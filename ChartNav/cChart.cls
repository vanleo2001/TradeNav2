VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Version (major.minor) for chart templates
'(won't load if major version does not match)
Const strChartVersion = "3.11" ' should increment just the 2-digit minor version

Const kTradeTimeFormat = "00000.###########"

Public Enum enumCursor
    eCursor_Default = 1
    eCursor_Pencil = 3
    eCursor_Hand = 4
    eCursor_CrossHair = 5
    eCursor_Vertical = 6
    eCursor_Horizontal = 7
    eCursor_ArrowNS = 8
    eCursor_ArrowEW = 9
    eCursor_ArrowNE = 10
    eCursor_ArrowNW = 11
    eCursor_Arrow4Way = 12
    eCursor_HorzSize = 13
    eCursor_NoDrop = 14
    eCursor_Magnify = 15
    eCursor_HandMoveUp = 16
    eCursor_HandMoveDown = 17
    eCursor_HandMoveLeft = 18
    eCursor_HandMoveRight = 19
    eCursor_HandDown = 20
    eCursor_HandLeft = 21
    eCursor_HandRight = 22
    eCursor_ChartMove = 23
    eCursor_Eraser = 24
    eCursor_OrderBuy = 25
    eCursor_OrderSell = 26
    eCursor_Blank = 27
End Enum

' levels of how much to redo when generating the chart
Public Enum eChartRedoMode
    eRedo1_Scrolled = 1
    eRedo2_ReloadAnnots = 2
    eRedo3_Settings = 3
    eRedo5_RecalcInd = 5
    eRedo6_ReloadInd = 6
    eRedo7_ReloadRT = 7
    eRedo9_ReloadData = 9
End Enum

Private Enum eDataLabelType
    eLBL_Chart = 0
    eLBL_Mouse = 1
    eLBL_DataWindow = 2
End Enum

Private Enum eLabelPeriod
    eLBL_Each = 0
    eLBL_Hour = 1
    eLBL_Day = 2
    eLBL_Week = 3
    eLBL_Month = 4
    eLBL_Year = 5
    eLBL_Decade = 6
End Enum

Private Enum eTradeField
    eTRADE_Date = 0
    eTRADE_Qty = 1
    eTRADE_Action = 2
    eTRADE_Price = 3
    eTRADE_Rule = 4
    eTRADE_Link = 5
    eTRADE_Profit = 6
    eTRADE_NetProfit = 7
    eTRADE_Position = 8
    eTRADE_AvgEntry = 9
End Enum

Public Enum eCompSym
    eCompSym_None
    eCompSym_PricePane
    eCompSym_NonPricePane
    eCompSym_PercentPane
End Enum

Public Enum eSeasonalCycleType
    eCycleType_Undefined = -1
    eCycleType_Year
    eCycleType_Quarter
    eCycleType_Month
    eCycleType_Week
    eCycleType_Days
    eCycleType_FullMoons
    eCycleType_NewMoons
End Enum

Public Enum eChartType
    eTypeChart_Undefined = -1
    eTypeChart_Standard = 0
    eTypeChart_Spread
    eTypeChart_PercentComp
    eTypeChart_Seasonal
    eTypeChart_Profile
End Enum

Private Type mPrivate
    eRedoMode As eChartRedoMode
    bTempAutoScale As Boolean
    nPartialLoadDate As Long
   
    dLastScreenDate As Double
    dExtraPriceScale As Double
    dLastActualDateOfData As Double
    dUserScreenDate As Double               'right most date on screen that user set chart at
    
    bRT As Boolean
    bDisableRT As Boolean
    
    nPrevNumPanes As Long
    
    nSymbolID As Long
    strExtData As String   ' external data file info
    strSpreadSymbols As String
    
    strTemplate As String
    strTemplateApplied As String
    strChartName As String
    strDesc As String
    strRequired As String
    
    Annots As cGdTree
    
    bCustomColors As Boolean
    nChartForeColor As Long
    nChartBackColor As Long
    nBorderForeColor As Long
    nBorderBackColor As Long
    nChartGradientColor As Long
    nUseGradient As Long ' 0 = off, 1 = on (-1 = not yet used)
    
    nPeriodicity As Long
    strAutoBreakout As String ' special override for DanielCode server
    
    'Design Note: resolution for aardvark 988
    'TLB: we decided to work with 2 separate variables ...
    '"BlankBars"
    '- can only be adjusted from chart settings
    '- when hit "End" key or first load chart or reset BlankBars, set ForecastBars to this value
    '"ForecastBars"
    '- used to determine end of scrollbar (and size of aXbar and aXdate arrays)
    '- can be adjusted by "Chart Move", but cannot be less than BlankBars
    bShowEmptyBars As Boolean
    nForecastBars As Long
    nBlankBars As Long
    nBlankBarsSave As Long
    
    'Design note: changes for Elliott Wave (4453)
    'Pixels per bar can be change by +,- keys, scrunching the x-scale, right-click on the chart etc.
    'DefaultPixelsPerBar can only be changed by right-clicking the red arrow key (Per Pete)
    nPixelsPerBar As Long
    nDefaultPixelsPerBar As Long
    
    bUnsplit As Boolean
    nShowSplitPane As Long  '0/1: hide/show extra info, e.g. Woodies stuff, to right of pane
    nVertGrid As Long       '0=coarse, 1=fine, 2=none
    bHorzGrid As Boolean
    
    iViewMode As Integer    '0=normal, 1=export, 2=print
    bHidePriceIndicators As Boolean
    
    dFromDate As Double
    dToDate As Double
    bToEndOfData As Boolean
    nMaxIntradayDays As Long
    
    '-----------------
    
    aSpeed As New cGdArray
    bZoomed As Boolean
    
    ' Associates a bar num from m.Bars to chart X axis
    aXBar As New cGdArray
    aXdate As New cGdArray
    ' (the total # of bars for chart is aXbar.Size,
    '  each item of aXbar holds bar# of m.Bars to use
    '  at that "X" position of chart -- e.g. holidays
    '  have an item in aXbar of NullData since there
    '  is no bar in m.Bars for it -- also aXbar goes
    '  beyond the end of m.Bars to display future bars)
    nLastGoodDataBarX As Long
        
    aTDOM As New cGdArray   ' array of TradeDayOfMonth
    aTDOY As New cGdArray   ' array of TradeDayOfYear/TradeWeekOfYear
    aTDOW As New cGdArray   ' array of TradeDayOfWeek
               
    'gdaDates As Long
    'or gdBars?
    Bars As cGdBars
    OtherBars() As cGdBars  ' 1-based.  Needs to be kept NOT ALIGNED to Market1 in this array of bars
    SpreadBars() As cGdBars ' 0-based.
    
    BarsProfile As cGdBars  'for market profile style display, this is a shared object, do not create or destroy
    BarsProfileTicks As cGdBars
    
    ReplayTicks As cGdTree
    bMoreDataBefore As Boolean
    bMoreDataAfter As Boolean
    bRemoveOvernightPriceGap As Boolean ' special feature for SpyderTrader
    Settles As cGdBars
    
    RollsTable As cGdTable
    SplitsTable As cGdTable
    DividendsTable As cGdTable
    nNextEarningsDate As Long
    
    Form As Form
    hsb As gdScrollBar  '(way more efficient to reference this directly than m.Form.hsb -- only because m.Form is a generic "Form")
    tbToolbar As SSActiveToolBars
    
    Tree As cGdTree

    'for pattern profit composite data (can contain up to 4 cIndicator objects)
    CompositeIndicators As cGdTree     'Node1=OHLC, Node2=close line, Node3=stdDevAbove, Node4=stdDevBelow
    
    ' Strategy & trade stuff:
    bAccountBar As Boolean
    nShowTrades As Long         '0=none,1=strategy,2=trade account (used to be boolean)
    nShowOrderHorzLine As Long  '0/1 - extended horizontal line for orders
    nShowAvgEntryLine As Long   '0/1 - show horizontal average entry line
    nHighlightPos As Long       '-2/-1/>=0 - (-2)=was never set/saved, (-1)=don't highlight, (>=0)=highlight with specified color
    nHighlightEquity As Long
    bReloadTrackerTrades As Boolean
    nProfitLineBox As Long      '0=none,1=line,2=hollow box,3=filled box
    nSystemID As Long
    strSystemName As String
    System As cSystem
    hTradesHandle As Long
    tblTrades As cGdTable
    aOrders As New cGdArray
    dNextBarDate As Double
    frmOrdersPopup As frmAlertPopup
    bAutoTrade As Boolean
    nAccountID As Long          'account ID to show orders/trades for
    nPseudoOrderType As Long    'pseudo order type for setting combo box (0=auto,1=limit,2=stop)
    OnlineOrders As cGdTree     'holds cOrderStruct objects
    OrderDailyBars As cGdBars   'daily bars to get bid/ask
    nTradeSymbolID As Long      'so don't have to keep converting -055,-057,-065,-067 contracts
    strTradeSymbol As String
    nSymbolPitType As Long      '-1 = not set else corresponds to enumerated future symbol type in mMain.bas

    ' for Image Server: to get and store Rex data
    strRexDataFile As String
    aRexData As New cGdArray
    nImgSrvFontSize As Long
    nImgSrvChartWidth As Long   'requested width of chart in pixels
    iMiniChart As Long          ' = how many bars of previous day to show
    iMiniChartStartX As Long    'startscreenX for mini chart data
    bHideExits As Boolean       'to show entries but not exits (special case for Markus image server)
    
    BadTicksInd As New cIndicator
    nFirstBar As Long           'bar# (of m.Bars) of first actual data bar visible on the screen
    nLastBar As Long            'bar# (of m.Bars) of last actual data bar visible on the screen
    
    ' for delayed recalc of indicators while real-time
    dLastRecalcDone As Double
    dAvgRecalcTime As Double    ' exp mov avg of last 10 recalc's
    iTryLastBarRecalc As Long   ' 0 = normal, 1 = try recalc last bar, -1 = can't recalc last bar
    dLastCountdownCheck As Double       'to prevent Woodies timer calling generatechart too often
    
    'flag for whether spread is using "dollar" multiplier
    bSpreadAsDollar As Boolean
    'flag for profile chart
    bProfileChart As Boolean
    profileChartInfo As profile_info
    
    'string for trade account settings
    strOrdBarCtrls As String    'POS|BM|SM|CLR|QTY|OT|BC|SC|BR|RV|CA|FL|PR|BA|SA|BB|SB
    strLimitTrades As String    'numeric| followed by T,B,D or M for trades,bars,days or minutes
        
    'for average entry line (don't want to have to search annots tree for this)
    AvgEntryAnnot As cAnnotation
    strAvgEntryProp As String           'color|lineStyle|font|fontSize|position (0=below, 1=above)
    strAcctBarCols As String            '
    dAvgEntry As Double                 'if zero then don't show avgEntryLine even if show flag is on
    
    aWhatIfSave As cGdArray             'saves real OHLC data of last bar in [0][1][2][3]
    
    'for bid/ask on chart
    BidAskAnnot As cAnnotation
    eBidAskMode As eBidAskColorMode
    eBidAskImage As eStockImage
    eBidAskImgDir As eImageDir
    nBidAskImgFill As Long              '0=hollow, 1=filled
    nBidAskImgSize As Long
    nBidColor As Long
    nAskColor As Long
    
    'grapheng.dll structure
    geChartObj As Long
    geChartStruct As chart_win
    geChartHeader As chart_annotation
    geChartFooter As chart_annotation
    strTitle As String
    strSubTitle As String
    
    Transactions As cTransactions       ' Table of transactions from the trade-by-trade array
    EquityFilter As cEquityFilter       ' Equity filter information
    
    strTemplateStartEndTimes As String
    strLinkedInputs As String           'format: &name;value;type - name starts with special char &, type is ParmTypeID in database (1,2,5,6)
    strOptionSym As String
    strPfpIndKeys As String             'pipe-delimited string for indicators user used for Pattern Profit
    
    strSeasonalCycle As String          'eg "1 year"
    eCycleType As eSeasonalCycleType
    nCycleLen As Long
    
    ' TLB 5/12/2008: temporary flags while debugging the new Recalc Last Bar method
    bUseNewRecalcMethod As Boolean
    iVerifyResults As Long
    strRecalcMsg As String
    
    ' JM 09/29/2008: array of string & flags for EWaves prototype
    aEWave As cGdArray
    bEWaveFile As Boolean
    nDegreeFilter As Long
    
    bRedrawTwice As Boolean
    bPriceAlertsChecked As Boolean
    nPriceTopMost As Long               '0:default draw order, 1:draw price on top of all indicators
    
    nShowToolbar As Long
    eTypeOfChart As eChartType
    
    LastEditedAnnot As cAnnotation
    LastCreatedAnnot As cAnnotation
    
    PivotVwapInd As cIndicator          'saved for determining line breaks on VWAP bands
    
    bUnpublishable As Boolean ' a flag just for "publishers" to mark chart so will NOT be published
    
    DrawingCommands As cGdArray
End Type
Private m As mPrivate

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.DrawingCommands = New cGdArray
    m.DrawingCommands.Create eGDARRAY_Strings, 0

    ' Use new method unless flag file exists at the time this chart is first loaded
    m.bUseNewRecalcMethod = False
    If HasNewNavEngine Then
        If FileLength(App.Path & "\Provided\OldRecalc.flg") <= 1 Then
            m.bUseNewRecalcMethod = True
        End If
    End If

    m.bEWaveFile = FileExist(App.Path & "\EWAVE\dji.txt")

    Set m.System = Nothing
    Set m.Bars = New cGdBars
    Set m.RollsTable = New cGdTable
    Set m.SplitsTable = New cGdTable
    Set m.DividendsTable = New cGdTable
    Set m.Tree = New cGdTree
    m.Tree.AllowMultipleObjectTypes = True
    Set m.Transactions = New cTransactions
    Set m.EquityFilter = New cEquityFilter

    m.aXBar.Create eGDARRAY_Longs
    m.aXdate.Create eGDARRAY_Doubles
    m.aTDOM.Create eGDARRAY_Longs
    m.aTDOY.Create eGDARRAY_Longs
    m.aTDOW.Create eGDARRAY_Longs
            
    ReDim m.OtherBars(0) As cGdBars
    ReDim m.SpreadBars(0) As cGdBars
    Set m.ReplayTicks = New cGdTree
    Set m.Settles = New cGdBars
            
    LoadDefaultSettings

    m.nSymbolID = 0

    Set m.Annots = New cGdTree
    Set m.OrderDailyBars = New cGdBars
    
    'initialize grapheng.dll structure
    m.geChartStruct.gdhDate = gdCreateArray(eGDARRAY_Doubles)
    m.geChartStruct.glhWinRect = gdCreateArray(eGDARRAY_Longs, 4)
    m.geChartStruct.glhChartRect = gdCreateArray(eGDARRAY_Longs, 4)
    m.geChartStruct.gshFont = gdCreateArray(eGDARRAY_gdString)
    m.geChartStruct.gdhPrintMargin = gdCreateArray(eGDARRAY_Doubles, 4)
    m.geChartStruct.gshTimeZoneInf = gdCreateArray(eGDARRAY_gdString)
    
    'profile chart struct
    m.profileChartInfo.styleTPO = 0
    m.profileChartInfo.ColorScheme = 0
    m.profileChartInfo.RTOnOff = 0
    m.profileChartInfo.reserved3 = Abs(FileExist("ExgridProfilerOn.flg"))
    
    m.profileChartInfo.glhItemShow = gdCreateArray(eGDARRAY_Longs, 8)
    m.profileChartInfo.glhItemColor = gdCreateArray(eGDARRAY_Longs, 8)
    m.profileChartInfo.gdhItemParm = gdCreateArray(eGDARRAY_Doubles, 8)
    m.profileChartInfo.gdhXDateCopy = gdCreateArray(eGDARRAY_Doubles, 0)
    m.profileChartInfo.gdhreserved1 = 0
    m.profileChartInfo.gdhreserved2 = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.Initialize", eGDRaiseError_Raise

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Dim i&, astrParms As New cGdArray
    
    ' clear the expression set in the engine
    If Not g.bUnloading And Not m.Form Is Nothing Then
        astrParms.Size = 1
        astrParms(0) = "Chart" & Str(m.Form.hWnd)
        SetupExpressions astrParms
        
        If HasNewNavEngine Then
            astrParms.Size = 1
            astrParms(0) = "Chart-" & Str(m.Form.hWnd)
            ClearLastBarSet astrParms.ArrayHandle, ByVal 0&
        End If
    End If
    
    RemoveAlerts True
    KillSystemReports
    
    ProfileBarsClear
    ReDim m.OtherBars(0) As cGdBars
    ReDim m.SpreadBars(0) As cGdBars
    Set m.ReplayTicks = Nothing
    Set m.Settles = Nothing
    
    Set m.DrawingCommands = Nothing
    Set m.Tree = Nothing
    Set m.tbToolbar = Nothing
    Set m.hsb = Nothing
    Set m.Form = Nothing
    Set m.RollsTable = Nothing
    Set m.SplitsTable = Nothing
    Set m.DividendsTable = Nothing
    Set m.Bars = Nothing
    Set m.Annots = Nothing
    Set m.System = Nothing
    Set m.Transactions = Nothing
    Set m.EquityFilter = Nothing
    
    Set m.CompositeIndicators = Nothing
        
    Set m.aOrders = Nothing
    Set m.OrderDailyBars = Nothing
    Set m.OnlineOrders = Nothing
    Set m.aSpeed = Nothing
    Set m.aRexData = Nothing
    Set m.frmOrdersPopup = Nothing
    Set m.BadTicksInd = Nothing
    Set m.AvgEntryAnnot = Nothing
    Set m.aWhatIfSave = Nothing
    Set m.aEWave = Nothing
    Set m.LastEditedAnnot = Nothing
    Set m.LastCreatedAnnot = Nothing
    Set m.PivotVwapInd = Nothing

    'MJM - destroy arrays
    Set m.aXBar = Nothing
    Set m.aXdate = Nothing
    Set m.aTDOM = Nothing
    Set m.aTDOY = Nothing
    Set m.aTDOW = Nothing
    
    '12-12-2012: destroy profile info structure
    gdDestroyArray m.profileChartInfo.glhItemShow
    gdDestroyArray m.profileChartInfo.glhItemColor
    gdDestroyArray m.profileChartInfo.gdhItemParm
    gdDestroyArray m.profileChartInfo.gdhXDateCopy
        
    'MJM - destroy grapheng.dll structure
    geCloseChart m.geChartObj
    m.geChartObj = 0
    gdDestroyArray m.geChartStruct.gdhDate
    gdDestroyArray m.geChartStruct.glhWinRect
    gdDestroyArray m.geChartStruct.glhChartRect
    gdDestroyArray m.geChartStruct.gshFont
    gdDestroyArray m.geChartStruct.gshTimeZoneInf
    gdDestroyArray m.geChartStruct.gdhPrintMargin
    
    gdDestroyArray m.geChartHeader.glhAnnType
    gdDestroyArray m.geChartHeader.gshText
    gdDestroyArray m.geChartHeader.gshFont
    gdDestroyArray m.geChartHeader.glhStyle
    gdDestroyArray m.geChartHeader.glhSize
    gdDestroyArray m.geChartHeader.gdhLeft
    gdDestroyArray m.geChartHeader.gdhBottom
    gdDestroyArray m.geChartHeader.gdhRight
    gdDestroyArray m.geChartHeader.glhAlign

    gdDestroyArray m.geChartFooter.glhAnnType
    gdDestroyArray m.geChartFooter.gshText
    gdDestroyArray m.geChartFooter.gshFont
    gdDestroyArray m.geChartFooter.glhStyle
    gdDestroyArray m.geChartFooter.glhSize
    gdDestroyArray m.geChartFooter.gdhLeft
    gdDestroyArray m.geChartFooter.gdhBottom
    gdDestroyArray m.geChartFooter.gdhRight
    gdDestroyArray m.geChartFooter.glhAlign

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.Terminate", eGDRaiseError_Raise

End Sub

' just need to clear all the other pointers to this cChart object so will terminate from Form
Public Sub DestroyChartRefs()
    
    On Error Resume Next
    Set m.Annots = Nothing
    Set m.Tree = Nothing

End Sub

Public Property Get Tree() As cGdTree
On Error GoTo ErrSection:

    If m.Tree Is Nothing Then
        Set m.Tree = New cGdTree
        m.Tree.AllowMultipleObjectTypes = True
    End If
    Set Tree = m.Tree

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.Tree.Get", eGDRaiseError_Raise

End Property

Public Property Get Bars() As cGdBars
On Error GoTo ErrSection:

    If m.Bars Is Nothing Then
        Set m.Bars = New cGdBars
    End If
    Set Bars = m.Bars

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.Bars.Get", eGDRaiseError_Raise

End Property

Public Property Get Form() As Form
    Set Form = m.Form
End Property

Public Property Set Form(frm As Form)
On Error GoTo ErrSection:
    
    If IsFrmChart(frm) Then
        If IsFrmChartMDI(frm) Then
            Set m.tbToolbar = frmMain.tbToolbar         'non-detached chart
        End If
        Set m.Form = frm
        Set m.hsb = frm.hsb
    Else
        Set m.tbToolbar = Nothing
        Set m.hsb = Nothing
        If frm Is Nothing Then Set m.Form = Nothing
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.Form.Set", eGDRaiseError_Raise

End Property

Public Property Get tbToolbar() As SSActiveToolBars
    Set tbToolbar = m.tbToolbar
End Property

Public Property Set tbToolbar(tb As SSActiveToolBars)
    
    If Not tb Is Nothing Then
        Set m.tbToolbar = Nothing
        Set m.tbToolbar = tb
    End If
        
End Property

Public Sub AddStudy(ByVal strStudy$, ByVal bDisplay As Boolean)
On Error GoTo ErrSection:

    AddStudyToChart Me, strStudy, bDisplay

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.AddStudy", eGDRaiseError_Raise

End Sub

Private Sub LoadDefaultSettings()
On Error GoTo ErrSection:

    Dim bOK As Boolean
    
    ' some defaults
    
    m.eCycleType = eCycleType_Undefined
    m.eTypeOfChart = eTypeChart_Undefined
    
    m.bShowEmptyBars = True
    m.bUnsplit = False
    m.nVertGrid = 0     'coarse grid
    m.bHorzGrid = False
    MyForecastBars = 12
    MyBlankBars = 12
    m.nPixelsPerBar = 0 '(so will get default)
    m.nDefaultPixelsPerBar = 0
    m.nShowSplitPane = 0
    m.nHighlightPos = -2
    m.nHighlightEquity = -2
    m.strAutoBreakout = ""
    m.bRemoveOvernightPriceGap = False
    m.bProfileChart = False
    m.bHideExits = False
    m.bUnpublishable = False
    
    'bCandlesticks = False
    
    'nGenTick = 0

    m.iViewMode = 0
    
    bOK = True
    If Not m.Form Is Nothing Then
        If Not m.Form.GameMode Is Nothing Then
            If m.Form.GameMode.CustomOrders > 0 Then
                bOK = False
            End If
        End If
    End If
    
    If bOK Then
        m.nPeriodicity = ePRD_Days + 1
        m.dFromDate = DateSerial(1900, 1, 1)
        m.dToDate = DateSerial(2001, 12, 31)
        m.bToEndOfData = True
        m.nMaxIntradayDays = 10
    End If
    
    m.bCustomColors = False
    m.nBorderBackColor = 0
    m.nBorderForeColor = 0
    m.nChartBackColor = 0
    m.nChartForeColor = 0
    m.nChartGradientColor = GradientDefault
    m.nUseGradient = 0
       
    m.strRequired = ""
    m.strChartName = ""
    m.strDesc = ""
    m.bHidePriceIndicators = False
        
    If ExtremeCharts >= 1 Then
        m.geChartStruct.LogoColor = 1 ' per Glen: do a Genesis logo for ETA version (since has www.BetterTrades.com)
    Else
        m.geChartStruct.LogoColor = -1 'no logo now per Glen's email 9/17/04  ' = 1 ' (-1 = off)
    End If
    
    'boolean indicator for highlighting bad ticks
    With m.BadTicksInd
        .CodedName = "Bad Ticks"
        .DataType = eINDIC_BooleanArray
        .BadTicksFlag = True
        .Display = False
        .DisplayType = eINDIC_HighlightBars
        .MarkerImage = eCNI_Arrow
        .MarkerDir = eCNI_North
        .MarkerLoc = 1
        .ChartLabel = "Bad Ticks"   'ask Tim
    End With
    
    With m.EquityFilter
        .MovingAverageType = "None"
        .MovingAveragePeriod = 0&
        .EquityFilterOn = False
    End With
    
    m.Transactions.EquityFilter = m.EquityFilter
    
    OrdBarGetDefaults
    
    BidAskPropLet eBidAskColorMode_None, eCNI_Diamond, eCNI_North, 1, 2, vbBlue, vbRed

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.LoadDefaultSettings", eGDRaiseError_Raise

End Sub

' return Trading day of Month for bar
Public Function TDOM(ByVal nBar&) As Long
    TDOM = m.aTDOM.Num(nBar)
End Function
' return Trading day of Year for bar
Public Function TDOY(ByVal nBar&) As Long
    TDOY = m.aTDOY.Num(nBar)
End Function
' return Trading day of Week for bar
Public Function TDOW(ByVal nBar&) As Long
    TDOW = m.aTDOW.Num(nBar)
End Function

Private Sub FixPaneSizes()
On Error GoTo ErrSection:

    Dim i&, iPricePane&, bBad As Boolean
    Dim dDefaultSize#, dTotalSize#, X#
    Dim Pane As cPane
    Dim nVisPanesIdx&, aVisiblePanes As New cGdArray       'MJM - added for grapheng.dll

    'if chart tree is empty then just exit
    If m.Tree.Count = 0 Then Exit Sub

    dDefaultSize = 0.19 '(default size for non-price panes)
    
    'MJM [begin] - added for grapheng.dll
    nVisPanesIdx = 0
    aVisiblePanes.Create eGDARRAY_Longs
    For i = 1 To m.Tree.Count
        Set Pane = Nothing
        If m.Tree.NodeLevel(i) = 0 Then
            Set Pane = m.Tree(i)
            If Pane.Display = True Then
                gdSetNum aVisiblePanes.ArrayHandle, nVisPanesIdx, i
                nVisPanesIdx = nVisPanesIdx + 1
                If Pane.Size = 0 Then Pane.Size = dDefaultSize
                X = Pane.Size
                If InStr(m.Tree.Key(i), "PRICE") > 0 Then
                    iPricePane = i
                Else
                    dTotalSize = dTotalSize + X
                End If
                'If X < 0.049999 Or X > 0.950001 Or dTotalSize > 1.0001 Then
                '    bBad = True
                '    Exit For
                'End If
            End If
        End If
    Next
    ' make price pane whatever's leftover
    If iPricePane > 0 Then
        X = 1 - dTotalSize
        dTotalSize = dTotalSize + X
        ' make sure enough room for price pane
        If X < 0.05 Or dTotalSize > 1.0001 Then
            bBad = True
        ' if have just added one or more panes, make sure price pane is still "big enough"
        ElseIf aVisiblePanes.Size > m.nPrevNumPanes And m.nPrevNumPanes > 0 Then
            If X < 2# / (aVisiblePanes.Size + 1) Then
                bBad = True
            End If
        End If
        If Not bBad Then
            Set Pane = m.Tree(iPricePane)
            Pane.Size = X
        End If
    End If
    ' check total size
    If dTotalSize < 0.9999 Or dTotalSize > 1.0001 Then
        bBad = True
    End If
        
    Set Pane = Nothing
    If bBad Then
        ' need to resize panes to default sizes
        ' (make price pane at least double the other panes)
        If iPricePane > 0 Then
            X = 1# / (aVisiblePanes.Size + 1)
        ElseIf aVisiblePanes.Size > 0 Then
            'this gdArray size is zero when tree.count is zero
            'this is divide by zero error bug that Dave saw on his laptop
            X = 1# / aVisiblePanes.Size
        End If
        If X > dDefaultSize Then
            X = dDefaultSize
        End If
        For i = 0 To aVisiblePanes.Size - 1
            Set Pane = m.Tree(aVisiblePanes(i))
            If aVisiblePanes(i) = iPricePane Then
                Pane.Size = 1 - (aVisiblePanes.Size - 1) * X
            Else
                Pane.Size = X
            End If
            Set Pane = Nothing
        Next
    End If
    
    m.nPrevNumPanes = aVisiblePanes.Size
    aVisiblePanes.Destroy
    Set Pane = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.FixPaneSizes", eGDRaiseError_Raise

End Sub

Private Sub LoadData()
On Error GoTo ErrSection:

    Dim nNum&, nCF%, strFileName$, strErrMsg$, strSymbol$, strFunc$, strText$, strFormat$
    Dim i&, j&, d#, dRange#, iX&, iBar&, dTimeout#
    Dim nFirstDate&, nLastDate&, nMPB&, nPoolRec&, nLoadFrom&
    Dim hBarsDateArray&, hXdateArray&, hXbarArray&, hTDOMArray&, hTDOYArray&, hTDOWArray&
    Dim dCurDate#, dBarDate#, dTempDate#
    Dim nTDOMctr&, nTDOYctr&, nTDOWctr&, nNextMonth&, nNextYear&, nNextWeek&
    Dim bLongDate As Boolean
    Dim nPeriodicity As Long
    Dim dMinuteBar#, dCrossoverOffset#, dSessionStart#, dSessionEnd#
    Dim aSpreadSymbols As New cGdArray
    Dim aHandles As New cGdArray
    Dim oFunction As cFunction
    Dim TempBars As New cGdBars
    Dim Indicator As cIndicator
    
    Dim nBlankBarsShown As Long
    Dim bSkipPartialLoad As Boolean
    Dim bSeasonal As Boolean
    
    Dim bAddToRecentSymbolList As Boolean

'JM 02-14-2013: these are never set in this routine (don't think they are needed)
'    Dim dDateSaveRT As Double
'    Dim bWasOnScreen As Boolean
    
    If Not m.Form Is Nothing Then
        m.Form.chkConfirmOrder.Visible = g.RealTime.Active
        If g.bLoadingChartPage Or g.bStarting Then
            m.Form.FixOrderBarControls True, True, m.strOrdBarCtrls
        Else
            m.Form.FixOrderBarControls False, True, m.strOrdBarCtrls
        End If
        If m.Form.IsInGameMode Then
            If m.Form.GameReplayMode <> eGDReplayMode_Play And m.Form.GameReplayMode <> eGDReplayMode_Sync Then
                Set m.System = Nothing
            End If
        End If
    End If
       
    m.dExtraPriceScale = 0
    If TypeOfChart = eTypeChart_Seasonal Then bSeasonal = True
    
    ' date range
    m.bMoreDataAfter = False
    m.bMoreDataBefore = False
    If m.nSymbolID <> 0 Then
        nPoolRec = g.SymbolPool.PoolRecForSymbolID(m.nSymbolID)
    ElseIf Len(m.strSpreadSymbols) > 0 Then
        'format: funcName~funcID;op,sym,mult;op,sym,mult...
        ' get spread function name and spread symbols
        aSpreadSymbols.SplitFields m.strSpreadSymbols, ";"
        strFunc = aSpreadSymbols(0)
        ' if function exists, get definition again in case it changed
        If Len(strFunc) > 0 Then
            Set oFunction = New cFunction
            j = InStr(strFunc, "~")
            If j <> 0 Then
                oFunction.FunctionID = Mid(strFunc, j + 1, Len(strFunc) - j)
                oFunction.Load
                strFunc = Left(strFunc, j - 1)
            End If
            If Len(oFunction.FunctionName) = 0 Or oFunction.FunctionID < 1 Then
                oFunction.FunctionID = FunctionIDFromName(strFunc)
                If oFunction.FunctionID > 0 Then oFunction.Load
            End If
            If oFunction.FunctionID > 0 Then
                strFunc = oFunction.Formatted
                j = InStr(strFunc, "{")
                If j <> 0 Then
                    strFunc = Mid(strFunc, j + 1, Len(strFunc) - j)
                    j = InStr(strFunc, "}")
                    If j <> 0 Then
                        strFunc = Left(strFunc, j - 1)
                        strFunc = Replace(strFunc, " ", "")         '6682
                        strFunc = oFunction.FunctionName & "~" & oFunction.FunctionID & strFunc
                        If strFunc <> m.strSpreadSymbols Then
                            m.strSpreadSymbols = strFunc
                            aSpreadSymbols.SplitFields m.strSpreadSymbols, ";"
                        End If
                    End If
                End If
            End If
        End If
        aSpreadSymbols.Remove 0, 1
        ' get first symbol in order to use for dates
        strSymbol = Parse(aSpreadSymbols(0), ",", 2)
        nPoolRec = g.SymbolPool.PoolRecForSymbol(strSymbol)
    Else
        nPoolRec = -1
    End If
    If Not m.Form.GameMode Is Nothing Then
        nLastDate = Int(m.Form.GameMode.GameDataTime) + 1
    ElseIf g.nReplaySession > 0 Then
        nLastDate = g.nReplaySession
    ElseIf m.bToEndOfData Then
        nLastDate = 0
    Else
        nLastDate = m.dToDate
        If nLastDate < g.SymbolPool.LastDate(nPoolRec) Then
            m.bMoreDataAfter = True
        End If
    End If
    nFirstDate = m.dFromDate
    If GetPeriodType(m.nPeriodicity) < ePRD_Days Then
        ' for intraday charts, let max intradays override from date
        ' (as # of weekdays back)
        
        'JM 02-15-2013: maxintraday is also used as #of profiles for intraday profile charts
        '   this code is slightly different from code determining start/end dates for trade profile form
        '   so the start date on trade profile ends up being 1 day after start date on chart.
        '   Either this or the code in trade profile form can be changed if the 2 need to match exactly
        
        i = g.SymbolPool.LastDate(nPoolRec)
        If i < 0 Then
            i = 0
        ElseIf i > nLastDate And nLastDate <> 0 Then
            i = nLastDate
        End If
        If i > 0 And i - m.nMaxIntradayDays * 7# / 5 >= nFirstDate Then
            nFirstDate = i - Int(m.nMaxIntradayDays / 5) * 7
            For i = (m.nMaxIntradayDays Mod 5) To 1 Step -1
                Do
                    nFirstDate = nFirstDate - 1
                Loop While Not IsWeekday(nFirstDate)
            Next
            If nFirstDate < m.dFromDate Then
                nFirstDate = m.dFromDate
            End If
        End If
    End If
    If nFirstDate > g.SymbolPool.FirstDate(nPoolRec) Then
        m.bMoreDataBefore = True
    End If
      
'Fix Option #1 for empty bar bug reported by Bernie
'To duplicate: use $DJIA on daily chart, set streaming replay to start on 12/26/2007 before $DJIA open time
'When this check is not here, you will see a blank bar for 12/25/2007 as data for 12/26/2007 comes in.
'    If g.RealTime.Active And Not m.bShowEmptyBars Then
'        If Not m.Bars.IsIntraday And RedoMode <> eRedo9_ReloadData Then
'            i = m.Bars.FindDateTime(Int(LastDailyDownload), True)
'            If i >= 0 And i < m.Bars.Size Then
'                If m.Bars(eBARS_Close, i) = kNullData Then RedoMode = eRedo9_ReloadData
'            End If
'        End If
'    End If
         
    ' TLB 3/10/2010: we don't really have a good way to reload just the data for today
    ' for the special bar types (P&F, Kagi, Renko), so we just have to reload it all
    ' TLB 7/24/2013: same for bar periods > 1 day per bar
    If RedoMode <> eRedo9_ReloadData Then
        If (m.nPeriodicity >= ePRD_IntRenko And m.nPeriodicity < ePRD_IntVol) Or m.nPeriodicity > ePRD_Days + 1 Then
            RedoMode = eRedo9_ReloadData
            bSkipPartialLoad = True ' (no reason to load in chunks since minutized already built)
        End If
    End If
         
    strSymbol = Symbol
    If RedoMode >= eRedo9_ReloadData Then
        ReDim m.SpreadBars(0) As cGdBars
        Set m.SpreadBars(0) = Nothing
        ''m.ReplayTicks.Clear               'aardvark 4748
        
'JM 08-03-2011: do not set dirty page flag here, this code is hit when loading detached
'   charts from a chart page AFTER all chart page loading flags have been cleared
'   the dirty page flag is now set when changing symbol & bars periodicity
'        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        
        g.RealTime.RemoveTickBuffer m.Bars
        
        'reset last screen date so red arrow will not show
        If m.geChartStruct.scaleArrow <> 2 Then
            ResetLastScreenDate
            If m.Form Is Nothing Then
                m.nForecastBars = m.nBlankBars
            ElseIf m.Form.GameMode Is Nothing Then
                m.nForecastBars = m.nBlankBars      '5118
            End If
        End If
        
        ' load bars (can only do up to daily prior to appending real-time)
        nPeriodicity = m.nPeriodicity
        If nPeriodicity > ePRD_Days + 1 Then
            nPeriodicity = ePRD_Days + 1 'Note: still need to do this for other things besides SpliceBars (e.g. SplitsTable, last date, etc.)
        ElseIf nPeriodicity >= ePRD_IntRenko And nPeriodicity < ePRD_IntVol Then
            ' for Renko/Kagi/P&F, we need to get 1 minute bars for both history and today (from streaming),
            ' then convert them from 1 minute bars only after splicing the data together
            nPeriodicity = ePRD_Minutes + 1
        ElseIf GetPeriodType(nPeriodicity) = ePRD_SMP Then
            ' for SMP bars, we need to get "n" minute bars for both history and today (from streaming),
            ' then convert them from minute bars only after splicing the data together
            nPeriodicity = ePRD_Minutes + Int(GetPeriodsPerBar(nPeriodicity) / 10000#)
        End If
        If m.nPartialLoadDate = 0 And Not bSkipPartialLoad Then
            Screen.MousePointer = vbHourglass
        End If
        dTimeout = gdTickCount + 750
        If Len(m.strExtData) > 0 Then
            ' load data from external file (e.g. CSI, MS)
            strSymbol = Parse(m.strExtData, "|", 1)
            strFileName = Parse(m.strExtData, "|", 2)
            strFormat = UCase(Parse(m.strExtData, "|", 5))
            If strFormat <> "GDB" Then
                strFileName = FilePath(strFileName)
            End If
            m.Bars.ArrayMask = eBARS_Eod
            SetBarProperties m.Bars, strSymbol
            i = 0 '19990501
            nNum = m.Bars.FromFile(strFormat, strFileName, strSymbol, "", strErrMsg, nFirstDate, nLastDate)
            m.nPartialLoadDate = 0
        ElseIf Len(m.strSpreadSymbols) > 0 Then
            ' load symbols for spread
            aHandles.Create eGDARRAY_Longs
            ' see if bar period is allowed for spreads
            If Len(BarPeriodError(m.strSpreadSymbols, nPeriodicity)) > 0 Then
                nFirstDate = Date + 100 ' (just out into the future so will not waste time loading anything)
            End If
            ReDim m.SpreadBars(aSpreadSymbols.Size - 1) As cGdBars
            For i = 0 To aSpreadSymbols.Size - 1
                Set m.SpreadBars(i) = New cGdBars
                strSymbol = Parse(aSpreadSymbols(i), ",", 2)
                ' TLB 7/8/2011: for spreads, only load the primary symbol at the requested bar period,
                ' but load the others as either daily or 1-minute so as to get time as close as possible
                If i = -1 Then ''= 0 Then
                    DM_GetBars m.SpreadBars(i), strSymbol, nPeriodicity, nFirstDate, nLastDate, , , m.bUnsplit
                ElseIf IsIntraday(nPeriodicity) Then
                    DM_GetBars m.SpreadBars(i), strSymbol, "1 min", nFirstDate, nLastDate, , , m.bUnsplit
                Else
                    DM_GetBars m.SpreadBars(i), strSymbol, "Daily", nFirstDate, nLastDate, , , m.bUnsplit
                End If
                aHandles(i) = m.SpreadBars(i).BarsHandle
                If i = 0 Then
                    SetBarProperties m.Bars, strSymbol
                    m.Bars.ArrayMask = m.SpreadBars(i).ArrayMask
                End If
                m.SpreadBars(i).Prop(eBARS_Desc) = aSpreadSymbols(i)
            Next
If 0 Then
            m.Bars.Size = 0
            m.Bars.Prop(eBARS_Periodicity) = m.SpreadBars(0).Prop(eBARS_Periodicity)
            aHandles.Add m.Bars.BarsHandle
            i = gdCalcSpread(aHandles.ArrayHandle, aSpreadSymbols.JoinFields(";"), 0)
            m.Bars.Prop(eBARS_Periodicity) = m.nPeriodicity
            m.Bars.BuildBars m.Bars.Prop(eBARS_PeriodicityStr)
Else
            Set TempBars = m.Bars.MakeCopy(True)
            TempBars.Prop(eBARS_Periodicity) = m.SpreadBars(0).Prop(eBARS_Periodicity)
            aHandles.Add TempBars.BarsHandle '' m.Bars.BarsHandle
            i = gdCalcSpread(aHandles.ArrayHandle, aSpreadSymbols.JoinFields(";"), 0)
            m.Bars.Size = 0
            m.Bars.Prop(eBARS_Periodicity) = m.nPeriodicity
            m.Bars.BuildBars m.Bars.Prop(eBARS_PeriodicityStr), TempBars.BarsHandle
            Set TempBars = Nothing
End If
            m.nPartialLoadDate = 0
            strSymbol = Symbol
            
            ' for spread charts, OHLC displays not valid for non-intraday data
            If Not IsIntraday(nPeriodicity) Then
                Set Indicator = m.Tree("PRICE")
                If Not Indicator Is Nothing Then
                    If Indicator.DisplayType < 0 Then
                        Indicator.DisplayType = eINDIC_Line ' change to just a Close line
                    End If
                    Set Indicator = Nothing
                End If
            End If
        ElseIf m.nPartialLoadDate <= 0 Then
            ' if realtime is running, then check if today's data has been requested yet
            ' (if not, make the request now before past history is loaded to help speed things up)
            If g.RealTime.SalmonIsRunning And Len(strSymbol) > 0 Then
                If g.RealTime.Active And m.bToEndOfData And Not m.bDisableRT Then
                    g.RealTime.SymbolInfo(strSymbol).GetDataRequestStatus m.nPeriodicity
                End If
            End If
            
            ' set start stop times from template (only if default times match)
            If Len(m.strTemplateStartEndTimes) > 0 And m.nSymbolID > 0 Then
                SetBarProperties m.Bars, m.nSymbolID '(call this to make sure the defaults are set)
                If m.Bars.Prop(eBARS_DefaultStartTime) = Val(Parse(m.strTemplateStartEndTimes, ";", 1)) _
                        And m.Bars.Prop(eBARS_DefaultEndTime) = Val(Parse(m.strTemplateStartEndTimes, ";", 2)) Then
                    m.Bars.Prop(eBARS_StartTime) = Val(Parse(m.strTemplateStartEndTimes, ";", 3))
                    m.Bars.Prop(eBARS_EndTime) = Val(Parse(m.strTemplateStartEndTimes, ";", 4))
                End If
                ' clear the string so we don't keep using it (since they could change it later)
                m.strTemplateStartEndTimes = ""
            End If
            
            ' load data for symbol
            nLoadFrom = nFirstDate
            If IsIntraday(nPeriodicity) And (nPoolRec >= 0) And Not bSkipPartialLoad Then
                ' TLB 6/30/2008: when loading lots of intraday data, load in weekly chunks
                nLoadFrom = g.SymbolPool.LastDate(nPoolRec)
                If nLastDate > 0 And nLoadFrom > nLastDate Then
                    nLoadFrom = nLastDate
                End If
                If nLoadFrom = 0 Or nLoadFrom > LastDailyDownload Then
                    nLoadFrom = LastDailyDownload
                End If
                nLoadFrom = nLoadFrom - 5
                Do While Weekday(nLoadFrom) <> vbSaturday
                    nLoadFrom = nLoadFrom - 1
                Loop
            End If
            If nLoadFrom <= nFirstDate + 2 Then
                ' all data will be loaded now
                nLoadFrom = nFirstDate
                m.nPartialLoadDate = 0
            Else
                ' only partial data will be loaded now, then will be chunked in later
                m.nPartialLoadDate = nLoadFrom
            End If
            
            If Not bSeasonal Then
                i = DM_GetBars(m.Bars, m.nSymbolID, nPeriodicity, nLoadFrom, nLastDate, , , m.bUnsplit)
            End If
        End If
        If m.nPartialLoadDate > 0 Then 'And Not g.bStarting Then
            ' TLB 6/30/2008: for intraday data, keep loading one week at a time and prepend to existing
            ' (but to load quicker: keep loading until reach a 3/4 second timeout -- then exit here to
            ' let the chart draw, then the next chart timer process will end up calling LoadData again)
            Set TempBars = m.Bars.MakeCopy(True)
            Do While gdTickCount <= dTimeout
                TempBars.Size = 0
                nLoadFrom = m.nPartialLoadDate - 7
                If nLoadFrom <= nFirstDate + 2 Then
                    nLoadFrom = nFirstDate
                End If
                i = DM_GetBars(TempBars, m.nSymbolID, nPeriodicity, nLoadFrom, m.nPartialLoadDate, , False, m.bUnsplit)
                ' check if flag had been set to 0 while data was loading (if so, then just exit this sub)
                If m.nPartialLoadDate = 0 Then
                    Exit Sub
                End If
                If TempBars.Size = 0 Then
                    ' TLB 4/29/2013: if an empty week, try backing up one more week (e.g. $SSE has full-week gaps)
                    m.nPartialLoadDate = m.nPartialLoadDate - 7 ' back up one more week
                    nLoadFrom = m.nPartialLoadDate - 7
                    If nLoadFrom <= nFirstDate + 2 Then
                        nLoadFrom = nFirstDate
                    End If
                    i = DM_GetBars(TempBars, m.nSymbolID, nPeriodicity, nLoadFrom, m.nPartialLoadDate, , False, m.bUnsplit)
                End If
                If TempBars.Size = 0 Then
                    ' if no data for 2nd consecutive week, must be at start of tick history -- we're done loading
                    m.nPartialLoadDate = 0
                Else
                    ' prepend these bars to the existing bars
                    gdAppendBars m.Bars.BarsHandle, TempBars.BarsHandle, True
                    If nLoadFrom <= nFirstDate Then
                        m.nPartialLoadDate = 0 ' done
                    ElseIf m.nPartialLoadDate > 0 Then
                        m.nPartialLoadDate = m.nPartialLoadDate - 7 ' back up one more week
                    End If
                End If
                If m.nPartialLoadDate = 0 Then Exit Do
                ' TLB 9/3/2008: do NOT put a DoEvents in this loop
                ' (will mess up when changing bar period during a long load)
                ''DoEvents
            Loop
            Set TempBars = Nothing
        End If
        
        ' add TickBuffer(s) now
        bAddToRecentSymbolList = False
        If (nLastDate = 0 Or nLastDate > LastDailyDownload) And Not m.bDisableRT Then
            If Not Screen.ActiveForm Is frmSymbolGrid Then
                bAddToRecentSymbolList = True
            End If
            If Len(m.strSpreadSymbols) > 0 Then
                For i = 0 To UBound(m.SpreadBars)
                    g.RealTime.AddTickBuffer m.SpreadBars(i), bAddToRecentSymbolList
                Next
            ElseIf m.nSymbolID > 0 Then
                g.RealTime.AddTickBuffer m.Bars, bAddToRecentSymbolList
            End If
        End If
        
        AddSpeed "DM_GetBars (" & CStr(m.Bars.Size) & " bars)"
    ElseIf RedoMode = eRedo7_ReloadRT Then
        ' 5/15/2008: new design to work more like CQG ...
        ' If the red arrow is not showing, then reset last screen date so red arrow will
        ' not be showing after the new bar gets added (if red arrow pointing left, then
        ' chart will not move until the bars fill in to the point where the arrow disappears)
        If m.geChartStruct.scaleArrow = 0 Then
            ResetLastScreenDate
        End If
    End If
            
'Fix Option #2 for empty bar bug reported by Bernie
'To duplicate: use $DJIA on daily chart, set streaming replay to start on 12/26/2007 before $DJIA open time
    ' chop blank bars at end
'    For i = m.Bars.Size - 1 To 0 Step -1
'        If m.Bars(eBARS_Close, i) <> kNullData Then
'            If m.bShowEmptyBars Or m.Bars.IsIntraday Then
'                m.Bars.Size = i + 1
'                Exit For
'            ElseIf m.Bars(eBARS_Close, i - 1) <> kNullData Then
'                m.Bars.Size = i + 1
'                Exit For
'            End If
'        End If
'    Next
   
   ' chop blank bars at end
    If Not bSeasonal Then
        For i = m.Bars.Size - 1 To 0 Step -1
            If m.Bars(eBARS_Close, i) <> kNullData Then
                m.Bars.Size = i + 1
                Exit For
            End If
        Next
        
        If m.bProfileChart Then
            If RedoMode >= eRedo7_ReloadRT Then
                ProfileBarsClear    '6848, 7003
            End If
            
            i = m.Bars.SessionDate(0)
            j = m.Bars.SessionDate(m.Bars.Size - 1)
            If j - i > 10 Then i = j - 10
            Set m.BarsProfile = ProfileBarsGet(m.Bars.Prop(eBARS_SymbolID), i, j, True)
            If Not m.BarsProfile Is Nothing Then
                'check if profile bars for entire data range has already been retrieved
                'this will be true if forecast bars are changed during chart move or change to blank bars at end
                If Bars.SessionDateForTime(m.BarsProfile(eBARS_DateTime, 0), True) <= m.Bars.SessionDate(0) And _
                   Bars.SessionDateForTime(m.BarsProfile(eBARS_DateTime, m.BarsProfile.Size - 1), True) <= m.Bars.SessionDate(m.Bars.Size - 1) Then
                
                    'already done, continue
                    i = i
                    
                ElseIf m.BarsProfile.Size > 0 And i <> m.Bars.SessionDate(0) Then
                    m.Form.tmrProfileLoad.Enabled = True        'turn on timer to do partial loading
                End If
            End If
        End If
    End If
    
    ' TLB 2/22/2010: the "delta's" should now be coming from the daily settles
    ' (use the pit symbol if it exists, else use the synthetic if it exists)
    m.Settles.Size = 0
    
    If Not bSeasonal Then
        If g.bUsePitSettlesForDeltas And m.nSymbolID > 0 And SecurityType(strSymbol) = "F" _
                    And (m.nPeriodicity <= ePRD_Days + 1) Then
            strText = GetSettleSymbol(strSymbol)
            nFirstDate = m.Bars.SessionDate(0) - 7
            nLastDate = m.Bars.SessionDate(m.Bars.Size - 1)
            nPeriodicity = ePRD_Days + 1 ' use Daily
            DM_GetBars m.Settles, strText, nPeriodicity, nFirstDate, nLastDate, , , m.bUnsplit
            ' but if the data from the pit symbol has stopped updating, then just use settles from actual symbol
            If strText <> strSymbol Then
                If m.Settles.SessionDate(m.Settles.Size - 1) < m.Bars.SessionDate(m.Bars.Size - 1) - 7 _
                    Or m.Settles.Size = 0 Then
                        DM_GetBars m.Settles, strSymbol, nPeriodicity, nFirstDate, nLastDate, , , m.bUnsplit
                End If
            End If
            If g.RealTime.Active And m.bToEndOfData And Not m.bDisableRT And m.Settles.Size > 0 Then
                ''g.RealTime.AddTickBuffer m.Settles ', bAddToRecentSymbolList
                g.RealTime.SpliceBars m.Settles
            End If
        End If
        If m.Settles.Size = 0 Then
            g.RealTime.RemoveTickBuffer m.Settles
        End If
    End If
    
    ' append real-time data
    m.bRT = False
    m.aRexData.Clear
    If Len(m.strRexDataFile) > 0 Then
        ' Image Server
        If frmTest.Visible Then
            frmTest.AddList "Data loaded for  " & FileBase(m.strRexDataFile), True
        End If
        ' while waiting for Rex, get the fundamental data
        frmSnapshot.ShowMe m.nSymbolID, App.Path & "\Charts\" & FileBase(m.strRexDataFile) & ".FND"
        If frmTest.Visible Then
            frmTest.AddList "Fundamental data retrieved for  " & FileBase(m.strRexDataFile), True
        End If
        ' wait for Rex data file (15-second timeout)
        dTimeout = Now + 15 / 86400#
        Do While Now <= dTimeout
            If FileExist(m.strRexDataFile & ".DON") Then
                If frmTest.Visible Then
                    frmTest.AddList "Rex data ready for  " & FileBase(m.strRexDataFile), True
                End If
                ' store data from file (so can merge secondary symbol data later)
                m.aRexData.FromFile m.strRexDataFile & ".DAT"
                ' merge today's data from Rex
                AppendRexData m.Bars
                AddSpeed "RexData (" & CStr(m.Bars.Size) & " bars)"
                If frmTest.Visible Then
                    frmTest.AddList "Rex data merged for  " & FileBase(m.strRexDataFile), True
                End If
                Exit Do
            End If
            Sleep 0.1
        Loop
    ElseIf m.Form.IsInGameMode Then
        SpliceReplayTicks m.Bars
        AddSpeed "SpliceReplayTicks (" & CStr(m.Bars.Size) & " bars)"
    ElseIf g.RealTime.Active And m.bToEndOfData And Not m.bDisableRT Then
        If m.nSymbolID > 0 Then
            g.RealTime.SpliceBars m.Bars
            AddSpeed "SpliceBars (" & CStr(m.Bars.Size) & " bars)"
        ElseIf Len(m.strSpreadSymbols) > 0 Then
            aHandles.Create eGDARRAY_Longs
            For i = 0 To UBound(m.SpreadBars)
                g.RealTime.SpliceBars m.SpreadBars(i)
                aHandles(i) = m.SpreadBars(i).BarsHandle
            Next
            Set TempBars = m.Bars.MakeCopy(True)
            TempBars.Prop(eBARS_Periodicity) = m.SpreadBars(0).Prop(eBARS_Periodicity)
            aHandles.Add TempBars.BarsHandle '' m.Bars.BarsHandle
            i = gdCalcSpread(aHandles.ArrayHandle, aSpreadSymbols.JoinFields(";"), 0)
            m.Bars.Size = 0
            m.Bars.Prop(eBARS_Periodicity) = m.nPeriodicity
            m.Bars.BuildBars m.Bars.Prop(eBARS_PeriodicityStr), TempBars.BarsHandle
            Set TempBars = Nothing
        End If
        m.bRT = True
    End If
       
    ' save last date of data (before going to weekly bars, etc.)
    m.dLastActualDateOfData = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
    If Not m.Bars.IsIntraday Then
        m.dLastActualDateOfData = Int(m.dLastActualDateOfData)
    ElseIf m.Bars.Prop(eBARS_LastTickTime) > 0 Then
        m.dLastActualDateOfData = Int(m.dLastActualDateOfData) _
            + m.Bars.Prop(eBARS_LastTickTime) / 1440#
    End If
    
    ' load splits table now (in case compressing unsplit data)
    m.nNextEarningsDate = 0
    m.SplitsTable.NumRecords = 0
    m.DividendsTable.NumRecords = 0
    If m.nSymbolID > 0 And SecurityType(strSymbol) <> "F" Then
        Set m.SplitsTable = GetSplitsTable(m.nSymbolID)
        Set m.DividendsTable = GetDividendsTable(m.nSymbolID, (m.bUnsplit = False))
        i = m.SplitsTable.NumRecords + m.DividendsTable.NumRecords
        If i > 0 Then
            AddSpeed "Load Splits/Dividends (" & Str(i) & " records)"
        End If
    
        ' TLB 6/26/2013: next earnings date
        If SecurityType(strSymbol) = "S" Then
            If DM_GetSnapData(strSymbol, "NextEarnings", d) Then
                m.nNextEarningsDate = d
            End If
        End If
    End If
    
    ' now compress bars (if greater than daily)
    If Not bSeasonal Then
        If m.nPeriodicity > ePRD_Days + 1 Or (m.nPeriodicity >= ePRD_IntRenko And m.nPeriodicity < ePRD_IntVol) Then
        ''If m.nPeriodicity >= ePRD_IntRenko And m.nPeriodicity <> ePRD_Days + 1 Then
            m.Bars.Prop(eBARS_Periodicity) = m.nPeriodicity
                
            ' when updating renko bars realtime, we have to ignore the prices of the current bar
            ' that just started -- so remove the current bar to build bars, then restore last close after
            d = kNullData
            If GetPeriodType(m.nPeriodicity) = ePRD_IntRenko And m.Bars.Size > 1 Then
                If g.RealTime.Active And bSkipPartialLoad Then
                    i = m.Bars.Size - 1
                    If m.Bars(eBARS_DateTime, i) > g.RealTime.FeedTime(strSymbol) Then
                        d = m.Bars(eBARS_Close, i) ' restore this after the renko bars are built
                        m.Bars.Size = i
                    End If
                End If
            End If
            
            If m.bUnsplit And m.Bars.SecurityType <> "F" Then
                m.Bars.BuildBars m.Bars.Prop(eBARS_PeriodicityStr), , , m.SplitsTable.FieldArrayHandle(0)
            Else
                m.Bars.BuildBars m.Bars.Prop(eBARS_PeriodicityStr)
            End If
            
            ' now we can restore the last price for realtime renko
            If d <> kNullData Then
                m.Bars(eBARS_Close, m.Bars.Size - 1) = d
            End If
            AddSpeed "Build Bars (" & CStr(m.Bars.Size) & " bars)"
        ElseIf GetPeriodType(m.nPeriodicity) = ePRD_SMP Then
            ' now build the SMP bars from the Minute bars
            Set TempBars = m.Bars.MakeCopy
            m.Bars.Size = 0
            i = GetPeriodsPerBar(m.nPeriodicity) Mod 10000
            m.Bars.BuildSMPBars TempBars, Int(i / 100), (i Mod 100)
            Set TempBars = Nothing
        End If
    End If

    ' do one more bar period check (e.g. in case EodVol)
    If Len(m.strSpreadSymbols) > 0 Then
        strText = BarPeriodError(m.strSpreadSymbols, m.nPeriodicity)
    Else
        strText = BarPeriodError(m.Bars.Prop(eBARS_Symbol), m.nPeriodicity)
    End If
    If Len(strText) > 0 Then
        Bars.Size = 0
    End If

    ' TLB 2/22/2010: special feature for SyderTrader
    ' (to back-adjust all price bars such that the overnight price gap is removed)
    If m.bRemoveOvernightPriceGap And m.Bars.Prop(eBARS_Periodicity) <= ePRD_Days + 1 Then
        gdRemoveOvernightPriceGap m.Bars.BarsHandle
    End If

    ' save array handles for efficiency
    hBarsDateArray = m.Bars.ArrayHandle(eBARS_DateTime)
    hXdateArray = m.aXdate.ArrayHandle
    hXbarArray = m.aXBar.ArrayHandle
    hTDOMArray = m.aTDOM.ArrayHandle
    hTDOYArray = m.aTDOY.ArrayHandle
    hTDOWArray = m.aTDOW.ArrayHandle
    
    nNum = Bars.Size                    'this is new last good data bar for adjusting last screen date for real time streaming
    ' resize the bars to include the forecast bars
    ' (so seasonals function will fill in forecast)
    If Not bSeasonal Then
        m.Bars.AddForecastBars m.nForecastBars
    End If
           
    If Bars.Size > 0 And Bars(eBARS_DateTime, Bars.Size - 1) > 0 Then
        If m.dUserScreenDate > 0 Then
            m.dUserScreenDate = gdFixDateTime(m.dUserScreenDate)
            'only time user screen date should be > 0 is right after template load
            If m.dUserScreenDate > m.Bars(eBARS_DateTime, Bars.Size - 1) Then
                If m.strTemplate = "Replay" Then
                    m.dUserScreenDate = 0
                Else
                    While m.dUserScreenDate > m.Bars(eBARS_DateTime, Bars.Size - 1)
                        m.Bars.AddForecastBars 1
                        m.nForecastBars = m.nForecastBars + 1
                    Wend
                    m.dLastScreenDate = m.dUserScreenDate
                End If
            ElseIf m.dUserScreenDate < m.Bars(eBARS_DateTime, Bars.Size - 1) Then
                m.dLastScreenDate = m.dUserScreenDate
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JM: 02-14-2013 - think this was a work-around "fix" that is no longer needed, dDateSaveRT is never set in
'   this routine so don't think this code was getting executed at all (leave awhile then remove if all ok
'
'        ElseIf RedoMode = eRedo7_ReloadRT And dDateSaveRT > 0 Then
'            'original spec: do not show new blank bars as new data comes in until less than blank bars setting is showing
'            'modified spec: maintain blank bars shown as set by user if blank bars shown is less than blank bars setting
'            'reason: if blank bars setting is fairly large, e.g. > 10, and user had position the chart such that there
'            '   are only 3 blank bars to right of last good data bar, implementing original spec verbatim would cause
'            '   chart to jump several bars to to the right (aardvark 4129)
'            If bWasOnScreen Then
'                If dDateSaveRT >= m.Bars(eBARS_DateTime, nNum + m.nBlankBars - 1) Then
'                    m.dLastScreenDate = dDateSaveRT
'                Else
'                    m.dLastScreenDate = m.Bars(eBARS_DateTime, nNum + nBlankBarsShown - 1)
'                End If
'            Else
'                m.dLastScreenDate = dDateSaveRT
'            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
    End If
    m.dUserScreenDate = 0
    nNum = Bars.Size
    
    ' load aXdate and aXbar arrays
    iX = 0
    iBar = 0
    m.aXBar.Size = 0
    m.aXdate.Size = 0
    ' pre-allocate enough space (with some extra)
    m.aXBar.Reserve nNum * 1.1
    m.aXdate.Reserve nNum * 1.1
    ' init TDOM and TDOY stuff
    m.aTDOM.Size = 0
    m.aTDOY.Size = 0
    m.aTDOW.Size = 0
    nNextMonth = 0
    nNextYear = 0
    nNextWeek = 0
    nTDOYctr = 0
    nTDOMctr = 0
    nTDOWctr = 0
    If m.nPeriodicity = ePRD_Days + 1 Then
        m.aTDOM.Size = nNum
        m.aTDOY.Size = nNum
        m.aTDOW.Size = nNum
        ' get first date of next year and next month and next week
        ' (can't do anything with the counters until then)
        i = JulToLong(m.Bars(eBARS_DateTime, 0), True)
        nNextYear = JulFromLong((i \ 10000 + 1) * 10000 + 101)
        i = i \ 100 + 1
        If i Mod 100 > 12 Then i = i + 88
        nNextMonth = JulFromLong(i * 100 + 1)
        nNextWeek = Int(m.Bars(eBARS_DateTime, 0) / 7) * 7 + 9
    ElseIf m.nPeriodicity = ePRD_Weeks + 1 Then
        m.aTDOY.Size = nNum
        ' get first date of next year
        ' (can't do anything with the counters until then)
        i = JulToLong(m.Bars(eBARS_DateTime, 0), True)
        nNextYear = JulFromLong((i \ 10000 + 1) * 10000 + 101)
        nNextWeek = Int(m.Bars(eBARS_DateTime, 0) / 7) * 7 + 9
    End If
    ' for each bar in m.Bars
    dMinuteBar = 0
    If GetPeriodType(m.nPeriodicity) = ePRD_Minutes Then
        dMinuteBar = GetPeriodsPerBar(m.nPeriodicity) / 1440#
        dCrossoverOffset = m.Bars.Prop(eBARS_CrossoverTime) / 1440#
        If dCrossoverOffset >= 1 Or dCrossoverOffset < 0 Then dCrossoverOffset = 0
        dSessionStart = m.Bars.Prop(eBARS_StartTime) / 1440#
        If dSessionStart >= 1 Or dSessionStart < 0 Then dSessionStart = 0
        dSessionEnd = m.Bars.Prop(eBARS_EndTime) / 1440#
        If dSessionEnd >= 1 Or dSessionEnd < 0 Then dSessionEnd = 0
'dSessionStart = 0
    Else
        dCrossoverOffset = 0
    End If
    Do While iBar < nNum
        dBarDate = gdGetNum(hBarsDateArray, iBar)
        If dBarDate = kNullData Then
            Exit Do '(no more good dates)
        ElseIf m.nPeriodicity = ePRD_Days + 1 Then
            ' add empty daily bars (holidays)
            If m.bShowEmptyBars And (iX > 0) Then
                dCurDate = dCurDate + 1
                Do While dBarDate > dCurDate
                    ' make sure not a weekend
                    If (Int(dCurDate) - 2) Mod 7 <= 4 Then
                        ' add a bar
                        'm.aXdate(iX) = dCurDate
                        gdSetNum hXdateArray, iX, gdFixDateTime(dCurDate)
                        iX = iX + 1
                    End If
                    dCurDate = dCurDate + 1
                Loop
            End If
        ElseIf m.nPeriodicity = ePRD_Weeks + 1 Then
            ' add empty weeks (e.g. for Gann contracts)
            If m.bShowEmptyBars And (iX > 0) Then
                dCurDate = dCurDate + 7
                Do While dBarDate > dCurDate
                    ' add a bar
                    'm.aXdate(iX) = dCurDate
                    gdSetNum hXdateArray, iX, gdFixDateTime(dCurDate)
                    iX = iX + 1
                    dCurDate = dCurDate + 7
                Loop
            End If
        ElseIf m.nPeriodicity = ePRD_Months + 1 Then
            ' add empty months (e.g. for Gann contracts)
            If m.bShowEmptyBars And (iX > 0) Then
                dCurDate = dCurDate + 45
                i = Month(dCurDate)
                Do While dBarDate > dCurDate
                    ' walk backwards looking for the last day of the month
                    dCurDate = dCurDate - 1
                    If Month(dCurDate) <> i Then
                        ' make sure not a weekend
                        If (Int(dCurDate) - 2) Mod 7 <= 4 Then
                            'm.aXdate(iX) = dCurDate
                            gdSetNum hXdateArray, iX, gdFixDateTime(dCurDate)
                            iX = iX + 1
                            ' setup for the next month
                            dCurDate = dCurDate + 45
                            i = Month(dCurDate)
                        End If
                    End If
                Loop
            End If
        ElseIf m.nPeriodicity = ePRD_Quarters + 1 Then
            ' add empty quarters (e.g. for Gann contracts)
            If m.bShowEmptyBars And (iX > 0) Then
                dCurDate = dCurDate + 105
                i = Month(dCurDate)
                Do While dBarDate > dCurDate
                    ' walk backwards looking for the last day of the month
                    dCurDate = dCurDate - 1
                    If Month(dCurDate) <> i Then
                        ' make sure not a weekend
                        If (Int(dCurDate) - 2) Mod 7 <= 4 Then
                            'm.aXdate(iX) = dCurDate
                            gdSetNum hXdateArray, iX, gdFixDateTime(dCurDate)
                            iX = iX + 1
                            ' setup for the next month
                            dCurDate = dCurDate + 105
                            i = Month(dCurDate)
                        End If
                    End If
                Loop
            End If
        ElseIf dMinuteBar > 0 And dBarDate <= gdFixDateTime(m.dLastActualDateOfData + dMinuteBar) Then
            ' add empty minute bars
            'If (iX > 0) Then
            If m.bShowEmptyBars And (iX > 0) Then
                ' check if just past the crossover
                If Int(dCurDate - dCrossoverOffset) <> Int(dBarDate - dCrossoverOffset) Then
                    ' for combined sessions, allow one empty bar between the
                    ' day and evening session of same date just for visual purposes
                    If Int(dBarDate) = Int(dCurDate) Then
                        dTempDate = Int(dBarDate - dCrossoverOffset) + dSessionStart ' dCrossoverOffset
                        'If dTempDate > dCurDate Then
                        If 0 Then
                            dCurDate = dTempDate
                            gdSetNum hXdateArray, iX, dCurDate
                            iX = iX + 1
                        End If
                    End If
                    ' get logical trading day for this bar
                    dTempDate = Int(dBarDate)
                    If dBarDate - dTempDate >= dCrossoverOffset And dCrossoverOffset > 0 Then
                        dTempDate = dTempDate + 1
                    End If
                    ' get start of session for logical trading day
                    If dSessionStart > dSessionEnd Then
                        dTempDate = dTempDate - 1 + dSessionStart
                    Else
                        dTempDate = dTempDate + dSessionStart
                    End If
                    dTempDate = gdFixDateTime(dTempDate + dMinuteBar)
                    If dTempDate > dCurDate Then dCurDate = dTempDate
                Else
                    ' go to next bar
                    dCurDate = dCurDate + dMinuteBar
                End If
                ' add empty bars as long as we're at least a full bar before the next one
                Do While dBarDate >= gdFixDateTime(dCurDate + dMinuteBar)
                    ' add a bar
                    'm.aXdate(iX) = dCurDate
                    gdSetNum hXdateArray, iX, gdFixDateTime(dCurDate)
                    iX = iX + 1
                    dCurDate = gdFixDateTime(dCurDate + dMinuteBar)
                Loop
            End If
        End If
        dCurDate = dBarDate
        'm.aXbar(iX) = iBar
        gdSetNum hXbarArray, iX, iBar
        'm.aXdate(iX) = dBarDate
        gdSetNum hXdateArray, iX, dBarDate
        
        ' get trading day of month/year
        If nNextMonth <> 0 Then
            If dBarDate >= nNextMonth Then
                nTDOMctr = 1
                ' get first date of next month
                i = JulToLong(dBarDate, True)
                i = i \ 100 + 1
                If i Mod 100 > 12 Then i = i + 88
                nNextMonth = JulFromLong(i * 100 + 1)
            ElseIf nTDOMctr > 0 Then
                nTDOMctr = nTDOMctr + 1
            End If
            If nTDOMctr > 0 Then
                gdSetNum hTDOMArray, iBar, nTDOMctr
            End If
        End If
        If nNextYear <> 0 Then
            If dBarDate >= nNextYear Then
                nTDOYctr = 1
                ' get first date of next year
                i = JulToLong(dBarDate, True)
                nNextYear = JulFromLong((i \ 10000 + 1) * 10000 + 101)
            ElseIf nTDOYctr > 0 Then
                nTDOYctr = nTDOYctr + 1
            End If
            If nTDOYctr > 0 Then
                gdSetNum hTDOYArray, iBar, nTDOYctr
            End If
        End If
        If nNextWeek <> 0 Then
            If dBarDate >= nNextWeek Then
                nTDOWctr = 1
                ' get first date of next week
                nNextWeek = Int(dBarDate / 7) * 7 + 9
            ElseIf nTDOWctr > 0 Then
                nTDOWctr = nTDOWctr + 1
            End If
            If nTDOWctr > 0 Then
                gdSetNum hTDOWArray, iBar, nTDOWctr
            End If
        End If
        
        iX = iX + 1
        iBar = iBar + 1
    Loop
    m.aXBar.Size = iX
    m.aXdate.Size = iX
    
    'copy the aXDate array to structure so grapheng can remove un-needed date values to make variable column widths
    If m.bProfileChart Then
        i = gdGetSize(m.profileChartInfo.gdhXDateCopy)
        If i = 0 Or gdGetNum(m.profileChartInfo.gdhXDateCopy, 0) < m.aXdate(0) Then
            gdClear m.profileChartInfo.gdhXDateCopy, 1
            For i = 0 To iX - 1
                gdSetNum m.profileChartInfo.gdhXDateCopy, i, gdGetNum(hXdateArray, i)
            Next
'            m.profileChartInfo.RTOnOff = -1
        ElseIf RedoMode = eRedo7_ReloadRT Then
            If gdGetNum(hXdateArray, iX - 1) > gdGetNum(m.profileChartInfo.gdhXDateCopy, i - 1) Then
                'true right after streaming start
                i = i - 1
                gdBinarySearch hXdateArray, gdGetNum(m.profileChartInfo.gdhXDateCopy, i), j, eGdSort_Default, iX, 0
                For i = j To iX - 1
                    gdSetNum m.profileChartInfo.gdhXDateCopy, i, gdGetNum(hXdateArray, i)
                Next
            End If
        ElseIf RedoMode = eRedo9_ReloadData Then
            If gdGetNum(m.profileChartInfo.gdhXDateCopy, i - 1) > m.aXdate(iX - 1) Then
                i = i - 1
                Do While gdGetNum(m.profileChartInfo.gdhXDateCopy, i) > gdGetNum(hXdateArray, iX - 1)
                    gdDeleteItems m.profileChartInfo.gdhXDateCopy, i, 1
                    i = i - 1
                Loop
            End If
        End If
    End If
    
    ' find last bar with good data
    '(store as chart's X as opposed to bar# since easier
    ' to convert X to bar# than other way around)
    m.nLastGoodDataBarX = -1
    For iX = m.aXBar.Size - 1 To 0 Step -1
        i = m.aXBar(iX)
        If i >= 0 Then
            If m.Bars(eBARS_Close, i) <> kNullData Then
                m.nLastGoodDataBarX = iX
                Exit For
            End If
        End If
    Next
    
    ' reset horizontal scroll bar
    With m.hsb
        .Max = m.aXBar.Size - 1
        If m.dLastScreenDate > 0 Then
            m.aXdate.BinarySearch m.dLastScreenDate, iX
            If iX > .Max Then                                       'Or iX >= m.nLastGoodDataBarX Then
                If RedoMode >= eRedo9_ReloadData Then
                    iX = m.nLastGoodDataBarX + m.nBlankBars
                    If iX > .Max Then iX = .Max
                Else
                    iX = .Max
                End If
            ElseIf iX < .Min Then
                iX = .Min
            End If
        ElseIf bSeasonal Then
            iX = .Max
        Else
            iX = m.nLastGoodDataBarX + m.nBlankBars
            If iX > .Max Then
                iX = .Max
            ElseIf Not m.Form Is Nothing Then
                If Not m.Form.GameMode Is Nothing Then
                    iX = .Max       '5118
                End If
            End If
        End If
        .Value = iX                                                 '.Max ' set scrollbar to end
    End With
            
    AddSpeed "Load Data (" & CStr(m.Bars.Size) & " bars)"
    
    ' load rolls and splits table
    Set m.RollsTable = GetRollsTable(m.nSymbolID)
    If m.RollsTable.NumRecords > 0 Then
        AddSpeed "Load Rolls"
        AddForecastRoll
        AddSpeed "ForecastRoll"
    End If
       
    'save off real data values
    Set m.aWhatIfSave = New cGdArray
    m.aWhatIfSave.Create eGDARRAY_Doubles, 5
    i = LastGoodDataBar(False)
    m.aWhatIfSave(0) = m.Bars(eBARS_Open, i)
    m.aWhatIfSave(1) = m.Bars(eBARS_High, i)
    m.aWhatIfSave(2) = m.Bars(eBARS_Low, i)
    m.aWhatIfSave(3) = m.Bars(eBARS_Close, i)
    m.aWhatIfSave(4) = m.Bars.Prop(eBARS_LastTickTime)
    
    If m.nSymbolID = 50 Then LoadEwave

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.LoadData"
End Sub

Private Sub SpliceReplayTicks(Bars As cGdBars)
On Error GoTo ErrSection:

    Dim i&, dDataTime#, dTime#, nSessionDate&, nFirstDate&, strSymbol$, strTrade$, dPrice#
    Dim SessionTicks As cGdBars, Ticks As cGdBars, Daily As cGdBars
    Dim dReplayDateFixed#, dGameInterval#, dCustomTime#
                
    strSymbol = Bars.Prop(eBARS_Symbol)
                
    ' get current session date
    If Not m.Form.GameMode Is Nothing Then
        dDataTime = m.Form.GameMode.GameDataTime
        dGameInterval = m.Form.GameMode.GameInterval
    End If
    nSessionDate = Bars.SessionDateForTime(dDataTime, False)
                
    ' chop back to end of last complete session
    For i = Bars.Size - 1 To 0 Step -1
        If Bars.SessionDate(i) < nSessionDate Then
            Bars.Size = i + 1
            Exit For
        ElseIf i = 0 Then
            'The function EnougData in the game config form calls g.SymbolPool.TickFirstDate or g.SymbolPool.EodFirstDate
            'to get date of first available data. On very rare occasions the TickFirstDate or EodFirstDate returns a date
            'that is earlier than actual available data and we end up here.
            Bars.Size = 1           'setting this to zero will cause TradeNav to hang
        End If
    Next
        
    ' Append current session (if has started)
    dTime = gdFixDateTime(nSessionDate + Bars.Prop(eBARS_StartTime) / 1440#)
    If Bars.Prop(eBARS_EndTime) < Bars.Prop(eBARS_StartTime) Then
        dTime = dTime - 1 '(starts prior evening)
    End If
    If dDataTime >= dTime And dDataTime > 0 Then
        ' load ticks for current session
        Set SessionTicks = m.ReplayTicks(strSymbol)
        If SessionTicks Is Nothing Then
            Set SessionTicks = New cGdBars
            SetBarProperties SessionTicks, strSymbol
            m.ReplayTicks.Add SessionTicks, strSymbol
        End If
        If dGameInterval >= 1 Then
            ' load daily data from after the last complete bar
            For i = Bars.Size - 1 To 0 Step -1
                If Int(Bars(eBARS_DateTime, i)) < nSessionDate Then
                    nFirstDate = Int(Bars(eBARS_DateTime, i)) + 1
                    Exit For
                End If
            Next
            i = DM_GetBars(SessionTicks, strSymbol, , nFirstDate, nSessionDate)
            Set Ticks = SessionTicks.MakeCopy
            i = Ticks.Size - 1
            ' if not end-of-day, only get the session's open
            If i >= 0 And dDataTime < Ticks(eBARS_DateTime, i) + 0.5 Then
                Ticks(eBARS_Close, i) = Ticks(eBARS_Open, i)
                Ticks(eBARS_Low, i) = Ticks(eBARS_Open, i)
                Ticks(eBARS_High, i) = Ticks(eBARS_Open, i)
                If Ticks(eBARS_Vol, i) > 0 Then
                    Ticks(eBARS_Vol, i) = 0
                End If
                If Ticks(eBARS_ContVol, i) > 0 Then
                    Ticks(eBARS_ContVol, i) = 0
                End If
            End If
        Else
            ' load intraday data for session date
            If SessionTicks.Size = 0 Or SessionTicks.SessionDate(0) <> nSessionDate Or SessionTicks.Prop(eBARS_PeriodType) <> ePRD_EachTick Then        '4748
                SessionTicks.Size = 0
                i = DM_GetBars(SessionTicks, strSymbol, ePRD_EachTick, nSessionDate, nSessionDate)
                ' if getting ticks for daily chart, verify the correct daily open and close
                If SessionTicks.Size > 0 And Not Bars.IsIntraday Then
                    Set Daily = New cGdBars
                    SetBarProperties Daily, strSymbol
                    i = DM_GetBars(Daily, strSymbol, 0, nSessionDate, nSessionDate)
                    If Daily.Size > 0 Then
                        ' if not match, replace first tick with daily open
                        If Daily(eBARS_Open, 0) <> SessionTicks(eBARS_Close, 0) Then
                            SessionTicks(eBARS_Close, 0) = Daily(eBARS_Open, 0)
                        End If
                        ' if not match, replace last tick with daily close
                        If Daily(eBARS_Close, 0) <> SessionTicks(eBARS_Close, SessionTicks.Size - 1) Then
                            SessionTicks(eBARS_Close, SessionTicks.Size - 1) = Daily(eBARS_Close, 0)
                        End If
                    End If
                End If
            End If
            ' chop off ticks after DateTime
            Set Ticks = SessionTicks.MakeCopy
            
            If m.Bars.IsIntraday Then
                'check custom start/stop times  (4433)
                If m.Bars.Prop(eBARS_StartTime) <> m.Bars.Prop(eBARS_DefaultStartTime) Then
                    i = Int(Ticks(eBARS_DateTime, 0))
                    If m.Bars.Prop(eBARS_DefaultStartTime) > m.Bars.Prop(eBARS_DefaultEndTime) And m.Bars.Prop(eBARS_StartTime) < m.Bars.Prop(eBARS_DefaultStartTime) Then
                        'overnight symbols
                        dCustomTime = i + 1 + m.Bars.Prop(eBARS_StartTime) / 1440#
                    Else
                        dCustomTime = i + m.Bars.Prop(eBARS_StartTime) / 1440#
                    End If
                    i = Ticks.FindDateTime(dCustomTime)
                    If i > 0 Then Ticks.DeleteFirstBars i
                End If
                If m.Bars.Prop(eBARS_EndTime) <> m.Bars.Prop(eBARS_DefaultEndTime) Then
                    dCustomTime = Int(Ticks(eBARS_DateTime, 0)) + m.Bars.Prop(eBARS_EndTime) / 1440#
                    i = Ticks.FindDateTime(dCustomTime)
                    If i > 0 Then Ticks.Size = i
                End If
                dReplayDateFixed = gdFixDateTime(dDataTime)
                ' only cut ticks off mid-bar if playing (i.e. if not forward/rewind buttons hit)    -   5123
                If m.Form.GameMode.UserForwardBack = 0 Then
                    'if replay date <= first bar then show only first bar, if replay date >= last bar then show all bars
                    If dReplayDateFixed < Ticks(eBARS_DateTime, Ticks.Size - 1) Then
                        gdBinarySearch Ticks.ArrayHandle(eBARS_DateTime), dReplayDateFixed, i, eGdSort_Default, 0, Ticks.Size - 1
                        If i > 0 And i < Ticks.Size Then
                            For i = i - 1 To Ticks.Size - 1
                                If m.Form.GameMode.KeepAtEndOfBar Then
                                    If Ticks(eBARS_DateTime, i) > dReplayDateFixed Then
                                        Ticks.Size = i
                                        Exit For
                                    End If
                                ElseIf Ticks(eBARS_DateTime, i) >= dReplayDateFixed Then
                                    Ticks.Size = i + 1
                                    Exit For
                                End If
                            Next
                        ElseIf i = 0 Then
                            Ticks.Size = 1
                        End If
                    End If
                End If
            Else
                For i = Ticks.Size - 1 To 0 Step -1
                    If Ticks(eBARS_DateTime, i) < dDataTime Then
                        If Not m.Form.GameMode Is Nothing Then
                            dReplayDateFixed = gdFixDateTime(dDataTime + m.Form.GameMode.GameInterval)
                        End If
                        If i < Ticks.Size - 1 And Ticks(eBARS_DateTime, i + 1) < dReplayDateFixed Then
                            Ticks.Size = i + 2
                        Else
                            Ticks.Size = i + 1
                        End If
                        Exit For
                    ElseIf i = 0 Then
                        Ticks.Size = 1 '(leave first tick)
                    End If
                Next
            End If
        End If
        
        ' compress to same periodicity and append
        Bars.BuildBars Bars.Prop(eBARS_PeriodicityStr), Ticks.BarsHandle, True
        
        ' if forward/rewind button hit, then chop off at a specific bar -   5123
        If m.Bars.IsIntraday And m.Form.GameMode.UserForwardBack <> 0 Then
            gdBinarySearch Bars.ArrayHandle(eBARS_DateTime), dReplayDateFixed, i, eGdSort_Default, 0, Bars.Size - 1     '5943
            If m.Form.GameMode.UserForwardBack > 0 And i < Bars.Size Then
                Bars.Size = i + 1
            ElseIf i > 0 Then
                Bars.Size = i
            End If
        End If
        
        Set Ticks = Nothing
        Set Daily = Nothing
        Set SessionTicks = Nothing
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.SpliceReplayTicks"
End Sub

Private Function MaxInitialNulls(aResults As cGdArray) As Long
On Error GoTo ErrSection:

    Dim iArray&, hArray&, iSize&, iMaxNulls&, i&
    
    For iArray = 0 To aResults.Size - 1
        hArray = aResults(iArray)
        If hArray <> 0 And hArray <> kNullData Then
            iSize = gdGetSize(hArray)
            For i = 0 To iSize - 1
                If gdGetNum(hArray, i) <> kNullData Then
                    If i > iMaxNulls Then iMaxNulls = i
                    Exit For
                End If
            Next
        End If
    Next

    MaxInitialNulls = iMaxNulls

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.MaxInitialNulls"
End Function

Private Sub LoadIndicators()
On Error GoTo ErrSection:

    Dim idx&, nCount&, nNum&, nX&, rc&, strText$, i&, nSymbolID&, nMPB&, iBars&, nPeriodicity&
    Dim dEngineTime#, dEngineSetupTime#
    Dim strSymbol$, strPeriod$, iPos&, hBars&, hArray&, strMarket$
    Dim bReloaded As Boolean, bAdd As Boolean, bRebuildCodedText As Boolean
    Dim Indicator As cIndicator, IndParent As cIndicator
    Dim Pane As cPane

    Dim nLastGoodBar&, hArray2&, dVal#, dVal2#, nSkipBars&, nMaxNulls&, iArray&, dDiffAllowed#
    Dim bNonNullFound As Boolean
    Dim aTempResults As New cGdArray
    Dim aSmallerBars As New cGdBars
    
    Dim astrParms As New cGdArray, astrBarNames As New cGdArray
    Dim astrExpr As New cGdArray, aArrayOfBars As New cGdArray
    Dim aArrayOfResults As New cGdArray
    Dim strCompMarkets As String
    
    Static astrBarNames2 As cGdArray
    Static astrExpr2 As cGdArray
    Static aArrayOfResults2 As cGdArray
    Static strCompMarkets2 As String
    Static dLastPrice As Double
    Static bInitExprDone As Boolean
    
    Dim barsOther As New cGdBars
    Dim aCompSymbolIdx As New cGdArray
        
    aCompSymbolIdx.Create eGDARRAY_Longs
    
    astrParms.Create eGDARRAY_Strings
    astrExpr.Create eGDARRAY_Strings
    
    aArrayOfResults.Create eGDARRAY_Longs
    aArrayOfBars.Create eGDARRAY_Longs
    
    ' make sure the primary bars is the first item in this array
    aArrayOfBars(0) = m.Bars.BarsHandle
    astrBarNames(0) = "Market1"
    
    nX = m.aXBar.Size
    nNum = m.Bars.Size
    nCount = m.Tree.Count
    For idx = 1 To nCount
        If m.Tree.NodeLevel(idx) = 0 Then
            Set Pane = m.Tree(idx)
        Else
            Set Indicator = m.Tree(idx)
            With Indicator
                If Pane.Display Then
                    If .DataType = eINDIC_BarData Then
                        ' load comparison symbol
                        If UCase(m.Tree.Key(idx)) = "PRICE" Then
                            .Expression = "Market1"
                        Else
                            .Expression = Chr(34) & UCase(.Name) & Chr(34) ' "Market" & Str(aCompSymbolIdx.Size)
                            bAdd = .Display
                            If Not bAdd Then
                                For i = idx + 1 To m.Tree.RelativeIndex(idx, eTREE_LastDescendant)
                                    If m.Tree(i).Display Then
                                        bAdd = True
                                        Exit For
                                    End If
                                Next
                            End If
                            If bAdd Then
                                strText = Chr(34) & UCase(.Name) & "," & GetPeriodStr(m.nPeriodicity) & Chr(34)
                                For i = 0 To astrBarNames.Size - 1
                                    If astrBarNames(i) = strText Then
                                        bAdd = False
                                        If UCase(.Name) = "HEIKIN-ASHI" Then
                                            ' can only show 1 set of H.A. bars
                                            .Display = False
                                        End If
                                        Exit For
                                    End If
                                Next
                                If bAdd Then
                                    aCompSymbolIdx.Add idx
                                    astrBarNames.Add strText
                                    strCompMarkets = strCompMarkets & "," & UCase(.Name)
                                End If
                            End If
                        End If
                        .BuildText Me, "", bRebuildCodedText
                    ElseIf .DataType = eINDIC_Array Or .DataType = eINDIC_BooleanArray Then
                        strMarket = ""
                        If m.Tree.NodeLevel(idx) > 1 Then
                            If .DataType = eINDIC_BooleanArray Then
                                i = m.Tree.AncestorIndex(idx, 1)
                                Set IndParent = m.Tree(i)
                                If IndParent.DataType = eINDIC_BarData Then
                                    strMarket = IndParent.Expression
                                End If
                            ElseIf .DataType = eINDIC_Array Then
                                ' linked indicator
                                i = InStr(UCase(.Parm(1)), " OF ")
                                If i > 0 Then
                                    .Parm(1) = Trim(Left(.Parm(1), i))
                                End If
                                i = m.Tree.RelativeIndex(idx, eTREE_Parent)
                                Set IndParent = m.Tree(i)
                                If IndParent.DataType <> eINDIC_BarData Then
                                    strText = IndParent.Expression
                                ElseIf ParmDefaultIsBarsArray(.Parm(1)) Then
                                    strText = .Parm(1) & " of " & IndParent.Expression
                                Else
                                    strText = "Close of " & IndParent.Expression
                                End If
                                If strText = "" Then strText = "Close of Market1"
                                If .ParmType(2) = 4 And Not IsDigit(.ParmDefault(2)) Then '(fix for MovingSum aardvark issues)
                                    ' if both of first 2 parms are "arrays" and 2nd parm default is not a number
                                    ' then link 1st parm to grandparent and 2nd parm to parent
                                    .Parm(2) = strText
                                    strText = ""
                                    If m.Tree.NodeLevel(i) > 1 Then
                                        i = m.Tree.RelativeIndex(i, eTREE_Parent)
                                        Set IndParent = m.Tree(i)
                                        If IndParent.DataType <> eINDIC_BarData Then
                                            strText = IndParent.Expression
                                        ElseIf ParmDefaultIsBarsArray(.Parm(1)) Then
                                            strText = .Parm(1) & " of " & IndParent.Expression
                                        Else
                                            strText = "Close of " & IndParent.Expression
                                        End If
                                    End If
                                    If strText = "" Then strText = "Close of Market1"
                                    .Parm(1) = strText
                                Else
                                    ' just link 1st parm to parent
                                    .Parm(1) = strText
                                End If
                            End If
                        End If
                        
                        .BuildText Me, strMarket, bRebuildCodedText
                        If .Display Or .IsAlert Or .IsWoodInd(False, False) Or .DisplayType = eINDIC_ClusterPrice Then
                            strText = Trim(.CodedText)
                            If Len(strText) = 0 Then
                                .ChartLabel = "INVALID"
                                gdClear .Data.ArrayHandle, False
                            Else
                                astrExpr.Add strText
                                aArrayOfResults.Add .Data.ArrayHandle
                            End If
                        End If
                    End If
                End If
            End With
        End If
    Next
    Set Indicator = Nothing
    Set IndParent = Nothing
    Set Pane = Nothing

    AddSpeed "Code Indicators"
    
    ' get symbols/periods from expressions (puts them into BarNames and fixes Expressions)
    MarketsInExpressions astrExpr, 0, False, astrBarNames, Nothing, GetPeriodStr(m.nPeriodicity)
        
    ' see if we need to recalculate: if any array, secondary symbols,
    ' or last price has changed since the previous call
    If dLastPrice <> m.Bars(eBARS_Close, Me.LastGoodDataBar(False)) Then
        dLastPrice = m.Bars(eBARS_Close, Me.LastGoodDataBar(False))
        RedoMode = eRedo5_RecalcInd
    End If
    If strCompMarkets <> strCompMarkets2 Then
        RedoMode = eRedo5_RecalcInd
        'g.bSkipSetChartFocus = true means charts are being restored from chart page switching or program startup
        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        bInitExprDone = False
    ElseIf Not ArraysAreSame(astrBarNames, astrBarNames2) Then
        RedoMode = eRedo5_RecalcInd
        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        bInitExprDone = False
    ElseIf Not ArraysAreSame(astrExpr, astrExpr2) Then
        RedoMode = eRedo5_RecalcInd
        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        bInitExprDone = False
    ElseIf Not ArraysAreSame(aArrayOfResults, aArrayOfResults2) Then
        RedoMode = eRedo5_RecalcInd
        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        bInitExprDone = False
    ElseIf RedoMode >= eRedo9_ReloadData Then
        bInitExprDone = False
    ElseIf RedoMode = eRedo6_ReloadInd Then
        bInitExprDone = False '(reload since a function may have been edited)
    End If
    
    If RedoMode >= eRedo5_RecalcInd Then
        ' see if we need to reload the secondary data
        If RedoMode >= eRedo9_ReloadData Or Not ArraysAreSame(astrBarNames, astrBarNames2) Or strCompMarkets <> strCompMarkets2 Then
            bReloaded = True
        
            ' first clear the existing secondary data bars
            ReDim m.OtherBars(astrBarNames.Size - 1) As cGdBars

            If astrBarNames.Size > 1 Then
                For i = 1 To astrBarNames.Size - 1
                    Set barsOther = Nothing
                    nSymbolID = 0
                    strSymbol = UCase(astrBarNames(i))
                    If Left(strSymbol, 1) = Chr(34) And Right(strSymbol, 1) = Chr(34) Then
                        strPeriod = Parse(Replace(strSymbol, Chr(34), ""), ",", 2)
                        strSymbol = Parse(Replace(strSymbol, Chr(34), ""), ",", 1)
                        If Right(strSymbol, 1) = ":" Then
                            strSymbol = strSymbol & Me.Symbol
                        End If
                        If Len(strSymbol) = 0 Then
                            nSymbolID = m.nSymbolID
                        Else
                            nSymbolID = GetSymbolID(strSymbol)
                        End If
                    ElseIf strSymbol = "DAILY" Or strSymbol = "WEEKLY" Or strSymbol = "MONTHLY" Then
                        strPeriod = strSymbol
                        nSymbolID = m.nSymbolID
                    ElseIf strSymbol = "GC" Or strSymbol = "TQ" Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol & "-067")
                    ElseIf UCase(strSymbol) <> "DEFAULT" Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
                    ElseIf Left(PrimaryFutureBase(g.SymbolPool.SymbolForID(m.nSymbolID)), 3) = "GC-" Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol("TQ-067")
                    Else
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol("GC-067")
                    End If
                    If nSymbolID <> 0 Then
                        ' load bars (can only do up to daily prior to appending real-time)
                        If Len(strPeriod) = 0 Then
                            strPeriod = GetPeriodStr(m.nPeriodicity)
                        End If
                        nPeriodicity = GetPeriodicity(strPeriod)
                        If nPeriodicity > ePRD_Days + 1 Then
'                            nPeriodicity = ePRD_Days + 1
                        ElseIf nPeriodicity >= ePRD_IntRenko And nPeriodicity < ePRD_IntVol Then
                            nPeriodicity = ePRD_Minutes + 1
                        End If
                        Set barsOther = New cGdBars
                        If nPeriodicity >= ePRD_Days Then 'And nSymbolID = m.nSymbolID Then
                            rc = DM_GetBars(barsOther, nSymbolID, nPeriodicity, , , , , m.bUnsplit)
                        ElseIf m.nPeriodicity >= ePRD_Days Then
                            ' TLB 9/30/2009: it is bogus to load intraday data for secondary bars when the
                            ' primary bars is not intraday (and it will bog down forever loading all that data!)
                            barsOther.Size = 0
                        Else
                            rc = DM_GetBars(barsOther, nSymbolID, nPeriodicity, m.Bars(eBARS_DateTime, 0), m.Bars(eBARS_DateTime, m.Bars.Size - 1), , , m.bUnsplit)
                        End If
                        ' Image Server: merge Rex data
                        If Len(m.strRexDataFile) > 0 Then
                            AppendRexData barsOther ', m.Bars(eBARS_DateTime, LastGoodDataBar(False))
                        End If
                    End If
                    Set m.OtherBars(i) = barsOther
                Next
                AddSpeed Str(astrBarNames.Size - 1) & " other symbols loaded"
            End If
        End If

        ' final stuff for the secondary data bars
        For iBars = 1 To UBound(m.OtherBars)
            Set barsOther = m.OtherBars(iBars)
            
            ' see if need to refresh the real-time portion
            If bReloaded Or RedoMode >= eRedo7_ReloadRT Then
                ' if Heikin-Ashi bars, recreate them now
                If InStr(UCase(astrBarNames(iBars)), "HEIKIN-ASHI") > 0 Then
                    Set barsOther = m.Bars.CreateHeikinAshi
                    barsOther.Prop(eBARS_CustomString) = "HEIKIN-ASHI" '(to flag this for realtime updating)
                    Set m.OtherBars(iBars) = barsOther
                Else
                    ' chop any blank bars at end
                    For i = barsOther.Size - 1 To 0 Step -1
                        If barsOther(eBARS_Close, i) <> kNullData Then
                            barsOther.Size = i + 1
                            Exit For
                        End If
                    Next
                    ' then add forecast bars
                    barsOther.AddForecastBars m.nForecastBars
                End If
                
                ' setup for realtime updating
                If Not m.bDisableRT Then
'bReloaded = True
                    ' TLB 12/21/2007: can skip this under the specific condition:
                    ' if other indicator is "Daily" of the same symbol and this chart is intraday
                    ' (this helps speed up when new bar starts -- e.g. Woodies templates with pivot points)
                    If bReloaded Or m.nSymbolID <= 0 Or barsOther.Prop(eBARS_SymbolID) <> m.nSymbolID _
                            Or m.nPeriodicity >= ePRD_Days Or barsOther.Prop(eBARS_Periodicity) <> ePRD_Days + 1 Then
                            
                        If Screen.ActiveForm Is frmSymbolGrid Then
                            g.RealTime.AddTickBuffer barsOther, False
                        Else
                            g.RealTime.AddTickBuffer barsOther, True
                        End If
                        ' don't splice bars for Heikin-Ashi (don't need to, and will mess them up)
                        If m.bRT And barsOther.Prop(eBARS_CustomString) <> "HEIKIN-ASHI" Then
                            g.RealTime.SpliceBars barsOther
                            ' TBL 2/25/2008: need to add forecast bars back on so future bars will be Null
                            barsOther.AddForecastBars m.nForecastBars
                        End If
                    End If
                End If
        
                ' now compress bars
                ' TLB 12/21/2007: this is messing up "of Daily" on intraday charts other than minute bars
                ' (e.g. Pivot Points on breakout bar charts)
                If 0 Then
                    Select Case GetPeriodType(m.nPeriodicity)
                    Case ePRD_Minutes, ePRD_Days, ePRD_Weeks, ePRD_Months, ePRD_Quarters, ePRD_Years
                        ' for fixed timeframes, compress to that timeframe before
                        ' calculating the indicators, then align to Market1 later
                        If barsOther.Prop(eBARS_Periodicity) <> m.nPeriodicity Then
                            barsOther.BuildBars GetPeriodStr(m.nPeriodicity)
                        End If
                    Case Else
                        ' for non-fixed timeframes, align to Market1 before calculating
                        ' indicators (but flat-line instead of empty bars so indicators work)
                        rc = gdAlignBars(barsOther.BarsHandle, m.Bars.BarsHandle, 1)
                    End Select
                End If
            End If
                
            ' build array of handles to bars
            aArrayOfBars.Add barsOther.BarsHandle
        Next
                       
hBars = 0
If 0 Then
    If FileExist(App.Path & "\Expr.tst") Then
        astrBarNames.Add Chr(127) & Chr(127) & "Profit" & vbTab & "4" & vbTab & "Market1"
        hBars = gdCreateArray(eGDARRAY_Doubles, m.Bars.Size)
        aArrayOfBars.Add hBars
        For i = 0 To m.Bars.Size - 1
            gdSetNum hBars, i, i + 0.5
        Next
        
        'aArrayOfBars.Size = 1
        'astrBarNames.Size = 1
    End If
End If
                       
        'KillFile App.Path & "\chk\Coded.chk"
        If IsIDE Then
            astrExpr.ToFile App.Path & "\chk\Coded.chk"
            astrBarNames.ToFile App.Path & "\chk\BarNames.chk"
            aArrayOfBars.ToFile App.Path & "\chk\BarNames.chk", True
        End If
                

If IsIDE Then
    If m.iTryLastBarRecalc = 0 Then
        'm.iTryLastBarRecalc = 1
    End If
    'bInitExprDone = False
    'm.iTryLastBarRecalc = 0 'clear flag
End If

        ' setup to run Expressions
        rc = 0
        astrParms.Size = 0 '(use defaults)
        astrParms(0) = "Chart" & Str(m.Form.hWnd)
        If Not bInitExprDone Then
            ' need to reinit the expression set
            dEngineSetupTime = gdTickCount(False)
            bInitExprDone = SetupExpressions(astrParms, astrBarNames, astrExpr)
            dEngineSetupTime = gdTickCount(False) - dEngineSetupTime
            m.iTryLastBarRecalc = 0 'clear flag
        ElseIf m.eRedoMode >= eRedo7_ReloadRT Then
            m.iTryLastBarRecalc = 0 'clear flag
        End If

        ' TLB 5/12/2006: To SPEED-UP calculations when just updating
        ' ticks on same bar from realtime streaming ...
        '   - use the results stored for calc of full data set
        '   - count max # initial null's
        '   - do calcs for smaller data set
        '   - verify same ending results (e.g. 20 bars prior to last data bar)
        '   - if matches, then just replace the last values (last data bar and after)
        '   - if not matches, then must recalc for full data set (set flag to -1)
        If m.iTryLastBarRecalc > 0 Then
            ' find max # of initial null's
            nSkipBars = 0
            nMaxNulls = MaxInitialNulls(aArrayOfResults)
            For nLastGoodBar = m.Bars.Size - 1 To 0 Step -1
                If m.Bars(eBARS_Close, nLastGoodBar) <> kNullData Then
                    nSkipBars = nLastGoodBar - nMaxNulls * 1 - 500 '50
                    dVal = m.Bars.SessionDate(nSkipBars)
                    ' search backwards for end of previous day
                    For nSkipBars = nSkipBars - 1 To 0 Step -1
                        If m.Bars.SessionDate(nSkipBars) < dVal Then
                            ' but make sure to have at least the complete previous day
                            If dVal >= m.Bars.SessionDate(nLastGoodBar) - 0 Then
                                dVal = m.Bars.SessionDate(nSkipBars)
                            Else
                                Exit For
                            End If
                        End If
                    Next
                    Exit For
                End If
            Next
            If nSkipBars < 1000 Then
                m.iTryLastBarRecalc = -1 ' set flag so won't keep trying
            ElseIf Not aSmallerBars.CopyFromHandle(aArrayOfBars(0)) Then
                m.iTryLastBarRecalc = -1 ' set flag so won't keep trying
            Else
                aSmallerBars.DeleteFirstBars nSkipBars
                nNum = aSmallerBars.Size
                aArrayOfBars(0) = aSmallerBars.BarsHandle
                
                aTempResults.Create eGDARRAY_Longs, aArrayOfResults.Size, 0
                For i = 0 To aTempResults.Size - 1
                    aTempResults(i) = gdCreateArray(eGDARRAY_Doubles, nNum)
                Next
                
                dEngineTime = gdTickCount(False)
                astrParms.Size = 1
                If bInitExprDone Then
                    astrParms(0) = "Chart" & Str(m.Form.hWnd)
LogCCI "Begin1 Run"
                    rc = RunExpressions(astrParms.ArrayHandle, astrBarNames.ArrayHandle, _
                        aArrayOfBars.ArrayHandle, aTempResults.ArrayHandle, ByVal 0&, ByVal 0&)
                Else
                    astrParms(0) = ""
LogCCI "Begin1 Exec"
                    rc = ExecuteExpressions(astrParms.ArrayHandle, _
                        astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                        astrExpr.ArrayHandle, aTempResults.ArrayHandle, ByVal 0&, ByVal 0&)
                End If
LogCCI "End1"
                If rc = 0 Then
                    g.bDirtyFunctionLibrary = False
                Else
                    m.iTryLastBarRecalc = -1 ' set flag so won't keep trying
                End If
                ' track exp moving avg of last 10 recalc's (unless < 25% of prev avg)
                m.dLastRecalcDone = gdTickCount(False)
                dEngineTime = m.dLastRecalcDone - dEngineTime
                If m.dAvgRecalcTime > 0 And dEngineTime > m.dAvgRecalcTime * 0.25 Then
                    m.dAvgRecalcTime = m.dAvgRecalcTime + (dEngineTime - m.dAvgRecalcTime) * 2 / 11#
                Else
                    m.dAvgRecalcTime = dEngineTime
                End If
                
                For iArray = 0 To aTempResults.Size - 1
                    hArray = aTempResults(iArray)
                    hArray2 = aArrayOfResults(iArray)
                    
                    ' calc largest diff allowed (6 significant digits of array's range)
                    dVal = gdRange(hArray2, 0, -1)
                    If dVal < 0 Then dVal = 0 '(e.g. if Null?)
                    dDiffAllowed = dVal * 0.000001
                    
#If 0 Then
                    ' check beginning of array to make sure # of initial nulls match
                    For iPos = 0 To gdGetSize(hArray) - 1
                        If m.iTryLastBarRecalc < 0 Then Exit For
                        dVal = gdGetNum(hArray, iPos)
                        dVal2 = gdGetNum(hArray2, iPos)
                        If dVal <> kNullData And dVal2 <> kNullData Then
                            Exit For ' done checking
                        ElseIf dVal - dVal2 <> 0 Then
                            ' something didn't match, so this method won't work
                            dVal = dVal
                            'm.iTryLastBarRecalc = -1 ' set flag so won't keep trying
                        End If
                    Next
#End If
                    
                    ' check end of array
                    bNonNullFound = False
                    For iPos = m.Bars.Size - 1 To 0 Step -1
                        If m.iTryLastBarRecalc < 0 Then Exit For
                        dVal = gdGetNum(hArray, iPos - nSkipBars)
                        If dVal <> kNullData Then
                            bNonNullFound = True
                        End If
                        If iPos >= nLastGoodBar Then
                            gdSetNum hArray2, iPos, dVal
                        Else
                            dVal2 = gdGetNum(hArray2, iPos)
                            'If dVal = 0 Or dVal2 = 0 Then
                            If Abs(dVal - dVal2) > dDiffAllowed Then
                                ' something didn't match, so this method won't work
                                dVal = dVal
                                m.iTryLastBarRecalc = -1 ' set flag so won't keep trying
                                DebugLog "Full Recalc required:" & vbTab & m.strTemplate & vbTab & m.strTemplateApplied _
                                    & vbTab & Str(dVal) & vbTab & Str(dVal2) & vbTab & Str(dDiffAllowed) _
                                    & vbTab & DateFormat(m.Bars(eBARS_DateTime, iPos), MM_DD_YYYY, HH_MM)
                            End If
#If 0 Then
                            If Abs(dVal) <= 0.0000005 Or Abs(dVal2) <= 0.0000005 Then
                                If Abs(dVal - dVal2) > 0.0000005 Then
                                    ' something didn't match, so this method won't work
                                    dVal = dVal
                                    m.iTryLastBarRecalc = -1 ' set flag so won't keep trying
                                End If
                            ElseIf Abs(dVal / dVal2 - 1) > 0.000005 Then '0.000005
                                ' something didn't match, so this method won't work
                                dVal = dVal
'                                m.iTryLastBarRecalc = -1 ' set flag so won't keep trying
                            End If
#End If
                            ' only need to check for 20 bars (but need at least 1 non-null found)
                            If iPos <= nLastGoodBar - 20 And bNonNullFound Then
                                Exit For
                            End If
                        End If
                    Next
                    
                    gdDestroyArray hArray
                Next
                aTempResults.Size = 0
                aArrayOfBars(0) = m.Bars.BarsHandle
            End If
            nNum = m.Bars.Size
        End If

        If m.iTryLastBarRecalc > 0 Then
            ' Can use the shorter recalc method (when updating new ticks on current bar),
            ' but always clear the flag (this flag will only get set from UpdateTicks)
            m.iTryLastBarRecalc = 0
        Else
            ' Do CALC for the entire data set:
            ' size and re-init result arrays
            For i = 0 To aArrayOfResults.Size - 1
                gdSetSize aArrayOfResults(i), nNum, False
            Next

            ' call engine to calculate results
            dEngineTime = gdTickCount(False)
            ' (ANOMALY: when SetupExpressions has already been successfully called,
            ' we should be calling RunExpressions instead of ExecuteExpressions,
            ' BUT for some odd reason calling RunExpressions actually takes longer!)
            'If 0 Then
            ' TLB 10/25/2006: now it does seem faster and we NEED it (e.g. for Woodies)
            astrParms.Size = 1
            If bInitExprDone Then
                astrParms(0) = "Chart" & Str(m.Form.hWnd)
LogCCI "Begin Run"
                rc = RunExpressions(astrParms.ArrayHandle, astrBarNames.ArrayHandle, _
                        aArrayOfBars.ArrayHandle, aArrayOfResults.ArrayHandle, ByVal 0&, ByVal 0&)
            Else
                astrParms(0) = ""
LogCCI "Begin Exec"
                rc = ExecuteExpressions(astrParms.ArrayHandle, _
                    astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                    astrExpr.ArrayHandle, aArrayOfResults.ArrayHandle, ByVal 0&, ByVal 0&)
            End If
LogCCI "End"
            If rc = 0 Then g.bDirtyFunctionLibrary = False
            ' track exp moving avg of last 10 recalc's (unless < 25% of prev avg)
            m.dLastRecalcDone = gdTickCount(False)
            dEngineTime = m.dLastRecalcDone - dEngineTime
            If m.dAvgRecalcTime > 0 And dEngineTime > m.dAvgRecalcTime * 0.25 Then
                m.dAvgRecalcTime = m.dAvgRecalcTime + (dEngineTime - m.dAvgRecalcTime) * 2 / 11#
            Else
                m.dAvgRecalcTime = dEngineTime
            End If
        End If

        m.strRecalcMsg = "Recalc = " & Format(dEngineSetupTime, "#0") & " + " & _
            Format(dEngineTime, "#0") & " ms (" & Format(m.dAvgRecalcTime, "#0") _
            & " avg), TryLB = " & Str(m.iTryLastBarRecalc)

If hBars <> 0 Then gdDestroyArray hBars
hBars = 0

        ' truncate trailing values for result arrays (i.e. if flat-lined past end of Market1)
If 0 Then
        For iArray = 0 To aArrayOfResults.Size - 1
            hArray = aArrayOfResults(iArray)
            If hArray <> 0 Then
                For idx = gdGetSize(hArray) - 1 To 1 Step -1
                    ' exit when hit end of Market1
                    If m.Bars(eBARS_Close, idx) <> kNullData Then Exit For
                    If gdGetNum(hArray, idx) <> kNullData Then
                        ' or exit when hit first value that differs from prior value
                        If gdGetNum(hArray, idx) <> gdGetNum(hArray, idx - 1) Then Exit For
                        gdSetNum hArray, idx, kNullData
                    End If
                Next
            End If
        Next
End If

        If rc = 0 Then
            AddSpeed "Calc Indicators (" & Str(Int(dEngineTime)) & " ms)"
        Else
            AddSpeed "Calc Indicators (Err = " & Str(rc) & ")"
            strText = astrParms(astrParms.Size - 1)
            FileFromString App.Path & "\chk\Coded.chk", _
                "ERROR " & CStr(rc) & ": " & strText, True, True
        End If
            
        ' now align comparison symbols to Market1
        If aCompSymbolIdx.Size > 0 Then
            For idx = 0 To aCompSymbolIdx.Size - 1
                Set Indicator = m.Tree(aCompSymbolIdx(idx))
                strSymbol = UCase(Chr(34) & Indicator.Name & "," & GetPeriodStr(m.nPeriodicity) & Chr(34))
                For i = 1 To astrBarNames.Size - 1
                    If strSymbol = UCase(astrBarNames(i)) Then
                        ' start with a COPY of the original non-aligned bars, and then align it to Market1
                        Set Indicator.Bars = m.OtherBars(i).MakeCopy
                        'rc = gdAlignBars(Indicator.Bars.BarsHandle, m.aXdate.ArrayHandle, 0)
                        'rc = gdAlignBars(Indicator.Bars.BarsHandle, m.Bars.ArrayHandle(eBARS_DateTime), 0)
                        rc = gdAlignBars(Indicator.Bars.BarsHandle, m.Bars.BarsHandle, 0)
                        Exit For
                    End If
                Next
            Next
            AddSpeed "Align bars"
        End If
    
  
If 0 Then
        ' for debugging: dump results to a file
        rc = FreeFile
        Open App.Path & "\Chk\Results.chk" For Output As #rc
        For i = 0 To m.Bars.Size - 1
            strText = DateFormat(m.Bars(eBARS_DateTime, i), MM_DD_YYYY, HH_MM) & vbTab & Str(m.Bars(eBARS_Close, i))
            For iArray = 0 To aArrayOfResults.Size - 1
                hArray = aArrayOfResults(iArray)
                strText = strText & vbTab & Str(gdGetNum(hArray, i))
            Next
            Print #rc, strText
        Next
        Close #rc
End If
    
        'save inputs for comparing during next call
        Set astrBarNames2 = astrBarNames
        Set astrExpr2 = astrExpr
        Set aArrayOfResults2 = aArrayOfResults
        strCompMarkets2 = strCompMarkets
        
        If Len(m.strRexDataFile) > 0 Then
            If frmTest.Visible Then
                frmTest.AddList "Indicators for  " & FileBase(m.strRexDataFile), True
            End If
        End If
    End If
  
ErrExit:
    Set barsOther = Nothing
    Set Indicator = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cChart.LoadIndicators", eGDRaiseError_Raise

End Sub

Private Sub LoadIndicators2()
On Error GoTo ErrSection:

    Dim idx&, nCount&, nNum&, nX&, rc&, strText$, i&, nSymbolID&, nMPB&, iBars&, nPeriodicity&, dEngineTime#, nLastDate&
    Dim strSymbol$, strPeriod$, iPos&, hBars&, hArray&, strMarket$, strFormat$, strFileName$
    Dim bReloaded As Boolean, bAdd As Boolean, bRebuildCodedText As Boolean, bStartLastBar As Boolean
    Dim Indicator As cIndicator, IndParent As cIndicator
    Dim Pane As cPane
    Dim bHideThisInd As Boolean

    Dim nLastGoodBar&, hArray2&, dVal#, dVal2#, nSkipBars&, nMaxNulls&, iArray&, dDiffAllowed#
    Dim bNonNullFound As Boolean
    Dim aTempResults As New cGdArray
    Dim aSmallerBars As New cGdBars
    
    Dim astrParms As New cGdArray, astrBarNames As New cGdArray
    Dim astrExpr As New cGdArray, aArrayOfBars As New cGdArray
    Dim aArrayOfResults As New cGdArray
    Dim strCompMarkets As String
    
    Static aFromBarIndex As cGdArray
    Static astrBarNames2 As cGdArray
    Static astrExpr2 As cGdArray
    Static aArrayOfResults2 As cGdArray
    Static strCompMarkets2 As String
    Static dLastPrice As Double
    Static bInitExprDone As Boolean
    Static dPrevSessionDate As Double
    Static bUsingArtPyramids As Boolean
    
    Dim barsOther As New cGdBars
    Dim aCompSymbolIdx As New cGdArray
    Dim aVerifyResults As New cGdArray
   
    ' when starting, don't do systems yet (will be done from tmrMain after all is initialized)
    'If g.bStarting Or g.bUnloading Then Exit Sub
    If g.bUnloading Then Exit Sub
    
    If TypeOfChart = eTypeChart_Seasonal Then Exit Sub
    
    ' TLB 11/16/2010: seems like we can have a problem if we try to calculate indicators
    ' too prematurely (i.e. before the form has started showing), so we'll skip this until then.
    If m.Form Is Nothing Then Exit Sub
    If m.Form.Visible = False And dPrevSessionDate = 0 Then Exit Sub
    
    If aFromBarIndex Is Nothing Then
        Set aFromBarIndex = New cGdArray
        aFromBarIndex.Create eGDARRAY_Longs, 0, 0
    End If
    
    aCompSymbolIdx.Create eGDARRAY_Longs
    astrParms.Create eGDARRAY_Strings
    astrExpr.Create eGDARRAY_Strings
    aArrayOfResults.Create eGDARRAY_Longs
    aArrayOfBars.Create eGDARRAY_Longs
    
    ' make sure the primary bars is the first item in this array
    aArrayOfBars(0) = m.Bars.BarsHandle
    astrBarNames(0) = "Market1"
    
    'make sure price cluster exists if time cluster exists
    InitClusterInd True
    
    nX = m.aXBar.Size
    nNum = m.Bars.Size
    nCount = m.Tree.Count
    For idx = 1 To nCount
        If m.Tree.NodeLevel(idx) = 0 Then
            Set Pane = m.Tree(idx)
        Else
            Set Indicator = m.Tree(idx)
            With Indicator
                bHideThisInd = False
                If Not Pane.Display And .Name = "Time Clusters" Then
                    If .trueRangeColor = 1 Then Pane.Display = True 'temporarily set to true to load data for heatmap without showing time cluster
                End If
                If Pane.Display Then
                    If .DataType = eINDIC_BarData Then
                        ' load comparison symbol
                        If UCase(m.Tree.Key(idx)) = "PRICE" Then
                            .Expression = "Market1"
                        Else
                            .Expression = Chr(34) & UCase(.Name) & Chr(34) ' "Market" & Str(aCompSymbolIdx.Size)
                            bAdd = .Display
                            If Not bAdd Then
                                For i = idx + 1 To m.Tree.RelativeIndex(idx, eTREE_LastDescendant)
                                    If m.Tree(i).Display Then
                                        bAdd = True
                                        Exit For
                                    End If
                                Next
                            End If
                            If bAdd Then
                                strText = Chr(34) & UCase(.Name) & "," & GetPeriodStr(m.nPeriodicity) & Chr(34)
                                For i = 0 To astrBarNames.Size - 1
                                    If astrBarNames(i) = strText Then
                                        bAdd = False
                                        If UCase(.Name) = "HEIKIN-ASHI" Then
                                            ' can only show 1 set of H.A. bars
                                            .Display = False
                                        End If
                                        Exit For
                                    End If
                                Next
                                
                                '5442 -always add index of indicator needing bars data
                                '   indicators will share bars if both use same sym & periodicity
                                aCompSymbolIdx.Add idx
                                
                                If bAdd Then
                                    astrBarNames.Add strText
                                    strCompMarkets = strCompMarkets & "," & UCase(.Name)
                                End If
                            End If
                        End If
                        .BuildText Me, "", bRebuildCodedText
                        ' TLB 4/11/2012: fix label if an external symbol (from hard drive)
                        strText = .ChartLabel
                        If InStr(strText, "|") > 0 Then
                            .ChartLabel = Parse(strText, "|", 1)
                        End If
                    ElseIf .DataType = eINDIC_Array Or .DataType = eINDIC_BooleanArray Or .DataType = eINDIC_DrawCommands Or .DataType = eINDIC_ProfileVolume Then
                        strMarket = ""
                        If m.Tree.NodeLevel(idx) > 1 Then
                            If .DataType = eINDIC_BooleanArray Then
                                i = m.Tree.AncestorIndex(idx, 1)
                                Set IndParent = m.Tree(i)
                                If IndParent.DataType = eINDIC_BarData Then
                                    strMarket = IndParent.Expression
                                End If
                                ' TLB 7/18/2014: if 1st parm is an "array" and default is
                                ' not a number, then link the 1st parm to it's parent's expression
                                If .ParmType(1) = 4 And Not IsDigit(.ParmDefault(1)) Then
                                    i = m.Tree.RelativeIndex(idx, eTREE_Parent)
                                    Set IndParent = m.Tree(i)
                                    If IndParent.DataType <> eINDIC_BarData Then
                                        strText = IndParent.Expression
                                        .Parm(1) = strText
                                    End If
                                End If
                            ElseIf .DataType = eINDIC_Array Or .DataType = eINDIC_DrawCommands Then
                                ' linked indicator
                                i = InStr(UCase(.Parm(1)), " OF ")
                                If i > 0 Then
                                    .Parm(1) = Trim(Left(.Parm(1), i))
                                End If
                                i = m.Tree.RelativeIndex(idx, eTREE_Parent)
                                Set IndParent = m.Tree(i)
                                If IndParent.DataType <> eINDIC_BarData Then
                                    strText = IndParent.Expression
                                    ' TLB: if parent of a TAS_Result is not displayed, then hide this as well
                                    If UCase(.CodedName) = "TAS_RESULT" And Not IndParent.Display Then
                                        bHideThisInd = True
                                    End If
                                ElseIf ParmDefaultIsBarsArray(.Parm(1)) Then
                                    strText = .Parm(1) & " of " & IndParent.Expression
                                Else
                                    strText = "Close of " & IndParent.Expression
                                End If
                                If strText = "" Then strText = "Close of Market1"
                                If .ParmType(2) = 4 And Not IsDigit(.ParmDefault(2)) Then '(fix for MovingSum aardvark issues)
                                    ' if both of first 2 parms are "arrays" and 2nd parm default is not a number
                                    ' then link 1st parm to grandparent and 2nd parm to parent
                                    .Parm(2) = strText
                                    strText = ""
                                    If m.Tree.NodeLevel(i) > 1 Then
                                        i = m.Tree.RelativeIndex(i, eTREE_Parent)
                                        Set IndParent = m.Tree(i)
                                        If IndParent.DataType <> eINDIC_BarData Then
                                            strText = IndParent.Expression
                                        ElseIf ParmDefaultIsBarsArray(.Parm(1)) Then
                                            strText = .Parm(1) & " of " & IndParent.Expression
                                        Else
                                            strText = "Close of " & IndParent.Expression
                                        End If
                                    End If
                                    If strText = "" Then strText = "Close of Market1"
                                    .Parm(1) = strText
                                Else
                                    ' just link 1st parm to parent
                                    .Parm(1) = strText
                                End If
                            End If
                        End If
                        
                        .BuildText Me, strMarket, bRebuildCodedText
                        If .Display Or .IsAlert Or .IsWoodInd(False, False, , , Me) Or Indicator Is m.PivotVwapInd Or _
                            .DisplayType = eINDIC_ClusterPrice Or .DisplayType = eINDIC_ClusterTime Then
                            strText = Trim(.CodedText)
                            If Len(strText) = 0 Or bHideThisInd Then
                                If bHideThisInd Then
                                    .ChartLabel = ""
                                Else
                                    .ChartLabel = "INVALID"
                                End If
                                gdClear .Data.ArrayHandle, False
                            Else
                                astrExpr.Add strText
                                aArrayOfResults.Add .Data.ArrayHandle
                            End If
                            'reset time cluster pane to correct display value in case it was temporarily turned on for time cluster heatmap
                            If .DisplayType = eINDIC_ClusterTime Then Pane.Display = .Display
                        End If
                    End If
                End If
            End With
        End If
    Next
    Set Indicator = Nothing
    Set IndParent = Nothing
    Set Pane = Nothing
    
    AddSpeed "Code Indicators"

    ' get symbols/periods from expressions (puts them into BarNames and fixes Expressions)
    MarketsInExpressions astrExpr, 0, False, astrBarNames, Nothing, GetPeriodStr(m.nPeriodicity), Me.Symbol
        
    AddSpeed "MarketsInExpressions"
    
    ' see if we need to recalculate: if any array, secondary symbols,
    ' or last price has changed since the previous call
    If dLastPrice <> m.Bars(eBARS_Close, Me.LastGoodDataBar(False)) Then
        dLastPrice = m.Bars(eBARS_Close, Me.LastGoodDataBar(False))
        RedoMode = eRedo5_RecalcInd
    End If
    If strCompMarkets <> strCompMarkets2 Then
        RedoMode = eRedo5_RecalcInd
        'g.bSkipSetChartFocus = true means charts are being restored from chart page switching or program startup
        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        bInitExprDone = False
    ElseIf Not ArraysAreSame(astrBarNames, astrBarNames2) Then
        RedoMode = eRedo5_RecalcInd
        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        bInitExprDone = False
    ElseIf Not ArraysAreSame(astrExpr, astrExpr2) Then
        RedoMode = eRedo5_RecalcInd
        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        bInitExprDone = False
    ElseIf Not ArraysAreSame(aArrayOfResults, aArrayOfResults2) Then
        RedoMode = eRedo5_RecalcInd
        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        bInitExprDone = False
    ElseIf RedoMode >= eRedo9_ReloadData Then
        bInitExprDone = False
    ElseIf RedoMode = eRedo6_ReloadInd Then
        bInitExprDone = False '(reload since a function may have been edited)
    ElseIf m.Bars.SessionDate(Me.LastGoodDataBar(False)) <> dPrevSessionDate Then
        bInitExprDone = False ' #4745: must reinit expressions when starting a new trading session
    End If
    
    If RedoMode >= eRedo5_RecalcInd Then
        dPrevSessionDate = m.Bars.SessionDate(Me.LastGoodDataBar(False))

        ' see if we need to reload the secondary data
        If RedoMode >= eRedo9_ReloadData Or Not ArraysAreSame(astrBarNames, astrBarNames2) Or strCompMarkets <> strCompMarkets2 Then
        
            ' first clear the existing secondary data bars
            ReDim m.OtherBars(astrBarNames.Size - 1) As cGdBars
            bReloaded = True
            bInitExprDone = False

            If astrBarNames.Size > 1 Then
                For i = 1 To astrBarNames.Size - 1
                    Set barsOther = Nothing
                    nSymbolID = 0
                    strSymbol = UCase(astrBarNames(i))
                    If Left(strSymbol, 1) = Chr(34) And Right(strSymbol, 1) = Chr(34) Then
                        strPeriod = Parse(Replace(strSymbol, Chr(34), ""), ",", 2)
                        strSymbol = Parse(Replace(strSymbol, Chr(34), ""), ",", 1)
                        If Right(strSymbol, 1) = ":" Then
                            strSymbol = strSymbol & Me.Symbol
                        End If
                        If Len(strSymbol) = 0 Then
                            nSymbolID = m.nSymbolID
                        Else
                            nSymbolID = GetSymbolID(strSymbol)
                        End If
                    ElseIf strSymbol = "DAILY" Or strSymbol = "WEEKLY" Or strSymbol = "MONTHLY" Then
                        strPeriod = strSymbol
                        nSymbolID = m.nSymbolID
                    ElseIf strSymbol = "GC" Or strSymbol = "TQ" Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol & "-067")
                    ElseIf UCase(strSymbol) <> "DEFAULT" Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
                    ElseIf Left(PrimaryFutureBase(g.SymbolPool.SymbolForID(m.nSymbolID)), 3) = "GC-" Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol("TQ-067")
                    Else
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol("GC-067")
                    End If
                    ' TLB 4/11/2012: allow if an external symbol (from hard drive)
                    If nSymbolID <> 0 Or InStr(strSymbol, "|") > 0 Then
                        ' TLB 10/12/2009: allow for putting ",2x" as the period (which means "double" the primary
                        ' -- e.g. if the chart's bar period is 5 minutes, then ",2x" would mean 10 minutes)
                        rc = 0
                        If Len(strPeriod) = 0 Then
                            strPeriod = GetPeriodStr(m.nPeriodicity)
                        ElseIf Right(strPeriod, 1) = "X" Then
                            rc = Val(Parse(strPeriod, "X", 1))
                        ElseIf Left(strPeriod, 1) = "X" Then
                            rc = Val(Parse(strPeriod, "X", 2))
                        End If
                        If rc > 0 Then
                            rc = GetPeriodsPerBar(m.nPeriodicity) * rc
                            nPeriodicity = GetPeriodType(m.nPeriodicity) + rc
                        Else
                            nPeriodicity = GetPeriodicity(strPeriod)
                        End If
                        If nPeriodicity > ePRD_Days + 1 Then
'                            nPeriodicity = ePRD_Days + 1
                        ElseIf nPeriodicity >= ePRD_IntRenko And nPeriodicity < ePRD_IntVol Then
                            nPeriodicity = ePRD_Minutes + 1
                        End If
                        Set barsOther = New cGdBars
                        If nPeriodicity >= ePRD_Days Then 'And nSymbolID = m.nSymbolID Then
                            ' TLB 5/21/2012: a user-specified ending date should be used for loading data
                            If m.bToEndOfData Then
                                nLastDate = 0
                            Else
                                nLastDate = m.dToDate
                            End If
                            If InStr(strSymbol, "|") > 0 Then
                                ' TLB 4/11/2012: load data for an external symbol (from hard drive)
                                strFormat = UCase(Parse(strSymbol, "|", 5))
                                strFileName = Parse(strSymbol, "|", 2)
                                strSymbol = Parse(strSymbol, "|", 1)
                                If strFormat <> "GDB" Then
                                    strFileName = FilePath(strFileName)
                                End If
                                barsOther.ArrayMask = eBARS_Eod
                                SetBarProperties barsOther, strSymbol
                                rc = barsOther.FromFile(strFormat, strFileName, strSymbol, strPeriod, , 0, nLastDate)
                            Else
                                rc = DM_GetBars(barsOther, nSymbolID, nPeriodicity, 0, nLastDate, , , m.bUnsplit)
                            End If
                        ElseIf m.nPeriodicity < ePRD_Days And m.nPartialLoadDate <= 0 Then
                            rc = DM_GetBars(barsOther, nSymbolID, nPeriodicity, m.Bars(eBARS_DateTime, 0), m.Bars(eBARS_DateTime, m.Bars.Size - 1), , , m.bUnsplit)
                        Else
                            ' TLB 9/30/2009: it is bogus to load intraday data for secondary bars when the
                            ' primary bars is not intraday (and it will bog down forever loading all that data!)
                            barsOther.Size = 0
                            SetBarProperties barsOther, nSymbolID
                        End If
                        ' Image Server: merge Rex data
                        If Len(m.strRexDataFile) > 0 Then
                            AppendRexData barsOther ', m.Bars(eBARS_DateTime, LastGoodDataBar(False))
                        End If
                    End If
                    Set m.OtherBars(i) = barsOther
                Next
                AddSpeed Str(astrBarNames.Size - 1) & " other symbols loaded"
            End If
        End If

        ' final stuff for the secondary data bars
        If UBound(m.OtherBars) > 0 Then
            For iBars = 1 To UBound(m.OtherBars)
                Set barsOther = m.OtherBars(iBars)
                
                ' see if need to refresh the real-time portion
                If bReloaded Or RedoMode >= eRedo7_ReloadRT Then
                    ' if Heikin-Ashi bars, recreate them now
                    If InStr(UCase(astrBarNames(iBars)), "HEIKIN-ASHI") > 0 Then
                        ' TLB 6/17/2008: must be careful with newer engine to NOT change the .BarsHandle
                        gdCopy barsOther.BarsHandle, m.Bars.CreateHeikinAshi.BarsHandle
                        barsOther.Prop(eBARS_CustomString) = "HEIKIN-ASHI" '(to flag this for realtime updating)
                    Else
                        ' chop any blank bars at end
                        For i = barsOther.Size - 1 To 0 Step -1
                            If barsOther(eBARS_Close, i) <> kNullData Then
                                barsOther.Size = i + 1
                                Exit For
                            End If
                        Next
                        ' then add forecast bars
                        If barsOther.Size > 0 Then
                            barsOther.AddForecastBars m.nForecastBars
                        End If
                    End If
                    
                    ' setup for realtime updating
                    If Not m.bDisableRT Then
                        ' TLB 12/21/2007: can skip this under the specific condition:
                        ' if other indicator is "Daily" of the same symbol and this chart is intraday
                        ' (this helps speed up when new bar starts -- e.g. Woodies templates with pivot points)
                        If bReloaded Or m.nSymbolID <= 0 Or barsOther.Prop(eBARS_SymbolID) <> m.nSymbolID _
                                Or m.nPeriodicity >= ePRD_Days Or barsOther.Prop(eBARS_Periodicity) <> ePRD_Days + 1 Then
                                
                            If Screen.ActiveForm Is frmSymbolGrid Then
                                g.RealTime.AddTickBuffer barsOther, False
                            Else
                                g.RealTime.AddTickBuffer barsOther, True
                            End If
                            ' don't splice bars for Heikin-Ashi (don't need to, and will mess them up)
                            If m.bRT And barsOther.Prop(eBARS_CustomString) <> "HEIKIN-ASHI" Then
                                g.RealTime.SpliceBars barsOther
                                ' TBL 2/25/2008: need to add forecast bars back on so future bars will be Null
                                barsOther.AddForecastBars m.nForecastBars
                            End If
                        End If
                    End If
            
                    ' now compress bars
                    ' TLB 12/21/2007: this is messing up "of Daily" on intraday charts other than minute bars
                    ' (e.g. Pivot Points on breakout bar charts)
                    If 0 Then
                        Select Case GetPeriodType(m.nPeriodicity)
                        Case ePRD_Minutes, ePRD_Days, ePRD_Weeks, ePRD_Months, ePRD_Quarters, ePRD_Years
                            ' for fixed timeframes, compress to that timeframe before
                            ' calculating the indicators, then align to Market1 later
                            If barsOther.Prop(eBARS_Periodicity) <> m.nPeriodicity Then
                                barsOther.BuildBars GetPeriodStr(m.nPeriodicity)
                            End If
                        Case Else
                            ' for non-fixed timeframes, align to Market1 before calculating
                            ' indicators (but flat-line instead of empty bars so indicators work)
                            rc = gdAlignBars(barsOther.BarsHandle, m.Bars.BarsHandle, 1)
                        End Select
                    End If
                End If
                    
                ' build array of handles to bars
                aArrayOfBars.Add barsOther.BarsHandle
            Next
            AddSpeed Str(UBound(m.OtherBars)) & " other symbols updated"
        End If
                       
        ' call engine to calculate results
        dEngineTime = gdTickCount(False)
        rc = 0
        ' if realtime and the indicators haven't changed, we just need to recalc the last bar
        ' -- otherwise we need to calc the entire set of data
        If RedoMode >= eRedo7_ReloadRT Then ' Or Not g.RealTime.Active Then
            For i = 0 To aFromBarIndex.Size - 1
                aFromBarIndex(i) = 0
            Next
            m.DrawingCommands.Size = 0
        End If
        If bInitExprDone Then
            ' then recalc the last bar
            bStartLastBar = False
            astrParms.Size = 1
            astrParms(0) = "Chart-" & Str(m.Form.hWnd)
' eventually need to remove drawing commands from last bar forward?
            rc = ContinueLastBarSet(astrParms.ArrayHandle, aFromBarIndex.ArrayHandle, aArrayOfResults.ArrayHandle, _
                    0, m.DrawingCommands.ArrayHandle) ' ByVal 0&)
            If rc <> 0 Then
                ' if an error with ContinueLastBarSet, then reset flag so will do StartLastBarSet
                bInitExprDone = False
            Else
                ' track/display time of recalc
                m.dLastRecalcDone = gdTickCount(False)
                dEngineTime = m.dLastRecalcDone - dEngineTime
                If m.dAvgRecalcTime <= 0 Then
                    m.dAvgRecalcTime = dEngineTime
                ElseIf aFromBarIndex(0) > 0 Then
                    ' keep exp. moving avg of last 10 recalc's
                    m.dAvgRecalcTime = m.dAvgRecalcTime + (dEngineTime - m.dAvgRecalcTime) * 2 / 11#
                End If
                m.strRecalcMsg = "Recalc = " & Format(dEngineTime, "#0") & " ms (" & _
                    Format(m.dAvgRecalcTime, "#0") & " avg)"
            End If
        End If
        If Not bInitExprDone Then
            'KillFile App.Path & "\chk\Coded.chk"
            If IsIDE Then
                astrExpr.ToFile App.Path & "\chk\Coded.chk"
                astrBarNames.ToFile App.Path & "\chk\BarNames.chk"
                aArrayOfBars.ToFile App.Path & "\chk\BarNames.chk", True
            End If
                                      
            ' first clear the old set, then run the new set
            bUsingArtPyramids = False
            bStartLastBar = True
            m.DrawingCommands.Size = 0
            astrParms.Size = 1
            astrParms(0) = "Chart-" & Str(m.Form.hWnd)
            ClearLastBarSet astrParms.ArrayHandle, ByVal 0&
            astrParms.Size = 1
            rc = StartLastBarSet(astrParms.ArrayHandle, astrBarNames.ArrayHandle, _
                aArrayOfBars.ArrayHandle, astrExpr.ArrayHandle, aArrayOfResults.ArrayHandle, _
                0, m.DrawingCommands.ArrayHandle)  'ByVal 0&)
            If rc = 0 Then
                bInitExprDone = True
                g.bDirtyFunctionLibrary = False
                For i = 0 To astrExpr.Size - 1
                    strText = astrExpr(i)
                    If InStr(UCase(strText), "PYRAMIDTP") > 0 Then
                        bUsingArtPyramids = True
                    ElseIf InStr(UCase(strText), "REVERSEBARTP") > 0 Then
                        bUsingArtPyramids = True
                    ElseIf Mid(strText, 7, 3) = "ART" Then
                        bUsingArtPyramids = True
                    End If
                Next
                strText = ""
            Else
                strText = " -> ERR = " & Str(rc)
            End If
            ' track/display time of recalc
            m.dLastRecalcDone = gdTickCount(False)
            dEngineTime = m.dLastRecalcDone - dEngineTime
            m.strRecalcMsg = "Full Calc = " & Format(dEngineTime, "#0") & " ms" & strText
        End If
        
If IsIDE Then
    If m.DrawingCommands.Size > 0 Then
        StatusMsg "#DrawingCommands = " & Str(m.DrawingCommands.Size)
        m.DrawingCommands.ToFile App.Path & "\Chk\DrawingCmds.txt"
    End If
End If
        

        ' truncate trailing values for result arrays (i.e. if flat-lined past end of Market1)
If 0 Then
        For iArray = 0 To aArrayOfResults.Size - 1
            hArray = aArrayOfResults(iArray)
            If hArray <> 0 Then
                gdSetShared hArray, True
                For idx = gdGetSize(hArray) - 1 To 1 Step -1
                    ' exit when hit end of Market1
                    If m.Bars(eBARS_Close, idx) <> kNullData Then Exit For
                    If gdGetNum(hArray, idx) <> kNullData Then
                        ' or exit when hit first value that differs from prior value
                        If gdGetNum(hArray, idx) <> gdGetNum(hArray, idx - 1) Then Exit For
                        gdSetNum hArray, idx, kNullData
                    End If
                Next
                gdSetShared hArray, False
            End If
        Next
End If

        If rc = 0 Then
            AddSpeed "Calc Indicators (" & Str(Int(dEngineTime)) & " ms)"
        Else
            AddSpeed "Calc Indicators (Err = " & Str(rc) & ")"
            strText = astrParms(astrParms.Size - 1)
            FileFromString App.Path & "\chk\Coded.chk", _
                "ERROR " & CStr(rc) & ": " & strText, True, True
        End If
            
        ' set each item of the FromBarIndex to the last bar of data
        aFromBarIndex.Size = aArrayOfBars.Size
        For iBars = 0 To UBound(m.OtherBars)
            aFromBarIndex(iBars) = 0 ' default
            If iBars = 0 Then
                Set barsOther = m.Bars
            Else
                Set barsOther = m.OtherBars(iBars)
            End If
            If Not barsOther Is Nothing Then
                For i = barsOther.Size - 1 To 1 Step -1
                    If barsOther(eBARS_Close, i) <> kNullData Then
                        ' safer to go one bar back (e.g. for spread indicators on daily charts)
                        ' -- TLB 1/26/2012: except when using the ArtPyramids function which was
                        ' originally only a drawing tool and does change results in the past based
                        ' on current data (so for alerts and highlight bars to work correctly while
                        ' streaming, we need the recalc to get done for about the last 250 bars or so).
                        If Not bUsingArtPyramids Then
                            aFromBarIndex(iBars) = i - 1
                        ElseIf i > 500 Then
                            aFromBarIndex(iBars) = i - 250
                        Else
                            aFromBarIndex(iBars) = 0
                        End If
                        Exit For
                    End If
                Next
            End If
        Next
        Set barsOther = Nothing
    
        ' now align comparison symbols to Market1
        If aCompSymbolIdx.Size > 0 Then
            For idx = 0 To aCompSymbolIdx.Size - 1
                Set Indicator = m.Tree(aCompSymbolIdx(idx))
                strSymbol = UCase(Chr(34) & Indicator.Name & "," & GetPeriodStr(m.nPeriodicity) & Chr(34))
                For i = 1 To astrBarNames.Size - 1
                    If strSymbol = UCase(astrBarNames(i)) Then
                        ' start with a COPY of the original non-aligned bars, and then align it to Market1
                        Set Indicator.Bars = m.OtherBars(i).MakeCopy
                        'rc = gdAlignBars(Indicator.Bars.BarsHandle, m.aXdate.ArrayHandle, 0)
                        'rc = gdAlignBars(Indicator.Bars.BarsHandle, m.Bars.ArrayHandle(eBARS_DateTime), 0)
                        rc = gdAlignBars(Indicator.Bars.BarsHandle, m.Bars.BarsHandle, 0)
                        Exit For
                    End If
                Next
            Next
            AddSpeed "Align bars"
        End If
    
        ' TLB: temporary code for verifying results of the new Recalc Last Bar method
        If m.iVerifyResults = 1 Then
            aVerifyResults.Create eGDARRAY_Longs, astrExpr.Size, 0
            For iArray = 0 To aVerifyResults.Size - 1
                aVerifyResults(iArray) = gdCreateArray(eGDARRAY_Doubles)
            Next
            astrParms.Size = 1
            astrParms(0) = ""
            rc = ExecuteExpressions(astrParms.ArrayHandle, _
                    astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                    astrExpr.ArrayHandle, aVerifyResults.ArrayHandle, ByVal 0&, ByVal 0&)
            strText = ""
            For iArray = 0 To aVerifyResults.Size - 1
                hArray = aArrayOfResults(iArray)
                hArray2 = aVerifyResults(iArray)
                If gdGetSize(hArray) <> gdGetSize(hArray2) Then
                    strText = "Size mismatch:  " & Str(gdGetSize(hArray)) & "  " & Str(gdGetSize(hArray2)) & vbCrLf & astrExpr(iArray)
                Else
                    For i = 0 To gdGetSize(hArray) - 1
                        dVal = gdGetNum(hArray, i)
                        dVal2 = gdGetNum(hArray2, i)
                        'If dval <> dVal2 Then
                        If Abs(dVal - dVal2) > 0.000000005 Then
                            rc = 0
                            If dVal = kNullData Or dVal2 = kNullData Then
                                ' due to a slight change in SwingPoints (old vs. new):
                                ' ... ignore "mismatch" if one value is Null and the other is Zero
                                If dVal <> 0 And dVal2 <> 0 Then
                                    ' ... and ignore if in first 1/4 of array and the non-zero is an integer
                                    If i >= gdGetSize(hArray) / 4 Or Int(dVal) <> dVal Or Int(dVal2) <> dVal2 Then
                                        rc = 1
                                    End If
                                End If
                            ElseIf Abs(dVal - dVal2) > 0.0000002 * Abs(dVal) Then
                            'ElseIf Abs(dval - dVal2) > 0.0000000000002 * Abs(dval) Then
                                rc = 2
                            End If
                            If rc > 0 Then
                                strText = "Mismatch at bar " & Str(i) & " of " & Str(gdGetSize(hArray)) _
                                    & vbCrLf & Str(dVal) & "  " & Str(dVal2) & vbCrLf & astrExpr(iArray)
                                Exit For
                            End If
                        End If
                    Next
                End If
                gdDestroyArray aVerifyResults(iArray)
                aVerifyResults(iArray) = 0
            Next
            aVerifyResults.Size = 0
            If Len(strText) > 0 Then
                InfBox strText, "!", "+-Close", "Verify Results - MISMATCH", True
                strText = vbCrLf & Format(Now, "yyyy-mm-dd hh:mm:ss, ") & Str(bStartLastBar) & vbCrLf _
                    & Trim(m.Form.vseCaption.Caption) & vbCrLf & strText
                FileFromString App.Path & "\Recalc.log", strText, True, True
            Else
                m.strRecalcMsg = "MATCHED"
                g.nNumVerifiedRecalcs = g.nNumVerifiedRecalcs + 1
            End If
        End If
        If m.iVerifyResults > 1 Then
            m.iVerifyResults = m.iVerifyResults - 1
        ElseIf m.iVerifyResults = 1 Or (m.iVerifyResults = 0 And RedoMode >= eRedo9_ReloadData) Then
            If FileExist("c:\common\files.exe") Then
                m.iVerifyResults = Val(FileToString(App.Path & "\Provided\NewRecalc.chk", , True))
            End If
        End If
        
  
If 0 Then
        ' for debugging: dump results to a file
        rc = FreeFile
        Open App.Path & "\Chk\Results.chk" For Output As #rc
        For i = 0 To m.Bars.Size - 1
            strText = DateFormat(m.Bars(eBARS_DateTime, i), MM_DD_YYYY, HH_MM) & vbTab & Str(m.Bars(eBARS_Close, i))
            For iArray = 1 To aArrayOfBars.Size - 1
                idx = aArrayOfBars(iArray)
                hArray = gdBarsArray(idx, eBARS_Close)
                strText = strText & vbTab & Str(gdGetNum(hArray, i))
            Next
            For iArray = 0 To aArrayOfResults.Size - 1
                hArray = aArrayOfResults(iArray)
                strText = strText & vbTab & Str(gdGetNum(hArray, i))
            Next
            Print #rc, strText
        Next
        Close #rc
End If
    
        'save inputs for comparing during next call
        Set astrBarNames2 = astrBarNames
        Set astrExpr2 = astrExpr
        Set aArrayOfResults2 = aArrayOfResults
        strCompMarkets2 = strCompMarkets
        
        If Len(m.strRexDataFile) > 0 Then
            If frmTest.Visible Then
                frmTest.AddList "Indicators for  " & FileBase(m.strRexDataFile), True
            End If
        End If
    End If
    
ErrExit:
    Set barsOther = Nothing
    Set Indicator = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cChart.LoadIndicators2", eGDRaiseError_Raise
End Sub

Private Function ArraysAreSame(a1 As cGdArray, a2 As cGdArray) As Boolean
On Error GoTo ErrSection:

    Dim bSame As Boolean, i&
    
    If Not a1 Is Nothing And Not a2 Is Nothing Then
        If a1.Size = a2.Size And a1.ArrayType = a2.ArrayType Then
            bSame = True
            For i = 0 To a1.Size - 1
                If a1(i) <> a2(i) Then
                    bSame = False
                    Exit For
                End If
            Next
        End If
    End If
    
    ArraysAreSame = bSame

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ArraysAreSame", eGDRaiseError_Raise

End Function

Private Sub SetGenTickMode()
    'nGenTick = 0 '???
End Sub

Private Sub AddSpeed(ByVal msg$)
On Error GoTo ErrSection:

    Dim dCurTime#
    Static dPrevTime#

    dCurTime = gdTickCount(False)
    If Len(msg) > 0 Then
        m.aSpeed.Add CStr(Int(dCurTime - dPrevTime + 0.5)) & "  " & msg
    End If
    dPrevTime = dCurTime
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.AddSpeed", eGDRaiseError_Raise

End Sub

Public Function ShowSpeed(Optional ByVal bDisplay As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strText$, i&, nSum&
    
    For i = 0 To m.aSpeed.Size - 1
        strText = strText & m.aSpeed(i) & " | "
        nSum = nSum + Val(Parse(m.aSpeed(i), " ", 1))
    Next
    strText = strText & "TOTAL: " & Str(nSum) & " ms"
    If bDisplay Then
        InfBox strText, , , "Chart processing time"
    End If
    
    ShowSpeed = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ShowSpeed", eGDRaiseError_Raise

End Function


Public Sub GenerateChart(Optional ByVal eRedoMode As eChartRedoMode = eRedo3_Settings, _
            Optional ByVal bRealTimeUpdate As Boolean = False, _
            Optional ByVal bClearPFP As Boolean = True)
On Error GoTo ErrSection:

    Dim rc&, X#, nPoints&, i&, idx&, strSymbol$, dStartTime#, strText$
    Dim Annot As cAnnotation
    Dim bRunSystem As Boolean
    Dim aPrevOrders As cGdArray
    Static nPrevHideAnnotations As Long
    Static nPrevTreeCount As Long
    Static bInProgress As Boolean
    Static nTimesCalled As Long
    
    Dim AnnotWhatIf As cAnnotation
    Dim bRecalcClusters As Boolean

    ' TLB 1/8/2009: we don't want this to get called until after page is loaded
    If g.bLoadingChartPage Then
        ' if this is called while loading a chart page, this is probably a programming error
        ' -- but even so, we don't want it to continue because it causes crashing issues
        ' if called "too early", esp. with spread charts
        Exit Sub
    End If
    
    If bInProgress Then Exit Sub
    bInProgress = True
            
    dStartTime = gdTickCount(False)
    m.strRecalcMsg = ""
    If nTimesCalled > 2000000000 Then
        nTimesCalled = -2000000000
    Else
        nTimesCalled = nTimesCalled + 1
    End If
    AddSpeed ""
'gdResetProfiles 300, 499
'gdStartProfile 300
'gdStartProfile 310

If IsIDE And 0 Then
    If frmTest.Visible Then
        frmTest.AddList Format(Now, "hh:mm:ss") & vbTab & Me.Symbol & vbTab & Str(eRedoMode) & " " & Str(m.eRedoMode)
    End If
End If

    
'for backwards-compatibility
If eRedoMode < 0 Then
    eRedoMode = eRedo1_Scrolled
ElseIf eRedoMode = 0 Then
    eRedoMode = eRedo3_Settings
End If
    
    'disable real time for the primary game mode form
    'un-sync non-primary game mode forms if user turns on real-time after start of game
    If m.Form.IsInGameMode Then
        If m.Form.GameReplayMode = eGDReplayMode_Sync Then
            m.bDisableRT = False
        Else
            m.bDisableRT = True
        End If
        m.Form.ToggleOrderBar False, False, m.nSymbolPitType
    ElseIf Len(m.strSpreadSymbols) > 0 Then     'Or (InStr(Symbol, "$") <> 0 And Not IsForex(Symbol)) Then - 5372
        m.nShowTrades = 0
        m.Form.ToggleOrderBar False, False, m.nSymbolPitType
    ElseIf m.nShowTrades = 2 Then
        m.Form.ToggleOrderBar True, False, m.nSymbolPitType
    Else
        m.Form.ToggleOrderBar False, False, m.nSymbolPitType             '4310
    End If
        
    ' set the redo flag (will only be set higher)
    RedoMode = eRedoMode
    
    ' TLB 6/30/2008: when loading lots of intraday data in weekly chunks
    ' (but if just scrolling, then don't add more data now -- this allows scrolling to happen during loading)
    If RedoMode >= eRedo9_ReloadData Then
        If bClearPFP And Not m.Form Is Nothing Then m.Form.PfpReset ePfpReset_ClearAll
        ' clear this flag when reloading data (e.g. if symbol or bar period has changed)
        m.nPartialLoadDate = 0
        m.nTradeSymbolID = 0
    ElseIf m.nPartialLoadDate > 0 And RedoMode > eRedo1_Scrolled Then
        ' need to load another week
        RedoMode = eRedo9_ReloadData
    End If
    
    ' special case for reloading user annots
    If m.eRedoMode = eRedo2_ReloadAnnots Then
        TemplateLoadUserAnnots
        m.eRedoMode = eRedo1_Scrolled
    End If
    
    If g.ChartGlobals.nHideAnnotations <> nPrevHideAnnotations Then
        nPrevHideAnnotations = g.ChartGlobals.nHideAnnotations
        For i = 1 To m.Annots.Count
            Set Annot = m.Annots(i)
            If Annot.eUsage = eANNOT_UserAdded Then
                Annot.Hide g.ChartGlobals.nHideAnnotations
            End If
        Next
        Set Annot = Nothing
    End If
    
    
    If m.nSymbolID <> 0 Then m.strExtData = ""
    
'm.iMiniChart = 1 '30

    'DebugLog "Starting GenerateChart " & Str(RedoMode)

    ' set primary bars
    If m.Bars Is Nothing Or RedoMode > eRedo1_Scrolled Then
        idx = m.Tree.Index("PRICE")
        If idx > 0 Then Set m.Bars = m.Tree(idx).Bars
    End If
'    If IsIntraday(m.nPeriodicity) Then
'        If m.Bars.ArrayMask <> eBARS_Intraday Then
'            m.Bars.Size = 0
'        End If
'    ElseIf m.Bars.ArrayMask = eBARS_Intraday Then
'        m.Bars.Size = 0
'    End If
    If m.Bars.Size = 0 Then
        RedoMode = eRedo9_ReloadData
    End If
    
    If m.bSpreadAsDollar Then
        ' change to dollar rounding, unless already at 1 cent moves (e.g. stocks)
        If m.Bars.Prop(eBARS_TickMove) <> 0.01 Or m.Bars.Prop(eBARS_TickValue) <> 0.01 _
                    Or m.Bars.Prop(eBARS_MinMoveInTicks) <> 1 Then
            m.Bars.Prop(eBARS_TickMove) = 1
            m.Bars.Prop(eBARS_TickValue) = 1
            m.Bars.Prop(eBARS_MinMoveInTicks) = 1
        End If
    End If
    
    m.aSpeed.Size = 0
    
    VerifyPanes False

'    If Not bRealTimeUpdate Or RedoMode >= eRedo7_ReloadRT Then
'        ResetTips
'    End If

    SetGenTickMode
    'If GenTick = 0 Then graph_one_date = 0
  
    If RedoMode > eRedo1_Scrolled Then
        ' init the chart settings
        If nPrevTreeCount <> m.Tree.Count Then
            rc = geChartSettings(True)
            geForceRecalc       '4330
        Else
            rc = geChartSettings(False)
        End If
        
        If rc <> 0 Then
            bInProgress = False
            Exit Sub
        End If
        AddSpeed "Setup"
'gdStopProfile 310
'gdStartProfile 320

        If (RedoMode >= eRedo7_ReloadRT) Then
            
            ' auto-breakout (for the DanielCode server)
            If Not g.FractZen.Allowed Or (GetPeriodType(m.nPeriodicity) <> ePRD_IntBreakout) Then
                m.strAutoBreakout = "" ' need to clear flag if not doing breakout bars
            End If
            If Len(m.strAutoBreakout) > 0 And m.nSymbolID > 0 Then
                ' check the breakout bar size
                ' (TLB 1/23/2012: now also pass in the ending date)
                If m.bToEndOfData Then
                    X = 0
                Else
                    ' TLB: to show how FZ would have looked that day based on the previous day
                    X = Int(m.dToDate) '- 1
                End If
                'i = CalcAutoBreakoutRange(GetSymbol(m.nSymbolID), Int(X), Val(Parse(m.strAutoBreakout, vbTab, 1)), Val(Parse(m.strAutoBreakout, vbTab, 2)))
                i = g.FractZen.GetFractZenRange(GetSymbol(m.nSymbolID), Int(X)) ', Val(Parse(m.strAutoBreakout, vbTab, 1)), Val(Parse(m.strAutoBreakout, vbTab, 2)))
                If m.nPeriodicity <> ePRD_IntBreakout + i Then
                    ' if it needs to change, then need to reload all data
                    m.nPeriodicity = ePRD_IntBreakout + i
                    RedoMode = eRedo9_ReloadData
                End If
                m.Bars.Prop(eBARS_FractZen) = Abs(g.FractZen.UseMerged)
            End If
            
            ' reload data
            If TypeOfChart = eTypeChart_Seasonal Then LoadSeasonalInd
            
            'SKIP the dm_getbars and addforecastbars in LoadData for seasonal charts
            LoadData
            
            bRecalcClusters = True
            
            ' resync the toolbar
            If RedoMode >= eRedo9_ReloadData Then   '5631
                SyncToolbar
            End If
        ElseIf m.nSymbolID = 0 And Len(Symbol) > 0 Then
            'this code is to get spread charts with bar period > daily to update real time
            If (RedoMode = eRedo5_RecalcInd And Not m.Bars.IsIntraday And UCase(m.Bars.Prop(eBARS_PeriodicityStr)) <> "DAILY") Then
                LoadData
                SyncToolbar
            End If
        End If
'gdStopProfile 320
'gdStartProfile 330
        
'gdStopProfile 330
'gdStartProfile 340

        ' see if the indicators need to be recalculated
        If m.nLastGoodDataBarX > 0 Then
            If m.bUseNewRecalcMethod Then
                'call function for seasonal chart
                LoadIndicators2 ' NewRecalcMethod
            Else
                LoadIndicators ' Old method
            End If
            bRecalcClusters = True
        End If
'gdStopProfile 340
    End If
'gdStopProfile 310
    
'gdStartProfile 350
    
    ' set # pixels per bar
    If m.geChartStruct.x_dataPointPix <> PixelsPerBar Then
    
'JM 08-03-2011: do not set dirty page flag here, this code is hit when loading detached
'   charts from a chart page AFTER all chart page loading flags have been cleared
'   the dirty page flag is now set when changing symbol & bars periodicity
'        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        
        m.geChartStruct.x_dataPointPix = PixelsPerBar
        geForceRecalc
    End If

    If m.iMiniChart > 0 Then
        m.iMiniChartStartX = 0
        ' find date of last bar with good data
        For i = m.Bars.Size - 1 To 0 Step -1
            If m.Bars(eBARS_Close, i) <> kNullData Then
                X = Int(m.Bars(eBARS_DateTime, i))
                Exit For
            End If
        Next
        ' find last bar on screen of previous date
        For i = m.aXBar.Size - 1 To 0 Step -1
            rc = m.aXBar(i)
            If rc >= 0 Then
                If m.Bars(eBARS_DateTime, rc) < X Then
                    ' set number of points such that shows specified
                    ' minutes of previous day through end of today
                    nPoints = m.aXBar.Size - i + m.iMiniChart - 1
                    'add previous close horz line
                    m.Annots.Clear
                    Set Annot = New cAnnotation
                    If (Annot.CreateNew(Me, eANNOT_HorzLine, 1, X, m.Bars(eBARS_Close, rc), m.Bars(eBARS_Close, rc), vbRed, eANNOT_DashSm) = 0) Then
                        Annot.Color = vbRed
                        idx = m.Annots.Add(Annot)
                    End If

                    m.iMiniChartStartX = i
                    Exit For
                End If
            End If
        Next
        Set Annot = Nothing
        'note: this is all hard-coded based on settings in GO file:
        'bar period = 5, minichart width = 110 pixels, font size = 290
        m.nVertGrid = 1 'fine
        m.geChartStruct.y_scaleLoc = 2  'sets scale on right, chart flush left
        m.geChartStruct.fSize = 6
        m.geChartStruct.y_labelMaxChar = 5
        'set pixels per data point
        m.geChartStruct.x_dataPointPix = 1
    End If

    Dim nTwipsX&, nTwipsY&, nHdrFontSize&
    
    If m.nImgSrvChartWidth > 0 Then
        nTwipsX = m.nImgSrvChartWidth * Screen.TwipsPerPixelX
        nTwipsY = m.nImgSrvChartWidth * 0.75 * Screen.TwipsPerPixelY
        If m.Form.Width < nTwipsX Or m.Form.Height < nTwipsY Then
            m.Form.Width = nTwipsX
            m.Form.Height = nTwipsY + (m.hsb.Height * 3) + (10 * Screen.TwipsPerPixelY)
        End If
        m.Form.pbChart.ScaleWidth = nTwipsX
        m.Form.pbChart.ScaleHeight = nTwipsY
        gdSetNum m.geChartStruct.glhWinRect, 0, 0
        gdSetNum m.geChartStruct.glhWinRect, 1, 0
        gdSetNum m.geChartStruct.glhWinRect, 2, nTwipsY
        gdSetNum m.geChartStruct.glhWinRect, 3, nTwipsX
        If m.iMiniChart = 0 Then
            If m.Bars.IsIntraday = True And m.nPeriodicity <= ePRD_Minutes + 15 Then
                m.nVertGrid = 1 'fine
            Else
                m.nVertGrid = 0 'coarse
            End If
            geSetImgSvrFontSize
        Else
            geRecalcChart m.geChartObj, m.geChartStruct, m.Form.pbChart.hDC
        End If
        m.nImgSrvChartWidth = 0
    End If
    
    nPoints = m.geChartStruct.x_dataPointCount
    If nPoints > m.aXBar.Size Then nPoints = m.aXBar.Size
        
    With m.hsb
        If .Min <> nPoints - 1 Then
            .Min = nPoints - 1
            .LargeChange = nPoints * 4 \ 5 ' 3 \ 4
        End If
        If .Value > .Max Then
            .Value = .Max
        ElseIf .Value < .Min Then
            .Value = .Min
        End If

        If .Max <= nPoints Then
            Disable m.hsb
        Else
            Enable m.hsb
            m.dLastScreenDate = m.aXdate(.Value)
        End If
    End With

    SetFormCaption

    If m.eRedoMode = eRedo9_ReloadData And Not m.Form Is Nothing Then
        m.Form.SetChartTabs
    End If

    LoadNotations

    Screen.MousePointer = 0

    'Enable tmr
    'Enable cmdOptions

    ' see if need to run (or re-run) the system
    bRunSystem = False
    If m.nShowTrades = 2 And Not m.Form.IsInGameMode Then
        If RedoMode >= eRedo9_ReloadData Then
            m.bReloadTrackerTrades = True
        End If
        ReadTrades False
    ElseIf m.nSystemID = 0 Or m.nShowTrades = 0 Or m.nPartialLoadDate > 0 Then
        If (m.nSystemID = 0 Or m.nShowTrades = 0) And m.nShowTrades = 1 Then
            m.nShowTrades = 0
        End If
        If m.Form.GameReplayMode = eGDReplayMode_Off Or m.Form.GameReplayMode = eGDReplayMode_Sync Then
            m.Transactions.NumRecords = 0
            m.aOrders.Clear
        ElseIf Not m.Form.GameMode Is Nothing Then
            'user removed strategy from chart when in game mode
            If m.nSystemID = 0 And (m.Form.GameMode.GameStrategyID > 0 And m.Form.GameMode.GameStrategyID <> kGameModeSysID) Then
                m.Transactions.NumRecords = 0
                m.aOrders.Clear
                m.Form.GameMode.GameStrategyID = 0
            End If
        End If
        ' We want to try to kill the system reports regardless of whether the
        ' system is nothing or not.  11/5/2003 DAJ
        'If Not m.System Is Nothing Then
            KillSystemReports
            Set m.System = Nothing
        'End If
    ElseIf m.Bars(eBARS_Close, 0) <> kNullData Then
        If m.System Is Nothing Then
            bRunSystem = True
        ElseIf m.System.SystemNumber <> m.nSystemID Then
            If m.nSystemID < 0 Then m.nSystemID = m.System.SystemNumber     'aardvark 3354 fix
            bRunSystem = True
        ElseIf RedoMode >= eRedo9_ReloadData Then
            bRunSystem = True
        ElseIf RedoMode >= eRedo7_ReloadRT Then
            If m.Form.GameReplayMode = 0 Then
                bRunSystem = True
            ElseIf m.dNextBarDate <> m.Bars(eBARS_DateTime, LastGoodDataBar(False)) Then
                bRunSystem = True
            End If
        End If
    End If

    Set aPrevOrders = m.aOrders.MakeCopy
    
    If m.bProfileChart Then AdjustProfileScrollBar
    
    If bRunSystem Then
        ' draw chart now without trades since could take awhile to run the system
        ' (clear current trades unless called from a real-time update)
        If Not bRealTimeUpdate Then
            m.Transactions.NumRecords = 0
            m.aOrders.Clear
        End If
        LoadScreenTrades
        LoadScreenData
        m.Form.pbChart.Cls
        rc = geDrawChart
        AddSpeed "geDrawChart (" & CStr(rc) & ")"
        'm.Form.pbChart.Refresh
        DoEvents
        Screen.MousePointer = vbHourglass
        
        ' run the system
        LoadSystem
        
        Screen.MousePointer = 0
    
        RemoveAnnots False, eANNOT_UndefinedType, eANNOT_Trades
        RemoveAnnots False, eANNOT_UndefinedType, eANNOT_ProfitLines
    End If
    EnablePerformanceButton
       
    Set AnnotWhatIf = m.Annots("WhatIf")
    If Not AnnotWhatIf Is Nothing Then
        i = LastGoodDataBar(False)
        'check for periodicity change, QB Refresh, template change etc.
        If AnnotWhatIf.Prop("WhatIfTemplate") <> m.strTemplateApplied Then
            DeactivateWhatIf False       'too many things may have changed, just deactivate
            SyncToolbar
        ElseIf AnnotWhatIf.BarPeriodAtCreate <> m.Bars.Prop(eBARS_PeriodicityStr) Or _
           AnnotWhatIf.dDate(1) <> m.Bars(eBARS_DateTime, i + 1) Or _
           AnnotWhatIf.Y(1) <> m.Bars(eBARS_Close, i) Then
            DeactivateWhatIf False
            ActivateWhatIf AnnotWhatIf.Prop("WhatIfBar")
        ElseIf InvalidWhatIf(False) Then
            DeactivateWhatIf False       'user changed display style
            SyncToolbar
        End If
        Set AnnotWhatIf = Nothing
    End If
    
    UpdateTradePrices       'do here so we get updated Ys for avgEntrLine before doing draw
    LoadScreenTrades
'gdStopProfile 350
'gdStartProfile 360
    
    If m.bRedrawTwice Then
        '01-08-2009 fix for issues with countdown panel
        m.bRedrawTwice = False
        LoadScreenData
        m.Form.pbChart.Cls
        rc = geDrawChart
    End If
    
    LoadScreenData , bRecalcClusters
'gdStopProfile 360
'gdStartProfile 370

    'sync the on/off form
    If m.eRedoMode <> eRedo1_Scrolled Then
        If bClearPFP Then
            'If nPrevTreeCount <> m.Tree.Count Or m.eRedoMode = eRedo9_ReloadData Then      'aardvark 6410
                If Not m.Form Is Nothing Then m.Form.PfpReset ePfpReset_GridInd, True
            'End If
        End If
        If m.Form Is ActiveChart Then frmChartOnOff.ShowData m.Form
    End If
    
    ' temporary solution until fix bug  (bug is fixed - leave awhile then remove 11-30-2004)
    'If m.eRedoMode >= eRedo9_ReloadData Then bDrawTwice = True
    
    ' clear the redo mode
    m.eRedoMode = eRedo1_Scrolled
    
    m.dLastCountdownCheck = gdTickCount(False)

    If Not m.System Is Nothing Then
        'If Not ArraysAreSame(aPrevOrders, m.aOrders) Then
        If 0 Then
            Set m.frmOrdersPopup = New frmAlertPopup
            m.frmOrdersPopup.ShowMe eGDAlertMode_NextBarAlert, Me.Symbol, "Orders have changed for:" & vbCrLf & m.System.SystemName, m.Form.hWnd
        End If
    End If
    
    ''SetFormCaption
ErrExit:
    On Error Resume Next
    nPrevTreeCount = Tree.Count
    
    bInProgress = False
    
    m.Form.pbChart.Cls

    'gdStopProfile 370
    'gdStartProfile 380
    If TypeOfChart = eTypeChart_Seasonal Then m.geChartStruct.paneCount = 999999
    
    rc = geDrawChart
    'gdStopProfile 380
    'gdStartProfile 390
    AddSpeed "geDrawChart (" & CStr(rc) & ")"
    
    If Not g.ChartGlobals.frmLastChartMouseMove Is Nothing Then
        g.ChartGlobals.frmLastChartMouseMove.GlobalCursorSync True
        AddSpeed "GlobalCursorSync"
    End If
'gdStopProfile 390
        
'gdStopProfile 300
    strText = ""
    If m.strRecalcMsg = "MATCHED" Then
        strText = "Recalc Verified"
    ElseIf g.bShowRecalcMsg Then
        strText = "#" & Str(nTimesCalled) & ":  " & Str(m.Bars.Size) & " bars,  "
        If Len(m.strRecalcMsg) > 0 Then
            strText = strText & m.strRecalcMsg & ",  "
        End If
        strText = strText & "Total = " & Str(Int(gdTickCount(False) - dStartTime)) & " ms"
    End If
    m.Form.vseDay.Tag = strText

If IsIDE Then
    If Not g.bStarting And Not g.bUnloading And Not g.RealTime.Active Then
        'InfBox gdGetProfiles(300, 499, vbCrLf), , , "GenerateChart", True
    End If
End If

    Exit Sub


ErrSection:
    bInProgress = False
    AddSpeed "ERROR raised"
    RaiseError "cChart.GenerateChart", eGDRaiseError_Raise

End Sub

Private Sub ResetTips()
On Error GoTo ErrSection:

    With m.Form.vseTipX
        '.BackColor = m.Peg.DeskColor   - MJM
        .Caption = ""
        .Top = -1000 - .Height
        '.Visible = g.ChartGlobals.bFloatingTips
    End With
    With m.Form.vseTipY
        '.BackColor = m.Peg.DeskColor   - MJM
        .Caption = ""
        .Top = -1000 - .Height
        '.Visible = g.ChartGlobals.bFloatingTips
    End With
    'With m.Peg
        'If 1 Then
           ' standard tooltip colors
            m.Form.lblTipX.BackColor = &H80000018
            m.Form.lblTipY.BackColor = &H80000018
            m.Form.lblTipX.ForeColor = &H80000017
            m.Form.lblTipY.ForeColor = &H80000017
        'ElseIf m.PicBox.TextColor = RGB(255, 255, 255) Then  'MJM
        '    m.Form.lblTipX.ForeColor = RGB(0, 0, 0)
        '    m.Form.lblTipY.ForeColor = RGB(0, 0, 0)
        'Else
        '    m.Form.lblTipX.ForeColor = .TextColor
        '    m.Form.lblTipY.ForeColor = .TextColor
        'End If
    'End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.ResetTips", eGDRaiseError_Raise

End Sub

Public Property Get Symbol() As String
On Error GoTo ErrSection:

    Static nPrevSymbolID&, strSymbol$
    Dim aSpreadSyms As cGdArray, i&
    
    If Len(m.strExtData) > 0 Then
        strSymbol = Parse(m.strExtData, "|", 1)
        nPrevSymbolID = 0
    ElseIf Len(m.strSpreadSymbols) > 0 Then
        Set aSpreadSyms = New cGdArray
        aSpreadSyms.SplitFields m.strSpreadSymbols, ";"
        ' use function name if exists, else concat all symbols
        strSymbol = aSpreadSyms(0)
        If Len(strSymbol) = 0 Then
            For i = 1 To aSpreadSyms.Size - 1
                If i = 1 Then
                    strSymbol = Parse(aSpreadSyms(i), ",", 2)
                Else
                    strSymbol = strSymbol & "," & Parse(aSpreadSyms(i), ",", 2)
                End If
            Next
        Else
            'remove the ~functionId
            i = InStr(strSymbol, "~")
            If i <> 0 Then strSymbol = Left(strSymbol, i - 1)
        End If
        nPrevSymbolID = 0
    ElseIf nPrevSymbolID <> m.nSymbolID Or m.eRedoMode >= eRedo9_ReloadData Then
        nPrevSymbolID = m.nSymbolID
        strSymbol = GetSymbol(m.nSymbolID)
    End If
    Symbol = strSymbol

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.Symbol.Get", eGDRaiseError_Raise

End Property

Public Property Get SymbolID() As Long
    SymbolID = m.nSymbolID
End Property

Public Property Get ExternalData() As String
    ExternalData = m.strExtData
End Property

Public Property Get SpreadSymbols() As String
    SpreadSymbols = m.strSpreadSymbols
End Property

' can pass SymbolID (numeric), Symbol (string), or External Data String (pipe-delimited string for CSI/MS file)
Public Sub SetSymbol(ByVal Symbol As Variant, Optional ByVal bGenerateNow As Boolean = False, _
                        Optional ByVal bTempAutoScale As Boolean = True)
On Error GoTo ErrSection:

    Dim nSymbolID&, strExtData$, strSpreadSymbols$, lValue&
    Dim strCopyFrom$, strReturn$
    Dim bChanged As Boolean
    
    If Not m.Form Is Nothing Then
        If m.Form.IsInGameMode Then
            If Not m.Form.GameMode Is Nothing Then
                If m.Form.GameMode.CustomOrders > 0 Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    ' get SymbolID or ExtData for new symbol (and see if has changed)
    m.bSpreadAsDollar = False
    If VarType(Symbol) = vbString Then
        If InStr(Symbol, "|") > 0 Then
            strExtData = Symbol
            If strExtData <> m.strExtData Then
                bChanged = True
            End If
        ElseIf InStr(Symbol, ";") > 0 Then
            'check if spread is using "dollar" multiplier
            strSpreadSymbols = Symbol
            m.bSpreadAsDollar = IsDollarMultiplier(strSpreadSymbols)
            If strSpreadSymbols <> m.strSpreadSymbols Then
                bChanged = True
            End If
        Else
            nSymbolID = GetSymbolID(Symbol)
        End If
    Else
        nSymbolID = Val(Symbol)
    End If
    If nSymbolID <> m.nSymbolID Then
        bChanged = True
        If m.nSymbolID <> 0 Then
            If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        End If
    End If
    
    If bChanged Then
        m.nBlankBars = m.nBlankBarsSave
        m.bAutoTrade = False
        DeactivateWhatIf False
        m.Form.tmr.Enabled = False
        m.bPriceAlertsChecked = False           '6609
        
        ' save things for old symbol
        If m.nSymbolID <> 0 Or Len(m.strExtData & m.strSpreadSymbols) > 0 Then
            StatusMsg '(to clear any notation related to previous symbol)
            TemplateSaveUserAnnots
            If Not m.Form Is Nothing Then       'aardvark 1223
                Form.vseMouse.Tag = Str(1) & Chr(9) & Str(-1)
                Form.DoMouseLabel
                UnzoomChart False
            End If
        ElseIf Len(m.strTemplate) = 0 Or m.strTemplate = "Replay" Then
            ' if template not yet loaded, then load a default
            If Not ActiveChart Is Nothing Then
                If Not ActiveChart.Chart Is Nothing Then
                    If ActiveChart.Chart.SeasonalCycleTypeEnum <> -1 Then
                        'this is to prevent usage of seasonal chart template
                        strCopyFrom = g.strAppPath & "\charts\templates\default.cht"
                    End If
                End If
            End If
            
            TemplateLoad "", , strCopyFrom
        End If
            
        ' set to new symbol
        m.ReplayTicks.Clear
        m.nSymbolID = 0
        m.strExtData = ""
        m.strSpreadSymbols = ""
        m.nTradeSymbolID = 0
        m.strTradeSymbol = ""
        m.strOptionSym = ""
        If Len(strExtData) > 0 Then
            m.strExtData = strExtData
        ElseIf Len(strSpreadSymbols) > 0 Then
            m.strSpreadSymbols = strSpreadSymbols
        Else
            m.nSymbolID = nSymbolID
        End If
        If TypeOfChart = eTypeChart_Seasonal Then m.Bars.Size = 0
        
        ' clear things for new symbol
        m.bTempAutoScale = bTempAutoScale '(false when called from restoring a chart page)
        g.RealTime.RemoveTickBuffer m.Bars
        m.Transactions.NumRecords = 0
        m.aOrders.Clear
        RedoMode = eRedo9_ReloadData
        m.nSymbolPitType = -1
        
        ProfileBarsClear        '6840
        
        ' load annots for new symbol
        'JM 01-31-2012: orginal code is not desired behavior (chart alerts should stay when changing symbol)
        '   this bug reported by Tim when working with ART Alerts
        'RemoveAlerts
        
        RemoveAnnots False, eANNOT_Icon, eANNOT_PriceAlert
        RemoveAnnots False, eANNOT_UndefinedType, eANNOT_UserAdded
        RemoveAnnots True, , eANNOT_AvgEntryLine        '4220
        RemoveAnnots False, , eANNOT_OptionInfo
        Set m.AvgEntryAnnot = Nothing
        
        TemplateLoadUserAnnots
               
        FixSpreadSymbols Me
               
        If bGenerateNow Then
            GenerateChart
        End If
        
        frmMain.SetWindowLink m.Form, eLink_Symbol
                
        m.Form.GetContractInformation
        m.Form.SetAutoExit
        m.Form.tmr.Enabled = ChartTimers
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.SetSymbol", eGDRaiseError_Raise

End Sub

Public Property Get gePixPerBar() As Long
On Error GoTo ErrSection:
    
    gePixPerBar = m.geChartStruct.x_dataPointPix

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.gePixPerBar.Get", eGDRaiseError_Raise

End Property

Public Property Let gePixPerBar(ByVal nPix As Long)
On Error GoTo ErrSection:
    
    m.geChartStruct.x_dataPointPix = nPix

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.gePixPerBar.Let", eGDRaiseError_Raise

End Property

Public Property Get PixelsPerBar() As Long
On Error GoTo ErrSection:

'Description: PPB = pixels per bar.
'This property is a semi-misnomer as it checks the PPB value and can change it
'prior to returning the PPB (TOTHINK - rework as a subroutine instead?)

    If m.nPixelsPerBar < 1 Then m.nPixelsPerBar = 7 '(default)
    PixelsPerBar = m.nPixelsPerBar

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.PixelsPerBar.Get", eGDRaiseError_Raise

End Property

Public Property Let PixelsPerBar(ByVal nPixelsPerBar As Long)
On Error GoTo ErrSection:

    Static bInProgress As Boolean

'Description:
'This property is a semi-misnomer as it inc/dec the pixel per bar value before
'doing the assignment (TOTHINK - rework as a subroutine instead?)

    Dim nMax&, nAction&, rc&
    nMax = 250
    
    If bInProgress Then Exit Property       '5303
    
    bInProgress = True
    
    If m.bZoomed = True Then
        'nAction:
        ' -1 = decrement by npixelsperbar if npixelsperbar > 1 else calculate pixels to decrement by
        '  0 = set pix per bar to npixelsperbar
        '  1 = increment by npixelsperbar if npixelsperbar > 1 else calculate pixels to increment by
        If nPixelsPerBar = -2 Then
            nAction = -1
        ElseIf nPixelsPerBar = -1 Then
            nAction = 1
        Else
            nAction = 0
        End If
        rc = geZoomPixPerBar(m.geChartObj, m.Form.pbChart.hDC, nPixelsPerBar, nAction)
        GoTo ErrExit
    End If
        
    If nPixelsPerBar = -2 Then
        ' more bars (decrease pixels per bar)
        If m.nPixelsPerBar <= 1 Then
            nPixelsPerBar = 1
            Beep
        Else
            nPixelsPerBar = m.nPixelsPerBar - 1
        End If
    ElseIf nPixelsPerBar = -1 Then
        ' less bars (increase pixels per bar)
        If m.nPixelsPerBar >= nMax Then
            nPixelsPerBar = nMax
            Beep
        Else
            nPixelsPerBar = m.nPixelsPerBar + 1
        End If
    End If
    If nPixelsPerBar < 1 Then nPixelsPerBar = 7 '(default)
    If nPixelsPerBar > nMax Then nPixelsPerBar = nMax
    
    m.nPixelsPerBar = nPixelsPerBar

ErrExit:
    bInProgress = False
    Exit Property
    
ErrSection:
    bInProgress = False
    RaiseError "cChart.PixelsPerBar.Let", eGDRaiseError_Raise

End Property

Public Property Get Annots() As cGdTree
On Error GoTo ErrSection:

    If m.Annots Is Nothing Then
        Set m.Annots = New cGdTree
    End If
    Set Annots = m.Annots

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.Annots.Get", eGDRaiseError_Raise

End Property

'Public Property Let Annots(ByVal NewAnnots As cGdTree)
'    Set m.Annots = NewAnnots
'End Property

Public Sub Repaint(Optional ByVal bReinit As Boolean = False)
On Error GoTo ErrSection:

    'MJM [begin] - added for grapheng.dll
    m.Form.pbChart.Refresh
    'MJM [end] - added for grapheng.dll

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.Repaint", eGDRaiseError_Raise

End Sub

Public Sub TemplateSave()
On Error GoTo ErrSection:

    Dim strFile$, fh%, n&, strTemp$, strName$
    Dim Sc As cPane, Ind As cIndicator, Pfp As cPatternProfit
    
    If m.Tree Is Nothing Then Exit Sub
    
    'aardvark 2950 - save drawing tools for instant replay
    If m.Form.IsInGameMode Then
        TemplateSaveUserAnnots
        Exit Sub
    End If
    
    strName = Template
    If Len(strName) = 0 Then Exit Sub
    
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strName & ".CHT"
    fh = FreeFile
    Open strFile For Output As #fh
    
    Print #fh, "[General]"
    'chart display
    Print #fh, IniString("ChartVersion", strChartVersion)
    Print #fh, IniString("Required", m.strRequired)
    Print #fh, IniString("DateApplied", Now)
    Print #fh, IniString("TemplateApplied", m.strTemplateApplied)
    Print #fh, IniString("ChartName", m.strChartName)
    Print #fh, IniString("Desc", m.strDesc)
    Print #fh, IniString("Periodicity", m.nPeriodicity)
    Print #fh, IniString("Unsplit", m.bUnsplit)
    Print #fh, IniString("ShowEmptyBars", m.bShowEmptyBars)
    Print #fh, IniString("BlankBars", m.nBlankBars)             'old templates called this ForecastBars
    Print #fh, IniString("PixelsPerBar", m.nPixelsPerBar)
    If m.nDefaultPixelsPerBar > 0 Then Print #fh, IniString("DefaultPixelsPerBar", m.nDefaultPixelsPerBar)
    Print #fh, IniString("MaxIntradayDays", m.nMaxIntradayDays)
    If m.geChartStruct.scaleArrow = 0 Then
        Print #fh, IniString("UserScreenDate", 0)
    Else
        Print #fh, IniString("UserScreenDate", gdGetNum(m.geChartStruct.gdhDate, gdGetSize(m.geChartStruct.gdhDate) - 1))
    End If
    Print #fh, IniString("FromDate", m.dFromDate)
    Print #fh, IniString("ToDate", m.dToDate)
    Print #fh, IniString("ToEndOfData", m.bToEndOfData)
    strTemp = ""
    If Not m.Bars Is Nothing Then
        If m.Bars.HasCustomSessionTime Then         '5539
            strTemp = Str(m.Bars.Prop(eBARS_DefaultStartTime)) & ";" & Str(m.Bars.Prop(eBARS_DefaultEndTime)) _
                & ";" & Str(m.Bars.Prop(eBARS_StartTime)) & ";" & Str(m.Bars.Prop(eBARS_EndTime))
        Else
            strTemp = ""
        End If
    End If
    Print #fh, IniString("StartEndTimes", strTemp)
    Print #fh, IniString("CoarseVertGrid", m.nVertGrid)
    Print #fh, IniString("HorzGridLines", m.bHorzGrid)
    Print #fh, IniString("CustomColors", m.bCustomColors)
    Print #fh, IniString("ChartForeColor", m.nChartForeColor)
    Print #fh, IniString("ChartBackColor", m.nChartBackColor)
    Print #fh, IniString("BorderForeColor", m.nBorderForeColor)
    Print #fh, IniString("BorderBackColor", m.nBorderBackColor)
    'TLB: if gradient color = the current default, then store as -1 instead
    If m.nChartGradientColor = GradientDefault Then
        Print #fh, IniString("ChartGradientColor", -1)
    Else
        Print #fh, IniString("ChartGradientColor", m.nChartGradientColor)
    End If
    Print #fh, IniString("UseGradient", m.nUseGradient)
    Print #fh, IniString("PriceTopMost", m.nPriceTopMost)
    Print #fh, IniString("SymbolID", m.nSymbolID)               'don't know if will use this
    'miscellaenous
    Print #fh, IniString("SystemID", m.nSystemID)
    Print #fh, IniString("SystemName", m.strSystemName)
    Print #fh, IniString("ShowSplitPane", m.nShowSplitPane)
    Print #fh, IniString("ShowToolbar", m.nShowToolbar)
    Print #fh, IniString("LinkedInputs", m.strLinkedInputs)
    'order bar
    Print #fh, IniString("BrOrdOverRide", 1)                    'do not move this line
    Print #fh, IniString("ShowTrades", m.nShowTrades)
    Print #fh, IniString("ProfitLineBox", m.nProfitLineBox)
    Print #fh, IniString("TradeAccountID", m.nAccountID)
    Print #fh, IniString("PseudoOrderType", m.nPseudoOrderType)
    Print #fh, IniString("ShowAccountBar", m.bAccountBar)
    Print #fh, IniString("OrdBarCtrls", m.strOrdBarCtrls)
    Print #fh, IniString("LimitTrades", m.strLimitTrades)
    Print #fh, IniString("ShowOrderHorzLine", m.nShowOrderHorzLine)
    Print #fh, IniString("ShowAvgEntryLine", m.nShowAvgEntryLine)
    Print #fh, IniString("AvgEntryProp", m.strAvgEntryProp)
    Print #fh, IniString("HighlightPos", m.nHighlightPos)
    Print #fh, IniString("HighlightEquity", m.nHighlightEquity)
    'bid/ask on chart
    Print #fh, IniString("BidAskMode", m.eBidAskMode)
    Print #fh, IniString("BidAskImage", m.eBidAskImage)
    Print #fh, IniString("BidAskImageDir", m.eBidAskImgDir)
    Print #fh, IniString("BidAskImageFill", m.nBidAskImgFill)
    Print #fh, IniString("BidAskImageSize", m.nBidAskImgSize)
    Print #fh, IniString("BidColor", m.nBidColor)
    Print #fh, IniString("AskColor", m.nAskColor)
    'seasonal chart
    If TypeOfChart = eTypeChart_Seasonal Then Print #fh, IniString("SeasonalCycle", SeasonalCycle)
    ' publishable
    If FileExist(g.strAppPath & "\SCP.flg") Then
        Print #fh, IniString("Unpublishable", m.bUnpublishable)
    End If
    
    If Not g.FractZen.Allowed Then
        m.strAutoBreakout = ""
    ElseIf Len(m.strAutoBreakout) > 0 Then
        Print #fh, IniString("AutoBreakout", m.strAutoBreakout) ' special setting for DanielCode
    End If
    If Not AllowRemoveOvernightGap Then
        m.bRemoveOvernightPriceGap = False
    ElseIf m.bRemoveOvernightPriceGap Then
        Print #fh, IniString("RemoveOvernightGap", True) ' special setting for SpyderTrader
    End If
    Print #fh, "END="
    
    For n = 1 To m.Tree.Count
        strTemp = m.Tree.Key(n) & ";" & Str(m.Tree.NodeLevel(n))
        If m.Tree.NodeLevel(n) = 0 Then
            'SUBCHART: ignore empty subcharts
            If m.Tree.NodeLevel(n + 1) > 0 Then
                Set Sc = m.Tree(n)
                Print #fh, "[SC; " & strTemp & "]"
                Sc.TemplateSave fh
            End If
        Else
            'INDICATOR
            Set Ind = m.Tree(n)
            Print #fh, "[IND; " & strTemp & "]"
            Ind.TemplateSave fh
        End If
    Next
    
    'indicator analyst
    Print #fh, "[IndAnalyst]"
    Print #fh, IniString("PfpIndicators", m.strPfpIndKeys)
    If Not m.Form Is Nothing Then Set Pfp = m.Form.PatternProfitObj
    If Not Pfp Is Nothing Then
        'general params
        Print #fh, IniString("ForecastLineStyle", Pfp.ForecastLineStyle)
        Print #fh, IniString("ForecastColor", Pfp.ForecastColor)
        Print #fh, IniString("FillColor", Pfp.FillColor)
        Print #fh, IniString("FillPattern", Pfp.FillPattern)
        Print #fh, IniString("ForecastBars", Pfp.NumForecastBars)
        Print #fh, IniString("StandardDev", Pfp.StandardDev)
        Print #fh, IniString("MinCorrelation", Pfp.PercentCorr)
        'optimization params
        Print #fh, IniString("PatternLen", Pfp.OptimizeMinBars)
        Print #fh, IniString("MaxBars", Pfp.OptimizeMaxBars)
        Print #fh, IniString("LowestCorr", Pfp.OptimizeLowestCorr)
        Print #fh, IniString("MinHits", Pfp.OptimizeMinHits)
        'matching methods
        Print #fh, IniString("WeightIndDiff", Pfp.WeightIndDiff)
        Print #fh, IniString("WeightCorrelation", Pfp.WeightCorr)
        Print #fh, IniString("WeightShape", Pfp.WeightShape)
        Print #fh, IniString("WeightDirectional", Pfp.WeightDirection)
        Print #fh, IniString("WeightSigns", Pfp.WeightSigns)
        Print #fh, IniString("WeightPeaks", Pfp.WeightPeaks)
    End If
    
    Print #fh, "END="
    
    Close #fh
    Set Sc = Nothing
    Set Ind = Nothing

    TemplateSaveUserAnnots 'strName
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.TemplateSave", eGDRaiseError_Raise

End Sub

Public Sub TemplateSaveStudy(ByVal strKey$, ByVal strName$)
On Error GoTo ErrSection:

    Dim strFile$, fh%, n&, nDesc&, strTemp$, strPane$
    Dim bSave As Boolean, bSaveInd As Boolean
    Dim Sc As cPane, Ind As cIndicator
    
    If m.Tree Is Nothing Then Exit Sub
    If Len(strName) = 0 Or Len(strKey) = 0 Then Exit Sub
    
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\Templates\" & strName & ".STU"
    fh = FreeFile
    Open strFile For Output As #fh
    
    For n = 1 To m.Tree.Count
        strTemp = m.Tree.Key(n) & ";" & Str(m.Tree.NodeLevel(n))
        If m.Tree.NodeLevel(n) = 0 Then
            'SUBCHART
            If bSave Then Exit For 'Done
            If UCase(m.Tree.Key(n)) = UCase(strKey) Then bSave = True
            If bSave Then
                strPane = m.Tree.Key(n)
                ''strTemp = strName & ";" & Str(m.Tree.NodeLevel(n))
                Set Sc = m.Tree(n)
                Print #fh, "[SC; " & strTemp & "]"
                Sc.TemplateSave fh
            End If
        ElseIf bSave Then
            'INDICATOR
            bSaveInd = False
            If UCase(strPane) = "PRICE PANE" Then
                'for price pane: save only if it or descendant is being displayed
                For nDesc = n To m.Tree.RelativeIndex(n, eTREE_LastDescendant)
                    Set Ind = m.Tree(nDesc)
                    If Ind.Display Then
                        bSaveInd = True
                        Exit For
                    End If
                Next
            Else 'if not price pane, save all indicators
                bSaveInd = True
            End If
            If bSaveInd Then
                Set Ind = m.Tree(n)
                Print #fh, "[IND; " & strTemp & "]"
                Ind.TemplateSave fh
            End If
        End If
    Next
    
ErrExit:
    Close #fh
    Set Sc = Nothing
    Set Ind = Nothing
    Exit Sub
    
ErrSection:
    Set Sc = Nothing
    Set Ind = Nothing
    RaiseError "cChart.TemplateSaveStudy", eGDRaiseError_Raise

End Sub

Public Function AnnotSymbol() As String
On Error GoTo ErrSection:

    Dim i&, strHex$, strMD5$
    Dim mb As cMemBuffer

    If m.nSymbolID <> 0 Then
        AnnotSymbol = Str(m.nSymbolID)
    ElseIf Len(m.strExtData & m.strSpreadSymbols) > 0 Then
        ' use MemBuffer since symbol of external data could have binary data
        Set mb = New cMemBuffer
        If Len(m.strExtData) > 0 Then
            mb.PutStr Parse(m.strExtData, "|", 1)
        Else
            mb.PutStr m.strSpreadSymbols
        End If
        strHex = Space(mb.Length * 2 + 1)
        i = gdHexMemory(strHex, mb.MemPtr, mb.Length)
        strHex = LCase(Trim(Left(strHex, i)))
        ' if hex string too long, then hash into an MD5 string
        If Len(strHex) > 32 Then
            strMD5 = Space(99)
            i = gdCalcMemMD5(strHex, Len(strHex), strMD5, True)
            strHex = LCase(Trim(Left(strMD5, i)))
        End If
        AnnotSymbol = "!" & strHex
        Set mb = Nothing
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.AnnotSymbol", eGDRaiseError_Raise
End Function

Private Sub TemplateSaveUserAnnots()
On Error GoTo ErrSection:

    Dim strFile$, fh%, n&, strTemp$, strName$
    Dim Annot As cAnnotation
    
    If m.Annots Is Nothing Then Exit Sub
    If m.nSymbolID = 0 And Len(m.strExtData & m.strSpreadSymbols) = 0 Then Exit Sub
    
    ' annotations for all template (Price Pane)
    fh = 0
    strName = "^" & AnnotSymbol ' Trim(Str(m.nSymbolID))
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strName & ".ANO"
    If 1 Then
        For n = 1 To m.Annots.Count
            Set Annot = m.Annots(n)
            If Annot.eUsage = eANNOT_UserAdded Then
                If Annot.Pane = "PRICE PANE" And Annot.MultiChartFlag Then
                    If Annot.IsValid Then
                        If fh = 0 Then
                            fh = FreeFile
                            Open strFile For Output As #fh
                        End If
                        Print #fh, "[ANNOT" & Str(n) & "]"
                        Annot.TemplateSave fh
                    End If
                End If
            End If
        Next
    End If
    If fh = 0 Then
        KillFile strFile
    Else
        Close #fh
    End If
    
    ' annotations for this template (chart window)
    If Len(Template) > 0 Then
        fh = 0
        strName = Template & "^" & AnnotSymbol 'Trim(Str(m.nSymbolID))
        strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strName & ".ANO"
        For n = 1 To m.Annots.Count
            Set Annot = m.Annots(n)
            If Annot.eUsage = eANNOT_UserAdded Or Annot.eUsage = eANNOT_FibClusters Then
                If Annot.Pane <> "PRICE PANE" Or Not Annot.MultiChartFlag Then
                    If Annot.IsValid Then
                        If fh = 0 Then
                            fh = FreeFile
                            Open strFile For Output As #fh
                        End If
                        Print #fh, "[ANNOT" & Str(n) & "]"
                        Annot.TemplateSave fh
                    End If
                End If
            End If
        Next
        If fh = 0 Then
            KillFile strFile
        Else
            Close #fh
        End If
    End If
    
ErrExit:
    Set Annot = Nothing
    Exit Sub
    
ErrSection:
    Set Annot = Nothing
    RaiseError "cChart.TemplateSaveUserAnnots", eGDRaiseError_Raise

End Sub

'JM - 03-12-2008
'the general section is getting unwieldy, moved and reorganized code for improved readability
Private Sub TemplateLoadGeneral(strProp$, strValue$, dValue#, bOK As Boolean, bResetMaxDays As Boolean)
On Error GoTo ErrSection:
   
    Select Case UCase(strProp)
        'charts stuff
        Case "REQUIRED"
            m.strRequired = strValue
        Case "TEMPLATEAPPLIED"
            m.strTemplateApplied = strValue
        Case "CHARTNAME"
            m.strChartName = strValue
        Case "DESC"
            m.strDesc = strValue
        Case "PERIODICITY"
            If bOK Then
                If dValue < 100 Then dValue = ePRD_Days + 1
                m.nPeriodicity = dValue
            End If
        Case "UNSPLIT"
            m.bUnsplit = dValue
        Case "SHOWEMPTYBARS"
            m.bShowEmptyBars = dValue
        Case "FORECASTBARS", "BLANKBARS"
            MyBlankBars = dValue                'historically the blank bars were saved to INI with property name of "ForecastBars"
            MyForecastBars = dValue
        Case "PIXELSPERBAR"
            PixelsPerBar = dValue
        Case "DEFAULTPIXELSPERBAR"
            m.nDefaultPixelsPerBar = dValue
        Case "MAXINTRADAYDAYS"
            If bOK Then
                If bResetMaxDays And (dValue > 10) Then
                    dValue = 10
                End If
                m.nMaxIntradayDays = dValue
            End If
        Case "USERSCREENDATE"
            m.dUserScreenDate = dValue
        Case "FROMDATE"
            If bOK Then m.dFromDate = dValue
        Case "TODATE"
            If bOK Then m.dToDate = dValue
        Case "TOENDOFDATA"
            If bOK Then m.bToEndOfData = dValue
        Case "STARTENDTIMES"
            If m.strTemplate = "Replay" Then
                m.strTemplateStartEndTimes = ""
            Else
                m.strTemplateStartEndTimes = strValue
            End If
        Case "COARSEVERTGRID"
            m.nVertGrid = dValue
            If m.nVertGrid < 0 Then m.nVertGrid = 0 'used to be boolean bCoarseGrid which would have -1 for true
        Case "HORZGRIDLINES"
            m.bHorzGrid = dValue
        Case "CUSTOMCOLORS"
            m.bCustomColors = dValue
        Case "CHARTFORECOLOR"
            m.nChartForeColor = dValue
        Case "CHARTBACKCOLOR"
            m.nChartBackColor = dValue
        Case "BORDERFORECOLOR"
            m.nBorderForeColor = dValue
        Case "BORDERBACKCOLOR"
            m.nBorderBackColor = dValue
        Case "CHARTGRADIENTCOLOR"
            m.nChartGradientColor = dValue
        Case "USEGRADIENT"
            m.nUseGradient = dValue
        Case "PRICETOPMOST"
            m.nPriceTopMost = dValue
        
        'miscellaenous stuff
        Case "SYSTEMID"
            If bOK Then m.nSystemID = dValue
        Case "SYSTEMNAME"
            If bOK Then m.strSystemName = strValue
        Case "SHOWSPLITPANE"
            'm.nShowSplitPane = dValue
            ShowSplitPane = dValue
        Case "SHOWTOOLBAR"
            If Not m.Form Is Nothing Then               '5731, 5827
                If TypeOf m.Form Is frmChart Then
                    m.nShowToolbar = dValue
                End If
            End If
        Case "LINKEDINPUTS"
            m.strLinkedInputs = strValue
        Case "SEASONALCYCLE"
            SeasonalCycle = strValue

        'order bar stuff
        Case "SHOWTRADES"
            If bOK Then m.nShowTrades = Abs(dValue)    'used to boolean
        Case "SHOWPROFITLINES"
            If m.nProfitLineBox = -1 Then m.nProfitLineBox = Abs(dValue)    'for backwards compatibility (new value is called ProfitLineBox)
        Case "PROFITLINEBOX"
            m.nProfitLineBox = dValue
        Case "TRADEACCOUNTID"
            m.nAccountID = Int(dValue)
        Case "PSEUDOORDERTYPE"
            m.nPseudoOrderType = Int(dValue)        'for setting order type combo box in form chart
        Case "SHOWACCOUNTBAR"
            m.bAccountBar = dValue
        Case "ORDBARCTRLS"
            m.strOrdBarCtrls = ConvertOrdBarButtons(strValue)
            If Len(m.strOrdBarCtrls) = 0 Then m.strOrdBarCtrls = kOrdBarDefaults  'something went wrong, just use new defaults
        Case "LIMITTRADES"
            m.strLimitTrades = strValue
        Case "SHOWORDERHORZLINE"
            m.nShowOrderHorzLine = dValue
        Case "SHOWAVGENTRYLINE"
            m.nShowAvgEntryLine = dValue
        Case "AVGENTRYPROP"
            m.strAvgEntryProp = strValue
        Case "HIGHLIGHTPOS"
            m.nHighlightPos = dValue
        Case "HIGHLIGHTEQUITY"
            m.nHighlightEquity = dValue

        'bid/ask on chart
        Case "BIDASKMODE"
            m.eBidAskMode = dValue
        Case "BIDASKIMAGE"
            m.eBidAskImage = dValue
        Case "BIDASKIMAGEDIR"
            m.eBidAskImgDir = dValue
        Case "BIDASKIMAGEFILL"
            m.nBidAskImgFill = dValue
        Case "BIDASKIMAGESIZE"
            m.nBidAskImgSize = dValue
        Case "BIDCOLOR"
            m.nBidColor = dValue
        Case "ASKCOLOR"
            m.nAskColor = dValue
        Case "AUTOBREAKOUT"
            If g.FractZen.Allowed Then
                m.strAutoBreakout = strValue
            Else
                m.strAutoBreakout = ""
            End If
        Case "REMOVEOVERNIGHTGAP"
            If AllowRemoveOvernightGap Then
                m.bRemoveOvernightPriceGap = dValue
            Else
                m.bRemoveOvernightPriceGap = False
            End If
        Case "UNPUBLISHABLE"
            m.bUnpublishable = dValue
        
        'commented out stuff
        Case "SHOWTIPS", "MINUTESPERBAR", "GENTICK", "CANDLESTICKS", "CHARTPERIOD", _
             "FONTSIZE", "GRIDLINES", "SYMBOLID", "END"
            'm.bShowTips = dValue
            '.nMinutesPerBar = dValue
            '.nGenTick = dValue
            '.bCandlesticks = dValue
            '.nChartPeriod = dValue
            'FontSize = dValue
            ''m.nGridLines = dValue
            'just so won't error below
        Case Else:
            'If FileExist("c:\COMMON\files.exe") Then
            '    InfBox "i=[] ; h=Chart Template ; Not found: " & strType
            'End If
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.TemplateLoadGeneral"

End Sub

Public Function TemplateLoad(Optional ByVal strName$ = "", _
        Optional ByVal bWithUserAnnots As Boolean = True, _
        Optional ByVal strCopyFrom$ = "") As Boolean
On Error GoTo ErrSection:

    Dim strFile$, fh%, n&
    Dim i&, strSection$, strProp$, strValue$, dValue#
    Dim idxPane&, idx&
    
    Dim bClearName As Boolean
    Dim bBrOrdDone As Boolean
    
    Dim Sc As cPane, Ind As cIndicator
    Dim bOK As Boolean, bResetMaxDays As Boolean
    Dim bAddSidewinder As Boolean, bAddPivots As Boolean
    Dim bResetGridLines As Boolean
    
    Dim strLinkedInputInUse As String
    Dim aLinkedInputs As cGdArray
    
    m.nHighlightPos = -2        'reset
    m.nHighlightEquity = -2
    
    bOK = True
    If Not m.Form Is Nothing Then
        If Not m.Form.GameMode Is Nothing Then
            If m.Form.GameMode.CustomOrders > 0 Then
                bOK = False
            End If
        End If
    End If
    
    If Len(strName) > 0 Then
        ' first kill the obsolete files
        If Len(m.strTemplate) > 0 And UCase(m.strTemplate) <> UCase(strName) Then
            If m.Form.ImgSrvState = eImgSrv_Busy Then
                KillFile g.ChartGlobals.strCPCRoot & "\Charts\" & m.strTemplate & ".*"
            End If
        End If
        ' assign name of new template
        m.strTemplate = strName
    ElseIf Len(m.strTemplate) > 0 Then
        ' name hasn't changed
        strName = m.strTemplate
    Else
        ' create a unique name
        strName = GetUnusedChartName
        m.strTemplate = strName
    End If
    
    ' if does not yet exist, then get a copy as a default
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strName & ".CHT"
    If Not FileExist(strFile) Then
        If Len(strCopyFrom) = 0 Then
            ' first try getting a copy of the active chart's template
            On Error Resume Next
            strCopyFrom = ActiveChart.Chart.Template
            If Len(strCopyFrom) > 0 Then
                ActiveChart.Chart.TemplateSave
                strCopyFrom = g.ChartGlobals.strCPCRoot & "\Charts\" & strCopyFrom & ".CHT"
                If Not FileExist(strCopyFrom) Then strCopyFrom = ""
            End If
            On Error GoTo ErrSection:
            If Len(strCopyFrom) = 0 Then
                ' else use the Default template
                strCopyFrom = g.ChartGlobals.strCPCRoot & "\Charts\Templates\Default.CHT"
            End If
        End If
        If FileExist(strCopyFrom) Then
            FileCopy strCopyFrom, strFile
            bClearName = True
        Else
            GoTo ErrSection:
        End If
    End If
        
    fh = FreeFile
    Open strFile For Input As #fh
    
    RemoveAlerts        'necessary when changing templates)
                
    Annots.Clear
    m.Tree.Clear
    Set m.Tree = Nothing
    i = Tree.Count
    
    LoadDefaultSettings
    
    m.nProfitLineBox = -1
    m.nUseGradient = -1 ' i.e. not used yet (so can fix it later if still not used)
    Do While Not EOF(fh)
        Line Input #fh, strProp
        If Left(strProp, 1) = "'" Then
            ' this is a comment line and should be ignored
        ElseIf Left(strProp, 1) = "[" Then
            strSection = strProp
            With Tree
                idx = 0
                Set Ind = Nothing
                Select Case UCase(Parse(strSection, ";", 1))
                    Case "[SC"
                        Set Sc = New cPane
                        Sc.TemplateLoad fh
                        If Sc.SeasonalPaneFlag = 1 Then m.eTypeOfChart = eTypeChart_Seasonal
                        idx = .Add(Sc)
                        idxPane = idx
                    Case "[IND"
                        Set Ind = New cIndicator
                        Ind.TemplateLoad fh
                        idx = .Add(Ind)
                        ' TLB 8/5/2008: check if need to add Sidewinder and/or Pivots
                        If UCase(Ind.CodedName) = "WCCISIDEWINDER" Then
                            bAddSidewinder = False
                        End If
                        If InStr(UCase(Ind.CodedName), "PIVOT") > 0 Then
                            bAddPivots = False
                        End If
                        If Ind.DataType = eINDIC_ProfileBars Then
                            If Sc.Display And Ind.Display Then m.bProfileChart = True
                        End If
                        If UCase(Ind.Name) = "PRICE" Then
                            If Not Sc Is Nothing Then Sc.PricePaneFlag = 1          '6799
                        End If
                        'get linked parms for "clean up" purposes
                        strLinkedInputInUse = strLinkedInputInUse & "|" & Ind.LinkedParmsString
                End Select
                If idx > 0 Then
                    .NodeLevel(idx) = Val(Parse(strSection, ";", 3))
                    ' set key to original (if does not already exist)
                    strValue = Trim(Parse(strSection, ";", 2))
                    If Not .Exists(strValue) Then
                        .Key(idx) = strValue
                    End If
                    If Not Ind Is Nothing Then
                        Ind.ConvertHighlightBar m.Tree, idx
                        Ind.MyKey = .Key(idx)
                    End If
                End If
            End With
        ElseIf Trim(UCase(strSection)) = "[GENERAL]" Then
            If ParseIniString(strProp, strValue, dValue) Then
                If UCase(strProp) = "CHARTVERSION" Then
                    ' make sure major version matches
                    If Val(Parse(strValue, ".", 1)) <> Val(Parse(strChartVersion, ".", 1)) Then
                        GoTo ErrSection:
                    End If
                    ' see if should reset MaxIntradayDays
                    If Val(strValue) < 3.1 Then
                        bResetMaxDays = True
                    End If
                    ' see if should reset Gridlines (due to new dots style)
                    If Val(strValue) < 3.11 Then
                        bResetGridLines = True
                    End If
                ElseIf UCase(strProp) = "ORDBARCTRLS" Then
                    'this is a one time override so bracket order button will show
                    If bBrOrdDone Then
                        TemplateLoadGeneral strProp, strValue, dValue, bOK, bResetMaxDays
                    'else just use the default which will show bracket order button if user has never explicity turned it off
                    End If
                Else
                    TemplateLoadGeneral strProp, strValue, dValue, bOK, bResetMaxDays
                    ' TLB 8/5/2008: check if need to add Sidewinder and/or Pivots
                    If UCase(strProp) = "TEMPLATEAPPLIED" And UCase(Left(strValue, 4)) = "WCCI" Then
                        bAddSidewinder = True
                        bAddPivots = True
                    ElseIf strProp = "BrOrdOverRide" Then
                        bBrOrdDone = True
                    End If
                End If
            End If
        ElseIf Trim(UCase(strSection)) = "[INDANALYST]" Then
            If ParseIniString(strProp, strValue, dValue) Then
                If UCase(strProp) = "PFPINDICATORS" Then m.strPfpIndKeys = strValue
            End If
        End If
    Loop
    
    If bClearName Then m.strChartName = ""
    If m.nSystemID = 0 And m.nShowTrades = 1 Then m.nShowTrades = 0
    If m.nProfitLineBox < 0 Then m.nProfitLineBox = 0
        
    ' if no "border color" properties yet, default all to 0 (in case of an old template)
    If m.nBorderBackColor = 0 And m.nBorderForeColor = 0 Then
        m.nChartBackColor = 0
        m.nBorderForeColor = 0
    End If
    
    ' TLB 5/16/2014: reset defaults for new style of "gridlines" (which are now dots)
    If bResetGridLines Then
        ' turn vert and horiz gridlines on by default -- then they can turn them off if desired
        m.bHorzGrid = True
        If m.nVertGrid = 2 Then     '0=coarse, 1=fine, 2=none
            m.nVertGrid = 0
        End If
        ' and change chart forecolor from old default to a darker color
        If m.nChartForeColor = RGB(192, 192, 192) Then
            m.nChartForeColor = RGB(128, 128, 128)
        End If
    End If
    
    ' TLB: fix some things for new gradient properties
    If m.nChartGradientColor = -1 Then  ' -1 means use the current default
        m.nChartGradientColor = GradientDefault
    End If
    If m.nUseGradient = -1 Then ' i.e. gradient props had not been saved to template yet
        m.nChartGradientColor = GradientDefault
        m.nUseGradient = 0
    End If
    
    If Not m.Tree.Exists("PRICE PANE") Then
        AddStudyToChart Me, "PRICE", True
    End If
    
    Set Ind = m.Tree("PRICE")
    If Not Ind Is Nothing Then
        Set m.Bars = Ind.Bars
    End If
    
    ' TLB 8/5/2008: check if need to add Sidewinder and/or Pivots
    If bAddSidewinder Then
        Set Ind = New cIndicator
        With Ind
            .CodedName = "WCCISidewinder"
            .Display = False
            .Color = 0
        End With
        idx = m.Tree.RelativeIndex("PRICE PANE", eTREE_NextSibling)
        If idx > 0 Then
            idx = m.Tree.Add(Ind, , idx, eTREE_FirstChild)
        End If
    End If
    If bAddPivots Then
        If TemplateAddStudy("Woodies Pivots") Then
            idx = m.Tree.Index("PRICE PANE")
            For idx = idx + 1 To m.Tree.Count
                If m.Tree.NodeLevel(idx) <= 0 Then
                    Exit For
                Else
                    Set Ind = m.Tree(idx)
                    If Not Ind Is Nothing Then
                        Select Case UCase(Ind.CodedName)
                        Case "WOODIESPIVOT", "WOODIESR1", "WOODIESR2", "WOODIESR3", _
                                "WOODIESS1", "WOODIESS2", "WOODIESS3"
                            Ind.Display = False
                        End Select
                    End If
                End If
            Next
        End If
    End If
    
    'clear out unused linked inputs
    If Len(m.strLinkedInputs) > 0 Then
        Set aLinkedInputs = New cGdArray
        aLinkedInputs.SplitFields m.strLinkedInputs, "|"
        For i = aLinkedInputs.Size - 1 To 0 Step -1
            If InStr(strLinkedInputInUse, Parse(aLinkedInputs(i), ";", 1)) = 0 Then
                aLinkedInputs.Remove i
            End If
        Next
        
        If aLinkedInputs.Size > 0 Then
            m.strLinkedInputs = aLinkedInputs.JoinFields("|")
        Else
            m.strLinkedInputs = ""
        End If
    End If
        
    If bWithUserAnnots Then
        TemplateLoadUserAnnots
    End If
        
    FixColors
    
    m.bTempAutoScale = True
    m.nPrevNumPanes = 0
    
    TemplateLoad = True
    frmChartOnOff.ClearPrevChart
    
ErrExit:
    On Error Resume Next
    RedoMode = eRedo9_ReloadData
    If fh <> 0 Then Close #fh
    Set Sc = Nothing
    Set Ind = Nothing
    ' make sure there's at least a price pane
    If m.Tree.Count = 0 Then
        AddStudyToChart Me, "PRICE", True
    End If
    Exit Function

ErrSection:
    TemplateLoad = False
    Resume ErrExit

End Function

Public Function TemplateAddStudy(ByVal strName$) As Boolean
On Error GoTo ErrSection:

    Dim strFile$, fh%, n&, strTemp$, strScName$, strScDesc$
    Dim i&, strSection$, strProp$, strValue$, dValue#
    Dim idxPane&, idx&, idxAddInd&
    Dim bPricePane As Boolean, bAddInd As Boolean
    Dim Sc As cPane, Ind As cIndicator
    
    Dim iRibbonCount&, iParm&, iExist&
    Dim strInputName$, strInputVal$, strInputType$, strLinkedInputs$
    Dim aExistLink As New cGdArray
    
    If Len(strName) = 0 Then GoTo ErrSection:
    
    If strName = "Volume Profile" Then
        If Not HasModule("TPRO") Then
            GoTo ErrExit
        ElseIf SecurityType(Symbol) = "I" Then
            InfBox "This study cannot be applied to Forex or Index charts.", , , "Volume Profile Study"
            GoTo ErrExit
        Else
            Set Ind = HasProfileVolume
            If Not Ind Is Nothing Then
                If Not Ind.Display Then
                    Ind.Display = True
                Else
                    InfBox "This chart already has a Volume Profile Study.", , , "Volume Profile Study"
                End If
                GoTo ErrExit
            End If
        End If
    End If
    
    If Len(m.strLinkedInputs) > 0 Then
        aExistLink.SplitFields m.strLinkedInputs, "|"
    End If
    
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\Templates\" & strName & ".STU"
    If Not FileExist(strFile) Then
        'check default "My Collection" path     - aardvark 6924 (EWI)
        strFile = g.strAppPath & "\Charts\Templates\" & strName & ".STU"
        
        If Not FileExist(strFile) Then GoTo ErrSection:
    End If
    
    fh = FreeFile
    Open strFile For Input As #fh
    
    idxAddInd = -1 '(default: add to end)
    Do While Not EOF(fh)
        Line Input #fh, strProp
        If Left(strProp, 1) = "[" Then
            strSection = strProp
            With Tree
                idx = 0
                Set Ind = Nothing
                Select Case UCase(Parse(strSection, ";", 1))
                    Case "[SC"
                        iRibbonCount = 0        'reset
                        Set Sc = New cPane
                        Sc.TemplateLoad fh
                        Sc.Display = True
                        strScName = Sc.Name
                        strScDesc = Sc.Desc
                        If UCase(Trim(Parse(strSection, ";", 2))) = "PRICE PANE" Then
                            If .Exists("PRICE PANE") Then bPricePane = True
                        End If
                        If bPricePane Then
                            ' work with existing Price Pane
                            Set Sc = .Item("PRICE PANE")
                            Sc.Display = True
                            Sc.Name = strScName
                            Sc.Desc = strScDesc
                            idxPane = .Index("PRICE PANE")
                            If .Exists("PRICE") Then
                                idxAddInd = .RelativeIndex("PRICE", eTREE_LastDescendant) + 1
                            Else
                                idxAddInd = .RelativeIndex("PRICE PANE", eTREE_LastDescendant) + 1
                            End If
                        Else
                            ' add new Pane
                            idx = .Add(Sc)
                            idxPane = idx
                            idxAddInd = -1 '(at end)
                        End If
                    Case "[IND"
                        bAddInd = True
                        If bPricePane Then
                            If UCase(Trim(Parse(strSection, ";", 2))) = "PRICE" Then
                                If .Exists("PRICE") Then bAddInd = False
                            End If
                        End If
                        If bAddInd Then
                            Set Ind = New cIndicator
                            Ind.TemplateLoad fh
                            
                            If Ind.Name = "Volume Profile" Then
                                If m.Bars.IsIntraday Then
                                    Ind.ProfilePeriodicityStr = "d"
                                ElseIf m.Bars.Prop(eBARS_PeriodType) = ePRD_Days Then
                                    Ind.ProfilePeriodicityStr = "w"
                                ElseIf m.Bars.Prop(eBARS_PeriodType) = ePRD_Weeks Then
                                    Ind.ProfilePeriodicityStr = "q"
                                ElseIf m.Bars.Prop(eBARS_PeriodType) = ePRD_Months Then
                                    Ind.ProfilePeriodicityStr = "q"
                                ElseIf m.Bars.Prop(eBARS_PeriodType) = ePRD_Quarters Then
                                    Ind.ProfilePeriodicityStr = "y"
                                Else
                                    Ind.ProfilePeriodicityStr = m.Bars.Prop(eBARS_PeriodicityStr)
                                End If
                                If SecurityType(m.Bars) = "S" Then
                                    Ind.TicksPerRow = 20
                                ElseIf SecurityType(m.Bars) = "F" Then
                                    Ind.TicksPerRow = 4
                                Else
                                    Ind.TicksPerRow = 10
                                End If
                            End If
                            'clear coded text so will get rebuilt
                            Ind.CodedText = ""
                            idx = .Add(Ind, , idxAddInd, eTREE_Myself)
                            If idxAddInd >= 0 Then idxAddInd = idxAddInd + 1
                            
                            'JM 03-08-2011: original design allows on one Ribbon pair per pane
                            If Ind.DisplayType = eINDIC_Ribbon And UCase(Ind.Name) <> "POWERZONES" Then
                                If iRibbonCount = 2 Then
                                    Ind.DisplayType = eINDIC_Line
                                Else
                                    iRibbonCount = iRibbonCount + 1
                                End If
                            End If
                            
                            'build linked inputs string for chart object
                            If Len(Ind.LinkedParmsString) > 0 Then
                                For iParm = 1 To Ind.ParmCount
                                    Ind.LinkedParmGet iParm, strInputName, strInputVal
                                    If Len(strInputName) > 0 And Len(strInputVal) > 0 Then
                                        strInputType = Str(Ind.ParmType(iParm))
                                        If Not aExistLink Is Nothing Then
                                            For iExist = 0 To aExistLink.Size - 1
                                                If Parse(aExistLink(iExist), ";", 1) = strInputName Then
                                                    If strInputType <> Parse(aExistLink(iExist), ";", 3) Then
                                                        'name exists, but type is different, do not use linked input from study
                                                        Ind.LinkedParmsDelete strInputName, iParm
                                                    End If
                                                    'linked input name exists in current chart, don't add
                                                    strInputName = ""
                                                    strInputVal = ""
                                                    strInputType = ""
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                        
                                        If Len(strInputName) > 0 And Len(strInputVal) > 0 And Len(strInputType) > 0 Then
                                            aExistLink.Add strInputName & ";" & strInputVal & ";" & strInputType
                                            If Len(strLinkedInputs) > 0 Then
                                                strLinkedInputs = strLinkedInputs & "|" & strInputName & ";" & strInputVal & ";" & strInputType
                                            Else
                                                strLinkedInputs = strInputName & ";" & strInputVal & ";" & strInputType
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                End Select
                If idx > 0 Then
                    .NodeLevel(idx) = Val(Parse(strSection, ";", 3))
                    If .NodeLevel(idx) > 0 And bAddInd Then
                        ' set key to original (if does not already exist) - 4314
                        strValue = Trim(Parse(strSection, ";", 2))
                        If strValue = Ind.GroupKey Then
                            If Not .Exists(strValue) Then
                                .Key(idx) = strValue
                            End If
                        End If
                    End If
                    If Not Ind Is Nothing Then Ind.ConvertHighlightBar m.Tree, idx
                End If
            End With
        End If
    Loop
    
    If Len(strLinkedInputs) > 0 Then
        If Len(m.strLinkedInputs) > 0 Then
            m.strLinkedInputs = m.strLinkedInputs & "|" & strLinkedInputs
        Else
            m.strLinkedInputs = strLinkedInputs
        End If
    End If
    
    Set Ind = Nothing
    Set Sc = Nothing
    
    If idxAddInd <> 0 Then TemplateAddStudy = True
    
ErrExit:
    On Error Resume Next
    RedoMode = eRedo9_ReloadData
    If fh <> 0 Then Close #fh
    Set Sc = Nothing
    Set Ind = Nothing
    Exit Function

ErrSection:
    TemplateAddStudy = False
    Resume ErrExit

End Function

Private Sub TemplateLoadUserAnnots()
On Error GoTo ErrSection:

    Dim strFile$, fh%, n&, strTemp$, strName$, strSymbol$, s$
    Dim i&, j&, idx&
    Dim strSection$, strType$, strValue$, nValue#
    Dim bGood As Boolean
    Dim Annot As cAnnotation
    
    Dim IndClusterPrice As cIndicator
    Dim AnnotClusterZone As cAnnotation
    
    Dim gdTempTree As New cGdTree
    Dim tbAnnotDate As New cGdTable
    Dim aIndex As New cGdArray
    
    ' first remove existing user annots
    RemoveAnnots False, eANNOT_UndefinedType, eANNOT_UserAdded
    RemoveAnnots False, , eANNOT_FibClusters        '6268
    If m.nSymbolID = 0 And Len(m.strExtData & m.strSpreadSymbols) = 0 Then Exit Sub
    
    ' create table
    tbAnnotDate.CreateField eGDARRAY_Longs, 0, "Index"
    tbAnnotDate.CreateField eGDARRAY_Doubles, 1, "Date"
    
    'look for price cluster indicator
    If Not m.Tree Is Nothing Then Set IndClusterPrice = m.Tree(kClusterPriceKey)

    strSymbol = GetSymbol(m.nSymbolID)

    ' annotations for all templates (Price Pane annotations)
    strName = "^" & AnnotSymbol ' Trim(Str(m.nSymbolID))
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strName & ".ANO"
    ' TLB 4/16/2016: convert drawing tools from OLD Forex symbols (GMT-style) to NEW Forex symbols (ET-style)
    If Len(strSymbol) = 8 And Left(strSymbol, 1) = "$" And Mid(strSymbol, 5, 1) = "-" Then
        If Not FileExist(strFile) Then
            ' if doesn't exist yet for NEW Forex symbol, then rename from OLD Forex symbol (if that does exist)
            i = GetSymbolID(strSymbol & "@GMT")
            If i > 0 Then
                s = Parse(strFile, "^", 1) & "^" & Str(i) & ".ANO"
                If FileExist(s) Then
                    RenameFile s, strFile
                End If
            End If
        End If
    End If
    If FileExist(strFile) Then
        fh = FreeFile
        Open strFile For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strType
            If Left(strType, 1) = "[" Then
                strSection = strType
                If Left(strSection, 6) = "[ANNOT" Then
                    Set Annot = New cAnnotation
                    Annot.TemplateLoad fh, Me
                    Annot.MultiChartFlag = True
                    If Annot.IsValid Then
                        i = gdTempTree.Add(Annot)
                        tbAnnotDate.AddRecord ""
                        tbAnnotDate(0, tbAnnotDate.NumRecords - 1) = i
                        tbAnnotDate(1, tbAnnotDate.NumRecords - 1) = Annot.CreateDate
                    End If
                End If
            End If
        Loop
        Close #fh
    End If
    
    ' annotations only for this template (non Price Pane)
    strName = Template & "^" & AnnotSymbol ' Trim(Str(m.nSymbolID))
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strName & ".ANO"
    ' TLB 4/16/2016: convert drawing tools from OLD Forex symbols (GMT-style) to NEW Forex symbols (ET-style)
    If Len(strSymbol) = 8 And Left(strSymbol, 1) = "$" And Mid(strSymbol, 5, 1) = "-" Then
        If Not FileExist(strFile) Then
            ' if doesn't exist yet for NEW Forex symbol, then rename from OLD Forex symbol (if that does exist)
            i = GetSymbolID(strSymbol & "@GMT")
            If i > 0 Then
                s = Parse(strFile, "^", 1) & "^" & Str(i) & ".ANO"
                If FileExist(s) Then
                    RenameFile s, strFile
                End If
            End If
        End If
    End If
    If Len(Template) > 0 And FileExist(strFile) Then
        fh = FreeFile
        Open strFile For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strType
            If Left(strType, 1) = "[" Then
                strSection = strType
                If Left(strSection, 6) = "[ANNOT" Then
                    Set Annot = New cAnnotation
                    Annot.TemplateLoad fh, Me
                    Annot.MultiChartFlag = False
                    If Annot.IsValid Then
                        i = gdTempTree.Add(Annot)
                        tbAnnotDate.AddRecord ""
                        tbAnnotDate(0, tbAnnotDate.NumRecords - 1) = i
                        tbAnnotDate(1, tbAnnotDate.NumRecords - 1) = Annot.CreateDate
                    End If
                End If
            End If
        Loop
        Close #fh
    End If
    
    Dim dAnnDate#, dAnnMin#, dAnnMax#
    
    'sort annotations based on create date then add to chart object's annotations tree
    aIndex.Create eGDARRAY_Longs, tbAnnotDate.NumRecords
    gdSortAsIndex aIndex.ArrayHandle, tbAnnotDate.FieldArrayHandle(1), 1, eGdSort_Default, 0, -1
    j = aIndex.Size - 1
    For i = 0 To j
        Set Annot = gdTempTree(tbAnnotDate(0, aIndex(i)))
        If Not Annot Is Nothing Then
            With Annots
                idx = .Add(Annot)
                If idx > 0 Then
                    ' set key to saved key if key was saved         '09-09-2004
                    strValue = Annot.Prop("AnnotKey")
                    If Annot.eUsage = eANNOT_FibClusters Then
                        .Key(idx) = kClusterZoneRect
                        Set AnnotClusterZone = Annot
                    ElseIf InStr(UCase(strValue), "AUTOKEY") > 0 And Not .Exists(strValue) Then
                        .Key(idx) = strValue
                    Else
                        Annot.Prop("AnnotKey") = .Key(idx)          'backwards compatibility
                    End If
                End If
            End With
        End If
    Next
    
    If Not IndClusterPrice Is Nothing Then
        'annot files are often deleted manually by techs or developers for troubleshooting purposes
        'users may be sharing a .CHT file for fib clusters but are not aware that the zone is in an .ANO file
        If AnnotClusterZone Is Nothing Then
            i = m.Tree.Index("PRICE PANE")
            
            'annot for cluster zone does not exist, create it
            Set Annot = New cAnnotation
            
            Annot.CreateNew Me, eANNOT_Rectangle, i, 0, 0, IndClusterPrice.ClusterEndDate, 0
            
            Annot.eUsage = eANNOT_FibClusters
            Annot.Prop("Shape") = 0
            Annot.PreIndicator = 1
            Annot.Prop("FillPattern") = 1
            Annot.Prop("FillColor") = Annot.ClusterPropDefault("ZoneFillColor")
            Annot.Color = Annot.ClusterPropDefault("ZoneColor")
            Annot.Prop("AnnotKey") = kClusterZoneRect
            
            idx = m.Annots.Add(Annot, kClusterZoneRect)
            
            'set to true so the dates for annotation will get set when price cluster gets calculated
            IndClusterPrice.Display = True
        End If
    End If
        
ErrExit:
    Set Annot = Nothing
    Set gdTempTree = Nothing
    Set tbAnnotDate = Nothing
    Set aIndex = Nothing
    
    Exit Sub

ErrSection:
    Set Annot = Nothing
    Set gdTempTree = Nothing
    Set tbAnnotDate = Nothing
    Set aIndex = Nothing
    
    RaiseError "cChart.TemplateLoadUserAnnots", eGDRaiseError_Raise
End Sub

Public Property Get Template() As String
    Template = m.strTemplate
End Property

Public Property Let Template(ByVal strTemplate As String)
    m.strTemplate = strTemplate
    'Should we load it here???
    
End Property

'Public Function RemoveAnnots(ByVal bJustLastAnnot As Boolean, _
'        Optional ByVal eType As eAnnotType = eANNOT_UndefinedType, _
'        Optional ByVal eUsage As eAnnotUsage = eANNOT_UserAdded, _
'        Optional ByVal bRemoveAll As Boolean = False) As Long

Public Function RemoveAnnots(ByVal bJustLastAnnot As Boolean, _
        Optional ByVal eType As eAnnotType = eANNOT_UndefinedType, _
        Optional ByVal eUsage As eAnnotUsage = eANNOT_UserAdded, _
        Optional ByVal bSaveForUndo As Boolean = True, _
        Optional ByRef annotSelected As cAnnotation = Nothing, _
        Optional ByVal dPriorToDate As Double = 0) As Long
On Error GoTo ErrSection:

    Dim i&, nRemoved&
    Dim bRemove As Boolean
    Dim Annot As cAnnotation
    
    If Annots.Count = 0 Then GoTo ErrExit
    
    For i = Annots.Count To 1 Step -1
        bRemove = False
        Set Annot = Annots(i)
        If Not Annot Is Nothing Then
            If Annot.eUsage = eANNOT_WhatIf Or (Annot.eUsage = eANNOT_FibClusters And eUsage <> eANNOT_FibClusters) Then        '6268
                'do nothing
            ElseIf Annot.eUsage = eANNOT_PatternProfit Then
                If eUsage = eANNOT_PatternProfit Then
                    bRemove = True
                End If
            ElseIf eUsage = Annot.eUsage Or eUsage = eANNOT_UndefinedUsage Then
                If eType = Annot.eType Or eType = eANNOT_UndefinedType Then
                    If bJustLastAnnot Then
                        If bSaveForUndo And Annot.eUsage = eANNOT_UserAdded Then
                            LastEditCreate Annot, False
                        End If
                        If Annot.eType <> eANNOT_MultiRects And Annot.eType <> eANNOT_BellAlert Then
                            bRemove = True
                        End If
                    ElseIf Not Annot Is annotSelected Then
                        'remove all or remove prior to date
                        If eUsage = eANNOT_UserAdded And Annot.eType = eANNOT_HorzLine And Annot.Prop("FakePriceAlert") = 1 Then
                            bRemove = False
                        Else
                            bRemove = True
                            If dPriorToDate > 0 Then
                                If Annot.dDate(1) > dPriorToDate Then
                                    bRemove = False
                                ElseIf Annot.dDate(2) > dPriorToDate Then
                                    bRemove = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            If bRemove Then
                Annot.geRemoveAnnotation m.geChartObj
                Annots.Remove i
                nRemoved = nRemoved + 1
                If bJustLastAnnot Then Exit For
            End If
        End If
    Next
    
    Set Annot = Nothing
    
    RemoveAnnots = nRemoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.RemoveAnnots", eGDRaiseError_Raise

End Function

'returns constant for Peg Line Type
'Private Function PegLineType(ByVal nStyle&, Optional ByVal bHorzLine As Boolean = False) As Long

' SAVE CODE for reference
'    Select Case nStyle
'        Case eINDIC_Thin:
'            nType = PELT_THINSOLID
'        Case eINDIC_MediumThin:
'            nType = PELT_MEDIUMTHINSOLID
'        Case eINDIC_Medium:
'            nType = PELT_MEDIUMSOLID
'        Case eINDIC_MediumThick:
'            nType = PELT_MEDIUMTHICKSOLID
'        Case eINDIC_Thick:
'            nType = PELT_THICKSOLID
'        Case eINDIC_ExtraThick:
'            nType = PELT_EXTRATHICKSOLID
'        Case eINDIC_Dash:
'            nType = PELT_DASH
'        Case eINDIC_Dot:
'            nType = PELT_DOT
'        Case eINDIC_DashDot:
'            nType = PELT_DASHDOT
'        Case Else:
'            nType = PELT_THINSOLID
'    End Select

'End Function

'Returns last "bar #" with good data (non-null close)
'- if bAsChartX is false:  bar # in the bars object
'- if bAsChartX is true:  "X" value for chart
Public Function LastGoodDataBar(ByVal bAsChartX As Boolean, Optional ByVal bOnScreen As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim nX&, nBar&, n&
    
    If bOnScreen Then
        ' find last good bar of data on the visible screen
        nX = -1
        For n = ScreenEndX To ScreenStartX Step -1
            If n <= m.nLastGoodDataBarX Then
                nBar = m.aXBar.Num(n)
                If nBar >= 0 Then
                    If m.Bars(eBARS_Close, nBar) <> kNullData Then
                        nX = n
                        Exit For
                    End If
                End If
            End If
        Next
    Else
        nX = m.nLastGoodDataBarX
    End If
    
    If nX < 0 Then
        LastGoodDataBar = kNullData
    ElseIf bAsChartX Then
        LastGoodDataBar = nX
    Else
        LastGoodDataBar = m.aXBar(nX)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.LastGoodDataBar", eGDRaiseError_Raise

End Function

Public Sub SyncToolbar(Optional ByVal bSyncNow As Boolean = False, Optional ByVal bResyncStudyBtns As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, strType$, nPPB&, bSync As Boolean
    Dim nSymbolID&, strSymbol$
    Dim strSub$, strSector$, strCmp$                'shown in caption of button
    Dim strSubDesc$, strSecDesc$, strCmpDesc$       'shown in tooltip of button
    Dim strPeriodCbo$, strPeriod$
    
    Dim bSaveRedraw As Boolean
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If m.nPartialLoadDate > 0 Then Exit Sub         '5166
    
    If Not m.Form Is Nothing Then
        If bSyncNow Then
            bSync = True
        ElseIf m.Form Is Screen.ActiveForm Then
            bSync = True
        ElseIf m.Form Is ActiveChart Then
            bSync = True
        End If
            
        If bSync Then
            Select Case BarDisplayType
            Case eINDIC_Line
                strType = "ID_CloseLine"
            Case eINDIC_Candlestick
                strType = "ID_Candlesticks"
            Case eINDIC_Area
                strType = "ID_Mountain"
            Case eINDIC_OHLC
                strType = "ID_OHLCBars"
            Case eINDIC_BollingerBar
                strType = "ID_BollingerBars"
            End Select
        
            strPeriodCbo = GetPeriodStr(m.nPeriodicity) ' m.Bars.Prop(eBARS_PeriodicityStr)
            If g.bHideAutoBreakoutNumber And Len(m.strAutoBreakout) > 0 Then
                strPeriodCbo = "FractZen" '"Auto Breakout" ' TLB 6/16/2011: Needham can now hide the exact #
            End If
        
            strPeriod = "ID_CustomPeriod"
            nPPB = m.Bars.Prop(eBARS_PeriodsPerBar)
            Select Case m.Bars.Prop(eBARS_PeriodType)
            Case ePRD_IntRenko, ePRD_EodRenko
                strType = "ID_Renko"
            Case ePRD_IntKagi, ePRD_EodKagi
                strType = "ID_Kagi"
            Case ePRD_IntPF, ePRD_EodPF
                strType = "ID_PointFigure"
            Case ePRD_Days
                If nPPB = 1 Then strPeriod = "ID_Daily"
            Case ePRD_Weeks
                If nPPB = 1 Then strPeriod = "ID_Weekly"
            Case ePRD_Months
                If nPPB = 1 Then strPeriod = "ID_Monthly"
            Case ePRD_Quarters
                If nPPB = 1 Then strPeriod = "ID_Quarterly"
            Case ePRD_Years
                If nPPB = 1 Then strPeriod = "ID_Yearly"
            Case ePRD_Minutes
                If nPPB = 1 Then
                    strPeriod = "ID_1minute"
                ElseIf nPPB = 3 Then
                    strPeriod = "ID_3minute"
                ElseIf nPPB = 5 Then
                    strPeriod = "ID_5minute"
                ElseIf nPPB = 10 Then
                    strPeriod = "ID_10minute"
                ElseIf nPPB = 15 Then
                    strPeriod = "ID_15minute"
                ElseIf nPPB = 30 Then
                    strPeriod = "ID_30minute"
                ElseIf nPPB = 60 Then
                    strPeriod = "ID_60minute"
                ElseIf nPPB = 90 Then
                    strPeriod = "ID_90minute"
                ElseIf nPPB = 120 Then
                    strPeriod = "ID_120minute"
                ElseIf nPPB = 180 Then
                    strPeriod = "ID_180minute"
                ElseIf nPPB = 240 Then
                    strPeriod = "ID_240minute"
                ElseIf nPPB = 360 Then
                    strPeriod = "ID_360minute"
                Else
                    strPeriod = "ID_CustomMinute"
                End If
            End Select
        
            GetSectorInfo strSector, strSub, strCmp, strSecDesc, strSubDesc, strCmpDesc
            SyncChartingBtns m.Form, strSector, strSecDesc, strSub, strSubDesc, strCmp, strCmpDesc, strPeriodCbo, strPeriod, bResyncStudyBtns
        
        End If
            
        If bSync And Not m.tbToolbar Is Nothing Then
            
            With m.tbToolbar
                bSaveRedraw = .Redraw
                .Redraw = False
                
                .Tools("ID_BarPeriod").ComboBox.Text = strPeriodCbo
                
                If m.Bars.Prop(eBARS_PeriodType) = ePRD_Minutes Then    '5266
                    If .Tools(strPeriod).Name <> Str(nPPB) Then
                        .Tools(strPeriod).ChangeAll ssChangeAllName, Str(nPPB)
                    End If
                End If
                
                If Len(strPeriod) > 0 Then
                    .Tools(strPeriod).State = ssChecked
                Else 'none:
                    .Tools("ID_Daily").State = ssChecked
                    .Tools("ID_Daily").State = ssUnchecked
                End If
                If Len(strType) > 0 Then
                    .Tools(strType).State = ssChecked
                Else 'none:
                    .Tools("ID_OHLCBars").State = ssChecked
                    .Tools("ID_OHLCBars").State = ssUnchecked
                End If
                
                .Tools("ID_ZoomOut").Enabled = m.bZoomed
                '.Tools("ID_ChartMove").Enabled = Not m.bZoomed
                '.Tools("ID_Eraser").Enabled = Not m.bZoomed
                
                .Tools("ID_Sectors").ChangeAll ssChangeAllName, Replace(strSector, "&", "&&")
                .Tools("ID_Subsectors").ChangeAll ssChangeAllName, Replace(strSub, "&", "&&")
                .Tools("ID_Components").ChangeAll ssChangeAllName, Replace(strCmp, "&", "&&")
                                
                If AutoScale Then
                    .Tools("ID_AutoScale").State = ssChecked
                    .Tools("ID_ChartMove").Picture = Picture16("kChartMoveHz")
                Else
                    .Tools("ID_AutoScale").State = ssUnchecked
                    .Tools("ID_ChartMove").Picture = Picture16("kChartMove")
                End If
                
                If IsInWhatIfMode Then
                    .Tools("ID_WhatIf").State = ssChecked
                Else
                    .Tools("ID_WhatIf").State = ssUnchecked
                End If
                
                If m.nShowTrades = 2 Then
                    .Tools("ID_ChartOrderbar").State = ssChecked
                Else
                    .Tools("ID_ChartOrderbar").State = ssUnchecked
                End If
                
                .Redraw = bSaveRedraw
            End With
        End If
                    
        SetCursor
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.SyncToolbar", eGDRaiseError_Raise

End Sub

Public Function TemplateApply(ByVal strTemplate$, Optional ByVal bFromCopyToOtherCharts As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strFile$, strPath$, i&, strMsg$, strChartName$
    Dim aTemplates As cGdArray
    Dim AnnotWhatIf As cAnnotation
        
    ' get name of template to apply
    strPath = App.Path & "\Charts\Templates\"
    strTemplate = Trim(strTemplate)
    If Left(strTemplate, 1) = "(" Or InStr(strTemplate, "<") > 0 Then
        strTemplate = ""
    ElseIf Len(strTemplate) = 1 And strTemplate >= "0" And strTemplate <= "9" Then
        Set aTemplates = GetAllowedList("T", True)
        strTemplate = Parse(aTemplates(Val(strTemplate)), vbTab, 1)
    ElseIf Left(strTemplate, 1) = "&" Then
        strTemplate = Trim(Mid(strTemplate, 4))
    End If

    'make sure chart is not zoomed
    If m.bZoomed = True Then
        If Len(strTemplate) > 0 Then
            UnzoomChart False
        Else
            UnzoomChart True
        End If
    End If
    
    Set m.PivotVwapInd = Nothing
    
    ' see if need to save existing template before applying new one
    If Len(strTemplate) > 0 And Len(m.strTemplateApplied) > 0 Then
        ' first save existing template and any new price pane annots
        ' (do this even if not saving template so price annots get saved)
        TemplateSave
        ' see if user wants to skip prompt
        If Not bFromCopyToOtherCharts Then
            If GetIniFileProperty("SaveTemplatePrompt", 1, "Charting", g.strIniFile) <> 0 Then
                ' see if existing template is dirty
                If TemplateIsDirty Then
                    ' ask if wish to save existing template
                    strMsg = "Save the current chart settings as template:|" & m.strTemplateApplied
                    strMsg = AskBox("i=? ; h=Save Chart Template? ; b=Save|+No|-Cancel ; " & strMsg)
                    If strMsg = "C" Then
                        Exit Function
                    ElseIf strMsg = "S" Then
                        TemplateSave
                        strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & Template & ".CHT"
                        On Error Resume Next
                        FileCopy strFile, strPath & m.strTemplateApplied & ".CHT"
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    End If
    
    If Len(strTemplate) > 0 Then
        'make copy of template to be applied
        If UCase(FileExt(strTemplate)) = "HID" Then
            strFile = strPath & strTemplate
        Else
            strFile = strPath & strTemplate & ".CHT"
        End If
        If FileExist(strFile) Then
            If m.strTemplate = "" Then
                'm.strTemplate = "TEMP"
                m.strTemplate = GetUnusedChartName
            End If
            'copy chart
            FileCopy strFile, g.ChartGlobals.strCPCRoot & "\Charts\" & m.strTemplate & ".CHT"
        Else
            strTemplate = ""
        End If
    End If

    If Len(strTemplate) > 0 Then
        'save what-if data
        Set AnnotWhatIf = m.Annots("WhatIf")
        'load new template and annotations
        strChartName = m.strChartName
        TemplateLoad
        m.strChartName = strChartName '(restore this chart's name)
        If Not bFromCopyToOtherCharts Then
            m.strTemplateApplied = strTemplate
            m.dUserScreenDate = 0
        End If
        If Not AnnotWhatIf Is Nothing Then
            If m.Annots("WhatIf") Is Nothing Then   'precautionary, should be true
                m.Annots.Add AnnotWhatIf, "WhatIf"
            End If
            Set AnnotWhatIf = Nothing
        End If
        TemplateApply = True
                
        If UCase(FileExt(strTemplate)) <> "HID" Then
            ' Now see if template just loaded is already considered "dirty"
            ' (e.g. due to newer properties since the template was created),
            ' if so then resave the template so this won't keep happening
            TemplateSave
            If TemplateIsDirty Then
                strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & Template & ".CHT"
                On Error Resume Next
                FileCopy strFile, strPath & m.strTemplateApplied & ".CHT"
                On Error GoTo 0
            End If
        End If
    
'JM(11-20-2009) Do this last otherwise causes intermitten crashes (5434)
        ResetSplitPane
        'geForceRecalc
        If m.nSymbolID <> 0 Or Len(m.strExtData & m.strSpreadSymbols) > 0 Then
            GenerateChart eRedo9_ReloadData
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.TemplateApply", eGDRaiseError_Raise

End Function

Public Function ChangeBarPeriod(ByVal vPeriod As Variant, _
    Optional ByVal bGenerateChart As Boolean = True, _
    Optional ByVal bGameModeInit As Boolean = False, _
    Optional ByVal bCalledByProfileEditor As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim s$, nNewPer&, eDisplayType As eIndicatorDisplayType
    Dim Ind As cIndicator
    
    If m.bProfileChart And Not bCalledByProfileEditor Then          '6847
        InfBox "Please use chart editor to change settings.", "i", , "Profile Chart"
        Exit Function
    End If
    
    If Not g.FractZen.Allowed Then
        m.strAutoBreakout = "" ' clear this if not allowed
    End If
    
    If m.Form.IsInGameMode Then
        If Not m.Form.GameMode Is Nothing Then
            If m.Form.GameMode.CustomOrders > 0 Then
                InfBox "You have open and/or filled orders in Instant Replay. Bars period cannot be changed."
                SyncToolbar
                Exit Function
            End If
        End If
    End If

    If VarType(vPeriod) <> vbString Then
        ' periodicity passed in directly
        nNewPer = CLng(vPeriod)
    ElseIf InStr("<C?", UCase(Left(Trim(vPeriod), 1))) > 0 Then
        ' custom: show bar period dialog
        Select Case UCase(Left(Trim(vPeriod), 2))
        Case "<P" 'Point&Figure
            If m.Bars.IsIntraday Then
                nNewPer = frmBarPeriod.ShowMe(ePRD_IntPF + 10, m.Bars, Me)          '6499
            Else
                nNewPer = frmBarPeriod.ShowMe(ePRD_EodPF + 100, m.Bars, Me)
            End If
        Case "<K" 'Kagi
            If m.Bars.IsIntraday Then
                nNewPer = frmBarPeriod.ShowMe(ePRD_IntKagi + 10, m.Bars, Me)
            Else
                nNewPer = frmBarPeriod.ShowMe(ePRD_EodKagi + 100, m.Bars, Me)
            End If
        Case "<R" 'Renko
            If m.Bars.IsIntraday Then
                nNewPer = frmBarPeriod.ShowMe(ePRD_IntRenko + 10, m.Bars, Me)
            Else
                nNewPer = frmBarPeriod.ShowMe(ePRD_EodRenko + 100, m.Bars, Me)
            End If
        Case Else
            nNewPer = frmBarPeriod.ShowMe(m.nPeriodicity, m.Bars, Me)
        End Select
        If nNewPer = 0 Then
            SyncToolbar
            Exit Function '(cancelled)
        End If
    ElseIf g.FractZen.Allowed And UCase(Left(Trim(vPeriod), 1)) = "F" Then
        If g.FractZen.UseMerged = 0 Then
            ' display auto-breakout settings
            '(TLB 5/10/2013: "Auto Breakout" renamed to "FractZen" per Needham)
            s = frmAutoBreakout.ShowMe(m.strAutoBreakout)
            If s = m.strAutoBreakout Then Exit Function
            m.strAutoBreakout = s
            If Len(s) = 0 Then Exit Function
        Else
            ' using merged FractZen bars
            m.strAutoBreakout = "0" ' to use default
            m.Bars.Prop(eBARS_FractZen) = 1
        End If
        nNewPer = ePRD_IntBreakout ' for now (will be reset later)
        RedoMode = eRedo9_ReloadData
    Else
        ' convert from string
        nNewPer = GetPeriodicity(vPeriod)
    End If
       
    If nNewPer <> m.nPeriodicity Then
        If g.FractZen.Allowed And Len(m.strAutoBreakout) > 0 And nNewPer <> ePRD_IntBreakout Then
            If g.FractZen.UseMerged = 0 Then
                s = frmAutoBreakout.ShowMe(m.strAutoBreakout)
                If s = m.strAutoBreakout Then Exit Function
                m.strAutoBreakout = s
                If Len(s) > 0 Then
                    nNewPer = ePRD_IntBreakout ' for now (will be reset later)
                    RedoMode = eRedo9_ReloadData
                End If
            Else
                ' no longer using merged FractZen bars
                m.strAutoBreakout = ""
                m.Bars.Prop(eBARS_FractZen) = 0
                RedoMode = eRedo9_ReloadData
            End If
        End If
        
        m.bAutoTrade = False
        If m.bZoomed = True Then UnzoomChart True
        
        If IsIntraday(nNewPer) And Not IsIntraday(m.nPeriodicity) Then
            'do this so chart won't stay blank if profile periodicity was something huge like M,Q or Yearly
            Set Ind = HasProfileVolume
            If Not Ind Is Nothing Then
                If InStr(Ind.ProfilePeriodicityStr, "Daily") = 0 Then Ind.ProfilePeriodicityStr = "Daily"
            End If
        End If
        
        ' fix display type (if switching to/from special type)
        Set Ind = m.Tree("PRICE")
        If Not Ind Is Nothing Then
            Select Case GetPeriodType(nNewPer)
            Case ePRD_EodPF, ePRD_IntPF
                Ind.DisplayType = eINDIC_PNF
            Case ePRD_EodKagi, ePRD_IntKagi
                Ind.DisplayType = eINDIC_Kagi
            Case ePRD_EodRenko, ePRD_IntRenko
                Ind.DisplayType = eINDIC_Renko
            Case Else
                If Ind.DisplayType < eINDIC_OHLC And Ind.DisplayType <> eINDIC_BollingerBar Then
                    Ind.DisplayType = eINDIC_OHLC
                End If
            End Select
        End If
        
        m.bTempAutoScale = True
        m.nPeriodicity = nNewPer
        m.Form.ClearAnnotFlags True, False
        
        If bGenerateChart Then
            GenerateChart eRedo9_ReloadData
            m.Form.SetChartTabs
        Else
            RedoMode = eRedo9_ReloadData
        End If

        If Not m.Form Is Nothing Then
            If m.Form.DetachStatus = eDetached And m.nShowToolbar = 1 Then
                SyncToolbar True            '5448
            ElseIf ActiveChart Is Nothing Then
                SyncToolbar
            ElseIf ActiveChart.DetachStatus = eDetached And ActiveChart.pbTbBack(0).Visible And m.Form Is g.ChartGlobals.frmActiveNonDetached Then
                'the active chart is detached with its own toolbar, force sync in case bar period on detached chart just got changed
                SyncToolbar True
            Else
                'there is no need to force a sync if chart is not detached or toolbar on detached chart is not shown
                SyncToolbar
            End If
        End If
        ChangeBarPeriod = True
        
        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
        
        frmMain.SetWindowLink m.Form, eLink_Period
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ChangeBarPeriod", eGDRaiseError_Raise

End Function

Public Sub SetFormCaption(Optional ByVal strCaption As String = "")
On Error GoTo ErrSection:

    Dim strBars$, strTitle$, strSubTitle$, strTemp$, strDesc$
    Dim nPoolRec&, dDate#, i&
    Dim aSpread As cGdArray
        
    If Not m.Form.GameMode Is Nothing Then
        strCaption = "Instant Replay:  "
        If m.Form.GameMode.GameDataTime > 0 Then
            dDate = m.Form.GameMode.GameDataTime
            If g.bShowInLocalTimeZone Then
                dDate = ConvertTimeZone(dDate, m.Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            End If
            strCaption = strCaption & DateFormat(dDate, MM_DD_YYYY, H_MM, AMPM_LOWER, True)
        End If
        If Not m.Form.GameMode.GameRandomSym Then
            strCaption = strCaption & " (" & Symbol & ")"
        End If
    End If
        
    If Len(strCaption) = 0 Then
        ' build default caption
        If m.Bars.HasCustomSessionTime And (m.nPartialLoadDate = 0) Then
            strCaption = Symbol & " (" & DateFormat(m.Bars.Prop(eBARS_StartTime) / 1440#, NO_DATE, H_MM, NO_AMPM) _
                    & "-" & DateFormat(m.Bars.Prop(eBARS_EndTime) / 1440#, NO_DATE, H_MM, NO_AMPM) & "):  "
        Else
            strCaption = Symbol & ":  "
        End If
        strTitle = strCaption
        If m.Bars.Size > 0 And m.Bars(eBARS_Close, 0) <> kNullData Then
            If m.Form.tmrProfileLoad.Enabled And m.nPartialLoadDate = 0 And Not m.BarsProfile Is Nothing Then
                dDate = m.BarsProfile(eBARS_DateTime, 0)
            Else
                ' use asterisk to indicate that more data exists before
                ' the first date and/or after the last date
                dDate = m.Bars(eBARS_DateTime, 0)
            End If
            
            strCaption = strCaption & DateFormat(dDate, MM_DD_YYYY)
            If m.bMoreDataBefore Then strCaption = strCaption & "*"
            dDate = m.Bars(eBARS_DateTime, LastGoodDataBar(False))
            strCaption = strCaption & " - " & DateFormat(dDate, MM_DD_YYYY)
            If m.bMoreDataAfter Then strCaption = strCaption & "*"
        End If
        strBars = GetPeriodStr(m.nPeriodicity) ' m.Bars.Prop(eBARS_PeriodicityStr)
        If Len(strBars) > 0 Then
            'If m.Bars.Prop(eBARS_Periodicity) > ePRD_Days And m.Bars.Prop(eBARS_PeriodsPerBar) > 1 Then
            If m.nPeriodicity > ePRD_Days And GetPeriodsPerBar(m.nPeriodicity) > 1 Then
                strCaption = strCaption & "  (" & strBars & " per bar)"
            ElseIf Len(m.strAutoBreakout) > 0 Then
                If g.bHideAutoBreakoutNumber Then
                    strCaption = strCaption & "  (FractZen)" '(Auto Breakout)"
                Else
                    strCaption = strCaption & "  (" & strBars & " AUTO)"
                End If
            Else
                strCaption = strCaption & "  (" & strBars & " bars)"
            End If
        End If
        If Len(m.strExtData) > 0 Then
            strDesc = Parse(m.strExtData, "|", 3)
        ElseIf Len(m.strSpreadSymbols) > 0 Then
            Set aSpread = New cGdArray
            aSpread.SplitFields m.strSpreadSymbols, ";"
            For i = 1 To aSpread.Size - 1
                If i = 1 Then
                    If Parse(aSpread(i), ",", 3) = "1" Then
                        strDesc = "(" & Parse(aSpread(i), ",", 2) & ")"
                    Else
                        strDesc = "(" & Parse(aSpread(i), ",", 3) & " " & Parse(aSpread(i), ",", 2) & ")"
                    End If
                Else
                    If Parse(aSpread(i), ",", 3) = "1" Then
                        strTemp = Parse(aSpread(i), ",", 1) & " (" & Parse(aSpread(i), ",", 2) & ")"
                    Else
                        strTemp = Parse(aSpread(i), ",", 1) & " (" & Parse(aSpread(i), ",", 3) & " " & Parse(aSpread(i), ",", 2) & ")"
                    End If
                    strDesc = strDesc & " " & strTemp
                End If
            Next
            strTemp = ""
        ElseIf m.nSymbolID <> 0 Then
            nPoolRec = g.SymbolPool.PoolRecForSymbolID(m.nSymbolID)
            strDesc = g.SymbolPool.Desc(nPoolRec)
        End If
        
        If m.nPartialLoadDate > 0 Or m.Form.tmrProfileLoad.Enabled Then
            If m.nPartialLoadDate > 0 Then
                strCaption = " " & Replace(strCaption, ":", " --> Loading:")
            Else
                strCaption = " " & Replace(strCaption, ":", " --> Loading Profile Bars:")
            End If
        Else
            strCaption = " " & strCaption & "   " & strDesc
            If Len(m.strTemplateApplied) > 0 Then
                strCaption = strCaption & "  --  Template: " & m.strTemplateApplied
            End If
        End If
    
        ' also set title/subtitle
        If m.iViewMode <> 0 And m.iMiniChart = 0 Then
            strTitle = Symbol & ":  " & strDesc & "  (" & strBars & " bars)"
            'strSubTitle = "Graphed by the 'Trade Navigator'  --  www.GenesisFT.com"
            m.strSubTitle = GetProvidedProperty("Website", , True)
        End If
    End If

    ' put caption in vseCaption and as Form caption (unless maximized child)
    strTemp = strCaption
    If m.Form.vseCaption.Top = 0 Then strTemp = ""
    If m.Form.Caption <> strTemp Then
        m.Form.Caption = strTemp
        ' if empty, refresh now (to reduce flicker of frmMain caption
        ' when chart is maximizing, etc.)
        If Len(strTemp) = 0 Then m.Form.Refresh
    End If
    
    strCaption = "    " & Replace(strCaption, "&", "&&")
    If m.Form.vseCaption.Caption <> strCaption Then
        m.Form.vseCaption.Caption = strCaption
        m.Form.vseCaption.Refresh
    End If

    'MJM [begin] - added for grapheng.dll
    m.strTitle = strTitle
    m.strSubTitle = strSubTitle
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.SetFormCaption", eGDRaiseError_Raise

End Sub

Public Function NetProfit() As Double
    NetProfit = m.Transactions.TblNum(eGDTransactionField_UnfilteredEquity)
End Function

Public Function OpenEquity(Optional ByVal dCurrentPrice# = 0) As Double
On Error GoTo ErrSection:

    Static BaseForexBars As cGdBars
    Static nPrevLDD&, strPrevSymbol$
    
    Dim strSymbol$, strBaseForex$
    Dim dDate1#, dDate2#
    Dim dPrice1#, dPrice2#, iQty&
    
    ' calculate from last trade if it was an entry
    iQty = m.Transactions.TblNum(eGDTransactionField_Position)
    
    If iQty <> 0 Then
        dPrice1 = m.Transactions.TblNum(eGDTransactionField_AvgEntry)
        If m.Form.GameMode Is Nothing Or dCurrentPrice = 0 Then
            dPrice2 = m.Bars(eBARS_Close, LastGoodDataBar(False))
        Else
            dPrice2 = dCurrentPrice             '6476
        End If
        
        dDate1 = m.Transactions.TblNum(eGDTransactionField_Date)
        dDate2 = m.Bars(eBARS_DateTime, LastGoodDataBar(False))
        
        strSymbol = m.Bars.Prop(eBARS_Symbol)
        strBaseForex = g.Profit.BaseForex(strSymbol)
        If Len(strBaseForex) = 0 Then
            Set BaseForexBars = Nothing
        ElseIf BaseForexBars Is Nothing Or nPrevLDD <> LastDailyDownload Or strPrevSymbol <> strSymbol Then
            strPrevSymbol = strSymbol
            nPrevLDD = LastDailyDownload
            Set BaseForexBars = New cGdBars
            DM_GetBars BaseForexBars, strBaseForex
        End If
        OpenEquity = g.Profit.Profit(m.Bars.Prop(eBARS_Symbol), dPrice2 - dPrice1, _
                iQty, MaxDouble(dDate1, dDate2), Bars, BaseForexBars, TradeAccountID)          '5620
    End If
    

'JM(02-24-2010): original code;leave awhile then remove if all okay
'    Dim strText$, dDiff#, strPos$, nQty&, iRec&, nPos&, dAvgEntry#
'
'    If nQty <> 0 And m.Bars.Prop(eBARS_TickMove) > 0 Then
'        If dCurrentPrice = 0 Then
'            dCurrentPrice = m.Bars(eBARS_Close, LastGoodDataBar(False))
'        End If
'        dDiff = dCurrentPrice - m.Transactions.TblNum(eGDTransactionField_AvgEntry)
'        OpenEquity = nQty * dDiff * m.Bars.Prop(eBARS_TickValue) / m.Bars.Prop(eBARS_TickMove)
'        If Not m.System Is Nothing And m.nSystemID <> 0 And m.nShowTrades = 1 Then
'            OpenEquity = OpenEquity - m.System.Expenses * Abs(nQty)
'        End If
'    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cChart.OpenEquity"

End Function

Public Function Position() As Long
    Position = m.Transactions.TblNum(eGDTransactionField_Position)
End Function

Private Sub ReadTrades(Optional ByVal bUseStrategy As Boolean = True)
On Error GoTo ErrSection:

    Dim strFile As String               ' Filename of file to read in
    Dim strTemp As String               ' Temporary string
    Dim strText As String               ' Temporary string
    Dim strCurrentPosition As String    ' Current position
    Dim strAction As String             ' Action of order
    Dim bNoConsolidated As Boolean      ' No consolidated orders?
    Dim bSeeOT As Boolean               ' Does the user need to see the order tracker?
    Dim bMarket As Boolean              ' Market order?
    Dim bBadOrders As Boolean           ' Do we have bad orders?
    Dim dEntryPrice As Double           ' Entry price
    Dim dEntryDate As Double            ' Entry date
    Dim dExitPrice As Double            ' Exit price
    Dim dNboPrice As Double             ' Next bar order price
    Dim dCheck As Double                ' Validation variable
    Dim lNextBar As Long                ' Next bar
    Dim lIndex As Long                  ' Index into a for loop
    Dim lQuantity As Long               ' Quantity
    Dim astrFile As cGdArray            ' Array read in from a file
    Dim astrFields As cGdArray          ' Array of split out fields

    Set astrFile = New cGdArray
    astrFile.Create eGDARRAY_Strings
    Set astrFields = New cGdArray
    astrFields.Create eGDARRAY_Strings

    m.Transactions.NumRecords = 0
    m.aOrders.Clear

    If (m.nSystemID = 0 Or m.System Is Nothing) And bUseStrategy Then Exit Sub
    
    If m.nShowTrades = 2 Then
        m.dNextBarDate = m.Bars(eBARS_DateTime, Me.LastGoodDataBar(False) + 1)
    End If
    
    ' get "next bar": bar currently being traded
    lNextBar = m.Bars.FindDateTime(m.dNextBarDate, True)
    dNboPrice = m.Bars(eBARS_Open, lNextBar)
    
' 1 line if can't do orders (N/A)
' 4 lines if can do orders but none exist

'20  Sample: Turtle strategy 1 minute    37971.5451388889    37972.6333333333    520 45  SP-067  0.1 25  0.5 20000   F
'37972.633333333 37972.632638889 1072.7  1072.7  1071.5  1071.5
'S   1   37972.632638889 1072.2  Sell when a lower low than prior days
'1071.6  1071.6
'0   1071.6
'  1 XS  1179    L   1070.7              1
'  1 EL  728 S   1072.7              1
'    2   XL  726 S   1072.2              1
'    2   XL  1178    L   1074.2              1
'  1 XS  1177    S   1072.5              1
    
    If bUseStrategy Then
        ' read orders for next bar
        strFile = m.System.NextBarFile(eGDNextBarMode_UseHwnd)
        If FileLength(strFile) < 10 Then
            ' if no consolidated, still need to check the RB file for the current position
            bNoConsolidated = True
            strTemp = FileBase(strFile)
            strFile = AddSlash(FilePath(strFile)) & "RB" & Mid(strTemp, 3) & ".TXT"
        End If
        dEntryPrice = 0
        
        astrFile.FromFile strFile
        If astrFile.Size > 0 Then
            bSeeOT = True
            m.aOrders.Add "(no orders)"
            For lIndex = 1 To astrFile.Size - 1
                strText = Trim(astrFile(lIndex))
                If Len(strText) > 0 Then
                    astrFields.SplitFields strText, vbTab
                    Select Case lIndex
                    Case 0
                    Case 1
                        ' check for correct next bar date
                        dEntryDate = RoundToSecond(Val(astrFields(0)))
                        If dEntryDate <> m.dNextBarDate Then Exit For
                    Case 2
                        ' save position to compare with trades report
                        strCurrentPosition = astrFields(0)
                        If bNoConsolidated Then
                            bSeeOT = True
                            Exit For
                        End If
                    Case 3
                        ' get range tested
                        dEntryPrice = Val(astrFields(0))
                        dExitPrice = 0
                        If astrFields.Size - 1 >= 1 Then
                            dExitPrice = Val(astrFields(1))
                        End If
                        If dEntryPrice <= 0 Then
                            Exit For '(something's wrong!)
                        ElseIf dExitPrice = 0 Then
                            dExitPrice = dEntryPrice
                        End If
                        bSeeOT = False
                    
                    Case Else ' ORDERS ...
                        Select Case Val(astrFields(0))
                        Case 0 ' NBO range:
                            ' if NBO price range doesn't exactly match then we can't do orders
                            dCheck = Val(astrFields(1))
                            If dCheck <> dEntryPrice Then
                                bSeeOT = True
                                Exit For
                            End If
                            dCheck = 0
                            If astrFields.Size - 1 >= 2 Then
                                dCheck = Val(astrFields(2))
                            End If
                            If dCheck = 0 Then
                                dCheck = dEntryPrice
                            End If
                            If dCheck <> dExitPrice Then
                                bSeeOT = True
                                Exit For
                            End If
                        
                        Case 1 ' Primary orders:
                            If dNboPrice = kNullData Then
                                ' can't do orders if need NBO but not known
                                If astrFields.Size - 1 >= 5 Then
                                    If astrFields(5) = "ONB" Then
                                        bSeeOT = True
                                        Exit For
                                    End If
                                End If
                            ElseIf astrFields.Size - 1 >= 4 Then
                                ' see if need to convert a stop/limit to a market order
                                ' (compare to open price of next bar)
                                bMarket = False
                                strAction = astrFields(1)
                                dEntryPrice = Val(astrFields(4))
                                Select Case astrFields(3)
                                Case "L"
                                    If strAction = "EL" Or strAction = "XS" Then
                                        If dEntryPrice >= dNboPrice Then bMarket = True
                                    Else
                                        If dEntryPrice <= dNboPrice Then bMarket = True
                                    End If
                                Case "S"
                                    If strAction = "EL" Or strAction = "XS" Then
                                        If dEntryPrice <= dNboPrice Then bMarket = True
                                    Else
                                        If dEntryPrice >= dNboPrice Then bMarket = True
                                    End If
                                End Select
                                If bMarket Then
                                    ' convert to a market order (since would fill at the open)
                                    strText = "1" & vbTab & strAction & vbTab & astrFields(2) & vbTab & "M"
                                End If
                            End If
                            m.aOrders.Add strText
                        
                        Case 2 ' Conditional orders: for now just append onto end of the primary order
                            If dNboPrice = kNullData Then
                                ' can't do orders if need NBO but not known
                                If astrFields.Size - 1 >= 5 Then
                                    If astrFields(5) = "ONB" Then
                                        bSeeOT = True
                                        Exit For
                                    End If
                                End If
                            End If
                            m.aOrders(m.aOrders.Size - 1) = m.aOrders(m.aOrders.Size - 1) & vbCrLf & strText
                        End Select
                    End Select
                End If
            Next
            If bSeeOT Then
                ' orders couldn't be fully determined, so must see Order Tracker
                m.aOrders.Clear
                m.aOrders.Add "(click here to see" & vbCrLf & "the Orders report)"
            End If
        End If
    End If

'169 a sp tester Daily   30062   37970   5466    45  SP-067  0.1 25  0.5 20000   F
'S   37964.000000    1068.1  DOWN CLOSE THEN NAKED UP AND OOPS!-1    37965.000000    1059.1  exit short 1/2 bailout  2205.00 2655.00 -45.00  1   4768    4693    2   7   237 0
'S   37970.000000    1081.9  sell as Bonds open <    37970.000000    1075.8      1480.00 2105.00 -45.00  0   5544    (null)  1   11  241 0
    
    ' read trades
    If bUseStrategy Then
        strFile = m.System.TradesFile(True)
    ElseIf m.bReloadTrackerTrades Then
        strFile = TradesFileForAccount(True)
        m.bReloadTrackerTrades = False
    Else
        strFile = TradesFileForAccount(False)
    End If
    
    astrFile.FromFile strFile
    m.Transactions.FromTradesArray astrFile
    m.Transactions.CalculateStats
    m.Transactions.CalculateEquityFilter

    ' make sure positions match between reports
    bBadOrders = False
    lQuantity = m.Transactions.TblNum(eGDTransactionField_Position)
    If lQuantity > 0 Then
        If strCurrentPosition <> "L" Then
            bBadOrders = True
        End If
    ElseIf lQuantity < 0 Then
        If strCurrentPosition <> "S" Then
            bBadOrders = True
        End If
    ElseIf strCurrentPosition = "L" Then
        bBadOrders = True
    ElseIf strCurrentPosition = "S" Then
        bBadOrders = True
    End If
    If bBadOrders And bUseStrategy Then
        m.aOrders.Clear
        m.aOrders.Add "ERROR with orders" & vbCrLf & "(click here for info)" '"(position mismatch)"
    End If
    
    If Left(Symbol, 1) = "$" And Not IsForex(Symbol) Then
        m.aOrders.Clear
        m.aOrders.Add "(orders are not" & vbCrLf & "given for indexes)"
    End If
    
    AddSpeed "Read Trades"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.ReadTrades"
    
End Sub

#If 0 Then
Private Sub ReadTradesOld(Optional ByVal bUseStrategy As Boolean = True)
On Error GoTo ErrSection:

    Dim i&, iRec&, iEntryRec&, iLinkRec&, iNumRecs&, strText$
    Dim nQty&, eDate#, xDate#, ePrice#, xPrice#, dChk#
    Dim nNextBar&, nboPrice#, dProfit#
    Dim bMarket As Boolean, bSeeOT As Boolean, bBadOrders As Boolean, bNoConsolidated As Boolean
    Dim strFile$, strRule$, strAction$, strCurrentPosition$, strTemp$
    Dim aStrings() As String, aFlds() As String
       
    m.tblTrades.NumRecords = 0
    m.aOrders.Clear
    If (m.nSystemID = 0 Or m.System Is Nothing) And bUseStrategy Then Exit Sub
    
    If m.nShowTrades = 2 Then
        m.dNextBarDate = m.Bars(eBARS_DateTime, Me.LastGoodDataBar(False) + 1)
    End If
    
    ' get "next bar": bar currently being traded
    nNextBar = m.Bars.FindDateTime(m.dNextBarDate, True)
    nboPrice = m.Bars(eBARS_Open, nNextBar)
    
' 1 line if can't do orders (N/A)
' 4 lines if can do orders but none exist

'20  Sample: Turtle strategy 1 minute    37971.5451388889    37972.6333333333    520 45  SP-067  0.1 25  0.5 20000   F
'37972.633333333 37972.632638889 1072.7  1072.7  1071.5  1071.5
'S   1   37972.632638889 1072.2  Sell when a lower low than prior days
'1071.6  1071.6
'0   1071.6
'  1 XS  1179    L   1070.7              1
'  1 EL  728 S   1072.7              1
'    2   XL  726 S   1072.2              1
'    2   XL  1178    L   1074.2              1
'  1 XS  1177    S   1072.5              1
    
    If bUseStrategy Then
        ' read orders for next bar
        strFile = m.System.NextBarFile(eGDNextBarMode_UseHwnd)
        If FileLength(strFile) < 10 Then
            ' if no consolidated, still need to check the RB file for the current position
            bNoConsolidated = True
            strTemp = FileBase(strFile)
            strFile = AddSlash(FilePath(strFile)) & "RB" & Mid(strTemp, 3) & ".TXT"
        End If
        ePrice = 0
        FileToArray strFile, aStrings
        If UBound(aStrings) > 0 Then
            bSeeOT = True
            m.aOrders.Add "(no orders)"
            For i = 1 To UBound(aStrings)
                strText = Trim(aStrings(i))
                If Len(strText) > 0 Then
                    aFlds = Split(strText, vbTab)
                    Select Case i
                    Case 1
                    Case 2
                        ' check for correct next bar date
                        eDate = RoundToSecond(Val(aFlds(0)))
                        If eDate <> m.dNextBarDate Then Exit For
                    Case 3
                        ' save position to compare with trades report
                        strCurrentPosition = aFlds(0)
                        If bNoConsolidated Then
                            bSeeOT = True
                            Exit For
                        End If
                    Case 4
                        ' get range tested
                        ePrice = Val(aFlds(0))
                        xPrice = 0
                        If UBound(aFlds) >= 1 Then
                            xPrice = Val(aFlds(1))
                        End If
                        If ePrice <= 0 Then
                            Exit For '(something's wrong!)
                        ElseIf xPrice = 0 Then
                            xPrice = ePrice
                        End If
                        bSeeOT = False
                    
                    Case Else ' ORDERS ...
                        Select Case Val(aFlds(0))
                        Case 0 ' NBO range:
                            ' if NBO price range doesn't exactly match then we can't do orders
                            dChk = Val(aFlds(1))
                            If dChk <> ePrice Then
                                bSeeOT = True
                                Exit For
                            End If
                            dChk = 0
                            If UBound(aFlds) >= 2 Then
                                dChk = Val(aFlds(2))
                            End If
                            If dChk = 0 Then
                                dChk = ePrice
                            End If
                            If dChk <> xPrice Then
                                bSeeOT = True
                                Exit For
                            End If
                        
                        Case 1 ' Primary orders:
                            If nboPrice = kNullData Then
                                ' can't do orders if need NBO but not known
                                If UBound(aFlds) >= 5 Then
                                    If aFlds(5) = "ONB" Then
                                        bSeeOT = True
                                        Exit For
                                    End If
                                End If
                            ElseIf UBound(aFlds) >= 4 Then
                                ' see if need to convert a stop/limit to a market order
                                ' (compare to open price of next bar)
                                bMarket = False
                                strAction = aFlds(1)
                                ePrice = Val(aFlds(4))
                                Select Case aFlds(3)
                                Case "L"
                                    If strAction = "EL" Or strAction = "XS" Then
                                        If ePrice >= nboPrice Then bMarket = True
                                    Else
                                        If ePrice <= nboPrice Then bMarket = True
                                    End If
                                Case "S"
                                    If strAction = "EL" Or strAction = "XS" Then
                                        If ePrice <= nboPrice Then bMarket = True
                                    Else
                                        If ePrice >= nboPrice Then bMarket = True
                                    End If
                                End Select
                                If bMarket Then
                                    ' convert to a market order (since would fill at the open)
                                    strText = "1" & vbTab & strAction & vbTab & aFlds(2) & vbTab & "M"
                                End If
                            End If
                            m.aOrders.Add strText
                        
                        Case 2 ' Conditional orders: for now just append onto end of the primary order
                            If nboPrice = kNullData Then
                                ' can't do orders if need NBO but not known
                                If UBound(aFlds) >= 5 Then
                                    If aFlds(5) = "ONB" Then
                                        bSeeOT = True
                                        Exit For
                                    End If
                                End If
                            End If
                            m.aOrders(m.aOrders.Size - 1) = m.aOrders(m.aOrders.Size - 1) & vbCrLf & strText
                        End Select
                    End Select
                End If
            Next
            If bSeeOT Then
                ' orders couldn't be fully determined, so must see Order Tracker
                m.aOrders.Clear
                m.aOrders.Add "(click here to see" & vbCrLf & "the Orders report)"
            End If
        End If
    End If

'169 a sp tester Daily   30062   37970   5466    45  SP-067  0.1 25  0.5 20000   F
'S   37964.000000    1068.1  DOWN CLOSE THEN NAKED UP AND OOPS!-1    37965.000000    1059.1  exit short 1/2 bailout  2205.00 2655.00 -45.00  1   4768    4693    2   7   237 0
'S   37970.000000    1081.9  sell as Bonds open <    37970.000000    1075.8      1480.00 2105.00 -45.00  0   5544    (null)  1   11  241 0
    
    ' read trades
    If bUseStrategy Then
        strFile = m.System.TradesFile(True)
    ElseIf m.bReloadTrackerTrades Then
        strFile = TradesFileForAccount(True)
        m.bReloadTrackerTrades = False
    Else
        strFile = TradesFileForAccount(False)
    End If
    
    FileToArray strFile, aStrings
    If UBound(aStrings) > 0 Then
        ' size table big enough now for efficiency (will resize later)
        m.tblTrades.NumRecords = UBound(aStrings) * 2
        iNumRecs = 0
        For i = 1 To UBound(aStrings)
            strText = Trim(aStrings(i))
            If Len(strText) > 0 Then
                aFlds = Split(strText, vbTab)
                Select Case aFlds(0)
                Case "L"
                    nQty = 1
                Case "S"
                    nQty = -1
                Case Else
                    nQty = 0
                End Select
'E 3/14  X 3/17
'E 3/15  X 3/16

'74 E 3/14  75
'75 X 3/17  74
'76

'74 E 3/14  76m
'75 E 3/15  -1
'76 X 3/17  74

'74 E 3/14  77m
'75 E 3/15  76
'76 X 3/16  75
'77 X 3/17  74
                If nQty <> 0 And UBound(aFlds) >= 7 Then
                    ' ENTRY
                    eDate = RoundToSecond(Val(aFlds(1)))
                    ePrice = Val(aFlds(2))
                    strRule = Trim(aFlds(3))
                    ' verify valid date and rule name exists
                    If eDate >= 10000 And eDate < m.dNextBarDate And Len(strRule) > 0 Then
                        ' insert before any later trades (if not at end, must insert a blank record)
                        For iRec = iNumRecs To 0 Step -1
                            If eDate >= TradeNum(eTRADE_Date, iRec - 1) Then Exit For
                            ' if linked to a record, increment the link for the linked record
                            iLinkRec = TradeNum(eTRADE_Link, iRec - 1)
                            If iLinkRec >= 0 Then
                                TradeNum(eTRADE_Link, iLinkRec) = TradeNum(eTRADE_Link, iLinkRec) + 1
                            End If
                        Next
                        If iRec < iNumRecs Then
                            m.tblTrades.AddRecord "", iRec
                        End If
                        iNumRecs = iNumRecs + 1
                        ' set fields in table
                        If nQty > 0 Then
                            strAction = "EL"
                        Else
                            strAction = "ES"
                        End If
                        TradeNum(eTRADE_Date, iRec) = eDate
                        TradeNum(eTRADE_Qty, iRec) = nQty
                        TradeStr(eTRADE_Action, iRec) = strAction
                        TradeNum(eTRADE_Price, iRec) = ePrice
                        TradeStr(eTRADE_Rule, iRec) = strRule
                        TradeNum(eTRADE_Link, iRec) = -1 '(for now)
                        iEntryRec = iRec
                        
                        'EXIT
                        xDate = RoundToSecond(Val(aFlds(4)))
                        ' for sorting purposes: if exit same bar as entered, bump up 1 second, else bump down 1 second
                        ' (because if exit date same as entry date then this exit happened after the entries on this bar,
                        ' otherwise this exit occurred before any other entries on this bar)
                        If xDate = eDate Then
                            xDate = xDate + 0.000001
                        Else
                            xDate = xDate - 0.000001
                        End If
                        xPrice = Val(aFlds(5))
                        strRule = Trim(aFlds(6))
                        ' verify valid date and rule name exists
                        If xDate >= 10000 And xDate < m.dNextBarDate And Len(strRule) > 0 Then
                            ' insert before any later trades (if not at end, must insert a blank record)
                            For iRec = iNumRecs To 0 Step -1
                                If xDate >= TradeNum(eTRADE_Date, iRec - 1) Then Exit For
                                ' if linked to a record, increment the link for the linked record
                                iLinkRec = TradeNum(eTRADE_Link, iRec - 1)
                                If iLinkRec >= 0 Then
                                    TradeNum(eTRADE_Link, iLinkRec) = TradeNum(eTRADE_Link, iLinkRec) + 1
                                End If
                            Next
                            If iRec < iNumRecs Then
                                m.tblTrades.AddRecord "", iRec
                            End If
                            iNumRecs = iNumRecs + 1
                            ' set fields in table
                            dProfit = Val(aFlds(7))
                            If nQty > 0 Then
                                strAction = "XL"
                            Else
                                strAction = "XS"
                            End If
                            TradeNum(eTRADE_Date, iRec) = xDate
                            TradeNum(eTRADE_Qty, iRec) = -nQty
                            TradeStr(eTRADE_Action, iRec) = strAction
                            TradeNum(eTRADE_Price, iRec) = xPrice
                            TradeStr(eTRADE_Rule, iRec) = strRule
                            TradeNum(eTRADE_Profit, iRec) = dProfit
                            ' link entry and exit
                            TradeNum(eTRADE_Link, iRec) = iEntryRec
                            TradeNum(eTRADE_Link, iEntryRec) = iRec
                        End If
                    End If
                End If
            End If
        Next
        ' now set to the correct size
        m.tblTrades.NumRecords = iNumRecs
    End If
    
    ' calc net profit, positions, etc.
    CalcTradeStats
    
    ' make sure positions match between reports
    bBadOrders = False
    nQty = Position
    If nQty > 0 Then
        If strCurrentPosition <> "L" Then
            bBadOrders = True
        End If
    ElseIf nQty < 0 Then
        If strCurrentPosition <> "S" Then
            bBadOrders = True
        End If
    ElseIf strCurrentPosition = "L" Then
        bBadOrders = True
    ElseIf strCurrentPosition = "S" Then
        bBadOrders = True
    End If
    If bBadOrders And bUseStrategy Then
        m.aOrders.Clear
        m.aOrders.Add "ERROR with orders" & vbCrLf & "(click here for info)" '"(position mismatch)"
    End If
    
    If Left(Symbol, 1) = "$" And Not IsForex(Symbol) Then
        m.aOrders.Clear
        m.aOrders.Add "(orders are not" & vbCrLf & "given for indexes)"
    End If
    
    AddSpeed "Read Trades"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.ReadTrades", eGDRaiseError_Raise

End Sub
#End If

Private Function FillOrders() As Boolean
On Error GoTo ErrSection:

    Dim i&, nNextBar&, iFillOrder&, nQty&
    Dim dPrice#, dBasePrice#, dProfit#
    Dim dBestPriceAbove#, dBestPriceBelow#, iBestOrderAbove&, iBestOrderBelow&
    Dim strText$, strRule$, strPos$, strCondOrders$
    Dim aFlds As New cGdArray
    Dim PtOrder As New cPtOrder
    Dim bForceLimitThrough As Boolean
    Dim dLimitAdjust#
    
    ' get "next bar": bar currently being traded
    nNextBar = m.Bars.FindDateTime(m.dNextBarDate, True)
    If m.Bars(eBARS_DateTime, nNextBar) <> m.dNextBarDate Then Exit Function
    If m.Bars(eBARS_Open, nNextBar) = kNullData Then Exit Function
    
    If Not m.System Is Nothing Then
        bForceLimitThrough = m.System.ForceLimitThrough
    End If
    
    ' loop until no orders are filled
    Do While m.aOrders.Size > 1
        ' init the variables
        iFillOrder = -1
        iBestOrderAbove = -1
        iBestOrderBelow = -1
        dBestPriceAbove = 999999
        dBestPriceBelow = -999999
        
        ' first get the "base price" (to check for orders above or below):
        ' use the "next bar open", unless there was an entry on the next bar
        dBasePrice = m.Bars(eBARS_Open, nNextBar)
        For i = m.Transactions.NumRecords - 1 To 0 Step -1
            If m.Transactions.TblNum(eGDTransactionField_Date, i) <> m.dNextBarDate Then Exit For
            If InStr(m.Transactions.TblStr(eGDTransactionField_Action, i), "E") > 0 Then
                dBasePrice = m.Transactions.TblNum(eGDTransactionField_Price, i)
            End If
        Next
    
    '  1 XS  1179    L   1070.7              1
    '  1 EL  728 S   1072.7              1
    '    2   XL  726 S   1072.2              1
    '    2   XL  1178    L   1074.2              1
    '  1 XS  1177    S   1072.5              1
        For i = 1 To m.aOrders.Size - 1
            strText = Parse(m.aOrders(i), vbCrLf, 1)
            aFlds.SplitFields strText, vbTab
            Select Case aFlds(3)
            Case "M"
                iFillOrder = i
                Exit For
            
            Case "S", "L"
                dPrice = Val(aFlds(4))
                
                ' make adjustment if want to force trading through limit price to fill
                If aFlds(3) = "L" And bForceLimitThrough Then
                    dLimitAdjust = 0.000000001
                Else
                    dLimitAdjust = 0
                End If
                
                ' if order was placed on this bar in Game Mode, see if order has filled SINCE it was placed
                If m.Form.IsInGameMode And Val(aFlds(9)) <= m.dNextBarDate Then
                    If dPrice > Val(aFlds(12)) Then
                        ' if order was placed above market (at the time order was placed), and the
                        ' market is now above the order price, or the high is above the order price and
                        ' the high has changed (since order was placed), then order should have filled
                        If m.Bars(eBARS_Close, nNextBar) - dLimitAdjust >= dPrice Or (m.Bars(eBARS_High, nNextBar) - dLimitAdjust >= dPrice _
                                    And Val(aFlds(10)) <> m.Bars(eBARS_High, nNextBar)) Then
                            If dPrice < dBestPriceAbove Then
                                iBestOrderAbove = i
                                
                                'if the order price < low of nextbar then fill at open (gap up/down aardvark issue 6731)
                                If dPrice < m.Bars(eBARS_Low, nNextBar) Or dPrice > m.Bars(eBARS_High, nNextBar) Then
                                    dBestPriceAbove = m.Bars(eBARS_Open, nNextBar)
                                Else
                                    dBestPriceAbove = dPrice
                                End If
                            End If
                        End If
                    ElseIf dPrice < Val(aFlds(12)) Then
                        ' if order was placed below market (at the time order was placed), and the
                        ' market is now below the order price, or the low is below the order price and
                        ' the low has changed (since order was placed), then order should have filled
                        If m.Bars(eBARS_Close, nNextBar) + dLimitAdjust <= dPrice Or (m.Bars(eBARS_Low, nNextBar) + dLimitAdjust <= dPrice _
                                    And Val(aFlds(11)) <> m.Bars(eBARS_Low, nNextBar)) Then
                            If dPrice > dBestPriceBelow Then
                                iBestOrderBelow = i
                                
                                'if the order price > high of nextbar then fill at open (gap up/down aardvark issue 6731)
                                If dPrice < m.Bars(eBARS_Low, nNextBar) Or dPrice > m.Bars(eBARS_High, nNextBar) Then
                                    dBestPriceBelow = m.Bars(eBARS_Open, nNextBar)
                                Else
                                    dBestPriceBelow = dPrice
                                End If
                            End If
                        End If
                    End If
                ' can check the bar's range in order to fill all primary orders (type = 1),
                ' or conditional orders (type = 2) if the entry happened at the bar's open
                ElseIf Val(aFlds(0)) = 1 Or dBasePrice = m.Bars(eBARS_Open, nNextBar) Then
                    ' see if order price is now within the bar's range (i.e. is fillable)
                    If dPrice <= m.Bars(eBARS_High, nNextBar) - dLimitAdjust And dPrice >= m.Bars(eBARS_Low, nNextBar) + dLimitAdjust Then
                        ' see if is best fillable order above or below the base price
                        If dPrice > dBasePrice Then
                            If dPrice < dBestPriceAbove Then
                                iBestOrderAbove = i
                                dBestPriceAbove = dPrice
                            End If
                        Else
                            If dPrice > dBestPriceBelow Then
                                iBestOrderBelow = i
                                dBestPriceBelow = dPrice
                            End If
                        End If
                    End If
                ' other conditional orders (type = 2) can only be filled for sure if
                ' the order price is between the entry price and the current price
                ' (else don't know if the order was hit before or after the entry)
                ElseIf dPrice > dBasePrice And dPrice <= m.Bars(eBARS_Close, nNextBar) - dLimitAdjust Then
                    If dPrice < dBestPriceAbove Then
                        iBestOrderAbove = i
                        dBestPriceAbove = dPrice
                    End If
                ElseIf dPrice < dBasePrice And dPrice >= m.Bars(eBARS_Close, nNextBar) + dLimitAdjust Then
                    If dPrice > dBestPriceBelow Then
                        iBestOrderBelow = i
                        dBestPriceBelow = dPrice
                    End If
                End If
            End Select
        Next
        If iFillOrder < 0 Then
            If iBestOrderAbove >= 0 And iBestOrderBelow < 0 Then
                iFillOrder = iBestOrderAbove
            ElseIf iBestOrderAbove < 0 And iBestOrderBelow >= 0 Then
                iFillOrder = iBestOrderBelow
            End If
        End If
        
        ' exit loop when no more orders to fill
        If iFillOrder < 0 Then Exit Do
        
        ' convert this order to a trade
        strText = Parse(m.aOrders(iFillOrder), vbCrLf, 1)
        aFlds.SplitFields strText, vbTab
        Select Case aFlds(3)
        Case "M"
            dPrice = m.Bars(eBARS_Open, nNextBar)
        Case Else
            If m.Form.IsInGameMode Then
                If dBestPriceAbove <> 999999 Then
                    dPrice = dBestPriceAbove        '6731
                ElseIf dBestPriceBelow <> -999999 Then
                    dPrice = dBestPriceBelow
                Else
                    dPrice = Val(aFlds(4))
                End If
            Else
                dPrice = Val(aFlds(4))
            End If
        End Select
        strPos = aFlds(1)
        nQty = Abs(Val(aFlds(8)))
        If nQty = 0 Then
            nQty = 1
        End If
        If strPos = "ES" Or strPos = "XL" Then
            nQty = -nQty
        End If
        On Error Resume Next
        strRule = Trim(aFlds(2))
        strRule = m.System.Rules.Item(strRule).Name
        On Error GoTo ErrSection:

dProfit = OpenEquity(dPrice)
        
        m.Transactions.NumRecords = m.Transactions.NumRecords + 1
        If Not m.Form.GameMode Is Nothing Then
            If m.Form.GameMode.GameStrategyID = 0 Then
                'fix for aardvark 3008 and aardvar 3027 (intraday data issue)
                'setting the trade date to next bar date causes the trade
                'to be remove in the LoadScreenTrades routine if the order
                'was filled on the bar but not exactly at the bar date
                'e.g. next bar date might be 10:15, but order was filled at 10:13
                m.Transactions.TblNum(eGDTransactionField_Date) = m.Form.GameMode.GameDataTime
            Else
                m.Transactions.TblNum(eGDTransactionField_Date) = m.dNextBarDate
            End If
        Else
            m.Transactions.TblNum(eGDTransactionField_Date) = m.dNextBarDate
        End If
        m.Transactions.TblNum(eGDTransactionField_Quantity) = nQty
        m.Transactions.TblStr(eGDTransactionField_Action) = strPos
        m.Transactions.TblNum(eGDTransactionField_Price) = dPrice
        m.Transactions.TblStr(eGDTransactionField_Rule) = strRule
        m.Transactions.TblNum(eGDTransactionField_Profit) = dProfit
                
        ' do a simulated market order (demo)
'        If m.bAutoTrade And Not g.bUnloading Then
'            If g.Photon.ConnectionStatus = eGDConnectionStatus_Connected Then
'                With PtOrder
'                    .Buy = (nQty > 0)
'                    .OrderType = eTT_OrderType_Market
'                    .Quantity = Abs(nQty)
'                    '.Quantity = 2  '(just assume stop-and-reverse)
'                    .SymbolOrSymbolID = RollSymbolForDate(m.nSymbolID, m.dNextBarDate)
'                    .AccountID = DefaultAccount
'                End With
'                If 1 And PtOrder.AccountID <> 0 Then
'                    PtOrder.Save
'                    SubmitOrder PtOrder
'                ElseIf frmTTEditOrder.ShowMe(PtOrder, PtOrder.Buy, eGDTTEditOrderMode_Normal) = eGDEditOrderReturn_Submit Then
'                    SubmitOrder PtOrder
'                End If
'            End If
'        End If
                
        ' if an entry, save new entry price for conditional exits
        If InStr(strPos, "E") > 0 Then
            dBasePrice = dPrice
        Else
            ' if an exit, link to the prior open entry
            For i = m.Transactions.NumRecords - 2 To 0 Step -1
                If m.Transactions.TblNum(eGDTransactionField_Link, i) < 0 Then
                    i = i
                    m.Transactions.TblNum(eGDTransactionField_Link, i) = m.Transactions.NumRecords - 1
                    m.Transactions.TblNum(eGDTransactionField_Link) = i
                    Exit For
                End If
            Next
        End If
        
        ' calc stats (open position, etc.)
        ''CalcTradeStats m.dNextBarDate
        m.Transactions.CalculateStats m.dNextBarDate
        m.Transactions.CalculateEquityFilter
        
        ' save this order (for the conditional orders) and remove it
        strCondOrders = m.aOrders(iFillOrder)
        m.aOrders.Remove iFillOrder
        
        ' cancel other orders
        For i = m.aOrders.Size - 1 To 1 Step -1
            strText = Parse(m.aOrders(i), vbCrLf, 1)
            ' if an entry then cancel all other existing orders,
            ' else if an exit then just cancel the other existing exits
            If InStr(strPos, "X") = 0 Or InStr(Parse(strText, vbTab, 2), "X") > 0 Then
                m.aOrders.Remove i
            End If
        Next
        
        If Not m.Form.GameMode Is Nothing Then
            m.Form.GameMode.OrderFilled iFillOrder, Abs(nQty), dPrice
        End If
        
        ' then add any conditional orders tied to the new trade
        ' (which had been appended to the end of this order)
        For i = 2 To 999
            strText = Parse(strCondOrders, vbCrLf, i)
            If Len(strText) = 0 Then Exit For
            aFlds.SplitFields strText, vbTab
            ' conditional orders may be an offset from the new entry price
            If aFlds(5) = "EP" Then
                dPrice = Val(aFlds(4))
                dPrice = dBasePrice + dPrice
                strText = aFlds(0) & vbTab & aFlds(1) & vbTab _
                    & aFlds(2) & vbTab & aFlds(3) & vbTab & Str(dPrice)
            End If
            m.aOrders.Add strText
        Next
        
        FillOrders = True
    Loop
    
    ''CalcTradeStats m.dNextBarDate
    m.Transactions.CalculateStats m.dNextBarDate
    m.Transactions.CalculateEquityFilter

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.FillOrders", eGDRaiseError_Raise
End Function

Private Sub LoadScreenTrades()
On Error GoTo ErrSection:

    Dim i&, nPaneID&, nStyle&, nColor&, nStartX&, nEndX&, iRec&
    Dim nPerRow&, dY#, nCol&, nRow&, dDate#
    Dim dyLink#, dDateLink#, dProfit#
    Dim iLinkRec&, nLineColor&, nLastBar&, iLastRec&
    Dim bAsTradingUnits As Boolean
    Dim bDrawLine As Boolean
    Dim bBeforeScreen As Boolean
    Dim strText$, strParm$, strPos$, strAction$, strKey$
    Dim aFlds As New cGdArray
    Dim Annot As cAnnotation
    Dim Pane As cPane
    
    Dim iEndLoop&
    
    ' clear orders if old (for a previous bar)
    If m.aOrders.Size > 0 Then
        ' but keep orders if doing custom orders in game mode
        If m.Form.IsInGameMode And m.nSystemID = kGameModeSysID Or m.nSystemID = 0 Then
            m.dNextBarDate = m.Bars(eBARS_DateTime, Me.LastGoodDataBar(False))
        End If
        If m.Bars(eBARS_DateTime, LastGoodDataBar(False)) > m.dNextBarDate Then
            m.aOrders.Clear
        End If
    End If
    
    If m.Form.GameReplayMode <> 0 And m.Transactions.NumRecords > 0 And m.System Is Nothing Then
        ' clear any custom trades in future
        For iRec = m.Transactions.NumRecords - 1 To 0 Step -1
            If m.Bars.Prop(eBARS_PeriodType) = ePRD_Days And Not m.Form.GameMode.HasIntradayTicks Then
                If m.Bars.SessionDateForTime(m.Transactions.TblNum(eGDTransactionField_Date, iRec), False) <= m.Bars.SessionDateForTime(m.Form.GameMode.GameDataTime, False) Then
                    Exit For
                End If
            ElseIf gdFixDateTime(m.Transactions.TblNum(eGDTransactionField_Date, iRec)) <= gdFixDateTime(m.Form.GameMode.GameDataTime) Then
                Exit For
            End If
            i = m.Transactions.TblNum(eGDTransactionField_Link, iRec)
            If i >= 0 And i < m.Transactions.NumRecords Then
                m.Transactions.TblNum(eGDTransactionField_Link, i) = -1 '(clear linked entry/exit)
            End If
            m.Transactions.RemoveRecords iRec, 1
        Next
    End If
            
    If m.Form.GameReplayMode = 0 Then
        If m.nShowTrades = 0 Or m.Transactions.NumRecords + m.aOrders.Size = 0 Then
            RemoveAnnots False, eANNOT_UndefinedType, eANNOT_Trades
            RemoveAnnots False, eANNOT_UndefinedType, eANNOT_ProfitLines
            Exit Sub
        End If
    End If
        
    If m.nShowTrades = 1 Or m.Form.IsInGameMode Then
        FillOrders
    End If
    
    ' see if display in trading units or not
    bAsTradingUnits = True '(default)
    nPaneID = m.Tree.Index("PRICE PANE")
    If nPaneID > 0 Then
        Set Pane = m.Tree(nPaneID)
        If Pane.DisplayFormat <> ePANE_PriceFormat Then
            bAsTradingUnits = False
        End If
        Set Pane = Nothing
    End If
    
    nEndX = ScreenEndX
    nStartX = ScreenStartX
    
    ' flag all trade and profit lines to see later if used again or not
    For i = Annots.Count To 1 Step -1
        Set Annot = Annots(i)
        If Annot.eUsage = eANNOT_Trades Then
            Annot.Color = -999
        ElseIf Annot.eUsage = eANNOT_ProfitLines Then
            If m.nProfitLineBox = 1 And Annot.eType <> eANNOT_SimpleLine Then
                Annot.geRemoveAnnotation m.geChartObj
                Annots.Remove i
            ElseIf (m.nProfitLineBox = 2 Or m.nProfitLineBox = 3) And Annot.eType = eANNOT_SimpleLine Then
                Annot.geRemoveAnnotation m.geChartObj
                Annots.Remove i
            Else
                Annot.Color = -999
            End If
        End If
    Next
    
    ' Trades
    If m.Transactions.NumRecords > 0 Then
        nLastBar = LastGoodDataBar(False)
        ' start at last trade on screen
        dDate = gdFixDateTime(m.aXdate(nEndX))
        m.Transactions.FieldArray(eTRADE_Date, False).BinarySearch dDate + 0.000001, iLastRec
        iEndLoop = LimitTradeLoop(iLastRec, nLastBar)
        iLastRec = iLastRec - 1
        ' then look backwards through the trades
        For iRec = iLastRec To iEndLoop Step -1
            ' get date and price
            dDate = m.Transactions.TblNum(eGDTransactionField_Date, iRec)
            If dDate < m.aXdate(nStartX) Then
                bBeforeScreen = True
            End If
            dY = m.Transactions.TblNum(eGDTransactionField_Price, iRec)
            
            ' trade triangle: ignore this if next trade has same date and price or if off-screen
            If Not bBeforeScreen And (dDate <> m.Transactions.TblNum(eGDTransactionField_Date, iRec + 1) _
                    Or dY <> m.Transactions.TblNum(eGDTransactionField_Price, iRec + 1)) Then
                ' get action and position
                strPos = ""
                Select Case m.Transactions.TblStr(eGDTransactionField_Action, iRec)
                Case "EL"
                    nStyle = PEGAT_SMALLUPTRIANGLESOLID
                    nColor = ChartGlbClrForCtl(m.Form.pbChart, g.ChartGlobals.nLongColor, "LongColor")
                    strPos = "Long Entry"
                    strAction = "Buy"
                Case "ES"
                    nStyle = PEGAT_SMALLDOWNTRIANGLESOLID
                    nColor = g.ChartGlobals.nShortColor
                    strPos = "Short Entry"
                    strAction = "Sell"
                Case "XL"
                    If Not m.bHideExits Then
                        nStyle = PEGAT_SMALLDOWNTRIANGLE
                        nColor = g.ChartGlobals.nShortColor 'nLongColor
                        strPos = "Long Exit"
                        strAction = "Sell"
                    End If
                Case "XS"
                    If Not m.bHideExits Then
                        nStyle = PEGAT_SMALLUPTRIANGLE
                        nColor = ChartGlbClrForCtl(m.Form.pbChart, g.ChartGlobals.nLongColor, "LongColor") 'nShortColor
                        strPos = "Short Exit"
                        strAction = "Buy"
                    End If
                End Select

                If Len(strPos) > 0 Then
                    strKey = "Trade " & CStr(iRec)
                    Set Annot = Annots(strKey)
                    If Not Annot Is Nothing Then
                        ' reset color to indicate it is still being used
                        'aardvark fixes note:
                        'we are re-using an annotation that may have been created for a different trade
                        Annot.Color = nColor
                        Annot.Y(1) = dY             'trade price can changed (aardvark 3530 - user edited price in trade tracker)
                        Annot.dDate(1) = dDate      'trade date can change (aardvark 3545 - user switched account)
                        Annot.Style = nStyle
                    Else
                        ' create the annotation
                        Set Annot = New cAnnotation
                        Annot.CreateNew Me, eANNOT_Icon, nPaneID, dDate, _
                            dY, 0, 0, nColor, nStyle, "", eANNOT_Trades
                        Annots.Add Annot, strKey
                    End If
                    
                    ' for Markus Image Server, only show the first Entry (and no exits)
                    If m.bHideExits Then
                        Exit For
                    End If
                End If
            End If
            
            ' profit lines: look for them until hit flat position before screen
            ' (since means there can be no more trades spanning the screen)
            If bBeforeScreen And m.Transactions.TblNum(eGDTransactionField_Position, iRec) = 0 Then
                Exit For '(nothing else to do)
            ElseIf m.nProfitLineBox > 0 And Not m.bHideExits Then
                bDrawLine = False
                iLinkRec = m.Transactions.TblNum(eGDTransactionField_Link, iRec)
                If iLinkRec = -1 Then
                    ' for open trade: use close of last bar
                    dDateLink = m.Bars(eBARS_DateTime, nLastBar)
                    dyLink = m.Bars(eBARS_Close, nLastBar)
                    Select Case m.Transactions.TblStr(eGDTransactionField_Action, iRec)
                    Case "ES", "XS"
                        dProfit = dY - dyLink
                    Case Else
                        dProfit = dyLink - dY
                    End Select
                    bDrawLine = True
                ElseIf iLinkRec > iRec Then
                    ' for entry: draw line if exit is after beginning of screen
                    dDateLink = m.Transactions.TblNum(eGDTransactionField_Date, iLinkRec)
                    If dDateLink >= m.aXdate(nStartX) Then
                        dyLink = m.Transactions.TblNum(eGDTransactionField_Price, iLinkRec)
                        dProfit = m.Transactions.TblNum(eGDTransactionField_Profit, iLinkRec)
                        bDrawLine = True
                    End If
                End If
                ' ignore profit lines that start and stop on same bar
                If bDrawLine And dDate <> dDateLink Then
                    If dProfit > 0 Then
                        nLineColor = g.ChartGlobals.nWinColor
                    Else
                        nLineColor = g.ChartGlobals.nLossColor
                    End If
                    strKey = "ProfitLine " & CStr(iRec)
                    Set Annot = Annots(strKey)
                    If Not Annot Is Nothing Then
                        ' reset color to indicate it is still being used
                        Annot.Color = nLineColor
                        If dDateLink > 0 Then
                            Annot.dDate(2) = dDateLink
                            Annot.Y(1) = dY                 '6749
                            Annot.Y(2) = dyLink
                        End If
                        If m.nProfitLineBox = 2 Then        '3961
                            Annot.Prop("FillPattern") = 0
                        ElseIf m.nProfitLineBox = 3 Then
                            Annot.Prop("FillPattern") = 1
                            Annot.Prop("FillColor") = nLineColor
                        End If
                    Else
                        Set Annot = New cAnnotation
                        If m.nProfitLineBox = 1 Then        'profit lines
                            Annot.CreateNew Me, eANNOT_SimpleLine, nPaneID, dDate, dY, _
                                dDateLink, dyLink, nLineColor, 0, "", eANNOT_ProfitLines
                        ElseIf m.nProfitLineBox = 2 Then    'hollow profit boxes
                            Annot.CreateNew Me, eANNOT_Rectangle, nPaneID, dDate, dY, _
                                dDateLink, dyLink, nLineColor, 0, "", eANNOT_ProfitLines
                            Annot.PreIndicator = 1
                            Annot.Prop("FillPattern") = 0
                        ElseIf m.nProfitLineBox = 3 Then    'filled profit boxes
                            Annot.CreateNew Me, eANNOT_Rectangle, nPaneID, dDate, dY, _
                                dDateLink, dyLink, g.ChartGlobals.nChartBackColor, 0, "", eANNOT_ProfitLines
                            Annot.PreIndicator = 1
                            Annot.Prop("FillColor") = nLineColor
                            Annot.Prop("FillPattern") = 1
                        End If
                        If m.nProfitLineBox = 3 Then
                            Annot.Style = eANNOT_Thin
                        Else
                            Annot.Style = g.ChartGlobals.eProfitLineStyle
                        End If
                        Annot.SkipHitTestFlag = 1
                        Annots.Add Annot, strKey
                    End If
                End If
            End If
        Next
    End If
    
    ' remove all orders and leftover trade annotations
    For i = Annots.Count To 1 Step -1
        Set Annot = Annots(i)
        If Annot.eUsage = eANNOT_Trades Or Annot.eUsage = eANNOT_ProfitLines Then
            If Annot.Color = -999 Then
                Annot.geRemoveAnnotation m.geChartObj
                Annots.Remove i
            End If
        End If
    Next
    Set Annot = Nothing
            
    ' Orders
    If nEndX >= LastGoodDataBar(True) And m.dNextBarDate > 0 Then
        For i = 1 To m.aOrders.Size - 1
            strText = Parse(m.aOrders(i), vbCrLf, 1)
            aFlds.SplitFields strText, vbTab
            'ORDER:  1 XS  1177    S   1072.5              1
            ' get action and position
            Select Case aFlds(1)
            Case "EL"
                nStyle = PEGAT_SMALLUPTRIANGLESOLID
                nColor = ChartGlbClrForCtl(m.Form.pbChart, g.ChartGlobals.nLongColor, "LongColor")
                strPos = "Long Entry"
                strAction = "Buy"
            Case "ES"
                nStyle = PEGAT_SMALLDOWNTRIANGLESOLID
                nColor = g.ChartGlobals.nShortColor
                strPos = "Short Entry"
                strAction = "Sell"
            Case "XL"
                nStyle = PEGAT_SMALLDOWNTRIANGLE
                nColor = g.ChartGlobals.nShortColor 'nLongColor
                strPos = "Long Exit"
                strAction = "Sell"
            Case "XS"
                nStyle = PEGAT_SMALLUPTRIANGLE
                nColor = ChartGlbClrForCtl(m.Form.pbChart, g.ChartGlobals.nLongColor, "LongColor") 'nShortColor
                strPos = "Short Exit"
                strAction = "Buy"
            End Select
                
            ' build the text to display
            dY = Val(aFlds(4))
            Select Case aFlds(3)
            Case "S"
                strText = strAction & ": " & m.Bars.PriceDisplay(dY, bAsTradingUnits) & " Stop"
            Case "L"
                strText = strAction & ": " & m.Bars.PriceDisplay(dY, bAsTradingUnits) & " Limit"
            Case "M"
                dY = m.Bars(eBARS_Open, LastGoodDataBar(False))
                strText = strAction & ": Market"
            Case "MOC"
                dY = m.Bars(eBARS_Close, LastGoodDataBar(False))
                strText = strAction & ": Market on Close"
            Case "SCO"
                dY = m.Bars(eBARS_Close, LastGoodDataBar(False))
                strText = strAction & ": Stop Close Only"
            Case "LCO"
                dY = m.Bars(eBARS_Close, LastGoodDataBar(False))
                strText = strAction & ": Limit Close Only"
            Case Else
                strText = ""
            End Select
                
            ' create the annotation (flag an order by setting dY2 to -1)
            If Len(strText) > 0 Then
                Set Annot = New cAnnotation
                Annot.CreateNew Me, eANNOT_Icon, nPaneID, m.dNextBarDate, _
                    dY, 0, -1, nColor, nStyle, strText, eANNOT_Trades
                Annots.Add Annot, "Trade " & CStr(i - 1000)
            End If
        Next
        ' show message? (e.g. "no orders" or "see report")
        If m.aOrders.Size = 1 Then
            strText = m.aOrders(0)
            If Len(strText) > 0 Then
                Set Annot = New cAnnotation
                nStyle = 0
                nColor = ChartGlbClrForCtl(m.Form.pbChart, g.ChartGlobals.nLongColor, "LongColor")
                dY = m.Bars(eBARS_Close, LastGoodDataBar(False))
                Annot.CreateNew Me, eANNOT_Icon, nPaneID, m.dNextBarDate, _
                    dY, 0, -1, nColor, nStyle, strText, eANNOT_Trades
                Annots.Add Annot, "Trade " & CStr(-1)
            End If
        End If
    End If
    
    Set Annot = Nothing
    AddSpeed "Load Trades"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.LoadScreenTrades", eGDRaiseError_Raise

End Sub

' Prints/displays all other view_modes other than "normal".
'   view_mode:  0 = Normal
'   =========   1 = Export Image
'               2 = Print for Color
'               3 = Print for B&W (append "dashed", etc. to labels)
Public Sub PrintChart(Optional ByVal nMode& = 2, Optional ByVal ShowDialog& = True, _
        Optional ByVal strExportFile$ = "")
On Error GoTo ErrSection:

    Dim strMsg$, rc&, strCmd$, nNumPixels&, lpPoint As POINTSTRUCT
    Dim strErrFile$, hActiveWind&
    'msg = AskBox("i=? ; b=Print|+-View ; Do you wish to Print or pre-View the current graph?")
    'If AskBox("i=? ; Force to Monochrome? ; b=Yes|+-No") = "Y" Then
    '    graph.ViewingStyle = 1
    'End If

    If m.bZoomed Then
        InfBox "Please unzoom the chart before printing.", "!", , "Print Chart"
        Exit Sub
    End If


    m.iViewMode = nMode
    
    nMode = Abs(nMode)

    ' for now, printing is same either way
    If nMode = 2 Then nMode = 3

    ' for auto-saving BMP and GIF
    If ShowDialog > 10 Then
        nNumPixels = ShowDialog
        ShowDialog = 0
    End If
       
    DoEvents
    hActiveWind = GetActiveWindow
    strMsg = ""
    Select Case nMode
        Case 1: ' Export Image
            'GenerateChart
            If ShowDialog Then
                If 0 Then
                    ' dialog box is no longer modal, so must wait
                    ' till it goes away before resetting
                    Sleep 0.25
                    hActiveWind = GetActiveWindow
                    Do
                        Sleep 0.25
                    Loop While GetActiveWindow = hActiveWind
                    Sleep 0.25
                End If
            ElseIf nNumPixels > 0 Then
                'MJM - ask Tim about this
                'auto-save as a bitmap
                'lpPoint.X = nNumPixels
                'lpPoint.Y = lpPoint.X * 0.75
                'If Len(strExportFile) = 0 Then
                '    rc = PEcopybitmaptoclipboard(m.Peg, lpPoint)
                'Else
                '    KillFile strExportFile
                '    rc = PEcopybitmaptofile(m.Peg, lpPoint, strExportFile)
                'End If
                'strMsg = ""
            Else
                'MJM [begin] - modified for graphen.dll
                Clipboard.Clear
                HdrPaneId 1, 12, 8, 1
'                rc = geDrawChart
                'need to do this so sextant will get drawn
                rc = geSaveChart(m.geChartObj, m.Form.pbChart.hWnd, m.Form.pbChart.hDC, 0, 0, 0, "Clipboard")
                If rc = 0 Then
                    'Clipboard.SetData m.Form.pbChart.Image
                    strMsg = "You can now paste the graph into |another application by selecting |'Edit-Paste'  (or hit 'Ctrl-V')."
                Else
                    InfBox "i=[] ; h=Export Image ; Image could not be exported!"
                End If
                HdrPaneId 0
                rc = geDrawChart
             End If
            
        Case 2, 3: ' Print
            'GenerateChart
            'MJM [begin] - modified for grapheng.dll
             m.geChartHeader.paneId = 1
             frmPrintPreview.ShowMe "CNV frmChart", Me.Form, 0, 0.5, 0.5, 0.5, 0.5, True
             m.geChartHeader.paneId = 0
            'MJM [end] - modified for grapheng.dll
    End Select
    DoEvents

    ' reset to viewing graph
    m.iViewMode = 0
    'GenerateChart
    
    If Len(strMsg) > 0 Then
        InfBox "i=i ; h=Process Graph Image ; " + strMsg
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.PrintChart", eGDRaiseError_Raise

End Sub

' Returns display type of primary set of bars in template
Public Property Get BarDisplayType() As eIndicatorDisplayType
On Error GoTo ErrSection:

    Dim iFlag%, Indicator As cIndicator
    Dim eType As eIndicatorDisplayType
    
    eType = eINDIC_OHLC  '(default)
    
    If Not m.Tree Is Nothing Then
        With m.Tree
            Set Indicator = m.Tree("PRICE")
            If Not Indicator Is Nothing Then
                eType = Indicator.DisplayType
            End If
        End With
    End If
    Set Indicator = Nothing
    BarDisplayType = eType

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.BarDisplayType.Get", eGDRaiseError_Raise

End Property

' Set all bars to OHLC, Candlesticks, Close line, etc.
Public Property Let BarDisplayType(ByVal eType As eIndicatorDisplayType)
On Error GoTo ErrSection:

    Dim idx&, Indicator As cIndicator, strPane$, s$
    
    If eType = -999 Then
        'toggle between 3 most common types
        Select Case BarDisplayType
        Case eINDIC_OHLC
            eType = eINDIC_Candlestick
        Case eINDIC_Candlestick
            eType = eINDIC_Line
        Case eINDIC_Line
            eType = eINDIC_BollingerBar
        Case Else
            eType = eINDIC_OHLC
        End Select
    End If
    
    ' if is changing
    If eType <> BarDisplayType Then
        ' special check for Spread charts
        If Len(m.strSpreadSymbols) > 0 And eType < 0 And Not IsIntraday(m.nPeriodicity) Then
            s = "Displaying daily Open-High-Low data for| Spread charts requires loading intraday data |(with a much smaller amount of history)."
            If InfBox(s, "?", "+OK|-Cancel", "Bar Display for Spread Charts") = "C" Then
                eType = eINDIC_Line
            Else
                ' change to 1440 minute bars
                m.nPeriodicity = ePRD_Minutes + 1440
                RedoMode = eRedo9_ReloadData
                If m.nVertGrid = 1 Then
                    m.nVertGrid = 0 ' coarse
                End If
            End If
        End If
    End If
    
    ' fix bar period (if switching from special type)
    If eType >= eINDIC_OHLC Then
        Select Case GetPeriodType(m.nPeriodicity)
        Case ePRD_EodPF, ePRD_EodKagi, ePRD_EodRenko
            ' reset to daily
            m.nPeriodicity = ePRD_Days + 1
            RedoMode = eRedo9_ReloadData
        Case ePRD_IntPF, ePRD_IntKagi, ePRD_IntRenko
            ' reset to 5 minute bars
            m.nPeriodicity = ePRD_Minutes + 5
            RedoMode = eRedo9_ReloadData
        End Select
    End If
    
    If Not m.Tree Is Nothing Then
        For idx = 1 To m.Tree.Count
            If m.Tree.NodeLevel(idx) = 0 Then
                strPane = m.Tree.Key(idx)
            Else
                Set Indicator = m.Tree(idx)
                With Indicator
                    If .DataType = eINDIC_BarData Then
                        If m.Tree.Key(idx) = "PRICE" Or strPane <> "PRICE PANE" Then
                            .DisplayType = eType
                        End If
                    End If
                End With
            End If
        Next
    End If

    SyncToolbar

ErrExit:
    Set Indicator = Nothing
    Exit Property
    
ErrSection:
    Set Indicator = Nothing
    RaiseError "cChart.BarsDisplayType.Let", eGDRaiseError_Raise

End Property

Public Property Get Periodicity() As Long
    Periodicity = m.nPeriodicity
End Property

'Public Property Let Periodicity(ByVal nPeriodicity As Long)
'    m.nPeriodicity = nPeriodicity
'End Property

Public Property Get ShowEmptyBars() As Boolean
    ShowEmptyBars = m.bShowEmptyBars
End Property

Public Property Let ShowEmptyBars(ByVal bShowEmptyBars As Boolean)
    m.bShowEmptyBars = bShowEmptyBars
End Property

Public Property Get ForecastBars(Optional frm As Form) As Long
    ForecastBars = m.nForecastBars
End Property

Public Property Let ForecastBars(frm As Form, ByVal nBars&)
    'only allow the chart's form to change this value
    If Not frm Is Nothing Then
        If IsFrmChart(frm) Then
            ' If we are adding more bars beyond the current hsb.Max (e.g. when dragging left)
            ' then we need to set the arrow left now so ResetLastScreenDate won't get called.
            If nBars > ForecastBars Then
                m.geChartStruct.scaleArrow = 1
            End If
            MyForecastBars = nBars
        End If
    End If
End Property

Private Property Let MyForecastBars(nBars&)
    m.nForecastBars = nBars
    If m.nForecastBars < m.nBlankBars Then m.nForecastBars = m.nBlankBars
    
    If g.ChartGlobals.bExtForecastBars Then
        '02-12-2013: code for Elliot Wave always save space to right of chart past end of data
        'nForecastBars is always saved on a chart move, but gets reset to blank bars on full eRedo9_ReloadData
        m.nBlankBars = m.nForecastBars
    End If

End Property

Public Property Get BlankBars(Optional frm As Form) As Long
    BlankBars = m.nBlankBars
End Property

Public Property Let BlankBars(frm As Form, ByVal nBars&)
    'only allow the chart editor to change this value
    If Not frm Is Nothing Then
        If TypeOf frm Is frmChartCfg Or (m.bProfileChart And TypeOf frm Is frmNewChart) Then
            MyBlankBars = nBars
            MyForecastBars = nBars
        End If
    End If
End Property

Private Property Let MyBlankBars(ByVal nBars&)
    m.nBlankBars = nBars
    If m.nBlankBars < 1 And TypeOfChart <> eTypeChart_Seasonal Then m.nBlankBars = 1
    m.nBlankBarsSave = m.nBlankBars         'aardvark 6780 - save to restore when changing symbols
End Property

Public Property Get Unsplit() As Boolean
    Unsplit = m.bUnsplit
End Property

Public Property Let Unsplit(ByVal bUnsplit As Boolean)
    m.bUnsplit = bUnsplit
End Property

Public Property Get VertGrid() As Long
    VertGrid = m.nVertGrid
End Property

Public Property Let VertGrid(ByVal nGrid&)
    m.nVertGrid = nGrid     '0=coarse, 1=fine, 2=none
End Property

Public Property Get HorzGrid() As Boolean
    HorzGrid = m.bHorzGrid
End Property

Public Property Let HorzGrid(ByVal bHorzGrid As Boolean)
    m.bHorzGrid = bHorzGrid
End Property

Public Property Get Zoomed() As Boolean
    Zoomed = m.bZoomed
End Property

Public Property Let Zoomed(ByVal bZoomed As Boolean)
    m.bZoomed = bZoomed
End Property

Public Property Get aXBar() As cGdArray
    Set aXBar = m.aXBar
End Property

Public Property Get aXdate() As cGdArray
    Set aXdate = m.aXdate
End Property

Public Property Get ShowTrades() As Long
    'used to be boolean: true=show trades from strategy
    '                    false=don't show trades from strategy
    'changed to long: 0=don't show trades or orders
    '                 1=show trades from strategy
    '                 2=show trades & orders from trading account
    ShowTrades = Abs(m.nShowTrades)
End Property

Public Property Let ShowTrades(ByVal nShowTrades As Long)
On Error GoTo ErrSection:

    If m.nShowTrades <> nShowTrades Then
        'remove all trades/orders since a change means we need to either clear
        'or load different trades/orders anyhow - aardvark 3354 fix
        RemoveAnnots False, eANNOT_UndefinedType, eANNOT_Trades
        RemoveAnnots False, eANNOT_UndefinedType, eANNOT_ProfitLines
        m.Transactions.NumRecords = 0
        m.aOrders.Clear
        If nShowTrades = 2 Then
            m.bReloadTrackerTrades = True   'need to request regeneration of trades file for account
        Else
            m.bReloadTrackerTrades = False
            If nShowTrades = 1 Then
                'set systemID to -1 so generate chart will re-run system
                If Not m.System Is Nothing Then
                    m.nSystemID = -1
                End If
            End If
        End If
        m.nShowTrades = Abs(nShowTrades)
        SyncToolbar
    End If
    
    Exit Property
    
ErrSection:
    RaiseError "cChart.ShowTrades.Let"
    
End Property

Public Function ScreenStartX() As Long
On Error GoTo ErrSection:

    Dim nPoints&

    If m.Form Is Nothing Then
        ScreenStartX = 0
    Else
        'MJM - modified for use with grapheng.dll
        If m.geChartObj <> 0 Then
            If m.geChartStruct.x_dataPointCount > nPoints Then nPoints = m.geChartStruct.x_dataPointCount
        End If
        If m.iMiniChart > 0 Then
            ScreenStartX = m.iMiniChartStartX
        Else
            ScreenStartX = m.hsb.Value - nPoints + 1
        End If
        If ScreenStartX < 0 Then ScreenStartX = 0
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ScreenStartX", eGDRaiseError_Raise

End Function

Public Function ScreenEndX() As Long
On Error GoTo ErrSection:

    Dim nEndX&
    If m.Form Is Nothing Then
        ScreenEndX = 0
    Else
        nEndX = m.hsb.Value
        
        If m.bZoomed = True And m.aXdate(nEndX) = kNullData Then
            'this is necessary because the hsb max has been faked out
            'to allow scrolling forward while zoomed
            'fixes zoom trade bug aardvark #229
            nEndX = LastGoodDataBar(True) + 1
        End If
        
        ScreenEndX = nEndX
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ScreenEndX", eGDRaiseError_Raise

End Function

Public Function ScreenNumX() As Long
On Error GoTo ErrSection:

    If m.Form Is Nothing Then
        ScreenNumX = 0
    Else
        ScreenNumX = m.geChartStruct.x_dataPointCount   'm.Peg.Points  - MJM
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ScreenNumX", eGDRaiseError_Raise

End Function

Public Function HighestHiForwardDate(ByVal dDate#, dHi#, dDateHi#) As Boolean
On Error GoTo ErrSection:

    Dim i&, iPos&, dY#
    
    dHi = kNullData
    dDateHi = kNullData
    gdBinarySearch m.Bars.ArrayHandle(eBARS_DateTime), gdFixDateTime(dDate), iPos, eGdSort_Default, 0, m.Bars.Size - 1
    For i = iPos To m.Bars.Size - 1
        dY = m.Bars(eBARS_High, i)
        If dY <> kNullData Then
            If dHi = kNullData Or dY > dHi Then
                dHi = dY
                dDateHi = m.Bars(eBARS_DateTime, i)
            End If
        End If
    Next
    
    If dHi = kNullData Then
        HighestHiForwardDate = False
    Else
        HighestHiForwardDate = True
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.HighestHiForwardDate", eGDRaiseError_Raise

End Function

Public Function LowestLowForwardDate(ByVal dDate#, dLow#, dDateLow#) As Boolean
On Error GoTo ErrSection:

    Dim i&, iPos&, dY#
    
    dLow = kNullData
    dDateLow = kNullData
    gdBinarySearch m.Bars.ArrayHandle(eBARS_DateTime), gdFixDateTime(dDate), iPos, eGdSort_Default, 0, m.Bars.Size - 1
    For i = iPos To m.Bars.Size - 1
        dY = m.Bars(eBARS_Low, i)
        If dY <> kNullData Then
            If dLow = kNullData Or dY < dLow Then
                dLow = dY
                dDateLow = m.Bars(eBARS_DateTime, i)
            End If
        End If
    Next
    
    If dLow = kNullData Then
        LowestLowForwardDate = False
    Else
        LowestLowForwardDate = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.LowestLowForwardDate", eGDRaiseError_Raise

End Function

Private Sub LoadNotations()
On Error GoTo ErrSection:

    Dim i&, nPaneID&, nStyle&, nColor&, nStartX&, nEndX&
    Dim nPerRow&, nX&, dY#, nCol&, nRow&, dDate#, dText#
    Dim strText$, strParm$, n&, dCrossoverOffset#
    Dim Annot As cAnnotation
    
    RemoveAnnots False, eANNOT_UndefinedType, eANNOT_Notation
    
    If Not g.ChartGlobals.bSplitsRolls Then Exit Sub
    
    nEndX = ScreenEndX
    nStartX = ScreenStartX
    
    'MJM NOTE - this check prevents infinite loop (see note later in this subroutine)
    If nStartX >= nEndX Then
        Exit Sub
    End If
    
    nPaneID = m.Tree.Index("PRICE PANE")
    'm.Peg.WorkingAxis = nPaneID  - MJM
    'dText = TextHeightAsY(nPaneID)
    'dy = m.Peg.ManualMinY + dText * 1.5    - MJM
    nColor = g.ChartGlobals.nSplitRollColor
    
    If m.RollsTable.NumRecords > 0 And m.nPeriodicity <= ePRD_Days + 1 Then
        ' calc crossover offset for intraday charts (in case has an overnight session)
        If m.Bars.IsIntraday Then
            dCrossoverOffset = m.Bars.Prop(eBARS_CrossoverTime) / 1440#
            If dCrossoverOffset <> 0 Then
                dCrossoverOffset = 1 - dCrossoverOffset
            End If
        End If
        'find first one after start of screen
        dDate = m.aXdate(nStartX)
        m.RollsTable.FieldArray(1, False).BinarySearch dDate, n
        If n <= 0 Then n = 1 '(skip the first record)
        nX = nStartX
        For i = n To m.RollsTable.NumRecords - 1
            dDate = gdFixDateTime(m.RollsTable(1, i) - dCrossoverOffset)
            If dDate > m.aXdate(nEndX) Then Exit For
            Do While dDate > m.aXdate(nX)
                nX = nX + 1     'MJM NOTE - infinite loop here (see note above)
            Loop
            'If dDate > LastDailyDownload Then
            If m.RollsTable(2, i) = kNullData Then
                strText = "?"
            Else
                strText = "R"
            End If
            
            Set Annot = New cAnnotation
            Annot.CreateNew Me, eANNOT_Icon, nPaneID, dDate, _
                dY - dText * 0.75, 0, 0, nColor, 0, strText, eANNOT_Notation
            Annots.Add Annot, "Roll2 " & CStr(i)
        Next
        
        AddSpeed "Load Rolls"
    End If

    If m.SplitsTable.NumRecords > 0 Then
        'find first one after start of screen
        dDate = m.aXdate(nStartX)
        m.SplitsTable.FieldArray(0, False).BinarySearch dDate, n
        nX = nStartX
        For i = n To m.SplitsTable.NumRecords - 1
            dDate = gdFixDateTime(m.SplitsTable(0, i))
            If dDate > m.aXdate(nEndX) Then Exit For
            Do While dDate > m.aXdate(nX)
                nX = nX + 1
            Loop
            strText = "S"
            
            Set Annot = New cAnnotation
            Annot.CreateNew Me, eANNOT_Icon, nPaneID, dDate, _
                dY - dText * 0.75, 0, 0, nColor, 0, strText, eANNOT_Notation
            Annots.Add Annot, "Split2 " & CStr(i)
        Next
        
        AddSpeed "Load Splits"
    End If
    
    ' show dividends (except for monthly bars and greater)
    If m.DividendsTable.NumRecords > 0 And m.nPeriodicity < ePRD_Months Then
        'find first one after start of screen
        dDate = m.aXdate(nStartX)
        m.DividendsTable.FieldArray(2, False).BinarySearch dDate, n
        nX = nStartX
        For i = n To m.DividendsTable.NumRecords - 1
            dDate = gdFixDateTime(m.DividendsTable(2, i))
            If dDate > m.aXdate(nEndX) Then Exit For
            Do While dDate > m.aXdate(nX)
                nX = nX + 1
            Loop
            strText = "D"
            
            Set Annot = New cAnnotation
            Annot.CreateNew Me, eANNOT_Icon, nPaneID, dDate, _
                dY - dText * 0.75, 0, 0, nColor, 0, strText, eANNOT_Notation
            Annots.Add Annot, "Dividend2 " & CStr(i)
        Next
        
        AddSpeed "Load Dividends"
    End If
    
    ' TLB 6/26/2013: display next earnings date
    dDate = m.nNextEarningsDate
    If dDate > 0 And dDate >= m.aXdate(nStartX) And dDate <= m.aXdate(nEndX) Then
        nX = nStartX
        Do While dDate > m.aXdate(nX)
            nX = nX + 1
        Loop
        strText = "E"
        
        Set Annot = New cAnnotation
        Annot.CreateNew Me, eANNOT_Icon, nPaneID, dDate, _
            dY - dText * 0.75, 0, 0, nColor, 0, strText, eANNOT_Notation
        Annots.Add Annot, "Earn2 " & CStr(m.nNextEarningsDate)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.LoadNotations", eGDRaiseError_Raise

End Sub

Private Sub DumpAnnots(ByVal strFile$)
On Error GoTo ErrSection:

    Dim i&, Annot As cAnnotation
    Dim a As New cGdArray

    For i = 1 To m.Annots.Count
        Set Annot = m.Annots(i)
        With Annot
            a.Add Str(i) & vbTab & .eUsage & vbTab & .eType & vbTab & .Pane & vbTab & .Y(1)
        End With
    Next
    Set Annot = Nothing
    a.ToFile App.Path & "\" & strFile & ".ann"
    Set a = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.DumpAnnots", eGDRaiseError_Raise

End Sub

Public Function PriceDisplay(ByVal idx&, ByVal dPrice#, Optional ByVal nForSessionDate& = 0, _
                    Optional ByVal iMaxDigits% = 0) As String
On Error GoTo ErrSection:

    Dim i&, idxPane&, iLen&, iPos&, strPrice$, strFmt$
    Dim Pane As cPane, Ind As cIndicator
    Dim b As cGdBars, idxBars&
           
    If iMaxDigits <= 0 Then iMaxDigits = 999
           
    ' get pane idx
    idxPane = m.Tree.RelativeIndex(idx, eTREE_Root)
    If idxPane > 0 Then
        Set Pane = m.Tree(idxPane)
        If Pane.DisplayFormat = ePANE_PriceFormat Then
            ' get price bars for pane or indicator
            If idx = idxPane Then
                idxBars = Pane.idxPriceBars
            Else
                Set Ind = m.Tree(idx)
                idxBars = Ind.idxPriceBars
            End If
            If idxBars = 0 Then
                Set b = m.Bars
            ElseIf idxBars > 0 Then
                Set Ind = m.Tree(idxBars)
                Set b = Ind.Bars
            Else
                Set b = Nothing
            End If
            If Not b Is Nothing Then
                If nForSessionDate <= 0 Then
                    ' default: use session date of right-most visible bar
                    i = LastGoodDataBar(False, True)
                    nForSessionDate = Me.Bars.SessionDate(i)
                End If
                strPrice = b.PriceDisplay(dPrice, True, nForSessionDate)
                Set b = Nothing
            End If
            iLen = Len(strPrice)
            If iLen > iMaxDigits Then
                ' try to strip off some of fraction
                iPos = InStr(strPrice, ".")
                If iPos > 0 Then
                    For i = iLen To iPos + 1 Step -1
                        If Mid(strPrice, iLen, 1) <> "0" Then Exit For
                        iLen = iLen - 1
                    Next
                    If Mid(strPrice, iLen, 1) = "." Then
                        iLen = iLen - 1
                    End If
                End If
                If iLen > iMaxDigits Then
                    ' try to strip off caret part
                    iPos = InStr(strPrice, "^")
                    If iPos > 0 Then
                        For i = iPos + 1 To iLen
                            If Mid(strPrice, i, 1) <> "0" Then Exit For
                            If i = iLen Then iLen = iPos - 1
                        Next
                    End If
                End If
                strPrice = Left(strPrice, iLen)
            End If
        ElseIf Pane.DisplayFormat = ePANE_RoundDecimals Then        '6683
            If Pane.DisplayDecimals > 0 Then
                strFmt = "0." & String(Pane.DisplayDecimals, "0")
            Else
                strFmt = "0"
            End If
            strPrice = Format(dPrice, strFmt)
        Else
            strPrice = RoundedValueStr(dPrice, Pane.gePaneMax - Pane.gePaneMin)
        End If
        Set Ind = Nothing
        Set Pane = Nothing
    End If
    If Len(strPrice) = 0 Then strPrice = CStr(dPrice)

    PriceDisplay = strPrice

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.PriceDisplay", eGDRaiseError_Raise

End Function

Public Property Get RollsTable() As cGdTable
    Set RollsTable = m.RollsTable
End Property

Public Sub ResetLastScreenDate()
    m.dLastScreenDate = 0
End Sub

' returns true if did a redisplay
Public Function UpdateNewTicks() As Boolean
On Error GoTo ErrSection:

    Dim bNewBar As Boolean, bRedisplay As Boolean, bDebug As Boolean, bVisible As Boolean
    Dim i&, d#, dHigh#, dLow#, dClose#, nBar&, nPeriodType&, dDelay#
    Dim strSpreadSymbols As String
    Dim barsOther As cGdBars
    Dim aHandles As New cGdArray
    Dim frm As Form
    Dim bResetNextBarDate As Boolean
    Static bPriceUpdated As Boolean, bVolUpdated As Boolean
    
    If m.bRT And Not m.bDisableRT And Not g.bUnloading Then

If IsIDE Then
    ' for debugging
    If Me.Form Is ActiveChart Then
        'bDebug = frmTest.Visible
    End If
    
    ' to simulate an active market (new tick every X milliseconds)
    If 0 Then ' gdTickCount(False) > m.dLastRecalcDone + 250 Then
        bRedisplay = True
        bPriceUpdated = True
        RedoMode = eRedo5_RecalcInd
    End If
End If
        
        ' update primary bars
        If m.nSymbolID > 0 Then
            nPeriodType = m.Bars.Prop(eBARS_PeriodType)
            nBar = LastGoodDataBar(False)
            dClose = m.Bars(eBARS_Close, nBar)
            dHigh = m.Bars(eBARS_High, nBar)
            dLow = m.Bars(eBARS_Low, nBar)
            If m.dNextBarDate = m.Bars(eBARS_DateTime, LastGoodDataBar(False)) Then
                bResetNextBarDate = True        '5517
            End If
            If g.RealTime.UpdateBars(m.Bars, bNewBar) Then
                If bDebug Then frmTest.AddList Str(Int(gdTickCount(False))) & vbTab & "update" & vbTab & Str(bNewBar)
                ' set redisplay flags based on what changed
                If bNewBar Then
                    RedoMode = eRedo7_ReloadRT
                    bResetNextBarDate = False
If IsIDE Then
    If frmTest.Visible Then
        frmTest.AddList m.Bars.Prop(eBARS_Symbol) & " at " & Format(m.Bars(eBARS_DateTime, nBar), "HH:MM") _
            & " = " & Str(m.Bars(eBARS_DownTicks, nBar) + m.Bars(eBARS_UpTicks, nBar)) & " trades"
    End If
End If
                ElseIf dClose <> m.Bars(eBARS_Close, nBar) Or dHigh <> m.Bars(eBARS_High, nBar) Or dLow <> m.Bars(eBARS_Low, nBar) Then
                    bRedisplay = True
                    bPriceUpdated = True
                Else
                    bVolUpdated = True
                End If
                If bResetNextBarDate Then
                    m.dNextBarDate = m.Bars(eBARS_DateTime, LastGoodDataBar(False))
                End If
                ' reset the current data/bar time
                m.dLastActualDateOfData = Int(m.dLastActualDateOfData) + m.Bars.Prop(eBARS_LastTickTime) / 1440#
                If nPeriodType < ePRD_Days And nPeriodType <> ePRD_Minutes Then
                    d = m.Bars(eBARS_DateTime, nBar)
                    i = LastGoodDataBar(True)
                    If d >= m.aXdate(i + 1) Then
                        RedoMode = eRedo7_ReloadRT '(need to update all the remaining bar times)
                    ElseIf d > m.aXdate(i) Then
                        m.aXdate(i) = d '(just update the current bar time)
                    End If
                End If
            End If
            
            If m.bProfileChart Then
                Dim nSession&, nSessionNext&, nSessionMostRecent&, j
                
                Dim BarsProfileRT As New cGdBars
                Dim eStatus As eProfileStartusRT
                
                nSession = LastDailyDownload
                nSessionNext = nSession + 1
                
                While gdIsHoliday(nSessionNext, "") Or Not IsWeekday(nSessionNext)
                    nSessionNext = nSessionNext + 1
                Wend
                
                nSessionMostRecent = m.Bars.SessionDate(m.Bars.Size - 1)
                
                If m.BarsProfileTicks Is Nothing Then
                    eStatus = ProfileUpdateRT(m.BarsProfileTicks, m.BarsProfile, BarsProfileRT, nSession, nSessionNext, nSessionMostRecent, True)
                Else
                    eStatus = ProfileUpdateRT(m.BarsProfileTicks, m.BarsProfile, BarsProfileRT, nSession, nSessionNext, nSessionMostRecent, False)
                End If
                
                If eStatus = eProfileRT_NewSession Then
                    eStatus = eStatus
                End If
                
            End If
            
        ElseIf Len(m.strSpreadSymbols) > 0 Then
            ' update spread bars
            aHandles.Create eGDARRAY_Longs
            For i = 0 To UBound(m.SpreadBars)
                Set barsOther = m.SpreadBars(i)
                If Not barsOther Is Nothing Then
                    aHandles(i) = barsOther.BarsHandle
                    If Len(strSpreadSymbols) = 0 Then
                        strSpreadSymbols = barsOther.Prop(eBARS_Desc)
                    Else
                        strSpreadSymbols = strSpreadSymbols & ";" & barsOther.Prop(eBARS_Desc)
                    End If
                    For nBar = barsOther.Size - 1 To 0 Step -1
                        If barsOther(eBARS_Close, nBar) <> kNullData Then Exit For
                    Next
                    dClose = barsOther(eBARS_Close, nBar)
                    dHigh = barsOther(eBARS_High, nBar)
                    dLow = barsOther(eBARS_Low, nBar)
                    If g.RealTime.UpdateBars(barsOther, bNewBar) Then
                        If bNewBar Then
                            RedoMode = eRedo7_ReloadRT
                        ElseIf dClose <> barsOther(eBARS_Close, nBar) Or dHigh <> barsOther(eBARS_High, nBar) Or dLow <> barsOther(eBARS_Low, nBar) Then
                            bRedisplay = True
                            bPriceUpdated = True
                        Else
                            'bVolUpdated = True
                        End If
                    End If
                End If
            Next
            If bPriceUpdated And Not bNewBar Then
                Set barsOther = m.Bars.MakeCopy(True)
                barsOther.Prop(eBARS_Periodicity) = m.SpreadBars(0).Prop(eBARS_Periodicity)
                aHandles.Add barsOther.BarsHandle '' m.Bars.BarsHandle
                If 0 Then
                    i = gdCalcSpread(aHandles.ArrayHandle, strSpreadSymbols, 0)
                    m.Bars.Size = 0
                    m.Bars.Prop(eBARS_Periodicity) = m.nPeriodicity
                    m.Bars.BuildBars m.Bars.Prop(eBARS_PeriodicityStr), barsOther.BarsHandle
                Else
                    i = gdCalcSpread(aHandles.ArrayHandle, strSpreadSymbols, 1)
                    d = barsOther(eBARS_Close, barsOther.Size - 1)
                    i = LastGoodDataBar(False)
                    If m.Bars(eBARS_Close, i) = kNullData Or d = kNullData Then
                        ' this probably shouldn't happen, but just in case ...
                        RedoMode = eRedo7_ReloadRT ' ??
                    Else
                        m.Bars(eBARS_Close, i) = d
                        'If m.Bars.Prop(eBARS_Periodicity) = ePRD_Minutes + 1 Then
                        If 0 Then
                            If d < m.Bars(eBARS_Open, i) Then
                                m.Bars(eBARS_High, i) = m.Bars(eBARS_Open, i)
                            Else
                                m.Bars(eBARS_High, i) = d
                            End If
                            If d > m.Bars(eBARS_Open, i) Then
                                m.Bars(eBARS_Low, i) = m.Bars(eBARS_Open, i)
                            Else
                                m.Bars(eBARS_Low, i) = d
                            End If
                        ElseIf d > m.Bars(eBARS_High, i) Then
                            m.Bars(eBARS_High, i) = d
                        ElseIf d < m.Bars(eBARS_Low, i) Then
                            m.Bars(eBARS_Low, i) = d
                        End If
                    End If
                End If
                Set barsOther = Nothing
            End If
        End If
        
        ' update secondary bars
        If Not bNewBar Then
            For i = 1 To UBound(m.OtherBars)
                Set barsOther = m.OtherBars(i)
                If Not barsOther Is Nothing Then
                    For nBar = barsOther.Size - 1 To 0 Step -1
                        If barsOther(eBARS_Close, nBar) <> kNullData Then Exit For
                    Next
                    dClose = barsOther(eBARS_Close, nBar)
                    dHigh = barsOther(eBARS_High, nBar)
                    dLow = barsOther(eBARS_Low, nBar)
                    If g.RealTime.UpdateBars(barsOther, bNewBar) Then
                        If UCase(barsOther.Prop(eBARS_CustomString)) = "HEIKIN-ASHI" Then
                            ' for Heikin-Ashi bars, the open/high/low should update correctly
                            ' but the Close needs to be recalculated (avg OHLC of source data)
                            barsOther(eBARS_Close, nBar) = (m.Bars(eBARS_Open, nBar) + m.Bars(eBARS_High, nBar) + m.Bars(eBARS_Low, nBar) + m.Bars(eBARS_Close, nBar)) / 4
                        End If
                        If bNewBar Then
                            RedoMode = eRedo7_ReloadRT
                        ElseIf dClose <> barsOther(eBARS_Close, nBar) Or dHigh <> barsOther(eBARS_High, nBar) Or dLow <> barsOther(eBARS_Low, nBar) Then
                            bRedisplay = True
                            bPriceUpdated = True
                        Else
                            bVolUpdated = True
                        End If
                    End If
                End If
            Next
        End If
        
        ' see if this chart is visible (e.g. not visible if another chart is maximized or this one is minimized)
        Set frm = ActiveChart
        If Not frm Is Nothing Then
            If m.Form.WindowState <> vbMinimized Then
                If frm Is m.Form Or frm.WindowState <> vbMaximized Or m.Form.DetachStatus = eDetached Then  '5357
                    bVisible = True
                End If
            End If
        End If
        Set frm = Nothing
        
        ' check if need to do a delayed recalc
        ' first see if user has set the recalcs to only be on a new bar
        If g.nRecalcIndRT >= 0 Then ' and bVisible (this was taken out from here because of chart alerts - 4492)
            ' no need to recalc until the price or vol has changed
            If bPriceUpdated Or bVolUpdated Then
                ' do a more immediate recalc if price changes on a visible chart
                If bPriceUpdated And bVisible Then
                    dDelay = m.dAvgRecalcTime * 4
                Else ' otherwise do a delayed recalc (e.g. once a second)
                    dDelay = m.dAvgRecalcTime * 8
                    If dDelay < 1000 Then dDelay = 1000
                End If
                ' but check if user has set the global setting to only recalc every X seconds
                If dDelay < g.nRecalcIndRT * 1000 Then dDelay = g.nRecalcIndRT * 1000
                ' see if time to recalc
                If gdTickCount(False) > m.dLastRecalcDone + dDelay Then
                    RedoMode = eRedo5_RecalcInd
                End If
            End If
        End If
        If RedoMode >= eRedo5_RecalcInd Then
            If bDebug Then frmTest.AddList Str(Int(gdTickCount(False))) & vbTab & "recalc" & vbTab & Str(bPriceUpdated) & vbTab & Str(bVolUpdated)
            bRedisplay = True
            bPriceUpdated = False
            bVolUpdated = False
        End If
        
        ' see if need to force a redisplay now
        If bRedisplay Or m.bReloadTrackerTrades Then
            bPriceUpdated = False
            If bDebug Then frmTest.AddList Str(Int(gdTickCount(False))) & vbTab & "display" & vbTab & Str(RedoMode)
            If m.iTryLastBarRecalc >= 0 And m.eRedoMode = eRedo5_RecalcInd Then
                m.iTryLastBarRecalc = 1
            End If
            GenerateChart eRedo1_Scrolled, True
            m.bReloadTrackerTrades = False
            UpdateNewTicks = True
        End If

    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.UpdateNewTicks", eGDRaiseError_Raise

End Function

Public Property Get HidePriceIndicators() As Boolean
    HidePriceIndicators = m.bHidePriceIndicators
End Property

Public Property Get geGridMode() As Long
    geGridMode = m.geChartStruct.gridMode
End Property

Public Property Let geGridMode(ByVal nMode&)
    m.geChartStruct.gridMode = nMode
End Property

Public Property Let HidePriceIndicators(ByVal bHide As Boolean)
    m.bHidePriceIndicators = bHide
End Property

Public Property Get FromDate() As Double
    FromDate = m.dFromDate
End Property

Public Property Let FromDate(ByVal dDate As Double)
    m.dFromDate = dDate
End Property

Public Property Get ToDate() As Double
    ToDate = m.dToDate
End Property

Public Property Let ToDate(ByVal dDate As Double)
    m.dToDate = dDate
End Property

Public Property Get ToEndOfData() As Boolean
    ToEndOfData = m.bToEndOfData
End Property

Public Property Let ToEndOfData(ByVal bToEndOfData As Boolean)
    m.bToEndOfData = bToEndOfData
End Property

Public Property Get MaxIntradayDays() As Long
    MaxIntradayDays = m.nMaxIntradayDays
End Property

Public Property Let MaxIntradayDays(ByVal nMax As Long)
    m.nMaxIntradayDays = nMax
End Property

' To auto-change black indicators to yellow when chart
' background changes from white to black (and vice versa).
Public Sub FixColors()
On Error GoTo ErrSection:

    Dim idx&, nLightYellow&, nChartBackColor&
    Dim Ind As cIndicator
    
    'nLightYellow = RGB(255, 255, 128)
    nLightYellow = RGB(0, 254, 0)
    If m.bCustomColors Then
        nChartBackColor = m.nChartBackColor
    Else
        nChartBackColor = g.ChartGlobals.nChartBackColor
    End If
    
    If Not m.Tree Is Nothing Then
        For idx = 1 To m.Tree.Count
            If m.Tree.NodeLevel(idx) > 0 Then
                Set Ind = m.Tree(idx)
                If nChartBackColor = 0 Then
                    If Ind.Color = 0 Then Ind.Color = nLightYellow
                ElseIf nChartBackColor = RGB(255, 255, 255) Then
                    If Ind.Color = nLightYellow Or Ind.Color = RGB(255, 255, 255) Then
                        Ind.Color = 0
                    End If
                End If
            End If
        Next
    End If
    
ErrExit:
    Set Ind = Nothing
    Exit Sub
    
ErrSection:
    Set Ind = Nothing
    RaiseError "cChart.FixColors", eGDRaiseError_Raise

End Sub

Public Function RollSymbol(ByVal nX&) As String
On Error GoTo ErrSection:

    Dim dDate#, nBar&, nSymbolID&, dCrossoverOffset#
    Static nRollRecord&, strSymbol$, nPrevSymbolID&
    
    nBar = m.aXBar(nX) 'could be < 0
    dDate = m.aXdate(nX)
    ' get logical trading date
    If dDate <> Int(dDate) Then
        dCrossoverOffset = m.Bars.Prop(eBARS_CrossoverTime) / 1440#
        If dCrossoverOffset <> 0 And dDate - Int(dDate) > dCrossoverOffset Then
            dDate = Int(dDate) + 1
        Else
            dDate = Int(dDate)
        End If
    End If
    
    With m.RollsTable
        ' allow showing roll for next bar (in case next bar has rolled to new contract)
        If .NumRecords <= 0 Or nBar > LastGoodDataBar(False) + 1 Then
            strSymbol = ""
            nPrevSymbolID = 0
        Else
            If nRollRecord >= .NumRecords Then nRollRecord = .NumRecords - 1
            If nRollRecord < 0 Then nRollRecord = 0
            For nRollRecord = nRollRecord To 0 Step -1
                If .Num(1, nRollRecord) <= dDate Then Exit For
            Next
            For nRollRecord = nRollRecord To .NumRecords - 1
                If nRollRecord = .NumRecords - 1 Then Exit For
                If .Num(1, nRollRecord + 1) > dDate Then Exit For
            Next
            nSymbolID = .Num(0, nRollRecord)
            If nPrevSymbolID <> nSymbolID Then
                nPrevSymbolID = nSymbolID
                strSymbol = SU_GetSymbol(nSymbolID)
            End If
        End If
    End With
    
    RollSymbol = strSymbol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.RollSymbol", eGDRaiseError_Raise

End Function

Private Function GetIndicatorDataLabel(ByVal nBar&, ByVal idx&, ByVal idxPane&, ByVal eLabelType As eDataLabelType) As String
On Error Resume Next

    Dim i&, hHandle&, dNum#, dNum2#, dDate#, dPrice#, dDelta#
    Dim s$, strText$, strValue$, strFmt$, strDelta$
    Dim bTradingUnits As Boolean, bAllFour As Boolean, bHA As Boolean
    Dim Ind As cIndicator, Pane As cPane
            
    ' build text to display
    If idxPane = 0 Then
        idxPane = m.Tree.RelativeIndex(idx, eTREE_Root)
    End If
    If idxPane > 0 And idxPane <= m.Tree.Count And idx > 0 And idx <= m.Tree.Count Then
        Set Pane = m.Tree(idxPane)
        If Pane.DisplayFormat = ePANE_PriceFormat Then
            bTradingUnits = True
        ElseIf Pane.DisplayFormat = ePANE_RoundDecimals Then
            If Pane.DisplayDecimals > 0 Then
                strFmt = "0." & String(Pane.DisplayDecimals, "0")
            Else
                strFmt = "0"
            End If
        End If
        Set Ind = m.Tree(idx)
        
        'only build text for current cycle and trends if this is a seasonal chart
        If TypeOfChart = eTypeChart_Seasonal Then
            If Ind.MyKey <> kSeasonalAvgIndKey And Ind.MyKey <> kSeasonalCurrIndKey And _
               Ind.MyKey <> kSeasonalBullIndKey And Ind.MyKey <> kSeasonalBearIndKey Then
               
                GoTo ErrExit
            End If
        End If
        
        If Ind.Display Then
            'If Pane.DisplayFormat = ePANE_RoundDecimals Then
            '    strFmt = strDecFmt
            'End If
            Select Case Ind.DataType
                Case eINDIC_BarData
                    hHandle = Ind.Bars.BarsHandle
                    If eLabelType = eLBL_Chart Then
                        bAllFour = False
                    ElseIf Ind.DisplayType < 0 Or m.Tree.Key(idx) = "PRICE" Then
                        bAllFour = True
                    End If
                    
                    ' special case for main chart price label
                    If eLabelType = eLBL_Chart And m.Tree.Key(idx) = "PRICE" Then
                        dPrice = m.Bars(eBARS_Close, nBar)
                        ' look for last bar of a previous date
                        dDelta = kNullData
                        dNum = m.Bars.SessionDate(nBar)
                        If m.Settles.Size > 0 Then
                            i = m.Settles.FindDateTime(dNum)
                            dDelta = m.Settles(eBARS_Close, i - 1)
                        Else
                            For i = nBar - 1 To 0 Step -1
                                If m.Bars.SessionDate(i) <> dNum Then
                                    dDelta = m.Bars(eBARS_Close, i)
                                    Exit For
                                End If
                            Next
                        End If
                        If dDelta <> kNullData Then
                            dDelta = dPrice - dDelta
                        End If
                        If Pane.DisplayFormat = ePANE_RoundDecimals Then
                            strValue = Format(dPrice, strFmt)
                            strDelta = Format(dDelta, strFmt)
                        Else
                            strValue = Ind.Bars.PriceDisplay(dPrice, bTradingUnits)
                            strDelta = Ind.Bars.PriceDisplay(dDelta, bTradingUnits)
                        End If
                        
                        dDate = 0
                        If nBar >= LastGoodDataBar(False) Then
                            ' see if actual last date is stored (in case weekly, etc.)
                            dDate = m.dLastActualDateOfData
                        End If
                        If dDate <= 0 Then
                            dDate = m.Bars(eBARS_DateTime, nBar)
                        End If
                        If IsIntraday(m.nPeriodicity) Then
                            ' show date/time of bar
                            dDate = DateTimeConvert(m.Bars, dDate)
                            strText = DateFormat(dDate, M_D_YY) & Format(dDate, " HH:MM")
                        ElseIf nBar >= LastGoodDataBar(False) And m.Bars.Prop(eBARS_PriceHasSettled) = 0 _
                                    And m.Bars.Prop(eBARS_LastTickTime) > 0 Then
                            ' for daily charts, show last tick date/time if a partial session on last bar
                            ' (i.e. if on last bar and has a tick time and has not "settled")
                            dDate = m.Bars.SessionDate(nBar) + m.Bars.Prop(eBARS_LastTickTime) / 1440#
                            If m.Bars.Prop(eBARS_LastTickTime) > m.Bars.Prop(eBARS_CrossoverTime) Then
                                dDate = dDate - 1 ' date/time is evening prior to session date
                            End If
                            If g.bShowInLocalTimeZone Then
                                dDate = ConvertTimeZone(dDate, m.Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                            End If
                            strText = DateFormat(dDate, M_D_YY) & Format(dDate, " HH:MM")
                        Else
                            ' show session date
                            dDate = m.Bars.SessionDate(nBar)
                            strText = DateFormat(dDate, MM_DD_YYYY)
                        End If
                        
                        strText = strText & " = " & strValue
                        If dDelta > 0 Then
                            strText = strText & " (+" & strDelta & ")"
                        ElseIf dDelta <> kNullData Then
                            strText = strText & " (" & strDelta & ")"
                        End If
                        If Ind.Overlayed Then strText = "*" & strText
                        Ind.ChartValue = strText
                    
                    ElseIf bAllFour Then
                        If UCase(Ind.Bars.Prop(eBARS_CustomString)) = "HEIKIN-ASHI" Then
                            bHA = True
                        Else
                            bHA = False
                        End If
                    
                        If eLabelType = eLBL_DataWindow Then
                            If bHA Then
                                strText = strText & "|H.A. Open" & vbTab
                            Else
                                strText = strText & "|Open" & vbTab
                            End If
                        End If
                        If nBar >= 0 Then
                            dNum = gdGetNum(gdBarsArray(hHandle, eBARS_Open), nBar)
                            If dNum <> kNullData Then
                                If Pane.DisplayFormat = ePANE_RoundDecimals Then
                                    strValue = Format(dNum, strFmt)
                                Else
                                    strValue = Ind.Bars.PriceDisplay(dNum, bTradingUnits)
                                End If
                                If eLabelType <> eLBL_Mouse Then
                                    strText = strText & strValue
                                ElseIf bHA Then
                                    strText = strText & "haO=" & strValue & "  "
                                Else
                                    strText = strText & "O=" & strValue & "  "
                                End If
                            End If
                        End If
                        
                        If eLabelType = eLBL_DataWindow Then
                            If bHA Then
                                strText = strText & "|H.A. High" & vbTab
                            Else
                                strText = strText & "|High" & vbTab
                            End If
                        End If
                        If nBar >= 0 Then
                            dNum = gdGetNum(gdBarsArray(hHandle, eBARS_High), nBar)
                            If dNum <> kNullData Then
                                If Pane.DisplayFormat = ePANE_RoundDecimals Then
                                    strValue = Format(dNum, strFmt)
                                Else
                                    strValue = Ind.Bars.PriceDisplay(dNum, bTradingUnits)
                                End If
                                If eLabelType <> eLBL_Mouse Then
                                    strText = strText & strValue
                                ElseIf bHA Then
                                    strText = strText & "haH=" & strValue & "  "
                                Else
                                    strText = strText & "H=" & strValue & "  "
                                End If
                            End If
                        End If
                        
                        If eLabelType = eLBL_DataWindow Then
                            If bHA Then
                                strText = strText & "|H.A. Low" & vbTab
                            Else
                                strText = strText & "|Low" & vbTab
                            End If
                        End If
                        If nBar >= 0 Then
                            dNum = gdGetNum(gdBarsArray(hHandle, eBARS_Low), nBar)
                            If dNum <> kNullData Then
                                If Pane.DisplayFormat = ePANE_RoundDecimals Then
                                    strValue = Format(dNum, strFmt)
                                Else
                                    strValue = Ind.Bars.PriceDisplay(dNum, bTradingUnits)
                                End If
                                If eLabelType <> eLBL_Mouse Then
                                    strText = strText & strValue
                                ElseIf bHA Then
                                    strText = strText & "haL=" & strValue & "  "
                                Else
                                    strText = strText & "L=" & strValue & "  "
                                End If
                            End If
                        End If
                        
                        If eLabelType = eLBL_DataWindow Then
                            If bHA Then
                                strText = strText & "|H.A. Close" & vbTab
                            Else
                                strText = strText & "|Close" & vbTab
                            End If
                        End If
                        If nBar >= 0 Then
                            dNum = gdGetNum(gdBarsArray(hHandle, eBARS_Close), nBar)
                            If dNum <> kNullData Then
                                If Pane.DisplayFormat = ePANE_RoundDecimals Then
                                    strValue = Format(dNum, strFmt)
                                Else
                                    strValue = Ind.Bars.PriceDisplay(dNum, bTradingUnits)
                                End If
                                If eLabelType <> eLBL_Mouse Then
                                    strText = strText & strValue
                                ElseIf bHA Then
                                    strText = strText & "haC=" & strValue & "  "
                                Else
                                    strText = strText & "C=" & strValue & "  "
                                End If
                            End If
                        End If
                    Else
                        ' comparison symbol
                        strValue = ""
                        If nBar >= 0 Then
                            dNum = gdGetNum(gdBarsArray(hHandle, eBARS_Close), nBar)
                            If dNum <> kNullData Then
                                If Pane.DisplayFormat = ePANE_RoundDecimals Then
                                    strValue = Format(dNum, strFmt)
                                Else
                                    strValue = Ind.Bars.PriceDisplay(dNum, bTradingUnits)
                                End If
                            End If
                        End If
                        Select Case eLabelType
                            Case eLBL_DataWindow
                                strText = "|" & Ind.Name & vbTab & strValue
                            Case eLBL_Mouse
                                If Len(strValue) > 0 Then
                                    strText = Ind.Name & "= " & strValue & "  "
                                End If
                            Case eLBL_Chart
                                Ind.ChartValue = strValue
                                If Len(strValue) > 0 Then
                                    strText = Ind.ChartLabel & " = " & strValue
                                Else
                                    strText = Ind.ChartLabel
                                End If
                        End Select
                    End If
                    
                Case eINDIC_Array, eINDIC_DrawCommands
                    strValue = ""
                    If nBar >= 0 And Not Ind.Data.IsConstantValue Then
                        If UCase(Ind.Name) = "TAS SIDEWAYS" Then
                            dNum = kNullData
                        Else
                            dNum = gdGetNum(Ind.Data.ArrayHandle, nBar)
                        End If
                        If dNum <> kNullData Then
                            If Ind.CodedName = "SeasonalTrend" Then
                                dNum2 = gdGetNum(Ind.Data.ArrayHandle, nBar - 1)
                                If dNum2 <> kNullData Then
                                    dNum = dNum - dNum2
                                    If dNum = Int(dNum) Then
                                        strFmt = "0"
                                    ElseIf Abs(dNum) > 10 Then
                                        strFmt = "0.00"
                                    ElseIf Abs(dNum) > 0.01 Then
                                        strFmt = "0.0000"
                                    Else
                                        strFmt = "0.000000"
                                    End If
                                    strValue = Format(Abs(dNum), strFmt)
                                    If dNum > 0 Then
                                        strValue = strValue & " UP     "
                                    Else
                                        strValue = strValue & " DOWN"
                                    End If
                                End If
                            ElseIf Ind.CodedName = "CycleTrend" Then
                                dNum2 = gdGetNum(Ind.Data.ArrayHandle, nBar - 1)
                                If dNum2 > 0 Then
                                    dNum = (dNum / dNum2 - 1) * 100#
                                    strFmt = "0.00"
                                    strValue = Format(dNum, strFmt) & "% avg price change"
                                End If
                            ElseIf Left(Ind.CodedName, 12) = "CyclePercent" Then
                                If dNum < 50 Then
                                    strValue = Format(100 - dNum, "0") & "% traded DOWN"
                                Else
                                    strValue = Format(dNum, "0") & "% traded UP"
                                End If
                            ElseIf Ind.CodedName = "RedLightGreenLight" Or Ind.CodedName = "SeasonalPercent" Then
                                If dNum < 50 Then
                                    strValue = Format(100 - dNum, "0") & "% traded DOWN"
                                Else
                                    strValue = Format(dNum, "0") & "% traded UP"
                                End If
                                strValue = strValue & " (for all years)"
                            ElseIf bTradingUnits And Ind.idxPriceBars >= 0 And TypeOfChart <> eTypeChart_Seasonal Then
                                strValue = PriceDisplay(idx, dNum) '?? , True)
                            ElseIf Pane.DisplayFormat = ePANE_RoundDecimals And Ind.Overlayed = False Then
                                strValue = Format(dNum, strFmt)
                            Else
                                If dNum = Int(dNum) Then
                                    strFmt = "0"
                                ElseIf Abs(dNum) > 10 Then
                                    strFmt = "0.00"
                                ElseIf Abs(dNum) > 0.01 Then
                                    strFmt = "0.0000"
                                Else
                                    strFmt = "0.000000"
                                End If
                                strValue = Format(dNum, strFmt)
                            End If
                            strValue = Trim(strValue)
                            If TypeOfChart = eTypeChart_Seasonal Then
                                strValue = strValue & "%"
                            End If
                        End If
                    End If
                    Select Case eLabelType
                        Case eLBL_DataWindow
                            strText = "|" & Ind.ChartLabel & vbTab & strValue
                        Case eLBL_Mouse
                            If Len(strValue) > 0 Then
                                If Ind.DisplayType = eINDIC_ClusterPrice Or Ind.DisplayType = eINDIC_ClusterPrice Then
                                    strText = "^Price Clusters " & "= " & strValue & "  "
                                ElseIf Ind.DisplayType = eINDIC_ClusterTime Then
                                    strText = "Time Clusters " & "= " & strValue & "  "
                                ElseIf Ind.MyKey = kSeasonalCurrIndKey Then
                                    strText = "Current Cycle= " & strValue & "  "
                                Else
                                    strText = Ind.Name & "= " & strValue & "  "
                                End If
                            End If
                        Case eLBL_Chart
                            If Len(strValue) > 0 Then
                                strText = Ind.ChartLabel & " = " & strValue
                            Else
                                strText = Ind.ChartLabel
                            End If
                    End Select
                
                Case eINDIC_BooleanArray '(highlight bars)
                    strValue = ""
                    If nBar >= 0 And Not Ind.Data.IsConstantValue Then
                        dNum = gdGetNum(Ind.Data.ArrayHandle, nBar)
                        If dNum = 0 Then
                            strValue = "False"
                        ElseIf dNum <> kNullData Then
                            strValue = "True"
                        End If
                    End If
                    Select Case eLabelType
                        Case eLBL_DataWindow
                            strText = "|" & Ind.ChartLabel & vbTab & strValue
                        Case eLBL_Mouse
                            If Len(strValue) > 0 Then
                                If UCase(strValue) = "TRUE" Then
                                    strText = "<" & Ind.Name & ">  "
                                End If
                                'strText = Ind.Name & "= " & strValue & "  "
                            End If
                        Case eLBL_Chart
                            If Len(strValue) > 0 Then
                                strText = Ind.ChartLabel & " = " & strValue
                            Else
                                strText = Ind.ChartLabel
                            End If
                    End Select
                    
            End Select
        End If
    End If
        
ErrExit:
    Set Ind = Nothing
    Set Pane = Nothing
    GetIndicatorDataLabel = strText

End Function

Private Function GetPaneDataLabel(ByVal nX&, ByVal idxPane&, ByVal eLabelType As eDataLabelType) As String
On Error Resume Next

    Dim i&, nBar&, idx&, idxLast&, hHandle&, dNum#
    Dim strText$, strValue$
            
    ' build text to display
    nBar = m.aXBar(nX) 'could be < 0
    If idxPane > 0 And idxPane <= m.Tree.Count Then
        If m.Tree.NodeLevel(idxPane) = 0 Then
            idxLast = m.Tree.RelativeIndex(idxPane, eTREE_LastDescendant)
            ' for each indicator in this pane
            For idx = idxPane + 1 To idxLast
                strValue = GetIndicatorDataLabel(nBar, idx, idxPane, eLabelType)
                strText = strText & strValue
            Next
            
            If eLabelType = eLBL_Mouse Then
                If UCase(m.Tree.Key(idxPane)) = "PRICE PANE" Then
                    If TypeOfChart = eTypeChart_Seasonal Then
                        If nBar >= 0 Then strText = strText & "CycleBar#=" & CStr(nBar + 1) & "  "
                    Else
                        dNum = m.aTDOM(nBar)
                        If dNum > 0 Then
                            strText = strText & "TDofM=" & CStr(dNum) & "  "
                        End If
                        dNum = m.aTDOY(nBar)
                        If dNum > 0 Then
                            If m.nPeriodicity >= ePRD_Weeks Then
                                strText = strText & "TWofY=" & CStr(dNum) & "  "
                            Else
                                strText = strText & "TDofY=" & CStr(dNum) & "  "
                            End If
                        End If
                    End If
                End If
            End If
        End If
    ElseIf eLabelType = eLBL_Mouse Then
        dNum = m.aTDOM(nBar)
        If dNum > 0 Then
            strText = strText & "TradeDayOfMonth= " & CStr(dNum) & "  "
        End If
        dNum = m.aTDOY(nBar)
        If dNum > 0 Then
            If m.nPeriodicity >= ePRD_Weeks Then
                strText = strText & "TradeWeekOfYear= " & CStr(dNum) & "  "
            Else
                strText = strText & "TradeDayOfYear= " & CStr(dNum) & "  "
            End If
        End If
    Else
        strText = strText & "|TradeDayOfMonth" & vbTab
        dNum = m.aTDOM(nBar)
        If dNum > 0 Then strText = strText & CStr(dNum)
        dNum = m.aTDOY(nBar)
        If m.nPeriodicity >= ePRD_Weeks Then
            strText = strText & "|TradeWeekOfYear" & vbTab
            If dNum > 0 Then strText = strText & CStr(dNum)
        Else
            strText = strText & "|TradeDayOfYear" & vbTab
            If dNum > 0 Then strText = strText & CStr(dNum)
        End If
    End If

    strText = Trim(strText)
    If Right(strText, 1) = "," Then
        strText = Trim(Left(strText, Len(strText) - 1))
    End If
        
    GetPaneDataLabel = strText

End Function

Public Function GetMouseLabel(ByVal nX&, ByVal nPaneID&) As String
On Error GoTo ErrSection:

    Dim i&, dDate#, nBar&, strText$, strRoll$
    
    nBar = m.aXBar(nX) 'could be < 0
    dDate = m.aXdate(nX)
        
    ' build text to display
    If dDate > 0 Then
        strRoll = RollSymbol(nX)
        If Len(strRoll) > 0 Then
            strText = " (" & strRoll & ")"
        End If
        dDate = DateTimeConvert(m.Bars, dDate)
        If GetPeriodType(m.nPeriodicity) = ePRD_Minutes Or GetPeriodType(m.nPeriodicity) = ePRD_SMP Then
            strText = DateFormat(dDate, MM_DD_YY) & Format(dDate, " Hh:Nn") & strText & "  "
        ElseIf IsIntraday(m.nPeriodicity) Then
If NewFullTick Then
            strText = DateFormat(dDate, MM_DD_YY) & Format(dDate, " Hh:Nn:Ss") & strText & "  "
Else
            strText = DateFormat(dDate, MM_DD_YY) & Format(dDate, " Hh:Nn") & strText & "  "
End If
        Else
            strText = DateFormat(dDate, MM_DD_YYYY) & strText & ":  "
        End If
        strText = WeekdayName(Weekday(dDate), True) & vbTab & strText
    Else
        strText = vbTab
    End If
    
    ' get text for pane
    strText = strText & GetPaneDataLabel(nX, nPaneID, eLBL_Mouse)

    GetMouseLabel = Trim(strText)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.GetMouseLabel", eGDRaiseError_Raise

End Function

Public Function GetDataWindowLabel(nX As Long) As String
On Error GoTo ErrSection:

    Dim i&, dDate#, dLocal#, nSession&, nBar&, strText$, strRoll$, strTZ$
    Dim Pane As cPane
    
    If nX < 0 Then
        nX = LastGoodDataBar(True, True)
    End If
    nBar = m.aXBar(nX) 'could be < 0
    dDate = m.aXdate(nX)
        
    ' build text to display
    strText = "Symbol" & vbTab & Symbol
    If dDate > 0 Then
        If Not IsIntraday(m.nPeriodicity) Then
            ' just show date for non-intraday
            strText = strText & "|Date" & vbTab & DateFormat(dDate, MM_DD_YYYY)
        Else
            ' for intraday, show Session date, TradeSense time, and Local time
            nSession = Int(dDate)
            i = m.Bars.Prop(eBARS_CrossoverTime)
            If i > 0 And gdFixDateTime(dDate) >= gdFixDateTime(nSession + i / 1440#) Then
                nSession = nSession + 1 ' is tomorrow if after crossover time
            End If
            strTZ = m.Bars.Prop(eBARS_ExchangeTimeZoneInf)
            dLocal = ConvertTimeZone(dDate, strTZ, "")
            Select Case UCase(strTZ)
            Case "NY", ""
                strTZ = "|NY Time"
            Case "GMT", "UTC"
                strTZ = "|GMT"
            Case Else
                strTZ = "|Exch Time"
            End Select
            strText = strText & "|Session" & vbTab & DateFormat(nSession, MM_DD_YYYY) _
                & strTZ & vbTab & DateFormat(dDate, M_D) & Format(dDate, " Hh:Mm") _
                & "|Local Time" & vbTab & DateFormat(dLocal, M_D) & Format(dLocal, " Hh:Mm")
        End If
        strRoll = RollSymbol(nX)
        If Len(strRoll) > 0 Then
            strText = strText & "|Contract" & vbTab & strRoll
        ElseIf m.RollsTable.NumRecords > 0 Then
            strText = strText & "|Contract" & vbTab
        End If
    End If
    
    ' append text for all panes
    For i = 1 To m.Tree.Count
        If m.Tree.NodeLevel(i) = 0 Then
            Set Pane = m.Tree(i)
            If Pane.Display Then
                strText = strText & GetPaneDataLabel(nX, i, eLBL_DataWindow)
            End If
        End If
    Next
    Set Pane = Nothing
    
    ' append TradeDay stuff
    strText = strText & GetPaneDataLabel(nX, -1, eLBL_DataWindow)
    
    GetDataWindowLabel = Trim(strText)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.GetDataWindowLabel", eGDRaiseError_Raise

End Function

Public Property Get ExtraPriceScale() As Double
    ExtraPriceScale = m.dExtraPriceScale
End Property

Public Property Let ExtraPriceScale(ByVal dNum As Double)
    If dNum < 0 Then dNum = 0
    m.dExtraPriceScale = dNum
End Property

Public Property Get TemplateApplied() As String
    TemplateApplied = m.strTemplateApplied
End Property

Public Property Let TemplateApplied(ByVal strApplied As String)
    m.strTemplateApplied = strApplied
    SetFormCaption
End Property

Private Function TemplateIsDirty() As Boolean
On Error GoTo ErrSection:

    Dim bDirty As Boolean, bManualScale As Boolean
    Dim strFile$, iCurr&, iOrig&
    Dim strOrig$, strCurr$, strProp$, strIgnore$
    Dim aCurrent As New cGdArray
    Dim aOriginal As New cGdArray
    
    ' read in the current template
    strCurr = Template
    If Len(strCurr) = 0 Then Exit Function
    strCurr = g.ChartGlobals.strCPCRoot & "\Charts\" & strCurr & ".CHT"
    aCurrent.FromFile strCurr
    
    ' read in the original template (when applied)
    strOrig = TemplateApplied
    If Len(strOrig) = 0 Then Exit Function
    strOrig = g.ChartGlobals.strCPCRoot & "\Charts\Templates\" & strOrig & ".CHT"
    aOriginal.FromFile strOrig
    
    ' list of properties to ignore for the purpose of determining if "dirty" or not
    strIgnore = UCase("|Desc|ChartName|TemplateApplied|SymbolID|Required|ChartLabel|Size|DateApplied|UserScreenDate|AutoBreakout|RemoveOvernightGap|ShowToolbar|")
    
    ' compare each line of the original with the current
    ' (allow current to have additional newer properties,
    ' but all existing properties of original must match,
    ' as well as the number of indicators and panes)
    iCurr = -1
    For iOrig = 0 To aOriginal.Size - 1
        strOrig = UCase(aOriginal(iOrig))
        strProp = UCase(Trim(Parse(strOrig, "=", 1)))
        
        ' TLB 10/7/03: need to ignore Max and Min value unless manual scale
        If Left(strOrig, 1) = "[" Then
            bManualScale = False
        ElseIf strOrig = "SCALING=1" Then
            bManualScale = True
        ElseIf strProp = "MAXVALUE" Or strProp = "MINVALUE" Then
            If Not bManualScale Then
                strProp = "DESC" '(just so will be ignored)
            End If
        End If
        
        If InStr(strIgnore, "|" & strProp & "|") = 0 Then
            Do While True
                iCurr = iCurr + 1
                If iCurr >= aCurrent.Size Then
                    Exit For '(dirty)
                End If
                strCurr = UCase(aCurrent(iCurr))
                strProp = UCase(Trim(Parse(strCurr, "=", 1)))
                If InStr(strIgnore, "|" & strProp & "|") = 0 Then
                    If strCurr = strOrig Then
                        Exit Do
                    ElseIf Left(strCurr, 1) = "[" Then
                        'at next indicator, so we must be dirty
                        iCurr = aCurrent.Size
                        Exit For
                    End If
                End If
            Loop
        End If
    Next
    ' if we didn't end up looking at the last line of the
    ' current, then we must be "dirty"
    If iCurr <> aCurrent.Size - 1 Then
        bDirty = True
    End If

    TemplateIsDirty = bDirty

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.TemplateIsDirty", eGDRaiseError_Raise

End Function

Public Property Get UseCustomColors() As Boolean
    UseCustomColors = m.bCustomColors
End Property

Public Property Let UseCustomColors(ByVal bCustom As Boolean)
    m.bCustomColors = bCustom
End Property

Public Property Get ChartForeColor() As Long
    ChartForeColor = m.nChartForeColor
End Property

Public Property Let ChartForeColor(ByVal nColor As Long)
    m.nChartForeColor = nColor
End Property

Public Property Get ChartBackColor() As Long
    ChartBackColor = m.nChartBackColor
End Property

Public Property Let ChartBackColor(ByVal nColor As Long)
    m.nChartBackColor = nColor
End Property

Public Property Get BorderForeColor() As Long
    BorderForeColor = m.nBorderForeColor
End Property

Public Property Let BorderForeColor(ByVal nColor As Long)
    m.nBorderForeColor = nColor
End Property

Public Property Get BorderBackColor() As Long
    BorderBackColor = m.nBorderBackColor
End Property

Public Property Let BorderBackColor(ByVal nColor As Long)
    m.nBorderBackColor = nColor
End Property

Public Property Get ChartGradientColor() As Long
    ChartGradientColor = m.nChartGradientColor
End Property

Public Property Let ChartGradientColor(ByVal nColor As Long)
    m.nChartGradientColor = nColor
End Property

Public Property Get UseGradient() As Long
    UseGradient = m.nUseGradient
End Property

Public Property Let UseGradient(ByVal nUse As Long)
    If nUse = 0 Or nUse = 1 Then m.nUseGradient = nUse
End Property

Public Function VerifyPanes(ByVal bShowErrorMsg As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim i&, j&, iPanesDisplayed&, iIndDisplayed&
    Dim strText$, bNonOverlayed As Boolean
    Dim bSpecialType As Boolean, bSpecialTypesMsg As Boolean
    Dim Ind As cIndicator
    Dim Pane As cPane
    Dim bSuccess As Boolean
    
    'first make sure things look OK
    bSuccess = True
    
    If TypeOfChart = eTypeChart_Seasonal Then GoTo ErrExit          '6398
    
    With m.Tree
        iPanesDisplayed = 0
        iIndDisplayed = 0
        bNonOverlayed = False
        bSpecialType = False
        bSpecialTypesMsg = False
        For i = .Count To 1 Step -1 '(easier to go backwards)
            If .NodeLevel(i) = 0 Then
                Set Pane = .Item(i)
                'if no indicator displayed in pane,
                'turn pane display off
                If iIndDisplayed = 0 Then
                    Pane.Display = False
                End If
                'count # of panes being displayed
                'need to check indicator displayed count as well because
                'price pane display flag is never turned off
                If Pane.Display And iIndDisplayed > 0 Then
                    iPanesDisplayed = iPanesDisplayed + 1
                    'if no non-overlayed indicator on an
                    'auto-scaled pane, set first indicator
                    If Not bNonOverlayed And Pane.Scaling = ePANE_ScaleModeAuto Then
                        Set Ind = .Item(iIndDisplayed)
                        If bShowErrorMsg Then
                            'InfBox "i=i ; h=WARNING ; " & UCase(Ind.Name) & "|is no longer set to 'Overlay' " _
                            '    & "|(at least one indicator in each pane must|be fixed to the pane's vertical scale)."
                            strText = "At least one indicator of an auto-scaled pane|must be fixed to the pane's vertical scale." _
                                & "||Do you wish to change the 'Overlay' setting for:|" & UCase(Ind.Name)
                            If AskBox("i=? ; h=Change Overlay? ; b=+Yes|-Cancel ; " & strText) = "C" Then
                                bSuccess = False
                                Exit For
                            End If
                        End If
                        Ind.Overlayed = 0
                    End If
                End If
                'reset for next pane
                iIndDisplayed = 0
                bNonOverlayed = False
                bSpecialType = False
            Else
                Set Ind = .Item(i)
                'keep track of first indicator being displayed in pane
                If Ind.Display And Ind.DataType <> eINDIC_Constant Then
                    iIndDisplayed = i
                    'to make sure at least one non-overlayed indicator in pane
                    If Ind.Overlayed = 0 Then
                        bNonOverlayed = True
                    End If
                    'to see if more than 1 special type
                    If Ind.DisplayType < 0 Or Ind.DisplayType = eINDIC_Area _
                            Or Ind.DisplayType = eINDIC_Histogram Then
                        If bSpecialType Then
                            ' with new graphics engine, we can now do more than
                            ' one of these on each pane!
                            'bSpecialTypesMsg = True 'too many
                        Else
                            bSpecialType = True
                        End If
                    End If
                End If
            End If
        Next
        If bSuccess Then
            'check # of panes displayed
            If iPanesDisplayed = 0 Then
                If bShowErrorMsg Then
                    InfBox "i=[] ; h=ERROR ; At least one Pane must be displayed."
                    bSuccess = False
                ElseIf .Count > 0 Then
                    'show price pane (or first pane if no price pane)
                    Set Ind = Nothing
                    Set Pane = .Item("PRICE PANE")
                    If Not Pane Is Nothing Then
                        i = .Index("PRICE")
                        If i > 0 Then Set Ind = .Item(i)
                    Else '(this shouldn't really happen -- there should be a Price Pane)
                        Set Pane = .Item(1)
                        i = .RelativeIndex(1, eTREE_FirstChild)
                        If i > 0 Then Set Ind = .Item(i)
                    End If
                    If Not Pane Is Nothing Then
                        Pane.Display = True
                        If Not Ind Is Nothing Then
                            Ind.Display = True
                        End If
                    End If
                End If
            'ElseIf iPanesDisplayed > 9 Then
                'If bShowErrorMsg Then
                '    InfBox "i=! ; h=WARNING ; Only up to 9 Panes can be shown at a time."
                'End If
            End If
        End If
        If bSuccess Then
            If Not bShowErrorMsg Then
                j = .Index(kClusterTimeKeyPane)
                If j > 0 Then
                    'make sure time cluster pane is always immediately below price pane
                    i = .RelativeIndex("PRICE PANE", eTREE_NextSibling)
                    If j <> i Then
                        .Move j, "PRICE PANE", eTREE_NextSibling
                        geResetPanes
                    End If
                End If
            ElseIf bSpecialTypesMsg And bShowErrorMsg Then
                'too many special types?
                strText = "A pane cannot have more than one indicator as a 'special display type' (OHLC bars, Histogram or Area chart)" _
                    & " -- the other indicators for that pane will be shown as lines regardless of their specified display type.|"
                If AskBox("i=! ; h=WARNING ; b=+OK|-Cancel ; " & strText) = "C" Then
                    bSuccess = False
                End If
            End If
        End If
    End With
    Set Ind = Nothing
    Set Pane = Nothing


ErrExit:
    VerifyPanes = bSuccess
    Exit Function
    
ErrSection:
    RaiseError "cChart.VerifyPanes", eGDRaiseError_Raise

End Function

Public Sub ImageServerCheck()
On Error Resume Next

    Dim i&, nSymbolID&, dStartTime#, nPixels&
    Dim dDate#, nCrossover&, dLast#, dPrev#, dOpen#, dHigh#, dLow#, dVol#
    Dim strImgPath$, strFile$, strID$, strText$, strRexQ$
    Dim strSymbol$, strCompare$, strErr$, strPeriod$, strPNG$, strTemplate$
    Dim aStrings As New cGdArray, aSpeed As New cGdArray
    Dim Bars As New cGdBars
    Dim bTime As Boolean, bUseRex As Boolean

    ' look for a trigger file
    For i = 1 To 99
        strImgPath = Parse(frmImageServer.ImgSrvQ, ";", i)
        If Len(strImgPath) = 0 Then Exit For
        strImgPath = AddSlash(strImgPath)
        strFile = Dir(strImgPath & "*.GO")
        If Len(strFile) > 0 Then
            ' try to rename file (so we own it)
            strID = FileBase(strFile)
            If RenameFile(strImgPath & strID & ".GO", _
                        strImgPath & strID & ".PRO") Then
                ' we now own this file to work with
                KillFile strImgPath & strID & ".DON"
                KillFile strImgPath & strID & ".PNG"
                Exit For
            End If
            strID = "" ' (another image server must have gotten it)
        End If
    Next
    
    If Len(strID) = 0 Then Exit Sub '(no file found)
    
    dStartTime = gdTickCount(False)
    bTime = frmTest.Visible
    If bTime Then frmTest.AddList "Starting image for: " & strID
    
    ' read trigger file: symbols, periodicity
    m.strRexDataFile = ""
    m.nImgSrvFontSize = 0
    m.iMiniChart = 0
    aStrings.FromFile strImgPath & strID & ".PRO"
    For i = 0 To aStrings.Size - 1
        strText = UCase(Parse(aStrings(i), "=", 2))
        Select Case UCase(Parse(aStrings(i), "=", 1))
        Case "SYMBOL"
            strSymbol = strText
        Case "COMPARESYMBOLS"
            strCompare = strText
        Case "BARPERIOD"
            strPeriod = strText
        Case "PIXELWIDTH"
            nPixels = Val(strText)
        Case "FONTSIZE"
            m.nImgSrvFontSize = Val(strText)
        Case "MINICHART"
            m.iMiniChart = Val(strText)
        Case "USEREX"
            If Left(strText, 1) = "Y" Or Left(strText, 1) = "T" Then
                bUseRex = True
            End If
        Case "TEMPLATE"
            strTemplate = FileBase(strText)
        End Select
    Next
    If Len(strPeriod) = 0 Then
        strPeriod = "D" 'default: Daily
    End If
    If nPixels = 0 Then
        nPixels = frmImageServer.PixelWidth
        If nPixels = 0 Then nPixels = 480 '(default)
    End If
    If m.nImgSrvFontSize = 0 Then
        m.nImgSrvFontSize = frmImageServer.DefaultFontSize
        If m.nImgSrvFontSize = 0 Then m.nImgSrvFontSize = 150
    End If
        
    strRexQ = AddSlash(frmImageServer.RexQ)
    If Len(strRexQ) = 0 Then
        bUseRex = False
    End If
        
    ' run symbol through bars to convert to new Genesis type
    Bars.Prop(eBARS_PeriodicityStr) = strPeriod
    Bars.Prop(eBARS_Symbol) = strSymbol
    strSymbol = Bars.Prop(eBARS_Symbol)
    nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
    If nSymbolID = 0 Then
        strErr = "Symbol not found: " & strSymbol
    ElseIf bUseRex Then
        ' create the REX request: date-date;15;I;S;INTC
        ' (do this now so will be processing while we load chart)
        If IsIntraday(Bars.Prop(eBARS_Periodicity)) Then
            strPeriod = "1" 'ask for 1-minute bars for intraday
        Else
            strPeriod = "D" 'ask for today's bar for non-intraday
        End If
        aStrings.Clear
        aStrings.Add "OPTIONS:LAST_TICK"
        ' request data for primary symbol
        strText = Str(JulToLong(LastDailyDownload + 1, True)) _
            & "-" & Str(JulToLong(Date + 1, True)) & ";" _
            & strPeriod & ";I;" & Bars.SecurityType & ";" & strSymbol
        aStrings.Add strText
        ' request data for comparison symbols (comma-delimited list)
        For i = 1 To 99
            strText = Parse(strCompare, ",", i)
            If Len(strText) = 0 Then Exit For
            ' run symbol through bars to convert to new Genesis type
            Bars.Prop(eBARS_Symbol) = strText
            strText = Bars.Prop(eBARS_Symbol)
            If g.SymbolPool.SymbolIDforSymbol(strText) > 0 Then
                strText = Str(JulToLong(LastDailyDownload + 1, True)) _
                    & "-" & Str(JulToLong(Date + 1, True)) & ";" _
                    & strPeriod & ";I;" & Bars.SecurityType & ";" & strText
                aStrings.Add strText
            End If
        Next
        KillFile strRexQ & strID & ".*"
        aStrings.ToFile strRexQ & strID & ".GO"
        If FileExist(strRexQ & strID & ".*") Then
            m.strRexDataFile = strRexQ & strID
        Else
            strErr = "Error creating data request."
        End If
        If bTime Then frmTest.AddList "Created REX request for  " & strID, True
    End If
    
    If Len(strErr) = 0 Then
        ' set this chart busy and designate another searcher
        ' (will look for another free chart to be searching
        ' while this one processes the data and chart image)
        m.Form.ImgSrvState = eImgSrv_Busy
        SetImgSrvSearcher
        If bTime Then frmTest.AddList "Reset searcher", True
        
        ' assign new symbol
        SetSymbol nSymbolID
        
        ' load chart template
        If Len(strTemplate) = 0 Then
            strTemplate = strID
        End If
        strFile = App.Path & "\Charts\" & strID & ".CHT"
        strPNG = ReplaceFileExt(strFile, ".PNG")
        KillFile strPNG
        If FileExist(strImgPath & strTemplate & ".CHT") Then
            ' if .CHT template exists in the ImgPath then copy it from there
            FileCopy strImgPath & strTemplate & ".CHT", strFile
        Else
            ' else copy it from our own Templates folder
            FileCopy App.Path & "\Charts\Templates\" & strTemplate & ".CHT", strFile
        End If
        If Not TemplateLoad(strID) Then
            strErr = "Could not load chart template."
        Else
            If bTime Then frmTest.AddList "Template loaded for  " & strID, True
            
            m.nPeriodicity = Bars.Prop(eBARS_Periodicity)
            m.bToEndOfData = False
            m.dToDate = LastDailyDownload
            m.dLastScreenDate = 0
            
            m.bHideExits = FileExist(App.Path & "\HideExits.flg") ' special case for Markus
            
            ' generate the chart
            m.nImgSrvChartWidth = nPixels
            If m.iMiniChart = 0 Then m.geChartHeader.paneId = m.Tree.Index("PRICE PANE")
            GenerateChart eRedo9_ReloadData
            If bTime Then frmTest.AddList "Generated chart for  " & strID, True
            
            ' export image to PNG file
            'PrintChart 1, nPixels, strBMP
            If m.iMiniChart = 0 Then
                GenerateChart eRedo1_Scrolled
            End If
            
            geSaveChart m.geChartObj, m.Form.pbChart.hWnd, m.Form.pbChart.hDC, _
                nPixels, nPixels * 0.75, 3, strPNG
            
            m.geChartHeader.paneId = 0
            m.bHideExits = False
            
            If Not FileExist(strPNG) Then
                strErr = "Error creating PNG file"
            End If
            
            'original code (begin) - save for awhile then remove 07-30-2003
            'If Not FileExist(strBMP) Then
            '    strErr = "Could not create BMP file."
            'Else
            '    If bTime Then frmTest.AddList "Created BMP for  " & strID, True
                ' convert BMP to PNG file
            '    strText = Chr(34) & strBMP & Chr(34) & " " & _
            '                Chr(34) & strPNG & Chr(34)
            '    RunProcess App.Path & "\BMP2PNG.EXE", strText, True
            '    If bTime Then frmTest.AddList "Created PNG for  " & strID, True
            '    If Not FileExist(strPNG) Then
            '        strErr = "Error creating PNG file"
            '    End If
            'End If
            'orignal code save (end)
        End If
    End If
    
    ' put output files into the image server queue
    strFile = strImgPath & strID & ".ERR"
    If Len(strErr) > 0 Then
        FileFromString strFile, strErr, True
    Else
        KillFile strFile
        FileCopy strPNG, strImgPath & strID & ".PNG"
        ' create text file with "quote" info
        aStrings.Clear
        aStrings.Add "Symbol=" & m.Bars.Prop(eBARS_Symbol)
        strText = RollSymbol(LastGoodDataBar(True))
        If Len(strText) > 0 Then
            aStrings.Add "Contract=" & strText
        End If
        aStrings.Add "Desc=" & m.Bars.Prop(eBARS_Desc)
        aStrings.Add "BarPeriod=" & m.Bars.Prop(eBARS_PeriodicityStr)
        i = m.Bars.Prop(eBARS_LastTickTime)
        If i > 0 Then
            aStrings.Add "LastTickTime=" & Format(i / 1440#, "h:mma/p")
        End If
        If m.Bars.IsIntraday Then
            nCrossover = m.Bars.Prop(eBARS_CrossoverTime)
            If nCrossover = 0 Then
                nCrossover = -1
            Else
                nCrossover = nCrossover - 1440
            End If
        Else
            nCrossover = 0
        End If
        dDate = 0
        dVol = kNullData
        dPrev = kNullData
        dLast = kNullData
        For i = m.Bars.Size - 1 To 0 Step -1
            If m.Bars(eBARS_Close, i) <> kNullData Then
                If dDate = 0 Then
                    'this is the last bar of data
                    dDate = Int(m.Bars(eBARS_DateTime, i))
                    dLast = m.Bars(eBARS_Close, i)
                    dOpen = m.Bars(eBARS_Open, i)
                    dHigh = m.Bars(eBARS_High, i)
                    dLow = m.Bars(eBARS_Low, i)
                    dVol = m.Bars(eBARS_Vol, i)
                ElseIf m.Bars(eBARS_DateTime, i) >= dDate + nCrossover / 1440# Then
                    'this is still within the last day
                    dOpen = m.Bars(eBARS_Open, i)
                    If m.Bars(eBARS_High, i) > dHigh Then dHigh = m.Bars(eBARS_High, i)
                    If m.Bars(eBARS_Low, i) < dLow Then dLow = m.Bars(eBARS_Low, i)
                    If m.Bars(eBARS_Vol, i) > 0 Then dVol = dVol + m.Bars(eBARS_Vol, i)
                Else
                    'this is close of previous day
                    dPrev = m.Bars(eBARS_Close, i)
                    Exit For
                End If
            End If
        Next
        If dLast <> kNullData Then
            aStrings.Add "Date=" & DateFormat(dDate)
            aStrings.Add "Open=" & m.Bars.PriceDisplay(dOpen, True)
            aStrings.Add "High=" & m.Bars.PriceDisplay(dHigh, True)
            aStrings.Add "Low=" & m.Bars.PriceDisplay(dLow, True)
            aStrings.Add "Last=" & m.Bars.PriceDisplay(dLast, True)
            If dVol <> kNullData Then
                aStrings.Add "Volume=" & m.Bars.PriceDisplay(dVol, True)
            End If
            If dPrev <> kNullData Then
                aStrings.Add "PrevClose=" & m.Bars.PriceDisplay(dPrev, True)
                strText = m.Bars.PriceDisplay(dLast - dPrev, True)
                If Left(strText, 1) <> "-" Then strText = "+" & strText
                aStrings.Add "Change=" & strText
                If dPrev > 0 And dLast > 0 Then
                    aStrings.Add "%Change=" & Format((dLast - dPrev) / dPrev, "+#0.00%;-#0.00%")
                End If
            End If
        End If
        ' append fundamental info
        strFile = ReplaceFileExt(strPNG, ".FND")
        If FileExist(strFile) Then
            aStrings.Add "**************"
            aStrings.FromFile strFile, True
            KillFile strFile
        End If
        aStrings.ToFile strImgPath & strID & ".TXT"
    End If
    Call RenameFile(strImgPath & strID & ".PRO", _
                    strImgPath & strID & ".DON")
    If bTime Then frmTest.AddList "Finished output for  " & strID, True
    
    ' save Rex data locally (for reference), and cleanup Rex queue
    If bUseRex Then
        m.aRexData.ToFile App.Path & "\Charts\" & strID & ".REX"
        KillFile strRexQ & strID & ".*"
    End If
        
    ' set this chart free and make sure there's a searcher
    ' (if all other charts are busy, this one can search again)
    m.strRexDataFile = ""
    m.nImgSrvFontSize = 0
    m.iMiniChart = 0
    m.Form.ImgSrvState = eImgSrv_Free
    SetImgSrvSearcher "  (" & Format((gdTickCount(False) - dStartTime) / 1000#, "0.000") _
        & " seconds for " & strSymbol & ")"
       
    Set aStrings = Nothing
    Set Bars = Nothing
                            
End Sub

' for Image Server: append Rex data
Private Sub AppendRexData(Bars As cGdBars) ', Optional ByVal dDateCutoff As Double = 999999)
On Error GoTo ErrSection:

    Dim bIsSymbol As Boolean
    Dim iLine&, strLine$, strSymbol$
    Dim iBar&, dDate#, dPrevDate#, dVol#, dOI#, dContVol#, dContOI#
    Dim nPeriodicity&, iFld&, hAppendBars&, nLastTickTime&
    Dim dMinMove#, dPrice#
    Dim aFlds As New cGdArray
    Dim AppendBars As New cGdBars
    
    If m.aRexData.Size = 0 Then Exit Sub

    ' get bars properties
    strSymbol = Bars.Prop(eBARS_Symbol)
    nPeriodicity = Bars.Prop(eBARS_Periodicity)
    SetBarProperties AppendBars, strSymbol
    If IsIntraday(nPeriodicity) Then
        ' for intraday, we'll be reading 1-minute bars and then
        ' simulating "minutized tick data" for building bars
        AppendBars.ArrayMask = eBARS_TickByTick Or eBARS_UpTicks Or eBARS_DownTicks
        AppendBars.Prop(eBARS_PeriodicityStr) = "Each Tick"
        AppendBars.Size = 2500 '(to reserve space)
    End If
    
    ' get min tick move (a little more efficient to do it once than to
    ' do it for every tick when doing intraday)
    hAppendBars = AppendBars.BarsHandle
    dMinMove = gdGetBarsMinMove(hAppendBars, Date)
    If dMinMove <= 0 Then dMinMove = 1 '(default)
       
    ' look through the data
    aFlds.Create eGDARRAY_Doubles, 10, 0
    For iLine = 0 To m.aRexData.Size - 1
        strLine = m.aRexData(iLine)
        Select Case Left(strLine, 1)
        ' !IBM
        Case "!", "#", "*", "$"
            ' see if this is the right symbol
            If Left(strLine, 1) <> "$" Then strLine = Mid(strLine, 2)
            AppendBars.Prop(eBARS_Symbol) = strLine
            If AppendBars.Prop(eBARS_Symbol) = strSymbol Then
                bIsSymbol = True
                AppendBars.Size = 0
                iBar = 0
                dPrevDate = LastDailyDownload
            Else
                bIsSymbol = False
            End If
            
        '20020715 68.400002 71.000000 66.599998 71.000000 10420900 0 0 0 1543 0.000000 0.000000 0 0 0.000000 0.000000 0.000000 0.000000
        '20020715 68.400002 68.550003 67.610001 68.309998 797600 0 0 0 0945 0.000000 0.000000 0 0 0.000000 0.000000 0.000000 0.000000
        Case "0" To "9"
            If bIsSymbol Then
                ' add data to bars
                aFlds.SplitFields strLine, " ", 10
                ' convert time from HHMM
                dDate = aFlds.Num(9)
                dDate = Int(dDate / 100) * 60 + (dDate Mod 100)
                nLastTickTime = dDate
                If IsIntraday(nPeriodicity) Then
                    '(since these are 1-minute bars, the ticks actually
                    ' occurred sometime within the previous minute)
                    dDate = JulFromLong(aFlds.Num(0)) + (dDate - 0.5) / 1440#
                Else
                    dDate = JulFromLong(aFlds.Num(0))
                End If
                ' make sure is later than last bar
                If dDate > dPrevDate Then
                    dPrevDate = dDate
                    If iBar >= AppendBars.Size Then
                        AppendBars.Size = iBar + 10
                    End If
                    ' add bar
                    If IsIntraday(nPeriodicity) Then
                        ' simulate "minutized tick data"
                        dVol = 0
                        For iFld = 1 To 4
                            gdSetNum gdBarsArray(hAppendBars, eBARS_DateTime), iBar, dDate
                            dPrice = aFlds.Num(iFld)
                            If dMinMove > 0 Then
                                dPrice = Int(dPrice / dMinMove + 0.5) * dMinMove
                            End If
                            gdSetNum gdBarsArray(hAppendBars, eBARS_Close), iBar, dPrice
                            If iFld = 4 Then
                                dVol = aFlds.Num(7)
                                If dVol = 0 Then dVol = aFlds.Num(5)
                            End If
                            gdSetNum gdBarsArray(hAppendBars, eBARS_Vol), iBar, dVol
                            '(don't know how many ticks, but must specify at least 1)
                            gdSetNum gdBarsArray(hAppendBars, eBARS_UpTicks), iBar, 1
                            gdSetNum gdBarsArray(hAppendBars, eBARS_DownTicks), iBar, 1
                            iBar = iBar + 1
                        Next
                    Else
                        ' add daily bar
                        AppendBars(eBARS_DateTime, iBar) = dDate
                        AppendBars(eBARS_Open, iBar) = AppendBars.RoundToPrice(aFlds.Num(1))
                        AppendBars(eBARS_High, iBar) = AppendBars.RoundToPrice(aFlds.Num(2))
                        AppendBars(eBARS_Low, iBar) = AppendBars.RoundToPrice(aFlds.Num(3))
                        AppendBars(eBARS_Close, iBar) = AppendBars.RoundToPrice(aFlds.Num(4))
                        dContVol = aFlds.Num(5)
                        dContOI = aFlds.Num(6)
                        dVol = aFlds.Num(7)
                        dOI = aFlds.Num(8)
                        If dVol = 0 Then dVol = dContVol
                        If dOI = 0 Then dOI = dContOI
                        AppendBars(eBARS_Vol, iBar) = dVol
                        AppendBars(eBARS_OI, iBar) = dOI
                        AppendBars(eBARS_ContVol, iBar) = dContVol
                        AppendBars(eBARS_ContOI, iBar) = dContOI
                        iBar = iBar + 1
                    End If
                End If
            End If
        End Select
    Next
    Bars.Prop(eBARS_LastTickTime) = nLastTickTime
    
    ' build correct type of bars and append
    AppendBars.Size = iBar
    If AppendBars.Size > 0 Then
        AppendBars.Prop(eBARS_Symbol) = strSymbol
        Bars.BuildBars Bars.Prop(eBARS_PeriodicityStr), AppendBars.BarsHandle, True
    End If
        
ErrExit:
    Set AppendBars = Nothing
    Exit Sub
    
ErrSection:
    Set AppendBars = Nothing
    RaiseError "cChart.AppendRexData", eGDRaiseError_Raise

End Sub

Public Function ShowTradeMsg(ByVal nTrade&, Optional ByVal bStatusMsg As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strText$, strPos$, dPrice#, dDate#, idx&, nPosition&
    Dim strRule$, strAction$, strOrderType$
    Dim bAsTradingUnits As Boolean
    Dim aFlds As New cGdArray
    Dim Pane As cPane
    
    If nTrade = -1 Then
        If bStatusMsg Then
            If UCase(Left(m.aOrders(0), 6)) = "ERROR " Then
                ' Error with orders is usually because of a position mismatch (the backtesting ended us in a different
                ' position than the generating orders started us in) -- which is usually a result the 'required bars'
                ' being set too low for this strategy.
                strText = "The 'Bars Required to Test' and/or 'Estimated Length of Longest Trade' may be set too low.||To adjust these settings, edit the strategy| and click on the 'Settings' tab."
                InfBox strText, "!", , "Error with Orders"
            Else
                ShowSystemReport True
            End If
        End If
        Exit Function
    ElseIf nTrade < 0 Then
        strText = Parse(m.aOrders(nTrade + 1000), vbCrLf, 1)
        aFlds.SplitFields strText, vbTab
        strPos = aFlds(1)
    Else
        strPos = m.Transactions.TblStr(eGDTransactionField_Action, nTrade)
    End If
    If Len(strPos) > 0 Then
        ' see if display in trading units or not
        bAsTradingUnits = True '(default)
        idx = m.Tree.Index("PRICE PANE")
        If idx > 0 Then
            Set Pane = m.Tree(idx)
            If Pane.DisplayFormat <> ePANE_PriceFormat Then
                bAsTradingUnits = False
            End If
            Set Pane = Nothing
        End If
        
        ' parse info ...
        Select Case strPos
            Case "EL"
                strPos = "Long Entry"
                strAction = "Buy"
            Case "XL"
                strPos = "Long Exit"
                strAction = "Sell"
            Case "ES"
                strPos = "Short Entry"
                strAction = "Sell"
            Case "XS"
                strPos = "Short Exit"
                strAction = "Buy"
        End Select
        If nTrade < 0 Then
            ' for an Order
            dPrice = Val(aFlds(4))
            Select Case aFlds(3)
            Case "S"
                strOrderType = "STOP"
            Case "L"
                strOrderType = "LIMIT"
            Case "M"
                strOrderType = "MARKET"
                dPrice = kNullData
            Case "MOC"
                strOrderType = "Market On Close"
                dPrice = kNullData
            Case "SCO"
                strOrderType = "Stop Close Only"
            Case "LCO"
                strOrderType = "Limit Close Only"
            Case Else
                strOrderType = aFlds(3)
            End Select
            strRule = Trim(aFlds(2))
            On Error Resume Next
            strRule = m.System.Rules.Item(strRule).Name
            On Error GoTo ErrSection:
            If bStatusMsg Then
                strText = "Place order to " & UCase(strAction) & " at "
            Else
                strText = UCase(strAction) & " at "
            End If
            If dPrice <> kNullData Then
                strText = strText & m.Bars.PriceDisplay(dPrice, bAsTradingUnits) & " "
            End If
            strText = strText & strOrderType & " (" & strPos & ")"
            nPosition = 0
        Else
            ' for a Trade
            dDate = DateTimeConvert(m.Bars, m.Transactions.TblNum(eGDTransactionField_Date, nTrade))
            dPrice = m.Transactions.TblNum(eGDTransactionField_Price, nTrade)
            strRule = m.Transactions.TblStr(eGDTransactionField_Rule, nTrade)
            strText = strPos & " at " & m.Bars.PriceDisplay(dPrice, bAsTradingUnits) _
                & " on " & DateFormat(dDate)
            If IsIntraday(m.nPeriodicity) Then
                strText = strText & Format(dDate, " HH:MM")
            End If
            nPosition = m.Transactions.TblNum(eGDTransactionField_Position, nTrade)
'strText = strText & " - " & Str(nTrade) & "," & Str(TradeNum(eTRADE_Link, nTrade))
        End If
        
        If bStatusMsg Then
            StatusMsg '(to force a little "blink")
            StatusMsg strText, vbBlue, strRule
            ' show help message (one-time only)
            If GetIniFileProperty("TradeMsgShown", False, "General", g.strIniFile) = False Then
                SetIniFileProperty "TradeMsgShown", True, "General", g.strIniFile
                strText = "Whenever you click on a trade, the details for that trade will be displayed in the status bar" _
                    & "|(next to the 'Help' menu at the top)"
                InfBox strText, "i", , "Please Note ..."
            End If
        ElseIf m.nShowTrades <> 2 Then
            ' for tooltip on chart
            If Len(strRule) > 0 Then
                strText = strText & vbCrLf & "Rule = " & strRule
            End If
            If nPosition > 0 Then
                strText = strText & vbCrLf & "Current Position = " & Str(nPosition) & " Long"
            ElseIf nPosition < 0 Then
                strText = strText & vbCrLf & "Current Position = " & Str(Abs(nPosition)) & " Short"
            End If
            ' display orders to the left
            If nTrade < 0 Then
                strText = Chr(27) & strText
            End If
        End If
    End If

    ShowTradeMsg = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ShowTradeMsg", eGDRaiseError_Raise
End Function

Public Function ShowNotationMsg(ByVal strNotation$, Optional ByVal bStatusMsg As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strText$, dDate#, i&, dValue#
    
    i = Val(Parse(strNotation, " ", 2))
    Select Case UCase(Left(strNotation, 4))
    Case "ROLL"
        If i > 0 And i < m.RollsTable.NumRecords Then
            If bStatusMsg Then
                strText = "from " & SU_GetSymbol(m.RollsTable(0, i - 1)) _
                    & " to " & SU_GetSymbol(m.RollsTable(0, i))
            Else
                strText = "from " & SU_GetSymbol(m.RollsTable(0, i - 1)) _
                    & " to " & vbCrLf & SU_GetSymbol(m.RollsTable(0, i))
            End If
            dDate = m.RollsTable(1, i)
            'If dDate > LastDailyDownload Then
            If m.RollsTable(2, i) = kNullData Then
                strText = "Forecasted to roll " & strText & " on about " & DateFormat(dDate, MM_DD_YYYY)
                If Not bStatusMsg Then
                    strText = Chr(27) & strText
                End If
            Else
                strText = "Rolled " & strText & " on " & DateFormat(dDate, MM_DD_YYYY)
            End If
        End If
    
    Case "SPLI"
        If i < m.SplitsTable.NumRecords Then
            dDate = m.SplitsTable(0, i)
            strText = "Split " & CStr(m.SplitsTable(1, i)) _
                & ":" & CStr(m.SplitsTable(2, i)) & " on " _
                & DateFormat(dDate, MM_DD_YYYY)
        End If
    
    Case "DIVI"
        If i < m.DividendsTable.NumRecords Then
            dDate = m.DividendsTable(2, i)
            dValue = m.DividendsTable(1, i)
            Select Case m.DividendsTable(0, i)
            Case 524 ' all cash-type dividends
                If dValue = Round(dValue, 2) Then
                    strText = Format(dValue, "$#0.00") & " dividend"
                Else
                    strText = Format(dValue, "$#0.0000") & " dividend"
                End If
            'Case 529 ' stock dividend (kind of like a split)
            '    strText = Str(dValue * 100) & "% stock dividend"
            'Case 530 ' stock consolidation (kind of like a reverse split)
            '    strText = Str((1 - dValue) * 100) & "% stock consolidation"
            Case Else
                strText = Format(dValue, "#0.0###") & " dividend"
            End Select
            strText = strText & " on " & DateFormat(dDate, MM_DD_YYYY)
        End If
    
    Case "EARN"
        If i > 0 Then
            strText = "Next Earnings announced on " & DateFormat(i, MM_DD_YYYY)
        End If
    
    Case "NOTE"
    
    End Select
    
    If Len(strText) > 0 And bStatusMsg Then
        StatusMsg '(to force a little "blink")
        StatusMsg strText, vbBlue
        
        ' show help message (one-time only)
        If GetIniFileProperty("NotationMsgShown", False, "General", g.strIniFile) = False Then
            SetIniFileProperty "NotationMsgShown", True, "General", g.strIniFile
            strText = "Whenever you click on a notation, the details for the notation will be displayed in the status bar" _
                & "|(next to the 'Help' menu at the top)"
            InfBox strText, "i", , "Please Note ..."
        End If
    End If

    ShowNotationMsg = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ShowNotationMsg", eGDRaiseError_Raise

End Function

Private Sub AddForecastRoll()
On Error GoTo ErrSection:

    Dim i&, nRollMonth&, nRecs&, nMonth&
    Dim nCurrentContractMonth&, nNextContractMonth&
    Dim nCount&, nSum&, lDate&, JDate&
    Dim strSymbol$, strPrevSymbol$
    Dim bIs55 As Boolean, bIs56 As Boolean, bIs57 As Boolean, bIsGann As Boolean
    Dim bIsStock As Boolean, bIsCurrency As Boolean

    nRecs = m.RollsTable.NumRecords
    If nRecs = 0 Then Exit Sub
'm.RollsTable(1, nRecs - 1) = DateSerial(2003, 11, 10)
    If m.RollsTable(1, nRecs - 1) > LastDailyDownload Then Exit Sub
    
    ' determine type of continuous contract
    strSymbol = SU_GetSymbol(m.nSymbolID)
    Select Case Val(Right(strSymbol, 2))
        Case 57, 67
            bIs57 = True
        Case 56, 66
            bIs56 = True
        Case 55, 65
            bIs55 = True
        Case 81 To 92
            bIsGann = True
    End Select
    
    If bIs57 And nRecs > 8 Then
        '(default these to true until know otherwise)
        bIsStock = True
        bIsCurrency = True
    End If
    
    ' get month of current contract
    strSymbol = SU_GetSymbol(m.RollsTable(0, nRecs - 1))
    nCurrentContractMonth = Val(Right(strSymbol, 2))
    If nCurrentContractMonth < 1 Or nCurrentContractMonth > 12 Then Exit Sub
    
    ' look for past rolls from same monthly contract
    strPrevSymbol = strSymbol
    For i = nRecs - 2 To 0 Step -1
        ' get date of roll (but only go back 5 years)
        JDate = m.RollsTable.Num(1, i + 1)
        If JDate < m.RollsTable.Num(1, nRecs - 1) - 5 * 365 Then Exit For
        lDate = JulToLong(JDate, True)
        nMonth = Int(lDate / 100) Mod 100
        
        ' check if follows currency rule
        If bIsCurrency Then
            If nMonth <> 9 Then '(Labor Day can goof this up)
                ' past rule for currencies could be either 2W-2 or 2W+1
                If JDate <> GetDateFromRule(lDate / 10000, nMonth, "2W-2") Then
                    If JDate <> GetDateFromRule(lDate / 10000, nMonth, "2W+1") Then
                        If JDate <> 38881 Then ' special case for Jun 13, 2006 (don't know why)
                            bIsCurrency = False
                        End If
                    End If
                End If
            End If
        End If
        ' check if follows stock rule
        If bIsStock Then
            If Int(lDate / 100) <> 200109 Then '(Sept. 2001 can goof this up)
                If JDate <> GetDateFromRule(lDate / 10000, nMonth, "2F-1") Then
                    bIsStock = False
                End If
            End If
        End If
        
        ' see if symbol is same contract month
        strSymbol = SU_GetSymbol(m.RollsTable.Num(0, i))
        If Val(Right(strSymbol, 2)) = nCurrentContractMonth Then
            ' make sure going to same contract
            If nNextContractMonth = 0 Then nNextContractMonth = Val(Right(strPrevSymbol, 2))
            If nNextContractMonth = Val(Right(strPrevSymbol, 2)) Then
                ' make sure same roll month
                If nRollMonth = 0 Then nRollMonth = nMonth
                If nRollMonth = nMonth Then
                    ' sum the day (to get average)
                    nCount = nCount + 1
                    nSum = nSum + (lDate Mod 100)
                End If
            End If
        End If
        strPrevSymbol = strSymbol
    Next
    If nCount = 0 Then Exit Sub
    
    ' build the next roll date
    lDate = JulToLong(m.RollsTable(1, nRecs - 1), True)
    If nRollMonth < Int(lDate / 100) Mod 100 Or bIsGann Then
        lDate = lDate + 10000 '(need to add a year)
    End If
    If bIsStock Then
        JDate = GetDateFromRule(lDate / 10000, nRollMonth, "2F-1")
    ElseIf bIsCurrency Then
        JDate = GetDateFromRule(lDate / 10000, nRollMonth, "2W+1")
    ElseIf bIs55 Then
        JDate = GetDateFromRule(lDate / 10000, nRollMonth, "FB")
    Else
        ' get average day in month
        lDate = Int(lDate / 10000) * 10000 + nRollMonth * 100 _
                    + (nSum + 0.5) / nCount
        JDate = JulFromLong(lDate)
        ' make sure not on a weekend
        i = JDate Mod 7
        If i = 0 Then
            JDate = JDate - 1
        ElseIf i = 1 Then
            JDate = JDate + 1
        End If
    End If
    
    ' must be after the day after the last date already downloaded
    i = LastDailyDownload + 1
    If Not IsWeekday(i) Then i = i + 2
    If JDate <= i Then
        JDate = i + 1
        If Not IsWeekday(JDate) Then JDate = JDate + 2
    End If
    
    ' make sure symbol for the next contract exists
    strSymbol = SU_GetSymbol(m.RollsTable(0, nRecs - 1))
    i = Val(Parse(strSymbol, "-", 2))
    i = Int(i / 100) * 100 + nNextContractMonth
    If nNextContractMonth < nCurrentContractMonth Then
        i = i + 100 '(need to add a year)
    End If
    strSymbol = Parse(strSymbol, "-", 1) & "-" & Format(i, "000000")
    i = GetSymbolID(strSymbol)
    If i < 1 Then Exit Sub
    
    ' set forecast as last record of table
    m.RollsTable(0, nRecs) = i
    m.RollsTable(1, nRecs) = JDate
    m.RollsTable(2, nRecs) = kNullData '(to flag as forecasted roll)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.AddForecastRoll", eGDRaiseError_Raise

End Sub

'MJM - [BEGIN] functions for grapheng.dll added here
Public Property Get geChartObj() As Long
    geChartObj = m.geChartObj
End Property

Public Property Get geChartPoints() As Long
    geChartPoints = m.geChartStruct.x_dataPointCount
End Property

Public Property Get geDateArray() As Long
    geDateArray = m.geChartStruct.gdhDate
End Property

Public Property Get geHitPix() As Long
    geHitPix = m.geChartStruct.hitTestPix
End Property

Public Property Let geHitPix(ByVal nPix As Long)
    m.geChartStruct.hitTestPix = nPix
End Property

' the default settings here are used when printing
' (the defaults are overridden when exporting to an image file)
Public Sub HdrPaneId(ByVal nPaneID&, Optional ByVal nTitleFontSize& = 14, _
        Optional ByVal nSubTitleFontSize& = 10, Optional ByVal nTitleStyle& = 0, _
        Optional ByVal nSubTitleStyle& = 2, _
        Optional ByVal strFont$ = "Arial")
On Error GoTo ErrSection:
    
    m.geChartHeader.paneId = nPaneID
    geSetHeaderStr nTitleFontSize, nSubTitleFontSize, nTitleStyle, nSubTitleStyle, strFont
    geForceRecalc
    
    Exit Sub
    
ErrSection:
    RaiseError "cChart.HdrPaneId"

End Sub

Public Sub FooterPaneId(ByVal nPaneID&, ByVal strText$, Optional ByVal nAlign& = 2)
On Error GoTo ErrSection:
    
    m.geChartFooter.paneId = nPaneID
    geSetFooterStr strText, nAlign      'alignment: 0=left, 1=right, 2=center
    geForceRecalc
    
    Exit Sub
    
ErrSection:
    RaiseError "cChart.FooterPaneId"
    
End Sub

Public Property Get geVisiblePaneCnt() As Long
On Error GoTo ErrSection:

    Dim i&, nCount&
    Dim Pane As cPane
    
    For i = 0 To m.Tree.Count
        Set Pane = Nothing
        If m.Tree.NodeLevel(i) = 0 Then
            Set Pane = m.Tree(i)
            If Not Pane Is Nothing Then
                If Pane.gePaneShow = 1 Then nCount = nCount + 1
            End If
        End If
    Next

    Set Pane = Nothing
    geVisiblePaneCnt = nCount

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.geVisiblePaneCnt.Get", eGDRaiseError_Raise

End Property

Public Sub geGetChartDim(nTop&, nLeft&, nBtm&, nRight&, Optional ByVal bSplitPaneWidth As Boolean = False)
On Error GoTo ErrSection:

    nTop& = gdGetNum(m.geChartStruct.glhChartRect, 0)
    nLeft = gdGetNum(m.geChartStruct.glhChartRect, 1)
    nBtm = gdGetNum(m.geChartStruct.glhChartRect, 2)
    nRight = gdGetNum(m.geChartStruct.glhChartRect, 3)
    
    If bSplitPaneWidth Then
        'precautionary: don't add negative number
        If m.geChartStruct.splitPaneWidth > 0 Then nRight = nRight + m.geChartStruct.splitPaneWidth
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.geGetChartDim", eGDRaiseError_Raise

End Sub

Private Function geSetWinDim(ByVal nTop&, ByVal nLeft&, ByVal nBtm&, ByVal nRight&) As Long
On Error GoTo ErrSection:

    Dim nWinTop&, nWinLeft&, nWinBtm&, nWinRight&, rc&
    
    nTop = nTop / Screen.TwipsPerPixelX
    nLeft = nLeft / Screen.TwipsPerPixelX
    nBtm = nBtm / Screen.TwipsPerPixelX
    nRight = nRight / Screen.TwipsPerPixelX
    
    nWinTop = gdGetNum(m.geChartStruct.glhWinRect, 0)
    nWinLeft = gdGetNum(m.geChartStruct.glhWinRect, 1)
    nWinBtm = gdGetNum(m.geChartStruct.glhWinRect, 2)
    nWinRight = gdGetNum(m.geChartStruct.glhWinRect, 3)
        
    If nTop <> nWinTop Or nLeft <> nWinLeft Or nBtm <> nWinBtm Or nRight <> nWinRight Then
        gdSetNum m.geChartStruct.glhWinRect, 0, nTop
        gdSetNum m.geChartStruct.glhWinRect, 1, nLeft
        gdSetNum m.geChartStruct.glhWinRect, 2, nBtm
        gdSetNum m.geChartStruct.glhWinRect, 3, nRight
        
        If m.geChartObj <> 0 Then rc = geRecalcChart(m.geChartObj, m.geChartStruct, Form.pbChart.hDC)
    End If
    
    geSetWinDim = rc
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.geSetWinDim", eGDRaiseError_Raise

End Function

Public Sub SetCursor(Optional ByVal nCursor& = 0)
On Error GoTo ErrSection:

    Dim rc&
    Static nPencilCursor&, nPrevCursor&


    If nCursor = 0 Then 'default
        If Len(g.strActiveDraw) > 0 Then
            m.Form.pbCursor = eCursor_Pencil
        ElseIf frmMain.tbToolbar.Tools("ID_ChartOrderBuy").State = ssChecked Then
            m.Form.pbCursor = eCursor_OrderBuy
        ElseIf frmMain.tbToolbar.Tools("ID_ChartOrderSell").State = ssChecked Then
            m.Form.pbCursor = eCursor_OrderSell
        ElseIf frmMain.tbToolbar.Tools("ID_CursorCrosshairs").State = ssChecked Then
            m.Form.pbCursor = eCursor_CrossHair
        ElseIf frmMain.tbToolbar.Tools("ID_CursorVertLine").State = ssChecked Then
            m.Form.pbCursor = eCursor_Vertical
        ElseIf frmMain.tbToolbar.Tools("ID_CursorHorizLine").State = ssChecked Then
            m.Form.pbCursor = eCursor_Horizontal
        Else
            m.Form.pbCursor = eCursor_Default
        End If
    End If
  
'    If Not g.oAvailToolButtons Is Nothing Then
'        If m.tbToolbar.Tools("ID_CursorCrosshairs").State = ssChecked Then
'            g.oAvailToolButtons.ToolBtnStateSync frmMain, frmMain.pbNotUsed, "ID_CursorCrosshairs", eBtnState_Selected, True
'        ElseIf m.tbToolbar.Tools("ID_CursorVertLine").State = ssChecked Then
'            g.oAvailToolButtons.ToolBtnStateSync frmMain, frmMain.pbNotUsed, "ID_CursorVertLine", eBtnState_Selected, True
'        ElseIf m.tbToolbar.Tools("ID_CursorHorizLine").State = ssChecked Then
'            g.oAvailToolButtons.ToolBtnStateSync frmMain, frmMain.pbNotUsed, "ID_CursorHorizLine", eBtnState_Selected, True
'        Else
'            g.oAvailToolButtons.ToolBtnStateSync frmMain, frmMain.pbNotUsed, "ID_CursorArrow", eBtnState_Selected, True
'        End If
'    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.SetCursor", eGDRaiseError_Raise

End Sub

Public Sub SetEqualSizePanes(PricePane As cPane)
On Error GoTo ErrSection:

    'this sub is called to set all non-price pane sizes the same
    'when user drags separator with shift-key down
    
    Dim Pane As cPane, idx&, i&
    Dim nVisPanes&      'number of visible panes
    Dim dEqualSize#     'height proportion of all non-price panes
    
    idx = PricePane.gePaneId
    nVisPanes = geVisiblePaneCnt - 1
    dEqualSize = (1 - PricePane.Size) / nVisPanes
    For i = 0 To m.Tree.Count
        Set Pane = Nothing
        If m.Tree.NodeLevel(i) = 0 Then
            If i <> idx Then
                Set Pane = m.Tree(i)
                If Not Pane Is Nothing Then
                    If Pane.Display = True Then
                        Pane.Size = dEqualSize
                    End If
                End If
            End If
        End If
    Next
    
ErrExit:
    Set Pane = Nothing
    Exit Sub
    
ErrSection:
    Set Pane = Nothing
    RaiseError "cChart.SetEqualSizePanes", eGDRaiseError_Raise

End Sub

Public Sub geResetPanes()
On Error GoTo ErrSection:

    Dim nSize&, i&
    Dim Annot As cAnnotation
    Dim treeObj As Variant
    
    If m.bZoomed = True Then Exit Sub
    
    RemoveAnnots False, eANNOT_UndefinedType, eANNOT_IndicatorLabel
    RemoveAnnots False, , eANNOT_AutoSwingTrend
    
    'reset flags for indicators
    nSize = m.Tree.Count
    For i = 1 To nSize
        Set treeObj = m.Tree(i)
        If Not treeObj Is Nothing Then treeObj.geResetData
    Next
    
    'reset flags for annotations
    nSize = m.Annots.Count
    For i = 1 To nSize
        Set Annot = m.Annots(i)
        If Not Annot Is Nothing Then Annot.geResetData
    Next
    
    'reset flags for composite indicators
    If Not m.CompositeIndicators Is Nothing Then
        nSize = m.CompositeIndicators.Count
        For i = 1 To nSize
            Set treeObj = m.CompositeIndicators(i)
            If Not treeObj Is Nothing Then treeObj.geResetData
        Next
    End If
    
    If m.geChartObj <> 0 Then
        i = geRemovePanes(m.geChartObj)
        If i <> 0 Then
            DebugLog "geRemovePanes error: " & Str(i) & "   symbol=" & Symbol
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.geResetPanes", eGDRaiseError_Raise

End Sub

Public Sub geForceRecalc()
On Error GoTo ErrSection:

    geRecalcChart m.geChartObj, m.geChartStruct, Form.pbChart.hDC

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.geForceRecalc", eGDRaiseError_Raise

End Sub

Public Function UnzoomChart(Optional ByVal bRefresh As Boolean = True) As Long
    
    Dim rc&
    
    m.bHidePriceIndicators = False
    m.hsb.Max = m.aXBar.Size - 1
    rc = geUnzoomChart(m.geChartObj, m.Form.pbChart.hDC)
    m.bZoomed = False
    If bRefresh Then
        If g.RealTime.Active Then geChartSettings True      'aardvark 4130
        GenerateChart eRedo1_Scrolled
    End If
    SyncToolbar
    
End Function

Public Function geDrawChart() As Long
On Error GoTo ErrSection:

    Dim rc&, strText$
    Dim Pane As cPane
    
    If m.Form Is Nothing Then Exit Function
    If m.Form.Visible = False Then Exit Function            '6682
    If m.Form.tmr.Tag = "UNLOADING" Then Exit Function
        
    ' If no data exists, just display a message
    ' (instead of trying to draw a chart with no data)
    If m.geChartObj = 0 Or m.nLastGoodDataBarX < 0 Then
        strText = "NO DATA"
        If Not m.Bars Is Nothing Then
            If Len(m.strSpreadSymbols) > 0 Then
                strText = BarPeriodError(m.strSpreadSymbols, m.nPeriodicity)
            Else
                strText = BarPeriodError(m.Bars.Prop(eBARS_Symbol), m.nPeriodicity)
            End If
            If Len(strText) = 0 Then
                If IsIntraday(m.nPeriodicity) Then
                    strText = "NO INTRADAY DATA"
                ElseIf TypeOfChart = eTypeChart_Seasonal Then
                    Set Pane = m.Tree("PRICE PANE")
                    If Pane Is Nothing Then
                        strText = "NO DATA"
                    Else
                        strText = Pane.Desc
                    End If
                Else
                    strText = "NO DATA"
                End If
            End If
        End If
        With m.Form.vseInvalid
            .Caption = strText
            .BackColor = g.ChartGlobals.nChartBackColor
            If g.ChartGlobals.nChartBackColor = 0 Then
                .ForeColor = vbWhite
            Else
                .ForeColor = 0
            End If
            .Visible = True
            .ZOrder
        End With
        m.Form.DoMouseLabel True
        Exit Function
    End If
    m.Form.vseInvalid.Visible = False
    
    If m.geChartStruct.x_dataPointPix <> m.nPixelsPerBar Then
        'If these 2 are not equal then chart settings were just reset to default of 7 pix per bar
        'by generate_chart. Don't draw since recalc_chart needs to be called to make these 2 the same.
        'This condition will be true when this function is called from outside generate_chart within
        'a timer like from chart form to refresh cursor and timing is such that generate_chart just
        'reset chart settings for new data bar, but didn't get a chance to recalc_chart (4140).
        Exit Function
    End If
       
'enum from grapheng header file - copied here for easy reference
'typedef enum{
'    e_vtCoarse,         //0=default
'    e_vtFine,           //1
'    e_bothCoarse,       //2
'    e_bothFine,         //3
'    e_hzOnly,           //4
'    e_noGrid            //5
'} GRID_MODE;
    
    'set grid style
    If m.nVertGrid < 0 Or m.nVertGrid > 2 Then m.nVertGrid = 0
    Select Case m.nVertGrid
        Case 0      'coarse vertical grid
            If m.bHorzGrid = True Then
                m.geChartStruct.gridMode = 2
            Else
                m.geChartStruct.gridMode = 0
            End If
        Case 1      'fine vertical grid
            If m.bHorzGrid = True Then
                m.geChartStruct.gridMode = 3
            Else
                m.geChartStruct.gridMode = 1
            End If
        Case 2
            If m.bHorzGrid = True Then
                m.geChartStruct.gridMode = 4
            Else
                m.geChartStruct.gridMode = 5
            End If
    End Select
    
    rc = geSetWinDim(m.Form.pbChart.ScaleTop, _
                     m.Form.pbChart.ScaleLeft, _
                     m.Form.pbChart.ScaleHeight, _
                     m.Form.pbChart.ScaleWidth)
    
    If rc = 0 Then
''gdStartProfile 381
        rc = geDrawWindow(m.geChartObj, m.Form.pbChart.hWnd, m.Form.pbChart.hDC)
''gdStopProfile 381
        If rc <> 0 And rc <> 232 Then
            If rc = 1064 Then
                geChartSettings True
                GenerateChart eRedo1_Scrolled       '4163 - one-bar bug
            ElseIf FileExist("c:\common\files.exe") Or rc <> 1064 Then
                StatusMsg "geDrawWindow failed - rc = " & CStr(rc)
            End If
        End If
    End If
    
    geDrawChart = rc
    m.Form.ShowCursor
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.geDrawChart", eGDRaiseError_Raise

End Function

Private Function geChartSettings(ByVal bResetPanes As Boolean) As Long
On Error GoTo ErrSection:

    Dim rc&, i&, j&, idxPane&, idxPriceInd&
    Dim Pane As cPane, Ind As cIndicator
    Dim bInd As Boolean
    Dim OrderSpec As chart_order_spec
    Dim strFile$

    If m.Tree.Count = 0 Then
        If m.geChartObj <> 0 Then
            geResetPanes
            geCloseChart m.geChartObj
            m.geChartObj = 0
        End If
        TemplateLoad m.strTemplate
    End If

    
    If RedoMode = eRedo9_ReloadData Or bResetPanes Then
        geResetPanes
    End If
        
    rc = geSetWinDim(m.Form.pbChart.ScaleTop, m.Form.pbChart.ScaleLeft, _
                     m.Form.pbChart.Height, m.Form.pbChart.ScaleWidth)
        
    ' chart colors
    If m.bCustomColors Then
        m.geChartStruct.borderBkColor = m.nBorderBackColor
        m.geChartStruct.borderFgColor = m.nBorderForeColor
        m.geChartStruct.chartBkColor = m.nChartBackColor
        m.geChartStruct.chartFgColor = m.nChartForeColor
'        m.Form.vseMouse.BackColor = m.nBorderBackColor
'        m.Form.vseMouse.ForeColor = m.nBorderForeColor
'        m.Form.vseDay.BackColor = m.nBorderBackColor
'        m.Form.vseDay.ForeColor = m.nBorderForeColor
        If m.nUseGradient = 1 Then
            m.geChartStruct.GradientColor = m.nChartGradientColor
        Else
            m.geChartStruct.GradientColor = -1
        End If
    Else
        m.geChartStruct.borderBkColor = g.ChartGlobals.nBorderBackColor
        m.geChartStruct.borderFgColor = g.ChartGlobals.nBorderForeColor
        m.geChartStruct.chartBkColor = g.ChartGlobals.nChartBackColor
        m.geChartStruct.chartFgColor = g.ChartGlobals.nChartForeColor
'        m.Form.vseMouse.BackColor = g.ChartGlobals.nBorderBackColor
'        m.Form.vseMouse.ForeColor = g.ChartGlobals.nBorderForeColor
'        m.Form.vseDay.BackColor = g.ChartGlobals.nBorderBackColor
'        m.Form.vseDay.ForeColor = g.ChartGlobals.nBorderForeColor
        If g.ChartGlobals.nUseGradient = 1 Then
            m.geChartStruct.GradientColor = g.ChartGlobals.nChartGradientColor
        Else
            m.geChartStruct.GradientColor = -1
        End If
        m.Form.DoMouseLabel
    End If
    If m.bUnpublishable Then
        m.Form.vseDay.BackColor = vbRed
    End If

    gdSetStr m.geChartStruct.gshFont, 0, g.ChartGlobals.strFontName
    m.geChartStruct.fSize = g.ChartGlobals.nFontSize
    m.geChartStruct.fStyle = g.ChartGlobals.nFontStyle
    
    m.geChartStruct.PriceTopMost = m.nPriceTopMost
    m.geChartStruct.y_labelMaxChar = 7
    m.geChartStruct.x_dataPointPix = 7      'Q & A - need to tie this to bars per inch..
    m.geChartStruct.paneSepWidth = 3
    m.geChartStruct.hitTestPix = 2
    If m.Bars Is Nothing Then
        m.geChartStruct.glhBars = 0
        gdSetStr m.geChartStruct.gshTimeZoneInf, 0, ""
    Else
        m.geChartStruct.glhBars = m.Bars.BarsHandle
        gdSetStr m.geChartStruct.gshTimeZoneInf, 0, Me.Bars.Prop(eBARS_ExchangeTimeZoneInf)
    End If
        
    'Programmer's Note: the pane & indicator idxPriceBars is the index
    '   of the indicator that controls how price data are format
    Set Pane = Nothing
    Set Ind = Nothing
    For i = 1 To m.Tree.Count
        If m.Tree.NodeLevel(i) = 0 Then
            Set Pane = m.Tree(i)
            If Not Pane Is Nothing Then
                idxPriceInd = 0
                Pane.idxPriceBars = 0 ' m.Tree.Index("PRICE")
                idxPane = i
            End If
        Else
            Set Ind = m.Tree(i)
        End If
        If Not Ind Is Nothing And Not Pane Is Nothing Then
            With Ind
                If .Overlayed = True Then
                    .idxPriceBars = -1
                ElseIf .DataType = eINDIC_BarData Then
                    idxPriceInd = i
                    .idxPriceBars = i
                    If Left(m.Tree.Key(idxPane), 5) <> "PRICE" And Pane.idxPriceBars = 0 Then
                        Pane.idxPriceBars = idxPriceInd
                    End If
                Else
                    .idxPriceBars = idxPriceInd
                End If
            End With
        End If
        Set Ind = Nothing
        If m.Tree.NodeLevel(i + 1) = 0 Then Set Pane = Nothing
    Next
    
    'FixPaneSizes
    'initialize grapheng chart object if necessary
    If rc = 0 And m.geChartObj = 0 Then
        m.geChartObj = geInitChart(m.geChartStruct, AddressOf fnFormatPrice, Form.pbChart.hDC)

'JM 11-06-2012: Watermark code for PNG file Liz sent (email forwarded from Tim)
'        If m.geChartObj <> 0 Then
'            If FileExist(kWaterPng) Then
'                strFile = g.strAppPath & "\" & kWaterPng
'                geWaterMarkFile m.geChartObj, strFile$
'            End If
'        End If

    ElseIf Not g.RealTime.Active Then
        geForceRecalc       'aardvark 3685 (don't do this if real time is on)
    End If
    
    OrderSpec.nLongColor = ChartGlbClrForCtl(m.Form.pbChart, g.ChartGlobals.nLongColor, "LongColor")
    OrderSpec.nShortColor = g.ChartGlobals.nShortColor
    OrderSpec.nOtherColor = RGB(128, 128, 128)
    OrderSpec.nEntryFill = 0
    
    geSetOrderSpec m.geChartObj, OrderSpec
    geSetProfileInfo m.geChartObj, m.profileChartInfo
    
    geChartSettings = rc
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.geChartSettings", eGDRaiseError_Raise

End Function

Private Sub geSetImgSvrFontSize()
On Error GoTo ErrSection:
    
    Dim nHdrFont&, nSubHdrFont&
    
    If m.nImgSrvChartWidth >= 350 Then
        nHdrFont = 10
        nSubHdrFont = 8
    Else
        nHdrFont = 6
        nSubHdrFont = 6
    End If
    
    m.geChartStruct.fSize = g.ChartGlobals.nFontSize
    gdSetStr m.geChartStruct.gshFont, 0, g.ChartGlobals.strFontName '"Times New Roman" works best for web server
    geSetHeaderStr nHdrFont, nSubHdrFont, 1, 2, g.ChartGlobals.strFontName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.geSetImgSvrFontSize", eGDRaiseError_Raise

End Sub

Private Sub geSetFooterStr(ByVal strText$, Optional ByVal nAlign = 0)

    Dim rc&
    
    If gdGetSize(m.geChartFooter.glhAnnType) = 0 Then
        m.geChartFooter.glhAnnType = gdCreateArray(eGDARRAY_Longs)
        m.geChartFooter.gshText = gdCreateArray(eGDARRAY_Strings)
        m.geChartFooter.gshFont = gdCreateArray(eGDARRAY_Strings)
        m.geChartFooter.glhStyle = gdCreateArray(eGDARRAY_Longs)
        m.geChartFooter.glhSize = gdCreateArray(eGDARRAY_Longs)
        m.geChartFooter.gdhTop = gdCreateArray(eGDARRAY_Doubles)
        m.geChartFooter.gdhLeft = gdCreateArray(eGDARRAY_Doubles)
        m.geChartFooter.gdhBottom = gdCreateArray(eGDARRAY_Doubles)
        m.geChartFooter.gdhRight = gdCreateArray(eGDARRAY_Doubles)
        m.geChartFooter.glhAlign = gdCreateArray(eGDARRAY_Longs)
    End If

'    gdSetNum m.geChartFooter.glhAnnType, 0, 17
'    'use same font type & size as header subtitle
'    gdSetNum m.geChartFooter.glhSize, 0, 10     'gdGetNum(m.geChartHeader.glhSize, 1) '+ 2
'    gdSetStr m.geChartFooter.gshFont, 0, gdGetStr(m.geChartHeader.gshFont, 1)
'    gdSetStr m.geChartFooter.gshText, 0, "TradeNavigator.com ©2007"
'    gdSetNum m.geChartFooter.glhAlign, 0, 0
    
    If Len(strText) > 0 Then
        gdSetNum m.geChartFooter.glhAnnType, 0, 17
        'use same font type & size as header subtitle
        gdSetNum m.geChartFooter.glhSize, 0, gdGetNum(m.geChartHeader.glhSize, 1) '+ 2
        gdSetStr m.geChartFooter.gshFont, 0, gdGetStr(m.geChartHeader.gshFont, 1)
        gdSetStr m.geChartFooter.gshText, 0, strText
        gdSetNum m.geChartFooter.glhAlign, 0, 0
    Else
        gdSetNum m.geChartFooter.glhAnnType, 0, -1      '4774
    End If
            
    If nAlign > 2 Then
        nAlign = 2   '0=left, 1=right, 2=center
    ElseIf nAlign < 0 Then
        nAlign = 0
    End If
        
    rc = geAddItem(m.geChartObj, 4, m.geChartFooter)
    
End Sub

Private Sub geSetHeaderStr(ByVal nTitleFontSize&, _
        ByVal nSubTitleFontSize&, Optional ByVal nTitleStyle& = 0, _
        Optional ByVal nSubTitleStyle& = 2, _
        Optional ByVal strFont$ = "Arial")
On Error GoTo ErrSection:

    Dim rc&, nPoolRec&, strBars$

    If Len(m.strSubTitle) < 1 Then
        If m.nSymbolID > 0 Then
            nPoolRec = g.SymbolPool.PoolRecForSymbolID(m.nSymbolID)
            strBars = GetPeriodStr(m.nPeriodicity) ' m.Bars.Prop(eBARS_PeriodicityStr)
            m.strTitle = g.SymbolPool.Symbol(nPoolRec) & ":  " & g.SymbolPool.Desc(nPoolRec) _
                            & "  (" & strBars & " bars)"
        ElseIf Len(m.strSpreadSymbols) > 0 Then     'aarvark 3344
            rc = InStr(m.Form.vseCaption, ")")
            If rc > 0 Then
                m.strTitle = Left(m.Form.vseCaption, rc)
            End If
        End If
        
        m.strSubTitle = GetProvidedProperty("Website", , True)
        If ExtremeCharts = 0 Then
            rc = Year(Date)
            If rc < 2008 Then rc = 2008
            'JM 07-25-2013: Glen wanted "All rights reserved"; he thinks it is redundant
            '   and removing it makes the subtitle "fits better" on chart export
            m.strSubTitle = m.strSubTitle & " © 1999-" & Str(rc) '& " All rights reserved"
        End If
    End If
    
    If gdGetSize(m.geChartHeader.glhAnnType) = 0 Then
        m.geChartHeader.glhAnnType = gdCreateArray(eGDARRAY_Longs)
        m.geChartHeader.gshText = gdCreateArray(eGDARRAY_Strings)
        m.geChartHeader.gshFont = gdCreateArray(eGDARRAY_Strings)
        m.geChartHeader.glhStyle = gdCreateArray(eGDARRAY_Longs)
        m.geChartHeader.glhSize = gdCreateArray(eGDARRAY_Longs)
        m.geChartHeader.gdhTop = gdCreateArray(eGDARRAY_Doubles)
        m.geChartHeader.gdhLeft = gdCreateArray(eGDARRAY_Doubles)
        m.geChartHeader.gdhBottom = gdCreateArray(eGDARRAY_Doubles)
        m.geChartHeader.gdhRight = gdCreateArray(eGDARRAY_Doubles)
        m.geChartHeader.glhAlign = gdCreateArray(eGDARRAY_Longs)
    End If
    
    'm.geChartHeader.paneId = 1
    
    gdSetNum m.geChartHeader.glhAnnType, 0, 10
    gdSetNum m.geChartHeader.glhAnnType, 1, 10
    
    gdSetNum m.geChartHeader.glhSize, 0, nTitleFontSize     '16 for printing
    gdSetNum m.geChartHeader.glhSize, 1, 10     'nSubTitleFontSize  '10 for printing
            
    gdSetNum m.geChartHeader.gdhTop, 0, 1
    gdSetNum m.geChartHeader.gdhLeft, 0, 1
    
    gdSetNum m.geChartHeader.gdhTop, 1, 1
    gdSetNum m.geChartHeader.gdhLeft, 1, 1

    gdSetStr m.geChartHeader.gshText, 0, m.strTitle
    gdSetStr m.geChartHeader.gshText, 1, m.strSubTitle
    
    gdSetNum m.geChartHeader.glhStyle, 0, nTitleStyle       '0 = regular for printing
    gdSetNum m.geChartHeader.glhStyle, 1, 0     'nSubTitleStyle    '2 = italic for printing
    
    gdSetNum m.geChartHeader.glhAlign, 0, 5
    gdSetNum m.geChartHeader.glhAlign, 1, 4
    
    gdSetStr m.geChartHeader.gshFont, 0, strFont            'arial for printing
    gdSetStr m.geChartHeader.gshFont, 1, strFont            'arial for printing
            
    rc = geAddItem(m.geChartObj, 4, m.geChartHeader)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.geSetHeaderStr", eGDRaiseError_Raise

End Sub

Public Sub FakeForecastBarsChange(ByVal nShift&)
On Error GoTo ErrSection:

    'temporarily fake an increase/decrease in the number of forecast bars
    Dim nSize&, i&
    Dim Ind As cIndicator
    Dim Annot As cAnnotation
        
    'shift date array
    gdSetShifted m.geChartStruct.gdhDate, nShift, True
    
    'step through indicators and fake data shift
    nSize = m.Tree.Count
    For i = 1 To nSize
        Set Ind = Nothing
        If m.Tree.NodeLevel(i) > 0 Then
            Set Ind = m.Tree(i)
            If Not Ind Is Nothing Then
                If Ind.Display = True Then Ind.FakeDataShift nShift
            End If
        End If
    Next
    
    'step through annotations and fake shift
    nSize = m.Annots.Count
    For i = 1 To nSize
        Set Annot = m.Annots(i)
        If Not Annot Is Nothing Then
            If Annot.eUsage = eANNOT_UserAdded Or Annot.eUsage = eANNOT_Notation _
                Or Annot.eUsage = eANNOT_Trades Or Annot.eUsage = eANNOT_ProfitLines _
                Or Annot.eType = eANNOT_MultiRects Then             '4546
                    Annot.FakeShift nShift
            End If
        End If
    Next
            
    Set Ind = Nothing
    Set Annot = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.FakeForecastBarsChange", eGDRaiseError_Raise

End Sub

Public Function LoadExportData(ByVal nPixWidth&, ByVal nPixHeight) As Long
On Error GoTo ErrSection:

    Dim nSaveDpoints&, nSaveWinLeft&, nSaveWinRight&, nSaveXPointPix&
    Dim nNewPix&, nNewDpointCount&, nDiff&, rc&
        
    'save settings
    nSaveWinLeft = gdGetNum(m.geChartStruct.glhWinRect, 1)
    nSaveWinRight = gdGetNum(m.geChartStruct.glhWinRect, 3)
    nSaveDpoints = m.geChartStruct.x_dataPointCount
    nSaveXPointPix = m.geChartStruct.x_dataPointPix
        
    rc = geSaveChart(m.geChartObj, m.Form.pbChart.hWnd, m.Form.pbChart.hDC, _
        nPixWidth, nPixHeight, 1, "", nNewPix, nNewDpointCount)
        
    If rc <> 0 Then
        LoadExportData = rc
        Exit Function
    End If
        
    nDiff = nSaveDpoints - nNewDpointCount
        
    'restore settings
    gdSetNum m.geChartStruct.glhWinRect, 1, nSaveWinLeft
    gdSetNum m.geChartStruct.glhWinRect, 3, nSaveWinRight
    m.geChartStruct.x_dataPointCount = nSaveDpoints
    m.geChartStruct.x_dataPointPix = nSaveXPointPix
    
    LoadScreenData ScreenStartX - Abs(nDiff)
              
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.LoadExportData", eGDRaiseError_Raise

End Function

Public Sub RestoreScreenData()
On Error GoTo ErrSection:

    LoadScreenData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.LoadExportData", eGDRaiseError_Raise

End Sub

'JM (revision note):
'SourceSafe version 777 11/25/09 12:57 is last good LoadScreenData prior to removal dinapoli "objects"
Private Sub LoadScreenData(Optional ByVal nStartX& = -1, Optional ByVal bCalcClusters As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, j&, rc&
    Dim nPaneID&
    Dim hXbarArray&, hXdateArray&
    Dim nStartScreenX&
    Dim nIndAlert&      '-1=inactive,0=none,1=active
    
    Dim dBarDate#, dTempDate#, strSysName$
    Dim Pane As cPane, PricePane As cPane
    
    Dim Ind As cIndicator
    Dim PriceInd As cIndicator
    Dim FirstVisibleInd As cIndicator
    
    Dim Annot As cAnnotation
    
    Dim bHasOverlayed As Boolean
    Dim bHasWoodPane As Boolean
    Dim bHasValidStyle As Boolean       'true = at least one indicator in pane has display type <> NoStyle
    Dim bExport As Boolean
    Static bAlreadyCheckedArrowMsg As Boolean
           
'gdStartProfile 361
           
    If nStartX > 0 Then
        nStartScreenX = nStartX
        bExport = True                      'As of 03-18-2008: nStartX is only passed in by the export data routine
    Else
        nStartScreenX = ScreenStartX
    End If
    ' save handles to arrays (for efficiency)
    ' and save any special offsets
    m.nFirstBar = 0
    m.nLastBar = 0
    hXbarArray = m.aXBar.ArrayHandle
    hXdateArray = m.aXdate.ArrayHandle
    
    For i = nStartScreenX To ScreenEndX
        If gdGetNum(hXbarArray, i) >= 0 Then
            m.nFirstBar = gdGetNum(hXbarArray, i) 'Q&A - skipping null data?
            Exit For
        End If
    Next
    For i = ScreenEndX To nStartScreenX Step -1
        If gdGetNum(hXbarArray, i) >= 0 Then
            m.nLastBar = gdGetNum(hXbarArray, i)
            Exit For
        End If
    Next
    
    'load dates data into grapheng.dll chart structure
    gdClear m.geChartStruct.gdhDate, False
    j = 0
    For i = nStartScreenX To ScreenEndX
        'gdSetNum m.geChartStruct.gdhDate, j, m.aXdate.Item(i)
        gdSetNum m.geChartStruct.gdhDate, j, gdGetNum(hXdateArray, i)
        j = j + 1
    Next
    
    If m.iMiniChart > 0 Then
        If j < m.geChartStruct.x_dataPointCount Then
            For i = j To m.geChartStruct.x_dataPointCount - 1
                gdSetNum m.geChartStruct.gdhDate, i, kNullData
            Next
        End If
    End If
    
    'set scale arrow flag
    If bExport Then
        m.geChartStruct.scaleArrow = 0
    ElseIf TypeOfChart = eTypeChart_Seasonal And m.hsb.Max = m.Bars.Size - 1 Then
        'all data in bars object are consider "good", just need to confirm scroll bar max is set correctly
        m.geChartStruct.scaleArrow = 0
    Else
        i = m.hsb.Value
        j = m.nLastGoodDataBarX + m.nBlankBars
        
        If i < j Then
            If IsInWhatIfMode Then
                'when user chooses "New Bar" in whatif mode j-i will = 1 because of extra "new bar"
                If j - i > 1 Then m.geChartStruct.scaleArrow = 2
            ElseIf m.bProfileChart And i = m.aXdate.Size - 1 Then
                m.geChartStruct.scaleArrow = 0
            Else
                m.geChartStruct.scaleArrow = 2          'point right
            End If
        ElseIf i > j Then
            m.geChartStruct.scaleArrow = 1
        Else
            m.geChartStruct.scaleArrow = 0
        End If
        If m.geChartStruct.scaleArrow <> 0 And Not g.bStarting And Not bAlreadyCheckedArrowMsg Then
            If Not GetIniFileProperty("DontShowArrowMsg", False, "Charting", g.strIniFile) Then
                InfBox "Clicking the red arrow in the lower-right| or using the 'End' key will reposition| chart to default blank bars.", "i", , "Note ..."
                SetIniFileProperty "DontShowArrowMsg", True, "Charting", g.strIniFile
            End If
            bAlreadyCheckedArrowMsg = True
        End If
    End If
    
    SetDateStyle
    
    'set minutes per bar for intraday data
    If m.Bars.IsIntraday = True Then
        m.geChartStruct.Periodicity = m.Bars.Prop(eBARS_PeriodsPerBar)
    End If
    
    FixPaneSizes
    SetBadTicksIndicator
    ValidateHighlightBars
    RemoveAnnots False, eANNOT_MultiRects, eANNOT_Indicator
    RemoveAnnots False, eANNOT_BellAlert
    
    If m.eRedoMode > eRedo1_Scrolled Then
        RemoveAnnots False, , eANNOT_AutoSwingTrend
    End If
                      
    ' see if need to temporarily set price pane to auto-scale
    ' (e.g. if change symbol, template, bar period, etc.)
    If m.bTempAutoScale Then
        Set Pane = m.Tree("Price Pane")
        If Not Pane Is Nothing Then
            If Pane.Scaling = ePANE_ScaleModeManual Then
                Pane.Scaling = Pane.ScaleAutoLastUsed       'aardvark 4652
            Else
                m.bTempAutoScale = False
            End If
        End If
    End If
            
'gdStopProfile 361

    'walk through tree adding panes and indicators to grapheng.dll
    For i = 1 To m.Tree.Count
        If m.Tree.NodeLevel(i) = 0 Then
'gdStartProfile 362
            If i > 1 And Not bHasValidStyle Then
                'user has set all display type to no style, change type of first indicator
                If Not FirstVisibleInd Is Nothing Then
                    If FirstVisibleInd.DataType = eINDIC_Array Or Ind.DataType = eINDIC_Constant Then
                        FirstVisibleInd.DisplayType = eINDIC_Line
                        bHasValidStyle = True
                    ElseIf FirstVisibleInd.DataType = eINDIC_BarData Then
                        FirstVisibleInd.DisplayType = eINDIC_OHLC
                        bHasValidStyle = True
                    End If
                    If bHasValidStyle Then
                        rc = FirstVisibleInd.geSetIndData(Me, nPaneID, FirstVisibleInd.geIndId, _
                            m.nFirstBar, m.nLastBar, nStartScreenX, ScreenEndX, m.dExtraPriceScale, bCalcClusters)
                    End If
                End If
            End If
            
            ' init new pane
            Set Pane = m.Tree(i)
            nPaneID = i
            bHasValidStyle = False
            Set FirstVisibleInd = Nothing
                                    
            If m.Tree.Key(i) = kClusterTimeKeyPane Then
                Pane.Scaling = ePANE_ScaleModeAuto
                Pane.YscaleLabelAll = 2
            End If
            
            rc = Pane.geAddPane(Me, m.geChartObj, i)
            Pane.gePaneMin = kNullData
            If m.Tree.Key(i) = "PRICE PANE" Then
                Set PricePane = Pane
                PricePane.PricePaneFlag = 1
            ElseIf InStr(UCase(Pane.Name), "WOODIES") > 0 And Pane.Display Then     '6060
                bHasWoodPane = True
            End If
'gdStopProfile 362
        Else
            Set Ind = m.Tree(i)
            'handle non-overlayed indicators first
            If Ind.Overlayed = True Then
'gdStartProfile 363
                If Ind.Display Then
                    If Ind.DataType <> eINDIC_BooleanArray And Ind.DisplayType <> eINDIC_NoStyle Then bHasValidStyle = True
                    If FirstVisibleInd Is Nothing Then
                        Set FirstVisibleInd = Ind
                    End If
                End If
                bHasOverlayed = True
'gdStopProfile 363
            Else
'gdStartProfile 364
                If m.iMiniChart > 0 Then Ind.IndLabelMode = eINDIC_Nothing
'gdStopProfile 364
'gdStartProfile 365
                rc = Ind.geSetIndData(Me, nPaneID, i, _
                    m.nFirstBar, m.nLastBar, nStartScreenX, ScreenEndX, m.dExtraPriceScale, bCalcClusters)
'gdStopProfile 365
'gdStartProfile 366
                If rc = 0 Then
'gdStartProfile 460
                    LoadIndLabel Pane, Ind

                    If Ind.DataType = eINDIC_ProfileBars Or Ind.DataType = eINDIC_ProfileVolume Then
                        m.profileChartInfo.ColorScheme = Ind.ProfileColorScheme
                        m.profileChartInfo.styleTPO = Ind.ProfileStyleTPO
                        
                        If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
                            m.profileChartInfo.RTOnOff = 1
                        ElseIf m.profileChartInfo.RTOnOff <> -1 Then        '-1 means force init
                            m.profileChartInfo.RTOnOff = 0
                        End If
                        
                        'profile: show/hide, color, parm
                        For j = 0 To 7
                            gdSetNum m.profileChartInfo.glhItemShow, j, Ind.ProfileShowHide(j)
                            gdSetNum m.profileChartInfo.glhItemColor, j, Ind.ProfileColor(j)
                            gdSetNum m.profileChartInfo.gdhItemParm, j, 0
                        Next
                        If Ind.DataType = eINDIC_ProfileVolume Then
                            gdSetNum m.profileChartInfo.glhItemShow, 8, Ind.ProfileShowHide(8)
                            gdSetNum m.profileChartInfo.glhItemColor, 8, Ind.ProfileColor(7)
                            gdSetNum m.profileChartInfo.gdhItemParm, 8, 0
                        End If
                        
                        'percent for TPO/Volume value area
                        '1st 4 items in parm array are ignored by grapheng.dll (exist only to align with other arrays)
                        gdSetNum m.profileChartInfo.gdhItemParm, ePCStruct_TPO_VA, Ind.ProfileParm(ePCStruct_TPO_VA)
                        gdSetNum m.profileChartInfo.gdhItemParm, ePCStruct_Volume_VA, Ind.ProfileParm(ePCStruct_Volume_VA)

                    End If
'gdStopProfile 460
                    If m.Tree.Key(Ind.geIndId) = "PRICE" Then
                        Ind.isPriceInd = 1
                        Set PriceInd = Ind
                    End If
                    
                    If m.eRedoMode > eRedo3_Settings Then
                        Ind.CheckIndAlert
                    End If
                    
                    If nIndAlert <> 1 Then
                        If Not Ind.AlertObject Is Nothing Then
                            Ind.geIndId = Parse(m.Tree.Key(i), " ", 2)
                            If Ind.AlertObject.Active Then
                                nIndAlert = 1
                            Else
                                nIndAlert = -1
                            End If
                        End If
                    End If
                    
                    If Ind.Display Then
                        If Ind.DataType <> eINDIC_BooleanArray And Ind.DisplayType <> eINDIC_NoStyle Then bHasValidStyle = True
                        If FirstVisibleInd Is Nothing Then
                            Set FirstVisibleInd = Ind
                        End If
                    End If
                End If
'gdStopProfile 366
            End If
        End If
        
        If rc = 0 Then
            If Not Ind Is Nothing Then
                If Ind.DisplayType = eINDIC_ClusterPrice Then
                    If Not Ind.Display Then
                        Pane.SplitPaneClear
                    ElseIf Ind.TrueRangeFlag <> 3 Then      '1=cluster & swing, 2=cluster only, 3=swing only
                        Pane.SetSplitPaneCluster Me, Ind
                    End If
                    If Ind.geIndId <> i Then Ind.geIndId = i
                End If
            End If
        Else
            Exit For    'Q&A - error message here?
        End If
    Next
'gdStartProfile 367
    
    'need to check last pane for dinapoli studies & validstyle
    If Not Pane Is Nothing Then
        If Not bHasValidStyle Then
            If Not FirstVisibleInd Is Nothing Then
                If FirstVisibleInd.DataType = eINDIC_Array Or Ind.DataType = eINDIC_Constant Then
                    FirstVisibleInd.DisplayType = eINDIC_Line
                    bHasValidStyle = True
                ElseIf FirstVisibleInd.DataType = eINDIC_BarData Then
                    FirstVisibleInd.DisplayType = eINDIC_OHLC
                    bHasValidStyle = True
                End If
                If bHasValidStyle Then
                    rc = FirstVisibleInd.geSetIndData(Me, nPaneID, FirstVisibleInd.geIndId, _
                        m.nFirstBar, m.nLastBar, nStartScreenX, ScreenEndX, m.dExtraPriceScale, bCalcClusters)
                End If
            End If
        End If
    End If

'gdStopProfile 367
'gdStartProfile 368
    'walk through indicator list a second time for overlayed indicators and Woodies pane settings
    For i = 1 To m.Tree.Count
        If m.Tree.NodeLevel(i) = 0 Then
            ' init new pane
            Set Pane = m.Tree(i)
            'have pane set split pane data as appropriate
            If Not Pane Is Nothing Then Pane.SetSplitPaneData Me, bHasWoodPane
            nPaneID = i
        ElseIf bHasOverlayed Then
            Set Ind = m.Tree(i)
            'handle overlayed indicators
            If Ind.Overlayed = True Then
                If m.iMiniChart > 0 Then Ind.IndLabelMode = eINDIC_Nothing
                rc = Ind.geSetIndData(Me, nPaneID, i, _
                    m.nFirstBar, m.nLastBar, nStartScreenX, ScreenEndX, m.dExtraPriceScale, bCalcClusters)
                If rc = 0 Then
                    LoadIndLabel Pane, Ind
                    If m.Tree.Key(Ind.geIndId) = "PRICE" Then
                        Ind.isPriceInd = 1
                        Set PriceInd = Ind
                    End If
                End If
            End If
        End If
        If rc <> 0 Then Exit For    'Q&A - error message here?
    Next
'gdStopProfile 368
'gdStartProfile 369
    
    'add bad ticks indicator last to maintain tree index & indicator ID relationship
    i = m.Tree.Count + 1
    If m.BadTicksInd.Display = True Then
        rc = m.BadTicksInd.geSetIndData(Me, PricePane.gePaneId, i, _
            m.nFirstBar, m.nLastBar, nStartScreenX, ScreenEndX, m.dExtraPriceScale, bCalcClusters)
        If rc = 0 Then LoadIndLabel PricePane, m.BadTicksInd
    End If
    
    'add system name as last indicator label
    strSysName = ""
    If m.nShowTrades = 1 And Not m.System Is Nothing _
        And Not PricePane Is Nothing And Not PriceInd Is Nothing Then
        strSysName = m.System.SystemName
    ElseIf m.Form.IsInGameMode And m.Form.GameReplayMode <> eGDReplayMode_Sync Then
        strSysName = m.strSystemName
'    ElseIf m.nShowTrades = 2 And TradeAccountID > 0 Then
'        strSysName = "Trade Account: " & AccountNameFromID(TradeAccountID)
    End If
    
    If Len(strSysName) > 0 Then
        LoadIndLabel PricePane, PriceInd, strSysName
    Else
        i = m.Annots.Index("system name")
        If i > 0 Then
            Set Annot = m.Annots(i)
            Annot.geRemoveAnnotation m.geChartObj
            m.Annots.Remove i
            Set Annot = Nothing
        End If
    End If
    
    'average entry line
    If Not m.Annots Is Nothing Then
       If m.nShowAvgEntryLine = 1 And m.dAvgEntry > 0# And m.nShowTrades = 2 Then
            If m.AvgEntryAnnot Is Nothing Then
                Set m.AvgEntryAnnot = New cAnnotation
                m.AvgEntryAnnot.CreateNew Me, eANNOT_HorzLine, m.Tree("PRICE PANE").gePaneId, 0, 0, 0, 0, 0, 0
                i = m.Annots.Add(m.AvgEntryAnnot)
                m.AvgEntryAnnot.Prop("AnnotKey") = m.Annots.Key(i)
                m.AvgEntryAnnot.eUsage = eANNOT_AvgEntryLine
                'properties
                '1=color, 2=pen style, 3=textPosition, 4=font, 5=font size, 6=font style, 7=font underline
                m.AvgEntryAnnot.PreIndicator = 0            'aardvark 6313
                m.AvgEntryAnnot.SkipHitTestFlag = 1
                m.AvgEntryAnnot.Color = Val(Parse(m.strAvgEntryProp, "|", 1))
                m.AvgEntryAnnot.Style = Val(Parse(m.strAvgEntryProp, "|", 2))
                m.AvgEntryAnnot.geTextAlign = Val(Parse(m.strAvgEntryProp, "|", 3))
                m.AvgEntryAnnot.Prop("Font") = Parse(m.strAvgEntryProp, "|", 4)
                m.AvgEntryAnnot.Prop("FontSize") = Val(Parse(m.strAvgEntryProp, "|", 5))
                m.AvgEntryAnnot.Prop("FontStyle") = Val(Parse(m.strAvgEntryProp, "|", 6))
                m.AvgEntryAnnot.Prop("FontUnderline") = Val(Parse(m.strAvgEntryProp, "|", 7))
            End If
            m.AvgEntryAnnot.dDate(1) = gdGetNum(m.geChartStruct.gdhDate, 0)
            m.AvgEntryAnnot.dDate(2) = gdGetNum(m.geChartStruct.gdhDate, gdGetSize(m.geChartStruct.gdhDate) - 1)
            m.AvgEntryAnnot.Y(1) = m.dAvgEntry
            m.AvgEntryAnnot.Y(2) = m.dAvgEntry
            
            ' 06/29/2009 DAJ: Want to display this in trading units instead of the decimal...
            'm.AvgEntryAnnot.Text = "AvgEntry " & Str(m.dAvgEntry)
            m.AvgEntryAnnot.Text = "AvgEntry " & m.Bars.PriceDisplay(m.dAvgEntry)
        ElseIf Not m.AvgEntryAnnot Is Nothing Then
            'order bar is off or avgEntryLine is off or AvgEntry value is zero --> delete annot
            RemoveAnnots True, , eANNOT_AvgEntryLine
            Set m.AvgEntryAnnot = Nothing
        End If
    End If
    
    'handle flags for split pane
    
    
    'walk through tree adding annotation to grapheng.dll
    Dim iEwaveAnnot&, iOptionAnnotColor&
    
    If g.ChartGlobals.nChartBackColor = 0 Then
        iOptionAnnotColor = RGB(0, 254, 0)
    ElseIf ChartBkIsLight Then      '5226
        iOptionAnnotColor = vbBlue
    Else
        iOptionAnnotColor = vbWhite
    End If
    
    'make sure cluster zone rect is first in list so will not overdraw other annots
    Set Annot = Annots(kClusterZoneRect)
    If Not Annot Is Nothing Then
        If Not Annots(1) Is Annot Then i = Annots.Move(kClusterZoneRect, 1, eTREE_Myself)
    End If
    
    For i = 1 To Annots.Count
        Set Annot = Annots(i)
        If Not Annot Is Nothing Then
            If Annot.gePaneId > 0 Then
                nPaneID = Annot.gePaneId
            Else
                nPaneID = m.Tree.Index(Annot.Pane)
            End If
            If Annot.eUsage = eANNOT_OptionInfo Then
                If Annot.Color <> iOptionAnnotColor Then Annot.Color = iOptionAnnotColor
            End If
            If nPaneID > 0 Then
                If bExport Then
                    Annot.geExportAnnot Me, nPaneID, i, nStartScreenX
                ElseIf Annot.eType = eANNOT_ElliotEwave Then
                    iEwaveAnnot = i
                Else
                    rc = Annot.geAddAnnotation(Me, nPaneID, i)
                    If rc = 0 Then Annot.CheckAnnotAlert False
                    Annot.Hide g.ChartGlobals.nHideAnnotations
                    'double check annot ID as things can get moved around
                    If Annot.geAnnId <> i Then Annot.geAnnId = i
                End If
            End If
        End If
    Next
    
    'prototype for EWI Ewaves
    If iEwaveAnnot = 0 Then
        If m.nSymbolID = 50 Then
            Set Annot = New cAnnotation
            If Not m.aEWave Is Nothing Then
                Annot.CreateNew Me, eANNOT_ElliotEwave, PricePane.gePaneId, m.aEWave.ArrayHandle, 0, 0, 0
                m.Annots.Add Annot
            End If
        End If
    Else
        Set Annot = Annots(iEwaveAnnot)
        If Not Annot Is Nothing Then
            If m.nSymbolID = 50 Then
                rc = Annot.geAddAnnotation(Me, PricePane.gePaneId, iEwaveAnnot)
            Else
                rc = Annot.geRemoveAnnotation(Me.geChartObj)        'theoretically should not get here
            End If
        End If
    End If
    
    'add one alert icon for all chart condition alerts
    If Not g.RealTime.Active Or (g.RealTime.Active And Len(g.strActiveDraw) = 0) Then
        'If user is drawing annot with multiple clicks when real time is on, adding the
        'bell as last icon will throw off annots ID and cause multiple problems.
        If nIndAlert <> 0 Then
            Set Annot = New cAnnotation
            If nIndAlert = 1 Then
                Annot.Prop("ImageType") = eCNI_Bell
            Else
                Annot.Prop("ImageType") = eCNI_Bell_Gray
            End If
            Annot.CreateNew Me, eANNOT_BellAlert, m.Tree.Index("PRICE PANE"), 0, -1, 0, -2
            m.Annots.Add Annot
        End If
    End If
        
    UpdateOnlineOrder Nothing, True
    
    'pattern for profit
    Dim oPFP As cPatternProfit
    
    Set oPFP = m.Form.PatternProfitObj
    If oPFP Is Nothing Then
        If Not m.CompositeIndicators Is Nothing Then
            For i = 1 To m.CompositeIndicators.Count
                If Not m.CompositeIndicators(i) Is Nothing Then m.CompositeIndicators(i).geResetData Me
            Next
            m.CompositeIndicators.Clear
        End If
    ElseIf oPFP.CompositeCount > 1 Then
        If m.CompositeIndicators Is Nothing Then
            Set Ind = New cIndicator
        ElseIf m.CompositeIndicators.Count = 0 Then
            Set Ind = New cIndicator
        Else
            Set Ind = m.CompositeIndicators(1)
        End If
        If Not Ind Is Nothing And Not PricePane Is Nothing Then
            If PricePane.gePaneId > 0 Then
                Ind.LoadCompositePFP Me, m.Form.PatternProfitObj, PricePane.gePaneId, Tree.Count + 1000, nStartScreenX, ScreenEndX
            End If
        End If
    ElseIf Not m.CompositeIndicators Is Nothing Then
        For i = 1 To m.CompositeIndicators.Count
            If Not m.CompositeIndicators(i) Is Nothing Then m.CompositeIndicators(i).geResetData Me
        Next
        m.CompositeIndicators.Clear
    End If
    
    If Not g.RealTime.Active Then
        geForceRecalc       'aardvark 3685 (don't do this if real time is on)
    End If
    
    'check for price pane hide flag
    If Not PricePane Is Nothing Then
        If m.bHidePriceIndicators = True Then
            PricePane.PricePaneFlag = -1
        Else
            PricePane.PricePaneFlag = 1
        End If
    End If
    
    ' set scale back to manual
    If m.bTempAutoScale Then
        Set Pane = m.Tree("Price Pane")
        If Not Pane Is Nothing Then
            Set Ind = m.Tree("PRICE")
            If Ind Is Nothing Then
                Pane.Scaling = ePANE_ScaleModeManual
                m.bTempAutoScale = False
            ElseIf FormIsLoaded(m.Form.Name) Then
                If m.Form.Visible Then
                    Pane.Scaling = ePANE_ScaleModeManual            'combined fix for 2365 & 3922
                    m.bTempAutoScale = False
                End If
            End If
        End If
    End If
    
    Set Pane = Nothing
    Set Ind = Nothing
    Set Annot = Nothing
                       
'gdStopProfile 369
    AddSpeed "LoadScreenData"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.LoadScreenData", eGDRaiseError_Raise

End Sub

Private Sub LoadIndLabel(Pane As cPane, Ind As cIndicator, Optional ByVal strSysName$)
On Error GoTo ErrSection:

    Dim rc&, i&, nPaneID&, nPos&
    Dim nBar&, nLastScreenDataBar&, dDate#, dPrice#, dDelta#
    Dim strText$, strKey$, strFmt$, strDelta$
    Dim Annot As cAnnotation
    Dim bDataExists As Boolean
    Dim geCoordInfo As coordinate_info
    Dim eLabelMode As eIndicatorLabelMode
    Dim bUsingDefault As Boolean

    'chart not initialized, no need to continue
    If m.geChartStruct.x_dataPointCount < 1 Then Exit Sub
    'no label, no need to continue
    If Len(Ind.ChartLabel) < 1 Then Exit Sub
    'get label mode
    eLabelMode = Ind.IndLabelMode
    
    If eLabelMode = eINDIC_DefaultSetting Then
        eLabelMode = g.ChartGlobals.eDefaultLabelMode
        bUsingDefault = True
    End If
        
    ' See if data exists
    nLastScreenDataBar = -1
    If m.nLastGoodDataBarX > 0 Then
        bDataExists = True
        nLastScreenDataBar = LastGoodDataBar(False, True)
    End If
                
    If strSysName = "" Then
        ' get text for label
        If bDataExists Then
            If Ind.BadTicksFlag = True Then
                strText = Ind.ChartLabel     'ask Tim
            ElseIf Ind.isPriceInd = 1 Then
                strText = GetIndicatorDataLabel(nLastScreenDataBar, Ind.geIndId, 0, eLBL_Chart)
                If Not bUsingDefault And (eLabelMode = eINDIC_NoLabel Or eLabelMode = eINDIC_Nothing) Then strText = ""     '4593
            ElseIf eLabelMode = eINDIC_NoLabel Or eLabelMode = eINDIC_Nothing Then
                strText = ""
            Else
'gdStartProfile 471
                strText = GetIndicatorDataLabel(nLastScreenDataBar, Ind.geIndId, 0, eLBL_Chart)
'gdStopProfile 471
            End If
        End If
        strKey = CStr(Ind.geIndId) & " idx label"
    Else
        strKey = "system name"
        If m.bAutoTrade Then
            strText = "*" & strSysName
        Else
            strText = strSysName
        End If
    End If
                   
    nPaneID = Pane.gePaneId
    
    'TODO: should reset string pane property of annotation here too
    Set Annot = m.Annots(strKey)
    If Annot Is Nothing Then
        If eLabelMode = eINDIC_NoLabel Or eLabelMode = eINDIC_Nothing Or Len(strText) = 0 Then
            If Ind.isPriceInd = 1 And bUsingDefault Then
                'don't exit
            Else
                Exit Sub    'no label needed
            End If
        End If
        ' create new annotation (grapheng.dll has code to position labels)
        Set Annot = New cAnnotation
        Annot.CreateNew Me, eANNOT_Icon, Pane.gePaneId, _
            -1, -1, 0, 0, Ind.Color, PEGAT_NOSYMBOL, _
            strText, eANNOT_IndicatorLabel
        Annot.geIndId = Ind.geIndId
        Annot.gePaneId = nPaneID
        m.Annots.Add Annot, strKey
    ElseIf Ind.Display = False Or eLabelMode = eINDIC_NoLabel Or eLabelMode = eINDIC_Nothing Then
        '(but always show price and delta in the price label)
        If m.Tree.Key(Ind.geIndId) <> "PRICE" Then
            Annot.geRemoveAnnotation m.geChartObj
            m.Annots.Remove strKey
            Set Annot = Nothing
        Else
            Annot.Text = strText
            Annot.Color = Ind.Color
            Annot.gePaneId = nPaneID
        End If
    Else
        'reset text & color
        Annot.Text = strText
        Annot.Color = Ind.Color
        Annot.gePaneId = nPaneID
    End If
        
    'check label mode and modify text if necessary
    If Not Annot Is Nothing Then
        'change text style for system name
        'If strSysName <> "" Then
        '    Annot.Style = 3
        'End If
        '(but always show price and delta in the price label)
        If m.Tree.Key(Ind.geIndId) <> "PRICE" Then
            If eLabelMode = eINDIC_NoValue Or eLabelMode = eINDIC_scale Then
                nPos = InStr(1, strText, "=")   'desc only in label
                If nPos > 0 Then
                    Annot.Text = Trim(Left(strText, nPos - 1))
                End If
            ElseIf eLabelMode = eINDIC_OnlyValue Then
                nPos = InStr(1, strText, "=")   'value only in label
                If nPos > 0 Then
                    Annot.Text = Trim(Right(strText, Len(strText) - nPos))
                End If
            End If
        End If
    End If
                   
ErrExit:
    Set Annot = Nothing
    Exit Sub
    
ErrSection:
    Set Annot = Nothing
    RaiseError "cChart.LoadIndLabel", eGDRaiseError_Raise

End Sub

Private Sub SetDateStyle()
On Error GoTo ErrSection:

'enum type from grapheng header file
'typedef enum{
'0    e_month,
'1    e_day,
'2    e_year,
'3    e_week,
'4    e_decade,
'5    e_hour,
'6    e_autoCoarse,
'7    e_autoFine,
'8    e_minichart,
'9    e_none
'} DATE_STYLE;
'
'nVertGrid: 0=coarse, 1=fine, 2=none
    
    Dim eLabelPrd As eLabelPeriod
    
    m.geChartStruct.crossOverTime = 0
    
    ' Date labels
    If m.iMiniChart > 0 Then
        eLabelPrd = eLBL_Each
        m.geChartStruct.DateStyle = 8
        Exit Sub
    End If
    ' determine types of labels to use (period)
        Select Case GetPeriodType(m.nPeriodicity)
        Case ePRD_Days
            If m.nVertGrid = 1 Then    'fine
                eLabelPrd = eLBL_Week
                m.geChartStruct.DateStyle = 3
            Else                        'coarse or none
                eLabelPrd = eLBL_Month
                m.geChartStruct.DateStyle = 0
            End If
        Case ePRD_Weeks, ePRD_EodVol
            If m.nVertGrid > 0 Then
                eLabelPrd = eLBL_Month
                m.geChartStruct.DateStyle = 0
            Else
                eLabelPrd = eLBL_Year
                m.geChartStruct.DateStyle = 2
            End If
        Case ePRD_Months
            eLabelPrd = eLBL_Year
            m.geChartStruct.DateStyle = 2
        Case ePRD_Quarters
            If m.nVertGrid > 0 Then
                eLabelPrd = eLBL_Year
                m.geChartStruct.DateStyle = 2
            Else
                eLabelPrd = eLBL_Decade
                m.geChartStruct.DateStyle = 4
            End If
        Case ePRD_Years
            eLabelPrd = eLBL_Decade
            m.geChartStruct.DateStyle = 4
        Case ePRD_EodPF, ePRD_EodKagi, ePRD_EodRenko
            eLabelPrd = eLBL_Month
            m.geChartStruct.DateStyle = 0
        'Case ePRD_IntPF, ePRD_IntKagi, ePRD_IntRenko
        '    eLabelPrd = eLBL_Day
        '    m.geChartStruct.DateStyle = 1
        Case Else
            m.geChartStruct.crossOverTime = m.Bars.Prop(eBARS_CrossoverTime)
            ' if minutes per bar >= 30 or if not intraday, use week/day for coarse/fine
            If Not IsIntraday(m.nPeriodicity) Or (GetPeriodType(m.nPeriodicity) = ePRD_Minutes And m.nPeriodicity >= ePRD_Minutes + 30) Or GetPeriodType(m.nPeriodicity) = ePRD_SMP Then
                'Intraday >= 30 minutes or non-intraday volume per bar
                If m.nVertGrid = 0 Then
                    eLabelPrd = eLBL_Week
                    m.geChartStruct.DateStyle = 3
                Else
                    eLabelPrd = eLBL_Day
                    m.geChartStruct.DateStyle = 1
                End If
            ElseIf GetPeriodType(m.nPeriodicity) = ePRD_Minutes Then
                If IsProfileChart() Then
                    eLabelPrd = eLBL_Day
                    m.geChartStruct.DateStyle = 1
                Else
                    ' else use day/hour for coarse/fine for minute bars (< 30 minutes)
                    If m.nVertGrid = 0 Then
                        eLabelPrd = eLBL_Day
                        m.geChartStruct.DateStyle = 1
                    Else
                        eLabelPrd = eLBL_Hour
                        m.geChartStruct.DateStyle = 5
                    End If
                End If
            Else
'JM: 09-27-2011
'For non-minute-bar intraday:
'- "Fine" = a label at every 10 bars
'- "Coarse" = date at session break (just like "Coarse" for 5-minute bars)
'- "None" = same (nothing displayed)
                Select Case m.nVertGrid
                    Case 0
                        eLabelPrd = eLBL_Day
                        m.geChartStruct.DateStyle = 1           '6425 - coarse grid line
                    Case 1
                        eLabelPrd = eLBL_Hour
                        m.geChartStruct.DateStyle = 6           '6425 - fine grid line
                    Case Else
                        eLabelPrd = eLBL_Day
                        m.geChartStruct.DateStyle = 1           '6436 - no grid line or unknown (use same label style as coarse grid line)
                End Select
            End If
        End Select
        
        'check & set time zone for display
        If g.bShowInLocalTimeZone Then
            gdSetStr m.geChartStruct.gshTimeZoneInf, 0, ""
        ElseIf Not m.Bars Is Nothing Then
            gdSetStr m.geChartStruct.gshTimeZoneInf, 0, m.Bars.Prop(eBARS_ExchangeTimeZoneInf)
        End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.SetDateStyle", eGDRaiseError_Raise

End Sub

Public Sub geSetPrintMargin(ByVal Top#, ByVal Left#, ByVal Bottom#, ByVal Right#)
On Error GoTo ErrSection:

    gdSetNum m.geChartStruct.gdhPrintMargin, 0, Top
    gdSetNum m.geChartStruct.gdhPrintMargin, 1, Left
    gdSetNum m.geChartStruct.gdhPrintMargin, 2, Bottom
    gdSetNum m.geChartStruct.gdhPrintMargin, 3, Right
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.geSetPrintMargin", eGDRaiseError_Raise

End Sub
'MJM - [END] functions for grapheng.dll added here

Private Sub LoadSystem()
On Error GoTo ErrSection

    Dim bOrders As Boolean

    ' when starting, don't do systems yet (will be done from tmrMain after all is initialized)
    If g.bStarting Or g.bUnloading Then Exit Sub
    
    If Not m.Form.GameMode Is Nothing Then
        If m.Form.GameMode.CustomOrders > 0 Then
            m.nSystemID = m.Form.GameMode.GameStrategyID
            m.strSystemName = m.Form.GameMode.GameStrategyName
        Else
            m.Form.GameMode.GameStrategyID = m.nSystemID
            m.Form.GameMode.GameStrategyName = m.strSystemName
        End If
    End If

    ' sync up the symbol name and ID (in case either has changed)
    If m.nSystemID > 0 Then
        If Not SyncSystemInfo(m.strSystemName, m.nSystemID) Then
            ' if can't find the system, turn "show trades" off (so won't keep trying)
            If m.nShowTrades = 1 Then m.nShowTrades = 0
            m.nSystemID = 0
        End If
    End If
    
    KillSystemReports False
    m.dNextBarDate = 0
    If m.Bars(eBARS_Close, 0) = kNullData Or m.nSystemID = 0 Then
        Set m.System = Nothing
        frmOrderTracker.ShowOrders m.Form.hWnd, "", "", False
    Else
        ' load the system
        If m.System Is Nothing Then
            Set m.System = New cSystem
        ElseIf m.System.SystemNumber <> m.nSystemID Then
            Set m.System = New cSystem
        End If
        m.dNextBarDate = m.System.LoadForChart(m.nSystemID, m.Bars, m.Form.hWnd, (m.Form.GameReplayMode <> 0))
        AddSpeed "System Loaded"
        ' backtest the system to get trades
        m.System.Test True
        AddSpeed "System BackTest"
        ' run orders for next bar (except for non-Forex indices, since can take so long for $DJIA)
        If Left(Me.Symbol, 1) = "$" And Mid(Me.Symbol, 5, 1) <> "-" Then
            'frmOrderTracker.ShowOrders m.Form.hWnd, "", "", False
            frmOrderTracker.ShowOrders m.Form.hWnd, Me.Symbol, "(orders are not created for indices)", False
        Else
            bOrders = m.System.NextBarReport(eGDNextBarMode_UseHwnd)
            AddSpeed "System NextBarReport"
            frmOrderTracker.ShowOrders m.Form.hWnd, Me.Symbol, m.System.SystemName, bOrders
        End If
        ReadTrades
        AddSpeed "Trades read"
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.LoadSystem", eGDRaiseError_Raise

End Sub

Public Property Get SystemID() As Long
    SystemID = m.nSystemID
End Property

Public Property Let SystemID(ByVal nSystemID As Long)
On Error GoTo ErrSection:

    If nSystemID <> m.nSystemID Then
        m.strSystemName = ""
        m.nSystemID = nSystemID
        If nSystemID = 0 Then
            Set m.System = Nothing
        Else
            SyncSystemInfo m.strSystemName, m.nSystemID
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.SystemID.Let", eGDRaiseError_Raise

End Property

Public Sub ShowSystemReport(ByVal bNextBarOrders As Boolean)
On Error GoTo ErrSection:

    Dim nPeriodType&
    Dim strTradesFile As String

    If m.Form.IsInGameMode Then
        If m.System Is Nothing Or m.nSystemID = 0 Or _
           m.Form.GameMode.GameStrategyID = 0 Or _
           m.Form.GameMode.GameStrategyID = kGameModeSysID Then
                m.Form.GenerateGameReport
        ElseIf m.nSystemID = m.Form.GameMode.GameStrategyID Then
            m.System.TestResults True
        End If
    ElseIf m.nShowTrades = 1 And Not m.System Is Nothing Then
        ' TLB 6/11/2012: backtesting is NOT valid for Renko, Kagi, and P&F
        ' (since high/low of bars does not reflect actual trading, which makes backtested fills invalid)
        nPeriodType = m.Bars.Prop(eBARS_PeriodType)
        If nPeriodType = ePRD_IntRenko Or nPeriodType = ePRD_IntKagi Or nPeriodType = ePRD_IntPF Or _
                nPeriodType = ePRD_EodRenko Or nPeriodType = ePRD_EodKagi Or nPeriodType = ePRD_EodPF Then
            InfBox "Backtesting is not valid for Point & Figure,| Kagi or Renko bars (since these types| of bars do not reflect the actual trades).", "!", , "Performance Report"
        ElseIf bNextBarOrders Then
            m.System.ShowNextBarReport eGDNextBarMode_UseHwnd
            frmOrderTracker.MarkAsRead m.Form.hWnd
        Else
            m.System.TestResults True
        End If
    ElseIf m.nShowTrades = 2 Then
        strTradesFile = TradesFileForAccount(False)     '4053
        If FileExist(strTradesFile) Then
            If m.System Is Nothing Then
                Set m.System = New cSystem
                m.System.Load 56  'this is the sample turtle strategy
            ElseIf m.System.SystemNumber <> 56 Then
                m.System.Load 56
            End If
            
            m.System.hWnd = m.Form.hWnd
            m.System.ShowReports , strTradesFile, True
        Else
            InfBox "No trades history for " & Symbol & " cannot show report.", "I", , "Performance Report"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.ShowSystemReport", eGDRaiseError_Raise

End Sub

Private Sub KillSystemReports(Optional ByVal bRemoveFromOrderTracking As Boolean = True)
On Error GoTo ErrSection:

    Dim hWnd&
    
    If Not m.Form Is Nothing Then
        hWnd = m.Form.hWnd
        KillFile App.Path & "\Trades\S_" & Str(hWnd) & ".TXT"
        KillFile App.Path & "\Trades\NB_" & Str(hWnd) & ".TXT"
        KillFile App.Path & "\Trades\RB_" & Str(hWnd) & ".TXT"
        
        If Not g.bUnloading Then
            If Not m.frmOrdersPopup Is Nothing Then
                If m.frmOrdersPopup.Visible Then
                    m.frmOrdersPopup.Hide
                    DoEvents
                End If
                Set m.frmOrdersPopup = Nothing
            End If
            
            If bRemoveFromOrderTracking Then
                frmOrderTracker.ShowOrders hWnd, "", "", False
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.KillSystemReports", eGDRaiseError_Raise

End Sub

Public Sub EnablePerformanceButton()
On Error GoTo ErrSection:

    Dim bEnable As Boolean
    
    If m.Form Is ActiveChart Then
        If m.Form.IsInGameMode Then
            bEnable = m.Form.GameRptEnable       'aardvark 2606
        ElseIf m.nShowTrades = 1 And Not m.System Is Nothing Then
            bEnable = True
        ElseIf m.nShowTrades = 2 Then
            bEnable = True
        End If
'        m.tbToolbar.Tools("ID_Performance").Enabled = bEnable
        frmMain.tbToolbar.Tools("ID_Performance").Enabled = bEnable
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.EnablePerformanceButton", eGDRaiseError_Raise
End Sub


Public Property Get RedoMode() As eChartRedoMode
    If m.eRedoMode <= 0 Then m.eRedoMode = eRedo9_ReloadData
    RedoMode = m.eRedoMode
End Property

Public Property Let RedoMode(ByVal eMode As eChartRedoMode)

    ' the redo mode can only be set higher (until cleared at end of GenerateChart)
    If eMode > m.eRedoMode Then
        If eMode = eRedo9_ReloadData Then
            eMode = eMode
        End If
        m.eRedoMode = eMode
    ElseIf m.eRedoMode <= 0 Then
        m.eRedoMode = eRedo9_ReloadData
    End If
    
    
'JM 08-03-2011: Do not set dirty page flag here, this code is hit way too often.
'    If eMode = eRedo3_Settings Then
''        'g.bSkipSetChartFocus = true means charts are being restored from chart page switching or program startup
'        If Not g.bSkipSetChartFocus And Not g.bStarting Then g.bDirtyChartPage = True
'    End If

End Property

Private Sub SetBadTicksIndicator()
On Error GoTo ErrSection:

    Dim nSize&, i&
    Dim hData As Long
    Dim PriceInd As cIndicator
    
    m.BadTicksInd.geResetData
    m.BadTicksInd.Display = False   'don't turn on until everything checks out okay
    
    If g.ChartGlobals.eBadTickShowMode = eHighlight_None Then Exit Sub
            
    Set PriceInd = m.Tree("PRICE")
    If PriceInd Is Nothing Then Exit Sub
            
    nSize = PriceInd.Bars.Size
    If nSize < 1 Then Exit Sub
    
    hData = m.BadTicksInd.Data.ArrayHandle
    gdClear hData, True
        
    With m.BadTicksInd
        .Color = g.ChartGlobals.nBadTickColor
        .Display = True
        If g.ChartGlobals.eBadTickShowMode = eHighLight_Bars Then
            .DisplayType = eINDIC_HighlightBars
        Else
            .DisplayType = eINDIC_HighlightMarkers
            .MarkerImage = g.ChartGlobals.eBadTickMarker
            .MarkerDir = g.ChartGlobals.eBadTickMarkerDir
            .MarkerFill = g.ChartGlobals.nBadTickMarkerFill
            If g.ChartGlobals.eBadTickShowMode = eHighlight_MarkersAbove Then
                .MarkerLoc = 0
            Else
                .MarkerLoc = 1
            End If
        End If
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.SetBadTicksIndicator", eGDRaiseError_Raise

End Sub

Private Sub ValidateHighlightBars()
On Error GoTo ErrSection:

    Dim IndToColor As cIndicator, Ind As cIndicator
    Dim BoolInd1 As cIndicator, BoolInd2 As cIndicator
    Dim aIndToColor As New cGdArray
    Dim aBoolData1 As Long, aBoolData2 As Long
    Dim hBoolRef As Long
    Dim bSkipPriceInd As Long
    Dim nIdx&, nSize&, nCount&, nDataSize&
    Dim i&, k&, j&, n&, t&
    Dim strMsg$

    'loop through tree and build array of indicators that have highlight bars
    For i = 1 To m.Tree.Count
        If m.Tree.NodeLevel(i) <> 0 Then
            Set Ind = m.Tree(i)
            If Not Ind Is Nothing Then
                If Ind.DataType = eINDIC_BooleanArray Then
                    nIdx = Ind.GetIndToColorIdx(m.Tree, i)
                    If nIdx >= 0 Then
                        Set IndToColor = m.Tree(nIdx)
                        If Not IndToColor Is Nothing Then
                            aIndToColor.Add nIdx
                            IndToColor.AddBoolRefInfo i
                        End If
                    End If
                End If
            End If
        End If
    Next

    aIndToColor.Sort eGdSort_Default Or eGdSort_DeleteDuplicates
    nSize = aIndToColor.Size
    If nSize < 1 Then Exit Sub

    'loop through array of indicators that have highlight bars and validate
    For i = 0 To nSize - 1
        Set IndToColor = m.Tree(aIndToColor(i))
        hBoolRef = IndToColor.BoolRefArray  'array of boolean indicators IDs used to color this indicator
        nCount = gdGetSize(hBoolRef)
        If nCount = 1 Then
            Set BoolInd1 = m.Tree(gdGetNum(hBoolRef, 0))
            BoolInd1.CanBeHighlight = 1
        ElseIf nCount > 1 Then
            For k = 0 To nCount - 1         'reset all highlight flags
                Set BoolInd1 = m.Tree(gdGetNum(hBoolRef, k))
                If Not BoolInd1 Is Nothing Then BoolInd1.CanBeHighlight = 1
            Next
            For k = 0 To nCount - 1
                Set BoolInd1 = m.Tree(gdGetNum(hBoolRef, k))
                If Not BoolInd1 Is Nothing Then
                    aBoolData1 = BoolInd1.Data.ArrayHandle
                    nDataSize = gdGetSize(aBoolData1)
                    For t = k + 1 To nCount - 1
                        Set BoolInd2 = m.Tree(gdGetNum(hBoolRef, t))
                        If Not BoolInd2 Is Nothing Then
                            If BoolInd2.Display = True And BoolInd2.MarkerPrompt = True Then
                                aBoolData2 = BoolInd2.Data.ArrayHandle
                                If nDataSize = gdGetSize(aBoolData2) And _
                                    BoolInd1.DisplayType = eINDIC_HighlightBars And _
                                    BoolInd1.Display = True Then
                                    For n = 0 To nDataSize
                                        If gdGetNum(aBoolData1, n) = 1 And gdGetNum(aBoolData2, n) = 1 Then
                                            strMsg = "The '" + BoolInd2.Name + "' highlight bars overlap other highlight bar(s)." + vbCrLf + "Would you like to use highlight markers instead?"
                                            If vbYes = MsgBox(strMsg, vbYesNo) Then
                                                BoolInd2.DisplayType = eINDIC_HighlightMarkers
                                                BoolInd2.MarkerImage = eCNI_Circle
                                                BoolInd2.MarkerFill = 1
                                                BoolInd2.CanBeHighlight = 0
                                                BoolInd2.MarkerPrompt = False
                                            End If
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    Next
                End If
            Next    'end loop of indicator's bool ref array
        End If
    Next    'end loop of indicators containing highlight bars

    If m.BadTicksInd.DisplayType <> eINDIC_HighlightBars Or m.BadTicksInd.Display = False Then Exit Sub
    
    'loop through and validate against bad ticks indicator
    m.BadTicksInd.CanBeHighlight = 1
    Set IndToColor = m.Tree("PRICE")
    hBoolRef = IndToColor.BoolRefArray
    nCount = gdGetSize(hBoolRef)
    If nCount > 1 Then
        aBoolData1 = m.BadTicksInd.Data.ArrayHandle
        nDataSize = gdGetSize(aBoolData1)
        For i = 0 To nCount - 1
            Set BoolInd2 = m.Tree(gdGetNum(hBoolRef, i))
            If Not BoolInd2 Is Nothing Then
                If BoolInd2.Display = True Then
                    aBoolData2 = BoolInd2.Data.ArrayHandle
                    If nDataSize = gdGetSize(aBoolData2) Then
                        For j = 0 To nDataSize
                            If gdGetNum(aBoolData1, j) = 1 And gdGetNum(aBoolData2, j) = 1 Then
                                BoolInd2.DisplayType = eINDIC_HighlightMarkers
                                BoolInd2.CanBeHighlight = 0
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        Next
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.ValidateHighlightBars", eGDRaiseError_Raise
End Sub

Public Property Get IndBadTicks() As cIndicator
    Set IndBadTicks = m.BadTicksInd
End Property

Public Property Get FirstBar() As Long
    FirstBar = m.nFirstBar
End Property

Public Property Get LastBar() As Long
    LastBar = m.nLastBar
End Property

Public Function ChartName(Optional ByVal bWithTemplate As Boolean = False, _
                    Optional ByVal bIgnoreCustomName As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strText$

    If Len(m.strChartName) > 0 And Not bIgnoreCustomName Then
        ' custom name
        strText = m.strChartName
    Else
        ' default name (symbol, period, template)
        If m.Bars Is Nothing Then
            strText = Symbol
        Else
            strText = Trim(m.Bars.Prop(eBARS_PeriodicityStr))
            If UCase(Mid(strText, 2, 7)) = " MINUTE" Then
                ' add spaces before single-digit minute bars (for sorting purposes)
                strText = "  " & strText
            ElseIf UCase(Mid(strText, 3, 7)) = " MINUTE" Then
                ' add a space before double-digit minute bars (for sorting purposes)
                strText = " " & strText
            End If
            If UCase(Right(strText, 7)) = " MINUTE" Then
                strText = Left(strText, Len(strText) - 6) & "min"
            End If
            'strText = Symbol & ", " & strText
            strText = Symbol & " (" & strText & ")"
        End If
        If bWithTemplate Then
            strText = strText & "  -  Template: " & TemplateApplied
        End If
    End If

    ChartName = strText
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ChartName", eGDRaiseError_Raise
End Function

Public Sub SetChartName(ByVal strChartName As String)
    
    m.strChartName = Trim(Replace(strChartName, "|", "/"))
    
End Sub

Public Sub AddReplayOrder(ByVal nQty#, ByVal strType$, ByVal dPrice#, _
    Optional ByVal dOrderDate# = 0#, _
    Optional ByVal strNonMarketOrder$ = "", _
    Optional Order As cPtOrder = Nothing)
On Error GoTo ErrSection:

    Dim i&, dDate#, dProfit#, strText$, nPos&, nEntryQty#, nExitQty#, strAction$

    If dOrderDate > 0 Then
        dDate = dOrderDate
    End If
    
    If dDate <= 0 Or nQty = 0 Then Exit Sub     'aardvark 1235 note: changed nQty, nEntryQty and nExitQty to doubles
    
    If m.Form.IsInGameMode And strType <> "M" And Len(strNonMarketOrder) > 0 Then
        i = LastGoodDataBar(False)
        m.dNextBarDate = m.Bars(eBARS_DateTime, i)
        If m.aOrders.Size = 0 Then m.aOrders.Add ""
        m.aOrders.Add strNonMarketOrder & vbTab & Str(m.dNextBarDate) & vbTab & Str(m.Bars(eBARS_High, i)) & vbTab & Str(m.Bars(eBARS_Low, i)) & vbTab & Str(m.Bars(eBARS_Close, i))
        If Not Order Is Nothing Then
            Order.OrderID = m.aOrders.Size - 1
        End If
        Exit Sub
    End If
    
    nPos = Position
    If strType = "M" Then
        ' TRADE: date EL price rule profit qty
        If nPos * nQty >= 0 Then
            ' entering or adding to same position
            nExitQty = 0
            nEntryQty = nQty
        ElseIf Abs(nQty) > Abs(nPos) Then
            ' reversing
            nExitQty = -nPos
            nEntryQty = nPos + nQty
        Else
            ' exiting
            nExitQty = nQty
            nEntryQty = 0
        End If
        ' add exit
        If nExitQty <> 0 Then
dProfit = OpenEquity(dPrice)
'm.dProfit = m.dProfit + dProfit
            If nExitQty > 0 Then
                strAction = "XS"
            Else
                strAction = "XL"
            End If
            m.Transactions.NumRecords = m.Transactions.NumRecords + 1
            m.Transactions.TblNum(eGDTransactionField_Date) = dDate
            m.Transactions.TblNum(eGDTransactionField_Quantity) = nExitQty
            m.Transactions.TblStr(eGDTransactionField_Action) = strAction
            m.Transactions.TblNum(eGDTransactionField_Price) = dPrice
            m.Transactions.TblStr(eGDTransactionField_Rule) = ""
            m.Transactions.TblNum(eGDTransactionField_Profit) = dProfit

' link exit with entry
For i = 0 To m.Transactions.NumRecords - 2
    If m.Transactions.TblNum(eGDTransactionField_Link, i) = -1 Then
        If m.Transactions.TblNum(eGDTransactionField_Quantity, i) = -nExitQty Then
            m.Transactions.TblNum(eGDTransactionField_Link, i) = m.Transactions.NumRecords - 1
            m.Transactions.TblNum(eGDTransactionField_Link) = i
            Exit For
        End If
    End If
Next
        
        End If
        ' add entry
        If nEntryQty <> 0 Then
            dProfit = 0
            If nEntryQty > 0 Then
                strAction = "EL"
            Else
                strAction = "ES"
            End If
            m.Transactions.NumRecords = m.Transactions.NumRecords + 1
            m.Transactions.TblNum(eGDTransactionField_Date) = dDate
            m.Transactions.TblNum(eGDTransactionField_Quantity) = nEntryQty
            m.Transactions.TblStr(eGDTransactionField_Action) = strAction
            m.Transactions.TblNum(eGDTransactionField_Price) = dPrice
            m.Transactions.TblStr(eGDTransactionField_Rule) = ""
            m.Transactions.TblNum(eGDTransactionField_Profit) = dProfit
            m.Transactions.TblNum(eGDTransactionField_Link) = -1
        End If
        ''m.nPosition = m.nPosition + nQty
        ''CalcTradeStats dDate
        m.Transactions.CalculateStats dDate
        m.Transactions.CalculateEquityFilter
    Else
        ' ORDER: 1 EL ruleID S price qty
        If nQty < 0 Then
            If nQty + nPos < 0 Then
                strAction = "ES"
            Else
                strAction = "XL"
            End If
        Else
            If nQty + nPos > 0 Then
                strAction = "EL"
            Else
                strAction = "XS"
            End If
        End If
        i = Me.LastGoodDataBar(False)
        m.dNextBarDate = m.Bars(eBARS_DateTime, i)
        strText = "1" & vbTab & strAction & vbTab & vbTab & strType & vbTab & Str(dPrice) & vbTab & Str(nQty)
        If m.aOrders.Size = 0 Then m.aOrders.Add ""
        m.aOrders.Add strText
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.AddReplayOrder", eGDRaiseError_Raise
End Sub

#If 0 Then
Private Sub CalcTradeStats(Optional ByVal dFromPositionBeforeDate As Double = 0)
On Error GoTo ErrSection:

    Dim i&, iStart&, nQty&, nPosition&, dPrice#, dProfit#, dSumEntry#
    Dim iEntry&, nEntryQtyUsed&, nEntryQty&, nExitQty&
    Dim strAction$
    
    If m.tblTrades.NumRecords = 0 Then Exit Sub
       
    If dFromPositionBeforeDate > 0 Then
        For i = m.tblTrades.NumRecords - 1 To 0 Step -1
            If TradeNum(eTRADE_Position, i) = 0 And TradeNum(eTRADE_Date, i) < dFromPositionBeforeDate Then
                iStart = i + 1
                Exit For
            End If
        Next
    End If
    
    For i = iStart To m.tblTrades.NumRecords - 1
        ' TRADE: Date Qty EL Price Rule Profit NetProfit Position AvgEntry
        TradeNum(eTRADE_Date, i) = RoundToSecond(TradeNum(eTRADE_Date, i)) '(need rounding due to adjustment made for sorting purposes)
        nQty = TradeNum(eTRADE_Qty, i)
        dProfit = TradeNum(eTRADE_Profit, i)
        strAction = TradeStr(eTRADE_Action, i)
        Select Case strAction
        Case "EL", "ES"
            If nPosition = 0 Then
                iEntry = i
                nEntryQtyUsed = 0
            End If
            dPrice = TradeNum(eTRADE_Price, i)
            dSumEntry = dSumEntry + dPrice * Abs(nQty)
            If dProfit = kNullData Then
                TradeNum(eTRADE_Profit, i) = 0
            End If
            
        Case "XL", "XS"
            If nQty = -nPosition Then
                ' if exiting entire position, clear out entry prices
                dSumEntry = 0
            Else
                ' drop off oldest entry prices
                nExitQty = Abs(nQty)
                Do While iEntry < i And nExitQty > 0
                    If Left(strAction, 1) = "E" Then
                        ' get unexited part of this entry
                        dPrice = TradeNum(eTRADE_Price, iEntry)
                        nEntryQty = Abs(TradeNum(eTRADE_Qty, iEntry)) - nEntryQtyUsed
                        If nExitQty < nEntryQty Then
                            ' if only exiting part of this entry
                            dSumEntry = dSumEntry - dPrice * nExitQty
                            nEntryQtyUsed = nEntryQtyUsed + nExitQty
                            Exit Do
                        End If
                        dSumEntry = dSumEntry - dPrice * nEntryQty
                        nEntryQtyUsed = 0
                        nExitQty = nExitQty - nEntryQty
                    End If
                    iEntry = iEntry + 1
                Loop
            End If
            If dProfit = kNullData Then
                ' calculate the profit for this trade
                TradeNum(eTRADE_Profit, i) = dProfit
            End If
        End Select
        nPosition = nPosition + nQty
        
        If nPosition = 0 Then
            TradeNum(eTRADE_AvgEntry, i) = 0
        Else
            TradeNum(eTRADE_AvgEntry, i) = dSumEntry / Abs(nPosition)
        End If
        
        TradeNum(eTRADE_NetProfit, i) = TradeNum(eTRADE_NetProfit, i - 1) + dProfit
        TradeNum(eTRADE_Position, i) = nPosition
    Next
  
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.CalcTradeStats", eGDRaiseError_Raise

End Sub

Private Property Get TradeStr(ByVal nField As eTradeField, Optional ByVal nRecord As Long = -99999) As String
    If nRecord = -99999 Then nRecord = m.tblTrades.NumRecords - 1
    'TradeStr = m.tblTrades.Item(nField, nRecord)
    TradeStr = gdGetTableString(m.hTradesHandle, nField, nRecord)
End Property

Private Property Let TradeStr(ByVal nField As eTradeField, Optional ByVal nRecord As Long = -99999, ByVal strData As String)
    If nRecord = -99999 Then nRecord = m.tblTrades.NumRecords - 1
    'm.tblTrades.Item(nField, nRecord) = strData
    gdSetTableStr m.hTradesHandle, nField, nRecord, strData
End Property

Private Property Get TradeNum(ByVal nField As eTradeField, Optional ByVal nRecord As Long = -99999) As Double
    If nRecord = -99999 Then nRecord = m.tblTrades.NumRecords - 1
    'TradeNum = m.tblTrades.Num(nField, nRecord)
    TradeNum = gdGetTableNum(m.hTradesHandle, nField, nRecord)
End Property

Private Property Let TradeNum(ByVal nField As eTradeField, Optional ByVal nRecord As Long = -99999, ByVal dData As Double)
    If nRecord = -99999 Then nRecord = m.tblTrades.NumRecords - 1
    'm.tblTrades.Num(nField, nRecord) = dData
    gdSetTableNum m.hTradesHandle, nField, nRecord, dData
End Property
#End If

Public Property Get ProfitLineBox() As Long
    ProfitLineBox = m.nProfitLineBox
End Property

Public Property Let ProfitLineBox(ByVal nFlag&)
    m.nProfitLineBox = nFlag
End Property

Public Function FirstTrueRangeClose() As Double
On Error GoTo ErrSection

    If m.nFirstBar > 0 Then
        FirstTrueRangeClose = Bars(eBARS_Close, m.nFirstBar - 1)
    Else
        FirstTrueRangeClose = Bars(eBARS_Close, m.nFirstBar)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.FirstTrueRangeClose", eGDRaiseError_Raise
End Function

Public Property Get AutoScale() As Boolean
On Error GoTo ErrSection:

    Dim Pane As cPane
    
    AutoScale = True '(default)
    Set Pane = m.Tree("Price Pane")
    If Not Pane Is Nothing Then
        If Pane.Scaling <> ePANE_ScaleModeAuto And Pane.Scaling <> ePANE_ScaleModeAutoPrice Then
            AutoScale = False
        End If
        Set Pane = Nothing
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.AutoScale_Get", eGDRaiseError_Raise
End Property

Public Property Let AutoScale(ByVal bAutoScale As Boolean)
On Error GoTo ErrSection:
    
    Dim Pane As cPane

    Set Pane = m.Tree("Price Pane")
    If Not Pane Is Nothing Then
        If bAutoScale Then
            Pane.Scaling = Pane.ScaleAutoLastUsed
        Else
            Pane.Scaling = ePANE_ScaleModeManual
        End If
        Set Pane = Nothing
        SyncToolbar
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.AutoScale_Let", eGDRaiseError_Raise
End Property

Public Property Get ChartLogFlag() As ePANE_LogFlag
On Error GoTo ErrSection:

    Dim Pane As cPane
    
    Set Pane = m.Tree("Price Pane")
    If Not Pane Is Nothing Then
        ChartLogFlag = Pane.PaneLogFlag
        Set Pane = Nothing
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.ChartLogFlag.Get", eGDRaiseError_Raise
End Property

Public Property Let ChartLogFlag(ByVal eFlag As ePANE_LogFlag)
On Error GoTo ErrSection:
    
    Dim Pane As cPane

    Set Pane = m.Tree("Price Pane")
    If Not Pane Is Nothing Then
        Pane.PaneLogFlag = eFlag
        Set Pane = Nothing
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.ChartLogFlag.Let", eGDRaiseError_Raise
End Property

Public Property Get PrintOrientation() As Long
    PrintOrientation = m.geChartStruct.prtOrientation
End Property

Public Property Let PrintOrientation(ByVal nOrientation&)
    m.geChartStruct.prtOrientation = nOrientation
End Property

Public Property Get LogoColor() As Long
    LogoColor = m.geChartStruct.LogoColor
End Property

Public Property Let LogoColor(ByVal nColor&)
    m.geChartStruct.LogoColor = nColor
End Property

Public Sub SyncGlobalAnnots(Annot As cAnnotation, Optional ByVal bWasMultiChart As Boolean = False)
On Error GoTo ErrSection:

    Dim bUpdateAll As Boolean

    If Not Annot Is Nothing Then
        If Annot.eUsage = eANNOT_UserAdded Then
            If Annot.MultiChartFlag Or bWasMultiChart Then
                bUpdateAll = True
            End If
        End If
    ElseIf bWasMultiChart Then
        bUpdateAll = True
    End If

    If bUpdateAll Then
        ' if something global changed, have all charts reload their annots and redraw
        TemplateSaveUserAnnots
        UpdateVisibleCharts eRedo2_ReloadAnnots, m.nSymbolID
    Else
        ' need to do this otherwise changes to non-global annotations can be lost
        TemplateSaveUserAnnots
        ''TemplateLoadUserAnnots
        GenerateChart eRedo1_Scrolled
    End If
    
    'set focus to chart so Ctrl-Z will work right away
    If Not m.Form Is Nothing Then SetFocus m.Form.pbChart.hWnd
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.SyncGlobalAnnots", eGDRaiseError_Raise
End Sub

Public Property Get DisableRT() As Boolean
    DisableRT = m.bDisableRT
End Property

Public Property Let DisableRT(ByVal bDisable As Boolean)
    m.bDisableRT = bDisable
End Property

Public Property Get AutoTrade() As Boolean
    AutoTrade = m.bAutoTrade
End Property

Public Property Let AutoTrade(ByVal bAutoTrade As Boolean)
    m.bAutoTrade = bAutoTrade
End Property

' just called by RestoreCharts in order to reuse an existing chart when loading a chart page
Public Sub ClearChartForReuse()
On Error GoTo ErrSection:

    m.strTemplate = ""
    m.nSymbolID = 0
    m.nTradeSymbolID = 0
    m.nShowTrades = 0
    m.eTypeOfChart = eTypeChart_Undefined
    m.bHideExits = False
    
    m.strLinkedInputs = ""
    m.strSpreadSymbols = ""
    m.strExtData = ""
    m.Form.WindowLink.SymbolColor = 0           '4693
    m.Form.WindowLink.PeriodColor = 0
    
    RemoveAlerts True                '4462

'4699 - this causes intermitten crashes due to the chart object getting closed out below (don't need to do this)
'    geResetPanes
    
    ResetSplitPane
    geCloseChart m.geChartObj
    m.geChartObj = 0
    m.nShowSplitPane = 0
    m.nBlankBars = m.nBlankBarsSave
    m.bProfileChart = False
    
    Set m.LastCreatedAnnot = Nothing
    Set m.LastEditedAnnot = Nothing
    
    If Not m.Annots Is Nothing Then m.Annots.Clear
    If Not m.aOrders Is Nothing Then m.aOrders.Clear
    If Not m.OnlineOrders Is Nothing Then m.OnlineOrders.Clear
    If Not m.Tree Is Nothing Then m.Tree.Clear
    
    If Not m.Form Is Nothing Then m.Form.ClearFormForReuse

'    If m.Form.DetachStatus = eNotDetached Then m.Form.ShowHideToolbar False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.ClearChartForReuse", eGDRaiseError_Raise

End Sub


Public Sub DeleteGameOrders()
On Error GoTo ErrSection:

    If m.Form.IsInGameMode Then
        m.aOrders.Clear
        m.aOrders.Size = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.DeleteGameOrders", eGDRaiseError_Raise

End Sub

Public Sub DeleteGameTrades(ByVal dDataTime#)
On Error GoTo ErrSection:
    
    Dim i&, j&, nCount&
    
    If dDataTime = 0 Then
        m.Transactions.NumRecords = 0
    ElseIf dDataTime > 0 Then
        nCount = m.Transactions.NumRecords
        For i = nCount - 1 To 0 Step -1
            If m.Transactions.TblNum(eGDTransactionField_Date, i) >= dDataTime Or m.Transactions.TblNum(eGDTransactionField_Date, i) <= 0 Then
                j = m.Transactions.TblNum(eGDTransactionField_Link, i)
                If j >= 0 And j < m.Transactions.NumRecords Then
                    m.Transactions.TblNum(eGDTransactionField_Link, j) = -1 '(clear linked entry/exit)  -aardvark 4767
                End If
                m.Transactions.RemoveRecords i
            End If
        Next
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.DeleteGameTrades", eGDRaiseError_Raise

End Sub

Public Sub RemoveAlerts(Optional ByVal bTerminating As Boolean = False)
On Error GoTo ErrSection:

    Dim Ind As cIndicator
    Dim Annot As cAnnotation
    Dim i&
    
    If g.bStarting Or g.bUnloading Then Exit Sub
    
    m.bPriceAlertsChecked = False
    
    If Not Tree Is Nothing Then
        For i = Tree.Count To 0 Step -1
            If Tree.NodeLevel(i) > 0 Then
                Set Ind = Tree(i)
                If Not Ind Is Nothing Then Ind.UpdateAlert 0, False
            End If
        Next
    End If
    
    If Not m.Annots Is Nothing Then
        For i = m.Annots.Count To 0 Step -1
            Set Annot = m.Annots(i)
            'bTerminate is true when: form unload, chart terminate or chart cleared for reuse
            If Not Annot Is Nothing Then Annot.UpdateAlert 0, , , bTerminating
        Next
    End If
    
    If FormIsLoaded("frmAlertsSetup") Then
        If frmAlertsSetup.Visible Then frmAlertsSetup.LoadGrid
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "cChart.RemoveAlerts"
    
End Sub

Public Sub ClearAnnotMultiFlag()
On Error GoTo ErrSection:

    Dim i&, strFile$
    Dim Annot As cAnnotation
    
    If m.Annots.Count = 0 Then Exit Sub
    
    For i = 1 To m.Annots.Count
        Set Annot = m.Annots(i)
        If Not Annot Is Nothing Then
            If Annot.MultiChartFlag Then
                Annot.MultiChartFlag = False
            End If
        End If
    Next
    
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\^" & AnnotSymbol & ".ANO"
    KillFile strFile
    
    TemplateSaveUserAnnots
    UpdateVisibleCharts eRedo2_ReloadAnnots

    Exit Sub
    
ErrSection:
    RaiseError "cChart.ClearAnnotMultiFlag"

End Sub

Public Property Get TradeAccountID() As Long

    If g.nReplaySession > 0 Or frmReplay.Visible Then
        TradeAccountID = g.nReplayAccountID
    Else
        TradeAccountID = m.nAccountID
    End If

End Property

Public Property Let TradeAccountID(ByVal nID&)
    
    'when user [clicks set date/time] button, replay session gets reset to zero, but
    'user is still in streaming replay therefore must check form's visibility as well
    If g.nReplaySession = 0 And Not frmReplay.Visible Then
        RemoveAnnots True, , eANNOT_AvgEntryLine        '4220
        Set m.AvgEntryAnnot = Nothing
        m.nAccountID = nID
    End If

End Property

Public Sub UpdateOnlineOrder(Order As cPtOrder, ByVal bAll As Boolean)
On Error GoTo ErrSection:

    Dim strKey$, i&, j&
    
    Dim OrderStruct As cOrderStruct
    Dim OpenOrders As cGdTree
    Dim CurrOrder As cPtOrder
    
    Dim aAutoExits As cGdArray          'holds cOrderStruct object of auto-exit orders
    Dim aAutoExitPrices As cGdArray     'holds prices of auto-exit orders
    Dim dCurrPrice#
    
    Dim bClearAll As Boolean

    If m.nShowTrades <> 2 Then
        bClearAll = True
    ElseIf bAll Then
        If m.Form.fraOrderBtns.Visible Then Set OpenOrders = g.Broker.PrimaryOrdersForSymbol(TradeAccountID, m.nSymbolID, 0)
    Else
        If Not Order Is Nothing And Not m.OnlineOrders Is Nothing Then
            'called from mouse event in chart's form to update this order only
            'if order tree does not already exist, something is very wrong (don't create tree)
            strKey = Order.OrderID
            Set OrderStruct = m.OnlineOrders(strKey)
            If Not OrderStruct Is Nothing Then
                If Order.Status = eTT_OrderStatus_Cancelled Then
                    'do nothing - drop from tree till next update (a little faster feedback for issue 5841)
                ElseIf OrderStruct.OrderToStruct(Me, Order) = 0 Then
                    If Order.OrderType <> eTT_OrderType_Market And _
                       Order.OrderType <> eTT_OrderType_MarketOnClose Then
                        OrderStruct.geDrawOrder Me
                    End If
                Else
                    'should show some kind of message here
                End If
            End If
        End If
        Exit Sub
    End If
            
    'remove all objects from tree and exit if there are no open orders
    If OpenOrders Is Nothing Then
        bClearAll = True
    ElseIf OpenOrders.Count = 0 Then
        bClearAll = True
    End If
    If bClearAll Then
        m.Form.cmdCancelAll.Enabled = False
        If Not m.OnlineOrders Is Nothing Then m.OnlineOrders.Clear
        If m.Form.BracketOrderOne Is Nothing Then Exit Sub
    End If
    m.Form.cmdCancelAll.Enabled = True
    
    If m.OnlineOrders Is Nothing Then
        Set m.OnlineOrders = New cGdTree        'okay to create tree here
    End If
    
    'remove objects from tree that are no longer open orders
    For i = m.OnlineOrders.Count To 1 Step -1
        Set OrderStruct = m.OnlineOrders(i)
        If Not OrderStruct Is Nothing Then
            Set CurrOrder = OpenOrders(Str(OrderStruct.ItemID))
            If CurrOrder Is Nothing Then
                m.OnlineOrders.Remove i
                Set OrderStruct = Nothing
            End If
        End If
    Next
    
    'add/update order objects
    Set aAutoExits = New cGdArray
    Set aAutoExitPrices = New cGdArray
    
    For i = 1 To OpenOrders.Count
        Set CurrOrder = OpenOrders(i)
        If Not CurrOrder Is Nothing Then
            If m.Form.OrderMoveInProg <> CurrOrder.OrderID Then     'aardvark 4071
                strKey = Str(CurrOrder.OrderID)
                Set OrderStruct = m.OnlineOrders(strKey)
                If OrderStruct Is Nothing Then
                    Set OrderStruct = New cOrderStruct
                    m.OnlineOrders.Add OrderStruct, strKey
                End If
                OrderStruct.OrderToStruct Me, CurrOrder
                UpdateTrigger CurrOrder
                If CurrOrder.IsAutoExit Then
                    aAutoExits.Add OrderStruct
                    aAutoExitPrices.Add CurrOrder.OrderPrice(False)
                End If
            End If
        End If
    Next
    
    'check if there are auto-exits on opposite side of the market
    aAutoExitPrices.Sort
    dCurrPrice = m.Bars(eBARS_Close, LastGoodDataBar(False, False))
    dCurrPrice = RoundToMinMove(dCurrPrice, m.Bars.MinMove(m.Bars(eBARS_DateTime, m.Bars.Size - 1)))
    
    If dCurrPrice >= aAutoExitPrices(0) Then
        If dCurrPrice <= aAutoExitPrices(aAutoExitPrices.Size - 1) Then         '5723
            For i = 0 To aAutoExits.Size - 1
                Set OrderStruct = aAutoExits(i)
                OrderStruct.ConnectAutoExit = 1
            Next
        End If
    End If
    Set aAutoExits = Nothing
    Set aAutoExitPrices = Nothing
    
    'add bracket order(s) if exist
    Set CurrOrder = m.Form.BracketOrderOne
    If Not CurrOrder Is Nothing Then
        If CurrOrder.Status = eTT_OrderStatus_Parked Then
            strKey = Str(CurrOrder.OrderID)
            Set OrderStruct = New cOrderStruct
            m.OnlineOrders.Add OrderStruct, strKey
            OrderStruct.OrderToStruct Me, CurrOrder, , -1
        End If
    End If
    
    Set CurrOrder = m.Form.BracketOrderTwo
    If Not CurrOrder Is Nothing Then
        If CurrOrder.Status = eTT_OrderStatus_Parked Then
            strKey = Str(CurrOrder.OrderID)
            Set OrderStruct = New cOrderStruct
            m.OnlineOrders.Add OrderStruct, strKey
            OrderStruct.OrderToStruct Me, CurrOrder, , -1
        End If
    End If

    
    'set pane scale to fit order
    Dim Pane As cPane
    Dim dMin#, dMax#, dPrice#
    
    dMax = kNullData
    dMin = -1 * kNullData
    For i = 1 To m.OnlineOrders.Count
        Set OrderStruct = m.OnlineOrders(i)
        If Not OrderStruct Is Nothing Then
            dPrice = OrderStruct.ItemPrice
            If dPrice <= 0 Then
                ' negative order prices only allowed for SpreadSymbols
                If Not IsSpreadSymbol(Me.Symbol) Then
                    dPrice = kNullData
                End If
            End If
            If dPrice <> kNullData Then
                If dPrice > dMax Then dMax = dPrice
                If dPrice < dMin Then dMin = dPrice
            End If
        End If
    Next
    
    Set Pane = m.Tree("PRICE PANE")
    If Not Pane Is Nothing Then
        With Pane
            If .Scaling = ePANE_ScaleModeAuto Then
                If .gePaneMax > 0 And .gePaneMax < dMax Then
                    .gePaneMax = dMax
                    .Max = .gePaneMax
                    .geSavePaneMax
                End If
                If .gePaneMin > 0 And .gePaneMin > dMin Then
                    .gePaneMin = dMin
                    .Min = .gePaneMin
                    .geSavePaneMin
                End If
            End If
        End With
    End If

    Exit Sub

ErrSection:
    RaiseError "cChart.UpdateOnlineOrder"

End Sub

Private Sub UpdateTrigger(Order As cPtOrder)
On Error GoTo ErrSection:

    Dim i&, strKey$
    Dim TrigTree As cGdTree
    Dim ChildOrder As cPtOrder
    Dim ChildStruct As cOrderStruct
    
    Set TrigTree = g.Broker.TriggeredOrdersForOrder(Order)
    
    If Not TrigTree Is Nothing Then
        For i = 1 To TrigTree.Count
            Set ChildOrder = TrigTree(i)
            If ChildOrder.SymbolID = Order.SymbolID Then
                If Not ChildOrder Is Nothing Then
                    strKey = Str(ChildOrder.OrderID)
                    Set ChildStruct = m.OnlineOrders(strKey)
                    If ChildStruct Is Nothing Then
                        Set ChildStruct = New cOrderStruct
                        m.OnlineOrders.Add ChildStruct, strKey
                    End If
                    ChildStruct.OrderToStruct Me, ChildOrder
                    UpdateTrigger ChildOrder
                End If
            End If
        Next
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cChart.UpdateTrigger"

End Sub

Private Function TradesFileForAccount(ByVal bGenerateFile As Boolean) As String
On Error GoTo ErrSection:

    Dim strFile As String               ' Name of the file to dump the trades to
    
    strFile = App.Path & "\Trades\S_" & Trim(Str(m.Form.hWnd)) & ".TXT"
    
    If bGenerateFile Then
        KillFile strFile, True
        If Not g.Broker.ToTradesFile(strFile, TradeAccountID, SymbolID, 0&, m.Bars.Prop(eBARS_PeriodicityStr)) Then
            strFile = ""
        End If
    End If
    
    TradesFileForAccount = strFile
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cChart.TradesFileForAccount"

End Function

Public Property Get OnlineOrders() As cGdTree
    Set OnlineOrders = m.OnlineOrders
End Property

Public Sub UpdateTradePrices()
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    
    Dim i&, strText$, dPrice#
    Dim strPos$, strPosQty$, strOpenEq$, strAvgEntry$
    Dim strSessionPL$, strSessionQty$
    Dim bInPosition As Boolean
    Dim bMismatch As Boolean            ' Is there currently a mismatch for the symbol?
    Dim strPrevCaption As String        ' Previous text of the string
    Dim nBroker As eTT_AccountType      ' Broker for the given account ID
                        
    If bInProgress Then Exit Sub
    If g.bUnloading Then Exit Sub
    If Not m.Form.vseOrderBar.Visible Then Exit Sub
    
    bInProgress = True
        
    'precautionary, theoretically should never happen
    If m.OrderDailyBars Is Nothing Then Set m.OrderDailyBars = New cGdBars
    
    If m.OrderDailyBars.Prop(eBARS_SymbolID) <> m.nSymbolID Or m.OrderDailyBars.Size = 0 Then
        Set m.OrderDailyBars = New cGdBars
        m.OrderDailyBars.Prop(eBARS_SymbolID) = m.nSymbolID
        m.OrderDailyBars.ArrayMask = eBARS_EodBidAsk
        i = DM_GetBars(m.OrderDailyBars, m.nSymbolID, "Daily", LastDailyDownload - 5, , , False, , False)
    End If
    
    With m.Form
        'bid/ask prices
        g.RealTime.UpdateBidAsk m.OrderDailyBars
        dPrice = m.OrderDailyBars(eBARS_Ask, m.OrderDailyBars.Size - 1)
        strText = m.OrderDailyBars.PriceDisplay(dPrice)
        If .lblAsk.Caption <> strText Then .lblAsk.Caption = strText
        
        dPrice = m.OrderDailyBars(eBARS_Bid, m.OrderDailyBars.Size - 1)
        strText = m.OrderDailyBars.PriceDisplay(dPrice)
        If .lblBid.Caption <> strText Then .lblBid.Caption = strText
        
        'current price
        dPrice = m.Bars(eBARS_Close, LastGoodDataBar(False))
        strText = m.Bars.PriceDisplay(dPrice)
        If .lblMarket.Caption <> strText Then .lblMarket.Caption = strText
        
        'current position, open equity on order bar
        strText = g.Broker.PositionString(TradeAccountID, TradeSymbolID(True), 0&)
        
        nBroker = g.Broker.AccountTypeForID(TradeAccountID)
        strPrevCaption = .lblTradePos.Text
        
        If Len(strText) > 0 Then
            strPos = Parse(strText, "|", 1)
            strPosQty = Parse(strText, "|", 2)
            strOpenEq = Parse(strText, "|", 3)
            strAvgEntry = Parse(strText, "|", 4)
            strSessionQty = Parse(strText, "|", 5)
            strSessionPL = Parse(strText, "|", 6)
            ' 06/29/2009 DAJ: Set this from the actual decimal value passed over
            ' instead of the trading units in case the chart's display units are
            ' different...
            ''m.dAvgEntry = m.Bars.PriceFromString(strAvgEntry)       '4835
            m.dAvgEntry = Val(Parse(strText, "|", 7))
                        
            bMismatch = False
            If InStr(UCase(strPos), "LONG") > 0 Then
                .lblTradePos.ForeColor = ChartGlbClrForCtl(.lblTradePos, g.ChartGlobals.nLongColor, "LongColor")
                .lblTradePos.Text = strPos & " " & strPosQty
                bInPosition = True
            ElseIf InStr(UCase(strPos), "SHORT") > 0 Then
                .lblTradePos.ForeColor = g.ChartGlobals.nShortColor
                .lblTradePos.Text = strPos & " " & strPosQty
                bInPosition = True
            ElseIf UCase(strPos) = "MISMATCH" Then
                .lblTradePos.Text = "Mismatch"
                .lblTradePos.ForeColor = 0
                bInPosition = True
                bMismatch = True
            Else
                .lblTradePos.Text = "Flat"
                .lblTradePos.ForeColor = 0
                bInPosition = False
            End If
                        
            If bMismatch = True Then
                .lblEquity.Caption = ""
            Else
                .lblEquity.Caption = strOpenEq
                If InStr(strOpenEq, "-") Or InStr(strOpenEq, "(") Then
                    .lblEquity.ForeColor = g.ChartGlobals.nLossColor
                Else
                    .lblEquity.ForeColor = ChartGlbClrForCtl(.lblEquity, g.ChartGlobals.nWinColor, "WinColor")
                End If
            End If
        Else
            .lblTradePos.Text = "Flat"
            .lblTradePos.ForeColor = 0
            .lblEquity.Caption = ""
            bInPosition = False
            m.dAvgEntry = 0#
        End If
        
        ' DAJ 09/26/2013: If the position label changes, dump it to the log...
        If (.lblTradePos.Text <> strPrevCaption) And (nBroker > 0) Then
            g.Broker.BrokerDebug nBroker, vbTab & vbTab & "Chart ( " & Str(.hWnd) & ", '" & Symbol & "', '" & g.Broker.AccountNameForID(TradeAccountID) & "' ): Position changed to '" & .lblTradePos.Text & "' ( " & strText & " )"
        End If
        
        'check exit favorite buttons in case position changed
        ExitFavoritesCheck m.Form, .lblTradePos.Text
        
        TSOGrpFavoritesCheck m.Form
        
        'update account bar
        If m.bAccountBar Then
            UpdateAccountBar .fgChartFlex(eFlexGridIdx_AcctBar), strPos, strPosQty, strOpenEq, strAvgEntry, _
                strSessionPL, strSessionQty, TradeAccountID, SecurityType(m.Bars)
        End If
        
        If (.CancelAllInProgress = True) And (.cmdCancelAll.Enabled = True) Then
            .cmdCancelAll.Enabled = False
        End If
        
        If (.FlattenInProgress = True) Or (.ReverseInProgress = True) Then
            If .cmdCancelAll.Enabled = True Then
                .cmdCancelAll.Enabled = False
            End If
            If .cmdReverse.Enabled = True Then
                .cmdReverse.Enabled = False
            End If
            If .cmdBailout.Enabled = True Then
                .cmdBailout.Enabled = False
                .cmdBailout.BackColor = .cmdReverse.BackColor
            End If
        ElseIf bInPosition Then
'JM 09-01-2010: original code commented out to implement aardvark 5900 (leave awhile then remove if all ok)
'            If .cmdReverse.Enabled = False Then
'                .cmdReverse.Enabled = True
'            End If
            If .cmdBailout.Enabled = False Then
                .cmdBailout.Enabled = True
                .cmdBailout.BackColor = RGB(192, 0, 0) ' &HFFFF&
            End If
        Else
            If .cmdReverse.Enabled = True Then
                .cmdReverse.Enabled = False
            End If
            If .cmdBailout.Enabled = True Then
                .cmdBailout.Enabled = False
                .cmdBailout.BackColor = .cmdReverse.BackColor
            End If
        End If
        'check disconnect
        If g.Broker.AccountTypeForID(TradeAccountID) = eTT_AccountType_SimTrade Then
            If m.Form.OrderBarMode = eOrdBarMode_BrokerDisconnect Then m.Form.OrderBarMode = eOrdBarMode_Order
        Else
            If m.Form.OrderBarMode = eOrdBarMode_BrokerDisconnect Then
                If g.Broker.ConnectionStatusForAccount(TradeAccountID) = eGDConnectionStatus_Connected Then m.Form.OrderbarWrapper eOrdBarMode_Order
            ElseIf m.Form.OrderBarMode = eOrdBarMode_Order Then
                If g.Broker.ConnectionStatusForAccount(TradeAccountID) <> eGDConnectionStatus_Connected Then m.Form.OrderbarWrapper eOrdBarMode_BrokerDisconnect
            End If
        End If
    End With
                       
    bInProgress = False
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.UpdateTradePrices"

End Sub

Public Sub SetTrackerTradesReload()
    m.bReloadTrackerTrades = True
End Sub

Public Property Get PseudoOrderType() As Long
    PseudoOrderType = m.nPseudoOrderType
End Property

Public Property Let PseudoOrderType(ByVal nType As Long)
    m.nPseudoOrderType = nType
End Property

Public Property Get ShowAccountBar() As Boolean
    
    If m.nShowTrades = 2 Then
        ShowAccountBar = m.bAccountBar
    End If
    
End Property

Public Property Let ShowAccountBar(ByVal bShow As Boolean)
    m.bAccountBar = bShow
End Property

Public Property Get SpreadAsDollar() As Boolean
    SpreadAsDollar = m.bSpreadAsDollar
End Property

Public Sub UpdateWoodCountDown()
On Error GoTo ErrSection:

    Static prevJ&
    Dim j&, rc&
    Dim Pane As cPane

    If m.nShowSplitPane = 0 Or Len(g.strActiveDraw) > 0 Then Exit Sub       'aardvark 4259
    
    If Not ActiveChart Is Nothing Then
        If ActiveChart.WindowState = vbMaximized And (m.Form.hWnd <> ActiveChart.hWnd) Then
            Exit Sub
        End If
    End If
    
    If m.Bars.IsIntraday Then
        If m.dLastScreenDate < m.dLastActualDateOfData Then
            'last data bar not visible on chart, countdown = 0
        ElseIf m.Bars.Prop(eBARS_PeriodType) = ePRD_Minutes Then
            If g.RealTime.Active And g.RealTime.FeedTime > 0 Then
                Set Pane = m.Tree("PRICE PANE")
                If Not Pane Is Nothing Then
                    If Not Pane.WoodPane Is Nothing Then
                        If gdTickCount(False) - m.dLastCountdownCheck > 950 Or prevJ < 0 Then
                            j = Pane.WoodPane.CountDownToNextBar(Me, LastGoodDataBar(False))
                            If j >= 0 And j <> prevJ Then
                                prevJ = j
                                If Len(Pane.WoodPane.TimerBkColor) = 0 Then
                                    geRepaintTimer m.geChartObj, m.Form.pbChart.hWnd, m.Form.pbChart.hDC, "-1", Str(j) & "s"
                                Else
                                    geRepaintTimer m.geChartObj, m.Form.pbChart.hWnd, m.Form.pbChart.hDC, Pane.WoodPane.TimerBkColor, Str(j) & "s"
                                End If
                            End If
                            m.dLastCountdownCheck = gdTickCount(False)
                        Else
                            'chart was generated less than a second ago, just set prevJ to what is shown on chart
                            prevJ = Pane.WoodPane.LastTimerVal
                        End If
                    End If
                End If
            Else
                prevJ = -1      'reset
            End If
        Else
            'don't need to do anything since countdown will not change until
            'a new tick comes in which triggers a call to generatechart anyways
        End If
    ElseIf Len(m.strSpreadSymbols) > 0 Then
        GenerateChart eRedo1_Scrolled
    End If
            
    Exit Sub
    
ErrSection:
    RaiseError "cChart.UpdateWoodCountDown"
    
End Sub

Public Property Get OrdBarCtrls() As String
On Error GoTo ErrSection:

    OrdBarCtrls = m.strOrdBarCtrls

    Exit Property
    
ErrSection:
    RaiseError "cChart.OrdBarCtrls.Get"
    
End Property

Public Property Let OrdBarCtrls(ByVal strSettings$)
On Error GoTo ErrSection:
    
    m.strOrdBarCtrls = strSettings
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.OrdBarCtrls.Let"
    
End Property

Public Property Get ShowSplitPane() As Long
On Error GoTo ErrSection:

    ShowSplitPane = m.nShowSplitPane
    
    Exit Property

ErrSection:
    RaiseError "cChart.ShowSplitPane.Get"
    
End Property

Public Property Let ShowSplitPane(ByVal nShow&)
On Error GoTo ErrSection:
    
    If m.nShowSplitPane <> nShow Then
        If nShow <> 0 Then m.bRedrawTwice = True
        m.nShowSplitPane = nShow
        ResetSplitPane
    End If
    
    Exit Property

ErrSection:
    RaiseError "cChart.ShowSplitPane.Let"
    
End Property

Public Property Get LimitTrades() As String
On Error GoTo ErrSection:

    LimitTrades = m.strLimitTrades
    
    Exit Property

ErrSection:
    RaiseError "cChart.LimitTrades.Get"
    
End Property

Public Property Let LimitTrades(ByVal strLimitTrades$)
On Error GoTo ErrSection:

    m.strLimitTrades = strLimitTrades
    
    Exit Property
    
ErrSection:
    RaiseError "cChart.LimitTrades.Let"
    
End Property

Private Function LimitTradeLoop(ByVal iLastRec&, ByVal nLastBar&) As Long
On Error GoTo ErrSection:

    Dim dLimitDate#, dLimitDatePrev#
    Dim iLimit&, i&, j&

    If Len(m.strLimitTrades) = 0 Or m.Form.IsInGameMode Then
        LimitTradeLoop = 0
        Exit Function
    End If
    
    iLimit = Int(ValOfText(Parse(m.strLimitTrades, "|", 1)))
    If iLimit < 0 Then iLimit = 0
        
    If InStr(m.strLimitTrades, "T") <> 0 Then
        iLimit = iLastRec - iLimit
    ElseIf InStr(m.strLimitTrades, "B") <> 0 Then
        dLimitDate = m.aXdate(nLastBar - iLimit)
        If iLimit > 0 Or m.Bars.IsIntraday Then
            dLimitDatePrev = m.aXdate(nLastBar - (iLimit + 1))
        Else
            dLimitDatePrev = dLimitDate
        End If
        m.Transactions.FieldArray(eGDTransactionField_Date, False).BinarySearch dLimitDate + 0.000001, iLimit
        j = iLimit - 1
        For i = j To 0 Step -1
            If m.Transactions.TblNum(eGDTransactionField_Date, i) > dLimitDatePrev And m.Transactions.TblNum(eGDTransactionField_Date, i) <= dLimitDate Then
                iLimit = i
            Else
                Exit For
            End If
        Next
'        StatusMsg DateFormat(dLimitDate, NO_DATE, HH_MM_SS) & "," & DateFormat(dLimitDatePrev, NO_DATE, HH_MM_SS) & "," & DateFormat(m.Transactions.TblNum(eGDTransactionField_Date, i), NO_DATE, HH_MM_SS)
    ElseIf InStr(m.strLimitTrades, "D") <> 0 Then
        dLimitDate = Int(m.aXdate(nLastBar)) - iLimit
        dLimitDatePrev = Int(dLimitDate) - 1
        m.Transactions.FieldArray(eGDTransactionField_Date, False).BinarySearch dLimitDate + 0.000001, iLimit
        j = iLimit - 1
        For i = j To 0 Step -1
            If Int(m.Transactions.TblNum(eGDTransactionField_Date, i)) > dLimitDatePrev And Int(m.Transactions.TblNum(eGDTransactionField_Date, i)) <= dLimitDate Then
                iLimit = i
            Else
                Exit For
            End If
        Next
    ElseIf InStr(m.strLimitTrades, "M") <> 0 Then
        dLimitDate = m.aXdate(nLastBar) - (iLimit / 1440#)
        If Not m.Bars.IsIntraday Then
            dLimitDate = dLimitDate + m.Bars.Prop(eBARS_LastTickTime) / 1440#
        End If
        
        If iLimit > 0 Then
            dLimitDatePrev = dLimitDate - (iLimit / 1440#)
        Else
            dLimitDatePrev = dLimitDate - (1 / 1440 / 60 * 58)
        End If
        
        m.Transactions.FieldArray(eGDTransactionField_Date, False).BinarySearch dLimitDate + 0.000001, iLimit
        j = iLimit - 1
        For i = j To 0 Step -1
            If m.Transactions.TblNum(eGDTransactionField_Date, i) > dLimitDatePrev And m.Transactions.TblNum(eGDTransactionField_Date, i) <= dLimitDate Then
                iLimit = i
            Else
                Exit For
            End If
        Next
'        StatusMsg DateFormat(dLimitDate, NO_DATE, HH_MM_SS) & "," & DateFormat(dLimitDatePrev, NO_DATE, HH_MM_SS) & "," & DateFormat(m.Transactions.TblNum(eGDTransactionField_Date, i), NO_DATE, HH_MM_SS)
    End If
    
    LimitTradeLoop = iLimit
    
    Exit Function

ErrSection:
    RaiseError "cChart.LimitTradeLoop"
    
End Function

Public Property Get ChartBkIsLight() As Boolean
On Error Resume Next
    
    Dim nColor&
    
    If m.bCustomColors Then
        nColor = m.nChartBackColor
    Else
        nColor = g.ChartGlobals.nChartBackColor
    End If

    ChartBkIsLight = IsLightColor(nColor)

End Property

Public Sub HandleDinapButtons(ByVal strBtnID$)
On Error GoTo ErrSection:

    Dim i&, iPricePaneIdx&, strCodedName$
    Dim Ind As cIndicator
    Dim Ind2 As cIndicator
    
    Dim IndDinapMA As cIndicator
    Dim IndMacdPredictor As cIndicator
    Dim IndOscPredLower As cIndicator
    Dim IndOscPredUpper As cIndicator
    
    Dim Pane As cPane
    Dim PaneMacD As cPane
    Dim PaneDetrendOsc As cPane
    Dim PanePrefStoch As cPane
    
    Dim aIndicators As cGdArray
    
    If m.Tree Is Nothing Then Exit Sub
    
    iPricePaneIdx = m.Tree.RelativeIndex("PRICE", eTREE_Parent)
    If iPricePaneIdx <= 0 Then Exit Sub
    
    Select Case strBtnID
        'indicators
        Case "ID_DisplacedMA", "ID_OscPredictor", "ID_MacdPredictor"
            'don't need to do anything, just continue
        'panes
        Case "ID_DiNapoliMACD", "ID_DetrendOsc", "ID_PrefStoch"
            Set aIndicators = New cGdArray
            aIndicators.Create eGDARRAY_Objects
        Case Else
            Exit Sub
    End Select
    
    For i = 1 To m.Tree.Count
        If Tree.NodeLevel(i) = 0 Then
            Set Pane = Tree(i)
            If Not aIndicators Is Nothing Then
                Pane.CheckDinapoliStudies PaneMacD, PaneDetrendOsc, PanePrefStoch, aIndicators
                aIndicators.Size = 0
                If strBtnID = "ID_DiNapoliMACD" And Not PaneMacD Is Nothing Then Exit For
                If strBtnID = "ID_DetrendOsc" And Not PaneDetrendOsc Is Nothing Then Exit For
                If strBtnID = "ID_PrefStoch" And Not PanePrefStoch Is Nothing Then Exit For
            End If
        ElseIf m.Tree.AncestorIndex(i, 0) = iPricePaneIdx Then
            If aIndicators Is Nothing Then
                Set Ind = Tree(i)
                If Not Ind Is Nothing Then
                    strCodedName = UCase(Ind.CodedName)
                    Select Case strBtnID
                        Case "ID_DisplacedMA"
                            If Ind.IsDinapoliDMA Then
                                Set IndDinapMA = Ind
                                Exit For
                            End If
                        Case "ID_OscPredictor"
                            If InStr(strCodedName, "UPPEROSCPRED") > 0 Then
                                Set IndOscPredUpper = Ind
                                If Not IndOscPredLower Is Nothing Then Exit For
                            ElseIf InStr(strCodedName, "LOWEROSCPRED") > 0 Then
                                Set IndOscPredLower = Ind
                                If Not IndOscPredUpper Is Nothing Then Exit For
                            End If
                        Case "ID_MacdPredictor"
                            If InStr(strCodedName, "DINAPOLIMACDPREDICTOR") > 0 Then
                                Set IndMacdPredictor = Ind
                                Exit For
                            End If
                    End Select
                End If
            End If
        ElseIf Not aIndicators Is Nothing Then
            aIndicators.Add Tree(i)
        End If
    Next
    
    'check the last pane if necessary
    If Not aIndicators Is Nothing Then
        If strBtnID = "ID_DiNapoliMACD" And PaneMacD Is Nothing Or _
           strBtnID = "ID_DetrendOsc" And PaneDetrendOsc Is Nothing Or _
           strBtnID = "ID_PrefStoch" And PanePrefStoch Is Nothing Then
        
            Pane.CheckDinapoliStudies PaneMacD, PaneDetrendOsc, PanePrefStoch, aIndicators
        End If
    End If
    
'toggle or add indicators/panes as needed
    Select Case strBtnID
        'indicators
        Case "ID_DisplacedMA"
            If IndDinapMA Is Nothing Then
                If TemplateAddStudy("CTP - Displaced MA") Then
                    'need to turn off the 7 & 25 bar MA that is part of this study
                    For i = 1 To m.Tree.Count
                        If m.Tree.NodeLevel(i) > 0 Then
                            Set Ind = m.Tree(i)
                            If Not Ind Is Nothing Then
                                If Ind.IsDinapoliDMA Then
                                    Set Ind = m.Tree(i + 1)
                                    Set Ind2 = m.Tree(i + 2)
                                    If Not Ind Is Nothing Then Ind.Display = False
                                    If Not Ind2 Is Nothing Then Ind2.Display = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                IndDinapMA.Display = Not IndDinapMA.Display
            End If
        
        Case "ID_OscPredictor"
            If IndOscPredLower Is Nothing Or IndOscPredUpper Is Nothing Then
               TemplateAddStudy "CTP - Osc Predictor"
            ElseIf IndOscPredUpper.Display = IndOscPredLower.Display Then
                IndOscPredUpper.Display = Not IndOscPredUpper.Display
                IndOscPredLower.Display = Not IndOscPredLower.Display
            Else
                IndOscPredUpper.Display = True      'if one is on then turn both on
                IndOscPredLower.Display = True
            End If
            
        Case "ID_MacdPredictor"
            If IndMacdPredictor Is Nothing Then
                TemplateAddStudy "DiNapoli MACD Predictor"
            Else
                IndMacdPredictor.Display = Not IndMacdPredictor.Display
            End If
    
        'panes
        Case "ID_DiNapoliMACD"
            If PaneMacD Is Nothing Then
                TemplateAddStudy "CTP - MACD"
            Else
                PaneMacD.Display = Not PaneMacD.Display
            End If
        
        Case "ID_DetrendOsc"
            If PaneDetrendOsc Is Nothing Then
                TemplateAddStudy "CTP - Detrend Oscillator"
            Else
                PaneDetrendOsc.Display = Not PaneDetrendOsc.Display
            End If
        
        Case "ID_PrefStoch"
            If PanePrefStoch Is Nothing Then
                TemplateAddStudy "CTP - Pref Stochastic"
            Else
                PanePrefStoch.Display = Not PanePrefStoch.Display
            End If
    
    End Select
    
ErrExit:
    Set aIndicators = Nothing
    
    GenerateChart eRedo3_Settings
    
    Exit Sub

ErrSection:
    RaiseError "cChart.HandleDinapButtons"
    
End Sub


Private Sub LogCCI(ByVal strDesc$)

    Dim strFile$, bAppend As Boolean
    Static iCount&
    
    strFile = App.Path & "\chk\cci.txt"
    If iCount = 0 Then
        If Not FileExist(strFile) Then iCount = -1
    End If
    If iCount >= 0 Then
        If iCount > 0 Then bAppend = True
        iCount = iCount + 1
        FileFromString strFile, strDesc & "  " & Str(iCount), True, bAppend
    End If

End Sub

Public Property Get ShowOrderHorzLine() As Long
On Error GoTo ErrSection:

    ShowOrderHorzLine = m.nShowOrderHorzLine

    Exit Property

ErrSection:
    RaiseError "cChart.ShowOrderHorzLine.Get"
    
End Property

Public Property Let ShowOrderHorzLine(ByVal nShow&)
On Error GoTo ErrSection:

    m.nShowOrderHorzLine = nShow

    Exit Property

ErrSection:
    RaiseError "cChart.ShowOrderHorzLine.Let"

End Property

Public Property Get ShowAvgEntryLine() As Long
On Error GoTo ErrSection:

    ShowAvgEntryLine = m.nShowAvgEntryLine

    Exit Property

ErrSection:
    RaiseError "cChart.ShowAvgEntryLint.Get"
    
End Property

Public Property Let ShowAvgEntryLine(ByVal nShow&)
On Error GoTo ErrSection:

    m.nShowAvgEntryLine = nShow

    Exit Property

ErrSection:
    RaiseError "cChart.ShowAvgEntryLine.Let"
    
End Property

Public Property Get AvgEntryProp(ByVal nProp&) As String
On Error GoTo ErrSection:

    '1=color, 2=pen style, 3=textPosition, 4=font, 5=font size, 6=font style, 7=font underline
        
    If nProp > 0 And nProp < 8 Then
        AvgEntryProp = Parse(m.strAvgEntryProp, "|", nProp)
    End If
    
    Exit Property

ErrSection:
    RaiseError "cChart.AvgEntryProp.Get"
    
End Property

Public Property Let AvgEntryProp(ByVal nProp&, ByVal strValue$)
On Error GoTo ErrSection:

    Dim aProp As cGdArray, i&
    
    '1=color, 2=pen style, 3=textPosition, 4=font, 5=font size, 6=font style, 7=font underline
    
    If nProp > 0 And i < 8 And Len(strValue) > 0 Then
        'save new value to property string regardless of whether average entry line is visible
        Set aProp = New cGdArray
        aProp.SplitFields m.strAvgEntryProp, "|"
        aProp(nProp - 1) = strValue         'array index is zero-based
        m.strAvgEntryProp = aProp.JoinFields("|")
        
        'update annot's property if annot exists
        If Not m.AvgEntryAnnot Is Nothing Then
            Select Case nProp
                Case 1
                    m.AvgEntryAnnot.Color = ValOfText(strValue)         'color
                Case 2
                    m.AvgEntryAnnot.Style = ValOfText(strValue)         'pen style
                Case 3
                    m.AvgEntryAnnot.geTextAlign = ValOfText(strValue)   'text position relative to line
                Case 4
                    m.AvgEntryAnnot.Prop("Font") = strValue
                Case 5
                    m.AvgEntryAnnot.Prop("FontSize") = strValue
                Case 6
                    m.AvgEntryAnnot.Prop("FontStyle") = strValue
                Case 7
                    m.AvgEntryAnnot.Prop("FontUnderline") = strValue
            End Select
        End If
    End If

    Exit Property

ErrSection:
    RaiseError "cChart.AvgEntryProp.Let"
    
End Property

Public Sub OrdBarSaveAsDefaults()
On Error Resume Next

    SetIniFileProperty "OrdBarCtrls", m.strOrdBarCtrls, "OrdBarDefaults", g.strIniFile
    SetIniFileProperty "ShowOrderHorzLine", m.nShowOrderHorzLine, "OrdBarDefaults", g.strIniFile
    SetIniFileProperty "ShowAvgEntryLine", m.nShowAvgEntryLine, "OrdBarDefaults", g.strIniFile
    SetIniFileProperty "AvgEntryProp", m.strAvgEntryProp, "OrdBarDefaults", g.strIniFile
    SetIniFileProperty "AcctBarCols", m.strAcctBarCols, "OrdBarDefaults", g.strIniFile
    SetIniFileProperty "HighlightPos", m.nHighlightPos, "OrdBarDefaults", g.strIniFile
    SetIniFileProperty "HighlightEquity", m.nHighlightEquity, "OrdBarDefaults", g.strIniFile
    
End Sub

Public Sub OrdBarGetDefaults()
On Error GoTo ErrSection:

    Dim i&, strControls$, strBR$

    'JM: Pete's email 10-21-2009:
    '    Glen wants the Bracket order button to be on the Order bar by default
    '    one-time override to turn on bracket order button
    
    strControls = GetIniFileProperty("OrdBarCtrls", kOrdBarDefaults, "OrdBarDefaults", g.strIniFile)
    i = GetIniFileProperty("BrOrdOverRide", 0, "OrdBarDefaults", g.strIniFile)
    
    If i = 0 Then
        i = InStr(strControls, "BRORD")
        If i <> 0 Then
            strBR = Mid(strControls, i, 7)
            strControls = Replace(strControls, strBR, "BRORD;1")
            SetIniFileProperty "OrdBarCtrls", strControls, "OrdBarDefaults", g.strIniFile
        End If
        
        SetIniFileProperty "BrOrdOverRide", 1, "OrdBarDefaults", g.strIniFile
    End If
        
    'default order bar controls to show all except buy/sell bid/ask buttons
    m.strOrdBarCtrls = ConvertOrdBarButtons(strControls)
    If Len(m.strOrdBarCtrls) = 0 Then m.strOrdBarCtrls = kOrdBarDefaults    'something went wrong just use new defaults
    
    '1=color, 2=pen style, 3=textPosition, 4=font, 5=font size, 6=font style, 7=font underline
    'default: dark green|thin|ctrLeft|arial|8|bold|0
    m.strAvgEntryProp = GetIniFileProperty("AvgEntryProp", "32768|1|9|arial|8|1|0", "OrdBarDefaults", g.strIniFile)
    
    m.nShowOrderHorzLine = GetIniFileProperty("ShowOrderHorzLine", 0, "OrdBarDefaults", g.strIniFile)
    m.nShowAvgEntryLine = GetIniFileProperty("ShowAvgEntryLine", 0, "OrdBarDefaults", g.strIniFile)
    m.nHighlightPos = GetIniFileProperty("HighlightPos", -2, "OrdBarDefaults", g.strIniFile)
    m.nHighlightEquity = GetIniFileProperty("HighlightEquity", -2, "OrdBarDefaults", g.strIniFile)
    m.strAcctBarCols = GetIniFileProperty("AcctBarCols", kABarCols, "OrdBarDefaults", g.strIniFile)
    
    Set m.AvgEntryAnnot = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "cChart.OrdBarGetDefaults"

End Sub

Public Property Get SpreadBars() As Variant
On Error Resume Next
    
    SpreadBars = m.SpreadBars

End Property

Public Sub ActivateWhatIf(Optional ByVal strWhatIfBar As String = "")
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    
    Dim dDate1#, dDate2#, dClose#
    Dim iLastBar&, iBlankBarsSave&, strMsg$
      
    If InvalidWhatIf(True) Then
        DeactivateWhatIf False
        SyncToolbar
        Exit Sub
    End If
        
    strWhatIfBar = UCase(Trim(strWhatIfBar))
    If strWhatIfBar = "L" Or strWhatIfBar = "N" Then
        'passed in value is okay, don't need to do anything
    Else
        strMsg = "'What-If' allows you to move the current price up/down to see how indicators are affected.||"
        strMsg = strMsg & "Would you like to adjust the close of the last bar or the open of a new bar?"
        strWhatIfBar = InfBox(strMsg, "?", "Last Bar|New Bar|Cancel", "What-If")
    End If
    
    If strWhatIfBar = "C" Then
        DeactivateWhatIf False
        SyncToolbar
    Else
        iBlankBarsSave = m.nBlankBars         '4671
        If m.nBlankBars < 10 Then
            MyBlankBars = 10
            MyForecastBars = 10
            GenerateChart eRedo9_ReloadData
        End If
                
        iLastBar = LastGoodDataBar(False)
        
        If strWhatIfBar = "N" Then
            m.Bars(eBARS_Open, iLastBar + 1) = m.Bars(eBARS_Close, iLastBar)
            m.Bars(eBARS_High, iLastBar + 1) = m.Bars(eBARS_Close, iLastBar)
            m.Bars(eBARS_Low, iLastBar + 1) = m.Bars(eBARS_Close, iLastBar)
            m.Bars(eBARS_Close, iLastBar + 1) = m.Bars(eBARS_Close, iLastBar)
            m.nLastGoodDataBarX = LastGoodDataBar(True) + 1
            iLastBar = iLastBar + 1
        End If
        
        dDate1 = m.Bars(eBARS_DateTime, iLastBar + 1)
        dDate2 = m.Bars(eBARS_DateTime, iLastBar + 6)
        dClose = m.Bars(eBARS_Close, iLastBar)
            
        Set Annot = m.Annots("WhatIf")
        
        If Annot Is Nothing Then
            Set Annot = New cAnnotation
            Annot.CreateNew Me, eANNOT_SimpleLine, 1, dDate1, dClose, dDate2, dClose, vbBlue, , , eANNOT_WhatIf
            m.Annots.Add Annot, "WhatIf"
        Else
            Annot.dDate(1) = dDate1
            Annot.dDate(2) = dDate2
            Annot.Y(1) = dClose
            Annot.Y(2) = dClose
        End If
                
        Annot.Prop("WhatIfBar") = strWhatIfBar
        Annot.Prop("WhatIfTemplate") = m.strTemplateApplied
        Annot.Prop("BlankBarsSave") = iBlankBarsSave
        
        m.nShowTrades = 0
        m.bDisableRT = True
        
        SyncToolbar
        GenerateChart eRedo1_Scrolled
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.ActivateWhatIf"
    
End Sub

Public Sub DeactivateWhatIf(Optional ByVal bRedraw As Boolean = True)
On Error GoTo ErrSection:
    
    Dim Annot As cAnnotation
    Dim eRedo As eChartRedoMode
    Dim i&
    
    Set Annot = m.Annots("WhatIf")
    
    i = LastGoodDataBar(False)
    
    If Not Annot Is Nothing Then
        If bRedraw And Not m.aWhatIfSave Is Nothing Then
            If Annot.Prop("WhatIfBar") = "N" Then
                m.Bars(eBARS_Open, i) = kNullData
                m.Bars(eBARS_High, i) = kNullData
                m.Bars(eBARS_Low, i) = kNullData
                m.Bars(eBARS_Close, i) = kNullData
                m.nLastGoodDataBarX = LastGoodDataBar(True) - 1
            Else
                m.Bars(eBARS_Open, i) = m.aWhatIfSave(0)
                m.Bars(eBARS_High, i) = m.aWhatIfSave(1)
                m.Bars(eBARS_Low, i) = m.aWhatIfSave(2)
                m.Bars(eBARS_Close, i) = m.aWhatIfSave(3)
                m.Bars.Prop(eBARS_LastTickTime) = m.aWhatIfSave(4)
            End If
        End If
        
        i = Annot.Prop("BlankBarsSave")         '5953
        
        Annot.geRemoveAnnotation m.geChartObj
        Annots.Remove "WhatIf"
        
        If m.nBlankBars = i Then
            eRedo = eRedo5_RecalcInd
        Else
            MyBlankBars = i
            MyForecastBars = i
            eRedo = eRedo9_ReloadData
        End If
    End If
    
    If eRedo = eRedo9_ReloadData Then
        RestoreChartNormal vbKeyClear       '4763
    ElseIf bRedraw And Not g.RealTime.Active Then
        GenerateChart eRedo5_RecalcInd
    End If
    
    m.bDisableRT = False
    
    'JM 12-16-2009: issue 4764 is same as issue 5522
    'calling generate chart with ReloadRT fixed issue 4764
    'but for some reason it is now necessary to use ReloadData for streaming to resume
    If g.RealTime.Active Then GenerateChart eRedo9_ReloadData       'eRedo7_ReloadRT     '4764
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.DeactivateWhatIf"
    
End Sub

Public Property Get IsInWhatIfMode() As Boolean
On Error GoTo ErrSection:

    If m.Annots("WhatIf") Is Nothing Then
        IsInWhatIfMode = False
    Else
        IsInWhatIfMode = True
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.IsInWhatIfMode"
    
End Property

Public Property Get AcctBarCols() As String
On Error GoTo ErrSection:

    AcctBarCols = m.strAcctBarCols
    
    Exit Property
    
ErrSection:
    RaiseError "cChart.AcctBarCols.Get"
    
End Property

Public Property Let AcctBarCols(ByVal strSettings$)
On Error GoTo ErrSection:

    m.strAcctBarCols = strSettings
    
    Exit Property
    
ErrSection:
    RaiseError "cChart.AcctBarCols.Let"
    
End Property

Public Sub EditEquityFilter()
On Error GoTo ErrSection:

    If frmEquityFilter.ShowMe(m.EquityFilter, Picture16("kBlank").Picture) Then
        m.Transactions.EquityFilter = m.EquityFilter
        m.Transactions.CalculateEquityFilter
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.EditEquityFilter"
    
End Sub

Private Function InvalidWhatIf(Optional ByVal bShowErrMsg As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strErr$

    If Not m.Bars Is Nothing Then
        Select Case m.Bars.Prop(eBARS_PeriodType)
            Case ePRD_IntPF, ePRD_EodPF
                strErr = "What-If mode will not work with Point & Figure."
            Case ePRD_IntKagi, ePRD_EodKagi
                strErr = "What-If mode will not work with Kagi bars."
            Case ePRD_IntRenko, ePRD_EodRenko
                strErr = "What-If mode will not work with Renko bars."
        End Select
    End If

    If Len(strErr) > 0 Then
        InvalidWhatIf = True
        If bShowErrMsg Then
            InfBox strErr, "I", , "What-If Mode"
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.InvalidWhatIf"
    
End Function

Private Sub SetTradeSymbolInfo()
On Error GoTo ErrSection:
    
    Dim strSymbol As String
    
    If m.nSymbolID > 0 And m.Bars.Size > 0 Then         '5112
        strSymbol = Symbol
        If InStr(strSymbol, "-055") <> 0 Or InStr(strSymbol, "-057") <> 0 Or _
           InStr(strSymbol, "-065") <> 0 Or InStr(strSymbol, "-067") <> 0 Then
           
            If g.nReplaySession > 0 Or frmReplay.Visible Then
                m.strTradeSymbol = RollSymbolForDate(strSymbol, m.Bars(eBARS_DateTime, m.Bars.Size - 1))        '6030
            Else
                m.strTradeSymbol = RollSymbolForDate(strSymbol) ', m.Bars(eBARS_DateTime, m.Bars.Size - 1))        '4844
            End If
            '6030 - for some reason this call returns 0 for the symbol ID for this aardvark issue
            m.nTradeSymbolID = g.SymbolPool.SymbolIDforSymbol(m.strTradeSymbol)
        Else
            m.nTradeSymbolID = m.nSymbolID
            m.strTradeSymbol = Symbol
        End If
    Else
        m.nTradeSymbolID = 0
        m.strTradeSymbol = ""
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.SetTradeSymbol"
    
End Sub

Public Property Get TradeSymbolID(Optional ByVal bSymOrSymID As Boolean = False) As Variant
On Error GoTo ErrSection:

    If m.nTradeSymbolID = 0 Then SetTradeSymbolInfo
    
    If bSymOrSymID And m.nTradeSymbolID = 0 And Len(m.strTradeSymbol) > 0 Then
        TradeSymbolID = m.strTradeSymbol        '6030 - see note in SetTradeSymbolInfo
    Else
        TradeSymbolID = m.nTradeSymbolID
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.TradeSymbolID.Get"

End Property

Public Property Get TradeSymbol() As String
On Error GoTo ErrSection:

    If Len(m.strTradeSymbol) = 0 Then SetTradeSymbolInfo
    TradeSymbol = m.strTradeSymbol

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.TradeSymbol.Get"

End Property

Public Sub BidAskPropGet(eMode As eBidAskColorMode, eImage As eStockImage, _
    eDir As eImageDir, nFill&, nSize&, nBidColor&, nAskColor&)
On Error GoTo ErrSection:

    eMode = m.eBidAskMode
    eImage = m.eBidAskImage
    eDir = m.eBidAskImgDir
    nFill = m.nBidAskImgFill
    nSize = m.nBidAskImgSize
    nBidColor = m.nBidColor
    nAskColor = m.nAskColor

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.BidAskPropGet"

End Sub

Public Sub BidAskPropLet(ByVal eMode As eBidAskColorMode, ByVal eImage As eStockImage, _
    ByVal eDir As eImageDir, ByVal nFill&, ByVal nSize&, ByVal nBidColor&, ByVal nAskColor&)
On Error GoTo ErrSection:

    Dim bChanged As Boolean

    If m.eBidAskImage <> eImage Or _
       m.eBidAskImgDir <> eDir Or _
       m.nBidAskImgFill <> nFill Or _
       m.nBidAskImgSize <> nSize Or _
       m.nBidColor <> nBidColor Or _
       m.nAskColor <> nAskColor Then
       
       bChanged = True
    
    End If
       

    m.eBidAskMode = eMode
    m.eBidAskImage = eImage
    m.eBidAskImgDir = eDir
    m.nBidAskImgFill = nFill
    m.nBidAskImgSize = nSize
    m.nBidColor = nBidColor
    m.nAskColor = nAskColor

    If Not m.BidAskAnnot Is Nothing Then
        With m.BidAskAnnot
            .PreIndicator = 1
            .Style = -1
            .Color = m.nBidColor
            .Prop("AskColor") = m.nAskColor
            .Prop("ImageType") = m.eBidAskImage
            .Prop("ImageSize") = m.nBidAskImgSize
            .Prop("ImageStyle") = m.nBidAskImgFill
            .Prop("ImageDir") = m.eBidAskImgDir
        End With
        If bChanged Then BidAskOnChart False, True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.BidAskPropLet"

End Sub

Public Sub RestoreChartNormal(ByVal KeyAscii As Integer)
On Error GoTo ErrSection:

    Dim i&
    Dim Pane As cPane
    Dim eBarMode As eOrderBarMode
    
    eBarMode = m.Form.OrderBarMode
    
    If KeyAscii = vbKeyEnd Or KeyAscii = vbKeyClear Then
        Set Pane = m.Tree("PRICE PANE")
        If Not Pane Is Nothing Then
            If Pane.Scaling = ePANE_ScaleModeManual Then Pane.CenterData Me
        End If
        m.geChartStruct.scaleArrow = 0
        If TypeOfChart = eTypeChart_Seasonal Then
            SeasonalIndClear
            RedoMode = eRedo7_ReloadRT
        ElseIf KeyAscii = vbKeyClear And m.nBlankBars <> m.nForecastBars Then
            m.nForecastBars = m.nBlankBars                  'vbkeyclear means user hit the red arrow
            GenerateChart eRedo9_ReloadData, , False        'Pete wants all extra bars beyond blank bars removed
        Else
            m.hsb.Value = m.nLastGoodDataBarX + m.nBlankBars
            GenerateChart eRedo1_Scrolled
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        'New spec from Pete:
        '   Reset tool button and esc is hotkey.
        '   Go to end of chart to comply with the Fixed Blank bar settings.
        '   Remove all extra forecast bars that were added by move tool
        '   Rescale the vertical axis and also default spacing.
        '   Return key also resets to auto scale mode or Last used state for the chart..
        
        If m.bZoomed Then UnzoomChart           'user used the ResetChart button
        
        If IsInWhatIfMode Then DeactivateWhatIf
        
        Set Pane = m.Tree("PRICE PANE")
        If Not Pane Is Nothing Then
            Pane.geClearMaxRatio
            Pane.geClearMinRatio
            Pane.geIncDecMaxRatio 0.03 ' 0.05
            Pane.geIncDecMinRatio -0.03 ' -0.05
            If Pane.Scaling <> ePANE_ScaleModeAuto And Pane.Scaling <> ePANE_ScaleModeAutoPrice Then
                Pane.Scaling = Pane.ScaleAutoLastUsed
            End If
            If TypeOfChart = eTypeChart_Seasonal Then
                SeasonalIndClear
                RedoMode = eRedo7_ReloadRT
            Else
                m.hsb.Value = m.nLastGoodDataBarX + m.nBlankBars
            End If
        End If
        Set Pane = Nothing
        ExtraPriceScale = 0
        If m.nDefaultPixelsPerBar > 0 Then
            m.nPixelsPerBar = m.nDefaultPixelsPerBar
        Else
            m.nPixelsPerBar = 7
        End If
        m.geChartStruct.x_dataPointPix = 7
        m.geChartStruct.scaleArrow = 0
        If m.nBlankBars <> m.nForecastBars Then
            m.nForecastBars = m.nBlankBars
            GenerateChart eRedo9_ReloadData, , False
        Else
            geForceRecalc
            GenerateChart eRedo1_Scrolled
        End If
        SyncToolbar
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.RestoreChartNormal"

End Sub

Public Sub BidAskOnChart(Optional ByVal bClear As Boolean = False, Optional ByVal bPropChanged As Boolean = False)
On Error GoTo ErrSection:

    Static bPrevShowHide As Boolean
    Static oBidAsk As cBidAskDepth
    Static nPrevSymID As Long
    Static dLastChecked As Double
    
    Dim aBids As cGdArray, aAsks As cGdArray
    Dim Pane As cPane
    
    Dim Y1#, Y2#, dDate1#, dDate2#, dBidPrice#, dAskPrice#
    Dim nLargestSize&, nLargestBid&, nLargestAsk&
    Dim i&, iLastDataBar&, iSymID&, strSym$
    
    Dim bRefresh As Boolean
    
    If m.eBidAskMode = eBidAskColorMode_None Then
        bClear = True
    Else
        iLastDataBar = LastGoodDataBar(False, False)
        If gdGetNum(m.geChartStruct.gdhDate, gdGetSize(m.geChartStruct.gdhDate) - 1) < _
            m.Bars(eBARS_DateTime, iLastDataBar) Then bClear = True
    End If

    If bClear Then
        If bPrevShowHide = True Then
            GoTo ErrExit
        Else
            Exit Sub
        End If
    End If
        
    If Not g.RealTime.Active Or m.Bars Is Nothing Then
        If oBidAsk Is Nothing Then
            Exit Sub
        Else
            Set oBidAsk = Nothing
            GoTo ErrExit
        End If
    End If
    
    If oBidAsk Is Nothing Then
        If nPrevSymID = 0 Or nPrevSymID <> m.nSymbolID Then
            bRefresh = True
        ElseIf gdTickCount(False) - dLastChecked > 3000 Then
            bRefresh = True
        End If
        
        If bRefresh Then
            strSym = m.Bars.Prop(eBARS_Symbol)
            If Left(strSym, 1) = "$" Then
                'do nothing (index & forex do not have depth of market)
            Else
                Set oBidAsk = g.RealTime.AddMarketDepthSymbol(strSym, False)
                If oBidAsk Is Nothing And SecurityType(strSym) = "F" Then
                    If InStr(strSym, "-0") = 0 Then
                        strSym = m.Bars.Prop(eBARS_BaseSymbol) & "-067"
                    Else
                        strSym = RollSymbolForDate(m.nSymbolID) ', m.Bars(eBARS_DateTime, m.Bars.Size - 1))   'chart has continuous, try current contract
                    End If
                    Set oBidAsk = g.RealTime.AddMarketDepthSymbol(strSym, False)
                End If
                If Not oBidAsk Is Nothing Then g.RealTime.RemoveMarketDepthSymbol strSym
            End If
            dLastChecked = gdTickCount(False)
        End If
        
        nPrevSymID = m.nSymbolID

    ElseIf oBidAsk.UsageCounter = 0 Then
        Set oBidAsk = Nothing           'ladder or depth of market got closed or symbol got changed
    ElseIf m.Bars.Prop(eBARS_SymbolID) <> nPrevSymID Then
        Set oBidAsk = Nothing           'symbol changed on chart
    End If
    
    If oBidAsk Is Nothing Then
        If bPrevShowHide = False Then
            Exit Sub
        Else
            GoTo ErrExit
        End If
    End If
        
    bRefresh = False
    
    Set aBids = oBidAsk.BestBids
    If aBids Is Nothing Then GoTo ErrExit
    
    Set aAsks = oBidAsk.BestAsks
    If aAsks Is Nothing Then GoTo ErrExit
    oBidAsk.LargestBidAskSize nLargestSize, nLargestBid, nLargestAsk, dBidPrice, dAskPrice
    
    Set Pane = m.Tree("PRICE PANE")
    If Pane Is Nothing Then GoTo ErrExit
    
    If dDate2 > dDate1 Then
        dDate1 = gdFixDateTime(m.Bars(eBARS_DateTime, iLastDataBar + 1))    'one bar to right of last good databar
        dDate2 = gdFixDateTime(m.Bars(eBARS_DateTime, iLastDataBar + 2))
    End If

    If m.BidAskAnnot Is Nothing Then
        Set m.BidAskAnnot = New cAnnotation
        m.BidAskAnnot.CreateNew Me, eANNOT_Icon, Pane.gePaneId, dDate1, 0, dDate2, 0, , , , eANNOT_BidAskChart
        If Not m.BidAskAnnot Is Nothing Then
            With m.BidAskAnnot
                .PreIndicator = 1
                .Style = -1
                .Color = m.nBidColor
                .Prop("AskColor") = m.nAskColor
                .Prop("ImageType") = m.eBidAskImage
                .Prop("ImageSize") = m.nBidAskImgSize
                .Prop("ImageStyle") = m.nBidAskImgFill
                .Prop("ImageDir") = m.eBidAskImgDir
            End With
        End If
    End If
    
    If Not m.BidAskAnnot Is Nothing Then
        With m.BidAskAnnot
            Y1 = .Y(1)
            Y2 = .Y(2)
            If m.eBidAskMode = eBidAskColorMode_BySize Then
                .Y(1) = dBidPrice
                .Y(2) = dAskPrice
            Else
                .Y(1) = aBids(0)
                .Y(2) = aAsks(0)
            End If
            If m.Bars.IsIntraday Then
                .dDate(1) = dDate1
                .dDate(2) = dDate1
            Else
                'offset ask position by 1 bar because the prices are often too close together
                .dDate(1) = dDate2      'bid
                .dDate(2) = dDate1      'ask (will show to left of bid)
            End If
            
            If Y1 <> .Y(1) Then
                bRefresh = True
            ElseIf Y2 <> .Y(2) Then
                bRefresh = True
            End If
            
            If bPrevShowHide = False Then
                .BidAskChartShowHide True
                .geDrawAnn Me
            ElseIf bRefresh Or bPropChanged Then
                .geDrawAnn Me
            End If
            bPrevShowHide = True
        End With
    End If
    
ErrExit:
    If oBidAsk Is Nothing Or bClear Then
        If Not m.BidAskAnnot Is Nothing Then
            m.BidAskAnnot.BidAskChartShowHide False
            If bPrevShowHide = True Then
                m.BidAskAnnot.geDrawAnn Me
                bPrevShowHide = False
            End If
        End If
    End If
    
    Set aBids = Nothing
    Set aAsks = Nothing
    Set Pane = Nothing
        
    Exit Sub

ErrSection:
    RaiseError "cChart.BidAskOnChart"
    Resume ErrExit

End Sub

Public Property Get DefaultPixelsPerBar(frm As Form) As Long
On Error GoTo ErrSection:

    Dim nPix As Long

    If m.nDefaultPixelsPerBar > 0 Then
        nPix = m.nDefaultPixelsPerBar
    Else
        nPix = 7             'this is hard-coded default
    End If
    
    If Not frm Is Nothing Then
        If TypeOf frm Is frmChartPixPerBar Then m.nPixelsPerBar = nPix
    End If
    
    DefaultPixelsPerBar = nPix

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.DefaultPixelsPerBar.Get"

End Property


Public Property Let DefaultPixelsPerBar(frm As Form, nPix As Long)
On Error GoTo ErrSection:

    If TypeOf frm Is frmChartPixPerBar Then
        m.nDefaultPixelsPerBar = nPix
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.DefaultPixelsPerBar.Let"

End Property

' TLB 5/12/2008: temporary routine to help debug the new Last Bar Recalc method
Public Sub VerifyResults()
On Error GoTo ErrSection:

    m.iVerifyResults = 1
    GenerateChart eRedo5_RecalcInd 'eRedo6_ReloadInd
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.VerifyResults"
    Resume ErrExit
End Sub

Public Property Get UseNewRecalcMethod() As Boolean
    UseNewRecalcMethod = m.bUseNewRecalcMethod
End Property

Public Property Let UseNewRecalcMethod(ByVal bUseNewRecalcMethod As Boolean)
    If bUseNewRecalcMethod And Not HasNewNavEngine Then
        bUseNewRecalcMethod = False
    End If
    If m.bUseNewRecalcMethod <> bUseNewRecalcMethod Then
        m.bUseNewRecalcMethod = bUseNewRecalcMethod
        m.dAvgRecalcTime = 0
        RedoMode = eRedo9_ReloadData
    End If
End Property

Public Property Get IsPartiallyLoaded() As Boolean
    If m.nPartialLoadDate > 0 Then
        IsPartiallyLoaded = True
    End If
End Property

Public Property Let IsPartiallyLoaded(ByVal bPartiallyLoaded As Boolean)
    If Not bPartiallyLoaded Then
        m.nPartialLoadDate = 0
    End If
End Property

Private Sub LoadEwave()
On Error GoTo ErrSection:

    Dim aFileContent As cGdArray
    Dim strChar As String
    Dim i&, nDegree&, dDate#

    If Not m.bEWaveFile Then Exit Sub
    
    If m.aEWave Is Nothing Then Set m.aEWave = New cGdArray
    If m.aEWave.Size > 0 Then Exit Sub          'already done
    
    Set aFileContent = New cGdArray
    aFileContent.FromFile App.Path & "\EWAVE\filter.txt"
    If aFileContent.Size > 0 Then
        strChar = Left(aFileContent(0), 1)
        If IsDigit(strChar) Then m.nDegreeFilter = Val(strChar)
    End If
    If m.nDegreeFilter <= 0 Or m.nDegreeFilter > 9 Then m.nDegreeFilter = 3
    
    aFileContent.FromFile App.Path & "\EWAVE\dji.txt"
    
    For i = 0 To aFileContent.Size - 1
        strChar = Parse(aFileContent(i), ",", 2)
        If Len(strChar) >= 8 Then
            nDegree = Parse(aFileContent(i), ",", 4)
            If nDegree <= m.nDegreeFilter Then
                strChar = Left(strChar, 8)
                dDate = DateOf(strChar)
                m.aEWave.Add dDate & "," & Parse(aFileContent(i), ",", 3) & "," & nDegree & "," & Parse(aFileContent(i), ",", 5) & "," & strChar
            End If
        End If
    Next
    
    m.aEWave.Sort eGdSort_Default, 0, m.aEWave.Size
    m.aEWave.ToFile "EWaveTest.txt"

    Set aFileContent = Nothing
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.LoadEwave"

End Sub

Public Sub ResetSplitPane()
On Error GoTo ErrSection:

    m.geChartStruct.splitPaneWidth = 0

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.ResetSplitPane"
End Sub

Private Function PriceAlertAnnotNew(AlertIn As cAlert) As Boolean

    Dim Annot As cAnnotation
    
    If AlertIn Is Nothing Then Exit Function
    If AlertIn.AlertType <> eGDAlertType_Price Then Exit Function
    If AlertIn.Symbol <> Symbol Then Exit Function
    If Not m.Annots(AlertIn.AlertKey) Is Nothing Then Exit Function     'already added

    Set Annot = New cAnnotation
    Annot.CreateNew Me, eANNOT_Icon, 1, m.Bars(eBARS_DateTime, m.Bars.Size - 1), 0, m.Bars(eBARS_DateTime, m.Bars.Size - 1), 0, vbBlack, , , eANNOT_PriceAlert
    Annot.AlertObject = AlertIn
    Annot.UpdateAlert 2, True

    'check that annot successfully initialized
    Set Annot = m.Annots(AlertIn.AlertKey)
    If Not Annot Is Nothing Then
        PriceAlertAnnotNew = True
    End If
    
    Set Annot = Nothing

End Function

Public Function PriceAlertAdd(Optional AlertIn As cAlert = Nothing) As Long
'JM: Do not call GenerateChart in this routine because it may be
'    too soon and Trade Navigator will crash on spread charts
    
    Dim Alert As cAlert, iCount&, i&

    If g.bUnloading Then Exit Function
    
    iCount = 0
    If AlertIn Is Nothing And Not m.bPriceAlertsChecked Then
        'only want to walk through alerts collection on template load or symbol change
        'always add a price alert object regardless of whether it is shown on charts or not
        'the geSetPoints routine will check whether the bells should be drawn
        For i = 1 To g.Alerts.Count
            Set Alert = g.Alerts(i)
            'tell alert object to add this chart to its list of charts
            'to notify when user update the alert object or remove it
            If PriceAlertAnnotNew(Alert) Then
                Alert.AddChartObject Me
                iCount = iCount + 1
            End If
        Next
    ElseIf m.Annots(AlertIn.AlertKey) Is Nothing Then
        If PriceAlertAnnotNew(AlertIn) Then iCount = iCount + 1
    End If
    
    m.bPriceAlertsChecked = True
    
    PriceAlertAdd = iCount

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.PriceAlertAdd"
    
End Function

Public Function PriceAlertUpdate(AlertIn As cAlert, ByVal bRemove As Boolean) As Boolean
On Error GoTo ErrSection:
'JM: Do not call GenerateChart in this routine because it may be
'    too soon and Trade Navigator will crash on spread charts

    Dim Annot As cAnnotation
    Dim bUpdated As Boolean
    
    If g.bStarting Or g.bUnloading Or m.Annots Is Nothing Then Exit Function
    
    If Not AlertIn Is Nothing Then
        Set Annot = m.Annots(AlertIn.AlertKey)
        If Not Annot Is Nothing Then
            If bRemove Then
                Annot.geRemoveAnnotation m.geChartObj
                m.Annots.Remove AlertIn.AlertKey
            Else
                Annot.UpdateAlert 2, True
                If FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid
            End If
            bUpdated = True
        End If
    End If
    
    PriceAlertUpdate = bUpdated

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.PriceAlertUpdate"
    
End Function

Public Property Get PriceAlertsChecked() As Boolean
    PriceAlertsChecked = m.bPriceAlertsChecked
End Property

Public Property Get PriceTopMost() As Long
    PriceTopMost = m.nPriceTopMost
End Property

Public Property Let PriceTopMost(ByVal nTopMost&)
    m.nPriceTopMost = nTopMost
End Property

Public Function ChainFromFile(bNewChain As Boolean, ByVal bReload As Boolean) As Boolean
On Error GoTo ErrSection:

    Static strPrevBaseSym$
    Static nTries&

    Dim i&, lNumWait&, iColor&
    Dim strFile$, strText$, strBaseSym$
    
    Dim dToday#, dFileDate#
    Dim dOptMonth#, dOptExpire#
    
    Dim aInfo As New cGdArray
    Dim aFile As New cGdArray
    
    Dim Annot As cAnnotation
    Dim bFirstAlreadySet As Boolean
    
    
    strBaseSym = Parse(Symbol, "-", 1)
    strFile = AddSlash(FilePath(g.strAppPath))
    
    If InStr(Symbol, "-") = 0 Then
        strFile = strFile & "OptionNav\Chains\" & strBaseSym & ".dat"
    Else
        strFile = strFile & "OptionNav\chains\" & strBaseSym & "-.dat"
        strBaseSym = strBaseSym & "-"
    End If
    
    
    If FileExist(strFile) Then      'check file date
        dFileDate = FileDate(strFile)
        dToday = Now
        If Int(dFileDate) < Int(dToday) Then
            KillFile strFile, True
        End If
    End If
    
    If strBaseSym = m.strOptionSym And FileExist(strFile) Then
        bNewChain = False
        ChainFromFile = True
        If Not bReload Then Exit Function       'already done
    End If

    'reset
    If Len(m.strOptionSym) > 0 Then
        RemoveAnnots False, eANNOT_VertLine, eANNOT_OptionInfo
        GenerateChart eRedo1_Scrolled
        DoEvents
    End If
    m.strOptionSym = ""
    
    If Not FileExist(strFile) Then
        If strPrevBaseSym = strBaseSym Then
            If nTries < 3 Then
                nTries = nTries + 1
                RequestOptionChainStructure strBaseSym
                StatusMsg "Requesting option chains ..."
            Else
                nTries = 0
                m.Form.OptionDataAvailable eGDOptNav_ChainBuilt, strBaseSym & vbTab & ";" & strFile & " not found. Request for option chain failed."
            End If
        Else
            nTries = 1
            RequestOptionChainStructure strBaseSym
            StatusMsg "Requesting option chains ..."
            strPrevBaseSym = strBaseSym
        End If
        
        Exit Function
    End If
    
    bNewChain = True
    aFile.FromFile strFile
    
    If g.ChartGlobals.nChartBackColor = 0 Then
        iColor = RGB(0, 254, 0)
    ElseIf ChartBkIsLight Then      '5226
        iColor = vbBlue
    Else
        iColor = vbWhite
    End If
    
    For i = 1 To aFile.Size - 1
        If Parse(aFile(i), "=", 1) = "DATE" Then
            If aInfo.Size > 0 And Not Annot Is Nothing Then
                Annot.LoadOptionsInfo aInfo
                Annots.Add Annot
                m.strOptionSym = strBaseSym                 'indicate data already available for this symb
            End If
            
            Set Annot = Nothing
            aInfo.Size = 0
            
            If Parse(aFile(i), ";", 3) = 4 Then
                strText = Parse(aFile(i), "=", 2)
                dOptMonth = DateOf(Left(Parse(strText, ";", 1), 6) & "01")
                dOptExpire = DateOf(Left(Parse(strText, ";", 2), 8))
                If dOptMonth > 0 And dOptExpire > 0 Then
                    Set Annot = New cAnnotation
                    Annot.CreateNew Me, eANNOT_VertLine, 1, dOptExpire, 0, dOptMonth, 0, iColor, , , eANNOT_OptionInfo
                    If Not bFirstAlreadySet Then
                        Annot.FirstOptionAnnnot = True
                        bFirstAlreadySet = True
                    End If
                End If
            Else
                'skip over strike prices for options that have expiration type <> 4 (TNE is an example)
                i = i + 1
                Do While Parse(aFile(i), "=", 1) = "STRIKE"
                    If i >= aFile.Size - 1 Then
                        Exit Do
                    Else
                        i = i + 1
                    End If
                Loop
                If i < aFile.Size - 1 Then i = i - 1 'set back one so we don't miss the DATE line
            End If
            
        ElseIf Parse(aFile(i), "=", 1) = "STRIKE" Then
            aInfo.Add Parse(aFile(i), "=", 2)
        End If
    Next
    
    'add the last one
    If aInfo.Size > 0 And Not Annot Is Nothing Then
        Annot.LoadOptionsInfo aInfo
        Annots.Add Annot
        m.strOptionSym = strBaseSym                 'indicate data available for this symb
    End If
    
    ChainFromFile = True
    
    Set aFile = Nothing
    Set aInfo = Nothing
    Set Annot = Nothing
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ChainFromFile"

End Function

Public Function ClosestOptionAnnot(ByVal dDate#) As cAnnotation
On Error GoTo ErrSection:

    Dim i&, iDiff&, dAnnotDate#
    
    Dim Annot As cAnnotation
    Dim AnnotOut As cAnnotation
        
    iDiff = 999999
    
        For i = 1 To Annots.Count
            Set Annot = Annots(i)
            If Annot.eUsage = eANNOT_OptionInfo Then
                If Annot.eType = eANNOT_VertLine Then
                    If dDate = 0 Then
                        If Annot.IsBalloonOptionsInfo Then
                            Set AnnotOut = Annot
                            Exit For
                        End If
                    Else
                        dAnnotDate = Annot.dDate(1)
                        If dAnnotDate = dDate Then
                            Set AnnotOut = Annot
                            Exit For
                        ElseIf Abs(dDate - dAnnotDate) < iDiff Then
                            Set AnnotOut = Annot
                            iDiff = Abs(dDate - dAnnotDate)
                        End If
                    End If
                End If
            End If
        Next
        
    Set ClosestOptionAnnot = AnnotOut
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.ClosestOptionAnnot"

End Function

Private Sub OptNavGraphData(ByVal strMsg$, aPrice As cGdArray, aNow As cGdArray, _
    aThird As cGdArray, aTwoThird As cGdArray, aExpire As cGdArray, _
    aProb1 As cGdArray, aProb2 As cGdArray, strGraphType$)
    
    Dim iSize&
    Dim strPrice$, strNow$, strThird$, strTwoThird$, strExpire$
    Dim strProb1$, strProb2$
    
    Dim bArraysCreated As Boolean
    Dim bErr As Boolean
    
    Set aPrice = Nothing
    Set aNow = Nothing
    Set aThird = Nothing
    Set aTwoThird = Nothing
    Set aExpire = Nothing
    
    Set aProb1 = Nothing
    Set aProb2 = Nothing
        
    If Len(strMsg) = 0 Then
        DebugLog "Risk graph error: appmail message string is empty"
        GoTo ErrExit
    End If
    
    strGraphType = Parse(strMsg, vbTab, 2)
    If strGraphType <> "ProfitLoss" And strGraphType <> "Delta" And strGraphType <> "Theta" _
        And strGraphType <> "Gamma" And strGraphType <> "Vega" And strGraphType <> "Probability" Then
        DebugLog "Risk graph error: unknown graph type - " & strGraphType
        GoTo ErrExit
    End If
    
    Dim aMsg As New cGdArray
    
    aMsg.SplitFields strMsg, vbTab
        
    strPrice = aMsg(2)  'Parse(strMsg, vbTab, 3)      'price scale
    strNow = aMsg(3)    'Parse(strMsg, vbTab, 4)      'x-values for now to expiration
    strThird = aMsg(4)  'Parse(strMsg, vbTab, 5)      'x-values for 1/3 to expiration
    strTwoThird = aMsg(5)   'Parse(strMsg, vbTab, 6)  'x-values for 2/3 to expiration
    strExpire = aMsg(6)     'Parse(strMsg, vbTab, 7)  'x-values at expiration
    
    If strGraphType = "Probability" Then
        strProb1 = aMsg(7)
        strProb2 = aMsg(8)
    End If
    
    If Len(strPrice) = 0 Then
        DebugLog "Risk graph error: missing price scale data points"
        GoTo ErrExit:
    ElseIf Len(strNow) = 0 Then
        DebugLog "Risk graph error: missing NOW data points"
        GoTo ErrExit:
    ElseIf Len(strThird) = 0 Then
        DebugLog "Risk graph error: missing 1/3 data points"
        GoTo ErrExit:
    ElseIf Len(strTwoThird) = 0 Then
        DebugLog "Risk graph error: missing 2/3 data points"
        GoTo ErrExit:
    ElseIf Len(strExpire) = 0 Then
        DebugLog "Risk graph error: missing expiration data points"
        GoTo ErrExit:
    End If
    
    If strGraphType = "Probability" Then
        If Len(strProb1) = 0 Then
            DebugLog "Risk graph error: missing probability data points for 5th array"
            GoTo ErrExit:
        ElseIf Len(strProb1) = 0 Then
            DebugLog "Risk graph error: missing probability data points for 5th array"
            GoTo ErrExit:
        End If
    End If
        
    Set aPrice = New cGdArray
    Set aNow = New cGdArray
    Set aThird = New cGdArray
    Set aTwoThird = New cGdArray
    Set aExpire = New cGdArray
    
    aPrice.Create eGDARRAY_Doubles
    aNow.Create eGDARRAY_Doubles
    aThird.Create eGDARRAY_Doubles
    aTwoThird.Create eGDARRAY_Doubles
    aExpire.Create eGDARRAY_Doubles
    
    If strGraphType = "Probability" Then
        Set aProb1 = New cGdArray
        Set aProb2 = New cGdArray
        aProb1.Create eGDARRAY_Doubles
        aProb2.Create eGDARRAY_Doubles
    End If
    
    bArraysCreated = True
    
    aPrice.SplitFields strPrice, ";"
    iSize = aPrice.Size
    
    aNow.SplitFields strNow, ";"
    If aNow.Size <> iSize Then
        DebugLog "Risk graph error: price array size (" & Str(iSize) & ") NOW array size (" & Str(aNow.Size) & ")"
        bErr = True
        GoTo ErrExit
    End If
    
    aThird.SplitFields strThird, ";"
    If aThird.Size <> iSize Then
        DebugLog "Risk graph error: price array size (" & Str(iSize) & ") 1/3 array size (" & Str(aThird.Size) & ")"
        bErr = True
        GoTo ErrExit
    End If
    
    aTwoThird.SplitFields strTwoThird, ";"
    If aTwoThird.Size <> iSize Then
        DebugLog "Risk graph error: price array size (" & Str(iSize) & ") 2/3 array size (" & Str(aTwoThird.Size) & ")"
        bErr = True
        GoTo ErrExit
    End If
    
    aExpire.SplitFields strExpire, ";"
    If aExpire.Size <> iSize Then
        DebugLog "Risk graph error: price array size (" & Str(iSize) & ") EXPIRE array size (" & Str(aExpire.Size) & ")"
        bErr = True
        GoTo ErrExit
    End If
    
    If strGraphType = "Probability" Then
        aProb1.SplitFields strProb1, ";"
        If aProb1.Size <> iSize Then
            DebugLog "Risk graph error: price array size (" & Str(iSize) & ") PROB1 array size (" & Str(aProb1.Size) & ")"
            bErr = True
            GoTo ErrExit
        End If
        
        aProb2.SplitFields strProb2, ";"
        If aProb2.Size <> iSize Then
            DebugLog "Risk graph error: price array size (" & Str(iSize) & ") PROB2 array size (" & Str(aProb2.Size) & ")"
            bErr = True
            GoTo ErrExit
        End If
        
        InfBox ""       'this infbox was displayed by frmChart when requesting Probability
    End If
    
    
ErrExit:
    If bArraysCreated And bErr Then
        If Not aPrice Is Nothing Then aPrice.Destroy
        If Not aNow Is Nothing Then aNow.Destroy
        If Not aThird Is Nothing Then aThird.Destroy
        If Not aTwoThird Is Nothing Then aTwoThird.Destroy
        If Not aExpire Is Nothing Then aExpire.Destroy
        
        If Not aProb1 Is Nothing Then aProb1.Destroy
        If Not aProb2 Is Nothing Then aProb2.Destroy
        
        Set aPrice = Nothing
        Set aNow = Nothing
        Set aThird = Nothing
        Set aTwoThird = Nothing
        Set aExpire = Nothing
        Set aProb1 = Nothing
        Set aProb2 = Nothing
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cChart.OptNavGraphData"
    Resume ErrExit
    
End Sub

Public Sub OptNavGraphInfoSet(ByVal strMsg$)
On Error GoTo ErrSection:

    Dim strGraphType$

    Dim Pane As cPane
    
    Dim aPrice As cGdArray          'y-axis values
    
    'for non-probability graph, these arrays mean x-values for now, 1/3 to expiration, 2/3 to expiration & at expiration
    'these are the dates that are seen in the tool tip of the graph
    Dim aNow As cGdArray
    Dim aThird As cGdArray
    Dim aTwoThird As cGdArray
    Dim aExpire As cGdArray
    
    'for Probability graph, the x-value arrays are paired as above/below 1,2,3 standard deviation (68%, 95% & 99%)
    'the aNow/aThird = 1st pair, aTwoThird/aExpire = 2nd pair & aProbValues1/aProbValues2 = 3rd pair
    Dim aProb1 As cGdArray
    Dim aProb2 As cGdArray
    
    'locate price pane
    Set Pane = m.Tree("PRICE PANE")
    If Pane Is Nothing Then Exit Sub
    
    'validate graph data
    OptNavGraphData strMsg, aPrice, aNow, aThird, aTwoThird, aExpire, aProb1, aProb2, strGraphType
    If aPrice Is Nothing Or aNow Is Nothing Or aThird Is Nothing Or aTwoThird Is Nothing Or aExpire Is Nothing Then Exit Sub

    m.nShowSplitPane = 1
    Pane.OptNavGraphInfoSet Me, aPrice, aNow, aThird, aTwoThird, aExpire, aProb1, aProb2, strGraphType
    
    GenerateChart eRedo3_Settings
    
    geForceRecalc
    GenerateChart eRedo1_Scrolled
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.OptNavGraphInfoSet"

End Sub

Public Property Get ShowToolbar() As Long
    ShowToolbar = m.nShowToolbar
End Property

Public Property Let ShowToolbar(ByVal nShow&)
    m.nShowToolbar = nShow
End Property

Public Sub GetSectorInfo(strSector$, strSub$, strCmp$, strSecDesc$, strSubDesc$, strCmpDesc$)
On Error GoTo ErrSection:

    Dim strSymbol$, nSymbolID&
    
    strSymbol = GetSymbol(m.nSymbolID)
    If Left(strSymbol, 3) = "$--" Then
        strSector = strSymbol
    ElseIf Left(strSymbol, 2) = "$-" Then
        strSub = strSymbol
        nSymbolID = SU_GetGroupParent(m.nSymbolID)
        If nSymbolID > 0 And nSymbolID < 999999 Then
            strSector = GetSymbol(nSymbolID)
        End If
    Else
        nSymbolID = SU_GetGroupParent(m.nSymbolID)
        If nSymbolID > 0 And nSymbolID < 999999 Then
            strCmp = strSymbol
            strSub = GetSymbol(nSymbolID)
            nSymbolID = SU_GetGroupParent(nSymbolID)
            If nSymbolID > 0 And nSymbolID < 999999 Then
                strSector = GetSymbol(nSymbolID)
            End If
        End If
    End If
    If Len(strSector) = 0 Then
        strSecDesc = "Sector"
        strSector = "Sector"
        strSub = ""
        strCmp = ""
    Else
        strSecDesc = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(strSector))
        strSector = g.SymbolPool.Symbol(g.SymbolPool.PoolRecForSymbol(strSector))       '5163
        If Len(strSector) = 0 Then
            strSecDesc = "Sector"
            strSector = "Sector"
            strSub = ""
            strCmp = ""
        End If
    End If
    If Len(strSub) = 0 Then
        strSubDesc = "Subsector"
        strSub = "Subsector"
        strCmp = ""
    Else
        strSubDesc = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(strSub))
        strSub = g.SymbolPool.Symbol(g.SymbolPool.PoolRecForSymbol(strSub))
        If Len(strSub) = 0 Then
            strSubDesc = "Subsector"
            strSub = "Subsector"
            strCmp = ""
        End If
    End If
    If Len(strCmp) = 0 Then
        strCmpDesc = "Component"
        strCmp = "Component"
    Else
        strCmpDesc = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(strCmp))
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.GetSectorInfo"

End Sub

Public Sub HandleHawkeyeButton(ByVal dDate#)
On Error GoTo ErrSection:
    
    Dim i&, iParentIndex&
    Dim dLastGoodDate As Double
    
    Dim Ind As cIndicator
    Dim IndAdds As cIndicator
    Dim IndLevels As cIndicator
    
    Dim Pane As cPane
    
    If m.Tree Is Nothing Then Exit Sub
    
    iParentIndex = m.Tree.RelativeIndex("PRICE", eTREE_Parent)
    
    If iParentIndex > 0 Then
        If m.Tree.NodeLevel(iParentIndex) = 0 Then
            Set Pane = m.Tree(iParentIndex)
        End If
    End If
    
    If Pane Is Nothing Then Exit Sub
    If Not Pane.Display Then Exit Sub
    
    For i = 1 To m.Tree.Count
        If m.Tree.NodeLevel(i) > 0 Then
            If m.Tree.RelativeIndex(i, eTREE_Parent) = iParentIndex Then
                Set Ind = m.Tree(i)
                If Ind.IsHawkeyeAdds Then
                    Set IndAdds = Ind
                    If Not IndLevels Is Nothing Then Exit For
                ElseIf Ind.IsHawkeyeLevels Then
                    Set IndLevels = Ind
                    If Not IndAdds Is Nothing Then Exit For
                End If
            End If
        End If
    Next
        
        
    i = LastGoodDataBar(False)
    dLastGoodDate = m.Bars(eBARS_DateTime, i)
        
    'Design note 09-17-2009:
    'add indicator(s) if they do not exist else just change date parm
    
    'Design note 10-14-2009
    '       commented out code that implemented original specs from Tim when when implemented
    '       "hiding" both on user right-clicked tool button or click past last good data bar
    'Original specs on from Tim
    '       if both exist and one is turned on the leave the other one turned off
    '       if both exist and both are turned off then turn both on
    
    If dDate <= 0 Or dDate > dLastGoodDate Then
        If Not IndAdds Is Nothing Then IndAdds.Display = False
        If Not IndLevels Is Nothing Then IndLevels.Display = False
        GoTo ErrExit
    End If
    
    If IndAdds Is Nothing Then
        Set Ind = New cIndicator
        With Ind
            .DisplayType = eINDIC_ValueMarkers
            .Display = True
            .Color = vbYellow
            .CodedName = "HawkeyeAdds"
            .Parm(3) = dDate
            .MarkerSize = 2
        End With
        If Ind.FunctionID <> 0 Then
            m.Tree.Add Ind, "", iParentIndex, eTREE_LastChild
        End If
    Else
        IndAdds.Parm(3) = dDate
        If Not IndAdds.Display Then IndAdds.Display = True
    End If
    
    If IndLevels Is Nothing Then
        Set Ind = New cIndicator
        With Ind
            .DisplayType = eINDIC_Line
            .Display = True
            .Color = vbWhite
            .CodedName = "HawkeyeLevels"
            .Parm(2) = dDate
        End With
        If Ind.FunctionID <> 0 Then
            m.Tree.Add Ind, "", iParentIndex, eTREE_LastChild
        End If
    Else
        IndLevels.Parm(2) = dDate
        If Not IndLevels.Display Then IndLevels.Display = True
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.HandleHawkeyeButton"

End Sub

Public Sub LinkedInputGet(ByVal strName$, ByRef strValue$, ByRef strType$)
On Error GoTo ErrSection:

    Dim i&
    Dim aInputs As New cGdArray
    
    If Len(strName) <= 0 Or Len(m.strLinkedInputs) <= 0 Then Exit Sub
    If InStr(m.strLinkedInputs, strName) = 0 Then Exit Sub
    
    aInputs.SplitFields m.strLinkedInputs, "|"
    
    For i = 0 To aInputs.Size - 1
        If strName = Parse(aInputs(i), ";", 1) Then
            strValue = Parse(aInputs(i), ";", 2)
            strType = Parse(aInputs(i), ";", 3)
            Exit For
        End If
    Next
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.LinkedInputVal.Get"

End Sub

Public Sub LinkedInputSet(ByVal strName$, ByVal strValue$, ByVal strType$, _
    Optional ByVal bUpdateIndicators As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, j&, strParm$
    Dim Ind As cIndicator
    Dim aInputs As New cGdArray
    
    If Len(strName) <= 0 Or Len(strValue) <= 0 Then Exit Sub
    
    If InStr(m.strLinkedInputs, strName) = 0 Then
        If Len(m.strLinkedInputs) = 0 Then
            m.strLinkedInputs = Trim(strName) & ";" & Trim(strValue) & ";" & Trim(strType)
        Else
            m.strLinkedInputs = Trim(m.strLinkedInputs) & "|" & Trim(strName) & ";" & Trim(strValue) & ";" & Trim(strType)
        End If
    Else
        aInputs.SplitFields m.strLinkedInputs, "|"
        For i = 0 To aInputs.Size - 1
            If strName = Parse(aInputs(i), ";", 1) Then
                aInputs(i) = Trim(strName) & ";" & Trim(strValue) & ";" & Trim(strType)
                m.strLinkedInputs = aInputs.JoinFields("|")
                Exit For
            End If
        Next
    End If
    
    If bUpdateIndicators Then
        For i = 1 To Tree.Count
            If Tree.NodeLevel(i) > 0 Then
                Set Ind = Tree(i)
                If Not Ind Is Nothing Then
                    Ind.LinkedParmsUpdate
                End If
            End If
        Next
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.LinkedInputVal.Let"

End Sub

Public Property Get LinkedInputString() As String
    LinkedInputString = m.strLinkedInputs
End Property

Public Property Get HighlightPos() As Long

    If m.nHighlightPos = -2 Then
        HighlightPos = 12632256    'property never set by user, default to gray
    Else
        HighlightPos = m.nHighlightPos
    End If
    
End Property

Public Property Let HighlightPos(ByVal nVal&)
    m.nHighlightPos = nVal
End Property

Public Property Get HighlightEquity() As Long

    If m.nHighlightEquity = -2 Then
        HighlightEquity = 12632256    'property never set by user, default to gray
    Else
        HighlightEquity = m.nHighlightEquity
    End If
    
End Property

Public Property Let HighlightEquity(ByVal nVal&)
    m.nHighlightEquity = nVal
End Property

Public Sub HandleJPButtons(ByVal strBtnID$)
On Error GoTo ErrSection:

    Dim i&, iFound&, iParentIndex&, idxSibling&
    Dim strPivPoint$, strPivResist$, strPivSupport$, strPivAvg$, strPeriodicity$, strFile$
    
    Dim Ind As cIndicator
    Dim IndPoint As cIndicator
    Dim IndResist As cIndicator
    Dim IndSupport As cIndicator
    Dim IndAvg As cIndicator
    
    If m.Tree Is Nothing Then Exit Sub
    
    iParentIndex = m.Tree.RelativeIndex("PRICE", eTREE_Parent)
    If iParentIndex <= 0 Then Exit Sub
    
    If InStr(strBtnID, "Daily") Then
        strPeriodicity = "Daily"
    ElseIf InStr(strBtnID, "Weekly") Then
        strPeriodicity = "Weekly"
    ElseIf InStr(strBtnID, "Monthly") Then
        strPeriodicity = "Monthly"
    ElseIf InStr(strBtnID, "Quarterly") Then
        strPeriodicity = "Quarterly"
    ElseIf InStr(strBtnID, "Expiration") Then
        strPeriodicity = "OptionsExp"
    Else
        Exit Sub
    End If
    
    strPivPoint = "JPPivotPoint" & strPeriodicity
    strPivResist = "JPPivotResistance" & strPeriodicity
    strPivSupport = "JPPivotSupport" & strPeriodicity
    strPivAvg = "JPAveragePivot" & strPeriodicity
    
    idxSibling = -1
    
    For i = 0 To m.Tree.Count
        Set Ind = Nothing
        If m.Tree.NodeLevel(i) > 0 Then
            Set Ind = m.Tree(i)
            If m.Tree.RelativeIndex(i, eTREE_Parent) = iParentIndex Then
                If Not Ind Is Nothing Then
                    If Ind.CodedName = strPivPoint Then
                        Set IndPoint = Ind
                        idxSibling = i
                        iFound = iFound + 1
                    ElseIf Ind.CodedName = strPivResist Then
                        Set IndResist = Ind
                        idxSibling = i
                        iFound = iFound + 1
                    ElseIf Ind.CodedName = strPivSupport Then
                        Set IndSupport = Ind
                        idxSibling = i
                        iFound = iFound + 1
                    ElseIf Ind.CodedName = strPivAvg Then
                        Set IndAvg = Ind
                        idxSibling = i
                        iFound = iFound + 1
                    End If
                End If
            End If
        End If
    Next
    
    ' see if all 4 were found
    If iFound = 4 Then
        'if all is on or all is off then toggle, else turn all on
        If IndPoint.Display = IndResist.Display And IndPoint.Display = IndSupport.Display And IndPoint.Display = IndAvg.Display Then
            IndPoint.Display = Not IndPoint.Display
            IndResist.Display = Not IndResist.Display
            IndSupport.Display = Not IndSupport.Display
            IndAvg.Display = Not IndAvg.Display
        Else
            IndPoint.Display = True
            IndResist.Display = True
            IndSupport.Display = True
            IndAvg.Display = True
        End If
    Else
        ' if not all were there, then turn off what was there
        If Not IndPoint Is Nothing Then IndPoint.Display = False
        If Not IndResist Is Nothing Then IndResist.Display = False
        If Not IndSupport Is Nothing Then IndSupport.Display = False
        If Not IndAvg Is Nothing Then IndAvg.Display = False
        
        ' and then just re-add the study
        strFile = "JP " & strPeriodicity
        If Not TemplateAddStudy(strFile) Then
            InfBox "Add study failed." & vbCrLf & "Please check that you have the file" & vbCrLf & strFile & ".STU", "I", , strFile & " Pivots"
        End If
    End If
    
#If 0 Then ' TLB: should be obsolete, but save for awhile
    ElseIf idxSibling > 0 Then
        'create & add missing indicator(s) to tree
        If IndResist Is Nothing Then
            Set IndResist = New cIndicator
            With IndResist
                .DisplayType = eINDIC_Rectangles
                .IndLabelMode = eINDIC_NoLabel
                .Style = eINDIC_Medium
                .Color = 255                'red
                .CodedName = strPivResist
                .Name = "JP Pivot Resistance " & strPeriodicity
                .Parm(2) = strPeriodicity
                .Parm(3) = 5
                .Parm(4) = 0
                .Display = True
                If .FunctionID = 0 Then
                    GoTo ErrExit            'something wrong - bailout
                Else
                    m.Tree.Add IndResist, , idxSibling, eTREE_NextSibling
                End If
            End With
        End If
    
        If IndPoint Is Nothing Then
            Set IndPoint = New cIndicator
            With IndPoint
                .DisplayType = eINDIC_Rectangles
                .IndLabelMode = eINDIC_NoLabel
                .Style = eINDIC_Medium
                .Color = 16711680           'blue
                .CodedName = strPivPoint
                .Name = "JP Pivot Point " & strPeriodicity
                .Parm(2) = strPeriodicity
                .Parm(3) = 0
                .Display = True
                If .FunctionID = 0 Then
                    GoTo ErrExit            'something wrong - bailout
                Else
                    m.Tree.Add IndPoint, , idxSibling, eTREE_NextSibling
                End If
            End With
        End If
    
        If IndSupport Is Nothing Then
            Set IndSupport = New cIndicator
            With IndSupport
                .DisplayType = eINDIC_Rectangles
                .IndLabelMode = eINDIC_NoLabel
                .Style = eINDIC_Medium
                .Color = 32768              'green
                .CodedName = strPivSupport
                .Name = "JP Pivot Support " & strPeriodicity
                .Parm(2) = strPeriodicity
                .Parm(3) = 5
                .Parm(4) = 0
                .Display = True
                If .FunctionID = 0 Then
                    GoTo ErrExit            'something wrong - bailout
                Else
                    m.Tree.Add IndSupport, , idxSibling, eTREE_NextSibling
                End If
            End With
        End If
    
        If IndAvg Is Nothing Then
            Set IndAvg = New cIndicator
            With IndAvg
                .DisplayType = eINDIC_Rectangles
                .IndLabelMode = eINDIC_NoLabel
                .Style = eINDIC_Medium
                .Color = 33023              'orange
                .CodedName = strPivAvg
                .Name = "JP Sentiment " & strPeriodicity
                .Parm(2) = strPeriodicity
                .Parm(3) = 3
                .Parm(4) = 0
                .Display = True
                If .FunctionID = 0 Then
                    GoTo ErrExit            'something wrong - bailout
                Else
                    m.Tree.Add IndAvg, , idxSibling, eTREE_NextSibling
                End If
            End With
        End If
        
        'turn all on
        IndPoint.Display = True
        IndResist.Display = True
        IndAvg.Display = True
        IndSupport.Display = True
    End If
#End If

ErrExit:
    Set Ind = Nothing
    Set IndResist = Nothing
    Set IndPoint = Nothing
    Set IndSupport = Nothing
    Set IndAvg = Nothing
    
    GenerateChart eRedo3_Settings
    
    Exit Sub

ErrSection:
    RaiseError "cChart.HandleJPButtons"

End Sub

Public Property Get SymbolPitType() As Long
On Error GoTo ErrExit:

    Dim ePitType As eFutureSymbolType
       
    If m.nSymbolPitType = -1 Then
        If Len(m.strSpreadSymbols) > 0 Then
            m.nSymbolPitType = ePrimarySymbol
        Else
            ePitType = GetSymbolPitType(Symbol)
            m.nSymbolPitType = ePitType
        End If
    End If
    
    SymbolPitType = m.nSymbolPitType

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.SymbolPitType.Get"

End Property

Public Property Get RemoveOvernightGap() As Boolean
    If AllowRemoveOvernightGap Then
        RemoveOvernightGap = m.bRemoveOvernightPriceGap
    End If
End Property

Public Property Let RemoveOvernightGap(ByVal bRemoveOvernightGap As Boolean)
    If AllowRemoveOvernightGap Then
        m.bRemoveOvernightPriceGap = bRemoveOvernightGap
    End If
End Property

Public Property Get CompositeIndTree() As cGdTree
On Error GoTo ErrSection:

    If m.CompositeIndicators Is Nothing Then
        Set m.CompositeIndicators = New cGdTree
    End If
    Set CompositeIndTree = m.CompositeIndicators

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cChart.CompositeIndTree.Get", eGDRaiseError_Raise

End Property

Public Function TemplateCopy(ByVal strDest$, ByVal strSource$, ByVal bUseDefaults As Boolean) As String
On Error GoTo ErrSection:

    Dim strFile$, strCopyFrom$
    
    If bUseDefaults Then
        strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strDest & ".CHT"
        strCopyFrom = ActiveChart.Chart.Template
    Else
        strFile = strDest
        strCopyFrom = strSource
    End If
    
    If Len(strCopyFrom) > 0 Then
        If bUseDefaults Then strCopyFrom = g.ChartGlobals.strCPCRoot & "\Charts\" & strCopyFrom & ".CHT"
        If Not FileExist(strCopyFrom) Then strCopyFrom = ""
        
        On Error GoTo ErrSection:
        If Len(strCopyFrom) = 0 Then
            ' else use the Default template
            strCopyFrom = g.ChartGlobals.strCPCRoot & "\Charts\Templates\Default.CHT"
        End If
        
        If FileExist(strCopyFrom) Then
            FileCopy strCopyFrom, strFile
            If FileExist(strFile) Then TemplateCopy = strFile
        Else
            GoTo ErrSection:
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cChart.TemplateCopy", eGDRaiseError_Raise

End Function

Public Property Get PfpIndicators() As String
    PfpIndicators = m.strPfpIndKeys
End Property

Public Property Let PfpIndicators(ByVal strKeys$)
    m.strPfpIndKeys = strKeys
End Property

' TLB: this may just be temporary (while developing)
Public Function GetSwingArray(Optional ByVal eType As eIndicatorDisplayType = eINDIC_ClusterPrice) As cGdArray

    Dim Ind As cIndicator

    If eType = eINDIC_ClusterPrice Then
        Set Ind = m.Tree(kClusterPriceKey)
        If Not Ind Is Nothing Then
            Set GetSwingArray = Ind.Data
        End If
    End If

End Function

'returns: -1 - error occured, clusters removed from tree
'          0 - new cluster indicators created & added to tree
'          1 - existing clusters validated
Public Function InitClusterInd(Optional ByVal bValidate As Boolean = False, _
    Optional ByVal dEndDate# = 0#, Optional ByVal nBars& = -1) As Long
On Error GoTo ErrSection:

    Dim i&, rc&, idx&, idxPricePane&
    
    Dim bNewPane As Boolean
    
    Dim Pane As cPane
    Dim IndTime As cIndicator
    Dim IndPrice As cIndicator
    Dim Annot As cAnnotation


    rc = 1      'assume existing validated
    idxPricePane = m.Tree.Index("PRICE PANE")
    
    If bValidate Then
        Set Pane = m.Tree(kClusterTimeKeyPane)
        Set IndTime = m.Tree(kClusterTimeKeyInd)
        Set IndPrice = m.Tree(kClusterPriceKey)
        
        If Not m.Annots Is Nothing Then Set Annot = m.Annots(kClusterZoneRect)
        
        If IndPrice Is Nothing Then
            'exit if neither price nor time cluster indicator exists, else continue on to create price cluster
            If IndTime Is Nothing Then
                If Not Annot Is Nothing Then RemoveFibClusters
                GoTo ErrExit
            End If
        ElseIf IndTime Is Nothing Then
            'price exists without time, turn it on so user has a way to access editor
            IndPrice.Display = True
            GoTo ErrExit
        ElseIf IndTime.Display And Pane.Display Then    'pane can get turned off without children indicators getting turned off
            'both exists, make sure Annot fill color matches heatmap setting, which can get out of sync during symbol change
            If Not Annot Is Nothing Then
                If IndPrice.trueRangeColor = 1 Or IndTime.trueRangeColor = 1 Then
                    i = Val(Annot.Prop("FillColor"))
                    If i > 0 Then Annot.Prop("FillColor") = i * -1      'make negative
                End If
            End If
            GoTo ErrExit
        Else
            'time cluster not visible, just turn on price cluster
            IndPrice.Display = True
            GoTo ErrExit
        End If
    End If
    
    rc = -1
    'get/create price cluster indicator
    Set IndPrice = m.Tree(kClusterPriceKey)
    If IndPrice Is Nothing Then
        Set IndPrice = New cIndicator
        IndPrice.ClusterIndDefaults Me, eINDIC_ClusterPrice, False, dEndDate, nBars

        If idxPricePane > 0 And idxPricePane <= m.Tree.Count Then
            idx = m.Tree.Add(IndPrice, kClusterPriceKey, idxPricePane, eTREE_LastChild)
            If idx > 0 And idx <= m.Tree.Count Then
                IndPrice.geIndId = idx
            End If
        End If
    ElseIf dEndDate > 0 And nBars > 0 Then
        IndPrice.ClusterEndDate = dEndDate      'user adjusted date using cluster zone rect
        IndPrice.ClusterBarsBack = nBars
        rc = m.Tree.Index(kClusterPriceKey)
        If IndPrice.geIndId <> rc Then IndPrice.geIndId = rc
        GoTo ErrExit
    End If
    If IndPrice Is Nothing Then GoTo ErrExit
    
    'get/create time cluster pane
    Set Pane = m.Tree(kClusterTimeKeyPane)
    If Pane Is Nothing Then
        If idxPricePane > 0 And idxPricePane <= m.Tree.Count Then
            Set Pane = New cPane
            idx = m.Tree.Add(Pane, kClusterTimeKeyPane, idxPricePane, eTREE_NextSibling)
            If idx > 0 And idx <= m.Tree.Count Then
                Pane.gePaneId = idx
                bNewPane = True
            End If
        End If
    End If
    If Pane Is Nothing Then GoTo ErrExit
    
    'get/create time cluster indicator
    Set IndTime = m.Tree(kClusterTimeKeyInd)
    If IndTime Is Nothing Then
        Set IndTime = New cIndicator
        IndTime.ClusterIndDefaults Me, eINDIC_ClusterTime, False, dEndDate, nBars
        idx = Pane.gePaneId
        If idx > 0 And idx <= m.Tree.Count Then
            'double, tripple, quadruple check that Pane object is valid
            If idx = m.Tree.Index(kClusterTimeKeyPane) Then
                If m.Tree(idx) Is Pane Then
                    Pane.Display = True
                    Pane.YscaleLabelAll = 2     'JM 07-28-2015 flag pane as time cluster pane
                    If Pane.gePaneId <> idx Then Pane.gePaneId = idx
                    idx = m.Tree.Add(IndTime, kClusterTimeKeyInd, idx, eTREE_FirstChild)
                End If
            End If
        End If
    End If
    If IndTime Is Nothing Then GoTo ErrExit
    
    If bValidate Then
        rc = 1
        'price cluster was added because time cluster exists
        IndPrice.Display = IndTime.Display
        Pane.Display = IndTime.Display
    Else
        rc = 0
        'new cluster added, set default parm values
        i = ValOfText(GetIniFileProperty("ClusterSwingParm2", "0", kClusterIniSection, g.strIniFile))
        If i <> 0 Then IndPrice.Parm(2) = i
        i = ValOfText(GetIniFileProperty("ClusterSwingParm3", "0", kClusterIniSection, g.strIniFile))
        If i <> 0 Then IndPrice.Parm(3) = i
        i = ValOfText(GetIniFileProperty("ClusterSwingParm4", "0", kClusterIniSection, g.strIniFile))
        If i <> 0 Then IndPrice.Parm(4) = i
    End If
    
    RedoMode = eRedo6_ReloadInd
    
ErrExit:
    If rc = 0 Then
        IndPrice.Display = True
        IndTime.Display = True
        Pane.Display = True
    ElseIf rc < 0 Then
        m.Tree.Remove kClusterPriceKey
        m.Tree.Remove kClusterTimeKeyInd
        If bNewPane Then m.Tree.Remove kClusterTimeKeyPane
    End If
    
    Set Pane = Nothing
    Set IndTime = Nothing
    Set IndPrice = Nothing
    
    InitClusterInd = rc
    
    Exit Function


ErrSection:
    InitClusterInd = -1
    
    Set Pane = Nothing
    Set IndTime = Nothing
    Set IndPrice = Nothing
    
    RaiseError "cChart.InitClusterInd"

End Function

Public Sub RemoveFibClusters()
On Error GoTo ErrSection:

    Dim i&
    Dim Annot As cAnnotation

    If Not m.Annots Is Nothing Then
        For i = m.Annots.Count To 1 Step -1
            Set Annot = m.Annots(i)
            If Not Annot Is Nothing Then
                If Annot.eUsage = eANNOT_FibClusters Then
                    Annot.geRemoveAnnotation m.geChartObj
                    m.Annots.Remove i
                End If
            End If
        Next
    End If
    
    If Not m.Tree Is Nothing Then
        m.Tree.Remove kClusterTimeKeyPane
        m.Tree.Remove kClusterPriceKey
        ResetSplitPane
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.RemoveFibClusters"

End Sub

Public Property Get CompSymType(Optional ByVal eLookForType As eCompSym = eCompSym_None) As eCompSym
On Error GoTo ErrSection:

    Dim i&, nPricePaneId&
    Dim Ind As cIndicator
    Dim Pane As cPane
    
    Dim bLookPercent As Boolean
    Dim bLookOthers As Boolean
    
    Dim eCompSymType As eCompSym

    If m.Tree Is Nothing Then Exit Property
    
    nPricePaneId = m.Tree.Index("PRICE PANE")
    
    
    Select Case eLookForType            '6542
        Case eCompSym_None
            bLookPercent = True
            bLookOthers = True
        Case eCompSym_PercentPane
            bLookPercent = True
        Case eCompSym_PricePane, eCompSym_NonPricePane
            bLookOthers = True
    End Select
    
    For i = 1 To Tree.Count
        If Tree.NodeLevel(i) = 0 Then
            Set Pane = Tree(i)
            If bLookPercent Then
                If Not Pane Is Nothing Then
                    If Pane.PaneLogFlag = ePANE_LogFlagPercent Then
                        eCompSymType = eCompSym_PercentPane
                        Exit For
                    End If
                End If
            End If
        ElseIf bLookOthers Then
            Set Ind = Tree(i)
            If Not Ind Is Nothing Then
                If Ind.isPriceInd = 0 Then
                    'an indicator is a comparison symbol if the data type is bardata and it is not the primary price bars
                    If Ind.MyKey <> "PRICE" Then
                        If Ind.DataType = eINDIC_BarData Then
                            If Tree.RelativeIndex(i, eTREE_Parent) = nPricePaneId Then
                                eCompSymType = eCompSym_PricePane
                            Else
                                eCompSymType = eCompSym_NonPricePane
                            End If
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    CompSymType = eCompSymType

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.CompSymType"

End Property

Public Property Get TypeOfChart() As eChartType
On Error GoTo ErrSection:
    
    Dim eType As eChartType
    Dim Pane As cPane
    
    If m.eTypeOfChart = eTypeChart_Undefined Then
        If Len(m.strSpreadSymbols) > 0 Then
            eType = eTypeChart_Spread
        Else
            Set Pane = Tree("PRICE PANE")
            
            If Pane Is Nothing Then
                eType = eTypeChart_Undefined
            ElseIf Pane.SeasonalPaneFlag = 1 Then
                eType = eTypeChart_Seasonal
            ElseIf CompSymType = eCompSym_None Then
                eType = eTypeChart_Standard
            Else
                eType = eTypeChart_PercentComp
            End If
        End If
        
        m.eTypeOfChart = eType
    End If
    
    TypeOfChart = m.eTypeOfChart
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.CompSymType"

End Property

Public Function SeasonalReset(ByRef aReturnInd As cGdArray) As Long
'returns 1)index of price pane if found
'        2)existing seasonal indicator objects if return array is passed in
On Error GoTo ErrSection:

    Dim i&, idx&
    
    Dim Pane As cPane
    Dim Ind As cIndicator
    Dim IndPrice As cIndicator

    idx = -1
    If Not aReturnInd Is Nothing Then aReturnInd.Size = 0
    m.eTypeOfChart = eTypeChart_Seasonal
    
    'remove all items from tree except price pane & price indicator
    With m.Tree
        Set Pane = .Item("PRICE PANE")
        Set IndPrice = .Item("PRICE")
        
        For i = .Count To 1 Step -1
            If TypeOf .Item(i) Is cPane Then
                If Not .Item(i) Is Pane Then .Remove i
            ElseIf TypeOf .Item(i) Is cIndicator Then
                Set Ind = .Item(i)
                
                If Not aReturnInd Is Nothing Then
                    If Ind.SeasonalIndFlag = 1 Then aReturnInd.Add Ind, 0
                End If
                
                If Not Ind Is IndPrice Then .Remove i
            End If
        Next
        
        If Not Pane Is Nothing Then
            idx = .Index("PRICE PANE")
            Pane.SplitPaneClear             '6363
            Pane.SeasonalPaneFlag = 1
            Pane.Desc = ""      'reset
        End If
        
        If Not IndPrice Is Nothing Then
            IndPrice.SeasonalIndFlag = 0
            IndPrice.Display = False
            IndPrice.DisplayType = eINDIC_Line
        End If
        
        MyBlankBars = 0         'reset
        ShowEmptyBars = False
        
    End With
    
    SeasonalReset = idx
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cChart.SeasonalReset"

End Function

Private Sub LoadSeasonalInd()
On Error GoTo ErrSection:

    Dim Pane As cPane
    
    Dim Ind As cIndicator
    Dim IndCurr As cIndicator
    Dim IndAvg As cIndicator
    Dim IndBull As cIndicator
    Dim IndBear As cIndicator
    
    Dim tbSeasonal As New cGdTable
    Dim aExist As New cGdArray          'holds any seasonal indicator currently in the chart
    
    Dim i&, j&, iColor&, iColorCounter&
    Dim iCurr&, iAvg&, iBull&, iBear&   'counters for restoring existing indicator properties
    Dim idx&, idxNew&, iRedrawSave&, iCount&
    
    Dim dValAtEnd#
    
    Dim iColorFrom&, iColorTo&
    Dim iBullTrendColor&, iBearTrendColor&
    
    Dim strStatus$, strName$
    Dim bShow As Boolean
    
    Dim bOverlayCycles As Boolean
    Dim bOverlayTrends As Boolean
    
    Dim fg As VSFlexGrid

    If g.bUnloading Then Exit Sub
    
    If TypeOfChart <> eTypeChart_Seasonal Then Exit Sub
    
    m.bDisableRT = True

    If m.Form Is Nothing Then Exit Sub
    If m.nSymbolID = 0 Then Exit Sub
    If m.Bars.Size > 0 Then Exit Sub
    
    'This call originally made in the load event of the form. If a form is reused, as when
    'reverting or switching chart pages, the load event of the form is not triggered. Makine
    'this call here ensures that the dynamic controls get loaded if not already done.
    InitSeasonalControls m.Form
    
    idx = SeasonalReset(aExist)
    If idx <= 0 Then Exit Sub           'something went wrong, there is no pane object in tree
    
    iCurr = -1
    iAvg = -1
    iBull = -1
    iBear = -1
    
    iCount = aExist.Size
    If iCount > 0 Then
        For i = 0 To aExist.Size - 1
            Set Ind = aExist(i)
            If Not Ind Is Nothing Then
                Select Case Ind.MyKey
                    Case kSeasonalCurrIndKey
                        iCurr = i
                    Case kSeasonalAvgIndKey
                        iAvg = i
                        bOverlayTrends = Ind.Overlayed
                    Case kSeasonalBullIndKey
                        iBull = i
                    Case kSeasonalBearIndKey
                        iBear = i
                    Case Else
                        bShow = Ind.Display
                        bOverlayCycles = Ind.Overlayed
                End Select
            End If
        Next
    Else
        bShow = True
    End If
    
    'pass m.Bars along with Table into function
    strStatus = CalcSeasonalChart(tbSeasonal, m.Bars, Symbol, GetPeriodStr(m.nPeriodicity), SeasonalCycle, , m.dFromDate)
    If Len(strStatus) > 0 Then
        Set Pane = m.Tree("PRICE PANE")
        If Not Pane Is Nothing Then Pane.Desc = "No data from seasonal calculations. " & strStatus
        
        Set tbSeasonal = Nothing
        Set Pane = Nothing
        m.Bars.Size = 0
        
        Exit Sub
    End If
    
    'change chart symbol to what was returned from function in bars object
    m.nSymbolID = m.Bars.Prop(eBARS_SymbolID)

    'create indicator average/bull/bear trends
    For i = 0 To 2
        strName = tbSeasonal.FieldName(i)
        
        Set Ind = New cIndicator
        With Ind
            .DisplayType = eINDIC_Line
            .IndLabelMode = eINDIC_Nothing
            .Overlayed = False
            .DataType = eINDIC_Array
            .Name = strName
            .Data.CopyFromHandle tbSeasonal.FieldArrayHandle(i)
            .SeasonalIndFlag = 1
        
            If strName = "Average Trend" Then
                m.Tree.Key(idxNew) = kSeasonalAvgIndKey
                .MyKey = kSeasonalAvgIndKey
                If iAvg = -1 Then
                    .Style = eINDIC_MediumThick         'seasonal defaults
                    .Color = vbBlack
                    .UpColor = 8421376
                    .DownColor = 16777152
                    .Display = True
                Else
                    .Style = aExist(iAvg).Style
                    .Color = aExist(iAvg).Color
                    .UpColor = aExist(iAvg).UpColor
                    .DownColor = aExist(iAvg).DownColor
                    .Display = aExist(iAvg).Display
                End If
                iColorFrom = Ind.UpColor
                iColorTo = Ind.DownColor
                .Overlayed = bOverlayTrends
                Set IndAvg = Ind
            ElseIf strName = "Bullish Trend" Then
                m.Tree.Key(idxNew) = kSeasonalBullIndKey
                .MyKey = kSeasonalBullIndKey
                If iBull = -1 Then
                    .Style = eINDIC_MediumThick         'default
                    .Color = 32768
                    .Display = True
                Else
                    .Style = aExist(iBull).Style
                    .Color = aExist(iBull).Color
                    .Display = aExist(iBull).Display
                End If
                iBullTrendColor = .Color
                .Overlayed = bOverlayTrends
                Set IndBull = Ind
            ElseIf strName = "Bearish Trend" Then
                m.Tree.Key(idxNew) = kSeasonalBearIndKey
                .MyKey = kSeasonalBearIndKey
                If iBear = -1 Then
                    .Style = eINDIC_MediumThick         'default
                    .Color = vbRed
                    .Display = True
                Else
                    .Style = aExist(iBear).Style
                    .Color = aExist(iBear).Color
                    .Display = aExist(iBear).Display
                End If
                iBearTrendColor = .Color
                .Overlayed = bOverlayTrends
                Set IndBear = Ind
            End If
        End With
    Next
    
    'create indicator for current cycle
    i = tbSeasonal.NumFields - 1
    Set IndCurr = New cIndicator
    With IndCurr
        .MyKey = kSeasonalCurrIndKey
        .DisplayType = eINDIC_Line
        .IndLabelMode = eINDIC_Nothing
        .Overlayed = bOverlayCycles Or bOverlayTrends
        .DataType = eINDIC_Array
        .Name = tbSeasonal.FieldName(i)
        .Data.CopyFromHandle tbSeasonal.FieldArrayHandle(i)
        .SeasonalIndFlag = 1
        If iCurr = -1 Then
            .Style = eINDIC_Medium                      'default
            .Color = vbBlue
            .Display = True
        Else
            .Style = aExist(iCurr).Style
            .Color = aExist(iCurr).Color
            .Display = aExist(iCurr).Display
        End If
    End With
    
    iCount = tbSeasonal.NumFields - 4
    iColorCounter = 1
    For i = 3 To tbSeasonal.NumFields - 2
        strName = tbSeasonal.FieldName(i)
        
        Set Ind = New cIndicator
        
        iColor = GradientColor(100 * iColorCounter / iCount, iColorTo, iColorFrom)      '6360
        
        With Ind
            .Display = bShow
            .DisplayType = eINDIC_Line
            .Style = eINDIC_Default
            .IndLabelMode = eINDIC_Nothing
            .Color = iColor
            .Overlayed = bOverlayCycles
        
            .DataType = eINDIC_Array
            .geIndpaneId = idx
            .Name = strName
        
            .Data.CopyFromHandle tbSeasonal.FieldArrayHandle(i)
            .SeasonalIndFlag = 1
        
            'set bull or bear color for text in grid
            For j = .Data.Size - 1 To 0 Step -1
                If .Data(j) <> kNullData Then
                    dValAtEnd = .Data(j)
                    Exit For
                End If
            Next
        
            If dValAtEnd > 0 Then
                .UpColor = iBullTrendColor
            ElseIf dValAtEnd < 0 Then
                .UpColor = iBearTrendColor
            Else
                .UpColor = vbBlack
            End If
        End With
        
        ' put the indicator into m.Tree
        idxNew = m.Tree.Add(Ind, "", idx, eTREE_LastChild)
        Ind.geIndId = idxNew
        
        iColorCounter = iColorCounter + 1
    Next
    
    'add current cycle and trend indicators to tree
    If Not IndCurr Is Nothing Then
        idxNew = m.Tree.Add(IndCurr, "", idx, eTREE_LastChild)
        m.Tree.Key(idxNew) = IndCurr.MyKey
        IndCurr.geIndId = idxNew
    End If
    
    If Not IndBear Is Nothing Then
        idxNew = m.Tree.Add(IndBear, "", idx, eTREE_LastChild)
        m.Tree.Key(idxNew) = IndBear.MyKey
        IndBear.geIndId = idxNew
    End If
    
    If Not IndBull Is Nothing Then
        idxNew = m.Tree.Add(IndBull, "", idx, eTREE_LastChild)
        m.Tree.Key(idxNew) = IndBull.MyKey
        IndBull.geIndId = idxNew
    End If
    
    If Not IndAvg Is Nothing Then
        idxNew = m.Tree.Add(IndAvg, "", idx, eTREE_LastChild)
        m.Tree.Key(idxNew) = IndAvg.MyKey
        IndAvg.geIndId = idxNew
    End If
    
    Set fg = m.Form.fgChartFlex(eFlexGridIdx_Seasonal)
    fg.Rows = fg.FixedRows
    iRedrawSave = fg.Redraw
    fg.Redraw = flexRDNone
    
    For i = m.Tree.Count To 1 Step -1
        If TypeOf m.Tree(i) Is cIndicator Then
            Set Ind = m.Tree(i)
            With fg
                If UCase(Ind.Name) = "PRICE" Then
                    'do nothing
                Else
                    .Rows = .Rows + 1
                    If Ind.MyKey = kSeasonalCurrIndKey Or Ind.MyKey = kSeasonalAvgIndKey Or _
                       Ind.MyKey = kSeasonalBullIndKey Or Ind.MyKey = kSeasonalBearIndKey Then
                    
                        .TextMatrix(fg.Rows - 1, 0) = Ind.Name
                        .Cell(flexcpFontBold, fg.Rows - 1, 0) = True
                        .Cell(flexcpForeColor, fg.Rows - 1, 0) = Ind.Color
                    Else
                        .TextMatrix(fg.Rows - 1, 0) = Ind.Name
                        .Cell(flexcpForeColor, fg.Rows - 1, 0) = Ind.UpColor
                    End If
                    
                    'index of indicator
                    .TextMatrix(fg.Rows - 1, 1) = Str(Ind.geIndId)
                End If
            End With
        End If
    Next
    fg.Redraw = iRedrawSave
    
    'add dashed horz line at 0 (do after populating grid so will not show up in grid)
    Set Ind = New cIndicator
    With Ind
        .Display = True
        .DisplayType = eINDIC_Line
        .Style = eINDIC_Dot
        .IndLabelMode = eINDIC_Nothing
        .Color = 4539717
        .DataType = eINDIC_Constant
        .geIndpaneId = idx
        .Data(0) = 0
    End With
    idxNew = m.Tree.Add(Ind, "", idx, eTREE_LastChild)
    m.Tree.Key(idxNew) = Ind.MyKey
    Ind.geIndId = idxNew
    
    If Not g.bStarting Then
        If m.Bars.Size > 0 Then
            m.geChartStruct.x_dataPointCount = m.Bars.Size
            
            If Int(gdGetNum(m.geChartStruct.glhChartRect, 3) / m.Bars.Size) = RoundNum(gdGetNum(m.geChartStruct.glhChartRect, 3) / m.Bars.Size) Then
                m.geChartStruct.x_dataPointPix = RoundNum(gdGetNum(m.geChartStruct.glhChartRect, 3) / m.Bars.Size)
            Else
                m.geChartStruct.x_dataPointPix = RoundNum(gdGetNum(m.geChartStruct.glhChartRect, 3) / m.Bars.Size) - 1
            End If
            PixelsPerBar = m.geChartStruct.x_dataPointPix
            geForceRecalc
        End If
    End If

ErrExit:
    Set Pane = Nothing
    Set Ind = Nothing
    
    Set tbSeasonal = Nothing
    Set aExist = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "cChart.LoadSeasonalInd"
    Resume ErrExit

End Sub

Public Property Get SeasonalCycle() As String
On Error GoTo ErrSection:
    
    SeasonalCycle = SeasonalCycleLen & " " & SeasonalCycleTypeStr

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.SeasonalCycle.Get"

End Property

Public Property Let SeasonalCycle(ByVal strCycle$)
On Error GoTo ErrSection:

    Dim strCycleLen$, strCycleType$, i&
    Dim eTypeSave As eSeasonalCycleType

    strCycleLen = Trim(Parse(strCycle, " ", 1))
    i = Len(strCycleLen)
    If i <= 0 Then Exit Property
    
    strCycleType = Trim(Right(strCycle, Len(strCycle) - i))
    If Len(strCycleType) < 0 Then Exit Property
    
    i = Int(Val(strCycleLen))
    eTypeSave = SeasonalCycleTypeEnum
    
    SeasonalCycleType = strCycleType
    If m.eCycleType = eCycleType_Undefined Then
        SeasonalCycleType = eTypeSave
    Else
        SeasonalCycleLen = i
        m.strSeasonalCycle = SeasonalCycle
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.SeasonalCycle.Let"

End Property

Public Property Get SeasonalCycleLen() As Long
    SeasonalCycleLen = m.nCycleLen
End Property

Public Property Let SeasonalCycleLen(ByVal nLen&)
    m.nCycleLen = nLen
End Property

Public Property Get SeasonalCycleTypeStr() As String
On Error GoTo ErrSection:

    Dim strCycle$
    
    strCycle = "Unknown"

    Select Case m.eCycleType
        Case eCycleType_Year
            strCycle = "Year"
        Case eCycleType_Quarter
            strCycle = "Quarter"
        Case eCycleType_Month
            strCycle = "Month"
        Case eCycleType_Week
            strCycle = "Week"
        Case eCycleType_Days
            strCycle = "Days"
        Case eCycleType_FullMoons
            strCycle = "Full Moons"
        Case eCycleType_NewMoons
            strCycle = "New Moons"
    End Select
    
    SeasonalCycleTypeStr = strCycle

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.SeasonalCycleString.Get"

End Property

Public Property Get SeasonalCycleTypeEnum() As eSeasonalCycleType
On Error GoTo ErrSection:

    SeasonalCycleTypeEnum = m.eCycleType

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.SeasonalCycleTypeEnum.Get"

End Property

Public Property Let SeasonalCycleType(ByVal vCycleType As Variant)
On Error GoTo ErrSection:

    Dim i&
    
    If VarType(vCycleType) = vbString Then
        Select Case Trim(UCase(vCycleType))
            Case "YEAR"
                m.eCycleType = eCycleType_Year
            Case "QUARTER"
                m.eCycleType = eCycleType_Quarter
            Case "MONTH"
                m.eCycleType = eCycleType_Month
            Case "WEEK"
                m.eCycleType = eCycleType_Week
            Case "DAYS"
                m.eCycleType = eCycleType_Days
            Case "FULL MOONS"
                m.eCycleType = eCycleType_FullMoons
            Case "NEW MOONS"
                m.eCycleType = eCycleType_NewMoons
        End Select
    Else
        i = CLng(vCycleType)
        If i >= eCycleType_Year And i <= eCycleType_NewMoons Then m.eCycleType = i
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.SeasonalCycleType.Let"

End Property

Public Function ValidSeaonalCycles(aTypes As cGdArray) As Long
On Error GoTo ErrSection:

    Set aTypes = New cGdArray
    
    aTypes.Add "Year"
    aTypes.Add "Quarter"
    aTypes.Add "Month"
    aTypes.Add "Week"
    aTypes.Add "Days"
    aTypes.Add "Full Moons"
    aTypes.Add "New Moons"
    
    ValidSeaonalCycles = SeasonalCycleTypeEnum
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cChart.ValidSeaonalCycles"

End Function

Public Function ValidSeasonalBarTypes(aTypes As cGdArray) As Long
On Error GoTo ErrSection:

    Dim nIdxOfCurrType&
    Dim strCurrentType As String
    
    Set aTypes = New cGdArray
    
    aTypes.Add "Daily"
    aTypes.Add "Weekly"
    aTypes.Add "Monthly"
    
    If Not m.Bars Is Nothing Then
        strCurrentType = m.Bars.Prop(eBARS_PeriodicityStr)
        If Len(strCurrentType) > 0 Then
            If Parse(strCurrentType, " ", 2) <> "" Then
                strCurrentType = Parse(strCurrentType, " ", 2)
            End If
        End If
        
        Select Case UCase(strCurrentType)
            Case "DAILY"
                ValidSeasonalBarTypes = 0
            Case "WEEKLY"
                ValidSeasonalBarTypes = 1
            Case "MONTHLY"
                ValidSeasonalBarTypes = 2
            Case Else
                ValidSeasonalBarTypes = -1
        End Select
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cChart.ValidSeasonalBarTypes"

End Function

Public Function SeasonalHasOverlay() As Boolean
On Error GoTo ErrSection:

    Dim i&
    Dim Ind As cIndicator
    Dim bOverlay As Boolean

    If Not m.Tree Is Nothing Then
        For i = 1 To m.Tree.Count
            If TypeOf m.Tree(i) Is cIndicator Then
                Set Ind = m.Tree(i)
                
                If Ind.MyKey = kSeasonalAvgIndKey Or Ind.MyKey = kSeasonalCurrIndKey Or _
                   Ind.MyKey = kSeasonalBullIndKey Or Ind.MyKey = kSeasonalBearIndKey Or _
                   Ind.MyKey = "PRICE" Then
                   
                    'do nothing
                
                ElseIf Ind.Overlayed Then
                    bOverlay = True
                    Exit For
                End If
            End If
        Next
    End If
    
    SeasonalHasOverlay = bOverlay

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cChart.SeasonalHasOverlay"

End Function

Public Sub SeasonalIndClear()
On Error GoTo ErrSection:

    Dim Pane As cPane
    
    If TypeOfChart = eTypeChart_Seasonal Then
        Set Pane = m.Tree("PRICE PANE")
        If Not Pane Is Nothing Then
            If Pane.SeasonalPaneFlag <> 0 Then
                m.nLastGoodDataBarX = 0
                m.Bars.Size = 0
                Pane.Desc = "Click Apply Changes when done."
                geDrawChart
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.SeasonalIndClear"

End Sub

Public Property Get LastEditedAnnot() As cAnnotation
On Error GoTo ErrSection:

    If m.LastEditedAnnot Is Nothing And Not m.LastCreatedAnnot Is Nothing Then
        Set LastEditedAnnot = m.LastCreatedAnnot
    Else
        Set LastEditedAnnot = m.LastEditedAnnot
    End If

ErrExit:
    Exit Property

ErrSection:
    Set LastEditedAnnot = Nothing
    RaiseError "cChart.LastEditedAnnot.Get"

End Property

Public Sub LastEditCreate(Annot As cAnnotation, ByVal bJustCreated As Boolean)
On Error GoTo ErrSection:

    If Annot Is Nothing Then
        Set m.LastEditedAnnot = Nothing
        Set m.LastCreatedAnnot = Nothing
    ElseIf bJustCreated Then
        Set m.LastEditedAnnot = Nothing
        Set m.LastCreatedAnnot = Annot
    Else
        Set m.LastEditedAnnot = Annot.MakeCopy
        Set m.LastCreatedAnnot = Nothing
    End If
    
    SyncChartingBtns m.Form, "", "", "", "", "", "", "", "", False, True

ErrExit:
    Exit Sub

ErrSection:
    Set m.LastEditedAnnot = Nothing
    RaiseError "cChart.LastEditedAnnot.Let"

End Sub

Public Sub LastEditedAnnotUndo()
On Error GoTo ErrSection:

    Dim strKey$, i&
    Dim bMultiChart As Boolean
    Dim bFound As Boolean
    
    Dim AnnotExist As cAnnotation
    
    Dim eType As eAnnotType
    
    If m.LastEditedAnnot Is Nothing Then
        If Not m.LastCreatedAnnot Is Nothing Then RemoveAnnots True, , , False
    Else
        strKey = m.LastEditedAnnot.Prop("AnnotKey")
        If Len(strKey) <> 0 Then Set AnnotExist = Annots(strKey)
        
        If Not AnnotExist Is Nothing Then
            If AnnotExist.eUsage = eANNOT_PriceAlert Then
                'do nothing
            Else
                'remove existing annot since this is one that has properties user wants to undo
                bMultiChart = AnnotExist.MultiChartFlag
                AnnotExist.geRemoveAnnotation m.geChartObj
                Annots.Remove (strKey)
                If bMultiChart Then SyncGlobalAnnots Nothing, bMultiChart
            End If
        End If
        
        If m.LastEditedAnnot.eUsage = eANNOT_PriceAlert Then
            'price alerts annots are not tied to chart objects, need to update or restore the alert they are tied to
            If Not m.LastEditedAnnot.AlertObject Is Nothing Then                '6458
                For i = 1 To g.Alerts.Count
                    'if alert exists, just update it
                    If g.Alerts(i).AlertKey = m.LastEditedAnnot.AlertObject.AlertKey Then
                        g.Alerts(i).FromFileString m.LastEditedAnnot.AlertObject.ToFileString
                        g.Alerts(i).UpdateChartObject False
                        bFound = True
                        Exit For
                    End If
                Next
                If Not bFound Then
                    'if alert not found then add it to alert collection & let alert handle updating the charts
                    g.Alerts.Add m.LastEditedAnnot.AlertObject, m.LastEditedAnnot.AlertObject.AlertKey
                    If FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid
                End If
            End If
        Else
            i = m.Annots.Add(m.LastEditedAnnot, strKey)
            
            eType = m.LastEditedAnnot.eType
            
            If eType = eANNOT_DNExpansion Or eType = eANNOT_DNRetracement _
                Or eType = eANNOT_DNExpansion2 Or eType = eANNOT_DNExpansion3 _
                Or eType = eANNOT_DNExpansion4 Or eType = eANNOT_Gartley _
                Or eType = eANNOT_RegressionLine Or eType = eANNOT_RegressionLine _
                Or eType = eANNOT_DollarLine Or eType = eANNOT_DollarLine2 _
                Or eType = eANNOT_DollarLine3 Or eType = eANNOT_DollarLine4 _
                Or eType = eANNOT_AndrewFork Or eType = eANNOT_Pattern _
                Or eType = eANNOT_TriangleWedge Or eType = eANNOT_ChannelHighlight _
                Or eType = eANNOT_ElliotTimeRatio Or eType = eANNOT_FibABCD Then

                    m.LastEditedAnnot.geMoveFlag = 1

            ElseIf eType = eANNOT_TextEdit Then
                m.LastEditedAnnot.geMoveFlag = 2                  '4352
            End If
            
            bMultiChart = m.LastEditedAnnot.MultiChartFlag
            If bMultiChart Then SyncGlobalAnnots Nothing, bMultiChart
        End If
    End If
    
    GenerateChart eRedo1_Scrolled
    
    Set m.LastEditedAnnot = Nothing
    Set m.LastCreatedAnnot = Nothing
    
    SyncChartingBtns m.Form, "", "", "", "", "", "", "", "", False, True

ErrExit:
    Exit Sub

ErrSection:
    Set m.LastEditedAnnot = Nothing
    RaiseError "cChart.LastEditedAnnotUndo"

End Sub

Public Property Get PivotVwapIndicator(ByVal nInterval&) As cIndicator
On Error GoTo ErrExit:

    Dim i&
    Dim Ind As cIndicator

    If m.PivotVwapInd Is Nothing Then
        
        For i = 1 To m.Tree.Count
            If TypeOf m.Tree(i) Is cIndicator Then
                Set Ind = m.Tree(i)
                If InStr(UCase(Ind.CodedText), "VWAP") <> 0 And InStr(UCase(Ind.CodedText), "PIVOT") <> 0 Then
                    'make sure pivot have same #of days, weeks etc. as requesting indicator
                    If Val(Ind.Parm(2)) = nInterval Then
                        Set m.PivotVwapInd = Ind
                        Exit For
                    End If
                End If
            End If
        Next
    
    End If

    Set PivotVwapIndicator = m.PivotVwapInd

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.PivotVwapIndicator"

End Property

Public Sub geReleaseChart()
    
    On Error Resume Next
    
    If m.geChartObj <> 0 Then
        geResetPanes
        geCloseChart m.geChartObj
        m.geChartObj = 0
    End If

End Sub

Public Sub RemoveAllTickBuffers()
On Error GoTo ErrSection:

    Dim i&

    If Not g.RealTime Is Nothing Then
        g.RealTime.RemoveTickBuffer m.Bars
        
        For i = 1 To UBound(m.OtherBars)
            g.RealTime.RemoveTickBuffer m.OtherBars(i)
        Next
        
        If Len(m.strSpreadSymbols) > 0 Then
            For i = 0 To UBound(m.SpreadBars)
                g.RealTime.RemoveTickBuffer m.SpreadBars(i)
            Next
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.RemoveAllTickBuffers"

End Sub

Public Sub AddAllTickBuffers()

    Dim i&

    If Not g.RealTime Is Nothing Then
        g.RealTime.AddTickBuffer m.Bars
               
        For i = 1 To UBound(m.OtherBars)
            g.RealTime.AddTickBuffer m.OtherBars(i)
        Next
        
        If Len(m.strSpreadSymbols) > 0 Then
            For i = 0 To UBound(m.SpreadBars)
                g.RealTime.AddTickBuffer m.SpreadBars(i)
            Next
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.AddAllTickBuffers"

End Sub

Public Sub TemplateLoadGlobalAnnots()
On Error GoTo ErrSection:

    Dim fh%, i&, idx&
    
    Dim strFile$, strName$
    Dim strSection$, strType$
    
    Dim Annot As cAnnotation
    Dim IndClusterPrice As cIndicator
    Dim AnnotClusterZone As cAnnotation

'JM: 09-22-2011
'This routine added to fix aardvark 6447. When a chart page gets cached, the cached chart object
'does not updated with newly added or edited annotations that are set to shown on all charts.
'Since this routine is meant to be called from a valid cached chart object, the annotation
'and pane & indicator trees are expected to exist. Do not call RemoveAnnots because cached
'chart objects do not have valid grapheng.dll structures.

    If m.Annots Is Nothing Then Exit Sub
    If m.Tree Is Nothing Then Exit Sub
    
    If m.nSymbolID = 0 And Len(m.strExtData & m.strSpreadSymbols) = 0 Then Exit Sub
    
    ' first remove existing global annots
    For i = Annots.Count To 1 Step -1
        If m.Annots(i).MultiChartFlag = True Then m.Annots.Remove i
    Next
    
    'look for price cluster indicator
    If Not m.Tree Is Nothing Then Set IndClusterPrice = m.Tree(kClusterPriceKey)

    ' annotations for all templates (Price Pane annotations)
    strName = "^" & AnnotSymbol ' Trim(Str(m.nSymbolID))
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strName & ".ANO"
    If FileExist(strFile) Then
        fh = FreeFile
        Open strFile For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strType
            If Left(strType, 1) = "[" Then
                strSection = strType
                If Left(strSection, 6) = "[ANNOT" Then
                    Set Annot = New cAnnotation
                    Annot.TemplateLoad fh, Me
                    Annot.MultiChartFlag = True
                    If Annot.IsValid Then
                        i = m.Annots.Add(Annot)
                    End If
                End If
            End If
        Loop
        Close #fh
    End If

    If Not IndClusterPrice Is Nothing Then
        'annot files are often deleted manually by techs or developers for troubleshooting purposes
        'users may be sharing a .CHT file for fib clusters but are not aware that the zone is in an .ANO file
        If AnnotClusterZone Is Nothing Then
            i = m.Tree.Index("PRICE PANE")
            
            'annot for cluster zone does not exist, create it
            Set Annot = New cAnnotation
            
            Annot.CreateNew Me, eANNOT_Rectangle, i, 0, 0, IndClusterPrice.ClusterEndDate, 0
            
            Annot.eUsage = eANNOT_FibClusters
            Annot.Prop("Shape") = 0
            Annot.PreIndicator = 1
            Annot.Prop("FillPattern") = 1
            Annot.Prop("FillColor") = Annot.ClusterPropDefault("ZoneFillColor")
            Annot.Color = Annot.ClusterPropDefault("ZoneColor")
            Annot.Prop("AnnotKey") = kClusterZoneRect
            
            idx = m.Annots.Add(Annot, kClusterZoneRect)
            
            'set to true so the dates for annotation will get set when price cluster gets calculated
            IndClusterPrice.Display = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChart.TemplateLoadGlobalAnnots"

End Sub

Public Sub LoadProfileHistory()
On Error GoTo ErrSection

    Static dPrevDate#           '6994
    
    Dim i&, j&
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If m.Bars Is Nothing Then Exit Sub
    
    If Not m.bProfileChart Then
        ProfileVolHistory Me       'profile volume indicator
        Exit Sub
    End If
    
    If m.BarsProfile Is Nothing Then Exit Sub       'something very wrong here!!!
    If m.BarsProfile.Size < 1 Then Exit Sub
    
    i = m.Bars.SessionDate(0)
    j = m.Bars.SessionDateForTime(m.BarsProfile(eBARS_DateTime, 0), False)
    If j - i > 10 Then
        i = j - 10
    End If
    
    Set m.BarsProfile = ProfileBarsGet(m.Bars.Prop(eBARS_SymbolID), i, j, False)

    If m.Bars.SessionDateForTime(m.BarsProfile(eBARS_DateTime, 0), False) <= m.Bars.SessionDate(0) Or _
       dPrevDate = m.BarsProfile(eBARS_DateTime, 0) Then    '6994 if no change then probably no more data to load
        m.Form.tmrProfileLoad.Enabled = False
        
        If Not m.Bars.IsIntraday Then
            i = m.Bars.Prop(eBARS_PeriodsPerBar)
            If PixelsPerBar < 4 * i Then
                PixelsPerBar = 4 * i
            End If
        End If
        
        GenerateChart eRedo1_Scrolled
    End If
    
    dPrevDate = m.BarsProfile(eBARS_DateTime, 0)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.LoadProfileHistory"

End Sub

Public Property Get IsProfileChart() As Boolean
    IsProfileChart = m.bProfileChart
End Property

Public Property Let IsProfileChart(ByVal bIsProfile As Boolean)
On Error GoTo ErrSection:

    Dim i&
    
    Dim Ind As cIndicator
    Dim bPaneShow As Boolean

    m.bProfileChart = False
    'set flag based on chart objects, not what is passed in
    If Not m.Tree Is Nothing Then
        For i = 1 To m.Tree.Count
            If TypeOf m.Tree(i) Is cPane Then
                bPaneShow = m.Tree(i).Display
            ElseIf TypeOf m.Tree(i) Is cIndicator Then
                Set Ind = m.Tree(i)
                If Ind.Display Then
                    If Ind.DataType = eINDIC_ProfileBars And Ind.DisplayType = eINDIC_Profile Then
                        m.bProfileChart = True
                        Exit For
                    End If
                End If
            End If
        Next
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cChart.IsProfileChart.Let"

End Property

Private Sub AdjustProfileScrollBar()
On Error GoTo ErrSection:

    Dim iX&, i&
    
    i = m.aXdate.Size
    iX = gdGetSize(m.profileChartInfo.gdhXDateCopy)
    If i = iX Then Exit Sub
    
    Dim dDateStop#, dDate#
    Dim hXdateArray&, hXbarArray
    
    dDateStop = gdGetNum(m.geChartStruct.gdhDate, 0) - 1       'go one day back to get full profile for overnight symbols
    hXdateArray = m.aXdate.ArrayHandle      'for efficiency
    hXbarArray = m.aXBar.ArrayHandle
    
    iX = iX - 1
    For i = m.aXdate.Size - 1 To 0 Step -1
        
        If gdGetNum(hXdateArray, i) <= dDateStop Or iX < 0 Then Exit For
        
        dDate = gdGetNum(m.profileChartInfo.gdhXDateCopy, iX)
        Do While gdGetNum(hXdateArray, i) <> dDate
            gdDeleteItems hXdateArray, i, 1
            gdDeleteItems hXbarArray, i, 1
            i = i - 1
            If gdGetNum(hXdateArray, i) <= dDateStop Then Exit Do
        Loop
        iX = iX - 1
    
    Next
    
    m.bRedrawTwice = True
    
    m.nLastGoodDataBarX = -1
    For iX = m.aXBar.Size - 1 To 0 Step -1
        i = m.aXBar(iX)
        If i >= 0 Then
            If m.Bars(eBARS_Close, i) <> kNullData Then
                m.nLastGoodDataBarX = iX
                Exit For
            End If
        End If
    Next

    dDate = gdGetNum(m.geChartStruct.gdhDate, gdGetSize(m.geChartStruct.gdhDate) - 1)
    If dDate = m.aXdate(m.aXdate.Size - 1) Then
        ResetLastScreenDate
    End If
    
    ' reset horizontal scroll bar
    With m.hsb
        .Max = m.aXBar.Size - 1
        If m.dLastScreenDate > 0 Then
            m.aXdate.BinarySearch m.dLastScreenDate, iX
            If iX > .Max Then                                       'Or iX >= m.nLastGoodDataBarX Then
                If RedoMode >= eRedo9_ReloadData Then
                    iX = m.nLastGoodDataBarX + m.nBlankBars
                    If iX > .Max Then iX = .Max
                Else
                    iX = .Max
                End If
            ElseIf iX < .Min Then
                iX = .Min
            End If
        Else
            iX = m.nLastGoodDataBarX + m.nBlankBars
            If iX > .Max Then
                iX = .Max
            ElseIf Not m.Form Is Nothing Then
                If Not m.Form.GameMode Is Nothing Then
                    iX = .Max       '5118
                End If
            End If
        End If
        .Value = iX                                                 '.Max ' set scrollbar to end
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.AdjustProfileScrollBar"

End Sub

Private Sub ProfileBarsClear()
On Error GoTo ErrSection:

    Dim dLastDate#, dLastDownload#

    If Not m.BarsProfile Is Nothing Then
    
        If Not m.Bars.IsIntraday And Not g.RealTime.Active Then         '6850
            dLastDownload = LastDailyDownload
            dLastDate = m.BarsProfile(eBARS_DateTime, m.BarsProfile.Size - 1)
            
            While Int(dLastDate) > dLastDownload
                m.BarsProfile.DeleteSomeBars m.BarsProfile.Size - 1, 1
                dLastDate = m.BarsProfile(eBARS_DateTime, m.BarsProfile.Size - 1)
            Wend
            
        End If
        
        Set m.BarsProfileTicks = Nothing
        ProfileBarsFree m.BarsProfile
        Set m.BarsProfile = Nothing
        gdSetSize m.profileChartInfo.gdhXDateCopy, 0, 0
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.ProfileBarsClear"

End Sub

Public Sub ShowHideAnnotsByType(ByVal typeToShowHide As eAnnotType, ByVal nHide&, ByVal bSetProperty As Boolean)
On Error GoTo ErrSection:
    
    Dim i&
    Dim Annot As cAnnotation
    
    If Annots.Count = 0 Then Exit Sub
    
    For i = Annots.Count To 1 Step -1
        Set Annot = Annots(i)
        If Not Annot Is Nothing Then
            If typeToShowHide = eANNOT_IndicLabel Then
                If Annot.eType = eANNOT_Icon And Annot.eUsage = eANNOT_IndicatorLabel Then
                    Annot.Hide nHide, Me, bSetProperty
                End If
            ElseIf Annot.eType = typeToShowHide Then
                Annot.Hide nHide, Me, bSetProperty
            End If
        End If
    Next
    
    Set Annot = Nothing
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.ShowHideAnnotsByType"

End Sub

'Return: 0 if specific type found and not hidden
'        1 if specific type found and is hidden
'       -1 if specific type not found
Public Function HasHiddenAnnots(ByVal specificType As eAnnotType) As Long
On Error GoTo ErrSection:

    Dim i&, Found&
    Dim Annot As cAnnotation
    
    If specificType = eANNOT_UndefinedType Then
        Found = 0
    Else
        Found = -1
    End If
    
    For i = 1 To m.Annots.Count
        If Not m.Annots(i) Is Nothing Then
            Set Annot = m.Annots(i)
            If Annot.eType = eANNOT_UndefinedType Then
                If Annot.Prop("IsHidden") = 1 Then
                    Found = 1
                    Exit For
                End If
            ElseIf Annot.eType = specificType Then
                'special case for EWI labels
                If Annot.eType = eANNOT_ElliotLabel And Not Annot.IsEndUserEWI Then
                    Found = Val(Annot.Prop("IsHidden"))
                    Exit For
                End If
            End If
        End If
    Next
    
    HasHiddenAnnots = Found

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cChart.HasHiddenAnnots"

End Function

Public Sub TestFZ()
On Error GoTo ErrSection:

    Dim i&, h&, idx&, nDate&, nBar&, nStopTicks&, nStartDate&
    Dim d#, dtEntry#, dtExit#, dEntry#, dExit#, dStopPrice#, dProfit#, dCumProfit#, dMonthlyProfit#, dTickValue#
    Dim s$, strSymbol$, strPos$, strSig$, strPrevSig$, strStop$
    Dim Indicator As cIndicator, Pane As cPane
    Dim Buys As cGdArray, Sells As cGdArray, Filter As cGdArray
    Dim aTrades As New cGdArray
    Static nEndDate&, dStopMult#

Exit Sub
    If Not g.FractZen.Allowed Then Exit Sub
    If m.bToEndOfData Then Exit Sub
    DoEvents
    
    If nEndDate < 1000 Then
        nEndDate = DateSerial(2014, 5, 30)
    End If
    i = DateOf(InfBox("Ending date to test through ...", "?", , "FractZen Backtest", , , , , , "d", DateFormat(nEndDate)))
    If i <= 0 Then Exit Sub
    nEndDate = i
    
    If dStopMult = 0 Then
        dStopMult = 2
    End If
    d = Val(InfBox("Stop Multiplier (if < 10),| or # of Ticks (if > 10) ...", "?", , "FractZen Backtest", , , , , , "n", Str(dStopMult)))
    If d = 0 Then Exit Sub
    dStopMult = Abs(d)
    
       
    ' header line
    s = "Position" & vbTab & "Symbol" & vbTab & "Entry Time (ET)" & vbTab & "Entry Price" & vbTab & "Exit Time (ET)" & vbTab & "Exit Price" & vbTab & "Profit" & vbTab & "Cum Profit" & vbTab & "Monthly" & vbTab & "#Ticks Stop"
    aTrades.Add s
       
    strSymbol = m.Bars.Prop(eBARS_Symbol)
    If IsForex(strSymbol) Then
        ' e.g. for $CAD-CHF the base symbol is $USD-CHF
        dTickValue = g.Profit.Profit(strSymbol, m.Bars.Prop(eBARS_TickMove), 100000, nEndDate + 1)
    Else
        dTickValue = m.Bars.Prop(eBARS_TickValue)
    End If
    
    dEntry = kNullData
    strPos = "F"
    nStartDate = Int(m.dToDate)
    For nDate = nStartDate To nEndDate
        ' load data through this date and generate the Buys/Sells signals
        Set Buys = Nothing
        Set Sells = Nothing
        Set Filter = Nothing
        If Month(nDate) <> Month(nDate - 1) Then
            dMonthlyProfit = 0
        End If
        If IsWeekday(nDate) Then
            'i = CalcAutoBreakoutRange(strSymbol, nDate)
            's = Str(i) & "b"
            m.dToDate = nDate
            GenerateChart eRedo9_ReloadData
            DoEvents
            Do While IsPartiallyLoaded
                GenerateChart eRedo5_RecalcInd
                DoEvents
            Loop
            nBar = LastGoodDataBar(False, False)
            If m.Bars.SessionDate(nBar) = nDate Then
                ' get the Buys/Sells signals
                For idx = 1 To m.Tree.Count
                    If m.Tree.NodeLevel(idx) = 0 Then
                        Set Pane = m.Tree(idx)
                    ElseIf Pane.Display Then
                        Set Indicator = m.Tree(idx)
                        With Indicator
                            If .Display Then 'And .DataType = eINDIC_BooleanArray Then
                                s = UCase(.ChartLabel)
                                If s = "FZ_BUYS" Then
                                    Set Buys = .Data
                                ElseIf s = "FZ_SELLS" Then
                                    Set Sells = .Data
                                ElseIf s = "FZ_FILTER" Then
                                    Set Filter = .Data
                                End If
                            End If
                        End With
                    End If
                Next
                Set Indicator = Nothing
                If Buys Is Nothing Then
                    InfBox "The 'FZ_Buys' highlight bar does not exist.", "e", , "ERROR"
                    Exit For
                ElseIf Sells Is Nothing Then
                    InfBox "The 'FZ_Sells' highlight bar does not exist.", "e", , "ERROR"
                    Exit For
                End If
            End If
        End If
        
        If Not Buys Is Nothing And Not Sells Is Nothing Then
            ' go back to first bar of this trading session
            For nBar = m.Bars.Size - 1 To 1 Step -1
                i = m.Bars.SessionDate(nBar - 1)
                If i > 0 And i < nDate Then
                    Exit For
                End If
            Next
            
            ' for each bar in this session
            For nBar = nBar To m.Bars.Size - 1
                If m.Bars.SessionDate(nBar) <> nDate Or m.Bars(eBARS_Close, nBar) = kNullData Then
                    Exit For
                End If
                                
                ' get current signal for this bar
                dExit = kNullData
                strSig = ""
If Not Filter Is Nothing Then
    strPrevSig = ""
End If
                If Abs(Buys(nBar)) = 1 Then
                    ' ignore this one if same direction as previous signal
                    If strPrevSig <> "Buy" Then
                        strSig = "Buy"
                        strPrevSig = strSig
                    End If
                ElseIf Abs(Sells(nBar)) = 1 Then
                    ' ignore this one if same direction as previous signal
                    If strPrevSig <> "Sell" Then
                        strSig = "Sell"
                        strPrevSig = strSig
                    End If
                End If
                ' ignore if already in this position
                If strSig = "Buy" And strPos = "L" Then
                    strSig = ""
                ElseIf strSig = "Sell" And strPos = "S" Then
                    strSig = ""
                End If
                
                ' get info for EXIT from old position (which will become the new entry)
                ' -- but if no entry signal, then see if hit the stop-loss on this bar
                If Len(strSig) > 0 Then
                    dExit = m.Bars(eBARS_Open, nBar)
                ElseIf strPos = "L" Then
                    If m.Bars(eBARS_Open, nBar) < dStopPrice Then
                        dExit = m.Bars(eBARS_Open, nBar)
                    ElseIf m.Bars(eBARS_Low, nBar) < dStopPrice Then
                        dExit = dStopPrice
                    End If
                ElseIf strPos = "S" Then
                    If m.Bars(eBARS_Open, nBar) > dStopPrice Then
                        dExit = m.Bars(eBARS_Open, nBar)
                    ElseIf m.Bars(eBARS_High, nBar) > dStopPrice Then
                        dExit = dStopPrice
                    End If
                End If
                If dExit <> kNullData Then
                    ' get time of exit (use NY time for everything)
                    If dExit <> m.Bars(eBARS_Open, nBar) Then
                        ' if price isn't the Open, then just use the date of this bar
                        dtExit = m.Bars(eBARS_DateTime, nBar)
                    ElseIf m.Bars.SessionDate(nBar - 1) = nDate Then
                        ' else if same session, add 1 second to date-time of previous bar
                        dtExit = m.Bars(eBARS_DateTime, nBar - 1) + 1 / 86400#
                    Else
                        ' else if first bar of session, then use start of session
                        dtExit = nDate + m.Bars.Prop(eBARS_StartTime) / 1440#
                        If m.Bars.Prop(eBARS_StartTime) > m.Bars.Prop(eBARS_EndTime) Then
                            dtExit = dtExit - 1
                        End If
                    End If
                    dtExit = ConvertTimeZone(dtExit, m.Bars.Prop(eBARS_ExchangeTimeZoneInf), "NY")
                    
                    If dEntry <> kNullData And strPos <> "F" Then
                        ' format info for this Trade (entry and exit)
                        ' Pos, Symbol, EntryTime, EntryPrice, ExitTime, ExitPrice, Profit, CumProfit, #StopTicks
                        strStop = Str(nStopTicks) & " tick Stop"
                        If Len(strSig) = 0 Then
                            strStop = strStop & " HIT"
                        End If
                        dProfit = dTickValue * (dExit - dEntry) / m.Bars.Prop(eBARS_TickMove)
                        dProfit = Round(dProfit, 4)
                        If strPos = "S" Then
                            dProfit = -dProfit
                        End If
                        dCumProfit = dCumProfit + dProfit
                        dMonthlyProfit = dMonthlyProfit + dProfit
                        If strPos = "L" Then
                            strPos = "Long"
                        ElseIf strPos = "S" Then
                            strPos = "Short"
                        End If
                        s = "YYYY-MM-DD hh:mm:ss"
                        s = strPos & vbTab & strSymbol & vbTab & Format(dtEntry, s) & vbTab & m.Bars.PriceDisplay(dEntry) _
                            & vbTab & Format(dtExit, s) & vbTab & m.Bars.PriceDisplay(dExit) & vbTab _
                            & Format(dProfit, "$#0.00") & vbTab & Format(dCumProfit, "$#0.00") & vbTab _
                            & Format(dMonthlyProfit, "$#0.00") & vbTab & strStop
                        aTrades.Add s
                    End If
                    strPos = "F"
                End If
                
                ' check filter to see if this signal will be a new entry or just an exit
                If Not Filter Is Nothing Then
                    If Filter(nBar - 1) = kNullData Then
                        InfBox "Not enough bars at " & DateFormat(nDate), "e", , "Error"
                        nDate = nEndDate + 1
                        Exit For
                    ElseIf strSig = "Buy" Then
                        If Filter(nBar - 1) <= 0 Then
                            strSig = ""
                        End If
                    ElseIf strSig = "Sell" Then
                        If Filter(nBar - 1) >= 0 Then
                            strSig = ""
                        End If
                    End If
                End If
                
                ' now get info for the new Entry, and the new Stop-price
                If Len(strSig) > 0 Then
                    ' get #StopTicks based on tick range from the previous weekend
                    For i = nDate To 0 Step -1
                        If Not IsWeekday(i) Then
                            If dStopMult > 10 Then
                                nStopTicks = Round(dStopMult)
                            Else
                                'nStopTicks = Round(dStopMult * CalcAutoBreakoutRange(strSymbol, i - 1, Val(Parse(m.strAutoBreakout, vbTab, 1)), Val(Parse(m.strAutoBreakout, vbTab, 2))))
                                nStopTicks = Round(dStopMult * g.FractZen.GetFractZenRange(strSymbol, i))
                            End If
                            Exit For
                        End If
                    Next
                    
                    dEntry = dExit
                    dtEntry = dtExit
                    If strSig = "Buy" Then
                        strPos = "L"
                        dStopPrice = dEntry - m.Bars.Prop(eBARS_TickMove) * nStopTicks
                    ElseIf strSig = "Sell" Then
                        strPos = "S"
                        dStopPrice = dEntry + m.Bars.Prop(eBARS_TickMove) * nStopTicks
                    End If
                End If
            Next
        End If
    Next
    
    m.dToDate = nStartDate
    GenerateChart eRedo9_ReloadData
    
    MakeDir App.Path & "\Trades\FZ\"
    s = App.Path & "\Trades\FZ\" & strSymbol & ".txt"
    aTrades.ToFile s
    DoEvents
    RunProcess s
    'RunProcess Chr(34) & s & Chr(34)
    'InfBox "Done", "i", , strSymbol

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cChart.TestFZ"

End Sub

' TLB: this is used primarily by the EWI analysts who have a published chart page with a mixture
' of daily and intraday charts, but where the daily and intraday charts are allowed to be seen
' by clients with either the daily and/or the intraday enablement codes respectively.
Public Sub SetRequired()

    Dim s$
    On Error Resume Next
    s = "'Enablement Code' required to see this chart ..."
    m.strRequired = UCase(Trim(InfBox(s, "i", "+OK", "Required Enablement", , , , , , "s", m.strRequired)))

End Sub

Public Property Get DrawCommands() As cGdArray
    Set DrawCommands = m.DrawingCommands
End Property

Public Property Get Unpublishable() As Boolean
    Unpublishable = m.bUnpublishable
End Property

Public Property Let Unpublishable(ByVal bUnpublishable As Boolean)
    m.bUnpublishable = bUnpublishable
End Property

Private Function HasProfileVolume() As cIndicator
On Error GoTo ErrSection:

    Dim i&
    
    If m.Tree Is Nothing Then Exit Function
    
    For i = 1 To Tree.Count
        If Not Tree(i) Is Nothing Then
            If Tree.NodeLevel(i) > 0 Then
                If Tree(i).DataType = eINDIC_ProfileVolume Then
                    Set HasProfileVolume = Tree(i)
                    Exit For
                End If
            End If
        End If
    Next

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cChart.HasVolumeProfile"

End Function
