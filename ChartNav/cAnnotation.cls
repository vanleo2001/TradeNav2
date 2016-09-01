VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAnnotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Special Bug Note:
'   file dated 12/15/2005 4:18pm released with v2.11 with alert issues 3191/3204 fixed!

Private Const kXfudge = 0.00001

Private Const kFibRatio = 1# / 1.61803398875 ' inverse of "Golden Ratio"

'misc constants for fib clusters
Private Const kClusterFibFillColor = 13750737           'light gray RGB(209,209,209)

'preset ratios for all fibs that don't have special defaults below
Private Const kFibRatios = "-1,0.25,0.382,0.5,0.618,0.75,1.618,2,2.618"
Private Const kFibShow = "0,0,1,1,1,0,0,0,0"
 
'preset ratios for fib time ratio
Private Const kFibTimeRatios = "0.25,0.382,0.5,0.618,0.75,1,1.618,2,2.618"
Private Const kFibTimeShow = "0,1,1,1,0,0,0,0,0"

'preset ratios for andrews fork parallel lines & time interval markers
Private Const kParallelRatios = "0.25,0.382,0.5,0.618,1.618,2.0"
Private Const kParallelShow = "0,1,1,1,0,0"
Private Const kParallelShowNone = "0,0,0,0,0,0"
Private Const kParallelColor = "255,255,16711680,255,255,255"
Private Const kIntervalRatios = "0.25,0.382,0.5,0.618,1,1.618,2.0"
Private Const kIntervalShow = "0,1,1,1,0,0,0"
Private Const kIntervalShowNone = "0,0,0,0,0,0,0"

'preset ratios for fib expansion, fib fan, adv risk reward
Private Const kFibExpFanRatios = "0.382,0.5,0.618,1.618,2,2.618"
Private Const kFibExpShow = "0,0,0,1,1,1"
Private Const kFibFanShow = "1,1,1,0,0,0"
Private Const kFibBold = "0,0,0,0,0,0"
Private Const kAdvRiskRewardRatios = "-1,-2,0.382,0.5,0.618,2,3"
Private Const kAdvRiskRewardShow = "0,0,0,0,0,1,1"

'preset ratios for Daniel Code Retrnacement
Private Const kDanCodeFib = "0.89,0.742,0.593,0.5,0.445,0.375,0.297"
Private Const kDanCodeFibShow = "0,1,1,1,1,1,0"
Private Const kDanCodeFibColor = "0,255,255,128,255,255,255"

'preset ratio for fib extension (this is variation of DnExp)    -JM 12-17-2009: new defaults for non-dinapoli version
Private Const kFibDnExRatios = "0.382,0.5,0.618,1,1.618,2,2.618"
Private Const kFibDnExShow = "1,0,1,1,1,0,0"
Private Const kFibDnExColors = "8388736,0,16711680,65280,255,0,0"

'preset ratios for speed resistance fan
Private Const kSpResistRatios = "0.33,0.67"

'preset ratios for balloon strangle
Private Const kBalloonRatios = "0.5,1,1.5,2,2.5,3,3.5,4,4.5,5,-0.5,-1,-1.5,-2,-2.5,-3,-3.5,-4,-4.5,-5"
Private Const kBalloonShow = "0,1,0,1,0,0,0,0,0,0,0,1,0,1,0,0,0,0,0,0"

'preset ratios for GreenBlatt version of fib time ratios tool
'JM: 5/22/2008
'   - the ratio/square root options moved to time zone tool
'   - to see original code look at sourcesafe history version 560 of this file
'   - companion editor is sourcesafe history version 271 of frmEditAnnot

'fixed bars count for GreenBlatt time zones
Private Const kZoneRatio = "14,23,38,61,78,127,161,261,423,685"
Private Const kZoneSqRoot = "38,48,61,78,88,112,127,161,205,261"

'fixed bars count for Daniel Code Zone
Private Const kZoneDanCode = "44,59,62,70"

'default colors for Pivot tool (same as pivot study)
Private Const kPivotColor = 16711680
Private Const kS1Color = 32768
Private Const kS2Color = 49344
Private Const kS3Color = 16711680
Private Const kR1Color = 192
Private Const kR2Color = 49344
Private Const kR3Color = 16711680

Private Enum eGANNacciSwing1
    eGannci_P1 = 1          ' [1]    price 1
    eGannci_P2              ' [2]    price 2
    eGannci_R               ' [3]    range (P1 - P2)
    eGannci_TB              ' [4]    trade bars
    eGannci_TB_P1           ' [5]    dif TB - P1
    eGannci_TB_P2           ' [6]    dif TB - P2
    eGannci_TB_R            ' [7]    dif TB - r
    eGannci_EQ_TB_P1        ' [8]    equals TB=P1
    eGannci_EQ_TB_P2        ' [9]    equals TB=P2
    eGannci_EQ_TB_R         '[10]    equals TB=F
    eGannci_CD              '[11]    calendar days
    eGannci_CD_P1           '[12]    dif CD - P1
    eGannci_CD_P2           '[13]    dif CD - P1
    eGannci_CD_R            '[14]    dif CD - R
    eGannci_EQ_CD_P1        '[15]    equals CD=P1
    eGannci_EQ_CD_P2        '[16]    equals CD=P2
    eGannci_EQ_CD_R         '[17]    equals CD=R
End Enum

Private Enum eGANNacciSwing2
    eGannci_R1 = 1          ' [1]    range swing 1 (PR1 - PR2)
    eGannci_R2              ' [2]    range swing 2 (PR2 - PR3)
    eGannci_TB1             ' [3]    trade bars swing 1
    eGannci_TB2             ' [4]    trade bars swing 2
    eGannci_TB1_R2          ' [5]    dif TB1 - R2
    eGannci_TB2_R1          ' [6]    dif TB2 - R1
    eGannci_EQ_TB1_R2       ' [7]    equals TB=P1
    eGannci_EQ_TB2_R1       ' [8]    equals TB=P2
    eGannci_CD1             ' [9]    calendar days swing 1
    eGannci_CD2             '[10]    calendar days swing 2
    eGannci_CD1_R2          '[11]    dif CD1 - R2
    eGannci_CD2_R1          '[12]    dif CD2 - R1
    eGannci_EQ_CD1_R2       '[13]    equals CD=P1
    eGannci_EQ_CD2_R1       '[14]    equals CD=P2
End Enum


Public Enum eImageSize
    eImgSmall = 8
    eImgMedium = 16
    eImgLarge = 24
End Enum

Public Enum eImageStyle
    eImgHollow = 0
    eImgFilled
End Enum

Public Enum eImageLabelAlign
    eImgLblAlign_Left = 3
    eImgLblAlign_Right = 4
    eImgLblAlign_Center = 5
    eImgLblAlign_Auto = 9           'this essentially defaults to left
End Enum

'NOTE: CNI stands for ChartNav Image
'Caution: These numbers match the enumerated type in grapheng.dll
Public Enum eStockImage
    eCNI_Arrow = 0
    eCNI_Plus = 1
    eCNI_Cross = 2
    eCNI_Circle = 3
    eCNI_Square = 4
    eCNI_Diamond = 5
    eCNI_Triangle = 6
    eCNI_LegLine = 7    'for chart legends only
    eCNI_LegRect = 8    'for chart legends only
    eCNI_SnowFlake = 9  'use only for marking PriceInd in addition to one other indicator
    eCNI_Arc = 10
    eCNI_Bell = 11
    eCNI_Bell_Gray = 12
    eCNI_Bell_Price = 13
    eCNI_Bell_Price_Gray = 14
    eCNI_PutSell = 15   'sell put icon for options navigator
    eCNI_Ascii = 29     'do not change - this matches max index in gdSelectIcon control
End Enum

'NOTE: Directions apply to arrow & trade annotations (trade annotations are E/W triangles)
'Caution: These numbers match enumerated type in grapheng.dll
Public Enum eImageDir
    eCNI_North = 0
    eCNI_NorthEast = 1
    eCNI_East = 2
    eCNI_SouthEast = 3
    eCNI_South = 4
    eCNI_SouthWest = 5
    eCNI_West = 6
    eCNI_NorthWest = 7
End Enum

'NOTE: cannot change the #'s assigned to
'each annotation type, since they are stored
'in template files by this number
Public Enum eAnnotType
    eANNOT_UndefinedType = 0
    eANNOT_IndicLabel = 1
    'eANNOT_Misc = 2
    eANNOT_Icon = 3
    eANNOT_SimpleLine = 4
    eANNOT_Trendline = 5
    eANNOT_DollarLine = 6
    eANNOT_HorzLine = 8
    eANNOT_VertLine = 9
    eANNOT_TargetShooter = 10
    eANNOT_Fibonacci = 11
    eANNOT_Rectangle = 12
    eANNOT_RegressionLine = 13
    eANNOT_DNExpansion = 14
    eANNOT_DNRetracement = 15
    eANNOT_TextEdit = 16
    eANNOT_Ellipse = 17
    eANNOT_GannLines = 18
    eANNOT_FibTimeRatio = 19
    eANNOT_FibTimeZones = 20
    eANNOT_FibArcs = 21
    eANNOT_TimeCycle = 22
    eANNOT_AndrewFork = 23
    eANNOT_SRLine = 24              'support/resistance line
    eANNOT_MultiRects = 25          'rects for highlight bars
    eANNOT_FibExpansion = 26
    eANNOT_FibFan = 27
    eANNOT_SpResistFan = 28         'speed resistance fan/lines
    eANNOT_Mirror = 29
    eANNOT_Pattern = 30
    eANNOT_RiskReward = 31
    eANNOT_TriangleWedge = 32
    eANNOT_ChannelHighlight = 33
    eANNOT_WaveLabels = 34
    eANNOT_BellAlert = 35
    eANNOT_ElliotLabel = 36         'image size must be 999999
    eANNOT_ElliotTimeRatio = 37
    eANNOT_Bracket = 38
    eANNOT_Pivot = 39
    eANNOT_FibABCD = 40
    eANNOT_Gartley = 41
    eANNOT_ArrowLine = 42
    eANNOT_TrendChannel = 43
    eANNOT_ElliotEwave = 45
    eANNOT_DanCodeFib = 46
    eANNOT_DanCodeZone = 47
    eANNOT_GannacciSwingSquare = 48
    eANNOT_GannacciCycle = 49
    eANNOT_GannacciTime = 50
    eANNOT_GannacciSwing1 = 51
    eANNOT_GannacciSwing2 = 52
    eANNOT_Trendline2 = 53
    eANNOT_Trendline3 = 54
    eANNOT_Trendline4 = 55
    eANNOT_TextEdit2 = 56
    eANNOT_TextEdit3 = 57
    eANNOT_TextEdit4 = 58
    eANNOT_BalloonStrangle = 59
    eANNOT_Fibonacci2 = 60
    eANNOT_Fibonacci3 = 61
    eANNOT_Fibonacci4 = 62
    eANNOT_DNExpansion2 = 63
    eANNOT_DNExpansion3 = 64
    eANNOT_DNExpansion4 = 65
    eANNOT_HorzLine2 = 66
    eANNOT_HorzLine3 = 67
    eANNOT_HorzLine4 = 68
    eANNOT_DollarLine2 = 69
    eANNOT_DollarLine3 = 70
    eANNOT_DollarLine4 = 71
    eANNOT_AdvRiskReward = 72
    eANNOT_SRLine2 = 73
    eANNOT_SRLine3 = 74
    eANNOT_SRLine4 = 75
End Enum

Public Enum eAnnotUsage
    eANNOT_UndefinedUsage = 0
    eANNOT_UserAdded = 1
    eANNOT_IndicatorLabel = 2
    eANNOT_Indicator = 3
    eANNOT_Trades = 4
    eANNOT_DateLabels = 5
    eANNOT_Notation = 6
    eANNOT_ProfitLines = 7
    eANNOT_AutoSwingTrend = 8
    eANNOT_SideWinderLabel = 9
    eANNOT_AvgEntryLine = 10
    eANNOT_WhatIf = 11
    eANNOT_BidAskChart = 12
    eANNOT_PriceAlert = 13
    eANNOT_OptionInfo = 14
    eANNOT_PatternProfit = 15
    eANNOT_FibClusters = 16
End Enum

Public Enum eAnnotPen
    eANNOT_Default = 0
    eANNOT_Thin = 1
    eANNOT_Medium = 2
    eANNOT_Thick = 3
    eANNOT_DashLg = 4
    eANNOT_DashSm = 5
    eANNOT_DashDot = 6
End Enum

Public Enum eGBZoneInUse
    eANNOT_GB_ZoneLucas = 0
    eANNOT_GB_ZoneFib = 1
    eANNOT_GB_Zone144 = 2
    eANNOT_GB_ZoneRatio = 3
    eANNOT_GB_ZoneSqRoot = 4
    eANNOT_GB_ZoneDanCode = 5
End Enum

'Note: Must keep this enum in sync with grapheng.dll
Public Enum eForkMode
    eANNOT_ForkModeUndefined = -1
    eANNOT_ForkNormal = 0
    eANNOT_Schiff = 1             'schiff
    eANNOT_SchiffMod = 2          'modified schiff
    eANNOT_SchiffModSchiff = 3    'modified schiff & schiff
    eANNOT_NormalSchiff = 4       'normal & schiff
    eANNOT_NormalSchiffMod = 5    'normal & modified schiff
    eANNOT_ForkAll = 6
End Enum

'JM 07-22-2013: for custom font sizes on EWIGMP pallette per Glen/Rob email
'   Note: Must keep this enum in sync with frmElliot.frm text & radio button control arrays
'         these enums are used as index values into the control arrays
Public Enum eCommonStyle
    eCommon_None = -1
    eCommon_Miniscule = 0
    eCommon_Submicro = 1
    eCommon_Micro = 2
    eCommon_Subminuette = 3
    eCommon_Minuette = 4
    eCommon_Minute = 5
    eCommon_Minor = 6
    eCommon_Intermediate = 7
    eCommon_Primary = 8
    eCommon_Cycle = 9
    eCommon_Supercycle = 10
    eCommon_GrandSupercycle = 11
    eCommon_CustomText = 12
End Enum

Private Type mPrivate
    Chart As cChart
    ' annotation type and usage
    eType As eAnnotType
    eUsage As eAnnotUsage
        
    ' common properties
    strPane As String
    dDate1 As Double
    dDate2 As Double
    Y1 As Double
    Y2 As Double
    nColor As Long
    nStyle As Long          'This is penstyle except for icontext type
    nPreIndicator As Long   'Flag: 1=draw before (i.e. behind) indicators, 0=draw after (i.e. in front of) indicators
    bMultiChart As Boolean
    strText As String
    
    'information to help troubleshoot switching between bar period
    dCreateDate As Double       'date time annotation was created
    strBarPeriod As String      'bar period of chart that annotation was created on
    strPeriodsPerBar As Long    'periods per bar of chart that annotation was created on
        
    ' collection of other annotation properties
    ' (can vary for different annotation types)
    ' note: do not use date or valu as property names (see TemplateLoad)
    Properties As cGdArray '(string array kept sorted)
    
    ' arrays of doubles to store additional date and Y values
    ' for annotations needing more than 2 date-Y value pairs
    ' e.g. - dinapoli retracement and expansion annotations
    aDates As cGdArray
    aYs As cGdArray
    
    'table & index for options information
    tbOptionsInfo As cGdTable
    aOptionsIdx As cGdArray
    
    'arrays of longs & doubles to hold:
    '   a) show flags, fib ratios & colors for fib tools
    '   b) show flags & calculated values for gannacci swing1 & swing 2 tools
    aRatioShow As cGdArray
    aRatios As cGdArray
    aRatioColor As cGdArray
    aRatioBold As cGdArray          'bold lines for fib supp/resist, fib expansion & dancode fib only
    
    ' used by SetPoints and MovePoint
    dReactPtBoundary As Double  'use by DnRetracement to identify reaction points left boundary
    X1 As Double
    X2 As Double
    dMoveOffSetX As Double
    dMoveOffSetY As Double
    dMoveOffSetLogY As Double
    nShiftKey As Long           '1=ShiftKey down, 2=CtrlKey down (passed in from chart form)
    
    nHide As Long
    nHitItemIdx As Long
    bMoveByMenu As Boolean
    bAddByMenu As Boolean
    bCanContinueLabel As Boolean
    bY1IsZeroRatio  As Boolean  'flag to determine which Y value is zero ratio for fib annots
    
    nGraphEngDll As Long        'Flag to determine whether to convert pen styles
    geTextAlign As Long         'currently only being used for ICO_TEXT annotations
    geChartObj As Long
    geAnnStruct As chart_annotation
    geAdded As Boolean
    
    oAlert As cAlert
    bAlertAdded As Boolean
    strAlert As String
    
    bGreenBlatt As Boolean
    bFirstOption As Boolean     'this is the very first, i.e. left-most, option annotation
    bOriginalVertical As Boolean
    
    bSchiffFork As Boolean
    bGannacciDebug As Boolean

    ' set True only if global Log drawing mode is set and this Annot is in a semi-log pane
    bUsingLogDrawingMode As Boolean
End Type
Dim m As mPrivate

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.eType = eANNOT_UndefinedType
    m.eUsage = eANNOT_UserAdded
    
    Set m.Properties = New cGdArray
    m.Properties.Create eGDARRAY_Strings
    
    Set m.aDates = New cGdArray
    m.aDates.Create eGDARRAY_Doubles
    
    Set m.aYs = New cGdArray
    m.aYs.Create eGDARRAY_Doubles
    
    Set m.aRatioShow = New cGdArray
    m.aRatioShow.Create eGDARRAY_Longs
    
    Set m.aRatioColor = New cGdArray
    m.aRatioColor.Create eGDARRAY_Longs
    
    Set m.aRatios = New cGdArray
    m.aRatios.Create eGDARRAY_Doubles
    
    Set m.aRatioBold = New cGdArray
    m.aRatioBold.Create eGDARRAY_Longs
        
    m.dReactPtBoundary = -1#
    m.nHitItemIdx = -1
    m.bGreenBlatt = HasModule("JGREEN")
    
    m.bGannacciDebug = FileExist("GannacciDebug.flg")
    
    m.geAnnStruct.glhAnnType = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhPenColor = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhPenStyle = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhPenSize = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhUnderline = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhFillColor = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhFillPattern = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhStyle = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhAlign = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhJustify = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhDirection = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhSize = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhImageType = gdCreateArray(eGDARRAY_Longs)
    
    m.geAnnStruct.glhPixTop = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhPixLeft = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhPixBottom = gdCreateArray(eGDARRAY_Longs)
    m.geAnnStruct.glhPixRight = gdCreateArray(eGDARRAY_Longs)
    
    m.geAnnStruct.gdhTop = gdCreateArray(eGDARRAY_Doubles)
    m.geAnnStruct.gdhLeft = gdCreateArray(eGDARRAY_Doubles)
    m.geAnnStruct.gdhBottom = gdCreateArray(eGDARRAY_Doubles)
    m.geAnnStruct.gdhRight = gdCreateArray(eGDARRAY_Doubles)
    m.geAnnStruct.gdhMisc = gdCreateArray(eGDARRAY_Doubles)
    
    m.geAnnStruct.gshText = gdCreateArray(eGDARRAY_Strings)
    m.geAnnStruct.gshFont = gdCreateArray(eGDARRAY_Strings)

    geResetData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.Intialize", eGDRaiseError_Raise

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Dim rc&
    
    Set m.Properties = Nothing
    Set m.aDates = Nothing
    Set m.aYs = Nothing
    Set m.aRatioShow = Nothing
    Set m.aRatios = Nothing
    Set m.aRatioColor = Nothing
    Set m.aRatioBold = Nothing
    Set m.Chart = Nothing
    Set m.oAlert = Nothing
    Set m.tbOptionsInfo = Nothing
    Set m.aOptionsIdx = Nothing
    
   
    'TODO: process non-success return code for removal of item?
    If m.geAdded Then rc = geRemoveItem(m.geChartObj, 4, m.geAnnStruct)
    
    gdDestroyArray m.geAnnStruct.glhAnnType
    gdDestroyArray m.geAnnStruct.glhPenColor
    gdDestroyArray m.geAnnStruct.glhPenStyle
    gdDestroyArray m.geAnnStruct.glhPenSize
    gdDestroyArray m.geAnnStruct.glhUnderline
    gdDestroyArray m.geAnnStruct.glhFillColor
    gdDestroyArray m.geAnnStruct.glhFillPattern
    gdDestroyArray m.geAnnStruct.glhStyle
    gdDestroyArray m.geAnnStruct.glhAlign
    gdDestroyArray m.geAnnStruct.glhJustify
    gdDestroyArray m.geAnnStruct.glhDirection
    gdDestroyArray m.geAnnStruct.glhSize
    gdDestroyArray m.geAnnStruct.glhImageType
    
    gdDestroyArray m.geAnnStruct.glhPixTop
    gdDestroyArray m.geAnnStruct.glhPixLeft
    gdDestroyArray m.geAnnStruct.glhPixBottom
    gdDestroyArray m.geAnnStruct.glhPixRight
    
    gdDestroyArray m.geAnnStruct.gdhTop
    gdDestroyArray m.geAnnStruct.gdhLeft
    gdDestroyArray m.geAnnStruct.gdhBottom
    gdDestroyArray m.geAnnStruct.gdhRight
    
    gdDestroyArray m.geAnnStruct.gdhMisc
    gdDestroyArray m.geAnnStruct.gshText
    gdDestroyArray m.geAnnStruct.gshFont
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.Terminate", eGDRaiseError_Raise

End Sub

Public Property Get eType() As eAnnotType
On Error GoTo ErrSection:

    eType = m.eType

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.eType.Get", eGDRaiseError_Raise

End Property

Public Property Get eUsage() As eAnnotUsage
On Error GoTo ErrSection:
    
    eUsage = m.eUsage

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.eUsage.Get", eGDRaiseError_Raise

End Property

Public Property Let eUsage(ByVal Usage As eAnnotUsage)
On Error GoTo ErrSection:
    
    m.eUsage = Usage
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.eUsage.Let", eGDRaiseError_Raise

End Property

Public Property Get PreIndicator() As Long
On Error GoTo ErrSection:
    
    PreIndicator = m.nPreIndicator

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.preIndicator.Get", eGDRaiseError_Raise

End Property

Public Property Let PreIndicator(ByVal nPreInd&)
On Error GoTo ErrSection:
    
    m.nPreIndicator = nPreInd

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.preIndicator.Let", eGDRaiseError_Raise

End Property

Public Property Get Pane() As String
On Error GoTo ErrSection:
    
    Pane = m.strPane

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Pane.Get", eGDRaiseError_Raise

End Property

Public Property Let Pane(ByVal strPane As String)
On Error GoTo ErrSection:
    
    m.strPane = strPane

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Pane.Let", eGDRaiseError_Raise

End Property

Public Property Get dDate(ByVal iPoint%) As Double
On Error GoTo ErrSection:
    
    If iPoint = 1 Then
        dDate = m.dDate1
    ElseIf iPoint = 2 Then
        dDate = m.dDate2
    ElseIf iPoint > 2 Then
        dDate = m.aDates(iPoint - 3)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.dDate.Get", eGDRaiseError_Raise

End Property

Public Property Let dDate(ByVal iPoint%, ByVal dDate As Double)
On Error GoTo ErrSection:
    
    If iPoint = 2 Then
        m.dDate2 = dDate
    Else
        m.dDate1 = dDate
    End If
       
    geSetPoints

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.dDate.Let", eGDRaiseError_Raise

End Property

Public Property Get Y(ByVal iPoint%) As Double
On Error GoTo ErrSection:
    
    Dim bSame As Boolean
    
    'Note: for balloon strangle Y values are stored as follows
    '   Use the balloon specific properties BalloonStock etc to access these
    '   Y1 = stock price
    '   Y2 = put strike
    '   aY(0) = call strike
    '   aY(1) = cost of put
    '   aY(2) = cost of call
    
    If m.eType = eANNOT_BalloonStrangle Then Exit Property
    
    If iPoint > 2 Then
        If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 _
           Or m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_FibABCD Then Y = m.aYs(0)
        Exit Sub
    End If
    
    If m.Y1 = m.Y2 And m.dDate1 = m.dDate2 Then bSame = True
    
    If iPoint = 2 Then
        If m.eType = eANNOT_GannLines And bSame Then
            '1=NE=up right, 7=NW=up left, 3=SE=down right, 5=SW=down left
            Select Case m.geAnnStruct.lenRatio
                Case 1, 7:
                    Y = m.Y1 + m.geAnnStruct.minorAxisLenData * m.geAnnStruct.lenPoints
                Case 3, 5:
                    Y = m.Y1 - m.geAnnStruct.minorAxisLenData * m.geAnnStruct.lenPoints
            End Select
        Else
            Y = m.Y2
        End If
    Else
        Y = m.Y1
    End If


ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Y.Get", eGDRaiseError_Raise

End Property

Public Property Let Y(ByVal iPoint%, ByVal dY As Double)
On Error GoTo ErrSection:
    
    If m.eType = eANNOT_BalloonStrangle Then Exit Property
    
    If iPoint = 2 Then
        m.Y2 = dY
    Else
        m.Y1 = dY
    End If
        
    geSetPoints

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Y.Let", eGDRaiseError_Raise

End Property

Public Sub FakeShift(ByVal nShift As Long)
On Error GoTo ErrSection:
    
    Dim dPoint#
    
    '08-18-2004:
    'Replaced all annot-specific shift code, except one, with new ShiftAnnot sub.
    'Leave original annot-specific shift code awhile then remove if all is okay.
    If m.eType = eANNOT_GannLines Then
        ShiftDTLine nShift
    Else
        ShiftAnnot nShift
    End If
    Exit Sub
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.FakeShift", eGDRaiseError_Raise

End Sub

Public Property Get DnrRptsCount() As Long
On Error GoTo ErrSection:
    
    DnrRptsCount = m.aDates.Size

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.DnrRptsCount.Get", eGDRaiseError_Raise

End Property

Public Property Get DnrRptDate(ByVal nPoint, ByVal bAsString As Boolean)
On Error GoTo ErrSection:
    
    If bAsString = True Then
        DnrRptDate = DateFormat(m.aDates(nPoint))
    Else
        DnrRptDate = m.aDates(nPoint)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.DnrRptDate.Get", eGDRaiseError_Raise

End Property

Public Property Get DnrLineageText(ByVal nPoint)
On Error GoTo ErrSection:
    
    Dim gdTextStr As New cGdArray
    
    gdTextStr.SplitFields m.strText, "~"
    DnrLineageText = gdTextStr(nPoint)
    Set gdTextStr = Nothing

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.DnrLineageText.Get", eGDRaiseError_Raise

End Property

Public Property Get X(ByVal iPoint%) As Double
On Error GoTo ErrSection:
    
    If iPoint = 2 Then
        X = m.X2
    Else
        X = m.X1
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.X.Get", eGDRaiseError_Raise

End Property

Public Property Let X(ByVal iPoint%, ByVal dX As Double)
On Error GoTo ErrSection:
    
    If iPoint = 2 Then
        m.X2 = dX
    Else
        m.X1 = dX
    End If
        
    geSetPoints

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.X.Let", eGDRaiseError_Raise

End Property

Public Sub geGetDim(Top#, Left#, Bottom#, Right#, ByVal nIdx&)
On Error GoTo ErrSection:

    If nIdx > gdGetSize(m.geAnnStruct.glhAnnType) Then Exit Sub
    
    Top = gdGetNum(m.geAnnStruct.gdhTop, nIdx)
    Left = gdGetNum(m.geAnnStruct.gdhLeft, nIdx)
    Bottom = gdGetNum(m.geAnnStruct.gdhBottom, nIdx)
    Right = gdGetNum(m.geAnnStruct.gdhRight, nIdx)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geGetDim", eGDRaiseError_Raise

End Sub

Public Property Get ArrowLocation() As Long
On Error GoTo ErrSection:

    ArrowLocation = gdGetNum(m.geAnnStruct.glhDirection, 1)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.ArrowLocation.Get", eGDRaiseError_Raise

End Property

Public Property Get MultiChartFlag() As Boolean
On Error GoTo ErrSection:

    MultiChartFlag = m.bMultiChart

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.MultiChartFlag.Get", eGDRaiseError_Raise

End Property

Public Property Let MultiChartFlag(ByVal bMulti As Boolean)
On Error GoTo ErrSection:

    m.bMultiChart = bMulti

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.MultiChartFlag.Let", eGDRaiseError_Raise

End Property

Public Property Get Color() As Long
On Error GoTo ErrSection:

    Color = m.nColor

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Color.Get", eGDRaiseError_Raise

End Property

Public Property Let Color(ByVal nColor As Long)
On Error GoTo ErrSection:

    m.nColor = nColor
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Color.Let", eGDRaiseError_Raise

End Property

Public Property Get Text() As String
On Error GoTo ErrSection:

    Text = m.strText

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Text.Get", eGDRaiseError_Raise

End Property

Public Property Let Text(ByVal strText As String)
On Error GoTo ErrSection:

    m.strText = strText

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Text.Let", eGDRaiseError_Raise

End Property

Public Property Get Style() As Long
On Error GoTo ErrSection:

    Style = m.nStyle

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Style.Get", eGDRaiseError_Raise

End Property

Public Property Let Style(ByVal nStyle As Long)
On Error GoTo ErrSection:

    m.nStyle = nStyle
    If m.eUsage = eANNOT_IndicatorLabel Then
        gdSetNum m.geAnnStruct.glhStyle, 0, nStyle
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Style.Let", eGDRaiseError_Raise

End Property

' to retrieve a property (varies by annotation type)
' NOTE: the "Get" always returns a string which is in English formatting,
' which means for numeric properties the caller should always be using "Val"
' instead of "ValOfText" -- e.g. = Val(Prop(...
Public Property Get Prop(ByVal strProperty$) As Variant
On Error GoTo ErrSection:
    
    Dim iPos&, iEq&, strValue$
    
    ' properties are kept sorted so can find quickly
    m.Properties.BinarySearch strProperty, iPos, eGdSort_IgnoreCase
    If iPos < m.Properties.Size Then
        strValue = m.Properties(iPos)
        iEq = InStr(strValue, "=")
        If iEq > 0 Then
            If UCase(Left(strValue, iEq - 1)) = UCase(strProperty) Then
                strValue = Mid(strValue, iEq + 1)
            Else
                strValue = ""
            End If
        End If
    End If
    Prop = strValue

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Prop.Get", eGDRaiseError_Raise

End Property

' to create or replace a property (varies by annotation type)
Public Property Let Prop(ByVal strProperty$, ByVal vValue As Variant)
On Error GoTo ErrSection:
    
    Dim iPos&, iEq&, strValue$, bExisting As Boolean
                        
    ' properties are kept sorted so can find quickly
    m.Properties.BinarySearch strProperty, iPos, eGdSort_IgnoreCase
        
    If iPos < m.Properties.Size Then
        strValue = m.Properties(iPos)
        iEq = InStr(strValue, "=")
        If iEq > 0 Then
            If UCase(Left(strValue, iEq - 1)) = UCase(strProperty) Then
                bExisting = True
            End If
        End If
    End If
    
    If bExisting Then
        If m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or _
           m.eType = eANNOT_DollarLine3 Or m.eType = eANNOT_DollarLine4 Then
            'JM 05-29-2014: fix issue of text not showing correctly when turning these on/off in editor
            If strProperty = "ProfitLoss" Or strProperty = "ProfitLossPercent" Or _
               strProperty = "PriceMove" Or strProperty = "SlopeLine" Or strProperty = "LenBars" Then
               
                If m.Properties(iPos) <> IniString(strProperty, vValue) Then
                    gdSetSize m.geAnnStruct.gdhMisc, 0, 0       'to trigger autoposition text in grapheng.dll
                End If
            
            End If
        End If
    End If
    
    ' create property string with value ("Prop=Value")
    strProperty = IniString(strProperty, vValue)
    If bExisting Then
        ' replace existing property
        m.Properties(iPos) = strProperty
    Else
        ' insert new property
        m.Properties.Add strProperty, iPos
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Prop.Let", eGDRaiseError_Raise

End Property

Public Sub LoadRatios(ByVal bUsePreset As Boolean, Optional ByVal bConvert As Boolean = False)
On Error GoTo ErrSection:

    Dim dValue#, nSize&, nColor&, i&
    Dim strRatio$, strShow$, strColor$, strBold$, strMsg$
    
    Dim aRatio As New cGdArray
    Dim aShow As New cGdArray
    Dim aColor As New cGdArray
    Dim aBold As New cGdArray

    Dim bDanCodeDefault As Boolean
               
    m.aRatioShow.Clear
    m.aRatios.Clear
    m.aRatioColor.Clear
    m.aRatioBold.Clear
    
    If m.eType = eANNOT_BalloonStrangle Then
        strRatio = Prop("BalloonRatios")
        If Len(strRatio) = 0 Then
            strRatio = kBalloonRatios
            Prop("BalloonRatios") = strRatio
        End If
        
        strShow = Prop("ShowRatios")
        If Len(strShow) = 0 Then
            strShow = kBalloonShow
            Prop("ShowRatios") = strShow
        End If
        
        strColor = Prop("BalloonColors")
        If Len(strColor) = 0 Then
            nColor = RGB(0, 128, 0)
            For i = 1 To 20
                m.aRatioColor.Add nColor
            Next
            Prop("BalloonColors") = m.aRatioColor.JoinFields(",")
        Else
            m.aRatioColor.SplitFields strColor, ","
        End If
        
        m.aRatios.SplitFields strRatio, ","
        m.aRatioShow.SplitFields strShow, ","
        
        GoTo ErrExit
        
    ElseIf bUsePreset = False Then
        If m.eType = eANNOT_SpResistFan Then
            strRatio = Prop("SpResistRatios")
        Else
            strRatio = Prop("FibRatios")
            strShow = Prop("ShowRatios")
            
            If Len(strRatio) < 1 Or Len(strShow) < 1 Then
                bUsePreset = True
            End If
            
            If m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 Or m.eType = eANNOT_Fibonacci3 Or _
               m.eType = eANNOT_Fibonacci4 Or m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_DanCodeFib Or _
               m.eType = eANNOT_AdvRiskReward Then
                strBold = Prop("FibBold")
            End If
        End If
    Else
        Select Case m.eType
            Case eANNOT_SpResistFan
                strRatio = kSpResistRatios
                
            Case eANNOT_FibExpansion
                strRatio = kFibExpFanRatios
                strShow = kFibExpShow
                strBold = kFibBold

            Case eANNOT_AdvRiskReward
                strRatio = kAdvRiskRewardRatios
                strShow = kAdvRiskRewardShow
                strBold = kFibBold
            
            Case eANNOT_FibFan
                strRatio = kFibExpFanRatios
                strShow = kFibFanShow
                
            Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
                 eANNOT_DNExpansion4, eANNOT_FibABCD
                strRatio = kFibDnExRatios
                strShow = kFibDnExShow
                Prop("FibColor") = kFibDnExColors
            
            Case eANNOT_FibTimeRatio
                strRatio = kFibTimeRatios
                strShow = kFibTimeShow
            
            Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, eANNOT_Fibonacci4, eANNOT_DanCodeFib
                strMsg = "Would you like to restore Fibonacci ratios or Daniel Code ratios?"
                If m.eType = eANNOT_DanCodeFib Then
                    If InfBox(strMsg, "?", "+Daniel|-Fibonacci", "Restore original ratios") = "D" Then
                        bDanCodeDefault = True
                    End If
                ElseIf HasModule("CODE") Then
                    If InfBox(strMsg, "?", "+Fibonacci|-Daniel", "Restore original ratios") = "D" Then
                        bDanCodeDefault = True
                    End If
                End If
                
                If bDanCodeDefault Then
                    strRatio = kDanCodeFib
                    strShow = kDanCodeFibShow
                    strColor = kDanCodeFibColor
                    Prop("ShowRatioZero") = 0
                    Prop("ShowRatioOne") = 0
                Else
                    strRatio = kFibRatios
                    strShow = kFibShow
                    strColor = "255"
                    Prop("ShowRatioZero") = 1
                    Prop("ShowRatioOne") = 1
                End If
                strBold = kFibBold
            
            Case eANNOT_FibArcs
                strRatio = Replace(kFibRatios, "-1,", "")
                strShow = Replace(kFibShow, "0,", "", , 1)
            
            Case Else
                strRatio = kFibRatios
                strShow = kFibShow
                
        End Select
    End If

    aRatio.SplitFields strRatio, ","
    
    If m.eType = eANNOT_SpResistFan Then
        aShow.SplitFields strRatio
    Else
        aShow.SplitFields strShow, ","
    End If
    
    If aRatio.Size <> aShow.Size Then Exit Sub
    
    nSize = aRatio.Size
    If m.eType = eANNOT_FibArcs Then
        For i = 0 To nSize - 1
            dValue = Abs(Val(aRatio(i)))    'disallow negatives
            If 1# - dValue = 0 Then         'disallow 1.00 ratio
                aRatio.Remove i
                aShow.Remove i
                nSize = nSize - 1
            End If
        Next
    ElseIf m.eType = eANNOT_SpResistFan Then
        For i = 0 To nSize - 1
            aShow(i) = 1
        Next
    End If
    
    If Len(strColor) > 0 Then
        m.aRatioColor.SplitFields strColor
    Else
        m.aRatioColor.SplitFields Prop("FibColor")
    End If
    
    If m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 Or m.eType = eANNOT_Fibonacci3 Or _
       m.eType = eANNOT_Fibonacci4 Or m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_DanCodeFib Or _
       m.eType = eANNOT_AdvRiskReward Then
        m.aRatioBold.SplitFields strBold
        For i = m.aRatioBold.Size To nSize - 1
            m.aRatioBold(i) = 0
        Next
        
        'bold fib lines is new as of 05-2012 (old templates will not have this property)
        If Len(strBold) = 0 Then Prop("FibBold") = m.aRatioBold.JoinFields(",")
    End If
    
    For i = m.aRatioColor.Size To nSize - 1
        m.aRatioColor(i) = vbRed
    Next
    
    'previous versions has the following properties
    'R50  - 0/1 = show/don't show 50% retracement line
    'Fibs - 0/1 = show/don't show fib lines ratios 0.618 and 0.382
    For i = 0 To nSize - 1
        dValue = Val(aRatio(i))
        m.aRatios.Add dValue
        If bUsePreset = True And bConvert = True Then
            If dValue = 0.5 Then
                dValue = Val(Prop("R50"))
            ElseIf dValue = 0.618 Or dValue = 0.382 Then
                dValue = Val(Prop("Fibs"))
            Else
                dValue = Val(aShow(i))
            End If
        Else
            dValue = Val(aShow(i))
        End If
        m.aRatioShow.Add Int(dValue)
    Next
    
    If m.eType = eANNOT_SpResistFan Then
        Prop("SpResistRatios") = strRatio
    Else
        Prop("FibRatios") = strRatio
        Prop("ShowRatios") = strShow
    End If
    
    Set aRatio = Nothing
    Set aShow = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.LoadRatios", eGDRaiseError_Raise

End Sub

Private Function LoadDefault(ByVal strProp$, ByVal vDefaultValue As Variant, _
        Optional ByVal bSetProp As Boolean = False) As Variant
On Error GoTo ErrSection:

    Dim strSection$, vValue As Variant
    
    strSection = "AnnotDefaults " & Str(m.eType)
    vValue = GetIniFileProperty(strProp, vDefaultValue, strSection, g.strIniFile)
    If bSetProp Then Prop(strProp) = vValue
    LoadDefault = vValue
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.LoadDefault", eGDRaiseError_Raise

End Function

Private Sub LoadDefaults()
On Error GoTo ErrSection:

    Dim i&
    
    'font (horz & vert lines don't use these, but can load anyways)
    LoadDefault "FontName", "arial", True
    LoadDefault "FontSize", 8, True
    LoadDefault "FontUnderline", 0, True
    
    If m.eType = eANNOT_TextEdit Or m.eType = eANNOT_TextEdit2 Or m.eType = eANNOT_TextEdit3 Or m.eType = eANNOT_TextEdit4 Then
        LoadDefault "FontStyle", 0, True    '0=reg,1=bold,2=italic,3=bold italic
    Else
        LoadDefault "FontStyle", 1, True    '0=reg,1=bold,2=italic,3=bold italic
    End If
        
    If m.eType = eANNOT_FibTimeRatio Or m.eType = eANNOT_FibTimeZones Or _
       m.eType = eANNOT_TimeCycle Or m.eType = eANNOT_VertLine Or _
       m.eType = eANNOT_TriangleWedge Or m.eType = eANNOT_HorzLine Or _
       m.eType = eANNOT_HorzLine2 Or m.eType = eANNOT_HorzLine3 Or m.eType = eANNOT_HorzLine4 Or _
       m.eType = eANNOT_ChannelHighlight Or m.eType = eANNOT_Bracket Or _
       m.eType = eANNOT_Gartley Or m.eType = eANNOT_DanCodeZone Then
        m.nPreIndicator = LoadDefault("PreIndicator", 1, False)
    Else
        m.nPreIndicator = LoadDefault("PreIndicator", 0, False)
    End If
        
    Select Case m.eType
        Case eANNOT_VertLine, eANNOT_AndrewFork, eANNOT_Mirror, _
             eANNOT_Pattern, eANNOT_WaveLabels, eANNOT_HorzLine, _
             eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4
            
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Default))
            m.bMultiChart = LoadDefault("MultiChart", False)
            
            If m.eType = eANNOT_AndrewFork Then
                'quarter lines - backwards compatibility
                LoadDefault "ShowQtrLines", 0, True
                
                LoadDefault "ForkMode", eANNOT_ForkNormal, True
                'parallel lines
                LoadDefault "ParallelStyle", eANNOT_DashSm, True
                LoadDefault "ParallelRatios", kParallelRatios, True
                LoadDefault "ParallelShow", kParallelShow, True
                LoadDefault "ParallelColor", kParallelColor, True
                LoadDefault "ParallelText", 1, True                         '0=don't show text
                'interval markers
                LoadDefault "MarkerStyle", eANNOT_Medium, True
                LoadDefault "MarkerRatios", kIntervalRatios, True
                LoadDefault "MarkerShow", kIntervalShow, True
                LoadDefault "MarkerColor", vbRed, True
                LoadDefault "MarkerText", 1, True                         '0=don't show text
                'miscellaneous
                LoadDefault "ShowHandle", 1, True           '1=extend base of fork to connect with parallel lines wih ratios > 1
                LoadDefault "HideABLine", 0, True           '1=don't draw AB line for Schiff/Modified Schiff fork
                LoadDefault "ShowInAllPanes", 0, True       '1=extend time interval marker to all panes
                LoadDefault "ParallelLocation", 0, True     '0=both,1=above,2=below (if there are no ratios checked then there are no parallels, ie equivalent to "none")
            ElseIf m.eType = eANNOT_HorzLine Or m.eType = eANNOT_HorzLine2 Or _
                   m.eType = eANNOT_HorzLine3 Or m.eType = eANNOT_HorzLine4 Or _
                   m.eType = eANNOT_VertLine Then
                LoadDefault "ShowInAxis", 0, True
            ElseIf m.eType = eANNOT_Pattern Then
                LoadDefault "ShowPatternName", 0, True
                LoadDefault "ForecastBars", 0, True
            ElseIf m.eType = eANNOT_WaveLabels Then
                LoadDefault "Lines", 1, True
                LoadDefault "LabelFirstPoint", 0, True
                LoadDefault "PointArc", 1, True
                LoadDefault "LabelPastEnd", "none", True
                m.strText = LoadDefault("DefaultWaveLabel", "1, 2, 3, 4, 5", True)
                If InStr(m.strText, ",") = 0 Then   'no commas means no labels, must have lines
                    Prop("Lines") = 1
                End If
                If Prop("Lines") = 0 Then           'no lines, must label first point
                    Prop("LabelFirstPoint") = 1
                End If
            End If
            
        Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4
            m.nColor = Val(LoadDefault("Color", RGB(0, 192, 0)))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Medium))
            LoadDefault "KeepAtEnd", 0, True
            LoadDefault "NumContracts", -1, True
            LoadDefault "NumForexContracts", -1, True
            LoadDefault "NumShares", 100, True
            LoadDefault "IndexMult", 1, True
            LoadDefault "ProfitLoss", 1, True
            LoadDefault "ProfitLossPercent", 0, True
            LoadDefault "PriceMove", 0, True
            LoadDefault "LenBars", 0, True
            LoadDefault "LenHzLine", 0, True
            LoadDefault "SlopeLine", 0, True
            LoadDefault "Volume", 0, True
            LoadDefault "PriceMovePoints", 0, True
            m.geTextAlign = LoadDefault("TextAlignment", 9, True) 'auto
            m.bMultiChart = LoadDefault("MultiChart", True)
            
        Case eANNOT_GannacciSwingSquare
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Medium))
            LoadDefault "IncludeBarOne", 0, True
            LoadDefault "PriceMove", 1, True
            LoadDefault "LenBars", 1, True
            LoadDefault "CalendarDays", 1, True
            LoadDefault "UseMutiplier", 0, True
            LoadDefault "MultiplierVal", 1, True
            LoadDefault "SameSwing", 1, True
            LoadDefault "SquareRange", 1, True
            LoadDefault "Ext", 0, True
            m.geTextAlign = LoadDefault("TextAlignment", 9, True) 'auto
            m.bMultiChart = LoadDefault("MultiChart", 0, False)
        
        Case eANNOT_GannacciTime
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Thin))
            LoadDefault "ShowTB", 1, True
            LoadDefault "ShowCD", 0, True
            
            LoadDefault "Show45", 1, True
            LoadDefault "Show90", 1, True
            LoadDefault "Show180", 1, True
            LoadDefault "Show270", 1, True
            LoadDefault "Show360", 1, True
            
            LoadDefault "ColorFor45", 32768, True
            LoadDefault "ColorFor90", 32768, True
            LoadDefault "ColorFor180", 32768, True
            LoadDefault "ColorFor270", 32768, True
            LoadDefault "ColorFor360", 32768, True
            
            LoadDefault "Ext", 1, True
            m.bMultiChart = LoadDefault("MultiChart", 0, False)
        
        Case eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Default))
            m.bMultiChart = LoadDefault("MultiChart", 0, False)
            
            LoadDefault "UseMutiplier", 0, True
            LoadDefault "MultiplierVal", 1, True
            LoadDefault "RoundDecimals", 0, True
            LoadDefault "Decimals", 0, True
            
            LoadDefault "IncludeBarOne", 0, True
            LoadDefault "Border", 1, True
            
            LoadDefault "FreeFloat", 1, True
            LoadDefault "ShowMarker", 1, True
            LoadDefault "ShowText", 1, True
            
            LoadDefault "SignalColorLow", vbRed, True           'strong signal (+/- 3 degrees)
            LoadDefault "SignalColorMedium", vbYellow, True     'stronger signal (+/- 2 degrees)
            LoadDefault "SignalColorHigh", vbGreen, True        'strongest signal (+/- 1/0 degrees)
            
            If m.eType = eANNOT_GannacciSwing1 Then
                LoadDefault "R", 1, True
                
                LoadDefault "Diff_TBP2", 1, True
                LoadDefault "Diff_TBR", 1, True
                LoadDefault "Diff_CDP2", 1, True
                LoadDefault "Diff_CDR", 1, True
                
                LoadDefault "Equal_TBP1", 1, True
                LoadDefault "Equal_CDP1", 1, True
                LoadDefault "Equal_TBR", 1, True
                LoadDefault "Equal_CDR", 1, True
                LoadDefault "Equal_TBP2", 1, True
                LoadDefault "Equal_CDP2", 1, True
            Else
                LoadDefault "TB1", 1, True
                LoadDefault "CD1", 1, True
                LoadDefault "TB2", 1, True
                LoadDefault "CD2", 1, True
                LoadDefault "R1", 1, True
                LoadDefault "R2", 1, True
            
                LoadDefault "Diff_TB1_R2", 1, True
                LoadDefault "Diff_CD1_R2", 1, True
                
                LoadDefault "Diff_TB2_R1", 1, True
                LoadDefault "Diff_CD2_R1", 1, True
            
                LoadDefault "Equal_TB1_R2", 1, True
                LoadDefault "Equal_CD1_R2", 1, True
            
                LoadDefault "Equal_TB2_R1", 1, True
                LoadDefault "Equal_CD2_R1", 1, True
            End If
        
        Case eANNOT_RegressionLine
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Thick))
            LoadDefault "FixLength", True, True
            LoadDefault "LenInBars", 0, True
            LoadDefault "KeepAtEnd", 0, True
            LoadDefault "StdDevVal", 1, True
            'LoadDefault "StdDevColor", RGB(0, 192, 0), True - save this awhile in case someone wants it
            'LoadDefault "StdDevStyle", eANNOT_Thin, True       -save awhile then remove 12-02-2003
            LoadDefault "StdDevStyle", eANNOT_Default, True
            LoadDefault "ChannelCount", 0, True
            LoadDefault "Ext", 0, True
            LoadDefault "ExtColor", vbRed, True
            LoadDefault "ExtStyle", eANNOT_DashLg, True
            LoadDefault "PriceField", 0, True '0=close,1=open,2=hi,3=low,4=AvgHiLow
            m.bMultiChart = LoadDefault("MultiChart", False)
            LoadDefault "ShowInAxis", 0, True
        
        Case eANNOT_GannacciCycle
            m.nColor = Val(LoadDefault("Color", vbBlue))
            LoadDefault "ImageType", eCNI_Triangle, True
            LoadDefault "ImageSize", 24, True        'small
            LoadDefault "ImageStyle", 0, True       'solid
            LoadDefault "ImageDir", eCNI_North, True
            LoadDefault "AsciiChar", "", True
            LoadDefault "GannacciYears", "5", True
            LoadDefault "Ext", 0, True
            
            m.nPreIndicator = LoadDefault("PreIndicator", 1, True)
            m.geTextAlign = LoadDefault("TextAlignment", 8, True) 'center
            m.nStyle = Val(LoadDefault("Style", -1))
            m.bMultiChart = LoadDefault("MultiChart", 0, False)

        Case eANNOT_Icon, eANNOT_ElliotLabel
            m.nColor = Val(LoadDefault("Color", vbRed))
            m.nStyle = Val(LoadDefault("Style", -1))    'using -1 as flag for backwards compatibility conversion
            If m.nStyle = -1 Then
                LoadDefault "ImageType", eCNI_Diamond, True
                LoadDefault "ImageSize", 8, True        'small
                LoadDefault "ImageStyle", 1, True       'solid
                LoadDefault "ImageDir", eCNI_North, True
                LoadDefault "AsciiChar", "", True
            End If
            m.strText = LoadDefault("Text", "1")
            m.geTextAlign = LoadDefault("TextAlignment", 8, True) 'center
            
            If m.eType = eANNOT_ElliotLabel Then
                LoadDefault "CustomText", "", True
                m.nColor = Val(LoadDefault("Color", vbBlack))   'this is used only for custom text in EWI palette
                LoadDefault "ImageType", eCNI_Ascii, True       'aardvark 3826
                Prop("ImageSize") = 999999
                LoadDefault "NumberStyle", 0, True
                LoadDefault "AlphaStyle", 0, True
                
                LoadDefault "UseCopyright", 1, True
                'JM 08-08-2013: aardvark 6898 - if user chooses the copyright text then
                '   the annotation is set up as if it were an edit text tool for drawing purposes
                '   the graphics engine needs these properties to draw fake edit text tool properly
                LoadDefault "TextJustify", 1, True  '0=left,1=right,2=center (align is being used to anchor text) - 4343
                LoadDefault "Shape", 0, True
                LoadDefault "Width", 20, True
                LoadDefault "Height", 10, True
                LoadDefault "Border", 0, True       '1=border, 0=no border
                LoadDefault "ArrowStyle", 0, True   '0=no arrow, 1=standard, 2=triangle
                LoadDefault "ArrowSize", 0, True    '0=small, 1=medium, 2=large
                LoadDefault "ArrowLine", eANNOT_Thin, True
                LoadDefault "Anchor", 8, True
                
'JM 07-22-2013: PresetStyle was originally inteded for use with one of the preset radio
'   buttons style, but was not needed and never used (never got around to removing it).
'   Since it's been so long, just commenting it out for now to make sure all is ok befor
'   removing it entirely from code.
'                LoadDefault "PresetStyle", -1, True
                
                
'JM 07-22-2013: Custom font sizes for EWIGMP palette per Glen/Rob email
                LoadDefault "UseCustomSizes", 0, True
'JM 07-25-2013: These are default custom sizes that EWI wants per Rob's email
                LoadDefault "CustomFontSizes", "12|12|12|16|16|14|18|18|20|22|22|22", True
                
                m.bMultiChart = LoadDefault("MultiChart", False, True)
            Else
                LoadDefault "Custom", 1, True
                m.bMultiChart = LoadDefault("MultiChart", True)
            End If
        
        Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
             eANNOT_TrendChannel
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Medium))
            m.strText = LoadDefault("Text", "")
            LoadDefault "Ext", 1, True
            LoadDefault "ExtColor", vbRed, True
            LoadDefault "ExtStyle", eANNOT_Default, True
            LoadDefault "ShowInAxis", 0, True
            m.geTextAlign = LoadDefault("TextAlignment", 9, True) 'auto
            m.bMultiChart = LoadDefault("MultiChart", True)
            LoadDefault "ChannelStyle", eANNOT_Default, True
            'LoadDefault "ChannelCount", 0, True
            'LoadDefault "ChannelPercent", 1, True
            'LoadDefault "ChannelPoints", 3, True
            LoadDefault "ChannelType", 0, True          '0=%price, 1=points
            LoadDefault "ChannelLocation", 3, True      '0=none,1=above,2=below,3=both
            LoadDefault "HideMainLine", 0, True
            
        Case eANNOT_SpResistFan
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Medium))
            m.bMultiChart = LoadDefault("MultiChart", False)
            LoadDefault "SpResistColor", vbRed, True
            LoadDefault "SpResistStyle", eANNOT_Default, True
            LoadDefault "SpResistRatios", kSpResistRatios, True
            LoadDefault "SpResistCount", 3, True
            Prop("ShowRatioZero") = 1
            Prop("ShowRatioOne") = 1
            LoadRatios False
            
        Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, eANNOT_Fibonacci4, _
             eANNOT_FibExpansion, eANNOT_DanCodeFib, eANNOT_FibArcs, eANNOT_FibFan, _
             eANNOT_FibTimeRatio, eANNOT_ElliotTimeRatio, eANNOT_AdvRiskReward
            
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.bMultiChart = LoadDefault("MultiChart", True)
            
            LoadDefault "FibColor", "255", True
            LoadDefault "FibStyle", eANNOT_Default, True
            LoadDefault "ShowValues", 0, True
            LoadDefault "ShowInAllPanes", 0, True
            LoadDefault "TextOnRight", -1, True
            LoadDefault "HideVerticalLine", 0, True
            LoadDefault "TextNextToMain", 0, True
            
            If eType = eANNOT_DanCodeFib Then
                LoadDefault "ShowRatioZero", 0, True
                LoadDefault "ShowRatioOne", 0, True
            Else
                LoadDefault "ShowRatioZero", 1, True
                LoadDefault "ShowRatioOne", 1, True
            End If
            
            If m.eType = eANNOT_ElliotTimeRatio Or m.eType = eANNOT_FibTimeRatio Then
                LoadDefault "ShowRatioTwo", 1, True
                LoadDefault "ShowNumBars", 0, True
                LoadDefault "ShowDates", 1, True
                If m.eType = eANNOT_FibTimeRatio Then
                    LoadDefault "RatioOneVal", 0, True
                End If
            End If
            
            If m.eType = eANNOT_FibArcs Then
                m.nStyle = Val(LoadDefault("Style", eANNOT_Default))
                LoadDefault "DirNE", 1, True
                LoadDefault "DirNW", 1, True
                LoadDefault "DirSE", 1, True
                LoadDefault "DirSW", 1, True
                LoadDefault "Ext", 0, True
                LoadDefault "Circular", 0, True
            Else
                m.nStyle = Val(LoadDefault("Style", eANNOT_Medium))
                LoadDefault "ExtColor", vbRed, True
                LoadDefault "ExtStyle", eANNOT_DashSm, True
                
                'JM 09-16-2015: this only applies to Fib retracement/expansion & DanCodeFib variant
                '   replaced with FibTextRatio, FibTextPrice for retracement & DanCodeFib
                '   added FibTextPL for FibExpansion to implement Coyote request for Risk Reward display
                '   leave commented out for reference / backwards compatibility
                'LoadDefault "PriceOnly", 0, True
                
                If m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 Or m.eType = eANNOT_Fibonacci3 Or _
                   m.eType = eANNOT_Fibonacci4 Or m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_DanCodeFib Or _
                   m.eType = eANNOT_AdvRiskReward Then
                   
                    LoadDefault "FibTextRatio", 1, True
                    LoadDefault "FibTextPrice", 1, True
                    LoadDefault "FibTextRatioLoc", 0, True      'loc: 0=left,1=right
                    LoadDefault "FibTextPriceLoc", 1, True      'loc: 0=left,1=right,2=y-scale
                    
                    If m.eType = eANNOT_AdvRiskReward Then
                        LoadDefault "FibTextPL", 0, True        '0=hide,1=show
                        LoadDefault "FibLabelPL", 1, True       '0=hide,1=show
                        LoadDefault "FibTextPLLoc", 0, True     'loc: 0=left,1=right
                        LoadDefault "FibLabelPLLoc", 1, True    'loc: 0=left,1=right
                        LoadDefault "FibRatiosText", "", True
                    End If
                End If
            End If
            
            If m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_AdvRiskReward Then
                If m.eType = eANNOT_FibExpansion Then
                    LoadDefault "FibRatios", kFibExpFanRatios, True
                    LoadDefault "ShowRatios", kFibExpShow, True
                Else
                    LoadDefault "FibRatios", kAdvRiskRewardRatios, True
                    LoadDefault "ShowRatios", kAdvRiskRewardShow, True
                    LoadDefault "FibColor", "255,255,255,255,255,32768,32768", True
                    LoadDefault "textZero", "", True
                    LoadDefault "textOne", "", True
                End If
                LoadDefault "ReverseDirection", 0, True
                LoadDefault "FibBold", kFibBold, True
                LoadDefault "BoldRatioZero", 0, True
                LoadDefault "BoldRatioOne", 0, True
                LoadDefault "KeepAlive", 0, True
                LoadDefault "Ext", 0, True                  '4313
            ElseIf m.eType = eANNOT_FibFan Then
                LoadDefault "FibRatios", kFibExpFanRatios, True
                LoadDefault "ShowRatios", kFibFanShow, True
                LoadDefault "Ext", 1, True
            ElseIf eType = eANNOT_DanCodeFib Then
                LoadDefault "FibRatios", kDanCodeFib, True
                LoadDefault "ShowRatios", kDanCodeFibShow, True
                LoadDefault "FibColor", kDanCodeFibColor, True
                LoadDefault "FibBold", kFibBold, True
                LoadDefault "BoldRatioZero", 0, True
                LoadDefault "BoldRatioOne", 0, True
                LoadDefault "KeepAlive", 0, True
                LoadDefault "Ext", 0, True
            ElseIf eType = eANNOT_FibTimeRatio Then
                LoadDefault "FibRatios", kFibTimeRatios, True
                LoadDefault "ShowRatios", kFibTimeShow, True
                LoadDefault "Ext", 0, True
            ElseIf eType = eANNOT_FibArcs Then
                LoadDefault "FibRatios", Replace(kFibRatios, "-1,", "", , 1), True
                LoadDefault "ShowRatios", Replace(kFibShow, "0,", "", , 1), True
                LoadDefault "Ext", 0, True
            Else
                LoadDefault "FibRatios", kFibRatios, True
                LoadDefault "ShowRatios", kFibShow, True
                LoadDefault "FibBold", kFibBold, True
                LoadDefault "BoldRatioZero", 0, True
                LoadDefault "BoldRatioOne", 0, True
                LoadDefault "KeepAlive", 0, True
                LoadDefault "Ext", 0, True
            End If
            LoadRatios False
            
        Case eANNOT_FibTimeZones, eANNOT_DanCodeZone
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_DashDot))
            LoadDefault "ShowInAllPanes", 0, True
            LoadDefault "ShowValues", 0, True
            LoadDefault "Arcs", 0, True
            LoadDefault "ArcPercentHeight", 0, True
            LoadDefault "FromPrevBar", 0, True          '0=draw lines as #of lines from first bar, 1=as #of lines from previous bar
            If m.eType = eANNOT_DanCodeZone Then
                LoadDefault "ZoneInUse", eANNOT_GB_ZoneDanCode, True
                LoadDefault "ShowValues", 1, True
            ElseIf IsGreenBlattTool Then
                LoadDefault "ZoneInUse", eANNOT_GB_ZoneLucas, True
            Else
                LoadDefault "ZoneInUse", eANNOT_GB_ZoneFib, True
            End If
               
        Case eANNOT_TargetShooter
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Default))
            LoadDefault "TargetColor", vbBlue, True
            LoadDefault "TargetStyle", eANNOT_Medium, True
            LoadDefault "Pullbacks", 0, True
            LoadDefault "Pt1", 1, True
            LoadDefault "Pt2", 1, True
            LoadDefault "Pt3", 1, True
            LoadDefault "ShowValues", 1, True
            LoadDefault "ExtMult", 4, True
            m.bMultiChart = LoadDefault("MultiChart", True)
                        
        Case eANNOT_Rectangle, eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4, eANNOT_ArrowLine
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Default))
            m.strText = LoadDefault("Text", "")
            geTextAlign = LoadDefault("TextAlignment", 8, True)   'e_center - 6801
            LoadDefault "Shape", 0, True
            LoadDefault "FillColor", m.nColor, True
            LoadDefault "FillPattern", 0, True
            If m.eType = eANNOT_ArrowLine Then
                LoadDefault "Width", 20, True
                LoadDefault "Height", 10, True
                LoadDefault "ArrowStyle", 1, True    '0=no arrow, 1=standard, 2=triangle
                LoadDefault "ArrowSize", 0, True    '0=small, 1=medium, 2=large
                LoadDefault "ArrowLine", eANNOT_Default, True
                LoadDefault "Anchor", 4, True
                m.strText = ""
            ElseIf m.eType = eANNOT_TextEdit Or m.eType = eANNOT_TextEdit2 _
                Or m.eType = eANNOT_TextEdit3 Or m.eType = eANNOT_TextEdit4 Then
                LoadDefault "Width", 20, True
                LoadDefault "Height", 10, True
                LoadDefault "Border", 1, True    '1=border, 0=no border
                LoadDefault "ArrowStyle", 0, True    '0=no arrow, 1=standard, 2=triangle
                LoadDefault "ArrowSize", 0, True    '0=small, 1=medium, 2=large
                LoadDefault "TextJustify", 0, True  '0=left,1=right,2=center (align is being used to anchor text) - 4343
                LoadDefault "ArrowLine", eANNOT_Default, True
                LoadDefault "Anchor", 4, True
                If Len(m.strText) > 0 Then m.strText = Replace(m.strText, "~", vbCrLf)  'aardvark 3407
            ElseIf m.eType = eANNOT_Rectangle Then
                m.nPreIndicator = LoadDefault("PreIndicator", 1, True)
                LoadDefault "FillColor", 8454143, True
                LoadDefault "FillPattern", 1, True
            End If
            m.bMultiChart = LoadDefault("MultiChart", False)
            
        Case eANNOT_Ellipse
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Medium))
            LoadDefault "MinorAxisLen", 0.25, True
            LoadDefault "MinorAxisLenData", 0, True     '0=ratio, 1=points, 2=ticks
            LoadDefault "ShowAxes", 1, True
            LoadDefault "ShowQtrLines", 0, True
        
        Case eANNOT_GannLines
            m.nColor = Val(LoadDefault("Color", vbRed))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Default))
            LoadDefault "Color1x2", vbBlue, True
            LoadDefault "Color1x3", vbBlue, True
            LoadDefault "Color1x4", vbBlue, True
            LoadDefault "Color1x8", vbBlue, True
            LoadDefault "GannFan", 1, True          '0=angle, 1=fan
            LoadDefault "GannRatioType", 0, True    '0=points, 1=ticks
            LoadDefault "Rise", 0.5, True
            LoadDefault "Run", 1, True
            LoadDefault "DirNE", 1, True
            LoadDefault "DirNW", 0, True
            LoadDefault "DirSE", 0, True
            LoadDefault "DirSW", 0, True
            LoadDefault "1x1", 1, True
            LoadDefault "1x2", 1, True
            LoadDefault "1x3", 1, True
            LoadDefault "1x4", 1, True
            LoadDefault "1x8", 1, True
            LoadDefault "2x1", 1, True
            LoadDefault "3x1", 1, True
            LoadDefault "4x1", 1, True
            LoadDefault "8x1", 1, True
            LoadDefault "SecondHotSpot", 0, True
            'hidden test flag for displaying angle value
            If Val(LoadDefault("ShowAxes", 0, False)) = 1 Then Prop("ShowAxes") = 1
       
        Case eANNOT_DNRetracement
            'Each ratio has a corresponding color, pen size and arc property.
            'Each arc has its own color and pen size properties.
            m.nColor = vbBlack      'ignored
            m.nStyle = 2            'ignored
            
            LoadDefault "FocusLow", 0, True     '0 -> focus value is bar's high else focus value is bar's low
            LoadDefault "FocusDynamic", 0, True '0 -> never move focus else update focus to highest or lowest bar
            LoadDefault "FibR0", 1#, True
            LoadDefault "FibR1", 0.618, True
            LoadDefault "FibR2", 0.382, True
            LoadDefault "FibColorR0", vbBlack, True
            LoadDefault "FibColorR1", vbMagenta, True
            LoadDefault "FibColorR2", vbRed, True
            LoadDefault "FibPenSizeR0", 2, True
            LoadDefault "FibPensizeR1", 2, True
            LoadDefault "FibPensizeR2", 2, True
            LoadDefault "ArcColorR1", RGB(128, 128, 128), True  'dark gray
            LoadDefault "ArcColorR2", RGB(128, 128, 128), True  'dark gray
            LoadDefault "ArcPensizeR1", 1, True
            LoadDefault "ArcPensizeR2", 1, True
            LoadDefault "ShowHandle", 1, True
            m.bMultiChart = LoadDefault("MultiChart", False)
            
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
             eANNOT_DNExpansion4, eANNOT_FibABCD
            m.nColor = vbBlack  'ignored
            m.nStyle = LoadDefault("Style", eANNOT_Default)
                        
            LoadDefault "HiLo", 0, True   '0=hi,1=low - this is for point A
                        
            If m.eType = eANNOT_DNExpansion And UseDiNapFib() Then
                'pen colors
                LoadDefault "colorA", vbBlack, True         'aardvark 6249
                LoadDefault "colorB", vbBlack, True
                LoadDefault "colorC", vbBlack, True
                LoadDefault "colorCOP", vbBlue, True
                LoadDefault "colorOP", vbGreen, True
                LoadDefault "colorXOP", vbRed, True
                'pen sizes
                LoadDefault "penSizeCOP", m.nStyle, True
                LoadDefault "penSizeOP", m.nStyle, True
                LoadDefault "penSizeXOP", m.nStyle, True
                'ratios
                LoadDefault "ratioCOP", 0.618, True
                LoadDefault "ratioOP", 1#, True
                LoadDefault "ratioXOP", 1.618, True
                'text
                If Len(LoadDefault("FontName", "")) = 0 Then        'aardvark 3825
                    LoadDefault "textA", "A", True
                    LoadDefault "textB", "B", True
                    LoadDefault "textC", "C", True
                    LoadDefault "textCOP", "COP", True
                    LoadDefault "textOP", "OP", True
                    LoadDefault "textXOP", "XOP", True
                Else
                    LoadDefault "textA", "", True
                    LoadDefault "textB", "", True
                    LoadDefault "textC", "", True
                    LoadDefault "textCOP", "", True
                    LoadDefault "textOP", "", True
                    LoadDefault "textXOP", "", True
                End If
            Else
                'pen colors & sizes for user-clicked points
                LoadDefault "colorA", vbBlack, True
                LoadDefault "colorB", vbBlack, True
                LoadDefault "colorC", vbBlack, True
                LoadDefault "penSizeA", m.nStyle, True
                LoadDefault "penSizeB", m.nStyle, True
                LoadDefault "penSizeC", m.nStyle, True
                
                LoadDefault "FibRatios", kFibDnExRatios, True
                LoadDefault "ShowRatios", kFibDnExShow, True
                LoadDefault "FibColor", kFibDnExColors, True
                LoadDefault "FibRatiosText", "", True
                LoadDefault "FibStyle", eANNOT_Default, True
                LoadDefault "TextOnRight", 0, True
                LoadDefault "ShowRatioZero", 1, True
                LoadDefault "ShowRatioOne", 1, True
                LoadDefault "ShowRatioTwo", 1, True
                LoadDefault "Ext", 1, True
                LoadDefault "ExtMain", 0, True
                LoadDefault "FreeFloat", 0, True
                LoadDefault "ShowValues", 1, True
                If m.eType = eANNOT_FibABCD Then
                    LoadDefault "RatioD", "1", True    'determines length of CD line
                    LoadDefault "colorD", vbBlack, True
                End If
                'text
                If Len(LoadDefault("FontName", "")) = 0 Then        'aardvark 3825
                    LoadDefault "textA", "A", True
                    LoadDefault "textB", "B", True
                    LoadDefault "textC", "C", True
                    If m.eType = eANNOT_FibABCD Then LoadDefault "textD", "D", True
                Else
                    LoadDefault "textA", "", True
                    LoadDefault "textB", "", True
                    LoadDefault "textC", "", True
                    If m.eType = eANNOT_FibABCD Then LoadDefault "textD", "", True
                End If
                LoadRatios False
            End If
            
            LoadDefault "ShowHandle", 1, True
            m.bMultiChart = LoadDefault("MultiChart", False)
    
        Case eANNOT_TimeCycle
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_DashDot))
            LoadDefault "ShowInAllPanes", 0, True
            LoadDefault "Arcs", 0, True
            LoadDefault "ArcPercentHeight", 0, True
    
        Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Thin))
            m.bMultiChart = LoadDefault("MultiChart", False)
            LoadDefault "ShowValues", 1, True
            LoadDefault "TextAlignment", 7, True
            LoadDefault "Ext", 0, True
            LoadDefault "ShowInAxis", 0, True
            LoadDefault "HideSRLineDot", 0, True
            
        Case eANNOT_Pivot
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Medium))
            m.bMultiChart = LoadDefault("MultiChart", False)
            LoadDefault "CalcMethod", 1, True
            LoadDefault "ShowValues", 1, True
            LoadDefault "PriceOnly", 0, True
            LoadDefault "Ext", 0, True                  'extension: 0=none,1=right,2=left,3=both
            LoadDefault "ExtStyle", eANNOT_DashSm, True
            LoadDefault "PivotColor", kPivotColor, True
            LoadDefault "S1Color", kS1Color, True
            LoadDefault "S2Color", kS2Color, True
            LoadDefault "S3Color", kS3Color, True
            LoadDefault "R1Color", kR1Color, True
            LoadDefault "R2Color", kR2Color, True
            LoadDefault "R3Color", kR3Color, True
            LoadDefault "PivotShow", 34, True       '34 is annotation type for pivot points in grapheng.dll
            LoadDefault "S1Show", 34, True
            LoadDefault "S2Show", 34, True
            LoadDefault "S3Show", 34, True
            LoadDefault "R1Show", 34, True
            LoadDefault "R2Show", 34, True
            LoadDefault "R3Show", 34, True
            
        Case eANNOT_MultiRects
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Thin))
            LoadDefault "FillColor", m.nColor, True
            LoadDefault "FillPattern", 0, True
            
        Case eANNOT_RiskReward
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Thin))
            LoadDefault "ShowProfitLoss", 1, True
            LoadDefault "ShowValues", 0, True
            LoadDefault "UserEndPoint1", -1, True
            LoadDefault "UserEndPoint2", -1, True
            LoadDefault "NumContracts", -1, True
            LoadDefault "NumForexContracts", -1, True
            LoadDefault "NumShares", 100, True
            LoadDefault "IndexMult", 1, True
            
        Case eANNOT_TriangleWedge, eANNOT_ChannelHighlight
            m.nColor = Val(LoadDefault("Color", RGB(0, 0, 192)))    'dark blue
            m.nStyle = Val(LoadDefault("Style", eANNOT_Thin))
            LoadDefault "FillColor", RGB(224, 224, 224), True   'light gray
            LoadDefault "FillPattern", 1, True
            If m.eType = eANNOT_ChannelHighlight Then
                Prop("LenRatio") = -1               'set to autodetect sides to outline
            End If
            
        Case eANNOT_Bracket
            m.nColor = Val(LoadDefault("Color", vbBlue, True))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Thin))
            LoadDefault "BracketStyle", 0, True                 '0=curly, 1=square
            LoadDefault "BracketDirection", 0, True             '0=open left, 1=open right
            m.bMultiChart = LoadDefault("MultiChart", False)

        Case eANNOT_Gartley
            m.nColor = Val(LoadDefault("Color", vbBlue, True))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Thin))
            LoadDefault "FontSize", 11, True
            
            LoadDefault "TriangleLabel", "", True
            LoadDefault "GartleyLabelStyle", 1, True            '0=none, 1=text, 2=text & value
            LoadDefault "GartleyLabels", "X|A|B|C|D", True
            
            LoadDefault "ShowHandle", 1, True                   'mark endpoints with circles
            
            LoadDefault "Triangle1_Fill", 2, True               '1=fill using triangle pen color, 2=flexUnchecked
            LoadDefault "Triangle2_Fill", 2, True
            LoadDefault "Triangle3_Fill", 2, True
            LoadDefault "Triangle4_Fill", 2, True
            
            LoadDefault "Triangle1_Color", vbBlue, True
            LoadDefault "Triangle2_Color", vbBlue, True
            LoadDefault "Triangle3_Color", vbBlue, True
            LoadDefault "Triangle4_Color", vbBlue, True

            LoadDefault "Triangle1", 1, True                'xB
            LoadDefault "Triangle2", 1, True                'xD
            LoadDefault "Triangle3", 1, True                'AC
            LoadDefault "Triangle4", 1, True                'BD

        Case eANNOT_BalloonStrangle
            m.nColor = Val(LoadDefault("Color", vbBlue))
            m.nStyle = Val(LoadDefault("Style", eANNOT_Medium))
            m.bMultiChart = LoadDefault("MultiChart", False)
            m.geTextAlign = LoadDefault("TextAlignment", 6, True) '6=right, 7=left
            LoadDefault "RiskBEStyle", eANNOT_Thin, True
            LoadDefault "RiskBEColor", RGB(0, 128, 0), True
            LoadDefault "ShowBE", 1, True
            LoadDefault "ShowPrices", 1, True
            LoadDefault "ShowTradeCost", 1, True
            LoadDefault "ShowProfitLoss", 1, True
            LoadDefault "ShowPriceMove", 1, True
            LoadDefault "ShowCollapsed", 0, True
            LoadDefault "ShowRatios", "", True
            LoadDefault "BalloonRatios", "", True
            LoadDefault "BalloonColors", "", True
            LoadDefault "ShowText", 1, True
            m.strText = LoadDefault("Text", "")
            If Len(m.strText) > 0 Then m.strText = Replace(m.strText, "~", vbCrLf)
            
            LoadRatios False
            
    End Select
    
    ' strip out old alignment prefix
    If Left(m.strText, 1) = "|" Then
        m.strText = Mid(m.strText, 3)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.LoadDefaults", eGDRaiseError_Raise

End Sub

Private Sub SaveDefault(ByVal strProp$, Optional ByVal vValue As Variant)
On Error GoTo ErrSection:

    Dim strSection$, strValue$
    
    If Len(Trim(strProp)) > 0 Then
        strSection = "AnnotDefaults " & Str(m.eType)
        If IsMissing(vValue) Then
            strValue = Prop(strProp)
        Else
            strValue = Str(vValue)
        End If
        SetIniFileProperty strProp, strValue, strSection, g.strIniFile
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SaveDefault", eGDRaiseError_Raise

End Sub

Public Sub SaveDefaults(Optional ByVal strSpecialProp$ = "")
On Error GoTo ErrSection:

    If m.eUsage = eANNOT_FibClusters Then
        SetIniFileProperty "ZoneFillColor", Prop("FillColor"), kClusterIniSection, g.strIniFile
        SetIniFileProperty "ZoneColor", m.nColor, kClusterIniSection, g.strIniFile
        GoTo ErrExit
    End If

    If m.strPane = "PRICE PANE" Then
        SaveDefault "MultiChart", m.bMultiChart
    End If
    
    SaveDefault "PreIndicator", m.nPreIndicator
    
    If m.eType <> eANNOT_ElliotLabel Or m.strText = "FakeText" Then
        'do not save font or color defaults for elliott label if user is not using custom text
        'the individual single-char label have their own color and font defaults
        SaveDefault "Color", m.nColor
        SaveDefault "Style", m.nStyle
        SaveDefault "FontName"
        SaveDefault "FontSize"
        SaveDefault "FontStyle"
        SaveDefault "FontUnderline"
        SaveDefault "Border"
    End If
    
    If m.eType = eANNOT_TextEdit Or m.eType = eANNOT_TextEdit2 _
       Or m.eType = eANNOT_TextEdit3 Or m.eType = eANNOT_TextEdit4 _
       Or m.eType = eANNOT_BalloonStrangle Then

        SaveDefault "Text", Replace(m.strText, vbCrLf, "~")     'aardvark 3407, 6813
    ElseIf m.strText = "FakeText" Then
        m.strText = Replace(Prop("CustomText"), vbCrLf, "~")
        SaveDefault "CustomText", m.strText
        m.strText = ""
    ElseIf m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or _
           m.eType = eANNOT_DollarLine3 Or m.eType = eANNOT_DollarLine4 Then
           
           'JM 09-14-2015: $Line text were always saved, but this used to be short 1-liner
           '    with all the additions, this can now be multi-lines that looks gunky in INI file
           '    although continuing to save the text doesn't hurt anything, it's nicer to not clutter INI file
           '    so - don't save text for $lines
    Else
        SaveDefault "Text", m.strText
    End If
    
' JM 09-14-2015: original code -
'                looks like a rushed bug-fix job for Elliott Wave custom text labels
'                rewrote above (leave awhile then remove if all ok)
'
'    If m.eType = eANNOT_TextEdit Or m.eType = eANNOT_TextEdit2 _
'       Or m.eType = eANNOT_TextEdit3 Or m.eType = eANNOT_TextEdit4 _
'       Or m.eType = eANNOT_BalloonStrangle Then
'
'        SaveDefault "Text", Replace(m.strText, vbCrLf, "~")     'aardvark 3407, 6813
'
'    ElseIf m.strText <> "FakeText" Then
'        SaveDefault "Text", m.strText
'    End If
'
'    SaveDefault "PreIndicator", m.nPreIndicator
'
'    'font (horz & vert lines don't use, but can save anyways)
'    If m.eType <> eANNOT_ElliotLabel Or m.strText = "FakeText" Then
'        'do not save font or color defaults for elliott label if user is not using custom text
'        'the individual single-char label have their own color and font defaults
'        SaveDefault "Color", m.nColor
'        SaveDefault "Style", m.nStyle
'        SaveDefault "FontName"
'        SaveDefault "FontSize"
'        SaveDefault "FontStyle"
'        SaveDefault "FontUnderline"
'        SaveDefault "Border"
'
'        If m.eType Or eANNOT_TextEdit Or m.eType = eANNOT_TextEdit2 _
'           Or m.eType = eANNOT_TextEdit3 Or m.eType = eANNOT_TextEdit4 _
'           Or m.eType = eANNOT_BalloonStrangle Then
'           'do nothing  -7033
'        Else
'            m.strText = Replace(Prop("CustomText"), vbCrLf, "~")
'            SaveDefault "CustomText", m.strText
'            m.strText = ""
'        End If
'    End If
    
    Select Case m.eType
        Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4, eANNOT_VertLine
            SaveDefault "ShowInAxis"

        Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
             eANNOT_TrendChannel
            SaveDefault "Ext"
            SaveDefault "ExtColor"
            SaveDefault "ExtStyle"
            SaveDefault "TextAlignment"
            SaveDefault "ShowInAxis"
                
            SaveDefault "ChannelCount"
            SaveDefault "ChannelStyle"
            SaveDefault "ChannelType"
            SaveDefault "ChannelPoints"
            SaveDefault "ChannelPercent"
            SaveDefault "ChannelLocation"
            
            If m.eType = eANNOT_TrendChannel Then SaveDefault "HideMainLine"
        
        Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4
            SaveDefault strSpecialProp
            SaveDefault "TextAlignment"
            SaveDefault "KeepAtEnd"
            SaveDefault "ProfitLoss"
            SaveDefault "PriceMove"
            SaveDefault "SlopeLine"
            SaveDefault "LenBars"
            SaveDefault "LenHzLine"
            SaveDefault "ProfitLossPercent"
            SaveDefault "Volume"
            SaveDefault "PriceMovePoints"
            
        Case eANNOT_GannacciSwingSquare
            SaveDefault "PriceMove"
            SaveDefault "LenBars"
            SaveDefault "CalendarDays"
            SaveDefault "UseMutiplier"
            SaveDefault "MultiplierVal"
            SaveDefault "IncludeBarOne"
            SaveDefault "SquareRange"
            SaveDefault "Ext"
        
        Case eANNOT_GannacciTime
            SaveDefault "ShowTB"
            SaveDefault "ShowCD"
            
            SaveDefault "Show45"
            SaveDefault "Show90"
            SaveDefault "Show180"
            SaveDefault "Show270"
            SaveDefault "Show360"
            
            SaveDefault "ColorFor45"
            SaveDefault "ColorFor90"
            SaveDefault "ColorFor180"
            SaveDefault "ColorFor270"
            SaveDefault "ColorFor360"
            
            SaveDefault "Ext"
        
        Case eANNOT_GannacciCycle
            SaveDefault "GannacciYears"
            SaveDefault "Ext"
        
        Case eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
            SaveDefault "UseMutiplier"
            SaveDefault "MultiplierVal"
            
            SaveDefault "RoundDecimals"
            SaveDefault "Decimals"
            
            SaveDefault "IncludeBarOne"
            SaveDefault "Border"
            
            SaveDefault "ShowMarker"
            SaveDefault "ShowText"
            
            SaveDefault "SignalColorLow"
            SaveDefault "SignalColorMedium"
            SaveDefault "SignalColorHigh"
            
            If m.eType = eANNOT_GannacciSwing1 Then
                SaveDefault "TB"
                SaveDefault "CD"
                SaveDefault "P1"
                SaveDefault "P2"
                SaveDefault "R"
                
                SaveDefault "Diff_TBP1"
                SaveDefault "Diff_TBP2"
                SaveDefault "Diff_TBR"
                SaveDefault "Diff_CDP1"
                SaveDefault "Diff_CDP2"
                SaveDefault "Diff_CDR"
                
                SaveDefault "Equal_TBP1"
                SaveDefault "Equal_CDP1"
                SaveDefault "Equal_TBR"
                SaveDefault "Equal_CDR"
            Else
                SaveDefault "TB1"
                SaveDefault "CD1"
                SaveDefault "TB2"
                SaveDefault "CD2"
                SaveDefault "R1"
                SaveDefault "R2"
            
                SaveDefault "Diff_TB1_R2"
                SaveDefault "Diff_CD1_R2"
                
                SaveDefault "Diff_TB2_R1"
                SaveDefault "Diff_CD2_R1"
            
                SaveDefault "Equal_TB1_R2"
                SaveDefault "Equal_CD1_R2"
            
                SaveDefault "Equal_TB2_R1"
                SaveDefault "Equal_CD2_R1"
            End If
            
        Case eANNOT_Icon
            SaveDefault "TextAlignment"
            SaveDefault "ImageType"
            SaveDefault "ImageSize"
            SaveDefault "ImageStyle"
            SaveDefault "ImageDir"
            SaveDefault "Style", -1
            
        Case eANNOT_ElliotLabel
            SaveDefault "UseCopyright"
            SaveDefault "TextJustify"
            SaveDefault "TextAlignment"
            SaveDefault "ImageType"
            SaveDefault "ImageSize"
            SaveDefault "ImageStyle"
            SaveDefault "NumberStyle"
            SaveDefault "AlphaStyle"
            SaveDefault "UseCustomSizes"
            SaveDefault "CustomFontSizes"
            SaveDefault "MultiChart", m.bMultiChart
                    
        Case eANNOT_Rectangle
            SaveDefault "Shape"
            SaveDefault "TextAlignment"
            SaveDefault "FillColor"
            SaveDefault "FillPattern"
            
        Case eANNOT_TriangleWedge, eANNOT_ChannelHighlight
            SaveDefault "FillColor"
            SaveDefault "FillPattern"
        
        Case eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4
            SaveDefault "Border"
            SaveDefault "ArrowStyle"
            SaveDefault "ArrowSize"
            SaveDefault "ArrowLine"
            SaveDefault "Anchor"
            SaveDefault "Width"
            SaveDefault "Height"
            SaveDefault "TextJustify"
            SaveDefault "TextAlignment"         '- 6801
        
        Case eANNOT_ArrowLine
            SaveDefault "ArrowStyle"
            SaveDefault "ArrowSize"
            SaveDefault "ArrowLine"
            
        Case eANNOT_Ellipse
            SaveDefault "MinorAxisLen"
            SaveDefault "MinorAxisLenData"     '0=ratio, 1=points, 2=ticks
            SaveDefault "ShowAxes"
            SaveDefault "ShowQtrLines"
            
        Case eANNOT_RegressionLine
            SaveDefault "Ext"
            SaveDefault "ExtColor"
            SaveDefault "ExtStyle"
            SaveDefault "StdDevVal"
            'SaveDefault "StdDevColor"
            SaveDefault "StdDevStyle"
            SaveDefault "KeepAtEnd"
            SaveDefault "PriceField"
            SaveDefault "ChannelCount"          '3761
            SaveDefault "FixLength"             '4222
            SaveDefault "ShowInAxis"
            
        Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, eANNOT_Fibonacci4, _
            eANNOT_FibTimeRatio, eANNOT_FibArcs, eANNOT_AdvRiskReward, _
            eANNOT_FibExpansion, eANNOT_FibFan, eANNOT_ElliotTimeRatio, _
            eANNOT_DanCodeFib
            SaveDefault "FibColor"
            SaveDefault "FibStyle"
            SaveDefault "FibRatios"
            SaveDefault "ShowValues"
            SaveDefault "ShowRatios"
            If m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 Or m.eType = eANNOT_Fibonacci3 Or _
               m.eType = eANNOT_Fibonacci4 Or m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_AdvRiskReward Or _
                m.eType = eANNOT_FibFan Or m.eType = eANNOT_DanCodeFib Then
                SaveDefault "Ext"
                SaveDefault "ExtColor"
                SaveDefault "ExtStyle"
                SaveDefault "ShowRatioZero"
                SaveDefault "ShowRatioOne"
                
                If m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_AdvRiskReward Then
                    If m.eType = eANNOT_AdvRiskReward Then
                        '11-09-2015: do not save ReverseDirection because on initial draw first click is always expected to be Entry
                        SaveDefault "FibTextPL"
                        SaveDefault "FibLabelPL"
                        SaveDefault "FibTextPLLoc"
                        SaveDefault "FibLabelPLLoc"
                        SaveDefault "FibRatiosText"
                        SaveDefault "textZero"
                        SaveDefault "textOne"
                    Else
                        SaveDefault "ReverseDirection"
                    End If
                End If
                If m.eType <> eANNOT_FibFan Then
                    SaveDefault "TextNextToMain"
                    SaveDefault "HideVerticalLine"
                    SaveDefault "FibBold"
                    SaveDefault "BoldRatioZero"
                    SaveDefault "BoldRatioOne"
                    SaveDefault "KeepAlive"
                    'SaveDefault "PriceOnly"        'JM 09-16-2015: deprecated
                    SaveDefault "FibTextRatio"
                    SaveDefault "FibTextPrice"
                    SaveDefault "FibTextRatioLoc"
                    SaveDefault "FibTextPriceLoc"
                End If
            ElseIf m.eType = eANNOT_FibTimeRatio Or m.eType = eANNOT_ElliotTimeRatio Then
                SaveDefault "ShowInAllPanes"
                SaveDefault "ShowDates"
                SaveDefault "ShowNumBars"
                If m.eType = eANNOT_FibTimeRatio Then SaveDefault "RatioOneVal"
            ElseIf m.eType = eANNOT_FibArcs Then
                SaveDefault "Ext"
                SaveDefault "DirNE"
                SaveDefault "DirNW"
                SaveDefault "DirSE"
                SaveDefault "DirSW"
                SaveDefault "Circular"
            End If
            
        Case eANNOT_FibTimeZones, eANNOT_TimeCycle, eANNOT_DanCodeZone
            SaveDefault "ShowInAllPanes"
            SaveDefault "Arcs"
            SaveDefault "ArcPercentHeight"
            SaveDefault "FromPrevBar"
            If m.eType <> eANNOT_TimeCycle Then
                SaveDefault "ShowValues"
                SaveDefault "ZoneInUse"
            End If
            
        Case eANNOT_SpResistFan
            SaveDefault "SpResistColor"
            SaveDefault "SpResistStyle"
            SaveDefault "SpResistRatios"
            SaveDefault "SpResistCount"
            
        Case eANNOT_TargetShooter
            SaveDefault "TargetColor"
            SaveDefault "TargetStyle"
            SaveDefault "Pullbacks"
            SaveDefault "Pt1"
            SaveDefault "Pt2"
            SaveDefault "Pt3"
            SaveDefault "ShowValues"
    
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
             eANNOT_DNExpansion4, eANNOT_FibABCD
            'pen colors
            SaveDefault "colorA"
            SaveDefault "colorB"
            SaveDefault "colorC"
            SaveDefault "colorD"
            SaveDefault "colorCOP"
            SaveDefault "colorOP"
            SaveDefault "colorXOP"
            'pen sizes
            SaveDefault "penSizeA"
            SaveDefault "penSizeB"
            SaveDefault "penSizeC"
            SaveDefault "penSizeD"
            SaveDefault "penSizeCOP"
            SaveDefault "penSizeOP"
            SaveDefault "penSizeXOP"
            'text
            SaveDefault "textA"
            SaveDefault "textB"
            SaveDefault "textC"
            SaveDefault "textD"
            SaveDefault "textCOP"
            SaveDefault "textOP"
            SaveDefault "textXOP"
            'ratios
            SaveDefault "ratioCOP"
            SaveDefault "ratioOP"
            SaveDefault "ratioXOP"
            SaveDefault "ShowHandle"
            
            If Not UseDiNapFib() Or m.eType = eANNOT_FibABCD Then
                SaveDefault "FibColor"
                SaveDefault "FibStyle"
                SaveDefault "FibRatios"
                SaveDefault "ShowValues"
                SaveDefault "ShowRatios"
                SaveDefault "FibRatiosText"
                SaveDefault "TextOnRight"
                SaveDefault "Ext"
                SaveDefault "ExtMain"
                SaveDefault "FreeFloat"
                SaveDefault "ShowValues"
            End If
        
        Case eANNOT_DNRetracement
            SaveDefault "FibR0"
            SaveDefault "FibR1"
            SaveDefault "FibR2"
            SaveDefault "FibColorR0"
            SaveDefault "FibColorR1"
            SaveDefault "FibColorR2"
            SaveDefault "FibPenSizeR0"
            SaveDefault "FibPensizeR1"
            SaveDefault "FibPensizeR2"
            SaveDefault "ArcColorR1"
            SaveDefault "ArcColorR2"
            SaveDefault "ArcPensizeR1"
            SaveDefault "ArcPensizeR2"
            SaveDefault "ShowHandle"
            
        Case eANNOT_GannLines
            SaveDefault "GannFan"      '0=angle, 1=fan
            SaveDefault "GannRatioType"    '0=points, 1=ticks
            SaveDefault "Rise"
            SaveDefault "Run"
            SaveDefault "DirNE"
            SaveDefault "DirNW"
            SaveDefault "DirSE"
            SaveDefault "DirSW"
            SaveDefault "1x1"
            SaveDefault "1x2"
            SaveDefault "1x3"
            SaveDefault "1x4"
            SaveDefault "1x8"
            SaveDefault "2x1"
            SaveDefault "3x1"
            SaveDefault "4x1"
            SaveDefault "8x1"
            SaveDefault "Color1x2"
            SaveDefault "Color1x3"
            SaveDefault "Color1x4"
            SaveDefault "Color1x8"
            
        Case eANNOT_AndrewFork
            SaveDefault "ForkMode"
            'time intervals
            SaveDefault "MarkerStyle"
            SaveDefault "MarkerRatios"
            SaveDefault "MarkerShow"
            SaveDefault "MarkerColor"
            SaveDefault "MarkerText"
            'parallels
            SaveDefault "ParallelStyle"
            SaveDefault "ParallelRatios"
            SaveDefault "ParallelShow"
            SaveDefault "ParallelColor"
            SaveDefault "ShowQtrLines"
            SaveDefault "ParallelText"
            'extend base of fork
            SaveDefault "ShowHandle"
            SaveDefault "HideABLine"
            SaveDefault "ShowInAllPanes"
            SaveDefault "ParallelLocation"
            
        Case eANNOT_WaveLabels
            SaveDefault "Lines"
            SaveDefault "LabelFirstPoint"
            SaveDefault "PointArc"
            SaveDefault "LabelPastEnd"
            SaveDefault "DefaultWaveLabel", m.strText
            
        Case eANNOT_RiskReward
            SaveDefault "ShowProfitLoss"
            SaveDefault "ShowValues"
            SaveDefault "NumContracts"
            SaveDefault "NumForexContracts"
            SaveDefault "NumShares"
            
        Case eANNOT_Pattern
            SaveDefault "ForecastBars"
            
        Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
            SaveDefault "ShowValues"        'aardvark 3828
            SaveDefault "TextAlignment"
            SaveDefault "Ext"
            SaveDefault "ShowInAxis"        'aardvark 6718
            SaveDefault "HideSRLineDot"
            
        Case eANNOT_Bracket
            SaveDefault "BracketStyle"
            
        Case eANNOT_Pivot
            SaveDefault "ShowValues"
            SaveDefault "TextNextToMain"
            SaveDefault "Ext"
            SaveDefault "ExtStyle"
            SaveDefault "CalcMethod"
            SaveDefault "PriceOnly"
            
            SaveDefault "PivotShow"
            SaveDefault "S1Show"
            SaveDefault "S2Show"
            SaveDefault "S3Show"
            SaveDefault "R1Show"
            SaveDefault "R2Show"
            SaveDefault "R3Show"
                        
            SaveDefault "PivotColor"
            SaveDefault "S1Color"
            SaveDefault "S2Color"
            SaveDefault "S3Color"
            SaveDefault "R1Color"
            SaveDefault "R2Color"
            SaveDefault "R3Color"
        
        Case eANNOT_Gartley             '6535
            SaveDefault "GartleyLabelStyle"
            SaveDefault "ShowHandle"
            
            SaveDefault "Triangle1"
            SaveDefault "Triangle2"
            SaveDefault "Triangle3"
            SaveDefault "Triangle4"
            
            SaveDefault "Triangle1_Color"
            SaveDefault "Triangle2_Color"
            SaveDefault "Triangle3_Color"
            SaveDefault "Triangle4_Color"
            
            SaveDefault "Triangle1_Fill"
            SaveDefault "Triangle2_Fill"
            SaveDefault "Triangle3_Fill"
            SaveDefault "Triangle4_Fill"

        Case eANNOT_BalloonStrangle
            SaveDefault "RiskBEStyle"
            SaveDefault "RiskBEColor"
            SaveDefault "ShowBE"
            SaveDefault "ShowPrices"
            SaveDefault "ShowTradeCost"
            SaveDefault "ShowProfitLoss"
            SaveDefault "ShowPriceMove"
            SaveDefault "ShowRatios"        '7206
            SaveDefault "BalloonColors"
            SaveDefault "BalloonRatios"
            SaveDefault "ShowText"

        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SaveDefaults", eGDRaiseError_Raise

End Sub

Public Function IsValid() As Boolean
On Error GoTo ErrSection:

    Dim bValid As Boolean

    bValid = True
    If m.eType = eANNOT_UndefinedType Then
        bValid = False
    ElseIf m.dDate1 = 0 Then
        
        If m.eType <> eANNOT_HorzLine And m.eType <> eANNOT_HorzLine2 And _
           m.eType <> eANNOT_HorzLine3 And m.eType <> eANNOT_HorzLine4 Then
            
            bValid = False
        End If
    
    End If
    IsValid = bValid

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.IsValid", eGDRaiseError_Raise

End Function

Public Sub TemplateSave(ByVal fh%)
On Error GoTo ErrSection:

    Dim i&, strText$
    
    'don't save trade triangles & rectangles for highlight bars
    If m.eType = eANNOT_SimpleLine Or m.eType = eANNOT_MultiRects Or _
       m.eUsage = eANNOT_AutoSwingTrend Or m.eUsage = eANNOT_AvgEntryLine Or _
       m.eUsage = eANNOT_BidAskChart Or m.eUsage = eANNOT_PriceAlert Or _
       m.eUsage = eANNOT_OptionInfo Or m.eUsage = eANNOT_PatternProfit Then
        Exit Sub
    End If
    
    If m.eType = eANNOT_RegressionLine Then
        If Len(Prop("IndicatorKey")) = 0 Then Exit Sub
    End If
        
    If fh <> 0 And IsValid() Then
        Print #fh, IniString("GraphEngDll", 1)
        
        If IsEndUserEWI Then
            'do this first so can distinguish end-user vs. analyst label on template load
            Print #fh, IniString("EndUserEWI", 1)
        End If
        
        Print #fh, IniString("Type", m.eType)
        Print #fh, IniString("Usage", m.eUsage)
        Print #fh, IniString("Pane", m.strPane)
        Print #fh, IniString("Date1", m.dDate1)
        Print #fh, IniString("Value1", m.Y1)
        Print #fh, IniString("Date2", m.dDate2)
        Print #fh, IniString("Value2", m.Y2)
        Print #fh, IniString("Color", m.nColor)
        Print #fh, IniString("Style", m.nStyle)
        Print #fh, IniString("Text", Replace(m.strText, vbCrLf, vbTab))
        Print #fh, IniString("PreIndicator", m.nPreIndicator)
        Print #fh, IniString("CreateDate", m.dCreateDate)
        Print #fh, IniString("CreateDateTime", DateFormat(m.dCreateDate, MM_DD_YYYY, HH_MM_SS))
        If Len(m.strBarPeriod) > 0 Then Print #fh, IniString("Periodicity", m.strBarPeriod)
        If Len(m.strPeriodsPerBar) > 0 Then Print #fh, IniString("PeriodsPerBar", m.strPeriodsPerBar)
        If Not m.oAlert Is Nothing Then m.strAlert = m.oAlert.ToFileString
        Print #fh, IniString("Alert", m.strAlert)
        
        'print flag to indicate annotation was saved after implementation of text options
        If m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or _
           m.eType = eANNOT_DollarLine3 Or m.eType = eANNOT_DollarLine4 Or _
           m.eType = eANNOT_GannacciSwingSquare Then
           Print #fh, IniString("TextOptions", 1)
        End If
        
        For i = 0 To m.Properties.Size - 1
            If InStr(m.Properties(i), "IsHidden") <> 0 Then
                If Not AllowEWI And Not AllowGMP Then
                    'don't save this property to template if they have flag files
                    Print #fh, m.Properties(i)
                End If
            ElseIf InStr(m.Properties(i), "EndUserEWI") <> 0 Then
                'this is written at top of file if = 1 else don't write it
            Else
                Print #fh, m.Properties(i)
            End If
        Next
        
        'for annotation having more than 2 date-y pairs (e.g. dinapoli annotations)
        For i = 0 To m.aDates.Size - 1
            strText = "Date" + Str(i + 3)
            Print #fh, IniString(strText, m.aDates(i))
        Next

        For i = 0 To m.aYs.Size - 1
            strText = "Value" + Str(i + 3)
            Print #fh, IniString(strText, m.aYs(i))
        Next
        
        Print #fh, "END="
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.TemplateSave", eGDRaiseError_Raise

End Sub

Public Sub TemplateLoad(ByVal fh%, Chart As cChart)
On Error GoTo ErrSection:

    Dim strProp$, strValue$, strMid$, dValue#, n&
    Dim bUsePresetFib As Boolean
    Dim bGannMultiColor As Boolean
    Dim nFillColor&
    Dim nDollarLineTextOptions&
    Dim bRatioZeroFound As Boolean
    Dim bForkFlagFound As Boolean
    Dim bShowValuesFound As Boolean
    
    bUsePresetFib = True
    nFillColor = -1
    If fh = 0 Then Exit Sub
    
    Set m.Chart = Chart
    
    Do While Not EOF(fh)
        Line Input #fh, strProp
        If ParseIniString(strProp, strValue, dValue) Then
            Select Case UCase(strProp)
            Case "END"
                Exit Do
            
            ' common properties
            Case "GRAPHENGDLL"
                m.nGraphEngDll = dValue
            Case "TYPE"
                m.eType = dValue
                If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 _
                    Or m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_DNRetracement Or m.eType = eANNOT_RegressionLine _
                    Or m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or m.eType = eANNOT_DollarLine3 _
                    Or m.eType = eANNOT_DollarLine4 Or m.eType = eANNOT_GannacciSwingSquare Or m.eType = eANNOT_AndrewFork _
                    Or m.eType = eANNOT_Pattern Or m.eType = eANNOT_TriangleWedge Or m.eType = eANNOT_ChannelHighlight _
                    Or m.eType = eANNOT_ElliotTimeRatio Or m.eType = eANNOT_FibABCD Or m.eType = eANNOT_Gartley _
                    Or m.eType = eANNOT_GannacciSwing1 Or m.eType = eANNOT_GannacciSwing2 Or m.eType = eANNOT_TrendChannel Then
                        m.geAnnStruct.moveable = 1                  '4304 - don't set this for text edit
                    
                        If m.eType = eANNOT_AndrewFork Then m.bSchiffFork = HasModule("ADVFIB")     'FileExist("SchiffFork.flg")
                    
                ElseIf m.eType = eANNOT_TextEdit Or m.eType = eANNOT_TextEdit2 Or _
                       m.eType = eANNOT_TextEdit3 Or m.eType = eANNOT_TextEdit4 Then
                    m.geAnnStruct.moveable = 2                  '4352
                ElseIf m.eType = eANNOT_ElliotLabel Then
                    If Not IsEndUserEWI Then
                        g.bPageHasEWILabels = True      '6926 - has elliot wave labels created by analysts
                    End If
                End If
            Case "USAGE"
                m.eUsage = dValue
                If m.eUsage = eANNOT_UndefinedUsage Then m.eUsage = eANNOT_UserAdded
            Case "PANE"
                m.strPane = strValue
            Case "DATE1"
                m.dDate1 = dValue
            Case "DATE2"
                m.dDate2 = dValue
            Case "VALUE1"
                m.Y1 = dValue
            Case "VALUE2"
                m.Y2 = dValue
            Case "COLOR"
                m.nColor = dValue
            Case "STYLE"
                m.nStyle = dValue
            Case "TEXT"
                If Left(strValue, 1) = "|" Then
                    m.strText = Mid(strValue, 3)
                Else
                    m.strText = strValue
                End If
                m.strText = Replace(m.strText, vbTab, vbCrLf)
            Case "TEXTALIGNMENT"
                m.geTextAlign = dValue
                Prop("TextAlignment") = dValue
            Case "PREINDICATOR"
                m.nPreIndicator = dValue
            Case "INDICATORID"
                m.geAnnStruct.indicatorId = dValue
            Case "FILLCOLOR"
                Prop("FillColor") = dValue
                nFillColor = dValue
            Case "FILLPATTERN"
                Prop("FillPattern") = dValue
            Case "TEXTOPTIONS"
                nDollarLineTextOptions = dValue
            Case "CREATEDATE"
                m.dCreateDate = dValue
            Case "PERIODICITY"
                m.strBarPeriod = strValue
            Case "PERIODSPERBAR"
                m.strPeriodsPerBar = strValue
            Case "ALERT"
                m.strAlert = strValue
            Case "CREATEDATETIME"
                'do nothing - this is just a string written at template save for quick reference, the double is what gets used
            
            Case "PARALLELTEXT", "MARKERTEXT"
                Prop(strProp) = strValue
                bForkFlagFound = True
            
            ' check for POINT1 and POINT2 for backward-compatibility
            Case "POINT1"
                n = Val(Parse(strValue, Chr(9), 1))
                'If n > 0 And n <> PEGAT_LINECONTINUE Then
                    m.nStyle = n
                'End If
                m.strPane = Trim(Parse(strValue, Chr(9), 2))
                m.dDate1 = Val(Parse(strValue, Chr(9), 3))
                m.Y1 = Val(Parse(strValue, Chr(9), 4))
                m.strText = Trim(Parse(strValue, Chr(9), 5))
            Case "POINT2"
                n = Val(Parse(strValue, Chr(9), 1))
                'If n > 0 And n <> PEGAT_LINECONTINUE Then
                    m.nStyle = n
                'End If
                m.strPane = Trim(Parse(strValue, Chr(9), 2))
                m.dDate2 = Val(Parse(strValue, Chr(9), 3))
                m.Y2 = Val(Parse(strValue, Chr(9), 4))
            
            ' other annotation properties & additional date-y data pairs
            Case Else
                strMid = Mid(strProp, 1, 4)
                If strMid = "Date" Then
                    m.aDates.Add dValue
                ElseIf strMid = "Valu" Then
                    m.aYs.Add dValue
                Else
                    If "LucasNumbers" = strProp Then
                        If Int(ValOfText(strValue)) = 0 Then      'convert old prop to new prop
                            Prop("ZoneInUse") = 1
                        Else
                            Prop("ZoneInUse") = 0
                        End If
                    Else
                        Prop(strProp) = strValue
                    End If
                    If "Color1x2" = strProp Then bGannMultiColor = True
                    If "ShowRatios" = strProp Or "SpResistRatios" = strProp Then
                        bUsePresetFib = False
                    End If
                    If "ShowRatioZero" = strProp Then bRatioZeroFound = True
                    If "ShowValues" = strProp Then bShowValuesFound = True
                End If
            End Select
        End If
    Loop
    
    n = 0
    If Not UseDiNapFib() Then
        If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or _
           m.eType = eANNOT_DNExpansion3 Or m.eType = eANNOT_DNExpansion4 Then
            n = True
        End If
    End If
    
    If m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 Or m.eType = eANNOT_Fibonacci3 _
        Or m.eType = eANNOT_Fibonacci4 Or m.eType = eANNOT_FibTimeRatio _
        Or m.eType = eANNOT_FibArcs Or m.eType = eANNOT_FibExpansion _
        Or m.eType = eANNOT_FibFan Or m.eType = eANNOT_SpResistFan _
        Or m.eType = eANNOT_ElliotTimeRatio Or m.eType = eANNOT_DanCodeFib _
        Or m.eType = eANNOT_AdvRiskReward Then
        
        LoadRatios bUsePresetFib, bUsePresetFib
        
        If Not bRatioZeroFound Then
            If m.eType = eANNOT_DanCodeFib Then
                Prop("ShowRatioZero") = 0       'JM 12-21-2015 fix for Jerry Winter issue reported by Heath
                Prop("ShowRatioOne") = 0
            Else
                Prop("ShowRatioZero") = 1
                Prop("ShowRatioOne") = 1
                If m.eType = eANNOT_ElliotTimeRatio Then Prop("ShowRatioTwo") = 1
            End If
        End If
        
        If m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 Or m.eType = eANNOT_Fibonacci3 _
            Or m.eType = eANNOT_Fibonacci4 Or m.eType = eANNOT_FibExpansion _
            Or m.eType = eANNOT_DanCodeFib Then
            
            If Len(Prop("FibTextRatioLoc")) = 0 Then
                If Len(Prop("PriceOnly")) > 0 Then
                    n = Val(Prop("PriceOnly"))
                Else
                    n = 0
                End If
                'JM 09-17-2015 - adjust new properties for backwards compatibility
                If n = 1 Then
                    Prop("FibTextPrice") = 1
                    Prop("FibTextRatio") = 0
                Else
                    'prior versions did not have option to show ratio only
                    Prop("FibTextPrice") = 1
                    Prop("FibTextRatio") = 1
                End If
                '09-17-2015: prior versions used ShowValues for both show & positioning
                '   0:no text
                '   1:ratio left, price/profit right
                '   2:all text on left
                '   3:all text on right,
                '   4:ratio right, price in y-scale
                n = Val(Prop("ShowValues"))
                Select Case n
                    Case 0:
                        Prop("FibTextPrice") = 0
                        Prop("FibTextRatio") = 0
                        Prop("FibTextPriceLoc") = 1
                        Prop("FibTextRatioLoc") = 0
                    Case 2:
                        Prop("FibTextPriceLoc") = 0
                        Prop("FibTextRatioLoc") = 0
                    Case 3:
                        Prop("FibTextPriceLoc") = 1
                        Prop("FibTextRatioLoc") = 1
                    Case 4:
                        Prop("FibTextPriceLoc") = 2
                        Prop("FibTextRatioLoc") = 1
                    Case Else
                        Prop("FibTextPriceLoc") = 1
                        Prop("FibTextRatioLoc") = 0
                End Select
            End If
        End If
    ElseIf m.eType = eANNOT_FibABCD Or n Then
        LoadRatios bUsePresetFib
        If Not bRatioZeroFound Then
            Prop("ShowRatioZero") = 1
            Prop("ShowRatioOne") = 1
            Prop("ShowRatioTwo") = 1
            If m.eType = eANNOT_FibABCD Then Prop("RatioD") = 1
        End If
        If Not bShowValuesFound Then Prop("ShowValues") = 1
    ElseIf m.eType = eANNOT_BalloonStrangle Then
        LoadRatios False
    ElseIf m.eType = eANNOT_GannLines Then
        If Val(Prop("SecondHotSpot")) = 0 Then
            m.dDate2 = m.dDate1
            m.Y2 = m.Y1
        End If
        If Not bGannMultiColor Then     'for backwards compatibility
            Prop("Color1x2") = m.nColor
            Prop("Color1x3") = m.nColor
            Prop("Color1x4") = m.nColor
            Prop("Color1x8") = m.nColor
        End If
    ElseIf m.eType = eANNOT_Rectangle And nFillColor = -1 Then
        Prop("FillColor") = m.nColor    'for backwards compatibility
        Prop("FillPattern") = 0
    ElseIf (m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or m.eType = eANNOT_DollarLine3 _
           Or m.eType = eANNOT_DollarLine4) And nDollarLineTextOptions = 0 Then
        Prop("ProfitLoss") = 1  'annotation was saved prior to implementation of text options, default to show trade profit/loss
    ElseIf m.eType = eANNOT_AndrewFork Then
        If Not bForkFlagFound Then
            Prop("MarkerText") = 1
            Prop("ParallelText") = 1
        End If
    End If
    'for backwards compatibility
    If m.dCreateDate = 0 Then m.dCreateDate = Now
    'do this outside loop because the graphengdll string may come after the style string
    geConvertPenStyle m.nStyle
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.TemplateLoad", eGDRaiseError_Raise

End Sub

Private Function GetX(Chart As cChart, ByVal dDate#) As Double
On Error GoTo ErrSection:

    Dim nBar&, i&, X#
    Dim dLastScreenDate#, dDateBar#
    
    If dDate = 0 Then
        X = kNullData
    ElseIf dDate < 0 Then
        ' negative means it was given as a screen X point
        X = Abs(dDate) + Chart.ScreenStartX
    Else
        If Not IsIntraday(Chart.Periodicity) Then
            ' need to check what session (e.g. if evening, then goes to next day)
            dDate = Chart.Bars.SessionDateForTime(dDate, False)
        ElseIf dDate <> Int(dDate) Then
            dDate = gdFixDateTime(dDate)    'aardvark 1192 fix
        End If
        ' see if can find exact date
        If Not Chart.aXdate.BinarySearch(dDate, nBar) Then
            If nBar = 0 Then
                dDateBar = Chart.Bars(eBARS_DateTime, 0)
                If Int(dDate) = Int(dDateBar) Then
                    If (dDate - Int(dDate)) <= (dDateBar - Int(dDateBar)) Then
                        X = 0               'aardvark 6672
                    Else
                        X = kNullData
                    End If
                Else
                    X = kNullData
                End If
            ' if not found, see if "out of range"
            ElseIf nBar < 0 Or nBar >= Chart.aXdate.Size Then
                X = kNullData
            End If
        End If
        If X <> kNullData Then
            X = nBar '- Chart.ScreenStartX
        Else
            'check for annot that was drawn into future on higher time frame (4112)
            dLastScreenDate = Chart.aXdate(m.Chart.aXdate.Size - 1)
            If dDate > dLastScreenDate Then
                X = Chart.aXdate.Size + (dDate - dLastScreenDate)
            End If
        End If
    End If

    GetX = X

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.GetX", eGDRaiseError_Raise

End Function

' Setup for all multi-point "Set..." routines
Private Function MultiPointSetup(Chart As cChart, XDelta#, YDelta#, dFudge#, _
        Optional ByVal bHasExtensions As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bShow As Boolean
    
    ' Get start and end points
    m.X1 = GetX(Chart, m.dDate1)
    m.X2 = GetX(Chart, m.dDate2)
    m.geAnnStruct.paneId = Chart.Tree.Index(m.strPane)
    XDelta = m.X2 - m.X1
    YDelta = m.Y2 - m.Y1
    
    ' See if should show or not
    bShow = True
    If m.geAnnStruct.paneId <= 0 Or m.X1 = kNullData Or m.X2 = kNullData _
        Or m.Y1 = kNullData Or m.Y2 = kNullData Then
            bShow = False
    'don't show if starts after end or finished before
    'beginning of visible screen
    ElseIf Not bHasExtensions And _
        ((m.X1 < Chart.ScreenStartX And m.X2 < Chart.ScreenStartX) _
        Or (m.X1 > Chart.ScreenEndX And m.X2 > Chart.ScreenEndX)) Then
            bShow = False
    'for multi-point annots with points on same bar:
    'only show if dates are exactly the same (otherwise,
    'came from a smaller time period and no use showing
    'since it's now just scrunched into a single bar)
    ElseIf m.X1 = m.X2 And m.dDate1 <> m.dDate2 Then
        If m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or m.eType = eANNOT_DollarLine3 Or _
           m.eType = eANNOT_DollarLine4 Or m.eType = eANNOT_GannacciSwingSquare Or _
           m.eType = eANNOT_Trendline Or m.eType = eANNOT_Trendline2 Or _
           m.eType = eANNOT_Trendline3 Or m.eType = eANNOT_Trendline4 Or _
           m.eType = eANNOT_TrendChannel Or m.eType = eANNOT_TargetShooter Then
            'these can be drawn vertically and should show in case it was drawn from
            'hi-to-low or vice versa on daily bar and time was appended to date (4212, 4217)
            If m.Y1 = m.Y2 Then bShow = False
        Else
            bShow = False
        End If
    End If

    MultiPointSetup = bShow

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MultiPointSetup", eGDRaiseError_Raise

End Function

' set point X,Y so just barely off the screen (either right or left)
Private Sub SetPointOffScreen(Chart As cChart, ByVal bRightSide As Boolean, X#, Y#)
On Error GoTo ErrSection:

    Dim dMult#

    If m.Y2 = m.Y1 Then
        ' if horizontal
        Y = m.Y1
        If bRightSide Then
            X = Chart.ScreenEndX + 2
        Else
            X = Chart.ScreenStartX - 2
        End If
    ElseIf m.X2 = m.X1 Then
        ' if vertical
        X = m.X1
    Else
        ' if slanted
        If bRightSide Then
            dMult = (Chart.ScreenEndX + 2 - m.X1) / (m.X2 - m.X1)
        Else
            dMult = (Chart.ScreenStartX - 2 - m.X1) / (m.X2 - m.X1)
        End If
        'MJM - doublecheck code
        X = m.X1 + (m.X2 - m.X1) * dMult
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetPointOffScreen", eGDRaiseError_Raise

End Sub

Private Function SetTrendLine(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim nColor&, nPenStyle&
    Dim iExt%, i&, k&
            
    Dim eUseLog As ePANE_LogFlag
    Dim Pane As cPane
    
    If m.Y1 > 0 And m.Y2 > 0 Then
        Set Pane = Chart.Tree(m.strPane)
        If Not Pane Is Nothing Then
            eUseLog = Pane.PaneLogFlag
            Set Pane = Nothing
        End If
    End If
    
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 8     'main line
        gdSetNum m.geAnnStruct.glhAnnType, 1, 8     'right extension line
        gdSetNum m.geAnnStruct.glhAnnType, 2, 8     'left extension line
        gdSetNum m.geAnnStruct.glhAnnType, 3, 0     'text
                
        'font (set for all since text is last item in array)
        'this is initial setting for place-holder purposes only
        'valid values for text are reset below
        For i = 0 To 2
            gdSetStr m.geAnnStruct.gshFont, i, "0"
            gdSetNum m.geAnnStruct.glhSize, i, 8
            gdSetNum m.geAnnStruct.glhStyle, i, 0
            gdSetNum m.geAnnStruct.glhUnderline, i, 0
            gdSetStr m.geAnnStruct.gshText, i, "0"      'place holder
            gdSetNum m.geAnnStruct.glhAlign, i, 4       'e_btmRight
        Next
        
        'backwards compatibility
        If m.eUsage <> eANNOT_AutoSwingTrend Then
            If Val(Prop("ChannelPoints")) = 0 Or Val(Prop("ChannelPercent")) = 0 Then
                Prop("ChannelCount") = 0
                Prop("ChannelPoints") = 3       'default
                Prop("ChannelPercent") = 1
'                Prop("ChannelStyle") = 0       'aardvark 6750
                Prop("ChannelType") = 0
                Prop("ChannelLocation") = 0
            End If
        End If
    End If
        
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    
    'font (reset for text only)
    gdSetStr m.geAnnStruct.gshFont, 3, Prop("FontName")
    gdSetNum m.geAnnStruct.glhSize, 3, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 3, Val(Prop("FontStyle"))
    gdSetNum m.geAnnStruct.glhUnderline, 3, Val(Prop("FontUnderline"))
    gdSetStr m.geAnnStruct.gshText, 3, m.strText
    
    'turn hittest off for auto-trendlines
    If m.eUsage = eANNOT_AutoSwingTrend Then
        m.geAnnStruct.skipHitTest = 1
    End If
    
    iExt = Val(Prop("Ext"))

    'assume no extension
    gdSetNum m.geAnnStruct.glhAnnType, 1, -1
    gdSetNum m.geAnnStruct.glhAnnType, 2, -1
   
    ' Extensions
    Select Case iExt
    Case 1 '(right)
        gdSetNum m.geAnnStruct.glhAnnType, 1, 8
    Case 2 '(left)
        gdSetNum m.geAnnStruct.glhAnnType, 2, 8
    Case 3 '(both)
        gdSetNum m.geAnnStruct.glhAnnType, 1, 8
        gdSetNum m.geAnnStruct.glhAnnType, 2, 8
    End Select
    
    'set pen color property
    nColor = Val(Prop("ExtColor"))
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor     'main line
    gdSetNum m.geAnnStruct.glhPenColor, 1, nColor       'right extension
    gdSetNum m.geAnnStruct.glhPenColor, 2, nColor       'left extension
    gdSetNum m.geAnnStruct.glhPenColor, 3, m.nColor     'text
    'set pen style property
    nPenStyle = Val(Prop("ExtStyle"))
    geSetPenStyle m.nStyle, 0   'main line
    geSetPenStyle nPenStyle, 1  'right extension
    geSetPenStyle nPenStyle, 2  'left extension
    gdSetNum m.geAnnStruct.glhPenSize, 3, 1             'text - always size 1
    gdSetNum m.geAnnStruct.glhPenStyle, 3, 0            'text - always solid
    'right extension
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y2
    gdSetNum m.geAnnStruct.gdhLeft, 1, dX2
    gdSetNum m.geAnnStruct.gdhRight, 1, Chart.geChartPoints + 1
    'left extension
    gdSetNum m.geAnnStruct.gdhBottom, 2, m.Y1
    gdSetNum m.geAnnStruct.gdhRight, 2, dX1
    gdSetNum m.geAnnStruct.gdhLeft, 2, -1
        
    If m.eUsage = eANNOT_AutoSwingTrend Then
        m.geAnnStruct.showAxes = 0
        Exit Function
    End If
    
    If Val(Prop("Y1OnHighLow")) = 1 Or Val(Prop("Y1OnOpenClose")) = 1 Then
        m.geAnnStruct.lenRatio = 1
    Else
        m.geAnnStruct.lenRatio = 0
    End If
    If Val(Prop("Y2OnHighLow")) = 1 Or Val(Prop("Y2OnOpenClose")) = 1 Then
        m.geAnnStruct.lenPoints = 1
    Else
        m.geAnnStruct.lenPoints = 0
    End If
    
    If m.eType = eANNOT_TrendChannel And m.geAnnStruct.moveable = 1 Then
        m.geAnnStruct.showQtrLines = Val(Prop("HideMainLine"))      '6717
    End If
    
    'check for text
    gdSetNum m.geAnnStruct.glhAnnType, 3, -1    'assume no text
    If Val(Prop("Ext")) = 0 Then
        If Len(m.strText) > 0 Then
            gdSetNum m.geAnnStruct.glhAnnType, 3, 0
            gdSetNum m.geAnnStruct.gdhTop, 3, m.Y2
            gdSetNum m.geAnnStruct.gdhLeft, 3, dX2
            gdSetNum m.geAnnStruct.gdhBottom, 3, m.Y2
            gdSetNum m.geAnnStruct.gdhRight, 3, dX2
            gdSetNum m.geAnnStruct.glhPenColor, 3, m.nColor
            gdSetStr m.geAnnStruct.gshText, 3, m.strText
            gdSetNum m.geAnnStruct.glhAlign, 3, m.geTextAlign       '6587
        End If
    End If
           
    Dim dChannel#, dPercenPoints#
    Dim nCount&, nLocation&, nStyle&
        
    'double check values
    nCount = Int(Val(Prop("ChannelCount")))
    nLocation = Int(Val(Prop("ChannelLocation")))   '1=above,2=below,3=both
    If nLocation = 0 Then
        gdSetSize m.geAnnStruct.glhAnnType, 4, 1
        gdSetSize m.geAnnStruct.glhPenColor, 4, 1
        gdSetSize m.geAnnStruct.glhPenStyle, 4, 1
        gdSetSize m.geAnnStruct.glhPenSize, 4, 1
        gdSetSize m.geAnnStruct.glhStyle, 4, 1
        gdSetSize m.geAnnStruct.glhSize, 4, 1
        gdSetSize m.geAnnStruct.gdhTop, 4, 1
        gdSetSize m.geAnnStruct.gdhLeft, 4, 1
        gdSetSize m.geAnnStruct.gdhBottom, 4, 1
        gdSetSize m.geAnnStruct.gdhRight, 4, 1
        nCount = 0
    ElseIf nLocation > 3 Then
        nLocation = 3
    End If
    If nLocation > 0 And nCount < 1 Then nCount = 1
            
    If nCount > 0 Then
    
        SetPercentPoints
            
        k = nCount * 2 + 4
        
        'set array sizes to match number of channels
        'this will correctly remove/add channels
        gdSetSize m.geAnnStruct.glhAnnType, k, 1
        gdSetSize m.geAnnStruct.glhPenColor, k, 1
        gdSetSize m.geAnnStruct.glhPenStyle, k, 1
        gdSetSize m.geAnnStruct.glhPenSize, k, 1
        gdSetSize m.geAnnStruct.glhStyle, k, 1
        gdSetSize m.geAnnStruct.glhSize, k, 1
        gdSetSize m.geAnnStruct.gdhTop, k, 1
        gdSetSize m.geAnnStruct.gdhLeft, k, 1
        gdSetSize m.geAnnStruct.gdhBottom, k, 1
        gdSetSize m.geAnnStruct.gdhRight, k, 1
                        
        If eUseLog = ePANE_LogFlagLog Then
            If Val(Prop("ChannelType")) = 1 Then
                dPercenPoints = Val(Prop("ChannelPoints"))
                If dX1 > dX2 Then
                    dChannel = Log(m.Y1 + dPercenPoints) - Log(m.Y1)
                Else
                    dChannel = Log(m.Y2 + dPercenPoints) - Log(m.Y2)
                End If
            Else
                dPercenPoints = Val(Prop("ChannelPercent")) / 100
                If dX1 > dX2 Then
                    dChannel = Log(m.Y1 + m.Y1 * dPercenPoints) - Log(m.Y1)
                Else
                    dChannel = Log(m.Y2 + m.Y2 * dPercenPoints) - Log(m.Y2)
                End If
            End If
        Else
            If Val(Prop("ChannelType")) = 1 Then
                dPercenPoints = Val(Prop("ChannelPoints"))
                dChannel = dPercenPoints
            Else
                dPercenPoints = Val(Prop("ChannelPercent")) / 100
                If dX1 > dX2 Then
                    dChannel = m.Y1 * dPercenPoints
                Else
                    dChannel = m.Y2 * dPercenPoints
                End If
                dChannel = Abs(dChannel)
            End If
        End If
        
        nCount = 1
        nStyle = Val(Prop("ChannelStyle"))
        For i = 4 To k - 1 Step 2
        
            gdSetNum m.geAnnStruct.glhAnnType, i, 8
            gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
            geSetPenStyle nStyle, i
            
            If eUseLog = ePANE_LogFlagLog Then
                gdSetNum m.geAnnStruct.gdhTop, i, Exp(Log(m.Y1) + dChannel * nCount)
                gdSetNum m.geAnnStruct.gdhLeft, i, dX1
                gdSetNum m.geAnnStruct.gdhBottom, i, Exp(Log(m.Y2) + dChannel * nCount)
                gdSetNum m.geAnnStruct.gdhRight, i, dX2
            Else
                gdSetNum m.geAnnStruct.gdhTop, i, m.Y1 + dChannel * nCount
                gdSetNum m.geAnnStruct.gdhLeft, i, dX1
                gdSetNum m.geAnnStruct.gdhBottom, i, m.Y2 + dChannel * nCount
                gdSetNum m.geAnnStruct.gdhRight, i, dX2
            End If
    
            gdSetNum m.geAnnStruct.glhAnnType, i + 1, 8
            gdSetNum m.geAnnStruct.glhPenColor, i + 1, m.nColor
            geSetPenStyle nStyle, i + 1
            
            If eUseLog = ePANE_LogFlagLog Then
                gdSetNum m.geAnnStruct.gdhTop, i + 1, Exp(Log(m.Y1) - dChannel * nCount)
                gdSetNum m.geAnnStruct.gdhLeft, i + 1, dX1
                gdSetNum m.geAnnStruct.gdhBottom, i + 1, Exp(Log(m.Y2) - dChannel * nCount)
                gdSetNum m.geAnnStruct.gdhRight, i + 1, dX2
            Else
                gdSetNum m.geAnnStruct.gdhTop, i + 1, m.Y1 - dChannel * nCount
                gdSetNum m.geAnnStruct.gdhLeft, i + 1, dX1
                gdSetNum m.geAnnStruct.gdhBottom, i + 1, m.Y2 - dChannel * nCount
                gdSetNum m.geAnnStruct.gdhRight, i + 1, dX2
            End If
            
            'turn off channels above/below main line as necessary
            If nLocation = 2 Then
                gdSetNum m.geAnnStruct.glhAnnType, i, -1
            ElseIf nLocation = 1 Then
                gdSetNum m.geAnnStruct.glhAnnType, i + 1, -1
            End If
                    
            nCount = nCount + 1
        Next
    End If
        
    'JM 08-10-2015: mod to show value in y-axis:
    '-1=label channels only, 0=no labels, 1=label main line only, >1=label main line & channels
    i = Val(Prop("ShowInAxis"))
    If i > 1 Then
        m.geAnnStruct.showAxes = gdGetSize(m.geAnnStruct.glhAnnType)
    Else
        m.geAnnStruct.showAxes = i
    End If
    
    'check for alert
    If Not m.oAlert Is Nothing Then
        SetAlertIcon gdGetSize(m.geAnnStruct.glhAnnType), m.Y2
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetTrendLine", eGDRaiseError_Raise

End Function

Private Function SetDollarLine(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim X#, Y#, XDelta#, YDelta#
    Dim dAmt#, dAmtPercent#, dFudge#, d1#, d2#
    Dim dTextX#, dTextY#
    Dim strAmt$, strPriceMove$, strBars$, s$
    Dim i&, nLen&, h1&, h2&
    
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 1     'main line
        gdSetNum m.geAnnStruct.glhAnnType, 1, 0     'text
        gdSetNum m.geAnnStruct.glhAnnType, 2, 1     'length as horizontal line
        
        'this is initial setting for place-holder purposes only
        'valid values for text are reset below
        For i = 0 To 2
            gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
            gdSetNum m.geAnnStruct.glhPenStyle, i, eANNOT_Thin
            gdSetStr m.geAnnStruct.gshFont, i, "0"
            gdSetNum m.geAnnStruct.glhSize, i, 8
            gdSetNum m.geAnnStruct.glhStyle, i, 0
            gdSetNum m.geAnnStruct.glhUnderline, i, 0
            gdSetStr m.geAnnStruct.gshText, i, ""
            'see note in grapheng.dll project's ExpFunctions.h pertaining to aardvark 1366
            gdSetNum m.geAnnStruct.glhAlign, i, -1
            gdSetNum m.geAnnStruct.gdhMisc, i, 0
        Next
        
        m.geAnnStruct.handleShow = 1        'aardvark 4498 (set to 1 to prevent grapheng.dll from drawing extension
    End If
    
    gdSetNum m.geAnnStruct.glhAnnType, 0, 1
    gdSetNum m.geAnnStruct.glhAnnType, 1, 0
    'check for keep at end option
    If Prop("KeepAtEnd") > 0 And m.geAnnStruct.moveable = 1 Then
        SetDynamicEndPoints Chart, dX1, dX2
    End If
    
    ' setup
    If MultiPointSetup(Chart, XDelta, YDelta, dFudge) = False Then
        SetDollarLine = 0
        'check if annotation's dates are within range of loaded data
        gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
        gdSetNum m.geAnnStruct.glhAnnType, 2, -1
        
        If m.eType = eANNOT_GannacciSwingSquare And Prop("SquareRange") <> 0 Then
            'continue - square range vertical lines may be within scop
        Else
            Exit Function
        End If
    End If
           
    strAmt = ""
    
    If m.dDate1 <= m.dDate2 Then ' TLB #5362
        dAmt = GetProfitLoss(Chart, m.Y1, m.Y2, strPriceMove)
    Else
        dAmt = GetProfitLoss(Chart, m.Y2, m.Y1, strPriceMove)
    End If
    
    If strPriceMove <> "INVALID" Then
        If m.X1 > m.X2 Then
            'dAmt = -dAmt
            If m.Y2 > 0 Then
                dAmtPercent = (m.Y1 / m.Y2 - 1) * 100
            Else
                dAmtPercent = 0
            End If
        ElseIf m.Y1 > 0 Then
            dAmtPercent = (m.Y2 / m.Y1 - 1) * 100
        Else
            dAmtPercent = 0
        End If
        
        If Val(Prop("ProfitLoss")) = 1 Then
            If dAmt = Int(dAmt) Then
                strAmt = Format(dAmt, "$#,##0;($#,##0)")
            Else
                strAmt = Format(dAmt, "$#,##0.00;($#,##0.00)")
            End If
        End If
        If Val(Prop("ProfitLossPercent")) = 1 Then
            If Len(strAmt) > 0 Then strAmt = strAmt & " / "
            strAmt = strAmt & Format(dAmtPercent, "#0.00") & "%"
        End If
    End If
    
    If m.X1 > Chart.ScreenEndX Then
        SetPointOffScreen Chart, True, X, Y
    ElseIf m.X1 < Chart.ScreenStartX Then
        SetPointOffScreen Chart, False, X, Y
    Else
        X = m.X1
        Y = m.Y1
    End If
                
    If m.X2 > Chart.ScreenEndX Then
        SetPointOffScreen Chart, True, X, Y
    ElseIf m.X2 < Chart.ScreenStartX Then
        SetPointOffScreen Chart, False, X, Y
    Else
        X = m.X2
        Y = m.Y2
    End If
    If X = m.X1 Then X = X + kXfudge
    If Y = m.Y1 Then Y = Y + dFudge
    
    'font and pen colors/styles
    For i = 0 To 2
        gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
        gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
    Next
    geSetPenStyle m.nStyle, 0       'main line
    geSetPenStyle eANNOT_DashSm, 2  'horizontal line
        
    'main line
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    
    'length as horizontal line
    If Val(Prop("LenHzLine")) = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 2, 1
        gdSetNum m.geAnnStruct.gdhLeft, 2, dX1
        gdSetNum m.geAnnStruct.gdhRight, 2, dX2
        If dX1 < dX2 Then
            gdSetNum m.geAnnStruct.gdhBottom, 2, m.Y1
            gdSetNum m.geAnnStruct.gdhTop, 2, m.Y1
        Else
            gdSetNum m.geAnnStruct.gdhBottom, 2, m.Y2
            gdSetNum m.geAnnStruct.gdhTop, 2, m.Y2
        End If
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 2, -1
    End If

    'set x-coordinate info for text
    If dX2 > dX1 Then
        dTextX = dX2
        dTextY = m.Y2
    Else
        dTextX = dX1
        dTextY = m.Y1
    End If

    m.strText = ""
    If strPriceMove <> "INVALID" Then
        'text (price move)
        If m.eType = eANNOT_GannacciSwingSquare Then
            If Val(Prop("PriceMove")) = 1 Then
                If dAmt < 0 Then
                    strPriceMove = "  -" & strPriceMove & " points"
                Else
                    strPriceMove = "  " & strPriceMove & " points"
                End If
                m.strText = strPriceMove
            End If
            SetGannacciSquareRange Chart, Abs(dAmt), dX1, dX2
        Else
            If Val(Prop("PriceMove")) = 1 Then
                ' TLB: for Forex, use the term "pips" instead
                s = " ticks"
                If Not Chart Is Nothing Then
                    If Not Chart.Bars Is Nothing Then
                        If IsForex(Chart.Bars.Prop(eBARS_Symbol)) Then
                            s = " pips"
                        End If
                    End If
                End If
                If dAmt < 0 Then
                    strPriceMove = "  -" & strPriceMove & s
                Else
                    strPriceMove = "  " & strPriceMove & s
                End If
                m.strText = strPriceMove
            End If
            If Val(Prop("PriceMovePoints")) = 1 Then
                d1 = Chart.Bars.Prop(eBARS_TickMove)
                If d1 > 0 Then
                    s = Chart.Bars.PriceDisplay(ValOfText(strPriceMove) * d1, True)
                    If Len(s) > 0 Then
                        s = s & " points"
                        If Len(m.strText) > 0 Then
                            m.strText = m.strText & vbCrLf & "  " & s
                        Else
                            m.strText = "  " & s
                        End If
                    End If
                End If
            End If
        End If
        
        'text (profit loss)
        If Val(Prop("ProfitLoss")) = 1 Or Val(Prop("ProfitLossPercent")) = 1 Then
            If Len(m.strText) > 0 Then
                m.strText = m.strText & vbCrLf & "  " & strAmt
            Else
                m.strText = "  " & strAmt
            End If
        End If
    End If
    
    'calendar days (same text location as profit loss)
    If Val(Prop("CalendarDays")) = 1 Then
        If Not Chart.Bars Is Nothing Then
            If Chart.Bars.Prop(eBARS_PeriodicityStr) = "Daily" Then
                If Prop("IncludeBarOne") = 1 Then
                    m.strText = m.strText & vbCrLf & "  CD = " & Str(Abs(Int(m.dDate2) - Int(m.dDate1)) + 1)
                Else
                    m.strText = m.strText & vbCrLf & "  CD = " & Str(Abs(Int(m.dDate2) - Int(m.dDate1)))
                End If
            End If
        End If
    End If
    
    'slope of line  - 6986
    If Val(Prop("SlopeLine")) = 1 Then
        If Int(dX2 - dX1) <> 0 Then
            ' TLB 5/30/2014: we really just need to do #ticks/bar for now (until Wayne can clarify what he really wants)
            dFudge = Chart.Bars.Prop(eBARS_TickMove)
            If dFudge > 0 Then
                dFudge = (m.Y2 - m.Y1) / dFudge / Int(dX2 - dX1)
                If IsForex(Chart.Bars.Prop(eBARS_Symbol)) Then
                    s = Format(dFudge, "#0.00") & " pips/bar"
                Else
                    s = Format(dFudge, "#0.00") & " ticks/bar"
                End If
            Else
                dFudge = (m.Y2 - m.Y1) / Int(dX2 - dX1)
                s = Str(Round(dFudge, 2)) & " / bar"
            End If
            If Len(m.strText) > 0 Then
                m.strText = m.strText & vbCrLf & "  " & s
            Else
                m.strText = "  " & s
            End If
        End If
    End If
    'text (length in bars)
    'JM 04-28-2011: There is an isssue that has existed prior to version 6.0 (minute charts)
    '   - draw dollar line with left end point on a bar that was visible prior to streaming, but right end point is in area of streamed data
    '   - turn of streaming, the length of the # bars of bars is now less than actual length when drawn
    '   - reason: the bars that existed with streaming on no longer exists & length is calculated on last good data bar
    '   - increasing the # of blank bars or moving chart to left corrects the length value
    If Val(Prop("LenBars")) = 1 Then
        If m.eType = eANNOT_GannacciSwingSquare Then
            If Chart.Bars Is Nothing Then
                If Prop("IncludeBarOne") = 1 Then
                    strBars = "  " & Str(Int(Abs(dX2 - dX1)) + 1) & " bars"
                Else
                    strBars = "  " & Str(Int(Abs(dX2 - dX1))) & " bars"
                End If
            ElseIf Chart.Bars.Prop(eBARS_PeriodicityStr) = "Daily" Then
                nLen = Abs(dX2 - dX1)
                'need to not count empty bars
                If m.Chart.ShowEmptyBars Then
                    h1 = m.Chart.aXdate.ArrayHandle
                    h2 = m.Chart.aXBar.ArrayHandle
                    
                    For i = m.X1 To m.X2
                        d1 = gdGetNum(h1, i)
                        d2 = gdGetNum(h2, i)
                        If d1 <> m.Chart.Bars(eBARS_DateTime, d2) Then
                            nLen = nLen - 1
                        End If
                    Next
                    
                End If
                If Prop("IncludeBarOne") = 1 Then nLen = nLen + 1
                strBars = "  TB = " & Str(nLen)
            ElseIf Prop("IncludeBarOne") = 1 Then
                strBars = "  TB = " & Str(Int(Abs(dX2 - dX1)) + 1)   'weekly, monthly etc
            Else
                strBars = "  TB = " & Str(Int(Abs(dX2 - dX1)))     'weekly, monthly etc
            End If
        Else
            strBars = "  " & Str(Int(Abs(dX2 - dX1))) & " bars"
        End If
        If Len(m.strText) > 0 Then
            m.strText = m.strText & vbCrLf & strBars
        Else
            m.strText = strBars
        End If
    End If
    
    If m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 _
       Or m.eType = eANNOT_DollarLine3 Or m.eType = eANNOT_DollarLine4 Then
        If Val(Prop("Volume")) = 1 And Not Chart.Bars Is Nothing Then
            h1 = Chart.Bars.FindDateTime(m.dDate1)
            h2 = Chart.Bars.FindDateTime(m.dDate2)
            If h1 >= 0 And h1 < Chart.Bars.Size Then
                If h2 >= 0 And h2 < Chart.Bars.Size Then
                    d1 = 0
                    If h1 > h2 Then
                        i = h2
                        h2 = h1
                        h1 = i
                    End If
                    For i = h1 To h2
                        If Chart.Bars(eBARS_Close, i) <> kNullData Then
                            d2 = Chart.Bars(eBARS_Vol, i)
                            d1 = d1 + d2
                        End If
                    Next
                    s = Format(d1, "#,##0")
                    m.strText = m.strText & vbCrLf & "  " & s
                End If
            End If
        End If
    End If
    
    If m.eType = eANNOT_GannacciSwingSquare And Prop("SameSwing") = 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
        gdSetNum m.geAnnStruct.glhAnnType, 2, -1
    ElseIf Len(m.strText) > 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 1, 0
        gdSetStr m.geAnnStruct.gshText, 1, m.strText
        gdSetNum m.geAnnStruct.gdhTop, 1, dTextY
        gdSetNum m.geAnnStruct.gdhLeft, 1, dTextX
        gdSetNum m.geAnnStruct.gdhBottom, 1, dTextY
        gdSetNum m.geAnnStruct.gdhRight, 1, dTextX
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
    End If
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetDollarLine", eGDRaiseError_Raise

End Function

Private Function SetRiskReward(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long

    Dim X#, Y#, XDelta#, YDelta#
    Dim dDiff#, dFudge#, dAmt#
    Dim dUserPoint1#, dUserPoint2#
    Dim strPriceMove$, strAmt$, i&
        
    Dim strFontName$, iFontSize&, iFontStyle&, iFontUnderline&
    
    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 1     'main line
        
        gdSetNum m.geAnnStruct.glhAnnType, 1, 6     'type ellipse
        gdSetNum m.geAnnStruct.glhStyle, 1, 0       'hollow circle
        gdSetNum m.geAnnStruct.glhFillPattern, 1, 0
        gdSetNum m.geAnnStruct.glhDirection, 1, 6   'e_west, shift left
    
        gdSetNum m.geAnnStruct.glhAnnType, 2, 6     'type ellipse
        gdSetNum m.geAnnStruct.glhStyle, 2, 0       'hollow circle
        gdSetNum m.geAnnStruct.glhFillPattern, 2, 0
        gdSetNum m.geAnnStruct.glhDirection, 2, 2   'e_east, shift right
    
        gdSetNum m.geAnnStruct.glhAnnType, 3, 1     'lower extension line
        gdSetNum m.geAnnStruct.glhAnnType, 4, 1     'lower extension line
        
        gdSetNum m.geAnnStruct.glhAnnType, 5, -1    'horizontal marks at end points
        gdSetNum m.geAnnStruct.glhAnnType, 6, -1
        
        gdSetNum m.geAnnStruct.glhAnnType, 7, 0     'text for profit/loss
        gdSetNum m.geAnnStruct.glhAnnType, 8, 0
        gdSetNum m.geAnnStruct.glhAnnType, 9, 0
        
        gdSetNum m.geAnnStruct.glhAnnType, 10, -1   'text for values
        gdSetNum m.geAnnStruct.glhAnnType, 11, -1   'default is show profit/loss only
        gdSetNum m.geAnnStruct.glhAnnType, 12, -1   'these are placeholders
        gdSetNum m.geAnnStruct.glhAnnType, 13, -1
        
        'place holders
        For i = 0 To 13
            gdSetStr m.geAnnStruct.gshFont, i, " "
            gdSetNum m.geAnnStruct.glhSize, i, 0
            gdSetStr m.geAnnStruct.gshText, i, " "
            gdSetNum m.geAnnStruct.glhAlign, i, 6   'e_ctrLeft
            gdSetNum m.geAnnStruct.gdhMisc, i, 0            'aardvark 4206

        Next
    End If
    
    If MultiPointSetup(Chart, XDelta, YDelta, dFudge) = False Then
        SetRiskReward = 0
        For i = 0 To 13
            gdSetNum m.geAnnStruct.glhAnnType, i, -1
        Next
        Exit Function
    End If
    
    If Val(Prop("UserEndPoint1")) > 0 And Val(Prop("UserEndPoint2")) > 0 Then
        'fix for aardvark 2159
        Dim aOrder As New cGdArray
        
        aOrder.Add m.Y1
        aOrder.Add m.Y2
        aOrder.Add Val(Prop("UserEndPoint1"))
        aOrder.Add Val(Prop("UserEndPoint2"))
        aOrder.Sort
        Prop("UserEndPoint1") = aOrder(0)
        m.Y1 = aOrder(1)
        m.Y2 = aOrder(2)
        Prop("UserEndPoint2") = aOrder(3)
    ElseIf m.Y1 > m.Y2 Then
        'make Y1 always less than Y2
        dFudge = m.Y1
        m.Y1 = m.Y2
        m.Y2 = dFudge
    End If

    'main line
    gdSetNum m.geAnnStruct.glhAnnType, 0, 1
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2

    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
    'circles
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhRight, 1, dX2
    
    gdSetNum m.geAnnStruct.gdhTop, 2, m.Y2
    gdSetNum m.geAnnStruct.gdhLeft, 2, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 2, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 2, dX2
    
    gdSetNum m.geAnnStruct.glhPenStyle, 1, 0    'always solid
    gdSetNum m.geAnnStruct.glhPenStyle, 2, 0    'always solid
    
    'extensions
    dUserPoint1 = Val(Prop("UserEndPoint1"))
    dUserPoint2 = Val(Prop("UserEndPoint2"))
    dDiff = Abs(m.Y1 - m.Y2)
    
    gdSetNum m.geAnnStruct.glhAnnType, 3, 1
    gdSetNum m.geAnnStruct.gdhLeft, 3, dX1
    gdSetNum m.geAnnStruct.gdhRight, 3, dX2
    
    gdSetNum m.geAnnStruct.glhAnnType, 4, 1
    gdSetNum m.geAnnStruct.gdhLeft, 4, dX1
    gdSetNum m.geAnnStruct.gdhRight, 4, dX2
    
    gdSetNum m.geAnnStruct.glhPenColor, 3, m.nColor
    geSetPenStyle m.nStyle, 3
    gdSetNum m.geAnnStruct.glhPenColor, 4, m.nColor
    geSetPenStyle m.nStyle, 4
    
    If dUserPoint1 <= 0 And dUserPoint2 <= 0 Then
        gdSetNum m.geAnnStruct.gdhTop, 3, m.Y1 - dDiff
        gdSetNum m.geAnnStruct.gdhBottom, 3, m.Y1
        gdSetNum m.geAnnStruct.gdhTop, 4, m.Y2
        gdSetNum m.geAnnStruct.gdhBottom, 4, m.Y2 + dDiff
        dUserPoint1 = m.Y1 - dDiff
        dUserPoint2 = m.Y2 + dDiff
    ElseIf dUserPoint1 <= 0 Then
        If dUserPoint2 > m.Y2 Then
            gdSetNum m.geAnnStruct.gdhTop, 3, m.Y1 - dDiff
            gdSetNum m.geAnnStruct.gdhBottom, 3, m.Y1
            gdSetNum m.geAnnStruct.gdhTop, 4, m.Y2
            gdSetNum m.geAnnStruct.gdhBottom, 4, dUserPoint2
            dUserPoint1 = m.Y1 - dDiff
        ElseIf dUserPoint1 < m.Y1 Then
            gdSetNum m.geAnnStruct.gdhTop, 3, dUserPoint2
            gdSetNum m.geAnnStruct.gdhBottom, 3, m.Y1
            gdSetNum m.geAnnStruct.gdhTop, 4, m.Y2
            gdSetNum m.geAnnStruct.gdhBottom, 4, m.Y2 + dDiff
            dUserPoint1 = m.Y2 + dDiff
        Else
            'should never happen ... something is very wrong here (don't draw extension)
            gdSetNum m.geAnnStruct.glhAnnType, 3, -1
            gdSetNum m.geAnnStruct.glhAnnType, 4, -1
        End If
    ElseIf dUserPoint2 <= 0 Then
        If dUserPoint1 > m.Y2 Then
            gdSetNum m.geAnnStruct.gdhTop, 3, m.Y1 - dDiff
            gdSetNum m.geAnnStruct.gdhBottom, 3, m.Y1
            gdSetNum m.geAnnStruct.gdhTop, 4, m.Y2
            gdSetNum m.geAnnStruct.gdhBottom, 4, dUserPoint1
            dUserPoint2 = m.Y1 - dDiff
        ElseIf dUserPoint1 < m.Y1 Then
            gdSetNum m.geAnnStruct.gdhTop, 3, dUserPoint1
            gdSetNum m.geAnnStruct.gdhBottom, 3, m.Y1
            gdSetNum m.geAnnStruct.gdhTop, 4, m.Y2
            gdSetNum m.geAnnStruct.gdhBottom, 4, m.Y2 + dDiff
            dUserPoint2 = m.Y2 + dDiff
        Else
            'should never happen ... something is very wrong here (don't draw extension)
            gdSetNum m.geAnnStruct.glhAnnType, 3, -1
            gdSetNum m.geAnnStruct.glhAnnType, 4, -1
        End If
    Else
        'make userpoint1 always less than userpoint2
        If dUserPoint1 > dUserPoint2 Then
            dFudge = dUserPoint1
            dUserPoint1 = dUserPoint2
            dUserPoint2 = dFudge
            Prop("UserEndPoint1") = dUserPoint1
            Prop("UserEndPoint2") = dUserPoint2
        End If
        gdSetNum m.geAnnStruct.gdhTop, 3, dUserPoint1
        gdSetNum m.geAnnStruct.gdhBottom, 3, m.Y1
        gdSetNum m.geAnnStruct.gdhTop, 4, m.Y2
        gdSetNum m.geAnnStruct.gdhBottom, 4, dUserPoint2
    End If
                
    'horizontal marks at user end points
    For i = 5 To 6
        gdSetNum m.geAnnStruct.glhAnnType, i, 1
        gdSetNum m.geAnnStruct.gdhLeft, i, dX1 - 1
        gdSetNum m.geAnnStruct.gdhRight, i, dX1 + 1
        gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
        geSetPenStyle m.nStyle, i
    Next
    
    gdSetNum m.geAnnStruct.gdhTop, 5, dUserPoint1
    gdSetNum m.geAnnStruct.gdhBottom, 5, dUserPoint1
    
    gdSetNum m.geAnnStruct.gdhTop, 6, dUserPoint2
    gdSetNum m.geAnnStruct.gdhBottom, 6, dUserPoint2
    
    strFontName = Prop("FontName")
    iFontSize = Val(Prop("FontSize"))
    iFontStyle = Val(Prop("FontStyle"))
    iFontUnderline = Val(Prop("FontUnderline"))
    
    'text
    If Val(Prop("ShowProfitLoss")) = 1 Then
        For i = 7 To 9
            gdSetNum m.geAnnStruct.glhAnnType, i, 0
            gdSetStr m.geAnnStruct.gshFont, i, strFontName
            gdSetNum m.geAnnStruct.glhSize, i, iFontSize
            gdSetNum m.geAnnStruct.glhStyle, i, iFontStyle          '5749
            gdSetNum m.geAnnStruct.glhUnderline, i, iFontUnderline
            gdSetNum m.geAnnStruct.gdhLeft, i, dX1
            gdSetNum m.geAnnStruct.gdhRight, i, dX1
            gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
            geSetPenStyle m.nStyle, i
        Next
        
        dAmt = GetProfitLoss(Chart, m.Y1, dUserPoint1, strPriceMove)
        If dAmt = Int(dAmt) Then                                        '5275
            strAmt = Format(dAmt, "$#,##0;($#,##0)")
        Else
            strAmt = Format(dAmt, "$#,##0.00;($#,##0.00)")
        End If
        gdSetStr m.geAnnStruct.gshText, 7, "  " & strAmt
        
        dAmt = GetProfitLoss(Chart, m.Y1, m.Y2, strPriceMove)
        If dAmt = Int(dAmt) Then
            strAmt = Format(dAmt, "$#,##0;($#,##0)")
        Else
            strAmt = Format(dAmt, "$#,##0.00;($#,##0.00)")
        End If
        gdSetStr m.geAnnStruct.gshText, 8, "  " & strAmt
        
        dAmt = GetProfitLoss(Chart, m.Y2, dUserPoint2, strPriceMove)
        If dAmt = Int(dAmt) Then
            strAmt = Format(dAmt, "$#,##0;($#,##0)")
        Else
            strAmt = Format(dAmt, "$#,##0.00;($#,##0.00)")
        End If
        gdSetStr m.geAnnStruct.gshText, 9, "  " & strAmt

        gdSetNum m.geAnnStruct.gdhTop, 7, (dUserPoint1 + m.Y1) / 2
        gdSetNum m.geAnnStruct.gdhTop, 8, (m.Y1 + m.Y2) / 2
        gdSetNum m.geAnnStruct.gdhTop, 9, (dUserPoint2 + m.Y2) / 2
    Else
        For i = 7 To 9
            gdSetNum m.geAnnStruct.glhAnnType, i, -1
        Next
    End If
    
    If Val(Prop("ShowValues")) = 1 Then
        For i = 10 To 13
            gdSetNum m.geAnnStruct.glhAnnType, i, 0
            gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
            gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
            If i = 11 Or i = 12 Then
                gdSetNum m.geAnnStruct.gdhLeft, i, dX1
                gdSetNum m.geAnnStruct.gdhRight, i, dX1
            Else
                gdSetNum m.geAnnStruct.gdhLeft, i, dX1 + 1
                gdSetNum m.geAnnStruct.gdhRight, i, dX1 + 1
            End If
            gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
            geSetPenStyle m.nStyle, i
        Next
        gdSetStr m.geAnnStruct.gshText, 10, "  " & Chart.PriceDisplay(m.geAnnStruct.paneId, dUserPoint1)
        gdSetStr m.geAnnStruct.gshText, 11, "  " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
        gdSetStr m.geAnnStruct.gshText, 12, "  " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y2)
        gdSetStr m.geAnnStruct.gshText, 13, "  " & Chart.PriceDisplay(m.geAnnStruct.paneId, dUserPoint2)

        gdSetNum m.geAnnStruct.gdhTop, 10, dUserPoint1
        gdSetNum m.geAnnStruct.gdhTop, 11, m.Y1
        gdSetNum m.geAnnStruct.gdhTop, 12, m.Y2
        gdSetNum m.geAnnStruct.gdhTop, 13, dUserPoint2
    Else
        For i = 10 To 13
            gdSetNum m.geAnnStruct.glhAnnType, i, -1
        Next
    End If
    
    'set initial user endpoints
    If m.geAnnStruct.moveable = 1 Then
        If Val(Prop("UserEndPoint1")) = -1 Then Prop("UserEndPoint1") = dUserPoint1
        If Val(Prop("UserEndPoint2")) = -1 Then Prop("UserEndPoint2") = dUserPoint2
    End If
    
    SetRiskReward = 1

End Function

Private Function SetSimpleLine(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:
        
    Dim X#, Y#, XDelta#, YDelta#, dFudge#
    
    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then gdSetNum m.geAnnStruct.glhAnnType, 0, 24     'line (no hittest)
    
    ' setup
    If MultiPointSetup(Chart, XDelta, YDelta, dFudge) = False Then
        SetSimpleLine = 0
        'check if annotation's dates are within range of loaded data
        'If m.dDate1 < Chart.Bars(eBARS_DateTime, 0) Or m.dDate2 < Chart.Bars(eBARS_DateTime, 0) Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        'End If
        Exit Function
    End If
       
    If m.eUsage = eANNOT_WhatIf Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 1
        gdSetStr m.geAnnStruct.gshText, 0, "?"
        gdSetNum m.geAnnStruct.glhAlign, 0, 6
        gdSetNum m.geAnnStruct.gdhMisc, 0, 0
        m.nStyle = eANNOT_DashSm
        If Not Chart.Tree Is Nothing Then
            If Not Chart.Tree("PRICE") Is Nothing Then
                m.nColor = Chart.Tree("PRICE").Color
            End If
        End If
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 0, 24
        m.nStyle = g.ChartGlobals.eProfitLineStyle
    End If

    If m.X1 > Chart.ScreenEndX Then
        SetPointOffScreen Chart, True, X, Y
    ElseIf m.X1 < Chart.ScreenStartX Then
        SetPointOffScreen Chart, False, X, Y
    Else
        X = m.X1
        Y = m.Y1
    End If
                
    If m.X2 > Chart.ScreenEndX Then
        SetPointOffScreen Chart, True, X, Y
    ElseIf m.X2 < Chart.ScreenStartX Then
        SetPointOffScreen Chart, False, X, Y
    Else
        X = m.X2
        Y = m.Y2
    End If
    If X = m.X1 Then X = X + kXfudge
    If Y = m.Y1 Then Y = Y + dFudge
    
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
    'line
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2

    SetSimpleLine = 1
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetSimpleLine", eGDRaiseError_Raise
    
End Function

Private Sub ShiftDnExpansion(ByVal nShift&)
On Error GoTo ErrSection:

    Dim dPoint#, i&
    
    For i = 0 To 3
        dPoint = gdGetNum(m.geAnnStruct.gdhLeft, i)
        If dPoint <> kNullData Then gdSetNum m.geAnnStruct.gdhLeft, i, dPoint + nShift
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ShiftDnExpansion", eGDRaiseError_Raise

End Sub

Private Sub ShiftDnRetracement(ByVal nShift&)
On Error GoTo ErrSection:

    Dim nSize&, idx&, i&
    Dim dPoint#

    dPoint = gdGetNum(m.geAnnStruct.gdhLeft, 0)
    If dPoint <> kNullData Then gdSetNum m.geAnnStruct.gdhLeft, 0, dPoint + nShift

    'set begining index of data for reaction point
    idx = 6 'hard-coded for now, can change if necessary
    
    nSize = m.aDates.Size
    For i = 0 To nSize - 1
        dPoint = gdGetNum(m.geAnnStruct.gdhLeft, i + idx)
        If dPoint <> kNullData Then gdSetNum m.geAnnStruct.gdhLeft, i + idx, dPoint + nShift
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ShiftDnRetracement", eGDRaiseError_Raise

End Sub

Private Sub ShiftDTLine(ByVal nShift&)
On Error GoTo ErrSection:
    
    Dim dPoint#
    
    dPoint = gdGetNum(m.geAnnStruct.gdhLeft, 0)
    If dPoint <> kNullData Then gdSetNum m.geAnnStruct.gdhLeft, 0, dPoint + nShift
    
    dPoint = gdGetNum(m.geAnnStruct.gdhRight, 0)
    If dPoint <> kNullData Then
        dPoint = dPoint + nShift
        gdSetNum m.geAnnStruct.gdhRight, 0, dPoint
        If m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or m.eType = eANNOT_DollarLine3 _
           Or m.eType = eANNOT_DollarLine4 Or m.eType = eANNOT_GannacciSwingSquare Then
            gdSetNum m.geAnnStruct.gdhLeft, 1, dPoint   '$Line text
        ElseIf m.eType <> eANNOT_SimpleLine Then
            gdSetNum m.geAnnStruct.gdhLeft, 3, dPoint   'trendline text
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ShiftDTLine", eGDRaiseError_Raise

End Sub

Private Sub ShiftAnnot(ByVal nShift&)
On Error GoTo ErrSection:
    
    Dim dPoint#, i&, j&
    
    j = gdGetSize(m.geAnnStruct.glhAnnType) - 1
    For i = 0 To j
        dPoint = gdGetNum(m.geAnnStruct.gdhLeft, i)
        If dPoint <> kNullData Then gdSetNum m.geAnnStruct.gdhLeft, i, dPoint + nShift
        dPoint = gdGetNum(m.geAnnStruct.gdhRight, i)
        If dPoint <> kNullData Then gdSetNum m.geAnnStruct.gdhRight, i, dPoint + nShift
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ShiftAnnot", eGDRaiseError_Raise

End Sub

Private Sub ShiftFibonacci(ByVal nShift&)
On Error GoTo ErrSection:

    Dim dPoint#, Size&, i&

    Size = gdGetSize(m.geAnnStruct.glhAnnType)
    For i = 0 To Size
        dPoint = gdGetNum(m.geAnnStruct.gdhLeft, i)
        If dPoint <> kNullData Then
            gdSetNum m.geAnnStruct.gdhLeft, i, dPoint + nShift
        End If
        dPoint = gdGetNum(m.geAnnStruct.gdhRight, i)
        If dPoint <> kNullData Then
            gdSetNum m.geAnnStruct.gdhRight, i, dPoint + nShift
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ShiftFibonacci", eGDRaiseError_Raise

End Sub

Private Sub ShiftRectangle(ByVal nShift&)
On Error GoTo ErrSection:

    Dim dPoint1#, dPoint2#
    
    dPoint1 = gdGetNum(m.geAnnStruct.gdhLeft, 0)
    If dPoint1 <> kNullData Then
        gdSetNum m.geAnnStruct.gdhLeft, 0, dPoint1 + nShift
        dPoint2 = gdGetNum(m.geAnnStruct.gdhRight, 0)
        If dPoint2 <> kNullData Then
            gdSetNum m.geAnnStruct.gdhRight, 0, dPoint2 + nShift
            gdSetNum m.geAnnStruct.gdhLeft, 1, (dPoint1 + dPoint2) / 2 'text
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ShiftRectangle", eGDRaiseError_Raise


End Sub

Private Sub ShiftTargetShooter(ByVal nShift&)
On Error GoTo ErrSection:

    Dim nSize&, i&
    Dim dPoint#

    For i = 0 To 17     '0-11 are lines, 12-17 are text
        dPoint = gdGetNum(m.geAnnStruct.gdhLeft, i)
        If dPoint <> kNullData Then gdSetNum m.geAnnStruct.gdhLeft, i, dPoint + nShift
        If i < 12 Then
            dPoint = gdGetNum(m.geAnnStruct.gdhRight, i)
            If dPoint <> kNullData Then gdSetNum m.geAnnStruct.gdhRight, i, dPoint + nShift
        End If
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ShiftTargetShooter", eGDRaiseError_Raise

End Sub

Private Sub ShiftTextEdit(ByVal nShift&)
On Error GoTo ErrSection:

    Dim dPoint#
    
    dPoint = gdGetNum(m.geAnnStruct.gdhLeft, 0)
    If dPoint <> kNullData Then
        gdSetNum m.geAnnStruct.gdhLeft, 0, dPoint + nShift
    End If

    dPoint = gdGetNum(m.geAnnStruct.gdhLeft, 1)
    If dPoint <> kNullData Then
        gdSetNum m.geAnnStruct.gdhLeft, 1, dPoint + nShift
    End If
    
    dPoint = gdGetNum(m.geAnnStruct.gdhRight, 1)
    If dPoint <> kNullData Then
        gdSetNum m.geAnnStruct.gdhRight, 1, dPoint + nShift
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ShiftTextEdit", eGDRaiseError_Raise

End Sub

Private Sub SetDynamicEndPoints(Chart As cChart, dX1#, dX2#)
On Error GoTo ErrSection:

    If m.eType <> eANNOT_DollarLine And m.eType <> eANNOT_DollarLine2 And m.eType <> eANNOT_DollarLine3 And _
       m.eType <> eANNOT_DollarLine4 And m.eType <> eANNOT_RegressionLine Then
       
       Exit Sub
    End If
    
    Dim dDate#, dClose#
    Dim nBarIdx&, nLength&
    
    If m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or _
       m.eType = eANNOT_DollarLine3 Or m.eType = eANNOT_DollarLine4 Then
        'get index of last good data bar
        nBarIdx = Chart.LastGoodDataBar(False, False)
    ElseIf m.eType = eANNOT_RegressionLine Then
        'get index of last good visible data bar
        nBarIdx = Chart.LastGoodDataBar(False, True)
    End If
    dDate = Chart.Bars(eBARS_DateTime, nBarIdx)
    dClose = Chart.Bars(eBARS_Close, nBarIdx)
    
    'dollar line
    If m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or _
       m.eType = eANNOT_DollarLine3 Or m.eType = eANNOT_DollarLine4 Then
        If m.X1 > m.X2 Then
            m.X1 = m.X2
            m.Y1 = m.Y2
            m.dDate1 = m.dDate2
        End If
        'set info for rightmost end point
        m.Y2 = dClose
        m.dDate2 = dDate
        m.X2 = GetX(Chart, m.dDate2)
        'set data bucket number of end points
        dX2 = GetX(Chart, m.dDate2) - Chart.ScreenStartX
        dX1 = GetX(Chart, m.dDate1) - Chart.ScreenStartX
        Exit Sub
    End If
    
    'regression line
    If Val(Prop("FixLength")) <> 0 Then       'aardvark 3150 fix
        'get length of line in x-direction
        If Val(Prop("LenInBars")) < 1 Then
            nLength = Abs(dX2 - dX1)
            Prop("LenInBars") = nLength
        Else
            nLength = Val(Prop("LenInBars"))
        End If
        'set info for leftmost end point
        nBarIdx = Abs(nBarIdx - nLength)
        m.dDate1 = Chart.Bars(eBARS_DateTime, nBarIdx)
        m.X1 = GetX(Chart, m.dDate1)
    End If
    'set info for rightmost end point
    m.dDate2 = dDate
    m.X2 = GetX(Chart, m.dDate2)
    'set data bucket number of end points
    dX2 = GetX(Chart, m.dDate2) - Chart.ScreenStartX
    dX1 = GetX(Chart, m.dDate1) - Chart.ScreenStartX

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetDynamicEndPoints", eGDRaiseError_Raise

End Sub

Private Function SetTriangleWedge(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&) As Long
On Error GoTo ErrSection:

    Dim dX3#, dY3#
    
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    If m.geAdded = False Or gdGetNum(m.geAnnStruct.glhAnnType, 0) < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 31
        gdSetNum m.geAnnStruct.glhAnnType, 1, 31
        gdSetNum m.geAnnStruct.glhAnnType, 2, 31
    End If

    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    gdSetNum m.geAnnStruct.glhFillColor, 0, Val(Prop("FillColor"))
    gdSetNum m.geAnnStruct.glhFillPattern, 0, Val(Prop("FillPattern"))  '0=transparent, 1=solid
    If m.aDates.Size > 0 Then
        dX3 = GetX(Chart, m.aDates(0)) - nStartScreenX          '4525
    Else
        dX3 = dX2
    End If
    If m.aYs.Size > 0 Then
        dY3 = m.aYs(0)
    Else
        dY3 = m.Y2
    End If
        
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 1, dX2
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y2
    gdSetNum m.geAnnStruct.gdhLeft, 2, dX3
    gdSetNum m.geAnnStruct.gdhTop, 2, dY3
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetTriangleWedge", eGDRaiseError_Raise

End Function

Private Function SetChannelHighlight(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&) As Long
On Error GoTo ErrSection:

    Dim dX3#, dY3#
    Dim dX4#, dY4#
    
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    If m.geAdded = False Or gdGetNum(m.geAnnStruct.glhAnnType, 0) < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 32
        gdSetNum m.geAnnStruct.glhAnnType, 1, 32
        gdSetNum m.geAnnStruct.glhAnnType, 2, 32
        gdSetNum m.geAnnStruct.glhAnnType, 3, 32
    End If

    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    gdSetNum m.geAnnStruct.glhFillColor, 0, Val(Prop("FillColor"))
    gdSetNum m.geAnnStruct.glhFillPattern, 0, Val(Prop("FillPattern"))  '0=transparent, 1=solid
    
    If m.aDates(0) < 0 Then
        dX3 = kNullData
    Else
        dX3 = GetX(Chart, m.aDates(0)) - nStartScreenX          '4525
    End If
    If m.aDates(1) < 0 Then
        dX4 = kNullData
    Else
        dX4 = GetX(Chart, m.aDates(1)) - nStartScreenX          '4525
    End If
    dY3 = m.aYs(0)
    dY4 = m.aYs(1)
        
    If Val(Prop("LenRatio")) = -1 And m.geAnnStruct.moveable = 1 Then
        If Abs(dX2 - dX1) > Abs(dX2 - dX3) Then
            Prop("LenRatio") = 1
        Else
            Prop("LenRatio") = 0
        End If
    End If
    m.geAnnStruct.lenRatio = Val(Prop("LenRatio"))
        
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 1, dX2
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y2
    gdSetNum m.geAnnStruct.gdhLeft, 2, dX3
    gdSetNum m.geAnnStruct.gdhTop, 2, dY3
    gdSetNum m.geAnnStruct.gdhLeft, 3, dX4
    gdSetNum m.geAnnStruct.gdhTop, 3, dY4
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetChannelHighlight", eGDRaiseError_Raise

End Function
Private Function SetETRatio(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&) As Long
On Error GoTo ErrSection:

    Dim dX3#, dY3#, dxLeft#, dTextY#, i&, j&
    Dim nPropColor&, nStyle&, nShowValues&, iExt&, iColor&
    Dim nX&, nBars&, iShowDates&, iShowNumBars&
    
    Dim aSort As New cGdArray
        
    'array item setup info
    '[0]    set to be ignored by graphics engine
    '[1]    color & style for drawing in all panes
    '[2]    line to be drawn at point of 1st mouse click (START line always)
    '[3]    line to be drawn at  point of 2nd mouse click (END line always)
    '[4]    line to be drawn at  point of 3rd mouse click (EXTEND line always)
    
    nPropColor = Val(Prop("FibColor"))
    nStyle = Val(Prop("ExtStyle"))
    nShowValues = Val(Prop("ShowValues"))
    iExt = Val(Prop("ShowInAllPanes"))
    iShowDates = Val(Prop("ShowDates"))
    iShowNumBars = Val(Prop("ShowNumBars"))
        
    m.geAnnStruct.showAxes = iExt
    gdSetNum m.geAnnStruct.glhAnnType, 0, -1
    'pen info for drawing across all panes
    gdSetNum m.geAnnStruct.glhAnnType, 1, 26
    gdSetNum m.geAnnStruct.glhPenColor, 1, -1
    geSetPenStyle nStyle, 1
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
    
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 2)
    If m.geAdded = False Or i <= 0 Then
        For i = 0 To 4
            gdSetStr m.geAnnStruct.gshText, i, ""   'place holder
            gdSetNum m.geAnnStruct.glhAlign, i, 5   'center the X
        Next
    End If
    
    'set annotation type
    If Prop("ShowRatioZero") = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 2, 1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 2, -1
    End If
    If Prop("ShowRatioOne") = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 3, 1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 3, -1
    End If
    If Prop("ShowRatioTwo") = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 4, 1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 4, -1
    End If
    
    gdSetNum m.geAnnStruct.glhPenColor, 2, m.nColor
    geSetPenStyle m.nStyle, 2
    gdSetNum m.geAnnStruct.glhPenColor, 3, m.nColor
    geSetPenStyle m.nStyle, 3
    gdSetNum m.geAnnStruct.glhPenColor, 4, m.nColor
    geSetPenStyle m.nStyle, 4
                    
    If m.aYs.Size > 0 Then
        aSort.Size = 0
        'find largest/smallest Ys
        dY3 = m.aYs(0)
        aSort.Add m.Y1
        aSort.Add m.Y2
        aSort.Add dY3
        aSort.Sort eGdSort_Default
        m.Y1 = aSort(0)
        m.aYs(0) = aSort(1)
        m.Y2 = aSort(2)
    End If
                    
    gdSetNum m.geAnnStruct.gdhLeft, 2, dX1
    gdSetNum m.geAnnStruct.gdhRight, 2, dX1
    
    gdSetNum m.geAnnStruct.gdhLeft, 3, dX2
    gdSetNum m.geAnnStruct.gdhRight, 3, dX2
    
    If m.aDates.Size > 0 Then
        dX3 = GetX(Chart, m.aDates(0)) - nStartScreenX          '4525
        
'JM (07-15-2011): Commented out to fix issue 6381. Leave awhile then remove if all ok.
        'gdSetNum m.geAnnStruct.glhAnnType, 4, 1
        'gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        
        gdSetNum m.geAnnStruct.gdhLeft, 4, dX3
        gdSetNum m.geAnnStruct.gdhRight, 4, dX3
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 4, -1

        gdSetStr m.geAnnStruct.gshText, 2, "x"
        gdSetStr m.geAnnStruct.gshText, 3, "x"
        
'JM (07-15-2011): Commented out to fix issue 6381. Leave awhile then remove if all ok.
'        gdSetNum m.geAnnStruct.glhAnnType, 0, 0
'        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
'        gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
'        gdSetStr m.geAnnStruct.gshText, 0, "x"
    End If
    
    For i = 2 To 4
        gdSetNum m.geAnnStruct.gdhTop, i, m.Y1
        gdSetNum m.geAnnStruct.gdhBottom, i, m.Y2

        gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
    Next
    
    Set aSort = Nothing
    
    If m.aDates.Size = 0 Then Exit Function
    
    nStyle = Val(Prop("FibStyle"))
    nBars = Abs(dX1 - dX2)
                
    'remove all text for # of bars
    j = m.aRatios.Size + 5
    gdSetSize m.geAnnStruct.glhAnnType, j, 1
    gdSetSize m.geAnnStruct.glhPenColor, j, 1
    gdSetSize m.geAnnStruct.glhPenStyle, j, 1
    gdSetSize m.geAnnStruct.glhPenSize, j, 1
    gdSetSize m.geAnnStruct.glhUnderline, j, 1
    gdSetSize m.geAnnStruct.glhStyle, j, 1
    gdSetSize m.geAnnStruct.glhAlign, j, 1
    gdSetSize m.geAnnStruct.glhSize, j, 1
    gdSetSize m.geAnnStruct.gdhTop, j, 1
    gdSetSize m.geAnnStruct.gdhLeft, j, 1
    gdSetSize m.geAnnStruct.gdhBottom, j, 1
    gdSetSize m.geAnnStruct.gdhRight, j, 1
    gdSetSize m.geAnnStruct.gdhMisc, j, 1
    gdSetSize m.geAnnStruct.gshText, j, 1
    gdSetSize m.geAnnStruct.gshFont, j, 1
    
    If m.Y1 < m.Y2 Then
        dTextY = m.Y1
    Else
        dTextY = m.Y2
    End If
    
    'fib lines
    For i = 0 To m.aRatios.Size - 1
        If m.aRatioShow(i) = 1 Then
            nX = dX3 + Int(nBars * m.aRatios(i) + 0.5)
            If m.aRatioColor.Size > i Then
               iColor = m.aRatioColor(i)
            Else
                iColor = nPropColor
            End If

            gdSetNum m.geAnnStruct.glhAnnType, i + 5, 1
            gdSetNum m.geAnnStruct.gdhTop, i + 5, m.Y1
            gdSetNum m.geAnnStruct.gdhLeft, i + 5, nX
            gdSetNum m.geAnnStruct.gdhBottom, i + 5, m.Y2
            gdSetNum m.geAnnStruct.gdhRight, i + 5, nX
            gdSetNum m.geAnnStruct.glhPenColor, i + 5, iColor
            geSetPenStyle nStyle, i + 5
            
            If nShowValues = 0 Then
                gdSetStr m.geAnnStruct.gshText, i + 5, ""
            Else
                gdSetStr m.geAnnStruct.gshText, i + 5, m.aRatios(i)
                'set font info
                gdSetStr m.geAnnStruct.gshFont, i + 5, Prop("FontName")
                gdSetNum m.geAnnStruct.glhSize, i + 5, Val(Prop("FontSize"))
                gdSetNum m.geAnnStruct.glhStyle, i + 5, Val(Prop("FontStyle"))
                gdSetNum m.geAnnStruct.glhUnderline, i + 5, Val(Prop("FontUnderline"))
            End If
            
            If iShowNumBars = 1 Then
                gdSetNum m.geAnnStruct.glhAnnType, j, 0
                gdSetNum m.geAnnStruct.gdhTop, j, dTextY
                gdSetNum m.geAnnStruct.gdhBottom, j, dTextY
                gdSetNum m.geAnnStruct.gdhLeft, j, nX
                gdSetNum m.geAnnStruct.gdhRight, j, nX
                gdSetStr m.geAnnStruct.gshText, j, Abs(dX1 - nX)
                gdSetNum m.geAnnStruct.glhPenColor, j, iColor
                gdSetStr m.geAnnStruct.gshFont, j, Prop("FontName")
                gdSetNum m.geAnnStruct.glhSize, j, Val(Prop("FontSize"))
                gdSetNum m.geAnnStruct.glhStyle, j, Val(Prop("FontStyle"))
                gdSetNum m.geAnnStruct.glhUnderline, j, Val(Prop("FontUnderline"))
                gdSetNum m.geAnnStruct.glhAlign, j, 10
                j = j + 1
            End If
        Else
            gdSetNum m.geAnnStruct.glhAnnType, i + 5, -1
        End If
    Next
            
    gdSetStr m.geAnnStruct.gshText, 2, ""
    gdSetStr m.geAnnStruct.gshText, 3, ""
    gdSetStr m.geAnnStruct.gshText, 4, ""
    
    If Prop("ShowRatioZero") = 1 And iShowDates = 1 Then
        gdSetStr m.geAnnStruct.gshText, 2, "Start " & DateFormat(m.dDate1, MM_DD_YYYY)
    End If
    
    If Prop("ShowRatioOne") = 1 Then
        If iShowDates = 1 Then gdSetStr m.geAnnStruct.gshText, 3, "End " & DateFormat(m.dDate2, MM_DD_YYYY)
        If iShowNumBars = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, j, 0
            gdSetNum m.geAnnStruct.gdhTop, j, dTextY
            gdSetNum m.geAnnStruct.gdhBottom, j, dTextY
            gdSetNum m.geAnnStruct.gdhLeft, j, dX2
            gdSetNum m.geAnnStruct.gdhRight, j, dX2
            gdSetStr m.geAnnStruct.gshText, j, Abs(dX1 - dX2)
            gdSetNum m.geAnnStruct.glhPenColor, j, m.nColor
            gdSetStr m.geAnnStruct.gshFont, j, Prop("FontName")
            gdSetNum m.geAnnStruct.glhSize, j, Val(Prop("FontSize"))
            gdSetNum m.geAnnStruct.glhStyle, j, Val(Prop("FontStyle"))
            gdSetNum m.geAnnStruct.glhUnderline, j, Val(Prop("FontUnderline"))
            gdSetNum m.geAnnStruct.glhAlign, j, 10
            j = j + 1
        End If
    End If
    
    If Prop("ShowRatioTwo") = 1 Then
        If iShowDates = 1 Then gdSetStr m.geAnnStruct.gshText, 4, "Extend " & DateFormat(m.aDates(0), MM_DD_YYYY)
        If iShowNumBars = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, j, 0
            gdSetNum m.geAnnStruct.gdhTop, j, dTextY
            gdSetNum m.geAnnStruct.gdhBottom, j, dTextY
            gdSetNum m.geAnnStruct.gdhLeft, j, dX3
            gdSetNum m.geAnnStruct.gdhRight, j, dX3
            gdSetStr m.geAnnStruct.gshText, j, Abs(dX1 - dX3)
            gdSetNum m.geAnnStruct.glhPenColor, j, m.nColor
            gdSetStr m.geAnnStruct.gshFont, j, Prop("FontName")
            gdSetNum m.geAnnStruct.glhSize, j, Val(Prop("FontSize"))
            gdSetNum m.geAnnStruct.glhStyle, j, Val(Prop("FontStyle"))
            gdSetNum m.geAnnStruct.glhUnderline, j, Val(Prop("FontUnderline"))
            gdSetNum m.geAnnStruct.glhAlign, j, 10
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetETRatio", eGDRaiseError_Raise

End Function

Private Function SetAndrewFork(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX) As Long
On Error GoTo ErrSection:

    Dim i&, j&, nStyle&, nColor&
    Dim dX3#, dxLeft#, dBarFirstDate#
    
    Dim a1 As New cGdArray
    Dim a2 As New cGdArray
    Dim a3 As New cGdArray
    
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        'set annotation type
        gdSetNum m.geAnnStruct.glhAnnType, 0, 23    'handle (leftmost x-value)
        gdSetNum m.geAnnStruct.glhAnnType, 1, 23    'one line
        gdSetNum m.geAnnStruct.glhAnnType, 2, 23    'other line
        'font
        gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
    End If
    
    'aardvark 4207
    If m.geAnnStruct.moveable = 1 Then
        dBarFirstDate = Chart.Bars(eBARS_DateTime, 0)
        If dBarFirstDate <= 0 Then
            Exit Function
        'ElseIf m.dDate1 < dBarFirstDate Or m.dDate2 < dBarFirstDate Or m.aDates(0) < dBarFirstDate Then
        ElseIf dX1 = kNullData Or dX2 = kNullData Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, -1
            gdSetNum m.geAnnStruct.glhAnnType, 1, -1
            gdSetNum m.geAnnStruct.glhAnnType, 2, -1
            Exit Function
        End If
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Graphics Engine Fields Info:
'
' 1. minorAxisLenData   :fork style (see enum)
' 2. showAxes           :show/hide vertical lines for time interval markers
' 3. handleShow         :extend/don't extend base of fork to outer parallel lines
' 4. showQtrLines       :show/hide parallels (quarter lines for old style fork)
' 5. lenRatio           :show/hide text labels for parallels
' 6. lenPoints          :show/hide text labels for time interval markers
' 7. glhJustify(0)      :1=hide AB line
' 8. glhJustify(1)      :1=extend vertical time interval lines through all panes (like fib time extension tool)
' 9. glhJustify(2)      :0=parallels above & below, 1=above, 2=below (like trendline channels option)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
    gdSetNum m.geAnnStruct.glhJustify, 0, Val(Prop("HideABLine"))
    gdSetNum m.geAnnStruct.glhJustify, 1, Val(Prop("ShowInAllPanes"))
    gdSetNum m.geAnnStruct.glhJustify, 2, Val(Prop("ParallelLocation"))
    
    If m.bSchiffFork Then
        i = Val(Prop("ForkMode"))
        m.geAnnStruct.minorAxisLenData = i
        
'JM 09-24-2010: use same pen style for Schiff Fork (project meeting mod)
'        i = Val(Prop("SchiffStyle"))
        geSetPenStyle m.nStyle, 1
        gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
    Else
        m.geAnnStruct.minorAxisLenData = 0
    End If
    
    
    If m.aDates.Size > 0 Then
        dX3 = GetX(Chart, m.aDates(0)) - nStartScreenX              '4525
    Else
        dX3 = dX1 + dX2     'just set to something bigger
    End If
                
    'find left most X
    a1.Add dX1
    a1.Add dX2
    a1.Add dX3
    a1.Sort eGdSort_Default
    
    dxLeft = a1(0)
    If dX1 = dxLeft Then
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
        
        If dX2 = a1(1) Then
            gdSetNum m.geAnnStruct.gdhLeft, 1, dX2
            gdSetNum m.geAnnStruct.gdhTop, 1, m.Y2
            gdSetNum m.geAnnStruct.gdhLeft, 2, dX3
            gdSetNum m.geAnnStruct.gdhTop, 2, m.aYs(0)
        Else
            gdSetNum m.geAnnStruct.gdhLeft, 1, dX3
            gdSetNum m.geAnnStruct.gdhTop, 1, m.aYs(0)
            gdSetNum m.geAnnStruct.gdhLeft, 2, dX2
            gdSetNum m.geAnnStruct.gdhTop, 2, m.Y2
        End If
    ElseIf dX2 = dxLeft Then
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX2
        gdSetNum m.geAnnStruct.gdhTop, 0, m.Y2
        
        If dX1 = a1(1) Then
            gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
            gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
            gdSetNum m.geAnnStruct.gdhLeft, 2, dX3
            gdSetNum m.geAnnStruct.gdhTop, 2, m.aYs(0)
        Else
            gdSetNum m.geAnnStruct.gdhLeft, 1, dX3
            gdSetNum m.geAnnStruct.gdhTop, 1, m.aYs(0)
            gdSetNum m.geAnnStruct.gdhLeft, 2, dX1
            gdSetNum m.geAnnStruct.gdhTop, 2, m.Y1
        End If
    ElseIf dX3 = dxLeft Then
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX3
        gdSetNum m.geAnnStruct.gdhTop, 0, m.aYs(0)
        
        If dX1 = a1(1) Then             '6090 - A,B,C must be in date order
            gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
            gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
            gdSetNum m.geAnnStruct.gdhLeft, 2, dX2
            gdSetNum m.geAnnStruct.gdhTop, 2, m.Y2
        Else
            gdSetNum m.geAnnStruct.gdhLeft, 1, dX2
            gdSetNum m.geAnnStruct.gdhTop, 1, m.Y2
            gdSetNum m.geAnnStruct.gdhLeft, 2, dX1
            gdSetNum m.geAnnStruct.gdhTop, 2, m.Y1
        End If
    End If
    
    'font
    gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
    gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
    gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
    
    'text (reset to off)
    m.geAnnStruct.lenRatio = 0
    m.geAnnStruct.lenPoints = 0
    
    If m.bSchiffFork Then
        Set a1 = New cGdArray
        'set flag & values for markers
        m.geAnnStruct.showAxes = 0
        a1.SplitFields Prop("MarkerRatios")
        a2.SplitFields Prop("MarkerShow")
        
        j = 3
        If a1.Size = a2.Size Then
            nStyle = Val(Prop("MarkerStyle"))
            nColor = Val(Prop("MarkerColor"))
            geSetPenStyle nStyle, 2
            gdSetNum m.geAnnStruct.glhPenColor, 2, nColor
            
            For i = 0 To a1.Size - 1
                If Val(a2(i)) <> 0 Then
                    gdSetNum m.geAnnStruct.gdhLeft, j, Val(a1(i))           '6559 - do not use valoftext; it returns 0 for 0.382 etc.
                Else
                    gdSetNum m.geAnnStruct.gdhLeft, j, kNullData
                End If
                j = j + 1
            Next
            
            If a1.Size > 0 Then
                m.geAnnStruct.showAxes = 1
                m.geAnnStruct.lenPoints = Val(Prop("MarkerText"))
            End If
        End If
        If gdGetSize(m.geAnnStruct.gdhLeft) > j Then
            For i = j To gdGetSize(m.geAnnStruct.gdhLeft) - 1
                gdSetSize m.geAnnStruct.gdhLeft, i, kNullData       'clear any removed ratios
            Next
        End If
        
        'set flag & values for drawing parallel lines  (used to be quarter lines)
        m.geAnnStruct.showQtrLines = 0      'assume no parallel lines
        a1.SplitFields Prop("ParallelRatios")
        a2.SplitFields Prop("ParallelShow")
        a3.SplitFields Prop("ParallelColor")
        
        j = 3
        If a1.Size = a2.Size And a1.Size = a3.Size Then
            nStyle = Val(Prop("ParallelStyle"))
            
            For i = 0 To a1.Size - 1
                If Val(a2(i)) <> 0 Then
                    gdSetNum m.geAnnStruct.gdhTop, j, Val(a1(i))           '6559 - do not use valoftext; it returns 0 for 0.382 etc.
                    m.geAnnStruct.showQtrLines = 1
                Else
                    gdSetNum m.geAnnStruct.gdhTop, j, kNullData
                End If
                geSetPenStyle nStyle, j
                gdSetNum m.geAnnStruct.glhPenColor, j, Val(a3(i))
                j = j + 1
            Next
        End If
        If gdGetSize(m.geAnnStruct.gdhTop) > j Then
            For i = j To gdGetSize(m.geAnnStruct.gdhTop) - 1
                gdSetSize m.geAnnStruct.gdhTop, i, kNullData       'clear any removed ratios
            Next
        End If
        
        If m.geAnnStruct.showQtrLines <> 0 Then
            m.geAnnStruct.lenRatio = Val(Prop("ParallelText"))
        End If
        
        'set flag for extending base of fork
        m.geAnnStruct.handleShow = Val(Prop("ShowHandle"))
        
    Else
        'set flag for drawing quarter lines
        m.geAnnStruct.showQtrLines = Int(Val(Prop("ShowQtrLines")))
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetAndrewFork", eGDRaiseError_Raise

End Function

Private Sub SetDnExpansion(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&)
On Error GoTo ErrSection:

    Dim i&, nSize&
    Dim nAnnType&, nLenABC&, nLenRatios&, nColor&
    Dim idxA&, idxC&
    Dim hDateArray&, hYArray&
    
    Dim dYA#, dYB#, dYC#, dRatioD#, dDate3#
    Dim bHasFib As Boolean
    Dim Ind As cIndicator
    Dim Bars As cGdBars
    
    Dim aRatioText As New cGdArray
        
    bHasFib = UseDiNapFib()
    
    If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 Or m.eType = eANNOT_DNExpansion4 Then
        '02-11-2011: since Dinap and NonDinap ABC tools are mutually exclusive and draw code is shared
        '   for Dinap ABC, Genesis ABC & ABCD tools, it is best to not mix up these tools in the template.
        '
        ' Only the Dinap ABC tool has the ratioCOP, ratioOP, and ratioXOP values set
        If bHasFib Then
            If Len(Prop("ratioCOP")) = 0 Or Len(Prop("ratioOP")) = 0 Or Len(Prop("ratioXOP")) = 0 Then
                Exit Sub    'originally drawn as Non-Dinapoli ABC
            End If
        ElseIf Len(Prop("ratioCOP")) > 0 Or Len(Prop("ratioOP")) > 0 Or Len(Prop("ratioXOP")) > 0 Then
            Exit Sub    'originally drawn as Dinapoli ABC
        End If
    End If
    
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    'or when ratios are added/removed from fib extension
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
        
    If m.geAdded = False Or i < 0 Then
        If bHasFib Then
            nSize = 6
        Else
            nSize = m.aRatios.Size + 3
        End If
                
        For i = 0 To nSize - 1
            gdSetNum m.geAnnStruct.glhAnnType, i, 14
            gdSetNum m.geAnnStruct.gdhLeft, i, 1#       'place holders
            gdSetNum m.geAnnStruct.glhImageType, i, -1  'default to non-resizeable horz lines
        Next
    End If
    
    m.geAnnStruct.handleShow = Val(Prop("ShowHandle"))
    gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
    gdSetStr m.geAnnStruct.gshFont, 0, Prop("Font")
        
    nAnnType = 14

    dYA = m.Y1
    dYB = m.Y2
    
    If bHasFib And m.eType = eANNOT_DNExpansion Then
        'COP
        geSetPenStyle Val(Prop("penSizeCOP")), 3
        gdSetNum m.geAnnStruct.glhPenColor, 3, Val(Prop("colorCOP"))
        gdSetStr m.geAnnStruct.gshText, 3, Prop("textCOP")
        gdSetNum m.geAnnStruct.gdhMisc, 3, Val(Prop("ratioCOP"))
        'OP
        geSetPenStyle Val(Prop("penSizeOP")), 4
        gdSetNum m.geAnnStruct.glhPenColor, 4, Val(Prop("colorOP"))
        gdSetStr m.geAnnStruct.gshText, 4, Prop("textOP")
        gdSetNum m.geAnnStruct.gdhMisc, 4, Val(Prop("ratioOP"))
        'XOP
        geSetPenStyle Val(Prop("penSizeXOP")), 5
        gdSetNum m.geAnnStruct.glhPenColor, 5, Val(Prop("colorXOP"))
        gdSetStr m.geAnnStruct.gshText, 5, Prop("textXOP")
        gdSetNum m.geAnnStruct.gdhMisc, 5, Val(Prop("ratioXOP"))
        nLenABC = -1
        
        m.geAnnStruct.showAxes = 0                      'text on left always
        m.geAnnStruct.showQtrLines = 0                  'never extend ratio lines
        m.geAnnStruct.minorAxisLenData = 0              'never extend primary lines
        m.geAnnStruct.lenPoints = 0                     'never draw D point
    Else
        nLenABC = Val(Prop("LenA"))
        nLenRatios = Val(Prop("LenCOP"))
        nColor = Val(Prop("FibColor"))
        aRatioText.SplitFields Prop("FibRatiosText"), ","
        
        Prop("penSizeA") = m.nStyle
        Prop("penSizeB") = m.nStyle
        Prop("penSizeC") = m.nStyle
        
        'JM 06-11-2012: aardvark 6685
        '   Elliott Wave requested that option be given for showing values in y-scale area
        '   0=left, 1=right, 2=in y-scale
        m.geAnnStruct.showAxes = Val(Prop("TextOnRight"))
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'original implementation only had Prop("Ext")     - 4294
        '1=extend ratios right, 2=extend primary lines right, 3=extend all lines right, 0=extend nothing
        '
        '09-11-2008: modified to extend lines to left as well
        'Prop("Ext"): 1=extend ratios right, 2=extend ratios left, 3=extend ratios both direction, 0=no extension
        'Prop("ExtMain"): 1=extend primary right, 2=extend primary left, 3=extend primary both direction, 0=no extension
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        m.geAnnStruct.showQtrLines = Int(Val(Prop("Ext")))
        m.geAnnStruct.minorAxisLenData = Int(Val(Prop("ExtMain")))
                
        Set Ind = Chart.Tree("PRICE")
        If Not Ind Is Nothing Then
            If Ind.DisplayType = eINDIC_Line Then       '4340
                Set Bars = Ind.Bars
                i = Bars.FindDateTime(m.dDate1)
                If i >= 0 And i < Bars.Size Then
                    If Int(Bars(eBARS_DateTime, i)) <> Int(m.dDate1) Then
                        If i > 0 Then
                            If Int(Bars(eBARS_DateTime, i - 1)) = Int(m.dDate1) Then i = i - 1
                        End If
                    End If
                    dYA = Bars(eBARS_Close, i)
                End If
                i = Bars.FindDateTime(m.dDate2)
                If i >= 0 And i < Bars.Size Then
                    If Int(Bars(eBARS_DateTime, i)) <> Int(m.dDate2) Then
                        If i > 0 Then
                            If Int(Bars(eBARS_DateTime, i - 1)) = Int(m.dDate2) Then i = i - 1
                        End If
                    End If
                    dYB = Bars(eBARS_Close, i)
                End If
            End If
        End If
        
        nSize = 3
        If m.eType = eANNOT_FibABCD Then
            dRatioD = Val(Prop("RatioD"))
            If dRatioD <= 0# Then
                gdSetNum m.geAnnStruct.glhAnnType, 3, -1
                gdSetNum m.geAnnStruct.gdhMisc, 3, dRatioD
            Else
                gdSetNum m.geAnnStruct.glhAnnType, 3, nAnnType
                gdSetNum m.geAnnStruct.gdhMisc, 3, dRatioD
                geSetPenStyle Val(Prop("FibStyle")), 3
                gdSetNum m.geAnnStruct.glhImageType, 3, nLenABC
                
                gdSetNum m.geAnnStruct.glhPenColor, 3, Val(Prop("colorD"))
                gdSetStr m.geAnnStruct.gshText, 3, Prop("textD")
            End If
            nSize = 4
        End If
        
        For i = 0 To m.aRatios.Size - 1
            If m.aRatioShow(i) = 0 Or m.aRatios(i) = dRatioD Then
                gdSetNum m.geAnnStruct.glhAnnType, i + nSize, -1
                gdSetNum m.geAnnStruct.gdhMisc, i + nSize, m.aRatios(i)
            Else
                gdSetNum m.geAnnStruct.glhAnnType, i + nSize, nAnnType
                gdSetNum m.geAnnStruct.gdhMisc, i + nSize, m.aRatios(i)
                geSetPenStyle Val(Prop("FibStyle")), i + nSize
                gdSetNum m.geAnnStruct.glhImageType, i + nSize, nLenRatios
                
                If m.aRatioColor.Size > i Then
                    gdSetNum m.geAnnStruct.glhPenColor, i + nSize, m.aRatioColor(i)
                Else
                    gdSetNum m.geAnnStruct.glhPenColor, i + nSize, nColor
                End If
                
                If aRatioText.Size > i Then
                    gdSetStr m.geAnnStruct.gshText, i + nSize, aRatioText(i)
                Else
                    gdSetStr m.geAnnStruct.gshText, i + nSize, ""
                End If
            End If
        Next
    End If
            
    idxA = 0    'assume user going left to right (i.e. forward in dates)
    idxC = 2
    'B
    geSetPenStyle Val(Prop("penSizeB")), 1
    gdSetNum m.geAnnStruct.glhPenColor, 1, Val(Prop("colorB"))
    gdSetStr m.geAnnStruct.gshText, 1, Prop("textB")
        
    hDateArray = m.aDates.ArrayHandle
    hYArray = m.aYs.ArrayHandle
    
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1  'point A
    gdSetNum m.geAnnStruct.gdhTop, 0, dYA   'm.Y1
    gdSetNum m.geAnnStruct.glhImageType, 0, nLenABC
    
    If m.dDate2 = -1 Then
        gdSetNum m.geAnnStruct.gdhLeft, 1, -1  'point B
        gdSetNum m.geAnnStruct.gdhTop, 1, -1
    Else
        gdSetNum m.geAnnStruct.gdhLeft, 1, dX2  'point B
        gdSetNum m.geAnnStruct.gdhTop, 1, dYB   'm.Y2
        gdSetNum m.geAnnStruct.glhImageType, 1, nLenABC
    End If
    
    If gdGetSize(hDateArray) > 0 Then
        dDate3 = gdGetNum(hDateArray, 0)
        
        If Bars Is Nothing Then
            dYC = gdGetNum(hYArray, 0)
        Else
            i = Bars.FindDateTime(dDate3)
            If i >= 0 And i < Bars.Size Then
                If Int(Bars(eBARS_DateTime, i)) <> Int(dDate3) Then
                    If i > 0 Then
                        If Int(Bars(eBARS_DateTime, i - 1)) = Int(dDate3) Then i = i - 1
                    End If
                End If
                dYC = Bars(eBARS_Close, i)
            End If
        End If
    
        dX2 = GetX(Chart, dDate3) - nStartScreenX           '4525
        gdSetNum m.geAnnStruct.gdhLeft, 2, dX2 'point C
        gdSetNum m.geAnnStruct.gdhTop, 2, dYC   'gdGetNum(hYArray, 0)
        gdSetNum m.geAnnStruct.glhImageType, 2, nLenABC
        
        'switch A & C properties if user is going backwards in dates
        If dX1 > dX2 Then
            If m.dDate1 = m.dDate2 Then
                'User's first 2 clicks are on same bar, but 3rd click is backwards in date.
                gdSetNum m.geAnnStruct.gdhLeft, 0, dX2 'point A
                gdSetNum m.geAnnStruct.gdhTop, 0, dYC   'gdGetNum(hYArray, 0)
                gdSetNum m.geAnnStruct.gdhLeft, 2, dX1  'point C
                gdSetNum m.geAnnStruct.gdhTop, 2, dYA   'm.Y1
            Else
                idxA = 2
                idxC = 0
            End If
        End If
    Else
        gdSetNum m.geAnnStruct.gdhLeft, 2, -1   'point C
        gdSetNum m.geAnnStruct.gdhTop, 2, -1
    End If
            
    'A
    geSetPenStyle Val(Prop("penSizeA")), idxA
    gdSetNum m.geAnnStruct.glhPenColor, idxA, Val(Prop("colorA"))
    gdSetStr m.geAnnStruct.gshText, idxA, Prop("textA")
    'C
    geSetPenStyle Val(Prop("penSizeC")), idxC
    gdSetNum m.geAnnStruct.glhPenColor, idxC, Val(Prop("colorC"))
    gdSetStr m.geAnnStruct.gshText, idxC, Prop("textC")

    'font information
    gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
    gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))  '0=reg,1=bold,2=italic,3=bold italic
    gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))  '0=don't underline,1=underline
    
    gdSetNum m.geAnnStruct.glhFillPattern, 0, 0                 '
    
    'hide anchor lines if applicable
    If Not bHasFib Then
        If Prop("ShowRatioZero") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, 14
        Else
            gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        End If
        If Prop("ShowRatioOne") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, 1, 14
        Else
            gdSetNum m.geAnnStruct.glhAnnType, 1, -1
        End If
        If Prop("ShowRatioTwo") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, 2, 14
        Else
            gdSetNum m.geAnnStruct.glhAnnType, 2, -1
        End If
        
        If Val(Prop("ShowValues")) = 0 Then gdSetNum m.geAnnStruct.glhFillPattern, 0, -1
    End If
    
    If m.eType = eANNOT_FibABCD Then
        m.geAnnStruct.lenPoints = 0             'flag for grapheng to draw connecting lines for Fib AB=CD tool
        If Prop("ShowRatioZero") = 1 Then
            If Prop("ShowRatioOne") = 1 Then
                If Prop("ShowRatioTwo") = 1 Then
                    If dRatioD > 0 Then m.geAnnStruct.lenPoints = dRatioD
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetDnExpansion", eGDRaiseError_Raise

End Sub

Private Sub SetDNRBoundary(Chart As cChart)
On Error GoTo ErrSection:
    
    'Locates and flags the first data bar left of the focus point that
    'is higher than a high focus point or lower than a low focus point.
    'All reaction points must be to the right of this data bar.
    
    Dim i&, iPos&, dY#, dDate#
    
    gdBinarySearch Chart.Bars.ArrayHandle(eBARS_DateTime), gdFixDateTime(m.dDate1), iPos, eGdSort_Default, 0, Chart.Bars.Size - 1
        
    dDate = kNullData
    dY = kNullData
    For i = iPos To 0 Step -1
        If gdGetNum(m.geAnnStruct.glhDirection, 0) = 1 Then  'focus low
            dY = Chart.Bars(eBARS_Low, i)
            If dY < m.Y1 Then
                dDate = Chart.Bars(eBARS_DateTime, i)
                Exit For
            End If
        Else
            dY = Chart.Bars(eBARS_High, i)
            If dY > m.Y1 Then
                dDate = Chart.Bars(eBARS_DateTime, i)
                Exit For
            End If
        End If
    Next
        
    If dDate = kNullData Then
        m.dReactPtBoundary = 0    'reaction points can go back all the way to the first data bar
    Else
        m.dReactPtBoundary = dDate
    End If


ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetDNRBoundary", eGDRaiseError_Raise

End Sub

Private Sub SetDnRetracement(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&)
On Error GoTo ErrSection:

    'Design Note. Since grapheng.dll only need the top, left values
    '   for this annotation, the field "bottom" is used to store the
    '   fib ratio as it is a double. This was done rather than changing
    '   the data structure for now. If neccessary, the data structure
    '   can be changed later.
    
    Dim Size&, idx&, i&
    Dim hDateArray&, hYArray&, fibRatio#
    Dim dHighestLowest#, dDateHiLow#
    Dim bRC As Boolean
    Dim gdTextStr As New cGdArray
        
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        'set annotation type
        gdSetNum m.geAnnStruct.glhAnnType, 0, 15    '1st fib ratio
        gdSetNum m.geAnnStruct.glhAnnType, 1, 6     'arc properties for 2nd fib ratio
        gdSetNum m.geAnnStruct.glhAnnType, 2, 1     'text & line properties for 2nd fib ratio
        gdSetNum m.geAnnStruct.glhAnnType, 3, 6     'arc properties for 3rd fib ratio
        gdSetNum m.geAnnStruct.glhAnnType, 4, 1     'text & line properties for 3rd fib ratio
        gdSetNum m.geAnnStruct.glhAnnType, 5, 0     'font information
        'set focus hi/low flag (default=1 --> focus is bar's high)
        gdSetNum m.geAnnStruct.glhDirection, 0, Val(Prop("FocusLow"))
    End If
    
    m.geAnnStruct.handleShow = Val(Prop("ShowHandle"))
    m.nColor = Val(Prop("FibColorR0"))  'using as temporary local variables
    m.nStyle = Val(Prop("FibPenSizeR0"))
    gdTextStr.SplitFields m.strText, "~"
                
    'properties for 2nd fib ratio (default = 0.618)
    fibRatio = Val(Prop("FibR1"))
    If fibRatio < 0 Then    'in case user input negatives
        fibRatio = -1 * fibRatio
        Prop("FibR1") = fibRatio
    End If
    gdSetNum m.geAnnStruct.gdhBottom, 1, fibRatio   'ratio value
    gdSetNum m.geAnnStruct.glhPenColor, 1, Val(Prop("ArcColorR1"))  'arc
    geSetPenStyle Val(Prop("ArcPensizeR1")), 1
    gdSetNum m.geAnnStruct.gdhBottom, 2, fibRatio   'ratio value
    gdSetNum m.geAnnStruct.glhPenColor, 2, Val(Prop("FibColorR1"))  'text & lines
    geSetPenStyle Val(Prop("FibPensizeR1")), 2
            
    'properties for 3rd fib ratio (default = 0.382)
    fibRatio = Val(Prop("FibR2"))
    If fibRatio < 0 Then
        fibRatio = -1 * fibRatio
        Prop("FibR2") = fibRatio
    End If
    gdSetNum m.geAnnStruct.gdhBottom, 3, fibRatio   'ratio value
    gdSetNum m.geAnnStruct.glhPenColor, 3, Val(Prop("ArcColorR2"))  'arc
    geSetPenStyle Val(Prop("ArcPensizeR2")), 3
    gdSetNum m.geAnnStruct.gdhBottom, 4, fibRatio   'ratio value
    gdSetNum m.geAnnStruct.glhPenColor, 4, Val(Prop("FibColorR2"))  'text & lines
    geSetPenStyle Val(Prop("FibPensizeR2")), 4
        
    'properties for 1st fib ratio (default = 1.000)
    'this is done last because this ratio is used in for-loop below
    fibRatio = Val(Prop("FibR0"))
    If fibRatio < 0 Then
        fibRatio = -1 * fibRatio
        Prop("FibR0") = fibRatio
    End If
    gdSetNum m.geAnnStruct.gdhBottom, 0, fibRatio    'ratio value
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor     'text & lines
    geSetPenStyle m.nStyle, 0
    
    'font information
    gdSetStr m.geAnnStruct.gshFont, 5, Prop("FontName")
    gdSetNum m.geAnnStruct.glhSize, 5, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 5, Val(Prop("FontStyle"))  '0=reg,1=bold,2=italic,3=bold italic
    gdSetNum m.geAnnStruct.glhUnderline, 5, Val(Prop("FontUnderline"))  '0=don't underline,1=underline
    
    hDateArray = m.aDates.ArrayHandle
    hYArray = m.aYs.ArrayHandle
    bRC = False

    'focus point
    If Val(Prop("FocusDynamic")) = 1 Then
        If m.aDates.Size > 0 And m.geAnnStruct.moveable = 1 Then
            dX2 = m.aDates(0) 'first reaction point date
            'see if annotation should be shown
            If m.aDates(0) > Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False)) Then
                ShowDnRetracement False
                Exit Sub
            End If
            If Val(Prop("focuslow")) = 0 Then
                bRC = Chart.HighestHiForwardDate(dX2, dHighestLowest, dDateHiLow)
            Else
                bRC = Chart.LowestLowForwardDate(dX2, dHighestLowest, dDateHiLow)
            End If
            If bRC = True Then
                m.Y1 = dHighestLowest
                m.dDate1 = dDateHiLow
                dX1 = GetX(Chart, m.dDate1) - nStartScreenX         '4525
                m.X1 = dX1
                SetDNRBoundary Chart    'focus point changed, recheck reaction points boundary
            End If
        End If
    Else
        If m.geAnnStruct.moveable = 1 Then  'drawing process is done
            If m.dDate1 > Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False)) Then
                ShowDnRetracement False
                Exit Sub
            End If
        End If
    End If
    
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    
    'set reaction points boundary if it has not been set
    If m.dReactPtBoundary < 0 Then SetDNRBoundary Chart
    
    'check for duplicates
    Size = gdGetSize(hDateArray)
    If gdGetNum(hYArray, Size - 1) = gdGetNum(hYArray, Size - 2) Then
        gdDeleteItems hDateArray, Size - 1, 1
        gdDeleteItems hYArray, Size - 1, 1
    End If

    'set begining of data for reaction point (type = hz_line = 3)
    idx = 6 'hard-coded for now, can change if necessary
    
    'set data value for reaction points
    Size = gdGetSize(hDateArray)
    For i = 0 To Size - 1
        m.dDate2 = gdGetNum(hDateArray, i)
        dX2 = GetX(Chart, m.dDate2) - nStartScreenX             '4525
        gdSetNum m.geAnnStruct.gdhLeft, i + idx, dX2
        gdSetNum m.geAnnStruct.gdhTop, i + idx, gdGetNum(hYArray, i)
        gdSetNum m.geAnnStruct.glhPenColor, i + idx, RGB(0, 0, 255)
        
        'If m.geAnnStruct.moveable = 0 And m.dDate2 < m.dReactPtBoundary Then
        If m.dDate2 < m.dReactPtBoundary Then
            'set color of circle handle to red
            gdSetNum m.geAnnStruct.glhPenColor, i + idx, RGB(255, 0, 0)
        End If
        'array size can increase as user add reaction points
        gdSetNum m.geAnnStruct.glhSize, i + idx, fibRatio
        gdSetNum m.geAnnStruct.glhAnnType, i + idx, 3
        gdSetStr m.geAnnStruct.gshText, i + idx, gdTextStr(i)
        geSetPenStyle m.nStyle, i + idx
    Next
    Set gdTextStr = Nothing
    
    i = gdGetSize(m.geAnnStruct.glhAnnType)
    If Size + idx < i Then
        If i > 0 Then
            gdDeleteItems m.geAnnStruct.glhAnnType, i - 1, 1
            gdDeleteItems m.geAnnStruct.glhPenColor, i - 1, 1
            gdDeleteItems m.geAnnStruct.glhPenSize, i - 1, 1
            gdDeleteItems m.geAnnStruct.glhSize, i - 1, 1
        End If
    End If
    ShowDnRetracement True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetDnRetracement", eGDRaiseError_Raise

End Sub
        
Public Sub SetTextEditSize(Chart As cChart, ByVal strText$)
On Error GoTo ErrSection:

    Dim cInfo As coordinate_info
    Dim nPixWidth&, nPixHeight&, rc&
        
    gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
    gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
    gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
    gdSetStr m.geAnnStruct.gshText, 0, strText
    rc = geGetTextDimension(Chart.Form.pbChart.hDC, m.geAnnStruct, 0, nPixWidth, nPixHeight)
    If rc = 0 Then
        Prop("Width") = nPixWidth
        Prop("Height") = nPixHeight
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetTextEditSize", eGDRaiseError_Raise

End Sub

Private Sub SetFibArcs(Chart As cChart, ByVal dX1#, ByVal dX2#)
On Error GoTo ErrSection:

    Dim i&, k&, nShowValues&, nColor&, nStyle&
    Dim nSize&, nBars&
    Dim dX#, dY#, dRatio#
        
    'Developer Note: as of 08-18-2003
    'dX1=1st mouse click, dX2=2nd mouse click (this must be maintained for code to work)
    'The second mouse click is always the 1.00 ratio arc (as implemented by eSignal)

    'array item setup info
    '[0] ... [4]
    '       set annotation type to be ignored by graphics engine
    '       set direction field to NE, NW, SE, SW to indicate arcs to show
    '       graphics engine is hard-coded to check for the direction field
    '[5]    (top, left) point of 1st mouse click
    '[5]    (bottom, right) point of 2nd mouse click
    '[i+6]  coordinate values for all other arcs
    'Distance from the first mouse click to the second mouse click is the
    'radius of the arc/ellipse with the first mouse click being the center.

    For i = 0 To 4
        gdSetNum m.geAnnStruct.glhAnnType, i, -1
        gdSetNum m.geAnnStruct.glhDirection, i, -1
        gdSetStr m.geAnnStruct.gshText, i, "N"      'N for no text
    Next

    nShowValues = Val(Prop("ShowValues"))
    
    'set coordinate & text info for 1.00 arc
    If Prop("ShowRatioZero") = 1 Or Prop("ShowRatioOne") = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 5, 21             'fib arcs
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 5, -1
    End If
    gdSetNum m.geAnnStruct.gdhTop, 5, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 5, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 5, m.Y2
    gdSetNum m.geAnnStruct.gdhMisc, 5, 0#
    gdSetNum m.geAnnStruct.gdhRight, 5, dX2
    gdSetNum m.geAnnStruct.glhPenColor, 5, m.nColor
    
    For i = 0 To 5
        geSetPenStyle m.nStyle, i
    Next
    
    If nShowValues = 1 Then
        gdSetStr m.geAnnStruct.gshText, 5, Format(1#, "0.00")
    Else
        gdSetStr m.geAnnStruct.gshText, 5, "N"      'N for no text
    End If

    nColor = Val(Prop("FibColor"))
    nStyle = Val(Prop("FibStyle"))          'pen style
    nBars = Abs(dX1 - dX2)
    
    nSize = m.aRatios.Size - 1
    
    'set coordinate info for all other arcs
    For i = 0 To nSize
        If m.aRatioShow(i) = 1 Then
            dRatio = m.aRatios(i)
            dX = dX1 + nBars * dRatio
            gdSetNum m.geAnnStruct.glhAnnType, i + 6, 21
            gdSetNum m.geAnnStruct.gdhTop, i + 6, m.Y1
            gdSetNum m.geAnnStruct.gdhLeft, i + 6, dX1
            gdSetNum m.geAnnStruct.gdhBottom, i + 6, m.Y2
            gdSetNum m.geAnnStruct.gdhRight, i + 6, dX
            gdSetNum m.geAnnStruct.gdhMisc, i + 6, dRatio
            
            If m.aRatioColor.Size > i Then
                gdSetNum m.geAnnStruct.glhPenColor, i + 6, m.aRatioColor(i)
            Else
                gdSetNum m.geAnnStruct.glhPenColor, i + 6, nColor
            End If
            
            geSetPenStyle nStyle, i + 6
            
            If nShowValues Then
                gdSetStr m.geAnnStruct.gshText, i + 6, Format(dRatio, "0.00#")
            Else
                gdSetStr m.geAnnStruct.gshText, i + 6, "N"  'N for no text
            End If
        Else
            gdSetNum m.geAnnStruct.glhAnnType, i + 6, -1
        End If
    Next
    
    nSize = gdGetSize(m.geAnnStruct.glhAnnType) - 1
    For i = 0 To nSize
        gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
        gdSetNum m.geAnnStruct.glhAlign, i, 3       'e_btmLeft
    Next
    
    'set direction for arcs to show
    Dim nDirNE&, nDirNW&, nDirSE&, nDirSW&
    
    nDirNE = Val(Prop("DirNE"))
    nDirNW = Val(Prop("DirNW"))
    nDirSE = Val(Prop("DirSE"))
    nDirSW = Val(Prop("DirSW"))
    
    If nDirNE = 0 And nDirNW = 0 And nDirSE = 0 And nDirSW = 0 Then
        Prop("DirNE") = 1
        Prop("DirNW") = 1
        Prop("DirSE") = 1
        Prop("DirSW") = 1
    End If
    
    If Val(Prop("DirNE")) = 1 Then gdSetNum m.geAnnStruct.glhDirection, 0, 1
    If Val(Prop("DirNW")) = 1 Then gdSetNum m.geAnnStruct.glhDirection, 1, 7
    If Val(Prop("DirSE")) = 1 Then gdSetNum m.geAnnStruct.glhDirection, 2, 3
    If Val(Prop("DirSW")) = 1 Then gdSetNum m.geAnnStruct.glhDirection, 3, 5
    
    '0=draw normal, 1=extend arcs
    m.geAnnStruct.showAxes = 0          'Val(Prop("Ext"))
    '0=draw normal, 1=keep circular
    m.geAnnStruct.showQtrLines = Val(Prop("Circular"))
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetFibArcs", eGDRaiseError_Raise

End Sub

Private Sub SetFibTime(Chart As cChart, ByVal dX1#, ByVal dX2#)
On Error GoTo ErrSection:
    
    Dim i&, k&, iExt&, iRatioOneVal&, nSize&, nRatios&, nBars&, nX&
    Dim nStyle&, nAlign&, nShowValues&
    Dim iShowDates&, iShowNumBars&
    Dim nPropColor&, iColor&, dTextY#

    'Design change: 10-14-2010
    'TLB email: change code so the 2 selected points are "-1" and "0"
    '   1st click changed from ratio 0 to -1
    '   2nd click changed from ratio 1 to 0
    'Mod 04-04-2012 per request from Elliot Wave - aardvark 6646

    'Developer Note: as of 08-08-2003
    'dX1=1st mouse click, dX2=2nd mouse click (this must be maintained for code to work)
    'The second mouse click is always the 1.00 ratio line (as implemented by eSignal)
    
    'array item setup info
    '[0]    set to be ignored by graphics engine
    '[1]    color & style for drawing in all panes
    '[2]    line to be drawn at point of 1st mouse click (-1.00 ratio)
    '[3]    line to be drawn at  point of 2nd mouse click (0.00 ratio)
    '[4]    text info for line at 1st mouse click
    '[5]    text info for line at 2nd mouse click
    
    nPropColor = Val(Prop("FibColor"))
    nStyle = Val(Prop("ExtStyle"))
    nShowValues = Val(Prop("ShowValues"))
    iExt = Val(Prop("ShowInAllPanes"))
    iShowDates = Val(Prop("ShowDates"))
    iShowNumBars = Val(Prop("ShowNumBars"))
        
    m.geAnnStruct.showAxes = iExt
    gdSetNum m.geAnnStruct.glhAnnType, 0, -1
    'pen info for drawing across all panes
    gdSetNum m.geAnnStruct.glhAnnType, 1, 26
    gdSetNum m.geAnnStruct.glhPenColor, 1, -1
    geSetPenStyle nStyle, 1
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
    
    nBars = Abs(dX1 - dX2)
    m.aDates.Size = 0
    iRatioOneVal = Val(Prop("RatioOneVal"))
    
    If m.Y2 > m.Y1 Then
        nAlign = 3      'e_btmLeft
    Else
        nAlign = 0      'e_topLeft
    End If
    
    'set coordinate & text info for line at 1st mouse click
    If Prop("ShowRatioZero") = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 2, 1             'line
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 2, -1
    End If
    gdSetNum m.geAnnStruct.gdhTop, 2, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 2, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 2, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 2, dX1
    gdSetNum m.geAnnStruct.glhPenColor, 2, m.nColor
    geSetPenStyle m.nStyle, 2
    If iShowDates = 1 And Prop("ShowRatioZero") = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 4, 0             'text
        If m.Y2 > m.Y1 Then
            gdSetNum m.geAnnStruct.gdhTop, 4, m.Y1
            gdSetNum m.geAnnStruct.gdhBottom, 4, m.Y2
        Else
            gdSetNum m.geAnnStruct.gdhTop, 4, m.Y2
            gdSetNum m.geAnnStruct.gdhBottom, 4, m.Y1
        End If
        gdSetNum m.geAnnStruct.gdhLeft, 4, dX1
        gdSetNum m.geAnnStruct.gdhRight, 4, dX1
        gdSetStr m.geAnnStruct.gshText, 4, DateFormat(m.dDate1, MM_DD_YYYY)
        gdSetNum m.geAnnStruct.glhAlign, 4, 10              'align vertical like time expansion
        gdSetNum m.geAnnStruct.glhPenColor, 4, m.nColor
        geSetPenStyle m.nStyle, 4
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 4, -1
    End If

    'set coordinate & text info for line at 2nd mouse click
    If Prop("ShowRatioOne") = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 3, 1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 3, -1
    End If
    gdSetNum m.geAnnStruct.gdhTop, 3, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 3, dX2
    gdSetNum m.geAnnStruct.gdhBottom, 3, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 3, dX2
    gdSetNum m.geAnnStruct.glhPenColor, 3, m.nColor
    geSetPenStyle m.nStyle, 3
    If iShowDates = 1 And Prop("ShowRatioOne") = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 5, 0             'text
        If m.Y2 > m.Y1 Then
            gdSetNum m.geAnnStruct.gdhTop, 5, m.Y1
            gdSetNum m.geAnnStruct.gdhBottom, 5, m.Y2
        Else
            gdSetNum m.geAnnStruct.gdhTop, 5, m.Y2
            gdSetNum m.geAnnStruct.gdhBottom, 5, m.Y1
        End If
        gdSetNum m.geAnnStruct.gdhLeft, 5, dX2
        gdSetNum m.geAnnStruct.gdhRight, 5, dX2
        gdSetStr m.geAnnStruct.gshText, 5, DateFormat(m.dDate2, MM_DD_YYYY)
        gdSetNum m.geAnnStruct.glhAlign, 5, 10
        gdSetNum m.geAnnStruct.glhPenColor, 5, m.nColor
        geSetPenStyle m.nStyle, 5
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 5, -1
    End If
    
    nStyle = Val(Prop("FibStyle"))
    nRatios = m.aRatios.Size
    
    'remove -1 from array
    For i = nRatios - 1 To 0 Step -1
        If -1 = m.aRatios(i) Then
            m.aRatios.Remove i
            m.aRatioShow.Remove i
            m.aRatioColor.Remove i
            Exit For
        End If
    Next
    nRatios = m.aRatios.Size
    
    'fib lines
    For i = 0 To nRatios - 1
        If m.aRatioShow(i) = 1 Then
            If iRatioOneVal = 1 Then
                If dX2 > dX1 Then
                    nX = dX1 + Int(nBars * m.aRatios(i) + 0.5)
                Else
                    nX = dX2 + Int(nBars * m.aRatios(i) + 0.5)
                End If
            Else
                If dX2 > dX1 Then
                    nX = dX2 + Int(nBars * m.aRatios(i) + 0.5)
                Else
                    nX = dX1 + Int(nBars * m.aRatios(i) + 0.5)
                End If
            End If
            gdSetNum m.geAnnStruct.glhAnnType, i + 6, 1
            gdSetNum m.geAnnStruct.gdhLeft, i + 6, nX
            gdSetNum m.geAnnStruct.gdhRight, i + 6, nX
            
            If m.Y2 > m.Y1 Then
                gdSetNum m.geAnnStruct.gdhTop, i + 6, m.Y1
                gdSetNum m.geAnnStruct.gdhBottom, i + 6, m.Y2
            Else
                gdSetNum m.geAnnStruct.gdhTop, i + 6, m.Y2
                gdSetNum m.geAnnStruct.gdhBottom, i + 6, m.Y1
            End If

            If m.aRatioColor.Size > i Then
                gdSetNum m.geAnnStruct.glhPenColor, i + 6, m.aRatioColor(i)
            Else
                gdSetNum m.geAnnStruct.glhPenColor, i + 6, nPropColor
            End If
            
            geSetPenStyle nStyle, i + 6
        Else
            gdSetNum m.geAnnStruct.glhAnnType, i + 6, -1
        End If
    Next
        
    If m.Y1 < m.Y2 Then
        dTextY = m.Y1
    Else
        dTextY = m.Y2
    End If
    
    nSize = m.aRatios.Size - 1
    k = gdGetSize(m.geAnnStruct.glhAnnType)
    
    'set fib text
    For i = 0 To nSize
        If m.aRatioShow(i) = 1 Then
            nX = gdGetNum(m.geAnnStruct.gdhLeft, i + 6)
            If m.aRatioColor.Size > i Then
                iColor = m.aRatioColor(i)
            Else
                iColor = nPropColor
            End If
                        
            If nShowValues = 1 Then
                gdSetStr m.geAnnStruct.gshText, i + 6, m.aRatios(i)
                gdSetNum m.geAnnStruct.glhAlign, i + 6, 10
            End If
            
            If iShowNumBars = 1 Then
                gdSetNum m.geAnnStruct.glhAnnType, k, 0
                gdSetNum m.geAnnStruct.gdhTop, k, dTextY
                gdSetNum m.geAnnStruct.gdhBottom, k, dTextY
                gdSetNum m.geAnnStruct.glhPenColor, k, iColor
                gdSetNum m.geAnnStruct.gdhLeft, k, nX
                gdSetNum m.geAnnStruct.gdhRight, k, nX
                gdSetStr m.geAnnStruct.gshText, k, Abs(nX - dX1)
                k = k + 1
            End If
        End If
    Next
    
    If iShowNumBars = 1 And Prop("ShowRatioOne") = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, k, 0
        gdSetNum m.geAnnStruct.gdhTop, k, dTextY
        gdSetNum m.geAnnStruct.gdhBottom, k, dTextY
        gdSetNum m.geAnnStruct.glhPenColor, k, m.nColor
        gdSetNum m.geAnnStruct.gdhLeft, k, dX2
        gdSetNum m.geAnnStruct.gdhRight, k, dX2
        gdSetStr m.geAnnStruct.gshText, k, Abs(dX2 - dX1)
        k = k + 1
    End If
    
    'set font info
    nSize = gdGetSize(m.geAnnStruct.glhAnnType) - 1
    For i = 0 To nSize
        gdSetNum m.geAnnStruct.glhAlign, i, 10
        gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetFibTime", eGDRaiseError_Raise

End Sub

Private Function SetFibZones(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim eZone As eGBZoneInUse, nShowValues&
    
    nShowValues = Val(Prop("ShowValues"))
    
    'time zone: first line is placed at left-most mouse click
    '   subsequent lines are placed date-forward of first line
    '   at increasingly wider intervals in accordance with the
    '   fibonacci integer sequence until end of data
    If Val(Prop("Arcs")) = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 27
        m.geAnnStruct.showAxes = 0
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 0, 20
        m.geAnnStruct.showAxes = Val(Prop("ShowInAllPanes"))
    End If
    m.geAnnStruct.lenRatio = Val(Prop("ArcPercentHeight"))
    
    eZone = Val(Prop("ZoneInUse"))
    
    'option to draw lines from first bar or previous bar applies only to fib series
    'lines for all others zone types are drawn from first bar
    If eZone = eANNOT_GB_ZoneFib And Val(Prop("FromPrevBar")) = 1 Then
        m.geAnnStruct.minorAxisLenData = 1
    Else
        m.geAnnStruct.minorAxisLenData = 0
    End If
        
'JM 5/22/2008: implementation of fixed bars count for Greenblatt ratio, square root series
'   line numbers for fib, lucas and 144 series are calculated and set by grapheng.dll
'   line numbers for ratio, square root are set on VB side then passed into grapheng.dll
'   line numbers are stored in gshText field
    If m.eType = eANNOT_DanCodeZone Then
        m.geAnnStruct.showQtrLines = 4
        gdSetStr m.geAnnStruct.gshText, 0, kZoneDanCode
    ElseIf IsGreenBlattTool Then
        Select Case eZone
            Case eANNOT_GB_ZoneLucas
                m.geAnnStruct.showQtrLines = 1
            Case eANNOT_GB_Zone144
                m.geAnnStruct.showQtrLines = 2
            Case eANNOT_GB_ZoneRatio
                m.geAnnStruct.showQtrLines = 3
                gdSetStr m.geAnnStruct.gshText, 0, kZoneRatio
            Case eANNOT_GB_ZoneSqRoot
                m.geAnnStruct.showQtrLines = 3
                gdSetStr m.geAnnStruct.gshText, 0, kZoneSqRoot
            Case Else
                m.geAnnStruct.showQtrLines = 0          'fib numbers
        End Select
    Else
        If eZone <> eANNOT_GB_ZoneFib And eZone <> eANNOT_GB_ZoneLucas Then
            'this can happen if zone was saved with greenblatt module then reopened without greenblatt module
            eZone = eANNOT_GB_ZoneFib
            Prop("ZoneInUse") = eANNOT_GB_ZoneFib
        End If
        If eZone = eANNOT_GB_ZoneLucas Then
            m.geAnnStruct.showQtrLines = 1
        Else
            m.geAnnStruct.showQtrLines = 0
        End If
    End If
        
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
    'specified by a point
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhRight, 0, dX1
    'place holders
    gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
    gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
    gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
    gdSetNum m.geAnnStruct.glhAlign, 0, 3           'bottom left
    
    'text
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhRight, 1, dX1
    gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
    
    gdSetStr m.geAnnStruct.gshFont, 1, Prop("FontName")
    gdSetNum m.geAnnStruct.glhSize, 1, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 1, Val(Prop("FontStyle"))
    gdSetNum m.geAnnStruct.glhUnderline, 1, Val(Prop("FontUnderline"))
    gdSetNum m.geAnnStruct.glhAlign, 1, 3           'bottom left
    'set annotation type for showing fib numbers
    If 1 = nShowValues Then
        gdSetNum m.geAnnStruct.glhAnnType, 1, 0
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
    End If
    
    If m.eType = eANNOT_DanCodeZone Then
        gdSetNum m.geAnnStruct.glhPenColor, 1, vbRed
        gdSetNum m.geAnnStruct.glhPenColor, 2, vbBlack
        gdSetNum m.geAnnStruct.glhPenColor, 3, vbBlue
        gdSetNum m.geAnnStruct.glhPenColor, 4, 8388608          'dark blue
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetFibZones", eGDRaiseError_Raise

End Function

Private Function SetFibonacci(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim i&, j&, k&
    
    If m.aRatios.Size < 1 Then Exit Function
        
    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 1     'extension info
        gdSetNum m.geAnnStruct.glhAlign, 0, 0       'place holder
        gdSetNum m.geAnnStruct.gdhMisc, 0, 0        'place holder (extension)
        gdSetNum m.geAnnStruct.glhAnnType, 1, 1     'vertical line (ignored by fib time/time zone)
        gdSetNum m.geAnnStruct.glhAlign, 1, 0       'place hodler
        gdSetNum m.geAnnStruct.gdhMisc, 1, 0        'placehoder (extension)
        
        gdSetNum m.geAnnStruct.glhAnnType, 2, 1     'top line (fib price)/left line (fib time)
        gdSetNum m.geAnnStruct.glhAlign, 2, 0       'place holder
        gdSetNum m.geAnnStruct.glhAnnType, 3, 1     'bottom line (fib price)/right line (fib time)
        gdSetNum m.geAnnStruct.glhAlign, 3, 0       'place holder
        
        gdSetNum m.geAnnStruct.glhAnnType, 4, 0     'text
        gdSetNum m.geAnnStruct.glhAlign, 4, 3       'e_btmLeft - text top line (fib price)/text left line (fib time)
        gdSetNum m.geAnnStruct.glhAnnType, 5, 0     'text
        gdSetNum m.geAnnStruct.glhAlign, 5, 0       'e_topLeft - text bottom line (fib price)/text right line (fib time)
    End If
                
    'remove all text items
    j = m.aRatios.Size + 6
    gdSetSize m.geAnnStruct.glhAnnType, j, 0
    gdSetSize m.geAnnStruct.glhPenColor, j, 0
    gdSetSize m.geAnnStruct.glhPenStyle, j, 0
    gdSetSize m.geAnnStruct.glhPenSize, j, 0
    gdSetSize m.geAnnStruct.glhUnderline, j, 0
    gdSetSize m.geAnnStruct.glhStyle, j, 0
    gdSetSize m.geAnnStruct.glhAlign, j, 0
    gdSetSize m.geAnnStruct.glhSize, j, 0
    gdSetSize m.geAnnStruct.gdhTop, j, 0
    gdSetSize m.geAnnStruct.gdhLeft, j, 0
    gdSetSize m.geAnnStruct.gdhBottom, j, 0
    gdSetSize m.geAnnStruct.gdhRight, j, 0
    gdSetSize m.geAnnStruct.gdhMisc, j, 0
    gdSetSize m.geAnnStruct.gshText, j, 0
    gdSetSize m.geAnnStruct.gshFont, j, 0
    
    If m.eType = eANNOT_FibTimeRatio Then
        SetFibTime Chart, dX1, dX2
        Exit Function
    End If
                                
    If m.eType = eANNOT_FibFan Or m.eType = eANNOT_SpResistFan Then
        SetFibFan Chart, dX1, dX2
        Exit Function
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'All code below here is for Fib Retracement, Fib Expansion, Daniel Code Fib & Advanced Risk Reward
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim Y#, YDelta#, YLeft#, YRight#
    Dim dxLeft#, dxRight#, dxText#
    
    Dim nShowValues&, nColor&, nStyle&, nSize&
    
    Dim iShowZero&, iShowOne&
    Dim iBoldZero&, iBoldOne&
    
    Dim iPriceOnly&, iKeepAlive&, iExt&
    Dim iTextOnRight&, iAlign&, iDiagSave&
    
    Dim bReverse As Boolean
    Dim bLowerHigher As Boolean
        
    nSize = m.aRatios.Size
    
    'nShowValues = Val(Prop("ShowValues"))
    nColor = Val(Prop("FibColor"))
    nStyle = Val(Prop("FibStyle"))
    
    iShowZero = Val(Prop("ShowRatioZero"))
    iShowOne = Val(Prop("ShowRatioOne"))
    iBoldZero = Val(Prop("BoldRatioZero"))
    iBoldOne = Val(Prop("BoldRatioOne"))
    
    If Prop("FibTextRatio") = 0 And Prop("FibTextPrice") = 1 Then iPriceOnly = 1
    iKeepAlive = Val(Prop("KeepAlive"))
    iExt = Val(Prop("Ext"))
    
    '05-12-2012: obsolete flag (read in from file for backwards compatibility)
    iTextOnRight = Val(Prop("TextOnRight"))
    'originally vertical line, changed to diagonal 05-2012
    iDiagSave = Val(Prop("HideVerticalLine"))

    If (m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_AdvRiskReward) And Prop("ReverseDirection") <> 0 Then
        bReverse = True
    End If
    
    '06-13-2012 - set reserved flag to prevent drawing text twice when user draws fib with a vertical line
    m.geAnnStruct.reserved = 100

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Original design:
'   Show values property used to be just 0/1 for show or don't show values
'   Text were preset on left or right based on extension property 0/1/2/3
'       no ext (0)     --> set values on right, ratios on left
'       ext right (1)  --> set values on left, no ratios
'       ext left (2)   --> set values on right, no ratios
'       ext both (3)   --> no values or ratios
'       ext only (4)   --> only show right extensions (see note below)
'   Text on right property = 1 overides pre-set text location
'
'04-12-2012
'   added an extensions only option (Prop("Ext") == 4) that works like right extension,
'   but without horizontal lines from original draw points (i.e. where users clicked to draw fib)
'
'02-21-2007 design change to accomodate Elliot Wave request (seems better)
'   Show values property can be 0/1/2/3/4
'       0: no values or ratios
'       1: values on right, ratios on left
'       2: values & ratios on left
'       3: values & ratios on right
'       4: values in y-scale area (05-09-2012)
'   Extension property unchanged, but no longer used to determine text location
'   Text on right property not used except to convert show values property
'       after coversion set to -1 to indicate done and don't convert again
'   TextNextToMain new property 0/1 (applicable when have extensions)
'       0: draw at end of extensions
'       1: draw between extension and main line
'
'Backwards compatibility:
'If (Text on right <> -1) then
'convert show values to reflect original preset text locations
'   Show    Extension       Text on right           New Show Value
'    0      don't care        don't care                0
'    1      don't care            1                     3 (all text on right)
'    1       0 (none)             0                     1 (ratio left/values right)
'    1       1 (right)            0                     2 (all text on left)
'    1       2 (left)             0                     3 (all text on right)
'    1       3 (both)             0                     3
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'check to see if need conversion (backwards compatibility)
    'at one time, the only text option was to show ration left/values right or all on right
'    If iTextOnRight <> -1 Then
'        If nShowValues = 1 Then
'            If iTextOnRight = 1 Then
'                nShowValues = 3
'            ElseIf iExt = 1 Then
'                nShowValues = 2
'            ElseIf iExt = 2 Then
'                nShowValues = 3
'            End If
'        End If
'        Prop("ShowValues") = nShowValues
'        Prop("TextOnRight") = -1
'    End If
    
'    If nShowValues = 3 Then
'        If dX1 = kNullData Or dX2 = kNullData Then
'            gdSetNum m.geAnnStruct.glhAnnType, 0, 1     'reset (aardvark 6672)
'            m.geAnnStruct.showAxes = 0
'        Else
'            gdSetNum m.geAnnStruct.glhAnnType, 0, 3     '05-09-2012: values in y-scale area (E Wave request)
'            m.geAnnStruct.showAxes = 1
'        End If
'    Else
'        gdSetNum m.geAnnStruct.glhAnnType, 0, 1     'reset
'        m.geAnnStruct.showAxes = 0
'    End If
    
'Pete 's email (03-22-2012):
'   A keep live setting or “Dynamic: stay on last swing” setting would be nice.
'   Where it attachs to the highest high / lowest low since the tool was draw.
'
'Note: the highest high or lowest low within the specified range of bars
'   will not necessarily be higher or lower than the fib's right-most end point
'   i.e. user could draw right-most endpoint above, below or in the middle of a bar

    Dim dNewY#, nNewX&, nBar&, nNewBar&

    If iKeepAlive = 1 Then
        If m.X1 <> 0 And m.X2 <> 0 Then
            If Not m.Chart Is Nothing Then
                If Not m.Chart.Bars Is Nothing Then
                    If m.dDate1 < m.dDate2 Then
                        If m.Chart.Bars.IsIntraday Then
                            nBar = m.Chart.Bars.FindDateTime(m.dDate2)
                        Else
                            nBar = m.Chart.Bars.FindDateTime(Int(m.dDate2))
                        End If
                        nNewX = m.X2
                        
                        If m.Y2 > m.Y1 Then
                            bLowerHigher = HighestHighInBarRange(m.Chart.Bars, dNewY, nNewBar, nBar, m.Chart.Bars.Size - 1)
                            nNewX = GetX(m.Chart, m.Chart.Bars(eBARS_DateTime, nNewBar))
                            If nNewBar = nBar And m.Y2 < dNewY Then
                                'highest high in specified range is on bar user drew to
                                m.Y2 = dNewY
                            ElseIf dNewY < m.Y1 And dNewY < m.Y2 Then
                                'all fib lines are above highest high in specified range
                                bLowerHigher = LowestLowInBarRange(m.Chart.Bars, dNewY, nNewBar, nBar, m.Chart.Bars.Size - 1)
                                nNewX = GetX(m.Chart, m.Chart.Bars(eBARS_DateTime, nNewBar))
                                If nNewBar = nBar And m.Y2 <> dNewY Then
                                    If dNewY = m.Chart.Bars(eBARS_Low, nBar) Then m.Y2 = dNewY 'highest high is on current bar
                                End If
                            End If
                        Else
                            bLowerHigher = LowestLowInBarRange(m.Chart.Bars, dNewY, nNewBar, nBar, m.Chart.Bars.Size - 1)
                            nNewX = GetX(m.Chart, m.Chart.Bars(eBARS_DateTime, nNewBar))
                            If nNewBar = nBar And m.Y2 <> dNewY Then
                                'lowest low in specified range is on bar user drew to
                                If dNewY = m.Chart.Bars(eBARS_Low, nBar) Then m.Y2 = dNewY
                            ElseIf dNewY > m.Y1 And dNewY > m.Y2 Then
                                'all fib lines are below lowest low in specified range
                                bLowerHigher = HighestHighInBarRange(m.Chart.Bars, dNewY, nNewBar, nBar, m.Chart.Bars.Size - 1)
                                If nNewBar = nBar And m.Y2 < dNewY Then m.Y2 = dNewY
                            End If
                        End If
                        If bLowerHigher And m.X2 <> nNewX Then
                            m.Y2 = dNewY
                            m.X2 = nNewX
                            dX2 = m.X2 - m.Chart.ScreenStartX
                            m.dDate2 = m.Chart.Bars(eBARS_DateTime, nNewBar)
                        End If
                    Else
                        If m.Chart.Bars.IsIntraday Then
                            nBar = m.Chart.Bars.FindDateTime(m.dDate1)
                        Else
                            nBar = m.Chart.Bars.FindDateTime(Int(m.dDate1))
                        End If
                        nNewX = m.X1
                        
                        If m.Y1 > m.Y2 Then
                            bLowerHigher = HighestHighInBarRange(m.Chart.Bars, dNewY, nNewBar, nBar, m.Chart.Bars.Size - 1)
                            nNewX = GetX(m.Chart, m.Chart.Bars(eBARS_DateTime, nNewBar))
                            If nNewBar = nBar And m.Y2 < dNewY Then
                                m.Y2 = dNewY
                            ElseIf dNewY < m.Y1 And dNewY < m.Y2 Then
                                'all fib lines are above highest high in specified range
                                bLowerHigher = LowestLowInBarRange(m.Chart.Bars, dNewY, nNewBar, nBar, m.Chart.Bars.Size - 1)
                                nNewX = GetX(m.Chart, m.Chart.Bars(eBARS_DateTime, nNewBar))
                                If nNewBar = nBar And m.Y2 <> dNewY Then
                                    If dNewY = m.Chart.Bars(eBARS_Low, nBar) Then m.Y2 = dNewY 'highest high is on current bar
                                End If
                            End If
                        Else
                            bLowerHigher = LowestLowInBarRange(m.Chart.Bars, dNewY, nNewBar, nBar, m.Chart.Bars.Size - 1)
                            nNewX = GetX(m.Chart, m.Chart.Bars(eBARS_DateTime, nNewBar))
                            If nNewBar = nBar And m.Y2 > dNewY Then
                                 m.Y2 = dNewY
                            ElseIf dNewY > m.Y1 And dNewY > m.Y2 Then
                                'all fib lines are below lowest low in specified range
                                bLowerHigher = HighestHighInBarRange(m.Chart.Bars, dNewY, nNewBar, nBar, m.Chart.Bars.Size - 1)
                                nNewX = GetX(m.Chart, m.Chart.Bars(eBARS_DateTime, nNewBar))
                                If nNewBar = nBar And m.Y2 < dNewY Then m.Y1 = dNewY
                            End If
                        End If
                        If bLowerHigher And m.X1 <> nNewX Then
                            m.Y1 = dNewY
                            m.X1 = nNewX
                            dX1 = m.X1 - m.Chart.ScreenStartX
                            m.dDate1 = m.Chart.Bars(eBARS_DateTime, nNewBar)
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If dX1 < dX2 Then
        dxLeft = dX1
        dxRight = dX2
    Else
        dxLeft = dX2
        dxRight = dX1
    End If
    
    If m.eType = eANNOT_AdvRiskReward Then
        If g.strActiveDraw = "ID_AdvRiskReward" Then
            If Len(Prop("textZero")) = 0 Then Prop("textZero") = "E"
            If Len(Prop("textOne")) = 0 Then Prop("textOne") = "S"
        End If
        If bReverse Then
            YLeft = m.Y2
            YRight = m.Y1
        Else
            YLeft = m.Y1
            YRight = m.Y2
        End If
        m.bY1IsZeroRatio = False
    ElseIf bReverse Then
       If m.dDate1 < m.dDate2 Then
            YLeft = m.Y2
            YRight = m.Y1
            m.bY1IsZeroRatio = True
        Else
            YLeft = m.Y1
            YRight = m.Y2
            m.bY1IsZeroRatio = False
        End If
    Else
        If m.dDate1 < m.dDate2 Then
            YLeft = m.Y1
            YRight = m.Y2
            m.bY1IsZeroRatio = False
        Else
            YLeft = m.Y2
            YRight = m.Y1
            m.bY1IsZeroRatio = True
        End If
    End If

    'extension info (i.e. line style & color)
    geSetPenStyle Val(Prop("ExtStyle")), 0
    
    'top line
    gdSetNum m.geAnnStruct.glhAnnType, 2, 1
    gdSetNum m.geAnnStruct.glhPenColor, 2, m.nColor
    gdSetNum m.geAnnStruct.gdhTop, 2, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 2, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 2, m.Y1
    gdSetNum m.geAnnStruct.gdhRight, 2, dX2
    'bottom line
    gdSetNum m.geAnnStruct.glhAnnType, 3, 1
    gdSetNum m.geAnnStruct.glhPenColor, 3, m.nColor
    gdSetNum m.geAnnStruct.gdhTop, 3, m.Y2
    gdSetNum m.geAnnStruct.gdhLeft, 3, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 3, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 3, dX2
    
    'old design had text in these 2 index location (set to -1 to be ignored by grapheng.dll)
    gdSetNum m.geAnnStruct.glhAnnType, 4, -1
    gdSetNum m.geAnnStruct.glhAnnType, 5, -1
    
    'extension flags
    gdSetNum m.geAnnStruct.gdhMisc, 0, 0            '4299
    gdSetNum m.geAnnStruct.gdhMisc, 1, 0
    gdSetNum m.geAnnStruct.gdhMisc, 2, iExt
    gdSetNum m.geAnnStruct.gdhMisc, 3, iExt
    'flag to draw text at end of extensions or next to main lines
    m.geAnnStruct.showQtrLines = Val(Prop("TextNextToMain"))
        
    For i = 2 To 3
        gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
    Next

    If m.bY1IsZeroRatio Then        '4493
        If iShowZero <> 1 Then gdSetNum m.geAnnStruct.glhAnnType, 2, -1
        If iShowOne <> 1 Then gdSetNum m.geAnnStruct.glhAnnType, 3, -1
        
        If iBoldZero = 1 Then
            geSetPenStyle -4, 2
        Else
            geSetPenStyle m.nStyle, 2
        End If
    
        If iBoldOne = 1 Then
            geSetPenStyle -4, 3
        Else
            geSetPenStyle m.nStyle, 3
        End If
    Else
        If iShowZero <> 1 Then gdSetNum m.geAnnStruct.glhAnnType, 3, -1
        If iShowOne <> 1 Then gdSetNum m.geAnnStruct.glhAnnType, 2, -1
    
        If iBoldZero = 1 Then
            geSetPenStyle -4, 3
        Else
            geSetPenStyle m.nStyle, 3
        End If
    
        If iBoldOne = 1 Then
            geSetPenStyle -4, 2
        Else
            geSetPenStyle m.nStyle, 2
        End If
    End If

'JM (10-04-2010): As user click and drag - right-most point = 0, left-most point = 1
    'fib lines
    For i = 0 To nSize - 1
    
        If m.aRatioShow(i) = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, i + 6, 1
            
            If m.bUsingLogDrawingMode Then
                Y = Exp((Log(YLeft) - Log(YRight)) * m.aRatios(i) + Log(YRight))
            Else
                Y = (YLeft - YRight) * m.aRatios(i) + YRight
            End If
            
            If m.aRatioColor.Size > i Then
                gdSetNum m.geAnnStruct.glhPenColor, i + 6, m.aRatioColor(i)
            Else
                gdSetNum m.geAnnStruct.glhPenColor, i + 6, nColor
            End If
            
            If m.aRatioBold.Size > i Then
                If m.aRatioBold(i) = 1 Then
                    geSetPenStyle -4, i + 6
                Else
                    geSetPenStyle nStyle, i + 6
                End If
            Else
                geSetPenStyle nStyle, i + 6
            End If
            
            gdSetNum m.geAnnStruct.gdhTop, i + 6, Y
            gdSetNum m.geAnnStruct.gdhBottom, i + 6, Y
            gdSetNum m.geAnnStruct.gdhLeft, i + 6, dX1
            gdSetNum m.geAnnStruct.gdhRight, i + 6, dX2
            gdSetNum m.geAnnStruct.gdhMisc, i + 6, iExt     'extension flag
            gdSetStr m.geAnnStruct.gshText, i + 6, ""       'clear out text
        Else
            gdSetNum m.geAnnStruct.glhAnnType, i + 6, -1
        End If

    Next
    
    If iDiagSave = 1 Then
        If iExt = 0 And m.X1 = m.X2 And nShowValues = 0 Then
            'override diag hide if diag is perfectly vertical and there are no extension and no text
            gdSetNum m.geAnnStruct.glhAnnType, 1, 1
            'iDiagSave = 0      - don't do this(fix for 6728)
        Else
            'hide diag
            gdSetNum m.geAnnStruct.glhAnnType, 1, -1
        End If
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 1, 1
    End If
    Prop("HideVerticalLine") = Abs(iDiagSave)   'restore temporary -1 while moving to 1
    
    gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
    gdSetNum m.geAnnStruct.glhPenSize, 1, 1
    gdSetNum m.geAnnStruct.glhPenStyle, 1, 2        'pen style is always small dash
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 1, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 1, dX2
        
    'flag for whether and what kind of text to draw:
    '  -1: no text (use -1 for backwards compatibility, previous versions did not have this flag and was defaulted to 0)
    '   1: price only
    '   2: text & ratio
'    If nShowValues = 3 Then     'Or (Prop("FibTextRatio") = 0 And Prop("FibTextPrice") = 0) Then
'        m.geAnnStruct.minorAxisLenData = -1
'    ElseIf iPriceOnly = 1 Then
'        m.geAnnStruct.minorAxisLenData = 1
'    Else
'        m.geAnnStruct.minorAxisLenData = 0  '2
'    End If
    
    Dim s1$, s2$
    
    s1 = ""
    s2 = ""
    
    'text for top/bottom line (0/1) ratio
    If m.bY1IsZeroRatio Then
        i = 0
        j = 1
    Else
        i = 1
        j = 0
    End If
    
    'JM 11-06-2015: text string per fib line are pipe (|) delimited with ^R,^L or ^Y following values for left,right,y-scale locations
    If Prop("FibTextRatio") = 1 Then
        s1 = Str(i)
        s2 = Str(j)
        If Prop("FibTextRatioLoc") = 1 Then
            s1 = s1 & "^R"
            s2 = s2 & "^R"
        Else
            s1 = s1 & "^L"
            s2 = s2 & "^L"
        End If
    End If
    
    gdSetNum m.geAnnStruct.glhAnnType, 0, 1     'reset
    m.geAnnStruct.showAxes = 0
    If Prop("FibTextPrice") = 1 Then
        s1 = s1 & "|" & Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
        s2 = s2 & "|" & Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y2)
        If Prop("FibTextPriceLoc") = 0 Then
            s1 = s1 & "^L"
            s2 = s2 & "^L"
        ElseIf Prop("FibTextPriceLoc") = 2 Then
            s1 = s1 & "^Y"
            s2 = s2 & "^Y"
            gdSetNum m.geAnnStruct.glhAnnType, 0, 3     '05-09-2012: values in y-scale area (E Wave request)
            m.geAnnStruct.showAxes = 1
        Else
            s1 = s1 & "^R"
            s2 = s2 & "^R"
        End If
    End If
    
    If Prop("FibLabelPL") = 1 Then
        If Len(Prop("textZero")) > 0 Then
            If Prop("FibLabelPLLoc") = 0 Then
                s1 = s1 & "|" & Prop("textZero") & "^L"
            Else
                s1 = s1 & "|" & Prop("textZero") & "^R"
            End If
        End If
        If Len(Prop("textOne")) > 0 Then
            If Prop("FibLabelPLLoc") = 0 Then
                s2 = s2 & "|" & Prop("textOne") & "^L"
            Else
                s2 = s2 & "|" & Prop("textOne") & "^R"
            End If
        End If
    End If
    
    gdSetStr m.geAnnStruct.gshText, 2, s1
    gdSetStr m.geAnnStruct.gshText, 3, s2
    
    Dim aRatioText As New cGdArray
    Dim strPriceMove$, dAmt#
    
    If m.eType = eANNOT_AdvRiskReward Then aRatioText.SplitFields Prop("FibRatiosText"), ","
    
    'text for fib values & ratios
    nSize = m.aRatios.Size - 1
    k = gdGetSize(m.geAnnStruct.glhAnnType)
    For i = 0 To nSize
        If m.aRatioShow(i) = 1 Then
            k = i + 6
            
            Y = gdGetNum(m.geAnnStruct.gdhTop, k)
            
            If m.aRatioColor.Size > i Then nColor = m.aRatioColor(i)
            
            'text for ratios & y-values
            s1 = ""
            If Prop("FibTextRatio") = 1 Then
                s1 = m.aRatios(i)
                If Prop("FibTextRatioLoc") = 1 Then
                    s1 = s1 & "^R"
                Else
                    s1 = s1 & "^L"
                End If
            End If
            If Prop("FibTextPrice") = 1 Then
                s1 = s1 & "|" & Chart.PriceDisplay(m.geAnnStruct.paneId, Y)
                If Prop("FibTextPriceLoc") = 0 Then
                    s1 = s1 & "^L"
                ElseIf Prop("FibTextPriceLoc") = 2 Then
                    s1 = s1 & "^Y"
                Else
                    s1 = s1 & "^R"
                End If
            End If
            If Prop("FibTextPL") = 1 Then
                If m.bY1IsZeroRatio Then
                    If bReverse Then
                        If YLeft < YRight Then
                            dAmt = GetProfitLoss(Chart, Y, YLeft, strPriceMove)
                        Else
                            dAmt = GetProfitLoss(Chart, YLeft, Y, strPriceMove)
                        End If
                    Else
                        If YLeft < YRight Then
                            dAmt = GetProfitLoss(Chart, YLeft, Y, strPriceMove)
                        Else
                            dAmt = GetProfitLoss(Chart, Y, YLeft, strPriceMove)
                        End If
                    End If
                ElseIf bReverse Then
                    If YLeft < YRight Then
                        dAmt = GetProfitLoss(Chart, YLeft, Y, strPriceMove)
                    Else
                        dAmt = GetProfitLoss(Chart, Y, YLeft, strPriceMove)
                    End If
                Else
                    If YLeft < YRight Then
                        dAmt = GetProfitLoss(Chart, Y, YLeft, strPriceMove)
                    Else
                        dAmt = GetProfitLoss(Chart, YLeft, Y, strPriceMove)
                    End If
                End If
                If strPriceMove <> "INVALID" Then
                    If dAmt = Int(dAmt) Then
                        s2 = Format(dAmt, "$#,##0;($#,##0)")
                    Else
                        s2 = Format(dAmt, "$#,##0.00;($#,##0.00)")
                    End If
                    s1 = s1 & "|" & s2
                    If Prop("FibTextPLLoc") = 0 Then
                        s1 = s1 & "^L"
                    Else
                        s1 = s1 & "^R"
                    End If
                End If
            End If
            If Prop("FibLabelPL") = 1 And aRatioText.Size > i Then
                s1 = s1 & "|" & aRatioText(i)
                If Prop("FibLabelPLLoc") = 0 Then
                    s1 = s1 & "^L"
                Else
                    s1 = s1 & "^R"
                End If
            End If
            gdSetStr m.geAnnStruct.gshText, k, s1
            
            
            gdSetStr m.geAnnStruct.gshFont, k, Prop("FontName")
            gdSetNum m.geAnnStruct.glhSize, k, Val(Prop("FontSize"))
            gdSetNum m.geAnnStruct.glhStyle, k, Val(Prop("FontStyle"))
            gdSetNum m.geAnnStruct.glhUnderline, k, Val(Prop("FontUnderline"))
        End If
    Next
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetFibonacci", eGDRaiseError_Raise

End Function

Private Function SetGannLines(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim i&, nCount&, hArray&, nIdx&, nSecondHotSpot&
    Dim dRise#, dRun#
    Dim bSame As Boolean
    Dim bDateError As Boolean

    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 19
        gdSetNum m.geAnnStruct.glhFillPattern, 0, 0      'hollow circle for handles
        m.geAnnStruct.showAxes = Val(Prop("ShowAxes"))   'hidden flag for testing angle
    End If
           
    If m.dDate1 = m.dDate2 Then bSame = True    'And m.Y1 = m.Y2 Then bSame = True
    
    'reset individual rise/run for lines
    For i = 0 To 8
        gdSetNum m.geAnnStruct.glhAnnType, i, -1
        'set rise/run for 1x1 angle if drawn with a single click
        m.geAnnStruct.minorAxisLenData = 1
        'set rise over run (i.e. 1x2 .. 1x8 and 2x1 .. 8x1) place holders
        gdSetNum m.geAnnStruct.gdhBottom, i, 0      'rise
        gdSetNum m.geAnnStruct.gdhRight, i, 0
        gdSetStr m.geAnnStruct.gshText, i, ""
    Next
        
    'reset line(s) direction
    'Gann lines can point in 1 to 4 directions at one time
    'valid directions are:
    '   1=NE=up right, 7=NW=up left, 3=SE=down right, 5=SW=down left
    gdSetNum m.geAnnStruct.glhDirection, 0, 1
    For i = 0 To 3
        gdSetNum m.geAnnStruct.glhDirection, i, 0
    Next
    
    nSecondHotSpot = Prop("SecondHotSpot")
    nCount = 0
    'set direction(s) for lines
    If m.geAnnStruct.moveable = 0 Then
        'user moved from one quadrant to another
        Select Case nSecondHotSpot
            Case 1:
                If Val(Prop("DirNE")) = 0 Then m.geAnnStruct.lenRatio = 0
            Case 3:
                If Val(Prop("DirSE")) = 0 Then m.geAnnStruct.lenRatio = 0
            Case 5:
                If Val(Prop("DirSW")) = 0 Then m.geAnnStruct.lenRatio = 0
            Case 7:
                If Val(Prop("DirNW")) = 0 Then m.geAnnStruct.lenRatio = 0
        End Select
        
        If Val(Prop("DirNE")) = 0 Then
            If m.geAnnStruct.lenRatio = 1 Then m.geAnnStruct.lenRatio = 0
        Else
            gdSetNum m.geAnnStruct.glhDirection, nCount, 1
            nCount = nCount + 1
            If m.geAnnStruct.lenRatio = 0 Then m.geAnnStruct.lenRatio = 1
        End If
        If Val(Prop("DirNW")) = 0 Then
            If m.geAnnStruct.lenRatio = 7 Then m.geAnnStruct.lenRatio = 0
        Else
            gdSetNum m.geAnnStruct.glhDirection, nCount, 7
            nCount = nCount + 1
            If m.geAnnStruct.lenRatio = 0 Then m.geAnnStruct.lenRatio = 7
        End If
        If Val(Prop("DirSE")) = 0 Then
            If m.geAnnStruct.lenRatio = 3 Then m.geAnnStruct.lenRatio = 0
        Else
            gdSetNum m.geAnnStruct.glhDirection, nCount, 3
            nCount = nCount + 1
            If m.geAnnStruct.lenRatio = 0 Then m.geAnnStruct.lenRatio = 3
        End If
        If Val(Prop("DirSW")) = 0 Then
            If m.geAnnStruct.lenRatio = 5 Then m.geAnnStruct.lenRatio = 0
        Else
            gdSetNum m.geAnnStruct.glhDirection, nCount, 5
            nCount = nCount + 1
            If m.geAnnStruct.lenRatio = 0 Then m.geAnnStruct.lenRatio = 5
        End If
        
        If nCount < 1 Then Exit Function    'no valid direction, can't draw
        
        If bSame Then
            hArray = Chart.geDateArray
            If m.X1 = 0 Then m.X1 = GetX(Chart, m.dDate1)
            gdBinarySearch hArray, m.dDate1, nIdx, eGdSort_Default, 0, -1
            If m.dDate1 = gdGetNum(hArray, nIdx) Then
                nIdx = nIdx + m.geAnnStruct.minorAxisLenData
                m.dDate2 = gdGetNum(hArray, nIdx)
                If m.dDate2 > 0 Then
                    m.X2 = m.X1 + m.geAnnStruct.minorAxisLenData
                    dX2 = dX2 + m.geAnnStruct.minorAxisLenData
                    m.Y2 = m.Y1 + m.geAnnStruct.minorAxisLenData * Val(Prop("Rise"))
                    Prop("SecondHotSpot") = 1
                Else
                    bDateError = True
                End If
            Else
                bDateError = True
            End If
        End If
    
        If Val(Prop("DirNE")) = 2 Then Prop("DirNE") = 1
        If Val(Prop("DirSE")) = 2 Then Prop("DirSE") = 1
        If Val(Prop("DirSW")) = 2 Then Prop("DirSW") = 1
        If Val(Prop("DirNW")) = 2 Then Prop("DirNW") = 1
        
        Prop("SecondHotSpot") = m.geAnnStruct.lenRatio
    End If
    
    If bDateError Then
        'happens during initialization
        nCount = gdGetSize(m.geAnnStruct.glhAnnType) - 1
        For i = 0 To nCount
            gdSetNum m.geAnnStruct.glhAnnType, i, -1
        Next
        Exit Function
    End If
    
    If m.geAnnStruct.moveable = 1 Then          'Or m.geAnnStruct.lenRatio = 0 Then
        'lenRatio tells grapheng.dll which quadrant 2nd hot spot is in
        'moveable = 1 --> click-n-drag draw or move is in progress
        'lenRatio = 0 --> template just loaded or quadrant anchor was in got turned off
        Select Case nSecondHotSpot
            'clear previous setting
            Case 1:
                If Val(Prop("DirNE")) = 1 Then Prop("DirNE") = 0
            Case 3:
                If Val(Prop("DirSE")) = 1 Then Prop("DirSE") = 0
            Case 5:
                If Val(Prop("DirSW")) = 1 Then Prop("DirSW") = 0
            Case 7:
                If Val(Prop("DirNW")) = 1 Then Prop("DirNW") = 0
        End Select
        Select Case m.geAnnStruct.lenRatio
            'clear previous setting
            Case 1:
                If Val(Prop("DirNE")) = 2 Then Prop("DirNE") = 0
            Case 3:
                If Val(Prop("DirSE")) = 2 Then Prop("DirSE") = 0
            Case 5:
                If Val(Prop("DirSW")) = 2 Then Prop("DirSW") = 0
            Case 7:
                If Val(Prop("DirNW")) = 2 Then Prop("DirNW") = 0
        End Select
        If dX2 > dX1 And m.Y2 > m.Y1 Then
            gdSetNum m.geAnnStruct.glhDirection, 0, 1
            m.geAnnStruct.lenRatio = 1
            If Val(Prop("DirNE")) = 0 Then Prop("DirNE") = 2    '2 means this was not originally turned on
        ElseIf dX2 > dX1 And m.Y2 < m.Y1 Then
            gdSetNum m.geAnnStruct.glhDirection, 0, 3
            m.geAnnStruct.lenRatio = 3
            If Val(Prop("DirSE")) = 0 Then Prop("DirSE") = 2
        ElseIf dX2 < dX1 And m.Y2 < m.Y1 Then
            gdSetNum m.geAnnStruct.glhDirection, 0, 5
            m.geAnnStruct.lenRatio = 5
            If Val(Prop("DirSW")) = 0 Then Prop("DirSW") = 2
        ElseIf dX2 < dX1 And m.Y2 > m.Y1 Then
            gdSetNum m.geAnnStruct.glhDirection, 0, 7
            m.geAnnStruct.lenRatio = 7
            If Val(Prop("DirNW")) = 0 Then Prop("DirNW") = 2
        End If
        Prop("SecondHotSpot") = m.geAnnStruct.lenRatio
    End If
            
    'set pen color and style
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
    'set points per bar
    gdSetNum m.geAnnStruct.glhAnnType, 0, 19
    If m.Y1 <> m.Y2 And m.dDate1 <> m.dDate2 Then
        Prop("Run") = Abs(dX2 - dX1)
        Prop("Rise") = Abs(m.Y2 - m.Y1)
    End If
    
    dRise = Abs(Val(Prop("Rise")))
    dRun = Abs(Val(Prop("Run")))
    If dRun = 0 Then dRun = 1
    m.geAnnStruct.lenPoints = dRise / dRun
       
    'set data for 1x1 line
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
        
    'set rise/run for non-1x1 lines
    nCount = 1
    If Prop("GannFan") = 1 Then
        
        Dim nColor1x2&, nColor1x3&, nColor1x4&, nColor1x8&
        
        nColor1x2 = Int(Val(Prop("Color1x2")))
        nColor1x3 = Int(Val(Prop("Color1x3")))
        nColor1x4 = Int(Val(Prop("Color1x4")))
        nColor1x8 = Int(Val(Prop("Color1x8")))
        
        If Prop("1x2") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, nCount, 19
            gdSetNum m.geAnnStruct.gdhBottom, nCount, 1
            gdSetNum m.geAnnStruct.gdhRight, nCount, 2
            gdSetStr m.geAnnStruct.gshText, nCount, "1x2"
            gdSetNum m.geAnnStruct.glhPenColor, nCount, nColor1x2
            nCount = nCount + 1
        End If

        If Prop("1x3") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, nCount, 19
            gdSetNum m.geAnnStruct.gdhBottom, nCount, 1
            gdSetNum m.geAnnStruct.gdhRight, nCount, 3
            gdSetStr m.geAnnStruct.gshText, nCount, "1x3"
            gdSetNum m.geAnnStruct.glhPenColor, nCount, nColor1x3
            nCount = nCount + 1
        End If

        If Prop("1x4") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, nCount, 19
            gdSetNum m.geAnnStruct.gdhBottom, nCount, 1
            gdSetNum m.geAnnStruct.gdhRight, nCount, 4
            gdSetStr m.geAnnStruct.gshText, nCount, "1x4"
            gdSetNum m.geAnnStruct.glhPenColor, nCount, nColor1x4
            nCount = nCount + 1
        End If

        If Prop("1x8") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, nCount, 19
            gdSetNum m.geAnnStruct.gdhBottom, nCount, 1
            gdSetNum m.geAnnStruct.gdhRight, nCount, 8
            gdSetStr m.geAnnStruct.gshText, nCount, "1x8"
            gdSetNum m.geAnnStruct.glhPenColor, nCount, nColor1x8
            nCount = nCount + 1
        End If

        If Prop("2x1") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, nCount, 19
            gdSetNum m.geAnnStruct.gdhBottom, nCount, 2
            gdSetNum m.geAnnStruct.gdhRight, nCount, 1
            gdSetStr m.geAnnStruct.gshText, nCount, "2x1"
            gdSetNum m.geAnnStruct.glhPenColor, nCount, nColor1x2
            nCount = nCount + 1
        End If

        If Prop("3x1") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, nCount, 19
            gdSetNum m.geAnnStruct.gdhBottom, nCount, 3
            gdSetNum m.geAnnStruct.gdhRight, nCount, 1
            gdSetStr m.geAnnStruct.gshText, nCount, "3x1"
            gdSetNum m.geAnnStruct.glhPenColor, nCount, nColor1x3
            nCount = nCount + 1
        End If
    
        If Prop("4x1") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, nCount, 19
            gdSetNum m.geAnnStruct.gdhBottom, nCount, 4
            gdSetNum m.geAnnStruct.gdhRight, nCount, 1
            gdSetStr m.geAnnStruct.gshText, nCount, "4x1"
            gdSetNum m.geAnnStruct.glhPenColor, nCount, nColor1x4
            nCount = nCount + 1
        End If
    
        If Prop("8x1") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, nCount, 19
            gdSetNum m.geAnnStruct.gdhBottom, nCount, 8
            gdSetNum m.geAnnStruct.gdhRight, nCount, 1
            gdSetStr m.geAnnStruct.gshText, nCount, "8x1"
            gdSetNum m.geAnnStruct.glhPenColor, nCount, nColor1x8
            nCount = nCount + 1
        End If
    End If
        

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetGannLines", eGDRaiseError_Raise

End Function

Private Function SetHorzLine(Chart As cChart) As Long
On Error GoTo ErrSection:

    Dim Pane As cPane, Bars As cGdBars, Ind As cIndicator
    Dim dMinMove#, i&

    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 3
    End If
        
    ' round Y values to the nearest min move for price
    If UCase(m.strPane) = "PRICE PANE" Then
        Set Pane = Chart.Tree(m.strPane)
        If Not Pane Is Nothing Then
            Set Ind = Chart.Tree(Pane.idxPriceBars)
            If Not Ind Is Nothing Then
                Set Bars = Ind.Bars
            End If
        End If
        If Bars Is Nothing Then
            Set Bars = Chart.Bars
        End If
        If Not Bars Is Nothing Then
            dMinMove = Bars.MinMove(m.dDate1)
            If dMinMove > 0 Then
                m.Y1 = Int(m.Y1 / dMinMove + 0.5) * dMinMove
            End If
        End If
    End If
    
    'show value in y-axis
    m.geAnnStruct.showQtrLines = Val(Prop("ShowInAxis"))
    
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, 0
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhRight, 0, Chart.geChartPoints - 1
    
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
    '02-14-2007: currently average entry line is the only horz line with text
    If Len(m.strText) > 0 And m.eUsage <> eANNOT_UserAdded Then
        gdSetStr m.geAnnStruct.gshText, 0, m.strText
        gdSetNum m.geAnnStruct.glhAlign, 0, m.geTextAlign
        gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
    End If
        
    If m.oAlert Is Nothing Then
        If gdGetSize(m.geAnnStruct.glhAnnType) > 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, 1, -1
        End If
    Else
        SetAlertIcon 1, m.Y1
        If Prop("FakePriceAlert") = 1 Then
            m.geAnnStruct.skipHitTest = 1
            If InStr(m.oAlert.ChartCondition, "up to") <> 0 Then
                gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_North
            Else
                gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_South
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetHorzLine", eGDRaiseError_Raise

End Function

Private Function SetIcon(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim i&, iImage&, iFontSize&, strAsciiChar$
    Dim eDir As eImageDir
        
    m.dDate2 = m.dDate1             '6987, keep this at top, fixed Robert Kelley leak
    
    If m.eType = eANNOT_ElliotLabel Then
        If Len(m.strText) > 6 Then
            SetTextEdit Chart, dX1, dX2
            Exit Function
        ElseIf InStr(m.strText, "(") <> 0 Then
            Prop("ImageType") = eCNI_Ascii      'aardvark 6913
        End If
    End If
    
    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 7     'default to image only
        gdSetNum m.geAnnStruct.glhImageType, 0, 4   'default to square
        gdSetNum m.geAnnStruct.glhSize, 0, 12       'default to medium icon
        gdSetNum m.geAnnStruct.glhAlign, 0, 8       'default to ctrCtr
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0    'solid
        gdSetNum m.geAnnStruct.glhPenSize, 0, 1
        
        If m.eUsage <> eANNOT_PriceAlert Then
            gdSetNum m.geAnnStruct.glhPenStyle, 1, 0    'solid
            gdSetNum m.geAnnStruct.glhPenSize, 1, 1
            gdSetNum m.geAnnStruct.glhStyle, 0, 0 'place holder
            gdSetNum m.geAnnStruct.glhStyle, 1, 0 'place holder
            gdSetNum m.geAnnStruct.glhFillPattern, 0, 0 'place holder
            gdSetNum m.geAnnStruct.glhFillPattern, 1, 0 'place holder
        End If
    End If
        
    iImage = Val(Prop("ImageType"))
    iFontSize = Val(Prop("FontSize"))
    eDir = Val(Prop("ImageDir"))
    
    If m.eUsage = eANNOT_BidAskChart Then
        m.geAnnStruct.showAxes = 1                  'flag as chart's bid/ask icons
        For i = 0 To 1
            gdSetNum m.geAnnStruct.glhImageType, i, iImage
            gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("ImageSize"))
            gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("ImageStyle"))
            gdSetNum m.geAnnStruct.glhFillPattern, i, Val(Prop("ImageStyle"))
            
            If eDir = eCNI_East Or eDir = eCNI_West Then
                gdSetNum m.geAnnStruct.glhDirection, i, eDir
            Else
                gdSetNum m.geAnnStruct.glhDirection, 0, eCNI_North      '5508 (bid is always < ask)
                gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_South
            End If
        Next
        
        gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor     'bid
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
        
        gdSetNum m.geAnnStruct.glhPenColor, 1, Val(Prop("AskColor"))        'ask
        gdSetNum m.geAnnStruct.gdhLeft, 1, dX2
        gdSetNum m.geAnnStruct.gdhTop, 1, m.Y2
    ElseIf m.eUsage = eANNOT_PriceAlert Then
        If m.oAlert Is Nothing Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        ElseIf m.oAlert.ShowOnCharts Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, 7     'icon
            gdSetNum m.geAnnStruct.glhImageType, 0, Val(Prop("ImageType"))
            gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("ImageSize"))
            gdSetNum m.geAnnStruct.glhDirection, 0, Val(Prop("ImageDir"))
        
            gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1  'text if there is text else is icon
            gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
            gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
            gdSetNum m.geAnnStruct.gdhRight, 0, dX1
        Else
            gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        End If
    Else
        geTransIconTextStyle
        gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1  'text if there is text else is icon
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y1
        gdSetNum m.geAnnStruct.gdhRight, 0, dX1
        gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1  'icon if there is text else is ignored
        gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
        
        'font (reset for text only)
        If gdGetNum(m.geAnnStruct.glhAnnType, 0) = 0 Then
            i = Val(Prop("FontStyle"))
            gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
            gdSetNum m.geAnnStruct.glhSize, 0, iFontSize
            gdSetNum m.geAnnStruct.glhStyle, 0, i
            gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
            strAsciiChar = Left(Prop("AsciiChar"), 1)
            If Len(strAsciiChar) = 0 Then
                gdSetStr m.geAnnStruct.gshText, 0, m.strText
            ElseIf strAsciiChar = "." And Len(m.strText) > 0 Then
                ' TLB 12/18/2012: special case (for Brandt) ...
                ' if AscChar = "." along with a label, then just the label should show
                gdSetStr m.geAnnStruct.gshText, 0, " |" & m.strText
            Else
                gdSetStr m.geAnnStruct.gshText, 0, strAsciiChar & "|" & m.strText
            End If
            
            'set pensize to 2 so circle looks darker for elliot labels when text is bolded
            If m.eType = eANNOT_ElliotLabel Then
                If i = 1 Or i = 3 Then
                    gdSetNum m.geAnnStruct.glhPenSize, 0, 2
                    gdSetNum m.geAnnStruct.glhPenSize, 1, 2
                Else
                    gdSetNum m.geAnnStruct.glhPenSize, 0, 1
                    gdSetNum m.geAnnStruct.glhPenSize, 1, 1
                End If
                If iImage = eCNI_Circle Then gdSetNum m.geAnnStruct.glhSize, 0, iFontSize - 2
            End If
        End If
    End If
    

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetIcon", eGDRaiseError_Raise

End Function

Private Function SetIndicatorLabel(Chart As cChart) As Long
On Error GoTo ErrSection:

    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 13
        gdSetNum m.geAnnStruct.glhStyle, 0, 1       'bold for text
        gdSetNum m.geAnnStruct.glhPenSize, 0, 1     'pen size is always 1
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0    'pen style is always solid
    End If
    
    'grapheng.dll draws labels right->left on first row then left->right thereafter
    gdSetNum m.geAnnStruct.gdhTop, 0, -1
    gdSetNum m.geAnnStruct.gdhRight, 0, 1
    If m.eUsage = eANNOT_SideWinderLabel Then
        gdSetNum m.geAnnStruct.gdhBottom, 0, -1
        gdSetNum m.geAnnStruct.gdhLeft, 0, 10
        gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
    Else
        gdSetNum m.geAnnStruct.gdhLeft, 0, -1
        gdSetNum m.geAnnStruct.glhStyle, 0, g.ChartGlobals.nFontStyle
        gdSetNum m.geAnnStruct.glhSize, 0, g.ChartGlobals.nFontSize
        gdSetStr m.geAnnStruct.gshFont, 0, g.ChartGlobals.strFontName
    End If
    
    gdSetStr m.geAnnStruct.gshText, 0, m.strText
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetIndicatorLabel", eGDRaiseError_Raise

End Function

Private Function SetNotation(Chart As cChart) As Long
On Error GoTo ErrSection:

    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 12
        gdSetNum m.geAnnStruct.glhSize, 0, 0        'defaults to 8
        gdSetNum m.geAnnStruct.glhAlign, 0, 0       'defaults to center
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0    'solid
        gdSetNum m.geAnnStruct.glhPenSize, 0, 1
        gdSetStr m.geAnnStruct.gshText, 0, m.strText
        gdSetStr m.geAnnStruct.gshFont, 0, "courier"
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetNotation", eGDRaiseError_Raise

End Function

Private Function SetTextEdit(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim i&, nArrowSize&
    
    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 0     'text
        gdSetNum m.geAnnStruct.glhAnnType, 1, 1     'line for arrow
                        
        'initialize arrays
        For i = 0 To 1
            gdSetNum m.geAnnStruct.gdhTop, i, 0
            gdSetNum m.geAnnStruct.gdhLeft, i, 0
            gdSetNum m.geAnnStruct.gdhBottom, i, 0
            gdSetNum m.geAnnStruct.gdhRight, i, 0
            gdSetNum m.geAnnStruct.glhPenStyle, i, 0
            gdSetNum m.geAnnStruct.glhPenSize, i, 1
            gdSetNum m.geAnnStruct.glhDirection, i, 0
            gdSetNum m.geAnnStruct.glhStyle, i, 0
            gdSetNum m.geAnnStruct.glhSize, i, 0
            gdSetNum m.geAnnStruct.glhAlign, i, 0
            gdSetStr m.geAnnStruct.gshFont, i, "0"
            gdSetStr m.geAnnStruct.gshText, i, ""
            gdSetNum m.geAnnStruct.gdhMisc, i, 1
            gdSetNum m.geAnnStruct.glhJustify, i, 0
        Next
        
        gdSetNum m.geAnnStruct.gdhMisc, 1, 0        'reset to 0 for arrow line
    End If
    
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, Val(Prop("Height"))
    gdSetNum m.geAnnStruct.gdhRight, 0, Val(Prop("Width"))
    
    'JM 10-15-2013 - prevent random line that EWI analysts have run into but we cannot duplicate
    If m.eType = eANNOT_ElliotLabel Then
        Prop("ArrowStyle") = 0
    ElseIf m.eType = eANNOT_ArrowLine Then
        m.strText = ""
        If Val(Prop("ArrowStyle")) = 0 Then Prop("ArrowStyle") = 1
    End If
    
    'check whether to draw arrow
    If Prop("ArrowStyle") = 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
        gdSetNum m.geAnnStruct.glhDirection, 1, 0
    Else
        nArrowSize = Val(Prop("ArrowSize"))
        gdSetNum m.geAnnStruct.glhAnnType, 1, 1
        gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
        gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
        If dX1 = dX2 And m.Y1 = m.Y2 And m.geAnnStruct.moveable = 1 Then
            'this code should only excute for text-edit created using
            'single click with no text (aardvark 1052)
            If Chart.geChartPoints >= 30 Then
                m.dDate2 = gdGetNum(Chart.geDateArray, 10)
            Else
                m.dDate2 = gdGetNum(Chart.geDateArray, 2)
            End If
        Else
            gdSetNum m.geAnnStruct.glhAlign, 1, 0
            If m.dDate1 > m.dDate2 Then
                gdSetNum m.geAnnStruct.glhDirection, 1, 1   '1=use top/left, 2=use bottom/right, 3=arrows at both ends
            Else
                gdSetNum m.geAnnStruct.glhDirection, 1, 2   '1=use top/left, 2=use bottom/right, 3=arrows at both ends
            End If
        End If
        gdSetNum m.geAnnStruct.gdhBottom, 1, m.Y2
        gdSetNum m.geAnnStruct.gdhRight, 1, dX2
        gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
        'arrow size = arrow sides in pixels
        If nArrowSize = 1 Then
            gdSetNum m.geAnnStruct.glhSize, 1, 10  'medium
        ElseIf nArrowSize = 2 Then
            gdSetNum m.geAnnStruct.glhSize, 1, 14  'large
        Else
            gdSetNum m.geAnnStruct.glhSize, 1, 8    'small (default)
        End If
        gdSetNum m.geAnnStruct.glhStyle, 1, Val(Prop("ArrowStyle")) '0=standard, 1=triangle
        geSetPenStyle Val(Prop("ArrowLine")), 1
    End If
    'check whether to draw rectangular border
    If Prop("Border") = 0 Then
        gdSetNum m.geAnnStruct.glhImageType, 0, 0
    Else
        gdSetNum m.geAnnStruct.glhImageType, 0, 1
        gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
        geSetPenStyle m.nStyle, 0
    End If
        
    'text
    If Len(m.strText) < 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, -1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 0, 0
        gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
        gdSetNum m.geAnnStruct.glhAlign, 0, m.geTextAlign           'use for anchoring text
        gdSetNum m.geAnnStruct.glhJustify, 0, Val(Prop("TextJustify"))   'aardvark 4343
        
        gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        gdSetNum m.geAnnStruct.gdhBottom, 0, Val(Prop("Height"))
        gdSetNum m.geAnnStruct.gdhRight, 0, Val(Prop("Width"))
        
        gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
        gdSetStr m.geAnnStruct.gshText, 0, m.strText
    End If
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetTextEdit", eGDRaiseError_Raise
End Function

Private Function SetEllipse(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:
    
    Dim nAxisLenData&, i&
    
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 18    'twentyman ellipse
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0    'solid
        gdSetNum m.geAnnStruct.glhPenSize, 0, 1     'pen size 1
        
        gdSetNum m.geAnnStruct.glhStyle, 0, 0       'hollow shape
        gdSetNum m.geAnnStruct.glhSize, 0, 0        'place holder
    End If
    
    'major axis
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
        
    'minor axis
    nAxisLenData = Val(Prop("MinorAxisLenData"))
    If nAxisLenData = 0 Then
        m.geAnnStruct.lenRatio = Val(Prop("MinorAxisLen"))
    Else
        m.geAnnStruct.lenPoints = Val(Prop("MinorAxisLen"))
    End If
    m.geAnnStruct.minorAxisLenData = nAxisLenData
    m.geAnnStruct.showAxes = Val(Prop("ShowAxes"))
    m.geAnnStruct.showQtrLines = Val(Prop("ShowQtrLines"))
        
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetEllipse", eGDRaiseError_Raise

End Function

Private Function SetPattern(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&) As Long

    Dim i&
    Dim dIndMax#, dIndMin#, dTemp#

    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    '[0]=copy, [1]=original
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 30
        gdSetNum m.geAnnStruct.glhFillPattern, 0, 0 '0=transparent, 1=color using fill color
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0    'default to solid
        gdSetNum m.geAnnStruct.glhPenSize, 0, 1     'default to size 1
        gdSetNum m.geAnnStruct.glhStyle, 0, 0       'place holder
        gdSetNum m.geAnnStruct.glhSize, 0, 0        'place holder
        gdSetNum m.geAnnStruct.glhAlign, 0, 1       'place holder
        gdSetStr m.geAnnStruct.gshFont, 0, " "      'place holder
        gdSetStr m.geAnnStruct.gshText, 0, " "      'place holder
    End If

    If m.dDate1 > m.dDate2 Then
        dTemp = m.dDate1
        m.dDate1 = m.dDate2
        m.dDate2 = dTemp
    End If
    If dX1 > dX2 Then
        dTemp = dX1
        dX1 = dX2
        dX2 = dTemp
    End If
    
    If m.aDates.Size > 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 1, 30
        gdSetNum m.geAnnStruct.glhFillPattern, 1, 0 '0=transparent, 1=color using fill color
        gdSetNum m.geAnnStruct.glhPenStyle, 1, 0    'default to solid
        gdSetNum m.geAnnStruct.glhPenSize, 1, 1     'default to size 1
        gdSetNum m.geAnnStruct.glhStyle, 1, 0       'place holder
        gdSetNum m.geAnnStruct.glhSize, 1, 0        'place holder
        gdSetNum m.geAnnStruct.glhAlign, 1, 1       'place holder
        gdSetStr m.geAnnStruct.gshFont, 1, "0"      'place holder
        gdSetStr m.geAnnStruct.gshText, 1, "0"      'place holder
        
        gdSetNum m.geAnnStruct.gdhTop, 1, m.aYs(0)
        gdSetNum m.geAnnStruct.gdhLeft, 1, GetX(Chart, m.aDates(0)) - nStartScreenX     '4525
        gdSetNum m.geAnnStruct.gdhBottom, 1, m.aYs(1)
        gdSetNum m.geAnnStruct.gdhRight, 1, GetX(Chart, m.aDates(1)) - nStartScreenX    '4525
        
        gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
        geSetPenStyle m.nStyle, 1
        
        If Val(Prop("ShowPatternName")) = 1 Then
            gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
            gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
            gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
            gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
            
            gdSetStr m.geAnnStruct.gshText, 0, m.strText & "'"
            gdSetStr m.geAnnStruct.gshText, 1, m.strText
        Else
            gdSetStr m.geAnnStruct.gshText, 0, ""
            gdSetStr m.geAnnStruct.gshText, 1, ""
        End If
        'needed by grapheng.dll to handle off-screen originals
        Dim Ind As cIndicator
        Set Ind = Chart.Tree("PRICE")
        If Ind Is Nothing Then
            m.geAnnStruct.reserved = 0
        Else
            m.geAnnStruct.reserved = Ind.Bars.BarsHandle
        End If
        m.geAnnStruct.lenRatio = m.aDates(0)
        m.geAnnStruct.lenPoints = m.aDates(1)
        m.geAnnStruct.showAxes = Abs(Chart.ShowEmptyBars)
    ElseIf gdGetSize(m.geAnnStruct.glhAnnType) < 2 Then
        PatternMinMax Chart, dIndMin, dIndMax, m.dDate1, m.dDate2
        If dIndMin <> kNullData And dIndMax <> kNullData Then
            m.Y1 = dIndMin
            m.Y2 = dIndMax
        End If
    End If
    
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0

    m.geAnnStruct.minorAxisLenData = Val(Prop("ForecastBars"))
    gdSetStr m.geAnnStruct.gshText, 2, " "

End Function

Private Function SetBracket(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long

    Dim i&

    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 33
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0    'default to solid
        gdSetNum m.geAnnStruct.glhPenSize, 0, 1     'default to size 1
        gdSetNum m.geAnnStruct.glhStyle, 0, 0       '0=curly, 1=square
        gdSetNum m.geAnnStruct.glhDirection, 0, 0   '0=open left, 1=open right
    End If

    'aardvark 3031
    If dX1 < 0 Then
        If m.dDate1 > Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False)) Then
            dX1 = Chart.geChartPoints + 2
        End If
    ElseIf dX2 < 0 Then
        If m.dDate2 > Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False)) Then
            dX2 = Chart.geChartPoints + 2
        End If
    End If
    
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
    i = Val(Prop("BracketStyle"))
    gdSetNum m.geAnnStruct.glhStyle, 0, i
    
    i = Val(Prop("BracketDirection"))
    gdSetNum m.geAnnStruct.glhDirection, 0, i

End Function

Private Function SetRectangle(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim i&
    Dim Pane As cPane

    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        If m.eType = eANNOT_Mirror Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, 29
        Else
            geSetShape -1                           'sets annotation type to rect, roundrect or oval
        End If
        gdSetNum m.geAnnStruct.glhFillPattern, 0, 0 '0=transparent, 1=color using fill color
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0    'default to solid
        gdSetNum m.geAnnStruct.glhPenSize, 0, 1     'default to size 1
        gdSetNum m.geAnnStruct.glhStyle, 0, 0       'place holder
        gdSetNum m.geAnnStruct.glhSize, 0, 0        'place holder
        gdSetNum m.geAnnStruct.glhAlign, 0, 1       'place holder
        gdSetStr m.geAnnStruct.gshFont, 0, "0"      'place holder
        gdSetStr m.geAnnStruct.gshText, 0, "0"      'place holder
        
        gdSetNum m.geAnnStruct.glhAnnType, 1, 0     'text
        gdSetNum m.geAnnStruct.glhPenStyle, 1, 0    'always solid
        gdSetNum m.geAnnStruct.glhPenSize, 1, 1     'always size 1
        
        If m.eType = eANNOT_TextEdit Or m.eType = eANNOT_TextEdit2 Or m.eType = eANNOT_TextEdit3 Or m.eType = eANNOT_TextEdit4 Then
            gdSetNum m.geAnnStruct.gdhMisc, 0, 1    'fixed width/height
            gdSetNum m.geAnnStruct.gdhMisc, 1, 1    'wrap text
        End If
    End If
    
    If m.eUsage = eANNOT_FibClusters Then
        Set Pane = Chart.Tree("PRICE PANE")
        If Not Pane Is Nothing Then
            m.Y1 = Pane.Max + 1000
            m.Y2 = Pane.Min - 1000
        End If
        gdSetNum m.geAnnStruct.gdhMisc, 0, 3            '6408
        Set Pane = Nothing
    ElseIf m.eUsage = eANNOT_PatternProfit Then
'        If m.geAnnStruct.showAxes = 1 Then
'            'this is search pattern, do not center on bar
'            gdSetNum m.geAnnStruct.gdhMisc, 0, 2        '07-14-2010 roll back to: gdSetNum m.geAnnStruct.gdhMisc, 0, 0
'        Else
            'set flag for grapheng.dll to center rectangle sides data bar
            gdSetNum m.geAnnStruct.gdhMisc, 0, 2
'        End If
    End If
    
    'aardvark 3031
    If dX1 < 0 Then
        If m.dDate1 > Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False)) Then
            dX1 = Chart.geChartPoints + 2
        End If
    ElseIf dX2 < 0 Then
        If m.dDate2 > Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False)) Then
            dX2 = Chart.geChartPoints + 2
        End If
    End If
    
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    geSetShape -1
    
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
    gdSetNum m.geAnnStruct.glhFillColor, 0, Val(Prop("FillColor"))
    gdSetNum m.geAnnStruct.glhFillPattern, 0, Val(Prop("FillPattern"))  '0=transparent, 1=solid
    
    'text
    If Len(m.strText) < 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 1, 0
        If m.geTextAlign = 9 Then
            gdSetNum m.geAnnStruct.glhAlign, 1, 3
        Else
            gdSetNum m.geAnnStruct.glhAlign, 1, m.geTextAlign
        End If
        'gdSetNum m.geAnnStruct.gdhTop, 1, ((m.Y1 + m.Y2) / 2)
        'gdSetNum m.geAnnStruct.gdhLeft, 1, ((dX1 + dX2) / 2)
        gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
        gdSetNum m.geAnnStruct.gdhLeft, 1, dX1
        gdSetNum m.geAnnStruct.gdhBottom, 1, m.Y2
        gdSetNum m.geAnnStruct.gdhRight, 1, dX2
        gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
        'font(reset for text only)
        gdSetStr m.geAnnStruct.gshFont, 1, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, 1, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, 1, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, 1, Val(Prop("FontUnderline"))
        gdSetStr m.geAnnStruct.gshText, 1, m.strText
    End If
    
    If m.eType = eANNOT_Mirror And Not Chart Is Nothing Then
        'needed by grapheng.dll to handle off-screen reflection
        Dim Ind As cIndicator
        Set Ind = Chart.Tree("PRICE")
        If Ind Is Nothing Then
            m.geAnnStruct.reserved = 0
        Else
            m.geAnnStruct.reserved = Ind.Bars.BarsHandle
        End If
        m.geAnnStruct.lenRatio = m.dDate1
        m.geAnnStruct.lenPoints = m.dDate2
        m.geAnnStruct.showAxes = Abs(Chart.ShowEmptyBars)
    End If
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetRectangle", eGDRaiseError_Raise

End Function

Private Sub SetRegressionLine(Chart As cChart, ByVal dX1#, ByVal dX2#)
On Error GoTo ErrSection:

    Dim XDelta#, YDelta#, dFudge#
    Dim Ind As cIndicator
    Dim hIndData As Long    'array of indicator data to draw best fit line on
    Dim dSlope#, dIntercept#, dStdDev#, dTemp#
    Dim nKeepAtEnd&, nPriceField&
    Dim aAvgHiLow As New cGdArray
    Dim hIndHi&, hIndLow&, nSize&, i&
    Dim strKey$
    
    'get handle to visible data array
    If m.geAnnStruct.moveable = 1 Then
        strKey = Prop("IndicatorKey")
        If Len(strKey) > 0 Then Set Ind = Chart.Tree(strKey)
        'try indicator ID for backwards compatibility (aardvark 1736)
        If Ind Is Nothing And m.geAnnStruct.indicatorId > 0 Then
            If Chart.Tree.NodeLevel(m.geAnnStruct.indicatorId) > 0 Then
                Set Ind = Chart.Tree(m.geAnnStruct.indicatorId)
            End If
        End If
    Else
        Set Ind = Chart.Tree(m.geAnnStruct.indicatorId)
    End If
    If Ind Is Nothing Then
        Prop("IndicatorKey") = ""       'set so will not save to file
        m.geAnnStruct.indicatorId = 0
        Exit Sub                        'can't do anything else
    End If
    
    nKeepAtEnd = Val(Prop("KeepAtEnd"))
    nPriceField = Val(Prop("PriceField"))
       
    'show value in y-axis
    m.geAnnStruct.showQtrLines = Val(Prop("ShowInAxis"))
    
    'set dX2 to be rightmost x-value
    If dX2 < dX1 Then
        dTemp = dX2
        dX2 = dX1
        dX1 = dTemp
        m.Y2 = dTemp
        m.Y2 = m.Y1
        m.Y1 = dTemp
    End If
                
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 16    'regression line
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1    'main line right extension
        gdSetNum m.geAnnStruct.glhAnnType, 2, -1    'main line left extension
        gdSetNum m.geAnnStruct.glhAnnType, 3, -1    'std dev line 1
        gdSetNum m.geAnnStruct.glhAnnType, 4, -1    'std dev line 2
        gdSetNum m.geAnnStruct.glhAnnType, 5, -1    'std dev 1 right extension
        gdSetNum m.geAnnStruct.glhAnnType, 6, -1    'std dev 1 left extension
        gdSetNum m.geAnnStruct.glhAnnType, 7, -1    'std dev 2 right extension
        gdSetNum m.geAnnStruct.glhAnnType, 8, -1    'std dev 2 left extension
        'place holders
        For i = 0 To 8
            gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
            geSetPenStyle m.nStyle, i
        Next
    End If
    
    'number of channels
    m.geAnnStruct.showAxes = Val(Prop("ChannelCount"))
                    
    'check for keep at end option
    'aardvark 3070 fix - do this before calling MultiPointSetup
    If nKeepAtEnd > 0 And m.geAnnStruct.moveable = 1 Then
        SetDynamicEndPoints Chart, dX1, dX2
    End If
    
    'check if annotation's dates are within range of loaded data
    'aardvark 1852 - regression line hangs fix
    If MultiPointSetup(Chart, XDelta, YDelta, dFudge) = False Then
        For i = 0 To 8
            gdSetNum m.geAnnStruct.glhAnnType, i, -1
        Next
        Exit Sub
    End If
                
    Prop("IndicatorKey") = Chart.Tree.Key(Ind.geIndId)
    m.geAnnStruct.indicatorId = Ind.geIndId
    
    If UCase(Prop("IndicatorKey")) <> "PRICE" Then m.bMultiChart = False
        
    If Ind.Display = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, -1    'regression line
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 0, 16    'regression line
    End If
    
    If Ind.DataType = eINDIC_BarData And Ind.DisplayType < 0 Then
        If nPriceField < 4 Then
            hIndData = Ind.geDataHandle(nPriceField)
        Else
            hIndHi = Ind.geDataHandle(2)        'avg high low
            hIndLow = Ind.geDataHandle(3)
            nSize = gdGetSize(hIndHi)
            For i = 0 To nSize - 1
                aAvgHiLow.Add ((gdGetNum(hIndHi, i) + gdGetNum(hIndLow, i)) / 2)
            Next
            hIndData = aAvgHiLow.ArrayHandle
        End If
    Else
        hIndData = Ind.geDataHandle(1)
    End If
    
    'get best fit line
    Dim aExtendData As New cGdArray, dAvg#
    
    If dX1 < 0 Then
        For i = Abs(dX1) To 1 Step -1
            If Ind.DataType = eINDIC_BarData And Ind.DisplayType < 0 Then
                Select Case nPriceField
                    Case 0:     'close
                        aExtendData.Add Chart.Bars(eBARS_Close, Chart.aXBar(Chart.ScreenStartX - i))
                    Case 1:     'open
                        aExtendData.Add Chart.Bars(eBARS_Open, Chart.aXBar(Chart.ScreenStartX - i))
                    Case 2:     'hi
                        aExtendData.Add Chart.Bars(eBARS_High, Chart.aXBar(Chart.ScreenStartX - i))
                    Case 3:     'low
                        aExtendData.Add Chart.Bars(eBARS_Low, Chart.aXBar(Chart.ScreenStartX - i))
                    Case 4:
                        dAvg = (Chart.Bars(eBARS_High, Chart.aXBar(Chart.ScreenStartX - i)) + Chart.Bars(eBARS_High, Chart.aXBar(Chart.ScreenStartX - i))) / 2
                        aExtendData.Add dAvg
                    Case Else
                        Set Ind = Nothing
                        Set aAvgHiLow = Nothing
                        Set aExtendData = Nothing
                        Exit Sub
                End Select
            Else
                aExtendData.Add Ind.Data(i)
            End If
        Next
        For i = 0 To dX2
            aExtendData.Add gdGetNum(hIndData, i)
        Next
    ElseIf dX2 >= Chart.geChartPoints Then
        For i = dX1 To gdGetSize(hIndData) - 1
            aExtendData.Add gdGetNum(hIndData, i)
        Next
        For i = 1 To (dX2 - Chart.geChartPoints) + 1
            If Ind.DataType = eINDIC_BarData And Ind.DisplayType < 0 Then
                Select Case nPriceField
                    Case 0:     'close
                        aExtendData.Add Chart.Bars(eBARS_Close, Chart.aXBar(Chart.ScreenEndX + i))
                    Case 1:     'open
                        aExtendData.Add Chart.Bars(eBARS_Open, Chart.aXBar(Chart.ScreenEndX + i))
                    Case 2:     'hi
                        aExtendData.Add Chart.Bars(eBARS_High, Chart.aXBar(Chart.ScreenEndX + i))
                    Case 3:     'low
                        aExtendData.Add Chart.Bars(eBARS_Low, Chart.aXBar(Chart.ScreenEndX + i))
                    Case 4:
                        dAvg = (Chart.Bars(eBARS_High, Chart.aXBar(Chart.ScreenEndX + i)) + Chart.Bars(eBARS_Low, Chart.aXBar(Chart.ScreenEndX + i))) / 2
                        aExtendData.Add dAvg
                    Case Else:
                        Set Ind = Nothing
                        Set aAvgHiLow = Nothing
                        Set aExtendData = Nothing
                        Exit Sub
                End Select
            Else
                aExtendData.Add Ind.Data(i)
            End If
        Next
    End If

    Dim eUseLog As ePANE_LogFlag
    Dim Pane As cPane
    Set Pane = Chart.Tree(m.strPane)
    If Not Pane Is Nothing Then
        eUseLog = Pane.PaneLogFlag
        Set Pane = Nothing
    End If

    If (dX1 >= 0 And dX2 < Chart.geChartPoints) Then
        If eUseLog = ePANE_LogFlagLog Then
            aExtendData.Size = dX2
            For i = dX1 To dX2
                If gdGetNum(hIndData, i) > 0 Then
                    aExtendData.Num(i) = Log(gdGetNum(hIndData, i))
                Else
                    aExtendData.Num(i) = kNullData
                End If
            Next
            hIndData = aExtendData.ArrayHandle
        End If
        If hIndData > 0 Then
            gdBestFitLine hIndData, dSlope, dIntercept, dStdDev, dX1, dX2
        End If
        m.Y1 = dSlope * dX1 + dIntercept
        m.Y2 = dSlope * dX2 + dIntercept
    Else
        'intercept will be off-chart if dx1 < 0
        hIndData = aExtendData.ArrayHandle
        If eUseLog = ePANE_LogFlagLog Then
            For i = 0 To gdGetSize(hIndData)
                If gdGetNum(hIndData, i) > 0 Then
                    aExtendData.Num(i) = Log(gdGetNum(hIndData, i))
                Else
                    aExtendData.Num(i) = kNullData
                End If
            Next
        End If
        gdBestFitLine hIndData, dSlope, dIntercept, dStdDev, 0, gdGetSize(hIndData)
        m.Y1 = dIntercept
        m.Y2 = dSlope * (aExtendData.Size - 1) + dIntercept
    End If
    If eUseLog = ePANE_LogFlagLog Then
        m.Y1 = Exp(m.Y1)
        m.Y2 = Exp(m.Y2)
    End If
        
    Set Ind = Nothing
    Set aAvgHiLow = Nothing

    'set regression line properties
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    'set pen color
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    'set pen style
    geSetPenStyle m.nStyle, 0
    
    Dim dStd1Y1#, dStd1Y2#, dStd2Y1#, dStd2Y2#, dStdDevConst#
    
    'assume no standard deviation channel
    gdSetNum m.geAnnStruct.glhAnnType, 3, -1    'std dev line 1
    gdSetNum m.geAnnStruct.glhAnnType, 4, -1    'std dev line 2
    
    dStdDevConst = Val(Prop("StdDevVal"))
    'set standard dev properties
    If dStdDevConst > 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 3, 16
        gdSetNum m.geAnnStruct.glhAnnType, 4, 16

        'y-values stdDev lines 1 & 2
        If eUseLog = ePANE_LogFlagLog And m.Y1 > 0 And m.Y2 > 0 Then
            dStd1Y1 = Exp(Log(m.Y1) + dStdDevConst * dStdDev)
            dStd1Y2 = Exp(Log(m.Y2) + dStdDevConst * dStdDev)
            dStd2Y1 = Exp(Log(m.Y1) - dStdDevConst * dStdDev)
            dStd2Y2 = Exp(Log(m.Y2) - dStdDevConst * dStdDev)
        Else
            dStd1Y1 = m.Y1 + (dStdDevConst * dStdDev)
            dStd1Y2 = m.Y2 + (dStdDevConst * dStdDev)
            dStd2Y1 = m.Y1 - (dStdDevConst * dStdDev)
            dStd2Y2 = m.Y2 - (dStdDevConst * dStdDev)
        End If

        'properties stdDev line 1
        gdSetNum m.geAnnStruct.gdhTop, 3, dStd1Y1
        gdSetNum m.geAnnStruct.gdhLeft, 3, dX1
        gdSetNum m.geAnnStruct.gdhBottom, 3, dStd1Y2
        gdSetNum m.geAnnStruct.gdhRight, 3, dX2
        gdSetNum m.geAnnStruct.glhPenColor, 3, m.nColor     'Val(Prop("StdDevColor"))
        geSetPenStyle Val(Prop("StdDevStyle")), 3
        'properties stdDev line 2
        gdSetNum m.geAnnStruct.gdhTop, 4, dStd2Y1
        gdSetNum m.geAnnStruct.gdhLeft, 4, dX1
        gdSetNum m.geAnnStruct.gdhBottom, 4, dStd2Y2
        gdSetNum m.geAnnStruct.gdhRight, 4, dX2
        gdSetNum m.geAnnStruct.glhPenColor, 4, m.nColor     'Val(Prop("StdDevColor"))
        geSetPenStyle Val(Prop("StdDevStyle")), 4
    End If
    Set aExtendData = Nothing
    
    'set extension properties
    Dim iExt&, nExtColor&, nExtStyle&
    
    'assume no extension
    gdSetNum m.geAnnStruct.glhAnnType, 1, -1    'main line right extension
    gdSetNum m.geAnnStruct.glhAnnType, 2, -1    'main line left extension
    gdSetNum m.geAnnStruct.glhAnnType, 5, -1    'std dev line 1 right extension
    gdSetNum m.geAnnStruct.glhAnnType, 6, -1    'std dev line 1 left extension
    gdSetNum m.geAnnStruct.glhAnnType, 7, -1    'std dev line 2 right extension
    gdSetNum m.geAnnStruct.glhAnnType, 8, -1    'std dev line 2 left extension
    
    iExt = Val(Prop("Ext"))
    If iExt < 1 Then Exit Sub
    
    nExtColor = Val(Prop("ExtColor"))
    nExtStyle = Val(Prop("ExtStyle"))
    Select Case iExt
    Case 1 '(right)
        gdSetNum m.geAnnStruct.glhAnnType, 1, 16
        gdSetNum m.geAnnStruct.glhAnnType, 5, 16    'std dev line 1 right extension
        gdSetNum m.geAnnStruct.glhAnnType, 7, 16    'std dev line 2 right extension
    Case 2 '(left)
        gdSetNum m.geAnnStruct.glhAnnType, 2, 16
        gdSetNum m.geAnnStruct.glhAnnType, 6, 16    'std dev line 1 left extension
        gdSetNum m.geAnnStruct.glhAnnType, 8, 16    'std dev line 2 left extension
    Case 3 '(both)
        gdSetNum m.geAnnStruct.glhAnnType, 1, 16
        gdSetNum m.geAnnStruct.glhAnnType, 2, 16
        gdSetNum m.geAnnStruct.glhAnnType, 5, 16    'std dev line 1 right extension
        gdSetNum m.geAnnStruct.glhAnnType, 6, 16    'std dev line 1 left extension
        gdSetNum m.geAnnStruct.glhAnnType, 7, 16    'std dev line 2 right extension
        gdSetNum m.geAnnStruct.glhAnnType, 8, 16    'std dev line 2 left extension
    End Select
    'pen color
    gdSetNum m.geAnnStruct.glhPenColor, 1, nExtColor        'right extension
    gdSetNum m.geAnnStruct.glhPenColor, 2, nExtColor        'left extension
    For i = 5 To 8
        gdSetNum m.geAnnStruct.glhPenColor, i, nExtColor    'std dev extensions
    Next
    'pen style
    geSetPenStyle nExtStyle, 1      'right extension
    geSetPenStyle nExtStyle, 2      'left extension
    For i = 5 To 8
        geSetPenStyle nExtStyle, i  'std dev extensions
    Next
    'main line right extension
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y2
    gdSetNum m.geAnnStruct.gdhLeft, 1, dX2
    gdSetNum m.geAnnStruct.gdhRight, 1, Chart.geChartPoints + 1
    'main line left extension
    gdSetNum m.geAnnStruct.gdhBottom, 2, m.Y1
    gdSetNum m.geAnnStruct.gdhRight, 2, dX1
    gdSetNum m.geAnnStruct.gdhLeft, 2, -1
    'std dev right extensions
    gdSetNum m.geAnnStruct.gdhTop, 5, dStd1Y2
    gdSetNum m.geAnnStruct.gdhLeft, 5, dX2
    gdSetNum m.geAnnStruct.gdhRight, 5, Chart.geChartPoints + 1
    gdSetNum m.geAnnStruct.gdhTop, 7, dStd2Y2
    gdSetNum m.geAnnStruct.gdhLeft, 7, dX2
    gdSetNum m.geAnnStruct.gdhRight, 7, Chart.geChartPoints + 1
    'std dev left extensions
    gdSetNum m.geAnnStruct.gdhBottom, 6, dStd1Y1
    gdSetNum m.geAnnStruct.gdhRight, 6, dX1
    gdSetNum m.geAnnStruct.gdhLeft, 6, -1
    gdSetNum m.geAnnStruct.gdhBottom, 8, dStd2Y1
    gdSetNum m.geAnnStruct.gdhRight, 8, dX1
    gdSetNum m.geAnnStruct.gdhLeft, 8, -1
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetRegressionLine", eGDRaiseError_Raise

End Sub


Private Function SetTargetShooter(Chart As cChart) As Long
On Error GoTo ErrSection:

    Dim X#, Y#, XDelta#, YDelta#, dFudge#
    Dim X1#, X2#, Y1#, Y2#
    Dim tp1#, tp2#, tp3#, tp4#, tp5#, tp6#
    Dim bPullbacks As Boolean
    Dim i&
    
    
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 1     'dash diagonal line
        gdSetNum m.geAnnStruct.glhStyle, 0, 0       'place holder
        gdSetNum m.geAnnStruct.glhDirection, 0, 0   'place holder
                
        gdSetNum m.geAnnStruct.glhAnnType, 1, 6     'type ellipse
        gdSetNum m.geAnnStruct.glhStyle, 1, 0       'hollow circle
        gdSetNum m.geAnnStruct.glhFillPattern, 1, 0
        gdSetNum m.geAnnStruct.glhDirection, 1, 6   'e_west, shift left
        
        gdSetNum m.geAnnStruct.glhAnnType, 2, 6     'type ellipse
        gdSetNum m.geAnnStruct.glhStyle, 2, 0       'hollow circle
        gdSetNum m.geAnnStruct.glhFillPattern, 2, 0
        gdSetNum m.geAnnStruct.glhDirection, 2, 2   'e_east, shift right
        
        gdSetNum m.geAnnStruct.glhAnnType, 3, 1     'horz line
        
        gdSetNum m.geAnnStruct.glhAnnType, 4, 1     'vert line 1
        gdSetNum m.geAnnStruct.glhAnnType, 5, 1     'vert line 2
        
        gdSetNum m.geAnnStruct.glhAnnType, 6, 1     'target point line 1
        gdSetNum m.geAnnStruct.glhAnnType, 7, 1     'target point line 2
        gdSetNum m.geAnnStruct.glhAnnType, 8, 1     'target point line 3
        
        gdSetNum m.geAnnStruct.glhAnnType, 9, 1     'line point line 4
        gdSetNum m.geAnnStruct.glhAnnType, 10, 1     'line point line 5
        gdSetNum m.geAnnStruct.glhAnnType, 11, 1     'line point line 6
        
        'place holders
        For i = 0 To 11
            gdSetNum m.geAnnStruct.glhAlign, i, 0       'place holder
        Next
        
        'text for top 3 target points (12-14)
        'text for bottom 3 target points (15-17)
        For i = 12 To 17
            gdSetNum m.geAnnStruct.glhAnnType, i, 0     'text type
            gdSetNum m.geAnnStruct.glhAlign, i, 7       'e_ctrRight
        Next
        
        'defaults & place holders
        For i = 0 To 17
            gdSetNum m.geAnnStruct.glhSize, i, 0 '8    'font size
            gdSetStr m.geAnnStruct.gshFont, i, "arial"
            gdSetStr m.geAnnStruct.gshText, i, ""
        Next
    End If
    
    ' setup
    If MultiPointSetup(Chart, XDelta, YDelta, dFudge) = False Then
        SetTargetShooter = 0
        Exit Function
    End If
            
    If Val(Prop("Pullbacks")) <> 0 Then bPullbacks = True
    
    ' Diagonal line (keep these points tied to m.X1
    ' and m.X2 regardless of which one is first)
    ' end pt
    X = m.X2
    If X = m.X1 Then X = X + kXfudge
    Y = m.Y2
    If Y = m.Y1 Then Y = Y + dFudge
            
    ' assign X1 and X2 such that X1 is first (earlier)
    If m.X1 <= m.X2 Then
        X1 = m.X1
        Y1 = m.Y1
        X2 = m.X2
        Y2 = m.Y2
    Else
        X1 = m.X2
        Y1 = m.Y2
        X2 = m.X1
        Y2 = m.Y1
    End If

    ' use FIBONACCI ratios to calculate target pts
    If bPullbacks Then
        ' pullbacks
        tp1 = Y2 + (Y1 - Y2) * (1# - kFibRatio)
        tp2 = Y2 + (Y1 - Y2) * 0.5
        tp3 = Y2 + (Y1 - Y2) * kFibRatio
    Else
        ' targets
        tp1 = Y2 + (Y1 - Y2) / kFibRatio  ' = 31.72 degrees
        tp2 = Y2 + (Y1 - Y2) / 0.5      ' = 45 degrees
        tp3 = Y2 + (Y1 - Y2) / (1 - kFibRatio) ' = 58.28 degrees
        ' when vertical, targets in both directions
        tp4 = Y1 - (Y1 - Y2) / kFibRatio
        tp5 = Y1 - (Y1 - Y2) / 0.5
        tp6 = Y1 - (Y1 - Y2) / (1# - kFibRatio)
        If Val(Prop("Pt2")) = 0 Then
            tp2 = tp1
            tp5 = tp4
        End If
        If Val(Prop("Pt3")) = 0 Then
            tp3 = tp2
            tp6 = tp5
        End If
    End If
    
    'font (reset for text only)
    For i = 12 To 17
        gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
    Next
    
    gdSetNum m.geAnnStruct.gdhTop, 6, tp1
    gdSetNum m.geAnnStruct.gdhTop, 7, tp2
    gdSetNum m.geAnnStruct.gdhTop, 8, tp3
    gdSetNum m.geAnnStruct.gdhTop, 9, tp4
    gdSetNum m.geAnnStruct.gdhTop, 10, tp5
    gdSetNum m.geAnnStruct.gdhTop, 11, tp6

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetTargetShooter", eGDRaiseError_Raise

End Function

Private Function SetTimeCycle(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    'time zone: first line is placed at left-most mouse click
    '   subsequent lines are spaced on either side at equal
    '   bars intervals specified by user
    
    'setup info:
    '   glhsize[0] holds number of bars between lines
    '   showAxes field is flag for showing in all panes or not
    '   showQtrLines field contains index of first line to right of base line (set by grapheng.dll)
    '   lenRatio field contains height of arc as percent of pane's height
    
    If Val(Prop("Arcs")) = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 28
        m.geAnnStruct.showAxes = 0
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 0, 22
        m.geAnnStruct.showAxes = Val(Prop("ShowInAllPanes"))
    End If
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    'specified by a point
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    m.geAnnStruct.lenRatio = Val(Prop("ArcPercentHeight"))
    gdSetNum m.geAnnStruct.glhSize, 0, Abs(dX2 - dX1)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetTargetShooter", eGDRaiseError_Raise

End Function

Private Function SetTrades(Chart As cChart) As Long
On Error GoTo ErrSection:

    Dim iDir%

    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 7     'image
        gdSetNum m.geAnnStruct.glhImageType, 0, 6   'triangle
        gdSetNum m.geAnnStruct.glhSize, 0, -60      'percent of datapoint width
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0    'solid pen
        gdSetNum m.geAnnStruct.glhPenSize, 0, 1
        gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
        gdSetStr m.geAnnStruct.gshFont, 0, ""
        gdSetStr m.geAnnStruct.gshText, 0, ""
        gdSetStr m.geAnnStruct.glhAlign, 0, 7       '6=left, 7=right .. and more
               
        If Len(m.strText) > 0 Then
            gdSetNum m.geAnnStruct.glhAnnType, 1, 0     'text
            
            'aardvark 6761: using same font size as indicator labels (but not bold like indicator labels)
            'indicator labels are hard-coded to be bold in grapheng.dll
            gdSetStr m.geAnnStruct.gshFont, 1, g.ChartGlobals.strFontName
            gdSetNum m.geAnnStruct.glhSize, 1, g.ChartGlobals.nFontSize        'font size in points
            gdSetNum m.geAnnStruct.glhStyle, 1, 0
            
            gdSetNum m.geAnnStruct.glhPenStyle, 1, 0    'solid pen
            gdSetNum m.geAnnStruct.glhPenSize, 1, 1
            gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
            gdSetStr m.geAnnStruct.gshText, 1, m.strText
            gdSetNum m.geAnnStruct.glhAlign, 1, 6       '6=left, 7=right .. and more
        End If
        
        ' set direction of triangle: 2=east, 6=west
        ' (orders are flagged by setting dY2 to -1)
        If m.Y2 < 0 Then
            iDir = 6 ' (east for orders)
        Else
            iDir = 2 ' (west for trades)
        End If
        
        'set triangle image property
        Select Case m.nStyle
            Case PEGAT_SMALLUPTRIANGLE:
                gdSetNum m.geAnnStruct.glhDirection, 0, iDir
                gdSetNum m.geAnnStruct.glhStyle, 0, 0       'hollow
                gdSetNum m.geAnnStruct.glhAlign, 0, 7
                
            Case PEGAT_SMALLUPTRIANGLESOLID:
                gdSetNum m.geAnnStruct.glhDirection, 0, iDir
                gdSetNum m.geAnnStruct.glhStyle, 0, 1       'solid
                gdSetNum m.geAnnStruct.glhAlign, 0, 7
                
            Case PEGAT_SMALLDOWNTRIANGLE:
                gdSetNum m.geAnnStruct.glhDirection, 0, iDir
                gdSetNum m.geAnnStruct.glhStyle, 0, 0       'hollow
                gdSetNum m.geAnnStruct.glhAlign, 0, 7
                
            Case PEGAT_SMALLDOWNTRIANGLESOLID:
                gdSetNum m.geAnnStruct.glhDirection, 0, iDir
                gdSetNum m.geAnnStruct.glhStyle, 0, 1       'solid
                gdSetNum m.geAnnStruct.glhAlign, 0, 7
                
            Case Else
                ' if no triangle, e.g. "(no orders)", set 1st
                ' item to blank text to eliminate the triangle
                gdSetNum m.geAnnStruct.glhAnnType, 0, 0     'text
        End Select
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetTrades", eGDRaiseError_Raise

End Function

Private Function SetVertLine(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim i&, dDateTime#, strDateTime$

    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 2
    End If

    'show value in x-axis
    i = Val(Prop("ShowInAxis"))
    If m.eUsage = eANNOT_OptionInfo Then
        m.geAnnStruct.showQtrLines = 2
        gdSetStr m.geAnnStruct.gshFont, 0, "Arial"
        gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, 0, 1         'bold
        gdSetStr m.geAnnStruct.gshText, 0, DateFormat(m.dDate1, MM_DD_YYYY)
    ElseIf i = 1 Then
        m.geAnnStruct.showQtrLines = 1
        If Chart.Bars.IsIntraday Then
            dDateTime = gdGetNum(Chart.geDateArray, dX1)
            If g.bShowInLocalTimeZone Then
                dDateTime = ConvertTimeZone(dDateTime, m.Chart.Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            End If
            gdSetStr m.geAnnStruct.gshText, 0, DateFormat(dDateTime, MM_DD_YY, HH_MM)
        Else
            gdSetStr m.geAnnStruct.gshText, 0, DateFormat(m.dDate1, MM_DD_YY)
        End If
    Else
        m.geAnnStruct.showQtrLines = 0          '6419
    End If
    
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX1
    
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetVertLine", eGDRaiseError_Raise

End Function

Private Function InitDinapoli(Chart As cChart, ByVal eType As eAnnotType) As Boolean
On Error GoTo ErrSection:

    Dim nHiLo&, nDynamic
    
    InitDinapoli = False
    
    nHiLo = Chart.Form.FocusHiLow
    
    If nHiLo < 1 Or nHiLo = 3 Then
        Exit Function    'precautionary (should never happen)
    End If
    
    m.dDate2 = -1
    m.Y2 = -1
    'set dynamic variable
    If nHiLo = 4 Or nHiLo = 5 Then
        nDynamic = 1
    End If
    'set hilo variable
    If nHiLo = 2 Or nHiLo = 5 Then
        nHiLo = 1   'low
    Else
        nHiLo = 0
    End If

    gdSetNum m.geAnnStruct.glhDirection, 0, nHiLo
    
    If eType = eANNOT_DNRetracement Then
        Prop("FocusLow") = nHiLo
        Prop("FocusDynamic") = nDynamic
    Else
        Prop("HiLo") = nHiLo
    End If
        
    'SetDNRBoundary Chart
    InitDinapoli = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.InitDinapoli", eGDRaiseError_Raise

End Function

' The user is moving one of the annotation points
Public Function MovePoint(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#, Optional nFlag As Long) As Boolean
On Error GoTo ErrSection:

    Dim X1#, Y1#, X2#, Y2#, bMoved As Boolean
        
    If nPoint < 0 Then Exit Function
        
    'save old points
    X1 = m.X1
    X2 = m.X2
    Y1 = m.Y1
    Y2 = m.Y2
    Select Case m.eType
        Case eANNOT_VertLine
            bMoved = MoveVertLine(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4
            bMoved = MoveHorzLine(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4, _
             eANNOT_GannacciSwingSquare, eANNOT_FibArcs
            bMoved = MoveDollarLine(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_Icon, eANNOT_ElliotLabel
            bMoved = MoveIcon(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_TimeCycle
            bMoved = MoveTimeCycle(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_FibFan, eANNOT_SpResistFan, eANNOT_ArrowLine, _
             eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4
            bMoved = MoveTrendline(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
            eANNOT_Fibonacci4, eANNOT_FibTimeRatio, eANNOT_AdvRiskReward, _
            eANNOT_FibExpansion, eANNOT_DanCodeFib
            bMoved = MoveFibonacci(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_FibTimeZones, eANNOT_DanCodeZone
            bMoved = MoveFibTimeZones(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_TargetShooter
            bMoved = MoveTargetShooter(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_Rectangle, eANNOT_Bracket
            bMoved = MoveRectangle(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_RegressionLine
            bMoved = MoveRegressionLine(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
             eANNOT_DNExpansion4, eANNOT_FibABCD
            bMoved = MoveDnExpansion(Chart, nPoint, X, Y)
        Case eANNOT_Ellipse
            If nPoint = 4 Then
                bMoved = MoveEllipse(Chart, nPoint, nPaneID, X, Y)
            Else
                bMoved = MoveTrendline(Chart, nPoint, nPaneID, X, Y)
            End If
        Case eANNOT_DNRetracement
            'nFlag is set to 1 if point is within dnrboundary
            bMoved = MoveDnRetracement(Chart, nPoint, X, Y, nFlag)
        Case eANNOT_GannLines
            bMoved = MoveGannLines(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_AndrewFork, eANNOT_TriangleWedge, eANNOT_ElliotTimeRatio
            bMoved = MoveAndrewFork(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
            bMoved = MoveSRLine(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
             eANNOT_TrendChannel
            bMoved = MoveTrendChannel(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_Mirror
            bMoved = MoveMirror(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_Pattern
            bMoved = MovePattern(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_RiskReward
            bMoved = MoveRiskReward(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_ChannelHighlight
            bMoved = MoveChannelHighlight(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_WaveLabels
            bMoved = MoveWaveLabels(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_Pivot
            bMoved = MovePivot(Chart, nPoint, nPaneID, X, Y)
        Case eANNOT_Gartley
            bMoved = MoveGartley(Chart, nPoint, X, Y)
        Case eANNOT_GannacciCycle
            bMoved = False
        Case eANNOT_GannacciTime
            If nPoint = 0 Then
                bMoved = MoveVertLine(Chart, nPoint, nPaneID, X, Y)
            Else
                bMoved = False
            End If
        Case eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
            bMoved = MoveGannacciSwing(Chart, nPoint, X, Y)
        Case eANNOT_BalloonStrangle
            bMoved = MoveBalloonStrangle(Chart, nPoint, nPaneID, X, Y)
    End Select
    
    If eType = eANNOT_DNExpansion Or eType = eANNOT_DNExpansion2 Or eType = eANNOT_DNExpansion3 _
        Or eType = eANNOT_DNExpansion4 Or eType = eANNOT_DNRetracement _
        Or eType = eANNOT_AndrewFork Or m.eType = eANNOT_RiskReward _
        Or eType = eANNOT_TriangleWedge Or eType = eANNOT_ChannelHighlight _
        Or eType = eANNOT_WaveLabels Or eType = eANNOT_ElliotTimeRatio _
        Or eType = eANNOT_FibABCD Or eType = eANNOT_Gartley _
        Or eType = eANNOT_GannacciSwing1 Or eType = eANNOT_GannacciSwing2 Then
        MovePoint = bMoved
        Exit Function
    ElseIf (eType = eANNOT_TimeCycle Or eType = eANNOT_FibTimeZones Or eType = eANNOT_DanCodeZone) And nPoint = 1 Then
        MovePoint = bMoved
        Exit Function
    ElseIf eType = eANNOT_Pattern And (nPoint = 7 Or nPoint = 3) Then
        MovePoint = bMoved
        Exit Function
    ElseIf m.eType = eANNOT_SRLine Or m.eType = eANNOT_SRLine2 Or m.eType = eANNOT_SRLine3 Or _
           m.eType = eANNOT_SRLine4 Or m.eType = eANNOT_BalloonStrangle Then
        'do nothing (4264)
    ElseIf m.eType <> eANNOT_Trendline And m.eType <> eANNOT_Trendline2 And _
           m.eType <> eANNOT_Trendline3 And m.eType <> eANNOT_Trendline4 And _
           m.eType <> eANNOT_TrendChannel Then
        ' see if really moved or not
        ' when target lines for targetshooter are moved,
        ' only the multiplier changes so don't do check
        If m.eType <> eANNOT_TargetShooter And nPoint <> 10 Then
            If m.eType <> eANNOT_Ellipse And nPoint <> 4 Then
                If m.X1 = X1 And m.X2 = X2 And m.Y1 = Y1 And m.Y2 = Y2 Then
                    bMoved = False
                End If
            End If
        End If
    End If
        
    If bMoved Then
        ' reset new points
        If m.dDate1 <> 0 Then m.dDate1 = Chart.aXdate(m.X1)
        If m.dDate2 <> 0 Then m.dDate2 = Chart.aXdate(m.X2)
    Else
        ' restore old points
        m.X1 = X1
        m.X2 = X2
        m.Y1 = Y1
        m.Y2 = Y2
    End If
    
    MovePoint = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MovePoint", eGDRaiseError_Raise

End Function

' The user is moving one of the annotation points
Private Function MoveVertLine(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    
    If m.eUsage <> eANNOT_UserAdded Then Exit Function
    
    bMoved = True
    m.X1 = RoundNum(X)
    m.Y1 = 0
    
    If m.eType = eANNOT_VertLine Then m.geAnnStruct.paneId = 0    'vert lines are not tied to a pane
    
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
    End If
    
    MoveVertLine = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveVertLine", eGDRaiseError_Raise

End Function

' The user is moving one of the annotation points
Private Function MoveGannLines(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim XDelta#, YDelta#
    Dim bMoved As Boolean
    
    If nPaneID <> m.geAnnStruct.paneId Then Exit Function
    
    bMoved = True
    XDelta = m.X2 - m.X1
    If m.bUsingLogDrawingMode Then
        YDelta = LogModeDelta(m.Y1, m.Y2)       'Log(m.Y2) - Log(m.Y1)
    Else
        YDelta = m.Y2 - m.Y1
    End If
    
    Select Case nPoint
        Case 0, 1:
            'on line - move whole annotation
            m.X1 = X - m.dMoveOffSetX
            m.X2 = m.X1 + XDelta
            
            If m.bUsingLogDrawingMode Then
                m.Y1 = Exp(Log(Y) - m.dMoveOffSetLogY)
                m.Y2 = Exp(Log(m.Y1) + YDelta)
            Else
                m.Y1 = Y - m.dMoveOffSetY
                m.Y2 = m.Y1 + YDelta
            End If
        
        Case 2:
            'other point
            m.geAnnStruct.moveable = 1
            m.X2 = RoundNum(X)
            If m.X2 = m.X1 Then
                bMoved = False
            Else
                m.Y2 = Y
            End If
        Case Else:
            bMoved = False
    End Select
    
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
    End If
    
    MoveGannLines = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveGannLines", eGDRaiseError_Raise

End Function

' The user is moving one of the annotation points
Private Function MoveHorzLine(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:
    
    If m.eUsage <> eANNOT_UserAdded And m.eUsage <> eANNOT_Indicator Then
        Exit Function
    End If
    
    m.X1 = 0
    m.Y1 = RoundedY(Chart, Y)
    m.geAnnStruct.paneId = nPaneID
    
    m.strPane = Chart.Tree.Key(nPaneID)
    
    MoveHorzLine = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveHorzLine", eGDRaiseError_Raise

End Function

' The user is moving one of the annotation points
Private Function MoveIcon(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim rc&, nPaneIdSave&
    Dim aFlds As cGdArray
    
    If m.eType = eANNOT_ElliotLabel Then
        If Not AllowEWI And Not AllowEWIGMP And Not AllowGMP Then   'aardvark 6939
            If IsEndUserEWI And HasModule("EWL") Then
                'allow moving
            Else
                Exit Function
            End If
        End If
    End If
    
    If m.eUsage = eANNOT_Trades Then
        If Chart.Form.IsInGameMode Then
            Set aFlds = New cGdArray
            aFlds.SplitFields m.strText
            aFlds(1) = Chart.Bars.PriceDisplay(Y)
            m.strText = aFlds.JoinFields(" ")
            gdSetStr m.geAnnStruct.gshText, 1, m.strText
        Else
            Exit Function
        End If
    ElseIf m.eUsage = eANNOT_PriceAlert Then
        If m.oAlert.UseGetsUpTo And m.oAlert.UseGetsDownTo Then
            If m.nHitItemIdx = 0 Then
                m.Y1 = Y
                gdSetNum m.geAnnStruct.gdhTop, 1, Y
                m.oAlert.GetsUpToPrice = Y
            Else
                m.Y2 = Y
                gdSetNum m.geAnnStruct.gdhBottom, 1, Y
                m.oAlert.GetsDownToPrice = Y
            End If
        Else
            m.Y1 = Y
            m.Y2 = Y
            gdSetNum m.geAnnStruct.gdhTop, 1, Y
            gdSetNum m.geAnnStruct.gdhBottom, 1, Y
            If m.oAlert.UseGetsUpTo Then
                m.oAlert.GetsUpToPrice = Y
            Else
                m.oAlert.GetsDownToPrice = Y
            End If
        End If
        
        MoveIcon = True
        Exit Function
    ElseIf m.eUsage <> eANNOT_UserAdded Then
        Exit Function
    End If
    
    nPaneIdSave = m.geAnnStruct.paneId
    
    If m.geAnnStruct.paneId <> nPaneID Then
        rc = geRemoveItem(Chart.geChartObj, 4, m.geAnnStruct)
        If rc = 0 Then
            m.geAnnStruct.paneId = nPaneID
            rc = geAddItem(Chart.geChartObj, 4, m.geAnnStruct)
            If rc <> 0 Then
                m.geAnnStruct.paneId = nPaneIdSave
                rc = geAddItem(Chart.geChartObj, 4, m.geAnnStruct)
                MoveIcon = False
                Exit Function
            End If
        End If
    End If
    
    bMoved = True
    m.X1 = RoundNum(X)
    m.Y1 = Y
    m.geAnnStruct.paneId = nPaneID
    
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If bMoved Then
            m.dDate2 = m.dDate1             '6987
            m.strPane = Chart.Tree.Key(nPaneID)
        End If
    End If
    
    MoveIcon = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveIcon", eGDRaiseError_Raise

End Function

' The user is moving one of the annotation points
Private Function MoveEllipse(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:
       
    Dim rc&
    
    MoveEllipse = False
    If nPoint = 0 And m.geAnnStruct.moveable = 0 Then Exit Function
    
    If ActiveChart Is Nothing Then Exit Function
    
    rc = geIncDecMinorAxis(m.geAnnStruct, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
    
    If rc = 0 Then
        MoveEllipse = True
        If m.geAnnStruct.minorAxisLenData = 0 Then
            Prop("MinorAxisLen") = m.geAnnStruct.lenRatio
        Else
            Prop("MinorAxisLen") = m.geAnnStruct.lenPoints
        End If
    Else
        MoveEllipse = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveTrendline", eGDRaiseError_Raise
    
End Function

Private Function MoveTimeCycle(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#, YDelta#
            
    XDelta = m.X2 - m.X1
    YDelta = m.Y2 - m.Y1
    
    If nPaneID <> m.geAnnStruct.paneId Then Exit Function
                                
    bMoved = True
    
    Select Case nPoint
        Case 0 'mid-point (move whole line)
            m.X1 = X - m.dMoveOffSetX
            m.X2 = m.X1 + XDelta
       
        Case 1      'arc's height
            geCalcArcHeight m.geAnnStruct, Y
            Prop("ArcPercentHeight") = m.geAnnStruct.lenRatio
                    
        Case 2      'resize width
            If X > m.X1 Then
                m.X2 = RoundNum(X)
            Else
                'line on other side of base line
                m.X2 = m.X2 + (m.X1 - X - XDelta)
            End If
            
        Case 4
            m.X1 = X - XDelta
            m.X2 = X + 1
                                
        Case Else
            bMoved = False
    End Select
       
    If bMoved Then
        If nPoint = 1 Then
            'do nothing
        Else
            ' make sure is in range
            If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
            If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
        End If
    End If
    
    MoveTimeCycle = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveTimeCycle", eGDRaiseError_Raise

End Function

Private Function MoveTrendline(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#, YDelta#
            
    XDelta = m.X2 - m.X1
    If m.bUsingLogDrawingMode Then
        YDelta = LogModeDelta(m.Y1, m.Y2)           'Log(m.Y2) - Log(m.Y1)
    Else
        YDelta = m.Y2 - m.Y1
    End If
    
    If nPaneID <> m.geAnnStruct.paneId Then Exit Function
    
    If m.eType = eANNOT_TrendChannel And nPoint > 2 Then
        Exit Function
    End If
    
    bMoved = True
    Select Case nPoint
        Case 0 'mid-point (move whole line)
            m.X1 = X - m.dMoveOffSetX
            m.X2 = m.X1 + XDelta
            If m.eType <> eANNOT_TimeCycle Then
                If m.bUsingLogDrawingMode Then
                    m.Y1 = Exp(Log(Y) - m.dMoveOffSetLogY)
                    m.Y2 = Exp(Log(m.Y1) + YDelta)
                Else
                    m.Y1 = Y - m.dMoveOffSetY
                    m.Y2 = m.Y1 + YDelta
                End If
            End If
       
        Case 1, 5 'start pt
            If m.eType = eANNOT_TimeCycle Then
                geCalcArcHeight m.geAnnStruct, Y
                Prop("ArcPercentHeight") = m.geAnnStruct.lenRatio
            Else
                m.X1 = RoundNum(X)
                m.Y1 = Y
            End If
                    
        Case 2, 3 'end pt
            m.X2 = RoundNum(X)
            m.Y2 = Y
                    
        Case Else
            bMoved = False
    End Select
       
    If bMoved Then
        If nPoint = 1 And m.eType = eANNOT_TimeCycle Then
            'do nothing
        Else
            ' make sure is in range
            If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
            If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
        End If
    End If
    
    MoveTrendline = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveTrendline", eGDRaiseError_Raise

End Function

' The user is moving one of the annotation points
Private Function MoveDollarLine(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#, YDelta#, XSave#, dMinMove#, dLastBarx#
    Dim Bars As cGdBars
    Dim Pane As cPane
    Dim Ind As cIndicator
    
    If nPaneID <> m.geAnnStruct.paneId Or m.eUsage <> eANNOT_UserAdded Then
        Exit Function
    End If
    
    XDelta = m.X2 - m.X1
    If m.bUsingLogDrawingMode Then
        YDelta = LogModeDelta(m.Y1, m.Y2)       'Log(m.Y2) - Log(m.Y1)
    Else
        YDelta = m.Y2 - m.Y1
    End If
    
    dLastBarx = GetX(Chart, Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False, True))) + m.Chart.BlankBars - 1
    
    bMoved = True
    Select Case nPoint
        Case 0 'mid-point (move whole line)
            If Val(Prop("KeepAtEnd")) = 0 Then
                XSave = m.X1
                m.X1 = X - m.dMoveOffSetX
                If m.X1 + XDelta > dLastBarx Then
                    m.X1 = XSave
                    bMoved = False
                Else
                    m.X2 = m.X1 + XDelta
                    If m.bUsingLogDrawingMode Then
                        m.Y1 = Exp(Log(Y) - m.dMoveOffSetLogY)
                        m.Y2 = Exp(Log(m.Y1) + YDelta)
                    Else
                        m.Y1 = Y - m.dMoveOffSetY
                        m.Y2 = m.Y1 + YDelta
                    End If
                End If
            Else
                bMoved = False
            End If
        
        Case 1 'start pt
            If m.X1 > m.X2 And Prop("KeepAtEnd") = 1 Then
                bMoved = False
            Else
                m.X1 = RoundNum(X)
                m.Y1 = RoundedY(Chart, Y)
            End If
                    
        Case 2 'end pt
            If m.X2 > m.X1 And Prop("KeepAtEnd") = 1 Then
                bMoved = False
            Else
                m.X2 = RoundNum(X)
                m.Y2 = RoundedY(Chart, Y)
            End If
        
        Case Else
            bMoved = False
    End Select
                
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
        ' round Y values to the nearest min move for price
        Set Pane = Chart.Tree(m.strPane)
        If Not Pane Is Nothing Then
            Set Ind = Chart.Tree(Pane.idxPriceBars)
            If Not Ind Is Nothing Then
                Set Bars = Ind.Bars
            End If
        End If
        If Bars Is Nothing Then
            Set Bars = Chart.Bars
        End If
        If Not Bars Is Nothing Then
            dMinMove = Bars.MinMove(m.dDate1)
            If dMinMove > 0 Then
                m.Y1 = Int(m.Y1 / dMinMove + 0.5) * dMinMove
                m.Y2 = Int(m.Y2 / dMinMove + 0.5) * dMinMove
            End If
        End If
    End If
    
    Set Pane = Nothing
    Set Ind = Nothing
    MoveDollarLine = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveDollarLine", eGDRaiseError_Raise

End Function

Private Function MoveFibonacci(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#, YDelta#
    
    If m.eUsage <> eANNOT_UserAdded Then Exit Function
    If nPaneID <> m.geAnnStruct.paneId Then Exit Function
               
    XDelta = m.X2 - m.X1
    If m.bUsingLogDrawingMode Then
        YDelta = LogModeDelta(m.Y1, m.Y2)           'Log(m.Y2) - Log(m.Y1)
    Else
        YDelta = m.Y2 - m.Y1
    End If
    
    bMoved = True
    Select Case nPoint
        Case 0 'mid-point (move whole line)
            m.X1 = X - m.dMoveOffSetX
            m.X2 = m.X1 + XDelta
            If m.bUsingLogDrawingMode Then
                m.Y1 = Exp(Log(Y) - m.dMoveOffSetLogY)
                m.Y2 = Exp(Log(m.Y1) + YDelta)
            Else
                m.Y1 = Y - m.dMoveOffSetY
                m.Y2 = m.Y1 + YDelta
            End If
        
        Case 1 'top
            If m.eType = eANNOT_FibArcs Then
                m.Y1 = Y - m.dMoveOffSetY       'm.Y1 is center, not edge
            Else
                m.Y1 = Y
            End If
            
        Case 2 'bottom
            m.Y2 = Y
        
        Case 3 'corner
            If m.eType = eANNOT_FibArcs Then
                m.X1 = X - Abs(m.X2 - m.X1)
                m.Y1 = Y - m.dMoveOffSetY
            Else
                m.X1 = RoundNum(X)
                m.Y1 = Y
            End If
            'm.X1 = RoundNum(X)
            'm.Y1 = Y
                    
        Case 4 'corner
            m.X2 = RoundNum(X)
            'm.Y1 = Y
            If m.eType = eANNOT_FibArcs Then
                m.Y1 = Y - m.dMoveOffSetY
            Else
                m.Y1 = Y
            End If
        
        Case 5 'corner
            If m.eType = eANNOT_FibArcs Then
                m.X1 = X - Abs(m.X2 - m.X1)
            Else
                m.X1 = RoundNum(X)
            End If
            'm.X1 = RoundNum(X)
            m.Y2 = Y
        
        Case 6 'corner
            m.X2 = RoundNum(X)
            m.Y2 = Y
        
        Case 7, 9, 11 'one side
            If m.eType = eANNOT_FibArcs Then
                m.X1 = X - Abs(m.X2 - m.X1)     'm.X1 is center, not edge
            Else
                m.X1 = RoundNum(X)
            End If
            
        Case 8, 10, 12 'other side
            m.X2 = RoundNum(X)
        
        Case Else
            bMoved = False
    End Select
    
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
        
        If m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 _
           Or m.eType = eANNOT_Fibonacci3 Or m.eType = eANNOT_Fibonacci4 _
           Or m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_DanCodeFib _
           Or m.eType = eANNOT_AdvRiskReward Then
            If Val(Prop("HideVerticalLine")) = 1 Then
                'tool is set to show extensions only AND hide diagonal
                'turn on diagonal temporarily while moving
                If Val(Prop("Ext")) = 4 Then Prop("HideVerticalLine") = -1
            End If
        End If
    End If
    
    MoveFibonacci = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveFibonacci", eGDRaiseError_Raise

End Function

' The user is moving one of the annotation points
Private Function MoveTargetShooter(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#, YDelta#, X1#, X2#, dMult#
    
    If nPaneID <> m.geAnnStruct.paneId Or m.eUsage <> eANNOT_UserAdded Then
        Exit Function
    End If
    
    XDelta = m.X2 - m.X1
    YDelta = m.Y2 - m.Y1
       
    bMoved = True
    Select Case nPoint
        Case 0 'mid-point (move whole line)
            MoveTargetShooter = False
        
        Case 1, 3, 7 'start pt
            m.X1 = RoundNum(X)
            m.Y1 = Y
                    
        Case 2, 5, 8 'end pt
            m.X2 = RoundNum(X)
            m.Y2 = Y
        
        Case 10, 12, 14, 16, 18, 20 ' extend mult
            X1 = m.X1
            X2 = m.X2
            If m.X1 > m.X2 Then
                X1 = m.X2
                X2 = m.X1
            End If
            If X <= X2 Then
                bMoved = False
            Else
                dMult = (X - X2) / (X2 - X1 + 1)
                Prop("ExtMult") = dMult
            End If
        
        Case Else
            bMoved = False
    End Select
    
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
    End If
    
    MoveTargetShooter = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveTargetShooter", eGDRaiseError_Raise

End Function

Private Function MoveAndrewFork(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim XDelta#, YDelta#, dX3#, dX3Delta#, dY3Delta#
    Dim dX1Old#, dX2Old#, dY1Old#, dY2Old#, dY3Old#
    Dim dDate1Old#, dDate2Old#, dDate3Old#
    
    Dim hArrayDate&, dNewDate#
    
    Dim bMoved As Boolean
    
    If nPaneID <> m.geAnnStruct.paneId Then Exit Function
    If nPoint = 0 And m.geAnnStruct.moveable = 0 Then Exit Function
    
    XDelta = m.X2 - m.X1
    
    dX3 = GetX(Chart, m.aDates(0))
    dX3Delta = dX3 - m.X1
    
    If m.bUsingLogDrawingMode Then
        YDelta = LogModeDelta(m.Y1, m.Y2)           'Log(m.Y2) - Log(m.Y1)
        If m.aYs.Size > 0 Then
            dY3Delta = LogModeDelta(m.Y1, m.aYs(0)) 'Log(m.aYs(0)) - Log(m.Y1)
        End If
    Else
        YDelta = m.Y2 - m.Y1
        dY3Delta = m.aYs(0) - m.Y1
    End If
    
    hArrayDate = Chart.geDateArray      '4240
    dNewDate = gdGetNum(hArrayDate, gdGetSize(hArrayDate) - 1)
    
    If X <= GetX(Chart, dNewDate) And X >= GetX(Chart, gdGetNum(hArrayDate, 0)) Then
        dNewDate = Chart.aXdate(X)
    Else
       Exit Function            'outside visible range of chart's date range
    End If
    
    bMoved = True
    
    Select Case nPoint
        Case 0:             'move whole line
            'save current values
            dX1Old = m.X1
            dX2Old = m.X2
            dY1Old = m.Y1
            dY2Old = m.Y2
            dY3Old = m.aYs(0)
            dDate1Old = m.dDate1
            dDate2Old = m.dDate2
            dDate3Old = m.aDates(0)
            'set moved values
            m.X1 = X - m.dMoveOffSetX
            m.X2 = m.X1 + XDelta
            dX3 = m.X1 + dX3Delta

            If m.bUsingLogDrawingMode Then
                m.Y1 = Exp(Log(Y) - m.dMoveOffSetLogY)
                m.Y2 = Exp(Log(m.Y1) + YDelta)
                If m.aYs.Size > 0 Then
                    m.aYs(0) = Exp(Log(m.Y1) + dY3Delta)
                End If
            Else
                m.Y1 = Y - m.dMoveOffSetY
                m.Y2 = m.Y1 + YDelta
                m.aYs(0) = m.Y1 + dY3Delta
            End If
            
            
            If m.dDate1 <> 0 Then m.dDate1 = Chart.aXdate(m.X1)
            If m.dDate2 <> 0 Then m.dDate2 = Chart.aXdate(m.X2)
            If m.aDates(0) <> 0 Then m.aDates(0) = Chart.aXdate(dX3)
            
            'check for going past end of data
            If m.dDate1 < 0 Or m.dDate2 < 0 Or m.aDates(0) < 0 Then bMoved = False      '2221
            If m.X1 < dX1Old Then
                If dX1Old - m.X1 >= Int(gdGetSize(hArrayDate) / 2) Then bMoved = False  '5977
            End If
            
            If Not bMoved Then
                m.X1 = dX1Old
                m.X2 = dX2Old
                m.Y1 = dY1Old
                m.Y2 = dY2Old
                m.aYs(0) = dY3Old
                m.dDate1 = dDate1Old
                m.dDate2 = dDate2Old
                m.aDates(0) = dDate3Old
                bMoved = False
            End If
        
        Case 1:
            m.X1 = X
            If m.eType = eANNOT_ElliotTimeRatio And m.aDates.Size > 0 Then
                If Abs(m.Y1 - Y) < Abs(m.Y2 - Y) Then
                    m.Y1 = Y
                ElseIf Abs(m.Y1 - Y) > Abs(m.Y2 - Y) Then
                    m.Y2 = Y
                End If
            Else
                m.Y1 = Y
            End If
            If m.dDate1 <> 0 Then m.dDate1 = dNewDate
        Case 2:
            m.X2 = X
            If m.eType = eANNOT_ElliotTimeRatio And m.aDates.Size > 0 Then
                If Abs(m.Y1 - Y) < Abs(m.Y2 - Y) Then
                    m.Y1 = Y
                ElseIf Abs(m.Y1 - Y) > Abs(m.Y2 - Y) Then
                    m.Y2 = Y
                End If
            Else
                m.Y2 = Y
            End If
            If m.dDate2 <> 0 Then m.dDate2 = dNewDate
        Case 3:
            If m.eType = eANNOT_ElliotTimeRatio And m.aDates.Size > 0 Then
                If Abs(m.Y1 - Y) < Abs(m.Y2 - Y) Then
                    m.Y1 = Y
                ElseIf Abs(m.Y1 - Y) > Abs(m.Y2 - Y) Then
                    m.Y2 = Y
                End If
            Else
                m.aYs(0) = Y
            End If
            If m.aDates(0) <> 0 Then m.aDates(0) = dNewDate
        Case Else
            bMoved = False
    End Select
        
    MoveAndrewFork = bMoved
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveAndrewFork", eGDRaiseError_Raise

End Function

Private Function MoveDnExpansion(Chart As cChart, ByVal nPoint&, _
        ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Points 1,2,3 --> initial draw in progress
'Points 11,12,13,14,15,16 --> user extending lines of existing annotation
'Point 17 --> B & C are on same bar, user extending line on B to left of chart
'nPoint 1 is always first mouse click
'nPoint 2 is always 2nd mouse click
'nPoint 3 is always 3rd mouse click
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim iStart&, iEnd&
    Dim dDateSave#, dYSave#
    Dim bSaveRatio As Boolean
    Dim bSaveMain As Boolean
    Dim bOkay As Boolean
    
    bOkay = True
    
    StatusMsg Str(nPoint)
    
    If nPoint = 2 Then
        dDateSave = m.dDate2
        dYSave = m.Y2
        m.dDate2 = X
        m.Y2 = Y
    ElseIf nPoint = 1 Then
        dDateSave = m.dDate1
        dYSave = m.Y1
        m.dDate1 = X
        m.Y1 = Y
    ElseIf nPoint = 3 Then
        dDateSave = gdGetNum(m.aDates.ArrayHandle, 0)
        dYSave = gdGetNum(m.aYs.ArrayHandle, 0)
        gdSetNum m.aDates.ArrayHandle, 0, X
        gdSetNum m.aYs.ArrayHandle, 0, Y
    ElseIf nPoint = 11 Or nPoint = 13 Then
        If gdGetNum(m.geAnnStruct.gdhTop, m.nHitItemIdx) = m.Y1 Then
            iStart = GetX(Chart, m.dDate1)
        Else
            iStart = GetX(Chart, m.aDates(0))
        End If
        iEnd = GetX(Chart, X)
        If iEnd > iStart Then bSaveMain = True
    ElseIf nPoint = 12 Then
        iStart = GetX(Chart, m.dDate2)  'B with text to right
        iEnd = GetX(Chart, X)
        If iEnd > iStart Then bSaveMain = True
    ElseIf nPoint = 999999 Then
        iStart = GetX(Chart, X)     'B with text to left
        iEnd = GetX(Chart, m.dDate2)
        If iEnd > iStart Then bSaveMain = True
    ElseIf nPoint > 13 Then         'And nPoint <= m.aRatios.Size + 11 Then (do not remember why this is here, commented out to fix 6315)
        If m.dDate1 > m.aDates(0) Then
            iStart = GetX(Chart, m.dDate1)
        Else
            iStart = GetX(Chart, m.aDates(0))
        End If
        iEnd = GetX(Chart, X)
        If iEnd > iStart Then bSaveRatio = True
    Else
        bOkay = False
    End If
    
    If geMoveFlag = 1 Or nPoint = 3 Then
        bOkay = ValidDNE
        If Not bOkay Then
            Select Case nPoint
                Case 1:
                    m.dDate1 = dDateSave
                    m.Y1 = dYSave
                Case 2:
                    m.dDate2 = dDateSave
                    m.Y2 = dYSave
                Case 3:
                    gdSetNum m.aDates.ArrayHandle, 0, dDateSave
                    gdSetNum m.aYs.ArrayHandle, 0, dYSave
            End Select
        End If
    End If
    
    If bSaveMain Then
        If (iEnd - iStart) < Chart.geChartPoints Then
            Prop("LenA") = iEnd - iStart
            Prop("LenB") = iEnd - iStart
            Prop("LenC") = iEnd - iStart
        End If
    ElseIf bSaveRatio Then
        Prop("LenCOP") = iEnd - iStart
        Prop("LenOP") = iEnd - iStart
        Prop("LenXOP") = iEnd - iStart
    End If
    
    MoveDnExpansion = bOkay

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveDnExpansion", eGDRaiseError_Raise

End Function
        
Private Sub ShowDnRetracement(ByVal bShow As Boolean)
On Error GoTo ErrSection:

    Dim i&, Size&
    
    Size = gdGetSize(m.geAnnStruct.glhAnnType)
    
    If bShow = False Then
        For i = 0 To Size - 1
            gdSetNum m.geAnnStruct.glhAnnType, i, -1
        Next
        Exit Sub
    End If
    
    gdSetNum m.geAnnStruct.glhAnnType, 0, 15
    gdSetNum m.geAnnStruct.glhAnnType, 1, 6     'arc properties type=ellipse
    gdSetNum m.geAnnStruct.glhAnnType, 2, 1     'text & line properties type=line
    gdSetNum m.geAnnStruct.glhAnnType, 3, 6     'arc properties type=ellipse
    gdSetNum m.geAnnStruct.glhAnnType, 4, 1     'text & line properties type=line
    gdSetNum m.geAnnStruct.glhAnnType, 5, 0     'font information
    
    For i = 6 To Size - 1
        gdSetNum m.geAnnStruct.glhAnnType, i, 3
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ShowDnRetracement", eGDRaiseError_Raise

End Sub

Private Function MoveDnRetracement(Chart As cChart, ByVal nPoint&, _
        ByVal X#, Y#, nFlag&) As Boolean
On Error GoTo ErrSection:

    'Note on input parameters:
    'nFlag: This is an [in][out] parameter.
    '   Has no meaning on input. Has following meanings on output:
    '   0: invalid point (not within prescribed boundary)
    '   1: point is valid
    '   2: point is the same as last point clicked or is to the right of
    '      last point clicked (can be used as trigger to end series)
    'X: This is a date value.

    Dim i&, Size&, nFocusLow&
    Dim hDate&, hY&
    'points to left & right of potential new point (requires dates going forward to right)
    Dim dyLeft#, dyRight#, dDateLeft#, dDateRight#, idxLeft#, idxRight#
    
    Dim bValid As Boolean
        
    nFocusLow = Val(Prop("FocusLow"))
   
   'cannot go forward in date past focus point
    If X > m.dDate1 Then
        nFlag = 2
        Exit Function
    End If
    
    'if focus is low, all reaction points must be high
    'if focus is high, all reaction points must be low
    If nFocusLow = 1 Then
        If Y < m.Y1 Then Exit Function
    Else
        If Y > m.Y1 Then Exit Function
    End If
    
    idxLeft = -1
    idxRight = -1
    Size = m.aDates.Size
    hDate = m.aDates.ArrayHandle
    hY& = m.aYs.ArrayHandle
        
    bValid = True
    gdSetNum hDate, nPoint, X   'assume valid and add to array
    gdSetNum hY, nPoint, Y
    
    If nPoint = 0 Then
        dDateRight = m.dDate1  'first reaction point, must be left of focus
        dyRight = m.Y1
        dDateLeft = X
        dyLeft = Y
    Else
        dDateRight = gdGetNum(hDate, nPoint - 1)  'assume using going left
        dyRight = gdGetNum(hY, nPoint - 1)
    End If

    If X < dDateRight Or (X = dDateRight And Size = 0) Then 'user is going left
        If nFocusLow = 1 Then
            If Y < dyRight Then bValid = False
        Else
            If Y > dyRight Then bValid = False
        End If
    Else
        bValid = False
        nFlag = 2
    End If
                
    If bValid = True Then
        'check dnrboundary
        If X > m.dReactPtBoundary Then
            nFlag = 1
        Else
            nFlag = 0
        End If
    Else
        gdDeleteItems hDate, nPoint, 1
        gdDeleteItems hY, nPoint, 1
    End If
    
    MoveDnRetracement = bValid
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveDnRetracement", eGDRaiseError_Raise

End Function

' The user is moving one of the annotation points
Private Function MoveRectangle(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#, YDelta#
    
    If nPaneID <> m.geAnnStruct.paneId Then Exit Function
    
    If m.eType = eANNOT_Rectangle Then
        If m.eUsage <> eANNOT_UserAdded And m.eUsage <> eANNOT_PatternProfit And m.eUsage <> eANNOT_FibClusters Then Exit Function
    ElseIf m.eUsage <> eANNOT_UserAdded Then
        Exit Function
    End If
    
    If m.eType = eANNOT_TextEdit Or m.eType = eANNOT_TextEdit2 Or m.eType = eANNOT_TextEdit3 Or m.eType = eANNOT_TextEdit4 Then
        If nPoint <> 0 Then Exit Function
    End If
    
    XDelta = m.X2 - m.X1
    If m.bUsingLogDrawingMode Then
        YDelta = LogModeDelta(m.Y1, m.Y2)   'Log(m.Y2) - Log(m.Y1)
    Else
        YDelta = m.Y2 - m.Y1
    End If
    
    bMoved = True
    Select Case nPoint
        Case 0 'mid-point (move whole line)
            m.X1 = X - m.dMoveOffSetX
            m.X2 = m.X1 + XDelta
            
            If m.bUsingLogDrawingMode Then
                m.Y1 = Exp(Log(Y) - m.dMoveOffSetLogY)
                m.Y2 = Exp(Log(m.Y1) + YDelta)
            Else
                m.Y1 = Y - m.dMoveOffSetY
                m.Y2 = m.Y1 + YDelta
            End If
                
        Case 2, 3     'corner
            If Abs(m.X1 - X) < Abs(m.X2 - X) Then
                m.X1 = RoundNum(X)
            Else
                m.X2 = RoundNum(X)
            End If
            
            If Abs(m.Y1 - Y) < Abs(m.Y2 - Y) Then
                m.Y1 = Y
            Else
                m.Y2 = Y
            End If
        
        Case 5
            If m.eUsage = eANNOT_PatternProfit Then
                bMoved = False      'theoretically should not get here
            Else
                If Abs(m.Y1 - Y) < Abs(m.Y2 - Y) Then
                    m.Y1 = Y
                Else
                    m.Y2 = Y
                End If
            End If
            
        Case 7
            If Abs(m.X1 - X) < Abs(m.X2 - X) Then
                m.X1 = X
            Else
                m.X2 = X
            End If
            
        Case Else
            bMoved = False
    End Select
    
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
    End If
        
    MoveRectangle = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveRectangle", eGDRaiseError_Raise

End Function

Private Function MoveRiskReward(Chart As cChart, ByVal nPoint&, _
    ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
    
    Dim bMoved As Boolean
    Dim XDelta#, YDelta#
    Dim dUserEnd1#, dUserEnd2#, dMinMove#
    Dim dDiff#, dDiffY2#, dDiffPoint1#, dDiffPoint2#

    bMoved = True           'assume success
    
    XDelta = m.X2 - m.X1
    If m.bUsingLogDrawingMode Then
        YDelta = LogModeDelta(m.Y1, m.Y2)       'Log(m.Y2) - Log(m.Y1)
    Else
        YDelta = m.Y2 - m.Y1
    End If
    
    If nPoint = 0 Then
        dUserEnd1 = Val(Prop("UserEndPoint1"))
        dUserEnd2 = Val(Prop("UserEndPoint2"))
        
        If m.bUsingLogDrawingMode Then
            dDiffPoint1 = LogModeDelta(m.Y1, dUserEnd1)     'Log(dUserEnd1) - Log(m.Y1)
            dDiffPoint2 = LogModeDelta(m.Y2, dUserEnd2)     'Log(dUserEnd2) - Log(m.Y2)
        Else
            dDiffPoint1 = dUserEnd1 - m.Y1
            dDiffPoint2 = dUserEnd2 - m.Y2
        End If
        
        'mid-point (move whole line)
        m.X1 = X - m.dMoveOffSetX
        m.X2 = m.X1 + XDelta
        
        If m.bUsingLogDrawingMode Then
            m.Y1 = Exp(Log(Y) - m.dMoveOffSetLogY)
            m.Y2 = Exp(Log(m.Y1) + YDelta)
            dUserEnd1 = Exp(Log(m.Y1) + dDiffPoint1)
            dUserEnd2 = Exp(Log(m.Y2) + dDiffPoint2)
        Else
            m.Y1 = Y - m.dMoveOffSetY
            m.Y2 = m.Y1 + YDelta
            dUserEnd1 = m.Y1 + dDiffPoint1
            dUserEnd2 = m.Y2 + dDiffPoint2
        End If
        
        'round to nearest min move
        dMinMove = Chart.Bars.MinMove(m.dDate1)
        Prop("UserEndPoint1") = Int(dUserEnd1 / dMinMove + 0.5) * dMinMove
        Prop("UserEndPoint2") = Int(dUserEnd2 / dMinMove + 0.5) * dMinMove
    Else
        'find closest point to Y
        dDiff = Abs(m.Y1 - Y)
        dDiffY2 = Abs(m.Y2 - Y)
        dDiffPoint1 = Abs(Val(Prop("UserEndPoint1")) - Y)
        dDiffPoint2 = Abs(Val(Prop("UserEndPoint2")) - Y)
        
        If dDiff <= dDiffY2 And dDiff < dDiffPoint1 And dDiff < dDiffPoint2 Then
            m.Y1 = Y
        ElseIf dDiffY2 < dDiff And dDiffY2 < dDiffPoint1 And dDiffY2 < dDiffPoint2 Then
            m.Y2 = Y
        ElseIf dDiffPoint1 < dDiff And dDiffPoint1 < dDiffY2 And dDiffPoint1 < dDiffPoint2 Then
            Prop("UserEndPoint1") = Y
        ElseIf dDiffPoint2 < dDiff And dDiffPoint2 < dDiffY2 And dDiffPoint2 < dDiffPoint1 Then
            Prop("UserEndPoint2") = Y
        Else
            bMoved = False
        End If
                        
    End If
            
    If bMoved Then
        ' reset new points
        If m.dDate1 <> 0 Then m.dDate1 = Chart.aXdate(m.X1)
        If m.dDate2 <> 0 Then m.dDate2 = Chart.aXdate(m.X2)
    End If
    
    MoveRiskReward = bMoved

End Function

Private Function MoveRegressionLine(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#, YDelta#, XSave
    Dim dLastBarx#, dFirstBarX#
    Dim Bars As cGdBars
    Dim Pane As cPane
    
    If nPaneID <> m.geAnnStruct.paneId Or m.eUsage <> eANNOT_UserAdded Then
        Exit Function
    End If
    
    'don't let user move annotation past last good data bar
    'or stretch line outside chart's left side
    dFirstBarX = Chart.ScreenStartX
    dLastBarx = Chart.LastGoodDataBar(True, False)
    If X > dLastBarx Then
        MoveRegressionLine = False
        Exit Function
    End If
    
    bMoved = True
    XDelta = m.X2 - m.X1
    YDelta = m.Y2 - m.Y1
    Select Case nPoint
        Case 0 'mid-point (move whole line)
            If Prop("KeepAtEnd") = 0 Then
                XSave = m.X1
                m.X1 = X - m.dMoveOffSetX
                'If m.X1 + Xdelta > dLastBarx Or m.X1 < dFirstBarX Then
                If m.X1 + XDelta > dLastBarx Then
                    m.X1 = XSave
                    bMoved = False
                Else
                    m.X2 = m.X1 + XDelta
                    m.Y1 = Y - m.dMoveOffSetY
                    m.Y2 = m.Y1 + YDelta
                End If
            Else
                bMoved = False
            End If
        
        Case 1 'start pt
            If RoundNum(X) > dFirstBarX Then
                m.X1 = RoundNum(X)
                m.Y1 = RoundedY(Chart, Y)
            Else
                bMoved = False
            End If
                    
        Case 2 'end pt
            If RoundNum(X) > dFirstBarX Then
                m.X2 = RoundNum(X)
                m.Y2 = RoundedY(Chart, Y)
            Else
                bMoved = False
            End If
        
        Case Else
            bMoved = False
    End Select
                
    If bMoved Then
        'make sure is in range
        If m.X1 < 0 Or m.X2 < 0 Or m.X1 >= Chart.aXdate.Size _
            Or m.X2 >= Chart.aXdate.Size Then
            bMoved = False
        Else
            'save new length
            Prop("LenInBars") = Abs(m.X2 - m.X1)
            'set moveflag to 0 to have grapheng.dll draw vertical guides
            m.geAnnStruct.moveable = 0
        End If
    End If
    
    MoveRegressionLine = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveRegressionLine", eGDRaiseError_Raise

End Function

' Caption to display while moving
Public Function MoveCaption(Chart As cChart) As String
On Error GoTo ErrSection:

    Dim strCaption$
    
    '07-14-2010 roll back to code without If InStr(g.strActiveDraw, "PFP") = 0
    
    If InStr(g.strActiveDraw, "PFP") = 0 Then
        strCaption = " Length = " & CStr(Abs(m.X2 - m.X1)) & " bars"
        If m.eType <> eANNOT_TimeCycle And m.eType <> eANNOT_RegressionLine _
            And m.eType <> eANNOT_Pattern Then
            strCaption = strCaption & ",  Height = " & _
                RoundedY(Chart, m.Y2 - m.Y1)
        End If
    Else
        strCaption = " Length = " & Str(PatternLength(Chart)) & " bars"
    End If
    
    MoveCaption = strCaption

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveCaption", eGDRaiseError_Raise

End Function

Friend Sub SetPrivateData(mCopyFrom As mPrivate, oAlert As cAlert)
On Error GoTo ErrSection:

    Dim SaveStruct As chart_annotation

    ' get copy of all private data
    ' (except must keep original geAnnStruct)
    SaveStruct = m.geAnnStruct
    m = mCopyFrom
    m.geAnnStruct = SaveStruct
    
    ' make a "copy" of the original properties
    Set m.Properties = mCopyFrom.Properties.MakeCopy
    Set m.aDates = mCopyFrom.aDates.MakeCopy
    Set m.aYs = mCopyFrom.aYs.MakeCopy
    
    If m.eUsage = eANNOT_PriceAlert And Not oAlert Is Nothing Then
        'copy alert object so can restore on annot undo             -6458
        Set m.oAlert = New cAlert
        m.oAlert.FromFileString oAlert.ToFileString
        m.oAlert.AlertKey = oAlert.AlertKey
    Else
        'reset so this will get its own alert object to be added to the alerts collection
        Set m.oAlert = Nothing
    End If
    
    m.bAlertAdded = False
    
    ' reset engine flags
    geResetData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetPrivateData", eGDRaiseError_Raise

End Sub

Public Function MakeCopy() As cAnnotation
On Error GoTo ErrSection:

    Dim aCopy As New cAnnotation
    
    aCopy.SetPrivateData m, m.oAlert            '6458
    
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MakeCopy", eGDRaiseError_Raise

End Function

' returns point which will be "active" when continuing to draw
Public Function CreateNew(Chart As cChart, ByVal eType As eAnnotType, _
        ByVal nPaneID&, ByVal dDate1#, ByVal dY1#, _
        ByVal dDate2#, ByVal dY2#, _
        Optional ByVal nColor& = -1, _
        Optional ByVal nStyle& = -1, _
        Optional ByVal strText As Variant, _
        Optional ByVal eUsage As eAnnotUsage = eANNOT_UserAdded, _
        Optional ByVal bDefaultsOnly As Boolean = False, _
        Optional ByVal bSchiff As Boolean = False) As Long 'array of dates & y-values for DnRetracement
On Error GoTo ErrSection:

    Dim nActivePoint&, strPane$
    
    nActivePoint = -1
    strPane = Chart.Tree.Key(nPaneID)
    m.bSchiffFork = bSchiff

    If Len(strPane) > 0 Then
        m.eType = eType
        m.eUsage = eUsage
    
        If m.eType = eANNOT_ElliotEwave Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, 35
            m.geAnnStruct.annotationId = Chart.Annots.Count + 1
            m.geAnnStruct.paneId = nPaneID
            m.geAnnStruct.lenPoints = Chart.Bars.MinMove
            
            For nActivePoint = 0 To gdGetSize(dDate1) - 1
                gdSetStr m.geAnnStruct.gshText, nActivePoint, gdGetStr(dDate1, nActivePoint)
            Next
            
            If geAddItem(Chart.geChartObj(), 4, m.geAnnStruct) = 0 Then
                m.geAdded = True
                m.geChartObj = Chart.geChartObj()
            End If
            Exit Function
        ElseIf m.eUsage = eANNOT_UserAdded Then
            LoadDefaults
            If bDefaultsOnly Then Exit Function
        End If
       
        m.strPane = strPane
        m.dDate1 = dDate1
        m.Y1 = dY1
        
        If eType = eANNOT_DNExpansion Or eType = eANNOT_DNExpansion2 Or eType = eANNOT_DNExpansion3 Or _
           eType = eANNOT_DNExpansion4 Or eType = eANNOT_FibABCD Or eType = eANNOT_Gartley Or _
           eType = eANNOT_DNRetracement Then
           
            If InitDinapoli(Chart, eType) = False Then
                nActivePoint = -1
                Exit Function
            End If
        
        ElseIf m.eType = eANNOT_Pivot And m.dDate1 > Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False)) Then
            MsgBox "Cannot go past last data bar for Pivot Points draw tool.", vbInformation
            CreateNew = -1
            Exit Function
        Else
            m.dDate2 = dDate2
            m.Y2 = dY2
            If nColor >= 0 Then m.nColor = nColor
            If nStyle >= 0 Then m.nStyle = nStyle
            If Not IsMissing(strText) Then m.strText = strText
        End If

        Select Case eType
            Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4, _
                eANNOT_VertLine, eANNOT_Icon, eANNOT_ElliotLabel, eANNOT_DNRetracement, _
                eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, eANNOT_DNExpansion4, _
                eANNOT_AndrewFork, eANNOT_BellAlert, eANNOT_ElliotTimeRatio, eANNOT_FibABCD, _
                eANNOT_Gartley, eANNOT_GannacciCycle, eANNOT_GannacciTime, eANNOT_GannacciSwing1, _
                eANNOT_GannacciSwing2
                nActivePoint = 0
            Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
                    eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4, _
                    eANNOT_RegressionLine, eANNOT_TargetShooter, eANNOT_Rectangle, eANNOT_TextEdit, _
                    eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4, _
                    eANNOT_Ellipse, eANNOT_GannLines, eANNOT_TimeCycle, _
                    eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4, _
                    eANNOT_MultiRects, eANNOT_FibFan, eANNOT_SpResistFan, _
                    eANNOT_Mirror, eANNOT_RiskReward, eANNOT_Pattern, eANNOT_ChannelHighlight, _
                    eANNOT_TriangleWedge, eANNOT_WaveLabels, eANNOT_Bracket, eANNOT_Pivot, _
                    eANNOT_ArrowLine, eANNOT_TrendChannel, eANNOT_GannacciSwingSquare, eANNOT_FibArcs
                nActivePoint = 2
                If eUsage = eANNOT_PatternProfit Then
                    m.geAnnStruct.showAxes = 1
                    m.geAnnStruct.skipHitTest = 1
                ElseIf eType = eANNOT_GannLines Then
                    m.geAnnStruct.moveable = 1
                End If
            Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, eANNOT_Fibonacci4, _
                eANNOT_FibTimeRatio, eANNOT_FibTimeZones, eANNOT_FibExpansion, eANNOT_AdvRiskReward, _
                eANNOT_DanCodeFib, eANNOT_DanCodeZone
                nActivePoint = 6
            Case eANNOT_BalloonStrangle
                nActivePoint = 0
                m.Y2 = 0#
                m.aYs(0) = 0#
                m.aYs(1) = 0#
                m.aYs(2) = 0#
                m.aYs(3) = 0#
        End Select
    End If
    
    geAddAnnotation Chart, nPaneID, Chart.Annots.Count + 1
    m.dCreateDate = Now
    m.strBarPeriod = Chart.Bars.Prop(eBARS_PeriodicityStr)
    m.strPeriodsPerBar = Chart.Bars.Prop(eBARS_PeriodsPerBar)
    CreateNew = nActivePoint

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.CreateNew", eGDRaiseError_Raise

End Function

' to show values on annotations
Private Function RoundedY(Chart As cChart, ByVal dY#) As Double
On Error GoTo ErrSection:

    Dim dRange#
    Dim Pane As cPane
    'Original code (save for reference)
    'With Chart.Peg
        '.WorkingAxis = m.nPaneID
        ' divide range by 10 for one more decimal of accuracy
        'dRange = Abs(.ManualMaxY - .ManualMinY) / 10#
    'End With

    Set Pane = Chart.Tree(m.geAnnStruct.paneId)
    If Not Pane Is Nothing Then
        dRange = Abs(Pane.gePaneMax - Pane.gePaneMin) / 10#
        Set Pane = Nothing
    End If
    RoundedY = ValOfText(RoundedValueStr(dY, dRange))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.RoundedY", eGDRaiseError_Raise

End Function

Public Sub FixRegressionEndPoint(Chart As cChart)
On Error GoTo ErrSection:
    
    '.nBar = m.Chart.aXBar(.nX) 'could be < 0
    '.dDate = m.Chart.aXdate(.nX)

    Dim nSize&, nX&, i&
    
    nSize = Chart.aXdate.Size
    
    If Chart.aXBar(m.X1) < 0 Then
        nX = m.X1
    ElseIf Chart.aXBar(m.X2) < 0 Then
        nX = m.X2
    Else
        Exit Sub
    End If
    
    For i = nX + 1 To nSize - 1
        If Chart.aXBar(i) >= 0 Then
            If nX = m.X1 Then
                m.dDate1 = Chart.aXdate(i)
                m.X1 = i
            Else
                m.dDate2 = Chart.aXdate(i)
                m.X2 = i
            End If
            Exit For
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.FixRegressionEndPoint", eGDRaiseError_Raise

End Sub

Public Property Get geMoveFlag() As Long
On Error GoTo ErrSection:

    geMoveFlag = m.geAnnStruct.moveable

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geMoveFlag.Get", eGDRaiseError_Raise

End Property

Public Property Let geMoveFlag(ByVal nMove&)
On Error GoTo ErrSection:

    m.geAnnStruct.moveable = nMove

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geMoveFlag.Let", eGDRaiseError_Raise

End Property

Public Property Get gePaneId() As Long
On Error GoTo ErrSection:

    gePaneId = m.geAnnStruct.paneId

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.gePaneId.Get", eGDRaiseError_Raise

End Property

Public Property Let gePaneId(ByVal nPaneID&)
On Error GoTo ErrSection:

    m.geAnnStruct.paneId = nPaneID

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.gePaneId.Let", eGDRaiseError_Raise

End Property

Public Property Get geAnnId() As Long
On Error GoTo ErrSection:

    geAnnId = m.geAnnStruct.annotationId

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geAnnId.Let", eGDRaiseError_Raise

End Property

Public Property Let geAnnId(ByVal nIdx As Long)
On Error GoTo ErrSection:

    m.geAnnStruct.annotationId = nIdx

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geAnnId.Let", eGDRaiseError_Raise

End Property

Public Property Get geIndId() As Long
On Error GoTo ErrSection:

    geIndId = m.geAnnStruct.indicatorId

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geIndId.Let", eGDRaiseError_Raise

End Property

Public Property Let geIndId(ByVal nIdx As Long)
On Error GoTo ErrSection:

    m.geAnnStruct.indicatorId = nIdx

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geIndId.Let", eGDRaiseError_Raise

End Property

Public Property Get geTextAlign() As Long
On Error GoTo ErrSection:

    geTextAlign = m.geTextAlign

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geTextAlign.Get", eGDRaiseError_Raise

End Property

Public Property Let geTextAlign(ByVal nAlign&)
On Error GoTo ErrSection:

    m.geTextAlign = nAlign
    Prop("TextAlignment") = nAlign

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geTextAlign.Let", eGDRaiseError_Raise

End Property

Public Property Get geCycleLhsIdx() As Long
On Error GoTo ErrSection:

    'index of the first line left of the base line is always 1
    geCycleLhsIdx = 1

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geCycleLhsIdx.Get", eGDRaiseError_Raise

End Property

Public Property Get geCycleRhsIdx() As Long
On Error GoTo ErrSection:

    'item index of the first line right of the base line (set by grapheng.dll)
    geCycleRhsIdx = m.geAnnStruct.showQtrLines

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geCycleRhsIdx.Get", eGDRaiseError_Raise

End Property

Public Property Get geMinorAxisLen() As Double
On Error GoTo ErrSection:

    If m.geAnnStruct.minorAxisLenData = 0 Then
        geMinorAxisLen = m.geAnnStruct.lenRatio
    Else
        geMinorAxisLen = m.geAnnStruct.lenPoints
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geMinorAxisLen.Get", eGDRaiseError_Raise

End Property

Public Property Let geMinorAxisLen(ByVal nLen As Double)
On Error GoTo ErrSection:

    Prop("MinorAxisLen") = nLen
    If m.geAnnStruct.minorAxisLenData = 0 Then
        m.geAnnStruct.lenRatio = nLen
    Else
        m.geAnnStruct.lenPoints = nLen
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geMinorAxisLen.Let", eGDRaiseError_Raise

End Property

Public Property Get geMinorAxisLenData() As Long
On Error GoTo ErrSection:

    geMinorAxisLenData = m.geAnnStruct.minorAxisLenData

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geMinorAxisLenData.Get", eGDRaiseError_Raise

End Property

Public Property Let geMinorAxisLenData(ByVal nVal As Long)
On Error GoTo ErrSection:

    Prop("MinorAxisLenData") = nVal
    m.geAnnStruct.minorAxisLenData = nVal

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.geMinorAxisLenData.Let", eGDRaiseError_Raise

End Property

Public Property Get PropertiesArray() As cGdArray
On Error GoTo ErrSection:

    Set PropertiesArray = m.Properties

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.PropertiesArray.Get", eGDRaiseError_Raise

End Property

Public Sub geMoveIndLabel(Chart As cChart, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrSection:

    Dim xPix#, yPix#
    
    xPix = X / Screen.TwipsPerPixelX
    yPix = Y / Screen.TwipsPerPixelY
    
    gdSetNum m.geAnnStruct.gdhTop, 0, yPix
    gdSetNum m.geAnnStruct.gdhLeft, 0, xPix
    gdSetNum m.geAnnStruct.gdhRight, 0, -1

    Chart.Form.pbChart.AutoRedraw = False
    Chart.geDrawChart
    Chart.Form.pbChart.AutoRedraw = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geMoveIndLabel", eGDRaiseError_Raise

End Sub

Public Function geDrawAnn(Chart As cChart) As Long
On Error GoTo ErrSection:

    Dim rc&
       
    If Chart Is Nothing Then
        If Not m.Chart Is Nothing Then
            geSetPoints m.Chart
            m.Chart.Form.pbChart.AutoRedraw = False
            rc = m.Chart.geDrawChart()
            m.Chart.Form.pbChart.AutoRedraw = True
        End If
    Else
        geSetPoints Chart
        Chart.Form.pbChart.AutoRedraw = False
        rc = Chart.geDrawChart()
        Chart.Form.pbChart.AutoRedraw = True
    End If
    
    geDrawAnn = rc
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.geDrawAnn", eGDRaiseError_Raise

End Function

Public Sub geSetShape(ByVal nShape)
On Error GoTo ErrSection:

   If m.eType <> eANNOT_Rectangle And m.eType <> eANNOT_TextEdit And _
      m.eType <> eANNOT_TextEdit2 And m.eType <> eANNOT_TextEdit3 And m.eType <> eANNOT_TextEdit4 Then
        Exit Sub
    End If
    
    If nShape < 1 Then nShape = Val(Prop("Shape"))
    
    Select Case nShape
        Case 1
            gdSetNum m.geAnnStruct.glhAnnType, 0, 5     'round rectangle
        Case 2
            gdSetNum m.geAnnStruct.glhAnnType, 0, 6     'ellipse
        Case Else
            gdSetNum m.geAnnStruct.glhAnnType, 0, 4     'rectangle
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geSetShape", eGDRaiseError_Raise

End Sub

Public Sub geResetData()
On Error GoTo ErrSection:

    m.geAdded = False
    m.geAnnStruct.paneId = -1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geResetData", eGDRaiseError_Raise

End Sub

Private Sub geConvertPenStyle(ByVal nStyle&)
On Error GoTo ErrSection:

'This sub translates Gigasoft constants to enum type for pen styles.
'This sub is for backwards compatibility with saved annotation files.
'See commented out sub LoadStyles in frmEditAnnot.frm for constant values.

If m.nGraphEngDll = 1 Or m.eType = eANNOT_Icon Then
    m.nStyle = nStyle
    Exit Sub   'no need to convert
End If

Select Case (nStyle)
    Case PELT_THINSOLID, PEGAT_THINSOLIDLINE
        m.nStyle = eANNOT_Thin
    Case PELT_MEDIUMSOLID, PEGAT_MEDIUMSOLIDLINE
        m.nStyle = eANNOT_Medium
    Case PELT_THICKSOLID, PEGAT_THICKSOLIDLINE
        m.nStyle = eANNOT_Thick
    Case PELT_DASH, PEGAT_DASHLINE
        m.nStyle = eANNOT_DashLg
    Case PELT_DOT, PEGAT_DOTLINE
        m.nStyle = eANNOT_DashSm
    Case PELT_DASHDOT, PEGAT_DASHDOTLINE
        m.nStyle = eANNOT_DashDot
End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geConvertPenStyle", eGDRaiseError_Raise

End Sub

Private Sub geSetPenStyle(ByVal nStyle&, ByVal nIdx&)
On Error GoTo ErrSection:
        
    gdSetNum m.geAnnStruct.glhPenStyle, nIdx, 0         'assume solid pen
    
    If nStyle = 0 Then
        nStyle = g.ChartGlobals.eDefaultAnnotStyle
    ElseIf nStyle < 0 Then
        gdSetNum m.geAnnStruct.glhPenSize, nIdx, Abs(nStyle)        '5229
        Exit Sub
    End If
    
    If nStyle > eANNOT_Thick Then
        gdSetNum m.geAnnStruct.glhPenSize, nIdx, 1
        If nStyle = eANNOT_DashLg Then
            gdSetNum m.geAnnStruct.glhPenStyle, nIdx, 1
        ElseIf nStyle = eANNOT_DashSm Then
            gdSetNum m.geAnnStruct.glhPenStyle, nIdx, 2
        ElseIf nStyle = eANNOT_DashDot Then
            gdSetNum m.geAnnStruct.glhPenStyle, nIdx, 3
        End If
    Else
        gdSetNum m.geAnnStruct.glhPenStyle, nIdx, 0
        gdSetNum m.geAnnStruct.glhPenSize, nIdx, nStyle
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geSetPenStyle", eGDRaiseError_Raise

End Sub

Private Sub geTransIconTextStyle()
On Error GoTo ErrSection:

    Dim eImageType As eStockImage
    Dim eDir As eImageDir
    
    If m.eType <> eANNOT_Icon And m.eType <> eANNOT_ElliotLabel And m.eType <> eANNOT_GannacciCycle Then Exit Sub
        
    'Drawing Order: grapheng.dll draws multi-components annotations
    'based on their array index. Since we always want the text item
    'drawn first, we assign array index 0 to the text component.
    
    'styles for text are bold, italic, regular etc.
    'styles for icons are size, shape & hollow/solid of image
        
    'set properties that apply only to text
    gdSetStr m.geAnnStruct.gshFont, 0, "arial"
    If Len(m.strText) > 0 Then    'check to see if there is text
        gdSetStr m.geAnnStruct.gshText, 0, m.strText
    ElseIf m.nStyle = 0 Then     'text only annotation
        gdSetStr m.geAnnStruct.gshText, 0, "Text"
    Else
        gdSetStr m.geAnnStruct.gshText, 0, ""
    End If
    
    'set properties that apply to both text & icon
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
    
    If m.geTextAlign = 9 Then     'auto alignment
        gdSetNum m.geAnnStruct.glhAlign, 0, 3
        gdSetNum m.geAnnStruct.glhAlign, 1, 1
    Else
        gdSetNum m.geAnnStruct.glhAlign, 0, m.geTextAlign
        gdSetNum m.geAnnStruct.glhAlign, 1, 8   'icon ctrCtr
    End If

    gdSetNum m.geAnnStruct.glhAnnType, 0, 0     'text
    gdSetNum m.geAnnStruct.glhSize, 0, 0 '8
    gdSetNum m.geAnnStruct.glhStyle, 0, 1       'bold
    gdSetNum m.geAnnStruct.glhImageType, 0, 0   'this is only a place holder (ignored by grapheng.dll)
    gdSetNum m.geAnnStruct.glhDirection, 0, 0   'this is only a place holder (ignored by grapheng.dll)
    
    'Disallow text only annotation (m.nStyle=0 used to mean text only)
    If m.eUsage = eANNOT_OptionInfo Then
        m.nStyle = -1
    ElseIf m.nStyle = 0 Then
        m.nStyle = 20   'small solid diamond
    End If

    'when style=-1 there is no need to convert image for backwards compatibility
    If m.nStyle = -1 Then
        eImageType = Val(Prop("ImageType"))
        eDir = Val(Prop("ImageDir"))
'        If eImageType < eCNI_Arrow Or eImageType > eCNI_Triangle Then
'            'eImageType = eCNI_Diamond
'            Prop("ImageType") = eImageType
'        End If
        If eDir < eCNI_North Or eDir > eCNI_NorthWest Then
            eDir = eCNI_North
            Prop("ImageDir") = eDir
        End If
        gdSetNum m.geAnnStruct.glhAnnType, 1, 7     'icon
        gdSetNum m.geAnnStruct.glhImageType, 1, eImageType
        gdSetNum m.geAnnStruct.glhSize, 1, Val(Prop("ImageSize"))
        gdSetNum m.geAnnStruct.glhStyle, 1, Val(Prop("ImageStyle"))
        gdSetNum m.geAnnStruct.glhFillPattern, 1, Val(Prop("ImageStyle"))
        gdSetNum m.geAnnStruct.glhDirection, 1, eDir
        Exit Sub
    End If
    
    gdSetNum m.geAnnStruct.glhAnnType, 1, 7     'icon
    'Note: as of 09-25-2003 m.nStyle is no longer used to determine
    '   icon style & size. Icon info are saved to ImageType, ImageSize,
    '   ImageStyle & ImageDir properties instead. Code below are kept
    '   for backwards compatibility and can be removed when compatiblity
    '   is no longer a concern.
    '   See LoadAnnotIcons in frmEditAnnot for icon styles and sizes.
    Select Case (m.nStyle)
        'determine size of icon
        Case 1 To 12
            gdSetNum m.geAnnStruct.glhSize, 1, 16   'medium
        Case 13 To 24
            gdSetNum m.geAnnStruct.glhSize, 1, 8    'small (6)
        Case 25 To 36
            gdSetNum m.geAnnStruct.glhSize, 1, 24   'large
        Case 92, 94, 96, 98
            gdSetNum m.geAnnStruct.glhSize, 1, 10   'arrow (north, east, south, west)
        Case 93, 95, 97, 99
            gdSetNum m.geAnnStruct.glhSize, 1, 8    'arrow (all others)
        Case Else
            gdSetNum m.geAnnStruct.glhSize, 1, 8    'small
    End Select
       
    If m.nStyle > 36 Then    'this is an arrow, set direction then exit
        gdSetNum m.geAnnStruct.glhImageType, 1, 0
        Prop("ImageType") = eCNI_Arrow
        Prop("ImageSize") = gdGetNum(m.geAnnStruct.glhSize, 1)
        Prop("ImageStyle") = 0
        Select Case (m.nStyle)
            Case 92
                gdSetNum m.geAnnStruct.glhDirection, 1, 0   'up = north
                Prop("ImageDir") = eCNI_North
            Case 93
                gdSetNum m.geAnnStruct.glhDirection, 1, 1   'northeast
                Prop("ImageDir") = eCNI_NorthEast
            Case 94
                gdSetNum m.geAnnStruct.glhDirection, 1, 2   'right = east
                Prop("ImageDir") = eCNI_East
            Case 95
                gdSetNum m.geAnnStruct.glhDirection, 1, 3   'southeast
                Prop("ImageDir") = eCNI_SouthEast
            Case 96
                gdSetNum m.geAnnStruct.glhDirection, 1, 4   'down = south
                Prop("ImageDir") = eCNI_South
            Case 97
                gdSetNum m.geAnnStruct.glhDirection, 1, 5   'southwest
                Prop("ImageDir") = eCNI_SouthWest
            Case 98
                gdSetNum m.geAnnStruct.glhDirection, 1, 6   'left = west
                Prop("ImageDir") = eCNI_West
            Case 99
                gdSetNum m.geAnnStruct.glhDirection, 1, 7   'northwest
                Prop("ImageDir") = eCNI_NorthWest
            Case Else
                gdSetNum m.geAnnStruct.glhDirection, 1, 0
                Prop("ImageDir") = eCNI_North
        End Select
        ConvertIconDefaults
        Exit Sub
    End If
    
    'default image direction to a valid number
    gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_North
    'assume solid icons
    gdSetNum m.geAnnStruct.glhStyle, 1, 1
    Select Case (m.nStyle)
        'determine shape of icon
        Case 1, 13, 25  'plus
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Plus
        Case 2, 14, 26  'cross
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Cross
        Case 3, 15, 27 'circle
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Circle
            gdSetNum m.geAnnStruct.glhSize, 1, gdGetNum(m.geAnnStruct.glhSize, 1) + 1
            gdSetNum m.geAnnStruct.glhStyle, 1, 0
        Case 4, 16, 28 'solid circle
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Circle
            gdSetNum m.geAnnStruct.glhSize, 1, gdGetNum(m.geAnnStruct.glhSize, 1) + 1
        Case 5, 17, 29 'square
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Square
            gdSetNum m.geAnnStruct.glhStyle, 1, 0
        Case 6, 18, 30 'solid square
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Circle
        Case 7, 19, 31 'diamond
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Diamond
            gdSetNum m.geAnnStruct.glhStyle, 1, 0
        Case 8, 20, 32 'solid diamond
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Circle
        Case 9, 21, 33 'up triangle
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Triangle
            gdSetNum m.geAnnStruct.glhStyle, 1, 0
            gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_North
        Case 10, 22, 34 'solid up triangle
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Triangle
            gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_North
        Case 11, 23, 35 'down triangle
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Triangle
            gdSetNum m.geAnnStruct.glhStyle, 1, 0
            gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_South
        Case 12, 24, 36 'solid down triangle
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Triangle
            gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_South
        Case Else
            m.nStyle = 18 'small solid square
            gdSetNum m.geAnnStruct.glhImageType, 1, eCNI_Diamond
            gdSetNum m.geAnnStruct.glhSize, 1, 6
    End Select
    
    'Save to properties so don't have to convert again
    Prop("ImageType") = gdGetNum(m.geAnnStruct.glhImageType, 1)
    Prop("ImageSize") = gdGetNum(m.geAnnStruct.glhSize, 1)
    Prop("ImageStyle") = gdGetNum(m.geAnnStruct.glhStyle, 1)
    Prop("ImageDir") = gdGetNum(m.geAnnStruct.glhDirection, 1)
    
    ConvertIconDefaults
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geTransIconTextStyle", eGDRaiseError_Raise

End Sub

Private Sub geTransAnnType(Chart As cChart)
On Error GoTo ErrSection:
   
    'set common values
    gdSetStr m.geAnnStruct.gshFont, 0, "arial"
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    
    Select Case m.eUsage
        Case eANNOT_IndicatorLabel, eANNOT_SideWinderLabel
            SetIndicatorLabel Chart
        Case eANNOT_Notation
            SetNotation Chart
        Case eANNOT_Trades
            SetTrades Chart
        Case eANNOT_UserAdded
            Select Case m.eType
                Case eANNOT_SimpleLine
                    gdSetNum m.geAnnStruct.glhAnnType, 0, 1
                Case eANNOT_VertLine
                    SetVertLine Chart, 0#, 0#
                Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4
                    SetHorzLine Chart
                Case eANNOT_Rectangle, eANNOT_Mirror
                    SetRectangle Chart, 0#, 0#
                Case eANNOT_Pattern
                    SetPattern Chart, 0#, 0#, Chart.ScreenStartX
                Case eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4
                    SetTextEdit Chart, 0#, 0#
                Case eANNOT_Ellipse
                    SetEllipse Chart, 0#, 0#
                Case eANNOT_RegressionLine
                    SetRegressionLine Chart, 0#, 0#
                Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
                     eANNOT_TrendChannel
                    SetTrendLine Chart, 0#, 0#
                Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, _
                     eANNOT_DollarLine4, eANNOT_GannacciSwingSquare
                    SetDollarLine Chart, 0#, 0#
                Case eANNOT_Icon, eANNOT_ElliotLabel
                    SetIcon Chart, 0#, 0#
                Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                     eANNOT_Fibonacci3, eANNOT_FibExpansion, eANNOT_DanCodeFib, _
                     eANNOT_FibFan, eANNOT_SpResistFan, eANNOT_FibTimeRatio, eANNOT_AdvRiskReward
                    SetFibonacci Chart, 0#, 0#
                Case eANNOT_FibTimeZones, eANNOT_DanCodeZone
                    SetFibZones Chart, 0#, 0#
                Case eANNOT_FibArcs
                    SetFibArcs Chart, 0#, 0#
                Case eANNOT_TargetShooter
                    SetTargetShooter Chart
                Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
                     eANNOT_DNExpansion4, eANNOT_FibABCD
                    SetDnExpansion Chart, 0#, 0#, Chart.ScreenStartX
                Case eANNOT_DNRetracement
                    SetDnRetracement Chart, 0#, 0#, Chart.ScreenStartX
                Case eANNOT_GannLines
                    SetGannLines Chart, 0#, 0#
                Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
                    SetSRLine Chart, 0#, 0#
                Case eANNOT_MultiRects
                    SetMultiRects Chart, Chart.ScreenStartX
                Case eANNOT_WaveLabels
                    SetWaveLabels Chart, 0#, 0#, Chart.ScreenStartX
                Case eANNOT_BellAlert
                    SetAlertIcon 0, -1      '-1 places icon in upper right corner
                Case eANNOT_Bracket
                    SetBracket Chart, 0#, 0#
                Case eANNOT_Pivot
                    SetPivotPoints Chart, 0#, 0#
                Case eANNOT_Gartley
                    SetGartley Chart, 0#, 0#, Chart.ScreenStartX
                Case eANNOT_GannacciCycle
                    SetGannacciCycle Chart, 0#, 0#
                Case eANNOT_GannacciTime
                    SetGannacciTime Chart, 0#, 0#
                Case eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
                    SetGannacciSwing Chart, 0#, 0#, Chart.ScreenStartX
                Case eANNOT_BalloonStrangle
                    SetBalloonStrangle Chart, 0#, 0#
            End Select
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geTransAnnType", eGDRaiseError_Raise

End Sub

Public Sub geGetFibTimeLines(Chart As cChart, aDates As cGdArray, _
    aLines As cGdArray)
On Error GoTo ErrSection:

    Dim i&, j&, nSize&, idx&, fib&
    Dim dDate#, dDateFix#
    Dim Bars As cGdBars
        
    Set Bars = Chart.Bars
    'look for date of first line (i.e. where user clicked)
    If Bars.IsIntraday Then
        dDateFix = gdFixDateTime(m.dDate1)
    Else
        dDateFix = Int(m.dDate1)
    End If
    idx = Bars.FindDateTime(dDateFix)
    dDate = Bars(eBARS_DateTime, idx)
       
    aDates.Add dDate
    aLines.SplitFields gdGetStr(m.geAnnStruct.gshText, 0)
    If m.geAnnStruct.showQtrLines >= 3 Then aLines.Add 1, 0         'ratio, sq root & daniel code numbers
    nSize = aLines.Size
    
    If nSize > 0 Then
    
        If m.geAnnStruct.showQtrLines = 0 Then
            'fib time zone
            idx = idx + 1
            dDate = Bars(eBARS_DateTime, idx)
            If dDate > 0 Then
                aDates.Add dDate
                idx = idx + 1
                dDate = Bars(eBARS_DateTime, idx)
                If dDate > 0 Then
                    aDates.Add dDate
            
                    For i = 3 To nSize - 1
                        idx = Val(aLines(i)) + idx
                        dDate = Bars(eBARS_DateTime, idx)
                        If dDate > 0 Then
                            aDates.Add dDate
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
        Else
            If m.eType = eANNOT_DanCodeZone Then aLines(0) = "0"
        
            idx = idx - 1
            For i = 1 To nSize - 1
                fib = Val(aLines(i)) + idx
                
                If m.eType = eANNOT_DanCodeZone Then
                    dDate = Bars(eBARS_DateTime, fib + 1)
                Else
                    dDate = Bars(eBARS_DateTime, fib)
                End If
                
                If dDate > 0 Then
                    aDates.Add dDate
                Else
                    Exit For
                End If
            Next
        End If
    End If
          
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geGetFibTimeLines", eGDRaiseError_Raise

End Sub

Private Sub SpecialRatiosToGrid(fg As VSFlexGrid)
On Error GoTo ErrSection:

    Dim i&, iUse, iColor, Str$
    Dim dY1#, dY2#, dRatioD#

    Select Case m.eType

        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
             eANNOT_DNExpansion4, eANNOT_FibABCD
            'A
            iUse = Val(Prop("ShowRatioZero"))
            If iUse = 1 Then
                iUse = flexChecked
            Else
                iUse = flexUnchecked
            End If
            iColor = Val(Prop("colorA"))
            If iColor = 0 Then iColor = -1
            Str = Prop("textA")
            
            With fg
                .Rows = .Rows + 1
                i = .Rows - 1
                .Cell(flexcpChecked, i, 0) = iUse
                .TextMatrix(i, 1) = "n/a"
                .TextMatrix(i, 2) = m.Chart.Bars.PriceDisplay(m.Y1)
                .Cell(flexcpBackColor, i, 3) = iColor
                .TextMatrix(i, 4) = Str
            End With
            
            'B
            iUse = Val(Prop("ShowRatioOne"))
            If iUse = 1 Then
                iUse = flexChecked
            Else
                iUse = flexUnchecked
            End If
            iColor = Val(Prop("colorB"))
            If iColor = 0 Then iColor = -1
            Str = Prop("textB")
            
            With fg
                .Rows = .Rows + 1
                i = .Rows - 1
                .Cell(flexcpChecked, i, 0) = iUse
                .TextMatrix(i, 1) = "n/a"
                .TextMatrix(i, 2) = m.Chart.Bars.PriceDisplay(m.Y2)
                .Cell(flexcpBackColor, i, 3) = iColor
                .TextMatrix(i, 4) = Str
            End With
            
            'C
            iUse = Val(Prop("ShowRatioTwo"))
            If iUse = 1 Then
                iUse = flexChecked
            Else
                iUse = flexUnchecked
            End If
            iColor = Val(Prop("colorC"))
            If iColor = 0 Then iColor = -1
            Str = Prop("textC")
            
            With fg
                .Rows = .Rows + 1
                i = .Rows - 1
                .Cell(flexcpChecked, i, 0) = iUse
                .TextMatrix(i, 1) = "n/a"
                .TextMatrix(i, 2) = m.Chart.Bars.PriceDisplay(m.aYs(0))
                .Cell(flexcpBackColor, i, 3) = iColor
                .TextMatrix(i, 4) = Str
                .Cell(flexcpFontBold, .FixedRows, 2, i, 2) = True
                .Cell(flexcpBackColor, .FixedRows, 0, i, 2) = ALT_GRID_ROW_COLOR
            End With
            
            If m.eType = eANNOT_FibABCD Then
                dRatioD = FormatNum(Val(Prop("RatioD")))
                If dRatioD > 0 Then
                    iUse = flexChecked
                Else
                    dRatioD = -1 * dRatioD      'for display purposes (don't want to show negative ratio in editor's grid)
                    iUse = flexUnchecked
                End If
                iColor = Val(Prop("colorD"))
                If iColor = 0 Then iColor = -1
                Str = Prop("textD")
                
                With fg
                    .Rows = .Rows + 1
                    i = .Rows - 1
                    .Cell(flexcpChecked, i, 0) = iUse
                    .TextMatrix(i, 1) = FormatNum(dRatioD)
                    .TextMatrix(i, 2) = m.Chart.Bars.PriceDisplay(gdGetNum(m.geAnnStruct.gdhTop, 3))   'm.Chart.Bars.PriceDisplay(m.aYs(0))
                    .Cell(flexcpBackColor, i, 3) = iColor
                    .TextMatrix(i, 4) = Str
                    .Cell(flexcpFontBold, .FixedRows, 2, i, 2) = True
                    .Cell(flexcpBackColor, .FixedRows, 0, i, 2) = ALT_GRID_ROW_COLOR
                End With
            End If
            
        Case eANNOT_ElliotTimeRatio, eANNOT_FibTimeRatio
            iColor = m.nColor
            If iColor = 0 Then iColor = -1
            
            'A
            iUse = Val(Prop("ShowRatioZero"))
            If iUse = 1 Then
                iUse = flexChecked
            Else
                iUse = flexUnchecked
            End If
            
            With fg
                .Rows = .Rows + 1
                i = .Rows - 1
                .Cell(flexcpChecked, i, 0) = iUse
                If m.eType = eANNOT_ElliotTimeRatio Then
                    .TextMatrix(i, 1) = "Start"
                ElseIf Prop("RatioOneVal") = 1 Then
                    'Mod 04-04-2012 per request from Elliot Wave - aardvark 6646
                    .TextMatrix(i, 1) = "0"
                Else
                    'JM: email from Tim 10-14=2010 - mod to Fib Time Ratio
                    '    change UI so the 2 selected points are "-1" and "0" (first 2 rows in grid)
                    .TextMatrix(i, 1) = "-1"
                End If
            End With
        
            'B
            iUse = Val(Prop("ShowRatioOne"))
            If iUse = 1 Then
                iUse = flexChecked
            Else
                iUse = flexUnchecked
            End If
        
            With fg
                .Rows = .Rows + 1
                i = .Rows - 1
                .Cell(flexcpChecked, i, 0) = iUse
                If m.eType = eANNOT_ElliotTimeRatio Then
                    .TextMatrix(i, 1) = "End"
                ElseIf m.eType = eANNOT_FibTimeRatio Then
                    .TextMatrix(i, 1) = Val(Prop("RatioOneVal"))
                Else
                    .TextMatrix(i, 1) = "0"
                End If
            End With
            
            'C
            If m.eType = eANNOT_ElliotTimeRatio Then
                iUse = Val(Prop("ShowRatioTwo"))
                If iUse = 1 Then
                    iUse = flexChecked
                Else
                    iUse = flexUnchecked
                End If
            
                With fg
                    .Rows = .Rows + 1
                    i = .Rows - 1
                    .Cell(flexcpChecked, i, 0) = iUse
                    .TextMatrix(i, 1) = "Extend"
                End With
            End If
            
            fg.Cell(flexcpBackColor, fg.FixedRows, 0, i, fg.Cols - 1) = ALT_GRID_ROW_COLOR
        
        Case Else
            iColor = m.nColor
            If iColor = 0 Then iColor = -1
            
            If m.bY1IsZeroRatio Then
                dY1 = m.Y1
                dY2 = m.Y2
            Else
                dY1 = m.Y2
                dY2 = m.Y1
            End If
            
            '0
            iUse = Val(Prop("ShowRatioZero"))
            If iUse = 1 Then
                iUse = flexChecked
            Else
                iUse = flexUnchecked
            End If
        
            With fg
                .Rows = .Rows + 1
                i = .Rows - 1
                .Cell(flexcpChecked, i, 0) = iUse
                .TextMatrix(i, 1) = "0"
                If m.eType <> eANNOT_FibArcs Then .TextMatrix(i, 2) = m.Chart.Bars.PriceDisplay(dY1)
            End With
        
            '1
            iUse = Val(Prop("ShowRatioOne"))
            If iUse = 1 Then
                iUse = flexChecked
            Else
                iUse = flexUnchecked
            End If
    
            With fg
                .Rows = .Rows + 1
                i = .Rows - 1
                .Cell(flexcpChecked, i, 0) = iUse
                .TextMatrix(i, 1) = "1"
                If m.eType <> eANNOT_FibArcs Then .TextMatrix(i, 2) = m.Chart.Bars.PriceDisplay(dY2)
                .Cell(flexcpBackColor, .FixedRows, 0, i, .Cols - 1) = ALT_GRID_ROW_COLOR
            End With
    
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.SpecialRatiosToGrid"

End Sub

Public Sub RatiosToFibGrid(fg As VSFlexGrid, Optional ByVal bSortDescending As Boolean = False)
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lTopRowSave As Long
    
    Dim iColorCol As Long
    Dim iValueCol As Long
    Dim iLabelCol As Long
    
    Dim i&, iRow&, iTemp&, dY#, dRatio#
    Dim Str$, strLabel$
    
    Dim tb As New cGdTable
    Dim aTemp As New cGdArray
    Dim Bars As cGdBars
    
    'double-check just to be sure
    If Not IsFibType Then Exit Sub
    
    'Fib Retracement, Expansion, DanielCode & Fan has 4 columns: use,ratios,values,color
    'Fib Arcs, Time Ratio, Time Extension & Parallel for Andrews Fork has 3 columns: use,ratios,color
    'Fib Extension (non-dinapoli version) has 5 columns: use,ratios,values,color,labels
    'Time Intervals for Andrews Fork has 2 columns: use,ratios
    
    iColorCol = -1
    iValueCol = -1
    iLabelCol = -1
    
    If fg.Cols > 3 Then
        iValueCol = 2
        iColorCol = 3
    ElseIf fg.Cols > 2 Then
        iColorCol = 2
    End If
    
    If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 Or _
       m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_FibABCD Or m.eType = eANNOT_FibExpansion Or _
       m.eType = eANNOT_AdvRiskReward Then
        If fg.Cols > 4 Then iLabelCol = 4
    End If
    
    
    If m.eType = eANNOT_Gartley Then
         iLabelCol = 5
         If Not m.Chart Is Nothing Then Set Bars = m.Chart.Bars
         
         With fg
            lRedraw = .Redraw
            .Redraw = flexRDNone
            .Rows = 10
            
            'use, ratio & color columns for X,A,B,C,D cannot be edited
            .Cell(flexcpBackColor, .FixedRows, .FixedCols, 5, 1) = ALT_GRID_ROW_COLOR
            .Cell(flexcpBackColor, .FixedRows, 3, 5, 3) = ALT_GRID_ROW_COLOR
            'value & text columns for triangles cannot be edited
            .Cell(flexcpBackColor, 6, 1, 9, 2) = ALT_GRID_ROW_COLOR
            .Cell(flexcpBackColor, 6, 5, 9, 5) = ALT_GRID_ROW_COLOR
            
            'ratio/values for points
            .TextMatrix(1, 1) = "n/a"
            .TextMatrix(2, 1) = "n/a"
            .TextMatrix(3, 1) = "n/a"
            .TextMatrix(4, 1) = "n/a"
            .TextMatrix(5, 1) = "n/a"
            
            'value for points (col 2)
            If Not Bars Is Nothing Then
                .TextMatrix(1, 2) = Bars.PriceDisplay(m.Y1)
                .TextMatrix(2, 2) = Bars.PriceDisplay(m.Y2)
                
                'B,C,D
                .TextMatrix(3, 2) = Bars.PriceDisplay(m.aYs(0))
                .TextMatrix(4, 2) = Bars.PriceDisplay(m.aYs(1))
                .TextMatrix(5, 2) = Bars.PriceDisplay(m.aYs(2))
                
                'ratios for line XB, XD
                dRatio = gdGetNum(m.geAnnStruct.gdhMisc, 1)
                .TextMatrix(6, 1) = FormatNum(dRatio, 5)
                dRatio = gdGetNum(m.geAnnStruct.gdhMisc, 2)
                .TextMatrix(7, 1) = FormatNum(dRatio, 5)
                
                'ratios for line AC, BD
                dRatio = gdGetNum(m.geAnnStruct.gdhMisc, 3)
                .TextMatrix(8, 1) = FormatNum(dRatio, 5)
                dRatio = gdGetNum(m.geAnnStruct.gdhMisc, 4)
                .TextMatrix(9, 1) = FormatNum(dRatio, 5)
            End If
            
            'color column
            .Cell(flexcpBackColor, 6, 3) = Val(Prop("Triangle1_Color"))
            .Cell(flexcpBackColor, 7, 3) = Val(Prop("Triangle2_Color"))
            .Cell(flexcpBackColor, 8, 3) = Val(Prop("Triangle3_Color"))
            .Cell(flexcpBackColor, 9, 3) = Val(Prop("Triangle4_Color"))
            
            'labels for points (col 4)
            Str = Prop("GartleyLabels")
            If Len(Str) > 0 Then
                For i = 1 To 5
                    strLabel = Parse(Str, "|", i)
                    If Len(strLabel) = 0 Then
                        Select Case i
                            Case 1
                                strLabel = "X"
                            Case 2
                                strLabel = "A"
                            Case 3
                                strLabel = "B"
                            Case 4
                                strLabel = "C"
                            Case 5
                                strLabel = "D"
                        End Select
                        
                    ElseIf InStr(strLabel, " (") <> 0 Then
                        strLabel = Parse(strLabel, " (", 1)
                    End If
                    .TextMatrix(i, iLabelCol) = strLabel
                Next
            End If
            
            'use column
            .Cell(flexcpChecked, 1, 0) = flexNoCheckbox
            .Cell(flexcpChecked, 2, 0) = flexNoCheckbox
            .Cell(flexcpChecked, 3, 0) = flexNoCheckbox
            .Cell(flexcpChecked, 4, 0) = flexNoCheckbox
            .Cell(flexcpChecked, 5, 0) = flexNoCheckbox
            .Cell(flexcpChecked, 6, 0) = Val(Prop("Triangle1"))
            .Cell(flexcpChecked, 7, 0) = Val(Prop("Triangle2"))
            .Cell(flexcpChecked, 8, 0) = Val(Prop("Triangle3"))
            .Cell(flexcpChecked, 9, 0) = Val(Prop("Triangle4"))
            
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexAlignCenterCenter
            
            'triangle label from end points
            .TextMatrix(6, iLabelCol) = .TextMatrix(1, iLabelCol) & .TextMatrix(2, iLabelCol) & .TextMatrix(3, iLabelCol)
            .TextMatrix(7, iLabelCol) = .TextMatrix(1, iLabelCol) & .TextMatrix(3, iLabelCol) & .TextMatrix(5, iLabelCol)
            .TextMatrix(8, iLabelCol) = .TextMatrix(2, iLabelCol) & .TextMatrix(3, iLabelCol) & .TextMatrix(4, iLabelCol)
            .TextMatrix(9, iLabelCol) = .TextMatrix(3, iLabelCol) & .TextMatrix(4, iLabelCol) & .TextMatrix(5, iLabelCol)
            
            'triangle fill option
            .Cell(flexcpChecked, 6, 4) = Val(Prop("Triangle1_Fill"))
            .Cell(flexcpChecked, 7, 4) = Val(Prop("Triangle2_Fill"))
            .Cell(flexcpChecked, 8, 4) = Val(Prop("Triangle3_Fill"))
            .Cell(flexcpChecked, 9, 4) = Val(Prop("Triangle4_Fill"))
            
            .Cell(flexcpPictureAlignment, 6, 4, .Rows - 1, 4) = flexAlignCenterCenter
            
            
            .TextMatrix(6, 2) = "n/a"
            .TextMatrix(7, 2) = "n/a"
            .TextMatrix(8, 2) = "n/a"
            .TextMatrix(9, 2) = "n/a"
            
            .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignRightCenter
            
            
            .Redraw = lRedraw
         End With
         
         GoTo ErrExit
         
    End If
    
    
    tb.CreateField eGDARRAY_Longs, 0, "Use", -1
    tb.CreateField eGDARRAY_Doubles, 1, "Ratios", kNullData
    tb.CreateField eGDARRAY_Doubles, 2, "Value", kNullData
    tb.CreateField eGDARRAY_Longs, 3, "Color", -1
    tb.CreateField eGDARRAY_Strings, 4, "Label", "-1"
    
    'ratios
    aTemp.SetArrayHandle tb.FieldArrayHandle(1), False
    
    If m.eType = eANNOT_AndrewFork Then
        If fg.Cols = 2 Then
            aTemp.SplitFields Prop("MarkerRatios")
            If aTemp.Size = 0 Then
                AndrewsDefaults "Markers", True       'fork read in from old template
                aTemp.SplitFields Prop("MarkerRatios")
            End If
        Else
            aTemp.SplitFields Prop("ParallelRatios")
            If aTemp.Size = 0 Then
                AndrewsDefaults "Parallel", True
                aTemp.SplitFields Prop("ParallelRatios")
            End If
        End If
    Else
        aTemp.SplitFields Prop("FibRatios"), ","
    End If
    iTemp = aTemp.Size
    If iTemp = 0 Then GoTo ErrExit              'theoretically should not happen
    
    'use
    aTemp.SetArrayHandle tb.FieldArrayHandle(0), False
    
    If m.eType = eANNOT_AndrewFork Then
        If fg.Cols = 2 Then
            aTemp.SplitFields Prop("MarkerShow")
        Else
            aTemp.SplitFields Prop("ParallelShow")
        End If
    Else
        aTemp.SplitFields Prop("ShowRatios"), ","
    End If
    
    If aTemp.Size > iTemp Then
        aTemp.Size = iTemp          'just truncate
    ElseIf aTemp.Size < iTemp Then
        For i = aTemp.Size To iTemp - 1
            aTemp(i) = 0
        Next
    End If
    If aTemp.Size <> iTemp Then GoTo ErrExit        'precautionary, theoretically should never happen

    'color
    aTemp.SetArrayHandle tb.FieldArrayHandle(3), False
    
    If m.eType = eANNOT_AndrewFork Then
        If fg.Cols > 2 Then aTemp.SplitFields Prop("ParallelColor"), ","
    Else
        aTemp.SplitFields Prop("FibColor"), ","
    End If
    
    If aTemp.Size > iTemp Then
        aTemp.Size = iTemp          'just truncate
    ElseIf aTemp.Size < iTemp Then
        For i = aTemp.Size To iTemp - 1
            aTemp(i) = vbRed        'use default color
        Next
    End If
    If aTemp.Size <> iTemp Then GoTo ErrExit        'precautionary, theoretically should never happen

    'values
    aTemp.SetArrayHandle tb.FieldArrayHandle(2), False
    aTemp.Size = iTemp
    'labels
    aTemp.SetArrayHandle tb.FieldArrayHandle(4), False
    If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 _
       Or m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_FibABCD Or m.eType = eANNOT_FibExpansion _
       Or m.eType = eANNOT_AdvRiskReward Then
        aTemp.SplitFields Prop("FibRatiosText"), ","
        If aTemp.Size <> iTemp Then aTemp.Size = iTemp
    Else
        aTemp.Size = iTemp
    End If

    tb.NumRecords = iTemp
    Set aTemp = Nothing

    If bSortDescending Then
        Set aTemp = tb.CreateSortedIndex(1, eGdSort_Descending Or eGdSort_Stable)
    Else
        Set aTemp = tb.CreateSortedIndex(1, eGdSort_Stable)
    End If

    If aTemp Is Nothing Then GoTo ErrExit        'precautionary, theoretically should never happen
    
    With fg
        lTopRowSave = .TopRow
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        'handle special ratios (e.g. zero, one, A, B, C etc.)
        If m.eType <> eANNOT_AndrewFork Then SpecialRatiosToGrid fg
        
        For i = 0 To aTemp.Size - 1
            .Rows = .Rows + 1
            iRow = .Rows - 1
            
            iTemp = tb(0, aTemp(i))
            If iTemp = 1 Then
                .Cell(flexcpChecked, iRow, 0) = flexChecked     'use
            Else
                .Cell(flexcpChecked, iRow, 0) = flexUnchecked
            End If
            .TextMatrix(iRow, 1) = FormatNum(tb(1, aTemp(i)))   'ratio
            
            If iValueCol > 0 Then
                If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or _
                   m.eType = eANNOT_DNExpansion3 Or m.eType = eANNOT_DNExpansion4 Then
                    dY = gdGetNum(m.geAnnStruct.gdhTop, i + 3)
                ElseIf m.eType = eANNOT_FibABCD Then
                    dY = gdGetNum(m.geAnnStruct.gdhTop, i + 4)
                ElseIf m.eType = eANNOT_FibFan Then
                    If m.dDate1 < m.dDate2 Then
                        dY = gdGetNum(m.geAnnStruct.gdhBottom, i + 6)
                    Else
                        dY = gdGetNum(m.geAnnStruct.gdhTop, i + 6)
                    End If
                Else
                    dY = gdGetNum(m.geAnnStruct.gdhTop, i + 6)
                End If
                .TextMatrix(iRow, iValueCol) = m.Chart.Bars.PriceDisplay(dY)    'y-value
            End If
            
            If iColorCol > 1 Then
                iTemp = tb(3, aTemp(i))
                If iTemp = 0 Then iTemp = -1
                .Cell(flexcpBackColor, iRow, iColorCol) = iTemp     'color
            End If
            
            If iLabelCol > 1 Then
                Str = tb(4, aTemp(i))
                If Len(Str) > 0 Then .TextMatrix(iRow, iLabelCol) = Str
            End If
        Next
        
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexPicAlignCenterCenter
        .Cell(flexcpAlignment, .FixedRows, 1, .Rows - 1, 1) = flexAlignRightCenter
        
        If lTopRowSave >= .FixedRows And lTopRowSave < .Rows Then .TopRow = lTopRowSave
        
        If m.eType = eANNOT_AdvRiskReward Then 'eANNOT_FibExpansion Then
            .TextMatrix(1, 4) = Prop("textOne")
            .TextMatrix(2, 4) = Prop("textZero")
        End If
        
        .Redraw = lRedraw
    End With
    
    
ErrExit:
    Set tb = Nothing
    Set aTemp = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.RatiosToFibGrid"

End Sub

Private Function SpecialRatiosToProp(fg As VSFlexGrid) As Long
On Error GoTo ErrSection:

    Dim i&, iUse&, iColor&, dRatioD#, s$
    Dim Bars As cGdBars
    
    If Not m.Chart Is Nothing Then Set Bars = m.Chart.Bars
    
    If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 Or _
       m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_FibABCD Then
        With fg
            'A
            iUse = 0
            i = .FixedRows
            If .Cell(flexcpChecked, i, 0) = flexChecked Then iUse = 1
            iColor = .Cell(flexcpBackColor, i, 3)
            If iColor = -1 Then iColor = 0
            s = Left(.TextMatrix(i, 4), 8)          'aardvark 6896
            
            Prop("ShowRatioZero") = iUse
            Prop("colorA") = iColor
            Prop("textA") = s
            'B
            iUse = 0
            i = .FixedRows + 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then iUse = 1
            iColor = .Cell(flexcpBackColor, i, 3)
            If iColor = -1 Then iColor = 0
            s = Left(.TextMatrix(i, 4), 8)          'aardvark 6896
            
            Prop("ShowRatioOne") = iUse
            Prop("colorB") = iColor
            Prop("textB") = s
            'C
            iUse = 0
            i = .FixedRows + 2
            If .Cell(flexcpChecked, i, 0) = flexChecked Then iUse = 1
            iColor = .Cell(flexcpBackColor, i, 3)
            If iColor = -1 Then iColor = 0
            s = Left(.TextMatrix(i, 4), 8)          'aardvark 6896
            
            Prop("ShowRatioTwo") = iUse
            Prop("colorC") = iColor
            Prop("textC") = s
            
            If InStr(.TextMatrix(1, 2), "^") > 0 Then
                If Not Bars Is Nothing Then
                    m.Y1 = Bars.PriceFromString(.TextMatrix(1, 2))          '6174
                    m.Y2 = Bars.PriceFromString(.TextMatrix(2, 2))
                    m.aYs(0) = Bars.PriceFromString(.TextMatrix(3, 2))
                End If
            Else
                m.Y1 = ValOfText(.TextMatrix(1, 2))
                m.Y2 = ValOfText(.TextMatrix(2, 2))
                m.aYs(0) = ValOfText(.TextMatrix(3, 2))
            End If
            
            'D
            If m.eType = eANNOT_FibABCD Then
                i = .FixedRows + 3
                dRatioD = ValOfText(.TextMatrix(i, 1))
                If .Cell(flexcpChecked, i, 0) = flexUnchecked Then dRatioD = -1 * dRatioD
                iColor = .Cell(flexcpBackColor, i, 3)
                If iColor = -1 Then iColor = 0
                s = Left(.TextMatrix(i, 4), 8)      'aardvark 6896
                
                Prop("RatioD") = dRatioD
                Prop("colorD") = iColor
                Prop("textD") = s
            End If
        End With
    Else
        With fg
            '0 ( or -1 for Fib Time Ratios )
            iUse = 0
            i = .FixedRows
            If .Cell(flexcpChecked, i, 0) = flexChecked Then iUse = 1
            Prop("ShowRatioZero") = iUse
            
            '1 ( or -1 for Fib Time Ratios )
            iUse = 0
            i = .FixedRows + 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then iUse = 1
            Prop("ShowRatioOne") = iUse
            Prop("RatioOneVal") = ValOfText(.TextMatrix(i, 1))
            
            If m.eType = eANNOT_ElliotTimeRatio Then
                iUse = 0
                i = .FixedRows + 2
                If .Cell(flexcpChecked, i, 0) = flexChecked Then iUse = 1
                Prop("ShowRatioTwo") = iUse
            End If
            
            '6144
            If m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 Or m.eType = eANNOT_Fibonacci3 _
               Or m.eType = eANNOT_Fibonacci4 Or m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_DanCodeFib _
               Or m.eType = eANNOT_AdvRiskReward Then
                
                s = .TextMatrix(.FixedRows, 2)
                If InStr(s, "^") > 0 Then
                    If m.bY1IsZeroRatio Then
                        m.Y1 = Bars.PriceFromString(s)        '6174
                        s = .TextMatrix(.FixedRows + 1, 2)
                        m.Y2 = Bars.PriceFromString(s)
                    Else
                        m.Y2 = Bars.PriceFromString(s)
                        s = .TextMatrix(.FixedRows + 1, 2)
                        m.Y1 = Bars.PriceFromString(s)
                    End If
                Else
                    If m.bY1IsZeroRatio Then
                        m.Y1 = ValOfText(.TextMatrix(.FixedRows, 2))
                        m.Y2 = ValOfText(.TextMatrix(.FixedRows + 1, 2))
                    Else
                        m.Y2 = ValOfText(.TextMatrix(.FixedRows, 2))
                        m.Y1 = ValOfText(.TextMatrix(.FixedRows + 1, 2))
                    End If
                End If
                
                If m.eType = eANNOT_AdvRiskReward Then
                    s = .TextMatrix(1, 4)
                    Prop("textOne") = s
                    s = .TextMatrix(2, 4)
                    Prop("textZero") = s
                End If
                
            End If
        
        End With
    End If
            
    
    SpecialRatiosToProp = i + 1

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.SpecialRatiosToProp"

End Function

Public Sub FibGridToRatios(fg As VSFlexGrid)
On Error GoTo ErrSection:

    Dim i&, j&, iRowStart&, iColorCol&
    Dim strShow$, strRatios$, strColor$, strLabel$
    
    If fg.Cols > 3 Then
        iColorCol = 3
    ElseIf fg.Cols > 2 Then
        iColorCol = 2
    End If
    
    If m.eType = eANNOT_Gartley Then
        With fg
            If .Rows >= 6 Then
                i = .Cell(flexcpBackColor, 6, 3)
                If i = -1 Then i = 0    '-1 is special grid value for black
                Prop("Triangle1_Color") = i
                
                i = .Cell(flexcpBackColor, 7, 3)
                If i = -1 Then i = 0
                Prop("Triangle2_Color") = i
                
                i = .Cell(flexcpBackColor, 8, 3)
                If i = -1 Then i = 0
                Prop("Triangle3_Color") = i
                
                i = .Cell(flexcpBackColor, 9, 3)
                If i = -1 Then i = 0
                Prop("Triangle4_Color") = i

                'triangle fill property
                Prop("Triangle1_Fill") = .Cell(flexcpChecked, 6, 4)
                Prop("Triangle2_Fill") = .Cell(flexcpChecked, 7, 4)
                Prop("Triangle3_Fill") = .Cell(flexcpChecked, 8, 4)
                Prop("Triangle4_Fill") = .Cell(flexcpChecked, 9, 4)
                
                'triangle show
                Prop("Triangle1") = .Cell(flexcpChecked, 6, 0)
                Prop("Triangle2") = .Cell(flexcpChecked, 7, 0)
                Prop("Triangle3") = .Cell(flexcpChecked, 8, 0)
                Prop("Triangle4") = .Cell(flexcpChecked, 9, 0)
                
                'make sure there is a label in the label column (this is needed to identify triangles
                If Len(.TextMatrix(1, 5)) = 0 Or Len(.TextMatrix(2, 5)) = 0 Or _
                   Len(.TextMatrix(3, 5)) = 0 Or Len(.TextMatrix(4, 5)) = 0 Or _
                   Len(.TextMatrix(5, 5)) = 0 Then
                   
                   If Len(.TextMatrix(1, 5)) = 0 Then .TextMatrix(1, 5) = "X"
                   If Len(.TextMatrix(2, 5)) = 0 Then .TextMatrix(2, 5) = "A"
                   If Len(.TextMatrix(3, 5)) = 0 Then .TextMatrix(3, 5) = "B"
                   If Len(.TextMatrix(4, 5)) = 0 Then .TextMatrix(4, 5) = "C"
                   If Len(.TextMatrix(5, 5)) = 0 Then .TextMatrix(5, 5) = "D"
                   
                   InfBox "Text label for Gartley endpoint cannot be blank. It is needed to identify triangles"
                End If
                
                'triangle label from end points
                .TextMatrix(6, 5) = .TextMatrix(1, 5) & .TextMatrix(2, 5) & .TextMatrix(3, 5)
                .TextMatrix(7, 5) = .TextMatrix(1, 5) & .TextMatrix(3, 5) & .TextMatrix(5, 5)
                .TextMatrix(8, 5) = .TextMatrix(2, 5) & .TextMatrix(3, 5) & .TextMatrix(4, 5)
                .TextMatrix(9, 5) = .TextMatrix(3, 5) & .TextMatrix(4, 5) & .TextMatrix(5, 5)
            End If
            
            strLabel = ""
            Prop("GartleyLabels") = ""       'reset
            i = Int(Val(Prop("GartleyLabelStyle")))
            If i = 2 Then
                strLabel = .TextMatrix(1, 5) & " (" & .TextMatrix(1, 2) & ")|"
                strLabel = strLabel & .TextMatrix(2, 5) & " (" & .TextMatrix(2, 2) & ")|"
                strLabel = strLabel & .TextMatrix(3, 5) & " (" & .TextMatrix(3, 2) & ")|"
                strLabel = strLabel & .TextMatrix(4, 5) & " (" & .TextMatrix(4, 2) & ")|"
                strLabel = strLabel & .TextMatrix(5, 5) & " (" & .TextMatrix(5, 2) & ")"
            Else
                strLabel = .TextMatrix(1, 5) & "|"
                strLabel = strLabel & .TextMatrix(2, 5) & "|"
                strLabel = strLabel & .TextMatrix(3, 5) & "|"
                strLabel = strLabel & .TextMatrix(4, 5) & "|"
                strLabel = strLabel & .TextMatrix(5, 5)
            End If
            
            Prop("GartleyLabels") = strLabel
        
        End With
        GoTo ErrExit
        
    ElseIf m.eType = eANNOT_AndrewFork Then
        iRowStart = fg.FixedRows
    Else
        iRowStart = SpecialRatiosToProp(fg)
        If iRowStart <= fg.FixedRows + 1 Or iRowStart >= fg.Rows Then GoTo ErrExit      '6531
    End If
    
    With fg
        For i = iRowStart To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                strShow = strShow & "1,"
            Else
                strShow = strShow & "0,"
            End If
            '6177 Italian regional setting issue:
            'must explicitly do Str() on value returned from ValOfText that uses period (.) as decimal point
            strRatios = strRatios & Str(ValOfText(.TextMatrix(i, 1))) & ","
            strColor = strColor & Int(Val(.Cell(flexcpBackColor, i, iColorCol))) & ","
            
            'aardvark 6896 - allow up to 8 characters for ratios labels
            If .Cols > 4 Then strLabel = strLabel & Left(.TextMatrix(i, 4), 8) & ","        'only ABC & ABCD tools has label
        Next
    End With
    
    'strip off last comma
    strShow = Left(strShow, Len(strShow) - 1)
    strRatios = Left(strRatios, Len(strRatios) - 1)
    strColor = Left(strColor, Len(strColor) - 1)
    
    strColor = Replace(strColor, "-1", "0")     'flex grid uses -1 for black, change this to standard RGB value for black
    
    If m.eType = eANNOT_AndrewFork Then
        If fg.Cols = 2 Then
            Prop("MarkerShow") = strShow
            Prop("MarkerRatios") = strRatios
        Else
            Prop("ParallelShow") = strShow
            Prop("ParallelRatios") = strRatios
            Prop("ParallelColor") = strColor
        End If
    Else
        Prop("ShowRatios") = strShow
        Prop("FibRatios") = strRatios
        Prop("FibColor") = strColor
    End If
    
    If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 _
       Or m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_FibABCD Or m.eType = eANNOT_FibExpansion _
       Or m.eType = eANNOT_AdvRiskReward Then
       
       Prop("FibRatiosText") = strLabel
    End If

    m.aRatios.SplitFields strRatios, ","
    m.aRatioShow.SplitFields strShow, ","
    m.aRatioColor.SplitFields strColor, ","

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.FibGridToRatios"

End Sub

Public Sub geSetAndrewForkChange(aYs As cGdArray, aDates As cGdArray)
On Error GoTo ErrSection:


    m.dDate1 = aDates(0)
    m.dDate2 = aDates(1)
    m.aDates(0) = aDates(2)
    
    m.Y1 = aYs(0)
    m.Y2 = aYs(1)
    m.aYs(0) = aYs(2)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geSetAndrewForkChange", eGDRaiseError_Raise

End Sub

Private Sub geSetTgtShooter(Chart As cChart, ByVal dX1#, ByVal dX2#)
On Error GoTo ErrSection:

    Dim bPullbacks As Boolean
    Dim nColor&, nPenStyle&, i&
    Dim dtsY1#, dtsY2#, dtsX1#, dtsX2#
    Dim dMult#, ExtX#   'x value for target points
    Dim dtsHzY#         'y value for horizontal line
    Dim dVal#           'temporary variable use to set tgt pts & text positions
    Dim nLineType&      'use to toggle annotation type for show/hide lines
    Dim nTextType1&     'use to toggle annotation type for show/hide text of upper tgt lines
    Dim nTextType2&     'use to toggle annotation type for show/hide text of lower tgt lines

    nLineType = -1
    dMult = Val(Prop("ExtMult"))
    If dMult <= 0 Then dMult = 2 '(default)
    If Val(Prop("Pullbacks")) <> 0 Then bPullbacks = True
          
    If dX1 <= dX2 Then
        dtsX1 = dX1
        dtsX2 = dX2
        dtsY1 = m.Y1
        dtsY2 = m.Y2
        dtsHzY = m.Y1
        ExtX = dX2 + (dX2 - dX1 + 1) * dMult
        If dtsX1 = dtsX2 Then nLineType = 1
    ElseIf dX1 > dX2 Then
        dtsX1 = dX2
        dtsX2 = dX1
        dtsY1 = m.Y2
        dtsY2 = m.Y1
        dtsHzY = m.Y2
        ExtX = dX1 + (dX1 - dX2 + 1) * dMult
    End If
        
    'set pen color & size for dashed line, horz line & circles
    For i = 0 To 3
        gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
        gdSetNum m.geAnnStruct.glhPenSize, i, 1
    Next
    'dash line
    gdSetNum m.geAnnStruct.glhPenStyle, 0, 2    'always small dash style
    gdSetNum m.geAnnStruct.gdhTop, 0, dtsY1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dtsX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, dtsY2
    gdSetNum m.geAnnStruct.gdhRight, 0, dtsX2
                    
    'circles
    gdSetNum m.geAnnStruct.glhPenStyle, 1, 0    'always solid
    gdSetNum m.geAnnStruct.gdhTop, 1, dtsY1
    gdSetNum m.geAnnStruct.gdhLeft, 1, dtsX1
    gdSetNum m.geAnnStruct.gdhBottom, 1, dtsY1
    gdSetNum m.geAnnStruct.gdhRight, 1, dtsX1
    
    gdSetNum m.geAnnStruct.glhPenStyle, 2, 0    'always solid
    gdSetNum m.geAnnStruct.gdhTop, 2, dtsY2
    gdSetNum m.geAnnStruct.gdhLeft, 2, dtsX2
    gdSetNum m.geAnnStruct.gdhBottom, 2, dtsY2
    gdSetNum m.geAnnStruct.gdhRight, 2, dtsX2
    
    'horz line
    geSetPenStyle m.nStyle, 3
    gdSetNum m.geAnnStruct.gdhLeft, 3, dtsX1
    gdSetNum m.geAnnStruct.gdhRight, 3, dtsX2
    gdSetNum m.geAnnStruct.gdhTop, 3, dtsHzY
    gdSetNum m.geAnnStruct.gdhBottom, 3, dtsHzY
    
    'set pen style & size for vertical & target lines & text
    For i = 4 To 17
        gdSetNum m.geAnnStruct.glhPenStyle, i, 0
        gdSetNum m.geAnnStruct.glhPenSize, i, 1
    Next
    'vertical line 1
    gdSetNum m.geAnnStruct.glhPenColor, 4, m.nColor
    geSetPenStyle m.nStyle, 4
    If dX1 = dX2 Then
        gdSetNum m.geAnnStruct.gdhTop, 4, dtsY1
    Else
        gdSetNum m.geAnnStruct.gdhTop, 4, dtsY2
    End If
    gdSetNum m.geAnnStruct.gdhLeft, 4, dtsX2
    gdSetNum m.geAnnStruct.gdhBottom, 4, gdGetNum(m.geAnnStruct.gdhTop, 8)
    gdSetNum m.geAnnStruct.gdhRight, 4, dtsX2
    
    'vertical line 2
    gdSetNum m.geAnnStruct.glhPenColor, 5, m.nColor
    geSetPenStyle m.nStyle, 5
    gdSetNum m.geAnnStruct.gdhTop, 5, dtsY2
    gdSetNum m.geAnnStruct.gdhLeft, 5, dtsX2
    gdSetNum m.geAnnStruct.gdhBottom, 5, gdGetNum(m.geAnnStruct.gdhTop, 11)
    gdSetNum m.geAnnStruct.gdhRight, 5, dtsX2
           
    'if pullbacks then draw from Y1 to Y2 or vice versa
    If bPullbacks = True Then
        gdSetNum m.geAnnStruct.gdhBottom, 4, dtsY1
        gdSetNum m.geAnnStruct.gdhBottom, 5, dtsY1
    End If
    
    'target lines & text
    nPenStyle = Val(Prop("TargetStyle"))
    nColor = Val(Prop("TargetColor"))
    For i = 6 To 11
        dVal = gdGetNum(m.geAnnStruct.gdhTop, i)
        gdSetNum m.geAnnStruct.glhPenColor, i, nColor
        geSetPenStyle nPenStyle, i
        gdSetNum m.geAnnStruct.gdhBottom, i, dVal
        gdSetNum m.geAnnStruct.gdhLeft, i, dtsX2
        gdSetNum m.geAnnStruct.gdhRight, i, ExtX
        'text
        gdSetNum m.geAnnStruct.glhPenColor, i + 6, nColor
        gdSetNum m.geAnnStruct.gdhTop, i + 6, dVal
        gdSetNum m.geAnnStruct.gdhLeft, i + 6, dtsX2
        gdSetNum m.geAnnStruct.gdhBottom, i + 6, dVal
        gdSetNum m.geAnnStruct.gdhRight, i + 6, dtsX2
        gdSetStr m.geAnnStruct.gshText, i + 6, Chart.PriceDisplay(m.geAnnStruct.paneId, dVal)
    Next
                                    
    'determine what to show
    nTextType1 = -1
    nTextType2 = -1
    If Val(Prop("ShowValues")) > 0 Then
        nTextType1 = 0
        If nLineType = 1 Then nTextType2 = 0
    End If
    
    gdSetNum m.geAnnStruct.glhAnnType, 5, nLineType
    'check property and reset as needed
    If Val(Prop("Pt1")) = 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 6, -1
        gdSetNum m.geAnnStruct.glhAnnType, 9, -1
        gdSetNum m.geAnnStruct.glhAnnType, 12, -1
        gdSetNum m.geAnnStruct.glhAnnType, 15, -1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 6, 1
        gdSetNum m.geAnnStruct.glhAnnType, 9, nLineType
        gdSetNum m.geAnnStruct.glhAnnType, 12, nTextType1
        gdSetNum m.geAnnStruct.glhAnnType, 15, nTextType2
    End If
    If Val(Prop("Pt2")) = 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 7, -1
        gdSetNum m.geAnnStruct.glhAnnType, 10, -1
        gdSetNum m.geAnnStruct.glhAnnType, 13, -1
        gdSetNum m.geAnnStruct.glhAnnType, 16, -1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 7, 1
        gdSetNum m.geAnnStruct.glhAnnType, 10, nLineType
        gdSetNum m.geAnnStruct.glhAnnType, 13, nTextType1
        gdSetNum m.geAnnStruct.glhAnnType, 16, nTextType2
    End If
    If Val(Prop("Pt3")) = 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 8, -1
        gdSetNum m.geAnnStruct.glhAnnType, 11, -1
        gdSetNum m.geAnnStruct.glhAnnType, 14, -1
        gdSetNum m.geAnnStruct.glhAnnType, 17, -1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 8, 1
        gdSetNum m.geAnnStruct.glhAnnType, 11, nLineType
        gdSetNum m.geAnnStruct.glhAnnType, 14, nTextType1
        gdSetNum m.geAnnStruct.glhAnnType, 17, nTextType2
    End If
        
    'Q & A - what are Prop("Pt4", "Pt5" & "Pt6") for?

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geSetTgtShooter", eGDRaiseError_Raise

End Sub

Private Sub geSetPoints(Optional Chart As cChart = Nothing, Optional ByVal nStartScreenX& = -999999)
On Error GoTo ErrSection:

    Dim dX1#, dX2#, dOptDateSave#
    Dim nColor&, nPenStyle&
    Dim i&, iStartX&
    Dim Pane As cPane
    
    If Chart Is Nothing Then Exit Sub
    If m.eType = eANNOT_ElliotEwave Then Exit Sub
    
    ' TLB 4/3/2012: new option for semi-log charts to drawing Fib lines as visual ratios
    ' (set True only if global Log drawing mode is set and this Annot is in a semi-log pane)
    m.bUsingLogDrawingMode = False
    If g.ChartGlobals.bLogModeDraw And Not Chart Is Nothing Then
        Set Pane = Chart.Tree(m.strPane)
        If Not Pane Is Nothing Then
            If Pane.PaneLogFlag = ePANE_LogFlagLog Then
                m.bUsingLogDrawingMode = True
            End If
            Set Pane = Nothing
        End If
    End If
    
    If m.eUsage = eANNOT_OptionInfo Then
        'm.dDate1 = option expiration date; m.dDate2 = option month
        dOptDateSave = m.dDate2
    ElseIf m.eUsage = eANNOT_IndicatorLabel Or m.eUsage = eANNOT_SideWinderLabel Then
        SetIndicatorLabel Chart
        GoTo ErrExit
    ElseIf m.nHide = 1 Then
        For i = 0 To gdGetSize(m.geAnnStruct.glhAnnType) - 1
            gdSetNum m.geAnnStruct.glhAnnType, i, -1
        Next
        GoTo ErrExit
    End If
       
    If nStartScreenX <> -999999 Then
        iStartX = nStartScreenX
    Else
        iStartX = Chart.ScreenStartX
    End If
    
    If m.eType = eANNOT_FibTimeZones Or m.eType = eANNOT_DanCodeZone Then
'JM 5/23/2008:
'While implementing the greenblatt time zones feature, noticed that if a fib time zone was drawn on an
'intraday chart then user switched to non-intraday, the annotation editor always showed the date one day off.
'Example: Draw on ES-067 5/21/2008 23:30. Switch chart to daily. The editor will show that the first line
'is drawn on 5/22 when in fact it is drawn on the 5/21 line. Do not know if this is an issue with other
'draw tools or not. To be on the safe side, only making this change for fib time zones.
        If Chart.Bars.IsIntraday Then
            dX1 = GetX(Chart, m.dDate1) - iStartX '+ 1
            dX2 = GetX(Chart, m.dDate2) - iStartX '+ 1
        Else
            dX1 = GetX(Chart, Int(m.dDate1)) - iStartX '+ 1
            dX2 = GetX(Chart, Int(m.dDate2)) - iStartX '+ 1
        End If
    Else
        If m.eType = eANNOT_VertLine Then
            If m.eUsage = eANNOT_OptionInfo Then
                 m.dDate2 = m.dDate1
            ElseIf Int(m.dDate1) = Int(m.dDate2) And m.dDate1 <> m.dDate2 Then
                m.dDate2 = m.dDate1         'aardvark 4537
            End If
        End If
        
        dX1 = GetX(Chart, m.dDate1) - iStartX '+ 1
        dX2 = GetX(Chart, m.dDate2) - iStartX '+ 1
    
    End If
    i = 0
    ' Get start and end points
    m.X1 = GetX(Chart, m.dDate1)
    m.X2 = GetX(Chart, m.dDate2)

    'check whether to show annotation: if 2 points are on same bar ...
    If dX1 = dX2 Then
        Select Case m.eType
        ' for DiNapoli Expansion, don't show if 3 points are on the same bar
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, eANNOT_DNExpansion4, _
             eANNOT_FibABCD, eANNOT_GannacciSwing2
            If dX1 = GetX(Chart, m.aDates(0)) - iStartX Then
                m.geAnnStruct.paneId = -1   'don't show
                GoTo ErrExit
            End If
            
        Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4, eANNOT_Pivot, _
             eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
            'do nothing - dates don't matter    (aardvark 6675)
            
        Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4, _
             eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
             eANNOT_TrendChannel, eANNOT_TargetShooter, eANNOT_TriangleWedge, _
             eANNOT_AndrewFork, eANNOT_GannacciSwingSquare    '4212,4217,4254
            'do nothing if Y1 & Y2 are different it is okay to draw as vertical line on same bar
            If m.dDate1 <> m.dDate2 And m.Y1 = m.Y2 Then
                m.geAnnStruct.paneId = -1   'don't show
                GoTo ErrExit
            End If
                                
        ' for all others: don't show if 2 points are on the same bar but
        ' have different dates (i.e. came from different bars in a less
        ' compressed bar period but now being compressed to the same bar)
        Case Else
            If m.dDate1 <> m.dDate2 Then
                m.geAnnStruct.paneId = -1   'don't show
                GoTo ErrExit
            End If
        End Select
    End If
    
    'set pre-indicator flag
    m.geAnnStruct.PreIndicator = m.nPreIndicator
    
    If m.eUsage = eANNOT_Notation Or m.eUsage = eANNOT_Trades Then
        gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        If m.eUsage = eANNOT_Trades And gdGetSize(m.geAnnStruct.glhAnnType) > 1 Then
            ' put order text on next bar so will not overlap the triangle
            gdSetNum m.geAnnStruct.gdhLeft, 1, dX1 + 1
            gdSetNum m.geAnnStruct.gdhRight, 1, dX1 + 1
            gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
            gdSetNum m.geAnnStruct.gdhBottom, 1, m.Y1
        End If
        GoTo ErrExit
    End If
    
    Select Case m.eType
    'items that do not need to handle log draw mode
        Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4
            SetHorzLine Chart
        Case eANNOT_VertLine
            SetVertLine Chart, dX1, dX2
        Case eANNOT_SimpleLine
            SetSimpleLine Chart, dX1, dX2       '04-04-2012 used only for what-if mode
        Case eANNOT_Pattern
            SetPattern Chart, dX1, dX2, iStartX
        Case eANNOT_RegressionLine
            SetRegressionLine Chart, dX1, dX2
        Case eANNOT_Icon, eANNOT_ElliotLabel
            SetIcon Chart, dX1, dX2
        Case eANNOT_DNRetracement
            SetDnRetracement Chart, dX1, dX2, iStartX
        Case eANNOT_TimeCycle
            SetTimeCycle Chart, dX1, dX2
        Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
            SetSRLine Chart, dX1, dX2
        Case eANNOT_MultiRects
            SetMultiRects Chart, iStartX        '04-04-2012 used only for highlight boxes
        Case eANNOT_BellAlert
            SetAlertIcon 0, -1
        Case eANNOT_GannacciCycle
            SetGannacciCycle Chart, dX1, dX2
        Case eANNOT_GannacciTime
            SetGannacciTime Chart, dX1, dX2
        Case eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
            SetGannacciSwing Chart, dX1, dX2, iStartX
        Case eANNOT_WaveLabels
            SetWaveLabels Chart, dX1, dX2, iStartX
        Case eANNOT_Pivot
            SetPivotPoints Chart, dX1, dX2
        Case eANNOT_BalloonStrangle
            SetBalloonStrangle Chart, dX1, dX2

    'items that already have code to handle log draw mode
        Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
             eANNOT_TrendChannel
            SetTrendLine Chart, dX1, dX2
        Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, _
             eANNOT_DollarLine4, eANNOT_GannacciSwingSquare
            SetDollarLine Chart, dX1, dX2
        Case eANNOT_Rectangle, eANNOT_Mirror
            SetRectangle Chart, dX1, dX2
        Case eANNOT_Bracket
            SetBracket Chart, dX1, dX2
        Case eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4, eANNOT_ArrowLine
            SetTextEdit Chart, dX1, dX2
        Case eANNOT_Ellipse
            SetEllipse Chart, dX1, dX2
        Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
            eANNOT_Fibonacci4, eANNOT_FibExpansion, eANNOT_DanCodeFib, _
            eANNOT_FibFan, eANNOT_SpResistFan, eANNOT_FibTimeRatio, eANNOT_AdvRiskReward
            SetFibonacci Chart, dX1, dX2
        Case eANNOT_FibTimeZones, eANNOT_DanCodeZone
            SetFibZones Chart, dX1, dX2
        Case eANNOT_FibArcs
            SetFibArcs Chart, dX1, dX2
        Case eANNOT_ElliotTimeRatio
            SetETRatio Chart, dX1, dX2, iStartX
        Case eANNOT_TriangleWedge
            SetTriangleWedge Chart, dX1, dX2, iStartX
        Case eANNOT_GannLines
            SetGannLines Chart, dX1, dX2
        Case eANNOT_ChannelHighlight
            SetChannelHighlight Chart, dX1, dX2, iStartX
        Case eANNOT_RiskReward
            SetRiskReward Chart, dX1, dX2
        Case eANNOT_AndrewFork
            SetAndrewFork Chart, dX1, dX2, iStartX
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
             eANNOT_DNExpansion4, eANNOT_FibABCD
            SetDnExpansion Chart, dX1, dX2, iStartX
        Case eANNOT_Gartley
            SetGartley Chart, dX1, dX2, iStartX
        
    'items that need code for handling log draw mode
        Case eANNOT_TargetShooter
            SetTargetShooter Chart
            geSetTgtShooter Chart, dX1, dX2
    
    End Select
    
ErrExit:
    If m.eUsage = eANNOT_OptionInfo And dOptDateSave > 0 Then m.dDate2 = dOptDateSave
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geSetPoints", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Public Function geAddAnnotation(Chart As cChart, ByVal nPaneID&, ByVal nID&) As Long
On Error GoTo ErrSection:

    Dim rc&, i&, strText$
    Dim Ind As cIndicator
    Dim bConvertedText As Boolean
    
    If 1 = Val(Prop("IsHidden")) Then
        If AllowEWI Or AllowGMP Then
            Prop("IsHidden") = 0
        Else
            Exit Function
        End If
    End If
    
    m.geAnnStruct.paneId = nPaneID
    m.geAnnStruct.annotationId = nID
    Set m.Chart = Chart
    
    'double check text for indicator label (fix for Aardvark issue #87)
    If m.eUsage = eANNOT_IndicatorLabel Then
        If Chart.Tree.NodeLevel(m.geAnnStruct.indicatorId) > 0 Or _
            m.geAnnStruct.indicatorId = Chart.Tree.Count + 1 Then   'this is label for bad ticks indicator
            Set Ind = Chart.Tree(m.geAnnStruct.indicatorId)
            If Not Ind Is Nothing Then
                If Ind.Display = False Then
                    m.strText = ""
                    gdSetStr m.geAnnStruct.gshText, 0, ""
                End If
            End If
        Else
            Exit Function   'invalid indicator ID, don't add
        End If
    End If
    
    If m.eType = eANNOT_ElliotLabel And Len(m.strText) > 6 Then
        m.strText = Replace(m.strText, "~", vbCrLf)
        SetTextEditSize Chart, m.strText
        
        gdSetNum m.geAnnStruct.gdhMisc, 0, 1    'flag telling grapheng.dll to draw text based on width, height
        gdSetNum m.geAnnStruct.gdhMisc, 1, 0
    End If
        
    If m.geAdded = True Then
        geSetPoints Chart
        Exit Function
    End If

    'convert previously saved icontext annotations that contain only text
    If m.eType = eANNOT_Icon And m.nStyle = 0 And m.eUsage = eANNOT_UserAdded Then
        m.eType = eANNOT_TextEdit
        LoadDefault "Shape", 0, True
        m.geTextAlign = LoadDefault("TextAlignment", 8, True)   'e_ctrCtr
        LoadDefault "Width", 20, True
        LoadDefault "Height", 10, True
        LoadDefault "Border", 1, True       '1=border, 0=no border
        LoadDefault "ArrowStyle", 0, True   '0=no arrow, 1=standard, 2=triangle
        LoadDefault "ArrowSize", 0, True    '0=small, 1=medium, 2=large
        LoadDefault "ArrowLine", eANNOT_Thin, True
        LoadDefault "Anchor", 4, True
        bConvertedText = True
    End If
    
    'add alert to global collection if necessary
    If Len(m.strAlert) > 0 And Not m.bAlertAdded Then
        Set m.oAlert = New cAlert
        m.oAlert.FromFileString m.strAlert
        If m.oAlert.ChartAlertId = 0 Then m.oAlert.ChartAlertId = -1
        AlertsCollection 1
    End If
    
    geTransAnnType Chart
    geSetPoints Chart

    rc = geAddItem(Chart.geChartObj(), 4, m.geAnnStruct)
    If rc = 0 Then
        m.geAdded = True
        m.geChartObj = Chart.geChartObj()
    End If
    
    If bConvertedText = True Then SetTextEditSize Chart, m.strText
    
    geAddAnnotation = rc
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.geAddAnnotation", eGDRaiseError_Raise

End Function

Public Function geRemoveAnnotation(ByVal chartObj As Long) As Long
On Error GoTo ErrSection:

    Dim rc&
   
    If m.geAdded = False Then Exit Function
    
    rc = geRemoveItem(chartObj, 4, m.geAnnStruct)
    If rc = 0 Then m.geAdded = False
    If Not m.oAlert Is Nothing Then
        AlertsCollection 0          'remove from alerts collection
    End If
    
    geRemoveAnnotation = rc
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.geRemoveAnnotation", eGDRaiseError_Raise

End Function

Public Sub SetMoveOffSet(ByVal dMouseNx#, ByVal dMouseDy#, ByVal nShift&)
On Error GoTo ErrSection:

    m.dMoveOffSetX = dMouseNx - m.X1
    m.dMoveOffSetY = dMouseDy - m.Y1
    If m.bUsingLogDrawingMode Then
        m.dMoveOffSetLogY = LogModeDelta(m.Y1, dMouseDy)    'Log(dMouseDy) - Log(m.Y1)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetMoveOffSet", eGDRaiseError_Raise

End Sub

Public Sub GannRiseChange(ByVal dPoints#)
On Error GoTo ErrSection:

    Dim dRise#

    dRise = RoundNum(Val(Prop("Rise")), 3)
    If dRise = dPoints Then Exit Sub
    
    If m.Y2 > m.Y1 Then
        m.Y2 = m.Y1 + dPoints
    Else
        m.Y2 = m.Y1 - dPoints
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.GannRiseChange", eGDRaiseError_Raise

End Sub

Public Sub GannRunChange(Chart As cChart, ByVal nRun&)
On Error GoTo ErrSection:
    
    Dim i&, j&, hArray&
    Dim dDateTemp#
    
    If Int(Val(Prop("Run"))) = nRun Then Exit Sub

    'JM 10-24-2013 - fix for issue of Gann lines not accepting change to Run value
    '   in editor when fan is drawn snapped to high or low of non-intraday bars
    dDateTemp = m.dDate1
    If Not Chart.Bars Is Nothing Then
        If Not Chart.Bars.IsIntraday Then
            dDateTemp = Int(m.dDate1)
        End If
    End If
    
    hArray = Chart.geDateArray
    gdBinarySearch hArray, dDateTemp, j, eGdSort_Default, 0, -1
    
    If dDateTemp = gdGetNum(hArray, j) Then
        If m.X2 > m.X1 Then
            j = j + nRun
        Else
            j = j - nRun
        End If
        dDateTemp = gdGetNum(hArray, j)
        If dDateTemp > 0 Then
            If m.X2 > m.X1 Then
                m.X2 = m.X1 + nRun
            Else
                m.X2 = m.X1 - nRun
            End If
            m.dDate2 = dDateTemp
            Prop("Run") = nRun
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.GannRunChange", eGDRaiseError_Raise

End Sub

Public Sub FixTextEditYs()
On Error GoTo ErrSection:

    If m.X1 <> m.X2 Or m.Y1 <> m.Y2 Then Exit Sub
    
    If m.eType <> eANNOT_TextEdit And m.eType <> eANNOT_TextEdit2 And _
       m.eType <> eANNOT_TextEdit3 And m.eType <> eANNOT_TextEdit4 Then
       
       Exit Sub
    End If
    
    m.Y2 = gdGetNum(m.geAnnStruct.gdhBottom, 1)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.FixTextEditYs", eGDRaiseError_Raise

End Sub

Public Property Get TimeCycleSpace(Optional Chart As cChart = 0) As Long
  
    TimeCycleSpace = Abs(m.X2 - m.X1)
    
End Property

Public Property Let TimeCycleSpace(Chart As cChart, ByVal nSpacing&)
    
    m.X2 = m.X1 + nSpacing
    m.dDate2 = Chart.aXdate(m.X2)
    
End Property

Public Property Get geImageSize(Optional ByVal bForDisplay As Boolean = False) As Long
    
    Dim nSize&
    
    nSize = Val(Prop("ImageSize"))
    If Not bForDisplay Then
        geImageSize = nSize
        Exit Property
    End If
    
    'this is just for convenience when setting the combo box in editannot form
    If nSize <= 8 Then
        nSize = 0       'small
    ElseIf nSize <= 16 Then
        nSize = 1       'medium
    Else
        nSize = 2
    End If
    
    geImageSize = nSize
End Property

Public Property Get geImageType(Optional ByVal bForDisplay As Boolean) As Long
On Error GoTo ErrSection:
    
    Dim i&, nSize&, nIdx&, nDisplayIdx&, nFill&
    Dim eImage As eStockImage
    Dim eDir As eImageDir
    Dim bFound As Boolean
   
    eImage = Val(Prop("ImageType"))
    If bForDisplay = False Then
        geImageType = eImage
        Exit Property
    End If
    
    'Note: this code converts the enumerated image type to an index value
    'matching the picture box array in the gdSelectIcon control (gdOcx project)
    'TODO: look into combining this with the routine in cIndicator
    
    eDir = Val(Prop("ImageDir"))
    nFill = Val(Prop("ImageStyle"))
    Select Case eImage
        Case eCNI_Arrow
            Select Case eDir
                Case eCNI_North
                    nDisplayIdx = 0
                Case eCNI_South
                    nDisplayIdx = 1
                Case eCNI_East
                    nDisplayIdx = 2
                Case eCNI_West
                    nDisplayIdx = 3
                Case eCNI_NorthEast
                    nDisplayIdx = 4
                Case eCNI_SouthWest
                    nDisplayIdx = 5
                Case eCNI_SouthEast
                    nDisplayIdx = 6
                Case eCNI_NorthWest
                    nDisplayIdx = 7
            End Select
        Case eCNI_Plus
            nDisplayIdx = 8
        Case eCNI_Cross
            nDisplayIdx = 9
        Case eCNI_Circle
            If nFill = 0 Then
                nDisplayIdx = 17
            Else
                nDisplayIdx = 12
            End If
        Case eCNI_Square
            If nFill = 0 Then
                nDisplayIdx = 18
            Else
                nDisplayIdx = 13
            End If
        Case eCNI_Diamond
            If nFill = 0 Then
                nDisplayIdx = 19
            Else
                nDisplayIdx = 14
            End If
        Case eCNI_Triangle
            If nFill = 0 Then
                If eDir = eCNI_North Then
                    nDisplayIdx = 15
                Else
                    nDisplayIdx = 16
                End If
            Else
                If eDir = eCNI_North Then
                    nDisplayIdx = 10
                Else
                    nDisplayIdx = 11
                End If
            End If
    End Select
    
    geImageType = nDisplayIdx

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.geImageType (Get)", eGDRaiseError_Raise

End Property

Private Sub ConvertIconDefaults()
On Error GoTo ErrSection:
    
    Dim nDefaultStyle&

    nDefaultStyle = Val(LoadDefault("Style", -1))
    If nDefaultStyle = -1 Or nDefaultStyle <> m.nStyle Then Exit Sub
    
    m.strText = LoadDefault("Text", "")
    m.geTextAlign = LoadDefault("TextAlignment", 8, True) 'center
    m.bMultiChart = LoadDefault("MultiChart", True)
    m.nStyle = -1
    
    SaveDefaults
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ConvertIconDefaults", eGDRaiseError_Raise

End Sub

Public Sub HideGannLine(ByVal nIdx&)
On Error GoTo ErrSection:
    
    Dim strLineToHide$
    Dim nRise&, nRun&
    
    If nIdx < 1 Or nIdx > 9 Then Exit Sub
    
    'build string for setting property based on rise & run value
    nRise = gdGetNum(m.geAnnStruct.gdhBottom, nIdx)
    nRun = gdGetNum(m.geAnnStruct.gdhRight, nIdx)
    strLineToHide = Str(nRise) + "x" + Str(nRun)
    
    Prop(strLineToHide) = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.HideGannLine", eGDRaiseError_Raise

End Sub

Public Sub HideTargetShooterLine(ByVal nIdx&)
On Error GoTo ErrSection:

    If nIdx < 6 Or nIdx > 17 Then Exit Sub

    Select Case nIdx
        Case 6, 9, 12, 15
            Prop("Pt1") = 0
        Case 7, 10, 13, 16
            Prop("Pt2") = 0
        Case 8, 11, 14, 17
            Prop("Pt3") = 0
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.HideTargetShooterLine", eGDRaiseError_Raise

End Sub

Public Sub HideFibLine(ByVal nIdx&)
On Error GoTo ErrSection:

    Dim i&, j&, dVal#

    If m.eType = eANNOT_ElliotTimeRatio Then
        If nIdx > 4 Then
            i = nIdx - 5
        Else
            If nIdx = 2 Then
                Prop("ShowRatioZero") = 0
            ElseIf nIdx = 3 Then
                Prop("ShowRatioOne") = 0
            ElseIf nIdx = 4 Then
                Prop("ShowRatioTwo") = 0
            End If
            Exit Sub
        End If
    ElseIf m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 _
           Or m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_FibABCD Then
        If nIdx = 0 Then
            Prop("ShowRatioZero") = 0
            Exit Sub
        ElseIf nIdx = 1 Then
            Prop("ShowRatioOne") = 0
            Exit Sub
        ElseIf nIdx = 2 Then
            Prop("ShowRatioTwo") = 0
            Exit Sub
        ElseIf nIdx > 2 Then
            i = nIdx - 3
        Else
            Exit Sub
        End If
    ElseIf m.eType = eANNOT_AndrewFork Then
        i = nIdx + 3
    ElseIf nIdx > 5 Then
        i = nIdx - 6
    ElseIf m.eType <> eANNOT_Fibonacci And m.eType <> eANNOT_Fibonacci2 And m.eType <> eANNOT_Fibonacci3 _
           And m.eType <> eANNOT_Fibonacci4 And m.eType <> eANNOT_FibExpansion And m.eType <> eANNOT_AdvRiskReward _
           And m.eType <> eANNOT_FibFan And m.eType <> eANNOT_FibArcs And m.eType <> eANNOT_FibTimeRatio Then
        Exit Sub
    End If
    
    If m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 Or m.eType = eANNOT_Fibonacci3 _
       Or m.eType = eANNOT_Fibonacci4 Or m.eType = eANNOT_FibExpansion Or m.eType = eANNOT_AdvRiskReward Then
        j = gdGetSize(m.geAnnStruct.glhAnnType)
        If nIdx = 1 Then
            Prop("HideVerticalLine") = 1
        ElseIf nIdx = 2 Then
            Prop("ShowRatioZero") = 0
        ElseIf nIdx = 3 Then
            Prop("ShowRatioOne") = 0
        ElseIf nIdx < j Then
            If gdGetNum(m.geAnnStruct.glhAnnType, nIdx) = 0 Then
                i = -1
                dVal = gdGetNum(m.geAnnStruct.gdhTop, nIdx)
                For j = 0 To m.aRatios.Size - 1
                    If gdGetNum(m.geAnnStruct.gdhTop, (j + 6)) = dVal Then
                        i = j
                        Exit For
                    End If
                Next
            End If
        Else
            i = -1      'so nothing will get changed
        End If
    ElseIf m.eType = eANNOT_FibFan Then
        If nIdx = 0 Then
            Prop("ShowRatioZero") = 0
            Prop("showRatioOne") = 0
        End If
    ElseIf m.eType = eANNOT_FibArcs Then
        If nIdx = 5 Then
            Prop("ShowRatioZero") = 0
            Prop("showRatioOne") = 0
        End If
    ElseIf m.eType = eANNOT_FibTimeRatio Then
        j = gdGetSize(m.geAnnStruct.glhAnnType)
        If nIdx = 2 Then
            Prop("ShowRatioZero") = 0
        ElseIf nIdx = 3 Then
            Prop("ShowRatioOne") = 0
        ElseIf nIdx > 3 And nIdx < j Then
            If gdGetNum(m.geAnnStruct.glhAnnType, nIdx) = 0 Then
                i = -1
                dVal = gdGetNum(m.geAnnStruct.gdhLeft, nIdx)
                For j = 0 To m.aRatios.Size - 1
                    If gdGetNum(m.geAnnStruct.gdhLeft, (j + 6)) = dVal Then
                        i = j
                        Exit For
                    End If
                Next
            End If
        End If
    End If
    
    If i >= 0 And i < m.aRatioShow.Size Then
        m.aRatioShow(i) = 0
        Prop("ShowRatios") = m.aRatioShow.JoinFields(",")
        If Not m.Chart Is Nothing Then m.Chart.GenerateChart eRedo1_Scrolled
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.HideFibLine", eGDRaiseError_Raise

End Sub

Public Property Get ChartTip(Chart As cChart, ByVal nSubItem&, ByVal dMouseDate#) As String

    Dim hString As Long
    Dim strTip$, dX#, i&
        
    strTip = ""
    
    Select Case m.eType
        Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4
            If nSubItem > 0 Then
                If Not m.oAlert Is Nothing Then strTip = m.oAlert.ChartCondition
            Else
                strTip = Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
            End If


        Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, eANNOT_Fibonacci4, _
             eANNOT_FibExpansion, eANNOT_DanCodeFib, eANNOT_AdvRiskReward
        
            hString = gdCreateArray(eGDARRAY_gdString)
            geAnnotAtVal Chart.geChartObj, 1, m.geAnnStruct.paneId, 15, 0, gdGetNum(m.geAnnStruct.gdhTop, nSubItem), hString
            strTip = Trim(gdGetStr(hString))
            strTip = Replace(strTip, "^L", "")
            strTip = Replace(strTip, "^Y", "")
            strTip = Replace(strTip, "^R", "")
            strTip = Replace(strTip, "|", Space(2))
            gdDestroyArray hString
        
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
             eANNOT_DNExpansion4, eANNOT_FibABCD
            
            hString = gdCreateArray(eGDARRAY_gdString)
            dX = GetX(Chart, dMouseDate) - Chart.ScreenStartX
            geAnnotAtVal Chart.geChartObj, 14, m.geAnnStruct.paneId, 15, dX, gdGetNum(m.geAnnStruct.gdhTop, nSubItem), hString
            strTip = Trim(gdGetStr(hString))
            strTip = Replace(strTip, "|", Space(2))
            gdDestroyArray hString

        Case eANNOT_ElliotTimeRatio, eANNOT_FibTimeRatio
            hString = gdCreateArray(eGDARRAY_gdString)
            geAnnotAtVal Chart.geChartObj, 26, m.geAnnStruct.paneId, 15, gdGetNum(m.geAnnStruct.gdhLeft, nSubItem), 0, hString
            strTip = Trim(gdGetStr(hString))
            strTip = Replace(strTip, "|", Space(2))
            gdDestroyArray hString
        
        Case eANNOT_TargetShooter
            If Val(Prop("ShowValues")) = 0 Then
                If nSubItem > 5 Then
                    strTip = gdGetStr(m.geAnnStruct.gshText, nSubItem + 6)
                End If
            End If

        Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
             eANNOT_TrendChannel
            If nSubItem > 0 Then
                If Not m.oAlert Is Nothing Then
                    strTip = m.oAlert.ChartCondition
                End If
            End If
    End Select
    
    ChartTip = strTip

End Property

Public Function SaveIconToString(Optional ByVal strName$) As String
On Error GoTo ErrSection:

    Dim strCfg$, strMulti$, strAlign$, strFile$, s$
    Dim strImage$, strImageSize$, strImageDir$, strImageStyle$
    Dim strFont$, strFontSize$, strFontStyle$, strUnderline$
    Dim strLabelType$, strAsciiChar$
    Dim aFile As New cGdArray
    Dim aFields As New cGdArray
    Dim i&
    
    If m.eType <> eANNOT_Icon Then
        Exit Function
    End If
    
    strFile = g.strAppPath & "\custom\IconAnnot.Cfg"
    s = "This will save the settings for this icon." & vbCrLf & "Please enter a name for the saved settings:"
       
    If Len(strName) = 0 Then
        strCfg = InfBox(s, "?", , "Save Icon Configuration", , , , , , "string", strName, eGDAlign_Left)
        strCfg = Trim(strCfg)
    Else
        strCfg = strName
    End If
    SaveIconToString = strCfg
        
    If Len(strCfg) > 0 Then
        'save new icon configuration
        strImage = Str(Prop("ImageType"))
        strImageSize = Str(Prop("ImageSize"))
        strImageStyle = Str(Prop("ImageStyle"))
        strImageDir = Str(Prop("ImageDir"))
        strFont = Prop("FontName")
        strFontSize = Str(Prop("FontSize"))
        strFontStyle = Str(Prop("FontStyle"))
        strUnderline = Str(Prop("FontUnderline"))
        strAlign = Str(Prop("TextAlignment"))
        strMulti = Str(m.bMultiChart)       '5514
        strAsciiChar = Prop("AsciiChar")
        
        'backwards compatibility:
        '   original implementation had label type in field 15, "!" is now place holder
        '   fields: name, color, style, preindicator flag, multichart flag,
        '           align, text, image type, image size, image style (hollow/solid),
        '           image direction, font, font size, font style, underline,
        '           label type (obsolete: using ! as place holder), ascii char
        strCfg = strCfg & vbTab & m.nColor & vbTab & "-1" & vbTab & m.nPreIndicator & vbTab & _
            strMulti & vbTab & strAlign & vbTab & m.strText & vbTab & strImage & vbTab & _
            strImageSize & vbTab & strImageStyle & vbTab & strImageDir & vbTab & strFont & vbTab & _
            strFontSize & vbTab & strFontStyle & vbTab & strUnderline & vbTab & "!" & vbTab & _
            strAsciiChar
        
        aFile.FromFile strFile
        If Len(strName) = 0 Then
            aFile.Add strCfg            'add to end
        Else
            'replace with new settings
            For i = aFile.Size - 1 To 0 Step -1
                aFields.SplitFields aFile(i), vbTab
                If aFields(0) = strName Then
                    aFile.Remove i
                    aFile.Add strCfg, i
                    Exit For
                End If
            Next
        End If
        aFile.ToFile strFile
    End If
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SaveIconToString", eGDRaiseError_Raise

End Function

Public Sub SetIconFromString(aIconList As cGdArray, ByVal nIdx&)
On Error GoTo ErrSection:

    Dim aFields As New cGdArray
    Dim i&
    
    aFields.SplitFields aIconList(nIdx), vbTab
    
    m.nColor = Val(aFields(1))
    m.nStyle = Val(aFields(2))
    m.nPreIndicator = Val(aFields(3))
    m.geAnnStruct.PreIndicator = m.nPreIndicator
    m.bMultiChart = Val(aFields(4))
    
    i = Val(aFields(5))
    If i = 3 Or i = 4 Or i = 5 Then
        m.geTextAlign = i
    Else
        m.geTextAlign = 5
    End If
    Prop("TextAlignment") = m.geTextAlign
    
    m.strText = aFields(6)
    
    Prop("ImageType") = Val(aFields(7))
    Prop("ImageSize") = Val(aFields(8))
    Prop("ImageStyle") = Val(aFields(9))
    Prop("ImageDir") = Val(aFields(10))
    
    Prop("FontName") = Val(aFields(11))
    Prop("FontSize") = Val(aFields(12))
    Prop("FontStyle") = Val(aFields(13))
    Prop("FontUnderline") = Val(aFields(14))
       
    If aFields.Size >= 17 Then
        'backwards compatibility:
        '   original implementation had label type in field 15, now ignored
        Prop("AsciiChar") = aFields(16)
    End If
    
    geTransIconTextStyle

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetIconFromString", eGDRaiseError_Raise

End Sub

Private Function MoveSRLine(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#
    
    If nPoint = 2 Then
        m.X2 = RoundNum(X)          'right end point
        'Prop("Ext") = "0"          '6789
        bMoved = True
    ElseIf nPoint = 1 Then
        m.X1 = RoundNum(X)          'left end point (4264)
        bMoved = True
    Else
        XDelta = m.X2 - m.X1        'middle (move whole line)
        m.X1 = X - m.dMoveOffSetX
        m.X2 = m.X1 + XDelta
        m.Y1 = RoundedY(Chart, Y)
        m.Y2 = m.Y1
        bMoved = True
    End If

    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size - 1 Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size - 1 Then bMoved = False
    End If
    
    MoveSRLine = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveSRLine", eGDRaiseError_Raise

End Function

Private Function MovePivot(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#, nBar&
    
    nBar = m.Chart.LastGoodDataBar(True)        '4329
    
    If X > nBar Then
        Exit Function
    End If
    
    If nPoint = 2 Then
        m.X2 = RoundNum(X)
        bMoved = True
    ElseIf nPoint = 1 Then
        m.X1 = RoundNum(X)
        bMoved = True
    Else
        XDelta = m.X2 - m.X1
        m.X1 = X - m.dMoveOffSetX       'move whole line
        m.X2 = m.X1 + XDelta
        bMoved = True
    End If

    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
    End If
    
    MovePivot = bMoved
    CalcPivots True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MovePivot", eGDRaiseError_Raise

End Function

Private Function MoveBalloonStrangle(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#, X1#
    
    If m.nHitItemIdx = gdGetSize(m.geAnnStruct.glhAnnType) - 1 Then
        If X >= 0 And X < Chart.aXdate.Size Then
            m.aDates(1) = Chart.aXdate(X)
            m.aYs(3) = Y
            bMoved = True
        End If
    ElseIf nPoint = 2 Then
        m.X2 = RoundNum(X)          'right end point
        bMoved = True
    ElseIf nPoint = 1 Then
        m.X1 = RoundNum(X)          'left end point (4264)
        bMoved = True
    Else
        XDelta = m.X2 - m.X1        'middle (move whole line)
        X1 = GetX(Chart, m.aDates(1)) - m.X1
        
        m.X1 = X - m.dMoveOffSetX
        m.X2 = m.X1 + XDelta
        
        XDelta = X1
        X1 = m.X1 + XDelta

        If X1 >= 0 And X1 < Chart.aXdate.Size Then
            m.aDates(1) = Chart.aXdate(X1)
        End If
        
        bMoved = True
    End If

    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size - 1 Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size - 1 Then bMoved = False
    End If
    
    MoveBalloonStrangle = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveSRLine", eGDRaiseError_Raise

End Function

Private Function SetSRLine(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim strValue$, dxDot#, i&

    If dX1 = 0 And dX2 = 0 Then Exit Function       'aardvark 1806

    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 1
        'for text
        gdSetStr m.geAnnStruct.gshText, 0, " "
        gdSetNum m.geAnnStruct.glhAlign, 0, 0
        gdSetStr m.geAnnStruct.gshFont, 0, "Arial"
        gdSetNum m.geAnnStruct.glhSize, 0, 8
        
        'place holder
        gdSetNum m.geAnnStruct.glhStyle, i, 0       'hollow circle
        gdSetNum m.geAnnStruct.glhFillPattern, i, 0
        gdSetNum m.geAnnStruct.glhDirection, i, 6   'e_west, shift left
        'dot
        gdSetNum m.geAnnStruct.glhAnnType, 1, 6     'type ellipse
        gdSetNum m.geAnnStruct.glhStyle, 1, 1       'hollow circle
        gdSetNum m.geAnnStruct.glhFillPattern, 1, 1
        gdSetNum m.geAnnStruct.glhDirection, 1, 6   'e_west, shift left
    End If
    
    'line
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y1
    If dX2 < dX1 Then
        If dX2 < 0 Then          'out of range (user changed to lower bar period - issue 3218)
            If m.dDate2 > m.dDate1 Then
                gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
                gdSetNum m.geAnnStruct.gdhRight, 0, Chart.geChartPoints - 1
            Else
                gdSetNum m.geAnnStruct.gdhLeft, 0, 0
                gdSetNum m.geAnnStruct.gdhRight, 0, dX1
            End If
        Else
            gdSetNum m.geAnnStruct.gdhLeft, 0, dX2
            gdSetNum m.geAnnStruct.gdhRight, 0, dX1
        End If
    Else
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    End If
        
    If Prop("Ext") = "1" Then
        gdSetNum m.geAnnStruct.gdhMisc, 0, 1
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        If dX1 > dX2 Then gdSetNum m.geAnnStruct.gdhRight, 0, dX1 + 1   '4264
        dxDot = dX1
    Else
        gdSetNum m.geAnnStruct.gdhMisc, 0, 0
        'dot
        If dX1 < dX2 Then
            dxDot = dX1
        ElseIf dX2 > 0 Then
            dxDot = dX2
        Else
            dxDot = dX1
        End If
    End If
        
    'hide dot   - 6990
    If Val(Prop("HideSRLineDot")) = 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 1, 6
    End If
    
    'drawn with single-click ---> dx1 = dx2
    If dX1 = dX2 Then
        dX2 = dX1 + 1
        gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    End If
    
    'show value in y-axis
    m.geAnnStruct.showQtrLines = Val(Prop("ShowInAxis"))      '6707
    
    'text
    i = Val(Prop("TextAlignment"))
    If i <> 6 And i <> 7 Then
        i = 7
        Prop("TextAlignment") = 7
    End If
    If Prop("ShowValues") = "1" Then
        gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")             '5749
        gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
        gdSetNum m.geAnnStruct.glhAlign, 0, i
        strValue = Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
        gdSetStr m.geAnnStruct.gshText, 0, strValue
    Else
        gdSetStr m.geAnnStruct.gshText, 0, ""
    End If
                
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhBottom, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 1, dxDot
    gdSetNum m.geAnnStruct.gdhRight, 1, dxDot
        
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
    geSetPenStyle m.nStyle, 0
    
    If m.oAlert Is Nothing Then
        'don't draw bell
        If gdGetSize(m.geAnnStruct.glhAnnType) > 3 Then gdSetNum m.geAnnStruct.glhAnnType, 3, -1
    Else
        SetAlertIcon 3, m.Y1
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetSRLine", eGDRaiseError_Raise

End Function

Private Function SetBalloonStrangle(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection

    Dim i&, j&, strValue$, dxRectLeft#, dxRectRight#
    Dim showPrices&, showCost&, showPL&, showPM&, align&
    Dim riskPenstyle&, riskPenColor&, fSize&, fStyle&, fName$
    
    Dim dAmt#, strAmt$, strPriceMove$

    If dX1 = 0 And dX2 = 0 Then Exit Function       'aardvark 1806

    showPrices = Val(Prop("ShowPrices"))
    showCost = Val(Prop("ShowTradeCost"))
    showPL = Val(Prop("ShowProfitLoss"))
    showPM = Val(Prop("ShowPriceMove"))
    align = Val(Prop("TextAlignment"))
    
    fSize = Val(Prop("FontSize"))
    fStyle = Val(Prop("FontStyle"))
    fName = Prop("FontName")
    riskPenstyle = Val(Prop("RiskBEStyle"))
    riskPenColor = Val(Prop("RiskBEColor"))
    
    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 1
    End If
    
    If dX2 < dX1 Then
        If dX2 < 0 Then          'out of range (user changed to lower bar period - issue 3218)
            If m.dDate2 > m.dDate1 Then
                gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
                gdSetNum m.geAnnStruct.gdhRight, 0, Chart.geChartPoints - 1
            Else
                gdSetNum m.geAnnStruct.gdhLeft, 0, 0
                gdSetNum m.geAnnStruct.gdhRight, 0, dX1
            End If
        Else
            gdSetNum m.geAnnStruct.gdhLeft, 0, dX2
            gdSetNum m.geAnnStruct.gdhRight, 0, dX1
        End If
    Else
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    End If
    
    dX1 = gdGetNum(m.geAnnStruct.gdhLeft, 0)
    dX2 = gdGetNum(m.geAnnStruct.gdhRight, 0)
    
    If dX2 - dX1 < 4 Then dX2 = dX1 + 4
    
    If Val(Prop("ShowCollapsed")) = 1 Then          '7019
        gdSetSize m.geAnnStruct.glhAnnType, 0, 0
        gdSetNum m.geAnnStruct.glhAnnType, 0, 1
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        gdSetNum m.geAnnStruct.gdhRight, 0, dX2
        gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
        gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y1
        geSetPenStyle riskPenstyle, 0
        gdSetNum m.geAnnStruct.glhPenColor, 0, riskPenColor
        
        gdSetStr m.geAnnStruct.gshFont, 0, fName
        gdSetNum m.geAnnStruct.glhSize, 0, fSize
        gdSetNum m.geAnnStruct.glhStyle, 0, fStyle
        gdSetNum m.geAnnStruct.glhUnderline, 0, 0
        gdSetNum m.geAnnStruct.glhAlign, 0, align
        gdSetNum m.geAnnStruct.gdhMisc, 0, 0
        gdSetNum m.geAnnStruct.glhFillPattern, 0, 0
        
        strValue = ""
        If showPrices = 1 Then strValue = Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
        If showCost Then
            strValue = strValue & " (" & "cost = $" & Chart.Bars.PriceDisplay(BalloonTradeCost) & ")"
        End If
        gdSetStr m.geAnnStruct.gshText, 0, strValue
        
        GoTo ErrExit
    End If
    
    dxRectLeft = dX1 + (dX2 - dX1) / 4
    dxRectRight = dX2 - 2
    
    'draw order for coloring purposes: boxes, ratio lines (descending order), price lines
    gdSetSize m.geAnnStruct.gshFont, 0, 0
    gdSetSize m.geAnnStruct.glhSize, 0, 0
    gdSetSize m.geAnnStruct.glhStyle, 0, 0
    gdSetSize m.geAnnStruct.glhUnderline, 0, 0
    gdSetSize m.geAnnStruct.glhAlign, 0, 0
    gdSetSize m.geAnnStruct.gdhMisc, 0, 0
    gdSetSize m.geAnnStruct.glhFillPattern, 0, 0
    gdSetSize m.geAnnStruct.gshText, 0, 0
    
    'set annot types
    If BalloonTradeCost > 0 Then
        gdSetSize m.geAnnStruct.glhAnnType, 0, 0
        
        'dA = strike price above stock price
        'dB = strike price below stock price
        'dR = ratio value in ratios array (e.g. 0.5, 1, 1.5 ...)
        Dim dxA#, dxB#, dR#
        
        If BalloonPutStrike > BalloonStockPrice Then
            dxA = BalloonPutStrike
            dxB = BalloonCallStrike
        Else
            dxA = BalloonCallStrike
            dxB = BalloonPutStrike
        End If
        
        j = 0
        For i = m.aRatioShow.Size - 1 To 0 Step -1
            If m.aRatioShow(i) = 1 Then
                dR = m.aRatios(i)
                If Int(dR) = dR Then
                    strValue = Str(Abs(dR)) & "R"
                Else
                    strValue = Format(Abs(dR), "0.00") & "R"
                End If
                
                gdSetNum m.geAnnStruct.glhAnnType, j, 4
                gdSetNum m.geAnnStruct.glhAnnType, j + 1, 1
                
                If dR > 0 Then
                    dR = dxA + BalloonTradeCost + BalloonTradeCost * dR
                    gdSetNum m.geAnnStruct.gdhTop, j + 1, dR
                    gdSetNum m.geAnnStruct.gdhBottom, j + 1, dR
                    gdSetNum m.geAnnStruct.gdhTop, j, dR
                    gdSetNum m.geAnnStruct.gdhBottom, j, dxA + BalloonTradeCost
                Else
                    dR = dxB - BalloonTradeCost + BalloonTradeCost * dR
                    gdSetNum m.geAnnStruct.gdhTop, j + 1, dR
                    gdSetNum m.geAnnStruct.gdhBottom, j + 1, dR
                    gdSetNum m.geAnnStruct.gdhTop, j, dR
                    gdSetNum m.geAnnStruct.gdhBottom, j, dxB - BalloonTradeCost
                End If
                gdSetNum m.geAnnStruct.gdhLeft, j, dxRectLeft
                gdSetNum m.geAnnStruct.gdhRight, j, dxRectRight
                
                If align = 7 Then
                    gdSetNum m.geAnnStruct.gdhLeft, j + 1, dX1
                    gdSetNum m.geAnnStruct.gdhRight, j + 1, dxRectRight - 1
                Else
                    gdSetNum m.geAnnStruct.gdhLeft, j + 1, dxRectLeft
                    gdSetNum m.geAnnStruct.gdhRight, j + 1, dX2
                End If
                
                If showPL Then
                    If m.aRatios(i) < 0 Then
                        'negative ratios
                        dAmt = GetProfitLoss(Chart, BalloonPutStrike, dR, strPriceMove)
                    Else
                        dAmt = GetProfitLoss(Chart, BalloonCallStrike, dR, strPriceMove)
                    End If
                    dAmt = Abs(dAmt) - BalloonTradeCost
                    
                    If strPriceMove <> "INVALID" Then
                        If dAmt = Int(dAmt) Then
                            strAmt = Format(dAmt, "$#,##0;($#,##0)")
                        Else
                            strAmt = Format(dAmt, "$#,##0.00;$#,##0.00")
                        End If
                        strValue = strValue & ": " & strAmt
                    End If
                End If
                
                If showPM Then
                    dAmt = GetProfitLoss(Chart, BalloonStockPrice, dR, strPriceMove)

                    If strPriceMove <> "INVALID" Then
                        If dAmt = Int(dAmt) Then
                            strAmt = Format(dAmt, "#,##0;(#,##0)")
                        Else
                            strAmt = Format(dAmt, "#,##0.00;#,##0.00")
                        End If
                        If showPL Then
                            strValue = strValue & " / " & strAmt
                        Else
                            strValue = strValue & ": " & strAmt
                        End If
                    End If
                End If
                gdSetStr m.geAnnStruct.gshText, j + 1, strValue
                
                gdSetNum m.geAnnStruct.glhPenColor, j, m.aRatioColor(i)
                gdSetNum m.geAnnStruct.glhPenColor, j + 1, m.aRatioColor(i)
'                gdSetNum m.geAnnStruct.glhPenColor, j, m.aRatioColor(i)
'                gdSetNum m.geAnnStruct.glhPenColor, j + 1, m.aRatioColor(i)
                
                j = j + 2
            End If
        Next

        'BE above/below
        If Val(Prop("ShowBE")) = 1 Then
            'BE above
            If BalloonCallCost > 0 Then
                gdSetNum m.geAnnStruct.glhAnnType, j, 4
                gdSetNum m.geAnnStruct.glhAnnType, j + 1, 1
            Else
                gdSetNum m.geAnnStruct.glhAnnType, j, -1
                gdSetNum m.geAnnStruct.glhAnnType, j + 1, -1
            End If
            gdSetNum m.geAnnStruct.gdhTop, j, dxA + BalloonTradeCost
            gdSetNum m.geAnnStruct.gdhBottom, j, dxA
            gdSetNum m.geAnnStruct.gdhLeft, j, dxRectLeft
            gdSetNum m.geAnnStruct.gdhRight, j, dxRectRight
            
            gdSetNum m.geAnnStruct.gdhTop, j + 1, dxA + BalloonTradeCost
            gdSetNum m.geAnnStruct.gdhBottom, j + 1, dxA + BalloonTradeCost
            
            If BalloonPutCost > 0 Then
                gdSetNum m.geAnnStruct.glhAnnType, j + 2, 4
                gdSetNum m.geAnnStruct.glhAnnType, j + 3, 1
            Else
                gdSetNum m.geAnnStruct.glhAnnType, j + 2, -1
                gdSetNum m.geAnnStruct.glhAnnType, j + 3, -1
            End If
            gdSetNum m.geAnnStruct.gdhTop, j + 2, dxB - BalloonTradeCost
            gdSetNum m.geAnnStruct.gdhBottom, j + 2, dxB
            gdSetNum m.geAnnStruct.gdhLeft, j + 2, dxRectLeft
            gdSetNum m.geAnnStruct.gdhRight, j + 2, dxRectRight
            
            gdSetNum m.geAnnStruct.gdhTop, j + 3, dxB - BalloonTradeCost
            gdSetNum m.geAnnStruct.gdhBottom, j + 3, dxB - BalloonTradeCost
        
            strValue = "B/E"
            If showPL Then strValue = strValue & ": $0"
            If showPM Then
                dAmt = GetProfitLoss(Chart, BalloonStockPrice, dxA + BalloonTradeCost, strPriceMove)
                If strPriceMove <> "INVALID" Then
                    If dAmt = Int(dAmt) Then
                        strAmt = Format(dAmt, "#,##0;(#,##0)")
                    Else
                        strAmt = Format(dAmt, "#,##0.00;#,##0.00")
                    End If
                    If showPL Then
                        strValue = strValue & " / " & strAmt
                    Else
                        strValue = strValue & ": " & strAmt
                    End If
                End If
            End If
            gdSetStr m.geAnnStruct.gshText, j + 1, strValue
            
            strValue = "B/E"
            If showPL Then strValue = strValue & ": $0"
            If showPM Then
                dAmt = GetProfitLoss(Chart, BalloonStockPrice, dxB - BalloonTradeCost, strPriceMove)
                If strPriceMove <> "INVALID" Then
                    If dAmt = Int(dAmt) Then
                        strAmt = Format(dAmt, "#,##0;(#,##0)")
                    Else
                        strAmt = Format(dAmt, "#,##0.00;#,##0.00")
                    End If
                    If showPL Then
                        strValue = strValue & " / " & strAmt
                    Else
                        strValue = strValue & ": " & strAmt
                    End If
                End If
            End If
            gdSetStr m.geAnnStruct.gshText, j + 3, strValue
            
            If align = 7 Then
                gdSetNum m.geAnnStruct.gdhLeft, j + 1, dX1
                gdSetNum m.geAnnStruct.gdhRight, j + 1, dxRectRight - 1
                gdSetNum m.geAnnStruct.gdhLeft, j + 3, dX1
                gdSetNum m.geAnnStruct.gdhRight, j + 3, dxRectRight - 1
            Else
                gdSetNum m.geAnnStruct.gdhLeft, j + 1, dxRectLeft
                gdSetNum m.geAnnStruct.gdhRight, j + 1, dX2
                gdSetNum m.geAnnStruct.gdhLeft, j + 3, dxRectLeft
                gdSetNum m.geAnnStruct.gdhRight, j + 3, dX2
            End If
            
            For i = 0 To 3
                gdSetNum m.geAnnStruct.glhPenColor, j + i, riskPenColor
            Next
        
        End If
        
        j = gdGetSize(m.geAnnStruct.glhAnnType)
        For i = 0 To j - 1
            geSetPenStyle riskPenstyle, i
        Next
    
    Else
        j = 0
    End If

    If dX1 = kNullData Then dX1 = 0
    'underlying stock price
    gdSetNum m.geAnnStruct.glhAnnType, j, 1
    gdSetNum m.geAnnStruct.glhAnnType, j + 1, 1
    gdSetNum m.geAnnStruct.glhAnnType, j + 2, 1

    gdSetNum m.geAnnStruct.gdhLeft, j, dX1
    gdSetNum m.geAnnStruct.gdhRight, j, dX2
    gdSetNum m.geAnnStruct.gdhTop, j, m.Y1
    gdSetNum m.geAnnStruct.gdhBottom, j, m.Y1
    geSetPenStyle riskPenstyle, j
    gdSetNum m.geAnnStruct.glhPenColor, j, riskPenColor
    
    strValue = ""
    If showPrices = 1 Then strValue = Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
    If showCost Then
        strValue = strValue & " (" & "cost = $" & Chart.Bars.PriceDisplay(BalloonTradeCost) & ")"
    End If
    gdSetStr m.geAnnStruct.gshText, j, strValue

    'call/strike prices
    If BalloonPutStrike > 0# Then
        gdSetNum m.geAnnStruct.glhPenColor, j + 2, m.nColor
        gdSetNum m.geAnnStruct.gdhLeft, j + 1, dX1
        gdSetNum m.geAnnStruct.gdhRight, j + 1, dX2
        gdSetNum m.geAnnStruct.gdhTop, j + 1, BalloonPutStrike
        gdSetNum m.geAnnStruct.gdhBottom, j + 1, BalloonPutStrike
        geSetPenStyle m.nStyle, j + 1
        gdSetNum m.geAnnStruct.glhPenColor, j + 1, m.nColor
        If showPrices = 1 Then
            strValue = Chart.PriceDisplay(m.geAnnStruct.paneId, BalloonPutStrike)
            gdSetStr m.geAnnStruct.gshText, j + 1, strValue
        End If
    End If

    If BalloonCallStrike > 0# Then
        gdSetNum m.geAnnStruct.gdhLeft, j + 2, dX1
        gdSetNum m.geAnnStruct.gdhRight, j + 2, dX2
        gdSetNum m.geAnnStruct.gdhTop, j + 2, BalloonCallStrike
        gdSetNum m.geAnnStruct.gdhBottom, j + 2, BalloonCallStrike
        geSetPenStyle m.nStyle, j + 2
        If showPrices = 1 Then
            strValue = Chart.PriceDisplay(m.geAnnStruct.paneId, BalloonCallStrike)
            gdSetStr m.geAnnStruct.gshText, j + 2, strValue
        End If
    End If

    For i = 0 To gdGetSize(m.geAnnStruct.glhAnnType) - 1
        gdSetStr m.geAnnStruct.gshFont, i, fName
        gdSetNum m.geAnnStruct.glhSize, i, fSize
        gdSetNum m.geAnnStruct.glhStyle, i, fStyle
        gdSetNum m.geAnnStruct.glhUnderline, i, 0
        gdSetNum m.geAnnStruct.glhAlign, i, align
        gdSetNum m.geAnnStruct.gdhMisc, i, 0
        gdSetNum m.geAnnStruct.glhFillPattern, i, 0
    Next

    If gdGetSize(m.geAnnStruct.glhAnnType) < 4 Then Exit Function

    'aDates(0) = options expiration date
    'aDates(1) = text location
    '   Y1 = stock price
    '   Y2 = put strike
    '   aY(0) = call strike
    '   aY(1) = cost of put
    '   aY(2) = cost of call
    '   aY(3) = text location
    If Len(m.strText) > 0 And Val(Prop("ShowText")) = 1 Then
        SetTextEditSize Chart, m.strText
        
        i = gdGetSize(m.geAnnStruct.glhAnnType)

        gdSetNum m.geAnnStruct.glhAnnType, i, 0

        If m.aDates(1) = kNullData Then
            gdSetNum m.geAnnStruct.gdhLeft, i, dX2
        Else
            gdSetNum m.geAnnStruct.gdhLeft, i, GetX(Chart, m.aDates(1)) - Chart.ScreenStartX
        End If

        If m.aYs(3) = kNullData Or m.aYs(3) = 0 Then
            gdSetNum m.geAnnStruct.gdhTop, i, BalloonStockPrice
        Else
            gdSetNum m.geAnnStruct.gdhTop, i, m.aYs(3)
        End If
        gdSetNum m.geAnnStruct.gdhRight, i, Val(Prop("Width")) + 10
        gdSetNum m.geAnnStruct.gdhBottom, i, Val(Prop("Height")) + 3
        gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
        gdSetNum m.geAnnStruct.glhPenSize, i, 1
        gdSetNum m.geAnnStruct.glhPenStyle, i, 0

        gdSetStr m.geAnnStruct.gshFont, i, fName
        gdSetNum m.geAnnStruct.glhSize, i, fSize
        gdSetNum m.geAnnStruct.glhStyle, i, fStyle
        gdSetNum m.geAnnStruct.gdhMisc, i, 1
        gdSetNum m.geAnnStruct.glhImageType, i, 1
        gdSetNum m.geAnnStruct.glhFillPattern, i, 0

        gdSetStr m.geAnnStruct.gshText, i, m.strText
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetBalloonStrangle", eGDRaiseError_Raise

End Function

Public Sub SetSRLineExt(Chart As cChart, ByVal nExtension As Long)
On Error GoTo ErrSection:

    If nExtension = Val(Prop("Ext")) Then Exit Sub
    
'11-05-2007: while fixing 4325, noticed line was going through text when ext gets turned on
'            this code changed fixes line going through text problem
    If m.dDate1 > m.dDate2 Then m.dDate1 = m.dDate2
    If (nExtension = 0 And (m.dDate1 = m.dDate2)) Or nExtension = 1 Then
        m.dDate2 = Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False, False))
    End If
    Prop("Ext") = nExtension

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetSRLineExt", eGDRaiseError_Raise

End Sub

Private Sub SetMultiRects(Chart As cChart, ByVal nStartScreenX&, Optional tblRectsDim As cGdTable = Nothing)
On Error GoTo ErrSection:

    Dim i&, nSize&
    Dim dX1#, dX2#

    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 25        'sets annotation type
        gdSetNum m.geAnnStruct.glhFillPattern, 0, 1     '0=transparent, 1=color using fill color
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0        'default to solid
        gdSetNum m.geAnnStruct.glhPenSize, 0, 2         'default to size 1
        gdSetNum m.geAnnStruct.glhFillColor, 0, vbYellow
        m.nPreIndicator = 1
    End If
    
    'these are not used
    m.dDate1 = 0
    m.dDate2 = 0
    m.Y1 = 0
    m.Y2 = 0
    m.X1 = 0
    m.X2 = 0
    
    m.geAnnStruct.PreIndicator = 1
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    gdSetNum m.geAnnStruct.glhFillColor, 0, Val(Prop("FillColor"))
    gdSetNum m.geAnnStruct.glhFillPattern, 0, Val(Prop("FillPattern"))
    m.geAnnStruct.showAxes = Val(Prop("ShowInAllPanes"))
    
    If tblRectsDim Is Nothing Then Exit Sub
    
    'clear annotation type info so rects will not be drawn to handle
    'changes in the number of rects as data changes when scrolling
'    nSize = gdGetSize(m.geAnnStruct.gdhTop) - 1
'    For i = 1 To nSize
'        gdSetNum m.geAnnStruct.glhAnnType, i, -1
'    Next
    
    gdSetSize m.geAnnStruct.glhAnnType, 1, 0
    gdSetSize m.geAnnStruct.gdhTop, 1, 0
    gdSetSize m.geAnnStruct.gdhLeft, 1, 0
    gdSetSize m.geAnnStruct.gdhBottom, 1, 0
    gdSetSize m.geAnnStruct.gdhRight, 1, 0
    
    Dim dTop#, dBottom#
    dTop = tblRectsDim.Num(2, 0)
    dBottom = tblRectsDim.Num(3, 0)
    'table record: 0=date1,1=date2,2=y1,3=y2
    'set valid type and coordinates info for applicable rects
    For i = 0 To tblRectsDim.NumRecords - 1
        gdSetNum m.geAnnStruct.glhAnnType, i, 25
        dX1 = GetX(Chart, tblRectsDim.Num(0, i)) - nStartScreenX '+ 1
        dX2 = GetX(Chart, tblRectsDim.Num(1, i)) - nStartScreenX '+ 1
        gdSetNum m.geAnnStruct.gdhTop, i, tblRectsDim.Num(2, i)
        gdSetNum m.geAnnStruct.gdhLeft, i, dX1
        gdSetNum m.geAnnStruct.gdhBottom, i, tblRectsDim.Num(3, i)
        gdSetNum m.geAnnStruct.gdhRight, i, dX2
    Next
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetMultiRects", eGDRaiseError_Raise

End Sub

Public Sub SetMultiRectsDim(Chart As cChart, ByVal nStartScreenX&, tblRectsDim As cGdTable)
    SetMultiRects Chart, nStartScreenX, tblRectsDim
End Sub

Public Sub Hide(ByVal nHide&, Optional Chart As cChart = Nothing, Optional ByVal bSetProperty As Boolean = False)
On Error Resume Next:

    If m.eUsage = eANNOT_UserAdded Then     '4189
    
        If m.eType = eANNOT_ElliotLabel Then
            'special case for EWI labels
            If bSetProperty Then
                'the bSetProperty flag is true when user clicks chart menu or button
                'to show/hide all or show/hide ewi labels created by analysts - 6984
                If Not IsEndUserEWI Then
                    Prop("IsHidden") = nHide
                    m.nHide = nHide
                End If
            ElseIf IsEndUserEWI Then
                'to see labels not created by analyst, must have flag files or EWL enablement
                If AllowEWI Or AllowGMP Or HasModule("EWL") Then
                    m.nHide = nHide
                Else
                    m.nHide = 1
                End If
            End If
        Else
            m.nHide = nHide
        End If
        
        If Not Chart Is Nothing Then
            geSetPoints Chart
        End If
    
    ElseIf m.eUsage = eANNOT_IndicatorLabel And bSetProperty Then
        'JM 03-15-2016 - quick mod for Glen's hot key toggle to show/hide chart's indicator labels
        If gdGetNum(m.geAnnStruct.glhAnnType, 0) = -1 Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, 13
        Else
            gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        End If
    End If
    
End Sub

Private Sub ShiftMultiRects(ByVal nShift&)
On Error Resume Next:

    Dim nSize&, nAnnType&, i&

    If nShift = 0 Then
        nAnnType = 25
    Else
        nAnnType = -1   'just hide these when moving chart
    End If
    
    nSize = gdGetSize(m.geAnnStruct.gdhTop) - 1
    For i = 0 To nSize
        gdSetNum m.geAnnStruct.glhAnnType, i, nAnnType
    Next

End Sub

Private Sub SetFibFan(Chart As cChart, ByVal dX1#, ByVal dX2#)
On Error GoTo ErrSection:

    'Developer Note: as of 07-30-2004
    'dX1=1st mouse click, dX2=2nd mouse click (this is where circle for handle is drawn)
    'The second mouse click is always the 1.00 ratio line (as implemented by eSignal)
    
    'array item setup info
    '[0]        set to be 1x1 line (end points being location of 1st & 2nd mouse clicks)
    '[1]        extension info
    '[2]..[5]   set to be ignored by graphics engine
    
    Dim i&, j&, k&
    Dim iExt&, nColor&, nStyle&
    Dim Y#, YLeft#, YRight#
    
    If m.eType = eANNOT_FibFan Then
        iExt = Val(Prop("Ext"))         'editor gives option to turn extension off
        nColor = Val(Prop("FibColor"))
        nStyle = Val(Prop("FibStyle"))
    ElseIf m.eType = eANNOT_SpResistFan Then
        iExt = 1                        'extension always on, user has no control
        nColor = Val(Prop("SpResistColor"))
        nStyle = Val(Prop("SpResistStyle"))
    End If
    
    If iExt = 0 Then
        If Prop("ShowRatioZero") = 1 Or Prop("ShowRatioOne") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, 1             'type=line (no extension)
        Else
            gdSetNum m.geAnnStruct.glhAnnType, 0, -1            'don't show 1x1 line
        End If
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
    Else
        If Prop("ShowRatioZero") = 1 Or Prop("ShowRatioOne") = 1 Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, 8             'type=trendline (extension calculated by grapheng.dll)
        Else
            gdSetNum m.geAnnStruct.glhAnnType, 0, -1             'don't show 1x1 line
        End If
        gdSetNum m.geAnnStruct.glhAnnType, 1, 8
        
        If m.eType = eANNOT_FibFan Then
            geSetPenStyle Val(Prop("ExtStyle")), 1
            gdSetNum m.geAnnStruct.glhPenColor, 1, -1
        Else
            geSetPenStyle nStyle, 1
            gdSetNum m.geAnnStruct.glhPenColor, 1, nColor
        End If
    End If
        
    '1x1 line
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    
    gdSetNum m.geAnnStruct.gdhMisc, 0, 1
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    geSetPenStyle m.nStyle, 0
    
    For i = 2 To 5
        gdSetNum m.geAnnStruct.glhAnnType, i, -1
    Next
    
    If m.dDate1 < m.dDate2 Then
        YLeft = m.Y1
        YRight = m.Y2
        m.bY1IsZeroRatio = False
    Else
        YLeft = m.Y2
        YRight = m.Y1
        m.bY1IsZeroRatio = True
    End If
    
    j = m.aRatios.Size - 1
    'fib lines
    For i = 0 To j
        If m.aRatioShow(i) = 1 Then
            If m.bUsingLogDrawingMode Then
                Y = Exp((Log(YLeft) - Log(YRight)) * m.aRatios(i) + Log(YRight))
            Else
                Y = (YLeft - YRight) * m.aRatios(i) + YRight
            End If
        
            If m.aRatioColor.Size > i Then
                gdSetNum m.geAnnStruct.glhPenColor, i + 6, m.aRatioColor(i)
            Else
                gdSetNum m.geAnnStruct.glhPenColor, i + 6, nColor
            End If
        
            If m.dDate1 < m.dDate2 Then
                gdSetNum m.geAnnStruct.glhAnnType, i + 6, 8
                If m.eType = eANNOT_SpResistFan Then
                    gdSetNum m.geAnnStruct.glhPenColor, i + 6, nColor
                End If
                geSetPenStyle nStyle, i + 6
                gdSetNum m.geAnnStruct.gdhTop, i + 6, m.Y1
                gdSetNum m.geAnnStruct.gdhBottom, i + 6, Y
            Else
                gdSetNum m.geAnnStruct.glhAnnType, i + 6, 8
                If m.eType = eANNOT_SpResistFan Then
                    gdSetNum m.geAnnStruct.glhPenColor, i + 6, nColor
                End If
                geSetPenStyle nStyle, i + 6
                gdSetNum m.geAnnStruct.gdhTop, i + 6, Y
                gdSetNum m.geAnnStruct.gdhBottom, i + 6, m.Y2
            End If
        
            gdSetNum m.geAnnStruct.gdhLeft, i + 6, dX1
            gdSetNum m.geAnnStruct.gdhRight, i + 6, dX2
            gdSetNum m.geAnnStruct.gdhMisc, i + 6, iExt     'extension flag
            gdSetStr m.geAnnStruct.gshText, i + 6, ""       'clear out text
        Else
            gdSetNum m.geAnnStruct.glhAnnType, i + 6, -1
        End If
    Next
    
    'misc field: 0=use extension color, 1=use main line color
    k = gdGetSize(m.geAnnStruct.glhAnnType) - 1
    For i = 1 To k
        gdSetNum m.geAnnStruct.gdhMisc, i, 0
    Next
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SetFibFan", eGDRaiseError_Raise

End Sub

Public Property Get SkipHitTestFlag() As Long
    SkipHitTestFlag = m.geAnnStruct.skipHitTest
End Property

Public Property Let SkipHitTestFlag(ByVal nFlag&)
    m.geAnnStruct.skipHitTest = nFlag
End Property

Public Sub SpeedResistRatiosChange(ByVal strRatios$, ByVal strShow$)
On Error GoTo ErrSection:

    If Len(strRatios) = 0 And Len(strShow) = 0 Then Exit Sub
    
    m.aRatios.Clear
    m.aRatioShow.Clear
    m.aRatioColor.Clear
    
    m.aRatios.SplitFields strRatios, ","
    m.aRatioShow.SplitFields strShow, ","
    
    Prop("SpResistRatios") = strRatios

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.SpeedResistRatiosChange", eGDRaiseError_Raise

End Sub

Public Property Get CreateDate() As Double
    CreateDate = m.dCreateDate
End Property

Public Property Get HitItemIndex() As Long
    HitItemIndex = m.nHitItemIdx
End Property

Public Property Let HitItemIndex(ByVal nIdx As Long)
    
    m.nHitItemIdx = nIdx
    
    If nIdx < 0 Then
    
        If m.eType = eANNOT_Trendline Or m.eType = eANNOT_Trendline2 Or _
           m.eType = eANNOT_Trendline3 Or m.eType = eANNOT_Trendline4 Or _
           m.eType = eANNOT_TrendChannel Then
           
           MoveTrendChannel Nothing, -1, m.geAnnStruct.paneId, 0, 0
           
        End If
    
    End If

End Property

Private Function MoveTrendChannel(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Static dPointsSave#, dPercentSave#
    
    Dim XDelta#, YDelta#
    Dim dPercent#, dY3#, i&

    Dim bMoved As Boolean
    
    If nPaneID <> m.geAnnStruct.paneId Then Exit Function
        
    If m.nShiftKey = 2 Then
        If X = m.X1 Or X = m.X2 Then
            If Not m.bOriginalVertical Then
                'if control key is down and user move endpoint such that it crosses over other one
                'don't let line go vertical if it did not start out as vertical
                Exit Function
            End If
        End If
    End If
    
    If Val(Prop("ChannelCount")) < 1 Then
        bMoved = MoveTrendline(Chart, nPoint, nPaneID, X, Y)
        MoveTrendChannel = bMoved
        Exit Function
    End If
    
    Dim eUseLog As ePANE_LogFlag
    Dim Pane As cPane
    If Not m.Chart Is Nothing Then
        Set Pane = m.Chart.Tree(m.strPane)
        If Not Pane Is Nothing Then
            eUseLog = Pane.PaneLogFlag
            Set Pane = Nothing
        End If
    End If
    
    XDelta = m.X2 - m.X1
    If m.bUsingLogDrawingMode Then
        YDelta = LogModeDelta(m.Y1, m.Y2)           'Log(m.Y2) - Log(m.Y1)
    Else
        YDelta = m.Y2 - m.Y1
    End If
    
    If m.nHitItemIdx = 1 Or m.nHitItemIdx = 2 Or m.nHitItemIdx = 3 Then
        Exit Function
    ElseIf m.nHitItemIdx = -1 Or dPointsSave = 0 Then
        dPointsSave = Val(Prop("ChannelPoints"))            '5403
        dPercentSave = Val(Prop("ChannelPercent"))
    ElseIf m.nHitItemIdx > 3 And nPoint <> 0 Then
        dPercent = Val(Prop("ChannelPercent")) / 100
        If Val(Prop("ChannelType")) = 1 Then
            dY3 = Val(Prop("ChannelPoints"))
        ElseIf m.X1 > m.X2 Then
            dY3 = m.Y1 * dPercent
        Else
            dY3 = m.Y2 * dPercent
        End If
        i = Int(m.nHitItemIdx / 2) - 1
        dY3 = Abs(dY3 * i)
        If m.nHitItemIdx Mod 2 = 0 Then dY3 = dY3 * -1
    End If
    
    bMoved = True
    
    i = Val(Prop("ChannelLocation"))
    
    Select Case nPoint
        Case 0
            If m.nHitItemIdx = 0 Then
                'mid-point on main line (move whole line)
                m.X1 = X - m.dMoveOffSetX
                m.X2 = m.X1 + XDelta
                If m.bUsingLogDrawingMode Then
                    m.Y1 = Exp(Log(Y) - m.dMoveOffSetLogY)
                    m.Y2 = Exp(Log(m.Y1) + YDelta)
                Else
                    m.Y1 = Y - m.dMoveOffSetY
                    m.Y2 = m.Y1 + YDelta
                End If
            ElseIf dPointsSave <> 0 Then
                'mid-point on channel (change points & percent values)
                dY3 = geTrendValueY(m.geAnnStruct.hObj, X - m.Chart.ScreenStartX)
                
                If i <> 3 Then
                    If dY3 > Y Then
                        Prop("ChannelLocation") = 2     'below
                    Else
                        Prop("ChannelLocation") = 1     'above
                    End If
                End If
                
                If eUseLog = ePANE_LogFlagLog And dY3 > 0 And Y > 0 And m.Y2 > 0 Then
                    dY3 = Exp(Log(m.Y2) + Abs(Log(dY3) - Log(Y)))       '6551
                Else
                    dY3 = m.Y2 + Abs(dY3 - Y)
                End If
                
                i = Int(m.nHitItemIdx / 2) - 1
                If i > 0 Then       'precautionary; should always be true
                    dPercent = (dY3 - m.Y2) / m.Y2 / i
                    Prop("ChannelPoints") = Abs(m.Y2 * dPercent)
                    Prop("ChannelPercent") = Abs(RoundNum(dPercent * 100, 2))
                End If
            End If
            
        Case 1
            m.X1 = RoundNum(X)
            If m.nHitItemIdx = 0 Then
                m.Y1 = Y
            Else
                m.Y1 = Y + dY3
            End If
                    
        Case 2
            m.X2 = RoundNum(X)
            If m.nHitItemIdx = 0 Then
                m.Y2 = Y
            Else
                m.Y2 = Y + dY3
            End If
                    
        Case Else
            bMoved = False
    End Select
    
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
    End If
    
    MoveTrendChannel = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveTrendChannel", eGDRaiseError_Raise

End Function

Public Sub AddChannelOnMove(ByVal YPrev#, ByVal Y#)
On Error GoTo ErrSection:

    Dim dY3#, dPercent#, i&
    
    Dim eUseLog As ePANE_LogFlag
    Dim Pane As cPane
    If Not m.Chart Is Nothing Then
        Set Pane = m.Chart.Tree(m.strPane)
        If Not Pane Is Nothing Then
            eUseLog = Pane.PaneLogFlag
            Set Pane = Nothing
        End If
    End If
    
    Prop("ChannelCount") = 1
    
    i = LoadDefault("ChannelLocation", 3)
    If i <= 0 Or i > 3 Then i = 3
    
    If m.eType = eANNOT_TrendChannel Then
        dY3 = geTrendValueY(m.geAnnStruct.hObj, GetX(m.Chart, YPrev) - m.Chart.ScreenStartX)
        
        If i <> 3 Then
            If dY3 > Y Then
                Prop("ChannelLocation") = 2     'below
            Else
                Prop("ChannelLocation") = 1     'above
            End If
        End If
        
        If eUseLog = ePANE_LogFlagLog And dY3 > 0 And Y > 0 And m.Y2 > 0 Then
            dY3 = Exp(Log(m.Y2) + Abs(Log(dY3) - Log(Y)))
        Else
            dY3 = m.Y2 + Abs(dY3 - Y)
        End If
        
        dPercent = ((dY3 - m.Y2) / m.Y2) * 100
    
        Prop("ChannelPercent") = Abs(RoundNum(dPercent, 2))
    Else
        'this code is hit when user holds down shift key and move an existing regular trend line
        Prop("ChannelLocation") = 1
    
        dY3 = m.Y2 + Abs(Y - YPrev)
        dPercent = ((dY3 - m.Y2) / m.Y2) * 100
    
        Prop("ChannelPercent") = Abs(RoundNum(dPercent, 2))
    
        m.nHitItemIdx = 4
    End If
    
    i = LoadDefault("ChannelCount", 1)
    If i > 1 Then Prop("ChannelCount") = i

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.AddChannelOnMove", eGDRaiseError_Raise

End Sub

Private Function MoveFibTimeZones(Chart, nPoint, nPaneID, X, Y) As Boolean
On Error GoTo ErrSection:

    Dim bMoved As Boolean
    Dim XDelta#
        
    XDelta = m.X2 - m.X1
    
    bMoved = True
    
    If nPoint = 1 And Val(Prop("Arcs")) = 1 Then
        geCalcArcHeight m.geAnnStruct, Y
        Prop("ArcPercentHeight") = m.geAnnStruct.lenRatio
    ElseIf nPoint = 0 Then      'mid-point (move whole line)
        m.X1 = X - m.dMoveOffSetX
        m.X2 = m.X1 + XDelta
    Else
        bMoved = False
    End If

    MoveFibTimeZones = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveFibTimeZones", eGDRaiseError_Raise

End Function

Private Function MoveMirror(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
        
    Dim bMoved As Boolean
    Dim XDelta#, X3#
    
    If nPaneID <> m.geAnnStruct.paneId Or m.eUsage <> eANNOT_UserAdded Then
        Exit Function
    End If
       
    XDelta = Abs(m.X2 - m.X1)
    
    bMoved = True
    Select Case nPoint
        Case 0, 2
            bMoved = MoveRectangle(Chart, nPoint, nPaneID, X, Y)
                
        Case 7
            If m.X1 < m.X2 Then
                If Abs(m.X1 - X) < Abs(m.X2 - X) Then
                    If X >= m.X2 Then
                        bMoved = False
                    Else
                        m.X1 = X
                    End If
                Else
                    If X <= m.X2 Then
                        bMoved = False
                    Else
                        X3 = m.X2 + XDelta
                        m.X1 = m.X1 + (X3 - X)
                    End If
                End If
            Else
                If Abs(m.X2 - X) < Abs(m.X1 - X) Then
                    If X >= m.X1 Then
                        bMoved = False
                    Else
                        m.X2 = X
                    End If
                Else
                    If X <= m.X1 Then
                        bMoved = False
                    Else
                        X3 = m.X1 + XDelta
                        m.X2 = m.X2 + (X3 - X)
                    End If
                End If
            End If
        
        Case Else
            bMoved = False
    End Select
    
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
    End If
        
    MoveMirror = bMoved
    
End Function

Private Function MovePattern(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean

    Dim bMoved As Boolean
    Dim dIndMin#, dIndMax#
    Dim XDelta#, XSave#
    
    If nPaneID <> m.geAnnStruct.paneId Or m.eUsage <> eANNOT_UserAdded Then
        Exit Function
    End If
          
    bMoved = True
    If nPoint = 3 Then
        If Val(Prop("ForecastBars")) > 0 Then
            XDelta = (X - GetX(Chart, m.aDates(1)))
            If XDelta >= 1 Then
                If XDelta <> Val(Prop("ForecastBars")) Then
                    ForecastBarsChange Chart, 1, XDelta
                    bMoved = True
                End If
            Else
                bMoved = False
            End If
        Else
            XSave = m.X2
            XDelta = X - GetX(Chart, m.aDates(1))
            m.X2 = m.X2 + XDelta
            If m.X2 > m.X1 Then
                m.dDate2 = Chart.aXdate(m.X2)
                m.aDates(1) = Chart.aXdate(X)
                PatternMinMax Chart, dIndMin, dIndMax, m.aDates(0), m.aDates(1)
                m.aYs(0) = dIndMin
                m.aYs(1) = dIndMax
                bMoved = True
            Else
                bMoved = False
                m.X2 = XSave
            End If
        End If
    ElseIf nPoint = 7 Then
        XSave = m.X1
        XDelta = X - GetX(Chart, m.aDates(0))
        m.X1 = m.X1 + XDelta
        If m.X1 < m.X2 Then
            m.dDate1 = Chart.aXdate(m.X1)
            m.aDates(0) = Chart.aXdate(X)
            PatternMinMax Chart, dIndMin, dIndMax, m.aDates(0), m.aDates(1)
            m.aYs(0) = dIndMin
            m.aYs(1) = dIndMax
            bMoved = True
        Else
            bMoved = False
            m.X1 = XSave
        End If
    ElseIf m.nHitItemIdx < 0 Then
        bMoved = MoveRectangle(Chart, nPoint, nPaneID, X, Y)
    ElseIf m.nHitItemIdx = 0 Then
        If m.aDates.Size = 0 Then
            m.aDates(0) = m.dDate1
            m.aDates(1) = m.dDate2
            m.aYs(0) = m.Y1
            m.aYs(1) = m.Y2
            gdSetNum m.geAnnStruct.gdhTop, 1, gdGetNum(m.geAnnStruct.gdhTop, 0)
            gdSetNum m.geAnnStruct.gdhLeft, 1, gdGetNum(m.geAnnStruct.gdhLeft, 0)
            gdSetNum m.geAnnStruct.gdhBottom, 1, gdGetNum(m.geAnnStruct.gdhBottom, 0)
            gdSetNum m.geAnnStruct.gdhRight, 1, gdGetNum(m.geAnnStruct.gdhRight, 0)
        ElseIf nPoint = 0 Then
            bMoved = MoveRectangle(Chart, nPoint, nPaneID, X, Y)
        End If
    End If
    
    If bMoved Then
        ' make sure is in range
        If m.X1 < 0 Or m.X1 >= Chart.aXdate.Size Then bMoved = False
        If m.X2 < 0 Or m.X2 >= Chart.aXdate.Size Then bMoved = False
    End If
        
    MovePattern = bMoved

End Function

Public Property Get geAnnotObject() As Long
    geAnnotObject = m.geAnnStruct.hObj
End Property

Public Property Get gePointNum(Chart As cChart, ByVal dDate#)
    gePointNum = GetX(Chart, dDate) - Chart.ScreenStartX
End Property

Private Function GetProfitLoss(Chart As cChart, dValue1#, dValue2, strPriceMove$) As Double
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim Pane As cPane
    Dim Ind As cIndicator
    Dim dAmt#, dMinMove#, dGannacci#, iQty&, strBaseForex$, strSymbol$
    Static BaseForexBars As cGdBars
    Static nPrevLDD&, strPrevSymbol$

    If Chart Is Nothing Then
        strPriceMove = "INVALID"
        Exit Function
    ElseIf Len(Chart.SpreadSymbols) > 0 Then
        'aardvark 4063: spread charts
        'auto-multiplier must be on or all components have same tickMove & tickValue
        If Not Chart.SpreadAsDollar Then
            If Not SpreadComponentsOk(Chart) Then
                strPriceMove = "INVALID"
                Exit Function
            End If
        End If
    End If
    
    ' round Y values to the nearest min move for price
    Set Pane = Chart.Tree(m.strPane)
    If Not Pane Is Nothing Then
        Set Ind = Chart.Tree(Pane.idxPriceBars)
        If Not Ind Is Nothing Then
            Set Bars = Ind.Bars
        End If
    End If
    If Bars Is Nothing Then
        Set Bars = Chart.Bars
    End If
    strPriceMove = ""
    
    If Bars Is Nothing Then GoTo ErrExit
        
    dMinMove = Bars.MinMove(m.dDate1)
    If dMinMove > 0 Then
        dValue1 = Int(dValue1 / dMinMove + 0.5) * dMinMove
        dValue2 = Int(dValue2 / dMinMove + 0.5) * dMinMove
        'set price move string
        dAmt = Bars.Prop(eBARS_TickMove)
        If dAmt > 0 Then
            If m.eType = eANNOT_GannacciSwingSquare Then
                dAmt = dValue2 - dValue1
                If Val(Prop("UseMutiplier")) = 1 Then
                    dGannacci = Val(Prop("MultiplierVal"))
                    If dGannacci <> 0 Then dAmt = dAmt * dGannacci
                End If
                strPriceMove = FormatNum(dAmt)
                GoTo ErrExit            'do not need profit loss calcuations
            Else
                strPriceMove = FormatNum(Abs((dValue1 - dValue2) / dAmt))
            End If
        End If
    End If
    If Bars.SecurityType = "S" Then
        iQty = Val(Prop("NumShares"))       '5215
    ElseIf IsForex(Bars.Prop(eBARS_Symbol)) Then
        iQty = Val(Prop("NumForexContracts"))
    Else
        iQty = Val(Prop("NumContracts"))
    End If
    
    If iQty = 0 Then iQty = 1           '5397
    
    ' DAJ 02/22/2010: Pass the second date in (which will act as the exit date) so that if
    ' the chart is for a non-US Dollar based Forex, we can go get the exchange rate from
    ' the day before to be consistent with trading profit and loss (Issue #5402)...
    strSymbol = Bars.Prop(eBARS_Symbol)
    strBaseForex = g.Profit.BaseForex(strSymbol)
    
    If Len(strBaseForex) = 0 Then
        Set BaseForexBars = Nothing
    ElseIf BaseForexBars Is Nothing Or nPrevLDD <> LastDailyDownload Or strPrevSymbol <> strSymbol Then
        strPrevSymbol = strSymbol
        nPrevLDD = LastDailyDownload
        Set BaseForexBars = New cGdBars
        DM_GetBars BaseForexBars, strBaseForex
    End If
    dAmt = g.Profit.Profit(Bars.Prop(eBARS_Symbol), dValue2 - dValue1, _
            iQty, MaxDouble(m.dDate1, m.dDate2), Bars, BaseForexBars, Chart.TradeAccountID) '5183


ErrExit:
    Set Pane = Nothing
    Set Ind = Nothing
    GetProfitLoss = dAmt
    
    Exit Function
    
ErrSection:
    Set Pane = Nothing
    Set Ind = Nothing
    GetProfitLoss = 0
    
    RaiseError "cAnnotation.GetProfitLoss", eGDRaiseError_Raise

End Function

Public Property Get YArray() As cGdArray
    Set YArray = m.aYs
End Property

Public Property Get DateArray() As cGdArray
    
    Set DateArray = m.aDates

End Property

Public Property Get DateFromArray(ByVal nIndex&)
    DateFromArray = m.aDates(nIndex)
End Property

Public Property Get YFromArray(ByVal nIndex&) As Double
    YFromArray = m.aYs(nIndex)
End Property

Public Property Let YFromArray(ByVal nIndex&, ByVal dValue#)
    m.aYs(nIndex) = dValue
End Property


Private Sub PatternMinMax(Chart As cChart, dyMin#, dyMax#, ByVal dDate1#, ByVal dDate2#, _
    Optional ByRef InputInd As cIndicator = Nothing)
On Error GoTo ErrSection:

    Dim nFirstBar&, nLastBar&, i&
    Dim dMax#, dMin#
    Dim bPrice As Boolean
    
    Dim Ind As cIndicator
        
    'initialize return values
    dyMin = kNullData
    dyMax = kNullData
    
    If InputInd Is Nothing Then
        Set Ind = Chart.Tree("PRICE")
    Else
        Set Ind = InputInd
    End If
    If Ind Is Nothing Then Exit Sub
    
    If UCase(Ind.Name) = "PRICE" Then
        bPrice = True
    ElseIf Ind.Data Is Nothing Then
        Exit Sub
    ElseIf Ind.Data.Size <> Chart.Bars.Size - Chart.ForecastBars Then
        Exit Sub
    End If
    
    nFirstBar = Chart.Bars.FindDateTime(dDate1, True)
    If nFirstBar >= 0 Then nLastBar = Chart.Bars.FindDateTime(dDate2, True)
    
    If nFirstBar >= 0 And nLastBar >= 0 Then
        If nLastBar < nFirstBar Then
            i = nFirstBar       'precautionary, should never happen
            nFirstBar = nLastBar
            nLastBar = i
        End If
        nLastBar = nLastBar + Val(Prop("ForecastBars"))
        If bPrice Then
            dMax = gdMaxValue(Chart.Bars.ArrayHandle(eBARS_High), nFirstBar, nLastBar)
            dMin = gdMinValue(Chart.Bars.ArrayHandle(eBARS_Low), nFirstBar, nLastBar)
        Else
            dMax = gdMaxValue(Ind.Data.ArrayHandle, nFirstBar, nLastBar)
            dMin = gdMinValue(Ind.Data.ArrayHandle, nFirstBar, nLastBar)
        End If
        
        dyMin = dMin
        dyMax = dMax
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.PatternMinMax", eGDRaiseError_Raise

End Sub

Public Sub ForecastBarsChange(Chart As cChart, ByVal nIdx&, ByVal nBars&)
On Error GoTo ErrSection:

    Dim dDate1#, dDate2#
    Dim dIndMin#, dIndMax#

    If nBars = m.geAnnStruct.minorAxisLenData Then Exit Sub

    If nIdx = 0 Then
        dDate1 = m.dDate1
        dDate2 = m.dDate2
    ElseIf nIdx = 1 Then
        dDate1 = m.aDates(0)
        dDate2 = m.aDates(1)
    Else
        Exit Sub
    End If

    Prop("ForecastBars") = nBars
    PatternMinMax Chart, dIndMin, dIndMax, dDate1, dDate2

    If dIndMin <> kNullData And dIndMax <> kNullData Then
        If nIdx = 0 Then
            m.Y1 = dIndMin
            m.Y2 = dIndMax
        Else
            m.aYs(0) = dIndMin
            m.aYs(1) = dIndMax
        End If
        m.geAnnStruct.minorAxisLenData = nBars
    Else
        Prop("ForecastBars") = m.geAnnStruct.minorAxisLenData
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.ForecastBarsChange", eGDRaiseError_Raise

End Sub

Private Function MoveChannelHighlight(Chart As cChart, ByVal nPoint&, _
        ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:
    
    Dim XDelta#, YDelta#, dX3#, dX4#
    Dim dX3Delta#, dY3Delta#, dX4Delta#, dY4Delta#
    Dim dX1Old#, dX2Old#, dY1Old#, dY2Old#, dY3Old#, dY4Old#
    Dim dDate1Old#, dDate2Old#, dDate3Old#, dDate4Old#
    
    Dim hArrayDate&, dNewDate#
    
    If nPaneID <> m.geAnnStruct.paneId Then Exit Function
    If nPoint = 0 And m.geAnnStruct.moveable = 0 Then Exit Function
    
    MoveChannelHighlight = True
            
    XDelta = m.X2 - m.X1
    
    dX3 = GetX(Chart, m.aDates(0))
    dX3Delta = dX3 - m.X1
    
    dX4 = GetX(Chart, m.aDates(1))
    dX4Delta = dX4 - m.X1

    If m.bUsingLogDrawingMode Then
        YDelta = LogModeDelta(m.Y1, m.Y2)                                   'Log(m.Y2) - Log(m.Y1)
        If m.aYs.Size >= 1 Then dY3Delta = LogModeDelta(m.Y1, m.aYs(0))     'Log(m.aYs(0)) - Log(m.Y1)
        If m.aYs.Size >= 2 Then dY4Delta = LogModeDelta(m.Y1, m.aYs(1))     'Log(m.aYs(1)) - Log(m.Y1)
    Else
        YDelta = m.Y2 - m.Y1
        dY3Delta = m.aYs(0) - m.Y1
        dY4Delta = m.aYs(1) - m.Y1
    End If

    
    hArrayDate = Chart.geDateArray      '4240
    dNewDate = gdGetNum(hArrayDate, gdGetSize(hArrayDate) - 1)
    If X < GetX(Chart, dNewDate) And X > GetX(Chart, gdGetNum(hArrayDate, 0)) Then
        dNewDate = Chart.aXdate(X)
    Else
       X = GetX(Chart, dNewDate)
    End If
    
    Select Case nPoint
        Case 0:
            'save current values
            dX1Old = m.X1
            dX2Old = m.X2
            dY1Old = m.Y1
            dY2Old = m.Y2
            dY3Old = m.aYs(0)
            dY4Old = m.aYs(1)
            dDate1Old = m.dDate1
            dDate2Old = m.dDate2
            dDate3Old = m.aDates(0)
            dDate4Old = m.aDates(1)
            'set moved values
            m.X1 = X - m.dMoveOffSetX
            m.X2 = m.X1 + XDelta
            dX3 = m.X1 + dX3Delta
            dX4 = m.X1 + dX4Delta
            
            If m.bUsingLogDrawingMode Then
                m.Y1 = Exp(Log(Y) - m.dMoveOffSetLogY)
                m.Y2 = Exp(Log(m.Y1) + YDelta)
                If m.aYs.Size >= 1 Then m.aYs(0) = Exp(Log(m.Y1) + dY3Delta)
                If m.aYs.Size >= 2 Then m.aYs(1) = Exp(Log(m.Y1) + dY4Delta)
            Else
                m.Y1 = Y - m.dMoveOffSetY
                m.Y2 = m.Y1 + YDelta
                m.aYs(0) = m.Y1 + dY3Delta
                m.aYs(1) = m.Y1 + dY4Delta
            End If
            
            If m.dDate1 <> 0 Then m.dDate1 = Chart.aXdate(m.X1)
            If m.dDate2 <> 0 Then m.dDate2 = Chart.aXdate(m.X2)
            If m.aDates(0) <> 0 Then m.aDates(0) = Chart.aXdate(dX3)
            If m.aDates(1) <> 0 Then m.aDates(1) = Chart.aXdate(dX4)
            'restore saved values if going past end of data
            If m.dDate1 < 0 Or m.dDate2 < 0 Or m.aDates(0) < 0 Or m.aDates(1) < 0 Then
                m.X1 = dX1Old
                m.X2 = dX2Old
                m.Y1 = dY1Old
                m.Y2 = dY2Old
                m.aYs(0) = dY3Old
                m.aYs(1) = dY4Old
                m.dDate1 = dDate1Old
                m.dDate2 = dDate2Old
                m.aDates(0) = dDate3Old
                m.aDates(1) = dDate4Old
                MoveChannelHighlight = False
            End If
        
        Case 1:
            m.X1 = X
            m.Y1 = Y
            If m.dDate1 > 0 Then m.dDate1 = dNewDate
        Case 2:
            m.X2 = X
            m.Y2 = Y
            If m.dDate2 > 0 Then m.dDate2 = dNewDate
        Case 3:
            m.aYs(0) = Y
            If m.aDates(0) <> 0 Then m.aDates(0) = dNewDate
        Case 4:
            m.aYs(1) = Y
            If m.aDates(1) <> 0 Then m.aDates(1) = dNewDate
        Case Else
            MoveChannelHighlight = False
    End Select
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveChannelHighlight", eGDRaiseError_Raise

End Function

Public Property Get ClosestPoint(Chart As cChart, ByVal dX#, ByVal dY#) As Long
On Error GoTo ErrSection:
'this routine intended for use by certain types of annotations
'that have more than 2 points and drawn with multiple clicks
'returns:
'1 if passed in values closest to X1 and Y1
'2 if passed in values closest to X2 and Y2
'3 .. n if passed in values closest to something in the aYs/aDates arrays
'-1 = error

    Dim tbPoints As New cGdTable
    Dim aIdx As New cGdArray
    Dim dAnnotX#, dAnnotY#
    Dim i&, j&, nSize&
    
    If m.eType <> eANNOT_AndrewFork And m.eType <> eANNOT_ChannelHighlight And _
       m.eType <> eANNOT_TriangleWedge And m.eType <> eANNOT_WaveLabels And _
       m.eType <> eANNOT_ElliotTimeRatio Then
       
       ClosestPoint = -1
       Exit Property
       
    End If
            
    nSize = m.aDates.Size
    If nSize <> m.aYs.Size Then     'precautionary
        ClosestPoint = -1
        Exit Property
    End If
        
    'create fields for table
    tbPoints.CreateField eGDARRAY_Shorts, 0, "PointNumber"
    tbPoints.CreateField eGDARRAY_Doubles, 1, "X"
    tbPoints.CreateField eGDARRAY_Doubles, 2, "Y"
    tbPoints.CreateField eGDARRAY_Doubles, 3, "DiffX"
    tbPoints.CreateField eGDARRAY_Doubles, 4, "DiffY"
    
    tbPoints.AddRecord ""
    j = tbPoints.NumRecords - 1
    dAnnotX = GetX(Chart, m.dDate1)
    dAnnotY = m.Y1
    tbPoints(0, j) = 1
    tbPoints(1, j) = dAnnotX
    tbPoints(2, j) = dAnnotY
    tbPoints(3, j) = Abs(dX - dAnnotX)
    tbPoints(4, j) = Abs(dY - dAnnotY)
    
    tbPoints.AddRecord ""
    j = tbPoints.NumRecords - 1
    dAnnotX = GetX(Chart, m.dDate2)
    dAnnotY = m.Y2
    tbPoints(0, j) = 2
    tbPoints(1, j) = dAnnotX
    tbPoints(2, j) = dAnnotY
    tbPoints(3, j) = Abs(dX - dAnnotX)
    tbPoints(4, j) = Abs(dY - dAnnotY)
    
    For i = 0 To nSize - 1
        tbPoints.AddRecord ""
        j = tbPoints.NumRecords - 1
        dAnnotX = GetX(Chart, m.aDates(i))
        dAnnotY = m.aYs(i)
        tbPoints(0, j) = i + 3
        tbPoints(1, j) = dAnnotX
        tbPoints(2, j) = dAnnotY
        tbPoints(3, j) = Abs(dX - dAnnotX)
        tbPoints(4, j) = Abs(dY - dAnnotY)
    Next
    
    aIdx.Create eGDARRAY_Longs, tbPoints.NumRecords
    gdSortAsIndex aIdx.ArrayHandle, tbPoints.FieldArrayHandle(4), 1, eGdSort_Stable, 0, -1
    gdSortAsIndex aIdx.ArrayHandle, tbPoints.FieldArrayHandle(3), 0, eGdSort_Stable, 0, -1
    
    ClosestPoint = tbPoints(0, aIdx(0))
    
    Set aIdx = Nothing
    Set tbPoints = Nothing
        
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.ClosestPoint", eGDRaiseError_Raise

End Property

Private Function SetWaveLabels(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&) As Long
On Error GoTo ErrSection:

    Dim i&, j&, k&, nLines&, nArc&
    Dim XDelta#, YDelta#, dFudge#
    Dim dX#, dY#, dPrevX#, dPrevY#, dNextY#
    Dim aLabels As New cGdArray
    Dim strLabels$

    
    gdSetSize m.geAnnStruct.glhAnnType, 0, 0    'clear this in case user added/removed points

'    commented out as fix for aardvark 2451 ( wave labels disappear ... )
'    If MultiPointSetup(Chart, XDelta, YDelta, dFudge) = False Then
'        SetWaveLabels = 0
'        Exit Function
'    End If

    nArc = Val(Prop("PointArc"))
    nLines = Val(Prop("Lines"))
    If nLines = 0 Then nLines = -1      'setting this to -1 will by pass drawing of lines
    
    gdSetNum m.geAnnStruct.glhAnnType, 0, nLines
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2

    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    gdSetNum m.geAnnStruct.glhAlign, 0, 0           'place holder
    geSetPenStyle m.nStyle, 0
    
    dPrevX = dX2
    dPrevY = m.Y2
        
    'set lines
    For i = 0 To m.aDates.Size - 1
        dX = GetX(Chart, m.aDates(i)) - nStartScreenX           '4525
        dY = m.aYs(i)
        
        If dX = dPrevX And dY = dPrevY Then
            m.aDates.Size = m.aDates.Size - 1
            m.aYs.Size = m.aYs.Size - 1
            Exit For
        End If

        gdSetNum m.geAnnStruct.glhAnnType, i + 1, nLines
        gdSetNum m.geAnnStruct.gdhTop, i + 1, dPrevY
        gdSetNum m.geAnnStruct.gdhLeft, i + 1, dPrevX
        gdSetNum m.geAnnStruct.gdhBottom, i + 1, dY
        gdSetNum m.geAnnStruct.gdhRight, i + 1, dX
        
        gdSetNum m.geAnnStruct.glhPenColor, i + 1, m.nColor
        gdSetNum m.geAnnStruct.glhAlign, i + 1, 0               'place holder
        gdSetNum m.geAnnStruct.glhImageType, i + 1, -1          'place holder
        gdSetNum m.geAnnStruct.glhDirection, i + 1, -1          'place holder
                
        dPrevX = dX
        dPrevY = dY
    Next
        
    dPrevX = 0
    dPrevY = 0
    i = 0
    If InStr(m.strText, ",") > 0 Then           'no commas in string means no labels
        If Val(Prop("LabelFirstPoint")) = 0 Then
            'original annotation: drew with lines and no label on first point
            'if user changed annotation to no lines then indicate first point with ?
            If nLines > 0 Then
                strLabels = " ," & m.strText
            Else
                strLabels = "?," & m.strText
            End If
        Else
            strLabels = m.strText
        End If
        aLabels.SplitFields strLabels, ","
        i = gdGetSize(m.geAnnStruct.glhAnnType)
    ElseIf nArc = 1 Then                        'user wants arcs with no labels, pad with spaces
        i = gdGetSize(m.geAnnStruct.glhAnnType)
        strLabels = " "
        For j = 1 To i
            strLabels = strLabels & ", "
        Next
        aLabels.SplitFields strLabels, ","
    End If
        
    AdjustLabelPastEnd aLabels, nArc
        
    'set labels
    For j = 0 To i - 1
        dX = gdGetNum(m.geAnnStruct.gdhLeft, j)
        dY = gdGetNum(m.geAnnStruct.gdhTop, j)
        If (j + 1) <= (i - 1) Then
            dNextY = gdGetNum(m.geAnnStruct.gdhTop, j + 1)
            If dNextY > 0 Then
                If dY < dNextY Then
                    gdSetNum m.geAnnStruct.glhDirection, j + i, 4
                Else
                    gdSetNum m.geAnnStruct.glhDirection, j + i, 0
                End If
            End If
        End If
        
        gdSetNum m.geAnnStruct.glhAnnType, j + i, 0
        
        gdSetNum m.geAnnStruct.gdhTop, j + i, dY
        gdSetNum m.geAnnStruct.gdhLeft, j + i, dX
        gdSetNum m.geAnnStruct.gdhBottom, j + i, dY
        gdSetNum m.geAnnStruct.gdhRight, j + i, dX
        
        gdSetStr m.geAnnStruct.gshText, j + i, aLabels(j)
        gdSetNum m.geAnnStruct.glhPenColor, j + i, m.nColor
        
        If gdGetNum(m.geAnnStruct.glhDirection, j + i) = 4 Then
            gdSetNum m.geAnnStruct.glhAlign, j + i, 2   'e_topCtr
        Else
            gdSetNum m.geAnnStruct.glhAlign, j + i, 5   'e_btmCtr
        End If
        
        If nArc = 1 Then
            gdSetNum m.geAnnStruct.glhImageType, j + i, 10
        Else
            gdSetNum m.geAnnStruct.glhImageType, j + i, -1
        End If
    Next
    
    'set last label
    j = gdGetSize(m.geAnnStruct.glhAnnType)

    dX = gdGetNum(m.geAnnStruct.gdhRight, i - 1)
    dY = gdGetNum(m.geAnnStruct.gdhBottom, i - 1)
    dPrevY = gdGetNum(m.geAnnStruct.gdhTop, i - 1)
    
    If dY > dPrevY Then
        gdSetNum m.geAnnStruct.glhDirection, j - 1, 4
        gdSetNum m.geAnnStruct.glhDirection, j, 0
        gdSetNum m.geAnnStruct.glhAlign, j, 5
    Else
        gdSetNum m.geAnnStruct.glhDirection, j - 1, 0
        gdSetNum m.geAnnStruct.glhDirection, j, 4
        gdSetNum m.geAnnStruct.glhAlign, j, 2
    End If
    
    If nArc = 1 Then
        gdSetNum m.geAnnStruct.glhImageType, j, 10
    Else
        gdSetNum m.geAnnStruct.glhImageType, j, -1
    End If

    gdSetNum m.geAnnStruct.glhAnnType, j, 0

    gdSetNum m.geAnnStruct.gdhTop, j, dY
    gdSetNum m.geAnnStruct.gdhLeft, j, dX
    gdSetNum m.geAnnStruct.gdhBottom, j, dY
    gdSetNum m.geAnnStruct.gdhRight, j, dX

    gdSetStr m.geAnnStruct.gshText, j, aLabels(m.aDates.Size + 1)
    gdSetNum m.geAnnStruct.glhPenColor, j, m.nColor
    
    'set font info
    For i = 0 To j
        gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
        gdSetNum m.geAnnStruct.glhStyle, i, Prop("FontStyle")
        gdSetNum m.geAnnStruct.glhSize, i, Prop("FontSize")
        gdSetNum m.geAnnStruct.glhUnderline, i, Prop("FontUnderline")
    Next
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetWaveLabels", eGDRaiseError_Raise

End Function

Private Function MoveWaveLabels(Chart As cChart, ByVal nPoint&, _
    ByVal nPaneID&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    If nPoint = 1 Then
        m.X1 = X
        m.Y1 = Y
        If m.dDate1 > 0 Then m.dDate1 = Chart.aXdate(X)
    ElseIf nPoint = 2 Then
        m.X2 = X
        m.Y2 = Y
        If m.dDate2 > 0 Then m.dDate2 = Chart.aXdate(X)
    Else
        m.aYs(nPoint - 3) = Y
        If m.aDates(nPoint - 3) <> 0 Then m.aDates(nPoint - 3) = Chart.aXdate(X)
    End If
    
    MoveWaveLabels = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.MoveWaveLabels", eGDRaiseError_Raise

End Function

Public Property Get Extendable(ByVal nPoint&) As Long
On Error GoTo ErrSection:
'returns:  0 = invalid annotation type
'          1 = selected point is first point
'         <0 = selected point is in the middle
'         >0 = selected point is last point
    
    Dim nExtend&
    
    nExtend = 0
    If m.eType = eANNOT_WaveLabels Then
        nExtend = nPoint
        If m.aDates.Size > 0 Then
            If nPoint <> 1 And nPoint <> m.aDates.Size + 2 Then
                nExtend = nPoint * -1
            End If
        End If
    End If
    
    Extendable = nExtend

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.Extendable", eGDRaiseError_Raise

End Property

Public Property Get MenuMove() As Boolean
    MenuMove = m.bMoveByMenu
End Property

Public Property Let MenuMove(ByVal bMenuMove As Boolean)
    m.bMoveByMenu = bMenuMove
End Property

Public Property Get MenuAdd() As Boolean
    MenuAdd = m.bAddByMenu
End Property

Public Property Let MenuAdd(ByVal bAdd As Boolean)
    m.bAddByMenu = bAdd
End Property

Public Property Get TotalPoints() As Long
    If m.eType = eANNOT_Icon Or m.eType = eANNOT_ElliotLabel Or m.eType = eANNOT_IndicLabel Then
        TotalPoints = 1
    Else
        TotalPoints = m.aDates.Size + 2
    End If
End Property

Public Sub DeletePoint(ByVal nPoint)
On Error GoTo ErrSection:

    If m.eType = eANNOT_WaveLabels Then
        If nPoint = m.aDates.Size + 2 Then
            m.aDates.Size = m.aDates.Size - 1
            m.aYs.Size = m.aYs.Size - 1
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.DeletePoint", eGDRaiseError_Raise

End Sub

Public Function AddPoint(ByVal dDate#, ByVal dY#) As Long
On Error GoTo ErrSection:

    Dim i&
    
    i = -1
    If m.eType = eANNOT_WaveLabels Then
        m.aDates(m.aDates.Size) = dDate
        m.aYs(m.aYs.Size) = dY
        i = m.aDates.Size + 2
    End If
    
    AddPoint = i

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.AddPoint", eGDRaiseError_Raise

End Function

Public Sub VerifyWavePoints()
On Error GoTo ErrSection:

    Dim i&
    
    If m.aDates.Size = m.aYs.Size Then      'theoretically should always be true
        For i = m.aDates.Size - 1 To 0 Step -1
            If m.aDates(i) < 0 Or m.aYs(i) < 0 Then
                m.aDates.Remove i
                m.aYs.Remove i
            End If
        Next
    End If
    
    AssignDateTime

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.VerifyWavePoints", eGDRaiseError_Raise

End Sub

Public Function CanContinueLabel(Optional ByVal strLabels$ = "") As Boolean
On Error GoTo ErrSection:

    Dim aLabels As New cGdArray
    Dim i&, j&
    
    m.bCanContinueLabel = True
    If Len(strLabels) > 0 Then
        aLabels.SplitFields strLabels
    Else
        aLabels.SplitFields m.strText
    End If
    
    j = aLabels.Size - 1
    
    For i = 0 To j
        If Len(aLabels(i)) > 1 Then
            m.bCanContinueLabel = False
            Exit For
        End If
        If Not IsAlpha(aLabels(i)) And Not IsDigit(aLabels(i)) Then
            m.bCanContinueLabel = False
            Exit For
        End If
    Next
    
    CanContinueLabel = m.bCanContinueLabel

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.CanContinueLabel", eGDRaiseError_Raise

End Function

Private Sub AdjustLabelPastEnd(aLabels As cGdArray, ByVal nArc&)
On Error GoTo ErrSection:

    Dim i&, j&, k&
    Dim nNextNumber&, strChar$
    Dim aOrigLabels As New cGdArray
    
    j = gdGetSize(m.geAnnStruct.glhAnnType)
    
    If Prop("LabelPastEnd") = "repeat" Then
        aOrigLabels.SplitFields m.strText, ","
        k = 0
        For i = aLabels.Size To j
            aLabels(i) = aOrigLabels(k)
            k = k + 1
            If k = aOrigLabels.Size Then k = 0
        Next
    ElseIf Prop("LabelPastEnd") = "continue" Then
        m.bCanContinueLabel = CanContinueLabel()
        If m.bCanContinueLabel Then
            aOrigLabels.SplitFields m.strText, ","
            strChar = Trim(aOrigLabels(aOrigLabels.Size - 1))
            If IsDigit(strChar) Then
                nNextNumber = Val(strChar)
                For i = aLabels.Size To j
                    nNextNumber = nNextNumber + 1
                    aLabels(i) = nNextNumber
                Next
            Else
                For i = aLabels.Size To j
                    k = Asc(strChar) + 1
                    If k = 91 Then
                        k = 65
                    ElseIf k = 123 Then
                        k = 97
                    End If
                    strChar = Chr(k)
                    aLabels(i) = strChar
                Next
            End If
        End If
    End If
    
    If nArc = 1 Then
        For i = aLabels.Size To j
            aLabels(i) = " "       'pad with space so arcs will be drawn
        Next
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.AdjustLabelPastEnd", eGDRaiseError_Raise

End Sub

Public Sub CopyCreateOptions(AnnotSrc As cAnnotation)
On Error GoTo ErrSection:

    If AnnotSrc Is Nothing Then Exit Sub

    If m.eType = eANNOT_WaveLabels Then
        m.strText = AnnotSrc.Text
        m.nColor = AnnotSrc.Color
        m.nStyle = AnnotSrc.Style
        Prop("Lines") = AnnotSrc.Prop("Lines")
        Prop("LabelFirstPoint") = AnnotSrc.Prop("LabelFirstPoint")
        Prop("FontName") = AnnotSrc.Prop("FontName")
        Prop("FontSize") = AnnotSrc.Prop("FontSize")
        Prop("FontUnderline") = AnnotSrc.Prop("FontUnderline")
        Prop("FontStyle") = AnnotSrc.Prop("FontStyle")
        Prop("PointArc") = AnnotSrc.Prop("PointArc")
        Prop("LabelPastEnd") = AnnotSrc.Prop("LabelPastEnd")
    ElseIf m.eType = eANNOT_Icon Or m.eType = eANNOT_ElliotLabel Then
        m.strText = AnnotSrc.Text
        m.nColor = AnnotSrc.Color
        m.nStyle = AnnotSrc.Style
        m.nPreIndicator = AnnotSrc.PreIndicator
        m.bMultiChart = AnnotSrc.MultiChartFlag
        geTextAlign = AnnotSrc.geTextAlign
        Prop("ImageType") = AnnotSrc.Prop("ImageType")
        Prop("ImageSize") = AnnotSrc.Prop("ImageSize")
        Prop("ImageStyle") = AnnotSrc.Prop("ImageStyle")
        Prop("ImageDir") = AnnotSrc.Prop("ImageDir")
        Prop("FontName") = AnnotSrc.Prop("FontName")
        Prop("FontSize") = AnnotSrc.Prop("FontSize")
        Prop("FontStyle") = AnnotSrc.Prop("FontStyle")
        Prop("FontUnderline") = AnnotSrc.Prop("FontUnderline")
        Prop("TextAlignment") = AnnotSrc.Prop("TextAlignment")
        Prop("Custom") = AnnotSrc.Prop("Custom")
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.CopyCreateOptions", eGDRaiseError_Raise

End Sub

Public Property Get AnnotChart() As cChart
    Set AnnotChart = m.Chart
End Property

Private Sub InitAlert()
On Error GoTo ErrSection:
        
    If Not m.Chart Is Nothing And CanHaveAlert() Then
        Set m.oAlert = New cAlert
        If Not m.oAlert Is Nothing Then
            m.oAlert.AlertType = eGDAlertType_Annot
            m.oAlert.Active = True
            If m.eType = eANNOT_HorzLine Or m.eType = eANNOT_HorzLine2 Or _
               m.eType = eANNOT_HorzLine3 Or m.eType = eANNOT_HorzLine4 Or _
               m.eType = eANNOT_SRLine Or m.eType = eANNOT_SRLine2 Or _
               m.eType = eANNOT_SRLine3 Or m.eType = eANNOT_SRLine4 Then
                SetAlertConditionHz
            ElseIf m.eType = eANNOT_Trendline Or m.eType = eANNOT_Trendline2 Or _
                   m.eType = eANNOT_Trendline3 Or m.eType = eANNOT_Trendline4 Or _
                   m.eType = eANNOT_TrendChannel Then
                SetAlertConditionTrendLine True
            End If
        End If
    Else
        UpdateAlert 0
        Set m.oAlert = Nothing
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.InitAlert"
    
End Sub

Public Property Get AlertObject(Optional ByVal bInit As Boolean = False) As cAlert
On Error GoTo ErrSection:
        
    If bInit Then InitAlert
    Set AlertObject = m.oAlert
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.AlertObject.Get"
    
End Property

Public Property Let AlertObject(Optional ByVal bInit As Boolean = False, AlertIn As cAlert)
On Error GoTo ErrSection:

    If AlertIn Is Nothing Then
        Set m.oAlert = Nothing
    ElseIf AlertIn.AlertType = eGDAlertType_Price Then
        Set m.oAlert = AlertIn
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.AlertObject.Let"
    
End Property

Private Sub PriceAlertProp(AlertIn As cAlert, ByVal bSetProperties As Boolean)

    If Not AlertIn Is Nothing Then
        'this is special case only for price alerts to show bell on chart
        If AlertIn.AlertType = eGDAlertType_Price Then
            If m.oAlert Is Nothing Then Set m.oAlert = AlertIn
            If bSetProperties Then
                m.eUsage = eANNOT_PriceAlert
                Prop("ImageSize") = 16
                m.bMultiChart = False
                m.strText = ""
                    
                If AlertIn.Active Then
                    Prop("ImageType") = eCNI_Bell_Price
                Else
                    Prop("ImageType") = eCNI_Bell_Price_Gray
                End If
                
                If AlertIn.UseGetsUpTo Then
                    Prop("ImageDir") = eCNI_North
                    
                    m.Y1 = AlertIn.GetsUpToPrice
                    If AlertIn.UseGetsDownTo Then
                        m.Y2 = AlertIn.GetsDownToPrice
                        'showAxes original set to 1 to fix issue 4706, but resulted 5556 bug when
                        'bid/ask on chart implemented setting showAxes to 2 is fix for 5556
                        m.geAnnStruct.showAxes = 2
                    Else
                        m.Y2 = AlertIn.GetsUpToPrice
                        m.geAnnStruct.showAxes = 0
                    End If
                ElseIf AlertIn.UseGetsDownTo Then
                    Prop("ImageDir") = eCNI_South
                    m.Y1 = AlertIn.GetsDownToPrice
                    m.Y2 = AlertIn.GetsDownToPrice
                End If
                
                gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
                gdSetNum m.geAnnStruct.gdhBottom, 1, m.Y2
                
                If Not m.Chart Is Nothing Then
                    m.Chart.Annots.Add Me, AlertIn.AlertKey
                    Prop("AnnotKey") = AlertIn.AlertKey
                    AlertIn.CheckAlert
                End If
            End If
        End If
    End If

End Sub

Public Property Get CanHaveAlert() As Boolean
On Error GoTo ErrSection:
    
    Dim bCanHave As Boolean
    Dim i&, nBar&
    
    If ExtremeCharts = 1 Then
        Exit Sub        'return false (cannot have annotation alerts)
    End If
    
    If Not m.Chart Is Nothing Then
        If m.eType = eANNOT_Trendline Or m.eType = eANNOT_Trendline2 Or _
           m.eType = eANNOT_Trendline3 Or m.eType = eANNOT_Trendline4 Or _
           m.eType = eANNOT_TrendChannel Then
           
            i = Val(Prop("Ext"))     '1=right, 2=left, 3=both
            If i = 1 Or i = 3 Then
                If m.oAlert Is Nothing Then
                    'last good data bar must be visible on chart trend line value for new alert
                    nBar = m.Chart.LastGoodDataBar(False)
                    i = gePointNum(m.Chart, m.Chart.Bars(eBARS_DateTime, nBar))
                    If i <= m.Chart.geChartPoints Then
                        bCanHave = True
                    End If
                Else
                    'alert already exist allow editing
                    bCanHave = True
                End If
            End If
        ElseIf m.geAnnStruct.paneId = m.Chart.Tree("PRICE PANE").gePaneId Then
            If m.eType = eANNOT_HorzLine Or m.eType = eANNOT_HorzLine2 Or _
               m.eType = eANNOT_HorzLine3 Or m.eType = eANNOT_HorzLine4 Then
                bCanHave = True
            ElseIf m.eType = eANNOT_SRLine Or m.eType = eANNOT_SRLine2 Or _
                   m.eType = eANNOT_SRLine3 Or m.eType = eANNOT_SRLine4 Then
                If Prop("Ext") = 1 Then
                    bCanHave = True
                End If
            End If
        End If
    End If
    
    CanHaveAlert = bCanHave
    
    Exit Property
    
ErrSection:
    RaiseError "cAnnotation.CanHaveAlert.Get"

End Property

Public Property Get AlertAdded() As Boolean
    AlertAdded = m.bAlertAdded
End Property

Private Sub SetAlertConditionHz(Optional ByVal eBarFlag As eGDAlertBarOption = eGDAlertBar9_Undefined)
On Error GoTo ErrSection:

    Dim dHigh#, dLow#, dClose#, i&
    Dim strSave$, strText$
    Dim strAnnotName$, strDirection$
    
    Dim eNewFlag As eGDAlertBarOption
    Dim eOldFlag As eGDAlertBarOption

'Design Note:
'   if price < horz value then check for high >= horz value
'   if price > horz value then check for low <= horz value
'   if price = horz value then ?
    
    If m.Chart Is Nothing Then Exit Sub
    
    eOldFlag = m.oAlert.AlertBarFlag
    If eBarFlag = eGDAlertBar9_Undefined Then
        eNewFlag = eOldFlag
    Else
        eNewFlag = eBarFlag
    End If
    
    strSave = m.oAlert.EnglishString
    
    With m.Chart
        dHigh = .Bars(eBARS_High, .LastGoodDataBar(False))
        dLow = .Bars(eBARS_Low, .LastGoodDataBar(False))
        dClose = .Bars(eBARS_Close, .LastGoodDataBar(False))
                
        strText = Trim(m.oAlert.ChartCondition)
        
        If m.eType = eANNOT_SRLine Or m.eType = eANNOT_SRLine2 Or m.eType = eANNOT_SRLine3 Or m.eType = eANNOT_SRLine4 Then
            strAnnotName = "Support Resistance Line  "
        ElseIf Prop("FakePriceAlert") = 1 Then
            strAnnotName = "Spread "
        Else
            strAnnotName = "Horizontal Line  "
        End If
        
        If Len(strText) = 0 Then
            'set alert parameters
            If dClose <= m.Y1 Then
                m.oAlert.Operator = ">="
                m.oAlert.Field = "High"
                strDirection = "(Price gets up to "
            ElseIf dClose > m.Y1 Then
                m.oAlert.Operator = "<="
                m.oAlert.Field = "Low"
                strDirection = "(Price gets down to "
            End If
        ElseIf InStr(strText, "down to") <> 0 Then
            strDirection = "(Price gets down to "
        Else
            strDirection = "(Price gets up to "
        End If
        m.oAlert.ChartCondition = strAnnotName & strDirection & .Bars.PriceDisplay(m.Y1) & ")"
        m.oAlert.Symbol = .Symbol
        m.oAlert.SymbolID = .SymbolID
        m.oAlert.Value = m.Y1
        m.oAlert.Annotation = Me
        m.oAlert.ResetUseCloseOfBar , , eNewFlag
        
        'add extra junk to saved string to trigger reset of last checked time
        If eNewFlag <> eOldFlag Then strSave = "Flag Changed " & strSave
        m.oAlert.ResetLastChecked strSave
        
        If m.oAlert.ChartAlertId = 0 Then AlertsCollection 1
        m.strAlert = m.oAlert.ToFileString
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.SetAlertConditionHz"

End Sub


Private Sub SetAlertConditionTrendLine(ByVal bSetOperators As Boolean, _
    Optional ByVal eBarFlag As eGDAlertBarOption = eGDAlertBar9_Undefined)
On Error GoTo ErrSection:

    Dim Ind As cIndicator
    Dim dData#, dDate#, dY#, dMinMove#
    Dim i&, j&, hArray&, strText$, strSave$, strOld$
    
    Dim bContinue As Boolean
    Dim eNewFlag As eGDAlertBarOption
    Dim eOldFlag As eGDAlertBarOption

    If m.Chart Is Nothing Then Exit Sub
        
    eOldFlag = m.oAlert.AlertBarFlag
    If eBarFlag = eGDAlertBar9_Undefined Then
        eNewFlag = eOldFlag
    Else
        eNewFlag = eBarFlag
    End If
    
    strText = Trim(m.oAlert.ChartCondition)
    If Len(strText) > 0 Then
        If bSetOperators Then
            strText = Parse(m.oAlert.ChartCondition, " gets ", 1)
        Else
            strText = Parse(m.oAlert.ChartCondition, " to ", 1) & " to "
            m.oAlert.ChartCondition = strText & m.Chart.Bars.PriceDisplay(m.oAlert.Value) & ")"
            m.strAlert = m.oAlert.ToFileString
            Exit Sub
        End If
    End If
    
    strOld = Trim(strText)
    
    'check that indicator has not been deleted
    strText = Prop("IndicatorKey")
    If Len(strText) > 0 Then
        Set Ind = m.Chart.Tree(Prop("IndicatorKey"))
        If Ind Is Nothing Then
            UpdateAlert 0
            Exit Sub
        End If
    End If

    If Ind Is Nothing Then
        If m.geAnnStruct.paneId = m.Chart.Tree("PRICE PANE").gePaneId Then
            Set Ind = m.Chart.Tree("PRICE")
        Else
            'find first non-boolean indicator in pane
            Set Ind = m.Chart.Tree.RelativeItem(m.geAnnStruct.paneId, eTREE_FirstChild)
            bContinue = True
            While bContinue
                If Ind Is Nothing Then
                    bContinue = False
                ElseIf Ind.DataType = eINDIC_BooleanArray Then
                    Set Ind = m.Chart.Tree.RelativeItem(Ind.geIndId, eTREE_NextSibling)
                Else
                    bContinue = False
                End If
            Wend
        End If
    End If
    If Ind Is Nothing Then Exit Sub
    
    strSave = m.oAlert.EnglishString
    strText = "Trend Line "
    Prop("IndicatorKey") = m.Chart.Tree.Key(Ind.geIndId)

    'PriceField: 0=Close,1=Open,2=High,3=Low
    If Ind.isPriceInd Then
        i = Val(Prop("PriceField"))
        Select Case i
            Case 1
                strText = strText & "(Open"
                dData = m.Chart.Bars(eBARS_Open, m.Chart.LastGoodDataBar(False))
            Case 2
                strText = strText & "(High"
                dData = m.Chart.Bars(eBARS_High, m.Chart.LastGoodDataBar(False))
            Case 3
                strText = strText & "(Low"
                dData = m.Chart.Bars(eBARS_Low, m.Chart.LastGoodDataBar(False))
            Case Else
                strText = strText & "(Close"
                dData = m.Chart.Bars(eBARS_Close, m.Chart.LastGoodDataBar(False))
        End Select
    Else
        strText = strText & "(" & Ind.ChartLabel
        If Ind.DataType = eINDIC_Array Then
            hArray = Ind.Data.ArrayHandle
            j = Ind.Data.Size - 1
            For i = j To 0 Step -1
                dData = gdGetNum(hArray, i)
                If dData <> kNullData Then Exit For
            Next
        End If
    End If
    
    
    With m.Chart
        dDate = .Bars(eBARS_DateTime, .LastGoodDataBar(False))
        'fixes trend value not snapping to price like in chart's form
        If dDate = m.dDate1 Then
            dY = m.Y1
        ElseIf dDate = m.dDate2 Then
            dY = m.Y2
        Else
            i = gePointNum(m.Chart, dDate)
            dY = geTrendValueY(m.geAnnStruct.hObj, i)
        End If
        
        If dY > 0 Then
            If Trim(strText) = strOld Then
                strText = Parse(m.oAlert.ChartCondition, " to ", 1) & " to "
                m.oAlert.ChartCondition = strText & m.Chart.Bars.PriceDisplay(dY) & ")"
            Else
                dMinMove = m.Chart.Bars.MinMove(m.Chart.Bars(eBARS_DateTime, m.Chart.Bars.Size - 1))
                dData = RoundToMinMove(dData, dMinMove)
                dY = RoundToMinMove(dY, dMinMove)
                'set alert parameters
                If dData <= dY Then
                    m.oAlert.Operator = ">="
                    m.oAlert.Field = "High"
                    strText = strText & " gets up to "
                ElseIf dData > dY Then
                    m.oAlert.Operator = "<="
                    m.oAlert.Field = "Low"
                    strText = strText & " gets down to "
                End If
                m.oAlert.ChartCondition = strText & .Bars.PriceDisplay(dY) & ")"
            End If
            m.oAlert.Symbol = .Symbol
            m.oAlert.SymbolID = .SymbolID
            m.oAlert.Value = dY
            m.oAlert.Annotation = Me
            m.oAlert.ResetUseCloseOfBar , , eNewFlag
            
            'add extra junk to saved string to trigger reset of last checked time
            If eNewFlag <> eOldFlag Then strSave = "Flag Changed " & strSave
            m.oAlert.ResetLastChecked strSave
            
            If m.oAlert.ChartAlertId = 0 Then AlertsCollection 1
            m.strAlert = m.oAlert.ToFileString
        End If
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.SetAlertConditionTrendLine"

End Sub

Private Function SetPivotPoints(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim strFontName$, iFontSize&, iFontStyle&, iFontUnderline&
    Dim iShowText&, iExt&, iPriceOnly&
    Dim i&, j&

    If dX1 = 0 And dX2 = 0 Then Exit Function       'aardvark 1806

    'This needs to be done only on initial add to grapheng.dll
    'or when the annotation type is set to -1 to hide it
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, -1        'reserved for extension pen info
        gdSetNum m.geAnnStruct.glhAnnType, 1, 34        'PP
        gdSetNum m.geAnnStruct.glhAnnType, 2, 34        'S1
        gdSetNum m.geAnnStruct.glhAnnType, 3, 34        'S2
        gdSetNum m.geAnnStruct.glhAnnType, 4, 34        'S3
        gdSetNum m.geAnnStruct.glhAnnType, 5, 34        'R1
        gdSetNum m.geAnnStruct.glhAnnType, 6, 34        'R2
        gdSetNum m.geAnnStruct.glhAnnType, 7, 34        'R3
        'place holder for extension, ignored by grapheng.dll
        gdSetNum m.geAnnStruct.gdhTop, 0, 0
        gdSetNum m.geAnnStruct.gdhBottom, 0, 0
        gdSetNum m.geAnnStruct.gdhLeft, 0, 0
        gdSetNum m.geAnnStruct.gdhRight, 0, 0
        gdSetNum m.geAnnStruct.glhPenColor, 0, 0
        gdSetStr m.geAnnStruct.gshText, 0, ""
    End If
        
    CalcPivots
    
    i = Val(Prop("PivotShow"))
    If i <> 34 Then i = -1
    gdSetNum m.geAnnStruct.glhAnnType, 1, i
    
    i = Val(Prop("S1Show"))
    If i <> 34 Then i = -1
    gdSetNum m.geAnnStruct.glhAnnType, 2, i
    
    i = Val(Prop("S2Show"))
    If i <> 34 Then i = -1
    gdSetNum m.geAnnStruct.glhAnnType, 3, i
    
    i = Val(Prop("S3Show"))
    If i <> 34 Then i = -1
    gdSetNum m.geAnnStruct.glhAnnType, 4, i
    
    i = Val(Prop("R1Show"))
    If i <> 34 Then i = -1
    gdSetNum m.geAnnStruct.glhAnnType, 5, i
    
    i = Val(Prop("R2Show"))
    If i <> 34 Then i = -1
    gdSetNum m.geAnnStruct.glhAnnType, 6, i
    
    i = Val(Prop("R3Show"))
    If i <> 34 Then i = -1
    gdSetNum m.geAnnStruct.glhAnnType, 7, i
    
    'Show values property can be 0/1/2/3
    '   0: no values or labels: S1,R1 ...
    '   1: values on right, labesl on left
    '   2: values & labels on left
    '   3: values & labels on right
    iShowText = Val(Prop("ShowValues"))
    'flag to draw text at end of extensions or next to main lines
    m.geAnnStruct.showQtrLines = Val(Prop("TextNextToMain"))
    
    iExt = Val(Prop("Ext"))
    
    iFontSize = Val(Prop("FontSize"))
    iFontStyle = Val(Prop("FontStyle"))
    iFontUnderline = Val(Prop("FontUnderline"))
    strFontName = Prop("FontName")
    
    For i = 0 To 7
        gdSetNum m.geAnnStruct.gdhMisc, i, iExt                'extension: 0=none,1=right,2=left,3=both
        gdSetStr m.geAnnStruct.gshFont, i, strFontName
        gdSetNum m.geAnnStruct.glhSize, i, iFontSize
        gdSetNum m.geAnnStruct.glhStyle, i, iFontStyle          '5749
        gdSetNum m.geAnnStruct.glhUnderline, i, iFontUnderline
        
        geSetPenStyle m.nStyle, i
        If iShowText = 0 Then
            gdSetStr m.geAnnStruct.gshText, i, ""
        ElseIf iShowText = 1 Then
            gdSetNum m.geAnnStruct.glhAlign, i, 9
        ElseIf iShowText = 2 Then
            gdSetNum m.geAnnStruct.glhAlign, i, 7
        Else
            gdSetNum m.geAnnStruct.glhAlign, i, 6
        End If
    Next
    
    iPriceOnly = Val(Prop("PriceOnly"))
        
    'pen style in array item [0] is for extension
    geSetPenStyle Val(Prop("ExtStyle")), 0
    
    'lines
    gdSetNum m.geAnnStruct.gdhTop, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhBottom, 1, m.Y1
    gdSetNum m.geAnnStruct.gdhTop, 2, m.Y2
    gdSetNum m.geAnnStruct.gdhBottom, 2, m.Y2
    
    'aYs[0/S2] [1/S3] [2/R1] [3/R2] [4/R3]
    For i = 0 To 4
        gdSetNum m.geAnnStruct.gdhTop, i + 3, m.aYs(i)
        gdSetNum m.geAnnStruct.gdhBottom, i + 3, m.aYs(i)
    Next
    
    
    'flag for whether and what kind of text to draw:
    '   -1: no text (use -1 for backwards compatibility, previous versions did not have this flag and was defaulted to 0)
    '    1: price only
    '    2: text & ratio
    If iShowText = 0 Then
        m.geAnnStruct.minorAxisLenData = -1
    ElseIf iPriceOnly = 1 Then
        m.geAnnStruct.minorAxisLenData = 1
    Else
        m.geAnnStruct.minorAxisLenData = 2
    End If
    
    'text
    If iShowText <> 0 Then
        If m.dDate1 < m.dDate2 Then
            If iPriceOnly = 1 Then
                gdSetStr m.geAnnStruct.gshText, 1, Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
                gdSetStr m.geAnnStruct.gshText, 2, Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y2)
            Else
                gdSetStr m.geAnnStruct.gshText, 1, "PP" & " | " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
                gdSetStr m.geAnnStruct.gshText, 2, "S1" & " | " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y2)
            End If
            gdSetNum m.geAnnStruct.glhPenColor, 1, Val(Prop("PivotColor"))
            gdSetNum m.geAnnStruct.glhPenColor, 2, Val(Prop("S1Color"))
        Else
            If iPriceOnly Then
                gdSetStr m.geAnnStruct.gshText, 1, Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
                gdSetStr m.geAnnStruct.gshText, 2, Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y2)
            Else
                gdSetStr m.geAnnStruct.gshText, 1, "S1" & " | " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y1)
                gdSetStr m.geAnnStruct.gshText, 2, "PP" & " | " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.Y2)
            End If
            gdSetNum m.geAnnStruct.glhPenColor, 1, Val(Prop("S1Color"))
            gdSetNum m.geAnnStruct.glhPenColor, 2, Val(Prop("PivotColor"))
        End If
        If iPriceOnly = 1 Then
            gdSetStr m.geAnnStruct.gshText, 3, Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(0))
            gdSetStr m.geAnnStruct.gshText, 4, Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(1))
            gdSetStr m.geAnnStruct.gshText, 5, Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(2))
            gdSetStr m.geAnnStruct.gshText, 6, Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(3))
            gdSetStr m.geAnnStruct.gshText, 7, Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(4))
        Else
            gdSetStr m.geAnnStruct.gshText, 3, "S2" & " | " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(0))
            gdSetStr m.geAnnStruct.gshText, 4, "S3" & " | " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(1))
            gdSetStr m.geAnnStruct.gshText, 5, "R1" & " | " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(2))
            gdSetStr m.geAnnStruct.gshText, 6, "R2" & " | " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(3))
            gdSetStr m.geAnnStruct.gshText, 7, "R3" & " | " & Chart.PriceDisplay(m.geAnnStruct.paneId, m.aYs(4))
        End If
    End If
    
    'pen colors
    gdSetNum m.geAnnStruct.glhPenColor, 3, Val(Prop("S2Color"))
    gdSetNum m.geAnnStruct.glhPenColor, 4, Val(Prop("S3Color"))
    gdSetNum m.geAnnStruct.glhPenColor, 5, Val(Prop("R1Color"))
    gdSetNum m.geAnnStruct.glhPenColor, 6, Val(Prop("R2Color"))
    gdSetNum m.geAnnStruct.glhPenColor, 7, Val(Prop("R3Color"))
        
    For i = 1 To 7
        If dX2 < dX1 Then
            If dX2 < 0 Then          'out of range (user changed to lower bar period - issue 3218)
                If m.dDate2 > m.dDate1 Then
                    gdSetNum m.geAnnStruct.gdhLeft, i, dX1
                    gdSetNum m.geAnnStruct.gdhRight, i, Chart.geChartPoints - 1
                Else
                    gdSetNum m.geAnnStruct.gdhLeft, i, 0
                    gdSetNum m.geAnnStruct.gdhRight, i, dX1
                End If
            Else
                gdSetNum m.geAnnStruct.gdhLeft, i, dX2
                gdSetNum m.geAnnStruct.gdhRight, i, dX1
            End If
        Else
            gdSetNum m.geAnnStruct.gdhLeft, i, dX1
            gdSetNum m.geAnnStruct.gdhRight, i, dX2
        End If
    Next
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetPivotPoints", eGDRaiseError_Raise

End Function

Public Sub UpdateAlert(ByVal nAddRemove&, Optional ByVal bSetCondition As Boolean = False, _
    Optional ByVal eBarFlag As eGDAlertBarOption = eGDAlertBar9_Undefined, _
    Optional ByVal bTerminating As Boolean = False)
On Error GoTo ErrSection:
'nAddRemove: 0=remove, 1=add, 2=just update alert string, 3=special case for caching chart object
    
    Dim strSave$, i&
    
    If m.oAlert Is Nothing Then Exit Sub
        
    If m.oAlert.AlertType = eGDAlertType_Price Then
        'this is not a chart drawing tool
        'alert functionalities are handled by the global alerts collection object
        'the purpose of this update is just to update draw information for bell on chart
        If bTerminating Or nAddRemove = 3 Then
            m.oAlert.RemoveChartObject m.Chart
        ElseIf nAddRemove = 2 Then
            PriceAlertProp m.oAlert, bSetCondition
        End If
        GoTo ErrExit
    ElseIf nAddRemove = 3 Then
        m.strAlert = m.oAlert.ToFileString          '6451
        For i = 1 To g.Alerts.Count
            If g.Alerts(i) Is m.oAlert Then
                If g.Alerts(i).ChartAlertId = m.oAlert.ChartAlertId Then
                    g.Alerts.Remove (i)
                    Exit For
                End If
            End If
        Next
    ElseIf nAddRemove = 0 Then
        m.strAlert = ""
        Prop("IndicatorKey") = ""
        Prop("PriceField") = ""
        If m.eType = eANNOT_HorzLine And Prop("FakePriceAlert") = 1 Then
            If Not m.Chart Is Nothing Then
                If m.Chart.Annots(Prop("AnnotKey")) Is Me Then      'precautionary
                    m.Chart.Annots.Remove Prop("AnnotKey")
                End If
            End If
        End If
    ElseIf bSetCondition Then
        'this is true when called from chart form after user moved annot
        'or from annot editor when user changes value of horz or SR line
        'or from alerts form when user edits condition of alert
        If m.eType = eANNOT_Trendline Or m.eType = eANNOT_Trendline2 Or _
           m.eType = eANNOT_Trendline3 Or m.eType = eANNOT_Trendline4 Or _
           m.eType = eANNOT_TrendChannel Then
           
            SetAlertConditionTrendLine True, eBarFlag
        Else
            SetAlertConditionHz eBarFlag
        End If
    End If
    
    If nAddRemove < 2 Then AlertsCollection nAddRemove
        
    If FormIsLoaded("frmAlertsSetup") Then
        If frmAlertsSetup.Visible Then frmAlertsSetup.LoadGrid
    End If
    
    CheckAnnotAlert True
    
    If Not m.Chart Is Nothing Then
        geDrawAnn m.Chart
        If m.bMultiChart Then
            m.Chart.SyncGlobalAnnots Me
        Else
            'set focus to chart so Ctrl-Z will work right away
            SetFocus m.Chart.Form.pbChart.hWnd
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.UpdateAlert"

End Sub

Private Function CheckHorzAlert(ByVal dLow#, ByVal dHigh#, ByVal dClose#, ByVal dMinMove#) As Boolean
On Error GoTo ErrSection:
   
    Dim strOp$, strField$, strErr$
    Dim dY#, dHzVal#
    Dim bTriggerAlert As Boolean
    
    With m.oAlert
        strOp = .Operator
        strField = .Field
    End With
        
    If m.oAlert.AlertBarFlag = eGDAlertBar3_CloseThisBar Or m.oAlert.AlertBarFlag = eGDAlertBar5_CloseCompleteBar Then
        dY = dClose
    ElseIf strField = "High" Then
        dY = dHigh
    ElseIf strField = "Low" Then
        dY = dLow
    ElseIf Not g.bStarting Then
        If m.eType = eANNOT_SRLine Or m.eType = eANNOT_SRLine2 Or m.eType = eANNOT_SRLine3 Or m.eType = eANNOT_SRLine4 Then
            strErr = " support resistance line alert."
        Else
            strErr = " horizontal line alert. "
        End If
        strErr = "Error in" & strErr & vbCrLf & "field = " & strField & ". Removing alert."
        InfBox strErr, "I", , "Check Alert"
        UpdateAlert 0
        GoTo ErrExit
    Else
        GoTo ErrExit
    End If
        
    dY = RoundToMinMove(dY, dMinMove)
    dHzVal = RoundToMinMove(m.Y1, dMinMove)
    
    If InStr(strOp, "=") > 0 And dY = dHzVal Then
        bTriggerAlert = True
    ElseIf InStr(strOp, ">") > 0 And dY > dHzVal Then
        bTriggerAlert = True
    ElseIf InStr(strOp, "<") > 0 And dY < dHzVal Then
        bTriggerAlert = True
    End If
    
    CheckHorzAlert = bTriggerAlert
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.CheckHorzAlert"

End Function

Private Function CheckTrendAlert(ByVal dLow#, ByVal dHigh#, _
    ByVal dClose#, ByVal dMinMove#, ByVal nPtNum&) As Boolean
On Error GoTo ErrSection:

    Dim strOp$, strField$, strText$, strTemp$
    Dim dY#, dTrend#, dDate#, i&
    Dim bTriggerAlert As Boolean
    
    Dim strIndicatorKey$
    Dim Ind As cIndicator
    
    With m.oAlert
        strOp = .Operator
        strField = .Field
    End With
        
    strIndicatorKey = Prop("IndicatorKey")
    If Len(strIndicatorKey) = 0 Or strIndicatorKey = "PRICE" Then
        'alert is set on price
        If m.oAlert.AlertBarFlag = eGDAlertBar3_CloseThisBar Then
            dY = dClose
        ElseIf strField = "High" Then
            dY = dHigh
        ElseIf strField = "Low" Then
            dY = dLow
        Else
            dY = -1
        End If
    ElseIf m.Chart Is Nothing Then
        Exit Function       'no chart object, cannot locate indicator
    Else
        Set Ind = m.Chart.Tree(strIndicatorKey)
        If Ind Is Nothing Then
            dY = -1
        Else
            i = m.Chart.LastGoodDataBar(False)
            dY = Ind.Data(i)
        End If
        Set Ind = Nothing
    End If
    
    're-calculate trendline value only if last good data bar is on screen
    If nPtNum <= m.Chart.geChartPoints Then
        dDate = m.Chart.Bars(eBARS_DateTime, nPtNum)
        If dDate = m.dDate1 Then
            dTrend = m.Y1
        ElseIf dDate = m.dDate2 Then
            dTrend = m.Y2
        Else
            dTrend = geTrendValueY(m.geAnnStruct.hObj, nPtNum)
        End If
    Else
        dTrend = m.oAlert.Value
    End If
        
    dY = RoundToMinMove(dY, dMinMove)
    dTrend = RoundToMinMove(dTrend, dMinMove)
    
    If m.oAlert.Value <> dTrend Then
        m.oAlert.Value = dTrend
        'change text for alert
        SetAlertConditionTrendLine False
    End If
    
    If dY > 0 Then
        If InStr(strOp, "=") > 0 And dY = dTrend Then
            bTriggerAlert = True
        ElseIf InStr(strOp, ">") > 0 And dY > dTrend Then
            bTriggerAlert = True
        ElseIf InStr(strOp, "<") > 0 And dY < dTrend Then
            bTriggerAlert = True
        End If
    End If
            
    CheckTrendAlert = bTriggerAlert
    
    Exit Function

ErrSection:
    RaiseError "cAnnotation.CheckTrendAlert"

End Function

Public Sub CheckAnnotAlert(ByVal bConfirmOverride As Boolean)
On Error GoTo ErrSection:

    Static nPrevPeriodicity As Long
    Dim nBar&, nLastDataBar&, nPtNum&
    Dim dLow#, dHigh#, dClose#, dY#, dMinMove#
    
    Dim eBarFlag As eGDAlertBarOption
    Dim Pane As cPane
    Dim PaneWood As cPaneWood
    
    If m.oAlert Is Nothing Then Exit Sub
    If m.oAlert.AlertType = eGDAlertType_Price Then Exit Sub
    
    If Not m.Chart Is Nothing Then
        nLastDataBar = m.Chart.LastGoodDataBar(False)
        nBar = nLastDataBar
        eBarFlag = m.oAlert.AlertBarFlag
        
        If nPrevPeriodicity <> 0 Then
            If Not bConfirmOverride Then        '4381
                If nPrevPeriodicity <> m.Chart.Bars.Prop(eBARS_Periodicity) Then bConfirmOverride = True
            End If
        End If
        
        nPrevPeriodicity = m.Chart.Bars.Prop(eBARS_Periodicity)
        
        If eBarFlag = eGDAlertBar4_HiLowCompleteBar Or eBarFlag = eGDAlertBar5_CloseCompleteBar Then
            If m.Chart.ShowSplitPane = 0 Or Not g.RealTime.Active Then
                nBar = nBar - 1
            Else
                Set Pane = m.Chart.Tree("PRICE PANE")
                If Not Pane Is Nothing Then Set PaneWood = Pane.WoodPane
                If PaneWood Is Nothing Then
                    nBar = nBar - 1
                ElseIf PaneWood.LastTimerVal > 0 Then
                    nBar = nBar - 1
                Else
                    nBar = nBar
                End If
            End If
        End If
        
        dLow = m.Chart.Bars(eBARS_Low, nBar)
        dHigh = m.Chart.Bars(eBARS_High, nBar)
        dClose = m.Chart.Bars(eBARS_Close, nBar)
        
        dMinMove = m.Chart.Bars.MinMove(m.Chart.Bars(eBARS_DateTime, nPtNum))
        If m.oAlert.Active Then
            If m.eType = eANNOT_HorzLine Or m.eType = eANNOT_HorzLine2 Or _
               m.eType = eANNOT_HorzLine3 Or m.eType = eANNOT_HorzLine4 Or _
               m.eType = eANNOT_SRLine Or m.eType = eANNOT_SRLine2 Or _
               m.eType = eANNOT_SRLine3 Or m.eType = eANNOT_SRLine4 Then
                If CheckHorzAlert(dLow, dHigh, dClose, dMinMove) Then
                    m.oAlert.CheckAlert , , True, bConfirmOverride
                ElseIf Not m.oAlert Is Nothing Then
                    'alert might have gotten removed if an error was encountered
                     m.oAlert.CheckAlert , , False
                End If
            ElseIf m.eType = eANNOT_Trendline Or m.eType = eANNOT_Trendline2 Or _
                   m.eType = eANNOT_Trendline3 Or m.eType = eANNOT_Trendline4 Then
                   
                nPtNum = gePointNum(m.Chart, m.Chart.Bars(eBARS_DateTime, nLastDataBar))
                If CheckTrendAlert(dLow, dHigh, dClose, dMinMove, nPtNum) Then
                    m.oAlert.CheckAlert , , True, bConfirmOverride
                Else
                    m.oAlert.CheckAlert , , False
                End If
            End If
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.CheckAnnotAlert"

End Sub

Private Sub AlertsCollection(ByVal nAddRemove&)
On Error GoTo ErrSection:
    
    Dim i&
    Dim bFound As Boolean, bAdded As Boolean
    Dim Alert As cAlert
    
    bAdded = m.bAlertAdded
    If Not m.oAlert Is Nothing And Not g.Alerts Is Nothing Then
        'check for existing alert in collection
        If m.oAlert.ChartAlertId <> 0 Then      'fix for aardvark 3204
            For i = 1 To g.Alerts.Count
                Set Alert = g.Alerts(i)
                If Alert.ChartAlertId = m.oAlert.ChartAlertId Then
                    bFound = True
                    Exit For
                End If
            Next
        End If
        If nAddRemove = 0 Then
            If bFound Then g.Alerts.Remove (i)
            Set m.oAlert = Nothing
            m.bAlertAdded = False
        ElseIf Not bFound Then
            g.Alerts.Add m.oAlert
            m.oAlert.Annotation = Me
            m.oAlert.ChartAlertId = RoundNum(CDbl(Now), 5) + g.Alerts.Count     'aardvark 3080
            m.bAlertAdded = True
        End If
    End If
       
    If FormIsLoaded("frmAlertsSetup") Then
        If frmAlertsSetup.Visible Then frmAlertsSetup.LoadGrid
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.AlertsCollection"
End Sub

Private Sub SetAlertIcon(ByVal nIdx&, ByVal dTop#)
On Error GoTo ErrSection:
    
    Dim eBellImg As eStockImage
    
    If nIdx = 0 Then
        'This needs to be done only on initial add to grapheng.dll
        If m.geAdded = False Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, 7
            gdSetNum m.geAnnStruct.glhImageType, 0, eCNI_Bell_Gray
            gdSetNum m.geAnnStruct.glhSize, 0, 1
            'fields below not used, but need to be set as place holders
            gdSetNum m.geAnnStruct.glhAlign, 0, 0
            gdSetNum m.geAnnStruct.glhPenSize, 0, 0
            gdSetNum m.geAnnStruct.glhPenStyle, 0, 0
            gdSetNum m.geAnnStruct.glhPenColor, 0, 0
            gdSetNum m.geAnnStruct.glhStyle, 0, 0
            gdSetNum m.geAnnStruct.glhFillPattern, 0, 0
        End If
    End If
    
    If m.oAlert Is Nothing Then
        'this is a stand alone icon for condition builder alert
        'want to check the prop to make sure it was set properly
        If Prop("ImageType") = eCNI_Bell Then
            eBellImg = eCNI_Bell
        Else
            eBellImg = eCNI_Bell_Gray
        End If
    ElseIf m.oAlert.Active Then
        If m.eType = eANNOT_HorzLine And Prop("FakePriceAlert") = 1 Then
            eBellImg = eCNI_Bell_Price
        Else
            eBellImg = eCNI_Bell
        End If
    Else
        If m.eType = eANNOT_HorzLine And Prop("FakePriceAlert") = 1 Then
            eBellImg = eCNI_Bell_Price_Gray
        Else
            eBellImg = eCNI_Bell_Gray
        End If
    End If
        
    If nIdx = 0 Then m.geAnnStruct.PreIndicator = 1         'aardvark 3908
    gdSetNum m.geAnnStruct.glhAnnType, nIdx, 7
    gdSetNum m.geAnnStruct.glhSize, nIdx, 1     'need to be >0
    gdSetNum m.geAnnStruct.glhImageType, nIdx, eBellImg
    
    gdSetNum m.geAnnStruct.gdhTop, nIdx, dTop
    gdSetNum m.geAnnStruct.gdhLeft, nIdx, 0
    gdSetNum m.geAnnStruct.gdhBottom, nIdx, 0
    gdSetNum m.geAnnStruct.gdhRight, nIdx, 0
    
    Exit Sub
                
ErrSection:
    RaiseError "cAnnotation.SetAlertIcon"
    
End Sub

Public Sub AlertTip(oToolTip As cToolTip, ByRef vCtrl As Control, ByVal nIdx&)
On Error GoTo ErrSection:

    Dim l&, t&, r&, b&, j&, i&
    Dim oAlert As cAlert
    Dim strText$

    Set oToolTip = Nothing
    If g.ChartGlobals.bChartTips Then
        j = gdGetSize(m.geAnnStruct.glhPixLeft)
        If j > nIdx Then
            l = gdGetNum(m.geAnnStruct.glhPixLeft, nIdx) - 1
            t = gdGetNum(m.geAnnStruct.glhPixTop, nIdx) - 1
            b = gdGetNum(m.geAnnStruct.glhPixBottom, nIdx)
            r = gdGetNum(m.geAnnStruct.glhPixRight, nIdx)
            If l >= 0 And t >= 0 And b >= 0 And r >= 0 Then
                Set oToolTip = New cToolTip
                oToolTip.Create m.Chart.Form
                
                oToolTip.DelayTime(ttDelayInitial) = 0
                oToolTip.DelayTime(ttDelayReshow) = 0
                oToolTip.DelayTime(ttDelayShow) = 3000
                oToolTip.MaxTipWidth = 1
                
                If m.oAlert Is Nothing Then
                    For i = 1 To g.Alerts.Count
                        Set oAlert = g.Alerts(i)
                        'tooltip for chart condition alerts bell should only show alert for that specific chart
                        If oAlert.AlertType = eGDAlertType_Chart And oAlert.Symbol = m.Chart.Symbol Then
                            If Len(strText) = 0 Then
                                strText = Parse(oAlert.ChartCondition, ":", 1)
                            Else
                                strText = strText & vbCrLf & Parse(oAlert.ChartCondition, ":", 1)
                            End If
                        End If
                    Next
                    oToolTip.AddRect vCtrl, l, t, r, b, strText, False
                Else
                    oToolTip.AddRect vCtrl, l, t, r, b, m.oAlert.ChartCondition, False
                End If
            End If
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.AlertTip"
    
End Sub

Public Property Get geGetX(ByVal idx&) As Long
On Error GoTo ErrSection:

    geGetX = gdGetNum(m.geAnnStruct.gdhLeft, idx)
    
    Exit Property

ErrSection:
    RaiseError "cAnnotation.geGetX.Get"

End Property

Public Function IsFibType() As Boolean
On Error GoTo ErrSection:

    If m.eType = eANNOT_ElliotTimeRatio Or m.eType = eANNOT_FibArcs Or m.eType = eANNOT_FibExpansion _
        Or m.eType = eANNOT_FibFan Or m.eType = eANNOT_Fibonacci Or m.eType = eANNOT_Fibonacci2 _
        Or m.eType = eANNOT_Fibonacci3 Or m.eType = eANNOT_Fibonacci4 Or m.eType = eANNOT_FibTimeRatio _
        Or m.eType = eANNOT_DanCodeFib Or m.eType = eANNOT_FibABCD Or m.eType = eANNOT_Gartley _
        Or m.eType = eANNOT_AdvRiskReward Then
        
        IsFibType = True
        
    ElseIf m.eType = eANNOT_AndrewFork And m.bSchiffFork Then
        IsFibType = True
    ElseIf Not UseDiNapFib() Then
        If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or _
           m.eType = eANNOT_DNExpansion3 Or m.eType = eANNOT_DNExpansion4 Then
            IsFibType = True
        End If
    End If

    Exit Function
        
ErrSection:
    RaiseError "cAnnotation.IsFibType"
    
End Function

Public Function AnnotTypeFromToolID(ByVal strToolID$) As eAnnotType
On Error GoTo ErrSection:

    Dim eType As eAnnotType
    
    Select Case strToolID
        Case "ID_DollarLine"
            eType = eANNOT_DollarLine
        Case "ID_DollarLine2"
            eType = eANNOT_DollarLine2
        Case "ID_DollarLine3"
            eType = eANNOT_DollarLine3
        Case "ID_DollarLine4"
            eType = eANNOT_DollarLine4
        Case "ID_Trendline"
            eType = eANNOT_Trendline
        Case "ID_Trendline2"
            eType = eANNOT_Trendline2
        Case "ID_Trendline3"
            eType = eANNOT_Trendline3
        Case "ID_Trendline4"
            eType = eANNOT_Trendline4
        Case "ID_Fibonacci"
            eType = eANNOT_Fibonacci
        Case "ID_Fibonacci2"
            eType = eANNOT_Fibonacci2
        Case "ID_Fibonacci3"
            eType = eANNOT_Fibonacci3
        Case "ID_Fibonacci4"
            eType = eANNOT_Fibonacci4
        Case "ID_DNRetracement"
            eType = eANNOT_DNRetracement
        Case "ID_DNExpansion"
            eType = eANNOT_DNExpansion
        Case "ID_DNExpansion2"
            eType = eANNOT_DNExpansion2
        Case "ID_DNExpansion3"
            eType = eANNOT_DNExpansion3
        Case "ID_DNExpansion4"
            eType = eANNOT_DNExpansion4
        Case "ID_TargetShooter"
            eType = eANNOT_TargetShooter
        Case "ID_Rectangle", "ID_Rectangle_PFP", "ID_FibClusters"
            eType = eANNOT_Rectangle
            If strToolID = "ID_Rectangle_PFP" Then m.eUsage = eANNOT_PatternProfit
        Case "ID_RegressionLine"
            eType = eANNOT_RegressionLine
        Case "ID_HorzLine"
            eType = eANNOT_HorzLine
        Case "ID_HorzLine2"
            eType = eANNOT_HorzLine2
        Case "ID_HorzLine3"
            eType = eANNOT_HorzLine3
        Case "ID_HorzLine4"
            eType = eANNOT_HorzLine4
        Case "ID_VertLine"
            eType = eANNOT_VertLine
        Case "ID_Icon"
            eType = eANNOT_Icon
        Case "ID_ElliotLabels", "ID_ElliotEndUser"
            eType = eANNOT_ElliotLabel
        Case "ID_Text"
            eType = eANNOT_TextEdit
        Case "ID_Text2"
            eType = eANNOT_TextEdit2
        Case "ID_Text3"
            eType = eANNOT_TextEdit3
        Case "ID_Text4"
            eType = eANNOT_TextEdit4
        Case "ID_Ellipse"
            eType = eANNOT_Ellipse
        Case "ID_GannLines"
            eType = eANNOT_GannLines
        Case "ID_FibCircle"
            eType = eANNOT_FibArcs
        Case "ID_FibTimeZones"
            eType = eANNOT_FibTimeZones
        Case "ID_FibTimeRatio"
            eType = eANNOT_FibTimeRatio
        Case "ID_TimeCycle"
            eType = eANNOT_TimeCycle
        Case "ID_AndrewFork"
            eType = eANNOT_AndrewFork
        Case "ID_SRLine"
            eType = eANNOT_SRLine
        Case "ID_SRLine2"
            eType = eANNOT_SRLine2
        Case "ID_SRLine3"
            eType = eANNOT_SRLine3
        Case "ID_SRLine4"
            eType = eANNOT_SRLine4
        Case "ID_FibExpansion"
            eType = eANNOT_FibExpansion
        Case "ID_FibFan"
            eType = eANNOT_FibFan
        Case "ID_SpResistFan"
            eType = eANNOT_SpResistFan
        Case "ID_Mirror"
            eType = eANNOT_Mirror
        Case "ID_Pattern"
            eType = eANNOT_Pattern
        Case "ID_RiskReward"
            eType = eANNOT_RiskReward
        Case "ID_Triangle"
            eType = eANNOT_TriangleWedge
        Case "ID_ChannelHighlight"
            eType = eANNOT_ChannelHighlight
        Case "ID_WaveLabels"
            eType = eANNOT_WaveLabels
        Case "ID_ElliotTimeRatio"
            eType = eANNOT_ElliotTimeRatio
        Case "ID_Bracket"
            eType = eANNOT_Bracket
        Case "ID_PivotPoints"
            eType = eANNOT_Pivot
        Case "ID_ArrowLine"
            eType = eANNOT_ArrowLine
        Case "ID_TrendChannel"
            eType = eANNOT_TrendChannel
        Case "ID_DanCodeFib"
            eType = eANNOT_DanCodeFib
        Case "ID_FibABCD"
            eType = eANNOT_FibABCD
        Case "ID_Gartley"
            eType = eANNOT_Gartley
        Case "ID_DanCodeZone"
            eType = eANNOT_DanCodeZone
        Case "ID_GannacciSwingSquare"
            eType = eANNOT_GannacciSwingSquare
        Case "ID_GannacciCycle"
            eType = eANNOT_GannacciCycle
        Case "ID_GannacciTime"
            eType = eANNOT_GannacciTime
        Case "ID_GannacciSwing1"
            eType = eANNOT_GannacciSwing1
        Case "ID_GannacciSwing2"
            eType = eANNOT_GannacciSwing2
        Case "ID_BalloonStrangle"
            eType = eANNOT_BalloonStrangle
        Case "ID_AdvRiskReward"
            eType = eANNOT_AdvRiskReward
        Case Else
            eType = eANNOT_UndefinedType
    End Select
    
    AnnotTypeFromToolID = eType
    
    Exit Function

ErrSection:
    AnnotTypeFromToolID = eANNOT_UndefinedType
    RaiseError "cAnnotation.AnnotTypeFromToolID"
    
End Function

Private Function SpreadComponentsOk(Chart As cChart) As Boolean
On Error GoTo ErrSection:

    Dim bOK As Boolean, i&
    Dim MultiBars As Variant            'array of bars for spread symbols
    Dim Bars As cGdBars
    
    Dim dTickMove#, dTickValue#

    MultiBars = Chart.SpreadBars
    Set Bars = Chart.Bars
    
    If Not Bars Is Nothing Then
        dTickMove = Bars.Prop(eBARS_TickMove)
        dTickValue = Bars.Prop(eBARS_TickValue)
        bOK = True
        For i = 0 To UBound(MultiBars)
            Set Bars = MultiBars(i)
            If Not Bars Is Nothing Then
                If dTickMove <> Bars.Prop(eBARS_TickMove) Then
                    bOK = False
                ElseIf dTickValue <> Bars.Prop(eBARS_TickValue) Then
                    bOK = False
                End If
            End If
            If Not bOK Then Exit For
        Next
    End If
    
    SpreadComponentsOk = bOK
    
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SpreadComponentsOk"

End Function

Private Function BarsForDate(InBars As cGdBars, ByVal nBar&, ByVal dDate#) As cGdBars
On Error GoTo ErrSection:
    
    Dim OutBars As cGdBars
    Dim dDate1#, dDate2#
    Dim nSymID&, rc&
    Dim strExtData$, strSymbol$, strFileName$, strErrMsg$
            
    If InBars Is Nothing Then Exit Function
    
    If nBar >= 0 And nBar < InBars.Size Then
        Set OutBars = New cGdBars
        dDate1 = InBars(eBARS_DateTime, nBar - 1) + 1
        dDate2 = InBars(eBARS_DateTime, nBar)
        nSymID = InBars.Prop(eBARS_SymbolID)
        'return 1-minute bars for 1-day bars else return daily bars
        If InBars.Prop(eBARS_Periodicity) < ePRD_Days + 2 Then
            If nSymID <> 0 Then
                rc = DM_GetBars(OutBars, nSymID, ePRD_Minutes, Int(dDate2), Int(dDate2))
            Else
                'SymID = 0 means not our symbol, no intraday data available anyhow
                'return nothing so calling routine won't continue unnecssary processing
                Set OutBars = Nothing
            End If
        Else
            If nSymID <> 0 Then
                rc = DM_GetBars(OutBars, nSymID, ePRD_Days, dDate1, dDate2)
            ElseIf Not InBars.IsIntraday And Not m.Chart Is Nothing Then
                strExtData = m.Chart.ExternalData
                If Len(strExtData) > 0 Then
                    ' load data from external file (e.g. CSI, MS)       - aardvark 6604
                    strSymbol = Parse(strExtData, "|", 1)
                    strFileName = Parse(strExtData, "|", 2)
                    OutBars.ArrayMask = eBARS_Eod
                    SetBarProperties OutBars, strSymbol
                    rc = OutBars.FromFile(Parse(strExtData, "|", 5), FilePath(strFileName), strSymbol, "", strErrMsg, dDate1, dDate2)
                End If
            End If
        End If
    End If

    Set BarsForDate = OutBars
    
    Set OutBars = Nothing
    
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.BarsForDate"

End Function

Private Function SetPointDateTime(ByVal dY#, ByVal dDate#, ByVal nBar&, InBars As cGdBars) As Double
On Error GoTo ErrSection:

    Dim strHigh$, strLow$, strY$
    Dim i&, iStart&, iEnd&
    Dim dDate1#, dDate2#
    
    If InBars Is Nothing Then Exit Function
    
    strY = InBars.PriceDisplay(dY)
    SetPointDateTime = dDate#
    
    iStart = -1
    iEnd = -1
    If m.Chart.Bars.IsIntraday Then
        dDate1 = m.Chart.Bars(eBARS_DateTime, nBar - 1)
        dDate2 = dDate
        'limit search to bar user clicked on and bar immediately before it
        iStart = InBars.FindDateTime(dDate1)
        iEnd = InBars.FindDateTime(dDate2)
        'make sure that time of 1-minute bar is > than time of bar immediately prior to bar user clicked on
        If InBars(eBARS_DateTime, iStart) <= dDate1 Then iStart = iStart + 1
    Else
        iStart = 0
        iEnd = InBars.Size - 1
    End If
    
    If iStart = -1 Or iEnd = -1 Then Exit Function
    
    For i = iStart To iEnd
        strHigh = InBars.PriceDisplay(InBars(eBARS_High, i))
        strLow = InBars.PriceDisplay(InBars(eBARS_Low, i))
        If strY = strHigh Or strY = strLow Then
            SetPointDateTime = InBars(eBARS_DateTime, i)
            Exit For
        End If
    Next

    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.SetPointDateTime"

End Function

Private Function IsOnHighLow(InBars As cGdBars, ByVal nBar&, ByVal dY#) As Boolean
On Error GoTo ErrSection:

    Dim strY$, strHigh$, strLow$, strOpen$, strClose$
    Dim bOnHighLow As Boolean
    Dim bOnOpenClose As Boolean

    If InBars Is Nothing Then Exit Function
    
    If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 _
       Or m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_DNRetracement Or m.eType = eANNOT_FibABCD Then
        IsOnHighLow = True
        Exit Function
    End If
    
    If nBar >= 0 And nBar < InBars.Size Then
        strY = InBars.PriceDisplay(dY)
        strHigh = InBars.PriceDisplay(InBars(eBARS_High, nBar))
        strLow = InBars.PriceDisplay(InBars(eBARS_Low, nBar))
        If strY = strHigh Or strY = strLow Then bOnHighLow = True
    End If
        
    IsOnHighLow = bOnHighLow
    
    If m.eType = eANNOT_Trendline Or m.eType = eANNOT_Trendline2 Or _
       m.eType = eANNOT_Trendline3 Or m.eType = eANNOT_Trendline4 Then
       
        strOpen = InBars.PriceDisplay(InBars(eBARS_Open, nBar))
        strClose = InBars.PriceDisplay(InBars(eBARS_Close, nBar))
        If strY = strOpen Or strY = strClose Then bOnOpenClose = True
        
        If dY = m.Y1 Then
            Prop("Y1OnHighLow") = Abs(bOnHighLow)
            Prop("Y1OnOpenClose") = Abs(bOnOpenClose)
        ElseIf dY = m.Y2 Then
            Prop("Y2OnHighLow") = Abs(bOnHighLow)
            Prop("Y2OnOpenClose") = Abs(bOnOpenClose)
        End If
    End If

    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.IsOnHighLow"

End Function

Public Sub AssignDateTime()
On Error GoTo ErrSection:

    Dim ChartBars As cGdBars
    Dim BarsDay As cGdBars, BarsMinute As cGdBars
    Dim nBar&, i&, dTemp#
    
    Dim bIsOnHighLow As Boolean
    
    m.bOriginalVertical = False

    If m.Chart Is Nothing Then Exit Sub     'no chart --> no bars, can't do anything
    If Len(m.Chart.SpreadSymbols) > 0 Then Exit Sub     'individual contract bars do not line up correctly - 5874
    
    If m.eUsage = eANNOT_PatternProfit Or m.eUsage = eANNOT_FibClusters Then
        Exit Sub
    ElseIf m.eType = eANNOT_Icon Or m.eType = eANNOT_ElliotLabel Or m.eType = eANNOT_VertLine Then
        m.dDate2 = m.dDate1         '6987
        Exit Sub
    ElseIf m.eType = eANNOT_HorzLine Or m.eType = eANNOT_HorzLine2 Or _
       m.eType = eANNOT_HorzLine3 Or m.eType = eANNOT_HorzLine4 Or _
       m.eType = eANNOT_FibTimeZones Or m.eType = eANNOT_TimeCycle Or _
       m.eType = eANNOT_Mirror Or m.eType = eANNOT_Pattern Or _
       m.eType = eANNOT_RegressionLine Or m.eType = eANNOT_Bracket Or _
       m.eType = eANNOT_RiskReward Or m.eType = eANNOT_TextEdit Or _
       m.eType = eANNOT_TextEdit2 Or m.eType = eANNOT_TextEdit3 Or m.eType = eANNOT_TextEdit4 Or _
       m.eType = eANNOT_DanCodeZone Or m.eType = eANNOT_GannacciCycle Or _
       m.eType = eANNOT_GannacciTime Or m.eType = eANNOT_GannacciSwing1 Or m.eType = eANNOT_GannacciSwing2 Then

        Exit Sub                            'don't need to do anything for these tools
    End If
    
    Set ChartBars = m.Chart.Bars
    If ChartBars Is Nothing Then Exit Sub
        
    If ChartBars.IsIntraday Then
        If ChartBars.Prop(eBARS_PeriodType) = ePRD_Minutes Then
            If ChartBars.Prop(eBARS_PeriodsPerBar) = 1 Then
                Exit Sub        '1-minute bars, don't need to do anything
            End If
        Else
            Exit Sub            'can't do anything with other intraday types - 4725
        End If
    ElseIf ChartBars.Prop(eBARS_PeriodType) = ePRD_EodRenko Or _
           ChartBars.Prop(eBARS_PeriodType) = ePRD_EodKagi Or _
           ChartBars.Prop(eBARS_PeriodType) = ePRD_EodPF Or _
           ChartBars.Prop(eBARS_PeriodType) = ePRD_EodVol Then
        
        Exit Sub
        
    End If
    
'Design note:
'1. identify bars on chart that annot date/y-value pairs are anchored on
'2. for each date/y-value pairs determine if y-value is on high or low of associated bar
'3. if #2 is false then we are done else do followings
'
'4. if chart bars = 1-day bars, get 1-minute bars for bar that annot is anchored on
'5. using 1-minute bars assign date-time to annot appropriately (we are done)
'
'6. if chart bars > 1-day bars, get daily bars for bar that annot is anchored on
'7. get 1-minute bars for day that high/low occured on
'8. using 1-minute bars assign date-time to annot appropriately (we are done)

    
    'check m.dDate1 / m.Y1 pair
    nBar = ChartBars.FindDateTime(m.dDate1, True)
    bIsOnHighLow = IsOnHighLow(ChartBars, nBar, m.Y1)
    
    If bIsOnHighLow Then
        If ChartBars.Prop(eBARS_Periodicity) < ePRD_Days + 2 Then
            Set BarsMinute = BarsForDate(ChartBars, nBar, m.dDate1)
        Else
            Set BarsDay = BarsForDate(ChartBars, nBar, m.dDate1)
        End If
    End If
    
    Select Case m.eType
        Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4, eANNOT_BalloonStrangle
            If bIsOnHighLow Then
                If BarsMinute Is Nothing And Not BarsDay Is Nothing Then
                    m.dDate1 = SetPointDateTime(m.Y1, m.dDate1, nBar, BarsDay)
                    nBar = BarsDay.FindDateTime(m.dDate1, True)
                    Set BarsMinute = BarsForDate(BarsDay, nBar, m.dDate1)
                End If
                dTemp = SetPointDateTime(m.Y1, m.dDate1, nBar, BarsMinute)
                If dTemp <> 0 Then
                    m.dDate1 = SetPointDateTime(m.Y1, m.dDate1, nBar, BarsMinute)
                End If
            End If
            'JM 05-23-2012: removed all code relating to the second date
            'for SR Line, only the first date matter as far as whether it is on high or low
        
        Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
             eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4, _
             eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, eANNOT_Fibonacci4, _
             eANNOT_FibExpansion, eANNOT_FibFan, eANNOT_FibTimeRatio, eANNOT_TargetShooter, _
             eANNOT_Ellipse, eANNOT_Rectangle, eANNOT_SpResistFan, eANNOT_DNExpansion, _
             eANNOT_DNExpansion2, eANNOT_DNExpansion3, eANNOT_DNExpansion4, _
             eANNOT_DNRetracement, eANNOT_AndrewFork, eANNOT_WaveLabels, eANNOT_TriangleWedge, _
             eANNOT_ChannelHighlight, eANNOT_GannLines, eANNOT_TrendChannel, eANNOT_FibABCD, _
             eANNOT_GannacciSwingSquare, eANNOT_AdvRiskReward
             
            Dim nBarDate1&
            
            If bIsOnHighLow Then
                If BarsMinute Is Nothing And Not BarsDay Is Nothing Then
                    m.dDate1 = SetPointDateTime(m.Y1, m.dDate1, nBar, BarsDay)
                    nBar = BarsDay.FindDateTime(m.dDate1, True)
                    Set BarsMinute = BarsForDate(BarsDay, nBar, m.dDate1)
                End If
                If Not BarsMinute Is Nothing Then
                    m.dDate1 = SetPointDateTime(m.Y1, m.dDate1, nBar, BarsMinute)
                End If
            End If
            
            Set BarsDay = Nothing
            Set BarsMinute = Nothing
            
            nBarDate1 = nBar
            'check m.dDate2 / m.Y2 pair
            nBar = ChartBars.FindDateTime(m.dDate2, True)
            bIsOnHighLow = IsOnHighLow(ChartBars, nBar, m.Y2)
            
            If nBar = nBarDate1 Then
                'User have drawn a vertical line from high to low of same bar.
                'Assigning a different time to 2nd date value will cause annotation to disappear on refresh (see geSetPoints)
                m.dDate2 = m.dDate1     '4328
            ElseIf bIsOnHighLow Then
                If ChartBars.Prop(eBARS_Periodicity) < ePRD_Days + 2 Then
                    Set BarsMinute = BarsForDate(ChartBars, nBar, m.dDate2)
                Else
                    Set BarsDay = BarsForDate(ChartBars, nBar, m.dDate2)
                End If
                If BarsMinute Is Nothing And Not BarsDay Is Nothing Then
                    m.dDate2 = SetPointDateTime(m.Y2, m.dDate2, nBar, BarsDay)
                    nBar = BarsDay.FindDateTime(m.dDate2, True)
                    Set BarsMinute = BarsForDate(BarsDay, nBar, m.dDate2)
                End If
                If Not BarsMinute Is Nothing Then
                    m.dDate2 = SetPointDateTime(m.Y2, m.dDate2, nBar, BarsMinute)
                End If
            End If
                        
            If m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 Or _
                m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_DNRetracement Or _
                m.eType = eANNOT_TriangleWedge Or m.eType = eANNOT_ChannelHighlight Or _
                m.eType = eANNOT_AndrewFork Or m.eType = eANNOT_WaveLabels Or _
                m.eType = eANNOT_FibABCD Then
                
                For i = 0 To m.aDates.Size - 1
                    Set BarsDay = Nothing
                    Set BarsMinute = Nothing
                    nBar = ChartBars.FindDateTime(m.aDates(i), True)
                    bIsOnHighLow = IsOnHighLow(ChartBars, nBar, m.aYs(i))
                    If bIsOnHighLow Then
                        If ChartBars.Prop(eBARS_Periodicity) < ePRD_Days + 2 Then
                            Set BarsMinute = BarsForDate(ChartBars, nBar, m.aDates(i))
                        Else
                            Set BarsDay = BarsForDate(ChartBars, nBar, m.aDates(i))
                        End If
                        If BarsMinute Is Nothing And Not BarsDay Is Nothing Then
                            m.aDates(i) = SetPointDateTime(m.aYs(i), m.aDates(i), nBar, BarsDay)
                            nBar = BarsDay.FindDateTime(m.aDates(i), True)
                            Set BarsMinute = BarsForDate(BarsDay, nBar, m.aDates(i))
                        End If
                        If Not BarsMinute Is Nothing Then
                            m.aDates(i) = SetPointDateTime(m.aYs(i), m.aDates(i), nBar, BarsMinute)
                        End If
                    End If
                Next
            ElseIf m.eType = eANNOT_DollarLine Or m.eType = eANNOT_DollarLine2 Or m.eType = eANNOT_DollarLine3 Or _
                   m.eType = eANNOT_DollarLine4 Or m.eType = eANNOT_GannacciSwingSquare Then
                gdSetSize m.geAnnStruct.gdhMisc, 0, 0           'to trigger autoposition text in grapheng.dll
            
            ElseIf m.eType = eANNOT_Trendline Or m.eType = eANNOT_Trendline2 _
                Or m.eType = eANNOT_Trendline3 Or m.eType = eANNOT_Trendline4 _
                Or m.eType = eANNOT_TrendChannel Then
                
                If m.X1 = m.X2 Then m.bOriginalVertical = True
            End If
    
    End Select
    
    Set ChartBars = Nothing
    Set BarsDay = Nothing
    Set BarsMinute = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.AssignDateTime"

End Sub

Private Sub CalcPivots(Optional ByVal bForceRecalc As Boolean = False)
On Error GoTo ErrSection:

    Static dPrevDate1#, dPrevDate2#, iPrevCalcMethod&

    Dim dHighestHi#, dLowestLow#, dClose#, dNextBarOpen#, dTemp#
    Dim dR1#, dR2#, dR3#, dS1#, dS2#, dS3#, dPivot#
    Dim nBarFirst&, nBarLast&, i&, iCalcMethod&
    
    Dim Bars As cGdBars
               
    If m.Chart Is Nothing Then Exit Sub
    Set Bars = m.Chart.Bars
    If Bars Is Nothing Then Exit Sub
    If Bars.Size <= 0 Then Exit Sub
        
    nBarLast = m.Chart.LastGoodDataBar(False)
    If m.dDate1 >= Bars(eBARS_DateTime, nBarLast) Or m.dDate2 >= Bars(eBARS_DateTime, nBarLast) Then
        Prop("CalcMethod") = 1
    End If
        
    iCalcMethod = Val(Prop("CalcMethod"))
    If iCalcMethod < 1 Or iCalcMethod > 4 Then
        Prop("CalcMethod") = 1
        iCalcMethod = 1
    End If
    
    If Not bForceRecalc Then
        'no significant change, no need to calculate
        If dPrevDate1 = m.dDate1 And dPrevDate2 = m.dDate2 And _
           iPrevCalcMethod = iCalcMethod Then Exit Sub
    End If
    
    m.aYs.Size = 0
    m.aYs(0) = 0
    m.aYs(1) = 0
    m.aYs(2) = 0
    m.aYs(3) = 0
    m.aYs(4) = 0
    
    If m.dDate1 < m.dDate2 Then
        If Bars.Prop(eBARS_PeriodicityStr) = m.strBarPeriod Then
            nBarFirst = Bars.FindDateTime(m.dDate1, True)
            nBarLast = Bars.FindDateTime(m.dDate2, True)
        ElseIf InStr(m.strBarPeriod, "minute") And Not Bars.IsIntraday Then
            nBarFirst = Bars.FindDateTime(Int(m.dDate1), False)
            nBarLast = Bars.FindDateTime(Int(m.dDate2), False)
        Else
            nBarFirst = Bars.FindDateTime(m.dDate1, False)
            nBarLast = Bars.FindDateTime(m.dDate2, False)
        End If
    Else
        If Bars.Prop(eBARS_PeriodicityStr) = m.strBarPeriod Then
            nBarFirst = Bars.FindDateTime(m.dDate2, True)
            nBarLast = Bars.FindDateTime(m.dDate1, True)
        ElseIf InStr(m.strBarPeriod, "minute") And Not Bars.IsIntraday Then
            nBarFirst = Bars.FindDateTime(Int(m.dDate2), False)
            nBarLast = Bars.FindDateTime(Int(m.dDate1), False)
        Else
            nBarFirst = Bars.FindDateTime(m.dDate2, False)
            nBarLast = Bars.FindDateTime(m.dDate1, False)
        End If
    End If
    
    If nBarFirst < 0 Or nBarFirst >= Bars.Size Or _
       nBarLast < 0 Or nBarLast >= Bars.Size Then
       
       Set Bars = Nothing
       Exit Sub
       
    End If
    
    dHighestHi = Bars(eBARS_High, nBarFirst)
    dLowestLow = Bars(eBARS_Low, nBarFirst)
    dClose = Bars(eBARS_Close, nBarLast)
    dNextBarOpen = Bars(eBARS_Open, nBarLast + 1)
    
    For i = nBarFirst To nBarLast
        dTemp = Bars(eBARS_High, i)
        If dTemp > dHighestHi Then dHighestHi = dTemp
        dTemp = Bars(eBARS_Low, i)
        If dTemp <> kNullData Then
            If dTemp < dLowestLow Then dLowestLow = dTemp
        End If
    Next
    
    'pivot
    If m.dDate1 < m.dDate2 Then
        If iCalcMethod = 2 Then
            m.Y1 = (dHighestHi + dLowestLow + dClose + dNextBarOpen) / 4
        ElseIf iCalcMethod = 3 Then
            m.Y1 = (dHighestHi + dLowestLow + dNextBarOpen) / 3
        ElseIf iCalcMethod = 4 Then
            m.Y1 = (dHighestHi + dLowestLow + (dNextBarOpen * 2)) / 4
        Else
            m.Y1 = (dHighestHi + dLowestLow + dClose) / 3
        End If
        m.Y2 = 2 * m.Y1 - dHighestHi    'S1
        dPivot = m.Y1
        dS1 = m.Y2
    Else
        If iCalcMethod = 2 Then
            m.Y2 = (dHighestHi + dLowestLow + dClose + dNextBarOpen) / 4
        ElseIf iCalcMethod = 3 Then
            m.Y2 = (dHighestHi + dLowestLow + dNextBarOpen) / 3
        ElseIf iCalcMethod = 4 Then
            m.Y2 = (dHighestHi + dLowestLow + (dNextBarOpen * 2)) / 4
        Else
            m.Y2 = (dHighestHi + dLowestLow + dClose) / 3
        End If
        m.Y1 = 2 * m.Y2 - dHighestHi    'S1
        dPivot = m.Y2
        dS1 = m.Y1
    End If
        
    
    dR1 = 2 * dPivot - dLowestLow
    dR2 = dPivot + dR1 - dS1
    dR3 = dPivot - dS1 + dR2
    dS2 = dPivot - (dR1 - dS1)
    dS3 = dPivot - (dR2 - dS1)
    
'   aYs[0/S2] [1/S3] [2/R1] [3/R2] [4/R3]
    m.aYs(0) = dS2
    m.aYs(1) = dS3
    m.aYs(2) = dR1
    m.aYs(3) = dR2
    m.aYs(4) = dR3
    
    dPrevDate1 = m.dDate1
    dPrevDate2 = m.dDate2
    
    iPrevCalcMethod = iCalcMethod
            
    Set Bars = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.CalcPivots"

End Sub

Public Property Get BarPeriodAtCreate() As String
    BarPeriodAtCreate = m.strBarPeriod
End Property

Public Property Get PeriodsPerBarAtCreate() As String
    PeriodsPerBarAtCreate = m.strPeriodsPerBar
End Property

Public Function EnableBarComplete(ByVal nIndIdx As Long) As Boolean
On Error GoTo ErrSection:

    Dim iBar As Long, iBarScreen As Long, i As Long
    Dim dIndVal As Double, dIndValPrev As Double
    Dim dAnnotVal As Double, dAnnotValPrev As Double
    
    Dim Chart As cChart
    Dim Ind As cIndicator
    
    Dim bEnable As Boolean
        
    Set Chart = m.Chart
    
    If Chart Is Nothing Then Exit Function
    
    If Chart.Tree.NodeLevel(nIndIdx) > 0 Then Set Ind = Chart.Tree(nIndIdx)
    
    If Ind Is Nothing Then Exit Function

    iBar = Chart.LastGoodDataBar(False)
    iBarScreen = Chart.LastGoodDataBar(False, True)
        
    If iBar > 0 And iBar < Chart.Bars.Size Then
        Select Case m.eType
            Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4
                If iBar = iBarScreen Then
                    i = gePointNum(Chart, Chart.Bars(eBARS_DateTime, iBar))
                    dAnnotVal = geTrendValueY(m.geAnnStruct.hObj, i)
                    i = gePointNum(Chart, Chart.Bars(eBARS_DateTime, iBar - 1))
                    dAnnotValPrev = geTrendValueY(m.geAnnStruct.hObj, i)
                Else
                    GoTo ErrExit            'last good data bar is not on screen
                End If
            Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4, eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
                dAnnotVal = m.Y1
                dAnnotValPrev = dAnnotVal
        End Select
        
        If Ind.isPriceInd Then
            If dAnnotVal > Ind.Bars(eBARS_High, iBar) And dAnnotValPrev > Ind.Bars(eBARS_High, iBar - 1) Then
                bEnable = True
            ElseIf dAnnotVal < Ind.Bars(eBARS_Low, iBar) And dAnnotValPrev < Ind.Bars(eBARS_Low, iBar - 1) Then
                bEnable = True
            End If
        ElseIf (dAnnotVal > Ind.Data(iBar) And dAnnotValPrev > Ind.Data(iBar - 1)) Or _
               (dAnnotVal < Ind.Data(iBar) And dAnnotValPrev < Ind.Data(iBar - 1)) Then
               bEnable = True
        End If
    End If
        
ErrExit:
    EnableBarComplete = bEnable
    Exit Function
    
ErrSection:
    RaiseError "cAnnotation.EnableBarComplete"

End Function

Public Sub BidAskChartShowHide(ByVal bShow As Boolean)
On Error GoTo ErrExit:

    Dim i&

    If bShow Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 7
        gdSetNum m.geAnnStruct.glhAnnType, 1, 7
        geSetPoints m.Chart
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAnnotation.BidAskChartShowHide"

End Sub

Public Property Get IsGreenBlattTool() As Boolean
On Error GoTo ErrSection:

    If m.eType = eANNOT_FibTimeZones Then IsGreenBlattTool = m.bGreenBlatt

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.IsGreenBlattTool"

End Property

Public Property Get ZoneInUse() As eGBZoneInUse
On Error GoTo ErrSection:

    Dim i&, eZone As eGBZoneInUse
    Dim eZoneTmp As eGBZoneInUse
    
    eZoneTmp = Int(Val(Prop("ZoneInUse")))
    
    
    If IsGreenBlattTool Then
        If eZoneTmp >= eANNOT_GB_ZoneLucas And eZoneTmp <= eANNOT_GB_ZoneSqRoot Then
            eZone = eZoneTmp
        Else
            eZone = eANNOT_GB_ZoneFib
        End If
    Else
        If eZoneTmp = eANNOT_GB_ZoneLucas Then
            eZone = eANNOT_GB_ZoneLucas
        Else
            eZone = eANNOT_GB_ZoneFib
        End If
    End If
    
    ZoneInUse = eZone

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.ZoneInUse.Get"

End Property

Public Property Let ZoneInUse(eZone As eGBZoneInUse)
On Error GoTo ErrSection:

    Prop("ZoneInUse") = eZone

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.ZoneInUse.Let"

End Property

'JM: 06-02-2008
'this is just a quick way to get a specific default for an annotation without creating it.
'Currently only implementing request for DNExpansion free-float property.
Public Property Get DefaultProp(ByVal eType As eAnnotType, ByVal strProp$) As Long
On Error GoTo ErrSection:

    Dim bOkay As Boolean

    bOkay = eType = eANNOT_DNExpansion Or eType = eANNOT_DNExpansion2 Or eType = eANNOT_DNExpansion3 Or _
            eType = eANNOT_DNExpansion4 Or eType = eANNOT_FibABCD

    If m.eType = eANNOT_UndefinedType Then
        If bOkay Or m.eType = eANNOT_Gartley Or m.eType = eANNOT_GannacciSwing1 Or m.eType = eANNOT_GannacciSwing2 Then
            m.eType = eType
        Else
            m.eType = eANNOT_DNExpansion
        End If
    ElseIf Not bOkay Then
        Exit Property
    End If
    
    If strProp = "FreeFloat" Then
        If UseDiNapFib() And m.eType = eANNOT_DNExpansion Then
            DefaultProp = 0
        ElseIf m.eType = eANNOT_GannacciSwing1 Or m.eType = eANNOT_GannacciSwing2 Then
            DefaultProp = 1
        Else
            DefaultProp = ValOfText(LoadDefault("FreeFloat", 0))
        End If
    End If
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.DefaultProp.Get"

End Property

'JM: 06-02-2008 (validity check for Dinapoli Expansion / Fib Extension tool)
'This tool original designed to snap to hi-low-hi or low-hi-low.
'For EWI "free floating" implementation this translates B must be > OR < both A & C
Private Property Get ValidDNE() As Boolean
On Error GoTo ErrSection:

    Dim a#, b#, c#
    Dim bValid As Boolean
    Dim bOkay As Boolean
    
    bOkay = m.eType = eANNOT_DNExpansion Or m.eType = eANNOT_DNExpansion2 Or m.eType = eANNOT_DNExpansion3 Or _
            m.eType = eANNOT_DNExpansion4 Or m.eType = eANNOT_FibABCD Or m.eType = eANNOT_GannacciSwing2

    If m.eType = eANNOT_GannacciSwing1 Or m.eType = eANNOT_GannacciSwing2 _
       Or (UseDiNapFib() And m.eType = eANNOT_DNExpansion) Then
        ValidDNE = True
        Exit Property
    ElseIf Not bOkay Then
        Exit Property
    End If
        
    a = m.dDate1
    b = m.dDate2
    c = gdGetNum(m.aDates.ArrayHandle, 0)
    If a > b And b > c Then
        bValid = True
    ElseIf a < b And b < c Then
        bValid = True
    ElseIf b = a Or b = c Then
        bValid = True
    Else
        bValid = False
    End If
    
    ValidDNE = bValid

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.ValidFocusDNE.Get"

End Property

Public Sub SnapFreeFloat(frm As Form)
On Error GoTo ErrSection:

    Dim dY#, dHigh#, dLow#
    Dim i&, iSize&
    
    If Not TypeOf frm Is frmEditAnnot Then Exit Sub
    If m.Chart Is Nothing Then Exit Sub

    iSize = m.Chart.Bars.Size
    
    i = m.Chart.Bars.FindDateTime(m.dDate1)
    If i >= 0 And i < iSize Then
        dHigh = m.Chart.Bars(eBARS_High, i)
        dLow = m.Chart.Bars(eBARS_Low, i)
        If m.Y1 > dHigh Then
            m.Y1 = dHigh
        ElseIf m.Y1 > dLow Then
            If dHigh - m.Y1 > m.Y1 - dLow Then
                m.Y1 = dLow
                Prop("HiLo") = 1
            Else
                m.Y1 = dHigh
                Prop("HiLo") = 0
            End If
        Else
            m.Y1 = dLow
            Prop("HiLo") = 1
        End If
    End If
    
    i = m.Chart.Bars.FindDateTime(m.dDate2)
    If i >= 0 And i < iSize Then
        dHigh = m.Chart.Bars(eBARS_High, i)
        dLow = m.Chart.Bars(eBARS_Low, i)
        If m.Y2 > dHigh Then
            m.Y2 = dHigh
        ElseIf m.Y2 > dLow Then
            If dHigh - m.Y2 > m.Y2 - dLow Then
                m.Y2 = dLow
            Else
                m.Y2 = dHigh
            End If
        Else
            m.Y2 = dLow
        End If
    End If
    
    i = m.Chart.Bars.FindDateTime(m.aDates(0))
    If i >= 0 And i < iSize Then
        dHigh = m.Chart.Bars(eBARS_High, i)
        dLow = m.Chart.Bars(eBARS_Low, i)
        dY = m.aYs(0)
        If dY > dHigh Then
            m.aYs(0) = dHigh
        ElseIf dY > dLow Then
            If dHigh - dY > dY - dLow Then
                m.aYs(0) = dLow
            Else
                m.aYs(0) = dHigh
            End If
        Else
            m.aYs(0) = dLow
        End If
    End If
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.SnapFreeFloat"

End Sub

Public Sub geExportAnnot(Chart As cChart, ByVal nPaneID&, ByVal nID&, ByVal nStartScreenX&)
On Error GoTo ErrSection:

    If m.geAdded = True Then
        geSetPoints Chart, nStartScreenX        '4525
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.geExportAnnot"

End Sub

Public Property Get DuplicateAllow() As Boolean
On Error GoTo ErrSection:

    Dim bAllow As Boolean
        
    If m.eUsage = eANNOT_UserAdded Then
        Select Case m.eType
            Case eANNOT_DNRetracement, eANNOT_DNExpansion, eANNOT_DNExpansion2, _
                 eANNOT_DNExpansion3, eANNOT_DNExpansion4, _
                 eANNOT_TimeCycle, eANNOT_FibTimeZones, _
                 eANNOT_Mirror, eANNOT_Pattern, _
                 eANNOT_WaveLabels, eANNOT_RegressionLine, _
                 eANNOT_TargetShooter, eANNOT_Pivot, eANNOT_FibABCD, _
                 eANNOT_Gartley, eANNOT_DanCodeZone, eANNOT_GannacciCycle, _
                 eANNOT_GannacciTime, eANNOT_BalloonStrangle
                 
                bAllow = False
                          
             Case eANNOT_ElliotLabel
                If AllowEWI Or AllowGMP Then
                    bAllow = True
                ElseIf IsEndUserEWI And HasModule("EWL") Then
                    bAllow = True
                Else
                    bAllow = False
                End If

            Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4
                If Prop("FakePriceAlert") <> 1 Then bAllow = True
             
             Case Else
                bAllow = True
        End Select
    End If
        
    DuplicateAllow = bAllow

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.DuplicateAllow.Get"

End Property

Private Sub SetPercentPoints()
On Error GoTo ErrSection:

    Dim dPercent#, dPoints#, dMinMove#, nChannelType&
    
    If m.eType <> eANNOT_Trendline And m.eType <> eANNOT_Trendline2 And _
       m.eType <> eANNOT_Trendline3 And m.eType <> eANNOT_Trendline4 And _
       m.eType <> eANNOT_TrendChannel Then
       
       Exit Sub
    
    End If
    
    If Not m.Chart Is Nothing Then
        If Not m.Chart.Bars Is Nothing Then
            dMinMove = m.Chart.Bars.MinMove()
        End If
    End If
    
    nChannelType = Val(Prop("ChannelType"))

    'set percent & point values to approx equal each other
    If nChannelType = 0 Then        'percent
        dPercent = Val(Prop("ChannelPercent")) / 100
        If dPercent = 0 Then
            dPercent = 1
            Prop("ChannelPercent") = dPercent
        End If
        dPoints = m.Y2 * dPercent
        Prop("ChannelPoints") = Abs(RoundToMinMove(dPoints, dMinMove))
    Else
        dPoints = Val(Prop("ChannelPoints"))
        If dPoints = 0 Then
            dPoints = 3
            Prop("ChannelPoints") = dPoints
        End If
        dPercent = ((m.Y2 + dPoints) - m.Y2) / m.Y2
        Prop("ChannelPercent") = Abs(RoundNum(dPercent * 100, 2))
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.SetPercentPoints"

End Sub

Public Sub LoadOptionsInfo(aInput As cGdArray)

    Dim i&, k&
    Dim strText$, strSym$, strContract$
    Dim bStocks As Boolean
    
    Dim hInput&, hStrike&, hCall&, hPut&

    If aInput Is Nothing Then Exit Sub
    
    If m.eType = eANNOT_VertLine And m.eUsage = eANNOT_OptionInfo Then
                
        If Not m.Chart Is Nothing Then
            If InStr(m.Chart.Symbol, "-") = 0 Then
                bStocks = True
            Else
                strSym = Parse(m.Chart.Symbol, "-", 1) & "-"
                strText = DateFormat(m.dDate2, MM_DD_YYYY)
                
                If Month(m.dDate2) < 10 Then            '5056
                    strContract = Year(m.dDate2) & "0" & Month(m.dDate2)
                Else
                    strContract = Year(m.dDate2) & Month(m.dDate2)
                End If
                
                strSym = strSym & strContract
            End If
        End If
        
        If m.tbOptionsInfo Is Nothing Then
            Set m.tbOptionsInfo = New cGdTable
            m.tbOptionsInfo.CreateField eGDARRAY_Doubles, , , -1
            m.tbOptionsInfo.CreateField eGDARRAY_Strings
            m.tbOptionsInfo.CreateField eGDARRAY_Strings
        End If
        
        m.tbOptionsInfo.NumRecords = 0
        
        hInput = aInput.ArrayHandle
        hStrike = m.tbOptionsInfo.FieldArrayHandle(0)
        hCall = m.tbOptionsInfo.FieldArrayHandle(1)
        hPut = m.tbOptionsInfo.FieldArrayHandle(2)
        
        For i = 0 To aInput.Size - 1
            strText = gdGetStr(hInput, i)
            
            m.tbOptionsInfo.AddRecord ""
            k = m.tbOptionsInfo.NumRecords - 1
                        
            If bStocks Then
                gdSetNum hStrike, k, ValOfText(Parse(strText, ";", 1))
                gdSetStr hCall, k, Parse(strText, ";", 2)
                gdSetStr hPut, k, Parse(strText, ";", 3)
            Else
                strText = Parse(strText, ";", 1)
                gdSetNum hStrike, k, ValOfText(strText)
                gdSetStr hCall, k, strSym & " C" & strText
                gdSetStr hPut, k, strSym & " P" & strText
            End If
        Next
        
    End If
    
    If m.tbOptionsInfo.NumRecords > 0 Then
        Set m.aOptionsIdx = Nothing
        Set m.aOptionsIdx = m.tbOptionsInfo.CreateSortedIndex(0, eGdSort_Default)
    End If
    

End Sub

Public Function ClosestStrike(ByVal dPrice#, _
    Optional ByVal bAbove As Boolean = False, _
    Optional ByVal bBelow As Boolean = False) As String

    Dim strInfo$, i&, k&
    Dim d1#, d2#, d3#

    If Not m.tbOptionsInfo Is Nothing Then
        If Not m.aOptionsIdx Is Nothing Then
            m.tbOptionsInfo.SearchAsIndex m.aOptionsIdx, 0, dPrice, i
            If i >= m.aOptionsIdx.Size Then
                'price passed in is greater than all strike prices
                k = m.aOptionsIdx.Size - 1
                strInfo = Str(m.tbOptionsInfo(0, k)) & ";" & m.tbOptionsInfo(1, k) & ";" & m.tbOptionsInfo(2, k)
            ElseIf i >= 0 And i < m.aOptionsIdx.Size Then
                k = m.aOptionsIdx(i)
                d1 = m.tbOptionsInfo(0, k)
                
                If bAbove And d1 < dPrice Then
                    k = m.aOptionsIdx(i + 1)
                ElseIf bBelow And d1 > dPrice Then
                    k = m.aOptionsIdx(i - 1)
                End If
                
                strInfo = Str(m.tbOptionsInfo(0, k)) & ";" & m.tbOptionsInfo(1, k) & ";" & m.tbOptionsInfo(2, k)
                If m.tbOptionsInfo.NumFields = 5 Then
                    strInfo = strInfo & ";" & m.tbOptionsInfo(3, k) & ";" & m.tbOptionsInfo(4, k)
                End If
            End If
        Else
            'theoretically should never get here
            DebugLog m.Chart.Symbol & " : " & Str(dPrice) & " : Annot.ClosestStrike failed. Index array is nothing."
        End If
    End If
    
    ClosestStrike = strInfo

End Function

Public Property Get geLeftX(ByVal idx&) As Long
    geLeftX = gdGetNum(m.geAnnStruct.gdhLeft, 0)
End Property

Public Property Get FirstOptionAnnnot() As Boolean
    FirstOptionAnnnot = m.bFirstOption
End Property

Public Property Let FirstOptionAnnnot(ByVal bFlag As Boolean)
    m.bFirstOption = bFlag
End Property

Public Sub SetChannelLocationOnAdd(ByVal dY#)
On Error Resume Next:           'this is called from MouseMove in chart's form
    
    Dim i&, dY3#
    
    If m.Chart Is Nothing Then Exit Sub

    i = LoadDefault("ChannelLocation", 3)
    If i = 0 Or i = 3 Then
        Prop("ChannelLocation") = 3
    Else
        dY3 = geTrendValueY(m.geAnnStruct.hObj, GetX(m.Chart, m.dDate1) - m.Chart.ScreenStartX)
    
        If dY3 > dY Then
            Prop("ChannelLocation") = 2     'below
        Else
            Prop("ChannelLocation") = 1     'above
        End If
    End If
    
    m.nHitItemIdx = 5

End Sub

Public Property Get AlertAddEditInprog() As Boolean
On Error GoTo ErrSection:

    If m.geAnnStruct.moveable = 99999 Then AlertAddEditInprog = True

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.AlertAddEditInprog.Get"

End Property

Public Property Let AlertAddEditInprog(ByVal bAddEdit As Boolean)
On Error GoTo ErrSection:

    'the graphics engine checks the moveable field for =0, <>0, =1 or <>1 in various cases
    '0=user is still selecting points for annotation, 1=points selection is complete
    'alerts can only be added once a drawing tool is complete, so when alert addition done rest field to 1
    
    If bAddEdit Then
        m.geAnnStruct.moveable = 99999
    Else
        m.geAnnStruct.moveable = 1
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.AlertAddEditInprog.Let"

End Property

Public Property Get ShiftKey() As Long
On Error Resume Next:           'called from chart's mouse move

    ShiftKey = m.nShiftKey

End Property

Public Property Let ShiftKey(ByVal nShift&)
On Error Resume Next:           'called from chart's mouse move

    If nShift = 0 Or m.nHitItemIdx = 0 Then     'only allow slope freeze if main line is selected (5664)
        m.nShiftKey = nShift
    End If

End Property

Public Property Get PatternLength(Chart As cChart) As Long
On Error GoTo ErrSection:

    Dim dTemp#, iLen&, i&
    Dim Bars As cGdBars

    If Chart Is Nothing Then Exit Property
    
    Set Bars = Chart.Bars
    If Bars Is Nothing Then Exit Property
    
    If m.eUsage = eANNOT_PatternProfit Then
        m.geAnnStruct.showAxes = 1
    ElseIf m.eUsage <> eANNOT_FibClusters Then
        Exit Property
    End If
    
    If m.dDate2 < m.dDate1 Then
        dTemp = m.dDate2
        m.dDate2 = m.dDate1
        m.dDate1 = dTemp
        
        dTemp = m.X2
        m.X2 = m.X1
        m.X1 = dTemp
    End If
    
    dTemp = Bars(eBARS_DateTime, Chart.LastGoodDataBar(False, False))
    If m.dDate2 > dTemp Then m.dDate2 = dTemp
    If m.dDate1 > dTemp Then m.dDate1 = dTemp
    
    If m.dDate1 = m.dDate2 Then Exit Property
    
    'check selection starting with a blank bar
    i = Bars.FindDateTime(m.dDate1, True)
    If i = -1 Then
        i = Bars.FindDateTime(m.dDate1)
        m.dDate1 = Bars(eBARS_DateTime, i)
        i = Bars.FindDateTime(m.dDate1, True)
        If i = -1 Then Exit Property
    End If
    
'JM 07-21-2010 - original code when grapheng.dll was drawing pattern lines in-between bars (leave awhile then remove if all okay)
'    dTemp = Chart.aXdate(m.X2 - 1)
'    i = Bars.FindDateTime(dTemp, True)
    
    'check selection ending with a blank bar
    i = Bars.FindDateTime(m.dDate2, True)
    If i = -1 Then
        m.dDate2 = Chart.aXdate(m.X2 - 1)
        m.X2 = m.X2 - 1
    Else
        dTemp = m.dDate2
    End If
    
    If m.dDate1 = m.dDate2 Then Exit Property
    
    i = Bars.FindDateTime(m.dDate2)
    Do While dTemp >= m.dDate1         '07-14-2010 roll back to: While dTemp > m.dDate1
        If i > 0 Then
            iLen = iLen + 1
            i = i - 1
            dTemp = Bars(eBARS_DateTime, i)
        Else
            Exit Do
        End If
    Loop
    
    If m.eUsage = eANNOT_FibClusters Then iLen = iLen - 1
    
    PatternLength = iLen
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.PatternLength"
    
End Property

Public Function OkayToMovePFP() As Boolean
On Error GoTo ErrSection:

    If m.eUsage = eANNOT_PatternProfit Then
        If m.geAnnStruct.showAxes = 1 Then OkayToMovePFP = True
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.OkayToMovePFP"

End Function

Public Sub SetMinMaxPFP(Chart As cChart, Ind As cIndicator, ByVal bCompare As Boolean)
On Error GoTo ErrSection:

    Dim dMin#, dMax#
    
    If Not m.eUsage = eANNOT_PatternProfit Then Exit Sub        'precautionary
    If m.geAnnStruct.showAxes = 1 Then Exit Sub                 'precautionary
    
    PatternMinMax Chart, dMin, dMax, m.dDate1, m.dDate2, Ind
    
    If bCompare Then
        If dMin < m.Y1 Then m.Y1 = dMin
        If dMax > m.Y2 Then m.Y2 = dMax
    Else
        m.Y1 = dMin
        m.Y2 = dMax
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.SetMinMaxPFP"
    
End Sub

Public Sub AndrewsDefaults(ByVal strProp$, Optional ByVal bShowNone As Boolean = False)
On Error GoTo ErrSection:

    If Not m.eType = eANNOT_AndrewFork Then Exit Sub
    
    If strProp = "Parallel" Then
        Prop("ParallelRatios") = kParallelRatios
        Prop("ParallelColor") = kParallelColor
        If bShowNone Then
            Prop("ParallelShow") = kParallelShowNone
        Else
            Prop("ParallelShow") = kParallelShow
        End If
    ElseIf strProp = "Markers" Then
        Prop("MarkerRatios") = kIntervalRatios
        If bShowNone Then
            Prop("MarkerShow") = kIntervalShowNone
        Else
            Prop("MarkerShow") = kIntervalShow
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.AndrewsDefaults"
    
End Sub

Public Property Get ClusterPropDefault(ByVal strProp$) As Long
On Error GoTo ErrSection:

    Dim i&

    'JM 07-30-2015: if want heatmap as default then make negative for strProp = ZoneFillColor
    i = GetIniFileProperty(strProp, kClusterFibFillColor, kClusterIniSection, g.strIniFile)
    
    ClusterPropDefault = i
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.ClusterPropDefault"
    
End Property

Private Function MoveGartley(Chart As cChart, ByVal nPoint&, ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:
    
    Dim i&, j&, k&
    Dim dDate#, dDateMid#, dPrev#, dTemp#, d3#
    Dim Bars As cGdBars
    
    Dim bOkay As Boolean
    Dim bHigh As Boolean
    Dim bAssign As Boolean
    
    bOkay = False
    
'geMoveFlag = 0 means initial draw in progress
'geMoveFlag = 1 means this is an existing Gartley (loaded from file or user just completed drawing)

    If geMoveFlag = 1 Then
        bOkay = True
        Select Case nPoint
            Case 0:
                If X < m.dDate2 Then
                    m.dDate1 = X
                    m.Y1 = Y
                Else
                    bOkay = False
                End If
            Case 1:
                If X > m.dDate1 And X < m.aDates(0) Then
                    m.dDate2 = X
                    m.Y2 = Y
                Else
                    bOkay = False
                End If
            Case 2:
                If X > m.dDate2 And X < m.aDates(1) Then
                    m.aDates(0) = X
                    m.aYs(0) = Y
                Else
                    bOkay = False
                End If
            Case 3:
                If X > m.aDates(0) And X < m.aDates(2) Then
                    m.aDates(1) = X
                    m.aYs(1) = Y
                Else
                    bOkay = False
                End If
            Case 4:
                If X > m.aDates(1) Then
                    m.aDates(2) = X
                    m.aYs(2) = Y
                Else
                    bOkay = False
                End If
        End Select
        GoTo ErrExit
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Points 1,2,3 --> initial draw in progress; 1=X, 2=A, 3=D
'   B & C are to be calculated per these rules:
'       if A is on high of bar then B must be on low of bar and vice versa
'           B cannot exceed X
'           length AB should be a minimum of 0.382 of length XA if possible
'       if A is on high of bar then C is also on high of bar and vice versa
'           C cannot exceed A
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If nPoint = 1 Then
        bOkay = True
        
        m.dDate1 = X
        m.Y1 = Y
    ElseIf nPoint = 2 Then
        If X < m.dDate1 Then GoTo ErrExit
        
        bOkay = True
        
        m.dDate2 = X
        m.Y2 = Y
    ElseIf nPoint = 3 Then
        If X < m.dDate2 Or Chart Is Nothing Then GoTo ErrExit
        
        Set Bars = Chart.Bars
        If Bars Is Nothing Then GoTo ErrExit
        
        i = Bars.FindDateTime(X, True)          'current bar cursor is on (ie point D)
        If i < 0 Or i >= Bars.Size Then GoTo ErrExit
        
        If i > Chart.LastGoodDataBar(False) Then i = Chart.LastGoodDataBar(False)
        
        j = Bars.FindDateTime(m.dDate2, True)   'bar that A is on
        If j < 0 Or j >= Bars.Size Then GoTo ErrExit
        
        k = j + (i - j) / 2
        dDateMid = Bars(eBARS_DateTime, k)      'bar that is midway between A & D
        
        'walk forward of A to assign B to lowest low or highest high between A & D (D is where user's cursor is at)
        If m.Y2 > m.Y1 Then
            dPrev = 999999  'A is high of bar
            dTemp = m.Y2 - (m.Y2 - m.Y1) * 0.382
            bHigh = True
        Else
            dPrev = -999999
            dTemp = m.Y2 + (m.Y2 - m.Y1) * 0.382
        End If
        
        For k = j + 1 To i - 1
            dDate = Bars(eBARS_DateTime, k)
            bAssign = False
            If bHigh Then
                d3 = Bars(eBARS_Low, k)
                'B must be higher than X (m.Y1), but lower than A (m.Y2)
                If d3 < dPrev And d3 > m.Y1 And d3 < m.Y2 Then bAssign = True
            Else
                d3 = Bars(eBARS_High, k)
                'B must be lower than X (m.Y1), but higher than A (m.Y2)
                If d3 > dPrev And d3 < m.Y1 And d3 > m.Y2 Then bAssign = True
            End If
            
            'B must be higher than X (m.Y1), but lower than A (m.Y2)
            If bAssign Then
                m.aDates(0) = dDate
                m.aYs(0) = d3
                dPrev = d3
                bOkay = True
            End If
            
            If bOkay Then
                'B should preferably be at least 0.382 (dTemp) of AX & not too close to C
                If dDate >= dDateMid Then
                    If (bHigh And d3 <= dTemp) Or (Not bHigh And d3 >= dTemp) Then Exit For
                End If
            End If
        Next
            
        If Not bOkay Then GoTo ErrExit
        bOkay = False   'reset
        
        'walk forward of B to assign C
        j = Bars.FindDateTime(m.aDates(0), True)        'bar that B is on
        If j < 0 Or j >= Bars.Size Then GoTo ErrExit    'theoretically should never happen
        
        If bHigh Then
            dPrev = -999999  'A is high of bar
        Else
            dPrev = 999999
        End If
        
        dTemp = m.aYs(0)
        For k = j + 1 To i - 1
            dDate = Bars(eBARS_DateTime, k)
            bAssign = False
            If bHigh Then
                d3 = Bars(eBARS_High, k)
                'C must be higher than B (dTemp), but lower than A (m.Y2)
                If d3 > dPrev And d3 > dTemp And d3 < m.Y2 Then bAssign = True
            Else
                d3 = Bars(eBARS_Low, k)
                'C must be lower than B (dTemp), but higher than A (m.Y2)
                If d3 < dPrev And d3 < dTemp And d3 > m.Y2 Then bAssign = True
            End If
            
            If bAssign Then
                m.aDates(1) = dDate
                m.aYs(1) = d3
                dPrev = d3
                bOkay = True
            End If
        Next
        
        m.aDates(2) = X
        m.aYs(2) = Y
        
        If Not bOkay Then
            'there were no C point that meet Gartley rules; just assign last date in range
            m.aDates(1) = dDate
            m.aYs(1) = d3
            bOkay = True
        End If
        
    End If
    
ErrExit:
    MoveGartley = bOkay
    Exit Function

ErrSection:
    RaiseError "cAnnotation.MoveGartley"

End Function


Private Sub SetGartley(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&)
On Error GoTo ErrSection:

    Dim i&, nSize&
    Dim iColor1&, iColor2&, iColor3&, iColor4&
    
    Dim Str$, str1$, Str2$, str3$, Str4$, Str5$

    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)

    '[0]=X, [1]=A, [2]=B, [3]=C, [4]=D
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 36
        For i = 1 To 4
            gdSetNum m.geAnnStruct.glhAnnType, i, -1    'place holders
            gdSetNum m.geAnnStruct.gdhLeft, i, -1#
            gdSetNum m.geAnnStruct.gdhTop, i, -1#
            gdSetNum m.geAnnStruct.gdhRight, 1, -1#
            gdSetNum m.geAnnStruct.gdhBottom, 1, -1
            gdSetNum m.geAnnStruct.gdhMisc, i, -1#
        Next
    End If
    
    If m.dDate1 > 0 Then
        gdSetNum m.geAnnStruct.gdhLeft, 0, GetX(Chart, m.dDate1) - nStartScreenX
        gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
        geSetPenStyle m.nStyle, 0
        
        If m.dDate2 > 0 Then
            gdSetNum m.geAnnStruct.glhAnnType, 1, 36
            gdSetNum m.geAnnStruct.gdhLeft, 1, GetX(Chart, m.dDate2) - nStartScreenX
            gdSetNum m.geAnnStruct.gdhTop, 1, m.Y2
            
            For i = 0 To m.aDates.Size - 1
                If m.aDates(i) > 0 Then
                    gdSetNum m.geAnnStruct.glhAnnType, i + 2, 36
                    gdSetNum m.geAnnStruct.gdhLeft, i + 2, GetX(Chart, m.aDates(i)) - nStartScreenX
                    gdSetNum m.geAnnStruct.gdhTop, i + 2, m.aYs(i)
                Else
                    gdSetNum m.geAnnStruct.glhAnnType, i + 2, -1
                End If
            Next
            
        Else
            gdSetNum m.geAnnStruct.glhAnnType, 1, -1
            Exit Sub
        End If
    End If
    
    'set triangle fill color
    gdSetNum m.geAnnStruct.glhFillColor, 0, i       'place holder
    gdSetNum m.geAnnStruct.glhFillColor, 1, Val(Prop("Triangle1_Fill"))
    gdSetNum m.geAnnStruct.glhFillColor, 2, Val(Prop("Triangle2_Fill"))
    gdSetNum m.geAnnStruct.glhFillColor, 3, Val(Prop("Triangle3_Fill"))
    gdSetNum m.geAnnStruct.glhFillColor, 4, Val(Prop("Triangle4_Fill"))
    
    m.geAnnStruct.handleShow = Val(Prop("ShowHandle"))
    
    'set pen colors
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    gdSetNum m.geAnnStruct.glhPenColor, 1, Val(Prop("Triangle1_Color"))
    gdSetNum m.geAnnStruct.glhPenColor, 2, Val(Prop("Triangle2_Color"))
    gdSetNum m.geAnnStruct.glhPenColor, 3, Val(Prop("Triangle3_Color"))
    gdSetNum m.geAnnStruct.glhPenColor, 4, Val(Prop("Triangle4_Color"))
    
    'set values to indicater whether to draw triangles
    gdSetNum m.geAnnStruct.glhImageType, 0, 0           'place holder
    gdSetNum m.geAnnStruct.glhImageType, 1, Val(Prop("Triangle1"))
    gdSetNum m.geAnnStruct.glhImageType, 2, Val(Prop("Triangle2"))
    gdSetNum m.geAnnStruct.glhImageType, 3, Val(Prop("Triangle3"))
    gdSetNum m.geAnnStruct.glhImageType, 4, Val(Prop("Triangle4"))

    'reset
    gdSetSize m.geAnnStruct.gshText, 0, 0
    
    For i = 0 To 4
        gdSetStr m.geAnnStruct.gshText, i, ""
        gdSetNum m.geAnnStruct.gdhMisc, i, -1#      'items [1] thru [4] holds ratios to draw on connecting XA,XD,BC,BD lines
    Next
    
    'set font
    gdSetSize m.geAnnStruct.gshFont, 0, 0
    gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
    gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
    gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
    
    'calculate ratio values
    Dim d1#, d2#, dRatio#
    Dim DB#, dC#, dD#
    Dim dY1#, dY2#
    
    If m.aYs.Size > 0 Then
        If m.bUsingLogDrawingMode Then
            If m.Y1 > 0 Then dY1 = Log(m.Y1)
            If m.Y2 > 0 Then dY2 = Log(m.Y2)
            
            If m.aYs(0) > 0 Then DB = Log(m.aYs(0))
            If m.aYs.Size > 1 And m.aYs(1) > 0 Then dC = Log(m.aYs(1))
            If m.aYs.Size > 2 And m.aYs(2) > 0 Then dD = Log(m.aYs(2))
            
            d1 = Abs(dY2 - dY1)
            d2 = Abs(dY2 - DB)
        Else
            dY1 = m.Y1
            dY2 = m.Y2
            
            DB = m.aYs(0)       'value of point B
            dC = m.aYs(1)       'value of point C
            dD = m.aYs(2)       'value of point D
            
            d1 = Abs(m.Y2 - m.Y1)                   'XA distance
            d2 = Abs(m.Y2 - m.aYs(0))               'AB distance
        End If
        
        'ratios for line XB, XD are based on fib retracement values from XA
        dRatio = Abs((dY2 - DB) / d1)
        gdSetNum m.geAnnStruct.gdhMisc, 1, dRatio
        
        dRatio = Abs((dY2 - dD) / d1)
        gdSetNum m.geAnnStruct.gdhMisc, 2, dRatio
        
        'ratios for line AC, BD are based on fib retracement values from AB
        If d2 > 0 Then
            If m.dDate2 < m.aDates(0) Then
                dRatio = Abs((DB - dC) / d2)
            Else
                dRatio = Abs((dY2 - dC) / d2)
            End If
            gdSetNum m.geAnnStruct.gdhMisc, 3, dRatio
            dRatio = Abs((dY2 - dD) / d2)
            gdSetNum m.geAnnStruct.gdhMisc, 4, dRatio
            
            ' TLB 5/31/2016: looks like the last ratio is supposed to be CD/BC rather than CD/AB
            ' (but Joe Duffy wants the old one as well, so I guess we'll just show both now)
            If dC <> DB Then
                dRatio = Abs((dC - dD) / (dC - DB))
                gdSetNum m.geAnnStruct.gdhMisc, 5, dRatio
            End If
        End If
    End If
    
    Dim Bars As cGdBars
    
    'labels for Gartley end points
    i = Int(Prop("GartleyLabelStyle"))
    If i > 0 Then
        Str = Prop("GartleyLabels")     'do this to get user specified letters (eg X,A,B,C,D)
        
        If Len(Str) > 0 Then
            str1 = Parse(Str, "|", 1)
            If Len(str1) > 1 Then str1 = Parse(str1, "(", 1)    'to get rid if price value if any
            
            Str2 = Parse(Str, "|", 2)
            If Len(Str2) > 1 Then Str2 = Parse(Str2, "(", 1)
            
            str3 = Parse(Str, "|", 3)
            If Len(str3) > 1 Then str3 = Parse(str3, "(", 1)
            
            Str4 = Parse(Str, "|", 4)
            If Len(Str4) > 1 Then Str4 = Parse(Str4, "(", 1)
            
            Str5 = Parse(Str, "|", 5)
            If Len(Str5) > 1 Then Str5 = Parse(Str5, "(", 1)
            
            If i = 2 And Not Chart Is Nothing Then
                Set Bars = Chart.Bars
                If Not Bars Is Nothing Then
                    str1 = str1 & " (" & Bars.PriceDisplay(m.Y1) & ")"
                    Str2 = Str2 & " (" & Bars.PriceDisplay(m.Y2) & ")"
                    str3 = str3 & " (" & Bars.PriceDisplay(m.aYs(0)) & ")"
                    Str4 = Str4 & " (" & Bars.PriceDisplay(m.aYs(1)) & ")"
                    Str5 = Str5 & " (" & Bars.PriceDisplay(m.aYs(2)) & ")"
                End If
            End If
                
            gdSetStr m.geAnnStruct.gshText, 0, str1
            gdSetStr m.geAnnStruct.gshText, 1, Str2
            gdSetStr m.geAnnStruct.gshText, 2, str3
            gdSetStr m.geAnnStruct.gshText, 3, Str4
            gdSetStr m.geAnnStruct.gshText, 4, Str5
            
            'do this to pick up any new y-values if user moved a point
            Prop("GartleyLabels") = str1 & "|" & Str2 & "|" & str3 & "|" & Str4 & "|" & Str5
        End If
    End If

ErrExit:
    Exit Sub


ErrSection:
    RaiseError "cAnnotation.SetGartley"
    
End Sub

Private Function SetGannacciCycle(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:
    
    Dim dY#, iYears&, i&
    Dim eDir As eImageDir
    Dim Pane As cPane
    Dim Bars As cGdBars
    
    'This needs to be done only on initial add to grapheng.dll
    If m.geAdded = False Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 37    'default to image only
        gdSetNum m.geAnnStruct.glhImageType, 0, 6   'default to triangle
        gdSetNum m.geAnnStruct.glhSize, 0, 16       'default to medium icon
        gdSetNum m.geAnnStruct.glhAlign, 0, 5       'default to ctrCtr
        gdSetNum m.geAnnStruct.glhPenStyle, 0, 0    'solid
        gdSetNum m.geAnnStruct.glhPenSize, 0, 1
    End If

    eDir = Val(Prop("ImageDir"))
    If eDir < eCNI_North Or eDir > eCNI_NorthWest Then
        eDir = eCNI_North
        Prop("ImageDir") = eDir
    End If
    
    m.geAnnStruct.lenPoints = Val(Prop("Ext"))      '=1 --> draw line from top to btm of pane
    
    iYears = Val(Prop("GannacciYears"))
    If iYears <= 0 Then iYears = 5
    
    gdSetNum m.geAnnStruct.glhAnnType, 0, 37
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    gdSetNum m.geAnnStruct.glhImageType, 0, 6
    gdSetNum m.geAnnStruct.glhFillPattern, 0, Val(Prop("ImageStyle"))
    gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
    gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
    gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
    geSetPenStyle m.nStyle, 0
    
    gdSetNum m.geAnnStruct.glhAnnType, 1, 37
    gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
    gdSetNum m.geAnnStruct.glhImageType, 1, 6
    gdSetNum m.geAnnStruct.glhFillPattern, 1, Val(Prop("ImageStyle"))
    gdSetStr m.geAnnStruct.gshFont, 1, Prop("FontName")
    gdSetNum m.geAnnStruct.glhSize, 1, Val(Prop("FontSize"))
    gdSetNum m.geAnnStruct.glhStyle, 1, Val(Prop("FontStyle"))
    gdSetNum m.geAnnStruct.glhUnderline, 1, Val(Prop("FontUnderline"))
    geSetPenStyle m.nStyle, 1
    
    gdSetNum m.geAnnStruct.glhDirection, 0, eDir
    If eDir = eCNI_North Then
        If iYears = 5 Then
            gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_South
        Else
            gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_North
        End If
    Else
        If iYears = 5 Then
            gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_North
        Else
            gdSetNum m.geAnnStruct.glhDirection, 1, eCNI_South
        End If
    End If
    
    If Not Chart Is Nothing Then
        Set Pane = Chart.Tree("PRICE PANE")
        If Not Pane Is Nothing Then dY = Pane.Min
        
        Set Bars = Chart.Bars
        m.dDate2 = m.dDate1 + 365 * iYears
        
        i = Bars.FindDateTime(m.dDate2) - Bars.FindDateTime(m.dDate1)
        dX2 = dX1 + i
        
        If dX2 >= Chart.geChartPoints Then dX2 = -1
        
    End If
    
    gdSetStr m.geAnnStruct.gshText, 0, DateFormat(m.dDate2, MM_DD_YYYY) & "|" & Str(iYears)
    gdSetStr m.geAnnStruct.gshText, 1, DateFormat(m.dDate1, MM_DD_YYYY) & "|" & Str(iYears)
    
    If dX1 < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, -1
    Else
        gdSetNum m.geAnnStruct.gdhTop, 0, dY
        gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
        gdSetNum m.geAnnStruct.gdhBottom, 0, dY
        gdSetNum m.geAnnStruct.gdhRight, 0, dX1
    End If
    
    If dX2 < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
    Else
        gdSetNum m.geAnnStruct.gdhTop, 1, dY
        gdSetNum m.geAnnStruct.gdhLeft, 1, dX2
        gdSetNum m.geAnnStruct.gdhBottom, 1, dY
        gdSetNum m.geAnnStruct.gdhRight, 1, dX2
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.SetGannacciCycle"
    
End Function

Private Sub GannacciTimeTB(ByVal nShow&, ByVal idx&, ByVal iSize&, ByVal iColor&, _
    ByVal dDate#, ByVal dX#)
On Error GoTo ErrSection:

    If dX >= iSize Or nShow <> 1 Then
        gdSetNum m.geAnnStruct.glhAnnType, idx, -1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, idx, 39
        gdSetNum m.geAnnStruct.gdhLeft, idx, dX
        gdSetNum m.geAnnStruct.gdhRight, idx, dX
        gdSetStr m.geAnnStruct.gshText, idx, DateFormat(dDate, MM_DD_YYYY)
    
        geSetPenStyle m.nStyle, idx
        gdSetNum m.geAnnStruct.glhPenColor, idx, iColor
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.GannacciTimeTB"
    
End Sub

Private Sub GannacciTimeCD(Chart As cChart, ByVal nShow&, ByVal idx&, ByVal iSize&, ByVal iColor&, ByVal dX#)
On Error GoTo ErrSection:

    Dim i&, dDate#, dCalendar#, strDayName$

    dCalendar = iSize + 1
    
    If nShow = 1 Then
        dDate = m.dDate1 + dX
        
        strDayName = UCase(WeekdayName(dDate))
        If strDayName = "SUN" Then
            dDate = dDate + 1   'bump to Monday
            gdSetNum m.geAnnStruct.glhPenStyle, idx, 1
        ElseIf strDayName = "SAT" Then
            dDate = dDate - 1
            gdSetNum m.geAnnStruct.glhPenStyle, idx, 1
        Else
            geSetPenStyle m.nStyle, idx
        End If
        
        If dDate < gdGetNum(Chart.geDateArray, iSize - 1) Then
            For i = 0 To iSize - 1
                If gdGetNum(Chart.geDateArray, i) = dDate Then
                    dCalendar = i
                    Exit For
                End If
            Next
        End If
    End If

    If dCalendar >= iSize Then
        gdSetNum m.geAnnStruct.glhAnnType, idx, -1
    Else
        gdSetNum m.geAnnStruct.glhAnnType, idx, 39
        gdSetNum m.geAnnStruct.gdhLeft, idx, dCalendar
        gdSetNum m.geAnnStruct.gdhRight, idx, dCalendar
        gdSetStr m.geAnnStruct.gshText, idx, DateFormat(m.dDate1 + dX, MM_DD_YYYY)
    
        gdSetNum m.geAnnStruct.glhPenColor, idx, iColor
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.GannacciTimeCD"
    
End Sub

Private Function SetGannacciTime(Chart As cChart, ByVal dX1#, ByVal dX2#) As Long
On Error GoTo ErrSection:

    Dim i&, iSize&, hArray&, dDate#, dX#
    Dim Bars As cGdBars
    Dim Pane As cPane
    
    If Not Chart Is Nothing Then
        Set Bars = Chart.Bars
        Set Pane = Chart.Tree("PRICE PANE")
    End If
    
    If Bars Is Nothing Then Exit Function
    If Pane Is Nothing Then Exit Function
    
    hArray = Chart.geDateArray
    If hArray = 0 Then Exit Function
    
    m.Y2 = Pane.geAdjustMin
    If Val(Prop("Ext")) = 1 Then
        m.Y1 = Pane.Max
    Else
        m.Y1 = m.Y2 + (Pane.Max - m.Y2) * 0.15
    End If
    
    gdSetNum m.geAnnStruct.glhAnnType, 0, 39
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhRight, 0, dX1
    gdSetStr m.geAnnStruct.gshText, 0, DateFormat(m.dDate1, MM_DD_YYYY)
    
    geSetPenStyle m.nStyle, 0
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    
    m.geAnnStruct.lenPoints = Val(Prop("Ext"))
    
    'set Trading Bars lines
    iSize = gdGetSize(hArray)
    If Val(Prop("ShowTB")) = 1 Then
        dX = dX1 + 45
        GannacciTimeTB Val(Prop("Show45")), 1, iSize, Val(Prop("ColorFor45")), gdGetNum(hArray, dX), dX
        
        dX = dX1 + 90
        GannacciTimeTB Val(Prop("Show90")), 2, iSize, Val(Prop("ColorFor90")), gdGetNum(hArray, dX), dX
        
        dX = dX1 + 180
        GannacciTimeTB Val(Prop("Show180")), 3, iSize, Val(Prop("ColorFor180")), gdGetNum(hArray, dX), dX
        
        dX = dX1 + 270
        GannacciTimeTB Val(Prop("Show270")), 4, iSize, Val(Prop("ColorFor270")), gdGetNum(hArray, dX), dX
        
        dX = dX1 + 360
        GannacciTimeTB Val(Prop("Show360")), 5, iSize, Val(Prop("ColorFor360")), gdGetNum(hArray, dX), dX
    Else
        GannacciTimeTB 0, 1, iSize, m.nColor, m.dDate1, dX1     'values don't really matter since not shown
        GannacciTimeTB 0, 2, iSize, m.nColor, m.dDate1, dX1
        GannacciTimeTB 0, 3, iSize, m.nColor, m.dDate1, dX1
        GannacciTimeTB 0, 4, iSize, m.nColor, m.dDate1, dX1
        GannacciTimeTB 0, 5, iSize, m.nColor, m.dDate1, dX1
    End If
    
    'set Calendar Days Lines
    If Chart.Bars.Prop(eBARS_PeriodicityStr) = "Daily" And Val(Prop("ShowCD")) = 1 Then
        GannacciTimeCD Chart, Val(Prop("Show45")), 6, iSize, Val(Prop("ColorFor45")), 45
        GannacciTimeCD Chart, Val(Prop("Show90")), 7, iSize, Val(Prop("ColorFor90")), 90
        GannacciTimeCD Chart, Val(Prop("Show180")), 8, iSize, Val(Prop("ColorFor180")), 180
        GannacciTimeCD Chart, Val(Prop("Show270")), 9, iSize, Val(Prop("ColorFor270")), 270
        GannacciTimeCD Chart, Val(Prop("Show360")), 10, iSize, Val(Prop("ColorFor360")), 360
    Else
        GannacciTimeCD Chart, 0, 6, iSize, m.nColor, dX1    'values don't really matter since not shown
        GannacciTimeCD Chart, 0, 7, iSize, m.nColor, dX1
        GannacciTimeCD Chart, 0, 8, iSize, m.nColor, dX1
        GannacciTimeCD Chart, 0, 9, iSize, m.nColor, dX1
        GannacciTimeCD Chart, 0, 10, iSize, m.nColor, dX1
    End If
    
    'info for all lines
    For i = 0 To 10
        gdSetNum m.geAnnStruct.gdhTop, i, m.Y1
        gdSetNum m.geAnnStruct.gdhBottom, i, m.Y2
        gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
        gdSetNum m.geAnnStruct.glhAlign, i, 10      'vertical text
    Next
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.SetGannacciTime"
    
End Function

Public Property Get AllowCalendarDays() As Boolean
On Error GoTo ErrSection:

    If m.eType = eANNOT_GannacciTime Then
        If Not m.Chart Is Nothing Then
            If Not m.Chart.Bars Is Nothing Then
                If m.Chart.Bars.Prop(eBARS_PeriodicityStr) = "Daily" Then
                    AllowCalendarDays = True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.AllowCalendarDays"
    
End Property

Private Function SetGannacciSwing(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&) As Long
On Error GoTo ErrSection:

    Dim rc&
    Dim X#, Y#, XDelta#, YDelta#, dFudge#
    
    rc = -1
    
    If Chart Is Nothing Then GoTo ErrExit
    If Chart.Bars Is Nothing Then GoTo ErrExit
    
    ' setup
    If MultiPointSetup(Chart, XDelta, YDelta, dFudge) = False Then
        SetGannacciSwing = 0
        gdSetNum m.geAnnStruct.glhAnnType, 0, -1
        Exit Function
    End If

    If m.X1 > Chart.ScreenEndX Then
        SetPointOffScreen Chart, True, X, Y
    ElseIf m.X1 < Chart.ScreenStartX Then
        SetPointOffScreen Chart, False, X, Y
    Else
        X = m.X1
        Y = m.Y1
    End If
                
    If m.X2 > Chart.ScreenEndX Then
        SetPointOffScreen Chart, True, X, Y
    ElseIf m.X2 < Chart.ScreenStartX Then
        SetPointOffScreen Chart, False, X, Y
    Else
        X = m.X2
        Y = m.Y2
    End If
    If X = m.X1 Then X = X + kXfudge
    If Y = m.Y1 Then Y = Y + dFudge
    
    m.geAnnStruct.lenPoints = Val(Prop("Border"))
    
    If m.eType = eANNOT_GannacciSwing1 Then
        rc = SetGannacciSwing1(Chart, dX1, dX2, nStartScreenX)
    ElseIf m.eType = eANNOT_GannacciSwing2 Then
        rc = SetGannacciSwing2(Chart, dX1, dX2, nStartScreenX)
    End If
    
ErrExit:
    SetGannacciSwing = rc
    Exit Function

ErrSection:
    RaiseError "cAnnotation.SetGannacciSwing"

End Function

Private Function SetGannacciSwing2(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&) As Long
On Error GoTo ErrSection:

    Dim i&, iSize&, dX3#
    Dim iColorLow&, iColorMedium&, iColorHigh&, iColor&
    
    Dim aText As New cGdArray
    
    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 38
        gdSetStr m.geAnnStruct.gshText, 0, ""
        gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
        gdSetNum m.geAnnStruct.glhFillPattern, 0, 0
    End If
    
    aText(0) = ""       'place holder (so can start for loop with i=1)
'group 1: R1, R2
    If Val(Prop("R1")) = 1 Then
        aText(eGannci_R1) = "R1="
    Else
        aText(eGannci_R1) = "-1"
    End If
    
    If Val(Prop("R2")) = 1 Then
        aText(eGannci_R2) = "R2="
    Else
        aText(eGannci_R2) = "-1"
    End If

'group 2: trading bars
    If Val(Prop("TB1")) = 1 Then
        aText(eGannci_TB1) = "TB1="
    Else
        aText(eGannci_TB1) = "-1"
    End If
    
    If Val(Prop("TB2")) = 1 Then
        aText(eGannci_TB2) = "TB2="
    Else
        aText(eGannci_TB2) = "-1"
    End If
    
    If Val(Prop("Diff_TB1_R2")) = 1 Then
        aText(eGannci_TB1_R2) = "TB1-R2="
    Else
        aText(eGannci_TB1_R2) = "-1"
    End If
    
    If Val(Prop("Diff_TB2_R1")) = 1 Then
        aText(eGannci_TB2_R1) = "TB2-R1="
    Else
        aText(eGannci_TB2_R1) = "-1"
    End If
    
    If Val(Prop("Equal_TB1_R2")) = 1 Then
        aText(eGannci_EQ_TB1_R2) = "TB1=R2"
    Else
        aText(eGannci_EQ_TB1_R2) = "-1"
    End If
    
    If Val(Prop("Equal_TB2_R1")) = 1 Then
        aText(eGannci_EQ_TB2_R1) = "TB2=R1"
    Else
        aText(eGannci_EQ_TB2_R1) = "-1"
    End If
    
'group 3: calendar days
    If Val(Prop("CD1")) = 1 Then
        aText(eGannci_CD1) = "CD1="
    Else
        aText(eGannci_CD1) = "-1"
    End If
    
    If Val(Prop("CD2")) = 1 Then
        aText(eGannci_CD2) = "CD2="
    Else
        aText(eGannci_CD2) = "-1"
    End If
    
    If Val(Prop("Diff_CD1_R2")) = 1 Then
        aText(eGannci_CD1_R2) = "CD1-R2="
    Else
        aText(eGannci_CD1_R2) = "-1"
    End If
    
    If Val(Prop("Diff_CD2_R1")) = 1 Then
        aText(eGannci_CD2_R1) = "CD2-R1="
    Else
        aText(eGannci_CD2_R1) = "-1"
    End If
    
    If Val(Prop("Equal_CD1_R2")) = 1 Then
        aText(eGannci_EQ_CD1_R2) = "CD1=R2"
    Else
        aText(eGannci_EQ_CD1_R2) = "-1"
    End If
    
    If Val(Prop("Equal_CD2_R1")) = 1 Then
        aText(eGannci_EQ_CD2_R1) = "CD2=R1"
    Else
        aText(eGannci_EQ_CD2_R1) = "-1"
    End If

    iColorLow = Prop("SignalColorLow")
    iColorMedium = Prop("SignalColorMedium")
    iColorHigh = Prop("SignalColorHigh")
    
    Dim iShowText&, iShowMarker&
    
    iShowText = Val(Prop("ShowText"))
    iShowMarker = Val(Prop("ShowMarker"))
    
    iSize = m.aRatioColor.Size
    If iSize > 14 Then iSize = 14       'theoretically should never happen
    
    If iShowMarker = 1 Then
        gdSetNum m.geAnnStruct.glhImageType, 0, eCNI_Circle
        m.geAnnStruct.showAxes = 2      'let grapheng.dll know starting index for drawing markers
    Else
        gdSetSize m.geAnnStruct.glhImageType, 0, 1
        m.geAnnStruct.showAxes = -1
    End If
    
    'coordinates for lines
    gdSetNum m.geAnnStruct.glhAnnType, 0, 38
    geSetPenStyle m.nStyle, 0
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    
    m.geAnnStruct.lenRatio = 0                  '0=draw text below bar's low, 1=draw text above bar's high
    
    If m.aDates.Size > 0 Then
        CalcGannacciSwing2 iColorLow, iColorMedium, iColorHigh
        dX3 = GetX(Chart, m.aDates(0)) - nStartScreenX           '4525
        
        If dX3 > dX1 Then
            gdSetNum m.geAnnStruct.glhAnnType, 0, 1
            gdSetNum m.geAnnStruct.glhAnnType, 1, 38
            If m.aYs(0) = Chart.Bars(eBARS_High, GetX(m.Chart, m.aDates(0))) Then m.geAnnStruct.lenRatio = 1
        Else
            gdSetNum m.geAnnStruct.glhAnnType, 0, 38
            gdSetNum m.geAnnStruct.glhAnnType, 1, 1
            If m.Y1 = Chart.Bars(eBARS_High, m.X1) Then m.geAnnStruct.lenRatio = 1
        End If
            
        geSetPenStyle m.nStyle, 1
        gdSetNum m.geAnnStruct.glhPenColor, 1, m.nColor
        gdSetNum m.geAnnStruct.gdhTop, 1, m.Y2
        gdSetNum m.geAnnStruct.gdhLeft, 1, dX2
        gdSetNum m.geAnnStruct.gdhBottom, 1, m.aYs(0)
        gdSetNum m.geAnnStruct.gdhRight, 1, dX3
    Else
        gdSetNum m.geAnnStruct.glhAnnType, 1, -1
    
        If m.X1 < m.X2 Then
            If m.Y2 = Chart.Bars(eBARS_High, m.X2) Then m.geAnnStruct.lenRatio = 1
        ElseIf m.Y1 = Chart.Bars(eBARS_High, m.X1) Then
            m.geAnnStruct.lenRatio = 1
        End If
        
        GoTo ErrExit
    End If
    
    'place holders
    gdSetStr m.geAnnStruct.gshText, 0, ""
    gdSetStr m.geAnnStruct.gshText, 1, ""
    
    For i = 1 To iSize - 1
        If iShowText = 1 Or iShowMarker = 1 Then
            iColor = m.aRatioColor(i)
            
            If iColor = iColorLow Or iColor = iColorMedium Or iColor = iColorHigh Or iColor = m.nColor Then
                If iShowText = 1 And aText(i) <> "-1" Then
                    Select Case i
                        Case eGannci_EQ_TB1_R2, eGannci_EQ_TB2_R1, eGannci_EQ_CD1_R2, eGannci_EQ_CD2_R1
                            gdSetStr m.geAnnStruct.gshText, i + 1, aText(i)
                        Case Else
                            gdSetStr m.geAnnStruct.gshText, i + 1, aText(i) & Str(m.aRatios(i))
                    End Select
                Else
                    gdSetStr m.geAnnStruct.gshText, i + 1, ""
                End If
            
                If aText(i) = "-1" Then
                    gdSetNum m.geAnnStruct.glhPenColor, i + 1, -1
                Else
                    gdSetNum m.geAnnStruct.glhPenColor, i + 1, iColor
                End If

                gdSetStr m.geAnnStruct.gshFont, i + 1, Prop("FontName")
                gdSetNum m.geAnnStruct.glhSize, i + 1, Val(Prop("FontSize"))
                gdSetNum m.geAnnStruct.glhStyle, i + 1, Val(Prop("FontStyle"))
                gdSetNum m.geAnnStruct.glhUnderline, i + 1, Val(Prop("FontUnderline"))
                gdSetNum m.geAnnStruct.glhFillPattern, i + 1, 0
            Else
                gdSetStr m.geAnnStruct.gshText, i + 1, ""
                gdSetNum m.geAnnStruct.glhPenColor, i + 1, -1
            End If
        Else
            gdSetStr m.geAnnStruct.gshText, i + 1, ""
            gdSetNum m.geAnnStruct.glhPenColor, i + 1, -1
        End If
    Next

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.SetGannacciSwing2"

End Function

Private Function SetGannacciSwing1(Chart As cChart, ByVal dX1#, ByVal dX2#, ByVal nStartScreenX&) As Long
On Error GoTo ErrSection:

    Dim i&, iSize&
    Dim iColorLow&, iColorMedium&, iColorHigh&, iColor&
    
    Dim aText As New cGdArray

    i = gdGetNum(m.geAnnStruct.glhAnnType, 0)
    If m.geAdded = False Or i < 0 Then
        gdSetNum m.geAnnStruct.glhAnnType, 0, 38
        gdSetStr m.geAnnStruct.gshText, 0, ""
        gdSetStr m.geAnnStruct.gshFont, 0, Prop("FontName")
        gdSetNum m.geAnnStruct.glhSize, 0, Val(Prop("FontSize"))
        gdSetNum m.geAnnStruct.glhStyle, 0, Val(Prop("FontStyle"))
        gdSetNum m.geAnnStruct.glhUnderline, 0, Val(Prop("FontUnderline"))
        gdSetNum m.geAnnStruct.glhFillPattern, 0, 0
    End If
    
    aText(0) = ""       'place holder (so can start for loop with i=1)
    
'group 1: P1,P2,R
    aText(eGannci_P1) = "-1"
    aText(eGannci_P2) = "-1"
    
    If Val(Prop("R")) = 1 Then
        aText(eGannci_R) = "R="
    Else
        aText(eGannci_R) = "-1"
    End If
    
'group 2: trade bars
    aText(eGannci_TB) = "-1"
    aText(eGannci_TB_P1) = "-1"
    
    If Val(Prop("Diff_TBP2")) = 1 Then
        aText(eGannci_TB_P2) = "TB - P2="
    Else
        aText(eGannci_TB_P2) = "-1"
    End If
    
    If Val(Prop("Diff_TBR")) = 1 Then
        aText(eGannci_TB_R) = "TB - R="
    Else
        aText(eGannci_TB_R) = "-1"
    End If
    
    If Val(Prop("Equal_TBP1")) = 1 Then
        aText(eGannci_EQ_TB_P1) = "TB=P1"
    Else
        aText(eGannci_EQ_TB_P1) = "-1"
    End If
        
    If Val(Prop("Equal_TBP2")) = 1 Then
        aText(eGannci_EQ_TB_P2) = "TB=P2"
    Else
        aText(eGannci_EQ_TB_P2) = "-1"
    End If

    If Val(Prop("Equal_TBR")) = 1 Then
        aText(eGannci_EQ_TB_R) = "TB=R"
    Else
        aText(eGannci_EQ_TB_R) = "-1"
    End If
    
'group 3: calendar days
    aText(eGannci_CD) = "-1"
    aText(eGannci_CD_P1) = "-1"
    
    If Val(Prop("Diff_CDP2")) = 1 Then
        aText(eGannci_CD_P2) = "CD - P2="
    Else
        aText(eGannci_CD_P2) = "-1"
    End If
    
    If Val(Prop("Diff_CDR")) = 1 Then
        aText(eGannci_CD_R) = "CD - R="
    Else
        aText(eGannci_CD_R) = "-1"
    End If
    
    If Val(Prop("Equal_CDP1")) = 1 Then
        aText(eGannci_EQ_CD_P1) = "CD=P1"
    Else
        aText(eGannci_EQ_CD_P1) = "-1"
    End If
    
    If Val(Prop("Equal_CDP2")) = 1 Then
        aText(eGannci_EQ_CD_P2) = "CD=P2"
    Else
        aText(eGannci_EQ_CD_P2) = "-1"
    End If
    
    If Val(Prop("Equal_CDR")) = 1 Then
        aText(eGannci_EQ_CD_R) = "CD=R"
    Else
        aText(eGannci_EQ_CD_R) = "-1"
    End If
    
    iColorLow = Prop("SignalColorLow")
    iColorMedium = Prop("SignalColorMedium")
    iColorHigh = Prop("SignalColorHigh")
    
    CalcGannacciSwing1 iColorLow, iColorMedium, iColorHigh
    
    'line
    gdSetNum m.geAnnStruct.glhAnnType, 0, 38
    geSetPenStyle m.nStyle, 0
    gdSetNum m.geAnnStruct.glhPenColor, 0, m.nColor
    gdSetNum m.geAnnStruct.gdhTop, 0, m.Y1
    gdSetNum m.geAnnStruct.gdhLeft, 0, dX1
    gdSetNum m.geAnnStruct.gdhBottom, 0, m.Y2
    gdSetNum m.geAnnStruct.gdhRight, 0, dX2
    
    m.geAnnStruct.lenRatio = 0
    If m.X1 < m.X2 Then
        If m.Y2 = Chart.Bars(eBARS_High, m.X2) Then m.geAnnStruct.lenRatio = 1
    ElseIf m.Y1 = Chart.Bars(eBARS_High, m.X1) Then
        m.geAnnStruct.lenRatio = 1
    End If
    
    Dim iShowText&, iShowMarker&
    
    iShowText = Val(Prop("ShowText"))
    iShowMarker = Val(Prop("ShowMarker"))
    
    iSize = m.aRatioColor.Size
    If iSize > 18 Then iSize = 18       'theoretically should never happen
    
    If iShowMarker = 1 Then
        gdSetNum m.geAnnStruct.glhImageType, 0, eCNI_Circle
        m.geAnnStruct.showAxes = 1      'let grapheng.dll know starting index for drawing markers
    Else
        gdSetSize m.geAnnStruct.glhImageType, 0, 1
        m.geAnnStruct.showAxes = -1
    End If
    
    For i = 1 To iSize - 1
        If iShowText = 1 Or iShowMarker = 1 Then
            iColor = m.aRatioColor(i)
            
            If iColor = iColorLow Or iColor = iColorMedium Or iColor = iColorHigh Or iColor = m.nColor Then
                If iShowText = 1 And aText(i) <> "-1" Then
                    Select Case i
                        Case eGannci_EQ_TB_P1, eGannci_EQ_TB_P2, eGannci_EQ_TB_R, _
                             eGannci_EQ_CD_P1, eGannci_EQ_CD_P2, eGannci_EQ_CD_R
                             
                             gdSetStr m.geAnnStruct.gshText, i, aText(i)
                             
                        Case Else
                            gdSetStr m.geAnnStruct.gshText, i, aText(i) & Str(m.aRatios(i))
                    
                    End Select
                    
                    gdSetNum m.geAnnStruct.glhPenColor, i, iColor
                Else
                    gdSetStr m.geAnnStruct.gshText, i, ""
                End If
                
                If aText(i) = "-1" Then
                    gdSetNum m.geAnnStruct.glhPenColor, i, -1
                Else
                    gdSetNum m.geAnnStruct.glhPenColor, i, iColor
                End If
                
                gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
                gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
                gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
                gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
                gdSetNum m.geAnnStruct.glhFillPattern, i, 0
            
                gdSetNum m.geAnnStruct.gdhTop, i, m.Y2
                gdSetNum m.geAnnStruct.gdhLeft, i, dX2
                gdSetNum m.geAnnStruct.gdhBottom, i, m.Y2
                gdSetNum m.geAnnStruct.gdhRight, i, dX2
            Else
                gdSetStr m.geAnnStruct.gshText, i, ""
                gdSetNum m.geAnnStruct.glhPenColor, i, -1
            End If
        Else
            gdSetStr m.geAnnStruct.gshText, i, ""
            gdSetNum m.geAnnStruct.glhPenColor, i, -1
        End If
    Next

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.SetGannacciSwing1"

End Function

Private Function MoveGannacciSwing(Chart As cChart, ByVal nPoint&, _
        ByVal X#, ByVal Y#) As Boolean
On Error GoTo ErrSection:

    Dim bOkay As Boolean
    
    If Chart Is Nothing Then Exit Function

    bOkay = True
    
'    StatusMsg "nPoint = " & str(nPoint)
    
    Select Case nPoint
        Case 2
            m.dDate2 = X
            m.Y2 = Y
            m.X2 = GetX(Chart, X)
        Case 1
            m.dDate1 = X
            m.Y1 = Y
            m.X1 = GetX(Chart, X)
        
        Case 3
            If m.eType = eANNOT_GannacciSwing2 Then
                m.aDates(0) = X
                m.aYs(0) = Y
            Else
                bOkay = False
            End If
        
        Case Else
            bOkay = False
    End Select
    
ErrExit:
    MoveGannacciSwing = bOkay
    Exit Function

ErrSection:
    RaiseError "cAnnotation.MoveGannacciSwing"

End Function

Private Sub CalcGannacciSwing2(ByVal iColorLow&, ByVal iColorMedium&, ByVal iColorHigh&)
On Error GoTo ErrSection:

    Dim i&, iColor&, iExtra&
    Dim d#, X1#, X2#, X3#, dY1#, dY2#, dY3#
    Dim dDate1#, dDate2#, dDate3#
    
    Dim Bars As cGdBars
    
    If m.Chart Is Nothing Then Exit Sub
    
    Set Bars = m.Chart.Bars
    If Bars Is Nothing Then Exit Sub
    
    If Val(Prop("IncludeBarOne")) = 1 Then iExtra = 1

    m.aRatios.Clear             'holds calculated value
    m.aRatioShow.Clear          'holds 0/1 for whether to display
    m.aRatioColor.Clear         'holds color coded display info
    
    m.aRatios(0) = 0#           'place holders
    m.aRatioShow(0) = 0
    m.aRatioColor(0) = 0
    
    If Prop("UseMutiplier") = 1 Then
        d = Val(Prop("MultiplierVal"))
        dY1 = m.Y1 * d
        dY2 = m.Y2 * d
        If m.aYs(0) <> kNullData Then dY3 = m.aYs(0) * d
        
        '05-04-2012: design change per Pete
        '   - Nick wants prices rounded to specified decimal places before calculations
        If Val(Prop("RoundDecimals") = 1) Then
            i = Val(Prop("Decimals"))
            If i >= 0 Then
                dY1 = Round(dY1, i)
                dY2 = Round(dY2, i)
                If m.aYs(0) <> kNullData Then dY3 = Round(dY3, i)
            End If
        End If
    Else
        dY1 = m.Y1
        dY2 = m.Y2
        If m.aYs(0) <> kNullData Then
            dY3 = m.aYs(0)
        End If
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' fix endpoints order in case user drew from right->left
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If m.X1 < m.X2 Then
        X1 = m.X1
        X2 = m.X2
        
        dDate1 = m.dDate1
        dDate2 = m.dDate2
        
        If m.aDates.Size > 0 Then
            X3 = GetX(m.Chart, m.aDates(0))
            dDate3 = m.aDates(0)
        End If
    ElseIf m.aYs(0) <> kNullData Then
        X1 = GetX(m.Chart, m.aDates(0))
        X2 = m.X2
        X3 = m.X1
        
        d = dY1
        dY1 = dY3
        dY3 = d
        
        dDate1 = m.aDates(0)
        dDate2 = m.dDate2
        dDate3 = m.dDate1
    Else
        X1 = m.X2               'user have only drawn the first leg going right->left
        X2 = m.X1
        d = dY1
        dY1 = dY2
        dY2 = d
        dDate1 = m.dDate2
        dDate2 = m.dDate1
    End If

'group 1: R1, R2
    i = GannacciFormula(dY1, dY2)                       'R1
    m.aRatios(eGannci_R1) = i
    m.aRatioColor(eGannci_R1) = GannacciSwingColor(i)

    i = GannacciFormula(dY2, dY3)                       'R2
    m.aRatios(eGannci_R2) = i
    m.aRatioColor(eGannci_R2) = GannacciSwingColor(i)


'group 2: trading bars
    i = GannacciFormula(X1, X2, iExtra, True)           'TB1 - trading bars count
    m.aRatios(eGannci_TB1) = i
    m.aRatioColor(eGannci_TB1) = GannacciSwingColor(i)

    If m.aDates.Size > 0 Then
        i = GannacciFormula(X2, X3, iExtra, True)       'TB2 - trading bars count
        m.aRatios(eGannci_TB2) = i
        m.aRatioColor(eGannci_TB2) = GannacciSwingColor(i)
    Else
        m.aRatios(eGannci_TB2) = -1
        m.aRatioColor(eGannci_TB2) = -1
    End If

    i = Abs(m.aRatios(eGannci_TB1) - m.aRatios(eGannci_R2))                'TB1 - R2
    m.aRatios(eGannci_TB1_R2) = i
    m.aRatioColor(eGannci_TB1_R2) = GannacciSwingColor(i)

    i = Abs(m.aRatios(eGannci_TB2) - m.aRatios(eGannci_R1))                'TB2 - R1
    m.aRatios(eGannci_TB2_R1) = i
    m.aRatioColor(eGannci_TB2_R1) = GannacciSwingColor(i)

    i = Abs(m.aRatios(eGannci_TB1) - m.aRatios(eGannci_R2))                 'TB1 = R2
    i = GannacciEqualColor(i)
    m.aRatios(eGannci_EQ_TB1_R2) = i
    m.aRatioColor(eGannci_EQ_TB1_R2) = i
    
    i = Abs(m.aRatios(eGannci_TB2) - m.aRatios(eGannci_R1))                 'TB2 = R1
    i = GannacciEqualColor(i)
    m.aRatios(eGannci_EQ_TB2_R1) = i
    m.aRatioColor(eGannci_EQ_TB2_R1) = i

'group 3: calendar days
    If Bars.Prop(eBARS_PeriodicityStr) = "Daily" Then
        i = GannacciFormula(dDate1, dDate2, iExtra)         'CD1 - calendar days count
        m.aRatios(eGannci_CD1) = i
        m.aRatioColor(eGannci_CD1) = GannacciSwingColor(i)
    
        If m.aDates.Size > 0 Then
            i = GannacciFormula(dDate2, dDate3, iExtra)     'CD2 - calendar days count
            m.aRatios(eGannci_CD2) = i
            m.aRatioColor(eGannci_CD2) = GannacciSwingColor(i)
        Else
            m.aRatios(eGannci_CD2) = -1
            m.aRatioColor(eGannci_CD2) = -1
        End If
        
        i = Abs(m.aRatios(eGannci_CD1) - m.aRatios(eGannci_R2))                'CD1 - R2
        m.aRatios(eGannci_CD1_R2) = i
        m.aRatioColor(eGannci_CD1_R2) = GannacciSwingColor(i)
    
        i = Abs(m.aRatios(eGannci_CD2) - m.aRatios(eGannci_R1))                'CD2 - R1
        m.aRatios(eGannci_CD2_R1) = i
        m.aRatioColor(eGannci_CD2_R1) = GannacciSwingColor(i)
    
        i = Abs(m.aRatios(eGannci_CD1) - m.aRatios(eGannci_R2))                 'CD1 = R2
        i = GannacciEqualColor(i)
        m.aRatios(eGannci_EQ_CD1_R2) = i
        m.aRatioColor(eGannci_EQ_CD1_R2) = i
    
        i = Abs(m.aRatios(eGannci_CD2) - m.aRatios(eGannci_R1))                 'CD2 = R1
        i = GannacciEqualColor(i)
        m.aRatios(eGannci_EQ_CD2_R1) = i
        m.aRatioColor(eGannci_EQ_CD2_R1) = i
    Else
        m.aRatios(eGannci_CD1) = -1
        m.aRatioColor(eGannci_CD1) = -1
    
        m.aRatios(eGannci_CD2) = -1
        m.aRatioColor(eGannci_CD2) = -1
    
        m.aRatios(eGannci_CD1_R2) = -1
        m.aRatioColor(eGannci_CD1_R2) = -1
    
        m.aRatios(eGannci_CD2_R1) = -1
        m.aRatioColor(eGannci_CD2_R1) = -1
    
        m.aRatios(eGannci_EQ_CD1_R2) = -1
        m.aRatioColor(eGannci_EQ_CD1_R2) = -1
    
        m.aRatios(eGannci_EQ_CD2_R1) = -1
        m.aRatioColor(eGannci_EQ_CD2_R1) = -1
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.CalcGannacciSwing2"

End Sub


Private Sub CalcGannacciSwing1(ByVal iColorLow&, ByVal iColorMedium&, ByVal iColorHigh&)
On Error GoTo ErrSection:

    Dim i&, iColor&, iExtra&
    Dim d#, dY1#, dY2#
    
    Dim Bars As cGdBars
    
    If m.Chart Is Nothing Then Exit Sub
    
    Set Bars = m.Chart.Bars
    If Bars Is Nothing Then Exit Sub
    
    If Val(Prop("IncludeBarOne")) = 1 Then iExtra = 1

    m.aRatios.Clear             'holds calculated value
    m.aRatioShow.Clear          'holds 0/1 for whether to display
    m.aRatioColor.Clear         'holds color coded display info
    
    m.aRatios(0) = 0#           'place holders
    m.aRatioShow(0) = 0
    m.aRatioColor(0) = 0
    
    If Prop("UseMutiplier") = 1 Then
        d = Val(Prop("MultiplierVal"))
        dY1 = m.Y1 * d
        dY2 = m.Y2 * d
        
        '05-04-2012: design change per Pete
        '   - Nick wants prices rounded to specified decimal places before calculations
        If Val(Prop("RoundDecimals") = 1) Then
            i = Val(Prop("Decimals"))
            If i >= 0 Then
                dY1 = Round(dY1, i)
                dY2 = Round(dY2, i)
            End If
        End If
    Else
        dY1 = m.Y1
        dY2 = m.Y2
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Group 1: price 1, price 2, price range
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'P1/P2  - make P1 price of left-most endpoint
    If m.X2 > m.X1 Then
        i = GannacciFormula(dY1, kNullData)
        m.aRatios(eGannci_P1) = i
        m.aRatioColor(eGannci_P1) = GannacciSwingColor(i)
        
        i = GannacciFormula(dY2, kNullData)
        m.aRatios(eGannci_P2) = i
        m.aRatioColor(eGannci_P2) = GannacciSwingColor(i)
    Else
        i = GannacciFormula(dY2, kNullData)
        m.aRatios(eGannci_P1) = i
        m.aRatioColor(eGannci_P1) = GannacciSwingColor(i)
        
        i = GannacciFormula(dY1, kNullData)
        m.aRatios(eGannci_P2) = i
        m.aRatioColor(eGannci_P2) = GannacciSwingColor(i)
    End If
    'R - range
    i = GannacciFormula(dY1, dY2)
    m.aRatios(eGannci_R) = i
    m.aRatioColor(eGannci_R) = GannacciSwingColor(i)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Group 2: trade bars
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    i = GannacciFormula(m.X1, m.X2, iExtra, True)                   'TB: trade bars count
    m.aRatios(eGannci_TB) = i
    m.aRatioColor(eGannci_TB) = GannacciSwingColor(i)

    'trade bars minus P1,P2,R
    i = Round(Abs(m.aRatios(eGannci_TB) - m.aRatios(eGannci_P1)))   'TB - P1: trade bars minus price 1
    m.aRatios(eGannci_TB_P1) = i
    m.aRatioColor(eGannci_TB_P1) = GannacciSwingColor(i)

    i = Round(Abs(m.aRatios(eGannci_TB) - m.aRatios(eGannci_P2)))   'TB - P2: trade bars minus price 2
    m.aRatios(eGannci_TB_P2) = i
    m.aRatioColor(eGannci_TB_P2) = GannacciSwingColor(i)

    i = Round(Abs(m.aRatios(eGannci_TB) - m.aRatios(eGannci_R)))    'TB - R: trade bars minus range
    m.aRatios(eGannci_TB_R) = i
    m.aRatioColor(eGannci_TB_R) = GannacciSwingColor(i)

    'trade bars equals P1,P2,R
    i = Abs(m.aRatios(eGannci_TB) - m.aRatios(eGannci_P1))
    i = GannacciEqualColor(i)
    m.aRatios(eGannci_EQ_TB_P1) = i
    m.aRatioColor(eGannci_EQ_TB_P1) = i

    i = Abs(m.aRatios(eGannci_TB) - m.aRatios(eGannci_P2))
    i = GannacciEqualColor(i)
    m.aRatios(eGannci_EQ_TB_P2) = i
    m.aRatioColor(eGannci_EQ_TB_P2) = i
    
    i = Abs(m.aRatios(eGannci_TB) - m.aRatios(eGannci_R))
    i = GannacciEqualColor(i)
    m.aRatios(eGannci_EQ_TB_R) = i
    m.aRatioColor(eGannci_EQ_TB_R) = i

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Group 3: calendar days
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Bars.Prop(eBARS_PeriodicityStr) = "Daily" Then
        i = GannacciFormula(m.dDate1, m.dDate2, iExtra)     'CD - calendar days count
        m.aRatios(eGannci_CD) = i
        m.aRatioColor(eGannci_CD) = GannacciSwingColor(i)
    
        'calendar days minus P1,P2,R
        i = Round(Abs(m.aRatios(eGannci_CD) - m.aRatios(eGannci_P1)))
        m.aRatios(eGannci_CD_P1) = i
        m.aRatioColor(eGannci_CD_P1) = GannacciSwingColor(i)
    
        i = Round(Abs(m.aRatios(eGannci_CD) - m.aRatios(eGannci_P2)))
        m.aRatios(eGannci_CD_P2) = i
        m.aRatioColor(eGannci_CD_P2) = GannacciSwingColor(i)
    
        i = Round(Abs(m.aRatios(eGannci_CD) - m.aRatios(eGannci_R)))
        m.aRatios(eGannci_CD_R) = i
        m.aRatioColor(eGannci_CD_R) = GannacciSwingColor(i)
    
        'calendar days equals P1,P2,R
        i = Abs(m.aRatios(eGannci_CD) - m.aRatios(eGannci_P1))
        i = GannacciEqualColor(i)
        m.aRatios(eGannci_EQ_CD_P1) = i
        m.aRatioColor(eGannci_EQ_CD_P1) = i
    
        i = Abs(m.aRatios(eGannci_CD) - m.aRatios(eGannci_P2))
        i = GannacciEqualColor(i)
        m.aRatios(eGannci_EQ_CD_P2) = i
        m.aRatioColor(eGannci_EQ_CD_P2) = i
    
        i = Abs(m.aRatios(eGannci_CD) - m.aRatios(eGannci_R))
        i = GannacciEqualColor(i)
        m.aRatios(eGannci_EQ_CD_R) = i
        m.aRatioColor(eGannci_EQ_CD_R) = i
    Else
        m.aRatios(eGannci_CD) = -1
        m.aRatioColor(eGannci_CD) = -1
        
        'calendar days minus P1,P2,R
        m.aRatios(eGannci_CD_P1) = -1
        m.aRatioColor(eGannci_CD_P1) = -1
        m.aRatios(eGannci_CD_P2) = -1
        m.aRatioColor(eGannci_CD_P2) = -1
        m.aRatios(eGannci_CD_R) = -1
        m.aRatioColor(eGannci_CD_R) = -1
        
        'calendar days equals P1,P2,R
        m.aRatios(eGannci_EQ_CD_P1) = -1
        m.aRatioColor(eGannci_EQ_CD_P1) = -1
        m.aRatios(eGannci_EQ_CD_P2) = -1
        m.aRatioColor(eGannci_EQ_CD_P2) = -1
        m.aRatios(eGannci_EQ_CD_R) = -1
        m.aRatioColor(eGannci_EQ_CD_R) = -1
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.CalcGannacciSwing1"

End Sub

'iExtra is passed in as 1 when calculating Trade Bars or Calendar Days and include firat bar is turned on
Private Function GannacciFormula(ByVal X1#, ByVal X2#, _
    Optional ByVal iExtra& = 0, Optional ByVal bTradeBarsDiff As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim d#, i&, h&
    
    If X2 = kNullData Then
        'gannacci swing1 tool applies this formula to the individual prices (i.e. P1,P2)
        d = Sqr(X1) * 180 - 225
        d = Round(d) Mod 360
        If d < 0 Then d = d + 360
    Else
        d = Abs(X1 - X2)
        
        If bTradeBarsDiff And Not m.Chart Is Nothing Then
            If m.Chart.ShowEmptyBars Then
                h = m.Chart.aXBar.ArrayHandle
                If h <> 0 And Not m.Chart.Bars Is Nothing Then
                    If m.Chart.Bars.Prop(eBARS_PeriodicityStr) = "Daily" Then
                        For i = X1 To X2
                            If m.Chart.Bars(eBARS_DateTime, gdGetNum(h, i)) = kNullData Then d = d - 1
                        Next
                    End If
                End If
            End If
        End If
        
        d = d + iExtra
        d = Sqr(d) * 180 - 225
        d = Round(d) Mod 360
        If d < 0 Then d = d + 360
    End If
    
    GannacciFormula = d

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.GannacciFormula"

End Function

Private Function GannacciSwingColor(ByVal i&) As Long
On Error GoTo ErrSection:

    Select Case i
        Case 87, 177, 267, 357, 93, 183, 273
            GannacciSwingColor = Val(Prop("SignalColorLow"))
        Case 88, 178, 268, 358, 92, 182, 272
            GannacciSwingColor = Val(Prop("SignalColorMedium"))
        Case 89, 90, 91, 179, 180, 181, 269, 270, 271, 359, 360
            GannacciSwingColor = Val(Prop("SignalColorHigh"))
        Case Else
            GannacciSwingColor = -1
    End Select

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.GannacciSwingColor"

End Function

Public Property Get GannacciSwingPrice(ByVal iPriceNum&) As String
On Error GoTo ErrSection:

    Dim i&, d#, dY#, strFmt$
    Dim Bars As cGdBars
    
    If m.eType <> eANNOT_GannacciSwing1 And m.eType <> eANNOT_GannacciSwing2 Then Exit Property
    
    If m.Chart Is Nothing Then Exit Property
    
    Set Bars = m.Chart.Bars
    If Bars Is Nothing Then Exit Property
    
    If iPriceNum = 1 Then
        If m.X1 < m.X2 Then
            dY = m.Y1
        ElseIf m.aYs.Size > 0 Then
            dY = m.aYs(0)
        Else
            dY = m.Y2
        End If
    ElseIf iPriceNum = 2 Then
        If m.eType = eANNOT_GannacciSwing1 Or m.aYs.Size = 0 Then
            If m.X1 < m.X2 Then
                dY = m.Y2
            Else
                dY = m.Y1
            End If
        Else
            dY = m.Y2
        End If
    ElseIf iPriceNum = 3 And m.eType = eANNOT_GannacciSwing2 And m.aYs.Size > 0 Then
        If m.X1 < m.X2 Then
            dY = m.aYs(0)
        Else
            dY = m.Y1
        End If
    End If
    
    If dY > 0 Then
        strFmt = "#0.00###"             'format if rounding is not selected
        
        If Val(Prop("UseMutiplier")) = 1 Then
            d = Val(Prop("MultiplierVal"))
            If d <> 0 Then
                dY = dY * d
                i = Val(Prop("Decimals"))
                If Val(Prop("RoundDecimals")) = 1 And i >= 0 Then
                    strFmt = "0." & String(i, "0")
                End If
            End If
        End If
    
        GannacciSwingPrice = Format(dY, strFmt)
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.GannacciSwingPrice.Get"

End Property

Public Property Get GannacciDebugFlag() As Boolean

    If m.eType = eANNOT_GannacciSwing1 Or m.eType = eANNOT_GannacciSwing2 Then
        GannacciDebugFlag = m.bGannacciDebug
    End If

End Property

Public Property Get GannacciDebugText() As String

    If m.eType = eANNOT_GannacciSwing1 Then
        GannacciDebugText = GannacciSwing1DebugText
    ElseIf m.eType = eANNOT_GannacciSwing2 Then
        GannacciDebugText = GannacciSwing2DebugText
    End If

End Property

Private Property Get GannacciSwing1DebugText() As String

    Dim strDebug$, Str$
    
    If m.Chart Is Nothing Then Exit Property
    
    strDebug = "TB ( " & BarsOrCalendarCount(1, 1) & " ) = " & m.aRatios(eGannci_TB) & vbCrLf
    strDebug = strDebug & "CD ( " & BarsOrCalendarCount(1, 2) & " ) = " & m.aRatios(eGannci_CD) & vbCrLf
    strDebug = strDebug & "P1 ( " & GannacciSwingPrice(1) & " ) = " & m.aRatios(eGannci_P1) & vbCrLf
    strDebug = strDebug & "P2 ( " & GannacciSwingPrice(2) & " ) = " & m.aRatios(eGannci_P2) & vbCrLf
    strDebug = strDebug & "R = " & m.aRatios(eGannci_R) & vbCrLf & vbCrLf
    
    strDebug = strDebug & "TB-P1 = " & m.aRatios(eGannci_TB_P1) & vbCrLf
    strDebug = strDebug & "CD-P1 = " & m.aRatios(eGannci_CD_P1) & vbCrLf
    strDebug = strDebug & "TB-P2 = " & m.aRatios(eGannci_TB_P2) & vbCrLf
    strDebug = strDebug & "CD-P2 = " & m.aRatios(eGannci_CD_P2) & vbCrLf
    strDebug = strDebug & "TB-R = " & m.aRatios(eGannci_TB_R) & vbCrLf
    strDebug = strDebug & "CD-R = " & m.aRatios(eGannci_CD_R) & vbCrLf & vbCrLf
    
    Str = "T"
    If m.aRatios(eGannci_EQ_TB_P1) = -1 Then Str = "F"
    strDebug = strDebug & "TB=P1 = " & Str & vbCrLf
    
    Str = "T"
    If m.aRatios(eGannci_EQ_CD_P1) = -1 Then Str = "F"
    strDebug = strDebug & "CD=P1 = " & Str & vbCrLf
    
    Str = "T"
    If m.aRatios(eGannci_EQ_TB_R) = -1 Then Str = "F"
    strDebug = strDebug & "TB=R = " & Str & vbCrLf
    
    Str = "T"
    If m.aRatios(eGannci_EQ_CD_R) = -1 Then Str = "F"
    strDebug = strDebug & "CD=R = " & Str & vbCrLf
    
    Str = "T"
    If m.aRatios(eGannci_EQ_TB_P2) = -1 Then Str = "F"
    strDebug = strDebug & "TB=P2 = " & Str & vbCrLf
    
    Str = "T"
    If m.aRatios(eGannci_EQ_CD_P1) = -1 Then Str = "F"
    strDebug = strDebug & "CD=P2 = " & Str
    
    GannacciSwing1DebugText = strDebug

End Property

Private Property Get GannacciSwing2DebugText() As String

    Dim strDebug$, Str$
    
    If m.Chart Is Nothing Then Exit Property
    
    strDebug = "P1 = " & GannacciSwingPrice(1) & vbCrLf
    strDebug = strDebug & "P2 = " & GannacciSwingPrice(2) & vbCrLf
    strDebug = strDebug & "P3 = " & GannacciSwingPrice(3) & vbCrLf & vbCrLf
    
    strDebug = strDebug & "TB1 ( " & BarsOrCalendarCount(1, 1) & " ) = " & m.aRatios(eGannci_TB1) & vbCrLf
    strDebug = strDebug & "CD1 ( " & BarsOrCalendarCount(1, 2) & " ) = " & m.aRatios(eGannci_CD1) & vbCrLf
    strDebug = strDebug & "TB2 ( " & BarsOrCalendarCount(2, 1) & " ) = " & m.aRatios(eGannci_TB2) & vbCrLf
    strDebug = strDebug & "CD2 ( " & BarsOrCalendarCount(2, 2) & " ) = " & m.aRatios(eGannci_CD2) & vbCrLf
    
    strDebug = strDebug & "R1 = " & m.aRatios(eGannci_R1) & vbCrLf
    strDebug = strDebug & "R2 = " & m.aRatios(eGannci_R2) & vbCrLf & vbCrLf
    
    strDebug = strDebug & "TB1-R2 = " & m.aRatios(eGannci_TB1_R2) & vbCrLf
    strDebug = strDebug & "CD1-R2 = " & m.aRatios(eGannci_CD1_R2) & vbCrLf
    strDebug = strDebug & "TB2-R1 = " & m.aRatios(eGannci_TB2_R1) & vbCrLf
    strDebug = strDebug & "CD2-R1 = " & m.aRatios(eGannci_CD2_R1) & vbCrLf & vbCrLf
    
    Str = "T"
    If m.aRatioColor(eGannci_EQ_TB1_R2) = -1 Then Str = "F"
    strDebug = strDebug & "TB1=R2 = " & Str & vbCrLf
    
    Str = "T"
    If m.aRatioColor(eGannci_EQ_CD1_R2) = -1 Then Str = "F"
    strDebug = strDebug & "CD1=R2 = " & Str & vbCrLf
    
    Str = "T"
    If m.aRatioColor(eGannci_EQ_TB2_R1) = -1 Then Str = "F"
    strDebug = strDebug & "TB2=R1 = " & Str & vbCrLf
    
    Str = "T"
    If m.aRatioColor(eGannci_EQ_CD2_R1) = -1 Then Str = "F"
    strDebug = strDebug & "CD2=R1 = " & Str & vbCrLf
    
    GannacciSwing2DebugText = strDebug

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' nItemNum 1 = 1st swing line, nItemNum 2 = 2nd swing line
' nItemType 1 = trade bars, nItemType 2 = calendar days
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get BarsOrCalendarCount(ByVal nItemNum&, ByVal nItemType&) As Double
On Error GoTo ErrSection:

    Dim i&, j&, h&, iCount&
    Dim X1#, X2#, X3#, d1#, d2#, d3#
    
    BarsOrCalendarCount = -1
    
    i = Val(Prop("IncludeBarOne"))
    If i < 0 Or i > 1 Then i = 0
    
    Select Case m.eType
        Case eANNOT_GannacciSwing1
            If nItemNum = 1 Then
                If nItemType = 1 Then
                    iCount = Abs(m.X1 - m.X2) + i
                    
                    If Not m.Chart Is Nothing Then
                        h = m.Chart.aXBar.ArrayHandle
                        If h <> 0 And Not m.Chart.Bars Is Nothing Then
                            For j = m.X1 To m.X2
                                If m.Chart.Bars(eBARS_DateTime, gdGetNum(h, j)) = kNullData Then iCount = iCount - 1
                            Next
                        End If
                    End If
                    
                    BarsOrCalendarCount = iCount
                    
                Else
                    BarsOrCalendarCount = Abs(m.dDate1 - m.dDate2) + i
                End If
            End If

        Case eANNOT_GannacciSwing2
            If m.X1 < m.X2 Then
                X1 = m.X1
                X2 = m.X2
                
                d1 = m.dDate1
                d2 = m.dDate2
                
                If m.aDates.Size > 0 Then
                    If Not m.Chart Is Nothing Then X3 = GetX(m.Chart, m.aDates(0))
                    d3 = m.aDates(0)
                End If
            ElseIf m.aDates.Size > 0 Then
                If Not m.Chart Is Nothing Then X1 = GetX(m.Chart, m.aDates(0))
                X2 = m.X2
                X3 = m.X1
                
                d1 = m.aDates(0)
                d2 = m.dDate2
                d3 = m.dDate1
            Else
                X1 = m.X2
                X2 = m.X1
                d1 = m.dDate2
                d2 = m.dDate1
            End If
    
            If nItemNum = 1 Then
                If nItemType = 1 Then
                    iCount = Abs(X1 - X2) + i
                    
                    If Not m.Chart Is Nothing Then
                        h = m.Chart.aXBar.ArrayHandle
                        If h <> 0 And Not m.Chart.Bars Is Nothing Then
                            For j = X1 To X2
                                If m.Chart.Bars(eBARS_DateTime, gdGetNum(h, j)) = kNullData Then iCount = iCount - 1
                            Next
                        End If
                    End If
                    
                    BarsOrCalendarCount = iCount
                Else
                    BarsOrCalendarCount = Abs(d1 - d2) + i
                End If
            ElseIf nItemNum = 2 Then
                If nItemType = 1 Then
                    iCount = Abs(X2 - X3) + i
                    
                    If Not m.Chart Is Nothing Then
                        h = m.Chart.aXBar.ArrayHandle
                        If h <> 0 And Not m.Chart.Bars Is Nothing Then
                            For j = X2 To X3
                                If m.Chart.Bars(eBARS_DateTime, gdGetNum(h, j)) = kNullData Then iCount = iCount - 1
                            Next
                        End If
                    End If
                    
                    BarsOrCalendarCount = iCount
                Else
                    BarsOrCalendarCount = Abs(d2 - d3) + i
                End If
            End If

    End Select

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.BarsOrCalendarCount.Get"

End Property

Private Function LogModeDelta(ByVal dY1#, ByVal dY2#) As Double
On Error GoTo ErrSection:

    Dim d#
    
    If dY1 <= 0 And dY2 <= 0 Then Exit Function         '6654
    
    If dY1 <= 0 Then
        d = Log(dY2)
    ElseIf dY2 <= 0 Then
        d = -Log(dY1)
    Else
        d = Log(dY2) - Log(dY1)
    End If
    
    LogModeDelta = d

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.LogModeDelta"

End Function

'this sub sets draw information for the vertical lines of the Gannacci SR tool
'd = price move in points lines are drawn int(d) and int(d)+1 distance forward in time
'    from the right-most end point of the Gannacci Same Swing (dollar style line)
Private Sub SetGannacciSquareRange(Chart As cChart, ByVal d#, ByVal dX1#, dX2#)
On Error GoTo ErrSection:

    Dim i&, nPenStyle&, strDay$
    Dim dDate#, dX3#, dX4#, dY3#, dY4#
    Dim Pane As cPane

    'reset any previously set square range lines
    For i = 3 To gdGetSize(m.geAnnStruct.glhAnnType) - 1
        gdSetNum m.geAnnStruct.glhAnnType, i, -1
        gdSetStr m.geAnnStruct.gshText, i, ""
    Next

    If d <= 0 Or Chart Is Nothing Then Exit Sub
    
    If Prop("SquareRange") = 0 Then Exit Sub
    Set Pane = Chart.Tree("PRICE PANE")
    If Pane Is Nothing Then Exit Sub
    
    dY4 = Pane.geAdjustMin + Chart.Bars.MinMove
    If Val(Prop("Ext")) = 1 Then
        dY3 = Pane.Max
        m.geAnnStruct.lenPoints = 1
    Else
        dY3 = dY4 + (Pane.Max - dY4) * 0.15
        m.geAnnStruct.lenPoints = 0
    End If
    
    If Prop("LenBars") = 1 Then
        If dX1 < dX2 Then
            dX3 = GetGannacciX(Chart, strDay, GetX(Chart, m.dDate2), -1, Int(d))
            dX4 = GetGannacciX(Chart, strDay, GetX(Chart, m.dDate2), -1, Int(d) + 1)
        Else
            dX3 = GetGannacciX(Chart, strDay, GetX(Chart, m.dDate1), -1, Int(d))
            dX4 = GetGannacciX(Chart, strDay, GetX(Chart, m.dDate1), -1, Int(d) + 1)
        End If
        
        If dX4 >= 0 And dX4 < gdGetSize(Chart.geDateArray) Then
            dDate = gdGetNum(Chart.geDateArray, dX4)
        End If
        
        For i = 3 To 4
            gdSetNum m.geAnnStruct.glhAnnType, i, 39
            geSetPenStyle m.nStyle, i
            gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
            
            gdSetNum m.geAnnStruct.gdhTop, i, dY3
            gdSetNum m.geAnnStruct.gdhBottom, i, dY4
            
            If i = 3 Then
                gdSetNum m.geAnnStruct.gdhLeft, i, dX3
                gdSetNum m.geAnnStruct.gdhRight, i, dX3
                gdSetStr m.geAnnStruct.gshText, i, Space(10)
            Else
                gdSetNum m.geAnnStruct.gdhLeft, i, dX4
                gdSetNum m.geAnnStruct.gdhRight, i, dX4
                If dDate > 0 Then
                    gdSetStr m.geAnnStruct.gshText, i, DateFormat(dDate, MM_DD_YYYY)
                Else
                    gdSetStr m.geAnnStruct.gshText, i, Space(10)
                End If
            End If
        
            gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
            gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
            gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
            gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
            gdSetNum m.geAnnStruct.glhAlign, i, 10      'vertical text
        Next
    End If
    
    If Prop("CalendarDays") = 1 Then
        nPenStyle = m.nStyle
        If dX1 < dX2 Then
            dX3 = GetGannacciX(Chart, strDay, -1, m.dDate2, Int(d))
            If Len(strDay) > 0 Then nPenStyle = eANNOT_DashLg
            dX4 = GetGannacciX(Chart, strDay, -1, m.dDate2, Int(d) + 1)
            If Len(strDay) > 0 Then nPenStyle = eANNOT_DashLg
            dDate = m.dDate2 + Int(d) + 1
        Else
            dX3 = GetGannacciX(Chart, strDay, -1, m.dDate1, Int(d))
            If Len(strDay) > 0 Then nPenStyle = eANNOT_DashLg
            dX4 = GetGannacciX(Chart, strDay, -1, m.dDate1, Int(d) + 1)
            If Len(strDay) > 0 Then nPenStyle = eANNOT_DashLg
            dDate = m.dDate1 + Int(d) + 1
        End If
        
        For i = 5 To 6
            gdSetNum m.geAnnStruct.glhAnnType, i, 39
            geSetPenStyle nPenStyle, i
            gdSetNum m.geAnnStruct.glhPenColor, i, m.nColor
            
            gdSetNum m.geAnnStruct.gdhTop, i, dY3
            gdSetNum m.geAnnStruct.gdhBottom, i, dY4
            
            If i = 5 Then
                If dX3 >= 0 Then
                    gdSetNum m.geAnnStruct.gdhLeft, i, dX3
                    gdSetNum m.geAnnStruct.gdhRight, i, dX3
                    gdSetStr m.geAnnStruct.gshText, i, Space(10)
                Else
                    gdSetNum m.geAnnStruct.glhAnnType, i, -1
                End If
            Else
                If dX4 >= 0 Then
                    gdSetNum m.geAnnStruct.gdhLeft, i, dX4
                    gdSetNum m.geAnnStruct.gdhRight, i, dX4
                    If dDate > 0 Then
                        gdSetStr m.geAnnStruct.gshText, i, DateFormat(dDate, MM_DD_YYYY)
                    Else
                        gdSetStr m.geAnnStruct.gshText, i, Space(10)
                    End If
                Else
                    gdSetNum m.geAnnStruct.glhAnnType, i, -1
                End If
            End If
        
            gdSetStr m.geAnnStruct.gshFont, i, Prop("FontName")
            gdSetNum m.geAnnStruct.glhSize, i, Val(Prop("FontSize"))
            gdSetNum m.geAnnStruct.glhStyle, i, Val(Prop("FontStyle"))
            gdSetNum m.geAnnStruct.glhUnderline, i, Val(Prop("FontUnderline"))
            gdSetNum m.geAnnStruct.glhAlign, i, 10      'vertical text
        Next
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.SetGannacciSquareRange"

End Sub

'dFromX: used to calculate #of bars (pass -1 to be iguored)
'dFromDate: used to calculate #of calendars (pass -1 to be ignored)
'if neither are -1 then function only returns result for #of bars
Private Function GetGannacciX(Chart As cChart, strWeekDay$, ByVal dFromX#, ByVal dFromDate#, ByVal nCount&) As Double
On Error GoTo ErrSection:

    Dim i&, hArray&, d#, strDay$
    
    Dim Bars As cGdBars
    
    GetGannacciX = -1#
    strWeekDay = ""
    
    If Chart Is Nothing Then Exit Function
    
    Set Bars = Chart.Bars
    If Bars Is Nothing Then Exit Function
    
    If dFromX >= 0 Then
        
        hArray = m.Chart.aXBar.ArrayHandle
        If hArray <> 0 Then
            d = dFromX
            'add nCount bars to start bar, making sure to skip over blank bars
            For i = 1 To nCount
                d = d + 1
                If Bars(eBARS_DateTime, gdGetNum(hArray, d)) = kNullData Then d = d + 1
            Next
            
            GetGannacciX = d - Chart.ScreenStartX
        End If
    
    ElseIf dFromDate >= 0 And Bars.Prop(eBARS_PeriodicityStr) = "Daily" Then
        
        d = dFromDate + nCount
        strDay = UCase(WeekdayName(d))
        
        If strDay = "SUN" Then
            d = d + 1   'bump to Monday
            strWeekDay = strDay
        ElseIf strDay = "SAT" Then
            d = d - 1   'bump to Friday
            strWeekDay = strDay
        End If
        GetGannacciX = GetX(Chart, d) - Chart.ScreenStartX
    
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.GetGannacciX"

End Function

Private Function GannacciEqualColor(ByVal i&) As Long
On Error GoTo ErrSection:

    Select Case i
        Case 3
            GannacciEqualColor = Val(Prop("SignalColorLow"))
        Case 2
            GannacciEqualColor = Val(Prop("SignalColorMedium"))
        Case 0, 1
            GannacciEqualColor = Val(Prop("SignalColorHigh"))
        Case Else
            GannacciEqualColor = -1
    End Select

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.GannacciEqualColor"

End Function

Public Property Get FibLineBold() As Boolean
On Error GoTo ErrSection:

    Dim i&
    
    If m.eType <> eANNOT_Fibonacci And m.eType <> eANNOT_Fibonacci2 And m.eType <> eANNOT_Fibonacci3 _
       And m.eType <> eANNOT_Fibonacci4 And m.eType <> eANNOT_FibExpansion And m.eType <> eANNOT_DanCodeFib _
       And m.eType <> eANNOT_AdvRiskReward Then
        Exit Property
    End If
    
    If m.nHitItemIdx = 2 Then
        If m.bY1IsZeroRatio Then
            i = Val(Prop("BoldRatioZero"))
        Else
            i = Val(Prop("BoldRatioOne"))
        End If
        If i = 1 Then FibLineBold = True
    ElseIf m.nHitItemIdx = 3 Then
        If m.bY1IsZeroRatio Then
            i = Val(Prop("BoldRatioOne"))
        Else
            i = Val(Prop("BoldRatioZero"))
        End If
        If i = 1 Then FibLineBold = True
    ElseIf m.nHitItemIdx > 5 Then
        i = m.nHitItemIdx - 6
        
        If i >= 0 And i < m.aRatioBold.Size Then
            If m.aRatioBold(i) = 1 Then FibLineBold = True
        End If
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.FibLineBold.Get"

End Property

Public Property Let FibLineBold(ByVal bBold As Boolean)
On Error GoTo ErrSection:

    Dim i&

    If m.eType <> eANNOT_Fibonacci And m.eType <> eANNOT_Fibonacci2 And m.eType <> eANNOT_Fibonacci3 _
       And m.eType <> eANNOT_Fibonacci4 And m.eType <> eANNOT_FibExpansion And m.eType <> eANNOT_DanCodeFib _
       And m.eType <> eANNOT_AdvRiskReward Then
        Exit Property
    End If
    
    If m.nHitItemIdx = 2 Then
        If m.bY1IsZeroRatio Then
            Prop("BoldRatioZero") = Str(Abs(bBold))
        Else
            Prop("BoldRatioOne") = Str(Abs(bBold))
        End If
    ElseIf m.nHitItemIdx = 3 Then
        If m.bY1IsZeroRatio Then
            Prop("BoldRatioOne") = Str(Abs(bBold))
        Else
            Prop("BoldRatioZero") = Str(Abs(bBold))
        End If
    ElseIf m.nHitItemIdx > 5 Then
        i = m.nHitItemIdx - 6
        If i >= 0 And i < m.aRatioBold.Size Then
            If bBold Then
                m.aRatioBold(i) = 1
            Else
                m.aRatioBold(i) = 0
            End If
            Prop("FibBold") = m.aRatioBold.JoinFields(",")
        End If
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.FibLineBold.Let"

End Property

Public Function FibDynamicValidate() As Boolean
On Error GoTo ErrSection:

    Dim dNewY#, nDontCare&, strMsg$
    Dim bErr As Boolean
    
    If m.Chart Is Nothing Then Exit Function
    If m.Chart.Bars Is Nothing Then Exit Function
    
    strMsg = "Enabling dynamic stay on last swing will reverse the direction of your drawing tool."

    If m.dDate1 < m.dDate2 Then
        If m.Chart.Bars.IsIntraday Then
            nDontCare = m.Chart.Bars.FindDateTime(m.dDate2)
        Else
            nDontCare = m.Chart.Bars.FindDateTime(Int(m.dDate2))
        End If
        
        If m.Y2 > m.Y1 Then
            If HighestHighInBarRange(m.Chart.Bars, dNewY, nDontCare, nDontCare, m.Chart.Bars.Size - 1) Then
                If dNewY < m.Y1 And dNewY < m.Y2 Then bErr = True
            End If
        Else
            If LowestLowInBarRange(m.Chart.Bars, dNewY, nDontCare, nDontCare, m.Chart.Bars.Size - 1) Then
                If dNewY > m.Y1 And dNewY > m.Y2 Then bErr = True
            End If
        End If
    Else
        If m.Chart.Bars.IsIntraday Then
            nDontCare = m.Chart.Bars.FindDateTime(m.dDate1)
        Else
            nDontCare = m.Chart.Bars.FindDateTime(Int(m.dDate1))
        End If
        
        If m.Y1 > m.Y2 Then
            If HighestHighInBarRange(m.Chart.Bars, dNewY, nDontCare, nDontCare, m.Chart.Bars.Size - 1) Then
                If dNewY < m.Y1 And dNewY < m.Y2 Then bErr = True
            End If
        Else
            If LowestLowInBarRange(m.Chart.Bars, dNewY, nDontCare, nDontCare, m.Chart.Bars.Size - 1) Then
                If dNewY > m.Y1 And dNewY > m.Y2 Then bErr = True
            End If
        End If
    End If
    
    If bErr Then
        If InfBox(strMsg, "I", "Ok|Cancel", "Dynamic stay on last swing") = "O" Then bErr = False
    End If
    
    FibDynamicValidate = Not bErr

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAnnotation.FibDynamicValidate"

End Function

Public Sub BalloonRatiosToGrid(fg As VSFlexGrid)
On Error GoTo ErrSection:
    
    Dim i&

    If m.eType <> eANNOT_BalloonStrangle Then Exit Sub
    
    i = m.aRatios.Size
    If i <> m.aRatioColor.Size Then Exit Sub     'precautionary
    If i <> m.aRatioShow.Size Then Exit Sub
    
    fg.Rows = m.aRatios.Size + 1
    For i = 0 To m.aRatios.Size - 1
        fg.TextMatrix(i + 1, 1) = m.aRatios(i)
        fg.Cell(flexcpBackColor, i + 1, 2) = m.aRatioColor(i)
    
        If m.aRatioShow(i) = 1 Then
            fg.Cell(flexcpChecked, i + 1, 0) = flexChecked   'show
        Else
            fg.Cell(flexcpChecked, i + 1, 0) = flexUnchecked
        End If
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.BalloonRatiosToGrid"

End Sub

Public Sub GridToBalloonRatios(fg As VSFlexGrid)
On Error GoTo ErrSection:

    Dim i&, Size&
    
    If m.eType <> eANNOT_BalloonStrangle Then Exit Sub
    
    Size = m.aRatios.Size
    If Size <> m.aRatioColor.Size Then Exit Sub     'precautionary
    If Size <> m.aRatioShow.Size Then Exit Sub
    If Size <> m.aRatios.Size Then Exit Sub
    
    For i = 1 To fg.Rows - 1
        If i - 1 >= Size Then Exit For
        
        m.aRatios(i - 1) = fg.TextMatrix(i, 1)
        m.aRatioColor(i - 1) = fg.Cell(flexcpBackColor, i, 2)
        
        If fg.Cell(flexcpChecked, i, 0) = flexChecked Then
            m.aRatioShow(i - 1) = 1
        Else
            m.aRatioShow(i - 1) = 0
        End If
    Next

    Prop("BalloonRatios") = m.aRatios.JoinFields(",")
    Prop("BalloonColors") = m.aRatioColor.JoinFields(",")
    Prop("ShowRatios") = m.aRatioShow.JoinFields(",")

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAnnotation.GridToBalloonRatios"

End Sub

Public Property Get BalloonExpiration() As Double
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    
    If m.aDates.Size > 0 Then
        BalloonExpiration = m.aDates(0)
    Else
        BalloonExpiration = m.dDate2
    End If

End Property

Public Property Let BalloonExpiration(ByVal expireDate As Double)
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    m.aDates(0) = expireDate
End Property

Public Property Get BalloonStockPrice() As Double
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    BalloonStockPrice = m.Y1
End Property

Public Property Let BalloonStockPrice(ByVal dY As Double)
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    m.Y1 = dY
End Property

Public Property Get BalloonPutStrike() As Double
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    BalloonPutStrike = m.Y2
End Property

Public Property Let BalloonPutStrike(ByVal dY As Double)
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    m.Y2 = dY
    geSetPoints
End Property

Public Property Get BalloonCallStrike() As Double
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    BalloonCallStrike = m.aYs(0)
End Property

Public Property Let BalloonCallStrike(ByVal dY As Double)
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    m.aYs(0) = dY
    geSetPoints
End Property

Public Property Get BalloonPutCost() As Double
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    BalloonPutCost = m.aYs(1)
End Property

Public Property Let BalloonPutCost(ByVal dY As Double)
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    m.aYs(1) = dY
    geSetPoints
End Property

Public Property Get BalloonCallCost() As Double
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    BalloonCallCost = m.aYs(2)
End Property

Public Property Let BalloonCallCost(ByVal dY As Double)
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    m.aYs(2) = dY
    geSetPoints
End Property

Public Property Get BalloonTradeCost() As Double
    If m.eType <> eANNOT_BalloonStrangle Then Exit Property
    BalloonTradeCost = BalloonPutCost + BalloonCallCost
End Property

Public Property Get IsBalloonOptionsInfo() As Boolean
    If m.tbOptionsInfo Is Nothing Then Exit Property
    If m.tbOptionsInfo.NumFields <> 5 Then Exit Property
    
    IsBalloonOptionsInfo = True
End Property

Public Property Let BalloonOptionsInfo(aInput As cGdArray)
On Error GoTo ErrSection:

    Dim i&, j&, k&, d#
    Dim hInput&, hStrike&, hCall&, hPut&, hCallAsk&, hPutAsk&
    Dim strText$, strChar$, strStrike$

    If aInput Is Nothing Then Exit Property
    
    If aInput.Size = 0 Then
        InfBox "Options data request returned array with size zero", "I", , "Balloon Strangle"
        Exit Property
    End If
    
    'put data into same format as used for options order bar (see LoadOptionsInfo)
    Set m.tbOptionsInfo = New cGdTable
    m.tbOptionsInfo.CreateField eGDARRAY_Doubles, , , -1
    m.tbOptionsInfo.CreateField eGDARRAY_Strings
    m.tbOptionsInfo.CreateField eGDARRAY_Strings
    m.tbOptionsInfo.CreateField eGDARRAY_Doubles, , , -1
    m.tbOptionsInfo.CreateField eGDARRAY_Doubles, , , -1

    m.tbOptionsInfo.NumRecords = 0
    
    hInput = aInput.ArrayHandle
    hStrike = m.tbOptionsInfo.FieldArrayHandle(0)
    hCall = m.tbOptionsInfo.FieldArrayHandle(1)
    hPut = m.tbOptionsInfo.FieldArrayHandle(2)
    hCallAsk = m.tbOptionsInfo.FieldArrayHandle(3)
    hPutAsk = m.tbOptionsInfo.FieldArrayHandle(4)
        
    strText = Parse(gdGetStr(hInput, 0), vbTab, 1)
    'get expiration date
    strChar = Parse(strText, " ", 2)
    m.dDate1 = DateOf(strChar)
    m.dDate2 = m.dDate1
    
    'get the char of first array item to use for controlling for loop
    strStrike = Parse(strText, " ", 3)
    If Len(strStrike) > 0 Then strChar = Left(strStrike, 1)
    If strChar <> "C" And strChar <> "P" Then Exit Property
    
    For i = 0 To aInput.Size - 1
        strText = gdGetStr(hInput, i)
        strStrike = Parse(Parse(strText, vbTab, 1), " ", 3)
        If Left(strStrike, 1) <> strChar Then
            strChar = Left(strStrike, 1)
            Exit For
        End If
        
        If Len(strStrike) > 0 Then
            m.tbOptionsInfo.AddRecord ""
            k = m.tbOptionsInfo.NumRecords - 1
            strStrike = Mid(strStrike, 2, Len(strStrike) - 1)
            gdSetNum hStrike, k, strStrike
            If strChar = "C" Then
                gdSetStr hCall, k, Parse(strText, vbTab, 1)
                gdSetNum hCallAsk, k, Val(Parse(strText, vbTab, 3))
            Else
                gdSetStr hPut, k, Parse(strText, vbTab, 1)
                gdSetNum hPutAsk, k, Val(Parse(strText, vbTab, 3))
            End If
        End If
    Next
    
    If strChar <> "C" And strChar <> "P" Then Exit Property
    
    j = i
    k = 0
    For i = j To aInput.Size - 1
        strText = gdGetStr(hInput, i)
        strStrike = Parse(Parse(strText, vbTab, 1), " ", 3)
        If Left(strStrike, 1) <> strChar Then
            Exit For
        End If
        
        If Len(strStrike) > 0 Then
            strStrike = Mid(strStrike, 2, Len(strStrike) - 1)
            If k < m.tbOptionsInfo.NumRecords Then
                If strStrike = m.tbOptionsInfo(0, k) Then
                    If strChar = "C" Then
                        gdSetStr hCall, k, Parse(strText, vbTab, 1)
                        gdSetNum hCallAsk, k, Val(Parse(strText, vbTab, 3))
                    Else
                        gdSetStr hPut, k, Parse(strText, vbTab, 1)
                        gdSetNum hPutAsk, k, Val(Parse(strText, vbTab, 3))
                    End If
                    k = k + 1
                End If
            End If
        End If
    Next
        
    If m.tbOptionsInfo.NumRecords > 0 Then
        Set m.aOptionsIdx = Nothing
        Set m.aOptionsIdx = m.tbOptionsInfo.CreateSortedIndex(0, eGdSort_Default)
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cAnnotation.BalloonOptionsInfo"

End Property

Public Property Get IsEndUserEWI() As Boolean
    If m.eType = eANNOT_ElliotLabel Then
        If Val(Prop("EndUserEWI")) = 1 Then IsEndUserEWI = True
    End If
End Property

Public Property Get AllowEWI() As Boolean
    AllowEWI = FileExist(g.strAppPath & "\ewave.flg")
End Property

Public Property Get AllowEWIGMP() As Boolean
    AllowEWIGMP = FileExist(g.strAppPath & "\ewave.flg") And FileExist(g.strAppPath & "\gmp.flg")
End Property

Public Property Get AllowGMP() As Boolean
    AllowGMP = FileExist(g.strAppPath & "\gmp.flg")
End Property
