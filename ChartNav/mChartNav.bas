Attribute VB_Name = "mChartNav"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mChartNav.bas
'' Description: Global routines and variables for charting
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 04/16/2014   DAJ         Fix for submitting TradeSense order group via favorites ( Pete Laverde )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kDefaultConstStyle = eINDIC_Default

Public Const kMaxAxes = 6
Public Const kGameModeSysID = 9999999
'min width/height set in chart.frm (used for arranging charts as well)
Public Const kMinChartWidth = 1000
Public Const kMinChartHeight = 500
'file containing saved alert messages used by alerts collection & alert messages form
Public Const kstrHistoryFile = "Custom\AlertHistory.TXT"
'message used by footprint forms
Public Const kFootPrintNoVol = "This symbol does not have buy/sell volume."

'special key for chart's TREE object for fib cluster indicators
Public Const kClusterTimeKeyPane = "ClusterTimePane"
Public Const kClusterTimeKeyInd = "ClusterTimeInd"
Public Const kClusterPriceKey = "ClusterPriceInd"
Public Const kClusterZoneRect = "eANNOT_FibClusters"

'icons for quote board
Public Const kActiveAlertIcon = "ID_Alerts"
Public Const kInactiveAlertIcon = "kGrayBell"

'fib clusters
Public Const kClusterIniSection = "FibClusters"

'seasonal chart
Public Const kSeasonalAvgIndKey = "SEASONAL AVG TREND"
Public Const kSeasonalBullIndKey = "SEASONAL BULL TREND"
Public Const kSeasonalBearIndKey = "SEASONAL BEAR TREND"
Public Const kSeasonalCurrIndKey = "SEASONAL CURRENT CYCLE"
Public Const kSeasonalUnavail = "This feature not available for seasonal charts."

Private Const kPercentCompGridHeight = 1515

'flag file for showing URL
Public Const kShowUrlFlagFile = "ShowURL.flg"

Public Enum eImgSrvState
    eImgSrv_Free = 0
    eImgSrv_Searching = 1
    eImgSrv_Busy = 2
End Enum

Public Enum eChartMode
    eMode_Zoom = 0
    eMode_Move = 1
    eMode_Erase = 2
    eMode_ChartOrder = 3
End Enum

Public Enum eDragModeScaleY
    eDragModeY_Both = 0
    eDragModeY_Each
End Enum

Public Enum eExitFavoritesMode
    eAEFMode_Global = 0          'AEF = auto exits favorites
    eAEFMode_Symbol
End Enum

Public Type ChartGlobalStruct
    nChartForeColor As Long
    nChartBackColor As Long
    nChartGradientColor As Long
    nBorderForeColor As Long
    nBorderBackColor As Long
    nUseGradient As Long
    nFontSize As Long
    nFontStyle As Long          'prior to light, dark theme font style was hard coded as bold for indLabels and normal for scale labels
    bHideScrollbars As Boolean
    bFloatingTips As Boolean
    bChartTips As Boolean
    bSplitsRolls As Boolean
    nSplitRollColor As Long
    eDefaultBarsStyle As eIndicatorStyle
    eDefaultIndStyle As eIndicatorStyle
    eDefaultHorzStyle As eIndicatorStyle
    eProfitLineStyle As eAnnotPen
    eDefaultLabelMode As eIndicatorLabelMode
    eDefaultAnnotStyle As eAnnotPen
    eChartMode As eChartMode
    ePrevChartMode As eChartMode
    eScaleMode As ePANE_ScaleMode
    eDragModeY As eDragModeScaleY
    
    bLogModeDraw As Boolean
    bAutoChartData As Boolean
    bChartDataSingleBar As Boolean
    nShortColor As Long
    nLongColor As Long
    nWinColor As Long
    nLossColor As Long
    strFontName As String
    nSquareBars As Long
    dSquareTicks As Double
    aSquareSymList As New cGdArray
    eBadTickShowMode As eHighlightMode
    eBadTickMarker As eStockImage
    eBadTickMarkerDir As eImageDir
    nBadTickMarkerFill As Long
    nBadTickColor As Long
    nHideAnnotations As Long
    nMagnetValue As Long
    nGameInProg As Long                     '0=no game, 1=game on 1 chart, 2=game on multiple charts
    
    strCPCRoot As String                    'CPCRoot=ChartPageCollectionRoot (this is folder one level up for the page collection \charts folder)
    
    eAEFMode As eExitFavoritesMode          'mode for auto exit favorites
    aExitFavorites As cGdArray              'array of global auto exits favorite (holds cExitStrategy objects)
    'aTSOGrpFavorites As cGdArray            'array of favorite TSO groups
    treeExitFavorites As cGdTree            'tree of exit favorites on per symbol basis
    astrTsogFavorites As cGdArray           ' Array of filenames for the TradeSense order group favorites
    
    frmLastChartMouseMove As Form
    frmPfpSelPattern As Form                'form that user clicked the "Select Pattern" (PFP button) in
    
    'for handling detached charts
    nDetached As Long                       'count of detached charts in a page
    frmActiveNonDetached As frmChart        'most recently active MDIChild chart form
    bMyPageFeature As Boolean               'true=show pages feature without Save,Manage or Create options
    
    'Flags for Elliot Wave special request features:
    '   fixed size charts based on pixels
    '   always save blank bars even on charts move
    '   snap non-max & non-cascade charts to dots on background
    bChartModeAutoSize As Boolean           'true=auto resize chart as ratio, false=maintain fixed twips size
    bExtForecastBars As Boolean             '6780
    bSnapToDots As Boolean                  '6782
End Type
'Public g.ChartGlobals As ChartGlobalStruct

Public Type ChartCoordinates
    MouseX As Single
    MouseY As Single
    nPaneID As Long         '-1 if point is not in chart area
    dY As Double
    nScreenX As Long
    nX As Long
    nBar As Long
    dDate As Double
    bOffChart As Boolean
    dMinY As Double
    dMaxY As Double
    iShift As Integer
    nButton As Integer      'stores vbRightButton or vbLeftButton on mouse clicks etc.
    strRoundedY As String
    dTickTime As Double
    nScalePaneId As Long    '-1=point in x-scale area, >0=paneId when point in y-scale area
End Type

Public Type AnnotMove
    nType As eAnnotType
    nPegIndex As Long
    nUserIdx As Long
    nMovePt As Long
    nOtherPt As Long
    nClickTime As Long
End Type

Public Function AddStudyToChart(Chart As cChart, _
    ByVal strStudy$, ByVal bDisplay As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim idxPane&, idx&
    Dim Indicator As cIndicator
    Dim Pane As cPane

    idxPane = -1
    With Chart.Tree
        ' new pane
        Set Pane = New cPane
        Pane.Name = Trim(strStudy)
        Pane.Display = bDisplay
        
        Select Case UCase(Trim(strStudy))
            
            Case "PRICE":
                idxPane = .Add(Pane)
                If Not .Exists("PRICE PANE") Then
                    .Key(idxPane) = "PRICE PANE"
                End If
                Pane.Size = 2
                Pane.DisplayFormat = ePANE_PriceFormat
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Price"
                    .Display = True
                    .DataType = eINDIC_BarData
                    .DisplayType = eINDIC_OHLC
                    '.DisplayType = eINDIC_Candlestick
                    .Color = vbBlack 'vbRed
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                If Not .Exists("PRICE") Then
                    .Key(idx) = "PRICE"
                End If
                
#If 0 Then
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    .InputsFrame = 1
                    .Display = True
                    .Color = RGB(0, 128, 0)
                    .Parm(2) = "18"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
#End If
                
#If 0 Then
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Bands"
                    '.nNumSubsets = 2
                    .InputsFrame = 1
                    .Display = False
                    .Color = vbBlack
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 3
#End If

#If 0 Then
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    .Name = "Mvg.Avg.#2"
                    .InputsFrame = 1
                    .Display = True
                    .Color = vbRed
                    .Parm(2) = "40"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
#End If
                
#If 0 Then
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Mvg.Avg.#3"
                    .InputsFrame = 1
                    .Display = False
                    .Color = vbBlue
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Swing Points"
                    .Display = False
                    .Color = vbBlack
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "PredictedRange"
                    .Display = False
                    .Color = vbBlack
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Aggr.Trend"
                    .Display = False
                    .Color = vbBlack
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
#End If
                
                
            Case "WILL-SPREAD":
                idxPane = .Add(Pane)
                Pane.DisplayFormat = ePANE_PriceFormat
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "GC-067"
                    .Display = True
                    .DataType = eINDIC_BarData
                    .DisplayType = eINDIC_OHLC
                    '.DisplayType = eINDIC_Candlestick
                    .Color = 8421376
                    .Overlayed = True
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "WillSpread"
                    .Display = True
                    .Color = vbRed
                    .Parm(1) = "Market1"
                    .Parm(2) = "default"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "0-line"
                    .Display = True
                    .DataType = eINDIC_Constant
                    .Color = vbRed
                    .Style = kDefaultConstStyle
                    .Parm(1) = "0"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                
            Case "RELATIVE STRENGTH RATIO":
                idxPane = .Add(Pane)
                Pane.Size = 2
                Pane.DisplayFormat = ePANE_PriceFormat
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = ""
                    .Name = "$DJIA"
                    .Display = True
                    .DataType = eINDIC_BarData
                    .DisplayType = eINDIC_OHLC
                    .Color = &HC0C000
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    .Display = False
                    .Color = RGB(0, 0, 192)
                    .Parm(2) = "18"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "RelativeStrengthRatio"
                    .Display = True
                    .Color = RGB(255, 0, 0)
                    .Style = eINDIC_Medium
                    .Overlayed = True
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                
            Case "MARKET SPREAD":
                idxPane = .Add(Pane)
                Pane.Size = 2
                Pane.DisplayFormat = ePANE_PriceFormat
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = ""
                    .Name = "$DJIA"
                    .Display = True
                    .DataType = eINDIC_BarData
                    .DisplayType = eINDIC_OHLC
                    .Color = &HC0C000
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    .Display = False
                    .Color = RGB(0, 0, 192)
                    .Parm(2) = "18"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "Spread"
                    .Display = True
                    .Color = vbRed
                    .Style = eINDIC_Medium
                    .Overlayed = True
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "0-line"
                    .Display = True
                    .DataType = eINDIC_Constant
                    .Color = vbRed
                    .Style = kDefaultConstStyle
                    .Parm(1) = "0"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                
            Case "RED/GREEN LIGHT":
                idxPane = .Add(Pane)
                With Pane
                    .Scaling = ePANE_ScaleModeManual '-1
                End With
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "RedLightGreenLight"
                    .Name = "Red/Green Light"
                    .Display = True
                    .Color = vbRed
                    .DisplayType = eINDIC_Histogram
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
            
            Case "SEASONALS":
                idxPane = .Add(Pane)
                With Pane
                    .Scaling = ePANE_ScaleModeManual
                End With
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "SeasonalPercent"
                    .Display = True
                    .Color = RGB(0, 192, 192)
                    .DisplayType = eINDIC_Histogram
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1

                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "SeasonalTrend"
                    .Display = True
                    .Overlayed = True
                    .Color = RGB(0, 0, 192)
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                
            Case "ADX":
                idxPane = .Add(Pane)
                With Pane
                    .Scaling = ePANE_ScaleModeManual
                End With
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "ADX"
                    .Display = True
                    .Color = vbRed
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "ADXR"
                    '.Name = "ADXR"
                    .Display = False
                    .Color = vbBlack
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
#If 0 Then
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "PlusDMI"
                    .Display = False
                    .Color = vbBlue
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "MinusDMI"
                    .Display = False
                    .Color = vbBlue
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2

                Set Indicator = New cIndicator
                With Indicator
                    .Name = "High Value"
                    .Display = False
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = vbBlue
                    .Parm(1) = "60"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
#End If
        
            
            Case "STOCHASTIC":
                idxPane = .Add(Pane)
                With Pane
                    .Scaling = ePANE_ScaleModeManual
                End With
                
#If 0 Then
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "PercentR"
                    .Display = False
                    .Color = vbMagenta
                    .Parm(2) = "14"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    .Name = "StochK"
                    .Display = True
                    .Color = vbBlack
                    .Parm(2) = "3"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    '.CodedName = "StochasticD"
                    .CodedName = "MovingAvg"
                    .Name = "StochD"
                    .Display = True
                    .Color = vbBlue
                    .Parm(2) = "3"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 3
#Else
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "StochK"
                    '.Name = "StochK"
                    .Display = True
                    .Color = RGB(0, 0, 192)
                    .Parm(2) = "14"
                    .Parm(3) = "3"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    '.CodedName = "StochD"
                    .CodedName = "MovingAvg"
                    .Name = "StochD"
                    .Display = True
                    .Color = RGB(192, 0, 192)
                    .Parm(2) = "3"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
#End If
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Overbought"
                    .Display = True
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "80"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Oversold"
                    .Display = True
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "20"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
        
            Case "CASHFLOW":
                idxPane = .Add(Pane)
                With Pane
                    '.Scaling = ePANE_ScaleModeManual
                End With
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "CashFlow"
                    .Name = "CashFlow Accum"
                    .Display = True
                    .Color = vbRed
                    '.Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    '.name = "ADXR"
                    .Display = False
                    .Color = RGB(0, 0, 128)
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
            
            Case "POWER BALANCE":
                idxPane = .Add(Pane)
                With Pane
                    '.Scaling = ePANE_ScaleModeManual
                End With
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "PowerBalanceAccumulation"
                    .Name = "Power Balance Accum"
                    .Display = True
                    .Color = RGB(128, 0, 128)
                    .DisplayType = eINDIC_Histogram
                    .Parm(2) = "17"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    '.name = "ADXR"
                    .Display = False
                    .Color = RGB(0, 0, 128)
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
            
            Case "TIME VOLUME ACCUM":
                idxPane = .Add(Pane)
                With Pane
                    '.Scaling = ePANE_ScaleModeManual
                End With
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "TimeVolumeAccumulation"
                    .Name = "Time Volume Accum"
                    .Display = True
                    .Color = vbBlue
                    '.Parm(2) = "17"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    '.name = "ADXR"
                    .Display = False
                    .Color = vbMagenta
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "0-line"
                    .Display = True
                    .DataType = eINDIC_Constant
                    .Color = vbBlue
                    .Style = kDefaultConstStyle
                    .Parm(1) = "0"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
            
            Case "RSI":
                idxPane = .Add(Pane)
                With Pane
                    .Scaling = ePANE_ScaleModeManual
                End With
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "RSI"
                    .Display = True
                    .Color = vbBlack
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    .InputsFrame = 1
                    .Display = False
                    .Color = vbMagenta
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Overbought"
                    .Display = False
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "80"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Oversold"
                    .Display = False
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "20"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
        
            
            Case "PRO-GO":
                idxPane = .Add(Pane)
                With Pane
                End With

                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "ProGoProfessional"
                    .Display = True
                    .Color = RGB(0, 128, 0)
                    .Parm(2) = "14"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "ProGoPublic"
                    .Display = True
                    .Color = vbRed
                    .Parm(2) = "14"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "0-line"
                    .Display = False
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "0"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
        
            
            Case "WILL-VAL":
                idxPane = .Add(Pane)
                Pane.Name = "WILL-VAL"
                Pane.Display = bDisplay
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Will-Val"
                    .Display = True
                    .Color = vbBlue
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Overbought"
                    .Display = False
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "80"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Oversold"
                    .Display = False
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "20"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
        
            
            Case "MACD":
                idxPane = .Add(Pane)
                Pane.Name = "MACD"
                Pane.Display = bDisplay
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MACD"
                    .Display = True
                    .Color = RGB(0, 0, 192)
                    .Parm(2) = "12"
                    .Parm(3) = "26"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvgX"
                    .InputsFrame = 1
                    .Display = True
                    .Color = RGB(192, 0, 192)
                    .Parm(2) = "9"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "Spread"
                    .Name = "Diff"
                    .InputsFrame = 1
                    .Display = True
                    .DisplayType = eINDIC_Histogram
                    .Color = &HC0C000
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 3
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "0-line"
                    .Display = True
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = &HC0C000
                    .Parm(1) = "0"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
#If 0 Then
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Diff"
                    .Display = False
                    .Color = vbBlue
                    .DisplayType = eINDIC_Histogram
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 3
#End If
            
            
            Case "VOLUME":
                idxPane = .Add(Pane)
                Pane.Name = "VOLUME"
                Pane.Display = bDisplay
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "Vol"
                    .Name = "Volume"
                    .Display = True
                    .Color = RGB(0, 0, 128)
                    .DisplayType = eINDIC_Histogram
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    .Display = False
                    .Color = vbBlack
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                
            Case "VOLUME AND O.I.":
                idxPane = .Add(Pane)
                Pane.Name = "VOLUME and O.I."
                Pane.Display = bDisplay
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "Vol"
                    .Name = "Volume"
                    .Display = True
                    .Color = RGB(0, 0, 128)
                    .DisplayType = eINDIC_Histogram
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    .Display = False
                    .Color = vbBlack
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "OI"
                    .Name = "Open Interest"
                    .Display = True
                    .Color = RGB(128, 0, 128)
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "ContractVol"
                    .Name = "Contract Volume"
                    .Display = False
                    .Color = RGB(0, 0, 192)
                    .DisplayType = eINDIC_Histogram
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "MovingAvg"
                    .Display = False
                    .Color = vbBlack
                    .Parm(2) = "7"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 2
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "ContractOI"
                    .Name = "Contract OI"
                    .Display = False
                    .Color = RGB(192, 0, 192)
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
            
            Case "COT/SENTIMENT":
                idxPane = .Add(Pane)
                With Pane
                    .Scaling = ePANE_ScaleModeManual
                End With
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "COTCommercialsIndex"
                    .Display = True
                    .Color = vbRed
                    .Style = eINDIC_Medium
                    '.Parm(1) = "3"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "GenesisSentimentIndex"
                    .Display = True
                    .Color = vbBlue
                    '.Parm(1) = "3"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "TNSentiment"
                    .Display = True
                    .Color = RGB(0, 128, 0)
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Overbought"
                    .Display = True
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "75"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "Oversold"
                    .Display = True
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "25"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
            Case "COT NET POSITIONS":
                idxPane = .Add(Pane)
                With Pane
                    '.Scaling = ePANE_ScaleModeManual
                End With
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "COTCommercials"
                    .Display = True
                    .Color = vbRed
                    .Style = eINDIC_Medium
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "COTLargeSpec"
                    .Display = True
                    .Color = vbBlack
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .CodedName = "COTSmallSpec"
                    .Display = True
                    .Color = RGB(128, 128, 128)
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
                Set Indicator = New cIndicator
                With Indicator
                    .Name = "0-line"
                    .Display = True
                    .DataType = eINDIC_Constant
                    .Style = kDefaultConstStyle
                    .Color = RGB(128, 128, 128)
                    .Parm(1) = "0"
                End With
                idx = .Add(Indicator)
                .NodeLevel(idx) = 1
                
        End Select
    End With
    Set Indicator = Nothing
    Set Pane = Nothing
    
    FixupTree Chart
    
    If idxPane >= 0 Then AddStudyToChart = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.AddStudyToChart", eGDRaiseError_Raise
    
End Function

Public Sub FixupTree(Chart As cChart)
On Error GoTo ErrSection:

    Dim idx&, idxParent&
    Dim Indicator As cIndicator
    Dim Pane As cPane
    
    With Chart.Tree
        For idx = 1 To .Count
            If .NodeLevel(idx) = 0 Then
                ' fix pane stuff
                Set Pane = .Item(idx)
                With Pane
                
                End With
            Else
                ' fix indicator stuff
                Set Indicator = .Item(idx)
                #If 0 Then
                With Indicator
                    Select Case .DataType
                        Case eINDIC_Constant
                            .nNumSubsets = 0
                        Case eINDIC_BarData
                            .nNumSubsets = 4
                        Case eINDIC_None
                            .nNumSubsets = 0
                    End Select
                End With
                #End If
            
                If .NodeLevel(idx) = 1 Then
                    ' unlinked
                    ''Indicator.strLinkedTo = ""
                Else
                    ' link to parent
                    idxParent = .RelativeIndex(idx, eTREE_Parent)
                    ''Indicator.strLinkedTo = .Key(idxParent)
                End If
            End If
        Next
    End With
    
ErrExit:
    Set Indicator = Nothing
    Set Pane = Nothing
    Exit Sub
    
ErrSection:
    Set Indicator = Nothing
    Set Pane = Nothing
    RaiseError "mChartNav.FixUpTree", eGDRaiseError_Raise
    
End Sub

Public Sub RandomArray(ByVal hArray&)
On Error GoTo ErrSection:

    Dim i&, n&, r#, p#
    
    n = gdGetSize(hArray)
    p = 50
    For i = 0 To n - 1
        r = gdRandomNumber(-100, 100) / 10#
        p = p + r
        If p < 0 Or p > 100 Then
            p = p - r * 2
        End If
        gdSetNum hArray, i, p
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.RandomArray", eGDRaiseError_Raise
    
End Sub

' force visible charts to update
' nSymbolID: 0=all symbols, else just charts of specified symbol
Public Sub UpdateVisibleCharts(Optional ByVal eRedoMode As eChartRedoMode = eRedo9_ReloadData, _
            Optional ByVal nSymbolID& = 0, Optional ByVal nSpreadFuncId& = 0, _
            Optional ByVal bCheckTradeSym As Boolean = False)
On Error Resume Next

    Dim iForm&, nTradeSymbolID&, nChartSymbolID&
    Dim bRealTimeStarting As Boolean
    
    If frmStatus.Status = eStatus_Running Then
        If InStr(UCase(frmStatus.Caption), "RETRIEVING DATA") <> 0 Then
            bRealTimeStarting = True
        End If
    End If
        
    For iForm = 0 To Forms.Count - 1
        If g.bUnloading Then Exit For
        If IsFrmChart(Forms(iForm)) Then
            If Forms(iForm).OkayToRefresh Then
                If eRedoMode < 0 Then
                    ' only if normal size (causes problems if try to resize a maximized form)
                    If Forms(iForm).WindowState = 0 Then
                        FormResize Forms(iForm)
                    End If
                Else
                    With Forms(iForm).Chart
                        nTradeSymbolID = 0
                        nChartSymbolID = .SymbolID
                        
                        If bCheckTradeSym Then          '5841
                            nTradeSymbolID = ConvertToTradeSymbol(nChartSymbolID, .Bars(eBARS_DateTime, .Bars.Size - 1))
                        End If
                        
                        If nSpreadFuncId > 0 Then
                            .GenerateChart eRedo9_ReloadData
                        ElseIf nSymbolID = 0 Or nChartSymbolID = nSymbolID Or nTradeSymbolID = nSymbolID Then
                            If bRealTimeStarting Then
                                'to help identify any charts that have issues loading when real time starts up
                                frmStatus.AddDetail Forms(iForm).Chart.Bars.Prop(eBARS_Symbol) & " (" & Forms(iForm).Chart.Bars.Prop(eBARS_PeriodicityStr) & ")"
                            End If
                            .GenerateChart eRedoMode
                        End If
                        
                    End With
                    
                    Forms(iForm).Refresh
                End If
            End If
        End If
    Next
    
    If bRealTimeStarting Then
        frmStatus.AddDetail "Finished loading charts."
    End If
    
        
End Sub

Public Function RoundedValueStr(ByVal dValue#, ByVal dRange#) As String
On Error GoTo ErrSection:

    Dim strFormat$
    
    dRange = Abs(dRange)
    If dRange = 0 Then '(if even possible!)
        strFormat = "0.#"
    ElseIf dRange >= 100000000 Then
        strFormat = "0M"
        dValue = RoundNum(dValue / 1000000#)
    ElseIf dRange >= 10000000 Then
        strFormat = "0K"
        dValue = RoundNum(dValue / 1000#, -2)
    ElseIf dRange >= 1000000 Then
        strFormat = "0K"
        dValue = RoundNum(dValue / 1000#, -1)
    ElseIf dRange >= 100000 Then
        strFormat = "0K"
        dValue = RoundNum(dValue / 1000#)
    ElseIf dRange >= 10000 Then
        strFormat = "0"
        dValue = RoundNum(dValue, -2)
    ElseIf dRange >= 100 Then
        strFormat = "0"
    ElseIf dRange >= 10 Then
        strFormat = "0.0"
    ElseIf dRange >= 1 Then
        strFormat = "0.00"
    ElseIf dRange >= 0.1 Then
        strFormat = "0.000"
    ElseIf dRange >= 0.01 Then
        strFormat = "0.0000"
    Else
        strFormat = "0.00000"
    End If
    RoundedValueStr = Format(dValue, strFormat)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.RoundedValueStr", eGDRaiseError_Raise
    
End Function

Public Sub TileCharts(Optional bCascade As Boolean = False)
On Error Resume Next

    Dim i&, iForm&, nNumCharts&, nNumRows&, nPerRow&, s$
    Dim nLeft&, nTop&, nHeight&, nWidth&, nMinimizedHeight&
    Dim nOrigMdiWidth&, nOrigMdiHeight&, iLoops%
    Dim aCharts As New cGdArray
    Dim frm As Form
    Static nPrevRows&

    ' find all the non-minimized and non-detached charts
    For iForm = 0 To Forms.Count - 1
        If IsFrmChartMDI(Forms(iForm)) Then
            Set frm = Forms(iForm)
            If frm.Visible Then
                If frm.WindowState = 1 Then
                    nMinimizedHeight = frm.Height - frm.ScaleHeight
                Else
                    ' save caption in order to sort
                    aCharts.Add Trim(Parse((frm.Caption), ":", 1)) + Chr(9) + Str(iForm)
                End If
            End If
        End If
    Next

    ' see how many rows to do
    nNumCharts = aCharts.Size
    If bCascade Then
        nNumRows = -1
    ElseIf nNumCharts <= 1 Then
        nNumRows = nNumCharts
    Else
        ' default to 2 charts per row
        If nPrevRows = 0 Or nPrevRows > nNumCharts Then
            nPrevRows = (nNumCharts + 1) \ 2
        End If
        'nNumRows = AskBox("i=? ; g=n ; h=Tile Charts ; m=Number of Rows: ; d=" & CStr(nPrevRows))
        s = InfBox("Number of Rows:", "?", , "Tile Charts", , , , , , "s", CStr(nPrevRows))
        If Len(s) > 0 Then
            nNumRows = Abs(ValOfText(s))
            If nNumRows = 0 Then nNumRows = -1 '(cascade)
            ' sort array of captions
            aCharts.Sort eGdSort_IgnoreCase
        Else
            nNumRows = 0
        End If
        If nNumRows > 0 Then nPrevRows = nNumRows
        'If nNumRows > nNumCharts Then nNumRows = nNumCharts
    End If

    ' arrange minimized chart icons
    frmMain.Arrange vbArrangeIcons
    
    If nNumRows < 0 Then
        ' cascade charts
        LockWindowUpdate frmMain.hWnd
        frmMain.Arrange vbCascade
        DoEvents
        LockWindowUpdate 0
    ElseIf nNumRows > 0 Then
        'DoEvents
        LockWindowUpdate frmMain.hWnd
        
        nOrigMdiWidth = frmMain.ScaleWidth
        nOrigMdiHeight = frmMain.ScaleHeight
        For iLoops = 1 To 2
            ' place charts
            nHeight = (frmMain.ScaleHeight - nMinimizedHeight) \ nNumRows
            nPerRow = (nNumCharts - 1) \ nNumRows + 1 ' # columns
            nWidth = frmMain.ScaleWidth \ nPerRow
            nLeft = 0
            nTop = -nHeight
            For i = 0 To nNumCharts - 1
                iForm = Val(Parse(aCharts(i), Chr(9), 2))
                If i Mod nPerRow = 0 Then
                    ' start a new row
                    nTop = nTop + nHeight
                    nLeft = -nWidth
                    ' if last row, shove to right
                    If nNumCharts - i < nPerRow Then
                        'nLeft = nLeft + nWidth * (nPerRow - (nNumCharts - i))
                    End If
                End If
                nLeft = nLeft + nWidth
                Set frm = Forms(iForm)
                With frm
                    ' temporarily set g.bUnloading so chart
                    ' resizing will not yet regenerate the chart
                    g.bUnloading = True
                    .WindowState = 0
                    .Move nLeft, nTop, nWidth, nHeight
                    g.bUnloading = False
                End With
            Next
            Set frm = Nothing
            
            ' if size of MDI client area did NOT change
            ' (due to scrollbars disappearing), then done
            DoEvents
            If nOrigMdiWidth = frmMain.ScaleWidth And _
                nOrigMdiHeight = frmMain.ScaleHeight Then
                    Exit For
            End If
            ' otherwise need to do this one more time
        Next
        
        ' now call resize for all charts so they will
        ' all regenerate now (this way, it's just once)
        For i = 0 To nNumCharts - 1
            iForm = Val(Parse(aCharts(i), Chr(9), 2))
            FormResize Forms(iForm)
        Next
        
        DoEvents
        LockWindowUpdate 0
    End If

End Sub

Private Sub GetExistingCompSym(Chart As cChart, aSymbols As cGdArray, idxIndPricePane As Long, _
    idxIndOtherPane As Long, ByVal bPercentCompOnly As Boolean)
On Error GoTo ErrSection:

    Dim idx&, iTemp&, strChartSymbol$, i&
    Dim Ind As cIndicator
    Dim Pane As cPane

    If Chart Is Nothing Or aSymbols Is Nothing Then Exit Sub
    If IsMissing(idxIndPricePane) Or IsMissing(idxIndOtherPane) Then Exit Sub

    'reset
    idxIndPricePane = 0
    idxIndOtherPane = 0
    
    If Len(Chart.SpreadSymbols) > 0 Then
        strChartSymbol = Parse(Chart.Symbol, ",", 1)        'grab first symbol of a spread chart - 5798
    Else
        strChartSymbol = Chart.Symbol
    End If
    
    'build list of existing comparison symbols
    With Chart.Tree
        For idx = 1 To .Count
            If .NodeLevel(idx) = 1 Then
                Set Ind = Chart.Tree(idx)
                If Ind.DataType = eINDIC_BarData Then
                    If UCase(.Key(idx)) <> "PRICE" Then
                        iTemp = .RelativeIndex(idx, eTREE_Parent)
                        Set Pane = Chart.Tree(iTemp)
                        If UCase(.Key(iTemp)) = "PRICE PANE" Then
                            'save first found indicators index for reuse
                            If idxIndPricePane = 0 Then idxIndPricePane = idx
                        Else
                            'save first found indicators index for reuse
                            If idxIndOtherPane = 0 Then idxIndOtherPane = idx
                        End If
                    End If
                    
                    If Ind.Bars.Prop(eBARS_Symbol) <> strChartSymbol Then
                        If Ind.Bars.Prop(eBARS_SymbolID) <> 0 Then  '=0 if hidden comparison symbol
                            If Ind.Display Then
                                iTemp = flexChecked
                            Else
                                iTemp = flexUnchecked
                            End If
                            
                            If bPercentCompOnly Then
                                'return only percent change comparison symbols
                                If Not Pane Is Nothing Then
                                    If Pane.PaneLogFlag = ePANE_LogFlagPercent Then
                                        'add to array to pass over to comparison symbol dialog
                                        aSymbols.Add Ind.Bars.Prop(eBARS_Symbol) & "|" & Ind.Color & "|" & iTemp
                                    End If
                                End If
                            Else
                                'return all comparison symbols
                                'add to array to pass over to comparison symbol dialog
                                aSymbols.Add Ind.Bars.Prop(eBARS_Symbol) & "|" & Ind.Color & "|" & iTemp
                            End If
                        End If
                    End If
                End If
                
                Set Ind = Nothing
                Set Pane = Nothing
            End If
        Next
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.GetExistingCompSym"

End Sub

Public Sub PercentChangeChartNew(ChartIn As cChart, aSymbols As cGdArray, _
    Optional ByVal strSymbolIn$ = "", Optional ByVal nPeriodicityIn& = -1)
On Error GoTo ErrSection:

    Dim strFile$, strSymbol$
    Dim i&, idx&, idxComp&, iTemp&, iRec&
    
    Dim frm As frmChart
    Dim ChartNew As cChart
    Dim Chart As cChart
    
    Dim Ind As cIndicator
    Dim Pane As cPane
    
    Dim bLocked As Boolean
    
    If aSymbols Is Nothing Then Exit Sub
    If aSymbols.Size = 0 Then Exit Sub
    
    If ChartIn Is Nothing Then
        If Not ActiveChart Is Nothing Then Set Chart = ActiveChart.Chart
    Else
        Set Chart = ChartIn
    End If
    If Chart Is Nothing Then Exit Sub
    
    If Len(strSymbolIn) > 0 Then
        strSymbol = strSymbolIn
    ElseIf Len(Chart.SpreadSymbols) > 0 Then
        strSymbol = Parse(Chart.Symbol, ",", 1) 'grab first symbol of a spread chart - 5798
    Else
        strSymbol = Chart.Symbol
    End If

    strFile = Chart.TemplateCopy("MyPercentChange", "", True)
    If Len(strFile) = 0 Then Exit Sub
    
    Set frm = New frmChart
    Set ChartNew = frm.Chart
    
    With ChartNew
        .TemplateLoad , , strFile
        .SetSymbol strSymbol
        If nPeriodicityIn = -1 Then
            .ChangeBarPeriod Chart.Periodicity, False
        Else
            .ChangeBarPeriod nPeriodicityIn, False
        End If
        .ShowTrades = 0
        .ShowToolbar = 0
    End With
    strSymbol = ""
    
    With ChartNew.Tree
        .Item("PRICE").DisplayType = 0
        idx = .Index("PRICE PANE")
        
        For i = .Count To idx Step -1
            If .NodeLevel(i) = 0 Then
                If i <> idx Then .Remove i
            End If
        Next
        Set Pane = .Item(idx)
        Pane.Display = True
        Pane.Scaling = ePANE_ScaleModeAuto
        Pane.PaneLogFlag = ePANE_LogFlagPercent
        Pane.DisplayFormat = ePANE_PriceFormat
        
        idxComp = .Index("PRICE")
        For iTemp = 0 To aSymbols.Size - 1
            strSymbol = Parse(aSymbols(iTemp), "|", 1)
            i = Val(Parse(aSymbols(iTemp), "|", 2))
            If strSymbol <> ChartNew.Symbol Then
                iRec = g.SymbolPool.PoolRecForSymbol(strSymbol, True)
                Set Ind = New cIndicator
                With Ind
                    .Name = g.SymbolPool.Symbol(iRec)
                    .Display = True
                    .DataType = eINDIC_BarData
                    .Color = i
                    .DisplayType = eINDIC_Line
                End With
                idx = .Add(Ind, , idxComp, eTREE_NextSibling)
            End If
        Next
    End With
    idxComp = 0
    
    frm.SkipFocusFix = True
    frm.Chart.TemplateSave      '5007
    ActiveChartFormSet frm
    
    bLocked = LockWindowUpdate(frmMain.hWnd)
        ShowForm frm
        MoveFocus frm.pbChart
    If bLocked Then LockWindowUpdate 0

ErrExit:
    Set frm = Nothing
    If Len(strFile) > 0 Then KillFile strFile
    
    Exit Sub

ErrSection:
    RaiseError "mChartNav.PercentChangeChartNew"

End Sub

Private Sub PercentChangePane(Chart As cChart, aSymbols As cGdArray, idxComp&, idxIndOtherPane&, nRec&)
On Error GoTo ErrSection:

    Dim i&, idx&, iTemp&
    Dim strSymbol$, strChartSymbol$
    
    Dim Ind As cIndicator
    Dim Pane As cPane
    
    If Chart Is Nothing Or aSymbols Is Nothing Then Exit Sub
    If aSymbols.Size = 0 Then Exit Sub
    
    If IsMissing(idxComp) Or IsMissing(idxIndOtherPane) Or IsMissing(nRec) Then Exit Sub
    
    idxComp = 0     'reset
    
    With Chart.Tree
        If idxIndOtherPane = 0 Then
            ' add to new pane (after price pane)
            Set Pane = New cPane
            idx = .Index("PRICE PANE")
            idx = .Add(Pane, , idx, eTREE_NextSibling)
            
            Set Ind = New cIndicator
            With Ind
                .Name = "$DJIA" '(for now)
                .Display = True
                .DataType = eINDIC_BarData
                .Color = RGB(192, 0, 192)
            End With
            idxComp = .Add(Ind, , idx, eTREE_FirstChild)
        Else
            idx = .RelativeIndex(idxIndOtherPane, eTREE_Parent)
            Set Pane = .Item(idx)
            Set Ind = Nothing
            
            'remove all existing symbols from pane
            For i = .Count To idx Step -1
                If .NodeLevel(i) > 0 Then
                    If .RelativeIndex(i, eTREE_Parent) = idx Then
                        Set Ind = .Item(i)
                        If Not Ind Is Nothing Then
                            If Ind.DataType = eINDIC_BarData Then .Remove i
                        End If
                    End If
                End If
            Next
            
            Set Ind = New cIndicator
            With Ind
                .Name = "$DJIA" '(for now)
                .Display = True
                .DataType = eINDIC_BarData
                .Color = RGB(192, 0, 192)
            End With
            idxComp = .Add(Ind, , idx, eTREE_LastChild)
            Chart.RedoMode = eRedo9_ReloadData
        End If
        
        Pane.Display = True
        Pane.Scaling = ePANE_ScaleModeAuto
        Pane.PaneLogFlag = ePANE_LogFlagPercent
        Pane.DisplayFormat = ePANE_PriceFormat
        
        'set the first ind to same color & symbol as chart's Price indicator
        Ind.DisplayType = eINDIC_Line
        
        If Len(Chart.SpreadSymbols) > 0 Then
            strChartSymbol = Parse(Chart.Symbol, ",", 1)   'grab first symbol of a spread chart - 5798
        Else
            strChartSymbol = Chart.Symbol
        End If
        
        Ind.Name = strChartSymbol
        idx = .Index("PRICE")
        If idx > 0 Then Ind.Color = Chart.Tree(idx).Color
    
        'loop through and add comparison symbols
        For iTemp = 0 To aSymbols.Size - 1
            strSymbol = Parse(aSymbols(iTemp), "|", 1)
            i = Val(Parse(aSymbols(iTemp), "|", 2))
            If strSymbol <> strChartSymbol Then
                nRec = g.SymbolPool.PoolRecForSymbol(strSymbol, True)
                Set Ind = New cIndicator
                With Ind
                    .Name = g.SymbolPool.Symbol(nRec)
                    .Display = True
                    .DataType = eINDIC_BarData
                    .Color = i
                    .DisplayType = eINDIC_Line
                End With
                idx = .Add(Ind, , idxComp, eTREE_NextSibling)
            End If
        Next
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.PercentChangePane"

End Sub

Private Sub CompSymOverlay(Chart As cChart, idxComp&, ByVal idxIndPricePane&, _
    Optional ByVal eDisplay As eIndicatorDisplayType = eINDIC_Line)
On Error GoTo ErrSection:

    Dim idx&
    Dim Ind As cIndicator

    If Chart Is Nothing Then Exit Sub
    If IsMissing(idxComp) Or IsMissing(idxIndPricePane) Then Exit Sub

    ' create new comparison indicator with defaults
    Set Ind = New cIndicator
    With Ind
        .Name = "$DJIA" '(for now)
        .Display = True
        .DataType = eINDIC_BarData
        .Color = RGB(192, 0, 192)
    End With
    
    With Chart.Tree
        If idxIndPricePane = 0 Then
            ' add to price pane as overlayed indicator
            Ind.DisplayType = eDisplay      'eINDIC_Line
            Ind.Overlayed = True
            idx = .Index("PRICE PANE")
            idxComp = .Add(Ind, , idx, eTREE_LastChild)
            .NodeLevel(idxComp) = 1
        Else
            idxComp = idxIndPricePane
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.CompSymOverlay"

End Sub

Private Sub CompSymPane(Chart As cChart, idxComp&, idxIndOtherPane&)

    Dim idx&
    Dim Ind As cIndicator
    Dim Pane As cPane

    If Chart Is Nothing Then Exit Sub
    If IsMissing(idxComp) Or IsMissing(idxIndOtherPane) Then Exit Sub
    
    ' create new comparison indicator with defaults
    Set Ind = New cIndicator
    With Ind
        .Name = "$DJIA" '(for now)
        .Display = True
        .DataType = eINDIC_BarData
        .Color = RGB(192, 0, 192)
    End With
    
    With Chart.Tree
        ' add to new pane (after price pane)
        If idxIndOtherPane = 0 Then
            Ind.DisplayType = eINDIC_OHLC
            Set Pane = New cPane
            Pane.Display = True
            Pane.Scaling = ePANE_ScaleModeAuto
            Pane.DisplayFormat = ePANE_PriceFormat
            idx = .Index("PRICE PANE")
            idx = .Add(Pane, , idx, eTREE_NextSibling)
            idxComp = .Add(Ind, , idx, eTREE_FirstChild)
        Else
            idxComp = idxIndOtherPane
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.CompSymPane"

End Sub

' returns IDX of comp symbol
Public Function AddCompSymbol(Chart As cChart, ByVal bOnlyIfNotYetExist As Boolean, _
                    Optional ByVal bFromHotKey As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim i&, idx&, idxComp&, nRec&
    Dim idxIndPricePane&, idxIndOtherPane&
    
    Dim strAnswer$
    Dim bFixedSpread As Boolean

    Dim Ind As cIndicator
    Dim Pane As cPane
    Dim aSymbols As New cGdArray
    
           
    With Chart.Tree
        ' first see if using 'C' to change the 2nd spread symbol
        ' (if using spread indicator with hidden price pane)
        If bFromHotKey Then
            bFixedSpread = FixSpreadSymbols(Chart, aSymbols(0))
        End If
        
'JM 12-12-2011: Original implementation replaces comparison symbol with new one rather than add it.
'   When decision was made to always add comparison symbols (aardvark 6497), the bOnlyIfNotYetExist
'   flag was added to preserve ability to replace comparison symbols (JIC - just in case).
'   As of 12-12-2011, this function is never called with this flag equals to true.

        If bOnlyIfNotYetExist Then
            ' get list of comparison symbols already on chart
            GetExistingCompSym Chart, aSymbols, idxIndPricePane, idxIndOtherPane, False
        ElseIf Chart.CompSymType(eCompSym_PercentPane) = eCompSym_PercentPane Then
            ' aardvark 6542 specifies adding percent change comparison symbols to existing pane
            GetExistingCompSym Chart, aSymbols, idxIndPricePane, idxIndOtherPane, True
        End If
        
        If Not bFixedSpread And .Exists("PRICE PANE") Then
            ' ask where to put it (O=overlay, N=new pane, P=% change current chart, C=%change new chart)
            aSymbols.Sort eGdSort_Default Or eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
            strAnswer = frmCompSymbol.ShowMe(aSymbols, Chart)
            
            idxComp = 0
            Select Case strAnswer
                Case "O"
                    nRec = g.SymbolPool.PoolRecForSymbol(aSymbols(0), True)
                    If Not bOnlyIfNotYetExist Then idxIndPricePane = 0
                    CompSymOverlay Chart, idxComp, idxIndPricePane
                
                Case "B"
                    nRec = g.SymbolPool.PoolRecForSymbol(aSymbols(0), True)
                    If Not bOnlyIfNotYetExist Then idxIndPricePane = 0
                    CompSymOverlay Chart, idxComp, idxIndPricePane, eINDIC_OHLC             '6517
                
                Case "N"
                    nRec = g.SymbolPool.PoolRecForSymbol(aSymbols(0), True)
                    CompSymPane Chart, idxComp, idxIndOtherPane
                    If bOnlyIfNotYetExist Then
                        If idxIndPricePane > 0 Then Chart.RedoMode = eRedo5_RecalcInd       '5798
                    End If
                
                Case "C"
                    PercentChangeChartNew Chart, aSymbols
                    
                Case "P"
                    PercentChangePane Chart, aSymbols, idxComp, idxIndOtherPane, nRec

            End Select
        End If
        
        ' fix name whether added new or using existing
        If idxComp > 0 Then
            Set Ind = Chart.Tree(idxComp)
            If strAnswer <> "P" Then
                ' TLB 4/11/2012: check if an external symbol (from hard drive)
                If InStr(aSymbols(0), "|") > 0 Then
                    Ind.Name = aSymbols(0)
                Else
                    Ind.Name = g.SymbolPool.Symbol(nRec)
                End If
            End If
            If Not bFixedSpread Then
                Ind.Display = True
                idx = .RelativeIndex(idxComp, eTREE_Root)
                If idx > 0 Then
                    Set Pane = Chart.Tree(idx)
                    If Not Pane Is Nothing Then
                        Pane.Display = True
                    End If
                End If
            End If
        End If
    End With
    
    Set Pane = Nothing
    Set Ind = Nothing
    AddCompSymbol = idxComp

ErrExit:
    Exit Function
    
ErrSection:
    Set Pane = Nothing
    Set Ind = Nothing
    RaiseError "mChartNav.AddCompSymbol", eGDRaiseError_Raise
    
End Function

Public Sub LoadChartGlobals(Optional ByVal strTheme As String)
On Error GoTo ErrSection:

    'JM 11-19-2015: need to to this because theme parameter is optional
    If Len(strTheme) = 0 Then strTheme = GetIniFileProperty("TradenavTheme", "", "General", g.strIniFile)
    
    If Len(strTheme) = 0 Then
        g.ChartGlobals.nFontStyle = GetIniFileProperty("FontStyle", 1, "Charting", g.strIniFile)
        g.ChartGlobals.strFontName = GetIniFileProperty("FontName", "arial", "Charting", g.strIniFile)
        g.ChartGlobals.nChartBackColor = GetIniFileProperty("ChartBackColor", RGB(255, 255, 255), "Charting", g.strIniFile)
        g.ChartGlobals.nChartForeColor = GetIniFileProperty("ChartForeColor", RGB(128, 128, 128), "Charting", g.strIniFile)
        g.ChartGlobals.nBorderBackColor = GetIniFileProperty("BorderBackColor", RGB(224, 224, 224), "Charting", g.strIniFile)
        g.ChartGlobals.nBorderForeColor = GetIniFileProperty("BorderForeColor", RGB(0, 0, 128), "Charting", g.strIniFile)
        g.ChartGlobals.eDefaultBarsStyle = GetIniFileProperty("DefaultBarsStyle", eINDIC_Auto, "Charting", g.strIniFile)
        g.ChartGlobals.eDefaultIndStyle = GetIniFileProperty("DefaultIndStyle", eINDIC_Thin, "Charting", g.strIniFile)
        g.ChartGlobals.bHideScrollbars = GetIniFileProperty("HideScrollbars", False, "Charting", g.strIniFile)
        g.ChartGlobals.nChartGradientColor = GetIniFileProperty("GradientColor", -1, "Charting", g.strIniFile)
        g.ChartGlobals.nUseGradient = GetIniFileProperty("UseGradient", 0, "Charting", g.strIniFile)
        g.ChartGlobals.nShortColor = GetIniFileProperty("ShortColor", vbRed, "Charting", g.strIniFile)
        g.ChartGlobals.nLongColor = GetIniFileProperty("LongColor", vbBlue, "Charting", g.strIniFile)
        g.ChartGlobals.nLossColor = GetIniFileProperty("LossColor", vbRed, "Charting", g.strIniFile)
        g.ChartGlobals.nWinColor = GetIniFileProperty("WinColor", RGB(0, 128, 0), "Charting", g.strIniFile)
    ElseIf strTheme = "Charcoal" Then
        g.ChartGlobals.nFontStyle = GetIniFileProperty("FontStyle", 0, strTheme, g.strIniFile)
        g.ChartGlobals.strFontName = GetIniFileProperty("FontName", "Microsoft Sans Serif", strTheme, g.strIniFile)
        g.ChartGlobals.nChartBackColor = GetIniFileProperty("ChartBackColor", 0, strTheme, g.strIniFile)    'set to zero to trigger special green color for indicators (see fixcolors in cChart)
        g.ChartGlobals.nChartForeColor = GetIniFileProperty("ChartForeColor", RGB(128, 128, 128), strTheme, g.strIniFile)
        g.ChartGlobals.nBorderBackColor = GetIniFileProperty("BorderBackColor", 0, strTheme, g.strIniFile)
        g.ChartGlobals.nBorderForeColor = GetIniFileProperty("BorderForeColor", RGB(128, 128, 128), strTheme, g.strIniFile)
        g.ChartGlobals.eDefaultBarsStyle = GetIniFileProperty("DefaultBarsStyle", eINDIC_Thin, strTheme, g.strIniFile)
        g.ChartGlobals.eDefaultIndStyle = GetIniFileProperty("DefaultIndStyle", eINDIC_Thin, strTheme, g.strIniFile)
        g.ChartGlobals.bHideScrollbars = GetIniFileProperty("HideScrollbars", True, strTheme, g.strIniFile)
        g.ChartGlobals.nChartGradientColor = GetIniFileProperty("GradientColor", -1, strTheme, g.strIniFile)
        g.ChartGlobals.nUseGradient = GetIniFileProperty("UseGradient", 0, strTheme, g.strIniFile)
        g.ChartGlobals.nShortColor = GetIniFileProperty("ShortColor", vbRed, strTheme, g.strIniFile)
        g.ChartGlobals.nLongColor = GetIniFileProperty("LongColor", vbCyan, strTheme, g.strIniFile)
        g.ChartGlobals.nLossColor = GetIniFileProperty("LossColor", vbRed, strTheme, g.strIniFile)
        g.ChartGlobals.nWinColor = GetIniFileProperty("WinColor", vbGreen, strTheme, g.strIniFile)
    ElseIf strTheme = "Ivory" Then
        g.ChartGlobals.nFontStyle = GetIniFileProperty("FontStyle", 0, strTheme, g.strIniFile)
        g.ChartGlobals.strFontName = GetIniFileProperty("FontName", "Microsoft Sans Serif", strTheme, g.strIniFile)
        g.ChartGlobals.nChartBackColor = GetIniFileProperty("ChartBackColor", vbWhite, strTheme, g.strIniFile)
        g.ChartGlobals.nChartForeColor = GetIniFileProperty("ChartForeColor", 0, strTheme, g.strIniFile)
        g.ChartGlobals.nBorderBackColor = GetIniFileProperty("BorderBackColor", vbWhite, strTheme, g.strIniFile)
        g.ChartGlobals.nBorderForeColor = GetIniFileProperty("BorderForeColor", 0, strTheme, g.strIniFile)
        g.ChartGlobals.eDefaultBarsStyle = GetIniFileProperty("DefaultBarsStyle", eINDIC_Thin, strTheme, g.strIniFile)
        g.ChartGlobals.eDefaultIndStyle = GetIniFileProperty("DefaultIndStyle", eINDIC_Thin, strTheme, g.strIniFile)
        g.ChartGlobals.bHideScrollbars = GetIniFileProperty("HideScrollbars", True, strTheme, g.strIniFile)
        g.ChartGlobals.nChartGradientColor = GetIniFileProperty("GradientColor", -1, strTheme, g.strIniFile)
        g.ChartGlobals.nUseGradient = GetIniFileProperty("UseGradient", 0, strTheme, g.strIniFile)
        g.ChartGlobals.nShortColor = GetIniFileProperty("ShortColor", vbRed, strTheme, g.strIniFile)
        g.ChartGlobals.nLongColor = GetIniFileProperty("LongColor", vbBlue, strTheme, g.strIniFile)
        g.ChartGlobals.nLossColor = GetIniFileProperty("LossColor", vbRed, strTheme, g.strIniFile)
        g.ChartGlobals.nWinColor = GetIniFileProperty("WinColor", RGB(0, 128, 0), strTheme, g.strIniFile)
    Else
        g.ChartGlobals.nFontStyle = GetIniFileProperty("FontStyle", 1, strTheme, g.strIniFile)
        g.ChartGlobals.strFontName = GetIniFileProperty("FontName", "arial", strTheme, g.strIniFile)
        g.ChartGlobals.nChartBackColor = GetIniFileProperty("ChartBackColor", RGB(255, 255, 255), strTheme, g.strIniFile)
        g.ChartGlobals.nChartForeColor = GetIniFileProperty("ChartForeColor", RGB(128, 128, 128), strTheme, g.strIniFile)
        g.ChartGlobals.nBorderBackColor = GetIniFileProperty("BorderBackColor", RGB(224, 224, 224), strTheme, g.strIniFile)
        g.ChartGlobals.nBorderForeColor = GetIniFileProperty("BorderForeColor", RGB(0, 0, 128), strTheme, g.strIniFile)
        g.ChartGlobals.eDefaultBarsStyle = GetIniFileProperty("DefaultBarsStyle", eINDIC_Auto, strTheme, g.strIniFile)
        g.ChartGlobals.eDefaultIndStyle = GetIniFileProperty("DefaultIndStyle", eINDIC_Thin, strTheme, g.strIniFile)
        g.ChartGlobals.bHideScrollbars = GetIniFileProperty("HideScrollbars", False, strTheme, g.strIniFile)
        g.ChartGlobals.nChartGradientColor = GetIniFileProperty("GradientColor", -1, strTheme, g.strIniFile)
        g.ChartGlobals.nUseGradient = GetIniFileProperty("UseGradient", 0, strTheme, g.strIniFile)
        g.ChartGlobals.nLossColor = GetIniFileProperty("LossColor", vbRed, strTheme, g.strIniFile)
        g.ChartGlobals.nWinColor = GetIniFileProperty("WinColor", RGB(0, 128, 0), strTheme, g.strIniFile)
        g.ChartGlobals.nShortColor = GetIniFileProperty("ShortColor", vbRed, strTheme, g.strIniFile)
        g.ChartGlobals.nLongColor = GetIniFileProperty("LongColor", vbBlue, strTheme, g.strIniFile)
    End If
    
    If g.ChartGlobals.nChartGradientColor = -1 Then g.ChartGlobals.nChartGradientColor = GradientDefault

    With g.ChartGlobals
        .bSplitsRolls = GetIniFileProperty("SplitsRolls", True, "Charting", g.strIniFile)
        .nSplitRollColor = GetIniFileProperty("SplitRollColor", RGB(0, 192, 192), "Charting", g.strIniFile)
        .bFloatingTips = GetIniFileProperty("FloatingTips", True, "Charting", g.strIniFile)
        .bChartTips = GetIniFileProperty("ChartTips", True, "Charting", g.strIniFile)
        
        ' TLB 5/16/2014: with new grid dots style, reset the default Global ChartForeColor (but just once)
        If GetIniFileProperty("GcfcDefaultCheck", 0, "Charting", g.strIniFile) = 0 Then
            SetIniFileProperty "GcfcDefaultCheck", 1, "Charting", g.strIniFile
            If .nChartForeColor = RGB(192, 192, 192) Then ' old default
                .nChartForeColor = RGB(128, 128, 128) ' new default (darker)
            End If
        End If
        
        .eDefaultHorzStyle = GetIniFileProperty("DefaultHorzStyle", eINDIC_Dot, "Charting", g.strIniFile)
        .eProfitLineStyle = GetIniFileProperty("ProfitLineStyle", eANNOT_Thin, "Charting", g.strIniFile)
        .eDefaultLabelMode = GetIniFileProperty("DefaultLabelMode", eINDIC_scale, "Charting", g.strIniFile)
        .eDefaultAnnotStyle = GetIniFileProperty("DefaultAnnotStyle", eANNOT_Default, "Charting", g.strIniFile)
        .nFontSize = GetIniFileProperty("FontSizePts", 8, "Charting", g.strIniFile)
        
        .bLogModeDraw = GetIniFileProperty("LogDrawMode", False, "Charting", g.strIniFile)
        .bAutoChartData = GetIniFileProperty("AutoChartData", True, "Charting", g.strIniFile)
        .bChartDataSingleBar = GetIniFileProperty("ChartDataSingleBar", True, "Charting", g.strIniFile)
        
        .nSquareBars = GetIniFileProperty("SquareBars", 1, "Charting", g.strIniFile)
        .dSquareTicks = GetIniFileProperty("SquareTicks", 100, "Charting", g.strIniFile)
        .aSquareSymList.FromFile g.strAppPath + "\SquareChart.dat"
        .aSquareSymList.Sort eGdSort_DeleteDuplicates Or eGdSort_IgnoreCase Or eGdSort_DeleteNullValues
        
        .eBadTickShowMode = GetIniFileProperty("BadTickShowMode", 0, "Charting", g.strIniFile)
        .nBadTickColor = GetIniFileProperty("BadTickColor", vbRed, "Charting", g.strIniFile)
        .eBadTickMarker = GetIniFileProperty("BadTickMarker", eCNI_Arrow, "Charting", g.strIniFile)
        .eBadTickMarkerDir = GetIniFileProperty("BadTickMarkerDir", eCNI_North, "Charting", g.strIniFile)
        .nBadTickMarkerFill = GetIniFileProperty("BadTickMarkerFill", 0, "Charting", g.strIniFile)
        
        .eChartMode = GetIniFileProperty("ChartMode", eMode_Move, "Charting", g.strIniFile)
        .ePrevChartMode = .eChartMode
        .eScaleMode = GetIniFileProperty("ScaleMode", ePANE_ScaleModeAuto, "Charting", g.strIniFile)
        .eDragModeY = GetIniFileProperty("DragModeY", eDragModeY_Each, "Charting", g.strIniFile)
        .nHideAnnotations = GetIniFileProperty("HideAnnotations", 0, "Charting", g.strIniFile)
        .nMagnetValue = GetIniFileProperty("MagnetValue", 5, "Charting", g.strIniFile)
        
        .strCPCRoot = g.strAppPath
        
        If .eDefaultBarsStyle <= 0 Or .eDefaultBarsStyle > eINDIC_Auto Then .eDefaultBarsStyle = eINDIC_Auto
        If .eDefaultIndStyle <= 0 Then .eDefaultIndStyle = eINDIC_Thin
        If .eDefaultHorzStyle <= 0 Then .eDefaultHorzStyle = eINDIC_Dot
        If .eDefaultAnnotStyle <= 0 Then .eDefaultAnnotStyle = eANNOT_Thin
        
        If .nFontSize < 4 Or .nFontSize > 30 Then
            .nFontSize = 8
        End If
        
        If FileExist(g.strAppPath & "\ShowCPC.flg") Then
            .bChartModeAutoSize = GetIniFileProperty("ChartModeAutoSize", False, "Charting", g.strIniFile)
            .bExtForecastBars = GetIniFileProperty("ExtForecastBars", True, "Charting", g.strIniFile)
        Else
            .bChartModeAutoSize = GetIniFileProperty("ChartModeAutoSize", True, "Charting", g.strIniFile)
            .bExtForecastBars = GetIniFileProperty("ExtForecastBars", False, "Charting", g.strIniFile)
        End If
        
        .bSnapToDots = GetIniFileProperty("SnapToDots", False, "Charting", g.strIniFile)
        If .bSnapToDots Then .bChartModeAutoSize = False

        
        '07-13-2011: Glen wants flag file taken out.
'        .bOverlayAll = True             'FileExist(g.strAppPath & "\OverlayAll.flg")
    End With
    
    If strTheme = "Charcoal" Then
        If IsBlueRange(g.ChartGlobals.nWinColor) Then
            g.ChartGlobals.nWinColor = vbCyan
        ElseIf IsGreenRange(g.ChartGlobals.nWinColor, True) Then
            g.ChartGlobals.nWinColor = vbGreen
        End If
        If IsBlueRange(g.ChartGlobals.nLossColor) Then
            g.ChartGlobals.nLossColor = vbCyan
        ElseIf IsGreenRange(g.ChartGlobals.nLossColor, True) Then
            g.ChartGlobals.nLossColor = vbGreen
        End If
        If IsBlueRange(g.ChartGlobals.nLongColor) Then
            g.ChartGlobals.nLongColor = vbCyan
        ElseIf IsGreenRange(g.ChartGlobals.nLongColor, True) Then
            g.ChartGlobals.nLongColor = vbGreen
        End If
        If IsBlueRange(g.ChartGlobals.nShortColor) Then
            g.ChartGlobals.nShortColor = vbCyan
        ElseIf IsGreenRange(g.ChartGlobals.nShortColor, True) Then
            g.ChartGlobals.nShortColor = vbGreen
        End If
    End If
    
    ExitFavoritesLoad
    TSOFavoritesLoad

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.LoadChartGlobals", eGDRaiseError_Raise
    
End Sub

Public Sub SaveChartGlobals(Optional ByVal strTheme As String)
On Error GoTo ErrSection:

    Dim bDefault As Boolean
    
    'JM 11-19-2015: need to to this because theme parameter is optional
    If Len(strTheme) = 0 Then
        strTheme = GetIniFileProperty("TradenavTheme", "", "General", g.strIniFile)
        If Len(strTheme) > 0 Then
            'JM 12-24-2015: user could have changed theme (theoretically should not get here when on XP)
            If g.nColorTheme = kDarkThemeColor Then
                strTheme = "Charcoal"
            ElseIf g.nColorTheme = vbWhite Then
                strTheme = "Ivory"
            Else
                strTheme = "Classic"
            End If
        End If
    End If
    
    If Len(strTheme) > 0 Then
        Call SetIniFileProperty("FontStyle", g.ChartGlobals.nFontStyle, strTheme, g.strIniFile)
        Call SetIniFileProperty("FontName", g.ChartGlobals.strFontName, strTheme, g.strIniFile)
        Call SetIniFileProperty("ChartBackColor", g.ChartGlobals.nChartBackColor, strTheme, g.strIniFile)
        Call SetIniFileProperty("ChartForeColor", g.ChartGlobals.nChartForeColor, strTheme, g.strIniFile)
        Call SetIniFileProperty("BorderBackColor", g.ChartGlobals.nBorderBackColor, strTheme, g.strIniFile)
        Call SetIniFileProperty("BorderForeColor", g.ChartGlobals.nBorderForeColor, strTheme, g.strIniFile)
        Call SetIniFileProperty("DefaultBarsStyle", g.ChartGlobals.eDefaultBarsStyle, strTheme, g.strIniFile)
        Call SetIniFileProperty("DefaultIndStyle", g.ChartGlobals.eDefaultIndStyle, strTheme, g.strIniFile)
        Call SetIniFileProperty("HideScrollbars", g.ChartGlobals.bHideScrollbars, strTheme, g.strIniFile)
        If g.ChartGlobals.nChartGradientColor = GradientDefault Then
            Call SetIniFileProperty("GradientColor", -1, strTheme, g.strIniFile)
        Else
            Call SetIniFileProperty("GradientColor", g.ChartGlobals.nChartGradientColor, strTheme, g.strIniFile)
        End If
        Call SetIniFileProperty("UseGradient", g.ChartGlobals.nUseGradient, strTheme, g.strIniFile)
        Call SetIniFileProperty("ShortColor", g.ChartGlobals.nShortColor, strTheme, g.strIniFile)
        Call SetIniFileProperty("LongColor", g.ChartGlobals.nLongColor, strTheme, g.strIniFile)
        Call SetIniFileProperty("LossColor", g.ChartGlobals.nLossColor, strTheme, g.strIniFile)
        Call SetIniFileProperty("WinColor", g.ChartGlobals.nWinColor, strTheme, g.strIniFile)
    Else
        Call SetIniFileProperty("FontStyle", g.ChartGlobals.nFontStyle, "Charting", g.strIniFile)
        Call SetIniFileProperty("FontName", g.ChartGlobals.strFontName, "Charting", g.strIniFile)
        Call SetIniFileProperty("ChartBackColor", g.ChartGlobals.nChartBackColor, "Charting", g.strIniFile)
        Call SetIniFileProperty("ChartForeColor", g.ChartGlobals.nChartForeColor, "Charting", g.strIniFile)
        Call SetIniFileProperty("BorderBackColor", g.ChartGlobals.nBorderBackColor, "Charting", g.strIniFile)
        Call SetIniFileProperty("BorderForeColor", g.ChartGlobals.nBorderForeColor, "Charting", g.strIniFile)
        Call SetIniFileProperty("DefaultBarsStyle", g.ChartGlobals.eDefaultBarsStyle, "Charting", g.strIniFile)
        Call SetIniFileProperty("DefaultIndStyle", g.ChartGlobals.eDefaultIndStyle, "Charting", g.strIniFile)
        Call SetIniFileProperty("HideScrollbars", g.ChartGlobals.bHideScrollbars, "Charting", g.strIniFile)
        ' if gradient color = the current default, then store as -1 instead
        If g.ChartGlobals.nChartGradientColor = GradientDefault Then
            Call SetIniFileProperty("GradientColor", -1, "Charting", g.strIniFile)
        Else
            Call SetIniFileProperty("GradientColor", g.ChartGlobals.nChartGradientColor, "Charting", g.strIniFile)
        End If
        Call SetIniFileProperty("UseGradient", g.ChartGlobals.nUseGradient, "Charting", g.strIniFile)
        Call SetIniFileProperty("ShortColor", g.ChartGlobals.nShortColor, "Charting", g.strIniFile)
        Call SetIniFileProperty("LongColor", g.ChartGlobals.nLongColor, "Charting", g.strIniFile)
        Call SetIniFileProperty("LossColor", g.ChartGlobals.nLossColor, "Charting", g.strIniFile)
        Call SetIniFileProperty("WinColor", g.ChartGlobals.nWinColor, "Charting", g.strIniFile)
    End If

    With g.ChartGlobals
        Call SetIniFileProperty("SplitRollColor", .nSplitRollColor, "Charting", g.strIniFile)
        Call SetIniFileProperty("SplitsRolls", .bSplitsRolls, "Charting", g.strIniFile)
        
        Call SetIniFileProperty("FloatingTips", .bFloatingTips, "Charting", g.strIniFile)
        Call SetIniFileProperty("ChartTips", .bChartTips, "Charting", g.strIniFile)
        
        Call SetIniFileProperty("DefaultHorzStyle", .eDefaultHorzStyle, "Charting", g.strIniFile)
        Call SetIniFileProperty("DefaultAnnotStyle", .eDefaultAnnotStyle, "Charting", g.strIniFile)
        Call SetIniFileProperty("DefaultLabelMode", .eDefaultLabelMode, "Charting", g.strIniFile)
        Call SetIniFileProperty("ProfitLineStyle", .eProfitLineStyle, "Charting", g.strIniFile)
        Call SetIniFileProperty("FontSizePts", .nFontSize, "Charting", g.strIniFile)
        Call SetIniFileProperty("AutoChartData", .bAutoChartData, "Charting", g.strIniFile)
        Call SetIniFileProperty("ChartDataSingleBar", .bChartDataSingleBar, "Charting", g.strIniFile)
        Call SetIniFileProperty("LogDrawMode", .bLogModeDraw, "Charting", g.strIniFile)
                
        Call SetIniFileProperty("BadTickShowMode", .eBadTickShowMode, "Charting", g.strIniFile)
        Call SetIniFileProperty("BadTickColor", .nBadTickColor, "Charting", g.strIniFile)
        Call SetIniFileProperty("BadTickMarker", .eBadTickMarker, "Charting", g.strIniFile)
        Call SetIniFileProperty("BadTickMarkerDir", .eBadTickMarkerDir, "Charting", g.strIniFile)
        Call SetIniFileProperty("BadTickMarkerFill", .nBadTickMarkerFill, "Charting", g.strIniFile)
        
        Call SetIniFileProperty("ChartMode", .eChartMode, "Charting", g.strIniFile)
        Call SetIniFileProperty("ScaleMode", .eScaleMode, "Charting", g.strIniFile)
        Call SetIniFileProperty("HideAnnotations", .nHideAnnotations, "Charting", g.strIniFile)
        Call SetIniFileProperty("MagnetValue", .nMagnetValue, "Charting", g.strIniFile)
        Call SetIniFileProperty("DragModeY", .eDragModeY, "Charting", g.strIniFile)
        
        Call SetIniFileProperty("SquareBars", .nSquareBars, "Charting", g.strIniFile)
        Call SetIniFileProperty("SquareTicks", .dSquareTicks, "Charting", g.strIniFile)
        
        ' since the CPC flag file could get in place after TradeNav has run, let's only write
        ' this setting if it got explicitly changed by the client
        bDefault = Not FileExist(g.strAppPath & "\ShowCPC.flg")
        If .bChartModeAutoSize <> GetIniFileProperty("ChartModeAutoSize", bDefault, "Charting", g.strIniFile) Then
            Call SetIniFileProperty("ChartModeAutoSize", .bChartModeAutoSize, "Charting", g.strIniFile)
        End If
        
        bDefault = FileExist(g.strAppPath & "ShowCPC.flg")  'want to be false for non-elliott wave users
        If .bExtForecastBars <> GetIniFileProperty("ExtForecastBars", bDefault, "Charting", g.strIniFile) Then
            Call SetIniFileProperty("ExtForecastBars", .bExtForecastBars, "Charting", g.strIniFile)
        End If
        
        If .bSnapToDots <> GetIniFileProperty("SnapToDots", bDefault, "Charting", g.strIniFile) Then
            Call SetIniFileProperty("SnapToDots", .bSnapToDots, "Charting", g.strIniFile)
        End If
        
        .aSquareSymList.ToFile g.strAppPath + "\squarechart.dat"
    End With
    
    ExitFavoritesSave
    TSOFavoritesSave

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.SaveChartGlobals", eGDRaiseError_Raise
    
End Sub

' To get list of allowed templates/studies/pages (checks "required" item in file)
' - returns a tab-delimited string array: Name & vbTab & ReqMod & vbTab & Date.Time & vbTab & Desc
' - stores in a .LST file for faster retrieval (so only have to recheck things that have changed)
Public Function GetAllowedList(ByVal strType$, Optional ByVal bReturnOnlyAllowed As Boolean = True) As cGdArray
On Error GoTo ErrSection:

    Dim i&, iFile&, iMatch&, iHotKey&, hArray&
    Dim strFileMask$, strFile$, strListFile$, strExt$
    Dim aFiles As New cGdArray, aFile As New cGdArray
    Dim aMatches As New cGdArray, aMatch As New cGdArray
    
    Dim iNewFileIdx&            'index of where new template or page file should be inserted
    
    ' create string array to return
    aMatches.Create eGDARRAY_Strings
        
    ' get Type: T=Templates, S=Studies, P=Pages
    strType = UCase(Left(Trim(strType), 1))
    Select Case strType
    Case "T"
        strFileMask = App.Path & "\Charts\Templates\*.CHT"
        strListFile = App.Path & "\Charts\Templates\Templates.LST"
        strExt = ".CHT"
    Case "S"
        strFileMask = App.Path & "\Charts\Templates\*.STU"
        strListFile = App.Path & "\Charts\Templates\Studies.LST"
        strExt = ".STU"
    Case "P"
        strFileMask = g.ChartGlobals.strCPCRoot & "\Charts\Pages\*.GZP"
        strListFile = g.ChartGlobals.strCPCRoot & "\Charts\Pages\Pages.LST"
        strExt = ".GZP"
    Case Else
        Exit Function ' coding ERROR
    End Select
    
    ' get existing match list (may be custom-ordered)
    aMatches.FromFile strListFile
    
    ' get current list of all matching files in folder
    aFiles.GetMatchingFiles strFileMask, False, False, True
    aFiles.Sort eGdSort_IgnoreCase
    
    ' for each item in current list, see if the file has changed
    ' (look backwards so can delete items if no longer there)
    For iMatch = aMatches.Size - 1 To 0 Step -1
        ' MATCHES:  Name, ReqMod, Date/Time, Desc, Group (i.e favorite) Flag
        ' FILES:  Name.Ext, Size, Date/Time, Attribs
        aMatch.SplitFields aMatches(iMatch), vbTab
                
        strFile = aMatch(0) & strExt
        aFiles.BinarySearch strFile & vbTab, iFile, eGdSort_IgnoreCase
        aFile.SplitFields aFiles(iFile), vbTab
        If UCase(aFile(0)) = UCase(strFile) Then

            ' if date/time has changed, mark this match as needing to be re-checked (for ReqMod)
            If aMatch(2) <> aFile(2) Then
                ' TLB&JM 6/21/2011: we need to keep the "F" favorite flag at the end
                aMatches(iMatch) = FileBase(aFile(0)) & vbTab & Chr(27) & vbTab & aFile(2) & vbTab & vbTab & aMatch(4)
            End If

            'if this is a favorite then save index for adding new file(s)
            If iNewFileIdx = 0 Then
                If aMatch(4) = "F" Then
                    iNewFileIdx = iMatch + 1
                End If
            End If
            
            ' remove files which were found (so won't add later)
            aFiles.Remove iFile
        Else
            ' file in list no longer exists, so delete it from list
            aMatches.Remove iMatch
        End If
    Next
    
    ' then add any new files that weren't already in the list
    ' (just add to the list for now -- we'll check the ReqMod later -- and they are not favorites)
    For iFile = 0 To aFiles.Size - 1
        aFile.SplitFields aFiles(iFile), vbTab
        strFile = FileBase(aFile(0))
        aMatches.Add strFile & vbTab & Chr(27) & vbTab & aFile(2), iNewFileIdx  '5055, 5544
    Next
    
    ' now look for files in the list which must be rechecked (e.g. ReqMod)
    'For iMatch = 0 To aMatches.Size - 1
    For iMatch = aMatches.Size - 1 To 0 Step -1
        ' MATCHES:  Name, ReqMod, Date/Time, Desc, Group (i.e favorite) Flag
        aMatch.SplitFields aMatches(iMatch), vbTab
        If aMatch(1) = Chr(27) Then
            aMatch(1) = ""
            aMatch(3) = ""
            If aMatch(4) <> "F" Then
                aMatch(4) = "N"
            End If
            If strType <> "P" Then
                ' read first section of file to get desc and required
                strFile = FilePath(strFileMask) & aMatch(0) & strExt
                aFile.FromFile strFile, False, "END="
                For i = 0 To aFile.Size - 1
                    Select Case UCase(Left(aFile(i), 5))
                    Case "[IND;"
                        Exit For
                    Case "NAME="
                        'strName = Trim(Mid(aFile(i), 6))
                    Case "DESC="
                        aMatch(3) = Trim(Mid(aFile(i), 6))
                    Case "REQUI"
                        aMatch(1) = UCase(Parse(aFile(i), "=", 2))
                        ' convert from old codes (backwards-compatibility)
                        Select Case aMatch(1)
                        Case "COT"
                            aMatch(1) = "CT"
                        Case "B800"
                            aMatch(1) = "BAT"
                        Case "IC"
                            aMatch(1) = "INC"
                        End Select
                    End Select
                Next
            End If
            
            aMatches(iMatch) = aMatch.JoinFields(vbTab)
        End If
    Next
    
    If strType = "S" Then
        ' sort list for studies
        aMatches.Sort eGdSort_IgnoreCase
    Else
        TemplatePageValidate aMatches, strType
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'JM: 12-19-2008
    ' save ordered list to file (BEFORE removing unallowed items)
    'this is to ensure that the list contains ALL files in the directory
    'otherwise we will keep reading file of unallowed items
    aMatches.ToFile strListFile
    
    ' remove unallowed items before returning
    If bReturnOnlyAllowed Then
        For iMatch = aMatches.Size - 1 To 0 Step -1
            aMatch.SplitFields aMatches(iMatch), vbTab
            If Not HasModule(aMatch(1), True) Then
                aMatches.Remove iMatch
            End If
        Next
        
        ' and add hot-keys to chart templates
'        If strType = "T" And bAddTemplateHotKeys Then
'            iHotKey = 0 '(hot-keys go from "0" to "9")
'            For iMatch = 0 To aMatches.Size - 1
'                aMatches(iMatch) = "&" & Str(iHotKey) & ": " & aMatches(iMatch)
'                iHotKey = iHotKey + 1
'                If iHotKey > 9 Then Exit For
'            Next
'        End If
    End If
        
    ' return matched list
    'If strType <> "S" Then TemplatePageValidate aMatches, True, strListFile
    Set GetAllowedList = aMatches
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.GetAllowedList", eGDRaiseError_Raise
    
End Function

' This will set the chart to be looking at the image server queue ...
' - if image server not active, makes sure no charts are looking
' - if image server is active, sets just one chart to be looking (less network traffic)
' - picks a free chart (i.e. not currently processing an image server chart)
' - tries to stay away from the currently active chart (in case being used by user)
Public Sub SetImgSrvSearcher(Optional ByVal strTime$ = "")
On Error Resume Next

    Dim iForm&, iSearchingForm&, iActiveForm&
    Dim strStatus$, iCount&
    Dim frm As Form, frmActive As Form
    Static iPrevSearchingForm&
    
    ' if image server is active, find a designated searcher
    iSearchingForm = -1
    If frmImageServer.Active Then
        Set frmActive = ActiveChart
        If frmActive Is Nothing Then
            ' need a new chart!
            Set frm = New frmChart          'for imageserver just use non-detached chart
            Set frmActive = frm
        End If
        
        ' find a free non-active chart to be the designated searcher
        ' (start looking with previous one so will not keep switching
        '  and so will rotate around evenly when need to find a new one)
        If iPrevSearchingForm >= Forms.Count Then iPrevSearchingForm = 0
        For iForm = iPrevSearchingForm To Forms.Count - 1
            If IsFrmChartMDI(Forms(iForm)) Then
                Set frm = Forms(iForm)
                If frm Is frmActive Then
                    iActiveForm = iForm
                ElseIf frm.ImgSrvState <> eImgSrv_Busy Then
                    iSearchingForm = iForm
                    Exit For
                End If
            End If
        Next
        ' if need to keep looking, wrap around to beginning
        If iSearchingForm < 0 Then
            For iForm = 0 To iPrevSearchingForm - 1
                If IsFrmChartMDI(Forms(iForm)) Then
                    Set frm = Forms(iForm)
                    If frm Is frmActive Then
                        iActiveForm = iForm
                    ElseIf frm.ImgSrvState <> eImgSrv_Busy Then
                        iSearchingForm = iForm
                        Exit For
                    End If
                End If
            Next
        End If
        ' last gasp: try the active chart
        If iSearchingForm < 0 Then
            Set frm = frmActive
            If frm.ImgSrvState <> eImgSrv_Busy Then
                iSearchingForm = iActiveForm
            End If
        End If
    End If
    
    ' make sure only the designated searcher is searching
    iCount = 0
    For iForm = 0 To Forms.Count - 1
        If IsFrmChartMDI(Forms(iForm)) Then
            iCount = iCount + 1
            Set frm = Forms(iForm)
            If iForm = iSearchingForm Then
                frm.ImgSrvState = eImgSrv_Searching
                iPrevSearchingForm = iSearchingForm
                'strStatus = "IMAGE SERVER ACTIVE (Chart " & CStr(iCount) _
                '    & ": " & Parse(frm.Caption, ":", 1) & ")" & strTime
                strStatus = "IMAGE SERVER ACTIVE" & strTime
            ElseIf frm.ImgSrvState = eImgSrv_Searching Then
                frm.ImgSrvState = eImgSrv_Free
            End If
        End If
    Next
    
    StatusMsg strStatus, vbRed
    Set frm = Nothing

End Sub

Public Function FixSpreadSymbols(Chart As cChart, Optional ByVal strCompSymbol As String = "") As Boolean
On Error GoTo ErrSection:

    Dim idx&, strParm$
    Dim Pane As cPane, Ind As cIndicator
        
    strCompSymbol = UCase(Trim(strCompSymbol))
    With Chart
        ' if hidden price pane, link primary symbol to first parm of the first displayed spread function
        Set Pane = .Tree("PRICE PANE")
        If Not Pane Is Nothing Then
            If Pane.Display = False And .SymbolID <> 0 Then
                For idx = 1 To .Tree.Count
                    If .Tree.NodeLevel(idx) > 0 Then
                        Set Ind = .Tree(idx)
                        If Not Ind Is Nothing Then
                            If Ind.Display And Left(UCase(Ind.Name), 6) = "SPREAD" And Ind.ParmCount >= 2 Then
                                If Ind.ParmType(1) = 5 And Ind.ParmType(2) = 5 Then
                                    strParm = Chr(34) & GetSymbol(.SymbolID) & Chr(34)
                                    If Ind.Parm(1) <> strParm Then
                                        Ind.Parm(1) = strParm
                                        Ind.CodedText = ""
                                        .RedoMode = eRedo5_RecalcInd
                                    End If
                                    If Len(strCompSymbol) > 0 Then
                                        strParm = Chr(34) & strCompSymbol & Chr(34)
                                        If Ind.Parm(2) <> strParm Then
                                            Ind.Parm(2) = strParm
                                            Ind.CodedText = ""
                                            .RedoMode = eRedo5_RecalcInd
                                        End If
                                    End If
                                    FixSpreadSymbols = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            Set Ind = Nothing
            Set Pane = Nothing
        End If
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.FixSpreadSymbols", eGDRaiseError_Raise
End Function

'this sub used by frmFunctionMgrCT and frmNewChart
Public Function SpreadExprToTable(strExpr As String, bDivide As Boolean) As cGdTable
On Error GoTo ErrSection:

    Dim i&, j&
    Dim dMult#, strTemp$
    Dim aTemp As New cGdArray
    Dim aFields As New cGdArray
    Dim tbData As New cGdTable
                            
    If Len(strExpr) < 1 Then Exit Function
    
    bDivide = False
    
    'create table fields
    tbData.CreateField eGDARRAY_Strings, 0, "Operator"
    tbData.CreateField eGDARRAY_Strings, 1, "Symbol"
    tbData.CreateField eGDARRAY_Doubles, 2, "Multiplier"
    tbData.CreateField eGDARRAY_Doubles, 3, "Contracts"
    
    'development note:
    '06-2005: spread functions were not saved with chart string in curly braces
    '         e.g. 84 * Close Of "HU-067" + 42 * Close Of "HO-067" - 3 * Close Of "CL-067"
    '08-2005: spread functions changed to save with chart string in curly braces
    '         e.g. 84 * Close Of "HU-067" + 42 * Close Of "HO-067" - 3 * Close Of "CL-067" {;+,HU-067,84;+,HO-067,42;-,CL-067,3}
    '03-2006: if chart string exists then parse chart string instead of function text
    
    i = InStr(strExpr, "{")
    If i <> 0 Then
        strExpr = Right(strExpr, Len(strExpr) - i)      'extract chart string
        strExpr = Replace(strExpr, "}", "")             'remove right curly brace
        aTemp.SplitFields strExpr, ";"
        
        For i = 0 To aTemp.Size - 1
            aFields.SplitFields aTemp(i), ","
            If aFields.Size >= 3 Then
                tbData.AddRecord ""
                j = tbData.NumRecords - 1
                strTemp = Trim(aFields(0))      '4921
                If strTemp = "/" Then
                    tbData(0, j) = "divide"
                ElseIf strTemp = "-" Then
                    tbData(0, j) = "minus"
                Else
                    tbData(0, j) = "plus"
                End If
                tbData(1, j) = aFields(1)
                tbData(2, j) = aFields(2)
                If aFields.Size >= 4 Then
                    tbData(3, j) = aFields(3)
                Else
                    tbData(3, j) = 1
                End If
            End If
        Next
        
        Set SpreadExprToTable = tbData
        
        Exit Function
        
    End If
   
    strExpr = Replace(strExpr, "- ", "|- ")
    strExpr = Replace(strExpr, "+ ", "|+ ")
    strExpr = Replace(strExpr, "/ ", "|/ ")
        
    aTemp.SplitFields strExpr, "|"
    If aTemp.Size < 0 Then Exit Function
    
    For i = 0 To aTemp.Size - 1
        tbData.AddRecord ""
        strExpr = aTemp(i)
        'check for parentheses which would indicate a ratio spread (i.e. division operator)
        'remove the parentheses and let the next if statement process the number
        If Left(strExpr, 1) = "(" Then
            strExpr = Mid(strExpr, 2)
        End If
        
        If Left(strExpr, 1) = "+" Then
            tbData(0, i) = "plus"
            strExpr = Mid(strExpr, 2)       'get rid of first char
        ElseIf Left(strExpr, 1) = "-" Then
            tbData(0, i) = "minus"
            strExpr = Mid(strExpr, 2)       'get rid of first char
        ElseIf Left(strExpr, 1) = "/" Then
            tbData(0, i) = "divide"
            strExpr = Mid(strExpr, 2)       'get rid of first char
            bDivide = True
        Else
            tbData(0, i) = "plus"            'first char must be a number
        End If
                
        'get rid of 2nd char if it is a space
        '  backwards compatibility: original design had -3 * Close of SYMBOL as first line
        '  which is why must make sure not to throw away the 2nd char unless it is a space
        If Left(strExpr, 1) = " " Then strExpr = Mid(strExpr, 2)
        
        'look for next space and extract multiplier
        j = InStr(strExpr, " ")
        If bDivide Then
            dMult = Val(Mid(strExpr, j - 1, 1)) 'skip left parentheses: e.g. (6 * close of "SYMBOL")
        Else
            dMult = Val(Left(strExpr, j - 1))
        End If
        tbData(2, i) = dMult
        
        'look for quotes and extract symbol name
        j = InStr(strExpr, Chr(34))
        strExpr = Mid(strExpr, j + 1)
        j = InStr(strExpr, Chr(34))
        strExpr = Left(strExpr, j - 1)
        tbData(1, i) = strExpr
        
        If bDivide Then Exit For
    Next
    Set aTemp = Nothing
    
    'backwards compatibility: set contracts field to 1 if necessary (03-10-2006)
    For i = 0 To tbData.NumRecords - 1
        If tbData(3, i) <= 0 Then tbData(3, i) = 1
    Next
    
    Set SpreadExprToTable = tbData
    
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.SpreadExprToTable"
    
End Function

'this sub used by frmFunctionMgrCT and frmNewChart
Public Function BuildSpreadExpr(fg As VSFlexGrid) As String
On Error GoTo ErrSection:

    Dim i&, strFunc$, strChart$
    Dim strSymbol$, strMultiplier$, strOperator$, strContracts$
    Dim bSuccess As Boolean
            
    
    bSuccess = True
    strChart = "{"
    With fg
        For i = .FixedRows To .Rows - 1
            strContracts = ""        'reset (precautionary)
            strSymbol = .TextMatrix(i, 1)
            strMultiplier = .TextMatrix(i, 2)
            If .Cols >= 4 Then strContracts = .TextMatrix(i, 3)
            If ValOfText(strContracts) <= 0 Then strContracts = "1"
            If .TextMatrix(i, 0) = "minus" Then
                strOperator = "-"
            ElseIf .TextMatrix(i, 0) = "plus" Then
                strOperator = "+"
            ElseIf .TextMatrix(i, 0) = "divide" Then
                strOperator = "/"
            Else
                strOperator = ""
            End If
            If i = .FixedRows Then
                If Len(strSymbol) = 0 Or Len(strMultiplier) = 0 Then
                    bSuccess = False
                    strFunc = ""
                    Exit For
                Else
                    strOperator = ""
                    strFunc = strOperator & strMultiplier & " * " & strContracts & " * Close Of " & Chr(34) & strSymbol & Chr(34)
                    strChart = strChart & ";+," & strSymbol & "," & strMultiplier & "," & strContracts
                End If
            ElseIf Len(strSymbol) > 0 And Len(strMultiplier) > 0 And Len(strOperator) > 0 Then
                If strOperator = "/" Then
                    strFunc = "(" & strFunc & ")"
                    strFunc = strFunc & " " & strOperator & " (" & strMultiplier & " * " & strContracts & " * Close Of " & Chr(34) & strSymbol & Chr(34) & ")"
                Else
                    strFunc = strFunc & " " & strOperator & " " & strMultiplier & " * " & strContracts & " * Close Of " & Chr(34) & strSymbol & Chr(34)
                End If
                strChart = strChart & ";" & strOperator & "," & strSymbol & "," & strMultiplier & "," & strContracts
            End If
        Next
    End With
    
    If bSuccess Then
        BuildSpreadExpr = strFunc & strChart & "}"
    Else
        BuildSpreadExpr = ""
    End If
    
    Exit Function

ErrSection:
    RaiseError "mChartNav.BuildSpreadExpr"
    
End Function

'this sub used by frmFunctionMgrCT and frmNewChart
Public Function IsDollarMultiplier(ByVal strSpread$) As Boolean
On Error GoTo ErrSection:

    Dim i&, nID&, strText$, strSym$
    Dim dTickValue#, dTickMove#
    
    Dim aSpreadSymbols As New cGdArray
    Dim aFields As New cGdArray
    Dim tbMultiplier As New cGdTable
    Dim Bars As cGdBars
    
    Dim bError As Boolean
    Dim bDollarMult As Boolean

    'format input string: operator,symbol,multiplier;operator,symbol,multiplier; .. etc ..
    aSpreadSymbols.SplitFields strSpread, ";"
    
    'create table fields
    tbMultiplier.CreateField eGDARRAY_Doubles, 0, "Multiplier"
    tbMultiplier.CreateField eGDARRAY_Doubles, 1, "TickVM"      'tick value/tick move
    
    For i = 0 To aSpreadSymbols.Size - 1
        strText = aSpreadSymbols(i)
        If InStr(strText, "~") = 0 Then
            aFields.SplitFields strText, ","
            If aFields.Size > 2 Then
                Set Bars = New cGdBars
                strSym = aFields(1)
                nID = GetMarketInfo(strSym, Bars)
                If nID > 0 Then
                    dTickValue = Bars.Prop(eBARS_TickValue)
                    dTickMove = Bars.Prop(eBARS_TickMove)
                    If dTickValue > 0 And dTickMove > 0 Then
                        tbMultiplier.AddRecord ""
                        tbMultiplier(0, tbMultiplier.NumRecords - 1) = ValOfText(aFields(2))
                        tbMultiplier(1, tbMultiplier.NumRecords - 1) = dTickValue / dTickMove
                    Else
                        bError = True       'tick move or tick value is invalid
                    End If
                Else
                    bError = True           'symbol id is invalid
                End If
            End If
        End If
        If bError Then Exit For
    Next
    
    If Not bError Then
        bDollarMult = True
        'compare saved mulitpliers against tick value / tick move
        For i = 0 To tbMultiplier.NumRecords - 1
            If tbMultiplier(0, i) <> tbMultiplier(1, i) Then
                bDollarMult = False
                Exit For
            End If
        Next
    End If
    
    IsDollarMultiplier = bDollarMult
    
    Exit Function

ErrSection:
    RaiseError "mChartNav.IsDollarMultiplier"
    
End Function

'this sub used by frmFunctionMgrCT and frmNewChart
Public Function ToggleAutoMultiplier(fg As VSFlexGrid, ByVal nOnOff&) As Boolean
On Error GoTo ErrSection:

    Dim i&, j&, nID&, strSym$
    Dim bError As Boolean
    Dim SpreadBars() As cGdBars
    Dim dTickVM#
        
    ReDim SpreadBars(0) As cGdBars
    Set SpreadBars(0) = Nothing
    
    If fg.Rows < 3 Or fg.Cols < 5 Then Exit Function            'precautionary
    
    With fg
        j = 0
        If .MergeRow(.Rows - 1) Then
            ReDim SpreadBars(.Rows - 2) As cGdBars    'last row has 'click to add...'
        Else
            ReDim SpreadBars(.Rows - 1) As cGdBars
        End If
        For i = .FixedRows To .Rows - 1
            strSym = Trim(.TextMatrix(i, 1))
            If Len(strSym) > 0 And j < UBound(SpreadBars) Then
                Set SpreadBars(j) = New cGdBars
                If SpreadBars(j) Is Nothing Then
                    bError = True
                Else
                    nID = GetMarketInfo(strSym, SpreadBars(j))
                    If nID > 0 Then
                        If SpreadBars(j).Prop(eBARS_TickValue) <= 0 Then
                            MsgBox "Invalid tick value." & SpreadBars(j).Prop(eBARS_Symbol) & " (" & ValOfText(SpreadBars(j).Prop(eBARS_TickValue)) & ")"
                            bError = True
                        ElseIf SpreadBars(j).Prop(eBARS_TickMove) <= 0 Then
                            MsgBox "Invalid tick move: " & SpreadBars(j).Prop(eBARS_Symbol) & " (" & ValOfText(SpreadBars(j).Prop(eBARS_TickMove)) & ")"
                            bError = True
                        Else
                            dTickVM = SpreadBars(j).Prop(eBARS_TickValue) / SpreadBars(j).Prop(eBARS_TickMove)
                            If nOnOff = 1 Then
                                'checkbox was off then got turned on: use tick V/M as multiplier
                                .TextMatrix(i, 2) = Str(dTickVM)
                            Else
                                'checkbox was on then got turned off: restore last saved multiplier entered by user
                                .TextMatrix(i, 2) = .TextMatrix(i, 4)
                            End If
                            j = j + 1
                        End If
                    Else
                        MsgBox "GetMarketInfo returned invalid ID (" & Str(nID) & ") for " & strSym
                        bError = True
                    End If
                End If
            End If
            If bError Then Exit For
        Next
    End With
    
    If bError Then
        'restore multiplier
        With fg
            For i = .FixedRows To .Rows - 2
                .TextMatrix(i, 2) = .TextMatrix(i, 4)
            Next
        End With
    End If
    
    'clean up
    ReDim SpreadBars(0) As cGdBars
    Set SpreadBars(0) = Nothing
    
    ToggleAutoMultiplier = bError
    
    Exit Function

ErrSection:
    RaiseError "mChartNav.ToggleAutoMultiplier"
    
End Function

Public Sub ClearAllBuySellBtns(Optional ctlExceptButton As Control, _
    Optional hWnd As Long = 0)
On Error GoTo ErrSection:

    Dim i&
        
    'force state of Buy/Sell buttons UP for all charts (except for the exception button)
    frmMain.tmrCheckBuySellButtons.Tag = CStr(hWnd)
    For i = 0 To Forms.Count - 1
        If IsFrmChart(Forms(i)) Then
            If Forms(i).hWnd <> hWnd Then
                Forms(i).ClearBuySellButtons
            End If
        End If
    Next
    
    If ctlExceptButton Is Nothing Or hWnd = 0 Then
        frmMain.tmrCheckBuySellButtons.Enabled = False
    Else
        frmMain.tmrCheckBuySellButtons.Enabled = True
    End If
        
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.ClearAllBuySellBtns"
    
End Sub

' TLB: this routine appears to be designed to only return 1 day of tick data
'ByRef arguments:
'- Bars [out]           contains data if there is data
'- nEndDate [out]       date for data or date of attempted data retrieval if no data
'BvVal arguments:
'- strSym               name of symbol to retrieve data for
'- nSymID               ID of symbol to retrieve data for
'- nSessionDate         date to retrieve data for or set to zero for current session
'- nDaysFromDownload    0: requesting current session
'                       <>0: requesting current or most recent since days from last download
'                            this should presumably be a negative value coming in
Public Sub GetAvailTickData(Bars As cGdBars, nEndDate As Long, _
    ByVal strSym$, ByVal nSymID&, ByVal nSessionDate&, ByVal nDaysFromDownload&)
On Error GoTo ErrSection:

    Dim rc&, i&, nFirstDate&, nLastDate&
    
    Bars.Size = 0       'initialize return variables
    nEndDate = 0
       
    If g.nReplaySession > 0 And g.RealTime.Active Then
        ' for streaming replay
        SetBarProperties Bars, nSymID
        g.RealTime.AddTickBuffer Bars
        Bars.Prop(eBARS_Periodicity) = ePRD_EachTick
        Bars.ArrayMask = eBARS_TickByTick
        g.RealTime.SpliceBars Bars
        nEndDate = g.nReplaySession
    ElseIf nSessionDate = 0 Then
        ' TLB: if Salmon is running, we can check which session should already have been trading for sure
        ' (e.g. if we have the daily data but do not yet have the full ticks for that session yet
        ' -- in which case we should then just wait for the data rather than loading a prior day)
        nFirstDate = 0
        If g.RealTime.SalmonIsRunning Then
            nFirstDate = g.RealTime.SymbolInfo(strSym).LastTradedSession
            nLastDate = nFirstDate
        End If
        If nFirstDate <= 0 Then
            If nDaysFromDownload = 0 Then
                nFirstDate = Date        'current session only
            Else
                nFirstDate = LastDailyDownload + nDaysFromDownload
            End If
            nLastDate = Date + 1
        End If
            
        SetBarProperties Bars, nSymID
        g.RealTime.AddTickBuffer Bars
        
        For i = nLastDate To nFirstDate Step -1
            Bars.Prop(eBARS_Periodicity) = ePRD_EachTick
            Bars.ArrayMask = eBARS_TickByTick                    '5465
            
            If g.RealTime.Active Then
                g.RealTime.SpliceBars Bars, i
            Else
                rc = DM_GetBars(Bars, nSymID, ePRD_EachTick, i, i)
            End If
            nEndDate = i
            
            If Bars.Size > 0 Then Exit For
        Next
        
        ' TLB 4/7/2011: possible bug?  this should identify and fix it
        If Bars.Size > 1 Then
            If Bars.SessionDate(0) <> Bars.SessionDate(Bars.Size - 1) Then
                DebugLog "KS bug for " & strSym
                nEndDate = Bars.SessionDate(Bars.Size - 1)
                For i = Bars.Size - 2 To 0 Step -1
                    If Bars.SessionDate(i) <> nEndDate Then
                        Bars.DeleteFirstBars i + 1
                        Exit For
                    End If
                Next
            End If
        End If
        
    ElseIf Not IsWeekday(nSessionDate) Then
        nEndDate = nSessionDate         'fix for TradeProfile having multiple copies on a Monday
    ElseIf g.RealTime.SalmonIsRunning And nSessionDate > LastDailyDownload Then
        'when asking for a specific session after the LastDailyDownload
        SetBarProperties Bars, nSymID
        g.RealTime.AddTickBuffer Bars
        Bars.Prop(eBARS_Periodicity) = ePRD_EachTick
        Bars.ArrayMask = eBARS_TickByTick                    '5465
        g.RealTime.SpliceBars Bars, nSessionDate
        
        'SpliceBars returns all data from nSessionDate forward in time
        'this routine is intended/expected to return data for a single session
        If Bars.SessionDate(0) > nSessionDate Then
            Bars.Size = 0
        Else
            For i = Bars.Size - 1 To 0 Step -1
                If Bars.SessionDate(i) = nSessionDate Then
                    Bars.Size = i + 1
                    Exit For
                End If
            Next
        End If
        
        nEndDate = nSessionDate
    Else
        rc = DM_GetBars(Bars, nSymID, ePRD_EachTick, nSessionDate, nSessionDate)
        nEndDate = nSessionDate
    End If
    
    Exit Sub

ErrSection:
    RaiseError "mChartNav.GetAvailTickData"
    
End Sub

Public Sub LoadAppBkImage(Optional ByVal bReset As Boolean = False)
On Error GoTo ErrSection:

    Dim nWindowState&, nScreenPointer&, strAppBkImgFile$, nLogoSize&, i&
    Static bAlreadyDone As Boolean, nPixelsWide&, nPixelsHigh&

    ' see if monitor size is now bigger than the last time the App Bkgd was created
    If nPixelsWide = 0 Then
        nPixelsWide = GetIniFileProperty("PixelsWide", 0, "AppBitmap", g.strIniFile)
        nPixelsHigh = GetIniFileProperty("PixelsHigh", 0, "AppBitmap", g.strIniFile)
    End If
    If frmMain.Width > nPixelsWide * Screen.TwipsPerPixelX Or frmMain.Height > nPixelsHigh * Screen.TwipsPerPixelY Then
        ' if so, then need to recreate the bkgrd bitmap
        ' (go 10% bigger, and at least size of primary monitor, but not smaller width/height than prior)
        bReset = True
        i = frmMain.Width
        If i < Screen.Width Then i = Screen.Width
        i = i * 1.1 / Screen.TwipsPerPixelX
        If i > nPixelsWide Then
            nPixelsWide = i
        End If
        i = frmMain.Height
        If i < Screen.Height Then i = Screen.Height
        i = i * 1.1 / Screen.TwipsPerPixelY
        If i > nPixelsHigh Then
            nPixelsHigh = i
        End If
        SetIniFileProperty "PixelsWide", nPixelsWide, "AppBitmap", g.strIniFile
        SetIniFileProperty "PixelsHigh", nPixelsHigh, "AppBitmap", g.strIniFile
    End If
    
    If bReset Or Not bAlreadyDone Then
        bAlreadyDone = True

        If g.ChartGlobals.bSnapToDots Then
            'remove the logo unless user explicitly set it after enabling the snap to dots feature
            If 0 = ValOfText(GetIniFileProperty("LogoSizeExplicit", 0, "AppBitmap", g.strIniFile)) Then
                nLogoSize = ValOfText(GetIniFileProperty("LogoSize", 40, "AppBitmap", g.strIniFile))
                If nLogoSize <> 0 Then
                    SetIniFileProperty "LogoSize", 0, "AppBitmap", g.strIniFile
                    bReset = True
                End If
            End If
        End If

        ' get name of image file
        If ExtremeCharts >= 1 Then
            'get rid of some old files
            strAppBkImgFile = App.Path & "\" & "AppBk.bmp"
            If FileExist(strAppBkImgFile) Then KillFile (strAppBkImgFile)
            strAppBkImgFile = App.Path & "\" & "BTAppBk.bmp"
            If FileExist(strAppBkImgFile) Then KillFile (strAppBkImgFile)
            'to make sure new BMP gets generated if customer already has existing AppBk.bmp file
            'strAppBkImgFile = GetIniFileProperty("AppBkImgFile", "BTAppBk.bmp", "AppBitmap", g.strIniFile)
            strAppBkImgFile = GetIniFileProperty("AppBkImgFile", "AppBkDots.bmp", "AppBitmap", g.strIniFile)
        Else
            strAppBkImgFile = GetIniFileProperty("AppBkImgFile", "AppBk.bmp", "AppBitmap", g.strIniFile)
        End If
        If InStr(strAppBkImgFile, "\") = 0 Then
            strAppBkImgFile = App.Path & "\" & strAppBkImgFile
        End If
        
        'app background: hwnd, dot color, logo color, dot size, logo size, pix between dots
        'dot color: <=0(black), >0(whatever is passed in)
        'dot size: 1=large size else smallsize
        'logo color: -1(COLOR_3DSHADOW), <=0(COLOR_3DDKSHADOW), >0(whatever is passed in)
        'logo size: default=50, minimum=10 (anything <10 defaults to 50)
        'pix between dots: default=30
        If Right(strAppBkImgFile, 1) <> "." Then ' (this provides an option to not do the image at all)
            gdResetProfiles 440, 449
            
            If bReset Then
                KillFile strAppBkImgFile
                nScreenPointer = Screen.MousePointer
                Screen.MousePointer = vbHourglass
            End If
            
            If FileExist(strAppBkImgFile) Then
                geAppBkDotsSpec ValOfText(GetIniFileProperty("DotSize", 1, "AppBitmap", g.strIniFile)), _
                                ValOfText(GetIniFileProperty("DotColor", 0, "AppBitmap", g.strIniFile)), _
                                ValOfText(GetIniFileProperty("DotPixSpace", 40, "AppBitmap", g.strIniFile))
            
            Else
                gdStartProfile 444
                If ExtremeCharts >= 1 Then
                    nLogoSize = 0              '0=no logo, 999999=BT Shield
                Else
                    nLogoSize = ValOfText(GetIniFileProperty("LogoSize", 40, "AppBitmap", g.strIniFile))
                End If
                geLoadAppBkBitmap frmMain.hWnd, _
                                  ValOfText(GetIniFileProperty("DotColor", 0, "AppBitmap", g.strIniFile)), _
                                  ValOfText(GetIniFileProperty("LogoColor", 0, "AppBitmap", g.strIniFile)), _
                                  ValOfText(GetIniFileProperty("DotSize", 1, "AppBitmap", g.strIniFile)), _
                                  nLogoSize, _
                                  ValOfText(GetIniFileProperty("DotPixSpace", 40, "AppBitmap", g.strIniFile)), _
                                  nPixelsWide, nPixelsHigh, strAppBkImgFile
                gdStopProfile 444
            End If
            
            If FileExist(strAppBkImgFile) Then
                gdStartProfile 445
On Error Resume Next
                frmMain.Picture = LoadPicture(strAppBkImgFile)
                gdStopProfile 445
                DebugLog "LoadAppBkgd: " & gdGetProfiles(440, 449, ", ")
            Else
                frmMain.Picture = LoadPicture() '(to clear image)
            End If
            
            ' if frmConfig is up modally, the bitmap won't repaint unless
            ' we do something to force it to be invalidated
            LockWindowUpdate frmMain.hWnd
            LockWindowUpdate 0
            
            If bReset Then Screen.MousePointer = nScreenPointer
        End If
    End If

    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.LoadAppBkImage"
End Sub

Public Function IsLightColor(ByVal nColor&) As Boolean
On Error Resume Next
'if 2 components >= 192 we consider it a light color

    Dim nRed&, nBlue&, nGreen&, rc&
    
    rc = geGetRedValue(nColor)
    If rc >= 0 And rc <= 255 Then
        nRed = rc
        rc = geGetBlueValue(nColor)
        If rc >= 0 And rc <= 255 Then
            nBlue = rc
            If nRed >= 192 And nBlue >= 192 Then
                IsLightColor = True
            Else
                rc = geGetGreenValue(nGreen)
                If rc >= 0 And rc <= 255 Then
                    If (nRed >= 192 And nGreen >= 192) Or (nBlue >= 192 And nGreen >= 192) Then
                        IsLightColor = True
                    End If
                End If
            End If
        End If
    End If

    Exit Function
        
End Function

Public Function PopulateIndicatorsCbo(cboIndicator As ctlUniComboImageXP, Chart As cChart, _
    ByVal nSelIndId&, ByVal nSelPaneId&, ByVal nSelPriceField&, _
    ByVal bAvgHiLow As Boolean) As Long
On Error GoTo ErrSection:

    Dim i&, k&, nCboIdx&
    Dim Ind As cIndicator
    
    If cboIndicator Is Nothing Or Chart Is Nothing Then Exit Function
    
    k = -1
    'walk through tree looking for selected pane then add indicator names to drop box
    'if indicator is price indicator then allow selection of price field
    For i = 1 To Chart.Tree.Count
        If Chart.Tree.NodeLevel(i) > 0 Then
            Set Ind = Chart.Tree(i)
            If Not Ind Is Nothing Then
                If Ind.geIndpaneId = nSelPaneId And Ind.DataType <> eINDIC_Constant _
                    And Ind.DataType <> eINDIC_BooleanArray Then
                    If Ind.isPriceInd Then
                        If Ind.geIndId = nSelIndId Then
                            nCboIdx = k + nSelPriceField + 1
                        End If
                        cboIndicator.AddItem ("Price (close)")
                        If bAvgHiLow Then
                            cboIndicator.ItemData(cboIndicator.ListCount - 1) = Ind.geIndId
                            cboIndicator.AddItem ("Price (open)")
                            cboIndicator.ItemData(cboIndicator.ListCount - 1) = Ind.geIndId
                            cboIndicator.AddItem ("Price (high)")
                            cboIndicator.ItemData(cboIndicator.ListCount - 1) = Ind.geIndId
                            cboIndicator.AddItem ("Price (low)")
                            cboIndicator.ItemData(cboIndicator.ListCount - 1) = Ind.geIndId
                            cboIndicator.AddItem ("Price (avg high low)")
                            cboIndicator.ItemData(cboIndicator.ListCount - 1) = Ind.geIndId
                            k = k + 5
                        Else
                            cboIndicator.ItemData(cboIndicator.ListCount - 1) = Ind.geIndId
                            k = k + 1
                        End If
                    Else
                        cboIndicator.AddItem Ind.ChartLabel
                        cboIndicator.ItemData(cboIndicator.ListCount - 1) = Ind.geIndId
                        k = k + 1
                        If Ind.geIndId = nSelIndId Then nCboIdx = k
                    End If
                End If
            End If
        End If
    Next
    
    PopulateIndicatorsCbo = nCboIdx
    
    Exit Function

ErrSection:
    RaiseError "mChartNav.PopulateIndicatorsCbo"
    
End Function

Public Sub IndicatorsCboToAnnot(cboIndicator As ctlUniComboImageXP, Chart As cChart, Annot As cAnnotation)
On Error GoTo ErrSection:

    Dim strText$

    If cboIndicator Is Nothing Or Chart Is Nothing Or Annot Is Nothing Then Exit Sub
    
    strText = cboIndicator.Text
    
    With Annot
        .geIndId = cboIndicator.ItemData(cboIndicator.ListIndex)  'CboItem(cboIndicator)
        .Prop("IndicatorKey") = Chart.Tree.Key(.geIndId)
        If InStr(strText, "(close)") Then
            .Prop("PriceField") = 0
        ElseIf InStr(strText, "(open)") Then
            .Prop("PriceField") = 1
        ElseIf InStr(strText, "(high)") Then
            .Prop("PriceField") = 2
        ElseIf InStr(strText, "(low)") Then
            .Prop("PriceField") = 3
        ElseIf InStr(strText, "(avg high low)") Then
            .Prop("PriceField") = 4
        End If
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.IndicatorsCboToAnnot"
    
End Sub

Public Sub MultiChartAlert(Chart As cChart, Annot As cAnnotation)
On Error GoTo ErrSection:

    Dim Alert As cAlert
    Dim NewAnnot As cAnnotation
    Dim i&, ch$, strKey$, strMsg$
    
    If Annot Is Nothing Or Chart Is Nothing Then Exit Sub
    If Not HasLevelForAlert(eGDAlertType_Annot, True) Then Exit Sub     '4610

    strMsg = "This drawing tool is set to show on multiple"
    strMsg = strMsg & vbCrLf & "charts. To add an alert, you can make"
    strMsg = strMsg & vbCrLf & "a copy or remove it from all other charts."
    
    ch = InfBox(strMsg, "?", "Make Copy|Remove|+Cancel", "Create Alert")
    
    If ch <> "C" Then
        Set NewAnnot = Annot.MakeCopy
        If Not NewAnnot Is Nothing Then
            NewAnnot.MultiChartFlag = False
            ' add to annots and save key
            NewAnnot.geAddAnnotation Chart, Annot.gePaneId, Chart.Annots.Count + 1
            i = Chart.Annots.Add(NewAnnot)
            strKey = Chart.Annots.Key(i)
            NewAnnot.Prop("AnnotKey") = strKey
            NewAnnot.geAnnId = i
            
            Set Alert = NewAnnot.AlertObject()
            If Alert Is Nothing Then Set Alert = NewAnnot.AlertObject(True)
            frmAlerts.ShowMe Alert, eGDAlertType_Annot
            
            If ch = "R" Then
                Annot.geRemoveAnnotation Chart.geChartObj
                Chart.Annots.Remove Annot.Prop("AnnotKey")
                Chart.SyncGlobalAnnots Annot, True
            End If
            
            Chart.TemplateSave
            Chart.GenerateChart eRedo2_ReloadAnnots
        End If
    End If

ErrExit:
    Set Alert = Nothing
    Set NewAnnot = Nothing
    Exit Sub

ErrSection:
    RaiseError "mChartNav.MultiChartAlert"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasLevelForAlert
'' Description: Check if user has right TradeNavigator version for type of alert
''              based on Product Features spreadsheet from Harry dated 08-06-2008
''              This is meant to be used when want to know, but do not want to show alert form.
'' Inputs:      Alert Type, Show upgrade message or not
'' Returns:     True if right version else false
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasLevelForAlert(ByVal eAlertType As eGDAlertType, ByVal bShowUpgradeMsg As Boolean) As Boolean
On Error GoTo ErrSection:
    
    Dim bOkay As Boolean
    
    bOkay = True
    
    If eAlertType = eGDAlertType_QuoteBoard Then
        ' TLB 8/28/2009: #5283 per Pete, allow BetterTrades to use QB alerts
        If ExtremeCharts >= 1 Then
            bOkay = True
        Else
            bOkay = HasLevel(eTN3_Standard, bShowUpgradeMsg)
        End If
    ElseIf eAlertType = eGDAlertType_Chart Or eAlertType = eGDAlertType_Annot Then
        bOkay = HasLevel(eTN3_Standard, bShowUpgradeMsg)
    ElseIf eAlertType = eGDAlertType_TradeSense Then
        bOkay = HasGold(True, "This type of alert", True)
    End If
        
    HasLevelForAlert = bOkay
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.HasLevelForAlert"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TemplatePageValidate (11-10-2008)
'' Description: This routine verifies the Favorites flag & Required Mod for templates
'' Inputs:      Array of templates or pages information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TemplatePageValidate(aItems As cGdArray, ByVal strType$)
On Error GoTo ErrSection:

    Dim aTemp As New cGdArray
    Dim i&, iPrevAllowed&
    Dim bFix As Boolean, bAllowed As Boolean
    
    If aItems Is Nothing Then Exit Sub
    If strType = "P" Then
        If FileExist("showcpc.flg") Then Exit Sub       '6872
    ElseIf strType <> "T" Then
        Exit Sub
    End If

    iPrevAllowed = -1
    'verify group field
    For i = 0 To aItems.Size - 1
        bFix = False
        aTemp.SplitFields aItems(i), vbTab
        If aTemp.Size < 5 Then
            bFix = True
            aTemp(3) = ""
            aTemp(4) = ""
        ElseIf aTemp.Size > 5 Then
            bFix = True
            aTemp.Size = 5
        End If
        
        If strType = "P" Then
            bAllowed = True
        Else
            bAllowed = HasModule(aTemp(1), True)
        End If
        
'        If strType = "T" And UCase(Parse(aItems(i), vbTab, 1)) = "DEFAULT" Then
'            If aTemp(4) <> "F" Then
'                aTemp(4) = "F"
'                bFix = True
'            End If
        If Len(aTemp(4)) = 0 Then
            bFix = True
            If bAllowed Then
                If i = aItems.Size - 1 Then
                    aTemp(4) = "N"
                ElseIf Parse(aItems(i - 1), vbTab, 5) = "F" And Parse(aItems(i + 1), vbTab, 5) = "F" Then
                    aTemp(4) = "F"
                Else
                    aTemp(4) = "N"
                End If
            Else
                aTemp(4) = "N"
            End If
        ElseIf Parse(aItems(i), vbTab, 5) = "F" Then
            If iPrevAllowed >= 0 Then
                If Parse(aItems(iPrevAllowed), vbTab, 5) <> "F" Then
                    bFix = True
                    aTemp(4) = "N"
                End If
            End If
        End If
        
        If bAllowed Then iPrevAllowed = i
        
        If bFix Then
            aItems(i) = aTemp.JoinFields(vbTab)
        End If
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.TemplatePageValidate"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveTempPageList (12-19-2008)
'' Description: This routine is intended to be called from frmTemplates or frmTemplatePage.
''              TemplatePageValidate should have already been done prior to calling this.
'' Inputs:      Array of templates or pages information to save
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveTempPageList(aListSave As cGdArray, ByVal eMode As eTemplateFormMode)
On Error GoTo ErrSection:

    Dim strFile$, iSize&, i&, j&
    
    Dim aMatches As cGdArray, aTempM As cGdArray, aTemp As cGdArray
    Dim hMatches&, hTempM&, hTemp&

    If aListSave Is Nothing Then Exit Sub
    If aListSave.Size <= 0 Then Exit Sub

    If eMode = eMode_Pages Then
        'JM: 12-19-2008 - pages do not currently have REQ MOD
        strFile = g.ChartGlobals.strCPCRoot & "\Charts\Pages\Pages.LST"
        aListSave.ToFile strFile
    ElseIf eMode = eMode_Templates Then
        strFile = App.Path & "\Charts\Templates\Templates.LST"
        
        If Len(strFile) > 0 Then Set aMatches = GetAllowedList("T", False)
        
        If aMatches Is Nothing Then
            aListSave.ToFile strFile
        ElseIf aMatches.Size <= 0 Then
            aListSave.ToFile strFile
        Else
            Set aTempM = New cGdArray
            Set aTemp = New cGdArray
            
            aTempM.Create eGDARRAY_Strings
            aTemp.Create eGDARRAY_Strings
            
            hMatches = aMatches.ArrayHandle()
            hTempM = aTempM.ArrayHandle()
            hTemp = aTemp.ArrayHandle()
            
            'the favorite field (F/N) may have changed so we need to search template names only
            'hTemp: list of template names to be save (contains only allowed templates)
            'hTempM: list of existing template names (contains all templates on disk)
            For i = 0 To aListSave.Size - 1
                gdInsertStr hTemp, Parse(aListSave(i), vbTab, 1), i
            Next
            
            For i = 0 To aMatches.Size - 1
                gdInsertStr hTempM, Parse(aMatches(i), vbTab, 1), i
            Next
            
            aTemp.Sort
            iSize = gdGetSize(hTempM)
            'remove items from existing list that are also in the new list to save
            For i = iSize - 1 To 0 Step -1
                If aTemp.BinarySearch(gdGetStr(hTempM, i), j, eGdSort_Default, 0, iSize) Then
                    gdDeleteItems hMatches, i, 1
                End If
            Next
            
            gdSetSize hTemp, 0, 0
            gdAppendFrom hTemp, aListSave.ArrayHandle, 0, aListSave.Size
            gdAppendFrom hTemp, aMatches.ArrayHandle, 0, aMatches.Size
            
            aTemp.ToFile strFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.SaveTempPageList"

End Sub

Public Function IsFrmChart(frm As Form) As Boolean
On Error GoTo ErrSection:

    If Not frm Is Nothing Then
        If TypeOf frm Is frmChart Or TypeOf frm Is frmChart2 Then
            IsFrmChart = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.IsFrmChart"

End Function

Public Function IsFrmChartMDI(frm As Form) As Boolean
On Error GoTo ErrSection:

    If Not frm Is Nothing Then
        If TypeOf frm Is frmChart Then IsFrmChartMDI = True
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.IsFrmChartMDI"

End Function

Public Sub InitPercentCompGrid(fg As VSFlexGrid, aSymbols As cGdArray, ByVal nHeight&, _
    Optional ByVal strSymbolIn$ = "")
On Error GoTo ErrSection:

    Dim i&, iTemp&, iColor&, strText$, strUpper$
    Dim strSector$, strSubsector$, strSymbol$
    Dim str1$, Str2$, str3$, Str4$

    If fg Is Nothing Then Exit Sub

    With fg
        .Redraw = flexRDNone
        SetupGrid fg, eGridMode_Grid
        .HighLight = flexHighlightNever
        .ExplorerBar = flexExMove
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .ExtendLastCol = False
        .ScrollBars = flexScrollBarVertical
        .FixedRows = 1
        .Rows = 4
        .Cols = 3
        'column headers
        .TextMatrix(0, 0) = "Add"
        .TextMatrix(0, 1) = "Symbol"
        .TextMatrix(0, 2) = "Color"
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .Redraw = flexRDBuffered
        
        .ColWidth(0) = 600          'Check box
        .ColWidth(2) = 900          'color
        
        .ColWidth(1) = .Width - .ColWidth(0) - .ColWidth(2)
        
        .Height = nHeight
        
        .TextMatrix(1, 1) = "Dow 30"
        .TextMatrix(2, 1) = "SP 500"
        .TextMatrix(3, 1) = "Nasdaq"
        
        .Cell(flexcpBackColor, 1, 2) = vbBlue
        .Cell(flexcpBackColor, 2, 2) = RGB(0, 128, 0)
        .Cell(flexcpBackColor, 3, 2) = vbMagenta
        
        If Len(strSymbolIn) > 0 Then
            If SecurityType(strSymbolIn) = "S" Then
                .Rows = .Rows + 2
                .TextMatrix(4, 1) = "Sector"
                .TextMatrix(5, 1) = "Subsector"
                
                .Cell(flexcpBackColor, 4, 2) = vbCyan
                .Cell(flexcpBackColor, 5, 2) = vbYellow
            End If
        End If
        
        For i = 1 To .Rows - 1
            .Cell(flexcpChecked, i, 0) = flexUnchecked
            .Cell(flexcpPictureAlignment, i, 0) = flexAlignCenterCenter
        Next
        
        If Not aSymbols Is Nothing Then
            For i = 0 To aSymbols.Size - 1
                strText = aSymbols(i)
                strSymbol = Parse(strText, "|", 1)
                iColor = ValOfText(Parse(strText, "|", 2))
                iTemp = ValOfText(Parse(strText, "|", 3))
                
                strUpper = UCase(strSymbol)
                
                If iColor = 0 Then iColor = -1
                Select Case strUpper
                    Case "$DJIA"
                        .Cell(flexcpChecked, 1, 0) = iTemp
                        .Cell(flexcpBackColor, 1, 2) = iColor
                    Case "$SPX"
                        .Cell(flexcpChecked, 2, 0) = iTemp
                        .Cell(flexcpBackColor, 2, 2) = iColor
                    Case "$COMPQ"
                        .Cell(flexcpChecked, 3, 0) = iTemp
                        .Cell(flexcpBackColor, 3, 2) = iColor
                    Case strSector
                        .Cell(flexcpChecked, 4, 0) = iTemp
                        .Cell(flexcpBackColor, 4, 2) = iColor
                    Case strSubsector
                        .Cell(flexcpChecked, 5, 0) = iTemp
                        .Cell(flexcpBackColor, 5, 2) = iColor
                    Case Else
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 1) = strSymbol
                        .Cell(flexcpChecked, .Rows - 1, 0) = iTemp
                        .Cell(flexcpPictureAlignment, .Rows - 1, 0) = flexAlignCenterCenter
                        .Cell(flexcpBackColor, .Rows - 1, 2) = iColor
                End Select
            Next
        End If
        
        .Rows = .Rows + 1
        i = .Rows - 1
        .Cell(flexcpChecked, i, 0) = flexNoCheckbox
        .TextMatrix(i, 1) = "Click to add..."
    
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.InitPercentCompGrid"
    
End Sub

Public Sub ParsePercentCompGrid(fg As VSFlexGrid, aSymbols As cGdArray, ByVal strSymbol$)
On Error GoTo ErrSection:
    
    Dim i&, strSector$, strSubsector$, strText$
    Dim str1$, Str2$, str3$, Str4$
    
    If fg Is Nothing Or aSymbols Is Nothing Or Len(strSymbol) = 0 Then Exit Sub
    
    aSymbols.Size = 0
    
    With fg
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                strText = .TextMatrix(i, 1)
                
                If strText = "Dow 30" Then
                    strText = "$DJIA"
                ElseIf strText = "SP 500" Then
                    strText = "$SPX"
                ElseIf strText = "Nasdaq" Then
                    strText = "$COMPQ"
                ElseIf InStr(UCase(strText), "SECTOR") <> 0 Then
                    GetSectorInfoForSymbol strSymbol, strSector, strSubsector, str1, Str2, str3, Str4
                    If strText = "Sector" Then
                        strText = strSector
                    ElseIf strText = "Subsector" Then
                        strText = strSubsector
                    End If
                End If
                If Len(strText) > 0 Then
                    If g.SymbolPool.PoolRecForSymbol(strText, True) > 0 Then
                        str1 = Str(.Cell(flexcpBackColor, i, 2))
                        aSymbols.Add strText & "|" & str1
                    End If
                End If
            End If
        Next
    End With
    
    aSymbols.Sort eGdSort_Default Or eGdSort_DeleteDuplicates

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.ParsePercentCompGrid"
    
End Sub

Public Sub GetSectorInfoForSymbol(strSymbol$, strSector$, strSub$, strCmp$, strSecDesc$, strSubDesc$, strCmpDesc$)
On Error GoTo ErrSection:

    Dim nSymbolID&, nGroupSymbolID&
    
    If Len(strSymbol) = 0 Then Exit Sub
    nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
    If nSymbolID <= 0 Then Exit Sub
    
    If Left(strSymbol, 3) = "$--" Then
        strSector = strSymbol
    ElseIf Left(strSymbol, 2) = "$-" Then
        strSub = strSymbol
        nGroupSymbolID = SU_GetGroupParent(nSymbolID)
        If nGroupSymbolID > 0 And nGroupSymbolID < 999999 Then
            strSector = GetSymbol(nGroupSymbolID)
        End If
    Else
        nGroupSymbolID = SU_GetGroupParent(nSymbolID)
        If nGroupSymbolID > 0 And nGroupSymbolID < 999999 Then
            strCmp = strSymbol
            strSub = GetSymbol(nGroupSymbolID)
            nGroupSymbolID = SU_GetGroupParent(nGroupSymbolID)
            If nGroupSymbolID > 0 And nGroupSymbolID < 999999 Then
                strSector = GetSymbol(nGroupSymbolID)
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
    RaiseError "mChartNav.GetSectorInfoForSymbol"

End Sub

Public Function UseDiNapFib() As Boolean
On Error GoTo ErrSection:

    Static bAlreadyChecked As Boolean
    Static bOverride As Boolean
    
    If Not bAlreadyChecked Then
        bOverride = FileExist(App.Path & "\DinapOverride.flg")
        bAlreadyChecked = True
    End If
    
    UseDiNapFib = HasModule("FIB") And Not bOverride

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartNav.UseDiNapFib"

End Function

'if an array is passed in, a list of indicators is returned as specified by nListFlag
'
'nClearFlag: 0 = don't change display type of any indicators
'            1 = change all indicators of display type ribbon to line EXCEPT the passed in indicator
'            2 = change ALL indicators of display type ribbon to line
'
'nListFlag:  0 = return indicators that can be displayed as riboon type, not including the passed in indicator
'            1 = return indicators that are of display type ribbon or the companion indicator if passed in indicator is display type ribbon
Public Sub RibbonList(aReturnList As cGdArray, Chart As cChart, InputInd As cIndicator, _
    ByVal nClearFlag&, ByVal nListFlag&)
On Error GoTo ErrSection:

    Dim nID&, i&
    
    Dim Ind As cIndicator
    Dim Tree As cGdTree

    If InputInd Is Nothing Then Exit Sub
    If Not Chart Is Nothing Then Set Tree = Chart.Tree
    If Tree Is Nothing Then Exit Sub
    
    If Not aReturnList Is Nothing Then aReturnList.Size = 0
    
    If Len(InputInd.MyKey) > 0 Then
        nID = Tree.RelativeIndex(InputInd.MyKey, eTREE_Root)
    End If
    
    If nID = 0 Then
        For i = 1 To Tree.Count
            If Tree(i) Is InputInd Then
                nID = Tree.RelativeIndex(i, eTREE_Root)
                Exit For
            End If
        Next
    End If
                    
    For i = 1 To Tree.Count
        If Tree.NodeLevel(i) >= 1 Then
            Set Ind = Tree(i)
            
            If Not Ind Is Nothing Then
                If Tree.RelativeIndex(i, eTREE_Root) = nID Then
                    'change ribbon display type to line type per clear flag
                    If Ind.DisplayType = eINDIC_Ribbon Then
                        Select Case nClearFlag
                            Case 1:
                                If Not Ind Is InputInd Then Ind.DisplayType = eINDIC_Line
                            Case 2:
                                Ind.DisplayType = eINDIC_Line
                        End Select
                    End If
                    
                    'make list of indicators per list flag if an array is passed in
                    If Not aReturnList Is Nothing Then
                        If nListFlag = 0 Then
                            'returns ribbon candidates
                            If Ind.DataType = eINDIC_Array And Not Ind Is InputInd Then
                                aReturnList.Add Ind.ChartLabel & "|" & Str(i)
                            End If
                        ElseIf Ind.DisplayType = eINDIC_Ribbon And Not Ind Is InputInd Then
                            'returns companion ribbon indicator if passed in indicator is of type ribbon else should be an array of size 2 since
                            aReturnList.Add Ind.ChartLabel & "|" & Str(i)
                        End If
                    End If
                    
                End If
            End If
        End If
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.RibbonList"

End Sub

Private Sub ExitFavoritesLoad()
On Error GoTo ErrExit

    Dim fh%
    Dim strName$, strFile$
    Dim strText$, strBaseSym$, strTemp$
    
    Dim aSymExits As cGdArray
    Dim oExit As cExitStrategy


    'get last used auto exit favorites mode from INI file
    g.ChartGlobals.eAEFMode = GetIniFileProperty("AEFMode", eAEFMode_Global, "ExitFavorites", g.strIniFile)

    'JM 03-29-2011: original design intention
    '   -this array always contains 4 object in A,B,C,D order corresponding to array index [0],[1],[2],[3]
    '   -if the saved favorite in the INI file loads successfully then the object will have valid file name property
    '   -if object[i] is demmed valid then corresponding letter is assumed to be assigned
    '   -there should be no array element that evaluates to "nothing"
    
    Set g.ChartGlobals.aExitFavorites = New cGdArray
    
    strName = GetIniFileProperty("A", "", "ExitFavorites", g.strIniFile)
    Set oExit = New cExitStrategy
    If Len(strName) > 0 Then
        'if load fails then revert to new object with default info
        If Not oExit.Load(strName) Then Set oExit = New cExitStrategy
    End If
    g.ChartGlobals.aExitFavorites.Add oExit

    strName = GetIniFileProperty("B", "", "ExitFavorites", g.strIniFile)
    Set oExit = New cExitStrategy
    If Len(strName) > 0 Then
        'if load fails then revert to new object with default info
        If Not oExit.Load(strName) Then Set oExit = New cExitStrategy
    End If
    g.ChartGlobals.aExitFavorites.Add oExit

    strName = GetIniFileProperty("C", "", "ExitFavorites", g.strIniFile)
    Set oExit = New cExitStrategy
    If Len(strName) > 0 Then
        'if load fails then revert to new object with default info
        If Not oExit.Load(strName) Then Set oExit = New cExitStrategy
    End If
    g.ChartGlobals.aExitFavorites.Add oExit

    strName = GetIniFileProperty("D", "", "ExitFavorites", g.strIniFile)
    Set oExit = New cExitStrategy
    If Len(strName) > 0 Then
        'if load fails then revert to new object with default info
        If Not oExit.Load(strName) Then Set oExit = New cExitStrategy
    End If
    g.ChartGlobals.aExitFavorites.Add oExit
    
    
    'JM 01-26-2012 - add autoexit by symbol to tree
    Set g.ChartGlobals.treeExitFavorites = New cGdTree      'always create tree
    
    strTemp = g.strAppPath & "\custom\ExitFavorites.Cfg"
    If Not FileExist(strTemp) Then GoTo ErrExit
    Set aSymExits = New cGdArray
    
    
    Dim oSymExits As cSymExitFavorites
    
    fh = FreeFile
    Open strTemp For Input As #fh

    Do While Not EOF(fh)
        Line Input #fh, strText
        
        Select Case Left(strText, 1)
            Case "["
                If Not oSymExits Is Nothing Then
                    If aSymExits.Size > 0 Then
                        oSymExits.ExitFavoritesSet strBaseSym, aSymExits
                        g.ChartGlobals.treeExitFavorites.Add oSymExits, oSymExits.BaseSym
                    End If
                End If
                
                strTemp = Right(strText, Len(strText) - 1)
                strBaseSym = Parse(strTemp, "]", 1)
                Set oSymExits = New cSymExitFavorites
                aSymExits.Size = 0
                
            Case "A", "B", "C", "D"
                If Len(strBaseSym) > 0 Then
                    strName = Parse(strText, "=", 2)
                    If Len(strName) > 0 Then
                        aSymExits.Add strText
                    End If
                End If
        End Select
    Loop
    
    If Not oSymExits Is Nothing Then
        If aSymExits.Size > 0 Then
            oSymExits.ExitFavoritesSet strBaseSym, aSymExits
            g.ChartGlobals.treeExitFavorites.Add oSymExits, oSymExits.BaseSym
        End If
    End If
    
ErrExit:
    If fh <> 0 Then Close #fh
    Exit Sub

ErrSection:
    If fh <> 0 Then Close #fh
    RaiseError "mChartNav.ExitFavoritesLoad"

End Sub

Private Sub ExitFavoritesSave()
On Error GoTo ErrSection:

    Dim fh%, i&, j&
    Dim strName$, strTemp$
    
    Dim aExitsInfo As New cGdArray
    Dim oExit As cExitStrategy
    Dim oSymExits As cSymExitFavorites
    
    'save last used auto exit favorites mode from INI file
    SetIniFileProperty "AEFMode", g.ChartGlobals.eAEFMode, "ExitFavorites", g.strIniFile
    
    If Not g.ChartGlobals.aExitFavorites Is Nothing Then
        strName = ""
        Set oExit = g.ChartGlobals.aExitFavorites(0)
        If Not oExit Is Nothing Then strName = oExit.FileName               'theoretically object should never be nothing
        SetIniFileProperty "A", strName, "ExitFavorites", g.strIniFile      'note: len(strName)=0 then nothing is written to file
    
        strName = ""
        Set oExit = g.ChartGlobals.aExitFavorites(1)
        If Not oExit Is Nothing Then strName = oExit.FileName
        SetIniFileProperty "B", strName, "ExitFavorites", g.strIniFile
    
        strName = ""
        Set oExit = g.ChartGlobals.aExitFavorites(2)
        If Not oExit Is Nothing Then strName = oExit.FileName
        SetIniFileProperty "C", strName, "ExitFavorites", g.strIniFile
    
        strName = ""
        Set oExit = g.ChartGlobals.aExitFavorites(3)
        If Not oExit Is Nothing Then strName = oExit.FileName
        SetIniFileProperty "D", strName, "ExitFavorites", g.strIniFile
    End If
    
    'save symbol specific favorite exits
    If Not g.ChartGlobals.treeExitFavorites Is Nothing Then
        If g.ChartGlobals.treeExitFavorites.Count > 0 Then
            strTemp = g.strAppPath & "\custom\ExitFavorites.Cfg"
            
            fh = FreeFile
            Open strTemp For Output As #fh
            
            For i = 1 To g.ChartGlobals.treeExitFavorites.Count
                Set oSymExits = g.ChartGlobals.treeExitFavorites(i)
                If Not oSymExits Is Nothing Then
                    oSymExits.ExitFavoritesGet oSymExits.BaseSym, aExitsInfo
                    If aExitsInfo.Size > 0 Then
                        Print #fh, "[" & oSymExits.BaseSym & "]"
                        For j = 0 To aExitsInfo.Size - 1
                            Print #fh, aExitsInfo(j)
                        Next
                    End If
                End If
            Next
        End If
    End If

ErrExit:
    If fh <> 0 Then Close #fh
    Exit Sub

ErrSection:
    RaiseError "mChartNav.ExitFavoritesSave"

End Sub

Public Function ExitFavoritesAssigned(ByVal strSymbol$) As String
On Error GoTo ErrSection:

    Dim i&
    Dim strFavorites$, strBaseSym$
    
    Dim oExitBySym As cSymExitFavorites
    Dim oExit As cExitStrategy

    If g.ChartGlobals.eAEFMode = eAEFMode_Symbol Then
        strBaseSym = BaseForAutoExitFavorites(strSymbol)
        For i = 1 To g.ChartGlobals.treeExitFavorites.Count
            Set oExitBySym = g.ChartGlobals.treeExitFavorites(i)
            If Not oExitBySym Is Nothing Then
                strFavorites = oExitBySym.ExitFavoritesAssigned(strBaseSym)
                If Len(strFavorites) > 0 Then Exit For
            End If
        Next
        
        If Len(strFavorites) > 0 Then GoTo ErrExit
    End If
    
    Set oExit = g.ChartGlobals.aExitFavorites(0)
    If Not oExit Is Nothing Then        'theoretically object should never be nothing
        If Len(oExit.FileName) > 0 Then strFavorites = "A"
    End If

    Set oExit = g.ChartGlobals.aExitFavorites(1)
    If Not oExit Is Nothing Then
        If Len(oExit.FileName) > 0 Then strFavorites = strFavorites & "|B"
    End If

    Set oExit = g.ChartGlobals.aExitFavorites(2)
    If Not oExit Is Nothing Then
        If Len(oExit.FileName) > 0 Then strFavorites = strFavorites & "|C"
    End If

    Set oExit = g.ChartGlobals.aExitFavorites(3)
    If Not oExit Is Nothing Then
        If Len(oExit.FileName) > 0 Then strFavorites = strFavorites & "|D"
    End If
    
ErrExit:
    ExitFavoritesAssigned = strFavorites
    Exit Function

ErrSection:
    RaiseError "mChartNav.ExitFavoritesAssigned"

End Function

Public Sub ExitFavoritesNotify(ByVal bRedraw As Boolean)
On Error GoTo ErrSection:

    Dim i&
    
    If g.bUnloading Then Exit Sub
    
    'This routine is intended to be called with TRUE when:
    '   - user remove all favorites when one or more were previously assigned (redraw will hide the buttons)
    '   - user assigns a favorite when none was previously assigned (redraw will show the buttons)
    '   - user deletes a previously assigned favorite (redraw will disable that favorite button)
    '
    'This routine is intended to be called with FALSE when:
    '   - a favorite button was reassigned to a different auto exit (just need to sync the button appearance)
    
    For i = 0 To Forms.Count - 1
        
        If g.bUnloading Then Exit For
        
        If IsFrmChart(Forms(i)) Or TypeOf Forms(i) Is frmTickDistribution Then
'            If bRedraw Then
'                Forms(i).tmr.Tag = "OrderbarReset"          'forces redraw of order bar from form's timer
'            Else
                ExitCtrlAppearance Forms(i), Nothing, "", True
'            End If
        End If
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.ExitFavoritesNotify"

End Sub

Public Function ExitFavoritesBySym(ByVal strSymbol$) As cSymExitFavorites

    Dim strBaseSym$, i&
    Dim oExitBySym As cSymExitFavorites

    If Len(strSymbol) > 0 And Not g.ChartGlobals.treeExitFavorites Is Nothing Then
        strBaseSym = BaseForAutoExitFavorites(strSymbol)
        For i = 1 To g.ChartGlobals.treeExitFavorites.Count
            Set oExitBySym = g.ChartGlobals.treeExitFavorites(i)
            If Not oExitBySym Is Nothing Then
                Set ExitFavoritesBySym = oExitBySym
                Exit For
            End If
        Next
    End If

End Function


Public Sub SeasonalChartNew(ByVal strSymbolIn As String, ByVal dDateFrom As Double, _
    ByVal nCycleNum As Long, ByVal strCycleType As String, ByVal strBarType As String)
On Error GoTo ErrSection:

    Dim strFile$, strFileLastUsed$, i&
    
    Dim frm As frmChart
    Dim Chart As cChart
    Dim Pane As cPane
    
    Dim bLocked As Boolean
    
    If Len(strSymbolIn) = 0 Then Exit Sub
    
    'use a custom default seasonal template if exist
    strFile = g.strAppPath & "\charts\templates\default seasonal.CHT"
    
    If Not FileExist(strFile) Then
        strFile = g.strAppPath & "\charts\templates\default.cht"
        If Not ActiveChart Is Nothing Then
            'use active chart's template if active chart is a seasonal chart
            Set Chart = ActiveChart.Chart
            If Not Chart Is Nothing Then
                If Chart.TypeOfChart = eTypeChart_Seasonal Then
                    strFile = g.ChartGlobals.strCPCRoot & "\charts\" & Chart.Template & ".CHT"
                End If
            End If
        End If
    End If
    Set Chart = Nothing
    
    Set frm = New frmChart
    Set Chart = frm.Chart
    
    With Chart
        If .TemplateLoad(, False, strFile) Then
        
            If .TypeOfChart <> eTypeChart_Seasonal Then i = .SeasonalReset(Nothing)
        
            .SetSymbol strSymbolIn
            .ChangeBarPeriod strBarType, False
            .ShowTrades = 0
            .ShowToolbar = 0
            .DisableRT = True
            .ShowEmptyBars = False
            .SeasonalCycle = Str(nCycleNum) & " " & strCycleType
            .FromDate = 0
            .ToEndOfData = True
            
            frm.SkipFocusFix = True
'            frm.Chart.TemplateSave      '5007

            If Not Chart.Tree Is Nothing Then
                If TypeOf Chart.Tree("PRICE PANE") Is cPane Then
                    Set Pane = Chart.Tree("PRICE PANE")
                    Pane.Scaling = ePANE_ScaleModeAuto                  '6359
                    Pane.PaneLogFlag = ePANE_LogFlagLinear
                End If
            End If
            
            'override so all seasonal charts have same template name in chart's caption     - 6363
            .TemplateApplied = "Seasonal/Cycle Analysis"
            
            bLocked = LockWindowUpdate(frmMain.hWnd)
            
            ShowForm frm
            MoveFocus frm.pbChart
        Else
            InfBox "Load template failed: " & strFile, "E", "Ok"
        End If
    End With

ErrExit:
    If bLocked Then LockWindowUpdate 0
    Set frm = Nothing
    Exit Sub

ErrSection:
    If bLocked Then LockWindowUpdate 0
    RaiseError "mChartNav.SeasonalChartNew"

End Sub

Public Sub InitSeaonalComboCtrl(cboCycle As ctlUniComboImageXP, Optional ChartIn As cChart = Nothing)
On Error GoTo ErrSection:

    Dim idx&, i&
    Dim aTypes As cGdArray
    Dim Chart As cChart

    If cboCycle Is Nothing Then Exit Sub
    
    If ChartIn Is Nothing Then
        Set Chart = New cChart
    Else
        Set Chart = ChartIn
    End If

    idx = Chart.ValidSeaonalCycles(aTypes)
    
    If aTypes Is Nothing Then Exit Sub
    If aTypes.Size <= 0 Then Exit Sub
    
    cboCycle.Clear
    
    For i = 0 To aTypes.Size - 1
        cboCycle.AddItem aTypes(i)
    Next
    
    If Not ChartIn Is Nothing And idx >= 0 And idx < cboCycle.ListCount Then
        cboCycle.ListIndex = idx
    Else
        cboCycle.ListIndex = 0
    End If
    
    Set aTypes = Nothing
    Set Chart = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.InitSeaonalComboCtrl"

End Sub

Public Sub InitSeasonalBartypeCombo(cboBarType As ctlUniComboImageXP, Optional ChartIn As cChart = Nothing)
On Error GoTo ErrSection:
    
    Dim idx&, i&
    Dim aTypes As cGdArray
    Dim Chart As cChart

    If cboBarType Is Nothing Then Exit Sub
    
    If ChartIn Is Nothing Then
        Set Chart = New cChart
    Else
        Set Chart = ChartIn
    End If

    idx = Chart.ValidSeasonalBarTypes(aTypes)
    If aTypes Is Nothing Then Exit Sub
    If aTypes.Size <= 0 Then Exit Sub
    
    cboBarType.Clear
    
    For i = 0 To aTypes.Size - 1
        cboBarType.AddItem aTypes(i)
    Next
    
    If Not ChartIn Is Nothing And idx >= 0 And idx < cboBarType.ListCount Then
        cboBarType.ListIndex = idx
    Else
        cboBarType.ListIndex = 1
    End If
    
    Set aTypes = Nothing
    Set Chart = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.InitSeasonalBartypeCombo"

End Sub

Public Sub InitChartFlex(frm As Form, ByVal nFlexType&)
On Error GoTo ErrSection:

    Dim eType As eChartFlexCtrlIndex
    Dim bLoaded As Boolean

    If Not IsFrmChart(frm) Then Exit Sub
    
    If frm.fgChartFlex Is Nothing Then Exit Sub
    
    eType = nFlexType
    bLoaded = frm.GridLoaded(eType)

    Select Case eType
        Case eFlexGridIdx_OrdWizard
            If Not bLoaded Then
                Load frm.fgChartFlex(eFlexGridIdx_OrdWizard)
                frm.GridLoaded(eFlexGridIdx_OrdWizard) = True
            
                With frm.fgChartFlex(eFlexGridIdx_OrdWizard)
                    .Redraw = flexRDNone
                    
                    .HighLight = flexHighlightNever
                    .ScrollBars = flexScrollBarNone
                    .Editable = flexEDKbdMouse
                    .ExtendLastCol = False
                    
                    .Width = frm.fraOrdWizard.Width - 15
                    .RowHeightMax = 240
                    .ColWidthMin = 15
                    .ColWidthMax = 1250
                    
                    .FixedRows = 1
                    .FixedCols = 0
                    .Rows = 5
                    .Cols = 5
                    
                    .ColAlignment(0) = flexAlignCenterCenter
                    .ColAlignment(1) = flexAlignLeftCenter
                    .ColAlignment(2) = flexAlignCenterCenter
                    
                    .TextMatrix(0, 1) = "Symbol"
                    .TextMatrix(0, 2) = "Qty"
                    .TextMatrix(0, 3) = ""
                    .TextMatrix(0, 4) = ""
                    
                    .ColHidden(3) = True
                    .ColHidden(4) = True
                    
                    .ColWidth(3) = 15
                    .ColWidth(4) = 15
                    .ColWidth(0) = 120
                    .ColWidth(2) = 330
                            
                    .Height = .Rows * .RowHeight(0) + 30
                    
                    .Redraw = flexRDBuffered
                    .Visible = False                'so won't show in chart's area before other order bar ctrls are positioned
                End With
            End If
        
        Case eFlexGridIdx_Seasonal
            If Not bLoaded Then
                Load frm.fgChartFlex(eFlexGridIdx_Seasonal)
                frm.GridLoaded(eFlexGridIdx_Seasonal) = True
                
                frm.fgChartFlex(eFlexGridIdx_Seasonal).RowHeightMax = 240
                frm.fgChartFlex(eFlexGridIdx_Seasonal).Visible = True
                frm.fgChartFlex(eFlexGridIdx_Seasonal).ZOrder
            
                With frm.fgChartFlex(eFlexGridIdx_Seasonal)
                    SetupGrid frm.fgChartFlex(eFlexGridIdx_Seasonal), eGridMode_Grid
                    
                    .HighLight = flexHighlightNever
                    .GridLines = flexGridFlatHorz
                    .ScrollBars = flexScrollBarVertical
                    .Cols = 2
                    .FixedCols = 0
                    .FixedRows = 1
                    .ExtendLastCol = True
                    .BackColorAlternate = ALT_GRID_ROW_COLOR
                    
                    .ColHidden(1) = True        'holds indicator ID index into chart's tree for quick access
                    
                    .TextMatrix(0, 0) = "Cycle"
                    
                    .RowHeightMax = 240
                    .Visible = True
                    .ZOrder
                End With
            
            End If
        
        Case eFlexGridIdx_PfpInd, eFlexGridIdx_PfpHits
            If Not bLoaded Then
                Load frm.fgChartFlex(eFlexGridIdx_PfpInd)
                Load frm.fgChartFlex(eFlexGridIdx_PfpHits)
                frm.GridLoaded(eFlexGridIdx_PfpInd) = True
                
                With frm.fgChartFlex(eFlexGridIdx_PfpInd)
                    .RowHeightMax = 240
                    .FixedRows = 0
                    .FixedCols = 0
                    .Rows = 4
                    .Cols = 1
                    .ScrollBars = flexScrollBarNone
                    .HighLight = flexHighlightNever
                    .ExtendLastCol = True
                End With
                
                InitGridPFP frm.fgChartFlex(eFlexGridIdx_PfpHits)
                frm.fgChartFlex(eFlexGridIdx_PfpHits).RowHeightMax = 240
                frm.fgChartFlex(eFlexGridIdx_PfpHits).Height = frm.fgChartFlex(eFlexGridIdx_PfpInd).Height * 3
                
                frm.fgChartFlex(eFlexGridIdx_PfpInd).Visible = True
                frm.fgChartFlex(eFlexGridIdx_PfpInd).ZOrder
                
                frm.fgChartFlex(eFlexGridIdx_PfpHits).Visible = True
                frm.fgChartFlex(eFlexGridIdx_PfpHits).ZOrder
            End If
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartNav.InitChartFlex"

End Sub

Public Sub InitSeasonalControls(frm As Form)
On Error GoTo ErrSection:
    
    Dim i&
    Dim Chart As cChart
    Dim Ind As cIndicator
    
    
    If Not IsFrmChart(frm) Then Exit Sub
    
    Set Chart = frm.Chart
    If Chart Is Nothing Then Exit Sub
    
    If Chart.TypeOfChart <> eTypeChart_Seasonal Then Exit Sub
    
    InitChartFlex frm, eFlexGridIdx_Seasonal
    InitSeaonalComboCtrl frm.cboCycle, Chart
    InitSeasonalBartypeCombo frm.cboBarType, Chart
    
    For i = 0 To 3
        With frm.cboTrendStyle(i)
            .AddItem "Default"
            .AddItem "Thin"
            .AddItem "Medium Thin"
            .AddItem "Medium"
            .AddItem "Medium Thick"
            .AddItem "Thick"
            .AddItem "Extra Thick"
            .AddItem "Dashed (Large)"
            .AddItem "Dashed (Small)"
            .AddItem "Dash Dot"
            
            .ListIndex = 0
        End With
    Next
    
    If frm.gdTrendColor.Count < 6 Then
        For i = 1 To 3
            Load frm.gdTrendColor(i)
            frm.gdTrendColor(i).Move frm.gdTrendColor(0).Left, frm.cboTrendStyle(i).Top
            frm.gdTrendColor(i).Visible = True
        Next
        Load frm.gdTrendColor(4)
        frm.gdTrendColor(4).Move frm.lblGradientFrom.Left + frm.lblGradientFrom.Width + 30, frm.lblGradientFrom.Top - 30
        frm.gdTrendColor(4).Visible = True
        
        Load frm.gdTrendColor(5)
        frm.gdTrendColor(5).Move frm.lblGradientTo.Left + frm.lblGradientTo.Width + 30, frm.lblGradientTo.Top - 30
        frm.gdTrendColor(5).Visible = True
    End If
    
    frm.cmdSeasonalApply.Enabled = False
    frm.SeasonalControlsReset
    
    FormResize frm

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.InitSeasonalControls"

End Sub

Public Sub CenterFormOnChart(frm As Form, Chart As cChart)
On Error GoTo ErrSection

    Dim iTop&, iLeft&
    
    If frm Is Nothing Then Exit Sub
    
    If Chart Is Nothing Then
        CenterTheForm frm
        Exit Sub
    End If
    
    If Chart.Form Is Nothing Then
        CenterTheForm frm
        Exit Sub
    End If

    If Chart.Form.DetachStatus = eDetached Then
        frm.Move Chart.Form.Left + (Chart.Form.Width - frm.Width) / 2, Chart.Form.Top + (Chart.Form.Height - frm.Height) / 2
    ElseIf ActiveChart.WindowState = vbNormal Then
        ' get top and left of the actual MDI client space
        iLeft = frmMain.Left + (frmMain.Width - frmMain.ScaleWidth)
        If frmMain.DockPro.DockedCount(HAlignRight) > 0 Then
            iLeft = iLeft - frmMain.DockPro.RightEdgeWidth
        End If
        iTop = frmMain.Top + (frmMain.Height - frmMain.ScaleHeight)
        If frmMain.DockPro.DockedCount(HAlignBottom) > 0 Then
            iTop = iTop - frmMain.DockPro.BottomEdgeHeight
        End If
        ' center on the chart in the MDI client space
        iLeft = iLeft + Chart.Form.Left + (Chart.Form.Width - frm.Width) / 2
        iTop = iTop + Chart.Form.Top + (Chart.Form.Height - frm.Height) / 2
        frm.Move iLeft, iTop
    Else
        CenterTheForm frm
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.CenterFormOnChart"

End Sub

Public Function HighestHighInBarRange(Bars As cGdBars, ByRef dY#, ByRef nBar&, ByVal nBarStart&, ByVal nBarEnd&) As Boolean
On Error GoTo ErrSection:

    Dim i&, j&, hArray&
    
    i = nBarStart
    j = nBarEnd
    
    If Bars Is Nothing Or i > j Or i < 0 Or j < 0 Then Exit Function
    
    If i >= Bars.Size Or j >= Bars.Size Then Exit Function
        
    hArray = Bars.ArrayHandle(eBARS_High)
    If hArray = 0 Then Exit Function
    
    dY = gdGetNum(hArray, i)
    nBar = i
    
    For i = nBar + 1 To j
        If gdGetNum(hArray, i) >= dY Then
            dY = gdGetNum(hArray, i)
            nBar = i
        End If
    Next
    
    HighestHighInBarRange = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.HighestHighInBarRange"

End Function

Public Function LowestLowInBarRange(Bars As cGdBars, ByRef dY#, ByRef nBar&, ByVal nBarStart&, ByVal nBarEnd&) As Boolean
On Error GoTo ErrSection:

    Dim i&, j&, hArray&
    Dim dLow#
    
    i = nBarStart
    j = nBarEnd
    
    If Bars Is Nothing Or i > j Or i < 0 Or j < 0 Then Exit Function
    
    If i >= Bars.Size Or j >= Bars.Size Then Exit Function
        
    hArray = Bars.ArrayHandle(eBARS_Low)
    If hArray = 0 Then Exit Function
    
    dY = gdGetNum(hArray, i)
    nBar = i
    
    For i = nBar + 1 To j
        dLow = gdGetNum(hArray, i)
        If dLow <> kNullData Then
            If gdGetNum(hArray, i) <= dY Then
                dY = gdGetNum(hArray, i)
                nBar = i
            End If
        End If
    Next
    
    LowestLowInBarRange = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.LowestLowInBarRange"

End Function

Private Sub TSOFavoritesLoad()
On Error GoTo ErrExit

#If 1 Then
    Dim lIndex As Long                  ' Index into a for loop
    Dim strTsogFilename As String       ' TradeSense order group filename
    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object

    Set g.ChartGlobals.astrTsogFavorites = New cGdArray
    g.ChartGlobals.astrTsogFavorites.Create eGDARRAY_Strings, 4
    
    For lIndex = 1 To 4
        strTsogFilename = GetIniFileProperty(Str(lIndex), "", "TSOFavorites", g.strIniFile)
        If Len(strTsogFilename) = 0 Then
            g.ChartGlobals.astrTsogFavorites(lIndex - 1) = ""
        Else
            Set tsoGrp = New cTradeSenseOrderGroup
            
            tsoGrp.FromFile strTsogFilename, (InStr(strTsogFilename, "Custom\") > 0)
            g.ChartGlobals.astrTsogFavorites(lIndex - 1) = TsogFavoriteString(tsoGrp)
        End If
    Next lIndex
#Else
    Dim strName As String
    Dim bCustom As Boolean
    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object

    Set g.ChartGlobals.aTSOGrpFavorites = New cGdArray

    'JM 06-05-2012: favorite TSO groups are implemented with same design principle as favorite auto-exits
    '   except favorite TSO groups do not have local mode (i.e. all favorite TSO groups are global)

    bCustom = False
    strName = GetIniFileProperty("1", "", "TSOFavorites", g.strIniFile)
    Set tsoGrp = New cTradeSenseOrderGroup
    If Len(strName) > 0 Then
        If InStr(strName, "Custom\") <> 0 Then bCustom = True
        
        'if load fails then revert to new object with default info
        tsoGrp.FromFile strName, bCustom
        If Len(tsoGrp.Name) <= 0 Or Len(tsoGrp.ID) <= 0 Then
            Set tsoGrp = New cTradeSenseOrderGroup
        End If
    End If
    g.ChartGlobals.aTSOGrpFavorites.Add tsoGrp

    bCustom = False
    strName = GetIniFileProperty("2", "", "TSOFavorites", g.strIniFile)
    Set tsoGrp = New cTradeSenseOrderGroup
    If Len(strName) > 0 Then
        If InStr(strName, "Custom\") <> 0 Then bCustom = True
        
        'if load fails then revert to new object with default info
        tsoGrp.FromFile strName, bCustom
        If Len(tsoGrp.Name) <= 0 Or Len(tsoGrp.ID) <= 0 Then
            Set tsoGrp = New cTradeSenseOrderGroup
        End If
    End If
    g.ChartGlobals.aTSOGrpFavorites.Add tsoGrp

    bCustom = False
    strName = GetIniFileProperty("3", "", "TSOFavorites", g.strIniFile)
    Set tsoGrp = New cTradeSenseOrderGroup
    If Len(strName) > 0 Then
        If InStr(strName, "Custom\") <> 0 Then bCustom = True
        
        'if load fails then revert to new object with default info
        tsoGrp.FromFile strName, bCustom
        If Len(tsoGrp.Name) <= 0 Or Len(tsoGrp.ID) <= 0 Then
            Set tsoGrp = New cTradeSenseOrderGroup
        End If
    End If
    g.ChartGlobals.aTSOGrpFavorites.Add tsoGrp

    bCustom = False
    strName = GetIniFileProperty("4", "", "TSOFavorites", g.strIniFile)
    Set tsoGrp = New cTradeSenseOrderGroup
    If Len(strName) > 0 Then
        If InStr(strName, "Custom\") <> 0 Then bCustom = True
        
        'if load fails then revert to new object with default info
        tsoGrp.FromFile strName, bCustom
        If Len(tsoGrp.Name) <= 0 Or Len(tsoGrp.ID) <= 0 Then
            Set tsoGrp = New cTradeSenseOrderGroup
        End If
    End If
    g.ChartGlobals.aTSOGrpFavorites.Add tsoGrp
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.TSOFavoritesLoad"

End Sub

Private Sub TSOFavoritesSave()
On Error GoTo ErrSection

#If 1 Then
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To 4
        SetIniFileProperty Str(lIndex), Parse(g.ChartGlobals.astrTsogFavorites(lIndex - 1), "|", 1), "TSOFavorites", g.strIniFile
    Next lIndex
#Else
    Dim tsoGrp As cTradeSenseOrderGroup

    If Not g.ChartGlobals.aTSOGrpFavorites Is Nothing Then
        Set tsoGrp = g.ChartGlobals.aTSOGrpFavorites(0)
        If tsoGrp Is Nothing Then
            'theoretically this should never be true, but if true then clear out name in INI file
            SetIniFileProperty "1", "", "TSOFavorites", g.strIniFile
        Else
            'note: by default a tradesense order group object has ..\.TSG for file name so check length of Name instead
            If Len(tsoGrp.Name) > 0 And Len(tsoGrp.ID) > 0 Then
                SetIniFileProperty "1", tsoGrp.FileName, "TSOFavorites", g.strIniFile
            Else
                SetIniFileProperty "1", "", "TSOFavorites", g.strIniFile        'clear out name in INI file
            End If
        End If
    
        Set tsoGrp = g.ChartGlobals.aTSOGrpFavorites(1)
        If tsoGrp Is Nothing Then
            SetIniFileProperty "2", "", "TSOFavorites", g.strIniFile
        Else
            If Len(tsoGrp.Name) > 0 And Len(tsoGrp.ID) > 0 Then
                SetIniFileProperty "2", tsoGrp.FileName, "TSOFavorites", g.strIniFile
            Else
                SetIniFileProperty "2", "", "TSOFavorites", g.strIniFile
            End If
        End If
    
        Set tsoGrp = g.ChartGlobals.aTSOGrpFavorites(2)
        If tsoGrp Is Nothing Then
            SetIniFileProperty "3", "", "TSOFavorites", g.strIniFile
        Else
            If Len(tsoGrp.Name) > 0 And Len(tsoGrp.ID) > 0 Then
                SetIniFileProperty "3", tsoGrp.FileName, "TSOFavorites", g.strIniFile
            Else
                SetIniFileProperty "3", "", "TSOFavorites", g.strIniFile
            End If
        End If
    
        Set tsoGrp = g.ChartGlobals.aTSOGrpFavorites(3)
        If tsoGrp Is Nothing Then
            SetIniFileProperty "4", "", "TSOFavorites", g.strIniFile
        Else
            If Len(tsoGrp.Name) > 0 And Len(tsoGrp.ID) > 0 Then
                SetIniFileProperty "4", tsoGrp.FileName, "TSOFavorites", g.strIniFile
            Else
                SetIniFileProperty "4", "", "TSOFavorites", g.strIniFile
            End If
        End If
    End If
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartNav.TSOFavoritesSave"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TsogFavorite
'' Description: Determine the favorite number for the given TradeSense order group
'' Inputs:      TradeSense order group
'' Returns:     Favorite index ( -1& if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TsogFavorite(ByVal tsoGroup As cTradeSenseOrderGroup) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into the for loop
    Dim strFavorite As String           ' Favorite
    
    lReturn = -1&
    If Not g.ChartGlobals.astrTsogFavorites Is Nothing Then
        For lIndex = 0 To 3
            strFavorite = g.ChartGlobals.astrTsogFavorites(lIndex)
            If Len(strFavorite) > 0 Then
                If tsoGroup.ID = Parse(strFavorite, "|", 2) Then
                    If tsoGroup.Name = Parse(strFavorite, "|", 3) Then
                        lReturn = lIndex
                        Exit For
                    End If
                End If
            End If
        Next lIndex
    End If
    
    TsogFavorite = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.TsogFavorite"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TsogFavoriteString
'' Description: Build a favorite string for the given TradeSense order group
'' Inputs:      TradeSense order group
'' Returns:     Favorite string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TsogFavoriteString(ByVal tsoGroup As cTradeSenseOrderGroup) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Not tsoGroup Is Nothing Then
        If (Len(tsoGroup.FileName) > 0) And (Len(tsoGroup.Name) > 0) And (Len(tsoGroup.ID) > 0) Then
            strReturn = tsoGroup.FileName & "|" & tsoGroup.ID & "|" & tsoGroup.Name
        End If
    End If
    
    TsogFavoriteString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.TsogFavoriteString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChartGlbClrForCtl
'' Description: gets chart global color based on control's back color & theme color
'' Inputs:      control
''              current loaded global color for INI property string
''              INI property string for chart global color (see LoadChartGlobals)
'' Returns:     modified chart global color for non-classic theme
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChartGlbClrForCtl(ByRef ctl As Control, ByVal nColor&, ByVal strType$) As Long
On Error GoTo ErrSection:
    
    Dim nReturn&
    
    nReturn = nColor
    If ctl Is Nothing Then GoTo ErrExit
    
    If ctl.BackColor = kDarkThemeColor Then
        If strType = "LongColor" And nColor = vbBlue Then
            nReturn = vbCyan
        ElseIf strType = "WinColor" And nColor = RGB(0, 128, 0) Then
            nReturn = vbGreen
        End If
    End If
    
ErrExit:
    ChartGlbClrForCtl = nReturn
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.ChartGlbClrForCtl"

End Function

Public Function IsBlueRange(ByVal nColor&)
On Error GoTo ErrSection:

    Dim r&
    
    If nColor < 0 Then Exit Function
    r = geGetRedValue(nColor)
    'red approx = green && blue > than both
    If Abs(r - geGetGreenValue(nColor)) <= 10 Then
        If geGetBlueValue(nColor) - r > 10 Then IsBlueRange = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.IsBlueRange"

End Function

Public Function IsGreenRange(ByVal nColor&, ByVal bSkipVbgreen As Boolean)
On Error GoTo ErrSection:

    Dim r&
    
    If nColor < 0 Then Exit Function
    If bSkipVbgreen And nColor = vbGreen Then Exit Function
    
    r = geGetRedValue(nColor)
    If Abs(r - geGetBlueValue(nColor)) <= 10 Then
        If geGetGreenValue(nColor) - r > 10 Then IsGreenRange = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartNav.IsGreenRange"

End Function
