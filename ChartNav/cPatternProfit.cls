VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPatternProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ePfpMethodType
    eMethod_IndDiff = 0
    eMethod_Corr
    eMethod_Shape
    eMehod_Directional
    eMethod_Signs
    eMethod_Peaks
End Enum

Private Enum ePfpIndFlags
    ePfp_CrossesZero = 1 ' should only be set for indicators that cross zero (max > 0 and min < 0) which are not in the price pane and are not comparison symbols
    ePfp_IsOverlayed = 2
    ePfp_IsBoolean = 4
    ePfp_IsConstant = 8
End Enum

Private Type mPrivate
    Bars As cGdBars
    OtherBars() As cGdBars  ' needs to be kept NOT ALIGNED to Market1 in this array of bars
    
    aIndNames As cGdArray
    
    Ind1 As cIndicator
    Ind2 As cIndicator
    Ind3 As cIndicator
    Ind4 As cIndicator
    
    strSymbol As String
    nSymbolID As Long
    
    nForecastBars As Long
    dStandardDev As Double
    
    nPercentCorrUser As Long        '%correlation explicitly set by user and saved to ini
    nPercentCorr As Long            'current %correlation of matches found (could be user-set or optimized result)
    
    'for optimizing (read in from INI file)
    nPtrnLen As Long
    nMaxBars As Long
    nMinHits As Long
    nLowestCorr As Long
    nBestCorr As Long
    
    nPatternSelLen As Long          '# of bars enclosed by annotation on chart
    nForecastColor As Long          'color to use for drawing forecast bars & rectangles outline
    nFillColor As Long              'color to use for shading rectangles
    nFillPattern As Long            '1=fill rectangles with fill color

    dPatternDateFrom As Double
    dPatternDateTo As Double
    
    dSearchDataFrom As Double
    dSearchDataTo As Double
    
    bOptimize As Boolean
    
    'arrays of hits
    pfp_hits() As Long
    pfp_corr() As Double
    pfp_strength() As Double
    
    'arrays of search data
    gDate() As Double
    gHourMin() As Long
    gArray1() As Double
    gArray2() As Double
    gArray3() As Double
    gArray4() As Double
    gVol() As Long
    gOI() As Long
    gTotVol() As Long
    gTotOI() As Long
    
    'arrays of search pattern
    pDate() As Double
    pHourMin() As Long
    pOpen() As Double
    pHigh() As Double
    pLow() As Double
    pClose() As Double
    pVol() As Long
    pOI() As Long
    pTotVol() As Long
    pTotOI() As Long
    
    'arrays of composited data
    mDate() As Double
    mHourMin() As Long
    mOpen() As Double
    mHigh() As Double
    mLow() As Double
    mClose() As Double
    mVol() As Long
    mOI() As Long
    mTotVol() As Long
    mTotOI() As Long
    
    'arrays of OHLC for display when correlating on indicators
    gOpen() As Double
    gHigh() As Double
    gLow() As Double
    gClose() As Double
    
    'core bars (IN/OUT parms for DLL)
    PatrnCore As CoreBars
    SearchCore As CoreBars
    CompositeCore As CoreBars
    
    'number of elements in arrays
    gNumLoaded As Long
    pNumLoaded As Long
    mNumLoaded As Long
    
    'misc
    match_type As Long                      'search flag for DLL
    eIndStyle As eIndicatorDisplayType      'plot as close line if no high,low or open of price selected
    eIndLineSylte As eIndicatorStyle


    '======================== NEW Method =========================
    aIndArrays As cGdArray
    aPaneIDs As cGdArray
    aIndFlags As cGdArray
    tMatches As cGdTable
    
    'these are values read in from INI file
    '   negative = value last used for method, but method is currently not used
    '   positive = value to use
    '   zero = treat as -0.5 for indicator differences & -1 for all other methods
    dWeightIndDiff As Double
    dWeightCorr As Double
    dWeightShape As Double
    dWeightDirectional As Double
    dWeightSigns As Double
    dWeightPeaks As Double
    
    nHeatmapStyle As Long           'for development testing
    nHeatmap As Long                'VB checkbox values: vbChecked, vbUnchecked & vbGrayed
    nShowSettings As Long           '-1=already shown, 0=don't show, 1=show
    bDefaults As Boolean            'true: settings are from INI file
End Type

Private m As mPrivate

Private Sub Class_Initialize()
On Error GoTo ErrSection

    Set m.aIndNames = New cGdArray
    Set m.tMatches = New cGdTable
    Set m.aIndArrays = New cGdArray
    Set m.aPaneIDs = New cGdArray
    Set m.aIndFlags = New cGdArray
    
    m.aIndArrays.Create eGDARRAY_Longs, 0, 0
    m.aPaneIDs.Create eGDARRAY_Longs, 0, 0
    m.aIndFlags.Create eGDARRAY_Longs, 0, 0
    
    LoadSettings

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPattnerProfit.ClassInitialize"

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    ClearPFP
    Set m.aIndNames = Nothing
    Set m.aIndArrays = Nothing
    Set m.aPaneIDs = Nothing
    Set m.aIndArrays = Nothing
    Set m.tMatches = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPattnerProfit.ClassTerminate"

End Sub

Private Function SettingCheck(ByVal strProp$, ByVal strFile$, _
    ByVal nPrev&, ByVal dPrev#, ByRef nSetting&, ByRef dSetting#) As Boolean
On Error GoTo ErrSection:
    
    Dim i&, d#
    Dim bUsingDefaults As Boolean
    
    If nPrev <> -1 Then
        i = GetIniFileProperty(strProp, -1, "IndAnalyst", strFile)
        If i = -1 Then
            nSetting = nPrev                'template does not have setting, use whatever was there
        ElseIf i <> nPrev Then
            nSetting = i
            If m.bDefaults Then
                bUsingDefaults = False      'previous setting was default, use setting from chart template
            Else
                m.nShowSettings = 1         'previous setting was not default, user may have changed template to one that has different setting saved to it
            End If
        Else
            nSetting = i                    'no change
        End If
    ElseIf dPrev <> -1# Then
        d = GetIniFileProperty(strProp, -1#, "IndAnalyst", strFile)
        If d = -1# Then
            dSetting = dPrev                'template does not have setting, use whatever was there
        ElseIf d <> dPrev Then
            dSetting = d
            If m.bDefaults Then
                bUsingDefaults = False      'previous setting was default, use setting from chart template
            Else
                m.nShowSettings = 1         'previous setting was not default, user may have changed template to one that has different setting saved to it
            End If
        Else
            dSetting = d
        End If
    End If
    
    SettingCheck = bUsingDefaults

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPattnerProfit.SettingCheck"

End Function

Public Sub LoadSettings(Optional Chart As cChart = Nothing)
On Error GoTo ErrSection:

    Dim i&, strFile$
    Dim bUsingDefaults As Boolean
    Dim bDontCare As Boolean
    
    Dim eStyle As eIndicatorStyle
    
    Dim nColor&, nFillColor&, nPattern&, nForecastBars&, nMaxBars&
    Dim nPercentCorr&, nLowestCorr&, nMinHits&, nLen&
    
    Dim dDiff#, dCorr#, dShape#, dDirection#, dSigns#, dPeaks#, dStdDev#

    If Not Chart Is Nothing Then strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & Chart.Template & ".CHT"

    If FileExist(strFile) Then
        'file is chart's template
        
        'previous settings
        eStyle = m.eIndLineSylte
        nColor = m.nForecastColor
        nFillColor = m.nFillColor
        nPattern = m.nFillPattern
        nForecastBars = m.nForecastBars
        nPercentCorr = m.nPercentCorrUser
        dStdDev = m.dStandardDev
        
        nLen = m.nPtrnLen           'optimization parms
        nMaxBars = m.nMaxBars
        nMinHits = m.nMinHits
        nLowestCorr = m.nLowestCorr
        
        dDiff = m.dWeightIndDiff    'matching methods
        dCorr = m.dWeightCorr
        dShape = m.dWeightShape
        dDirection = m.dWeightDirectional
        dSigns = m.dWeightSigns
        dPeaks = m.dWeightPeaks
        
        bUsingDefaults = True
        
        i = GetIniFileProperty("ForecastBars", -1, "IndAnalyst", strFile)
        If i = -1 Then
            'template does not have any existing indicator analyst settings
            If m.nShowSettings <> -1 Then m.nShowSettings = 1    'set to true only once
            GoTo ErrExit
        Else
            'if current setting is default then check to see if new settings is also default; once any
            'setting is not the default then don't care, but don't change module level flag until end of routine
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("ForecastLineStyle", strFile, eStyle, -1, m.eIndLineSylte, -1)
            Else
                bDontCare = SettingCheck("ForecastLineStyle", strFile, eStyle, -1, m.eIndLineSylte, -1)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("ForecastColor", strFile, nColor, -1, m.nForecastColor, -1)
            Else
                bDontCare = SettingCheck("ForecastColor", strFile, nColor, -1, m.nForecastColor, -1)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("FillColor", strFile, nFillColor, -1, m.nFillColor, -1)
            Else
                bDontCare = SettingCheck("FillColor", strFile, nFillColor, -1, m.nFillColor, -1)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("FillPattern", strFile, nPattern, -1, m.nFillPattern, -1)
            Else
                bDontCare = SettingCheck("FillPattern", strFile, nPattern, -1, m.nFillPattern, -1)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("ForecastBars", strFile, nForecastBars, -1, m.nForecastBars, -1)
            Else
                bDontCare = SettingCheck("ForecastBars", strFile, nForecastBars, -1, m.nForecastBars, -1)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("MinCorrelation", strFile, nPercentCorr, -1, m.nPercentCorrUser, -1)
            Else
                bDontCare = SettingCheck("MinCorrelation", strFile, nPercentCorr, -1, m.nPercentCorrUser, -1)
            End If
            m.nPercentCorr = m.nPercentCorrUser
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("StandardDev", strFile, -1, dStdDev, -1, m.dStandardDev)
            Else
                bDontCare = SettingCheck("StandardDev", strFile, -1, dStdDev, -1, m.dStandardDev)
            End If
        
            'optimization params
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("PatternLen", strFile, nLen, -1, m.nPtrnLen, -1)
            Else
                bDontCare = SettingCheck("PatternLen", strFile, nLen, -1, m.nPtrnLen, -1)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("LowestCorr", strFile, nLowestCorr, -1, m.nLowestCorr, -1)
            Else
                bDontCare = SettingCheck("LowestCorr", strFile, nLowestCorr, -1, m.nLowestCorr, -1)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("MaxBars", strFile, nMaxBars, -1, m.nMaxBars, -1)
            Else
                bDontCare = SettingCheck("MaxBars", strFile, nMaxBars, -1, m.nMaxBars, -1)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("MinHits", strFile, nMinHits, -1, m.nMinHits, -1)
            Else
                bDontCare = SettingCheck("MinHits", strFile, nMinHits, -1, m.nMinHits, -1)
            End If
            
            'matching methods
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("WeightIndDiff", strFile, -1, dDiff, -1, m.dWeightIndDiff)
            Else
                bDontCare = SettingCheck("WeightIndDiff", strFile, -1, dDiff, -1, m.dWeightIndDiff)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("WeightCorrelation", strFile, -1, dCorr, -1, m.dWeightCorr)
            Else
                bDontCare = SettingCheck("WeightCorrelation", strFile, -1, dCorr, -1, m.dWeightCorr)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("WeightShape", strFile, -1, dShape, -1, m.dWeightShape)
            Else
                bDontCare = SettingCheck("WeightShape", strFile, -1, dShape, -1, m.dWeightShape)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("WeightDirectional", strFile, -1, dDirection, -1, m.dWeightDirectional)
            Else
                bDontCare = SettingCheck("WeightDirectional", strFile, -1, dDirection, -1, m.dWeightDirectional)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("WeightSigns", strFile, -1, dSigns, -1, m.dWeightSigns)
            Else
                bDontCare = SettingCheck("WeightSigns", strFile, -1, dSigns, -1, m.dWeightSigns)
            End If
            
            If m.bDefaults And bUsingDefaults Then
                bUsingDefaults = SettingCheck("WeightPeaks", strFile, -1, dPeaks, -1, m.dWeightPeaks)
            Else
                bDontCare = SettingCheck("WeightPeaks", strFile, -1, dPeaks, -1, m.dWeightPeaks)
            End If
            
            m.bDefaults = bUsingDefaults
        End If
    Else
        'this code executes as part of class initialization
        m.eIndLineSylte = GetIniFileProperty("ForecastLineStyle", eINDIC_Thin, "IndAnalyst", g.strIniFile)
        m.nForecastColor = GetIniFileProperty("ForecastColor", 32768, "IndAnalyst", g.strIniFile)    'forest green default
        m.nFillColor = GetIniFileProperty("FillColor", 16777152, "IndAnalyst", g.strIniFile)         'light blue
        m.nFillPattern = GetIniFileProperty("FillPattern", 1, "IndAnalyst", g.strIniFile)            '1=fill rectangles with fill color
        m.nForecastBars = GetIniFileProperty("ForecastBars", 6, "IndAnalyst", g.strIniFile)
        m.dStandardDev = GetIniFileProperty("StandardDev", 1, "IndAnalyst", g.strIniFile)
        m.nPercentCorrUser = GetIniFileProperty("MinCorrelation", 90, "IndAnalyst", g.strIniFile)
        m.nPercentCorr = m.nPercentCorrUser
    
        'optimization params
        m.nPtrnLen = GetIniFileProperty("PatternLen", 2, "IndAnalyst", g.strIniFile)
        m.nMaxBars = GetIniFileProperty("MaxBars", 6, "IndAnalyst", g.strIniFile)
        m.nLowestCorr = GetIniFileProperty("LowestCorr", 80, "IndAnalyst", g.strIniFile)
        m.nMinHits = GetIniFileProperty("MinHits", 10, "IndAnalyst", g.strIniFile)
        
        'matching methods
        m.dWeightIndDiff = GetIniFileProperty("WeightIndDiff", 0.7, "IndAnalyst", g.strIniFile)
        m.dWeightCorr = GetIniFileProperty("WeightCorrelation", 1, "IndAnalyst", g.strIniFile)
        m.dWeightShape = GetIniFileProperty("WeightShape", 5000, "IndAnalyst", g.strIniFile)
        m.dWeightDirectional = GetIniFileProperty("WeightDirectional", 5000, "IndAnalyst", g.strIniFile)
        m.dWeightSigns = GetIniFileProperty("WeightSigns", 10000, "IndAnalyst", g.strIniFile)
        m.dWeightPeaks = GetIniFileProperty("WeightPeaks", 3000, "IndAnalyst", g.strIniFile)
        
        m.bDefaults = True
    End If
    
    If m.dWeightCorr = 0 And m.dWeightShape = 0 And m.dWeightDirectional = 0 And m.dWeightSigns = 0 And m.dWeightPeaks = 0 Then
        'if all methods=0 then turn on correlation
        m.dWeightCorr = 1
    End If
    
    If m.dWeightIndDiff = 0 Then m.dWeightIndDiff = -0.7
    If m.dWeightCorr = 0 Then m.dWeightCorr = -1
    If m.dWeightShape = 0 Then m.dWeightShape = -5000
    If m.dWeightDirectional = 0 Then m.dWeightDirectional = -5000
    If m.dWeightSigns = 0 Then m.dWeightSigns = -10000
    If m.dWeightPeaks = 0 Then m.dWeightPeaks = -3000
    
    If g.bPatProfitFlag Then
        m.nHeatmapStyle = GetIniFileProperty("HeatmapStyle", 200, "IndAnalyst", g.strIniFile)
    Else
        m.nHeatmapStyle = 200
    End If
    
    If m.nFillPattern <> 0 And m.nFillPattern <> 1 Then m.nFillPattern = 0
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPattnerProfit.LoadSettings"

End Sub

Private Sub SaveSettings(Optional Chart As cChart = Nothing)
On Error GoTo ErrSection:
    
    Dim strFile$

    If Not Chart Is Nothing Then
        strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & Chart.Template & ".CHT"
        If FileExist(strFile) Then
            SetIniFileProperty "ForecastLineStyle", m.eIndLineSylte, "IndAnalyst", strFile
            SetIniFileProperty "ForecastColor", m.nForecastColor, "IndAnalyst", strFile
            SetIniFileProperty "FillColor", m.nFillColor, "IndAnalyst", strFile
            SetIniFileProperty "FillPattern", m.nFillPattern, "IndAnalyst", strFile
            SetIniFileProperty "ForecastBars", m.nForecastBars, "IndAnalyst", strFile
            SetIniFileProperty "StandardDev", m.dStandardDev, "IndAnalyst", strFile
            SetIniFileProperty "MinCorrelation", m.nPercentCorrUser, "IndAnalyst", strFile
            'optimization params
            SetIniFileProperty "PatternLen", m.nPtrnLen, "IndAnalyst", strFile
            SetIniFileProperty "MaxBars", m.nMaxBars, "IndAnalyst", strFile
            SetIniFileProperty "LowestCorr", m.nLowestCorr, "IndAnalyst", strFile
            SetIniFileProperty "MinHits", m.nMinHits, "IndAnalyst", strFile
            'matching methods
            SetIniFileProperty "WeightIndDiff", m.dWeightIndDiff, "IndAnalyst", strFile
            SetIniFileProperty "WeightCorrelation", m.dWeightCorr, "IndAnalyst", strFile
            SetIniFileProperty "WeightShape", m.dWeightShape, "IndAnalyst", strFile
            SetIniFileProperty "WeightDirectional", m.dWeightDirectional, "IndAnalyst", strFile
            SetIniFileProperty "WeightSigns", m.dWeightSigns, "IndAnalyst", strFile
            SetIniFileProperty "WeightPeaks", m.dWeightPeaks, "IndAnalyst", strFile
        End If
    End If

'save to chartnavigator.INI file as last-used settings
    SetIniFileProperty "ForecastLineStyle", m.eIndLineSylte, "IndAnalyst", g.strIniFile
    SetIniFileProperty "ForecastColor", m.nForecastColor, "IndAnalyst", g.strIniFile
    SetIniFileProperty "FillColor", m.nFillColor, "IndAnalyst", g.strIniFile
    SetIniFileProperty "FillPattern", m.nFillPattern, "IndAnalyst", g.strIniFile
    SetIniFileProperty "ForecastBars", m.nForecastBars, "IndAnalyst", g.strIniFile
    SetIniFileProperty "StandardDev", m.dStandardDev, "IndAnalyst", g.strIniFile
    SetIniFileProperty "MinCorrelation", m.nPercentCorrUser, "IndAnalyst", g.strIniFile

    'optimization params
    SetIniFileProperty "PatternLen", m.nPtrnLen, "IndAnalyst", g.strIniFile
    SetIniFileProperty "MaxBars", m.nMaxBars, "IndAnalyst", g.strIniFile
    SetIniFileProperty "LowestCorr", m.nLowestCorr, "IndAnalyst", g.strIniFile
    SetIniFileProperty "MinHits", m.nMinHits, "IndAnalyst", g.strIniFile
    
    'matching methods
    SetIniFileProperty "WeightIndDiff", m.dWeightIndDiff, "IndAnalyst", g.strIniFile
    SetIniFileProperty "WeightCorrelation", m.dWeightCorr, "IndAnalyst", g.strIniFile
    SetIniFileProperty "WeightShape", m.dWeightShape, "IndAnalyst", g.strIniFile
    SetIniFileProperty "WeightDirectional", m.dWeightDirectional, "IndAnalyst", g.strIniFile
    SetIniFileProperty "WeightSigns", m.dWeightSigns, "IndAnalyst", g.strIniFile
    SetIniFileProperty "WeightPeaks", m.dWeightPeaks, "IndAnalyst", g.strIniFile

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPattnerProfit.SaveSettings"

End Sub

Public Sub ClearPFP()
On Error GoTo ErrSection:

    Set m.Bars = Nothing
    ReDim m.OtherBars(0)
    
    m.aIndNames.Size = 0
    Set m.Ind1 = Nothing
    Set m.Ind2 = Nothing
    Set m.Ind3 = Nothing
    Set m.Ind4 = Nothing
    
    m.strSymbol = ""
    m.nSymbolID = 0
    
    m.dSearchDataFrom = 0
    m.dSearchDataTo = 0
    m.dPatternDateFrom = 0
    m.dPatternDateTo = 0
    
    m.gNumLoaded = 0
    m.pNumLoaded = 0
    m.mNumLoaded = 0
    
    m.match_type = 0
    
    ReDim m.pfp_hits(0)
    ReDim m.pfp_corr(0)
    ReDim m.pfp_strength(0)
    
    'arrays of search data
    ReDim m.gDate(0)
    ReDim m.gHourMin(0)
    ReDim m.gArray1(0)
    ReDim m.gArray2(0)
    ReDim m.gArray3(0)
    ReDim m.gArray4(0)
    ReDim m.gVol(0)
    ReDim m.gOI(0)
    ReDim m.gTotVol(0)
    ReDim m.gTotOI(0)
    
    'arrays of search pattern
    ReDim m.pDate(0)
    ReDim m.pHourMin(0)
    ReDim m.pOpen(0)
    ReDim m.pHigh(0)
    ReDim m.pLow(0)
    ReDim m.pClose(0)
    ReDim m.pVol(0)
    ReDim m.pOI(0)
    ReDim m.pTotVol(0)
    ReDim m.pTotOI(0)
    
    'arrays of composited data
    ReDim m.mDate(0)
    ReDim m.mHourMin(0)
    ReDim m.mOpen(0)
    ReDim m.mHigh(0)
    ReDim m.mLow(0)
    ReDim m.mClose(0)
    ReDim m.mVol(0)
    ReDim m.mOI(0)
    ReDim m.mTotVol(0)
    ReDim m.mTotOI(0)
    
    'arrays of OHLC for display when correlating on indicators
    ReDim m.gOpen(0)
    ReDim m.gHigh(0)
    ReDim m.gLow(0)
    ReDim m.gClose(0)
    
    m.bOptimize = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.ClearPFP"

End Sub

Public Property Get ForecastColor() As Long
    ForecastColor = m.nForecastColor
End Property

Public Property Get NumForecastBars() As Long
    NumForecastBars = m.nForecastBars
End Property

Public Property Let NumForecastBars(ByVal nBars&)
    m.nForecastBars = nBars
End Property

Public Property Get FillColor() As Long
    FillColor = m.nFillColor
End Property

Public Property Get FillPattern() As Long
    FillPattern = m.nFillPattern
End Property

Public Property Get ForecastLineStyle() As Long
    ForecastLineStyle = m.eIndLineSylte
End Property

Public Property Get StandardDev() As Double
    StandardDev = m.dStandardDev
End Property

Public Property Get PercentCorr() As Long
    PercentCorr = m.nPercentCorrUser
End Property

Public Property Let PercentCorr(ByVal nPercent&)
    If nPercent > 0 And nPercent < 101 Then
        m.nPercentCorrUser = nPercent
    End If
End Property

Public Property Get PatternDateFrom() As Double
    PatternDateFrom = m.dPatternDateFrom
End Property

Public Property Let PatternDateFrom(ByVal dDate#)
    m.dPatternDateFrom = dDate
End Property

Public Property Get PatternDateTo() As Double
    PatternDateTo = m.dPatternDateTo
End Property

Public Property Let PatternDateTo(ByVal dDate#)
    m.dPatternDateTo = dDate
End Property

Public Property Get PatternSelLength() As Long
    PatternSelLength = m.nPatternSelLen
End Property

Public Property Let PatternSelLength(ByVal nLen&)
    m.nPatternSelLen = nLen
End Property

Public Property Get OptimizeMinBars() As Long
    OptimizeMinBars = m.nPtrnLen
End Property

Public Property Let OptimizeMinBars(ByVal nMin&)
    m.nPtrnLen = nMin
End Property

Public Property Get OptimizeMaxBars() As Long
    OptimizeMaxBars = m.nMaxBars
End Property

Public Property Let OptimizeMaxBars(ByVal nMax&)
    m.nMaxBars = nMax
End Property

Public Property Get OptimizeMinHits() As Long
    OptimizeMinHits = m.nMinHits
End Property

Public Property Let OptimizeMinHits(ByVal nMinHits&)
    m.nMinHits = nMinHits
End Property

Public Property Get OptimizeLowestCorr() As Long
    OptimizeLowestCorr = m.nLowestCorr
End Property

Public Property Let OptimizeLowestCorr(ByVal nLowest&)
    m.nLowestCorr = nLowest
End Property

Public Property Get IndDisplayStyle() As eIndicatorDisplayType
    IndDisplayStyle = m.eIndStyle
End Property

Private Function ValidateParams() As String

    Dim dTemp#, strErr$
    
    'check date range of search data
    If m.dSearchDataFrom <= 0 Then
        strErr = "Invalid date from: " & Str(m.dSearchDataFrom)
        GoTo ErrExit
    End If
    If m.dSearchDataTo <= 0 Then
        strErr = "Invalid date to: " & Str(m.dSearchDataTo)
        GoTo ErrExit
    End If
    
    'check date range of pattern
    If m.dPatternDateFrom <= 0 Then
        strErr = "Invalid pattern date from: " & Str(m.dPatternDateFrom)
        GoTo ErrExit
    End If
    If m.dPatternDateTo <= 0 Then
        strErr = "Invalid pattern date to: " & Str(m.dPatternDateTo)
        GoTo ErrExit
    End If
    
    If m.nPatternSelLen > 260 Then
        strErr = "Pattern cannot exceed 260 bars."
        GoTo ErrExit
    End If
    
    'check other parameters
    If m.dStandardDev > 10# Then
        strErr = "Standard deviation cannot exceed 10."
        GoTo ErrExit
    End If
    
    If Len(m.strSymbol) = 0 Then
        strErr = "Invalid symbol name."
        GoTo ErrExit
    End If
    
    If m.nSymbolID <= 0 Then
        strErr = "Invalid symbol ID."
        GoTo ErrExit
    End If

ErrExit:
    ValidateParams = strErr

End Function

Public Function FindMatches(Chart As cChart, fgIndicators As VSFlexGrid, fgHits As VSFlexGrid, _
    ByVal nPercentCorr As Long, _
    Optional ByVal bOptimize As Boolean = False, _
    Optional ByVal bValidateNoFind As Boolean = False) As String
On Error GoTo ErrSection:

    Dim i&, iOptimizeLen&
    Dim strErr$, strText$
    Dim bFound As Boolean
    Dim Annot As cAnnotation
    
    m.bOptimize = bOptimize
    m.nPercentCorr = nPercentCorr
    
    If Chart Is Nothing Then
        strErr = "Invalid chart object."
        GoTo ErrExit
    End If
    
    If fgIndicators Is Nothing Or fgHits Is Nothing Then
        strErr = "Invalid grid parameters."
        GoTo ErrExit
    End If
    
    With fgIndicators
        For i = .FixedRows To .Rows - 1
            strText = .Cell(flexcpTextDisplay, i, 0)
            If Len(strText) > 0 And UCase(strText) <> "NONE" Then bFound = True
            If bFound Then Exit For
        Next
    End With
    
    If Not bFound Then
        strErr = "Invalid indicators list."
        GoTo ErrExit
    End If
    
    i = Chart.LastGoodDataBar(False)
    m.nSymbolID = Chart.Bars.Prop(eBARS_SymbolID)
    m.strSymbol = Chart.Bars.Prop(eBARS_Symbol)
    m.dSearchDataFrom = Chart.Bars(eBARS_DateTime, 0)
    m.dSearchDataTo = Chart.Bars(eBARS_DateTime, i)
    
    Set m.Bars = New cGdBars
    m.Bars.CopyFromHandle Chart.Bars.BarsHandle
    m.Bars.Size = i + 1         'remove bars beyond last good data bar (i.e. forecast bars on chart)
    
    If m.Bars.Size = 0 Then
        strErr = "No data available."
        GoTo ErrExit
    ElseIf m.Bars.Size <> Chart.Bars.Size - Chart.ForecastBars Then
        strErr = "Data size mismatch."
        GoTo ErrExit
    ElseIf m.dSearchDataTo <> m.Bars(eBARS_DateTime, m.Bars.Size - 1) Then
        strErr = "Date discrepancy. Requested: " & DateFormat(m.dSearchDataTo, MM_DD_YYYY) & "  Actual: " & DateFormat(m.Bars(eBARS_DateTime, m.Bars.Size - 1), MM_DD_YYYY)
        GoTo ErrExit
    End If
    
    For i = 1 To Chart.Annots.Count
        Set Annot = Chart.Annots(i)
        If Not Annot Is Nothing Then
            If Annot.eUsage = eANNOT_PatternProfit Then
                m.dPatternDateFrom = Annot.dDate(1)
                m.nPatternSelLen = Annot.PatternLength(Chart)
                Exit For
            End If
        End If
    Next
    
    If m.dPatternDateFrom > 0 Then
        i = m.Bars.FindDateTime(m.dPatternDateFrom, True)
        If m.Bars(eBARS_DateTime, i) = m.dPatternDateFrom Then
            m.dPatternDateTo = m.Bars(eBARS_DateTime, i + m.nPatternSelLen - 1)
            ' TLB 11/11/2011: when optimizing, exclude data past the end of the pattern
            ' (so optimized results won't keep changing as the forecasted bars come into existence)
            If bOptimize Then
                m.Bars.Size = i + m.nPatternSelLen
            End If
        End If
    End If
    
    strErr = ValidateParams()
    If Len(strErr) > 0 Or bValidateNoFind Then GoTo ErrExit

'TLB: ??? don't think that we should still be doing this anymore (esp. not when also doing indicators?)
'    AdjustBarPrices
    
    strErr = LoadIndicators(Chart, fgIndicators)
    If Len(strErr) > 0 Then GoTo ErrExit

    strErr = LoadSearchCore()
    If m.gNumLoaded <= 0 Or Len(strErr) > 0 Then GoTo ErrExit
    
    m.eIndStyle = eINDIC_Line
    For i = 0 To m.aIndNames.Size - 1
        strText = UCase(Parse(m.aIndNames(i), ";", 1))
        Select Case strText
            Case "OPEN", "HIGH", "LOW"
                m.eIndStyle = eINDIC_OHLC
        End Select
        If m.eIndStyle = eINDIC_OHLC Then Exit For
    Next
    
'07-14-2010: code for old DLL function
'    iOptimizeLen = GetMatches(fgHits, Chart)

    iOptimizeLen = GetMatches3(Chart)
    
    If iOptimizeLen > 0 Then
        If iOptimizeLen <> m.nPatternSelLen Then
            i = m.Bars.FindDateTime(m.dPatternDateTo, True)
            If m.Bars(eBARS_DateTime, i) = m.dPatternDateTo Then
                m.dPatternDateFrom = m.Bars(eBARS_DateTime, i - iOptimizeLen + 1)
                m.nPatternSelLen = iOptimizeLen
                If Not Annot Is Nothing Then
                    If Annot.eUsage = eANNOT_PatternProfit Then
                        If Annot.dDate(1) < Annot.dDate(2) Then
                            Annot.dDate(1) = m.dPatternDateFrom
                        Else
                            Annot.dDate(2) = m.dPatternDateFrom
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    PopulateGrid fgHits, m.nPatternSelLen
    RebuildComposite fgHits

    m.bOptimize = False
    
    InfBox ""
    
    Chart.GenerateChart eRedo1_Scrolled
    
    If m.dPatternDateTo = Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False)) Then
        Chart.Form.hsb.Value = Chart.Form.hsb.Max
    Else
        Chart.Form.CenterTheDate m.dPatternDateTo
    End If
    
ErrExit:
    FindMatches = strErr
    If FormIsLoaded("frmPatternProfitOpt") Then Unload frmPatternProfitOpt
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.FindMatches"

End Function

Private Sub AdjustBarPrices()
On Error GoTo ErrSection:

    Dim i&, dLowest#, dHighest#, dAdjust#
    
    dLowest = gdMinValue(m.Bars.ArrayHandle(eBARS_Low), 0, m.Bars.Size)
    dHighest = gdMaxValue(m.Bars.ArrayHandle(eBARS_High), 0, m.Bars.Size)
    If dLowest < 0 Then
        ' TLB: if the lowest is negative, we are most likely dealing with a back-adjusted future
        ' so we should shift everything up so the lowest is about half the current highest
        ' in order to make all the percentage changes about right.
        If dHighest > 0 Then
            dAdjust = dHighest / 2 + Abs(dLowest)
        Else
            dAdjust = Abs(dLowest) * 2
        End If
        For i = 0 To m.Bars.Size - 1
            m.Bars(eBARS_Open, i) = m.Bars(eBARS_Open, i) + dAdjust
            m.Bars(eBARS_High, i) = m.Bars(eBARS_High, i) + dAdjust
            m.Bars(eBARS_Low, i) = m.Bars(eBARS_Low, i) + dAdjust
            m.Bars(eBARS_Close, i) = m.Bars(eBARS_Close, i) + dAdjust
        Next
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.AdjustBarPrices"
End Sub

Private Function LoadIndicators(Chart As cChart, fgIndicators As VSFlexGrid) As String
On Error GoTo ErrSection:

    Dim i&, j&, rc&, iBars&
    Dim nID&, nPricePaneId&, strName$
    Dim strText$
    
    Dim Ind As cIndicator
    Dim ChartInd As cIndicator
    
    If Chart Is Nothing Then Exit Function
    
    'reset indicator objects
    Set m.Ind1 = Nothing
    Set m.Ind2 = Nothing
    Set m.Ind3 = Nothing
    Set m.Ind4 = Nothing
    
    m.aIndNames.Size = 0
    
    With fgIndicators
        For i = .FixedRows To .Rows - 1
            Set Ind = Nothing
            strName = UCase(.Cell(flexcpTextDisplay, i, 0))
            
            nID = -1
            If strName = "OPEN" Or strName = "HIGH" Or strName = "LOW" Or strName = "CLOSE" Then
                Set Ind = New cIndicator
                Set ChartInd = Chart.Tree("PRICE")
                If Not ChartInd Is Nothing Then nID = ChartInd.geIndId
            ElseIf strName <> "NONE" Then
                strText = Parse(UCase(.Cell(flexcpData, i, 0)), ";", 1)     'indicator ID format: #X
                strText = Right(strText, Len(strText) - 1)                  'get rid of the # sign (part of grid's colcombo list)
                nID = ValOfText(strText)
                If Left(strName, 9) = "CLOSE OF " Then
                    'don need to do anything
                ElseIf Left(strName, 8) = "OPEN OF " Then
                    '!Important!
                    'this code relies on the mPFPCore.InitIndicatorGrid
                    'routine assigning the numeric using equivalent
                    'reverse math so need to keep code in sync
                    nID = nID / (Chart.Tree.Count + 1)
                ElseIf Left(strName, 8) = "HIGH OF " Then
                    nID = nID / (Chart.Tree.Count + 2)
                ElseIf Left(strName, 7) = "LOW OF " Then
                    nID = nID / (Chart.Tree.Count + 3)
                ElseIf Left(.Cell(flexcpData, i, 0), 1) = "#" Then
                    strText = Parse(UCase(.Cell(flexcpData, i, 0)), ";", 1)
                    strText = Right(strText, Len(strText) - 1)
                    nID = ValOfText(strText)
                End If
            
                If Chart.Tree.NodeLevel(nID) > 0 Then
                    Set Ind = New cIndicator
                    Set ChartInd = Chart.Tree(nID)
                End If
            End If
            
            If Not Ind Is Nothing And Not ChartInd Is Nothing Then
                'copy properties that are needed for pattern matching
                Ind.Name = ChartInd.Name
                Ind.geIndId = nID
                Ind.geIndpaneId = ChartInd.geIndpaneId
                Ind.Overlayed = ChartInd.Overlayed
                Ind.DataType = ChartInd.DataType
                Ind.ChartLabel = ChartInd.ChartLabel
                
                If ChartInd.isPriceInd = 1 Then
                    Set Ind.Bars = m.Bars           'm.Bars is already a copy
                ElseIf ChartInd.DataType = eINDIC_BarData Then
                    If ChartInd.Bars Is Nothing Then
                        InfBox "Internal Error: Bars data for " & ChartInd.ChartLabel & " not available.", "E", "Okay"
                    Else
                        Ind.Data.Size = 0
                        If Left(strName, 9) = "CLOSE OF " Then
                            Ind.Data.CopyFromHandle ChartInd.Bars.ArrayHandle(eBARS_Close)
                        ElseIf Left(strName, 8) = "OPEN OF " Then
                            Ind.Data.CopyFromHandle ChartInd.Bars.ArrayHandle(eBARS_Open)
                        ElseIf Left(strName, 8) = "HIGH OF " Then
                            Ind.Data.CopyFromHandle ChartInd.Bars.ArrayHandle(eBARS_High)
                        ElseIf Left(strName, 7) = "LOW OF " Then
                            Ind.Data.CopyFromHandle ChartInd.Bars.ArrayHandle(eBARS_Low)
                        End If
                        Ind.Data.Size = m.Bars.Size     'remove forecast bars "buckets" from chart's indicator array
                        Ind.DataType = eINDIC_Array
                    End If
                Else
                    Ind.Data.Size = 0
                    Ind.Data.AppendFromArray ChartInd.Data
                    Ind.Data.Size = m.Bars.Size     'remove forecast bars "buckets" from chart's indicator array
                End If

                If m.Ind1 Is Nothing Then
                    Set m.Ind1 = Ind
                ElseIf m.Ind2 Is Nothing Then
                    Set m.Ind2 = Ind
                ElseIf m.Ind3 Is Nothing Then
                    Set m.Ind3 = Ind
                ElseIf m.Ind4 Is Nothing Then
                    Set m.Ind4 = Ind
                End If
                
                m.aIndNames.Add .Cell(flexcpTextDisplay, i, 0) & ";" & Str(nID)
            End If
        Next
    End With

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.LoadIndicators"

End Function

Private Function LoadPatternCore(ByVal nPatternLen&) As String
On Error GoTo ErrSection:

    Dim idxFirst&, idxLast&, iSize&, i&, j&
    
    Dim hArray1&, hArray2&, hArray3&, hArray4&
               
    If m.Bars Is Nothing Then Exit Function
               
    i = nPatternLen
    If i <= 0 Then Exit Function
    
    idxLast = m.Bars.FindDateTime(m.dPatternDateTo)
    idxFirst = m.Bars.FindDateTime(m.Bars(eBARS_DateTime, idxLast - i + 1))
    
    i = i + 2
    
    ReDim m.pDate(i) As Double
    ReDim m.pHourMin(i) As Long
    ReDim m.pOpen(i) As Double
    ReDim m.pHigh(i) As Double
    ReDim m.pLow(i) As Double
    ReDim m.pClose(i) As Double
    ReDim m.pVol(i) As Long
    ReDim m.pOI(i) As Long
    ReDim m.pTotVol(i) As Long
    ReDim m.pTotOI(i) As Long
        
    SetArrayHandles hArray1, hArray2, hArray3, hArray4
    
    j = 1
    For i = idxFirst To idxLast
        'm.pDate(j) = Int(m.Bars(eBARS_DateTime, i))
        m.pDate(j) = m.Bars(eBARS_DateTime, i)
        
        m.pOpen(j) = gdGetNum(hArray1, i)
        m.pHigh(j) = gdGetNum(hArray2, i)
        m.pLow(j) = gdGetNum(hArray3, i)
        m.pClose(j) = gdGetNum(hArray4, i)
        
        'hourmin & vol info are not used by PFP dll
        'vol info arrays are currently arrays of "short"
        'can cause overflow for volume of bars > daily for certain symbols (e.g. AXP)
        m.pHourMin(j) = 0
        m.pVol(j) = 0
        m.pOI(j) = 0
        m.pTotVol(j) = 0
        m.pTotOI(j) = 0
        j = j + 1
    Next
    j = j - 1
    
    m.PatrnCore.jdate_ptr = GetAddress(m.pDate(0))
    m.PatrnCore.hourmin_ptr = GetAddress(m.pHourMin(0))
    m.PatrnCore.open_ptr = GetAddress(m.pOpen(0))
    m.PatrnCore.high_ptr = GetAddress(m.pHigh(0))
    m.PatrnCore.low_ptr = GetAddress(m.pLow(0))
    m.PatrnCore.close_ptr = GetAddress(m.pClose(0))
    m.PatrnCore.vol_ptr = GetAddress(m.pVol(0))
    m.PatrnCore.oi_ptr = GetAddress(m.pOI(0))
    m.PatrnCore.tot_vol_ptr = GetAddress(m.pTotVol(0))
    m.PatrnCore.tot_oi_ptr = GetAddress(m.pTotOI(0))
    
    m.PatrnCore.max_bar = j
    m.PatrnCore.first_bar = 1
    m.PatrnCore.last_bar = j
    m.pNumLoaded = j
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.LoadPatternCore"

End Function

Private Function LoadSearchCore() As String
On Error GoTo ErrSection:
    
    Dim i&, j&
    Dim idxLast&, idxFirst&
    Dim dLowest#, dHighest#, dAdjust#
    
    Dim hArray1&, hArray2&, hArray3&, hArray4&
        
    i = m.Bars.Size + 1
    
    ReDim m.gDate(i) As Double
    ReDim m.gHourMin(i) As Long
    ReDim m.gArray1(i) As Double
    ReDim m.gArray2(i) As Double
    ReDim m.gArray3(i) As Double
    ReDim m.gArray4(i) As Double
        
    ReDim m.gOpen(i) As Double
    ReDim m.gHigh(i) As Double
    ReDim m.gLow(i) As Double
    ReDim m.gClose(i) As Double
    
    ReDim m.gVol(i) As Long
    ReDim m.gOI(i) As Long
    ReDim m.gTotVol(i) As Long
    ReDim m.gTotOI(i) As Long
    
    SetArrayHandles hArray1, hArray2, hArray3, hArray4
    
'JM 07-23-2010
'Begin - code copied from AdjustBarPrices
    dLowest = gdMinValue(m.Bars.ArrayHandle(eBARS_Low), 0, m.Bars.Size - 1)
    dHighest = gdMaxValue(m.Bars.ArrayHandle(eBARS_High), 0, m.Bars.Size - 1)
    If dLowest < 0 Then
        ' TLB: if the lowest is negative, we are most likely dealing with a back-adjusted future
        ' so we should shift everything up so the lowest is about half the current highest
        ' in order to make all the percentage changes about right.
        If dHighest > 0 Then
            dAdjust = dHighest / 2 + Abs(dLowest)
        Else
            dAdjust = Abs(dLowest) * 2
        End If
    End If
'End - code copied from AdjustBarPrices
    
    j = 1
    For i = 0 To m.Bars.Size - 1
        'hourmin & vol info are not used by PFP dll
        m.gDate(j) = m.Bars(eBARS_DateTime, i)
        m.gHourMin(j) = 0
        
        'array of data for correlation matches
        m.gArray1(j) = gdGetNum(hArray1, i)
        m.gArray2(j) = gdGetNum(hArray2, i)
        m.gArray3(j) = gdGetNum(hArray3, i)
        m.gArray4(j) = gdGetNum(hArray4, i)
        
        'array of OHLC for display
        m.gOpen(j) = m.Bars(eBARS_Open, i) + dAdjust
        m.gHigh(j) = m.Bars(eBARS_High, i) + dAdjust
        m.gLow(j) = m.Bars(eBARS_Low, i) + dAdjust
        m.gClose(j) = m.Bars(eBARS_Close, i) + dAdjust
        
        'vol info arrays are currently arrays of "short"
        'can cause overflow for volume of bars > daily for certain symbols (e.g. AXP)
        m.gVol(j) = 0
        m.gOI(j) = 0
        m.gTotVol(j) = 0
        m.gTotOI(j) = 0
        j = j + 1
    Next
    j = j - 1

    m.SearchCore.jdate_ptr = GetAddress(m.gDate(0))
    m.SearchCore.hourmin_ptr = GetAddress(m.gHourMin(0))
    m.SearchCore.open_ptr = GetAddress(m.gArray1(0))
    m.SearchCore.high_ptr = GetAddress(m.gArray2(0))
    m.SearchCore.low_ptr = GetAddress(m.gArray3(0))
    m.SearchCore.close_ptr = GetAddress(m.gArray4(0))
    m.SearchCore.vol_ptr = GetAddress(m.gVol(0))
    m.SearchCore.oi_ptr = GetAddress(m.gOI(0))
    m.SearchCore.tot_vol_ptr = GetAddress(m.gTotVol(0))
    m.SearchCore.tot_oi_ptr = GetAddress(m.gTotOI(0))
    
    m.SearchCore.max_bar = j
    m.SearchCore.first_bar = 1
    m.SearchCore.last_bar = j
    m.gNumLoaded = j

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.LoadSearchCore"

End Function

Private Sub LoadCompCore(ByVal Num&)
On Error GoTo ErrSection:

    ReDim m.mDate(Num) As Double
    ReDim m.mHourMin(Num) As Long
    ReDim m.mOpen(Num) As Double
    ReDim m.mHigh(Num) As Double
    ReDim m.mLow(Num) As Double
    ReDim m.mClose(Num) As Double
    ReDim m.mVol(Num) As Long
    ReDim m.mOI(Num) As Long
    ReDim m.mTotVol(Num) As Long
    ReDim m.mTotOI(Num) As Long
    
    m.CompositeCore.jdate_ptr = GetAddress(m.mDate(0))
    m.CompositeCore.hourmin_ptr = GetAddress(m.mHourMin(0))
    m.CompositeCore.open_ptr = GetAddress(m.mOpen(0))
    m.CompositeCore.high_ptr = GetAddress(m.mHigh(0))
    m.CompositeCore.low_ptr = GetAddress(m.mLow(0))
    m.CompositeCore.close_ptr = GetAddress(m.mClose(0))
    m.CompositeCore.vol_ptr = GetAddress(m.mVol(0))
    m.CompositeCore.oi_ptr = GetAddress(m.mOI(0))
    m.CompositeCore.tot_vol_ptr = GetAddress(m.mTotVol(0))
    m.CompositeCore.tot_oi_ptr = GetAddress(m.mTotOI(0))

    m.CompositeCore.max_bar = Num
    m.CompositeCore.first_bar = 1
    m.CompositeCore.last_bar = Num

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.LoadCompCore"

End Sub

Public Sub CompositeData(ByVal idx&, ByVal dBase#, ByVal dDate#, _
    ByRef dOpen#, ByRef dHigh#, ByRef dLow#, ByRef dClose#, _
    ByRef dStdAbove#, ByRef dStdBelow#)
On Error GoTo ErrSection:
    
    Dim i&
    
    dOpen = kNullData
    dHigh = kNullData
    dLow = kNullData
    dClose = kNullData
    dStdAbove = kNullData
    dStdBelow = kNullData
    
    For i = 1 To m.mNumLoaded
        If m.mDate(i) = dDate Then
            dOpen = (m.mOpen(i) - 100) / 100 * dBase + dBase
            dHigh = (m.mHigh(i) - 100) / 100 * dBase + dBase
            dLow = (m.mLow(i) - 100) / 100 * dBase + dBase
            dClose = (m.mClose(i) - 100) / 100 * dBase + dBase
            
            If m.mClose(i) = 0 And m.pfp_strength(i) = 0 Then
                'do nothing, this will simply return zero for std dev values
                'the chart will end up drawing a line to the bottom of the price pane
                'known case of when this is true is when there is only one hit
            Else
                dStdAbove = (m.mClose(i) + m.pfp_strength(i) * m.dStandardDev - 100) / 100 * dBase + dBase
                dStdBelow = (m.mClose(i) - m.pfp_strength(i) * m.dStandardDev - 100) / 100 * dBase + dBase
            End If
            
            Exit For
        End If
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.CompositeData"

End Sub

Public Property Get CompositeCount() As Long
On Error Resume Next
    
    CompositeCount = m.mNumLoaded

End Property

Public Property Get MatchCount() As Long
On Error Resume Next
    
    If UBound(m.pfp_hits) > 0 Then MatchCount = m.pfp_hits(0)
    
End Property

' Find the BEST point on the worst-case std dev line
' (i.e. look at the std dev line that is closest to zero,
' and look for the highest point on it if it's the lower line
' or the lowest point on it if it's the upper line)
Private Function ForecastRank(ByVal nPatternLen&) As Double
On Error GoTo ErrSection:

    Dim b#, F#, i&, p&, l#, h#, max_f#

    p = nPatternLen
    max_f = -99999
    b = m.mClose(p) 'base
    For i = p + 1 To p + m.nForecastBars
        l = (m.mClose(i) - m.pfp_strength(i)) - b
        h = (m.mClose(i) + m.pfp_strength(i)) - b
        If Abs(l) < Abs(h) Then
            F = Abs(l)
        Else
            F = Abs(h)
        End If
        If h * l < 0 Then F = -F
        If F > max_f Then
            max_f = F
        End If
    Next i

    ForecastRank = max_f

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.ForecastRank"

End Function

Private Sub AddIndicator(ByVal iIndNum&, Ind As cIndicator, Chart As cChart)
On Error GoTo ErrSection:

    Dim hArray&, eFlag As ePfpIndFlags, bCheckZero As Boolean
    Dim Pane As cPane

    hArray = GetDataArrayHandle(Ind, iIndNum)
    If hArray <> 0 Then
        ' set the Flags (bit-mask) for this indicator
        eFlag = 0
        If Ind.Overlayed Then
            eFlag = eFlag Or ePfp_IsOverlayed
        End If
        If Ind.DataType = eINDIC_BooleanArray Then
            eFlag = eFlag Or ePfp_IsBoolean
        End If
        If Ind.DataType = eINDIC_Constant Then
            eFlag = eFlag Or ePfp_IsConstant
        End If
        ' "crosses zero" should only be set for indicators that cross zero (max > 0 and min < 0)
        ' but which are not "price data" (e.g. bars) or tied to price data (e.g. non-overlayed in a price pane)
        If Ind.DataType <> eINDIC_BarData And Ind.DataType <> eINDIC_BooleanArray And Ind.DataType <> eINDIC_Constant Then
            bCheckZero = True
            If Not Ind.Overlayed Then
                ' check if this indicator is tied to a "price scale" (i.e. scale of a bars)
                Set Pane = Chart.Tree.RelativeItem(Ind.geIndId, eTREE_Root)
                If Not Pane Is Nothing Then
                    If Pane.DisplayFormat = ePANE_PriceFormat Then
                        bCheckZero = False ' do not check for crossing zero if tied to a price scale
                    End If
                End If
            End If
            If bCheckZero Then
                If gdMinValue(hArray, 0, gdGetSize(hArray)) < 0 Then
                    If gdMaxValue(hArray, 0, gdGetSize(hArray)) > 0 Then
                        eFlag = eFlag Or ePfp_CrossesZero
                    End If
                End If
            End If
        End If
        
        m.aIndArrays.Add hArray
        m.aPaneIDs.Add Ind.geIndpaneId
        m.aIndFlags.Add eFlag
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.AddIndicator"
End Sub

Private Function GetMatches3(Chart As cChart) As Long
On Error GoTo ErrSection:

    Dim i&, iPricePaneID&, iPtrSave&
    
    Dim ptrn_len&, fcast_len&, num_comp&
    Dim max_bars&, min_hits&, max_hits&
    
    Dim lowest_corr&, min_corr&
    Dim best_bars&, best_rank#, best_corr&, p#

    If Chart Is Nothing Then Exit Function
    If m.Bars Is Nothing Then Exit Function
    If m.Ind1 Is Nothing And m.Ind2 Is Nothing And m.Ind3 Is Nothing And m.Ind4 Is Nothing Then Exit Function
        
    If Not Chart.Tree("PRICE PANE") Is Nothing Then
        iPricePaneID = Chart.Tree("PRICE PANE").gePaneId
    End If
    
    m.aIndArrays.Size = 0
    m.aPaneIDs.Size = 0
    m.aIndFlags.Size = 0

    AddIndicator 0, m.Ind1, Chart
    AddIndicator 1, m.Ind2, Chart
    AddIndicator 2, m.Ind3, Chart
    AddIndicator 3, m.Ind4, Chart
    
    best_bars = 0
    best_rank = -99999
    min_corr = m.nPercentCorr
    
    ptrn_len = m.nPatternSelLen             'm.nPatternSelLen: selected & enclosed within rectangle by user
    max_bars = ptrn_len
    fcast_len = m.nForecastBars
    
    m.nBestCorr = -1
    
    iPtrSave = Screen.MousePointer
    
    If m.bOptimize Then
        ptrn_len = m.nPtrnLen               'm.nPtrnLen: set in optimizer form and read in from INI file
        If ptrn_len < 2 Then ptrn_len = 2
        max_bars = m.nMaxBars
        If max_bars < ptrn_len Then max_bars = ptrn_len
        min_hits = m.nMinHits
        lowest_corr = m.nLowestCorr
        If lowest_corr < 1 Or lowest_corr > 100 Then lowest_corr = 80
        min_corr = 98
        Screen.MousePointer = vbHourglass
    End If
    If ptrn_len < 1 Then
        InfBox "i=[] ; Invalid pattern length!"
        GoTo ErrExit
    End If
    If fcast_len < 1 Then
        InfBox "i=[] ; Invalid forecast length!"
        GoTo ErrExit
    End If
        
    max_hits = 1000

    ' when optimizing, just need to loop once for each possible pattern length,
    ' then one more time again to redo the best one (bOptimize will be false by then)
    Do While True
        If m.bOptimize Then
            min_corr = lowest_corr
            InfBox "Optimizing:  pattern length = " & Str(ptrn_len) & " bars", , , "Pattern for Profit", True
        End If
        ApplyNewFilter Chart, ptrn_len, min_corr
        If Not m.bOptimize Then Exit Do
        
        ' if optimizing, find the highest correlation that has at least
        ' the minimum # of hits (these arrays are now sorted by correlation)
        If m.pfp_hits(0) >= min_hits Then
            ' get the correlation at the min_hits spot (use integer at or below)
            min_corr = Int(m.pfp_corr(min_hits) * 100)
            If min_corr > 99 Then min_corr = 99
            For i = min_hits + 1 To m.pfp_hits(0)
                ' look for where it drops to a lower correlation
                If Int(m.pfp_corr(i) * 100) < min_corr Then
                    ' set the "size" of the pfp arrays such that the rest will be ignored
                    m.pfp_hits(0) = i - 1
                    m.pfp_corr(0) = i - 1
                    Exit For
                End If
            Next
        End If
    
        If m.pfp_hits(0) >= min_hits Then
            ' build the composite forecast for these hits
            m.mNumLoaded = ptrn_len + fcast_len
            LoadCompCore m.mNumLoaded + 10
            ReDim m.pfp_strength#(m.mNumLoaded + 10)
            num_comp = m.pfp_hits(0)
            
            m.SearchCore.open_ptr = GetAddress(m.gOpen(0))
            m.SearchCore.high_ptr = GetAddress(m.gHigh(0))
            m.SearchCore.low_ptr = GetAddress(m.gLow(0))
            m.SearchCore.close_ptr = GetAddress(m.gClose(0))
            
            i = PFP_BuildComposite2(m.SearchCore, _
                                    m.match_type, _
                                    ptrn_len, _
                                    fcast_len, _
                                    num_comp, _
                                    m.pfp_hits(0), _
                                    m.CompositeCore, _
                                    m.pfp_strength(0))
            
            ' compare how this forecast ranks against the others
            ' (looking for the one that provides the best worst-case scenario)
            p = ForecastRank(ptrn_len)
            If p > best_rank Then
                best_rank = p
                best_bars = ptrn_len
                best_corr = min_corr
            End If
        End If
    
        ptrn_len = ptrn_len + 1
        If ptrn_len > max_bars Then
            ' rerun one more time for the best pattern length
            m.bOptimize = False
            ptrn_len = best_bars
            min_corr = best_corr
            m.nBestCorr = best_corr
        End If
    Loop

    If min_corr = 0 Then
        m.pfp_hits(0) = 0
    ElseIf m.pfp_hits(0) >= max_hits - 2 Then
        InfBox "i=! ; Too many matches -- please provide a greater restriction."
        m.pfp_hits(0) = 0
        m.mNumLoaded = 0
    End If

ErrExit:
    GetMatches3 = ptrn_len
    Screen.MousePointer = iPtrSave
    Exit Function

ErrSection:
    GetMatches3 = 0
    Screen.MousePointer = iPtrSave
    RaiseError "cPatternProfit.GetMatches3"

End Function

Private Function GetMatches(fgHits As VSFlexGrid, Chart As cChart) As Long
On Error GoTo ErrSection:

    Dim rc&, Num&, temp$, i&, min_corr&, d, j&, p#
    Dim max_hits&, fcast_len&, ptrn_len&
    Dim num_rules&, num_comp&, max_bars&
    Dim best_bars&, best_rank#, best_corr&, min_hits&, lowest_corr&

    Dim filtered&()
    Dim fh&, fName$, warning&
    
    best_bars = 0
    best_rank = -99999
    min_corr = m.nPercentCorr
    
    ptrn_len = m.nPatternSelLen             'm.nPatternSelLen: selected & enclosed within rectangle by user
    max_bars = ptrn_len
    fcast_len = m.nForecastBars
    
    If m.bOptimize Then
        ptrn_len = m.nPtrnLen               'm.nPtrnLen: set in optimizer form and read in from INI file
        If ptrn_len < 2 Then ptrn_len = 2
        max_bars = m.nMaxBars
        If max_bars < ptrn_len Then max_bars = ptrn_len
        min_hits = m.nMinHits
        lowest_corr = m.nLowestCorr
        If lowest_corr < 1 Or lowest_corr > 100 Then lowest_corr = 80
        min_corr = 98
    End If
    If ptrn_len < 1 Then
        InfBox "i=[] ; Invalid pattern length!"
        GoTo ErrExit
    End If
    If fcast_len < 1 Then
        InfBox "i=[] ; Invalid forecast length!"
        GoTo ErrExit
    End If
        
    max_hits = 1000
    
    ReDim m.pfp_hits(max_hits)
    ReDim m.pfp_corr#(max_hits)
    ReDim filtered(m.gNumLoaded + 10) As Long

Do While min_corr <> 0
    LoadPatternCore ptrn_len
    If m.pNumLoaded <= 0 Then Exit Do

    m.SearchCore.open_ptr = GetAddress(m.gArray1(0))
    m.SearchCore.high_ptr = GetAddress(m.gArray2(0))
    m.SearchCore.low_ptr = GetAddress(m.gArray3(0))
    m.SearchCore.close_ptr = GetAddress(m.gArray4(0))
    
    rc = PFP_CorrelationMatches2(m.PatrnCore, _
                                 m.SearchCore, _
                                 m.match_type, _
                                 filtered(0), _
                                 min_corr, _
                                 max_hits, _
                                 m.pfp_hits(0), _
                                 m.pfp_corr(0))
    If m.bOptimize Then
        If m.pfp_hits(0) < min_hits Then
            min_corr = min_corr - 2
            rc = 0
        End If
    End If

    If rc > 0 Then
        m.mNumLoaded = ptrn_len + fcast_len
        LoadCompCore m.mNumLoaded + 10
        ReDim m.pfp_strength#(m.mNumLoaded + 10)
        num_comp = m.pfp_hits(0)
        'If chkExcludeLast <> 0 And num_comp > 0 Then num_comp = num_comp - 1
        
        m.SearchCore.open_ptr = GetAddress(m.gOpen(0))
        m.SearchCore.high_ptr = GetAddress(m.gHigh(0))
        m.SearchCore.low_ptr = GetAddress(m.gLow(0))
        m.SearchCore.close_ptr = GetAddress(m.gClose(0))
        
        rc = PFP_BuildComposite2(m.SearchCore, _
                                 m.match_type, _
                                 ptrn_len, _
                                 fcast_len, _
                                 num_comp, _
                                 m.pfp_hits(0), _
                                 m.CompositeCore, _
                                 m.pfp_strength(0))
        If m.bOptimize Then
            p = ForecastRank(ptrn_len)
            If p > best_rank Then
                best_rank = p
                best_bars = ptrn_len
                best_corr = min_corr
            End If
            min_corr = -min_corr
        Else
            fName = "pfp.chk"
            KillFile fName
            
            fh = FreeFile
            If fh Then Open fName For Output As #fh
            If fh Then Print #fh, m.pNumLoaded, m.mNumLoaded
            For i = 1 To m.pNumLoaded
                If fh Then Print #fh, m.pDate(i), m.pOpen(i), m.pHigh(i), m.pLow(i), m.pClose(i)
            Next
            For i = 1 To m.mNumLoaded
                If fh Then Print #fh, i, m.mOpen(i), m.mHigh(i), m.mLow(i), m.mClose(i), m.pfp_strength(i)
                If m.mOpen(i) > m.mHigh(i) Or m.mLow(i) > m.mHigh(i) Or m.mClose(i) > m.mHigh(i) Then
                    warning = i
                End If
                If m.mOpen(i) < m.mLow(i) Or m.mHigh(i) < m.mLow(i) Or m.mClose(i) < m.mLow(i) Then
                    warning = i
                End If
            Next
            If fh Then
                Close #fh
                If warning Then InfBox "i=! ; h=Warning ; OHLC anomaly at" + Str(warning)
            End If
            Exit Do
        End If
    End If

    If Not m.bOptimize Then
        Exit Do
    ElseIf min_corr < lowest_corr Then
        min_corr = 100 'Abs(min_corr) + 2
        If min_corr > 98 Then min_corr = 98
        ptrn_len = ptrn_len + 1
        If ptrn_len > max_bars Then
            m.bOptimize = False
            ptrn_len = best_bars
            min_corr = best_corr
        End If
    End If
Loop

    If min_corr = 0 Then
        m.pfp_hits(0) = 0
    Else
        m.nPercentCorr = min_corr
    End If
    
    If m.pfp_hits(0) >= max_hits - 2 Then
        InfBox "i=! ; Too many matches -- please provide a greater restriction."
        m.pfp_hits(0) = 0
        m.mNumLoaded = 0
        GoTo ErrExit
    End If
    
'    PopulateGrid fgHits, ptrn_len
    
ErrExit:
    GetMatches = ptrn_len
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.GetMatches"

End Function

Private Sub SetArrayHandles(ByRef h1&, ByRef h2&, ByRef h3&, ByRef h4)
On Error GoTo ErrSection:

    Dim strText$
    Dim i As Long

    Dim bOpen As Boolean
    Dim bHigh As Boolean
    Dim bLow As Boolean
    Dim bClose As Boolean
    
    h1 = 0
    h2 = 0
    h3 = 0
    h4 = 0
    
    For i = 0 To m.aIndNames.Size - 1
        strText = UCase(Parse(m.aIndNames(i), ";", 1))
    
        Select Case strText
            Case "OPEN"
                h1 = m.Bars.ArrayHandle(eBARS_Open)
            Case "HIGH"
                h2 = m.Bars.ArrayHandle(eBARS_High)
            Case "LOW"
                h3 = m.Bars.ArrayHandle(eBARS_Low)
            Case "CLOSE"
                h4 = m.Bars.ArrayHandle(eBARS_Close)
        End Select
    Next
       
    If Not m.Ind1 Is Nothing Then
        If m.Ind1.Data.Size = m.Bars.Size Then     'price will have size=0 for Ind.Data
            If h1 = 0 Then
                h1 = m.Ind1.Data.ArrayHandle
            ElseIf h2 = 0 Then
                h2 = m.Ind1.Data.ArrayHandle
            ElseIf h3 = 0 Then
                h3 = m.Ind1.Data.ArrayHandle
            ElseIf h4 = 0 Then
                h4 = m.Ind1.Data.ArrayHandle
            End If
        End If
    End If
    
    If Not m.Ind2 Is Nothing Then
        If m.Ind2.Data.Size = m.Bars.Size Then
            If h1 = 0 Then
                h1 = m.Ind2.Data.ArrayHandle
            ElseIf h2 = 0 Then
                h2 = m.Ind2.Data.ArrayHandle
            ElseIf h3 = 0 Then
                h3 = m.Ind2.Data.ArrayHandle
            ElseIf h4 = 0 Then
                h4 = m.Ind2.Data.ArrayHandle
            End If
        End If
    End If

    If Not m.Ind3 Is Nothing Then
        If m.Ind3.Data.Size = m.Bars.Size Then
            If h1 = 0 Then
                h1 = m.Ind3.Data.ArrayHandle
            ElseIf h2 = 0 Then
                h2 = m.Ind3.Data.ArrayHandle
            ElseIf h3 = 0 Then
                h3 = m.Ind3.Data.ArrayHandle
            ElseIf h4 = 0 Then
                h4 = m.Ind3.Data.ArrayHandle
            End If
        End If
    End If

    If Not m.Ind4 Is Nothing Then
        If m.Ind4.Data.Size = m.Bars.Size Then
            If h1 = 0 Then
                h1 = m.Ind4.Data.ArrayHandle
            ElseIf h2 = 0 Then
                h2 = m.Ind4.Data.ArrayHandle
            ElseIf h3 = 0 Then
                h3 = m.Ind4.Data.ArrayHandle
            ElseIf h4 = 0 Then
                h4 = m.Ind4.Data.ArrayHandle
            End If
        End If
    End If

    m.match_type = 0
    
    If h1 = 0 Then
        h1 = m.Bars.ArrayHandle(eBARS_Open)     'don't want to pass null pointer into DLL
    Else
        m.match_type = m.match_type Or MATCH_OPEN
    End If
    
    If h2 = 0 Then
        h2 = m.Bars.ArrayHandle(eBARS_High)
    Else
        m.match_type = m.match_type Or MATCH_HIGH
    End If
    
    If h3 = 0 Then
        h3 = m.Bars.ArrayHandle(eBARS_Low)
    Else
        m.match_type = m.match_type Or MATCH_LOW
    End If
    
    If h4 = 0 Then
        h4 = m.Bars.ArrayHandle(eBARS_Close)
    Else
        m.match_type = m.match_type Or MATCH_CLOSE
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPatternProfit.SetArrayHandles"

End Sub

Public Function PatternAnnotCheck(Chart As cChart, frm As Form, ByVal bAdd As Boolean, _
    Optional ByRef AnnotIn As cAnnotation) As cAnnotation
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    Dim Pane As cPane
    Dim Tree As cGdTree
    
    Dim i&, dDateTo#, dDateFrom#, strKey$
    Dim bFound As Boolean
    
    If Chart Is Nothing Then Exit Function
    If Chart.Tree Is Nothing Then Exit Function
    If Chart.Annots Is Nothing Then Exit Function
    
    Set Tree = Chart.Tree
    For i = 1 To Tree.Count
        If Tree.NodeLevel(i) = 0 Then
            Set Pane = Tree(i)
            If Pane.Display Then Exit For       'want annot to be associated with first visible pane
        End If
    Next
    If Pane Is Nothing Then Exit Function
    
    If AnnotIn Is Nothing Then
        dDateTo = Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False) + 1)
        'm.nPatternSelLen = 3
        Set Tree = Chart.Annots
        For i = 1 To Tree.Count
            Set Annot = Tree(i)
            If Not Annot Is Nothing Then
                If Annot.eUsage = eANNOT_PatternProfit And InStr(Tree.Key(i), "PFP") = 0 Then
                    bFound = True
                    Exit For
                End If
            End If
        Next
    Else
        If AnnotIn.dDate(2) > AnnotIn.dDate(1) Then
            dDateTo = AnnotIn.dDate(2)
            dDateFrom = AnnotIn.dDate(1)
        Else
            dDateTo = AnnotIn.dDate(1)
            dDateFrom = AnnotIn.dDate(2)
        End If
        Chart.RemoveAnnots True, , eANNOT_PatternProfit
        bAdd = True
    End If
    
    
    If Not bAdd Or bFound Then GoTo ErrExit
    
    Set Annot = New cAnnotation
    If dDateFrom = 0 Then
        i = Annot.CreateNew(Chart, eANNOT_Rectangle, Pane.gePaneId, dDateTo, 0, dDateTo, 0, m.nForecastColor, , , eANNOT_PatternProfit)
    Else
        i = Annot.CreateNew(Chart, eANNOT_Rectangle, Pane.gePaneId, dDateFrom, 0, dDateTo, 0, m.nForecastColor, , , eANNOT_PatternProfit)
    End If
    
    If i <> 0 Then
        Annot.eUsage = eANNOT_PatternProfit
        Annot.Prop("Shape") = 0
        Annot.Prop("FillPattern") = m.nFillPattern
        Annot.Prop("FillColor") = m.nFillColor
        Annot.Color = m.nForecastColor
        Annot.PreIndicator = 1
        Annot.SkipHitTestFlag = 1
        
        If dDateFrom = 0 Then
            dDateFrom = Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False) - m.nPatternSelLen + 1)
            Annot.dDate(1) = dDateFrom
            Annot.dDate(2) = dDateTo
        End If
        
        i = Chart.Annots.Add(Annot)
        strKey = Chart.Annots.Key(i)
        Annot.Prop("AnnotKey") = strKey
        i = Annot.PatternLength(Nothing)        'this is just to get a flag set
        Chart.GenerateChart eRedo1_Scrolled
        bFound = True
    End If

ErrExit:
    If bFound Then Set PatternAnnotCheck = Annot
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.PatternAnnotCheck"

End Function

Public Sub LoadMatchAnnot(Chart As cChart, frm As Form, ByVal dDateFrom#, ByVal dDateTo#)
On Error GoTo ErrSection:

    Dim i&, dDateMatch#
    
    For i = m.pfp_hits(0) To 1 Step -1
        If m.pfp_hits(i) + (m.nPatternSelLen - 1) <= m.gNumLoaded Then
            dDateMatch = m.gDate(m.pfp_hits(i) + (m.nPatternSelLen - 1))
            If dDateMatch >= dDateFrom And dDateMatch <= dDateTo Then
                MatchAnnotShow Chart, frm, dDateMatch
            End If
        End If
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.LoadMatchAnnot"

End Sub

Public Sub MatchAnnotShow(Chart As cChart, frm As Form, ByVal dDate#)
On Error GoTo ErrSection:

    Dim i&, j&, nID&
    Dim strKey$, dDateFrom#, dDateTo#
    Dim bCompare As Boolean
    
    Dim Annot As cAnnotation
    Dim Pane As cPane
    Dim Tree As cGdTree
    
    Dim aPaneIDs As New cGdArray
    
    If m.bOptimize Then Exit Sub            'optimization in progress, don't process annots yet
    
    If Chart Is Nothing Then Exit Sub
    If Chart.Annots Is Nothing Then Exit Sub
    If dDate <= 0 Then Exit Sub
    If gdFixDateTime(dDate) = gdFixDateTime(m.dPatternDateTo) Then Exit Sub
    
    Set Pane = Chart.Tree("PRICE PANE")
    If Pane Is Nothing Then Exit Sub
    
    'add pane ID to an array to determine which pane(s) will need a pattern-match annotation to be drawn on chart
    If Not m.Ind1 Is Nothing Then
        nID = m.Ind1.geIndpaneId
        If nID <> Pane.gePaneId Then aPaneIDs.Add nID
    End If
    
    If Not m.Ind2 Is Nothing Then
        nID = m.Ind2.geIndpaneId
        If nID <> Pane.gePaneId Then aPaneIDs.Add nID
    End If
    
    If Not m.Ind3 Is Nothing Then
        nID = m.Ind3.geIndpaneId
        If nID <> Pane.gePaneId Then aPaneIDs.Add nID
    End If
    
    If Not m.Ind4 Is Nothing Then
        nID = m.Ind4.geIndpaneId
        If nID <> Pane.gePaneId Then aPaneIDs.Add nID
    End If
    
    If aPaneIDs.Size > 0 Then
        aPaneIDs.Sort eGdSort_Default Or eGdSort_DeleteDuplicates
    End If
    aPaneIDs.Add Pane.gePaneId, 0
    
    Set Tree = Chart.Annots
    
    'loop through tree & add any pattern-match annotation not already in tree
    For i = 0 To aPaneIDs.Size - 1
        Set Pane = Nothing
        nID = aPaneIDs(i)
        If Chart.Tree.NodeLevel(nID) = 0 Then Set Pane = Chart.Tree(nID)
        
        If Not Pane Is Nothing Then
            strKey = "PFP_" & Str(dDate) & "_" & Str(nID)
    
            Set Annot = Tree(strKey)
            If Annot Is Nothing Then
                bCompare = False
                Set Annot = New cAnnotation
                j = Annot.CreateNew(Chart, eANNOT_Rectangle, nID, dDate, 0, dDate, 0, m.nForecastColor)
                If j <> 0 Then
                    Annot.eUsage = eANNOT_PatternProfit
                    Annot.Prop("Shape") = 0
                    Annot.Prop("FillPattern") = m.nFillPattern
                    Annot.Prop("FillColor") = m.nFillColor
                    Annot.PreIndicator = 1
                    Annot.SkipHitTestFlag = 1
                    j = m.Bars.FindDateTime(dDate, True)
                    If m.Bars(eBARS_DateTime, j) = gdFixDateTime(dDate) Then

'JM 07-08-2010 - code prior to grapheng.dll chane (save for reference)
'rectangles are drawn with side lines "in betwen" bars so need to add 1 to the "To" date
'                        dDateTo = m.Bars(eBARS_DateTime, j + 1)
                        
                        dDateTo = m.Bars(eBARS_DateTime, j)
                        dDateFrom = m.Bars(eBARS_DateTime, j - m.nPatternSelLen + 1)
                        Annot.dDate(1) = dDateFrom
                        Annot.dDate(2) = dDate
                        
                        If i = 0 Then
                            'set y-values based on min/max of price bars
                            Annot.SetMinMaxPFP Chart, Nothing, False
                            bCompare = True
                        End If
                        
                        'if bCompare = false, annot y-values will be set = to min/max of passed-in indicator
                        'if bcompare = true, current y-value of annot will be changed only if min/max of
                        '                    passed-in indicator is lower or higher
                        If Not m.Ind1 Is Nothing Then
                            If nID = m.Ind1.geIndpaneId Then
                                Annot.SetMinMaxPFP Chart, m.Ind1, bCompare
                                bCompare = True
                            End If
                        End If
                        
                        If Not m.Ind2 Is Nothing Then
                            If nID = m.Ind2.geIndpaneId Then
                                Annot.SetMinMaxPFP Chart, m.Ind2, bCompare
                                bCompare = True
                            End If
                        End If
                        
                        If Not m.Ind3 Is Nothing Then
                            If nID = m.Ind3.geIndpaneId Then
                                Annot.SetMinMaxPFP Chart, m.Ind3, bCompare
                                bCompare = True
                            End If
                        End If
                        
                        If Not m.Ind4 Is Nothing Then
                            If nID = m.Ind4.geIndpaneId Then Annot.SetMinMaxPFP Chart, m.Ind4, bCompare
                        End If
                        
                        Annot.dDate(2) = dDateTo
                    
                        j = Chart.Annots.Add(Annot, strKey)
                    End If
                End If
            End If
        End If
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPatternProfit.MatchAnnotShow"

End Sub

Public Function RebuildComposite(fgHits As VSFlexGrid) As Boolean
On Error GoTo ErrSection:

    Dim i&, j&, rc&, iCount&
    Dim ptrn_len&, fcast_len&
    Dim filtered_hits() As Long
    

    If fgHits Is Nothing Then Exit Function
    If fgHits.Rows <= fgHits.FixedRows Then Exit Function
    If m.pfp_hits(0) <= 0 Then Exit Function
    
    ReDim filtered_hits(m.pfp_hits(0)) As Long
    
    iCount = 0
    With fgHits
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, eColsPFP_Use) = flexChecked Then
                j = Val(.TextMatrix(i, eColsPFP_Index))
                filtered_hits(iCount) = m.pfp_hits(j)
                iCount = iCount + 1
            End If
        Next
    End With
    
    If iCount < 1 Then
        InfBox "There must be at least 1 match selected.", "I", , "Patterns for Profit"     '5887
        Exit Function
    End If
    
    filtered_hits(0) = iCount - 1
    
    ptrn_len = m.nPatternSelLen
    fcast_len = m.nForecastBars
        
    m.mNumLoaded = ptrn_len + fcast_len
    LoadCompCore m.mNumLoaded + 10
    ReDim m.pfp_strength(m.mNumLoaded + 10)
    iCount = filtered_hits(0)
        
    m.SearchCore.open_ptr = GetAddress(m.gOpen(0))
    m.SearchCore.high_ptr = GetAddress(m.gHigh(0))
    m.SearchCore.low_ptr = GetAddress(m.gLow(0))
    m.SearchCore.close_ptr = GetAddress(m.gClose(0))
    
    rc = PFP_BuildComposite2(m.SearchCore, _
                             m.match_type, _
                             ptrn_len, _
                             fcast_len, _
                             iCount, _
                             filtered_hits(0), _
                             m.CompositeCore, _
                             m.pfp_strength(0))

    
    If rc = 0 Then RebuildComposite = True
    
    SetCompositeDates fgHits.Parent.Chart

ErrExit:
    Exit Function

ErrSection:
    m.mNumLoaded = 0
    RaiseError "cPatternProfit.RebuildComposite"

End Function

Public Sub UpdateProperty(Chart As cChart, ByVal bRedraw As Boolean, _
    Optional ByVal nColor& = -1, Optional ByVal nFillColor& = -1, _
    Optional ByVal nFillStyle& = -1, Optional ByVal dStandardDev# = -1#, _
    Optional ByVal nLineStyle& = -1, Optional ByVal bMethodChanged As Boolean = False)
On Error GoTo ErrSection:

    Dim i&
    Dim Annot As cAnnotation
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If nColor = -1 And nFillColor = -1 And nFillStyle = -1 And dStandardDev# = -1# And nLineStyle = -1 And Not bMethodChanged Then Exit Sub 'nothing passed in
    If m.nForecastColor = nColor And m.nFillColor = nFillColor And m.nFillPattern = nFillStyle And Not bMethodChanged Then Exit Sub    'nothing changed
    
    If nColor <> -1 And nColor <> m.nForecastColor Then m.nForecastColor = nColor
    If nFillColor <> -1 And nFillColor <> m.nFillColor Then m.nFillColor = nFillColor
    If nFillStyle <> -1 And nFillStyle <> m.nFillPattern Then m.nFillPattern = nFillStyle
    If dStandardDev <> -1 And dStandardDev <> m.dStandardDev Then m.dStandardDev = dStandardDev
    If nLineStyle <> -1 And nLineStyle <> m.eIndLineSylte Then m.eIndLineSylte = nLineStyle
    
    SaveSettings Chart
    
    If Chart Is Nothing Then Exit Sub
    If Chart.Annots Is Nothing Then Exit Sub
    
    If m.nPatternSelLen <= 0 Then Exit Sub   'no pattern no need to do anything
    
    If m.pfp_hits(0) <= 0 Then
        For i = 1 To Chart.Annots.Count
            Set Annot = Chart.Annots(i)
            If Not Annot Is Nothing Then
                If Annot.eUsage = eANNOT_PatternProfit Then
                    Annot.Color = m.nForecastColor
                    Annot.Prop("FillColor") = m.nFillColor
                    Annot.Prop("FillPattern") = m.nFillPattern
                    Exit For                'no matches, no need to continue
                End If
            End If
        Next
    Else
        For i = 1 To Chart.Annots.Count
            Set Annot = Chart.Annots(i)
            If Not Annot Is Nothing Then
                If Annot.eUsage = eANNOT_PatternProfit Then
                    Annot.Color = m.nForecastColor
                    Annot.Prop("FillColor") = m.nFillColor
                    Annot.Prop("FillPattern") = m.nFillPattern
                End If
            End If
        Next
    End If
    
    If bRedraw Then Chart.GenerateChart eRedo1_Scrolled

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.UpdateProperty"

End Sub

Private Sub SetCompositeDates(Chart As cChart)
On Error GoTo ErrSection:

    Dim i&, j&, hDates&, iSize&
    
    'get handle to the date array for the chart (this array includes blank & forecast bars)
    If Chart Is Nothing Then Exit Sub
    hDates = Chart.aXdate.ArrayHandle
    
    If hDates = 0 Then Exit Sub
    iSize = gdGetSize(hDates)
    If iSize <= 0 Then Exit Sub
    
    If gdBinarySearch(hDates, m.dPatternDateFrom, i, eGdSort_Default, 0, iSize - 1) = 0 Then Exit Sub
    
    For j = 1 To m.mNumLoaded
        m.mDate(j) = gdGetNum(hDates, i)
        i = i + 1
        If i >= iSize Then Exit For
    Next
    
    Dim strDate$
    For i = 1 To m.mNumLoaded
        strDate = DateFormat(m.mDate(i), MM_DD_YYYY)
    Next
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.SetCompositeDates"

End Sub

Private Sub PopulateGrid(fgHits As VSFlexGrid, ByVal ptrn_len&)
On Error GoTo ErrSection:

    Dim i&, dDate#, iBars&
    Dim strText$, temp$
    
    Dim bResetIbars As Boolean
    Dim frmParent As Form
    
    iBars = ptrn_len - 1
    
    Do While m.pfp_hits(m.pfp_hits(0)) + iBars > m.gNumLoaded
        iBars = iBars - 1
        bResetIbars = True
    Loop
    
    fgHits.Redraw = flexRDNone
    fgHits.Rows = fgHits.FixedRows
    
    fgHits.Cols = eColsPFP_DateDouble + 1
    fgHits.ColHidden(eColsPFP_DateDouble) = True
    ' TLB 7/16/2010: we want the CorrPercent column to contain the non-rounded
    ' correlation for sorting purposes (but will still display as a rounded %)
    fgHits.ColDataType(eColsPFP_CorrPercent) = flexDTDouble
    fgHits.ColFormat(eColsPFP_CorrPercent) = "#0%"
    For i = m.pfp_hits(0) To 1 Step -1
        dDate = m.gDate(m.pfp_hits(i) + iBars)
        If bResetIbars Then
            iBars = ptrn_len - 1
            bResetIbars = False
        End If
        
        With fgHits
            .Rows = .Rows + 1
            .Cell(flexcpChecked, .Rows - 1, eColsPFP_Use) = flexChecked
            .Cell(flexcpPictureAlignment, .Rows - 1, eColsPFP_Use) = flexAlignCenterCenter
            .TextMatrix(.Rows - 1, eColsPFP_Day) = Weekday(dDate)
            temp = DateFormat(dDate, MM_DD_YYYY)
            If m.Bars.IsIntraday Then temp = Left(temp, 5) & "  " & DateFormat(dDate, NO_DATE, HH_MM)
            .TextMatrix(.Rows - 1, eColsPFP_Date) = temp
            .TextMatrix(.Rows - 1, eColsPFP_CorrPercent) = m.pfp_corr(i) 'Format(m.pfp_corr(i), "00%")
            .TextMatrix(.Rows - 1, eColsPFP_Index) = Str(i)
            .TextMatrix(.Rows - 1, eColsPFP_DateDouble) = dDate
        End With
    Next
    
    If fgHits.Rows > fgHits.FixedRows Then
        fgHits.Col = eColsPFP_CorrPercent            '04-16-2009: Request from LW per Chad
        fgHits.Sort = flexSortGenericDescending
        fgHits.Row = fgHits.FixedRows
        fgHits.Select fgHits.Row, eColsPFP_Date
        fgHits.SetFocus
    End If
    
    Set frmParent = fgHits.Parent
    
    If IsFrmChart(frmParent) Then
        frmParent.lblHitsFoundPFP.Caption = "Found: " & Str(fgHits.Rows - fgHits.FixedRows)
        
        If ptrn_len > 0 Then
            temp = DateFormat(m.dPatternDateTo, MM_DD_YYYY)
            If m.Bars.IsIntraday Then temp = Left(temp, 5) & "  " & DateFormat(m.dPatternDateTo, NO_DATE, HH_MM)
            temp = "Pattern length: " & Trim(Str(ptrn_len)) & "  (" & temp & ")"
            If m.nBestCorr > 0 Then
                frmParent.txtCorrPercentPFP.Text = Trim(Str(m.nBestCorr))
            Else
                frmParent.txtCorrPercentPFP.Text = Trim(Str(m.nPercentCorr))
            End If
        Else
            temp = "Pattern length: N/A"
            frmParent.txtCorrPercentPFP.Text = Trim(Str(m.nLowestCorr))
        End If
        
        frmParent.lblPatternLen.Caption = temp
    End If

    fgHits.Redraw = flexRDBuffered

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.PopulateGrid"

End Sub

Private Function ApplyNewFilter(Chart As cChart, ByVal iLen&, ByVal nMinCorr&) As Long
On Error GoTo ErrSection:

    Dim i&, k&, n&, iFrom&, iTo&, strLogFile$
    Dim aSorted As cGdArray, aMethodWeighting As New cGdArray

    ValidateBeforeRun

    ' Fill out weightings for each type of correlation
    aMethodWeighting.Create eGDARRAY_Doubles, 0, 0
    
    aMethodWeighting(eMethod_IndDiff) = MethodWeightGet(eMethod_IndDiff)    ' overall ratio for all the "difference" indicators
    aMethodWeighting(eMethod_Corr) = MethodWeightGet(eMethod_Corr)
    aMethodWeighting(eMethod_Shape) = MethodWeightGet(eMethod_Shape)
    aMethodWeighting(eMehod_Directional) = MethodWeightGet(eMehod_Directional)
    aMethodWeighting(eMethod_Signs) = MethodWeightGet(eMethod_Signs)
    aMethodWeighting(eMethod_Peaks) = MethodWeightGet(eMethod_Peaks)
    
    iTo = m.Bars.FindDateTime(m.dPatternDateTo, True)
    iFrom = iTo - iLen + 1
    
    If Not m.bOptimize Then ' don't waste time leaving log files while optimizing
        strLogFile = App.Path & "\pfp.chk"
    End If
        
    k = PFP_IndicatorMatches(aMethodWeighting.ArrayHandle, m.Bars.ArrayHandle(eBARS_DateTime), m.aIndArrays.ArrayHandle, _
            m.aPaneIDs.ArrayHandle, m.aIndFlags.ArrayHandle, iTo, iLen, nMinCorr, m.tMatches.TableHandle, strLogFile)
    If k > 0 Then
        k = m.tMatches.NumRecords
    Else
        k = 0
        m.tMatches.NumRecords = 0
    End If
    
    ' put into the pfp_hits arrays sorted by correlation (necessary when optimizing)
    Set aSorted = m.tMatches.CreateSortedIndex(1, eGdSort_Descending)
    ReDim m.pfp_hits(k) As Long
    ReDim m.pfp_corr(k) As Double
    For i = 0 To k - 1
        n = aSorted(i)
        m.pfp_hits(i + 1) = m.tMatches(0, n) - iLen + 2 ' TLB: don't know why this adjustment is needed???
        m.pfp_corr(i + 1) = m.tMatches(1, n)
    Next
    m.pfp_hits(0) = k
    m.pfp_corr(0) = k
        
ErrExit:
    ApplyNewFilter = k
    Set aSorted = Nothing
    Set aMethodWeighting = Nothing
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.ApplyNewFilter"

End Function

Private Function GetDataArrayHandle(Ind As cIndicator, ByVal iNameArrayIdx As Long) As Long
On Error GoTo ErrSection:

    Dim hHandle&, strName$
    
    If Not Ind Is Nothing Then
        If Ind.DataType = eINDIC_None Then
            'do nothing, unknown data type
        ElseIf Ind.DataType = eINDIC_BarData Then
            If iNameArrayIdx >= 0 And iNameArrayIdx < m.aIndNames.Size Then
                strName = UCase(Parse(m.aIndNames(iNameArrayIdx), ";", 1))
                Select Case strName
                    Case "OPEN"
                        hHandle = Ind.Bars.ArrayHandle(eBARS_Open)
                    Case "HIGH"
                        hHandle = Ind.Bars.ArrayHandle(eBARS_High)
                    Case "LOW"
                        hHandle = Ind.Bars.ArrayHandle(eBARS_Low)
                    Case "CLOSE"
                        hHandle = Ind.Bars.ArrayHandle(eBARS_Close)
                    Case Else
                        If Not Ind.Data Is Nothing Then hHandle = Ind.Data.ArrayHandle
                End Select
            End If
        ElseIf Not Ind.Data Is Nothing Then
            hHandle = Ind.Data.ArrayHandle
        End If
    End If
    
    GetDataArrayHandle = hHandle

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.GetDataArrayHandle"
    
End Function

Public Function MethodWeightGet(ByVal eType As ePfpMethodType, _
    Optional ByVal bReturnLastUsed As Boolean = False) As Double
On Error GoTo ErrSection:

    Dim d#

    d = 0#
    
    Select Case eType
        Case eMethod_IndDiff
            If bReturnLastUsed Then
                If m.dWeightIndDiff = 0 Then
                    d = -0.5
                Else
                    d = m.dWeightIndDiff
                End If
            ElseIf m.dWeightIndDiff > 0 Then
                d = m.dWeightIndDiff
            End If
        Case eMethod_Corr
            If bReturnLastUsed Then
                If m.dWeightCorr = 0 Then
                    d = -1
                Else
                    d = m.dWeightCorr
                End If
            ElseIf m.dWeightCorr > 0 Then
                d = m.dWeightCorr
            End If
        Case eMethod_Shape
            If bReturnLastUsed Then
                If m.dWeightShape = 0 Then
                    d = -1
                Else
                    d = m.dWeightShape
                End If
            ElseIf m.dWeightShape > 0 Then
                d = m.dWeightShape
            End If
        Case eMehod_Directional
            If bReturnLastUsed Then
                If m.dWeightDirectional = 0 Then
                    d = -1
                Else
                    d = m.dWeightDirectional
                End If
            ElseIf m.dWeightDirectional > 0 Then
                d = m.dWeightDirectional
            End If
        Case eMethod_Signs
            If bReturnLastUsed Then
                If m.dWeightSigns = 0 Then
                    d = -1
                Else
                    d = m.dWeightSigns
                End If
            ElseIf m.dWeightSigns > 0 Then
                d = m.dWeightSigns
            End If
        Case eMethod_Peaks
            If bReturnLastUsed Then
                If m.dWeightPeaks = 0 Then
                    d = -1
                Else
                    d = m.dWeightPeaks
                End If
            ElseIf m.dWeightPeaks > 0 Then
                d = m.dWeightPeaks
            End If
    End Select
    
    MethodWeightGet = d

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPatternProfit.MethodWeight.Get"

End Function

Public Sub MethodWeightSet(ByVal eType As ePfpMethodType, ByVal d#)
On Error GoTo ErrSection:

    If d = 0# Then
        Select Case eType
            Case eMethod_IndDiff
                If m.dWeightIndDiff > 0 Then m.dWeightIndDiff = -m.dWeightIndDiff
            Case eMethod_Corr
                If m.dWeightCorr > 0 Then m.dWeightCorr = -m.dWeightCorr
            Case eMethod_Shape
                If m.dWeightShape > 0 Then m.dWeightShape = -m.dWeightShape
            Case eMehod_Directional
                If m.dWeightDirectional > 0 Then m.dWeightDirectional = -m.dWeightDirectional
            Case eMethod_Signs
                If m.dWeightSigns > 0 Then m.dWeightSigns = -m.dWeightSigns
            Case eMethod_Peaks
                If m.dWeightPeaks > 0 Then m.dWeightPeaks = -m.dWeightPeaks
        End Select
    Else
        Select Case eType
            Case eMethod_IndDiff
                m.dWeightIndDiff = d
            Case eMethod_Corr
                m.dWeightCorr = d
            Case eMethod_Shape
                m.dWeightShape = d
            Case eMehod_Directional
                m.dWeightDirectional = d
            Case eMethod_Signs
                m.dWeightSigns = d
            Case eMethod_Peaks
                m.dWeightPeaks = d
        End Select
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.MethodWeight.Let"

End Sub

Private Sub ValidateBeforeRun()
On Error GoTo ErrSection:
    
    'double check values before sending them to DLL
    'values may have gotten saved wrong or INI file got corrupted or something ...
    
    If Abs(m.dWeightIndDiff) > 1 Or m.dWeightIndDiff = 0 Then
        If m.dWeightIndDiff <= 0 Then
            m.dWeightIndDiff = -0.5
        Else
            m.dWeightIndDiff = 0.5
        End If
    End If
    
    If m.dWeightCorr < 0 And m.dWeightShape < 0 And m.dWeightDirectional < 0 Then
        'reset to standard values
        m.dWeightCorr = -1
        m.dWeightShape = 1
        m.dWeightDirectional = 1
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPatternProfit.ValidateBeforeRun"

End Sub

Public Property Get WeightIndDiff() As Double
    WeightIndDiff = m.dWeightIndDiff
End Property

Public Property Get WeightCorr() As Double
    WeightCorr = m.dWeightCorr
End Property

Public Property Get WeightShape() As Double
    WeightShape = m.dWeightShape
End Property

Public Property Get WeightDirection() As Double
    WeightDirection = m.dWeightDirectional
End Property

Public Property Get WeightSigns() As Double
    WeightSigns = m.dWeightSigns
End Property

Public Property Get WeightPeaks() As Double
    WeightPeaks = m.dWeightPeaks
End Property

Public Property Get ShowSettings() As Boolean
    If m.nShowSettings = 1 Then
        ShowSettings = True
        m.nShowSettings = -1
    End If
End Property

Public Property Get Heatmap() As Long
On Error GoTo ErrSection
    
    Dim i&
    
    i = m.nHeatmap
    If m.nHeatmap = 1 Then
        If m.dStandardDev <> 1 Then i = vbGrayed
    End If

    Heatmap = i

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cPatternProfit.Heatmap.Get"

End Property

Public Property Let Heatmap(ByVal nState As Long)
On Error GoTo ErrSection

    If nState <> m.nHeatmap Then
        m.nHeatmap = nState
        If nState = 1 Then m.dStandardDev = 1
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cPatternProfit.Heatmap.Let"

End Property

Public Property Get HeatmapStyle() As Long
    HeatmapStyle = m.nHeatmapStyle
End Property
