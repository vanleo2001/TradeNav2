VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cIndicators"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const kHawkeyeLevelsColors = "16777215;255;16776960;32768;8388736;255;65535;16777215;255"

Public Enum eIndicatorDataType
    eINDIC_BarData = -1
    eINDIC_Array = 0
    eINDIC_Constant = 1
    eINDIC_BooleanArray = 2
    eINDIC_ProfileBars = 3
    eINDIC_DrawCommands = 4
    eINDIC_ProfileVolume = 5
    eINDIC_None = 999
End Enum

Public Enum eIndicatorDisplayType
    ' < 0: relevant only for price bars
    eINDIC_BollingerBar = -8
    eINDIC_Renko = -7
    eINDIC_Kagi = -6
    eINDIC_PNF = -5                 'point & figure
    eINDIC_OHLC = -4
    eINDIC_HLC = -3
    eINDIC_HL = -2
    eINDIC_Candlestick = -1
    ' >= 0: for indicators (or price bars)
    eINDIC_Line = 0
    eINDIC_Histogram = 1
    eINDIC_Area = 2
    eINDIC_Points = 3
    eINDIC_Steps = 4
    eINDIC_Rectangles = 5
    eINDIC_Ribbon = 6
    eINDIC_HighlightBars = 10
    eINDIC_HighlightMarkers = 11
    eINDIC_HighlightBoxes = 12
    eINDIC_HighlightZones = 13
    eINDIC_ArtPyramid = 14
    eINDIC_ArtReversal = 15
    eINDIC_ValueMarkers = 16
    eINDIC_ClusterPrice = 17
    eINDIC_ClusterTime = 18
'    eINDIC_ClusterPriceLine = 19        'JM 08-05-2015: DEPRECATED
'    eINDIC_ClusterPriceNone = 20        'JM 08-05-2015: DEPRECATED
    eINDIC_Profile = 21
    eINDIC_NoStyle = 30
End Enum

Public Enum eIndicatorStyle
    eINDIC_Default = 0
    eINDIC_Thin = 1
    eINDIC_MediumThin = 2
    eINDIC_Medium = 3
    eINDIC_MediumThick = 4
    eINDIC_Thick = 5
    eINDIC_ExtraThick = 6
    eINDIC_Dash = 7
    eINDIC_Dot = 8
    eINDIC_DashDot = 9
    eINDIC_Auto = 7
End Enum

Public Enum eIndicatorLabelMode
    eINDIC_DefaultSetting = 0                                       'Text in chart editor
    eINDIC_scale = 1            'scale:value    label:desc          ("Value in axis")
    eINDIC_Label = 2            'scale:nothing  label:desc,value    ("Value in label")
    eINDIC_LabelScale = 3       'scale:value    label:desc,value    ("Value in label/axis")
    eINDIC_NoValue = 4          'scale:nothing  label:desc          ("No values")
    eINDIC_NoLabel = 5          'scale:value    label:nothing       ("No labels")
    eINDIC_Nothing = 6          'scale:nothing  label:nothing       ("No labels or values")
    eINDIC_OnlyValue = 7        'scale:nothing  label:value         ("Only values")
End Enum

Public Enum eHighlightMode
    eHighlight_None = 0
    eHighLight_Bars = 1
    eHighlight_MarkersAbove = 2
    eHighlight_MarkersBelow = 3
End Enum

Private Enum eVwapBreak
    eVwapBreak1_NA = -2                'not a VWAP chart, NA = not applicable
    eVwapBreak2_Pivot = -1             'faster to break VWAP pivots based on data instead of date boundary
    eVwapBreak3_Daily = 0
    eVwapBreak4_Weekly
    eVwapBreak5_Monthly
    eVwapBreak6_Quarterly
    eVwapBreak7_Yearly
End Enum

Private Type mPassed
    nColor As Long
End Type

Private Type mPrivate
    'Passed As mPassed
    Chart As cChart         'chart object needed when removing alerts
    ' general function attributes
    strName As String
    eDataType As eIndicatorDataType
    strChartLabel As String
    strChartValue As String
    strExpression As String '(as if typed into an editor)
    strCodedText As String
    dCodedTextDate As Double ' date coded text was last built
    
    ' the following will only be valid if the function is
    ' from the database (i.e. NOT a custom expression)
    nFunctionID As Long '(= 0 for custom expressions)
    strCodedName As String
    Inputs As cInputs
    nShiftBars As Long
    nInputsFrame As Long
    
    ' display attributes
    bHideWick As Boolean ' TLB 6/29/2016: option Peter Brandt wants to hide the Candlestick wick (i.e. display body only)
    bFlip As Boolean
    bDisplay As Boolean
    eDisplayType As eIndicatorDisplayType
    eStyle As eIndicatorStyle
    eLabelMode As eIndicatorLabelMode
    nColor As Long
    nUpColor As Long
    nDownColor As Long
    nBollingerColor As Long
    nUpDownColorFlag As Long            '1=use up/down colors for HL, HLC, OHLC bars
    nUpDownColorSpecial As Long         '1=use up/down colors for candles, kagi, renko & other special types display
    bOverlayed As Boolean
    bColorNextBar As Boolean
    bBadTicksInd As Boolean
    eMarker As eStockImage
    eMarkerDir As eImageDir
    nMarkerLoc As Long          '0=above indicator, 1=below indicator
    nMarkerFill As Long         '0=hollow, 1=solid
    nMarkerSize As Long         '1=small, 2=medium, 3=large
    strMarkerAscii As String    'ascii char for custom markers, ignored unless marker type is eCNI_Ascii
    
    'highlightbars/colored indicators cross reference & flags
    nColorPriceInd As Long  '0=do not color price bars, 1=color price bars in addition to coloring this indicator
    nIndToColor As Long     'index of indicator to color (ignored by indicators not of type boolean array)
    aBoolIndRef As cGdArray 'holds indexes of all boolean arrays used to color this indicator (ignored by indicators of type boolean array)
    nCanBeHighlight As Long
    bMarkerPrompt As Boolean
    
    'variables for drawing squares around highlight bars
    nBoxBarsRight As Long
    nBoxBarsLeft As Long
    nBoxFillColor As Long
    nBoxPenStyle As Long
    nBoxFillStyle As Long       '0=hollow, 1=filled
    nBoxAnnotID As Long
    nShowInAllPanes As Long     'flag for highlight box's height
        
    'variables for coloring a histogram or area chart with different colors above/below a y-value
    dColorSeperator As Double   'y-value
    dBaseLineY As Double
    dUserBaseLineY As Double
    nColorBelow As Long
    
    'variables for auto trend lines
    nShowTrendHistory As Long
    nExtendTrend As Long
    
    'variables for ART Pyramid
    strFontName As String
    nFontSize As Long
    nFontBold As Long
    nFontItalic As Long
    
    'variables for alerts
    bAlert As Boolean
    bAlertAdded As Boolean      '0=not added to g.Alerts, 1=added to g.Alerts
    strAlert As String
    strCondBuilder As String
    oAlert As cAlert
    
    'variables for indicator grouping
    strMyKey As String
    strGroupKey As String           'key identifying a group of indicators
    bUpdateGroupInProg As Boolean   'to prevent recursion when updating group of indicators
    aGroupKeys As cGdArray          'to move group
    
    'variables for price/time clusters
    tbClusters As cGdTable              'table of values for price cluster
    
    aClusterSwingWeights As cGdArray    'should be array[4] where [0] item is ignored by routine calculating price cluster values
    tbClusterRatios As cGdTable
    
    dClusterEndDate As Double           'stores bar end date for fib cluster
    nClusterBarsBack As Long            'stores # of bars from cluster end date
    nClusterProximity As Long
    nClusterMaxBars As Long
    nClusterSwingLabel As Long          '0/1: label swing point indicator with L,I,S
    nClusterSwingColorL As Long         'colors for labelling swing points
    nClusterSwingColorM As Long
    nClusterSwingColorS As Long
    
    bResistanceOrL2H As Boolean         'resistance for price, L2H for time
    bSupportOrH2L As Boolean            'support for price, H2L for time
    bRetraceOrL2L As Boolean            'retracement for price, L2L for time
    bExpanOrH2H As Boolean              'expansion for price, H2H for time
    bClusterABCs As Boolean
    
    nIndSeasonal As Long                'not zero = seasonal chart
    nTicksPerRow As Long                'for profile style display
    nSessionsPerProfile As Long         'for multi-sessions profile style display
    strProfilePeriodicity As String
    nStyleTPO As Long
    'show/hide flags
    nShowProfileTPO As Long
    nShowProfileVol As Long
    nShowProfileTpoPOC As Long
    nShowProfileVolPOC As Long
    nShowProfileVolVA As Long
    nShowProfileTpoVA As Long
    nShowProfileOC As Long
    nProfileOnRight As Long
    'colors
    nColorProfileFrom As Long
    nColorProfileTo As Long
    nColorProfileUp As Long
    nColorProfileDown As Long
    nColorProfileVol As Long
    nColorProfileTpoPOC As Long
    nColorProfileVolPOC As Long
    nColorProfileTpoVA As Long
    nColorProfileVolVA As Long
    nColorProfileOC As Long
    'parms
    dPercentTpoVA As Double
    dPercentVolVA As Double
    'color scheme
    eColorProfileScheme As MktProfile_Char_ColorScheme
    
    ' data stored here
    aData As cGdArray
    Bars As cGdBars
    BarsProfile As cGdBars
    BarsProfileAlign As cGdBars
    BarsProfileTicks As cGdBars
    idxPriceBars As Long
    
    'grapheng.dll structure
    geIndStruct As chart_indicator
    geAdded As Boolean
    
    aSymbolReplacements As cGdArray
    
    ' TLB: so doesn't keep trying to regenerate coded text which is erroring
    bCodedTextError As Boolean
    strErrorText As String
    
    strHawkeyeColors As String
    strLinkedParms As String        'save linked input names only - eg &BarsPeriod|&ShowInitialBars|&LinkedParmN
        
    bConfirmAlert As Boolean
    bSyncOverlayFlip As Boolean     'this is for chart editor purposes -6519 (do not save to template)
End Type
Private m As mPrivate

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.aSymbolReplacements = New cGdArray

    Set m.aData = New cGdArray
    Set m.aBoolIndRef = New cGdArray
    m.aData.Create eGDARRAY_Doubles, 0
    SetDefaults
        
    Set m.Bars = Nothing
    
    If Not m.BarsProfile Is Nothing Then
        ProfileBarsFree m.BarsProfile           'theoretically should not get here ??
        Set m.BarsProfile = Nothing
    End If
    
    'initialize grapheng.dll structure
    m.geIndStruct.gdhData1 = gdCreateArray(eGDARRAY_Doubles)
    m.geIndStruct.gdhData2 = gdCreateArray(eGDARRAY_Doubles)
    m.geIndStruct.gdhData3 = gdCreateArray(eGDARRAY_Doubles)
    m.geIndStruct.gdhData4 = gdCreateArray(eGDARRAY_Doubles)
    m.geIndStruct.gdhYScaleVal = gdCreateArray(eGDARRAY_Doubles)
    
    m.geIndStruct.glhPenColor = gdCreateArray(eGDARRAY_Longs)
    m.geIndStruct.glhPenStyle = gdCreateArray(eGDARRAY_Longs)
    m.geIndStruct.glhPenWidth = gdCreateArray(eGDARRAY_Longs)
    
    m.geIndStruct.labelMode = eINDIC_DefaultSetting
    m.geIndStruct.glhImageLoc = gdCreateArray(eGDARRAY_Longs)
    m.geIndStruct.glhImageDir = gdCreateArray(eGDARRAY_Longs)
    m.geIndStruct.glhImageFill = gdCreateArray(eGDARRAY_Longs)
    m.geIndStruct.glhImageColor = gdCreateArray(eGDARRAY_Longs)
    m.geIndStruct.glhImageSize = gdCreateArray(eGDARRAY_Longs)
    m.geIndStruct.gdshImage = gdCreateArray(eGDARRAY_Strings)
    
    m.geIndStruct.glhFlags = gdCreateArray(eGDARRAY_Longs)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.Intialize", eGDRaiseError_Raise
    
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.aSymbolReplacements = Nothing
    Set m.Inputs = Nothing
    Set m.aData = Nothing
    Set m.aBoolIndRef = Nothing
    Set m.Chart = Nothing
    
    Set m.Bars = Nothing
    ProfileBarsFree m.BarsProfile
    Set m.BarsProfile = Nothing
    
    AlertsCollection 0, True
    
    'destroy grapheng.dll structure
    gdDestroyArray m.geIndStruct.gdhData1
    gdDestroyArray m.geIndStruct.gdhData2
    gdDestroyArray m.geIndStruct.gdhData3
    gdDestroyArray m.geIndStruct.gdhData4
    gdDestroyArray m.geIndStruct.gdhYScaleVal
    gdDestroyArray m.geIndStruct.gdshImage
    
    gdDestroyArray m.geIndStruct.glhImageLoc
    gdDestroyArray m.geIndStruct.glhImageDir
    gdDestroyArray m.geIndStruct.glhImageFill
    gdDestroyArray m.geIndStruct.glhImageColor
    gdDestroyArray m.geIndStruct.glhImageSize
    gdDestroyArray m.geIndStruct.glhPenColor
    gdDestroyArray m.geIndStruct.glhPenStyle
    gdDestroyArray m.geIndStruct.glhPenWidth
    
    gdDestroyArray m.geIndStruct.glhFlags

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.Terminate", eGDRaiseError_Raise
    
End Sub

Public Sub SetDefaults()
On Error GoTo ErrSection:

    ' defaults
    m.nUpDownColorSpecial = -1
    m.nCanBeHighlight = 1
    Display = True
    DataType = eINDIC_Array
    DisplayType = eINDIC_HighlightMarkers
    Style = eINDIC_Default
    Color = vbBlue
    m.nUpColor = RGB(0, 192, 0)
    m.nDownColor = vbRed
    m.nBollingerColor = -1
    m.nShowTrendHistory = 1
    
    'use same defaults as Trade Profile form (frmMarketProfile.frm)
    m.nColorProfileUp = vbGreen
    m.nColorProfileDown = vbRed
    m.nColorProfileFrom = vbYellow
    m.nColorProfileTo = vbBlack
    
    m.eMarker = eCNI_Arrow
    m.eMarkerDir = eCNI_South
    m.nMarkerLoc = 0
    m.strMarkerAscii = "A"
    
    m.geIndStruct.trueRangeColor = RGB(192, 192, 192) ' default to light gray
    m.geIndStruct.TrueRangeFlag = kNullData
    
    m.nBoxFillColor = vbYellow
    m.nColorBelow = -1
    m.dColorSeperator = 0
    m.dUserBaseLineY = kNullData        'to indicate user has not set baseline
    
    m.strFontName = "Arial"
    m.nFontSize = 8
    m.nFontBold = 0
    
    m.bFlip = False
    m.bHideWick = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.SetDefaults", eGDRaiseError_Raise
    
End Sub

Private Sub GetDatabaseFunc()
On Error GoTo ErrSection:

    Dim i&
    
    If m.nFunctionID = 0 Then
        m.strCodedName = ""
        Set m.Inputs = Nothing
    ElseIf g.Functions.Found(CStr(m.nFunctionID)) Then
        With g.Functions.Item(CStr(m.nFunctionID))
            ' get name
            m.strName = .FunctionName
            m.strCodedName = .CodedName
            'spread function has hidden auto-mutliplier we don't want displayed - 3939
            ' get inputs
            If g.Functions.Item(CStr(m.nFunctionID)).FunctionCategoryID <> 27 Then
                Set m.Inputs = New cInputs
                With .Inputs
                    For i = 1 To .Count
                        With .Item(i)
                            m.Inputs.Add .RuleName, .ParmSeq, .ParmName, _
                                .ParmDesc, .ParmID, .Value, .IfOptimize, _
                                .OptFromValue, .OptToValue, .OptStepValue, _
                                .OptListID, .RuleID, .ParmTypeID, .DefaultValue, _
                                .Required, .FromValue, .ToValue, .ListID, _
                                .FillPre, .FillPost
                        End With
                        m.Inputs.Item(i).Value = m.Inputs.Item(i).DefaultValue
                        m.Inputs.Item(i).IfOptimize = False
                    Next
                End With
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.GetDatabaseFunc", eGDRaiseError_Raise
    
End Sub

Public Property Get FunctionID() As Long
On Error GoTo ErrSection:
    
    FunctionID = m.nFunctionID

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.FunctionID.Get", eGDRaiseError_Raise
    
End Property

'Public Property Let FunctionID(ByVal nFunctionID As Long)
'    m.nFunctionID = nFunctionID
'    GetDatabaseFunc
'End Property

Public Property Get CodedName() As String
On Error GoTo ErrSection:

    CodedName = m.strCodedName

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.CodedName.Get", eGDRaiseError_Raise
    
End Property

Public Property Let CodedName(ByVal strCodedName As String)
On Error GoTo ErrSection:

    strCodedName = Trim(strCodedName)
    
    ' TLB 9/20/2012: rename "LW Sentiment" and "LW Sentiment Index" indicators
    If UCase(Left(strCodedName, 11)) = "LWSENTIMENT" Then
        strCodedName = "TNConsensus" & Mid(strCodedName, 12)
    End If
    
    If Len(strCodedName) > 0 And Left(strCodedName, 1) <> "*" Then
        m.nFunctionID = GetFunctionIDFromCodedName(strCodedName)
    Else
        m.nFunctionID = 0
    End If
    GetDatabaseFunc
    
    ' TLB 7/25/2006: if function isn't accessable (enablements?), we
    ' should still keep the existing CodedName so will work later when
    ' the module does become enabled (but prepend with "*" so flagged)
    If Len(m.strCodedName) = 0 And Len(strCodedName) > 0 Then
        If Left(strCodedName, 1) = "*" Then
            m.strCodedName = strCodedName
        Else
            m.strCodedName = "*" & strCodedName
        End If
    ElseIf m.strCodedName = "TAS_MarketMap" Then
        m.eDataType = eINDIC_DrawCommands
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.CodedName.Let", eGDRaiseError_Raise
    
End Property

Public Property Get Name() As String
    
    ' TLB 7/21/2014: special case to strip off a right parenthesis if no left parenthesis exists
    ' -- to allow trick for custom assigning a name like "Blah )" in the chart editor if you wish
    ' to not display the argument list on either the chart label or the chart name
    Dim s$
    s = m.strName
    If Right(s, 1) = ")" Then
        If InStr(s, "(") = 0 Then
            s = Trim(Left(s, Len(s) - 1))
        End If
    End If
    Name = s
    
End Property

Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get Display() As Boolean
    Display = m.bDisplay
End Property

Public Property Let Display(ByVal bDisplay As Boolean)
    
    m.bDisplay = bDisplay
    
    If Len(m.strGroupKey) > 0 Then
        UpdateGroup m.strGroupKey, bDisplay
    End If
    
End Property

Public Property Get Inputs() As cInputs
    Set Inputs = m.Inputs
End Property

Public Property Get Expression() As String
    Expression = m.strExpression
End Property

Public Property Let Expression(ByVal strExpression As String)
    m.strExpression = strExpression
End Property

Public Property Get ShiftBars() As Long
    ShiftBars = m.nShiftBars
End Property

Public Property Let ShiftBars(ByVal nShiftBars As Long)
    m.nShiftBars = nShiftBars
End Property

Public Property Get InputsFrame() As Long
    InputsFrame = m.nInputsFrame
End Property

Public Property Let InputsFrame(ByVal nInputsFrame As Long)
    m.nInputsFrame = nInputsFrame
End Property

Public Property Get DataType() As eIndicatorDataType
    DataType = m.eDataType
End Property

Public Property Let DataType(ByVal eDataType As eIndicatorDataType)
On Error GoTo ErrSection:

    m.eDataType = eDataType
    If m.eDataType = eINDIC_BarData Then
        Set m.Bars = New cGdBars
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.DataType.Let", eGDRaiseError_Raise
    
End Property

Public Property Get DisplayType() As eIndicatorDisplayType
    DisplayType = m.eDisplayType
End Property

Public Property Let DisplayType(ByVal eDisplayType As eIndicatorDisplayType)
    m.eDisplayType = eDisplayType
End Property

Public Property Get Style() As eIndicatorStyle
    Style = m.eStyle
End Property

Public Property Let Style(ByVal eStyle As eIndicatorStyle)
    m.eStyle = eStyle
    If m.geIndStruct.indicatorType <> 10 Then geSetPenStyle
End Property

Public Property Get Color() As Long
    If m.eDisplayType = eINDIC_BollingerBar Then
        If m.nBollingerColor < 0 Then
            Color = RGB(0, 0, 192)
        Else
            Color = m.nBollingerColor
        End If
    Else
        Color = m.nColor
    End If
End Property

Public Property Let Color(ByVal nColor As Long)
    If m.eDisplayType = eINDIC_BollingerBar Then
        m.nBollingerColor = nColor
    Else
        m.nColor = nColor
    End If
End Property

Public Property Get Overlayed() As Boolean
    Overlayed = m.bOverlayed
End Property

Public Property Let Overlayed(ByVal bOverlayed As Boolean)
    m.bOverlayed = bOverlayed
End Property

Public Property Get ColorNextBar() As Boolean
    ColorNextBar = m.bColorNextBar
End Property

Public Property Let ColorNextBar(ByVal bColorNextBar As Boolean)
    m.bColorNextBar = bColorNextBar
End Property

Public Property Get Data() As cGdArray
    Set Data = m.aData
End Property

Public Property Get Bars() As cGdBars
    If m.Bars Is Nothing Then
        Set m.Bars = New cGdBars
    End If
    Set Bars = m.Bars
End Property

Public Property Set Bars(Bars As cGdBars)
    Set m.Bars = Bars
End Property

Public Property Get CodedText() As String
    CodedText = m.strCodedText
End Property

Public Property Let CodedText(ByVal strText As String)
    m.strCodedText = strText
End Property

Public Property Get CodedTextDate() As Double
    CodedTextDate = m.dCodedTextDate
End Property

'Make constants (horz lines) act like has 1 parm
Public Property Get ParmCount() As Long
On Error GoTo ErrSection:

    If m.eDataType = eINDIC_Constant Then
        ParmCount = 1
    ElseIf m.Inputs Is Nothing Then
        ParmCount = 0
    Else
        ParmCount = m.Inputs.Count
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ParmCount.Get", eGDRaiseError_Raise
    
End Property

Public Property Get Parm(ByVal nItem As Long) As String
On Error GoTo ErrSection:

    If nItem > ParmCount Then
        Parm = ""
    ElseIf m.eDataType = eINDIC_Constant Then
        Parm = m.strExpression
    Else
        Parm = m.Inputs.Item(nItem).Value
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.Parm.Get", eGDRaiseError_Raise
    
End Property

Public Property Let Parm(ByVal nItem As Long, ByVal strValue As String)
On Error GoTo ErrSection:

    Dim Size&
    
    If nItem <= ParmCount Then
        If m.eDataType = eINDIC_Constant Then
            m.strExpression = Trim(strValue)
            If m.geIndStruct.gdhData1 <> 0 Then
                Size = gdGetSize(m.geIndStruct.gdhData1)
                gdClear m.geIndStruct.gdhData1, True
                gdMakeConstantValue m.geIndStruct.gdhData1, ValOfText(m.strExpression), Size
            End If
        ElseIf Not m.Inputs Is Nothing Then
            m.Inputs.Item(nItem).Value = Trim(strValue)
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.Parm.Let", eGDRaiseError_Raise
    
End Property

Public Property Get ParmType(ByVal nItem As Long) As Long
On Error GoTo ErrSection:

    If nItem > ParmCount Then
        ParmType = 0
    ElseIf m.eDataType = eINDIC_Constant Then
        ParmType = 1
    Else
        ParmType = m.Inputs.Item(nItem).ParmTypeID
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.ParmType.Get", eGDRaiseError_Raise
    
End Property

Public Property Get ParmDefault(ByVal nItem As Long) As String
On Error GoTo ErrSection:

    If nItem > ParmCount Or nItem < 1 Then      'ask Tim about this
        ParmDefault = ""
    ElseIf m.eDataType = eINDIC_Constant Then
        ParmDefault = "0"
    Else
        ParmDefault = m.Inputs.Item(nItem).DefaultValue
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.ParmDefault.Get", eGDRaiseError_Raise
    
End Property

Public Property Get ParmName(ByVal nItem As Long) As String
On Error GoTo ErrSection:

    If nItem > ParmCount Then
        ParmName = ""
    ElseIf m.eDataType = eINDIC_Constant Then
        ParmName = "Value"
    Else
        ParmName = m.Inputs.Item(nItem).ParmName
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.ParmName.Get", eGDRaiseError_Raise
    
End Property

Public Sub TemplateSave(ByVal fh%)
On Error GoTo ErrSection:

    Dim i&
    Dim strExpression As String         ' Expression to save
    
    If fh = 0 Then Exit Sub
    
    ' Replace any carriage return/line feed characters in the expression before saving it...
    strExpression = m.strExpression
    strExpression = Replace(strExpression, vbCr, "<cr>")
    strExpression = Replace(strExpression, vbLf, "<lf>")
        
    Print #fh, IniString("Display", m.bDisplay)
    Print #fh, IniString("CodedName", m.strCodedName)
    Print #fh, IniString("Name", m.strName)
    Print #fh, IniString("Expression", strExpression)
    Print #fh, IniString("CodedText", m.strCodedText)
    Print #fh, IniString("CodedTextDate", m.dCodedTextDate)
    Print #fh, IniString("DataType", m.eDataType)
    Print #fh, IniString("DisplayType", m.eDisplayType)
    Print #fh, IniString("Color", m.nColor)
    Print #fh, IniString("InputsFrame", m.nInputsFrame)
    'Print #fh, IniString("NumSubsets", .nNumSubsets)
    Print #fh, IniString("Style", m.eStyle)
    'Print #fh, IniString("Width", .nWidth)
    Print #fh, IniString("ShiftBars", m.nShiftBars)
    Print #fh, IniString("Overlayed", m.bOverlayed)
    Print #fh, IniString("ChartLabel", m.strChartLabel)
    Print #fh, IniString("ColorNextBar", m.bColorNextBar)
    Print #fh, IniString("IndicatorLabelMode", m.eLabelMode)
    Print #fh, IniString("ColorPriceIndFlag", m.nColorPriceInd)
    Print #fh, IniString("MarkerImage", m.eMarker)
    Print #fh, IniString("MarkerDirection", m.eMarkerDir)
    Print #fh, IniString("MarkerFill", m.nMarkerFill)
    Print #fh, IniString("MarkerLocation", m.nMarkerLoc)
    Print #fh, IniString("MarkerAscii", m.strMarkerAscii)
    Print #fh, IniString("MarkerSize", m.nMarkerSize)
    Print #fh, IniString("TrueRangeColor", m.geIndStruct.trueRangeColor)
    Print #fh, IniString("TrueRangeFlag", m.geIndStruct.TrueRangeFlag)
    Print #fh, IniString("UpDownColorFlag", m.nUpDownColorFlag)
    Print #fh, IniString("UpDownColorSpecial2", m.nUpDownColorSpecial)
    Print #fh, IniString("UpColor", m.nUpColor)
    Print #fh, IniString("DownColor", m.nDownColor)
    Print #fh, IniString("BollingerColor", m.nBollingerColor)
    Print #fh, IniString("GroupKey", m.strGroupKey)
    Print #fh, IniString("FlipDisplay", m.bFlip)
    Print #fh, IniString("HideWick", m.bHideWick)
    
    'profile chart
    If m.eDisplayType = eINDIC_Profile Or m.eDataType = eINDIC_ProfileVolume Then
        Print #fh, IniString("TicksPerRow", m.nTicksPerRow)
        Print #fh, IniString("SessionsPerProfile", m.nSessionsPerProfile)
        
        'volume
        Print #fh, IniString("ShowProfileVol", m.nShowProfileVol)
        Print #fh, IniString("ShowVolPOC", m.nShowProfileVolPOC)
        Print #fh, IniString("ShowVolVA", m.nShowProfileVolVA)
        Print #fh, IniString("ColorVolPOC", m.nColorProfileVolPOC)
        Print #fh, IniString("ColorVolVA", m.nColorProfileVolVA)
        Print #fh, IniString("PercentVolVA", m.dPercentVolVA)
        
        'gradient & up/down colors applying to both volume & TPOs display
        Print #fh, IniString("ColorFrom", m.nColorProfileFrom)
        Print #fh, IniString("ColorTo", m.nColorProfileTo)
        Print #fh, IniString("ColorScheme", m.eColorProfileScheme)
        Print #fh, IniString("ColorTpoPOC", m.nColorProfileTpoPOC)
        Print #fh, IniString("ColorTpoVA", m.nColorProfileTpoVA)
        Print #fh, IniString("ShowProfileOC", m.nShowProfileOC)
        
        Print #fh, IniString("StyleTPO", m.nStyleTPO)
        'TPO properties used by trade profile chart but not volume profile
        If m.eDisplayType = eINDIC_Profile Then
            Print #fh, IniString("ColorProfileVol", m.nColorProfileVol)
            Print #fh, IniString("ColorUp", m.nColorProfileUp)
            Print #fh, IniString("ColorDown", m.nColorProfileDown)
            Print #fh, IniString("ColorOpenClose", m.nColorProfileOC)
            Print #fh, IniString("ShowTPO", m.nShowProfileTPO)
            Print #fh, IniString("ShowTpoPOC", m.nShowProfileTpoPOC)
            Print #fh, IniString("ShowTpoVA", m.nShowProfileTpoVA)
            Print #fh, IniString("PercentTpoVA", m.dPercentTpoVA)
        Else
            Print #fh, IniString("ProfilePeriodicity", m.strProfilePeriodicity)
            Print #fh, IniString("ProfileOnRight", m.nProfileOnRight)
        End If
        
    End If
    
    If Len(m.strAlert) > 0 Then Print #fh, IniString("AlertString", m.strAlert)
    If Len(m.strCondBuilder) > 0 Then Print #fh, IniString("CondBuilder", m.strCondBuilder)
    TemplateSaveLinkedParms fh
    
    'bicolor histogram/area charts options
    Print #fh, IniString("ColorBelow", m.nColorBelow)
    Print #fh, IniString("ColorSeperatorValue", m.dColorSeperator)
    If m.dUserBaseLineY <> kNullData Then
        Print #fh, IniString("UserBaseLineY", m.dUserBaseLineY)
    End If
    
    'write highlight bars box options if applicable
    If m.eDataType = eINDIC_BooleanArray Then
        Print #fh, IniString("BoxPenStyle", m.nBoxPenStyle)
        Print #fh, IniString("BoxFillStyle", m.nBoxFillStyle)
        Print #fh, IniString("BoxFillColor", m.nBoxFillColor)
        Print #fh, IniString("BoxBarsLeft", m.nBoxBarsLeft)
        Print #fh, IniString("BoxBarsRight", m.nBoxBarsRight)
        Print #fh, IniString("ShowInAllPanes", m.nShowInAllPanes)
    ElseIf m.eDisplayType = eINDIC_ArtPyramid Or m.eDisplayType = eINDIC_ArtReversal Then
        Print #fh, IniString("BoxFillColor", m.nBoxFillColor)       'potential pyramid or voided reversal color
        Print #fh, IniString("BoxFillStyle", m.nBoxFillStyle)       'reversal bar letters flag (stacked or side by side)
    End If
    
    'swing auto trend lines
    Print #fh, IniString("TrendHistory", m.nShowTrendHistory)
    Print #fh, IniString("ExtendTrend", m.nExtendTrend)
    
    'art pyramid
    If Len(m.strFontName) > 0 Then
        Print #fh, IniString("FontName", m.strFontName)
        Print #fh, IniString("FontSize", m.nFontSize)
        Print #fh, IniString("FontBold", m.nFontBold)
        Print #fh, IniString("FontItalic", m.nFontItalic)
    End If
    
    'colors for Hawkeye Levels
    If IsHawkeyeLevels Then
        Print #fh, IniString("HawkeyeLevelsColors", m.strHawkeyeColors)
    End If
    
    'indicator for seasonal chart
    If m.nIndSeasonal = 1 Then
        Print #fh, IniString("IndSeasonal", m.nIndSeasonal)
    End If
    
    'clusters
    If m.eDisplayType = eINDIC_ClusterPrice Or m.eDisplayType = eINDIC_ClusterTime Then
        
        Dim aFib As cGdArray
        
        If Not m.tbClusterRatios Is Nothing Then
            Set aFib = m.tbClusterRatios.FieldArray(0)
            If Not aFib Is Nothing Then
                Print #fh, IniString("ClusterFibRatios", aFib.JoinFields("|"))
            End If
            Set aFib = m.tbClusterRatios.FieldArray(1)
            If Not aFib Is Nothing Then
                Print #fh, IniString("ClusterFibWeights", aFib.JoinFields("|"))
            End If
        End If
        
        If Not m.aClusterSwingWeights Is Nothing Then
            Print #fh, IniString("ClusterSwingWeights", m.aClusterSwingWeights.JoinFields("|"))
        End If
        
        If m.eDisplayType = eINDIC_ClusterPrice Then
            Print #fh, IniString("ClusterSwingLabel", m.nClusterSwingLabel)
            Print #fh, IniString("ClusterSwingColorL", m.nClusterSwingColorL)
            Print #fh, IniString("ClusterSwingColorM", m.nClusterSwingColorM)
            Print #fh, IniString("ClusterSwingColorS", m.nClusterSwingColorS)
            Print #fh, IniString("PriceColor", m.geIndStruct.labelColor)
        End If
        
        Print #fh, IniString("ClusterProximity", m.nClusterProximity)
        Print #fh, IniString("ClusterBarsBack", m.nClusterBarsBack)
        Print #fh, IniString("ClusterEndDate", m.dClusterEndDate)
        Print #fh, IniString("ClusterMaxBars", m.nClusterMaxBars)
        
        Print #fh, IniString("ClusterResistanceOrL2H", m.bResistanceOrL2H)
        Print #fh, IniString("ClusterSupportOrH2L", m.bSupportOrH2L)
        Print #fh, IniString("ClusterRetraceOrL2L", m.bRetraceOrL2L)
        Print #fh, IniString("ClusterExpanOrH2H", m.bExpanOrH2H)
        Print #fh, IniString("ClusterABCs", m.bClusterABCs)
    End If

    If Not m.Inputs Is Nothing Then
        For i = 1 To m.Inputs.Count
            Print #fh, IniString("Parm" & Trim(Str(i)), Parm(i))
        Next
    End If
    
    Print #fh, "END="

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.TemplateSave", eGDRaiseError_Raise
    
End Sub

Public Sub TemplateLoad(ByVal fh%)
On Error GoTo ErrSection:

    Dim i&, strProp$, strValue$, dValue#, iNumParmsInFile&
    
    Dim strClusterRatios$, strClusterRatioWeights$
    Dim aTemp As New cGdArray
    
    m.nUpDownColorSpecial = -1       'default to on for candlesticks & other special display types like kagi, P&F, renko etc.
    
    If fh = 0 Then Exit Sub
    m.strName = ""
    Do While Not EOF(fh)
        Line Input #fh, strProp
        If ParseIniString(strProp, strValue, dValue) Then
            Select Case UCase(strProp)
                Case "END"
                    Exit Do
                Case "DISPLAY"
                    m.bDisplay = dValue
                Case "CODEDNAME"
                    CodedName = strValue
                Case "NAME"
                    '(fix old names)
                    If UCase(Left(strValue, 9)) = "MOV.AVG.#" Or UCase(Left(strValue, 9)) = "MVG.AVG.#" Then
                        strValue = "MovingAvg"
                    End If
                    strValue = Replace(strValue, "LW Sentiment", "TN Consensus")
                    If Len(strValue) > 0 Then
                        m.strName = strValue
                    End If
                Case "EXPRESSION"
                    strValue = Replace(strValue, "<cr>", vbCr)
                    strValue = Replace(strValue, "<lf>", vbLf)
                    strValue = Replace(strValue, "LW Sentiment", "TN Consensus")
                    m.strExpression = strValue
                Case "CODEDTEXT"
                    strValue = Replace(strValue, "LWSentiment", "TNConsensus")
                    m.strCodedText = strValue
                Case "CODEDTEXTDATE"
                    m.dCodedTextDate = dValue
                Case "CHARTLABEL"
                    strValue = Replace(strValue, "LW Sentiment", "TN Consensus")
                    m.strChartLabel = strValue
                Case "DATATYPE"
                    DataType = dValue
                Case "DISPLAYTYPE"
                    m.eDisplayType = dValue
                Case "COLOR"
                    m.nColor = dValue
                Case "UPCOLOR"
                    m.nUpColor = dValue
                Case "DOWNCOLOR"
                    m.nDownColor = dValue
                Case "INPUTSFRAME"
                    m.nInputsFrame = dValue
                Case "NUMSUBSETS"
                    '.nNumSubsets = dValue
                Case "STYLE"
                    m.eStyle = dValue
                Case "WIDTH"
                    '.nWidth = dValue
                Case "SHIFTBARS"
                    m.nShiftBars = dValue
                Case "OVERLAYED"
                    m.bOverlayed = dValue
                Case "COLORNEXTBAR"
                    m.bColorNextBar = dValue
                Case "INDICATORLABELMODE"
                    m.eLabelMode = dValue
                Case "COLORPRICEINDFLAG"
                    m.nColorPriceInd = dValue
                Case "MARKERIMAGE"
                    m.eMarker = Int(dValue)
                Case "MARKERDIRECTION"
                    m.eMarkerDir = Int(dValue)
                Case "MARKERFILL"
                    m.nMarkerFill = Int(dValue)
                Case "MARKERLOCATION"
                    m.nMarkerLoc = Int(dValue)
                Case "MARKERASCII"
                    m.strMarkerAscii = Left(strValue, 1)
                Case "MARKERSIZE"
                    m.nMarkerSize = Int(dValue)
                Case "TRUERANGECOLOR"
                    m.geIndStruct.trueRangeColor = Int(dValue)
                Case "TRUERANGEFLAG"
                    m.geIndStruct.TrueRangeFlag = dValue
                Case "PRICECOLOR"
                    m.geIndStruct.labelColor = dValue
                Case "UPDOWNCOLORFLAG"
                    m.nUpDownColorFlag = dValue
                Case "UPDOWNCOLORSPECIAL2"
                    m.nUpDownColorSpecial = dValue
                Case "BOLLINGERCOLOR"
                    m.nBollingerColor = dValue
                Case "BOXPENSTYLE"
                    m.nBoxPenStyle = dValue
                Case "BOXFILLSTYLE"
                    m.nBoxFillStyle = dValue
                Case "BOXFILLCOLOR"
                    m.nBoxFillColor = dValue
                Case "BOXBARSLEFT"
                    m.nBoxBarsLeft = dValue
                Case "BOXBARSRIGHT"
                    m.nBoxBarsRight = dValue
                Case "SHOWINALLPANES"
                    m.nShowInAllPanes = dValue
                Case "COLORBELOW"
                    m.nColorBelow = dValue
                Case "COLORSEPERATORVALUE"
                    m.dColorSeperator = dValue
                Case "TRENDHISTORY"
                    m.nShowTrendHistory = dValue
                Case "EXTENDTREND"
                    m.nExtendTrend = dValue
                Case "ALERTSTRING"
                    m.strAlert = strValue
                Case "CONDBUILDER"
                    m.strCondBuilder = strValue
                Case "BASELINEY", "USERBASELINEY"
                    m.dUserBaseLineY = dValue       'aardvark 3986
                Case "GROUPKEY"
                    m.strGroupKey = strValue
                Case "FONTNAME"
                    m.strFontName = strValue
                Case "FONTSIZE"
                    m.nFontSize = dValue
                Case "FONTBOLD"
                    m.nFontBold = dValue
                Case "FONTITALIC"
                    m.nFontItalic = dValue
                Case "HAWKEYELEVELSCOLORS"
                    m.strHawkeyeColors = strValue
                Case "LINKEDPARMS"
                    m.strLinkedParms = strValue
                Case "CLUSTERPROXIMITY"
                    m.nClusterProximity = dValue
                    If m.nClusterProximity > 99 Then
                        m.nClusterProximity = 99
                    ElseIf m.nClusterProximity < 0 Then
                        m.nClusterProximity = 0
                    End If
                Case "CLUSTERMAXBARS"
                    m.nClusterMaxBars = dValue
                Case "CLUSTERRESISTANCEORL2H"
                    m.bResistanceOrL2H = dValue
                Case "CLUSTERSUPPORTORH2L"
                    m.bSupportOrH2L = dValue
                Case "CLUSTERRETRACEORL2L"
                    m.bRetraceOrL2L = dValue
                Case "CLUSTEREXPANORH2H"
                    m.bExpanOrH2H = dValue
                Case "CLUSTERABCS"
                    m.bClusterABCs = dValue
                Case "CLUSTERSWINGLABEL"
                    m.nClusterSwingLabel = dValue
                Case "CLUSTERSWINGCOLORL"
                    m.nClusterSwingColorL = dValue
                Case "CLUSTERSWINGCOLORM"
                    m.nClusterSwingColorM = dValue
                Case "CLUSTERSWINGCOLORS"
                    m.nClusterSwingColorS = dValue
                Case "CLUSTERFIBRATIOS"
                    strClusterRatios = strValue
                Case "CLUSTERFIBWEIGHTS"
                    strClusterRatioWeights = strValue
                Case "CLUSTERSWINGWEIGHTS"
                    Set m.aClusterSwingWeights = New cGdArray
                    aTemp.SplitFields strValue, "|"
                    For i = 0 To aTemp.Size - 1
                        m.aClusterSwingWeights.Add ValOfText(aTemp(i))
                    Next
                Case "CLUSTERBARSBACK"
                    m.nClusterBarsBack = dValue
                Case "CLUSTERENDDATE"
                    m.dClusterEndDate = dValue
                Case "INDSEASONAL"
                    m.nIndSeasonal = dValue
                Case "FLIPDISPLAY"
                    m.bFlip = dValue
                Case "HIDEWICK"
                    m.bHideWick = dValue
                
                'profile charts
                Case "TICKSPERROW"
                    m.nTicksPerRow = dValue
                Case "SESSIONSPERPROFILE"
                    m.nSessionsPerProfile = dValue
                Case "STYLETPO"
                    m.nStyleTPO = dValue
                Case "SHOWTPO"
                    m.nShowProfileTPO = dValue
                Case "SHOWTPOPOC"
                    m.nShowProfileTpoPOC = dValue
                Case "SHOWTPOVA"
                    m.nShowProfileTpoVA = dValue
                Case "SHOWPROFILEOC"
                    m.nShowProfileOC = dValue
                Case "PROFILEONRIGHT"
                    m.nProfileOnRight = dValue
                Case "PROFILEPERIODICITY"
                    m.strProfilePeriodicity = strValue

                Case "COLORFROM"
                    m.nColorProfileFrom = dValue
                Case "COLORTO"
                    m.nColorProfileTo = dValue
                Case "COLORUP"
                    m.nColorProfileUp = dValue
                Case "COLORDOWN"
                    m.nColorProfileDown = dValue
                Case "COLORTPOPOC"
                    m.nColorProfileTpoPOC = dValue
                Case "COLORTPOVA"
                    m.nColorProfileTpoVA = dValue
                Case "COLOROPENCLOSE"
                    m.nColorProfileOC = dValue
                Case "COLORSCHEME"
                    m.eColorProfileScheme = dValue

                Case "PERCENTTPOVA"
                    m.dPercentTpoVA = dValue
                'volume
                Case "SHOWPROFILEVOL"
                    m.nShowProfileVol = dValue
                Case "SHOWVOLPOC"
                    m.nShowProfileVolPOC = dValue
                Case "SHOWVOLVA"
                    m.nShowProfileVolVA = dValue

                Case "COLORPROFILEVOL"
                    m.nColorProfileVol = dValue
                Case "COLORVOLPOC"
                    m.nColorProfileVolPOC = dValue
                Case "COLORVOLVA"
                    m.nColorProfileVolVA = dValue

                Case "PERCENTVOLVA"
                    m.dPercentVolVA = dValue

                Case Else
                    If UCase(Left(strProp, 4)) = "PARM" Then
                        If Not m.Inputs Is Nothing Then
                            i = Val(Mid(strProp, 5))
                            If i > iNumParmsInFile Then
                                iNumParmsInFile = i
                            End If
                            If i > 0 And i <= m.Inputs.Count Then
                                Parm(i) = strValue
                            End If
                        End If
                    'ElseIf FileExist("c:\COMMON\files.exe") Then
                    '    InfBox "i=[] ; h=Indicator Template ; Not found: " & strType
                    End If
            End Select
        End If
    Loop
    
    ' TLB 4/17/2007: implement this for backwards-compatibility of Pivot Point functions
    ' (if first 2 parms are both Bars but the first Market1 has been inserted since template was saved)
    If Not m.Inputs Is Nothing Then
        If iNumParmsInFile > 0 And iNumParmsInFile < m.Inputs.Count And m.Inputs.Count >= 2 Then
            If m.Inputs.Item(1).ParmTypeID = 5 And m.Inputs.Item(2).ParmTypeID = 5 Then
                If UCase(m.Inputs.Item(1).DefaultValue) = "MARKET1" And UCase(m.Inputs.Item(1).Value) <> "MARKET1" _
                            And UCase(m.Inputs.Item(1).Value) = UCase(m.Inputs.Item(2).DefaultValue) Then
                    For i = iNumParmsInFile To 1 Step -1
                        m.Inputs.Item(i + 1).Value = m.Inputs.Item(i).Value
                    Next
                    m.Inputs.Item(1).Value = "Market1"
                    DebugLog "Inserted Market1 parm for chart indicator: " & m.strName
                End If
            End If
        End If
    End If
    
    If Len(strClusterRatios) > 0 And Len(strClusterRatioWeights) > 0 Then
        RatioStringToTable strClusterRatios, strClusterRatioWeights
        If m.strName = "Price Clusters" Then
            'backwards compatibility check
            If m.eDisplayType = eINDIC_ClusterPrice Then
                'valid trueRangeColor values:   0=fib zone, 1=heatmap
                'valid TrueRangeFlag values:    0=no show, 1=swing & cluster, 2=cluster, 3=swing
                i = m.geIndStruct.trueRangeColor
                If i = 0 Or i = 1 Then
                    i = m.geIndStruct.TrueRangeFlag
                    If i <> 0 And i <> 1 And i <> 2 And i <> 3 Then i = 99
                Else
                    i = 99
                End If
            Else
                i = 99  'invalid or DEPRECATED display type
            End If
            If i = 99 Then
                m.eDisplayType = eINDIC_ClusterPrice
                m.geIndStruct.trueRangeColor = 0
                m.geIndStruct.TrueRangeFlag = 2
                m.geIndStruct.labelColor = m.nColor
                ClusterPriceShow = 1
                ClusterSwingShow = 0
                m.bDisplay = True
            End If
        ElseIf m.strName = "Time Clusters" Then
            If (m.geIndStruct.trueRangeColor <> 0 And m.geIndStruct.trueRangeColor <> 1) Or _
               (m.geIndStruct.TrueRangeFlag <> 0 And m.geIndStruct.TrueRangeFlag <> 1) Then
                m.geIndStruct.trueRangeColor = 0
                m.geIndStruct.TrueRangeFlag = 0
                m.bDisplay = False
            End If
        End If
    End If
    
    'set boolean flag so don't have to keep checking string length
    If Len(m.strAlert) > 0 Then
        m.bAlert = True
    Else
        m.bAlert = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.TemplateLoad", eGDRaiseError_Raise
    
End Sub

Public Function FixCodedName() As Boolean
On Error GoTo ErrSection:

    If Not IsAlert Then         'fix for chart setting giving error if an alert has same name as matching indicator
        If Not IsCustom And (DataType = eINDIC_Array Or DataType = eINDIC_BooleanArray) Then
            If Len(CodedName) = 0 Or Left(CodedName, 1) = "*" Then
                If Len(CodedName) = 0 Then
                    CodedName = StripStr(m.strName, " ")
                Else
                    CodedName = Mid(CodedName, 2)
                End If
                If Len(CodedName) > 0 And Left(CodedName, 1) <> "*" Then
                    FixCodedName = True
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.FixCodedName", eGDRaiseError_Raise
End Function

Public Sub ConvertHighlightBar(Tree As cGdTree, ByVal nIndex&)
On Error GoTo ErrSection:

'Fix for backwards compatibility bug (aardvark #297)
'Previous implementation of highlight bars applied only to price bars
'at same node level. Parent of these highlight bars is the price pane
'rather than an indicator.

    Dim i&, nParentIdx&, nPriceIndIdx&
    Dim Ind As cIndicator
    Dim strKey$
    
    nPriceIndIdx = Tree.Index("PRICE")
    strKey = m.strName
    If m.eDataType = eINDIC_BooleanArray Then
        If Tree.NodeLevel(nIndex) < 2 Then
            Tree.Move nIndex, nPriceIndIdx, eTREE_LastChild
            m.nIndToColor = nPriceIndIdx
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.ConvertHighlightBars", eGDRaiseError_Raise

End Sub

Public Property Get idxPriceBars() As Long
    idxPriceBars = m.idxPriceBars
End Property

Public Property Let idxPriceBars(ByVal idxBars As Long)
    m.idxPriceBars = idxBars
End Property

Public Property Get isPriceInd() As Long
    isPriceInd = m.geIndStruct.isPriceInd
End Property

Public Property Let isPriceInd(ByVal nIsPriceInd&)
    m.geIndStruct.isPriceInd = nIsPriceInd
End Property

Public Property Get ColorPriceIndFlag() As Long
    ColorPriceIndFlag = m.nColorPriceInd
End Property

Public Property Let ColorPriceIndFlag(ByVal nFlag&)
    m.nColorPriceInd = nFlag
End Property

Public Sub ToggleIndToColor(Tree As cGdTree, ByVal bDisplay As Boolean)
On Error GoTo ErrSection:

    Dim Ind As cIndicator

    Set Ind = Tree(m.nIndToColor)
    If Not Ind Is Nothing Then Ind.Display = bDisplay
    
    Set Ind = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.ToggleIndToColor", eGDRaiseError_Raise

End Sub

Public Sub ToggleBoolRefInd(Tree As cGdTree, ByVal bDisplay As Boolean)
On Error GoTo ErrSection:

    'this call to hide all associated reference boolean array indicators
    Dim Ind As cIndicator
    Dim nSize&, i&
    
    nSize = m.aBoolIndRef.Size
    
    For i = 0 To nSize - 1
        Set Ind = Tree(m.aBoolIndRef(i))
        If Not Ind Is Nothing Then Ind.Display = bDisplay
    Next
    
    Set Ind = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.ToggleBoolRefInd", eGDRaiseError_Raise

End Sub

Public Sub AddBoolRefInfo(ByVal nBoolIndIdx)
On Error GoTo ErrSection:

    If m.aBoolIndRef.BinarySearch(nBoolIndIdx) = False Then
        m.aBoolIndRef.Add nBoolIndIdx
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.AddBoolRefInfo", eGDRaiseError_Raise
    
End Sub

' build the label to display on the chart
Private Function BuildLabel() As String
On Error GoTo ErrSection:

    Dim iParm&, strText$, strParm$, dValue#, strChk$, iPos&, strMkt$
    Dim bParen As Boolean, bShow As Boolean, bCluster As Boolean
    
    
    If m.eDisplayType = eINDIC_ClusterPrice Then bCluster = True
    
    strText = m.strName
    If Right(strText, 2) = "#1" Or Right(strText, 2) = "#2" Then
        strText = Trim(Left(strText, Len(strText) - 2))
    End If
    If Right(strText, 1) = ")" Then
        If InStr(strText, "(") = 0 Then
            strText = Trim(Left(strText, Len(strText) - 1))
        End If
    ElseIf Not m.Inputs Is Nothing Then
        bParen = False
        For iParm = 1 To m.Inputs.Count
            With m.Inputs.Item(iParm)
                bShow = False
                strParm = Trim(.Value)
                If UCase(.DefaultValue) = "GOLD" Then
                    If UCase(strParm) <> "DEFAULT" Then
                        bShow = True
                        strParm = StripStr(strParm, Chr(34))
                    End If
                ElseIf .ParmTypeID = 5 Then ' Bars
                    If Left(.ParmName, 1) <> Chr(34) And UCase(.ParmName) <> "MARKET1" And UCase(.ParmName) <> "DEFAULT" Then
                        bShow = True
                        strParm = StripStr(strParm, Chr(34))
                    End If
                ElseIf .ParmTypeID = 6 Then ' Boolean
                    bShow = True
                    dValue = ValOfText(strParm)
                    If dValue = 0 Then
                        strParm = "F"
                    ElseIf Abs(dValue) = 1 Then
                        strParm = "T"
                    End If
                ElseIf .ParmTypeID = 2 Then
                    ' strip quotes from text parm
                    strParm = Replace(strParm, Chr(34), "")
                    bShow = True
                ElseIf IsDigit(strParm, 1) Then
                    bShow = True
                ElseIf IsDigit(strParm, 2) Then
                    If Left(strParm, 1) = "-" Then
                        bShow = True
                        
                        If bCluster Then
                            'From Tim's email for cluster 10-25-2010:
                            '   Negative strength values are not valid and simply act the same as 0 (zero).
                            'JM 10-25-2010: They are saved as negative to CHT file to preserve last-used value
                            strParm = "0"
                        End If
                        
                    ElseIf Left(strParm, 1) = "." Then
                        strParm = "0" & strParm
                        bShow = True
                    End If
                Else
                    ' special stuff for array from price bars
                    strMkt = ""
                    strChk = UCase(strParm)
                    iPos = InStr(UCase(strChk), " OF ")
                    If iPos > 0 Then
                        strMkt = Trim(Mid(strChk, iPos + 4))
                        strChk = Trim(Left(strChk, iPos))
                        If Left(strMkt, 1) = Chr(34) And Right(strMkt, 1) = Chr(34) Then
                            If InStr(strMkt, ",") = 0 Then
                                strMkt = Mid(strMkt, 2, Len(strMkt) - 2)
                            End If
                        'ElseIf Left(strMkt, 6) = "MARKET" And Val(Mid(strMkt, 7)) > 1 Then
                        '    strMkt = Parse(strOtherMarkets, vbTab, Val(Mid(strMkt, 7)))
                        Else
                            strMkt = ""
                        End If
                    End If
                    If strChk = "CLOSE" Or strChk = "OPEN" Or strChk = "HIGH" Or strChk = "LOW" Then
                        bShow = True
                        strParm = Left(strChk, 1)
                    ElseIf strChk = "WCLOSE" Then
                        bShow = True
                        strParm = "WClose"
                    ElseIf Left(strChk, 3) = "AVG" Then
                        bShow = True
                        strParm = "Avg" & Mid(strChk, 4)
                    End If
                    If Len(strMkt) > 0 Then
                        strParm = strParm & " of " & strMkt
                    End If
                End If
                If bShow Then
                    If bParen Then
                        strText = strText & "," & strParm
                    Else
                        strText = strText & " (" & strParm
                        bParen = True
                    End If
                End If
            End With
        Next
        If bParen Then
            ' take off any ending commas (e.g. if just blank text after it)
            ' and ending 0's (e.g. ",0")
            strText = Trim(strText)
            Do While Len(strText) > 0
                If Right(strText, 1) = "," Then
                    strText = Trim(Left(strText, Len(strText) - 1))
                ElseIf Right(strText, 2) = ",0" Then
                    strText = Trim(Left(strText, Len(strText) - 2))
                Else
                    Exit Do
                End If
            Loop
            strText = strText & ")"
        End If
    End If
    ' append shift
    If m.nShiftBars > 0 Then
        strText = strText & "." & CStr(m.nShiftBars)
    ElseIf m.nShiftBars < 0 Then
        strText = strText & ".(" & CStr(m.nShiftBars) & ")"
    End If
    If m.bOverlayed Then strText = "*" & strText
    
    If bCluster Then
        'last 2 parms for price cluster's swing points indicator are always 0,1 so no need to show them
        strText = Parse(strText, ",", 1) & "," & Parse(strText, ",", 2) & "," & Parse(strText, ",", 3) & ")"
    End If
    
    BuildLabel = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.BuildLabel", eGDRaiseError_Raise
    
End Function

' build the expression as if the user had typed it into an editor
Private Function BuildExpressionOld() As String
On Error GoTo ErrSection:

    Dim iParm&, strText$, strMarket$, strParm$, bParen As Boolean
    
    m.aSymbolReplacements.Clear
    
    ' if a database function, build expression
    If m.nFunctionID > 0 Then
        If g.Functions.Found(CStr(m.nFunctionID)) Then
            ' get english name of function
            strText = g.Functions.Item(CStr(m.nFunctionID)).FunctionName
            If Not m.Inputs Is Nothing Then
                ' append function parms
                bParen = False
                For iParm = 1 To m.Inputs.Count
                    With m.Inputs.Item(iParm)
                        strParm = Trim(.Value)
                        ' check for text parm
                        If .ParmTypeID = 2 Then
                            If Len(strParm) = 0 Then strParm = " " '(temp fix for TradeSense bug)
                            ' put quotes around text parm (if not already exist)
                            If Left(strParm, 1) <> Chr(34) Then
                                strParm = Chr(34) & strParm & Chr(34)
                            End If
                        ElseIf .ParmTypeID = 5 Then
                            If Left(.ParmName, 1) = Chr(34) Then
                                ' symbol/period is hard-coded in the expression
                                strParm = Trim(.ParmName)
                            ElseIf Left(strParm, 1) = Chr(34) Then
                                ' bars parm where symbol is being set in the chart settings
                                ' (must let it build coded text first, then replace symbol later)
                                m.aSymbolReplacements.Add .ParmName & vbTab & strParm
                            ElseIf Len(strParm) = 0 Or UCase(strParm) = "DEFAULT" Then
                                strParm = "Market1" '(default if no market specified)
                                m.aSymbolReplacements.Add .ParmName & vbTab & strParm
                            End If
                        End If
                        ' if first market then save for later, else append parm to text
                        If iParm = 1 And .ParmTypeID = 5 Then ' And Left(.ParmName, 1) <> Chr(34) Then
                            strMarket = strParm
                        ElseIf bParen Then
                            strText = strText & ", " & strParm
                        Else
                            If UCase(.DefaultValue) = "GOLD" And UCase(strParm) <> "GOLD" Then
                                strParm = "Close of " & strParm
                            End If
                            strText = strText & " (" & strParm
                            bParen = True
                        End If
                    End With
                Next
                If bParen Then strText = strText & ")"
            End If
            ' append shift
            If m.nShiftBars > 0 Then
                strText = "(" & strText & ")." & CStr(m.nShiftBars)
            ElseIf m.nShiftBars < 0 Then
                'strText = "(" & strText & ").(" & CStr(m.nShiftBars) & ")"
                strText = "ShiftBars(" & strText & ", " & CStr(m.nShiftBars) & ")"
            End If
            ' append market
            If Len(strMarket) > 0 And UCase(strMarket) <> "MARKET1" Then
                strText = strText & " of " & strMarket
            End If
        End If
    End If
    
    BuildExpressionOld = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.BuildExpressionOld", eGDRaiseError_Raise
    
End Function

' build the expression as if the user had typed it into an editor
Private Function BuildExpression() As String
On Error GoTo ErrSection:

    Dim iParm&, strText$, strMarket$, strParm$, bParen As Boolean
    
    m.aSymbolReplacements.Clear
    
    ' if a database function, build expression
    If m.nFunctionID > 0 Then
        If g.Functions.Found(CStr(m.nFunctionID)) Then
            ' get english name of function
            strText = g.Functions.Item(CStr(m.nFunctionID)).FunctionName
            If Not m.Inputs Is Nothing Then
                ' append function parms
                bParen = False
                For iParm = 1 To m.Inputs.Count
                    With m.Inputs.Item(iParm)
                        strParm = Trim(.Value)
                        ' check for text parm
                        If .ParmTypeID = 2 Then
                            If Len(strParm) = 0 Then strParm = " " '(temp fix for TradeSense bug)
                            ' put quotes around text parm (if not already exist)
                            If Left(strParm, 1) <> Chr(34) Then
                                strParm = Chr(34) & strParm & Chr(34)
                            End If
                        ElseIf .ParmTypeID = 5 Then
                            If Left(.ParmName, 1) = Chr(34) Then
                                ' symbol/period is hard-coded in the expression
                                strParm = Trim(.ParmName)
                            ElseIf Left(strParm, 1) = Chr(34) Then
                                ' bars parm where symbol is being set in the chart settings
                                ' (must let it build coded text first, then replace symbol later)
                                m.aSymbolReplacements.Add .ParmName & vbTab & strParm
                            ElseIf Len(strParm) = 0 Or UCase(strParm) = "DEFAULT" Then
                                strParm = "Market1" '(default if no market specified)
                                m.aSymbolReplacements.Add .ParmName & vbTab & strParm
                            End If
                        End If
                        If .ParmTypeID = 5 Then
                            ' if market is first parm then save for later, else ignore it for now
                            If iParm = 1 Then
                                strMarket = strParm
                            End If
                        ElseIf bParen Then
                            strText = strText & ", " & strParm
                        Else
                            If UCase(.DefaultValue) = "GOLD" And UCase(strParm) <> "GOLD" Then
                                strParm = "Close of " & strParm
                            End If
                            strText = strText & " (" & strParm
                            bParen = True
                        End If
                    End With
                Next
                If bParen Then strText = strText & ")"
            End If
            ' append shift
            If m.nShiftBars > 0 Then
                strText = "(" & strText & ")." & CStr(m.nShiftBars)
            ElseIf m.nShiftBars < 0 Then
                'strText = "(" & strText & ").(" & CStr(m.nShiftBars) & ")"
                strText = "ShiftBars(" & strText & ", " & CStr(m.nShiftBars) & ")"
            End If
            ' append market
            If Len(strMarket) > 0 And UCase(strMarket) <> "MARKET1" Then
                strText = strText & " of " & strMarket
            End If
        End If
    End If
    
    BuildExpression = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.BuildExpression", eGDRaiseError_Raise
    
End Function


Public Function BuildText(Chart As cChart, Optional strMarketOverride As String = "", _
        Optional ByVal bModified As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strText$, strCodedText$, i&, strText2$, strCodedText2$
    Dim Funct As cFunction
    Dim Ind As cIndicator
    
    m.aSymbolReplacements.Clear
    
    FixInputs
    
    If m.eDataType = eINDIC_ProfileVolume Then
        m.strChartLabel = m.strName
        CodedName = ""
        CodedText = ""
        m.strExpression = ""
    ElseIf m.eDisplayType = eINDIC_ClusterTime Then
        m.strChartLabel = m.strName
        CodedName = ""
        CodedText = ""
        m.strExpression = ""
        'time cluster shares swing point levels indicator with price cluster
        If Not Chart Is Nothing Then
            If Not Chart.Tree Is Nothing Then
                Set Ind = Chart.Tree(kClusterPriceKey)
                If Not Ind Is Nothing Then
                    m.strChartLabel = Replace(Ind.ChartLabel, "Price", "Time")
                End If
            End If
        End If
    Else
        ' build the label
        m.strChartLabel = BuildLabel
    End If
    
    m.bColorNextBar = False
    
    ' if a database function, build expression
    If m.nFunctionID > 0 Then
        If g.Functions.Found(CStr(m.nFunctionID)) Then
            Set Funct = g.Functions.Item(CStr(m.nFunctionID))
            strText = Trim(BuildExpression)
If IsIDE Then
    strText2 = Trim(BuildExpressionOld)
    If strText2 <> strText Then
        'bModified = True
    End If
End If
            ' if expression has changed, or if function has been changed since we
            ' built the coded text, then rebuild the coded text
            If strText <> m.strExpression Or Left(m.strCodedText, 1) <> "~" Or _
                    (m.dCodedTextDate > 0 And m.dCodedTextDate < CDbl(Funct.LastModified)) Then
                bModified = True
                m.strExpression = strText
            End If
            If InStr(UCase(Funct.CodedText), "NEXTBAR") > 0 Then
                m.bColorNextBar = True
            End If
            Set Funct = Nothing
        End If
    ElseIf m.eDataType = eINDIC_Array Or m.eDataType = eINDIC_BooleanArray Then
        ' custom functions
        If Len(m.strCodedText) = 0 And Len(m.strExpression) > 0 Then
            bModified = True
        End If
    End If
    
    If bModified Then
        ' rebuild coded text
        strText = m.strExpression
        If Len(Trim(strText)) > 0 Then
            'TLB 10/12/2004: fix for WillSpread in a custom highlight bar
            ''If Len(strMarketOverride) > 0 Then
            If Len(strMarketOverride) > 0 And UCase(strMarketOverride) <> "MARKET1" Then
                strText = "(" & strText & ") of " & strMarketOverride
            End If
            
            ' TLB 5/22/2007: if the coded text generation has errored, then flag it
            ' so won't keep trying (until the text has changed)
            If m.bCodedTextError = False Or strText <> m.strErrorText Then
                m.bCodedTextError = False
                m.strErrorText = ""
                If frmCustomFunction.ShowMe(strText, strCodedText, True, True) <= 0 Then
                    strCodedText = ""
                    m.bCodedTextError = True
                    m.strErrorText = strText
                ElseIf IsIDE And Len(strText2) > 0 And strText <> strText2 Then
                    If frmCustomFunction.ShowMe(strText2, strCodedText2, True, True) <= 0 Then
                        strCodedText2 = ""
                    End If
                    If strCodedText <> strCodedText2 Then
                        DebugLog strCodedText
                        DebugLog strCodedText2
                    End If
                End If
            End If
        End If
        ' replace any bar parms that were custom-set in chart settings
        For i = 0 To m.aSymbolReplacements.Size - 1
            strText = Parse(m.aSymbolReplacements(i), vbTab, 1)
            strText = "~07" & Format(Len(strText), "000") & strText
            strText2 = Parse(m.aSymbolReplacements(i), vbTab, 2)
            strText2 = "~07" & Format(Len(strText2), "000") & strText2
            strCodedText = Replace(strCodedText, strText, strText2)
        Next
        m.strCodedText = Trim(strCodedText)
        m.dCodedTextDate = CDbl(Now)
    End If
    
    If InStr(UCase(m.strCodedText), "NEXTBAR") > 0 Then
        m.bColorNextBar = True
    End If
       
    BuildText = bModified
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.BuildText", eGDRaiseError_Raise
    
End Function

Public Property Get ChartLabel() As String
    ChartLabel = m.strChartLabel
End Property

Public Property Let ChartLabel(ByVal strLabel As String)
    If m.eDisplayType = eINDIC_ClusterTime And UCase(strLabel) = "INVALID" Then Exit Property
    m.strChartLabel = strLabel
End Property

Public Property Get ChartValue() As String
    ChartValue = m.strChartValue
End Property

Public Property Let ChartValue(ByVal strValue As String)
    m.strChartValue = strValue
End Property

Public Function IsCustom() As Boolean
On Error GoTo ErrSection:

    If m.nInputsFrame = 2 Then IsCustom = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.IsCustom", eGDRaiseError_Raise
    
End Function

' Returns true if custom function has been modified
Public Function EditCustom(Chart As cChart) As Boolean
On Error GoTo ErrSection:

    Dim strCodedText$, strExpression$, rc&
    Dim bModified As Boolean
    
    Dim strBoolMsg$
    
    If IsCustom Then
        strExpression = m.strExpression
        strCodedText = m.strCodedText
        rc = frmCustomFunction.ShowMe(strExpression, strCodedText, , , , Chart)         '6499
        If strExpression <> m.strExpression Or strCodedText <> m.strCodedText Then
            bModified = True
        End If
        Select Case rc
        Case 0: 'Invalid
            m.strCodedText = ""
        
        Case 1: 'Numeric
            If m.eDataType = eINDIC_BooleanArray Then
                m.strName = "Custom Indicator"
                InfBox "This Custom HighlightBars does not return a boolean and will be shown as a Custom Indicator.", "i", , "Custom HighlightBars"
            End If
            m.eDataType = eINDIC_Array
            m.strExpression = strExpression
            m.strCodedText = strCodedText
        
        Case 2: 'Boolean
            If m.eDataType = eINDIC_Array Then
                InfBox "This Custom Indicator returns a boolean and will be shown as Custom HighlightBars.", "i", , "Custom Indicator"
                m.strName = "Custom HighlightBars"
            End If
            m.eDataType = eINDIC_BooleanArray
            m.strExpression = strExpression
            m.strCodedText = strCodedText

        End Select
    End If

    EditCustom = bModified

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.EditCustom", eGDRaiseError_Raise
    
End Function

Public Property Get IndLabelMode() As eIndicatorLabelMode
    IndLabelMode = m.eLabelMode
End Property

Public Property Let IndLabelMode(ByVal eMode As eIndicatorLabelMode)
    m.eLabelMode = eMode
End Property

Public Property Get geIndId() As Long
    geIndId = m.geIndStruct.indicatorId
End Property

Public Property Let geIndId(ByVal nID As Long)
    m.geIndStruct.indicatorId = nID
End Property

Public Property Get geIndpaneId() As Long
    geIndpaneId = m.geIndStruct.paneId
End Property

Public Property Let geIndpaneId(ByVal nID As Long)
    m.geIndStruct.paneId = nID
End Property

Public Property Get geDataHandle(ByVal nIdx&) As Long
On Error GoTo ErrSection:

    Select Case nIdx
        Case 1
            geDataHandle = m.geIndStruct.gdhData1
        Case 2
            geDataHandle = m.geIndStruct.gdhData2
        Case 3
            geDataHandle = m.geIndStruct.gdhData3
        Case 0
            geDataHandle = m.geIndStruct.gdhData4
        Case Else
            geDataHandle = 0
    End Select
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.geDataHandle", eGDRaiseError_Raise

End Property

Public Sub geResetData(Optional Chart As cChart)
On Error GoTo ErrSection:

    Dim rc&

    If Not Chart Is Nothing And m.geAdded Then
        rc = geRemoveItem(Chart.geChartObj, 3, m.geIndStruct)
    End If
    
    If m.eDataType = eINDIC_ProfileVolume Then
        If Not m.Chart Is Nothing Then
            If m.Chart.RedoMode = eRedo9_ReloadData Then
                ProfileBarsFree m.BarsProfile
                Set m.BarsProfile = Nothing
                Set m.BarsProfileAlign = Nothing
                
                If Not m.BarsProfileTicks Is Nothing Then
                    g.RealTime.RemoveTickBuffer m.BarsProfileTicks
                    Set m.BarsProfileTicks = Nothing
                End If
                
                gdSetNum m.geIndStruct.glhFlags, 0, 0
                gdSetNum m.geIndStruct.glhFlags, 1, 0
                gdSetNum m.geIndStruct.glhFlags, 2, 0
            End If
        End If
    End If
    
    m.geAdded = False
    m.aBoolIndRef.Clear
    m.geIndStruct.paneId = -1
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.geResetData", eGDRaiseError_Raise
    
End Sub

Public Sub geSetPenColor(ByVal nIndex&, ByVal nColor&)
On Error GoTo ErrSection:

    If m.eDataType = eINDIC_DrawCommands Then
        Exit Sub        'precautionary, theoretically should never get here
    End If
    
    ' check for the "special color" that means "just use whatever the chart background color is"
    If nColor = RGB(1, 1, 1) Then
        nColor = g.ChartGlobals.nChartBackColor
        If Not m.Chart Is Nothing Then
            If m.Chart.UseCustomColors Then
                nColor = m.Chart.ChartBackColor
            End If
        End If
    End If
    
    gdSetNum m.geIndStruct.glhPenColor, nIndex, nColor

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.geSetPenColor", eGDRaiseError_Raise
    
End Sub

Public Sub FakeDataShift(ByVal nShift&)
On Error GoTo ErrSection:
        
    gdSetShifted m.geIndStruct.gdhData1, nShift, True
    
    If m.eDataType = eINDIC_BarData Then
        gdSetShifted m.geIndStruct.gdhData2, nShift, True
        gdSetShifted m.geIndStruct.gdhData3, nShift, True
        gdSetShifted m.geIndStruct.gdhData4, nShift, True
        gdSetShifted m.geIndStruct.glhPenColor, nShift, True        '4546
    ElseIf m.eDisplayType = eINDIC_ArtPyramid Or m.eDisplayType = eINDIC_ArtReversal Then
        gdSetShifted m.geIndStruct.gdhData2, nShift, True
        gdSetShifted m.geIndStruct.gdhData3, nShift, True
    End If
    
    FakeMarkersShift nShift             '4546
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.FakeDataShift", eGDRaiseError_Raise

End Sub

Public Function GetIndToColorIdx(Tree As cGdTree, ByVal nBoolIndIdx&) As Long
On Error GoTo ErrSection:

    Dim nIdx&, nParentIdx&
    Dim Parent As Variant
    Dim Ind As cIndicator
    
    nIdx = -1
    
    If IsHawkeyeAdds Then
        nIdx = Tree.Index("PRICE")
    Else
        nParentIdx = Tree.AncestorIndex(nBoolIndIdx, -1)
        
        If Tree.NodeLevel(nParentIdx) > 0 Then  'check that parent is of type indicator
            Set Ind = Tree(nParentIdx)
            If Not Ind Is Nothing Then
                nIdx = nParentIdx
                Set Ind = Nothing
            End If
        End If
    End If
        
    GetIndToColorIdx = nIdx

ErrExit:
    Exit Function
    
ErrSection:
   RaiseError "cIndicator.GetIndToColorIdx", eGDRaiseError_Raise

End Function


Public Function geSetIndData(Chart As cChart, _
    ByVal paneId&, ByVal nID&, ByVal nFirstBar&, ByVal nLastBar&, _
    ByVal nScreenStartX&, ByVal nScreenEndX&, ByVal dExtraPriceScale#, _
    ByVal bCalcClusters As Boolean) As Long
On Error GoTo ErrSection:

    Static nPrevID As Long

    Dim i&, j&, k&, rc&, nDpoints&, nFlag&
    Dim dIndMax#, dIndMin#, dOp#, dHi#, dLo#, dCl#, dData#, dMidBar#
    Dim hOpen&, hHigh&, hLow&, hClose&, hBarsDT&, hFlags&
    Dim ParentPane As cPane
    Dim PriceInd As cIndicator, IndToColor As cIndicator
    Dim hXBar&, hDataArray&, hBoolScreenArray&
    Dim strImage$, strColor$
    Dim nSaveOHLCtype As eIndicatorDisplayType
    Dim nStartCloseLine As Long
    
    Dim ePeriodType As eBarsPeriodType
    Dim eBoolDisplayType As eIndicatorDisplayType       '3630
    
    Dim Tree As cGdTree
    Dim chartObj As Long
    Dim tResults As cGdTable
    
    Dim bUseUpdownColor As Boolean
    Dim bOverlayData As Boolean
    Dim dBaseLineY As Double        '5413
    
    Dim dFirstVisibleValue#, bPercentChangeScaling As Boolean
    Dim eLogFlag As ePANE_LogFlag
    
    Dim eVwapBreakType As eVwapBreak
    Dim d1#, d2#            'data values for VWAP indicators
    
    'assume success
    geSetIndData = 0
                    
    chartObj = Chart.geChartObj
    Set m.Chart = Chart
    Set Tree = Chart.Tree
    Set ParentPane = Tree(paneId)
       
    If ParentPane Is Nothing Then Exit Function
    If ParentPane.Display = False Then
        If m.eDisplayType = eINDIC_ClusterTime And m.geIndStruct.trueRangeColor = 1 Then
            'time cluster heatmap is turned on don't exit
        Else
            Exit Function
        End If
    End If
    If m.eDataType = eINDIC_None Then Exit Function     'Q & A - is this a valid check
    
'gdStartProfile 400

    eVwapBreakType = eVwapBreak1_NA
    
    'check for VWAP indicator
    If InStr(m.strCodedText, "VWAP") <> 0 And m.eDisplayType = eINDIC_Line Then
        
        strImage = UCase(m.strCodedText)
        
            If InStr(strImage, "PIVOT") <> 0 Then
                eVwapBreakType = eVwapBreak2_Pivot
            ElseIf InStr(strImage, "YEAR") <> 0 Then
                eVwapBreakType = eVwapBreak7_Yearly
            ElseIf InStr(strImage, "QUARTER") <> 0 Then
                eVwapBreakType = eVwapBreak6_Quarterly
            ElseIf InStr(strImage, "MONTH") <> 0 Then
                eVwapBreakType = eVwapBreak5_Monthly
            ElseIf InStr(strImage, "WEEK") <> 0 Then
                eVwapBreakType = eVwapBreak4_Weekly
            Else
                eVwapBreakType = eVwapBreak3_Daily
            End If
            
        strImage = ""
        
        If m.Bars Is Nothing Then Set m.Bars = m.Chart.Bars
    ElseIf m.eDataType = eINDIC_DrawCommands Or m.eDataType = eINDIC_ProfileVolume Then
        If m.Bars Is Nothing Then Set m.Bars = m.Chart.Bars
    End If
    
    hXBar = Chart.aXBar.ArrayHandle
    If Not m.Bars Is Nothing Then
        hBarsDT = m.Bars.ArrayHandle(eBARS_DateTime)
        hOpen = m.Bars.ArrayHandle(eBARS_Open)
        hHigh = m.Bars.ArrayHandle(eBARS_High)
        hLow = m.Bars.ArrayHandle(eBARS_Low)
        hClose = m.Bars.ArrayHandle(eBARS_Close)
        hFlags = m.Bars.ArrayHandle(eBARS_Flags)
    End If
    
    'save unique key value for handling indicator grouping
    m.strMyKey = Chart.Tree.Key(nID)
    
    If m.bDisplay = False Then
'gdStartProfile 401
        If m.eDisplayType = eINDIC_ClusterPrice And m.geIndStruct.trueRangeColor = 1 Then
            'heatmap is on, process data and send to grapheng.dll
            If m.tbClusters Is Nothing Then
                bCalcClusters = True
            ElseIf m.tbClusters.NumRecords = 0 Then
                bCalcClusters = True
            End If
        ElseIf m.eDisplayType = eINDIC_ClusterTime And m.geIndStruct.trueRangeColor = 1 Then
            'time cluster heat map is on, don't exit
            If m.aData.Size = 0 Then bCalcClusters = True
        Else
            If m.bAlert Then
                If m.oAlert Is Nothing Or Parse(m.strAlert, "|", 2) <> Chart.Symbol Then
                    InitAlert Chart             'aardvark 3662 fix
                End If
                If m.oAlert.ChartAlertId = 0 Then m.oAlert.ChartAlertId = -1
                AlertsCollection 1, True
            End If
            If m.geAdded = True Then rc = geRemoveItem(chartObj, 3, m.geIndStruct)
            If rc = 0 Then m.geAdded = False
            
            geSetIndData = rc
    'gdStopProfile 401
    'gdStopProfile 400
            Exit Function
        End If
    End If
    
'gdStartProfile 410

    If IsHawkeyeLevels Then
        If ParmCount >= 2 Then
            i = m.Chart.Bars.FindDateTime(Val(Parm(2)))
            If i >= 0 And i < m.Chart.Bars.Size Then
                dMidBar = (m.Chart.Bars(eBARS_High, i) + m.Chart.Bars(eBARS_Low, i)) / 2
            End If
        End If
    End If
    
    If bCalcClusters Then
        If m.eDisplayType = eINDIC_ClusterPrice Then
            ClusterIndPriceCalc Chart, nFirstBar
            geSetFontInfo 0
        ElseIf m.eDisplayType = eINDIC_ClusterTime Then
            ClusterIndTimeCalc Chart, nFirstBar
            'm.aData array can get re-assigned in ClusterIndTimeCalc and original hDataArray becomes invalid
            'original symptom is chart getting "hosed" after a QB refresh when time-cluster is on chart
            If m.aData Is Nothing Then GoTo ErrExit
        End If
    End If
    
    hDataArray = m.aData.ArrayHandle
    If hDataArray = 0 Then GoTo ErrExit
    
    eLogFlag = ParentPane.PaneLogFlag
    If eLogFlag = ePANE_LogFlagPercent Then bPercentChangeScaling = True
    
    ' TLB 5/16/2005: if displaying 1 tick per bar as a Close line but minutized data
    ' is on the screen, then display as OHLC bars instead
    nSaveOHLCtype = m.eDisplayType
    If Tree.Key(nID) = "PRICE" Then
        If m.Bars.Prop(eBARS_Periodicity) = ePRD_Ticks + 1 Or m.Bars.Prop(eBARS_Periodicity) = ePRD_EachTick Then
            If m.eDisplayType >= 0 Then
                ' check backwards for end of the minutized data
                For i = nScreenEndX To nScreenStartX Step -1
                    j = gdGetNum(hXBar, i) ' Chart.aXBar(i)
                    If j >= 0 Then
                        If m.Bars(eBARS_UpTicks, j) + m.Bars(eBARS_DownTicks, j) > 1 Then
                            ' then go to beginning of the next trading session
                            ' (since minutized data comes in whole sessions)
                            For k = i + 1 To nScreenEndX + 1
                                j = gdGetNum(hXBar, k) 'Chart.aXBar(k)
                                If j >= 0 Then
                                    If m.Bars.SessionDate(k) <> m.Bars.SessionDate(k - 1) Then
                                        Exit For
                                    End If
                                End If
                            Next
                            ' allow for split mode: OHLC bars up to end of minutized data,
                            ' then close line from non-minutized data forward
                            nStartCloseLine = k
                            m.eDisplayType = eINDIC_OHLC
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    nDpoints = nScreenEndX - nScreenStartX  'number of visible data points

    'reset prompt for marker's flag
    m.bMarkerPrompt = False
    'set indicator common properties
    m.geIndStruct.paneId = paneId
    m.geIndStruct.indicatorId = nID
    m.geIndStruct.FillPattern = -1
    m.geIndStruct.FillColor = -1
    m.geIndStruct.fillColor2 = -1
    If m.eDisplayType = eINDIC_BollingerBar Then
        If m.nBollingerColor < 0 Then
            m.geIndStruct.labelColor = RGB(0, 0, 192)
        Else
            m.geIndStruct.labelColor = m.nBollingerColor
        End If
    ElseIf m.eDisplayType = eINDIC_Ribbon Then
        m.geIndStruct.FillColor = m.nColor          '6224, 6395
    ElseIf m.eDisplayType = eINDIC_ClusterPrice Then
        If m.geIndStruct.TrueRangeFlag = 3 And m.eLabelMode = eINDIC_Nothing Then
            'make sure indicator label is on when both time & price clusters are set to not show
            If Tree.Item(kClusterTimeKeyPane) Is Nothing Then
                m.eLabelMode = eINDIC_Label 'theoretically should not get here, but just in case
            ElseIf Tree.Item(kClusterTimeKeyPane).Display = False Then
                m.eLabelMode = eINDIC_Label
            End If
        End If
    Else
        m.geIndStruct.labelColor = m.nColor     'for label color in y-scale area
    End If
    'If nID = m.idxPriceBars Then m.geIndStruct.isPriceInd = 1
            
    'don't show values or labels for horizontal indicators
    If m.eDisplayType = eINDIC_Line And m.eDataType = eINDIC_Constant Then
        m.geIndStruct.labelMode = eINDIC_Nothing - 1
    ElseIf m.eLabelMode = eINDIC_DefaultSetting Then
        If m.eDisplayType = eINDIC_NoStyle Then
            m.eLabelMode = eINDIC_Label
            m.geIndStruct.labelMode = 1
        Else
            m.geIndStruct.labelMode = g.ChartGlobals.eDefaultLabelMode - 1
        End If
    Else
        m.geIndStruct.labelMode = m.eLabelMode - 1
    End If

    'clear indicator's data arrays
    gdClear m.geIndStruct.gdhData1, 1
    gdClear m.geIndStruct.gdhData2, 1
    gdClear m.geIndStruct.gdhData3, 1
    gdClear m.geIndStruct.gdhData4, 1
    gdClear m.geIndStruct.gdhYScaleVal, 1
    gdClear m.geIndStruct.glhFlags, 1
    
    'clear indicator's image / highlight info arrays
    gdClear m.geIndStruct.gdshImage, 1
    gdClear m.geIndStruct.glhImageLoc, 1
    gdClear m.geIndStruct.glhImageDir, 1
    gdClear m.geIndStruct.glhImageFill, 1
    gdClear m.geIndStruct.glhImageColor, 1
    
    'clear indicator's pen info arrays
    gdClear m.geIndStruct.glhPenColor, 1
    gdClear m.geIndStruct.glhPenStyle, 1
    gdClear m.geIndStruct.glhPenWidth, 1
    
    'clear pnf flag
    m.geIndStruct.pnfX = 0
    
    'clear extra info used for kagi & fib cluster swing point indicator
    m.geIndStruct.prevBarHigh = kNullData
    m.geIndStruct.prevBarLow = kNullData
        
'gdStopProfile 410
'gdStartProfile 420

    If IsHawkeyeAdds Then
        If m.eDisplayType = eINDIC_HighlightBars Or m.eDisplayType = eINDIC_HighlightBoxes Or _
            m.eDisplayType = eINDIC_HighlightMarkers Or m.eDisplayType = eINDIC_HighlightZones Then
            m.eDataType = eINDIC_BooleanArray
        ElseIf m.eDisplayType <> eINDIC_NoStyle Then
            m.eDataType = eINDIC_Array
            m.eDisplayType = eINDIC_ValueMarkers
        End If
    ElseIf Not m.tbClusters Is Nothing Then
        For i = 0 To m.tbClusters.NumRecords - 1
            gdSetNum m.geIndStruct.gdhData2, i, m.tbClusters(0, i)
            gdSetNum m.geIndStruct.gdhData3, i, m.tbClusters(1, i)
        Next
    End If
    
    'get min/max of visible portion of indicator
    Select Case m.eDataType
        Case eINDIC_DrawCommands
            'JM 08-27-2014: display style would be set based on type specified in draw string
            '               depending on display type, can have semi-colon delimited info for
            '               (x,y) pairs or pen color, brush color etc.
            '               e.g. x1;x2;xN|y1;y2;yN|displayType|penColor|penStyle
            m.eDisplayType = eINDIC_Line
            m.geIndStruct.indicatorType = 1     'line style
            
            If Len(Parm(2)) > 0 Then
                If Val(Parm(2)) > m.Bars.Size Then GoTo ErrExit
            End If
            
        Case eINDIC_ProfileBars
            If m.Bars Is Nothing Then Set m.Bars = Chart.Bars
            If Not m.Bars Is Nothing Then
                dIndMin = gdMinValue(hLow, nFirstBar, nLastBar) - m.Bars.MinMove(gdGetNum(hBarsDT, nLastBar))
                dIndMax = gdMaxValue(hHigh, nFirstBar, nLastBar) + m.Bars.MinMove(gdGetNum(hBarsDT, nLastBar))
                
                If Chart.Form.tmrProfileLoad.Enabled Then
                    'the chart object is still requesting data, treat like an OHLC so
                    'things will run as normal until profile bars are retrieved
                    m.eDisplayType = eINDIC_OHLC
                    m.geIndStruct.indicatorType = 6
                Else
                    i = m.Bars.SessionDate(0)
                    j = Chart.LastGoodDataBar(False)
                    
                    Set m.BarsProfile = ProfileBarsGet(m.Bars.Prop(eBARS_SymbolID), i, j, False)
                    If m.BarsProfile Is Nothing Then
                        gdClear m.geIndStruct.glhFlags, 1
                        m.geIndStruct.indicatorType = 50
                    Else
                        gdSetNum m.geIndStruct.glhFlags, 0, m.Bars.BarsHandle
                        gdSetNum m.geIndStruct.glhFlags, 1, m.BarsProfile.BarsHandle
                        m.geIndStruct.indicatorType = 23
                        m.eDisplayType = eINDIC_Profile
                        geSetFontInfo 0
                    End If
                End If
                If Tree.Key(nID) = "PRICE" Then
                    m.geIndStruct.isPriceInd = 1
                End If
            End If

        Case eINDIC_ProfileVolume
            If SecurityType(m.Bars.Prop(eBARS_Symbol)) = "I" Then m.bDisplay = False
            If Not m.Bars Is Nothing And m.bDisplay Then
                dIndMin = kNullData
                dIndMax = kNullData
                
                If ParentPane.Scaling = ePANE_ScaleModeAuto Then
                    'JM 09-04-2015 - when pane scale is auto-fit to indicators
                    '       adjust min/max to include high/lows of profile data
                    If Not m.BarsProfileAlign Is Nothing Then
                        i = m.BarsProfileAlign.FindDateTime(m.Bars.SessionDate(nFirstBar))
                        j = m.BarsProfileAlign.FindDateTime(m.Bars.SessionDate(nLastBar))
                        If j >= m.BarsProfileAlign.Size Then j = m.BarsProfileAlign.Size - 1
                        
                        If j > 0 And i >= 0 And i < m.BarsProfileAlign.Size Then
                            d1 = gdMinValue(m.BarsProfileAlign.ArrayHandle(eBARS_Low), i, j)
                            d2 = gdMaxValue(m.BarsProfileAlign.ArrayHandle(eBARS_High), i, j)
                                
                            If i > 0 And (ProfileShowHide(ePCStruct_Volume_POC) > 1 Or ProfileShowHide(ePCStruct_Volume_VA) > 1) Then
                                'JM 09-04-2015: check profile off-chart left if show previous volume stats is on
                                If m.BarsProfileAlign(eBARS_Open, i - 1) <> kNullData And m.BarsProfileAlign(eBARS_Close, i - 1) <> kNullData Then
                                    dIndMin = gdMinValue(m.BarsProfileAlign.ArrayHandle(eBARS_Close), i - 1, j)
                                    dIndMax = gdMaxValue(m.BarsProfileAlign.ArrayHandle(eBARS_Open), i - 1, j)
                                ElseIf m.BarsProfileAlign(eBARS_Vol, i - 1) <> kNullData Then
                                    dIndMin = gdMinValue(m.BarsProfileAlign.ArrayHandle(eBARS_Vol), i - 1, j)
                                    dIndMax = gdMaxValue(m.BarsProfileAlign.ArrayHandle(eBARS_Vol), i - 1, j)
                                End If
                            End If
                                
                            If dIndMin = kNullData Or d1 < dIndMin Then dIndMin = d1
                            If dIndMax = kNullData Or d2 > dIndMax Then dIndMax = d2
                        End If
                    End If
                End If
                
                If dIndMin = kNullData Or dIndMax = kNullData Then
                    'JM 09-04-2015: if min/max = null because pane's scale mode != auto-fit to indicators
                    '       or volume stats not yet calculated or whatever ... use chart's visible data min/max
                    dIndMin = gdMinValue(m.Bars.ArrayHandle(eBARS_Low), nFirstBar, nLastBar)
                    dIndMax = gdMaxValue(m.Bars.ArrayHandle(eBARS_High), nFirstBar, nLastBar)
                End If
                
                If Len(m.strProfilePeriodicity) > 0 Then
                    gdSetNum m.geIndStruct.glhFlags, 0, Chart.Bars.BarsHandle
                    
                    i = GetPeriodicity(m.strProfilePeriodicity)
                    If i > ePRD_Weeks And m.Chart.Bars.IsIntraday Then
                        If GetPeriodType(i) = ePRD_Months Then
                            m.strProfilePeriodicity = "Monthly"
                        ElseIf GetPeriodType(i) = ePRD_Weeks Then
                            i = Val(Parse(m.strProfilePeriodicity, " ", 1))
                            If i > 4 Then
                                m.strProfilePeriodicity = "Weekly"
                            End If
                        Else
                            m.strProfilePeriodicity = "Daily"
                        End If
                    ElseIf Chart.Bars.Prop(eBARS_Periodicity) > i Then
                        If Chart.Bars.IsIntraday Then
                            m.strProfilePeriodicity = "Daily"
                        Else
                            m.strProfilePeriodicity = Chart.Bars.Prop(eBARS_PeriodicityStr)
                        End If
                    End If
                    If Not m.BarsProfileAlign Is Nothing Then
                        If m.BarsProfileAlign.Prop(eBARS_PeriodicityStr) <> m.strProfilePeriodicity Then
                            Set m.BarsProfileAlign = Nothing
                        End If
                    End If
                
                    If m.BarsProfileAlign Is Nothing Then
                        Set m.BarsProfileAlign = New cGdBars
                        SetBarProperties m.BarsProfileAlign, Chart.Symbol
                        
                        i = GetPeriodicity(m.strProfilePeriodicity)
                        DM_GetBars m.BarsProfileAlign, Chart.Symbol, i, Chart.Bars.SessionDate(0), Chart.Bars.SessionDate(Chart.Bars.Size - 1)
                        m.strProfilePeriodicity = m.BarsProfileAlign.Prop(eBARS_PeriodicityStr)
                        
                        'JM 09-03-2015: don't need vol info of these bars
                        'JM 09-09-2015: OI, ConVol do not work for stocks, changed to use Open, Close & Vol
                        '   init arrays to null values and let grapheng.dll save VA, POC info here
                        '   eBARS_Vol = grapheng will save price of POC to this array
                        '   eBARS_Open = grapheng will save price of VAH to this array
                        '   eBARS_Close = grapheng will save price of VAL to this array
                        For i = 0 To m.BarsProfileAlign.Size - 1
                            m.BarsProfileAlign(eBARS_Vol, i) = kNullData
                            m.BarsProfileAlign(eBARS_Open, i) = kNullData
                            m.BarsProfileAlign(eBARS_Close, i) = kNullData
                        Next
                    End If
                    
                    If m.BarsProfileAlign.Size > 0 Then
                        If m.Bars.IsIntraday Then
                            If Not m.BarsProfile Is Nothing Then
                                If m.BarsProfile.Prop(eBARS_PeriodicityStr) <> m.Bars.Prop(eBARS_PeriodicityStr) Then
                                    ProfileBarsFree m.BarsProfile
                                    Set m.BarsProfile = Nothing
                                End If
                            End If
                            If m.BarsProfile Is Nothing Then
                                'user may have changed chart's periodicity from say 30m to 60m before all profiles for 30m were loaded
                                'turn timer off to stop loading profile bars for previous periodicity
                                If Chart.Form.tmrProfileLoad.Enabled Then
                                    StatusMsg ""
                                    Chart.Form.tmrProfileLoad.Enabled = False
                                End If
                                Set m.BarsProfile = VolumeBarsGet(Chart, Me, m.BarsProfileAlign, m.BarsProfileTicks, True)
                            Else
                                Set m.BarsProfile = VolumeBarsGet(Chart, Me, m.BarsProfileAlign, m.BarsProfileTicks, False)
                            End If
                            If m.BarsProfile Is Nothing Then
                                m.bDisplay = False
                            ElseIf m.BarsProfile.Size > 0 Then
                                gdSetNum m.geIndStruct.glhFlags, 1, m.BarsProfileAlign.BarsHandle
                                gdSetNum m.geIndStruct.glhFlags, 2, m.BarsProfile.BarsHandle
                                If m.eDisplayType <> eINDIC_NoStyle Then m.geIndStruct.indicatorType = 24
                            End If
                        Else
                            gdSetNum m.geIndStruct.glhFlags, 1, m.BarsProfileAlign.BarsHandle
                            If m.eDisplayType <> eINDIC_NoStyle Then m.geIndStruct.indicatorType = 24
                        End If
                        m.geIndStruct.boxSize = m.nTicksPerRow
                        m.geIndStruct.reversal = m.nSessionsPerProfile
                    Else
                        m.bDisplay = False
                    End If
                End If
            End If

        Case eINDIC_BarData
'gdStartProfile 421
            dIndMin = gdMinValue(hLow, nFirstBar, nLastBar) - m.Bars.MinMove(gdGetNum(hBarsDT, nLastBar))
            dIndMax = gdMaxValue(hHigh, nFirstBar, nLastBar) + m.Bars.MinMove(gdGetNum(hBarsDT, nLastBar))
            
            If m.Chart.ScreenStartX > m.Chart.LastGoodDataBar(True, True) Then      '3922
                'chart has been scrolled beyond last good data bar
                If Int(dIndMin) < 0 And Int(dIndMax) < 0 Then
                    'double check in case the annot values are invalid
                    dIndMin = gdGetNum(hLow, m.Chart.LastGoodDataBar(False, False))
                    dIndMax = gdGetNum(hHigh, m.Chart.LastGoodDataBar(False, False))
                End If
            End If
            
            ePeriodType = m.Bars.Prop(eBARS_PeriodType)
            If Tree.Key(nID) = "PRICE" Then
                If ePeriodType = ePRD_EodPF Or ePeriodType = ePRD_IntPF Then
                    m.geIndStruct.indicatorType = 9
                    m.eDisplayType = eINDIC_PNF
                    If m.nUpDownColorSpecial <> 0 Then bUseUpdownColor = True
                ElseIf ePeriodType = ePRD_EodKagi Or ePeriodType = ePRD_IntKagi Then
                    m.eDisplayType = eINDIC_Kagi
                    m.geIndStruct.indicatorType = 10
                    If m.nUpDownColorSpecial <> 0 Then bUseUpdownColor = True
                ElseIf ePeriodType = ePRD_EodRenko Or ePeriodType = ePRD_IntRenko Then
                    m.eDisplayType = eINDIC_Renko
                    m.geIndStruct.indicatorType = 11
                    If m.nUpDownColorSpecial <> 0 Then bUseUpdownColor = True
                ElseIf m.eDisplayType = eINDIC_Candlestick Then
                    m.geIndStruct.indicatorType = 5
                    If m.nUpDownColorSpecial <> 0 Then bUseUpdownColor = True
                ElseIf m.eDisplayType = eINDIC_OHLC Then
                    m.geIndStruct.indicatorType = 6
                    If m.nUpDownColorFlag <> 0 Then bUseUpdownColor = True
                ElseIf m.eDisplayType = eINDIC_HLC Then
                    m.geIndStruct.indicatorType = 7
                    If m.nUpDownColorFlag <> 0 Then bUseUpdownColor = True
                ElseIf m.eDisplayType = eINDIC_HL Then
                    m.geIndStruct.indicatorType = 8
                    If m.nUpDownColorFlag <> 0 Then bUseUpdownColor = True
                ElseIf m.eDisplayType = eINDIC_BollingerBar Then
                    m.geIndStruct.indicatorType = 13
                    bUseUpdownColor = True
                ElseIf m.eDisplayType = eINDIC_Histogram Or m.eDisplayType = eINDIC_Area Then
                    'need baseline a little less than min so histogram/area won't 'float'
                    dBaseLineY = dIndMin - m.Bars.Prop(eBARS_MinMoveInTicks)
                End If
                m.geIndStruct.isPriceInd = 1
            ElseIf m.eDisplayType = eINDIC_Candlestick Then
                m.geIndStruct.indicatorType = 5
                If m.nUpDownColorSpecial <> 0 Then bUseUpdownColor = True       '5319
            ElseIf m.eDisplayType = eINDIC_OHLC Then
                m.geIndStruct.indicatorType = 6
                If m.nUpDownColorFlag <> 0 Then bUseUpdownColor = True
            ElseIf m.eDisplayType = eINDIC_HLC Then
                m.geIndStruct.indicatorType = 7
                If m.nUpDownColorFlag <> 0 Then bUseUpdownColor = True
            ElseIf m.eDisplayType = eINDIC_HL Then
                m.geIndStruct.indicatorType = 8
                If m.nUpDownColorFlag <> 0 Then bUseUpdownColor = True
            ElseIf m.eDisplayType = eINDIC_BollingerBar Then
                m.geIndStruct.indicatorType = 13
                bUseUpdownColor = True
            ElseIf m.eDisplayType < eINDIC_OHLC Then '(can't do kagi, renko, PNF for secondary symbols, so just do OHLC)
                m.eDisplayType = eINDIC_OHLC
                m.geIndStruct.indicatorType = 6
                If m.nUpDownColorFlag <> 0 Then bUseUpdownColor = True
            End If
            
            ' TLB: if a "TAS Vega" indicator is being displayed in this pane,
            ' we need to ignore the normal up/down colors for the price bars
            If bUseUpdownColor Then
                For i = nID + 1 To Tree.Count
                    If Tree.NodeLevel(i) = 0 Then Exit For
                    Set IndToColor = Tree(i)
                    If UCase(IndToColor.Name) = "TAS VEGA" Then
                        If IndToColor.Display Then
                            bUseUpdownColor = False
                        End If
                    End If
                Next
                Set IndToColor = Nothing
            End If
            
            'Get true range value
            If m.geIndStruct.TrueRangeFlag <> kNullData Then
                If Chart.FirstTrueRangeClose = kNullData Then
                    m.geIndStruct.TrueRangeFlag = -1
                Else
                    m.geIndStruct.TrueRangeFlag = Chart.FirstTrueRangeClose
                End If
            End If
            'set up/down colors
            If m.eDisplayType >= 0 Then
                m.geIndStruct.FillColor = -1
                m.geIndStruct.fillColor2 = -1
            ElseIf bUseUpdownColor Then
                m.geIndStruct.FillColor = m.nUpColor
                m.geIndStruct.fillColor2 = m.nDownColor
            End If
'gdStopProfile 421
            
        Case eINDIC_Array
'gdStartProfile 422
            If IsHawkeyeAdds Or IsHawkeyeLevels Then
                dIndMin = ParentPane.Min
                dIndMax = ParentPane.Max
            ElseIf m.eDisplayType = eINDIC_ClusterTime Then
                dIndMin = 0
                dIndMax = gdMaxValue(hDataArray, 0, gdGetSize(hDataArray))
                dBaseLineY = 0
                m.geIndStruct.TrueRangeFlag = Abs(m.bDisplay)     'flag for grapheng to not adjust pane's minimum to be exactly zero
            ElseIf IsCustom And m.aData.IsConstantValue Then           '5593
                If Tree.RelativeIndex(ParentPane.gePaneId, eTREE_FirstChild) = Me.geIndId And Tree.RelativeIndex(ParentPane.gePaneId, eTREE_LastChild) = Me.geIndId Then
                    'this is the only indicator in this pane, set min/max accordingly
                    dIndMin = gdGetNum(hDataArray, 0)
                    If dIndMin > 0 Then
                        dIndMax = dIndMin * 1.5
                        dIndMin = 0
                    ElseIf dIndMin < 0 Then
                        dIndMax = 0
                        dIndMin = dIndMin * 1.5
                    Else
                        dIndMax = 10
                        dIndMin = -10
                    End If
                Else
                    'there are other indicators in this pane, follow rule that horizontal lines should not be used for determining pane's min/max
                    dIndMin = kNullData
                    dIndMax = kNullData
                End If
            Else
                dIndMin = gdMinValue(hDataArray, nFirstBar, nLastBar)
                dIndMax = gdMaxValue(hDataArray, nFirstBar, nLastBar)
            End If
'gdStopProfile 422
                    
        'note: horz lines indicator are not to be used in determining pane's min/max
        'Case eINDIC_Constant
        '    dIndMin = Val(Parm(0))
        '    dIndMax = Val(Parm(0))
            
        Case eINDIC_BooleanArray
'gdStartProfile 423
            'boolean arrays are highlight bars for the price and/or one other indicator
'gdStartProfile 461
            Set PriceInd = Tree("PRICE")
'gdStopProfile 461
                                               
'gdStartProfile 462
            If m.bBadTicksInd = True Then
                Set IndToColor = PriceInd
            Else
                Set IndToColor = Tree(GetIndToColorIdx(Tree, nID))
            End If
            If IndToColor Is Nothing Then GoTo ErrExit
            m.bOverlayed = IndToColor.Overlayed 'TODO: this is slightly inefficient, find better method
            m.nIndToColor = IndToColor.geIndId
'gdStopProfile 462
'gdStartProfile 463
            j = 0
            If (IndToColor.DisplayType = eINDIC_BollingerBar Or IndToColor.UpDownColorFlag <> 0) _
                    And m.eDisplayType = eINDIC_HighlightBars Then
                
                If IndToColor.DisplayType = eINDIC_BollingerBar Or IndToColor.UpDownColorFlag = 1 Then
                    eBoolDisplayType = eINDIC_HighlightMarkers
                Else
                    IndToColor.UpDownColorFlag = 0
                    eBoolDisplayType = m.eDisplayType
                End If
            Else
                eBoolDisplayType = m.eDisplayType
            End If
'gdStopProfile 463

            If m.eDisplayType <> eINDIC_HighlightBars And m.eDisplayType <> eINDIC_HighlightBoxes And _
               m.eDisplayType <> eINDIC_HighlightMarkers And m.eDisplayType <> eINDIC_HighlightZones And _
               m.eDisplayType <> eINDIC_NoStyle Then
               
                If IndToColor.UpDownColorFlag = 0 Then
                    m.eDisplayType = eINDIC_HighlightBars         '5646
                Else
                    m.eDisplayType = eINDIC_NoStyle
                End If
               
            End If
            ' just for efficiency, create an array of boolean values up-front
            ' (for highlight boxes: # is the length of the box, 0 = no box)
'gdStartProfile 464
            hBoolScreenArray = gdCreateArray(eGDARRAY_Longs, nScreenEndX + 1, 0)
'gdStopProfile 464
'gdStartProfile 465
            For i = nScreenStartX - 1 To nScreenEndX
                dData = gdGetNum(hDataArray, gdGetNum(hXBar, i))
                If dData <> 0 And dData <> kNullData Then
                    ' TLB 7/8/2008: when displaying highlight boxes, only do the last box for
                    ' patterns which are simply incrementing (if the next bar = length + 1)
                    If eBoolDisplayType = eINDIC_HighlightBoxes And i < nScreenEndX Then
                        If gdGetNum(hDataArray, gdGetNum(hXBar, i) + 1) = dData + 1 Then
                            If IndToColor.Bars(eBARS_Close, gdGetNum(hXBar, i + 1)) <> kNullData Then
                                dData = 0
                            End If
                        End If
                    End If
                    If dData <> 0 Then
                        gdSetNum hBoolScreenArray, i, Abs(dData)
                    End If
                End If
            Next
'gdStopProfile 465
'gdStopProfile 423
            If eBoolDisplayType = eINDIC_HighlightMarkers Then
'gdStartProfile 424
                strImage = ""
                For i = nScreenStartX To nScreenEndX
                    'If BooleanValue(m.aData(gdGetNum(hXBar, i))) <> 0 Then
                    If gdGetNum(hBoolScreenArray, i) <> 0 Then
                        strImage = strImage + Str(m.eMarker) + ","
                    '01-27-2004: Tim does not want the next bar to be marked as is the case with highlight bars
                    'ElseIf m.bColorNextBar = True And Abs(m.aData(gdGetNum(hXBar, i - 1))) = 1 Then
                    '    strImage = strImage + Str(m.eMarker) + ","
                    Else
                        strImage = strImage + "-1,"
                    End If
                Next
                If Len(strImage) > 2 Then
                    strImage = Left(strImage, Len(strImage) - 1)    'drop trailing comma
                    IndToColor.MarkerAscii = m.strMarkerAscii
                    If m.nColorPriceInd = 0 Then    'aardvark 1075
                        If m.eMarker = eCNI_Ascii Then
                            IndToColor.geSetMarkersData strImage, m.nMarkerLoc, m.eMarkerDir, _
                                Asc(m.strMarkerAscii), m.nColor, m.nMarkerSize
                        Else
                            IndToColor.geSetMarkersData strImage, m.nMarkerLoc, m.eMarkerDir, _
                                m.nMarkerFill, m.nColor, m.nMarkerSize
                        End If
                    ElseIf m.nColorPriceInd = 1 And Not PriceInd Is Nothing Then
                        If m.geIndStruct.paneId <> PriceInd.geIndpaneId Then    'fix for aardvark 1018
                            If m.eMarker = eCNI_Ascii Then
                                PriceInd.geSetMarkersData strImage, m.nMarkerLoc, m.eMarkerDir, _
                                    Asc(m.strMarkerAscii), m.nColor, m.nMarkerSize
                            Else
                                PriceInd.geSetMarkersData strImage, m.nMarkerLoc, m.eMarkerDir, _
                                    m.nMarkerFill, m.nColor, m.nMarkerSize
                            End If
                        End If
                    End If
                End If
'gdStopProfile 424
            ElseIf m.eDisplayType = eINDIC_HighlightBoxes Or m.eDisplayType = eINDIC_HighlightZones Then
                'don't need to do anything since we are going to validate box annots anyways
            Else
                'Note: previous version does not set display type for boolean data
                If Not PriceInd Is Nothing And m.eDisplayType <> eINDIC_NoStyle Then    'original fix for 5646 caused 5660; restored code to fix 5660; moved fix for 5646 up
'gdStartProfile 425
                    Dim iUpDownColorFlag&
                    iUpDownColorFlag = PriceInd.UpDownColorFlag
                    For i = nScreenStartX To nScreenEndX
                        'If BooleanValue(m.aData(gdGetNum(hXBar, i))) <> 0 Then
                        If gdGetNum(hBoolScreenArray, i) <> 0 Then
                            IndToColor.geSetPenColor j, m.nColor
                            If m.nColorPriceInd = 1 And iUpDownColorFlag = 0 Then PriceInd.geSetPenColor j, m.nColor
                        'ElseIf m.bColorNextBar And BooleanValue(m.aData(gdGetNum(hXBar, i - 1))) <> 0 Then
                        ElseIf m.bColorNextBar Then
                            If gdGetNum(hBoolScreenArray, i - 1) <> 0 Then
                                IndToColor.geSetPenColor j, m.nColor
                                If m.nColorPriceInd = 1 And iUpDownColorFlag = 0 Then PriceInd.geSetPenColor j, m.nColor
                            End If
                        End If
                        j = j + 1
                    Next
'gdStopProfile 425
                End If
            End If
'gdStartProfile 429
            ValidateBoxAnnot Chart, IndToColor, Chart.aXdate.ArrayHandle, hXBar, nScreenStartX, nScreenEndX, hBoolScreenArray
'gdStopProfile 429
            IndToColor.AddBoolRefInfo nID
            gdDestroyArray hBoolScreenArray
            hBoolScreenArray = 0
            Set PriceInd = Nothing
            Set IndToColor = Nothing
            GoTo ErrExit 'Exit Function
        Case eINDIC_None
            GoTo ErrExit 'Exit Function
        Case Default
            GoTo ErrExit 'Exit Function
    End Select
'gdStopProfile 420
'gdStartProfile 430
        
    If ParentPane.PricePaneFlag = 1 Then
        If nPrevID > 0 And nPrevID <> Chart.SymbolID Then
            m.dUserBaseLineY = kNullData        'reset custom baseline for price data (4376)
            m.nColorBelow = -1
            m.dColorSeperator = 0
        End If
        If m.eDisplayType = eINDIC_ValueMarkers Then
            geSetPenColor 0, m.nColor
            'gdSetNum m.geIndStruct.glhPenColor, 0, m.nColor
            gdSetNum m.geIndStruct.glhImageColor, 0, m.nColor
            gdSetNum m.geIndStruct.glhImageLoc, 0, m.nMarkerLoc
            gdSetNum m.geIndStruct.glhImageSize, 0, MarkerSize
        End If
    End If
    nPrevID = Chart.SymbolID
   
    If m.dUserBaseLineY <> kNullData Then
        dBaseLineY = m.dUserBaseLineY
    ElseIf Me.CodedName = "RedLightGreenLight" Or Me.CodedName = "SeasonalPercent" Or Me.CodedName = "CycleAsPercent" Then
        dBaseLineY = 50
    ElseIf ParentPane.PricePaneFlag = 1 And (m.eDisplayType = eINDIC_Area Or m.eDisplayType = eINDIC_Histogram) Then
        If (Tree.Key(nID) = "PRICE" Or m.eDataType <> eINDIC_BarData) And dBaseLineY = 0 Then
            dBaseLineY = ParentPane.gePaneMin - 0.005
            'test case: ES daily with MovingAvg(C,18) scroll back so first bar on chart=3/26/2009
            If dBaseLineY > dIndMin Then dBaseLineY = dIndMin   'this line prevents test case from showing histogram above/below
        ElseIf m.eDataType = eINDIC_BarData Then
            'comparison symbol in price pane
            If dBaseLineY = 0 Then dBaseLineY = dIndMin - m.Bars.Prop(eBARS_MinMoveInTicks)
        End If
    ElseIf m.eDataType = eINDIC_BarData And Tree.Key(nID) <> "PRICE" Then
        'comparison symbol in non-price pane (aardvark 4026)
        If m.eDisplayType = eINDIC_Area Or m.eDisplayType = eINDIC_Histogram Then
            If dBaseLineY = 0 Then dBaseLineY = dIndMin - m.Bars.Prop(eBARS_MinMoveInTicks)
        End If
    End If
    
    m.geIndStruct.y_baseline = dBaseLineY
    
    If m.eDisplayType = eINDIC_ClusterPrice Then
        m.bOverlayed = False        'do this so * will not show next to chart's label (looks confusing)
        bOverlayData = True         'always overlay swing point levels data for display purposes since primary scale is price
    ElseIf m.eDataType = eINDIC_DrawCommands Then
        m.bOverlayed = False        '08-28-2014: for now don't allow overlay
        bOverlayData = False
    ElseIf Not bPercentChangeScaling And m.bOverlayed Then
        'data can only be overlayed for:
        'Bollinger, OHLC, HLC, HL, Line, Histogram, Area, Points, Steps & Rectangles
        If m.eDisplayType = eINDIC_BollingerBar Then
            bOverlayData = True
        ElseIf m.eDisplayType >= eINDIC_OHLC And m.eDisplayType < eINDIC_HighlightBars Then
            bOverlayData = True
        End If
    End If
    
    If bOverlayData Then
        If ParentPane.gePaneMin = kNullData Or ParentPane.gePaneMax = kNullData Then    'aardvark 1255
            geSetPaneMinMax ParentPane, Tree, dIndMax, dIndMin, nID, paneId, dExtraPriceScale
        End If
        geSetOverlayData dIndMin, dIndMax, ParentPane.gePaneMin, ParentPane.gePaneMax, nScreenStartX, nScreenEndX, eLogFlag, hXBar
        If m.geAdded = False Then rc = geAddItem(chartObj, 3, m.geIndStruct)
        If 0 = rc Then m.geAdded = True
        geSetIndData = rc
    Else
        'set data arrays
        j = 0
        'candlestick = -1, hl = -2, hlc = -3, ohlc = -4
        'PNF = -5, Kagi = -6, Renko = -7
        'line = 0, histogram = 1, area = 2, points = 3
        If m.eDisplayType < 0 Or m.eDataType = eINDIC_ProfileBars Or m.eDataType = eINDIC_ProfileVolume _
           Or (m.eDataType = eINDIC_BarData And m.eDisplayType = eINDIC_NoStyle) Then
'gdStartProfile 431
            
            ' TLB: can't do a percent change if indicator goes to 0 or negative
            If dIndMin <= 0 Then
                bPercentChangeScaling = False
            End If
            dFirstVisibleValue = kNullData
            
            For i = nScreenStartX To nScreenEndX
                dOp = gdGetNum(hOpen, gdGetNum(hXBar, i))
                dHi = gdGetNum(hHigh, gdGetNum(hXBar, i))
                dLo = gdGetNum(hLow, gdGetNum(hXBar, i))
                dCl = gdGetNum(hClose, gdGetNum(hXBar, i))
                nFlag = gdGetNum(hFlags, gdGetNum(hXBar, i))
                
                If bPercentChangeScaling And dCl > 0 Then
                    ' TLB: displayed value = this bar's value as a % change from the value of the first visible bar
                    If dFirstVisibleValue = kNullData Then
                        dFirstVisibleValue = dCl
                        dCl = 0
                        ' and do a one-time adjustment to the min/max
                        dIndMin = dIndMin / dFirstVisibleValue * 100 - 100
                        dIndMax = dIndMax / dFirstVisibleValue * 100 - 100
                    Else
                        dCl = dCl / dFirstVisibleValue * 100 - 100
                    End If
                    dOp = dOp / dFirstVisibleValue * 100 - 100
                    dHi = dHi / dFirstVisibleValue * 100 - 100
                    dLo = dLo / dFirstVisibleValue * 100 - 100
                End If
                
                gdSetNum m.geIndStruct.gdhData1, j, dOp
                gdSetNum m.geIndStruct.gdhData2, j, dHi
                gdSetNum m.geIndStruct.gdhData3, j, dLo
                gdSetNum m.geIndStruct.gdhData4, j, dCl
                gdSetNum m.geIndStruct.gdhYScaleVal, j, dCl
                
                If m.eDataType <> eINDIC_ProfileBars And m.eDataType <> eINDIC_ProfileVolume Then
                    'flags array for Wyckoff PNF chart
                    gdSetNum m.geIndStruct.glhFlags, j, nFlag
                    If nFlag <> kNullData Then m.geIndStruct.pnfX = -1
                End If
                
                j = j + 1
            Next
'gdStopProfile 431
        ElseIf m.eDataType = eINDIC_Constant Then
            gdMakeConstantValue m.geIndStruct.gdhData1, ValOfText(Parm(0)), nDpoints + 1
            gdMakeConstantValue m.geIndStruct.gdhYScaleVal, ValOfText(Parm(0)), nDpoints + 1
        ElseIf IsAutoSwingTrendlines() Then
            SetAutoTrend Chart, paneId, nID, nFirstBar, nLastBar, nScreenStartX, nScreenEndX, dExtraPriceScale
        ElseIf m.eDisplayType = eINDIC_ArtPyramid Or m.eDisplayType = eINDIC_ArtReversal Then
            SetArtPyramidData hDataArray, Chart.Bars.ArrayHandle(eBARS_High), Chart.Bars.ArrayHandle(eBARS_Low), hXBar, nScreenStartX, nScreenEndX
            dIndMin = gdMinValue(Chart.Bars.ArrayHandle(eBARS_Low), nFirstBar, nLastBar)
            dIndMax = gdMaxValue(Chart.Bars.ArrayHandle(eBARS_High), nFirstBar, nLastBar)
        ElseIf m.eDataType = eINDIC_DrawCommands Then
            dIndMin = gdMinValue(Chart.Bars.ArrayHandle(eBARS_Low), nFirstBar, nLastBar)
            dIndMax = gdMaxValue(Chart.Bars.ArrayHandle(eBARS_High), nFirstBar, nLastBar)
            
            'initialize data values to null
            j = 0
            
            hDataArray = Chart.geDateArray
            Dim aCmds As cGdArray
            Set aCmds = Chart.DrawCommands

            '08-28-2014: string format for TAS Market Map
            'name\t symbol\t    bar1;val1;bar2;val2;color
            'TASMarketMap    IBM 2742;183.58;2743;183.58;65280
            For i = 0 To aCmds.Size - 1
                
                strImage = Parse(aCmds(i), vbTab, 3)
                
                d1 = Val(Parse(strImage, ";", 1))
                d2 = Val(Parse(strImage, ";", 3))
                
                If d1 >= nScreenStartX Or d2 >= nScreenStartX Then
                    gdSetNum m.geIndStruct.gdhData1, j, gdGetNum(hBarsDT, d1)
                    gdSetNum m.geIndStruct.gdhData2, j, Val(Parse(strImage, ";", 2))
                    gdSetNum m.geIndStruct.gdhData3, j, gdGetNum(hBarsDT, d2)
                    gdSetNum m.geIndStruct.gdhData4, j, Val(Parse(strImage, ";", 4))
                    
                    gdSetNum m.geIndStruct.glhPenColor, j, Val(Parse(strImage, ";", 5))
                    j = j + 1
                End If
            Next
        Else
            If Len(m.Chart.SpreadSymbols) > 0 And Tree.Key(nID) = "PRICE" Then
                dIndMin = -1 * kNullData        'aardvark 3900 $--Auto DIVIDE $-CARMFG
                dIndMax = kNullData             'don't know why calling gdMinValue & gdMaxValue on Bars return values not in data
            End If
            
            ' TLB: can't do a percent change if indicator goes to 0 or negative
            If dIndMin <= 0 Then
                bPercentChangeScaling = False
            End If
            dFirstVisibleValue = kNullData
            
            For i = nScreenStartX To nScreenEndX
                If m.eDataType = eINDIC_Array Then
                    'dData = m.aData(gdGetNum(hXBar, i))
                    dData = gdGetNum(hDataArray, gdGetNum(hXBar, i))
                ElseIf m.eDataType = eINDIC_BarData Then
                    dData = gdGetNum(hClose, gdGetNum(hXBar, i))
                    If Tree.Key(nID) = "PRICE" Then     'aardvark 4026 fix
                        If Len(m.Chart.SpreadSymbols) > 0 Then
                            If dData <> kNullData And dIndMin > dData Then dIndMin = dData
                            If dIndMax < dData Then dIndMax = dData
                        End If
                    End If
                Else
                    dData = kNullData   'unknown dataType (theoretically should never get here)
                End If
                
                If bPercentChangeScaling And dData > 0 Then
                    ' TLB: displayed value = this bar's value as a % change from the value of the first visible bar
                    If dFirstVisibleValue = kNullData Then
                        dFirstVisibleValue = dData
                        dData = 0
                        ' and do a one-time adjustment to the min/max
                        dIndMin = dIndMin / dFirstVisibleValue * 100 - 100
                        dIndMax = dIndMax / dFirstVisibleValue * 100 - 100
                    Else
                        dData = dData / dFirstVisibleValue * 100 - 100
                    End If
                End If

                If eVwapBreakType = eVwapBreak1_NA Or m.Bars Is Nothing Then
                    d1 = d2
                ElseIf i > nScreenStartX Then
                    'apply special line break for VWAP charts
                    If eVwapBreakType = eVwapBreak2_Pivot Then
                        'faster to use date for breaking (just break where the pivot value changes)
                        d1 = m.aData(i - 1)
                        d2 = m.aData(i)
                    Else
                        d1 = m.Bars(eBARS_DateTime, gdGetNum(hXBar, i - 1))
                        d2 = m.Bars(eBARS_DateTime, gdGetNum(hXBar, i))
                    
                        If d1 = kNullData Then
                            'empty bar, get session date for two bars ago
                            d1 = m.Bars.SessionDate(gdGetNum(hXBar, i - 2))
                            d2 = m.Bars.SessionDate(gdGetNum(hXBar, i))
                        ElseIf d2 = kNullData Then
                            'empty bar, get session date for next bar
                            d1 = m.Bars.SessionDate(gdGetNum(hXBar, i - 1))
                            d2 = m.Bars.SessionDate(gdGetNum(hXBar, i + 1))
                        Else
                            d1 = m.Bars.SessionDate(gdGetNum(hXBar, i - 1))
                            d2 = m.Bars.SessionDate(gdGetNum(hXBar, i))
                        End If
                        
                        SetVwapBreak d1, d2, i, eVwapBreakType
                    End If
                End If

                If d1 = d2 Then
                    gdSetNum m.geIndStruct.gdhData1, j, dData
                    gdSetNum m.geIndStruct.gdhYScaleVal, j, dData
                Else
                    gdSetNum m.geIndStruct.gdhData1, j, kNullData
                    gdSetNum m.geIndStruct.gdhYScaleVal, j, dData
                End If
                
                j = j + 1
            Next
            If m.eDisplayType = eINDIC_Line Or m.eDisplayType = eINDIC_Steps Or m.eDisplayType = eINDIC_Rectangles Then
                'If m.strCodedName = "SwingPoints" Or m.strCodedName = "ZeroBalance" Or m.strCodedName = "ZigZag" Then
                If InStr(m.strCodedName, "Swing") > 0 Or m.strCodedName = "ZeroBalance" Or InStr(m.strCodedName, "ZigZag") > 0 Then
                    geSetSwingFirstLast hXBar, nScreenStartX, nScreenEndX, dIndMin, dIndMax
                End If
            End If
        End If  'end if display type is less than zero, i.e. LC, HLC, OHLC etc.
        
        If m.eDataType <> eINDIC_Constant Then
            If m.eDisplayType = eINDIC_Histogram Or m.eDisplayType = eINDIC_ClusterTime Then
                If dIndMin = dIndMax And dIndMin <> kNullData Then
                    If m.geIndStruct.y_baseline <> 0 Then
                        If m.geIndStruct.y_baseline > dIndMin Then
                            dIndMax = m.geIndStruct.y_baseline
                        Else
                            dIndMin = m.geIndStruct.y_baseline      'aardvark 3812
                        End If
                    Else
                        dIndMin = dIndMin - m.Chart.Bars.Prop(eBARS_MinMoveInTicks)     'aardvark 4628
                        dIndMax = dIndMax + m.Chart.Bars.Prop(eBARS_MinMoveInTicks)
                    End If
                End If
            End If
            geSetPaneMinMax ParentPane, Tree, dIndMax, dIndMin, nID, paneId, dExtraPriceScale
        End If
        
        If m.eDisplayType = eINDIC_Area Or m.eDisplayType = eINDIC_Histogram Or m.eDisplayType = eINDIC_ClusterTime Then
            m.geIndStruct.reversal = m.dColorSeperator
            m.geIndStruct.fillColor2 = m.nColorBelow
        End If
        
    End If  'end if overlayed
'gdStopProfile 430
'gdStartProfile 440
       
    'set box size & reversal for PNF & Kagi, Renko
    If m.geIndStruct.indicatorType > 8 And Not m.Bars Is Nothing Then
        If m.geIndStruct.indicatorType = 9 Then
            m.geIndStruct.boxSize = (m.Bars.Prop(eBARS_PeriodsPerBar) Mod 100000) * m.Bars.Prop(eBARS_TickMove)
            m.geIndStruct.reversal = Int(m.Bars.Prop(eBARS_PeriodsPerBar) / 100000)
        ElseIf m.geIndStruct.indicatorType = 10 Then
            m.geIndStruct.reversal = (m.Bars.Prop(eBARS_PeriodsPerBar) Mod 100000) * m.Bars.Prop(eBARS_TickMove)
        ElseIf m.geIndStruct.indicatorType = 11 Then
            m.geIndStruct.boxSize = (m.Bars.Prop(eBARS_PeriodsPerBar) Mod 100000) * m.Bars.Prop(eBARS_TickMove)
            ParentPane.gePaneMax = ParentPane.gePaneMax + m.geIndStruct.boxSize / 2
        ElseIf m.geIndStruct.indicatorType = 23 Then
            'market profile style display
            If m.nTicksPerRow <> -1 And m.nTicksPerRow = 0 Then m.nTicksPerRow = -1     'ticksPerRow = -1 means auto ticks per row
            If m.nSessionsPerProfile <= 0 Then m.nSessionsPerProfile = 2
            m.geIndStruct.boxSize = m.nTicksPerRow
            m.geIndStruct.reversal = m.nSessionsPerProfile
        End If
        If m.geIndStruct.pnfX <> -1 Then
            If nScreenStartX Mod 2 = 0 Then
                If m.Bars.Prop(eBARS_OddBarsUp) <> 0 Then
                    m.geIndStruct.pnfX = 0
                Else
                    m.geIndStruct.pnfX = 1
                End If
            Else
                If m.Bars.Prop(eBARS_OddBarsUp) = 0 Then
                    m.geIndStruct.pnfX = 0
                Else
                    m.geIndStruct.pnfX = 1
                End If
            End If
        End If
    End If
    
    'set additional info needed by Kagi charts
    'a. high/low of bar immediately preceding first visible data bar
    'b. flag indicating whether to use normal or thicker pen for first bar
    gdSetNum m.geIndStruct.glhPenStyle, 0, 0    '0=use normal, 1=use thicker
    If m.geIndStruct.indicatorType = 9 Then     'P&F
        If nScreenStartX > 0 Then
            m.geIndStruct.prevBarHigh = gdGetNum(hHigh, gdGetNum(hXBar, nScreenStartX - 1))
            m.geIndStruct.prevBarLow = gdGetNum(hLow, gdGetNum(hXBar, nScreenStartX - 1))
        End If
    ElseIf m.geIndStruct.indicatorType = 10 Then    'kagi
        geSetPenStyle eVwapBreakType
        If nScreenStartX > 0 Then
            m.geIndStruct.prevBarHigh = gdGetNum(hHigh, gdGetNum(hXBar, nScreenStartX - 1))
            m.geIndStruct.prevBarLow = gdGetNum(hLow, gdGetNum(hXBar, nScreenStartX - 1))
            If m.geIndStruct.prevBarHigh > ParentPane.gePaneMax Then
                ParentPane.gePaneMax = m.geIndStruct.prevBarHigh
            End If
            If m.geIndStruct.prevBarLow < ParentPane.gePaneMin Then
                ParentPane.gePaneMin = m.geIndStruct.prevBarLow
            End If
            For i = 1 To nScreenStartX
                If gdGetNum(hHigh, i) > gdGetNum(hHigh, i - 1) Then
                    gdSetNum m.geIndStruct.glhPenStyle, 0, 1
                ElseIf gdGetNum(hLow, i) < gdGetNum(hLow, i - 1) Then
                    gdSetNum m.geIndStruct.glhPenStyle, 0, 0
                End If
            Next
        End If
    End If
    
'gdStopProfile 440
'gdStartProfile 450
    
    'set pen style, size & color
    If m.geIndStruct.indicatorType <> 10 Then
        geSetPenStyle eVwapBreakType
    End If
    

#If 0 Then ' Test for TAS stuff
    'j = m.strName
    If UCase(Left(m.strCodedName, 10)) = "TAS_RESULT" Then
        j = nID
        j = Tree.NodeLevel(nID)
        For j = nID To paneId + 1 Step -1
            If Tree.NodeLevel(j) < Tree.NodeLevel(nID) Then
                Set IndToColor = Tree(j)
                If IndToColor.Data.IsConstantValue Then
                    k = IndToColor.Data(0)
                End If
                Exit For
            End If
        Next
        Set IndToColor = Nothing
    End If
#End If
    
    
    If m.strCodedName = "RedLightGreenLight" Then
        ' walk through each bucket on the screen
        For j = nScreenStartX To nScreenEndX
            ' get the bar# and data for that bucket (and ignore if an empty bucket)
            i = Chart.aXBar(j)
            dData = m.aData(i)
            If i <> kNullData And dData <> kNullData Then
                If dData >= 50 Then
                    gdSetNum m.geIndStruct.glhPenColor, j - nScreenStartX, RGB(0, 128, 0)
                Else
                    gdSetNum m.geIndStruct.glhPenColor, j - nScreenStartX, vbRed
                End If
            End If
        Next
    ElseIf UCase(m.strName) = "TAS NAVHIST" Then
        ' walk through each bucket on the screen
        For j = nScreenStartX To nScreenEndX
            ' get the bar# and data for that bucket (and ignore if an empty bucket)
            i = Chart.aXBar(j)
            dData = m.aData(i)
            If i <> kNullData And dData <> kNullData Then
                ' get color for the histogram (a hack: it's just the 10th digit after the decimal point)
                k = Round(Abs(dData - Round(dData, 9)) * 10000000000#)
                Select Case k
                Case 1
                    k = RGB(0, 224, 0) ' lime
                Case 2
                    k = RGB(0, 128, 0) ' darker green
                Case 3
                    k = RGB(160, 0, 0) ' maroon
                Case 4
                    k = RGB(255, 0, 0) ' red
                Case 5
                    k = RGB(255, 0, 255) ' pink
                Case Else
                    k = m.nColor
                End Select
                gdSetNum m.geIndStruct.glhPenColor, j - nScreenStartX, k
            End If
        Next
    ElseIf UCase(m.strName) = "TAS VEGA" Then
        Set tResults = Nothing
        If m.aData.IsConstantValue And m.bDisplay Then
            i = m.aData(m.aData.Size - 1) ' get the TAS IndicatorID that is stored here
            Set tResults = TASGetResultsTable(i)
            Set PriceInd = Tree("PRICE")
        End If
        If Not PriceInd Is Nothing And Not tResults Is Nothing Then
            ' walk through each bucket on the screen
            For j = nScreenStartX To nScreenEndX
                ' get the bar# and data for that bucket (and ignore if an empty bucket)
                i = Chart.aXBar(j)
                ' get the color for the price bars
                k = tResults(16, i)
                If i <> kNullData And k <> kNullData Then
                    Select Case k
                    Case 1
                        k = RGB(255, 165, 0)
                    Case 2
                        k = RGB(128, 128, 128) ' gray
                    Case 4
                        k = RGB(0, 224, 0) ' lime
                    Case 5
                        k = vbRed
                    Case 6
                        k = vbMagenta
                    Case Else
                        k = 0 ' black
                    End Select
                    PriceInd.geSetPenColor j - nScreenStartX, k
                End If
            Next
        End If
    ElseIf IsHawkeyeLevels Then
        strImage = HawkeyeLevelsColors
        For i = 1 To 9
            strColor = Parse(strImage, ";", i)
            If Len(strColor) > 0 Then
                gdSetNum m.geIndStruct.glhPenColor, i - 1, Val(strColor)
            Else
                geSetPenColor i - 1, m.nColor
                'gdSetNum m.geIndStruct.glhPenColor, i - 1, m.nColor
            End If
        Next
    ElseIf m.eDataType <> eINDIC_BooleanArray And m.eDataType <> eINDIC_ProfileBars And _
        m.eDataType <> eINDIC_ProfileVolume And m.eDataType <> eINDIC_DrawCommands Then
        For i = 0 To nDpoints
            If m.eDisplayType = eINDIC_BollingerBar Then
                If m.nBollingerColor < 0 Then
                    gdSetNum m.geIndStruct.glhPenColor, i, RGB(0, 0, 192)
                Else
                    gdSetNum m.geIndStruct.glhPenColor, i, m.nBollingerColor
                End If
            Else
                geSetPenColor i, m.nColor
                'gdSetNum m.geIndStruct.glhPenColor, i, m.nColor
            End If
        Next
    End If
    
    If m.geAdded = False Then
        rc = geAddItem(chartObj, 3, m.geIndStruct)
    End If
    If 0 = rc Then m.geAdded = True
    geSetIndData = rc

    ' restore original style if had done split mode (for 1-tick-per-bar using close line)
    Dim nAltSize&
    
    If nStartCloseLine > 0 Then
        m.eDisplayType = nSaveOHLCtype
        m.geIndStruct.prevBarHigh = nStartCloseLine - nScreenStartX
        'set alternate indicator type for split mode
        If m.eDisplayType = eINDIC_Points Then
            m.geIndStruct.prevBarLow = 0
            If m.eStyle <> eINDIC_Default Then
                nAltSize = m.eStyle
            Else
                nAltSize = eINDIC_Medium
            End If
            m.geIndStruct.altPenSize = nAltSize
        ElseIf m.eDisplayType = eINDIC_Area Then
            m.geIndStruct.prevBarLow = 3
        ElseIf m.eDisplayType = eINDIC_Histogram Or m.eDisplayType = eINDIC_ClusterTime Then
            m.geIndStruct.prevBarLow = 4
            If m.eStyle <> eINDIC_Default Then
                nAltSize = m.eStyle
            Else
                nAltSize = eINDIC_Medium
            End If
            m.geIndStruct.fillPct = nAltSize * 15 + 10
        Else
            m.geIndStruct.prevBarLow = 1    'default alternate type to line
            If m.eStyle <> eINDIC_Default Then
                nAltSize = m.eStyle
            Else
                nAltSize = g.ChartGlobals.eDefaultIndStyle
            End If
            If nAltSize = eINDIC_Dash Then
                m.geIndStruct.altPenStyle = 1
            ElseIf nAltSize = eINDIC_Dot Then
                m.geIndStruct.altPenStyle = 2
            ElseIf nAltSize = eINDIC_DashDot Then
                m.geIndStruct.altPenStyle = 3
            Else
                m.geIndStruct.altPenSize = nAltSize
            End If
        End If
    End If
    
'gdStopProfile 450

ErrExit:
    If hBoolScreenArray <> 0 Then gdDestroyArray hBoolScreenArray
'gdStopProfile 420
'gdStopProfile 400
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.geSetIndData", eGDRaiseError_Raise

End Function

' Sets the pen width and style along with the indicator type
Private Sub geSetPenStyle(Optional ByVal eVwapBreakType As eVwapBreak = eVwapBreak1_NA)
On Error GoTo ErrSection:

    Dim nUseStyle As eIndicatorStyle
    
    ' defaults
    m.geIndStruct.nullValStyle = 1      'default to span-points for null values
    gdSetNum m.geIndStruct.glhPenWidth, 0, 1
    gdSetNum m.geIndStruct.glhPenStyle, 0, 0
    
    ' based on display type
    Select Case m.eDisplayType
        Case eINDIC_OHLC, eINDIC_HLC, eINDIC_HL, eINDIC_PNF, eINDIC_Kagi, eINDIC_Renko, eINDIC_Profile
            If m.eStyle <> eINDIC_Default Then
                nUseStyle = m.eStyle
            Else
                nUseStyle = g.ChartGlobals.eDefaultBarsStyle
            End If
            If nUseStyle >= eINDIC_Auto Then
                ' make pen size variable (negative size = thickness as % of pixel width)
                gdSetNum m.geIndStruct.glhPenWidth, 0, -15
            Else
                ' make pen size fixed (positive size)
                gdSetNum m.geIndStruct.glhPenWidth, 0, nUseStyle
            End If
            m.geIndStruct.fillPct = 80
        
        Case eINDIC_Candlestick, eINDIC_BollingerBar:
            If m.eDisplayType = eINDIC_Candlestick Then
                m.geIndStruct.indicatorType = 5
            Else
                m.geIndStruct.indicatorType = 13
            End If
            If m.eStyle <> eINDIC_Default Then
                nUseStyle = m.eStyle
            Else
                nUseStyle = g.ChartGlobals.eDefaultBarsStyle
            End If
            If nUseStyle >= eINDIC_Auto Then
                ' make pen size variable (negative size = thickness as % of pixel width)
                gdSetNum m.geIndStruct.glhPenWidth, 0, -9
            Else
                ' make pen size fixed (positive size)
                gdSetNum m.geIndStruct.glhPenWidth, 0, nUseStyle
            End If
            m.geIndStruct.fillPct = 75
            If m.bHideWick And m.eDisplayType = eINDIC_Candlestick Then
                m.geIndStruct.pnfX = -1 ' to hide wick
            Else
                m.geIndStruct.pnfX = 0
            End If
        
        Case eINDIC_Line, eINDIC_Steps, eINDIC_Rectangles, eINDIC_ArtPyramid, eINDIC_ArtReversal, _
            eINDIC_ClusterPrice, eINDIC_Ribbon
            'set indicator type
            If m.eDataType = eINDIC_Constant Then
                m.geIndStruct.indicatorType = 2
            ElseIf m.eDisplayType = eINDIC_Steps Then
                m.geIndStruct.indicatorType = 14
            ElseIf m.eDisplayType = eINDIC_Rectangles Then
                m.geIndStruct.indicatorType = 16
                m.geIndStruct.nullValStyle = 2          'don't draw for blank bars (i.e. holidays)
            ElseIf m.eDisplayType = eINDIC_ArtPyramid Then
                m.geIndStruct.indicatorType = 17
            ElseIf m.eDisplayType = eINDIC_ArtReversal Then
                m.geIndStruct.indicatorType = 18
            ElseIf m.eDisplayType = eINDIC_ClusterPrice Then
                m.geIndStruct.indicatorType = 20
            ElseIf IsAutoSwingTrendlines() Then
                m.geIndStruct.indicatorType = 15
            ElseIf m.eDisplayType = eINDIC_Ribbon Then
                m.geIndStruct.indicatorType = 22
            ElseIf m.eDataType = eINDIC_ProfileVolume Then
                m.geIndStruct.indicatorType = 24
            Else
                m.geIndStruct.indicatorType = 1
                If InStr(m.strCodedName, "HawkeyeBarrier") <> 0 Then
                    m.geIndStruct.nullValStyle = 2          'don't draw null values for this indicator per Nigel - 09-29-2009
                End If
            End If
            'set pen style and size
            If m.eStyle <> eINDIC_Default Then
                nUseStyle = m.eStyle
            ElseIf m.eDataType = eINDIC_Constant Then   'Hz line
                nUseStyle = g.ChartGlobals.eDefaultHorzStyle
            Else
                nUseStyle = g.ChartGlobals.eDefaultIndStyle
            End If
            If nUseStyle = eINDIC_Dash Then
                gdSetNum m.geIndStruct.glhPenStyle, 0, 1
            ElseIf nUseStyle = eINDIC_Dot Then
                gdSetNum m.geIndStruct.glhPenStyle, 0, 2
            ElseIf nUseStyle = eINDIC_DashDot Then
                gdSetNum m.geIndStruct.glhPenStyle, 0, 3
            Else
                gdSetNum m.geIndStruct.glhPenWidth, 0, nUseStyle
            End If
            
            'set special line break style for VWAP charts
            If eVwapBreakType <> eVwapBreak1_NA Then
                If m.eDisplayType = eINDIC_Line Then m.geIndStruct.nullValStyle = 4
            End If
        
        Case eINDIC_Histogram, eINDIC_ClusterTime
            m.geIndStruct.indicatorType = 4     'histogram
            m.geIndStruct.nullValStyle = 0      'draw null values as zero
            If m.eStyle <> eINDIC_Default Then
                nUseStyle = m.eStyle
            Else
                nUseStyle = eINDIC_Medium
            End If
            'note that pen size stays at the default size 1 set above
            'if it becomes desireable to use pen sizes other than 1 for
            'histograms -- a change is required in the grapheng.dll
            m.geIndStruct.fillPct = nUseStyle * 15 + 10
        Case eINDIC_Area:
            m.geIndStruct.indicatorType = 3
            m.geIndStruct.nullValStyle = 0      'draw null values as zero
        
        Case eINDIC_Points:
            m.geIndStruct.indicatorType = 0
            m.geIndStruct.nullValStyle = 2      'don't draw null values
            If m.eStyle <> eINDIC_Default Then
                nUseStyle = m.eStyle
            Else
                nUseStyle = eINDIC_Medium
            End If
            gdSetNum m.geIndStruct.glhPenWidth, 0, nUseStyle
        
        Case eINDIC_ValueMarkers
            m.geIndStruct.indicatorType = 19
                
        Case eINDIC_NoStyle
            m.geIndStruct.indicatorType = 30
    
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.geSetPenStyle", eGDRaiseError_Raise
    
End Sub

Private Sub geSetSwingFirstLast(hXBar&, nScreenStartX&, nScreenEndX&, dIndMin#, dIndMax#)
On Error GoTo ErrSection:

    Dim dData#, dRatio#, hDataArray&
    Dim i&, k&, nLast&
    
    ' see if need to interpolate at beginning of chart
    hDataArray = m.aData.ArrayHandle
    If gdGetNum(m.geIndStruct.gdhData1, 0) = kNullData Then
        dData = kNullData
        For i = nScreenStartX To nScreenEndX
            'dData = m.aData(gdGetNum(hXBar, i))
            dData = gdGetNum(hDataArray, gdGetNum(hXBar, i))
            If dData <> kNullData Then
                For k = nScreenStartX - 1 To 0 Step -1
                    'If m.aData(gdGetNum(hXBar, k)) <> kNullData Then
                    If gdGetNum(hDataArray, gdGetNum(hXBar, k)) <> kNullData Then
                        ' do the interpolation
                        dRatio = Abs(nScreenStartX - i) / Abs(i - k)
                        'dData = dRatio * (m.aData(gdGetNum(hXBar, k)) - dData) + dData
                        dData = dRatio * (gdGetNum(hDataArray, gdGetNum(hXBar, k)) - dData) + dData
                        gdSetNum m.geIndStruct.gdhData1, 0, dData
                        gdSetNum m.geIndStruct.gdhYScaleVal, 0, dData
                        If dData < dIndMin Then dIndMin = dData
                        If dData > dIndMax Then dIndMax = dData
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    End If
    ' see if need to interpolate at end of chart
    nLast = gdGetSize(m.geIndStruct.gdhData1) - 1
    If gdGetNum(m.geIndStruct.gdhData1, nLast) = kNullData Then
        dData = kNullData
        For i = nScreenEndX To nScreenStartX Step -1
            'dData = m.aData(gdGetNum(hXBar, i))
            dData = gdGetNum(hDataArray, gdGetNum(hXBar, i))
            If dData <> kNullData Then
                For k = nScreenEndX To gdGetSize(hXBar) - 1
                    If gdGetNum(hDataArray, gdGetNum(hXBar, k)) <> kNullData Then
                        ' do the interpolation
                        dRatio = Abs(nScreenEndX - i) / Abs(i - k)
                        dData = dRatio * (gdGetNum(hDataArray, gdGetNum(hXBar, k)) - dData) + dData
                        gdSetNum m.geIndStruct.gdhData1, nLast, dData
                        gdSetNum m.geIndStruct.gdhYScaleVal, nLast, dData
                        If dData < dIndMin Then dIndMin = dData
                        If dData > dIndMax Then dIndMax = dData
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    End If  'end if need to interpolate at end of chart
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.geSetSwingFirstLast", eGDRaiseError_Raise

End Sub

Private Sub CalcOverlayVars(dMultipler As Double, dOffSet As Double, _
    ByVal dIndMin#, ByVal dIndMax#, ByVal dPaneMin#, ByVal dPaneMax#, _
    ByVal eLogFlag As ePANE_LogFlag)
On Error GoTo ErrSection:
    
    'initialize return variables
    dMultipler = 0
    dOffSet = 0
    
    If dIndMax = kNullData Or dIndMin = kNullData Then Exit Sub
    
    If dIndMin = dIndMax Then
        dMultipler = 0
    ElseIf eLogFlag = ePANE_LogFlagLog Then
        dMultipler = (Log(dPaneMax) - Log(dPaneMin)) / (dIndMax - dIndMin)
        dOffSet = Log(dPaneMin) - (dIndMin * dMultipler)
    Else
        dMultipler = (dPaneMax - dPaneMin) / (dIndMax - dIndMin)
        dOffSet = dPaneMin - (dIndMin * dMultipler)
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cIndicator.CalcOverlayVars"
    
End Sub

Private Sub geSetOverlayData(dIndMin#, dIndMax#, ByVal dPaneMin#, ByVal dPaneMax#, _
        ByVal nScreenStartX&, ByVal nScreenEndX&, ByVal eLogFlag As ePANE_LogFlag, ByVal hXBar&)
On Error GoTo ErrSection:

    Dim i&, j&
    Dim dMult#, dOffSet#, dData#
    Dim dOp#, dHi#, dLo#, dCl#
    
    Dim bClusterPrice As Boolean
    
    If m.bFlip And dIndMin <> kNullData Then
        ' need to invert the min/max when flipping the data
        dData = dIndMin
        dIndMin = -dIndMax
        dIndMax = -dData
    End If
    
    CalcOverlayVars dMult, dOffSet, dIndMin, dIndMax, dPaneMin, dPaneMax, eLogFlag
    
    j = 0
    'candlestick = -1, hl = -2, hlc = -3, ohlc = -4
    'line = 0, histogram = 1, area = 2, points = 3
    If m.eDisplayType < 0 Then
        For i = nScreenStartX To nScreenEndX
            dCl = m.Bars.Item(eBARS_Close, gdGetNum(hXBar, i))
            If m.bFlip And dCl <> kNullData Then
                dOp = -m.Bars.Item(eBARS_Open, gdGetNum(hXBar, i))
                dLo = -m.Bars.Item(eBARS_High, gdGetNum(hXBar, i))
                dHi = -m.Bars.Item(eBARS_Low, gdGetNum(hXBar, i))
                dCl = -dCl
            Else
                dOp = m.Bars.Item(eBARS_Open, gdGetNum(hXBar, i))
                dHi = m.Bars.Item(eBARS_High, gdGetNum(hXBar, i))
                dLo = m.Bars.Item(eBARS_Low, gdGetNum(hXBar, i))
            End If
            
            gdSetNum m.geIndStruct.gdhYScaleVal, j, dCl
            
            If dOp <> kNullData Then
                If eLogFlag = ePANE_LogFlagLog Then
                    dOp = Exp(dOp * dMult + dOffSet)
                Else
                    dOp = dOp * dMult + dOffSet
                End If
            End If
                        
            If dHi <> kNullData Then
                If eLogFlag = ePANE_LogFlagLog Then
                    dHi = Exp(dHi * dMult + dOffSet)
                Else
                    dHi = dHi * dMult + dOffSet
                End If
            End If
            
            If dLo <> kNullData Then
                If eLogFlag = ePANE_LogFlagLog Then
                    dLo = Exp(dLo * dMult + dOffSet)
                Else
                    dLo = dLo * dMult + dOffSet
                End If
            End If
            
            If dCl <> kNullData Then
                If eLogFlag = ePANE_LogFlagLog Then
                    dCl = Exp(dCl * dMult + dOffSet)
                Else
                    dCl = dCl * dMult + dOffSet
                End If
            End If
            
            gdSetNum m.geIndStruct.gdhData1, j, dOp
            gdSetNum m.geIndStruct.gdhData2, j, dHi
            gdSetNum m.geIndStruct.gdhData3, j, dLo
            gdSetNum m.geIndStruct.gdhData4, j, dCl
            
            j = j + 1
        Next
    Else
        
        If m.eDisplayType = eINDIC_ClusterPrice And m.geIndStruct.TrueRangeFlag <> 2 Then
            If Not m.Chart Is Nothing Then
                If Not m.Chart.Bars Is Nothing Then bClusterPrice = True
            End If
        End If
        
        For i = nScreenStartX To nScreenEndX
            dData = kNullData
            
            If m.eDataType = eINDIC_Array Then
                'dData = m.aData(i)
                dData = m.aData(gdGetNum(hXBar, i))
                
                If bClusterPrice Then
                    'JM 11-10-2010: for now, always label swing points (m.nClusterSwingLabel flag not used)
                    'do this for 1-to-1 correspondence with gdhData1 array
                    gdSetNum m.geIndStruct.glhImageColor, j, -1
                    gdSetStr m.geIndStruct.gdshImage, j, ""
                    Select Case Abs(dData)
                        Case 1:
                            gdSetNum m.geIndStruct.glhImageColor, j, m.nClusterSwingColorS
                            gdSetStr m.geIndStruct.gdshImage, j, "S"
                            If dData = -1 Then gdSetStr m.geIndStruct.gdshImage, j, "-S"
                        Case 2:
                            gdSetNum m.geIndStruct.glhImageColor, j, m.nClusterSwingColorM
                            gdSetStr m.geIndStruct.gdshImage, j, "I"
                            If dData = -2 Then gdSetStr m.geIndStruct.gdshImage, j, "-I"
                        Case 3:
                            gdSetNum m.geIndStruct.glhImageColor, j, m.nClusterSwingColorL
                            gdSetStr m.geIndStruct.gdshImage, j, "L"
                            If dData = -3 Then gdSetStr m.geIndStruct.gdshImage, j, "-L"
                    End Select
                    If dData <> kNullData Then
                        If dData > 0 Then
                            dData = m.Chart.Bars.Item(eBARS_High, gdGetNum(hXBar, i))
                        ElseIf dData < 0 Then
                            dData = m.Chart.Bars.Item(eBARS_Low, gdGetNum(hXBar, i))
                        ElseIf dData = 0 Then
                            dData = kNullData
                        End If
                    End If
                End If
            
            ElseIf m.eDataType = eINDIC_BarData Then
                dData = m.Bars.Item(eBARS_Close, gdGetNum(hXBar, i))
            End If
            
            If m.bFlip And dData <> kNullData Then dData = -dData
            
            gdSetNum m.geIndStruct.gdhYScaleVal, j, dData   'aardvark issue 669 fix
            
            If Not bClusterPrice And dData <> kNullData Then
                If eLogFlag = ePANE_LogFlagLog Then
                    dData = Exp(dData * dMult + dOffSet)
                Else
                    dData = dData * dMult + dOffSet
                End If
            End If
            
            gdSetNum m.geIndStruct.gdhData1, j, dData
            
            j = j + 1
        Next
        
        If bClusterPrice Then
            'first y-value on screen is null, step backwards through bars and set low/high of first swing
            'point that is off chart-left as well as x-bucket number so grapheng can extrapolate line
            If gdGetNum(m.geIndStruct.gdhData1, 0) = kNullData Then
                For i = nScreenStartX To 0 Step -1
                    dData = m.aData(gdGetNum(hXBar, i))
                    If dData <> kNullData And dData <> 0 Then
                        If dData > 0 Then
                            'first swing visible on-screen was high, so set first swing off-chart left as low
                            m.geIndStruct.prevBarHigh = m.Chart.Bars.Item(eBARS_High, gdGetNum(hXBar, i))
                        ElseIf dData < 0 Then
                            m.geIndStruct.prevBarHigh = m.Chart.Bars.Item(eBARS_Low, gdGetNum(hXBar, i))
                        End If
                        m.geIndStruct.prevBarLow = i - nScreenStartX        'bucket number
                        Exit For
                    End If
                Next
            End If
            
            j = gdGetSize(m.geIndStruct.gdshImage)
            'color does not matter, just need glhImageColor to be same size as gdshImage
            gdSetNum m.geIndStruct.glhImageColor, j, m.nClusterSwingColorS
            geSetFontInfo j
            
        ElseIf m.eDisplayType = eINDIC_Histogram Or m.eDisplayType = eINDIC_Area Then
            dData = m.geIndStruct.y_baseline
            If eLogFlag = ePANE_LogFlagLog Then
                dData = Exp(dData * dMult + dOffSet)
            Else
                dData = dData * dMult + dOffSet
            End If
            m.geIndStruct.y_baseline = dData
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.geSetOverlayData", eGDRaiseError_Raise
    
End Sub

Private Sub geSetPaneMinMax(Pane As cPane, Tree As cGdTree, IndMax#, IndMin#, _
    ByVal indId&, ByVal paneId&, ByVal dExtraPriceScale)
On Error GoTo ErrSection:

    Dim z#
        
    If Pane.Scaling = ePANE_ScaleModeManual Then
        If Pane.Max > Pane.Min Then
            'user selected manual scale & specified pane's min/max
            Pane.gePaneMax = Pane.Max
            Pane.gePaneMin = Pane.Min
            Pane.geSavePaneMax
            Pane.geSavePaneMin
            Exit Sub
        Else
            Pane.Scaling = ePANE_ScaleModeAuto
        End If
    ElseIf Pane.Scaling = ePANE_ScaleModeAutoPrice Then
        If Tree.Key(indId) <> "PRICE" Then
            Exit Sub
        End If
    End If
    
    'don't use horizontal indicators to determine pane's min/ max (aardvark #255)
    If m.eDisplayType = eINDIC_NoStyle Then
        Exit Sub
    ElseIf m.eDisplayType = eINDIC_Line And m.eDataType = eINDIC_Constant Then
        Exit Sub
    End If
        
    If Me.CodedName = "RedLightGreenLight" Or Me.CodedName = "SeasonalPercent" Then
        Pane.gePaneMin = 0
        Pane.gePaneMax = 100
    'ElseIf indId - paneId = 1 Then
    ElseIf Pane.gePaneMin = kNullData Then
        'if this is first indicator of pane then set pane's min max without checking
        If IndMin <> kNullData Then Pane.gePaneMin = IndMin
        If IndMax <> kNullData Then Pane.gePaneMax = IndMax
    Else
        If IndMin <> kNullData And IndMin < Pane.gePaneMin Then Pane.gePaneMin = IndMin
        If IndMax <> kNullData And IndMax > Pane.gePaneMax Then Pane.gePaneMax = IndMax
    End If
    If Tree.Key(paneId) = "PRICE PANE" Then
        z = (IndMax - IndMin) * dExtraPriceScale
        IndMax = IndMax + z
        If Pane.PaneLogFlag = ePANE_LogFlagLog And IndMin > 0 And z >= IndMin Then
            IndMin = IndMin / 2 '(so won't go <= 0)
        Else
            IndMin = IndMin - z
        End If
    'Need substitute code for manualscale function
    'Histograms should not start above 0 if scaletype is regular
'    ElseIf Pane.gePaneMin >= 0 And Pane.geScaleType = 0 And m.eDisplayType = eINDIC_Histogram Then
'        Pane.gePaneMin = 0     - no longer true (aardvark 3611)
    End If
    
    With Pane
        .Max = .gePaneMax           '3922
        .Min = .gePaneMin
        .geSavePaneMax
        .geSavePaneMin
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.geSetPaneMinMax", eGDRaiseError_Raise
    
End Sub

Private Sub FixInputs()
On Error GoTo ErrSection:

    ' see if a market parm needs to be converted to an expression parm
    ' (e.g. first arg of RSI was converted in July 2003)
    If Not m.Inputs Is Nothing Then
        If m.Inputs.Count >= 1 Then
            With m.Inputs.Item(1)
                If .ParmTypeID = 4 And UCase(Trim(.Value)) = "MARKET1" Then
                    .Value = "Close" '(default)
                End If
            End With
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.FixInputs", eGDRaiseError_Raise
    
End Sub

Public Property Get MarkerImage(Optional ByVal bForDisplay As Boolean = False) As eStockImage
On Error GoTo ErrSection:

    Dim nDisplayIdx&

    If bForDisplay = False Or m.eMarker = eCNI_Ascii Then
        MarkerImage = m.eMarker
        Exit Property
    End If

    'Note: this code converts the enumerated image type to an index value
    'matching the picture box array in the gdSelectIcon control (gdOcx project)
    
    Select Case m.eMarker
        Case eCNI_Arrow
            Select Case m.eMarkerDir
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
            If m.nMarkerFill = 0 Then
                nDisplayIdx = 17
            Else
                nDisplayIdx = 12
            End If
        Case eCNI_Square
            If m.nMarkerFill = 0 Then
                nDisplayIdx = 18
            Else
                nDisplayIdx = 13
            End If
        Case eCNI_Diamond
            If m.nMarkerFill = 0 Then
                nDisplayIdx = 19
            Else
                nDisplayIdx = 14
            End If
        Case eCNI_Triangle
            If m.nMarkerFill = 0 Then
                If m.eMarkerDir = eCNI_North Then
                    nDisplayIdx = 15
                Else
                    nDisplayIdx = 16
                End If
            Else
                If m.eMarkerDir = eCNI_North Then
                    nDisplayIdx = 10
                Else
                    nDisplayIdx = 11
                End If
            End If
    End Select
    
    MarkerImage = nDisplayIdx

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerImage.Get", eGDRaiseError_Raise
    
End Property

Public Property Let MarkerImage(Optional ByVal bForDisplay As Boolean = False, ByVal eImage As eStockImage)
On Error GoTo ErrSection:

    m.eMarker = eImage

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerImage.Let", eGDRaiseError_Raise
    
End Property

Public Property Get MarkerDir() As eImageDir
On Error GoTo ErrSection:

    MarkerDir = m.eMarkerDir

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerDir.Get", eGDRaiseError_Raise
    
End Property

Public Property Let MarkerDir(ByVal eDir As eImageDir)
On Error GoTo ErrSection:

    m.eMarkerDir = eDir

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerDir.Let", eGDRaiseError_Raise
    
End Property

Public Property Get MarkerFill() As Long
 On Error GoTo ErrSection:

    MarkerFill = m.nMarkerFill
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerFill.Get", eGDRaiseError_Raise
    
End Property

Public Property Let MarkerFill(ByVal nFill&)
On Error GoTo ErrSection:

    m.nMarkerFill = nFill
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerFill.Let", eGDRaiseError_Raise
    
End Property

Public Property Get MarkerLoc() As Long
On Error GoTo ErrSection:

    MarkerLoc = m.nMarkerLoc

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerLoc.Get", eGDRaiseError_Raise
    
End Property

Public Property Let MarkerLoc(ByVal nLoc&)
On Error GoTo ErrSection:

    m.nMarkerLoc = nLoc

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerLoc.Let", eGDRaiseError_Raise
    
End Property

Public Property Get MarkerAscii() As String
On Error GoTo ErrSection:

    MarkerAscii = Left(m.strMarkerAscii, 1)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerAscii.Get", eGDRaiseError_Raise
    
End Property

Public Property Let MarkerAscii(ByVal strAscii As String)
On Error GoTo ErrSection:

    m.strMarkerAscii = Left(strAscii, 1)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.MarkerAscii.Let", eGDRaiseError_Raise
    
End Property

Public Sub geSetMarkersData(strImage$, nMarkerLoc&, nDir&, nFill&, nColor&, nImageSize&)
On Error GoTo ErrSection:

    Dim Size&

    Size = gdGetSize(m.geIndStruct.gdshImage)
    
    gdSetStr m.geIndStruct.gdshImage, Size, strImage
    
    If nImageSize < 1 Then
        nImageSize = 1
    End If
    
    gdSetNum m.geIndStruct.glhImageLoc, Size, nMarkerLoc
    gdSetNum m.geIndStruct.glhImageDir, Size, nDir
    gdSetNum m.geIndStruct.glhImageFill, Size, nFill
    gdSetNum m.geIndStruct.glhImageColor, Size, nColor
    gdSetNum m.geIndStruct.glhImageSize, Size, nImageSize
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.geSetMarkersData", eGDRaiseError_Raise
    
End Sub

Public Property Get BoolRefArray() As Long
    BoolRefArray = m.aBoolIndRef.ArrayHandle
End Property

Public Property Let BadTicksFlag(ByVal bFlag As Boolean)
    m.bBadTicksInd = bFlag
End Property

Public Property Get BadTicksFlag() As Boolean
    BadTicksFlag = m.bBadTicksInd
End Property

Public Property Get CanBeHighlight() As Long
    CanBeHighlight = m.nCanBeHighlight
End Property

Public Property Let CanBeHighlight(ByVal nFlag&)
    m.nCanBeHighlight = nFlag
End Property

Public Property Get MarkerPrompt() As Boolean
    MarkerPrompt = m.bMarkerPrompt
End Property

Public Property Let MarkerPrompt(ByVal bPrompt As Boolean)
    m.bMarkerPrompt = bPrompt
End Property

Public Property Get IndMin(ByVal nFirstBar&, ByVal nLastBar&) As Double
On Error GoTo ErrSection:

    Select Case m.eDataType
        Case eINDIC_BarData
            IndMin = gdMinValue(m.Bars.ArrayHandle(eBARS_Low), nFirstBar, nLastBar)
        Case eINDIC_Array
            IndMin = gdMinValue(m.aData.ArrayHandle, nFirstBar, nLastBar)
        Case eINDIC_Constant
            IndMin = Val(Parm(0))
        Case Else
            IndMin = kNullData
    End Select
    
    Exit Property

ErrSection:
    RaiseError "cIndicator.IndMin.Get"
    
End Property

Public Property Get IndMax(ByVal nFirstBar&, ByVal nLastBar&) As Double
On Error GoTo ErrSection:

    Select Case m.eDataType
        Case eINDIC_BarData
            IndMax = gdMaxValue(m.Bars.ArrayHandle(eBARS_High), nFirstBar, nLastBar)
        Case eINDIC_Array
            IndMax = gdMaxValue(m.aData.ArrayHandle, nFirstBar, nLastBar)
        Case eINDIC_Constant
            IndMax = Val(Parm(0))
        Case Else
            IndMax = kNullData
    End Select
    
    Exit Property

ErrSection:
    RaiseError "cIndicator.IndMax.Get"
    
End Property

Public Property Get UpColor() As Long
    UpColor = m.nUpColor
End Property

Public Property Let UpColor(ByVal nColor&)
On Error GoTo ErrSection:
    
    Dim Annot As cAnnotation
    Dim Chart As cChart
    Dim i&
    
    m.nUpColor = nColor

    If IsAutoSwingTrendlines Then
        If Not ActiveChart Is Nothing Then
            Set Chart = ActiveChart.Chart
            For i = 0 To Chart.Annots.Count
                Set Annot = Chart.Annots(i)
                If Not Annot Is Nothing Then
                    If Annot.eUsage = eANNOT_AutoSwingTrend Then
                        If Annot.Prop("Trend") = "Up" Then
                            Annot.Color = m.nUpColor
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    Exit Property

ErrSection:
    RaiseError "cIndicator.UpColor.Let"
    
End Property

Public Property Get DownColor() As Long
    DownColor = m.nDownColor
End Property

Public Property Let DownColor(ByVal nColor&)
On Error GoTo ErrSection:
    
    Dim Annot As cAnnotation
    Dim Chart As cChart
    Dim i&
    
    m.nDownColor = nColor

    If IsAutoSwingTrendlines Then
        If Not ActiveChart Is Nothing Then
            Set Chart = ActiveChart.Chart
            For i = 0 To Chart.Annots.Count
                Set Annot = Chart.Annots(i)
                If Not Annot Is Nothing Then
                    If Annot.eUsage = eANNOT_AutoSwingTrend Then
                        If Annot.Prop("Trend") = "Down" Then
                            Annot.Color = m.nDownColor
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    Exit Property

ErrSection:
    RaiseError "cIndicator.DownColor.Let"
    
End Property

Public Property Get IndToColor() As Long

    If IsHawkeyeAdds Then
        If m.nIndToColor <= 0 Then
            If Not m.Chart Is Nothing Then
                If Not m.Chart.Tree Is Nothing Then
                    m.nIndToColor = m.Chart.Tree.Index("PRICE")
                End If
            End If
        End If
    End If
    
    IndToColor = m.nIndToColor
    
End Property

Public Property Get TrueRangeFlag() As Double
    If m.geIndStruct.TrueRangeFlag = kNullData Then
        TrueRangeFlag = 0
    Else
        TrueRangeFlag = 1
    End If
End Property

Public Property Let TrueRangeFlag(ByVal nFlag#)
    m.geIndStruct.TrueRangeFlag = nFlag
End Property

Public Property Get trueRangeColor() As Long
    trueRangeColor = m.geIndStruct.trueRangeColor
End Property

Public Property Let trueRangeColor(ByVal nColor&)
    m.geIndStruct.trueRangeColor = nColor
End Property

Public Property Get UpDownColorFlag() As Long

    If m.eDisplayType = eINDIC_Candlestick Or m.eDisplayType < eINDIC_OHLC Then
        UpDownColorFlag = m.nUpDownColorSpecial
    Else
        UpDownColorFlag = m.nUpDownColorFlag
    End If

End Property

Public Property Let UpDownColorFlag(ByVal nFlag&)
    
    If m.eDisplayType = eINDIC_Candlestick Or m.eDisplayType < eINDIC_OHLC Then
        m.nUpDownColorSpecial = nFlag
        If nFlag = 0 Then
            m.geIndStruct.FillColor = -1
            m.geIndStruct.fillColor2 = -1
        End If
    Else
        m.nUpDownColorFlag = nFlag
    End If
    
End Property

Public Property Get ParmDesc(ByVal nItem&) As String
On Error GoTo ErrSection:

    Dim strDesc$

    If m.Inputs Is Nothing Then
        strDesc = ""
    Else
        If nItem > ParmCount Or nItem < 1 Then
            strDesc = ""
        ElseIf m.eDataType = eINDIC_Constant Then
            strDesc = ""
        Else
            strDesc = m.Inputs.Item(nItem).ParmDesc
        End If
        
        If Len(strDesc) > 0 Then
            strDesc = Replace(strDesc, "|", vbCrLf)
        End If
    End If
    
    ParmDesc = strDesc

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.ParmDesc", eGDRaiseError_Raise
    
End Property

Public Property Get HighlightBarsRight() As Long
    HighlightBarsRight = m.nBoxBarsRight
End Property

Public Property Let HighlightBarsRight(ByVal nCount As Long)
    m.nBoxBarsRight = Abs(nCount)
End Property

Public Property Get HighlightBarsLeft() As Long
    HighlightBarsLeft = m.nBoxBarsLeft
End Property

Public Property Let HighlightBarsLeft(ByVal nCount As Long)
    m.nBoxBarsLeft = Abs(nCount)
End Property

Public Property Get BoxFilLColor() As Long
    BoxFilLColor = m.nBoxFillColor
End Property

Public Property Let BoxFilLColor(ByVal nColor As Long)
    m.nBoxFillColor = nColor
End Property

Public Property Get BoxAnnotationID() As Long
    BoxAnnotationID = m.nBoxAnnotID
End Property

Public Property Get BoxPenStyle() As Long
    BoxPenStyle = m.nBoxPenStyle
End Property

Public Property Let BoxPenStyle(ByVal nStyle As Long)
    m.nBoxPenStyle = nStyle
End Property

Public Property Get BoxFillStyle() As Long
    BoxFillStyle = m.nBoxFillStyle
End Property

Public Property Let BoxFillStyle(ByVal nStyle As Long)
    
    If nStyle = 0 Then
        m.nBoxFillStyle = 0
    Else
        m.nBoxFillStyle = 1
    End If
    
End Property

Private Sub ValidateBoxAnnot(Chart As cChart, Ind As cIndicator, _
    ByVal hXdate&, ByVal hXBar&, _
    ByVal nScreenStartX&, ByVal nScreenEndX&, ByVal hBoolScreenArray&)
On Error GoTo ErrSection:
    
    Dim Annot As cAnnotation
    Dim tblData As cGdTable        'table to hold date1, date2, y1 and y2 for annotation
    Dim i&, j&, k&, nIdxLeft&, nIdxRight&, nEnd&
    Dim strRecord$
    
    Dim eLogFlag As ePANE_LogFlag
    
    Dim dDate1#, dDate2#, dY1#, dY2#
    Dim dMax#, dMin#, dTemp#
    Dim dIndMax#, dIndMin#, dMult#, dOffSet#
    Dim IndBars As cGdBars
    Dim IndData As cGdArray
    Dim AnnotType As eAnnotType
    Dim Pane As cPane
       
    Dim iDataVal&, hIndHigh&, hIndLow&, hIndData&
    Dim bAutoExtendBoxes As Boolean
        
    'remove annotation if necessary
    If m.eDisplayType <> eINDIC_HighlightBoxes And m.eDisplayType <> eINDIC_HighlightZones Then
        If m.nBoxAnnotID > 0 Then Set Annot = Chart.Annots(m.nBoxAnnotID)
        If Not Annot Is Nothing Then
            Annot.geRemoveAnnotation Chart.geChartObj
            Chart.Annots.Remove Annot.Prop("AnnotKey")
            Set Annot = Nothing
        End If
        m.nBoxAnnotID = 0
        m.nBoxBarsLeft = 0
        m.nBoxBarsRight = 0
        Exit Sub
    End If
    
    'double check for negative numbers (theoretically should never happen)
    If m.nBoxBarsLeft < 0 Then m.nBoxBarsLeft = 0
    If m.nBoxBarsRight < 0 Then m.nBoxBarsRight = 0
    
    ' TLB 12/9/2009: we can default back to auto-extending the boxes (like we
    ' used to do long ago) only if none of the data in the array is greater than 1
    ' (since > 1 means the function is designed to explicitly set the length of each box)
    If m.nBoxBarsLeft = 0 And m.nBoxBarsRight = 0 Then
        If m.aData.MaxValue <= 1 Then
            bAutoExtendBoxes = True
        End If
    End If
    
    If Ind.DataType = eINDIC_BarData Then
        Set IndBars = Ind.Bars
    ElseIf Ind.DataType = eINDIC_Array Then
        Set IndData = Ind.Data
    Else
        Exit Sub
    End If
    
    If IndBars Is Nothing And IndData Is Nothing Then Exit Sub
           
    Set Annot = New cAnnotation
    Annot.CreateNew Chart, eANNOT_MultiRects, m.geIndStruct.paneId, 0, 0, 0, 0, , , , eANNOT_Indicator
    i = Chart.Annots.Add(Annot)
    If i > 0 Then
        Annot.geAnnId = i
    Else
        Exit Sub
    End If
    
    If Annot Is Nothing Then Exit Sub
    
    m.nBoxAnnotID = Annot.geAnnId
        
    ' Flag for box height
    '   0: height = min/max of bar(s)
    '   1: height = pane's height
    '   2: height = chart's height (i.e. draw across panes)
    If m.eDisplayType = eINDIC_HighlightBoxes Then
        Annot.Prop("ShowInAllPanes") = 0
    Else
        Annot.Prop("ShowInAllPanes") = m.nShowInAllPanes
    End If

    'build table of rectangle dimension (i.e. top,left,bottom,right info)
    Set tblData = New cGdTable
    
    tblData.CreateField eGDARRAY_Doubles, 0
    tblData.CreateField eGDARRAY_Doubles, 1
    tblData.CreateField eGDARRAY_Doubles, 2
    tblData.CreateField eGDARRAY_Doubles, 3
        
    If Not IndBars Is Nothing Then
        hIndHigh = IndBars.ArrayHandle(eBARS_High)
        hIndLow = IndBars.ArrayHandle(eBARS_Low)
    End If
    If Not IndData Is Nothing Then
        hIndData = IndData.ArrayHandle
    End If
        
    If Ind.Overlayed = True Then
        If Ind.DataType = eINDIC_BarData Then
            dIndMin = gdMinValue(hIndLow, nScreenStartX, nScreenEndX)
            dIndMax = gdMaxValue(hIndHigh, nScreenStartX, nScreenEndX)
        Else
            dIndMin = gdMinValue(hIndData, nScreenStartX, nScreenEndX)
            dIndMax = gdMaxValue(hIndData, nScreenStartX, nScreenEndX)
        End If
        Set Pane = Chart.Tree(Chart.Tree.AncestorIndex(Ind.geIndId, -1))
        If Pane Is Nothing Then
            Exit Sub        'can't do anything, just quit
        Else
            eLogFlag = Pane.PaneLogFlag
            CalcOverlayVars dMult, dOffSet, dIndMin, dIndMax, Pane.gePaneMin, Pane.gePaneMax, eLogFlag
        End If
    End If
    
    nEnd = Chart.LastGoodDataBar(True, True)
    ' TLB 2/11/2013: allow HighlightZones to go beyond the data
    ' (since they could potentially be related only to the dates and not to the data)
    If m.eDisplayType = eINDIC_HighlightZones And nScreenEndX > nEnd Then
        nEnd = nScreenEndX
    End If
    For i = nScreenStartX To nEnd
        ''iDataVal = BooleanValue(m.aData(gdGetNum(hXBar, i)))
        iDataVal = gdGetNum(hBoolScreenArray, i)
        If iDataVal >= 1 Then
            nIdxLeft = i
            nIdxRight = i
            ' only do before/after bars for Boxes
            If m.eDisplayType = eINDIC_HighlightBoxes Then
                If bAutoExtendBoxes Then
                    k = i
                    nIdxLeft = i
                    ''While k < nScreenEndX And m.aData(gdGetNum(hXBar, k + 1)) = 1
                    While k < nScreenEndX And gdGetNum(hBoolScreenArray, k + 1) = 1
                        k = k + 1
                    Wend
                    nIdxRight = k
                    i = k
                Else
                    ' count "n" bars to the right (must be actual good data bars)
                    k = m.nBoxBarsRight  '(# bars to add)
                    Do While k > 0
                        nIdxRight = nIdxRight + 1
                        If nIdxRight > nScreenEndX + 1 Then Exit Do
                        If gdGetNum(hXBar, nIdxRight) > 0 Then
                            k = k - 1
                        End If
                    Loop
                    ' count "n" bars to the left (must be actual good data bars)
                    k = m.nBoxBarsLeft + iDataVal - 1  '(# bars to add)
                    Do While k > 0
                        nIdxLeft = nIdxLeft - 1
                        ' TLB 8/15/2008: allow searching further off-screen (to beginning of box)
                        ''If nIdxLeft < 0 Or nIdxLeft < nScreenStartX - 1 Then Exit Do
                        If nIdxLeft <= 0 Then Exit Do
                        If gdGetNum(hXBar, nIdxLeft) > 0 Then
                            k = k - 1
                        End If
                    Loop
                End If
            End If
            
            dDate1 = gdGetNum(hXdate, nIdxLeft)
            dDate2 = gdGetNum(hXdate, nIdxRight + 1)
                                                
            If Ind.DataType = eINDIC_BarData Then
                dMax = kNullData
                dMin = 999999999
                For j = nIdxLeft To nIdxRight
                    'dTemp = IndBars.Item(eBARS_High, gdGetNum(hXBar, j))
                    dTemp = gdGetNum(hIndHigh, gdGetNum(hXBar, j))
                    If Ind.Overlayed And dTemp <> kNullData Then
                        If eLogFlag = ePANE_LogFlagLog Then
                            dTemp = Exp(dTemp * dMult + dOffSet)
                        Else
                            dTemp = dTemp * dMult + dOffSet
                        End If
                    End If
                    If dTemp > dMax Then dMax = dTemp
                    
                    'dTemp = IndBars.Item(eBARS_Low, gdGetNum(hXBar, j))
                    dTemp = gdGetNum(hIndLow, gdGetNum(hXBar, j))
                    If Ind.Overlayed And dTemp <> kNullData Then
                        If eLogFlag = ePANE_LogFlagLog Then
                            dTemp = Exp(dTemp * dMult + dOffSet)
                        Else
                            dTemp = dTemp * dMult + dOffSet
                        End If
                    End If
                    If dTemp < dMin And dTemp <> kNullData Then
                        dMin = dTemp
                        ' TLB 8/5/2008: for the Lateral Move indicator, just
                        ' set the high/low based on the very first bar's range
                        If UCase(Left(m.strName, 7)) = "LATERAL" Then
                            Exit For
                        End If
                    End If
                Next
            Else
                dMax = gdGetNum(hIndData, gdGetNum(hXBar, nIdxLeft))
                dMin = dMax
                For j = nIdxLeft + 1 To nIdxRight
                    dTemp = gdGetNum(hIndData, gdGetNum(hXBar, j))
                    If dTemp > dMax Then dMax = dTemp
                    If dTemp < dMin Then dMin = dTemp
                Next
            End If
            
            If dMax <> kNullData Then dY2 = dMax
            If dMin <> kNullData Then dY1 = dMin
            
'            strRecord = Trim(Str(dDate1)) & "," & Trim(Str(dDate2)) & ", " & Trim(Str(dY1)) & "," & Trim(Str(dY2))
'            tblData.AddRecord strRecord, -1, ","
            'aardvark 3085 fix
            tblData.NumRecords = tblData.NumRecords + 1
            tblData.Num(0, tblData.NumRecords - 1) = dDate1
            tblData.Num(1, tblData.NumRecords - 1) = dDate2
            tblData.Num(2, tblData.NumRecords - 1) = dY1
            tblData.Num(3, tblData.NumRecords - 1) = dY2
        End If
    Next
        
    Annot.Color = m.nColor
    Annot.Style = m.nBoxPenStyle
    Annot.Prop("FillColor") = m.nBoxFillColor
    Annot.Prop("FillPattern") = m.nBoxFillStyle
    
    Annot.SetMultiRectsDim Chart, nScreenStartX, tblData        '4525
        
    Set tblData = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.ValidateBoxAnnot"
    
End Sub

Public Property Get OkayToMirror() As Boolean

    Dim bOkay As Boolean
    
    bOkay = False
    
    If m.eDisplayType = eINDIC_BollingerBar Or m.eDisplayType = eINDIC_Candlestick Or _
       m.eDisplayType = eINDIC_HL Or m.eDisplayType = eINDIC_HLC Or m.eDisplayType = eINDIC_OHLC Or _
       m.eDisplayType = eINDIC_Line Or m.eDisplayType = eINDIC_Points Or m.eDisplayType = eINDIC_Steps Then
       
       bOkay = True
       
    End If
    
    OkayToMirror = bOkay

End Property

Public Property Get ShowInAllPanes() As Long
    ShowInAllPanes = m.nShowInAllPanes
End Property

Public Property Let ShowInAllPanes(ByVal nShow&)
    m.nShowInAllPanes = nShow
End Property

Public Property Get ColorSeperatorVal() As Double
    ColorSeperatorVal = m.dColorSeperator
End Property

Public Property Let ColorSeperatorVal(ByVal dValue#)
    m.dColorSeperator = dValue
End Property

Public Property Get HistogramColorBelow() As Long
    If m.nColorBelow >= 0 Then
        HistogramColorBelow = m.nColorBelow
    Else
        HistogramColorBelow = m.nBoxFillColor
    End If
End Property

Public Property Let HistogramColorBelow(ByVal nColor&)
    m.nColorBelow = nColor
End Property

Public Property Get IsBiColorHistogram() As Long
On Error GoTo ErrSection:
    
    Dim nIsBiColor&
    
    If m.eDisplayType = eINDIC_Area Or m.eDisplayType = eINDIC_Histogram Or m.eDisplayType = eINDIC_ClusterTime Then
        If m.nColorBelow >= 0 Then
            nIsBiColor = 1
        End If
    End If
    
    IsBiColorHistogram = nIsBiColor
    
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.IsBiColorHistogram.Get"
    
End Property

Public Property Get BaseLineY() As Double
On Error GoTo ErrSection:

    If m.dUserBaseLineY = kNullData Then
        'custom base line was never set, return what is currently in use
        BaseLineY = m.geIndStruct.y_baseline
    Else
        BaseLineY = m.dUserBaseLineY
    End If
    
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.BaseLineY.Get"
    
End Property

Public Property Let BaseLineY(ByVal dBase#)
On Error GoTo ErrSection:

    m.dUserBaseLineY = dBase
    
    Exit Property

ErrSection:
    RaiseError "cIndicator.BaseLineY.Let"
    
End Property

Private Sub AddNewAutoTrend(Chart As cChart, _
    aSwingHigh As cGdArray, aSwingLow As cGdArray, hXBar&, _
    strAnnotUpKey$, strAnnotDownKey$, nStrengthHigh&, nStrengthLow&, _
    ByVal paneId&, ByVal dDataBarsAgo#, ByVal i&, ByVal nStrength&, _
    ByVal nScreenStartX&, ByVal nScreenEndX&, ByVal nPointNum&, _
    Optional ByVal bDoLow As Boolean = False)
On Error GoTo ErrSection:
    
    Dim k&, n&, nCount&
    Dim strTrend$, strKey$, nColor&
    Dim dY#, dData#, dHigh#, dLow#
    Dim dDate1#, dDate2#, dDateEnd#
    Dim bAddNew As Boolean
    Dim aSwingArrayToUse As cGdArray
    Dim PrevAnnot As cAnnotation, Annot As cAnnotation
    Dim bSwingHigh As Boolean, bSwingLow As Boolean
    
    Dim eStyle As eIndicatorStyle
    
    dHigh = Chart.Bars.Item(eBARS_High, gdGetNum(hXBar, i - nStrength))
    dLow = Chart.Bars.Item(eBARS_Low, gdGetNum(hXBar, i - nStrength))
    
    'determine whether we are processing a swing high or swing low data bar
    If bDoLow Then
        'spread charts have high=low on every bar - aardvark 4158
        'when strength is the same, this is called twice to get both up & down trend lines
        If dDataBarsAgo = dLow Then
            bSwingLow = True
        ElseIf dDataBarsAgo = dHigh Then
            bSwingHigh = True
        End If
    Else
        If ParmCount = 3 And Parm(2) <> Parm(3) Then
            If nStrength = Parm(3) And dDataBarsAgo = dLow Then
                bSwingLow = True
            ElseIf nStrength = Parm(2) And dDataBarsAgo = dHigh Then
                bSwingHigh = True
            End If
        ElseIf dDataBarsAgo = dHigh Then
            bSwingHigh = True
        ElseIf dDataBarsAgo = dLow Then
            bSwingLow = True
        End If
    End If
            
    'set values for array holding high/low swingpoints
    If bSwingLow Then
        strTrend = "Up"
        For k = 8 To 0 Step -1
            aSwingLow(k + 1) = aSwingLow(k)
        Next
        nStrengthLow = nStrength
        aSwingLow(0) = i - nStrength
        Set aSwingArrayToUse = aSwingLow
        Set PrevAnnot = Chart.Annots(strAnnotUpKey)
    ElseIf bSwingHigh Then
        strTrend = "Down"
        For k = 8 To 0 Step -1
            aSwingHigh(k + 1) = aSwingHigh(k)
        Next
        nStrengthHigh = nStrength
        aSwingHigh(0) = i - nStrength
        Set aSwingArrayToUse = aSwingHigh
        Set PrevAnnot = Chart.Annots(strAnnotDownKey)
    Else
        Exit Sub
    End If
    
    'double check
    If aSwingArrayToUse Is Nothing Then
        Exit Sub
    End If

    'add new trend line
    For k = 0 To 9
        bAddNew = False
        If bSwingLow Then
            dData = Chart.Bars.Item(eBARS_Low, gdGetNum(hXBar, aSwingArrayToUse(k)))
            If dData < dDataBarsAgo And dData <> kNullData Then
                bAddNew = True
            End If
        ElseIf bSwingHigh Then
            dData = Chart.Bars.Item(eBARS_High, gdGetNum(hXBar, aSwingArrayToUse(k)))
            If dData > dDataBarsAgo Then
                bAddNew = True
            End If
        Else
            Exit For
        End If
        
        If bAddNew Then
            dDate1 = Chart.Bars.Item(eBARS_DateTime, gdGetNum(hXBar, aSwingArrayToUse(k)))
            dDate2 = Chart.Bars.Item(eBARS_DateTime, gdGetNum(hXBar, aSwingArrayToUse(0)))
            'If dDate1 > Chart.aXdate(Chart.ScreenEndX) Or dDate2 < Chart.aXdate(Chart.ScreenStartX) Then
            '    bAddNew = False
            If dDate1 = kNullData Or dDate2 = kNullData Then
                bAddNew = False
            End If
        End If

        If bAddNew Then
            'truncate previous annot of the same trend if not already truncated
            If Not PrevAnnot Is Nothing Then
                If PrevAnnot.Prop("Truncated") = 0 And PrevAnnot.Prop("Trend") = strTrend Then
                    dDateEnd = Chart.Bars.Item(eBARS_DateTime, gdGetNum(hXBar, aSwingArrayToUse(0) + nStrength))
                    n = nPointNum
                    If n >= 0 Then
                        dY = geTrendValueY(PrevAnnot.geAnnotObject, n)
                        If dY >= 0 Then
                            If m.nExtendTrend > 0 Then
                                dDateEnd = Chart.Bars.Item(eBARS_DateTime, gdGetNum(hXBar, aSwingArrayToUse(0) + nStrength + m.nExtendTrend))
                                n = n + m.nExtendTrend
                                dY = geTrendValueY(PrevAnnot.geAnnotObject, n)
                            End If
                            
                            PrevAnnot.dDate(2) = dDateEnd
                            PrevAnnot.Y(2) = dY
                            PrevAnnot.Prop("Truncated") = 1
                            PrevAnnot.Prop("Ext") = 0
                        End If
                    End If
                End If
                
                If PrevAnnot.Prop("Trend") = strTrend Then
                    If m.nShowTrendHistory = 0 Or _
                        PrevAnnot.dDate(1) > Chart.aXdate(nScreenEndX) Or _
                        PrevAnnot.dDate(2) < Chart.aXdate(nScreenStartX) Then
                        
                        PrevAnnot.geRemoveAnnotation Chart.geChartObj
                        Chart.Annots.Remove (PrevAnnot.Prop("AnnotKey"))
                        Set PrevAnnot = Nothing
                        
                    End If
                End If
            End If

            'add new annot
            If strTrend = "Up" Then
                nColor = m.nUpColor
            ElseIf strTrend = "Down" Then
                nColor = m.nDownColor
            Else
                nColor = m.nColor
            End If

            Set Annot = New cAnnotation
            Annot.CreateNew Chart, eANNOT_Trendline, paneId, dDate1, dData, dDate2, dDataBarsAgo, nColor, , , eANNOT_AutoSwingTrend
            
            If m.eStyle = eINDIC_Default Then       '5229
                eStyle = g.ChartGlobals.eDefaultIndStyle
            Else
                eStyle = m.eStyle
            End If
            
            Select Case eStyle
                Case eINDIC_Dash
                    Annot.Style = eANNOT_DashLg
                Case eINDIC_Dot
                    Annot.Style = eANNOT_DashSm
                Case eINDIC_DashDot
                    Annot.Style = eANNOT_DashDot
                Case Else
                    Annot.Style = Int(eStyle) * -1
            End Select
            
            Annot.Prop("Ext") = 0
            Annot.Prop("ExtColor") = Annot.Color
            Annot.Prop("Truncated") = 0
            Annot.Prop("Trend") = strTrend
            Annot.geIndId = m.geIndStruct.indicatorId
            Annot.geAnnId = Chart.Annots.Add(Annot)
            strKey = Chart.Annots.Key(Annot.geAnnId)
            Annot.Prop("AnnotKey") = strKey
                        
            If strTrend = "Up" Then
                strAnnotUpKey = strKey
            ElseIf strTrend = "Down" Then
                strAnnotDownKey = strKey
            End If
            Exit For
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.AddNewAutoTrend", eGDRaiseError_Raise

End Sub

Private Sub SetAutoTrend(Chart As cChart, _
    ByVal paneId&, ByVal nID&, ByVal nFirstBar&, ByVal nLastBar&, _
    ByVal nScreenStartX&, ByVal nScreenEndX&, ByVal dExtraPriceScale#)
On Error GoTo ErrSection:
        
    Static nSymID&
        
    Dim i&, j&, k&, n&, nStart&, nEnd&, rc&
    Dim hXBar&
    Dim strUpKey$, strDownKey$
    Dim dData#, dDataBarsAgo#, dClose#, dMinMove#
    Dim dHigh#, dLow#
    Dim nStrengthHigh&, nStrengthLow&
    
    Dim Ind As cIndicator
    Dim AnnotUp As cAnnotation, AnnotDown As cAnnotation
    Dim dDateStart#, dDateEnd#, dY#
    
    Dim aSwingHigh As New cGdArray      'holds index of bars that have high swing points
    Dim aSwingLow As New cGdArray       'holds index of bars that have low swing points
    Dim nPointNum&, nLoop&

    Dim oFunction As cFunction
    Dim bSwingHigh As Boolean, bSwingLow As Boolean
    
    Dim bSpreadChart As Boolean

    If nSymID <> Chart.SymbolID Then Chart.geForceRecalc

    nStrengthHigh = -1
    nStrengthLow = -1
    
    hXBar = Chart.aXBar.ArrayHandle
        
    dDateStart = Chart.aXdate(nScreenStartX)
    dDateEnd = Chart.aXdate(nScreenEndX)
    
    If Len(Chart.SpreadSymbols) > 0 Then bSpreadChart = True
            
    'set min move value for rounding
    dMinMove = Chart.Bars.MinMove(dDateEnd)
    
    'initialize all elements
    aSwingHigh.Create eGDARRAY_Longs, 10
    aSwingLow.Create eGDARRAY_Longs, 10
    For i = 0 To 9
        aSwingHigh(i) = kNullData
        aSwingLow(i) = kNullData
    Next
    
    If m.nFunctionID <= 0 Then
        nStart = nScreenStartX
        nEnd = nScreenEndX
    Else
        Set oFunction = New cFunction
        With oFunction
            .FunctionID = m.nFunctionID
            .Load
            If InStr(.CodedText, "SwingSystemHigh") Then bSwingHigh = True
            If InStr(.CodedText, "SwingSystemLow") Then bSwingLow = True
        End With
        nStart = GetStartX(Chart, hXBar, nScreenStartX, bSwingHigh, bSwingLow)
        nEnd = GetEndX(Chart, hXBar, nScreenEndX, bSwingHigh, bSwingLow)
    End If
    
    If m.nShowTrendHistory = 0 And nEnd < Chart.LastGoodDataBar(False) Then
        Exit Sub
    End If
        
    Chart.RemoveAnnots False, , eANNOT_AutoSwingTrend

    Set AnnotUp = New cAnnotation
    nPointNum = AnnotUp.gePointNum(Chart, Chart.aXdate(nStart))
    Set AnnotUp = Nothing
    nLoop = -1
        
    Set Ind = Chart.Tree("PRICE")
    For i = nStart To nEnd
        nLoop = nLoop + 1
        dData = m.aData(gdGetNum(hXBar, i))
        If i >= nScreenStartX And i <= nScreenEndX Then
            gdSetNum m.geIndStruct.gdhData1, j, dData
            gdSetNum m.geIndStruct.gdhYScaleVal, j, dData
            If dData <> kNullData Then
                If Not Ind Is Nothing And m.nColorPriceInd = 1 Then
                    dHigh = Chart.Bars.Item(eBARS_High, gdGetNum(hXBar, i))
                    dLow = Chart.Bars.Item(eBARS_Low, gdGetNum(hXBar, i))
                    If dData = dHigh Then
                        Ind.geSetPenColor j, m.nDownColor
                    ElseIf dData = dLow Then
                        Ind.geSetPenColor j, m.nUpColor
                    End If
                End If
            End If
            j = j + 1
        End If
            
        Set AnnotUp = Chart.Annots(strUpKey)
        Set AnnotDown = Chart.Annots(strDownKey)
        
        'see if there was a swing point nStrength bars ago
        dDataBarsAgo = m.aData(gdGetNum(hXBar, i - Parm(2)))

        If dDataBarsAgo <> kNullData Then
            AddNewAutoTrend Chart, aSwingHigh, aSwingLow, hXBar, strUpKey, strDownKey, _
                nStrengthHigh, nStrengthLow, paneId, dDataBarsAgo, i, Parm(2), _
                nScreenStartX, nScreenEndX, nPointNum + nLoop
        End If
        If ParmCount = 3 Then
            If bSpreadChart And Parm(2) = Parm(3) Then
                AddNewAutoTrend Chart, aSwingHigh, aSwingLow, hXBar, strUpKey, strDownKey, _
                    nStrengthHigh, nStrengthLow, paneId, dDataBarsAgo, i, Parm(2), _
                    nScreenStartX, nScreenEndX, nPointNum + nLoop, True         '4158
            ElseIf Parm(2) <> Parm(3) Then
                dDataBarsAgo = m.aData(gdGetNum(hXBar, i - Parm(3)))
                If dDataBarsAgo <> kNullData Then
                    AddNewAutoTrend Chart, aSwingHigh, aSwingLow, hXBar, strUpKey, strDownKey, _
                    nStrengthHigh, nStrengthLow, paneId, dDataBarsAgo, i, Parm(3), _
                    nScreenStartX, nScreenEndX, nPointNum + nLoop
                End If
            End If
        End If
                
        'truncate if close is higher/lower than trend line value
        If Not AnnotUp Is Nothing Then
            TruncateTrend Chart, AnnotUp, hXBar, nScreenStartX, nPointNum + nLoop, dMinMove, nStrengthLow, i
        End If
        If Not AnnotDown Is Nothing Then
             rc = TruncateTrend(Chart, AnnotDown, hXBar, nScreenStartX, nPointNum + nLoop, dMinMove, nStrengthHigh, i)
        End If
    Next

    'add extension to last trend line(s) if applicable
    If Not AnnotUp Is Nothing Then
        If AnnotUp.Prop("Truncated") = 0 Then AnnotUp.Prop("Ext") = 1
    End If
    
    If Not AnnotDown Is Nothing Then
        If AnnotDown.Prop("Truncated") = 0 Then AnnotDown.Prop("Ext") = 1
    End If
    
    aSwingHigh.Destroy
    aSwingLow.Destroy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.SetAutoTrend", eGDRaiseError_Raise

End Sub

Private Function TruncateTrend(Chart As cChart, Annot As cAnnotation, _
    hXBar&, ByVal nScreenStartX&, ByVal nPointNum&, _
    ByVal dMinMove#, ByVal nStrength&, ByVal i&) As Long
On Error GoTo ErrSection:

'{ if most recent DnTL/UpTL exists AND has not yet been truncated here AND was drawn
'  at least BarsPast ago AND was breached BarsPast bars ago THEN truncate it here and
'  set break flag }
    
    Dim dDateEnd#, dY#
    Dim n&, nIdx&, nEnd&, nExtend&
    Dim dClose#, dClose1#, dClose2#
        
    If nStrength < 0 Or Annot.Prop("Truncated") = 1 Then
        Exit Function
    End If
    
    nIdx = i
    
    With Annot
        nEnd = .X(2) - nScreenStartX
        If .Prop("Truncated") = 0 Then
            dDateEnd = Chart.Bars.Item(eBARS_DateTime, gdGetNum(hXBar, nIdx))
            If dDateEnd > 0 Then
                n = nPointNum
            Else
                n = -1
            End If
            If n >= 0 Then
                dY = geTrendValueY(.geAnnotObject, n)
                dY = Int(dY / dMinMove + 0.5) * dMinMove
                dClose = Chart.Bars.Item(eBARS_Close, gdGetNum(hXBar, nIdx))
                If (.Prop("Trend") = "Down" And dClose > dY) Or _
                   (.Prop("Trend") = "Up" And dClose < dY) Then
                    If m.nExtendTrend > 0 Then
                        nExtend = nEnd + nStrength + m.nExtendTrend + 1
                        dDateEnd = Chart.Bars.Item(eBARS_DateTime, gdGetNum(hXBar, nIdx + m.nExtendTrend))
                        n = .gePointNum(Chart, dDateEnd)
                        dY = geTrendValueY(.geAnnotObject, n)
                        dY = Int(dY / dMinMove + 0.5) * dMinMove
                    End If
                    If dY >= 0 And dDateEnd <> kNullData And dClose <> kNullData Then
                        .dDate(2) = dDateEnd
                        .Y(2) = dY
                        .Prop("Truncated") = 1
                        .Prop("Ext") = 0
                        'check for qualifying trend lines
                        If HasModule("INC,LWST,PRG11") Then
                            If .Prop("Trend") = "Down" Then
                                If nIdx - 2 > 0 Then
                                    dClose1 = Chart.Bars.Item(eBARS_Close, gdGetNum(hXBar, nIdx - 1))
                                    dClose2 = Chart.Bars.Item(eBARS_Close, gdGetNum(hXBar, nIdx - 2))
                                    If dClose1 < dClose2 Then
                                        Annot.Style = eANNOT_DashSm
                                    End If
                                End If
                            ElseIf .Prop("Trend") = "Up" Then
                                If nIdx - 2 > 0 Then
                                    dClose1 = Chart.Bars.Item(eBARS_Close, gdGetNum(hXBar, nIdx - 1))
                                    dClose2 = Chart.Bars.Item(eBARS_Close, gdGetNum(hXBar, nIdx - 2))
                                    If dClose1 > dClose2 Then
                                        Annot.Style = eANNOT_DashSm
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
    
    TruncateTrend = dY
    
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.TruncateTrend"
    
End Function

Private Function GetStartX(Chart As cChart, hXBar&, ByVal nScreenStartX&, _
    ByVal bSwingHigh As Boolean, ByVal bSwingLow As Boolean) As Long
On Error GoTo ErrSection:

    Dim bBoth As Boolean
    Dim nCountHigh&, nCountLow&, i&
    Dim dData#, dHigh#, dLow#
        
    If bSwingHigh And bSwingLow Then bBoth = True
    
    For i = nScreenStartX To 0 Step -1
        dData = m.aData(gdGetNum(hXBar, i))
        If dData <> kNullData Then
            dHigh = Chart.Bars.Item(eBARS_High, gdGetNum(hXBar, i))
            dLow = Chart.Bars.Item(eBARS_Low, gdGetNum(hXBar, i))
            If dData = dHigh Then
                nCountHigh = nCountHigh + 1
            ElseIf dData = dLow Then
                nCountLow = nCountLow + 1
            Else
                nCountHigh = nCountHigh + 1
                nCountLow = nCountLow + 1
            End If
        End If
        
        If bBoth Then
            If nCountHigh >= 10 And nCountLow >= 10 Then Exit For
        ElseIf bSwingHigh Then
            If nCountHigh >= 10 Then Exit For
        ElseIf bSwingLow Then
            If nCountLow >= 10 Then Exit For
        Else
            Exit For
        End If
    Next

    If i < 0 Then i = 0
    If bSwingLow Or bSwingHigh Then
        GetStartX = i
    Else
        GetStartX = nScreenStartX
    End If
    
    Exit Function

ErrSection:
    RaiseError "cIndicator.GetStartX"
    
End Function

Private Function GetEndX(Chart As cChart, hXBar&, ByVal nScreenEndX&, _
    ByVal bSwingHigh As Boolean, bSwingLow As Boolean) As Long
On Error GoTo ErrSection:

    Dim oFunction As New cFunction
    Dim bBoth As Boolean
    Dim nCountHigh&, nCountLow&, i&
    Dim dData#, dHigh#, dLow#
    
    If bSwingHigh And bSwingLow Then bBoth = True
    
    For i = nScreenEndX To m.aData.Size - 1
        dData = m.aData(gdGetNum(hXBar, i))
        If dData <> kNullData Then
            dHigh = Chart.Bars.Item(eBARS_High, gdGetNum(hXBar, i))
            dLow = Chart.Bars.Item(eBARS_Low, gdGetNum(hXBar, i))
            If dData = dHigh Then
                nCountHigh = nCountHigh + 1
            ElseIf dData = dLow Then
                nCountLow = nCountLow + 1
            Else
                nCountHigh = nCountHigh + 1
                nCountLow = nCountLow + 1
            End If
        End If

        If bBoth Then
            If nCountHigh >= 10 And nCountLow >= 10 Then Exit For
        ElseIf bSwingHigh Then
            If nCountHigh >= 10 Then Exit For
        ElseIf bSwingLow Then
            If nCountLow >= 10 Then Exit For
        Else
            Exit For
        End If
    Next

    If bSwingLow Or bSwingHigh Then
        GetEndX = i
    Else
        GetEndX = nScreenEndX
    End If
    
    Exit Function

ErrSection:
    RaiseError "cIndicator.GetEndX"
    
End Function

Public Function IsAutoSwingTrendlines() As Boolean
On Error GoTo ErrSection:

    Dim strCodedText As String
    
    If m.nFunctionID > 0 Then
        If g.Functions.Found(Str(m.nFunctionID)) Then
            strCodedText = g.Functions.Item(Str(m.nFunctionID)).CodedText
            If InStr(strCodedText, "SwingSystemHigh") Or InStr(strCodedText, "SwingSystemLow") Then
                If InStr(UCase(m.strCodedText), "TRENDLINES") Then
                    IsAutoSwingTrendlines = True
                End If
            End If
        End If
    End If
    
    Exit Function

ErrSection:
    RaiseError "cIndicator.IsAutoSwingTrendlines"
    
End Function

Public Property Get ShowTrendHistory() As Long
    ShowTrendHistory = m.nShowTrendHistory
End Property

Public Property Let ShowTrendHistory(ByVal nShow&)
    m.nShowTrendHistory = nShow
End Property

Public Property Get ExtendTrend() As Long
    ExtendTrend = m.nExtendTrend
End Property

Public Property Let ExtendTrend(ByVal nBars&)
    If nBars >= 0 Then
        m.nExtendTrend = nBars
    Else
        m.nExtendTrend = 0
    End If
End Property

Public Property Get MarkerSize() As Long
    If m.nMarkerSize = 0 Then
        MarkerSize = 1
    ElseIf m.nMarkerSize > 3 Then
        m.nMarkerSize = 3
        MarkerSize = 3
    Else
        MarkerSize = m.nMarkerSize
    End If
End Property

Public Property Let MarkerSize(ByVal nSize&)
    If nSize > 0 And nSize < 4 Then m.nMarkerSize = nSize
End Property

Public Property Get IsAlert() As Boolean
    IsAlert = m.bAlert
End Property

Public Property Get AlertAdded() As Boolean
    AlertAdded = m.bAlertAdded
End Property

Public Property Let AlertAdded(ByVal bAdded As Boolean)
    m.bAlertAdded = bAdded
End Property

Public Sub InitAlert(Chart As cChart, Optional ByVal bClear As Boolean = False)
On Error GoTo ErrSection:
    
    Dim aAlertStr As cGdArray
    Dim strSymInMsgField$
    
    If m.oAlert Is Nothing Then Set m.oAlert = New cAlert

    m.bAlert = True
    m.bDisplay = False
    If Len(m.strName) = 0 Then m.strName = Chart.Symbol
    Set m.Chart = Chart
    
    If bClear Then m.strAlert = ""
    
    With m.oAlert
        If Len(m.strAlert) > 0 Then
            Set aAlertStr = New cGdArray            '01-06-2010 - improved efficiency fix for 3052
            aAlertStr.SplitFields m.strAlert, "|"
            
            strSymInMsgField = Parse(aAlertStr(3), " ", 1)
            
            aAlertStr(1) = Chart.Symbol
            aAlertStr(2) = Chart.SymbolID
            
            'aardvark 6869, 6870
            'the default text is "<symbol name> Chart Alert ..."
            'if the first word is a valid symbol name and it does not match chart symbol then replace it
            'if first word of custom alert msg is valid symbol name this code will also change that
            If strSymInMsgField <> Chart.Symbol Then
                If g.SymbolPool.SymbolIDforSymbol(strSymInMsgField) <> 0 Then
                    If strSymInMsgField <> Chart.Symbol Then
                        aAlertStr(3) = Replace(aAlertStr(3), strSymInMsgField, Chart.Symbol)
                    End If
                End If
            End If
            
            m.strAlert = aAlertStr.JoinFields("|")
            .AlertType = eGDAlertType_Chart
            .FromFileString m.strAlert
            
            Set aAlertStr = Nothing
        Else
            .Symbol = Chart.Symbol
            .SymbolID = Chart.SymbolID
            .AlertType = eGDAlertType_Chart
            .Active = True
            m.strExpression = Replace(m.strExpression, vbCrLf, "")
            .ChartCondition = m.strName & ":" & Space(3) & m.strExpression
        End If
        .Indicator = Me
    End With
    
    If FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid
    
    Exit Sub

ErrSection:
    RaiseError "cIndicator.InitAlert"
    
End Sub

Public Property Get AlertObject() As cAlert
    Set AlertObject = m.oAlert
End Property

Public Sub UpdateAlert(ByVal nAddRemove&, Optional ByVal bRefresh As Boolean = True)
On Error GoTo ErrSection:
'nAddRemove: 0=remove, 1=add, 2=just update alert string, 3=cached page
    
    Dim i&
    
    If Not m.oAlert Is Nothing Then
        If nAddRemove = 0 Then
            m.strAlert = ""
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
        Else
            m.strAlert = m.oAlert.ToFileString
        End If
        If nAddRemove < 2 Then AlertsCollection nAddRemove, bRefresh
        
        If Not m.Chart Is Nothing And bRefresh Then
            m.bConfirmAlert = True
            If nAddRemove = 2 Then
                If m.oAlert.Active Then
                    m.Chart.GenerateChart eRedo5_RecalcInd      'aardvark 3650 fix
                Else
                    m.Chart.GenerateChart eRedo3_Settings
                End If
            Else
                m.Chart.GenerateChart eRedo3_Settings
            End If
        End If
    End If
    
    Exit Sub
        
ErrSection:
    RaiseError "cIndicator.UpdateAlert"
    
End Sub

Private Sub AlertsCollection(ByVal nAddRemove&, ByVal bRefreshForm As Boolean)
On Error GoTo ErrSection:
    
    Dim i&, j&
    Dim bFound As Boolean, bAdded As Boolean
    Dim Alert As cAlert
    
    bAdded = m.bAlertAdded
    If Not m.oAlert Is Nothing Then
        'check for existing alert in collection
        For i = 1 To g.Alerts.Count
            Set Alert = g.Alerts(i)
            If Not Alert Is Nothing Then
                If Alert.ChartAlertId = m.oAlert.ChartAlertId Then
                    bFound = True
                    Exit For
                End If
            End If
        Next
        If nAddRemove = 0 Then
            If bFound Then g.Alerts.Remove (i)
            If Not m.Chart Is Nothing Then
                'fix for art alerts not getting removed from alert manager form
                'oritinal code used: "AutoKey# " & Me.geIndId (no good cause geIndId is zero)
                If m.Chart.Tree(m.strMyKey) Is Me Then
                    m.Chart.Tree.Remove m.strMyKey
                Else
                    'fail safe check in case the indicator key got goofed somehow
                    For j = 1 To m.Chart.Tree.Count
                        If Not m.Chart.Tree(j) Is Nothing Then
                            If m.Chart.Tree(j) Is Me Then
                                m.Chart.Tree.Remove j
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
            Set m.oAlert = Nothing
            m.bAlertAdded = False
        ElseIf Not bFound Then
            g.Alerts.Add m.oAlert
            m.oAlert.Indicator = Me
            m.oAlert.ChartAlertId = RoundNum(CDbl(Now), 5) + g.Alerts.Count 'aardvark 3080
            m.bAlertAdded = True
        End If
    End If
    
    If bRefreshForm Then
        If bAdded <> m.bAlertAdded Then
            If FormIsLoaded("frmAlertsSetup") Then
                If frmAlertsSetup.Visible Then frmAlertsSetup.LoadGrid
            End If
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.AlertsCollection"
    
End Sub

Public Sub CheckIndAlert()
On Error GoTo ErrSection:

    Dim nBar&
    
    If Not m.oAlert Is Nothing And Not m.Chart Is Nothing Then
        If m.oAlert.Active Then
            nBar = m.Chart.LastGoodDataBar(False)
            If m.aData(nBar) = 1 Then
                m.oAlert.CheckAlert , , True, m.bConfirmAlert          'aardvark 3650 fix
            Else
                m.oAlert.CheckAlert , , False
            End If
            m.strAlert = m.oAlert.ToFileString          'update string
            m.bConfirmAlert = False                     '4307
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cIndicator.CheckIndAlert"
    
End Sub

Public Property Get IndChart() As cChart
    Set IndChart = m.Chart
End Property

Public Property Get CondBuilderStr() As String
    CondBuilderStr = m.strCondBuilder
End Property

Public Property Let CondBuilderStr(ByVal strText$)
    m.strCondBuilder = strText
End Property

Private Property Get LastIndValue() As Double
On Error GoTo ErrSection:

    Dim i&, j&, dData#
    
    dData = kNullData           'to indicate no data available
    
    'only implementing for array type right now - can add others later
    If m.eDataType = eINDIC_Array Then
        If Not m.Chart Is Nothing Then
            j = m.Chart.aXBar(m.Chart.ScreenEndX)
            If j > m.Chart.LastGoodDataBar(False) Then j = m.Chart.LastGoodDataBar(False)
            dData = m.aData(j)
        End If
    End If
        
    LastIndValue = dData
    
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.LastIndValue.Get"
    
End Property

Public Function IsWoodInd(ByVal bRetName As Boolean, ByVal bRetLastValue As Boolean, _
    Optional strName As String, Optional dData As Double, Optional ChartIn As cChart = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim s$, s1$
    Dim ChartInUse As cChart
    
    Set ChartInUse = ChartIn                    'fix for Sidewinder not showing up on Woodies CCI template
    If ChartInUse Is Nothing Then Set ChartInUse = m.Chart
    
    If ChartInUse Is Nothing Then Exit Function
    
    If ChartInUse.ShowSplitPane = 1 Then
        s = UCase(m.strCodedName)
        s1 = UCase(m.strExpression)
        
        If Len(s) > 0 Then          '4218
            If InStr(s, "EMA ANGLE") Or InStr(s1, "EMA ANGLE") Or _
               InStr(s, "SIDEWINDER") Or InStr(s1, "SIDEWINDER") Or _
               InStr(s, "WOODIESPIVOT") Or InStr(s1, "WOODIES PIVOT") Or _
               InStr(s, "WOODIESR1") Or InStr(s1, "WOODIES R1") Or _
               InStr(s, "WOODIESR2") Or InStr(s1, "WOODIES R2") Or _
               InStr(s, "WOODIESR3") Or InStr(s1, "WOODIES R3") Or _
               InStr(s, "WOODIESS1") Or InStr(s1, "WOODIES S1") Or _
               InStr(s, "WOODIESS2") Or InStr(s1, "WOODIES S2") Or _
               InStr(s, "WOODIESS3") Or InStr(s1, "WOODIES S3") Then
               
               If bRetName Then strName = s
               If bRetLastValue Then dData = LastIndValue
               IsWoodInd = True
            
            ElseIf s = "MOVINGAVGX" And (InStr(s1, "34") <> 0 And InStr(s1, "CLOSE") <> 0) Then
            
               If bRetName Then strName = m.strCodedName
               If bRetLastValue Then dData = LastIndValue
               IsWoodInd = True
            
            End If
        End If
    End If

    Exit Function
    
ErrSection:
    RaiseError "cIndicator.IsWoodInd"
    
End Function

Public Property Get IsDinapoliDMA() As Boolean
On Error GoTo ErrSection:

    Dim bRC As Boolean

    If UCase(m.strCodedName) = "MOVINGAVG" Then
        If ParmCount = 2 Then
            If Parm(2) = 3 Then
                If InStr(UCase(Parm(1)), "CLOSE OF MARKET") > 0 Then
                    If m.nShiftBars = 3 Then bRC = True
                End If
            End If
        End If
    End If
    
    IsDinapoliDMA = bRC
    
    Exit Property
    
ErrSection:
    RaiseError "cIndicator.IsDinapoliDMA"
    
End Property

Public Function WhatIfMove(ByVal dY#) As Boolean
On Error GoTo ErrSection:

    Dim i&, iLastBar&
    Dim dHigh#, dLow#, dDiff#
    Dim bMoved As Boolean
    
    Dim Pane As cPane
    Dim Annot As cAnnotation
    Dim ePeriodType As eBarsPeriodType

    If m.geIndStruct.isPriceInd <> 1 Then Exit Function
    
    Set Pane = m.Chart.Tree("PRICE PANE")
    If Pane Is Nothing Then Exit Function
    
    Set Annot = m.Chart.Annots("WhatIf")
    If Annot Is Nothing Then Exit Function
    
    If dY > Pane.geAdjustMax Or dY < Pane.geAdjustMin Then Exit Function
    
    iLastBar = m.Chart.LastGoodDataBar(False)
    ePeriodType = m.Bars.Prop(eBARS_PeriodType)
    
    bMoved = True
    If Annot.Prop("WhatIfBar") <> "N" Then
        'only need this check if user is doing what-if with last data bar
        If ePeriodType = ePRD_IntBreakout Or ePeriodType = ePRD_EodBreakout Then
            If m.Bars.Prop(eBARS_TickMove) > 0 Then
                dHigh = m.Bars(eBARS_High, iLastBar)
                dLow = m.Bars(eBARS_Low, iLastBar)
                If dY > dHigh Then
                    dDiff = (dY - dLow) / m.Bars.Prop(eBARS_TickMove)
                    If dDiff > m.Bars.Prop(eBARS_PeriodsPerBar) Then dY = dHigh
                ElseIf dY < dLow Then
                    dDiff = (dHigh - dY) / m.Bars.Prop(eBARS_TickMove)
                    If dDiff > m.Bars.Prop(eBARS_PeriodsPerBar) Then dY = dLow
                End If
            Else
                bMoved = False
            End If
        End If
    End If
        
    If bMoved Then
        bMoved = False          'reset
        m.Bars(eBARS_Close, iLastBar) = dY
        m.Bars.Prop(eBARS_LastTickTime) = m.Bars.Prop(eBARS_LastTickTime) + 1
        If Annot.Prop("WhatIfBar") = "N" Then
            m.Bars(eBARS_Open, iLastBar) = dY
            m.Bars(eBARS_High, iLastBar) = dY
            m.Bars(eBARS_Low, iLastBar) = dY
            For i = gdGetSize(m.geIndStruct.gdhData4) - 1 To 0 Step -1
                If gdGetNum(m.geIndStruct.gdhData4, i) <> kNullData Then
                    gdSetNum m.geIndStruct.gdhData1, i, dY
                    gdSetNum m.geIndStruct.gdhData2, i, dY
                    gdSetNum m.geIndStruct.gdhData3, i, dY
                    gdSetNum m.geIndStruct.gdhData4, i, dY
                    bMoved = True
                    Exit For
                End If
            Next
        Else
            If Bars(eBARS_High, iLastBar) < dY Then
                Bars(eBARS_High, iLastBar) = dY
            ElseIf Bars(eBARS_Low, iLastBar) > dY Then
                Bars(eBARS_Low, iLastBar) = dY
            End If
            For i = gdGetSize(m.geIndStruct.gdhData4) - 1 To 0 Step -1
                If gdGetNum(m.geIndStruct.gdhData4, i) <> kNullData Then
                    gdSetNum m.geIndStruct.gdhData4, i, dY
                    If gdGetNum(m.geIndStruct.gdhData2, i) < dY Then
                        gdSetNum m.geIndStruct.gdhData2, i, dY
                    ElseIf gdGetNum(m.geIndStruct.gdhData3, i) > dY Then
                        gdSetNum m.geIndStruct.gdhData3, i, dY
                    End If
                    bMoved = True
                    Exit For
                End If
            Next
        End If
    End If
    
    WhatIfMove = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.WhatIfMove"

End Function

' For boolean indicators: pass back false if 0 or Null, else pass # of bars for HighlightBox
Private Function BooleanValue(ByVal dValue#) As Long

    If dValue = 0 Or dValue = kNullData Then
        BooleanValue = 0
    Else
        BooleanValue = Abs(dValue)
    End If

End Function

Public Property Get GroupKey() As String
    GroupKey = m.strGroupKey
End Property

Public Property Let GroupKey(ByVal strKey$)
    m.strGroupKey = strKey
End Property

Public Property Get UpdateGroupFlag() As Boolean
    UpdateGroupFlag = m.bUpdateGroupInProg
End Property

Public Property Let UpdateGroupFlag(ByVal bFlag As Boolean)
    m.bUpdateGroupInProg = bFlag
End Property

Private Sub UpdateGroup(ByVal strGroupKey$, ByVal bDisplay As Boolean)
On Error GoTo ErrSection:
    
    Dim i&, iDone&
    Dim Ind As cIndicator
    Dim Tree As cGdTree
        
    If m.bUpdateGroupInProg Then Exit Sub
    
    If m.Chart Is Nothing Then Exit Sub
    Set Tree = m.Chart.Tree
    If Tree Is Nothing Then Exit Sub
        
    iDone = -1
    For i = 1 To Tree.Count
        If Tree.NodeLevel(i) > 0 Then
            Set Ind = Tree(i)
            If Not Ind Is Nothing Then
                If Ind.GroupKey = strGroupKey Then
                    iDone = i
                    If Ind.Display <> bDisplay Then
                        Ind.UpdateGroupFlag = True
                        Ind.Display = bDisplay
                        Ind.UpdateGroupFlag = False
                    End If
                ElseIf iDone <> -1 Then
                    Exit For
                End If
            End If
        End If
    Next
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.UpdateGroup"
    
End Sub

Public Function IAmNotInMyGroup(strPrevSibGroup$, strNextSibGroup$) As Boolean
On Error GoTo ErrSection:

    Dim Tree As cGdTree
    Dim IndPrev As cIndicator
    Dim IndNext As cIndicator
    
    Dim bOutsideGroup As Boolean, i&
        
    If m.Chart Is Nothing Then Exit Function
    Set Tree = m.Chart.Tree
    If Tree Is Nothing Then Exit Function
    
    i = Tree.RelativeIndex(m.strMyKey, eTREE_PrevSibling)
    If Tree.NodeLevel(i) > 0 Then
        Set IndPrev = Tree.RelativeItem(m.strMyKey, eTREE_PrevSibling)
    End If
    
    i = Tree.RelativeIndex(m.strMyKey, eTREE_NextSibling)
    If Tree.NodeLevel(i) > 0 Then
        Set IndNext = Tree.RelativeItem(m.strMyKey, eTREE_NextSibling)
    End If
    
    bOutsideGroup = True
    
    If Not IndPrev Is Nothing Then
        If IndPrev.GroupKey = m.strGroupKey Then bOutsideGroup = False
        strPrevSibGroup = IndPrev.GroupKey
    End If
    
    If Not IndNext Is Nothing Then
        If IndNext.GroupKey = m.strGroupKey Then bOutsideGroup = False
        strNextSibGroup = IndNext.GroupKey
    End If
        
    IAmNotInMyGroup = bOutsideGroup

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.IAmNotInMyGroup"

End Function

Public Sub CheckGroup()
On Error GoTo ErrSection:

    Dim strPrevGroup$, strNextGroup$
    Dim bReset As Boolean
    Dim bLeader As Boolean
    
    If Len(m.strGroupKey) = 0 Then Exit Sub     'not in a group, don't need to do anything
    
    bLeader = IAmGroupLeader
    
    If IAmNotInMyGroup(strPrevGroup, strNextGroup) Then
        If Len(strPrevGroup) = 0 And Len(strNextGroup) = 0 Then
            'indicator moved outside its group into area of ungrouped indicators
            MoveGroup
        ElseIf Len(strPrevGroup) > 0 And Len(strNextGroup) > 0 Then
            'indicator moved outside its group into the middle of a different group or in between two groups
            If strPrevGroup = strNextGroup Then
                m.strGroupKey = strPrevGroup     'indicator moved into middle of a different group
            Else
                MoveGroup             'indicator moved in between two groups
            End If
        ElseIf Len(m.strGroupKey) > 0 Then
            'an indicator in one group moved to top or bottom of another group
            MoveGroup
        End If
    ElseIf Len(m.strGroupKey) > 0 Then
        If strPrevGroup = strNextGroup Then
            'indicator moved within same group, reset if indicator was group leader
            If bLeader Then ResetGroupLeader
        ElseIf m.strGroupKey = strPrevGroup And Len(strNextGroup) = 0 Then
            'indicator moved to bottom of its own group, reset if indicator was group leader
            If bLeader Then ResetGroupLeader
        Else
            'group leader moved out of position
            ResetGroupLeader
        End If
    End If
    
    Set m.aGroupKeys = Nothing      'reset
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.CheckGroup"

End Sub

Private Sub ResetGroupLeader()
On Error GoTo ErrSection:

    Dim strOldKey$, strNewKey$
    Dim strDonCare1$, strDontCare2$
    Dim i&, iPaneIdx&
    
    Dim Tree As cGdTree
    Dim Pane As cPane
    Dim Ind As cIndicator
    
    If m.Chart Is Nothing Then Exit Sub
    Set Tree = m.Chart.Tree
    If Tree Is Nothing Then Exit Sub
    
    strOldKey = m.strGroupKey
    
    'precautionary check, theoretically should never happen
    If IAmGroupLeader And IAmNotInMyGroup(strDonCare1, strDontCare2) Then
        MoveGroup
        Exit Sub
    End If
        
    i = Tree.RelativeIndex(m.strMyKey, eTREE_Parent)
    If Tree(i) Is Nothing Then
        DebugLog "ResetGroupLeader error: null parent object."
        Exit Sub
    ElseIf Tree.NodeLevel(i) = 0 Then
        Set Pane = Tree(i)
        iPaneIdx = i
    Else
        DebugLog "ResetGroupLeader error: invalid parent - " & Tree(i).Name
        Exit Sub
    End If
        
    i = Tree.RelativeIndex(iPaneIdx, eTREE_FirstChild)
    If Tree(i) Is Nothing Then
        DebugLog "ResetGroupLeader error: null first child - " & Tree(iPaneIdx).Name
        Exit Sub
    ElseIf Tree.NodeLevel(i) > 0 Then
        Set Ind = Tree(i)
    Else
        DebugLog "ResetGroupLeader error: invalid first child - " & Tree(i).Name
        Exit Sub
    End If
    
    While Not Ind Is Nothing
        If Len(strNewKey) = 0 Then
            If Ind.GroupKey = strOldKey Then
                strNewKey = Tree.Key(i)
                Ind.GroupKey = strNewKey
            End If
        ElseIf Ind.GroupKey = strOldKey Then
            Ind.GroupKey = strNewKey
        End If
        i = Tree.RelativeIndex(i, eTREE_NextSibling)
        If Tree.NodeLevel(i) > 0 Then
            Set Ind = Tree(i)
        Else
            Set Ind = Nothing
        End If
    Wend
    
    If Len(strNewKey) > 0 Then m.strGroupKey = strNewKey
    
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.ResetGroupLeader"

End Sub

Public Property Get IAmGroupLeader() As Boolean
    If m.strGroupKey = m.strMyKey Then IAmGroupLeader = True
End Property

Public Property Get MyKey() As String
    MyKey = m.strMyKey
End Property

Public Property Let MyKey(ByVal strKey$)
    m.strMyKey = strKey
End Property

Public Sub SaveGroupInfo()
On Error GoTo ErrSection:

    Dim i&
    Dim Tree As cGdTree
    Dim Ind As cIndicator
        
    If Len(m.strGroupKey) > 0 Then
        If Not m.Chart Is Nothing Then Set Tree = m.Chart.Tree
        If Not Tree Is Nothing Then
            Set m.aGroupKeys = New cGdArray
            For i = 1 To Tree.Count
                If Tree.NodeLevel(i) > 0 Then
                    Set Ind = Tree(i)
                    If Not Ind Is Nothing Then
                        If Ind.GroupKey = m.strGroupKey Then m.aGroupKeys.Add Ind.MyKey
                    End If
                End If
            Next
        End If
    End If
    
    'this is the only indicator in group (can happen when user removes indicator in chart's editor) - clear group key
    If Not m.aGroupKeys Is Nothing Then     '4424
        If m.aGroupKeys.Size = 1 Then m.strGroupKey = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.SaveGroupInfo"

End Sub

Private Sub MoveGroup()
On Error GoTo ErrSection:

    Dim i&, idx&, strLeaderKey$
        
    Dim Ind As cIndicator
    Dim IndLeader As cIndicator
    Dim Tree As cGdTree
    
    Dim aKeys As New cGdArray
    
    If m.aGroupKeys Is Nothing Then Exit Sub
    If m.aGroupKeys.Size = 0 Then Exit Sub
    If Not m.Chart Is Nothing Then Set Tree = m.Chart.Tree
    If Tree Is Nothing Then Exit Sub
        
    If Not IAmGroupLeader Then
        Set IndLeader = Tree(m.aGroupKeys(0))
        If IndLeader Is Nothing Then
            Exit Sub
        Else
            idx = Tree.Move(IndLeader.MyKey, m.strMyKey, eTREE_PrevSibling)
            Set Ind = Tree(idx)
            If Ind Is Nothing Then
                Exit Sub
            Else
                Ind.geIndId = idx
                strLeaderKey = Ind.MyKey
            End If
        End If
    End If
            
    For i = 1 To m.aGroupKeys.Size - 1
        idx = Tree.Move(m.aGroupKeys(i), m.aGroupKeys(i - 1), eTREE_NextSibling)
        Set Ind = Tree(idx)
        If Not Ind Is Nothing Then Ind.geIndId = idx
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.MoveGroup"

End Sub

Public Function RemoveMyGroup() As Boolean
On Error GoTo ErrSection:

    Dim i As Long
    Dim bRC As Boolean
    
    Dim Tree As cGdTree
    Dim Ind As cIndicator
    
    If Len(m.strGroupKey) = 0 Then Exit Function        'double check that this is part of a group
    
    If Not m.Chart Is Nothing Then Set Tree = m.Chart.Tree
    If Tree Is Nothing Then Exit Function
    
    If m.aGroupKeys Is Nothing Then SaveGroupInfo
    If m.aGroupKeys Is Nothing Then Exit Function
    
    For i = 0 To m.aGroupKeys.Size - 1
        bRC = Tree.Remove(m.aGroupKeys(i))
        If Not bRC Then Exit For
    Next
    
    Set m.aGroupKeys = Nothing
    Set Tree = Nothing
    RemoveMyGroup = bRC

ErrExit:
    Exit Function
    
ErrSection:
    Set m.aGroupKeys = Nothing
    Set Tree = Nothing
    RaiseError "cIndicator.RemoveMyGroup"

End Function

Private Sub FakeMarkersShift(ByVal nShift&)
On Error GoTo ErrSection:

    Dim i&, j&, idx&, iSize&, hImage&
    Dim strImage$, strNum$
    Dim aImage As New cGdArray
   
    For idx = 0 To gdGetSize(m.geIndStruct.gdshImage) - 1
        strImage = gdGetStr(m.geIndStruct.gdshImage, idx)
        aImage.SplitFields strImage, ","
        iSize = aImage.Size
        hImage = aImage.ArrayHandle
    
        If nShift > 0 Then
            For i = iSize - 1 To 0 Step -1
                strNum = gdGetStr(hImage, i)
                If Val(strNum) <> -1 Then
                    j = i + nShift
                    If j >= 0 And j < iSize Then
                        gdSetStr hImage, j, strNum
                    End If
                    gdSetStr hImage, i, "-1"
                End If
            Next
        Else
            For i = 0 To iSize - 1
                strNum = gdGetStr(hImage, i)
                If Val(strNum) <> -1 Then
                    j = i + nShift
                    If j >= 0 And j < iSize Then
                        gdSetStr hImage, j, strNum
                    End If
                    gdSetStr hImage, i, "-1"
                End If
            Next
        End If
        
        strImage = aImage.JoinFields(",")
        gdSetStr m.geIndStruct.gdshImage, idx, strImage
    Next

ErrExit:
    Set aImage = Nothing
    Exit Sub

ErrSection:
    Set aImage = Nothing
    RaiseError "cIndicator.FakeMarkersShift"

End Sub

Private Sub SetArtPyramidData(hData&, hHigh&, hLow&, hXBar&, ByVal idxStart&, ByVal idxEnd&)
On Error GoTo ErrSection:

    Dim i&, j&
            
    If m.eLabelMode <> eINDIC_NoValue Then m.eLabelMode = eINDIC_Nothing
    
    m.geIndStruct.FillColor = m.nUpColor
    m.geIndStruct.fillColor2 = m.nDownColor
    m.geIndStruct.trueRangeColor = m.nBoxFillColor          'color for potential pyramid or voided reversal bar
    
    geSetFontInfo 0
    
    If m.eDisplayType = eINDIC_ArtPyramid Then
        m.geIndStruct.indicatorType = 17
    Else
        m.geIndStruct.indicatorType = 18
        m.geIndStruct.FillPattern = m.nBoxFillStyle
    End If
            
    For i = idxStart To idxEnd
        gdSetNum m.geIndStruct.gdhData1, j, gdGetNum(hData, gdGetNum(hXBar, i))     'this is results array from BuiltIn.dll
        gdSetNum m.geIndStruct.gdhData2, j, gdGetNum(hHigh, gdGetNum(hXBar, i))
        gdSetNum m.geIndStruct.gdhData3, j, gdGetNum(hLow, gdGetNum(hXBar, i))
        j = j + 1
    Next
                
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.SetArtPyramidData"

End Sub

Public Property Get FontName() As String
On Error GoTo ErrExit:

    FontName = m.strFontName
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.Font.Name.Get"
End Property

Public Property Let FontName(ByVal strName$)
On Error GoTo ErrExit:

    m.strFontName = strName
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.Font.Name.Let"
End Property

Public Property Get FontSize() As Long
On Error GoTo ErrExit:

    FontSize = m.nFontSize
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.Font.Size.Get"
End Property

Public Property Let FontSize(ByVal nSize&)
On Error GoTo ErrExit:

    m.nFontSize = nSize
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.Font.Size.Let"
End Property

Public Property Get FontBold() As Boolean
On Error GoTo ErrExit:

    If m.nFontBold = 0 Then
        FontBold = False
    Else
        FontBold = True
    End If
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.Font.Bold.Get"
End Property

Public Property Let FontBold(ByVal bBold As Boolean)
On Error GoTo ErrExit:

    If bBold Then
        m.nFontBold = 1
    Else
        m.nFontBold = 0
    End If
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.Font.Bold.Let"
End Property

Public Property Get FontItalic() As Boolean
On Error GoTo ErrExit:

    If m.nFontItalic = 0 Then
        FontItalic = False
    Else
        FontItalic = True
    End If
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.FontItalic.Get"
End Property

Public Property Let FontItalic(ByVal bItalic As Boolean)
On Error GoTo ErrExit:

    If bItalic Then
        m.nFontItalic = 1
    Else
        m.nFontItalic = 0
    End If
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.FontItalic.Let"
End Property

Public Function HasAssociatedAlerts(ByVal bRemove As Boolean, ByVal bRefresh As Boolean) As Boolean
On Error GoTo ErrExit:

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Caution: Use the bRemove flag = true with care.
'   When flag is true indicators will be removed from chart's tree object.
'   If calling routine is holding on to an index into the tree, this index
'   will not be pointing to the same object if items are removed here.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim bHasAlert As Boolean
    Dim i&, strText$, strKeyword$, strCodedNames
    
    Dim Tree As cGdTree
    Dim Ind As cIndicator

    If g.bUnloading Or g.bLoadingChartPage Then Exit Function
    If m.Chart Is Nothing Then Exit Function
    If Len(m.strCodedName) = 0 Then Exit Function
    
    If m.bAlert And Not bRemove Then
        HasAssociatedAlerts = True
        Exit Function
    End If
    
    Set Tree = m.Chart.Tree
    If Tree Is Nothing Then Exit Function
    
    'special case for ART Pyramid & ART Reverse Bars:
    'alerts are tied to the highlight bars, but logically associated with the ART Pyramid or Reverse Bars
    If InStr(m.strCodedName, "PyramidTP") <> 0 Then
        strKeyword = "Pyramid"
    ElseIf InStr(m.strCodedName, "ReverseBarTP") <> 0 Then
        strKeyword = "ARTReverseBars"
    End If
    
    'compare the coded text of any Ind-type alert against this indicator's coded name or special key word
    For i = Tree.Count To 0 Step -1
        If Tree.NodeLevel(i) > 0 Then
            Set Ind = Tree(i)
            If Not Ind Is Nothing Then
                If Ind.IsAlert Then
                    strText = Ind.CodedText
                    If InStr(strText, m.strCodedName) <> 0 Then
                        bHasAlert = True
                        If bRemove Then
                            Ind.UpdateAlert 0, bRefresh
                            Tree.Remove i
                        Else
                            Exit For        'found an alert, remove not requested no need to continue
                        End If
                    ElseIf Len(strKeyword) > 0 Then
                        If InStr(strText, strKeyword) <> 0 Then
                            bHasAlert = True
                            If bRemove Then
                                Ind.UpdateAlert 0, bRefresh
                                Tree.Remove i
                            Else
                                Exit For    'found an alert, remove not requested no need to continue
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    HasAssociatedAlerts = bHasAlert
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cIndicator.HasAssociatedAlert.Get"

End Function

Public Property Get IsHawkeyeLevels() As Boolean
    If UCase(m.strCodedName) = "HAWKEYELEVELS" Then IsHawkeyeLevels = True
End Property

Public Property Get IsHawkeyeAdds() As Boolean
    If UCase(m.strCodedName) = "HAWKEYEADDS" Then IsHawkeyeAdds = True
End Property

Public Property Get HawkeyeLevelsColors() As String
    
    If Len(m.strHawkeyeColors) = 0 Then m.strHawkeyeColors = kHawkeyeLevelsColors
    HawkeyeLevelsColors = m.strHawkeyeColors

End Property

Public Property Let HawkeyeLevelsColors(ByVal strColors$)
    m.strHawkeyeColors = strColors
End Property

Public Property Get IsLinkedParm(ByVal nItem) As Boolean
On Error GoTo ErrSection:

    Dim strName$, strVal$, strType$
    Dim aInputs As cGdArray
    Dim bRet As Boolean

    strName = Parse(m.strLinkedParms, "|", nItem)
    
    'double check that linked parm has not been removed
    If Len(strName) > 0 And Not m.Chart Is Nothing Then
        m.Chart.LinkedInputGet strName, strVal, strType
        If Len(strVal) > 0 Then
            bRet = True
        Else
            Set aInputs = New cGdArray
            aInputs.SplitFields m.strLinkedParms, "|"
            aInputs(nItem) = ""
            m.strLinkedParms = aInputs.JoinFields("|")
            bRet = False
        End If
    ElseIf Len(strName) > 0 Then
        bRet = True
    End If
    
    IsLinkedParm = bRet

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.IsLinkedParm.Get"

End Property

Public Property Get LinkedParmAllowed(ByVal nItem) As Boolean
On Error GoTo ErrSection:

    Dim i&

    If nItem <= ParmCount Then
        i = ParmType(nItem)
        'JM 09-24-2009: for now only doing parm types:
        '   contant number (1), constant text (2), constant boolean (5) or market (6)
        If i = 1 Or i = 2 Or i = 5 Or i = 6 Then LinkedParmAllowed = True
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.LinkedParmAllowed"

End Property

Private Function LinkedValValid(ByVal nItem&, ByVal strValue$) As Boolean
On Error GoTo ErrSection:

    Dim i&
    Dim bValid As Boolean
        
    If Len(strValue) = 0 Then Exit Function
    
    bValid = True       'assume valid
    
    If ParmType(nItem) = 6 Then
        If strValue <> "1" And strValue <> "0" Then
            If UCase(strValue) <> "FALSE" And UCase(strValue) <> "TRUE" Then
                bValid = False
            End If
        End If
    End If
        
    LinkedValValid = bValid

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.LinkedValValid"

End Function

Public Function LinkedParmGet(ByVal nItem&, ByRef strName$, ByRef strVal$) As Boolean
On Error GoTo ErrSection:
    
'JM 10-01-2009 design notes:
'   the value associated with the linked input name is save in string format only at the chart level
'       &Name;Value|&Name2;Value2|&Name3;Value3 ...
'
'   the indicator object saves only the names of the linked input it uses in exact order of its parms count
'       eg. if an indicator has 3 parms & the 2nd parm uses a linked input the string looks like so:
'           |&linkedinputname|
'
'   this design ensures backwards compatibility as well as minimal changes to existing parms code
'       parm values are still saved directly to the parm settings regardless of whether linked inputs are used
'       a user running a version of TN that does not have linked inputs implemented can still load templates that use linked inputs
    
    Dim i&, strLinkedName$, strLinkedVal$, strLinkedType$
    
    strLinkedName = Parse(m.strLinkedParms, "|", nItem)
    
    'initialize return variables
    strName = ""
    strVal = ""
    
    'get value from chart's object
    If Len(strLinkedName) > 0 Then
        If m.Chart Is Nothing Then
            strName = strLinkedName
            strLinkedVal = Parm(nItem)
        Else
            m.Chart.LinkedInputGet strLinkedName, strLinkedVal, strLinkedType
        End If
        If Len(strLinkedVal) > 0 Then
            LinkedParmGet = True
            strName = strLinkedName
            strVal = strLinkedVal
        End If
    End If
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.LinkdParmsGet"

End Function

Public Function LinkedParmSet(ByVal nItem&, ByVal strName$, Optional ByVal strVal$ = "") As Boolean
On Error GoTo ErrSection:

    Dim i&, strParmVal$, strType$
    
    Dim bValid As Boolean
    Dim aNew As cGdArray
    
    If Not LinkedParmAllowed(nItem) Then Exit Function
    
    If Len(strVal) > 0 Then
        strParmVal = strVal
    ElseIf Not m.Chart Is Nothing Then
        m.Chart.LinkedInputGet strName, strParmVal, strType
    End If
    
    If Len(strParmVal) <= 0 Then GoTo ErrExit
    
    bValid = LinkedValValid(nItem, strParmVal)
        
    If bValid Then
        Set aNew = New cGdArray
        
        aNew.SplitFields m.strLinkedParms, "|"
        aNew(nItem - 1) = strName
        m.strLinkedParms = aNew.JoinFields("|")
        
        aNew.SplitFields m.strLinkedParms, "|"
        
        Parm(nItem) = strParmVal
    
    End If
    
    LinkedParmSet = bValid

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.LinkedParmSet"

End Function

Public Sub LinkedParmsDelete(ByVal strName$, ByVal nItem&)
On Error GoTo ErrSection:

    Dim i&
    Dim aParms As New cGdArray
    
    aParms.SplitFields m.strLinkedParms, "|"
    
    For i = aParms.Size - 1 To 0 Step -1
        If aParms(i) = strName Then
            If nItem = -1 Then
                aParms.Remove i     'remove the linked reference completely
            ElseIf nItem = i + 1 Then
                aParms(i) = ""      'remove linked reference only for this parm
                Exit For
            End If
        End If
    Next
    
    m.strLinkedParms = aParms.JoinFields("|")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.LinkdParmsDelete"

End Sub

Public Sub LinkedParmsUpdate()
On Error GoTo ErrSection:

    Dim i&, strName$, strValue$, strRet$

    If Not m.Chart Is Nothing Then
        For i = 1 To ParmCount
            If IsLinkedParm(i) Then
                LinkedParmGet i, strName, strValue
                If Len(strName) > 0 And Len(strValue) > 0 Then
                    Parm(i) = strValue
                End If
            End If
        Next
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.LinkdParmsUpdate"

End Sub

Public Property Get LinkedParmsString() As String
    LinkedParmsString = m.strLinkedParms
End Property

Private Sub TemplateSaveLinkedParms(ByVal fh%)
On Error GoTo ErrSection:

    Dim i&
    Dim aParms As New cGdArray
    
    Dim strName As String
    Dim strValue As String
    Dim strType As String
    Dim strLinkedInputs As String

    If Not m.Chart Is Nothing And Len(m.strLinkedParms) > 0 Then
        'need to validate & update linked input values from chart level prior to saving to template
        'templates might have been swapped back & forth between users of TN with/without linked inputs
        strLinkedInputs = m.Chart.LinkedInputString
        
        If Len(strLinkedInputs) > 0 Then
            aParms.SplitFields m.strLinkedParms, "|"
            
            For i = aParms.Size - 1 To 0 Step -1
                strName = aParms(i)
                If Len(strName) > 0 Then
                    m.Chart.LinkedInputGet strName, strValue, strType
                    If Len(strValue) = 0 Then
                        aParms.Remove i
                    Else
                        If LinkedValValid(i + 1, strValue) Then Parm(i + 1) = strValue
                    End If
                End If
            Next
        End If
    
        If aParms.Size > 0 Then
            m.strLinkedParms = aParms.JoinFields("|")
        Else
            m.strLinkedParms = ""
        End If
    End If
    
    Print #fh, IniString("LinkedParms", m.strLinkedParms)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cIndicator.TemplateSaveLinkedParms"

End Sub

Public Function geAddToDLL(Chart As cChart) As Long
On Error GoTo ErrSection:
    
    Dim rc&
    
    If Not m.geAdded Then
        rc = geAddItem(Chart.geChartObj, 3, m.geIndStruct)
        m.geAdded = True
    End If
    
    geAddToDLL = rc
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cIndicator.geAddToDLL"

End Function

Public Sub geResetIndStruct()
On Error GoTo ErrSection:

    'set indicator common properties
    m.geIndStruct.paneId = -1
    m.geIndStruct.indicatorId = -1
    m.geIndStruct.FillPattern = -1
    m.geIndStruct.FillColor = -1
    m.geIndStruct.fillColor2 = -1
    
    'clear indicator's data arrays
    gdClear m.geIndStruct.gdhData1, 1
    gdClear m.geIndStruct.gdhData2, 1
    gdClear m.geIndStruct.gdhData3, 1
    gdClear m.geIndStruct.gdhData4, 1
    gdClear m.geIndStruct.gdhYScaleVal, 1
    gdClear m.geIndStruct.glhFlags, 1
    
    'clear indicator's image / highlight info arrays
    gdClear m.geIndStruct.gdshImage, 1
    gdClear m.geIndStruct.glhImageLoc, 1
    gdClear m.geIndStruct.glhImageDir, 1
    gdClear m.geIndStruct.glhImageFill, 1
    gdClear m.geIndStruct.glhImageColor, 1
    
    'clear indicator's pen info arrays
    gdClear m.geIndStruct.glhPenColor, 1
    gdClear m.geIndStruct.glhPenStyle, 1
    gdClear m.geIndStruct.glhPenWidth, 1
    
    'clear pnf flag
    m.geIndStruct.pnfX = 0
    
    'set defaults
    geSetPenStyle
    geSetPenColor 0, m.nColor
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.geResetIndStruct"

End Sub

Public Sub LoadCompositePFP(Chart As cChart, oPFP As cPatternProfit, _
    ByVal paneId&, ByVal nID&, ByVal nScreenStartX&, ByVal nScreenEndX&)
On Error GoTo ErrSection:

    Dim i&, j&, idx&, nCount&, nColor&
    
    Dim dOpen#, dHigh#, dLow#, dClose#, dStdAbove#, dStdBelow#
    Dim dBase#, dDateFromPFP#, dDateToPFP#, dSd1#
    Dim dBarDate#, dBarDateStart#, dBarDateEnd#
    
    Dim hDataClose&, hDataStdAbove&, hDataStdBelow&         'array handles to m.data member of cIndicator object
    Dim hIndClose&, hIndStdAbove&, hIndStdBelow&            'array handles to structure in graheng dll
    Dim hXBar&, hBarsDT&
    
    Dim CompositeTree As cGdTree
    Dim IndClose As cIndicator
    Dim IndStdAbove As cIndicator
    Dim IndStdBelow As cIndicator
    
    Dim bShow As Boolean
    
    If oPFP Is Nothing Then Exit Sub
    If Chart Is Nothing Then Exit Sub
    If Chart.Bars Is Nothing Then Exit Sub
    If Chart.aXBar Is Nothing Then Exit Sub
    
    Set CompositeTree = Chart.CompositeIndTree
    If CompositeTree Is Nothing Then Exit Sub
    
    'set array handles to access correct datetime in bars even when empty bars are turned on
    hXBar = Chart.aXBar.ArrayHandle
    hBarsDT = Chart.Bars.ArrayHandle(eBARS_DateTime)
    
    m.eStyle = oPFP.ForecastLineStyle
    geResetIndStruct
    m.geIndStruct.paneId = paneId
    m.geIndStruct.indicatorId = nID
    m.eDataType = eINDIC_BarData
    m.eDisplayType = oPFP.IndDisplayStyle
    m.strName = "PFP_OHLC"
    m.strChartLabel = "PFP"
    
    m.geIndStruct.indicatorType = 6
    geSetPenColor 0, oPFP.ForecastColor
    geSetPenStyle
    
    'initialize close line indicator
    If CompositeTree.Count = 0 Then
        Set IndClose = New cIndicator
    Else
        Set IndClose = CompositeTree(2)
    End If
    If IndClose Is Nothing Then GoTo ErrExit
    
    IndClose.DataType = eINDIC_Array
    IndClose.DisplayType = eINDIC_Line
    IndClose.Name = "PFP_Close"
    IndClose.Color = oPFP.ForecastColor
    IndClose.Style = eINDIC_Dot
    IndClose.geResetIndStruct
    
    IndClose.geIndpaneId = paneId
    IndClose.geIndId = nID + 1
    hDataClose = IndClose.Data.ArrayHandle
    hIndClose& = IndClose.geDataHandle(1)
    
    'initialize standard dev above
    If CompositeTree.Count = 0 Then
        Set IndStdAbove = New cIndicator
    Else
        Set IndStdAbove = CompositeTree(3)
    End If
    If IndStdAbove Is Nothing Then GoTo ErrExit
    
    IndStdAbove.DataType = eINDIC_Array
    IndStdAbove.DisplayType = eINDIC_Line
    IndStdAbove.Name = "PFP_StdAbove"
    IndStdAbove.Color = oPFP.ForecastColor
    IndStdAbove.Style = eINDIC_Dot
    IndStdAbove.geResetIndStruct
    
    IndStdAbove.geIndpaneId = paneId
    IndStdAbove.geIndId = nID + 2
    hDataStdAbove = IndStdAbove.Data.ArrayHandle
    hIndStdAbove = IndStdAbove.geDataHandle(1)
    
    'initialize standard dev below
    If CompositeTree.Count = 0 Then
        Set IndStdBelow = New cIndicator
    Else
        Set IndStdBelow = CompositeTree(4)
    End If
    If IndStdBelow Is Nothing Then GoTo ErrExit
    
    IndStdBelow.DataType = eINDIC_Array
    IndStdBelow.DisplayType = eINDIC_Line
    IndStdBelow.Name = "PFP_StdBelow"
    IndStdBelow.Color = oPFP.ForecastColor
    IndStdBelow.Style = eINDIC_Dot
    IndStdBelow.geResetIndStruct
    
    IndStdBelow.geIndpaneId = paneId
    IndStdBelow.geIndId = nID + 3
    hDataStdBelow = IndStdBelow.Data.ArrayHandle
    hIndStdBelow = IndStdBelow.geDataHandle(1)
    
    'set variables for charting composite values
    i = Chart.Bars.FindDateTime(oPFP.PatternDateTo)
    dBase = Chart.Bars(eBARS_Close, i)
    dDateFromPFP = oPFP.PatternDateFrom
    dDateToPFP = oPFP.PatternDateTo
    
    'set counters
    j = 1
    idx = 0
    nCount = oPFP.CompositeCount
    For i = nScreenStartX To nScreenEndX
        dOpen = kNullData
        dHigh = kNullData
        dLow = kNullData
        dClose = kNullData
        dStdAbove = kNullData
        dStdBelow = kNullData
        
        gdSetNum hDataClose, idx, kNullData
        gdSetNum hIndClose, idx, kNullData
        
        gdSetNum hDataStdAbove, idx, kNullData
        gdSetNum hIndStdAbove, idx, kNullData
        
        gdSetNum hDataStdBelow, idx, kNullData
        gdSetNum hIndStdBelow, idx, kNullData
        
        dBarDate = gdGetNum(hBarsDT, gdGetNum(hXBar, i))
        
        If dBarDate >= dDateFromPFP Then
            If j > nCount Then
                Exit For
            ElseIf dBarDate >= dDateToPFP Then
                oPFP.CompositeData j, dBase, dBarDate, dOpen, dHigh, dLow, dClose, dStdAbove, dStdBelow
                
                dSd1 = dStdAbove - dClose
                If dOpen = 0 And dHigh = 0 And dLow = 0 And dClose = 0 Then
                    dOpen = kNullData
                    dHigh = kNullData
                    dLow = kNullData
                    dClose = kNullData
                ElseIf m.eDisplayType <> eINDIC_Line And dBarDate = dDateToPFP Then
                    dOpen = kNullData       'don't want to overlap OHLC bar with last bar in pattern, but do want standard dev values
                    dHigh = kNullData
                    dLow = kNullData
                    dClose = kNullData
                Else
                    gdSetNum hDataClose, idx, dBase     'close line
                    gdSetNum hIndClose, idx, dBase
                End If
                
                'standard dev lines
                gdSetNum hDataStdAbove, idx, dStdAbove
                gdSetNum hIndStdAbove, idx, dStdAbove
                gdSetNum hDataStdBelow, idx, dStdBelow
                gdSetNum hIndStdBelow, idx, dStdBelow
                
                If Not bShow Then
                    If dOpen <> kNullData Or dHigh <> kNullData Or dLow <> kNullData Or dClose <> kNullData Then bShow = True
                End If
            End If
            j = j + 1
        End If
        
        If m.eDisplayType = eINDIC_Line Then
            gdSetNum m.geIndStruct.gdhData1, idx, dClose
            gdSetNum m.geIndStruct.gdhData2, idx, dSd1
        Else
            gdSetNum m.geIndStruct.gdhData1, idx, dOpen
            gdSetNum m.geIndStruct.gdhData2, idx, dHigh
            gdSetNum m.geIndStruct.gdhData3, idx, dLow
            gdSetNum m.geIndStruct.gdhData4, idx, dClose
        End If
        
        idx = idx + 1
    Next
    
    If oPFP.Heatmap = vbChecked Then
        IndStdBelow.trueRangeColor = 100
        m.geIndStruct.trueRangeColor = oPFP.HeatmapStyle
        IndStdAbove.trueRangeColor = 300
    Else
        IndStdBelow.trueRangeColor = 0
        m.geIndStruct.trueRangeColor = 0
        IndStdAbove.trueRangeColor = 0
    End If
    
    dBarDateStart = gdGetNum(hBarsDT, gdGetNum(hXBar, nScreenStartX))
    If dBarDateStart <= 0 Then dBarDateStart = gdGetNum(hBarsDT, gdGetNum(hXBar, nScreenStartX - 1))    'could be blank bar
    If dBarDateStart <= 0 Then dBarDateStart = gdGetNum(hBarsDT, gdGetNum(hXBar, nScreenStartX - 2))
    
    dBarDateEnd = gdGetNum(hBarsDT, gdGetNum(hXBar, nScreenEndX))
    If dBarDateEnd <= 0 Then dBarDateEnd = gdGetNum(hBarsDT, gdGetNum(hXBar, nScreenEndX + 1))
    If dBarDateEnd <= 0 Then dBarDateEnd = gdGetNum(hBarsDT, gdGetNum(hXBar, nScreenEndX + 2))
    
    If dBarDateStart <= 0 Or dBarDateEnd <= 0 Then
        DebugLog "LoadCompositePFP: unable to get dates " & Chart.Bars.Prop(eBARS_Symbol) & " " & Chart.Bars.Prop(eBARS_PeriodicityStr)
    Else
        If bShow And Not CompositeTree Is Nothing Then
            oPFP.LoadMatchAnnot Chart, Chart.Form, dBarDateStart, dBarDateEnd
            i = geAddToDLL(Chart)
            If i = 0 Then CompositeTree.Add Me, m.strName
            
            i = IndClose.geAddToDLL(Chart)
            If i = 0 Then CompositeTree.Add IndClose, IndClose.Name
            
            i = IndStdAbove.geAddToDLL(Chart)
            If i = 0 Then CompositeTree.Add IndStdAbove, IndStdAbove.Name
        
            i = IndStdBelow.geAddToDLL(Chart)
            If i = 0 Then CompositeTree.Add IndStdBelow, IndStdBelow.Name
        
        ElseIf oPFP.MatchCount > 0 Then
            oPFP.LoadMatchAnnot Chart, Chart.Form, dBarDateStart, dBarDateEnd
        End If
    End If
    
ErrExit:
    Set CompositeTree = Nothing
    Set IndClose = Nothing
    Set IndStdAbove = Nothing
    Set IndStdBelow = Nothing
    Exit Sub

ErrSection:
    RaiseError "cIndicator.LoadCompositePFP"

End Sub

Public Property Get CanBeNonLinear() As Boolean
On Error GoTo ErrSection:

    Dim iSize As Long
    Dim bOkay As Boolean
    Dim hHandle As Long
    
    bOkay = True
    If Not m.bDisplay Then GoTo ErrExit
    
    If m.eDataType = eINDIC_Array Then
        hHandle = m.aData.ArrayHandle
    ElseIf m.eDataType = eINDIC_BarData Then
        If Not m.Bars Is Nothing Then hHandle = m.Bars.ArrayHandle(eBARS_Low)
    ElseIf m.eDataType = eINDIC_None Then
        bOkay = False
    End If
    
    If hHandle <> 0 Then
        iSize = gdGetSize(hHandle)
        If iSize > 0 Then
            If gdMinValue(hHandle, 0, iSize) <= 0 Then bOkay = False
        End If
    End If
    

ErrExit:
    CanBeNonLinear = bOkay
    Exit Property

ErrSection:
    CanBeNonLinear = False
    RaiseError "cIndicator.CanBeNonLinear.Get"

End Property

Public Sub ClusterIndDefaults(Chart As cChart, ByVal eType As eIndicatorDisplayType, _
    Optional ByVal bRatiosOnly As Boolean = False, Optional ByVal dEndDate# = 0#, Optional ByVal nBars& = -1)
On Error GoTo ErrExit

    Dim strRatios$, strRatioWeights$, strSwingWeights$

    If Chart Is Nothing Then Exit Sub
    If Chart.Bars Is Nothing Then Exit Sub
    If eType <> eINDIC_ClusterPrice And eType <> eINDIC_ClusterTime Then Exit Sub
        
    m.eDisplayType = eType
    m.eDataType = eINDIC_Array
    m.bDisplay = True
    m.bOverlayed = False
    
    If bRatiosOnly Then
        strRatios = "0.382|0.5|0.618|1|1.382|1.5|1.618|2"
        strRatioWeights = "1|1|1|1|1|1|1|1"
        RatioStringToTable strRatios, strRatioWeights
        GoTo ErrExit
    End If
    
    If m.eDisplayType = eINDIC_ClusterPrice Then
        CodedName = "SwingPointLevels"
        m.strName = "Price Clusters"
        
        'first value in string must be 0 then it is weight for short term|medium term|long term swings
        strSwingWeights = GetIniFileProperty("ClusterSwingWeights", "0|100|200|300", kClusterIniSection, g.strIniFile)
        Set m.aClusterSwingWeights = New cGdArray
        m.aClusterSwingWeights.SplitFields strSwingWeights, "|"
        
        m.nClusterProximity = GetIniFileProperty("PriceProximity", 20, kClusterIniSection, g.strIniFile)
        m.nClusterSwingColorL = GetIniFileProperty("ClusterSwingColorL", vbBlue, kClusterIniSection, g.strIniFile)
        m.nClusterSwingColorM = GetIniFileProperty("ClusterSwingColorM", RGB(0, 128, 0), kClusterIniSection, g.strIniFile)
        m.nClusterSwingColorS = GetIniFileProperty("ClusterSwingColorS", vbRed, kClusterIniSection, g.strIniFile)
        
        m.nClusterSwingLabel = ValOfText(GetIniFileProperty("ClusterSwingLabel", 1, kClusterIniSection, g.strIniFile))
    
        m.strFontName = GetIniFileProperty("ClusterSwingFontName", "Arial", kClusterIniSection, g.strIniFile)
        m.nFontBold = ValOfText(GetIniFileProperty("ClusterSwingFontBold", 0, kClusterIniSection, g.strIniFile))
        m.nFontItalic = ValOfText(GetIniFileProperty("ClusterSwingFontItalic", 0, kClusterIniSection, g.strIniFile))
        m.nFontSize = ValOfText(GetIniFileProperty("ClusterSwingFontSize", 8, kClusterIniSection, g.strIniFile))
    
        m.eStyle = ValOfText(GetIniFileProperty("ClusterSwingStyle", 0, kClusterIniSection, g.strIniFile))
        m.geIndStruct.labelColor = ValOfText(GetIniFileProperty("PriceColor", vbBlue, kClusterIniSection, g.strIniFile))
        m.geIndStruct.trueRangeColor = ValOfText(GetIniFileProperty("TrueRangeColor", 0, kClusterIniSection, g.strIniFile)) '1=heatmap colors
        m.geIndStruct.TrueRangeFlag = ValOfText(GetIniFileProperty("TrueRangeFlag", 1, kClusterIniSection, g.strIniFile))   '1=cluster & swing, 2=cluster only, 3=swing only
    
        m.bResistanceOrL2H = GetIniFileProperty("UseResistance", True, kClusterIniSection, g.strIniFile)     'used to be rallies
        m.bSupportOrH2L = GetIniFileProperty("UseSupport", True, kClusterIniSection, g.strIniFile)        'used to be declines
        m.bClusterABCs = GetIniFileProperty("UseABCs", False, kClusterIniSection, g.strIniFile)
        
        m.bRetraceOrL2L = True          'not exposed to user
        m.bExpanOrH2H = True            'not exposed to user
    
        strRatios = GetIniFileProperty("PriceFibRatios", "0.382|0.5|0.618|1|1.382|1.5|1.618|2", kClusterIniSection, g.strIniFile)
        strRatioWeights = GetIniFileProperty("PriceFibWeights", "1|1|1|1|1|1|1|1", kClusterIniSection, g.strIniFile)
        RatioStringToTable strRatios, strRatioWeights
    Else
        m.strName = "Time Clusters"
        
        m.nClusterProximity = ValOfText(GetIniFileProperty("TimeProximity", 1, kClusterIniSection, g.strIniFile))
        m.nColor = ValOfText(GetIniFileProperty("TimeColor", vbBlue, kClusterIniSection, g.strIniFile))
        m.geIndStruct.trueRangeColor = ValOfText(GetIniFileProperty("TrueRangeColor", 0, kClusterIniSection, g.strIniFile))     '1=heatmap colors
        m.geIndStruct.TrueRangeFlag = ValOfText(GetIniFileProperty("TrueRangeFlag", 1, kClusterIniSection, g.strIniFile))       'will always = abs(display boolean) so theoretically always be 1 in grapheng.dll
    
        m.bExpanOrH2H = GetIniFileProperty("HighToHigh", True, kClusterIniSection, g.strIniFile)
        m.bSupportOrH2L = GetIniFileProperty("HighToLow", True, kClusterIniSection, g.strIniFile)
        m.bResistanceOrL2H = GetIniFileProperty("LowToHigh", True, kClusterIniSection, g.strIniFile)
        m.bRetraceOrL2L = GetIniFileProperty("LowToLow", True, kClusterIniSection, g.strIniFile)
        
        m.bClusterABCs = False          'not exposed to user
        
        strRatios = GetIniFileProperty("TimeFibRatios", "0.382|0.5|0.618|1|1.382|1.5|1.618|2", kClusterIniSection, g.strIniFile)
        strRatioWeights = GetIniFileProperty("TimeFibWeights", "1|1|1|1|1|1|1|1", kClusterIniSection, g.strIniFile)
        RatioStringToTable strRatios, strRatioWeights
    End If
    
    m.nClusterMaxBars = ValOfText(GetIniFileProperty("ClusterMaxBars", 100, kClusterIniSection, g.strIniFile), False)
    If dEndDate > 0 Then
        m.dClusterEndDate = dEndDate
    Else
        m.dClusterEndDate = Chart.LastGoodDataBar(False, False)
    End If
    
    If nBars > 0 Then
        m.nClusterBarsBack = nBars
    Else
        m.nClusterBarsBack = ValOfText(GetIniFileProperty("ClusterBarsBack", 210, kClusterIniSection, g.strIniFile))
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.ClusterIndDefaults"
    
End Sub

Private Function ClusterStartBar(Chart As cChart, ByVal nFirstBar&) As Long
On Error GoTo ErrSection:

    Dim i&, j&, strText$
    
    If m.eDisplayType = eINDIC_ClusterPrice Then
        strText = "Price"
    Else
        strText = "Time"
    End If

    j = Chart.Bars.FindDateTime(m.dClusterEndDate)
    If j = 0 Or j >= Chart.Bars.Size Then
        j = Chart.LastGoodDataBar(False, False)
        m.dClusterEndDate = Chart.Bars(eBARS_DateTime, j)
    End If
    
    i = j - m.nClusterBarsBack
    If i < 0 Then
        DebugLog ("Adjusting " & strText & " Cluster data for: " & Chart.Symbol & ", " & Chart.Bars.Prop(eBARS_PeriodicityStr))
        DebugLog ("Adjusting " & strText & " Cluster End Date - From: " & DateFormat(m.dClusterEndDate, MM_DD_YYYY, HH_MM_SS))
        DebugLog ("Adjusting " & strText & " Cluster Bars Back - From: " & Str(m.nClusterBarsBack))
        
            j = Chart.LastGoodDataBar(False, False)
            i = nFirstBar
            m.dClusterEndDate = Chart.Bars(eBARS_DateTime, j)
            m.nClusterBarsBack = j - i
        
        DebugLog ("Adjusting " & strText & " Cluster End Date - To: " & DateFormat(m.dClusterEndDate, MM_DD_YYYY, HH_MM_SS))
        DebugLog ("Adjusting " & strText & " Cluster Bars Back - To: " & Str(m.nClusterBarsBack))
    End If
    
    ClusterStartBar = i

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cIndicator.ClusterStartBar"
    
End Function

Private Sub ClusterIndPriceCalc(Chart As cChart, ByVal nFirstBar&)
On Error GoTo ErrSection:

    Dim i&, j&
    
    Dim Annot As cAnnotation
    Dim aRatio As cGdArray
    Dim aRatioWeights As cGdArray
    
    If Chart Is Nothing Then Exit Sub
    If Chart.Bars Is Nothing Then Exit Sub
    If Chart.Bars.Size <= 0 Then Exit Sub
    
    If m.aData Is Nothing Then Exit Sub
    
    If m.aClusterSwingWeights Is Nothing Then
        ClusterIndDefaults Chart, eINDIC_ClusterPrice
        If m.aClusterSwingWeights Is Nothing Then Exit Sub
    End If
    If m.aClusterSwingWeights.Size <= 0 Then Exit Sub
    
    If m.tbClusterRatios Is Nothing Then
        ClusterIndDefaults Chart, eINDIC_ClusterPrice, True
        If m.tbClusterRatios Is Nothing Then Exit Sub
    End If
    If m.tbClusterRatios.NumRecords <= 0 Then Exit Sub
    
    Set aRatio = m.tbClusterRatios.FieldArray(0)
    Set aRatioWeights = m.tbClusterRatios.FieldArray(1)
    If aRatio Is Nothing Or aRatioWeights Is Nothing Then Exit Sub
    
    If Not m.bResistanceOrL2H And Not m.bSupportOrH2L Then
        m.bResistanceOrL2H = True   'precautionary, theoretically should never get here
    End If
    
    Set m.tbClusters = Nothing
    
    i = ClusterStartBar(Chart, nFirstBar)
    j = i + m.nClusterBarsBack
    
    If j = 0 Or j >= Chart.Bars.Size Or i < 0 Then
        Set m.tbClusters = Nothing
        DebugLog "Price cluster start/end data invalid. Skipping calculation."
    Else
        Set m.tbClusters = CalcPriceClusters(Chart.Bars, i, j, _
                m.aData, m.aClusterSwingWeights, aRatio, aRatioWeights, _
                m.nClusterProximity, m.bSupportOrH2L, m.bResistanceOrL2H, _
                m.bRetraceOrL2L, m.bExpanOrH2H, m.bClusterABCs)

        'start/end date may have changed when changing bar periods
        Set Annot = m.Chart.Annots(kClusterZoneRect)
        If Not Annot Is Nothing Then
            Annot.dDate(1) = m.Chart.Bars(eBARS_DateTime, i)
            Annot.dDate(2) = m.Chart.Bars(eBARS_DateTime, j)
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.ClusterIndPriceCalc"
    
End Sub

Private Sub ClusterIndTimeCalc(Chart As cChart, ByVal nFirstBar&)
On Error GoTo ErrSection:

    Dim i&, j&
    Dim aRatio As cGdArray
    Dim aRatioWeights As cGdArray

    Dim Ind As cIndicator
    Dim aData As cGdArray
    Dim aSwingWeights As cGdArray

    If Chart Is Nothing Then Exit Sub
    If Chart.Bars Is Nothing Then Exit Sub
    If Chart.Bars.Size <= 0 Then Exit Sub
        
    Set Ind = m.Chart.Tree(kClusterPriceKey)
    If Ind Is Nothing Then Exit Sub
    
'JM 11-04-2010: temporary backwards compatibity for when time/price had separate swing points level indicators
    Set m.aClusterSwingWeights = Nothing
    
    Set aData = Ind.Data
    If aData Is Nothing Then Exit Sub
    If aData.Size <= 0 Then Exit Sub
    
    Set aSwingWeights = Ind.SwingWeightsArray
    If aSwingWeights Is Nothing Then
        Ind.ClusterIndDefaults Chart, eINDIC_ClusterPrice
        If aSwingWeights Is Nothing Then Exit Sub
    End If
    If aSwingWeights.Size <= 0 Then Exit Sub
    
    If m.tbClusterRatios Is Nothing Then
        ClusterIndDefaults Chart, eINDIC_ClusterPrice, True
        If m.tbClusterRatios Is Nothing Then Exit Sub
    End If
    If m.tbClusterRatios.NumRecords <= 0 Then Exit Sub
    
    Set aRatio = m.tbClusterRatios.FieldArray(0)
    Set aRatioWeights = m.tbClusterRatios.FieldArray(1)
    If aRatio Is Nothing Or aRatioWeights Is Nothing Then Exit Sub
    
    Set m.aData = Nothing
    
    i = ClusterStartBar(Chart, nFirstBar)
    j = i + m.nClusterBarsBack
    
    If j = 0 Or j >= Chart.Bars.Size Or i < 0 Then
        Set m.tbClusters = Nothing
        DebugLog "Time cluster start/end data invalid. Skipping calculation."
    Else
        'JM 02-07-2011:
        '   Per Tim ClusterMaxBars not relevant right now because not doing time clusters
        '   over entire history range (just set to something big & not configurable by user)
        Set m.aData = CalcTimeClusters(Chart.Bars, i, j, 999999, aData, _
                            aSwingWeights, aRatio, aRatioWeights, m.nClusterProximity, _
                            m.bResistanceOrL2H, m.bSupportOrH2L, m.bRetraceOrL2L, _
                            m.bExpanOrH2H, m.bClusterABCs)
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.ClusterIndTimeCalc"
    
End Sub

Public Property Get SwingWeightsArray() As cGdArray
    Set SwingWeightsArray = m.aClusterSwingWeights
End Property

Public Property Get ClusterBarsMax() As Long
    ClusterBarsMax = m.nClusterMaxBars
End Property

Public Property Let ClusterBarsMax(ByVal nMax&)
    m.nClusterMaxBars = Abs(nMax)
End Property

Public Property Get ClusterBarsBack() As Long
On Error GoTo ErrSection:

    Dim Ind As cIndicator
    
    If m.eDisplayType = eINDIC_ClusterPrice Then
        ClusterBarsBack = m.nClusterBarsBack
    ElseIf Not m.Chart Is Nothing Then
        Set Ind = m.Chart.Tree(kClusterPriceKey)
        If Not Ind Is Nothing Then ClusterBarsBack = Ind.ClusterBarsBack
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ClusterBarsBack.Get"

End Property

Public Property Let ClusterBarsBack(ByVal nBars&)
On Error GoTo ErrSection:
    
    Dim Ind As cIndicator
    
    If nBars > 0 Then
        m.nClusterBarsBack = Abs(nBars)
    
        If m.eDisplayType = eINDIC_ClusterPrice And Not m.Chart Is Nothing Then
            Set Ind = m.Chart.Tree(kClusterTimeKeyInd)
            If Not Ind Is Nothing Then Ind.ClusterBarsBack = m.nClusterBarsBack
        End If
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ClusterBarsBack.Let"

End Property

Public Property Get ClusterEndDate() As Double
On Error GoTo ErrSection:
    
    Dim Ind As cIndicator
    
    If m.eDisplayType = eINDIC_ClusterPrice Then
        ClusterEndDate = m.dClusterEndDate
    ElseIf Not m.Chart Is Nothing Then
        Set Ind = m.Chart.Tree(kClusterPriceKey)
        If Not Ind Is Nothing Then ClusterEndDate = Ind.ClusterEndDate
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ClusterEndDate.Get"

End Property

Public Property Let ClusterEndDate(ByVal dDate#)
On Error GoTo ErrSection:
    
    Dim Ind As cIndicator
    
    m.dClusterEndDate = dDate
    
    If m.eDisplayType = eINDIC_ClusterPrice And Not m.Chart Is Nothing Then
        Set Ind = m.Chart.Tree(kClusterTimeKeyInd)
        If Not Ind Is Nothing Then Ind.ClusterEndDate = m.dClusterEndDate
    End If
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ClusterEndDate.Let"

End Property

Public Property Get ClusterProximity() As Long
    ClusterProximity = m.nClusterProximity
End Property

Public Property Let ClusterProximity(ByVal nProximity&)
    If nProximity > 0 And nProximity < 100 Then m.nClusterProximity = nProximity
End Property

Public Property Get ClusterResistance() As Boolean
    ClusterResistance = m.bResistanceOrL2H
End Property

Public Property Let ClusterResistance(ByVal bResistance As Boolean)
    m.bResistanceOrL2H = bResistance
End Property

Public Property Get ClusterSupport() As Boolean
    ClusterSupport = m.bSupportOrH2L
End Property

Public Property Let ClusterSupport(ByVal bSupport As Boolean)
    m.bSupportOrH2L = bSupport
End Property

Public Property Get ClusterABCs() As Boolean
    ClusterABCs = m.bClusterABCs
End Property

Public Property Let ClusterABCs(ByVal bABCs As Boolean)
    m.bClusterABCs = bABCs
End Property

Public Property Get ClusterRetracements() As Boolean
    ClusterRetracements = m.bRetraceOrL2L
End Property

Public Property Let ClusterRetracements(ByVal bRetracements As Boolean)
    m.bRetraceOrL2L = bRetracements
End Property

Public Property Get ClusterExpansions() As Boolean
    ClusterExpansions = m.bExpanOrH2H
End Property

Public Property Let ClusterExpansions(ByVal bExpansions As Boolean)
    m.bExpanOrH2H = bExpansions
End Property

Public Property Get ClusterH2H() As Boolean
    ClusterH2H = m.bExpanOrH2H
End Property

Public Property Let ClusterH2H(ByVal bH2H As Boolean)
    m.bExpanOrH2H = bH2H
End Property

Public Property Get ClusterH2L() As Boolean
    ClusterH2L = m.bSupportOrH2L
End Property

Public Property Let ClusterH2L(ByVal bH2L As Boolean)
    m.bSupportOrH2L = bH2L
End Property

Public Property Get ClusterL2H() As Boolean
    ClusterL2H = m.bResistanceOrL2H
End Property

Public Property Let ClusterL2H(ByVal bL2H As Boolean)
    m.bResistanceOrL2H = bL2H
End Property

Public Property Get ClusterL2L() As Boolean
    ClusterL2L = m.bRetraceOrL2L
End Property

Public Property Let ClusterL2L(ByVal bL2L As Boolean)
    m.bRetraceOrL2L = bL2L
End Property

Public Property Get ClusterSwingShow() As Long
On Error GoTo ErrSection:

    If m.bDisplay Then
        If m.geIndStruct.TrueRangeFlag = 2 Then
            ClusterSwingShow = 0
        Else
            ClusterSwingShow = 1
        End If
    Else
        ClusterSwingShow = 0
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ClusterSwingShow.Get"

End Property

Public Property Let ClusterSwingShow(ByVal nShow&)
On Error GoTo ErrSection:
    
    If m.bDisplay Then
        Select Case m.geIndStruct.TrueRangeFlag
            Case 0
                If nShow = 0 Then m.bDisplay = False        'theoretically should not get here
            Case 1
                If nShow = 0 Then m.geIndStruct.TrueRangeFlag = 2   'show cluster only
            Case 2
                If nShow = 1 Then m.geIndStruct.TrueRangeFlag = 1   'show both
            Case 3
                If nShow = 0 Then
                    m.bDisplay = False
                    m.geIndStruct.TrueRangeFlag = 0
                End If
        End Select
    ElseIf nShow = 1 Then
        m.bDisplay = True
        m.geIndStruct.TrueRangeFlag = 3     'show swing only
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ClusterSwingShow.Let"

End Property

Public Property Get ClusterPriceShow() As Long
On Error GoTo ErrSection:

    If m.bDisplay Then
        If m.geIndStruct.TrueRangeFlag = 3 Then
            ClusterPriceShow = 0
        Else
            ClusterPriceShow = 1
        End If
    Else
        ClusterPriceShow = 0
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ClusterPriceShow.Get"

End Property

Public Property Let ClusterPriceShow(ByVal nShow&)
On Error GoTo ErrSection:

    If m.bDisplay Then
        Select Case m.geIndStruct.TrueRangeFlag
            Case 0:
                If nShow = 0 Then m.bDisplay = False    'theoretically should not get here
            Case 1:
                If nShow = 0 Then m.geIndStruct.TrueRangeFlag = 3   'show swing only
            Case 2:
                If nShow = 0 Then
                    m.bDisplay = False
                    m.geIndStruct.TrueRangeFlag = 0
                End If
            Case 3:
                If nShow = 1 Then m.geIndStruct.TrueRangeFlag = 1   'show both
        End Select
    ElseIf nShow = 1 Then
        m.bDisplay = True
        m.geIndStruct.TrueRangeFlag = 2     'show cluster only
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ClusterPriceShow.Let"

End Property

Public Property Get ClusterSwingColor(ByVal strLevel$) As Long

    Select Case strLevel
        Case "L":
            ClusterSwingColor = m.nClusterSwingColorL       'for L label
        Case "M":
            ClusterSwingColor = m.nClusterSwingColorM       'for M label
        Case "S":
            ClusterSwingColor = m.nClusterSwingColorS       'for S label
        Case "P"
            'JM 07-29-2015: hack for price cluster color so won't need new functions etc.
            ClusterSwingColor = m.geIndStruct.labelColor
    End Select
    
End Property

Public Property Let ClusterSwingColor(ByVal strLevel$, ByVal nColor&)
    
    Select Case strLevel
        Case "L":
            m.nClusterSwingColorL = nColor
        Case "M":
            m.nClusterSwingColorM = nColor
        Case "S":
            m.nClusterSwingColorS = nColor
        Case "P"
            m.geIndStruct.labelColor = nColor
    End Select
    
End Property

Public Property Get ClusterSwingLabel() As Long
    ClusterSwingLabel = m.nClusterSwingLabel
End Property

Public Property Let ClusterSwingLabel(ByVal nShow&)
    m.nClusterSwingLabel = nShow
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Email from Tim (10-25-2010): Inputs to Swing Point Levels Function
' Strength:
'   Negative strength values are not valid and simply act the same as 0 (zero),
'   so they should not be allowed in the interface.
' Method & Altmode:
'   To be less confusing, for now we should not even bother allowing the user to change either of these args
'   -- for clustering purposes, we will just always be setting "Method = 0" (which is the method that counts
'   # swing bars on each side), and set "AltMode = 1" (which is what forces swing points to always alternate
'   from High to Low to High to Low, etc -- i.e. never 2 highs or 2 lows consecutively).
'
' JM (10-25-2010): How swing strength & fib weight values are saved to CHT file
'   Saved as positive if user checked the Use[] checkbox in UI.
'   Saved as negative to preserve last-used value.
'
' JM (10-25-2010): Parm(x) for Swing Point Levels Function
'   Parm1 = Market1
'   Parm2 = Strength for Long Term
'   Parm3 = Strength for Intermediate
'   Parm4 = Strenght for Short Term
'   Parm5 = Method
'   Parm6 = AltMode
'
' JM (10-25-2010): m.aClusterSwingWeights[x]
'   [0] = reserved (must be zero)
'   [1] = weight for Short Term swing
'   [2] = weight for Intermediate swing
'   [3] = weight for Long Term swing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ClusterSwingInfoGet(ByVal nOrder&) As cGdTable
On Error GoTo ErrSection:

    Dim i&, j&
    Dim Ind As cIndicator
    Dim tb As cGdTable
    
    If m.eDisplayType = eINDIC_ClusterTime Then
        'time cluster shares swing point levels indicator with Price Cluster
        If Not m.Chart Is Nothing Then
            If Not m.Chart.Tree Is Nothing Then
                Set Ind = m.Chart.Tree(kClusterPriceKey)
                If Not Ind Is Nothing Then Set tb = Ind.ClusterSwingInfoGet(nOrder)
            End If
        End If
    Else
        If m.aClusterSwingWeights Is Nothing Then Exit Function
        If m.aClusterSwingWeights.Size <= 0 Then Exit Function
        
        Set tb = New cGdTable
        
        tb.CreateField eGDARRAY_Strings, 0, "InputName"
        tb.CreateField eGDARRAY_Longs, 1, "Strength", kNullData
        tb.CreateField eGDARRAY_Doubles, 2, "Weights", kNullData
        
'See note above:
'   the parms order (L,I,S) is opposite of swing weights array (S,I,L)
'   parms order is dictated by Swing Points Levels function
'   swing weights array order is dictated by CalcPriceClusters & CalcTimeClusters functions

        If nOrder = 1 Then
            'nOrder=1: short,intermediate,long
            For i = 4 To 2 Step -1
                tb.AddRecord ""
                j = tb.NumRecords - 1
                tb(0, j) = ParmName(i)
                tb(1, j) = Parm(i)
            Next
            tb(2, 0) = m.aClusterSwingWeights(1)
            tb(2, 1) = m.aClusterSwingWeights(2)
            tb(2, 2) = m.aClusterSwingWeights(3)
        Else
            'nOrder=0: long,intermediate,short
            For i = 2 To 4
                tb.AddRecord ""
                j = tb.NumRecords - 1
                tb(0, j) = ParmName(i)
                tb(1, j) = Parm(i)
            Next
            tb(2, 0) = m.aClusterSwingWeights(3)
            tb(2, 1) = m.aClusterSwingWeights(2)
            tb(2, 2) = m.aClusterSwingWeights(1)
        End If
    End If
        
    Set ClusterSwingInfoGet = tb

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cIndicator.ClusterSwingInfoGet"
    
End Function

Public Sub ClusterSwignInfoSet(tb As cGdTable)
On Error GoTo ErrSection:

    Dim i&, j&, k&
    Dim nCount&, nStrength&, dWeight#
    Dim strName$, strLong$, strIntermediate$, strShort$
    
    Dim Ind As cIndicator
    
    If tb Is Nothing Then Exit Sub
    nCount = tb.NumRecords
    If nCount <> 3 Then Exit Sub        'something wrong don't make changes
    
    If m.eDisplayType = eINDIC_ClusterTime Then
        If Not m.Chart Is Nothing Then
            If Not m.Chart.Tree Is Nothing Then
                Set Ind = m.Chart.Tree(kClusterPriceKey)
                If Not Ind Is Nothing Then Ind.ClusterSwignInfoSet tb
            End If
        End If
    Else
        m.aClusterSwingWeights.Size = 4     'precautionary
        m.aClusterSwingWeights(0) = 0
        
'Parms order: L,I,S
'Swing weights arrray order: S,I,L
        
        strLong = ParmName(2)
        strIntermediate = ParmName(3)
        strShort = ParmName(4)
        
        For i = 0 To nCount
            strName = tb(0, i)
            nStrength = tb(1, i)
            dWeight = tb(2, i)
            
            If Abs(nStrength) > 0 And Abs(dWeight) > 0 Then
                If strName = strLong Then
                    Parm(2) = nStrength
                    m.aClusterSwingWeights(3) = dWeight
                ElseIf strName = strIntermediate Then
                    Parm(3) = nStrength
                    m.aClusterSwingWeights(2) = dWeight
                ElseIf strName = strShort Then
                    Parm(4) = nStrength
                    m.aClusterSwingWeights(1) = dWeight
                End If
            End If
        Next
        If Not m.Chart Is Nothing Then m.Chart.RedoMode = eRedo5_RecalcInd
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.ClusterSwingInfoSet"
    
End Sub

Public Function ClusterRatioInfoGet() As cGdTable
On Error GoTo ErrSection:

    Set ClusterRatioInfoGet = m.tbClusterRatios

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cIndicator.ClusterRatioInfoGet"
    
End Function

Public Sub ClusterRatioInfoSet(tb As cGdTable)
On Error GoTo ErrSection:

    Dim i&, d1#, d2#
    Dim aIndex As cGdArray

    If tb Is Nothing Then Exit Sub
    If tb.NumRecords <= 0 Then Exit Sub
    
    Set aIndex = tb.CreateSortedIndex(0, eGdSort_Default Or eGdSort_DeleteDuplicates Or eGdSort_Stable)
    If aIndex Is Nothing Then GoTo ErrExit
    If aIndex.Size <> tb.NumRecords Then GoTo ErrExit
        
    Set m.tbClusterRatios = New cGdTable
    m.tbClusterRatios.CreateField eGDARRAY_Doubles, 0, 0#
    m.tbClusterRatios.CreateField eGDARRAY_Doubles, 1, 0#
    
    For i = 0 To tb.NumRecords - 1
        d1 = tb(0, aIndex(i))
        d2 = tb(1, aIndex(i))
                    
        If d1 <> 0# And d2 <> 0# Then
            m.tbClusterRatios.AddRecord ""
            m.tbClusterRatios(0, i) = d1
            m.tbClusterRatios(1, i) = d2
        End If
    Next

    If m.tbClusterRatios.NumRecords > 0 And Not m.Chart Is Nothing Then
        m.Chart.RedoMode = eRedo5_RecalcInd
    End If
    
ErrExit:
    Set aIndex = Nothing
    Exit Sub

ErrSection:
    RaiseError "cIndicator.ClusterRatioInfoSet"
    
End Sub

Private Sub RatioStringToTable(ByVal strRatios$, ByVal strWeights$)
On Error GoTo ErrSection:

    Dim i&, j&
    Dim aRatios As New cGdArray
    Dim aWeights As New cGdArray
    
    aRatios.SplitFields strRatios, "|"
    aWeights.SplitFields strWeights, "|"
    
    If aRatios.Size = 0 And aWeights.Size = 0 Then Exit Sub
    
    If aRatios.Size <> aWeights.Size Then
        DebugLog "Cluster ratio string err: " & strRatios & " : " & strWeights
    End If
    
    'precaustionary check in case arrays are not same size
    If aRatios.Size <= aWeights.Size Then
        j = aRatios.Size - 1
    Else
        j = aWeights.Size - 1
    End If
    
    If j <= 0 Then Exit Sub
    
    Set m.tbClusterRatios = New cGdTable
    m.tbClusterRatios.CreateField eGDARRAY_Doubles, 0, 0#
    m.tbClusterRatios.CreateField eGDARRAY_Doubles, 1, 0#
    
    For i = 0 To j
        m.tbClusterRatios.AddRecord ""
        m.tbClusterRatios(0, i) = ValOfText(aRatios(i))
        m.tbClusterRatios(1, i) = ValOfText(aWeights(i))
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.RatioStringToTable"
    
End Sub

Private Sub geSetFontInfo(ByVal nIdx&)
On Error GoTo ErrSection:

    m.geIndStruct.altPenSize = m.nFontSize
    gdSetStr m.geIndStruct.gdshImage, nIdx, m.strFontName
    
    If m.nFontBold = 1 And m.nFontItalic = 1 Then
        m.geIndStruct.altPenStyle = 3
    ElseIf m.nFontBold = 1 Then
        m.geIndStruct.altPenStyle = 1
    ElseIf m.nFontItalic = 1 Then
        m.geIndStruct.altPenStyle = 2
    Else
        m.geIndStruct.altPenStyle = 0
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.geSetFontInfo"

End Sub

Public Sub SaveDefaults()
On Error GoTo ErrSection:
'JM: As of 12-09-2010 this function is valid only for the price & time fib cluster indicators
'    Retracement/Expansion not exposed to user in UI and are always true

    Dim aTemp As cGdArray
    
    If m.nClusterProximity > 99 Then
        m.nClusterProximity = 99
    ElseIf m.nClusterProximity < 0 Then
        m.nClusterProximity = 0
    End If

    If m.eDisplayType = eINDIC_ClusterPrice Then
    
        SetIniFileProperty "PriceColor", m.geIndStruct.labelColor, kClusterIniSection, g.strIniFile
        SetIniFileProperty "PriceProximity", m.nClusterProximity, kClusterIniSection, g.strIniFile
        SetIniFileProperty "UseResistance", m.bResistanceOrL2H, kClusterIniSection, g.strIniFile        'used to be rallies
        SetIniFileProperty "UseSupport", m.bSupportOrH2L, kClusterIniSection, g.strIniFile              'used to be declines
        SetIniFileProperty "UseABCs", m.bClusterABCs, kClusterIniSection, g.strIniFile
        SetIniFileProperty "ClusterMaxBars", m.nClusterMaxBars, kClusterIniSection, g.strIniFile
    
        SetIniFileProperty "ClusterSwingFontName", m.strFontName, kClusterIniSection, g.strIniFile
        SetIniFileProperty "ClusterSwingFontBold", m.nFontBold, kClusterIniSection, g.strIniFile
        SetIniFileProperty "ClusterSwingFontItalic", m.nFontItalic, kClusterIniSection, g.strIniFile
        SetIniFileProperty "ClusterSwingFontSize", m.nFontSize, kClusterIniSection, g.strIniFile
        SetIniFileProperty "ClusterSwingStyle", m.eStyle, kClusterIniSection, g.strIniFile
        
        SetIniFileProperty "ClusterSwingLabel", m.nClusterSwingLabel, kClusterIniSection, g.strIniFile
        SetIniFileProperty "ClusterSwingColorL", m.nClusterSwingColorL, kClusterIniSection, g.strIniFile
        SetIniFileProperty "ClusterSwingColorM", m.nClusterSwingColorM, kClusterIniSection, g.strIniFile
        SetIniFileProperty "ClusterSwingColorS", m.nClusterSwingColorS, kClusterIniSection, g.strIniFile
        
        SetIniFileProperty "ClusterSwingParm2", Parm(2), kClusterIniSection, g.strIniFile       'swing point strength L
        SetIniFileProperty "ClusterSwingParm3", Parm(3), kClusterIniSection, g.strIniFile       'swing point strength M
        SetIniFileProperty "ClusterSwingParm4", Parm(4), kClusterIniSection, g.strIniFile       'swing point strength S
        
        If Not m.aClusterSwingWeights Is Nothing Then
            SetIniFileProperty "ClusterSwingWeights", m.aClusterSwingWeights.JoinFields("|"), kClusterIniSection, g.strIniFile
        End If
        
        If Not m.tbClusterRatios Is Nothing Then
            Set aTemp = m.tbClusterRatios.FieldArray(0)
            If Not aTemp Is Nothing Then
                SetIniFileProperty "PriceFibRatios", aTemp.JoinFields("|"), kClusterIniSection, g.strIniFile
            End If
            Set aTemp = m.tbClusterRatios.FieldArray(1)
            If Not aTemp Is Nothing Then
                SetIniFileProperty "PriceFibWeights", aTemp.JoinFields("|"), kClusterIniSection, g.strIniFile
            End If
        End If
        
    ElseIf m.eDisplayType = eINDIC_ClusterTime Then
        
        SetIniFileProperty "TimeColor", m.nColor, kClusterIniSection, g.strIniFile
        SetIniFileProperty "TimeProximity", m.nClusterProximity, kClusterIniSection, g.strIniFile
        SetIniFileProperty "HighToHigh", m.bExpanOrH2H, kClusterIniSection, g.strIniFile
        SetIniFileProperty "HighToLow", m.bSupportOrH2L, kClusterIniSection, g.strIniFile
        SetIniFileProperty "LowToHigh", m.bResistanceOrL2H, kClusterIniSection, g.strIniFile
        SetIniFileProperty "LowToLow", m.bRetraceOrL2L, kClusterIniSection, g.strIniFile
        SetIniFileProperty "ClusterBarsBack", m.nClusterMaxBars, kClusterIniSection, g.strIniFile
    
        If Not m.tbClusterRatios Is Nothing Then
            Set aTemp = m.tbClusterRatios.FieldArray(0)
            If Not aTemp Is Nothing Then
                SetIniFileProperty "TimeFibRatios", aTemp.JoinFields("|"), kClusterIniSection, g.strIniFile
            End If
            Set aTemp = m.tbClusterRatios.FieldArray(1)
            If Not aTemp Is Nothing Then
                SetIniFileProperty "TimeFibWeights", aTemp.JoinFields("|"), kClusterIniSection, g.strIniFile
            End If
        End If

    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.SaveDefaults"

End Sub

Public Property Get SeasonalIndFlag() As Long
    SeasonalIndFlag = m.nIndSeasonal
End Property

Public Property Let SeasonalIndFlag(ByVal nFlag&)
    m.nIndSeasonal = nFlag
End Property

Private Sub SetVwapBreak(ByRef d1#, ByVal d2#, ByVal i1&, ByVal eBreakType As eVwapBreak)
On Error GoTo ErrSection:

    Dim nInterval&, dPivot1#, dPivot2#
    Dim hXBar&, hDataArray&
    
    Dim PivotVWAP As cIndicator
    
    nInterval = Val(Parm(2))     'this is # of days, weeks, months etc. to break
    

'JM 08-10-2011: Design note
'   interval limitations as set by Heath - Yr=1,Qrtr=2,Mo/Wk=4,Day=5
    
    Select Case eBreakType
        Case eVwapBreak7_Yearly
            nInterval = 1
        Case eVwapBreak6_Quarterly
            If nInterval > 2 Then nInterval = 2
        Case eVwapBreak5_Monthly, eVwapBreak4_Weekly
            If nInterval > 4 Then nInterval = 4
        Case eVwapBreak3_Daily
            If nInterval > 5 Then nInterval = 5
    End Select
    
    If nInterval <= 0 Then nInterval = 1        'precautionary, theoretically should never happen
    
    If nInterval > 1 Then
        Set PivotVWAP = m.Chart.PivotVwapIndicator(nInterval)
        If PivotVWAP Is Nothing Then nInterval = 1
    End If
    
    If nInterval > 1 Then
        hXBar = m.Chart.aXBar.ArrayHandle
        hDataArray = PivotVWAP.Data.ArrayHandle
        
        dPivot1 = gdGetNum(hDataArray, gdGetNum(hXBar, i1))
        dPivot2 = gdGetNum(hDataArray, gdGetNum(hXBar, i1 - 1))
        
        If dPivot1 = dPivot2 Then
            d1 = d2
        Else
            d1 = -d2
        End If
    Else
        Select Case eBreakType
            Case eVwapBreak4_Weekly
                If WkNum(d1) = WkNum(d2) Then d1 = d2
            Case eVwapBreak5_Monthly
                If Month(d1) = Month(d2) Then d1 = d2
            Case eVwapBreak6_Quarterly
                If Int((Month(d1) - 1) / 3) = Int((Month(d2) - 1) / 3) Then d1 = d2
            Case eVwapBreak7_Yearly
                 If Year(d1) = Year(d2) Then d1 = d2
        End Select
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.SetVwapBreak"

End Sub

Public Property Get HideWick() As Boolean
On Error Resume Next

    HideWick = m.bHideWick

End Property

Public Property Let HideWick(ByVal bHideWick As Boolean)
On Error Resume Next

    m.bHideWick = bHideWick

End Property

Public Property Get Flip() As Boolean
On Error GoTo ErrSection:

    Flip = m.bFlip

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.Flip.Get"

End Property

Public Property Let Flip(ByVal bFlip As Boolean)
On Error GoTo ErrSection:

    m.bFlip = bFlip

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.Flip.Let"

End Property

Public Property Get SyncOverlayFlip() As Boolean
    SyncOverlayFlip = m.bSyncOverlayFlip
End Property

Public Property Let SyncOverlayFlip(ByVal bSync As Boolean)
    m.bSyncOverlayFlip = bSync
End Property

Public Sub UpdateAlertParms()
On Error GoTo ErrSection:
    
    Dim i&
    Dim strText$, strKeyword$
    Dim strUpdate$, strExprOld$, strLabelOld$
    
    Dim Tree As cGdTree
    Dim Ind As cIndicator
    Dim bUpdateParm As Boolean

    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If m.Chart Is Nothing Then Exit Sub
    If Len(m.strCodedName) = 0 Then Exit Sub
    
    Set Tree = m.Chart.Tree
    If Tree Is Nothing Then Exit Sub
    
    'special case for ART Pyramid & ART Reverse Bars:
    'alerts are tied to the highlight bars, but logically associated with the ART Pyramid or Reverse Bars
    If InStr(m.strCodedName, "PyramidTP") <> 0 Then
        strKeyword = "Pyramid"
    ElseIf InStr(m.strCodedName, "ReverseBarTP") <> 0 Then
        strKeyword = "ARTReverseBars"
    End If
    
    strExprOld = m.strExpression
    strLabelOld = m.strChartLabel
    
    'compare the coded text of any Ind-type alert against this indicator's coded name or special key word
    For i = Tree.Count To 0 Step -1
        If Tree.NodeLevel(i) > 0 Then
            bUpdateParm = False
            Set Ind = Tree(i)
            If Not Ind Is Nothing Then
                If Ind.IsAlert Then
                    strText = Ind.CodedText
                    
                    If InStr(strText, m.strCodedName) <> 0 Then
                        bUpdateParm = True
                    ElseIf Len(strKeyword) > 0 Then
                        If InStr(strText, strKeyword) <> 0 Then bUpdateParm = True
                    End If
                    
                    If bUpdateParm Then
                        'force rebuild of expression to pick up parm changes
                        If BuildText(m.Chart, , True) Then
                            'replace matching portion of alert's indicator with rebuilt expression
                            strUpdate = Replace(Ind.Expression, strExprOld, m.strExpression)
                            Ind.Expression = strUpdate
                            
                            'replace matching portion of alert's condition builder string with rebuilt chart label
                            strUpdate = Replace(Ind.CondBuilderStr, strLabelOld, m.strChartLabel)
                            Ind.CondBuilderStr = strUpdate
                            
                            'force alert's indicator to rebuild its alert string
                            Ind.InitAlert m.Chart, True
                            
                            'clear alert's indicator coded text
                            Ind.CodedText = ""
                        End If
                    End If
                End If
            End If
        End If
    Next
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cIndicator.UpdateAlertParms"
End Sub

Public Property Get TicksPerRow() As Long
    TicksPerRow = m.nTicksPerRow
End Property

Public Property Let TicksPerRow(ByVal nTicks&)
    If nTicks = -1 Or nTicks > 0 Then       '-1 = auto
        m.nTicksPerRow = nTicks
    End If
End Property

Public Property Get SessionsPerProfile() As Long
    SessionsPerProfile = m.nSessionsPerProfile
End Property

Public Property Let SessionsPerProfile(ByVal nCount&)
    'for multi-day, profile must have more than 1 sessions per profile
    If nCount > 1 Then
        m.nSessionsPerProfile = nCount
    End If
End Property

Public Property Get ProfileStyleTPO() As Long
    ProfileStyleTPO = m.nStyleTPO
End Property

Public Property Let ProfileStyleTPO(ByVal nStyle&)
    'grapheng.dll enum
    'typedef enum{
    '    e_StandardAlpha,        //TPO letter or block style for trade profile chart
    '    e_ClassicAlpha,
    '    e_blocks,
    '    e_vp_filled,            //filled rectangles for volume profile on bar chart displayed as volume at price
    '    e_vp_hollow,
    '    e_vp_outline,
    '    e_vt_filled,            //filled blocks for volume profile on bar chart displayed as time at price
    '    e_vt_hollow
    '}PROFILE_CHAR_STYLE;
    
    m.nStyleTPO = nStyle

End Property

Public Property Get ProfileShowHide(eProperty As ePCStruct) As Long
    
    Select Case eProperty
        'TPO show/hide
        Case ePCStruct_TPO
            ProfileShowHide = m.nShowProfileTPO
        Case ePCStruct_TPO_ColorTo
            If m.nColorProfileFrom = m.nColorProfileTo Then
                ProfileShowHide = 0
            Else
                ProfileShowHide = m.nShowProfileTPO
            End If
        Case ePCStruct_TPO_POC
            ProfileShowHide = m.nShowProfileTpoPOC
        Case ePCStruct_TPO_VA
            ProfileShowHide = m.nShowProfileTpoVA
        'volume show/hide
        Case ePCStruct_Volume
            ProfileShowHide = m.nShowProfileVol
        Case ePCStruct_Volume_POC
            ProfileShowHide = m.nShowProfileVolPOC
        Case ePCStruct_Volume_VA
            ProfileShowHide = m.nShowProfileVolVA
        Case ePCStruct_Open
            ProfileShowHide = m.nShowProfileOC
        Case ePCStruct_Close
            If m.eDataType = eINDIC_ProfileVolume Then
                ProfileShowHide = m.nProfileOnRight
            Else
                ProfileShowHide = m.nShowProfileOC
            End If
    End Select
    
End Property

Public Property Let ProfileShowHide(eProperty As ePCStruct, ByVal nShow&)
On Error GoTo ErrSection:
    
    Select Case eProperty
        'TPO show/hide
        Case ePCStruct_TPO
            m.nShowProfileTPO = nShow
        Case ePCStruct_TPO_POC
            m.nShowProfileTpoPOC = nShow
        Case ePCStruct_TPO_VA
            m.nShowProfileTpoVA = nShow
        'volume show/hide
        Case ePCStruct_Volume
            m.nShowProfileVol = nShow
        Case ePCStruct_Volume_POC
            'JM 09-04-2015: for trade profile this value is: 0=hide, 1=show
            '   for volume profile this value can be 0,1,2 or 3
            '       0 = hide current - hide previous
            '       1 = show current - hide previous
            '       2 = show current - show previous
            '       3 = hide current - show previous
            m.nShowProfileVolPOC = nShow
        Case ePCStruct_Volume_VA
            'JM 09-04-2015: for trade profile this value is: 0=hide, 1=show
            '   for volume profile this value can be 0,1,2 or 3 (same as for show/hid POC)
            m.nShowProfileVolVA = nShow
        Case ePCStruct_Open
            'JM 09-08-2015: for trade profile this value is: 0/1 = hide/show profile open as triangle on left
            '   for volume profile this value is:
            '       0 = hide current values - hide previous values
            '       1 = show current values - hide previous values
            '       2 = show current values - show previous values
            '       3 = hide current values - show previous values
            m.nShowProfileOC = nShow
        Case ePCStruct_Close
            If m.eDataType = eINDIC_ProfileVolume Then
                m.nProfileOnRight = nShow
            Else
                m.nShowProfileOC = nShow
            End If
    End Select
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cIndicator.ProfileShowHide.Let"

End Property

Public Property Get ProfileColor(eProperty As ePCStruct) As Long
    
    Select Case eProperty
        'TPO colors are color scheme dependent (use same defaults as trade profile)
        Case ePCStruct_TPO
            If m.eColorProfileScheme = MktProf_Color_OpenClose Then
                ProfileColor = m.nColorProfileUp
            Else
                ProfileColor = m.nColorProfileFrom
            End If
        Case ePCStruct_TPO_ColorTo
            If m.eColorProfileScheme = MktProf_Color_OpenClose Then
                ProfileColor = m.nColorProfileDown
            Else
                ProfileColor = m.nColorProfileTo
            End If
        Case ePCStruct_TPO_POC
            ProfileColor = m.nColorProfileTpoPOC
        Case ePCStruct_TPO_VA
            ProfileColor = m.nColorProfileTpoVA
        Case ePCStruct_Open, ePCStruct_Close
            ProfileColor = m.nColorProfileOC
        'volume oolors
        Case ePCStruct_Volume
            ProfileColor = m.nColorProfileVol
        Case ePCStruct_Volume_POC
            ProfileColor = m.nColorProfileVolPOC
        Case ePCStruct_Volume_VA
            ProfileColor = m.nColorProfileVolVA
    End Select
    
End Property

Public Property Let ProfileColor(eProperty As ePCStruct, ByVal nColor&)
    
    Select Case eProperty
        'TPO colors
        Case ePCStruct_TPO
            If m.eColorProfileScheme = MktProf_Color_Gradient Or m.eDataType = eINDIC_ProfileVolume Then
                m.nColorProfileFrom = nColor
            ElseIf m.eColorProfileScheme = MktProf_Color_OpenClose Then
                m.nColorProfileUp = nColor
            End If
        Case ePCStruct_TPO_ColorTo
            If m.eColorProfileScheme = MktProf_Color_Gradient Or m.eDataType = eINDIC_ProfileVolume Then
                m.nColorProfileTo = nColor
            ElseIf m.eColorProfileScheme = MktProf_Color_OpenClose Then
                m.nColorProfileDown = nColor
            End If
        Case ePCStruct_TPO_POC
            m.nColorProfileTpoPOC = nColor
        Case ePCStruct_TPO_VA
            m.nColorProfileTpoVA = nColor
        Case ePCStruct_Open, ePCStruct_Close
            m.nColorProfileOC = nColor
        'volume oolors
        Case ePCStruct_Volume
            m.nColorProfileVol = nColor
        Case ePCStruct_Volume_POC
            m.nColorProfileVolPOC = nColor
        Case ePCStruct_Volume_VA
            m.nColorProfileVolVA = nColor
    End Select
    
End Property

Public Property Get ProfileParm(eProperty As ePCStruct) As Double

    Select Case eProperty
        Case ePCStruct_TPO_VA
            ProfileParm = m.dPercentTpoVA
        Case ePCStruct_Volume_VA
            ProfileParm = m.dPercentVolVA
    End Select

End Property

Public Property Let ProfileParm(eProperty As ePCStruct, ByVal dParm#)

    Select Case eProperty
        Case ePCStruct_TPO_VA
            m.dPercentTpoVA = dParm
        Case ePCStruct_Volume_VA
            m.dPercentVolVA = dParm
    End Select

End Property

Public Property Get ProfileColorScheme() As MktProfile_Char_ColorScheme
    
    ProfileColorScheme = m.eColorProfileScheme

End Property

Public Property Let ProfileColorScheme(ByVal eScheme As MktProfile_Char_ColorScheme)

    Select Case eScheme
        
        'disallow volume iterator scheme
        Case MktProf_Color_Gradient, MktProf_Color_Rainbow, MktProf_Color_OpenClose, MktProf_Color_BidAsk, MktProf_Color_NoScheme
            m.eColorProfileScheme = eScheme

    End Select

End Property

Public Property Get ProfilePeriodicityStr() As String
    ProfilePeriodicityStr = m.strProfilePeriodicity
End Property

Public Property Let ProfilePeriodicityStr(ByVal strPeriodicity$)
    m.strProfilePeriodicity = strPeriodicity
End Property

Public Sub PopulateCboProfilePeriodicity(cbo As ctlUniComboImageXP, txtPerBar As ctlUniTextBoxXP) 'RH was TextBox

    Dim i&, str1$, Str2$
    
    cbo.Clear
    
    If m.Bars Is Nothing Then
        cbo.AddItem "Daily"
        cbo.ListIndex = 0
        txtPerBar.Text = "1"
        Exit Sub
    End If
    
    If 0 = InStr(m.strProfilePeriodicity, " ") Then
        Str2 = m.strProfilePeriodicity
    Else
        str1 = Parse(m.strProfilePeriodicity, " ", 1)
        Str2 = Parse(m.strProfilePeriodicity, " ", 2)
    End If
    
    If Len(str1) > 0 Then
        txtPerBar.Text = str1
    Else
        txtPerBar.Text = "1"
    End If
    
    If m.Bars.IsIntraday Then
        cbo.AddItem "Daily"
        cbo.AddItem "Weekly"
        cbo.AddItem "Monthly"
    Else
        i = m.Bars.Prop(eBARS_PeriodType)
        
        Select Case i
            Case ePRD_Days
                cbo.AddItem "Daily"
                cbo.AddItem "Weekly"
                cbo.AddItem "Monthly"
                cbo.AddItem "Quarterly"
                cbo.AddItem "Yearly"
            Case ePRD_Weeks
                cbo.AddItem "Weekly"
                cbo.AddItem "Monthly"
                cbo.AddItem "Quarterly"
                cbo.AddItem "Yearly"
            Case ePRD_Months
                cbo.AddItem "Monthly"
                cbo.AddItem "Quarterly"
                cbo.AddItem "Yearly"
            Case ePRD_Quarters
                cbo.AddItem "Quarterly"
                cbo.AddItem "Yearly"
            Case ePRD_Years
                cbo.AddItem "Yearly"
        End Select
    End If
    
    str1 = UCase(Left(Str2, 1))
    
    For i = 0 To cbo.ListCount - 1
        If str1 = Left(cbo.List(i), 1) Then
            cbo.ListIndex = i
            Exit For
        End If
    Next
    
End Sub

Public Function ChkboxValForProfileStruct(eProperty As ePCStruct, ByVal bPrevious As Boolean) As Long
On Error GoTo ErrExit
    
    Dim chkboxval&
    
    chkboxval = vbUnchecked

    If m.eDataType = eINDIC_ProfileVolume Then
        If eProperty = ePCStruct_Volume_POC Then
            'check box for show/hide current/previous POC
            If m.nShowProfileVolPOC = 2 Then
                chkboxval = vbChecked
            ElseIf bPrevious Then
                If m.nShowProfileVolPOC = 3 Then
                    chkboxval = vbChecked
                Else
                    chkboxval = vbUnchecked
                End If
            ElseIf m.nShowProfileVolPOC = 1 Then
                chkboxval = vbChecked
            End If
        ElseIf eProperty = ePCStruct_Volume_VA Then
            'check box for show/hide current/previous VA
            If m.nShowProfileVolVA = 2 Then
                chkboxval = vbChecked
            ElseIf bPrevious Then
                If m.nShowProfileVolVA = 3 Then
                    chkboxval = vbChecked
                Else
                    chkboxval = vbUnchecked
                End If
            ElseIf m.nShowProfileVolVA = 1 Then
                chkboxval = vbChecked
            End If
        ElseIf eProperty = ePCStruct_Open Then
            'check box for show/hide current/previous VA, POC values
            If m.nShowProfileVolPOC = 0 And m.nShowProfileVolVA = 0 Then
                chkboxval = vbGrayed
            ElseIf bPrevious Then
                If m.nShowProfileVolPOC > 1 Or m.nShowProfileVolVA > 1 Then
                    If m.nShowProfileOC > 1 Then
                        chkboxval = vbChecked
                    Else
                        chkboxval = vbUnchecked
                    End If
                Else
                    chkboxval = vbGrayed
                End If
            Else
                If m.nShowProfileVolPOC = 1 Or m.nShowProfileVolPOC = 2 Or _
                   m.nShowProfileVolVA = 1 Or m.nShowProfileVolVA = 2 Then
                    If m.nShowProfileOC = 1 Or m.nShowProfileOC = 2 Then
                        chkboxval = vbChecked
                    Else
                        chkboxval = vbUnchecked
                    End If
                Else
                    chkboxval = vbGrayed
                End If
            End If
        End If
    End If
    
    ChkboxValForProfileStruct = chkboxval
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cIndicator.ChkboxValForProfileStruct.Let"

End Function
