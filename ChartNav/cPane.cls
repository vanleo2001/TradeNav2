VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum ePANE_DisplayFormat
    ePANE_AutoFormat = 0
    ePANE_PriceFormat = 1
    ePANE_RoundDecimals = 2
End Enum

Public Enum ePANE_LogFlag
    ePANE_LogFlagLog = -1
    ePANE_LogFlagLinear = 0
    ePANE_LogFlagPercent = 1
End Enum

'Note: These numbers match the enumerated type in grapheng.dll
Public Enum ePANE_ScaleType
    ePANE_ScaleNormal = 0       'normal
    ePANE_ScaleLog = 1          'log
    ePANE_ScaleSquare = 2       'normal scale, square the chart
    ePANE_ScalePrice = 3        'normal scale, fit to price
    ePANE_ScaleLogDraw = 4      'log AND draw mode log
    ePANE_DoNotUse = 99         'reserved for internal use by Graphics Engine
End Enum

Public Enum ePANE_ScaleMode
    ePANE_ScaleModeAuto = 0
    ePANE_ScaleModeManual = 1
    ePANE_ScaleModeSquare = 2
    ePANE_ScaleModeAutoPrice = 3
End Enum

'this is to determine what data to show in split-pane area
Public Enum ePANE_SplitPaneType
    ePANE_SplitPaneNone = 0
    ePANE_SplitPaneWood = 1
    ePANE_SplitPaneTimer = 2
    ePANE_SplitPaneCluster = 3
    ePANE_SplitPaneOptGraph = 4
End Enum

Private Type mPrivate
    bDisplay As Boolean
    bLogFlagChecked As Boolean
    bCanBeNonLinear As Boolean
    
    strName As String
    strDesc As String
    dSize As Double                 'proportion of pane as percent of chart's height
    
    eScaleMode As ePANE_ScaleMode
    eScaleAutoLastUsed As ePANE_ScaleMode
    eLogFlag As ePANE_LogFlag       '-1=log, 0=linear, 1=percent change
    
    dMax As Double                  'for scale mode manual
    dMin As Double                  'for scale mode manual
    dMaxSave As Double              'for stretching y-scale
    dMinSave As Double              'for stretching y-scale
    nDisplayFormat As ePANE_DisplayFormat
    nDisplayDecimals As Long
    nHideSeparator As Long
    nPtsOrTicksFlag As Long          '0=points,1=ticks display format
    dSquareTicks As Double           'for calculating points/ticks per bar(s) when squaring chart
    nSquareBars As Long              'for calculating points per bar(s) when squaring chart
    nYscaleLabelAll As Long
    nRestoreTimerPane As Long
    nPaneSeasonal As Long            'not zero = seasonal chart
    
    'This is the index in the tree of the PriceBars indicator
    'that is "primarily associated" with this pane (for scaling
    'purposes, etc.) -- it will be 0 if the pane is the "Price Pane"
    'or if there is no bars data in this pane, else it will be the
    'index in the tree of the indicator in this pane with the bars data
    idxPriceBars As Long
        
    eSplitPaneType As ePANE_SplitPaneType
    oWoodData As cPaneWood
    
    gePaneStruct As chart_pane
    geAdded As Boolean
End Type
Private m As mPrivate

Private Sub Class_Initialize()
On Error GoTo ErrSection:
        
    m.eScaleMode = ePANE_ScaleModeAuto
    m.eLogFlag = ePANE_LogFlagLinear
    m.nHideSeparator = 0
    m.dMax = 100
    m.dMin = 0
    m.nDisplayFormat = ePANE_AutoFormat
    m.nDisplayDecimals = 2

    m.nSquareBars = g.ChartGlobals.nSquareBars
    m.dSquareTicks = g.ChartGlobals.dSquareTicks
    If m.nSquareBars <= 0 Then m.nSquareBars = 1
    If m.dSquareTicks <= 0 Then m.dSquareTicks = 100
    
    m.gePaneStruct.y_scaleMin = 0
    m.gePaneStruct.y_scaleMax = 100
    
    m.gePaneStruct.y_spAboveRatio = 0.03      '0.05
    m.gePaneStruct.y_spBelowRatio = -0.03     '-0.05
    
    'array of strings for vertical split pane text (current usage: Woodies)
    'format: text|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|yPixel
    m.gePaneStruct.gdshSplitPane = gdCreateArray(eGDARRAY_Strings)
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.Intialize", eGDRaiseError_Raise

End Sub

Public Sub TemplateSave(ByVal fh%)
On Error GoTo ErrSection:

    If fh = 0 Then Exit Sub
    Print #fh, IniString("Display", m.bDisplay)
    Print #fh, IniString("Name", m.strName)
    Print #fh, IniString("Desc", m.strDesc)
    Print #fh, IniString("Size", Size)
    Print #fh, IniString("Scaling", m.eScaleMode)
    Print #fh, IniString("ScaleAutoLastUsed", m.eScaleAutoLastUsed)
    Print #fh, IniString("LogScale", m.eLogFlag)
    Print #fh, IniString("HiddenSeparator", m.nHideSeparator)
    Print #fh, IniString("MaxValue", m.dMax)
    Print #fh, IniString("MinValue", m.dMin)
    Print #fh, IniString("SpaceAboveRatio", m.gePaneStruct.y_spAboveRatio)
    Print #fh, IniString("SpaceBelowRatio", m.gePaneStruct.y_spBelowRatio)
    Print #fh, IniString("DisplayFormat", m.nDisplayFormat)
    Print #fh, IniString("DisplayDecimals", m.nDisplayDecimals)
    Print #fh, IniString("PtsOrTicksFlag", m.nPtsOrTicksFlag)
    Print #fh, IniString("YscaleLabelAll", m.nYscaleLabelAll)
    
    If m.nPaneSeasonal <> 0 Then Print #fh, IniString("PaneSeasonal", 1)
    If Not m.oWoodData Is Nothing Then m.oWoodData.TemplateSave Me, fh
    
    Print #fh, "END="
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.TemplateSave", eGDRaiseError_Raise

End Sub

Public Sub TemplateLoad(ByVal fh%)
On Error GoTo ErrSection:

    Dim i&, strProp$, strValue$, dValue#
    
    m.eSplitPaneType = ePANE_SplitPaneNone
    If Not m.oWoodData Is Nothing Then m.oWoodData.Reset
    
    If fh = 0 Then Exit Sub
    Do While Not EOF(fh)
        Line Input #fh, strProp
        If ParseIniString(strProp, strValue, dValue) Then
            Select Case UCase(strProp)
                Case "END"
                    Exit Do
                Case "DISPLAY"
                    m.bDisplay = dValue
                Case "NAME"
                    m.strName = strValue
                    If strValue = "PRICE" Then '(from older version)
                        m.nDisplayFormat = ePANE_PriceFormat '(default for price pane)
                    End If
                Case "DESC"
                    m.strDesc = strValue
                Case "SIZE"
                    Size = dValue
                Case "SCALING"
                    If dValue = -1 Then dValue = 1
                    m.eScaleMode = dValue
                Case "SCALEAUTOLASTUSED"
                    If dValue = 0 Or dValue = 1 Then m.eScaleAutoLastUsed = dValue
                Case "LOGSCALE"
                    Select Case dValue
                        Case ePANE_LogFlagLog, ePANE_LogFlagLinear, ePANE_LogFlagPercent
                            m.eLogFlag = dValue
                        Case Else
                            m.eLogFlag = ePANE_LogFlagLinear
                    End Select
                Case "HIDDENSEPARATOR"
                    m.nHideSeparator = dValue
                Case "MAXVALUE"
                    m.dMax = dValue
                Case "MINVALUE"
                    m.dMin = dValue
                Case "DISPLAYFORMAT"
                    Select Case dValue
                        Case ePANE_AutoFormat, ePANE_RoundDecimals, ePANE_PriceFormat
                            m.nDisplayFormat = dValue
                        Case Else
                            m.nDisplayFormat = ePANE_PriceFormat
                    End Select
                Case "DISPLAYDECIMALS"
                    m.nDisplayDecimals = dValue
                Case "PTSORTICKSFLAG"
                    m.nPtsOrTicksFlag = dValue
                Case "YSCALELABELALL"
                    m.nYscaleLabelAll = dValue
                    m.gePaneStruct.y_labelAll = dValue
                Case "SPACEABOVERATIO"
                    m.gePaneStruct.y_spAboveRatio = dValue
                Case "SPACEBELOWRATIO"
                    m.gePaneStruct.y_spBelowRatio = dValue
                Case "PANESEASONAL"
                    m.nPaneSeasonal = dValue
                Case Else:
                    'If FileExist("c:\COMMON\files.exe") Then
                    '    InfBox "i=[] ; h=Pane Template ; Not found: " & strType
                    'End If
                    If InStr(UCase(strProp), "PANEWOOD") <> 0 Then
                        If m.oWoodData Is Nothing Then
                            Set m.oWoodData = New cPaneWood
                            m.oWoodData.Reset
                        End If
                        If InStr(strValue, "PP_Price") <> 0 Then
                            m.eSplitPaneType = ePANE_SplitPaneTimer
                        ElseIf InStr(strValue, "PP_Cluster") <> 0 Then
                            m.eSplitPaneType = ePANE_SplitPaneCluster
                        ElseIf InStr(strValue, "SideWinder") <> 0 Then
                            m.eSplitPaneType = ePANE_SplitPaneWood
                        End If
                        m.oWoodData.TemplateLoad UCase(strProp), strValue, dValue
                    End If
            End Select
        End If
    Loop
        
    ' from older version, change from all uppercase if length is > 4
    If UCase(m.strName) = m.strName And Len(m.strName) > 4 _
            And Len(m.strDesc) = 0 Then
        m.strName = FileNameDisplay(m.strName)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.TemplateLoad", eGDRaiseError_Raise

End Sub

Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    ' get copy of all private data
    m = mCopyFrom

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.SetPrivateData", eGDRaiseError_Raise

End Sub

Public Function MakeCopy() As cPane
On Error GoTo ErrSection:

    Dim aCopy As New cPane
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPane.MakeCopy", eGDRaiseError_Raise

End Function

Public Property Get Display() As Boolean
    Display = m.bDisplay
End Property

Public Property Let Display(ByVal bDisplay As Boolean)
    
    'Verifypanes in cChart.cls sets pane display to false when all indicators are hidden
    
    m.bDisplay = bDisplay
    'MJM [begin] - added for grapheng.dll
    If bDisplay = True Then
        m.gePaneStruct.paneShow = 1
    Else
        m.gePaneStruct.paneShow = 0
    End If
    'MJM [end] - added for grapheng.dll
End Property

Public Property Get PricePaneFlag() As Long
    PricePaneFlag = m.gePaneStruct.IsPricePane
End Property

Public Property Let PricePaneFlag(ByVal nFlag&)
    m.gePaneStruct.IsPricePane = nFlag
End Property

Public Property Get Name() As String
    Name = m.strName
End Property

Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get Desc() As String
    Desc = m.strDesc
End Property

Public Property Let Desc(ByVal strDesc As String)
    m.strDesc = strDesc
End Property

Public Property Get Size() As Double
    Size = m.dSize
End Property

Public Property Let Size(ByVal dSize As Double)
    If dSize < 0.025 Then dSize = 0.025         '4487
    m.dSize = dSize
    m.gePaneStruct.htProportion = dSize * 100
End Property

Public Property Get Max() As Double
    Max = m.dMax
End Property

Public Property Let Max(ByVal dMax As Double)
    
    'Develper's note:
    'm.dMax is the value shown in the edit chart form and is set as the pane's
    'max y-value only if the scaling type is manual.
    m.dMax = dMax
    
End Property

Public Property Get Min() As Double
    Min = m.dMin
End Property

Public Property Let Min(ByVal dMin As Double)
    
    'Develper's note:
    'm.dMin is the value shown in the edit chart form and is set as the pane's
    'min y-value only if the scaling type is manual.
    
    m.dMin = dMin
        
End Property

Public Property Get ScaleAutoLastUsed() As ePANE_ScaleMode
    If m.eScaleAutoLastUsed = ePANE_ScaleModeAuto Or m.eScaleAutoLastUsed = ePANE_ScaleModeAutoPrice Then
        ScaleAutoLastUsed = m.eScaleAutoLastUsed
    Else
        ScaleAutoLastUsed = ePANE_ScaleModeAuto
    End If
End Property

Public Property Get Scaling() As ePANE_ScaleMode
    Scaling = m.eScaleMode
End Property

Public Property Let Scaling(ByVal eScaleMode As ePANE_ScaleMode)
    
    If eScaleMode = ePANE_ScaleModeAuto Or eScaleMode = ePANE_ScaleModeAutoPrice Then
        m.eScaleAutoLastUsed = eScaleMode               'save last used auto scale mode
    Else
        If m.eScaleMode = ePANE_ScaleModeAuto Or m.eScaleMode = ePANE_ScaleModeAutoPrice Then
            m.eScaleAutoLastUsed = m.eScaleMode         'save last used auto scale mode
        End If
    End If
    m.eScaleMode = eScaleMode
    
    'JM: 06-06-2008 (aardvark 4534 - do away with temporary space)
    'max/min ratio used to be "temporary space" above/below and was cleared based on user interaction
    If m.eScaleMode <> ePANE_ScaleModeManual Then
        geClearMaxRatio
        geClearMinRatio
    End If

End Property

Public Property Get PaneLogFlag() As ePANE_LogFlag
    PaneLogFlag = m.eLogFlag
End Property

Public Property Let PaneLogFlag(ByVal eFlag As ePANE_LogFlag)
    
    'Developer's note:
    'Scale types on the graphics engine side are: 0=normal, 1=log, 2=square chart, 3=do not use
    'Scale modes on the VB side are: 0=auto, 1=manual, 2=square
    '   log scale is a true/false flag valid only if pane's min > 0 and
    '   scale mode is auto or manual
    'The graphics engine does not care whether the scale is auto or manual since
    'that only affects the min and max y-values of the pane which are preset
    'on the VB side.

    m.eLogFlag = eFlag

End Property

Public Property Get PointsOrTicksFlag() As Long
    PointsOrTicksFlag = m.nPtsOrTicksFlag
End Property

Public Property Let PointsOrTicksFlag(ByVal nVal&)
    m.nPtsOrTicksFlag = nVal
End Property

Public Property Get SquareTicks(Optional Chart As cChart = 0) As Double

    SquareTicks = m.dSquareTicks
    
End Property

Public Property Let SquareTicks(Chart As cChart, ByVal dVal#)
    
    If dVal > 0# And m.dSquareTicks <> dVal Then
        m.dSquareTicks = dVal
        g.ChartGlobals.dSquareTicks = dVal
        SaveCustomSquare Chart
    End If

End Property

Public Property Get SquareBars(Optional Chart As cChart = 0) As Long
    
    SquareBars = m.nSquareBars

End Property

Public Property Let SquareBars(Chart As cChart, ByVal nVal&)
    
    If nVal > 0 And m.nSquareBars <> nVal Then
        m.nSquareBars = nVal
        g.ChartGlobals.nSquareBars = nVal
        SaveCustomSquare Chart
    End If

End Property

Public Property Get HideSeparator() As Long
    'This is a temp fix for the "Enable Tabs" error bug in LW Weekly Setup Template.
    'TODO: Find out when, where, why & how this got set to -1! 01-23-2003
    If m.nHideSeparator < 0 Then
        m.nHideSeparator = 0
    End If
    HideSeparator = m.nHideSeparator
End Property

Public Property Let HideSeparator(ByVal Hide As Long)
    If Hide < 0 Then
        Hide = 0
    End If
    m.nHideSeparator = Hide
    m.gePaneStruct.paneSepHide = m.nHideSeparator
End Property

Public Property Get idxPriceBars() As Long
    idxPriceBars = m.idxPriceBars
End Property

Public Property Let idxPriceBars(ByVal idxBars As Long)
    m.idxPriceBars = idxBars
End Property

Public Property Get DisplayFormat() As ePANE_DisplayFormat
    If m.nDisplayFormat < ePANE_AutoFormat Or m.nDisplayFormat > ePANE_RoundDecimals Then m.nDisplayFormat = ePANE_PriceFormat
    DisplayFormat = m.nDisplayFormat
End Property

Public Property Let DisplayFormat(ByVal nDisplayFormat As ePANE_DisplayFormat)
    If nDisplayFormat >= ePANE_AutoFormat And nDisplayFormat <= ePANE_RoundDecimals Then
        m.nDisplayFormat = nDisplayFormat
    End If
End Property

Public Property Get DisplayDecimals() As Long
    DisplayDecimals = m.nDisplayDecimals
End Property

Public Property Let DisplayDecimals(ByVal nDisplayDecimals As Long)
    m.nDisplayDecimals = nDisplayDecimals
End Property

Public Property Get gePaneProportion() As Long
On Error GoTo ErrSection:
    
    gePaneProportion = m.gePaneStruct.htProportion

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.gePaneProportion (Get)", eGDRaiseError_Raise

End Property

Public Property Get geTopSep() As Long
On Error GoTo ErrSection:
    
    'this is the pane's top separator y value in pixels
    geTopSep = m.gePaneStruct.y_topSepCoord

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.geTopSep (Get)", eGDRaiseError_Raise

End Property

Public Property Get geBtmSep() As Long
On Error GoTo ErrSection:
    
    'this is the pane's bottom separator y value in pixels
    geBtmSep = m.gePaneStruct.y_btmSepCoord

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.geBtmSep (Get)", eGDRaiseError_Raise

End Property

Public Property Get geScaleType() As ePANE_ScaleType
On Error GoTo ErrSection:
    
    geScaleType = m.gePaneStruct.y_scaleType

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.geScaleType (Get)", eGDRaiseError_Raise

End Property

Public Property Get gePaneId() As Long
On Error GoTo ErrSection:
    
    gePaneId = m.gePaneStruct.paneId

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.gePaneId (Get)", eGDRaiseError_Raise

End Property

Public Property Let gePaneId(ByVal nID As Long)
On Error GoTo ErrSection:
    
    m.gePaneStruct.paneId = nID


ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.gePaneId (Let)", eGDRaiseError_Raise

End Property

Public Property Get gePaneMin() As Double
On Error GoTo ErrSection:
    
    gePaneMin = m.gePaneStruct.y_scaleMin

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.gePaneMin (Get)", eGDRaiseError_Raise

End Property

Public Property Let gePaneMin(ByVal dMin#)
On Error GoTo ErrSection:
    
    m.gePaneStruct.y_scaleMin = dMin

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.gePaneMin (Let)", eGDRaiseError_Raise

End Property

Public Property Get gePaneMax() As Double
On Error GoTo ErrSection:
    
    gePaneMax = m.gePaneStruct.y_scaleMax

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.gePaneMax (Get)", eGDRaiseError_Raise

End Property

Public Property Let gePaneMax(ByVal dMax#)
On Error GoTo ErrSection:
    
    m.gePaneStruct.y_scaleMax = dMax

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.gePaneMax (Let)", eGDRaiseError_Raise

End Property

Public Property Get gePaneShow() As Long
On Error GoTo ErrSection:
    
    gePaneShow = m.gePaneStruct.paneShow

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPane.gePaneShow (Get)", eGDRaiseError_Raise

End Property

Public Sub geSavePaneMax()
On Error GoTo ErrSection:
    
    m.dMaxSave = m.gePaneStruct.y_scaleMax

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.geSavePaneMax", eGDRaiseError_Raise

End Sub

Public Sub geSavePaneMin()
On Error GoTo ErrSection:

    m.dMinSave = m.gePaneStruct.y_scaleMin

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.geSavePaneMin", eGDRaiseError_Raise

End Sub

Public Sub geClearMaxRatio()
On Error GoTo ErrSection:
    
    m.gePaneStruct.y_spAboveRatio = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.geClearMaxRatio", eGDRaiseError_Raise

End Sub

Public Sub geClearMinRatio()
On Error GoTo ErrSection:
    
    m.gePaneStruct.y_spBelowRatio = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.geClearMinRatio", eGDRaiseError_Raise

End Sub

Public Sub geIncDecMaxRatio(ByVal dIncDecAmt#, Optional ByVal bFromMouse As Boolean = False)
On Error GoTo ErrSection:

    If m.gePaneStruct.y_scaleType = ePANE_ScaleLog Or m.gePaneStruct.y_scaleType = ePANE_ScaleLogDraw Then
        m.gePaneStruct.y_spAboveRatio = m.gePaneStruct.y_spAboveRatio + dIncDecAmt      '4586
    Else
        If bFromMouse And g.ChartGlobals.eDragModeY = eDragModeY_Both And m.gePaneStruct.y_scaleType <> ePANE_ScaleLog And _
           m.gePaneStruct.y_scaleType <> ePANE_ScaleLogDraw Then
            m.gePaneStruct.y_spAboveRatio = (m.gePaneStruct.y_spAboveRatio + 0.5) * (1 + dIncDecAmt) - 0.5
        Else
            m.gePaneStruct.y_spAboveRatio = m.gePaneStruct.y_spAboveRatio + dIncDecAmt
        End If
    End If
    
    If m.gePaneStruct.y_spAboveRatio < -0.45 Then
        m.gePaneStruct.y_spAboveRatio = -0.45
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.geIncDecMaxRatio", eGDRaiseError_Raise

End Sub

Public Sub geIncDecMinRatio(ByVal dIncDecAmt#, Optional ByVal bFromMouse As Boolean = False)
On Error GoTo ErrSection:
    
    If m.gePaneStruct.y_scaleType = ePANE_ScaleLog Or m.gePaneStruct.y_scaleType = ePANE_ScaleLogDraw Then
        m.gePaneStruct.y_spBelowRatio = m.gePaneStruct.y_spBelowRatio + dIncDecAmt      '4586
    Else
        If bFromMouse And g.ChartGlobals.eDragModeY = eDragModeY_Both And m.gePaneStruct.y_scaleType <> ePANE_ScaleLog And _
           m.gePaneStruct.y_scaleType <> ePANE_ScaleLogDraw Then
            m.gePaneStruct.y_spBelowRatio = -((-m.gePaneStruct.y_spBelowRatio + 0.5) * (1 - dIncDecAmt) - 0.5)
        Else
            m.gePaneStruct.y_spBelowRatio = m.gePaneStruct.y_spBelowRatio + dIncDecAmt
        End If
    End If

    If m.gePaneStruct.y_spBelowRatio > 0.45 Then
        m.gePaneStruct.y_spBelowRatio = 0.45
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.geIncDecMinRatio", eGDRaiseError_Raise

End Sub

Public Sub geSetMaxRatio(ByVal nButton&)
On Error GoTo ErrSection:
    
    'Developer's note:
    'This sub is intended for use ony by frmChart.frm for resizing the y-scale
    'with the hand cursor in the y-scale area. To set this ratio to a specific
    'value, call the geClearMaxRatio first then follow with geIncDecMaxRatio.
    
    Dim dDiff#
    
    If m.eScaleMode = ePANE_ScaleModeSquare Then Exit Sub     'square scale
    
    'manual scale
    If m.eScaleMode = ePANE_ScaleModeManual Then
        m.gePaneStruct.y_spAboveRatio = 0
        m.dMax = m.gePaneStruct.y_scaleMax
        If nButton = vbRightButton Then m.dMin = m.gePaneStruct.y_scaleMin
        Exit Sub
    End If
    
    'auto scale
    If m.eScaleMode = ePANE_ScaleModeAuto Or m.eScaleMode = ePANE_ScaleModeAutoPrice Then
        If nButton = vbLeftButton Then
            If m.dMaxSave - m.gePaneStruct.y_scaleMin <> 0 Then
                dDiff = m.gePaneStruct.y_scaleMax - m.dMaxSave
                m.gePaneStruct.y_spAboveRatio = dDiff / (m.dMaxSave - m.gePaneStruct.y_scaleMin)
                m.gePaneStruct.y_scaleMax = m.dMaxSave
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.geSetMaxRatio", eGDRaiseError_Raise

End Sub

Public Sub geSetMinRatio(ByVal nButton&)
On Error GoTo ErrSection:
    
    'Developer's note:
    'This sub is intended for use ony by frmChart.frm for resizing the y-scale
    'with the hand cursor in the y-scale area. To set this ratio to a specific
    'value, call the geClearMinRatio first then follow with geIncDecMinRatio.
    
    Dim dDiff#
    
    If m.eScaleMode = ePANE_ScaleModeSquare Then Exit Sub     'square scale
        
    'manual scale
    If m.eScaleMode = ePANE_ScaleModeManual Then
        m.gePaneStruct.y_spBelowRatio = 0
        m.dMin = m.gePaneStruct.y_scaleMin
        If nButton = vbRightButton Then m.dMax = m.gePaneStruct.y_scaleMax
        Exit Sub
    End If
    
    'auto scale
    If m.eScaleMode = ePANE_ScaleModeAuto Or m.eScaleMode = ePANE_ScaleModeAutoPrice Then
        If nButton = vbLeftButton Then
            If m.gePaneStruct.y_scaleMax - m.dMinSave <> 0 Then
                dDiff = m.gePaneStruct.y_scaleMin - m.dMinSave
                m.gePaneStruct.y_spBelowRatio = dDiff / (m.gePaneStruct.y_scaleMax - m.dMinSave)
                m.gePaneStruct.y_scaleMin = m.dMinSave
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.geSetMinRatio", eGDRaiseError_Raise

End Sub

Public Sub geResetData()
On Error GoTo ErrSection:

    m.geAdded = False
    m.bCanBeNonLinear = False
    m.bLogFlagChecked = False
    
    m.gePaneStruct.paneId = -1
    m.gePaneStruct.IsPricePane = 0
    m.gePaneStruct.indicatorCount = 0
    m.gePaneStruct.splitPaneShow = 0
    m.gePaneStruct.splitPaneWidth = 0
    gdSetSize m.gePaneStruct.gdshSplitPane, 0, 0
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.geResetData", eGDRaiseError_Raise

End Sub

Public Function geAddPane(Chart As cChart, ByVal chartObj&, ByVal nID&) As Long
On Error GoTo ErrSection:

    Dim rc&, dMove#, dPoints#
    Dim b As cGdBars
    
    ''If InStr(UCase(Chart.TemplateApplied), "WOODIES") Or InStr(UCase(Me.Name), "WOODIES") Then
''    If Chart.ShowSplitPane <> 0 Then
''        m.eSplitPaneType = ePANE_SplitPaneWood
''    Else
''        m.eSplitPaneType = ePANE_SplitPaneNone
''    End If
    
    m.gePaneStruct.y_base2Move = 0
    
    If m.idxPriceBars > 0 Then
        If Chart.Tree.NodeLevel(m.idxPriceBars) > 0 Then
            Set b = Chart.Tree(m.idxPriceBars).Bars
        ElseIf m.nDisplayFormat = ePANE_PriceFormat Or Chart.Tree.Key(nID) = "PRICE PANE" Then
            Set b = Chart.Bars
        Else
            Set b = Nothing
        End If
    ElseIf m.nDisplayFormat = ePANE_PriceFormat Or Chart.Tree.Key(nID) = "PRICE PANE" Then
        Set b = Chart.Bars
    Else
        Set b = Nothing
    End If
    
    If Not b Is Nothing Then
        dMove = b.Prop(eBARS_TickMove)
        If dMove = 0.25 Or dMove = 0.125 Or dMove = 1 / 16# Or _
            dMove = 1 / 32# Or dMove = 1 / 64# Or dMove = 1 / 128# Then
                m.gePaneStruct.y_base2Move = dMove
        End If
    End If
    
    m.gePaneStruct.paneId = nID
    m.gePaneStruct.htProportion = RoundNum(Size * 100)
        
    m.gePaneStruct.paneShow = 0
    If m.bDisplay = True Then m.gePaneStruct.paneShow = 1
    
    LogFlagOk Chart, True
    
    'm.eScaleMode:  0=auto, 1=manual, 2=square
    'y_scaleType: 0=normal, 1=log, 2=square, 3=scale to price, 4=log AND log mode draw, 99=do not use
    If m.eLogFlag = ePANE_LogFlagLog Then
        If g.ChartGlobals.bLogModeDraw Then
            m.gePaneStruct.y_scaleType = ePANE_ScaleLogDraw
        Else
            m.gePaneStruct.y_scaleType = ePANE_ScaleLog
        End If
    Else
        If m.eScaleMode = ePANE_ScaleModeSquare Then
            m.gePaneStruct.y_scaleType = ePANE_ScaleSquare
            GetCustomSquare Chart
            If m.dSquareTicks > 0 And m.nSquareBars > 0 Then
                dPoints = m.dSquareTicks * Chart.Bars.Prop(eBARS_TickMove) 'convert ticks to points
                m.gePaneStruct.ptsPerBar = dPoints / m.nSquareBars
            Else
                m.gePaneStruct.ptsPerBar = 1
            End If
        ElseIf m.eScaleMode = ePANE_ScaleModeAutoPrice Then
            m.gePaneStruct.y_scaleType = ePANE_ScalePrice
        Else
            m.gePaneStruct.y_scaleType = ePANE_ScaleNormal
        End If
    End If

    If m.gePaneStruct.y_scaleType = ePANE_DoNotUse Then
        m.gePaneStruct.y_scaleType = ePANE_ScaleNormal  'precautionary, should never happen
    End If
    
    m.gePaneStruct.paneSepHide = 0
    m.gePaneStruct.paneSepHide = m.nHideSeparator
    m.gePaneStruct.y_labelAll = m.nYscaleLabelAll
            
    If m.geAdded = False Then
        rc = geAddItem(chartObj, 1, m.gePaneStruct)
        If rc = 0 Then m.geAdded = True
    End If
    
    geAddPane = rc
            
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPane.geAddPane", eGDRaiseError_Raise

End Function

Private Sub GetCustomSquare(Chart As cChart)
On Error GoTo ErrSection:

    Dim strSym$, strPeriod$
    Dim strSearch$, strFound$, nPos&
    Dim gdStr As New cGdArray
    
    If g.ChartGlobals.aSquareSymList.Size < 1 Then Exit Sub
    
    
    strPeriod = Chart.Bars.Prop(eBARS_PeriodicityStr) + ","
    strSym = Chart.Symbol
    nPos = InStr(strSym, "-")
    
    If nPos > 0 And Left(strSym, 1) <> "$" Then
        strSym = Left(strSym, nPos - 1) + ","
    Else
        strSym = CStr(Chart.SymbolID) + ","
    End If
    strSearch = strSym + strPeriod
    
    g.ChartGlobals.aSquareSymList.BinarySearch strSearch, nPos
    strFound = g.ChartGlobals.aSquareSymList(nPos)
    If Left(strFound, Len(strSym)) = strSym Then
        gdStr.SplitFields strFound, ","
        m.dSquareTicks = Val(gdStr(2))
        m.nSquareBars = Val(gdStr(3))
    Else
        m.dSquareTicks = g.ChartGlobals.dSquareTicks
        m.nSquareBars = g.ChartGlobals.nSquareBars
    End If
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.GetCustomSquare", eGDRaiseError_Raise

End Sub

Private Sub SaveCustomSquare(Chart As cChart)
On Error GoTo ErrSection:

    Dim strSym$, strPeriod$
    Dim strSearch$, strFound$, strSave$
    Dim nPos&
    
    strSym = Chart.Symbol
    nPos = InStr(strSym, "-")
    
    If nPos > 0 And Left(strSym, 1) <> "$" Then
        strSym = Left(strSym, nPos - 1) + ","
    Else
        strSym = CStr(Chart.SymbolID) + ","
    End If
    
    strPeriod = Chart.Bars.Prop(eBARS_PeriodicityStr) + ","
    strSearch = strSym + strPeriod
    strSave = strSym + strPeriod + Str(m.dSquareTicks) + "," + Str(m.nSquareBars)
    
    If g.ChartGlobals.aSquareSymList.Size < 1 Then
        g.ChartGlobals.aSquareSymList.Add strSave
        Exit Sub
    End If

    g.ChartGlobals.aSquareSymList.BinarySearch strSearch, nPos
    strFound = g.ChartGlobals.aSquareSymList(nPos)
    If Left(strFound, Len(strSym)) = strSym Then
        g.ChartGlobals.aSquareSymList(nPos) = strSave
    Else
        g.ChartGlobals.aSquareSymList.Add strSave
        g.ChartGlobals.aSquareSymList.Sort eGdSort_IgnoreCase Or eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPane.SaveCustomSquare", eGDRaiseError_Raise

End Sub

Public Sub CenterData(Chart As cChart)

    Dim i&, nFirstBar&, nLastBar&
    Dim hXbarArray&, hXdateArray&
    Dim dMin#, dMax#, dNewMin#, dNewMax#
    Dim dR1#, dR2#, dDiff#
    Dim Ind As cIndicator
    
    If m.eScaleMode <> ePANE_ScaleModeManual Then Exit Sub  'precautionary, should never happen
    
    dNewMin = kNullData
    dNewMax = kNullData
    
    nFirstBar = Chart.FirstBar
    nLastBar = Chart.LastBar
    
    If nFirstBar <= 0 Or nLastBar <= 0 Or nFirstBar = nLastBar Then Exit Sub 'precautionary
    
    If m.eScaleAutoLastUsed = ePANE_ScaleModeAutoPrice Then
        Set Ind = Chart.Tree("PRICE")           'aardvark 4582
        If Not Ind Is Nothing Then
            dMin = Ind.IndMin(nFirstBar, nLastBar)
            dMax = Ind.IndMax(nFirstBar, nLastBar)
        End If
    Else
        'walk through nodes of indicators belonging to this pane
        'and extract the min and max values of the visible indicators
        For i = m.gePaneStruct.paneId + 1 To Chart.Tree.Count
            If Chart.Tree.NodeLevel(i) = 0 Then Exit For    'have hit next pane
            Set Ind = Chart.Tree(i)
            If Not Ind Is Nothing Then
                If Ind.Display = True Then
                    If Ind.DisplayType <> eINDIC_ArtPyramid And _
                       Ind.DisplayType <> eINDIC_ArtReversal And _
                       Ind.DataType <> eINDIC_BooleanArray Then
                       
                        dMin = Ind.IndMin(nFirstBar, nLastBar)
                        dMax = Ind.IndMax(nFirstBar, nLastBar)
                        If dMin <> kNullData Then
                            If dNewMin = kNullData Or dNewMin > dMin Then dNewMin = dMin
                        End If
                        If dMax <> kNullData Then
                            If dNewMax = kNullData Or dNewMax < dMax Then dNewMax = dMax
                        End If
                        
                    End If
                End If
            End If
        Next
    End If
    
    If dNewMin = kNullData Or dNewMax = kNullData Then Exit Sub
        
    dR1 = m.dMax - m.dMin
    dR2 = dNewMax - dNewMin
    dDiff = dR2 - dR1
            
    'adjust min and max to make range the same
    If dDiff > 0 Then
        dNewMin = dNewMin + dDiff / 2
        dNewMax = dNewMax - dDiff / 2
    Else
        dNewMin = dNewMin - Abs(dDiff) / 2
        dNewMax = dNewMax + Abs(dDiff) / 2
    End If
        
    m.dMin = dNewMin
    m.gePaneStruct.y_scaleMin = dNewMin

    m.dMax = dNewMax
    m.gePaneStruct.y_scaleMax = dNewMax

    Chart.GenerateChart eRedo1_Scrolled

End Sub

Public Property Get YscaleLabelAll() As Long
    YscaleLabelAll = m.nYscaleLabelAll
End Property

Public Property Let YscaleLabelAll(ByVal nAll&)
    m.nYscaleLabelAll = nAll
    m.gePaneStruct.y_labelAll = nAll
End Property

Private Sub Class_Terminate()
        
    Set m.oWoodData = Nothing
    gdDestroyArray m.gePaneStruct.gdshSplitPane
    
End Sub

Public Sub SetSplitPaneCluster(Chart As cChart, Ind As cIndicator)
On Error GoTo ErrSection:

    Dim iMin&, iMax&, iSize&
    Dim hPrice&, hClusterValue&

    If Chart Is Nothing Or Ind Is Nothing Then GoTo ErrExit
    If m.gePaneStruct.IsPricePane <> 1 Then GoTo ErrExit
    
    If Not Ind.Display Then GoTo ErrExitClear
    
    hPrice = Ind.geDataHandle(2)
    hClusterValue = Ind.geDataHandle(3)
    
    If hPrice = 0 Or hClusterValue = 0 Then GoTo ErrExitClear
    
    iSize = gdGetSize(hClusterValue)
    If iSize <= 0 Then GoTo ErrExitClear
    
    If iSize <> gdGetSize(hPrice) Then GoTo ErrExitClear
        
    Set m.oWoodData = Nothing
    If m.eSplitPaneType = ePANE_SplitPaneTimer Then m.nRestoreTimerPane = 1
    m.eSplitPaneType = ePANE_SplitPaneCluster
    m.gePaneStruct.splitPaneShow = 1
        
    iMin = gdMinValue(hClusterValue, 0, iSize - 1)
    iMax = gdMaxValue(hClusterValue, 0, iSize - 1)
    'clear array in case it was previously used for timer or Woodies CCI info
    gdClear m.gePaneStruct.gdshSplitPane, 1
    'format for string: PP_CLUSTER|min cluster value|max cluster value|optional fixed pane width
    gdSetStr m.gePaneStruct.gdshSplitPane, 0, "PP_CLUSTER|" & Str(iMin) & "|" & Str(iMax) & "|0"

ErrExit:
    Exit Sub

ErrExitClear:
    Set m.oWoodData = Nothing
    m.eSplitPaneType = ePANE_SplitPaneNone
    m.gePaneStruct.splitPaneShow = 0
    gdClear m.gePaneStruct.gdshSplitPane, 1
    Exit Sub

ErrSection:
    RaiseError "cPane.SetSplitPaneCluster"

End Sub

Public Sub SetSplitPaneData(Chart As cChart, ByVal bHasWoodPane As Boolean)
On Error GoTo ErrSection:

    Dim Pane As cPane
    Dim Ind As cIndicator
    
    Dim strKey$, strDelta$
    Dim i&, iBar&, iNum&, iNum2&
    Dim dPrice#, dPrevPrice#, dDelta#, dPercent#
    
    Dim Bars As cGdBars


    If Chart Is Nothing Then Exit Sub
    
    'JM 01-15-2011
    'timer & cluster pane are mutually exclusive
    'design intentions: If timer is on when user adds cluster then replace timer with price
    '   cluster in split pane area and restore the timer display in split pane area when
    '   user removes fib cluster.
    If m.nRestoreTimerPane = 1 Then
        If Chart.Tree(kClusterTimeKeyInd) Is Nothing And Chart.Tree(kClusterPriceKey) Is Nothing Then
            m.eSplitPaneType = ePANE_SplitPaneTimer
            m.nRestoreTimerPane = 0
        Else
            Exit Sub
        End If
    ElseIf m.eSplitPaneType = ePANE_SplitPaneCluster Then
        If Chart.Tree(kClusterTimeKeyInd) Is Nothing And Chart.Tree(kClusterPriceKey) Is Nothing Then
            If m.nRestoreTimerPane Then
                m.eSplitPaneType = ePANE_SplitPaneTimer
            Else
                m.eSplitPaneType = ePANE_SplitPaneNone
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    
    Set Bars = Chart.Bars
    
    If Chart.ShowSplitPane <> 1 Then
        m.gePaneStruct.splitPaneShow = 0
        Chart.RemoveAnnots True, eANNOT_Icon, eANNOT_SideWinderLabel
        Exit Sub
    End If
    
    If bHasWoodPane Then
        If m.gePaneStruct.IsPricePane Or UCase(m.strName) = "WOODIES CCI" Then
            m.gePaneStruct.splitPaneShow = 1
            m.eSplitPaneType = ePANE_SplitPaneWood
            If m.oWoodData Is Nothing Then Set m.oWoodData = New cPaneWood
            m.oWoodData.SetDrawData Chart, Me
        Else
            m.gePaneStruct.splitPaneShow = 0
            m.eSplitPaneType = ePANE_SplitPaneNone
        End If
    ElseIf m.gePaneStruct.IsPricePane Then
        m.gePaneStruct.splitPaneShow = 1
        
        If m.eSplitPaneType = ePANE_SplitPaneNone Then m.eSplitPaneType = ePANE_SplitPaneTimer
        If m.oWoodData Is Nothing Then Set m.oWoodData = New cPaneWood
        
        If m.eSplitPaneType = ePANE_SplitPaneTimer Then
            Set Ind = Chart.Tree("PRICE")
            If Ind Is Nothing Then
                m.oWoodData.DeltaLabel = ""
            Else
                strDelta = Parse(Ind.ChartValue, "(", 2)
                If Len(strDelta) > 0 Then strDelta = Left(strDelta, Len(strDelta) - 1) 'strip off ")"
                m.oWoodData.DeltaLabel = strDelta
                
                If Bars Is Nothing Then
                    m.oWoodData.PercentLabel = ""
                Else
                    iBar = Chart.LastGoodDataBar(False, True)
                    dPrice = Bars(eBARS_Close, iBar)
                    
                    dDelta = Bars.PriceFromString(strDelta)         '6448
                    dPrevPrice = dPrice - dDelta
                    
                    If dPrevPrice = 0 Or dPrevPrice * dPrice <= 0 Then
                        m.oWoodData.PercentLabel = ""
                    Else
                        dPercent = dPrice / dPrevPrice - 1
                        m.oWoodData.PercentLabel = Format(dPercent, "+#0.00%;-#0.00%")
                    End If
                End If
                
            End If
        End If
        
        m.oWoodData.SetDrawData Chart, Me
    Else
        m.gePaneStruct.splitPaneShow = 0
        m.eSplitPaneType = ePANE_SplitPaneNone
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cPane.SetSplitPaneData"
    
End Sub

Public Property Get SplitPaneStrArray() As Long
    SplitPaneStrArray = m.gePaneStruct.gdshSplitPane
End Property

Public Property Get WoodPane() As cPaneWood
    Set WoodPane = m.oWoodData
End Property

Public Property Get SplitPaneType() As ePANE_SplitPaneType
    SplitPaneType = m.eSplitPaneType
End Property

Public Property Let SplitPaneType(ByVal eType As ePANE_SplitPaneType)
    
    If m.eSplitPaneType <> eType Then Set m.oWoodData = Nothing       'there's been a type change; remove object
    m.eSplitPaneType = eType

End Property

Public Property Get SplitPaneLabelIdxCCI(ByVal nIdx&) As Long
    
    Dim i&
    
    i = -1
    
    If nIdx >= 0 And nIdx < gdGetSize(m.gePaneStruct.gdshSplitPane) Then
        i = ValOfText(Parse(gdGetStr(m.gePaneStruct.gdshSplitPane, nIdx), "|", 10))
    End If
    
    SplitPaneLabelIdxCCI = i
    
End Property

Public Sub CheckDinapoliStudies(PaneMacD As cPane, PaneDetrendOsc As cPane, PanePrefStoch As cPane, aIndicators As cGdArray)
On Error GoTo ErrSection:

    Dim i&
    Dim strUpperName As String
    Dim bFound As Boolean
    Dim Ind As cIndicator
    
    strUpperName = UCase(Me.Name)
    
    If InStr(strUpperName, "CTP") Then
        If InStr(strUpperName, "MACD") Then
            Set PaneMacD = Me
            bFound = True
        ElseIf InStr(strUpperName, "DETREND") Then
            If InStr(strUpperName, "OSC") Then
                Set PaneDetrendOsc = Me
                bFound = True
            End If
        ElseIf InStr(strUpperName, "PREF") Then
            If InStr(strUpperName, "STOCH") Then
                Set PanePrefStoch = Me
                bFound = True
            End If
        End If
    End If
    
    If bFound Then Exit Sub
    If aIndicators Is Nothing Then Exit Sub
    If aIndicators.Size = 0 Then Exit Sub
    
    Dim bIndDetrendOsc As Boolean
    Dim bIndHorz As Boolean
    Dim bIndPrefStochK As Boolean
    Dim bIndPrefStochD As Boolean
    Dim bIndOverBought As Boolean
    Dim bIndOverSold As Boolean
    Dim bIndMacd As Boolean
    Dim bIndMacdXma As Boolean
    Dim bIndMacdDiff As Boolean
    
    For i = 0 To aIndicators.Size - 1
        Set Ind = aIndicators(i)
        If Not Ind Is Nothing Then
            strUpperName = UCase(Ind.Name)
            If InStr(strUpperName, "HORIZONTAL") Then
                If InStr(strUpperName, "LINE") Then bIndHorz = True
            ElseIf InStr(strUpperName, "OVERBOUGHT") Then
                bIndOverBought = True
            ElseIf InStr(strUpperName, "OVERSOLD") Then
                bIndOverSold = True
            ElseIf InStr(strUpperName, "DETREND") Then
                If InStr(strUpperName, "OSC") Then bIndDetrendOsc = True
            ElseIf InStr(strUpperName, "DINAPOLI") Then
                If InStr(strUpperName, "PREF") Then
                    If InStr(strUpperName, "STOCHK") Then
                        bIndPrefStochK = True
                    ElseIf InStr(strUpperName, "STOCHD") Then
                        bIndPrefStochD = True
                    End If
                ElseIf InStr(strUpperName, "MACD") Then
                    If InStr(strUpperName, "XMA") Then
                        bIndMacdXma = True
                    ElseIf InStr(strUpperName, "DIFF") Then
                        bIndMacdDiff = True
                    Else
                        bIndMacd = True
                    End If
                End If
            End If
        End If
        If bIndDetrendOsc And bIndHorz Then
            Set PaneDetrendOsc = Me
            Exit For
        ElseIf bIndMacd And bIndMacdXma And bIndMacdDiff Then
            Set PaneMacD = Me
            Exit For
        ElseIf bIndPrefStochK And bIndPrefStochD And bIndOverBought And bIndOverSold Then
            Set PanePrefStoch = Me
            Exit For
        End If
    Next
    
    Exit Sub
        
ErrSection:
    RaiseError "cPane.CheckDinapoliStudies"
    
End Sub

Public Property Get geAdjustMin() As Double
    geAdjustMin = m.gePaneStruct.y_adjMin
End Property

Public Property Get geAdjustMax() As Double
    geAdjustMax = m.gePaneStruct.y_adjMax
End Property

Public Sub OptNavGraphInfoSet(Chart As cChart, aPrice As cGdArray, aNow As cGdArray, _
    aThird As cGdArray, aTwoThird As cGdArray, aExpire As cGdArray, _
    aProb1 As cGdArray, aProb2 As cGdArray, ByVal strGraph$)
On Error GoTo ErrSection:
    
    If m.eSplitPaneType <> ePANE_SplitPaneNone And m.eSplitPaneType <> ePANE_SplitPaneOptGraph Then Set m.oWoodData = Nothing
    
    If m.oWoodData Is Nothing Then
        Set m.oWoodData = New cPaneWood
        Chart.RedoMode = eRedo3_Settings
    End If
        
    m.eSplitPaneType = ePANE_SplitPaneOptGraph
    m.oWoodData.OptNavGraphInfoSet strGraph, aPrice, aNow, aThird, aTwoThird, aExpire, aProb1, aProb2
    m.oWoodData.SetDrawData Chart, Me
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPane.OptNavGraphInfoSet"

End Sub

Public Sub OptNavGraphInfoClear()
On Error GoTo ErrSection:

    If m.eSplitPaneType = ePANE_SplitPaneOptGraph Then
        m.eSplitPaneType = ePANE_SplitPaneNone
        gdSetSize m.gePaneStruct.gdshSplitPane, 0, 0
        m.gePaneStruct.splitPaneShow = 0
        Set m.oWoodData = Nothing
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPane.OptNavGraphInfoClear"

End Sub

Public Sub OptNavGraphValues(ByVal dPrice#, dNow#, dThird#, dTwoThird#, dExpire#)
    If Not m.oWoodData Is Nothing Then m.oWoodData.OptNavGraphValues dPrice, dNow, dThird, dTwoThird, dExpire
End Sub

Public Sub OptNavGraphDate(dDateNow#, dDateThird#, dDateTwoThird#, dDateExpire#)
On Error GoTo ErrSection:

    If Not m.oWoodData Is Nothing Then m.oWoodData.OptNavGraphDate dDateNow, dDateThird, dDateTwoThird, dDateExpire

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPane.OptNavGraphDate"
    
End Sub

Public Property Get CanBeNonLinear(Chart As cChart) As Boolean
On Error GoTo ErrSection
    
    If Not m.bLogFlagChecked Then
        If Not Chart Is Nothing Then
            LogFlagOk Chart, False
        End If
    End If
    
    CanBeNonLinear = m.bCanBeNonLinear

ErrExit:
    Exit Property

ErrSection:
    CanBeNonLinear = False
    RaiseError "cPane.CanBeNonLinear.Get"

End Property

Public Sub LogFlagReset()
On Error GoTo ErrSection:

    m.bLogFlagChecked = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPane.LogFlagReset"

End Sub

Private Sub LogFlagOk(Chart As cChart, ByVal bFixFlag As Boolean)
On Error GoTo ErrSection:

'this routine:
' - sets the private data member bCanBeNonLinear to True/False as appropriate
' - fixes the private data member eLogFlag if fix is requested and fixing is needed

    Dim idx As Long
    Dim bOkay As Boolean
    
    Dim Tree As cGdTree
    Dim Ind As cIndicator

    If m.bLogFlagChecked Then
        If Not m.bCanBeNonLinear And bFixFlag Then m.eLogFlag = ePANE_LogFlagLinear
        GoTo ErrExit
    End If
    
    If m.eScaleMode = ePANE_ScaleModeSquare Then
        If bFixFlag Then m.eLogFlag = ePANE_LogFlagLinear
        m.bCanBeNonLinear = False
        GoTo ErrExit
    End If
    
    If Not m.bDisplay Then
        m.bCanBeNonLinear = True
        GoTo ErrExit
    End If
    
    If Chart Is Nothing Then GoTo ErrExit
    Set Tree = Chart.Tree
    If Tree Is Nothing Then GoTo ErrExit
    
    If Tree(m.gePaneStruct.paneId) Is Nothing Then GoTo ErrExit
    If Tree.NodeLevel(m.gePaneStruct.paneId) > 0 Then GoTo ErrExit      'precautionary
    If Not Tree.HasChildren(m.gePaneStruct.paneId) Then GoTo ErrExit
    
    idx = Tree.RelativeIndex(m.gePaneStruct.paneId, eTREE_FirstChild)
    If Tree(idx) Is Nothing Then GoTo ErrExit
    If Tree.NodeLevel(idx) <= 0 Then GoTo ErrExit
    
    bOkay = True
    Set Ind = Tree(idx)
    
    While bOkay And Not Ind Is Nothing
        
        bOkay = Ind.CanBeNonLinear()
        
        Set Ind = Nothing
        If bOkay Then
            idx = Tree.RelativeIndex(idx, eTREE_NextSibling)
            If Not Tree(idx) Is Nothing Then
                If Tree.NodeLevel(idx) > 0 Then
                    Set Ind = Tree(idx)
                End If
            End If
        End If
    
    Wend
    
    m.bLogFlagChecked = True
    m.bCanBeNonLinear = bOkay
    
    If Not bOkay And bFixFlag Then m.eLogFlag = ePANE_LogFlagLinear

ErrExit:
    Exit Sub

ErrSection:
    m.bCanBeNonLinear = False
    RaiseError "cPane.LogFlagOk"

End Sub

Public Sub SplitPaneClear()
On Error GoTo ErrSection:

    m.eSplitPaneType = ePANE_SplitPaneNone
    m.gePaneStruct.splitPaneShow = 0
    m.gePaneStruct.splitPaneWidth = 0

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPane.SplitPaneClear"

End Sub

Public Property Get SeasonalPaneFlag() As Long
On Error GoTo ErrSection:


    SeasonalPaneFlag = m.nPaneSeasonal

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cPane.Get.SeasonalPaneFlag"

End Property

Public Property Let SeasonalPaneFlag(ByVal nFlag&)
On Error GoTo ErrSection:


    m.nPaneSeasonal = nFlag

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cPane.Let.SeasonalPaneFlag"


End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JM 10-17-2013:
'These functions are meant to be used by chart copy to maintain temporary y-scale values
'as requested by Steven Craig from EWI. Original design is not to save temporary y-scale.
Public Property Get geSpaceAbove() As Double
On Error GoTo ErrSection:
    
    geSpaceAbove = m.gePaneStruct.y_spAboveRatio

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cPane.Get.geSpaceAbove"

End Property

Public Property Get geSpaceBelow() As Double
On Error GoTo ErrSection:
    
    geSpaceBelow = m.gePaneStruct.y_spBelowRatio

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cPane.Get.geSpaceBelow"

End Property

Public Sub geCopyScaleStructInfo(fromPane As cPane)
On Error GoTo ErrExit

    m.gePaneStruct.y_adjMax = fromPane.geAdjustMax
    m.gePaneStruct.y_adjMin = fromPane.geAdjustMin
    m.gePaneStruct.y_scaleMax = fromPane.gePaneMax
    m.gePaneStruct.y_scaleMin = fromPane.gePaneMin
    m.gePaneStruct.y_spAboveRatio = fromPane.geSpaceAbove
    m.gePaneStruct.y_spBelowRatio = fromPane.geSpaceBelow

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPane.geCopyScaleStructInfo"

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

