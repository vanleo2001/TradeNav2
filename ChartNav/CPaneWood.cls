VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPaneWood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DefaultFont = "Times New Roman"
Private Const DefaultFontSize = 10
Private Const kMaroon = 128

'format for string that graphics engine uses to draw
'yLocation: T=top of pane, B=bottom of pane, numeric=at yValue (price scale)
'xLocation: optional (ignored by grapheng.dll if len()=0; everything after this also ignored by grapheng.dll)
'<CCIIndex>: not used by grapheng.dll, included for hittest reference on VB side
'   text|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|yLocation|xLocation|<CCIIndex>

'format for saving text specs to pane section of chart template file
'   strings delimited by | and ~
'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|location|show

'Note that the actual value, i.e. text, sent to graphics engine for drawing should not
'   need to be saved to template file. A descriptive text type, i.e. textType, which
'   is not sent to the graphics engine, is saved to the template file for reference
'   and can be used for debugging display issues relating to font size, location, etc.
'   The text type is a saved string as described below.

'SplitPaneWoodPrice (text type)
'   PP_Price, PP_Timer, PP_Cluster, PP_PriceDelta, PP_PricePercent, PP_OptGraph

'SplitPaneWoodCCI (text type & corresponding array index)
'   text-type(string)       array index(numeric)         description
'   DayHigh                 DayHigh                     'high of day
'   PivotAbove              PivotAbove                  'closest pivot point above price (eg R1 +9)
'   BarHigh                 BarHigh                     'prev bar high [space] curr bar high
'   Price                   Price                       'current price
'   BarLow                  BarLow                      'prev bar low [space] curr bar low
'   EMA                     EMA                         'EMA ticks away from price, and EMA angle
'   Timer                   Timer                       'count down to next bar in seconds (not doing ticks countdown)
'   PivotBelow              PivotBelow                  'closest pivot point below price (eg S1 -3)
'   DayLow                  DayLow                      'low of day
'   SideWinder              SideWinder                  'sidewinder value in upper left corner under indicator labels
Public Enum eCCIText_Index
    eCCI_DayHigh = 0
    eCCI_PivotAbove
    eCCI_BarHigh
    eCCI_Price
    eCCI_BarLow
    eCCI_EMA
    eCCI_Timer
    eCCI_PivotBelow
    eCCI_DayLow
    eCCI_SideWinder
End Enum

Private Type eWOOD_IndData
    dIndR1 As Double
    dIndR2 As Double
    dIndR3 As Double
    dIndPP As Double
    dIndS1 As Double
    dIndS2 As Double
    dIndS3 As Double
    dIndEMA As Double                'this is MovingAvgX(close, 34)
    dIndEMAAngle As Double
    dIndSideWinder As Double
    nIndSideWinderID As Long
End Type

'when to highlight text
Private Type eWOOD_Highlight
    nColor As Long                  'color to highlight background of text with
    nDayHigh As Long                'highlight day high when it is this #of ticks within price
    nDayLow As Long                 'highlight day low when it is this #of ticks within price
    nUpperPivot As Long             'highlight closest pivot when it is this #of ticks within price
    nLowerPivot As Long             'highlight 2nd closest pivot when it is this #of ticks within price
    dTimerSeconds As Double         'highlight seconds timer when there is this % of current bar left (eg 0.15)
    dTimerTicks As Double           'highlight tick count timer when there is this % of current bar left (eg 0.7)
End Type

Private Type mPrivate
    eWoodData As eWOOD_IndData                  'holds indicators values
    eWoodHighlight As eWOOD_Highlight           'holds alert levels for when to highlight text
    
    aTemplateStr As New cGdArray
    
    strCountDownBkColor As String               'bk color for countdown
    strUpperPivotBkColor As String              'bk color for pivots closest to price
    strLowerPivotBkColor As String
    strSoundFile As String
        
    eDisplayFormat As ePANE_DisplayFormat
    nDisplayDecimals As Long
    nCounterToNextBar As Long
    nPlaySound As Long              '-1=play till end of bar, 0=none, 1=play once per bar
    bSoundDone As Long              'flag to determine whether sound already played for once per bar option
    
    dTickMove As Double
    dFlat As Double
    dTrending As Double
    
    strDeltaLabel As String
    strPercentLabel As String
    bPriceDeltaChecked As Boolean
    
    strOptNavGraphType As String    'type of OptNav risk graph
    dDateNow As Double
    dDateThird As Double
    dDateTwoThird As Double
    dDateExpire As Double
    aPrice As cGdArray
    aNow As cGdArray
    aThird As cGdArray
    aTwoThird As cGdArray
    aExpire As cGdArray
    aProb1 As cGdArray
    aProb2 As cGdArray
End Type

Private m As mPrivate

Public Sub Reset()
    m.eWoodData.dIndEMA = 0
    m.eWoodData.dIndEMAAngle = 0
    m.eWoodData.dIndPP = 0
    m.eWoodData.dIndR1 = 0
    m.eWoodData.dIndR2 = 0
    m.eWoodData.dIndR3 = 0
    m.eWoodData.dIndS1 = 0
    m.eWoodData.dIndS2 = 0
    m.eWoodData.dIndS3 = 0
    m.eWoodData.dIndSideWinder = 0
    m.eWoodData.nIndSideWinderID = 0
    
    m.eWoodHighlight.dTimerSeconds = 0
    m.eWoodHighlight.dTimerTicks = 0
    m.eWoodHighlight.nColor = vbYellow
    m.eWoodHighlight.nDayHigh = 0
    m.eWoodHighlight.nDayLow = 0
    m.eWoodHighlight.nLowerPivot = 0
    m.eWoodHighlight.nUpperPivot = 0
    
    m.aTemplateStr.Size = 0
    
    m.strCountDownBkColor = ""
    m.strUpperPivotBkColor = ""
    m.strLowerPivotBkColor = ""
    m.strSoundFile = ""
    
    m.eDisplayFormat = ePANE_AutoFormat
    m.nDisplayDecimals = 0
    m.nCounterToNextBar = 0
    m.nPlaySound = 0
    m.bSoundDone = False
    
    m.dTickMove = 0
    m.dFlat = 0
    m.dTrending = 0
    
    m.strDeltaLabel = ""
    m.strPercentLabel = ""
    m.bPriceDeltaChecked = False
End Sub

Private Sub SetWoodIndData(Chart As cChart, ByVal dMinMove#)
On Error GoTo ErrSection:

    Dim i&, nFound&
    Dim strName$, dData#
        
    Dim Tree As cGdTree
    Dim Ind As cIndicator
    
    If Chart Is Nothing Then Exit Sub
    
    Set Tree = Chart.Tree
    If Tree Is Nothing Then Exit Sub
        
    'reset to null values
    m.eWoodData.dIndR1 = kNullData
    m.eWoodData.dIndR2 = kNullData
    m.eWoodData.dIndR3 = kNullData
    m.eWoodData.dIndPP = kNullData
    m.eWoodData.dIndS1 = kNullData
    m.eWoodData.dIndS2 = kNullData
    m.eWoodData.dIndS3 = kNullData
    m.eWoodData.dIndEMA = kNullData
    m.eWoodData.dIndEMAAngle = kNullData
    m.eWoodData.dIndSideWinder = kNullData
            
    For i = 1 To Tree.Count
        If Tree.NodeLevel(i) > 0 Then
            Set Ind = Tree(i)
            If Not Ind Is Nothing Then
                If Ind.IsWoodInd(True, True, strName, dData) Then
                    strName = UCase(strName)
                    If InStr(strName, "WOODIESPIVOT") <> 0 Or InStr(strName, "WOODIES PIVOT") <> 0 Then
                        If m.eWoodData.dIndPP = kNullData Then
                            m.eWoodData.dIndPP = dData
                            nFound = nFound + 1
                        End If
                    ElseIf InStr(strName, "WOODIESR1") <> 0 Or InStr(strName, "WOODIES R1") <> 0 Then
                        If m.eWoodData.dIndR1 = kNullData Then
                            m.eWoodData.dIndR1 = dData
                            nFound = nFound + 1
                        End If
                    ElseIf InStr(strName, "WOODIESR2") <> 0 Or InStr(strName, "WOODIES R2") <> 0 Then
                        If m.eWoodData.dIndR2 = kNullData Then
                            m.eWoodData.dIndR2 = dData
                            nFound = nFound + 1
                        End If
                    ElseIf InStr(strName, "WOODIESR3") <> 0 Or InStr(strName, "WOODIES R3") <> 0 Then
                        If m.eWoodData.dIndR3 = kNullData Then
                            m.eWoodData.dIndR3 = dData
                            nFound = nFound + 1
                        End If
                    ElseIf InStr(strName, "WOODIESS1") <> 0 Or InStr(strName, "WOODIES S1") <> 0 Then
                        If m.eWoodData.dIndS1 = kNullData Then
                            m.eWoodData.dIndS1 = dData
                            nFound = nFound + 1
                        End If
                    ElseIf InStr(strName, "WOODIESS2") <> 0 Or InStr(strName, "WOODIES S2") <> 0 Then
                        If m.eWoodData.dIndS2 = kNullData Then
                            m.eWoodData.dIndS2 = dData
                            nFound = nFound + 1
                        End If
                    ElseIf InStr(strName, "WOODIESS3") <> 0 Or InStr(strName, "WOODIES S3") <> 0 Then
                        If m.eWoodData.dIndS3 = kNullData Then
                            m.eWoodData.dIndS3 = dData
                            nFound = nFound + 1
                        End If
                    ElseIf InStr(strName, "MOVINGAVGX") <> 0 Then
                        If m.eWoodData.dIndEMA = kNullData Then
                            m.eWoodData.dIndEMA = dData
                            nFound = nFound + 1
                        End If
                    ElseIf InStr(strName, "EMAANGLE") <> 0 Or InStr(strName, "EMA ANGLE") <> 0 Then
                        If m.eWoodData.dIndEMAAngle = kNullData Then
                            m.eWoodData.dIndEMAAngle = dData
                            nFound = nFound + 1
                        End If
                    ElseIf InStr(strName, "SIDEWINDER") <> 0 Then
                        If m.eWoodData.dIndSideWinder = kNullData Then
                            m.eWoodData.dIndSideWinder = dData
                            m.eWoodData.nIndSideWinderID = Ind.geIndId
                            nFound = nFound + 1
                        End If
                    End If
                End If
            End If
        End If
        If nFound >= 10 Then Exit For
    Next
    
    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.SetWoodInData"
    
End Sub

Private Sub ClosestPivots(Bars As cGdBars, strAbovePrice$, strBelowPrice$, ByVal dPrice#)
On Error GoTo ErrSection:
    
    Dim tbTemp As New cGdTable
    Dim aIdx As cGdArray
    Dim i&, iAbove&, iBelow&
    Dim dPivot#, dPrice2#, dDiff#
    
    tbTemp.CreateField eGDARRAY_Doubles, 0
    tbTemp.CreateField eGDARRAY_Strings, 1
    
    tbTemp.AddRecord ""
    tbTemp(0, 0) = dPrice
    tbTemp(1, 0) = "Price"
    
    dPrice2 = ValOfText(FormatData(Bars, dPrice))

    If m.eWoodData.dIndPP <> kNullData Then
        tbTemp.AddRecord ""
        i = tbTemp.NumRecords - 1
        tbTemp(0, i) = m.eWoodData.dIndPP
        tbTemp(1, i) = "PP"
    End If
    
    If m.eWoodData.dIndR1 <> kNullData Then
        tbTemp.AddRecord ""
        i = tbTemp.NumRecords - 1
        tbTemp(0, i) = m.eWoodData.dIndR1
        tbTemp(1, i) = "R1"
    End If
    
    If m.eWoodData.dIndR2 <> kNullData Then
        tbTemp.AddRecord ""
        i = tbTemp.NumRecords - 1
        tbTemp(0, i) = m.eWoodData.dIndR2
        tbTemp(1, i) = "R2"
    End If

    If m.eWoodData.dIndR3 <> kNullData Then
        tbTemp.AddRecord ""
        i = tbTemp.NumRecords - 1
        tbTemp(0, i) = m.eWoodData.dIndR3
        tbTemp(1, i) = "R3"
    End If

    If m.eWoodData.dIndS1 <> kNullData Then
        tbTemp.AddRecord ""
        i = tbTemp.NumRecords - 1
        tbTemp(0, i) = m.eWoodData.dIndS1
        tbTemp(1, i) = "S1"
    End If

    If m.eWoodData.dIndS2 <> kNullData Then
        tbTemp.AddRecord ""
        i = tbTemp.NumRecords - 1
        tbTemp(0, i) = m.eWoodData.dIndS2
        tbTemp(1, i) = "S2"
    End If

    If m.eWoodData.dIndS3 <> kNullData Then
        tbTemp.AddRecord ""
        i = tbTemp.NumRecords - 1
        tbTemp(0, i) = m.eWoodData.dIndS3
        tbTemp(1, i) = "S3"
    End If
    
    If tbTemp.NumRecords > 1 Then
        Set aIdx = tbTemp.CreateSortedIndex(0)
        'go through table look for price
        For i = 0 To aIdx.Size - 1
            If tbTemp(1, aIdx(i)) = "Price" Then Exit For
        Next
        
        If i = 0 Then
            'price is above all pivot points
            iAbove = 1
            iBelow = 2
        ElseIf i = aIdx.Size - 1 Then
            'price is below all pivot points
            iAbove = i - 1
            iBelow = i - 2
        Else
            iAbove = i + 1
            iBelow = i - 1
        End If
        
        'closest pivot to price (shown in upper half of pane)
        dPivot = tbTemp(0, aIdx(iAbove))
        dPivot = ValOfText(FormatData(Bars, dPivot))
        
        dDiff = (dPivot - dPrice2) / m.dTickMove
        dDiff = RoundNum(dDiff)
        If dDiff >= 0 Then
            strAbovePrice = tbTemp(1, aIdx(iAbove)) & " +" & Str(dDiff)
        Else
            strAbovePrice = tbTemp(1, aIdx(iAbove)) & " " & Str(dDiff)
        End If
        If (Abs(dDiff) <= m.eWoodHighlight.nUpperPivot) Then
            m.strUpperPivotBkColor = Str(m.eWoodHighlight.nColor)
        Else
            m.strUpperPivotBkColor = "-1"
        End If
        
        '2nd closest pivot to price (shown in lower half of pane)
        dPivot = tbTemp(0, aIdx(iBelow))
        dPivot = ValOfText(FormatData(Bars, dPivot))
        
        dDiff = (dPivot - dPrice2) / m.dTickMove
        dDiff = RoundNum(dDiff)
        If dDiff >= 0 Then
            strBelowPrice = tbTemp(1, aIdx(iBelow)) & " +" & Str(dDiff)
        Else
            strBelowPrice = tbTemp(1, aIdx(iBelow)) & " " & Str(dDiff)
        End If
        If (Abs(dDiff) <= m.eWoodHighlight.nLowerPivot) Then
            m.strLowerPivotBkColor = Str(m.eWoodHighlight.nColor)
        Else
            m.strLowerPivotBkColor = "-1"
        End If
    End If
    
    Set tbTemp = Nothing
    Set aIdx = Nothing

    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.ClosestPivots"

End Sub

Private Function FormatData(Bars As cGdBars, ByVal dData#) As String

    Dim bTradingUnits As Boolean
    Dim strFmt$, strReturn$
    
    If m.eDisplayFormat = ePANE_PriceFormat Then
        If Not Bars Is Nothing Then
            FormatData = Bars.PriceDisplay(dData)
        End If
    ElseIf m.eDisplayFormat = ePANE_RoundDecimals Then
        If m.nDisplayDecimals > 0 Then
            strFmt = "0." & String(m.nDisplayDecimals, "0")
        Else
            strFmt = "0"
        End If
        FormatData = Format(dData, strFmt)
    Else
        'auto round
        If dData = Int(dData) Then
            strFmt = "0"
        ElseIf Abs(dData) > 10 Then
            strFmt = "0.00"
        ElseIf Abs(dData) > 0.01 Then
            strFmt = "0.0000"
        Else
            strFmt = "0.000000"
        End If
        FormatData = Format(dData, strFmt)
    End If

End Function

Private Function SetDrawString(ByVal nIdx&, Bars As cGdBars, ByVal dPrice#, ByVal nColor&) As String
On Error GoTo ErrSection:

    Dim strReturn$, strFormat$, strBkColor$, strLocation$, strTextType$
    Dim aFields As New cGdArray
       
'format of saved string stored in array for text
'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|y-location|show
'format of saved string stored in array for price cluster horizontal histogram
'   PP_CLUSTER|min cluster value|max cluster value|optional fixed pane width
    
    strFormat = m.aTemplateStr(nIdx)
    If (Parse(strFormat, "|", 9)) <> "1" Then Exit Function
            
'format for graphics engine to draw text
'    text|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|yLocation|xLocation|<arrayIndex>
    
    strTextType = UCase(Parse(strFormat, "|", 1))
        
    Select Case strTextType
        Case "PP_CLUSTER"
            strReturn = Left(strFormat, Len(strFormat) - 2)
            strReturn = strReturn & "||" & Str(nIdx)        'add "||" as place holder for optional xLocation & array index for hittest purposes
        Case "PP_PRICE"
            strFormat = Right(strFormat, Len(strFormat) - Len(Parse(strFormat, "|", 1)) - 1)    'strip out the textType field
            strFormat = Left(strFormat, Len(strFormat) - 2)                                     'strip off the show flag at end of string in array
            strReturn = FormatData(Bars, dPrice)
            If Len(strReturn) > 0 Then
                strReturn = strReturn & "|" & strFormat
                strReturn = strReturn & "||" & Str(nIdx)    'add "||" as place holder for optional xLocation & array index for hittest purposes
                aFields.SplitFields strReturn, "|"
                aFields(5) = Str(nColor)
                strReturn = aFields.JoinFields("|")
            Else
                strReturn = ""
            End If
        Case "PP_TIMER"
            strBkColor = Parse(strFormat, "|", 7)
            strLocation = Parse(strFormat, "|", 8)
            strFormat = Right(strFormat, Len(strFormat) - Len(Parse(strFormat, "|", 1)) - 1)    'strip out the textType field
            strFormat = Left(strFormat, Len(strFormat) - 2)                                     'strip off the show flag at end of string in array
            strReturn = Left(strFormat, Len(strFormat) - (Len(strBkColor) + Len(strLocation) + 1)) & _
                m.strCountDownBkColor & "|" & strLocation                                       'replace text bkColor with correct background or highlight color for timer
            strReturn = strReturn & "||" & Str(nIdx)        'add "||" as place holder for optional xLocation & array index for hittest purposes
            aFields.SplitFields strReturn, "|"
            aFields(4) = Str(nColor)
            strReturn = aFields.JoinFields("|")
        Case "PP_PRICEDELTA", "PP_PRICEPERCENT"
            strFormat = Right(strFormat, Len(strFormat) - Len(Parse(strFormat, "|", 1)) - 1)    'strip out the textType field
            strFormat = Left(strFormat, Len(strFormat) - 2)                                     'strip off the show flag at end of string in array
            If strTextType = "PP_PRICEDELTA" Then
                strReturn = m.strDeltaLabel
            Else
                strReturn = m.strPercentLabel
            End If
            If Len(strReturn) > 0 Then
                'strReturn = Left(strReturn, Len(strReturn) - 1) 'strip off right parentheses ")"
                strReturn = strReturn & "|" & strFormat
                If InStr(strReturn, "+") <> 0 Then
                    nColor = frmQuotes.UpColor
                ElseIf InStr(strReturn, "-") <> 0 Then
                    nColor = frmQuotes.DownColor
                End If
                strReturn = strReturn & "||" & Str(nIdx)        'add "||" as place holder for optional xLocation & array index for hittest purposes
                aFields.SplitFields strReturn, "|"
                aFields(5) = Str(nColor)
                strReturn = aFields.JoinFields("|")
            Else
                strReturn = ""
            End If
    End Select
    
    SetDrawString = strReturn
    
    Exit Function
    
ErrSection:
    RaiseError "cPaneWood.SetDrawString"

End Function

Private Function SetDrawStringCCI(ByVal eIdx As eCCIText_Index, Bars As cGdBars, _
    strAbove As String, strBelow As String, _
    dHigh#, dPrevHigh#, dDayHigh#, dLow#, dPrevLow#, dDayLow#, dPrice#, _
    iChartBkColor&, iIndColor&) As String
On Error GoTo ErrSection:
    
    Dim strReturn$, dDiff#
    Dim strFont$, strFontSize$, strBold$, strItalic$, strColor$, strBkColor$, strLocation$
                
'format of saved string stored in array
'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|location|show
    strReturn = m.aTemplateStr(eIdx)
    If (Parse(strReturn, "|", 9)) <> "1" Then Exit Function
        
    strFont = Parse(strReturn, "|", 2)
    strFontSize = Parse(strReturn, "|", 3)
    strBold = Parse(strReturn, "|", 4)
    strItalic = Parse(strReturn, "|", 5)
    strColor = Parse(strReturn, "|", 6)
    strBkColor = Parse(strReturn, "|", 7)
    strLocation = Parse(strReturn, "|", 8)
    
    If iChartBkColor = 0 And strColor = "0" Then strColor = Str(iIndColor)
    
    strReturn = ""
    Select Case eIdx
        Case eCCI_DayHigh
            dDiff = (dDayHigh - dPrice) / m.dTickMove
            dDiff = RoundNum(dDiff)
            'turn bk yellow if within specified range of curr price
            If dDiff <= m.eWoodHighlight.nDayHigh Then
                strBkColor = Str(m.eWoodHighlight.nColor)
            Else
                strBkColor = "-1"
            End If
            strReturn = FormatData(Bars, dDayHigh) & " +" & Str(dDiff)
        Case eCCI_PivotAbove
            strReturn = strAbove
            strBkColor = m.strUpperPivotBkColor
        Case eCCI_BarHigh
            'turn bk green if curr bar high > prev bar high
            If dHigh > dPrevHigh Then
                If strColor = "0" Then
                    strBkColor = Str(vbGreen)
                Else
                    strBkColor = Str(RGB(0, 128, 0))
                End If
            Else
                strBkColor = "-1"
            End If
            strReturn = FormatData(Bars, dPrevHigh) & "  " & FormatData(Bars, dHigh)
        Case eCCI_Price
            strReturn = FormatData(Bars, dPrice)
        Case eCCI_BarLow
            'turn bk red if curr bar low < prev bar low
            If dLow < dPrevLow Then
                strBkColor = Str(vbRed)
            Else
                strBkColor = "-1"
            End If
            strReturn = FormatData(Bars, dPrevLow) & "  " & FormatData(Bars, dLow)
        Case eCCI_EMA
            'this line shows diff(EMA, curr price) [space] EMA angle
            If m.eWoodData.dIndEMA <> kNullData Then
                dDiff = (m.eWoodData.dIndEMA - dPrice) / m.dTickMove
                strReturn = RoundNum(dDiff)
                If m.eWoodData.dIndEMAAngle <> kNullData Then
                    strReturn = strReturn & Space(5) & FormatData(Bars, m.eWoodData.dIndEMAAngle)
                End If
            ElseIf m.eWoodData.dIndEMAAngle <> kNullData Then
                strReturn = Space(10) & FormatData(Bars, m.eWoodData.dIndEMAAngle)
            End If
            If m.eWoodData.dIndEMAAngle <> kNullData Then
                'set text color to red if angle < 0 else set text color to green
                If m.eWoodData.dIndEMAAngle > 0 Then
                    strColor = Str(vbCyan)
                ElseIf m.eWoodData.dIndEMAAngle < 0 Then
                    strColor = Str(kMaroon)
                Else
                    strColor = "0"
                End If
            End If
        Case eCCI_Timer
            strBkColor = m.strCountDownBkColor
        Case eCCI_PivotBelow
            strReturn = strBelow
            strBkColor = m.strLowerPivotBkColor
        Case eCCI_DayLow
            dDiff = (dPrice - dDayLow) / m.dTickMove
            strReturn = FormatData(Bars, dDayLow) & " -" & Str(RoundNum(dDiff))
            'turn bk yellow if within specified range of curr price
            If dDiff <= m.eWoodHighlight.nDayLow Then
                strBkColor = Str(m.eWoodHighlight.nColor)
            Else
                strBkColor = "-1"
            End If
        Case eCCI_SideWinder
            If m.eWoodData.dIndSideWinder <> kNullData Then
                strReturn = FormatData(Bars, m.eWoodData.dIndSideWinder)
                If Abs(m.eWoodData.dIndSideWinder) >= m.dTrending Then
                    strReturn = strReturn & "  " & "Trending"
                    strColor = Str(vbGreen)
                ElseIf Abs(m.eWoodData.dIndSideWinder) <= m.dFlat Then
                    strReturn = strReturn & "  " & "Flat"
                    strColor = Str(vbRed)
                Else
                    strReturn = strReturn & "  " & "Normal"
                    strColor = Str(vbYellow)
                End If
            End If
    End Select
        
    If Len(strReturn) > 0 Then
        'format for string that graphics engine uses to draw
        'yLocation: T=top of pane, B=bottom of pane, numeric=at yValue per price scale
        '   text|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|yLocation
        strReturn = strReturn & "|" & strFont & "|" & strFontSize & "|" & strBold & "|" & strItalic & "|" _
            & strColor & "|" & strBkColor & "|" & strLocation & "||" & Str(eIdx)
    ElseIf eIdx = eCCI_Timer Then
        strReturn = strFont & "|" & strFontSize & "|" & strBold & "|" & strItalic & "|" & strColor & "|" _
            & strBkColor & "|" & strLocation & "||" & Str(eIdx)
    End If

    SetDrawStringCCI = strReturn

    Exit Function
    
ErrSection:
    RaiseError "cPaneWood.SetDrawStringCCI"

End Function

Public Function CountDownToNextBar(Chart As cChart, nLastDataBar&) As Long

    Dim nCountDown&
    Dim Bars As cGdBars
    
    nCountDown = -1
    
    Set Bars = Chart.Bars
    
    If Not Bars Is Nothing Then
        Select Case Bars.Prop(eBARS_PeriodType)
            Case ePRD_Minutes
                nCountDown = SecondsToNextBar(Chart, nLastDataBar)
            Case ePRD_Ticks
                nCountDown = TicksToNextBar(Chart, nLastDataBar)
            Case ePRD_IntVol
                nCountDown = VolToNextBar(Chart, nLastDataBar)
        End Select
    End If
    
    Set Bars = Nothing
    CountDownToNextBar = nCountDown

End Function

Private Function VolToNextBar(Chart As cChart, nLastDataBar&)

    Dim nVol As Long
    Dim nVolPerBar As Long
    
    m.strCountDownBkColor = -1
    
    nVol = Chart.Bars(eBARS_Vol, nLastDataBar)
    nVolPerBar = Chart.Bars.Prop(eBARS_PeriodsPerBar)
    
    nVol = nVolPerBar - nVol
    If nVol < 0 Then
        nVol = -1
    ElseIf nVolPerBar > 0 Then
        If nVol / nVolPerBar <= m.eWoodHighlight.dTimerSeconds Then
            m.strCountDownBkColor = Str(m.eWoodHighlight.nColor)
            CheckPlaySound
        End If
    End If
    
    m.nCounterToNextBar = nVol
    
    VolToNextBar = nVol

End Function

Private Function TicksToNextBar(Chart As cChart, nLastDataBar&) As Long

    Dim nTicks As Long
    Dim nTicksPerBar As Long
    
    m.strCountDownBkColor = -1
    
    nTicks = Chart.Bars(eBARS_UpTicks, nLastDataBar) + Chart.Bars(eBARS_DownTicks, nLastDataBar)
    nTicksPerBar = Chart.Bars.Prop(eBARS_PeriodsPerBar)
    
    nTicks = nTicksPerBar - nTicks
    If nTicks < 0 Then
        nTicks = -1
    ElseIf nTicksPerBar > 0 Then
        If nTicks / nTicksPerBar <= m.eWoodHighlight.dTimerSeconds Then
            m.strCountDownBkColor = Str(m.eWoodHighlight.nColor)
            CheckPlaySound
        End If
    End If
    
    m.nCounterToNextBar = nTicks
    
    TicksToNextBar = nTicks

End Function

Private Sub CheckPlaySound()
On Error Resume Next

    Dim i&

    If m.nPlaySound = 0 Then Exit Sub
    
    For i = 0 To m.aTemplateStr.Size - 1
        If Parse(m.aTemplateStr(i), "|", 1) = "PP_Timer" Then Exit For      '4425
    Next
    
    If i >= 0 And i < m.aTemplateStr.Size Then
        'check whether timer is shown then play sound if applicable
        If (Parse(m.aTemplateStr(i), "|", 9)) = 1 Then
            If m.nCounterToNextBar > 0 Then
                If m.nPlaySound = -1 Or Not m.bSoundDone Then
                    PlaySoundFile m.strSoundFile
                    m.bSoundDone = True
                End If
            End If
        End If
    End If

End Sub

Private Function SecondsToNextBar(Chart As cChart, nLastDataBar&) As Long
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim dNow#, dBarTime#, dDiff#, i&
    
    Dim strPeriodicity$, nPeriodicityMinute&
    Dim strDiff$, nMinute&, nSeconds&
    
    ' if streaming replay is paused then don't change the values
    If g.nReplaySession = 0 Or g.RealTime.ReplaySpeed > 0 Then
    
        Set Bars = Chart.Bars
                    
        If Bars.Prop(eBARS_PeriodType) <> ePRD_Minutes Then
            Set Bars = Nothing
            Exit Function
        End If
        
        If g.RealTime.Active And g.RealTime.FeedTime > 0 Then
            strPeriodicity = UCase(Bars.Prop(eBARS_PeriodicityStr))
            If InStr(strPeriodicity, "MINUTE") = 0 Then
                m.nCounterToNextBar = -1
                m.strCountDownBkColor = "-1"
            Else
                nPeriodicityMinute = Bars.Prop(eBARS_PeriodsPerBar)
                dNow = g.RealTime.FeedTime(Bars.Prop(eBARS_Symbol))
                dBarTime = Bars.DateTimeConvert(nLastDataBar, "NY")
                dDiff = (dBarTime - dNow) * 86400#
                                
                If dDiff >= 0 Then
                    m.nCounterToNextBar = Int(dDiff)
                    dDiff = m.nCounterToNextBar / (nPeriodicityMinute * 60)
                    If dDiff <= m.eWoodHighlight.dTimerSeconds Then
                        m.strCountDownBkColor = Str(m.eWoodHighlight.nColor)
                        CheckPlaySound
                    Else
                        m.strCountDownBkColor = "-1"
                        m.bSoundDone = False
                    End If
                Else
                    m.nCounterToNextBar = 0
                    m.strCountDownBkColor = Str(m.eWoodHighlight.nColor)
                End If
            End If
        Else
            m.nCounterToNextBar = -1
            m.strCountDownBkColor = "-1"
            m.bSoundDone = False
        End If
    End If
    
    Set Bars = Nothing
    
    SecondsToNextBar = m.nCounterToNextBar
    
    Exit Function

ErrSection:
    RaiseError "cPaneWood.SecondsToNextBar"
    
End Function

Public Sub SetDrawData(Chart As cChart, Pane As cPane)
On Error GoTo ErrSection:

    Dim i&, iColor&, iChartBkColor&, nSessionDate&
    Dim strData$, nLastDataBar&, nCountDown&
    Dim strAbove$, strBelow$, strCountDown$
    
    Dim dMinMove#, dPrice#
    Dim dLow#, dLowPrevBar#
    Dim dHigh#, dHighPrevBar#
    Dim dDayHigh#, dDayLow#
    
    Dim Bars As cGdBars
    Dim Ind As cIndicator
    Dim hArrayStr As Long
        
    Dim Annot As cAnnotation
    Dim iBold&, iItalic&, iStyle&
    Dim strKey$
    
    If g.bStarting Or g.bUnloading Then Exit Sub
    If Chart Is Nothing Or Pane Is Nothing Then Exit Sub
    If Chart.Bars Is Nothing Then Exit Sub
    
    If Pane.SplitPaneType = ePANE_SplitPaneOptGraph Then
        SetOptNavDrawData Chart, Pane
        Exit Sub
    ElseIf m.aTemplateStr.Size = 0 Then
        If Pane.SplitPaneType = ePANE_SplitPaneTimer Then
            DefaultTemplatePriceStr
        ElseIf Pane.SplitPaneType = ePANE_SplitPaneWood Then
            If UCase(Pane.Name) = "PRICE PANE" Then
                DefaultTemplatePriceStr
            Else
                DefaultTemplateCCIStr
            End If
        Else
            Exit Sub        'no data nothing to do
        End If
    End If
            
    nLastDataBar = Chart.aXBar(Chart.ScreenEndX)
    i = Chart.LastGoodDataBar(False)
    If nLastDataBar <= 0 Then
        Exit Sub          'this can occur during template changes etc.
    ElseIf nLastDataBar > i Then
        nLastDataBar = i
    End If
    
    hArrayStr = Pane.SplitPaneStrArray
    
    'clear out existing data
    gdSetSize hArrayStr, 0, 0
    
    Set Bars = Chart.Bars
    
    dMinMove = Bars.MinMove(Bars.SessionDate(nLastDataBar))
    dPrice = Bars(eBARS_Close, nLastDataBar)
    
    m.dTickMove = Bars.Prop(eBARS_TickMove)
    m.eDisplayFormat = Pane.DisplayFormat
    m.nDisplayDecimals = Pane.DisplayDecimals
        
    strCountDown = ""
    If g.RealTime.Active Then
        Select Case Bars.Prop(eBARS_PeriodType)
            Case ePRD_Minutes
                nCountDown = SecondsToNextBar(Chart, nLastDataBar)
                If nCountDown >= 0 Then strCountDown = Str(nCountDown) & "s"
            Case ePRD_Ticks
                nCountDown = TicksToNextBar(Chart, nLastDataBar)
                If nCountDown >= 0 Then strCountDown = Str(nCountDown) & "t"
            Case ePRD_IntVol
                nCountDown = VolToNextBar(Chart, nLastDataBar)
                If nCountDown >= 0 Then strCountDown = Str(nCountDown) & "v"
        End Select
    End If
        
    iColor = 0
    If Chart.UseCustomColors Then
        iChartBkColor = Chart.ChartBackColor
    Else
        iChartBkColor = g.ChartGlobals.nChartBackColor
    End If
    If iChartBkColor = 0 Then
        Set Ind = Chart.Tree("PRICE")
        If Not Ind Is Nothing Then iColor = Ind.Color
        Set Ind = Nothing
    End If
    If Pane.PricePaneFlag Then
        If Pane.SplitPaneType = ePANE_SplitPaneCluster Then
            For i = 2 To m.aTemplateStr.Size - 1
                If Not IsNumeric(Parse(m.aTemplateStr(i), "|", 8)) Then
                    strData = SetDrawString(i, Bars, dPrice, 0)
                Else
                    strData = SetDrawString(i, Bars, Parse(m.aTemplateStr(i), "|", 8), 0)
                End If
                If Len(strData) > 0 Then gdSetStr hArrayStr, i, strData
            Next
        Else
            CheckPriceDelta Pane
            For i = 0 To m.aTemplateStr.Size - 1
                If InStr(m.aTemplateStr(i), "PP_Timer") = 0 Then
                    strData = SetDrawString(i, Bars, dPrice, iColor)
                    If Len(strData) > 0 Then gdSetStr hArrayStr, i, strData
                ElseIf g.RealTime.Active Then
                    If (Parse(m.aTemplateStr(i), "|", 9)) = 1 Then
                        If m.nPlaySound <> 0 And Len(strCountDown) > 0 Then strCountDown = strCountDown & "A"
                        strData = strCountDown & "|" & SetDrawString(i, Bars, dPrice, iColor)
                        gdSetStr hArrayStr, i, strData
                    End If
                End If
            Next
        End If
    
    ElseIf UCase(Pane.Name) = "WOODIES CCI" Then
        
        SetWoodIndData Chart, dMinMove
        ClosestPivots Bars, strAbove, strBelow, dPrice

        dHigh = Bars(eBARS_High, nLastDataBar)
        dLow = Bars(eBARS_Low, nLastDataBar)
        dHighPrevBar = Bars(eBARS_High, nLastDataBar - 1)
        dLowPrevBar = Bars(eBARS_Low, nLastDataBar - 1)
        
        dDayHigh = dHigh
        dDayLow = dLow
        If Bars.IsIntraday Then
            ' get high & low for day (same trading session)
            nSessionDate = Bars.SessionDate(nLastDataBar)
            For i = nLastDataBar - 1 To 0 Step -1
                If Bars.SessionDate(i) <> nSessionDate Then Exit For
                If Bars(eBARS_High, i) > dDayHigh Then dDayHigh = Bars(eBARS_High, i)
                If Bars(eBARS_Low, i) < dDayLow Then dDayLow = Bars(eBARS_Low, i)
            Next
        End If
        
        'high of day at pane's top
        strData = SetDrawStringCCI(eCCI_DayHigh, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
        If Len(strData) > 0 Then gdSetStr hArrayStr, gdGetSize(hArrayStr), strData
        'closest pivot above price @ y=210
        strData = SetDrawStringCCI(eCCI_PivotAbove, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
        If Len(strData) > 0 Then gdSetStr hArrayStr, gdGetSize(hArrayStr), strData
        'prev/curr bar high @ y=180
        strData = SetDrawStringCCI(eCCI_BarHigh, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
        If Len(strData) > 0 Then gdSetStr hArrayStr, gdGetSize(hArrayStr), strData
        'price @ y=140
        strData = SetDrawStringCCI(eCCI_Price, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
        If Len(strData) > 0 Then gdSetStr hArrayStr, gdGetSize(hArrayStr), strData
        'prev/curr bar low @ y=100
        strData = SetDrawStringCCI(eCCI_BarLow, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
        If Len(strData) > 0 Then gdSetStr hArrayStr, gdGetSize(hArrayStr), strData
        'EMA @ y = -100
        strData = SetDrawStringCCI(eCCI_EMA, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
        If Len(strData) > 0 Then gdSetStr hArrayStr, gdGetSize(hArrayStr), strData
        'count down @ y = -150
        If g.RealTime.Active Then
            If (Parse(m.aTemplateStr(eCCI_Timer), "|", 9)) = 1 Then
                strData = strCountDown & "|" & SetDrawStringCCI(eCCI_Timer, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
                gdSetStr hArrayStr, gdGetSize(hArrayStr), strData
            End If
        End If
        'closest pivot below price @ y = -200
        strData = SetDrawStringCCI(eCCI_PivotBelow, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
        If Len(strData) > 0 Then gdSetStr hArrayStr, gdGetSize(hArrayStr), strData
        'low of day at pane's bottom
        strData = SetDrawStringCCI(eCCI_DayLow, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
        If Len(strData) > 0 Then gdSetStr hArrayStr, gdGetSize(hArrayStr), strData
        
        'sidewinder value in upper left
        strData = SetDrawStringCCI(eCCI_SideWinder, Bars, strAbove, strBelow, dHigh, dHighPrevBar, dDayHigh, dLow, dLowPrevBar, dDayLow, dPrice, iChartBkColor, iColor)
        
        strKey = CStr(m.eWoodData.nIndSideWinderID) & " SideWinderLabel"
        Set Annot = Chart.Annots(strKey)
        
        If Len(strData) > 0 Then
            If Annot Is Nothing Then
                Set Annot = New cAnnotation
                Annot.CreateNew Chart, eANNOT_Icon, Pane.gePaneId, _
                    -1, -1, 0, 0, 0, PEGAT_NOSYMBOL, _
                    Parse(strData, "|", 1), eANNOT_SideWinderLabel
                Chart.Annots.Add Annot, strKey
            End If
            
            iBold = ValOfText(Parse(strData, "|", 4))
            iItalic = ValOfText(Parse(strData, "|", 5))
            If iBold = 0 And iItalic = 0 Then
                iStyle = 0
            ElseIf iBold = 1 And iItalic = 0 Then
                iStyle = 1
            ElseIf iBold = 0 And iItalic = 1 Then
                iStyle = 2
            ElseIf iBold = 1 And iItalic = 1 Then
                iStyle = 3
            End If
        
            Annot.geIndId = m.eWoodData.nIndSideWinderID
            Annot.gePaneId = Pane.gePaneId
            Annot.Text = Parse(strData, "|", 1)
            Annot.Prop("FontName") = Parse(strData, "|", 2)
            Annot.Prop("FontSize") = Parse(strData, "|", 3)
            Annot.Prop("FontStyle") = iStyle
            Annot.Color = Parse(strData, "|", 6)
        ElseIf Not Annot Is Nothing Then
            Chart.RemoveAnnots True, eANNOT_Icon, eANNOT_SideWinderLabel
        End If
    End If

    Set Bars = Nothing
    Set Ind = Nothing
    Set Annot = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.SetDrawData"

End Sub

Private Sub DefaultTemplatePriceStr()
On Error GoTo ErrSection:

'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|location|show
    m.aTemplateStr.Size = 0
    m.aTemplateStr.Add "PP_Price|" & DefaultFont & "|12|1|0|0|-1|T|1"
    m.aTemplateStr.Add "PP_PriceDelta|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|-1|N|1"
    m.aTemplateStr.Add "PP_PricePercent|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|-1|N|1"
    m.aTemplateStr.Add "PP_Timer|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|-1|B|1"

    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.DefaultTemplatePriceStr"

End Sub

Private Sub DefaultAlertLevel()
    m.eWoodHighlight.nColor = vbYellow
    m.eWoodHighlight.nDayHigh = 10
    m.eWoodHighlight.nDayLow = 10
    m.eWoodHighlight.nUpperPivot = 3
    m.eWoodHighlight.nLowerPivot = 3
    m.eWoodHighlight.dTimerSeconds = 0.15     'this is a % value (cannot be > 1)
    m.eWoodHighlight.dTimerTicks = 0.07       'this is a % value
End Sub

Private Sub DefaultTemplateCCIStr()
On Error GoTo ErrSection:

    Dim i&
    Dim strColor&, strFontSize$
    
    m.aTemplateStr.Size = 0
        
'SplitPaneWoodCCI (text type & corresponding array index)
'   text-type       array index         description
'   DayHigh         DayHigh             'high of day
'   PivotAbove      PivotAbove          'closest pivot point above price (eg R1 +9)
'   BarHigh         BarHigh             'prev bar high [space] curr bar high
'   Price           Price               'current price
'   BarLow          BarLow              'prev bar low [space] curr bar low
'   EMA             EMA                 'EMA ticks away from price and EMA angle
'   Timer           Timer               'count down to next bar in ticks or seconds
'   PivotBelow      PivotBelow          'closest pivot point below price (eg S1 -3)
'   DayLow          DayLow              'low of day
'   SideWinder      SideWinder          'sidewinder value in upper left corner under indicator labels
        
'default location specified by Pete
'high of day at pane's top
'closest pivot above price @ y=210
'prev/curr bar high @ y=180
'price @ y=140
'prev/curr bar low @ y=100
'EMA @ y = -100
'count down @ y = -150
'closest pivot below price @ y = -200
'low of day at pane's bottom

'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|location|show
    m.aTemplateStr.Add "DayHigh|" & DefaultFont & "|" & DefaultFontSize & "|1|0|" & Str(vbBlue) & "|-1|T|1"
    m.aTemplateStr.Add "PivotAbove|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|-1|210|1"
    m.aTemplateStr.Add "BarHigh|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|-1|180|1"
    m.aTemplateStr.Add "Price|" & DefaultFont & "|" & "12|1|0|0|-1|140|1"
    m.aTemplateStr.Add "BarLow|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|-1|100|1"
    m.aTemplateStr.Add "EMA|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|-1|-100|1"
    m.aTemplateStr.Add "Timer|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|-1|-150|1"
    m.aTemplateStr.Add "PivotBelow|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|-1|-200|1"
    m.aTemplateStr.Add "DayLow|" & DefaultFont & "|" & DefaultFontSize & "|1|0|" & Str(vbRed) & "|-1|B|1"
    
    'sidewinder value is actually implemented as an annot with fixed colors based on trending text
    'colors & location are set in CCI drawstring routines (colors & location set here are ignored)
    m.aTemplateStr.Add "SideWinder|" & DefaultFont & "|" & DefaultFontSize & "|1|0|0|0|0|1"
    
    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.DefaultTemplateCCIStr"

End Sub

Public Sub TemplateSave(Pane As cPane, ByVal fh%)
On Error GoTo ErrSection:

    Dim strSave$, i&
    
    
    If Pane.SplitPaneType = ePANE_SplitPaneOptGraph Then Exit Sub
    
    If Pane.PricePaneFlag Then
        strSave = m.aTemplateStr.JoinFields("~")
        If Len(Trim(strSave)) > 0 Then Print #fh, IniString("SplitPaneWoodPrice", strSave)
        Print #fh, IniString("PaneWoodHLColor", m.eWoodHighlight.nColor)
        Print #fh, IniString("PaneWoodHLTimerS", m.eWoodHighlight.dTimerSeconds)
        Print #fh, IniString("PaneWoodPlaySound", m.nPlaySound)
        Print #fh, IniString("PaneWoodSoundFile", m.strSoundFile)
    ElseIf UCase(Pane.Name) = "WOODIES CCI" Then
        strSave = m.aTemplateStr.JoinFields("~")
        If Len(Trim(strSave)) > 0 Then Print #fh, IniString("SplitPaneWoodCCI", strSave)
        Print #fh, IniString("PaneWoodHLColor", m.eWoodHighlight.nColor)
        Print #fh, IniString("PaneWoodHLDayHigh", m.eWoodHighlight.nDayHigh)
        Print #fh, IniString("PaneWoodHLDayLow", m.eWoodHighlight.nDayLow)
        Print #fh, IniString("PaneWoodHLPivotUpper", m.eWoodHighlight.nUpperPivot)
        Print #fh, IniString("PaneWoodHLPivotLower", m.eWoodHighlight.nLowerPivot)
        Print #fh, IniString("PaneWoodHLTimerS", m.eWoodHighlight.dTimerSeconds)
        Print #fh, IniString("PaneWoodHLTimerT", m.eWoodHighlight.dTimerTicks)
        Print #fh, IniString("PaneWoodTrending", m.dTrending)
        Print #fh, IniString("PaneWoodFlat", m.dFlat)
    End If

    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.TemplateSave"

End Sub

' TLB #5904: should just pass and use dValue into this routine instead of using
' the "ValOfText" function (which was causing a regional settings issue)
Public Sub TemplateLoad(ByVal strProp$, ByVal strValue$, ByVal dValue#)
On Error GoTo ErrSection:

    Dim i&, j&, nIdx&
    Dim aRows As New cGdArray
    Dim aFields As New cGdArray
    
    Select Case UCase(strProp)
    
        Case "SPLITPANEWOODCCI"
            If Len(Trim(strValue)) > 0 Then
                aRows.SplitFields strValue, "~"
                For i = 0 To aRows.Size - 1
                    m.aTemplateStr(i) = aRows(i)
                Next
            End If
        
        Case "SPLITPANEWOODPRICE"
            If Len(Trim(strValue)) > 0 Then aRows.SplitFields strValue, "~"
            If aRows.Size > 3 Then
                For i = 0 To aRows.Size - 1
                    m.aTemplateStr(i) = aRows(i)
                Next
            Else
                DefaultTemplatePriceStr     'get new delta percent row
            End If
        Case "PANEWOODHLCOLOR"
            m.eWoodHighlight.nColor = dValue
        
        Case "PANEWOODHLDAYHIGH"
            m.eWoodHighlight.nDayHigh = dValue
        
        Case "PANEWOODHLDAYLOW"
            m.eWoodHighlight.nDayLow = dValue
                    
        Case "PANEWOODHLPIVOTUPPER"
            m.eWoodHighlight.nUpperPivot = dValue
        
        Case "PANEWOODHLPIVOTLOWER"
            m.eWoodHighlight.nLowerPivot = dValue
        
        Case "PANEWOODHLTIMERS"
            If dValue <= 0 Then dValue = 0.15 ' TLB: set back to default if messed up due to previous regional settings bug
            m.eWoodHighlight.dTimerSeconds = dValue
        
        Case "PANEWOODHLTIMERT"
            m.eWoodHighlight.dTimerTicks = dValue
        
        Case "PANEWOODTRENDING"
            m.dTrending = dValue
        
        Case "PANEWOODFLAT"
            m.dFlat = dValue
        
        Case "PANEWOODPLAYSOUND"
            m.nPlaySound = dValue
                        
        Case "PANEWOODSOUNDFILE"
            m.strSoundFile = strValue
    End Select
    
    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.TemplateLoad"

End Sub

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    DefaultAlertLevel
    
    'default for sidewinder limits
    m.dFlat = 75#
    m.dTrending = 100#
            
    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.Class_Initialize"

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.aTemplateStr = Nothing
    
    Set m.aPrice = Nothing
    Set m.aNow = Nothing
    Set m.aThird = Nothing
    Set m.aTwoThird = Nothing
    Set m.aExpire = Nothing
    
    Set m.aProb1 = Nothing
    Set m.aProb2 = Nothing

    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.Class_Terminate"

End Sub

Public Function GetCCITextInfo(ByVal iIndex As Long, Optional ByVal strTextType$) As String
On Error GoTo ErrSection:
    
    Dim i As Long
    Dim strText As String

'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|location|show
    If strTextType = "PP_Price" Or strTextType = "PP_Timer" Or strTextType = "PP_PriceDelta" Or strTextType = "PP_PricePercent" Then
        For i = 0 To m.aTemplateStr.Size - 1
            strText = m.aTemplateStr(i)
            If Parse(strText, "|", 1) = strTextType Then
                Exit For
            Else
                strText = ""
            End If
        Next
        GetCCITextInfo = strText
    ElseIf strTextType = "PP_Cluster" Then
        GetCCITextInfo = Trim(m.aTemplateStr(iIndex))
    Else
        GetCCITextInfo = Trim(m.aTemplateStr(iIndex))
    End If
    
    Exit Function

ErrSection:
    RaiseError "cPaneWood.GetCCITextInfo"
    
End Function

Public Sub SetCCITextInfo(ByVal iIndex As Long, ByVal strInfo As String, Optional ByVal strTextType$)
On Error GoTo ErrSection:

    Dim i As Long

'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|location|show
    If strTextType = "PP_Price" Or strTextType = "PP_PriceDelta" Or strTextType = "PP_PricePercent" Or strTextType = "PP_Timer" Then
        For i = 0 To m.aTemplateStr.Size - 1
            If Parse(m.aTemplateStr(i), "|", 1) = strTextType Then
                m.aTemplateStr(i) = strInfo
                Exit For
            End If
        Next
    ElseIf strTextType = "PP_Timer" Then
        m.aTemplateStr(1) = strInfo
    ElseIf strTextType = "PP_Cluster" Then
        m.aTemplateStr(iIndex) = strInfo
    Else
        m.aTemplateStr(iIndex) = strInfo
    End If
    
    Exit Sub
        
ErrSection:
    RaiseError "cPaneWood.SetCCITextInfo"
    
End Sub

Public Sub GetHighlightLevels(nColor&, nDayHigh&, nDayLow&, nUpperPivot&, nLowerPivot&, dTimerS#, dTimerT#)
On Error GoTo ErrSection:

    nColor = m.eWoodHighlight.nColor
    nDayHigh = m.eWoodHighlight.nDayHigh
    nDayLow = m.eWoodHighlight.nDayLow
    nUpperPivot = m.eWoodHighlight.nUpperPivot
    nLowerPivot = m.eWoodHighlight.nLowerPivot
    dTimerS = m.eWoodHighlight.dTimerSeconds
    dTimerT = m.eWoodHighlight.dTimerTicks
    
    Exit Sub
    
ErrSection:
    RaiseError "cPaneWood.GetHighlightLevels"
    
End Sub

Public Property Let HighlightColor(ByVal nColor&)
    m.eWoodHighlight.nColor = nColor
End Property

Public Property Let HighlightDayHigh(ByVal nDayHigh&)
    m.eWoodHighlight.nDayHigh = nDayHigh
End Property

Public Property Let HighlightDayLow(ByVal nDayLow&)
    m.eWoodHighlight.nDayLow = nDayLow
End Property

Public Property Let HighlightPivotUpper(ByVal nPivotUpper&)
    m.eWoodHighlight.nUpperPivot = nPivotUpper
End Property

Public Property Let HighlightPivotLower(ByVal nPivotLower&)
    m.eWoodHighlight.nLowerPivot = nPivotLower
End Property

Public Property Let HighlightTimerS(ByVal dTimerS#)
    If dTimerS > 0 And dTimerS <= 1 Then
        m.eWoodHighlight.dTimerSeconds = dTimerS
    End If
End Property

Public Property Get SideWinderFlat() As Double
    SideWinderFlat = m.dFlat
End Property

Public Property Let SideWinderFlat(ByVal dFlat#)
    m.dFlat = dFlat
End Property

Public Property Get SideWinderTrending() As Double
    SideWinderTrending = m.dTrending
End Property

Public Property Let SideWinderTrending(ByVal dTrending#)
    m.dTrending = dTrending
End Property

Public Property Get LastTimerVal() As Long
    LastTimerVal = m.nCounterToNextBar
End Property

Public Property Get SoundFile() As String
    SoundFile = m.strSoundFile
End Property

Public Property Let SoundFile(ByVal strFile$)
    m.strSoundFile = strFile
End Property

Public Property Get PlaySound() As Long
    PlaySound = m.nPlaySound
End Property

Public Property Let PlaySound(ByVal nPlay&)
    m.nPlaySound = nPlay
End Property

Public Property Get ArrayPriceStrSize() As Long
    ArrayPriceStrSize = m.aTemplateStr.Size
End Property

Public Property Get TimerOnOff() As Long
    
    If m.aTemplateStr.Size > 1 Then
        PriceOnOff = ValOfText(Parse(m.aTemplateStr(1), "|", 9))
    End If
    
End Property

Public Property Let TimerOnOff(ByVal nOnOff&)

    Dim aFields As cGdArray
    
    If m.aTemplateStr.Size > 1 Then
        Set aFields = New cGdArray
        aFields.SplitFields m.aTemplateStr(1), "|"
        If aFields.Size >= 8 Then
            aFields(8) = nOnOff
        End If
        m.aTemplateStr(1) = aFields.JoinFields("|")
    End If

End Property

Public Property Get PriceOnOff() As Long
    
    If m.aTemplateStr.Size > 0 Then
        PriceOnOff = ValOfText(Parse(m.aTemplateStr(0), "|", 9))
    End If
    
End Property

Public Property Let PriceOnOff(ByVal nOnOff&)

    Dim aFields As cGdArray
    
    If m.aTemplateStr.Size > 0 Then
        Set aFields = New cGdArray
        aFields.SplitFields m.aTemplateStr(0), "|"
        If aFields.Size >= 8 Then
            aFields(8) = nOnOff
        End If
        m.aTemplateStr(0) = aFields.JoinFields("|")
    End If

End Property

Public Sub ClearPriceCluster()
    If Not m.aTemplateStr Is Nothing Then m.aTemplateStr.Size = 2
End Sub

Public Property Get TimerBkColor() As String
    TimerBkColor = m.strCountDownBkColor
End Property

Public Property Let DeltaLabel(ByVal strLabel$)
    m.strDeltaLabel = strLabel
End Property

Public Property Let PercentLabel(ByVal strLabel$)
    m.strPercentLabel = strLabel
End Property

Private Sub CheckPriceDelta(Pane As cPane)

    Dim aFields As cGdArray
    
    If Pane Is Nothing Or m.bPriceDeltaChecked Then Exit Sub
        
    If Pane.SplitPaneType = ePANE_SplitPaneTimer Then
        If m.aTemplateStr.Size > 0 Then
            If Parse(m.aTemplateStr(1), "|", 1) <> "PP_PriceDelta" Then
                Set aFields = New cGdArray
                aFields.SplitFields m.aTemplateStr(0), "|"
                aFields(0) = "PP_PriceDelta"
                aFields(7) = "N"
                aFields(8) = "1"
                m.aTemplateStr.Add aFields.JoinFields("|"), 1
            End If
        End If
    End If
            
    m.bPriceDeltaChecked = True

End Sub

Public Sub OptNavGraphInfoSet(ByVal strGraph$, aPrice As cGdArray, aNow As cGdArray, aThird As cGdArray, _
    aTwoThird As cGdArray, aExpire As cGdArray, aProb1 As cGdArray, aProb2 As cGdArray)
On Error GoTo ErrSection:
    
    m.strOptNavGraphType = strGraph
    Set m.aPrice = aPrice
    Set m.aNow = aNow
    Set m.aThird = aThird
    Set m.aTwoThird = aTwoThird
    Set m.aExpire = aExpire
    
    Set m.aProb1 = aProb1
    Set m.aProb2 = aProb2
    
    
    If m.aProb1 Is Nothing Or m.aProb2 Is Nothing Then
        'save of dates for tooltip
        m.dDateNow = DateOf(Left(Parse(aNow(0), ";", 1), 8))
        m.dDateThird = DateOf(Left(Parse(aThird(0), ";", 1), 8))
        m.dDateTwoThird = DateOf(Left(Parse(aTwoThird(0), ";", 1), 8))
        m.dDateExpire = DateOf(Left(Parse(aExpire(0), ";", 1), 8))
        
        Set m.aProb1 = Nothing
        Set m.aProb2 = Nothing
    Else
        m.dDateNow = aPrice(0)          'volatility
        m.dDateThird = aNow(0)          '1st probability %
        m.dDateTwoThird = aTwoThird(0)  '2nd probability %
        m.dDateExpire = aProb2(0)       '3rd probability %
        
        m.aProb1.Remove 0
        m.aProb2.Remove 0
        
    End If
    
    'remove the first array element (all remaining items are data points for plotting graph)
    m.aPrice.Remove 0
    m.aNow.Remove 0
    m.aThird.Remove 0
    m.aTwoThird.Remove 0
    m.aExpire.Remove 0
    
    m.aTemplateStr.Size = 0
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPaneWood.OptNavGraphInfoSet"

End Sub

Private Sub SetOptNavDrawData(Chart As cChart, Pane As cPane)
On Error GoTo ErrSection:

    Dim strData$
    Dim bOkay As Boolean
    Dim hArrayStr As Long
    
    m.aTemplateStr.Size = 0         'not used for drawing risk graphs
    
    'make sure we have y-values
    If m.aPrice Is Nothing Then Exit Sub
    If m.aPrice.Size = 0 Then Exit Sub
    
    'make sure we have at least one array of x-values
    If Not m.aNow Is Nothing Then
        If m.aNow.Size > 0 Then bOkay = True
    End If
    If Not bOkay And Not m.aThird Is Nothing Then
        If m.aThird.Size > 0 Then bOkay = True
    End If
    If Not bOkay And Not m.aTwoThird Is Nothing Then
        If m.aTwoThird.Size > 0 Then bOkay = True
    End If
    If Not bOkay And Not m.aExpire Is Nothing Then
        If m.aExpire.Size > 0 Then bOkay = True
    End If
    If Not bOkay Then Exit Sub
    
    'clear out existing data
    hArrayStr = Pane.SplitPaneStrArray
    gdSetSize hArrayStr, 0, 0
    
    'Format: PP_OPTGRAPH|penSize|penColor|data points
    
    strData = m.aPrice.JoinFields(";")
    gdSetStr hArrayStr, 0, "PP_OPTGRAPH|1|0|" & strData
    
    If Not m.aNow Is Nothing Then
        strData = m.aNow.JoinFields(";")
        If Len(strData) > 0 Then
            gdSetStr hArrayStr, gdGetSize(hArrayStr), "PP_OPTGRAPH|1|255|" & strData      'red
        End If
    End If

    If Not m.aThird Is Nothing Then
        strData = m.aThird.JoinFields(";")
        If Len(strData) > 0 Then
            gdSetStr hArrayStr, gdGetSize(hArrayStr), "PP_OPTGRAPH|1|16711680|" & strData  'blue
        End If
    End If

    If Not m.aTwoThird Is Nothing Then
        strData = m.aTwoThird.JoinFields(";")
        If Len(strData) > 0 Then
            gdSetStr hArrayStr, gdGetSize(hArrayStr), "PP_OPTGRAPH|1|16384|" & strData      'green
        End If
    End If

    If Not m.aExpire Is Nothing Then
        strData = m.aExpire.JoinFields(";")
        If Len(strData) > 0 Then
            If g.ChartGlobals.nChartBackColor = 0 Then
                gdSetStr hArrayStr, gdGetSize(hArrayStr), "PP_OPTGRAPH|1|" & Str(RGB(0, 254, 0)) & "|" & strData    'same color as FixColors in cChart -5227
            Else
                gdSetStr hArrayStr, gdGetSize(hArrayStr), "PP_OPTGRAPH|1|0|" & strData                  'black
            End If
        End If
    End If
    
    If Not m.aProb1 Is Nothing Then
        strData = m.aProb1.JoinFields(";")
        If Len(strData) > 0 Then
            gdSetStr hArrayStr, gdGetSize(hArrayStr), "PP_OPTGRAPH|1|16384|" & strData      'green
        End If
    End If

    If Not m.aProb2 Is Nothing Then
        strData = m.aProb2.JoinFields(";")
        If Len(strData) > 0 Then
            gdSetStr hArrayStr, gdGetSize(hArrayStr), "PP_OPTGRAPH|1|16711680|" & strData  'blue
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPaneWood.SetOptNavDrawData"

End Sub

Public Sub OptNavGraphDate(dDateNow#, dDateThird#, dDateTwoThird#, dDateExpire#)
On Error GoTo ErrSection:

    dDateNow = m.dDateNow
    dDateThird = m.dDateThird
    dDateTwoThird = m.dDateTwoThird
    dDateExpire = m.dDateExpire

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPaneWood.OptNavGraphDate"

End Sub

Public Sub OptNavGraphValues(ByVal dPrice#, dNow#, dThird#, dTwoThird#, dExpire#)
On Error GoTo ErrSection:

    Dim i&

    'initialize return values
    dNow = kNullData
    dThird = kNullData
    dTwoThird = kNullData
    dExpire = kNullData

    If Not m.aPrice Is Nothing Then
        m.aPrice.BinarySearch dPrice, i
        
        If i >= 0 And i < m.aPrice.Size Then
            If Not m.aNow Is Nothing Then dNow = m.aNow(i)
            If Not m.aThird Is Nothing Then dThird = m.aThird(i)
            If Not m.aTwoThird Is Nothing Then dTwoThird = m.aTwoThird(i)
            If Not m.aExpire Is Nothing Then dExpire = m.aExpire(i)
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPaneWood.OptNavGraphValues"

End Sub

Public Function OptNavGraphToolTip() As String
On Error GoTo ErrExit

    Dim strTip$
    
    If m.aProb1 Is Nothing Or m.aProb2 Is Nothing Then Exit Function
    
    strTip = "Price = " & RoundNum(m.aNow(0), 2) & vbCrLf
    strTip = strTip & "Volatility = " & RoundNum(m.dDateNow, 2)
    
    OptNavGraphToolTip = strTip

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPaneWood.OptNavGraphToolTip"

End Function
