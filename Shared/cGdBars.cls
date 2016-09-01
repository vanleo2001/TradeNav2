VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGdBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Requires "cGdArray.cls", "mGdDll.bas", and "G32_GD.DLL"
Option Explicit

Private Type mPrivate
    hBars As Long
    bAutoDestroy As Boolean '(normally TRUE)
    
    strPriceThresholds As String        ' Semicolon delimited list of price thresholds
    strSecondaryMinMoves As String      ' Semicolon delimited list of secondary min moves
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    
    m.hBars = 0
    m.bAutoDestroy = True 'by default
    ArrayMask = eBARS_Eod
    
End Sub

Private Sub Class_Terminate()
       
    Destroy
    
End Sub

Private Sub Destroy()
    
    If m.hBars <> 0 And m.bAutoDestroy Then
#If TRADENAV_EXE Then
        If Not g.RealTime Is Nothing Then
            g.RealTime.RemoveTickBuffer Me
        End If
#End If
        gdDestroyBars m.hBars
        If m.hBars <> 0 Then
            'HOUSTON, WE HAVE A PROBLEM!  Somehow, we still have
            'a pointer to something that used to be a valid gdBars
            'but it had already been destroyed.
            '-- THIS INDICATES A PROGRAMMING ERROR (most likely
            'the "SetBarsHandle" was used in an invalid fashion).
            m.hBars = 0
        End If
    End If
    m.hBars = 0
    m.bAutoDestroy = True 'reset flag to default
    
End Sub

Public Property Get Size() As Long
    Size = gdGetSize(m.hBars)
End Property

Public Property Let Size(ByVal nNewSize As Long)
    gdSetSize m.hBars, nNewSize, True
End Property

' Frees any extra memory from all the price/vol arrays
Public Sub FreeExtra()
    gdFreeExtra m.hBars
End Sub

Public Property Get Item(ByVal WhichArray As eBarsArray, ByVal nOffset&) As Double
Attribute Item.VB_UserMemId = 0
    Item = gdBarsData(m.hBars, WhichArray, nOffset)
End Property

Public Property Let Item(ByVal WhichArray As eBarsArray, ByVal nOffset&, ByVal dValue As Double)
    Dim hArray&
    hArray = gdBarsArray(m.hBars, WhichArray)
    If hArray Then
        gdSetNum hArray, nOffset, dValue
    End If
End Property

Public Property Get ArrayHandle(ByVal WhichArray As eBarsArray) As Long
    ArrayHandle = gdBarsArray(m.hBars, WhichArray)
End Property

'Returns the "handle" to the gdBars (mostly used in DLL calls)
'- if bDetach = True, then the gdBars is detached from the
'   cGdBars object (up to caller to destroy the handle later)
Public Property Get BarsHandle(Optional ByVal bDetach As Boolean = False) As Long
    Dim h&, bAutoDestroy As Boolean
    h = m.hBars
    If bDetach Then
        bAutoDestroy = m.bAutoDestroy
        m.bAutoDestroy = False
        Destroy
        m.bAutoDestroy = bAutoDestroy
    End If
    BarsHandle = h
End Property

Public Property Get SymbolOrSymbolID() As Variant
    
    If Prop(eBARS_SymbolID) = 0 Then
        SymbolOrSymbolID = Prop(eBARS_Symbol)
    Else
        SymbolOrSymbolID = Prop(eBARS_SymbolID)
    End If

End Property

'This is allowed for some flexibility -- BUT SHOULD BE USED
'WITH GREAT CAUTION (know what you're doing and be careful)!
Public Sub SetBarsHandle(ByVal NewBarsHandle As Long, _
        ByVal bAutoDestroy As Boolean)
    
    Destroy 'destroy current bars if exists
    m.hBars = NewBarsHandle
    
    ' set flag to determine if will handle auto-destroying
    ' the gdBars when this object gets done with it
    m.bAutoDestroy = bAutoDestroy
    
End Sub

' To read bars from a data file.
' - strFormat: "CSI", "MS7", "GEN" (Genesis EOD), "GT" (GenTick)
' - strPath: path where data file is located
' - strSymbol: pass either the symbol (leave blank to use the current bars.Symbol)
'      or pass the name of the data file to load (e.g. "F104.DTA")
' - strPeriod: see "BuildBars" for periodicity options
' - nAppendMode: >0 = append at bar record, 0 = by dates, -1 = by file record #
' - returns # of bars loaded, or negative number for error
Public Function FromFile(ByVal strFormat$, _
        ByVal strPath$, ByVal strSymbol$, _
        Optional ByVal strPeriod$ = "", Optional strErrMsg$, _
        Optional ByVal vFromDate As Variant = 0, _
        Optional ByVal vToDate As Variant = 0, _
        Optional ByVal nAppendMode& = 0, _
        Optional ByVal bAutoSetBarType = True) As Long
    
    Dim nNumBars&, hErrMsg&, hDates&, i&
    
    hErrMsg = gdCreateString(0)
    hDates = 0
    
    ' fix some things with Symbol and Path
    strFormat = UCase(Trim(strFormat))
    strPath = UCase(Trim(strPath))
    If strFormat = "GT" Then
        If InStr(strSymbol, ".GT") = 0 Then
            strSymbol = strSymbol & ".GT"
        End If
    ElseIf strFormat = "CSI" Or Left(strFormat, 2) = "MS" Then
        If Right(strPath, 8) = "\QMASTER" Or Right(strPath, 7) = "\MASTER" Or Right(strPath, 4) = ".DTA" Or Right(strPath, 4) = ".DAT" Then
            For i = Len(strPath) To 1 Step -1
                If Mid(strPath, i, 1) = "\" Then
                    strPath = Left(strPath, i)
                    Exit For
                End If
            Next
        End If
    End If
    
    If bAutoSetBarType Then
        If strFormat = "GT" Then
            ArrayMask = eBARS_Intraday
        Else
            ArrayMask = eBARS_Eod
        End If
    End If
    
    If nAppendMode = 0 Then Size = 0
    
    nNumBars = gdBarsFromFile(m.hBars, strFormat, strPath, strSymbol, _
        strPeriod, CLng(vFromDate), CLng(vToDate), nAppendMode, hErrMsg, hDates)
    
    If Not IsMissing(strErrMsg) Then
        If nNumBars >= 0 Then
            strErrMsg = ""
        Else
            strErrMsg = gdGetStr(hErrMsg, 0)
        End If
    End If
    gdDestroyString hErrMsg
    
    FromFile = nNumBars
End Function

' To write bars to a data file.
' - strFormat: "CSI", "MS7", "GEN" (Genesis EOD), "GT" (GenTick), "GDB" (Genesis Ascii)
' - strPath: path where data file is located (or filename for GDB)
' - strSymbol: pass either the symbol (leave blank to use the current bars.Symbol)
'      or pass the name of the data file to load (e.g. "F104.DTA")
' - strDesc: if strSymbol and strDesc are blank, will use current bars.Desc
' - nConvFactor: if < -5, will use current bars.ConvFactor
' - nAddMode: 0=overwrite file, 1=add to existing file, 2=add + auto-detect conv_factor
' - returns # of bars written, or negative number for error
Public Function ToFile(ByVal strFormat$, ByVal strPath$, _
        Optional ByVal strSymbol$ = "", Optional ByVal strDesc$ = "", _
        Optional ByVal nConvFactor& = -99, Optional strErrMsg$, _
        Optional ByVal nAddMode& = 0) As Long
    
    Dim nNumBars&, hErrMsg&
    
    hErrMsg = gdCreateString(0)
    
    nNumBars = gdBarsToFile(m.hBars, strFormat, strPath, strSymbol, _
        strDesc, nConvFactor, nAddMode, hErrMsg)
    
    If Not IsMissing(strErrMsg) Then
        If nNumBars >= 0 Then
            strErrMsg = ""
        Else
            strErrMsg = gdGetStr(hErrMsg, 0)
        End If
    End If
    gdDestroyString hErrMsg
    
    ToFile = nNumBars
End Function

Public Property Get ArrayMask() As eBarsArray
    If m.hBars = 0 Then
        ArrayMask = 0 'eBARS_Eod
    Else
        ArrayMask = gdGetBarsNumProp(m.hBars, eBARS_ArrayMask)
    End If
End Property

Public Property Let ArrayMask(ByVal nArrayMask As eBarsArray)
    If m.hBars = 0 Then
        m.hBars = gdCreateBars(0, nArrayMask)
    ElseIf nArrayMask <> ArrayMask Then
        gdSetBarsNumProp m.hBars, eBARS_ArrayMask, nArrayMask
    End If
End Property

' To build bars (e.g. Daily, Weekly, Monthly, Quarterly, Yearly)
' - strPeriod: as a string (optional number then letter/word)
'   - "Daily", "Day", "D"  =  Daily  (same for Weekly, Monthly, etc.)
'   - "2 days", "2d"  =  2 days per bar (same for weeks per bar, etc.)
'   - "30 minute", "30m", "30"  =  30 minute bars
'   - "100 ticks", "100t"  =  100 ticks per bar
'   - "500 vol", "500v"  =  1000 intraday volume per bar
'   - "500k vol", "500kv"  =  500000 end-of-day volume per bar
' - hFromBars: if 0, will use current Bars as source
' - bAppendFromBars: to append FromBars to source bars
' - hSplitDates: can pass array of split dates (gdArray Longs),
'       for unsplit data to separate weekly/monthly bars at the splits
' - returns success
Public Function BuildBars(ByVal strPeriod$, _
        Optional ByVal hFromBars& = 0, _
        Optional ByVal bAppendFromBars As Boolean = False, _
        Optional ByVal hSplitDates& = 0) As Boolean
    BuildBars = gdBuildBars2(m.hBars, strPeriod, hFromBars, bAppendFromBars, hSplitDates)
End Function

' TEMPORARY?: to build "Steidlmayer Market Profile" bars
Public Function BuildSMPBars(FromMinuteBars As cGdBars, Optional ByVal iMinimumPeriods& = 12, Optional ByVal iBreakoutPeriods& = 9) As Boolean

    Dim iMPB&, strPeriod$
    
    If iMinimumPeriods > 0 And iMinimumPeriods <= 99 And iBreakoutPeriods > 0 And iBreakoutPeriods <= iMinimumPeriods Then
        iMPB = FromMinuteBars.Prop(eBARS_PeriodsPerBar)
        If iMPB > 0 And FromMinuteBars.Size > 0 And FromMinuteBars.Prop(eBARS_PeriodType) = ePRD_Minutes Then
            strPeriod = Str(iMinimumPeriods) & "/" & Str(iBreakoutPeriods) & "/" & Str(iMPB) & "smp"
            BuildSMPBars = BuildBars(strPeriod, FromMinuteBars.BarsHandle)
        End If
    End If

End Function

' To append blank bars into the future (based on periodicity)
Public Sub AddForecastBars(ByVal nNumForecastBars&)
    gdAddForecastBars m.hBars, nNumForecastBars, ByVal 0&
End Sub

Public Property Get Prop(ByVal ePropID As eBarsPropID) As Variant
    If ePropID >= 100 Then
        ' get string
        Dim hString&
        hString = gdGetBarsStrProp(m.hBars, ePropID)
        Prop = gdGetStr(hString)
        gdDestroyString hString
    Else
        ' get number
        Prop = gdGetBarsNumProp(m.hBars, ePropID)
    End If
End Property

Public Property Let Prop(ByVal ePropID As eBarsPropID, ByVal vNewValue As Variant)
    If ePropID >= 100 Then
        ' set string
        gdSetBarsStrProp m.hBars, ePropID, Str(vNewValue)
    Else
        ' set number
        If VarType(vNewValue) = vbString Then
            gdSetBarsNumProp m.hBars, ePropID, Val(vNewValue)
        Else
            gdSetBarsNumProp m.hBars, ePropID, CDbl(vNewValue)
        End If
    End If
End Property

'To format a price (either as decimal or as trading units)
Public Function PriceDisplay(ByVal dPrice#, Optional ByVal bAsTradingUnits As Boolean = True, _
                        Optional ByVal nForSessionDate As Long = 0)
    Dim hString&, nFlags&
    If bAsTradingUnits Then nFlags = 1
    hString = gdBarsPriceDisplay2(m.hBars, dPrice, nFlags, nForSessionDate)
    PriceDisplay = gdGetStr(hString)
    gdDestroyString hString
End Function

'Returns a "copy" of these bars.
'(Note: array data not actually copied yet -- both
' bars are pointing to same array data until one of them
' is edited, then that one gets its own copy of the data)
' TLB 5/12/2008: if bOfZeroSize = True, then a separate copy of array data is made of 0 size
Public Function MakeCopy(Optional ByVal bOfZeroSize As Boolean = False) As cGdBars
    'first create a new set of bars
    Dim bSharedDataMode As Boolean
    Dim aCopy As cGdBars
    Set aCopy = New cGdBars
    'then make it a "copy" of this one
    gdCopy aCopy.BarsHandle, m.hBars
    If bOfZeroSize Then
        'set size of copy to 0 (but first make sure SharedDataMode is set to False)
        bSharedDataMode = SharedDataMode
        If bSharedDataMode Then SharedDataMode = False
        aCopy.Size = 0 '(forces the new array copies to be made while SharedDataMode is off)
        If bSharedDataMode Then SharedDataMode = True
    End If
    
    aCopy.PriceThresholds = PriceThresholds
    aCopy.SecondaryMinMoves = SecondaryMinMoves
    
    Set MakeCopy = aCopy
End Function

'Makes current object a copy of the handle being passed.
Public Function CopyFromHandle(ByVal CopyFromBarsHandle As Long) As Boolean
    
    Destroy 'destroy current bars if exists
    
    'make copy of bars
    m.hBars = gdCreateBars(0, eBARS_Eod) '(must start with a valid object)
    If gdCopy(m.hBars, CopyFromBarsHandle) <> 0 Then
        CopyFromHandle = True
    End If
    m.bAutoDestroy = True
    
End Function

' Returns true if the bar period is intraday (i.e. < Daily)
Public Property Get IsIntraday() As Boolean
    If Prop(eBARS_Periodicity) < ePRD_Days Then
        IsIntraday = True
    Else
        IsIntraday = False
    End If
End Property

' Returns true if the specified array is active.
Public Property Get IsActiveArray(ByVal WhichArray As eBarsArray) As Boolean
    If gdIsConstantValue(gdBarsArray(m.hBars, WhichArray)) = 0 Then
        IsActiveArray = True
    End If
End Property

' Rounds a value to the nearest even "price"
' (using the symbol's TickMove and MinMoveInTicks)
Public Function RoundToPrice(ByVal dPrice#, Optional ByVal dForDate# = 0) As Double

    Dim dMinMove#
    dMinMove = gdGetBarsMinMove(m.hBars, dForDate)
    RoundToPrice = RoundToMinMove(dPrice, dMinMove)
    
End Function

' Converts a price to a double from a string (e.g. from "101^31.5")
Public Function PriceFromString(ByVal strPrice As String) As Double

    PriceFromString = gdBarsPriceFromString(m.hBars, strPrice)
    
End Function

' To round all prices to minimum price increments (using TickMove and MinMoveInTicks)
Public Sub FixPrices(Optional ByVal dPriceAdjust As Double = 0)
    If m.hBars Then gdBarsFixPrices2 m.hBars, 0, dPriceAdjust
End Sub

' To delete one or more of the bars at the beginning.
Public Sub DeleteFirstBars(ByVal nCount&)
    If m.hBars Then
        gdDeleteFirstBars m.hBars, nCount
        FreeExtra
    End If
End Sub

' To delete one or more of the bars
Public Sub DeleteSomeBars(ByVal nFromBar&, Optional ByVal nCount& = 1)
    If m.hBars Then
        gdDeleteSomeBars m.hBars, nFromBar, nCount
        FreeExtra
    End If
End Sub

' Performs a binary search on the DateTime array.
' - returns bar number where the date matches or would be inserted
'      (i.e. bar number with the lowest date >= dDateTimeToFind)
' - returns -1 if bExactMatch = true and no exact match is found
Public Function FindDateTime(ByVal dDateTimeToFind As Double, Optional ByVal bExactMatch As Boolean = False) As Long

    FindDateTime = gdBarsFindDateTime(m.hBars, dDateTimeToFind, bExactMatch)

End Function

' Returns Session Date for the specified bar (tomorrow if time is after crossover)
' But when asked to validate, it will return 0 if the date/time is invalid:
' - if the trading session is a Sat or Sun
' - if it's before the custom start time (when start time different from default start)
' - if it's after the custom end time (when end time different from default end)
Public Function SessionDate(ByVal nBar As Long, Optional ByVal bValidate As Boolean = False) As Long

    SessionDate = SessionDateForTime(Item(eBARS_DateTime, nBar), bValidate)
        
End Function

' Returns Session Date for the specified date/time (tomorrow if time is after crossover)
' But when asked to validate, it will return 0 if the date/time is invalid:
' - if the trading session is a Sat or Sun
' - if it's before the custom start time (when start time different from default start)
' - if it's after the custom end time (when end time different from default end)
Public Function SessionDateForTime(ByVal dExchangeDateTime As Double, ByVal bValidate As Boolean) As Long

    If dExchangeDateTime > 0 Then
        SessionDateForTime = gdBarsSessionDate(m.hBars, dExchangeDateTime, bValidate)
    Else
        SessionDateForTime = 0
    End If

End Function

' This function should be used for "trade times" (e.g. fills)
Public Function SessionDateForTradeTime(ByVal dExchangeDateTime As Double) As Long

    Dim lCrossover As Long              ' Crossover time
    Dim lEndTime As Long
    Dim dReturn As Double               ' Return value from the function
    
    If dExchangeDateTime > 0 Then
        If Prop(eBARS_DefaultStartTime) = 0 And Prop(eBARS_DefaultEndTime) = 0 Then
            lEndTime = Prop(eBARS_EndTime)
        Else
            lEndTime = Prop(eBARS_DefaultEndTime)
        End If
        
        ' use a "crossover" that is not more than 5 minutes after default session end
        lCrossover = Prop(eBARS_CrossoverTime)
        If lCrossover = 0 Then lCrossover = 1440
        If lCrossover > lEndTime + 5 Then
            lCrossover = lEndTime + 5
        End If
        dReturn = dExchangeDateTime
        
        If lCrossover > 0 Then
            ' see if after crossover time
            If dReturn - Int(dReturn) >= gdFixDateTime(lCrossover / 1440#) Then
                dReturn = dReturn + 1
            End If
        End If
        dReturn = Int(dReturn)
        
        ' consider Friday night fills part of Monday's trading session
        Do While Not IsWeekday(dReturn)
            dReturn = dReturn + 1
        Loop
    End If
    SessionDateForTradeTime = Int(dReturn)

End Function

Public Function LastOpenSessionForTradeTime(ByVal dExchangeDateTime As Double) As Long

    Dim lStartTime As Long              ' Starting time for the session
    Dim lEndTime As Long                ' Endint time for the session
    Dim dReturn As Double               ' Return value from the function
    
    dReturn = 0
    If dExchangeDateTime > 0 Then
        dReturn = dExchangeDateTime
        
        If Prop(eBARS_DefaultStartTime) = 0 And Prop(eBARS_DefaultEndTime) = 0 Then
            lStartTime = Prop(eBARS_StartTime)
            lEndTime = Prop(eBARS_EndTime)
        Else
            lStartTime = Prop(eBARS_DefaultStartTime)
            lEndTime = Prop(eBARS_DefaultEndTime)
        End If
        
        If lStartTime < lEndTime Then
            If ((dReturn - Int(dReturn)) < gdFixDateTime(lStartTime / 1440#)) Or (IsWeekday(dReturn) = False) Then
                Do
                    dReturn = dReturn - 1&
                Loop While Not IsWeekday(dReturn)
            End If
        Else
            If Weekday(dReturn) = vbSaturday Then
                dReturn = dReturn - 1&
            ElseIf Weekday(dReturn) <> vbFriday Then
                If (dReturn - Int(dReturn)) >= gdFixDateTime(lStartTime / 1440#) Then
                    dReturn = dReturn + 1&
                ElseIf Weekday(dReturn) = vbSunday Then
                    dReturn = dReturn - 2&
                End If
            End If
        End If
    End If
    
    LastOpenSessionForTradeTime = Int(dReturn)
    
End Function

' Determines if the given time is valid for trading or not
Public Function IsValidTradingTime(ByVal dExchangeDateTime As Double, Optional ByVal dSessionSuspend = kNullData, Optional ByVal dSessionResume = kNullData) As Boolean
    
    Dim dSessionNow As Double           ' Time right now with session date
    Dim dSessionStart As Double         ' Time the session started
    Dim dSessionEnd As Double           ' Time the session ends
    Dim dCrossOver As Double            ' Crossover time
    Dim bReturn As Boolean              ' Return value for the function
    
    ' use default session start/end times (if exist)
    dSessionStart = Prop(eBARS_DefaultStartTime)
    dSessionEnd = Prop(eBARS_DefaultEndTime)
    If dSessionStart = 0 And dSessionEnd = 0 Then
        dSessionStart = Prop(eBARS_StartTime)
        dSessionEnd = Prop(eBARS_EndTime)
    End If
    
    ' use a "crossover" that is not more than 5 minutes after default session end
    dCrossOver = Prop(eBARS_CrossoverTime)
    If dCrossOver = 0 Then dCrossOver = 1440
    If dCrossOver > dSessionEnd + 5 Then
        dCrossOver = dSessionEnd + 5
    End If

    dSessionNow = dExchangeDateTime
    dCrossOver = Int(dSessionNow) + (dCrossOver / 1440#)
    If dSessionNow > dCrossOver Then dSessionNow = dSessionNow + 1
    
    dSessionStart = Int(dSessionNow) + (dSessionStart / 1440#)
    dSessionEnd = Int(dSessionNow) + (dSessionEnd / 1440#)
    If dSessionStart >= dSessionEnd Then dSessionStart = dSessionStart - 1

    bReturn = (IsWeekday(dSessionNow) And (dExchangeDateTime >= dSessionStart) And (dExchangeDateTime <= dSessionEnd))
    If bReturn Then
        If dSessionSuspend = kNullData Then
            dSessionSuspend = Prop(eBARS_SuspendTime)
        End If
        If dSessionResume = kNullData Then
            dSessionResume = Prop(eBARS_ResumeTime)
        End If
        
        If (dSessionSuspend <> 0) Or (dSessionResume <> 0) Then
            dSessionSuspend = Int(dSessionNow) + (dSessionSuspend / 1440#)
            If dSessionSuspend >= dSessionEnd Then dSessionSuspend = dSessionSuspend - 1
            
            dSessionResume = Int(dSessionNow) + (dSessionResume / 1440#)
            If dSessionResume >= dSessionEnd Then dSessionResume = dSessionResume - 1
            
            bReturn = ((dExchangeDateTime <= dSessionSuspend) Or (dExchangeDateTime >= dSessionResume))
        End If
    End If
    
    IsValidTradingTime = bReturn

End Function

Public Property Get SecurityType() As String
    Dim i&
    i = gdGetBarsNumProp(m.hBars, eBARS_SecurityType)
    If i > 0 Then SecurityType = Chr(i)
End Property

' This only works for external symbols (i.e. if symbol starts with '*')
Public Property Let SecurityType(ByVal strSecType As String)
    If Len(strSecType) = 0 Then
        gdSetBarsNumProp m.hBars, eBARS_SecurityType, 0
    Else
        gdSetBarsNumProp m.hBars, eBARS_SecurityType, Asc(strSecType)
    End If
End Property

' Returns true if symbol from an "external source" (e.g. CSI file, other data vendor, etc)
' - is external if symbol starts with '*'
Public Property Get IsExternalSymbol() As Boolean
    If gdGetBarsNumProp(m.hBars, eBARS_IsExternalSymbol) <> 0 Then
        IsExternalSymbol = True
    End If
End Property

' Returns DateTime of bar converted to the specified time zone:
' - "" (empty string) for machine's local time zone
' - or "GMT" for GMT/UTC, "NY" for New York, "CHI" for Chicago
' - or custom time zone info string (see "ConvertTimeZone" for format spec)
Public Function DateTimeConvert(ByVal nOffset&, Optional ByVal strToTimeZone$ = "") As Double
    DateTimeConvert = gdBarsDateTimeConvert(m.hBars, nOffset, strToTimeZone)
End Function

' Returns the Minimum Movement for this symbol
' - "min move" is basically = tick move * min move in ticks
' - will use date if passed (e.g. min move for SP used to be 0.05 but is now 0.1)
Public Function MinMove(Optional ByVal dForDate# = 0) As Double
    MinMove = gdGetBarsMinMove(m.hBars, dForDate)
End Function

' Returns the Tick Move for this symbol (price movement for a "logical tick")
Public Function TickMove() As Double
    TickMove = gdGetBarsNumProp(m.hBars, eBARS_TickMove)
End Function

' Returns the Tick Value for this symbol (dollar profit/loss when the market moves one tick)
Public Function TickValue() As Double
    TickValue = gdGetBarsNumProp(m.hBars, eBARS_TickValue)
End Function

' Returns the Tick Value divided by the Tick Move for profit calculations (zero if either is zero)
Public Function TickValuePerMove() As Double

    Dim dTickValue As Double            ' Tick value
    Dim dTickMove As Double             ' Tick move
    
    dTickValue = TickValue
    dTickMove = TickMove
    
    If (dTickValue <> 0#) And (dTickMove <> 0#) Then
        TickValuePerMove = dTickValue / dTickMove
    Else
        TickValuePerMove = 0#
    End If

End Function

' Returns a set of bars which is the Heikin-Ashi version of this bar's data
' (the OHLC data is modified)
Public Function CreateHeikinAshi() As cGdBars

    Dim b As cGdBars
    Set b = MakeCopy
    gdBarsConvertToHeikinAshi b.BarsHandle
    Set CreateHeikinAshi = b

End Function

' Returns the time of the last tick
Public Function LastTickDateTime() As Double

    Dim dLastTickTime As Double         ' Last tick time in fraction of day
    Dim dReturn As Double               ' Return value for the function
    
    ' Try to use the last tick time -- if it is zero, try to use the ending time of the session...
    dLastTickTime = Prop(eBARS_LastTickTime)
    If dLastTickTime = 0 Then
        dLastTickTime = Prop(eBARS_EndTime)
    End If
    
    dReturn = SessionDate(Size - 1) + (dLastTickTime / 1440#)
    If (dLastTickTime / 1440#) > (Prop(eBARS_CrossoverTime) / 1440#) Then
        dReturn = dReturn - 1&
    End If
    
    LastTickDateTime = dReturn

End Function

Public Property Get SharedDataMode() As Boolean
    SharedDataMode = gdGetShared(m.hBars)
End Property

Public Property Let SharedDataMode(ByVal bSharedData As Boolean)
    gdSetShared m.hBars, bSharedData
End Property

' returns true if custom session start/end time has been defined
Public Function HasCustomSessionTime() As Boolean

    Dim nDefaultStart&, nDefaultEnd&

    If IsIntraday Then
        nDefaultStart = Prop(eBARS_DefaultStartTime)
        nDefaultEnd = Prop(eBARS_DefaultEndTime)
        If nDefaultStart > 0 Or nDefaultEnd > 0 Then
            If Prop(eBARS_StartTime) <> nDefaultStart Or Prop(eBARS_EndTime) <> nDefaultEnd Then
                HasCustomSessionTime = True
            End If
        End If
    End If

End Function

Public Function SerializeDataArrays(ByVal vFile As Variant, ByVal bPut As Boolean) As Boolean

    Dim bSuccess As Boolean, hFile As Long
    
    If m.hBars <> 0 Then
        If VarType(vFile) <> vbString Then
            'file handle is being passed (file already open)
            hFile = CLng(vFile)
        ElseIf bPut Then
            'open file for binary output
            hFile = gdFileOpen(vFile, "wb")
        Else
            'open file for binary input
            hFile = gdFileOpen(vFile, "rb")
        End If
        If hFile <> 0 Then
            bSuccess = gdSerializeBarsArrays(m.hBars, hFile, bPut)
            'close file (if opened here)
            If VarType(vFile) = vbString Then
                gdFileClose hFile
            End If
        End If
    End If

    SerializeDataArrays = bSuccess
End Function

Public Property Get PriceThresholds() As String
    PriceThresholds = m.strPriceThresholds
End Property
Public Property Let PriceThresholds(ByVal strPriceThresholds As String)
    m.strPriceThresholds = strPriceThresholds
End Property

Public Property Get SecondaryMinMoves() As String
    SecondaryMinMoves = m.strSecondaryMinMoves
End Property
Public Property Let SecondaryMinMoves(ByVal strSecondaryMinMoves As String)
    m.strSecondaryMinMoves = strSecondaryMinMoves
End Property

' Note: the include header option is meant to support dumping data that TradeStation can chart
' (although we still have not yet figured out how to get the tick-by-tick data into TradeStation)
Public Function DumpToFile(ByVal strFileName$, Optional ByVal nBegAndEnd& = 0, Optional ByVal bIncludeHeader As Boolean = False) As Boolean

    Dim i&, fh%, strDateFmt$, strPriceFmt$, strBuffer$, strHeader$
    Dim nArrayMask As eBarsArray
    
    On Error GoTo ErrExit ' will simply return False if error creating file
    fh = FreeFile
    If InStr(strFileName, ".") = 0 Then
        strFileName = strFileName & Prop(eBARS_Symbol) & ".txt"
    End If
    Open strFileName For Output As #fh
    
    nArrayMask = ArrayMask
    If Not IsIntraday Then
        strDateFmt = "YYYYMMDD"
    ElseIf GetPeriodType(Prop(eBARS_Periodicity)) = ePRD_Minutes Then
        strDateFmt = "YYYYMMDD HHMM"
    Else
        strDateFmt = "YYYYMMDD HHMMSS"
    End If
    strPriceFmt = "#0.00##"
    
    If bIncludeHeader Then
        strDateFmt = Replace(strDateFmt, " ", vbTab)
    End If
    
    If nBegAndEnd = 0 Then nBegAndEnd = Size
    For i = 0 To Size - 1
        If (i < nBegAndEnd Or i >= Size - nBegAndEnd) And Item(eBARS_DateTime, i) > 0 Then
            If nArrayMask = eBARS_TickByTick Then
                'eBARS_TickByTick = (eBARS_DateTime Or eBARS_Close Or eBARS_Vol Or eBARS_Flags)
                If bIncludeHeader Then
                    strHeader = "Date,Time,Close,Volume,Other"
                End If
                strBuffer = Format(Item(eBARS_DateTime, i), strDateFmt) & vbTab & Format(Item(eBARS_Close, i), strPriceFmt) _
                & vbTab & Format(Item(eBARS_Vol, i), "0") & vbTab & Format(Item(eBARS_Flags, i), "0")
            ElseIf nArrayMask = eBARS_Minutized Then
                'eBARS_Minutized = (eBARS_DateTime Or eBARS_Close Or eBARS_Vol Or eBARS_UpTicks Or eBARS_DownTicks Or eBARS_BidVol Or eBARS_AskVol)
                If bIncludeHeader Then
                    strHeader = "Date,Time,Close,Volume,UpTicks,DownTicks,BidVol,AskVol"
                End If
                strBuffer = Format(Item(eBARS_DateTime, i), strDateFmt) & vbTab _
                    & Format(Item(eBARS_Close, i), strPriceFmt) & vbTab & Format(Item(eBARS_Vol, i), "0") & vbTab _
                    & Format(Item(eBARS_UpTicks, i), "0") & vbTab & Format(Item(eBARS_DownTicks, i), "0") & vbTab _
                    & Format(Item(eBARS_BidVol, i), "0") & vbTab & Format(Item(eBARS_AskVol, i), "0")
            ElseIf nArrayMask = eBARS_Intraday Then
                'eBARS_Intraday = (eBARS_Prices Or eBARS_Vol Or eBARS_UpTicks Or eBARS_DownTicks Or eBARS_BidVol Or eBARS_AskVol)
                If bIncludeHeader Then
                    strHeader = "Date,Time,Open,High,Low,Close,Volume,UpTicks,DownTicks,BidVol,AskVol"
                End If
                strBuffer = Format(Item(eBARS_DateTime, i), strDateFmt) & vbTab & Format(Item(eBARS_Open, i), strPriceFmt) & vbTab _
                    & Format(Item(eBARS_High, i), strPriceFmt) & vbTab & Format(Item(eBARS_Low, i), strPriceFmt) & vbTab _
                    & Format(Item(eBARS_Close, i), strPriceFmt) & vbTab & Format(Item(eBARS_Vol, i), "0") & vbTab _
                    & Format(Item(eBARS_UpTicks, i), "0") & vbTab & Format(Item(eBARS_DownTicks, i), "0") & vbTab _
                    & Format(Item(eBARS_BidVol, i), "0") & vbTab & Format(Item(eBARS_AskVol, i), "0")
            ElseIf (nArrayMask And eBARS_Bid) = 0 Then
                'eBars_Eod = (eBARS_Prices Or eBARS_VolOI)
                If bIncludeHeader Then
                    strHeader = "Date,Open,High,Low,Close,Volume,OpenInt"
                End If
                strBuffer = Format(Item(eBARS_DateTime, i), strDateFmt) & vbTab _
                    & Format(Item(eBARS_Open, i), strPriceFmt) & vbTab & Format(Item(eBARS_High, i), strPriceFmt) & vbTab _
                    & Format(Item(eBARS_Low, i), strPriceFmt) & vbTab & Format(Item(eBARS_Close, i), strPriceFmt) & vbTab _
                    & Format(Item(eBARS_Vol, i), "0") & vbTab & Format(Item(eBARS_OI, i), "0")
            Else 'eBARS_EodBidAsk = (eBars_Eod Or eBARS_BidAsk)
                If bIncludeHeader Then
                    strHeader = "Date,Open,High,Low,Close,BidPrice,BidSize,AskPrice,AskSize,Volume,OpenInt"
                End If
                strBuffer = Format(Item(eBARS_DateTime, i), strDateFmt) & vbTab _
                    & Format(Item(eBARS_Open, i), strPriceFmt) & vbTab & Format(Item(eBARS_High, i), strPriceFmt) & vbTab _
                    & Format(Item(eBARS_Low, i), strPriceFmt) & vbTab & Format(Item(eBARS_Close, i), strPriceFmt) & vbTab _
                    & Format(Item(eBARS_Bid, i), strPriceFmt) & vbTab & Format(Item(eBARS_BidSize, i), "0") & vbTab _
                    & Format(Item(eBARS_Ask, i), strPriceFmt) & vbTab & Format(Item(eBARS_AskSize, i), "0") & vbTab _
                    & Format(Item(eBARS_Vol, i), "0") & vbTab & Format(Item(eBARS_OI, i), "0")
            End If
            If bIncludeHeader Then
                strHeader = Replace(strHeader, ",", vbTab)
                Print #fh, strHeader
                bIncludeHeader = False ' so only does the header once
            End If
            Print #fh, strBuffer
        End If
    Next
    
    Close #fh
    DumpToFile = True

ErrExit: ' will simply return False if error creating file
    Exit Function
    
End Function

Public Property Get HighestHigh() As Double
    HighestHigh = gdMaxValue(ArrayHandle(eBARS_High), 0, Size - 1)
End Property
Public Property Get LowestLow() As Double
    LowestLow = gdMinValue(ArrayHandle(eBARS_Low), 0, Size - 1)
End Property

Public Function LastDayOfContractMonth() As Long

    Dim lReturn As Long                 ' Return value for the function
    Dim lContract As Long               ' Contract
    Dim lYear As Long                   ' Year of the contract
    Dim lMonth As Long                  ' Month of the contract
    Dim lDay As Long                    ' Last trading day of the month
    Dim strSecType As String            ' Security type
    
    lReturn = 0&
    strSecType = SecurityType
    If (strSecType = "F") Or (strSecType = "FO") Then
        lContract = Prop(eBARS_Contract)
        lYear = lContract \ 100
        
        lMonth = lContract Mod 100
        If Prop(eBARS_ExpiresPriorMonth) <> 0 Then
            If lMonth = 12 Then
                lMonth = 1
                lYear = lYear - 1
            Else
                lMonth = lMonth - 1
            End If
        End If
        
        Select Case lMonth
            Case 1, 3, 5, 7, 8, 10, 12
                lDay = 31
            Case 2
                lDay = 28
            Case 4, 6, 9, 11
                lDay = 30
        End Select
        
        lReturn = (lYear * 10000) + (lMonth * 100) + lDay
    End If
    
    LastDayOfContractMonth = lReturn

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NearestPriceDown
'' Description: Find the nearest valid price less than or equal to the price
''              passed in
'' Inputs:      Price
'' Returns:     Nearest Valid Price Down
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NearestPriceDown(ByVal dPrice As Double) As Double

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = RoundToPrice(dPrice)
    If dReturn > dPrice Then
        dReturn = dReturn - MinMove
    End If
    
    NearestPriceDown = dReturn

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NearestPriceUp
'' Description: Find the nearest valid price greater than or equal to the price
''              passed in
'' Inputs:      Price
'' Returns:     Nearest Valid Price Up
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NearestPriceUp(ByVal dPrice As Double) As Double

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = RoundToPrice(dPrice)
    If dReturn < dPrice Then
        dReturn = dReturn + MinMove
    End If
    
    NearestPriceUp = dReturn

End Function

Public Function GetHoliday(ByVal nRecord&) As Long

    Dim hHolidayTable&, d#

    hHolidayTable = gdGetBarsNumProp(m.hBars, 99)
    If hHolidayTable <> 0 Then
        d = gdGetTableNum(hHolidayTable, 0, nRecord)
        If d > 0 Then
            GetHoliday = d
        End If
    End If
    
End Function

Public Function NumHolidays() As Long

    Dim hHolidayTable&

    hHolidayTable = gdGetBarsNumProp(m.hBars, 99)
    If hHolidayTable <> 0 Then
        NumHolidays = gdGetSize(hHolidayTable)
    End If
    
End Function

Public Function IsHoliday(ByVal nDate&) As Boolean
    
    Dim hHolidayTable&, hDates&, iPos&

    If nDate > 0 Then
        hHolidayTable = gdGetBarsNumProp(m.hBars, 99)
        If gdGetSize(hHolidayTable) > 0 Then
            hDates = gdFieldArrayHandle(hHolidayTable, 0)
            If hDates <> 0 Then
                If gdBinarySearch(hDates, nDate, iPos, eGdSort_Default, 0, 999999) <> 0 Then
                    IsHoliday = True
                End If
            End If
        End If
    End If

End Function

Public Function NextTradingDay(ByVal lDate As Long) As Long

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = lDate + 1&
    Do While (IsWeekday(lReturn) = False) Or (IsHoliday(lReturn) = True)
        lReturn = lReturn + 1&
    Loop
    
    NextTradingDay = lReturn

End Function
