VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFractZen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    strIniFile As String        ' stores the default FractZen parms, and overrides for specific symbols
    strResultsFile As String    ' file for the calculated results
    aResults As cGdArray        ' to store the calculated breakout period for each symbol/date/parms combination
    bUseMerged As Boolean       ' if using merged FractZen bars
    SpeedInfo As cGdTree        ' cache to hold the "speed info" for each symbol (just while program is running)
    
    iLoadCount As Long
End Type

Private m As mPrivate

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.SpeedInfo = New cGdTree
    Set m.aResults = New cGdArray
    
    m.strIniFile = App.Path & "\Provided\FractZenTest.INI" ' override just for John's machine
    If Not FileExist(m.strIniFile) Then
        m.strIniFile = App.Path & "\Provided\FractZen.INI"
    End If
    
    m.strResultsFile = App.Path & "\Custom\FractZen.DAT"
    
    If Val(GetIniFileProperty("UseMerged", 1, "", m.strIniFile)) <> 0 Then
        m.bUseMerged = True
    End If
    LoadResults

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFractZen.Class_Initialize"
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    SaveResults
    Set m.aResults = Nothing
    Set m.SpeedInfo = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFractZen.Class_Terminate"
End Sub

Public Property Get UseMerged() As Boolean
    UseMerged = m.bUseMerged
End Property

Public Function Allowed() As Boolean
On Error GoTo ErrSection:
    
    Static iFlagFile%
    
    If iFlagFile = 0 Then
        If Trim(UCase(FileToString(App.Path & "\AutoBreakout.flg", , True))) = "PROJECTX" Then
            iFlagFile = 1
        Else
            iFlagFile = -1
        End If
    End If
    
    ' return True if has enablement code or if the special flag file exists
    If iFlagFile > 0 Or HasModule("DCABO") Then
        Allowed = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFractZen.Allowed"
End Function

Public Function AllowTSOG() As Boolean
On Error GoTo ErrSection:
    
    AllowTSOG = Me.Allowed And HasModule("TSOGPRJX")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFractZen.AllowTSOG"
End Function

Private Sub LoadResults()
On Error GoTo ErrSection:

    Dim i&, nDate&, s$
    Static nPrevDailyDownload As Long
    
    If Allowed Then
        If m.aResults.Size = 0 Then
            ' Symbol, MaxBarsPerDay, LookbackDays, Date, BreakoutTickRange
            m.aResults.FromFile m.strResultsFile
            m.aResults.Sort eGdSort_DeleteNullValues Or eGdSort_DeleteDuplicates
        End If
        
        ' after a new daily download, remove any records >= PrevDailyDownload (in case of recent data fixes)
        If nPrevDailyDownload > 0 And nPrevDailyDownload <> LastDailyDownload Then
            For i = m.aResults.Size - 1 To 0 Step -1
                s = m.aResults(i)
                nDate = Val(Parse(s, vbTab, 4))
                If nDate >= nPrevDailyDownload Then
                    m.aResults.Remove i
                End If
            Next
            SaveResults ' re-save after a daily download
            ' and clear the SpeedInfo cache (in case of updated FractZen.INI settings)
            m.SpeedInfo.Clear
        End If
        nPrevDailyDownload = LastDailyDownload
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFractZen.LoadResults"
End Sub

Private Sub SaveResults()
On Error GoTo ErrSection:

    If Allowed And m.aResults.Size > 0 Then
        m.aResults.ToFile m.strResultsFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFractZen.SaveResults"
End Sub

Public Function FzSymbol(ByVal strSymbol$) As String
On Error GoTo ErrSection:

    Dim s$

    ' if first char = Esc, then was already converted (no need to do multiple times)
    If Left(strSymbol, 1) = Chr(27) Then
        strSymbol = Mid(strSymbol, 2)
    Else
        Select Case SecurityType(strSymbol)
        Case "F"
            ' convert futures symbols to Electronic-067
            s = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
            If Len(s) > 0 Then
                strSymbol = s
            End If
            strSymbol = Parse(strSymbol, "-", 1) & "-067"
        Case "I"
            If IsForex(strSymbol) Then
                ' convert broker-specific forex to generic forex symbol
                strSymbol = Parse(strSymbol, "@", 1)
            End If
        End Select
    End If
    
    FzSymbol = strSymbol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFractZen.FzSymbol"
End Function

' Returns the speed info for this symbol (as a comma-delimited string)
' e.g. "8,3,10" means MaxBarsPerDay = 8, LookbackDays = 3, MinTicks = 10
Public Function GetSpeedInfo(ByVal strSymbol$) As String
On Error GoTo ErrSection:

    Dim nMaxBarsPerDay&, nLookbackDays&, nMinTicks&
    Dim strDefSpeed$, strSymSpeed$, strSpeed$
    Static dIniFileDate As Double, strSpeedSection$
    
    ' if INI file changed, then clear cache so will reload
    If FileDate(m.strIniFile) <> dIniFileDate Then
        dIniFileDate = FileDate(m.strIniFile)
        m.SpeedInfo.Clear
        
        strSpeedSection = UCase(FileToString(m.strIniFile))
        If InStr(strSpeedSection, "[SPEED]") > 0 Then
            strSpeedSection = "Speed" ' new
        ElseIf InStr(strSpeedSection, "[SPEEDS]") > 0 Then
            strSpeedSection = "Speeds" ' new
        Else
            strSpeedSection = "Symbols" ' will become obsolete
        End If
    End If
    
    ' convert symbol
    strSymbol = FzSymbol(strSymbol)
    
    ' see if speed info for this symbol is already cached
    If Len(strSymbol) = 0 Or Not Me.Allowed Then
        strSpeed = ""
    ElseIf m.SpeedInfo.Exists(strSymbol) Then
        strSpeed = m.SpeedInfo(strSymbol)
    Else
        ' from the INI file, get the "Symbol=MaxBarsPerDay,LookbackDays" setting
        strSymSpeed = Trim(GetIniFileProperty(strSymbol, "", strSpeedSection, m.strIniFile))
        
        ' get default for security type
        If SecurityType(strSymbol) = "S" Then
            strDefSpeed = Trim(GetIniFileProperty("Stocks", "", "Defaults", m.strIniFile))
        ElseIf IsForex(strSymbol) Then
            strDefSpeed = Trim(GetIniFileProperty("Forex", "", "Defaults", m.strIniFile))
        End If
        If Len(strDefSpeed) = 0 Then
            strDefSpeed = Trim(GetIniFileProperty("Default", "", "Defaults", m.strIniFile))
        End If
        
        ' get MaxBarsPerDay (unless was passed in as an override)
        If nMaxBarsPerDay <= 0 Then
            If Len(strSymSpeed) > 0 Then
                nMaxBarsPerDay = Val(Parse(strSymSpeed, ",", 1))
            End If
            If nMaxBarsPerDay <= 0 Then
                nMaxBarsPerDay = Val(Parse(strDefSpeed, ",", 1))
            End If
            If nMaxBarsPerDay <= 0 Then
                nMaxBarsPerDay = 8
            End If
        End If
        
        ' get LookbackDays (unless was passed in as an override)
        If nLookbackDays <= 0 Then
            If Len(strSymSpeed) > 0 Then
                nLookbackDays = Val(Parse(strSymSpeed, ",", 2))
            End If
            If nLookbackDays <= 0 Then
                nLookbackDays = Val(Parse(strDefSpeed, ",", 2))
            End If
            If nLookbackDays <= 0 Then
                nLookbackDays = 3
            End If
        End If
        
        ' get MinTicks (unless was passed in as an override)
        If nMinTicks <= 0 Then
            If Len(strSymSpeed) > 0 Then
                nMinTicks = Val(Parse(strSymSpeed, ",", 3))
            End If
            If nMinTicks <= 0 Then
                nMinTicks = Val(Parse(strDefSpeed, ",", 3))
            End If
            If nMinTicks <= 0 Then
                nMinTicks = 10
            End If
        End If
        
        ' cache the speed info for this symbol
        strSpeed = Str(nMaxBarsPerDay) & "," & Str(nLookbackDays) & "," & Str(nMinTicks)
        m.SpeedInfo(strSymbol) = strSpeed
    End If
        
ErrExit:
    GetSpeedInfo = strSpeed
    Exit Function
    
ErrSection:
    RaiseError "cFractZen.GetSpeedInfo"
End Function

' Setting speed info is really ONLY for testing purposes (by John Needham and Terry).
' It is only in the cache, and just meant to be very temporary (e.g. TradeNav restarts, or new daily download).
' And can just pass all 0's to clear the setting back to what's in the INI file.
Public Sub SetSpeedInfo(ByVal strSymbol$, Optional ByVal nMaxBarsPerDay& = 0, Optional ByVal nLookbackDays& = 0, Optional ByVal nMinTicks& = 0)
On Error GoTo ErrSection:

    Dim strSpeed$
    
    ' convert symbol
    strSymbol = FzSymbol(strSymbol)
    If Len(strSymbol) > 0 And Me.Allowed Then
        If nMaxBarsPerDay > 0 Or nLookbackDays > 0 Or nMinTicks > 0 Then
            ' get default speed info
            strSpeed = GetSpeedInfo(strSymbol)
            If nMaxBarsPerDay <= 0 Then
                nMaxBarsPerDay = Val(Parse(strSpeed, ",", 1))
            End If
            If nLookbackDays <= 0 Then
                nLookbackDays = Val(Parse(strSpeed, ",", 2))
            End If
            If nMinTicks <= 0 Then
                nMinTicks = Val(Parse(strSpeed, ",", 3))
            End If
            
            ' cache the speed info for this symbol
            strSpeed = Str(nMaxBarsPerDay) & "," & Str(nLookbackDays) & "," & Str(nMinTicks)
            m.SpeedInfo(strSymbol) = strSpeed
        ElseIf m.SpeedInfo.Exists(strSymbol) Then
            ' just clear the custom setting
            m.SpeedInfo.Remove strSymbol
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFractZen.SetSpeedInfo"
End Sub

' Returns the FractZen period as a string (e.g. "25 tick breakout")
Public Function GetFractZenPeriod(ByVal strSymbol$, Optional ByVal nForSessionDate& = 0) As String
On Error GoTo ErrSection:

    Dim nTicks&
    nTicks = GetFractZenRange(strSymbol, nForSessionDate)
    If nTicks > 0 Then
        GetFractZenPeriod = GetPeriodStr(ePRD_IntBreakout + nTicks)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFractZen.GetFractZenPeriod"
End Function

' TLB: returns the breakout range for John Needham (their DanielCode server)
' (the smallest breakout bar size such that it's not over the specified max bars per day during the lookback days)
Public Function GetFractZenRange(ByVal strSymbol$, Optional ByVal nForSessionDate& = 0) As Long

On Error GoTo ErrSection:

    Dim i&, nTicks&, nEndDate&, nDate&, nMaxBarsPerDay&, nLookbackDays&, nMinTicks&
    Dim strCalcKey$, strSpeed$
    Dim DailyBars As New cGdBars
    
    If Not Allowed Then Exit Function
    
    ' make sure table has been loaded (this also runs after a daily download in order to clear recent dates)
    LoadResults
    
    ' convert symbol
    strSymbol = FzSymbol(strSymbol)
    If Len(strSymbol) = 0 Then Exit Function
    
    ' speed parms for this symbol
    strSpeed = GetSpeedInfo(Chr(27) & strSymbol) ', nMaxBarsPerDay, nLookbackDays, nMinTicks
    nMaxBarsPerDay = Val(Parse(strSpeed, ",", 1))
    nLookbackDays = Val(Parse(strSpeed, ",", 2))
    nMinTicks = Val(Parse(strSpeed, ",", 3))
       
    nTicks = 0
    If nForSessionDate > 0 Then
        ' for quick lookup, just assume for the moment that the last trading session was the prior weekday
        nEndDate = nForSessionDate - 1
        Do While Not IsWeekday(nEndDate)
            nEndDate = nEndDate - 1
        Loop
        
        ' see if this specific result has already been calculated
        strCalcKey = strSymbol & vbTab & Str(nMaxBarsPerDay) & vbTab & Str(nLookbackDays) _
                        & vbTab & Str(nEndDate) & vbTab
        If m.aResults.BinarySearch(strCalcKey, i, eGdSort_MatchUsingSearchStringLength) Then
            ' if already calculated, just return the result
            nTicks = Val(Parse(m.aResults(i), vbTab, 5))
        End If
    End If
    
    If nTicks <= 0 Then
        ' else use Daily bars to get the Start and End session dates
        If nForSessionDate > 0 Then 'And nEndingSessionDate <= LastDailyDownload Then
            DM_GetBars DailyBars, strSymbol, "Daily", nForSessionDate - nLookbackDays * 2 - 10, nForSessionDate - 1
        Else
            DM_GetBars DailyBars, strSymbol, "Daily", nForSessionDate - nLookbackDays * 2 - 10
            g.RealTime.SpliceBars DailyBars
            ' End date is the last completed session
            ' (Note: the last completed session could still be after the LastDailyDownload --
            '  e.g. if tomorrow's session has already started then today would be completed)
            If DailyBars.SessionDate(DailyBars.Size - 1) > LastDailyDownload Then
                DailyBars.Size = DailyBars.Size - 1 ' throw away the current session in progress
            End If
        End If
        nEndDate = DailyBars.SessionDate(DailyBars.Size - 1)
        If DailyBars.Size > 0 And nEndDate > 0 Then
            ' see if this specific result has already been calculated
            strCalcKey = strSymbol & vbTab & Str(nMaxBarsPerDay) & vbTab & Str(nLookbackDays) _
                            & vbTab & Str(nEndDate) & vbTab
            If m.aResults.BinarySearch(strCalcKey, i, eGdSort_MatchUsingSearchStringLength) Then
                ' if already calculated, just return the result
                nTicks = Val(Parse(m.aResults(i), vbTab, 5))
            End If
            If nTicks <= 0 Then
                ' if not already calculated, then must do the calculation for this symbol/date/parms
                nTicks = CalcFractZenRange(strSymbol, DailyBars, nMaxBarsPerDay, nLookbackDays)
                If nTicks > 0 Then
                    ' now store the result (so won't have to keep recalculating)
                    If m.aResults.BinarySearch(strCalcKey, i, eGdSort_MatchUsingSearchStringLength) Then
                        m.aResults(i) = strCalcKey & Str(nTicks)
                    Else ' (insert it into correct position of the sorted array to maintain the sortation)
                        m.aResults.Add strCalcKey & Str(nTicks), i
                    End If
                End If
            End If
        End If
    End If
    
    ' can just apply various MinTicks now (after the symbol/date/parms calculation)
    If nTicks <= 0 Then
        nTicks = 0
    ElseIf nTicks < nMinTicks Then
        nTicks = nMinTicks
    End If
    GetFractZenRange = nTicks

ErrExit:
    Set DailyBars = Nothing
    Exit Function
    
ErrSection:
    RaiseError "cFractZen.GetFractZenRange"
End Function


' TLB: to calculate the breakout range for John Needham (their DanielCode server)
' (returns the smallest breakout bar size such that it's not over the specified max bars per day during the lookback days)
Private Function CalcFractZenRange(ByVal strSymbol$, DailyBars As cGdBars, ByVal nMaxBarsPerDay&, ByVal nLookbackDays&) As Long
On Error GoTo ErrSection:

    Dim i&, nTicks&, nTickAdjust&, dMaxRange#, nMaxRangeDate&, nDayCount&
    Dim nSessionDate&, nDate&, nStartDate&, nEndDate&
    Dim bStillDoubling As Boolean
    Dim Bars As New cGdBars, TickBars As cGdBars
    Dim aMBPD As New cGdArray ' to store the MaxBarsPerDay for various breakout bar sizes
    
    Static TickBarsCache As cGdTree ' collection of cached TickBars (for most recent lookback period)
    Static strPrevSymbol$, nPrevLDD&
    
    ' clear the TickBarsCache if Symbol or LDD has changed
    If strSymbol <> strPrevSymbol Or LastDailyDownload <> nPrevLDD Or TickBarsCache Is Nothing Then
        Set TickBarsCache = New cGdTree
        strPrevSymbol = strSymbol
        nPrevLDD = LastDailyDownload
    End If
    
    ' Start date is the session date for nDays back (from last completed session)
    nEndDate = DailyBars.SessionDate(DailyBars.Size - 1)
    nDayCount = 0
    For i = DailyBars.Size - 1 To 0 Step -1
        If DailyBars(eBARS_High, i) - DailyBars(eBARS_Low, i) > dMaxRange Then
            dMaxRange = DailyBars(eBARS_High, i) - DailyBars(eBARS_Low, i)
            nMaxRangeDate = DailyBars.SessionDate(i)
        End If
        nStartDate = DailyBars.SessionDate(i)
        nDayCount = nDayCount + 1
        If nDayCount >= nLookbackDays Then
            Exit For
        End If
    Next
    
    ' just wipe out anything in the cache outside of this lookback period for this symbol
    For i = TickBarsCache.Count To 1 Step -1
        nSessionDate = Val(TickBarsCache.Key(i))
        If nSessionDate < nStartDate Or nSessionDate > nEndDate Then
            TickBarsCache.Remove i
        End If
    Next
    
    ' To calculate: must load ticks for each session and then search for the correct size
    SetBarProperties Bars, strSymbol
    aMBPD.Create eGDARRAY_Longs, 150
    For nDate = nStartDate - 1 To nEndDate
        Set TickBars = Nothing
        nSessionDate = 0
        If nDate < nStartDate Then
            nSessionDate = nMaxRangeDate ' do date of max range first (a little more efficient)
        ElseIf IsWeekday(nDate) And nDate <> nMaxRangeDate Then
            nSessionDate = nDate
        End If
        If nSessionDate > 0 Then
            ' first see if the ticks for this day is already in the cache
            Set TickBars = TickBarsCache(Str(nSessionDate))
            If TickBars Is Nothing Then
                ' if not in the cache, then load ticks for the trading session
                Set TickBars = New cGdBars
                DM_GetBars TickBars, strSymbol, "each tick", nSessionDate, nSessionDate
                If TickBars.Size > 0 Then
                    ' add this to the local TickBarsCache
                    TickBarsCache.Add TickBars, Str(nSessionDate)
                    m.iLoadCount = m.iLoadCount + 1
                End If
            End If
            If nSessionDate > LastDailyDownload Then
                g.RealTime.SpliceBars TickBars, nSessionDate
            End If
            ' remove any trailing ticks for next session (e.g. if added from SpliceBars)
            For i = TickBars.Size - 1 To 0 Step -1
                If TickBars.SessionDate(i) = nSessionDate Then
                    TickBars.Size = i + 1
                    Exit For
                ElseIf i = 0 Then
                    TickBars.Size = 0
                End If
            Next
            If TickBars.Size = 0 Then
                Set TickBars = Nothing
            End If
        End If
        If Not TickBars Is Nothing Then
            Bars.Size = 0
            If nTicks > 0 Then
                Bars.BuildBars Str(nTicks) & "b", TickBars.BarsHandle
            End If
            If Bars.Size > nMaxBarsPerDay Or nTicks = 0 Then
                ' do a binary search to hone in on the breakout size to meet the criteria
                ' (but store the MaxBarsPerDay for each breakout size to avoid extra recalcs)
                aMBPD.Size = 0
                i = nTicks
                nTicks = 64 ' just make a reasonable guess as to a starting point
                Do While nTicks <= i ' but we know we need to start at least bigger than we were
                    nTicks = nTicks * 2
                Loop
                bStillDoubling = True
                Do While True
                    ' if we haven't yet calculated the max bars per day for this size of
                    ' breakout bars, then we need to calc it now (then store so only calc once)
                    If aMBPD(nTicks) <= 0 Then
                        Bars.Size = 0
                        Bars.BuildBars Str(nTicks) & "b", TickBars.BarsHandle
                        If Bars.Size = 0 Then
                            ' check if can't build breakout bars
                            If nSessionDate < nStartDate Then
                                nTicks = 0
                                GoTo ErrExit
                            End If
                            Exit Do
                        End If
                        aMBPD(nTicks) = Bars.Size
                    End If
                    
                    ' We are done when we get to a size where we don't have too many bars per day
                    ' but where the next smaller size does have too many bars per day
                    If aMBPD(nTicks) <= nMaxBarsPerDay And aMBPD(nTicks - 1) > nMaxBarsPerDay Then
                        Exit Do ' nTicks is the correct answer
                    End If
                    
                    ' if we don't have too many bars per day, try a smaller breakout size
                    If aMBPD(nTicks) <= nMaxBarsPerDay Then
                        If bStillDoubling Then
                            ' now that we're too high, we can stop doubling and start the binary search
                            bStillDoubling = False
                            nTickAdjust = Int(nTicks / 2)
                        End If
                        nTicks = nTicks - nTickAdjust
                    ElseIf bStillDoubling Then
                        nTicks = nTicks * 2
                    Else ' otherwise we need to try a bigger breakout size
                        nTicks = nTicks + nTickAdjust
                    End If
                    ' we can halve the increment/decrement amount each time through the binary search
                    If nTickAdjust >= 2 Then
                        nTickAdjust = Int(nTickAdjust / 2)
                    End If
                    If nTicks <= 1 Then
                        nTicks = 1
                        Exit Do
                    End If
                Loop
            End If
        End If
    Next
   
    DebugLog "CalcFractZen: " & strSymbol & " " & Format(nEndDate, "YYYYMMDD") & " " _
            & Str(nMaxBarsPerDay) & " " & Str(nLookbackDays) & " = " & Str(nTicks)
    CalcFractZenRange = nTicks

ErrExit:
    Set aMBPD = Nothing
    Set Bars = Nothing
    Set TickBars = Nothing
    Exit Function
    
ErrSection:
    RaiseError "cFractZen.CalcFractZenRange"
End Function

' Call this to set the Bars to FractZen
Public Function SetFractZen(Bars As cGdBars, Optional ByVal nForSessionDate& = 0) As Boolean
On Error GoTo ErrSection:

    Dim nTicks&, strSymbol$
        
    If Not Bars Is Nothing Then
        strSymbol = Bars.Prop(eBARS_Symbol)
    End If
    
    ' make sure enabled for FractZen
    If Me.Allowed And Len(strSymbol) > 0 Then
        ' get # ticks for the ending date
        nTicks = Me.GetFractZenRange(Bars.Prop(eBARS_Symbol), nForSessionDate)
        If nTicks > 0 Then
            ' clear the size (unless is already breakout bars and using merged)
            If Bars.Prop(eBARS_PeriodType) <> ePRD_IntBreakout Or Bars.Prop(eBARS_FractZen) <> Abs(Me.UseMerged) Then
                Bars.Size = 0
            End If
            ' set merge flag and current breakout bar range
            Bars.Prop(eBARS_FractZen) = Abs(Me.UseMerged)
            Bars.Prop(eBARS_Periodicity) = ePRD_IntBreakout + nTicks
            SetFractZen = True
        End If
    End If
    If nTicks <= 0 And Not Bars Is Nothing Then
        ' not allowed or invalid
        Bars.Size = 0
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFractZen.SetFractZen"
End Function

