VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSymbolInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The purpose of this structure is to provide a binary signature of the
' messages passed between the salmon client and trade navigator.
Private Type DailySummaryStructure
    nTradeCount As Long
    nBidAskCount As Long
    strSymbol As String * 32
    strExchange As String * 8
    strSecType As String * 8
    dOpen As Double
    dPrevOpen As Double
    dHigh As Double
    dPrevHigh As Double
    dLow As Double
    dPrevLow As Double
    dClose As Double
    dPrevClose As Double
    dVolume As Double
    dPrevVolume As Double
    dBidPrice As Double
    dAskPrice As Double
    dBidSize As Double
    dAskSize As Double
    dSettle As Double
    dPrevSettle As Double
    dLastTickTime As Double
    dBidTime As Double
    dAskTime As Double
    dSessionDate As Double
    dPrevSessionDate As Double
    nSymbolID As Long
    nGmtOffset As Long
    nDomestic As Long
    nInternalPointer As Long
    dPastGoodPrice As Double
    dMinMove As Double
    nLastTickDown As Long ' 0 = up, 1 = down
    nLastTickVol As Long
    nBrokerOverride As Long ' set if broker override is enabled for this symbol
    ' added 9/10/2010 (all these are only used internally by the salmon DLL)
    nAutoRefresh As Long    ' 1 if salmon client will automatically refresh
    nDelayAmount As Long    ' delay for this symbol?
    nSlotConsumed As Long   ' 1 if consumed, 0 if not
    dOpenInterest As Double ' primarily used for options (current session)
    dPrevOpenInterest As Double ' (previous session)
    ' used by TradeNav
    nNewsCount As Long      ' # of news items for this symbol
    nPrevNewsCount As Long
    nPadding(1 To 4) As Long ' future growth (keep 8-byte alignment)
End Type ' (320 bytes)

Private Type mPrivate
    DailySummary As DailySummaryStructure
    Bars As cGdBars
    NewTicks As cGdBars
    HolidayBar As cGdBars
    
    nSymbolID As Long
    strSymbol As String
    strSecType As String
    strExchange As String
    strRequestSettle As String
    bAddedToGenesisRT As Boolean

    bCanHaveMktDepth As Boolean
    bIsSpread As Boolean
End Type
Private m As mPrivate

' Get??? returns the SalmonResponse enum (1 = pending, 2 = available, 3 = error)
Private Declare Function GetEodData Lib "SalmonClient.dll" Alias "GetEOD" (ByVal hBars As Long, DailyInf As DailySummaryStructure) As Long
Private Declare Function GetMinuteBarData Lib "SalmonClient.dll" Alias "GetMinute" (ByVal hBars As Long, DailyInf As DailySummaryStructure) As Long
Private Declare Function GetFullTickData Lib "SalmonClient.dll" Alias "GetFull" (ByVal hBars As Long, DailyInf As DailySummaryStructure) As Long
Private Declare Function GetMinuteBarStatus Lib "SalmonClient.dll" Alias "GetMinuteStatus" (ByVal hBars As Long) As Long
Private Declare Function GetFullTickStatus Lib "SalmonClient.dll" Alias "GetFullStatus" (ByVal hBars As Long) As Long
' returns number of ticks since last called
Private Declare Function GetStreamedTicks Lib "SalmonClient.dll" Alias "GetStreamBars" (ByVal hBars As Long, DailyInf As DailySummaryStructure) As Long
' to remove a symbol from the stream (and invalidate its data)
Private Declare Function RemoveSymbol Lib "SalmonClient.dll" (DailyInf As DailySummaryStructure) As Long
' for depth-of-market
Private Declare Function RequestDOM Lib "SalmonClient.dll" (ByVal hBars As Long, ByVal hBidTable As Long, ByVal hAskTable As Long, nChangeCounter As Long) As Long
' for OptionNav symbol
Private Declare Function RequestEODForOptionNavigator Lib "SalmonClient.dll" (ByVal hBars As Long, DailyInf As DailySummaryStructure, ByVal hAddSymbols As Long, ByVal hBaseSymbols As Long) As Long


Private Sub Class_Initialize()

    'Set m.Bars = New cGdBars

End Sub

Private Sub Class_Terminate()

    If Len(m.strSymbol) > 0 And Not g.bUnloading Then
        If Not g.RealTime Is Nothing Then
            If g.RealTime.SalmonIsRunning Then
                '??? RemoveSymbol m.DailySummary
            End If
        End If
    End If

    Set m.Bars = Nothing
    Set m.NewTicks = Nothing
    Set m.HolidayBar = Nothing

End Sub

Public Sub SetSymbol(ByVal Symbol As Variant)
On Error GoTo ErrSection:
   
    Dim i&, strText$
    Dim TimeInfo As TradeTimeInfo
   
    ' This is designed to ONLY set the symbol ONCE -- i.e. it can
    ' be reset for the same symbol, but not to a different symbol.
    strText = RollSymbolForDate(Symbol)
    If Len(strText) > 0 Then
        If Len(m.strSymbol) = 0 Then
            m.strSymbol = strText
        ElseIf m.strSymbol <> strText Then
            DebugLog "cSymbolInfo ERROR: tried to change symbol from " & m.strSymbol & " to " & strText
            Exit Sub ' ERROR: cannot change the symbol
        End If
    End If
    
    If Len(m.strSymbol) > 0 Then
        If IsIDE Then
            If Len(m.DailySummary) <> 320 Then
                i = i / 0 ' STOP: there must be a problem!
            End If
'If Left(m.strSymbol, 3) = "TQ-" Or Left(m.strSymbol, 3) = "US-" Then
If m.strSymbol = "BAX" Then
    i = i
End If
        End If
        
        m.bAddedToGenesisRT = False
        m.nSymbolID = GetSymbolID(m.strSymbol)
        
        ' if a different settle symbol, then set flag so will request it first
        m.strRequestSettle = GetSettleSymbol(m.strSymbol)
        If m.strRequestSettle = m.strSymbol Then
            m.strRequestSettle = ""
        End If
        
        Set m.HolidayBar = Nothing
            
        ' get the last daily download's settle and store as "past good price"
        Set m.Bars = New cGdBars
        m.Bars.ArrayMask = eBARS_EodBidAsk
        DM_GetBars m.Bars, m.strSymbol, "Daily", LastDailyDownload - 7, 0, False, , False, False
        m.strSecType = SecurityType(m.strSymbol, True)
        
        ' check if it's a spread symbol (e.g. HE-201302-S1)
        m.bIsSpread = IsSpreadSymbol(m.strSymbol)
        
        ' keep just the last good data bar
        For i = m.Bars.Size - 1 To 0 Step -1
            If m.Bars(eBARS_Close, i) <> kNullData Then
                m.Bars.DeleteFirstBars i
                Exit For
            ElseIf i = 0 Then
                m.Bars.Size = 0
            End If
        Next
        If m.Bars.Size > 1 Then m.Bars.Size = 1
        
        With m.DailySummary
            ' allow SalmonDLL to use the PastGoodPrice (except for Options and certain Indices)
            .dPastGoodPrice = 0
            ' TLB: need to ignore "spread" symbols as well as options
            If m.Bars.Size > 0 And Len(m.strSecType) = 1 And Not m.bIsSpread Then ' if not an Option
                ' but ignore non-Forex Indices unless TickMove = 0.01
                If m.strSecType <> "I" Or m.Bars.TickMove = 0.01 Or IsForex(m.strSymbol) Then
                    .dPastGoodPrice = m.Bars(eBARS_Close, m.Bars.Size - 1)
                End If
            End If
            
            ' allow SalmonDLL to use the MinMove (except for Future Options, since can't be believed)
            If m.strSecType = "FO" Then
                .dMinMove = 0
            Else
                .dMinMove = m.Bars.MinMove(LastDailyDownload + 2)
            End If

            .nTradeCount = -1
            .nBidAskCount = -1
            .nSymbolID = m.nSymbolID
            .strSymbol = m.strSymbol & Chr(0)
            .strExchange = Chr(0)
            
            If g.RealTime.IsBrokerSymbol(m.strSymbol) Then
                .nBrokerOverride = 0
            Else
                .nBrokerOverride = -1
            End If

            ' get the GMT offset (only for Foreign symbols)
            .nGmtOffset = 0
            .nDomestic = 1 ' (default)
            If m.nSymbolID > 0 Then
                If SU_GetTimeInfo(g.SU, m.nSymbolID, Date, TimeInfo) <> 0 Then
                    If TimeInfo.cFeedTime <> "N" Then
                        .nDomestic = 0
                        .nGmtOffset = TimeInfo.iLocalToGmtOffset
                    End If
                End If
            End If
            
            .dBidPrice = kNullData
            .dAskPrice = kNullData
            .dOpen = kNullData
            .dHigh = kNullData
            .dLow = kNullData
            .dClose = kNullData
            .dSettle = kNullData
            .dPrevOpen = kNullData
            .dPrevHigh = kNullData
            .dPrevLow = kNullData
            .dPrevClose = kNullData
            .dPrevSettle = kNullData
        End With
        
        Set m.NewTicks = m.Bars.MakeCopy(True)
        m.NewTicks.ArrayMask = eBARS_TickByTick
        m.NewTicks.Prop(eBARS_Periodicity) = ePRD_EachTick + 1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolInfo.SetSymbol"
End Sub

' Fix Price:
' - rounds price to the symbol's min move
' - auto-fix conversion factor issues for prices from the data feed (futures only)
' - remove really obvious spikes
Private Function FixPrice(ByVal dPrice As Double) As Double
On Error GoTo ErrSection:

'The quick and dirty algorithm for Rex can be summarized as follows:
'for consecutive ticks ABC, look for a spike or dagger at B
'    a spike is when B > (A * 1.5) and B > (C * 1.5)
'    a dagger is when B < (A / 1.5) and B < (C / 1.5)  -- equivalent to 67% of A and C
'    [ Note: for futures, the boundary value for spikes and daggers is 1.015]
'for consecutive ticks ABCD, look for B & C both being spikes or B & C both being daggers
'    a spike is when B and C > (A * 1.5) and B and C > (D * 1.5)
'    a dagger is when B and C < (A / 1.5) and B and C < (D / 1.5)
'Corner cases:  first four ticks ABCD, look for a spike or dagger in A in relation to B, C, and D
'                      last four ticks WXYZ, look for spike or dagger in Z in relation to W, X, and Y

    Dim bFixed As Boolean, dPastGoodPrice As Double
    
    ' TLB 10/29/2010: for a little while, let's keep checking when in IDE just to make sure
    ' we don't really need to be fixing the prices anymore (should be ok now)
    FixPrice = dPrice
    If IsIDE Then
        dPrice = RoundToMinMove(dPrice, m.DailySummary.dMinMove)
        If FixPrice <> dPrice Then
            frmMain.Caption = "FixPrice ERROR for: " & m.strSymbol & " at " & Format(FixPrice, "#0.00000000000000000000") _
                & "  " & Str(dPrice)
            DebugLog "FixPrice ERROR for: " & m.strSymbol & " at " & Format(FixPrice, "#0.00000000000000000000") _
                & "  " & Str(dPrice)
        End If
    End If
    
#If 0 Then
    ' ORIGINAL code:
    FixPrice = dPrice
    If FixPrice <> kNullData Then
        ' only for Futures, check the PastGoodPrice for conv factor adjustment
        'If m.strSecType = "F" Then
        If 0 Then ' this functionality is getting moved over to the SalmonDLL
            dPastGoodPrice = m.DailySummary.dPastGoodPrice
            If dPastGoodPrice > 0 And dPrice > 0 And dPastGoodPrice < 1000000 Then
                While FixPrice < dPastGoodPrice * 0.33
                    FixPrice = FixPrice * 10#
                    bFixed = True
                Wend
                While FixPrice > dPastGoodPrice * 6.7
                    FixPrice = FixPrice / 10#
                    bFixed = True
                Wend
                If bFixed Then
                    ' if not within range, set back to original
                    If FixPrice < dPastGoodPrice * 0.8 Or FixPrice > dPastGoodPrice * 1.5 Then
                        FixPrice = dPrice
                    End If
                End If
            End If
        End If
        
        ' (TLB: we wanted to move this over to the SalmonDLL, but the rounding issues still exist)
        ' round to nearest MinMove (except for Future Options -- MinMove = 0)
        If m.DailySummary.dMinMove > 0 Then
            If 0 Then ' IsIDE Then
                'If m.strSecType = "S" Then
                    If FixPrice <> Int(FixPrice / m.DailySummary.dMinMove + 0.5) * m.DailySummary.dMinMove Then
                        frmMain.Caption = "FixPrice ERROR for: " & m.strSymbol & " at " & Format(dPrice, "#0.00000000000000000000") _
                            & "  " & Str(FixPrice - Int(FixPrice / m.DailySummary.dMinMove + 0.5) * m.DailySummary.dMinMove)
                    End If
                'End If
            End If
            'FixPrice = Int(FixPrice / m.DailySummary.dMinMove + 0.5) * m.DailySummary.dMinMove
            FixPrice = RoundToMinMove(FixPrice, m.DailySummary.dMinMove)
            If 0 Then ' IsIDE Then
                If FixPrice <> Int(FixPrice / m.DailySummary.dMinMove + 0.5) * m.DailySummary.dMinMove Then
                    frmMain.Caption = "FixPrice ERROR for: " & m.strSymbol & " at " & Format(dPrice, "#0.00000000000000000000") _
                        & "  " & Str(FixPrice - Int(FixPrice / m.DailySummary.dMinMove + 0.5) * m.DailySummary.dMinMove)
                End If
            End If
            
            If IsIDE Then
                If FixPrice <> dPrice Then
                    frmMain.Caption = "FixPrice ERROR for: " & m.strSymbol & " at " & Format(dPrice, "#0.00000000000000000000") _
                        & "  " & Str(FixPrice)
                    DebugLog "FixPrice ERROR for: " & m.strSymbol & " at " & Format(dPrice, "#0.00000000000000000000") _
                        & "  " & Str(FixPrice)
                End If
            End If
        End If
    End If
#End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolInfo.FixPrice"
End Function

Private Sub SwapData(ByRef dPrevious As Double, ByRef dCurrent As Double)
    Dim d As Double
    d = dPrevious
    dPrevious = dCurrent
    dCurrent = d
End Sub

' This routine refreshes our copy of "Daily data" from the SalmonClient
' - includes OHLCV data after the last daily download (up to 2 sessions)
' - also includes the current bid/ask info
' - very efficient since the SalmonDLL first compares TradeCount and BidAskCount
' (if it's counts are still the same, then it doesn't need to check anything else)
Private Function RefreshEOD() As eSalmonResponse
On Error GoTo ErrSection:

    Dim strDebug$, nPrevTradeCount&, nPrevBidAskCount&, nPrevDate&, dChk#, strChk$
    Dim dPrevBid#, dPrevAsk#
    Dim bFixBidAsk As Boolean
    Static bErrorLogged As Boolean
    
If Not g.RealTime.SalmonIsRunning Then
    nPrevDate = nPrevDate
End If
    
    ' if we're connecting to streaming, don't ask for a refresh prematurely
    ' -- we should wait until the stream is connected (otherwise we would
    ' just have to throw this data away and ask a 2nd time in a few seconds)
    If g.RealTime.ConnectionStatus <> eGDConnectionStatus_Disconnected Then
        If Not g.RealTime.GenRtConnected Then Exit Function
    End If
    
    If m.strSecType = "M" Then
        ' for Mutual Funds, don't bother requesting a refresh (since we don't want it added to the list)
        RefreshEOD = eSalmonNotRequested
    ElseIf Len(m.strSymbol) > 0 Then
        nPrevTradeCount = m.DailySummary.nTradeCount
        nPrevBidAskCount = m.DailySummary.nBidAskCount
        nPrevDate = m.DailySummary.dSessionDate
        dPrevBid = m.DailySummary.dBidPrice
        dPrevAsk = m.DailySummary.dAskPrice
        
        ' TLB 8/14/2014: make sure gets added to GenesisRT before Salmon so GMT offset will work correctly
        If Not m.bAddedToGenesisRT Then
            m.bAddedToGenesisRT = True
            g.RealTime.GetSymbolOffsetForGenesisRT m.strSymbol, True
        End If

        ' request settle symbol first if not yet done
        ' (so the settle price will be available when this symbol's data is)
        If Len(m.strRequestSettle) > 0 Then
            g.RealTime.SymbolInfo(m.strRequestSettle).GetDataRequestStatus 0
            m.strRequestSettle = ""
        End If
        
        RefreshEOD = GetEodData(m.Bars.BarsHandle, m.DailySummary)
        Select Case RefreshEOD
        Case eSalmonPending
            strDebug = "RefreshEOD for " & m.strSymbol & " = pending"
        Case eSalmonError
            strDebug = "RefreshEOD for " & m.strSymbol & " = ERROR"
        Case eSalmonAvailable
            ''strDebug = "RefreshEOD for " & m.strSymbol & ",  Close = " & Str(RoundNum(m.DailySummary.dClose, 5))
'If m.strSymbol = "$DJIA" Then 'ES-200909" Then
If m.bIsSpread Then
    dChk = dChk
End If
If m.DailySummary.dSettle = 0 Then
    dChk = dChk
End If

            ' fix some things if a trade has happened since last called
            With m.DailySummary
                If .nBidAskCount > nPrevBidAskCount Or .nBidAskCount = 0 Then
                    bFixBidAsk = True
                    If m.strSecType = "F" And .dClose > 0 And Not m.bIsSpread Then
                        strChk = ""
                        If .dAskPrice <> dPrevAsk Then
                            If .dAskPrice < .dClose * 0.8 Or .dAskPrice > .dClose * 1.2 Then
                                strChk = "*** BAD ASK"
                            End If
                        End If
                        If .dBidPrice <> dPrevBid Then
                            If .dBidPrice < .dClose * 0.8 Or .dBidPrice > .dClose * 1.2 Then
                                strChk = "*** BAD BID"
                            End If
                        End If
                        If Len(strChk) > 0 Then
                            strChk = strChk & " for " & m.strSymbol & ":  Bid=" & Str(.dBidPrice) _
                                & "  Ask=" & Str(.dAskPrice) & "  Last=" & Str(.dClose) _
                                & "  BidAskCount=" & Str(.nBidAskCount)
                            DebugLog strChk
                            If IsIDE Then
                                frmMain.Caption = strChk
                            End If
                        End If
                    End If
                End If
                If .nTradeCount > nPrevTradeCount Or .dSessionDate <> nPrevDate Then
                    bFixBidAsk = True
                    If .dPrevSessionDate > 0 Then
                        If .dSessionDate <= .dPrevSessionDate Then
                            ' *** ERROR: this should never happen (but it is sometimes during our initial development)
                            If Not bErrorLogged Then
                                bErrorLogged = True
                                strDebug = "SESSION ERROR with " & m.strSymbol & " at " & Format(Now, "mm/dd/yy hh:mm:ss") & vbCrLf _
                                    & "  PrevSession = " & Format(.dPrevSessionDate, "yyyymmdd") & ", OHLCV = " & Format(.dPrevOpen, " 0#.0###") & Format(.dPrevHigh, " 0#.0###") & Format(.dPrevLow, " 0#.0###") & Format(.dPrevClose, " 0#.0###") & Format(.dPrevVolume, " 0#.0###") & vbCrLf _
                                    & "  CurrSession = " & Format(.dSessionDate, "yyyymmdd") & ", OHLCV = " & Format(.dOpen, " 0#.0###") & Format(.dHigh, " 0#.0###") & Format(.dLow, " 0#.0###") & Format(.dClose, " 0#.0###") & Format(.dVolume, " 0#.0###")
                                FileFromString App.Path & "\SalmonErr.txt", strDebug, True, True
                                frmTest.AddList strDebug
                                strDebug = ""
                            End If
                            If .dSessionDate = .dPrevSessionDate Then
                                ' the current session needs to be combined from the 2 sessions
                                .dSessionDate = 0
                                .dOpen = .dPrevOpen
                                If .dPrevHigh > .dHigh Then .dHigh = .dPrevHigh
                                If .dPrevLow < .dLow Then .dLow = .dPrevLow
                                .dVolume = .dVolume + .dPrevVolume
                            Else
                                ' the sessions need to be swapped
                                SwapData .dPrevSessionDate, .dSessionDate
                                SwapData .dPrevClose, .dClose
                                SwapData .dPrevHigh, .dHigh
                                SwapData .dPrevLow, .dLow
                                SwapData .dPrevOpen, .dOpen
                                SwapData .dPrevVolume, .dVolume
                                SwapData .dPrevSettle, .dSettle
                            End If
                        End If
                        
                        ' fix prices for previous session
                        .dPrevOpen = FixPrice(.dPrevOpen)
                        .dPrevHigh = FixPrice(.dPrevHigh)
                        .dPrevLow = FixPrice(.dPrevLow)
                        .dPrevClose = FixPrice(.dPrevClose)
                        .dPrevSettle = FixPrice(.dPrevSettle)
                    End If
                    If .dSessionDate > 0 Then
                        ' fix prices for current session
                        .dOpen = FixPrice(.dOpen)
                        .dHigh = FixPrice(.dHigh)
                        .dLow = FixPrice(.dLow)
                        .dClose = FixPrice(.dClose)
                        .dSettle = FixPrice(.dSettle)
                        ' store the "fixed" exchange
                        m.strExchange = .strExchange
                        FixNullTermStr m.strExchange
                        m.strExchange = Trim(m.strExchange)
                    End If
                End If
                ' fix bid/ask prices (TLB 1/25/2010: just do these in all
                ' cases since it may have been overwritten with trades, etc)
                ''If .nBidAskCount > nPrevBidAskCount Or .nBidAskCount = 0 Then
                If .dBidPrice = 0 And .dAskPrice = 0 Then
                    .dBidPrice = kNullData
                    .dAskPrice = kNullData
                ElseIf bFixBidAsk Then
                    .dBidPrice = FixPrice(.dBidPrice)
                    .dAskPrice = FixPrice(.dAskPrice)
                ElseIf IsIDE Then
                    dChk = FixPrice(.dBidPrice)
                    If .dBidPrice <> dChk Then
                        .dBidPrice = dChk
                    End If
                    dChk = FixPrice(.dAskPrice)
                    If .dAskPrice <> dChk Then
                        .dAskPrice = dChk
                    End If
                End If
            End With
        End Select
    Else
        strDebug = "Symbol not initialized"
    End If
    
    If Len(strDebug) > 0 Then
        If frmTest.Visible Then
            frmTest.AddList strDebug
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolInfo.RefreshEOD"
End Function

' To update "daily bars" (or greater in periodicity -- but not for intraday bars)
' - it first checks for a refresh from the SalmonDLL
' - then it makes sure the Bars (passed in) are updated with the most current data
' - the volume will get updated only for daily bars (not for bar periods greater than "1 day per bar")
' - returns True if the Bars has been updated
' - sets bNewBar (if passed) to true if data exists for a new bar (e.g. so you can reload the chart)
Public Function UpdateDailyBars(Bars As cGdBars, Optional bNewBar As Boolean = False, _
            Optional strPrevUpdated As String = "", Optional ByVal bIgnoreSettles As Boolean = False, _
            Optional ByVal bIncludeBidAskTimeWhenUpdatingLastTickTime As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim i&, nBar&, dNew#, dOld#, s$, nSessionDate&, strUpdated$, dLastTickTime#
    Dim BarsType As eBarsArray
    Dim bUpdated As Boolean, bSettled As Boolean, bDailyBars As Boolean
    Static nPrevLDD&
    
    If Bars Is Nothing Then Exit Function
    
    ' check bar period
    nBar = Bars.Prop(eBARS_Periodicity)
    If nBar < ePRD_Days Then
        Exit Function ' this routine does not support intraday bars
    ElseIf nBar > ePRD_Days + 1 Then
        bDailyBars = False ' e.g. weekly, monthly, multiple days per bar, etc
    Else
        bDailyBars = True ' exactly 1 bar per day (most common)
    End If

    ' get a refresh from Salmon -- just return if no data yet available
    If RefreshEOD <> eSalmonAvailable Then Exit Function
    
If IsIDE Then
    If Not bDailyBars Then
        nBar = nBar
        'Exit Function
    End If
End If
    
    ' allow for storing the last updated counts (for efficiency)
    strUpdated = Str(m.DailySummary.nTradeCount) & " " & Str(m.DailySummary.nBidAskCount)
    If strUpdated = strPrevUpdated Then
        Exit Function
    End If
    strPrevUpdated = strUpdated
    
If IsIDE Then
    If Left(m.strSymbol, 3) = "XY-" Then
    'If m.strSymbol = "ES-200909" Then
        nBar = nBar
    End If
End If
    
    BarsType = Bars.ArrayMask
    
    ' get back to the first bar after the last daily download
    For nBar = Bars.Size To 0 Step -1
        If nBar <= 0 Then Exit For
        dOld = Bars(eBARS_DateTime, nBar - 1)
        If dOld <> kNullData And dOld <= LastDailyDownload Then
            ' TLB 4/9/2012: then need to continue going back until the first good data bar
            If Bars(eBARS_Close, nBar - 1) <> kNullData Then
                Exit For
            End If
        End If
    Next
    ' but just set nBar = 0 if no data coming from the Data Manager (e.g. for Options)
    If m.nSymbolID = 0 Then
        nBar = 0
    End If
    
    ' TLB 10/15/2013: if the last daily download was a holiday, then load the holiday bar so we can merge into streaming
    If nPrevLDD <> LastDailyDownload Or m.HolidayBar Is Nothing Then
        nPrevLDD = LastDailyDownload
        Set m.HolidayBar = New cGdBars
        If m.nSymbolID > 0 And Bars.IsHoliday(LastDailyDownload) Then
            DM_GetBars m.HolidayBar, m.nSymbolID, "Daily", LastDailyDownload, LastDailyDownload
            If m.HolidayBar(eBARS_Close, 0) <> kNullData Then
                ' set date back to Holiday date (we need to override whatever GetBars might do)
                m.HolidayBar(eBARS_DateTime, 0) = LastDailyDownload
            Else
                m.HolidayBar.Size = 0
            End If
        End If
    End If
    
    ' PREVIOUS SESSION:
    nSessionDate = m.DailySummary.dPrevSessionDate
    If nSessionDate > LastDailyDownload And m.nSymbolID > 0 And m.DailySummary.dPrevOpen <> kNullData Then
        ' if LastDailyDownload was a holiday (i.e. a m.HolidayBar exists) then merge that in here
        If m.HolidayBar.Size > 0 Then
            If m.DailySummary.dPrevSessionDate > m.HolidayBar(eBARS_DateTime, 0) Then
                m.DailySummary.dPrevOpen = m.HolidayBar(eBARS_Open, 0)
                dOld = m.HolidayBar(eBARS_High, 0)
                If dOld > m.DailySummary.dPrevHigh Then
                    m.DailySummary.dPrevHigh = dOld
                End If
                dOld = m.HolidayBar(eBARS_Low, 0)
                If dOld < m.DailySummary.dPrevLow Then
                    m.DailySummary.dPrevLow = dOld
                End If
            End If
        End If
        ' if PrevSession is a holiday, then merge it in with Session instead
        If Bars.IsHoliday(nSessionDate) Then
            m.DailySummary.dOpen = m.DailySummary.dPrevOpen
            If m.DailySummary.dPrevHigh > m.DailySummary.dHigh Then
                m.DailySummary.dHigh = m.DailySummary.dPrevHigh
            End If
            If m.DailySummary.dPrevLow < m.DailySummary.dLow Then
                m.DailySummary.dLow = m.DailySummary.dPrevLow
            End If
            'Note: doesn't work to merge volumes here
            nSessionDate = 0 ' so will skip the PrevSession section below
        End If
    End If
    If nSessionDate > LastDailyDownload Or (m.nSymbolID = 0 And nSessionDate > 0) Then
        If Bars.Size <= nBar Then
            Bars.Size = nBar + 1
            bNewBar = True
            bUpdated = True
        End If
        
        ' check the date
        dOld = Bars(eBARS_DateTime, nBar)
        If dOld <> nSessionDate Then
            If bDailyBars Or dOld = kNullData Or nSessionDate > dOld Then
                Bars(eBARS_DateTime, nBar) = nSessionDate
                bUpdated = True
            End If
        End If
        
        ' check the prices
        If (BarsType And eBARS_Open) <> 0 Then
            dNew = m.DailySummary.dPrevOpen
            dOld = Bars(eBARS_Open, nBar)
            If dOld <> dNew Then
                If bDailyBars Or dOld = kNullData Then
                    Bars(eBARS_Open, nBar) = dNew
                    bUpdated = True
                End If
            End If
        End If
        If (BarsType And eBARS_High) <> 0 Then
            dNew = m.DailySummary.dPrevHigh
            dOld = Bars(eBARS_High, nBar)
            If dOld <> dNew Then
                If bDailyBars Or dOld = kNullData Or dNew > dOld Then
                    Bars(eBARS_High, nBar) = dNew
                    bUpdated = True
                End If
            End If
        End If
        If (BarsType And eBARS_Low) <> 0 Then
            dNew = m.DailySummary.dPrevLow
            dOld = Bars(eBARS_Low, nBar)
            If dOld <> dNew Then
                If bDailyBars Or dOld = kNullData Or dNew < dOld Then
                    Bars(eBARS_Low, nBar) = dNew
                    bUpdated = True
                End If
            End If
        End If
        If (BarsType And eBARS_Close) <> 0 Then
            ' if there is a valid settle, then use it
            If m.DailySummary.dPrevSettle <> kNullData Then
                ' unless we want to ignore it and we also have a valid close
                If bIgnoreSettles And m.DailySummary.dPrevClose <> kNullData Then
                    dNew = m.DailySummary.dPrevClose
                    bSettled = False
                Else
                    dNew = m.DailySummary.dPrevSettle
                    bSettled = True
                End If
            Else
                ' otherwise use the last price
                dNew = m.DailySummary.dPrevClose
                If dNew <> kNullData Then bSettled = False
            End If
            dOld = Bars(eBARS_Close, nBar)
            If dOld <> dNew Then
                If dOld = kNullData Then
                    bNewBar = True
                End If
                ' update the previous session if daily bars or a new bar
                If bDailyBars Or dOld = kNullData Then
                    Bars(eBARS_Close, nBar) = dNew
                    bUpdated = True
                ' or if the current session will start a new bar
                ElseIf m.DailySummary.dSessionDate > Bars(eBARS_DateTime, nBar) Then
                    Bars(eBARS_Close, nBar) = dNew
                    bUpdated = True
                End If
            End If
        End If
        
        ' volume is only valid to update for DailyBars
        If bDailyBars Then
            If (BarsType And eBARS_Vol) <> 0 Then
                dNew = m.DailySummary.dPrevVolume
                If Bars(eBARS_Vol, nBar) <> dNew Then
                    Bars(eBARS_Vol, nBar) = dNew
                    bUpdated = True
                End If
            End If
            If (BarsType And eBARS_ContVol) <> 0 Then
                dNew = m.DailySummary.dPrevVolume
                If Bars(eBARS_ContVol, nBar) <> dNew Then
                    Bars(eBARS_ContVol, nBar) = dNew
                    bUpdated = True
                End If
            End If
            If (BarsType And eBARS_OI) <> 0 Then
                dNew = m.DailySummary.dPrevOpenInterest
                If dNew <= 0 Then
                    dNew = kNullData
                End If
                If Bars(eBARS_OI, nBar) <> dNew Then
                    Bars(eBARS_OI, nBar) = dNew
                    bUpdated = True
                End If
            End If
        End If
        
        ' bump up to the next bar (unless staying within the same bar)
        If bDailyBars Then
            nBar = nBar + 1
        ElseIf m.DailySummary.dSessionDate > Bars(eBARS_DateTime, nBar) Then
            nBar = nBar + 1
        End If
    End If
            
    ' CURRENT SESSION
    nSessionDate = m.DailySummary.dSessionDate
    If nSessionDate > LastDailyDownload And m.nSymbolID > 0 And m.DailySummary.dOpen <> kNullData Then
        ' if LastDailyDownload was a holiday (i.e. a m.HolidayBar exists) then merge that in here
        If m.HolidayBar.Size > 0 Then
            ' unless was merged into the PrevSession
            If m.DailySummary.dPrevSessionDate <= m.HolidayBar(eBARS_DateTime, 0) Then
                m.DailySummary.dOpen = m.HolidayBar(eBARS_Open, 0)
                dOld = m.HolidayBar(eBARS_High, 0)
                If dOld > m.DailySummary.dHigh Then
                    m.DailySummary.dHigh = dOld
                End If
                dOld = m.HolidayBar(eBARS_Low, 0)
                If dOld < m.DailySummary.dLow Then
                    m.DailySummary.dLow = dOld
                End If
            End If
        End If
        ' if Session is a holiday, then bump the date ahead
        If Bars.IsHoliday(nSessionDate) Then
            Do
                nSessionDate = nSessionDate + 1
            Loop While Not IsWeekday(nSessionDate)
        End If
    End If
    If nSessionDate > LastDailyDownload Or (m.nSymbolID = 0 And nSessionDate > 0) Then
        If Bars.Size <= nBar Then
            Bars.Size = nBar + 1
            bNewBar = True
            bUpdated = True
        End If
        
        ' check the date
        dOld = Bars(eBARS_DateTime, nBar)
        If dOld <> nSessionDate Then
            If bDailyBars Or dOld = kNullData Or nSessionDate > dOld Then
                Bars(eBARS_DateTime, nBar) = nSessionDate
                bUpdated = True
            End If
        End If
        
        ' check the prices
        If (BarsType And eBARS_Open) <> 0 Then
            dNew = m.DailySummary.dOpen
            dOld = Bars(eBARS_Open, nBar)
            If dOld <> dNew Then
                If bDailyBars Or dOld = kNullData Then
                    Bars(eBARS_Open, nBar) = dNew
                    bUpdated = True
                End If
            End If
        End If
        If (BarsType And eBARS_High) <> 0 Then
            dNew = m.DailySummary.dHigh
            dOld = Bars(eBARS_High, nBar)
            If dOld <> dNew Then
                If bDailyBars Or dOld = kNullData Or dNew > dOld Then
                    Bars(eBARS_High, nBar) = dNew
                    bUpdated = True
                End If
            End If
        End If
        If (BarsType And eBARS_Low) <> 0 Then
            dNew = m.DailySummary.dLow
            dOld = Bars(eBARS_Low, nBar)
            If dOld <> dNew Then
                If bDailyBars Or dOld = kNullData Or dNew < dOld Then
                    Bars(eBARS_Low, nBar) = dNew
                    bUpdated = True
                End If
            End If
        End If
        If (BarsType And eBARS_Close) <> 0 Then
            ' if there is a valid settle, then use it
If m.DailySummary.dSettle = 0 Then
    dNew = dNew
End If
            If m.DailySummary.dSettle <> kNullData Then
                ' unless we want to ignore it and we also have a valid close
                If bIgnoreSettles And m.DailySummary.dClose <> kNullData Then
                    dNew = m.DailySummary.dClose
                    bSettled = False
                Else
                    dNew = m.DailySummary.dSettle
                    bSettled = True
                End If
            Else
                ' otherwise use the last price
                dNew = m.DailySummary.dClose
                If dNew <> kNullData Then bSettled = False
            End If
            dOld = Bars(eBARS_Close, nBar)
            If dOld <> dNew Then
                If dOld = kNullData Then
                    bNewBar = True
                End If
                Bars(eBARS_Close, nBar) = dNew
                bUpdated = True
            End If
        End If
        
        ' volume is only valid to update for DailyBars
        If bDailyBars Then
            If (BarsType And eBARS_Vol) <> 0 Then
                dNew = m.DailySummary.dVolume
                If Bars(eBARS_Vol, nBar) <> dNew Then
                    Bars(eBARS_Vol, nBar) = dNew
                    bUpdated = True
                End If
            End If
            If (BarsType And eBARS_ContVol) <> 0 Then
                dNew = m.DailySummary.dVolume
                If Bars(eBARS_ContVol, nBar) <> dNew Then
                    Bars(eBARS_ContVol, nBar) = dNew
                    bUpdated = True
                End If
            End If
            If (BarsType And eBARS_OI) <> 0 Then
                dNew = m.DailySummary.dOpenInterest
                If dNew <= 0 Then
                    dNew = kNullData
                End If
                If Bars(eBARS_OI, nBar) <> dNew Then
                    Bars(eBARS_OI, nBar) = dNew
                    bUpdated = True
                End If
            End If
        End If
        
        dLastTickTime = m.DailySummary.dLastTickTime
        
        ' update the Exchange field (for the last trade)
        If Len(m.strExchange) > 0 And bUpdated Then
            If Bars.Prop(eBARS_Exchange) <> m.strExchange Then
                Bars.Prop(eBARS_Exchange) = m.strExchange
                bUpdated = True
            End If
        End If
    End If
   
    ' TLB 6/12/2012: if higher than Daily period, need to rebuild in the correct periodicity
    ' TLB 7/24/2013: NO, we can ONLY build higher than Daily period by using Daily bars
    ' as the source -- so we CANNOT do this here!
    'If bUpdated And Not bDailyBars Then
    '    Bars.BuildBars Bars.Prop(eBARS_PeriodicityStr)
    'End If
    
    ' update the Settle flag
    If Abs(Bars.Prop(eBARS_PriceHasSettled)) <> Abs(bSettled) Then
        Bars.Prop(eBARS_PriceHasSettled) = bSettled
        bUpdated = True
    End If
    
    ' update misc properties
    Bars.Prop(eBARS_LastTickDown) = m.DailySummary.nLastTickDown
    Bars.Prop(eBARS_LastTickVol) = m.DailySummary.nLastTickVol '(do we even use this?)
    
    ' the most recent Bid/Ask info should just go into the last data bar
    For nBar = nBar To 0 Step -1
        If Bars(eBARS_Close, nBar) <> kNullData Then
            If (BarsType And eBARS_Bid) <> 0 Then
                dNew = m.DailySummary.dBidPrice
                If Bars(eBARS_Bid, nBar) <> dNew And (dNew > 0 Or m.bIsSpread) Then
                    Bars(eBARS_Bid, nBar) = dNew
                    bUpdated = True
                    If bIncludeBidAskTimeWhenUpdatingLastTickTime Then
                        If m.DailySummary.dBidTime > dLastTickTime Then
                            dLastTickTime = m.DailySummary.dBidTime
                        End If
                    End If
                End If
            End If
            If (BarsType And eBARS_Ask) <> 0 Then
                dNew = m.DailySummary.dAskPrice
                If Bars(eBARS_Ask, nBar) <> dNew And (dNew > 0 Or m.bIsSpread) Then
                    Bars(eBARS_Ask, nBar) = dNew
                    bUpdated = True
                    If bIncludeBidAskTimeWhenUpdatingLastTickTime Then
                        If m.DailySummary.dAskTime > dLastTickTime Then
                            dLastTickTime = m.DailySummary.dAskTime
                        End If
                    End If
                End If
            End If
            If (BarsType And eBARS_BidSize) <> 0 Then
                dNew = m.DailySummary.dBidSize
                If Bars(eBARS_BidSize, nBar) <> dNew And dNew > 0 Then
                    Bars(eBARS_BidSize, nBar) = dNew
                    bUpdated = True
                    If bIncludeBidAskTimeWhenUpdatingLastTickTime Then
                        If m.DailySummary.dBidTime > dLastTickTime Then
                            dLastTickTime = m.DailySummary.dBidTime
                        End If
                    End If
                End If
            End If
            If (BarsType And eBARS_AskSize) <> 0 Then
                dNew = m.DailySummary.dAskSize
                If Bars(eBARS_AskSize, nBar) <> dNew And dNew > 0 Then
                    Bars(eBARS_AskSize, nBar) = dNew
                    bUpdated = True
                    If bIncludeBidAskTimeWhenUpdatingLastTickTime Then
                        If m.DailySummary.dAskTime > dLastTickTime Then
                            dLastTickTime = m.DailySummary.dAskTime
                        End If
                    End If
                End If
            End If
            Exit For
        End If
    Next
    
    ' LastTickTime in gdBars is stored as # minutes since midnight
    ' (in fractional minutes -- calculated to the nearest second)
    If dLastTickTime > 0 Then
        dNew = MinutesFromMidnight(dLastTickTime)
        If Bars.Prop(eBARS_LastTickTime) <> dNew Then
            Bars.Prop(eBARS_LastTickTime) = dNew
            bUpdated = True
        End If
    End If
    
    UpdateDailyBars = bUpdated

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolInfo.UpdateDailyBars"
End Function

Public Function GetIntradayData(ByVal nPeriodicity&) As cGdBars
On Error GoTo ErrSection:

    Dim i&, strDebug$, dBarTime#, dPrevBarTime#, dTempBarTime#, nSkipCount&
    Dim bFullTicks As Boolean
    Dim nSalmonResponse As eSalmonResponse
    Dim Bars As cGdBars
    Static bMinuteErrorLogged As Boolean, bFullTicksErrorLogged As Boolean

    Set Bars = m.Bars.MakeCopy(True)
    
' tempfix ????: since requesting intraday data for Options will crash
If InStr(m.strSymbol, " ") > 0 Then
    If Not FileExist(App.Path & "\SalmonDev.flg") Then
        GoTo ErrExit:
    End If
End If
    
    ' we just want to make sure EOD data has been requested first
    ' (it doesn't really matter whether it's available yet or not)
    If RefreshEOD > eSalmonNotRequested Then
        If GetPeriodType(nPeriodicity) = ePRD_Minutes Or nPeriodicity >= ePRD_Days Then
            ' Get 1-Minute Bars
            bFullTicks = False
            Bars.ArrayMask = eBARS_Intraday
            Bars.Prop(eBARS_Periodicity) = ePRD_Minutes + 1
'DebugLog "v=" & Str(Int(gdTickCount))
            nSalmonResponse = GetMinuteBarData(Bars.BarsHandle, m.DailySummary)
'DebugLog "w=" & Str(Int(gdTickCount))
        Else
            ' Get Full Ticks
            bFullTicks = True
            Bars.ArrayMask = eBARS_TickByTick
            Bars.Prop(eBARS_Periodicity) = ePRD_EachTick + 1
            nSalmonResponse = GetFullTickData(Bars.BarsHandle, m.DailySummary)
        End If
        
        Select Case nSalmonResponse
        Case eSalmonPending
            If bFullTicks Then
                strDebug = "GetFullTicks for " & m.strSymbol & " = pending"
            Else
                strDebug = "GetMinuteBars for " & m.strSymbol & " = pending"
            End If
        Case eSalmonError
            If bFullTicks Then
                strDebug = "GetFullTicks for " & m.strSymbol & " = ERROR"
            Else
                strDebug = "GetMinuteBars for " & m.strSymbol & " = ERROR"
            End If
        Case eSalmonAvailable
            If FileExist(App.Path & "\DumpToCache.flg") Then
                If bFullTicks Then
                    Bars.DumpToFile App.Path & "\Cache\ft~"
                Else
                    Bars.DumpToFile App.Path & "\Cache\mb~"
                End If
                
' TLB 10/29/2010: looks like we no longer need to do this kind of fixing (except when dumping cache file) ...
                ' fix the incoming bar data
                nSkipCount = 0
'DebugLog "x=" & Str(Int(gdTickCount))
                dPrevBarTime = 999999
                For i = Bars.Size - 1 To 0 Step -1
                    dBarTime = RoundToSecond(Bars(eBARS_DateTime, i))
                    ' check for out-of-sequence bars (but do allow for 2 consecutive bars with same DateTime)
                    If dBarTime > dPrevBarTime Or dBarTime <= 0 Then
                        ' ERROR: bars are not chronological!
                        ' (so we must skip over the bad bars until back in sequence)
                        nSkipCount = nSkipCount + 1
                        dTempBarTime = Bars(eBARS_DateTime, i + 1)
                        If dTempBarTime < 0 Then dTempBarTime = 0 '(just so invalid date won't error in the "Format" below)
                        If dBarTime < 0 Then dBarTime = 0 '(just so invalid date won't error in the "Format" below)
                        If bFullTicks And Not bFullTicksErrorLogged Then
                            bFullTicksErrorLogged = True
                            On Error Resume Next
                            strDebug = "FullTicks SEQUENCE ERROR with " & m.strSymbol & " at " & Format(Now, "mm/dd/yy hh:mm:ss") & vbCrLf _
                                & "  Bar " & Str(i) & " = " & Format(dBarTime, "yyyymmdd hh:mm:ss") & "  " & Format(Bars(eBARS_Close, i), " 0#.0###") & vbCrLf _
                                & "  Bar " & Str(i + 1) & " = " & Format(dTempBarTime, "yyyymmdd hh:mm:ss") & "  " & Format(Bars(eBARS_Close, i + 1), " 0#.0###")
                            FileFromString App.Path & "\SalmonErr.txt", strDebug, True, True
                            frmTest.AddList strDebug
                        ElseIf Not bFullTicks And Not bMinuteErrorLogged Then
                            bMinuteErrorLogged = True
                            On Error Resume Next
                            strDebug = "MinuteBar SEQUENCE ERROR with " & m.strSymbol & " at " & Format(Now, "mm/dd/yy hh:mm:ss") & vbCrLf _
                                & "  Bar " & Str(i) & " = " & Format(dBarTime, "yyyymmdd hh:mm:ss") & "  " & Format(Bars(eBARS_Close, i), " 0#.0###") & vbCrLf _
                                & "  Bar " & Str(i + 1) & " = " & Format(dTempBarTime, "yyyymmdd hh:mm:ss") & "  " & Format(Bars(eBARS_Close, i + 1), " 0#.0###")
                            FileFromString App.Path & "\SalmonErr.txt", strDebug, True, True
                            frmTest.AddList strDebug
                        End If
                    Else
                        ' if there were bad bars, remove them now
                        If nSkipCount > 0 Then
                            If Not FileExist(App.Path & "\SalmonDev.flg") Then
                                Bars.DeleteSomeBars i + 1, nSkipCount
                            End If
                            nSkipCount = 0
                        End If
                        ' fix the data for this bar
                        dPrevBarTime = dBarTime
                        Bars(eBARS_DateTime, i) = dBarTime
                        If m.DailySummary.dMinMove > 0 Then
                            If Not bFullTicks Then
                                Bars(eBARS_Open, i) = FixPrice(Bars(eBARS_Open, i))
                                Bars(eBARS_High, i) = FixPrice(Bars(eBARS_High, i))
                                Bars(eBARS_Low, i) = FixPrice(Bars(eBARS_Low, i))
                            End If
                            Bars(eBARS_Close, i) = FixPrice(Bars(eBARS_Close, i))
                        End If
                    End If
                Next
                If nSkipCount > 0 Then
                    If Not FileExist(App.Path & "\SalmonDev.flg") Then
                        Bars.DeleteFirstBars nSkipCount
                    End If
                    nSkipCount = 0
                End If
            End If
            
            If bFullTicks Then
                strDebug = "GetFullTicks for " & m.strSymbol & ", #Ticks=" & Str(Bars.Size) _
                    & ", LastTick=" & Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
            Else
                strDebug = "GetMinuteBars for " & m.strSymbol & ", #Minutes=" & Str(Bars.Size) _
                    & ", LastClose=" & Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
            End If
'DebugLog "y=" & Str(Int(gdTickCount))
        End Select
    End If

    If Len(strDebug) > 0 And frmTest.Visible Then
        frmTest.AddList strDebug
    End If

ErrExit:
    Set GetIntradayData = Bars
    Exit Function
    
ErrSection:
    RaiseError "cSymbolInfo.GetIntradayData"
End Function

Public Function GetDataRequestStatus(ByVal nPeriodicity&) As eSalmonResponse
On Error GoTo ErrSection:
    
    Dim nResponse As eSalmonResponse
    Dim TempBars As cGdBars
    
If Not g.RealTime.SalmonIsRunning Then
    nResponse = nResponse
End If
    
    nResponse = eSalmonNotRequested
    If Len(m.strSymbol) > 0 And Not m.Bars Is Nothing Then
        If nPeriodicity >= ePRD_Days Or nPeriodicity = 0 Then
            ' Daily data: this call will submit request and return status
            nResponse = RefreshEOD
        Else ' Intraday data
            If GetPeriodType(nPeriodicity) = ePRD_Minutes Then
                nResponse = GetMinuteBarStatus(m.Bars.BarsHandle)
            Else
                nResponse = GetFullTickStatus(m.Bars.BarsHandle)
            End If
            ' if intraday data not yet requested, request it now
            If nResponse = eSalmonNotRequested Then
                Set TempBars = GetIntradayData(nPeriodicity)
                Set TempBars = Nothing
                ' and return status (most likely now pending)
                If GetPeriodType(nPeriodicity) = ePRD_Minutes Then
                    nResponse = GetMinuteBarStatus(m.Bars.BarsHandle)
                Else
                    nResponse = GetFullTickStatus(m.Bars.BarsHandle)
                End If
            End If
        End If
    End If
    
ErrExit:
    GetDataRequestStatus = nResponse
    Exit Function
    
ErrSection:
    RaiseError "cSymbolInfo.GetDataRequestStatus"
End Function

Public Function GetNewlyStreamedTicks() As cGdBars
On Error GoTo ErrSection:

    Dim strDebug$, i&

If Not g.RealTime.SalmonIsRunning Then
    i = i
End If
    
    If Not m.NewTicks Is Nothing Then
        m.NewTicks.Size = 0
        If GetStreamedTicks(m.NewTicks.BarsHandle, m.DailySummary) > 0 Then
            ' fix the incoming prices
            If m.DailySummary.dMinMove > 0 Then
                For i = 0 To m.NewTicks.Size - 1
                    m.NewTicks(eBARS_Close, i) = FixPrice(m.NewTicks(eBARS_Close, i))
                    'frmTest.AddList m.NewTicks(eBARS_Flags, i)
                Next
            End If
            'strDebug = "GetStreamedTicks for " & m.strSymbol & ", #Ticks=" & Str(Bars.Size) _
                    & ", LastTick=" & Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
        Else
            m.NewTicks.Size = 0
        End If
    End If

    If Len(strDebug) > 0 And frmTest.Visible Then
        frmTest.AddList strDebug
    End If

ErrExit:
    Set GetNewlyStreamedTicks = m.NewTicks
    Exit Function
    
ErrSection:
    RaiseError "cSymbolInfo.GetNewlyStreamedTicks"
End Function


'========= Functions to quickly retrieve current info for this symbol =======

Public Function Symbol() As String
    Symbol = m.strSymbol
End Function

Public Function LastTradedSession() As Long
On Error GoTo ErrSection:
    
    Dim nSessionDate&
    If Len(m.strSymbol) > 0 Then
        If m.DailySummary.dSessionDate > 0 Then
            nSessionDate = m.DailySummary.dSessionDate
        ElseIf m.Bars.Size > 0 Then
            nSessionDate = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
        End If
    End If
    LastTradedSession = nSessionDate
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolInfo.LastTradedSession"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastKnownPrice
'' Description: Get the last known price for the given symbol
'' Inputs:      Symbol, Local Time?, Date/Time of last price(return)
''      iBidTradeOrAsk:  -1 = LastBid,  0 = LastTrade,  1 = LastAsk
'' Returns:     Last Known Price (kNullData if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastKnownPrice(Optional ByVal iBidTradeOrAsk As Integer = 0, _
                    Optional ByVal bLocalTime As Boolean = False, Optional dDateTime As Double) As Double
On Error GoTo ErrSection:

    Dim i&, dPrice#
    
    dPrice = kNullData
    dDateTime = kNullData
    If Len(m.strSymbol) > 0 Then
        ' get the price and time from the Salmon data
        If RefreshEOD = eSalmonAvailable Then
            With m.DailySummary
                If iBidTradeOrAsk < 0 Then
                    If (.dBidPrice > 0 Or m.bIsSpread) Then
                        dPrice = .dBidPrice
                        dDateTime = .dBidTime
                    End If
                ElseIf iBidTradeOrAsk > 0 Then
                    If (.dAskPrice > 0 Or m.bIsSpread) Then
                        dPrice = .dAskPrice
                        dDateTime = .dAskTime
                    End If
                ElseIf .dSessionDate > 0 Then
                    dPrice = .dClose
                    If dPrice = kNullData And .dSettle <> kNullData Then
                        dPrice = .dSettle
                    End If
                    dDateTime = .dLastTickTime
                ElseIf .dPrevSessionDate > 0 Then
                    dPrice = .dPrevClose
                    If dPrice = kNullData And .dPrevSettle <> kNullData Then
                        dPrice = .dPrevSettle
                    End If
                    dDateTime = .dLastTickTime
                End If
            End With
        End If
        ' otherwise get it from the last daily download
        If dPrice = kNullData And m.Bars.Size > 0 Then
            If iBidTradeOrAsk < 0 Then
                dPrice = m.Bars(eBARS_Bid, m.Bars.Size - 1)
            ElseIf iBidTradeOrAsk > 0 Then
                dPrice = m.Bars(eBARS_Ask, m.Bars.Size - 1)
            Else
                dPrice = m.Bars(eBARS_Close, m.Bars.Size - 1)
            End If
            dDateTime = m.Bars.LastTickDateTime
        End If
        
        ' fix the DateTime
        If dPrice = kNullData Or dPrice = 0 Or dDateTime <= 0 Then
            dDateTime = kNullData
        Else
            dDateTime = RoundToSecond(dDateTime)
            ' convert the date/time if necessary
            If bLocalTime And (dDateTime <> Int(dDateTime)) Then
                dDateTime = ConvertTimeZone(dDateTime, m.Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            End If
        End If
    End If

ErrExit:
    LastKnownPrice = dPrice
    Exit Function
    
ErrSection:
    RaiseError "cSymbolInfo.LastKnownPrice"
End Function


Public Sub RemoveFromStream()
On Error GoTo ErrSection:
    
    Dim i&
    
    If g.RealTime Is Nothing Or g.bUnloading Then Exit Sub
    If Not g.RealTime.SalmonIsRunning Then Exit Sub
    
    If Len(m.strSymbol) > 0 Then
        RemoveSymbol m.DailySummary
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolInfo.RemoveFromStream"
End Sub

' Returns false if the Bid/Ask tables have not been changed
Public Function UpdateDOM(BidTable As cGdTable, AskTable As cGdTable) As Boolean
On Error GoTo ErrSection:

    Dim nPrevCounter&, i&
    Static nCounter&
    
    If Len(m.strSymbol) > 0 And m.strSecType <> "M" Then
        ' first time called: initialize things
        If nCounter = 0 Then
            nCounter = 1
            RefreshEOD ' just to make sure symbol gets in the stream
            BidTable.NumRecords = 0
            AskTable.NumRecords = 0
            UpdateDOM = True
        End If
        If nCounter <> 0 Then
            nPrevCounter = nCounter
            If RequestDOM(m.Bars.BarsHandle, BidTable.TableHandle, AskTable.TableHandle, nCounter) = 0 Then
                If nCounter <> nPrevCounter Then
                    For i = 0 To BidTable.NumRecords - 1
                        BidTable.Num(1, i) = RoundToMinMove(BidTable.Num(1, i), m.Bars.MinMove)
                    Next
                    For i = 0 To AskTable.NumRecords - 1
                        AskTable.Num(1, i) = RoundToMinMove(AskTable.Num(1, i), m.Bars.MinMove)
                    Next
                    UpdateDOM = True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymbolInfo.UpdateDOM"
End Function

Public Sub AddForOptNav(aSymbols As cGdArray, Optional ByVal strChainBase$ = "")
On Error GoTo ErrSection:

    Dim i&, aBase As New cGdArray
    Dim Bars As cGdBars
    
    aBase.Create eGDARRAY_Strings
    If Len(strChainBase) > 0 Then
        aBase.Add strChainBase
    End If
    
    If aSymbols.Size > 1 Then
        ' TLB: when passing multiple symbols at one time, pass just a copy of our Bars
        ' (just so no chance of getting the original messed up as Salmon processes each symbol)
        Set Bars = m.Bars.MakeCopy(True)
    Else
        ' but if only this symbol, we can simply pass this Bars directly
        Set Bars = m.Bars
    End If
    
    i = RequestEODForOptionNavigator(Bars.BarsHandle, m.DailySummary, aSymbols.ArrayHandle, aBase.ArrayHandle)
    Set Bars = Nothing
    
ErrExit:
    Set aBase = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cSymbolInfo.AddForOptNav"
End Sub

