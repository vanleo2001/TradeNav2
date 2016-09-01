VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGameMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cGameMode.cls
'' Description: Allow the user to replay a time frame of data to trade with
''
'' Author:      Genesis Financial Data Services
''              425 Wind Chime Place
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kResultsPath = "\GameResults\"

'core code in this file copied from frmReplay.frm

Public Enum eGDReplayMode
    eGDReplayMode_Off = 0
    eGDReplayMode_Pause = 1
    eGDReplayMode_Play = 2
    eGDReplayMode_Sync = 3
End Enum

Private Type mPrivate
    frmOwner As Form
    OrdersOpen As New cGdTree          'tree holding cPtOrder objects
    OrdersFilled As New cGdTree
    GameSystem As cSystem
    
    TargetOrder As cPtOrder
    StopLossOrder As cPtOrder
    BaseOrder As cPtOrder
            
    strStrategy As String
    nStrategyId As Long
    nStrategyLibId As Long
    
    dDataTime As Double                 ' use data through this time
    dSessionStart As Double             ' Session start time
    dSessionEnd As Double               ' Session end time
    dInterval As Double                 ' Interval to run at
    dUserInterval As Double             'use for fixing game interval on daily charts
    dGameInitTime As Double
    strReplaySpeed As String
    
    bAutoInterval As Boolean            ' Use automatic interval?
    bRandomSym As Boolean               ' is using random symbol?
    bShowResults As Boolean
    bReplayAll As Boolean
    bEndOfBar As Boolean                'show data only to end of bar
    bHasIntradayTicks As Boolean        'use for fixing game interval on daily charts
    nTickFirstDate As Long
    
    dAutoTargetDollars As Double
    dAutoTargetPoints As Double
    dAutoStopDollars As Double
    dAutoStopPoints As Double
    bAutoExitsDone As Boolean
    bResultFiles As Boolean
    
    dTargetDollars As Double
    dStopDollars As Double
    dTargetPoints As Double
    dStopPoints As Double
    
    nHasStop As Long
    nHasTarget As Long
    nUserForwardBack As Long            '0=normal, 1=user clicked forward, -1=user clicked rewind
End Type

Private m As mPrivate

Public Property Get GameDataTime() As Double
    GameDataTime = m.dDataTime
End Property

Public Property Let GameDataTime(ByVal dTime#)
    m.dDataTime = dTime
End Property

Public Property Get GameSessionStart() As Double
    GameSessionStart = m.dSessionStart
End Property

Public Property Let GameSessionStart(ByVal dStart#)
    m.dSessionStart = dStart
End Property

Public Property Get GameSessionEnd() As Double
    GameSessionEnd = m.dSessionEnd
End Property

Public Property Let GameSessionEnd(ByVal dEnd#)
    m.dSessionEnd = dEnd
End Property

Public Property Get GameInterval() As Double
    If Not m.bHasIntradayTicks Then CheckForTicks
    GameInterval = m.dInterval
End Property

Public Property Let GameInterval(ByVal dInterval#)
    m.dInterval = dInterval
    If Not m.bHasIntradayTicks Then CheckForTicks
End Property

Public Property Get GameAutoInterval() As Boolean
    GameAutoInterval = m.bAutoInterval
End Property

Public Property Let GameAutoInterval(ByVal bAuto As Boolean)
    m.bAutoInterval = bAuto
End Property

Public Property Get GameRandomSym() As Boolean
    GameRandomSym = m.bRandomSym
End Property

Public Property Let GameRandomSym(ByVal bRandom As Boolean)
    m.bRandomSym = bRandom
End Property

Public Property Get GameReplayMode(ByVal bTimerEnabled As Boolean) As eGDReplayMode
    
    GameReplayMode = eGDReplayMode_Off
    
    If m.frmOwner Is Nothing Then
        GameReplayMode = eGDReplayMode_Sync
    Else
        If m.dDataTime > 0 Then
            If bTimerEnabled Then
                GameReplayMode = eGDReplayMode_Play
            Else
                GameReplayMode = eGDReplayMode_Pause
            End If
        End If
    End If

End Property

Public Property Get GameLogicalDate() As Long
    If m.dDataTime > Int(m.dDataTime) + m.dSessionEnd Then
        GameLogicalDate = Int(m.dDataTime) + 1
    Else
        GameLogicalDate = Int(m.dDataTime)
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGame
'' Description: Initialize values for replaying data for game mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InitGame(frm As frmChart) As Boolean        'game always start as non-detached chart
On Error GoTo ErrSection:

    Dim dDate As Double                 ' Date to run
    Dim dInterval As Double             ' Interval to run
    Dim bInitialized As Boolean

    'StopPlay
    
    If frm Is Nothing Then Exit Function
    
    Set m.frmOwner = frm
    m.frmOwner.GameMode = Me
    m.frmOwner.IsInGameMode = True
        
    If m.dDataTime = 0 Then m.dDataTime = DateSerial(2003, 12, 1)
    dDate = m.dDataTime
    
    m.dUserInterval = 0
    If m.bAutoInterval Then
        dInterval = 0
        bInitialized = GetSessionTimes()
    Else
        FixInterval
        dInterval = m.dInterval
        m.dUserInterval = gdFixDateTime(dInterval / 1440)   'save in case user requested minute interval for daily charts that have no tick data at game start
        m.dInterval = gdFixDateTime(dInterval / 1440)
        bInitialized = GetSessionTimes()        'true
    End If
        
    If Not bInitialized Then Exit Function
    
    m.dDataTime = dDate + m.dSessionStart
    If m.dSessionStart > m.dSessionEnd Then
        m.dDataTime = m.dDataTime - 1
    End If
    
    AdjustTime -99
    
    m.dGameInitTime = m.dDataTime
        
    InitGame = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cGameMode.Init", eGDRaiseError_Raise
    
End Function

Private Sub CheckSessionTimes(frm As Form, nBarInterval&, _
    bEvening As Boolean)
On Error GoTo ErrSection:

    Dim dStart#, dEnd#, strSymbol$
    
    With frm.Chart.Bars
        dStart = .Prop(eBARS_StartTime)
        dEnd = .Prop(eBARS_EndTime)
        If dStart > 0 Or dEnd > 0 Then
            strSymbol = .Prop(eBARS_Symbol)
            If Left(strSymbol, 1) = "$" And Not IsForex(strSymbol) Then
                If dEnd > 16 * 60 Then dEnd = 16 * 60
            End If
            ' get earliest start time
            If dStart > dEnd Then
                ' this starts prior evening
                If dStart < m.dSessionStart Or Not bEvening Then
                    m.dSessionStart = dStart
                End If
                bEvening = True
            ElseIf Not bEvening Then
                If dStart < m.dSessionStart Or m.dSessionStart = -1 Then
                    m.dSessionStart = dStart
                End If
            End If
            ' get latest end time
            If dEnd > m.dSessionEnd Then
                m.dSessionEnd = dEnd
            End If
        End If
        ' determine smallest bar interval
        If .IsIntraday Then
            If .Prop(eBARS_PeriodType) = ePRD_Minutes Then
                If .Prop(eBARS_PeriodsPerBar) < nBarInterval Then
                    nBarInterval = .Prop(eBARS_PeriodsPerBar)
                End If
            ElseIf nBarInterval > 5 Then
                nBarInterval = 5
            End If
        ElseIf .Prop(eBARS_Periodicity) = ePRD_Days + 1 And nBarInterval > 1440 Then
            nBarInterval = 1440
        End If
    End With
    
    Exit Sub

ErrSection:
    RaiseError "cGameMode.CheckSessionTimes"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSessionTimes
'' Description: Get the session times for the current symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSessionTimes() As Boolean
On Error GoTo ErrSection:

    Dim i&, nBarInterval&
    Dim frm As Form
    Dim bEvening As Boolean, bMaximized As Boolean, bStrategy As Boolean
                    
    If m.frmOwner Is Nothing Then Exit Function
    If g.bUnloading Then Exit Function
            
    m.dSessionStart = -1
    m.dSessionEnd = -1
    nBarInterval = 1441
    
    ' see if owner chart is maximized
    If m.frmOwner.WindowState = vbMaximized Then bMaximized = True
    
    'check for strategies
    If m.frmOwner.Chart.ShowTrades And m.frmOwner.Chart.SystemID > 0 Then
        bStrategy = True
    End If
    
    'set start/end date for data
    CheckSessionTimes m.frmOwner, nBarInterval, bEvening
    
    'set game mode objects for other charts
    If m.bReplayAll Then
        For i = 0 To Forms.Count - 1
            If IsFrmChart(Forms(i)) Then
                If g.bUnloading Then Exit Function
                Set frm = Forms(i)
                If frm.GameMode Is Nothing Then
                    frm.GameMode = New cGameMode
                    frm.GameMode.GameDataTime = m.dDataTime
                    frm.IsInGameMode = True
                End If
            End If
        Next
    End If
    Set frm = Nothing
        
    ' check that didn't wrap around
    If m.dSessionStart < m.dSessionEnd And bEvening Then
        m.dSessionStart = m.dSessionEnd + 1
    End If
        
    CheckForTicks
    
    m.dSessionStart = gdFixDateTime(m.dSessionStart / 1440)
    m.dSessionEnd = gdFixDateTime(m.dSessionEnd / 1440)
    
    GetSessionTimes = True
    

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cGameMode.GetSessionTimes", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdjustTime
'' Description: Adjust time accordingly
'' Inputs:      Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' iMode: 0=normal play,
' 1=forward non-intraday (alternating: end of current session, start of next session)
' 2=forward (always skip to start of next session)
' -1=rewind non-intraday (alternating: start of current session, end of previous session)
'
'JM: 08-08-2006:
'This is correct fix that Tim made which takes care of aardvark issues 3754, 3250,
'3471 & 3432, which also made the ForwindTime & RewindTime functions unnecessary.
'
'JM: 10-31-2006:
'Forward and rewind for intraday data is handled in GameForward & GameBack.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AdjustTime(ByVal iMode%)
On Error GoTo ErrSection:

    Dim dTime As Double                 ' Time of the replay
    Dim dSessionDate As Double
    Dim Bars As cGdBars
        
    If m.dDataTime = 0 Then Exit Sub
    
    Set Bars = m.frmOwner.Chart.Bars
    dSessionDate = Bars.SessionDateForTime(m.dDataTime, False)
        
    ' if already at or past end of session, go to start of next session
    dTime = gdFixDateTime(m.dDataTime - Int(m.dDataTime))
    If iMode = 0 Then
        If m.dSessionEnd >= m.dSessionStart Then
            If dTime >= m.dSessionEnd Then
                iMode = 2
            ElseIf dTime < m.dSessionStart Then
                iMode = 2
                m.dDataTime = m.dDataTime - 1
            End If
        ElseIf m.strReplaySpeed = "day" Then      '4094
            If dTime >= m.dSessionEnd Then
                iMode = 1
            ElseIf dTime <= m.dSessionStart Then
                iMode = 2
            End If
        Else
            If dTime >= m.dSessionEnd And dTime < m.dSessionStart Then
                iMode = 2
            End If
        End If
    End If
        
    If iMode = -99 Then
        ' when first start, go back to midnight (i.e. before day session starts)
        'm.dDataTime = Int(m.dDataTime) '- 1 + m.dSessionEnd
        
        ' when first start, go back to end of previous session
        m.dDataTime = dSessionDate - 1
        Do While Not IsWeekday(m.dDataTime)
            m.dDataTime = m.dDataTime - 1 '(bump back to Friday)
        Loop
        m.dDataTime = m.dDataTime + m.dSessionEnd
    ElseIf iMode < 0 Then
        If dTime <> m.dSessionStart Then
            ' go to start of current session
            m.dDataTime = dSessionDate + m.dSessionStart
            If m.dSessionEnd < m.dSessionStart Then
                m.dDataTime = m.dDataTime - 1 '(session starts prior evening)
            End If
        Else
            ' go to end of previous session
            m.dDataTime = dSessionDate - 1
            Do While Not IsWeekday(m.dDataTime)
                m.dDataTime = m.dDataTime - 1 '(bump back to Friday)
            Loop
            m.dDataTime = m.dDataTime + m.dSessionEnd
        End If
    ElseIf iMode > 0 Then
        ' iMode=1: alternate to end of current session, then start of next session
        If iMode = 1 And dTime <> m.dSessionEnd Then
'JM:06-16-2008 - fix for 4094 implemented above did away with necessity for this code
' go to end of current session (intraday is not handled here, but check just in case)
'If IsForex(Bars.Prop(eBARS_Symbol)) And Not Bars.IsIntraday Then
'    dSessionDate = dSessionDate + 1
'End If
            m.dDataTime = dSessionDate + m.dSessionEnd
        Else
            ' go to start of next session
            m.dDataTime = dSessionDate + 1
            Do While Not IsWeekday(m.dDataTime)
                m.dDataTime = m.dDataTime + 1 '(bump up to Monday)
            Loop
            m.dDataTime = m.dDataTime + m.dSessionStart
            If m.dSessionEnd < m.dSessionStart Then
                m.dDataTime = m.dDataTime - 1 '(session starts prior evening)
            End If
            
            If IsForex(Bars.Prop(eBARS_Symbol)) Then
'forex session starts in evening of one day and ends evening of next day
'e.g. starts=[10-23-YYYY]22:00 ends=[10-24-YYYY]21:59
'A data time of [10-24-YYYY]22:00 should be new bar for new session date [10-25-YYYY].
'Does not work as expected. Just go full bar for now per Tim.
'JM:06-16-2008 fix for 4094 implemented above and below now has Forex working properly
                If Bars.IsIntraday Then
                    m.dDataTime = m.dDataTime + 1 / 86400#
                Else
                    m.dDataTime = m.dDataTime + 1 / 1440#           '4094
                End If
            End If
            
        End If
    ElseIf dTime < m.dSessionEnd And gdFixDateTime(dTime + m.dInterval) >= m.dSessionEnd Then
        ' if adding interval goes past end of session then set to end of session
        m.dDataTime = Int(m.dDataTime) + m.dSessionEnd
    Else
        ' else just add interval
        m.dDataTime = m.dDataTime + m.dInterval
        If m.dInterval > 4 / 1440# Then
            m.dDataTime = Int(m.dDataTime * 1440 / 5# + 0.000001) * 5 / 1440#
        End If
    End If
    m.dDataTime = gdFixDateTime(m.dDataTime)
        
    If iMode = 0 Then
        UpdateCharts
    Else
        UpdateCharts 1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGameMode.AdjustTime", eGDRaiseError_Raise
    
End Sub
    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateCharts
'' Description: Update the charts accordingly
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateCharts(Optional ByVal iAction% = 0)
On Error GoTo ErrSection:
   
    Dim i&
    Dim frm As Form
    Dim bMaximized As Boolean
    
    If m.frmOwner Is Nothing Or g.bUnloading Then Exit Sub
    
    ' see if active chart is maximized
    If m.frmOwner.GameReplayMode = eGDReplayMode_Sync Then
        Exit Sub        'precautionary - should never happen
    ElseIf m.frmOwner.WindowState = vbMaximized Then
        bMaximized = True
    End If
            
    'update main instant replay chart
    With m.frmOwner.Chart
        .ResetLastScreenDate
        Select Case iAction
            Case 1
                .GenerateChart eRedo9_ReloadData, True
            
            Case Else
                .GenerateChart eRedo7_ReloadRT, True
        
        End Select
    End With
    
    'check real time in case user started real time after starting game
    If g.RealTime.Active Then Exit Sub
    
    Dim nPrevSession&, nCurrSession&
    Dim dNewDataTime#
    
    If m.bReplayAll Then        'And Not bMaximized Then    -6473 (Pete wants to replay on all charts all the time)
        'update all charts
        For i = 0 To Forms.Count - 1
            If g.bUnloading Then Exit For
            
            If IsFrmChart(Forms(i)) Then
                Set frm = Forms(i)
                With frm
                    If .GameReplayMode = eGDReplayMode_Sync Then
                        If Not .GameMode Is Nothing Then
                            dNewDataTime = ConvertTimeZone(m.dDataTime, m.frmOwner.Chart.Bars.Prop(eBARS_ExchangeTimeZoneInf), frm.Chart.Bars.Prop(eBARS_ExchangeTimeZoneInf))
                            nPrevSession = TradingSession(frm.Chart.Bars, .GameMode.GameDataTime)
                            nCurrSession = TradingSession(frm.Chart.Bars, dNewDataTime)
                            .GameMode.GameDataTime = dNewDataTime
                            .Chart.ResetLastScreenDate
                            If Not g.bUnloading Then
                                If nPrevSession = nCurrSession Then
                                    .Chart.GenerateChart eRedo7_ReloadRT, True
                                Else
                                    .Chart.GenerateChart eRedo9_ReloadData, True
                                End If
                            End If
                        End If
                    End If
                End With
            End If
        Next
    End If
    Set frm = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError "cGameMode.UpdateCharts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GamePlayTime
''              this is the ShowTime() sub original instant replay form
'' Description: Return string for showing time information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get GamePlayTime() As String
On Error GoTo ErrSection:

    If m.dDataTime > 0 Then
        GamePlayTime = "Instant Replay:  " & DateFormat(m.dDataTime, MM_DD_YYYY, H_MM, AMPM_LOWER, True)
    Else
        GamePlayTime = "Instant Replay"
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cGameMode.GamePlayTime", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GameTimer
''              this is the tmr_Timer sub in original instant replay form
'' Description: Check some things every so often
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GameTimer() As Boolean
On Error GoTo ErrSection:


    Dim bProcessed As Boolean
    Dim i&
    Dim Bars As cGdBars
    Static dLastBarDate#, dLastTickTime#
    
    Do While m.dInterval <> 0 And m.dDataTime < Date + 2
        
        AdjustTime 0
               
        Set Bars = m.frmOwner.Chart.Bars
        For i = Bars.Size - 1 To 0 Step -1
            If Bars(eBARS_Close, i) <> kNullData Then
                If dLastBarDate <> Bars(eBARS_DateTime, i) Then
                    dLastBarDate = Bars(eBARS_DateTime, i)
                    bProcessed = True
                    Exit For
                End If
                'Exit For
            End If
        Next
        If dLastTickTime <> Bars.Prop(eBARS_LastTickTime) Then
            dLastTickTime = Bars.Prop(eBARS_LastTickTime)
            bProcessed = True
        End If
        Set Bars = Nothing
        
        If bProcessed Then Exit Do
    Loop
    
    GameTimer = bProcessed
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cGameMode.GameTimer", eGDRaiseError_Show
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GameStart
''              this is StartPlay sub in original instant replay form
'' Description: Start play mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GameStart() As Boolean
On Error GoTo ErrSection:

    Dim bOkay As Boolean

    bOkay = GetSessionTimes()

    If bOkay Then
        If m.dDataTime = 0 Then bOkay = False
    End If

    GameStart = bOkay

    Exit Function

ErrSection:
    RaiseError "cGameMode.GameStart"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GameModeBack
''              this is cmdBack_Click in original instant replay form
'' Description: Move backwards in time
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GameBack()
On Error GoTo ErrSection:

    Dim i&, Count&
        
    Dim iBar&, iPeriod&
    Dim Bars As cGdBars
    Dim dBarTime#, dPrevBarTime#
    Dim dTime#, dDiff#
    
    Set Bars = m.frmOwner.Chart.Bars
    
    If Bars Is Nothing Then Exit Sub        'precautionary, theoretically should never happen
            
    m.bEndOfBar = False
        
    iBar = m.frmOwner.Chart.LastGoodDataBar(False)
    If Bars(eBARS_DateTime, iBar) <= m.dGameInitTime Then
        InfBox "Cannot rewind past replay start date/time."
        GoTo ErrExit
    ElseIf iBar <= 1 Then
        InfBox "Cannot rewind past current bar."
        GoTo ErrExit
    End If
    
    If Bars.IsIntraday Then
        If iBar > 0 Then
            UpdateCharts
            iBar = m.frmOwner.Chart.LastGoodDataBar(False)
            m.dDataTime = Bars(eBARS_DateTime, iBar)
            m.frmOwner.Chart.SetFormCaption
        End If
    Else
        GetSessionTimes
        AdjustTime -1
    End If
    
    m.bEndOfBar = False
    
    If m.nStrategyId = 0 Or m.nStrategyId = kGameModeSysID Then
        'remove trades from chart
        m.frmOwner.Chart.DeleteGameOrders
        m.frmOwner.Chart.DeleteGameTrades m.dDataTime
        
        RemoveAutoExits True, True
        m.OrdersOpen.Clear
        
        Count = m.OrdersFilled.Count
        For i = Count To 1 Step -1
            If m.OrdersFilled(i).OrderDate >= m.dDataTime Then
                m.OrdersFilled.Remove i
            End If
        Next
    End If
       
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGameMode.GameBack"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GameModeForward
''              this is cmdForward_Click in original instant replay form
'' Description: Move forwards in time
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GameForward()
On Error GoTo ErrSection:

    Dim iBar&, iCount&
    Dim Bars As cGdBars
    Dim dTime#, dBarTime#
    
    Set Bars = m.frmOwner.Chart.Bars
    
    If Bars Is Nothing Then Exit Sub        'precautionary, theoretically should never happen
            
    m.bEndOfBar = False
    
    If m.dDataTime > Date + 2 Then Exit Sub
       
    If Bars.IsIntraday Then
        m.bEndOfBar = False
        
        ' bump up game date/time to after last bar
        iBar = m.frmOwner.Chart.LastGoodDataBar(False)
        dBarTime = gdFixDateTime(Bars(eBARS_DateTime, iBar))
        dTime = gdFixDateTime(m.dDataTime)
        m.dDataTime = dBarTime + 1 / 1440#
            
        ' update the chart, but if we didn't add a bar then we need to go to the next session
        UpdateCharts
        
        Do While True
            iBar = m.frmOwner.Chart.LastGoodDataBar(False)
            If dBarTime < gdFixDateTime(Bars(eBARS_DateTime, iBar)) Then
                Exit Do
            End If
            
            If iCount > 90 Then
                Exit Do         '!don't want to get into an infinite loop!
            Else
                AdjustTime 2                '5943
                iCount = iCount + 1
                m.frmOwner.Chart.SetFormCaption
                DoEvents
            End If
        Loop
        
        ' reset the game date/time to the date/time of the last bar
        iBar = m.frmOwner.Chart.LastGoodDataBar(False)
        dBarTime = gdFixDateTime(Bars(eBARS_DateTime, iBar))
        m.dDataTime = dBarTime
        m.frmOwner.Chart.SetFormCaption
        
        If iCount > 90 Then
            DebugLog "Instant Replay FF failed after 90 attemps. Symbol=" & Bars.Prop(eBARS_Symbol) & _
                " Periodicity=" & Bars.Prop(eBARS_PeriodicityStr) & " GameDataTime=" & Str(m.dDataTime) & "(" & DateFormat(m.dDataTime, MM_DD_YYYY, HH_MM_SS) & ")"
            
            InfBox "Data appears too sparse for Instant Replay" & vbCrLf & "Fast Forward function. Please try" & vbCrLf & "different replay settings.", "I", "Ok", "Instant Replay"
        End If
        
    Else
        GetSessionTimes
        AdjustTime 1
    End If
    
    m.bEndOfBar = False
    
    Exit Sub
    
ErrSection:
    RaiseError "cGameMode.GameForward"
    
End Sub

Public Function GameNewOrder(NewOrder As cPtOrder, _
    Optional ByVal bAddToTree As Boolean = True) As Boolean

    Dim eType As eTT_OrderType
    Dim Ind As cIndicator
    Dim Bars As cGdBars
    Dim strText$, strKeySave$, dMarketPrice#
    Dim aFlds As New cGdArray
    Dim dOrderDate#, iLastDataBar&
        
    Set Ind = m.frmOwner.Chart.Tree("PRICE")
    If Ind Is Nothing Then Exit Function
        
    Set Bars = Ind.Bars
    If Bars Is Nothing Then Exit Function
    If Not bAddToTree Then strKeySave = Str(NewOrder.OrderID)
        
    If NewOrder.OrderType = eTT_OrderType_Market Then
        If NewOrder.Status <> eTT_OrderStatus_Filled Then
            strText = UCase(NewOrder.EngineString)
            iLastDataBar = m.frmOwner.Chart.LastGoodDataBar(False)
            If m.frmOwner.Chart.Bars.Prop(eBARS_PeriodType) = ePRD_Days And Not m.bHasIntradayTicks Then
                'dOrderDate = Bars(eBARS_DateTime, m.frmOwner.Chart.LastGoodDataBar(False))
                dOrderDate = Bars(eBARS_DateTime, iLastDataBar)
                If m.dDataTime > MidSessionTime Then
                    dMarketPrice = Bars(eBARS_Close, iLastDataBar)
                Else
                    dMarketPrice = Bars(eBARS_Open, iLastDataBar)
                End If
            Else
                dMarketPrice = Bars(eBARS_Close, iLastDataBar)
                dOrderDate = m.dDataTime
            End If
            If InStr(strText, "EL") Or InStr(strText, "XS") Then
                m.frmOwner.Chart.AddReplayOrder NewOrder.Quantity, "M", dMarketPrice, dOrderDate    'm.dDataTime
            ElseIf InStr(strText, "ES") Or InStr(strText, "XL") Then
                m.frmOwner.Chart.AddReplayOrder -1 * NewOrder.Quantity, "M", dMarketPrice, dOrderDate   'm.dDataTime
            End If
            If Not bAddToTree Then
                'this is an existing non-market order that got changed to a market order
                m.OrdersOpen.Remove Str(NewOrder.OrderID)
                bAddToTree = True
            End If
            NewOrder.StatusDate = dOrderDate  'm.dDataTime
            NewOrder.OrderDate = DateFormat(dOrderDate, MM_DD_YY, HH_MM_SS)    'DateFormat(m.dDataTime, MM_DD_YY, HH_MM_SS)
            NewOrder.Status = eTT_OrderStatus_Filled
            NewOrder.Fees = dMarketPrice
            NewOrder.OrderID = m.OrdersFilled.Count + 1
        End If
    ElseIf NewOrder.OrderType = eTT_OrderType_StopWithLimit Then
        aFlds.SplitFields (UCase(NewOrder.EngineString))
        strText = aFlds(0) & vbTab & aFlds(1) & vbTab & "0" & vbTab & "S" & vbTab & Str(NewOrder.StopPrice) & vbTab & vbTab & vbTab & vbTab & Str(NewOrder.Quantity)
        m.frmOwner.Chart.AddReplayOrder NewOrder.Quantity, "", dMarketPrice, m.dDataTime, strText, NewOrder
        strText = aFlds(0) & vbTab & aFlds(1) & vbTab & "0" & vbTab & "L" & vbTab & Str(NewOrder.LimitPrice) & vbTab & vbTab & vbTab & vbTab & Str(NewOrder.Quantity)
        m.frmOwner.Chart.AddReplayOrder NewOrder.Quantity, "", dMarketPrice, m.dDataTime, strText, NewOrder
    Else
        strText = UCase(NewOrder.EngineString)
        m.frmOwner.Chart.AddReplayOrder NewOrder.Quantity, "", dMarketPrice, m.dDataTime, NewOrder.EngineString, NewOrder
    End If
    
    If bAddToTree Then
        If NewOrder.OrderID > 0 Then
            If NewOrder.Status = eTT_OrderStatus_Filled Then
                m.OrdersFilled.Add NewOrder, Str(NewOrder.OrderID)
                If m.bAutoExitsDone Then
                    If m.BaseOrder Is Nothing Then
                        If m.nHasTarget Then
                            If m.dTargetDollars = 0 Then
                                SetTargetPrice m.dAutoTargetPoints, m.dAutoTargetDollars, NewOrder
                            Else
                                SetTargetPrice m.dTargetPoints, m.dTargetDollars, NewOrder
                            End If
                        End If
                        If m.nHasStop Then
                            If m.dAutoStopDollars = 0 Then
                                SetStopPrice m.dAutoStopPoints, m.dAutoStopDollars, NewOrder
                            Else
                                SetStopPrice m.dStopPoints, m.dStopDollars, NewOrder
                            End If
                        End If
                    End If
                Else
                    CalcAutoExits
                    frmGameTargetLoss.ShowMe Me, m.frmOwner, NewOrder
                    m.bAutoExitsDone = True
                End If
            Else
                m.OrdersOpen.Add NewOrder, Str(NewOrder.OrderID)
            End If
        Else
            strText = "Error" 'should never get here
        End If
    Else
        m.OrdersOpen.Key(strKeySave) = Str(NewOrder.OrderID)
    End If
            
End Function

Public Sub GameUpdateOrder()
    
    Dim i&
    Dim Order As cPtOrder
    
    m.frmOwner.Chart.DeleteGameOrders
    For i = 1 To m.OrdersOpen.Count
        Set Order = m.OrdersOpen(i)
        If Not Order Is Nothing Then
            GameNewOrder Order, False
        End If
    Next
    
End Sub

Public Sub GameDeleteOrder(Order As cPtOrder)

    Dim SWLOrder As cPtOrder
    Dim nID&

    nID = Order.OrderID
    If m.OrdersOpen(nID) Is Nothing Then
        Set SWLOrder = m.OrdersOpen(nID - 1)
        If SWLOrder Is Nothing Then
            nID = -1
        Else
            If SWLOrder.OrderType = eTT_OrderType_StopWithLimit Then
                nID = nID - 1
            Else
                nID = -1
            End If
        End If
    End If

    If Not m.OrdersOpen(nID) Is Nothing Then
        m.OrdersOpen.Remove nID
        GameUpdateOrder
    End If

End Sub

Public Sub ResetGameMode()

    Dim strFiles$
    
    m.OrdersOpen.Clear
    m.OrdersFilled.Clear
    Set m.TargetOrder = Nothing
    Set m.StopLossOrder = Nothing
    Set m.BaseOrder = Nothing
    
    m.bAutoInterval = True
    m.bReplayAll = False
    m.dDataTime = 0
    m.dInterval = 0
    m.dSessionEnd = 0
    m.dSessionStart = 0
    m.nStrategyId = 0
    m.nStrategyLibId = 0
    m.strStrategy = ""
        
    m.dAutoStopDollars = 0
    m.dAutoTargetDollars = 0
    m.dAutoStopPoints = 0
    m.dAutoTargetPoints = 0
    
    m.dStopDollars = 0
    m.dTargetDollars = 0
    m.dStopPoints = 0
    m.dTargetPoints = 0
    
    m.nHasStop = 0
    m.nHasTarget = 0
    m.bAutoExitsDone = False
    m.bShowResults = True
    
    'check for game result files
    strFiles = App.Path & kResultsPath & "*.*"
    If FileExist(strFiles) Then
        m.bResultFiles = True
    Else
        m.bResultFiles = False
    End If

End Sub

Public Property Get OpenOrder(ByVal nIdx&) As cPtOrder
    
    Dim Order As cPtOrder
    
    If nIdx >= 0 Then
        Set Order = m.OrdersOpen(nIdx)
        'stop with limit orders have 2 entries in the aOrders array in frmChart
        'but only one order object in the open orders tree in this class
        If Order Is Nothing Then
            Set Order = m.OrdersOpen(nIdx - 1)
            If Not Order Is Nothing Then
                If Order.OrderType <> eTT_OrderType_StopWithLimit Then
                    Set Order = Nothing
                End If
            End If
        End If
        
    End If

    Set OpenOrder = Order
    
End Property

Public Function GameReport() As String

    Dim i&, j&
    Dim strTradesFile$, strText$
    Dim aTrades As New cGdArray
    Dim Bars As cGdBars
    
    Dim bEntryFound As Boolean
    Dim bExitFound As Boolean
        
    If m.OrdersFilled.Count = 0 Then
        GameReport = "There are no trades to create report for."
        Exit Function
    ElseIf m.OrdersFilled.Count = 1 Then
        GameReport = "At least one entry and one exit trade is needed to create a report."
        Exit Function
    End If
    
    For i = 1 To m.OrdersFilled.Count
        strText = m.OrdersFilled(i).EngineString
        If InStr(strText, "EL") Or InStr(strText, "ES") Then bEntryFound = True
        If InStr(strText, "XL") Or InStr(strText, "XS") Then bExitFound = True
        If bEntryFound And bExitFound Then
            Exit For
        End If
    Next
    
    If Not bEntryFound Then
        GameReport = "No entry trade found. Cannot create report."
        Exit Function
    ElseIf Not bExitFound Then
        GameReport = "No exit trade found. Cannot create report."
        Exit Function
    End If
    
    
    Dim tbReport As New cGdTable
    Dim Order As cPtOrder
    Dim EntryOrder As cPtOrder
    Dim strPosition$
    
    tbReport.CreateField eGDARRAY_Shorts, 0, "NotUsed", 0
    tbReport.CreateField eGDARRAY_Strings, 1, "Position"
    tbReport.CreateField eGDARRAY_Doubles, 2, "EntryDate", 0
    tbReport.CreateField eGDARRAY_Doubles, 3, "EntryPrice", 0
    tbReport.CreateField eGDARRAY_Strings, 4, "EntrySignal"
    tbReport.CreateField eGDARRAY_Doubles, 5, "ExitDate", 0
    tbReport.CreateField eGDARRAY_Doubles, 6, "ExitPrice", 0
    tbReport.CreateField eGDARRAY_Strings, 7, "ExitSignal"
    tbReport.CreateField eGDARRAY_Doubles, 8, "ProfitLoss", 0
    tbReport.CreateField eGDARRAY_Doubles, 9, "MaxProfit", 0
    tbReport.CreateField eGDARRAY_Doubles, 10, "MaxLoss", 0
    tbReport.CreateField eGDARRAY_Longs, 11, "BarsInTrade", 0
    tbReport.CreateField eGDARRAY_Longs, 12, "EntryRule", -1
    tbReport.CreateField eGDARRAY_Longs, 13, "ExitRule", -2
    tbReport.CreateField eGDARRAY_Longs, 14, "TDOW", 0
    tbReport.CreateField eGDARRAY_Longs, 15, "TDOM", 0
    tbReport.CreateField eGDARRAY_Longs, 16, "TDOY", 0
    tbReport.CreateField eGDARRAY_Longs, 17, "MaxProfitLossFirst", 0
    tbReport.CreateField eGDARRAY_Longs, 18, "EntryOrderID", 0
    tbReport.CreateField eGDARRAY_Longs, 19, "EntryQuantity", 0
            
    Set Bars = m.frmOwner.Chart.Bars
    If Bars Is Nothing Then Exit Function
    
    strTradesFile = App.Path & "\Trades\GameTrades.Txt"
    
    aTrades.Create eGDARRAY_Strings
    aTrades(0) = BuildHeader(Bars, Bars(eBARS_DateTime, 0), Bars(eBARS_DateTime, Bars.Size - 1), strTradesFile)
    
    'go through tree adding trades information to table
    For i = 1 To m.OrdersFilled.Count
        strPosition = ""
        Set Order = m.OrdersFilled(i)
        If Not Order Is Nothing Then
            strText = Order.EngineString
            If InStr(strText, "EL") Then
                strPosition = "L"
                tbReport.AddRecord ""
                SetTblEntryFields Order, Bars, tbReport, strPosition
                Set EntryOrder = Order
            ElseIf InStr(strText, "ES") Then
                strPosition = "S"
                tbReport.AddRecord ""
                SetTblEntryFields Order, Bars, tbReport, strPosition
                Set EntryOrder = Order
            ElseIf InStr(strText, "XL") Then
                strPosition = "L"
                If Not EntryOrder Is Nothing Then
                    SetTblExitFields EntryOrder, Order, Bars, tbReport, strPosition
                    Set EntryOrder = Nothing
                End If
            ElseIf InStr(strText, "XS") Then
                strPosition = "S"
                If Not EntryOrder Is Nothing Then
                    SetTblExitFields EntryOrder, Order, Bars, tbReport, strPosition
                    Set EntryOrder = Nothing
                End If
            End If
        End If
    Next
    
    'add entry with matching exit trades to array
    For i = 0 To tbReport.NumRecords - 1
        strText = ""
        If tbReport(18, i) = 0 Then
            strText = tbReport(1, i)
            For j = 2 To 17
                strText = strText & vbTab & tbReport(j, i)
            Next
            aTrades.Add strText
        End If
    Next
    
    If aTrades.Size > 1 Then
        aTrades.ToFile strTradesFile
        GameReport = strTradesFile
    Else
        GameReport = "Replay report generation failed."
    End If
    
    aTrades.Destroy

End Function

Private Function BuildHeader(ByVal Bars As cGdBars, ByVal dFromDate As Double, ByVal dToDate As Double, _
                            Optional ByVal strReportName As String = "") As String
On Error GoTo ErrSection:

    Dim astrHeader As New cGdArray
    
    Set astrHeader = New cGdArray
    astrHeader.Create eGDARRAY_Strings

    astrHeader(0) = CStr(kGameModeSysID)
    astrHeader(1) = strReportName
    
    astrHeader(2) = Bars.Prop(eBARS_PeriodicityStr)
    astrHeader(3) = dFromDate
    astrHeader(4) = dToDate
    astrHeader(5) = Str(Bars.Size)
    astrHeader(6) = "0"
    astrHeader(7) = Bars.Prop(eBARS_Symbol)
    astrHeader(8) = Str(Bars.Prop(eBARS_TickMove))
    astrHeader(9) = Str(Bars.Prop(eBARS_TickValue))
    astrHeader(10) = Str(Bars.Prop(eBARS_MinMoveInTicks))
    astrHeader(11) = Str(Bars.Prop(eBARS_Margin))
    astrHeader(12) = Bars.SecurityType
    astrHeader(13) = Str(Bars.Prop(eBARS_StartTime))
    astrHeader(14) = Str(Bars.Prop(eBARS_EndTime))
    astrHeader(15) = Bars.Prop(eBARS_ExchangeTimeZoneInf)
    
    BuildHeader = astrHeader.JoinFields(vbTab)

ErrExit:
    Set astrHeader = Nothing
    Exit Function
    
ErrSection:
    Set astrHeader = Nothing
    RaiseError "cSystem.BuildHeader", eGDRaiseError_Raise
    
End Function

Public Sub SetStopPrice(ByVal dPoints#, ByVal dDollars#, FilledOrder As cPtOrder)
    
    Dim strText$, dPrice#
    Dim bNewOrder As Boolean
    
    If m.BaseOrder Is Nothing And FilledOrder Is Nothing Then
        If m.frmOwner.Chart.Position > 0 Then
            Set m.BaseOrder = m.OrdersFilled(m.OrdersFilled.Count)
            If m.BaseOrder Is Nothing Then
                Exit Sub
            Else
                m.BaseOrder.Message = "BaseOrder"
            End If
        Else
            Exit Sub
        End If
    End If
    
    If m.StopLossOrder Is Nothing Then
        Set m.StopLossOrder = New cPtOrder
        m.StopLossOrder.Message = "AutoStop"
        bNewOrder = True
    End If
    
    If m.BaseOrder Is Nothing Then
        Set m.BaseOrder = FilledOrder
        m.BaseOrder.Message = "BaseOrder"
    ElseIf Not FilledOrder Is Nothing Then
        If FilledOrder.Message <> m.BaseOrder.Message Then Exit Sub
    End If
        
    If m.BaseOrder Is Nothing Then Exit Sub     'precautionary should never happen
    
    strText = UCase(m.BaseOrder.OrderText)
    With m.StopLossOrder
        If m.BaseOrder.Buy Then
            dPrice = m.BaseOrder.Fees - dPoints
        Else
            dPrice = m.BaseOrder.Fees + dPoints
        End If
        If InStr(strText, "SELL") Then
            .Buy = True
        Else
            .Buy = False
        End If
        .OrderType = eTT_OrderType_Stop
        .LimitPrice = kNullData
        .StopPrice = Abs(dPrice)
        .Quantity = m.BaseOrder.Quantity
    End With
    
    If bNewOrder Then
        GameNewOrder m.StopLossOrder
    Else
        GameUpdateOrder
    End If
    
    m.dStopPoints = dPoints
    m.dStopDollars = dDollars

End Sub

Public Sub SetTargetPrice(ByVal dPoints#, ByVal dDollars#, FilledOrder As cPtOrder)
    
    Dim strText$, dPrice#
    Dim bNewOrder As Boolean
        
    If m.BaseOrder Is Nothing And FilledOrder Is Nothing Then
        If m.frmOwner.Chart.Position > 0 Then
            Set m.BaseOrder = m.OrdersFilled(m.OrdersFilled.Count)
            If m.BaseOrder Is Nothing Then
                Exit Sub
            Else
                m.BaseOrder.Message = "BaseOrder"
            End If
        Else
            Exit Sub
        End If
    End If
        
    If m.TargetOrder Is Nothing Then
        Set m.TargetOrder = New cPtOrder
        m.TargetOrder.Message = "AutoTarget"
        bNewOrder = True
    End If
        
    If m.BaseOrder Is Nothing Then
        Set m.BaseOrder = FilledOrder
        m.BaseOrder.Message = "BaseOrder"
    ElseIf Not FilledOrder Is Nothing Then
        If FilledOrder.Message <> m.BaseOrder.Message Then Exit Sub
    End If
        
    If m.BaseOrder Is Nothing Then Exit Sub     'precautionary should never happen
            
    strText = UCase(m.BaseOrder.OrderText)
    With m.TargetOrder
        If m.BaseOrder.Buy Then
            dPrice = m.BaseOrder.Fees + dPoints
        Else
            dPrice = m.BaseOrder.Fees - dPoints
        End If
        If InStr(strText, "SELL") Then
            .Buy = True
        Else
            .Buy = False
        End If
        .OrderType = eTT_OrderType_Limit
        .LimitPrice = Abs(dPrice)
        .StopPrice = kNullData
        .Quantity = m.BaseOrder.Quantity
    End With
    
    If bNewOrder Then
        GameNewOrder m.TargetOrder
    Else
        GameUpdateOrder
    End If
    
    m.dTargetPoints = dPoints
    m.dTargetDollars = dDollars

End Sub

Public Function IsTargetOrder(Order As cPtOrder) As Boolean

    Dim bAutoTarget As Boolean
    
    If Order.Message = "AutoTarget" Then bAutoTarget = True
    
    IsTargetOrder = bAutoTarget

End Function

Public Function IsStopLossOrder(Order As cPtOrder) As Boolean

    Dim bAutoStop As Boolean
    
    If Order.Message = "AutoStop" Then bAutoStop = True
    
    IsStopLossOrder = bAutoStop
    
End Function

Public Property Get TargetOrder() As cPtOrder

    Set TargetOrder = m.TargetOrder
    
End Property

Public Property Get StopLossOrder() As cPtOrder

    Set StopLossOrder = m.StopLossOrder
    
End Property

Public Property Get BaseOrder() As cPtOrder

    Set BaseOrder = m.BaseOrder

End Property

Public Sub OrderFilled(ByVal nIdx&, ByVal nQuantity&, ByVal dPrice#)

    Dim Order As cPtOrder
    Dim nOldId&, strText$
    
    Set Order = m.OrdersOpen(Str(nIdx))
    
    If Order Is Nothing Then
        Set Order = m.OrdersOpen(Str(nIdx - 1))
        If Not Order Is Nothing Then
            If Order.OrderType <> eTT_OrderType_StopWithLimit Then
                Set Order = Nothing
            End If
        End If
    End If
    
    If Not Order Is Nothing Then
        nOldId = Order.OrderID
        strText = Order.EngineString
        Order.Status = eTT_OrderStatus_Filled
        Order.StatusDate = m.dDataTime
        Order.OrderDate = DateFormat(m.dDataTime, MM_DD_YY, HH_MM_SS)
        Order.Fees = dPrice
        Order.OrderID = m.OrdersFilled.Count + 1
        m.OrdersOpen.Remove nIdx
        m.OrdersFilled.Add Order, Str(Order.OrderID)
        If Order.Message = "AutoTarget" Or Order.Message = "AutoStop" Then
            RemoveAutoExits True, True
        ElseIf m.bAutoExitsDone Then
            'if the order that just got filled is an entry (EL or ES) then trigger auto exit orders if they are defined
            If InStr(strText, "EL") Or InStr(strText, "ES") Then         '6165
                If m.BaseOrder Is Nothing Then
                    If m.nHasTarget Then
                        If m.dTargetDollars = 0 Then
                            SetTargetPrice m.dAutoTargetPoints, m.dAutoTargetDollars, Order
                        Else
                            SetTargetPrice m.dTargetPoints, m.dTargetDollars, Order
                        End If
                    End If
                    If m.nHasStop Then
                        If m.dStopDollars = 0 Then
                            SetStopPrice m.dAutoStopPoints, m.dAutoStopDollars, Order
                        Else
                            SetStopPrice m.dStopPoints, m.dStopDollars, Order
                        End If
                    End If
                End If
            End If
        Else
            CalcAutoExits
            frmGameTargetLoss.ShowMe Me, m.frmOwner, Order
            m.bAutoExitsDone = True
        End If
        m.frmOwner.EnableGameControls
    End If

End Sub

Private Function OpenEntryTradeIdx(tbData As cGdTable, ByVal strEntryType$) As Long

    Dim nIdx&, i&
    
    nIdx = -1
    
    For i = 0 To tbData.NumRecords - 1
        If tbData(18, i) > 0 And tbData(19, i) > 0 Then
            If tbData(1, i) = strEntryType Then
                nIdx = i
                Exit For
            End If
        End If
    Next

    OpenEntryTradeIdx = nIdx

End Function

Private Function BarsInTrade(Bars As cGdBars, nEntryIdx&, nExitIdx&, _
    ByVal dDateEntry#, ByVal dDateExit#) As Long

    Dim hArray&, iEntry&, iExit&
    Dim nBars&, dDate1#, dDate2#
    
    hArray = Bars.ArrayHandle(eBARS_DateTime)
    
    If Bars.IsIntraday Then
       dDate1 = dDateEntry
       dDate2 = dDateExit
    Else
        dDate1 = Int(dDateEntry)
        dDate2 = Int(dDateExit)
    End If
    
    nEntryIdx = -1
    nExitIdx = -1
    gdBinarySearch hArray, dDate1, iEntry, eGdSort_Default, -1, -1
    If Bars(eBARS_DateTime, iEntry) = dDate1 Then
        gdBinarySearch hArray, dDate2, iExit, eGdSort_Default, -1, -1
        If Bars(eBARS_DateTime, iExit) = dDate2 Then
            nBars = iExit - iEntry
            nEntryIdx = iEntry
            nExitIdx = iExit
        End If
    End If
    
    BarsInTrade = nBars

End Function

Private Sub MaxProfitLoss(Bars As cGdBars, dMaxProfit#, dMaxLoss#, _
    ByVal iEntryIdx&, ByVal iExitIdx&, ByVal nQty&, ByVal strPosition$, _
    ByVal dEntryPrice#)

    Dim hArray&
    Dim dHighestHigh#, dLowestLow#
    
    hArray = Bars.ArrayHandle(eBARS_High)
    dHighestHigh = gdMaxValue(hArray, iEntryIdx, iExitIdx)
    
    hArray = Bars.ArrayHandle(eBARS_Low)
    dLowestLow = gdMinValue(hArray, iEntryIdx, iExitIdx)
    
    If strPosition = "L" Then
        dMaxProfit = dHighestHigh - dEntryPrice
        dMaxProfit = dMaxProfit * nQty * Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
    
        dMaxLoss = dLowestLow - dEntryPrice
        dMaxLoss = dMaxLoss * nQty * Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
    Else
        dMaxProfit = dEntryPrice - dLowestLow
        dMaxProfit = dMaxProfit * nQty * Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
        
        dMaxLoss = dEntryPrice - dHighestHigh
        dMaxLoss = dMaxLoss * nQty * Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
    End If

End Sub

Public Sub RemoveAutoExits(ByVal bRemoveStop As Boolean, _
    ByVal bRemoveTarget As Boolean)

    Dim bUpdate As Boolean
    
    If bRemoveStop Then
        If Not m.StopLossOrder Is Nothing Then
            m.OrdersOpen.Remove (Str(m.StopLossOrder.OrderID))
            bUpdate = True
        End If
        Set m.StopLossOrder = Nothing
    End If
    
    If bRemoveTarget Then
        If Not m.TargetOrder Is Nothing Then
            m.OrdersOpen.Remove Str(m.TargetOrder.OrderID)
            bUpdate = True
        End If
        Set m.TargetOrder = Nothing
    End If
    
    If m.StopLossOrder Is Nothing And m.TargetOrder Is Nothing Then
        If Not m.BaseOrder Is Nothing Then
            'm.BaseOrder.Message = ""
            Set m.BaseOrder = Nothing
        End If
    End If
    
    If bUpdate Then GameUpdateOrder
    
End Sub

Public Property Get GameStrategyID() As Long
    GameStrategyID = m.nStrategyId
End Property

Public Property Let GameStrategyID(ByVal nID&)
    m.nStrategyId = nID
End Property

Public Property Get GameStrategyName() As String
    GameStrategyName = m.strStrategy
End Property

Public Property Let GameStrategyName(ByVal strName$)
    m.strStrategy = strName
End Property

Private Sub Class_Initialize()
    ResetGameMode
End Sub

Private Sub Class_Terminate()
    
    Set m.frmOwner = Nothing
    Set m.OrdersOpen = Nothing
    Set m.OrdersFilled = Nothing
    
    Set m.TargetOrder = Nothing
    Set m.StopLossOrder = Nothing
    Set m.BaseOrder = Nothing
    
    Set m.GameSystem = Nothing

End Sub

Public Property Get AutoTargetDollars() As Double
    AutoTargetDollars = m.dAutoTargetDollars
End Property

Public Property Get AutoTargetPoints() As Double
    AutoTargetPoints = m.dAutoTargetPoints
End Property

Public Property Get AutoStopDollars() As Double
    AutoStopDollars = m.dAutoStopDollars
End Property

Public Property Get AutoStopPoints() As Double
    AutoStopPoints = m.dAutoStopPoints
End Property

Public Property Get TargetDollars()
    TargetDollars = m.dTargetDollars
End Property

Public Property Get TargetPoints()
    TargetPoints = m.dTargetPoints
End Property

Public Property Get StopDollars()
    StopDollars = m.dStopDollars
End Property

Public Property Get StopPoints()
    StopPoints = m.dStopPoints
End Property

Public Property Get HasStop() As Long
    HasStop = m.nHasStop
End Property

Public Property Let HasStop(ByVal nStop&)
    m.nHasStop = nStop
End Property

Public Property Get HasTarget() As Long
    HasTarget = m.nHasTarget
End Property

Public Property Let HasTarget(ByVal nTarget&)
    m.nHasTarget = nTarget
End Property

Private Sub CalcAutoExits()

    Dim Bars As cGdBars
    Dim dAvg#, dDollars#, dTicks#
    Dim i&, j&, k&
    
    Set Bars = m.frmOwner.Chart.Bars
    If Bars Is Nothing Then Exit Sub
    
    j = m.frmOwner.Chart.LastGoodDataBar(False) - 31
    If j < 0 Then j = 0
    
    For i = m.frmOwner.Chart.LastGoodDataBar(False) To j Step -1
        dAvg = dAvg + (Bars(eBARS_High, i) - Bars(eBARS_Low, i))
        k = k + 1
    Next
    
    If k > 0 Then
        dAvg = dAvg / k
    
        dDollars = dAvg * Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
        dTicks = RoundNum(dDollars / Bars.Prop(eBARS_TickValue), 0) * 2
            
        'convert to points
        m.dAutoTargetPoints = dTicks * Bars.Prop(eBARS_TickMove)
        m.dAutoStopPoints = dTicks * Bars.Prop(eBARS_TickMove)
        
        'make dollar values match points
        m.dAutoTargetDollars = RoundToSigDigits(dTicks * Bars.Prop(eBARS_TickValue), 2)
        m.dAutoStopDollars = RoundToSigDigits(dTicks * Bars.Prop(eBARS_TickValue), 2)
    End If

End Sub

Public Sub ExitPosition(ByVal strPos$, Optional ByVal bReverse As Boolean = False)
On Error GoTo ErrSection:

    Dim LastFilledOrd As cPtOrder
    Dim Order As New cPtOrder
    Dim strPosition$, nQty&, bBuy As Boolean
    
    Set LastFilledOrd = m.OrdersFilled(m.OrdersFilled.Count)
    If LastFilledOrd Is Nothing Then Exit Sub
    
    strPosition = UCase(Parse(strPos, " ", 1))
    nQty = ValOfText(Parse(strPos, " ", 2))
        
    If nQty <= 0 Then Exit Sub
    
    If strPosition = "SHORT" Then
        bBuy = True
    ElseIf strPosition = "LONG" Then
        bBuy = False
    Else
        Exit Sub
    End If
    
    'Order.Symbol = LastFilledOrd.Symbol
    'Order.SymbolID = LastFilledOrd.SymbolID
    Order.SymbolOrSymbolID = LastFilledOrd.SymbolOrSymbolID
    
    Order.Quantity = nQty
    Order.Enter = False
    Order.Buy = bBuy
        
    Order.OrderType = eTT_OrderType_Market
    GameNewOrder Order
    
    RemoveAutoExits True, True
    
    If bReverse Then
        Set Order = New cPtOrder
        'Order.Symbol = LastFilledOrd.Symbol
        'Order.SymbolID = LastFilledOrd.SymbolID
        Order.SymbolOrSymbolID = LastFilledOrd.SymbolOrSymbolID
        Order.Quantity = nQty
        Order.Buy = bBuy
        Order.Enter = True
        Order.OrderType = eTT_OrderType_Market
        GameNewOrder Order
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "cGameMode.ExitPosition"
    
End Sub

Public Sub CalcNewTargetValues(ByVal dPrice#)

    Dim dTicks#
  
    m.dTargetPoints = Abs(m.BaseOrder.Fees - dPrice)
    dTicks = m.dTargetPoints / m.frmOwner.Chart.Bars.Prop(eBARS_TickMove)
    m.dTargetDollars = RoundToSigDigits(dTicks * m.frmOwner.Chart.Bars.Prop(eBARS_TickValue))
 
End Sub

Public Sub CalcNewStopValues(ByVal dPrice#)

    Dim dTicks#
  
    m.dStopPoints = Abs(m.BaseOrder.Fees - dPrice)
    dTicks = m.dStopPoints / m.frmOwner.Chart.Bars.Prop(eBARS_TickMove)
    m.dStopDollars = RoundToSigDigits(dTicks * m.frmOwner.Chart.Bars.Prop(eBARS_TickValue))
 
End Sub

Private Sub SetTblEntryFields(Order As cPtOrder, Bars As cGdBars, _
    tbReport As cGdTable, ByVal strPosition$)

    Dim dDate#, dPrice#, nIdx&
    Dim nTDOW&, nTDOM&, nTDOY&
    Dim hBarsDateTime&, iPos&

    dDate = Order.StatusDate
    dPrice = Order.Fees
    nIdx = tbReport.NumRecords - 1
    'find bars number to get TDOW, TDOM, TDOY
    hBarsDateTime = Bars.ArrayHandle(eBARS_DateTime)
    If Not Bars.IsIntraday Then
        If gdBinarySearch(hBarsDateTime, Int(dDate), iPos, eGdSort_Default, 0, -1) Then
            If Bars(eBARS_DateTime, iPos) = Int(dDate) Then
                nTDOW = m.frmOwner.Chart.TDOW(iPos)
                nTDOM = m.frmOwner.Chart.TDOM(iPos)
                nTDOY = m.frmOwner.Chart.TDOY(iPos)
            End If
        End If
        If nTDOW < 0 Then nTDOW = 0
        If nTDOM < 0 Then nTDOM = 0
        If nTDOY < 0 Then nTDOY = 0
    End If
                    
    tbReport(1, nIdx) = strPosition
    tbReport(2, nIdx) = dDate
    tbReport(3, nIdx) = dPrice
    tbReport(4, nIdx) = "Replay"
    tbReport(14, nIdx) = nTDOW
    tbReport(15, nIdx) = nTDOM
    tbReport(16, nIdx) = nTDOY
    tbReport(18, nIdx) = Order.OrderID
    tbReport(19, nIdx) = Order.Quantity

End Sub

Private Sub SetTblExitFields(OrderE As cPtOrder, OrderX As cPtOrder, Bars As cGdBars, _
    tbReport As cGdTable, ByVal strPosition$)

    Dim dDate#, dPrice#, nQty&, nIdx&
    Dim dMaxProfit#, dMaxLoss#, dProfitLoss#
    Dim nBarsInTrade&, nBarsEntryIdx&, nBarsExitIdx&

    dDate = OrderX.StatusDate
    dPrice = OrderX.Fees
    nQty = OrderX.Quantity
    nIdx = tbReport.NumRecords - 1
    
    If strPosition = "L" Then
        dProfitLoss = OrderX.Fees - OrderE.Fees
    Else
        dProfitLoss = OrderE.Fees - OrderX.Fees
    End If
    dProfitLoss = dProfitLoss * nQty * Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
    
    nBarsInTrade = BarsInTrade(Bars, nBarsEntryIdx, nBarsExitIdx, OrderE.StatusDate, OrderX.StatusDate)
    If nBarsEntryIdx >= 0 And nBarsExitIdx >= 0 Then
        If nBarsExitIdx > nBarsEntryIdx Then
            MaxProfitLoss Bars, dMaxProfit, dMaxLoss, nBarsEntryIdx, nBarsExitIdx - 1, nQty, strPosition, OrderE.Fees
        Else
            MaxProfitLoss Bars, dMaxProfit, dMaxLoss, nBarsEntryIdx, nBarsExitIdx, nQty, strPosition, OrderE.Fees
        End If
    End If
    
    tbReport(5, nIdx) = dDate
    tbReport(6, nIdx) = dPrice
    tbReport(7, nIdx) = "Replay"
    tbReport(8, nIdx) = dProfitLoss
    tbReport(9, nIdx) = RoundNum(dMaxProfit, 2)
    tbReport(10, nIdx) = RoundNum(dMaxLoss, 2)
    tbReport(11, nIdx) = nBarsInTrade
    tbReport(18, nIdx) = 0
    tbReport(19, nIdx) = 0

End Sub

Public Property Get HasResultFiles() As Boolean
    HasResultFiles = m.bResultFiles
End Property

Public Property Get ResultFilePath() As String
    ResultFilePath = App.Path & kResultsPath
End Property

Public Function SaveGameResult() As String
On Error GoTo ErrSection:

    Dim strFrom$, strTo$       'file names
    Dim strMsg$, i&
    Dim aFile As New cGdArray
    Dim dDate As Double
    
    strFrom = GameReport
    aFile.FromFile strFrom
    
    If aFile.Size > 0 Then
        strMsg = "Would you like to save the result of your game?"
        strMsg = InfBox(strMsg, "?", "+Save|-Cancel", "Game Result")
        If strMsg = "S" Then
            'set default file name
            strTo = Format(Now, "YYYYMMDD")
            For i = 65 To 90
                strMsg = App.Path & kResultsPath & strTo & "_" & Chr(i) & ".txt"
                If Not FileExist(strMsg) Then
                    Exit For
                End If
            Next
            
            If i <= 90 Then
                strTo = strTo & "_" & Chr(i)
            Else
                strTo = ""
            End If
            
            strMsg = InfBox("Enter name for game result.", "?", , "Game Result", , , , , , "strTo", strTo)
            strMsg = Trim(strMsg)
            If Len(strMsg) > 0 Then
                i = InStr(strMsg, ".")      'remove period so can add file extension
                If i > 0 Then
                    strMsg = Left(strMsg, i - 1)
                End If
                strTo = App.Path & kResultsPath & strMsg & ".txt"
                aFile.ToFile (strTo)
                SaveGameResult = strMsg
                ShowGameResult
            End If
        End If
    End If
    
    Set aFile = Nothing
    
    Exit Function
    
ErrSection:
    RaiseError "cGameMode.SaveGameResult"

End Function

Public Sub ShowGameResult()
On Error GoTo ErrSection:
   
    Dim strFile$, i&, rc&, s$
    Dim aResultFiles As New cGdArray
    Dim aTrades As New cGdArray
    Dim aParms As New cGdArray
    
    If Not m.bShowResults Then Exit Sub
    
    aResultFiles.Size = 0
    
#If 0 Then
    Dim fs As Object
    Dim F As Folder
    Dim f1 As File, fc As Files
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFolder(App.Path & kResultsPath)
    Set fc = F.Files
    For Each f1 In fc
        aResultFiles.Add App.Path & kResultsPath & f1.Name
        aParms.Add vbTab & f1.Name & vbTab & "0.00000" & vbTab & "false"
    Next
#Else
    aResultFiles.GetMatchingFiles App.Path & kResultsPath & "*.*"
    For i = 0 To aResultFiles.Size - 1
        s = aResultFiles(i)
        s = Mid(s, Len(FilePath(s)) + 1)
        aParms.Add vbTab & s & vbTab & "0.00000" & vbTab & "false"
    Next
#End If
    
    If aResultFiles.Size > 0 Then
        If frmOptimizer.Init(aResultFiles.Size, Nothing, eGDOptMode_GameMode, Me) Then
            For i = 0 To aResultFiles.Size - 1
                strFile = aResultFiles(i)
                aTrades.Size = 0
                aTrades.FromFile strFile
                CheckTradeFile aTrades, strFile
                rc = frmOptimizer.Add(i + 1, aTrades, aParms)
                
                If rc = kSN_OPTIMIZATION_ERROR Then
                    InfBox "There was an error processing result file " & strFile & ".", "E", , "Instant Replay Results"
                    Exit For
                End If
                
            Next
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cGameMode.ShowGameResults"

End Sub

Public Function RenameResult(ByVal strName$) As String
On Error GoTo ErrSection:

    Dim strOld$, strNew$, strSave$
    Dim aFile As New cGdArray
        
    strOld = strName
    If InStr(strOld, ".txt") = 0 Then strOld = strOld & ".txt"
    strOld = App.Path & kResultsPath & strOld
    
    If Not FileExist(strOld) Then
        InfBox "The game result " & strName & " was not found.", "E", "Game Results"
        Exit Function
    End If
    
    strNew = InfBox("Enter name for game result.", "?", , "Game Result", , , , , , "strName", strName)
    strNew = Trim(strNew)
    
    If Len(strNew) > 0 Then
        If InStr(strNew, ".txt") = 0 Then strNew = strNew & ".txt"
        aFile.FromFile strOld
        aFile.ToFile App.Path & kResultsPath & strNew
        If FileExist(App.Path & kResultsPath & strNew) Then
            KillFile strOld
        End If
    End If
    
    Set aFile = Nothing
    RenameResult = strNew
    
    Exit Function
    
ErrSection:
    RaiseError "cGameMode.RenameResult"
    
End Function

Public Sub DeleteResult(ByVal strName$)
On Error GoTo ErrSection:

    Dim strFile$
    
    strFile = strName
    If InStr(strFile, ".txt") = 0 Then
        strFile = strFile & ".txt"
    End If
    strFile = App.Path & kResultsPath & strFile
    
    KillFile strFile
    ShowGameResult          'issue 6475
    
    Exit Sub

ErrSection:
    RaiseError "cGameMode.DeleteResult"

End Sub

Public Sub ShowGameReport(ByVal strFile$)
On Error GoTo ErrSection:

    Dim i&
    
    i = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If FileExist(strFile) Then
        If m.GameSystem Is Nothing Then
            Set m.GameSystem = New cSystem
            m.GameSystem.Load 56        'this is the sample turtle strategy
        End If
        m.GameSystem.ShowReports , strFile, True
    Else
        InfBox strFile, "E", , "Instant Replay Report"
    End If
    
    Screen.MousePointer = i
    
    Exit Sub

ErrSection:
    RaiseError "cGameMode.ShowGameReport"
    
End Sub

Public Property Get CustomOrders() As Long
On Error GoTo ErrSection:
    
    Dim nCount&
    
    nCount = m.OrdersFilled.Count + m.OrdersOpen.Count
    
    CustomOrders = nCount
    
    Exit Property

ErrSection:
    RaiseError "cGameMode.CustomOrders", eGDRaiseError_Default

End Property

Public Property Get ShowResultsFlag() As Boolean
    ShowResultsFlag = m.bShowResults
End Property

Public Property Let ShowResultsFlag(ByVal bShow As Boolean)
    m.bShowResults = bShow
End Property

Public Property Get ReplayAll() As Boolean
    ReplayAll = m.bReplayAll
End Property

Public Property Let ReplayAll(ByVal bAll As Boolean)
    m.bReplayAll = bAll
End Property

Public Sub ClearReplayAll()
On Error GoTo ErrSection:

    Dim i&
    Dim frm As Form
    
    If m.bReplayAll Then
        For i = 0 To Forms.Count - 1
            If IsFrmChart(Forms(i)) Then
                Set frm = Forms(i)
                frm.ClearReplaySync
            End If
        Next
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cGameMode.ClearReplayAll"
    
End Sub

Private Function TradingSession(Bars As cGdBars, ByVal dDateTime#) As Long
On Error GoTo ErrSection:
    
    Dim nSession&
    Dim dCrossOver#
    
    dCrossOver = Bars.Prop(eBARS_CrossoverTime)
    If dCrossOver > 0 Then
        nSession = Int(dDateTime + (1 - dCrossOver / 1440#))
    Else
        nSession = Int(dDateTime)
    End If

    TradingSession = nSession
    
    Exit Function

ErrSection:
    RaiseError "cGameMode.TradingSession"
    
End Function

Private Sub CheckTradeFile(aData As cGdArray, ByVal strFile$)
On Error GoTo ErrSection:

'aardvark 2761 fix:
'create two different "bogus" rule ID's -1 and -2 that can be put
'into the reporst lookup table as "Manual Entry" and "Manual Exit"

    Dim i&, bConvert As Boolean
    Dim aFields As New cGdArray
    
    If aData Is Nothing Then Exit Sub
    
    For i = 1 To aData.Size - 1
        aFields.SplitFields aData(i), vbTab
        If aFields(11) <> -1 Or aFields(12) <> -2 Then
            bConvert = True
            aFields(11) = -1
            aFields(12) = -2
            aData(i) = aFields.JoinFields(vbTab)
        End If
    Next
    
    If bConvert And Len(strFile) > 0 Then aData.ToFile strFile
    
    Exit Sub
    
ErrSection:
    RaiseError "cGameMode.CheckTradeFile"
    
End Sub

Public Property Get KeepAtEndOfBar() As Boolean
    KeepAtEndOfBar = m.bEndOfBar
End Property

Public Property Get HasIntradayTicks() As Boolean
    HasIntradayTicks = m.bHasIntradayTicks
End Property

Public Property Get MidSessionTime() As Double
On Error GoTo ErrSection:

    Dim dMinutes#
    
    If m.dSessionEnd > m.dSessionStart Then
        dMinutes = (m.dSessionEnd - m.dSessionStart) * 1440#
        dMinutes = m.dSessionStart + dMinutes / 2880 'add half of total minutes in session to session start
        MidSessionTime = Int(m.dDataTime) + dMinutes
    Else
        MidSessionTime = m.frmOwner.Chart.Bars.SessionDateForTime(m.dDataTime, False) 'just use midnight of new day for overnight symbols
    End If
    
    Exit Property
        
ErrSection:
    RaiseError "cGameMode.MidSessionTime"
    
End Property

Private Sub CheckForTicks()
On Error GoTo ErrSection:

    Dim Bars As cGdBars, SessionTicks As cGdBars
    Dim strSymbol$, i&, nBarInterval&
    Dim bEvening As Boolean
        
    If m.frmOwner Is Nothing Then Exit Sub
    If m.frmOwner.Chart Is Nothing Then Exit Sub
    Set Bars = m.frmOwner.Chart.Bars
    If Bars Is Nothing Then Exit Sub
    
    If Bars.Prop(eBARS_PeriodType) <> ePRD_Days And (m.dUserInterval = 0 Or m.dUserInterval >= 1) Then
        m.bHasIntradayTicks = True  'don't care, just set to true
    Else
        strSymbol = Bars.Prop(eBARS_Symbol)
        If m.nTickFirstDate = 0 Then m.nTickFirstDate = g.SymbolPool.TickFirstDate(strSymbol)
            
        If Int(m.dDataTime) < m.nTickFirstDate Then
            m.bHasIntradayTicks = False
        Else
            'Fix for aardvark 3789:
            'If daily bar then make sure game interval is >= 1 day if there are no intraday ticks
            Set SessionTicks = New cGdBars
            
            SetBarProperties SessionTicks, strSymbol
            i = DM_GetBars(SessionTicks, strSymbol, ePRD_EachTick, Int(m.dDataTime), Int(m.dDataTime))
            If SessionTicks.Size > 0 Then
                m.bHasIntradayTicks = True
            Else
                m.bHasIntradayTicks = False
            End If
            
            Set SessionTicks = Nothing
        End If
    End If
        
    ' interval to add during replay
    If Not m.bHasIntradayTicks Then
        If m.dInterval < 1 Then m.dInterval = 1
    ElseIf m.bAutoInterval Then
        nBarInterval = 1441
        'get correct nBarInterval
        CheckSessionTimes m.frmOwner, nBarInterval, bEvening
        If nBarInterval < 30 Then
            m.dInterval = 1
        ElseIf nBarInterval < 1440 Then
            m.dInterval = 5
        ElseIf nBarInterval = 1440 Then
            m.dInterval = 30
        Else
            m.dInterval = 1440
        End If
        m.dInterval = gdFixDateTime(m.dInterval / 1440)
    ElseIf m.dUserInterval <> 0 Then
        m.dInterval = m.dUserInterval
    End If
    
    Exit Sub
        
ErrSection:
    RaiseError "cGameMode.CheckForTicks"

End Sub

Private Sub FixInterval()
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    
    If m.bAutoInterval Then Exit Sub
    If m.frmOwner Is Nothing Then Exit Sub
    If m.frmOwner.Chart Is Nothing Then Exit Sub
    Set Bars = m.frmOwner.Chart.Bars
    If Bars Is Nothing Then Exit Sub
    'If Bars.Prop(eBARS_Periodicity) <> ePRD_Days + 1 Then Exit Sub
    
    'saved as auto, minutes, or day in INI file
    m.strReplaySpeed = GetIniFileProperty("ReplaySpeed", "auto", "Game Mode", g.strIniFile)
    
    'fix for 4094 - custom interval of one day does not work the same of different periodicity & symbol types
    If m.strReplaySpeed = "day" Then
        If Bars.Prop(eBARS_StartTime) > Bars.Prop(eBARS_EndTime) And Not IsForex(Bars.Prop(eBARS_Symbol)) Then
            m.dInterval = 1440 - (Bars.Prop(eBARS_StartTime) - Bars.Prop(eBARS_EndTime))
        ElseIf Bars.IsIntraday Then
            'do this so the interval will not end up being 1 (ie 1440/1440)
            'SpliceReplayTicks loads entire day if game interval >= 1
            m.dInterval = 1439
        Else
            m.dInterval = 1440
        End If
    End If
    
    Exit Sub
        
ErrSection:
    RaiseError "cGameMode.FixInterval"

End Sub

Public Sub FormOwnerChange(frmNewOwner As Form)
On Error GoTo ErrSection:

    Set m.frmOwner = Nothing        'to prevent anything else from executing from here on out
    
    If frmNewOwner Is Nothing Then Exit Sub
    If Not IsFrmChart(frmNewOwner) Then Exit Sub
    
    Set m.frmOwner = frmNewOwner
    
    Exit Sub

ErrSection:
    RaiseError "cGameMode.FormOwnerChange"

End Sub

Public Property Get UserForwardBack() As Long
    UserForwardBack = m.nUserForwardBack
End Property

Public Property Let UserForwardBack(ByVal nAction&)

    If nAction = 0 Or nAction = 1 Or nAction = -1 Then
        m.nUserForwardBack = nAction
    End If

End Property
