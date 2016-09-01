Attribute VB_Name = "mProfileBars"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mProfileBars.bas
'' Description: module for managing profile bars for market profile display style
''              on charts & Trade Profile form
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 10/19/2012   MJM         Initial create/addition to VB project
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private treeProfileBars As cGdTree          'tree of profile bars
Private treeHistory As cGdTree
Private tbUsage As cGdTable                 'table of usage count for destroying bars (0 usage = ok to destroy)

Public Enum eProfileStartusRT
    eProfileRT_NewSession = -1
    eProfileRT_InProgress = -2
    eProfileRT_NoChange = 0
    eProfileRT_NewData = 1
End Enum

Public Function VolumeBarsGet(Chart As cChart, Ind As cIndicator, _
    BarsProfileAlign As cGdBars, BarsProfileTicks As cGdBars, ByVal bInit As Boolean) As cGdBars
On Error GoTo ErrSection:

'JM 07-16-2015: data retrieval & storage
'   - retrieve most recent 10 sessions of profile bars on initial request and add to treeProfileBars
'   - previous sessions are loaded one at a time as triggered by form's timer and added to treeHistory
'   - when all previous sessions are loaded move them from treeHistory to treeProfileBars
    
    Dim i&, d#
    Dim Bars As cGdBars
    Dim BarsProfileNew As cGdBars
    
    If Chart Is Nothing Or Ind Is Nothing Or BarsProfileAlign Is Nothing Then Exit Function
    If BarsProfileAlign.Size = 0 Then Exit Function
    
    If Not treeProfileBars Is Nothing Then
        For i = 1 To treeProfileBars.Count
            Set Bars = treeProfileBars(i)
            If Not Bars Is Nothing Then
                If Bars.Prop(eBARS_Symbol) = Chart.Symbol Then
                    If Bars.Prop(eBARS_PeriodicityStr) = Chart.Bars.Prop(eBARS_PeriodicityStr) Then
                        Exit For
                    End If
                End If
                Set Bars = Nothing
            End If
        Next
    End If
    
    If Bars Is Nothing Then
        Set Bars = New cGdBars
        SetBarProperties Bars, Chart.Symbol
        Bars.Prop(eBARS_PeriodicityStr) = Chart.Bars.Prop(eBARS_PeriodicityStr)
    End If

    i = Chart.Bars.SessionDate(Chart.Bars.Size - 1) - Chart.Bars.SessionDate(0)
    If i > 10 And Bars.Size = 0 Then
        d = Chart.Bars.SessionDate(Chart.Bars.Size - 1) - 10
        i = BarsProfileAlign.FindDateTime(d)
        d = BarsProfileAlign.SessionDate(i)
        i = Chart.Bars.FindDateTime(d)
        While Chart.Bars.SessionDate(i) >= d
            i = i - 1   'get to beginning of session profile
        Wend
        If Chart.Bars.SessionDate(i) < d Then i = i + 1
        d = Chart.Bars.SessionDate(i)
    ElseIf Bars.SessionDate(0) > Chart.Bars.SessionDate(0) Then
        If Not Chart.Form.tmrProfileLoad.Enabled Then
            If treeHistory Is Nothing Then Set treeHistory = New cGdTree
            For i = 1 To treeHistory.Count
                Set BarsProfileNew = treeHistory(i)
                If Not BarsProfileNew Is Nothing Then
                    If BarsProfileNew.Prop(eBARS_SymbolID) = Bars.Prop(eBARS_SymbolID) Then
                        If BarsProfileNew.Prop(eBARS_PeriodicityStr) = Bars.Prop(eBARS_PeriodicityStr) Then
                            Exit For
                        Else
                            Set BarsProfileNew = Nothing
                        End If
                    Else
                        Set BarsProfileNew = Nothing
                    End If
                End If
            Next
            
            If BarsProfileNew Is Nothing Then
                Set BarsProfileNew = Bars.MakeCopy(True)
                treeHistory.Add BarsProfileNew
            End If
            
            If Not BarsProfileNew Is Nothing Then Chart.Form.tmrProfileLoad.Enabled = True
        End If
    End If
    
    If Bars.Size = 0 Then
        BuildProfileBars Bars, Chart.SymbolID, d, Chart.Bars.SessionDate(Chart.Bars.Size - 1)
    ElseIf Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False)) > Bars(eBARS_DateTime, Bars.Size - 1) Then
        'true when streaming just turned on
        i = Bars.Size
        BuildProfileBars Bars, Chart.SymbolID, 0, 0
        If i = Bars.Size Then
            For i = Bars.Size - 1 To 0 Step -1
                If Chart.Bars.SessionDate(Chart.LastGoodDataBar(False)) = Bars.SessionDate(i) Then
                    Bars.DeleteSomeBars i
                Else
                    Exit For
                End If
            Next
            BuildProfileBars Bars, Chart.SymbolID, 0, 0
        End If
    ElseIf g.RealTime.Active Then
        If BarsProfileTicks Is Nothing Then
            Set BarsProfileTicks = New cGdBars
        ElseIf BarsProfileTicks.Size = 0 Then
            GetAvailTickData BarsProfileTicks, i, Chart.Symbol, Chart.SymbolID, 0, 0
        Else
            Dim iPrevSize&, bNewBar As Boolean, bNewData As Boolean
            iPrevSize = BarsProfileTicks.Size
            bNewData = g.RealTime.UpdateBars(BarsProfileTicks, bNewBar)
            i = BarsProfileTicks.Size - iPrevSize
            If i > 0 Then
                BarsProfileTicks.DeleteFirstBars iPrevSize
                Set BarsProfileNew = Bars.MakeCopy(True)
                If BarsProfileNew.BuildBars(Bars.Prop(eBARS_PeriodicityStr), BarsProfileTicks.BarsHandle) Then
                    MergeBarsProfile BarsProfileNew, Bars
                End If
            End If
        End If
    End If
    
    If BarsProfileAlign.SessionDate(BarsProfileAlign.Size - 1) < Chart.Bars.SessionDate(Chart.LastGoodDataBar(False)) Then
        BarsProfileAlign.AddForecastBars (1)
    End If
    
    Set VolumeBarsGet = Bars
    If bInit Then UsageIncrement Bars
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mProfileBars.VolumeBarsGet"

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JM: 10-22-2012 design notes
'
' This module maintains a tree of Profile bars for displaying market profile data.
'
' Usage:
'   1) call ProfileBarsGet to get profile bars with historical data
'   2) call ProfileUpdateRT to update profile bars with streaming data
'
'   IF ProfileBarsGet was called with bCreate = TRUE then
'       3) call ProfileBarsFree when done allowing bars to be destroyed
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function ProfileBarsGet(ByVal nSymbolID&, _
    Optional ByVal iSessionStart As Long = -1, _
    Optional ByVal iSessionEnd As Double = -1, _
    Optional ByVal bCreate As Boolean = False, _
    Optional frmNotify As Form = Nothing) As cGdBars
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim BarsAvail As cGdBars
    Dim bCreateOk As Boolean
    
    Dim i&, iChunk&, iSize&, iStart&, iEnd&
    Dim dtTimeout#
    
    If bCreate Then
        If iSessionStart > 0 And iSessionEnd > 0 And iSessionEnd >= iSessionStart Then bCreateOk = True
    End If

    If treeProfileBars Is Nothing Then
        If Not bCreateOk Then Exit Function
    Else
        For i = 1 To treeProfileBars.Count
            Set Bars = treeProfileBars(i)
            If Not Bars Is Nothing Then
                If Bars.Prop(eBARS_SymbolID) = nSymbolID Then
                    If Bars.Prop(eBARS_PeriodicityStr) = "5 minute" Then
                        Exit For
                    End If
                End If
                Set Bars = Nothing
            End If
        Next
    End If
    
    iChunk = 2
    Dim SymInf As cSymbolInfo
    
    Set BarsAvail = New cGdBars         'aardvark 7003
    SetBarProperties BarsAvail, nSymbolID
    BarsAvail.ArrayMask = eBARS_Eod Or eBARS_BidAsk
    If DM_GetBars(BarsAvail, nSymbolID, "Daily", 0, 0) Then
    If g.RealTime.Active Then
        g.RealTime.SpliceBars BarsAvail
    End If
    End If
    If BarsAvail.Size > 0 Then
        iSessionEnd = BarsAvail.SessionDate(BarsAvail.Size - 1, True)
    End If
    
    If Bars Is Nothing Then
        If bCreateOk Then
            Set Bars = New cGdBars
            'JM 10/17/2012: copied code (see TLB note in frmMarketProfile.InitialShow)
            dtTimeout = gdTickCount + 10000
            If g.RealTime.Active And g.RealTime.SalmonIsRunning Then
                Set SymInf = g.RealTime.SymbolInfo(nSymbolID)
                Do While SymInf.GetDataRequestStatus(ePRD_EachTick) = eSalmonPending And gdTickCount < dtTimeout
                    DoEvents
                Loop
            End If
            
            For i = iSessionStart To iSessionEnd
                iEnd = i + iChunk
                If iEnd > iSessionEnd Then iEnd = iSessionEnd
                
                iSize = Bars.Size
                Bars.Prop(eBARS_PeriodicityStr) = "5 minute"
                BuildProfileBars Bars, nSymbolID, i, iEnd
                
                If frmNotify Is Nothing Then
                    DoEvents
                Else
                    frmNotify.ProfileBarsGetStatus "Profile bars retrieved: " & _
                            DateFormat(Bars(eBARS_DateTime, iSize), MM_DD_YYYY, HH_MM_SS) & " - " & _
                            DateFormat(Bars(eBARS_DateTime, Bars.Size - 1), MM_DD_YYYY, HH_MM_SS)
                End If
                i = iEnd
            Next
        End If
    Else
        iStart = Bars.SessionDateForTime(Bars(eBARS_DateTime, 0), False)
        If iStart > iSessionStart Then
            For iEnd = iStart To iSessionStart Step -1
                i = iEnd - iChunk
                If i < iSessionStart Then i = iSessionStart
                
                iSize = Bars.Size
                BuildProfileBars Bars, nSymbolID, i, iEnd
                
                iSize = Bars.Size - iSize - 1
                
                If frmNotify Is Nothing Then
                    DoEvents
                Else
                    frmNotify.ProfileBarsGetStatus "Profile bars built: " & _
                            DateFormat(Bars(eBARS_DateTime, 0), MM_DD_YYYY, HH_MM_SS) & " - " & _
                            DateFormat(Bars(eBARS_DateTime, iSize), MM_DD_YYYY, HH_MM_SS)
                End If
                iEnd = i
            Next
        End If
        
        iEnd = Bars.SessionDateForTime(Bars(eBARS_DateTime, Bars.Size - 1), False)
        If iEnd < iSessionEnd Then
            For i = iEnd To iSessionEnd
                iEnd = i + iChunk
                If iEnd > iSessionEnd Then iEnd = iSessionEnd
                
                iSize = Bars.Size
                BuildProfileBars Bars, nSymbolID, i, iEnd
                
                If frmNotify Is Nothing Then
                    DoEvents
                Else
                    frmNotify.ProfileBarsGetStatus "Profile bars built: " & _
                            DateFormat(Bars(eBARS_DateTime, iSize), MM_DD_YYYY, HH_MM_SS) & " - " & _
                            DateFormat(Bars(eBARS_DateTime, Bars.Size - 1), MM_DD_YYYY, HH_MM_SS)
                End If
                i = iEnd
            Next
        End If
    End If
    
    Set ProfileBarsGet = Bars
    If bCreate Then UsageIncrement Bars

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mProfileBars.ProfileBarsGet"

End Function

Public Sub ProfileBarsFree(Bars As cGdBars)
    UsageDecrement Bars
End Sub

Private Sub UsageIncrement(Bars As cGdBars)
On Error GoTo ErrSection:

    Dim i&, iPos&, iCount&
    Dim aUsageIdx As cGdArray

    If Bars Is Nothing Then Exit Sub

    If treeProfileBars Is Nothing Then Set treeProfileBars = New cGdTree
    
    If tbUsage Is Nothing Then
        Set tbUsage = New cGdTable
        tbUsage.CreateField eGDARRAY_Longs, 0, "SymbolID", 0
        tbUsage.CreateField eGDARRAY_Longs, 1, "UsageCount", 0
        tbUsage.CreateField eGDARRAY_Longs, 2, "Periodicity", 0
    End If

    iCount = 0
    For iPos = 1 To treeProfileBars.Count
        If Not treeProfileBars(iPos) Is Nothing Then
            If treeProfileBars(iPos).Prop(eBARS_Symbol) = Bars.Prop(eBARS_Symbol) Then
                If treeProfileBars(iPos).Prop(eBARS_PeriodicityStr) = Bars.Prop(eBARS_PeriodicityStr) Then
                    iCount = 1
                    Exit For
                End If
            End If
        End If
    Next
    If iCount = 0 Then
        treeProfileBars.Add Bars
        tbUsage.AddRecord ""
        iPos = tbUsage.NumRecords - 1
        tbUsage(0, iPos) = Bars.Prop(eBARS_SymbolID)
        tbUsage(1, iPos) = 1
        tbUsage(2, iPos) = GetPeriodicity(Bars.Prop(eBARS_PeriodicityStr))
    Else
        i = GetPeriodicity(Bars.Prop(eBARS_PeriodicityStr))
        For iCount = 0 To tbUsage.NumRecords - 1
            If tbUsage(0, iCount) = Bars.Prop(eBARS_SymbolID) Then
                If tbUsage(2, iCount) = i Then
                    tbUsage(1, iCount) = tbUsage(1, iCount) + 1
                    Exit For
                End If
            End If
        Next
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mProfileBars.UsageIncrement"

End Sub

Private Sub UsageDecrement(Bars As cGdBars)
On Error GoTo ErrSection:

    Dim i&, iCount&

    If Bars Is Nothing Then Exit Sub
    If treeProfileBars Is Nothing Then Exit Sub
    If tbUsage Is Nothing Then Exit Sub            'bogus request, just exit
    
    i = GetPeriodicity(Bars.Prop(eBARS_PeriodicityStr))
    For iCount = 0 To tbUsage.NumRecords - 1
        If tbUsage(0, iCount) = Bars.Prop(eBARS_SymbolID) Then
            If tbUsage(2, iCount) = i Then
                If tbUsage(1, iCount) > 0 Then
                    tbUsage(1, iCount) = tbUsage(1, iCount) - 1
                End If

'JM 06-01-2015: for now do not remove profile bars since user could be switching
'   between intraday periodicities and keeping the bars around gives much better
'   performance with neglible effect on memory usage
'
'                If tbUsage(1, iCount) = 0 Then
'                    For i = 1 To treeProfileBars.Count
'                        If treeProfileBars(i) Is Bars Then
'                            treeProfileBars.Remove i
'                            Exit For
'                        End If
'                    Next
'                End If
            End If
        End If
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mProfileBars.UsageDecrement"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JM: 10-22-2012 Parameters
'
'       [out]BarsTick:           bars of periodicity ePRD_EachTick from RT stream
'   [in][out]BarsProfile:        special profile bars that contains summary trade info at every price for 5min intervals
'       [out]BarsProfileNew:     changes from RT stream used to update the above BarsProfile
'        [in]nSessionDate:       session date of last available data caller has access to
'        [in]nSessionNext:       what caller expects to be the next session date
'        [in]nSessionMostRecent: what caller expects the most recent sessiond date to be
'        [in]bInit:              true to initialize the tick bars for RT updates
'
' The historical BarsProfile are created & populated only once. The various session dates
' passed in by different callers are used to return correct status to caller.
'
' Example: caller A creates BarsProfile & starts updating it with RT stream
'          caller B requests same BarsProfile minutes later
'          caller B will be notified that BarsProfile contains "New Day" data since
'               the nSessionDate, nSessionNext, nSessionMostRecent passed in by caller B
'               will be based on caller's B initial data load which will not yet have RT data
'
' NOTE; To update displays with RT data, the caller should use the [out] BarsProfileNew bars.
'       BarsProfileNew contains only changed data when this routine exits.
'
'       BarsProfile will have the BarsProfileNew changes merged in it as well as all history data.
'       In other words, it will not be obvious to caller, what has changed in BarsProfile without
'       the caller having some sort of before & after snapshots that can be used for comparison.
'       This routine is designed so that caller can simply use the [out]BarsProfileNew to see
'       what the changes are.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ProfileUpdateRT(BarsTick As cGdBars, _
    BarsProfile As cGdBars, _
    BarsProfileNew As cGdBars, _
    ByVal nSessionDate As Long, _
    ByVal nSessionNext As Long, _
    ByVal nSessionMostRecent As Long, _
    ByVal bInit As Boolean) As eProfileStartusRT
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    
    Dim iDontCare&, iPrevSize&
    Dim strSymbol$, nSymbolID&
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Function
    If BarsProfile Is Nothing Or BarsProfileNew Is Nothing Then Exit Function
    
    If bInProgress Then
        ProfileUpdateRT = eProfileRT_InProgress
        Exit Function
    End If
    
    strSymbol = BarsProfile.Prop(eBARS_Symbol)
    nSymbolID = BarsProfile.Prop(eBARS_SymbolID)
    If Len(strSymbol) = 0 Or nSymbolID = 0 Then Exit Function
    
'JM 06-01-2015: original code no longer needed - leave awhile then remove when all ok
'    If Not bInit Then
'        If BarsTick Is Nothing Then
'            Exit Function
'        ElseIf strSymbol <> BarsTick.Prop(eBARS_Symbol) Then
'            Exit Function
'        ElseIf nSymbolID <> BarsTick.Prop(eBARS_SymbolID) Then
'            Exit Function
'        End If
'    End If

    If Not g.RealTime.Active Then Exit Function
    
    Dim i&, j&, k&
    
    bInProgress = True
    
    If BarsTick Is Nothing Then
        Set BarsTick = New cGdBars
        GetAvailTickData BarsTick, iDontCare, strSymbol, nSymbolID, 0, 0
        If BarsTick(eBARS_DateTime, BarsTick.Size - 1) <= BarsProfile(eBARS_DateTime, BarsProfile.Size - 1) Then
            GoTo ErrExit
        Else
            iPrevSize = -1
        End If
    End If

    Dim bNewBar As Boolean
    Dim bNewData As Boolean
    
    If iPrevSize <> -1 Then
        iPrevSize = BarsTick.Size
        bNewData = g.RealTime.UpdateBars(BarsTick, bNewBar)
        
        If BarsTick.Prop(eBARS_CustomString) = "<PENDING>" Then
            DoEvents
            bNewData = g.RealTime.UpdateBars(BarsTick, bNewBar)
        End If
    End If
    
    If bNewBar Then
        i = nSessionNext
        j = nSessionMostRecent
        
        If i <= j Then
            'need to do this loop because overnight symbol will have an entire session that is available only from RT stream
            For k = i To j
                GetAvailTickData BarsTick, iDontCare, strSymbol, nSymbolID, k, 0
                Set BarsProfileNew = BarsProfile.MakeCopy(True)
                
                If BarsProfileNew.BuildBars(BarsProfile.Prop(eBARS_PeriodicityStr), BarsTick.BarsHandle) Then
                    'if user starts RT, stops RT then restart RT some of the bars in the stream will already
                    'be part of the profile bars so need to merge instead of just append
                    If BarsProfile.SessionDate(BarsProfile.Size - 1) < k Then
                        'JM 06-21-2015 - fixes bug where volume is re-added when streaming is off then back on
                        '   example: if volume POC is 100 on initial show with streaming on
                        '            turning streaming off then back on will cause POC volume to go to 200
                        '   this bug has always been there and not noticeable because the volume profile
                        '   are almost always displayed as percent by users (bug was discovered when testing
                        '   and comparing against the volume profile study feature)
                        '   Last Upgrade tested that has this bug is 6.9 build 1451 dated 5/13/2015
                        MergeBarsProfile BarsProfileNew, BarsProfile
                    End If
                End If
            Next
            ProfileUpdateRT = eProfileRT_NewSession
            GoTo ErrExit
        End If
    End If
    
    If iPrevSize <> -1 Then
        i = BarsTick.Size - iPrevSize
        If i > 0 Then
            BarsTick.DeleteFirstBars iPrevSize
        Else
            GoTo ErrExit
        End If
    End If

    Set BarsProfileNew = BarsProfile.MakeCopy(True)
    bNewData = BarsProfileNew.BuildBars(BarsProfile.Prop(eBARS_PeriodicityStr), BarsTick.BarsHandle)
    
    If Not bNewData Or BarsProfileNew.Size < 1 Then
        iPrevSize = 0       'should log err message here
    Else
        MergeBarsProfile BarsProfileNew, BarsProfile
        iPrevSize = BarsProfile.Size
        ProfileUpdateRT = eProfileRT_NewData
    End If

ErrExit:
    bInProgress = False
    Exit Function

ErrSection:
    BarsProfileNew.Size = 0
    bInProgress = False
    RaiseError "mProfileBars.ProfileUpdateRT"

End Function

Private Sub MergeBarsProfile(BarsSource As cGdBars, BarsDest As cGdBars)
On Error GoTo ErrSection:

    Dim i&, j&, idx&, Size&
    Dim dDateTime#
    
    Dim ProfileNotFound As cGdBars

    If BarsSource Is Nothing Then Exit Sub
    If BarsDest Is Nothing Then Exit Sub
    
    dDateTime = BarsSource(eBARS_DateTime, 0)
    
    If dDateTime > BarsDest(eBARS_DateTime, BarsDest.Size - 1) Then
        'this is all new data, no need to step through the bars
        gdAppendBars BarsDest.BarsHandle, BarsSource.BarsHandle, 0
    Else
        Set ProfileNotFound = BarsSource.MakeCopy
        Size = BarsDest.Size
        
        For i = BarsSource.Size - 1 To 0 Step -1
            dDateTime = BarsSource(eBARS_DateTime, i)
            idx = BarsDest.FindDateTime(dDateTime, True)
            
            If idx >= 0 And idx < BarsDest.Size Then
                For j = idx To Size - 1
                    If BarsSource(eBARS_Close, i) = BarsDest(eBARS_Close, j) Then
                        Exit For
                    End If
                Next
            
                If j >= 0 And j < Size Then
                    BarsDest(eBARS_Vol, j) = BarsDest(eBARS_Vol, j) + BarsSource(eBARS_Vol, i)
                    BarsDest(eBARS_BidVol, j) = BarsDest(eBARS_BidVol, j) + BarsSource(eBARS_BidVol, i)
                    BarsDest(eBARS_AskVol, j) = BarsDest(eBARS_AskVol, j) + BarsSource(eBARS_AskVol, i)
                    
                    ProfileNotFound.DeleteSomeBars i, 1
                Else
                    i = i
                End If
            Else
                j = j
            End If
        Next
        
        If ProfileNotFound.Size > 0 Then
            gdAppendBars BarsDest.BarsHandle, ProfileNotFound.BarsHandle, 0
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mProfileBars.MergeProfileBars"

End Sub

Public Sub ProfileVolHistory(Chart As cChart)
On Error GoTo ErrSection:

    If g.bStarting Or g.bLoadingChartPage Or g.bUnloading Then Exit Sub

    If Chart Is Nothing Then Exit Sub
    If Chart.Bars Is Nothing Then Exit Sub
    
    If treeHistory Is Nothing Or treeProfileBars Is Nothing Then
        Chart.Form.tmrProfileLoad.Enabled = False
        Exit Sub
    End If
    If treeHistory.Count = 0 Or treeProfileBars.Count = 0 Then
        Chart.Form.tmrProfileLoad.Enabled = False
        Exit Sub
    End If
    
    Dim i&, j&, iPos&, d#
    Dim Bars As cGdBars
    Dim ProfileBars As cGdBars
    
    For i = 1 To treeHistory.Count
        Set Bars = treeHistory(i)
        If Not Bars Is Nothing Then
            If Bars.Prop(eBARS_SymbolID) = Chart.Bars.Prop(eBARS_SymbolID) Then
                If Bars.Prop(eBARS_PeriodicityStr) = Chart.Bars.Prop(eBARS_PeriodicityStr) Then
                    iPos = i
                    Exit For
                End If
            End If
        End If
    Next
    
    If Bars Is Nothing Then
        Chart.Form.tmrProfileLoad.Enabled = False
        Exit Sub
    End If
    
    For i = 1 To treeProfileBars.Count
        Set ProfileBars = treeProfileBars(i)
        If Not ProfileBars Is Nothing Then
            If Bars.Prop(eBARS_SymbolID) = ProfileBars.Prop(eBARS_SymbolID) Then
                If Bars.Prop(eBARS_PeriodicityStr) = ProfileBars.Prop(eBARS_PeriodicityStr) Then
                    Exit For
                End If
            End If
        End If
    Next
    
    If ProfileBars Is Nothing Then
        Chart.Form.tmrProfileLoad.Enabled = False
        Exit Sub
    End If
    
    If Bars.Size = 0 Then
        d = ProfileBars.SessionDate(0)
    Else
        d = Bars.SessionDate(0)
    End If
    
    If Not Chart.Bars.IsIntraday Then
        StatusMsg ""
        'user changed from intraday to non-intraday chart before all profiles were loaded
        'turn off timer to stop loading profile for previous periodicity
        Chart.Form.tmrProfileLoad.Enabled = False
    ElseIf d > Chart.Bars.SessionDate(0) Then
        j = d
        d = d - 1
        While gdIsHoliday(d, "") Or Not IsWeekday(d)
            d = d - 1
        Wend
        If Chart.Form Is ActiveChart Then
            StatusMsg "Loading volume profile " & DateFormat(d)
        ElseIf InStr(ActiveChart.Caption, "Loading Profile") = 0 Then
            StatusMsg ""
        End If
        BuildProfileBars Bars, Chart.SymbolID, d, j
    Else
        StatusMsg ""
        gdAppendBars ProfileBars.BarsHandle, Bars.BarsHandle, 1
        treeHistory.Remove iPos
        Chart.Form.tmrProfileLoad.Enabled = False
        If Chart.Form Is ActiveChart Then Chart.GenerateChart eRedo1_Scrolled
    End If
    

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mProfileBars.ProfileVolHistory"

End Sub
