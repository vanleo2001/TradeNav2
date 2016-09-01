VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmMarketProfile 
   Caption         =   "Trade Profile"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbExgrid 
      Height          =   2055
      Left            =   600
      ScaleHeight     =   1995
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
      _cx             =   2143
      _cy             =   1085
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Timer tmr 
      Interval        =   125
      Left            =   4350
      Top             =   1785
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   4440
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   8
      DisplayContextMenu=   0   'False
      Tools           =   "frmMarketProfile.frx":0000
      ToolBars        =   "frmMarketProfile.frx":4D17
   End
End
Attribute VB_Name = "frmMarketProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kColorGreen = 32768                  'RGB(0,128,0)      - forest green
Private Const kColorPurple = 8388736               'RGB(128,0,128)    - weird purplish

Public Enum eProfileStatsType
    eProfileStats_VAVolume = 0
    eProfileStats_VATPO
    eProfileStats_StdDev
    eProfileStats_IB
    eProfileStats_POCVol
    eProfileStats_POCTPO
    eProfileStats_Mean
End Enum

Public Enum eProfilePropType
    eProfileProp_ColorScheme = 0
    eProfileProp_GradientFrom
    eProfileProp_GradientTo
    eProfileProp_UpColor
    eProfileProp_DownColor
    eProfileProp_BidColor
    eProfileProp_AskColor
    eProfileProp_SMPColor
    eProfileProp_VolItColor0
    eProfileProp_VolItColor1
    eProfileProp_OtherTextColor
    eProfileProp_BackgroundColor
    eProfileProp_VolShadeColor
    eProfileProp_VolShade
    eProfileProp_VolPercentShow
    eProfileProp_VolActualShow
    eProfileProp_GridLines
    eProfileProp_ExtraRows
    eProfileProp_MarginLeft
    eProfileProp_MarginRight
    eProfileProp_BoxFirst
    eProfileProp_TPOCountShow
    eProfileProp_IteratorLookback
    eProfileProp_IteratorHighVol
    eProfileProp_SMPMinTPO
    eProfileProp_SMPBreakOut
End Enum

Private Type mPrivate
    WindowLink As New cWindowLink
    
    BarsTicks As cGdBars                'for streaming updates
    BarsMinute As cGdBars               'minute bars for building Steildmayer bars
    BarsSMP As cGdBars                  'Steidlmayer bars
    BarsProfile As cGdBars

    nSessionDate As Double
    
    strSymbol As String
    nSymbolID As Long
    
    nDaysBack As Long
    nDaysProfile As Long                'days per profile cluster
    
    nSMPMinTPO As Long
    nSMPBreakOut As Long
    
    nIteratorLookback As Long
    nIteratorHighVol As Long
    
    nMarginLeft As Long                 'left margin in pixels
    nMarginRight As Long                'right margin in pixels
    
    nGradientFrom As Long               'these color values are saved to INI file
    nGradientTo As Long                 'then passed to DLL as color1, color2
    nUpColor As Long                    'based on user-specified color scheme
    nDownColor As Long
    nBidColor As Long                   'color for sell volume
    nAskColor As Long                   'color for buy volume
    nSMPColor As Long
    nVolItColor0 As Long                'vol iterator color for flag = 0; default=32768 (forest green)
    nVolItColor1 As Long                'vol iterator color for flag = 1; default=vbRed
    
    nShadeVolume As Long                '0=no shading
    nVolumeColor As Long                'color to shade volume with
    
    nProfileCount As Long               '# of profiles
    nProfileInterval As Long            'number of minutes or days per interval
    nIntervalType As Long               '0=minutes interval, 1=days interval
    nYScaleMinMove As Long              'y-scale ticks per row: -1 = auto
    dTimeStart As Double                'custom start time
    dTimeEnd As Double                  'custom end time
    dProfileStart As Double             'date time of first profile currently displayed
    dProfileEnd As Double               'date time of last profile currently displayed
    
    gridControl As Long
    
    gridExtendedProp As ExGrid_Extended_Properties
    gridTextSpecs As ExGrid_Text_Specs
    
    profileProp As MktProfile_Display
    statsProp As MktProfile_Stats_Prop
    
    bReload As Boolean
    bRefresh As Boolean
    bUnloadNow As Boolean
    bInitialShow As Boolean
    bCurrentSession As Boolean
    bSMPProfile As Boolean
    bSMPEnabled As Boolean
End Type
Private m As mPrivate

Public Sub ShowMe(Optional ByVal nSymbolID& = 0)
On Error GoTo ErrSection:

    If Not FileExist("Exgrid.dll") Then
        InfBox "File Exgrid.dll not found." & vbCrLf & "Trade Profile feature not available.", "I", , "Trade Profile"
        Unload Me
        Exit Sub
    End If

    m.strSymbol = GetSymbol(nSymbolID)
    m.bReload = False
    m.bRefresh = False
    tmr.Enabled = False
    
    If Not LoadSettings Then
        Unload Me
        Exit Sub
    End If
    
    m.WindowLink.Init Me
    
    If ChangeSymbol(nSymbolID) Then
        ShowForm Me, eForm_Nonmodal, frmMain
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.ShowMe"

End Sub

Private Function InitialShow(Optional ByVal nShowStatus As Long = 1, _
    Optional ByVal bNewBar As Boolean = False) As Long '0=success
On Error GoTo ErrSection

    Dim i&, j&, k&, rc&, iPtrSave&, strText$
    
    Dim dtBegin#, dtEnd#, d#, dtTimeout#
    
    Dim bEnabled As Boolean
    Dim bSuccess As Boolean
    Dim bCustomStart As Boolean
    Dim bCustomEnd As Boolean
    
    Dim aHandles As Long
    
    Dim Bars As cGdBars
    Dim SymInf As cSymbolInfo

    If m.bInitialShow Then Exit Function
    m.bInitialShow = True
    
    iPtrSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    bEnabled = tmr.Enabled
    tmr.Enabled = False
    
    fg.FixedRows = 0
    fg.FixedCols = 0
    fg.Col = 1
    fg.Rows = 1
    fg.ColWidthMin = Me.ScaleWidth
    
    If 0 <> nShowStatus Then         '0=no, 1=yes, -1=write to grid but don't toggle visibility
        If 1 = nShowStatus Then
            fg.Visible = True
            pbExgrid.Visible = False
        End If

        If m.gridControl = 0 Then
            strText = "Loading data ..."
            fg.TextMatrix(fg.FixedRows, 0) = strText
            DoEvents
        Else
            strText = "DLL reset ..."
            fg.TextMatrix(fg.FixedRows, 0) = strText
            DoEvents
            rc = gxMktInitPrices(m.gridControl, 0, 0, 0, -1, -1, -1, -1, -1)  'clear out previous data on DLL side
            If rc = -10 Then
                DoEvents        '-10 means UpdateRT or some other procedure has not completed, try 2 more times
                rc = gxMktInitPrices(m.gridControl, 0, 0, 0, -1, -1, -1, -1, -1)  'clear out previous data on DLL side
                If rc = -10 Then
                    DoEvents
                    rc = gxMktInitPrices(m.gridControl, 0, 0, 0, -1, -1, -1, -1, -1)  'clear out previous data on DLL side
                End If
            End If
        End If
    ElseIf m.gridControl <> 0 Then
        rc = gxMktInitPrices(m.gridControl, 0, 0, 0, -1, -1, -1, -1, -1)  'clear out previous data on DLL side
    End If
    
    If rc <> 0 Then
        fg.Visible = True
        pbExgrid.Visible = False
        strText = "DLL reset failed system error:" & Str(rc) & ". Please close Trade Profile window and try again."
        fg.TextMatrix(fg.FixedRows, 0) = strText
        DoEvents
        bEnabled = False
        GoTo ErrExit
    End If
    
    If bNewBar Then
        bSuccess = True
        dtBegin = m.dProfileStart
        
        j = m.BarsMinute.FindDateTime(m.dProfileEnd)
        For i = j To m.BarsMinute.Size - 1
            k = m.BarsMinute.SessionDateForTime(m.BarsMinute(eBARS_DateTime, i), True)
            If k > m.nSessionDate Then
                m.nProfileCount = m.nProfileCount + 1
                m.nSessionDate = k
            End If
        Next
        
        m.nSessionDate = m.BarsMinute.SessionDateForTime(m.BarsMinute(eBARS_DateTime, m.BarsMinute.Size - 1), True)
        dtEnd = m.nSessionDate + m.BarsMinute.Prop(eBARS_EndTime) / 1440#
        m.dProfileEnd = dtEnd
    Else
        ResetData
        
        ' TLB 10/17/2012: make sure today's tick data has been requested from salmon and is now available
        ' before continuing (but with a timeout just in case)       -6735
        dtTimeout = gdTickCount + 10000
        If g.RealTime.Active And g.RealTime.SalmonIsRunning Then
            Set SymInf = g.RealTime.SymbolInfo(m.nSymbolID)
            Do While SymInf.GetDataRequestStatus(ePRD_EachTick) = eSalmonPending And gdTickCount < dtTimeout
                DoEvents
            Loop
        End If
        
        'get daily bars to determine first/last dates
        Set Bars = New cGdBars
        Bars.ArrayMask = eBARS_Eod Or eBARS_BidAsk
        bSuccess = DM_GetBars(Bars, m.nSymbolID, "Daily", 0, m.nSessionDate)
        If Bars.Size <= 0 Then
            GoTo ErrExit        'TODO: error message here
        End If
        If m.bCurrentSession Or m.nSessionDate = 0 Then
            If g.RealTime.Active Then
                g.RealTime.AddTickBuffer Bars
                g.RealTime.UpdateBars Bars
            End If
            m.nSessionDate = Bars.SessionDateForTime(Bars(eBARS_DateTime, Bars.Size - 1), True)
            g.RealTime.RemoveTickBuffer Bars
        End If
        
        If m.nSessionDate = 0 Then
            'theoretically should never get here
            DebugLog "frmMarketPofile sessiondate = 0 symbol:" & Bars.Prop(eBARS_Symbol)
            GoTo ErrExit
        End If
        
        'set initial first date
        dtEnd = m.nSessionDate
        i = Bars.FindDateTime(dtEnd)        '6975
        
        If i - m.nDaysBack >= 0 Then
            For k = 1 To m.nDaysBack
                dtBegin = Bars.SessionDate(i - k)
            Next
        ElseIf i > 0 Then
            dtBegin = Bars.SessionDate(0)
            m.nDaysBack = i - 1
        End If
        
        If dtBegin = 0 Then
            m.nDaysBack = 0
            dtBegin = dtEnd
        End If
    
        If 0 <> nShowStatus Then
            With fg
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "Building minute bars ..."
                .ShowCell .Rows - 1, 0
            End With
            DoEvents
        End If
    End If
    
    If m.bSMPProfile Then
        'add 10 more days back to first date so we can be sure we have complete profile for the first requested session
        j = dtBegin - 10
        
        If m.BarsMinute Is Nothing Then Set m.BarsMinute = New cGdBars
        
        Dim bBuildMinBars As Boolean
        
        bBuildMinBars = True
        If m.bCurrentSession And g.RealTime.Active And m.BarsMinute.Size > 0 Then
            d = m.BarsMinute(eBARS_DateTime, m.BarsMinute.Size - 1)
            i = m.BarsMinute.SessionDateForTime(d, True)
            If m.BarsMinute(eBARS_DateTime, 0) <= dtBegin And i >= m.BarsMinute.SessionDateForTime(dtEnd, True) Then
                bBuildMinBars = False
                If i > dtEnd Then dtEnd = i
            End If
        End If
        
        If bBuildMinBars Then
            Set m.BarsMinute = New cGdBars
            SetBarProperties m.BarsMinute, m.nSymbolID
            
            If m.dTimeStart > 0 And m.dTimeStart <> m.BarsMinute.Prop(eBARS_DefaultStartTime) Then
                m.BarsMinute.Prop(eBARS_StartTime) = m.dTimeStart
            End If
            If m.dTimeEnd > 0 And m.dTimeEnd <> m.BarsMinute.Prop(eBARS_DefaultEndTime) Then
                m.BarsMinute.Prop(eBARS_EndTime) = m.dTimeEnd
            End If
            
            bSuccess = DM_GetBars(m.BarsMinute, m.nSymbolID, Str(m.nProfileInterval) & "m", j, dtEnd)
            If m.bCurrentSession Or m.nSessionDate = 0 Then
                If g.RealTime.Active Then
                    g.RealTime.AddTickBuffer m.BarsMinute
                    g.RealTime.UpdateBars m.BarsMinute
                    g.RealTime.SpliceBars m.BarsMinute
                Else
                    g.RealTime.AddTickBuffer m.BarsMinute       'fixes SMP bars not getting updated if streaming starts after profile was opened
                End If
            End If
        Else
            bSuccess = True
        End If
        
        If bSuccess And m.BarsMinute.Size > 0 Then
            If 0 <> nShowStatus Then
                With fg
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = "Building SMP bars ... "
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = "Begin date: " & DateFormat(dtBegin, MM_DD_YYYY, HH_MM_SS) & " (" & Str(dtBegin) & ")"
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = "End date: " & DateFormat(dtEnd, MM_DD_YYYY, HH_MM_SS) & " (" & Str(dtEnd) & ")"
                    
                    If Not m.BarsSMP Is Nothing Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "SMPBars previous size = " & Str(m.BarsSMP.Size)
                    End If
                    
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = "Minute Bars Start = " & DateFormat(m.BarsMinute(eBARS_DateTime, 0), MM_DD_YYYY, HH_MM_SS)
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = "Minute Bars End = " & DateFormat(m.BarsMinute(eBARS_DateTime, m.BarsMinute.Size - 1), MM_DD_YYYY, HH_MM_SS)
                    
                    .ShowCell .Rows - 1, 0
                End With
                DoEvents
            End If
            
            Set m.BarsSMP = New cGdBars
            strText = Str(m.nSMPMinTPO) & "/" & Str(m.nSMPBreakOut) & "/" & Str(m.nProfileInterval) & "smp"
            i = GetPeriodicity(strText)
            m.BarsSMP.Prop(eBARS_Periodicity) = i
            SetBarProperties m.BarsSMP, m.nSymbolID
            bSuccess = m.BarsSMP.BuildSMPBars(m.BarsMinute, m.nSMPMinTPO, m.nSMPBreakOut)
            
            If bSuccess And m.BarsSMP.Size > 0 Then
                ' determine begin date.time
                i = m.BarsSMP.FindDateTime(dtBegin)
                If i >= m.BarsSMP.Size Then
                    i = m.BarsSMP.Size - 1
                    dtBegin = m.BarsSMP(eBARS_DateTime, i)
                End If
                Do While i > 0
                    If m.BarsSMP.SessionDate(i) < dtBegin Then
                        dtBegin = m.BarsSMP(eBARS_DateTime, i) ' get the date.time that next SMP bar starts
                        m.BarsSMP.DeleteFirstBars i + 1 ' then throw away these bars
                        Exit Do
                    End If
                    i = i - 1
                Loop
                
                If 0 <> nShowStatus Then
                    With fg
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "SMPBars built begin date: " & DateFormat(dtBegin, MM_DD_YYYY, HH_MM_SS) & " (" & Str(dtBegin) & ")"
                        .ShowCell .Rows - 1, 0
                    End With
                    DoEvents
                End If
                
                
                i = m.BarsSMP.FindDateTime(dtEnd)
                d = m.BarsSMP(eBARS_DateTime, i) - m.BarsSMP.SessionDate(i)        'get time portion of bar datetime
                ' determine end date.time
                For i = m.BarsSMP.Size - 1 To 0 Step -1
                    If m.BarsSMP.SessionDate(i) <= dtEnd Then
                        ' does part of the desired ending session fall into the next SMP profile?
                        If d * 1440 = m.BarsSMP.Prop(eBARS_EndTime) Then
                            ' if the bar time matches the session ending time, then no data for this session is in the next SMP bar
                            m.BarsSMP.Size = i + 1
                        ElseIf i < m.BarsSMP.Size - 1 Then
                            'make sure actual bar size is >= to what we want to set it to else will get invalid values & TN hangs
                            If m.BarsSMP.Size >= i + 2 Then
                                m.BarsSMP.Size = i + 2      ' we need the next bar since part of the session is in there
                            ElseIf m.BarsSMP.Size >= i + 1 Then
                                m.BarsSMP.Size = i + 1      'just use the last bar since this is the valid number of bars available
                            End If
                        ElseIf m.BarsSMP.Size >= i + 1 Then
                            m.BarsSMP.Size = i + 1
                        End If
                        Exit For
                    End If
                Next
                dtEnd = m.BarsSMP(eBARS_DateTime, m.BarsSMP.Size - 1)
            
                If 0 <> nShowStatus Then
                    With fg
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "SMPBars built end date: " & DateFormat(dtEnd, MM_DD_YYYY, HH_MM_SS) & " (" & Str(dtEnd) & ")"
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "SMPBars size = " & Str(m.BarsSMP.Size)
                        .ShowCell .Rows - 1, 0
                    End With
                    DoEvents
                End If
            End If
        End If
    Else
    
        If Not bNewBar Then
            Set m.BarsMinute = New cGdBars
            
            SetBarProperties m.BarsMinute, m.nSymbolID
            
            If m.dTimeStart > 0 And m.dTimeStart <> Bars.Prop(eBARS_DefaultStartTime) Then
                m.BarsMinute.Prop(eBARS_StartTime) = m.dTimeStart
                bCustomStart = True
            End If
            If m.dTimeEnd > 0 And m.dTimeEnd <> m.BarsMinute.Prop(eBARS_DefaultEndTime) Then
                m.BarsMinute.Prop(eBARS_EndTime) = m.dTimeEnd
                bCustomEnd = True
            End If
            
            bSuccess = DM_GetBars(m.BarsMinute, m.nSymbolID, Str(m.nProfileInterval) & "m", dtBegin, dtEnd)
            If m.bCurrentSession Or m.nSessionDate = 0 Then
                If g.RealTime.Active Then
                    g.RealTime.AddTickBuffer m.BarsMinute
                    g.RealTime.UpdateBars m.BarsMinute
                    g.RealTime.SpliceBars m.BarsMinute
                Else
                    g.RealTime.AddTickBuffer m.BarsMinute
                End If
            End If
        
            '6975
            'set begin & end times to minute bars datetime; the call to DM_GetBars will have returned
            'correct data range for default or custom start/stop times even for overnight symbols
            dtBegin = m.BarsMinute(eBARS_DateTime, 0)
            dtEnd = m.BarsMinute(eBARS_DateTime, m.BarsMinute.Size - 1)
        End If
    
    End If
    
    If Not bSuccess Then
        bEnabled = False
        GoTo ErrExit
    End If
    
    If Not bNewBar Then
        If 0 <> nShowStatus Then
            With fg
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "Building profile bars ..."
                .ShowCell .Rows - 1, 0
            End With
            DoEvents
        End If
        
        'if m.BarsProfile Is Nothing Then Set m.BarsProfile = New cGdBars
        'BuildProfileBars m.BarsProfile, m.nSymbolID, m.BarsMinute.SessionDateForTime(dtBegin, True), m.BarsMinute.SessionDateForTime(dtEnd, True)
        Set m.BarsProfile = ProfileBarsGet(m.nSymbolID, _
                                           m.BarsMinute.SessionDateForTime(dtBegin, True), _
                                           m.BarsMinute.SessionDateForTime(dtEnd, True), _
                                           True, _
                                           Me)
        
        If m.BarsProfile Is Nothing Then
            Set m.BarsProfile = New cGdBars     'so the check for profile bars size below will not fail
            bSuccess = False
        ElseIf g.RealTime.Active And m.bCurrentSession Then
            If m.BarsProfile.Size > 0 Then d = m.BarsProfile(eBARS_DateTime, m.BarsProfile.Size - 1)
            
            Dim nProfileSession As Long
            nProfileSession = m.BarsMinute.SessionDateForTime(d, False)
            If m.bCurrentSession And nProfileSession > m.nSessionDate Then
                m.nSessionDate = nProfileSession
            End If
            
            If nProfileSession <> m.nSessionDate Then
                If 0 <> nShowStatus Then
                    With fg
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "Data mismatch. Please close Trade Profile and try again."
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "BarsProfile session: " & Str(nProfileSession) & " nSessionDate: " & Str(m.nSessionDate)
                        .ShowCell .Rows - 1, 0
                    End With
                    DoEvents
                    bSuccess = False
                    m.BarsProfile.Size = 0
                End If
            End If
        End If
        
        If bSuccess Then i = ProfileStream(True)
    End If
    
    If m.BarsProfile.Size > 0 Then
        m.BarsProfile.Prop(eBARS_DefaultStartTime) = m.BarsMinute.Prop(eBARS_DefaultStartTime)
        m.BarsProfile.Prop(eBARS_StartTime) = m.BarsMinute.Prop(eBARS_StartTime)
        m.BarsProfile.Prop(eBARS_DefaultEndTime) = m.BarsMinute.Prop(eBARS_DefaultEndTime)
        m.BarsProfile.Prop(eBARS_EndTime) = m.BarsMinute.Prop(eBARS_EndTime)
        
        UpdateVolIterator True
        
        If Not bNewBar Then
            If m.bSMPProfile Then
                m.nProfileCount = m.BarsSMP.Size
            Else
                Set Bars = New cGdBars
                SetBarProperties Bars, m.nSymbolID
                
                Dim first As Long
                Dim last As Long
                
                first = m.BarsProfile.SessionDate(0, True)
                last = m.BarsProfile.SessionDate(m.BarsProfile.Size - 1, True)
                
                'check for a # of profile change that is less than initial # of profile requested
                'eg: user requested 60 profiles then went into settings and changed to 10 profiles
                If m.BarsProfile.SessionDateForTime(dtBegin, True) > first Then
                    first = m.BarsProfile.SessionDateForTime(dtBegin, True)
                End If
                If m.BarsProfile.SessionDateForTime(dtEnd, True) < last Then
                    last = m.BarsProfile.SessionDateForTime(dtEnd, True)
                End If
                
                'request 1440m bars instead of daily bars to make sure we have minute bars
                'regardless of whether there is a daily bar or not - YC2-current contract
                'sometimes has daily bars at very beginning of chart, but no intraday data
                'YC2-201409 has daily bar on 05-13-2014, but no intraday data until 05-30-2014
                bSuccess = DM_GetBars(Bars, m.nSymbolID, "1440" & "m", first, last)
                If m.bCurrentSession Or m.nSessionDate = 0 Then
                    If g.RealTime.Active Then
                        g.RealTime.SpliceBars Bars
                    End If
                End If
                
                If Bars.Size = 0 Then
                    'if no intraday data then something very wrong - do not continue
                    InfBox "DM returned no data to request for 1440m bars; cannot do trade profile", "i", , Bars.Prop(eBARS_Symbol)
                    GoTo ErrExit
                Else
                    'adjust profile count to actual data available - aardvark 7002
                    m.nProfileCount = Bars.Size
                End If
            
            End If
        End If
        
        If m.nYScaleMinMove <= 0 Then m.nYScaleMinMove = -1
        
        If m.gridControl = 0 Then
            m.gridControl = gxMktProfileNew()
            rc = gxMktProfileInit(Me.hWnd, pbExgrid.hWnd, m.gridControl)
        End If
        
        If m.gridControl <> 0 And rc = 0 Then
            rc = SetGridTextSpecs()
            If rc = 0 Then rc = SetProfileSpecs()
            
            If rc = 0 Then
                If 0 <> nShowStatus Then
                    With fg
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "Initializing prices ..."
                        .ShowCell .Rows - 1, 0
                    End With
                    DoEvents
                End If
                
                gxMktProfileCallback m.gridControl, AddressOf fnDLLStatus
                
                If Not m.bSMPProfile Or m.BarsSMP Is Nothing Then
                    'if this is not SMP profile or the SMP bars is nothing (precautionary check since SMP bars should not be nothing for SMP profiles)
                    rc = gxMktInitPrices(m.gridControl, m.BarsProfile.BarsHandle, m.BarsMinute.BarsHandle, 0, dtBegin, dtEnd, m.nProfileCount, m.profileProp.ExtraRows, m.nYScaleMinMove)
                Else
                    rc = gxMktInitPrices(m.gridControl, m.BarsProfile.BarsHandle, m.BarsMinute.BarsHandle, m.BarsSMP.BarsHandle, dtBegin, dtEnd, m.nProfileCount, m.profileProp.ExtraRows, m.nYScaleMinMove)
                End If
                
                If rc = 0 Then
                    m.gridExtendedProp.reserved3 = gxMktAutoTicks(m.gridControl)
                    rc = SetColorSchemeDLL
                    If rc = 0 Then rc = SetProfileSpecs()
                End If
            End If
        End If
    
        If m.gridControl = 0 Or rc <> 0 Then bSuccess = False
    Else
        fg.Visible = True
        pbExgrid.Visible = False
        fg.Rows = fg.Rows + 1
        fg.TextMatrix(fg.Rows - 1, 0) = "No profile data, profile bars size = zero."
        bSuccess = False
        bEnabled = False
        rc = 1          'to indicate non-success
    End If
    
    If bSuccess Then
        Dim bLocked As Boolean
        
        If 0 <> nShowStatus Then
            With fg
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "Processing display (" & Str(m.nProfileCount) & " profiles)..."
                .ShowCell .Rows - 1, 0
            End With
        Else
            bLocked = LockWindowUpdate(pbExgrid.hWnd)
            pbExgrid.Cls
        End If
        DoEvents
        
        aHandles = gdCreateArray(eGDARRAY_Longs, 2, 0)
        If m.bSMPProfile Then
            gdSetNum aHandles, 0, m.BarsMinute.BarsHandle
            gdSetNum aHandles, 1, m.BarsSMP.BarsHandle
            rc = gxMktProfileData(m.gridControl, aHandles, m.nProfileInterval, m.nProfileCount, dtBegin, dtEnd, 1)
        Else
            gdSetNum aHandles, 0, m.BarsMinute.BarsHandle
            gdSetNum aHandles, 1, m.BarsProfile.BarsHandle
            rc = gxMktProfileData(m.gridControl, aHandles, m.nProfileInterval, m.nProfileCount, dtBegin, dtEnd, -1)
        End If
        
        If bLocked Then
            LockWindowUpdate 0
            gxMktGridRefresh m.gridControl
        End If
        
        gdDestroyArray aHandles
    
        If rc = 0 Then
            m.dProfileStart = dtBegin
            m.dProfileEnd = dtEnd
            
            fg.Visible = False
            pbExgrid.Visible = True
            
            If m.bSMPProfile Then
                Me.Caption = m.strSymbol & " " & DateFormat(m.BarsSMP.SessionDate(dtEnd), MM_DD_YYYY) & " (" & Str(m.nSMPMinTPO) & "/" & Str(m.nSMPBreakOut) & "/" & Str(m.nProfileInterval) & "smp)"
            Else
                Me.Caption = m.strSymbol & " " & DateFormat(m.BarsProfile.SessionDate(dtEnd), MM_DD_YYYY) & " (" & Str(m.nProfileInterval) & "min)"
            End If
            If 0 <> nShowStatus Or m.nYScaleMinMove > 0 Then SetScaleDropdown
        Else
            fg.Visible = True
            pbExgrid.Visible = False
            If rc <> -1 Then
                fg.Rows = fg.Rows + 1
                fg.TextMatrix(fg.Rows - 1, 0) = "Processing display failed error  = " & Str(rc)
            End If
        End If
    End If
    
    DoEvents
    rc = gxMktCenterPrice(m.gridControl)
    
    m.bReload = False
    m.bRefresh = False
    
    InitialShow = rc
    
ErrExit:
    Screen.MousePointer = iPtrSave
    m.bInitialShow = False
    tmr.Enabled = bEnabled
    
    Exit Function

ErrSection:
    Screen.MousePointer = iPtrSave
    tmr.Enabled = False
    m.bInitialShow = False
    
    RaiseError "frmMarketProfile.InitialShow"
    
End Function

Private Sub Form_Activate()
On Error GoTo ErrSection:

    TextIncDecRegisterForm Me, True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.Form_Activate"

End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:
    
    TextIncDecUnregisterForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.Form_Deactivate"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection

    Dim strText$
    
    Me.Icon = Picture16(ToolbarIcon("ID_MarketProfile"), , True)

    g.Styler.StyleForm Me
    
    With tbToolbar
        strText = "Classic"
        If g.nTbIconStyle = 1 Then
            If g.nColorTheme = kDarkThemeColor Then
                strText = "Light"
            Else
                strText = "Dark"
            End If
            .Tools("ID_Centered").Picture = g.CoreBridge.ImgListToolbarExt(strText, "kCenterPrice", "", 16).ExtractIcon
        End If
        .Tools("ID_Symbol").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_Symbol"), "", 16).ExtractIcon
        .Tools("ID_Settings").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_Settings"), "", 16).ExtractIcon
        .Tools("ID_TextIncrease").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_TextIncrease"), "", 16).ExtractIcon
        .Tools("ID_TextDecrease").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_TextDecrease"), "", 16).ExtractIcon
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
        
        .Tools("ID_Scale").ComboBox.AddItem "Auto"
        .Tools("ID_Scale").ComboBox.AddItem "1 tick"
        .Tools("ID_Scale").ComboBox.AddItem "2 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "3 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "4 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "5 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "6 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "10 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "15 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "20 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "25 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "50 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "75 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "100 ticks"
        .Tools("ID_Scale").ComboBox.AddItem "200 ticks"
    End With

    'Restore/set form size & location
    strText = GetIniFileProperty(Me.Name, "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText
    End If
   
    fg.Font.Name = "Arial"
    fg.Font.Size = 10
    fg.Font.Bold = True
   
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.Form_Load"
    Unload Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Me.tmr.Enabled = False
    
    If m.bInitialShow Then
        m.bUnloadNow = True
        Cancel = 1
        Exit Sub
    End If

    TextIncDecUnregisterForm Me
    
    If FormIsLoaded("frmMarketProfileCfg") Then Unload frmMarketProfileCfg      '4720

    If Cancel = 0 Then m.WindowLink.Unhook

    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.Form_QueryUnload"
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim i&
    
    i = m.nMarginLeft * Screen.TwipsPerPixelX + m.nMarginRight * Screen.TwipsPerPixelX

    With pbExgrid
        .Move m.nMarginLeft * Screen.TwipsPerPixelX, 0, Me.ScaleWidth - i, Me.ScaleHeight
    End With
    
    fg.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    Me.BackColor = m.gridExtendedProp.gridBkColor
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    Dim rc&, i&

    Me.tmr.Enabled = False
    
    If m.bInitialShow = True Then Exit Sub
    
    If m.gridControl <> 0 Then
        rc = gxMktProfileDestroy(m.gridControl)
        Me.Hide
    End If
    
    Set m.WindowLink = Nothing
    
    ProfileBarsFree m.BarsProfile

    If m.gridTextSpecs.gshFontName <> 0 Then gdDestroyArray m.gridTextSpecs.gshFontName

    'save form size & location
    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.Form_Unload"

End Sub

Private Sub ResetData()
On Error GoTo ErrSection:

    Dim strTemp$, i&
    
    If m.nDaysBack < 0 Then m.nDaysBack = 0
    If m.nDaysProfile = 0 Then m.nDaysProfile = 0
    If m.nProfileInterval = 0 Then m.nProfileInterval = 30
    m.nProfileCount = 0
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.ResetData"

End Sub

Public Property Get ProfileInterval() As Long
On Error GoTo ErrSection:

    ProfileInterval = m.nProfileInterval

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.ProfileIntervalGet"

End Property

Public Property Let ProfileInterval(ByVal nInterval&)

    If m.nProfileInterval <> nInterval Then
        m.nProfileInterval = nInterval
        m.bReload = True
    End If

End Property

Public Property Get IntervalType() As Long
On Error GoTo ErrSection:

    IntervalType = m.nIntervalType

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.IntervalTypeGet"

End Property

Public Property Let IntervalType(ByVal nType&)

    If m.nIntervalType <> nType Then
        m.nIntervalType = nType
        m.bReload = True
    End If

End Property

Public Property Get DaysBack() As Long
On Error GoTo ErrSection:
    
    DaysBack = m.nDaysBack

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfileGet.DaysBack"

End Property

Public Property Let DaysBack(ByVal nDays&)
On Error GoTo ErrSection:

    If nDays >= 0 Then
        If m.nDaysBack <> nDays Then
            m.nDaysBack = nDays
            m.bReload = True
        End If
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfileLet.DaysBack"

End Property

Public Property Get DaysProfile() As Long
On Error GoTo ErrSection:
    
    DaysProfile = m.nDaysProfile

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.DaysProfile"

End Property

Public Property Let DaysProfile(ByVal nDaysProfile&)

    If m.nDaysProfile <> nDaysProfile Then
        m.nDaysProfile = nDaysProfile
        m.bReload = True
    End If
    
End Property

Public Property Get SessionDate() As Long
On Error GoTo ErrSection:
    
    If m.bCurrentSession Then
        SessionDate = 0
    Else
        SessionDate = m.nSessionDate
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfileGet.SessionDateGet"

End Property

Public Property Let SessionDate(ByVal nDate&)
On Error GoTo ErrSection

    If m.nSessionDate <> nDate Then
        m.nSessionDate = nDate
        m.bReload = True
    End If
    
    If m.nSessionDate = 0 Then
        m.bCurrentSession = True
    Else
        m.bCurrentSession = False
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfileLet.SessionDateGet"

End Property

Public Property Get GridFontBold() As Long
On Error GoTo ErrSection:
    
    GridFontBold = m.gridTextSpecs.FontBold

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.GridFontBoldGet"

End Property

Public Property Get GridFontItalic() As Long
On Error GoTo ErrSection:
    
    GridFontItalic = m.gridTextSpecs.FontItalic

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.GridFontItalicGet"

End Property

Public Property Get GridFontSize() As Long
On Error GoTo ErrSection:
    
    GridFontSize = m.gridTextSpecs.FontSize

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.GridFontSizeGet"

End Property

Public Property Get GridFontName() As String
On Error GoTo ErrSection:
    
    GridFontName = gdGetStr(m.gridTextSpecs.gshFontName, 0)

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.GridFontNameGet"

End Property

Public Property Get ProfileProperty(ByVal ePropType As eProfilePropType) As Long
    
    Dim nReturn As Long
    
    Select Case ePropType
        Case eProfileProp_ColorScheme
            nReturn = m.profileProp.ColorScheme
            
        Case eProfileProp_GradientFrom
            nReturn = m.nGradientFrom
        Case eProfileProp_GradientTo
            nReturn = m.nGradientTo
        Case eProfileProp_UpColor
            nReturn = m.nUpColor
        Case eProfileProp_DownColor
            nReturn = m.nDownColor
        Case eProfileProp_BidColor
            nReturn = m.nBidColor
        Case eProfileProp_AskColor
            nReturn = m.nAskColor
        
        'volume iterator
        Case eProfileProp_IteratorLookback
            nReturn = m.nIteratorLookback
        Case eProfileProp_IteratorHighVol
            nReturn = m.nIteratorHighVol
        Case eProfileProp_VolItColor0
            nReturn = m.nVolItColor0
        Case eProfileProp_VolItColor1
            nReturn = m.nVolItColor1
            
        'auto split SMP profile
        Case eProfileProp_SMPColor
            nReturn = m.nSMPColor
        Case eProfileProp_SMPMinTPO
            nReturn = m.nSMPMinTPO
        Case eProfileProp_SMPBreakOut
            nReturn = m.nSMPBreakOut
        
        Case eProfileProp_BackgroundColor
            nReturn = m.gridExtendedProp.gridBkColor
        Case eProfileProp_GridLines
            nReturn = m.gridExtendedProp.GridLines
        
        Case eProfileProp_OtherTextColor
            nReturn = m.gridTextSpecs.Color
        Case eProfileProp_VolShadeColor
            nReturn = m.nVolumeColor
        Case eProfileProp_VolShade
            nReturn = m.nShadeVolume
        Case eProfileProp_VolPercentShow
            If m.profileProp.VolumeText = MktProf_Text_Percent Or m.profileProp.VolumeText = MktProf_Text_Both Or _
               m.profileProp.VolumeText = MktProf_Text_TpoVolPercent Or m.profileProp.VolumeText = MktProf_Text_All Then
                nReturn = 1
            End If
        Case eProfileProp_VolActualShow
            If m.profileProp.VolumeText = MktProf_Text_Actual Or m.profileProp.VolumeText = MktProf_Text_Both Or _
               m.profileProp.VolumeText = MktProf_Text_TpoVolActual Or m.profileProp.VolumeText = MktProf_Text_All Then
                nReturn = 1
            End If
        Case eProfileProp_ExtraRows
            nReturn = m.profileProp.ExtraRows
        Case eProfileProp_MarginLeft
            nReturn = m.nMarginLeft
        Case eProfileProp_MarginRight
            nReturn = m.nMarginRight
        Case eProfileProp_BoxFirst
            nReturn = m.profileProp.BoxFirst
        Case eProfileProp_TPOCountShow
            If m.profileProp.VolumeText >= MktProf_Text_Tpo Then nReturn = 1
    End Select
    
    ProfileProperty = nReturn
    
End Property

Public Property Let ProfileProperty(ByVal ePropType As eProfilePropType, ByVal nValue As Long)

    Dim bResize As Boolean

    If nValue < 0 Then nValue = 0

    Select Case ePropType
        Case eProfileProp_ColorScheme
            If m.profileProp.ColorScheme <> nValue Then
                m.profileProp.ColorScheme = nValue
                m.bRefresh = True
            End If
        
        'context-dependent TPO colors (begin)
        Case eProfileProp_GradientFrom
            If m.nGradientFrom <> nValue Then
                m.nGradientFrom = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_GradientTo
            If m.nGradientTo <> nValue Then
                m.nGradientTo = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_UpColor
            If m.nUpColor <> nValue Then
                m.nUpColor = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_DownColor
            If m.nDownColor <> nValue Then
                m.nDownColor = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_BidColor
            If m.nBidColor <> nValue Then
                m.nBidColor = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_AskColor
            If m.nAskColor <> nValue Then
                m.nAskColor = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_SMPColor
            If m.nSMPColor <> nValue Then
                m.nSMPColor = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_VolItColor0
            If m.nVolItColor0 <> nValue Then
                m.nVolItColor0 = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_VolItColor1
            If m.nVolItColor1 <> nValue Then
                m.nVolItColor1 = nValue
                m.bRefresh = True
            End If
        'context-dependent TPO colors (end)
        
        'volume iterator
        Case eProfileProp_IteratorLookback
            If m.nIteratorLookback <> nValue Then
                m.nIteratorLookback = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_IteratorHighVol
            If m.nIteratorHighVol <> nValue Then
                m.nIteratorHighVol = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_VolItColor0
            If m.nVolItColor0 <> nValue Then
                m.nVolItColor0 = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_VolItColor1
            If m.nVolItColor1 <> nValue Then
                m.nVolItColor1 = nValue
                m.bRefresh = True
            End If
            
        'auto split SMP profile
        Case eProfileProp_SMPColor
            If m.nSMPColor <> nValue Then
                m.nSMPColor = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_SMPMinTPO
            If m.nSMPMinTPO <> nValue Then
                If m.nSMPBreakOut < nValue Then
                    m.nSMPMinTPO = nValue
                    m.bReload = True
                End If
            End If
        Case eProfileProp_SMPBreakOut
            If m.nSMPBreakOut <> nValue Then
                If m.nSMPMinTPO > nValue Then
                    m.nSMPBreakOut = nValue
                    m.bReload = True
                End If
            End If
            
        Case eProfileProp_GridLines
            If m.gridExtendedProp.GridLines <> nValue Then
                If nValue = 0 Or nValue = 1 Then
                    m.gridExtendedProp.GridLines = nValue
                    m.bRefresh = True
                End If
            End If
        Case eProfileProp_BackgroundColor
            If m.gridExtendedProp.gridBkColor <> nValue Then
                m.gridExtendedProp.gridBkColor = nValue
                Me.BackColor = nValue
                m.bRefresh = True
            End If
        
        Case eProfileProp_OtherTextColor
            If m.gridTextSpecs.Color <> nValue Then
                m.gridTextSpecs.Color = nValue
                m.bRefresh = True
            End If
        Case eProfileProp_VolShadeColor
            If m.nVolumeColor <> nValue Then
                m.nVolumeColor = nValue
                If m.nShadeVolume = 1 Then
                    m.profileProp.VolumeColor = m.nVolumeColor      '6735
                    m.bRefresh = True
                ElseIf m.nShadeVolume = 0 Then
                    If m.profileProp.VolumeColor <> -1 Then
                        m.profileProp.VolumeColor = -1
                        m.bRefresh = True
                    End If
                ElseIf m.profileProp.VolumeColor = -1 Then
                    m.profileProp.VolumeColor = m.nVolumeColor      'theoretically should not get here
                    m.bRefresh = True                               'but if does, then just set the color
                End If
            End If
        Case eProfileProp_VolShade
            If m.nShadeVolume <> nValue Then
                m.nShadeVolume = nValue
                If nValue = 0 Then
                    m.profileProp.VolumeColor = -1
                Else
                    m.profileProp.VolumeColor = m.nVolumeColor
                End If
                m.bRefresh = True
            End If
        
        Case eProfileProp_VolPercentShow
            Select Case m.profileProp.VolumeText
                Case MktProf_Text_Both
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Actual
                        m.bRefresh = True
                    End If
                Case MktProf_Text_Percent
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_None
                        m.bRefresh = True
                    End If
                Case MktProf_Text_None
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Percent
                        m.bRefresh = True
                    End If
                Case MktProf_Text_Actual
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Both
                        m.bRefresh = True
                    End If
                Case MktProf_Text_Tpo
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_TpoVolPercent
                        m.bRefresh = True
                    End If
                Case MktProf_Text_TpoVolPercent
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Tpo
                        m.bRefresh = True
                    End If
                Case MktProf_Text_TpoVolActual
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_All
                        m.bRefresh = True
                    End If
                Case MktProf_Text_All
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_TpoVolActual
                        m.bRefresh = True
                    End If
            End Select
            
        Case eProfileProp_VolActualShow
            Select Case m.profileProp.VolumeText
                Case MktProf_Text_Both
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Percent
                        m.bRefresh = True
                    End If
                Case MktProf_Text_Actual
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_None
                        m.bRefresh = True
                    End If
                Case MktProf_Text_None
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Actual
                        m.bRefresh = True
                    End If
                Case MktProf_Text_Percent
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Both
                        m.bRefresh = True
                    End If
                Case MktProf_Text_Tpo
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_TpoVolActual
                        m.bRefresh = True
                    End If
                Case MktProf_Text_TpoVolActual
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Tpo
                        m.bRefresh = True
                    End If
                Case MktProf_Text_TpoVolPercent
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_All
                        m.bRefresh = True
                    End If
                Case MktProf_Text_All
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_TpoVolPercent
                        m.bRefresh = True
                    End If
            End Select
        
        Case eProfileProp_ExtraRows
            If nValue >= 0 Then
                If nValue <> m.profileProp.ExtraRows Then
                     m.profileProp.ExtraRows = nValue
                     m.bReload = True
                End If
            End If
        
        Case eProfileProp_MarginLeft
            If nValue >= 0 Then
                If nValue <> m.nMarginLeft Then
                    m.nMarginLeft = nValue
                    bResize = True
                End If
            End If
    
        Case eProfileProp_MarginRight
            If nValue >= 0 Then
                If nValue <> m.nMarginRight Then
                    m.nMarginRight = nValue
                    bResize = True
                End If
            End If
        
        Case eProfileProp_BoxFirst
            If nValue < 0 Then
                nValue = 0
            ElseIf nValue > 1 Then
                nValue = 1
            End If
            If nValue <> m.profileProp.BoxFirst Then
                m.profileProp.BoxFirst = nValue
                m.bRefresh = True
            End If
        
        Case eProfileProp_TPOCountShow
            Select Case m.profileProp.VolumeText
                Case MktProf_Text_Both
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_All
                        m.bRefresh = True
                    End If
                Case MktProf_Text_Actual
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_TpoVolActual
                        m.bRefresh = True
                    End If
                Case MktProf_Text_None
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Tpo
                        m.bRefresh = True
                    End If
                Case MktProf_Text_Percent
                    If nValue <> 0 Then
                        m.profileProp.VolumeText = MktProf_Text_TpoVolPercent
                        m.bRefresh = True
                    End If
                Case MktProf_Text_Tpo
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_None
                        m.bRefresh = True
                    End If
                Case MktProf_Text_TpoVolPercent
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Percent
                        m.bRefresh = True
                    End If
                Case MktProf_Text_TpoVolActual
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Actual
                        m.bRefresh = True
                    End If
                Case MktProf_Text_All
                    If nValue = 0 Then
                        m.profileProp.VolumeText = MktProf_Text_Both
                        m.bRefresh = True
                    End If
            End Select
    
    End Select
    
    If bResize Then FormResize Me

End Property

Public Property Get StatShow(ByVal eStatsType As eProfileStatsType) As Long

    Dim nShow&

     Select Case eStatsType
        Case eProfileStats_VAVolume
            If m.statsProp.VA_Vol_Color >= 0 Then nShow = 1       'neg color means not shown
        Case eProfileStats_VATPO
            If m.statsProp.VA_TPO_Color >= 0 Then nShow = 1
        Case eProfileStats_StdDev
            If m.statsProp.StdDev_Color >= 0 Then nShow = 1
        Case eProfileStats_IB
            If m.statsProp.IB_Color >= 0 Then nShow = 1
        Case eProfileStats_POCVol
            If m.statsProp.POC_Vol_Color >= 0 Then nShow = 1
        Case eProfileStats_POCTPO
            If m.statsProp.POC_TPO_Color >= 0 Then nShow = 1
        Case eProfileStats_Mean
            If m.statsProp.Mean_Color >= 0 Then nShow = 1
    End Select
    
    StatShow = nShow

End Property

Public Property Let StatShow(ByVal eStatsType As eProfileStatsType, ByVal nShow As Long)

    Dim bShow As Boolean
    
    'this property is intended to be passed in only 0 or 1 for nShow, so if nShow <> 1 then assume zero
    'nShow can be passed in as 2(unchecked checkbox) from flexcpChecked value in vsFlexGrid
    If nShow = 1 Then
        bShow = True
    Else
        bShow = False
    End If

    Select Case eStatsType
        Case eProfileStats_VAVolume
            If bShow And m.statsProp.VA_Vol_Color < 0 Then
                m.statsProp.VA_Vol_Color = Abs(m.statsProp.VA_Vol_Color)
                m.bRefresh = True
            ElseIf Not bShow And m.statsProp.VA_Vol_Color >= 0 Then
                m.statsProp.VA_Vol_Color = -m.statsProp.VA_Vol_Color
                m.bRefresh = True
            End If
        
        Case eProfileStats_VATPO
            If bShow And m.statsProp.VA_TPO_Color < 0 Then
                m.statsProp.VA_TPO_Color = Abs(m.statsProp.VA_TPO_Color)
                m.bRefresh = True
            ElseIf Not bShow And m.statsProp.VA_TPO_Color >= 0 Then
                m.statsProp.VA_TPO_Color = -m.statsProp.VA_TPO_Color
                m.bRefresh = True
            End If
        
        Case eProfileStats_POCVol
            If bShow And m.statsProp.POC_Vol_Color < 0 Then
                m.statsProp.POC_Vol_Color = Abs(m.statsProp.POC_Vol_Color)
                m.bRefresh = True
            ElseIf Not bShow And m.statsProp.POC_Vol_Color >= 0 Then
                m.statsProp.POC_Vol_Color = -m.statsProp.POC_Vol_Color
                m.bRefresh = True
            End If
        
        Case eProfileStats_POCTPO
            If bShow And m.statsProp.POC_TPO_Color < 0 Then
                m.statsProp.POC_TPO_Color = Abs(m.statsProp.POC_TPO_Color)
                m.bRefresh = True
            ElseIf Not bShow And m.statsProp.POC_TPO_Color >= 0 Then
                m.statsProp.POC_TPO_Color = -m.statsProp.POC_TPO_Color
                m.bRefresh = True
            End If
        
        Case eProfileStats_StdDev
            If bShow And m.statsProp.StdDev_Color < 0 Then
                m.statsProp.StdDev_Color = Abs(m.statsProp.StdDev_Color)
                m.bRefresh = True
            ElseIf Not bShow And m.statsProp.StdDev_Color >= 0 Then
                m.statsProp.StdDev_Color = -m.statsProp.StdDev_Color
                m.bRefresh = True
            End If
        
        Case eProfileStats_IB
            If bShow And m.statsProp.IB_Color < 0 Then
                m.statsProp.IB_Color = Abs(m.statsProp.IB_Color)
                m.bRefresh = True
            ElseIf Not bShow And m.statsProp.IB_Color >= 0 Then
                m.statsProp.IB_Color = -m.statsProp.IB_Color
                m.bRefresh = True
            End If
        
        Case eProfileStats_Mean
            If bShow And m.statsProp.Mean_Color < 0 Then
                m.statsProp.Mean_Color = Abs(m.statsProp.Mean_Color)
                m.bRefresh = True
            ElseIf Not bShow And m.statsProp.Mean_Color >= 0 Then
                m.statsProp.Mean_Color = -m.statsProp.Mean_Color
                m.bRefresh = True
            End If
    End Select

End Property

Public Property Get StatColor(ByVal eStatsType As eProfileStatsType) As Long

    Dim nColor&

    nColor = -1
    
    Select Case eStatsType
        Case eProfileStats_VAVolume
            nColor = Abs(m.statsProp.VA_Vol_Color)        'neg color means item is not shown
        Case eProfileStats_VATPO
            nColor = Abs(m.statsProp.VA_TPO_Color)
        Case eProfileStats_StdDev
            nColor = Abs(m.statsProp.StdDev_Color)
        Case eProfileStats_IB
            nColor = Abs(m.statsProp.IB_Color)
        Case eProfileStats_POCVol
            nColor = Abs(m.statsProp.POC_Vol_Color)
        Case eProfileStats_POCTPO
            nColor = Abs(m.statsProp.POC_TPO_Color)
        Case eProfileStats_Mean
            nColor = Abs(m.statsProp.Mean_Color)
    End Select
    
    StatColor = nColor

End Property

Public Property Let StatColor(ByVal eStatsType As eProfileStatsType, ByVal nColor As Long)

    If nColor < 0 Then nColor = 0
        
    Select Case eStatsType
        Case eProfileStats_VAVolume
            If Abs(m.statsProp.VA_Vol_Color) <> nColor Then
                If m.statsProp.VA_Vol_Color >= 0 Then
                    m.statsProp.VA_Vol_Color = nColor
                    m.bRefresh = True
                Else
                    m.statsProp.VA_Vol_Color = -nColor          'item not shown, just save color - no need to refresh
                End If
            End If
        
        Case eProfileStats_VATPO = nColor
            If Abs(m.statsProp.VA_TPO_Color) <> nColor Then
                If m.statsProp.VA_TPO_Color >= 0 Then
                    m.statsProp.VA_TPO_Color = nColor
                    m.bRefresh = True
                Else
                    m.statsProp.VA_TPO_Color = -nColor
                End If
            End If
        
        Case eProfileStats_StdDev
            If Abs(m.statsProp.StdDev_Color) <> nColor Then
                If m.statsProp.StdDev_Color >= 0 Then
                    m.statsProp.StdDev_Color = nColor
                    m.bRefresh = True
                Else
                    m.statsProp.StdDev_Color = -nColor
                End If
            End If
        
        Case eProfileStats_IB
            If Abs(m.statsProp.IB_Color) <> nColor Then
                If m.statsProp.IB_Color >= 0 Then
                    m.statsProp.IB_Color = nColor
                    m.bRefresh = True
                Else
                    m.statsProp.IB_Color = -nColor
                End If
            End If
        
        Case eProfileStats_POCVol
            If Abs(m.statsProp.POC_Vol_Color) <> nColor Then
                If m.statsProp.POC_Vol_Color >= 0 Then
                    m.statsProp.POC_Vol_Color = nColor
                    m.bRefresh = True
                Else
                    m.statsProp.POC_Vol_Color = -nColor
                End If
            End If
        
        Case eProfileStats_POCTPO
            If Abs(m.statsProp.POC_TPO_Color) <> nColor Then
                If m.statsProp.POC_TPO_Color >= 0 Then
                    m.statsProp.POC_TPO_Color = nColor
                    m.bRefresh = True
                Else
                    m.statsProp.POC_TPO_Color = -nColor
                End If
            End If
        
        Case eProfileStats_Mean
            If Abs(m.statsProp.Mean_Color) <> nColor Then
                If m.statsProp.Mean_Color >= 0 Then
                    m.statsProp.Mean_Color = nColor
                    m.bRefresh = True
                Else
                    m.statsProp.Mean_Color = -nColor
                End If
            End If
    End Select
    
End Property

Public Property Get StatPenSize(ByVal eStatsType As eProfileStatsType) As Long

    Dim nSize&

    nSize = -1
    
    Select Case eStatsType
        Case eProfileStats_VAVolume
            nSize = m.statsProp.VA_Vol_PenSize
        Case eProfileStats_VATPO
            nSize = m.statsProp.VA_TPO_PenSize
        Case eProfileStats_StdDev
            nSize = m.statsProp.StdDev_PenSize
        Case eProfileStats_IB
            nSize = m.statsProp.IB_PenSize
        Case eProfileStats_POCVol
            nSize = m.statsProp.POC_Vol_PenSize
        Case eProfileStats_POCTPO
            nSize = m.statsProp.POC_TPO_PenSize
        Case eProfileStats_Mean
            nSize = m.statsProp.Mean_PenSize
    End Select
    
    StatPenSize = nSize
    
End Property

Public Property Let StatPenSize(ByVal eStatsType As eProfileStatsType, ByVal nSize As Long)

    If nSize < 1 Then
        nSize = 1
    ElseIf nSize > 15 Then
        nSize = 15
    End If

    Select Case eStatsType
        Case eProfileStats_VAVolume
            If m.statsProp.VA_Vol_PenSize <> nSize Then
                m.statsProp.VA_Vol_PenSize = nSize
                m.bRefresh = True
            End If
        Case eProfileStats_VATPO
            If m.statsProp.VA_TPO_PenSize <> nSize Then
                m.statsProp.VA_TPO_PenSize = nSize
                m.bRefresh = True
            End If
        Case eProfileStats_StdDev
            If m.statsProp.StdDev_PenSize <> nSize Then
                m.statsProp.StdDev_PenSize = nSize
                m.bRefresh = True
            End If
        Case eProfileStats_IB
            If m.statsProp.IB_PenSize <> nSize Then
                m.statsProp.IB_PenSize = nSize
                m.bRefresh = True
            End If
        Case eProfileStats_POCVol
            If m.statsProp.POC_Vol_PenSize <> nSize Then
                m.statsProp.POC_Vol_PenSize = nSize
                m.bRefresh = True
            End If
        Case eProfileStats_POCTPO
            If m.statsProp.POC_TPO_PenSize <> nSize Then
                m.statsProp.POC_TPO_PenSize = nSize
                m.bRefresh = True
            End If
        Case eProfileStats_Mean
            If m.statsProp.Mean_PenSize <> nSize Then
                m.statsProp.Mean_PenSize = nSize
                m.bRefresh = True
            End If
    End Select

End Property

Public Property Get StatPercent(ByVal eStatsType As eProfileStatsType) As Double

    Dim dPercent#

    dPercent = 0#
    
    Select Case eStatsType
        Case eProfileStats_VAVolume
            dPercent = m.statsProp.VA_Vol_Percent
        Case eProfileStats_VATPO
            dPercent = m.statsProp.VA_TPO_Percent
    End Select
    
    StatPercent = dPercent

End Property

Public Property Let StatPercent(ByVal eStatsType As eProfileStatsType, ByVal dPercent As Double)

    Select Case eStatsType
        Case eProfileStats_VAVolume
            If m.statsProp.VA_Vol_Percent <> dPercent Then
                m.statsProp.VA_Vol_Percent = dPercent
                m.bRefresh = True
            End If
        Case eProfileStats_VATPO
            If m.statsProp.VA_TPO_Percent <> dPercent Then
                m.statsProp.VA_TPO_Percent = dPercent
                m.bRefresh = True
            End If
    End Select

End Property


Public Property Get StatPercentStr(ByVal eStatsType As eProfileStatsType) As String

    Dim dPercent#, strPercent$

    dPercent = 0#
    
    Select Case eStatsType
        Case eProfileStats_VAVolume
            dPercent = m.statsProp.VA_Vol_Percent
        Case eProfileStats_VATPO
            dPercent = m.statsProp.VA_TPO_Percent
    End Select
    
    strPercent = Format(dPercent, "0.##")
    StatPercentStr = strPercent

End Property

Public Property Get CharacterSequence() As MktProfile_Char_Sequence
On Error GoTo ErrSection:
    
    CharacterSequence = m.profileProp.CharSequence

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.CharacterSequenceGet"

End Property

Public Property Let CharacterSequence(ByVal eSequence As MktProfile_Char_Sequence)

    If m.profileProp.CharSequence <> eSequence Then
        m.profileProp.CharSequence = eSequence
        m.bRefresh = True
    End If

End Property

Public Property Get TimeStart() As Double
On Error GoTo ErrSection:

    If m.dTimeStart > 0 Then
        TimeStart = m.dTimeStart / 1440
    ElseIf Not m.BarsMinute Is Nothing Then
        TimeStart = m.BarsMinute.Prop(eBARS_DefaultStartTime) / 1440
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.TimeStartGet"

End Property

Public Property Get TimeEnd() As Double
On Error GoTo ErrSection:

    If m.dTimeEnd > 0 Then
        TimeEnd = m.dTimeEnd / 1440
    ElseIf Not m.BarsMinute Is Nothing Then
        TimeEnd = m.BarsMinute.Prop(eBARS_DefaultEndTime) / 1440
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.TimeEndGet"

End Property

Private Sub tbToolbar_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)

    Dim strText$, i&
    
    strText = Parse(Tool.ComboBox.Text, " ", 1)
    If strText = "Auto" Then
        i = -1
    Else
        i = Val(Int(strText))
    End If
    
    YScaleMinMove = i
    If m.bReload Then
        SaveSettings
        pbExgrid.SetFocus
        InitialShow False
    End If
    
End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim strName$, iSize&, rc&
    Dim eText As ExGrid_Text_Specs
    Dim bSaved As Boolean
    
    If FormIsLoaded("frmMarketProfileCfg") Then
        InfBox "Please finish editing the Settings first.", "I", "Ok", "Trade Profile"
        frmMarketProfileCfg.Show
        Exit Sub
    End If

    Select Case Tool.ID
        Case "ID_Settings"
            frmMarketProfileCfg.ShowMe Me, m.BarsProfile, m.bSMPEnabled
        
        Case "ID_Close"
            Unload Me
        
        Case "ID_Symbol"
            ChangeSymbol
        
        Case "ID_TextIncrease"
            strName = gdGetStr(m.gridTextSpecs.gshFontName, 0)
            eText.FontBold = m.gridTextSpecs.FontBold
            eText.FontItalic = m.gridTextSpecs.FontItalic
            iSize = m.gridTextSpecs.FontSize
            If UCase(strName) = "SMALL FONTS" Then
                If iSize < 7 Then
                    eText.FontSize = iSize + 1
                Else
                    strName = "Arial"
                    eText.FontSize = 10
                End If
            ElseIf iSize < 72 Then
                eText.FontSize = iSize + 2
            ElseIf eText.FontBold <> 0 Then
                eText.FontBold = 0
            Else
                strName = ""        'do nothing, cannot go bigger
            End If
            
            If Len(strName) > 0 Then
                If GridTextChanged(eText, strName, True) And m.gridControl <> 0 Then
                    gxMktGridRefresh m.gridControl
                End If
            End If
            
        Case "ID_TextDecrease"
            strName = gdGetStr(m.gridTextSpecs.gshFontName, 0)
            eText.FontBold = m.gridTextSpecs.FontBold
            eText.FontItalic = m.gridTextSpecs.FontItalic
            iSize = m.gridTextSpecs.FontSize
    
            If UCase(strName) = "SMALL FONTS" Then
                If iSize > 2 Then
                    eText.FontSize = iSize - 1
                ElseIf eText.FontBold <> 0 Then
                    eText.FontBold = 0
                Else
                    strName = ""        'do nothing, cannot go smaller
                End If
            ElseIf iSize > 8 Then
                eText.FontSize = iSize - 2
            Else
                strName = "Small Fonts"
                eText.FontSize = 7
            End If
    
            If Len(strName) > 0 Then
                If GridTextChanged(eText, strName, True) And m.gridControl <> 0 Then
                    gxMktGridRefresh m.gridControl
                End If
            End If
        
        Case "ID_ToggleGrid"
            fg.Visible = Not fg.Visible
            fg.ZOrder
        
'If IsIDE Then
'    Me.WindowState = vbMinimized
'End If
        
        Case "ID_Centered"
            rc = gxMktCenterPrice(m.gridControl)     'returns 0 for success
            If rc = 0 Then gxMktGridRefresh m.gridControl
            
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.tbToolbar_ToolClick"

End Sub

Private Function GridTextChanged(eNewTextSpecs As ExGrid_Text_Specs, _
    ByVal strNewFontName$, ByVal bNotify As Boolean) As Boolean
On Error GoTo ErrSection:
    
    Dim bChanged As Boolean, rc&
    
    If gdGetStr(m.gridTextSpecs.gshFontName, 0) <> strNewFontName Then
        gdSetStr m.gridTextSpecs.gshFontName, 0, strNewFontName
        bChanged = True
    End If
    If m.gridTextSpecs.FontBold <> eNewTextSpecs.FontBold Then
        m.gridTextSpecs.FontBold = eNewTextSpecs.FontBold
        bChanged = True
    End If
    If m.gridTextSpecs.FontItalic <> eNewTextSpecs.FontItalic Then
        m.gridTextSpecs.FontItalic = eNewTextSpecs.FontItalic
        bChanged = True
    End If
    If m.gridTextSpecs.FontSize <> eNewTextSpecs.FontSize Then
        m.gridTextSpecs.FontSize = eNewTextSpecs.FontSize
        bChanged = True
    End If
    
    If bChanged And bNotify Then
        rc = SetGridTextSpecs()
        If rc <> 0 Then bChanged = False
    End If
    
    GridTextChanged = bChanged

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmMarketProfile.GridTextChanged"

End Function

Private Sub tbToolbar_ToolKeyDown(ByVal Tool As ActiveToolBars.SSTool, ByVal KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next:

    If KeyCode = vbKeyReturn Then
        If Tool.ID = "ID_Scale" Then tbToolbar_ComboCloseUp Tool
    End If

End Sub

Private Sub tmr_Timer()
On Error GoTo ErrSection:

    Static bInprog As Boolean
    
    Dim rc&, d1#, d2#, dDiff#
    
    Dim bNewBar As Boolean
    Dim bSuccess As Boolean
    Dim bLocked As Boolean
    
    TimerStart "frmMarketProfile.tmr"
    If g.bUnloading Or m.bInitialShow Then Exit Sub
    
    If bInprog Then Exit Sub
    
    If m.bUnloadNow Or m.bInitialShow Then GoTo ErrExit
    If FormIsLoaded("frmMarketProfileCfg") Then GoTo ErrExit
    
    bInprog = True
    
    If tbToolbar.Tools("ID_Scale").ComboBox.Text = "Auto" Then SetScaleDropdown
    
    If m.BarsProfile Is Nothing Then
        InitialShow
        GoTo ErrExit
    End If
    If m.BarsProfile.Size = 0 Then
        InitialShow
        GoTo ErrExit
    End If
    
    If m.bCurrentSession Then
        If g.RealTime.UpdateBars(m.BarsMinute, bNewBar) Then
            g.RealTime.SpliceBars m.BarsMinute
            If m.profileProp.ColorScheme = MktProf_Color_VolIterator Then UpdateVolIterator False
        End If
        
        If bNewBar And m.bSMPProfile Then
            d1 = m.BarsSMP(eBARS_DateTime, m.BarsSMP.Size - 2)
            d2 = m.BarsMinute(eBARS_DateTime, m.BarsMinute.Size - 1)
            dDiff = (d2 - d1) * 1440
            If Int(dDiff) > m.nProfileInterval * m.nSMPMinTPO Then
                bLocked = LockWindowUpdate(Me.pbExgrid.hWnd)
                    InitialShow -1, False
                If bLocked Then LockWindowUpdate 0
            Else
                rc = ProfileStream(False)
                If -1 = rc Then
                    bLocked = LockWindowUpdate(Me.pbExgrid.hWnd)
                        InitialShow -1, False
                    If bLocked Then LockWindowUpdate 0
                End If
            End If
        Else
            rc = ProfileStream(False)
            If -1 = rc Then
                bLocked = LockWindowUpdate(Me.pbExgrid.hWnd)
                    InitialShow -1, True
                If bLocked Then LockWindowUpdate 0
            End If
        End If
    End If
    
    If tmr.Interval <> 250 Then tmr.Interval = 250
    TimerEnd "frmMarketProfile.tmr", tmr.Interval

ErrExit:
    If m.bUnloadNow Then Unload Me
    bInprog = False
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.tmr_Timer"

End Sub

Private Function ChangeSymbol(Optional ByVal nSymID& = 0) As Boolean
On Error GoTo ErrSection:

    Dim astrSymbols As cGdArray     ' Symbol(s) back from the symbol selector
    Dim lSymbolID As Long               ' Symbol ID for the symbol selected
    Dim rc&, strMsg$
    Dim bGetData As Boolean
    Dim bGetSettings As Boolean
    
    tmr.Enabled = False
    
    If nSymID = 0 Then
        Set astrSymbols = frmSymbolSelector.ShowMe("", False)
        If astrSymbols.Size > 0 Then
            lSymbolID = g.SymbolPool.SymbolIDforSymbol(astrSymbols(0))
            If m.nSymbolID <> 0 Then
                If lSymbolID <> m.nSymbolID And Len(GetSymbol(lSymbolID)) > 0 Then
                    SaveSettings
                    bGetSettings = True
                End If
            End If
        End If
    Else
        lSymbolID = nSymID
    End If
    If lSymbolID = 0 Then
        Beep
    Else
        m.strSymbol = GetSymbol(lSymbolID)
        m.nSymbolID = lSymbolID
        If bGetSettings Then LoadSettings
        
        If IsForex(m.strSymbol) Then        '4599
            strMsg = "This feature does not support Forex symbols."
        ElseIf Left(m.strSymbol, 1) = "$" Then
            strMsg = "This feature does not support index symbols."
        Else
            bGetData = True
            strMsg = "Loading data for " & m.strSymbol & " ..."
        End If
        
        With fg
            .Redraw = flexRDNone
            .FixedCols = 0
            .FixedRows = 0
            .Rows = .FixedRows '(to clear all rows)
            .Rows = .FixedRows + 1
            .TextMatrix(.FixedRows, 0) = strMsg         '6191
            .MergeCells = flexMergeSpill
            .Redraw = flexRDBuffered
            .Refresh
            .Visible = True
        End With
        
        ShowForm Me, eForm_Nonmodal, frmMain
        FormResize Me

'If IsIDE Then
'    Me.WindowState = vbMinimized
'End If
        DoEvents
        
        If bGetData Then
            rc = InitialShow()
            If rc = 0 Then
                frmMain.SetWindowLink Me
                tmr.Interval = 1000
                tmr.Enabled = True
            End If
        End If
    End If

ErrExit:
    ChangeSymbol = bGetData
    Exit Function

ErrSection:
    RaiseError "frmMarketProfile.ChangeSymbol"

End Function

Private Function LoadSettings() As Boolean
On Error GoTo ErrSection:

    Dim strFile$, strSection$
    
    Dim bSuccess As Boolean
    
    bSuccess = True
    
    'trade profile always start with current session, but user can change the end date to something prior
    'this flag is only changed when user explicitly select a specific date in the config dialog
    m.bCurrentSession = True
    
    tbToolbar.Tools("ID_ToggleGrid").Visible = False
    m.bSMPEnabled = HasModule("SMPPI")

    'grid general settings
    m.gridExtendedProp.gridBkColor = GetIniFileProperty("GridBkColor", vbWhite, "Market Profile", g.strIniFile)
    m.gridExtendedProp.GridLines = GetIniFileProperty("GridLines", 1, "Market Profile", g.strIniFile)
    
    'grid text settings
    m.gridTextSpecs.Color = GetIniFileProperty("GridTextColor", vbBlack, "Market Profile", g.strIniFile)
    m.gridTextSpecs.FontBold = GetIniFileProperty("GridFontBold", 0, "Market Profile", g.strIniFile)
    m.gridTextSpecs.FontItalic = GetIniFileProperty("GridFontItalic", 0, "Market Profile", g.strIniFile)
    m.gridTextSpecs.FontUnderline = GetIniFileProperty("GridFontUnderline", 0, "Market Profile", g.strIniFile)
    m.gridTextSpecs.FontSize = GetIniFileProperty("GridFontSize", 8, "Market Profile", g.strIniFile)

    If m.gridTextSpecs.gshFontName = 0 Then
        m.gridTextSpecs.gshFontName = gdCreateArray(eGDARRAY_gdString)
    End If
    If m.gridTextSpecs.gshFontName <> 0 Then
        gdSetStr m.gridTextSpecs.gshFontName, 0, GetIniFileProperty("GridFontName", "MS Sans Serif", "Market Profile", g.strIniFile)
    End If
    
    'Steidlmayer minimum & breakout periods
    m.nSMPColor = GetIniFileProperty("SMPColor", vbBlue, "Market Profile", g.strIniFile)
'    m.nSMPMinTPO = GetIniFileProperty("SMPMinTPO", 12, "Market Profile", g.strIniFile)
'    m.nSMPBreakOut = GetIniFileProperty("SMPBreakOut", 9, "Market Profile", g.strIniFile)
'    If m.bSMPEnabled Then m.bSMPProfile = GetIniFileProperty("SMPProfile", 0, "Market Profile", g.strIniFile)
    
    'volume iterator lookback & highvol specs
    m.nIteratorLookback = GetIniFileProperty("IteratorLookback", 300, "Market Profile", g.strIniFile)
    m.nIteratorHighVol = GetIniFileProperty("IteratorHighVol", 20, "Market Profile", g.strIniFile)
    m.nVolItColor0 = GetIniFileProperty("VolIteratorColor0", 32768, "Market Profile", g.strIniFile)
    m.nVolItColor1 = GetIniFileProperty("VolIteratorColor1", vbRed, "Market Profile", g.strIniFile)
    
    'profile display settings
    m.profileProp.ColorScheme = GetIniFileProperty("ProfileColorScheme", MktProf_Color_Gradient, "Market Profile", g.strIniFile)
    m.profileProp.CharSequence = GetIniFileProperty("ProfileCharSeq", 0, "Market Profile", g.strIniFile)
    
    m.nGradientFrom = GetIniFileProperty("TPOGradientFrom", vbYellow, "Market Profile", g.strIniFile)
    m.nGradientTo = GetIniFileProperty("TPOGradientTo", vbBlack, "Market Profile", g.strIniFile)
    m.nUpColor = GetIniFileProperty("TPOUpColor", vbGreen, "Market Profile", g.strIniFile)
    m.nDownColor = GetIniFileProperty("TPODownColor", vbRed, "Market Profile", g.strIniFile)
    m.nBidColor = GetIniFileProperty("TPOBidColor", vbRed, "Market Profile", g.strIniFile)
    m.nAskColor = GetIniFileProperty("TPOAskColor", vbBlue, "Market Profile", g.strIniFile)
    
    m.nShadeVolume = GetIniFileProperty("ProfileShadeVol", 1, "Market Profile", g.strIniFile)
    m.nVolumeColor = GetIniFileProperty("ProfileVolColor", 12648447, "Market Profile", g.strIniFile)  'light yellow
    m.profileProp.VolumeText = GetIniFileProperty("ProfileVolText", MktProf_Text_Percent, "Market Profile", g.strIniFile)
    If m.nShadeVolume = 0 Then
        m.profileProp.VolumeColor = -1
    Else
        m.profileProp.VolumeColor = m.nVolumeColor
    End If
    
    m.nMarginLeft = GetIniFileProperty("ProfileMarginLeft", 0, "Market Profile", g.strIniFile)
    m.nMarginRight = GetIniFileProperty("ProfileMarginRight", 0, "Market Profile", g.strIniFile)
    m.profileProp.ExtraRows = GetIniFileProperty("ProfileExtraRows", 15, "Market Profile", g.strIniFile)
    m.profileProp.BoxFirst = GetIniFileProperty("ProfileBoxFirst", 0, "Market Profile", g.strIniFile)
    
    'statistics settings
    m.statsProp.POC_Vol_Color = GetIniFileProperty("POCVolColor", kColorPurple, "Market Profile", g.strIniFile)
    m.statsProp.POC_Vol_PenSize = GetIniFileProperty("POCVolPenSize", 2, "Market Profile", g.strIniFile)
    
    m.statsProp.POC_TPO_Color = GetIniFileProperty("POCTPOColor", vbGreen, "Market Profile", g.strIniFile)
    m.statsProp.POC_TPO_PenSize = GetIniFileProperty("POCTPOPenSize", 2, "Market Profile", g.strIniFile)
    
    m.statsProp.VA_Vol_Color = GetIniFileProperty("VAVolColor", vbBlue, "Market Profile", g.strIniFile)
    m.statsProp.VA_Vol_PenSize = GetIniFileProperty("VAVolPenSize", 2, "Market Profile", g.strIniFile)
    m.statsProp.VA_Vol_Percent = GetIniFileProperty("VAVolPercent", 70, "Market Profile", g.strIniFile)
    
    m.statsProp.VA_TPO_Color = GetIniFileProperty("VATPOColor", vbCyan, "Market Profile", g.strIniFile)
    m.statsProp.VA_TPO_PenSize = GetIniFileProperty("VATPOPenSize", 2, "Market Profile", g.strIniFile)
    m.statsProp.VA_TPO_Percent = GetIniFileProperty("VATPOPercent", 70, "Market Profile", g.strIniFile)
    
    m.statsProp.IB_Color = GetIniFileProperty("IBColor", vbMagenta, "Market Profile", g.strIniFile)
    m.statsProp.IB_PenSize = GetIniFileProperty("IBPenSize", 2, "Market Profile", g.strIniFile)
    
    m.statsProp.StdDev_Color = GetIniFileProperty("StdDevColor", vbRed, "Market Profile", g.strIniFile)
    m.statsProp.StdDev_PenSize = GetIniFileProperty("StdDevPenSize", 2, "Market Profile", g.strIniFile)
    
    m.statsProp.Mean_Color = GetIniFileProperty("MeanColor", kColorGreen, "Market Profile", g.strIniFile)
    m.statsProp.Mean_PenSize = GetIniFileProperty("MeanPenSize", 2, "Market Profile", g.strIniFile)
        
    'profile data settings
    m.nDaysBack = GetIniFileProperty("DaysBack", 0, "Market Profile", g.strIniFile)
    m.nDaysProfile = GetIniFileProperty("DaysProfile", 1, "Market Profile", g.strIniFile)
    m.nProfileInterval = GetIniFileProperty("ProfileInterval", 30, "Market Profile", g.strIniFile)
    m.nIntervalType = GetIniFileProperty("IntervalType", 0, "Market Profile", g.strIniFile)
    m.nYScaleMinMove = GetIniFileProperty("YScaleMinMove", 1, "Market Profile", g.strIniFile)
    
    strFile = g.strAppPath & "\custom\TradeProfile.Cfg"
    If FileExist(strFile) Then
        strSection = BaseForAutoExitFavorites(m.strSymbol)
        If Len(strSection) > 0 Then
            m.nDaysBack = GetIniFileProperty("DaysBack", m.nDaysBack, strSection, strFile)
            m.nDaysProfile = GetIniFileProperty("DaysProfile", m.nDaysProfile, strSection, strFile)
            m.nProfileInterval = GetIniFileProperty("ProfileInterval", m.nProfileInterval, strSection, strFile)
            m.nIntervalType = GetIniFileProperty("IntervalType", m.nIntervalType, strSection, strFile)
            m.dTimeStart = GetIniFileProperty("CustomStartTime", 0, strSection, strFile)
            m.dTimeEnd = GetIniFileProperty("CustomEndTime", 0, strSection, strFile)
            m.nYScaleMinMove = GetIniFileProperty("YScaleMinMove", 1, strSection, strFile)
        
            m.nSMPMinTPO = GetIniFileProperty("SMPMinTPO", 12, strSection, strFile)
            m.nSMPBreakOut = GetIniFileProperty("SMPBreakOut", 9, strSection, strFile)
            If m.bSMPEnabled Then m.bSMPProfile = GetIniFileProperty("SMPProfile", 0, strSection, strFile)
        End If
    End If
    
    If bSuccess Then ChooseCharSeq
    
    'turn on gdProfile logging / viewing
    If FileExist("ExgridProfilerOn.flg") Then
        m.gridExtendedProp.gdProfiler = 1
        tbToolbar.Tools("ID_ToggleGrid").Visible = True
    Else
        If GetIniFileProperty("ShowProfileBtn", 0, "Market Profile", g.strIniFile) = 1 Then
            tbToolbar.Tools("ID_ToggleGrid").Visible = True
        End If
    End If
    
    LoadSettings = bSuccess
        
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmMarketProfile.LoadSettings"

End Function

Private Sub SaveSettings()
On Error GoTo ErrSection:
    
    Dim strFile$, strSection$
    
    'grid general settings
    SetIniFileProperty "GridBkColor", m.gridExtendedProp.gridBkColor, "Market Profile", g.strIniFile
    SetIniFileProperty "GridLines", m.gridExtendedProp.GridLines, "Market Profile", g.strIniFile
    
    'grid text settings
    SetIniFileProperty "GridTextColor", m.gridTextSpecs.Color, "Market Profile", g.strIniFile
    SetIniFileProperty "GridFontBold", m.gridTextSpecs.FontBold, "Market Profile", g.strIniFile
    SetIniFileProperty "GridFontItalic", m.gridTextSpecs.FontItalic, "Market Profile", g.strIniFile
    SetIniFileProperty "GridFontUnderline", m.gridTextSpecs.FontUnderline, "Market Profile", g.strIniFile
    SetIniFileProperty "GridFontSize", m.gridTextSpecs.FontSize, "Market Profile", g.strIniFile
    
    If m.gridTextSpecs.gshFontName <> 0 Then
        SetIniFileProperty "GridFontName", gdGetStr(m.gridTextSpecs.gshFontName, 0), "Market Profile", g.strIniFile
    End If
    
    'Steidlmayer profile
    SetIniFileProperty "SMPColor", m.nSMPColor, "Market Profile", g.strIniFile
'    SetIniFileProperty "SMPMinTPO", m.nSMPMinTPO, "Market Profile", g.strIniFile
'    SetIniFileProperty "SMPBreakOut", m.nSMPBreakOut, "Market Profile", g.strIniFile
'    If m.bSMPEnabled Then SetIniFileProperty "SMPProfile", m.bSMPProfile, "Market Profile", g.strIniFile
    
    'volume iterator
    SetIniFileProperty "IteratorLookback", m.nIteratorLookback, "Market Profile", g.strIniFile
    SetIniFileProperty "IteratorHighVol", m.nIteratorHighVol, "Market Profile", g.strIniFile
    SetIniFileProperty "VolIteratorColor0", m.nVolItColor0, "Market Profile", g.strIniFile
    SetIniFileProperty "VolIteratorColor1", m.nVolItColor1, "Market Profile", g.strIniFile
    
    'normal profile display settings
    SetIniFileProperty "ProfileColorScheme", m.profileProp.ColorScheme, "Market Profile", g.strIniFile
    SetIniFileProperty "ProfileCharSeq", m.profileProp.CharSequence, "Market Profile", g.strIniFile
    
    SetIniFileProperty "TPOGradientFrom", m.nGradientFrom, "Market Profile", g.strIniFile
    SetIniFileProperty "TPOGradientTo", m.nGradientTo, "Market Profile", g.strIniFile
    SetIniFileProperty "TPOUpColor", m.nUpColor, "Market Profile", g.strIniFile
    SetIniFileProperty "TPODownColor", m.nDownColor, "Market Profile", g.strIniFile
    SetIniFileProperty "TPOBidColor", m.nBidColor, "Market Profile", g.strIniFile
    SetIniFileProperty "TPOAskColor", m.nAskColor, "Market Profile", g.strIniFile
    
    SetIniFileProperty "ProfileVolText", m.profileProp.VolumeText, "Market Profile", g.strIniFile
    SetIniFileProperty "ProfileShadeVol", m.nShadeVolume, "Market Profile", g.strIniFile
    SetIniFileProperty "ProfileVolColor", m.nVolumeColor, "Market Profile", g.strIniFile
    
    SetIniFileProperty "ProfileMarginLeft", m.nMarginLeft, "Market Profile", g.strIniFile
    SetIniFileProperty "ProfileMarginRight", m.nMarginRight, "Market Profile", g.strIniFile
    SetIniFileProperty "ProfileExtraRows", m.profileProp.ExtraRows, "Market Profile", g.strIniFile
    SetIniFileProperty "ProfileBoxFirst", m.profileProp.BoxFirst, "Market Profile", g.strIniFile
    
    'statistics prop
    SetIniFileProperty "POCVolColor", m.statsProp.POC_Vol_Color, "Market Profile", g.strIniFile
    SetIniFileProperty "POCVolPenSize", m.statsProp.POC_Vol_PenSize, "Market Profile", g.strIniFile
    
    SetIniFileProperty "POCTPOColor", m.statsProp.POC_TPO_Color, "Market Profile", g.strIniFile
    SetIniFileProperty "POCTPOPenSize", m.statsProp.POC_TPO_PenSize, "Market Profile", g.strIniFile
    
    SetIniFileProperty "VAVolColor", m.statsProp.VA_Vol_Color, "Market Profile", g.strIniFile
    SetIniFileProperty "VAVolPenSize", m.statsProp.VA_Vol_PenSize, "Market Profile", g.strIniFile
    SetIniFileProperty "VAVolPercent", m.statsProp.VA_Vol_Percent, "Market Profile", g.strIniFile
    
    SetIniFileProperty "VATPOColor", m.statsProp.VA_TPO_Color, "Market Profile", g.strIniFile
    SetIniFileProperty "VATPOPenSize", m.statsProp.VA_TPO_PenSize, "Market Profile", g.strIniFile
    SetIniFileProperty "VATPOPercent", m.statsProp.VA_TPO_Percent, "Market Profile", g.strIniFile
    
    SetIniFileProperty "IBColor", m.statsProp.IB_Color, "Market Profile", g.strIniFile
    SetIniFileProperty "IBPenSize", m.statsProp.IB_PenSize, "Market Profile", g.strIniFile
    
    SetIniFileProperty "StdDevColor", m.statsProp.StdDev_Color, "Market Profile", g.strIniFile
    SetIniFileProperty "StdDevPenSize", m.statsProp.StdDev_PenSize, "Market Profile", g.strIniFile
    
    SetIniFileProperty "MeanColor", m.statsProp.Mean_Color, "Market Profile", g.strIniFile
    SetIniFileProperty "MeanPenSize", m.statsProp.Mean_PenSize, "Market Profile", g.strIniFile
    
    strFile = g.strAppPath & "\custom\TradeProfile.Cfg"
    strSection = BaseForAutoExitFavorites(m.strSymbol)
    
    If Len(strSection) > 0 Then
        SetIniFileProperty "DaysBack", m.nDaysBack, strSection, strFile
        SetIniFileProperty "DaysProfile", m.nDaysProfile, strSection, strFile
        SetIniFileProperty "ProfileInterval", m.nProfileInterval, strSection, strFile
        SetIniFileProperty "IntervalType", m.nIntervalType, strSection, strFile
        SetIniFileProperty "YScaleMinMove", m.nYScaleMinMove, strSection, strFile
        
        SetIniFileProperty "SMPMinTPO", m.nSMPMinTPO, strSection, strFile
        SetIniFileProperty "SMPBreakOut", m.nSMPBreakOut, strSection, strFile
        If m.bSMPEnabled Then SetIniFileProperty "SMPProfile", m.bSMPProfile, strSection, strFile
        
        If Not m.BarsMinute Is Nothing Then
            If m.dTimeStart = m.BarsMinute.Prop(eBARS_DefaultStartTime) Then
                SetIniFileProperty "CustomStartTime", 0, strSection, strFile
            ElseIf m.dTimeStart > 0 Then
                SetIniFileProperty "CustomStartTime", m.dTimeStart, strSection, strFile
            End If
            
            If m.dTimeEnd = m.BarsMinute.Prop(eBARS_DefaultEndTime) Then
                SetIniFileProperty "CustomEndTime", 0, strSection, strFile
            ElseIf m.dTimeEnd > 0 Then
                SetIniFileProperty "CustomEndTime", m.dTimeEnd, strSection, strFile
            End If
        End If
    Else
        'profile data settings
        SetIniFileProperty "DaysBack", m.nDaysBack, "Market Profile", g.strIniFile
        SetIniFileProperty "DaysProfile", m.nDaysProfile, "Market Profile", g.strIniFile
        SetIniFileProperty "ProfileInterval", m.nProfileInterval, "Market Profile", g.strIniFile
        SetIniFileProperty "IntervalType", m.nIntervalType, "Market Profile", g.strIniFile
        SetIniFileProperty "YScaleMinMove", m.nYScaleMinMove, "Market Profile", g.strIniFile
    End If
                
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.SaveSettings"

End Sub

Private Function SetGridTextSpecs() As Long
On Error GoTo ErrSection:
    
    Dim rc As Long
        
    rc = 1              '0 means success
    If m.gridControl <> 0 Then
        rc = gxMktGridTextSpecs(m.gridControl, m.gridTextSpecs)
    End If
    
    SetGridTextSpecs = rc

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmMarketProfile.SetGridTextSpecs"

End Function

Private Function SetProfileSpecs() As Long
On Error GoTo ErrSection:

    Dim rc As Long
    Dim SMPProp As MktProfile_Stats_Prop
    
    rc = 1              '0 means success
    
    If m.gridControl <> 0 Then
        rc = gxMktGridExtendedProp(m.gridControl, m.gridExtendedProp)
        
        If rc = 0 Then rc = gxMktProfileProperties(m.gridControl, m.profileProp)
        
        If rc = 0 Then
            If m.bSMPProfile Then
                SMPProp.IB_Color = -1
                SMPProp.Mean_Color = -1
                SMPProp.POC_TPO_Color = -1
                SMPProp.POC_Vol_Color = -1
                SMPProp.VA_Vol_Color = -1
                SMPProp.VA_TPO_Color = -1
                SMPProp.StdDev_Color = -1
                rc = gxMktStatisticsProp(m.gridControl, SMPProp)
            Else
                rc = gxMktStatisticsProp(m.gridControl, m.statsProp)
            End If
        End If
    End If
    
    SetProfileSpecs = rc
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmMarketProfile.SetProfileSpecs"

End Function

Private Function SetColorSchemeDLL() As Long

    Dim rc&
    
    If m.bSMPProfile And m.profileProp.ColorScheme <> MktProf_Color_VolIterator Then
        m.profileProp.Color1 = m.nSMPColor
    Else
        Select Case m.profileProp.ColorScheme
            Case MktProf_Color_OpenClose
                m.profileProp.Color1 = m.nUpColor
                m.profileProp.Color2 = m.nDownColor
            Case MktProf_Color_BidAsk
                m.profileProp.Color1 = m.nBidColor
                m.profileProp.Color2 = m.nAskColor
            Case MktProf_Color_VolIterator
                rc = gxMktVolIteratorBars(m.gridControl, m.BarsMinute.BarsHandle)
                If rc = 0 Then
                    m.profileProp.Color1 = m.nVolItColor0
                    m.profileProp.Color2 = m.nVolItColor1
                Else
                    InfBox "gxMktVolIteratorBars failed. Switching to alternate color scheme.", "E", , "Trade PRofile"
                    m.profileProp.ColorScheme = MktProf_Color_Rainbow
                    
                    If m.bSMPProfile Then
                        m.profileProp.Color1 = m.nSMPColor
                    End If
                End If
            Case Else
                m.profileProp.Color1 = m.nGradientFrom
                m.profileProp.Color2 = m.nGradientTo
        End Select
    End If
    
    SetColorSchemeDLL = rc

End Function

Private Sub ChooseCharSeq()
On Error GoTo ErrSection:
'Per Pete: user must choose character sequence on first time use for legal reason

    Dim iAlreadyDone&
    
    iAlreadyDone = GetIniFileProperty("CharSeqChosen", 0, "Market Profile", g.strIniFile)
    If iAlreadyDone = 0 Then
        If m.bSMPEnabled Then
            m.statsProp.IB_Color = -1 * m.statsProp.IB_Color
            m.statsProp.Mean_Color = -1 * m.statsProp.Mean_Color
            m.statsProp.POC_TPO_Color = -1 * m.statsProp.POC_TPO_Color
            m.statsProp.POC_Vol_Color = -1 * m.statsProp.POC_Vol_Color
            m.statsProp.VA_Vol_Color = -1 * m.statsProp.VA_Vol_Color
            m.statsProp.VA_TPO_Color = -1 * m.statsProp.VA_TPO_Color
            m.statsProp.StdDev_Color = -1 * m.statsProp.StdDev_Color
            m.profileProp.VolumeText = MktProf_Text_None
            ProfileProperty(eProfileProp_VolShade) = 0
            
            m.gridExtendedProp.GridLines = 0    'default to no grid lines for SMP profile
        End If
        m.bInitialShow = True
        frmMarketProfileCfg.ShowMe Me, Nothing, m.bSMPEnabled
        SetIniFileProperty "CharSeqChosen", 1, "Market Profile", g.strIniFile
        SaveSettings
        m.bInitialShow = False
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.ChooseCharSeq"

End Sub

Public Property Get SymbolID() As Long
    SymbolID = m.nSymbolID
End Property

Public Property Let SymbolID(ByVal nSymbolID As Long)
On Error GoTo ErrSection:
    
    ChangeSymbol nSymbolID

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmMarketProfile.LetSymbolID", eGDRaiseError_Raise

End Property

Public Sub RefreshData()
    If Not g.RealTime.Active Then InitialShow
End Sub

Public Property Get WindowLink() As cWindowLink
On Error GoTo ErrSection:

    Set WindowLink = m.WindowLink

    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.WindowLinkGet"
End Property

Public Sub UpdateSettings(ByVal strFontName$, ByVal nFontSize&, ByVal nBold&, ByVal nItalic&, _
    ByVal bSMPProfile As Boolean)
On Error GoTo ErrSection:

    Dim i&, rc&, strTemp$
    Dim bBarsChanged As Boolean
    Dim bTextChanged As Boolean

    If m.bSMPProfile <> bSMPProfile Then
        m.bSMPProfile = bSMPProfile
        m.bReload = True
    End If
    
    'font information
    If GridFontName <> strFontName Then
        gdSetStr m.gridTextSpecs.gshFontName, 0, strFontName
        bTextChanged = True
    End If
    If m.gridTextSpecs.FontBold <> nBold Then
        m.gridTextSpecs.FontBold = nBold
        bTextChanged = True
    End If
    If m.gridTextSpecs.FontItalic <> nItalic Then
        m.gridTextSpecs.FontItalic = nItalic
        bTextChanged = True
    End If
    If m.gridTextSpecs.FontSize <> nFontSize Then
        m.gridTextSpecs.FontSize = nFontSize
        bTextChanged = True
    End If
    
    If Not m.BarsProfile Is Nothing Then
        If m.dTimeStart <> m.BarsProfile.Prop(eBARS_StartTime) Then
            m.dTimeStart = m.BarsProfile.Prop(eBARS_StartTime)
            m.bReload = True
            bBarsChanged = True
        End If
        If m.dTimeEnd <> m.BarsProfile.Prop(eBARS_EndTime) Then
            m.dTimeEnd = m.BarsProfile.Prop(eBARS_EndTime)
            m.bReload = True
            bBarsChanged = True
        End If
    End If
    
    SaveSettings        '6511 - if this is call after the If() block below then TN crashes on Windows 7 & Windows 8
    
    If m.bReload Then
        InitialShow
    ElseIf m.gridControl <> 0 Then
        If m.bRefresh And m.profileProp.ColorScheme = MktProf_Color_VolIterator Then UpdateVolIterator True
        SetColorSchemeDLL
        If m.bRefresh Then
            rc = SetProfileSpecs()
            If rc = 0 Then rc = SetGridTextSpecs()
        ElseIf bTextChanged Then
            rc = SetGridTextSpecs()
        End If
        If rc = 0 Then rc = gxMktCenterPrice(m.gridControl)
        If rc = 0 Then gxMktGridRefresh m.gridControl
    End If
    
    m.bRefresh = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarketProfile.LetSettings", eGDRaiseError_Raise

End Sub

Public Sub GridTextIncrease()
On Error Resume Next

    tbToolbar_ToolClick Me.tbToolbar.Tools("ID_TextIncrease")

End Sub

Public Sub GridTextDecrease()
On Error Resume Next

    tbToolbar_ToolClick Me.tbToolbar.Tools("ID_TextDecrease")

End Sub

Public Sub DLLStatusInfo(ByVal nStatus&, ByVal hString&)
On Error Resume Next

    Dim strText$, strProfile$

    If hString <> 0 Then
    
        strText = gdGetStr(hString, 0)
        
        If Len(strText) > 0 Then
            If nStatus = 1 And Not fg.Visible Then
                pbExgrid.Visible = False
                fg.Visible = True
            ElseIf nStatus = 5 And fg.Visible Then
                pbExgrid.Visible = True
                fg.Visible = False
            End If
            
            With fg
                If .Visible Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = strText
                    .ShowCell .Rows - 1, 0
                End If
            End With
            
            DoEvents
        End If
    
    End If

End Sub

Public Sub ProfileBarsGetStatus(ByVal strMsg$)
On Error Resume Next

    If m.bInitialShow And fg.Visible Then
        With fg
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = strMsg
            .ShowCell .Rows - 1, 0
            DoEvents
        End With
    End If

End Sub

Public Property Get YScaleMinMove() As Long
On Error GoTo ErrSection:

    YScaleMinMove = m.nYScaleMinMove

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.YScaleMinMove.Get"

End Property

Public Property Let YScaleMinMove(ByVal nValue&)
On Error GoTo ErrSection:

    If nValue <> m.nYScaleMinMove Then
        m.nYScaleMinMove = nValue
        m.bReload = True
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMarketProfile.YScaleMinMove.Let"

End Property

Private Sub UpdateVolIterator(ByVal bShowMsg As Boolean)

    If m.BarsMinute Is Nothing Then Exit Sub
    
    If m.nIteratorLookback <= 0 Then m.nIteratorLookback = 300
    If m.nIteratorHighVol <= 0 Then m.nIteratorHighVol = 20
    
    If m.nIteratorHighVol > m.nIteratorLookback Then
        If bShowMsg And m.profileProp.ColorScheme = MktProf_Color_VolIterator Then
            InfBox "High volume periods cannot be larger than lookback bars for Volume Iterator. Please edit the settings and try again.", "I", , "Trade Profile"
            Exit Sub
        End If
    End If

    If bShowMsg Then
        fg.Rows = fg.Rows + 1
        fg.TextMatrix(fg.Rows - 1, 0) = "Building volume iterator."
    End If

    If Not GetVolumeIterators(m.BarsMinute, m.nIteratorLookback, m.nIteratorHighVol) Then
        If m.BarsMinute.Size < m.nIteratorLookback Then
            If bShowMsg Then
                If m.profileProp.ColorScheme = MktProf_Color_VolIterator Then
                    InfBox "The volume iterator returned no high volume periods. You may not have enough data for the specified lookback bars. Adjust the number of days back or lookback periods and try again.", "I", , "Trade Profile"
                    fg.Rows = fg.Rows + 1
                    fg.TextMatrix(fg.Rows - 1, 0) = "The volume iterator returned no high volume periods."
                End If
            End If
        Else
            DebugLog "Trade Profile GetVolumeIterators failed: " & m.BarsMinute.Prop(eBARS_Symbol)
            fg.Rows = fg.Rows + 1
            fg.TextMatrix(fg.Rows - 1, 0) = "GetVolumeIterators returned false."
        End If
    End If
    
End Sub

Public Property Get IsSMPProfile() As Boolean
    IsSMPProfile = m.bSMPProfile
End Property

Private Sub SetScaleDropdown()
On Error GoTo ErrSection

    If m.nYScaleMinMove < 0 Then
        tbToolbar.Tools("ID_Scale").ComboBox.Text = "Auto (" & Str(m.gridExtendedProp.reserved3) & " ticks)"
    Else
        tbToolbar.Tools("ID_Scale").ComboBox.Text = Str(Int(m.nYScaleMinMove)) & " ticks"
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfile.SetScaleDropdown"

End Sub

Private Function ProfileStream(ByVal bInit As Boolean) As Long
On Error GoTo ErrExit

    Dim eStatus As eProfileStartusRT
    
    Dim BarsProfileRT As New cGdBars
    
    Dim nSessionDate&, nSessionDateNext&, nSessionDateRecent&, i&
    
    nSessionDate = m.BarsMinute.SessionDateForTime(m.nSessionDate, True)
    nSessionDateNext = m.BarsMinute.SessionDateForTime(m.dProfileEnd, True) + 1    'next session date
    nSessionDateRecent = m.BarsMinute.SessionDateForTime(m.BarsMinute(eBARS_DateTime, m.BarsMinute.Size - 1), True)  'most recent session date
    
    eStatus = ProfileUpdateRT(m.BarsTicks, m.BarsProfile, BarsProfileRT, nSessionDate, nSessionDateNext, nSessionDateRecent, bInit)
                
    If eStatus = eProfileRT_NewData And Not BarsProfileRT Is Nothing Then
        i = gxMktUpdateRT(m.gridControl, BarsProfileRT.BarsHandle, 0)
    End If

ErrExit:
    ProfileStream = eStatus
    Exit Function

ErrSection:
    RaiseError "frmMarketProfile.ProfileStream"

End Function

