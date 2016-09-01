VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRealTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eSalmonResponse
    eSalmonNotRequested = 0
    eSalmonPending = 1
    eSalmonAvailable = 2
    eSalmonError = 3
End Enum

Private Enum eLastQuote
    eLQ_TradeTime = 0
    eLQ_TradePrice
    eLQ_TradeVol
    eLQ_BidAskTime
    eLQ_BidPrice
    eLQ_BidSize
    eLQ_AskPrice
    eLQ_AskSize
    eLQ_MinMove
    eLQ_PastGoodPrice
    eLQ_NUM_ITEMS
End Enum
 
Private Type mQuoteMsg
    strSymbol As String * 24
    strExch As String * 8
    dDateTime As Double
    dPrice As Double
    dPrice2 As Double
    nSize As Long
    nSize2 As Long
    nFlags As Long
    strFlags As String * 4
End Type

Private Type mPrivate
    RT As gdAppMail
    nConnectionStatus As eGDConnectionStatus
    iRtInterval As Long ' how often GenesisRT will send messages
    
    strServer As String
    strServerNode As String
    nServerStatus As Long
    dFeedTime As Double     ' time of feed (probably NY time)
    dLastFeedTimeMsg As Double ' gdTickCount of last message from RT client
    dLastMsgTime As Double  ' gdTickCount of last message from RT client
    dQuoteCount As Double   ' use a double so no chance of overflowing a long
    
    bDirtySymbolList As Boolean
    TickBuffers As cGdTree      ' collection of cGdBars (kept sorted by symbol) for new ticks
    TickBufferSymbols As cGdArray     ' sorted array of symbols (one-to-one correspondence with buffers)
    UniqueSymbols As cGdArray   ' sorted array of unique symbols
    TickRetrieval As cGdTree    ' collection of cGdBars for retrieving sets of ticks
    
    ' collection of the most recent bid, ask, and last trade for each symbol
    LastQuote As cGdTree
    
    RecentSymbols As cGdArray  ' recently used symbols (from past week)
    
    RtSymbolInfo As cGdArray    ' RT info for each symbol (e.g. delay, feed symbol, MktDepth enablement)
    
    nBrokerFeed As eTT_AccountType ' which feed to use for the BrokerRtSymbols
    BrokerRtSymbols As cGdArray ' sorted string array of symbols with broker real-time access
    BrokerSubscriptionList As cGdArray
    
    PastGoodPrice As cGdTree    ' last good price for each base symbol (futures only) -- for FixTicks
    
    MktDepthList As cGdTree
    
    bDebug As Boolean
    bRefreshAllSymbols As Boolean
    strIPandKey As String
    strSplitterMsg As String
    
    dEminiOffset As Double
    bReconnecting As Boolean
    bSocketLogging As Boolean
    nTickHistRequestCount As Long
    nTickHistTimeoutCount As Long
    
    OptNavStreamingList As New cGdArray
    OptNavHistSymbols As New cGdArray

    ' only used when replaying a real-time stream for a previous day
    dReplaySpeed As Double ' 0=paused, 0.5=half, 1.0=normal, 2.0=double
    
    ' New Salmon stuff ...
    bSalmonStarted As Boolean
    nSalmonPort As Long
    SymbolInfo As cGdTree ' collection of cSymbolInfo (1 per symbol)
    
    strDisconnectReason As String
End Type
Private m As mPrivate

Public bDisableForTesting As Boolean ' TLB: this flag is just for testing/debugging

Private Declare Function Start_SalmonDLL Lib "SalmonClient.dll" Alias "Start" (ByVal hStringArray As Long) As Long
Private Declare Function Stop_SalmonDLL Lib "SalmonClient.dll" Alias "Stop" () As Long
Private Declare Function StartStreaming_SalmonDLL Lib "SalmonClient.dll" Alias "StartStreaming" (ByVal hStringArray As Long) As Long
Private Declare Function SetWindowHandle_SalmonDLL Lib "SalmonClient.dll" Alias "SetWindowHandle" (ByVal hStringArray As Long) As Long


Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Dim i&
    
    KillFile App.Path & "\Cache\*.*"

    Set m.SymbolInfo = New cGdTree
    Set m.OptNavHistSymbols = New cGdArray
    Set m.OptNavStreamingList = New cGdArray
    Set m.PastGoodPrice = New cGdTree
    Set m.MktDepthList = New cGdTree
    Set m.RtSymbolInfo = New cGdArray
    Set m.UniqueSymbols = New cGdArray
    Set m.BrokerSubscriptionList = New cGdArray
    Set m.BrokerRtSymbols = New cGdArray
    Set m.RecentSymbols = New cGdArray
    Set m.TickBuffers = New cGdTree
    Set m.TickBufferSymbols = New cGdArray
    Set m.TickRetrieval = New cGdTree
    Set m.LastQuote = New cGdTree
    Set m.RT = frmMain.apmRTClient
    With m.RT
        .Active = False
        .ControlName = "RTClient"
    End With
    
    m.nBrokerFeed = 0
    
    m.iRtInterval = 250 ' millisecond interval for GenesisRT to send newly streamed ticks
    
    m.bDebug = FileExist(AddSlash(App.Path) & "RTDBG.LOG")
    DebugLog "cRealTime object initialized"
    
    ' load the Recent Symbol list
    m.RecentSymbols.FromFile App.Path & "\Custom\RecentSymbols.Lst"
    For i = m.RecentSymbols.Size - 1 To 0 Step -1
        ' delete any that haven't been used in a week
        If Val(Parse(m.RecentSymbols(i), vbTab, 2)) < Now - 7 Then
            m.RecentSymbols.Remove i
        ElseIf InStr(m.RecentSymbols(i), "-05") > 0 Or InStr(m.RecentSymbols(i), "-06") > 0 Then
            m.RecentSymbols.Remove i
        End If
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.Initialize", eGDRaiseError_Raise
    
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Dim i&

    Active = False
    SalmonStop
    
    ' save the Recent Symbol list
    m.RecentSymbols.ToFile App.Path & "\Custom\RecentSymbols.Lst"
    
    Set m.SymbolInfo = Nothing
    Set m.RtSymbolInfo = Nothing
    Set m.RT = Nothing
    Set m.UniqueSymbols = Nothing
    Set m.BrokerRtSymbols = Nothing
    Set m.RecentSymbols = Nothing
    Set m.BrokerSubscriptionList = Nothing
    Set m.TickBuffers = Nothing
    Set m.TickBufferSymbols = Nothing
    Set m.TickRetrieval = Nothing
    Set m.LastQuote = Nothing
    Set m.MktDepthList = Nothing
    Set m.PastGoodPrice = Nothing
    Set m.OptNavStreamingList = Nothing
    Set m.OptNavHistSymbols = Nothing

    DebugLog "cRealTime object terminated"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.Terminate", eGDRaiseError_Raise
    
End Sub

Public Property Get ActiveRTG() As Boolean
    If m.RT.Active Then
        If UCase(m.strServerNode) = "GENESISRT" Then
            ActiveRTG = True
        End If
    End If
End Property

Public Property Get Active() As Boolean
    Active = m.RT.Active
End Property

' Starts/Stops the GenesisRT program (along with our AppMail connection to it)
Private Property Let Active(ByVal bActive As Boolean)
On Error GoTo ErrSection:
    
    Dim strProgram$
    
    If m.RT.Active <> bActive Then
        m.nServerStatus = 0
        m.dLastMsgTime = 0
        If Len(m.strServer) = 0 Then
            ' must first init the settings
            Init
            If Len(m.strServer) = 0 Or m.RT.Active = bActive Then
                Exit Property
            End If
        End If
        With m.RT
            ' set AutoSendInterval (for sending ticks)
            If bActive Then
                m.strSplitterMsg = ""
                .ClearInbox
                .Active = True
                
                ' make sure RT server is not already running
                DoEvents
                If Not FileExist(App.Path & "\debug.rt") Then
                    If IsServerActive Then
                        ' try to shut it down
                        .CreateMessage m.strServerNode, 21, "0"
                        .CreateMessage m.strServerNode, 99, m.strDisconnectReason
                        .Active = False
                        Sleep 3
                        .Active = True
                        DoEvents
                        If IsServerActive Then
                            .Active = False
                            InfBox "Please first shut down the GenesisRT program,|then try starting the real-time feed.", "e", , "Realtime Feed"
                            Exit Property
                        End If
                    End If
                    ' TLB 5/15/2006: in case a thread in GenesisRT had gotten hung
                    If KillProcess("GenesisRT") > 0 Then
                        .Active = False
                        Sleep 3
                        .Active = True
                    End If
                End If
                
                m.strDisconnectReason = ""
                
                ' start the RT server program
                strProgram = App.Path & "\..\RealTime\" & m.strServer _
                    & "\" & m.strServer & "RT"
                If SalmonIsRunning Then
                    If FileExist(strProgram & ".WS") Then
                        If Abs(FileDate(strProgram & ".WS") - FileDate(strProgram & ".EXE")) > 1 / 1440# Then
                            FileCopy strProgram & ".WS", strProgram & ".EXE", True
                        End If
                    End If
                Else
                    SalmonStop
                    If FileExist(strProgram & ".NS") Then
                        If Abs(FileDate(strProgram & ".NS") - FileDate(strProgram & ".EXE")) > 1 / 1440# Then
                            FileCopy strProgram & ".NS", strProgram & ".EXE", True
                        End If
                    End If
                End If
                strProgram = strProgram & ".exe"
                If FileDate(strProgram) < DateSerial(2005, 12, 1) Then
                    RunProcess strProgram, , , vbMinimizedNoFocus
                Else
                    ' high priority doesn't work for streaming replay (e.g. at 8x)
                    RunProcess strProgram, , , vbHide ', , HIGH_PRIORITY_CLASS
                End If
                frmMain.tmrMain.Tag = "REALTIME"
                m.bSocketLogging = False
            Else
                ' stop sending ticks and tell RT client to shut down
                .CreateMessage m.strServerNode, 21, "0"
                .CreateMessage m.strServerNode, 99, m.strDisconnectReason
                DebugLog "STREAMING OFF: " & m.strDisconnectReason
                .Active = False
                m.strIPandKey = ""
                m.dFeedTime = 0
                m.RtSymbolInfo.Clear
                m.BrokerSubscriptionList.Clear
                m.bSocketLogging = False
            End If
        End With
        'm.strDisconnectReason = "" ' TLB 10/22/2012: this line was moved up to the "If bActive" section
            
        CopyRecalcLog
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cRealTime.Active.Let", eGDRaiseError_Raise
    
End Property

' Event called whenever new message comes from RealTime server
Public Sub RTMessage(msg As gdOCX.gdAppMailMsg)
On Error GoTo ErrSection:

    Dim i&, iPos&, strSymbol$, iTick&, nFld&, iRow&, lCol&, hFile&
    Dim dPrice#, dPrevClose#, dTemp#
    Dim strTemp$, strKey$, strFile$
    Dim aFlds As New cGdArray
    Dim TickBars As cGdBars
    
#If 0 Then
    ' for debugging: display count of each type of message per second
    Static iMsgCount&, iTradeCount&, iBidAskCount&, iDomCount&, dPrevTime#
    If IsIDE Then
        If gdTickCount > dPrevTime + 1000 Then
            dPrevTime = gdTickCount
            If frmTest.Visible Then frmTest.AddList "#msgs = " & Str(iMsgCount) _
                & " = " & Str(iTradeCount) & " t, " & Str(iBidAskCount) & " ba, " & Str(iDomCount) & " dom, " _
                & Str(iMsgCount - iTradeCount - iBidAskCount - iDomCount)
            iMsgCount = 0
            iTradeCount = 0
            iBidAskCount = 0
            iDomCount = 0
        End If
        iMsgCount = iMsgCount + 1
        Select Case msg.MsgType
        Case 20, 201
            iTradeCount = iTradeCount + 1
        Case 25, 26, 202
            iBidAskCount = iBidAskCount + 1
        Case 52, 203
            iDomCount = iDomCount + 1
        End Select
    End If
#End If

    ' ignore messages from other servers
    ' (except for message to sync the symbol list)
    If UCase(msg.FromControlName) <> UCase(m.strServerNode) Then
        ''If msg.MsgType <> 11 Then Exit Sub
        Exit Sub
    End If
    
    Select Case msg.MsgType
    Case 2 ' Status message comes once a minute (right "on" the minute) or when the status changes
        ' Status <tab> FeedTime <tab> RT latency <tab> gdTickCount
        If m.bDebug Then frmTest2.AddList msg.Message
        aFlds.SplitFields msg.Message, vbTab
        m.dFeedTime = RoundToSecond(Val(aFlds(1)))
        If g.nReplaySession > 0 Then
            frmReplay.UpdateTime m.dFeedTime
        End If
        m.dLastFeedTimeMsg = gdTickCount
        If m.nServerStatus <> Val(aFlds(0)) Then
            i = m.nServerStatus
            m.nServerStatus = Val(aFlds(0))
            If (i = 0) And SalmonIsRunning Then
                RequestInitialSymbolBatch
            End If
        End If
        If m.nServerStatus <> 0 Then
            m.dLastMsgTime = gdTickCount ' store tick count in order to know when last received something
        End If
        CheckLatency msg.MsgType, msg.Message
        ' call this once a minute to avoid computer going into hibernation/sleep mode
        DoNotHibernateNow
    
    Case 110, 106, 108
        ' 110 = Each new second of data (not necessarily every second, and not right "on" the second)
        ' 106 = Socket logging (106, gdTickCount, # bytes read, # ms waiting for data, # ms since last read)
        ' 108 = Latency msg from splitter
        If m.dFeedTime > 0 Then
            m.dLastMsgTime = gdTickCount ' store tick count in order to know when last received something
            If g.nReplaySession = 0 Then
                CheckLatency msg.MsgType, msg.Message
            ElseIf msg.MsgType = 110 Then
                ' During streaming replay, need to get feed time from each second message
                ' status msg: Status <tab> FeedTime <tab> RT latency <tab> gdTickCount when sent
                dTemp = Val(Parse(msg.Message, vbTab, 2))
                If dTemp > 0 Then
                    m.dFeedTime = RoundToSecond(dTemp)
                    m.dLastFeedTimeMsg = gdTickCount
                End If
            End If
        End If
    
    Case 11 ' Sync list of symbols (sent when RT starts up AND when we go active)
        If m.bDebug Then frmTest2.AddList msg.Message
        ''m.strSymbolDelays = "" '(clear symbol delays whenever resynching)
        'aFlds.SplitFields msg.Message, vbTab
        'SyncSymbolLists aFlds, msg.FromControlName
    
    Case 13, 14 ' List of symbols with delay times
        If m.bDebug Then frmTest2.AddList msg.Message
        If msg.MsgType = 13 Then
            m.RtSymbolInfo.Clear
        End If
        'format of string: Message.Format("C\t%s|%d|%s|%d\t",(LPCTSTR)GenSymbol,GRTSymbol.Delay,FeedSymbol,FlagsByte);
        aFlds.SplitFields msg.Message, vbTab
        For i = 0 To aFlds.Size - 1
            strTemp = Trim(UCase(aFlds(i)))
            If InStr(strTemp, "|") > 0 Then
                DebugLog strTemp
                ''frmTest.AddList msg.Message & vbTab & Format(Now, "hh:mm:ss") & vbTab & Str(msg.MsgType)
                strSymbol = Parse(strTemp, "|", 1)
                If Not m.RtSymbolInfo.BinarySearch(strSymbol & "|", iPos, eGdSort_MatchUsingSearchStringLength Or eGdSort_IgnoreCase) Then
                    m.RtSymbolInfo.Add strTemp, iPos '(insert for new symbol)
                ElseIf m.RtSymbolInfo(iPos) <> strTemp Then
                    ' replace info for this symbol
                    m.RtSymbolInfo(iPos) = strTemp '(replace for existing symbol)
                    ' set flag to reload data (since delay or something has changed)
                    SetTickBufferStatus strSymbol
                    ClearCache strSymbol
                End If
                ' feed symbol
                strTemp = Parse(strTemp, "|", 3)
                If Len(strTemp) > 0 Then
                    frmQuotes.FeedSymbol strSymbol, strTemp
                    'frmTest.AddList "GenSymbol = " & strSymbol & vbTab & " FeedSymbol = " & strTemp
                End If
            End If
        Next

    ' new tick, bid, ask, market-depth
    Case 20, 25, 26, 52
        If Not SalmonIsRunning Then
            StreamedMessage msg.MsgType, msg.Message
        End If
    Case 201, 202
        If Not SalmonIsRunning Then
            StreamedQuote msg.MsgType, msg.MsgPtr, msg.MsgLength
        End If
    Case 203
        'If Not SalmonIsRunning Then
            StreamedQuote msg.MsgType, msg.MsgPtr, msg.MsgLength
        'End If

    Case 22 ' Settle (GenSymbol, Date.Time, SettlePrice)
        m.dLastMsgTime = gdTickCount ' store tick count in order to know when last received something
        aFlds.SplitFields msg.Message, vbTab
        strSymbol = UCase(Trim(aFlds(0)))
        DebugLog "SETTLE: " & msg.Message
        SetTickBufferStatus strSymbol
        m.dQuoteCount = m.dQuoteCount + 1
        ClearCache strSymbol
    
    Case 31 ' Ticks for symbol
        i = msg.MsgLength
        If i > 50 Then i = 50
        strTemp = Space(i)
        CopyMemory ByVal strTemp, ByVal msg.MsgPtr, i
        iPos = InStr(strTemp, vbTab)
        If iPos > 0 Then
            strTemp = Left(strTemp, iPos - 1)
            iPos = msg.MsgPtr + iPos '(start of DateTime array)
            aFlds.SplitFields strTemp, ","
            strKey = Str(msg.MsgRepliedFrom)
            Set TickBars = m.TickRetrieval(strKey)
            If Not TickBars Is Nothing Then
                TickBars.Prop(eBARS_Symbol) = aFlds(0)
                TickBars.ArrayMask = eBARS_TickByTick
                TickBars.Prop(eBARS_PeriodicityStr) = "Each Tick"
                TickBars.Size = Val(aFlds(1))
                If TickBars.Size > 0 Then
                    i = gdGetDataPtr(TickBars.ArrayHandle(eBARS_DateTime))
                    If i <> 0 Then CopyMemory ByVal i, ByVal iPos, TickBars.Size * 8
                    iPos = iPos + TickBars.Size * 8
                    i = gdGetDataPtr(TickBars.ArrayHandle(eBARS_Close))
                    If i <> 0 Then CopyMemory ByVal i, ByVal iPos, TickBars.Size * 8
                    iPos = iPos + TickBars.Size * 8
                    i = gdGetDataPtr(TickBars.ArrayHandle(eBARS_Vol))
                    If i <> 0 Then CopyMemory ByVal i, ByVal iPos, TickBars.Size * 8
                    iPos = iPos + TickBars.Size * 8
                    If msg.MsgLength > TickBars.Size * 32 Then
                        i = gdGetDataPtr(TickBars.ArrayHandle(eBARS_Flags))
                        If i <> 0 Then CopyMemory ByVal i, ByVal iPos, TickBars.Size * 8
                    End If
                    TickBars.Prop(eBARS_LastTickTime) = MinutesFromMidnight(TickBars(eBARS_DateTime, TickBars.Size - 1))
                    FixTicks TickBars
                Else
                    TickBars.Prop(eBARS_LastTickTime) = -1
                End If
            End If
        End If
        Set TickBars = Nothing
        
    Case 60 ' Real-time client requires an upgrade
        m.dLastMsgTime = 0 '(so won't trigger a reconnect)
        strTemp = InfBox("You need to upgrade the program in order to connect to the real-time data stream.||Would you like to Upgrade now?", "?", "+Upgrade|-Not now", "Upgrade Required")
        g.RealTime.Init False, "Upgrade required"
        If UCase(strTemp) = "U" Then frmMain.tmrMain.Tag = "UPGRADE"
        
    Case 62 ' Connection is dropping, but store message to display reason
        m.strSplitterMsg = Parse(msg.Message, vbTab, 1)
        If m.dLastMsgTime = 0 Then
            ' in this special case, store an old time (couple minutes ago) just so will trigger an auto-reconnect
            m.dLastMsgTime = gdTickCount - 120000
        End If
        
    Case 66 ' Another machine with same data service has just connected
        m.dLastMsgTime = 0 '(so won't trigger a reconnect)
        m.strSplitterMsg = Parse(msg.Message, vbTab, 1)
        DebugLog "Streaming dropped (66): " & msg.Message
        If Len(m.strSplitterMsg) = 0 Then m.strSplitterMsg = "Your realtime streaming connection has been dropped."
        ' TLB 3/15/2013: per Glen, make this message non-modal (so will continue to do daily downloads)
        ShowAlertPopup m.strSplitterMsg, "Realtime Streaming", vbCenter
        'InfBox m.strSplitterMsg, "e", , "Realtime Streaming"
        'InfBox "Your realtime connection has been dropped. |A different computer has just connected using the same account and data service.", "e", , "Realtime Streaming"
    
    Case 70 ' File downloaded through stream
        '(reserved for later)
    
    Case 72 ' Tick corrections exist (tab-delimited list of Genesis symbols)
        aFlds.SplitFields msg.Message, vbTab
        For i = 0 To aFlds.Size - 1
            strSymbol = aFlds(i)
            
        Next
    
    Case 73 ' Reload RT data for symbol(s)
        m.dLastMsgTime = gdTickCount ' store tick count in order to know when last received something
        aFlds.SplitFields msg.Message, vbTab
        For i = 0 To aFlds.Size - 1
            strSymbol = UCase(Trim(aFlds(i)))
            SetTickBufferStatus strSymbol
            m.dQuoteCount = m.dQuoteCount + 1
            ClearCache strSymbol
        Next
    
    Case 74 ' TradeServer fill/expiration (need to unzip this file -- .TRD is in the zip file)
        'strTemp = msg.Message
        'strFile = App.Path & "\SimTrade\Trades.GZP"
        'KillFile strFile
        'hFile = FileOpen(strFile, "wb")
        'If hFile Then
        '    FileBinaryIO hFile, ByVal strTemp, Len(strTemp), True
        '    FileClose hFile
        '    If FileLength(strFile) > 10 Then
        '        ZipExecute "U", strFile, AddSlash(App.Path) & "SimTrade\In", "*.TRD"
        '        frmTTSummary.tmrSimFill.interval = 250
        '        frmTTSummary.tmrSimFill.Enabled = True
        '    End If
        'End If
        
    Case 80 ' Splitter buffer latency
        'strTemp = Parse(msg.Message, vbTab, 1)
        'If Len(strTemp) > 0 And InStr(strTemp, Chr(0)) = 0 Then
        '    dSplitterLatency = Val(Parse(strTemp, vbTab, 1))
        'End If
        
    Case 151 ' OptNav streaming list
        m.OptNavStreamingList.Size = 0
        strTemp = Trim(Replace(msg.Message, vbCrLf, vbTab))
        If Len(strTemp) > 0 Then
            m.OptNavStreamingList.SplitFields strTemp, vbTab
            m.OptNavStreamingList.Sort eGdSort_DeleteNullValues
        End If
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.RTMessage", eGDRaiseError_Raise
    
End Sub

' To add a "bars" to the buffer (for updating upon request)
Public Sub AddTickBuffer(ByVal Bars As cGdBars, Optional ByVal bAddToRecentSymbolList As Boolean = True)
On Error GoTo ErrSection:

    Dim idx&, strSymbol$, strKey$, s$, iPos&
    Dim TickBuffer As cGdBars, TickBuffer2 As cGdBars
    Dim bAddedGeneral As Boolean
    
    strKey = Str(Bars.BarsHandle)
    
    ' first make sure existing one is removed
    RemoveTickBuffer Bars, True
    
    ' don't add for an external symbol or mutual fund
    strSymbol = ""
    If Not Bars.IsExternalSymbol Then
        strSymbol = UCase(Trim(Bars.Prop(eBARS_Symbol)))
        Select Case SecurityType(strSymbol)
        Case "M":
            strSymbol = ""
        Case "F":
            strSymbol = RollSymbolForDate(strSymbol)
        End Select
    End If
    If Len(strSymbol) = 0 Then Exit Sub
    
    ' set up new tick bars (use copy of Bars so all properties will get copied)
    Set TickBuffer = Bars.MakeCopy(True)
    TickBuffer.Size = 0
    TickBuffer.ArrayMask = eBARS_TickByTick Or eBARS_OI
    TickBuffer.Prop(eBARS_PeriodicityStr) = "Each Tick"
    
    ' add to buffers
    If m.TickBufferSymbols.Size = 0 Then m.TickBufferSymbols.Add ""
    m.TickBufferSymbols.BinarySearch strSymbol, idx
    m.TickBufferSymbols.Add strSymbol, idx
    m.TickBuffers.Add TickBuffer, strKey, idx, eTREE_Myself
    ' TLB 1/19/2010: due to timing issues, need to always create the general tickbuffer for intraday data
    ' (even if running salmon -- it's just that salmon will ignore them)
    If Bars.IsIntraday Then
        ' TLB 5/18/2010: for heikin-ashi intraday bars when running salmon,
        ' we need to set the OK flag now since SpliceBars will get bypassed
        If TickBuffer.Prop(eBARS_CustomString) = "HEIKIN-ASHI" Then
            TickBuffer.Prop(eBARS_CustomString) = "<OK>"
        End If

        ' TLB 6/21/2008: create a general tickbuffer for SpliceBars to use for intraday bars
        ' (much more efficient since will help avoid calls to GenesisRT for every new bar)
        strKey = "^" & UCase(Trim(strSymbol))
        If Not m.TickBuffers.Exists(strKey) Then
            ClearCache strSymbol
            ' make copy of bars for generic tick buffer, but must clear any custom start/stop time
            Set TickBuffer2 = TickBuffer.MakeCopy(True)
            TickBuffer2.Prop(eBARS_StartTime) = TickBuffer2.Prop(eBARS_DefaultStartTime)
            TickBuffer2.Prop(eBARS_EndTime) = TickBuffer2.Prop(eBARS_DefaultEndTime)
            ' keep the generic tick buffer as the last in the sortation (for RemoveObsolete to work correctly)
            Do While m.TickBufferSymbols(idx) = strSymbol
                idx = idx + 1
            Loop
            m.TickBufferSymbols.Add strSymbol, idx
            m.TickBuffers.Add TickBuffer2, strKey, idx, eTREE_Myself
            Set TickBuffer2 = Nothing
            bAddedGeneral = True
        End If
    End If
    Set TickBuffer = Nothing
    m.bDirtySymbolList = True
    
    If bAddToRecentSymbolList Then
        s = strSymbol & vbTab & Format(CDbl(Now), "#0.####")
        If m.RecentSymbols.BinarySearch(strSymbol & vbTab, iPos, eGdSort_MatchUsingSearchStringLength) Then
            m.RecentSymbols(iPos) = s
        Else
            m.RecentSymbols.Add s, iPos
        End If
    End If

    If frmTest.Visible Then
        s = Str(m.TickBuffers.Count) & " tick buffers"
        If bAddedGeneral Then
            s = s & ", added general for " & strSymbol
        End If
        frmTest.AddList s
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.AddTickBuffer", eGDRaiseError_Raise
    
End Sub

' To remove a "bars" from the buffer
Public Sub RemoveTickBuffer(ByVal Bars As cGdBars, Optional ByVal bFromAddTickBuffer As Boolean = False)
On Error GoTo ErrSection:

    Dim idx&, iPos&, strKey$, strSymbol$, hBars&
    Dim TickBuffer As cGdBars
    
    If Bars Is Nothing Then Exit Sub
    strKey = Str(Bars.BarsHandle)
    idx = m.TickBuffers.Index(strKey)
    If idx > 0 Then
        m.TickBuffers.Remove idx
        m.TickBufferSymbols.Remove idx
        m.bDirtySymbolList = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.RemoveTickBuffer", eGDRaiseError_Raise
    
End Sub

' To remove any obolete tick buffers (e.g. the "general" one used by SpliceBars for efficiency)
Public Sub RemoveObsoleteTickBuffers()
On Error GoTo ErrSection:

    Dim i&, idx&, strSymbol$, hBars&
       
    ' If no intraday bars are being used for a symbol, then we can remove the generic tick buffer for that symbol.
    ' Since the generic tick buffer will always be the last item for that symbol in the sorted list ...
    ' - walk backwards through the sorted list to look symbol-by-symbol
    ' - when hit a generic tick buffer, then start looking to see if any intraday bars are currently using it for that symbol
    ' - if at least one is, then just look for the next generic tick buffer
    ' - else if no intraday bars are using it, then remove that generic tick buffer
    ' (Note: the 0th item does not exist in the collection -- it does in the array of symbols, but is unused)
    idx = 0
    For i = m.TickBufferSymbols.Size - 1 To 0 Step -1
        If m.TickBufferSymbols(i) <> strSymbol Then
            ' remove obsolete tick buffer if still "flagged" as unused
            If idx > 0 Then
                m.TickBufferSymbols.Remove idx
                m.TickBuffers.Remove idx
                If IsIDE And Not g.bUnloading Then
                    If frmTest.Visible Then
                        frmTest.AddList Str(m.TickBuffers.Count) & " tick buffers, removed general for " & strSymbol
                    End If
                End If
            End If
            ' new symbol
            strSymbol = m.TickBufferSymbols(i)
            idx = 0
        End If
        If i > 0 Then
            If Left(m.TickBuffers.Key(i), 1) = "^" Then
                idx = i
            ElseIf idx > 0 Then
                ' see if the bars that is using this tick buffer is intraday
                hBars = Val(m.TickBuffers.Key(i))
                If hBars <> 0 Then
                    If gdGetBarsNumProp(hBars, eBARS_PeriodType) < ePRD_Days Then
                        ' Yes, there is still at least one still using it
                        idx = 0 '(so we won't remove it)
                    End If
                End If
            End If
        End If
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.RemoveObsoleteTickBuffers", eGDRaiseError_Raise
End Sub

' To update the "bars" with recent ticks (must have been added to the buffer)
' - returns True if something has changed
' - sets bNewBar (if passed) to true if data exists for a new bar (so you can reload the data)
Public Function UpdateBars(Bars As cGdBars, Optional bNewBar As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim iTick&, iBar&, nPerBar&, strStatus$, strSymbol$, nPeriodType As eBarsPeriodType
    Dim dPrice#, dTime#, dCrossoverTime#, dEndTime#, dPrev#, dDiff#, nSessionDate&
    Dim bUpdated As Boolean
    Dim TickBuffer As cGdBars
    Dim SymInf As cSymbolInfo
    
    If Bars Is Nothing Then Exit Function
    If bDisableForTesting Then Exit Function
    
    ' see if there are new ticks for these bars
    nPeriodType = Bars.Prop(eBARS_PeriodType)
    nPerBar = Bars.Prop(eBARS_PeriodsPerBar)
    Set TickBuffer = m.TickBuffers(Str(Bars.BarsHandle))
    If m.RT.Active And Not TickBuffer Is Nothing Then
        strStatus = TickBuffer.Prop(eBARS_CustomString)
        strSymbol = Bars.Prop(eBARS_Symbol)
        
        If SalmonIsRunning Then
            Set SymInf = SymbolInfo(strSymbol)
            If nPeriodType >= ePRD_Days Then
                ' for Daily bars (non-intraday), just call UpdateDailyBars
                bUpdated = SymInf.UpdateDailyBars(Bars, bNewBar, strStatus)
                If bUpdated Then
                    TickBuffer.Prop(eBARS_CustomString) = strStatus
                End If
                TickBuffer.Size = 0
            Else
                ' for Intraday, see if the data is available
                If SymInf.GetDataRequestStatus(nPeriodType) <> eSalmonAvailable Then
                    ' and just continue to wait for the data
                    TickBuffer.Prop(eBARS_CustomString) = "<PENDING>"
                    TickBuffer.Size = 0
                ElseIf strStatus <> "<OK>" Then
                    ' if the data has become available since the last SpliceBars call, then force a reload
                    bNewBar = True
                    bUpdated = True
                    TickBuffer.Prop(eBARS_CustomString) = "<RELOAD>"
                    TickBuffer.Size = 0
                    strStatus = ""
                Else
                    ' get the newly streamed ticks for this symbol (since last called)
                    SalmonGetNewTicks strSymbol
                End If
            End If
            Set SymInf = Nothing
        End If
        
        If strStatus = "<RELOAD>" Then
            ' if QB refresh just finished for this symbol, then flag to reload
            bNewBar = True
            bUpdated = True
            TickBuffer.Size = 0
            TickBuffer.Prop(eBARS_CustomString) = ""
        ElseIf TickBuffer.Size > 0 Then
            If Bars.Size = 0 Then
                ' make sure valid bars type (e.g. if need full ticks)
                If Len(BarPeriodError(strSymbol, nPeriodType)) = 0 Then
                    ' new ticks but no bars, so reload
                    bNewBar = True
                    bUpdated = True
                End If
            Else
                Bars.Prop(eBARS_Exchange) = TickBuffer.Prop(eBARS_Exchange)
                
                ' get crossover and end time as fraction of day
                dCrossoverTime = Bars.Prop(eBARS_CrossoverTime) / 1440#
                dEndTime = Bars.Prop(eBARS_EndTime) / 1440#
                
                ' find starting tick (ignore if any are less than LastTickTime,
                ' e.g. if quote-board refresh has more recent data than delayed feed)
                iTick = 0
                If Bars.Prop(eBARS_LastTickTime) > 0 Then
                    Do While iTick < TickBuffer.Size
                        dTime = RoundToSecond(TickBuffer(eBARS_DateTime, iTick))
                        dDiff = MinutesFromMidnight(dTime) - Bars.Prop(eBARS_LastTickTime) + 0.00000001
                        If dDiff >= 0 Or dDiff < -60 Then
                            Exit Do
                        End If
                        iTick = iTick + 1
                    Loop
                Else
                    dTime = RoundToSecond(TickBuffer(eBARS_DateTime, iTick))
                End If
                
                ' find starting bar
                For iBar = Bars.Size - 1 To 0 Step -1
                    If dTime >= RoundToSecond(Bars(eBARS_DateTime, iBar - 1)) Then
                        If Bars(eBARS_Close, iBar) <> kNullData Then
                            Exit For
                        End If
                    End If
                Next
                
                ' update bars with new ticks
                Do While iTick < TickBuffer.Size '(use "Do" instead of "For" since .Size may be changing)
                    ' check for valid session date (e.g. if within custom start/stop time)
                    dTime = RoundToSecond(TickBuffer(eBARS_DateTime, iTick))
                    nSessionDate = Bars.SessionDateForTime(dTime, True)
                    If nSessionDate > 0 Then
                        ' if tick is past crossover time from last bar/tick, then just reload
                        If nSessionDate > Bars.SessionDate(iBar) Then
                            bUpdated = True
                            bNewBar = True
                            Exit Do
                        End If
                        
                        If nPeriodType = ePRD_EachTick Then '(adding to ticks)
                            ' add new tick
                            iBar = iBar + 1
                            Bars.Size = iBar + 1
                            Bars.Prop(eBARS_LastTickTime) = MinutesFromMidnight(dTime)
                            Bars(eBARS_DateTime, iBar) = dTime
                            Bars(eBARS_Close, iBar) = TickBuffer(eBARS_Close, iTick)
                            If (Bars.ArrayMask And eBARS_Vol) <> 0 And (TickBuffer.ArrayMask And eBARS_Vol) <> 0 Then
                                Bars(eBARS_Vol, iBar) = TickBuffer(eBARS_Vol, iTick)
                            End If
                            If (Bars.ArrayMask And eBARS_Flags) <> 0 And (TickBuffer.ArrayMask And eBARS_Flags) <> 0 Then
                                Bars(eBARS_Flags, iBar) = TickBuffer(eBARS_Flags, iTick)
                            End If
                        Else
                            ' for intraday bars, see if in current bar
                            ' (i.e. Exit this Do Loop if this is the right bar for this tick)
                            Do While iBar < Bars.Size And Bars.IsIntraday
                                Select Case nPeriodType
                                Case ePRD_IntVol
                                    ' intraday volume per bar
                                    If Bars(eBARS_Vol, iBar) + TickBuffer(eBARS_Vol, iTick) <= nPerBar Then
                                        ' update time of bar
                                        Bars(eBARS_DateTime, iBar) = dTime
                                        Exit Do
                                    ElseIf Bars(eBARS_Close, iBar + 1) = kNullData Then
                                        bNewBar = True
                                        Exit Do
                                    End If
                                
                                Case ePRD_Ticks
                                    ' ticks per bar
                                    If Bars(eBARS_UpTicks, iBar) + Bars(eBARS_DownTicks, iBar) + 1 <= nPerBar Then
                                        ' update time of bar
                                        Bars(eBARS_DateTime, iBar) = dTime
                                        Exit Do
                                    ElseIf Bars(eBARS_Close, iBar + 1) = kNullData Then
                                        bNewBar = True
                                        Exit Do
                                    End If
                                
                                Case ePRD_IntBreakout
                                    dDiff = 0
                                    dPrice = TickBuffer(eBARS_Close, iTick)
                                    If dPrice > Bars(eBARS_High, iBar) Then
                                        dDiff = dPrice - Bars(eBARS_Low, iBar)
                                    ElseIf dPrice < Bars(eBARS_Low, iBar) Then
                                        dDiff = Bars(eBARS_High, iBar) - dPrice
                                    End If
                                    If nPerBar * Bars.Prop(eBARS_TickMove) - dDiff > -0.00000001 Then
                                        ' update time of bar
                                        Bars(eBARS_DateTime, iBar) = dTime
                                        Exit Do
                                    ElseIf Bars(eBARS_Close, iBar + 1) = kNullData Then
                                        bNewBar = True
                                        Exit Do
                                    End If
                                
                                Case ePRD_IntPF, ePRD_IntKagi, ePRD_IntRenko
                                    ' make sure we're on the last data bar
                                    If Bars(eBARS_Close, iBar + 1) = kNullData Then
                                        ' all we can really do is rebuild bars whenever a new minute starts
                                        dPrev = RoundToSecond(Bars(eBARS_DateTime, iBar))
                                        If Int(dTime * 1440# + 0.001) > Int(dPrev * 1440# + 0.001) Then
                                            bNewBar = True
                                        ElseIf dTime <> dPrev Then
                                            ' update time of bar
                                            Bars(eBARS_DateTime, iBar) = dTime
                                        End If
                                        Exit Do
                                    End If
                                
                                Case Else ' minutes per bar
                                    If dTime < RoundToSecond(Bars(eBARS_DateTime, iBar)) Then
                                        Exit Do
                                    ElseIf dEndTime > 0 And dCrossoverTime > 0 Then
                                        ' or see if tick is between session end and crossover time
                                        If dTime - Int(dTime) >= dEndTime Then
                                            If dTime - Int(dTime) < dCrossoverTime Then
                                                ' and current bar is the session end
                                                If Bars(eBARS_DateTime, iBar) >= Int(Bars(eBARS_DateTime, iBar)) + dEndTime - 0.0001 Then
                                                    Exit Do
                                                End If
                                            End If
                                        End If
                                    End If
                                End Select
                                iBar = iBar + 1
                            Loop
                            
                            If (iBar >= Bars.Size) Or bNewBar Then
                                bUpdated = True
                                bNewBar = True
                                Exit Do
                            End If
                        
                            ' update the bar
                            dPrev = Bars(eBARS_Close, iBar)
                            If dPrev = kNullData Then
                                dPrev = Bars(eBARS_Close, iBar - 1)
                            End If
                            dPrice = TickBuffer(eBARS_Close, iTick)
                            Bars.Prop(eBARS_LastTickTime) = MinutesFromMidnight(dTime)
                            If Bars(eBARS_Open, iBar) = kNullData Then
                                bNewBar = True
                                Bars(eBARS_Open, iBar) = dPrice
                            End If
                            If nPeriodType <> ePRD_IntRenko Then ' (CANNOT update Highs/Lows for Renko bars)
                                If Bars(eBARS_High, iBar) < dPrice Or Bars(eBARS_High, iBar) = kNullData Then
                                    Bars(eBARS_High, iBar) = dPrice
                                End If
                                If Bars(eBARS_Low, iBar) > dPrice Or Bars(eBARS_Low, iBar) = kNullData Then
                                    Bars(eBARS_Low, iBar) = dPrice
                                End If
                            End If
                            Bars(eBARS_Close, iBar) = dPrice
                            If Bars(eBARS_Vol, iBar) = kNullData Then
                                Bars(eBARS_Vol, iBar) = TickBuffer(eBARS_Vol, iTick)
                            Else
                                Bars(eBARS_Vol, iBar) = Bars(eBARS_Vol, iBar) + TickBuffer(eBARS_Vol, iTick)
                            End If
                            
                            ' up/down ticks
                            If dPrice > dPrev Then
                                Bars.Prop(eBARS_LastTickDown) = 0
                            ElseIf dPrice < dPrev Then
                                Bars.Prop(eBARS_LastTickDown) = 1
                            End If
                            If Bars.IsIntraday Then
                                If Bars.Prop(eBARS_LastTickDown) = 0 Then
                                    Bars(eBARS_UpTicks, iBar) = Bars(eBARS_UpTicks, iBar) + 1
                                Else
                                    Bars(eBARS_DownTicks, iBar) = Bars(eBARS_DownTicks, iBar) + 1
                                End If
                            End If
                            
                            ' BidVol and AskVol
                            If TickBuffer(eBARS_Vol, iTick) > 0 And (Bars.ArrayMask And eBARS_BidVol) > 0 Then
                                Select Case TickBuffer(eBARS_Flags, iTick)
                                Case eTICK_AtBid
                                    If Bars(eBARS_BidVol, iBar) > 0 Then
                                        Bars(eBARS_BidVol, iBar) = Bars(eBARS_BidVol, iBar) + TickBuffer(eBARS_Vol, iTick)
                                    Else
                                        Bars(eBARS_BidVol, iBar) = TickBuffer(eBARS_Vol, iTick)
                                    End If
                                Case eTICK_AtAsk
                                    If Bars(eBARS_AskVol, iBar) > 0 Then
                                        Bars(eBARS_AskVol, iBar) = Bars(eBARS_AskVol, iBar) + TickBuffer(eBARS_Vol, iTick)
                                    Else
                                        Bars(eBARS_AskVol, iBar) = TickBuffer(eBARS_Vol, iTick)
                                    End If
                                End Select
                            End If
                        End If
                        
                        bUpdated = True
                    End If
                    iTick = iTick + 1
                Loop
            End If
            ' TLB 12/23/2009: only want to do this if not flagged for a new bar to start
            ' (so if this is called again before reloading then will still flag for a new bar)
            If Not bNewBar Then
                TickBuffer.Size = 0
            End If
        End If
        
        If Not bNewBar And Not SalmonIsRunning Then
            bUpdated = bUpdated Or UpdateBidAsk(Bars)
        End If
    End If
    Set TickBuffer = Nothing

    UpdateBars = bUpdated

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.UpdateBars", eGDRaiseError_Raise
    
End Function

' TLB 10/17/2008: Design for the Tick Cache (to avoid continually merging ticks from DMDLL and GenesisRT)
'   If the generic tickbuffer and cache file both exist then
'       Read the ticks from the cache file
'       Append the ticks from the generic tick buffer
'   Else
'       Merge the ticks from the DMDLL and GenesisRT
'       Clear the generic tick buffer
'       Store the ticks into the cache file
'   End If

Public Function SpliceBars(Bars As cGdBars, Optional ByVal lStartSession = 0, _
                Optional ByVal bIgnoreSettle As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim i&, d#
    Dim DataMgrTicks As cGdBars     ' Ticks from data manager from LastDailyDownload to end
    Dim RealTimeTicks As cGdBars    ' Ticks from the RealTime Server from DM Ticks
    Dim TickBuffer As cGdBars       ' Ticks from the tick buffer
    Dim DailyBars As New cGdBars
    Dim lMsgNum As Long                 ' Message number from the RT Server
    Dim dStartTime As Double            ' Starting date/time of ticks from RT Server
    Dim lIndex As Long                  ' Index for a for loop
    Dim lStart As Long                  ' Starting index
    Dim lOldSize As Long                ' Old size of bars
    Dim dCrossOver As Double            ' Crossover time in minutes (fractional)
    Dim dTime As Double
    Dim lMemFrom As Long, lMemTo As Long
    Dim strSymbol$, strTemp$
    Dim bLog As Boolean
    Dim lEndDay As Long
    Dim bIncludeSettles As Boolean, bTimedOut As Boolean
    Dim iRT&, iDM&, bOverlap As Boolean
    Dim strProfile$, dProfile#, dStartingProfile#
    Dim lStartSave&, lEndSave&, lStartDay&
    
    Dim hDmClose As Long, hDmDateTime, hDmVol As Long
    Dim hRtClose As Long, hRtDateTime, hRtVol As Long

    ' Use new routine if Salmon is running ...
    If SalmonIsRunning Then
        SpliceBars = SpliceBarsSalmon(Bars, lStartSession, bIgnoreSettle)
        Exit Function
    End If
    
    ' don't do anything if real-time is not active
    If Not Active Or Bars Is Nothing Then Exit Function

    ' and don't do anything if the symbol has not yet made it into the real-time symbol list
    ' (e.g. if walking up/down quote board then keep data from current session update)
    strSymbol = RollSymbolForDate(Bars.Prop(eBARS_Symbol))
    If Not m.UniqueSymbols.BinarySearch(strSymbol) Then
        Exit Function
    End If
    
    ' Intraday bar periods other than minute bars now require the full tick database
    If Len(BarPeriodError(strSymbol, Bars.Prop(eBARS_Periodicity))) > 0 Then
        Bars.Size = 0
        Exit Function
    End If
    
    ' Make sure that nobody else can steal the ticks in the tick buffer
    ' until we are done
    If m.TickBuffers.Exists(Str(Bars.BarsHandle)) Then
        m.TickBuffers.Key(Str(Bars.BarsHandle)) = Str(Bars.BarsHandle) & "X"
    ElseIf m.TickBuffers.Exists(Str(Bars.BarsHandle) & "X") Then
        Exit Function '(already in progress for this gdBars!)
    End If
    
    bLog = frmTest.Visible
    dProfile = gdTickCount(False)
    dStartingProfile = dProfile
    
    ' Get the day after the last date of "good" data
    lStartDay = CLng(LastDailyDownload) + 1
    If g.nReplaySession > 0 Then lEndDay = lStartDay
    If bLog Then
        frmTest.AddList "--> " & strSymbol & ", " & Bars.Prop(eBARS_PeriodicityStr) _
            & ",  TickRetrieval = " & Str(m.TickRetrieval.Count)
    End If

    ' save custom start/end time to restore later (but need to clear it to use tickbuffer cache)
    lStartSave = Bars.Prop(eBARS_StartTime)
    lEndSave = Bars.Prop(eBARS_EndTime)
    Bars.Prop(eBARS_StartTime) = Bars.Prop(eBARS_DefaultStartTime)
    Bars.Prop(eBARS_EndTime) = Bars.Prop(eBARS_DefaultEndTime)

    ' include settles with the ticks only if periodicity >= daily
    If (Bars.Prop(eBARS_PeriodType) >= ePRD_Days) And Not bIgnoreSettle Then
        bIncludeSettles = True
    End If
    Bars.Prop(eBARS_PriceHasSettled) = False

    ' setup the DataMgrTicks bars
    If Bars.Prop(eBARS_PeriodType) = ePRD_EachTick Then
        Set DataMgrTicks = Bars
    Else
        ' make copy so all properties get copied
        Set DataMgrTicks = Bars.MakeCopy(True)
        DataMgrTicks.Size = 0
        DataMgrTicks.ArrayMask = eBARS_TickByTick
    End If
    DataMgrTicks.Size = 0

    ' TLB 10/17/2008: Design for the Tick Cache (to avoid continually merging ticks from DMDLL and GenesisRT)
    '   If the generic tickbuffer and cache file both exist then
    '       Read the ticks from the cache file
    '       Append the ticks from the generic tick buffer
    '   Else
    '       Merge the ticks from the DMDLL and GenesisRT
    '       Clear the generic tick buffer
    '       Store the ticks into the cache file
    '   End If
    Set TickBuffer = m.TickBuffers("^" & strSymbol)
    If Not TickBuffer Is Nothing Then
        strTemp = App.Path & "\Cache\" & strSymbol & ".TBC" ' (TickBufferCache file)
        If FileExist(strTemp) Then
            ' read the ticks from the cache file
            DataMgrTicks.Prop(eBARS_PeriodicityStr) = "Each Tick"
            If Not DataMgrTicks.SerializeDataArrays(strTemp, False) Then
                DataMgrTicks.Size = 0
            Else
                ' append the ticks from the generic tick buffer for this symbol
                If bLog Then
                    d = gdTickCount(False)
                    strProfile = strProfile & " RC=" & Format(d - dProfile, "#0")
                    dProfile = d
                End If
                lStart = 0
                'Do While TickBuffer(eBARS_OI, lStart) <= RealTimeTicks.Size And TickBuffer.Size > 0
                '    lStart = lStart + 1
                'Loop
                If TickBuffer.Size > 0 Then
                    If lStart = 0 Then
                        gdAppendBars DataMgrTicks.BarsHandle, TickBuffer.BarsHandle, False
                    Else
                        lOldSize = DataMgrTicks.Size
                        DataMgrTicks.Size = lOldSize + (TickBuffer.Size - lStart)
                        For lIndex = lStart To TickBuffer.Size - 1
                            If lOldSize + (lIndex - lStart) > DataMgrTicks.Size Then
                                DataMgrTicks.Size = lOldSize + (lIndex - lStart)
                            End If
                            DataMgrTicks(eBARS_DateTime, lOldSize + (lIndex - lStart)) = TickBuffer(eBARS_DateTime, lIndex)
                            DataMgrTicks(eBARS_Close, lOldSize + (lIndex - lStart)) = TickBuffer(eBARS_Close, lIndex)
                            If DataMgrTicks.IsActiveArray(eBARS_Vol) Then
                                DataMgrTicks(eBARS_Vol, lOldSize + (lIndex - lStart)) = TickBuffer(eBARS_Vol, lIndex)
                            End If
                            If DataMgrTicks.IsActiveArray(eBARS_Flags) Then
                                DataMgrTicks(eBARS_Flags, lOldSize + (lIndex - lStart)) = TickBuffer(eBARS_Flags, lIndex)
                            End If
                        Next lIndex
                    End If
                End If
                'If bLog Then frmTest.AddList "Read Serialize: " & Str(lOldSize) & " " & Str(DataMgrTicks.Size)
                                
                Set TickBuffer = m.TickBuffers(Str(Bars.BarsHandle) & "X")
                If Not TickBuffer Is Nothing Then
                    TickBuffer.Size = 0
                End If
                If bLog Then
                    d = gdTickCount(False)
                    strProfile = strProfile & " AC=" & Format(d - dProfile, "#0")
                    dProfile = d
                End If
            End If
        End If
    End If
    Set TickBuffer = Nothing

    ' If didn't get merged ticks from the cache, then need to go get all the ticks and merge them
    If DataMgrTicks.Size = 0 Then
        ' 1. Get ticks from DM since LastDailyDownload
        If bLog Then
            d = gdTickCount(False)
            strProfile = strProfile & " 1a=" & Format(d - dProfile, "#0")
            dProfile = d
        End If
        ' NOTE: should get history from DataMgr using Bars.Prop(eBARS_Symbol) instead of strSymbol
        ' (because we want to get -067 history for -067's and the DataMgr will handle it correctly)
        DM_GetBars DataMgrTicks, Bars.Prop(eBARS_Symbol), "Each Tick", lStartDay, lEndDay
        If bLog Then
            d = gdTickCount(False)
            strProfile = strProfile & ", 1b=" & Format(d - dProfile, "#0")
            dProfile = d
        End If
        If DataMgrTicks.Size > 0 Then
            If m.bDebug Then DataMgrTicks.DumpToFile AddSlash(App.Path) & "DMTicks1.ASC"
            ' TLB 5/13/2005: in order to try to get an overlap in RT and DM data,
            ' back up dStartTime to the minute of the tick prior to the last minute of data
            dStartTime = gdFixDateTime(DataMgrTicks(eBARS_DateTime, DataMgrTicks.Size - 1))
            For lIndex = DataMgrTicks.Size - 2 To 0 Step -1
                If dStartTime <> gdFixDateTime(DataMgrTicks(eBARS_DateTime, lIndex)) Then
                    dStartTime = gdFixDateTime(DataMgrTicks(eBARS_DateTime, lIndex))
                    Exit For
                End If
            Next
        ElseIf Bars.Prop(eBARS_CrossoverTime) = 0 Then
            dStartTime = CDbl(lStartDay)
        Else
            dCrossOver = Bars.Prop(eBARS_CrossoverTime) / 1440#
    '        dCrossOver = (((dCrossOver / 100) * 60) + (dCrossOver Mod 100)) / 1440
            dStartTime = CDbl(lStartDay - 1) + dCrossOver
        End If
        dStartTime = gdFixDateTime(dStartTime)
        
        If bLog Then
            d = gdTickCount(False)
            strProfile = strProfile & ", 1c=" & Format(d - dProfile, "#0")
            dProfile = d
        End If
        
        ' only if the server is active ...
        Set RealTimeTicks = Bars.MakeCopy(True)
        RealTimeTicks.Size = 0
        RealTimeTicks.ArrayMask = eBARS_TickByTick
        RealTimeTicks.Prop(eBARS_PeriodicityStr) = "Each Tick"
        If IsServerActive And m.RT.Active Then
        
            ' 2. Request ticks from RT server since time of last tick gotten from DM
            If g.nReplaySession > 0 Then
                If Bars.SecurityType = "F" Then
                    dStartTime = g.nReplaySession
                Else
                    dStartTime = Int(m.dFeedTime * 1440# + 0.00001) / 1440#
                End If
            End If
            strTemp = strSymbol & vbTab & Str(dStartTime) & vbTab & Str(Abs(bIncludeSettles))
            lMsgNum = m.RT.CreateMessage(m.strServerNode, 30, strTemp)
            If bLog Then
                d = gdTickCount(False)
                strProfile = strProfile & ", 2=" & Format(d - dProfile, "#0")
                dProfile = d
            End If
            If lMsgNum <> 0 Then
                m.nTickHistRequestCount = m.nTickHistRequestCount + 1
                RealTimeTicks.Prop(eBARS_LastTickTime) = -29999
                m.TickRetrieval.Add RealTimeTicks, Str(lMsgNum)
                
                FixTicks DataMgrTicks, False, True
        
                ' 3. Wait for ticks from RT server
                dTime = Now
                'Do While m.TickRetrieval(Str(lMsgNum)).Prop(eBARS_LastTickTime) <= -1440
                Do While RealTimeTicks.Prop(eBARS_LastTickTime) <= -1440
                    If Now > dTime + 5# / 84600# Then
                        DebugLog "Timeout waiting for RT Ticks for " & strSymbol
                        m.nTickHistTimeoutCount = m.nTickHistTimeoutCount + 1
                        bTimedOut = True
                        Exit Do
                    End If
                    Sleep 0
                    'DoEvents
                    'Exit Do
                Loop
                'If bLog Then frmTest.AddList "GenesisRT from " & Format(dStartTime, "YYYYMMDD HHMM") & ", TickRetrieval=" & Str(m.TickRetrieval.Count)
                'Set RealTimeTicks = m.TickRetrieval(Str(lMsgNum))
                m.TickRetrieval.Remove Str(lMsgNum)
                'Set m.TickRetrieval(Str(lMsgNum)) = Nothing
                lStart = RealTimeTicks(eBARS_DateTime, RealTimeTicks.Size - 1)
                If m.bDebug Then RealTimeTicks.DumpToFile AddSlash(App.Path) & "RTTicks1.ASC"
            End If
        End If
        
        If bLog Then
            d = gdTickCount(False)
            strProfile = strProfile & ", 3=" & Format(d - dProfile, "#0")
            dProfile = d
        End If
        
        ' 4. Append ticks from tick buffer (if later) and clear tick buffer
        Set TickBuffer = m.TickBuffers(Str(Bars.BarsHandle) & "X")
        If Not TickBuffer Is Nothing Then
            'm.TickBuffers(Str(Bars.BarsHandle) & "X").Size = 0
            TickBuffer.Size = 0
            If m.bDebug Then TickBuffer.DumpToFile AddSlash(App.Path) & "TickBuf.ASC"
            lStart = 0
            Do While TickBuffer(eBARS_OI, lStart) <= RealTimeTicks.Size And TickBuffer.Size > 0
                lStart = lStart + 1
            Loop
            lOldSize = RealTimeTicks.Size
            RealTimeTicks.Size = RealTimeTicks.Size + (TickBuffer.Size - lStart)
            For lIndex = lStart To TickBuffer.Size - 1
                RealTimeTicks(eBARS_DateTime, lOldSize + (lIndex - lStart)) = TickBuffer(eBARS_DateTime, lIndex)
                RealTimeTicks(eBARS_Close, lOldSize + (lIndex - lStart)) = TickBuffer(eBARS_Close, lIndex)
                If RealTimeTicks.IsActiveArray(eBARS_Vol) Then
                    RealTimeTicks(eBARS_Vol, lOldSize + (lIndex - lStart)) = TickBuffer(eBARS_Vol, lIndex)
                End If
                If RealTimeTicks.IsActiveArray(eBARS_Flags) Then
                    RealTimeTicks(eBARS_Flags, lOldSize + (lIndex - lStart)) = TickBuffer(eBARS_Flags, lIndex)
                End If
            Next lIndex
            If m.bDebug Then RealTimeTicks.DumpToFile AddSlash(App.Path) & "RTTicks2.ASC"
            Set TickBuffer = Nothing
        End If
    
        ' tweak things for simulated real-time streaming
        If g.nReplaySession > 0 Then
            dStartTime = RealTimeTicks(eBARS_DateTime, 0)
            If dStartTime <= 0 Then
                dStartTime = ConvertTimeZone(m.dFeedTime, "NY", DataMgrTicks.Prop(eBARS_ExchangeTimeZoneInf))
            End If
            gdBinarySearch DataMgrTicks.ArrayHandle(eBARS_DateTime), gdFixDateTime(dStartTime), _
                    lStart, eGdSort_Default, 0, DataMgrTicks.Size
            If lStart < DataMgrTicks.Size Then
                DataMgrTicks.Size = lStart
            End If
        End If
    
        If bLog Then
            d = gdTickCount(False)
            strProfile = strProfile & ", 4=" & Format(d - dProfile, "#0")
            dProfile = d
        End If
        
        ' 5. Append ticks from RT to ticks from DM (need to identify overlap)
        If RealTimeTicks.Size > 0 Then
            ' if last realtime tick has vol = -1, then it is a settle
            If RealTimeTicks(eBARS_Vol, RealTimeTicks.Size - 1) = -1 Then
                Bars.Prop(eBARS_PriceHasSettled) = True
                If RealTimeTicks.Size = 1 Then
                    ' when only the settle gets passed then RT client doesn't know a time to use,
                    ' so set the time to session end time of the date passed
                    dTime = Int(RealTimeTicks(eBARS_DateTime, 0)) + Bars.Prop(eBARS_EndTime) / 1440#
                    RealTimeTicks(eBARS_DateTime, 0) = dTime
                End If
            End If
        
If 0 And IsIDE Then
    If strSymbol = "SP-067" And bLog Then
        For lStart = RealTimeTicks.Size - 10 To RealTimeTicks.Size - 1
            frmTest.AddList RealTimeTicks(eBARS_Close, lStart)
        Next
    End If
End If
            
            gdBinarySearch RealTimeTicks.ArrayHandle(eBARS_DateTime), gdFixDateTime(dStartTime), _
                    lStart, eGdSort_Default, 0, RealTimeTicks.Size
            If lStart < RealTimeTicks.Size Then
                ' TLB 5/13/2005: skip any RT ticks that overlap the last minute of DM ticks
                If DataMgrTicks.Size > 0 Then
                    ' first see if the minute of first RT tick is prior to the minute of last DM tick
                    dStartTime = gdFixDateTime(DataMgrTicks(eBARS_DateTime, DataMgrTicks.Size - 1))
If 0 And IsIDE And bLog And Bars.IsIntraday Then
    For iDM = DataMgrTicks.Size - 200 To DataMgrTicks.Size - 1
        If gdFixDateTime(DataMgrTicks(eBARS_DateTime, iDM)) >= dStartTime Then
            frmTest.AddList "DM" & vbTab & Format(DataMgrTicks(eBARS_DateTime, iDM), "hh:mm") & vbTab & _
                DataMgrTicks(eBARS_Close, iDM) & vbTab & DataMgrTicks(eBARS_Vol, iDM)
        End If
    Next
    For iRT = 0 To RealTimeTicks.Size - 1
        If gdFixDateTime(RealTimeTicks(eBARS_DateTime, iRT)) > dStartTime Then Exit For
        frmTest.AddList Str(iRT) & vbTab & Format(RealTimeTicks(eBARS_DateTime, iRT), "hh:mm") & vbTab & _
                RealTimeTicks(eBARS_Close, iRT) & vbTab & RealTimeTicks(eBARS_Vol, iRT)
    Next
End If
                    
                    If gdFixDateTime(RealTimeTicks(eBARS_DateTime, lStart)) < dStartTime Then
                        ' if so then can use easy method: can chop off last partial minute of DM ...
                        For lIndex = DataMgrTicks.Size - 1 To 0 Step -1
                            If gdFixDateTime(DataMgrTicks(eBARS_DateTime, lIndex)) < dStartTime Then
                                DataMgrTicks.Size = lIndex + 1
                                Exit For
                            End If
                        Next
                        ' ... and chop the RT ticks prior to that minute
                        For lIndex = lStart To RealTimeTicks.Size - 1
                            lStart = lIndex
                            If gdFixDateTime(RealTimeTicks(eBARS_DateTime, lIndex)) >= dStartTime Then Exit For
                        Next
                        If IsIDE Then
                            If bLog And Bars.IsIntraday Then
                                frmTest.AddList "OVERLAP CHOPPED at minute boundary " & Format(dStartTime, "hh:mm")
                            End If
                        End If
                    Else
                        ' otherwise, we have to use harder method: try to splice the overlapped minute
                        ' together by skipping any RT ticks that overlap the last minute of DM ticks
                        For lIndex = lStart To RealTimeTicks.Size - 1
                            ' if beyond the start time then we're done looking
                            If gdFixDateTime(RealTimeTicks(eBARS_DateTime, lIndex)) > dStartTime Then Exit For
                            ' check if all RT ticks from here back match the DM ticks (for this minute)
                            bOverlap = False
                            iDM = DataMgrTicks.Size - 1
                            For iRT = lIndex To 0 Step -1
                                If gdFixDateTime(RealTimeTicks(eBARS_DateTime, iRT)) < dStartTime Then Exit For
                                ' if not a match (time, price, vol), then an overlap does not start here
                                If RoundToSigDigits(RealTimeTicks(eBARS_Close, iRT)) <> RoundToSigDigits(DataMgrTicks(eBARS_Close, iDM)) Then
                                    bOverlap = False
                                    Exit For
                                ElseIf gdFixDateTime(RealTimeTicks(eBARS_DateTime, iRT)) <> gdFixDateTime(DataMgrTicks(eBARS_DateTime, iDM)) Then
                                    bOverlap = False
                                    Exit For
                                ElseIf DataMgrTicks(eBARS_Vol, iDM) > 0 And RealTimeTicks(eBARS_Vol, iRT) <> DataMgrTicks(eBARS_Vol, iDM) Then
                                    bOverlap = False
                                    Exit For
                                End If
                                iDM = iDM - 1
                                bOverlap = True
                            Next
                            ' if this overlaps, then move the start up (but keep looking for
                            ' the rest of this minute in case a bigger overlap is found)
                            If bOverlap Then
                                If IsIDE Then
                                    If bLog And Bars.IsIntraday Then
                                        frmTest.AddList "OLD START = " & Str(lStart) & ", NEW START = " & Str(lIndex + 1)
                                    End If
                                End If
                                lStart = lIndex + 1
                            End If
                        Next
                    End If
                End If
                            
                ' if RealTime has volumes but no ticks from DataManager, then create the volume and flags arrays
                If DataMgrTicks.Size = 0 Then
                    If RealTimeTicks.IsActiveArray(eBARS_Vol) Then
                        DataMgrTicks.ArrayMask = eBARS_Vol Or DataMgrTicks.ArrayMask
                    End If
                    If RealTimeTicks.IsActiveArray(eBARS_Flags) Then
                        DataMgrTicks.ArrayMask = eBARS_Flags Or DataMgrTicks.ArrayMask
                    End If
                End If
                            
                ' append RT ticks (from "lStart") onto the DM ticks
                lOldSize = DataMgrTicks.Size
                DataMgrTicks.Size = DataMgrTicks.Size + RealTimeTicks.Size - lStart
                
                lMemFrom = gdGetDataPtr(RealTimeTicks.ArrayHandle(eBARS_DateTime)) + lStart * 8
                lMemTo = gdGetDataPtr(DataMgrTicks.ArrayHandle(eBARS_DateTime)) + lOldSize * 8
                If lMemFrom <> 0 And lMemTo <> 0 Then
                    CopyMemory ByVal lMemTo, ByVal lMemFrom, (RealTimeTicks.Size - lStart) * 8
                End If
            
                lMemFrom = gdGetDataPtr(RealTimeTicks.ArrayHandle(eBARS_Close)) + lStart * 8
                lMemTo = gdGetDataPtr(DataMgrTicks.ArrayHandle(eBARS_Close)) + lOldSize * 8
                If lMemFrom <> 0 And lMemTo <> 0 Then
                    CopyMemory ByVal lMemTo, ByVal lMemFrom, (RealTimeTicks.Size - lStart) * 8
                End If
            
                If DataMgrTicks.IsActiveArray(eBARS_Vol) And RealTimeTicks.IsActiveArray(eBARS_Vol) Then
                    lMemFrom = gdGetDataPtr(RealTimeTicks.ArrayHandle(eBARS_Vol)) + lStart * 8
                    lMemTo = gdGetDataPtr(DataMgrTicks.ArrayHandle(eBARS_Vol)) + lOldSize * 8
                    If lMemFrom <> 0 And lMemTo <> 0 Then
                        CopyMemory ByVal lMemTo, ByVal lMemFrom, (RealTimeTicks.Size - lStart) * 8
                    End If
                End If
            
                If DataMgrTicks.IsActiveArray(eBARS_Flags) And RealTimeTicks.IsActiveArray(eBARS_Flags) Then
                    lMemFrom = gdGetDataPtr(RealTimeTicks.ArrayHandle(eBARS_Flags)) + lStart * 8
                    lMemTo = gdGetDataPtr(DataMgrTicks.ArrayHandle(eBARS_Flags)) + lOldSize * 8
                    If lMemFrom <> 0 And lMemTo <> 0 Then
                        CopyMemory ByVal lMemTo, ByVal lMemFrom, (RealTimeTicks.Size - lStart) * 8
                    End If
                End If
            End If
        End If
        If m.bDebug Then DataMgrTicks.DumpToFile AddSlash(App.Path) & "DMTicks2.ASC"
        
        If bLog Then
            d = gdTickCount(False)
            strProfile = strProfile & ", 5=" & Format(d - dProfile, "#0")
            dProfile = d
        End If
    
        ClearCache strSymbol
    End If

If strSymbol = "ES-067" And Bars.IsIntraday Then
    d = d
End If

    ' store the merged ticks to a cache file (for efficiency when a new bar starts)
    Set TickBuffer = m.TickBuffers("^" & strSymbol)
    If Not bTimedOut And Not TickBuffer Is Nothing Then
        ' if the cache file does not yet exist, or if the memory cache has grown very large
        strTemp = App.Path & "\Cache\" & strSymbol & ".TBC" ' (TickBufferCache file)
        If DataMgrTicks.Size = 0 Then
            ' but must delete cache file if no ticks yet exist (otherwise it won't get started correctly)
            KillFile strTemp
        ElseIf Not FileExist(strTemp) Or TickBuffer.Size > 50000 Then
            TickBuffer.Size = 0
            TickBuffer.FreeExtra
            MakeDir App.Path & "\Cache\"
            DataMgrTicks.SerializeDataArrays strTemp, True
            If bLog Then
                d = gdTickCount(False)
                strProfile = strProfile & ", WC=" & Format(d - dProfile, "#0")
                dProfile = d
            End If
        End If
    End If
    Set TickBuffer = Nothing

    ' now that cache has been stored, we can restore the custom start/end
    ' and chop off ticks prior to the requested start session
    Bars.Prop(eBARS_StartTime) = lStartSave
    Bars.Prop(eBARS_EndTime) = lEndSave
    DataMgrTicks.Prop(eBARS_StartTime) = lStartSave
    DataMgrTicks.Prop(eBARS_EndTime) = lEndSave
    If lStartSession > 0 Then
        ' chop ticks prior to start session
        d = Bars.Prop(eBARS_CrossoverTime) / 1440#
        If d = 0 Then d = 0.99999
        d = lStartSession + d - 1
        i = DataMgrTicks.FindDateTime(d)
        If i >= 0 Then
            DataMgrTicks.DeleteFirstBars i
        End If
    End If
    
If 0 And IsIDE Then
    DataMgrTicks.DumpToFile AddSlash(App.Path) & "DMTicks.ASC"
End If

    ' skip these steps if updating ticks (i.e. if "Set DataMgrTicks = Bars" at beginning)
    If Bars.BarsHandle <> DataMgrTicks.BarsHandle Then
    
        ' 6. Chop the bars at the LastDailyDownload + CrossOver Time
        If Bars.Size > 0 Then
            dCrossOver = Bars.Prop(eBARS_CrossoverTime)
            If dCrossOver = 0 Then dCrossOver = 1440
            dCrossOver = dCrossOver / 1440 '(((dCrossOver / 100) * 60) + (dCrossOver Mod 100)) / 1440
            dStartTime = gdFixDateTime(LastDailyDownload + dCrossOver)
            gdBinarySearch Bars.ArrayHandle(eBARS_DateTime), dStartTime, _
                    lStart, eGdSort_Default, 0, Bars.Size
#If 0 Then
            If Bars.Prop(eBARS_PeriodType) = ePRD_IntPF Then
                If Bars(eBARS_Close, lStart + 1) <> kNullData Then
                    lStart = lStart + 1
                    dStartTime = Bars(eBARS_DateTime, lStart)
                    gdBinarySearch DataMgrTicks.ArrayHandle(eBARS_DateTime), dStartTime, _
                            lStart, eGdSort_Default, 0, DataMgrTicks.Size
                    If lStart > 0 Then
                        DataMgrTicks.DeleteFirstBars lStart
                    End If
                    lStart = Bars.Size
                End If
            End If
#End If
            Bars.Size = lStart
            'If bLog Then frmTest.AddList "Chopping: " & dStartTime & " " & lStart, True
        End If
               
        If bLog Then
            d = gdTickCount(False)
            strProfile = strProfile & ", 6=" & Format(d - dProfile, "#0")
            dProfile = d
        End If
        
        ' 7. Make bars from ticks and append to the bars from step number 3
        If Bars.Size > 0 And Bars.Prop(eBARS_Periodicity) > ePRD_Days + 1 Then
            ' if greater than daily bars (e.g. weekly, 2 days, etc.), first get daily bars from last complete bar,
            ' append to daily bars from ticks, then append daily bars to build into correct periodicity
            ' NOTE: should get history from DataMgr using Bars.Prop(eBARS_Symbol) instead of strSymbol
            ' (because we want to get -067 history for -067's and the DataMgr will handle it correctly)
            DM_GetBars DailyBars, Bars.Prop(eBARS_Symbol), "Daily", Bars(eBARS_DateTime, Bars.Size - 1) + 1, LastDailyDownload
            DailyBars.BuildBars "Daily", DataMgrTicks.BarsHandle, True
            Bars.BuildBars Bars.Prop(eBARS_PeriodicityStr), DailyBars.BarsHandle, True
        Else
            ' else just append ticks onto bars using correct periodicity
            Bars.BuildBars Bars.Prop(eBARS_PeriodicityStr), DataMgrTicks.BarsHandle, True
        End If
        If DataMgrTicks.Size > 0 Then
            dTime = RoundToSecond(DataMgrTicks(eBARS_DateTime, DataMgrTicks.Size - 1))
            Bars.Prop(eBARS_LastTickTime) = MinutesFromMidnight(dTime)
        End If
                
        If bLog Then
            d = gdTickCount(False)
            strProfile = strProfile & ", 7=" & Format(d - dProfile, "#0")
            dProfile = d
        End If
        
        ' 8. Put bid/ask in for options
        If DataMgrTicks.IsActiveArray(eBARS_Bid) And DataMgrTicks.Size >= 4 And DataMgrTicks.Prop(eBARS_SymbolID) = 0 Then
            Bars(eBARS_Bid, Bars.Size - 1) = DataMgrTicks(eBARS_Bid, 3)
            Bars(eBARS_Ask, Bars.Size - 1) = DataMgrTicks(eBARS_Ask, 3)
            Bars(eBARS_BidSize, Bars.Size - 1) = DataMgrTicks(eBARS_BidSize, 3)
            Bars(eBARS_AskSize, Bars.Size - 1) = DataMgrTicks(eBARS_AskSize, 3)
        End If
        UpdateBidAsk Bars

        If bLog Then
            d = gdTickCount(False)
            strProfile = strProfile & ", 8=" & Format(d - dProfile, "#0")
            dProfile = d
        End If
        
        ' 9. get settle prices from DataMgr (if daily bars)
        If bIncludeSettles Then
            ' check one day at a time
            dTime = -1
            For lIndex = Bars.Size - 1 To 0 Step -1
                ' TLB 12/19/2007: this check is VERY important, since some odd bar alignments can cause
                ' "daily" periodicity with LOTS of bars with same date -- only call GetBars once per day
                If dTime <> Int(Bars(eBARS_DateTime, lIndex)) Then
                    dTime = Int(Bars(eBARS_DateTime, lIndex))
                    If dTime <= LastDailyDownload Then
                        If lIndex = Bars.Size - 1 Then
                            Bars.Prop(eBARS_PriceHasSettled) = True
                        End If
                        Exit For
                    End If
                    DailyBars.Size = 0
                    ' NOTE: should get history from DataMgr using Bars.Prop(eBARS_Symbol) instead of strSymbol
                    ' (because we want to get -067 history for -067's and the DataMgr will handle it correctly)
                    DM_GetBars DailyBars, Bars.Prop(eBARS_Symbol), 0, dTime, dTime
                    If DailyBars.Size = 1 And DailyBars.Prop(eBARS_PriceHasSettled) Then
                        Bars(eBARS_Close, lIndex) = DailyBars(eBARS_Close, 0)
                        If lIndex = Bars.Size - 1 Then
                            Bars.Prop(eBARS_PriceHasSettled) = True
                        End If
                    End If
                End If
            Next
        End If

        If m.bDebug Then Bars.DumpToFile AddSlash(App.Path) & "Final.ASC"
    End If
    
    If bLog Then
        d = gdTickCount(False)
        strProfile = strProfile & ", 9=" & Format(d - dProfile, "#0")
        dProfile = d
    End If
    
    ' Let other people have control of the tick buffer again
    If m.TickBuffers.Exists(Str(Bars.BarsHandle) & "X") Then
        m.TickBuffers.Key(Str(Bars.BarsHandle) & "X") = Str(Bars.BarsHandle)
    End If
    
    Set DataMgrTicks = Nothing
    Set RealTimeTicks = Nothing
    Set TickBuffer = Nothing
    Set DailyBars = Nothing
    SpliceBars = True
                
    If bLog Then
        d = Round(gdTickCount(False) - dStartingProfile)
        frmTest.AddList Str(d) & " ms for: " & strSymbol & ", " & Bars.Prop(eBARS_PeriodicityStr)
        frmTest.AddList strProfile
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.SpliceBars"
End Function

Public Function SpliceBarsSalmon(Bars As cGdBars, Optional ByVal nStartSession = 0, _
                Optional ByVal bIgnoreSettle As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim i&, d#, dTime#
    Dim strSymbol$, strText$
    Dim TickBuffer As cGdBars       ' Ticks from the tick buffer
    Dim SalmonData As cGdBars      ' intraday data (minutes or ticks) from Salmon (after LastDailyDownload)
    Dim SessionData As cGdBars
    Dim SymInf As cSymbolInfo
    
    Dim bLog As Boolean
    Dim strProfile$, dProfile#, dStartingProfile#
    
    ' don't do anything if real-time is not active
    If Not Active Or Bars Is Nothing Then Exit Function
    
    If Not SalmonIsRunning Then Exit Function

    strSymbol = RollSymbolForDate(Bars.Prop(eBARS_Symbol))
    
    ' and don't do anything if the symbol has not yet made it into the real-time symbol list
    ' (e.g. if walking up/down quote board then keep data from current session update)
'??? do we still need this?
If 0 Then
    If Not m.UniqueSymbols.BinarySearch(strSymbol) Then
        Exit Function
    End If
End If

    ' Intraday bar periods other than minute bars now require the full tick database
    If Len(BarPeriodError(strSymbol, Bars.Prop(eBARS_Periodicity))) > 0 Then
        Bars.Size = 0
        Exit Function
    End If
    
    Set SymInf = SymbolInfo(strSymbol)
    If SymInf Is Nothing Then Exit Function
    
    ' clear the TickBuffer
    Set TickBuffer = m.TickBuffers(Str(Bars.BarsHandle))
    If Not TickBuffer Is Nothing Then
        TickBuffer.Size = 0
        TickBuffer.Prop(eBARS_CustomString) = ""
    End If
    
    ' TLB 2/17/2010: we need to chop off forecast bars since that is how the old SpliceBars worked
    ' (otherwise the AddForecast bars doubles things up and can mess up the calcs for charts)
    ' TLB 4/26/2011: also needed for intraday bars (to avoid null data bars between sessions)
    For i = Bars.Size - 1 To 0 Step -1
        If Bars(eBARS_Close, i) <> kNullData Or (i = 0) Then
            Bars.Size = i + 1
            ' TLB 4/9/2012: there was a bug that caused the need to double-check for a null bar just
            ' prior to the last good data bar (the bug's been fixed, but this is still a good check)
            If Bars(eBARS_Close, i - 1) <> kNullData Then
                Exit For
            End If
        End If
    Next
    
    If Bars.Prop(eBARS_PeriodType) >= ePRD_Days Then
        ' For DAILY data (i.e. non-intraday) ...
        strText = ""
        SpliceBarsSalmon = SymInf.UpdateDailyBars(Bars, , strText, bIgnoreSettle)
        If Not TickBuffer Is Nothing Then
            TickBuffer.Prop(eBARS_CustomString) = strText
        End If
    Else
        ' For INTRADAY data ...
        If SymInf.GetDataRequestStatus(Bars.Prop(eBARS_Periodicity)) <> eSalmonAvailable Then
            If Not TickBuffer Is Nothing Then
                TickBuffer.Prop(eBARS_CustomString) = "<PENDING>"
            End If
        Else
            bLog = frmTest.Visible
            If bLog Then
                dProfile = gdTickCount(False)
                dStartingProfile = dProfile
            End If
            
            ' chop off any bars after the LastDailyDownload
            dTime = Bars.Prop(eBARS_CrossoverTime) / 1440#
            If dTime = 0 Then dTime = 0.99999
            dTime = LastDailyDownload + dTime
            i = Bars.FindDateTime(dTime)
            If Bars.Size > i Then
                Bars.Size = i
            End If
            
            ' call GetStreamedTicks just to flush the newest ticks into Salmon's history
            SalmonGetNewTicks strSymbol
            ' and reset this TickBuffer
            If Not TickBuffer Is Nothing Then
                TickBuffer.Prop(eBARS_CustomString) = "<OK>"
                TickBuffer.Size = 0
            End If
        
' setup the DataMgrTicks bars
'??? If Bars.Prop(eBARS_PeriodType) = ePRD_EachTick Then

            ' Get TODAY'S data from Salmon
'DebugLog "f=" & Str(Int(gdTickCount))
            Set SalmonData = SymInf.GetIntradayData(Bars.Prop(eBARS_Periodicity))
'DebugLog "g=" & Str(Int(gdTickCount))
            If SalmonData.Size > 0 Then
                ' set custom start/stop time
                SalmonData.Prop(eBARS_StartTime) = Bars.Prop(eBARS_StartTime)
                SalmonData.Prop(eBARS_EndTime) = Bars.Prop(eBARS_EndTime)
        
                ' chop off any ticks prior to the requested start session
                If nStartSession <= LastDailyDownload Then
                    nStartSession = LastDailyDownload + 1
                End If
                dTime = SalmonData.Prop(eBARS_CrossoverTime) / 1440#
                If dTime = 0 Then dTime = 0.99999
                dTime = nStartSession + dTime - 1
                i = SalmonData.FindDateTime(dTime)
                If i > 0 Then
                    SalmonData.DeleteFirstBars i
                End If
            End If
            
            If SalmonData.Size > 0 Then
                SpliceBarsSalmon = True
                
                If Bars.Prop(eBARS_PeriodType) = ePRD_EachTick Then
                    ' just append tick-by-tick onto the tick-by-tick array
                    gdAppendBars Bars.BarsHandle, SalmonData.BarsHandle, 0
                ElseIf Bars.Prop(eBARS_FractZen) = 0 Then
                    ' append ticks onto bars using correct periodicity
                    ' (BuildBars will now also build X-minute bars from either 1-minute bars or ticks)
                    Bars.BuildBars Bars.Prop(eBARS_PeriodicityStr), SalmonData.BarsHandle, True
                Else
                    ' for merged FractZen bars over multiple sessions, need to build for each session
                    For nStartSession = SalmonData.SessionDate(0) To SalmonData.SessionDate(SalmonData.Size - 1)
                        Set SessionData = SalmonData.MakeCopy
                        dTime = SalmonData.Prop(eBARS_CrossoverTime) / 1440#
                        If dTime = 0 Then dTime = 0.99999
                        dTime = nStartSession + dTime - 1
                        i = SessionData.FindDateTime(dTime)
                        If i > 0 Then
                            SessionData.DeleteFirstBars i
                        End If
                        i = SessionData.FindDateTime(dTime + 1)
                        If i < SessionData.Size Then
                            SessionData.Size = i
                        End If
                        i = g.FractZen.GetFractZenRange(Bars.Prop(eBARS_Symbol), nStartSession)
                        If i > 0 And SessionData.Size > 0 Then
                            Bars.BuildBars Str(i) & "b", SessionData.BarsHandle, True
                        End If
                    Next
                End If
                
                ' set LastTickTime property (use time of last trade if tick-by-tick)
                dTime = MinutesFromMidnight(SalmonData(eBARS_DateTime, SalmonData.Size - 1))
                If SalmonData.Prop(eBARS_PeriodType) = ePRD_Minutes Then
                    ' but for minute bars, use the LastTickTime if accurate
                    ' (i.e. must be within 1 minute from the last minute bar)
                    d = SalmonData.Prop(eBARS_LastTickTime)
                    If d >= dTime - 1 And d <= dTime Then
                        dTime = d
                    Else
                        DebugLog "SpliceBarSalmon time oddity for " & strSymbol & _
                            ":  MB = " & Str(dTime) & ", LTT = " & Str(d)
                        dTime = dTime - 1 ' must just default to 1 minute ago
                    End If
                End If
                If dTime >= 0 Then
                    Bars.Prop(eBARS_LastTickTime) = dTime
                End If
                        
                If bLog Then
                    d = gdTickCount(False)
                    strProfile = strProfile & ", 7=" & Format(d - dProfile, "#0")
                    dProfile = d
                End If
                
                If m.bDebug Then Bars.DumpToFile App.Path & "\Cache\" & strSymbol & "_.txt"
            End If
            
            If bLog Then
                d = Round(gdTickCount(False) - dStartingProfile)
                frmTest.AddList Str(d) & " ms for: " & strSymbol & ", " & Bars.Prop(eBARS_PeriodicityStr)
                frmTest.AddList strProfile
            End If
        End If
    End If
    
ErrExit:
    Set TickBuffer = Nothing
    Set SymInf = Nothing
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.SpliceBarsSalmon"
End Function


#If 0 Then
Public Function UpdateBidAsk_OLD(Bars As cGdBars, Optional ByVal bUpdateLastTickTime As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strTemp As String, dPrice As Double, i&, dTime As Double, bNewPrice As Boolean, bNewSize As Boolean
    Dim dOldBid#, dOldAsk#, dOldBidSize#, dOldAskSize#

' New Bid (GenSymbol, Date.Time, BidPrice, BidSize)

    dOldBid = Bars(eBARS_Bid, Bars.Size - 1)
    dOldAsk = Bars(eBARS_Ask, Bars.Size - 1)
    dOldBidSize = Bars(eBARS_BidSize, Bars.Size - 1)
    dOldAskSize = Bars(eBARS_AskSize, Bars.Size - 1)

    If (Bars.ArrayMask And eBARS_Bid) <> 0 Then
        i = m.LastBid.Index(Bars.Prop(eBARS_Symbol))
        If i > 0 Then
            strTemp = m.LastBid(i)
            dPrice = Val(Parse(strTemp, vbTab, 3))
            'dPrice = Bars.RoundToPrice(dPrice)
            dTime = RoundToSecond(Val(Parse(strTemp, vbTab, 2)))
            If dPrice <> 0 Then
                bNewPrice = True
                If Bars.Size = 0 Then Bars.Size = 1
                Bars(eBARS_Bid, Bars.Size - 1) = dPrice
            End If
            If (Bars.ArrayMask And eBARS_BidSize) <> 0 Then
                bNewSize = True
                dPrice = Val(Parse(strTemp, vbTab, 4))
                Bars(eBARS_BidSize, Bars.Size - 1) = dPrice
            End If
        End If
    End If

    If (Bars.ArrayMask And eBARS_Ask) <> 0 Then
        i = m.LastAsk.Index(Bars.Prop(eBARS_Symbol))
        If i > 0 Then
            strTemp = m.LastAsk(i)
            dPrice = Val(Parse(strTemp, vbTab, 3))
            'dPrice = Bars.RoundToPrice(dPrice)
            dTime = RoundToSecond(Val(Parse(strTemp, vbTab, 2)))
            If dPrice <> 0 Then
                bNewPrice = True
                If Bars.Size = 0 Then Bars.Size = 1
                Bars(eBARS_Ask, Bars.Size - 1) = dPrice
            End If
            If (Bars.ArrayMask And eBARS_AskSize) <> 0 Then
                bNewSize = True
                dPrice = Val(Parse(strTemp, vbTab, 4))
                Bars(eBARS_AskSize, Bars.Size - 1) = dPrice
            End If
        End If
    End If
    
    If bNewPrice Or bNewSize Then
        If bNewPrice Then FixTicks Bars, True

        If dOldBid <> Bars(eBARS_Bid, Bars.Size - 1) Or dOldAsk <> Bars(eBARS_Ask, Bars.Size - 1) _
            Or dOldBidSize <> Bars(eBARS_BidSize, Bars.Size - 1) Or dOldAskSize <> Bars(eBARS_AskSize, Bars.Size - 1) Then
        
            UpdateBidAsk = True
            If bUpdateLastTickTime Then
                ' convert date/time to minutes since midnight
                dTime = (dTime - Int(dTime)) * 1440
                Bars.Prop(eBARS_LastTickTime) = dTime
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.UpdateBidAsk", eGDRaiseError_Raise
    
End Function
#End If

Public Function UpdateBidAsk(Bars As cGdBars, Optional ByVal bUpdateLastTickTime As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim dPrice#, dTime#, dSize#, strSymbol$, strStatus$, bNewBar As Boolean
    Dim aLastQuote As cGdArray
    Dim SymInf As cSymbolInfo

    strSymbol = UCase(Trim(Bars.Prop(eBARS_Symbol)))
    
    If SalmonIsRunning Then
        ' TLB: when salmon is running, the only calls to this routine are most likely not updating
        ' trades (else UpdateBars would be called instead of this), so a tick buffer most likely
        ' was not created, so we need to just store the "status" directly in these bars
        Set SymInf = SymbolInfo(strSymbol)
        strStatus = Bars.Prop(eBARS_CustomString)
        If SymInf.UpdateDailyBars(Bars, bNewBar, strStatus, , bUpdateLastTickTime) Then
            Bars.Prop(eBARS_CustomString) = strStatus
            UpdateBidAsk = True
        End If
        Set SymInf = Nothing
        Exit Function
    End If

    strSymbol = RollSymbolForDate(strSymbol)
    If (Bars.ArrayMask And eBARS_Bid) <> 0 And (Bars.ArrayMask And eBARS_Ask) <> 0 Then
        Set aLastQuote = m.LastQuote(strSymbol)
        If Not aLastQuote Is Nothing Then
            ' see if a new bid price
            dPrice = aLastQuote.Num(LQ(eLQ_BidPrice))
            If dPrice > 0 And Bars(eBARS_Bid, Bars.Size - 1) <> dPrice Then
                UpdateBidAsk = True
                If Bars.Size = 0 Then Bars.Size = 1
                Bars(eBARS_Bid, Bars.Size - 1) = dPrice
            End If
            ' see if a new ask price
            dPrice = aLastQuote.Num(LQ(eLQ_AskPrice))
            If dPrice > 0 Then
                UpdateBidAsk = True
                If Bars.Size = 0 Then Bars.Size = 1
                Bars(eBARS_Ask, Bars.Size - 1) = dPrice
            End If
            ' fix ticks if there was a new price
            If UpdateBidAsk Then
                FixTicks Bars, True
            End If
            
            ' see if a new bid size
            dSize = aLastQuote.Num(LQ(eLQ_BidSize))
            If (Bars.ArrayMask And eBARS_BidSize) <> 0 Then
                If dSize > 0 And Bars(eBARS_BidSize, Bars.Size - 1) <> dSize Then
                    UpdateBidAsk = True
                    Bars(eBARS_BidSize, Bars.Size - 1) = dSize
                End If
            End If
            ' see if a new ask size
            dSize = aLastQuote.Num(LQ(eLQ_AskSize))
            If (Bars.ArrayMask And eBARS_AskSize) <> 0 Then
                If dSize > 0 And Bars(eBARS_AskSize, Bars.Size - 1) <> dSize Then
                    UpdateBidAsk = True
                    Bars(eBARS_AskSize, Bars.Size - 1) = dSize
                End If
            End If
            
            ' update the time
            If bUpdateLastTickTime Then
                dTime = aLastQuote.Num(LQ(eLQ_BidAskTime))
                If dTime > 0 Then
                    ' convert date/time to minutes since midnight
                    dTime = (dTime - Int(dTime)) * 1440
                    If Bars.Prop(eBARS_LastTickTime) < dTime Then
                        Bars.Prop(eBARS_LastTickTime) = dTime
                    End If
                End If
            End If
            
            Set aLastQuote = Nothing
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.UpdateBidAsk", eGDRaiseError_Raise
    
End Function

Public Function IsServerActive(Optional bAndConnectedToFeed As Boolean = False) As Boolean
On Error GoTo ErrSection:

    If Not g.bUnloading Then
        If m.RT.FindControlID(m.strServerNode) <> 0 Then
            If bAndConnectedToFeed Then
                ' if good status and received a message within last 75 seconds
                If m.nServerStatus <> 0 And m.dLastMsgTime > gdTickCount - 75000 Then
                    IsServerActive = True
                End If
            Else
                IsServerActive = True
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.IsServerActive", eGDRaiseError_Raise
    
End Function

' Called at startup and when config changes -- bSetActive:
'  missing = use current setting in the registry (off/on)
Public Sub Init(Optional ByVal bSetActive As Variant, Optional ByVal strDisconnectReason$)
On Error GoTo ErrSection:

    Dim i&, iMax&
    Dim strKey$, strServer$
    Dim bActive As Boolean, bWasActive As Boolean
    Dim Bars As cGdBars
    Dim strTemp As String
    Dim iPos&
    Dim aDirs As New cGdArray
    Dim frm As Form
    Static bAlreadyDone As Boolean

    StatusMsg
    
    bWasActive = m.RT.Active
    
    ' get settings from registry
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    If IsMissing(bSetActive) Then
        ' get registry setting
        If GetRegistryValue(rkLocalMachine, strKey, "ActivateRT", vbUnchecked) <> vbUnchecked Then
            bActive = True
        End If
    Else
        ' save registry setting
        bActive = bSetActive
        If bActive Then
            SetRegistryValue rkLocalMachine, strKey, "ActivateRT", vbChecked
        Else
            SetRegistryValue rkLocalMachine, strKey, "ActivateRT", vbUnchecked
        End If
    End If
    
    ' see if a valid RealTime server and is currently updating tick data
    strServer = ""
    If HasModule("RTG") Or HasModule("RTE") Then
        strServer = GetRegistryValue(rkLocalMachine, strKey, "ActiveFeedRT", "")
        If Not DirExist(App.Path & "\..\RealTime\" & strServer) Then
            strServer = ""
        End If
        If Len(strServer) = 0 Then
            ' if active feed not yet saved, set to first that exists
            aDirs.GetMatchingFiles AddSlash(App.Path) & "..\RealTime\*", False, True
            For i = 0 To aDirs.Size - 1
                strTemp = aDirs(i)
                If Right(strTemp, 1) = "\" Then
                    strServer = Left(strTemp, Len(strTemp) - 1)
                    SetRegistryValue rkLocalMachine, strKey, "ActiveFeedRT", strServer
                    Exit For
                End If
            Next
        End If
    End If
    If Len(strServer) = 0 Then
        bActive = False
    Else
        With frmMain.tbToolbar.Tools("ID_RealTime")
            .ToolTipText = "Connect with " & strServer & _
                " RealTime server (for symbols on the Quote Board)"
            .Visible = True
        End With
    End If
    
    ' check for valid codes
    If bActive Then
        If UCase(strServer) = "ESIGNAL" And Not HasModule("RTE") Then
            bActive = False
            InfBox "Please contact Genesis Sales to add the eSignal streaming functionality to Trade Navigator.", "i=!", , "Not Authorized"
        ElseIf UCase(strServer) = "GENESIS" And Not HasModule("RTG") Then
            bActive = False
            InfBox "Please contact Genesis Sales to add the Genesis streaming functionality to Trade Navigator.", "i=!", , "Not Authorized"
        End If
    End If
    
    If Not bActive Then
        g.nReplaySession = 0
        If g.nReplayAccountID <> 0 Then
            g.nReplayAccountID = 0
            RefreshAccountCombos
            
            If Not g.ConsoleForms Is Nothing Then
                g.ConsoleForms.FilterGrids
            End If
        End If
    End If
    
    ' Make sure to cancel quick fill orders before deactivating streaming (07/21/2006 DAJ)...
    If (bWasActive = True) And (bActive = False) Then
        ' NOTE: if we change the text of the reason below, then also change it in cAutoTradeItem.Disable
        DeactivateTrading "Realtime streaming going inactive"
        
        ' TLB 6/15/2010: and store the disconnect reason
        DisconnectReason = strDisconnectReason
    End If

    ' see if changing real-time servers
    If UCase(strServer) <> UCase(m.strServer) Then
        If Active Then
            ' must first deactivate the old one
            Active = False
            Sleep 5 '(allow time for messages to be passed)
        End If
        m.strServer = strServer
        m.strServerNode = strServer & "RT"
    End If

    ' set active or inactive
    If (bActive <> bWasActive) Or Not bAlreadyDone Then
        bAlreadyDone = True
        frmQuotes.tmrRealTime.Enabled = False
        m.dLastMsgTime = 0 '(so won't trigger a reconnect)
        frmMain.tmrPredLabs.Enabled = False
        
        m.LastQuote.Clear
        m.OptNavStreamingList.Size = 0
        m.OptNavHistSymbols.Size = 0
                
        If bActive Then
            ConnectionStatus = eGDConnectionStatus_Connecting
        Else
            ConnectionStatus = eGDConnectionStatus_Disconnecting
        End If
        
        If bActive <> bWasActive Then
            If bActive Then
                SalmonStart True, True
                SalmonInitNewStream False ' just to initiate request (will get the response later)
                m.strIPandKey = ""
            End If
        
            SetRTbutton 1 '(yellow)
        End If
        
        ClearCache
        
        iMax = 0
        For i = 1 To m.TickBuffers.Count
            Set Bars = m.TickBuffers(i)
            If Not Bars Is Nothing Then
                If Bars.Size > iMax Then iMax = Bars.Size
                Bars.Size = 0
            End If
        Next
        DebugLog "RT: TickBufs=" & Str(m.TickBuffers.Count) _
            & ", MaxBufSize=" & Str(iMax) _
            & ", TickSyms=" & Str(m.TickBufferSymbols.Size) _
            & ", TickRetr=" & Str(m.TickRetrieval.Count) _
            & ", BidAsk=" & Str(m.LastQuote.Count)
        
        If bActive And Not m.bReconnecting Then
            If g.nReplaySession > 0 Then
                g.nReplayAccountID = StreamReplayAccount("", m.dFeedTime)
                RefreshAccountCombos
                
                If Not g.ConsoleForms Is Nothing Then
                    g.ConsoleForms.FilterGrids
                End If
            ElseIf Not FileExist(App.Path & "\RTS\Record.flg") Then
                ' connect brokers now?
                AutoBrokerConnect
            End If
        End If
        
        Active = bActive
        
        If Not bActive Then
            SalmonStop
        End If
        
        'un-sync non-primary game mode forms if user turns on real-time after start of game
        If g.ChartGlobals.nGameInProg = 2 Then
            g.ChartGlobals.nGameInProg = 1
            For i = 0 To Forms.Count - 1
                If IsFrmChart(Forms(i)) Then
                    Set frm = Forms(i)
                    If frm.IsInGameMode Then
                        If frm.GameReplayMode = eGDReplayMode_Sync Then
                            frm.ClearReplaySync
                        Else
                            frm.GameMode.ReplayAll = False
                        End If
                    End If
                End If
            Next
            Set frm = Nothing
        End If
        
        If Not Active Then
            ' if just starting, need to refresh the quote board
            If g.bStarting Then
                frmQuotes.TotalRefresh True
            ' else if real-time is turning inactive, then need to refresh all data
            ElseIf bWasActive Then
                If Not g.bUnloading Then
                    ' reset data from just data manager
                    RefreshAllFormData True
                End If
            End If
        End If
    End If
    
    If Not Active Then
        ConnectionStatus = eGDConnectionStatus_Disconnected
    End If
    
    SetRTbutton
    If g.nReplaySession > 0 Then frmReplay.FixButtons

    ' resync broker feed list
    ResyncSubscriptionList m.nBrokerFeed

    ' TLB 4/16/2007: need to set next auto-update time when starting/stopping realtime
    CalcNextTryTime

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.Init", eGDRaiseError_Raise
    
End Sub

Private Function RtSymbolInfo(ByVal strSymbol$) As String
On Error GoTo ErrSection:

    Dim iPos&, dTimeout#

    If IsServerActive And m.RT.Active Then
        ' need to retrieve the symbol delays if empty
        If m.RtSymbolInfo.Size = 0 And Not ActiveRTG Then
            m.RT.CreateMessage m.strServerNode, 12, "", , True
            dTimeout = gdTickCount + 2000
            Do
                DoEvents
            Loop While m.RtSymbolInfo.Size = 0 And gdTickCount <= dTimeout
        End If
        
        ' find info for specified symbol
        strSymbol = RollSymbolForDate(UCase(Trim(strSymbol)))
        If m.RtSymbolInfo.BinarySearch(strSymbol & "|", iPos, eGdSort_MatchUsingSearchStringLength Or eGdSort_IgnoreCase) Then
            RtSymbolInfo = m.RtSymbolInfo(iPos)
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.RtSymbolInfo", eGDRaiseError_Raise
End Function

' Returns the real-time delay (in minutes) for a symbol
' (-1 is returned if an unknown delay or symbol)
Public Function SymbolDelay(ByVal strSymbol$) As Long
On Error GoTo ErrSection:

    Dim iPos&, strDelay$
    
    ' get delay time (after symbol)
    strSymbol = RollSymbolForDate(strSymbol)
    strDelay = Parse(RtSymbolInfo(strSymbol), "|", 2)
    If Len(strDelay) > 0 Then
        SymbolDelay = Val(strDelay)
    Else
        SymbolDelay = -1 '(unknown)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.SymbolDelay", eGDRaiseError_Raise
    
End Function

' Returns the RT feed symbol for a symbol
Public Function FeedSymbol(ByVal strSymbol$) As String
On Error GoTo ErrSection:

    strSymbol = RollSymbolForDate(strSymbol)
    FeedSymbol = Parse(RtSymbolInfo(strSymbol), "|", 3)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.FeedSymbol", eGDRaiseError_Raise
End Function

' Returns true if user is enabled for depth-of-market for the specified symbol
Public Function CanHaveMarketDepth(ByVal strSymbol$) As Boolean
On Error GoTo ErrSection:

    Dim nByte&, bEnabled As Boolean

    strSymbol = RollSymbolForDate(strSymbol)
    If ActiveRTG Then
        ' for US Stocks, first just check for the D_NASD enablement
        If SecurityType(strSymbol) = "S" Then
            If InStr(strSymbol, "@") = 0 Then
                bEnabled = HasModule("D_NASD", False)
            End If
        End If
        If Not bEnabled Then
            ' check first bit of flags byte (4th field) in RtSymbolInfo for this symbol
            ' (this check will only work if this symbol had already been in the stream)
            nByte = Val(Parse(RtSymbolInfo(strSymbol), "|", 4))
            bEnabled = GetBit(nByte, 1)
        End If
    ElseIf Active Then
        ' e.g. for eSignal
        bEnabled = True
    End If

ErrExit:
    CanHaveMarketDepth = bEnabled
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.CanHaveMarketDepth", eGDRaiseError_Raise
End Function

Private Sub DrawRealtimeBtn(eState As eBtnState)
On Error Resume Next:

    Dim i&
    Dim aButtons As cGdArray
    Dim pB As PictureBox
    Dim oButton As cPicBoxButton
        
    If Not g.bUnloading Then
        Set pB = frmMain.pbTbBack(0)
        Set aButtons = frmMain.TbButtonsArray(kTbGeneral)
        If Not pB Is Nothing And Not aButtons Is Nothing Then
            For i = 0 To aButtons.Size - 1
                Set oButton = aButtons(i)
                If Not oButton Is Nothing Then
                    If oButton.BtnID = "ID_RealTime" Then
                        oButton.BtnState = eState
                        oButton.Draw pB, aButtons, 0
                        pB.Refresh
                        Exit For
                    End If
                End If
            Next
        End If
        Set pB = Nothing
        Set aButtons = Nothing
    End If

End Sub

Public Sub SetRTbutton(Optional ByVal iPicture% = -1)
On Error GoTo ErrSection:

    Static iCurrPicture% '0=Red, 1=Yellow, 2=Green
    Static bInProgress As Boolean

    If g.bUnloading Or bInProgress Then Exit Sub ' so won't be re-entrant from ToolbarReset
    bInProgress = True
    
    If iPicture < 0 Then
        If Not m.RT.Active Then
            iPicture = 0 'Red
        ElseIf g.RealTime.IsServerActive(True) Then
            iPicture = 2 'Green
        Else
            iPicture = 1 'Yellow
        End If
    End If
    If iPicture <> iCurrPicture Then
        If Not g.bStarting And (iCurrPicture = 2 Or iPicture = 2) Then
            ' reset things in case have RTGPLAT or RTGGOLD (Plat or Gold only when streaming)
            SetMainCaption
            ToolbarReset
        End If
        iCurrPicture = iPicture
        With frmMain.tbToolbar.Tools("ID_RealTime")
            frmMain.tbToolbar.Tag = "IGNORE"
            Select Case iPicture
            Case 2
                Set .Picture = Picture16("kGreenLight")
                .State = ssChecked
                DrawRealtimeBtn eBtnState_Selected
            Case 1
                Set .Picture = Picture16("kYellowLight")
                .State = ssChecked
                DrawRealtimeBtn eBtnState_InProg
            Case Else
                Set .Picture = Picture16("kRedLight")
                .State = ssUnchecked
                DrawRealtimeBtn eBtnState_Neutral
            End Select
            frmMain.tbToolbar.Tag = ""
            If g.nReplaySession > 0 Then
                .Enabled = False
            Else
                .Enabled = True
            End If
        End With
        With frmMain.tbToolbar
            If .Redraw = False Then
                ' do this so display will update
                .Redraw = True
                .Redraw = False
            End If
        End With
    End If
    
    ' see if need to automatically reconnect
    If iPicture = 1 Then
        If m.dLastMsgTime > 0 And Not g.bStarting And Not g.bUnloading And (g.nReplaySession = 0) Then
            If ConnectionStatus = eGDConnectionStatus_Connected Or ConnectionStatus = eGDConnectionStatus_Connecting Then
                If UCase(m.strServerNode) = "GENESISRT" Then
                    DisconnectReason = "Stream Dropped"
                    frmMain.tmrMain.Tag = "RECONNECT"
                End If
            End If
        End If
    End If
    
ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "cRealTime.SetRTbutton", eGDRaiseError_Raise
End Sub

Public Sub FixTicks(Ticks As cGdBars, Optional ByVal bJustLastTick As Boolean = False, _
                        Optional ByVal bFromDataMgr As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, nFirst&, dMinMove#, hDates&, hPrices&, iInvalidStart&, dDateTime#
    Dim dPastGoodPrice#, strSymbol$, strBaseSymbol$
    Dim bSkipPastGood As Boolean, bEachTick As Boolean
    Dim DailyBars As cGdBars
    
    ' just check for skip files once every 15 seconds
    Static dLastChecked As Double, bSkipRT As Boolean, bSkipDM As Boolean
    If gdTickCount > dLastChecked + 15000 Then
        bSkipDM = FileExist(App.Path & "\SkipFixTicks.DM")
        bSkipRT = FileExist(App.Path & "\SkipFixTicks.RT")
        dLastChecked = gdTickCount
    End If
    If bFromDataMgr Then
        If bSkipDM Then bSkipPastGood = True
    Else
        If bSkipRT Then bSkipPastGood = True
    End If

    If Ticks.Size = 0 Then Exit Sub

    hDates = Ticks.ArrayHandle(eBARS_DateTime)
    hPrices = Ticks.ArrayHandle(eBARS_Close)
    dMinMove = Ticks.MinMove(Date)

    ' for futures, get past good tick (e.g. previous day's settle) to use in FixPrice
    If Ticks.SecurityType = "F" Then
        ' get base symbol, but also make sure it's not an option symbol
        strSymbol = Ticks.Prop(eBARS_Symbol)
        If InStr(strSymbol, " ") > 0 Then
            ' for a futures option symbol, don't count on the min move
            dMinMove = 0
        ElseIf Not bSkipPastGood Then
            strBaseSymbol = Parse(strSymbol, "-", 1)
            If m.PastGoodPrice.Exists(strBaseSymbol) Then
                dPastGoodPrice = m.PastGoodPrice(strBaseSymbol)
            Else
                ' get past good price and store in collection for quick retrieval
                Set DailyBars = New cGdBars
                ' TLB 10/26/2011: during streaming replay, we need to get daily data from week prior to replay date
                If g.nReplaySession > 0 Then
                    i = g.nReplaySession
                Else
                    i = Date
                End If
                If DM_GetBars(DailyBars, strSymbol, 0, i - 7, 0, , , , False) Then
                    dPastGoodPrice = DailyBars(eBARS_Close, DailyBars.Size - 1)
                End If
                Set DailyBars = Nothing
                m.PastGoodPrice(strBaseSymbol) = dPastGoodPrice
            End If
        End If
    End If

    ' fix last or all ticks (go backwards so can easily delete blocks of invalid times)
    If bJustLastTick Then
        nFirst = Ticks.Size - 1
    Else
        nFirst = 0
    End If
#If 0 Then
    ' TLB: this C++ call didn't help as much as I thought it might, and not yet thoroughly tested
    gdBarsFixTicksHelper Ticks.BarsHandle, nFirst, dPastGoodPrice
#Else
    If Ticks.Prop(eBARS_PeriodType) = ePRD_EachTick Then
        bEachTick = True
    End If
    iInvalidStart = -1
    For i = Ticks.Size - 1 To nFirst Step -1
        ' fix date (round to second and make sure it's a valid time)
        dDateTime = gdFixDateTime(gdGetNum(hDates, i))
        ' check for valid session date (e.g. if outside custom start/end time)
        If Ticks.SessionDateForTime(dDateTime, True) <= 0 Then
            ' mark the start of invalid ticks
            If iInvalidStart < 0 Then iInvalidStart = i
        Else
            If iInvalidStart >= 0 Then
                ' delete the block of invalid ticks
                Ticks.DeleteSomeBars i + 1, iInvalidStart - i
                iInvalidStart = -1
            End If
            gdSetNum hDates, i, dDateTime
            
            ' fix price (conversion factor and round to nearest tick move)
            If bEachTick Then
                gdSetNum hPrices, i, FixPrice(gdGetNum(hPrices, i), dMinMove, dPastGoodPrice)
            Else
                Ticks(eBARS_Bid, i) = FixPrice(Ticks(eBARS_Bid, i), dMinMove, dPastGoodPrice)
                Ticks(eBARS_Ask, i) = FixPrice(Ticks(eBARS_Ask, i), dMinMove, dPastGoodPrice)
            End If
        End If
    Next
    If iInvalidStart >= 0 Then
        ' delete the block of invalid ticks
        Ticks.DeleteFirstBars iInvalidStart + 1
        iInvalidStart = -1
    End If
#End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.FixTicks", eGDRaiseError_Raise
End Sub

' Fix Price:
' - rounds price to the symbol's min move
' - auto-fix conversion factor issues for prices from the data feed (futures only)
' - remove really obvious spikes
Private Function FixPrice(ByVal dPrice#, ByVal dMinMove#, ByVal dPastGoodPrice#) As Double
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

    Dim bFixed As Boolean
    
    FixPrice = dPrice
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
    If dMinMove > 0 Then
        'FixPrice = Int(FixPrice / dMinMove + 0.5) * dMinMove
        FixPrice = RoundToMinMove(FixPrice, dMinMove)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.FixPrice"
End Function

' To set the symbol for the depth-of-market messages
' - pass empty string in order to stop the depth-of-market stream
' - changing the symbol automatically clears the depth-of-market buffer
' - JM(03-27-2008): added create flag to allow chart to get object without creating new one (to show on chart)
Public Function AddMarketDepthSymbol(ByVal strSymbol$, ByVal bCreate As Boolean) As cBidAskDepth
On Error GoTo ErrSection:

    Dim i&, iMaxAllowed&, strChk$, str67$
    Dim BidAskDepth As cBidAskDepth

    iMaxAllowed = MaxSymbolsAllowed(True)
    strSymbol = RollSymbolForDate(Trim(UCase(strSymbol)))
       
    ' new method
    If Len(strSymbol) > 0 Then
        Set BidAskDepth = m.MktDepthList(strSymbol)
        If BidAskDepth Is Nothing Then
            If bCreate Then
                ' TLB 3/15/2011: if already at the max, look for an "old" contract (i.e. non-67)
                ' to replace (which is at least better than to simply not allow this one)
                If m.MktDepthList.Count >= iMaxAllowed Then
                    For i = 1 To m.MktDepthList.Count
                        strChk = m.MktDepthList.Key(i)
                        If SecurityType(strChk) = "F" And strChk <> strSymbol Then
                            str67 = Parse(strChk, "-", 1) & "-067"
                            If strChk <> RollSymbolForDate(str67) Then
                                RemoveMarketDepthSymbol strChk
                                If m.MktDepthList.Count < iMaxAllowed Then
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                
                'limit number of symbols to max allowed
                If m.MktDepthList.Count < iMaxAllowed Then
                    Set BidAskDepth = New cBidAskDepth
' ??? TLB do we need to set dMinMove here (for the new cBidAskDepth object) ???

                    m.MktDepthList.Add BidAskDepth, strSymbol
                    ' add new symbol to stream
                    If m.RT.Active Then
                        m.RT.CreateMessage m.strServerNode, 50, strSymbol
                        If InStr(strSymbol, "-") > 0 Then
                            ''If Not FileExist(App.Path & "\SkipBrokerPrices.flg") Then
                            If FileExist(App.Path & "\UseBrokerPrices.flg") Then
                                If Not m.BrokerSubscriptionList.BinarySearch(strSymbol, i) Then
                                    m.BrokerSubscriptionList.Add strSymbol, i
                                    ResyncSubscriptionList m.nBrokerFeed
                                End If
                            End If
                        End If
                    End If
                    BidAskDepth.UsageCounter = BidAskDepth.UsageCounter + 1
                    If (m.nBrokerFeed = eTT_AccountType_TransAct) And (Not g.Transact Is Nothing) Then
                        g.Transact.ClearPrevPrices
                    End If
                End If
            End If
        Else
            BidAskDepth.UsageCounter = BidAskDepth.UsageCounter + 1
        End If
        Set AddMarketDepthSymbol = BidAskDepth
        Set BidAskDepth = Nothing
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.AddMarketDepthSymbol", eGDRaiseError_Raise
    Resume ErrExit
End Function

Public Sub RemoveMarketDepthSymbol(ByVal strSymbol$)
On Error GoTo ErrSection:

    Dim i&, BidAskDepth As cBidAskDepth
    
    If Len(strSymbol) = 0 Then Exit Sub
    strSymbol = RollSymbolForDate(strSymbol)
    
    Set BidAskDepth = m.MktDepthList(strSymbol)
    If Not BidAskDepth Is Nothing Then
        BidAskDepth.UsageCounter = BidAskDepth.UsageCounter - 1
        If BidAskDepth.UsageCounter = 0 Then
            m.MktDepthList.Remove strSymbol
            ' delete old symbol from stream
            If m.RT.Active Then
                m.RT.CreateMessage m.strServerNode, 51, strSymbol
                If m.BrokerSubscriptionList.BinarySearch(strSymbol, i) Then
                    m.BrokerSubscriptionList.Remove i
                    ResyncSubscriptionList m.nBrokerFeed
                End If
            End If
        End If
        Set BidAskDepth = Nothing
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.RemoveMarketDepthSymbol", eGDRaiseError_Raise
    Resume ErrExit
End Sub

Private Function DebugKey() As String
    'This version will wait for a message #40 with the IP-Address<TAB>Security key.
    DebugKey = "10.1.230.21" & vbTab & Chr(0) & "ABCDEFGHIJKLMNOP"
End Function

' convert "binary" exchange strings -- e.g. ascii 13 + "S" becomes "[0D]S"
Private Function EncodeExchange(ByVal strExchange As String) As String
    
    Dim strText$, i&
    
    ' if exchange is hex 0D-0F (decimal 13-15) then change to "US"
    'i = Asc(strExchange)
    'If i >= 13 And i <= 15 And Len(strExchange) = 1 Then
    If 0 Then
        strText = "US"
    Else
        ' check each character for ascii values < 32
        For i = 1 To Len(strExchange)
            If Asc(Mid(strExchange, i, 1)) >= 32 Then
                strText = strText & Mid(strExchange, i, 1)
            Else
                strText = strText & "[" & Right("0" & Hex(Asc(Mid(strExchange, i, 1))), 2) & "]"
            End If
        Next
    End If
    EncodeExchange = strText
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.EncodeExchange", eGDRaiseError_Raise
    Resume ErrExit
End Function

' this needs to be called both from RefreshSymbols and from cSymbolInfo.RefreshEOD
Public Function GetSymbolOffsetForGenesisRT(ByVal strSymbol$, Optional ByVal bAddNowToGenesisRT As Boolean = False) As String
On Error GoTo ErrSection:

    Dim nGmtOffset&, nSymbolID&, strReturn$
    Dim TimeInfo As TradeTimeInfo

    If Len(strSymbol) > 0 Then
        ' TLB 7/7/2009: pass GMT offset to GenesisRT
        ' (so we no longer have to rely on manually maintaining the extra info multiple times a year)
        nGmtOffset = 0
        nSymbolID = GetSymbolID(Parse(strSymbol, " ", 1))
        If nSymbolID > 0 Then
            If SU_GetTimeInfo(g.SU, nSymbolID, Date, TimeInfo) <> 0 Then
                If TimeInfo.cFeedTime <> "N" Then
                    nGmtOffset = TimeInfo.iLocalToGmtOffset
                    DebugLog "GMT offset for " & GetSymbol(nSymbolID) & " = " & Str(nGmtOffset)
                End If
            End If
        End If
        strReturn = strSymbol & vbTab & Str(nGmtOffset)
    
        If bAddNowToGenesisRT Then
            If SalmonIsRunning Then
                ' but ONLY if not already in the stream (seems to mess it up otherwise)
                If Not m.UniqueSymbols.BinarySearch(strSymbol) Then
                    m.RT.CreateMessage m.strServerNode, 15, strReturn, , True
                    If IsIDE Then
                        DebugLog strSymbol & " GMT offset added to GenesisRT now"
                    End If
                End If
            End If
        End If
    End If
    
    GetSymbolOffsetForGenesisRT = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.GetSymbolOffsetForGenesisRT"
    Resume ErrExit
End Function


' RefreshSymbols (iRefreshMode):
' - iRefreshMode: 0 = just new symbols, 1 = all symbols, 2 = just non-streaming symbols
' - if dirty flag, rebuild unique symbol list
' - if refreshing all, clear previous list (so all symbols are considered new)
' - get list of new symbols (compared to previous list)
' - if any new symbols (or removed symbols):
'   - need to pass all symbols to RT client
'   - need to send request for all symbols to be added to RT stream
'   - request today's data for just the new symbols
' - if RT:
'   - send RT requests (skip REX request if no IP address yet)
'   - pass symbols to RT client while waiting
'   - pass IP address to RT client
' - send REX request (if not done yet)
' - if distribution file:
'   - distribute data
'   - flag tick buffers to reload for new symbols
' - if refresh all, verify that most recent daily download has been done
Public Function RefreshSymbolList(Optional ByVal iRefreshMode As Integer = 0, _
                        Optional ByVal bAfterDailyDownload As Boolean = False, _
                        Optional ByVal bCurrentSessionUpdate As Boolean = False) As Boolean
On Error GoTo ErrSection:
    
    Dim i&, iPos&, iSymbol&
    Dim lNumSeconds&, dNyDate#, dMult#, nCrossoverTime&, nGmtOffset&, nStartTime&, nSymbolID&
    Dim strSymbol$, strSecType$, strText$, strFeed$, strFeedSymbol$, strFeedExchange$, strDates$
    Dim strFile$, strTemp$, strSynthetic$, strStreamed$
    Dim bDistData As Boolean, bBrokerRT As Boolean, bOnlyNewFromOptNav As Boolean
    Dim bStartTheStream As Boolean, bWasStreamingWhenStarted As Boolean
    Dim aFiles As New cGdArray
    Dim aTempSymbols As cGdArray
    Dim aNewSymbols As New cGdArray, aRemoveSymbols As New cGdArray
    Dim aCompSymbols As New cGdArray, aCompExchanges As New cGdArray
    Dim aRex As New cGdArray, aStream As New cGdArray, aRtClient As New cGdArray ' string arrays for the requests
    Dim Bars As New cGdBars
    
    Static aPrevSymbols As New cGdArray
    Static bInProgress As Boolean
                    
    If Me.Active Then bWasStreamingWhenStarted = True
                    
    ' with salmon, we no longer want to do a "quote board refresh" after streaming has started
    If SalmonIsRunning And Len(m.strIPandKey) > 0 Then
        If iRefreshMode <> 0 And Not bAfterDailyDownload And Not bCurrentSessionUpdate Then
            Exit Function
        End If
    End If
                    
    ' set flag so if process is busy now, total refresh will happen next time this is able to continue
    If Abs(iRefreshMode) = 1 Then m.bRefreshAllSymbols = True
    
    ' return if busy or if nothing to refresh ...
    If bInProgress Then Exit Function
    If Not m.bRefreshAllSymbols And iRefreshMode = 0 And Not m.bDirtySymbolList Then Exit Function
    If Not bAfterDailyDownload And Not bCurrentSessionUpdate Then
        If ProcessIsBusy(True) Then Exit Function
    End If
    
    ' build unique list of symbols (get copy of symbol list, sort, remove duplicates and empty strings)
    Set m.UniqueSymbols = m.TickBufferSymbols.MakeCopy
    m.UniqueSymbols.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
    
    ' TLB 2/19/07: add recent symbols (e.g. used in past week)
    Set aTempSymbols = m.UniqueSymbols.MakeCopy
    For i = 0 To m.RecentSymbols.Size - 1
        strText = Parse(m.RecentSymbols(i), vbTab, 1)
        If g.RealTime.SalmonIsRunning Then
            ' TLB 4/26/2010: for salmon, ignore any symbols not used today
            If Val(Parse(m.RecentSymbols(i), vbTab, 2)) < Now - 0.5 Then
                strText = ""
            End If
        ElseIf SecurityType(strText) = "S" Then
            ' TLB 6/8/07: only include stocks used in the past day
            If Val(Parse(m.RecentSymbols(i), vbTab, 2)) < Now - 1 Then
                strText = ""
            End If
        End If
        If Len(strText) > 0 Then
            If Not m.UniqueSymbols.BinarySearch(strText, iPos) Then
                m.UniqueSymbols.Add strText, iPos
            End If
        End If
    Next
    If m.UniqueSymbols.Size >= MaxSymbolsAllowed Then
        ' if too many, try adding just for the last day (stocks for today's session)
        Set m.UniqueSymbols = aTempSymbols.MakeCopy
        For i = 0 To m.RecentSymbols.Size - 1
            strText = Parse(m.RecentSymbols(i), vbTab, 1)
            If SecurityType(strText) = "S" Then
                If Val(Parse(m.RecentSymbols(i), vbTab, 2)) < Now - 0.5 Then
                    strText = ""
                End If
            Else
                If Val(Parse(m.RecentSymbols(i), vbTab, 2)) < Now - 1 Then
                    strText = ""
                End If
            End If
            If Len(strText) > 0 Then
                If Not m.UniqueSymbols.BinarySearch(strText, iPos) Then
                    m.UniqueSymbols.Add strText, iPos
                End If
            End If
        Next
        If m.UniqueSymbols.Size >= MaxSymbolsAllowed Then
            ' if too many, then just forget adding them
            Set m.UniqueSymbols = aTempSymbols.MakeCopy
        End If
    End If
    
    ' TLB 4/26/2010: add the pit contract for any non-pit contracts in list
    ' (since we now want the settles coming from the pit symbols)
    ' 4/29/2010: and also add the base forex (for forex profit calcs)
    For iSymbol = m.UniqueSymbols.Size - 1 To 0 Step -1
        strText = ""
        strSymbol = m.UniqueSymbols(iSymbol)
        Select Case SecurityType(strSymbol, True)
        Case "F"
            strText = GetSettleSymbol(strSymbol)
        Case "I"
            If IsForex(strSymbol) Then
                If Not g.Profit Is Nothing Then
                    strText = g.Profit.BaseForex(strSymbol)
                End If
            End If
        End Select
        If Len(strText) > 0 And strText <> strSymbol Then
            m.UniqueSymbols.Add strText
        End If
    Next
    m.UniqueSymbols.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
    m.bDirtySymbolList = False
       
    ' see what new symbols there are to add (since last time)
    If m.OptNavHistSymbols.Size > 0 Then
        bOnlyNewFromOptNav = True
    Else
        bOnlyNewFromOptNav = False
    End If
    If m.bRefreshAllSymbols Then aPrevSymbols.Size = 0
    For iSymbol = 0 To m.UniqueSymbols.Size - 1
        strSymbol = m.UniqueSymbols(iSymbol)
        If Not aPrevSymbols.BinarySearch(strSymbol) Then
            aNewSymbols.Add strSymbol
            ' check if all the new symbols are from OptNav (i.e. OHLC data already distributed)
            If bOnlyNewFromOptNav Then
                If Not m.OptNavHistSymbols.BinarySearch(strSymbol) Then
                    bOnlyNewFromOptNav = False
                End If
            End If
        ElseIf Abs(iRefreshMode) = 2 Then
            ' add symbols not being streamed
            If Not m.RtSymbolInfo.BinarySearch(strSymbol & "|", iPos, eGdSort_MatchUsingSearchStringLength Or eGdSort_IgnoreCase) Then
                aNewSymbols.Add strSymbol
                bOnlyNewFromOptNav = False
            End If
        End If
    Next
    
    ' if no symbols have been added, then nothing to do
    If aNewSymbols.Size = 0 And Not bCurrentSessionUpdate Then Exit Function ' And aPrevSymbols.Size = m.UniqueSymbols.Size Then Exit Sub
    
    ' for real-time feeds, see what old symbols there are to remove
    If Active Then ' And Not ActiveRTG Then
        For iSymbol = 0 To aPrevSymbols.Size - 1
            strSymbol = aPrevSymbols(iSymbol)
            If Not m.UniqueSymbols.BinarySearch(strSymbol) Then
                aRemoveSymbols.Add strSymbol
                i = m.LastQuote.Index(strSymbol)
                If i > 0 Then m.LastQuote.Remove strSymbol
                If SalmonIsRunning Then
                    SymbolInfo strSymbol, True
                    If frmTest.Visible Then
                        frmTest.AddList "Removed: " & strSymbol
                    End If
                End If
            End If
        Next
    End If
    
    Set aPrevSymbols = m.UniqueSymbols.MakeCopy
    
    bInProgress = True
       
    ' find out what date it is in NY right now
    dNyDate = ConvertTimeZone(Now)
    dNyDate = Int(dNyDate - 1 / 24#) '(adjust by 1 hour just to make sure)
    ' don't need to request today's session if already did daily download for today
    If dNyDate <= LastDailyDownload Then
        dNyDate = LastDailyDownload + 1
    End If
    Do While Not IsWeekday(dNyDate)
        dNyDate = dNyDate + 1
    Loop
    ' ask for today and tomorrow (next weekday)
    i = dNyDate + 1
    Do While Not IsWeekday(i)
        i = i + 1
    Loop
    strDates = "@" & Format(dNyDate, "YYYYMMDD") & "-" & Format(i, "YYYYMMDD")
            
    If UCase(m.strServerNode) = "ESIGNALRT" Then
        strFeed = "S"
    Else
        strFeed = "C"
    End If
            
    ' build all the requests (Rex, Stream, RtClient) into the string arrays for now
    frmOnlineBroker.GetLastKnownBrokerSymbols
    ''m.BrokerSubscriptionList.Clear
    For iSymbol = 0 To m.UniqueSymbols.Size - 1
        strSymbol = m.UniqueSymbols(iSymbol)
        bBrokerRT = False
        
        If m.RT.Active Then
            If ActiveRTG Then
                'add real-time streaming requests:  "+GenSymbol=SP-067;SP-200503"  (GenSymbol;Underlying)
                strText = RollSymbolForDate(strSymbol)
                If strText = strSymbol Then strText = ""
                strText = "+GenSymbol=" & strSymbol & ";" & strText
                
                ' see if allowed realtime access due to broker
                If IsBrokerSymbol(strSymbol, i) Then
                    strText = strText & ";0;" & Parse(m.BrokerRtSymbols(i), vbTab, 2)
                    bBrokerRT = True
                    ''m.BrokerSubscriptionList.Add strSymbol
                End If
                
                aStream.Add strText
                'aRtClient(0) = " "
                
                strText = GetSymbolOffsetForGenesisRT(strSymbol)
                aRtClient.Add strText
            
            ElseIf TranslateSymbol(strSymbol, strFeed, strSecType, strFeedSymbol, strFeedExchange, _
                                dMult, nCrossoverTime, nGmtOffset, False, aCompSymbols, aCompExchanges) Then
                
                frmQuotes.FeedSymbol strSymbol, strFeedSymbol
                    
                ' convert from MinSinceMid to HHMM (NY time for American, GMT for foreign)
                If nCrossoverTime > 0 Then nCrossoverTime = Int(nCrossoverTime / 60) * 100 + (nCrossoverTime Mod 60)
                    
                ' for RT client need to add:
                ' - exchange offset from GMT in minutes (+/-)
                ' - for day-session synthetic symbols, add session starting time as HHMM (else leave blank)
                strText = ""
                If strSecType = "F" Then
                    SetBarProperties Bars, strSymbol
                    If Bars.Prop(eBARS_Session) = Asc("S") Then
                        ' convert StartTime from MinSinceMid to HHMM
                        nStartTime = Bars.Prop(eBARS_StartTime)
                        nStartTime = Int(nStartTime / 60) * 100 + (nStartTime Mod 60)
                        strText = Str(nStartTime)
                    End If
                End If
                ' RT client:  "WP-067  F   SP!M4      1625    1   SPM4,SP&M4  , 0"
                strText = strSymbol & vbTab & strSecType & vbTab & strFeedSymbol _
                    & vbTab & strFeedExchange & vbTab & Str(nCrossoverTime) & vbTab & Str(dMult) _
                    & vbTab & aCompSymbols.JoinFields(",") & vbTab & aCompExchanges.JoinFields(",") _
                    & vbTab & Str(nGmtOffset) & vbTab & strText
                aRtClient.Add strText
            End If
        End If
        
        ' TLB 11/23/05: don't do Rex request for continuous contract (since individual contract
        ' is now always being asked for and the continuous is now redundant)
        If InStr(strSymbol, "-0") > 0 Then
            If InStr(strSymbol, "-05") > 0 Or InStr(strSymbol, "-06") > 0 _
                Or InStr(strSymbol, "-08") > 0 Or InStr(strSymbol, "-09") > 0 Then
                    If Not InStr(strSymbol, "-099") Then
                        strSymbol = ""
                    End If
            End If
        End If
        ' do Rex request only for the new symbols
        If aNewSymbols.BinarySearch(strSymbol) Then
            ' if they are authorized for ticks, ask for both, else just ask for end-of-day
            strSecType = SecurityType(strSymbol, True)
            If InStr(g.strAuthorizationString, "," & Left(strSecType, 1) & "T,") > 0 And Right(strSecType, 1) <> "O" Then
                strText = ";B;I;"
            Else
                strText = ";E;I;"
            End If
            ' TLB 12/14/2007: convert from synthetic for Rex requests
            strSynthetic = ""
            If Left(strSecType, 1) = "F" Then
                strTemp = ConvertSynthetic(strSymbol, False)
                If Len(strTemp) > 0 Then
                    strSynthetic = strSymbol
                    strSymbol = strTemp
                End If
            End If
            ' delay override (if hooked up to realtime broker prices)
            If bBrokerRT Then
                strTemp = ";0"  ' X:<12>
            Else
                strTemp = ";-1" ' X:<12>
            End If
            ' add the Rex request
            strText = strDates & strText & strSecType & ";" & strSymbol & strTemp
            aRex.Add strText
            If Len(strSynthetic) > 0 Then
                ' if this is a synthetic, then also ask for just end-of-day for the synthetic
                ' as well as asking for the "base symbol"
                aRex.Add strDates & ";E;I;" & strSecType & ";" & strSynthetic & strTemp
            End If
        End If
    Next
    Set Bars = Nothing
    
    ' TLB 8/10/2010: should just exit out if "status" has changed since this routine started
    If g.bUnloading Or (Me.Active <> bWasStreamingWhenStarted) Then
        GoTo ErrExit ' Exit Function
    End If
    
If SalmonIsRunning Then
    aRex.Size = 0
    If Len(m.strIPandKey) > 0 Then
        ''aStream.Size = 0
    End If
End If
    
    aRex.Sort eGdSort_IgnoreCase Or eGdSort_DeleteDuplicates
    aStream.Sort eGdSort_IgnoreCase Or eGdSort_DeleteDuplicates
    aRtClient.Sort eGdSort_IgnoreCase Or eGdSort_DeleteDuplicates
    ''m.BrokerSubscriptionList.Sort eGdSort_IgnoreCase Or eGdSort_DeleteDuplicates
    
'aStream.ToFile "c:\stream.txt"
'aRex.ToFile "c:\rex.txt"
'aRtClient.ToFile "c:\rt.txt"
'Exit Sub
    
    ' for a quote board refresh, do a one-time refresh of Prediction Labs data
    If m.bRefreshAllSymbols And Not g.RealTime.Active Then
        frmMain.tmrPredLabs.Enabled = True
    End If

    If Not frmStatus.IsBusy And Not bAfterDailyDownload And Not bCurrentSessionUpdate And g.nReplaySession = 0 Then
        frmStatus.IsBusy = True
        frmStatus.Status = eStatus_Initialized
        frmStatus.ShowDetails False
    End If

    If bCurrentSessionUpdate Then
        If SalmonIsRunning Then
            aRex.Clear ' when salmon is running, a CSU should not do Rex requests for symbols
        End If
        aStream.Clear
        aRtClient.Clear
        aRemoveSymbols.Clear
        
        ' prepend things to request array
        ' - add CRC's for all TODAY_* files
        aFiles.GetMatchingFiles App.Path & "\FTP\Backup\Today_*.gzp"
        For i = 0 To aFiles.Size - 1
            aRex.Add "+CRC:" & FileCrcString(aFiles(i)), 0
        Next
        ' - add exclude/include modules
        strText = Trim(GetIniFileProperty("DownloadInclude", "", "General", g.strIniFile))
        If Len(strText) > 0 Then
            aRex.Add "+INCLUDE:" & strText, 0
        End If
        strText = Trim(GetIniFileProperty("DownloadExclude", "", "General", g.strIniFile))
        If Len(strText) > 0 Then
            aRex.Add "+EXCLUDE:" & strText, 0
        End If
        ' - intraday request
        aRex.Add "+DATE:" & ConvertDate(0), 0
        aRex.Add "%INTRADAY DOWNLOAD", 0
    End If


    ' RealTime Client: pass new symbols if connection already open
    If aRtClient.Size > 0 And Len(m.strIPandKey) > 0 Then
        ' pass all symbols over to RT client
        strTemp = aRtClient.JoinFields(vbCrLf)
        If ActiveRTG Then
            m.RT.CreateMessage m.strServerNode, 15, strTemp, , False
        Else
            For i = 0 To aRtClient.Size - 1
                m.RT.CreateMessage m.strServerNode, 15, aRtClient(i), , False
            Next
        End If
        ' and start streaming
        m.RT.CreateMessage m.strServerNode, 21, Str(m.iRtInterval), , False
        m.RT.SendNow
        aRtClient.Size = 0
    End If

    ' Streaming requests:
    If aStream.Size > 0 Then
        ' if no IPandKey at this point then we will be needing to start the stream
        If Len(m.strIPandKey) = 0 Then
            bStartTheStream = True
            If SalmonIsRunning Then
                ' when starting the stream with Salmon
                StatusMsg "Connecting to realtime streaming ...", vbRed
                SalmonInitNewStream True
                If Len(m.strIPandKey) > 0 Then
                    frmStatus.Status = eStatus_Completed
                End If
            End If
            ' but we also need to wait until GenesisRT is up and running and receiving AppMails
            dMult = gdTickCount
            Do While m.RT.CreateMessage(m.strServerNode, 21, "0") = 0
                Sleep 0.1
                If gdTickCount > dMult + 5000 Then Exit Do ' 5 second timeout
            Loop
        End If
        ' store tab-limited list of symbols to send to GenesisRT
        strStreamed = aStream.JoinFields(vbTab)
        If bOnlyNewFromOptNav Then
            aStream.Size = 0
            aRex.Size = 0
        Else
            m.OptNavHistSymbols.Size = 0
            
            If Not SalmonIsRunning Then
                ' append streaming list from GenesisRT
                aStream.AppendFromArray m.OptNavStreamingList
                aStream.Add "%RT Symbols", 0
            
                If g.nReplaySession = 0 Then
                    If Len(m.strIPandKey) > 0 Then
                        ' Change the timeout of GClient based on the number of symbols in the list
                        ' (do 1 second per symbol + a 30 second buffer)
                        lNumSeconds = aRex.Size + 30
                        aStream.Add "+StreamerIP=" & Parse(m.strIPandKey, vbTab, 1)
                        aStream.Add "+AccessKey=" & Parse(m.strIPandKey, vbTab, 2)
                        
                        ' if connection already open, then include the Rex requests in now
                        If aRex.Size > 0 Then
                            aStream.AppendFromArray aRex
                            aRex.Size = 0
                            bDistData = True
                            If Not bAfterDailyDownload Then
                                frmStatus.SetTitle "Retrieving Data"
                            End If
                        Else
                            frmStatus.SetTitle "Updating Data Stream"
                        End If
                        frmStatus.AddDetail "Updating RealTime stream ..."
                    Else
                        lNumSeconds = 90 'allow 90 seconds for now in case need to go to backup splitters
                        frmStatus.SetTitle "Opening Data Stream"
                        frmStatus.Status = eStatus_Initialized
                        frmStatus.AddDetail "Requesting RealTime stream ..."
                    End If
                End If
                
                ' Make the ftp request, and get IP info returned
                strFile = AddSlash(App.Path) & "FTP\request.rtg"
                KillFile strFile
                Set MsgForm = frmStatus
                If g.nReplaySession > 0 Then
                    ' if replaying stream from a file (subtract 1/10 second due to rounding issues)
                    ''Sleep 5 '(wait for RT client to start up)
                    'FileFromString strFile, "STREAMERIP=" & App.Path & "\RTS\" & Format(g.nReplaySession, "YYYYMMDD") & ".rts" _
                        & vbCrLf & "ACCESSKEY=" & Str(ConvertTimeZone(m.dFeedTime - 0.000001, "NY", "GMT")), True
                    m.strIPandKey = App.Path & "\RTS\" & Format(g.nReplaySession, "YYYYMMDD") & ".rts" _
                        & vbTab & Str(ConvertTimeZone(m.dFeedTime - 0.000001, "NY", "GMT"))
                    frmStatus.Status = eStatus_Completed
                    ' set speed
                    m.RT.CreateMessage m.strServerNode, 42, Str(ReplaySpeed)
                ElseIf FtpRequest(aStream, , lNumSeconds) Then
                    aStream.FromFile strFile
                    If Len(m.strIPandKey) = 0 And aStream.Size > 0 Then
                        ' get Splitter IP and Key from file
                        strText = ""
                        For i = 0 To aStream.Size - 1
                            iPos = InStr(aStream(i), "=")
                            If iPos > 1 Then
                                Select Case Trim(UCase(Left(aStream(i), iPos - 1)))
                                Case "STREAMERIP"
                                    m.strIPandKey = Trim(Mid(aStream(i), iPos + 1))
                                Case "ACCESSKEY"
                                    strText = Trim(Mid(aStream(i), iPos + 1))
                                End Select
                            End If
                        Next
                        If Len(strText) > 0 Then
                            m.strIPandKey = m.strIPandKey & vbTab & strText
                        End If
                    End If
                Else
                    bDistData = False
                End If
                Set MsgForm = Nothing
            End If
            
            aStream.Size = 0
            If frmStatus.Status = eStatus_Aborted Or frmStatus.Status = eStatus_Aborting Then
                Init False, "Aborting startup"
                GoTo ErrExit
            ElseIf bAfterDailyDownload Then
                frmStatus.Status = eStatus_Completed
            End If
            
            If FileExist(App.Path & "\RTS\Record.flg") And g.nReplaySession = 0 And bStartTheStream Then
                ' record the stream to a file (bump to next day if after 5:30pm NY)
                dNyDate = ConvertTimeZone(Now, "", "NY")
                If dNyDate - Int(dNyDate) >= 0.73 Then
                    dNyDate = Int(dNyDate) + 1
                Else
                    dNyDate = Int(dNyDate)
                End If
                MakeDir App.Path & "\RTS"
                strFile = App.Path & "\RTS\" & Format(dNyDate, "YYYYMMDD") & ".RTS"
                ' if file already exists and after 8am, rename it (if backup does not yet exist)
                If FileExist(strFile) And Hour(Now) > 7 Then
                    strTemp = ReplaceFileExt(strFile, "BAK")
                    If Not FileExist(strTemp) Then
                        Name strFile As strTemp
                    End If
                End If
                m.RT.CreateMessage m.strServerNode, 44, strFile
            End If
            
            ' RealTime Client: pass new symbols now if not already passed
            If aRtClient.Size > 0 Then
                ' pass all symbols over to RT client
                strTemp = aRtClient.JoinFields(vbCrLf)
                If ActiveRTG Then
                    m.RT.CreateMessage m.strServerNode, 15, strTemp, , False
                Else
                    For i = 0 To aRtClient.Size - 1
                        m.RT.CreateMessage m.strServerNode, 15, aRtClient(i), , False
                    Next
                End If
                ' and start streaming
                m.RT.CreateMessage m.strServerNode, 21, Str(m.iRtInterval), , False
                m.RT.SendNow
                aRtClient.Size = 0
            End If
            
            If bStartTheStream Then
                If Len(m.strIPandKey) > 0 Then
                    ' must send the SalmonPort to GenesisRT before sending the IP-Key
                    If m.nSalmonPort > 0 And SalmonIsRunning Then
                        m.RT.CreateMessage m.strServerNode, 110, Str(m.nSalmonPort)
                    End If
                    ' send the IP, Key, etc. over to GenesisRT
                    DebugLog "Splitter IP and Key:" & vbTab & m.strIPandKey
                    strText = m.strIPandKey & vbTab & Str(RI_GetDataServiceID) & vbTab _
                                & g.strAuthorizationString & vbTab
                    If SalmonIsRunning Then
                        strText = strText & Trim(FileToString(App.Path & "\Provided\Salmon.IP", , True)) ' "~10.1.191.2" 'Salmon IP addresses
                    Else
                        strText = strText & Trim(FileToString(App.Path & "\Provided\DataSrv.IP", , True)) ' "~10.1.191.2" 'Salmon IP addresses
                    End If
                    DebugLog strText
                    m.RT.CreateMessage m.strServerNode, 40, strText
                    ' and pass list of broker override symbols
                    strText = m.BrokerRtSymbols.JoinFields(Chr(27))
                    m.RT.CreateMessage m.strServerNode, 160, strText ' should be obsolete for Salmon as of 9/10/2010
                ElseIf frmStatus.Status <> eStatus_Aborted And frmStatus.Status <> eStatus_Aborting Then
                    frmStatus.SetTitle "Opening Data Stream"
                    frmStatus.Status = eStatus_Error
                    frmStatus.AddDetail "ERROR opening RealTime stream"
                End If
                SalmonStart True
            ElseIf bAfterDailyDownload And SalmonIsRunning Then
                SalmonStart
            End If
             
            aStream.Size = 0
        End If
        
        ' update GenesisRT with TradeNav's streaming list
        'If Not SalmonIsRunning Then
            m.RT.CreateMessage m.strServerNode, 150, strStreamed ' should be obsolete for Salmon as of 9/10/2010
        'End If
    End If

    ' RealTime Client: pass new symbols now if not already passed
    If aRtClient.Size > 0 Then
        ' if not Genesis, wait for client to get started
        If Not ActiveRTG Then
            dMult = gdTickCount
            Do While m.nServerStatus = 0
                DoEvents
                If gdTickCount > dMult + 20000 Then Exit Do '(timed out)
            Loop
        End If
        
        ' pass all symbols over to RT client
        If ActiveRTG Then
            m.RT.CreateMessage m.strServerNode, 15, aRtClient.JoinFields(vbCrLf), , False
        Else
            For i = 0 To aRtClient.Size - 1
                m.RT.CreateMessage m.strServerNode, 15, aRtClient(i), , False
            Next
        End If
        ' and start streaming
        m.RT.CreateMessage m.strServerNode, 21, Str(m.iRtInterval), , False
        m.RT.SendNow
        aRtClient.Size = 0
    End If

    ' RealTime Client (except Genesis RT): pass symbols to remove
    If aRemoveSymbols.Size > 0 Then
        For i = 0 To aRemoveSymbols.Size - 1
            m.RT.CreateMessage m.strServerNode, 16, aRemoveSymbols(i), , False
        Next
        m.RT.SendNow
    End If
    
If IsIDE Then
'aRex.Clear
End If
    
    ' REX requests:
    If g.nReplaySession > 0 Then
        bDistData = False
        aRex.Size = 0
    ElseIf aRex.Size > 0 And (frmStatus.Status = eStatus_Completed Or frmStatus.Status = eStatus_Initialized) Then
        If m.bRefreshAllSymbols Or iRefreshMode = 2 Then
            g.dNextQuoteBoardRefresh = 0
            g.dLastQuoteBoardRefresh = Now
        End If

        If Not bCurrentSessionUpdate Then
            frmStatus.Status = eStatus_Initialized
            frmStatus.SetTitle "Retrieving Data"
            frmStatus.AddDetail "Refreshing Quote Board ..."
        End If

        ' Change the timeout of GClient based on the number of symbols in the list
        ' (do 1 second per symbol + a 30 second buffer)
        lNumSeconds = aRex.Size + 30
        ' Make the ftp request
        Set MsgForm = frmStatus
        If FtpRequest(aRex, , lNumSeconds) Then
            bDistData = True
        Else
            KillFile AddSlash(App.Path) & "FTP\Data.dat"
            ' TLB 5/22/2014: however, if we do have "current session" data, we can at least distribute that
            If FileExist(App.Path & "\FTP\Today_*.gzp") Then
                bDistData = True
            End If
        End If
        Set MsgForm = Nothing
    End If
    
    ' reset the GenesisRT priority to normal
    SetRTPriority False
    
    ' log QB files?
    If Abs(iRefreshMode) = 1 Then
        If FileExist(App.Path & "\ftp\QBLogs\QB.log") Then
            KillFile App.Path & "\ftp\*.gzp"
            strFile = App.Path & "\ftp\QBLogs\" & Format(Now, "yyyymmdd-hhnnss") & ".zip"
            ZipExecute "c", strFile, App.Path & "\ftp\", "*.*"
        End If
    End If
    
    ' Distribute data from REX
    If g.nReplaySession > 0 Then
        ' (TLB 12/17/2009: for Streaming Replay, we need to set reload flag for the new symbols)
        ' set flag for just the new symbols so the next call to UpdateBars will cause them to reload
        For i = 1 To m.TickBufferSymbols.Size - 1
            strSymbol = m.TickBufferSymbols(i)
            If aNewSymbols.BinarySearch(strSymbol) Then
                Set Bars = m.TickBuffers(i)
                If Not Bars Is Nothing Then
                    Bars.Prop(eBARS_CustomString) = "<RELOAD>"
                End If
                ClearCache strSymbol
            End If
        Next
    ElseIf bDistData And frmStatus.Status = eStatus_Completed Then
            
        ' reset GenTick export list
        If FileExist(App.Path & "\GenTick.EXP") Then
            FileFromString App.Path & "\Data\TickDist.Lst", " ", True
        End If
            
        ' Distribute the data (if data to distribute)
        If Not DistributeData("Distributing Data") Then
            'State.eCurStatus = 600
'            frmStatus.Status = eStatus_Error
            If Not bAfterDailyDownload Then
                frmStatus.AddDetail "ERROR downloading data"
            End If
        End If
    
        If frmStatus.Status = eStatus_Completed Then
            
            frmStatus.AddDetail "Final Updating"
            DM_DistribData ""
            frmStatus.Status = eStatus_Running
            
            If Not m.bRefreshAllSymbols And Not bCurrentSessionUpdate Then
                ' set flag for just the new symbols so the next call to UpdateBars will cause them to reload
                For i = 1 To m.TickBufferSymbols.Size - 1
                    strSymbol = m.TickBufferSymbols(i)
                    If aNewSymbols.BinarySearch(strSymbol) Then
                        Set Bars = m.TickBuffers(i)
                        If Not Bars Is Nothing Then
                            Bars.Prop(eBARS_CustomString) = "<RELOAD>"
                        End If
                        ClearCache strSymbol
                    End If
                Next
            End If
            If m.bRefreshAllSymbols Or bAfterDailyDownload Then
                ClearCache
            End If
            
            UpdateGenTick
           
        End If
        
    End If
    
    If m.bRefreshAllSymbols Or bCurrentSessionUpdate Or bAfterDailyDownload Then
        ' refresh data on all forms
        If frmStatus.Visible And g.nReplaySession = 0 Then
            frmStatus.AddDetail "Reloading Data"
            frmStatus.UpdateProgress "Reloading Data"
        End If
        If Not (SalmonIsRunning And bCurrentSessionUpdate) Then
            RefreshAllFormData bAfterDailyDownload
        End If
        
        CalcNextQuoteRefresh
        m.bRefreshAllSymbols = False
    ElseIf iRefreshMode = 2 Then
        CalcNextQuoteRefresh
    End If
    
    If frmStatus.Status = eStatus_Completed Or frmStatus.Status = eStatus_Running Then
        frmStatus.Status = eStatus_Completed
        If frmStatus.Visible And Not bAfterDailyDownload Then
            frmStatus.AddDetail "Finished"
        End If
    End If
    
    If bAfterDailyDownload Then
        ' tell RT client to delete old real-time data
        If FileExist(App.Path & "\DeleteOld.rtg") Then '(this may be causing problems -- being debugged)
            m.RT.CreateMessage m.strServerNode, 5, Str(LastDailyDownload)
        End If
    End If
    
' - if dirty flag, rebuild unique symbol list
' - if refreshing all, clear previous list (so all symbols are considered new)
' - get list of new symbols (compared to previous list)
' - if any new symbols (or removed symbols):
'   - need to pass all symbols to RT client
'   - need to send request for all symbols to be added to RT stream
'   - request today's data for just the new symbols
' - if RT:
'   - send RT requests (skip REX request if no IP address yet)
'   - pass symbols to RT client while waiting
'   - pass IP address to RT client
' - send REX request (if not done yet)
' - if distribution file:
'   - distribute data
'   - flag tick buffers to reload for new symbols
' - if refresh all, verify that most recent daily download has been done

    
    ' resync broker feed list
    ResyncSubscriptionList m.nBrokerFeed

    If ConnectionStatus = eGDConnectionStatus_Connecting Then
        ConnectionStatus = eGDConnectionStatus_Connected
    End If

    ' verify the broker positions and auto-trading status
    g.Broker.VerifyBrokerPositions
    g.OrderStrategies.ActivateExits

    ' TLB 5/22/2006: just for debugging purposes (this stops storing ticks in T&E after X ms delay)
    If ActiveRTG Then
        'strText = FileToString(App.Path & "\RtStorageCutoff.flg", , True)
        'If Len(strText) > 0 Then m.RT.CreateMessage m.strServerNode, 22, strText
    End If

    ' now bump up the GenesisRT priority
    ''SetRTPriority True

ErrExit:
    frmStatus.IsBusy = False
    bInProgress = False
    RefreshSymbolList = True
    Exit Function
    
ErrSection:
    bInProgress = False
    If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    frmStatus.IsBusy = False
    RaiseError "cRealTime.RefreshSymbolList", eGDRaiseError_Raise
    Resume ErrExit
End Function

' to refresh the data on all forms after a distribution or realtime on/off
Public Sub RefreshAllFormData(Optional ByVal bReloadHistory As Boolean = False)
On Error GoTo ErrSection:

    Dim iInterval&, i&
    Dim frm As Form, frms As New cForms
    
    If bReloadHistory And Not g.ChartPageCache Is Nothing Then
        ' clear the Chart Page cache (since older data corrections may have come in)
        g.ChartPageCache.Clear
    End If
    
    ' clear the Salmon cache (we need it rebuilt correctly)
    ClearCache

    ' if streaming replay but just pre-loading data (not yet caught up to start time),
    ' then don't want to reload now -- reloading will happen when get to the start time
    If g.nReplaySession > 0 Then
        If Not frmReplay.IsPlaying Then Exit Sub
    End If

    ' if real-time inactive, make sure all real-time timers are off
    If Not g.RealTime.Active Or g.bUnloading Then
        frmMain.tmrPredLabs.Enabled = False
        frmQuotes.tmrRealTime.Enabled = False
        frmTTSummary.tmrRealTime.Enabled = False
        frmOnlineBroker.tmrAutoTradeData.Enabled = False
        frmOptionChain.tmrRealTime.Enabled = False
        
        If FormIsLoaded("frmAlertsSetup") Then
            frmAlertsSetup.tmrRealTime.Enabled = False
            frmAlertsSetup.ClearUpdatedColors
        End If
        
        If FormIsLoaded("frmBrokerView") Then
            frmBrokerView.tmrRealTime.Enabled = False
        End If
    
        g.CattleBridge.StreamTimerEnabled = False
    End If
    If g.bUnloading Then Exit Sub

    ' get charts updated first
    If g.RealTime.Active Then ChartTimers = False
    If m.bReconnecting Then
        UpdateVisibleCharts eRedo1_Scrolled
    ElseIf g.RealTime.Active And Not bReloadHistory Then
        UpdateVisibleCharts eRedo7_ReloadRT
    Else
        UpdateVisibleCharts eRedo9_ReloadData
    End If
    If g.bUnloading Then Exit Sub
                    
#If 0 Then
    If bReloadHistory Then
        For i = 0 To Forms.Count - 1
            If g.bUnloading Then Exit Sub
            If i >= Forms.Count Then Exit For
            Set frm = Forms(i)
            If TypeOf frm Is frmTickDistribution Then
                'get all forms to release their market depth object so the data stream
                'can be closed and re-opened to get history data, this is necessary
                'when user changes the show as local time flag or on a re-connect
                frm.ReleaseMarketDepth
            End If
            Set frm = Nothing
        Next
        Sleep 2
    End If
   
    For i = 0 To Forms.Count - 1
        If g.bUnloading Then Exit Sub
        If i >= Forms.Count Then Exit For
        Set frm = Forms(i)
        If TypeOf frm Is frmTickDistribution Then
            frm.RefreshData
            If bReloadHistory Then
                'give the depth of market object a chance to fill its table
                'with history data before the next form try to grab it
                Sleep 3
            End If
        ElseIf TypeOf frm Is frmTimeSales Then
            frm.RefreshData
        ElseIf TypeOf frm Is frmTimeSalesAnalyzer Then
            frm.RefreshData
        ElseIf TypeOf frm Is frmPriceVol Then
            frm.RefreshData
        ElseIf TypeOf frm Is frmBidAskDir Then
            frm.RefreshData
        ElseIf TypeOf frm Is frmMarketProfile Then
            frm.RefreshData
        End If
        Set frm = Nothing
    Next
#Else
    If bReloadHistory Then
        frms.Init frmTickDistribution
        Do While Not g.bUnloading
            Set frm = frms.NextForm
            If frm Is Nothing Then Exit Do
            'get all forms to release their market depth object so the data stream
            'can be closed and re-opened to get history data, this is necessary
            'when user changes the show as local time flag or on a re-connect
            frm.ReleaseMarketDepth
            Set frm = Nothing
        Loop
        Sleep 1
    End If
   
    frms.Init
    Do While Not g.bUnloading
        Set frm = frms.NextForm
        If frm Is Nothing Then Exit Do
        
        If TypeOf frm Is frmTickDistribution Then
            frm.RefreshData
            If bReloadHistory Then
                'give the depth of market object a chance to fill its table
                'with history data before the next form try to grab it
                Sleep 3
            End If
        ElseIf TypeOf frm Is frmTimeSales Then
            frm.RefreshData
        ElseIf TypeOf frm Is frmTimeSalesAnalyzer Then
            frm.RefreshData
        ElseIf TypeOf frm Is frmPriceVol Then
            frm.RefreshData
        ElseIf TypeOf frm Is frmBidAskDir Then
            frm.RefreshData
        ElseIf TypeOf frm Is frmMarketProfile Then
            frm.RefreshData
        End If
        Set frm = Nothing
    Loop
#End If
                
    ' if real-time active, make sure all real-time timers are on
    If g.RealTime.Active Then
        iInterval = 250
        If m.UniqueSymbols.Size < 100 Then
            'iInterval = 125
        End If
'If IsIDE Then iInterval = 1000
    
        frmQuotes.tmrRealTime.Interval = iInterval
        frmQuotes.tmrRealTime.Enabled = True
        
        frmTTSummary.tmrRealTime.Interval = iInterval
        frmTTSummary.tmrRealTime.Enabled = True
        
        frmOnlineBroker.tmrAutoTradeData.Interval = iInterval
        frmOnlineBroker.tmrAutoTradeData.Enabled = True
    
        ' TLB 4/2/2008: no more realtime for the old option chain
        If 0 Then ' HasModule("OPTNAV") Then
            frmOptionChain.tmrRealTime.Interval = iInterval
            frmOptionChain.tmrRealTime.Enabled = True
        End If
    
        If FormIsLoaded("frmAlertsSetup") Then
            frmAlertsSetup.tmrRealTime.Interval = iInterval
            frmAlertsSetup.tmrRealTime.Enabled = True
        End If
        
        If FormIsLoaded("frmBrokerView") Then
            frmBrokerView.tmrRealTime.Interval = iInterval
            frmBrokerView.tmrRealTime.Enabled = True
        End If
    
        g.CattleBridge.StreamInterval = iInterval
        g.CattleBridge.StreamTimerEnabled = True
    
        frmMain.tmrPredLabs.Enabled = True
    End If

    ' lastly, update data on other forms (keep checking g.bUnloading
    ' in case flag gets set after one of the data refreshes)
    If Not g.bUnloading Then frmQuotes.TotalRefresh True
    If Not g.bUnloading Then frmTTSummary.RefreshData
    If Not g.bUnloading Then frmOptionChain.LoadData
    
    If Not g.bUnloading Then g.TradingItems.ReloadData
    If Not g.bUnloading Then g.TsoGroups.ReloadData
   
    ChartTimers = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.RefreshAllFormData", eGDRaiseError_Raise
    Resume ErrExit
End Sub

Public Property Get QuoteCount() As Double
    QuoteCount = m.dQuoteCount
End Property

' To restart the real-time data stream -- for timeout, can pass:
'  -1 = random # from 30-120,  >0 = specified timeout,  0 = skip message (just do it now)
Public Sub Reconnect(Optional ByVal nTimeoutSeconds As Long = -1, Optional ByVal bCompactTradeTracker As Boolean = False)
On Error GoTo ErrSection:

    Dim strMsg$, i&
    Dim strDB$, strBU1$, strBU2$, strTemp$

    If g.nReplaySession > 0 Then Exit Sub

    If nTimeoutSeconds < 0 Then
        ' pick random number (so everybody won't all be reconnecting at the same time when splitter goes down)
        nTimeoutSeconds = RandomNum(30, 120)
    End If

    If nTimeoutSeconds > 0 Then
        DebugLog "Streaming dropped: " & m.strSplitterMsg
        strMsg = Trim(m.strSplitterMsg)
        If Len(strMsg) = 0 Then strMsg = "The connection to the realtime |data stream has been dropped."
        strMsg = strMsg & "||Try to reconnect?"
        If InfBox(strMsg, "?", "+Reconnect|-Not Now", "Realtime Streaming", , nTimeoutSeconds) <> "R" Then
            m.dLastMsgTime = 0
            Init False, "Not reconnecting"
            Exit Sub
        End If
    End If
    
    ' log the reconnection
    If bCompactTradeTracker Then
        strMsg = "Reconnecting data stream (normal Sunday morning reconnect)."
    ElseIf Len(Trim(m.strSplitterMsg)) = 0 Then
        strMsg = "Reconnecting data stream."
    Else
        strMsg = Trim(m.strSplitterMsg)
    End If
    g.Broker.BrokerDebugAll strMsg
    
    m.bReconnecting = True
    frmMain.tbToolbar.Tools("ID_RealTime").Enabled = False
    Init False, "Reconnecting"
    Sleep 1
    'For i = 5 To 1 Step -1
    '    StatusMsg "Will reconnect to realtime data stream in " & Str(i) & " seconds ...", vbRed
    '    Sleep 1
    'Next
    'StatusMsg
    
    If bCompactTradeTracker Then
        strDB = App.Path + "\TradeTracker.MDB"
        If FileExist(strDB) And FileLength(App.Path & "\Provided\SkipCompact.flg") < 5 Then
            g.dbPaper.Close ' Must first close the database
            Set g.dbPaper = Nothing
            DoEvents ' TLB 9/12/2013: we think we may need this for the Compact to work correctly
            strBU1 = App.Path + "\TTBak1.mdb"
            strBU2 = App.Path + "\TTBak2.mdb"
            strTemp = App.Path + "\_Temp.mdb"
            KillFile strTemp
            On Error Resume Next 'GoTo CompactError
            DBEngine.CompactDatabase strDB, strTemp, , dbDecrypt
            If FileExist(strTemp) Then
                DebugLog "TradeTracker.MDB Compacted"
                Set g.dbPaper = OpenDatabase(strTemp, "")
                If Not g.dbPaper Is Nothing Then
                    If CheckTtIndexes("after compact") = True Then
                        g.dbPaper.Close
                        DoEvents
                
                        ' save backups
                        KillFile strBU2
                        If FileExist(strBU1) Then Name strBU1 As strBU2
                        KillFile strBU1
                        Name strDB As strBU1
                        Name strTemp As strDB
                    Else
                        g.dbPaper.Close
                    End If
                End If
            End If
            ' now re-open the compacted database
            Set g.dbPaper = OpenDatabase(strDB, "")
            
            ' 11/03/2014 DAJ: Once the database is reopened, need to send it to the appropriate
            ' DLL's as well otherwise they yield an "Object invalid or no longer set" error...
            mTradeTracker.SetDllBridgeDatabases
        End If
    End If
    
    frmMain.tbToolbar.Tools("ID_RealTime").Enabled = True
    Init True

ErrExit:
    m.bReconnecting = False
    Exit Sub
    
ErrSection:
    m.bReconnecting = False
    RaiseError "cRealTime.Reconnect", eGDRaiseError_Raise
    Resume ErrExit
End Sub

' Returns TRUE only if in the process of disconnecting and then immediately reconnecting
Public Property Get Reconnecting() As Boolean
    Reconnecting = m.bReconnecting
End Property

Public Sub TestMsg(ByVal strMsg$)

    Dim iPos&, strSymbol$, iTick&
    Dim aFlds As New cGdArray
    Dim TickBuffer As cGdBars
    
    aFlds.SplitFields strMsg, vbTab
    If 1 Then
        strSymbol = UCase(Trim(aFlds(0)))
        If m.TickBufferSymbols.BinarySearch(strSymbol, iPos) Then
            For iPos = iPos To m.TickBufferSymbols.Size - 1
                If m.TickBufferSymbols(iPos) <> strSymbol Then Exit For
                Set TickBuffer = m.TickBuffers(iPos)
                If Not TickBuffer Is Nothing Then
                    If TickBuffer.Prop(eBARS_CustomString) <> "<RELOAD>" Then
                        TickBuffer.Size = TickBuffer.Size + 1
                        iTick = TickBuffer.Size - 1
                        TickBuffer(eBARS_DateTime, iTick) = RoundToSecond(Val(aFlds(1)))
                        TickBuffer(eBARS_Close, iTick) = Val(aFlds(2))
                        TickBuffer(eBARS_Vol, iTick) = Val(aFlds(3))
                        TickBuffer(eBARS_OI, iTick) = Val(aFlds(5))
                        TickBuffer.Prop(eBARS_Exchange) = Trim(aFlds(6))
                        FixTicks TickBuffer, True
                    End If
                End If
            Next
            Set TickBuffer = Nothing
        End If
    End If

End Sub

Public Sub CreateTransFile()

    Dim i&, iPos&, iSymbol&
    Dim dMult#, nCrossoverTime&, nGmtOffset&, nStartTime&
    Dim strSymbol$, strSecType$, strText$, strBase$, strFeedSymbol$, strFeedExchange$, strFile$
    Dim aCompSymbols As New cGdArray, aCompExchanges As New cGdArray
    Dim aFile As New cGdArray, aBases As New cGdArray
    Dim Bars As New cGdBars
    
    For iSymbol = 0 To g.SymbolPool.NumRecords - 1
        strSymbol = ""
        Select Case g.SymbolPool.SecType(iSymbol)
        Case eSYMType_Index
            ' do all indices
            strSymbol = g.SymbolPool.Symbol(iSymbol)
            If Left(strSymbol, 1) <> "$" Then
                strSymbol = ""
            Else
                strBase = strSymbol
            End If
        Case eSYMType_Future
            ' just do each base symbol once
            strSymbol = g.SymbolPool.Symbol(iSymbol)
            strBase = Parse(strSymbol, "-", 1)
            strSymbol = strBase & "-199912"
            If aBases.BinarySearch(strBase, iPos) Then
                strSymbol = "" ' already done this base
            Else
                aBases.Add strBase, iPos
            End If
        Case eSYMType_Stock
            ' just do each exchange once
            strSymbol = g.SymbolPool.Symbol(iSymbol)
            strBase = Parse(strSymbol, "@", 2)
            If Len(strBase) > 0 Then
                strBase = "@" & strBase
            Else
                strBase = "@US"
            End If
            If aBases.BinarySearch(strBase, iPos) Then
                strSymbol = "" ' already done this base
            Else
                aBases.Add strBase, iPos
            End If
        End Select
        
        If Len(strSymbol) > 0 Then
If Left(strSymbol, 4) = "YM1-" Then
    dMult = dMult
End If
            If TranslateSymbol(strSymbol, "C", strSecType, strFeedSymbol, strFeedExchange, _
                            dMult, nCrossoverTime, nGmtOffset, False, aCompSymbols, aCompExchanges) Then
                
                ' encode exchanges (into Hex)
                strFeedExchange = EncodeExchange(strFeedExchange)
                For i = 0 To aCompExchanges.Size - 1
                    aCompExchanges(i) = EncodeExchange(aCompExchanges(i))
                Next
                           
                ' convert from MinSinceMid to HHMM (NY time for American, GMT for foreign)
                If nCrossoverTime > 0 Then nCrossoverTime = Int(nCrossoverTime / 60) * 100 + (nCrossoverTime Mod 60)
                                    
                ' for RT client need to add:
                ' - exchange offset from GMT in minutes (+/-)
                ' - for day-session synthetic symbols, add session starting time as HHMM (else leave blank)
                strText = ""
                If strSecType = "F" Then
                    SetBarProperties Bars, strSymbol
                    If Bars.Prop(eBARS_Session) = Asc("S") Then
                        ' convert StartTime from MinSinceMid to HHMM
                        nStartTime = Bars.Prop(eBARS_StartTime)
                        nStartTime = Int(nStartTime / 60) * 100 + (nStartTime Mod 60)
                        strText = Str(nStartTime)
                    End If
                End If
                                
                ' RT client:  "WP-067  F   SP!M4      1625    1   SPM4,SP&M4  , 0"
                strText = strBase & vbTab & strSymbol & vbTab & strSecType & vbTab & strFeedSymbol _
                    & vbTab & strFeedExchange & vbTab & Str(nCrossoverTime) & vbTab & Str(dMult) _
                    & vbTab & aCompSymbols.JoinFields(",") & vbTab & aCompExchanges.JoinFields(",") _
                    & vbTab & Str(nGmtOffset) & vbTab & strText
                
                aFile.Add strText
            End If
        End If
    Next
    
    aFile.Add "ZZZ ZZ" & vbTab & "ZZZ ZZ" & vbTab & "SO" & vbTab & "ZZZZZ" & vbTab & "[0C]" & vbTab & "0" & vbTab & "1"
    
    strFile = "C:\SymInf.txt"
    aFile.Sort
    aFile.ToFile strFile
    InfBox strFile & " created"

End Sub

Public Property Get IsMarketDepthMax(ByVal strSymbol$, _
    Optional ByVal strDropSymbol$ = "") As Boolean
On Error GoTo ErrSection:

    Dim bIsMax As Boolean
    Dim BidAskDepth As cBidAskDepth
    
    If ActiveRTG Then
        If m.MktDepthList.Count = MaxSymbolsAllowed(True) Then
            strSymbol = RollSymbolForDate(strSymbol)
            Set BidAskDepth = m.MktDepthList(strSymbol)
            If BidAskDepth Is Nothing Then
                If Len(strDropSymbol) > 0 Then
                    strDropSymbol = RollSymbolForDate(strDropSymbol)
                    Set BidAskDepth = m.MktDepthList(strDropSymbol)
                    If BidAskDepth Is Nothing Then
                        bIsMax = True
                    ElseIf BidAskDepth.UsageCounter > 1 Then
                        bIsMax = True
                    End If
                Else
                    bIsMax = True
                End If
            End If
        End If
    End If
    
    IsMarketDepthMax = bIsMax

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cRealTime.IsMarketDepthMax", eGDRaiseError_Raise
    Resume ErrExit

End Property

' iMsgType: 20 = trade, 25 = bid, 26 = ask, 52 = depth-of-market
Public Sub StreamedMessage(ByVal iMsgType&, ByVal strMessage$, Optional nAcctType As eTT_AccountType = 0)
On Error GoTo ErrSection

    Dim strSymbol$, i&, iPos&, iTick&, dPrice#, dOrigPrice#, dMinMove#, strBaseSymbol$, dTime#
    Dim bSkip As Boolean
    Dim aFlds As New cGdArray
    Dim TickBuffer As cGdBars
    Dim BidAskDepth As cBidAskDepth
    Dim aLastQuote As cGdArray
    Static dLastLogged As Double
    Static aLogging As cGdArray
        
Static nChkCount&, nPrevMSM&
           
    If nAcctType = 0 Then m.dLastMsgTime = gdTickCount ' store tick count in order to know when last received something

    ' first time in, determine whether to enable ES logging or not (based on flag file)
    If dLastLogged = 0 Then
        If FileExist(App.Path & "\ES.LOG") Then
            Set aLogging = New cGdArray
            aLogging.Add "ES log"
            aLogging.ToFile App.Path & "\ES.LOG", False
            aLogging.Size = 0
            dLastLogged = gdTickCount
        Else
            dLastLogged = -1
        End If
    End If

    ' check if message is from the "correct" feed (on a per symbol basis)
    aFlds.SplitFields strMessage, vbTab
    strSymbol = UCase(Trim(aFlds(0)))
    If iMsgType = 20 And dLastLogged > 0 Then
        ' log all ES trades (individual contracts)
        If Left(strSymbol, 5) = "ES-20" Then
            If nAcctType = 0 Then
                aLogging.Add Format(Now, "hh:mm:ss") & vbTab & Format(Me.FeedTime, "hh:mm:ss") _
                    & vbTab & Str(nAcctType) & vbTab & strMessage
            Else
                aLogging.Add Format(Now, "hh:mm:ss") & vbTab & Format(Me.FeedTime, "hh:mm:ss") _
                    & vbTab & Str(nAcctType) & vbTab & vbTab & strMessage
            End If
            If gdTickCount > dLastLogged + 15000 Then
                aLogging.ToFile App.Path & "\ES.LOG", True
                aLogging.Size = 0
                dLastLogged = gdTickCount
            End If
        End If
    End If
    If m.nBrokerFeed = 0 And nAcctType <> 0 Then
        strSymbol = ""
    ElseIf m.nBrokerFeed > 0 And nAcctType = 0 Then
        i = InStr(strSymbol, "-2")
        If i > 0 Then
            If m.BrokerSubscriptionList.BinarySearch(strSymbol) Then
                If IsBrokerSymbol(strSymbol) Then
                    strSymbol = ""
                End If
            End If
        End If
    End If
    
    ' loop for standard and synthetic symbol
    Do While Len(strSymbol) > 0
        Select Case iMsgType
        Case 20 ' New Tick (GenSymbol, Date.Time, Price, Vol, ??, OI, Exch, Flags)
If IsIDE Then
    If strSymbol = "ES-067" Then
        dPrice = RoundToSecond(Val(aFlds(1)))
        i = Int(1440 * (dPrice - Int(dPrice)) + 0.0000001)
        If i <> nPrevMSM Then
            If nPrevMSM <> 0 And frmTest.Visible Then
                frmTest.AddList strSymbol & " at " & Format(Int(i / 60), "#00") & ":" _
                    & Format(i Mod 60, "#00") & " = " & Str(nChkCount) & " streamed ticks"
            End If
            nPrevMSM = i
            nChkCount = 0
        End If
        nChkCount = nChkCount + 1
    End If
End If
            If Len(aFlds(2)) = 0 Then
                dPrice = kNullData
            Else
                dPrice = Val(aFlds(2))
            End If
            dTime = RoundToSecond(Val(aFlds(1)))
            If m.TickBufferSymbols.BinarySearch(strSymbol, iPos) Then
                If m.dEminiOffset <> 0 Then
                    If Left(strSymbol, 4) = "ES-2" Then
                        dPrice = dPrice + m.dEminiOffset
                    End If
                End If
                For iPos = iPos To m.TickBufferSymbols.Size - 1
                    If m.TickBufferSymbols(iPos) <> strSymbol Then Exit For
                    Set TickBuffer = m.TickBuffers(iPos)
                    If Not TickBuffer Is Nothing Then
                        bSkip = False
                        If TickBuffer.Prop(eBARS_CustomString) = "<RELOAD>" Then
                            If Left(m.TickBuffers.Key(iPos), 1) = "^" Then
                                TickBuffer.Prop(eBARS_CustomString) = "" ' (reset for generic tick buffer)
                            Else
                                bSkip = True
                            End If
                        End If
                        If SalmonIsRunning Then
                            If Left(m.TickBuffers.Key(iPos), 1) = "^" Then bSkip = True
                        End If
                        If Not bSkip Then
                            TickBuffer.Size = TickBuffer.Size + 1
                            iTick = TickBuffer.Size - 1
                            TickBuffer(eBARS_DateTime, iTick) = dTime
                            TickBuffer(eBARS_Close, iTick) = dPrice
                            TickBuffer(eBARS_Vol, iTick) = Val(aFlds(3))
                            TickBuffer(eBARS_OI, iTick) = Val(aFlds(5))
                            TickBuffer(eBARS_Flags, iTick) = Val(aFlds(7))
                            TickBuffer.Prop(eBARS_Exchange) = Trim(aFlds(6))
'If Left(strSymbol, 4) = "YI2-" Then
'    iTick = iTick
'End If
                            FixTicks TickBuffer, True
                            ' save fixed price (if it didn't get deleted due to custom stop/end)
                            If iTick < TickBuffer.Size Then
                                dPrice = TickBuffer(eBARS_Close, iTick)
                            End If
                        End If
                    End If
                Next
                Set TickBuffer = Nothing
                m.dQuoteCount = m.dQuoteCount + 1
            End If
            
            Set aLastQuote = m.LastQuote(strSymbol)
            If aLastQuote Is Nothing Then
                Set aLastQuote = New cGdArray
                aLastQuote.Create eGDARRAY_Doubles, LQ(eLQ_NUM_ITEMS)
                m.LastQuote.Add aLastQuote, strSymbol
            End If
            If dPrice <> kNullData Then
                aLastQuote.Num(LQ(eLQ_TradePrice)) = dPrice
            End If
            iPos = Val(aFlds(3))
            If iPos > 0 Then
                aLastQuote.Num(LQ(eLQ_TradeVol)) = iPos
            End If
            If dTime > 0 Then
                aLastQuote.Num(LQ(eLQ_TradeTime)) = dTime
            End If
            If IsForex(strSymbol) And Right(strSymbol, 4) <> "@PFG" Then
                If dPrice <> kNullData Then
                    aLastQuote.Num(LQ(eLQ_BidPrice)) = dPrice
                End If
                If iPos > 0 Then
                    aLastQuote.Num(LQ(eLQ_BidSize)) = iPos
                End If
                If dTime > 0 Then
                    aLastQuote.Num(LQ(eLQ_BidAskTime)) = dTime
                End If
            End If
            Set aLastQuote = Nothing
    
    
        Case 25, 26 ' New Bid or Ask (GenSymbol, Date.Time, BidPrice, BidSize)
            If Len(aFlds(2)) = 0 Then
                dPrice = kNullData
            Else
                dPrice = Val(aFlds(2))
            End If
            If m.dEminiOffset <> 0 Then
                If Left(strSymbol, 4) = "ES-2" Then
                    dPrice = dPrice + m.dEminiOffset
                End If
            End If
            ' for futures, fix the price (using past good tick from previous day's settle)
            If SecurityType(strSymbol, True) = "F" Then
                strBaseSymbol = Parse(strSymbol, "-", 1)
                If m.PastGoodPrice.Exists(strBaseSymbol) Then
                    ' try to get the MinMove from a TickBuffer (one will most likely exist for this symbol)
                    dMinMove = 0
                    If m.TickBufferSymbols.BinarySearch(strSymbol, iPos) Then
                        Set TickBuffer = m.TickBuffers(iPos)
                        If Not TickBuffer Is Nothing Then
                            dMinMove = TickBuffer.MinMove(Date)
                        End If
                        Set TickBuffer = Nothing
                    End If
                    dPrice = FixPrice(dPrice, dMinMove, m.PastGoodPrice(strBaseSymbol))
                End If
            End If
            
            ' store as bid or ask
            Set aLastQuote = m.LastQuote(strSymbol)
            If aLastQuote Is Nothing Then
                Set aLastQuote = New cGdArray
                aLastQuote.Create eGDARRAY_Doubles, LQ(eLQ_NUM_ITEMS)
                m.LastQuote.Add aLastQuote, strSymbol
            End If
            iPos = Val(aFlds(3))
            If iMsgType = 25 Then
                If dPrice > 0 Then
                    aLastQuote.Num(LQ(eLQ_BidPrice)) = dPrice
                End If
                If iPos > 0 Then
                    aLastQuote.Num(LQ(eLQ_BidSize)) = iPos
                End If
            Else
                If dPrice > 0 Then
                    aLastQuote.Num(LQ(eLQ_AskPrice)) = dPrice
                End If
                If iPos > 0 Then
                    aLastQuote.Num(LQ(eLQ_AskSize)) = iPos
                End If
            End If
            dTime = RoundToSecond(Val(aFlds(1)))
            If dTime > 0 Then
                aLastQuote.Num(LQ(eLQ_BidAskTime)) = dTime
            End If
            Set aLastQuote = Nothing
            m.dQuoteCount = m.dQuoteCount + 1
            
        Case 52 ' Depth-of-market (GenSymbol, MarketMaker, DateTime, BidPrice, BidSize, AskPrice, AskSize, BidTick, QCond)
            Set BidAskDepth = m.MktDepthList(strSymbol)
            If Not BidAskDepth Is Nothing Then
                BidAskDepth.NewData strMessage
                Set BidAskDepth = Nothing
                If 0 And IsIDE Then
                    If frmTest.Visible Then 'And Parse(strMessage, vbTab, 2) = "0" Then
                        frmTest.AddList strMessage
                    End If
                End If
            End If
        
        End Select
    
Exit Do
        ' if not GenesisRT, then duplicate for synthetic symbol (if exists)
        If nAcctType = 0 Then
            strSymbol = ""
        Else
            strSymbol = ConvertSynthetic(strSymbol, True)
            If Len(strSymbol) > 0 Then
                aFlds(0) = strSymbol
                strMessage = aFlds.JoinFields(vbTab)
            End If
        End If
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.StreamedMessage"
    Resume ErrExit
End Sub

' iMsgType: 20 = trade, 25 = bid, 26 = ask, 52 = depth-of-market
Public Sub StreamedQuote(ByVal iMsgType&, ByVal nMsgPtr&, ByVal nMsgLen&, Optional nAcctType As eTT_AccountType = 0)
On Error GoTo ErrSection

    Dim strSymbol$, strExch$, strBaseSymbol$
    Dim i&, iPos&, iTick&, dPrice#, dPastGoodPrice#, dTime#, dMinMove#
    Dim bSkip As Boolean
    Dim TickBuffer As cGdBars
    Dim BidAskDepth As cBidAskDepth
    Dim QuoteMsg As mQuoteMsg
    Dim aLastQuote As cGdArray
           
    If nAcctType = 0 Then m.dLastMsgTime = gdTickCount ' store tick count in order to know when last received something

    CopyMemory QuoteMsg, ByVal nMsgPtr, nMsgLen
    
    strSymbol = QuoteMsg.strSymbol
    FixNullTermStr strSymbol
    strSymbol = UCase(Trim(strSymbol))
    
'If InStr(strSymbol, " ") > 0 Then
'    i = i
'End If
    
    strExch = QuoteMsg.strExch
    FixNullTermStr strExch
    strExch = UCase(Trim(strExch))
    
    ' check if message is from the "correct" feed (on a per symbol basis)
    If m.nBrokerFeed = 0 And nAcctType <> 0 Then
        strSymbol = ""
    ElseIf m.nBrokerFeed > 0 And nAcctType = 0 Then
        i = InStr(strSymbol, "-2")
        If i > 0 Then
            If m.BrokerSubscriptionList.BinarySearch(strSymbol) Then
                If IsBrokerSymbol(strSymbol) Then
                    strSymbol = ""
                End If
            End If
        End If
    End If
    
    ' loop for standard and synthetic symbol
    Do While Len(strSymbol) > 0
        Select Case iMsgType
        Case 201 ' New Tick (GenSymbol, Date.Time, Price, Vol, ??, OI, Exch, Flags)
            If m.TickBufferSymbols.BinarySearch(strSymbol, iPos) Then
                dTime = RoundToSecond(QuoteMsg.dDateTime)
                dPrice = QuoteMsg.dPrice
                For iPos = iPos To m.TickBufferSymbols.Size - 1
                    If m.TickBufferSymbols(iPos) <> strSymbol Then Exit For
                    Set TickBuffer = m.TickBuffers(iPos)
                    If Not TickBuffer Is Nothing Then
                        bSkip = False
                        dMinMove = TickBuffer.MinMove(Date)
                        If TickBuffer.Prop(eBARS_CustomString) = "<RELOAD>" Then
                            If Left(m.TickBuffers.Key(iPos), 1) = "^" Then
                                TickBuffer.Prop(eBARS_CustomString) = "" ' (reset for generic tick buffer)
                            Else
                                bSkip = True
                            End If
                        End If
                        If SalmonIsRunning Then
                            If Left(m.TickBuffers.Key(iPos), 1) = "^" Then bSkip = True
                        End If
                        If Not bSkip Then
                            TickBuffer.Size = TickBuffer.Size + 1
                            iTick = TickBuffer.Size - 1
                            TickBuffer(eBARS_DateTime, iTick) = dTime
                            TickBuffer(eBARS_Close, iTick) = dPrice
                            TickBuffer(eBARS_Vol, iTick) = QuoteMsg.nSize
                            TickBuffer(eBARS_OI, iTick) = QuoteMsg.nSize2
                            TickBuffer(eBARS_Flags, iTick) = QuoteMsg.nFlags
                            TickBuffer.Prop(eBARS_Exchange) = strExch
'If Left(strSymbol, 4) = "YI2-" Then
'    iTick = iTick
'End If
                            ' save fixed price (if it didn't get deleted due to custom stop/end)
                            FixTicks TickBuffer, True
                            If iTick < TickBuffer.Size Then
                                dPrice = TickBuffer(eBARS_Close, iTick)
                            End If
                        End If
                    End If
                Next
                Set TickBuffer = Nothing
                m.dQuoteCount = m.dQuoteCount + 1
            
                ' replace last quote
                If dPrice <> kNullData Then
                    Set aLastQuote = m.LastQuote(strSymbol)
                    If aLastQuote Is Nothing Then
                        Set aLastQuote = New cGdArray
                        aLastQuote.Create eGDARRAY_Doubles, LQ(eLQ_NUM_ITEMS)
                        m.LastQuote.Add aLastQuote, strSymbol
                    End If
                    aLastQuote.Num(LQ(eLQ_TradeTime)) = dTime
                    aLastQuote.Num(LQ(eLQ_TradePrice)) = dPrice
                    If QuoteMsg.nSize > 0 Then
                        aLastQuote.Num(LQ(eLQ_TradeVol)) = QuoteMsg.nSize
                    End If
                    If dMinMove > 0 Then
                        aLastQuote.Num(LQ(eLQ_MinMove)) = dMinMove
                    End If
                    ' and replace bid if it's a Forex trade other than PFG
                    If IsForex(strSymbol) And Right(strSymbol, 4) <> "@PFG" Then
                        aLastQuote.Num(LQ(eLQ_BidPrice)) = dPrice
                        If QuoteMsg.nSize > 0 Then
                            aLastQuote.Num(LQ(eLQ_BidSize)) = QuoteMsg.nSize
                        End If
                        aLastQuote.Num(LQ(eLQ_BidAskTime)) = dTime
                    End If
                    Set aLastQuote = Nothing
                End If
            End If
    
        Case 202 ' New Bid and/or Ask
            Set aLastQuote = m.LastQuote(strSymbol)
            If aLastQuote Is Nothing Then
                Set aLastQuote = New cGdArray
                aLastQuote.Create eGDARRAY_Doubles, LQ(eLQ_NUM_ITEMS)
                m.LastQuote.Add aLastQuote, strSymbol
            End If
            
            ' for futures, fix the price (using past good tick from previous day's settle)
            dMinMove = aLastQuote.Num(LQ(eLQ_MinMove))
            dPastGoodPrice = kNullData
            If SecurityType(strSymbol, True) = "F" Then
                dPastGoodPrice = aLastQuote.Num(LQ(eLQ_PastGoodPrice))
                If dPastGoodPrice = kNullData Then
                    strBaseSymbol = Parse(strSymbol, "-", 1)
                    If m.PastGoodPrice.Exists(strBaseSymbol) Then
                        dPastGoodPrice = m.PastGoodPrice(strBaseSymbol)
                        aLastQuote.Num(LQ(eLQ_PastGoodPrice)) = dPastGoodPrice
                    End If
                End If
            End If
            
            ' bid/ask prices and sizes can only be positive
            dTime = RoundToSecond(QuoteMsg.dDateTime)
            dPrice = QuoteMsg.dPrice
            If dPrice > 0 Then
                dPrice = FixPrice(dPrice, dMinMove, dPastGoodPrice)
                aLastQuote.Num(LQ(eLQ_BidPrice)) = dPrice
                If QuoteMsg.nSize > 0 Then
                    aLastQuote.Num(LQ(eLQ_BidSize)) = QuoteMsg.nSize
                End If
                If dTime > 0 Then
                    aLastQuote.Num(LQ(eLQ_BidAskTime)) = dTime
                    dTime = 0 ' since don't need to store it again below
                End If
            End If
            dPrice = QuoteMsg.dPrice2
            If dPrice > 0 Then
                dPrice = FixPrice(dPrice, dMinMove, dPastGoodPrice)
                aLastQuote.Num(LQ(eLQ_AskPrice)) = dPrice
                If QuoteMsg.nSize2 > 0 Then
                    aLastQuote.Num(LQ(eLQ_AskSize)) = QuoteMsg.nSize2
                End If
                If dTime > 0 Then
                    aLastQuote.Num(LQ(eLQ_BidAskTime)) = dTime
                End If
            End If

            Set aLastQuote = Nothing
            m.dQuoteCount = m.dQuoteCount + 1
            
        Case 203 ' Depth-of-market (GenSymbol, MarketMaker, DateTime, BidPrice, BidSize, AskPrice, AskSize, BidTick, QCond)
            Set BidAskDepth = m.MktDepthList(strSymbol)
            If Not BidAskDepth Is Nothing Then
                BidAskDepth.NewQuote strSymbol, strExch, QuoteMsg.dDateTime, _
                    QuoteMsg.dPrice, QuoteMsg.nSize, QuoteMsg.dPrice2, QuoteMsg.nSize2, Mid(QuoteMsg.strFlags, 2, 1)
                Set BidAskDepth = Nothing
            End If
        
        End Select
        
Exit Do
        ' if not GenesisRT, then duplicate for synthetic symbol (if exists)
        If nAcctType = 0 Then
            strSymbol = ""
        Else
            strSymbol = ConvertSynthetic(strSymbol, True)
            If Len(strSymbol) > 0 Then
                ''aFlds(0) = strSymbol
                ''strMessage = aFlds.JoinFields(vbTab)
            End If
        End If
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.StreamedQuote"
    Resume ErrExit
End Sub

Public Function FeedTime(Optional ByVal strSymbol$ = "") As Double
On Error GoTo ErrSection
    
    Dim dDelay#, dMult#
    
    If m.dFeedTime > 0 Then
        ' if accounting for delayed vs. realtime for specified symbol
        If Len(strSymbol) > 0 And g.nReplaySession = 0 Then
            dDelay = SymbolDelay(strSymbol)
            If dDelay < 0 Then dDelay = 0
        End If
        
        ' account for replay speed
        If g.nReplaySession > 0 Then
            dMult = m.dReplaySpeed
        End If
        If dMult = 0 Then dMult = 1
        
        FeedTime = m.dFeedTime + dMult * (gdTickCount - m.dLastFeedTimeMsg) / 86400000 - (dDelay / 1440#)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.FeedTime"
    Resume ErrExit
End Function

Public Sub AddOptNavHistSymbol(ByVal strGenSymbol$)
On Error GoTo ErrSection

    Dim i&
   
    ' add if an option symbol that does not already exist
    strGenSymbol = UCase(Trim(strGenSymbol))
    If InStr(strGenSymbol, " ") > 0 Then
        If Not m.OptNavHistSymbols.BinarySearch(strGenSymbol, i) Then
            m.OptNavHistSymbols.Add strGenSymbol, i
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.AddOptNavHistSymbol"
    Resume ErrExit
End Sub

Public Sub AddBrokerRtSymbol(ByVal strGenSymbol$, ByVal strBroker$)
On Error GoTo ErrSection

    Dim i&, strText$

    ' TLB for internal use only: check for special Realtime AP override
    If m.BrokerRtSymbols.Size = 0 Then
        If FileExist(App.Path & "\RealtimeAP.flg") Then
            m.BrokerRtSymbols.Add "AP-" & vbTab
        End If
    End If

    ' strip contract and add broker
    strGenSymbol = UCase(Trim(strGenSymbol))
    i = InStr(strGenSymbol, "-")
    If i > 0 Then strGenSymbol = Left(strGenSymbol, i)
    strText = strGenSymbol & vbTab & UCase(Trim(strBroker))
    
    ' add if not already exist
    If Not m.BrokerRtSymbols.BinarySearch(strText, i) Then
        m.BrokerRtSymbols.Add strText, i
        If 0 Then
            strText = ConvertSynthetic(strText, True)
            If Len(strText) > 0 Then
                If Not m.BrokerRtSymbols.BinarySearch(strText, i) Then
                    'm.BrokerRtSymbols.Add strText, i
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.AddBrokerRtSymbol"
    Resume ErrExit
End Sub

Public Property Get UseBrokerFeed() As eTT_AccountType
    UseBrokerFeed = m.nBrokerFeed
End Property

Public Property Let UseBrokerFeed(ByVal nUseBrokerFeed As eTT_AccountType)
    If g.Broker.IsLiveAccount(nUseBrokerFeed) = False Then
        m.nBrokerFeed = 0
    Else
        m.nBrokerFeed = nUseBrokerFeed
    End If
End Property

Public Function BrokerSubscriptionList(ByVal nAcctType As eTT_AccountType) As cGdArray
    Dim aEmpty As cGdArray
    If 0 Then ' m.nBrokerFeed = 0 Then
        Set aEmpty = New cGdArray
        aEmpty.Create eGDARRAY_Strings, 0
        Set BrokerSubscriptionList = aEmpty
    Else
        Set BrokerSubscriptionList = m.BrokerSubscriptionList
    End If
End Function

Public Function IsBrokerSymbol(ByVal strGenSymbol$, Optional iPos As Long = 0) As Boolean
On Error GoTo ErrSection

    Dim i&, strTemp$
    
    ' TLB 5/2/2011: for ASX stocks (and indices?) check if TWS has been installed
    If Right(strGenSymbol, 4) = "@ASX" Then
        If Len(Trim(GetRegistryValue(rkLocalMachine, "Software\Trader Workstation", "jtspath", ""))) > 2 Then
            IsBrokerSymbol = True
        ElseIf InStr(UCase(GetRegistryValue(rkClassesRoot, "tws", "", "")), "TRADER") > 0 Then ' newer version of TWS
            IsBrokerSymbol = True
        End If
    ' TLB 2/9/2015: for $Forex@Broker symbols, just allow realtime (since we don't even support delayed data)
    ElseIf IsForex(strGenSymbol) And InStr(strGenSymbol, "@") > 1 Then
        IsBrokerSymbol = True
    ' for everything else ...
    ElseIf m.BrokerRtSymbols.Size > 0 Then
        i = InStr(strGenSymbol, "-")
        If i > 0 Then
            ' strip off the contract
            strGenSymbol = Left(strGenSymbol, i)
            ' convert from a synthetic (i.e. the non-synthentic should be checked)
            strTemp = ConvertSynthetic(strGenSymbol, False)
            If Len(strTemp) > 0 Then strGenSymbol = strTemp
        End If
        If m.BrokerRtSymbols.BinarySearch(strGenSymbol & vbTab, iPos, eGdSort_MatchUsingSearchStringLength) Then
            IsBrokerSymbol = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.IsBrokerSymbol"
    Resume ErrExit
End Function

Public Property Get EminiOffset() As Double
    EminiOffset = m.dEminiOffset
End Property

Public Property Let EminiOffset(ByVal dOffSet As Double)
    If FileExist("c:\common\files.exe") Then
        m.dEminiOffset = dOffSet
    End If
End Property

Public Property Get SocketLogging() As Boolean
    SocketLogging = m.bSocketLogging
End Property

Public Property Let SocketLogging(ByVal bLog As Boolean)
    If bLog <> m.bSocketLogging Then
        m.bSocketLogging = bLog
        m.RT.CreateMessage m.strServerNode, 105, Str(Abs(Val(bLog))), , True
    End If
End Property

Public Function GetSecondsBehind() As Long
    On Error Resume Next
    GetSecondsBehind = CheckLatency(kNullData, "")
End Function

Private Function CheckLatency(ByVal nMsgType&, ByVal strMsg$) As Long
On Error GoTo ErrSection

    Dim i&, dBehind#, dTickTime#, dGenRtLatency#, dWhenSent#, strFile$, bLog As Boolean
    Static nDisplayCutoff As Long
    Static iSecondsBehind As Long
    Static dMaxBehind As Double
    Static dLastLogged As Double
    Static aLogging As cGdArray
    Static dFeedDate As Double
    Static strPrevIpKey As String

    If m.dFeedTime > 0 And m.nServerStatus > 0 And UCase(m.strServerNode) = "GENESISRT" And g.nReplaySession = 0 Then
        ' if nMsgType = kNullData then we're just returning the current latency
        If nMsgType <> kNullData Then
            ' init some things (first time)
            If dLastLogged = 0 Then
                Set aLogging = New cGdArray
                ' get "display cutoff" from file (i.e. display only if over # seconds)
                aLogging.FromFile "c:\ShowDelay.flg"
                If aLogging.Size = 0 Then
                    aLogging.FromFile App.Path & "\Provided\ShowDelay.flg"
                End If
                For i = 0 To aLogging.Size - 1
                    If HasModule(Parse(aLogging(i), vbTab, 2)) Then
                        nDisplayCutoff = Val(Parse(aLogging(i), vbTab, 1))
                        Exit For
                    End If
                Next
                If nDisplayCutoff <= 0 Then nDisplayCutoff = 5 ' (default)
                DebugLog "DisplayCutoff = " & Str(nDisplayCutoff)
                
                aLogging.Size = 0
                dLastLogged = gdTickCount
            End If
                
            If Right(strMsg, 1) = vbTab Then strMsg = Left(strMsg, Len(strMsg) - 1)
            Select Case nMsgType
            Case 2, 110
                ' Parse status msg: Status <tab> FeedTime <tab> RT latency <tab> gdTickCount when sent
                dTickTime = RoundToSecond(Val(Parse(strMsg, vbTab, 2)))
                dGenRtLatency = Val(Parse(strMsg, vbTab, 3))
                dWhenSent = Val(Parse(strMsg, vbTab, 4))
                If dFeedDate <> Int(dTickTime) Then
                    dFeedDate = Int(dTickTime)
                    dMaxBehind = 0
                End If
                    
                If dWhenSent > 0 Then
                    ' Calculate seconds behind
                    bLog = False
                    dBehind = (dGenRtLatency + (gdTickCount - dWhenSent)) / 1000
                    If dBehind >= 0 Then 'TLB 6/9/2006: need this check due to multiprocessor issues
                        If dBehind > dMaxBehind Then
                            dMaxBehind = dBehind
                        End If
                        If Int(dBehind) <> iSecondsBehind Then
                            iSecondsBehind = Int(dBehind)
                            If iSecondsBehind > nDisplayCutoff Then
                                ' display 1 second less since not sure about sub-second accuracy?
                                StatusMsg "Data stream is " & Str(iSecondsBehind - 1) & " seconds behind", RGB(250, 0, 0)
                            ElseIf frmMain.tbToolbar.Tools("ID_Status").ForeColor = RGB(250, 0, 0) Then
                                StatusMsg ""
                            End If
                            bLog = True
                        ElseIf nMsgType = 2 Then
                            bLog = True
                        End If
                        If bLog Then
                            i = Int(dBehind)
                            If i < 0 Then
                                i = 0
                            ElseIf i > 100 Then
                                i = 100
                            End If
                            aLogging.Add Format(Time, "hh:mm:ss") & vbTab & Str(nMsgType) & vbTab & _
                                Format(dTickTime, "hh:mm:ss") & vbTab & Format(dBehind, "#0.000") & _
                                vbTab & String(i, "#") & vbTab & Format(dMaxBehind, "#0.0") & vbTab & _
                                strMsg & vbTab & Str(m.nTickHistTimeoutCount) & "/" & Str(m.nTickHistRequestCount)
                            m.nTickHistRequestCount = 0
                            m.nTickHistTimeoutCount = 0
                        End If
                    End If
                
                    ' turn socket logging on/off (extra debugging info for streaming)
                    If nMsgType = 2 Then
                        SocketLogging = FileExist(App.Path & "\SocketLog.All")
                    End If
                End If
                
            Case Else
                ' just log these messages
                If nMsgType = 106 Then
                    strMsg = Parse(strMsg, vbCrLf, 1) ' strip off the gdProfiles stuff
                End If
                aLogging.Add Format(Time, "hh:mm:ss") & vbTab & Str(nMsgType) & vbTab & strMsg
            End Select
            
            ' flush to log once every 10 seconds
            If gdTickCount > dLastLogged + 10000 And aLogging.Size > 0 And dFeedDate > 0 Then
                strFile = App.Path & "\Stream\" & Format(dFeedDate, "YYYYMMDD") & ".Log"
                If strPrevIpKey <> m.strIPandKey Or Not FileExist(strFile) Then
                    strPrevIpKey = m.strIPandKey
                    aLogging.Add "Splitter IP and Key:" & vbTab & m.strIPandKey, 0
                End If
                aLogging.ToFile strFile, True
                aLogging.Size = 0
                dLastLogged = gdTickCount
            End If
        End If
    Else
        iSecondsBehind = 0
    End If

ErrExit:
    CheckLatency = iSecondsBehind
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.CheckLatency"
    Resume ErrExit
End Function

Public Property Get ReplaySpeed() As Double
    ReplaySpeed = m.dReplaySpeed
End Property

Public Property Let ReplaySpeed(ByVal dSpeed As Double)
    m.dReplaySpeed = dSpeed
    If g.RealTime.ActiveRTG And g.nReplaySession > 0 Then
        m.RT.CreateMessage m.strServerNode, 42, Str(m.dReplaySpeed)
    End If
End Property

Public Sub SetReplayTime(ByVal dStartTime As Double)
On Error GoTo ErrSection

    Dim s$

    m.dFeedTime = dStartTime
    g.nReplaySession = Int(dStartTime)
    If g.RealTime.ActiveRTG And g.nReplaySession > 0 Then
        ' subtract 1/10 second due to rounding issues when converting to/from a string
        s = Str(ConvertTimeZone(m.dFeedTime - 0.000001, "NY", "GMT"))
        m.RT.CreateMessage m.strServerNode, 43, s
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.SetReplayTime"
    Resume ErrExit
End Sub

Public Property Get ConnectionStatus() As eGDConnectionStatus
    ConnectionStatus = m.nConnectionStatus
End Property
Public Property Let ConnectionStatus(ByVal nConnectionStatus As eGDConnectionStatus)
    m.nConnectionStatus = nConnectionStatus
    If (g.nReplaySession = 0) And (Not g.SimTradeStream Is Nothing) Then
        g.SimTradeStream.Broker.HandleConnectionInfo nConnectionStatus, "", ""
    'Else
        'g.SimTradeReplay.Broker.HandleConnectionInfo nConnectionStatus, "", ""
    End If
End Property

' to reset the priority on the GenesisRT thread
Public Sub SetRTPriority(Optional ByVal bHighPriority As Boolean = False)
On Error GoTo ErrSection
    
    Dim i&
    If g.RealTime.ActiveRTG Then
        ' Do NOT set to high for streaming replay (can lock things up, e.g. at 8x speed)
        If g.nReplaySession > 0 Or FileExist(App.Path & "\Provided\HighRT.skp") Then
            bHighPriority = False
        End If
        If Not bHighPriority Then
            i = NORMAL_PRIORITY_CLASS
        ElseIf IsAtLeastXP Then
            i = ABOVE_PRIORITY_CLASS
        Else
            i = HIGH_PRIORITY_CLASS
        End If
        m.RT.CreateMessage m.strServerNode, 90, Str(i)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.SetRTPriority"
    Resume ErrExit
End Sub

Public Property Get RtInterval() As Long
    RtInterval = m.iRtInterval
End Property

Public Property Let RtInterval(ByVal iRtInterval As Long)
    m.iRtInterval = iRtInterval
    If g.RealTime.Active Then
        m.RT.CreateMessage m.strServerNode, 21, Str(m.iRtInterval)
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastKnownPrice
'' Description: Get the last known price for the given symbol
'' Inputs:      Symbol, Local Time?, Date/Time of last price(return)
''      iBidTradeOrAsk:  -1 = LastBid,  0 = LastTrade,  1 = LastAsk
'' Returns:     Last Known Price (kNullData if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastKnownPrice(ByVal vSymbolOrSymbolID As Variant, Optional ByVal iBidTradeOrAsk As Integer = 0, _
                    Optional ByVal bLocalTime As Boolean = False, Optional dDateTime As Double) As Double
On Error GoTo ErrSection:

    Dim i&, strSymbol$, dPrice#
    Dim Bars As cGdBars             ' Bars structure
    Dim aLastQuote As cGdArray
    
    If SalmonIsRunning Then
        LastKnownPrice = g.RealTime.SymbolInfo(vSymbolOrSymbolID).LastKnownPrice(iBidTradeOrAsk, bLocalTime, dDateTime)
        Exit Function
    End If
    
    dPrice = kNullData
    dDateTime = kNullData
    
    ' first see if we can get it from what was last sent in the realtime stream (much more efficient)
    strSymbol = RollSymbolForDate(GetSymbol(vSymbolOrSymbolID))
    If g.RealTime.Active Then
        Set aLastQuote = m.LastQuote(strSymbol)
        If Not aLastQuote Is Nothing Then
            ' get the price and time
            If iBidTradeOrAsk < 0 Then
                dPrice = aLastQuote.Num(LQ(eLQ_BidPrice))
                dDateTime = aLastQuote.Num(LQ(eLQ_BidAskTime))
            ElseIf iBidTradeOrAsk > 0 Then
                dPrice = aLastQuote.Num(LQ(eLQ_AskPrice))
                dDateTime = aLastQuote.Num(LQ(eLQ_BidAskTime))
            Else
                dPrice = aLastQuote.Num(LQ(eLQ_TradePrice))
                dDateTime = aLastQuote.Num(LQ(eLQ_TradeTime))
            End If
            If dPrice = kNullData Or dPrice = 0 Or dDateTime <= 0 Then
                dDateTime = kNullData
            Else
                dDateTime = RoundToSecond(dDateTime)
            End If
            Set aLastQuote = Nothing
        End If
    End If
    
    ' otherwise try to get it from loading the data (less efficient)
    If dPrice = kNullData Then
        Set Bars = New cGdBars
        If DM_GetBars(Bars, vSymbolOrSymbolID, "Daily", LastDailyDownload - 5) Then
            g.RealTime.SpliceBars Bars
            If iBidTradeOrAsk < 0 Then
                dPrice = Bars(eBARS_Bid, Bars.Size - 1)
            ElseIf iBidTradeOrAsk > 0 Then
                dPrice = Bars(eBARS_Ask, Bars.Size - 1)
            Else
                dPrice = Bars(eBARS_Close, Bars.Size - 1)
            End If
            dDateTime = Bars.LastTickDateTime
        End If
    End If
    
    ' convert the date/time if necessary
    If (bLocalTime = True) And (dDateTime <> Int(dDateTime)) Then
        If Bars Is Nothing Then
            Set Bars = New cGdBars
            SetBarProperties Bars, vSymbolOrSymbolID
        End If
        dDateTime = ConvertTimeZone(dDateTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
    End If

ErrExit:
    LastKnownPrice = dPrice
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.LastKnownPrice"
End Function

Public Sub ClearCache(Optional ByVal strSymbol$ = "", Optional ByVal bAndRemoveTickBuffer As Boolean = False)
On Error GoTo ErrSection:

    Dim idx&
    
    If m.bSalmonStarted Then Exit Sub
    
    strSymbol = RollSymbolForDate(UCase(Trim(strSymbol)))
    If Len(strSymbol) > 0 Then
        ' clear tick buffer and cache file for just this symbol
        KillFile App.Path & "\Cache\" & strSymbol & ".TBC" ' (TickBufferCache file)
        If bAndRemoveTickBuffer Then
            idx = m.TickBuffers.Index("^" & strSymbol)
            If idx > 0 Then
                m.TickBuffers.Remove idx
                m.TickBufferSymbols.Remove idx
            End If
        End If
    Else
        ' clear all generic tick buffers and cache files
        KillFile App.Path & "\Cache\*.TBC" ' (TickBufferCache file)
        If bAndRemoveTickBuffer Then
            For idx = m.TickBuffers.Count To 1 Step -1
                If Left(m.TickBuffers.Key(idx), 1) = "^" Then
                    m.TickBuffers.Remove idx
                    m.TickBufferSymbols.Remove idx
                End If
            Next
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealTime.ClearCache"
End Sub

Private Function LQ(ByVal LastQuote As eLastQuote) As eLastQuote
    LQ = LastQuote
End Function

Public Function DumpTickBufferInfo() As String
On Error GoTo ErrSection:

    Dim i&, strText$, hBars&, hString&, strPeriod$
    Dim tb As cGdBars
    
    For i = 1 To m.TickBuffers.Count
        strText = strText & vbCrLf & Str(i) & ", " & m.TickBufferSymbols(i) & ", " & m.TickBuffers.Key(i)
        Set tb = m.TickBuffers(i)
        If Not tb Is Nothing Then
            hBars = 0
            strPeriod = ""
            If Left(m.TickBuffers.Key(i), 1) <> "^" Then
                hBars = Val(m.TickBuffers.Key(i))
                If hBars <> 0 Then
                    hString = gdGetBarsStrProp(hBars, eBARS_PeriodicityStr)
                    If hString <> 0 Then
                        strPeriod = gdGetStr(hString)
                        gdDestroyString hString
                    End If
                End If
            End If
            strText = strText & ", tbSize=" & Str(tb.Size) & ", hBarsSize=" _
                & gdGetSize(hBars) & ", " & strPeriod
        End If
    Next

ErrExit:
    DumpTickBufferInfo = strText
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.DumpTickBufferInfo"
End Function

Public Property Get SplitterIP() As String
    SplitterIP = Parse(m.strIPandKey, vbTab, 1)
End Property

Public Property Get SplitterKey() As String
    SplitterKey = Parse(m.strIPandKey, vbTab, 2)
End Property

Public Function GetRecentSymbols() As cGdArray
    Set GetRecentSymbols = m.RecentSymbols.MakeCopy
End Function

Public Property Get ConvertContinuous() As Boolean
    ConvertContinuous = True 'm.bConvertCont
End Property

Public Property Get GenRtConnected() As Boolean
    If m.nServerStatus <> 0 Then
        GenRtConnected = True
    End If
End Property

Private Sub SetTickBufferStatus(ByVal strSymbol$, Optional ByVal strStatus$ = "<RELOAD>")
On Error GoTo ErrSection:

    Dim iPos&
    Dim TickBuffer As cGdBars

    strSymbol = RollSymbolForDate(strSymbol)

    ' set flag to reload data (since delay or something has changed)
    If m.TickBufferSymbols.BinarySearch(strSymbol, iPos) Then
        For iPos = iPos To m.TickBufferSymbols.Size - 1
            If m.TickBufferSymbols(iPos) <> strSymbol Then Exit For
            Set TickBuffer = m.TickBuffers(iPos)
            If Not TickBuffer Is Nothing Then
                TickBuffer.Prop(eBARS_CustomString) = strStatus
            End If
        Next
    End If

ErrExit:
    Set TickBuffer = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cRealtime.SetTickBufferStatus"
End Sub


'================== NEW SALMON STUFF BELOW ===========================

Public Property Get UseNextGen() As Boolean
On Error GoTo ErrSection:
    
    Dim strKey$, i&
    
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    i = GetRegistryValue(rkLocalMachine, strKey, "UseNextGen", -99)
    If i < -1 Then
        ' default the setting if it has not yet been overridden by the user
        If 1 Then
            i = True
        ElseIf HasModule("OPTNAV,E_US,BTX") Then
            i = True
        ElseIf FileExist(App.Path & "\UseSalmon.flg") Then
            i = True
        Else
            i = False
        End If
    End If
    If i = 0 Then
        UseNextGen = False
    Else
        UseNextGen = True
    End If

#If 0 Then
    If Not g.RealTime.Active Then
        UseNextGen = False
    ElseIf m.bSalmonStarted Then
        UseNextGen = True
    Else
        UseNextGen = False
    End If
#End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cRealtime.UseNextGen.Get"
End Property

Public Property Let UseNextGen(ByVal bUseNextGen As Boolean)
On Error GoTo ErrSection:
    
    Dim strKey$
    
    ' set in registry if user has overridden the normal default
    If bUseNextGen <> UseNextGen Then
        strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
        SetRegistryValue rkLocalMachine, strKey, "UseNextGen", Int(bUseNextGen)
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cRealtime.UseNextGen.Let"
End Property

Public Sub SalmonSetWindow()
If Not IsIDE Then
    On Error Resume Next
End If
    
    Dim i&
    Dim aStrings As New cGdArray
    aStrings(0) = Str(frmOnlineBroker.txtSalmonCallback.hWnd)
    ChangePath App.Path
    i = SetWindowHandle_SalmonDLL(aStrings.ArrayHandle)
    Set aStrings = Nothing

End Sub

Public Function SalmonIsRunning() As Boolean
On Error GoTo ErrSection:
    
    If m.bSalmonStarted Then
        If g.nReplaySession > 0 Then
            SalmonStop
        Else
            SalmonIsRunning = True
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealtime.SalmonIsRunning"
End Function

Public Sub SalmonStop()
On Error GoTo ErrSection:
    
    If m.bSalmonStarted Then
        m.bSalmonStarted = False
        Stop_SalmonDLL
        m.SymbolInfo.Clear
        ClearCache
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealtime.SalmonStop"
End Sub

' bForceRestart: can be True anytime we want to force calling Start_SalmonDLL again
' bResetAll: should be True only when restarting streaming (i.e. if it had been off)
Private Function SalmonStart(Optional ByVal bForceRestart As Boolean = False, _
                Optional ByVal bResetAll As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim i&
    Dim aStrings As cGdArray
    Static nPrevLDD&, nLenPrevAuthStr&
    
    If g.bUnloading Or g.nReplaySession > 0 Then Exit Function
    If UCase(m.strServerNode) <> "GENESISRT" Then Exit Function
    If Not UseNextGen Then Exit Function
        
    If bResetAll Or Not m.bSalmonStarted Or nPrevLDD <> LastDailyDownload _
            Or nLenPrevAuthStr <> Len(g.strAuthorizationString) Then
        bForceRestart = True
        nPrevLDD = LastDailyDownload
        nLenPrevAuthStr = Len(g.strAuthorizationString)
    End If
    
    If bForceRestart Then
        ChangePath App.Path
        If Not m.bSalmonStarted Then
            ClearCache
            m.SymbolInfo.Clear
            'm.RecentSymbols.Size = 0
        End If
        
        Set aStrings = New cGdArray
        'aStrings(0) = "10.1.191.76"
        aStrings(0) = Trim(FileToString(App.Path & "\Provided\Salmon.IP", , True))
        If Len(aStrings(0)) > 5 Then
            aStrings(1) = g.strAuthorizationString
            aStrings(2) = Str(frmOnlineBroker.txtSalmonCallback.hWnd)
            aStrings(3) = Str(LastDailyDownload)
            aStrings(4) = g.RealTime.SplitterIP
            aStrings(5) = g.RealTime.SplitterKey
            aStrings(6) = Str(RI_GetLastDataServiceID) ' TLB: should use Last Known DSID (to avoid security breach)
            aStrings(7) = "8" ' max requests to send at once in a batch
            aStrings(8) = "100" ' millisecond timeout to batch up requests
            aStrings(9) = Str(Abs(bResetAll)) ' when restarting realtime
            aStrings(10) = "60" '"300" ' seconds timeout to re-request history (in case a response never comes back)
            aStrings(11) = "300" ' seconds between re-request history of just the non-streamed symbols
            aStrings(12) = "3" ' # times to re-request history (e.g. if no response from Rex) before giving up and displaying just the streamed data (without the history)
            aStrings(13) = "TradeNav " & FormatVersion & ", b" & Str(App.Revision) & ", " & WindowsVersionStr
            aStrings(14) = RI_GetMachineID ' MachineID
            aStrings(15) = RI_GetUserPassword ' password
            m.bSalmonStarted = True
            i = Start_SalmonDLL(aStrings.ArrayHandle)
            If i <> 0 Then m.bSalmonStarted = False
            DebugLog "*** Start_SalmonDLL = " & Str(i) & vbCrLf & aStrings.JoinFields(vbCrLf)
            
            If 1 Then ' bResetAll Then
                ' just clear the SymbolInfo collection and let it all rebuild
                m.SymbolInfo.Clear
            Else
                ' reload all the Symbol Info objects
                For i = 1 To m.SymbolInfo.Count
                    m.SymbolInfo(i).SetSymbol ""
                Next
            End If
        End If
        Set aStrings = Nothing
    End If

    SalmonStart = m.bSalmonStarted

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealtime.SalmonStart"
End Function

Private Function SalmonInitNewStream(Optional ByVal bWaitForResponse As Boolean = True) As Boolean
On Error GoTo ErrSection:
    
    Dim dStartTime#, strErr$, nRC&
    Dim aStrings As New cGdArray
    
    If Not m.bSalmonStarted Then
        'SalmonStart
    End If
    If m.bSalmonStarted Then
        aStrings.Create eGDARRAY_Strings, 2
        dStartTime = gdTickCount
        Do While gdTickCount <= dStartTime + 30000 ' timeout
            nRC = StartStreaming_SalmonDLL(aStrings.ArrayHandle)
            Select Case nRC
            'Case eSalmonNotRequested
            'Case eSalmonPending
            Case eSalmonAvailable
                SalmonInitNewStream = True
                m.strIPandKey = aStrings(1) & vbTab & aStrings(0) ' IP <tab> Key
                dStartTime = gdTickCount - dStartTime
                Exit Do
            Case eSalmonError
                SalmonInitNewStream = False
                strErr = aStrings(0) ' Error message?
                Exit Do
            End Select
            If Not bWaitForResponse Then
                Exit Do
            End If
            Sleep 0.25
        Loop
    End If
    Set aStrings = Nothing
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealtime.SalmonInitNewStream"
End Function

' this routine gets called by the SalmonClient DLL via a txtWindow message (from frmOnlineBroker)
Public Sub SalmonCallback(ByVal strText$)
On Error GoTo ErrSection:

    Dim iCount&, nDataType&
    Dim nResponseStatus As eSalmonResponse
    Static strPrev$, iPrevCount&

    ' TLB: let this routine run when salmon is stopping (to send msg to OptNav)
    'If Not m.bSalmonStarted Then Exit Sub

    If strText <> strPrev Then
        strPrev = strText
        iCount = Val(Parse(strText, vbTab, 3))
        If iCount > 0 And iCount <> iPrevCount + 1 Then
            strText = "*** CALLBACK ERROR: " & Str(iPrevCount) & "  " & Str(iCount)
            DebugLog strText
            ''frmMain.Caption = strText
        End If
        iPrevCount = iCount
    
        If frmTest.Visible Then
            frmTest.AddList "callback = " & strText
        End If

        nDataType = Val(Parse(strText, vbTab, 2))
        nResponseStatus = Val(Parse(strText, vbTab, 4))
        If nDataType = 4 Then
            ' Salmon client wants to send an appmail to OptNav
            SendMessageToOptNav nResponseStatus, Parse(strText, vbTab, 1), True
        ElseIf nDataType = 3 Then
            m.nSalmonPort = Val(Parse(strText, vbTab, 1))
            frmTest.AddList "Port = " & Str(m.nSalmonPort)
        ElseIf nResponseStatus = eSalmonAvailable Then
            If m.bSalmonStarted Then
                ' if new data has arrived for a symbol,
                ' set flag in order to reload data
                'If nDataType > 0 And nDataType < 3 Then
                    SetTickBufferStatus Parse(strText, vbTab, 1) ' first field = symbol
                'End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealtime.SalmonCallback", eGDRaiseError_Show
End Sub


Public Function SymbolInfo(Symbol As Variant, Optional ByVal bRemoveFromStream As Boolean = False, _
                Optional ByVal bSkipAutoCreate As Boolean = False) As cSymbolInfo
On Error GoTo ErrSection:
    
    Dim nIdx&, strKey$
    Static nPrevIdx&, PrevSymbol As Variant
    
    ''If Not m.bSalmonStarted Then SalmonStart
    
    ' efficiency shortcut when getting same symbol consecutively
    If nPrevIdx > 0 And Not bRemoveFromStream Then
        If VarType(Symbol) = VarType(PrevSymbol) Then
            If Symbol = PrevSymbol Then
                Set SymbolInfo = m.SymbolInfo(nPrevIdx)
                If Not SymbolInfo Is Nothing Then ' (in case was removed)
If IsIDE Then
    If Len(SymbolInfo.Symbol) = 0 Then
        nIdx = nIdx
    End If
End If
                    Exit Function
                End If
            End If
        End If
    End If
       
    ' key is the symbol (individual contract is used for continuous)
    strKey = RollSymbolForDate(Symbol)
    nIdx = m.SymbolInfo.Index(strKey)
    If nIdx > 0 Then
        Set SymbolInfo = m.SymbolInfo(nIdx)
        If bRemoveFromStream Then
            SymbolInfo.RemoveFromStream
            m.SymbolInfo.Remove nIdx
            nIdx = 0
            Set SymbolInfo = Nothing
        End If
    ElseIf Not bRemoveFromStream And Not bSkipAutoCreate Then
        Set SymbolInfo = New cSymbolInfo
        SymbolInfo.SetSymbol strKey
        m.SymbolInfo.Add SymbolInfo, strKey
        nIdx = m.SymbolInfo.Count
    End If
    PrevSymbol = Symbol
    nPrevIdx = nIdx
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealtime.SymbolInfo"
End Function

' Get the newly streamed ticks for this symbol (since last time this was called)
' and append them to every TickBuffer for this symbol
Public Function SalmonGetNewTicks(ByVal strSymbol$) As Boolean
On Error GoTo ErrSection:

    Dim i&, j&, iPos&, strStatus$
    Dim Ticks As cGdBars, TickBuffer As cGdBars
    
    strSymbol = RollSymbolForDate(strSymbol)
    
    Set Ticks = SymbolInfo(strSymbol).GetNewlyStreamedTicks
    If Not Ticks Is Nothing Then
        If Ticks.Size > 0 Then
            SalmonGetNewTicks = True
            ' look for each tickbuffer for this symbol
            If m.TickBufferSymbols.BinarySearch(strSymbol, iPos) Then
                For iPos = iPos To m.TickBufferSymbols.Size - 1
                    If m.TickBufferSymbols(iPos) <> strSymbol Then Exit For
                    Set TickBuffer = m.TickBuffers(iPos)
                    If Not TickBuffer Is Nothing Then
                        ' only use the intraday tick buffers that are "ready"
                        If Left(m.TickBuffers.Key(iPos), 1) = "^" Then
                            strStatus = "" ' to skip the generic tickbuffer (if exists)
                        Else
                            strStatus = TickBuffer.Prop(eBARS_CustomString)
                        End If
                        If strStatus = "<OK>" Then
                            ' append the new ticks
                            For i = 0 To Ticks.Size - 1
                                j = TickBuffer.Size
                                TickBuffer.Size = j + 1
                                TickBuffer(eBARS_DateTime, j) = Ticks(eBARS_DateTime, i)
                                TickBuffer(eBARS_Close, j) = Ticks(eBARS_Close, i)
                                TickBuffer(eBARS_Vol, j) = Ticks(eBARS_Vol, i)
                                TickBuffer(eBARS_Flags, j) = Ticks(eBARS_Flags, i)
                            Next
                            ' only need to set the Exchange for the last trade
                            TickBuffer.Prop(eBARS_Exchange) = Ticks.Prop(eBARS_Exchange)
                        End If
                    End If
                Next
                Set TickBuffer = Nothing
            End If
        End If
    End If

ErrExit:
    Set Ticks = Nothing
    Exit Function
    
ErrSection:
    RaiseError "cRealtime.SalmonGetNewTicks"
End Function


Private Sub RequestInitialSymbolBatch()
On Error GoTo ErrSection:

    Dim i&, iRex&, strSymbol$, dLastBatchSent#
    Dim aSymbols As cGdArray
    Dim aRex1 As New cGdArray
    Dim aRex2 As New cGdArray
    Dim aRex3 As New cGdArray
    
    If Not SalmonIsRunning Then Exit Sub
    
    ' get list of symbols (load visible symbols first)
    Set aSymbols = frmQuotes.GetVisibleSymbols
    aSymbols.AppendFromArray m.UniqueSymbols
    
    ' break up list by data type based on Rex servers
    ' (1=Indices and Stocks, 2=PFG, 3=everything else)
    For i = 0 To aSymbols.Size - 1
        strSymbol = UCase(Parse(aSymbols(i), vbTab, 1))
        Select Case SecurityType(strSymbol, True)
        Case "S"
            aRex1.Add strSymbol
        Case "F", "FO", "IO", "SO"
            aRex3.Add strSymbol
        Case "M"
            ' don't add mutual funds
        Case Else
            If Right(strSymbol, 4) = "@PFG" Then
                aRex2.Add strSymbol
            ElseIf IsForex(strSymbol) Then
                aRex3.Add strSymbol
            Else
                aRex1.Add strSymbol
            End If
        End Select
    Next
    
    ' now request EOD data for symbols based on data type
    For iRex = 3 To 1 Step -1
        Select Case iRex
        Case 1
            Set aSymbols = aRex1
        Case 2
            Set aSymbols = aRex2
        Case 3
            Set aSymbols = aRex3
        End Select
        If aSymbols.Size > 0 Then
            Do While gdTickCount < dLastBatchSent + 100
                ' wait for last data type to get sent
            Loop
            For i = 0 To aSymbols.Size - 1
                SymbolInfo(aSymbols(i)).GetDataRequestStatus 0
            Next
            dLastBatchSent = gdTickCount
        End If
    Next
    
    Set aSymbols = Nothing
    Set aRex1 = Nothing
    Set aRex2 = Nothing
    Set aRex3 = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRealtime.RequestInitialSymbolBatch"
End Sub

Public Property Get DisconnectReason() As String
    DisconnectReason = m.strDisconnectReason
End Property

Public Property Let DisconnectReason(ByVal strReason As String)
    ' only store the first reason given
    If Len(m.strDisconnectReason) = 0 Then
        m.strDisconnectReason = strReason
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RtDataAvailable
'' Description: Determine if streaming data is available for the given symbol
'' Inputs:      Symbol, Period, Error?
'' Returns:     True if streaming data available, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RtDataAvailable(ByVal vSymbolOrSymbolID As Variant, ByVal lPeriodicity As Long, Optional bError As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nStatus As eSalmonResponse      ' Salmon response

    bReturn = False
    bError = False
    If ConnectionStatus = eGDConnectionStatus_Connected Then
        If SalmonIsRunning = True Then
            nStatus = SymbolInfo(vSymbolOrSymbolID).GetDataRequestStatus(lPeriodicity)
            
            If nStatus = eSalmonAvailable Then
                bReturn = True
            ElseIf nStatus = eSalmonError Then
                bError = True
            End If
        Else
            bReturn = True
        End If
    End If

    RtDataAvailable = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealtime.RtDataAvailable"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RtDataAvailableForBars
'' Description: Determine if streaming data is available for the given bars
'' Inputs:      Bars, Error?
'' Returns:     True if streaming data available, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RtDataAvailableForBars(ByVal Bars As cGdBars, Optional bError As Boolean) As Boolean
On Error GoTo ErrSection:

    RtDataAvailableForBars = RtDataAvailable(Bars.SymbolOrSymbolID, Bars.Prop(eBARS_Periodicity), bError)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRealTime.RtDataAvailableForBars"

End Function

