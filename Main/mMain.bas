Attribute VB_Name = "mMain"
Option Explicit

'S8=2 (2 seconds for comma), M1 (Sound ON), L3 (Sound High)
Public Const MODEM_INIT_DEFAULT As String = "AT S8=2 M1 L3"
Public Const SORT_BY_PREFIX As String = "Click to sort by: "

Public Const kFlagCol = 0
Public Const kSymbolCol = 1

Public Enum eTradeNavLevels
    eTN0_Unknown = 0
    eTN1_Silver = 1
    eTN2_Lite = 2
    eTN3_Standard = 3
    eTN4_Gold = 4
    eTN5_Professional = 5
    eTN6_Platinum = 6
End Enum

Public Enum eDockStates
    eShowAsPrevious = -1
    eHidden = 0
    eDocked = 1
    eUndocked = 2
End Enum

Private Enum eFormTypes
    eOther = 0
    eSymGrid = 1
    eChart = 2
    eFundamental = 3
End Enum

Public Enum eFutureSymbolType
    ePrimarySymbol = 0 ' returns Pit if exists, else returns Electronic
    ePitSymbol = 1
    eElectronicSymbol = 2 ' (any symbol not in SymbolMap.csv assumed to be electronic)
    eSyntheticSymbol = 3
    eCombinedSymbol = 4
End Enum

Public Enum eToolbarExtraInfo
    eTbExtraInfo_None = 0
    eTbExtraInfo_PFPNewPattern = 1
End Enum

'RH
Public Enum eStyleColorTypes
    'form
    eForm_Background = 0
    
    'frame
    eFrame_Background = 1
    eFrame_Border = 2
    
    'button
    eButton_Background = 3
    eButton_Border = 4
    eButton_Text = 5
    
    'checkbox
    eCheck_Border = 6
    eCheck_Background = 7
    eCheck_Forecolor = 8
    
    'flexgrid
    eGrid_Background = 9
    
End Enum

Public Type gGlobalStruct

    'RH
    Styler As New cStyler
    
    eTradeNavLevel As eTradeNavLevels
    
    WrkJet As Workspace
    dbNav As Database

    Functions As New cFunctions
    CommonBridge As New cCommonBridge
    'CommonMMBridge As New cCommonMMBridge

    bDirtyFunctionLibrary As Boolean
    bDirtyLibrariesMDB As Boolean   ' if saved function/rule/strategy so will know if should backup when shutdown
    bSkipMdbCompact As Boolean

    'Expression As cExpression
    'EditorOptions As cEditorOptions

    Universe As New cUniverse
    SymbolPool As New cSymbolPool
    RealTime As cRealTime
    
    bStarting As Boolean
    bUnloading As Boolean
    strRunWhenExit As String 'program to run when exiting (mostly for upgrading)
    
    dNextDownloadTry As Double
    dNextQuoteBoardRefresh As Double
    dLastQuoteBoardRefresh As Double
    dLastMouseActivity As Double
    
    iScrubLevel As Integer
    bShowInLocalTimeZone As Boolean
    
    nAltGridRowColor As Long
    nUpdatedColorDuration As Long
    
    strIniFile As String
    strAppPath As String
    strTitle As String

    strAuthorizationString As String
    strInstalledString As String

    DMS As DataMgrStruct
    SU As SymbolUniverseStruct

    ChartGlobals As ChartGlobalStruct
    strActiveDraw As String
    bDivAdjust As Boolean ' True = also adjust for dividends when adjusting for splits
    
    strChartPage As String
    bLoadingChartPage As Boolean
    bDirtyChartPage As Boolean
    bLoadPageTime As Boolean        'true=show time it took to load chart page
    bLoadPageOldMethod As Boolean   'true=use old method to load chart page
    
    bImportStrategyBaskets As Boolean   ' Do we need to import the strategy baskets into the database?
    
    ' TradeTracker:
    bDYKShown As Boolean
    'wsPaper As Workspace
    dbPaper As Database
    lAccountNum As Long
    
    lLCD As Long ' "last known" Customer ID (from last valid connection)

    'Supporting tables
    Coloring As cColoring
    Security As cSecurity

    RptBridge As cRptBridge
    
    CurrentSystem As cSystem

    ' 10/18/99 TLB: for frmMain to know which editor is active
    ActiveEditor As Editor
    
    ' gdTable versions of database tables
    tblLibrary As cGdTable
    tblFunction As cGdTable
    tblFunctionParm As cGdTable
    tblRule As cGdTable
    astrFunctionCategory As cGdArray
    
    Help As cHelp
    Alerts As cAlerts
    
    ' Online broker objects
    Broker As cBrokerDispatch           ' Broker dispatch object
    AdvFut As cBrokerTt                 ' Advantage Futures (Trading Technologies)
    AlpariCnx As cBrokerCurrenex        ' Alpari (Currenex)
    AlpariPats As cBrokerPats           ' Alpari (PATS)
    AlpariZenFire As cRithmic           ' Alpari (Zen-Fire)
    Ameritrade As cBrokerAmeritrade     ' TD Ameritrade
    AmpCqg As cBrokerCqg                ' AMP Futures (CQG)
    BornPats As cBrokerPats             ' Born Capital (PATS)
    CQG As cBrokerCqg                   ' CQG
    CtgCqg As cBrokerCqg                ' Capital Trading Group (CQG)
    CtgPats As cBrokerPats              ' Capital Trading Group (PATS)
'    CtgPfg As cPFG                      ' Capital Trading Group (PFG)
    Currenex As cBrokerCurrenex         ' Currenex
    DemoPats As cBrokerPats             ' PATS (Demo)
    Etrade As cBrokerEtrade             ' E-Trade
'    FintecPfg As cPFG                   ' Fintec Group (PFG)
    FptCqg As cBrokerCqg                ' Future Path Trading ( CQG )
    FptOec As cBrokerOec                ' Future Path Trading ( Open E-Cry )
    FxddCnx As cBrokerCurrenex          ' FXDD (Currenex)
    Gft As cBrokerGft                   ' GFT Forex
    Ideal As cIntBrokers                ' I-Deal (Interactive Brokers)
    IntBroker As cIntBrokers            ' Interactive Brokers
    KnightCnx As cBrokerCurrenex        ' Knight (Currenex)
    KnightCqg As cBrokerCqg             ' Knight (CQG)
'    LindWaldock As cXpress              ' Lind Waldock (LindXpress)
'    ManExpress As cXpress               ' Man Express (LindXpress)
    Oec As cBrokerOec                   ' Open E-Cry
    Optimus As cRithmic                 ' Optimus (Rithmic)
    OpVest As cRithmic                  ' OpVest (Rithmic)
    Pats As cBrokerPats                 ' PATS
'    PFG As cPFG                         ' PFG
    RcgPats As cBrokerPats              ' Rosenthal Collins Group (New PATS)
    Rithmic As cRithmic                 ' Rithmic
    RjoCqg As cBrokerCqg                ' RJ OBrien (CQG)
    RjoHkPats As cBrokerPats            ' RJ OBrien Hong Kong (PATS)
    RjoPats As cBrokerPats              ' RJ OBrien (PATS)
    RobbinsCqg As cBrokerCqg            ' Robbins (CQG)
    SimTradeTs As cSimTradeTs           ' Simulated Trading via the Trade Server
    SimTradeStream As cSimTradeStream   ' Simulated Trading via the stream
    SimTradeReplay As cSimTradeStream   ' Simulated Trading via streaming replay
    Tradier As cBrokerTradier           ' Tradier
    Transact As cTransact               ' TransAct
    TT As cBrokerTt                     ' Trading Technologies
    VanKarCnx As cBrokerCurrenex        ' VanKar Trading (Currenex)
    Vision As cRithmic                  ' Vision (Rithmic)
    VisionCqg As cBrokerCqg             ' Vision (CQG)
    ZanerCnx As cBrokerCurrenex         ' Zaner (Currenex)
    ZanerCqg As cBrokerCqg              ' Zaner (CQG)
    ZanerPats As cBrokerPats            ' Zaner (PATS)
    ZanerRithmic As cRithmic            ' Zaner (Rithmic)
    ZanerZenFire As cRithmic            ' Zaner (Zen-Fire)
    ZenFire As cRithmic                 ' Zen-Fire (Rithmic)
    
    CoreBridge As cCoreBridge           ' Core bridge object
    TnCore As cTnCore                   ' Application side bridge for common Navigator Suite routines
    
    CattleBridge As cCattleBridge       ' Cattle bridge object
    TnCattle As cTnCattle               ' Application side bridge for the Cattle DLL
    BrokerBridge As cBrokerBridge       ' Broker bridge object
    BrokerEnums As cBrokerEnums         ' Broker enumerations object
    TnBroker As cTnBroker               ' Application side bridge for the Broker DLL
    
    JournalBridge As cJournalBridge     ' Journal bridge object
    TnJournal As cTnJournal             ' Application side bridge for the Journal DLL
    
    nOptNavStatus As eGDOptNavStatus    ' Option Navigator status
    
    Profit As cProfit                   ' Global object for calculating profit/loss information
'    OrderLinks As cOrderLinks           ' Global collection of order links

    FractZen As cFractZen               ' Global object for FractZen stuff

    bSkipSetChartFocus As Boolean
    
    nRecalcIndRT As Long    'when streaming, 0:new tick, -1:new bar, >0:# seconds
    
    nReplaySession As Long  'for streaming replay: date of session to replay (0=off)
    nReplayAccountID As Long 'the trading account to use during streaming replay
    
    bShowRecalcMsg As Boolean
    nNumVerifiedRecalcs As Long
    
    eTbSkin As eTbSkin                  'background skin on toolbar
    vbeTbAlignDraw As Long              'VB enumeration for picturebox alignment propert
    nTbLargeIcons As Long               'JM 12-21-2010: 1=use 32x32 icons
    nTbIncludeText As Long              'JM 12-21-2010: 1=include text on toolbar buttons
    nTbIconStyle As Long                'JM 10-01-2015: 0=classic, 1=chrome
    nColorTheme As Long                 'JM 10-01-2015: RGB value depending on theme
    bPatProfitFlag As Boolean           'JM 04-12-2010: if true then allow right-click to open old PFP form
    
    bUsePitSettlesForDeltas As Boolean
    
    ConsoleForms As cTradeConsoleForms  ' Collection of trade console forms
    ActivityLogs As cActivityLogs       ' Collection of activity log objects
    OrderStrategies As cOrderStrategies ' Collection of active order strategies
    TradingItems As cAutoTradeItems     ' Collection of automated trading items
    FlattenQueue As cFlattenQueue       ' Collection of items to be flattened
    CondOrders As cConditionalOrders    ' Collection of conditional orders
    ExitAllOrders As cExitAllOrders     ' Collection of "exit all" orders
    TsoGroups As cActiveTsOrderGroups   ' Collection of trade sense order groups

    nLastWebReportID As Long ' increments each time a TNWebReport is run
    bShowAlertMsgForm As Boolean
    bHideAutoBreakoutNumber As Boolean  ' just for John Needham (when at seminars)
    bPageHasEWILabels As Boolean        ' 6926

    ChartPageCache As cGdTree
    FtpDownloader As New cDownloader
    
    iLanguage As Long ' 0 = English, 1 = German
End Type
Public g As gGlobalStruct

Private Type mPrivate
    dPrevActiveFormTime As Double
    PrevActiveForm As Form
    ActiveChartForm As Form
    bChartTimers As Boolean
    'iExtremeChartsMode As Integer ' 2 = Rule1U, 1 = BetterTrades, 0 = not either
    iExtremeChartsMode As Integer ' 2 = Advanced, 1 = Basic, 0 = is not ExtremeCharts
    IrxBars As cGdBars ' to store daily $IRX for reports (only need to retrieve once per day)
    
    TimerStarts As cGdTree
End Type
Private m As mPrivate

Private Declare Function Shutdown_SalmonDLL Lib "SalmonClient.dll" Alias "Shutdown" () As Long
Private Declare Sub DumpSymbolState_SalmonDLL Lib "SalmonClient.dll" Alias "DumpSymbolState" ()

' before calling, call Start as normal
' parameters[0] - type (0|1|2) (daily|minutes|fullTick)
' parameters[1] - start date (ASCII) YYYYMMDD
' parameters[2] - stop date (ASCII) YYYYMMDD
' parameters[3] - flags (currently undefined)
' tradeNavBars - contains the symbol to retreive; salmon client will
' populate individual arrays as needed.
' return value will be >= 0 on success and negative on failure
Public Declare Function GetSalmonHistory Lib "SalmonClient.dll" Alias "GetHistory" (ByVal hStrParms&, ByVal hBars&) As Long

Public Sub Main()
On Error GoTo ErrSection:

    Dim i&, d#, s$, strTemp$, dFileDate#, strMDB$, strCmd$, strFile$, strRegKey$, strBackup$
    Dim nSaveMainTop&
    Dim bHidden As Boolean, bEmpty As Boolean
    Dim bReset As Boolean
    Dim aStrings As New cGdArray
    Dim aUndocked As New cGdArray
    Dim SymbolGroup As cSymbolGroup
    Dim frm As Form
    Dim DbUpdates As cDatabaseUpdates   ' Database Updates class
    Dim TTUpdates As cTTUpdates
    Dim lCustID As Long                 ' Customer ID
    Dim strPKey As String               ' Product Key for System Navigator
    Dim rs As Recordset                 ' Recordset into the database
    Dim bValidDB As Boolean             ' Is this a valid Database?
    Dim lLCD As Long                    ' Last successful download customer ID
    Dim bResetCheckSums As Boolean
    Dim lPrevBuild As Long
    Dim lRestoreAttempt As Long         ' Number of restore attempts

    ' set current directory and store app path
    g.strAppPath = App.Path
    g.strIniFile = App.Path & "\ChartNavigator.INI"
    g.bStarting = True
    ChangePath App.Path
        
        
    ' exit if another instance already running
    If App.PrevInstance Then
        If Not bHidden Then ActivatePrevInstance
        End
    End If
    
    ' if an NVS machine needs to be rebooted, do it now (e.g. if restarting right after an Upgrade)
    If RebootNVSIfRequired Then
        End
    End If
    
    ' for Vista run the TradeNavStartup program (which will keep running)
    If IsAtLeastVista And Not IsIDE Then
        ' if Command args start with a ~, then this was started by the TradeNavStartup program
        If Left(Trim(Command$), 1) <> "~" Then
            If FileExist(App.Path & "\TradeNavStartup.exe") Then
                On Error Resume Next
                strTemp = AddSlash(App.Path) & App.EXEName & ".exe" & vbCrLf & "~" & Command$
                If 0 Then 'TLB 12/3/2015: don't disable Aero anymore
                    If FileExist(App.Path & "\AllowAero.flg") Then
                        strTemp = strTemp & vbCrLf & "AllowAero"
                    Else
                        strTemp = strTemp & vbCrLf & "DisableAero"
                        If AeroIsEnabled Then
                            FileFromString App.Path & "\DisablingAero.Now", "Aero is being disabled"
                        End If
                    End If
                End If
                FileFromString App.Path & "\TradeNavStartup.Run", strTemp, True
                If KillProcess("TradeNavStartup", True) = 0 Then
                    ' MUST use "ShellExecute" here instead of CreateProcess/RunProcess
                    ' (since needs to be able to elevate the called process to administrator)
                    ShellExecute ByVal 0&, "open", App.Path & "\TradeNavStartup.exe", "", App.Path, 0&
                End If
                End
            Else
                InfBox "TradeNavStartup.exe could not be found", "e", , "ERROR"
                End
            End If
        End If
    
        If 0 Then 'TLB 12/3/2015: don't disable Aero anymore
            If Not FileExist(App.Path & "\AllowAero.flg") Then
                ' Disable the Aero effects in Vista and Windows7 while TradeNav is running
                ' (since it causes both display and performance issues)
                EnableAero False
            End If
        End If
    End If
    
    ' TLB 4/9/2014: due to Kaspersky's bug last night, we can at least delete this file if it ever shows up
    ' (since a 0-byte System.MDB file really messes up Microsoft's JET engine!)
    strFile = AddSlash(App.Path) & "System.MDB"
    If FileLength(strFile) = 0 Then
        KillFile strFile, True
    End If
    
    If IsDBCS Then
        StartupLog "Started (DBCS windows)", 1
    Else
        StartupLog "Started", 1
    End If
       
    g.bHideAutoBreakoutNumber = GetIniFileProperty("HideAutoBreakout#", False, "General", g.strIniFile)
    
    If AllowDivAndMF Then
        g.bDivAdjust = GetIniFileProperty("DividendAdjust", True, "General", g.strIniFile)
    End If
    
    ' TLB 4/2010: new method for deltas -- use settle from the pit symbol's prior session
    If FileExist(App.Path & "\UsePitSettles.FLG") Then
        g.bUsePitSettlesForDeltas = True
    ElseIf FileLength(App.Path & "\Provided\UsePitSettles.NOT") > 5 Then
        g.bUsePitSettlesForDeltas = False
    Else
        g.bUsePitSettlesForDeltas = True
    End If
    
    ' Run the RegFiles.Bat if TradeNav was not previously shut down normally
    ' (at this point with Vista, we are now running in full Administrator mode)
    strFile = App.Path & "\..\SharedSelfReg\RegFiles.Bat"
    If FileExist(strFile) And Not FileExist(App.Path & "\SkipReg.flg") And Not Is9598orMe Then
        If Not FileExist(App.Path & "\Main\*.frm") Then
            StartupLog "Running RegFiles.Bat"
            frmRegFiles.Show 0
            RunProcess strFile, , True, vbHide, , , FilePath(strFile)
            Unload frmRegFiles
            StartupLog "Finished RegFiles.Bat"
        ElseIf 0 Then
            ' if just "debugging" for the programmers machines
            frmRegFiles.Show 0
            Sleep 5
            Unload frmRegFiles
        End If
    End If
    KillFile App.Path & "\SkipReg.flg"

    ' TLB: this property must be set BEFORE the HasDotNet is called
    SetIniFileProperty "ProgramPath32", AddSlash(App.Path), "GENERAL", WindowsPath() & "NavWin.INI"
    ' TLB 8/15/2014: and let's now verify upfront that the .NET Framework has been installed
    HasDotNet True
    
    ' process Command Line
    strCmd = Trim(Command$)
    If Left(strCmd, 1) = "~" Then strCmd = Trim(Mid(strCmd, 2))
    If UCase(strCmd) = ".CMD" Then
        strCmd = ""
        strTemp = App.Path & "\" & App.EXEName & ".CMD"
        If FileExist(strTemp) Then
            aStrings.FromFile strTemp
            strCmd = Trim(aStrings(0))
        End If
    End If
    If UCase(Left(strCmd, 6)) = "HIDDEN" Then
        bHidden = True
    ElseIf UCase(strCmd) = "EMPTY" Or UCase(strCmd) = "QUICK" Then
        bEmpty = True
    End If
    If InStr(UCase(strCmd), "RESET") > 0 Then
        bReset = True
    End If
                 
    ' try setting heap to "low frag" -- but if ever crashes trying, then leave flag file so will ignore from now on
    strFile = App.Path & "\SkipSetLowFrag.flg"
    If Not FileExist(strFile) Then
        FileFromString strFile, "x"
        SetLowFragHeap
        KillFile strFile '(must have worked, so can kill the file)
    End If
                 
    ' make sure some of the external processes are not running
    KillOtherPrograms
                 
    ' clear read-only flags from files
    If FileExist(App.Path & "\ChartNav\*.frm") Then
        ' for developers (with source code), just clear certain things
        ClearReadOnlyFlags App.Path & "\Info\*.*"
        ClearReadOnlyFlags App.Path & "\TradeTracker.mdb"
        ClearReadOnlyFlags App.Path & "\FixedStats.txt"
        ClearReadOnlyFlags App.Path & "\IpAddr.gcl"
    Else
        ' clear read-only flags from all files (just safer to do all)
        aStrings.GetMatchingFiles App.Path & "\*.* /s /a=r"
        For i = 0 To aStrings.Size - 1
            ClearReadOnlyFlags aStrings(i)
        Next
    End If
    
    strRegKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    ALT_GRID_ROW_COLOR = GetColorFromString(GetProvidedProperty("AltGridRowColor", "&HC8F0FF"))
    'ALT_GRID_ROW_COLOR = GetRegistryValue(rkLocalMachine, strRegKey, "AltGridRowColor", &HC8F0FF)
'ALT_GRID_ROW_COLOR = &HE8F0E8
'ALT_GRID_ROW_COLOR = &HF0F4F0
'ALT_GRID_ROW_COLOR = &HD0F0FF
    lPrevBuild = GetRegistryValue(rkLocalMachine, strRegKey, "BuildNumber", 0)
    g.iScrubLevel = GetRegistryValue(rkLocalMachine, strRegKey, "ScrubLevel", 5&)
                 
    ' if first time: set auto-download time to a random time between 7pm-8pm NY
    If GetRegistryValue(rkLocalMachine, strRegKey, "AutoStart", -99.9) = -99.9 Then
        SetRegistryValue rkLocalMachine, strRegKey, "AutoStart", TimeSerial(19, RandomNum(0, 59), 0)
    End If
    
    ' allow reset to clear out some things
    If bReset Then
        SetIniFileProperty "MDI_Placement", "", "Forms", g.strIniFile
        SetIniFileProperty "MDI_State", 0, "Forms", g.strIniFile
    End If
                                                     
    ' set ExtremeChartsMode as follows:
    ' - if cmd-line arg has ETA, then BetterTrades
    ' - else if cmd-line arg has R1U or source code is Rule1, then Rule 1
    ' - else if has ETA program, then BetterTrades
    If InStr(UCase(strCmd), "ETA") > 0 Then
        m.iExtremeChartsMode = 1
    ElseIf InStr(UCase(strCmd), "R1U") > 0 Or GetSourceCode = "R1U" Then
        m.iExtremeChartsMode = 1 '2 --> TLB 2/19/2014: R1U no longer supported
    ElseIf FileDate(App.Path & "\..\Eta\Eta.exe") > DateSerial(2005, 10, 1) Then
        m.iExtremeChartsMode = 1
    End If
    ' - but if HasGold (but without the special flag file -- for backwards-compatibility?)
    '   then it's not Extreme Charts at all (neither BetterTrades nor R1U)
    If HasGold(False, , False) And Not FileExist(App.Path & "\Extreme.flg") Then
        m.iExtremeChartsMode = 0
    End If
    
    ' TLB 2/19/2014: for new "Extreme Charts Advanced" demo
    If HasModule("BTXA") Then
        m.iExtremeChartsMode = 2
    End If
    
    ' create the menu and desktop shortcuts now (only if a newer build)
    If App.Revision > lPrevBuild Then
        CreateShortcuts
    End If
    
    Set m.TimerStarts = New cGdTree
    
    ' This should be initialized before frmMain is loaded because the frmMain.Load calls Picture16...
    Set g.CoreBridge = New cCoreBridge
Set g.TnCore = New cTnCore
                                                                      
    ' make sure MDI parent is first form loaded at run-time
    StartupLog "Loading frmMain"
    Load frmMain
    StartupLog "frmMain loaded"
              
    'Show splash form
    frmSplash.Show
    DoEvents
    frmSplash.Message 0, "Initializing ..."
    StartupLog "Splash shown"
    SetMainCaption
   
    Set g.FtpDownloader = New cDownloader
  
    ' display info message about disabling Aero
    If FileExist(App.Path & "\DisablingAero.Now") Then
        KillFile App.Path & "\DisablingAero.Now"
        If 0 Then ' IsAtLeastVista Then
            If GetIniFileProperty("AeroMessage", 0, "DontAsk", g.strIniFile) = 0 Then
                strTemp = "The Windows Aero 'glass' effect is being temporarily disabled in order to improve graphics performance and to allow compatibility with some special display features within this application."
                strTemp = InfBox(strTemp, "i", "OK", "Performance Setting", , , , , , , , , True)
                If Right(strTemp, 1) = "-" Then
                    SetIniFileProperty "AeroMessage", 1, "DontAsk", g.strIniFile
                End If
            End If
        End If
    End If
  
    'Make sure required directories exist
    MakeDir App.Path & "\Cache", False
    MakeDir App.Path & "\RTS", False
    MakeDir App.Path & "\Stream", False
    MakeDir App.Path & "\Trades", False
    MakeDir App.Path & "\GameResults", False
    MakeDir App.Path & "\SimTrade", False
    MakeDir App.Path & "\Chk", False
    MakeDir App.Path & "\Info\Remote\", False
    MakeDir App.Path & "\Backup\", False
    MakeDir App.Path & "\Provided\", False
    MakeDir App.Path & "\Custom\", False
    MakeDir App.Path & "\Charts\Templates\", False
    MakeDir App.Path & "\Charts\Pages\", False
    MakeDir App.Path & "\Ftp\Backup\", False
    MakeDir App.Path & "\Ftp\Dist\", False
    MakeDir AddSlash(App.Path) & "..\LibraryDLLs", False
    MakeDir AddSlash(App.Path) & "Help", False
    MakeDir AddSlash(App.Path) & "QBT", False
    MakeDir AddSlash(App.Path) & "SavedImages", False
        
    ' and kill obsolete files
    KillFile App.Path & "\News.txt"
    KillFile App.Path & "\Fileinfo.*"
    KillFile App.Path & "\HotKeys.txt"
    KillFile App.Path & "\chk\Coded.chk"
    
    ' ONLY ON INSTALLS MACHINE: see if need to re-encrypt the symbol filters
    If IsIDE Then
        If FileDate(App.Path & "\Provided\SymbolFilter.TXT") > FileDate(App.Path & "\Provided\SymbolFilter.CFG") Then
            strTemp = FileToString(App.Path & "\Provided\SymbolFilter.TXT")
            strTemp = EncryptToHex(strTemp)
            FileFromString App.Path & "\Provided\SymbolFilter.CFG", strTemp, False
        End If
    End If
    
    ' copy any newer Remote files over (this is done in case the GRemote is running when the
    ' upgrade was done -- the new GRemote files get put into the RemoteNew folder and copied over now)
    If FileDate(App.Path & "\Info\RemoteNew\GRemote.exe") > FileDate(App.Path & "\Info\Remote\GRemote.exe") Then
        On Error Resume Next
        'fs.CopyFile App.Path & "\Info\RemoteNew\*", App.Path & "\Info\Remote\", True
        CopyFiles App.Path & "\Info\RemoteNew\*", App.Path & "\Info\Remote\", True
        On Error GoTo ErrSection:
    End If
    ' fix initial Remote Assist settings
    'station=101
    SetIniFileProperty "inputs", "1", "Settings", App.Path & "\Info\Remote\Remote.INI"
    SetIniFileProperty "clipboard", "1", "Settings", App.Path & "\Info\Remote\Remote.INI"
    If ExtremeCharts >= 1 Then
        strTemp = GetProvidedProperty("CompanyName", , True)
    Else
        strTemp = ""
    End If
    SetIniFileProperty "company", strTemp, "Settings", App.Path & "\Info\Remote\Remote.INI"
    strTemp = GetProvidedProperty("RemoteAssist", "tech.genesisft.com", True)
    SetIniFileProperty "host", strTemp, "Settings", App.Path & "\Info\Remote\Remote.INI"
    
    ' clear out gclient's log (if exists) -- TLB 1/11/2007: now always create this for everybody
    ''If FileExist(App.Path & "\Gclient_.txt") Then
        FileFromString App.Path & "\Gclient_.txt", "Debug log for GClient ..." & vbCrLf
    ''End If
      
    'Check for PROVIDED.GZP and CUSTOM.GZP, etc. (but NOT yet news and msg)
    CheckForSpecialDownloadFiles App.Path & "\Ftp\", False
    
    'Check for new SymTran files to copy into data area
    strFile = DataPath & "SymTran\*.*"
    If FileExist(strFile) Then
        'fs.CopyFile strFile, DataPath, True
        'fs.DeleteFile strFile, True
        CopyFiles strFile, AddSlash(DataPath), True
        DeleteFiles strFile, True
    End If

    ' set # milliseconds for duration of updated color (blue) when in real-time
    g.nUpdatedColorDuration = 2000

    'Get authorization string
    GetAuthorizationStringFromRegistry
    
    'See if newer data from recent install
    If NewerCdDataExists Then
        ' throw away the current downloading data
        ' (make the user install the new data)
        If FileExist(DataPath & "*.*") Then
            If 1 Then
                ' temporarily move them into an OLD directory
                On Error Resume Next
                strFile = DataPath
                MakeDir strFile & "Old"
                KillFile strFile & "Old\*.*"
                'fs.MoveFile strFile & "*.*", strFile & "Old\"
                MoveFiles strFile & "*.*", strFile & "Old\"
                On Error GoTo ErrSection
            Else
                strFile = App.Path & "\Backup\OldData.GZP"
                KillFile strFile
                'SplashMessage "Backing up old data ..."
                'ZipExecute "Z", strFile, App.Path & "\Data\"
            End If
            KillFile DataPath & "*.*"
        End If
    End If
    strTemp = DataPath
    If Not FileExist(strTemp & "*.dbf") Then
        MakeDir strTemp, False
        ' if need to init data, cancel "HIDDEN" mode
        bHidden = False
    End If
    
    ' remove path from data table refs for tables in the Data folder
    FixMasterDataTablePaths

    'JM 11-16-2015: need theme color to set toolbar menu icons etc
    strTemp = GetIniFileProperty("TradenavTheme", "", "General", g.strIniFile)
    If Len(strTemp) = 0 Then
        ' theme has never been set, so prompt for theme
        If IsAtLeastVista Then
            strTemp = frmTheme.ShowMe(True)
        Else
            strTemp = "Classic"
        End If
    End If
    Select Case strTemp
        Case "Classic"
            g.nColorTheme = 0  'this will get changed in call to SetTheAppBackColor below
            g.nTbIconStyle = 0
            Picture16 "Theme=Classic"
        Case "Charcoal"
            g.nColorTheme = kDarkThemeColor
            g.nTbIconStyle = 1
            Picture16 "Theme=Charcoal"
            ALT_GRID_ROW_COLOR = GetColorFromString(GetProvidedProperty("AltGridRowColorCharcoal", Str(RGB(65, 65, 65))))
        Case "Ivory"
            g.nColorTheme = vbWhite
            g.nTbIconStyle = 1
            Picture16 "Theme=Ivory"
        Case Else
            ' check if AppBackColor is being supplied by main INI file
            g.nColorTheme = GetColorFromString(GetProvidedProperty("AppBackColor", ""))
            If g.nColorTheme > 0 Then g.nColorTheme = Abs(Val(Parse(GetProvidedProperty("AppBackColor", ""), ";", 2)))
    End Select
    ' get these INI properties after the initial "theme" is set at startup
    g.eTbSkin = GetIniFileProperty("ToolbarSkin", eTbSkin_Silver, "Toolbars", g.strIniFile)
    g.vbeTbAlignDraw = GetIniFileProperty("ToolbarAlignDraw", vbAlignRight, "Toolbars", g.strIniFile)
    
    ' load chart global properties from INI file
    LoadChartGlobals
    
    ' load default available tool buttons that can be used by any form
    BtnConfigLoad App.Path & "\Provided\Toolbuttons.cfg"
        
    lRestoreAttempt = 0&
            
    'Make sure Libraries.MDB exists
MDBCopy:
    strMDB = App.Path & "\Libraries.MDB"
    '(first check for possibility of rename not working after last compact)
    strTemp = App.Path + "\_Temp.mdb"
    If FileExist(strMDB) Then
        KillFile strTemp
    ElseIf FileExist(strTemp) Then
        Name strTemp As strMDB
        strTemp = "_Temp.mdb renamed: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        FileFromString App.Path & "\CheckMDB.log", strTemp, True, True
    End If
    If Not FileExist(strMDB) Then
        ' kill an existing Market.MDB since might be way too old
        KillFile App.Path & "\Market.MDB"
        
        ' If not exist, try to copy over from SysNav directory first
        If FileExist(App.Path & "\..\System Navigator\Navigator.mdb") Then
            Select Case InfBox("Do you wish to convert your systems, rules and functions from the System Navigator product,|or start with a new libraries database?", "?", "+Convert|New|-Abort", "Libraries Database")
            Case "A" 'Abort
                End
            Case "N" 'New
                If InfBox("NOTE:  After upgrading to Trade Navigator|using a new set of libraries, please do not continue using the System Navigator product.", "i", "+Continue|-Abort", "Confirmation") = "A" Then
                    End
                End If
            Case Else 'Convert
                If InfBox("NOTE:  After converting your libraries to|Trade Navigator, please do not continue|using the System Navigator product.", "i", "+Continue|-Abort", "Convert Libraries") = "A" Then
                    End
                End If
                FileCopy App.Path & "\..\System Navigator\Navigator.mdb", App.Path & "\Libraries.mdb"
                'FileCopy App.Path & "\..\System Navigator\Market.mdb", App.Path & "\Market.mdb"
            End Select
        End If
        ' If not in SysNav directory, unzip it from the NavStart
        If Not FileExist(strMDB) And FileExist(App.Path & "\NavStart.GZP") Then
            ZipExecute "U", App.Path & "\NavStart.GZP", App.Path, "Libr*.mdb"
            'ZipExecute "U", App.Path & "\NavStart.GZP", App.Path, "Genes*.gzp"
        End If
        ' If not found anywhere, then must quit
        If Not FileExist(strMDB) Then
            InfBox "i=[] ; h=Program Error ; Cannot find database:  Libraries.MDB"
            End
        End If
        
        ' If just unzipped from NavStart or copied from SysNav area, see if a
        ' GenesisLib.GZP exists (either in App path or in Backup folder) which
        ' is newer than the Libraries.MDB file -- if so, needs to be auto-imported.
        MakeDir App.Path & "\Backup", True
        strFile = App.Path & "\GenesisLib.GZP"
        strBackup = App.Path & "\Backup\GenesisLib.GZP"
        If FileDate(strFile) - 0.25 < FileDate(strMDB) Then
            ' delete file in App path if older so won't auto-import it
            ' (but if no backup exists yet, then move it to backup area
            ' -- so the first Upgrade will not auto-import unless is newer)
            If Not FileExist(strBackup) Then
                'fs.MoveFile strFile, strBackup
                MoveFiles strFile, strBackup
            Else
                KillFile strFile
            End If
        End If
        ' see if a newer copy exists in Backup area
        If FileDate(strBackup) - 0.25 > FileDate(strMDB) And FileDate(strBackup) > FileDate(strFile) Then
            ' move to App path so will auto-import
            'fs.CopyFile strBackup, strFile, True
            CopyFiles strBackup, strFile, True
            If FileExist(strFile) Then KillFile strBackup
        End If
    ElseIf FileExist(App.Path & "\GenesisLib.GZP") Then
        ' if about to do an auto-import, then make a backup copy first
        ' (unless the existing backup copy is less than a week old)
        strTemp = App.Path & "\Backup\BeforeAutoImport.mdb"
        If FileDate(strTemp) < Date - 7 Then
            KillFile strTemp
            'fs.CopyFile strMDB, strTemp
            CopyFiles strMDB, strTemp
        End If
    End If
    
    'If Not FileExist(App.Path & "\Market.MDB") Then
    '    ZipExecute "U", App.Path & "\NavStart.GZP", App.Path, "Market.MDB"
    'End If
    If Not FileExist(App.Path & "\TradeTracker.MDB") Then
        ZipExecute "U", App.Path & "\NavStart.GZP", App.Path, "TradeT*.MDB"
    End If
    
    ' If the RestoreMDB.FLG file exists, then we need to copy OLD.MDB to
    ' Libraries.MDB because an import failed...
    If FileExist(AddSlash(App.Path) & "RestoreMDB.FLG") Then
        'fs.CopyFile AddSlash(App.Path) & "Old.MDB", AddSlash(App.Path) & "Libraries.MDB", True
        CopyFiles AddSlash(App.Path) & "Old.MDB", AddSlash(App.Path) & "Libraries.MDB", True
        KillFile AddSlash(App.Path) & "RestoreMDB.FLG"
    End If
       
    ' TLB 4/2/2013: unzip OptionNav.GZP (if exists, must have just been installed)
    strFile = App.Path & "\..\OptionNav\OptionNav.gzp"
    If FileExist(strFile) Then
        ZipExecute "U", strFile, App.Path & "\..\OptionNav\"
        KillFile strFile ' then delete it (so won't keep unzipping)
    End If
       
    frmSplash.Message 5
    
    If Not bEmpty Then
        
        ' open Libraries.MDB
        'On Error Resume Next
        Set g.WrkJet = CreateWorkspace("NavSuite", "admin", "", dbUseJet)
        If g.WrkJet Is Nothing Then
            Err.Raise vbObjectError + 1000, , "Error creating DAO workspace."
            End
        End If
        
        ' Update the gdSettings.dat if command line is "Settings"
        If UCase(strCmd) = "SETTINGS" Then
            Set g.dbNav = g.WrkJet.OpenDatabase(AddSlash(App.Path) & "Settings.MDB")
            BuildSettingsFile
            DoEvents
            g.dbNav.Close
            DoEvents
        End If
        
        ' TLB 3/22/2015: a 1-time check to reset the RiskBasedOn (MM settings) back to default
        If GetIniFileProperty("BuildChecked", 0, "275", App.Path & "\Reports.INI") < 1445 Then
            SetIniFileProperty "BuildChecked", App.Revision, "275", App.Path & "\Reports.INI"
            ' reset RiskBasedOn to default (empty string)
            SetIniFileProperty "RiskBasedOn", "", "275", App.Path & "\Reports.INI"
        End If
        
        ' kill rules table file if libraries mdb has changed
        If GetIniFileProperty("LastLib", "", "General", g.strIniFile) <> Str(FileDate(strMDB)) & ";" & Str(FileLength(strMDB)) Then
            KillRulesFile
        End If
        
        'DBEngine.Workspaces.Append g.WrkJet
        ''Set g.dbNav = g.WrkJet.OpenDatabase(strMDB, False, False, "; PWD=" & DbPassword)
        Set g.dbNav = OpenDatabase(strMDB, DbPassword)
        
        On Error GoTo ErrSection
        If g.dbNav Is Nothing Then
            Err.Raise vbObjectError + 1000, , "Could not open Libraries.MDB"
            End
        End If
        'If Not LinkTableToDb(g.dbNav, "tblMarkets", App.Path & "\Market.mdb") Then
        '    g.dbNav.Close
        '    End
        'End If
        
        ' Check to see if the database has indexes...
        If (CheckLibIndexes(Str(lRestoreAttempt)) = False) Then
            lRestoreAttempt = lRestoreAttempt + 1
            g.dbNav.Close
            g.WrkJet.Close
            
            If RestoreLibDatabase(lRestoreAttempt) Then
                GoTo MDBCopy
            Else
                End
            End If
        End If
        
        'open Paper trader MDB
        strFile = App.Path & "\TradeTracker.mdb"
        ''Set g.dbPaper = g.WrkJet.OpenDatabase(strFile, False, False)
        Set g.dbPaper = OpenDatabase(strFile, "")
        
        ' Check to see if the database has indexes...
        If (CheckTtIndexes(Str(lRestoreAttempt)) = False) Then
            lRestoreAttempt = lRestoreAttempt + 1
            g.dbPaper.Close
            g.dbNav.Close
            g.WrkJet.Close
            
            If RestoreTtDatabase(lRestoreAttempt) Then
                GoTo MDBCopy
            Else
                End
            End If
        End If
        
        StartupLog "Database opened"
        frmSplash.Message 10
        
        ' Verify TradeTracker.MDB is up to date
        Set TTUpdates = New cTTUpdates
        With TTUpdates
            .DB = g.dbPaper
            If .Upgrade Then StartupLog "TT database upgraded"
        End With
        Set TTUpdates = Nothing
    
        ' Verify Libraries.MDB is up to date BEFORE getting customer ID:
        ' just for upgrading purposes, we need to use whatever CustID
        ' was here in the first place (if this was a stolen database,
        ' we will upgrade, but still not allow it to be used)
        g.bImportStrategyBaskets = False
        Set DbUpdates = New cDatabaseUpdates
        With DbUpdates
            .DB = g.dbNav
            If .Upgrade Then
                StartupLog "Database upgraded"
                KillRulesFile '(force a rebuild of the rules table)
            End If
        End With
        Set DbUpdates = Nothing
            
        ' now get the last known customer ID (from download)
        ChangePath App.Path ' (TLB: need to insure RInfoDLL.DLL can get found and loaded on first call)
        lLCD = RI_GetLastDataServiceID \ 1000
        g.lLCD = lLCD
        s = Str(RI_GetDataServiceID) & "  " & Str(g.lLCD) & "  " & UCase(RI_GetMachineID) & "  " _
            & FormatVersion & ", b" & Str(App.Revision) & ", " & WindowsVersionStr
        DebugLog s & vbCrLf & g.strAuthorizationString

        ' Make sure that the CheckSums on the Libraries are valid...
        bValidDB = True
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblDatabase];", dbOpenDynaset)
        lCustID = rs!CID
        ' TLB 3/14/2007: to solve the Vista upgrade issue (due to "virtual registry")
        ' if g.lLCD is 0 (not in registry) but lCustID > 0 and an Upgrade exists in the FTP folder
        ' then let's be nice and get the CustID and password from the Ibis.txt file
        ' (otherwise they will lose their Libraries.mdb the next time they startup)
        If g.lLCD = 0 And lCustID > 0 Then
            If FileExist(App.Path & "\Ftp\Upgrd32.exe") Then
                aStrings.FromFile App.Path & "\Ftp\Ibis.txt"
                For i = 0 To aStrings.Size - 1
                    Select Case UCase(Parse(aStrings(i), "=", 1))
                    Case "DSRVID"
                        If Val(Parse(aStrings(i), "=", 2)) \ 1000 = lCustID Then
                            lLCD = lCustID
                            g.lLCD = lCustID
                            RI_SetDataServiceID Val(Parse(aStrings(i), "=", 2))
                        Else
                            Exit For ' don't set password if CustID does not match
                        End If
                    Case "PASSWORD"
                        If g.lLCD > 0 Then
                            RI_SetUserPassword Parse(aStrings(i), "=", 2)
                        End If
                    End Select
                Next
            End If
        End If
        ValidateCheckSums rs, "tblDatabase", lCustID
        If rs!CheckSum = 0.5 Then
            bValidDB = False '(user must have manually changed the CID in the MDB)
        Else
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys];", dbOpenDynaset)
            ValidateCheckSums rs, "tblLibrarys", lCustID
            If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
            Do While Not rs.EOF
                If rs!CheckSum = 0.5 Then
                    bValidDB = False
                    Exit Do
                End If
                rs.MoveNext
            Loop
            If bResetCheckSums And g.lLCD = 0 Then
                End
            End If
        End If
        
        StartupLog "Libraries validated"
        
        ' If the Libraries were valid, then make sure the Customer ID is correct...
        If bValidDB Then
            ' If no successful download yet, only allow the BuiltIn libraries...
            If lLCD = 0 Then
                If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
                Do While Not rs.EOF
                    rs.Edit
                    rs!Ignore = (Not (rs!BuiltIn)) And (rs!LibraryID <> kSN_UserLibrary)
                    rs!CheckSum = BuildCheckSum(rs, "tblLibrarys", lCustID)
                    rs.Update
                    
                    rs.MoveNext
                Loop
                
            ' If the Customer ID is blank in the database, put in LCD and update CheckSums...
            ElseIf lCustID = 0 Then
                Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblDatabase];", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then
                    rs.Edit
                    rs!CID = lLCD
                    rs!CheckSum = BuildCheckSum(rs, "tblDatabase")
                    rs.Update
                End If
                
                Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys];", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
                Do While Not rs.EOF
                    rs.Edit
                    rs!CheckSum = BuildCheckSum(rs, "tblLibrarys")
                    rs.Update
                    
                    rs.MoveNext
                Loop
            
            ' If the Customer ID's don't match, it is invalid
            ElseIf lCustID <> lLCD Then
                bValidDB = False
            End If
        End If
        rs.Close
        
        ' If the user is not authorized for this database, let them know about it...
        If Not bValidDB Then
            g.dbNav.Close
            If InfBox("You are not authorized for this libraries database.||Would you like to get a new copy|of the default libraries database?|", "!", "+Yes|-Abort", "Error") = "A" Then
                End
            End If
            ' move the "bad" file into the backup area
            If FileExist(strMDB) Then
                On Error Resume Next
                strTemp = App.Path & "\Backup\Unauthorized.mdb"
                KillFile strTemp
                Name strMDB As strTemp
                KillFile strMDB '(just in case still exists)
                On Error GoTo ErrSection
            End If
            g.dbPaper.Close
            Set g.dbPaper = Nothing
            GoTo MDBCopy
        End If
            
        ' If a GenesisLib.GZP exists, we need to auto-import library(s)
        If FileExist(App.Path & "\GenesisLib.GZP") Then
            ' see if same file as one already in the backup area
            If CalcFileCrc(App.Path & "\GenesisLib.GZP") = CalcFileCrc(App.Path & "\Backup\GenesisLib.GZP") Then
                'If InfBox("Do you wish to re-import the Genesis Libraries|(if unsure, please select 'Import')?", "?", "+Import|-Cancel", "Import Genesis Libraries") = "C" Then
                If InfBox("Do you need to re-import the Genesis Libraries?|(e.g. if the previous import had an error)", "?", "Re-Import|+-No", "Import Genesis Libraries") <> "R" Then
                    KillFile App.Path & "\GenesisLib.GZP"
                End If
            End If
        End If
        If FileExist(App.Path & "\GenesisLib.GZP") Then
            ' Make sure that the temporary library directory exists
            MakeDir App.Path & "\TempLib", True
            KillFile App.Path & "\TempLib\*.*", True
            
            ' Unzip the GenesisLib.GZP into the temporary library directory
            ZipExecute "U", App.Path & "\GenesisLib.GZP", App.Path & "\TempLib"
            
            ' Import all of the txt files in the GenesisLib.GZP file
            frmAutoImportLib.ShowMe App.Path & "\TempLib\GenesisLib.CFG"
            DoEvents
            
            ' Copy any necessary DLL's over to the application directory
            aStrings.GetMatchingFiles App.Path & "\TempLib\*.DLL", False
            For i = 0 To aStrings.Size - 1
                strFile = aStrings(i)
                KillFile App.Path & "\" & strFile, True
                'fs.CopyFile App.Path & "\TempLib\" & strFile, App.Path & "\" & strFile, True
                CopyFiles App.Path & "\TempLib\" & strFile, App.Path & "\" & strFile, True
            Next
            
            ' Copy the GenesisLib.GZP into the backup directory and delete it
            MakeDir App.Path & "\Backup", True
            'fs.CopyFile App.Path & "\GenesisLib.GZP", App.Path & "\Backup\GenesisLib.GZP", True
            CopyFiles App.Path & "\GenesisLib.GZP", App.Path & "\Backup\GenesisLib.GZP", True
            KillFile App.Path & "\GenesisLib.GZP"
            
            StartupLog "GenesisLib.GZP Imported"
            KillRulesFile '(force a rebuild of the rules table)
        End If
        
        ' If the AutoImport.CFG file exists, we need to auto-import the given library
        ' because it needs to replace a DLL...
        If FileExist(AddSlash(App.Path) & "AutoImport.CFG") Then
            frmAutoImportLib.ShowMe AddSlash(App.Path) & "AutoImport.CFG"
            DoEvents
            KillFile AddSlash(App.Path) & "AutoImport.CFG"
            
            StartupLog "AutoImport.CFG Imported"
            KillRulesFile '(force a rebuild of the rules table)
        End If
        
        ' Check to make sure that the database still has its indexes...
        If (CheckLibIndexes("after auto-import") = False) Then
            InfBox "The libraries database is corrupted.  The next time that you start Trade Navigator, an attempt will be made to restore a backup copy.", "!", , "Database Error"
            End
        End If
        
        ' Clean out the temporary library directory
        KillFolder App.Path & "\TempLib", True
        
        'Pass connection to Access through bridges...
        g.CommonBridge.AppPath = App.Path
        g.CommonBridge.CustomerID = g.lLCD
        g.CommonBridge.dbNavRef = g.dbNav
        
        InitializeBrokerBridge
        InitializeCattleBridge
        InitializeJournalBridge
        g.JournalBridge.AltGridRowColor = ALT_GRID_ROW_COLOR
        
        StartupLog "Common Bridge setup"
        
        'Build supporting collections and vectors...
        frmSplash.Message 15
        g.Functions.Load
        
        ' Filter out any functions with a required mod that the user does not have...
        FilterFunctions
        
        StartupLog "Functions loaded"
        
        Set g.Security = New cSecurity
        Set g.Coloring = New cColoring
        
        StartupLog "Collections loaded"
    
        'Init engine
        frmSplash.Message 20, "Loading functions ..."
        i = InitEngine(True, strTemp)
        StartupLog "Engine initialized: " & CStr(i) & " " & strTemp
        
        ' Auto import any quote board tab files as necessary...
        AutoImportQuoteBoardTabs
        
        ' Startup the TAS Authenticator now (if needs to be running and not already running)
        TASDllExists
    End If
                           
    'open data mgr database and load symbols
    frmSplash.Message 30, "Loading symbols ..."
    If Not bEmpty Then
        DM_Init True
        StartupLog "DataManager initialized"
        ' 9/20/2013: check if need to add Mutual Fund tables
        If Not FileExist(DataPath & "Mut_*.dbf") Then
            strFile = App.Path & "\NewMut.gzp"
            If FileExist(strFile) Then
                ZipExecute "U", strFile, DataPath
                If FileExist(DataPath & "Mut_*.dbf") Then
                    UpdateDBConfig
                    KillFile strFile
                End If
            End If
        
            ' and if upgrading and had been downloading stocks
            If 0 Then 'If FileExist("") Then
                s = "NOTE: for better performance comparison,| the historical stock prices are now being back-adjusted for dividends (along with splits).| If for some reason you desire to turn this off,| see the Misc. tab of the Program Settings."
                InfBox s, "i", , "Dividend Adjusting for Stocks"
            End If
        End If
    End If
    frmSplash.Message 40
    
    ' 12/13/2010 DAJ: Make sure that this object is intialized before the symbol pool because
    ' g.SymbolPool.Load.LoadSymbols.SetAccessFlags.SfeAllowedByBroker calls g.Broker.IsBrokerUser now...
    Set g.Broker = New cBrokerDispatch
    'Set g.Turnkey = New cTurnkey
    
    If g.Universe.OpenDb Then
        g.SymbolPool.Load True
        StartupLog "Symbols loaded"
    End If
    
    Set g.FractZen = New cFractZen
    
    If g.bImportStrategyBaskets = True Then
        mSysNav.ImportStrategyBaskets
        StartupLog "Strategy Baskets Imported..."
    End If
    
    FixSystemSecuritySymbolIDs
    StartupLog "System Security ID's Fixed..."
    CleanOutCustomIndexes
    StartupLog "Custom Indexes Cleaned out..."
    FixDuplicateBasketNames
    StartupLog "Fixed Duplicate Basket Names..."
    
    ' TLB 5/3/2015: just in case the auto-trading items get all screwed up ...
    s = App.Path & "\DeleteAutoTradeItems.ask"
    If FileExist(s) Then
        KillFile s
        If InfBox("Delete ALL the existing AutoTrade Items?", "?", "Delete|-No", "Delete AutoTrade Items") = "D" Then
            DeleteAutoTradeItems
        End If
    End If
    
#If 0 Then
    'make sure quote list exists (create default list if not)
    If Not FileExist(App.Path & "\Custom\QuoteList.GRP") Then
        Set SymbolGroup = New cSymbolGroup
        SymbolGroup.MakeSpecialType "QuoteList.GRP", "Quote List", eGROUP_QuoteList
        SymbolGroup.ToFile
        SymbolGroup.AddToPool True
        Set SymbolGroup = Nothing
    End If
#End If
    
    ' get setting for displaying in local time zone or not
    If GetRegistryValue(rkLocalMachine, strRegKey, "DisplayLocalTimeZone", 999) = 999 Then
        SetRegistryValue rkLocalMachine, strRegKey, "DisplayLocalTimeZone", vbChecked ' default
        ' if just upgraded, notify of the change
        If GetRegistryValue(rkLocalMachine, strRegKey, "chkSymbolExpire", 999) <> 999 Then
            strTemp = "By default, Trade Navigator now displays| all trade times in your local time zone |(e.g. quote board, minute bars, trade reports).||However, this is a setting which can be changed on the 'Misc' tab of the Program Settings."
            InfBox strTemp, "i", , "PLEASE NOTE ..."
        End If
    End If
    g.bShowInLocalTimeZone = GetRegistryValue(rkLocalMachine, strRegKey, "DisplayLocalTimeZone", vbChecked)
StartupLog "Display Time Zone loaded..."
    
    ' show main form (off-screen)
    frmSplash.Message 50, "Loading forms ..."
    If Not FormIsLoaded("frmMain") Then Load frmMain
    
    Set g.RealTime = New cRealTime
    g.nRecalcIndRT = GetIniFileProperty("RecalcIndRT", 0, "General", g.strIniFile)
    ' TLB 5/1/2014: since "Classic" is now mostly obsolete, the first step to phase it out is to
    ' always reset to "NextGen" at startup (while still leaving option to change it in Program Settings)
    g.RealTime.UseNextGen = True
    
    If IsAtLeastVista And FileExist(WinSysPath & "UxTheme.dll") Then
        If g.nColorTheme = kDarkThemeColor Then
            i = SetWindowTheme(frmMain.hWnd, "", 0)
            i = geHighContrastOn(frmMain.hWnd, kDarkThemeColor, vbWhite)   'hWnd not used, just passing it in case have use for it in future
            SendMessage frmMain.hWnd, WM_THEMECHANGED, 0, 0
        ElseIf g.nColorTheme = vbWhite Then
            i = SetWindowTheme(frmMain.hWnd, "", 0)
            i = geHighContrastOn(frmMain.hWnd, vbWhite, vbBlack)
            SendMessage frmMain.hWnd, WM_THEMECHANGED, 0, 0
        Else
            'i = SetWindowTheme(frmMain.hWnd, "", 0)
        End If
    End If
    
    ToolbarReset
StartupLog "Toolbar Reset..."
    
'    LockWindowUpdate frmMain.hWnd
    'nSaveMainTop = frmMain.Top
    'frmMain.Top = -1000 - frmMain.Height 'Screen.Height + 9000
    'StartupLog "frmMain loaded"
    
    ' Image Server
    If FileExist(App.Path & "\ImageServer.flg") Then
        frmMain.tbToolbar.Tools("ID_ImageServer").Visible = True
        Load frmImageServer
    Else
        frmMain.tbToolbar.Tools("ID_ImageServer").Visible = False
    End If
    
    ' rename old manual if no new manual yet
    strFile = App.Path & "\TradeNavigatorManual.pdf"
    strTemp = App.Path & "\ChartNavigatorManual.pdf"
    If Not FileExist(strFile) Then
        If FileExist(strTemp) Then
            Name strTemp As strFile
        End If
    End If
    
    ' Initialize the global profit object...
    Set g.Profit = New cProfit
    
    ' Initialize the Help Object...
    Set g.Help = New cHelp
    g.Help.Init frmMain.hWnd, AddSlash(App.Path) & "Help"
    
    Set g.ActivityLogs = New cActivityLogs
    Set g.OrderStrategies = New cOrderStrategies
    g.OrderStrategies.Load
StartupLog "OrderStrategies Loaded..."
    Set g.TradingItems = New cAutoTradeItems
    g.TradingItems.Load
StartupLog "TradingItems Loaded..."
    Set g.FlattenQueue = New cFlattenQueue
    Set g.CondOrders = New cConditionalOrders
    Set g.ExitAllOrders = New cExitAllOrders
    Set g.TsoGroups = New cActiveTsOrderGroups
StartupLog "Profit/Help/Broker Loaded..."
    
    ' MUST load these forms BEFORE adding to docked control
    ' (else may not show at startup)
    frmSplash.Message 55
    Load frmSymbolGrid
StartupLog "Symbol Grid Loaded..."
    frmSplash.Message 60
    Load frmQuotes
    frmSplash.Message 65
StartupLog "Quote Board Loaded..."
    
    Load frmMessage
    Load frmStatus
    Load frmOptionChain
    Load frmChartData
    Load frmPlanetData
    Load frmSnapshot
    Load frmChartOnOff
StartupLog "Other Forms Loaded..."
    Load frmTTSummary
    'Load frmChartCfg
    'Load frmEditAnnot
       
    'JM 12-21-2015: need to this AFTER the above forms have been loaded, but BEFORE filtering trade console grids
    '   seems to be the only way to get the docked trade console forms to do alternate grid row colors as expected
    SetTheAppBackColor
    g.JournalBridge.SetAppBackColor GetAppBackColor
    
    Set g.ConsoleForms = New cTradeConsoleForms
    g.ConsoleForms.FilterGrids
StartupLog "Console forms loaded"
    
    'frmTTSummary.Caption = "TRADE CONSOLE  (info for live accounts believed to be correct - contact your broker if questions)"
    frmTTSummary.Caption = "TRADE CONSOLE  (info for live accounts based on data transmitted from broker - contact your broker if questions)"
    
    With frmMain.DockPro
        'JM 12-01-2015 set OS Version so DockPro knows whether to disable windows theme
        .OsVersion = WindowsVersion()
        ' don't get previous docked settings if "reset" is passed
        ' on command line, or if build before upgrade was prior to 496
        ' (since a docked form was removed and seems to be causing problems)
        If bReset Or lPrevBuild < 496 Then
            .Persistent = False
        Else
            .Persistent = True
        End If
        
        ''.SysMenuCaption = "Allow Docking"
        .LeftEdgeWidth = 3660
        .BottomEdgeHeight = 2280
        .TopEdgeHeight = 1650
        .MDIMenuBarLetters = "FEVRCTPDH"
                
        ' add all dock-type forms with default settings
        If ExtremeCharts >= 1 Then
            .AddForm frmSymbolGrid, DPDocked, HAlignLeft, , , 0
        Else
            .AddForm frmSymbolGrid, -DPDocked, HAlignLeft, , , 0
        End If
        .AddForm frmQuotes, -DPDocked, HAlignBottom
        .AddForm frmSnapshot, -DPUndocked
        .AddForm frmOptionChain, -DPUndocked
        .AddForm frmMessage, -DPUndocked
        .AddForm frmChartData, -DPUndocked
        .AddForm frmOrderTracker, -DPUndocked
        .AddForm frmPlanetData, -DPUndocked
        .AddForm frmChartOnOff, -DPDocked, HAlignLeft
        .AddForm frmTTSummary, -DPDocked, HAlignTop
        '.AddForm frmTTSummary, -DPUndocked, HAlignTop
        '.AddForm frmChartCfg, -DPUndocked
        '.AddForm frmEditAnnot, -DPUndocked
        
        ' set defaults for these forms
        .Dockable("frmSnapshot") = False
        .Dockable("frmOptionChain") = False
        .Dockable("frmMessage") = False
        .Dockable("frmChartData") = False
        .Dockable("frmPlanetData") = False
        .Dockable("frmSnapshot") = False
       
        ' now reload settings from when used last
        If .Persistent Then .LoadPersistanceSettings
                
        ' DAJ 02/26/2010: With the Trade Console in the new mode with the toolbar on top, it no
        ' longer makes sense to allow it to be docked on either side (make sure to do this after
        ' persisting)...
        .CanDockRight("frmTTSummary") = False
        .CanDockLeft("frmTTSummary") = False
        
        ' make sure these forms start hidden
        ''.State("frmStatus") = DPHidden
        .State("frmOptionChain") = DPHidden
        '.State("frmChartCfg") = DPHidden
        '.State("frmEditAnnot") = DPHidden
        .State("frmMessage") = DPHidden
        .State("frmSnapshot") = DPHidden
        If g.SymbolPool.NumRecords = 0 Then
            .State("frmSymbolGrid") = DPHidden
        End If

        ' get list of undocked forms which will be shown later
        If .State("frmSymbolGrid") = DPUndocked Then aUndocked.Add "frmSymbolGrid"
        If .State("frmQuotes") = DPUndocked Then aUndocked.Add "frmQuotes"
        If .State("frmChartData") = DPUndocked Then aUndocked.Add "frmChartData"
        If .State("frmPlanetData") = DPUndocked Then aUndocked.Add "frmPlanetData"
        If .State("frmOrderTracker") = DPUndocked Then aUndocked.Add "frmOrderTracker"
        If .State("frmChartOnOff") = DPUndocked Then aUndocked.Add "frmChartOnOff"
        If .State("frmTTSummary") = DPUndocked Then aUndocked.Add "frmTTSummary"
        'If .State("frmSnapshot") = DPUndocked Then aUndocked.Add "frmSnapshot"
        
        If bHidden Then
            ' when starting minimized, hide undocked forms initially
            For i = 0 To aUndocked.Size - 1
                .State(aUndocked(i)) = DPHidden
            Next
            aUndocked.Size = 0
        End If
        
        ' set toolbar buttons
        frmMain.tbToolbar.Redraw = False
        If .State("frmSymbolGrid") <> DPHidden Then
            frmMain.tbToolbar.Tools("ID_SymbolGrid").State = ssChecked
        End If
        If .State("frmQuotes") <> DPHidden Then
            frmMain.tbToolbar.Tools("ID_Quote").State = ssChecked
        End If
        If .State("frmSnapshot") <> DPHidden Then
            frmMain.tbToolbar.Tools("ID_Snapshot").State = ssChecked
        End If
        If .State("frmChartData") <> DPHidden Then
            frmMain.tbToolbar.Tools("ID_ChartData").State = ssChecked
        End If
        If .State("frmPlanetData") <> DPHidden Then
            frmMain.tbToolbar.Tools("ID_PlanetData").State = ssChecked
        End If
        If .State("frmOrderTracker") <> DPHidden Then
            frmMain.tbToolbar.Tools("ID_Orders").State = ssChecked
        End If
        If .State("frmChartOnOff") <> DPHidden Then
            frmMain.tbToolbar.Tools("ID_ChartOnOff").State = ssChecked
        End If
        If .State("frmTTSummary") <> DPHidden Then
            frmMain.tbToolbar.Tools("ID_TradeTracker").State = ssChecked
        End If
        frmMain.tbToolbar.Redraw = True
        
        ' hide undocked forms for now, show later
        For i = 0 To aUndocked.Size - 1
            .State(aUndocked(i)) = DPHidden
        Next
        
        .Paint
        
        'now turn persistence on
        If Not .Persistent Then .Persistent = True
    End With
    StartupLog "Forms loaded"
    frmSplash.Message 70
    
    ' show off-screen first
    frmMain.Show
    DoEvents
    StartupLog "Main form shown"
            
    'show symbol grid
    'Load frmSymbolGrid
    If g.SymbolPool.NumRecords > 0 Then
        frmSymbolGrid.ShowSymbol ""
        frmSplash.Message 75, "Loading charts ..."
        strFile = App.Path & "\Charts\Loading.tmp"
        If FileExist(strFile) Then
            ' this file should only exist if the previous startup had an error
            ' while loading the charts, which means they're probably corrupt
            strTemp = "Restore the previously loaded charts? |(choose 'No' if there were errors the |last time the charts were being loaded)"
            If InfBox(strTemp, "?", "+Restore|-No", "Restore charts") = "N" Then
                KillFile App.Path & "\Charts\*.cht"
                KillFile App.Path & "\Charts\Charts.cfg"
            End If
        End If
        FileFromString strFile, "Loading charts"
        
        RestoreCharts True
        KillFile strFile
        StartupLog "Charts restored"
        
        'If bHidden Then
        '    SplashMessage "Creating chart ..."
        'ElseIf Not bEmpty And InStr(UCase(strCmd), "NOCHART") = 0 Then
            ' go to symbol
        '    SplashMessage "Loading chart ..."
        '    frmSymbolGrid.ShowSymbol ""
        '    Set frm = ActiveChart
        '    If Not frm Is Nothing Then frm.WindowState = 2
        '    StartupLog "Initial symbol selected and charted"
        'End If
    End If
    
    ' Rename any MRG files to SB files and add the Symbol ID to appropriate lines...
    FixStrategyBasketFiles
       
    ' now make main form visible on screen
    frmSplash.Message 100
    StatusMsg
    'StatusMsg "To see charting Hot-keys and Tips, hit 'Ctrl-H'"
    If Not FileExist(App.Path & "\ImageServer.flg") Then
        i = 0
        Do While frmSplash.Visible
            Select Case UCase(Left(Trim(frmSplash.Tag), 1))
            Case "N" ' if disclaimer cancelled, shut down
                g.bStarting = False
                Unload frmSplash
                Unload frmMain
                Exit Sub
            Case "W" ' wait for response
                If i = 0 Then
                    MoveFocus frmSplash.cmdOK
                    i = 1
                ElseIf i = 1 Then
                    i = 2
                    frmSplash.Message 0, "The 'ACCEPT' button must be clicked ..."
                Else
                    i = 1
                    frmSplash.Message 100, "The 'ACCEPT' button must be clicked ..."
                End If
                Sleep 0.5
            Case Else
                Exit Do
            End Select
        Loop
    End If
    Unload frmSplash
    DoEvents
    
    ' load toolbar stuff before frmMain gets shown (so charts will initially be the right size)
    g.bLoadingChartPage = True ' (set this temporarily just so GenerateChart will get bypassed)
    frmMain.pbNotUsed.Visible = False
    ToolbarInit2 frmMain, frmMain.TbButtonsArray(kTbGeneral)
    ToolbarResize2 frmMain, frmMain.pbTbBack, frmMain.imgTbBack, frmMain.TbButtonsArray(kTbGeneral), frmMain.ToolBarWrapGet(kTbGeneral)
    
    ToolbarInit2 frmMain, frmMain.TbButtonsArray(kTbDraw), , kTbDraw, , g.vbeTbAlignDraw
    ToolbarResize2 frmMain, frmMain.pbTbBackDraw, frmMain.imgTbBackDraw, frmMain.TbButtonsArray(kTbDraw), frmMain.ToolBarWrapGet(kTbDraw)
    
    g.bLoadingChartPage = False
    
    frmMain.InitialShow bHidden
    
    ' and show undocked forms
    For i = 0 To aUndocked.Size - 1
        frmMain.DockPro.ShowForm aUndocked(i), EAlignPrevious
    Next
    StartupLog "Forms shown on screen"
    
    ' set initial focus to appropriate window
    Set frm = ActiveChart
    If Not frm Is Nothing Then
        MoveFocus frm
        frmChartData.ShowData -1
        frmPlanetData.ShowData -1
        frmChartOnOff.ShowData
    ElseIf DockState(frmSymbolGrid) <> eHidden Then
        MoveFocus frmSymbolGrid
    ElseIf DockState(frmQuotes) <> eHidden Then
        MoveFocus frmQuotes
    Else
        MoveFocus frmMain
    End If
    
    SetRegistryValue rkLocalMachine, strRegKey, "BuildNumber", App.Revision
    
    g.RealTime.SalmonSetWindow
    
    SetMainCaption
    LockWindowUpdate 0
    frmMain.tmrMain.Enabled = True
    'Unload frmSplash
    
    g.bDirtyChartPage = FileExist(App.Path & "\Charts\Page.flg")
    g.bLoadPageOldMethod = FileExist(App.Path & "\LoadPageOld.flg")
    g.bLoadPageTime = FileExist(App.Path & "\LoadPageTime.flg")
    
    ' do initial symbol link for old AutoSync setting (which is now obsolete)
    strRegKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    If GetRegistryValue(rkLocalMachine, strRegKey, "AutoSync", False) Then
        SetRegistryValue rkLocalMachine, strRegKey, "AutoSync", False
        i = RGB(255, 0, 255)
        'frmSymbolGrid.WindowLink.SymbolColor = i
        'frmSnapshot.WindowLink.SymbolColor = i
        'If Not ActiveChart Is Nothing Then
        '    ActiveChart.WindowLink.SymbolColor = i
        'End If
        'g.bDirtyChartPage = True
    End If
    
    StartupLog "End of main", -1
    DebugLog "End of main"
      
ErrExit:
    Exit Sub
    
ErrSection:
    LockWindowUpdate 0
    RaiseError "mMain.Main", eGDRaiseError_Show
    If FormIsLoaded("frmMain") Then
        Unload frmMain
        DoEvents
    End If
    End
    Exit Sub
Resume 'RH
End Sub

Public Function LinkTableToDb(DB As Database, strLinkedTable As String, strLinkToDb As String) As Boolean
    
    ' if the current connection for the given table is different,
    ' relink to the new given connection (strLinkToDb)
    Dim tdTmp As TableDef
    Dim strConnect As String, strTmp As String
    Dim NeedToRelink As Boolean
    
    LinkTableToDb = False
    ' find the given table (try to relink if table not found)
    On Error Resume Next
    Set tdTmp = DB.TableDefs(strLinkedTable)
    On Error GoTo LinkTable_Error
    ' form the connect string
    strConnect = ";DATABASE=" & strLinkToDb
    If tdTmp Is Nothing Then
        NeedToRelink = True ' table not found, try relink anyway
    Else
        ' only need to relink if the connection is different
        If UCase(strConnect) <> UCase(tdTmp.Connect) Then
            NeedToRelink = True
            ' must delete current tabledef from collection
            DB.TableDefs.Delete strLinkedTable
        End If
    End If
    If NeedToRelink Then
        ' create new tabledef and setup new link
        Set tdTmp = DB.CreateTableDef(strLinkedTable)
        tdTmp.SourceTableName = strLinkedTable
        tdTmp.Connect = strConnect
        ' append new tabledef to collection
        DB.TableDefs.Append tdTmp
        'tdTmp.RefreshLink
    End If
    If UCase(strConnect) = UCase(tdTmp.Connect) Then
        LinkTableToDb = True
    End If

LinkTable_Exit:
    If Not (tdTmp Is Nothing) Then Set tdTmp = Nothing
    Exit Function

LinkTable_Error:
    strTmp = "Error linking " & strLinkedTable & " to " & strLinkToDb
    InfBox strTmp, "[]", , "LinkTableToDb"
    Resume LinkTable_Exit

End Function

Public Function DbPassword() As String
On Error GoTo ErrSection:

    Dim strKey$, i&
    Dim mb As cMemBuffer, mbKey As cMemBuffer
    Static strPW$
    
    'if already gotten, just return it
    If Len(strPW) > 0 Then
        DbPassword = strPW
        Exit Function
    End If
    
    Set mb = New cMemBuffer
    Set mbKey = New cMemBuffer
    
    ' key used to get the password
    mbKey.Buffer = "ToEncryptThePassword"
    
    ' include the following ONLY when first getting the encrypted password
    ' (code inside the "#If 0 Then" is NOT compiled into the EXE)
    #If 0 Then
        ' here's the real password
        mb.PutStr "v10sysnav"
        ' encrypt it, then show the ascii #'s of the encrypted password
        gdEncrypt True, mb, mbKey
        strPW = ""
        For i = 0 To mb.Length - 1
            strPW = strPW & Str(mb.GetByte(i)) & " "
        Next
        MsgBox strPW
        Exit Function
    #End If
            
    ' decrypt the password
    mb.PutByte 247
    mb.PutByte 80
    mb.PutByte 137
    mb.PutByte 158
    mb.PutByte 134
    mb.PutByte 49
    mb.PutByte 98
    mb.PutByte 48
    mb.PutByte 129
    gdEncrypt False, mb, mbKey
    strPW = mb.Buffer
    
    DbPassword = strPW

ErrExit:
    Set mb = Nothing
    Exit Function
    
ErrSection:
    RaiseError "mMain.dbPassword", eGDRaiseError_Raise
    
End Function

' called when frmMain unloads
Public Sub CleanupWhenExit()

    Dim rc&, strTemp$, strDB$, strBU1$, strBU2$, i%, strKey$, dDeleteZip#

    StartupLog "------ Cleanup -------"
        
'    g.RealTime.SalmonStop
        
    ChartTimers = False
    frmTTSummary.DisableTimers

    Set g.RptBridge = Nothing
    Set g.Help = Nothing
    Set MsgForm = Nothing
    
    StartupLog "Cleanup: disconnections"
    
    ' set flag file for dirty chart page
    strTemp = g.ChartGlobals.strCPCRoot & "\Charts\Page.flg"
    If Not g.bDirtyChartPage Then
        KillFile strTemp
    ElseIf Not FileExist(strTemp) Then
        FileFromString strTemp, "Dirty page"
    End If
    
    ' Close the trade console forms...
    Set g.ConsoleForms = Nothing
    
    ' first close all forms except main
    ' (may be stray non-child non-modal forms open)
    UnloadEditors
    InfBox ""
    DoEvents
    On Error Resume Next
    SetPrevActiveForm Nothing
    If FormIsLoaded("frmOnlineBroker") Then
        StartupLog "Closing: frmOnlineBroker"
        Unload frmOnlineBroker
    End If
    If FileDate(App.Path & "\SalmonClient.DLL") >= 40702 Then ' if >= 6/8/2011
        StartupLog "Closing: Shutdown_SalmonDLL"
        Shutdown_SalmonDLL
    End If
    StartupLog "Closing: other forms"
    Do
        rc = Forms.Count
        For i = Forms.Count - 1 To 0 Step -1
            If i < Forms.Count Then ' need this check, just in case 1 form closing also caused another to close
                If Not TypeOf Forms(i) Is frmMain Then
                    Unload Forms(i)
                End If
            End If
        Next
        DoEvents
        ' keep looping until an iteration where no more got closed
    Loop While rc <> Forms.Count
    
    LockWindowUpdate 0
      
    StartupLog "Cleanup: forms closed"
    
    SaveChartGlobals
    Set g.ChartGlobals.frmLastChartMouseMove = Nothing
    Set g.Alerts = Nothing
   
    Set g.RealTime = Nothing
    
    Set g.Profit = Nothing
    
    Set g.FractZen = Nothing
        
    ' clean up old files
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    dDeleteZip = Int(GetRegistryValue(rkLocalMachine, strKey, "DeleteZipFiles", dDeleteZip))
    If dDeleteZip <= 0 Then dDeleteZip = 31
    KillFile App.Path & "\ftp\Backup\* /i=zip,gzp /o=-" & Str(dDeleteZip - 1)
    ' Let the TransAct object take care of this (DAJ 10/26/2007)...
    'KillFile App.Path & "\Transact\*.log /o=-14"
    KillFile App.Path & "\Stream\*.log /o=-31"
    KillFile App.Path & "\Trades\*.txt"
   
    StartupLog "Cleanup: delete old files"
    
    ' Let engine unload
    rc = InitEngine(False, strTemp)
    
    StartupLog "Cleanup: engine unloaded"
    
    ' Close database
    DoEvents
    If Not g.dbNav Is Nothing Then
        If (CheckLibIndexes("before compact") = False) Then g.bSkipMdbCompact = True
        g.dbNav.Close
    End If
    If Not g.dbPaper Is Nothing Then
        If (CheckTtIndexes("before compact") = False) Then g.bSkipMdbCompact = True
        g.dbPaper.Close
    End If
    DoEvents
    
    StartupLog "Cleanup: db's closed"
    
    'rc = gAllocChk
    g.SymbolPool.MakeDebugFile
    g.SymbolPool.SaveFlagGroup
    Set g.SymbolPool = Nothing
    'rc = gAllocChk
    Set g.Universe = Nothing
    'rc = gAllocChk
    
    StartupLog "Cleanup: symbol pool unloaded"
    
    ' close codebase stuff
    DM_Init False
    Cb4Finish
    
    StartupLog "Cleanup: codebase closed"
    
    ' Cleanup global objects ...
    Set g.Security = Nothing
    Set g.Functions = Nothing
    Set g.CommonBridge = Nothing
    Set g.ActiveEditor = Nothing
    Set g.RptBridge = Nothing
    Set g.Coloring = Nothing
    Set g.CurrentSystem = Nothing
    Set g.ActivityLogs = Nothing
    Set g.OrderStrategies = Nothing
    Set g.TradingItems = Nothing
    Set g.FlattenQueue = Nothing
    Set g.CondOrders = Nothing
    Set g.ExitAllOrders = Nothing
    Set g.TsoGroups = Nothing
    Set g.Broker = Nothing
    Set g.tblLibrary = Nothing
    Set g.tblFunction = Nothing
    Set g.tblFunctionParm = Nothing
    Set g.tblRule = Nothing
    Set g.astrFunctionCategory = Nothing
    
    CopyRecalcLog
    
    InfBox ""
    DoEvents ' give chance for all of above to completely terminate
    
    StartupLog "Cleanup: objects destructed"
    
    ' Try to compact databases
    If Not g.bSkipMdbCompact Then
        strDB = App.Path + "\Libraries.MDB"
        If FileExist(strDB) And g.bDirtyLibrariesMDB Then
            strBU1 = App.Path + "\LibBak1.mdb"
            strBU2 = App.Path + "\LibBak2.mdb"
            strTemp = App.Path + "\_Temp.mdb"
            KillFile strTemp
            On Error GoTo CompactError
            DBEngine.CompactDatabase strDB, strTemp, , dbDecrypt, ";pwd=" & DbPassword
            StartupLog "Libraries.MDB Compacted..."
            If FileExist(strTemp) Then
                Set g.dbNav = OpenDatabase(strTemp, DbPassword)
                StartupLog "Libraries.MDB reopened..."
                If Not g.dbNav Is Nothing Then
                    If CheckLibIndexes("after compact") = True Then
                        g.dbNav.Close
                        StartupLog "Libraries.MDB checked for indexes..."
                        DoEvents
                        ' save backups
                        KillFile strBU2
                        If FileExist(strBU1) Then Name strBU1 As strBU2
                        KillFile strBU1
                        Name strDB As strBU1
                        Name strTemp As strDB
                    Else
                        g.dbNav.Close
                    End If
                End If
            End If
        End If
        
        strDB = App.Path + "\TradeTracker.MDB"
        If FileExist(strDB) Then
            strBU1 = App.Path + "\TTBak1.mdb"
            strBU2 = App.Path + "\TTBak2.mdb"
            strTemp = App.Path + "\_Temp.mdb"
            KillFile strTemp
            On Error GoTo CompactError
            DBEngine.CompactDatabase strDB, strTemp, , dbDecrypt
            StartupLog "TradeTracker.MDB Compacted..."
            If FileExist(strTemp) Then
                Set g.dbPaper = OpenDatabase(strTemp, "")
                StartupLog "TradeTracker.MDB reopened..."
                If Not g.dbPaper Is Nothing Then
                    If CheckTtIndexes("after compact") = True Then
                        g.dbPaper.Close
                        StartupLog "TradeTracker.MDB checked for indexes..."
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
        End If
        
        StartupLog "Cleanup: databases compacted"
    End If
    
    ' store the Libraries file info
    strDB = App.Path + "\Libraries.MDB"
    SetIniFileProperty "LastLib", Str(FileDate(strDB)) & ";" & Str(FileLength(strDB)), "General", g.strIniFile
    
CompactErrExit:
    On Error Resume Next
    Set g.dbNav = Nothing
    Set g.dbPaper = Nothing
    Set g.WrkJet = Nothing
    
    If Len(g.strRunWhenExit) > 0 Then
        ' TLB: skip this check since could be an Upgrade, Beta, Alpha, etc.
        'If InStr(UCase(g.strRunWhenExit), "UPGRD") > 0 Then
            If FileExist(App.Path & "\ftp\NavSuite.exe") Then
                KillFile App.Path & "\SymPool.mem"
            End If
        'End If
        If FileDate(App.Path & "\ftp\TradeNavStartup.exe") > FileDate(App.Path & "\TradeNavStartup.exe") Then
            FileFromString App.Path & "\TradeNavStartup.Kil", "kill now"
        End If
        ' if a newer TAS Authenticator is in the upgrade, then kill the currently running process
        If FileDate(App.Path & "\ftp\TASAuthServer.exe") > FileDate(App.Path & "\TASAuthServer.exe") Then
            KillProcess "TAS Authenticator"
        End If
        If FileDate(App.Path & "\ftp\TASLaunchPad.exe") > FileDate(App.Path & "\TASLaunchPad.exe") Then
            KillProcess "TAS Launch Pad"
        End If
        KillOtherPrograms
        Shell g.strRunWhenExit, vbNormalFocus
    Else
        ' if an NVS machine needs to be rebooted, initiate that now
        RebootNVSIfRequired
    End If
    
    StartupLog "Cleanup: finished"
    Exit Sub

CompactError:
    Resume CompactErrExit:
    
End Sub

Public Function ConvertDate(dNewDate As Date)
On Error GoTo ErrSection:

    ConvertDate = Right(Str(JulToLong(CLng(dNewDate), True)), 6)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ConvertDate", eGDRaiseError_Raise
    
End Function

Public Sub AutoSizeChart()
On Error GoTo ErrSection:

    Exit Sub ' obsolete: now handled in timer of frmMain

#If 0 Then
    With frmChart
        If .bSingleChart Then
            If .Visible And .AutoSize <> 0 And WindowStateX(frmChart) = wsNormal Then
                .AutoSize = -2
            End If
        End If
    End With
#End If
    
    If Not g.bStarting And Not g.bUnloading Then
        If Not ActiveChart Is Nothing Then
            If ActiveChart.DetachStatus = eNotDetached Then
                If ActiveChart.WindowState <> vbMaximized Then
                    frmArrange.ArrangeCharts True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.AutoSizeChart", eGDRaiseError_Raise
    
End Sub

Public Sub StartupLog(ByVal strMsg$, Optional ByVal iStartStop% = 0)
On Error GoTo ErrSection:

    Dim dCurTime#, strFile$
    Static dPrevTime#, dStartTime#

    strFile = App.Path & "\Startup.Log"

    'see if starting new log
    If iStartStop > 0 Then
        KillFile strFile
        dStartTime = gdTickCount
        dPrevTime = dStartTime
    End If
            
    'write message to startup log
    dCurTime = gdTickCount
    strMsg = Format((dCurTime - dPrevTime) / 1000#, "#0.00  ") & strMsg
    FileFromString strFile, strMsg, True, True
    dPrevTime = dCurTime
    
    'see if ending log
    If iStartStop < 0 Then
        strMsg = Format((dCurTime - dStartTime) / 1000#, "#0.00  ") & "TOTAL STARTUP TIME"
        FileFromString strFile, strMsg, True, True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.StartupLog", eGDRaiseError_Raise
    
End Sub

Public Sub CopyToClipboard(ctl As Control)
On Error GoTo ErrSection:

    With Clipboard
        .Clear
        If TypeOf ctl Is ctlUniTextBoxXP Or TypeOf ctl Is TextBox Then 'RH was TextBox
            .SetText ctl.SelText
        ElseIf TypeOf ctl Is ctlUniComboImageXP Then 'was ComboBox
            .SetText ctl.Text
        ElseIf TypeOf ctl Is PictureBox Then
            .SetData ctl.Picture
        ElseIf TypeOf ctl Is ListBox Then
            .SetText ctl.Text
        ElseIf TypeOf ctl Is RichTextBox Then
            .SetText ctl.TextRTF, vbCFRTF
        Else
            'Try default property
            .SetText ctl
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    Beep

End Sub

'Called from Toolbar events of frmMain and frmChart2
Public Sub ToolBarClick(ByVal Tool As ActiveToolBars.SSTool, frmSource As Form, _
        Optional ByVal bFromDropDown As Boolean = False, _
        Optional ByVal eExtraInfo As eToolbarExtraInfo = eTbExtraInfo_None)
On Error GoTo ErrSection:

    Dim i&, menuItemCount&, d#
    Dim strFile$, strPath$, strTemp$, strFormType$
    Dim aStrings As New cGdArray
    Dim frmActive As Form, frm As Form
    Dim eFormType As eFormTypes
    Dim tbToolbar As SSActiveToolBars
    Dim strKey$, strSymbol$, nSymbolID&
    Dim Pane As cPane
    Dim PtOrder As cPtOrder
    
    Dim bChart As Boolean
    Dim bKeepChecking As Boolean
    Dim bClearFocus As Boolean
    
    Set tbToolbar = frmSource.tbToolbar
    
    If Not ActiveChart Is Nothing Then
        If ActiveChart.SkipFocusFix Then
            'user clicked on the menu when a detached chart was active      -4883
            If Tool.ID = "ID_Tile" Then
                If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                    SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_NCACTIVATE, 1, 0
                    SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_MOUSEACTIVATE, 1, 0
                End If
            End If
        End If
    End If
    
    If Not g.bStarting And Not g.bUnloading And Not g.bLoadingChartPage And tbToolbar.Redraw <> False Then
        FixFocusChart frmSource, Tool
        If Tool.ID <> "ID_Templates" Then
            If tbToolbar.Tools("ID_Templates").State = ssChecked Then
                tbToolbar.Tools("ID_Templates").State = ssUnchecked
            End If
        End If
        If Tool.ID <> "ID_Pages" Then
            If tbToolbar.Tools("ID_Pages").State = ssChecked Then
                tbToolbar.Tools("ID_Pages").State = ssUnchecked
            End If
        End If
        If TypeOf frmSource Is frmMain Then
            frmMain.LastClickedToolID = Tool.ID
            If Not ActiveChart Is Nothing Then
                'user clicked a tool button on main app toolbar when detached chart has focus
'06-30-2009: commented out to implement no toolbar on detached chart
'                If ActiveChart.DetachStatus = eDetached Then
'                    If Tool.Category = "Charting" Then
'                        If Tool.ID <> "ID_Tile" Then
'                            If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
'                                ActiveChart.SkipFocusFix = True
'                                SendMessage ActiveChart.hWnd, WM_NCACTIVATE, 0, 0
'                                SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_NCACTIVATE, 1, 0
'                                SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_MOUSEACTIVATE, 1, 0
'                            End If
'                        End If
'                    End If
'                End If
            End If
        Else
            frmMain.LastClickedToolID = ""
        End If
    End If
    
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    
    ' see if just clearing status msg
    If Tool.ID = "ID_Status" Then
        StatusMsg
        Exit Sub
    End If
        
    'Check for a button dropdown first (since this is kind of a modal
    ' thing and should be done before the .Redraw is turned off)
    If bFromDropDown Then
        ' toolbar button dropdown
        If tbToolbar.Redraw <> False Then
            On Error Resume Next
            strTemp = Tool.ID & "Popup"
            tbToolbar.PopupMenu tbToolbar.Tools(strTemp)
        End If
        ' for drop-downs, trigger this so the popup will
        ' display smoother and faster even during other processing
        frmMain.tmrMain.Tag = "SLEEP 0.5"
        Exit Sub
    End If
    
    'If Redraw is False, then this routine is being entered
    ' due to some code just updating some item states in which
    ' case we do not need to process anything -- the Redraw
    ' would be true if the user was actually clicking on something.
    If tbToolbar.Redraw = False Then Exit Sub
    '(and turn it off so won't take a chance on any recursion,
    ' e.g. when loading menu items or filling combo boxes)
    'Set tbToolbar = frmSource.tbToolbar
    If tbToolbar.Tag = "IGNORE" Then Exit Sub
    tbToolbar.Redraw = False
    
    '1)----------------------------------------------
    
    'Check if a menu dropdown
    If Tool.Type = ssTypeMenu Then
        ' menu bar dropdowns
        Select Case Tool.ID
        Case "ID_File"
            tbToolbar.Tools("ID_RecalcCriteria").Enabled = ScansEnabled
        
        Case "ID_Window"
            'Chart menu: build current list of chart windows
            ToolbarWindowList
        
        Case "ID_Help"
            'Help menu: get list of custom help items
            ToolbarHelpList
        
        Case "ID_TemplatesList"
            'fill menu with templates
            With tbToolbar.Tools("ID_TemplatesList").Menu
                'get list of templates
                Set aStrings = GetAllowedList("T")
                aStrings.Add " < &Manage chart templates >", 0
                aStrings.Add " < Copy settings to other charts >", 1

                'add menu item for each template
                For i = 1 To aStrings.Size
                    If i > .Tools.Count Then
                        .Tools.Add "Template #" & CStr(i)
                    End If
                    With .Tools(i)
                        .Name = Parse(aStrings(i - 1), vbTab, 1)
                    End With
                Next
                'remove extras
                For i = .Tools.Count To aStrings.Size + 1 Step -1
                    .Tools.Remove i
                Next
            End With
        
        Case "ID_PagesList"
            'fill menu with pages
            With tbToolbar.Tools("ID_PagesList").Menu
                If g.ChartGlobals.bMyPageFeature Then
                    ' first time: add static menu items
                    If .Tools.Count = 0 Then
                        .Tools.Add "Page #M", ssTypeButton
                        .Tools(1).Name = " < &My Page >"
                    ElseIf .Tools.Count >= 2 Then
                        ' temporarily remove the separator (in case no pages exist)
                        .Tools.Remove 2
                    End If
                    menuItemCount = 2
                Else
                    ' first time: add static menu items
                    If .Tools.Count = 0 Then
                        .Tools.Add "Page #M", ssTypeButton
                        .Tools(1).Name = " < &Manage chart pages >"
                        .Tools.Add "Page #S", ssTypeButton
                        .Tools(2).Name = " < &Save chart page >"
                        .Tools.Add "Page #N", ssTypeButton
                        .Tools(3).Name = " < &Create new chart page >"
                    ElseIf .Tools.Count >= 4 Then
                        ' temporarily remove the separator (in case no pages exist)
                        .Tools.Remove 4
                    End If
                    menuItemCount = 4
                End If
                
                'get list of pages
                Set aStrings = GetAllowedList("P")
                If Len(g.strChartPage) = 0 And Not g.ChartGlobals.bMyPageFeature Then
                    aStrings.Add "(unnamed)", 0
                End If
                'add menu item for each page
                For i = 0 To aStrings.Size - 1
                    If i + menuItemCount > .Tools.Count Then
                        strTemp = "Page #" & CStr(i)
                        .Tools.Add strTemp, ssTypeStateButton
                        tbToolbar.Tools(strTemp).PictureDown = Picture16(ToolbarIcon("kChecked"))
                    End If
                    With .Tools(i + menuItemCount)
                        strTemp = Parse(aStrings(i), vbTab, 1)
                        .Name = Replace(strTemp, "&", "&&")
                        If i = 0 And Len(g.strChartPage) = 0 Then
                            If Not g.ChartGlobals.bMyPageFeature Then .State = ssChecked
                        ElseIf UCase(strTemp) = UCase(g.strChartPage) Then
                            .State = ssChecked
                        Else
                            .State = ssUnchecked
                        End If
                    End With
                Next
                'remove extras
                For i = .Tools.Count To aStrings.Size + menuItemCount Step -1
                    .Tools.Remove i
                Next
                'insert the separator
                If .Tools.Count > 3 Then
                    .Tools.Add "separator", ssTypeSeparator, menuItemCount
                End If
            End With
        
        Case "ID_Sectors", "ID_Subsectors", "ID_Components"
            ToolbarSectorMenu tbToolbar, Tool.ID
        
        Case "ID_PrintMenu"
            'Print sub-menu: enable only visible items for printing
            With frmMain.tbToolbar
                
                .Tools("ID_PrintTradeConsole").Enabled = (DockState(frmTTSummary) <> eHidden)
                .Tools("ID_PrintQuoteBoard").Enabled = (DockState(frmQuotes) <> eHidden)
                .Tools("ID_PrintSymbolGrid").Enabled = (DockState(frmSymbolGrid) <> eHidden)
                .Tools("ID_PrintSnapshot").Enabled = (DockState(frmSnapshot) <> eHidden)
                .Tools("ID_PrintOptionsChain").Enabled = (DockState(frmOptionChain) <> eHidden)
                .Tools("ID_PrintChartData").Enabled = (DockState(frmChartData) <> eHidden)
                
                If DockState(frmMessage) = eHidden Then
                    .Tools("ID_PrintNews").Enabled = False
                    .Tools("ID_PrintNews").ChangeAll ssChangeAllName, "News"
                Else
                    .Tools("ID_PrintNews").Enabled = True
                    .Tools("ID_PrintNews").ChangeAll ssChangeAllName, frmMessage.Caption
                End If
                
                If ActiveChart Is Nothing Then
                    .Tools("ID_PrintChart").Enabled = False
                    .Tools("ID_PrintChart").ChangeAll ssChangeAllName, "Chart"
                Else
                    .Tools("ID_PrintChart").Enabled = True
                    .Tools("ID_PrintChart").ChangeAll ssChangeAllName, ActiveChart.Chart.Symbol
                End If
            End With
        
        Case "ID_Trading"
            With frmMain.tbToolbar
                .Tools("ID_Cattle").Visible = g.CattleBridge.IsCattleUser
                .Tools("ID_Turnkey").Visible = g.CattleBridge.IsTurnkeyUser
                .Tools("ID_TurnkeyAdministration").Visible = g.CattleBridge.IsTurnkeyAdminUser
            End With
            g.Broker.LoadTradingMenu
        
        Case "ID_DrawingTools"
            If g.nColorTheme = kDarkThemeColor Then
                If Not tbToolbar.Tools("ID_ChartMove") Is Nothing Then
                    strTemp = "kChartMoveHz"
                    If Not ActiveChart Is Nothing Then
                        If Not ActiveChart.Chart Is Nothing Then
                            If Not ActiveChart.Chart.AutoScale Then strTemp = "kChartMove"
                        End If
                    End If
                    If g.ChartGlobals.eChartMode <> eMode_Move Then
                        tbToolbar.Tools("ID_ChartMove").Picture = g.CoreBridge.ImgListToolbarExt("Light", strTemp, "", 16).ExtractIcon
                    End If
                End If
            End If

        Case Else
            If UCase(Left(Tool.ID, 11)) = "ID_TRADING_" Then
                g.Broker.HandleTradingMenu Tool.ID
            End If
        
        End Select
        
        ' for menu drop-downs, trigger this so the popup will
        ' display smoother and faster even during other processing
        frmMain.tmrMain.Tag = "SLEEP 0.5"
        GoTo Cleanup '(nothing more to look for)
    End If
    
    
    'Check if the "Window list" or "Help list"
    Select Case Tool.Group
    Case "WindowList"
        ' Tool.ID has hWnd of child window to activate,
        ' but activate it in the timer of frmMain to
        ' avoid toolbar crashing under some circumstances
        ' (don't know why!)
        frmMain.tmrMain.Tag = "ACTIVATE " & CStr(Val(Tool.ID))
        GoTo Cleanup '(nothing more to look for)
    
    Case "HelpList"
        strTemp = Tool.TagVariant
        d = Val(Parse(strTemp, vbTab, 1))
        strTemp = FixURL(Parse(strTemp, vbTab, 2))
        Select Case Int(d)
        Case 4 '.HLP
            RunProcess Parse(strTemp, "|", 1), Parse(strTemp, "|", 2)
        Case 1, 2 'HTML
            RunProcess InternetBrowser, Chr(34) & strTemp & Chr(34)
        Case Else 'Message window
            frmMessage.ShowMe Tool.Name, "@" & strTemp
        End Select
        GoTo Cleanup '(nothing more to look for)
    End Select
    
    'See if this tool requires a specific type of form
    Select Case UCase(Tool.Category)
        Case "CHARTING"
            strFormType = "FRMCHART"
            strTemp = "Chart"
        'Case "SYMBOLGRID"
        '    strFormType = "FRMSYMBOLGRID"
        '    strTemp = "SymbolGrid"
        Case Else
            strFormType = ""
    End Select
    
    'Get active form (active when menu was accessed)
    If frmSource Is frmMain Then
        If UCase(Tool.Category) = "CHARTING" Then
            Set frmActive = ActiveChart
            If frmActive Is Nothing Then
                ' no charts!
                Beep
                GoTo Cleanup
            End If
            ' set focus to form to work with (some functionality
            ' doesn't work right if form does not have focus)
            ' This was causing problems, but we don't think that we
            ' need it anymore anyway (DAJ 10/15/2002)
            ''MoveFocus frmActive
        Else
            Set frmActive = ActiveForm
            If frmActive Is Nothing Or frmActive Is frmMain Then
                ' if no charts, pretend a docked form is "active" (has focus)
                If DockState(frmSymbolGrid) <> eHidden Then
                    Set frmActive = frmSymbolGrid
                ElseIf DockState(frmQuotes) <> eHidden Then
                    Set frmActive = frmQuotes
                End If
            End If
        End If
    Else
        'active form has the menu on it (e.g. frmChart2)
        Set frmActive = frmSource
    End If
    'See if active form is a chart
    bChart = False
    If Not frmActive Is Nothing Then
        If IsFrmChart(frmActive) Then
            If Not frmActive Is ActiveChart Then Set frmActive = ActiveChart
            bChart = True
        End If
    End If
    
    '2)----------------------------------------------
    'Check ONLY for things that DO NOT require an active form
    bKeepChecking = False
    Select Case Tool.ID
        Case "ID_Templates"
            If Tool.State <> ssUnchecked Then
                frmTemplatePage.ShowMe frmSource, eMode_Templates, Tool.Left, Tool.Top
            ElseIf FormIsLoaded("frmTemplatePage") Then
                Unload frmTemplatePage
            End If
                
        Case "ID_Pages"
            If Tool.State <> ssUnchecked Then
                frmTemplatePage.ShowMe frmSource, eMode_Pages, Tool.Left, Tool.Top
            ElseIf FormIsLoaded("frmTemplatePage") Then
                Unload frmTemplatePage
            End If
                        
        Case "ID_Publish"
            PublishSharedChartPage

        Case "ID_SharedPage"
            DisplaySharedChartPages
                        
        Case "ID_HelpTopics"
            g.Help.ShowHelpDefault
        
        Case "ID_RealTime"
            If Not g.RealTime.Active And Tool.State = ssChecked Then
                ' if turning on, must wait until process is not busy
                If Not ProcessIsBusy Then
                    ' start realtime
                    If HasModule("RTG") Or HasModule("RTE") Then
                        If g.FtpDownloader.DownloaderIsRunning Then
                            strTemp = "We recommend you PAUSE the Historical Data Downloader while data streaming is turned on."
                            InfBox strTemp, "i", , "Realtime Streaming"
                        End If
                        g.nReplaySession = 0
                        g.RealTime.Init True
                    Else
                        ' not enabled
                        strTemp = "Your data service is not currently enabled for Realtime Streaming.  Please call Genesis at " _
                            & GetIniFileProperty("SalesContact", "800-808-3282", "GenesisFT", App.Path & "\Provided\Provided.INI") & " for more info."
                        InfBox strTemp, "i", , "Realtime Streaming"
                    End If
                End If
            ElseIf g.RealTime.Active And Tool.State = ssUnchecked Then
                ' stop realtime
                If g.Broker.CanStopStreaming("User stopping streaming", False) Then
                    g.RealTime.Init False, "User turned off"
                End If
            End If
            tbToolbar.Tag = "IGNORE"
            If g.RealTime.Active Then
                Tool.State = ssChecked
            Else
                Tool.State = ssUnchecked
            End If
            tbToolbar.Tag = ""
                    
        Case "ID_ConditionBuilder"
            If ActiveChart Is Nothing Then
                Beep
            Else
                frmConditionBuilder.ShowMe ActiveChart.Chart
            End If
        
        Case "ID_ProcessingStatus"
            frmStatus.ShowMe True
        
        Case "ID_StatusLabel"
            SetMainCaption
            StatusMsg
            g.Security.ClearGoodPasswords
            If KeyIsPressed(VK_CONTROL) Then
                ''EnableAero 2 ' to toggle it from previous state
            End If
            If Not IsAtLeastXP Then
                i = GetFreeSystemResources
                If i >= 0 And i < 90 Then
                    StatusMsg CStr(GetFreeSystemResources) & "% system resources free"
                Else
                    StatusMsg
                End If
            End If
            
        Case "ID_WhatsNew"
            frmWhatsNew.ShowMe
            
        Case "ID_Symbol"
            On Error Resume Next
            MoveFocus frmActive
            SendKeys "S"
    
        Case "ID_Copy"
            If bChart Then
                frmActive.Chart.PrintChart 1, False
            Else
                CopyToClipboard Screen.ActiveControl
            End If
    
        Case "ID_Print"
            If bChart Then
                frmActive.TopMost = False
                frmActive.Chart.PrintChart 2, True
            Else
                On Error Resume Next
                frmActive.PrintMe
            End If
            
        Case "ID_PrintChart"
            Set frmActive = ActiveChart
            If Not frmActive Is Nothing Then
                frmActive.TopMost = False
                frmActive.Chart.PrintChart 2, True
            End If
        
        Case "ID_PrintChartData"
            frmChartData.PrintMe
            
        Case "ID_PrintSymbolGrid"
            frmSymbolGrid.PrintMe
        
        Case "ID_PrintQuoteBoard"
            frmQuotes.PrintMe
            
        'Case "ID_PrintTradeTracker"
        '    frmTTPositions.PrintMe
        
        Case "ID_PrintTradeConsole"
            frmTTSummary.PrintMe
            
        Case "ID_PrintSnapshot"
            frmSnapshot.PrintMe
            
        Case "ID_PrintNews"
            frmMessage.PrintMe
            
        Case "ID_PrintOptionsChain"
            frmOptionChain.PrintMe
    
        Case "ID_Manual"
            strTemp = "TradeNavigatorManual.pdf"
            If FileExist(App.Path & "\" & strTemp) Then
                RunProcess App.Path & "\" & strTemp
            Else
                InfBox "Manual not found:|" & strTemp, "e", , "View Manual"
            End If
        
        Case "ID_ImageServer"
            If Tool.Visible Then
                frmImageServer.ShowMe
            End If
            
        Case "ID_CustomizeToolbar"
            frmToolbar.ShowMe
               
        'Case "ID_HumeTools"
            'frmStatus.ShowMe
            'ShowForm frmHumeMain, True
        
        Case "ID_Exit"
            If Not ProcessIsBusy Then
                Unload frmSource
                Set frmSource = Nothing '(so won't turn redraw back on)
            End If
    
        Case "ID_News"
            strFile = App.Path & "\Info\News"
            If FileExist(strFile & ".htm") Then
                RunProcess InternetBrowser, Chr(34) & strFile & ".htm" & Chr(34)
            Else
                frmMessage.ShowMe "NEWS from Genesis", "@" & strFile, , True
            End If
    
        Case "ID_HotKeys"
            frmMessage.ShowMe "Charting Hot Keys and Tips", "@" & App.Path & "\Info\HotKeys"
    
        Case "ID_ProVersionInfo"
            frmMessage.ShowMe "Upgrading to 'GOLD' or 'PLATINUM'", "@" & App.Path & "\Info\ProVersion"
    
        Case "ID_RecalcCriteria"
            If Not ProcessIsBusy Then
                If Not ScansEnabled Then
                    InfBox "Filters are not currently enabled.", "i", , "Recalculate Criteria"
                    frmToolbox.ShowMe eTab_Filters
                    'frmConfig.ShowMe eMiscTab
                Else
                    frmStatus.tmrRecalc.Tag = ""
                    If GetRegistryValue(rkLocalMachine, strKey, "SessionUpdate", False) = True Then
                        If InfBox("Do you want to recalculate criteria based on Current Session Update or last End-of-Day Update?", _
                            "?", "+Session|-EndOfDay", "Criteria") = "S" Then
                                frmStatus.tmrRecalc.Tag = "-1"
                        End If
                    End If
                    strTemp = "Recalculate just the modified criteria,| or recalculate all criteria?"
                    strTemp = InfBox(strTemp, "?", "+Modified|All|-Cancel", "Recalculate Criteria")
                    If strTemp <> "C" Then
                        If strTemp = "A" Then 'ALL
                            g.SymbolPool.DirtyCriteria = True
                        End If
                        frmStatus.tmrRecalc.Enabled = True
                    End If
                End If
            End If
    
        Case "ID_About"
            ShowForm frmAbout, True
            
        Case "ID_TradeFilter"
            ShowTradeFilter
            
        Case "ID_ProgramFiles"
            frmFileInfo.ShowMe
    
        Case "ID_ExportData"
            If HasModule("EXP") Then
                ShowForm frmExport, True, , , ALT_GRID_ROW_COLOR
            ElseIf HasGold(True, , False) Then
                ShowForm frmExport, True, , , ALT_GRID_ROW_COLOR
            End If
    
        Case "ID_InstallData"
            If Not ProcessIsBusy Then
                If Not FileExist(App.Path & "\Provided\Install.CFG") Then
                    ''ShowForm frmDataInstall, True
                    Beep
                Else
                    InstallData
                End If
            End If
            
        Case "ID_ImportLibrary"
            If Not ProcessIsBusy Then
                ImportLibrary
            End If
            
        Case "ID_BackupRestore"
            strFile = App.Path & "\TNArchive.exe"
            If Not FileExist(strFile) Then
                InfBox "Archive program not found:" & vbCrLf & strFile, "e", , "Error"
            ElseIf ProcessIsBusy Then
                Beep
            Else
                'RunProcess strFile ', Chr(34) & g.strTitle & Chr(34)
                strTemp = "This program must first be shut down in order to Backup or Restore all your settings."
                If InfBox(strTemp, "?", "ShutDown|+-Cancel", "Backup/Restore Settings") <> "C" Then
                    g.strRunWhenExit = Chr(34) & strFile & Chr(34) '& " " & Chr(34) & g.strTitle & Chr(34)
                    Unload frmSource
                    Set frmSource = Nothing '(so won't turn redraw back on)
                End If
            End If
            
        Case "ID_ZipTradeLogs"
            strTemp = "Zip up how many days of Trade Logs?"
            i = Val(InfBox(strTemp, "?", , "Zip Trade Logs", , , , , , "N", "7"))
            If i > 0 Then
                If g.RealTime.SalmonIsRunning Then
                    DumpSymbolState_SalmonDLL
                End If
                strPath = AddSlash(FilePath(App.Path))
                strFile = strPath & "TradeLogs.zip"
                InfBox "Zipping the trade log files ...", "t", , "Processing", True
                KillFile strFile
                'ZipExecute "z", strFile, strPath, "* /i=*20*.log,error*.log /s /n=-" & Str(i), True
                'ZipExecute "z", strFile, App.Path, "* /i=TradeTracker.mdb,*.log"
                ZipExecute "z", strFile, strPath, "* /i=*.log,R20*.000 /s /n=-" & Str(i), True
                ZipExecute "z", strFile, App.Path, "* /i=TradeTracker.mdb"
                ZipExecute "z", strFile, App.Path, "* /i=debug.log,dlog.log,dlog1.log"
                ZipExecute "z", strPath & "TTBak.zip", App.Path, "TTBak?.mdb"
                InfBox "The trade logs have been zipped into|" & strFile, "i", , "Trade Logs"
            End If
        
        Case "ID_ApplicationBackground"
            frmAppBkCfg.ShowMe
  
        Case "ID_Settings"
            frmConfig.ShowMe
  
        Case "ID_Test1"
            If FileExist("C:\Common\*.exe") Then
                ShowForm frmTest, eForm_Nonmodal, frmMain
            End If
        
        Case "ID_Test2"
            If FileExist("C:\Common\*.exe") Then
                ShowForm frmTest2, eForm_Nonmodal, frmMain
            End If
        
        Case "ID_Timers"
            frmTimers.ShowMe
        
        Case "ID_NewAccount"
            frmNewAccount.ShowMe
            
        Case "ID_RebuildRollFiles"
            frmBuildRolls.ShowMe
        
        Case "ID_DataComparison"
            frmBuildRolls.ShowDiffs
            
        Case "ID_Replay"
            frmGameModeCfg.ShowMe
            
        'Case "ID_Monitoring"
            'If Not FormIsLoaded("frmMonitoring") Then
            '    frmMonitoring.ShowMe
            'End If
        
        Case "ID_COTReport"
            If Not ProcessIsBusy Then
                If HasModule("CT") Then
                    frmCotSettings.ShowMe
                Else
                    InfBox "This report is not valid when 'Commitment of Traders' data is not being updated.||(contact Genesis Sales for current updating cost)", "i", , "COT Report"
                    'If InfBox("This report is only valid if you are updating 'Commitment of Traders' data.", "i", "Proceed|+-Cancel", "COT Report") = "P" Then
                    '    frmCotSettings.ShowMe
                    'End If
                End If
            End If
            
        Case "ID_SAIReport"
            frmSaiReport.ShowMe
            
        Case "ID_SAIElite"
            frmSaiElite.ShowMe
            
        Case "ID_RollsTable"
            frmRollsTable.ShowMe
        
        Case "ID_DanCodeWeb"
            strTemp = DanielCodeProcess
            If Len(strTemp) > 0 Then
                If Not FileExist(strTemp) Then
                    strTemp = ""
                End If
            End If
            
            If Len(strTemp) > 0 Then
                StartDanielCodeProcess
            Else
                'strTemp = GetProvidedProperty("DanielCode2")
                'RunWebReport "Danielcode Trade Signals", strTemp, "kDanielCode", 0
            End If
            'Set frm = New frmWebReport
            'frm.ShowMe "Danielcode Trade Signals", ToolbarIcon("ID_DanCodeWeb")
            'Set frm = Nothing
        
        Case "ID_GmajPro"
            strTemp = GmajProcess
            If Len(strTemp) > 0 Then
                If Not FileExist(strTemp) Then
                    strTemp = ""
                End If
            End If
            
            If Len(strTemp) > 0 Then
                StartGmajProcess
            End If
        
        Case "ID_SectorWeb"
            strTemp = GetProvidedProperty("SectorWeb")
            RunWebReport "Sector Analysis", strTemp, "kSectorAnalysis", 2

        Case "ID_ScreenerWeb"
            strTemp = GetProvidedProperty("ScreenerWeb")
            RunWebReport "Stock Screener", strTemp, "kScreenerWeb", 2
        
        Case "ID_Alerts"
            If frmAlertsSetup.WindowState = vbMinimized Then
                frmAlertsSetup.WindowState = 0
            End If
            If frmAlertsSetup.Visible Then
                MoveFocus frmAlertsSetup
            Else
                frmAlertsSetup.ShowMe
            End If
        
        Case "ID_Quote"
            If Tool.State = ssChecked Then
                DockState(frmQuotes) = eShowAsPrevious
                'ShowForm frmQuotes
                'If frmQuotes.fgQuotes.Rows <= frmQuotes.fgQuotes.FixedRows Then
                    'frmQuotes.EditList
                    'frmQuotes.LoadGrid True
                'End If
            Else
                'frmQuotes.Hide
                DockState(frmQuotes) = eHidden
            End If
            AutoSizeChart
        
        Case "ID_SymbolGrid"
            If Tool.State = ssChecked Then
                DockState(frmSymbolGrid) = eShowAsPrevious
            Else
                DockState(frmSymbolGrid) = eHidden
            End If
            AutoSizeChart
        
        Case "ID_ChartOnOff"
            If Tool.State = ssChecked Then
                DockState(frmChartOnOff) = eShowAsPrevious
                frmChartOnOff.ShowData
            Else
                DockState(frmChartOnOff) = eHidden
            End If
        
        Case "ID_ChartData"
            If Tool.State = ssChecked Then
                DockState(frmChartData) = eShowAsPrevious
                frmChartData.ShowData -1
            Else
                DockState(frmChartData) = eHidden
            End If
    
        Case "ID_PlanetData"
            If Tool.State = ssChecked Then
                DockState(frmPlanetData) = eShowAsPrevious
                frmPlanetData.ShowData -1
            Else
                DockState(frmPlanetData) = eHidden
            End If
    
        Case "ID_Orders"
            If Tool.State = ssChecked Then
                DockState(frmOrderTracker) = eShowAsPrevious
            Else
                DockState(frmOrderTracker) = eHidden
                ''AutoHideStatusForm
            End If
    
        Case "ID_ReloadSymbols"
            If Not ProcessIsBusy Then
                InfBox "w=NOWAIT ; i=t ; Loading Symbols ..."
                g.SymbolPool.Load False
                InfBox ""
            End If

        Case "ID_Download"
            If Not ProcessIsBusy Then
                ShowForm frmDownload, True
            End If
    
        Case "ID_TradeTracker"
            ''frmTTAccounts.ShowMe True
            'If Tool.State = ssChecked Then         -JM: originally a state button leave awhile then remove 01-12-2010
            If DockState(frmTTSummary) = eHidden Then
                i = GetIniFileProperty("TradeConsoleMsg", 0, "General", g.strIniFile)
                If i < 1 Then
                    strTemp = App.Path & "\Info\TradeConsole.rtf"
                    If FileExist(strTemp) Then
                        Set frm = New frmMessage
                        frm.ShowMe "Trade Console", "@" & strTemp, eNormalMessage
                        Set frm = Nothing
                        SetIniFileProperty "TradeConsoleMsg", i + 1, "General", g.strIniFile
                    End If
                End If
                DockState(frmTTSummary) = eShowAsPrevious
            Else
                DockState(frmTTSummary) = eHidden
            End If
            g.ConsoleForms.ShowForms
            AutoSizeChart
    
        Case "ID_Eta"
            GoToETA
            
        Case "ID_CloseWindow"
            ' do this so it will act as if user clicked on "X" of window
            ' (so form's QueryUnload will handle it correctly).
            SendMessage frmActive.hWnd, WM_CLOSE, ByVal 0&, ByVal 0&
        
        Case "ID_ArrangeIcons"
            frmMain.Arrange vbArrangeIcons
        Case "ID_Tile"
            'TileCharts
            If Not ActiveChart Is Nothing Then
                If ActiveChart.DetachStatus = eNotDetached Then
                    frmArrange.ShowMe
                End If
            End If
            
        Case "ID_Cascade"
            TileCharts True
                       
        Case "ID_Toolbox"
            frmToolbox.ShowMe
        
        Case "ID_SymbolGroups"
            frmToolbox.ShowMe eTab_SymbolGroups
        
        Case "ID_Criteria"
            frmToolbox.ShowMe eTab_Criteria
        
        Case "ID_Filters"
            frmToolbox.ShowMe eTab_Filters
            
        Case "ID_Functions"
            frmToolbox.ShowMe eTab_Functions
        
        Case "ID_Rules"
            frmToolbox.ShowMe eTab_Rules
        
        Case "ID_Strategies"
            frmToolbox.ShowMe eTab_Systems
        
        Case "ID_StrategyBaskets"
            frmToolbox.ShowMe eTab_StrategyBaskets
        
        Case "ID_Libraries"
            frmToolbox.ShowMe eTab_Libraries
            
        Case "ID_ManageAccounts"
            frmTTAccounts.ShowMe
        
        Case "ID_ManageAutoExits"
            frmOrderStrategies.ShowMeManage
            
        Case "ID_TradingPerformanceReports"
            frmTradeReportFilter.ShowMe
            
        Case "ID_ViewJournals"
            g.TnJournal.ShowJournals
            
        Case "ID_Cattle"
            g.CattleBridge.ShowCattleForm False
            
        Case "ID_Turnkey"
            g.CattleBridge.ShowCattleForm True
        
        Case "ID_TurnkeyAdministration"
            g.CattleBridge.ShowCattleAdminForm
        
        Case "ID_SeasonalSP"
            strTemp = GetProvidedProperty("SeasonalWeb")
            RunWebReport "Seasonal Sweet Spot", strTemp, "kSeasonalSP", 2
        
        Case Else:
            'apply page?
            If Left(Tool.ID, 6) = "Page #" Then
                strFile = Replace(Tool.Name, "&&", "&")
                If UCase(strFile) = "(UNNAMED)" Then
                    Beep
                ElseIf InStr(strFile, "<") > 0 Then
                    If InStr(UCase(strFile), "SAVE") > 0 Then
                        SaveChartPage ""
                    ElseIf InStr(UCase(strFile), "NEW") > 0 Or InStr(UCase(strFile), "MY PAGE") > 0 Then
                        LoadChartPage ""
                    Else
                        frmTemplates.ShowMe eMode_Pages
                    End If
                Else
                    LoadChartPage strFile
                End If
            ElseIf UCase(Left(Tool.ID, 11)) = "ID_TRADING_" Then
                g.Broker.HandleTradingMenu Tool.ID
            Else
                bKeepChecking = True
            End If
    End Select
    
    If bKeepChecking = False Then
        GoTo Cleanup '(nothing more to look for)
    End If
    
    '3)----------------------------------------------------------
    'Check for loading a form with a symbol from the active form (if exists)
    strTemp = "ID_Chart,ID_Chain,ID_Snapshot,ID_SectorBrowser,ID_TickDistribution,ID_MarketDepth,ID_TimeSales,ID_VolumeAtPrice,ID_TimeSalesAnalyzer,ID_BidAskDir,ID_MarketProfile,ID_NewsBrowser"
    If InStr(strTemp & ",", Tool.ID & ",") > 0 Then
    
        ' first get symbol from active form (if exists)
        strSymbol = ""
        nSymbolID = 0
        On Error Resume Next
        
        If TypeOf frmActive Is frmMain Or IsFrmChart(frmActive) Then
            Set frmActive = ActiveChart
        End If
        nSymbolID = frmActive.SymbolID
        
        On Error GoTo ErrSection:
        If nSymbolID <> 0 Then
            strSymbol = GetSymbol(nSymbolID)
        End If
        
        ' then load the requested form
        bKeepChecking = False
        Select Case Tool.ID
            Case "ID_TickDistribution", "ID_MarketDepth"
                ' get symbol of active chart
                i = 0
                If Not ActiveChart Is Nothing Then
                    i = ActiveChart.Chart.SymbolID
                    ' use symbol of active chart as default
                    If nSymbolID = 0 Then nSymbolID = i
                End If
                ' if symbol does not match active chart,
                ' ask user to confirm desired symbol
                'If nSymbolID = 0 Or nSymbolID <> i Then
                If 1 Then
                    If Tool.ID = "ID_MarketDepth" Then
                        If g.RealTime.Active Then
                            Set aStrings = frmSymbolSelector.ShowMe(g.SymbolPool.SymbolForID(nSymbolID), _
                                False, True, "Get Market Depth for ...")
                        Else
                            InfBox "Realtime streaming must be active.", "!", , "Market Depth"
                        End If
                    Else
                        Set aStrings = frmSymbolSelector.ShowMe(g.SymbolPool.SymbolForID(nSymbolID), _
                            False, True, "Get Price Ladder for ...")
                    End If
                    If aStrings.Size > 0 Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(aStrings(0))
                    Else
                        nSymbolID = 0
                    End If
                End If
                If nSymbolID <> 0 Then
                    Set frm = New frmTickDistribution
                    frm.ShowMe nSymbolID, (Tool.ID = "ID_MarketDepth")
                End If
        
            Case "ID_TimeSales"
                ' get symbol of active chart
                i = 0
                If Not ActiveChart Is Nothing Then
                    i = ActiveChart.Chart.SymbolID
                    ' use symbol of active chart as default
                    If nSymbolID = 0 Then nSymbolID = i
                End If
                ' if symbol does not match active chart,
                ' ask user to confirm desired symbol
                'If nSymbolID = 0 Or nSymbolID <> i Then
                If 1 Then
                    Set aStrings = frmSymbolSelector.ShowMe(g.SymbolPool.SymbolForID(nSymbolID), _
                            False, True, "Get Time and Sales for ...")
                    If aStrings.Size > 0 Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(aStrings(0))
                    Else
                        nSymbolID = 0
                    End If
                End If
                If nSymbolID <> 0 Then
                    Set frm = New frmTimeSales
                    frm.ShowMe nSymbolID
                End If
        
            Case "ID_TimeSalesAnalyzer"
                ' get symbol of active chart
                i = 0
                If Not ActiveChart Is Nothing Then
                    i = ActiveChart.Chart.SymbolID
                    ' use symbol of active chart as default
                    If nSymbolID = 0 Then nSymbolID = i
                End If
                ' if symbol does not match active chart,
                ' ask user to confirm desired symbol
                'If nSymbolID = 0 Or nSymbolID <> i Then
                If 1 Then
                    Set aStrings = frmSymbolSelector.ShowMe(g.SymbolPool.SymbolForID(nSymbolID), _
                            False, True, "Time and Sales Analysis for ...")
                    If aStrings.Size > 0 Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(aStrings(0))
                    Else
                        nSymbolID = 0
                    End If
                End If
                If nSymbolID <> 0 Then
                    Set frm = New frmTimeSalesAnalyzer
                    frm.ShowMe GetSymbol(nSymbolID)
                End If
        
            Case "ID_VolumeAtPrice"
                ' get symbol of active chart
                i = 0
                If Not ActiveChart Is Nothing Then
                    i = ActiveChart.Chart.SymbolID
                    ' use symbol of active chart as default
                    If nSymbolID = 0 Then nSymbolID = i
                End If
                ' if symbol does not match active chart,
                ' ask user to confirm desired symbol
                'If nSymbolID = 0 Or nSymbolID <> i Then
                If 1 Then
                    Set aStrings = frmSymbolSelector.ShowMe(g.SymbolPool.SymbolForID(nSymbolID), _
                            False, True, "Get Volume At Price for ...")
                    If aStrings.Size > 0 Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(aStrings(0))
                    Else
                        nSymbolID = 0
                    End If
                End If
                If nSymbolID <> 0 Then
                    Set frm = New frmPriceVol
                    frm.ShowMe GetSymbol(nSymbolID)
                End If
            
            Case "ID_MarketProfile"
                ' get symbol of active chart
                i = 0
                If Not ActiveChart Is Nothing Then
                    i = ActiveChart.Chart.SymbolID
                    ' use symbol of active chart as default
                    If nSymbolID = 0 Then nSymbolID = i
                End If
                ' if symbol does not match active chart,
                ' ask user to confirm desired symbol
                'If nSymbolID = 0 Or nSymbolID <> i Then
                If 1 Then
                    Set aStrings = frmSymbolSelector.ShowMe(g.SymbolPool.SymbolForID(nSymbolID), _
                            False, True, "Show Trade Profile for ...")
                    If aStrings.Size > 0 Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(aStrings(0))
                    Else
                        nSymbolID = 0
                    End If
                End If
                If nSymbolID <> 0 Then
                    Set frm = New frmMarketProfile
                    frm.ShowMe nSymbolID
                End If
            
            Case "ID_BidAskDir"
                ' get symbol of active chart
                i = 0
                If Not ActiveChart Is Nothing Then
                    i = ActiveChart.Chart.SymbolID
                    ' use symbol of active chart as default
                    If nSymbolID = 0 Then nSymbolID = i
                End If
                ' if symbol does not match active chart,
                ' ask user to confirm desired symbol
                'If nSymbolID = 0 Or nSymbolID <> i Then
                If 1 Then
                    Set aStrings = frmSymbolSelector.ShowMe(g.SymbolPool.SymbolForID(nSymbolID), _
                            False, True, "Get Volume At Price for ...")
                    If aStrings.Size > 0 Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(aStrings(0))
                    Else
                        nSymbolID = 0
                    End If
                End If
                If nSymbolID <> 0 Then
                    Set frm = New frmBidAskDir
                    frm.ShowMe GetSymbol(nSymbolID)
                End If
        
            Case "ID_Chain"
                ' see if running old Option Chain or new Option Navigator
                strTemp = OptNavIP
                If HasModule("OPTNAV") And Len(strTemp) > 0 And FileExist(OptNavExeFile) Then
                    StartOptNav     'this routine will give focus to OptNav or start it as appropriate
                    Tool.State = ssUnchecked
                ElseIf Tool.State = ssUnchecked Then
                    DockState(frmOptionChain) = eHidden
                Else
                    ' get symbol of active chart
                    i = 0
                    If Not ActiveChart Is Nothing Then
                        i = ActiveChart.Chart.SymbolID
                        ' use symbol of active chart as default
                        If nSymbolID = 0 Then nSymbolID = i
                    End If
                    Set aStrings = frmSymbolSelector.ShowMe(g.SymbolPool.SymbolForID(nSymbolID), _
                            False, True, "Get Option Chain for ...")
                    If aStrings.Size > 0 Then
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(aStrings(0))
                    Else
                        nSymbolID = 0
                    End If
                    If nSymbolID = 0 Then
                        'Beep
                        Tool.State = ssUnchecked
                    ElseIf Not ProcessIsBusy Then
                        frmOptionChain.ShowMe nSymbolID
                        DockState(frmOptionChain) = eShowAsPrevious
                    End If
                End If
          
            Case "ID_Chart"
                If frmActive Is Nothing Then
                    strSymbol = ""
                ElseIf IsFrmChart(frmActive) Then
                    strSymbol = frmActive.Chart.SpreadSymbols
                    If Len(strSymbol) = 0 Then
                        strSymbol = frmActive.Chart.Symbol
                    End If
                End If
If 1 Then ' FileExist(App.Path & "\SpreadChart.flg") Then
                frmNewChart.ShowMe strSymbol
Else
                Set aStrings = frmSymbolSelector.ShowMe(strSymbol, False, True, "Symbol for New Chart", True)
                If aStrings.Size > 0 Then
                    Set frm = Nothing
                    If InStr(aStrings(0), "|") > 0 Then
                        Set frm = New frmChart          'new chart is always non-detached
                        frm.Chart.SetSymbol aStrings(0)
                    Else
                        nSymbolID = g.SymbolPool.SymbolIDforSymbol(aStrings(0))
                        If nSymbolID <> 0 Then
                            Set frm = New frmChart          'new chart is always non-detached
                            frm.Chart.SetSymbol nSymbolID
                        End If
                    End If
                    If frm Is Nothing Then
                        Beep
                    Else
                        frm.Chart.ShowTrades = False
                        frm.Chart.GenerateChart
                        ShowForm frm
                        Set frm = Nothing
                    End If
                End If
End If
          
            Case "ID_Snapshot"
                If Tool.State <> ssChecked Then
                    DockState(frmSnapshot) = eHidden
                Else
                    frmSnapshot.ShowMe
                End If
    
            Case "ID_SectorBrowser"
                If Tool.State = ssChecked Then
                    frmSectorTree.ShowMe strSymbol
                ElseIf FormIsLoaded("frmSectorTree") Then
                    frmSectorTree.Hide
                End If
            
            Case "ID_NewsBrowser"
                'StartNewsBrowser
                strTemp = GetProvidedProperty("NewsWeb")
                If Len(strSymbol) > 0 Then
                    strTemp = strTemp & vbTab & "&S=" & strSymbol
                End If
                RunWebReport "News Navigator", strTemp, "kNewsBrowser", 2
                
            Case Else
                bKeepChecking = True
        End Select
    End If

    If frmActive Is Nothing Or bKeepChecking = False Then
        GoTo Cleanup '(nothing more to look for)
    End If
        
    With frmActive
        '4)----------------------------------------------------------
        'Check for other non-chart-specific things from an Active form
        If bKeepChecking Then
            bKeepChecking = False
            Select Case Tool.ID
                Case "ID_TextIncrease"
                    GridTextIncrease
                Case "ID_TextDecrease"
                    GridTextDecrease
                
                Case Else
                    bKeepChecking = True
            End Select
        End If

        '5)----------------------------------------------------------
        'Now check for CHART-SPECIFIC types of things
        If bKeepChecking And bChart Then
            If .Chart.TypeOfChart = eTypeChart_Seasonal Then
                If Tool.ID = "ID_Rectangle" And eExtraInfo = eTbExtraInfo_PFPNewPattern Then
                    InfBox kSeasonalUnavail, "I", "Ok", "Seasonal chart"
                Else
                    ToolbarClickSeasonal frmActive, tbToolbar, Tool
                End If
            Else
                Select Case Tool.ID
                    Case "ID_RepeatDraw"
                        ToolbarSyncCursorGroup tbToolbar, Tool.ID
                    
                    Case "ID_Magnet"
                        i = Abs(g.ChartGlobals.nMagnetValue)
                        If i = 0 Then i = 5
                        If Tool.State = ssUnchecked Then
                            i = -i
                        End If
                        g.ChartGlobals.nMagnetValue = i
                        ToolbarSyncCursorGroup tbToolbar, Tool.ID
                
                    Case "ID_AutoScale"
                        If Tool.State = ssChecked Then
                            .Chart.AutoScale = True
                        Else
                            .Chart.AutoScale = False
                        End If
                        .Chart.GenerateChart eRedo1_Scrolled
                
                    Case "ID_Performance"
                        .Chart.ShowSystemReport False
                
                    Case "ID_EditChart"
                        .TopMost = False
                        .tmr.Tag = "EditSettings"
                
                    Case "ID_AddToChart"
                        .TopMost = False
                        .tmr.Tag = "AddItem"
                
                    Case "ID_TopMost"
                        If Tool.State = ssChecked Then
                            .TopMost = True
                        Else
                            .TopMost = False
                        End If
                        
                    Case "ID_Delete/Hide"
                        If g.ChartGlobals.nHideAnnotations = 0 Then
                            tbToolbar.Tools("ID_HideAnnotations").State = ssUnchecked
                        Else
                            tbToolbar.Tools("ID_HideAnnotations").State = ssChecked
                        End If
                        
                    Case "ID_DeleteLastAnnotation"
                        If .Chart.RemoveAnnots(True) > 0 Then
                            '.Chart.GenerateChart eRedo1_Scrolled
                            .Chart.SyncGlobalAnnots Nothing, True
                        Else
                            Beep
                        End If
                        
                    Case "ID_DeleteAllAnnotations"
                        If .Chart.RemoveAnnots(False) > 0 Then
                            '.Chart.GenerateChart eRedo1_Scrolled
                            .Chart.SyncGlobalAnnots Nothing, True
                        Else
                            Beep
                        End If
                        
                    Case "ID_HideAnnotations"
                        HideAnnotations Tool.State
                        
                    Case "ID_Trendline", "ID_Trendline2", "ID_Trendline3", "ID_Trendline4", _
                            "ID_DollarLine", "ID_DollarLine2", "ID_DollarLine3", "ID_DollarLine4", "ID_Icon", _
                            "ID_SRLine", "ID_SRLine2", "ID_SRLine3", "ID_SRLine4", _
                            "ID_HorzLine", "ID_HorzLine2", "ID_HorzLine3", "ID_HorzLine4", _
                            "ID_VertLine", "ID_Fibonacci", "ID_TargetShooter", "ID_ElliotLabels", _
                            "ID_Rectangle", "ID_Ellipse", "ID_DNRetracement", "ID_DNExpansion", _
                            "ID_DNExpansion2", "ID_DNExpansion3", "ID_DNExpansion4", _
                            "ID_Fibonacci2", "ID_Fibonacci3", "ID_Fibonacci4", _
                            "ID_RegressionLine", "ID_Text", "ID_Text2", "ID_Text3", "ID_Text4", _
                            "ID_AndrewFork", "ID_GannLines", "ID_FibCircle", _
                            "ID_TimeCycle", "ID_FibTimeZones", "ID_FibTimeRatio", _
                            "ID_FibFan", "ID_FibExpansion", "ID_SpResistFan", "ID_Mirror", _
                            "ID_Pattern", "ID_RiskReward", "ID_Triangle", "ID_ChannelHighlight", _
                            "ID_WaveLabels", "ID_ElliotTimeRatio", "ID_Bracket", "ID_PivotPoints", _
                            "ID_ArrowLine", "ID_TrendChannel", "ID_PriceAlert", "ID_DanCodeFib", _
                            "ID_Hawkeye", "ID_FibClusters", "ID_FibABCD", "ID_Gartley", "ID_DanCodeZone", _
                            "ID_GannacciSwingSquare", "ID_GannacciCycle", "ID_GannacciTime", "ID_GannacciSwing1", _
                            "ID_GannacciSwing2", "ID_ElliotEndUser", "ID_BalloonStrangle", "ID_AdvRiskReward"
                        'turn on annotations if they are hidden
                        If g.ChartGlobals.nHideAnnotations = 1 Then HideAnnotations False
                        If Tool.ID = "ID_Rectangle" And eExtraInfo = eTbExtraInfo_PFPNewPattern Then
                            ToolbarSetCursorGroup tbToolbar, True, Tool.ID & "_PFP"
                        Else
                            ToolbarSetCursorGroup tbToolbar, True, Tool.ID
                        End If
                    
                    Case "ID_ShowEWI"
                        If .Chart.HasHiddenAnnots(eANNOT_ElliotLabel) = 0 Then
                            .Chart.ShowHideAnnotsByType eANNOT_ElliotLabel, 1, True
                            .Chart.GenerateChart eRedo1_Scrolled
                            .Chart.SyncToolbar True
                        ElseIf .Chart.HasHiddenAnnots(eANNOT_ElliotLabel) = 1 Then
                            .Chart.ShowHideAnnotsByType eANNOT_ElliotLabel, 0, True
                            .Chart.GenerateChart eRedo1_Scrolled
                            .Chart.SyncToolbar True
                        End If
                    
                    Case "ID_CursorArrow", "ID_CursorCrosshairs", "ID_CursorHorizLine", "ID_CursorVertLine"
                        ToolbarSetCursorGroup tbToolbar, False, Tool.ID
                    
                    Case "ID_ChartMove"
                        g.ChartGlobals.eChartMode = eMode_Move
                        ToolbarSetCursorGroup tbToolbar, False, Tool.ID
                        
                    Case "ID_DragModeY"
                        If g.ChartGlobals.eDragModeY = eDragModeY_Each Then
                            g.ChartGlobals.eDragModeY = eDragModeY_Both
                        Else
                            g.ChartGlobals.eDragModeY = eDragModeY_Each
                        End If
                        ToolbarSyncCursorGroup tbToolbar, Tool.ID
                    
                    Case "ID_Eraser"
                        g.ChartGlobals.eChartMode = eMode_Erase     '6183
                        ToolbarSetCursorGroup tbToolbar, False
                                                        
                    Case "ID_ChartOrderBuy", "ID_ChartOrderSell"
                        g.ChartGlobals.eChartMode = eMode_ChartOrder
                        ToolbarSetCursorGroup tbToolbar, False, Tool.ID
                        
                    Case "ID_WhatIf"
                        If tbToolbar.Tools("ID_WhatIf").State = ssChecked And Not .Chart.IsInWhatIfMode Then
                            .Chart.ActivateWhatIf
                        ElseIf .Chart.IsInWhatIfMode Then
                            .Chart.DeactivateWhatIf
                        End If
                                        
                    Case "ID_PatternProfit"
                        If Not FormIsLoaded("frmPatternProfit") Then frmPatternProfit.ShowMe
                    
                    Case "ID_IndAnalyst"
                        .OrderbarWrapper eOrdBarMode_PFP
    
                    Case "ID_HBReporter"
                        frmHighlightBarReporter.ShowMe
    
                    Case "ID_ZoomIn"
                        If Len(g.strActiveDraw) = 0 Then
                            StatusMsg "To ZOOM, click on the chart and drag over the area to zoom."
                        End If
                        g.ChartGlobals.eChartMode = eMode_Zoom
                        'StatusMsg 'so will force the "ZOOM" msg
                        ToolbarSetCursorGroup tbToolbar, False
                    
                    Case "ID_ZoomOut"
                        If .Chart.Zoomed Then
                            .Chart.UnzoomChart True
                        End If
                        
                    Case "ID_MoreBars"
                        .Chart.PixelsPerBar = -2
                        .Chart.GenerateChart eRedo1_Scrolled
                    Case "ID_LessBars"
                        .Chart.PixelsPerBar = -1
                        .Chart.GenerateChart eRedo1_Scrolled
                    Case "ID_MoreAboveBelow"
                        Set Pane = .Chart.Tree("PRICE PANE")
                        If Not Pane Is Nothing Then
                            Pane.geIncDecMaxRatio 0.05
                            Pane.geIncDecMinRatio -0.05
                            Set Pane = Nothing
                        End If
                        .Chart.GenerateChart eRedo1_Scrolled
                    Case "ID_LessAboveBelow"
                        Set Pane = .Chart.Tree("PRICE PANE")
                        If Not Pane Is Nothing Then
                            Pane.geIncDecMaxRatio -0.05
                            Pane.geIncDecMinRatio 0.05
                            Set Pane = Nothing
                        End If
                        .Chart.GenerateChart eRedo1_Scrolled
            
                    Case "ID_Crosshair"
                        If IsFrmChart(frmActive) Then
                            'set for all visible charts
                            For i = 0 To Forms.Count - 1
                                Set frm = Forms(i)
                                If IsFrmChart(frm) Then
                                    frm.Chart.SetCursor
                                End If
                            Next
                        'Else
                            'set just for this chart
                            '.Chart.SetCursor
                        End If
    
                    Case "ID_ChartOrderbar"
                        .OrderbarWrapper
                    
                    Case "ID_UndoDraw"
                        .Chart.LastEditedAnnotUndo
                
                    Case "ID_OHLCBars"
                        .Chart.BarDisplayType = eINDIC_OHLC
                        .Chart.GenerateChart
                    Case "ID_Candlesticks"
                        .Chart.BarDisplayType = eINDIC_Candlestick
                        .Chart.GenerateChart
                    Case "ID_BollingerBars"
                        .Chart.BarDisplayType = eINDIC_BollingerBar
                        .Chart.GenerateChart
                    Case "ID_CloseLine"
                        .Chart.BarDisplayType = eINDIC_Line
                        .Chart.GenerateChart
                    Case "ID_Mountain"
                        .Chart.BarDisplayType = eINDIC_Area
                        .Chart.GenerateChart
                    Case "ID_PointFigure"
                        .Chart.ChangeBarPeriod "<P", True
                    Case "ID_Kagi"
                        .Chart.ChangeBarPeriod "<K", True
                    Case "ID_Renko"
                        .Chart.ChangeBarPeriod "<R", True
                
                    Case "ID_BarPeriod"
                        strTemp = Trim(Tool.ComboBox.Text)
                        .Chart.ChangeBarPeriod strTemp
                        MoveFocus frmActive
                        
                    Case "ID_Yearly"
                        .Chart.ChangeBarPeriod "Yearly"
                    Case "ID_Quarterly"
                        .Chart.ChangeBarPeriod "Quarterly"
                    Case "ID_Monthly"
                        .Chart.ChangeBarPeriod "Monthly"
                    Case "ID_Weekly"
                        .Chart.ChangeBarPeriod "Weekly"
                    Case "ID_Daily"
                        .Chart.ChangeBarPeriod "Daily"
                    Case "ID_360minute"
                        .Chart.ChangeBarPeriod "360"
                    Case "ID_240minute"
                        .Chart.ChangeBarPeriod "240"
                    Case "ID_180minute"
                        .Chart.ChangeBarPeriod "180"
                    Case "ID_120minute"
                        .Chart.ChangeBarPeriod "120"
                    Case "ID_90minute"
                        .Chart.ChangeBarPeriod "90"
                    Case "ID_60minute"
                        .Chart.ChangeBarPeriod "60"
                    Case "ID_30minute"
                        .Chart.ChangeBarPeriod "30"
                    Case "ID_15minute"
                        .Chart.ChangeBarPeriod "15"
                    Case "ID_10minute"
                        .Chart.ChangeBarPeriod "10"
                    Case "ID_5minute"
                        .Chart.ChangeBarPeriod "5"
                    Case "ID_3minute"
                        .Chart.ChangeBarPeriod "3"
                    Case "ID_1minute"
                        .Chart.ChangeBarPeriod "1"
                    Case "ID_CustomMinute"
                        .Chart.ChangeBarPeriod Tool.Name
                    Case "ID_CustomPeriod"
                        .Chart.ChangeBarPeriod "Custom"
                    Case "ID_DisplacedMA", "ID_MacdPredictor", "ID_OscPredictor", "ID_DetrendOsc", "ID_DiNapoliMACD", "ID_PrefStoch"
                        .Chart.HandleDinapButtons Tool.ID
                    Case "ID_JPDaily", "ID_JPWeekly", "ID_JPMonthly", "ID_JPQuarterly", "ID_JPExpiration"
                        .Chart.HandleJPButtons Tool.ID
                    Case "ID_ResetChart"
                        If Not frmActive.Chart Is Nothing Then frmActive.Chart.RestoreChartNormal vbKeyReturn
                            
                    Case Else:
                        'apply template?
                        If Left(Tool.ID, 10) = "Template #" Then
                            strFile = Trim(Tool.Name)
                            If InStr(strFile, "<") = 0 Then
                                If Not .Chart.TemplateApply(strFile) Then
                                    Beep
                                End If
                            ElseIf Left(UCase(strFile), 6) = "< COPY" Then
                                .TopMost = False
                                CopySettingsToOtherCharts ActiveChart
                            Else
                                .TopMost = False
                                frmTemplates.ShowMe eMode_Templates, .Chart
                            End If
                        ElseIf Left(Tool.ID, 8) = "Symbol #" Then
                            strSymbol = Parse(Tool.Name, ":", 1)
                            nSymbolID = GetSymbolID(strSymbol)
                            If nSymbolID <> 0 Then
                                .Chart.SetSymbol nSymbolID, True
                            Else
                                Beep
                            End If
                        End If
                End Select
            End If      'end if type of chart is seasonal/or not
            
            .Chart.SetCursor

        End If
    End With

Cleanup:
    'cleanup
    On Error Resume Next
        
    Set frm = Nothing
    Set frmActive = Nothing
    Set aStrings = Nothing
    Set tbToolbar = Nothing
    If Not frmSource Is Nothing Then
        frmSource.tbToolbar.Redraw = True
    End If

    ' temporary fix -- TLB 6/22/2009: since the new toolbars can now take the "focus" (away from the chart),
    ' let's always try to move the focus back to the active chart's chart
    'MoveFocus ActiveChart.pbChart

ErrExit:
    If Tool.Type = ssTypeStateButton And eExtraInfo = eTbExtraInfo_None Then        '5807
        If Tool.State = ssChecked Then
            SyncStateButton Tool.ID, Tool.Group, Tool.Category, eBtnState_Selected
        Else
            SyncStateButton Tool.ID, Tool.Group, Tool.Category, eBtnState_Neutral
        End If
    End If
    Exit Sub
    
ErrSection:
    Set frm = Nothing
    Set frmActive = Nothing
    Set aStrings = Nothing
    Set tbToolbar = Nothing
    If Not frmSource Is Nothing Then
        frmSource.tbToolbar.Redraw = True
    End If
    RaiseError "mMain.ToolBarClick", eGDRaiseError_Raise
    
End Sub

Public Sub SetDropDownTool(DropDownTool As SSTool, SetToTool As SSTool)
On Error Resume Next

    Dim strTip$

Exit Sub
    With SetToTool
        DropDownTool.TagVariant = .ID
        DropDownTool.Picture = .Picture
        
        strTip = .ToolTipText
        If Len(strTip) = 0 Then strTip = .Name
        DropDownTool.ToolTipText = TipStr(strTip)
        
        .State = ssChecked
    End With

End Sub

Public Function Picture16(ByVal strPicture$, Optional ByVal iImageList As Integer = 0, _
    Optional isFormIcon As Boolean = False) As Object

    If iImageList = 0 And g.nTbIconStyle = 1 Then
        If g.nColorTheme = kDarkThemeColor Then
            iImageList = 3
        ElseIf isFormIcon Then
            iImageList = 5
        Else
            iImageList = 4
        End If
    End If
    
    Set Picture16 = g.CoreBridge.Picture16(strPicture, iImageList)

End Function

' Returns true if customer has at least the specified level (Gold, Plat, etc.)
' (can optionally display a "need to upgrade" message)
Public Function HasLevel(ByVal eLevel As eTradeNavLevels, _
        Optional ByVal bShowUpgradeMsg As Boolean = True, _
        Optional ByVal strFeature$ = "This feature") As Boolean
On Error GoTo ErrSection:

    Dim strMsg$, strAnswer$
    
    ' see if need to get authorization string
    If Len(g.strAuthorizationString) = 0 Then
        GetAuthorizationStringFromRegistry
    End If
    
    If g.eTradeNavLevel >= eLevel Then
        HasLevel = True
    ElseIf bShowUpgradeMsg Then
        If ExtremeCharts = 1 Then
            strMsg = strFeature & " is not allowed in this version."
            InfBox strMsg, "i", , "Not Enabled"
        Else
            If g.eTradeNavLevel > eTN0_Unknown Then
                Select Case eLevel
                Case eTN6_Platinum
                    strMsg = "the 'PLATINUM'"
                Case eTN5_Professional
                    strMsg = "the 'PROFESSIONAL' or 'PLATINUM'"
                Case eTN4_Gold
                    strMsg = "at least the 'GOLD'"
                Case eTN3_Standard
                    strMsg = "at least the 'STANDARD'"
                Case eTN2_Lite
                    strMsg = "at least the 'LITE'"
                End Select
            End If
            If Len(strMsg) > 0 Then
                strMsg = strFeature & " requires upgrading to " & strMsg & " version."
            Else
                strMsg = strFeature & " is not allowed in this version."
            End If
            strAnswer = InfBox(strMsg, "i", , "Upgrade Required")
            'strAnswer = InfBox(strMsg, "i", "+More Info|-Cancel", "Upgrade Required")
            If strAnswer = "M" Then
                ' display upgrading information
                frmMessage.ShowMe "Upgrading to 'GOLD' or 'PLATINUM'", _
                    "@" & App.Path & "\Info\ProVersion"
            End If
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.HasLevel"
End Function

' Returns true if customer has any of the specified modules in their IBIS string
' - strModules can be a comma-delimited string of one or more modules
' - a module can be prepended with minus to exclude if has that module -- the
'       order is significant (e.g. "GOLD,-ETA" true if has Gold regardless of ETA,
'       but "-ETA,GOLD" is false if has ETA regardless of Gold)
Public Function HasModule(ByVal strModules As String, Optional ByVal bIncludeSourceCode As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim i&, strModule$, strEnablements$, strSearchFor$
    Dim bMatch As Boolean, bExclude As Boolean
    Dim aModules() As String
    Static iHasETA As Integer
    
   ' see if need to get authorization string
    If Len(g.strAuthorizationString) = 0 Then
        GetAuthorizationStringFromRegistry
    End If
    
    strEnablements = g.strAuthorizationString
    If bIncludeSourceCode Then
        strEnablements = strEnablements & GetSourceCode & ","
    End If
    
    ' need to strip out opposite codes for BetterTrades and Rule1U
    If IsRule1U Then
        strEnablements = Replace(strEnablements, ",ETA,", ",") & "R1U,"
    ElseIf ExtremeCharts >= 1 Then
        strEnablements = Replace(strEnablements, ",R1U,", ",") & "ETA,"
    End If
    
If IsIDE Then
    'strEnablements = strEnablements & "EWL,"
    
    If InStr(strEnablements, "TRAN") = 0 Then
'        strEnablements = strEnablements & "TRAN,"
    End If
'strEnablements = Replace(strEnablements, ",WOODCCI", ",")
'strEnablements = Replace(strEnablements, ",RTE", ",")
End If
    
    HasModule = True 'default
    
    If Len(strModules) > 0 Then
        ' check each module from comma-delimited string
        aModules = Split(UCase(Trim(strModules)), ",")
        For i = 0 To UBound(aModules)
            strModule = Trim(aModules(i))
            If Len(strModule) > 0 Then
                ' check for a "not"
                If Left(strModule, 1) = "-" Then
                    strModule = Mid(strModule, 2)
                    bExclude = True
                Else
                    bExclude = False
                    HasModule = False 'default if at least 1 required module
                End If
                
                ' if a module matches, return true
                If Right(strModule, 1) = "*" Then '(basically a wild card)
                    strSearchFor = "," & Left(strModule, Len(strModule) - 1)
                Else
                    strSearchFor = "," & strModule & ","
                End If
                If InStr(strEnablements, strSearchFor) > 0 Then
                    bMatch = True
                Else
                    ' other special cases
                    Select Case strModule
                    Case "PLAT"
                        If HasLevel(eTN6_Platinum, False) Then
                            bMatch = True
                        End If
                    Case "RTGPLAT", "PROF"
                        If HasLevel(eTN5_Professional, False) Then
                            bMatch = True
                        End If
                    Case "GOLD"
                        If HasLevel(eTN4_Gold, False) Then
                            bMatch = True
                        End If
                    Case "RTGGOLD", "STANDARD"
                        If HasLevel(eTN3_Standard, False) Then
                            bMatch = True
                        End If
                    Case "LITE"
                        If HasLevel(eTN2_Lite, False) Then
                            bMatch = True
                        End If
                    Case "ETA"
                        If Not IsRule1U Then
                            If iHasETA = 0 Then
                                iHasETA = -1
                                ' see if the ETA program exists or has run
                                If Len(GetIniFileProperty("SimutradeDirectory", "", "GENERAL", "navwin.ini")) > 0 Then
                                    iHasETA = 1
                                ElseIf FileExist(App.Path & "\..\Eta\Eta.EXE") Then
                                    iHasETA = 1
                                End If
                            End If
                            If iHasETA = 1 Then
                                bMatch = True
                            End If
                        End If
                    Case "JURIK"
                        bMatch = JurikDllExists
                    'Case "IOAMT" '(temporary override)
                    '    bMatch = FileExist(App.Path & "\IOAMT.FLG")
                    End Select
                End If
                If bMatch Then
                    If bExclude Then
                        HasModule = False
                    Else
                        HasModule = True
                    End If
                    Exit For
                End If
            End If
        Next
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.HasModule", eGDRaiseError_Raise
    Resume 'RH
End Function

Public Function HasPlatinum(Optional ByVal bShowUpgradeMsg As Boolean = True, _
        Optional ByVal strFeature$ = "This feature") As Boolean
On Error GoTo ErrSection:

    Dim strMsg$, strAnswer$, bRealtime As Boolean
    
    If Not g.RealTime Is Nothing Then
        bRealtime = g.RealTime.IsServerActive(True)
    End If
    
    ' check if has PLATINUM/SNV version
    If g.eTradeNavLevel >= eTN5_Professional Then
        HasPlatinum = True
    ElseIf HasModule("PLAT") Then
        HasPlatinum = True
    ElseIf HasModule("RTGPLAT") And bRealtime Then
        HasPlatinum = True
    ElseIf bShowUpgradeMsg Then
        ' show message to user
        strMsg = strFeature & " requires upgrading to the 'PLATINUM' version."
        strAnswer = InfBox(strMsg, "i", "+More Info|-Cancel", "Upgrade Required")
        If strAnswer <> "C" Then
            ' display upgrading information
            frmMessage.ShowMe "Upgrading to 'GOLD' or 'PLATINUM'", _
                "@" & App.Path & "\Info\ProVersion"
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.HasPlatinum", eGDRaiseError_Raise
    
End Function

' This function is from years past and was called for features in the "old Gold"
' -- which is probably more like the "Standard" in the new 6-Levels hierarchy
Public Function HasGold(Optional ByVal bShowUpgradeMsg As Boolean = True, _
        Optional ByVal strFeature$ = "This feature", Optional ByVal bOrExtreme As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim strMsg$, strAnswer$, bRealtime As Boolean
    
    If Not g.RealTime Is Nothing Then
        bRealtime = g.RealTime.IsServerActive(True)
    End If
    
    ' check if has GOLD/PRO (or Platinum) version
    If g.eTradeNavLevel >= eTN3_Standard Then 'TLB: the "older 'Gold'" is the new "Standard"
        HasGold = True
    ElseIf HasModule("GOLD") Then
        HasGold = True
    ElseIf HasModule("RTGGOLD,RTGPLAT") Then ' And bRealtime Then
        HasGold = True
    ElseIf bOrExtreme And HasModule("BTX") Then
        HasGold = True
    ElseIf bShowUpgradeMsg Then
        ' show message to user
        If ExtremeCharts = 1 Then
            strMsg = strFeature & " is not allowed in this version."
            strAnswer = InfBox(strMsg, "i", , "Not Enabled")
        Else
            strMsg = strFeature & " requires upgrading to the 'GOLD' or 'PLATINUM' version."
            strAnswer = InfBox(strMsg, "i", , "Upgrade Required")
            'strAnswer = InfBox(strMsg, "i", "+More Info|-Cancel", "Upgrade Required")
            If strAnswer = "M" Then
                ' display upgrading information
                frmMessage.ShowMe "Upgrading to 'GOLD' or 'PLATINUM'", _
                    "@" & App.Path & "\Info\ProVersion"
            End If
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.HasGold", eGDRaiseError_Raise
    
End Function

Public Sub CheckForSpecialDownloadFiles(ByVal strPath$, Optional ByVal bNewsAndMessages As Boolean = True)
On Error GoTo ErrSection:

    Dim i&, iHighest&, iFile&
    Dim strFile$, strBak$, strDest$, strArgs$, strExt$
    Dim aFiles As New cGdArray
    Dim frm As frmMessage
    
    strPath = AddSlash(strPath)
    
    ' Check for PROVIDED.GZP, CUSTOM.GZP, INFO.GZP
    strFile = strPath & "Provided.GZP"
    If FileExist(strFile) Then
        ZipExecute "U", strFile, App.Path & "\Provided", "", True, True
        strBak = ReplaceFileExt(strFile, ".BAK")
        KillFile strBak
        Name strFile As strBak
    End If
    
    strFile = strPath & "Custom.GZP"
    If FileExist(strFile) Then
        ZipExecute "U", strFile, App.Path & "\Custom", "", True, True
        strBak = ReplaceFileExt(strFile, ".BAK")
        KillFile strBak
        Name strFile As strBak
    End If
    
    strFile = strPath & "Info.GZP"
    If FileExist(strFile) Then
        ZipExecute "U", strFile, App.Path & "\Info", "", True, True
        strBak = ReplaceFileExt(strFile, ".BAK")
        KillFile strBak
        Name strFile As strBak
    End If
    
    strFile = strPath & "Templates.GZP"
    If FileExist(strFile) Then
        ZipExecute "U", strFile, App.Path & "\Charts\Templates", "", True, True
        strBak = ReplaceFileExt(strFile, ".BAK")
        KillFile strBak
        Name strFile As strBak
    End If
    
    strFile = strPath & "GenRT.GZP"
    If FileExist(strFile) Then
        ZipExecute "U", strFile, App.Path & "\..\Realtime\Genesis", "", True, True
        strBak = ReplaceFileExt(strFile, ".BAK")
        KillFile strBak
        Name strFile As strBak
    End If
    
    strFile = strPath & "SymTran.GZP"
    If FileExist(strFile) Then
        strDest = DataPath & "SymTran"
        MakeDir strDest
        ZipExecute "U", strFile, strDest, "", True, True
        strBak = ReplaceFileExt(strFile, ".BAK")
        KillFile strBak
        Name strFile As strBak
    End If
    
    strFile = strPath & "HelpUpd.GZP"
    If FileExist(strFile) Then
        'only update help files if help files already exist -- unless downloaded full help
        If FileExist(App.Path & "\Help\*.chm") Or FileLength(strFile) > 3000000 Then
            ZipExecute "U", strFile, App.Path & "\Help", "", True, True
            strBak = ReplaceFileExt(strFile, ".BAK")
            KillFile strBak
            Name strFile As strBak
        End If
    End If
    
    ' we don't want News, Msg, and WhatsNew to happen before frmMain is visible
    ' (so this gets called again in tmrMain of frmMain for this to get checked)
    If bNewsAndMessages Then
        ' check for a newer "News" file
        strFile = strPath & "News"
        strDest = App.Path & "\Info\News"
        strExt = ".TXT"
        If FileExist(strFile & ".RTF") Then
            strExt = ".RTF"
            KillFile strDest & ".TXT"
        End If
        If FileExist(strFile & ".HTM") Then
            ' make sure they have a browser enabled
            If Len(InternetBrowser) > 0 Then
                strExt = ".HTM"
                KillFile strDest & ".TXT"
                KillFile strDest & ".RTF"
            End If
        End If
        strFile = strFile & strExt
        If FileExist(strFile) Then
            strDest = strDest & strExt
            If FileDate(strFile) > FileDate(strDest) Then
                FileCopy strFile, strDest
                If strExt = ".HTM" Then
                    RunProcess InternetBrowser, Chr(34) & strDest & Chr(34)
                Else
                    frmMessage.ShowMe "NEWS from Genesis", "@" & strDest, , True
                End If
            End If
        End If
        
        ' check for new "MSG" file
        strExt = ".rtf"
        strFile = strPath & "Msg" & strExt
        If Not FileExist(strFile) Then
            strExt = ".txt"
            strFile = strPath & "Msg" & strExt
        End If
        If FileExist(strFile) Then
            ' find highest existing msg #
            MakeDir App.Path & "\Messages", True
            aFiles.GetMatchingFiles App.Path & "\Messages\Msg*" & strExt, False
            For iFile = 0 To aFiles.Size - 1
                i = Val(Mid(aFiles(iFile), 4))
                If i > iHighest Then iHighest = i
            Next
            strDest = App.Path & "\Messages\Msg" & Format(iHighest + 1, "00000") & strExt
            FileCopy strFile, strDest
            KillFile strFile
            Set frm = New frmMessage
            frm.ShowMe "Message:  " & strDest, "@" & strDest
            Set frm = Nothing
        End If
    
        ' check for a newer "WhatsNew" file
        strFile = strPath & "WhatsNew.txt"
        strDest = App.Path & "\Info\WhatsNew.txt"
        If FileExist(strFile) Then
            If FileDate(strFile) > FileDate(strDest) + 0.05 Then '(more than an hour newer)
                ' check for a "required modules" file
                strArgs = Trim(FileToString(strPath & "WhatsNew.REQ", , True))
                If HasModule(strArgs, True) Then
                    ' copy file to Info folder and display it
                    FileCopy strFile, strDest
                    'If ExtremeCharts = 0 Then
                        frmWhatsNew.ShowMe
                    'End If
                End If
            End If
        End If
    End If
    
    ' Check for Batch/Exe files to run
    ' - args:  AppPath, WinSysPath, WinPath, AppVersion
    strArgs = Chr(34) & App.Path & Chr(34) _
        & " " & Chr(34) & WinSysPath(True) & Chr(34) _
        & " " & Chr(34) & WindowsPath(True) & Chr(34) _
        & " " & FormatVersion
    
    ' visible, wait until completes
    strFile = strPath & "RUN.BAT"
    If Not FileExist(strFile) Then strFile = ReplaceFileExt(strFile, ".EXE")
    If FileExist(strFile) Then
        If Not RunProcess(strFile, strArgs, True) Then
            frmStatus.AddDetail "ERROR processing " & strFile
        End If
        strBak = ReplaceFileExt(strFile, ".BAK")
        KillFile strBak
        Name strFile As strBak
    End If

    ' hidden, wait until completes
    strFile = strPath & "HIDDEN.BAT"
    If Not FileExist(strFile) Then strFile = ReplaceFileExt(strFile, ".EXE")
    If FileExist(strFile) Then
        frmStatus.AddDetail "Updating Files"
        If Not RunProcess(strFile, strArgs, True, vbHide) Then
            frmStatus.AddDetail "ERROR processing " & strFile
        End If
        strBak = ReplaceFileExt(strFile, ".BAK")
        KillFile strBak
        Name strFile As strBak
    End If

    ' visible, don't wait until completes
    strFile = strPath & "RUN_A.BAT"
    If Not FileExist(strFile) Then strFile = ReplaceFileExt(strFile, ".EXE")
    If FileExist(strFile) Then
        If Not RunProcess(strFile, strArgs, False) Then
            frmStatus.AddDetail "ERROR processing " & strFile
            '(can only rename if didn't run -- if runs,
            ' don't rename it cause it's still running!)
            strBak = ReplaceFileExt(strFile, ".BAK")
            KillFile strBak
            Name strFile As strBak
        End If
    End If

    ' hidden, don't wait until completes
    strFile = strPath & "HIDDEN_A.BAT"
    If Not FileExist(strFile) Then strFile = ReplaceFileExt(strFile, ".EXE")
    If FileExist(strFile) Then
        If Not RunProcess(strFile, strArgs, False, vbHide) Then
            frmStatus.AddDetail "ERROR processing " & strFile
            '(can only rename if didn't run -- if runs,
            ' don't rename it cause it's still running!)
            strBak = ReplaceFileExt(strFile, ".BAK")
            KillFile strBak
            Name strFile As strBak
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.CheckForSpecialDownloadFiles", eGDRaiseError_Raise
    
End Sub

Public Sub CheckCriteria(Optional ByVal bIfDirtyThenAsk As Boolean = True)
On Error GoTo ErrSection:

    Dim strKey$
    
    ' See if have criteria enabled
    If Not ScansEnabled Then
        Exit Sub
    End If
    
    ' make sure another process isn't in progress
    If ProcessIsBusy Then Exit Sub
    
    If bIfDirtyThenAsk Then
        If Not g.SymbolPool.DirtyCriteria Then Exit Sub
        
        If AskBox("h=Criteria ; i=? ; b=+Recalculate|-Not now ; One or more Criteria have changed,| do you wish to recalculate them now?") = "N" Then
            Exit Sub
        End If
    End If
    
    frmStatus.tmrRecalc.Tag = ""
    frmStatus.tmrRecalc.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.CheckCriteria", eGDRaiseError_Raise
    
End Sub

Public Function WholeGridRow(fg As VSFlexGrid, Optional ByVal nRow& = -1) As String
On Error GoTo ErrSection:

    Dim strRow$, i&
    
    If nRow < 0 Then nRow = fg.Row
    If nRow >= 0 And nRow < fg.Rows Then
        For i = 0 To fg.Cols - 1
            strRow = strRow & vbTab & fg.TextMatrix(nRow, i)
        Next
    End If
    WholeGridRow = Mid(strRow, 2)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.WholeGridRow", eGDRaiseError_Raise
    
End Function

Public Function ProcessIsBusy(Optional ByVal bSkipMessage As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bBusy As Boolean
    
    If g.bStarting Then
        bBusy = True
    ElseIf FormIsLoaded("frmStatus") Then
        If frmStatus.IsBusy Then
            bBusy = True
        Else
            'check one more time after a DoEvents
            DoEvents
            If frmStatus.IsBusy Then
                bBusy = True
            End If
        End If
    End If
    
    If bBusy And Not bSkipMessage And Not g.bStarting Then
        Beep
        If frmStatus.IsBusy And Not frmStatus.Visible Then
            On Error Resume Next
            ShowForm frmStatus, , frmMain
        End If
        InfBox "Please wait until the current processing| has completed ...", "t", "+-OK", "Cannot be performed now", True
    End If

    ProcessIsBusy = bBusy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ProcessIsBusy", eGDRaiseError_Raise
    
End Function

' So tooltip messages don't get too long
Public Function TipStr(ByVal strMsg$) As String
On Error GoTo ErrSection:

    Dim nMaxLen&
    
    nMaxLen = 60
    
    If Len(strMsg) > nMaxLen Then
        strMsg = Left(strMsg, nMaxLen) & "..."
    End If
    TipStr = strMsg

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.TipStr", eGDRaiseError_Raise
    
End Function

Public Function FormatVersion(Optional ByVal bIncludeRevision As Boolean = False, Optional ByVal bIncludeFileDate As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = Format(App.Major, "#0") & "." & Format(App.Minor, "#0")
    
    If bIncludeRevision Then
        strReturn = strReturn & "." & Str(App.Revision)
    End If
    
    If bIncludeFileDate Then
        strReturn = strReturn & " " & DateFormat(FileDate(App.Path & "\" & App.EXEName & ".EXE"), MM_DD_YYYY, HH_MM, AMPM_UPPER)
    End If
    
    FormatVersion = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.FormatVersion", eGDRaiseError_Raise
    
End Function

Private Sub GoToETA()
On Error GoTo ErrSection:

    Dim strPgm$
    
    strPgm = App.Path & "\..\Eta\Eta.exe"
    If Not FileExist(strPgm) Then
        strPgm = GetIniFileProperty("SimutradeDirectory", "", "GENERAL", "navwin.ini")
        If Len(strPgm) > 0 Then
            strPgm = AddSlash(strPgm) & "eta.exe"
        Else
            strPgm = "c:\program files\genesis\eta\eta.exe"
        End If
    End If
    If IsIDE Then
        ' can't run from here or ETA will try to startup another instance of NavSuite.exe
        Err.Raise vbObjectError + 1000, , "ETA cannot be run from the IDE"
    ElseIf Len(strPgm) > 0 And FileExist(strPgm) Then
        Shell Chr(34) & strPgm & Chr(34) & " TRADENAV", vbNormalFocus
    Else
        Err.Raise vbObjectError + 1000, , "Please run the first time setup in your ETA Program"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.GoToETA", eGDRaiseError_Raise
    
End Sub

Public Sub ToolbarSetCursorGroup(tbToolbar As SSActiveToolBars, _
            ByVal bDrawingTool As Boolean, Optional ByVal strName$ = "")
On Error GoTo ErrSection:

    Dim iSaveRedraw&, i&, strText$
    Dim frm As Form
    Static strPrevDraw$
    
    iSaveRedraw = tbToolbar.Redraw
    tbToolbar.Redraw = False
    
    If g.ChartGlobals.eChartMode <> eMode_ChartOrder Then
        g.ChartGlobals.ePrevChartMode = g.ChartGlobals.eChartMode
    End If
    
    If bDrawingTool Then
        If Len(strName) > 0 Then
            If InStr(strName, "PFP") = 0 Then strPrevDraw = strName  '(save for next time)  -5807
        ElseIf Len(strPrevDraw) > 0 Then
            strName = strPrevDraw '(set to last tool used)
        Else
            strName = "ID_Trendline"
        End If
        g.strActiveDraw = strName
        If Len(strName) > 0 Then
            If InStr(strName, "_PFP") Then strName = "ID_Rectangle"
            tbToolbar.Tools(strName).State = ssChecked
            Select Case strName
            Case "ID_DNRetracement"
                StatusMsg "To start drawing a Retracement, click on a focus point ..."
            Case "ID_DNExpansion", "ID_DNExpansion2", "ID_DNExpansion3", "ID_DNExpansion4"
                If UseDiNapFib() Then
                    StatusMsg "To draw Expansion, click on three peaks of the chart ..."
                Else
                    StatusMsg "To draw Extension, click on three peaks of the chart ..."
                End If
            Case "ID_FibABCD"
                StatusMsg "To draw Fib AB=CD, click on three peaks of the chart ..."
            Case "ID_Gartley"
                StatusMsg "To draw Gartley, click on three peaks of the chart ..."
            Case "ID_AndrewFork"
                StatusMsg "To start drawing an Andrews Pitchfork, click on the first point ..."
            Case "ID_Triangle"
                StatusMsg "To start drawing a triangle, click on the first point ..."
            Case "ID_ChannelHighlight"
                StatusMsg "To start drawing a channel, click on the first point ..."
            Case "ID_WaveLabels"
                StatusMsg "To start drawing wave labels, click on the first point ..."
            Case "ID_Icon", "ID_ElliotLabels", "ID_ElliotEndUser", "ID_HorzLine", "ID_HorzLine2", _
                 "ID_HorzLine3", "ID_HorzLine4", "ID_VertLine", "ID_GannLines", "ID_FibTimeZones"
                StatusMsg "To DRAW, click on the chart ..."
            Case Else
                StatusMsg "To DRAW, click and drag on the chart ..."
            End Select
        End If
    Else
        If Len(g.strActiveDraw) > 0 Then
            ' clear any remaining messages from invalid drawing
            'strText = Trim(UCase(frmMain.tbToolbar.Tools("ID_Status").Name))
            'If InStr(strText, "DRAW") > 0 Or Len(strText) = 0 Then
            'If frmMain.tbToolbar.Tools("ID_Status").ForeColor = vbRed Then
                StatusMsg
            'End If
            g.strActiveDraw = ""
        End If
        If Len(strName) > 0 Then
            tbToolbar.Tools(strName).State = ssChecked
        End If
        If g.ChartGlobals.eChartMode = eMode_Move Then
            tbToolbar.Tools("ID_ChartMove").State = ssChecked
        ElseIf g.ChartGlobals.eChartMode = eMode_Erase Then
            tbToolbar.Tools("ID_Eraser").State = ssChecked
        ElseIf g.ChartGlobals.eChartMode = eMode_ChartOrder Then
            'don't need to do anything (this is here to prevent zoomin tool getting set
        Else
            tbToolbar.Tools("ID_ZoomIn").State = ssChecked
        End If
        
        If TypeOf tbToolbar.Parent Is frmMain Then
            'set for all visible charts
            For i = 0 To Forms.Count - 1
                Set frm = Forms(i)
                If IsFrmChart(frm) Then
                    frm.Chart.SetCursor
                End If
            Next
            Set frm = Nothing
        End If
    End If
        
    Set frm = ActiveChart
    If Not frm Is Nothing Then
        frm.ClearAnnotFlags True, False
        If strName = "ID_Icon" Then
            ' show icon pallete
            If Not frmIconAnnot.Visible Then frmIconAnnot.ShowMe frm.Chart
        ElseIf strName = "ID_ElliotLabels" Then
            If Not frmElliot.Visible Then frmElliot.ShowMe Nothing
        ElseIf strName = "ID_ElliotEndUser" Then
            If Not frmElliot.Visible Then frmElliot.ShowMe Nothing, True
        End If
        Set frm = Nothing
    End If
    
    If strName <> "ID_Icon" Then
        If FormIsLoaded("frmIconAnnot") Then Unload frmIconAnnot
    End If
    
    If strName <> "ID_ElliotLabels" And strName <> "ID_ElliotEndUser" Then
        If strName <> "ID_ChartMove" Then
            If FormIsLoaded("frmElliot") Then Unload frmElliot
        End If
    End If
    
ErrExit:
    tbToolbar.Redraw = iSaveRedraw
    Exit Sub
    
ErrSection:
    tbToolbar.Redraw = iSaveRedraw
    RaiseError "mMain.ToolbarSetCursorGroup", eGDRaiseError_Raise
    
End Sub

Private Sub ToolbarWindowList()
On Error GoTo ErrSection:

    Dim i&, iPos&, strText$, strID$, strActive$
    Dim aCharts As New cGdArray
    Dim frm As Form
    
    If Not ActiveChart Is Nothing Then
        strActive = CStr(ActiveChart.hWnd)
    End If
   
    ' build list of current chart windows
    For i = 0 To Forms.Count - 1
        If IsFrmChart(Forms(i)) Then
            Set frm = Forms(i)
            strText = frm.Chart.ChartName(True)
            aCharts.Add strText & vbTab & CStr(frm.hWnd)
        End If
    Next
    aCharts.Sort eGdSort_IgnoreCase
    
    With frmMain.tbToolbar
        .Redraw = False
        ' remove previous window items
        With .Tools("ID_Window").Menu
            For i = .Tools.Count To 1 Step -1
                strText = .Tools(i).Group
                If strText = "WindowList" Then
                    strID = .Tools(i).ID
                    frmMain.tbToolbar.Tools.Remove strID
                End If
            Next
        End With
    
        ' add new window items
        For i = 0 To aCharts.Size - 1
            strText = aCharts(i)
            iPos = InStr(strText, vbTab)
            If iPos > 1 Then
                ' add the window item
                strID = Mid(strText, iPos + 1)
                strText = Left(strText, iPos - 1)
                'strText = Replace(strText, ",   ", ", ")
                'strText = Replace(strText, ",  ", ", ")
                strText = Replace(strText, "(  ", "(")
                strText = Replace(strText, "( ", "(")
                strText = Replace(strText, "&", "&&")
                .Tools.Add strID, ssTypeStateButton
                With .Tools(strID)
                    .Name = strText
                    .Group = "WindowList"
                    .GroupAllowAllUp = False
                    .PictureDown = Picture16(ToolbarIcon("kChecked"))
                    If strActive = strID Then
                        .State = ssChecked
                    End If
                End With
                ' add window item to window list
                .Tools("ID_Window").Menu.Tools.Add strID
            End If
        Next
        .Redraw = True
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ToolbarWindowList", eGDRaiseError_Raise
    
End Sub

' The "HelpItems.MNU" file adds items to the help menu ...
' - Format: Type<tab>Menu Name<tab>Location<tab>Modules (optional comma-delimited string)
' - Types:
'   1 = remote web page (displayed with the browser, requires internet connection)
'   2 = local HTML file (displayed with the browser, no internet connection)
'   3 = local .TXT or .RTF file (displayed in the "News" window)
'   4 = local .HLP file (displayed using the operating system's help program)
' - Location: web address or name of local file
' - Modules: item will only show if user is enabled for one of the specified modules
Private Sub ToolbarHelpList()
On Error GoTo ErrSection:

    Dim bShow As Boolean
    Dim i&, iFld&, strText$, strID$, strIcon$
    Dim nType&, strName$, strLocation$, strModules$
    Dim aItems As New cGdArray
    
       
    aItems.FromFile App.Path & "\Info\HelpItems.mnu"
   
    With frmMain.tbToolbar
        .Redraw = False
        
        ' remove previous help items
        With .Tools("ID_Help").Menu
            For i = .Tools.Count To 1 Step -1
                strText = .Tools(i).Group
                If strText = "HelpList" Then
                    strID = .Tools(i).ID
                    frmMain.tbToolbar.Tools.Remove strID
                End If
            Next
        End With
    
        ' add new help items
        For i = 0 To aItems.Size - 1
            ' parse the line
            strText = Trim(aItems(i))
            nType = Val(Parse(strText, vbTab, 1))
            strName = Parse(strText, vbTab, 2)
            strLocation = Parse(strText, vbTab, 3)
            strModules = Parse(strText, vbTab, 4)
            strIcon = Parse(strText, vbTab, 5)
            
            ' check for valid fields
            If nType < 1 Or nType > 4 Or Len(strText) = 0 Or Len(strLocation) = 0 Then
                bShow = False
            Else
                ' check for required module (optional comma-delimited string)
                bShow = HasModule(strModules, True)
            End If
            
            ' if local file, make sure it exists
            If bShow = True And nType > 1 Then
                If InStr(strLocation, "\") = 0 Then
                    strLocation = App.Path & "\Info\" & strLocation
                ElseIf InStr(strLocation, ":") = 0 Then
                    strLocation = App.Path & "\" & strLocation
                End If
                bShow = FileExist(Parse(strLocation, "|", 1))
            End If
            
            ' add the window item
            If bShow Then
                ' don't show local manual when web page is ready with manuals
                If InStr(UCase(strName), "MANUAL") > 0 Then
                    .Tools("ID_Manual").Visible = False
                End If
                strID = "Help" & CStr(i)
                .Tools.Add strID, ssTypeButton
                With .Tools(strID)
                    .Name = strName
                    .Group = "HelpList"
                    .TagVariant = CStr(nType) & vbTab & strLocation
                    If Len(strIcon) > 0 Then
                        .Picture = LoadPicture(App.Path & "\Info\" & strIcon)
                    Else
                        Select Case nType
                        Case 1
                            .Picture = Picture16(ToolbarIcon("kInternet"))
                        Case 4
                            .Picture = Picture16(ToolbarIcon("kHelp"))
                        Case Else
                            .Picture = Picture16(ToolbarIcon("kInfo"))
                        End Select
                    End If
                End With
                ' add window item to window list
                .Tools("ID_Help").Menu.Tools.Add strID, , .Tools("ID_Help").Menu.Tools.Count - 1
            End If
        Next
        
        .Redraw = True
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ToolbarHelpList", eGDRaiseError_Raise
End Sub

' ToolbarSync should be called at the END of the form's
' Activate event (so rest of event will process before frm.Refresh),
' and in the Unload event (pass bShowing = False)
Public Sub ToolbarSync(frm As Form, Optional ByVal bShowing As Boolean = True)
On Error GoTo ErrSection:

    Dim bSaveRedraw As Boolean

  Exit Sub '(don't need this anymore)

    ' this helps a form display much quicker
    ' (instead of waiting for main toolbar to refresh)
    ''If bShowing Then frm.Refresh
    
    With frmMain.tbToolbar
        bSaveRedraw = .Redraw
        .Redraw = False
            
        ''If Not bShowing Then
        If 0 Then
            If TypeOf frm Is frmSymbolGrid Then
                .Tools("ID_SymbolGrid").State = ssUnchecked
            ElseIf TypeOf frm Is frmQuotes Then
                .Tools("ID_Quote").State = ssUnchecked
            ElseIf TypeOf frm Is frmSnapshot Then
                .Tools("ID_Snapshot").State = ssUnchecked
            ElseIf TypeOf frm Is frmOptionChain Then
                .Tools("ID_Chain").State = ssUnchecked
            End If
        End If
        
        If IsMDIChild(frm) Then
            ''ToolbarShow frm, bShowing
            ''ToolbarWindowList frm, bShowing
        End If
        
        ' to reset the toolbar
        .Enabled = False
        .Enabled = True
        .Redraw = bSaveRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ToolbarSync", eGDRaiseError_Raise
    
End Sub

Public Sub StatusMsg(Optional ByVal strMsg$ = "", Optional ByVal nColor& = -1, _
        Optional ByVal strTooltip$ = "")
On Error Resume Next

    Dim i&
    Static strPrevMsg$, nPrevColor&

    If Len(Trim(strMsg)) = 0 Then
        nColor = &H80000004 '("invisible": menu bar color)
        strMsg = " "
    ElseIf nColor = -1 Then
        'nColor = &H80000008 '(window text)
        'nColor = RGB(0, 0, 128)
        If g.nColorTheme = 1 Or g.nColorTheme = kDarkThemeColor Then
            nColor = vbGreen
        Else
            nColor = RGB(0, 0, 160)
        End If
        'nColor = RGB(128, 0, 0)
    End If
    If strMsg <> strPrevMsg Or nColor <> nPrevColor Then
        With frmMain.tbToolbar.Tools("ID_Status")
            If strMsg <> strPrevMsg Then .ChangeAll ssChangeAllName, strMsg
            If nColor <> nPrevColor Then .ChangeAll ssChangeAllForeColor, nColor
            If Len(Trim(strTooltip)) = 0 Then
                .ToolTipText = " " '(to disable it)
            Else
                .ToolTipText = strTooltip
            End If
        End With
        With frmMain.tbToolbar
            If .Redraw = False Then
                ' do this so display will update
                .Redraw = True
                .Redraw = False
            End If
        End With
        strPrevMsg = strMsg
        nPrevColor = nColor
    End If
    
End Sub

Public Property Get DockState(frm As Form) As eDockStates
    DockState = frmMain.DockPro.State(frm.Name)
End Property

Public Property Let DockState(frm As Form, ByVal State As eDockStates)
    With frmMain.DockPro
        If .State(frm.Name) <> State Then
            If State = eHidden Then
                ShowFormLog frm, False
            Else
                ShowFormLog frm, True
            End If
            .State(frm.Name) = State
        End If
    End With
End Property

Public Sub ShowUndocked(frm As Form, Optional ByVal nLeft& = -1, Optional ByVal nTop& = -1, _
        Optional ByVal nWidth& = -1, Optional ByVal nHeight& = -1)

    With frmMain.DockPro
        .ShowUndocked frm.Name, nLeft, nTop, nWidth, nHeight
    End With

End Sub

Public Sub ActiveChartFormSet(frm As Form)

    If g.bUnloading Then Exit Sub

    If Not frm Is Nothing Then
        If Not m.ActiveChartForm Is Nothing And Not m.ActiveChartForm Is frm Then       '4898, 4904
            If m.ActiveChartForm.tmr.Tag <> "UNLOADING" And m.ActiveChartForm.tmr.Tag <> "UNLOAD_NOW" Then
                m.ActiveChartForm.SkipFocusFix = False
                SendMessage m.ActiveChartForm.hWnd, WM_NCACTIVATE, 0, 0
                m.ActiveChartForm.ClearBuySellButtons True
            End If
            'sync forms
            frmSymbolGrid.SymbolID = frm.Chart.SymbolID
            frmSnapshot.SymbolID = frm.Chart.SymbolID
            
            frmChartData.ShowData -1
            frmPlanetData.ShowData -1
            frmChartOnOff.ShowData frm
        End If
    End If
    
    Set m.ActiveChartForm = frm
    
    If Not frm Is Nothing Then
        frm.SkipFocusFix = False        'reset  - must do this last
        If frm.DetachStatus = eNotDetached Then Set g.ChartGlobals.frmActiveNonDetached = frm
        frmMain.SetWindowLink frm   '5444
    End If
        
End Sub

Public Function ActiveChart() As Form
On Error GoTo ErrSection:

    Dim frm As Form
        
    If m.ActiveChartForm Is Nothing Then
        Set ActiveChart = g.ChartGlobals.frmActiveNonDetached
    Else
        Set ActiveChart = m.ActiveChartForm
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ActiveChart", eGDRaiseError_Raise
    
End Function

' Returns the "form that was just active" (as of 200 milliseconds ago)
' - if click on toolbar from a docked form, etc., this will return that form
' - will return chart if that was the active form
Public Function ActiveForm() As Form
On Error GoTo ErrSection:

    Dim dDiff#

    dDiff = gdTickCount - m.dPrevActiveFormTime
    If dDiff < 200 And Not m.PrevActiveForm Is Nothing Then
        Set ActiveForm = m.PrevActiveForm
    Else
        Set m.PrevActiveForm = Nothing
        Set ActiveForm = Screen.ActiveForm
        If ActiveForm Is frmMain Then
            If Not frmMain.ActiveForm Is Nothing Then
                Set ActiveForm = frmMain.ActiveForm
            End If
        End If
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mMain.ActiveForm", eGDRaiseError_Raise

End Function

Public Sub SetPrevActiveForm(ByVal frm As Form)
On Error GoTo ErrSection:

    m.dPrevActiveFormTime = gdTickCount
    Set m.PrevActiveForm = frm

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.SetPrevActiveForm", eGDRaiseError_Raise
    
End Sub

Public Sub UnloadEditors()
On Error GoTo ErrSection:

    If FormIsLoaded("frmEditAnnot") Then Unload frmEditAnnot
    If FormIsLoaded("frmChartCfg") Then
        If Not frmChartCfg.bNowAdding Then
            Unload frmChartCfg
        End If
    End If
    If FormIsLoaded("frmTbMoreButtons") Then Unload frmTbMoreButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.UnloadEditors", eGDRaiseError_Raise
    
End Sub

' NOTE: this now returns Minutes since midnight but rounded to the nearest second
' (i.e. returns a fractional part of a minute)
Public Function MinutesFromMidnight(ByVal dDateTime As Double) As Double
On Error GoTo ErrSection:

    ' first convert to number of seconds (round to nearest second)
    dDateTime = Int((dDateTime - Int(dDateTime)) * 86400# + 0.5)
    ' then return as number of minutes
    If dDateTime < 86400 Then
        MinutesFromMidnight = dDateTime / 60# 'Int(dDateTime / 60#)
    Else
        ' if was rounded up to midnight, return as if 11:59pm
        MinutesFromMidnight = 1439 + 59 / 60# ' 1439
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.MinutesFromMidnight", eGDRaiseError_Raise
    
End Function

Public Sub CalcNextTryTime(Optional ByVal bFromDownload As Boolean = False)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the registry
    Dim dNewYorkNow As Double           ' Time in New York right now
    Dim dStartTime As Double            ' New York Time to start trying
    Dim dEndTime As Double              ' New York Time to stop trying
    Dim iMinutes As Integer             ' Interval of download attempts
    Dim iHours As Integer               ' Number of hours to try to download
    Dim d#
 
    ' Key into the registry
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    
    ' do auto-download if chosen by user or if real-time is currently active
    If GetRegistryValue(rkLocalMachine, strKey, "AutoUpdate", vbChecked) <> vbUnchecked _
            Or (g.RealTime.Active And g.nReplaySession = 0) Then
                    
        ' Get current time in New York
        dNewYorkNow = ConvertTimeZone(Now)
       
If 0 Then ' OLD method:

        ' Get Start Time from Registry
        dStartTime = CDbl(LastDailyDownload + 1) + GetRegistryValue(rkLocalMachine, strKey, "AutoStart", TimeSerial(18, 30, 0))
        Do While Not IsWeekday(dStartTime - 0.5)
            dStartTime = dStartTime + 1#
        Loop
        ' Convert it to New York time if stored in Local Time
        If GetRegistryValue(rkLocalMachine, strKey, "LocalTime", False) Then
            ' Convert Start Time to New York Time
            dStartTime = ConvertTimeZone(dStartTime)
        End If
        
        ' Calculate the Ending Time from the Start Time
        iMinutes = GetRegistryValue(rkLocalMachine, strKey, "TryInterval", 5)
        iHours = GetRegistryValue(rkLocalMachine, strKey, "TryHours", 2)
        dEndTime = dStartTime + (iHours / 24#)
        
        ' Bump up the End Time until it is greater than the current NY time
        Do While dEndTime < dNewYorkNow
            dStartTime = dStartTime + 1#
            Do While Not IsWeekday(dStartTime - 0.5)
                dStartTime = dStartTime + 1#
            Loop
            dEndTime = dStartTime + (iHours / 24#)
        Loop
        
        ' If called from a download, either bump up to next day or next interval
        If bFromDownload Then
            If dStartTime < dNewYorkNow + (iMinutes / 1440#) Then
                If dNewYorkNow + (iMinutes / 1440#) <= dEndTime Then
                    dStartTime = Int((dNewYorkNow + (iMinutes / 1440#)) * 1440) / 1440#
                Else
                    dStartTime = dStartTime + 1#
                    Do While Not IsWeekday(dStartTime - 0.5)
                        dStartTime = dStartTime + 1#
                    Loop
                End If
            End If
        End If
    
Else ' NEW method 11/17/2014: delayed daily update when streaming, and ignore "end time" ...

        ' Get Start Time from Registry
        dStartTime = GetRegistryValue(rkLocalMachine, strKey, "AutoStart", TimeSerial(18, 30, 0))
        If GetRegistryValue(rkLocalMachine, strKey, "LocalTime", False) Then
            ' Convert to New York Time (if stored in Local time)
            dStartTime = ConvertTimeZone(dStartTime)
        End If
        ' TLB 11/17/2014: to help eliminate bad data from first run being downloaded by clients who are
        ' streaming (e.g. live and/or auto trading, etc), make sure daily download is after 5:15pm MT
        If g.RealTime.Active And g.nReplaySession = 0 Then
            If dStartTime < TimeSerial(19, 15, 0) Then
                ' reset Start time randomly between 7:15-7:45pm ET
                dStartTime = TimeSerial(19, RandomNum(15, 45), 0)
                ' and store that way (so new time will be obvious from frmConfig)
                If GetRegistryValue(rkLocalMachine, strKey, "LocalTime", False) Then
                    d = ConvertTimeZone(dStartTime, "NY", "")
                Else
                    d = dStartTime
                End If
                SetRegistryValue rkLocalMachine, strKey, "AutoStart", d
            End If
        End If
        ' determine start time for the next daily download
        d = LastDailyDownload + 1
        Do While Not IsWeekday(d)
            d = d + 1
        Loop
        dStartTime = d + dStartTime

        ' If called from a download, bump up to next interval
        If bFromDownload Then
            iMinutes = GetRegistryValue(rkLocalMachine, strKey, "TryInterval", 5)
            If dStartTime < dNewYorkNow + (iMinutes / 1440#) Then
                dStartTime = Int((dNewYorkNow + (iMinutes / 1440#)) * 1440) / 1440#
            End If
        End If
End If
        
        ' Convert StartTime to Local Time
        dStartTime = ConvertTimeZone(dStartTime, "NY", "")
        
        ' add random 0-59 seconds (but tie it to their CustID so will stay the same)
        dStartTime = dStartTime + (g.lLCD Mod 60) / 86400#
    End If
    
    ' Set the next try time
    g.dNextDownloadTry = dStartTime
    DebugLog "Next auto Daily Update: " & Format(dStartTime, "YYYYMMDD HH:MM:SS")
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.CalcNextTryTime", eGDRaiseError_Raise
    
End Sub

' TLB 9/7/2004: this version allows for a more random auto-refresh (so
' most customers won't be all refreshing at exactly the same time)
Public Sub CalcNextQuoteRefresh()
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the registry
    Dim dStartTime As Double            ' Time to start refreshing quote board
    Dim dEndTime As Double              ' Time to stop refreshing quote board
    Dim dNextTime As Double             ' Time of next interval
    Dim iInterval As Long               ' Interval to refresh quote board
    Dim dNewYorkNow As Double           ' Time in New York
    Dim strText$

    ' Key into the registry
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    
    If GetRegistryValue(rkLocalMachine, strKey, "AutoQuotes", vbUnchecked) <> vbUnchecked Then
        ' Get current time locally and in New York
        dNewYorkNow = ConvertTimeZone(Now)
        
        ' Get Start and End Times from Registry and make sure they are in NY
        iInterval = CLng(GetRegistryValue(rkLocalMachine, strKey, "QuoteInterval", 10#))
        dStartTime = GetRegistryValue(rkLocalMachine, strKey, "QuoteStart", TimeSerial(8, 20, 0))
        dEndTime = GetRegistryValue(rkLocalMachine, strKey, "QuoteEnd", TimeSerial(16, 30, 0))
        If GetRegistryValue(rkLocalMachine, strKey, "LocalTime", False) Then
            dStartTime = ConvertTimeZone(dStartTime)
            dEndTime = ConvertTimeZone(dEndTime)
        End If
        
        ' start back a couple of days
        dStartTime = Int(dNewYorkNow) + dStartTime - 2
        dEndTime = Int(dNewYorkNow) + dEndTime - 2
        ' move End time forward until end time is a NY weekday in the future
        Do While dEndTime < dNewYorkNow Or Not IsWeekday(dEndTime)
            dEndTime = dEndTime + 1
        Loop
        ' move Start time forward until within a day of the End time
        Do While dStartTime + 1 < dEndTime
            dStartTime = dStartTime + 1
        Loop
        ' convert start time back to local
        dStartTime = ConvertTimeZone(dStartTime, "NY", "")
        
        ' calc next refresh time, must be after the start time (local time)
        If g.dLastQuoteBoardRefresh = 0 Then g.dLastQuoteBoardRefresh = Now
        dNextTime = g.dLastQuoteBoardRefresh
        Do
            dNextTime = dNextTime + iInterval / 1440#
        Loop While dNextTime < dStartTime
    
        ' text for tooltip
        strText = "Next auto-refresh scheduled for " & DateFormat(dNextTime, MM_DD_YY, H_MM_SS, AMPM_LOWER)
        frmTest2.AddList strText
    End If
    
    frmQuotes.cmdRefresh(0).ToolTipText = strText
    g.dNextQuoteBoardRefresh = dNextTime

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.CalcNextQuoteRefresh", eGDRaiseError_Raise
    
End Sub

Public Sub SetActiveChartSymbol(ByVal vSymbol As Variant)
On Error GoTo ErrSection:

    Dim nSymbolID&, strSym$, bNew As Boolean, dTimeout#
    Dim frm As Form
       
    If VarType(vSymbol) = vbString Then
        nSymbolID = g.SymbolPool.SymbolIDforSymbol(vSymbol)
        strSym = vSymbol
    Else
        nSymbolID = CLng(vSymbol)
        strSym = g.SymbolPool.SymbolForID(nSymbolID)
    End If
    If nSymbolID <> 0 Then
        ' first allow any current symbol linking to finish
        dTimeout = gdTickCount + 10000
        Do While frmMain.tmrWindowLink.Enabled
            DoEvents
            If gdTickCount > dTimeout Then Exit Do
        Loop
        ' change symbol on the active chart
        Set frm = ActiveChart
        If frm Is Nothing Then
            Set frm = New frmChart          'new chart is always non-detached
            bNew = True
        End If
        With frm
            If nSymbolID <> .SymbolID Then
                If Left(strSym, 1) = "$" And Not IsForex(strSym) Then
                    .Chart.ShowTrades = False   'aardvark 3959
                End If
                .Chart.SetSymbol nSymbolID
                .Chart.GenerateChart
            End If
            If bNew Then ShowForm frm
        End With
    Else
        Beep
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.SetActiveChartSymbol", eGDRaiseError_Raise
    
End Sub

Public Sub DoPrintHeader(Optional ByVal nFontSize& = 12, Optional ByVal vp As VSPrinter = Nothing)
On Error GoTo ErrSection:

    Dim strText$

    If vp Is Nothing Then
        Set vp = frmPrintPreview.vp
    End If

    With vp
        .LineSpacing = 100
        .HdrFontName = "Times New Roman"
        .HdrFontSize = nFontSize
        .Header = g.TnCore.GetPrintHeader
        .Footer = "|Page: %d|"
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.DoPrintHeader", eGDRaiseError_Raise
    
End Sub

Public Sub SetMainCaption(Optional ByVal strFeedTime$ = "")
On Error GoTo ErrSection:

    Dim strText$, strIcon$, strFile$, s$, d#
    Static strPrevCaption$
    
    If Len(strFeedTime) > 0 Then
        ' if just feed time is changing, then use prev caption
        strText = strPrevCaption & " -- " & strFeedTime
    Else
        ' fix caption
        strIcon = ToolbarIcon("ID_About")       '"kNav"
        strText = g.strTitle
        If IsRule1U Then
            strIcon = "kRule1U"
        ElseIf ExtremeCharts >= 1 Then
            strIcon = "kExtCharts"
            
            ' TLB 2/19/2014: for new Advanced version demo
            If HasModule("BTXA") Then
                strText = strText & " ADVANCED"
                m.iExtremeChartsMode = 2
                g.eTradeNavLevel = eTN6_Platinum ' (just to turn everything on for now)
            End If
            
        Else
            Select Case g.eTradeNavLevel
            Case eTN6_Platinum
                strText = strText & " PLATINUM"
            Case eTN5_Professional
                strText = strText & " PROFESSIONAL"
            Case eTN4_Gold
                strText = strText & " GOLD"
            Case eTN3_Standard
                If HasModule("LWMC", True) Then
                    strText = strText & " MONEY CODE"
                ElseIf InStr(UCase(strText), "TRADESMART") > 0 Then
                    strText = strText & " ELITE"
                Else
                    strText = strText & " STANDARD"
                End If
            Case eTN2_Lite
                If HasModule("LWMC", True) Then
                    strText = strText & " MONEY CODE"
                Else
                    strText = strText & " LITE"
                End If
            Case eTN1_Silver
                If HasModule("LWMC", True) Then
                    strText = strText & " MONEY CODE"
                ElseIf InStr(UCase(strText), "TRADESMART") > 0 Then
                    strText = strText & " PORTFOLIO"
                Else
                    strText = strText & " SILVER"
                End If
            End Select
            If UCase(Left(strText, 4)) = "WOOD" Then
                strIcon = ToolbarIcon("ID_TradeFilter")
            End If
        End If
        strText = strText & " v" & FormatVersion
        
        If FileExist(App.Path & "\TitleBar.txt") Then
            strText = strText & FileToString(App.Path & "\TitleBar.txt", , True)
        ElseIf Len(g.strChartPage) > 0 Then
            ' see if chart page still exists
            If Not FileExist(g.ChartGlobals.strCPCRoot & "\Charts\Pages\" & g.strChartPage & ".GZP") Then
                g.strChartPage = ""
            Else
                strText = strText & "  --  Chart Page: " & g.strChartPage
                strFile = g.ChartGlobals.strCPCRoot & "\Charts\SCP.INI"
                d = GetIniFileProperty("Published", 0, "", strFile)
                If d > 0 Then
                    s = GetIniFileProperty("PageName", "", "", strFile)
                    If UCase(s) <> UCase(g.strChartPage) Then
                        d = 0
                        KillFile strFile
                    End If
                End If
                If d > 0 Then
                    strText = strText & "  [published " & DateFormat(d, MM_DD_YYYY, H_MM, AP_LOWER) & " ET]"
                End If
            End If
        ElseIf g.ChartGlobals.bMyPageFeature Then
            strText = strText & "  --  Chart Page: <My Page>"
        ElseIf HasGold(False) Then
            strText = strText & "  --  Chart Page: (unnamed)"
        End If
        
        ' TLB 3/12/2013: resetting a toolbar picture with a toolbar dropdown displayed was causing issues,
        ' so the easy fix is that we only need to redo the picture when the feedtime is not passed in.
        frmMain.Icon = Picture16(strIcon)
        frmMain.tbToolbar.Tools("ID_About").Picture = frmMain.Icon
        
        strPrevCaption = strText
    End If
    
    frmMain.Caption = strText
    
ErrExit:
    Exit Sub
    
ErrSection:
    frmMain.tbToolbar.Redraw = True
    RaiseError "mMain.SetMainCaption", eGDRaiseError_Raise
    
End Sub

Public Function ChangeGridFont(fg As VSFlexGrid, Optional ByVal bResizeColumns As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the Redraw
    Dim lAutoSizeMode As Long           ' Current status of the Auto Size Mode

    If ProcessIsBusy Then Exit Function

    If CommonDialogFont(frmMain.CommonDialog1, fg.Font) Then
        With fg
            lRedraw = .Redraw
            lAutoSizeMode = .AutoSizeMode
            
            .Redraw = flexRDNone
            .Font = .Font '(this is required to trigger the grid to reset itself!)
            
            If bResizeColumns Then
                .AutoSizeMode = flexAutoSizeColWidth
                .AutoSize 0, .Cols - 1, , 75
            End If
            
            .AutoSizeMode = lAutoSizeMode
            .Redraw = lRedraw
        End With
        ChangeGridFont = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ChangeGridFont", eGDRaiseError_Raise
    
End Function

' TLB 11/7/2011: Check when just their "data package" will expire (i.e. the latest date for F, I, S)
Public Sub ExpiringDataPkgWarning(Optional ByVal bAsModalDialog As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, strMsg$
    Dim ModuleInfo As New cGdTable
    Dim lExpireDate As Long
    Dim lDataPkgExpire As Long
    Dim lNumDays As Long
    Dim lWarnDays As Long
    Dim frm As frmMessage
       
    ' get the table of module info
    If SU_GetPurchaseInfo(ModuleInfo.TableHandle) <> 0 Then
        ' only look for the primary data pkg records ("F", "S", "I")
        For i = 0 To ModuleInfo.NumRecords - 1
            Select Case Trim(UCase(ModuleInfo(kPurchaseCode, i)))
            Case "F", "S", "I"
                ' see which one has the latest expiration (furthest into the future)
                lExpireDate = JulFromLong(ModuleInfo.Num(kExpireDate, i))
                If lExpireDate > lDataPkgExpire Then
                    lDataPkgExpire = lExpireDate
                    ' *** NOTE: the following call doesn't really work for some reason ***
                    'If ModuleInfo(kPurchased, i) = ePurchased_Evaluation Then
                        'strTrial = " trial"
                    'End If
                End If
            End Select
        Next
    End If
    Set ModuleInfo = Nothing
        
    If lDataPkgExpire > 0 And g.lLCD > 0 Then
        ' see if # days until data pkg expires < # of days for Warning msg
        lNumDays = lDataPkgExpire - JulFromLong(RI_HonestDate) ' # days until expire
        lWarnDays = GetProvidedProperty("ExpireWarningDays", 0)
        If lNumDays <= lWarnDays Then
            If lNumDays <= 0 Then
                strMsg = "Your data subscription has expired.||To update your data,"
            Else
                strMsg = "Your data will stop updating in " & Str(lNumDays) & " days.||To avoid interruption,"
            End If
            strMsg = strMsg & " please call:|" & GetProvidedProperty("SalesContact", "800-808-3282", True)
            If bAsModalDialog Then
                ' at startup, we can bring up a modal dialog
                InfBox strMsg, "!", , "WARNING: Data Subscription Expiration", , , &HD8F8FF
            Else
                ' after a daily download, we'll bring up a non-modal but stay-on-top message window
                strMsg = vbCrLf & Replace(strMsg, "|", vbCrLf)
                Set frm = New frmMessage
                frm.rtbMessage.Font.Bold = True
                frm.ShowMe "WARNING: Data Subscription Expiration", strMsg, eStayOnTopMessage, True
                Set frm = Nothing
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ExpiringDataPkgWarning", eGDRaiseError_Raise
End Sub

Public Sub ExpiringModuleWarning()
On Error GoTo ErrSection:

    Dim ModuleInfo As New cGdTable
    Dim lIndex As Long
    Dim lExpireDate As Long
    Dim lRealDate As Long
    Dim strExpireMsg As String
    Dim strWarnMsg As String
    Dim lValue As Long
    Dim lNumDays As Long
    
 Exit Sub ' TLB 5/31/2011: not working correctly, so disabled until we can fix it
    
    lRealDate = JulFromLong(RI_HonestDate)  'Date
    
    If SU_GetPurchaseInfo(ModuleInfo.TableHandle) <> 0 Then
        If IsIDE Then
            strWarnMsg = ModuleInfo.ToString(vbCrLf, vbTab)
            FileFromString App.Path & "\ModuleInfo.txt", strWarnMsg, True
            strWarnMsg = ""
        End If
        For lIndex = 0 To ModuleInfo.NumRecords - 1
            If ModuleInfo(kPurchaseType, lIndex) = ePurchaseType_ModulePurchased Or _
                        ModuleInfo(kPurchaseType, lIndex) = ePurchaseType_ModuleLeased Then
                If ModuleInfo(kPurchased, lIndex) = ePurchased_Evaluation Then
                    lExpireDate = JulFromLong(ModuleInfo(kExpireDate, lIndex))
                    If lExpireDate <= lRealDate Then
                        lValue = GetIniFileProperty(ModuleInfo(kPurchaseCode, lIndex), 0, "Expire", g.strIniFile)
                        If lValue = 0 Then
                            strExpireMsg = strExpireMsg & ModDescription(ModuleInfo(kPurchaseCode, lIndex)) & "|"
                            SetIniFileProperty ModuleInfo(kPurchaseCode, lIndex), 1, "Expire", g.strIniFile
                        End If
                    ElseIf lExpireDate - lRealDate < 10 Then
                        strWarnMsg = strWarnMsg & ModDescription(ModuleInfo(kPurchaseCode, lIndex)) & "|" '& " in " & lRealDate - lExpireDate & " days|"
                        lNumDays = lExpireDate - lRealDate
                    End If
                End If
            End If
        Next lIndex
    End If
    
    If strWarnMsg <> "" Then
        InfBox "The following module(s) are set|to expire in " & lNumDays & " days:||" & _
            strWarnMsg & "|Please call Genesis Sales at|(800) 808-DATA to purchase these modules.", "!", , "Authorization Warning"
    End If
    
    If strExpireMsg <> "" Then
        InfBox "The following module(s) have expired:||" & strExpireMsg & _
            "|Please call Genesis Sales at|(800) 808-DATA to purchase these modules.", "!", , "Authorization Warning"
    End If
    
    Set ModuleInfo = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ExpiringModuleWarning", eGDRaiseError_Raise
    
End Sub

Public Function ModDescription(ByVal strModule As String) As String
On Error GoTo ErrSection:

    Dim strTemp As String
    
    strTemp = GetIniFileProperty(strModule, "", "Descriptions", AddSlash(App.Path) & "Info\Mods.INI")
    If strTemp = "" Then
        ModDescription = strModule
    Else
        ModDescription = strTemp
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ModDescription", eGDRaiseError_Raise
    
End Function

Public Function DialPager(ByVal nComPort&, _
        ByVal strPhone$, ByVal nWaitSeconds&, ByVal strMessage$, _
        ByVal strModemInit$, Optional ByVal bConfirm As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim i&, strTemp$, strSend$, dSleep#, dStart#

    ' set defaults
    If nComPort <= 0 Then nComPort = 1
    strModemInit = UCase(StripStr(strModemInit, " "))
    If Left(strModemInit, 2) <> "AT" Then
        strModemInit = "AT" & strModemInit
    End If
    strPhone = UCase(Trim(strPhone))
    strMessage = UCase(Trim(strMessage))
    If Len(strPhone) = 0 Or Len(strMessage) = 0 Then Exit Function
                
    If bConfirm Then
        strTemp = "Sending a page ..." _
            & "|To:  " & strPhone _
            & "|Msg:  " & strMessage
        If InfBox(strTemp, "!", "+Dial|-Abort", _
            "Trade Navigator Alert", , 5) = "A" Then
                Exit Function
        End If
    End If
                    
    Screen.MousePointer = 11
                    
    With frmMain.MSComm1

        ' try multiple times
        For i = 1 To 9
            ' Open com port
            .CommPort = nComPort
            ' 9600 baud, no parity, 8 data, and 1 stop bit.
            .Settings = "9600,N,8,1"
            ' Tell the control to read entire buffer when Input is used.
            .InputLen = 0
            ' Open the port.
            frmTest.AddList "Opening ComPort " & CStr(nComPort)
            .PortOpen = True
            ' Send the attention command to the modem.
            .Output = "AT" & Chr$(13)
            ' Wait for data to come back to the serial port.
            dStart = Now
            Do
                DoEvents
                If .InBufferCount >= 1 Then
                    Sleep 1 '(allow entire input to get buffered)
                    Exit Do
                End If
            Loop Until Now > dStart + 5 / 86400#
            ' Read the "OK" response data in the serial port.
            strTemp = .Input
            frmTest.AddList strTemp
            If InStr(strTemp, "OK") > 0 Then
                DialPager = True
                Exit For
            End If
    
            ' If error, close the serial port
            .PortOpen = False
            If i >= 2 Then  ' just try twice
                Beep
                Exit For
            End If
            Sleep 5 ' wait 5 seconds, then try again
        Next
        
        If DialPager Then
            ' String to send: Phone + Comma for each 2 seconds + Message
            strSend = StripStr(strPhone, " -()")
            For i = 1 To nWaitSeconds Step 2
                strSend = strSend & ","
            Next
            strSend = strSend & strMessage
    
            ' Dial
            .Output = strModemInit & "DT" & strSend & Chr$(13)
            dStart = Now
            Do
                DoEvents
                If .InBufferCount >= 1 Then
                    Sleep 1 '(allow entire input to get buffered)
                    Exit Do
                End If
            Loop Until Now > dStart + (nWaitSeconds + 1) / 86400#
            strTemp = .Input
            frmTest.AddList strTemp
            
            ' Wait until message should be finished
            dSleep = 5 ' (some extra buffer)
            For i = 1 To Len(strSend)
                Select Case Mid(strSend, i, 1)
                    Case "," ' add 2 seconds for each comma
                        dSleep = dSleep + 2
                    Case Else ' add some time for each digit
                        dSleep = dSleep + 0.75
                End Select
            Next
            Sleep dSleep
        
            ' Close the serial port.
            .PortOpen = False
            frmTest.AddList "ComPort closed"
        End If
    End With
    
ErrExit:
    Screen.MousePointer = 0
    Exit Function
    
ErrSection:
    DialPager = False
    Resume Next '(need to let it continue in order to close the port, etc.)

End Function

#If 0 Then
Public Function HasGrayedChartMenu(frm As Form) As Boolean

    Dim hMenu&, uFlags&
    
    hMenu = GetSystemMenu(frm.hWnd, 0)
    If hMenu <> 0 Then
        uFlags = GetMenuState(hMenu, SC_MINIMIZE, 0)
        If (uFlags And MF_DISABLED) Or (uFlags And MF_GRAYED) Then
            HasGrayedChartMenu = True
        End If
    End If

End Function
#End If

Public Function GetUnusedChartName() As String
On Error GoTo ErrSection:

    Dim i&, strNew$

    ' find first unused #
    For i = 1 To 99999
        strNew = "Cus" & Format(i, "00000")
        If Not FileExist(g.ChartGlobals.strCPCRoot & "\Charts\" & strNew & ".CHT") Then Exit For
    Next
    GetUnusedChartName = strNew

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.GetUnusedChartName"
End Function

Public Sub LoadChartPage(ByVal strPage$, Optional ByVal bSkipSave As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, strPagePath$, strText$, dStarted#, dEnded#
    Dim aPages As cGdArray
    
    Dim aReuseableForms As cGdArray
    Dim bLocked As Boolean
    
    Dim bFileRestore As Boolean
    Dim bSaveChanges As Boolean

    strPagePath = g.ChartGlobals.strCPCRoot & "\Charts\Pages\"
    
    If strPage = "+" Or strPage = "-" Then
        ' find the current page
        Set aPages = GetAllowedList("P")
        For i = 0 To aPages.Size - 1
            If UCase(Parse(aPages(i), vbTab, 1)) = UCase(g.strChartPage) Then
                Exit For
            End If
        Next
        ' want to load the next or previous page
        If strPage = "+" Then
            i = i + 1
            If i >= aPages.Size Then i = 0
        Else
            i = i - 1
            If i < 0 Then i = aPages.Size - 1
        End If
        strPage = Parse(aPages(i), vbTab, 1)
    End If
    
    If g.ChartGlobals.bMyPageFeature Then
        bSkipSave = True
        If Len(g.strChartPage) = 0 Then
            'this is (unnamed) chart page that shows up as <My Page> in menu item & tool button
            SaveChartPage "My Page"
            'if both the current chart page is blank and the page to load is blank then user
            'have simply clicked the <My Page> option while currently on <My Page> so just exit
            'else the result will be the equivalent of <create new chart page>
            If Len(strPage) = 0 Then GoTo ErrExit
        ElseIf Len(strPage) = 0 And FileExist(g.strAppPath & "\charts\pages\My Page.gzp") Then
            strPage = "My Page"
        End If
    End If
    
    If bSkipSave Then
        bSaveChanges = False
    Else
        bSaveChanges = True
        ' see if current page is "dirty" (has changed since last loaded)
        If (g.bDirtyChartPage Or Len(g.strChartPage) = 0) And Not ActiveChart Is Nothing Then
            strText = ""
            If Len(g.strChartPage) = 0 Then
                strText = "Do you want to save the current page?"
                strText = InfBox(strText, "?", "+Save|No|-Cancel", "Chart Page")
            ElseIf UCase(strPage) = UCase(g.strChartPage) Then
                strText = "Revert to when this page was previously| saved?  (current changes will be lost)"
                strText = InfBox(strText, "?", "Revert|+-Cancel", "Chart Page")
            ElseIf GetIniFileProperty("AutoSavePage", 0, "Charting", g.strIniFile) = 0 Then
                strText = "Do you want to save the current page:|" & g.strChartPage
                strText = InfBox(strText, "?", "+Save|No|-Cancel", "Chart Page")
            Else ' auto-save
                strText = "S"
            End If
            
            If strText = "C" Then
                Exit Sub
            ElseIf strText = "S" Then
                If Not SaveChartPage(g.strChartPage) Then
                    Exit Sub
                End If
            ElseIf strText = "R" Then
                bFileRestore = True         'force restore from files
                bSaveChanges = False        '6401
            ElseIf strText = "N" Then
                bSaveChanges = False        'user said No to prompt for saving changes to page
                
                If Not g.ChartPageCache Is Nothing Then
                    g.ChartPageCache.Remove g.strChartPage
                End If
            End If
        End If
    End If
    
gdResetProfiles 800, 899
gdStartProfile 800
gdStartProfile 801
    dStarted = gdTickCount
   
' testing: if want to delete all charts first
Dim frm As Form
If 0 Then
    Set frm = Nothing
    For i = Forms.Count - 1 To 0 Step -1
        If IsFrmChart(Forms(i)) Then
            If Forms(i).WindowState = vbMaximized Then
                Set frm = Forms(i)
            Else
                Unload Forms(i)
            End If
        End If
    Next
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    DoEvents
End If
    
    ' delete files for current chart page
    KillFile g.ChartGlobals.strCPCRoot & "\Charts\* /i=CHT,Charts.cfg,INI,Cus*.ano,SCP.cfg" ' /x=^*.ano /x=Replay^*.ano"
    
    ' unzip files for the new chart page
'    LockWindowUpdate frmMain.hWnd
    If Len(strPage) > 0 Then
        InfBox strPage, "t", , "Loading chart page ...", True
        MoveFocus ActiveChart '(so focus won't flicker to/from InfBox)
        ZipExecute "U", strPagePath & strPage & ".GZP", g.ChartGlobals.strCPCRoot & "\Charts\"
    End If
    
    If g.bLoadPageOldMethod Then
        bFileRestore = True
    ElseIf bSaveChanges Then
        If g.ChartGlobals.bMyPageFeature Then
            If Not g.bDirtyChartPage Or g.strChartPage = "My Page" Then
                'only cache page if it is the user's page or if no changes have been made
                Set aReuseableForms = CachePageSave(g.strChartPage)
            End If
        Else
            'save chart objects to cache & get returned array of reuseable frmChart
            Set aReuseableForms = CachePageSave(g.strChartPage)
        End If
    End If
    
    g.strChartPage = strPage
    

'JM 07-27-2011 - code moved here from RestoreCharts
ChartTimers = False
g.bLoadingChartPage = True
g.bSkipSetChartFocus = True
g.ChartGlobals.nDetached = 0
Set m.ActiveChartForm = Nothing
Set g.ChartGlobals.frmActiveNonDetached = Nothing
bLocked = LockWindowUpdate(frmMain.hWnd)

    
    If Not bFileRestore Then
        bFileRestore = True         'assume need to load from files
gdStopProfile 801
gdStartProfile 802
        ' load charts for new chart page from cache
        If Not g.ChartPageCache Is Nothing Then
            If Not g.ChartPageCache(g.strChartPage) Is Nothing Then
                If CachePageRestore(aReuseableForms, g.strChartPage) Then
                    bFileRestore = False
                    frmMain.tmrWindowLink.Enabled = True        '6433
                End If
            End If
        End If
gdStopProfile 802
gdStartProfile 803
    End If
    
    ' load charts for new chart page from files
    If bFileRestore Then
gdStopProfile 801
gdStartProfile 802
        RestoreCharts False, aReuseableForms
gdStopProfile 802
gdStartProfile 803
    End If
    
    g.strChartPage = strPage
    g.bDirtyChartPage = False
    
    If Not ActiveChart Is Nothing Then
        ActiveChart.SetChartTabs
    End If
    
    SetMainCaption
    DoEvents
    
    If Not ActiveChart Is Nothing Then
        MoveFocus ActiveChart.pbChart
    End If
    
    ChartTimers = True
    dEnded = gdTickCount
gdStopProfile 803
gdStopProfile 800
    
    If g.ChartGlobals.bMyPageFeature And g.strChartPage = "My Page" Then
        g.strChartPage = ""
    End If
       
    If IsIDE Or g.bLoadPageTime Then
        StatusMsg "Page loaded: " & Format((dEnded - dStarted) / 1000#, "0.00") & " seconds"
        'InfBox gdGetProfiles(800, 830, "|"), "i", , "LoadChartPage"
    End If

    DebugLog g.RealTime.DumpTickBufferInfo

ErrExit:
    g.bLoadingChartPage = False
    g.bSkipSetChartFocus = False
    If bLocked Then LockWindowUpdate 0
    If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        FormResize g.ChartGlobals.frmActiveNonDetached
    End If
    Exit Sub
    
ErrSection:
    g.bLoadingChartPage = False
    g.bSkipSetChartFocus = False
    If bLocked Then LockWindowUpdate 0
    
    RaiseError "mMain.LoadChartPage", eGDRaiseError_Raise

End Sub

Private Sub RestoreCharts(Optional ByVal bStartup As Boolean = True, _
    Optional ByRef aReuseForms As cGdArray = Nothing)
On Error GoTo ErrSection:

    Dim i&, strTemplate$, ws%, iForm&, strText$, iPos&, strOld$
    Dim aUnique As New cGdArray
    Dim bActiveFound As Boolean
    Dim bRatiosFound As Boolean
       
    Dim frm As Form                     'working variable used in FOR loop
    Dim frmActive As frmChart           'non-detached form to set as active
    Dim frmLast As frmChart             'last form processed in FOR loop
        
    Dim aCharts As New cGdArray
    Dim aFlds As New cGdArray
    Dim strActive As String
    
    g.bPageHasEWILabels = False

gdStartProfile 811

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JM 07-27-2011 - this code moved to calling routine
'
'ChartTimers = False
'g.bLoadingChartPage = True
'g.bSkipSetChartFocus = True

    'reset temporary detached chart count
'    g.ChartGlobals.nDetached = 0
    
'    Set m.ActiveChartForm = Nothing
'    Set g.ChartGlobals.frmActiveNonDetached = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If aReuseForms Is Nothing Then
        iForm = Forms.Count - 1
        For i = iForm To 0 Step -1
            If TypeOf Forms(i) Is frmChart2 Then
                Unload Forms(i)         'JM: cannot reuse detached charts
            ElseIf TypeOf Forms(i) Is frmChart Then
                If Forms(i).Tag = "BOGUS" Or Forms(i).WindowState = vbMinimized Then
                    Unload Forms(i)     '6659
                ElseIf Forms(i).WindowState = vbMaximized Then
                    Forms(i).WindowState = vbNormal ' TEST THIS
                End If
            End If
        Next
    End If
    
'JM 05-19-2009: From Tim - requires gold
'JM 07-13-2013: For Glen - everyone gets detached charts
'    bAllowDetach = HasGold(False)         'FileExist("DetachCharts.flg")

gdStopProfile 811
gdStartProfile 812

    ' load charts config from file
    aCharts.FromFile g.ChartGlobals.strCPCRoot & "\Charts\Charts.cfg"
    For i = aCharts.Size - 1 To 1 Step -1
        aFlds.SplitFields aCharts(i), vbTab
        
        ' replace existing .CHT files containing older Woodies Templates
        strOld = g.ChartGlobals.strCPCRoot & "\Charts\" & aFlds(1) & ".CHT"
        If FileExist(strOld) Then
            strTemplate = GetIniFileProperty("TemplateApplied", "", "General", strOld)
            If InStr(UCase(strTemplate), "WOODIES CCI") <> 0 Then
                ' 1/4/2008: replace obsolete "Club Patterns" templates with the newer one
                If UCase(strTemplate) = "WOODIES CCI CLUB PATTERNS" Or UCase(strTemplate) = "WOODIES CCI PATTERNS" Then
                    strTemplate = g.ChartGlobals.strCPCRoot & "\Charts\Templates\WCCI Basic Room.CHT"

                ' replace older CCI templates with newer one (if DateApplied does not yet exist)
                ElseIf GetIniFileProperty("DateApplied", "-1", "General", strOld) = "-1" Then
                    strTemplate = g.ChartGlobals.strCPCRoot & "\Charts\Templates\" & strTemplate & ".CHT"
                Else
                    strTemplate = ""
                End If
                If Len(strTemplate) > 0 Then
                    If FileExist(strTemplate) Then
                        FileCopy strTemplate, strOld, True
                        SetIniFileProperty "DateApplied", Now, "General", strOld
                    End If
                End If
            End If
        End If
        
        ' move active chart to end of the array
        ' (since focus doesn't always go there unless it's the last chart restored)
        If Not bActiveFound Then
            If i >= 1 And Val(aFlds(4)) = 1 Then
                If Len(aFlds(0)) > 0 Then
                    strActive = aCharts(i)      'save to string
                    aCharts.Remove i            'remove from array
                    bActiveFound = True
                End If
            End If
        End If
    Next
    
    If bActiveFound And Len(strActive) > 0 Then aCharts.Add strActive
                   
gdStopProfile 812
gdStartProfile 813
                   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'aFlds[0]   symbol ID
'aFlds[1]   template file (cusxxxx.cht)
'aFlds[2]   window normal position (left,top,width,height,window state,visible)
'aFlds[3]   window state (normal=0, min=1, max=2)
'aFlds[4]   active flag
'aFlds[5]   symbol link color
'aFlds[6]   period link color
'aFlds[7]   window ratio position (left,top,width,height,window state,visible)
'aFlds[8]   detached flag: 3=detached (see enumDetachStatus)
'aFlds[9]   window detached position (left,top,width,height,window state,visible)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' it's faster to re-use existing chart forms
    iForm = Forms.Count
    
    ' restore each chart
    g.ChartGlobals.nDetached = 0
    For i = 1 To aCharts.Size - 1
        aFlds.SplitFields aCharts(i), vbTab
        strTemplate = aFlds(1)
        If Len(aFlds(0)) <> 0 And Len(strTemplate) <> 0 And FileExist(g.ChartGlobals.strCPCRoot & "\Charts\" & strTemplate & ".CHT") Then
            ' TLB: fix bug where duplicates could occur
            If aUnique.BinarySearch(strTemplate, iPos, eGdSort_IgnoreCase) Or UCase(strTemplate) = "TEMP" Then
                strOld = strTemplate
                strTemplate = GetUnusedChartName
                FileCopy g.ChartGlobals.strCPCRoot & "\Charts\" & strOld & ".CHT", g.ChartGlobals.strCPCRoot & "\Charts\" & strTemplate & ".CHT"
                aUnique.BinarySearch strTemplate, iPos, eGdSort_IgnoreCase
            End If
            aUnique.Add strTemplate, iPos
        
            ' see if can re-use an existing chart form
            Set frm = Nothing
            Do While iForm > 0
                iForm = iForm - 1
                If TypeOf Forms(iForm) Is frmChart Then
                    ' clear out old template and form
                    Set frm = Forms(iForm)
                    If frm.WindowState <> 0 Then frm.WindowState = 0
                    frm.Chart.ClearChartForReuse
                    frm.Tag = ""
                    frm.tmr.Tag = ""
                    Exit Do
                End If
            Loop
            If frm Is Nothing Then
                ' else create a new one
                Set frm = New frmChart          'new chart is always non-detached
            End If
            
            ' load new template and symbol
            If Not frm.Chart.TemplateLoad(strTemplate) Then
                Unload frm
            Else
                If InStr(aFlds(0), "|") > 0 Or InStr(aFlds(0), ";") > 0 Then
                    frm.Chart.SetSymbol aFlds(0), , False
                Else
                    frm.Chart.SetSymbol Val(aFlds(0)), , False
                End If
                frm.WindowLink.SymbolColor = Val(aFlds(5))
                frm.WindowLink.PeriodColor = Val(aFlds(6))
                frm.Chart.ResetLastScreenDate
                If InStr(aFlds(7), ";") > 0 And InStr(aFlds(7), ",") = 0 Then
                    bRatiosFound = True ' (flag for backwards-compatibility)
                Else
                    aFlds(7) = ""
                End If
                
                If aFlds.Size > 8 Then
                    If aFlds(8) = eDetached Then
                        frm.DetachStatus = eDetached
                    Else
                        frm.DetachStatus = eNotDetached
                    End If
                    If aFlds.Size > 9 Then
                        frm.CopyPlacements , aFlds(2), aFlds(7), aFlds(9)
                    Else
                        frm.CopyPlacements , aFlds(2), aFlds(7)
                    End If
                Else
                    frm.CopyPlacements , aFlds(2), aFlds(7)
                End If
                
                ' always set the position as ratio of MDI client (even if blank, to clear it)
                frm.SetRatioPlacement aFlds(7)
                ' then set position as # twips only if during startup (since it
                ' looks smoother) and for older chart pages (with no ratios)
                If frm.DetachStatus = eNotDetached Then
                    If bStartup Or InStr(aFlds(7), ";") = 0 Then
                        SetFormPlacement frm, aFlds(2), "P"
                    End If
                End If
                ws = Val(aFlds(3))
                
                If frm.DetachStatus = eDetached Then
                    frm.tmr.Tag = "DETACH_NOW"
                    'a new window is always shown as windowstate normal
                    'save the state now so we know to maximize it later if necessary
                    frm.Tag = aFlds(3)
                    g.ChartGlobals.nDetached = g.ChartGlobals.nDetached + 1
                Else
                    frm.Show
                    frm.ZOrder
                    'let the chart's timer miminize the form after all data members/objects are fully set/initialized
                    'this is fix for grey & out-of-sync charts menu drop down reported by Vanessa
                    If ws = 1 Then
                        frm.tmr.Tag = "MINIMIZE_NOW"
                    Else
                        Set frmActive = frm
                    End If
                    Set frmActive = frm
                End If
                ' TLB 1/7/2009: looks like we CANNOT call generate chart here (while reloading a chart page)
                ' -- it's causing something to run "too early" and causes crashes esp. with spread charts
                ' (and the fix below ended up not needing to be done right here anyway)
                ''frm.Chart.GenerateChart eRedo1_Scrolled ' 10/29/2008 needed to fix issues with countdown panel
                If Val(aFlds(4)) <> 0 Then
                    If frm.DetachStatus = eNotDetached Then Set frmActive = frm
                End If
                Set frmLast = frm
            End If
            Set frm = Nothing
        End If
                
        If bStartup Then
            frmSplash.Message 75 + (20# * i) / (aCharts.Size - 1)
        End If
    Next
        
    If FileExist(App.Path & "\ewave.flg") Or FileExist(App.Path & "\gmp.flg") Then      '6926
        If frmMain.tbToolbar.Tools("ID_ShowEWI").Visible Then
            frmMain.tbToolbar.Tools("ID_ShowEWI").Visible = False
            ToolbarReset False
            ToolbarInit2 frmMain, frmMain.TbButtonsArray(kTbDraw), , kTbDraw, , g.vbeTbAlignDraw
            ToolbarResize2 frmMain, frmMain.pbTbBackDraw, frmMain.imgTbBackDraw, frmMain.TbButtonsArray(kTbDraw), frmMain.ToolBarWrapGet(kTbDraw)
        End If
    Else
        If frmMain.tbToolbar.Tools("ID_ShowEWI").Visible <> g.bPageHasEWILabels Then
            frmMain.tbToolbar.Tools("ID_ShowEWI").Visible = g.bPageHasEWILabels
            ToolbarReset False
            ToolbarInit2 frmMain, frmMain.TbButtonsArray(kTbDraw), , kTbDraw, , g.vbeTbAlignDraw
            ToolbarResize2 frmMain, frmMain.pbTbBackDraw, frmMain.imgTbBackDraw, frmMain.TbButtonsArray(kTbDraw), frmMain.ToolBarWrapGet(kTbDraw)
        End If
    End If

gdStopProfile 813
gdStartProfile 814
    
    KillFile g.ChartGlobals.strCPCRoot & "\Charts\Temp.cht"
    
    Do While iForm > 0
        iForm = iForm - 1
        If IsFrmChart(Forms(iForm)) Then
            Unload Forms(iForm)
        End If
    Loop
    
    If NonDetachCount = 0 And Not frmLast Is Nothing Then
        Set frmActive = frmLast
        frmLast.DetachStatus = eNotDetached
        frmLast.tmr.Tag = ""
        frmLast.Show
    End If
    
    'aCharts.Size = 0 when:
    '   1. creating new chart pages from pages dropdown menu because LoadChartPage clears files
    '   2. user closed all charts right before quitting TradeNavigator
    If aCharts.Size = 0 Then
        aCharts.Add "MAX" 'default is to maximize the chart
        Set frmActive = New frmChart        'new chart is always non-detached
        frmActive.Chart.SetSymbol g.SymbolPool.SymbolIDforSymbol("$DJIA")
        frmActive.Show
    End If
        
    If Not frmActive Is Nothing Then Set g.ChartGlobals.frmActiveNonDetached = frmActive
        
    strText = aCharts(0)
    g.strChartPage = Parse(strText, vbTab, 2)
'    g.bLoadingChartPage = False
    If Not frmActive Is Nothing Then
        frmMain.SetWindowLink frmActive
        ''MoveFocus frmActive ' TEST THIS OUT
        If Left(strText, 3) = "MAX" Then
            frmActive.WindowState = 2
        ElseIf frmActive.WindowState <> 0 Then
            frmActive.WindowState = 0
        End If
        Set frmActive = Nothing
    End If
    
    i = Val(Parse(strText, vbTab, 3))
    If i <= 1 Then i = GetIniFileProperty(frmSymbolGrid.Name, 1, "LinkDefaults", g.strIniFile)
    frmSymbolGrid.WindowLink.SymbolColor = i
    If i = 1 And Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        frmSymbolGrid.SymbolID = g.ChartGlobals.frmActiveNonDetached.SymbolID
    End If
    
    i = Val(Parse(strText, vbTab, 4))
    If i <= 1 Then i = GetIniFileProperty(frmSnapshot.Name, 1, "LinkDefaults", g.strIniFile)
    frmSnapshot.WindowLink.SymbolColor = i
    If i = 1 And Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        frmSnapshot.SymbolID = g.ChartGlobals.frmActiveNonDetached.SymbolID
    End If
        
gdStopProfile 814
gdStartProfile 815
    
g.bLoadingChartPage = False ' TESTING THIS HERE
    
    If bStartup Then
        UpdateVisibleCharts -1
    ElseIf Left(strText, 3) = "MAX" Then
        ' looks a little faster to show the maximized chart now,
        ' then update the background charts later
        UpdateVisibleCharts -1
        InfBox
    Else
        ' need to update all charts before showing
        UpdateVisibleCharts -1
        ' TLB 4/17/2008: AutoTile is now ONLY used when loading an older
        ' chart page that does not have the MDI client ratios assigned
        ' (i.e. for backwards-compatibility of pages intended to be tiled)
        If Not bStartup And Not bRatiosFound Then
            frmArrange.ArrangeCharts True
        End If
        InfBox
    End If
        
gdStopProfile 815
        
ErrExit:
    Set aUnique = Nothing
    Set frm = Nothing
    Set frmActive = Nothing
    Set frmLast = Nothing
'    g.bSkipSetChartFocus = False           'moved to calling routine
'    LockWindowUpdate 0
    Exit Sub
    
ErrSection:
'    g.bLoadingChartPage = False            'moved to calling routine
'    g.bSkipSetChartFocus = False
'    LockWindowUpdate 0
    
    RaiseError "mMain.RestoreCharts", eGDRaiseError_Raise
    
End Sub

Public Function SaveChartPage(Optional ByVal strPage$ = "", Optional ByVal strPublishZip As String = "") As Boolean
On Error GoTo ErrSection:

    Dim i&, strPath$, strFile$
    Dim aChartFiles As New cGdArray
    
    If Len(strPublishZip) > 0 Then
        KillFile strPublishZip, True
    End If

    ' get name of page
    strPage = Trim(strPage)
    If Len(strPage) = 0 Then
        strPage = g.strChartPage
        Do
            ' ask for name until valid or cancelled
            If Len(strPublishZip) > 0 Then
                strPage = Trim(InfBox("Publish this as a shared chart page:", "?", _
                    "+Publish|-Cancel", "Publish Chart Page", , , , , , "s", strPage))
            Else
                strPage = Trim(InfBox("Name for this chart page:", "?", _
                    "+Save|-Cancel", "Save Chart Page", , , , , , "s", strPage))
            End If
            If Len(strPage) = 0 Then Exit Function
            If IsValidFileBase(strPage) Then Exit Do
        Loop
    End If
    g.strChartPage = strPage
    
    ' save the current charts
    LockWindowUpdate frmMain.hWnd
    SaveCharts
    LockWindowUpdate 0
    
    ' zip up the files needed for just this chart page
    ' (but not global annotation files, which start with a caret)
    strPath = g.ChartGlobals.strCPCRoot & "\Charts\"
    If Len(strPublishZip) > 0 Then
        'FileFromString strPath & "SCP.cfg", Str(ConvertTimeZone(Now, "", "NY")) & vbTab & strPage
        SetIniFileProperty "PageName", strPage, "", strPath & "SCP.INI"
        SetIniFileProperty "Published", Str(ConvertTimeZone(Now, "", "NY")), "", strPath & "SCP.INI"
        
        ' TLB 12/12/2014: but ignore the "unpublishable" charts for the Publishing zip file
        KillFile strPath & "*.UNP", True
        aChartFiles.GetMatchingFiles strPath & "*.CHT"
        For i = aChartFiles.Size - 1 To 0 Step -1
            strFile = aChartFiles(i)
            If GetIniFileProperty("Unpublishable", 0, "", strFile) <> 0 Then
                RenameFile strFile, ReplaceFileExt(strFile, ".UNP")
            Else
                aChartFiles.Remove i
            End If
        Next
        ZipExecute "C", strPublishZip, strPath, "* /i=CHT,Charts.cfg,INI,Cus*.ano" ' /x=^*.ano /x=Replay^*.ano"
        For i = 0 To aChartFiles.Size - 1
            ' now restore the .CHT extensions
            strFile = aChartFiles(i)
            RenameFile ReplaceFileExt(strFile, ".UNP"), strFile
        Next
    End If
    ZipExecute "C", strPath & "Pages\" & strPage & ".GZP", strPath, "* /i=CHT,Charts.cfg,INI,Cus*.ano" ' /x=^*.ano /x=Replay^*.ano"
    
    g.bDirtyChartPage = False
    SetMainCaption
    SaveChartPage = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.SaveChartPage", eGDRaiseError_Raise
End Function


Public Sub SaveCharts()
On Error GoTo ErrSection:

    Dim i&, hActive&, strFile$, s$, ws%
    Dim aCharts As New cGdArray
    Dim frm As Form
    Dim Charts As cGdTree

    If g.SymbolPool.NumRecords = 0 Then Exit Sub
    
    'unload any chart that is in gamemode so we don't save it to the Cfg file
    'use window state of non-detached chart for saving to first line in Cfg file
    If frmMain.ActiveForm Is Nothing Then
        Set frm = ActiveChart       'this happens when user closes all non-detached charts then immediately quits TN
    ElseIf TypeOf frmMain.ActiveForm Is frmChart2 Then
        Set frm = frmMain.ActiveForm    'a little fast since this will always be non-detached
    Else
        Set frm = ActiveChart
    End If
    
    If Not frm Is Nothing Then
        If frm.IsInGameMode Then
            Unload frm
            Set frm = Nothing
        ElseIf frm.DetachStatus <> eNotDetached Then
            Set frm = Nothing           'theoretically should never get here
        End If
    End If
    
    If frm Is Nothing Then
        For i = Forms.Count - 1 To 0 Step -1
            Set frm = Nothing
            If IsFrmChart(Forms(i)) Then
                Set frm = Forms(i)
                If frm.IsInGameMode Then
                    Unload frm
                ElseIf frm.DetachStatus = eNotDetached Then
                    Exit For
                End If
            End If
        Next
    End If
    
    If frm Is Nothing Then
        'frm will be nothing if all charts are detached
        hActive = 0
        s = "NORMAL"
    Else
        hActive = frm.hWnd
        If frm.WindowState = vbMaximized Then
            s = "MAX" ' Maximized charts
        Else
            s = "NORMAL"
        End If
    End If
    
    aCharts.Add s & vbTab & g.strChartPage & vbTab & Str(frmSymbolGrid.WindowLink.SymbolColor) _
                    & vbTab & Str(frmSnapshot.WindowLink.SymbolColor)
    
#If 0 Then
    Dim tbForms As New cGdTable
    Dim aIndex As cGdArray
    
    'create table's fields
    tbForms.CreateField eGDARRAY_Longs, 0, "FormLeft"
    tbForms.CreateField eGDARRAY_Longs, 1, "FormIndex"
    'save forms to table then sort by form's left
    For i = 0 To Forms.Count - 1
        If IsFrmChart(Forms(i)) Then
            Set frm = Forms(i)
            If Not frm.IsInGameMode Then
                If Len(frm.Chart.Symbol) > 0 Then
                    tbForms.AddRecord ""
                    tbForms(0, tbForms.NumRecords - 1) = frm.Left
                    tbForms(1, tbForms.NumRecords - 1) = i
                End If
            End If
        End If
    Next
    Set aIndex = tbForms.CreateSortedIndex(0)
    
    For i = 0 To aIndex.Size - 1
        Set frm = Forms(tbForms(1, aIndex(i)))
        If IsFrmChart(frm) Then
            If Not frm.IsInGameMode Then
                ws = frm.WindowState
                If g.bUnloading Then
                    frm.Visible = False
                End If
                frm.Chart.TemplateSave
                If frm.Chart.SymbolID = 0 Then
                    s = frm.Chart.ExternalData
                    If Len(s) = 0 Then s = frm.Chart.SpreadSymbols
                Else
                    s = CStr(frm.Chart.SymbolID)
                End If
                s = s & vbTab & frm.Chart.Template _
                    & vbTab & frm.GetNormalPlacement & vbTab & CStr(ws) & vbTab
                If frm.hWnd = hActive Then
                    s = s & "1"
                Else
                    s = s & "0"
                End If
                s = s & vbTab & Str(frm.WindowLink.SymbolColor) & vbTab _
                    & Str(frm.WindowLink.PeriodColor) & vbTab & frm.GetRatioPlacement & vbTab & Str(frm.DetachStatus)
                
                If Len(frm.GetDetachedPlacement) > 0 Then
                    s = s & vbTab & frm.GetDetachedPlacement
                End If
                
                aCharts.Add s
            End If
        End If
    Next
#Else

    ' for each chart (attached and detached), in reverse Zorder
    Set Charts = GetChartsInZorder
    For i = Charts.Count To 1 Step -1
        Set frm = Charts(i)
        If Not frm.IsInGameMode Then
            If Len(frm.Chart.Symbol) > 0 Then
                ws = frm.WindowState
                If g.bUnloading Then
                    frm.Visible = False
                End If
                frm.Chart.TemplateSave
                If frm.Chart.SymbolID = 0 Then
                    s = frm.Chart.ExternalData
                    If Len(s) = 0 Then s = frm.Chart.SpreadSymbols
                Else
                    s = CStr(frm.Chart.SymbolID)
                End If
                s = s & vbTab & frm.Chart.Template _
                    & vbTab & frm.GetNormalPlacement & vbTab & CStr(ws) & vbTab
                If frm.hWnd = hActive Then
                    s = s & "1"
                Else
                    s = s & "0"
                End If
                s = s & vbTab & Str(frm.WindowLink.SymbolColor) & vbTab _
                    & Str(frm.WindowLink.PeriodColor) & vbTab & frm.GetRatioPlacement & vbTab & Str(frm.DetachStatus)
                
                If Len(frm.GetDetachedPlacement) > 0 Then
                    s = s & vbTab & frm.GetDetachedPlacement
                End If
                
                aCharts.Add s
            End If
        End If
    Next
    Set Charts = Nothing
#End If

    strFile = g.ChartGlobals.strCPCRoot & "\Charts\Charts.cfg"
    If aCharts.Size > 1 Then            'first line is not chart's symbol & template info
        aCharts.ToFile strFile
    Else
        KillFile strFile
    End If
    
    ' kill off old charts (older than 10 minutes ago)
    aCharts.GetMatchingFiles g.ChartGlobals.strCPCRoot & "\Charts\*.CHT /o=-10m"
    For i = 0 To aCharts.Size - 1
        strFile = aCharts(i)
        KillFile strFile
        ' and any annotation files that were with it
        KillFile Left(strFile, Len(strFile) - 4) & "^*.ANO"
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.SaveCharts", eGDRaiseError_Raise
End Sub

' Call this to get a collection of all the chart forms in zOrder:
' - the topmost MDI/attached chart is in the #1 spot
' - followed by all the other MDI/attached charts in zOrder
' - then followed by all the detached charts (not necessarily in any order)
Public Function GetChartsInZorder() As cGdTree
On Error GoTo ErrSection:

    Dim i&, iZorder&, hWnd&, strKey$
    Dim frm As Form
    Dim Charts As New cGdTree

    ' get all charts (attached and detached)
    Charts.AllowMultipleObjectTypes = True ' (since both frmChart and frmChart2)
    For i = Forms.Count - 1 To 0 Step -1
        Set frm = Forms(i)
        If IsFrmChart(frm) Then
            ' but just throw away any bogus charts
            If frm.Tag = "BOGUS" Then
                Set frm = Nothing
                Unload Forms(i)
            Else
                Charts.Add frm, Str(frm.hWnd)
            End If
        End If
    Next
    Set frm = Nothing

    ' get Zorder of all the MdiChild forms
    ' (first MdiChild is actually the "grandchild" of frmMain)
    iZorder = 0
    hWnd = GetWindow(frmMain.hWnd, GW_CHILD)
    hWnd = GetWindow(hWnd, GW_CHILD)
    Do While hWnd <> 0
        strKey = Str(hWnd)
        Set frm = Charts(strKey)
        If Not frm Is Nothing Then
            ' re-insert this chart into it's proper Zorder
            iZorder = iZorder + 1
            i = Charts.Index(strKey)
            If i <> iZorder Then
                Charts.Remove strKey
                Charts.Add frm, strKey, iZorder, eTREE_Myself
                'i = Charts.Index(strKey)
                'If i <> iZorder Then
                '    i = i
                'End If
            End If
        End If
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop
    
ErrExit:
    Set GetChartsInZorder = Charts
    Exit Function
    
ErrSection:
    RaiseError "mMain.GetChartsInZorder"
End Function

' This needs to be called from the "KeyDown" event of FlexGrid's (esp. on docked
' forms), since the form does not end up previewing this keystroke from the grid
Function fgKeyDown(KeyCode As Integer, Shift As Integer) As Boolean

    On Error Resume Next
    Dim s$
    If Shift = 2 Then
        If KeyCode = vbKeyPageUp Then
            s = "+"
        ElseIf KeyCode = vbKeyPageDown Then
            s = "-"
        End If
        If Len(s) > 0 Then
            KeyCode = 0
            fgKeyDown = True
            LoadChartPage s
        End If
    End If

End Function

Private Sub CheckNewButtons(aShow As cGdArray)
On Error GoTo ErrSection:

    Dim iVer&
    
    If aShow Is Nothing Then Exit Sub
    If aShow.Size <= 0 Then Exit Sub
    
    ' use ToolbarVersion to add new buttons that should be shown by default for existing users
    iVer = Val(Parse(aShow(0), vbTab, 1))
    If iVer = 0 Then
        ' for older file: get version from INI and insert
        iVer = GetIniFileProperty("ToolbarVersion", 0, "Toolbars", g.strIniFile)
        aShow.Add Str(iVer), 0
    End If
    
    frmToolbar.iVersion = 42 '(store current toolbar version)
        
    If iVer < frmToolbar.iVersion Then
        aShow(0) = Str(frmToolbar.iVersion)
        If iVer < 3 Then
            aShow.Add "ID_OHLCBars"
            aShow.Add "ID_Candlesticks"
            aShow.Add "ID_CloseLine"
            aShow.Add "ID_AndrewFork"
        End If
        If iVer < 4 Then
            aShow.Add "ID_ChartMove"
            aShow.Add "ID_Eraser"
        End If
        If iVer < 5 Then
            aShow.Add "ID_Replay"
            aShow.Add "ID_PlanetData"
        End If
        If iVer < 6 Then
            aShow.Add "ID_SRLine"
        End If
        If iVer < 7 Then
            aShow.Add "ID_ChartOnOff"
        End If
        If iVer < 8 Then
            aShow.Add "ID_ConditionBuilder"
        End If
        If iVer < 9 Then
            aShow.Add "ID_FibFan"
            aShow.Add "ID_FibExpansion"
        End If
        If iVer < 10 Then
            aShow.Add "ID_SpResistFan"
        End If
        If iVer < 12 Then
            aShow.Add "ID_TradeTracker"
        End If
        If iVer < 13 Then
            aShow.Add "ID_Mirror"
            aShow.Add "ID_Pattern"
            aShow.Add "ID_RiskReward"
            aShow.Add "ID_Triangle"
            aShow.Add "ID_ChannelHighlight"
            aShow.Add "ID_WaveLabels"
        End If
        If iVer < 14 Then
            aShow.Add "ID_Magnet"
        End If
        If iVer < 15 Then
            aShow.Add "ID_Alerts"
        End If
        If iVer < 16 Then
            aShow.Add "ID_Eta"
        End If
        If iVer < 17 Then
            aShow.Add "ID_TickDistribution"
        End If
        If iVer < 18 Then
            aShow.Add "ID_ElliotLabels"
        End If
        If iVer < 19 Then
            aShow.Add "ID_ElliotTimeRatio"
        End If
        If iVer < 20 Then
            aShow.Add "ID_Bracket"
        End If
        If iVer < 21 Then
            aShow.Add "ID_RepeatDraw"
        End If
        If iVer < 22 Then
            aShow.Add "ID_DisplacedMA"
            aShow.Add "ID_OscPredictor"
            aShow.Add "ID_DetrendOsc"
            aShow.Add "ID_DiNapoliMACD"
            aShow.Add "ID_PrefStoch"
        End If
        If iVer < 23 Then
            aShow.Add "ID_MacdPredictor"
        End If
        If iVer < 24 Then
            aShow.Add "ID_WhatIf"
        End If
        If iVer < 25 Then
            aShow.Add "ID_TradeFilter"
        End If
        If iVer < 26 Then
            aShow.Add "ID_AutoScale"
            aShow.Add "ID_ResetChart"
        End If
        If iVer < 27 Then
            aShow.Add "ID_ArrowLine"
        End If
        If iVer < 28 Then
            aShow.Add "ID_TrendChannel"
        End If
        If iVer < 29 Then
            aShow.Add "ID_PatternProfit"
        End If
        If iVer < 30 Then
            aShow.Add "ID_PriceAlert"
        End If
        If iVer < 31 Then
            aShow.Add "ID_DanCodeFib"
        End If
        If iVer < 32 Then
            aShow.Add "ID_ChartOrderbar"
        End If
        If iVer < 33 Then
            aShow.Add "ID_Hawkeye"
        End If
        If iVer < 34 Then
            aShow.Add "ID_JPDaily"
            aShow.Add "ID_JPWeekly"
            aShow.Add "ID_JPMonthly"
        End If
        If iVer < 35 Then
            aShow.Add "ID_IndAnalyst"
        End If
        If iVer < 36 Then
            aShow.Add "ID_DanCodeZone"
        End If
        If iVer < 37 Then
            aShow.Add "ID_HBReporter"
        End If
        If iVer < 38 Then
            aShow.Add "ID_UndoDraw"
        End If
        If iVer < 39 Then
            aShow.Add "ID_ShowEWI"
        End If
        If iVer < 40 Then
            aShow.Add "ID_ElliotEndUser"
        End If
        If iVer < 41 Then
            aShow.Add "ID_SAIElite"
        End If
        If iVer < 42 Then
            aShow.Add "ID_SharedPage"
            aShow.Add "ID_Publish"
        End If
        aShow.ToFile App.Path & "\Toolbar.sho"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.CheckNewButtons", eGDRaiseError_Raise

End Sub

Private Sub ToolbarMakeCompatible(tbToolbar As SSActiveToolBars)
On Error GoTo ErrSection:

    Dim frm As Form
        
    Set frm = tbToolbar.Parent
    
    If TypeOf frm Is frmMain Or IsFrmChart(frm) Then
        tbToolbar.ToolBars("General").Visible = False
        tbToolbar.ToolBars(kTbChartSettings).Visible = False
        tbToolbar.ToolBars(kTbWindows).Visible = False
        tbToolbar.ToolBars(kTbDraw).Visible = False
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mMain.ToolbarMakeComaptible", eGDRaiseError_Raise

End Sub

Public Sub BarPeriodCboInit(cbo As Variant)
On Error GoTo ErrSection:

    If Not cbo Is Nothing Then
        With cbo
            .Clear
            If g.FractZen.Allowed Then
                .AddItem "FractZen" '"Auto Breakout"
            End If
            If (ExtremeCharts <> 1) Or HasModule("IT") Then
                .AddItem "1 minute"
                .AddItem "3 minute"
                .AddItem "5 minute"
                .AddItem "10 minute"
                .AddItem "15 minute"
                .AddItem "30 minute"
                .AddItem "60 minute"
                .AddItem "90 minute"
                .AddItem "120 minute"
                .AddItem "180 minute"
                .AddItem "240 minute"
                .AddItem "360 minute"
            End If
            .AddItem "Daily"
            .AddItem "Weekly"
            .AddItem "Monthly"
            .AddItem "Quarterly"
            .AddItem "Yearly"
            .AddItem "< Custom >"
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.BarPeriodCboInit", eGDRaiseError_Raise

End Sub

Private Sub ToolbarInitialShow(tbToolbar As SSActiveToolBars, aShow As cGdArray, ByVal bChart2 As Boolean)
On Error GoTo ErrSection:

    Dim i&, strText$, strFile$
    Dim dDate1#, dDate2#
    Dim Tool As SSTool
    
    If tbToolbar Is Nothing Then Exit Sub
    
    With tbToolbar
        If .Tools("ID_BarPeriod").ComboBox.ListCount = 0 Then
            BarPeriodCboInit .Tools("ID_BarPeriod").ComboBox
        End If
    
        If bChart2 Then
            ' for toolbar on frmChart2 (non-child: to show for ETA)
            .DisplayContextMenu = False
            .ToolBars.Remove "Menu"
            ' buttons to show
            strText = "ID_TopMost;ID_Print;" _
                & "ID_AddToChart;ID_EditChart;ID_Symbol;ID_BarPeriod;" _
                & "ID_Templates;ID_MoreBars;ID_LessBars;ID_MoreAboveBelow;ID_LessAboveBelow;" _
                & "ID_CursorArrow;ID_CursorCrosshairs;ID_ZoomIn;ID_ZoomOut;" _
                & "ID_Trendline;ID_Trendline2;ID_Trendline3;ID_Trendline4;" _
                & "ID_DollarLine;ID_RegressionLine;ID_SRLine;ID_HorzLine;ID_VertLine;" _
                & "ID_Text;ID_Text2;ID_Text3;ID_Text4;ID_Icon;ID_Rectangle;ID_Ellipse;ID_FibCircle;" _
                & "ID_AndrewFork;ID_GannLines;ID_Fibonacci;ID_SpResistFan;ID_TargetShooter;ID_DNExpansion;ID_DNRetracement;" _
                & "ID_TimeCycle;ID_FibTimeZones;ID_FibTimeRatio;ID_FibFan;ID_FibExpansion;ID_ElliotLabels;ID_ElliotEndUser;ID_ElliotTimeRatio;" _
                & "ID_GannacciSwingSquare;ID_GannacciCycle;ID_GannacciTime;ID_GannacciSwing1;ID_GannacciSwing2;ID_AdvRiskReward"
            aShow.SplitFields strText, ";"
        Else
            ' for toolbar on frmMain
            .DisplayContextMenu = False
        
            ' fix names for bar periods on main menu
            With .Tools("ID_BarTimePeriod").Menu
                .Tools("ID_1minute").Name = "1 minute"
                .Tools("ID_3minute").Name = "3 minute"
                .Tools("ID_5minute").Name = "5 minute"
                .Tools("ID_10minute").Name = "10 minute"
                .Tools("ID_15minute").Name = "15 minute"
                .Tools("ID_30minute").Name = "30 minute"
                .Tools("ID_60minute").Name = "60 minute"
                .Tools("ID_90minute").Name = "90 minute"
                .Tools("ID_120minute").Name = "120 minute"
                .Tools("ID_180minute").Name = "180 minute"
                .Tools("ID_240minute").Name = "240 minute"
                .Tools("ID_360minute").Name = "360 minute"
                .Tools("ID_Daily").Name = "&Daily"
                .Tools("ID_Weekly").Name = "&Weekly"
                .Tools("ID_Monthly").Name = "&Monthly"
                .Tools("ID_Quarterly").Name = "&Quarterly"
                .Tools("ID_Yearly").Name = "&Yearly"
                .Tools("ID_CustomPeriod").Name = "&Custom Bar Period"
            End With
        
            ' save which buttons are available on toolbars for use on customizing form
            aShow.Clear
            aShow.Add "General"
            aShow.Add kTbWindows
            aShow.Add kTbChartSettings
            aShow.Add kTbDraw
            frmToolbar.aItems.Clear
            For i = 0 To aShow.Size - 1
                frmToolbar.aItems.Add "=" & aShow(i)
                For Each Tool In .ToolBars(aShow(i)).Tools
                    With Tool
                        If .Type <> ssTypeSeparator Then
                            frmToolbar.aItems.Add .ID
                        End If
                    End With
                Next
            Next
            
            ' ==================== DEFAULTS for which buttons show initially on the Toolbars =================
            ' =================== (no module checking should be done here) ==============================
            ' Get list of buttons to be shown on the toolbar
            If FileExist(App.Path & "\Toolbar.sho") Then
                aShow.FromFile App.Path & "\Toolbar.sho"
                'check file date time & length (to replace provided toolbar template if newer)
                strText = GetIniFileProperty(kTbTemplate, "", kTbIniSection, g.strIniFile)
                If Len(strText) > 0 And UCase(strText) <> "CUSTOM" Then
                    strFile = App.Path & "\Provided\" & strText & ".sho"
                    If FileExist(strFile) Then
                        If FileLen(strFile) > 10 Then
                            dDate1 = GetIniFileProperty(kTbTemplateDate, 0, kTbIniSection, g.strIniFile)
                            dDate2 = CDbl(FileDate(strFile))
                            If RoundNum(dDate2, 9) > RoundNum(dDate1, 9) Then
                                aShow.FromFile strFile
                                If aShow.Size > 0 Then
                                    FileCopy strFile, App.Path & "\Toolbar.sho"
                                    SetIniFileProperty kTbTemplateDate, dDate2, kTbIniSection, g.strIniFile
                                Else
                                    aShow.FromFile App.Path & "\Toolbar.sho"
                                End If
                            End If
                        Else
                            KillFile strFile, True
                        End If
                    End If
                End If
            Else
                ' First-time user: get default toolbar template (4th line of Install.Cfg)
                aShow.FromFile App.Path & "\Provided\Install.cfg"
                strText = Trim(StripStr(aShow(3), Chr(34)))
                aShow.Size = 0
                ' if not exist, just default to "Basic.sho"
                If Len(strText) = 0 Or Not FileExist(App.Path & "\Provided\" & strText) Then
                    If ExtremeCharts >= 1 Then
                        strText = "ETA.sho"
                    Else
                        strText = "Basic.sho"
                    End If
                End If
                strFile = App.Path & "\Provided\" & strText
                aShow.FromFile strFile
                If aShow.Size > 0 Then
                    'since Toolbar.sho does not exist, we want to use icon size & include text specified by template
                    i = Val(Parse(aShow(0), vbTab, 5))
                    'write this to the INI now so it will get picked up later in the toolbar reset code
                    SetIniFileProperty "LargeIcons", i, kTbIniSection, g.strIniFile
                    
                    i = Val(Parse(aShow(0), vbTab, 6))
                    SetIniFileProperty "LargeButtons", i, kTbIniSection, g.strIniFile
                    
                    SetIniFileProperty kTbTemplate, FileBase(strFile), kTbIniSection, g.strIniFile
                    SetIniFileProperty kTbTemplateDate, CDbl(FileDate(strFile)), kTbIniSection, g.strIniFile
                
                    FileCopy strFile, App.Path & "\Toolbar.sho"
                End If
            
                ' And show symbol grid at startup for TSU clients
                If HasModule("TSU", True) Then
                    ' default group for symbol grid
                    strText = "SP100.GRP" ' "FOSO-10.GRP"
                    If FileExist(App.Path & "\Provided\" & strText) Then
                        '[Grid]
                        'ComboID = grp: All Symbols.grp
                        SetIniFileProperty "ComboID", "GRP:" & strText, "Grid", g.strIniFile
                        DockState(frmSymbolGrid) = eDocked
                    End If
                End If
            End If
            
            ' OLD METHOD for initializing the first-time toolbar (is ONLY used now if no provided .SHO files):
            If aShow.Size = 0 Then
                ' default buttons to show for FIRST-TIME USER
                If ExtremeCharts >= 1 Then
                    strText = "17;ID_BarPeriod;ID_Candlesticks;ID_Chain;ID_CloseLine;ID_Components;ID_CursorArrow;" _
                        & "ID_CursorCrosshairs;ID_CursorHorizLine;ID_DNRetracement;ID_DollarLine;ID_Download;ID_EditChart;" _
                        & "ID_Eta;ID_FibExpansion;ID_Fibonacci;ID_FibTimeRatio;ID_FibTimeZones;ID_HorzLine;ID_Mirror;" _
                        & "ID_OHLCBars;ID_Orders;ID_Pattern;ID_Performance;ID_PlanetData;ID_Print;ID_RealTime;ID_Replay;" _
                        & "ID_RiskReward;ID_SectorBrowser;ID_Sectors;ID_Settings;ID_Snapshot;ID_SpResistFan;ID_SRLine;" _
                        & "ID_Subsectors;ID_Symbol;ID_SymbolGrid;ID_TargetShooter;ID_Templates;ID_Text;ID_TickDistribution;" _
                        & "ID_TimeCycle;ID_Toolbox;ID_TradeTracker;ID_Trendline;ID_WaveLabels;ID_ZoomIn;ID_ZoomOut;" _
                        & "ID_AutoScale;ID_ResetChart;ID_ArrowLine;ID_TrendChannel;ID_AdvRiskReward"
                ElseIf HasModule("TSU", True) Then
                    ' default group for symbol grid
                    strText = "SP100.GRP" ' "FOSO-10.GRP"
                    If FileExist(App.Path & "\Provided\" & strText) Then
                        '[Grid]
                        'ComboID = grp: All Symbols.grp
                        SetIniFileProperty "ComboID", "GRP:" & strText, "Grid", g.strIniFile
                        DockState(frmSymbolGrid) = eDocked
                    End If
                    ' toolbar defaults for TradeSmart
                    strText = "35;ID_Download;ID_RealTime;ID_Print;ID_Toolbox;ID_Settings;" _
                        & "ID_Quote;ID_SymbolGrid;ID_Chain;ID_TradeTracker;" _
                        & "ID_Chart;ID_AddToChart;ID_EditChart;ID_Symbol;ID_BarPeriod;" _
                        & "ID_OHLCBars;ID_Candlesticks;ID_CloseLine;" _
                        & "ID_Templates;ID_Pages;ID_MoreBars;ID_LessBars;ID_AutoScale;ID_ResetChart;" _
                        & "ID_CursorArrow;ID_CursorCrosshairs;" _
                        & "ID_ZoomIn;ID_ZoomOut;ID_ChartMove;ID_Eraser;ID_Magnet;ID_PriceAlert;" _
                        & "ID_Trendline;ID_Trendline2;ID_Trendline3;ID_Trendline4;ID_TrendChannel;" _
                        & "ID_SRLine;ID_HorzLine;ID_VertLine;ID_ArrowLine;ID_Text;" _
                        & "ID_Icon;ID_Rectangle;ID_Bracket"
                Else
                    ' normal toolbar defaults
                    strText = "1;ID_Download;ID_RealTime;ID_Print;ID_Toolbox;ID_Settings;ID_CustomizeToolbar;ID_ConditionBuilder;ID_PatternProfit;ID_IndAnalyst;" _
                        & "ID_Quote;ID_Alerts;ID_SymbolGrid;ID_ChartOnOff;ID_ChartData;ID_PlanetData;ID_Snapshot;ID_Chain;ID_TradeTracker;ID_TradeFilter;ID_Performance;ID_Orders;ID_Tile;ID_Replay;ID_Eta;ID_TickDistribution;" _
                        & "ID_Chart;ID_AddToChart;ID_EditChart;ID_Symbol;ID_BarPeriod;" _
                        & "ID_OHLCBars;ID_Candlesticks;ID_BollingerBars;ID_CloseLine;" _
                        & "ID_Templates;ID_Pages;ID_MoreBars;ID_LessBars;ID_MoreAboveBelow;ID_LessAboveBelow;ID_AutoScale;ID_ResetChart;ID_WhatIf;" _
                        & "ID_CursorArrow;ID_CursorCrosshairs;" _
                        & "ID_ZoomIn;ID_ZoomOut;ID_ChartMove;ID_Eraser;ID_Magnet;ID_PriceAlert;" _
                        & "ID_Trendline;ID_Trendline2;ID_Trendline3;ID_Trendline4;ID_DollarLine;ID_RegressionLine;" _
                        & "ID_SRLine;ID_HorzLine;ID_VertLine;ID_Text;ID_Text2;ID_Text3;ID_Text4;ID_Icon;ID_Rectangle;" _
                        & "ID_Ellipse;ID_FibCircle;ID_AndrewFork;ID_GannLines;ID_FibFan;ID_SpResistFan;" _
                        & "ID_Fibonacci;ID_FibExpansion;ID_TargetShooter;ID_DNExpansion;ID_DNRetracement;ID_AdvRiskReward;" _
                        & "ID_TimeCycle;ID_FibTimeZones;ID_FibTimeRatio;ID_Mirror;ID_Pattern;ID_RiskReward;ID_Triangle;" _
                        & "ID_ChannelHighlight;ID_WaveLabels;ID_ElliotLabels;ID_ElliotTimeRatio;ID_ElliotEndUser;ID_Bracket;" _
                        & "ID_DisplacedMA;ID_MacdPredictor;ID_OscPredictor;ID_DetrendOsc;ID_DiNapoliMACD;ID_PrefStoch;ID_ArrowLine;ID_TrendChannel;" _
                        & "ID_GannacciSwingSquare;ID_GannacciCycle;ID_GannacciTime;ID_GannacciSwing1;ID_GannacciSwing2"
                End If
                aShow.SplitFields strText, ";"
                aShow.ToFile App.Path & "\Toolbar.sho"
            End If
            
            CheckNewButtons aShow
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ToolbarInitialShow", eGDRaiseError_Raise

End Sub

Public Sub ToolbarReset(Optional ByVal bReset As Boolean = False, Optional tbToolbar As SSActiveToolBars = Nothing)

    Dim i&, strText$
    Dim bSaveRedraw As Boolean
    Dim bChart As Boolean, bChart2 As Boolean, bShow As Boolean, bShowReports As Boolean
    Dim aShow As New cGdArray
    Dim Toolbar As SSToolBar
    Dim Tool As SSTool
    Dim img As ListImage
    
    Dim bReInitPicBoxTB As Boolean '6437 - if reinit unnecessarily then gets symptoms from issue 6255
    
    Static bAlreadyDone As Boolean

    If IsIDE Then
        On Error GoTo ErrSection '(during development)
    Else
        On Error Resume Next '(for users)
    End If
    
    If tbToolbar Is Nothing Then
        Set tbToolbar = frmMain.tbToolbar
    ElseIf IsFrmChart(tbToolbar.Parent) Then
        bChart = True
    End If
    
    ' init toolbar
    With tbToolbar
        ' MUST turn Redraw off so won't execute any of the toolbar
        ' "Click" code during the Form_Load, etc.
        bSaveRedraw = .Redraw
        .Redraw = False
                
        ' reset the toolbars (to include buttons when first loaded)
        For i = 1 To .ToolBars.Count
            With .ToolBars(i)
                If .Style <> ssMenuBar Then
                    .Reset False
                    .DisplayMoreToolsButton = True
                    .AllowCustomize = False
                    If bChart2 Then
                        .DisplayGrabHandles = False
                        .DockFlags = .DockFlags Or ssPositionLocked
                    End If
                End If
            End With
        Next
        
        If bReset Or Not bAlreadyDone Then
            bReInitPicBoxTB = True

            ToolbarInitialShow tbToolbar, aShow, bChart2
        
            ' Set icons and whether to show on toolbar
            aShow.Sort eGdSort_IgnoreCase Or eGdSort_DeleteNullValues Or eGdSort_DeleteDuplicates
            For Each Tool In .Tools
                Set img = Nothing
                strText = TranslatedText(Tool.ID, Tool.Name)
                If strText <> Tool.Name Then
                    Tool.ChangeAll ssChangeAllName, strText
                End If

                If Tool.Group <> "HelpList" Then
                    If bChart And (Tool.Category = "Window" Or (Tool.Category = "General" And Tool.ID <> "ID_Chart")) Then
                        'don't add general or window buttons to detached chart toolbar
                    Else
                        strText = ToolbarIcon(Tool.ID)
                        If Left(strText, 1) = "k" Then
                            If g.nTbIconStyle = 1 Then
                                If g.nColorTheme = kDarkThemeColor Then
                                    Set img = g.CoreBridge.ImgListToolbarExt("Light", strText, "", 16)
                                    'no need to set disabled picture if image not found (code below will load default classic icon)
                                    If Not img Is Nothing Then Tool.PictureDisabled = img.ExtractIcon
                                Else
                                    Set img = g.CoreBridge.ImgListToolbarExt("Dark", strText, "", 16)
                                End If
                            Else
                                Set img = g.CoreBridge.ImgListToolbarExt("Classic", strText, "", 16)
                            End If

                            If img Is Nothing Then
                                Tool.Picture = Picture16(strText)
                            Else
                                Tool.Picture = img.ExtractIcon
                            End If
                        End If
                        Tool.TagVariant = aShow.BinarySearch(Tool.ID, , eGdSort_IgnoreCase)

                        ' TLB: must make sure all state buttons that aren't in a group are able
                        ' to toggle back up (since settings sometimes get changed on their own!)
                        If Tool.Type = ssTypeStateButton Then
                            If Tool.Group = "" And Not Tool.GroupAllowAllUp Then
                                Tool.GroupAllowAllUp = True
                            End If
                        End If
                    End If
                End If
            Next
            Set img = Nothing
            ' if large fonts, make combo box wider
            .Tools("ID_BarPeriod").ChangeAll ssChangeAllCtlWidth, Int(1200 / Screen.TwipsPerPixelY)
#If 0 Then
            If Screen.TwipsPerPixelY < 15 Then
                'Large fonts
                '.Tools("ID_Download").PictureSize = 17 '20
                i = .Tools("ID_BarPeriod").CtlWidth
                .Tools("ID_BarPeriod").ChangeAll ssChangeAllCtlWidth, Int(i * 1.25 + 6)
            End If
#End If
        
            ' restore cursor
            strText = GetIniFileProperty("CursorType", "ID_CursorCrosshairs", "Charting", g.strIniFile)
            On Error Resume Next
            i = 0
            i = .Tools(strText).Enabled '(just to check for an invalid cursor ID)
            If i = 0 Then strText = "ID_CursorCrosshairs"
            On Error GoTo ErrSection
            .Tools(strText).State = ssChecked
        
            ' restore toolbar positions
            ToolbarLoadPositions tbToolbar, bReset
        
            If g.ChartGlobals.nMagnetValue < 0 Then
                .Tools("ID_Magnet").State = ssUnchecked
            Else
                .Tools("ID_Magnet").State = ssChecked
            End If
            
            If g.ChartGlobals.eDragModeY = eDragModeY_Each Then
                .Tools("ID_DragModeY").State = ssUnchecked
            Else
                .Tools("ID_DragModeY").State = ssChecked
            End If
        End If          'end if bReset or not already done
             
        ' Note: the invisible item needs to exist just so won't get an error related to separators
        .Tools("ID_Invisible").Visible = False
        .Tools("ID_ProcessingStatus").Enabled = False
        
        '======================= CHECK ENABLEMENTS BELOW HERE ===========================
        
        ' REPORTS: need to only show "Reports" on the main menu if one of the reports is enabled ...
        bShowReports = False
        bShow = HasModule("F")
        .Tools("ID_RollsTable").Visible = bShow
        .Tools("ID_COTReport").Visible = bShow
        If bShow = True Then bShowReports = True
         
        ' SAI reports
        bShow = HasModule("SAI_*")
        .Tools("ID_SAIReport").Visible = bShow
        If bShow = True Then bShowReports = True
        bShow = HasModule("SAIE_*")
        .Tools("ID_SAIElite").Visible = bShow
        If bShow = True Then bShowReports = True
        
        ' Highlight Bar Reporter
        bShow = HasPlatinum(False) Or HasModule("HBR")
        .Tools("ID_HBReporter").Visible = bShow
        If bShow = True Then bShowReports = True
        
        ' Sector Analysis
        bShow = HasModule("SECTOR")
        .Tools("ID_SectorWeb").Visible = bShow
        If bShow = True Then bShowReports = True
        
        ' Seasonal Sweet Spot
        bShow = HasModule("SWEET,MOSWT")
        .Tools("ID_SeasonalSP").Visible = bShow
        If bShow = True Then bShowReports = True
        
        ' Stock Screener
        bShow = FileExist(App.Path & "\ScreenerWeb.flg")
        .Tools("ID_ScreenerWeb").Visible = bShow
        If bShow = True Then bShowReports = True
        
        ' News Browser
        bShow = HasModule("NEWS") ' FileExist(App.Path & "\News\NewsBrowser.flg") 'exe")
        .Tools("ID_NewsBrowser").Visible = bShow
        If bShow = True Then bShowReports = True
        
        ' DanielCode Genie
        bShow = HasModule("DCPLUS,DCFOREX,DCFUTURE")
        .Tools("ID_DanCodeWeb").Visible = bShow
        If bShow = True Then bShowReports = True
        bShow = HasModule("GMAJPRO,DCPROFX,DCPROFUT")
        .Tools("ID_GmajPro").Visible = bShow
        If bShow = True Then bShowReports = True
        
        ' only show "Reports" on main menu if one of the above reports is enabled
        .Tools("ID_Reports").Visible = bShowReports
        
        
        ' TOOLS: show based on flag files
        .Tools("ID_TopMost").Visible = bChart2 '(just used for frmChart2)
        ''.Tools("ID_ImageServer").Visible = False
        .Tools("ID_BackupRestore").Visible = True 'False
        ''.Tools("ID_HelpTopics").Visible = FileExist(App.Path & "\Help\*.*")
        If Not FileExist(App.Path & "\testing.mnu") Then
            .Tools("ID_Testing").Visible = False
        End If
        If Not FileExist(App.Path & "\Hume.mod") Then
            .Tools("ID_HumeTools").Visible = False
        End If
        ' only show ETA icon if it has been run
        'If Len(GetIniFileProperty("SimutradeDirectory", "", "GENERAL", "navwin.ini")) = 0 Then
        If IsRule1U Or Not FileExist(App.Path & "\..\Eta\Eta.exe") Then
            .Tools("ID_Eta").Visible = False
        'ElseIf Not FileExist("c:\common\files.exe") Then
        '    .Tools("ID_TradeTracker").Visible = False
        End If

        ' only show RealTime if have server and updating tick data
        'If (HasModule("ST") Or HasModule("FT")) And HasGold(False) _
                And DirExist(App.Path & "\..\RealTime\") Then
        If HasModule("RTG") Or HasModule("RTE") Then
            If .Tools("ID_RealTime").Visible = False Then bReInitPicBoxTB = True
            .Tools("ID_RealTime").Visible = True
        Else
            ' TLB 5/20/2013: to solve #6837 (and other related issues), just always show the traffic light
            ' (enablement will now be checked when they try to turn it on)
            ' TLB 6/7/2013: except for BetterTrades
            If ExtremeCharts = 1 Then
                .Tools("ID_RealTime").Visible = False
            Else
                .Tools("ID_RealTime").Visible = True
            End If
            If g.RealTime.Active Then g.RealTime.Init False, "Not enabled"
        End If
        
        .Tools("ID_MarketProfile").Visible = HasModule("TPRO") Or FileExist(App.Path & "\Provided\MktProf.flg")
        .Tools("ID_PatternProfit").Visible = HasModule("LWPFP,PFP") ' HasPlatinum(False) And FileExist(App.Path & "\Patterns4Profit.flg")
        .Tools("ID_IndAnalyst").Visible = HasModule("LWIA,APFP")
        g.bPatProfitFlag = FileExist(App.Path & "\Patterns4Profit.flg")
        
        ' only if Plat
        bShow = HasPlatinum(False)
        .Tools("ID_Performance").Visible = bShow
        .Tools("ID_ProVersionInfo").Visible = False 'bShow
        
        ' only if at least Gold (not Extreme)
        bShow = HasGold(False, , False)
        .Tools("ID_Orders").Visible = bShow
        .Tools("ID_TickDistribution").Visible = bShow Or HasModule("RTG") '(so shows even before starting realtime)
        .Tools("ID_MarketDepth").Visible = bShow
        .Tools("ID_TimeSales").Visible = bShow
        
        ' only if at least Extreme or Gold
        bShow = HasGold(False, , True)
        .Tools("ID_FibTimeZones").Visible = bShow
        .Tools("ID_FibTimeRatio").Visible = bShow
        .Tools("ID_ElliotTimeRatio").Visible = bShow
        .Tools("ID_TimeCycle").Visible = bShow
        .Tools("ID_Mirror").Visible = bShow
        .Tools("ID_Pattern").Visible = bShow
        .Tools("ID_RiskReward").Visible = bShow
        .Tools("ID_Triangle").Visible = bShow
        .Tools("ID_ChannelHighlight").Visible = bShow
        .Tools("ID_WaveLabels").Visible = bShow
        
        '.Tools("ID_Pages").Visible = HasGold(False, , True) Or HasModule("ROCK*")
        bShow = HasGold(False, , True) Or HasModule("ROCK*")
        If bShow Then
            'this turns on limited chart pages feature so do not do if already have full chart pages feature
            g.ChartGlobals.bMyPageFeature = False
        Else
            'turn on limited chart pages feature if applicable
            g.ChartGlobals.bMyPageFeature = HasModule("LWMC", True) 'Or FileExist(g.strAppPath & "\LWMCPages.flg")
            bShow = g.ChartGlobals.bMyPageFeature
        End If
        .Tools("ID_Pages").Visible = bShow
        .Tools("ID_PagesList").Visible = bShow      'this is ID for menu item (fix for issue 6567)
        
        ' only if not Extreme basic
        bShow = (ExtremeCharts <> 1)
        .Tools("ID_ExportData").Visible = bShow
        .Tools("ID_Rules").Visible = bShow
        .Tools("ID_Strategies").Visible = bShow
        .Tools("ID_StrategyBaskets").Visible = bShow
        .Tools("ID_News").Visible = bShow
        .Tools("ID_HelpTopics").Visible = bShow
        .Tools("ID_WhatsNew").Visible = bShow
        .Tools("ID_1minute").Visible = bShow
        .Tools("ID_3minute").Visible = bShow
        .Tools("ID_5minute").Visible = bShow
        .Tools("ID_10minute").Visible = bShow
        .Tools("ID_15minute").Visible = bShow
        .Tools("ID_30minute").Visible = bShow
        .Tools("ID_60minute").Visible = bShow
        .Tools("ID_90minute").Visible = bShow
        .Tools("ID_120minute").Visible = bShow
        .Tools("ID_180minute").Visible = bShow
        .Tools("ID_240minute").Visible = bShow
        .Tools("ID_360minute").Visible = bShow
        If ExtremeCharts = 1 And Not HasModule("IT") Then
            ' remove the intraday periods from the dropdown
            For i = .Tools("ID_BarPeriod").ComboBox.ListCount - 1 To 0 Step -1
                If InStr(UCase(.Tools("ID_BarPeriod").ComboBox.List(i)), "MINUTE") > 0 Then
                    .Tools("ID_BarPeriod").ComboBox.RemoveItem i
                    If .Parent.cboBarPeriod.ListCount > i Then
                        .Parent.cboBarPeriod.RemoveItem i
                    End If
                End If
            Next
        End If

        ' only show certain drawing tools if have correct modules
        .Tools("ID_IOAMT").Visible = HasModule("IOAMT")
        .Tools("ID_VolumeAtPrice").Visible = HasModule("IOAMT")
        .Tools("ID_TimeSalesAnalyzer").Visible = HasModule("IOAMT")
        .Tools("ID_BidAskDir").Visible = HasModule("IOAMT")
        .Tools("ID_PlanetData").Visible = HasModule("ASTR")
        .Tools("ID_TargetShooter").Visible = HasModule("BAT,PPWL,PRG11,LW*") 'LWCRACK,LWART,LWST,LWSTX
        .Tools("ID_DNRetracement").Visible = HasModule("FIB")
        .Tools("ID_PivotPoints").Visible = HasModule("PVT") Or HasLevel(eTN4_Gold, False)
        
        If FileExist(App.Path & "\ewave.flg") Or FileExist(App.Path & "\gmp.flg") Then
            .Tools("ID_ElliotLabels").Visible = True
            .Tools("ID_ElliotEndUser").Visible = False
            .Tools("ID_ShowEWI").Visible = False
        Else
            .Tools("ID_ElliotLabels").Visible = False
            .Tools("ID_ElliotEndUser").Visible = HasModule("EWL")
        End If
        
        ' Option Navigator
        .Tools("ID_Chain").Visible = HasModule("FO,SO")
        .Tools("ID_OptionNavigator").Visible = False
        
        ' Fib clusters
        .Tools("ID_FibClusters").Visible = HasModule("ADVFIB") ' FileExist(App.Path & "\Cluster.flg")
        
        .Tools("ID_Hawkeye").Visible = HasModule("HKADDS", False)
        
        ' WhatIf -- Gold, or if a TSU user has Standard (per Pete 4/13/2011)
        If HasLevel(eTN4_Gold, False) Then
            .Tools("ID_WhatIf").Visible = True
        ElseIf HasLevel(eTN3_Standard, False) And HasModule("TSU") Then
            .Tools("ID_WhatIf").Visible = True
        Else
            .Tools("ID_WhatIf").Visible = False
        End If
        
        ' if FIB then give tool as "DiNapoli Expansion", else if has Gold then
        ' give tool as "Fibonacci Extension", else don't give tool at all
        'If HasModule("FIB") Then
        .Tools("ID_FibABCD").Visible = HasModule("ADVFIB") ' FileExist(App.Path & "\FibOverride.flg")
        .Tools("ID_Gartley").Visible = HasModule("ADVFIB") 'FileExist(App.Path & "\FibOverride.flg")
        If UseDiNapFib Then
            .Tools("ID_DNExpansion").Visible = True
            bShow = False
        ElseIf HasGold(False) Then
            .Tools("ID_DNExpansion").Visible = True
            strText = "Fibonacci Extension"
            If .Tools("ID_DNExpansion").Name <> strText Then
                .Tools("ID_DNExpansion").ToolTipText = strText
                .Tools("ID_DNExpansion").ChangeAll ssChangeAllName, strText
            End If
            bShow = True
        Else
            .Tools("ID_DNExpansion").Visible = False
            bShow = False
        End If
        .Tools("ID_DNExpansion2").Visible = bShow
        .Tools("ID_DNExpansion3").Visible = bShow
        .Tools("ID_DNExpansion4").Visible = bShow
    
        ' if GREENBLATT then give tool as "Greenblatt Time Zone"
        If HasModule("JGREEN") Then
            strText = "Greenblatt Time Zone"
            If .Tools("ID_FibTimeZones").Name <> strText Then
                .Tools("ID_FibTimeZones").ToolTipText = strText
                .Tools("ID_FibTimeZones").ChangeAll ssChangeAllName, strText
            End If
        End If
        
        'if CODE then show Daniel Code Retracement tool
        .Tools("ID_DanCodeFib").Visible = HasModule("CODE")
        .Tools("ID_DanCodeZone").Visible = HasModule("CODE")
        
        'if JPC then show JP buttons
        bShow = HasModule("JPC")
        .Tools("ID_JPDaily").Visible = bShow
        .Tools("ID_JPWeekly").Visible = bShow
        .Tools("ID_JPMonthly").Visible = bShow
        .Tools("ID_JPQuarterly").Visible = bShow
        .Tools("ID_JPExpiration").Visible = bShow
        
        ' buttons for Coast Trading Package
        bShow = HasModule("CTP")
        .Tools("ID_DisplacedMA").Visible = bShow
        .Tools("ID_OscPredictor").Visible = bShow
        .Tools("ID_DetrendOsc").Visible = bShow
        .Tools("ID_DiNapoliMACD").Visible = bShow
        .Tools("ID_PrefStoch").Visible = bShow
        .Tools("ID_MacdPredictor").Visible = HasModule("FIB") '(this one requires the Fib module)
        
        ' Woodies CCI
        '.Tools("ID_TradeFilter").Visible = HasModule("WOODCCI")
        .Tools("ID_TradeFilter").Visible = True
        If IsWoodiesVersion Then
            If .Tools("ID_TradeFilter").Name <> "Trade Filter" Then
                .Tools("ID_TradeFilter").ChangeAll ssChangeAllName, "Trade Filter"
            End If
        Else
            If .Tools("ID_TradeFilter").Name <> "Trade Reports" Then
                .Tools("ID_TradeFilter").ChangeAll ssChangeAllName, "Trade Reports"
            End If
        End If
        
        ' Turnkey
        .Tools("ID_Cattle").Visible = g.CattleBridge.IsCattleUser
        .Tools("ID_Turnkey").Visible = g.CattleBridge.IsTurnkeyUser
        
        'Better Trades RPM (aka Balloon Strangle) tool
        .Tools("ID_BalloonStrangle").Visible = HasModule("BTRPM")

        'Gannacci tools
        bShow = HasModule("WCT")        'FileExist("Gannacci.flg")
        .Tools("ID_GannacciSwingSquare").Visible = bShow
        .Tools("ID_GannacciCycle").Visible = bShow
        .Tools("ID_GannacciTime").Visible = bShow
        .Tools("ID_GannacciSwing1").Visible = bShow
        .Tools("ID_GannacciSwing2").Visible = bShow
        
        'Publish
        bShow = HasGold(False) And FileExist(g.strAppPath & "\SCP.flg")
        .Tools("ID_Publish").Visible = bShow
        'Shared Chart Page
        If bShow Or (HasGold(False) And HasModule("SCP_*")) Then bShow = True
        .Tools("ID_SharedPage").Visible = bShow
        
        ' Large Icons
        .LargeIcons = GetIniFileProperty("LargeIcons", ssUnchecked, "Toolbars", g.strIniFile)
        g.nTbLargeIcons = Abs(.LargeIcons)
        ' Buttons with text

'08-08-2012: Originally hard-coded for TSU to include text & exclude text for everyone else
'            Per Tim/Pete, change this to be picked up from toolbar template file when applicatble
'        If HasModule("TSU", True) Then
'            g.nTbIncludeText = 1
'        Else
'            g.nTbIncludeText = 0
'        End If
'        g.nTbIncludeText = Abs(GetIniFileProperty("LargeButtons", g.nTbIncludeText, "Toolbars", g.strIniFile))
        g.nTbIncludeText = Abs(GetIniFileProperty("LargeButtons", 0, "Toolbars", g.strIniFile))
        
        If TypeOf tbToolbar.Parent Is frmMain Then
            If g.nTbIncludeText = 1 Then
                If g.nTbLargeIcons Then
                    tbToolbar.Parent.ToolBarBtnSizeSet "", kBtnLargeIcoTextWd, kBtnLargeIcoTextHt
                Else
                    tbToolbar.Parent.ToolBarBtnSizeSet "", kBtnSmallIcoTextWd, kBtnSmallIcoTextHt
                End If
            ElseIf g.nTbLargeIcons Then
                tbToolbar.Parent.ToolBarBtnSizeSet "", kBtnLargeIco, kBtnLargeIco
            Else
                tbToolbar.Parent.ToolBarBtnSizeSet "", kBtnSmallIco, kBtnSmallIco
            End If
        End If
        'for new picture box toolbar
        tbToolbar.Parent.ToolBarWrapSet "", GetIniFileProperty("ToolbarWrap", False, "Toolbars", g.strIniFile)
        
        ' Hide buttons on the toolbar, or include text with buttons if using that option
        On Error Resume Next
        For Each Toolbar In .ToolBars
            If Toolbar.Style <> ssMenuBar Then
                For i = Toolbar.Tools.Count To 1 Step -1 '(go backwards since may be removing some)
                    With Toolbar.Tools(i)
                        .AutoWrap = False ' (so won't wrap to next line)
                        If .Type <> ssTypeSeparator Then
                            If .TagVariant = False Then
                            'If Not aShow.BinarySearch(.ID, , eGdSort_IgnoreCase) Then
                                Toolbar.Tools.Remove i
                            ElseIf .Type = ssTypeButton Or .Type = ssTypeStateButton Then
                                .CaptionAlignment = ssRightOfImage
                                If .DisplayStyle <> ssDisplayTextOnlyAlways Then
                                    ' don't include text on toolbars docked to right or left
                                    If g.nTbIncludeText And Toolbar.DockedStatus <> ssDockedLeft And Toolbar.DockedStatus <> ssDockedRight Then
                                        .DisplayStyle = ssDisplayImageAndText
                                    Else
                                        .DisplayStyle = ssDisplayDefaultStyle
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next
                If Toolbar.Tools.Count = 0 Then
                    Toolbar.Visible = False
                End If
            End If
        Next
        .Tools("ID_Status").AutoWrap = False 'True
                
        If g.ChartGlobals.eChartMode = eMode_Move Then
            .Tools("ID_ChartMove").State = ssChecked
        ElseIf g.ChartGlobals.eChartMode = eMode_Erase Then
            .Tools("ID_Eraser").State = ssChecked
        Else
            .Tools("ID_Zoom").State = ssChecked
        End If
        
        ToolbarMakeCompatible tbToolbar
                        
        .Redraw = bSaveRedraw
    End With
        
    bAlreadyDone = True
    Set tbToolbar = Nothing
    
    If Not g.bStarting Then
        If bReInitPicBoxTB Then ToolbarInit2 frmMain, frmMain.TbButtonsArray(kTbGeneral)
        ToolbarResize2 frmMain, frmMain.pbTbBack, frmMain.imgTbBack, frmMain.TbButtonsArray(kTbGeneral), frmMain.ToolBarWrapGet(kTbGeneral)
    End If
                
    Exit Sub
    
ErrSection:
    bAlreadyDone = True
    tbToolbar.Redraw = bSaveRedraw
    Set tbToolbar = Nothing
    RaiseError "mMain.ToolbarReset", eGDRaiseError_Raise
End Sub

Private Sub ToolbarLoadPositions(tbToolbar As SSActiveToolBars, Optional ByVal bReset As Boolean = False)
On Error Resume Next
    
    Dim s$, i&, bAToolbarIsVisible As Boolean
    Static bAlreadyDone As Boolean
    
    For i = 1 To tbToolbar.ToolBars.Count
        With tbToolbar.ToolBars(i)
            If .Style <> ssMenuBar Then
                If Not bReset Then
                    s = GetIniFileProperty(StripStr("TB3_" & .Name, " "), "", "Toolbars", g.strIniFile)
                End If
                If Len(Parse(s, ";", 4)) > 0 Then
                    .DockedStatus = Val(Parse(s, ";", 1))
                    .DockedRow = Val(Parse(s, ";", 2))
                    .DockedColumn = Val(Parse(s, ";", 3))
                    .Visible = Val(Parse(s, ";", 4))
                    If Val(Parse(s, ";", 4)) <> 0 Then
                        bAToolbarIsVisible = True
                    End If
                    If Val(Parse(s, ";", 5)) >= 0 Then
                        .FloatingLeft = Val(Parse(s, ";", 5))
                        .FloatingTop = Val(Parse(s, ";", 6))
                        .FloatingWidth = Val(Parse(s, ";", 7))
                        .FloatingHeight = Val(Parse(s, ";", 8))
                    End If
                Else
                    ' default positions:
                    bAToolbarIsVisible = True
                    .Visible = True
                    Select Case .Name
                    Case "General"
                        .DockedStatus = ssDockedTop
                        .DockedRow = 2
                        .DockedColumn = 1
                    Case kTbWindows
                        .DockedStatus = ssDockedTop
                        .DockedRow = 2
                        .DockedColumn = 2
                    Case kTbChartSettings
                        .DockedStatus = ssDockedTop
                        .DockedRow = 2
                        .DockedColumn = 3
                    Case kTbDraw
                        If TypeOf tbToolbar.Parent Is frmChart2 Then    'JM: not sure this matters
                            .DockedStatus = ssDockedTop
                            .DockedRow = 2
                            .DockedColumn = 4
                        Else
                            .DockedStatus = ssDockedRight
                            .DockedRow = 1
                            .DockedColumn = 1
                        End If
                    End Select
                End If
            End If
        End With
    Next
    
    ' first time: if no toolbars are visible (most likely
    ' due to a shutdown error), then turn them all on
    If Not bAlreadyDone And Not bAToolbarIsVisible Then
        For i = 1 To tbToolbar.ToolBars.Count
            With tbToolbar.ToolBars(i)
                If .Style <> ssMenuBar Then
                    .Visible = True
                End If
            End With
        Next
    End If
    bAlreadyDone = True

End Sub

Public Sub ToolbarSavePositions()
On Error Resume Next
    
    Dim s$, i&
    
    For i = 1 To frmMain.tbToolbar.ToolBars.Count
        With frmMain.tbToolbar.ToolBars(i)
            If .Style <> ssMenuBar Then
                s = Str(.DockedStatus) & ";" & Str(.DockedRow) & ";" & Str(.DockedColumn) _
                    & ";" & Str(.Visible) & ";" & Str(.FloatingLeft) & ";" & Str(.FloatingTop) _
                    & ";" & Str(.FloatingWidth) & ";" & Str(.FloatingHeight)
                SetIniFileProperty StripStr("TB3_" & .Name, " "), s, "Toolbars", g.strIniFile
            End If
        End With
    Next

End Sub

' return the Icon (in the images collection) associated with the toolbar button
Public Function ToolbarIcon(ByVal strID$) As String
On Error Resume Next

    Dim strIcon$
    
    'JM:03-30-2009 - running all Picture16 icons through this routine so can easily change icons on forms

    'these are not associated with a tool so just return the passed in string
    'code will be added to return new icons as they become available
    If InStr(strID, "_") = 0 Then
        ToolbarIcon = strID
        Exit Function
    End If
    
    ' for these, use same as another tool
    Select Case strID
    Case "ID_PrintChartData"
        strID = "ID_ChartData"
    Case "ID_PrintSnapshot"
        strID = "ID_Snapshot"
    Case "ID_PrintNews"
        strID = "ID_News"
    Case "ID_PrintOptionsChain"
        strID = "ID_Chain"
    Case "ID_PrintQuoteBoard"
        strID = "ID_Quote"
    Case "ID_PrintSymbolGrid"
        strID = "ID_SymbolGrid"
    Case "ID_PrintTradeConsole"
        strID = "ID_TradeTracker"
    Case "ID_PrintMenu"
        strID = "ID_Print"
    Case "ID_PrintNews"
        strID = "ID_News"
    End Select
    
    Select Case strID
    Case "ID_About"
        strIcon = "kNav"
    Case "ID_Alerts"
        strIcon = "kBell"
    Case "ID_AndrewFork"
        strIcon = "kPitchfork"
    Case "ID_AutoScale"
        strIcon = "kAutoScale"
    Case "ID_BarPeriod"
        strIcon = "kCombo"
    Case "ID_BollingerBars"
        strIcon = "kBollBars"
    Case "ID_Candlesticks"
        strIcon = "kCandlesticks"
    Case "ID_Chain"
        strIcon = "kChain"
    Case "ID_ChannelHighlight"
        strIcon = "kChannel"
    Case "ID_ChartData"
        strIcon = "kChartData"
    Case "ID_ChartMove"
        'JM 05-18-2009: g.ChartGlobals.eScaleMode does not seem to be used anywhere
        '   look into removing this after May 27 release is stable
        If g.ChartGlobals.eScaleMode = ePANE_ScaleModeManual Then
            strIcon = "kChartMove"
        Else
            strIcon = "kChartMoveHz"
        End If
    Case "ID_ChartOnOff"
        strIcon = "kChartOnOff_Show"
    Case "ID_CloseLine"
        strIcon = "kCloseLine"
    Case "ID_ConditionBuilder"
        strIcon = "kConditionBuilder"
    Case "ID_COTReport"
        strIcon = "kReport"
    Case "ID_Criteria"
        strIcon = "kCriteria"
    Case "ID_CursorArrow"
        strIcon = "kArrow"
    Case "ID_CursorCrosshairs"
        strIcon = "kCrosshairs"
    Case "ID_CursorHorizLine"
        strIcon = "kHorizCursor"
    Case "ID_CursorVertLine"
        strIcon = "kVertCursor"
    Case "ID_CustomizeToolbar"
        strIcon = "kToolbar"
    Case "ID_CustomizeToolbar"
        strIcon = "kToolbar"
    Case "ID_DNExpansion"
        If UseDiNapFib() Then
            strIcon = "kExpansion"
        Else
            strIcon = "kExpansionGen"
        End If
    Case "ID_DNExpansion2"
        strIcon = "kExpansionGen2"
    Case "ID_DNExpansion3"
        strIcon = "kExpansionGen3"
    Case "ID_DNExpansion4"
        strIcon = "kExpansionGen4"
    Case "ID_DNRetracement"
        strIcon = "kRetracement"
    Case "ID_DollarLine"
        strIcon = "kDollarLine"
    Case "ID_DollarLine2"
        strIcon = "kDollarLine2"
    Case "ID_DollarLine3"
        strIcon = "kDollarLine3"
    Case "ID_DollarLine4"
        strIcon = "kDollarLine4"
    Case "ID_GannacciSwingSquare"
        strIcon = "kGannacciSwingSquare"
    Case "ID_GannacciCycle"
        strIcon = "kGannacciCycle"
    Case "ID_GannacciTime"
        strIcon = "kGannacciTime"
    Case "ID_GannacciSwing1"
        strIcon = "kGannacciSwing1"
    Case "ID_GannacciSwing2"
        strIcon = "kGannacciSwing2"
    Case "ID_Download"
        strIcon = "kDownload"
    Case "ID_DragModeY"
        strIcon = "kDragModeY"
    Case "ID_ElliotLabels"
        strIcon = "kElliotLabels"
    Case "ID_ElliotTimeRatio"
        strIcon = "kElliotTimeRatio"
    Case "ID_Ellipse"
        strIcon = "kEllipse"
    Case "ID_Eraser"
        strIcon = "kEraser"
    Case "ID_Eta"
        strIcon = "kEta"
    Case "ID_ExportData"
        strIcon = "kExport"
    Case "ID_FibCircle"
        strIcon = "kFibCircle"
    Case "ID_FibExpansion"
        strIcon = "kFibExpansion"
    Case "ID_FibFan"
        strIcon = "kFibFan"
    Case "ID_Fibonacci"
        strIcon = "kFib"
    Case "ID_Fibonacci2"
        strIcon = "kFib2"
    Case "ID_Fibonacci3"
        strIcon = "kFib3"
    Case "ID_Fibonacci4"
        strIcon = "kFib4"
    Case "ID_FibTimeRatio"
        strIcon = "kFibTime"
    Case "ID_FibTimeZones"
        strIcon = "kFibTimeZones"
    Case "ID_Filters"
        strIcon = "kFilter"
    Case "ID_Functions"
        strIcon = "kFunction"
    Case "ID_GannLines"
        strIcon = "kGannFan"
    Case "ID_HorzLine"
        strIcon = "kHorzLine"
    Case "ID_HorzLine2"
        strIcon = "kHorzLine2"
    Case "ID_HorzLine3"
        strIcon = "kHorzLine3"
    Case "ID_HorzLine4"
        strIcon = "kHorzLine4"
    Case "ID_Icon"
        strIcon = "kIcon"
    Case "ID_Kagi"
        strIcon = "kKagi"
    Case "ID_LessAboveBelow"
        strIcon = "kLessSpace"
    Case "ID_LessBars"
        strIcon = "kLessBars"
    Case "ID_Libraries"
        strIcon = "kLibrary"
    Case "ID_Magnet"
        strIcon = "kMagnet"
    Case "ID_MarketDepth"
        strIcon = "kMarketDepth"
    Case "ID_MarketProfile"
        strIcon = "kMarketProfile"
    Case "ID_Mirror"
        strIcon = "kMirror"
    Case "ID_MoreAboveBelow"
        strIcon = "kMoreSpace"
    Case "ID_MoreBars"
        strIcon = "kMoreBars"
    Case "ID_Mountain"
        strIcon = "kMountain"
    Case "ID_News"
        strIcon = "kNote"
    Case "ID_OHLCBars"
        strIcon = "kOHLC"
    Case "ID_Orders"
        strIcon = "kOrders"
    Case "ID_Pages"
        'strIcon = "kSelect"
    Case "ID_Pattern"
        strIcon = "kCopyPattern"
    Case "ID_Performance"
        strIcon = "kPerformance"
    Case "ID_PlanetData"
        strIcon = "kPlanet"
    Case "ID_PointFigure"
        strIcon = "kPointFigure"
    Case "ID_Print"
        strIcon = "kPrint"
    Case "ID_Quote"
        strIcon = "kQuote"
    Case "ID_RealTime"
        strIcon = "kRedLight"
    Case "ID_RecalcCriteria"
        strIcon = "kCriteria"
    Case "ID_Rectangle"
        strIcon = "kRectEllipse"
    Case "ID_RegressionLine"
        strIcon = "kRegressionLine"
    Case "ID_Renko"
        strIcon = "kRenko"
    Case "ID_Replay"
        strIcon = "kReplay"
    Case "ID_ResetChart"
        strIcon = "kResetChart"
    Case "ID_RiskReward"
        strIcon = "kRiskReward"
    Case "ID_RollsTable"
        strIcon = "kReportRolls"
    Case "ID_Rules"
        strIcon = "kRule"
    Case "ID_SectorBrowser"
        strIcon = "kSectors"
    Case "ID_Settings"
        strIcon = "kSettings"
    Case "ID_Snapshot"
        strIcon = "kFundamentals"
    Case "ID_SpResistFan"
        strIcon = "kSpeedResistance"
    Case "ID_SRLine"
        strIcon = "kSRLine"
    Case "ID_SRLine2"
        strIcon = "kSRLine2"
    Case "ID_SRLine3"
        strIcon = "kSRLine3"
    Case "ID_SRLine4"
        strIcon = "kSRLine4"
    Case "ID_Strategies"
        strIcon = "kSystem"
    Case "ID_StrategyBaskets"
        strIcon = "kBasket"
    Case "ID_SymbolGrid"
        strIcon = "kSymbolGrid"
    Case "ID_SymbolGroups"
        strIcon = "kSymbolGroup"
    Case "ID_TargetShooter"
        strIcon = "kTargetShooter"
    Case "ID_Templates"
        'strIcon = "kSelect"
    Case "ID_Text"
        strIcon = "kText"
    Case "ID_Text2"
        strIcon = "kText2"
    Case "ID_Text3"
        strIcon = "kText3"
    Case "ID_Text4"
        strIcon = "kText4"
    Case "ID_TickDistribution"
        strIcon = "kTickDist"
    Case "ID_TimeCycle"
        strIcon = "kTimeCycle"
    Case "ID_TimeSales"
        strIcon = "kTimeSales"
    Case "ID_TimeSalesAnalyzer"
        strIcon = "kTSAnalyzer"
    Case "ID_Toolbox"
        strIcon = "kTools"
    Case "ID_TradeTracker"
        strIcon = "kDollar"
    Case "ID_TradeFilter"
        If IsWoodiesVersion Then
            strIcon = "kWoodCCI"
        Else
            strIcon = "kTradeFilter"            '6412
        End If
    Case "ID_Trendline"
        strIcon = "kTrendline"
    Case "ID_Trendline2"
        strIcon = "kTrendline2"
    Case "ID_Trendline3"
        strIcon = "kTrendline3"
    Case "ID_Trendline4"
        strIcon = "kTrendline4"
    Case "ID_Triangle"
        strIcon = "kTriangle"
    Case "ID_VertLine"
        strIcon = "kVertLine"
    Case "ID_ViewJournals"
        strIcon = "kScroll"
    Case "ID_VolumeAtPrice"
        strIcon = "kVolumeAtPrice"
    Case "ID_WaveLabels"
        strIcon = "kWave"
    Case "ID_WhatIf"
        strIcon = "kWhatIf"
    Case "ID_WhatsNew"
        strIcon = "kExclamation"
    Case "ID_ZoomIn"
        strIcon = "kZoomIn"
    Case "ID_ZoomOut"
        strIcon = "kZoomOut"
    Case "ID_Bracket"
        strIcon = "kBracket"
    Case "ID_RepeatDraw"
        strIcon = "kRepeatDraw"
    Case "ID_DisplacedMA"
        strIcon = "kDisplacedMA"
    Case "ID_OscPredictor"
        strIcon = "kOscPredictor"
    Case "ID_DetrendOsc"
        strIcon = "kDetrendOsc"
    Case "ID_DiNapoliMACD"
        strIcon = "kDinapoliMACD"
    Case "ID_PrefStoch"
        strIcon = "kPrefStoch"
    Case "ID_MacdPredictor"
        strIcon = "kMacdPredictor"
    Case "ID_PivotPoints"
        strIcon = "kPivotPoints"
    Case "ID_FibClusters"
        strIcon = "kFibClusters"
    Case "ID_ArrowLine"
        strIcon = "kArrowLine"
    Case "ID_TrendChannel"
        strIcon = "kTrendChannel"
    Case "ID_PriceAlert"
        strIcon = "kPriceAlert"
    Case "ID_PatternProfit"
        strIcon = "kMoneyBag"
    Case "ID_IndAnalyst"
        strIcon = "kIndAnalyst"
    Case "ID_DanCodeFib"
        strIcon = "kDanCodeFib"
    Case "ID_Hawkeye"
        strIcon = "kHawkeye"
    Case "ID_DanCodeWeb"
        strIcon = "kDanCodeWeb"
    Case "ID_GmajPro"
        strIcon = "kGmajProW"
    Case "ID_ChartOrderbar"
        strIcon = "kLightning"
    Case "ID_SectorWeb"
        strIcon = "kSectorWeb"
    Case "ID_JPDaily"
        strIcon = "kJPDaily"
    Case "ID_JPWeekly"
        strIcon = "kJPWeekly"
    Case "ID_JPMonthly"
        strIcon = "kJPMonthly"
    Case "ID_JPQuarterly"
        strIcon = "kJPQuarterly"
    Case "ID_JPExpiration"
        strIcon = "kJPExpiration"
    Case "ID_NewsBrowser"
        strIcon = "kNewsBrowser"
    Case "ID_FibABCD"
        strIcon = "kFibABCD"
    Case "ID_Gartley"
        strIcon = "kGartley"
    Case "ID_SeasonalSP"
        strIcon = "kSeasonalSP"
    Case "ID_DanCodeZone"
        strIcon = "kDanCodeZone"
    Case "ID_HBReporter"
        strIcon = "kHBReporter"
    Case "ID_UndoDraw"
        strIcon = "kUndoEnabled"
    Case "ID_ScreenerWeb"
        strIcon = "kScreenerWeb"
    Case "ID_TextIncrease"
        strIcon = "kTextIncrease"
    Case "ID_TextDecrease"
        strIcon = "kTextDecrease"
    Case "ID_SAIReport"
        strIcon = "kSAI"
    Case "ID_SAIElite"
        strIcon = "kSAI"
    Case "ID_Cattle"
        strIcon = "kCattle"
    Case "ID_Turnkey"
        strIcon = "kHedgeLinc"
    Case "ID_Publish"
        strIcon = "kPublish"
    Case "ID_SharedPage"
        strIcon = "kSharedChartPage"
    Case "ID_ShowEWI"
        strIcon = "kElliotLabelsOn"
    Case "ID_ElliotEndUser"
        strIcon = "kElliotEndUser"
    Case "ID_BalloonStrangle"
        strIcon = "kBalloonStrangle"
    Case "ID_AdvRiskReward"
        strIcon = "kAdvRiskReward"
    
    Case "ID_Chart"
        If g.nTbIconStyle = 1 Then
            strIcon = "kChartNew4"
        ElseIf g.eTbSkin = eTbSkin_Blue Or g.eTbSkin = eTbSkin_ALuminumBlue Then
            strIcon = "kChartNew2"
        Else
            strIcon = "kChartNew3"
        End If
    
    Case "ID_Tile"
        If g.nTbIconStyle = 1 Then
            strIcon = "kCharts4"
        ElseIf g.eTbSkin = eTbSkin_Blue Or g.eTbSkin = eTbSkin_ALuminumBlue Then
            strIcon = "kCharts2"
        Else
            strIcon = "kCharts3"
        End If
    
    Case "ID_EditChart"
        If g.nTbIconStyle = 1 Then
            strIcon = "kChartEdit4"
        ElseIf g.eTbSkin = eTbSkin_Blue Or g.eTbSkin = eTbSkin_ALuminumBlue Then
            strIcon = "kChartEdit2"
        Else
            strIcon = "kChartEdit3"
        End If
    
    Case "ID_AddToChart"
        If g.nTbIconStyle = 1 Then
            strIcon = "kChartAdd4"
        ElseIf g.eTbSkin = eTbSkin_Blue Or g.eTbSkin = eTbSkin_ALuminumBlue Then
            strIcon = "kChartAdd2"
        Else
            strIcon = "kChartAdd3"
        End If
    
    Case "ID_Symbol"
        If g.nTbIconStyle = 1 Then
            strIcon = "kChangeSymbol4"
        ElseIf g.eTbSkin = eTbSkin_Blue Or g.eTbSkin = eTbSkin_ALuminumBlue Then
            strIcon = "kChangeSymbol2"
        Else
            strIcon = "kChangeSymbol3"
        End If
    
    End Select
    
    ToolbarIcon = strIcon
    
End Function

' To display tooltips for grid
' - for header row, will display "Show by ..."
' - if not header row, will display tooltip contained in nTipCol (if >= 0)
Public Sub GridTooltip(Grid As VSFlexGrid, Optional ByVal nTipCol As Long = -1, _
        Optional ByVal strColName As String = "")
On Error Resume Next

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    Dim strTip As String
    Dim strSymbol As String
    Static strPrevSymbol As String
    Static strDesc As String
    
    With Grid
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        If lMouseRow >= 0 And lMouseRow < .Rows And lMouseCol >= 0 And lMouseCol < .Cols Then
            If lMouseRow < .FixedRows Then
                ' if header, show "Sort by ..."
                If Len(strColName) = 0 Then
                    strColName = Trim(.TextMatrix(lMouseRow, lMouseCol))
                End If
                If Len(strColName) > 0 And strColName <> "-" Then
                    strTip = SORT_BY_PREFIX & strColName
                End If
            ElseIf nTipCol >= 0 And nTipCol < .Cols Then
                ' show tooltip contained in specified column
                strTip = TipStr(.TextMatrix(lMouseRow, nTipCol))
            ElseIf .FixedRows > 0 Then
                ' if symbol, show description as tooltip
                If (Trim(UCase(.TextMatrix(0, lMouseCol))) = "SYMBOL") Then
                    If Len(strColName) = 0 Then
                        strSymbol = UCase(Trim(Parse(.TextMatrix(lMouseRow, lMouseCol), "(", 1)))
                        If strSymbol <> strPrevSymbol Then
                            strPrevSymbol = strSymbol
                            strDesc = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(strSymbol))
                        End If
                        strTip = strDesc
                    Else
                        strTip = strColName
                    End If
                End If
            End If
        End If
        .ToolTipText = strTip
    End With

End Sub

Public Function OpenDatabase(ByVal strMDB As String, ByVal strPassword As String) As Database
On Error GoTo ErrSection:

    If Len(strPassword) > 0 Then
        Set OpenDatabase = g.WrkJet.OpenDatabase(strMDB, False, False, "; PWD=" & strPassword)
    Else
        Set OpenDatabase = g.WrkJet.OpenDatabase(strMDB, False, False)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    Dim strBackup$, dBackupDate#, strMsg$, bRestored As Boolean
    
    ' Check for specific types of errors related to a corrupted database:
    ' 3343 = database format not recognized
    ' 3011 = could not find database object (e.g. hidden or system object in the MDB)
    If Err.Number = 3343 Or Err.Number = 3011 Then
        Select Case strMDB
            Case AddSlash(App.Path) & "Libraries.MDB"
                strBackup = AddSlash(App.Path) & "LibBak1.MDB"
                strMsg = "Libraries.MDB"
            Case AddSlash(App.Path) & "TradeTracker.mdb"
                strBackup = AddSlash(App.Path) & "TTBak1.MDB"
                strMsg = "TradeTracker.MDB"
        End Select
        If FileExist(strBackup) Then
            dBackupDate = FileDate(strBackup)
            strMsg = "The " & strMsg & " is corrupted.|Would you like to restore your backup copy |from " _
                & DateAndTime(dBackupDate) & "?"
            If InfBox(strMsg, "?", "+Yes|-No", "Error") = "Y" Then
                FileCopy strBackup, strMDB, True
                Set OpenDatabase = OpenDatabase(strMDB, strPassword)
                bRestored = True
            Else
                Err.Raise vbObjectError + 1000, , strMDB & " could not be opened"
            End If
        End If
    End If
    If Not bRestored Then
        RaiseError "mMain.OpenDatabase", eGDRaiseError_Raise
    End If
    
End Function

Public Sub ToolbarSectorMenu(tbToolbar As SSActiveToolBars, ByVal strMenu$, _
    Optional aRetString As cGdArray = Nothing, Optional ByVal bPopulateMenu As Boolean = True, _
    Optional ByVal bAddDesc As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, nSymbolID&, strSymbol$, strTemp$
    Dim aIDs As New cGdArray, aStrings As cGdArray
    
    If aRetString Is Nothing Then
        Set aStrings = New cGdArray
    Else
        Set aStrings = aRetString
    End If

    'fill menu with sectors, subsectors, or components
    aIDs.Create eGDARRAY_Longs
    If Not ActiveChart Is Nothing Then
        nSymbolID = ActiveChart.Chart.SymbolID
        If nSymbolID > 0 Then
            strSymbol = GetSymbol(nSymbolID)
            If Left(strSymbol, 3) = "$--" Then
                ' for SECTOR:
                If strMenu = "ID_Sectors" Then
                    ' fill sectors menu with siblings
                    SU_GetGroupSiblings nSymbolID, aIDs
                ElseIf strMenu = "ID_Subsectors" Then
                    ' fill subsectors menu with children
                    SU_GetGroupChildren nSymbolID, aIDs
                End If
            ElseIf Left(strSymbol, 2) = "$-" Then
                ' for SUBSECTOR:
                If strMenu = "ID_Sectors" Then
                    ' fill sectors menu with siblings of parent
                    nSymbolID = SU_GetGroupParent(nSymbolID)
                    If nSymbolID > 0 Then
                        SU_GetGroupSiblings nSymbolID, aIDs
                    End If
                ElseIf strMenu = "ID_Subsectors" Then
                    ' fill subsectors menu with siblings
                    SU_GetGroupSiblings nSymbolID, aIDs
                Else
                    ' fill components menu with children
                    SU_GetGroupChildren nSymbolID, aIDs
                End If
            Else
                ' for SYMBOL:
                If strMenu = "ID_Sectors" Then
                    ' fill sectors menu with siblings of grandparent
                    nSymbolID = SU_GetGroupParent(nSymbolID)
                    If nSymbolID > 0 Then
                        nSymbolID = SU_GetGroupParent(nSymbolID)
                        If nSymbolID > 0 Then
                            SU_GetGroupSiblings nSymbolID, aIDs
                        End If
                    End If
                ElseIf strMenu = "ID_Subsectors" Then
                    ' fill subsectors menu with siblings of parent
                    nSymbolID = SU_GetGroupParent(nSymbolID)
                    If nSymbolID > 0 Then
                        SU_GetGroupSiblings nSymbolID, aIDs
                    End If
                Else
                    ' fill components menu with siblings
                    SU_GetGroupSiblings nSymbolID, aIDs
                End If
            End If
        End If
        If aIDs.Size = 0 And strMenu = "ID_Sectors" Then
            ' fill sectors menu with siblings of grandparent
            nSymbolID = SU_GetGroupParent("IBM")
            If nSymbolID > 0 Then
                nSymbolID = SU_GetGroupParent(nSymbolID)
                If nSymbolID > 0 Then
                    SU_GetGroupSiblings nSymbolID, aIDs
                End If
            End If
            nSymbolID = 0
        End If
    End If
    
    ' get list of symbols and sort
    aStrings.Clear
    If aIDs.Size > 0 Then
        aStrings.Size = aIDs.Size '(to preallocate space)
        aStrings.Size = 0
        For i = 0 To aIDs.Size - 1
            strTemp = GetSymbol(aIDs(i))
            ' ignore any expired symbols that may still be in the lists
            If Len(strTemp) > 0 And Left(strTemp, 1) <> "#" Then
                If bAddDesc Then
                    strTemp = strTemp & ": " & g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(strTemp))
                End If
                aStrings.Add strTemp
            End If
        Next
    End If
    If aStrings.Size > 0 Then
        aStrings.Sort eGdSort_DeleteNullValues
    ElseIf SecurityType(strSymbol) = "S" Then
        ' this stock must not be in a sector
        aStrings.Add "(" & strSymbol & " is not assigned to a sector)"
    Else
        ' if not a stock or if trying to show components for a sector
        aStrings.Add "(a stock or subsector must be charted)"
    End If
    
    If Not bPopulateMenu Then GoTo ErrExit
        
    If nSymbolID > 0 Then
        strSymbol = GetSymbol(nSymbolID)
    Else
        strSymbol = ""
    End If
        
    ' add menu item for each symbol
    With tbToolbar.Tools(strMenu).Menu
        For i = 1 To aStrings.Size
            If i > .Tools.Count Then
                strTemp = "Symbol #" & CStr(i) & "," & strMenu
                .Tools.Add strTemp, ssTypeStateButton
                tbToolbar.Tools(strTemp).Group = "Symbol"
                tbToolbar.Tools(strTemp).PictureDown = Picture16(ToolbarIcon("kChecked"))
            End If
            With .Tools(i)
                strTemp = aStrings(i - 1) & ":  " & g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(aStrings(i - 1)))
                strTemp = Trim(Replace(strTemp, "&", "&&"))
                If Right(strTemp, 1) = ":" Then
                    strTemp = Trim(Left(strTemp, Len(strTemp) - 1))
                End If
                .Name = strTemp
                If aStrings(i - 1) = strSymbol Then
                    .State = ssChecked
                Else
                    .State = ssUnchecked
                End If
            End With
        Next
        'remove extras
        For i = .Tools.Count To aStrings.Size + 1 Step -1
            .Tools.Remove i
        Next
    End With

ErrExit:
    Set aStrings = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ToolbarSectorMenu", eGDRaiseError_Raise
End Sub

Public Sub HideAnnotations(ByVal bHide As Boolean)
On Error GoTo ErrSection:

    If bHide Then
        g.ChartGlobals.nHideAnnotations = 1
        If Len(g.strActiveDraw) > 0 Then
            'toggle off any active drawing tool
            ToolbarSetCursorGroup frmMain.tbToolbar, False
            g.strActiveDraw = ""
            If Not ActiveChart Is Nothing Then
                ActiveChart.Chart.SetCursor
            End If
        End If
    Else
        g.ChartGlobals.nHideAnnotations = 0
    End If
    
    UpdateVisibleCharts eRedo1_Scrolled

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.HideAnnotations", eGDRaiseError_Raise
End Sub


' Returns DateTime of bar converted to the specified time zone:
' - "" (empty string) for machine's local time zone
' - or "GMT" for GMT/UTC, "NY" for New York, "CHI" for Chicago
' - or custom time zone info string (see "ConvertTimeZone" for format spec)
Public Function DateTimeConvert(PrimaryBars As cGdBars, ByVal dDateTimeOrBarsOffset#, _
                                Optional ByVal bArgIsBarsOffset As Boolean = False) As Double
On Error GoTo ErrSection:
    
    If PrimaryBars.IsIntraday And g.bShowInLocalTimeZone Then
        If bArgIsBarsOffset Then
            DateTimeConvert = gdBarsDateTimeConvert(PrimaryBars.BarsHandle, Int(dDateTimeOrBarsOffset), "")
        Else
            DateTimeConvert = ConvertTimeZone(dDateTimeOrBarsOffset, PrimaryBars.Prop(eBARS_ExchangeTimeZoneInf), "")
        End If
    ElseIf bArgIsBarsOffset Then
        DateTimeConvert = PrimaryBars(eBARS_DateTime, Int(dDateTimeOrBarsOffset))
    Else
        DateTimeConvert = dDateTimeOrBarsOffset
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.DateTimeConvert", eGDRaiseError_Raise
End Function

' returns true if expression is one of the standard defaults for a bars array parm
Public Function ParmDefaultIsBarsArray(ByVal strExpr$) As Boolean
On Error GoTo ErrSection:

    If InStr(UCase("|Close|High|Low|Open|AvgHL|AvgHLC|AvgOHLC|AvgOC|WClose|"), UCase("|" & strExpr & "|")) > 0 Then
        ParmDefaultIsBarsArray = True
    End If
    
ErrExit:
    Exit Function
        
ErrSection:
    RaiseError "mMain.ParmDefaultIsBarsArray", eGDRaiseError_Raise
    Resume ErrExit
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InstallData
'' Description: Allow the user to install data from a download or from a CD
'' Inputs:      None
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InstallData() As Boolean
On Error GoTo ErrSection:

' As a "backdoor": this file needs to be recently created in order
' to allow installing an older (minutized-style) data set
'If Int(FileDate(App.Path & "\OldInstall.flg")) < Date - 1 Then
    
    ' NEW "FULL TICK" METHOD
    InstallData = frmDataInstall2.ShowMe

'Else
'    ' OLD METHOD
'    Dim frmActive As Form               ' Active chart
'    Dim lDateOfCDReg As Long            ' Date of the installed CD in registry
'    Dim lDateOfCdDrv As Long            ' Date of CD in the drive
'    Dim lDateOfCD As Long               ' Date of the installed CD
'    Dim strCDInfoReg As String          ' Info for the installed CD
'    Dim strReturn As String             ' Return from in an infbox
'    Dim astrFile As New cGdArray        ' New Way file for configuration
'    Dim lStartRange As Long             ' Starting number of days
'    Dim lEndRange As Long               ' Ending number of days
'    Dim strCdDrive As String            ' CD Drive where the Genesis CD is
'    Dim strStarter As String            ' Starter data file
'    Dim strInfo As String               ' Starter data information
'
'    strCDInfoReg = GetCDDataInf
'    lDateOfCDReg = Int(Val(Parse(strCDInfoReg, vbTab, 2)))
'    astrFile.FromFile App.Path & "\Provided\Install.CFG"
'    lStartRange = CLng(Val(Parse(astrFile(0), "-", 1)))
'    lEndRange = CLng(Val(Parse(astrFile(0), "-", 2)))
'    If lStartRange = 0 Then lStartRange = 60
'    If lEndRange = 0 Then lEndRange = 180
'
'    lDateOfCD = lDateOfCDReg
'    strCdDrive = GenesisCDInDrive
'    If Len(strCdDrive) > 0 Then
'        lDateOfCdDrv = Int(FileDate(strCdDrive & ":\Data\Starter.GZP"))
'    End If
'
'    If lDateOfCdDrv <> 0& Then
'        If (Len(strCDInfoReg) = 0) Or (lDateOfCdDrv > lDateOfCDReg) Then
'            lDateOfCD = lDateOfCdDrv
'        End If
'    End If
'
'    If (Date <= (lDateOfCD + lEndRange)) And (Date >= (lDateOfCD + lStartRange)) Then
'        strReturn = InfBox("Would you like to install from the CD or download a new dataset?", "?", "+CD|-Download", "Data Install")
'    ElseIf (Date > (lDateOfCD + lEndRange)) Then
'        strReturn = "D"
'    Else
'        strReturn = "C"
'    End If
'
'    ' install data
'    If strReturn = "C" Then
'        ' from a CD
'        StartupLog "Data installation"
'        ShowForm frmDataInstall, True
'        StartupLog "Data installed"
'    Else
'        ' download starter set
'        If DownloadDataSet = False Then
'            InfBox ""
'            InstallData = False
'            Exit Function
'        Else
'            strStarter = AddSlash(DataPath) & "Symbols.DBF"
'            strInfo = strStarter & vbTab & Str(Val(FileDate(strStarter)))
'            FileFromString AddSlash(DataPath) & "Starter.INF", strInfo
'
'            DM_Close g.DMS
'            DM_Init True
'            g.Universe.OpenDb
'            g.SymbolPool.Load False
'            frmSymbolGrid.InitForm
'
'            Set frmActive = ActiveChart
'            If frmActive Is Nothing Then
'                Set frmActive = New frmChart        'new chart is always non-detached
'                frmActive.Chart.SetSymbol g.SymbolPool.SymbolIDforSymbol("$DJIA")
'                frmActive.WindowState = 2
'                frmActive.Show
'            End If
'
'            InfBox ""
'
'            frmSymbolGrid.RefreshGrid '.fg.Refresh
'            UpdateVisibleCharts
'            frmSymbolGrid.ShowInitialSymbol
'
'            frmQuotes.LoadTable
'            frmQuotes.TotalRefresh True
'        End If
'    End If
'
'    InstallData = True
'End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.InstallData", eGDRaiseError_Raise
    
End Function

' Returns max # of symbols allowed based on authorization codes
' (pass False for # of quote board symbols, True for # of depth-of-market symbols)
Public Function MaxSymbolsAllowed(Optional ByVal bMarketDepth As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim iMax&, i&, s$
    
    ' default (if no override codes)
    If bMarketDepth Then
        iMax = 1
    ElseIf HasGold(False) Or HasModule("RTGGOLD") Or HasModule("RTGPLAT") Then
        iMax = 250
    Else
        iMax = 25
    End If
    
    ' find largest Max# code (in case more than 1 happens to exist)
    s = Replace(g.strAuthorizationString, ",", ";")
's = ";SYM#500;15" & s & "DOM#4;SYM#1000;"
    Do While True
        If bMarketDepth Then
            i = InStr(s, ";DOM#")
        Else
            i = InStr(s, ";SYM#")
        End If
        If i < 1 Then Exit Do
        s = Mid(s, i + 5)
        i = Val(s)
        If i > iMax Then iMax = i
    Loop
    
If IsIDE Then
    iMax = 99999
End If
        
ErrExit:
    MaxSymbolsAllowed = iMax
    Exit Function
    
ErrSection:
    RaiseError "mMain.MaxSymbolsAllowed", eGDRaiseError_Raise
End Function

' Return the symbol ID for the color link of the specified form (i.e. looking at other forms)
Public Function FindWindowLinkSymbolID(frmCalledFrom As Form) As Long

    Dim iForm&, nColor&, nForColor&
    Dim frm As Form

    On Error Resume Next
    nForColor = frmCalledFrom.WindowLink.SymbolColor
    If nForColor > 0 Then
        For iForm = 0 To Forms.Count - 1
            Set frm = Forms(iForm)
            If Not frm Is frmCalledFrom Then
                nColor = 0
                nColor = frm.WindowLink.SymbolColor
                If nColor = nForColor Then
                    FindWindowLinkSymbolID = frm.SymbolID
                    Exit For
                End If
            End If
        Next
    End If
    Set frm = Nothing

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisCDInDrive
'' Description: Is there a Genesis CD in the drive?
'' Inputs:      None
'' Returns:     Drive the CD is in
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisCDInDrive(Optional bNewInstall As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strFile$, strDrive$, strDrives$, i&, iType&
    
    If bNewInstall Then
        strFile = ":\DataInst\DataInst.CFG"
    Else
        strFile = ":\Data\Starter.GZP"
    End If
    GenesisCDInDrive = ""
    
#If 0 Then
    Dim fs As New FileSystemObject      ' File system object
    Dim drv As Drive                    ' Drive object
    
    For Each drv In fs.Drives
        If Len(GenesisCDInDrive) = 0 Then
            strDrive = Left(UCase(Trim(drv.DriveLetter)), 1)
            If drv.DriveType = CDRom Or drv.DriveType = Removable Or strDrive = "C" Then
                If drv.IsReady Then
                    If FileExist(strDrive & strFile) Then
                        GenesisCDInDrive = strDrive
                        Exit For
                    End If
                End If
            End If
        End If
    Next drv
    Set drv = Nothing
    Set fs = Nothing
#Else
    
    strDrives = GetAllDrives
    For i = 1 To Len(strDrives)
        strDrive = Mid(strDrives, i, 1)
        iType = GetDriveType(strDrive & ":")
        If iType = DRIVE_REMOVABLE Or iType = DRIVE_CDROM Then ' Or strDrive = "C" Then
            If GetDiskSize(strDrive & ":") > 0 Then
                If FileExist(strDrive & strFile) Then
                    GenesisCDInDrive = strDrive
                    Exit For
                End If
            End If
        End If
    Next
#End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.GenesisCDInDrive", eGDRaiseError_Raise
End Function

' to remove full path from data table refs for tables in the Data folder
' (so the "Genesis" folder can more easily be moved to a different drive/location)
Private Sub FixMasterDataTablePaths()
On Error GoTo ErrSection:

    Dim nTable&, nField&, strDataPath$, strAppPath$, strText$, nLen&

    strDataPath = UCase(AddSlash(DataPath))
    strAppPath = UCase(AddSlash(App.Path))
    nLen = Len(strAppPath)
    
    If Left(strDataPath, nLen) = strAppPath Then
        If FileExist(strDataPath & "Master.GDM") Then
            nTable = TblOpen(strDataPath & "Master.GDM", True)
            If nTable <> 0 Then
                If d4top(nTable) = r4success Then
                    nField = d4field(nTable, "FilePath")
                    If nField <> 0 Then
                        Do While Not d4eof(nTable)
                            strText = f4memoStr(nField)
                            If Left(UCase(strText), nLen) = strAppPath Then
                                ' replace with relative path (e.g. ".\Data\Fut_Eod")
                                strText = "." & Mid(strText, nLen)
                                f4memoAssign nField, strText
                            End If
                            If d4skip(nTable, 1) <> r4success Then Exit Do
                        Loop
                    End If
                End If
                TblClose nTable
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.FixMasterDataTablePaths", eGDRaiseError_Raise
End Sub

Public Sub CopySettingsToOtherCharts(frmFrom As Form)
On Error GoTo ErrSection:

    Dim i&, strTemp$, strFrom$, strTo$
    Dim frm As Form

    If frmFrom Is Nothing Then Exit Sub

    strTemp = "Copy the chart settings from the| active chart to all other open charts?"
    If InfBox(strTemp, "?", "+Yes|-Cancel", "Copy Chart Settings") <> "Y" Then
        Exit Sub
    End If

    ' save this template and create temporary copy in templates area
    frmFrom.Chart.TemplateSave
    strTemp = "_CopyToOtherCharts_.HID"
    strTo = g.ChartGlobals.strCPCRoot & "\Charts\Templates\" & strTemp
    strFrom = g.ChartGlobals.strCPCRoot & "\Charts\" & frmFrom.Chart.Template & ".CHT"
    FileCopy strFrom, strTo
    
    ' copy template to all other charts
    For i = 0 To Forms.Count - 1
        If IsFrmChart(Forms(i)) Then
            Set frm = Forms(i)
            If Not frm Is frmFrom Then
                frm.Chart.TemplateApply strTemp, True
            End If
        End If
    Next
    Set frm = Nothing

ErrExit:
    If Len(strTo) > 0 Then KillFile strTo
    Exit Sub
    
ErrSection:
    RaiseError "mMain.CopySettingsToOtherCharts"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RestoreLibDatabase
'' Description: Attempt to restore a backup of the Libraries.MDB database
'' Inputs:      Restore Attempt
'' Returns:     True if restore, False if End
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RestoreLibDatabase(lRestoreAttempt As Long) As Boolean
On Error GoTo ErrSection:

    Dim bBackupExists As Boolean        ' Does the backup exist?
    Dim strBackup As String             ' Backup database to attempt to restore
    Dim strAttempt As String            ' "first" or "second"
    Dim bReturn As Boolean              ' Return value for the function

    Select Case lRestoreAttempt
        Case 1
            ' Backup the corrupted database...
            KillFile AddSlash(App.Path) & "Backup\Corrupted.MDB"
            FileCopy AddSlash(App.Path) & "Libraries.MDB", AddSlash(App.Path) & "Backup\Corrupted.MDB", True
                                
            ' Try to restore one of the backup databases...
            If FileExist(AddSlash(App.Path) & "LibBak1.MDB") Then
                strAttempt = "first"
                strBackup = AddSlash(App.Path) & "LibBak1.MDB"
                bBackupExists = True
            ElseIf FileExist(AddSlash(App.Path) & "LibBak2.MDB") Then
                lRestoreAttempt = lRestoreAttempt + 1
                strAttempt = "second"
                strBackup = AddSlash(App.Path) & "LibBak2.MDB"
                bBackupExists = True
            Else
                bBackupExists = False
            End If
            
        Case 2
            If FileExist(AddSlash(App.Path) & "LibBak2.MDB") Then
                strAttempt = "second"
                strBackup = AddSlash(App.Path) & "LibBak2.MDB"
                bBackupExists = True
            Else
                bBackupExists = False
            End If
        
        Case Is > 2
            bBackupExists = False
    
    End Select
    
    If bBackupExists Then
        InfBox "The libraries database is corrupted --|Restoring " & strAttempt & " backup.", "!", , "Database Error"
        FileCopy strBackup, AddSlash(App.Path) & "Libraries.MDB", True
        bReturn = True
    Else
        If InfBox("The libraries database is corrupted|and there are no good backups.||Would you like to get a new copy|of the default libraries database?|", "!", "+Yes|-Abort", "Error") = "A" Then
            bReturn = False
        Else
            KillFile AddSlash(App.Path) & "Libraries.MDB"
            bReturn = True
        End If
    End If
    
    RestoreLibDatabase = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.RestoreLibDatabase"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckLibIndexes
'' Description: Check to see if any indexes exist in the Libraries.MDB
'' Inputs:      None
'' Returns:     True if indexes exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckLibIndexes(Optional ByVal strWhen$ = "") As Boolean
On Error Resume Next

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (g.dbNav.TableDefs("tblSystems").Indexes.Count > 0) And (g.dbNav.TableDefs("tblRules").Indexes.Count > 0) And (g.dbNav.TableDefs("tblFunctions").Indexes.Count > 0) Then
        bReturn = True
    Else
        strWhen = "CheckLibIndexes failed: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & " " & strWhen
        FileFromString App.Path & "\CheckMDB.log", strWhen, True, True
        If IsIDE Then InfBox "*** Corrupt Libraries.MDB? ***", "e", , "IDE Message"
    End If
    
    CheckLibIndexes = bReturn

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RestoreTtDatabase
'' Description: Attempt to restore a backup of the TradeTracker.MDB database
'' Inputs:      Restore Attempt
'' Returns:     True if restore, False if End
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RestoreTtDatabase(lRestoreAttempt As Long) As Boolean
On Error GoTo ErrSection:

    Dim bBackupExists As Boolean        ' Does the backup exist?
    Dim strBackup As String             ' Backup database to attempt to restore
    Dim strAttempt As String            ' "first" or "second"
    Dim bReturn As Boolean              ' Return value for the function

    Select Case lRestoreAttempt
        Case 1
            ' Backup the corrupted database...
            KillFile AddSlash(App.Path) & "Backup\CorruptedTt.MDB"
            FileCopy AddSlash(App.Path) & "TradeTracker.MDB", AddSlash(App.Path) & "Backup\CorruptedTt.MDB", True
                                
            ' Try to restore one of the backup databases...
            If FileExist(AddSlash(App.Path) & "TTBak1.MDB") Then
                strAttempt = "first"
                strBackup = AddSlash(App.Path) & "TTBak1.MDB"
                bBackupExists = True
            ElseIf FileExist(AddSlash(App.Path) & "TTBak2.MDB") Then
                lRestoreAttempt = lRestoreAttempt + 1
                strAttempt = "second"
                strBackup = AddSlash(App.Path) & "TTBak2.MDB"
                bBackupExists = True
            Else
                bBackupExists = False
            End If
            
        Case 2
            If FileExist(AddSlash(App.Path) & "TTBak2.MDB") Then
                strAttempt = "second"
                strBackup = AddSlash(App.Path) & "TTBak2.MDB"
                bBackupExists = True
            Else
                bBackupExists = False
            End If
        
        Case Is > 2
            bBackupExists = False
    
    End Select
    
    If bBackupExists Then
        InfBox "The trade tracker database is corrupted --|Restoring " & strAttempt & " backup.", "!", , "Database Error"
        FileCopy strBackup, AddSlash(App.Path) & "TradeTracker.MDB", True
        bReturn = True
    Else
        If InfBox("The trade tracker database is corrupted|and there are no good backups.||Would you like to get a new copy|of the default trade tracker database?|", "!", "+Yes|-Abort", "Error") = "A" Then
            bReturn = False
        Else
            KillFile AddSlash(App.Path) & "TradeTracker.MDB"
            bReturn = True
        End If
    End If

    RestoreTtDatabase = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.RestoreTtDatabase"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckTtIndexes
'' Description: Check to see if any indexes exist in the TradeTracker.MDB
'' Inputs:      None
'' Returns:     True if indexes exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckTtIndexes(Optional ByVal strWhen$ = "") As Boolean
On Error Resume Next

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (g.dbPaper.TableDefs("tblAccounts").Indexes.Count > 0) And (g.dbPaper.TableDefs("tblOrders").Indexes.Count > 0) And (g.dbPaper.TableDefs("tblFills").Indexes.Count > 0) Then
        bReturn = True
    Else
        strWhen = "CheckTtIndexes failed: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & " " & strWhen
        FileFromString App.Path & "\CheckTtMDB.log", strWhen, True, True
        If IsIDE Then InfBox "*** Corrupt TradeTracker.MDB? ***", "e", , "IDE Message"
    End If
    
    CheckTtIndexes = bReturn

End Function

' Returns "Extreme Charts" mode (1 = Basic, 2 = Advanced)
Public Function ExtremeCharts() As Integer

    ExtremeCharts = m.iExtremeChartsMode

End Function

' Returns true if this is the Rule 1 University "Extreme" version
' TLB 2/19/2014: R1U no longer supported (now considered a Better Trades basic version)
Public Function IsRule1U() As Boolean

    'If m.iExtremeChartsMode = 2 Then
        'IsRule1U = True
    'End If

End Function

' Returns true if this is the PFG Best version
Public Function IsPfgVersion(Optional ByVal bAlsoCheckEnablement As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bHasEnablement As Boolean
    
    If bAlsoCheckEnablement Then
        bHasEnablement = HasModule("B_PFG")
    End If
    
    If bHasEnablement Or UCase(Left(GetSourceCode, 3)) = "PFG" Then
        If FileExist(App.Path & "\Info\BDNav.jpg") Then
            IsPfgVersion = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.IsPfgVersion"
End Function

' Returns true if this is the FxProbe version
Public Function IsFxProbeVersion() As Boolean
On Error GoTo ErrSection:

    If GetSourceCode = "FXPROBE" Then
        If FileExist(App.Path & "\Info\FxProbe.jpg") Then
            IsFxProbeVersion = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.IsFxProbeVersion"
End Function

' Returns true if this is the Learn:Forex version
Public Function IsLearnFxVersion() As Boolean
On Error GoTo ErrSection:

    'If GetSourceCode = "LEARNFX" Then
    '    If FileExist(App.Path & "\Info\LearnFX.jpg") Then
    '        IsLearnFxVersion = True
    '    End If
    'End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.IsLearnFxVersion"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsWoodiesVersion
'' Description: Is this the Woodies CCI Club version of the software?
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsWoodiesVersion() As Boolean
On Error GoTo ErrSection:

    IsWoodiesVersion = HasModule("WOODCCI") ' FileExist(AddSlash(App.Path) & "WoodiesClub.FLG")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.IsWoodiesVersion"
    
End Function

Public Sub SetupBrokerLayout(Optional ByVal bTest As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, strSymbol$, nSymbolID&, strName$, strChartPage$
    Dim frmC As frmChart            'JM: this seems to be used to create new charts, so leave form type as non-detached
    Dim frmT As frmTickDistribution
    Dim tRolls As cGdTable
    Dim Bars As New cGdBars
    Dim aStrings As New cGdArray

    ' calling this here will also allow it to happen at startup and after every download
    CheckReqPages

    If GetIniFileProperty("BrokerLayout", 0, "Brokers", g.strIniFile) = 0 Or bTest Then
        ' get default chart page (3rd line of Install.Cfg)
        aStrings.FromFile App.Path & "\Provided\Install.cfg"
        strChartPage = Trim(StripStr(aStrings(2), Chr(34)))
        aStrings.Size = 0
        
        If UCase(Left(GetSourceCode, 2)) = "FX" Then
            SetIniFileProperty "BrokerLayout", 1, "Brokers", g.strIniFile
            
            ' show trade console and hide symbolgrid
            DockState(frmSymbolGrid) = eHidden
            DockState(frmTTSummary) = eHidden
            
            ' show quote board -- display grid-style tab and add some symbols to it
            strName = "Forex"
            frmQuotes.CreateTab strName, eGDQuoteStyle_Forex, "$EUR-USD,$USD-JPY,$USD-CHF,$GBP-USD,$AUD-USD,$USD-CAD,$EUR-JPY"
            DockState(frmQuotes) = eShowAsPrevious
            
            Sleep 0.1 '(TLB: need this here to allow resize events to complete properly - else can cause infinite loop)
            
            nSymbolID = GetSymbolID("$EUR-USD")
            If FileExist(g.ChartGlobals.strCPCRoot & "\Charts\Pages\" & strChartPage & ".gzp") Then
                LoadChartPage strChartPage, True
            ElseIf nSymbolID > 0 Then
                ' first make sure intraday data exists for this symbol
                DM_GetBars Bars, nSymbolID, "60 min", LastDailyDownload - 5
                If Bars.Size > 0 Then
                    ' add new chart with order bar turned on
                    Set frmC = New frmChart
                    frmC.Chart.SetSymbol nSymbolID
                    frmC.Chart.ChangeBarPeriod "5 min", False
                    frmC.Chart.ShowTrades = 2
                    frmC.WindowState = 2
                    frmC.Chart.GenerateChart
                    ShowForm frmC
                    Set frmC = Nothing
                End If
            End If
        
            Sleep 0.1 '(TLB: need this here to allow resize events to complete properly - else can cause infinite loop)
        
            ' default group for symbol selector
            strName = "FOREX.GRP"
            If FileExist(App.Path & "\Provided\" & strName) Then
                '[SymbolSelector]
                'SymbolGroup=GRP:ELECFUT.GRP
                SetIniFileProperty "SymbolGroup", "GRP:" & strName, "SymbolSelector", g.strIniFile
            End If
        
            ' remove some buttons from main toolbar
            strName = App.Path & "\Toolbar.sho"
            aStrings.FromFile strName
            For i = aStrings.Size - 1 To 0 Step -1
                Select Case UCase(aStrings(i))
                Case "ID_SNAPSHOT", "ID_CHAIN"
                    aStrings.Remove i
                End Select
            Next
            aStrings.ToFile strName
            ToolbarReset True
        
        ElseIf g.Broker.IsBrokerUser(eTT_AccountType_TransAct) Or IsPfgVersion Then
            SetIniFileProperty "BrokerLayout", 1, "Brokers", g.strIniFile
            
            ' show trade console and hide symbolgrid
            DockState(frmSymbolGrid) = eHidden
            DockState(frmTTSummary) = eShowAsPrevious
            
            ' show quote board -- display box-style tab and add some symbols to it
            If InStr(UCase(GetSourceCode), "-CBOT") > 0 Then
                frmQuotes.CreateTab "CBOT", eGDQuoteStyle_OHLC, "$DJIA,YM-067,ZB-067,ZW-067,ZC-067,ZG-067,ZN-067,ZF-067,ZS-067,ZI-067", _
                    "SymbolID;-1|SecType;-1|Symbol;0|Period;-1|Session|Last Tick|Open|High|Low|Last|T|Bid|Bid Size|Ask|Ask Size|Volume|Prev Close|Change"
            Else
                frmQuotes.CreateTab "Elect Futures", eGDQuoteStyle_OHLC, "$DJIA,ES-067,TF-067,NQ-067,YM-067,ZB-067,ZN-067,G6E-067", _
                    "SymbolID;-1|SecType;-1|Symbol;0|Period;-1|Session|Last Tick|Open|High|Low|Last|T|Bid|Bid Size|Ask|Ask Size|Volume|Prev Close|Change"
            End If
            DockState(frmQuotes) = eShowAsPrevious
            
            Sleep 0.1 '(TLB: need this here to allow resize events to complete properly - else can cause infinite loop)
            
            ' create chart and price ladder for front-month ES
            If InStr(UCase(GetSourceCode), "-CBOT") > 0 Then
                Set tRolls = GetRollsTable("YM-067")
            Else
                Set tRolls = GetRollsTable("ES-067")
            End If
            If tRolls.NumRecords > 0 Then
                nSymbolID = tRolls.Num(0, tRolls.NumRecords - 1)
            End If
            Set tRolls = Nothing
            If FileExist(g.ChartGlobals.strCPCRoot & "\Charts\Pages\" & strChartPage & ".gzp") Then
                LoadChartPage strChartPage, True
            ElseIf nSymbolID > 0 Then
                ' first make sure intraday data exists for this symbol
                DM_GetBars Bars, nSymbolID, "60 min", LastDailyDownload - 5
                If Bars.Size > 0 Then
                    ' add new chart with order bar turned on
                    Set frmC = New frmChart
                    frmC.Chart.SetSymbol nSymbolID
                    frmC.Chart.ChangeBarPeriod "5 min", False
                    frmC.Chart.ShowTrades = 2
                    frmC.WindowState = 2
                    frmC.Chart.GenerateChart
                    ShowForm frmC
                    Set frmC = Nothing
                    
                    ' price ladder with order bar turned on
                    SetIniFileProperty "OrderBar", 1, "Price Ladder", g.strIniFile
                    'Set frmT = New frmTickDistribution
                    'frmT.ShowMe nSymbolID, 0, True
                    Set frmT = Nothing
                End If
            End If
            
            Sleep 0.1 '(TLB: need this here to allow resize events to complete properly - else can cause infinite loop)
        
            ' default group for symbol selector
            strName = "ELECFUT.GRP"
            If FileExist(App.Path & "\Provided\" & strName) Then
                '[SymbolSelector]
                'SymbolGroup=GRP:ELECFUT.GRP
                SetIniFileProperty "SymbolGroup", "GRP:" & strName, "SymbolSelector", g.strIniFile
            End If
        
            ' remove some buttons from main toolbar
            strName = App.Path & "\Toolbar.sho"
            aStrings.FromFile strName
            For i = aStrings.Size - 1 To 0 Step -1
                Select Case UCase(aStrings(i))
                Case "ID_SNAPSHOT", "ID_CHAIN"
                    aStrings.Remove i
                End Select
            Next
            aStrings.ToFile strName
            ToolbarReset True
        ElseIf HasModule("TSU", True) Then
            ' TradeSmart University layout
            SetIniFileProperty "BrokerLayout", 1, "Brokers", g.strIniFile
            
            ' show symbol grid and hide quote board
            DockState(frmQuotes) = eHidden
            DockState(frmTTSummary) = eHidden
            DockState(frmSymbolGrid) = eDocked
            
            Sleep 0.1 '(TLB: need this here to allow resize events to complete properly - else can cause infinite loop)
            
            If FileExist(g.ChartGlobals.strCPCRoot & "\Charts\Pages\" & strChartPage & ".gzp") Then
                LoadChartPage strChartPage, True
                Sleep 0.1 '(TLB: need this here to allow resize events to complete properly - else can cause infinite loop)
            End If
            
            ToolbarReset True
        
        ElseIf g.lLCD > 0 Then
            ' if a good connection has been made, check if need to load a default chart page
            SetIniFileProperty "BrokerLayout", -1, "Brokers", g.strIniFile
            If FileExist(g.ChartGlobals.strCPCRoot & "\Charts\Pages\" & strChartPage & ".gzp") Then
                LoadChartPage strChartPage, True
            End If
        End If
    
        ' TLB 6/20/2011: use default Templates list for specific source code
        strName = g.ChartGlobals.strCPCRoot & "\Charts\Req\Templates-" & GetSourceCode & ".LST"
        If FileExist(strName) Then
            FileCopy strName, g.ChartGlobals.strCPCRoot & "\Charts\Templates\Templates.LST"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.SetupBrokerLayout"
End Sub


Public Function LiveTradingAllowed(AcctType As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim strText$, bAllowed As Boolean

    If HasModule("BRKRLIVE") Then
        bAllowed = True
    Else
        If g.Broker.IsLiveAccount(AcctType) = False Then
            bAllowed = True '(shouldn't really even get here for these account types)
        ElseIf AcctType = eTT_AccountType_TransAct Then
            If Not g.Transact Is Nothing Then
                If UCase(g.Transact.UserName) = UCase(g.Transact.SimUserUserName) Or Len(g.Transact.UserName) = 0 Then
                    bAllowed = True
                End If
            End If
        End If
    End If
    
    If Not bAllowed Then
        ' they need to sign the liability release
        strText = "To enable live trading, you will need to sign the risk disclosure agreement.  Would you like to complete an 'online agreement' now?"
        If InfBox(strText, "?", "+Continue|-Not Now", "Live Trading Setup") <> "N" Then
            ' ?s=serviceid&p=password&agree=BRKRLIVE
            strText = "https://www.TradeNavigator.com/orderwizard/LogIn.aspx?s=*&p=*&agree=BRKRLIVE"
            strText = FixURL(GetProvidedProperty("AgreeBrokerLive", strText))
            RunProcess InternetBrowser, strText
            
            ' connect to refresh enablements
            strText = "The online agreement will be displayed in your browser.  After completing the agreement, a connection with Genesis is required in order to validate your new account enablements."
            Do While InfBox(strText, "i", "+Validate|-Abort", "Live Trading Setup") <> "A"
                GetNYTime
                If frmStatus.Status <> eStatus_Error Then frmStatus.Hide
                Sleep 1
                If HasModule("BRKRLIVE") Then
                    bAllowed = True
                    Exit Do
                End If
            Loop
        End If
    End If

ErrExit:
    LiveTradingAllowed = bAllowed
    Exit Function
    
ErrSection:
    RaiseError "mMain.LiveTradingAllowed"
End Function

' To convert a Future symbol (with or without the contract) to one of the following:
' 0 = Primary (returns Pit if exists, else returns Electronic)
' 1 = Pit
' 2 = Electronic (any symbol not in SymbolMap.csv is assumed to be electronic-only)
' 3 = Synthetic (electronic day-session)
' 4 = Combined (pit + electronic)
Public Function ConvertFutureSymbol(ByVal strSymbol$, ByVal eToType As eFutureSymbolType) As String
On Error GoTo ErrSection:

    Dim i&, iRec&, iFld&, strBase$, strContract$
    Dim aStrings As cGdArray, aFields As cGdArray
    Static aLookup As cGdArray
    Static SymTable As cGdTable
    Static nPrevLastDailyDownload As Long

    If Not IsAlpha(strSymbol, 1) Then Exit Function
    
    ' just load the table once or after a new daily download
    If nPrevLastDailyDownload <> LastDailyDownload Or aLookup Is Nothing Then
        nPrevLastDailyDownload = LastDailyDownload
        
        ' read in symbol table from file and create the Lookup array
        ' e.g.: TQ,ZB,ZB1,US  (pit, electronic, synthetic, combined)
        Set SymTable = New cGdTable
        Set aLookup = New cGdArray
        Set aStrings = New cGdArray
        Set aFields = New cGdArray
        aStrings.FromFile App.Path & "\Info\SymbolMap.csv"
        SymTable.CreateField eGDARRAY_Strings, ePitSymbol
        SymTable.CreateField eGDARRAY_Strings, eElectronicSymbol
        SymTable.CreateField eGDARRAY_Strings, eSyntheticSymbol
        SymTable.CreateField eGDARRAY_Strings, eCombinedSymbol
        SymTable.NumRecords = aStrings.Size
        For iRec = 0 To aStrings.Size - 1
            aFields.SplitFields UCase(aStrings(iRec)), ","
            For iFld = ePitSymbol To eCombinedSymbol
                strBase = Trim(aFields(iFld - 1))
                If Len(strBase) > 0 Then
                    SymTable(iFld, iRec) = strBase
                    ' add "SYMBOL REC#" to the lookup array for each symbol
                    aLookup.Add strBase & vbTab & Str(iRec)
                End If
            Next
        Next
        Set aStrings = Nothing
        Set aFields = Nothing
        aLookup.Sort
    End If
    
    ' parse the contract off the symbol (if exists) -- will append back on later
    i = InStr(strSymbol, "-")
    If i > 0 Then
        strBase = Left(strSymbol, i - 1)
        strContract = Mid(strSymbol, i)
    Else
        strBase = strSymbol
        strContract = ""
    End If
    
    ' search for base symbol in the Lookup array
    If aLookup.BinarySearch(strBase & vbTab, i, eGdSort_MatchUsingSearchStringLength) Then
        ' get Rec# in table and convert the base symbol
        iRec = Val(Parse(aLookup(i), vbTab, 2))
        If eToType = ePrimarySymbol Then
            ' return Pit if exists, else return Electronic
            strBase = SymTable(ePitSymbol, iRec)
            If Len(strBase) = 0 Then
                strBase = SymTable(eElectronicSymbol, iRec)
            End If
        Else
            strBase = SymTable(eToType, iRec)
        End If
    ' else if not in table then base symbol must be an electronic-only symbol
    ElseIf eToType <> eElectronicSymbol And eToType <> ePrimarySymbol Then
        strBase = ""
    End If
    
    ' return converted symbol (append contract back on)
    If Len(strBase) > 0 Then
        ConvertFutureSymbol = strBase & strContract
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ConvertFutureSymbol"
End Function

' bConvertToSynthentic: True = return synthetic for the specified symbol,
'                       False = return regular symbol for this synthetic
Public Function ConvertSynthetic(ByVal strSymbol$, ByVal bConvertToSynthetic As Boolean) As String
On Error GoTo ErrSection:

    Dim i&, iEnd&, strSym2$, strSym3$, strSym4$, strSym5$, strContract$
    Dim aStrings As cGdArray
    Static strSynthetics As String
    
    If Not IsAlpha(strSymbol, 1) Then Exit Function
    
    If Len(strSynthetics) = 0 Then
        'File: TQ,ZB,ZB1,US
        'strSynthetic:  <tab>Elec,Synth<tab>Elec,Synth
        Set aStrings = New cGdArray
        aStrings.FromFile App.Path & "\Info\SymbolMap.csv"
        For i = 0 To aStrings.Size - 1
            strSym2 = Parse(aStrings(i), ",", 2)
            strSym3 = Parse(aStrings(i), ",", 3)
            strSym4 = Parse(aStrings(i), ",", 4)
            strSym5 = Parse(aStrings(i), ",", 5)
                        
            If Len(strSym2) > 0 And Len(strSym3) > 0 Then
                strSynthetics = strSynthetics & vbTab & UCase(strSym2) & "," & UCase(strSym3)
            ' TLB 12/18/2007: don't do this for pseudo-pit anymore (causes problems with QB refreshes)
            'ElseIf Len(strSym4) > 0 And Len(strSym5) > 0 Then
                'strSynthetics = strSynthetics & vbTab & UCase(strSym4) & "," & UCase(strSym5)
            End If
        Next
        strSynthetics = strSynthetics & vbTab
        Set aStrings = Nothing
    End If
    
    i = InStr(strSymbol, "-")
    If i > 0 Then
        strContract = Mid(strSymbol, i)
        strSymbol = Left(strSymbol, i - 1)
    End If
    
    If bConvertToSynthetic Then
        ' return synthetic for the given symbol
        i = InStr(strSynthetics, vbTab & UCase(Trim(strSymbol)) & ",")
        If i > 0 Then
            i = InStr(i + 1, strSynthetics, ",")
            If i > 0 Then
                i = i + 1
                iEnd = InStr(i, strSynthetics, vbTab)
                If iEnd > 0 Then
                    ConvertSynthetic = Trim(Mid(strSynthetics, i, iEnd - i))
                End If
            End If
        End If
    Else
        ' return regular symbol for the synthetic
        i = InStr(strSynthetics, "," & UCase(Trim(strSymbol)) & vbTab)
        If i > 0 Then
            iEnd = i - 1
            For i = iEnd - 1 To 1 Step -1
                If Mid(strSynthetics, i, 1) = vbTab Then
                    ConvertSynthetic = Trim(Mid(strSynthetics, i + 1, iEnd - i))
                    Exit For
                End If
            Next
        End If
    End If
    
    If Len(ConvertSynthetic) > 0 And Len(strContract) > 0 Then
        ConvertSynthetic = ConvertSynthetic & strContract
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ConvertSynthetic"
End Function

' Returns the "trade symbol" for a given symbol
' - converts a synthetic to the normal trading symbol (e.g. returns ES for ES1)
' - if a date is passed, will convert continuous to the front-month
' - returns same type as passed (ID if ID is passed, or Symbol if Symbol is passed)
Public Function ConvertToTradeSymbol(ByVal vSymbolOrSymbolID As Variant, Optional ByVal dDate# = 0) As Variant
On Error GoTo ErrSection:

    Dim strSymbol$, strNonSynth$
    
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    If Len(strSymbol) > 0 Then
        ' convert synthetic symbol to normal (e.g. ES1 to ES)
        strNonSynth = ConvertSynthetic(strSymbol, False)
        If Len(strNonSynth) > 0 Then
            strSymbol = strNonSynth
        End If
    
        ' convert continuous contract to the front-month for that date
        If InStr(strSymbol, "-0") > 0 And dDate > 0 Then
            strSymbol = RollSymbolForDate(strSymbol, dDate)
        End If
    End If
    
    ' return same type as was passed in (Symbol or SymbolID)
    If VarType(vSymbolOrSymbolID) = vbString Then
        ConvertToTradeSymbol = strSymbol
    Else
        ConvertToTradeSymbol = GetSymbolID(strSymbol)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ConvertToTradeSymbol"
End Function


Public Property Get ChartTimers() As Boolean
    ChartTimers = m.bChartTimers
End Property

Public Property Let ChartTimers(ByVal bEnabled As Boolean)
    Dim i&
    If m.bChartTimers <> bEnabled Then
        m.bChartTimers = bEnabled
        For i = 0 To Forms.Count - 1
            If IsFrmChart(Forms(i)) Then
                Forms(i).tmr.Enabled = bEnabled
            End If
        Next
    End If
End Property

' returns the Source Code embedded in the Install.Cfg file of the Provided folder
Public Function GetSourceCode() As String
On Error GoTo ErrSection:

    Dim i&, aStrings As cGdArray
    Static strSrcCode As String
    Static bAlreadyDone As Boolean
    
    If Not bAlreadyDone Then
        bAlreadyDone = True
        Set aStrings = New cGdArray
        aStrings.FromFile App.Path & "\Provided\Install.cfg"
        strSrcCode = UCase(Replace(aStrings(1), "?", "&"))
        i = InStr(strSrcCode, "&SOURCE=")
        If i > 0 Then
            strSrcCode = Mid(strSrcCode, i + 8)
            i = InStr(strSrcCode, "&")
            If i > 1 Then strSrcCode = Left(strSrcCode, i - 1)
            strSrcCode = Trim(strSrcCode)
        Else
            strSrcCode = ""
        End If
        Set aStrings = Nothing
    End If
    GetSourceCode = strSrcCode

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.GetSourceCode"
End Function

' Check the \Charts\Req\ folder for any chart pages which require a module
' (will move the page over to the Pages folder if module is enabled)
Private Sub CheckReqPages()
On Error GoTo ErrSection:

    Dim i&, strReq$, strPage$, strFrom$, strTo$
    Dim aStrings As New cGdArray
    
    aStrings.GetMatchingFiles g.ChartGlobals.strCPCRoot & "\Charts\Req\*.req"
    For i = 0 To aStrings.Size - 1
        strReq = UCase(Trim(FileToString(aStrings(i), , True)))
        If HasModule(strReq, True) Then
            strPage = FileBase(aStrings(i)) & ".gzp"
            strFrom = g.ChartGlobals.strCPCRoot & "\Charts\Req\" & strPage
            strTo = g.ChartGlobals.strCPCRoot & "\Charts\Pages\" & strPage
            If FileExist(strFrom) And Not FileExist(strTo) Then
                If FileLength(strFrom) < 200 Then
                    KillFile strFrom
                Else
                    Name strFrom As strTo
                End If
            End If
            KillFile aStrings(i)
        End If
    Next
    Set aStrings = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.CheckReqPages"
End Sub

' Calls a web page and returns the data passed back
Public Function GetWebPageData(ByVal strUrl$, Optional ByVal nTimeout& = 30) As String

'On Error GoTo ErrSection:
On Error Resume Next ' to avoid the popup error message if times out

    If Not g.bUnloading Then
        frmMain.INet.RequestTimeout = nTimeout
        GetWebPageData = frmMain.INet.OpenURL(strUrl, icString)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    GetWebPageData = ""
    RaiseError "mMain.GetWebPageData"
End Function

Public Function GetPredictionLabsData() As Boolean
On Error GoTo ErrSection:

    Dim i&, d#, dDate#, strText$, strAll$, strPath$, iCfg&, strCfg$
    Dim aRecords As New cGdArray, aCfg As New cGdArray

    'djia,russell2000,sp500,sp100,nasdaq100,nasdaq

    ' read config file
    strPath = App.Path & "\PredLabs\"
    aCfg.FromFile strPath & "PredLabs.cfg"
    If aCfg.Size = 0 Then
        KillFile strPath & "*.dat"
    Else
        ' get today's date in NY (backup to Friday if during weekend)
        dDate = Int(ConvertTimeZone(Now, "", "NY")) ' current date in NY
        Do While Not IsWeekday(dDate)
            dDate = dDate - 1
        Loop
        
        ' get URL (first line of CFG file) and append Cust ID
        aCfg(0) = Parse(aCfg(0), vbTab, 1) & Trim(Str(g.lLCD))
        ' get data for each symbol in config file
        For iCfg = 1 To aCfg.Size - 1
            If g.bUnloading Then Exit For
            strCfg = Trim(aCfg(iCfg))
            If Len(strCfg) > 0 Then
                strAll = ""
                strText = GetWebPageData(aCfg(0) & Parse(strCfg, vbTab, 1))
                aRecords.SplitFields strText, Chr(10)
                If aRecords.Size < 5 Then
                    FileFromString strPath & "$DJIA.err", strText
                Else
                    For i = 0 To aRecords.Size - 1
                        strText = aRecords(i)
                        d = Val(Parse(strText, " ", 1))
                        If d > 0 Then
                            GetPredictionLabsData = True
                            If i = 0 Then
                                ' first record of file is the previous day's close,
                                ' so replace it with --> 9:34am NY = Null
                                ' (second record of file is first real data point: 9:35am NY)
                                d = RoundToMinute(dDate + 9.5 / 24# + 4 / 1440#)
                                strText = "-999999"
                            Else
                                d = RoundToMinute(dDate + (d + 9.5) / 24# - 5 / 1440#) ' 9:30am + time offset - 5 minutes
                                strText = Parse(strText, " ", 2)
                            End If
                            strAll = strAll & Format(d, "#0.0000000000") & vbTab & strText & vbCrLf
                        End If
                    Next
                End If
            
                ' save data in files to be read by the Get Text Data function
                strCfg = Parse(strCfg, vbTab, 2)
                aRecords.SplitFields strCfg, ","
                For i = 0 To aRecords.Size - 1
                    FileFromString strPath & aRecords(i) & ".dat", strAll
                Next
            End If
        Next
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.GetPredictionLabsData"
End Function

' for futures, lookup the primary base symbol
' (the first symbol in record of SymbolMap.csv file)
Public Function PrimaryFutureBase(ByVal strSymbol$) As String
On Error GoTo ErrSection:

    Dim i&, strBase$
    Static strSymbolMap As String
       
    ' first time: load symbol map in order to lookup the primary base symbol
    If Len(strSymbolMap) = 0 Then
        strSymbolMap = UCase(StripStr(FileToString(App.Path & "\Info\SymbolMap.csv"), " "))
        strSymbolMap = Replace(vbCrLf & strSymbolMap & vbCrLf, vbCrLf, "," & vbTab & ",")
        strSymbolMap = Replace(strSymbolMap, ",,,,", ",")
        strSymbolMap = Replace(strSymbolMap, ",,,", ",")
        strSymbolMap = Replace(strSymbolMap, ",,", ",")
        strSymbolMap = Replace(strSymbolMap, vbTab & ",", vbTab)
    End If
        
    ' lookup first symbol in record
    strBase = UCase(Parse(strSymbol, "-", 1))
    i = InStr(strSymbolMap, "," & strBase & ",")
    If i > 0 Then
        For i = i - 1 To 1 Step -1
            If Mid(strSymbolMap, i, 1) = vbTab Then
                strBase = Mid(strSymbolMap, i + 1)
                i = InStr(strBase, ",")
                If i > 0 Then
                    strBase = Left(strBase, i - 1)
                End If
                Exit For
            End If
        Next
    End If
    i = InStr(strSymbol, "-")
    If i > 0 Then
        strBase = strBase & Mid(strSymbol, i)
    End If
    
    PrimaryFutureBase = strBase

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.PrimaryFutureBase"
End Function

' for futures, lookup the Category for a symbol (in record of SymbolMap.csv file)
' returns first letter of category (or empty if no category assigned):
' C - Currencies (green)
' E - Energies (red)
' G - Grains (yellow)
' I - Equity Indexes (blue)
' L - Meats / Livestock (brown)
' M - Metals (silver)
' S - Softs (pink)
' T - Treasuries / Interest Rates (purple)
Public Function GetFuturesCategory(ByVal strSymbol$) As String
On Error GoTo ErrSection:

    Dim i&, strBase$, strCat$
    Static strSymbolMap As String
       
    ' first time: load symbol map in order to lookup the primary base symbol
    If Len(strSymbolMap) = 0 Then
        strSymbolMap = UCase(StripStr(FileToString(App.Path & "\Info\SymbolMap.csv"), " "))
        strSymbolMap = Replace(vbCrLf & strSymbolMap & vbCrLf, vbCrLf, "," & vbTab & ",")
        strSymbolMap = Replace(strSymbolMap, ",,,,", ",")
        strSymbolMap = Replace(strSymbolMap, ",,,", ",")
        strSymbolMap = Replace(strSymbolMap, ",,", ",")
        'strSymbolMap = Replace(strSymbolMap, vbTab & ",", vbTab)
    End If
        
    ' lookup symbol
    strBase = UCase(Parse(strSymbol, "-", 1))
    i = InStr(strSymbolMap, "," & strBase & ",")
    If i > 0 Then
        For i = i + 1 To Len(strSymbolMap)
            If Mid(strSymbolMap, i, 1) = vbTab Then
                Exit For
            ElseIf Mid(strSymbolMap, i, 1) = "/" Then
                strCat = UCase(Mid(strSymbolMap, i + 1, 1))
                Exit For
            End If
        Next
    End If
    
    GetFuturesCategory = strCat

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.GetFuturesCategory"
End Function

' for futures, lookup the Category Color for a symbol (in record of SymbolMap.csv file)
' returns RGB Color (or White if no category assigned):
' C - Currencies (green)
' E - Energies (red)
' G - Grains (yellow)
' I - Equity Indexes (blue)
' L - Meats / Livestock (brown)
' M - Metals (silver)
' S - Softs (pink)
' T - Treasuries / Interest Rates (purple)
Public Function GetFuturesCategoryColor(ByVal strSymbol$) As Long
On Error GoTo ErrSection:

    Dim strCat$, nColor&
        
    ' lookup category
    If Left(strSymbol, 1) = "/" Then
        strCat = Right(strSymbol, 1)
    Else
        strCat = GetFuturesCategory(strSymbol)
    End If
    
    Select Case UCase(strCat)
    Case "C" ' Currencies
        nColor = &H80FF80    ' green
    Case "E" ' Energies
        nColor = &H8080FF    ' red
    Case "G" ' Grains
        nColor = &H80FFFF    ' yellow
    Case "I" ' Indexes (equities)
        nColor = &HFFFFC0    ' light blue
    Case "L" ' Livestock (meats)
        nColor = &H80C0FF    ' brown
    Case "M" ' Metals
        nColor = &HC0C0C0    ' silver
    Case "S" ' Softs
        nColor = &HFFC0FF    ' pink
    Case "T" ' Treasuries
        nColor = &HFFC0C0    ' purple
    Case Else ' unassigned
        nColor = RGB(255, 255, 255) ' white
    End Select
    
    GetFuturesCategoryColor = nColor

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.GetFuturesCategoryColor"
End Function

Public Sub SaveVisibleForms()
On Error GoTo ErrSection:

    Dim i&, s$
    Dim aStrings As New cGdArray
    Dim frmTD As frmTickDistribution
    Dim frmTS As frmTimeSales
    Dim frmMP As frmMarketProfile
    
    For i = 0 To Forms.Count - 1
        s = ""
        If TypeOf Forms(i) Is frmTickDistribution Then
            'TD FormPlacement SymbolLink PeriodLink SymbolID ViewMode
            Set frmTD = Forms(i)
            s = "TD" & vbTab & GetFormPlacement(frmTD) & vbTab & Str(frmTD.WindowLink.SymbolColor) _
                & vbTab & Str(frmTD.WindowLink.PeriodColor) & vbTab & Str(frmTD.SymbolID) & vbTab & Str(frmTD.DisplayStyle)
            Set frmTD = Nothing
        ElseIf TypeOf Forms(i) Is frmTimeSales Then
            'TS FormPlacement SymbolLink PeriodLink SymbolID
            Set frmTS = Forms(i)
            s = "TS" & vbTab & GetFormPlacement(frmTS) & vbTab & Str(frmTS.WindowLink.SymbolColor) _
                & vbTab & Str(frmTS.WindowLink.PeriodColor) & vbTab & Str(frmTS.SymbolID)
            Set frmTS = Nothing
        ElseIf TypeOf Forms(i) Is frmMarketProfile Then
            'MP FormPlacement SymbolLink PeriodLink SymbolID
            Set frmMP = Forms(i)
            s = "MP" & vbTab & GetFormPlacement(frmMP) & vbTab & Str(frmMP.WindowLink.SymbolColor) _
                & vbTab & Str(frmMP.WindowLink.PeriodColor) & vbTab & Str(frmMP.SymbolID)
            Set frmMP = Nothing
        End If
        If Len(s) > 0 Then aStrings.Add s
    Next
    
    If aStrings.Size > 0 Then
        aStrings.ToFile App.Path & "\VisibleForms.dat"
    Else
        KillFile App.Path & "\VisibleForms.dat"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.SaveVisibleForms"
End Sub

Public Sub RestoreVisibleForms()
On Error GoTo ErrSection:

    Dim i&
    Dim aStrings As New cGdArray, aFlds As New cGdArray
    Dim frmTD As frmTickDistribution
    Dim frmTS As frmTimeSales
    Dim frmMP As frmMarketProfile
    
    aStrings.FromFile App.Path & "\VisibleForms.dat"
    For i = 0 To aStrings.Size - 1
        aFlds.SplitFields aStrings(i)
        Select Case UCase(aFlds(0))
        Case "TD"
            'TD FormPlacement SymbolLink PeriodLink SymbolID ViewMode
            ' only restore if in ladder view
            If Val(aFlds(5)) = 0 Then
                Set frmTD = New frmTickDistribution
                Load frmTD
                SetFormPlacement frmTD, aFlds(1)
                frmTD.ShowMe Val(aFlds(4)), Val(aFlds(5)), True
                frmTD.WindowLink.SymbolColor = Val(aFlds(2))
                Set frmTD = Nothing
            End If
        Case "TS"
            'TS FormPlacement SymbolLink PeriodLink SymbolID ViewMode
            Set frmTS = New frmTimeSales
            Load frmTS
            SetFormPlacement frmTS, aFlds(1)
            frmTS.ShowMe Val(aFlds(4))
            frmTS.WindowLink.SymbolColor = Val(aFlds(2))
            Set frmTS = Nothing
'        Case "MP"
'            'TS FormPlacement SymbolLink PeriodLink SymbolID ViewMode
'            'JM 10-31-2012: we are not going to restore these files because of the # of ticks
'            Set frmMP = New frmMarketProfile
'            Load frmMP
'            SetFormPlacement frmMP, aFlds(1)
'            frmMP.ShowMe Val(aFlds(4))
'            frmMP.WindowLink.SymbolColor = Val(aFlds(2))
'            Set frmMP = Nothing
        End Select
    Next
    
    MoveFocus ActiveChart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.RestoreVisibleForms"
End Sub

Public Sub ExportData(Optional ByVal bRtExport As Boolean = False)
On Error GoTo ErrSection:

    Dim i&
    Dim aStrings As New cGdArray
    Dim ExportGroup As New cExportGroup

    aStrings.FromFile App.Path & "\Custom\Export.TXT"
    For i = 0 To aStrings.Size - 1
        ExportGroup.FromString aStrings(i)
        If ExportGroup.AutoExport Then
            'frmStatus.AddDetail "Exporting " & ExportGroup.SymbolGroup & " (" & ExportGroup.Format & ")"
            ExportGroup.Export bRtExport
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ExportData"
End Sub

Public Function NewFullTick() As Boolean
On Error GoTo ErrSection:

    Static iNewDLL As Integer
    If iNewDLL = 0 Then
        'If FileDate(App.Path & "\DMDLL.DLL") >= DateSerial(2007, 7, 22) Then
        If App.Major >= 4 Then
            iNewDLL = 1
        Else
            iNewDLL = -1
        End If
    End If
    If iNewDLL > 0 Then NewFullTick = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.NewFullTick"
End Function

' Returns true if DotNet dependencies have been installed
' (will prompt user to download and install if not done yet)
Public Function HasDotNet(Optional ByVal bFromStartup As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strPath$, dTimeout#, strMsg$, bHasDotNet As Boolean

    strPath = AddSlash(FilePath(App.Path)) & "DotNetSetup\"
    If FileExist(strPath & "Setup.EXE") Then
        KillFile strPath & "Message.txt"
        ' check if dependencies need to be installed (if Setup.EXE newer than Setup.DON file)
        If FileDate(strPath & "Setup.EXE") <= FileDate(strPath & "Setup.DON") Then
            bHasDotNet = True
        ElseIf InfBox("We first need to verify if the Microsoft .NET dependencies have been installed.", "i", "+OK|-Cancel", "Microsoft .NET Setup") <> "C" Then
            KillFile strPath & "Setup.DON"
            ' run Setup (this Setup program is from a dummy .Net app just to check for the dependencies)
            RunProcess strPath & "Setup.EXE"
            ' see if the ".DON" file gets created within 10 seconds
            ' - if so, then there must have been no new dependencies, so return true
            ' - if not, then must be in the middle of installing new dependencies,
            '       so just return false for now and leave a message to be displayed
            dTimeout = gdTickCount + 10000
            Do While gdTickCount < dTimeout
                DoEvents
                If FileExist(strPath & "Setup.DON") Then
                    bHasDotNet = True
                    Exit Do
                End If
            Loop
            If Not bHasDotNet Then
                ' create the message to be displayed by our Dependency checker after dependencies have been installed
                If bFromStartup Then
                    strMsg = "The Microsoft .NET dependencies have now been installed."
                Else
                    strMsg = "The Microsoft .NET dependencies have now been installed.  Please retry your request again now."
                End If
                FileFromString strPath & "Message.txt", strMsg
            
                ' if from the Startup, then ask them to wait before continuing
                If bFromStartup Then
                    strMsg = "Please wait until the Microsoft .NET dependencies have been installed,|then hit 'Continue' ..."
                    If InfBox(strMsg, "i", "+Continue|-Quit", ".NET Framework dependency check") = "Q" Then
                        End
                    End If
                End If
            End If
        End If
    End If
    
    HasDotNet = bHasDotNet

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.HasDotNet"
End Function

Public Function HasNewNavEngine() As Boolean
On Error GoTo ErrSection:

    Static iHasNewNavEngine As Integer
    If iHasNewNavEngine = 0 Then
        iHasNewNavEngine = -1
        If FileDate(App.Path & "\NavEngineAdv.dll") >= DateSerial(2008, 5, 9) Then
            iHasNewNavEngine = 1
        End If
    End If
    If iHasNewNavEngine > 0 Then HasNewNavEngine = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.HasNewNavEngine"
End Function

' just for employees: copy Recalc.Log file over to network
Public Sub CopyRecalcLog()
    
    Dim strFile$, strText$
    On Error Resume Next
    If FileExist("c:\common\ask.exe") And g.lLCD > 0 Then
        If FileExist("k:\allusers\recalc\*.*") Then
            strFile = App.Path & "\Recalc.log"
            If g.nNumVerifiedRecalcs > 0 And g.bUnloading Then
                strText = vbCrLf & Format(Now, "yyyy-mm-dd hh:mm:ss, ") & "#Verified = " & Str(g.nNumVerifiedRecalcs)
                FileFromString strFile, strText, True, True
                g.nNumVerifiedRecalcs = 0
            End If
            strText = "k:\allusers\recalc\" & Str(RI_GetDataServiceID) & ".log"
            If FileDate(strFile) > FileDate(strText) + 0.0001 Then
                FileCopy strFile, strText
            End If
        End If
    End If

End Sub

Public Sub ToolbarSyncCursorGroup(tbToolbar As SSActiveToolBars, ByVal strToolID$)
On Error Resume Next:

    Dim frm As Form
    Dim frmParent As Form
    Dim i&
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    
    'must be a chart's or main app's tool bar
    If IsFrmChart(tbToolbar.Parent) Then
        Set frmParent = tbToolbar.Parent
    ElseIf Not TypeOf tbToolbar.Parent Is frmMain Then
        Exit Sub
    End If
    
    If strToolID = "ID_RepeatDraw" Or strToolID = "ID_Magnet" Or strToolID = "ID_DragModeY" Then
        For i = 0 To Forms.Count - 1
            If IsFrmChart(Forms(i)) Then
                Set frm = Forms(i)
                If frm.DetachStatus = eDetached Then
                    If Not frm Is frmParent Then
                        If frm.DetachStatus = eDetached Then
                            frm.tbToolbar.Tools(strToolID).State = tbToolbar.Tools(strToolID).State
                        End If
                    End If
                End If
            End If
        Next
        
        If Not frmParent Is Nothing Then
            frmMain.tbToolbar.Tools(strToolID).State = tbToolbar.Tools(strToolID).State
        End If
    End If

End Sub

Public Sub FixFocusChart(frmSource As Form, Optional Tool As ActiveToolBars.SSTool = Nothing, _
    Optional ByVal bFromActivateEvent As Boolean = False)
On Error Resume Next

    '04-17-2009
    'Fix for 4883 in cWindowLink.cls makes this not necessary.
    'Leave awhile for reference then remove if all okay.
    Exit Sub

End Sub

Public Sub CreateShortcuts()

    On Error Resume Next ' (since this may not work on all operating systems)
    Dim strAppPath$, strDesc$
    
    ' delete the old desktop links (from the really old install)
    KillFile SpecialFolderPath(CSIDL_ALLUSERS_DESKTOP) & "Navigator Suite.lnk"
    KillFile SpecialFolderPath(CSIDL_USER_DESKTOP) & "Navigator Suite.lnk"
    
    strAppPath = AddSlash(App.Path)
    strDesc = "To start the Trade Navigator platform."
    
    ' create some shortcuts in the Start Menu (under the Programs\Genesis folder)
    CreateShortcut strAppPath & "TNArchive.exe", "*\Genesis", "Backup and Restore Settings", "To backup and/or restore all of your custom settings."
    If ExtremeCharts = 0 Then
        CreateShortcut strAppPath & "NavSuite.exe", "*\Genesis", "Trade Navigator", strDesc
        CreateShortcut strAppPath & "Info\Remote\GRemote.exe", "*\Genesis", "Remote Assistance", "To allow a Technical Support specialist to connect to your machine."
        
        ' create a desktop shortcut for TradeNav
        CreateShortcut strAppPath & "NavSuite.exe", , "Trade Navigator", strDesc
    
        ' create a shortcut in the Quick Launch area
        'CreateShortcut strAppPath & "NavSuite.exe", SpecialFolderPath(CSIDL_QUICKLAUNCH), "Trade Navigator", strDesc
    End If

End Sub

Public Function AllowRemoveOvernightGap() As Boolean
On Error GoTo ErrSection:
    
    AllowRemoveOvernightGap = HasModule("SPYDER")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.AllowRemoveOvernightGap"
End Function

' To make sure other programs we have started have been killed
' (e.g. when we startup, or before upgrading)
Private Sub KillOtherPrograms()
On Error Resume Next
    
    If Not IsIDE Then
        If KillProcess("OptionNav") > 0 Then
            StartupLog "Option Navigator killed"
        End If
        If KillProcess("GenTransact") > 0 Then
            StartupLog "GenTransact killed"
        End If
        If KillProcess("GenPFG") > 0 Then
            StartupLog "GenPFG killed"
        End If
        If Not FileExist(App.Path & "\debug.rt") Then
            If KillProcess("GenesisRT") > 0 Then
                StartupLog "GenesisRT killed"
            End If
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BaseURL
'' Description: Determine the base URL to be used for our website
'' Inputs:      None
'' Returns:     Base Url
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BaseURL() As String
On Error GoTo ErrSection:

    BaseURL = mMain.GetProvidedProperty("BaseWeb", "www.TradeNavigator.com")
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.BaseURL"
    
End Function

' To decrypt a URL and/or embed the encrypted UserID and Password
Public Function FixURL(ByVal strUrl$) As String
On Error GoTo ErrSection:

    Dim i&, a&, iFld&
    Dim aURL As New cGdArray
    
    ' TLB 5/12/2011: allow for tab-delimited pieces (since one piece may be encrypted while another may not)
    aURL.SplitFields strUrl, vbTab
    For iFld = 0 To aURL.Size - 1
        strUrl = aURL(iFld)
        ' if an encrypted hex string, then first decrypt it
        For i = 1 To Len(strUrl)
            ' see if char is 0-9 or A-F
            a = Asc(UCase(Mid(strUrl, i, 1)))
            If (a >= 48 And a <= 57) Or (a >= 65 And a <= 70) Then
                If i = Len(strUrl) Then
                    ' it must be a Hex string -- so decrypt it
                    strUrl = DecryptFromHex(strUrl)
                End If
            Else
                Exit For ' it must not be a hex string
            End If
        Next
        
        ' insert encrypted Username and Password into URL
        i = InStr(UCase(strUrl), "U=*&P=*")
        If i = 0 Then
            i = InStr(UCase(strUrl), "S=*&P=*") ' old method
        End If
        If i > 0 Then
            strUrl = Left(strUrl, i) & "=" & EncryptToHex(RI_GetDataServiceID) _
                & "&P=" & EncryptToHex(RI_GetUserPassword) & Mid(strUrl, i + 7)
        End If
        
        ' insert Source Code into URL
        i = InStr(UCase(strUrl), "&S=*")
        If i = 0 Then
            i = InStr(UCase(strUrl), "?S=*")
        End If
        If i > 0 Then
            strUrl = Left(strUrl, i + 2) & GetSourceCode & Mid(strUrl, i + 4)
        End If
        
        ' insert Build# into URL
        i = InStr(UCase(strUrl), "&B=*")
        If i = 0 Then
            i = InStr(UCase(strUrl), "?B=*")
        End If
        If i > 0 Then
            strUrl = Left(strUrl, i + 2) & Str(App.Revision) & Mid(strUrl, i + 4)
            strUrl = Replace(strUrl, "&B=*", "") ' and just remove any extras
        End If
        
        ' TLB 11/6/2013: replace GenesisFT.com (too many issues with the old domain name)
        strUrl = Replace(strUrl, "GenesisFT.com", "TradeNavigator.com", , , vbTextCompare)
        
        ' DAJ 08/21/2015: If the base URL is something other than TradeNavigator.com, replace it here...
        If InStr(UCase(strUrl), "TRADENAVIGATOR.COM") > 0 Then
            If InStr(UCase(strUrl), "WWW.TRADENAVIGATOR.COM") = 0 And InStr(UCase(strUrl), ".TRADENAVIGATOR.COM") = 0 Then
                strUrl = Replace(strUrl, "TradeNavigator.com", "www.TradeNavigator.com")
            End If
            strUrl = Replace(strUrl, "www.TradeNavigator.com", BaseURL, , , vbTextCompare)
        End If
        
        aURL(iFld) = strUrl
    Next
    strUrl = aURL.JoinFields("")
    
    If FileExist("c:\common\files.exe") Then
        DebugLog strUrl
    End If

ErrExit:
    FixURL = strUrl
    Exit Function
    
ErrSection:
    RaiseError "mMain.FixURL"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UrlEncodeField
'' Description: Replace certain characters in the string with their hex equivalent
'' Inputs:      URL, Replace Space With
'' Returns:     Encoded string
''
'' Note:        Can only run this on each field, not the entire URL string because
''              it replaces characters such as '&' and '=' which would screw up
''              the URL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UrlEncodeField(ByVal strUrlField As String, Optional ByVal strReplaceSpaceWith As String = "%20") As String
On Error GoTo ErrSection:
    
    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim cChr As String                  ' Character out of the original string
    Dim iAsc As Integer                 ' ASCII representation of the character
    
    strReturn = ""
    For lIndex = 1 To Len(strUrlField)
        cChr = Mid(strUrlField, lIndex, 1)
        iAsc = Asc(cChr)
        
        If iAsc = vbKeySpace Then
            strReturn = strReturn & strReplaceSpaceWith
        Else
            Select Case iAsc
                Case Is <= 47
                    strReturn = strReturn & "%" & Right("0" & Hex(iAsc), 2)
                Case Is <= 57
                    strReturn = strReturn & cChr
                Case Is <= 64
                    strReturn = strReturn & "%" & Right("0" & Hex(iAsc), 2)
                Case Is <= 90
                    strReturn = strReturn & cChr
                Case Is <= 96
                    strReturn = strReturn & "%" & Right("0" & Hex(iAsc), 2)
                Case Is <= 122
                    strReturn = strReturn & cChr
                Case Is <= 255
                    strReturn = strReturn & "%" & Right("0" & Hex(iAsc), 2)
            End Select
        End If
    Next lIndex
    
    UrlEncodeField = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.UrlEncodeField"
    
End Function

' To display a web report in a separate process (so no chance of crashing TradeNav)
Public Function RunWebReport(ByVal strTitle$, ByVal strUrl$, Optional ByVal strIcon$ = "", Optional ByVal nDefaultWindowState& = vbMaximized) As Boolean
On Error GoTo ErrSection:

    Dim strProgram$, strTemp$, strMsg$
    Dim iWebID&, iMaxSpecificID&
    
    If Len(strUrl) = 0 Then Exit Function
    
    strProgram = App.Path & "\TNWebReport.EXE" ' "\WebShell.EXE"
    If FileExist(strProgram) Then
        'Parameters - tab delimited:
        '1. Window Title
        '2. URL
        '3. form icon
        '4. default window state (if first time for this title)
        '5. WebReportAppMailID
        
        ' for specific contexts (per the strIcon), only allow 1 instance for each
        iMaxSpecificID = 5
        If g.nLastWebReportID < iMaxSpecificID Then
            g.nLastWebReportID = iMaxSpecificID ' need to init this to the max specific context
        End If
        Select Case strIcon
        Case "kSectorAnalysis"
            iWebID = 1
        Case "kNewsBrowser"
            iWebID = 2
        Case "kSeasonalSP"
            iWebID = 3
        Case "kScreenerWeb"
            iWebID = 4
        Case "kSharedChartPage"
            iWebID = 5
            strMsg = "REFRESH" ' to force the web page to refresh if already running
        Case Else
            g.nLastWebReportID = g.nLastWebReportID + 1
            iWebID = g.nLastWebReportID
        End Select
        
        ' for those specific contexts, see if it's already running
        If iWebID <= iMaxSpecificID Then
            ' try to send a message to it to tell it to restore itself
            If frmMain.apmNews.CreateMessage("TNWebReport" & Str(iWebID), 2, strMsg, , True) <> 0 Then
                ' if the message is sent successfully, then we don't want to start another instance
                iWebID = 0
                RunWebReport = True
            End If
        End If
        
        ' if need to start an instance
        If iWebID <> 0 Then
            strTemp = FixURL(strUrl)
            ' TLB 5/14/2013: convert "Seas" (unlimited) to "Seasonal" (limited) if not enabled
            If Right(strTemp, 7) = "&T=Seas" Then
                If Not HasModule("MOSWT") Then
                    strTemp = strTemp & "onal"
                End If
            End If
            strTemp = strTitle & vbTab & strTemp & vbTab & strIcon & vbTab & Str(nDefaultWindowState) & vbTab & Str(iWebID)
            RunWebReport = RunProcess(strProgram, strTemp)
        End If
    Else
        InfBox "The program " & strProgram & " was not found.", , , "Run Web Report"
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.RunWebReport"
End Function

Public Function GetProvidedProperty(ByVal strPropName$, Optional ByVal vDefaultValue As Variant = "", _
                    Optional ByVal bCompanySpecific As Boolean = False) As Variant
On Error GoTo ErrSection:

    Dim strSection$
    
    If Not bCompanySpecific Then
        strSection = "General"
    ElseIf IsRule1U Then
        strSection = "Rule1U"
    ElseIf ExtremeCharts >= 1 Then
        strSection = "BetterTrades"
    Else
        strSection = "GenesisFT"
    End If

    GetProvidedProperty = GetIniFileProperty(strPropName, vDefaultValue, strSection, App.Path & "\Provided\Provided.INI")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.GetProvidedProperty"
End Function

' return the symbol to use for the "settle"
' (for determining what a symbol's net change is for the day)
Public Function GetSettleSymbol(ByVal strSymbol$) As String
On Error GoTo ErrSection:
                
    Dim i&, strSettleSymbol$
    
    ' see if the symbol is a future (and not an option)
    i = InStr(strSymbol, "-")
    If i > 0 And g.bUsePitSettlesForDeltas Then
        If Left(strSymbol, 1) <> "$" And InStr(strSymbol, " ") = 0 Then
            ' get the pit symbol
            strSettleSymbol = ConvertFutureSymbol(strSymbol, ePitSymbol)
            If Len(strSettleSymbol) > 0 Then
                strSymbol = strSettleSymbol
            Else
                ' but if no pit symbol, then get the synthetic (since the "settle"
                ' needs to come from the daily close of the last non-holiday)
                strSettleSymbol = ConvertFutureSymbol(strSymbol, eSyntheticSymbol)
                If Len(strSettleSymbol) > 0 Then
                    strSymbol = strSettleSymbol
                End If
            End If
        End If
    End If

ErrExit:
    GetSettleSymbol = strSymbol
    Exit Function
    
ErrSection:
    RaiseError "mMain.GetSettleSymbol"
End Function

' return the "settle": close of the settle symbol for the prior session
Public Function GetPrevCloseForQB(Bars As cGdBars) As Double
On Error GoTo ErrSection:
    
    Dim i&, iLastBar&, nSymbolID&, dPrevClose#, nCurrentSession&
    Dim strSymbol$, strSettleSymbol$, strNew$, strOld$
    Dim Settles As cGdBars
    
    dPrevClose = kNullData
    For iLastBar = Bars.Size - 1 To 0 Step -1
        If Bars(eBARS_Close, iLastBar) <> kNullData Then
            Exit For
        End If
    Next
    nSymbolID = Bars.Prop(eBARS_SymbolID)
    If iLastBar > 0 And nSymbolID > 0 Then
        ' for quote board: use pit settle if a future using daily bars
        If g.bUsePitSettlesForDeltas And Bars.SecurityType = "F" _
                    And Bars.Prop(eBARS_Periodicity) = ePRD_Days + 1 Then
            ' need to store SymbolID, CurrentSession, PrevClose
            ' (if first 2 haven't changed then we can just use the same PrevClose)
            nCurrentSession = Bars.SessionDate(iLastBar)
            strOld = Bars.Prop(eBARS_CustomString)
            strNew = Str(nSymbolID) & vbTab & Str(nCurrentSession) & vbTab
            i = Len(strNew)
            If strNew = Left(strOld, i) Then
                dPrevClose = Val(Mid(strOld, i + 1))
            Else
                ' get the symbol to use for settles and load the last 7 days
                strSymbol = GetSettleSymbol(Bars.Prop(eBARS_Symbol))
                Set Settles = New cGdBars
                DM_GetBars Settles, strSymbol, "Daily", nCurrentSession - 7
                If Settles.Size > 0 Then
                    ''g.RealTime.AddTickBuffer Settles
                    'If Settles.SessionDate(Settles.Size - 1) < nCurrentSession - 1 Then
                        ' need to load data since last daily download
                        g.RealTime.SpliceBars Settles
                    'End If
                    ' search for session settle prior to the current session
                    For i = Settles.Size - 1 To 0 Step -1
                        If Settles.SessionDate(i) < nCurrentSession Then
                            dPrevClose = Settles(eBARS_Close, i)
                            Exit For
                        End If
                    Next
                End If
                Bars.Prop(eBARS_CustomString) = strNew & Str(dPrevClose)
                Set Settles = Nothing
            End If
        End If
        If dPrevClose = kNullData Then
            dPrevClose = Bars(eBARS_Close, iLastBar - 1)
        End If
    End If

ErrExit:
    GetPrevCloseForQB = dPrevClose
    Exit Function

ErrSection:
    RaiseError "mMain.GetPrevCloseForQB"
End Function

' Returns a unique custom filename that includes the MachineID and a random "number"
' (in the format _MACHINEID_ABCDE.ext -- with whatever extension is passed in)
Public Function GetUniqueCustomFilename(ByVal strExt$) As String
On Error GoTo ErrSection:

    Dim i&, r&, strChars$, strFile$, strMachID$, bFound As Boolean
       
    strExt = UCase(Trim(strExt))
    If Left(strExt, 1) <> "." Then
        strExt = "." & strExt
    End If
    strMachID = StripStr(UCase(RI_GetMachineID), "- ")
    
    strChars = "123456789ABCDEFGHJKMNPQRSTUVWXYZ"
    Do
        ' build a filename using MachineID and 5 random characters
        ' (32 ^ 5 = over 32 million possibilities)
        strFile = ""
        For i = 1 To 5
            r = RandomNum(1, Len(strChars))
            strFile = strFile & Mid(strChars, r, 1)
        Next
        strFile = UCase("_" & strMachID & "_" & strFile & strExt)
        
        bFound = False
        If FileExist(App.Path & "\Custom\" & strFile) Then
            bFound = True
        ElseIf FileExist(App.Path & "\Provided\" & strFile) Then
            bFound = True
        End If
    Loop While bFound

    GetUniqueCustomFilename = strFile
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mMain.GetUniqueCustomFilename"
End Function

' returns RGB color from either a single value or comma-delimited RGB values
' (returns -1 if string is empty)
Public Function GetColorFromString(ByVal strColor$) As Long
    
    Dim nColor&, nRed&, nGreen&, nBlue&
    
    On Error Resume Next
    strColor = Parse(strColor, ";", 1)
    If Len(strColor) = 0 Then
        nColor = -1
    ElseIf InStr(strColor, ",") = 0 Then
        nColor = Val(strColor)
    Else
        nRed = Val(Parse(strColor, ",", 1))
        If nRed < 0 Or nRed > 256 Then
            nColor = nRed
        Else
            nGreen = Val(Parse(strColor, ",", 2))
            nBlue = Val(Parse(strColor, ",", 3))
            nColor = RGB(nRed, nGreen, nBlue)
        End If
    End If
    
    GetColorFromString = nColor
    
End Function

' Gets the default gradient color from the main INI file (or hard-coded if missing).
' MAIN POINT: we want to be able to change/tweak our default gradient color whenever we want,
'   so it is stored in the Provided.INI file, and whenever the user chooses our default gradient
'   color we want to store that as -1 so will change to our newer default color when we change.
' NOTE: we only want to get this once at startup since we always want to store each chart's gradient
'   color as -1 if it still matches the current default gradient color -- but if the default changes
'   mid-stream (e.g. after a daily download), then it would not match the default anymore and not get
'   stored as -1.  So we just wait until the next startup to get the new default gradient from the INI file.
Public Function GradientDefault() As Long

    On Error Resume Next
    Static nColor As Long
    
    If nColor = 0 Then
        nColor = GetColorFromString(GetProvidedProperty("GradientDefaultColor", ""))
        If nColor <= 0 Then
            'nColor = RGB(255, 236, 184) 'gold
            'nColor = RGB(168, 208, 232) 'blue
            nColor = RGB(192, 192, 192) 'silver = 12632256
        End If
    End If
    GradientDefault = nColor

End Function

' We are now able to set the App's BackColor to a custom color
' (except a few controls such as command buttons and combo boxes).
' So for now, we will just be setting it a slight bit darker (or lighter)
' than the command buttons in order to help them stand out better
Private Sub SetTheAppBackColor()

    On Error Resume Next
    Dim nColor&, nRed&, nGreen&, nBlue&, nAvg&, nAdj&, dwFlags&, i&
    Dim strColor$, strFile$
    Dim bWhiteForeColor As Boolean
    Dim eStyle As FormBorderStyleConstants
    
    nColor = g.nColorTheme
    
    If nColor > 0 Then
        If nColor = 1 Or nColor = kDarkThemeColor Then
            bWhiteForeColor = True
        End If
    Else
        ' get ButtonFace color
        nColor = GetSysColor(15)
        If nColor <> 0 Then
            ' if the ButtonFace color is a very light shade, then make the
            ' app background a little darker, else make it a little lighter
            nRed = nColor Mod 256
            nGreen = Int(nColor / 256) Mod 256
            nBlue = Int(nColor / 65536)
            nAvg = (nRed + nGreen + nBlue) / 3
            If nAvg > 232 Then
                nAdj = -16 ' darker than the buttons
            ElseIf nAvg > 212 Then
                nAdj = -12 ' a little darker than the buttons
            Else
                nAdj = 14 ' lighter than the buttons
            End If
            'frmTest.AddList "RGB = " & Str(nRed) & " " & Str(nGreen) & " " & Str(nBlue) & ", Avg = " & Str(nAvg) & ", Adj = " & Str(nAdj)
            nRed = nRed + nAdj
            nGreen = nGreen + nAdj
            nBlue = nBlue + nAdj
            If nRed > 255 Then nRed = 255
            If nRed < 0 Then nRed = 0
            If nGreen > 255 Then nGreen = 255
            If nGreen < 0 Then nGreen = 0
            If nBlue > 255 Then nBlue = 255
            If nBlue < 0 Then nBlue = 0
            nColor = RGB(nRed, nGreen, nBlue)
        End If
    End If
    SetAppBackColor nColor, bWhiteForeColor
    g.nColorTheme = GetAppBackColor()       'JM 11-16-2015: need to call this to get right theme color for Classic theme
    If g.nColorTheme = kDarkThemeColor Then g.ConsoleForms.Refresh
End Sub

Private Sub StartNewsBrowser()
On Error GoTo ErrSection:

    Dim strExe$, strCfg$, strErrMsg$
    Dim aArgs As New cGdArray

    strExe = App.Path & "\News\NewsBrowser.exe"
    strCfg = App.Path & "\News\NewsBrowser.cfg"
    aArgs(0) = GetProvidedProperty("NewsServerIP")
    If Not FileExist(strExe) Then
        strErrMsg = "News Browser not found."
    ElseIf HasDotNet Then
        aArgs.ToFile strCfg
        RunProcess strExe, Chr(34) & strCfg & Chr(34)
    End If
    
    If Len(strErrMsg) > 0 Then
        InfBox strErrMsg, "e", , "Error"
    End If

ErrExit:
    Set aArgs = Nothing
    Exit Sub

ErrSection:
    RaiseError "mMain.StartNewsBrowser"
End Sub

'For both PRICE and TIME clusters:
'
'#1. We first need to identify where the swing points are and what weight to give them.
'- we will allow for 3 types of swing points: long-term, intermediate, and short-term
'- the user can select which of the above types of swing points to use along with the swing-point-strength for each type (defaults: long-term = 21 bars, intermediate = 14 bars, short-term = 7 bars)
'- to allow for the more major swing points to have more weight in the clustering, the user can also set the weight for each swing point type (default: long-term = 300, intermediate = 200, short-term = 100)
'- we should be able to display the swing point indicator for the lowest type selected (e.g. the short-term if turned on) and then label each swing point according to it's type (L = long-term, I = intermediate, S = short-term)
'- NOTE: this entire step should also be wrapped up into a "swing point" drawing tool so it can be easily added to a chart simply for display purposes (except for the 100/200/300 weightings)
'
'For PRICE clusters only:
'
'#2. Allow user to select the start-end -- here's some various options:
'- user clicks on the end point and specifies the # of bars back to use
'- user clicks on the end point and specifies the # of long-term swing points back to use
'- user clicks on the chart twice to specify both the start and end point
'
'#3. Identify which swing point combinations to use for the PRICE clusters.
'- user can select options for "Rallies" (Low->High) and/or "Declines" (High->Low) -- we will use every combination of high-low and/or low-high such that the 2 end points (A & B) are the highest and lowest of the prices (i.e. no higher or lower prices can exist between A & B)
'- user can also select "fib extensions" (like the 3-point ABC tool which uses 3 swing points) -- we will use every valid combination of ABC such that A & B are the highest and lowest of all prices between A and C
'- so in practical terms, here's the way we should probably code it: for each starting low swing point ("A") we will keep looking at high swing points ("B") in the future until the price drops lower than "A" (which ends our need to continue looking), and we will use each "B" only if it is the highest so far (but if it is lower than a previous high swing point, then it is ignored for this particular "A") -- and if including the fib extensions, then for each "B" that is used we will also look for all valid low swing points ("C") to use in the future (and we can stop looking as soon as the price goes above "B" or below "A")
'
'#4. For each swing point combination, calculate where each fib line would be and add to the "cluster totals"
'- user can select/edit/add which fib ratios to use along with a custom weighting for each ratio (ratio's from 0-1 would be for retracements, ratio's < 0 and > 1 would be for expansions)
'- user can select "retracements" (using ratios from 0-1), "expansions" (using ratios < 0 and > 1), and/or "extensions" (using all selected ratio's > 0 of the A-B difference from the C point)
'- for each fib line, calculate it's total weight = (swing A weight + swing B weight) * ratio weight
'- and use the "fib price window" (FPW) to add to the cluster totals for each possible price (i.e. min moves) -- so if the fib line is at price "P", then do a "For N = P-FPW To P+FPW" and add to each price: (FPW - abs(P-N)) * fib line weight
'
'#5. Display the price cluster totals at the right of the chart
'- display the cluster total for each possible price at the right as a histogram
'- the horiz scale for the cluster totals at the right should go from 0 (zero) to the Max of the highest cluster total for ALL the prices (i.e. the scale should not change as you move the chart up/down)
'
'PRICE CLUSTER Function:
'Inputs:
'   - Bars data (price bars), Start bar # and End bar # (window in which to look for swing points)
'   - Swing point array (i.e. result of "Swing Point Levels" function: 3=Long-term, 2=Interm, 1=Short-term, positive=High, negative=Low, 0=none)
'   - Swing point weightings (doubles array where item 0=0, 1=short-term weight, 2=intermediate weight, 3=long-term weight)
'   - Fib Ratios (doubles array: each fib ratio that is selected)
'   - Fib Ratio Weights (doubles array: the corresponding weight to use for each fib ratio selected)
'   - Tick Proximity: include # of ticks away from actual fib price (value decreases as get further away)
'   - Flags: support, resistance, retracements, expansions, extensions
'Output: cluster table sorted by price -- Column 0 = Price, Column 1 = Cluster value at that price
Public Function CalcPriceClusters(Bars As cGdBars, ByVal iStartBar&, ByVal iEndBar&, aSwingPoints As cGdArray, aSwingPointWeights As cGdArray, _
            aFibRatios As cGdArray, aFibRatioWeights As cGdArray, ByVal iTickProximity&, ByVal bSupport As Boolean, ByVal bResistance As Boolean, _
            ByVal bUseRetracements As Boolean, ByVal bUseExpansions As Boolean, ByVal bUseABCs As Boolean) As cGdTable
On Error GoTo ErrSection:

    Dim i&, j&, d#
    Dim iBarA&, iSwingA&, dPriceA#
    Dim iBarB&, iSwingB&, dPriceB#
    Dim iBarC&, iSwingC&, dPriceC#
    Dim iFib&, dPrice#, dValue#, dMinMove#, iWindow&, dRatio#
    Dim aClusterPrices As New cGdArray, aClusterValues As New cGdArray, aClusterSort As cGdArray
    Dim aClusters As New cGdTable
    
    aSwingPointWeights(0) = 0
    
    aClusterPrices.Create eGDARRAY_Doubles
    aClusterValues.Create eGDARRAY_Doubles
    aClusters.CreateField eGDARRAY_Doubles, 0, "Price"
    aClusters.CreateField eGDARRAY_Doubles, 1, "ClusterValue"
    dMinMove = Bars.MinMove

    ' start looking for each swing point
    For iBarA = iStartBar To iEndBar
        iSwingA = aSwingPoints.Num(iBarA)
        If aSwingPointWeights.Num(Abs(iSwingA)) <= 0 Then
            iSwingA = 0 ' ignore if the weighting for this level is 0
        End If
        If iSwingA <> 0 Then
            If iSwingA < 0 Then
                dPriceA = Bars(eBARS_Low, iBarA)
            Else
                dPriceA = Bars(eBARS_High, iBarA)
            End If
            dPriceB = dPriceA
            ' for each starting swing point ("A") we will keep looking at opposite swing points ("B")
            ' in the future (i.e. Low->High or High->Low)
            For iBarB = iBarA + 1 To iEndBar
                iSwingB = aSwingPoints.Num(iBarB)
                If aSwingPointWeights.Num(Abs(iSwingB)) <= 0 Then
                    iSwingB = 0
                End If
                If iSwingA < 0 Then ' if "A" is a Low:
                    ' look until the price drops lower than "A" (which ends our need to continue looking)
                    If Bars(eBARS_Low, iBarB) < dPriceA Then
                        Exit For
                    End If
                    ' and we will use each High "B" only if it is the highest so far
                    ' (otherwise ignore it for this particular "A")
                    If iSwingB <= 0 Then
                        iSwingB = 0 ' and ignore Low->Low
                    ElseIf Bars(eBARS_High, iBarB) <= dPriceB Then
                        iSwingB = 0
                    Else
                        dPriceB = Bars(eBARS_High, iBarB)
                    End If
                Else ' if "A" is a High:
                    ' look until the price goes higher than "A" (which ends our need to continue looking)
                    If Bars(eBARS_High, iBarB) > dPriceA Then
                        Exit For
                    End If
                    ' and we will use each Low "B" only if it is the lowest so far
                    ' (otherwise ignore it for this particular "A")
                    If iSwingB >= 0 Then
                        iSwingB = 0 ' and ignore High->High
                    ElseIf Bars(eBARS_Low, iBarB) >= dPriceB Then
                        iSwingB = 0
                    Else
                        dPriceB = Bars(eBARS_Low, iBarB)
                    End If
                End If
                
                If iSwingB <> 0 Then
                    ' now calc each Fib line for this A->B
                    For iFib = 0 To aFibRatios.Size - 1
                        dRatio = aFibRatios.Num(iFib)
                        If Not bSupport Then
                            ' Support = positive ratio's for Low->High, negative ratio's for High->Low
                            If dRatio > 0 And iSwingB > 0 Then
                                dRatio = 0
                            ElseIf dRatio < 0 And iSwingB < 0 Then
                                dRatio = 0
                            End If
                        End If
                        If Not bResistance Then
                            ' Resistance = positive ratio's for High->Low, negative ratio's for Low->High
                            If dRatio > 0 And iSwingB < 0 Then
                                dRatio = 0
                            ElseIf dRatio < 0 And iSwingB > 0 Then
                                dRatio = 0
                            End If
                        End If
                        
                        ' calculate the total fib line weight = ratio weight * (Avg weight for swings A and B)
                        dValue = aFibRatioWeights.Num(iFib)
                        If dValue > 0 And dRatio <> 0 Then
                            dValue = dValue * (aSwingPointWeights.Num(Abs(iSwingA)) + aSwingPointWeights.Num(Abs(iSwingB))) / 2
                            ' calculate the fib price = B - (B - A) * ratio
                            dPrice = dPriceB - (dPriceB - dPriceA) * dRatio
                            ' calculate value for prices based on proximity to the actual fib price
                            ' (i.e. assign decreasing values as move away from the actual fib price)
                            For iWindow = 0 To iTickProximity
                                If iWindow = 0 Then
                                    aClusterPrices.Add dPrice
                                    aClusterValues.Add dValue
                                Else
                                    d = dValue * (iTickProximity + 1 - iWindow) / (iTickProximity + 1)
                                    aClusterPrices.Add dPrice + iWindow * dMinMove
                                    aClusterValues.Add d
                                    aClusterPrices.Add dPrice - iWindow * dMinMove
                                    aClusterValues.Add d
                                End If
                            Next
                        End If
                    Next
                    
                    If bUseABCs Then
                        ' and if including the fib extensions, then for each "B" that is used we will also
                        ' look for all valid swing points ("C") to use in the future
                        For iBarC = iBarB + 1 To iEndBar
                            iSwingC = aSwingPoints.Num(iBarC)
                            If aSwingPointWeights.Num(Abs(iSwingC)) <= 0 Then
                                iSwingC = 0
                            End If
                            If iSwingA < 0 Then ' if "A" is a Low:
                                ' can stop looking as soon as the price goes below "A" or above "B"
                                If Bars(eBARS_Low, iBarC) < dPriceA Or Bars(eBARS_High, iBarC) > dPriceB Then
                                    Exit For
                                End If
                                ' "C" must also be a Low
                                If iSwingC >= 0 Then
                                    iSwingC = 0
                                Else
                                    dPriceC = Bars(eBARS_Low, iBarC)
                                End If
                            Else ' if "A" is a High:
                                ' can stop looking as soon as the price goes below "B" or above "A"
                                If Bars(eBARS_Low, iBarC) < dPriceB Or Bars(eBARS_High, iBarC) > dPriceA Then
                                    Exit For
                                End If
                                ' "C" must also be a High
                                If iSwingC <= 0 Then
                                    iSwingC = 0
                                Else
                                    dPriceC = Bars(eBARS_High, iBarC)
                                End If
                            End If
                            If iSwingC <> 0 Then
                                ' now calc all the Fib lines for this A->B->C
                                For iFib = 0 To aFibRatios.Size - 1
                                    dRatio = aFibRatios.Num(iFib)
                                    If Not bSupport Then
                                        ' Support = positive ratio's for High->Low->High, negative ratio's for Low->High->Low
                                        If dRatio > 0 And iSwingC > 0 Then
                                            dRatio = 0
                                        ElseIf dRatio < 0 And iSwingC < 0 Then
                                            dRatio = 0
                                        End If
                                    End If
                                    If Not bResistance Then
                                        ' Resistance = positive ratio's for Low->High->Low, negative ratio's for High->Low->High
                                        If dRatio > 0 And iSwingC < 0 Then
                                            dRatio = 0
                                        ElseIf dRatio < 0 And iSwingC > 0 Then
                                            dRatio = 0
                                        End If
                                    End If
                                    
                                    ' calculate the total fib line weight = ratio weight * (Avg weight for swings A, B and C)
                                    dValue = aFibRatioWeights.Num(iFib)
                                    If dValue > 0 And dRatio <> 0 Then
                                        dValue = dValue * (aSwingPointWeights.Num(Abs(iSwingA)) + aSwingPointWeights.Num(Abs(iSwingB)) + aSwingPointWeights.Num(Abs(iSwingC))) / 3
                                        ' calculate the fib price = C + (B - A) * ratio
                                        dPrice = dPriceC + (dPriceB - dPriceA) * dRatio
                                        ' calculate value for prices based on proximity to the actual fib price
                                        ' (i.e. assign decreasing values as move away from the actual fib price)
                                        For iWindow = 0 To iTickProximity
                                            If iWindow = 0 Then
                                                aClusterPrices.Add dPrice
                                                aClusterValues.Add dValue
                                            Else
                                                d = dValue * (iTickProximity + 1 - iWindow) / (iTickProximity + 1)
                                                aClusterPrices.Add dPrice + iWindow * dMinMove
                                                aClusterValues.Add d
                                                aClusterPrices.Add dPrice - iWindow * dMinMove
                                                aClusterValues.Add d
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            Next
        End If
    Next
    
    ' now sort by ClusterPrice and add into the Cluster table
    ' (sum the values for all fib lines which are at the same rounded price)
    If aClusterPrices.Size > 0 Then
        dPrice = kNullData
        Set aClusterSort = aClusterPrices.CreateSortedIndex
        For i = 0 To aClusterSort.Size - 1
            j = aClusterSort.Num(i)
            dValue = aClusterValues.Num(j)
            dPriceA = Bars.RoundToPrice(aClusterPrices.Num(j))
            If dPrice <> dPriceA Then
                ' if a new price, add a new record
                dPrice = dPriceA
                aClusters.NumRecords = aClusters.NumRecords + 1
                aClusters.Num(0, aClusters.NumRecords - 1) = dPrice
                aClusters.Num(1, aClusters.NumRecords - 1) = dValue
            Else
                ' if same price, just add into that record
                aClusters.Num(1, aClusters.NumRecords - 1) = aClusters.Num(1, aClusters.NumRecords - 1) + dValue
            End If
        Next
    End If

ErrExit:
    Set CalcPriceClusters = aClusters
    Exit Function
    
ErrSection:
    RaiseError "CalcPriceClusters"
    Resume ErrExit
End Function

'TIME CLUSTER Function:
'Inputs:
'   - Bars data (price bars), Start bar # and End bar # (window in which to look for swing points)
'   - Max bars between swing points
'   - Swing point array (i.e. result of "Swing Point Levels" function: 3=Long-term, 2=Interm, 1=Short-term, positive=High, negative=Low, 0=none)
'   - Swing point weightings (doubles array where item 0=0, 1=short-term weight, 2=intermediate weight, 3=long-term weight)
'   - Fib Ratios (doubles array: each fib ratio that is selected)
'   - Fib Ratio Weights (doubles array: the corresponding weight to use for each fib ratio selected)
'   - Bar Proximity: include # of bars away from actual fib bar (value decreases as get further away)
'   - Flags: L2H, H2L, L2L, H2H, extensions?
'Output: cluster array (cluster value for each bar#)
Public Function CalcTimeClusters(Bars As cGdBars, ByVal iStartBar&, ByVal iEndBar&, ByVal iMaxBars&, aSwingPoints As cGdArray, aSwingPointWeights As cGdArray, _
            aFibRatios As cGdArray, aFibRatioWeights As cGdArray, ByVal iBarProximity&, ByVal bUseL2H As Boolean, ByVal bUseH2L As Boolean, _
            ByVal bUseL2L As Boolean, ByVal bUseH2H As Boolean, ByVal bUseABCs As Boolean) As cGdArray
On Error GoTo ErrSection:

    Dim i&, j&, d#
    Dim iBarA&, iSwingA& ', dPriceA#
    Dim iBarB&, iSwingB& ', dPriceB#
    Dim iBarC&, iSwingC& ', dPriceC#
    Dim iFib&, iFibBar&, dValue#, iWindow&
    Dim aClusterBars As New cGdArray
    
    aSwingPointWeights(0) = 0
    
    aClusterBars.Create eGDARRAY_Doubles, Bars.Size, 0

'#3. Identify which swing point combinations to use for the TIME clusters.
' user can select options for High->Low, Low->High, High->High, Low->Low (any or all can be selected),
' and can 'select the max # of bars between swing points to be used together (i.e. between A & B)

    ' start looking for each swing point
    For iBarA = iStartBar To iEndBar
        iSwingA = aSwingPoints.Num(iBarA)
        If aSwingPointWeights.Num(Abs(iSwingA)) <= 0 Then
            iSwingA = 0 ' ignore if the weighting for this level is 0
        End If
        If (iSwingA < 0 And (bUseL2H Or bUseL2L)) Or (iSwingA > 0 And (bUseH2L Or bUseH2H)) Then
            ' for each starting swing point ("A") we will look at each swing point ("B") what is within MaxBars
            For iBarB = iBarA + 1 To iEndBar
                If iBarB - iBarA > iMaxBars Then
                    Exit For
                End If
                iSwingB = aSwingPoints.Num(iBarB)
                If aSwingPointWeights.Num(Abs(iSwingB)) <= 0 Then
                    iSwingB = 0
                End If
                If iSwingA < 0 Then ' if "A" is a Low:
                    If iSwingB < 0 And Not bUseL2L Then
                        iSwingB = 0
                    ElseIf iSwingB > 0 And Not bUseL2H Then
                        iSwingB = 0
                    End If
                Else ' if "A" is a High:
                    If iSwingB < 0 And Not bUseH2L Then
                        iSwingB = 0
                    ElseIf iSwingB > 0 And Not bUseH2H Then
                        iSwingB = 0
                    End If
                End If
                If iSwingB <> 0 Then
                    ' now calc each Fib line for this A->B
                    For iFib = 0 To aFibRatios.Size - 1
                        ' calculate the total fib line weight = ratio weight * (Avg weight for swings A and B)
                        dValue = aFibRatioWeights.Num(iFib)
                        If dValue > 0 Then
                            dValue = dValue * (aSwingPointWeights.Num(Abs(iSwingA)) + aSwingPointWeights.Num(Abs(iSwingB))) / 2
                            ' calculate the fib bar = B + (B - A) * ratio   (round to nearest bar)
                            iFibBar = Int(iBarB + (iBarB - iBarA) * aFibRatios.Num(iFib) + 0.5)
                            ' calculate value for bars based on proximity to the actual fib bar
                            ' (i.e. assign decreasing values as move away from the actual fib bar)
                            For iWindow = 0 To iBarProximity
                                If iWindow = 0 Then
                                    If iFibBar < aClusterBars.Size Then
                                        aClusterBars.Num(iFibBar) = aClusterBars.Num(iFibBar) + dValue
                                    End If
                                Else
                                    d = dValue * CDbl(iBarProximity + 1 - iWindow) / CDbl(iBarProximity + 1)
                                    If iFibBar + iWindow < aClusterBars.Size Then
                                        aClusterBars.Num(iFibBar + iWindow) = aClusterBars.Num(iFibBar + iWindow) + d
                                    End If
                                    If iFibBar - iWindow < aClusterBars.Size Then
                                        aClusterBars.Num(iFibBar - iWindow) = aClusterBars.Num(iFibBar - iWindow) + d
                                    End If
                                End If
                            Next
                        End If
                    Next
                    
#If 0 Then
                    If bUseABCs Then
                        ' and if including the fib extensions, then for each "B" that is used we will also
                        ' look for all valid swing points ("C") to use within MaxBars of "B"
                        For iBarC = iBarB + 1 To iEndBar
                            If iBarC - iBarB > iMaxBars Then
                                Exit For
                            End If
                            iSwingC = aSwingPoints.Num(iBarC)
                            If aSwingPointWeights.Num(Abs(iSwingC)) <= 0 Then
                                iSwingC = 0
                            End If
                            
                If iSwingA < 0 Then ' if "A" is a Low:
                    If iSwingC < 0 And Not bUseL2L Then
                        iSwingC = 0
                    ElseIf iSwingC > 0 And Not bUseL2H Then
                        iSwingC = 0
                    End If
                Else ' if "A" is a High:
                    If iSwingB < 0 And Not bUseH2L Then
                        iSwingB = 0
                    ElseIf iSwingB > 0 And Not bUseH2H Then
                        iSwingB = 0
                    End If
                End If
                            
                            
                            
                            If iSwingA < 0 Then ' if "A" is a Low:
                                ' can stop looking as soon as the price goes below "A" or above "B"
                                If Bars(eBARS_Low, iBarC) < dPriceA Or Bars(eBARS_High, iBarC) > dPriceB Then
                                    Exit For
                                End If
                                ' "C" must also be a Low
                                If iSwingC >= 0 Then
                                    iSwingC = 0
                                Else
                                    dPriceC = Bars(eBARS_Low, iBarC)
                                End If
                            Else ' if "A" is a High:
                                ' can stop looking as soon as the price goes below "B" or above "A"
                                If Bars(eBARS_Low, iBarC) < dPriceB Or Bars(eBARS_High, iBarC) > dPriceA Then
                                    Exit For
                                End If
                                ' "C" must also be a High
                                If iSwingC <= 0 Then
                                    iSwingC = 0
                                Else
                                    dPriceC = Bars(eBARS_High, iBarC)
                                End If
                            End If
                            If iSwingC <> 0 Then
                                ' now calc all the Fib lines for this A->B->C
                                For iFib = 0 To aFibRatios.Size - 1
                                    ' calculate the total fib line weight = ratio weight * (Avg weight for swings A, B and C)
                                    dValue = aFibRatioWeights.Num(iFib)
                                    If dValue > 0 Then
                                        dValue = dValue * (aSwingPointWeights.Num(Abs(iSwingA)) + aSwingPointWeights.Num(Abs(iSwingB)) + aSwingPointWeights.Num(Abs(iSwingC))) / 3
                                        ' calculate the fib price = C + (B - A) * ratio
                                        dPrice = dPriceC + (dPriceB - dPriceA) * aFibRatios.Num(iFib)
                                        ' calculate value for prices based on proximity to the actual fib price
                                        ' (i.e. assign decreasing values as move away from the actual fib price)
                                        For iWindow = 0 To iTickProximity
                                            If iWindow = 0 Then
                                                aClusterPrices.Add dPrice
                                                aClusterValues.Add dValue
                                            Else
                                                d = dValue * (iTickProximity + 1 - iWindow) / (iTickProximity + 1)
                                                aClusterPrices.Add dPrice + iWindow * dMinMove
                                                aClusterValues.Add d
                                                aClusterPrices.Add dPrice - iWindow * dMinMove
                                                aClusterValues.Add d
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
#End If
                End If
            Next
        End If
    Next

ErrExit:
    Set CalcTimeClusters = aClusterBars
    Exit Function
    
ErrSection:
    RaiseError "CalcTimeClusters"
    Resume ErrExit
End Function


' TLB: this routine just used for testing while in development
Public Sub TestPriceCluster()

    Dim i&, iStart&, iEnd&, iTickProximity&, dPrice#, dProcTime#, dMaxValue#, iLen&
    Dim Clusters As cGdTable
    Dim Chart As cChart
    Dim Bars As cGdBars
    Dim aSwingPoints As New cGdArray, aSwingPointWeights As New cGdArray
    Dim aFibRatios As New cGdArray, aFibRatioWeights As New cGdArray
    
    Set Chart = ActiveChart.Chart
    If Not Chart Is Nothing Then
        Set Bars = Chart.Bars
        Set aSwingPoints = Chart.GetSwingArray
        If Not Bars Is Nothing And Not aSwingPoints Is Nothing Then
            iEnd = Bars.Size - 1
            iStart = iEnd - 210
            
            aFibRatios.Create eGDARRAY_Doubles, 0
            aFibRatioWeights.Create eGDARRAY_Doubles, 0
            aSwingPointWeights.Create eGDARRAY_Doubles, 4
            
            aSwingPointWeights(3) = 300
            aSwingPointWeights(2) = 200
            aSwingPointWeights(1) = 100
            
            aFibRatios.Add 0.618
            aFibRatioWeights.Add 1
                
            aFibRatios.Add 0.382
            aFibRatioWeights.Add 1
                
            aFibRatios.Add 0.5
            aFibRatioWeights.Add 1
                
            iTickProximity = 4
                    
            dProcTime = gdTickCount
            Set Clusters = CalcPriceClusters(Bars, iStart, iEnd, aSwingPoints, aSwingPointWeights, aFibRatios, aFibRatioWeights, _
                                iTickProximity, True, True, True, True, False)
            dProcTime = gdTickCount - dProcTime
            frmTest.AddList Str(Int(dProcTime)) & " ms, #records = " & Str(Clusters.NumRecords)
                                
            dMaxValue = 0
            For i = 0 To Clusters.NumRecords - 1
                'frmTest.AddList str(i) & vbTab & Bars.PriceDisplay(Clusters(0, i)) & vbTab & str(Clusters(1, i))
                If Clusters.Num(1, i) > dMaxValue Then
                    dMaxValue = Clusters.Num(1, i)
                End If
            Next
                                
            ' display a histogram
            i = Clusters.NumRecords - 1
            dPrice = Clusters(0, i)
            Do While i >= 0
                If Abs(dPrice - Clusters(0, i)) < Bars.MinMove / 2 Then
                    iLen = Int(100 * Clusters(1, i) / dMaxValue + 0.5)
                    If iLen < 1 Then iLen = 1
                    frmTest.AddList Bars.PriceDisplay(Clusters(0, i)) & vbTab & Str(Round(Clusters(1, i), 2)) & vbTab & String(iLen, "*")
                    i = i - 1
                Else
                    frmTest.AddList Bars.PriceDisplay(dPrice)
                End If
                dPrice = dPrice - Bars.MinMove
            Loop
            
        End If
    End If

End Sub

' ScansEnabled: to recalc all criteria after daily downloads for use in Filters
Public Property Get ScansEnabled() As Boolean
On Error GoTo ErrSection:

    Dim strKey$
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    If GetRegistryValue(rkLocalMachine, strKey, "DoScan", 0) = 0 Then
        ScansEnabled = False
    Else
        ScansEnabled = True
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "mMain.ScansEnabled-Get"
    Resume ErrExit
End Property

Public Property Let ScansEnabled(ByVal bScansEnabled As Boolean)
On Error GoTo ErrSection:

    Dim strKey$
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    If bScansEnabled = False Then
        SetRegistryValue rkLocalMachine, strKey, "DoScan", 0
    Else
        SetRegistryValue rkLocalMachine, strKey, "DoScan", 1
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "mMain.ScansEnabled-Let"
    Resume ErrExit
End Property

' func(symbol, cycle, start date, year filter, etc, bars, tree of arrays) -- return err msg
' Used to calculate all the #'s for the Seasonal Chart
' - returns err msg (if an error)
' - results returned in the tResults table -- all #'s are already scaled as % changes
'   (first column is Avg Trend, other columns are data from past cycles, fieldnames are descriptions)
' - and the primary Bars data is also returned (i.e. dates and prices for the "current cycle")
' - need to pass the symbol, bar period, cycle period, detrend, start date (JJJJJ or YYYYMMDD)
Public Function CalcSeasonalChart(tResults As cGdTable, Bars As cGdBars, _
            ByVal strSymbol$, Optional ByVal strBarPeriod$ = "", Optional ByVal strCycle$ = "1 year", _
            Optional ByVal bDetrend As Boolean = False, Optional ByVal nStartDate& = 0) As String
On Error GoTo ErrSection:

    Dim rc&, i&, s$, iSize&, iCycle&, dBase#, dRatio#, iStartBar&, iEndBar&, iCycleBar&
    Dim strYearFilter$, strStart$
    
    Dim astrParms As New cGdArray
    Dim astrExpr As New cGdArray
    Dim astrBarNames As New cGdArray
    Dim aArrayOfBars As New cGdArray
    Dim aExprResults As New cGdArray

    Dim aCycleBar As New cGdArray
    Dim aAvgTrend As New cGdArray
    Dim aBullTrend As New cGdArray
    Dim aBearTrend As New cGdArray
    Dim Bars67 As New cGdBars

    ' clear the table of results (passed back)
    tResults.Clear
    tResults.CreateField eGDARRAY_Doubles, 0, "Average Trend"
    tResults.CreateField eGDARRAY_Doubles, 1, "Bullish Trend"
    tResults.CreateField eGDARRAY_Doubles, 2, "Bearish Trend"
    tResults.NumRecords = 0
    
    ' convert future symbol to the combined continuous flavor
    If SecurityType(strSymbol) = "F" Then
        s = ConvertFutureSymbol(strSymbol, eCombinedSymbol)
        If Len(s) > 0 Then
            strSymbol = s
        End If
        ' and also load the 67 bars now
        strSymbol = Parse(strSymbol, "-", 1) & "-067"
        DM_GetBars Bars67, strSymbol, strBarPeriod
        strSymbol = Parse(strSymbol, "-", 1) & "-057" ' 57 will be the "primary" bars (passed back)
    Else
        Bars67.Size = 0
    End If
    
    ' load the primary bars
    DM_GetBars Bars, strSymbol, strBarPeriod
    If Bars.Size = 0 Then
        Exit Function
    End If
    Bars.AddForecastBars 1200 ' ? (need forecast length to cover at least one full cycle)
    iSize = Bars.Size
    If Bars67.Size = 0 Then
        Set Bars67 = Bars ' i.e. if not a future
    Else
        Bars67.AddForecastBars 1200 ' ? (need forecast length to cover at least one full cycle)
    End If
    
    ' setup the 3 markets for the engine call
    astrBarNames.Create eGDARRAY_Strings, 3
    astrBarNames(0) = "Market1"
    astrBarNames(1) = Chr(34) & "-057" & Chr(34)
    astrBarNames(2) = Chr(34) & "-067" & Chr(34)
    
    aArrayOfBars.Create eGDARRAY_Longs, astrBarNames.Size, 0
    aArrayOfBars(0) = Bars.BarsHandle
    aArrayOfBars(1) = Bars.BarsHandle
    aArrayOfBars(2) = Bars67.BarsHandle
   
    ' setup the 4 expressions (Cycle Bar Number, All Trend, Bullish Trend, Bearish Trend) for the engine call
    astrExpr.Create eGDARRAY_Strings, 4
    ' start date (convert to YYYYMMDD)
    If nStartDate < 10000000 Then
        nStartDate = JulToLong(nStartDate, True)
    End If
    strStart = Str(nStartDate)
    strStart = " ~22001, ~13" & Format(Len(strStart), "000") & strStart
    ' Cycle Bar Number:
    'Expression=Cycle Bar Number ("1 Year", 1900, " ") of "-067"
    'CodedText=~01014CycleBarNumber ~16001( ~07006"-067" ~22001, ~200061 Year ~22001, ~130041900 ~22001, ~20000 ~17001)
    strYearFilter = ""
    astrExpr(0) = "~01014CycleBarNumber ~16001( ~07006""-067"" ~22001, ~20" _
                    & Format(Len(strCycle), "000") & strCycle & strStart & " ~22001, ~20" _
                    & Format(Len(strYearFilter), "000") & strYearFilter & " ~17001)"
    ' Cycle Trend:
    'Expression=Cycle Trend ("1 Year", 0, 1900, " ") of "-067"
    'CodedText=~01010CycleTrend ~16001( ~07006"-067" ~22001, ~200061 Year ~22001, ~130010 ~22001, ~130041900 ~22001, ~20000 ~17001)
    If bDetrend Then
        s = " ~22001, ~130011"
    Else
        s = " ~22001, ~130010"
    End If
    ' Avg Trend (all years)
    strYearFilter = ""
    astrExpr(1) = "~01010CycleTrend ~16001( ~07006""-067"" ~22001, ~20" _
                    & Format(Len(strCycle), "000") & strCycle & s & strStart & " ~22001, ~20" _
                    & Format(Len(strYearFilter), "000") & strYearFilter & " ~17001)"
    ' Bull Trend (just bullish years)
    strYearFilter = "Bullish"
    astrExpr(2) = "~01010CycleTrend ~16001( ~07006""-067"" ~22001, ~20" _
                    & Format(Len(strCycle), "000") & strCycle & s & strStart & " ~22001, ~20" _
                    & Format(Len(strYearFilter), "000") & strYearFilter & " ~17001)"
    ' Bear Trend (just bearish years)
    strYearFilter = "Bearish"
    astrExpr(3) = "~01010CycleTrend ~16001( ~07006""-067"" ~22001, ~20" _
                    & Format(Len(strCycle), "000") & strCycle & s & strStart & " ~22001, ~20" _
                    & Format(Len(strYearFilter), "000") & strYearFilter & " ~17001)"
    
    ' and create arrays to hold the results from the engine call
    aCycleBar.Create eGDARRAY_Doubles, iSize
    aAvgTrend.Create eGDARRAY_Doubles, iSize
    aBullTrend.Create eGDARRAY_Doubles, iSize
    aBearTrend.Create eGDARRAY_Doubles, iSize
    aExprResults.Create eGDARRAY_Longs, astrExpr.Size, 0
    aExprResults(0) = aCycleBar.ArrayHandle
    aExprResults(1) = aAvgTrend.ArrayHandle
    aExprResults(2) = aBullTrend.ArrayHandle
    aExprResults(3) = aBearTrend.ArrayHandle

    ' call the engine
    astrParms.Create eGDARRAY_Strings, 1
    rc = ExecuteExpressionSet(astrParms, astrBarNames, aArrayOfBars, astrExpr, aExprResults)
    
    If rc <> 0 Then
        ' error when running the Cycle functions through the engine
        CalcSeasonalChart = "Error " & Str(rc) & ": " & astrParms(1)
    Else
        ' walk through the historical prices
        iEndBar = iSize - 1
        iStartBar = 0
        iCycle = 0
        For i = 0 To iSize - 1
            ' see if a new cycle is starting
            iCycleBar = aCycleBar.Num(i)
            If iCycleBar = 1 And aCycleBar.Num(i + 1) <> 1 Then
                ' see if we're now at or past the end of the data
                If Bars(eBARS_Close, i + 1) = kNullData Then
                    ' if so, then we're done
                    iEndBar = i
                    Exit For
                End If
                
                ' do last bar of prev cycle
                If iCycle > 0 And Bars(eBARS_Close, i - 1) > 0 Then
                    dRatio = (Bars67(eBARS_Close, i) - Bars67(eBARS_Close, i - 1)) / Bars(eBARS_Close, i - 1) + 1
                    dBase = dBase * dRatio
                    tResults.Num(iCycle + 2, i - iStartBar) = dBase - 100
                End If
                
                ' setup for new cycle (create new field in the table)
                iCycle = iCycle + 1
                iStartBar = i
                dBase = 100
                tResults.CreateField eGDARRAY_Doubles, iCycle + 2, DateFormat(Bars(eBARS_DateTime, i))
            End If
            If iCycle > 0 And Bars(eBARS_Close, i) <> kNullData Then
                ' calc price as % change from last bar
                If i = iStartBar Then
                    dBase = 100 ' new cycle always starts at 100
                ElseIf Bars(eBARS_Close, i - 1) <= 0 Then
                    Exit For ' ERROR?
                Else
                    ' this is best way to do a ratio for Futures (diff of 67 divided by the 57)
                    dRatio = (Bars67(eBARS_Close, i) - Bars67(eBARS_Close, i - 1)) / Bars(eBARS_Close, i - 1) + 1
                    dBase = dBase * dRatio
                End If
                'tResults.Num(iCycle + 2, i - iStartBar) = dBase - 100 ' i.e. as % change
                If iCycleBar > 0 Then
                    tResults.Num(iCycle + 2, iCycleBar - 1) = dBase - 100 ' i.e. as % change
                ElseIf IsIDE Then
                    InfBox "Isn't this an Error?", , , "DEBUG -- CalcSeasonalChart"
                End If
                
                If IsIDE Then
                    'frmTest.AddList DateFormat(Bars(eBARS_DateTime, i)) & vbTab & Str(iCycle) & vbTab & Str(aCycleBar(i)) & vbTab & Format(aAvgTrend(i), "#0.0000")
                End If
            End If
        Next
        
        ' chop everything down to the size of the "current cycle" (i.e. where data ends)
        iSize = iEndBar - iStartBar + 1
        Bars.DeleteFirstBars iStartBar
        Bars.Size = iSize
        tResults.NumRecords = iSize
        
        ' and put the Trends for the current cycle into the table
        For i = iStartBar To iEndBar
            If i = iStartBar Then
                dBase = 100
            ElseIf aAvgTrend.Num(i - 1) > 0 And aAvgTrend.Num(i) > 0 Then
                dRatio = aAvgTrend.Num(i) / aAvgTrend.Num(i - 1)
                dBase = dBase * dRatio
            End If
            tResults.Num(0, i - iStartBar) = dBase - 100 ' i.e. as % change
        Next
        For i = iStartBar To iEndBar
            If i = iStartBar Then
                dBase = 100
            ElseIf aBullTrend.Num(i - 1) > 0 And aBullTrend.Num(i) > 0 Then
                dRatio = aBullTrend.Num(i) / aBullTrend.Num(i - 1)
                dBase = dBase * dRatio
            End If
            tResults.Num(1, i - iStartBar) = dBase - 100 ' i.e. as % change
        Next
        For i = iStartBar To iEndBar
            If i = iStartBar Then
                dBase = 100
            ElseIf aBearTrend.Num(i - 1) > 0 And aBearTrend.Num(i) > 0 Then
                dRatio = aBearTrend.Num(i) / aBearTrend.Num(i - 1)
                dBase = dBase * dRatio
            End If
            tResults.Num(2, i - iStartBar) = dBase - 100 ' i.e. as % change
        Next
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.CalcSeasonalChart"
    Resume ErrExit
End Function

' Returns handle to daily bars of $IRX data for reports (needed for Sharpe Ratio)
Public Function GetIrxBarsHandle() As Long
On Error GoTo ErrSection:

    Static nLDD As Long
    
    If m.IrxBars Is Nothing Then
        Set m.IrxBars = New cGdBars
    End If
    
    ' only need to load once per day
    If nLDD < LastDailyDownload Or m.IrxBars.Size = 0 Then
        nLDD = LastDailyDownload
        DM_GetBars m.IrxBars, "$IRX"
    End If
    
    GetIrxBarsHandle = m.IrxBars.BarsHandle

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.GetIrxBarsHandle"
    Resume ErrExit
End Function

Private Sub ToolbarClickSeasonal(frmActiveChart As Form, tbToolbar As SSActiveToolBars, _
    ByVal Tool As ActiveToolBars.SSTool)
On Error Resume Next

    Dim i&
    Dim Pane As cPane
    Dim frm As Form
    Dim eExtraInfo As eToolbarExtraInfo

    If frmActiveChart Is Nothing Then Exit Sub
    If Not IsFrmChart(frmActiveChart) Then Exit Sub
    If frmActiveChart.Chart Is Nothing Then Exit Sub

    With frmActiveChart
        Select Case Tool.ID
            Case "ID_RepeatDraw"
                ToolbarSyncCursorGroup tbToolbar, Tool.ID
            
            Case "ID_Magnet"
                i = Abs(g.ChartGlobals.nMagnetValue)
                If i = 0 Then i = 5
                If Tool.State = ssUnchecked Then
                    i = -i
                End If
                g.ChartGlobals.nMagnetValue = i
                ToolbarSyncCursorGroup tbToolbar, Tool.ID
        
            Case "ID_AutoScale"
                If Tool.State = ssChecked Then
                    .Chart.AutoScale = True
                Else
                    .Chart.AutoScale = False
                End If
                .Chart.GenerateChart eRedo1_Scrolled
            
            Case "ID_Delete/Hide"
                If g.ChartGlobals.nHideAnnotations = 0 Then
                    tbToolbar.Tools("ID_HideAnnotations").State = ssUnchecked
                Else
                    tbToolbar.Tools("ID_HideAnnotations").State = ssChecked
                End If
                
            Case "ID_DeleteLastAnnotation"
                If .Chart.RemoveAnnots(True) > 0 Then
                    '.Chart.GenerateChart eRedo1_Scrolled
                    .Chart.SyncGlobalAnnots Nothing, True
                Else
                    Beep
                End If
                
            Case "ID_DeleteAllAnnotations"
                If .Chart.RemoveAnnots(False) > 0 Then
                    '.Chart.GenerateChart eRedo1_Scrolled
                    .Chart.SyncGlobalAnnots Nothing, True
                Else
                    Beep
                End If
                
            Case "ID_HideAnnotations"
                HideAnnotations Tool.State
            
            Case "ID_Trendline", "ID_Trendline2", "ID_Trendline3", "ID_Trendline4", "ID_TrendChannel", _
                 "ID_SRLine", "ID_HorzLine", "ID_HorzLine2", "ID_HorzLine3", "ID_HorzLine4", "ID_VertLine", _
                 "ID_ArrowLine", "ID_Text", "ID_Text2", "ID_Text3", "ID_Text4", "ID_Icon", _
                 "ID_ElliotLabels", "ID_ElliotEndUser", "ID_Bracket", "ID_Ellipse", "ID_Rectangle"
                
                'turn on annotations if they are hidden
                If g.ChartGlobals.nHideAnnotations = 1 Then HideAnnotations False
                ToolbarSetCursorGroup tbToolbar, True, Tool.ID
            
            Case "ID_CursorArrow", "ID_CursorCrosshairs", "ID_CursorHorizLine", "ID_CursorVertLine"
                ToolbarSetCursorGroup tbToolbar, False, Tool.ID
            
            Case "ID_ChartMove"
                g.ChartGlobals.eChartMode = eMode_Move
                ToolbarSetCursorGroup tbToolbar, False, Tool.ID
                
            Case "ID_DragModeY"
                If g.ChartGlobals.eDragModeY = eDragModeY_Each Then
                    g.ChartGlobals.eDragModeY = eDragModeY_Both
                Else
                    g.ChartGlobals.eDragModeY = eDragModeY_Each
                End If
                ToolbarSyncCursorGroup tbToolbar, Tool.ID
            
            Case "ID_Eraser"
                g.ChartGlobals.eChartMode = eMode_Erase     '6183
                ToolbarSetCursorGroup tbToolbar, False
            
            Case "ID_CursorArrow", "ID_CursorCrosshairs", "ID_CursorHorizLine", "ID_CursorVertLine"
                ToolbarSetCursorGroup tbToolbar, False, Tool.ID
            
            Case "ID_ZoomIn"
                If Len(g.strActiveDraw) = 0 Then
                    StatusMsg "To ZOOM, click on the chart and drag over the area to zoom."
                End If
                g.ChartGlobals.eChartMode = eMode_Zoom
                'StatusMsg 'so will force the "ZOOM" msg
                ToolbarSetCursorGroup tbToolbar, False
            
            Case "ID_ZoomOut"
                If .Chart.Zoomed Then
                    .Chart.UnzoomChart True
                End If
                
            Case "ID_MoreBars"
                .Chart.PixelsPerBar = -2
                .Chart.GenerateChart eRedo1_Scrolled
            Case "ID_LessBars"
                .Chart.PixelsPerBar = -1
                .Chart.GenerateChart eRedo1_Scrolled
            Case "ID_MoreAboveBelow"
                Set Pane = .Chart.Tree("PRICE PANE")
                If Not Pane Is Nothing Then
                    Pane.geIncDecMaxRatio 0.05
                    Pane.geIncDecMinRatio -0.05
                    Set Pane = Nothing
                End If
                .Chart.GenerateChart eRedo1_Scrolled
            Case "ID_LessAboveBelow"
                Set Pane = .Chart.Tree("PRICE PANE")
                If Not Pane Is Nothing Then
                    Pane.geIncDecMaxRatio -0.05
                    Pane.geIncDecMinRatio 0.05
                    Set Pane = Nothing
                End If
                .Chart.GenerateChart eRedo1_Scrolled
            Case "ID_BarPeriod", "ID_Yearly", "ID_Quarterly", "ID_Monthly", "ID_Weekly", "ID_Daily", _
                 "ID_360minute", "ID_240minute", "ID_180minute", "ID_90minute", "ID_60minute", _
                 "ID_30minute", "ID_15minute", "ID_10minute", "ID_5minute", "ID_3minute", "ID_1minute", _
                 "ID_CustomMinute", "ID_CustomPeriod"
                
                InfBox "Please use Seasonal Sidebar on chart.", "I", , "Seasonal Chart"
            
            Case "ID_Crosshair"
                'set for all visible charts
                For i = 0 To Forms.Count - 1
                    Set frm = Forms(i)
                    If IsFrmChart(frm) Then
                        frm.Chart.SetCursor
                    End If
                Next
            
            Case "ID_ResetChart"
                .Chart.RestoreChartNormal vbKeyReturn
            
            Case Else:
                InfBox kSeasonalUnavail, "I", "Ok", "Seasonal chart"
        End Select
    End With

End Sub

Public Sub ClearChartPointers(frm As Form)

    On Error Resume Next

    If m.ActiveChartForm Is frm Then Set m.ActiveChartForm = Nothing
    If m.PrevActiveForm Is frm Then Set m.PrevActiveForm = Nothing
    
    If g.ChartGlobals.frmActiveNonDetached Is frm Then Set g.ChartGlobals.frmActiveNonDetached = Nothing
    If g.ChartGlobals.frmLastChartMouseMove Is frm Then Set g.ChartGlobals.frmLastChartMouseMove = Nothing
    If g.ChartGlobals.frmPfpSelPattern Is frm Then Set g.ChartGlobals.frmPfpSelPattern = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeAccessFieldDataType
'' Description: Change the data type of a field in an Access database
'' Inputs:      Database, Table Name, Field Name, New Data Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeAccessFieldDataType(DB As Database, ByVal strTable As String, ByVal strField As String, Optional ByVal nNewDataType As DataTypeEnum = dbLongBinary)
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table in the database that contains the field
    Dim fld As field                    ' Field in the table to change
    Dim q As QueryDef                   ' Query to move the data over
    Dim rs As Recordset                 ' Recordset into the database
    Dim mb As New cMemBuffer            ' Buffer to hold data during conversion

    If ItemExists(DB.TableDefs, strTable) Then
        Set t = DB.TableDefs(strTable)
        With t
            If .Fields(strField).Type <> nNewDataType Then
                ' create a new field of the correct data type
                ' and insert into table after field being replaced
                Set fld = .CreateField("temp_new", nNewDataType)
                fld.OrdinalPosition = .Fields(strField).OrdinalPosition
                .Fields.Append fld
                
                ' duplicate some of the properties
                On Error Resume Next
                With .Fields(strField)
                    fld.DefaultValue = .DefaultValue
                    fld.Required = .Required
                    fld.AllowZeroLength = True ' .AllowZeroLength
                End With
                On Error GoTo ErrSection:
                
                ' copy the data to the new field
                If (nNewDataType = dbBinary Or nNewDataType = dbLongBinary) And _
                        (.Fields(strField).Type = dbText Or .Fields(strField).Type = dbMemo) Then
                    ' to convert from text to binary
                    Set rs = DB.OpenRecordset("SELECT * FROM [" & strTable & "];", dbOpenDynaset)
                    Do While Not rs.EOF
                        mb.Buffer = NullChk(rs.Fields(strField))
                        If mb.Length > 0 Then
                            rs.Edit
                            rs!temp_new = mb.Bytes
                            rs.Update
                        End If
                        rs.MoveNext
                    Loop
                    Set rs = Nothing
                Else
                    Set q = DB.CreateQueryDef("") '(temp update query)
                    q.SQL = "UPDATE " & strTable & " SET [" & strTable & "].[temp_new] = [" & strTable & "].[" & strField & "];"
                    q.Execute
                End If
                
                ' delete the old field and rename the new field
                .Fields.Delete strField
                .Fields("temp_new").Name = strField
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ChangeAccessFieldDataType"

End Sub


Public Function CreateGroupFromScreener(ByVal strText$) As Boolean
On Error GoTo ErrSection:

    Dim i&, iSymID&, strName$, strID$
    Dim aSymbols As New cGdArray
    Dim alSymbolIds As New cGdArray
    Dim frm As New frmSymbolGroup
    Dim bSuccess As Boolean
    
    strName = Parse(strText, vbTab, 1)
    
    ' get symbol ID for each symbol
    aSymbols.SplitFields Parse(strText, vbTab, 2), ","
    alSymbolIds.Create eGDARRAY_Longs, 0, 0
    For i = 0 To aSymbols.Size - 1
        iSymID = GetSymbolID(aSymbols(i))
        If iSymID > 0 Then
            alSymbolIds.Add iSymID
        End If
    Next
    alSymbolIds.Sort eGdSort_DeleteDuplicates
    
    If IsIDE And False Then
        aSymbols.Size = 0
        frmTest.AddList strName
        For i = 0 To alSymbolIds.Size - 1
            aSymbols.Add GetSymbol(alSymbolIds(i))
        Next
        aSymbols.Sort
        For i = 0 To aSymbols.Size - 1
            frmTest.AddList Str(i) & vbTab & GetSymbolID(aSymbols(i)) & vbTab & aSymbols(i)
        Next
        frmTest.AddList "finished CreateGroup"
    End If
    
    ' make ID same as name, except replace all invalid filename characters with an underscore
    strID = "~" & strName
    For i = 1 To Len(strID)
        If InStr(" .:\/*?|><+=", Mid(strID, i, 1)) > 0 Then
            Mid(strID, i, 1) = "_"
        End If
    Next
    strID = strID & ".GRP"
    
    AllowSetForegroundWindow GetCurrentProcess
    frm.ShowMe AddSlash(App.Path) & "Custom\", strID, False, alSymbolIds, False, , False, strName
    

#If 0 Then
    Select Case UCase(strReturn)
        Case "A":
            ' Append to the current Symbol Group
            ''SymbolGroup.Edit AddSlash(App.Path) & "Custom", strID, , , m.alSymbolIDs, True, True
            frm.ShowMe AddSlash(App.Path) & "Custom\", strID, False, m.alSymbolIds, True
        Case "O":
            ' Overwrite the current Symbol Group
            ''SymbolGroup.Edit AddSlash(App.Path) & "Custom", strID, , , m.alSymbolIDs, False, True
            frm.ShowMe AddSlash(App.Path) & "Custom\", strID, False, m.alSymbolIds, False
        Case Else:
            ' Create new Symbol Group with the Symbol ID's passed in
            ''SymbolGroup.Edit AddSlash(App.Path) & "Custom", strID, , , m.alSymbolIDs
            bSaveNew = optNew.Value
            frm.ShowMe AddSlash(App.Path) & "Custom\", strID, False, m.alSymbolIds, False, , bSaveNew
    End Select
#End If

    CreateGroupFromScreener = bSuccess
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.CreateGroupFromScreener"
    Resume ErrExit
End Function

Public Sub PageCollectionLoad()
On Error GoTo ErrSection:

    Dim bLocked As Boolean
    Dim bTimersSaved As Boolean

    If Not g.ChartPageCache Is Nothing Then g.ChartPageCache.Clear
    
    bTimersSaved = ChartTimers
    
    ChartTimers = False
    g.bLoadingChartPage = True
    g.bSkipSetChartFocus = True
    g.ChartGlobals.nDetached = 0
    Set m.ActiveChartForm = Nothing
    Set g.ChartGlobals.frmActiveNonDetached = Nothing
    bLocked = LockWindowUpdate(frmMain.hWnd)

    RestoreCharts False

    g.bLoadingChartPage = False
    g.bSkipSetChartFocus = False
    If bLocked Then LockWindowUpdate 0
    ChartTimers = bTimersSaved

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.PageCollectionLoad"

End Sub

' returns the primary base symbol used for the section header of the auto-exit favorites
Public Function BaseForAutoExitFavorites(ByVal strSymbol$) As String
On Error GoTo ErrSection:
    
    Dim strBase$
    
    Select Case SecurityType(strSymbol)
    Case "S"
        strBase = "STOCKS"
    Case "F"
        strBase = PrimaryFutureBase(Parse(strSymbol, "-", 1)) & "-"
    Case Else
        ' (TLB: still waiting to hear what to do with Forex symbols)
        strBase = strSymbol
    End Select
    BaseForAutoExitFavorites = strBase
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.BaseForAutoExitFavorites"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PlaceForm
'' Description: Place the given form appropriately
'' Inputs:      Form to Place
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PlaceForm(FormToPlace As Form)
On Error GoTo ErrSection:

    mGenesis.PlaceTheForm FormToPlace, g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.PlaceForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveFormPlacement
'' Description: Save the placement of the given form
'' Inputs:      Form to Save
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveFormPlacement(FormToSave As Form)
On Error GoTo ErrSection:

    mGenesis.SaveTheFormPlacement FormToSave, g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.SaveFormPlacement"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowAlertPopup
'' Description: Show an alert message in a non-modal, show on top window
'' Inputs:      Message, Caption, Text Alignment
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowAlertPopup(ByVal strMessage As String, ByVal strCaption As String, Optional ByVal nTextAlignment As AlignmentConstants = vbLeftJustify, _
            Optional ByVal bBoldFont As Boolean = False, Optional ByVal nBackColor As Long = vbButtonFace)
On Error GoTo ErrSection:

    Dim frm As frmAlertPopup            ' New alert popup form
    
    Set frm = New frmAlertPopup
    frm.ShowMessageBox strMessage, strCaption, nTextAlignment, bBoldFont, nBackColor

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ShowAlertPopup"
    
End Sub

' TLB: this routine is to help fix the inadvertant scrolling issues with the FlexGrids while streaming is on.
' To use, just add the following line into the grid's fg_BeforeScroll event:
'    GridScrollCheck fgGridName??, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
Public Sub GridScrollCheck(fg As VSFlexGrid, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    On Error Resume Next
    
    Dim pt As POINTAPI
    Static iVertHorizMode As Integer ' mode: 0=None, 1=Vert, 2=Horiz
       
    ' If mouse is not pressed, then just clear the mode
    If fg Is Nothing Or Not MouseIsPressed Then
        iVertHorizMode = 0
    ElseIf iVertHorizMode = 0 Then
        ' But if mouse is pressed and the mode has not yet been set, then user must be
        ' pressing a grid scrollbar -- so need to find out which one (vert or horiz).
        ' First get the X,Y coords of the mouse:
        If GetCursorPos(pt) <> 0 Then
            ' then convert to coords based on the grid itself
            If ScreenToClient(fg.hWnd, pt) <> 0 Then
                ' then see if the mouse is either below the grid (horiz scrollbar)
                ' or is to the right of the grid (vert scrollbar)
                If pt.X >= fg.ClientWidth / Screen.TwipsPerPixelX Then
                    iVertHorizMode = 1 ' Vert scrollbar is right of grid
                ElseIf pt.Y >= fg.ClientHeight / Screen.TwipsPerPixelY Then
                    iVertHorizMode = 2 ' Horiz scrollbar is below grid
                End If
            End If
        End If
        ' and start the timer which will clear this stuff as soon as the mouse is no longer pressed
        If iVertHorizMode <> 0 Then
            frmMain.tmrGridScrollPressed.Enabled = True
        End If
    End If
        
    If iVertHorizMode = 1 Then
        ' while the vertical scrollbar is pressed, we will NOT be scrolling columns
        If OldLeftCol <> NewLeftCol Then
            Cancel = True
        End If
    ElseIf iVertHorizMode = 2 Then
        ' while the horiz scrollbar is pressed, we will NOT be scrolling rows
        If OldTopRow <> NewTopRow Then
            Cancel = True
        End If
    End If

End Sub

' displays the list of shared chart pages (on our web server)
Public Sub DisplaySharedChartPages()
On Error GoTo ErrSection:

    Dim i&, strCodes$, strUrl$
    Dim aCodes As New cGdArray
        
    ' build string of just the SCP_* enablements
    strCodes = UCase(g.strAuthorizationString) & ",SCP_TEST"
    aCodes.SplitFields strCodes, ","
    strCodes = ","
    For i = 0 To aCodes.Size - 1
        If Left(aCodes(i), 4) = "SCP_" Then
            strCodes = strCodes & aCodes(i) & ","
        End If
    Next
    Set aCodes = Nothing
    
    ' append codes and the user's TradeNav build# (in case a minimum build is required for a shared chart page)
    strUrl = GetProvidedProperty("SharedChartPages", "http://www.TradeNavigator.com/SharedChartPage/index.php?U=*&P=*")
    strUrl = strUrl & "&B=*&C=" & strCodes
    
    ' display the web page with the list of shared chart pages
    RunWebReport "Shared Chart Pages", strUrl, "kSharedChartPage", 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.DisplaySharedChartPages"
End Sub

' Saves the chart page downloaded from RunWebReport and then loads it.
' from tab-delimited fields: ??, PageName, HexEncryptedGzp
Public Sub LoadSharedChartPage(ByVal strText$)
On Error GoTo ErrSection:

    Dim i&, strPage$, dTime#, strZip$, strFile$, strPath$, strCodeI$, strCodeD$
    Dim aCfg As New cGdArray, aFlds As New cGdArray, aFiles As New cGdArray
        
    ' name of shared chart page
    strPage = Parse(strText, vbTab, 2)
    If UCase(Right(strPage, 4)) = ".SCP" Then
        strPage = Left(strPage, Len(strPage) - 4)
    End If
    
    ' get published time (from # of seconds from 1/1/1970 in GMT)
    dTime = Val(Parse(strText, vbTab, 4))
    If dTime > 100000000# Then
        dTime = DateSerial(1970, 1, 1) + dTime / 86400#
        ' but we'll display this in ET (to be consistent with our web page)
        dTime = ConvertTimeZone(dTime, "GMT", "NY")
    Else
        dTime = 0
    End If
    
    ' decrypt the GZP file from the encrypted hex field
    strText = Parse(strText, vbTab, 3)
    If Len(strText) > 0 And dTime > 0 Then
        strZip = g.ChartGlobals.strCPCRoot & "\Charts\Pages\" & strPage & ".GZP"
        strText = DecryptFromHex(strText)
        FileFromString strZip, strText, , , True
        
        ' unzip it into the Temp area
        strPath = g.ChartGlobals.strCPCRoot & "\Charts\Pages\Temp\"
        MakeDir strPath
        KillFile strPath & "*.*"
        ZipExecute "U", strZip, strPath
        KillFile strZip
        
        ' set the published time for this chart page (to sync up with time of file on the Genesis server)
        If dTime > 0 Then
            'FileFromString strPath & "SCP.cfg", Str(dTime) & vbTab & strPage
            SetIniFileProperty "PageName", strPage, "", strPath & "SCP.INI"
            SetIniFileProperty "Published", Str(dTime), "", strPath & "SCP.INI"
        End If
        
        ' change to MAX
        strFile = strPath & "Charts.cfg"
        aCfg.FromFile strFile
        aFlds.SplitFields aCfg(0), vbTab
        If aFlds.Size > 1 Then
            aFlds(0) = "MAX"
            aCfg(0) = aFlds.JoinFields(vbTab)
            aCfg.ToFile strFile
        End If

        ' check enablements for intraday/daily charts (delete if not enabled)
        strCodeI = GetIniFileProperty("IntradayCode", "", "", strPath & "SCP.INI")
        strCodeD = GetIniFileProperty("DailyCode", "", "", strPath & "SCP.INI")
        If Len(strCodeI) > 0 Or Len(strCodeD) > 0 Then
            aFiles.GetMatchingFiles strPath & "*.CHT"
            For i = 0 To aFiles.Size - 1
                'Periodicity=285212673
                strFile = aFiles(i)
                If IsIntraday(GetIniFileProperty("Periodicity", 0, "General", strFile)) Then
                    ' intraday chart
                    If HasModule(strCodeI) Then
                        strFile = "" ' this one's ok
                    End If
                Else
                    ' daily (or above) chart
                    If HasModule(strCodeD) Then
                        strFile = "" ' this one's ok
                    End If
                End If
                If Len(strFile) > 0 Then
                    ' not enabled for this chart, so delete it (along with its .ANO files)
                    KillFile strFile
                    strFile = Left(strFile, Len(strFile) - 4) & "^*.ANO"
                    KillFile strFile
                End If
            Next
        End If
        
        ' zip it back up into the Pages folder
        ZipExecute "C", strZip, strPath
        KillFile strPath & "*.*"
        
        ' need to remove this chart page from the cache (if exists there)
        CachePageRemove strPage
        
        ' if they are currently on this chart page, do NOT save their changes
        If UCase(strPage) = UCase(g.strChartPage) Then
            LoadChartPage strPage, True
        Else
            ' but if on a different page, then ask if they want to save their changes
            LoadChartPage strPage, False
        End If
    Else
        InfBox "Error downloading shared chart page:|" & strPage, "e", , "Error"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.LoadSharedChartPage"
End Sub

' for certain users to publish (upload) their chart page to our server
Public Function PublishSharedChartPage() As Boolean
On Error GoTo ErrSection:

    Dim i&, j&
    Dim strPage$, strText$, strPublishZip$
    Dim Annots As cGdTree
    Dim Annot As cAnnotation
    
    ' convert all their global annots to "local" (only showing on this chart page)
    For i = 0 To Forms.Count - 1
        If IsFrmChart(Forms(i)) Then
            Set Annots = Forms(i).Chart.Annots
            If Not Annots Is Nothing Then
                For j = 1 To Annots.Count
                    Set Annot = Annots(j)
                    If Annot.MultiChartFlag Then
                        Annot.MultiChartFlag = False
                    End If
                Next
            End If
        End If
    Next
    
    strPublishZip = App.Path & "\_Publishing.Zip"
    If SaveChartPage("", strPublishZip) Then
        strPage = g.strChartPage
        strText = FileToString(strPublishZip, , , True)
        If Len(strText) > 0 Then
            strText = EncryptToHex(strText)
            SendWebPage strPage & ".SCP", strText
            PublishSharedChartPage = True
            i = ZipExecute("N", strPublishZip, "", "*.CHT")
            InfBox Str(i) & " charts published to:|" & strPage, "i", , "Publish"
        End If
    End If
    KillFile strPublishZip, True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.PublishSharedChartPage"
End Function

Public Sub TimerStart(ByVal strName As String)
On Error GoTo ErrSection:

    Static bDumpTimers As Boolean       ' Dump all the timers always?
    Static dCheckAgain As Double        ' Next time to check for flag file

    ' check once a minute for the flag file to start the timer dumps
    If gdTickCount > dCheckAgain Then
        dCheckAgain = gdTickCount + 60000
        bDumpTimers = FileExist(AddSlash(App.Path) & "DumpTimers.FLG")
        If bDumpTimers = False And m.TimerStarts.Count > 0 Then
            m.TimerStarts.Clear
        End If
    End If

    If bDumpTimers Then
        If m.TimerStarts.Exists(strName) Then
            m.TimerStarts(strName) = gdTickCount
        Else
            m.TimerStarts.Add gdTickCount, strName
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.TimerStart"
    
End Sub

Public Function TimerEnd(ByVal strName As String, ByVal lInterval As Long) As Boolean
On Error GoTo ErrSection:

    Dim dStartTime As Double            ' Start time for the timer
    Dim dElapsedTime As Double          ' Elapsed time for the timer

    ' if not logging timers, the collection count will stay at 0
    If m.TimerStarts.Count > 0 Then
        If lInterval >= 100 Then
            If m.TimerStarts.Exists(strName) Then
                dStartTime = m.TimerStarts(strName)
                dElapsedTime = gdTickCount - dStartTime
                If dElapsedTime > lInterval Then
                    TimerLog "Start = " & Str(dStartTime) & vbTab & "Time = " & Str(dElapsedTime) & vbTab & "Interval = " & Str(lInterval) & vbTab & strName
                    TimerEnd = True ' return true if logged the time
                End If
            End If
        End If
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mMain.TimerEnd"
    
End Function

Private Sub TimerLog(ByVal strMessage$)
On Error Resume Next

    Dim fh%
    Dim strPath$
    Static bAlreadyDone As Boolean
    
    If Len(strMessage) = 0 Then Exit Sub
    
    strPath = AddSlash(App.Path) & "TimerLogs\"
        
    If Not bAlreadyDone Then
        bAlreadyDone = True
        MakeDir strPath
        KillFile strPath & "*.LOG /o=-90"
    End If

#If 0 Then
    
    fh = FreeFile
    Open strPath & Format(Date, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ")" & vbTab & strMessage
        Close #fh
    End If

#Else

    Static LogFile As cLogFile
    If LogFile Is Nothing Then
        Set LogFile = New cLogFile
        LogFile.OpenFile strPath & "*.LOG"
    End If
    LogFile.WriteText strMessage
    
#End If

End Sub

Public Function CanDoAcctStatusWebPage() As Boolean

    'If HasLevel(eTN5_Professional, False) Then
    If HasLevel(eTN4_Gold, False) Then
        'If HasModule("BRKRLIVE") Then
            If HasModule("RTG") Then 'And HasModule("B_*") Then
                CanDoAcctStatusWebPage = True
            End If
        'End If
    End If

End Function

' to allow stock dividends and mutual funds
Public Function AllowDivAndMF() As Boolean

    If App.Major > 6 Or App.Minor >= 7 Then ' if version 6.7 or higher
        AllowDivAndMF = True
    ElseIf FileExist(App.Path & "\DivAdjust.flg") Then
        AllowDivAndMF = True
    End If

End Function

Public Sub FtpInstallCheck()

If Not IsIDE Then
    On Error Resume Next
End If

    Dim s$
    Static dNextCheck#

    ' check if an FTP Data Download is ready to install
    If Not g.RealTime.Active Then
        If gdTickCount > dNextCheck Then
            dNextCheck = gdTickCount + 3000
            If frmDataInstall2.ReadyToLink Then
                If Not ProcessIsBusy(True) Then
                    s = "Link in the downloaded historical data now?||(or can select 'Install Data' under 'File' menu later)"
                    If InfBox(s, "?", "+OK|-Not Now", "Data Install", , 10) = "O" Then
                        InstallData
                    Else
                        ' wait a few minutes?
                        dNextCheck = gdTickCount + 60000# * 5
                    End If
                End If
            End If
        End If
    End If

End Sub

' returns True if initiating a reboot
Private Function RebootNVSIfRequired() As Boolean

    Dim strFile$, bRebootNow As Boolean

    ' see if this is an NVS machine (check for a Restart.bat file)
    strFile = AddSlash(FilePath(App.Path)) & "Restart.bat"
    If FileExist(strFile) Then
        ' see if a reboot is required (e.g. due to a Windows Update)
        If IsRebootRequired Then
            bRebootNow = True
        ElseIf FileExist(AddSlash(FilePath(App.Path)) & "Reboot.Now") Then ' just for testing?
            KillFile AddSlash(FilePath(App.Path)) & "Reboot.Now"
            bRebootNow = True
        End If
        If bRebootNow Then
            ' first display a message so they know why the machine is being rebooted
            MsgBox "A recent Windows Update now requires this machine to be rebooted.", vbExclamation, "Reboot Required"
            Shell strFile, vbNormalFocus
            RebootNVSIfRequired = True
        End If
    End If

End Function

' Get messages from file on our web server:
' - each line of message file has a message (so multiple messages can be kept)
' - TN keeps track of last message ID# which it has already "processed"
' - message: Sequence#, ReqModules/CustIDs, MsgAction, Message
Public Sub CheckForTradeNavMessages()

If Not IsIDE Then
    On Error Resume Next ' we just don't ever want this routine to cause any errors on client machines
End If

    Dim i&, iLine&, nMsgSeq&, nLastSeq&, nAction&, iSeconds&
    Dim s$, strReq$, strMessage$, strCaption$
    Dim bFirstTime As Boolean
    Dim aLines As New cGdArray
    Dim aFlds As New cGdArray
    Static dPrevCheckTime#, iRecheckSeconds&
    
    ' only check once every couple of minutes (and not when a modal form is up)
    If g.bStarting Or g.bUnloading Then Exit Sub
    If iRecheckSeconds <= 0 Then iRecheckSeconds = 120 ' default to 2 minutes between checks
    If gdTickCount < dPrevCheckTime + iRecheckSeconds * 1000 Then Exit Sub
    If frmMain.Enabled = False Or ProcessIsBusy(True) Then Exit Sub
    dPrevCheckTime = gdTickCount

    ' get the last message sequence# that was already processed
    nLastSeq = GetIniFileProperty("TradeNavMsgSeq", 0, "", g.strIniFile)
    If nLastSeq <= 0 Then
        bFirstTime = True ' if has not ever processed any messages yet
        nLastSeq = 0
    End If

    ' get messages from file on our web server
    s = FixURL("www.TradeNavigator.com/DataInst/TradeNavMessages.txt")
    s = GetWebPageData(s)
    aLines.SplitFields s, vbCrLf
    ' make sure we got a text file and not some goofy default HTML file
    s = aLines(0)
    If Len(s) > 0 And InStr(s, "//") = 0 Then
        For iLine = 0 To aLines.Size - 1
            s = Trim(aLines(iLine))
            If IsDigit(s, 1) Then
                aFlds.SplitFields s, "|"
                ' only process for Sequence# > Last Msg processed
                nMsgSeq = Val(aFlds(0))
                If nMsgSeq > nLastSeq Then
                    nLastSeq = nMsgSeq
                    ' if first time, then we just want to ignore all the existing messages
                    ' but set the last processed message to the highest current sequence#
                    If Not bFirstTime Then
                        For i = 1 To aFlds.Size - 1
                            aFlds(i) = Trim(aFlds(i))
                        Next
                    
                        nAction = Val(aFlds(2))
                        ' but check if required Modules or CustomerID's (comma-delimited list)
                        strReq = aFlds(1)
                        If IsDigit(strReq) Then
                            If InStr(strReq, "-") > 1 And InStr(strReq, ",") = 0 Then
                                ' range of build #'s (e.g. "0-1352" or "1353-1360" or "1361-99999")
                                i = Val(Parse(strReq, "-", 1))
                                If App.Revision < i Then
                                    nAction = 0
                                Else
                                    i = Val(Parse(strReq, "-", 2))
                                    If App.Revision > i Then
                                        nAction = 0
                                    End If
                                End If
                            Else
                                ' list of CustomerID's
                                i = Int(RI_GetDataServiceID \ 1000)
                                If InStr("," & strReq & ",", "," & Str(i) & ",") = 0 Then
                                    nAction = 0 ' ignore if not in list of CustomerIDs
                                End If
                            End If
                        ElseIf Not HasModule(strReq) Then
                            nAction = 0
                        End If
                    End If
                    
                    ' do action based on type
                    Select Case nAction
                    Case 1 ' Popup message
                        strMessage = aFlds(3)
                        strCaption = aFlds(4)
                        If Len(strCaption) = 0 Then
                            strCaption = "IMPORTANT NOTE ..."
                        End If
                        ShowAlertPopup strMessage, strCaption, vbLeftJustify, True, &H80FFFF
                    
                    Case 2 ' Need to redo the daily download (if already done for today)
                        If LastDailyDownload = Int(CurrentTime("NY") - 0.5) Then
                            ' determine "window" during which each client will randomly redownload daily update
                            i = Val(aFlds(3)) '(# of minutes can be passed down)
                            If i > 180 Then
                                i = 180 ' max = 3 hours
                            ElseIf i < 10 Then ' min = 10 minutes
                                i = 60 ' default = 1 hour
                            End If
                            iSeconds = i * 60
                            
                            If g.RealTime.ActiveRTG Then
                                ' do anyone streaming in the first half of the window
                                ' (but allow all live auto-traders in first)
                                If HasModule("BRKRAUTO") Then
                                    iSeconds = RandomNum(1, 90)
                                Else
                                    iSeconds = RandomNum(100, Int(iSeconds / 2))
                                End If
                            Else
                                ' non-streaming gets put into the second half of the window
                                iSeconds = RandomNum(Int(iSeconds / 2), iSeconds)
                            End If
                            g.dNextDownloadTry = Now + iSeconds / 86400#
                        End If
                    End Select
                End If
            ElseIf Left(s, 8) = "RECHECK=" Then
                i = Val(Mid(s, 9))
                If i >= 10 And i <= 9999 Then
                    iRecheckSeconds = i
                End If
            End If
        Next
    End If
    
    SetIniFileProperty "TradeNavMsgSeq", nLastSeq, "", g.strIniFile

End Sub

Public Function TranslatedText(ByVal strTranslationID$, Optional ByVal strDefaultText$ = "") As String

    Dim iPos&, strText$
    Static Translations As cGdArray

    If Translations Is Nothing Then
        ' first time initialization
        Set Translations = New cGdArray
        If FileExist(App.Path & "\German.flg") Then
            g.iLanguage = 1
        End If
        ' load the Translations text
        If g.iLanguage > 0 Then
            Translations.FromFile App.Path & "\Provided\TextTranslations.txt"
            Translations.Sort eGdSort_DeleteNullValues Or eGdSort_IgnoreCase, 0
            If Translations.Size = 0 Then
                g.iLanguage = 0
            End If
        End If
    End If

    If g.iLanguage > 0 Then
        If Translations.BinarySearch(strTranslationID & vbTab, iPos, eGdSort_MatchUsingSearchStringLength Or eGdSort_IgnoreCase) Then
            strText = Parse(Translations(iPos), vbTab, g.iLanguage + 1, False)
        End If
    End If
    
    If Len(strText) > 0 Then
        TranslatedText = strText
    Else
        TranslatedText = strDefaultText
    End If

End Function

Private Sub DeleteAutoTradeItems()

    On Error Resume Next
    Dim rs As Recordset
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] ;", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    Set rs = Nothing

End Sub




