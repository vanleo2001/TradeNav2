Attribute VB_Name = "mDataNav"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mDataNav.bas
'' Description: Download routines for calling FRED
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 02/05/2010   DAJ         New functionality for new stock option symbols
'' 03/05/2010   DAJ         Only ask import QBT questions once
'' 05/17/2010   DAJ         Added the ChopDailyBars routine
'' 08/10/2010   DAJ         Don't do an FtpRequest when unloading
'' 08/19/2010   DAJ         Enhanced IsExpiredContract function
'' 08/20/2010   DAJ         Added SFE broker allowance
'' 08/23/2010   DAJ         Fixes for flattening expired stock option position
'' 08/30/2010   DAJ         Added MFGlobal and LindWaldock to SFE broker allowance
'' 09/28/2010   DAJ         Added TriggeredBy price to NumTicksFromMarket function (#5947)
'' 12/10/2010   DAJ         Changed over to the IsBrokerUser function
'' 02/03/2011   DAJ         Added the ValidPfgFxTradingTime function
'' 04/28/2011   DAJ         Modified ValidPfgFxTradingTime for Friday, Saturday, and Sunday
'' 05/16/2011   DAJ         Fix for CurrentTime in case delayed streaming for a symbol
'' 06/09/2011   DAJ         Added the AutoBreakoutPeriod function
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 12/09/2011   DAJ         Modified the IsExpiredContract function
'' 01/10/2012   DAJ         Added minimum breakout range for auto breakout calculation
'' 01/13/2012   DAJ         Fixed IsExpiredContract so that non F,FO,SO don't expire
'' 07/13/2012   DAJ         Added check for valid trading time for IB Forex
'' 08/17/2012   DAJ         Made generic valid trading time for broker forex
'' 01/18/2013   DAJ         SFE symbols allowed for CQG
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 05/01/2013   DAJ         Shadow Trading
'' 07/24/2013   DAJ         Beefed up the IsExpiredContract function
'' 05/05/2014   DAJ         Added optional Ending Date to AutoBreakoutPeriod
'' 06/06/2014   DAJ         Tell the automated trading items collection that enablements changed
'' 10/23/2014   DAJ         Mods to Spread component functions
'' 07/15/2015   DAJ         Added IsRth and IsEth functions
'' 07/16/2015   DAJ         Pass the time to check into the IsEth function
'' 07/20/2015   DAJ         Added debug line to IsRth function
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Declare Function Opt_BlackSholes Lib "OptLib2.dll" _
    (ByVal dStockPrice#, ByVal dStrikePrice#, ByVal iDaysToExpire%, _
     ByVal dVolatility#, ByVal dInterestRate#, ByVal dCarryCost#, _
     ByVal bIsPut As Byte, dDelta#) As Double

Public Declare Function Opt_GetVolatility Lib "OptLib2.dll" _
    (ByVal dOptionPrice#, ByVal dStockPrice#, ByVal dStrikePrice#, _
     ByVal iDaysToExpire%, ByVal dInterestRate#, ByVal dCarryCost#, _
     ByVal bIsPut As Byte, dDelta#) As Double
                                                     
Public Declare Function Opt_Delta Lib "OptLib2.dll" _
    (ByVal dStockPrice#, ByVal dStrikePrice#, ByVal iDaysToExpire%, _
     ByVal dVolatility#, ByVal dInterestRate#, ByVal dCarryCost#) As Double

Public Declare Function Opt_Gamma Lib "OptLib2.dll" _
    (ByVal dStockPrice#, ByVal dStrikePrice#, ByVal iDaysToExpire%, _
     ByVal dVolatility#, ByVal dInterestRate#, ByVal dCarryCost#) As Double

Public Declare Function Opt_Vega Lib "OptLib2.dll" _
    (ByVal dStockPrice#, ByVal dStrikePrice#, ByVal iDaysToExpire%, _
     ByVal dVolatility#, ByVal dInterestRate#, ByVal dCarryCost#) As Double

Public Declare Function Opt_Theta Lib "OptLib2.dll" _
    (ByVal dStockPrice#, ByVal dStrikePrice#, ByVal iDaysToExpire%, _
     ByVal dVolatility#, ByVal dInterestRate#, ByVal dCarryCost#, _
     ByVal bIsPut As Byte) As Double

Public Declare Function Opt_Rho Lib "OptLib2.dll" _
    (ByVal dStockPrice#, ByVal dStrikePrice#, ByVal iDaysToExpire%, _
     ByVal dVolatility#, ByVal dInterestRate#, ByVal dCarryCost#, _
     ByVal bIsPut As Byte) As Double

Private Const strPasswordKey$ = "ToEncryptThePassword" '"TheKeyForEncryptingPasswords"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FtpRequest
'' Description: Calls gclient to get a response back from fred with the
''              appropriate data files
'' Inputs:      Form that is calling the function
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FtpRequest(aRequest As cGdArray, Optional ByVal bTransfer As Boolean = False, _
                    Optional ByVal lSecondsToWait As Long = 150, _
                    Optional ByVal bSkipSymbolReload As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim s$
    Dim strDataLine As String           ' Line from the open file
    Dim strDirectoryPath As String      ' Directory to search in
    Dim strDownloadDate As String       ' Date of the last download
    Dim strOldDate As String            ' Old date
    Dim bValidUser As Boolean           ' Is this a valid user?
    Dim lValidate As Long               '
    Dim strTemp As String               ' Temporary string
    Dim strPgm As String, strArgs As String
    Dim bSuccess As Boolean             '
    Dim strKey As String                ' Key for the registry
    Dim dStartDate As Double            ' Starting date of account from registry
    Dim strPurchased As String
    
    Dim lDataServID As Long
    Dim strMachineID As String
    Dim lExpirationDate As Long
    Dim strReturn As String
    Dim lNumDays As Long
    Dim strStatus As String
    Dim strOldAuth As String
    Dim strEngineErr As String
    Dim strAuthBillMode As String
    Dim lNumTries As Long
    Dim astrIbisFile As New cGdArray
    Dim bTryOtherProtocol As Boolean
    Dim bUseFTP As Boolean
    Dim dStartTime As Double

TryAgain:
DebugLog "FtpRequest 1"

    ' DAJ 08/10/2010: We don't want to do a GClient request when Trade Navigator is unloading...
    If (g.bUnloading = True) Then
        If IsIDE Then
            Err.Raise vbObjectError + 1000, "Downloading on exit"
        Else
            Exit Function
        End If
    End If

    If aRequest Is Nothing Then Exit Function
    If aRequest.Size = 0 Then Exit Function

    ' GClient will ask for the file every 1.5 seconds -- so should determine
    ' the number of tries based on how long to wait (in seconds)
    lNumTries = Round(lSecondsToWait / 1.5)

    ' Make sure that necessary directories exist...
    MakeDir AddSlash(App.Path) & "FTP\Backup", False
    MakeDir AddSlash(App.Path) & "FTP\Dist", False
    MakeDir AddSlash(App.Path) & "SimTrade\In", False
    
    ' Set ftp directory
    strDirectoryPath = App.Path & "\ftp"

    ' Clean out ftp directory
    ClearReadOnlyFlags App.Path & "\ftp\*.*"
    KillFile AddSlash(App.Path) & "Ftp\*.*", True
    
    ' Sync Gclient according to FTP/HTTP setting
    If Not bTryOtherProtocol Then
        bUseFTP = SyncGclient '(first time, get preferred Gclient: FTP or HTTP)
    ElseIf bUseFTP Then
        frmStatus.AddDetail "FTP failed, now trying HTTP"
        bUseFTP = SyncGclient(False)
    Else
        frmStatus.AddDetail "HTTP failed, now trying FTP"
        bUseFTP = SyncGclient(True)
    End If
    DoEvents
    
DebugLog "FtpRequest 2"

    ' create request file
    FileFromString App.Path & "\ftp\Request.txt", "/ZIP:yes", True
    FileFromString App.Path & "\ftp\Request.txt", "/File Type:ASC=GTX", True, True
    FileFromString App.Path & "\ftp\Request.txt", "/Bad Tick File:yes", True, True
    
    ' 02/04/2010 DAJ: Added the following change for new option symbology...
    FileFromString App.Path & "\ftp\Request.txt", "/New Opt Symbols:Yes", True, True
    
    strTemp = "/ProtocolMode:" ''& Str(GetIniFileProperty("UseFTP", 0, "Mode", App.Path & "\Gclient.ini"))
    If bTryOtherProtocol Then
        strTemp = strTemp & " (first protocol failed)"
    End If
    FileFromString App.Path & "\ftp\Request.txt", strTemp, True, True
    aRequest.ToFile App.Path & "\ftp\Request.txt", True
    
DebugLog "FtpRequest 3"

    lDataServID = RI_GetDataServiceID
    strMachineID = RI_GetMachineID
    lExpirationDate = RI_GetExpirationDate
    strStatus = RI_GetIBISStatus

DebugLog "FtpRequest 4"

    astrIbisFile.Clear
    If bTransfer = False Then
        astrIbisFile.Add "Action=ValDSrv"
        astrIbisFile.Add "MachineID=" & strMachineID
        astrIbisFile.Add "DSrvID=" & Trim(Str(lDataServID))
        astrIbisFile.Add "Password=" & RI_GetUserPassword
        astrIbisFile.Add "DateTime=" & Format(Date, "YYYYMMDD") & " " & Format(Time, "HHMMSS")
    Else
        astrIbisFile.Add "Action=IBISREQUEST"
        astrIbisFile.Add "MachineID=" & strMachineID
        astrIbisFile.Add "DateTime=" & Format(Date, "YYYYMMDD") & " " & Format(Time, "HHMMSS")
    End If
    ' NOTE: the ", b????," portion must be maintained since that is checked by the servers
    astrIbisFile.Add "Source=TradeNav " & FormatVersion & ", b" & Str(App.Revision) & ", " & WindowsVersionStr
    astrIbisFile.ToFile App.Path & "\Ftp\Ibis.TXT"

DebugLog "FtpRequest 5"

    ' If the user is running SimTrade, zip up the necessary stuff in the "Out"
    ' directory to send to the Trade Server...
    If Not g.SimTradeTs Is Nothing Then
        If g.SimTradeTs.UseSalmon = False Then
            If FileExist(AddSlash(App.Path) & "SimTrade\Out\*.*") Then
                ZipExecute "A", AddSlash(App.Path) & "FTP\Orders.GZP", "", AddSlash(App.Path) & "SimTrade\Out\*.ORD"
                ZipExecute "A", AddSlash(App.Path) & "FTP\Orders.GZP", "", AddSlash(App.Path) & "SimTrade\Out\*.ACK"
            End If
        End If
    End If
    
    ' see if anything exists to upload
    If FileExist(App.Path & "\FTP\Upload\*.*") Then
        astrIbisFile.ToFile App.Path & "\FTP\Upload\Ibis.TXT"
        ZipExecute "A", App.Path & "\FTP\Upload.GZP", "", App.Path & "\FTP\Upload\*.*"
    End If
      
DebugLog "FtpRequest 6"

    'frm.txtHwnd.Text = "Connecting to GFDS ..."
    'frmStatus.UpdateProgress "Connecting to GFDS"
      
    ' Internal IP address 10.1.1.80
    ' External IP address ftp request 12.10.144.230
    KillFile App.Path & "\gclient.can"
    frmStatus.Status = eStatus_Running
    frmStatus.AddDetail "Sending Request"
    
    If bUseFTP Then
        strPgm = App.Path & "\GClientF.exe"
    Else
        strPgm = App.Path & "\GClient.exe"
    End If
    'If FileDate(strPgm) < DateSerial(2005, 11, 25) Then
        strArgs = Chr(34) & "/s=" & App.Path & "\ftp" & Chr(34) & " /c=" & Str(lNumTries) & " " & Chr(34) & "/d=" & App.Path & "\ftp" & Chr(34) & " /h=" & frmStatus.txtHwnd.hWnd
    'Else
    '    strArgs = Chr(34) & "/s=" & App.Path & "\ftp" & Chr(34) & " /c=" & Str(lNumTries) & " " & Chr(34) & "/d=" & App.Path & "\ftp" & Chr(34) & " /m=" & Str(frmStatus.MsgHwnd)
    'End If
    If bTransfer Then strArgs = strArgs & " /n"
    FileFromString App.Path & "\gclient.cal", strPgm & vbCrLf & strArgs, True, False
    ' "RunProcess" works much better than "Shell" -- returns immediately
    RunProcess strPgm, strArgs, , vbHide
    
    'frmUpdateBar.Caption = "Downloading..."
  
DebugLog "FtpRequest 7"

    ' Loop until done or error is received
    dStartTime = gdTickCount
    Do
        LoadAppBkImage ' (may as well load background image while waiting)
        Sleep 0.1
        Select Case frmStatus.Status
            Case eStatus_Completed
                bSuccess = True
                Exit Do
            Case eStatus_Aborted, eStatus_Error
                ' If an error occured while downloading or the user aborted the
                ' download, make sure that we randomly pick a new Twdl the next
                ' time.  11/05/2003 DAJ
                strKey = "Software\Genesis Financial Data Services\GClient"
                SetRegistryValue rkLocalMachine, strKey, "IPLastTime", 0&
            
                bSuccess = False
                Exit Do
            Case Else
                ' timeout if not connected within 2 minutes
                If frmStatus.LastKnownStatusCode = 100 Then
                    If gdTickCount > dStartTime + 120000 Then
                        FileFromString App.Path & "\Gclient.can", "Abort"
                        frmStatus.AddDetail "Error connecting"
                        frmStatus.Status = eStatus_Error
                        Sleep 3
                        Exit Do
                    End If
                End If
        End Select
    Loop
    
    ' TLB 9/16/05: Try other protocol if fails the first time
If 0 Then
    If Not bTryOtherProtocol And (frmStatus.Status = eStatus_Error) Then
        If Not FileExist(App.Path & "\ftp\request.inf") Then
            bTryOtherProtocol = True
            Sleep 1
            GoTo TryAgain
        End If
    End If
End If
    
    ' must wait a second or registry won't be updated yet!
    Sleep 1
    
DebugLog "FtpRequest 8"

    'Unzip DATA.GZP if exists
    If FileExist(App.Path & "\Ftp\ReqData.GZP") Then
        ZipExecute "U", App.Path & "\Ftp\ReqData.GZP", App.Path & "\Ftp\"
    End If
    
    ' TLB 8/1/2008: check the clock difference (based on the Genesis servers)
    If gdTickCount < dStartTime + 600000 Then  '(if download took over 10 minutes then skip this)
        If Abs(ClockDiff) > 4 / 24# Then ' if over 4 hours off
            If frmMain.Enabled Then ' (don't show modal dialog if another one is up right now)
                s = "Your computer's Date, Time or Time Zone may need to be adjusted.  It is currently set to:||" _
                    & DateFormat(Now, MM_DD_YYYY, H_MM, AMPM_LOWER)
                'InfBox s, "!", , "PLEASE CHECK"
                ShowAlertPopup s, "PLEASE CHECK"
            End If
        End If
    End If
    
    'Get authorization string (in case changed)
    strOldAuth = g.strAuthorizationString
    GetAuthorizationStringFromRegistry
    
    strPurchased = DM_GetPurchased(True)
    DM_LoadAuth strPurchased
    
DebugLog "FtpRequest 9"

    ' If the authorization string changed, reinit the engine so proper libraries will be recognized
    If strOldAuth <> g.strAuthorizationString Then
        SetMainCaption
        ToolbarReset True ' #6837 - call with bReset = True to force a toolbar reset when enablements change
        
        'If frmStatus.Visible And Not bSkipSymbolReload Then
        If 0 Then
            frmStatus.AddDetail "Reloading Symbols"
            frmStatus.UpdateProgress "Reloading Symbols"
            g.SymbolPool.Load False
        End If
        
        InitEngine False, strEngineErr
        InitEngine True, strEngineErr
        
        ' Refresh the Trade Console form in case module codes changed...
        g.Broker.InitBrokerObjects
        If FormIsLoaded("frmTTSummary") Then
            frmTTSummary.RefreshForm
        End If
        
        mSysNav.CreateGuruAutoTradeItems
        
        If Not g.TradingItems Is Nothing Then
            g.TradingItems.EnablementsChanged
        End If
    End If
    
    If FileExist(AddSlash(App.Path) & "Ftp\Purchase.OK") Then
        KillFile App.Path & "\Install.flg"
    End If
    KillFile App.Path & "\FTP\*.OK"
    
DebugLog "FtpRequest 10"

    ' If there is an error with the users account information extract the error
    ' and display it to the customer
    If frmStatus.Status <> eStatus_Aborted Then
        bSuccess = CheckAuthorization
        If bSuccess Then bSuccess = CheckData
        
'        If bSuccess Then
'            If CheckDataForError Then
'                frmStatus.Status = eStatus_Error
'            End If
'        End If
        
    End If
    
DebugLog "FtpRequest 11"

    ' If we had a successful download, and the SimTrade stuff is running, check to
    ' see if we have anything new for SimTrade...
    If bSuccess Then
        If Not g.SimTradeTs Is Nothing Then
            If g.SimTradeTs.UseSalmon = False Then
                ZipExecute "U", AddSlash(App.Path) & "FTP\Trades.GZP", AddSlash(App.Path) & "SimTrade\In", "*.TRD"
                ZipExecute "U", AddSlash(App.Path) & "FTP\Trades.GZP", AddSlash(App.Path) & "FTP", "*.TXT"
                
                frmOnlineBroker.tmrTradeServer.Enabled = True
            End If
        End If
        KillFile App.Path & "\FTP\Upload\*.*"
    End If
    
    CopyRecalcLog
    
    FtpRequest = bSuccess
    
DebugLog "FtpRequest 12"

ErrExit:
    Exit Function

ErrSection:
    If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    RaiseError "mDataNav.FtpRequest", eGDRaiseError_Raise
End Function

' TLB: started developing this in case we need to trigger uploading some info (but not yet very well tested)
Public Function FtpUploadCheck() As Boolean

If IsIDE Then
    On Error GoTo ErrSection:
Else
    On Error Resume Next
End If

    Dim i&, nInactive&, bRecursive As Boolean, bClearMsgForm As Boolean
    Dim dtStart As Date, dtEnd As Date, dtNow As Date
    Dim s$, strTriggerFile$, strUploadPath$, strZipFile$, strTime$
    Dim aFile As New cGdArray, aRequest As New cGdArray, aFlds As New cGdArray
    Static bInProgress As Boolean
    
    If g.bStarting Or g.bUnloading Or bInProgress Then Exit Function
    
    ' if no trigger file, then exit
    strTriggerFile = App.Path & "\Info\UZP.req"
    If Not FileExist(strTriggerFile) Then Exit Function
    aFile.FromFile strTriggerFile
    If aFile.Size = 0 Then Exit Function
    If ProcessIsBusy(True) Then Exit Function
          
    ' check if filtering on MachineID, DataServiceID, or Enablement code
    ' if not a match, then delete trigger file and exit
    For i = 0 To aFile.Size - 1
        aFlds.SplitFields aFile(i), vbTab
        Select Case UCase(aFlds(0))
        Case "MID"
            If InStr("," & UCase(aFlds(1)) & ",", "," & UCase(RI_GetMachineID) & ",") = 0 Then
                aRequest.Size = 0
                Exit For
            End If
        Case "DSID"
            If InStr("," & aFlds(1) & ",", "," & Str(RI_GetDataServiceID) & ",") = 0 Then
                aRequest.Size = 0
                Exit For
            End If
        Case "REQUIRED"
            If Not HasModule(aFlds(1), True) Then
                aRequest.Size = 0
                Exit For
            End If
        Case "TIME"
            strTime = aFlds(1)
        Case "INACTIVE"
            nInactive = Val(aFlds(1))
        Case "FILE", "DIR"
            aRequest.Add aFile(i)
        End Select
    Next
    If aRequest.Size = 0 Then
        KillFile strTriggerFile
        Exit Function
    End If
    
    ' check for valid time range and #minutes of inactivity (else just exit and wait for later)
    If nInactive > 0 Then
        If (gdTickCount - g.dLastMouseActivity) / 60000# < nInactive Then
            Exit Function
        End If
    End If
    If Len(strTime) > 0 Then
        dtStart = CDate(Parse(strTime, "-", 1))
        dtEnd = CDate(Parse(strTime, "-", 2))
        dtNow = Now - Date
        If dtEnd > dtStart Then
            If dtNow < dtStart Or dtNow > dtEnd Then
                Exit Function
            End If
        Else
            If dtNow < dtStart And dtNow > dtEnd Then
                Exit Function
            End If
        End If
    End If
    
    bInProgress = True
    KillFile strTriggerFile
    
    strUploadPath = App.Path & "\FTP\Upload\"
    MakeDir strUploadPath, False
    KillFile strUploadPath & "*.*"
    
    strZipFile = strUploadPath & RI_GetMachineID & ".gzp"
    For i = 0 To aRequest.Size - 1
        aFlds.SplitFields aRequest(i), vbTab
        Select Case UCase(aFlds(0))
        Case "FILE"
            If Val(aFlds(2)) = 0 Then
                bRecursive = False
            Else
                bRecursive = True
            End If
            ChangePath App.Path
            ZipExecute "A", strZipFile, "", aFlds(1), bRecursive
        Case "DIR"
            ChangePath App.Path
            aFile.GetMatchingFiles aFlds(1), True, False, True
            aFile.Add aFlds(1), 0
            aFile.Add ""
            aFile.ToFile strUploadPath & "Dir.txt", True
        End Select
    Next
    
    If FileExist(strUploadPath & "*.*") Then
        aRequest.Size = 0
        aRequest.Add "%NYTIME"
        frmStatus.Status = eStatus_Initialized
        If MsgForm Is Nothing Then
            Set MsgForm = frmStatus
            bClearMsgForm = True
        End If
        FtpUploadCheck = FtpRequest(aRequest)
        If bClearMsgForm Then
            Set MsgForm = Nothing
        End If
    End If
    
ErrExit:
    frmStatus.AddDetail "Finished"
    frmStatus.Status = eStatus_Completed
    frmStatus.Hide
    KillFile strUploadPath & "*.*"
    KillFile App.Path & "\FTP\Upload.GZP"
    bInProgress = False
    Exit Function

ErrSection:
    bInProgress = False
    If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    RaiseError "mDataNav.FtpUploadCheck", eGDRaiseError_Raise
    Resume ErrExit
End Function

Public Sub SendWebPage(ByVal strFileName$, ByVal strContents$)

    Dim i&, strUrl$, strPostData$
    Dim frmWeb As frmWebReport
    
    If Not IsIDE Then
        On Error Resume Next
    End If
    
    strUrl = FixURL(GetProvidedProperty("WebSend", "http://www.TradeNavigator.com/ClientUpload/Index.aspx?U=*&P=*"))
    If InStr(strUrl, ".") > 0 And Len(strFileName) > 0 Then
        ' encode the Filename in case it has ampersands or other special characters
        strFileName = UrlEncodeField(strFileName)
    
        ' move other args to PostData
        i = InStr(strUrl, "?")
        If i > 0 Then
            strPostData = Trim(Mid(strUrl, i + 1)) & "&"
            strUrl = Trim(Left(strUrl, i - 1))
        End If
        strPostData = strPostData & "filename=" & strFileName & "&contents=" & EncryptToHex(strContents)
        
        Set frmWeb = New frmWebReport
        frmWeb.ShowMe "", , strUrl, strPostData
        Set frmWeb = Nothing
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetNavwinIniValue
'' Description: Get a value from the Navwin.INI file
'' Inputs:      Variable to get
'' Returns:     Value of the variable
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNavwinIniValue(strPropName As String) As String
On Error GoTo ErrSection:
  
    GetNavwinIniValue = GetIniFileProperty(strPropName, "", "LOGIN", "navwin.ini")
  
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mDataNav.GetNavwinIniValue", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckReqCode
'' Description: Returns whether or not the customer is a valid user
'' Inputs:      Mode to call ReqCode with
'' Returns:     Ture if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckReqCode(strMode As String) As Boolean
On Error GoTo ErrSection:

    Dim strRightNow As String           ' Current Date/Time
    Dim strLine As String               ' Line from an input file
    Dim strCommandLine As String        ' Command line to send to shell
    Dim astrFile As New cGdArray        ' Input file
    
    ' Set up the command line for ReqCode
    strRightNow = ConvertDate(Now) & ".txt"
    Select Case strMode
        Case "DLD"
            strCommandLine = "ReqCode.exe" & Chr(34) & " DLD /D=30 /p=ETA /f=" & Chr(34) & App.Path & "\" & strRightNow & Chr(34)
        Case "SIM"
            strCommandLine = "ReqCode.exe" & Chr(34) & " SIM /D=30 /p=ETA /f=" & Chr(34) & App.Path & "\" & strRightNow & Chr(34)
    End Select

    ' Send the command to ReqCode
    Shell Chr(34) & App.Path & "\" & strCommandLine, vbNormalFocus  '"c:\progra~1\simutrade\reqcode.exe SIM /p=Simutrade /f=c:\progra~1\simutrade\"
    
    ' Wait for the response file to come back
    Do While Not FileExist(AddSlash(App.Path) & strRightNow)
        Sleep 0.1
    Loop
 
    ' Get the first line out of the file
    astrFile.FromFile AddSlash(App.Path) & strRightNow
    strLine = astrFile(0)
    KillFile AddSlash(App.Path) & strRightNow
 
    ' Post an appropriate message and return whether the user is valid or not
    Select Case UCase(Trim(strLine))
        Case "CANCELED BY USER."
            If strMode <> "DLD" Then MsgBox "You cannot use the software until you register"
            CheckReqCode = False
        Case "REQUEST MADE."
            If strMode <> "DLD" Then MsgBox "Thank you for submitting your registration information, you should receive your access code via email shortly"
            CheckReqCode = False
        Case "(C)GFDS"
            CheckReqCode = True
        Case Else
            If strMode <> "DLD" Then MsgBox "There was an error during registration please try again." & Chr(13) & Chr(10) & "Invalid response : " & strLine
            CheckReqCode = False
    End Select
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mDataNav.CheckReqCode", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOptionChain
'' Description: Requests an option chain from FRED for a particular symbol
'' Inputs:      Underlying Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetOptionChain(ByVal strSymbol$) As Boolean
On Error GoTo ErrSection:

    Dim strPath As String               ' FTP directory for the app
    Dim strError As String              ' Error message to display to the user
    Dim aRequest As New cGdArray
    Dim strTemp As String
    Dim SecType As eSYM_SecType
    
    frmStatus.Status = eStatus_Initialized
    frmStatus.AddDetail "Retrieving Option Chain"
    frmStatus.SetTitle "Retrieving Option Chain"
    
    strPath = AddSlash(App.Path) & "FTP\"
    
    ' First get the security type of the symbol passed in
    SecType = g.SymbolPool.SecType(g.SymbolPool.PoolRecForSymbol(strSymbol))
    
    ' Create the request file
    'aRequest.Add "@0-0;E;I;I;$IRX"
    aRequest.Add RequestLine("$IRX", "I")
    
    Select Case SecType
        Case eSYMType_Index
            'aRequest.Add "@0-0;E;I;I;" & UCase(strSymbol)
            aRequest.Add RequestLine(UCase(strSymbol), "I")
            If Left(strSymbol, 1) = "$" Then
                strTemp = Mid(strSymbol, 2) '(strip $ off index options)
            Else
                strTemp = strSymbol
            End If
            aRequest.Add "@0-0;E;I;SO;" & UCase(strTemp) & "-*/*"
        Case eSYMType_Stock
            'aRequest.Add "@0-0;E;I;S;" & UCase(strSymbol)
            aRequest.Add RequestLine(UCase(strSymbol), "S")
            aRequest.Add "@0-0;E;I;SO;" & UCase(strSymbol) & "-*/*"
        Case eSYMType_Future
            'aRequest.Add "@0-0;D;I;F;" & UCase(strSymbol)
            aRequest.Add RequestLine(UCase(strSymbol), "F")
            aRequest.Add "@0-0;D;I;FO;" & UCase(Parse(strSymbol, "-", 1)) & "-*/*"
    End Select
    
    ' Make the request
    'ShowForm frmDownloadStatus
    If Not FtpRequest(aRequest) Then
        KillFile strPath & "data.dat"
        frmStatus.Status = eStatus_Error
    End If
    
    If frmStatus.Status = eStatus_Completed Then
        
        ' Distribute the data (if data to distribute)
        If Not DistributeData("Distributing Data") Then
            'State.eCurStatus = 600
            frmStatus.Status = eStatus_Error
            frmStatus.AddDetail "ERROR downloading data"
        End If
    
        ' Distribute the data (if data to distribute)
        If frmStatus.Status = eStatus_Completed Then
            
            frmStatus.AddDetail "Final Updating"
            DM_DistribData ""
            frmStatus.Status = eStatus_Running
            
            ' Update any visible charts
            'UpdateVisibleCharts
            
            ' Refresh the grid
            'frmQuotes.TotalRefresh True
            
            frmStatus.Status = eStatus_Completed
            frmStatus.AddDetail "Finished"
            GetOptionChain = True
        End If
    End If
            
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mDataNav.GetOptionChain", eGDRaiseError_Raise

End Function

' Returns array of tab-delimited strings: OptionSymbol <tab> BidPrice <tab> AskPrice
' e.g. "AAPL 20140816 C67.86<tab>30.31<tab>30.6"
' - if no ExpirationDate specified, will return bids and asks for all expirations
' - if ExpirationDate specified, will return only for that expiration (and for stocks, the
'       ExpirationDate will automatically get bumped forward to the next valid weekly or monthly expiration)
Public Function GetOptionChainBidAskData(ByVal strSymbol$, Optional ByVal nExpirationDate& = 0) As cGdArray
On Error GoTo ErrSection:

    Dim i&, s$, strOptionSymbol$, strExpDate$, dBid#, dAsk#, strCheck$
    Dim strPath As String               ' FTP directory for the app
    Dim strError As String              ' Error message to display to the user
    Dim aRequest As New cGdArray
    Dim SecType As eSYM_SecType
    Dim aResults As New cGdArray
    Dim aFile As New cGdArray
    
    ' First get the security type of the symbol passed in
    SecType = g.SymbolPool.SecType(g.SymbolPool.PoolRecForSymbol(strSymbol))
    
    ' And get the expiration date
    If nExpirationDate > 0 Then
        nExpirationDate = DateOf(nExpirationDate)
        ' for stocks, move ExpDate to the Friday, except 3rd Friday should be moved to Saturday
        If SecType <> eSYMType_Future Then
            Do While Weekday(nExpirationDate) < vbFriday
                nExpirationDate = nExpirationDate + 1
            Loop
            i = GetDateFromRule(Year(nExpirationDate), Month(nExpirationDate), "3F")
            If nExpirationDate = i Then
                nExpirationDate = nExpirationDate + 1
            End If
        End If
        strExpDate = Format(nExpirationDate, "YYYYMMDD")
    End If
    
    frmStatus.Status = eStatus_Initialized
    frmStatus.AddDetail "Retrieving Options Data"
    frmStatus.SetTitle "Retrieving Options Data"
    
    strPath = AddSlash(App.Path) & "FTP\"
    
    ' Create the request file
    Select Case SecType
        Case eSYMType_Index
            If Left(strSymbol, 1) = "$" Then
                s = Mid(strSymbol, 2) '(strip $ off index options)
            Else
                s = strSymbol
            End If
            aRequest.Add "@0-0;E;I;SO;" & UCase(s) & "-*/*"
        Case eSYMType_Stock
            aRequest.Add "@0-0;E;I;SO;" & UCase(strSymbol) & "-*/*"
        Case eSYMType_Future
            aRequest.Add "@0-0;D;I;FO;" & UCase(Parse(strSymbol, "-", 1)) & "-*/*"
    End Select
    
    ' Make the request
    'ShowForm frmDownloadStatus
    If Not FtpRequest(aRequest) Then
        KillFile strPath & "data.dat"
        frmStatus.Status = eStatus_Error
    End If
    
    If frmStatus.Status = eStatus_Completed Then
        frmStatus.Status = eStatus_Completed
        frmStatus.AddDetail "Finished"
        
'@B/20140816/C22.5 (B)
'20140729 0.000000 0.000000 0.000000 0.000000 0 0 0 0 0000 11.900000 14.900000 52 111 0.000000 0.000000 0.000000 0.000000
'^ZB-201409/P145
'20140729 6.109375 6.109375 6.109375 6.109375 0 1 0 0 0000 6.015625 6.218750 12 12 0.000000 0.000000 0.000000 0.000000 1

        aFile.FromFile strPath & "data.dat"
        strOptionSymbol = ""
        strCheck = "@" & UCase(strSymbol) & "/"
        For i = 0 To aFile.Size - 1
            s = Trim(aFile(i))
            If Len(s) = 0 Then
                strOptionSymbol = ""
            ElseIf Not IsDigit(s, 1) Then
                ' get symbol, and check expiration date
                If SecType = eSYMType_Stock Then
                    If Left(s, Len(strCheck)) <> strCheck Then
                        s = "" ' ignore oddball symbols (e.g. AAPL7)
                    End If
                End If
                strOptionSymbol = s
                If Len(strExpDate) > 0 And Len(s) > 0 Then
                    s = Parse(s, "/", 2)
                    If s <> strExpDate Then
                        strOptionSymbol = "" ' ignore this one
                    End If
                End If
            ElseIf Len(strOptionSymbol) > 0 Then
                ' get bid and ask prices
                dBid = Val(Parse(s, " ", 11))
                dAsk = Val(Parse(s, " ", 12))
                If SecType <> eSYMType_Future Then
                    dBid = Round(dBid, 5)
                    dAsk = Round(dAsk, 5)
                End If
                ' put symbol together
                strOptionSymbol = Parse(Mid(strOptionSymbol, 2), "(", 1)
                strOptionSymbol = Replace(strOptionSymbol, "/", " ")
                s = strOptionSymbol & vbTab & Str(dBid) & vbTab & Str(dAsk)
                aResults.Add s
            End If
        Next
    End If
            
ErrExit:
    Set GetOptionChainBidAskData = aResults
    Exit Function

ErrSection:
    RaiseError "mDataNav.GetOptionChainBidAskData", eGDRaiseError_Raise

End Function

Public Function NewCustomObjectName(ByVal strExt$) As String
On Error GoTo ErrSection:

    Dim strPath$, strFile$, nHighest&, nNum&, i&
    Dim aFiles As New cGdArray
    
    strPath = App.Path & "\Custom\"
    If Left(strExt, 1) <> "." Then strExt = "." & strExt
    
    'look for highest existing number in directory
    aFiles.GetMatchingFiles strPath & "CUS*" & strExt, False
    For i = 0 To aFiles.Size - 1
        nNum = Val(Mid(aFiles(i), 4))
        If nNum > nHighest Then nHighest = nNum
    Next
    
    NewCustomObjectName = "Cus" & Format(nHighest + 1, "00000") & UCase(strExt)

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mDataNav.NewCustomObjectName", eGDRaiseError_Raise

End Function

' Returns the last date of data from a daily download
Public Function LastDailyDownload(Optional ByVal bRefresh As Boolean = False) As Date
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' End of day data for a symbol
    Dim lSymbolID As Long               ' Symbol ID for a symbol
    Dim lFromDate As Long
    Static lLastDate As Long
    
If 0 Then
    LastDailyDownload = DateSerial(2003, 12, 11)
    Exit Function
End If
    
    ' check for streaming replay mode
    If g.nReplaySession > 0 Then
        ' make negative to flag this simulated streaming mode
        LastDailyDownload = g.nReplaySession - 1
        lLastDate = 0 '(to force a refresh once simulated streaming is off)
        Exit Function
    End If
    
    If (lLastDate <= 0) Or bRefresh Then
        ' Get last download date from where stored
        ' (may be later than major symbols if an American holiday)
        lLastDate = Val(FileToString(DataPath & "LastDown.txt", 256))
        If lLastDate >= 20000101 And lLastDate <= 29990101 Then
            lLastDate = JulFromLong(lLastDate)
        Else
            ' look at data for major symbols ...
            lLastDate = 0
    
            ' see if $DJIA has later data: 50
            lSymbolID = 50 'g.SymbolPool.SymbolIDforSymbol("$DJIA")
            lFromDate = Date - 730
            If lLastDate > lFromDate Then lFromDate = lLastDate
            If DM_GetBars(Bars, lSymbolID, 0, lFromDate, , , , , False) Then
                If Bars.Size > 0 Then
                    lLastDate = Bars(eBARS_DateTime, Bars.Size - 1)
                End If
            End If
            
            ' see if IBM has later data: 11936
            lSymbolID = 11936 'g.SymbolPool.SymbolIDforSymbol("IBM")
            lFromDate = Date - 730
            If lLastDate >= lFromDate Then lFromDate = lLastDate + 1
            If DM_GetBars(Bars, lSymbolID, 0, lFromDate, , , , , False) Then
                If Bars.Size > 0 Then
                    lLastDate = Bars(eBARS_DateTime, Bars.Size - 1)
                End If
            End If
            
            ' see if SP-067 has later data: 41180
            lSymbolID = 41180 'g.SymbolPool.SymbolIDforSymbol("SP-067")
            lFromDate = Date - 730
            If lLastDate >= lFromDate Then lFromDate = lLastDate + 1
            If DM_GetBars(Bars, lSymbolID, 0, lFromDate, , , , , False) Then
                If Bars.Size > 0 Then
                    lLastDate = Bars(eBARS_DateTime, Bars.Size - 1)
                End If
            End If
        End If
    End If
    
    LastDailyDownload = lLastDate

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mDataNav.LastDailyDownload", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetUserInfo
'' Description: Retrieve the user information from the registry (or from the
''              INI file if the registry is blank)
'' Inputs:      None
'' Returns:     User Name, Password
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetUserInfo(strUserName As String, strPassword As String)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the registry

    ' Retrieve from ini file
    strUserName = Trim(GetIniFileProperty("FIRSTNAME", "", "LOGIN", "navwin.ini") & " " & GetIniFileProperty("LASTNAME", "", "LOGIN", "navwin.ini"))
    strPassword = Trim(GetIniFileProperty("PASSWORD", "", "LOGIN", "navwin.ini"))
    
    ' Retrieve from registry
    strKey = "Software\Genesis Financial Data Services\Account"
    strUserName = GetRegistryValue(rkLocalMachine, strKey, "UserName", strUserName)
    
    ' Password is encrypted in registry
    If Not IsDBCS Then '(VbEncrypt does not work under DBCS)
        strPassword = Left(strPassword + Space(25), 25)
        VbEncrypt strPasswordKey, strPassword, Len(strPassword)
        strPassword = GetRegistryValue(rkLocalMachine, strKey, "Password", strPassword, 100)
        VbEncrypt strPasswordKey, strPassword, Len(strPassword)
    End If
    
    strUserName = Trim(strUserName)
    strPassword = Trim(strPassword)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mDataNav.GetUserInfo", eGDRaiseError_Raise

End Sub

Public Sub SetUserInfo(ByVal pstrUserName As String, ByVal pstrPassword As String)
On Error GoTo ErrSection:

    Dim strKey As String
    Dim strPassword As String

    If IsDBCS Then Exit Sub '(VbEncrypt does not work under DBCS)

    strKey = "Software\Genesis Financial Data Services\Account"
    
    SetIniFileProperty "FIRSTNAME", Trim(Parse(pstrUserName, " ", 1)), "LOGIN", "NavWin.INI"
    SetIniFileProperty "LASTNAME", Trim(Parse(pstrUserName, " ", 2)), "LOGIN", "NavWin.INI"
    SetIniFileProperty "PASSWORD", Trim(pstrPassword), "LOGIN", "NavWin.INI"
    
    SetRegistryValue rkLocalMachine, strKey, "UserName", pstrUserName
    strPassword = Left(Trim(pstrPassword) + Space(25), 25)
    VbEncrypt strPasswordKey, strPassword, Len(strPassword)
    SetRegistryValue rkLocalMachine, strKey, "Password", strPassword, True
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mDataNav.SetUserInfo", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetRegisterProgram
'' Description: Downloads the Subscribe program via FTP
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRegisterProgram() As Boolean
On Error GoTo ErrSection:

    Dim fh As Integer                   ' File handle to the request file
    Dim strTemp As String               ' Temporary string variable
    Dim lReturn As Long                 ' Return value from shell
    Dim bSuccess As Boolean             ' Return value for the function

    ' Clean out the FTP directory
    If Not DirExist(AddSlash(App.Path) & "Ftp") Then MakeDir AddSlash(App.Path) & "Ftp"
    If Dir(App.Path & "\FTP\*.*") <> "" Then
        If Not DirExist(AddSlash(App.Path) & "Ftp\Temp") Then
            MakeDir AddSlash(App.Path) & "Ftp\Temp"
        Else
            KillFile AddSlash(App.Path) & "Ftp\Temp\*.*"
        End If
        
        If FileExist(AddSlash(App.Path) & "Ftp\*.*") Then FileCopy AddSlash(App.Path) & "Ftp\*.*", AddSlash(App.Path) & "Ftp\Temp\"
        KillFile AddSlash(App.Path) & "Ftp\*.*"
    End If
    
    ' Write out the IBIS.TXT file with the guest account
    fh = FreeFile
    Open App.Path & "\FTP\Ibis.TXT" For Output As #fh
    Print #fh, "Action=STRT"
    Print #fh, "UserName=ACCESS REQUEST"
    Print #fh, "Password=guest"
    Close #fh
    
    ' Write out the request file
    fh = FreeFile
    Open App.Path & "\FTP\Retrieve.TXT" For Output As #fh
    Print #fh, "<Subscrib.EXE"
    Print #fh, "+CRC:" & FileCrcString(App.Path & "\Subscribe.exe")
    Close #fh
    
    ' Make the request with gclient
    KillFile App.Path & "\GClient.can"
    frmStatus.Status = eStatus_Running
    frmStatus.AddDetail "Downloading Subscription Information"
    If SyncGclient Then '(get correct Gclient: FTP or HTTP)
        strTemp = Chr(34) & App.Path & "\GClientF.exe" & Chr(34) & " " & Chr(34) & "/s=" & App.Path & "\ftp" & Chr(34) & " /c=100 " & Chr(34) & "/d=" & App.Path & "\ftp" & Chr(34) & " /h=" & frmStatus.txtHwnd.hWnd & " /n"
    Else
        strTemp = Chr(34) & App.Path & "\GClient.exe" & Chr(34) & " " & Chr(34) & "/s=" & App.Path & "\ftp" & Chr(34) & " /c=100 " & Chr(34) & "/d=" & App.Path & "\ftp" & Chr(34) & " /h=" & frmStatus.txtHwnd.hWnd & " /n"
    End If
    DoEvents
    lReturn = Shell(strTemp, vbHide)
  
    ' Loop until done or error is received
    Do
        Sleep 0.1
        Select Case frmStatus.Status
            Case eStatus_Completed
                frmStatus.UpdateProgress ""
                bSuccess = True
                Exit Do
            Case eStatus_Aborted, eStatus_Error
                bSuccess = False
                Exit Do
        End Select
    Loop
    
    If bSuccess Then
        If FileExist(App.Path & "\FTP\Subscrib.EXE") Then FileCopy App.Path & "\FTP\Subscrib.EXE", App.Path & "\Subscribe.EXE"
        KillFile App.Path & "\FTP\*.*"
        If FileExist(App.Path & "\FTP\Temp\*.*") Then
            FileCopy App.Path & "\FTP\Temp\*.*", App.Path & "\FTP\"
            KillFile App.Path & "\FTP\Temp\*.*"
        End If
    End If
    
    If Not FileExist(App.Path & "\Subscribe.EXE") Then
        InfBox "h=Error ; i=! ; Failed to download the subscription|information.  Please try again."
    End If
    
    GetRegisterProgram = bSuccess
    
ErrExit:
    Exit Function

ErrSection:
    If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    RaiseError "mDataNav.GetRegisterProgram", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportSymbolGroup
'' Description: Exports a symbol group to a given format in a given path
'' Inputs:      Symbol Group ID (without the GRP:), Format, Path
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExportSymbolGroup(ByVal pstrSymbolGroupID$, ByVal pstrFormat$, ByVal pstrPath$)
On Error GoTo ErrSection:

    Dim lFieldNum As Long               ' Field number for the symbol group
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID for the current symbol
    Dim Bars As New cGdBars             ' Bars structure to hold the data

    ' Get the field number for the symbol group ID passed in
    lFieldNum = g.SymbolPool.FieldNumForID("GRP:" & pstrSymbolGroupID)
    
    ' Walk through the symbol pool
    For lIndex = 0 To g.SymbolPool.NumRecords - 1
        ' If the symbol is in the symbol group, export it
        If g.SymbolPool.ArrayTable(lFieldNum, lIndex) = 1 Then
            ' Get the symbol ID of the symbol to export
            lSymbolID = g.SymbolPool.SymbolID(lIndex)
            
            ' Get the bars from the Data Manager and export the data
            If DM_GetBars(Bars, lSymbolID) = True Then
                If Not DirExist(pstrPath) Then MakeDir pstrPath, False
                Bars.ToFile pstrFormat, pstrPath, g.SymbolPool.Symbol(lIndex), g.SymbolPool.Desc(lIndex), Bars.Prop(eBARS_ConvFactor)
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mDataNav.ExportSymbolGroup", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetNYTime
'' Description: Downloads the New York time from the FTP server
'' Inputs:      Purchase OK if want to know if the purchase is now ok
'' Returns:     String of the date/time in New York
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNYTime() As String
On Error GoTo ErrSection:
    
    Dim aFile As New cGdArray
    
    ' Make the request file to request the New York time
    aFile.Add "%NYTIME"
    aFile.Add "%VERIFY PORTFOLIO"
    
    ' Issue the request
    FtpRequest aFile
    
    ' Get the New York time out of the file
    aFile.FromFile App.Path & "\FTP\Request.INF"
    If aFile.Size > 0 Then
        GetNYTime = Parse(aFile(0), "=", 2)
    Else
        GetNYTime = ""
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mDataNav.GetNYTime", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AskForActivate
'' Description: Asks if the user would like to activate their data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AskForActivate(Optional ByVal bForceAsk As Boolean = False)
On Error GoTo ErrSection:

    Dim strMsg$

    ' if we don't have the true enablement codes (i.e. still the default), try to get them now
    Do While bForceAsk Or UCase(Left(g.strAuthorizationString, 9)) = ",DEFAULT,"
        strMsg = "Newly installed data or modules may need to be activated (this requires an internet connection)." & _
            "||Would you like to connect now?"
        If InfBox(strMsg, "?", "+Yes|-No", "Activation") = "N" Then
            Exit Do
        End If
        GetNYTime
        If UCase(Left(g.strAuthorizationString, 9)) <> ",DEFAULT," Then
            If frmStatus.Visible Then
                frmStatus.AddDetail "Activation successful"
                frmStatus.UpdateProgress "Finished"
            End If
            Exit Do
        End If
    Loop

#If 0 Then
' Commented out 11/2/2001 by DAJ for Hume People
'    If FileExist(App.Path & "\Hume.MOD") Then Exit Sub

    strMsg = "Newly installed data or modules may need to be activated (this requires an internet connection)." & _
        "||Would you like to connect now?"
    strReturn = InfBox(strMsg, "?", "+Yes|-No", "Activation")

    If strReturn = "N" Then
        Exit Sub
    Else
        GetNYTime
    End If

    If Not FileExist(App.Path & "\Install.flg") Then
        If frmStatus.Visible Then
            frmStatus.AddDetail "Activation successful"
            frmStatus.UpdateProgress "Finished"
        End If
        InfBox "Your data has been successfully activated.", "i", , "Activation"
        ' update charts to show newly activated data
        UpdateVisibleCharts
    Else
        If frmStatus.Visible Then
            frmStatus.AddDetail "Activation failed"
            frmStatus.UpdateProgress "ERROR"
        End If
        InfBox "Genesis was unable to activate|your data at this time.", "i", , "Activation"
    End If
#End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mDataNav.AskForActivate", eGDRaiseError_Raise
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFutOptChain
'' Description: Downloads a future option chain
'' Inputs:      Underlying symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFutOptChain(ByVal strUnderlying As String)
On Error GoTo ErrSection:

    Dim strDate As String               ' Date to request from
    Dim aRequest As New cGdArray        ' Array of ftp requests

    If ProcessIsBusy Then Exit Sub

    ' Make sure that the ftp directory exists
    If Not DirExist(App.Path & "\ftp") Then MakeDir App.Path & "\ftp"
    
    ' Initialize the Status form
    frmStatus.Status = eStatus_Initialized
    frmStatus.AddDetail "Retrieving Option Chain"
    frmStatus.SetTitle "Retrieving Option Chain"
    
    ' Clean out the ftp directory
    If Dir(App.Path & "\ftp\*.*") <> "" Then KillFile App.Path & "\ftp\*.*", True

    ' Set up the request
    strDate = Format(Date, "yyyymmdd")
    aRequest.Add "@" & strDate & "-" & strDate & ";D;I;F;" & strUnderlying & "*"
    aRequest.Add "@" & strDate & "-" & strDate & ";D;I;FO;" & strUnderlying & "-*/*"
 
    ' Make the ftp request
    If Not FtpRequest(aRequest) Then
        frmStatus.Status = eStatus_Error
    End If

    If frmStatus.Status < eStatus_Aborting Or frmStatus.Status = eStatus_Completed Then
    
        If Not DistributeData Then
            frmStatus.Status = eStatus_Error
            frmStatus.AddDetail "ERROR downloading data"
        End If
        
        If frmStatus.Status = eStatus_Completed Then
            DM_DistribData ""
            
            frmStatus.Status = eStatus_Completed
            frmStatus.AddDetail "Finished"
        End If
    
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mDataNav.GetFutOptChain", eGDRaiseError_Raise

End Sub

Public Function FileCrcString(ByVal strFile$) As String
On Error GoTo ErrSection:

    Dim lCRC&, lFileSize&, dFileDate#, strFileDate$
    
    ' get stats for file
    If FileExist(strFile) Then
        lCRC = CalcFileCrc(strFile)
        lFileSize = FileLength(strFile)
        dFileDate = FileDate(strFile)
        If dFileDate > 0 Then
            strFileDate = Format(dFileDate, "yyyymmdd HH:MM:SS")
        End If
    End If
    
    ' strip off path
    strFile = Right(strFile, Len(strFile) - Len(FilePath(strFile)))
    
    ' build string
    FileCrcString = strFile & "," & Str(lCRC) & "," & strFileDate & "," & Str(lFileSize)
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mDataNav.FileCrcString", eGDRaiseError_Raise

End Function

Private Function RequestLine(ByVal strSymbol As String, ByVal strSecType As String)
On Error GoTo ErrSection:

    Dim strDates As String              ' Dates to request data for
    Dim strTemp As String               ' Symbol and security type
    Dim lIndex As Long                  ' Index into a for loop
    Dim lDelay As Long                  ' Delay for the symbol (if real-time)

    strDates = "@" & Format(Date, "YYYYMMDD") & "-" & Format(Date + 1, "YYYYMMDD")
    strTemp = strSecType & ";" & strSymbol
    
    ' If real-time is active, then get the delay
    If g.RealTime.Active Then
        lDelay = g.RealTime.SymbolDelay(strSymbol)
    Else
        lDelay = -1&
    End If
    
    ' If they are authorized for ticks, ask for both...
    If InStr(g.strAuthorizationString, "," & Left(strSecType, 1) & "T,") > 0 Then
        RequestLine = strDates & ";B;I;" & strTemp & ";" & Str(lDelay)
        
    ' Otherwise just ask for end of day...
    Else
        RequestLine = strDates & ";E;I;" & strTemp & ";" & Str(lDelay)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.RequestLine", eGDRaiseError_Raise
    
End Function

Public Sub ConvertQBF()
On Error GoTo ErrSection:

    Dim astrFiles As New cGdArray       ' List of QBF files in Custom directory
    Dim QBF As New cCriteria            ' QBF from the file
    Dim strFields As String             ' Quote Board fields from the ini file
    Dim lIndex As Long                  ' Index into a for loop
    Dim Criteria As New cCriteria       ' Criteria from the pool
    Dim bFound As Boolean               ' Name was found in existing criteria
    
    ' Get a list of QBF files in the custom directory...
    If astrFiles.GetMatchingFiles(AddSlash(g.strAppPath) & "Custom\*.QBF", False) > 0 Then
        ' Get the list of fields in the quote board from the ini file...
        strFields = GetIniFileProperty("DisplayFields", "", "QuoteList", g.strIniFile)
        
        ' Walk through the list of QBF files...
        For lIndex = 0 To astrFiles.Size - 1
            ' If the QBF file was being used, we need to convert it to a criteria...
            If InStr(UCase(strFields), UCase(astrFiles(lIndex))) > 0 Then
                Set QBF = New cCriteria
                If QBF.FromFile(AddSlash(g.strAppPath) & "Custom", UCase(astrFiles(lIndex))) Then
                    QBF.ID = ""
                    QBF.UsageType = eCriteria_FilterCriteria
                    QBF.IsActive = False
                    
                    ' If the name of the QBF is already used, append a (QBF)...
                    bFound = False
                    For Each Criteria In g.SymbolPool.Criterias
                        If Criteria.Name = QBF.Name Then
                            bFound = True
                            Exit For
                        End If
                    Next Criteria
                    If bFound Then QBF.Name = QBF.Name & " (QBF)"
                    
                    QBF.Save True
                    
                    ' Change the field in the ini file string...
                    strFields = Replace(strFields, UCase(astrFiles(lIndex)), QBF.ID)
                End If
            End If
            
            ' Delete the QBF file...
            KillFile AddSlash(g.strAppPath) & "Custom\" & astrFiles(lIndex), True
        Next lIndex
        
        ' Write the ini file string back out...
        SetIniFileProperty "DisplayFields", strFields, "QuoteList", g.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.ConvertQBF", eGDRaiseError_Raise
    
End Sub

' Pass: true to set to FTP, false to set to HTP, nothing to just get return value
' Returns: true if using FTP, false if using HTP
Public Function SyncGclient(Optional ByVal bUseFTP As Variant) As Boolean
On Error GoTo ErrSection

    Dim strSource$, strDest$, strIniFile$

bUseFTP = False
    If 0 Then ' IsMissing(bUseFTP) Then
        ' if nothing passed in, then just use current setting
        strIniFile = AddSlash(g.strAppPath) & "GClient.INI"
        Select Case GetIniFileProperty("UseFTP", 0, "Mode", strIniFile)
        Case 2
            bUseFTP = True
        Case 1
            ' see if should revert back to "Try HTTP first"
            If Int(CDbl(Now)) <> Int(GetIniFileProperty("WhenSetFTP", 0#, "Mode", strIniFile)) Then
                bUseFTP = False
                SetIniFileProperty "UseFTP", 0, "Mode", strIniFile
            Else
                bUseFTP = True
            End If
        Case Else
            bUseFTP = False
        End Select
    End If

    ' kill off any current Gclient's that might be running
    If KillProcess("Lil'Fred") Then
        Sleep 5
    End If
    
    ' sync with the correct file
    If bUseFTP Then
        strSource = AddSlash(g.strAppPath) & "GClient.FTP"
        strDest = AddSlash(g.strAppPath) & "GClientF.EXE"
        SyncGclient = True '(is using FTP)
    Else
        strSource = AddSlash(g.strAppPath) & "GClient.HTP"
        strDest = AddSlash(g.strAppPath) & "GClient.EXE"
        SyncGclient = False '(is using HTTP)
    End If
    If FileExist(strSource) Then
        If FileDate(strSource) <> FileDate(strDest) Then
            FileCopy strSource, strDest, True
        End If
        ' also sync the Gclient in the ETA folder
        strDest = AddSlash(g.strAppPath) & "..\Eta\GClient.EXE"
        If FileExist(strDest) Then
            If FileDate(strSource) <> FileDate(strDest) Then
                FileCopy strSource, strDest, True
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.SyncGclient", eGDRaiseError_Raise

End Function

Public Function IsForex(ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    strSymbol = Trim(strSymbol)
    If Left(strSymbol, 1) = "$" Then
        If Mid(strSymbol, 5, 1) = "-" Then
            If Len(strSymbol) = 8 Or Mid(strSymbol, 9, 1) = "@" Then
                IsForex = True
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.IsForex"
End Function

' Return true if this is a "spread symbol" (e.g. HE-201302-S1)
' (this should move to mDataNav when it's available)
Public Function IsSpreadSymbol(ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim i&

    If Len(strSymbol) >= 11 Then
        If Left(strSymbol, 1) <> "$" Then
            ' look for first dash
            i = InStr(2, strSymbol, "-")
            If i > 0 Then
                ' look for a "-S" after the contract year/month
                If UCase(Mid(strSymbol, i + 7, 2)) = "-S" Then
                    IsSpreadSymbol = True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.IsSpreadSymbol"
End Function

' Returns the roll symbol (individual contract) for a specific date
' (can pass -999 for session date in order to clear cache -- e.g. during daily download)
Public Function RollSymbolForDate(ByVal Symbol As Variant, Optional ByVal dSessionDate As Double = 999999) As String
On Error GoTo ErrSection:

    Dim iRec&, dLDL#, strSymbol$, bFound As Boolean
    Dim RollsTable As cGdTable
    Static dPrevLDL As Double
    Static CurrentContracts As cGdArray ' to provide a fast lookup table for current roll symbol
    
    If dSessionDate = -999 Then
        ' just clear the prev LDL so the cache will get rebuilt on next call
        dPrevLDL = 0
        Exit Function
    End If
    
    ' only need to perform this if symbol is a continuous contract
    strSymbol = GetSymbol(Symbol)
    RollSymbolForDate = strSymbol
    If IsAlpha(strSymbol, 1) And InStr(strSymbol, "-0") > 0 Then
        ' whenever LastDailyDownload has changed, clear out the fast lookup table
        dLDL = LastDailyDownload
        If dLDL <> dPrevLDL Or dPrevLDL = 0 Then
            dPrevLDL = dLDL
            Set CurrentContracts = New cGdArray
            CurrentContracts.Create eGDARRAY_Strings, 0
        End If
        
        ' if Date > LastDailyDownload then adjust it (so will work correctly for Streaming Replay)
        dSessionDate = Int(dSessionDate)
        If dSessionDate > dLDL Then
            dSessionDate = dLDL + 1
            ' see if already in fast lookup table
            If CurrentContracts.BinarySearch(strSymbol & vbTab, iRec, eGdSort_MatchUsingSearchStringLength) Then
                RollSymbolForDate = Parse(CurrentContracts(iRec), vbTab, 2)
                bFound = True
            End If
        End If
        
        If Not bFound Then
            ' load the rolls table and lookup the symbol
            Set RollsTable = GetRollsTable(Symbol)
            If RollsTable.NumRecords > 0 Then
                ' make sure we bump up a weekend date to the next Monday
                Do While Not IsWeekday(dSessionDate)
                    dSessionDate = dSessionDate + 1
                Loop
                For iRec = RollsTable.NumRecords - 1 To 0 Step -1
                    If RollsTable.Num(1, iRec) <= dSessionDate Then Exit For
                Next
                If iRec < 0 Then iRec = 0
                RollSymbolForDate = GetSymbol(RollsTable.Num(0, iRec))
            End If
            
            ' add to the fast lookup table (if it's a current roll)
            If dSessionDate > dLDL Then
                If Not CurrentContracts.BinarySearch(strSymbol & vbTab, iRec, eGdSort_MatchUsingSearchStringLength) Then
                    CurrentContracts.Add strSymbol & vbTab & RollSymbolForDate, iRec
                End If
            End If
        End If
    End If

ErrExit:
    Set RollsTable = Nothing
    Exit Function
    
ErrSection:
    Set RollsTable = Nothing
    RaiseError "mDataNav.RollSymbolForDate", eGDRaiseError_Raise
    
End Function

Public Sub UpdateGenTick()

'Config file:
'BaseSym  GenTick  StartTime(ET)  EndTime(ET)
'TQ US 820 1310
'$DJIA $DJ 930 1615

    Dim i&, nSym&, nPos&, nBar&, fh%
    Dim nSymbolID&, nDate&, nPoolRec&
    Dim nMonth&, nYear&, nStartTime&, nEndTime&, dTime#
    Dim strSymbol$, strBase$, strTickPath$
    Dim aGT As New cGdArray, aTickDist As New cGdArray
    Dim Ticks As New cGdBars
    
    If Not FileExist(App.Path & "\GenTick.exe") Then Exit Sub
    aGT.FromFile App.Path & "\GenTick.EXP"
    aTickDist.FromFile App.Path & "\Data\TickDist.LST"
    For i = 0 To aTickDist.Size - 1
        If Len(Trim(aTickDist(i))) > 0 Then Exit For
        aTickDist.Remove 0
    Next
    If aGT.Size = 0 Or aTickDist.Size = 0 Then Exit Sub
    
    frmStatus.AddDetail "Updating GenTick files"
    KillFile App.Path & "\Data\TickDist.LST"
    
    If InStr(aGT(0), "\") > 0 Then
        strTickPath = aGT(0)
        aGT.Remove 0
    Else
        strTickPath = "c:\gd\tick"
    End If
    
    For i = aGT.Size - 1 To 0 Step -1
        If Len(Trim(aGT(i))) = 0 Then
            aGT.Remove i
        ElseIf InStr(aGT(i), vbTab) = 0 Then
            aGT(i) = Trim(aGT(i)) & vbTab
        End If
    Next
    
    aGT.Sort eGdSort_IgnoreCase Or eGdSort_DeleteNullValues
    aTickDist.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
    
    For nSym = 0 To aTickDist.Size - 1
        nSymbolID = Val(Parse(aTickDist(nSym), vbTab, 1))
        nDate = Val(Parse(aTickDist(nSym), vbTab, 2))
        If nSymbolID > 0 And nDate > Date - 366 Then
            nPoolRec = g.SymbolPool.PoolRecForSymbolID(nSymbolID)
            strSymbol = g.SymbolPool.Symbol(nPoolRec)
If strSymbol = "$DJIA" Or Left(strSymbol, 3) = "TQ-" Then
    nDate = nDate
End If
            strBase = ""
            If InStr(strSymbol, " ") > 0 Then
                strBase = "" '(don't do options)
            ElseIf Left(strSymbol, 1) = "$" Then
                strBase = strSymbol
            ElseIf InStr(strSymbol, "-") > 0 Then
                nMonth = Val(Parse(strSymbol, "-", 2))
                nYear = nMonth / 100
                nMonth = nMonth Mod 100
                If nMonth >= 1 And nMonth <= 12 Then
                    strBase = Parse(strSymbol, "-", 1)
                Else
                    strBase = ""
                End If
            End If
            If Len(strBase) > 0 Then
                aGT.BinarySearch strBase & vbTab, nPos, eGdSort_IgnoreCase
                If Left(aGT(nPos), Len(strBase) + 1) = strBase & vbTab Then
                    If DM_GetBars(Ticks, nSymbolID, ePRD_EachTick, nDate, nDate) Then
                        UnminutizeTicks Ticks
                        If fh = 0 Then
                            fh = FreeFile
                            Open App.Path & "\GenTick.asc" For Output As #fh
                        End If
                        ' convert symbol to GenTick:  *US_03Z 20030910
                        strSymbol = Parse(aGT(nPos), vbTab, 2)
                        If Len(strSymbol) = 0 Then
                            strSymbol = strBase
                        End If
                        If Left(strSymbol, 1) <> "$" Then
                            strSymbol = Left(strSymbol & "___", 3) & Format(nYear Mod 100, "00")
                            Select Case nMonth 'FGHJKMNQUVXZ
                            Case 1: strSymbol = strSymbol & "F"
                            Case 2: strSymbol = strSymbol & "G"
                            Case 3: strSymbol = strSymbol & "H"
                            Case 4: strSymbol = strSymbol & "J"
                            Case 5: strSymbol = strSymbol & "K"
                            Case 6: strSymbol = strSymbol & "M"
                            Case 7: strSymbol = strSymbol & "N"
                            Case 8: strSymbol = strSymbol & "Q"
                            Case 9: strSymbol = strSymbol & "U"
                            Case 10: strSymbol = strSymbol & "V"
                            Case 11: strSymbol = strSymbol & "X"
                            Case 12: strSymbol = strSymbol & "Z"
                            End Select
                        End If
                        Print #fh, "*" & strSymbol & " " & Str(JulToLong(nDate, True))
                        ' get start and end times
                        nStartTime = Val(Parse(aGT(nPos), vbTab, 3))
                        nEndTime = Val(Parse(aGT(nPos), vbTab, 4))
                        If nEndTime = 0 Then nEndTime = 1630
                        nStartTime = Int(nStartTime / 100) * 60 + nStartTime Mod 100
                        nEndTime = Int(nEndTime / 100) * 60 + (nEndTime Mod 100)
                        ' check each tick
                        For nBar = 0 To Ticks.Size - 1
                            dTime = Ticks(eBARS_DateTime, nBar)
                            dTime = Int((dTime - Int(dTime)) * 1440 + 0.5)
                            If dTime >= nStartTime And dTime <= nEndTime Then
                                ' convert time to HHMM:  0820 106.531
                                dTime = Int(dTime / 60) * 100 + (dTime Mod 60)
                                Print #fh, Format(dTime, "0000") & " " & Str(Ticks(eBARS_Close, nBar))
                            End If
                        Next
                    End If
                End If
            End If
        End If
    Next
    
    If fh <> 0 Then
        Close #fh
        RunProcess App.Path & "\Gentick.exe", "g Gentick.asc " & strTickPath, False, vbHide
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CompareBars
'' Description: Output a list of differences between two sets of bars
'' Inputs:      Two sets of Bars to compare
'' Returns:     Array of differences
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CompareBars(ByVal Bars1 As cGdBars, ByVal Bars2 As cGdBars, _
        Optional ByVal strDesc1 As String = "Bars1", Optional ByVal strDesc2 As String = "Bars2", _
        Optional ByVal bDiffsFile As Boolean = True, Optional ByVal bUseDM As Boolean = True) As cGdArray
On Error GoTo ErrSection:

    Dim astrOutput As New cGdArray      ' Array of output information
    Dim astrLine As New cGdArray        ' Line of output information
    Dim lFromDate As Long               ' Date to test from
    Dim lToDate As Long                 ' Date to test to
    Dim lDate As Long                   ' Index into a for loop
    Dim dValue1 As Double               ' Value from Bars1
    Dim dValue2 As Double               ' Value from Bars2
    Dim lIndex1 As Long                 ' Index into Bars1
    Dim lIndex2 As Long                 ' Index into Bars2
    Dim strFirstChar As String          ' First character in output file
    Dim strMJKSymbol As String          ' MJK version of the symbol
    Dim bPriceDiff As Boolean           ' Were any prices different?
    Dim bVolDiff As Boolean             ' Were any vol/oi different?

    ' Create the string arrays...
    astrOutput.Create eGDARRAY_Strings
    astrLine.Create eGDARRAY_Strings

    ' Compare start dates and determine where to start test...
    dValue1 = Bars1(eBARS_DateTime, 0)
    dValue2 = Bars2(eBARS_DateTime, 0)
    If dValue1 < dValue2 Then
        If bDiffsFile Then astrOutput.Add strDesc1 & " StartDate=" & DateFormat(dValue1) & " ;" & strDesc2 & " StartDate=" & DateFormat(dValue2)
        lFromDate = CLng(dValue2)
    ElseIf dValue1 > dValue2 Then
        If bDiffsFile Then astrOutput.Add strDesc1 & " StartDate=" & DateFormat(dValue1) & " ;" & strDesc2 & " StartDate=" & DateFormat(dValue2)
        lFromDate = CLng(dValue1)
    Else
        lFromDate = CLng(dValue1)
    End If

    ' Compare end dates and determine where to end test...
    dValue1 = Bars1(eBARS_DateTime, Bars1.Size - 1)
    dValue2 = Bars2(eBARS_DateTime, Bars2.Size - 1)
    If dValue1 < dValue2 Then
        If bDiffsFile Then astrOutput.Add strDesc1 & " EndDate=" & DateFormat(dValue1) & " ;" & strDesc2 & " EndDate=" & DateFormat(dValue2)
        lToDate = CLng(dValue1)
    ElseIf dValue1 > dValue2 Then
        If bDiffsFile Then astrOutput.Add strDesc1 & " EndDate=" & DateFormat(dValue1) & " ;" & strDesc2 & " EndDate=" & DateFormat(dValue2)
        lToDate = CLng(dValue2)
    Else
        lToDate = CLng(dValue1)
    End If
    
    ' Walk through all of the dates and do comparisons...
    For lDate = lFromDate To lToDate
        bPriceDiff = False
        bVolDiff = False
    
        lIndex1 = Bars1.FindDateTime(lDate, True)
        lIndex2 = Bars2.FindDateTime(lDate, True)
        
        If lIndex1 >= 0 And lIndex2 = -1 Then
            If bDiffsFile Then astrOutput.Add strDesc2 & " is missing " & DateFormat(lDate)
        ElseIf lIndex1 = -1 And lIndex2 >= 0 Then
            If bDiffsFile Then astrOutput.Add strDesc1 & " is missing " & DateFormat(lDate)
        ElseIf lIndex1 >= 0 And lIndex2 >= 0 Then
            astrLine.Clear
            If Bars1(eBARS_Open, lIndex1) <> Bars2(eBARS_Open, lIndex2) Then
                astrLine.Add "Open (" & strDesc1 & ": " & Str(Bars1(eBARS_Open, lIndex1)) & "; " & strDesc2 & ": " & Str(Bars2(eBARS_Open, lIndex2)) & ")"
                bPriceDiff = True
            End If
            If Bars1(eBARS_High, lIndex1) <> Bars2(eBARS_High, lIndex2) Then
                astrLine.Add "High (" & strDesc1 & ": " & Str(Bars1(eBARS_High, lIndex1)) & "; " & strDesc2 & ": " & Str(Bars2(eBARS_High, lIndex2)) & ")"
                bPriceDiff = True
            End If
            If Bars1(eBARS_Low, lIndex1) <> Bars2(eBARS_Low, lIndex2) Then
                astrLine.Add "Low (" & strDesc1 & ": " & Str(Bars1(eBARS_Low, lIndex1)) & "; " & strDesc2 & ": " & Str(Bars2(eBARS_Low, lIndex2)) & ")"
                bPriceDiff = True
            End If
            If Bars1(eBARS_Close, lIndex1) <> Bars2(eBARS_Close, lIndex2) Then
                astrLine.Add "Close (" & strDesc1 & ": " & Str(Bars1(eBARS_Close, lIndex1)) & "; " & strDesc2 & ": " & Str(Bars2(eBARS_Close, lIndex2)) & ")"
                bPriceDiff = True
            End If
            If Bars1(eBARS_ContVol, lIndex1) <> Bars2(eBARS_ContVol, lIndex2) Then
                astrLine.Add "ContVol (" & strDesc1 & ": " & Str(Bars1(eBARS_ContVol, lIndex1)) & "; " & strDesc2 & ": " & Str(Bars2(eBARS_ContVol, lIndex2)) & ")"
                bVolDiff = True
            End If
            If Bars1(eBARS_ContOI, lIndex1) <> Bars2(eBARS_ContOI, lIndex2) Then
                astrLine.Add "ContOI (" & strDesc1 & ": " & Str(Bars1(eBARS_ContOI, lIndex1)) & "; " & strDesc2 & ": " & Str(Bars2(eBARS_ContOI, lIndex2)) & ")"
                bVolDiff = True
            End If
            If Bars1(eBARS_Vol, lIndex1) <> Bars2(eBARS_Vol, lIndex2) Then
                astrLine.Add "TotVol (" & strDesc1 & ": " & Str(Bars1(eBARS_Vol, lIndex1)) & "; " & strDesc2 & ": " & Str(Bars2(eBARS_Vol, lIndex2)) & ")"
                bVolDiff = True
            End If
            If Bars1(eBARS_OI, lIndex1) <> Bars2(eBARS_OI, lIndex2) Then
                astrLine.Add "TotOI (" & strDesc1 & ": " & Str(Bars1(eBARS_OI, lIndex1)) & "; " & strDesc2 & ": " & Str(Bars2(eBARS_OI, lIndex2)) & ")"
                bVolDiff = True
            End If
            
            If astrLine.Size > 0 Then
                If bDiffsFile Then
                    astrOutput.Add DateFormat(lDate) & ": " & astrLine.JoinFields(", ")
                ElseIf bVolDiff And Not bPriceDiff Then
                    If bUseDM Then
                        astrOutput.Add Format(lDate, "YYYYMMDD") & " " & Str(Bars1(eBARS_Open, lIndex1)) & " " & Str(Bars1(eBARS_High, lIndex1)) & " " & Str(Bars1(eBARS_Low, lIndex1)) & " " & Str(Bars1(eBARS_Close, lIndex1)) & " " & Str(Bars1(eBARS_ContVol, lIndex1)) & " " & Str(Bars1(eBARS_ContOI, lIndex1)) & " " & Str(Bars1(eBARS_Vol, lIndex1)) & " " & Str(Bars1(eBARS_OI, lIndex1)) & " 0"
                    Else
                        astrOutput.Add Format(lDate, "YYYYMMDD") & " " & Str(Bars2(eBARS_Open, lIndex2)) & " " & Str(Bars2(eBARS_High, lIndex2)) & " " & Str(Bars2(eBARS_Low, lIndex2)) & " " & Str(Bars2(eBARS_Close, lIndex2)) & " " & Str(Bars2(eBARS_ContVol, lIndex2)) & " " & Str(Bars2(eBARS_ContOI, lIndex2)) & " " & Str(Bars2(eBARS_Vol, lIndex2)) & " " & Str(Bars2(eBARS_OI, lIndex2)) & " 0"
                    End If
                End If
            End If
        End If
    Next lDate
    
    If astrOutput.Size = 0 And bDiffsFile Then
        astrOutput.Add "No differences"
    ElseIf astrOutput.Size > 0 And Not bDiffsFile Then
        Select Case Bars1.SecurityType
            Case "F"
                strFirstChar = "#"
                strMJKSymbol = Replace(Bars1.Prop(eBARS_Symbol), "-", "/")
                strMJKSymbol = Replace(strMJKSymbol, "/0", "/99")
                
            Case "S"
                strFirstChar = "!"
                strMJKSymbol = Bars1.Prop(eBARS_Symbol)
            
            Case "I"
                strFirstChar = "$"
                strMJKSymbol = Bars1.Prop(eBARS_Symbol)
            
            Case "M"
                strFirstChar = "~"
                strMJKSymbol = Bars1.Prop(eBARS_Symbol)
            
        End Select
        astrOutput.Add strFirstChar & strMJKSymbol, 0
    End If
    
    Set CompareBars = astrOutput

ErrExit:
    Set astrOutput = Nothing
    Set astrLine = Nothing
    Exit Function
    
ErrSection:
    Set astrOutput = Nothing
    Set astrLine = Nothing
    RaiseError "mDataNav.CompareBars", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CompareDataDirectory
'' Description: Compare the data files in a directory to the data manager
'' Inputs:      Path of the data files, Format of the data files
'' Returns:     Array of output information
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CompareDataDirectory(ByVal strPath As String, ByVal strFormat As String, _
        Optional ByVal bDiffsFile As Boolean = True, Optional ByVal bUseDM As Boolean = True) As cGdArray
On Error GoTo ErrSection:

    Dim astrOutput As New cGdArray      ' Array of output information
    Dim astrReturn As New cGdArray      ' Array returned from CompareBars
    Dim astrFiles As New cGdArray       ' Array of files to check
    Dim lFile As Long                   ' Index into a for loop
    Dim ExtBars As New cGdBars          ' Bars containing the external data
    Dim DmBars As New cGdBars           ' Bars containing data manager data

    ' Create string arrays...
    astrOutput.Create eGDARRAY_Strings
    astrReturn.Create eGDARRAY_Strings
    astrFiles.Create eGDARRAY_Strings
    
    ' Get the list of files in the given path...
    Select Case UCase(strFormat)
        Case "CSI"
            astrFiles.GetMatchingFiles AddSlash(strPath) & "*.DTA /s"
        Case "MS7"
            astrFiles.GetMatchingFiles AddSlash(strPath) & "*.DAT /s"
    End Select
    astrFiles.Sort
    
    ' Compare each file with it's data manager counterpart...
    For lFile = 0 To astrFiles.Size - 1
        ExtBars.FromFile strFormat, FilePath(astrFiles(lFile)), FileBase(astrFiles(lFile)) & "." & FileExt(astrFiles(lFile))
        DM_GetBars DmBars, Right(ExtBars.Prop(eBARS_Symbol), Len(ExtBars.Prop(eBARS_Symbol)) - 1), "Daily"
        
        If bDiffsFile Then astrOutput.Add strFormat & ": " & ExtBars.Prop(eBARS_Symbol) & ", DM: " & DmBars.Prop(eBARS_Symbol)
        If DmBars.Size = 0 Then
            If bDiffsFile Then astrOutput.Add "No data in Data Manager"
        Else
            Set astrReturn = CompareBars(ExtBars, DmBars, strFormat, "DM", bDiffsFile, bUseDM)
            astrOutput.AppendFromArray astrReturn
        End If
        If bDiffsFile Then astrOutput.Add ""
    Next lFile
    
    Set CompareDataDirectory = astrOutput

ErrExit:
    Set astrOutput = Nothing
    Set astrReturn = Nothing
    Set astrFiles = Nothing
    Set ExtBars = Nothing
    Set DmBars = Nothing
    Exit Function
    
ErrSection:
    Set astrOutput = Nothing
    Set astrReturn = Nothing
    Set astrFiles = Nothing
    Set ExtBars = Nothing
    Set DmBars = Nothing
    RaiseError "mDataNav.CompareDataDirectory", eGDRaiseError_Raise
    
End Function

Public Function IsMinutized(ByVal Bars As cGdBars, bFullTicksAvail As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim strFile As String
    Dim aFullTicks As cGdArray

    IsMinutized = False
    bFullTicksAvail = False
    
    If Bars.IsActiveArray(eBARS_DownTicks) And Bars.IsActiveArray(eBARS_UpTicks) Then
        If Bars(eBARS_DownTicks, 0) <> gdNullValue(Bars.ArrayHandle(eBARS_DownTicks)) Then
            If Bars(eBARS_UpTicks, 0) <> gdNullValue(Bars.ArrayHandle(eBARS_UpTicks)) Then
                IsMinutized = True
            End If
        End If
    End If

    strFile = g.strAppPath & "\provided\fullticks.flg"
    If FileExist(strFile) Then
        Set aFullTicks = New cGdArray
        aFullTicks.FromFile (strFile)
        If aFullTicks.Size > 0 Then
            If UCase(Left(aFullTicks(0), 3)) = "YES" Then
                bFullTicksAvail = True
            End If
        End If
    End If
    
    Set aFullTicks = Nothing

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.IsMinutized", eGDRaiseError_Raise
    
End Function

Public Function DownloadTicks(ByVal Bars As cGdBars, ByVal dSessionDate As Double) As Boolean
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol for the symbol ID
    Dim strSecType As String            ' Security Type for the symbol ID
    Dim astrRequest As cGdArray         ' Request to send
    
    DownloadTicks = False
    
    strSymbol = Bars.Prop(eBARS_Symbol)
    strSecType = Bars.SecurityType
    Set astrRequest = New cGdArray
    astrRequest.Add "@" & Format(dSessionDate, "YYYYMMDD") & "-" & _
            Format(dSessionDate, "YYYYMMDD") & ";T;I;" & strSecType & ";" & strSymbol
                        
    If FtpRequest(astrRequest) Then
    
        ' Distribute the data (if data to distribute)
        If Not DistributeData("Distributing Data", False) Then
            'State.eCurStatus = 600
            frmStatus.Status = eStatus_Error
            frmStatus.AddDetail "ERROR downloading data"
        End If
    
        ' Distribute the data (if data to distribute)
        If frmStatus.Status = eStatus_Completed Then
            frmStatus.AddDetail "Final Updating"
            DM_DistribData ""
            frmStatus.Status = eStatus_Running
            
            frmStatus.Status = eStatus_Completed
            frmStatus.AddDetail "Finished"
            DownloadTicks = True
        End If
    End If

ErrExit:
    Set astrRequest = Nothing
    Exit Function
    
ErrSection:
    Set astrRequest = Nothing
    RaiseError "mDataNav.DownloadTicks", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFunctionFavorites
'' Description: Get a list of the user's favorite functions
'' Inputs:      None
'' Returns:     Array of Favorites
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetFunctionFavorites() As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As New cGdArray
    
    astrReturn.Create eGDARRAY_Strings
    
    If FileLength(AddSlash(App.Path) & "Custom\Function.FAV") > 0 Then
        astrReturn.FromFile AddSlash(App.Path) & "Custom\Function.FAV"
    ElseIf FileExist(AddSlash(App.Path) & "Provided\Function.FAV") Then
        astrReturn.FromFile AddSlash(App.Path) & "Provided\Function.FAV"
    ElseIf FileExist(AddSlash(App.Path) & "Info\Common.IND") Then
        FileCopy AddSlash(App.Path) & "Info\Common.IND", AddSlash(App.Path) & "Provided\Function.FAV"
        astrReturn.FromFile AddSlash(App.Path) & "Provided\Function.FAV"
    End If
    
    astrReturn.Sort
    
    Set GetFunctionFavorites = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.GetFunctionFavorites", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PriceDisplay
'' Description: Display a given price in the trading units of the given symbol
'' Inputs:      Price to Convert, Symbol, Display in Trading Units?, Return
''              Empty if Zero?
'' Returns:     Formatted Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PriceDisplay(ByVal dPrice As Double, ByVal vSymbol As Variant, Optional ByVal bTradingUnits As Boolean = True, Optional ByVal bBlankIfZero As Boolean = True) As String
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID to convert price for
    Dim dTickMove As Double             ' Tick Move for the symbol
    Dim dTickValue As Double            ' Tick Value for the symbol
    Dim dMinMoveInTicks As Double       ' Minimum Move in Ticks for the symbol
    Dim strPriceFormat As String        ' Price format for the symbol
    Dim Bars As New cGdBars             ' Bars to convert the price
    
    lSymbolID = GetSymbolID(vSymbol)
    
    If dPrice = 0# And bBlankIfZero Then
        PriceDisplay = ""
    ElseIf SU_GetTickInfoWithFormat(lSymbolID, dTickValue, dTickMove, dMinMoveInTicks, strPriceFormat) = True Then
        Bars.Prop(eBARS_TickMove) = dTickMove
        Bars.Prop(eBARS_MinMoveInTicks) = dMinMoveInTicks
        Bars.Prop(eBARS_PriceDisplayFormat) = strPriceFormat
        
        PriceDisplay = Bars.PriceDisplay(dPrice, bTradingUnits)
    Else
        PriceDisplay = Str(dPrice)
    End If

ErrExit:
    Set Bars = Nothing
    Exit Function
    
ErrSection:
    Set Bars = Nothing
    RaiseError "mDataNav.PriceDisplay", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PriceFromDisplay
'' Description: Get the value of a price from the displayed version
'' Inputs:      Formatted Price, Symbol
'' Returns:     Value of the string passed in
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PriceFromDisplay(ByVal strPrice As String, ByVal vSymbol As Variant) As Double
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID to convert price for
    Dim dTickMove As Double             ' Tick Move for the symbol
    Dim dTickValue As Double            ' Tick Value for the symbol
    Dim dMinMoveInTicks As Double       ' Minimum Move in Ticks for the symbol
    Dim strPriceFormat As String        ' Price format for the symbol
    Dim Bars As New cGdBars             ' Bars to convert the price
    
    lSymbolID = GetSymbolID(vSymbol)
    
    If SU_GetTickInfoWithFormat(lSymbolID, dTickValue, dTickMove, dMinMoveInTicks, strPriceFormat) = True Then
        Bars.Prop(eBARS_TickMove) = dTickMove
        Bars.Prop(eBARS_MinMoveInTicks) = dMinMoveInTicks
        Bars.Prop(eBARS_PriceDisplayFormat) = strPriceFormat
        
        PriceFromDisplay = Bars.PriceFromString(strPrice)
    Else
        PriceFromDisplay = Val(strPrice)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.PriceFromDisplay", eGDRaiseError_Raise
    
End Function

Public Sub UnminutizeTicks(Ticks As cGdBars)
On Error GoTo ErrSection:

    Dim i&, iTick&, iTick2&, iUp&, iDown&, iUpDown&
    Dim dt#, dOpen#, dHigh#, dLow#, dClose#, dPrevClose#, dVol#
    Dim bUseVol As Boolean
    Dim Ticks2 As cGdBars
    Dim bFullTicksAvail As Boolean

    If IsMinutized(Ticks, bFullTicksAvail) Then
        Set Ticks2 = Ticks.MakeCopy
        Ticks.ArrayMask = eBARS_TickByTick
        If Ticks.SecurityType = "S" And Ticks(eBARS_Vol, 0) >= 0 Then
            bUseVol = True
        End If
        iTick = 0
        For iTick2 = 0 To Ticks2.Size '(need to go past end)
            If Ticks2(eBARS_DateTime, iTick2) <> dt Then
                ' store ticks for previous minute
                If iUp + iDown > 0 Then
                    If dHigh = dLow Then
                        ' just do "n" ticks
                        For i = 1 To iUp + iDown
                            Ticks(eBARS_DateTime, iTick) = dt
                            Ticks(eBARS_Close, iTick) = dClose
                            If bUseVol Then
                                Ticks(eBARS_Vol, iTick) = dVol
                                dVol = 0
                            End If
                            iTick = iTick + 1
                        Next
                    Else
                        ' open tick
                        Ticks(eBARS_DateTime, iTick) = dt
                        Ticks(eBARS_Close, iTick) = dOpen
                        If bUseVol Then
                            Ticks(eBARS_Vol, iTick) = dVol
                            dVol = 0
                        End If
                        iTick = iTick + 1
                        If dOpen < dPrevClose Then
                            iDown = iDown - 1
                            iUpDown = -1
                        ElseIf dOpen > dPrevClose Then
                            iUp = iUp - 1
                            iUpDown = 1
                        ElseIf iUpDown < 0 Then
                            iDown = iDown - 1
                        Else
                            iUp = iUp - 1
                        End If
                        If dLow < dOpen Or iUpDown < 0 Then
                            ' adjust for close tick
                            If dClose < dHigh Then
                                iDown = iDown - 1
                            Else
                                iUp = iUp - 1
                            End If
                            If iUp < 0 Then
                                iDown = iDown + Abs(iUp)
                                iUp = 0
                            ElseIf iDown < 0 Then
                                iUp = iUp + Abs(iDown)
                                iDown = 0
                            End If
                            ' do all the down ticks
                            For i = 1 To iDown
                                Ticks(eBARS_DateTime, iTick) = dt
                                Ticks(eBARS_Close, iTick) = dLow
                                If bUseVol Then
                                    Ticks(eBARS_Vol, iTick) = 0
                                End If
                                iTick = iTick + 1
                            Next
                            ' do all the up ticks
                            For i = 1 To iUp
                                Ticks(eBARS_DateTime, iTick) = dt
                                Ticks(eBARS_Close, iTick) = dHigh
                                If bUseVol Then
                                    Ticks(eBARS_Vol, iTick) = 0
                                End If
                                iTick = iTick + 1
                            Next
                        Else
                            ' adjust for close tick
                            If dClose > dLow Then
                                iUp = iUp - 1
                            Else
                                iDown = iDown - 1
                            End If
                            If iUp < 0 Then
                                iDown = iDown + Abs(iUp)
                                iUp = 0
                            ElseIf iDown < 0 Then
                                iUp = iUp + Abs(iDown)
                                iDown = 0
                            End If
                            ' do all the up ticks
                            For i = 1 To iUp
                                Ticks(eBARS_DateTime, iTick) = dt
                                Ticks(eBARS_Close, iTick) = dHigh
                                If bUseVol Then
                                    Ticks(eBARS_Vol, iTick) = 0
                                End If
                                iTick = iTick + 1
                            Next
                            ' do all the down ticks
                            For i = 1 To iDown
                                Ticks(eBARS_DateTime, iTick) = dt
                                Ticks(eBARS_Close, iTick) = dLow
                                If bUseVol Then
                                    Ticks(eBARS_Vol, iTick) = 0
                                End If
                                iTick = iTick + 1
                            Next
                        End If
                        ' close tick
                        Ticks(eBARS_DateTime, iTick) = dt
                        Ticks(eBARS_Close, iTick) = dClose
                        If bUseVol Then
                            Ticks(eBARS_Vol, iTick) = 0
                        End If
                        iTick = iTick + 1
                    End If
                    If iTick2 = Ticks2.Size Then Exit For '(done)
                End If
                ' init for new minute
                dt = Ticks2(eBARS_DateTime, iTick2)
                dPrevClose = dClose
                dClose = Ticks2(eBARS_Close, iTick2)
                dHigh = dClose
                dLow = dClose
                dOpen = dClose
                iUp = Ticks2(eBARS_UpTicks, iTick2)
                If iUp < 0 Then iUp = 0
                iDown = Ticks2(eBARS_DownTicks, iTick2)
                If iDown < 0 Then iDown = 0
                If bUseVol Then
                    dVol = Ticks2(eBARS_Vol, iTick2)
                    If dVol < 0 Then dVol = 0
                End If
            Else
                ' tick in same minute
                dClose = Ticks2(eBARS_Close, iTick2)
                If dClose > dHigh Then dHigh = dClose
                If dClose < dLow Then dLow = dClose
                i = Ticks2(eBARS_UpTicks, iTick2)
                If i > 0 Then iUp = iUp + i
                i = Ticks2(eBARS_DownTicks, iTick2)
                If i > 0 Then iDown = iDown + i
                If bUseVol Then
                    i = Ticks2(eBARS_Vol, iTick2)
                    If i > 0 Then dVol = dVol + i
                End If
            End If
        Next
        Ticks.Size = iTick
        Set Ticks2 = Nothing
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mDataNav.UnminutizeTicks", eGDRaiseError_Raise
End Sub

' Returns single-character security type ("S", "F", "I")
' - can pass either a cGdBars, Symbol, or SymbolID
' - but SHOULD pass the Bars object if have it (to support external symbols from MS or CSI files)
Public Function SecurityType(BarsOrSymbol As Variant, Optional bIfOptionAppendO As Boolean = False) As String
On Error GoTo ErrSection:
    
    Dim strSymbol$, bOption As Boolean
    
    If VarType(BarsOrSymbol) = vbObject Then
        SecurityType = BarsOrSymbol.SecurityType
    Else
        strSymbol = GetSymbol(BarsOrSymbol)
        If Len(strSymbol) > 0 Then
            If InStr(strSymbol, " ") > 0 Then
                bOption = True
                strSymbol = Parse(strSymbol, " ", 1)
            Else
                bOption = False
            End If
            If Left(strSymbol, 1) = "$" Or Left(strSymbol, 1) = "#" Then
                SecurityType = "I"
            ElseIf InStr(strSymbol, "-") > 0 Then
                SecurityType = "F"
            Else
                SecurityType = "S"
                ' Mutual Fund: if Len >= 5, last char = "X", and all Alpha (i.e. no underscores, @, etc)
                If Len(strSymbol) >= 5 And Right(strSymbol, 1) = "X" Then
                    If IsAlpha(strSymbol, -1) Then
                        SecurityType = "M"
                    End If
                End If
            End If
            If bOption And bIfOptionAppendO Then
                SecurityType = SecurityType & "O"
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.SecurityType", eGDRaiseError_Raise
End Function

' To check if need to do daily updates
Public Function NeedDailyUpdate() As Boolean
On Error GoTo ErrSection:
        
    Dim dDate#, dNyTime#
    
    ' see if past 8pm NY time
    dNyTime = ConvertTimeZone(Now)
    If dNyTime - Int(dNyTime) > 20 / 24# Then
        dDate = Int(dNyTime) ' today's file should be ready
    Else
        dDate = Int(dNyTime) - 1 ' yesterday's file should be ready
    End If
    
    ' back up if a weekend
    If Month(dDate) = 1 And Day(dDate) = 1 Then
        dDate = dDate - 1 ' no file on New Year's day
    End If
    If Weekday(dDate) = vbSunday Then
        dDate = dDate - 2
    ElseIf Weekday(dDate) = vbSaturday Then
        dDate = dDate - 1
    End If
    
    ' see if last daily download was prior
    If LastDailyDownload < dDate Then
        NeedDailyUpdate = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.NeedDailyUpdate", eGDRaiseError_Raise
End Function

' Copies/unzips distribution files into the "\ftp\dist\" folder, creates the .CTL file,
' and distributes the data (returns true if data was distributed)
Public Function DistributeData(Optional ByVal strStatusMsg As String = "", _
    Optional ByVal bSnapshot As Boolean = True, _
    Optional ByVal strFlag As String = ",0") As Boolean
On Error GoTo ErrSection:

    Dim strDate$, strFile$, strType$, strFtpPath$, strDistPath$, i&
    Dim aTodayFiles As New cGdArray

    ' use date of last daily download so won't wipe out the current snapshot area
    'strDate = "," & Right(Str(JulToLong(LastDailyDownload, True)), 6) & ",0"
    'mod per Karl for single symbol history download:
    '- date format: yyyymmdd
    '- flag needs to be 4 instead of 0
    strDate = "," & Str(JulToLong(LastDailyDownload, True)) & strFlag

    If bSnapshot Then
        strType = "Latest "
    Else
        strType = "Distribute "
    End If

    ' get paths
    strFtpPath = AddSlash(App.Path) & "FTP\"
    strDistPath = strFtpPath & "Dist\"

    ' Clean out the distribution directory
    KillFile strDistPath & "*.*", True

    ' Unzip the current session update files if they exist
    aTodayFiles.GetMatchingFiles strFtpPath & "Today_*.gzp"
    For i = 0 To aTodayFiles.Size - 1
        ZipExecute "U", aTodayFiles(i), strDistPath
    Next

    ' Copy files into the distribution directory and create control file
    strFile = "Data.DTX"
    If FileExist(strFtpPath & strFile) Then
        FileCopy strFtpPath & strFile, strDistPath & strFile
        FileFromString strDistPath & "zzz.ctl", _
            strType & "EOD," & strFile & strDate, True, True
    End If
    strFile = "Data.DAT"
    If FileExist(strFtpPath & strFile) Then
        FileCopy strFtpPath & strFile, strDistPath & strFile
        FileFromString strDistPath & "zzz.ctl", _
            strType & "EOD," & strFile & strDate, True, True
    End If
    strFile = "Data.GTX"
    If FileExist(strFtpPath & strFile) Then
        FileCopy strFtpPath & strFile, strDistPath & strFile
        FileFromString strDistPath & "zzz.ctl", _
            strType & "TICK," & strFile & strDate, True, True
    End If
    strFile = "Data.ASC"
    If FileExist(strFtpPath & strFile) Then
        FileCopy strFtpPath & strFile, strDistPath & strFile
        FileFromString strDistPath & "zzz.ctl", _
            strType & "TICK," & strFile & strDate, True, True
    End If
    strFile = "Data.BTF"
    If FileExist(strFtpPath & strFile) Then
        FileCopy strFtpPath & strFile, strDistPath & strFile
        FileFromString strDistPath & "zzz.ctl", _
            strType & "Bad Tick," & strFile & strDate, True, True
    End If
    
    ' Distribute the data
    If FileExist(strDistPath & "*.*") Then
        If Len(strStatusMsg) > 0 Then
            frmStatus.AddDetail strStatusMsg
        End If
        DistributeData = DM_DistribData(strDistPath)
    End If
    
    ' Backup the data files
    If FileExist(strFtpPath & "Today*.GZP") Then
        FileCopy strFtpPath & "Today*.GZP", strFtpPath & "Backup\", True
    End If
    If FileExist(strFtpPath & "Data.*") Then
        FileCopy strFtpPath & "Data.*", strFtpPath & "Backup\", True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.DistributeData", eGDRaiseError_Raise
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewUser
'' Description: Is this a new user (no stored customer information)?
'' Inputs:      None
'' Returns:     True if no stored customer information, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewUser() As Boolean
On Error GoTo ErrSection:

    If RI_GetDataServiceID = 0 Then NewUser = True Else NewUser = False

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.NewUser", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DownloadDataSet
'' Description: Download the appropriate data set from Genesis servers
'' Inputs:      None
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DownloadDataSet() As Boolean
On Error GoTo ErrSection:

    Dim astrRequest As New cGdArray     ' Array of requests
    Dim bSuccess As Boolean             ' Success from the download
    Dim lValid As Long                  ' Return from Unzip command
    Dim strButtons As String            ' Buttons to display to the user

    DownloadDataSet = False
    
    If g.SymbolPool.NumRecords = 0 Then
        strButtons = "+OK|-Exit"
    Else
        strButtons = "+OK|-Cancel"
    End If
    
    Do While InfBox("We will need to connect to Genesis and download your starting data set.|(This will require an internet connection)", "i", strButtons, "Starting Data Set") = "O"
        If Dir(App.Path & "\ftp\*.*") <> "" Then KillFile AddSlash(App.Path) & "Ftp\*.*", True
        frmStatus.Status = eStatus_Initialized
        
        frmStatus.SetTitle "Downloading Data Set"
        astrRequest.Size = 0
        astrRequest.Add "%Download Data Set"
        If frmStatus.Status < eStatus_Aborting Then
            ' after 10 seconds, show the quick start info
            ' (this allows getting the downloading started)
            frmMain.tmrQuickStart.Interval = 10000
            frmMain.tmrQuickStart.Enabled = True
            bSuccess = FtpRequest(astrRequest)
            frmMain.tmrQuickStart.Enabled = False
        End If
        
        If frmStatus.Status = eStatus_Aborted Then
            '(nothing more to do)
        ElseIf frmStatus.Status = eStatus_Error Or bSuccess = False Then
            frmStatus.AddDetail "ERROR downloading data set"
            frmStatus.Status = eStatus_Error
        ElseIf frmStatus.Status = eStatus_Completed Then
            If FileExist(AddSlash(App.Path) & "FTP\DataSet.GZP") Then
                FileCopy AddSlash(App.Path) & "FTP\DataSet.GZP", AddSlash(App.Path) & "FTP\Backup\", True
                InfBox "Please wait while Trade Navigator|sets up the data", , , , True
                CleanOutDataFolder
                lValid = ZipExecute("U", AddSlash(App.Path) & "FTP\DataSet.GZP", AddSlash(App.Path) & "Data", "", False, False)
                DownloadDataSet = True
            ElseIf FileExist(AddSlash(App.Path) & "FTP\Backup\DataSet.GZP") Then
                InfBox "Please wait while Trade Navigator|sets up the data", , , , True
                CleanOutDataFolder
                lValid = ZipExecute("U", AddSlash(App.Path) & "FTP\Backup\DataSet.GZP", AddSlash(App.Path) & "Data", "", False, False)
                DownloadDataSet = True
            Else
                frmStatus.AddDetail "Data set not found"
                frmStatus.Status = eStatus_Error
            End If
        End If
        
        If frmStatus.Status <> eStatus_Aborted And frmStatus.Status <> eStatus_Error Then
            frmStatus.Status = eStatus_Completed
            frmStatus.AddDetail "Finished"
        End If
        
        If frmStatus.Status = eStatus_Completed Then Exit Do
    Loop

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.DownloadDataSet", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectionInfo
'' Description: Build a delimited string of connection information
'' Inputs:      None
'' Returns:     String of Conection Info
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConnectionInfo() As String
On Error GoTo ErrSection:

    Dim strIniFile As String            ' Path and name of the ini file
    Dim bUseFTP As Boolean              ' FTP or HTTP?
    Dim lUseProxy As Long               ' Using Proxy?
    Dim lLogin As Long                  ' Using Login to Proxy?
    Dim astrInfo As New cGdArray        ' Array of information
    
    astrInfo.Create eGDARRAY_Strings
    strIniFile = AddSlash(App.Path) & "GClient.INI"
    
    bUseFTP = GetIniFileProperty("UseFTP", False, "Mode", strIniFile)
    If bUseFTP Then
        astrInfo.Add "FTP"
    Else
        astrInfo.Add "HTTP"
    End If

    lUseProxy = GetIniFileProperty("UseProxy", vbUnchecked, "Proxy", strIniFile)
    If lUseProxy = vbUnchecked Then
        astrInfo.Add "0"
    Else
        astrInfo.Add "1"
    End If
    
    astrInfo.Add GetIniFileProperty("ProxyServer", "", "Proxy", strIniFile)
    astrInfo.Add GetIniFileProperty("ProxyPort", "", "Proxy", strIniFile)
    
    lLogin = GetIniFileProperty("SendLogin", vbUnchecked, "Proxy", strIniFile)
    If lLogin = vbUnchecked Then
        astrInfo.Add "0"
    Else
        astrInfo.Add "1"
    End If
    
    astrInfo.Add GetIniFileProperty("LoginUser", "", "Proxy", strIniFile)
    astrInfo.Add DecryptFromHex(GetIniFileProperty("LoginPassword", "", "Proxy", strIniFile), frmHTTP.PasswordKey)
    
    ConnectionInfo = astrInfo.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.ConnectionInfo", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CleanOutDataFolder
'' Description: Clean out the data folder
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CleanOutDataFolder()
On Error GoTo ErrSection:

    Dim strDataPath As String           ' Path of the data

    strDataPath = AddSlash(DataPath)
    
    ClearReadOnlyFlags strDataPath & "*.*"
    
    ' Wipe out the temporary "Old" folder
    KillFile strDataPath & "Old\*.*"
    
    ' Clean out the Data directory
    KillFile AddSlash(App.Path) & "SYMPOOL.MEM"
    KillFile strDataPath & "*.*"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mDataNav.CleanOutDataFolder", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumBusinessDays
'' Description: Number of business days from one date to the other
'' Inputs:      From Date, To Date
'' Returns:     Number of Week days
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NumBusinessDays(ByVal lFromDate As Long, ByVal lToDate As Long) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value from the function
    
    lReturn = 0&
    For lIndex = lFromDate To lToDate
        If IsWeekday(lIndex) Then lReturn = lReturn + 1
    Next lIndex
    
    NumBusinessDays = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.NumBusinessDays", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CleanOutCustomIndexes
'' Description: Clean out any custom indexes for which the symbol group no
''              longer exists
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CleanOutCustomIndexes()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID of the custom index
    Dim strSymbol As String             ' Symbol of the custom index
    Dim strGroup As String              ' Name of the symbol group
    Dim lSymGroup As Long               ' Index into a for loop
    Dim bFound As Boolean               ' Did we find the symbol group?
    
    ' Walk through the symbols in the pool looking for custom indexes...
    For lIndex = g.SymbolPool.NumRecords - 1 To 0 Step -1
        lSymbolID = g.SymbolPool.SymbolID(lIndex)
        
        ' Only look at custom indexes...
        If lSymbolID < 0 Then
            strSymbol = g.SymbolPool.Symbol(lIndex)
            If Len(strSymbol) > 0 Then
                ' Determine the name of the underlying symbol group...
                strGroup = strSymbol
                If Left(strGroup, 1) = "#" Then strGroup = Mid(strGroup, 2)
                
                ' See if the corresponding symbol group exists...
                bFound = False
                For lSymGroup = 1 To g.SymbolPool.SymbolGroups.Count
                    If UCase(g.SymbolPool.SymbolGroups(lSymGroup).Name) = UCase(strGroup) Then
                        If g.SymbolPool.SymbolGroups(lSymGroup).IsIndex = True Then
                            bFound = True
                            Exit For
                        End If
                    End If
                Next lSymGroup
                
                ' If the symbol group does not exist, we need to delete the custom index...
                If bFound = False Then
                    SU_DeleteComposite lSymbolID, strSymbol
                    g.SymbolPool.RemoveCustomIndex lSymbolID
                End If
                
                DebugLog strSymbol & "(" & Str(lSymbolID) & "): " & strGroup & " = " & CStr(bFound)
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mDataNav.CleanOutCustomIndexes"
    
End Sub

Public Function GclientCallbackProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim strText$
    Dim cds As COPYDATASTRUCT
    Dim mb As cMemBuffer

    If iMsg = WM_COPYDATA Then
        ' first copy "CopyData" structure (so can read it)
        CopyMemory cds, ByVal lParam, Len(cds)
        ' then copy string
        Set mb = New cMemBuffer
        mb.PutFromMemory cds.lpData, cds.cbData
        strText = mb.Buffer
        Set mb = Nothing
        ' then send it to frmStatus
        frmStatus.ProcessStatusMsg strText
        
        GclientCallbackProc = 1
    Else
        GclientCallbackProc = DefWindowProc(hWnd, iMsg, wParam, ByVal lParam)
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsExpiredContract
'' Description: Determine if the futures contract is expired
'' Inputs:      Symbol (or Symbol ID), Expiration Date, Use 56?, Compare against today?
'' Returns:     True if expired, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsExpiredContract(ByVal vSymbolOrSymbolID As Variant, Optional ByVal lExpirationDate As Long = kNullData, Optional ByVal bUse56 As Boolean = False, Optional ByVal bUseToday As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strSymbol As String             ' String representation of the symbol
    Dim strSecType As String            ' Security type for the symbol
    Dim lContract As Long               ' Contract
    Dim lDate As Long                   ' First of contract expiration month
    Dim Bars As New cGdBars             ' Bars structure
    Dim lLDL As Long                    ' Last daily download plus one business
    Dim dExpirationDate As Double       ' Expiration date back from LoadSymbolInfo call
    Dim dLotSize As Double              ' Lot size back from LoadSymbolInfo call
    Dim strFront56 As String            ' Front contract for the 56
    
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    strSecType = SecurityType(strSymbol, True)
    
    If bUseToday Then
        lLDL = CurrentTime
    Else
        lLDL = LastDailyDownload + 1
        While IsWeekday(lLDL) = False
            lLDL = lLDL + 1
        Wend
    End If
    
    lDate = 999999
    If InStr(strSymbol, "-0") = 0 Then
        If lExpirationDate > 0 Then
            lDate = lExpirationDate
        ElseIf g.Broker.LoadSymbolInfo(strSymbol, dExpirationDate, dLotSize) Then
            lDate = CLng(dExpirationDate)
        ElseIf (strSecType = "F") Or (strSecType = "FO") Then
            If (strSecType = "F") And (bUse56 = True) Then
                strFront56 = ConvertToTradeSymbol(Parse(strSymbol, "-", 1) & "-056", CurrentTime)
                If CLng(Val(Parse(strSymbol, "-", 2))) < CLng(Val(Parse(strFront56, "-", 2))) Then
                    lDate = -1&
                End If
            Else
                SetBarProperties Bars, strSymbol
                lDate = JulFromLong(Bars.LastDayOfContractMonth)
            End If
        ElseIf (strSecType = "SO") And (Len(strSymbol) > 6) Then
            lDate = JulFromLong(Parse(strSymbol, " ", 2))
        End If
    End If
    
    IsExpiredContract = (lDate < lLDL)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.IsExpiredContract"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentTime
'' Description: If streaming get the feed time, otherwise return now
'' Inputs:      To Time Zone, Symbol, Allow Replay Time?
'' Returns:     Feed Time or Now
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentTime(Optional ByVal strToTimeZone As String = "", Optional ByVal strSymbol As String = "", Optional ByVal bAllowReplayTime As Boolean = False) As Double
On Error GoTo ErrSection:

    ' (TLB 2/14/2013: reorganized code to simplify it and only call the g.RealTime.FeedTime once)
    Dim dTime As Double

    ' if the realtime object is active...
    If Not g.RealTime Is Nothing Then
        If g.RealTime.Active Then
            ' and if not in a replay session or want to use the replay time, then use the feed time...
            If (g.nReplaySession = 0) Or bAllowReplayTime Then
                ' but only if the feed time is valid...
                dTime = g.RealTime.FeedTime(strSymbol)
                If dTime > 0 Then
                    dTime = ConvertTimeZone(dTime, "NY", strToTimeZone)
                End If
            End If
        End If
    End If
    
    ' otherwise just use the system time...
    If dTime <= 0 Then
        dTime = ConvertTimeZone(Now, "", strToTimeZone)
    End If
    
    CurrentTime = dTime

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.CurrentTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumTicksFromMarket
'' Description: Return the number of ticks the given price is from the market
'' Inputs:      Price, Symbol, Use Min Move?, Trigger Price
'' Returns:     Number of ticks from the market (Null if data not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NumTicksFromMarket(ByVal dPrice As Double, ByVal vSymbolOrSymbolID As Variant, Optional ByVal bUseMinMove As Boolean = True, Optional ByVal dTriggerPrice As Double = kNullData) As Long
On Error GoTo ErrSection:

    Dim dLastKnownPrice As Double       ' Last known price for symbol
    Dim dDiff As Double                 ' Difference in price
    Dim lReturn As Long                 ' Return value for the function
    Dim Bars As cGdBars                 ' Bars structure
    Dim dTickMove As Double             ' Tick move
    
    lReturn = kNullData
    
    Set Bars = New cGdBars
    SetBarProperties Bars, vSymbolOrSymbolID
    
    If dTriggerPrice <> kNullData Then
        dLastKnownPrice = dTriggerPrice
    Else
        dLastKnownPrice = g.RealTime.LastKnownPrice(vSymbolOrSymbolID)
    End If
    
    If (dLastKnownPrice <> kNullData) And (Bars.TickMove <> 0) Then
        dDiff = dPrice - dLastKnownPrice
        If bUseMinMove Then
            dTickMove = Bars.MinMove
        Else
            dTickMove = Bars.TickMove
        End If
        
        If dTickMove <> 0 Then
            lReturn = dDiff / dTickMove
        End If
    End If
    
    NumTicksFromMarket = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.NumTicksFromMarket"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PreviousCloseForSymbol
'' Description: Return the previous close for the given symbol
'' Inputs:      Symbol
'' Returns:     Previous Close
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PreviousCloseForSymbol(ByVal vSymbolOrSymbolID As Variant, Optional dPreviousCloseTime As Double) As Double
On Error GoTo ErrSection:
   
    Dim Bars As cGdBars                 ' Bars object
    Dim lCurrentSession As Long         ' Current session date
    Dim lIndex As Long                  ' Index variable
    Dim dReturn As Double               ' Return value
    
    dReturn = kNullData
    
    Set Bars = New cGdBars
    If DM_GetBars(Bars, vSymbolOrSymbolID, , LastDailyDownload - 5) Then
        g.RealTime.SpliceBars Bars
        lCurrentSession = Bars.SessionDateForTradeTime(CurrentTime(Bars.Prop(eBARS_ExchangeTimeZoneInf)))
        
        lIndex = Bars.Size - 1
        Do While (Bars(eBARS_DateTime, lIndex) >= lCurrentSession) And (lIndex >= 0)
            lIndex = lIndex - 1
        Loop
        
        If lIndex >= 0 Then
            dReturn = Bars(eBARS_Close, lIndex)
            dPreviousCloseTime = Int(Bars(eBARS_DateTime, lIndex)) + (Bars.Prop(eBARS_EndTime) / 1440#)
        End If
    End If
    
    PreviousCloseForSymbol = dReturn
   
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mDataNav.PreviousCloseForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoImportQuoteBoardTabs
'' Description: Auto import any quote board tab files that are in the app path
'' Inputs:      Startup?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoImportQuoteBoardTabs(Optional ByVal bStartup As Boolean = True)
On Error GoTo ErrSection:

    Dim astrFiles As cGdArray           ' Array of filenames
    Dim lIndex As Long                  ' Index into a for loop
    Dim strQbtFile As String            ' File converted to a string
    Dim strQbtFileName As String        ' QBT file name with path
    Dim strDonFileName As String        ' Done file to compare with
    Dim strQbtDate As String            ' Date/Time stamp for the QBT file
    Dim strDonDate As String            ' Date/Time stamp for the DON file
    
    Set astrFiles = New cGdArray
    astrFiles.Create eGDARRAY_Strings
    
    astrFiles.GetMatchingFiles AddSlash(App.Path) & "QBT\*.QBT", False
    If (astrFiles.Size > 0) And (FileExist(AddSlash(App.Path) & "Custom\QuoteBoard.INF") <> 0) Then
        If bStartup Then
            frmSplash.Message 25, "Importing Quote Board Tabs"
        End If
        
        For lIndex = 0 To astrFiles.Size - 1
            strQbtFileName = AddSlash(App.Path) & "QBT\" & astrFiles(lIndex)
            strDonFileName = AddSlash(App.Path) & "QBT\" & FileBase(astrFiles(lIndex)) & ".DON"
            
            strQbtDate = FileToString(strQbtFileName, , True)
            strDonDate = FileToString(strDonFileName, , True)
            
            If Val(strQbtDate) > Val(strDonDate) Then
                strQbtFile = FileToString(strQbtFileName)
                ImportQuoteBoardTab strQbtFile
                FileCopy strQbtFileName, strDonFileName
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mDataNav.AutoImportQuoteBoardTabs"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportQuoteBoardTab
'' Description: Create a quote board tab from the given string
'' Inputs:      Quote Board Tab Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportQuoteBoardTab(ByVal strTabString As String)
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' Information broken out into an array
    Dim astrQuoteBoard As cGdArray      ' Quote board file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lVersion As Long                ' File format version
    Dim strTabInfo As String            ' Tab information
    Dim astrCriteria As cGdArray        ' Array of criteria information
    Dim astrAlerts As cGdArray          ' Array of alert information
    Dim bFound As Boolean               ' Is the tab already in the table information?
    Dim bContinue As Boolean            ' Do we want to continue?
    Dim lPos As Long                    ' Position in a string
    Dim strFileName As String           ' Filename
    Dim strCriteria As String           ' Criteria
    Dim bReload As Boolean              ' Do we need to reload?
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strName As String               ' Name of the criteria
    Dim Alert As cAlert                 ' Alert to add to the collection
    Dim lExistingTab As Long            ' Existing quote board tab
    Dim astrAlertsFile As cGdArray      ' Alerts file
    Static bSkipTabCheck As Boolean     ' Skip the check if the tab exists?
    Static bSkipCriteriaCheck As Boolean ' Skip the check if the criteria exists?
    
    ' Initialize the arrays...
    Set astrFile = New cGdArray
    astrFile.Create eGDARRAY_Strings
    Set astrQuoteBoard = New cGdArray
    astrQuoteBoard.Create eGDARRAY_Strings
    Set astrCriteria = New cGdArray
    astrCriteria.Create eGDARRAY_Strings
    Set astrAlerts = New cGdArray
    astrAlerts.Create eGDARRAY_Strings
    Set astrAlertsFile = New cGdArray
    astrAlertsFile.Create eGDARRAY_Strings
    
    ' Read in the information from the string (ignore the date/time stamp that is the first line)...
    astrFile.SplitFields Replace(strTabString, vbCrLf, vbLf), vbLf
    For lIndex = 1 To astrFile.Size - 1
        Select Case UCase(astrFile(lIndex))
            Case "[VERSION]"
                lIndex = lIndex + 1
                lVersion = CLng(Val(astrFile(lIndex)))
                
            Case "[TAB INFO]"
                lIndex = lIndex + 1
                strTabInfo = astrFile(lIndex)
                
            Case "[CRITERIA]"
                Do While (Left(astrFile(lIndex + 1), 1) <> "[") And (lIndex + 1 < astrFile.Size)
                    lIndex = lIndex + 1
                    astrCriteria.Add astrFile(lIndex)
                Loop
                
            Case "[ALERTS]"
                Do While (Left(astrFile(lIndex + 1), 1) <> "[") And (lIndex + 1 < astrFile.Size)
                    lIndex = lIndex + 1
                    astrAlerts.Add astrFile(lIndex)
                Loop
                
        End Select
    Next lIndex
    
    ' Verify the version of the file...
    If lVersion < 1 Then
        InfBox "The version of the file that you are trying to import is too old", "!", , "Quote Board Tab Import Error"
        bContinue = False
    ElseIf lVersion > 1 Then
        InfBox "The version of the file that you are trying to import is too new", "!", , "Quote Board Tab Import Error"
        bContinue = False
    Else
        bContinue = True
    End If
    
    ' See if the quote board tab already exists (there is a tab by the same name)...
    If bContinue Then
        bFound = False
        lExistingTab = -1&
        
        If FileExist(AddSlash(App.Path) & "Custom\QuoteBoard.INF") Then
            astrQuoteBoard.SplitFields FileToString(AddSlash(App.Path) & "Custom\QuoteBoard.INF"), vbLf
        
            For lIndex = 0 To astrQuoteBoard.Size - 1
                If Parse(astrQuoteBoard(lIndex), vbTab, 1) = Parse(strTabInfo, vbTab, 1) Then
                    lExistingTab = lIndex
                    bFound = True
                    Exit For
                End If
            Next lIndex
            
            bContinue = True
            If bFound Then
                ' 03/05/2010 DAJ: If they said yes to this question once, apply it to all...
                If bSkipTabCheck Then
                    bContinue = True
                Else
                    If InfBox("There is already a quote board tab named|'" & Parse(strTabInfo, vbTab, 1) & "'.||Would you like to overwrite it?", "?", "+Yes|-No", "Quote Board Tab Import") = "Y" Then
                        bContinue = True
                        bSkipTabCheck = True
                    Else
                        bContinue = False
                    End If
                End If
            End If
        End If
    End If
    
    ' If the quote board tab doesn't exist, make sure none of the criteria exist...
    If (bFound = False) And (bContinue = True) Then
        For lIndex = 0 To astrCriteria.Size - 1
            lPos = InStr(astrCriteria(lIndex), vbTab)
            strFileName = Parse(Left(astrCriteria(lIndex), lPos - 1), "=", 2)
            strCriteria = Mid(astrCriteria(lIndex), lPos + 1)
            
            If FileExist(AddSlash(App.Path) & strFileName) = True Then
                bFound = True
                Exit For
            End If
        Next lIndex
    
        If bFound Then
            ' 03/05/2010 DAJ: If they said yes to this question once, apply it to all...
            If bSkipCriteriaCheck Then
                bContinue = True
            Else
                If InfBox("One or more of the criteria used as a field for this quote board tab already exist.||Would you like to overwrite them?", "?", "+Yes|-No", "Quote Board Tab Import") = "Y" Then
                    bContinue = True
                    bSkipCriteriaCheck = True
                Else
                    bContinue = False
                End If
            End If
        End If
    End If
    
    ' Import the criteria...
    If bContinue Then
        bReload = False
        
        For lIndex = 0 To astrCriteria.Size - 1
            lPos = InStr(astrCriteria(lIndex), vbTab)
            strFileName = Parse(Left(astrCriteria(lIndex), lPos - 1), "=", 2)
            strCriteria = Mid(astrCriteria(lIndex), lPos + 1)
            
            FileFromString AddSlash(App.Path) & strFileName, Replace(strCriteria, vbTab, vbCrLf)
            bReload = True
        Next lIndex
        
        If bReload Then
            KillFile AddSlash(App.Path) & "SymPool.MEM"
        End If
    End If
    
    ' Import the quote board tab...
    If bContinue Then
        If astrQuoteBoard.Size > 0 Then
            If lExistingTab = -1& Then
                If UCase(Parse(astrQuoteBoard(astrQuoteBoard.Size - 1), vbTab, 1)) = "(FILTER)" Then
                    astrQuoteBoard.Add strTabInfo, astrQuoteBoard.Size - 1
                Else
                    astrQuoteBoard.Add strTabInfo
                End If
            Else
                astrQuoteBoard(lExistingTab) = strTabInfo
            End If
            
            FileFromString AddSlash(App.Path) & "Custom\QuoteBoard.INF", astrQuoteBoard.JoinFields(vbLf)
        End If
    End If
    
    ' Import the alerts...
    If bContinue Then
        bReload = False
        astrAlertsFile.FromFile AddSlash(App.Path) & "Custom\QuoteList.ALR"
        
        For lIndex = 0 To astrAlerts.Size - 1
            If AlertExistsInFile(astrAlerts(lIndex), astrAlertsFile) = False Then
                astrAlertsFile.Add astrAlerts(lIndex)
                bReload = True
            End If
        Next lIndex
        
        If bReload Then
            astrAlertsFile.ToFile AddSlash(App.Path) & "Custom\QuoteList.ALR"
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mDataNav.ImportQuoteBoardTab"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AlertExistsInFile
'' Description: Does the given alert exist in the alerts file?
'' Inputs:      Alert String, Alerts File
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AlertExistsInFile(ByVal strAlert As String, Optional ByVal astrAlerts As cGdArray = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim strThisAlert As String          ' Alert string for this alert
    Dim astrThisAlert As cGdArray       ' This alert split out into an array
    Dim astrThatAlert As cGdArray       ' The passed in alert split out into an array
    Dim lAlert As Long                  ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastChecked As Long            ' Index for the Last Checked value
    Dim lActive As Long                 ' Index for the Active value
    Dim bReturn As Boolean              ' Return value for the function
    
    Set astrThisAlert = New cGdArray
    Set astrThatAlert = New cGdArray
    
    bReturn = False
    
    If astrAlerts Is Nothing Then
        Set astrAlerts = New cGdArray
        astrAlerts.FromFile AddSlash(App.Path) & "Custom\QuoteList.ALR"
    End If
    
    If astrAlerts.Size > 1 Then
        astrThisAlert.SplitFields strAlert, "|"
        
        For lAlert = 1 To astrAlerts.Size - 1
            astrThatAlert.SplitFields astrAlerts(lAlert), "|"
            
            If astrThisAlert.Size = astrThatAlert.Size Then
                Select Case CLng(Val(astrThatAlert(0)))
                    Case eGDAlertType_QuoteBoard
                        lLastChecked = 13
                        lActive = 7
                    Case eGDAlertType_AutoTrade
                        lLastChecked = 8
                        lActive = 4
                    Case eGDAlertType_Status
                        lLastChecked = 6
                        lActive = 2
                    Case eGDAlertType_Price
                        lLastChecked = 9
                        lActive = 7
                    Case eGDAlertType_Time
                        lLastChecked = 11
                        lActive = 5
                    Case eGDAlertType_Chart
                        lLastChecked = 11
                        lActive = 7
                    Case eGDAlertType_Annot
                        lLastChecked = 11
                        lActive = 7
                    Case eGDAlertType_TradeSense
                        lLastChecked = 8
                        lActive = 4
                    Case Else
                        lLastChecked = -1
                        lActive = -1
                End Select
            
                bReturn = True
                For lIndex = 0 To astrThisAlert.Size - 1
                    If (lIndex <> lLastChecked) And (lIndex <> lActive) Then
                        If UCase(astrThisAlert(lIndex)) <> UCase(astrThatAlert(lIndex)) Then
                            bReturn = False
                            Exit For
                        End If
                    End If
                Next lIndex
                
                If bReturn = True Then
                    Exit For
                End If
            End If
        Next lAlert
    End If

    AlertExistsInFile = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.AlertExistsInFile"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameCriteriaFile
'' Description: Allow the user to rename the custom criteria file
'' Inputs:      Old Filename (without path), Old Criteria Name
'' Returns:     True if renamed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RenameCriteriaFile(ByVal strOldFileBase As String, ByVal strOldName As String) As Boolean
On Error GoTo ErrSection:

    Dim strOldFilePath As String        ' Old filename with path
    Dim strNewFileBase As String        ' Base filename for the new criteria
    Dim strNewFilePath As String        ' New filename for the criteria
    Dim lIndex As Long                  ' Index into a for loop
    Dim Filter As cFilter               ' Filter object
    Dim strSelection As String          ' Current selection in the symbol grid
    Dim lFieldNum As Long               ' ID of the field in the symbol pool
    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
        
    strNewFileBase = InfBox("What is the new filename for '" & strOldName & "'?", "?", "+OK|-Cancel", "Rename Criteria File", , , , , , "string", strOldFileBase)
    If (Len(strNewFileBase) > 0) And (strNewFileBase <> strOldFileBase) Then
        strNewFileBase = FileBase(strNewFileBase) & ".SCN"
        strNewFilePath = AddSlash(App.Path) & "Custom\" & strNewFileBase
        strOldFilePath = AddSlash(App.Path) & "Custom\" & strOldFileBase
        
        If (Left(UCase(FileBase(strNewFileBase)), 3) = "CUS") And (IsNumeric(Right(FileBase(strNewFileBase), 5))) Then
            InfBox "Cannot rename a criteria file to a custom filename", "!", , "Rename Criteria Error"
        ElseIf IsValidFileBase(strNewFileBase, False) = False Then
            InfBox strNewFileBase & " is not a valid filename.  Could not rename file.", "!", , "Rename Criteria Error"
        ElseIf FileExist(strNewFilePath) Then
            InfBox strNewFileBase & " already exists.  Could not rename filename.", "!", , "Rename Criteria Error"
        ElseIf Not FileExist(strOldFilePath) Then
            InfBox strOldFileBase & " does not exist", "!", , "Rename Criteria Error"
        Else
            ' Rename file on hard drive...
            If RenameFile(strOldFilePath, strNewFilePath) Then
                InfBox "Updating forms.  Please wait...", , , "Rename Criteria", True
                
                ' Change the ID's in the symbol pool for appropriate items...
                g.SymbolPool.Criterias.Key(strOldFileBase) = strNewFileBase
                g.SymbolPool.Criterias(strNewFileBase).ID = strNewFileBase
                
                lFieldNum = g.SymbolPool.FieldNumForID("DSV:" & strOldFileBase)
                If lFieldNum <> -1& Then
                    g.SymbolPool.FieldID(lFieldNum) = "DSV:" & strNewFileBase
                End If
                lFieldNum = g.SymbolPool.FieldNumForID("DSP:" & strOldFileBase)
                If lFieldNum <> -1& Then
                    g.SymbolPool.FieldID(lFieldNum) = "DSP:" & strNewFileBase
                End If
            
                ' Make sure that any filters that were looking at the old criteria are now looking at
                ' the new criteria instead...
                For lIndex = 1 To g.SymbolPool.Filters.Count
                    Set Filter = g.SymbolPool.Filters(lIndex)
                    Filter.RenameCriteria strOldFileBase, strNewFileBase
                Next lIndex
                
                ' Change quote board field if necessary...
                frmQuotes.RenameCriteriaFile strOldFileBase, strNewFileBase
                
                ' Change symbol grid if necessary...
                frmSymbolGrid.RenameCriteria strOldFileBase, strNewFileBase
            
                ' Change snapshot form if necessary...
                frmSnapshot.RenameCriteria strOldFileBase, strNewFileBase
                
                InfBox ""
                bReturn = True
            Else
                InfBox strOldFileBase & " could not be renamed", "!", , "Rename Criteria Error"
            End If
        End If
    End If
    
    RenameCriteriaFile = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.RenameCriteriaFile"
    
End Function

' Returns the difference between the local clock and the Genesis servers (after adjusting for time zone)
' - uses the NYTIME from the Request.inf file which gets sent down in a download
Public Function ClockDiff() As Double
On Error Resume Next

    Dim s$, dNY#, iMonth&, iDay&, iYear&, iHour&, iMin&
    Dim aStrings As New cGdArray
    
    ' NYTIME=August 1, 2008 12:52pm ET
    s = UCase(FileToString(App.Path & "\Ftp\Request.inf", 50, True))
    If InStr(s, "=") > 0 Then
        s = Replace(s, "=", " ")
        s = Replace(s, ",", " ")
        s = Replace(s, ":", " ")
        aStrings.SplitFields s, " "
        ' Space-delimited fields become ...
        '0: NYTIME
        '1: AUGUST
        '2: 1
        '3: 2008
        '4: 12
        '5: 52PM
        '6: ET
        iMonth = MonthNumber(aStrings(1))
        iDay = Val(aStrings(2))
        iYear = Val(aStrings(3))
        iHour = Val(aStrings(4))
        iMin = Val(Left(aStrings(5), 2))
        If iHour = 12 Then iHour = 0
        If InStr(aStrings(5), "PM") > 0 Then
            iHour = iHour + 12
        End If
        If iMonth >= 1 And iMonth <= 12 And iDay >= 1 And iDay <= 31 And iYear > 2000 _
            And iHour >= 0 And iHour < 24 And iMin >= 0 And iMin < 60 Then
                dNY = DateSerial(iYear, iMonth, iDay) + TimeSerial(iHour, iMin, 0)
        End If
        If dNY > 0 Then
            ClockDiff = ConvertTimeZone(Now, "", "NY") - dNY
        End If
        Set aStrings = Nothing
    End If
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsPitFuture
'' Description: Is the given symbol a pit future?
'' Inputs:      Symbol
'' Returns:     True if Pit Future, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsPitFuture(ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    Static strPitSymbols As String      ' List of pit symbols
    Dim astrFile As cGdArray            ' Symbol map file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strPitSymbol As String          ' Pit symbol on the line
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If SecurityType(strSymbol) = "F" Then
        If Len(strPitSymbols) = 0 Then
            Set astrFile = New cGdArray
            If astrFile.FromFile(AddSlash(App.Path) & "Info\SymbolMap.CSV") Then
                For lIndex = 0 To astrFile.Size - 1
                    strPitSymbol = Parse(astrFile(lIndex), ",", 1)
                    If Len(strPitSymbol) > 0 Then
                        strPitSymbols = strPitSymbols & "," & strPitSymbol & ","
                    End If
                Next lIndex
            End If
        End If
        
        If Len(strPitSymbols) > 0 Then
            bReturn = (InStr(strPitSymbols, "," & Parse(strSymbol, "-", 1) & ",") <> 0)
        End If
    End If
    
    IsPitFuture = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.IsPitFuture"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChopDailyBars
'' Description: Chop the given daily bars with the given dates
'' Inputs:      Bars, Begin time, End time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChopDailyBars(Bars As cGdBars, Optional ByVal dBeginDateTime As Double = kNullData, Optional ByVal dEndDateTime As Double = kNullData, Optional ByVal bTickLevel As Boolean = False)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of something
    Dim lSessionDate As Long            ' Session date for the time given
    Dim TickData As New cGdBars         ' Bars object for the tick data for a day

    If Not Bars Is Nothing Then
        If (UCase(Bars.Prop(eBARS_PeriodicityStr)) = "DAILY") And (Bars.Size > 0) Then
            If dBeginDateTime <> kNullData Then
                lSessionDate = Bars.SessionDateForTradeTime(dBeginDateTime)
            End If
            If (dBeginDateTime <> kNullData) And (lSessionDate >= Bars(eBARS_DateTime, 0)) Then
                If lSessionDate > Bars(eBARS_DateTime, Bars.Size - 1) Then
                    Bars.Size = 0
                Else
                    lPos = Bars.FindDateTime(lSessionDate)
                    If lPos > 0 Then
                        Bars.DeleteFirstBars lPos
                    End If
                    
                    If bTickLevel Then
                        If DM_GetBars(TickData, Bars.SymbolOrSymbolID, ePRD_EachTick, lSessionDate, lSessionDate) Then
                            g.RealTime.SpliceBars TickData
                            TickData.DeleteFirstBars TickData.FindDateTime(dBeginDateTime)
                            
                            Bars(eBARS_Open, 0) = TickData(eBARS_Close, 0)
                            Bars(eBARS_High, 0) = gdMaxValue(TickData.ArrayHandle(eBARS_Close), 0, TickData.Size - 1)
                            Bars(eBARS_Low, 0) = gdMinValue(TickData.ArrayHandle(eBARS_Close), 0, TickData.Size - 1)
                            Bars(eBARS_Close, 0) = TickData(eBARS_Close, TickData.Size - 1)
                        End If
                    Else
                        If DM_GetBars(TickData, Bars.SymbolOrSymbolID, "1 Minute", lSessionDate, lSessionDate) Then
                            g.RealTime.SpliceBars TickData
                            TickData.DeleteFirstBars TickData.FindDateTime(dBeginDateTime) + 1
                            
                            Bars(eBARS_Open, 0) = TickData(eBARS_Open, 0)
                            Bars(eBARS_High, 0) = gdMaxValue(TickData.ArrayHandle(eBARS_High), 0, TickData.Size - 1)
                            Bars(eBARS_Low, 0) = gdMinValue(TickData.ArrayHandle(eBARS_Low), 0, TickData.Size - 1)
                            Bars(eBARS_Close, 0) = TickData(eBARS_Close, TickData.Size - 1)
                        End If
                    End If
                End If
            End If
        
            If dEndDateTime <> kNullData Then
                lSessionDate = Bars.SessionDateForTradeTime(dEndDateTime)
            End If
            If (dEndDateTime <> kNullData) And (lSessionDate <= Bars(eBARS_DateTime, Bars.Size - 1)) Then
                If lSessionDate < Bars(eBARS_DateTime, 0) Then
                    Bars.Size = 0
                Else
                    lPos = Bars.FindDateTime(lSessionDate)
                    If lPos > 0 Then
                        Bars.DeleteSomeBars lPos + 1, Bars.Size - lPos
                    End If
                    
                    If bTickLevel Then
                        If DM_GetBars(TickData, Bars.SymbolOrSymbolID, ePRD_EachTick, lSessionDate, lSessionDate) Then
                            lPos = TickData.FindDateTime(dEndDateTime)
                            TickData.DeleteSomeBars lPos + 1, TickData.Size - lPos
                            
                            Bars(eBARS_Open, 0) = TickData(eBARS_Close, 0)
                            Bars(eBARS_High, 0) = gdMaxValue(TickData.ArrayHandle(eBARS_Close), 0, TickData.Size - 1)
                            Bars(eBARS_Low, 0) = gdMinValue(TickData.ArrayHandle(eBARS_Close), 0, TickData.Size - 1)
                            Bars(eBARS_Close, 0) = TickData(eBARS_Close, TickData.Size - 1)
                        End If
                    Else
                        If DM_GetBars(TickData, Bars.SymbolOrSymbolID, "1 Minute", lSessionDate, lSessionDate) Then
                            lPos = TickData.FindDateTime(dEndDateTime)
                            TickData.DeleteSomeBars lPos + 1, TickData.Size - lPos
                            
                            Bars(eBARS_Open, 0) = TickData(eBARS_Open, 0)
                            Bars(eBARS_High, 0) = gdMaxValue(TickData.ArrayHandle(eBARS_High), 0, TickData.Size - 1)
                            Bars(eBARS_Low, 0) = gdMinValue(TickData.ArrayHandle(eBARS_Low), 0, TickData.Size - 1)
                            Bars(eBARS_Close, 0) = TickData(eBARS_Close, TickData.Size - 1)
                        End If
                    End If
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mDataNav.ChopDailyBars"
    
End Sub

' Returns true if local time zone is in North/South America
' (i.e. basically if the local time is less than GMT - 1 hour),
' or false if local time zone is in Europe/Asia/Africa
Public Function IsInWesternHemisphere() As Boolean
On Error GoTo ErrSection:

    Dim dtLocal#
    ' convert a summer date (e.g. 40000 = July 6 2009) from GMT to local
    ' (there's a bigger time diff between Brazil and Iceland in the summer)
    dtLocal = ConvertTimeZone(40000, "GMT", "")
    ' check if it is < GMT - 1 hour
    If dtLocal < 40000 - 1.1 / 24# Then
        IsInWesternHemisphere = True
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mDataNav.IsInWesternHemisphere"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SfeAllowed
'' Description: Is the user allowed to see the SFE data?
'' Inputs:      Grace days
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SfeAllowed(Optional ByVal lGraceDays As Long = 30) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If HasModule("E_SFE") Then
        bReturn = True
    ElseIf Len(Trim(GetRegistryValue(rkLocalMachine, "Software\Trader Workstation", "jtspath", ""))) > 2 Then
        bReturn = True
    ElseIf InStr(UCase(GetRegistryValue(rkClassesRoot, "tws", "", "")), "TRADER") > 0 Then ' newer version of TWS
        bReturn = True
    Else
        bReturn = SfeAllowedByBroker(lGraceDays)
    End If
    
    SfeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.SfeAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SfeAllowedByBroker
'' Description: Is the SFE data allowed by a broker?
'' Inputs:      Grace days
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SfeAllowedByBroker(Optional ByVal lGraceDays As Long = 30) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrExchanges As New cGdArray   ' List of exchanges
    Dim lLastDate As Long               ' Last successful connection to broker
    Dim lIndex As Long                  ' Index into a for loop
    
#If 0 Then
    bReturn = False
    If g.Broker.IsBrokerUser(eTT_AccountType_Ideal) Then
        lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", AddSlash(App.Path) & "Ideal.INI"))))
        
        If lLastDate >= Date - lGraceDays Then
            If FileExist(AddSlash(App.Path) & "Provided\LkExch.IDL") Then
                astrExchanges.Serialize AddSlash(App.Path) & "Provided\LkExch.IDL", False
                bReturn = astrExchanges.BinarySearch("SNFE")
            End If
        End If
    End If
    
    If (bReturn = False) And g.Broker.IsBrokerUser(eTT_AccountType_IntBrokers) Then
        lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", AddSlash(App.Path) & "IntBrokers2.INI"))))
        
        If lLastDate >= Date - lGraceDays Then
            If FileExist(AddSlash(App.Path) & "Provided\LkExch.IB") Then
                astrExchanges.Serialize AddSlash(App.Path) & "Provided\LkExch.IB", False
                bReturn = astrExchanges.BinarySearch("SNFE")
            End If
        End If
    End If
    
    If (bReturn = False) And g.Broker.IsBrokerUser(eTT_AccountType_ManExpress) Then
        lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", AddSlash(App.Path) & "ManExpress.INI"))))
        
        If lLastDate >= Date - lGraceDays Then
            bReturn = True
        End If
    End If
    
    If (bReturn = False) And g.Broker.IsBrokerUser(eTT_AccountType_LindWaldock) Then
        lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", AddSlash(App.Path) & "LindWaldock.INI"))))
        
        If lLastDate >= Date - lGraceDays Then
            bReturn = True
        End If
    End If
#Else
    bReturn = False
    If Not g.Broker Is Nothing Then
        For lIndex = 1 To kNumBrokers - 1
            If g.Broker.IsBrokerUser(lIndex) Then
                lLastDate = g.Broker.SfeAllowed(lIndex)
                If lLastDate <> kNullData Then
                    If lLastDate > Int(CurrentTime) - lGraceDays Then
                        bReturn = True
                        Exit For
                    End If
                End If
            End If
        Next lIndex
    End If
#End If
    
    SfeAllowedByBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.SfeAllowedByBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidForexTradingTime
'' Description: Is the given time a valid Forex trading time?
'' Inputs:      Symbol, Time in NY
'' Returns:     True if valid trading time, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidForexTradingTime(ByVal strSymbol As String, Optional ByVal dNyTime As Double = kNullData) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim dStartTime As Double            ' Starting time for the session
    Dim dEndTime As Double              ' Ending time for the session
    Dim strTimes As String              ' Times read in from the INI file
    Dim strPropName As String           ' Property name for the INI file
    Dim strBrokerCode As String         ' Broker code from the symbol
    
    bReturn = True
    If IsForex(strSymbol) Then
        If InStr(strSymbol, "@") <> 0 Then
            strBrokerCode = Parse(strSymbol, "@", 2)
            strBrokerCode = UCase(Left(strBrokerCode, 1)) & LCase(Mid(strBrokerCode, 2))
            strPropName = strBrokerCode & "Fx"
            strTimes = GetIniFileProperty(strPropName, "", "SessionTimes", AddSlash(App.Path) & "Provided\Provided.INI")
        End If
    End If
    
    If Len(strTimes) > 0 Then
        If dNyTime = kNullData Then
            dNyTime = CurrentTime("NY")
        End If
        dStartTime = Int(dNyTime) + HHMMtoMinutes(Parse(strTimes, "-", 1)) / 1440#
        dEndTime = Int(dNyTime) + HHMMtoMinutes(Parse(strTimes, "-", 2)) / 1440#
        
        If Weekday(dNyTime) = vbSaturday Then
            bReturn = False
        ElseIf (Weekday(dNyTime) = vbFriday) And (dNyTime > dEndTime) Then
            bReturn = False
        ElseIf (Weekday(dNyTime) = vbSunday) And (dNyTime < dStartTime) Then
            bReturn = False
        ElseIf dStartTime > dEndTime Then
            bReturn = (dNyTime >= dStartTime) Or (dNyTime < dEndTime)
        Else
            bReturn = (dNyTime >= dStartTime) And (dNyTime < dEndTime)
        End If
    End If
    
    ValidForexTradingTime = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.ValidForexTradingTime"
    
End Function

Public Sub DoPFCheck()

    Dim i&, nID&, nDate&, nWeek&, nMinWeeks&, nTrades&
    Dim dProfit#, dWin#, dLoss#, dPF#
    Dim s$, bCheck As Boolean, bLiveOnly As Boolean
    Dim aWin As New cGdArray, aLoss As New cGdArray, aNumWin As New cGdArray, aNumLoss As New cGdArray
    Dim aFile As New cGdArray
    Dim rsFills As Recordset, rsAccts As Recordset
    Dim Acct As cPtAccount
    Static bInProgress As Boolean
   
    If bInProgress Or g.bUnloading Or g.bStarting Then Exit Sub
    If g.dbPaper Is Nothing Then Exit Sub
    
    If Not IsIDE Then
        On Error Resume Next
    End If
    
    nMinWeeks = Val(GetProvidedProperty("PFCheck"))
    If nMinWeeks > 0 Then
        bLiveOnly = True
    Else
        nMinWeeks = Abs(nMinWeeks)
        bLiveOnly = False
    End If
    If nMinWeeks = 0 Then Exit Sub
        
    Set rsAccts = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
    If rsAccts.BOF And rsAccts.EOF Then Exit Sub
    
    bInProgress = True
    Do While Not rsAccts.EOF And Not g.bUnloading
        bCheck = True
        ' !AccountID, !AccountNumber, !Name, !Broker, !AccountType
        nID = rsAccts!AccountID
        Set Acct = g.Broker.Account(nID)
        If Acct Is Nothing Then
            bCheck = False
        ElseIf bLiveOnly Then
            If Not g.Broker.IsLiveAccount(Acct.AccountType) Then
                bCheck = False
            End If
        End If
        If bCheck Then
            aWin.Create eGDARRAY_Doubles, 520, 0
            aLoss.Create eGDARRAY_Doubles, aWin.Size, 0
            aNumWin.Create eGDARRAY_Longs, aWin.Size, 0
            aNumLoss.Create eGDARRAY_Longs, aWin.Size, 0
        
            Set rsFills = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] WHERE [AccountID]=" & Str(nID) & ";", dbOpenDynaset)
            If Not (rsFills.BOF And rsFills.EOF) Then
                rsFills.MoveFirst
                Do While Not rsFills.EOF And Not g.bUnloading
                    dProfit = 0
                    nDate = 0
                    dProfit = rsFills!ClosedProfit
                    nDate = rsFills!SessionDate
                    If dProfit <> kNullData And dProfit <> 0 And nDate > 0 Then
                        nWeek = WkNum(Date) - WkNum(nDate)
                        If nWeek >= 0 And nWeek < aWin.Size Then
                            If dProfit > 0 Then
                                aWin.Num(nWeek) = aWin.Num(nWeek) + dProfit
                                aNumWin.Num(nWeek) = aNumWin.Num(nWeek) + 1
                            Else
                                aLoss.Num(nWeek) = aLoss.Num(nWeek) + dProfit
                                aNumLoss.Num(nWeek) = aNumLoss.Num(nWeek) + 1
                            End If
                        End If
                    End If
                    rsFills.MoveNext
                Loop
            End If
            Set rsFills = Nothing
            
            dWin = 0
            dLoss = 0
            dPF = 0
            nTrades = 0
            For nWeek = 0 To aWin.Size
                If aNumWin.Num(nWeek) + aNumLoss.Num(nWeek) > 0 Then
                    If nTrades = 0 Then
                        If nWeek > nMinWeeks Then
                            Exit For ' skip if no recent trading activity
                        End If
                        aFile.Add Acct.AccountNumber & vbTab & Acct.Name & vbTab & g.Broker.BrokerName(Acct.AccountType) _
                                & vbTab & Str(Acct.ConnectionStatus) & vbTab & Str(Int(Acct.ClosedProfit)) & vbTab & Str(Int(Acct.OpenProfit))
                    End If
                    nTrades = nTrades + aNumWin.Num(nWeek) + aNumLoss.Num(nWeek)
                    
                    dWin = dWin + aWin.Num(nWeek)
                    dLoss = dLoss + aLoss.Num(nWeek)
                    If dWin = 0 Then
                        dPF = 0
                    ElseIf dLoss = 0 Then
                        dPF = 99
                    Else
                        dPF = dWin / Abs(dLoss)
                    End If
                    
                    aFile.Add Str(nWeek) & vbTab & Str(aNumWin.Num(nWeek)) & vbTab & Str(aNumLoss.Num(nWeek)) _
                        & vbTab & Str(Int(aWin.Num(nWeek))) & vbTab & Str(Int(aLoss.Num(nWeek))) _
                        & vbTab & Str(Int(dWin + dLoss)) & vbTab & Format(dPF, "#0.000")
                End If
            Next
            If nTrades > 0 Then
                i = 0
                If nTrades > 10 Then
                    If dPF > 10 Then dPF = 10
                    i = Int(dPF * 10)
                End If
                aFile.Add String(i + 1, "#")
            End If
        End If
        Set Acct = Nothing
        rsAccts.MoveNext
    Loop
    Set rsAccts = Nothing
    
    If aFile.Size > 1 And Not g.bUnloading Then
        s = Format(Date, "yyyy-mm-dd") & vbTab & Format(RI_GetDataServiceID / 1000, "#000000") & ":" & Format(RI_GetDataServiceID Mod 1000, "000") _
                & vbTab & Str(g.RealTime.ConnectionStatus)
        aFile.Add s, 0
        If IsIDE And frmTest.Visible Then
            For i = 0 To aFile.Size - 1
                frmTest.AddList aFile(i)
            Next
        End If
        
        SendWebPage RI_GetMachineID & "-$.txt", aFile.JoinFields(vbCrLf)
    End If

    bInProgress = False

End Sub

' see if any special things for this MachineID (just testing this for now)
Public Function GetMidCmd() As String
If Not IsIDE Then
    On Error Resume Next
End If

    Dim i&, s$, dt#, nID&, strCmd$
    Dim aList As New cGdArray
    Static strReturn$, dtPrev#

    ' only need to re-check if the datetime of the file has changed
    s = App.Path & "\Info\MidCmd.cfg"
    dt = FileDate(s)
    If dt <> dtPrev Then
        dtPrev = dt
        strReturn = ""
        aList.FromFile s
        For i = 0 To aList.Size - 1
            s = DecryptFromHex(aList(i))
            If UCase(Parse(s, vbTab, 1)) = RI_GetMachineID Then
                strCmd = Parse(s, vbTab, 2)
                If strCmd = "-" Then
                    ' check for a bad MachID/DSID combination
                    nID = RI_GetDataServiceID
                    If nID > 0 And nID = Val(Parse(s, vbTab, 3)) Then
                        RI_RegenerateMID
                        strReturn = ""
                        Exit For
                    End If
                ElseIf Len(strCmd) > 1 Then
                    strReturn = s
                End If
            End If
        Next
        Set aList = Nothing
    End If

    GetMidCmd = strReturn
End Function


' to assign the Volume Iterator "color flag" to the Minute bars
' (the "Flags" array values will be set such that: Null = Black, 0 = Green, 1 = Red)
Public Function GetVolumeIterators(Bars As cGdBars, Optional ByVal nNumLookbackBars& = 300, Optional ByVal nNumTopVolumeBars& = 20) As Boolean
On Error GoTo ErrSection:

    Dim i&, nCount&, dVolume#, nStartBar&
    Dim aVolume As New cGdArray

    ' make sure parms are valid
    If nNumLookbackBars > 0 And nNumTopVolumeBars > 0 And nNumTopVolumeBars < nNumLookbackBars And Bars.Size > 0 Then
        ' add the "Flags" array to the bars to designate the color flag
        Bars.ArrayMask = Bars.ArrayMask Or eBARS_Flags
        
        ' get the Volumes for the lookback bars
        aVolume.Create eGDARRAY_Doubles, nNumLookbackBars, 0
        aVolume.Size = 0
        nStartBar = 0
        For i = Bars.Size - 1 To 0 Step -1
            ' make sure Close and Volume data is valid on this bar (i.e. non-Null)
            If Bars(eBARS_Close, i) <> kNullData Then
                dVolume = Bars(eBARS_Vol, i)
                If dVolume >= 0 Then
                    aVolume.Add dVolume
                    nCount = nCount + 1
                    If nCount >= nNumLookbackBars Then
                        nStartBar = i
                        Exit For
                    End If
                End If
            End If
        Next
                
        If aVolume.Size > nNumTopVolumeBars Then
            ' sort the Volumes in order to find the "Nth" highest volume
            aVolume.Sort eGdSort_Descending
            dVolume = aVolume(nNumTopVolumeBars - 1)
            ' assign "1" for all volumes above the Nth highest volume
            For i = nStartBar To Bars.Size - 1
                If Bars(eBARS_Vol, i) >= dVolume Then
                    Bars(eBARS_Flags, i) = 1
                Else
                    Bars(eBARS_Flags, i) = 0
                End If
            Next
            GetVolumeIterators = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.GetVolumeIterators"
End Function

' SPECIALIZED: this routine is only used by the Market Profile stuff ...
' Maintains a set of "bars" to contain a record for each price traded during each 5-minute period
' (arrays used are: Date.Time, Close, Volume, BidVol, AskVol, Flags=#Trades)
' - if nStartSessionDate > 0 then it's a date, or if <= 0 then indicates # of days back
' - if nEndSessionDate > 0 then it's a date, or if = 0 then indicates most recent session
Public Function BuildProfileBars(ProfileBars As cGdBars, ByVal nSymbolID&, ByVal nStartSessionDate&, ByVal nEndSessionDate&) As Boolean
On Error GoTo ErrSection:

    Dim d#, i&, nDate&, strSymbol$, bAppend As Boolean, bPrepend As Boolean
    Dim Ticks As cGdBars, ProfileDay As cGdBars
    
    If nSymbolID = 0 Then Exit Function
        
    ' need to clear the bars data whenever the symbol has changed
    If nSymbolID <> ProfileBars.Prop(eBARS_SymbolID) Or ProfileBars.ArrayMask <> eBARS_Profiled Then
        ProfileBars.Size = 0
        SetBarProperties ProfileBars, nSymbolID
        ProfileBars.Prop(eBARS_PeriodicityStr) = ProfileBars.Prop(eBARS_PeriodicityStr)
        ProfileBars.ArrayMask = eBARS_Profiled
    End If
    strSymbol = GetSymbol(nSymbolID)
    
    ' get date range
    If nEndSessionDate > 19000000 Then
        nEndSessionDate = JulFromLong(nEndSessionDate)
    ElseIf nEndSessionDate <= 0 Then
        nEndSessionDate = Date + 1
    End If
    If nStartSessionDate > 19000000 Then
        nStartSessionDate = JulFromLong(nStartSessionDate)
    ElseIf nStartSessionDate <= 0 Then
        If nEndSessionDate < LastDailyDownload Then
            nStartSessionDate = nEndSessionDate + nStartSessionDate
        Else
            nStartSessionDate = LastDailyDownload + nStartSessionDate
        End If
    End If
    
    ' make sure no gaps get created in the date ranges
    If ProfileBars.Size > 0 Then
        If nStartSessionDate > ProfileBars.SessionDate(ProfileBars.Size - 1) Then
            nStartSessionDate = ProfileBars.SessionDate(ProfileBars.Size - 1) + 1
        End If
        If nEndSessionDate < ProfileBars.SessionDate(0) Then
            nEndSessionDate = ProfileBars.SessionDate(0) - 1
        End If
    End If
        
  
    ' APPEND any session dates not yet added
    ' add one day of profile bars at a time (if not already in the data)
    For nDate = nStartSessionDate To nEndSessionDate
        bAppend = False
        bPrepend = False
        If IsWeekday(nDate) Then
            If ProfileBars.Size = 0 Or nDate > ProfileBars.SessionDate(ProfileBars.Size - 1) Then
                bAppend = True
            ElseIf nDate < ProfileBars.SessionDate(0) Then
                bPrepend = True
            End If
            If bAppend Then
                ' load the full ticks for this session
                Set Ticks = New cGdBars
'd = gdTickCount
                'DM_GetBars Ticks, strSymbol, "Each Tick", nDate, nDate
                i = 0
                GetAvailTickData Ticks, i, strSymbol, nSymbolID, nDate, 0
'frmTest.AddList "GetBars " & Str(gdTickCount - d)
                If Ticks.Size > 0 And i = nDate Then
                    ' build profile bars from the ticks
                    Set ProfileDay = ProfileBars.MakeCopy(True)
'd = gdTickCount
                    If ProfileDay.BuildBars(ProfileBars.Prop(eBARS_PeriodicityStr), Ticks.BarsHandle) Then
                        ' add to the big set
                        If ProfileDay.Size > 0 Then
                            gdAppendBars ProfileBars.BarsHandle, ProfileDay.BarsHandle, Abs(bPrepend)
                            BuildProfileBars = True ' return True if any bars added
                        End If
                    End If
'frmTest.AddList "BuildBars " & Str(gdTickCount - d)
                End If
            End If
        End If
    Next
    
    ' Then PREPEND any session dates not yet added
    ' add one day of profile bars at a time (if not already in the data)
    For nDate = nEndSessionDate To nStartSessionDate Step -1
        bAppend = False
        bPrepend = False
        If IsWeekday(nDate) Then
            If ProfileBars.Size = 0 Or nDate > ProfileBars.SessionDate(ProfileBars.Size - 1) Then
                bAppend = True
            ElseIf nDate < ProfileBars.SessionDate(0) Then
                bPrepend = True
            End If
            If bPrepend Then
                ' load the full ticks for this session
                Set Ticks = New cGdBars
'd = gdTickCount
                'DM_GetBars Ticks, strSymbol, "Each Tick", nDate, nDate
                i = 0
                GetAvailTickData Ticks, i, strSymbol, nSymbolID, nDate, 0
'frmTest.AddList "GetBars " & Str(gdTickCount - d)
                If Ticks.Size > 0 And i = nDate Then
                    ' build profile bars from the ticks
                    Set ProfileDay = ProfileBars.MakeCopy(True)
'd = gdTickCount
                    If ProfileDay.BuildBars(ProfileBars.Prop(eBARS_PeriodicityStr), Ticks.BarsHandle) Then
                        ' add to the big set
                        If ProfileDay.Size > 0 Then
                            gdAppendBars ProfileBars.BarsHandle, ProfileDay.BarsHandle, Abs(bPrepend)
                            BuildProfileBars = True ' return True if any bars added
                        End If
                    End If
'frmTest.AddList "BuildBars " & Str(gdTickCount - d)
                End If
            End If
        End If
    Next
    
    Set Ticks = Nothing
    Set ProfileDay = Nothing

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.BuildProfileBars"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IndexOfLastNonDigit
'' Description: Find the index of the last non-digit
'' Inputs:      String
'' Returns:     Index of the last non-digit (-1 if none)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IndexOfLastNonDigit(ByVal strString As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    For lIndex = Len(strString) To 1 Step -1
        If IsDigit(Mid(strString, lIndex, 1)) = False Then
            lReturn = lIndex
            Exit For
        End If
    Next lIndex
    
    IndexOfLastNonDigit = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.IndexOfLastNonDigit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SpreadComponentsForBars
'' Description: Determine the components of the given Futures Spread
'' Inputs:      Bars
'' Returns:     Array of Components
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SpreadComponentsForBars(Bars As cGdBars) As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of spread components to return
    Dim strSymbol As String             ' Symbol for the Bars passed in
    Dim astrContracts As cGdArray       ' Contacts that make up the spread
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    If Not Bars Is Nothing Then
        strSymbol = Bars.Prop(eBARS_Symbol)
        If IsSpreadSymbol(strSymbol) Then
            Set astrContracts = SpreadContractsFromDescription(Bars.Prop(eBARS_Desc))
            If astrContracts.Size = 2 Then
                astrReturn.Add Bars.Prop(eBARS_BaseSymbol) & "-" & astrContracts(0)
                astrReturn.Add Bars.Prop(eBARS_BaseSymbol) & "-" & astrContracts(1)
            End If
        End If
    End If
    
    Set SpreadComponentsForBars = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.SpreadComponentsForBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SpreadComponentsForSymbol
'' Description: Determine the components of the given Futures Spread symbol
'' Inputs:      Symbol
'' Returns:     Array of Components
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SpreadComponentsForSymbol(ByVal strSymbol As String) As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of spread components to return
    Dim lPoolRec As Long                ' Symbol pool record number
    Dim strDescription As String        ' Description from the symbol pool
    Dim astrContracts As cGdArray       ' Contacts that make up the spread
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    If IsSpreadSymbol(strSymbol) Then
        lPoolRec = g.SymbolPool.PoolRecForSymbol(strSymbol)
        If lPoolRec > -1& Then
            strDescription = g.SymbolPool.Desc(lPoolRec)
            
            Set astrContracts = SpreadContractsFromDescription(strDescription)
            If astrContracts.Size = 2 Then
                astrReturn.Add Parse(strSymbol, "-", 1) & "-" & astrContracts(0)
                astrReturn.Add Parse(strSymbol, "-", 1) & "-" & astrContracts(1)
            End If
        End If
    End If
    
    Set SpreadComponentsForSymbol = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.SpreadComponentsForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SpreadContractsFromDescription
'' Description: Determine the contracts of the given Futures Spread
'' Inputs:      Description
'' Returns:     Array of Components
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SpreadContractsFromDescription(ByVal strDescription As String) As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of spread components to return
    Dim astrDescription As cGdArray     ' Description for the Bars passed in
    Dim lIndex As Long                  ' Index into a for loop
    Dim BrokerSymbol As cBrokerSymbol   ' Broker symbol object for contract conversion
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    If InStr(UCase(strDescription), " SPRD ") > 0 Then
        Set astrDescription = New cGdArray
        astrDescription.SplitFields strDescription, " "
        
        For lIndex = 0 To astrDescription.Size - 1
            If InStr(astrDescription(lIndex), ",") > 0 Then
                Set BrokerSymbol = New cBrokerSymbol
                
                astrReturn.Add BrokerSymbol.ContractFromMMMYY(Parse(astrDescription(lIndex), ",", 1))
                astrReturn.Add BrokerSymbol.ContractFromMMMYY(Parse(astrDescription(lIndex), ",", 2))
                
                Exit For
            End If
        Next lIndex
    End If
    
    Set SpreadContractsFromDescription = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.SpreadContractsFromDescription"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SpreadSymbolForComponents
'' Description: Determine the calendar spread symbol for the given components
'' Inputs:      Components
'' Returns:     Spread Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SpreadSymbolForComponents(ByVal astrComponents As cGdArray) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lCounter As Long                ' Counter variable
    Dim BrokerSymbol As cBrokerSymbol   ' Broker symbol object
    Dim strComponents As String         ' String version of the components
    Dim strSymbol As String             ' Symbol to check
    Dim lPoolRec As Long                ' Record for the symbol in the symbol pool
    Dim strDescription As String        ' Description from the symbol pool
    
    strReturn = ""
    If astrComponents.Size = 2 Then
        lCounter = 0&
        
        Set BrokerSymbol = New cBrokerSymbol
        strComponents = " Sprd " & BrokerSymbol.ContractToMMMYY(Parse(astrComponents(0), "-", 2)) & "," & BrokerSymbol.ContractToMMMYY(Parse(astrComponents(1), "-", 2))
        
        Do While True
            lCounter = lCounter + 1&
            strSymbol = astrComponents(0) & "-S" & Str(lCounter)
            lPoolRec = g.SymbolPool.PoolRecForSymbol(strSymbol)
            If lPoolRec = -1& Then
                Exit Do
            Else
                strDescription = g.SymbolPool.Desc(lPoolRec)
                If InStr(UCase(strDescription), UCase(strComponents)) > 0 Then
                    strReturn = strSymbol
                    Exit Do
                End If
            End If
        Loop
    End If
    
    SpreadSymbolForComponents = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.SpreadSymbolForComponents"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LeadContractForSpread
'' Description: Determine the lead contract for a calendar spread
'' Inputs:      Spread Symbol
'' Returns:     Lead contract
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LeadContractForSpread(ByVal strSpreadSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If IsSpreadSymbol(strSpreadSymbol) Then
        strReturn = Left(strSpreadSymbol, InStr(strSpreadSymbol, "-S") - 1)
    End If
    
    LeadContractForSpread = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.LeadContractForSpread"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsRth
'' Description: Determine if the current time is within regular trading hours
'' Inputs:      Symbol or Symbol ID, Bars, Time to Check
'' Returns:     True if Regular Trading Hours, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsRth(ByVal vSymbolOrSymbolID As Variant, Optional Bars As cGdBars = Nothing, Optional ByVal dTimeToCheck As Double = kNullData) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strKey As String                ' Key into the INI file
    Dim dRthStart As Double             ' Regular trading hours start time
    Dim dRthEnd As Double               ' Reguler trading hours end time
    
    bReturn = True
    If Bars Is Nothing Then
        Set Bars = New cGdBars
        SetBarProperties Bars, vSymbolOrSymbolID
    End If
    If dTimeToCheck = kNullData Then
        dTimeToCheck = CurrentTime(Bars.Prop(eBARS_ExchangeTimeZoneInf), GetSymbol(vSymbolOrSymbolID), True)
    End If
    If SecurityType(Bars) = "F" Then
        strKey = Bars.Prop(eBARS_BaseSymbol) & "-"
    Else
        strKey = Bars.Prop(eBARS_BaseSymbol)
    End If
    
    dRthStart = GetIniFileProperty(strKey, Bars.Prop(eBARS_StartTime), "RthStart", AddSlash(App.Path) & "Provided\Provided.INI")
    dRthEnd = GetIniFileProperty(strKey, Bars.Prop(eBARS_EndTime), "RthEnd", AddSlash(App.Path) & "Provided\Provided.INI")
    
    dRthStart = Int(dTimeToCheck) + (dRthStart / 1440#)
    dRthEnd = Int(dTimeToCheck) + (dRthEnd / 1440#)
    
    If dRthStart > dRthEnd Then
        bReturn = (dTimeToCheck >= dRthStart) Or (dTimeToCheck <= dRthEnd)
    Else
        bReturn = (dTimeToCheck >= dRthStart) And (dTimeToCheck <= dRthEnd)
    End If
    
    DebugLog "IsRth - Symbol=" & GetSymbol(vSymbolOrSymbolID) & ", strKey=" & strKey & ", RthStart=" & DateFormat(dRthStart, MM_DD_YYYY, HH_MM_SS) & ", RthEnd=" & DateFormat(dRthEnd, MM_DD_YYYY, HH_MM_SS) & ", TimeToCheck=" & DateFormat(dTimeToCheck, MM_DD_YYYY, HH_MM_SS) & ", Return=" & Str(bReturn)
    
    IsRth = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.IsRth"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsEth
'' Description: Determine if the current time is within extended trading hours
'' Inputs:      Symbol or Symbol ID, Bars, Time to Check
'' Returns:     True if Extended Trading Hours, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEth(ByVal vSymbolOrSymbolID As Variant, Optional Bars As cGdBars = Nothing, Optional ByVal dTimeToCheck As Double = kNullData) As Boolean
On Error GoTo ErrSection:

    IsEth = Not IsRth(vSymbolOrSymbolID, Bars, dTimeToCheck)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDataNav.IsEth"
    
End Function
