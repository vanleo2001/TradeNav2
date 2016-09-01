Attribute VB_Name = "mDmDll"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mDmDll.bas
'' Description: Functions for referencing the data manager and symbol universe
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/10/2009   DAJ         Added PutOptionSnap and changed GetOptionSnap
'' 03/10/2009   DAJ         Allowed for time zone in future option market table
'' 04/13/2009   DAJ         Added support for price threshold and secondary min moves
'' 08/26/2009   DAJ         Fixed trading time properties for stock options
'' 02/05/2010   DAJ         New functionality for new stock option symbols
'' 02/10/2010   DAJ         Fixed the description for option symbols
'' 07/21/2010   DAJ         Don't assume that external symbol is a stock option if has spaces
'' 08/05/2010   DAJ         Handle base symbol column in option chain table
'' 03/16/2011   DAJ         Set DefaultStartTime and DefaultEndTime properties for options
'' 07/11/2011   DAJ         Added the LiveContracts function
'' 05/13/2013   DAJ         Allow special users to add enablements based on a flag file
'' 05/14/2013   DAJ         Fix for allowing special users to add enablements
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Global Const kPurchaseCode = 0
Global Const kPurchaseType = 1
Global Const kPurchased = 2
Global Const kExpireDate = 3
Global Const kDate1 = 4
Global Const kDate2 = 5

Public Enum eDM_Status
    eDMStatus_Init = 0       'initializing
    eDMStatus_Checking = 100 'checking distribution files
    eDMStatus_Process = 200  'processing files
    eDMStatus_Clean = 300    'cleaning up
    eDMStatus_Done = 400     'finished
    eDMStatus_Err = 500      'error occurred during processing
    eDMStatus_Cancel = 600   'cancel
End Enum

Public Type DM_Status
    eCurStatus As eDM_Status
    nNumSymbols As Long
    nNumFiles As Long
End Type

Public Type DataMgrStruct
    DM_Handle As Long
    C4Error As Long
    DMError As Long
    DM_DateStamp As Long
    DM_TimeStamp As Long
    reserved1 As Long
    DM_ThreadSafe As Byte
    DM_OpenReadOnly As Byte
    DM_IgnoreReadOnly As Byte
    DM_LastDownload As Byte
    DM_IsOptionSymbol As Byte
    DM_ReplaceSnapshot As Byte
    DM_Unsplit As Byte
    DM_ActiveDates As Byte
    DM_CondenseTicks As Byte
    DM_RemoveBadTicks As Byte '<- remove bad ticks (using the info stored in the bad ticks table)
    DM_BadTickLevel As Byte '<- level for removing auto-scrubbed ticks (0=remove no ticks, 5=standard, 10=remove even slightest offenders)
    DM_TimeFlag As Byte ' ???
    DM_DistContins As Byte ' ???
    DM_GetBidAsk As Byte ' ???
    DM_AdjustDivs As Byte '<- adjust data for dividends (True = also adjust for dividends when split-adjusted)
    DM_Bool3 As Byte
    DM_Long2 As Long
End Type

Public Type SymbolUniverseStruct
    SU_Handle As Long
    C4Error As Long
    SUError As Long
    SU_ThreadSafe As Byte
    SU_OpenReadOnly As Byte
    SU_OpenExclusive As Byte
    SU_Reserved2 As Byte
End Type

Private Type SymbolInfo
    SymbolID As Long
    hstrSecurityType As Long
    hstrSymbol As Long
    hstrBase As Long
    hstrExchange As Long
    hstrDescription As Long
    Access As Byte
    reserved As Byte
    Flags As Integer
End Type

Public Type vbSymbolInfo
    SymbolID As Long
    SecurityType As String
    Symbol As String
    Base As String
    Exchange As String
    Description As String
    Access As Boolean
    reserved As Byte
    Flags As Integer
End Type

Public Type TradeTimeInfo ' all times in minutes since midnight
    iLocalTradeStart As Integer
    iLocalTradeEnd As Integer
    iLocalCrossover As Integer
    iLocalSessionSuspend As Integer
    iFeedTradeStart As Integer
    iFeedTradeEnd As Integer
    iFeedCrossover As Integer
    iFeedSessionSuspend As Integer
    nCurDate As Long
    nNextChange As Long ' date at which next change to feed times occurs, 0 = none
    hstrTimeZone As Long ' gdString
    cFeedTime As String * 1 ' N=NY time, G=GMT time, L=Local time
    cReserved1 As String * 3
    iLocalToGmtOffset As Integer ' # minutes to add to local time to get GMT time
    iLocalSessionResume As Integer
    iFeedSessionResume As Integer
    iReserved2 As Integer
    nReserved(10) As Long
End Type

Public Enum ePurchaseType
    ePurchaseType_DataUpdating = 1
    ePurchaseType_DataHistory = 2
    ePurchaseType_ModulePurchased = 3
    ePurchaseType_ModuleLeased = 4
End Enum

Public Enum ePurchased
    ePurchased_Purchased = 1
    ePurchased_Evaluation = 2
    ePurchased_MoneyBackGuarantee = 3
    ePurchased_Subscription = 4
End Enum


Declare Function DM_ActiveCodeBase Lib "DmDll.dll" (DMS As DataMgrStruct) As Long

Declare Sub DM_Construct Lib "DmDll.dll" (DMS As DataMgrStruct)
Declare Function DM_Open Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hString&, ByVal c4Ptr&) As Byte
Declare Function DM_OpenRW Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hString&, ByVal c4Ptr&) As Byte
Declare Function DM_Close Lib "DmDll.dll" (DMS As DataMgrStruct) As Byte
Declare Function DM_Setup Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hMasterPath&, ByVal hCdPath&, ByVal hDescFile&, ByVal c4Ptr&) As Byte
Declare Function DM_Setup2 Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hMasterPath&, ByVal hCdPath&, ByVal hDescFile&, ByVal hAuthString&, ByVal c4Ptr&, ByVal hMessageWindow&) As Long
Declare Function DM_Setup3 Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hMasterPath&, ByVal hCdPath&, ByVal hDescFile&, ByVal hAuthString&, ByVal c4Ptr&, ByVal hMessageWindow&, ByVal hMainWindow&) As Long

Declare Function DM_SetupStatus Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStatusHandle&, ByVal hStatus&, DMStatusStruct As DM_Status) As Byte

' to load Bars with symbol info
Declare Function DM_LoadSymbolInfo Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal hBars&) As Byte

' to get daily data
Declare Function DM_GetDataEOD Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal nFirstDate&, ByVal nLastDate&, ByVal hBars&) As Byte
' to edit daily data
Declare Function DM_ChangeEOD Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal hBars&, ByVal hChangedSymbolIDs&) As Byte
' to delete daily data (for this bar)
Declare Function DM_ClearEOD Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal nDeleteDate&, ByVal hChangedSymbolIDs&) As Byte
' to clear all the snapshot data
Declare Function DM_ClearSnapshot Lib "DmDll.dll" (DMS As DataMgrStruct) As Byte

'bool DM_GetTickBars(DATAMGR* dm, DMID symbolID, short minutesPerBar, DateType firstDate, DateType lastDate, gdBars& bars);
Declare Function DM_GetTickBars Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal iMinutesPerBar%, ByVal nFirstDate&, ByVal nLastDate&, ByVal hBars&) As Byte

'bool DM_GetTickData2(DATAMGR* dm, DMID symbolID, DateType firstDate, DateType lastDate, gdBars& tickData, short tickMode = 0);
' tickMode: 2 = Minutized (ignore full ticks), else will use full ticks if exist
'Declare Function DM_GetTickData Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal nFirstDate&, ByVal nLastDate&, ByVal hBars&) As Byte
Declare Function DM_GetTickData2 Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal nFirstDate&, ByVal nLastDate&, ByVal hBars&, ByVal iTickMode As Integer) As Byte

'bool DM_GetBadTickTable(DATAMGR* dm, DMID symbolID, DateType fromDate, DateType toDate, gdArrayTable& badTicks);
Declare Function DM_GetBadTickTable Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal nFirstDate&, ByVal nLastDate&, ByVal hBadTickTable&) As Byte

'bool STD_API     DM_PutBadTickTable(DATAMGR* dm, long symID, const gdArrayTable& changeTable);
Declare Function DM_PutBadTickTable Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal hBadTickTable&) As Byte

Declare Function DM_GetTickRawInfo Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal JDate&, jUpdateDate&, nBufLen&, jNearestDate&) As Byte

Declare Function DM_GetSnapAll Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal hlKindIDs&, ByVal hdValues&, ByVal hlDates&) As Byte
Declare Function DM_GetDataKindName Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal nKindId&, ByVal hName&) As Byte
Declare Function DM_GetDataKindID Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hName&, nKindId&) As Byte
Declare Function DM_GetDataKinds Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hlKindIDs&, ByVal hstrNames&) As Byte
Declare Function DM_GetDataKindDesc Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal nKindId&, ByVal hstrDesc&) As Byte
Declare Function DM_GetDataKindsAll Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal nKindId&, ByVal hstrNames&) As Byte

Declare Function DM_GetDataKindInactive Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal nKindId&, iInactive As Integer) As Byte
Declare Function DM_GetDataKindLifetime Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal nKindId&, iLifetime As Integer) As Byte
Declare Function DM_GetDataKindConvfac Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal nKindId&, iConvFactor As Integer) As Byte

Declare Function DM_GetSnap1Sym Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hstrSymbol&, ByVal lKindID&, dValue#, lActiveDate&) As Byte
Declare Function DM_GetSnap1 Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal lKindID&, dValue#, lActiveDate&) As Byte
Declare Function DM_GetDataHist Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal nSymbolID&, ByVal nKindId&, ByVal lFirstDate&, ByVal lLastDate&, ByVal hDates&, ByVal hdValues&, ByVal bShowActiveDates As Long) As Byte
Declare Function DM_GetSnapSelectedData Lib "DmDll.dll" Alias "DM_GetSnapData" (DMS As DataMgrStruct, ByVal lSymbolID&, ByVal hlKindIDs As Long, ByVal hdValues As Long, ByVal hlDates As Long) As Byte

''Declare Function DM_GetDataKindID Lib "DmDll.dll" (ByVal hDMS&, ByVal hName&, nKindID&) As Byte
''Declare Function DM_GetDataHist Lib "DmDll.dll" (ByVal hDMS&, ByVal nSymbolID&, ByVal nKindID&, ByVal hDates&, ByVal hValues&) As Byte
''Declare Function DM_GetDataCurrent Lib "DmDll.dll" (ByVal hDMS&, ByVal nSymbolID&, ByVal nKindID&, dValue#, lDate&) As Byte

Declare Function DM_GetOptSnap Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal hMonths&, ByVal hYears&, ByVal hStrikes&, ByVal hIsCalls&, ByVal hBars&, ByVal hSymbols&) As Byte
Declare Function DM_GetOptSnap1 Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hstrSymbol&, ByVal hstrSecType&, iMonth%, iYear%, dStrike#, bIsCall As Long, ByVal hBars&, lUnderSymID&) As Byte
Declare Function DM_PutOptSnap Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal hlMonths&, ByVal hlYears&, ByVal hdStrikes&, ByVal hlIsCalls&, ByVal hBars&, ByVal hstrSymbols&) As Byte

' 02/04/2010 DAJ: The following are the new calls for retrieving options with the new symbology...
' Retrieve option chain for a particular underlying stock or index by symbolid...
Declare Function DM_GetOptSnap2 Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal lSymbolID&, ByVal hlDays&, ByVal hlMonths&, ByVal hlYears&, ByVal hdStrikes&, ByVal hlIsCalls&, ByVal hBars&, ByVal hSymbols&) As Byte
' Retrieve option chain for a particular underlying stock or index by symbol...
Declare Function DM_GetOptSnapSym2 Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hstrSymbol&, ByVal hlDay&, ByVal hlMonth&, ByVal hlYear&, ByVal hdStrike&, ByVal hlIsCall&, ByVal hBars&, ByVal hSymbols&) As Byte
' Retrieve data for a single option given the symbol...
Declare Function DM_GetOptSnap1Sym Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hstrSymbol&, ByVal hstrSecType&, iDay As Integer, iMonth As Integer, iYear As Integer, dStrike As Double, lIsCall As Long, ByVal hBars&, lUnderSymID As Long) As Byte

Declare Function DM_Distribute Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&, ByVal hDate&, ByVal bDoRecalc As Long, hStatus&, ByVal hWndMsg&) As Long
Declare Function DM_DistributeLatest Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&, hstats&, ByVal hWndMsg&) As Long
Declare Function DM_DistributeDaily Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&, ByVal nDistribDate&, hStatus&, ByVal hWndMsg&) As Long
Declare Function DM_DistStatus Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStatus&, ByVal hStrStatus&, DMStatusStruct As DM_Status) As Byte
Declare Function DM_DistCancel Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStatus&) As Byte
Declare Sub DM_GetSymbolUniverse Lib "DmDll.dll" (DMS As DataMgrStruct, SU As SymbolUniverseStruct)

'bool DM_LinkStaticDB(DATAMGR* dm, const gdString& dbPath, gdString& securityType, gdString& exchange);
'bool DM_UnlinkStaticDB(DATAMGR* dm, const gdString& dbPath);
'bool DM_LinkOptDB(DATAMGR* dm, const gdString& dbPath, gdString& securityType);
'bool DM_UnlinkOptDB(DATAMGR* dm, const gdString& dbPath);
'bool DM_LinkAssocDB(DATAMGR* dm, const gdString& dbPath, gdString& securityType);
'bool DM_UnlinkAssocDB(DATAMGR* dm, const gdString& dbPath);
'bool DM_LinkTickDB(DATAMGR* dm, const gdString& dbPath, gdString& securityType);
'bool DM_UnlinkTickDB(DATAMGR* dm, const gdString& dbPath);
Declare Function DM_LinkStaticDB Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&, ByVal hstrSecType&, ByVal hstrExchange&) As Byte
Declare Function DM_LinkTickDB Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&, ByVal hstrSecType&) As Byte
Declare Function DM_LinkAssocDB Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&, ByVal hstrSecType&) As Byte
Declare Function DM_LinkOptDB Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&, ByVal hstrSecType&) As Byte
Declare Function DM_UnlinkStaticDB Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&) As Byte
Declare Function DM_UnlinkTickDB Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&) As Byte
Declare Function DM_UnlinkAssocDB Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&) As Byte
Declare Function DM_UnlinkOptDB Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&) As Byte

' for new install method
Declare Function DM_UpdateDBConfig Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrPath&) As Byte
Declare Function DM_GetDBInfo Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hastrSecTypes&, ByVal hStrFileType&, ByVal halFileAccess&, ByVal hastrFilePaths&) As Byte

' GetCalendarSpreadData: Returns difference in daily close between given symbol and next contract out (or +2, +3, etc)
'   symbolData -- contains continuous futures symbol and date range
'   values -- returns the difference between given symbol and further out contract (see numContractsOut)
'   numContractsOut -- specifies the number of contracts out to take difference of
'   spreadFlags -- bit flags can be combined:
'     0 => return currentContract less furtherOutContract (typically negative)
'     1 => return furtherOutContract less currentContract (typically positive)
'     2 => return the values back-adjusted (to smooth out contract rolls)
'   ptrSymbolIDs -- returns the symbolID of "furtherOutContract" for each element of the values (ignore if NULL)
' returns true if successful, false otherwise (non-futures symbol, non-continuous symbol, etc)
Declare Function DM_GetCalendarSpreadData2 Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hBars&, ByVal hResultArray&, ByVal iNumContractsOut%, ByVal dwFlags&, ByVal hSymbolsArray&) As Byte

Declare Function DM_TestFunction Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hStrings&) As Long

'C_CALL bool DM_GetDividendInfo(DATAMGR* dm, DMID symbolID, DateType firstDate, DateType lastDate, gdTable& divInfo, short flags);
Declare Function DM_GetDividendInfo Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal SymbolID&, ByVal lFirstDate&, ByVal lLastDate&, ByVal hDivTable&, ByVal iFlags As Integer) As Long

Declare Function SU_GetSymbolID Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal hstrSymbol&) As Long
Declare Function SU_GetSymbolsAll Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal hlSymbols As Long) As Byte
Declare Function SU_GetSymbolsMore Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal hlSymbols As Long, ByVal hstrIndex As Long, ByVal hastrSecTypes As Long, ByVal hastrSymbols As Long, ByVal hastrDescription As Long, ByVal hlFlags As Long) As Byte
Declare Function SU_GetSymbolInfo Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal SymbolID&, Info As SymbolInfo) As Byte
Declare Function SU_GetCompositeInfo Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal SymbolID As Long, ByVal hstrName&, ByVal hstrDescription&, Divisor As Double, ByVal hlComponents&, ByVal hdWeights&, ByVal hlFlags&, dVolDivisor#, ByVal hdVolWeights&) As Byte
Declare Function SU_SetCompositeInfo Lib "SuDll.dll" (SU As SymbolUniverseStruct, SymbolID As Long, ByVal hstrName&, ByVal hstrDesc&, ByVal Divisor#, ByVal hlComponents&, ByVal hdWeights&, ByVal hlFlags&, ByVal dVolDivisor#, ByVal hdVolWeights&) As Byte
Declare Function SU_GetRollInfo Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal SymbolID As Long, ByVal hlSymbols As Long, ByVal hlDates As Long, ByVal hdDeltas As Long) As Byte
Declare Function SU_GetSplitInfo Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal SymbolID As Long, ByVal hlDates As Long, ByVal hfFactors As Long) As Byte
Declare Function SU_GetSplitInfo2 Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal SymbolID As Long, ByVal hlDates As Long, ByVal hlNewShares As Long, ByVal hlOldShares As Long) As Byte
Declare Function SU_GetMarketContracts Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal hList As Long, ByVal lFlags As Long) As Byte
Declare Function SU_GetContractList Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal lSymbolID As Long, ByVal hList As Long, ByVal lFlags As Long) As Byte
Declare Function SU_GetTickInfo Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal lSymbolID As Long, dTickValue As Double, dTickMove As Double, dMinMoveInTicks As Double) As Byte
Declare Function SU_GetTickInfoFmt Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal lSymbolID As Long, dTickValue As Double, dTickMove As Double, dMinMoveInTicks As Double, ByVal hStrFormat As Long) As Byte

Declare Function SU_Feed2Gen Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal hFeedSymbol&, ByVal hFeedExchange&, ByVal hFeed&, ByVal hSecType&, ByVal hSymbol&, nSymbolID&, dMult#) As Byte
Declare Function SU_Gen2Feed Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal hSymbol&, ByVal nSymbolID&, ByVal hFeed&, ByVal hSecType&, ByVal hFeedSymbol&, ByVal hFeedExchange&, dMult#) As Byte
Declare Function SU_GetTimeInfo Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal nSymbolID&, ByVal nCurDate&, TimeInfo As TradeTimeInfo) As Byte
'GetComponentSymbols -- provide list of symbols that contribute to the given symbol
Declare Function SU_GetComponentSymbols Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal cFeed As Byte, ByVal hSymbol&, ByVal hExchange&, ByVal hComponentSymbols&, ByVal hComponentExchanges&) As Byte

Declare Function DM_GetPurchaseString Lib "DmDll.dll" (ByVal hstrPurchaseString As Long) As Byte
Declare Function DM_LoadAuthorization Lib "DmDll.dll" (DMS As DataMgrStruct, ByVal hstrAuthString As Long) As Byte
Declare Function SU_GetUpdatingString Lib "SuDll.dll" (ByVal hAuthString As Long, nDaysUsed As Long) As Byte
Declare Function SU_GetPurchaseInfo Lib "SuDll.dll" (ByVal hPurchaseTable As Long) As Byte

Declare Function SU_GetGMTDateTime Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal dDateTime#, ByVal hstrTimeZone&, dGMTDateTime#) As Byte
Declare Function SU_GetDateTimeFromGMT Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal dGMTDateTime#, ByVal hstrTimeZone&, dDateTime#) As Byte
Declare Function SU_GetLocalTimeZone Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal hstrLocalTimeZone&) As Byte
Declare Function SU_GetNYTimeZone Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal hstrNYTimeZone&) As Byte
Declare Function SU_GetLocalDateTime Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal dDateTime#, ByVal hstrOtherTimeZone&, dLocalDateTime#) As Byte

'C_CALL bool STD_API SU_GetFOExpDate(SYMUNIV* su, const gdString& symbol, int& ExpDate);
Declare Function SU_GetFOExpDate Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal Symbol As Long, lExpDate As Long) As Byte
'C_CALL bool STD_API SU_GetFOStrike(SYMUNIV* su, const gdString& GenSymbol, double& Strike);
Declare Function SU_GetFOStrike Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal GenSymbol As Long, dStrike As Double) As Byte

Declare Function SU_GetChildren Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal lSymbolID As Long, ByVal hlChildrenIDs As Long, ByVal sFamily As Integer) As Byte
Declare Function SU_GetParent Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal lSymbolID As Long, lParentID As Long, ByVal sFamily As Integer) As Byte

' GetHolidayTable returns available dates, reasons, and times for a symbol's holidays after (and including) a particular date
' - returns false if no dates recorded
Declare Function SU_GetHolidayTable Lib "SuDll.dll" (SU As SymbolUniverseStruct, ByVal lSymbolID As Long, ByVal hHolidayTable As Long, ByVal lFromDate As Long) As Byte
'#define HOLIDAY_CLOSED      1
'#define HOLIDAY_EARLYCLOSE  2
'#define HOLIDAY_LATEOPEN    3
'#define HOLIDAY_UNUSUAL     4


Public Function DM_Init(ByVal bOpen As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim errDM&, errC4&
    Dim strPath$
    Dim gdsDmPath As New cGdArray
    Dim strPurchased As String
    Static bOpened As Boolean
    
    'for our codebase ...
    Cb4Start '(make sure our codebase has been initialized)
    TblFlush '(flush all our stuff)
    
    If bOpen And Not bOpened Then
        ChangePath App.Path 'in case need to load DLL
        
        ' get path for data mgr
        strPath = DataPath
        gdsDmPath.Create eGDARRAY_gdString
        gdsDmPath(0) = strPath
        
        DM_Construct g.DMS
        If FileExist(strPath & "SYMBOLS.DBF") Then
            ' OPEN
            'If DM_OpenRW(g.DMS, gdsDmPath.ArrayHandle, cb4Ptr) <> 0 Then
            If DM_OpenRW(g.DMS, gdsDmPath.ArrayHandle, 0) <> 0 Then
                bOpened = True
            Else
                errDM = g.DMS.DMError
                errC4 = g.DMS.C4Error
            End If
        End If
        
        'open Symbol Universe
        If bOpened Then
            DM_GetSymbolUniverse g.DMS, g.SU
            
            ' Load Authorization (Purchased)
            strPurchased = DM_GetPurchased(True)
            DM_LoadAuth strPurchased
        End If
        
    ElseIf bOpened And Not bOpen Then
        ' close DM
        DM_Close g.DMS
        DM_Construct g.DMS
        bOpened = False
        
        g.SU.SU_Handle = 0
    End If

    DM_Init = bOpened
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_Init", eGDRaiseError_Raise
    
End Function

' To load bars with data (daily bars, intraday bars, or ticks)
' - Symbol: can be either string or numeric SymbolID
' - vPeriodicity: can be either string or numeric periodicity
' - nFirstDate: can be < 0 for intraday bars to specify loading at least that many
' - bAppend: only valid for daily and higher
' - bIncludeSnapshotData: True means to include data after LDD (but only if not using salmon)
'   (but setting bIncludeSnapshotData = 2 will override to use even if salmon -- e.g. for option chain)
Public Function DM_GetBars(Bars As cGdBars, ByVal Symbol As Variant, _
        Optional ByVal vPeriodicity As Variant = 0, _
        Optional ByVal nFirstDate As Long = 0, _
        Optional ByVal nLastDate As Long = 0, _
        Optional ByVal bAppend As Boolean = False, _
        Optional ByVal bAutoSetBarType = True, _
        Optional ByVal bUnsplit As Boolean = False, _
        Optional ByVal bIncludeSnapshotData As Integer = True, _
        Optional ByVal bScrubTicks As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bRC As Byte, nSymbolID&, nPeriodicity&, nEmptyCount&, iTickMode%
    Dim i&, iYear%, iMonth%, iDay%, bIsCall As Boolean, dStrike#, lUnderID&, dAdjust#
    Dim d#, dDate#, nBar&
    Dim s$, strCall$, strPeriod$, strSecType$, strSymbol$
    Dim eBarsType As eBarsArray
    Dim TempBars As cGdBars, OneDay As cGdBars
    Dim Rolls As cGdTable
    
    If bIncludeSnapshotData <> 0 Then
        ' the default for using Snapshot data is to only load it from the DM if salmon is not running
        ' (but bIncludeSnapshotData = 2 will override -- e.g. for option chain and criteria recalc)
        If Abs(bIncludeSnapshotData) = 1 And Not g.RealTime Is Nothing Then
            If g.RealTime.SalmonIsRunning Then
                ' TLB: with salmon, data after the LastDailyDownload should no longer be coming from the DataMgr
                bIncludeSnapshotData = False
            End If
        End If
        If bIncludeSnapshotData <> 0 Then
            bIncludeSnapshotData = True
        End If
    End If
    
    ' Changed from nSymbolID to Symbol because we may be passing in an option which does not
    ' have a Symbol ID (01/11/2005 DAJ)...
    nSymbolID = GetSymbolID(Symbol)
    SetBarProperties Bars, Symbol
    strSymbol = Bars.Prop(eBARS_Symbol)
    
    If VarType(vPeriodicity) <> vbString Then
        nPeriodicity = vPeriodicity
    ElseIf Len(vPeriodicity) > 0 Then
        If UCase(Left(vPeriodicity, 1)) = "F" And InStr(UCase(vPeriodicity), "Z") > 0 Then
            ' setup for FractZen bars
            g.FractZen.SetFractZen Bars
            vPeriodicity = Bars.Prop(eBARS_PeriodicityStr)
        End If
        nPeriodicity = GetPeriodicity(vPeriodicity)
    End If
    If nPeriodicity = 0 Then
        nPeriodicity = ePRD_Days + 1
    ElseIf GetPeriodType(nPeriodicity) = ePRD_EachTick Then
        If GetPeriodsPerBar(nPeriodicity) = 2 Then ' special flag to ignore full ticks
            iTickMode = 2 ' (use minutized)
        End If
        nPeriodicity = ePRD_EachTick + 1
    End If
       
    ' appending bars isn't really valid when using the "load at least X # of bars" mode (when nFirstDate < 0)
    If nFirstDate < 0 Or Not bAppend Then
        Bars.Size = 0
        Bars.Prop(eBARS_LastTickTime) = 0
    End If
    Bars.Prop(eBARS_PriceHasSettled) = False
       
    If nFirstDate > 18000000 Then
        nFirstDate = JulFromLong(nFirstDate)
    End If
    If nLastDate <= 0 And nLastDate > -999999 Then
        nLastDate = Date + 1
        Do While Not IsWeekday(nLastDate)
            nLastDate = nLastDate + 1
        Loop
    ElseIf nLastDate > 18000000 Then
        nLastDate = JulFromLong(nLastDate)
    End If
    If nFirstDate > nLastDate Or nLastDate <= -999999 Then
        DM_GetBars = False
        Exit Function ' no data to get
    End If
       
    g.DMS.DM_Unsplit = Abs(bUnsplit)
    g.DMS.DM_AdjustDivs = 0
    If g.bDivAdjust And Not bUnsplit Then
        If Bars.SecurityType <> "F" Then
            g.DMS.DM_AdjustDivs = 1 ' to also adjust for dividends when adjusting for splits
        End If
    End If
    If bIncludeSnapshotData Then
        g.DMS.DM_LastDownload = 0
    Else
        g.DMS.DM_LastDownload = 1
    End If
    
    If bScrubTicks Then
        g.DMS.DM_RemoveBadTicks = 1 '(to remove bad ticks and turn on auto-scrubbing)
    Else
        g.DMS.DM_RemoveBadTicks = 0
    End If
    g.DMS.DM_BadTickLevel = g.iScrubLevel '(to set level for auto-scrubbing)
''g.DMS.DM_BadTickLevel = 0 'Turn off for now
    
    If Len(BarPeriodError(strSymbol, nPeriodicity)) > 0 Then
        Bars.Size = 0
    ElseIf Not IsIntraday(nPeriodicity) Then
        ' Daily, weekly, etc.
        If nFirstDate < 0 Then
            ' determine # days ago to load (provide some buffer)
            i = GetPeriodsPerBar(nPeriodicity) * (Abs(nFirstDate) + 1)
            Select Case GetPeriodType(nPeriodicity)
            Case ePRD_Days
                ' 365 days = about 250 trading days = about 1.5 multiplier
                nFirstDate = LastDailyDownload - Int(i * 1.5 + 2)
            Case ePRD_Weeks
                nFirstDate = LastDailyDownload - i * 7
            Case ePRD_Months
                nFirstDate = LastDailyDownload - i * 31
            Case ePRD_Quarters
                nFirstDate = LastDailyDownload - i * 91
            Case ePRD_Years
                nFirstDate = LastDailyDownload - i * 365
            End Select
        End If
        If nFirstDate < 2 Then nFirstDate = 2
        If bAutoSetBarType Then
            eBarsType = eBARS_Eod
            If nSymbolID = 0 Then
                ' see if a stock or future option
                If VarType(Symbol) = vbString Then
                    If InStr(Symbol, " ") > 0 Then
                        eBarsType = eBARS_Prices Or eBARS_VolOI Or eBARS_BidAsk
                    End If
                End If
            'ElseIf g.SymbolPool.SecType(g.SymbolPool.PoolRecForSymbolID(nSymbolID)) <> eSYMType_Future Then
            ElseIf Bars.SecurityType <> "F" Then
                eBarsType = eBARS_Prices Or eBARS_Vol
            End If
            Bars.ArrayMask = eBarsType
        End If
        Bars.Prop(eBARS_Periodicity) = ePRD_Days + 1
        If nSymbolID <> 0 Then
            bRC = DM_GetDataEOD(g.DMS, nSymbolID, nFirstDate, nLastDate, Bars.BarsHandle)
        ElseIf VarType(Symbol) = vbString Then
            If InStr(Symbol, " ") > 0 Then
                If InStr(Symbol, "-") > 0 Then
                    strSecType = "F"
                Else
                    strSecType = "S"
                End If
                'bRC = DM_GetOptionSnap(Symbol, strSecType, iMonth, iYear, dStrike, bIsCall, Bars, lUnderID)
                bRC = DM_GetOptionSnapNew(Symbol, strSecType, iDay, iMonth, iYear, dStrike, bIsCall, Bars, lUnderID)
                If bRC = 0 Then Bars.Size = 0
                Bars.Prop(eBARS_Symbol) = Symbol
            End If
        End If
        LoadHolidays Bars
        If Bars.Size > 0 Then
#If 1 Then
            ' while we still have daily bars, we can do some massaging
            If nSymbolID <> 0 Then
                If Bars.SecurityType = "M" Then
                    ' TLB 10/11/2013: for mutual fund history, need to delete any bogus prices (if <= 0)
                    For i = Bars.Size - 1 To 0 Step -1
                        If Bars(eBARS_Close, i) <= 0 Then
                            Bars.DeleteSomeBars i
                        End If
                    Next
                ElseIf Bars.SecurityType = "F" Then
                    ' TLB 10/11/2013 (for EWI): for Combined symbols, merge the pit holiday bar with the next daily bar
                    If ConvertFutureSymbol(strSymbol, eCombinedSymbol) = strSymbol Then
                        ' look for each holiday
                        For i = 0 To 99999
                            dDate = Bars.GetHoliday(i)
                            ' TLB 11/11/2015: EWI no longer wants it this way after summer of 2015 (since pits are now closed)
                            If dDate <= 0 Or dDate > 42156 Then
                                Exit For ' no more holidays
                            End If
                            nBar = Bars.FindDateTime(dDate, True)
                            If nBar >= 0 Then
                                ' if bar exists for the holiday, merge it in with the next day
                                d = Bars(eBARS_Open, nBar)
                                If d <> kNullData Then
                                    ' but if it's the last bar, just change the date of the bar
                                    If nBar = Bars.Size - 1 Then
                                        Do
                                            dDate = dDate + 1
                                        Loop While Not IsWeekday(dDate)
                                        Bars(eBARS_DateTime, nBar) = dDate
                                    Else
                                        ' set the open and check the high/low's
                                        Bars(eBARS_Open, nBar + 1) = d
                                        d = Bars(eBARS_High, nBar)
                                        If d > Bars(eBARS_High, nBar + 1) Then
                                            Bars(eBARS_High, nBar + 1) = d
                                        End If
                                        d = Bars(eBARS_Low, nBar)
                                        If d < Bars(eBARS_Low, nBar + 1) Then
                                            Bars(eBARS_Low, nBar + 1) = d
                                        End If
                                        ' sum the volumes
                                        If Bars(eBARS_Vol, nBar) > 0 Then
                                            d = Bars(eBARS_Vol, nBar + 1)
                                            If d > 0 Then
                                                d = d + Bars(eBARS_Vol, nBar)
                                            Else
                                                d = Bars(eBARS_Vol, nBar)
                                            End If
                                            Bars(eBARS_Vol, nBar + 1) = d
                                        End If
                                        If Bars(eBARS_ContVol, nBar) > 0 Then
                                            d = Bars(eBARS_ContVol, nBar + 1)
                                            If d > 0 Then
                                                d = d + Bars(eBARS_ContVol, nBar)
                                            Else
                                                d = Bars(eBARS_ContVol, nBar)
                                            End If
                                            Bars(eBARS_ContVol, nBar + 1) = d
                                        End If
                                        ' after merging, delete the holiday bar
                                        Bars.DeleteSomeBars nBar
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        
            ' now we can turn the daily bars into weekly/monthly/etc
            If nPeriodicity > ePRD_Days + 1 Then
                Bars.BuildBars GetPeriodStr(nPeriodicity)
            End If
#Else
            Bars.BuildBars GetPeriodStr(nPeriodicity)
#End If
        End If
    ElseIf GetPeriodType(nPeriodicity) = ePRD_EachTick Then
        ' tick-by-tick
        If nFirstDate < 2 Then nFirstDate = 2
        If bAutoSetBarType Then
            Bars.ArrayMask = eBARS_TickByTick
        End If
        Bars.Prop(eBARS_Periodicity) = nPeriodicity
        If nSymbolID <> 0 Then
            bRC = DM_GetTickData2(g.DMS, nSymbolID, nFirstDate, nLastDate, Bars.BarsHandle, iTickMode)
            dDate = RoundToSecond(nLastDate + Bars.Prop(eBARS_CrossoverTime) / 1440#)
            If Bars(eBARS_DateTime, Bars.Size - 1) >= dDate Then
                ' temp fix for YC2 issue (was sometimes getting ticks from first part of next day)
                s = "Extra ticks for " & Bars.Prop(eBARS_Symbol) & " on " & DateFormat(nLastDate, MM_DD_YYYY)
                DebugLog "*** " & s
                If IsIDE Then
''                    InfBox s, "e", , "DM_GetTickData2 ERROR!"
                End If
                For i = Bars.Size To 0 Step -1
                    If Bars(eBARS_DateTime, i - 1) < dDate Then
                        Bars.Size = i
                        Exit For
                    End If
                Next
            End If
        ElseIf VarType(Symbol) = vbString Then
            If InStr(Symbol, " ") Then
                Set TempBars = New cGdBars
                Bars.ArrayMask = eBARS_TickByTick Or eBARS_BidAsk
                Bars.Size = 4
                If InStr(Symbol, "-") > 0 Then
                    strSecType = "F"
                Else
                    strSecType = "S"
                End If
                'bRC = DM_GetOptionSnap(Symbol, strSecType, iMonth, iYear, dStrike, bIsCall, TempBars, lUnderID)
                bRC = DM_GetOptionSnapNew(Symbol, strSecType, iDay, iMonth, iYear, dStrike, bIsCall, TempBars, lUnderID)
                If bRC = 0 Then
                    Bars.Size = 0
                Else
                    Bars(eBARS_Close, 0) = TempBars(eBARS_Open, TempBars.Size - 1)
                    Bars(eBARS_Close, 1) = TempBars(eBARS_High, TempBars.Size - 1)
                    Bars(eBARS_Close, 2) = TempBars(eBARS_Low, TempBars.Size - 1)
                    Bars(eBARS_Close, 3) = TempBars(eBARS_Close, TempBars.Size - 1)
                    
                    Bars(eBARS_DateTime, 0) = TempBars(eBARS_DateTime, TempBars.Size - 1) + (TempBars.Prop(eBARS_LastTickTime) / 1440)
                    Bars(eBARS_DateTime, 1) = TempBars(eBARS_DateTime, TempBars.Size - 1) + (TempBars.Prop(eBARS_LastTickTime) / 1440)
                    Bars(eBARS_DateTime, 2) = TempBars(eBARS_DateTime, TempBars.Size - 1) + (TempBars.Prop(eBARS_LastTickTime) / 1440)
                    Bars(eBARS_DateTime, 3) = TempBars(eBARS_DateTime, TempBars.Size - 1) + (TempBars.Prop(eBARS_LastTickTime) / 1440)
                    
                    Bars(eBARS_Bid, 3) = TempBars(eBARS_Bid, TempBars.Size - 1)
                    Bars(eBARS_Ask, 3) = TempBars(eBARS_Ask, TempBars.Size - 1)
                    Bars(eBARS_BidSize, 3) = TempBars(eBARS_BidSize, TempBars.Size - 1)
                    Bars(eBARS_AskSize, 3) = TempBars(eBARS_AskSize, TempBars.Size - 1)
                    
                    Bars(eBARS_Vol, 3) = TempBars(eBARS_Vol, TempBars.Size - 1)
                End If
                Bars.Prop(eBARS_Symbol) = Symbol
            End If
        End If
        LoadHolidays Bars
    Else
        ' intraday
        If bAutoSetBarType Then
            Bars.ArrayMask = eBARS_Intraday
        End If
        Bars.Prop(eBARS_Periodicity) = nPeriodicity
        
If IsIDE Then
    If nSymbolID = 41142 Then '(ES-067)
        'Bars.Prop(eBARS_StartTime) = 1080
    End If
End If
        
        ' make sure it's a valid type of bars (e.g. if needs full tick database)
        If nSymbolID <> 0 Then
            'If 0 Then ' IsIDE And GetPeriodType(nPeriodicity) = ePRD_IntBreakout Then
            'If IsIDE And nPeriodicity = ePRD_IntBreakout + 9876 Then
            If 0 Then
                Bars.Size = 0
                SetBarProperties Bars, nSymbolID
                For dDate = nFirstDate To nLastDate
                    If IsWeekday(dDate) Then
                        ' TLB 12/18/2012: found out we MUST make a new copy every time through the loop
                        ' (otherwise 2nd time through will fail -- one of those "but we don't know why" scenarios)
                        Set TempBars = Bars.MakeCopy(True)
                        TempBars.Size = 0
                        i = g.FractZen.GetFractZenRange(Bars.Prop(eBARS_Symbol), dDate)
                        TempBars.Prop(eBARS_PeriodicityStr) = Str(i) & "b"
                        bRC = DM_GetTickBars(g.DMS, nSymbolID, 0, dDate, dDate, TempBars.BarsHandle)
                        If TempBars.Size > 0 Then
                            nEmptyCount = 0
                            gdAppendBars Bars.BarsHandle, TempBars.BarsHandle, False
                        Else
                            ' check if 15 consecutive days of no intraday data
                            nEmptyCount = nEmptyCount + 1
                            If nEmptyCount > 15 Then Exit For ' must not be any tick data before this
                        End If
                    End If
                Next
            ElseIf IsIDE And GetPeriodType(nPeriodicity) = ePRD_IntRenko Then
                ' testing/debugging
                'nFirstDate = Date - 7
                'nLastDate = Date
                Bars.Size = 0
                SetBarProperties Bars, nSymbolID
                Set TempBars = New cGdBars
                TempBars.ArrayMask = eBARS_TickByTick
                TempBars.Prop(eBARS_Periodicity) = ePRD_EachTick
                iTickMode = 0
                For i = nFirstDate To nLastDate
                    TempBars.Size = 0
                    bRC = DM_GetTickData2(g.DMS, nSymbolID, i, i, TempBars.BarsHandle, iTickMode)
                    If TempBars.Size > 0 Then
                        Bars.BuildBars GetPeriodStr(nPeriodicity), TempBars.BarsHandle, True
                    End If
                Next
                Set TempBars = Nothing
            ElseIf Bars.Prop(eBARS_FractZen) <> 0 Then
                ' for FractZen, go backwards one day at a time (prepending into a Temp set of bars)
                SetBarProperties Bars, nSymbolID
                Set TempBars = Bars.MakeCopy(True)
                TempBars.Size = 0
                For dDate = nLastDate To 0 Step -1
                    If nFirstDate >= 0 And dDate < nFirstDate Then
                        Exit For ' until we are before the nFirstDate
                    ElseIf nFirstDate < 0 And TempBars.Size > Abs(nFirstDate) Then
                        Exit For ' or until we get enough bars
                    ElseIf IsWeekday(dDate) Then
                        ' TLB 12/18/2012: found out we MUST make a new copy every time through the loop
                        ' (otherwise 2nd time through will fail -- one of those "but we don't know why" scenarios)
                        Set OneDay = Bars.MakeCopy(True)
                        OneDay.Size = 0
                        i = g.FractZen.GetFractZenRange(Bars.Prop(eBARS_Symbol), dDate)
                        OneDay.Prop(eBARS_PeriodicityStr) = Str(i) & "b"
                        bRC = DM_GetTickBars(g.DMS, nSymbolID, 0, dDate, dDate, OneDay.BarsHandle)
                        If OneDay.Size > 0 Then
                            ' prepend this day into the TempBars
                            nEmptyCount = 0
                            gdAppendBars TempBars.BarsHandle, OneDay.BarsHandle, True
                        Else
                            ' check if 15 consecutive days of no intraday data
                            nEmptyCount = nEmptyCount + 1
                            If nEmptyCount > 15 Then Exit For ' must not be any tick data before this
                        End If
                    End If
                Next
                ' then append the TempBars onto the real Bars
                gdAppendBars Bars.BarsHandle, TempBars.BarsHandle, False
            ElseIf nFirstDate >= 0 Then
                ' (we now pass 0 for minutes so will use existing periodicity)
                If bAppend And Bars.Size > 0 Then
                    ' TLB 8/29/2014: needed to add this code to do the "append" properly for intraday data
                    Set TempBars = Bars.MakeCopy(True)
                    TempBars.Size = 0
                    bRC = DM_GetTickBars(g.DMS, nSymbolID, 0, nFirstDate, nLastDate, TempBars.BarsHandle)
                    If TempBars.Size > 0 Then
                        ' TLB 1/28/2015: make sure there is no overlap
                        For i = Bars.Size - 1 To 0 Step -1
                            If Bars(eBARS_DateTime, i) < TempBars(eBARS_DateTime, 0) Then
                                Bars.Size = i + 1
                                Exit For
                            End If
                        Next
                        gdAppendBars Bars.BarsHandle, TempBars.BarsHandle, False
                    End If
                Else
                    bRC = DM_GetTickBars(g.DMS, nSymbolID, 0, nFirstDate, nLastDate, Bars.BarsHandle)
                End If
            Else
                ' TLB 2/21/2007: to get at least X number of bars, start at last date and
                ' work backwards 1 week at a time prepending the data until have enough bars
                ' get 1 week's worth of data at a time
                Do While Bars.Size < Abs(nFirstDate)
                    ' TLB 12/18/2012: found out we MUST make a new copy every time through the loop
                    ' (otherwise 2nd time through will fail -- one of those "but we don't know why" scenarios)
                    Set TempBars = Bars.MakeCopy(True)
                    TempBars.Size = 0
                    ' load from just the past Saturday
                    dDate = nLastDate - 7
                    Do While IsWeekday(dDate)
                        dDate = dDate + 1
                    Loop
                    ' (we now pass 0 for minutes so will use existing periodicity)
                    bRC = DM_GetTickBars(g.DMS, nSymbolID, 0, dDate + 1, nLastDate, TempBars.BarsHandle)
                    If TempBars.Size > 0 Then
                        nEmptyCount = 0
                        gdAppendBars Bars.BarsHandle, TempBars.BarsHandle, True
                    Else
                        ' check if 4 consecutive weeks of no intraday data
                        nEmptyCount = nEmptyCount + 1
                        If nEmptyCount > 4 Then Exit Do ' must not be any tick data before this
                    End If
                    ' move back 1 more week
                    nLastDate = dDate
                Loop
            End If
            
            ' TLB 10/3/2005: reset Symbol info since GetTickBars will not use custom settings
            SetBarProperties Bars, nSymbolID
        End If
        LoadHolidays Bars
    End If
    g.DMS.DM_Unsplit = 0
    g.DMS.DM_AdjustDivs = 0
    g.DMS.DM_LastDownload = 0
    
    ' reset Symbol info if need to (since some calls like GetTickBars or GetTickData
    ' can change the symbol, e.g. when doing tick data for a continuous contract)
    If Bars.Prop(eBARS_SymbolID) <> nSymbolID Then
        SetBarProperties Bars, nSymbolID
    End If
    
    Set TempBars = Nothing
    Bars.FreeExtra
    
    FixBarVolumes Bars
    
    ' TLB 7/17/2013: for streaming replay, we need to un-adjust the backadjusted prices in the future
    dAdjust = 0
    If g.nReplaySession > 0 Then
        If Bars.SecurityType = "F" Then
            strSymbol = Bars.Prop(eBARS_Symbol)
            If InStr(strSymbol, "-06") > 0 Then
                ' look at the rolls that happened in the future
                Set Rolls = GetRollsTable(strSymbol)
                For i = Rolls.NumRecords - 1 To 0 Step -1
                    If Rolls.Num(1, i) > g.nReplaySession Then
                        ' and do an un-adjust for them
                        dAdjust = dAdjust - Rolls.Num(2, i)
                    Else
                        Exit For
                    End If
                Next
            End If
        End If
    End If
    
    ' round all prices to nearest price increment
    Bars.FixPrices dAdjust
    
    If Bars(eBARS_Close, 0) = kNullData Then Bars.Prop(eBARS_PriceHasSettled) = False
    
    If bRC = 0 Then
        DM_GetBars = False
    Else
        DM_GetBars = True
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_GetBars", eGDRaiseError_Raise
    
End Function

Public Function DM_GetAllSnapData(ByVal lSymbolID&, _
    alDataIDs As cGdArray, adValues As cGdArray, alDates As cGdArray, _
    Optional astrDataNames As cGdArray = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bSuccess As Boolean
    Dim lIndex As Long
    Dim hArray As Long
    
    Static slSymbolID As Long
    Static sdPrevTime As Double
    Static salDataIDs As cGdArray
    Static sastrDataNames As cGdArray
    Static sadValues As cGdArray
    Static salDates As cGdArray
    
    ' If some time has passed or the symbol id is different,
    ' go get the information again
    If lSymbolID <> slSymbolID Or gdTickCount > sdPrevTime + 1000 Then
        
        ' If this is the first call, create the arrays
        If salDataIDs Is Nothing Then
            Set salDataIDs = New cGdArray
            salDataIDs.Create eGDARRAY_Longs
        End If
        If sastrDataNames Is Nothing Then
            Set sastrDataNames = New cGdArray
            sastrDataNames.Create eGDARRAY_Strings
        End If
        If sadValues Is Nothing Then
            Set sadValues = New cGdArray
            sadValues.Create eGDARRAY_Doubles
        End If
        If salDates Is Nothing Then
            Set salDates = New cGdArray
            salDates.Create eGDARRAY_Longs
        End If
        
        ' Clear out the arrays
        salDataIDs.Size = 0
        sastrDataNames.Size = 0
        sadValues.Size = 0
        salDates.Size = 0
'gdStartProfile 780
        If DM_GetSnapAll(g.DMS, lSymbolID, salDataIDs.ArrayHandle, sadValues.ArrayHandle, salDates.ArrayHandle) <> 0 Then
'gdStopProfile 780
'gdStartProfile 781
            bSuccess = True
            ' round all values to 5 decimals
            hArray = sadValues.ArrayHandle
            For lIndex = 0 To gdGetSize(hArray) - 1
                gdSetNum hArray, lIndex, RoundNum(gdGetNum(hArray, lIndex), 5)
            Next
            slSymbolID = lSymbolID
            sdPrevTime = gdTickCount
'gdStopProfile 781
        End If
'gdStopProfile 780
    Else
        bSuccess = True
    End If
    
    If bSuccess Then
'gdStartProfile 782
        Set alDataIDs = salDataIDs.MakeCopy
        Set adValues = sadValues.MakeCopy
        Set alDates = salDates.MakeCopy
'gdStopProfile 782
        ' names may not be requested
        If Not astrDataNames Is Nothing Then
            ' if names not gotten yet, then must get them
            If sastrDataNames.Size = 0 Then
'gdStartProfile 783
                If Not DM_GetDataKinds(g.DMS, salDataIDs.ArrayHandle, sastrDataNames.ArrayHandle) Then
                    sastrDataNames.Size = 0
                End If
'gdStopProfile 783
            End If
'gdStartProfile 784
            Set astrDataNames = sastrDataNames.MakeCopy
'gdStopProfile 784
        End If
    
    End If

    DM_GetAllSnapData = (sadValues.Size > 0)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_GetAllSnapData", eGDRaiseError_Raise
    
End Function

Public Function DM_GetDataKindDescription(ByVal nKindId&) As String
On Error GoTo ErrSection:
    
    Dim gdsDesc As New cGdArray

    gdsDesc.Create eGDARRAY_gdString
    
    If DM_GetDataKindDesc(g.DMS, nKindId, gdsDesc.ArrayHandle) <> 0 Then
        DM_GetDataKindDescription = gdsDesc(0)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_GetDataKindDescription", eGDRaiseError_Raise
    
End Function

Public Function DM_GetDataKindNameForID(ByVal nKindId&) As String
On Error GoTo ErrSection:

    Dim gdsName As New cGdArray
    
    gdsName.Create eGDARRAY_gdString
    
    If DM_GetDataKindName(g.DMS, nKindId, gdsName.ArrayHandle) <> 0 Then
        DM_GetDataKindNameForID = gdsName(0)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_GetDataKindNameForID", eGDRaiseError_Raise
    
End Function

Public Function DM_GetAllDataKinds(aKindIds As cGdArray, aKindNames As cGdArray) As Boolean
On Error GoTo ErrSection:

    If DM_GetDataKindsAll(g.DMS, aKindIds.ArrayHandle, aKindNames.ArrayHandle) = 0 Then
        DM_GetAllDataKinds = False
    End If
    DM_GetAllDataKinds = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_GetAllDataKinds", eGDRaiseError_Raise
    
End Function

Public Function DM_GetSnapData(ByVal pstrSymbol$, ByVal pstrField$, dValue#, Optional plActiveDate&) As Boolean
On Error GoTo ErrSection:

    Dim lFieldID As Long
    Dim gdsField As New cGdArray
    Dim gdsSymbol As New cGdArray
    
    gdsField.Create eGDARRAY_gdString
    gdsField(0) = Trim(pstrField)
    gdsSymbol.Create eGDARRAY_gdString
    gdsSymbol(0) = Trim(pstrSymbol)
    
    '????If DM_GetDataKindName(g.DMS, lFieldID, gdsField.ArrayHandle) <> 0 Then
    If DM_GetDataKindID(g.DMS, gdsField.ArrayHandle, lFieldID) <> 0 Then
        If DM_GetSnap1Sym(g.DMS, gdsSymbol.ArrayHandle, lFieldID, dValue, plActiveDate) Then
            DM_GetSnapData = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_GetSnapData", eGDRaiseError_Raise
    
End Function

' Loads an Option Chain into a gdTable -- each field is an array:
'  0 = Months (longs)
'  1 = Years (longs)
'  2 = Strikes (doubles)
'  3 = IsCalls (longs)
'  4 = Symbols (strings)
Public Function DM_GetOptChain(ByVal nSymbolID&, tblOptionInfo As cGdTable, Bars As cGdBars) As Boolean
On Error GoTo ErrSection:

    Dim bSuccess As Boolean
    Dim i&, iPos&, strSymbol$
    Dim SymInfo As vbSymbolInfo
    Dim nReturn As Byte
    
    tblOptionInfo.NumRecords = 0&
    Bars.SetBarsHandle gdCreateBars(0&, eBARS_EodBidAsk), True
    
    g.DMS.DM_LastDownload = 0
    
    SU_GetSymbolInf nSymbolID, SymInfo
   
    If tblOptionInfo.NumFields > 7 Then
        nReturn = DM_GetOptSnap2(g.DMS, nSymbolID, tblOptionInfo.FieldArrayHandle(7), _
                tblOptionInfo.FieldArrayHandle(0), tblOptionInfo.FieldArrayHandle(1), _
                tblOptionInfo.FieldArrayHandle(2), tblOptionInfo.FieldArrayHandle(3), _
                Bars.BarsHandle, tblOptionInfo.FieldArrayHandle(4))
    Else
        nReturn = DM_GetOptSnap(g.DMS, nSymbolID, tblOptionInfo.FieldArrayHandle(0), _
                tblOptionInfo.FieldArrayHandle(1), tblOptionInfo.FieldArrayHandle(2), _
                tblOptionInfo.FieldArrayHandle(3), Bars.BarsHandle, _
                tblOptionInfo.FieldArrayHandle(4))
    End If
    
    If nReturn <> 0 Then
        bSuccess = True
        
        With tblOptionInfo
            .NumRecords = .FieldArray(0, False).Size
            
            ' fix symbol names for display purposes
            Select Case SymInfo.SecurityType
                Case "F" ' FUTURE Options: if slash is used, replace with space
                    For i = 0 To .NumRecords - 1
                        strSymbol = .Item(4, i)
                        iPos = InStr(strSymbol, "/")
                        If iPos = 0 Then Exit For '(slashes not used)
                        .Item(4, i) = Left(strSymbol, iPos - 1) & " " & Mid(strSymbol, iPos + 1)
                        
                        If .NumFields > 8 Then
                            .Item(8, i) = Parse(strSymbol, "/", 1)
                        End If
                    Next
            
                Case "S", "I" ' STOCK/INDEX Options: add space if not there already
                    For i = 0 To .NumRecords - 1
                        strSymbol = .Item(4, i)
                        iPos = InStr(strSymbol, " ")
                        
                        ' If we have an old symbol without a space, add one in...
                        If (iPos = 0) Then
                            If (Len(strSymbol) >= 3) And (Len(strSymbol) <= 5) Then
                                .Item(4, i) = Left(strSymbol, Len(strSymbol) - 2) & " " & Right(strSymbol, 2)
                            End If
                        End If
                        
                        ' If we have the "Days" field, but it is null, fill it in with the
                        ' default 3F+1 date...
                        If .NumFields > 7 Then
                            If .Item(7, i) = kNullData Then
                                .Item(7, i) = VBA.Day(GetDateFromRule(.Item(1, i), .Item(0, i), "3F+1"))
                            End If
                        End If
                        
                        If .NumFields > 8 Then
                            .Item(8, i) = Parse(strSymbol, " ", 1)
                        End If
                        
                        ' Temporary code to generate new looking stock option symbols...
'                        If Len(.Item(4, i)) <= 6 Then
'                            If .Item(3, i) = 1 Then
'                                .Item(4, i) = Parse(.Item(4, i), " ", 1) & " " & Str(.Item(1, i)) & Format(.Item(0, i), "00") & Format(.Item(7, i), "00") & " C" & Str(.Item(2, i))
'                            Else
'                                .Item(4, i) = Parse(.Item(4, i), " ", 1) & " " & Str(.Item(1, i)) & Format(.Item(0, i), "00") & Format(.Item(7, i), "00") & " P" & Str(.Item(2, i))
'                            End If
'                        End If
                    Next
            End Select
        End With
    End If

    DM_GetOptChain = bSuccess
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_GetOptChain", eGDRaiseError_Raise
    
End Function

Public Function DM_PutOptChain(ByVal SymbolID&, ByVal alMonths As cGdArray, ByVal alYears As cGdArray, _
        ByVal adStrikes As cGdArray, ByVal alIsCalls As cGdArray, ByVal Bars As cGdBars, _
        ByVal astrSymbols As cGdArray) As Byte
On Error GoTo ErrSection:

    Dim bSuccess As Boolean
    
    g.DMS.DM_LastDownload = 0
    g.DMS.DM_ReplaceSnapshot = 1
    
    If DM_PutOptSnap(g.DMS, SymbolID, alMonths.ArrayHandle, alYears.ArrayHandle, _
            adStrikes.ArrayHandle, alIsCalls.ArrayHandle, Bars.BarsHandle, astrSymbols.ArrayHandle) <> 0 Then
        bSuccess = True
    End If

    DM_PutOptChain = bSuccess
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_PutOptChain", eGDRaiseError_Raise
    
End Function

Public Function DM_PutOptionSnap(ByVal lUnderlyingSymbolID&, ByVal strOptionSymbol As String, ByVal lOptionMonth As Long, ByVal lOptionYear As Long, ByVal dOptionStrike As Double, ByVal lOptionIsCall As Long, ByVal OptionBars As cGdBars) As Byte
On Error GoTo ErrSection:

    Dim bSuccess As Boolean             ' Return of the call to the data manager
    Dim alMonths As New cGdArray        ' Array of contract months
    Dim alYears As New cGdArray         ' Array of contract years
    Dim adStrikes As New cGdArray       ' Array of strike prices
    Dim alIsCalls As New cGdArray       ' Array of Put/Call information
    Dim astrSymbols As New cGdArray     ' Array of symbols
    
    g.DMS.DM_LastDownload = 0
    g.DMS.DM_ReplaceSnapshot = 0
    
    alMonths.Create eGDARRAY_Longs, 1
    alMonths(0) = lOptionMonth
    
    alYears.Create eGDARRAY_Longs, 1
    alYears(0) = lOptionYear
    
    adStrikes.Create eGDARRAY_Doubles, 1
    adStrikes(0) = dOptionStrike
    
    alIsCalls.Create eGDARRAY_Longs, 1
    alIsCalls(0) = lOptionIsCall
    
    astrSymbols.Create eGDARRAY_Strings, 1
    astrSymbols(0) = strOptionSymbol
    
    bSuccess = False
    If DM_PutOptSnap(g.DMS, lUnderlyingSymbolID, alMonths.ArrayHandle, alYears.ArrayHandle, adStrikes.ArrayHandle, alIsCalls.ArrayHandle, OptionBars.BarsHandle, astrSymbols.ArrayHandle) <> 0 Then
        bSuccess = True
    End If

    DM_PutOptionSnap = bSuccess
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_PutOptionSnap"
    
End Function

Public Function DM_DistribData(ByVal strPath$) As Boolean
On Error GoTo ErrSection:

    Dim rc&, bAborted As Boolean
    Dim nDistribDate& '(no longer used)
    Dim DoFinalUpdating As Long
    Dim gdString As New cGdArray
    Dim hStatus As Long
    Dim DMStatusStruct As DM_Status
    
    ' Setup the parms
    If Len(strPath) = 0 Then DoFinalUpdating = 1
    gdString.Create eGDARRAY_gdString
    gdString(0) = Trim(strPath)
    
    ' Start the distribution
    DebugLog "Distributing: " & strPath & " " & Str(nDistribDate) & " " & Str(DoFinalUpdating)
    frmStatus.Status = eStatus_Running
    'rc = DM_Distribute(g.DMS, gdString.ArrayHandle, nDistribDate, DoFinalUpdating, hStatus, frmStatus.txtHwnd.hWnd)
    rc = DM_Distribute(g.DMS, gdString.ArrayHandle, nDistribDate, DoFinalUpdating, hStatus, 0&)  ' frmStatus.txtHwnd.hWnd)
    If rc = 0 Then
        ' Wait for it to finish
        Do While frmStatus.Status < eStatus_Aborted
            If frmStatus.Status = eStatus_Aborting And Not bAborted Then
                DM_DistCancel g.DMS, hStatus
                bAborted = True
            End If
            ' get status
            gdString(0) = ""
            If DM_DistStatus(g.DMS, hStatus, gdString.ArrayHandle, DMStatusStruct) Then
                frmStatus.ProcessStatusMsg gdString(0)
            End If
            Sleep 0.5
        Loop
        If frmStatus.Status = eStatus_Completed Then
            DM_DistribData = True
        End If
        ' if last distribution in sequence, need to sync the caches
        If DoFinalUpdating <> 0 Then
            SyncCodebaseCaches
        End If
    End If
    DebugLog "Distribution finished:  rc=" & Str(rc) & "  hStatus=" & Str(hStatus)
    
ErrExit:
    Exit Function
    
ErrSection:
    If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    RaiseError "mDMDll.DM_DistribData", eGDRaiseError_Raise
    
End Function

Public Function SU_GetAllSymbolsInfo(lSymbolIDs As cGdArray, strIndex As String, astrSecType As cGdArray, astrSymbols As cGdArray, astrDescription As cGdArray, alFlags As cGdArray) As Boolean
On Error GoTo ErrSection:

    Dim hstrIndex As Long
    
    hstrIndex = gdCreateArray(eGDARRAY_gdString)
    lSymbolIDs.Create eGDARRAY_Longs
    astrSecType.Create eGDARRAY_Strings
    astrSymbols.Create eGDARRAY_Strings
    astrDescription.Create eGDARRAY_Strings
    alFlags.Create eGDARRAY_Longs
    
    gdSetStr hstrIndex, 0, strIndex
    
    If SU_GetSymbolsMore(g.SU, lSymbolIDs.ArrayHandle, hstrIndex, astrSecType.ArrayHandle, astrSymbols.ArrayHandle, astrDescription.ArrayHandle, alFlags.ArrayHandle) <> 0 Then
        SU_GetAllSymbolsInfo = True
    Else
        SU_GetAllSymbolsInfo = False
    End If
    
    gdDestroyArray hstrIndex
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.SU_GetAllSymbolsInfo", eGDRaiseError_Raise
    
End Function

Public Function SU_GetAllSymbolIDs(lSymbolIDs As cGdArray) As Boolean
On Error GoTo ErrSection:

    lSymbolIDs.Create eGDARRAY_Longs
    If SU_GetSymbolsAll(g.SU, lSymbolIDs.ArrayHandle) <> 0 Then
        SU_GetAllSymbolIDs = True
    Else
        SU_GetAllSymbolIDs = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.SU_GetAllSymbolIDs", eGDRaiseError_Raise
    
End Function

Public Function SU_GetSymbolInf(ByVal nSymbolID&, Info As vbSymbolInfo) As Boolean
On Error GoTo ErrSection:

    Dim Inf As SymbolInfo
    
    Inf.hstrBase = gdCreateArray(eGDARRAY_gdString)
    Inf.hstrDescription = gdCreateArray(eGDARRAY_gdString)
    Inf.hstrExchange = gdCreateArray(eGDARRAY_gdString)
    Inf.hstrSecurityType = gdCreateArray(eGDARRAY_gdString)
    Inf.hstrSymbol = gdCreateArray(eGDARRAY_gdString)
    
    If SU_GetSymbolInfo(g.SU, nSymbolID, Inf) <> 0 Then
        SU_GetSymbolInf = True
        Info.Base = gdGetStr(Inf.hstrBase)
        Info.Description = gdGetStr(Inf.hstrDescription)
        Info.Exchange = gdGetStr(Inf.hstrExchange)
        Info.SecurityType = gdGetStr(Inf.hstrSecurityType)
        Info.Symbol = gdGetStr(Inf.hstrSymbol)
        Info.Access = Inf.Access
        Info.Flags = Inf.Flags
        Info.reserved = Inf.reserved
        Info.SymbolID = Inf.SymbolID
    Else
        SU_GetSymbolInf = False
    End If
      
ErrExit:
    gdDestroyArray Inf.hstrBase
    gdDestroyArray Inf.hstrDescription
    gdDestroyArray Inf.hstrExchange
    gdDestroyArray Inf.hstrSecurityType
    gdDestroyArray Inf.hstrSymbol
    Exit Function
    
ErrSection:
    gdDestroyArray Inf.hstrBase
    gdDestroyArray Inf.hstrDescription
    gdDestroyArray Inf.hstrExchange
    gdDestroyArray Inf.hstrSecurityType
    gdDestroyArray Inf.hstrSymbol
    RaiseError "mDMDll.SU_GetSymbolInf", eGDRaiseError_Raise
    
End Function

Public Function SU_GetSymbol(ByVal nSymbolID&) As String
On Error GoTo ErrSection:

    Dim Inf As SymbolInfo
    
    Inf.hstrSymbol = gdCreateArray(eGDARRAY_gdString)
    
    If SU_GetSymbolInfo(g.SU, nSymbolID, Inf) <> 0 Then
        SU_GetSymbol = gdGetStr(Inf.hstrSymbol)
    Else
        SU_GetSymbol = ""
    End If
    
ErrExit:
    gdDestroyArray Inf.hstrSymbol
    Exit Function
    
ErrSection:
    gdDestroyArray Inf.hstrSymbol
    RaiseError "mDMDll.SU_GetSymbol", eGDRaiseError_Raise
    
End Function

Public Function SU_GetCompositeInf(ByVal SymbolID As Long, strName$, strDescription$, dDivisor As Double, lComponents As cGdArray, dWeights As cGdArray, lFlags As cGdArray, dVolDivisor As Double, dVolWeights As cGdArray) As Boolean
On Error GoTo ErrSection:

    Dim hstrName As Long, hstrDesc As Long
    
    lComponents.Create eGDARRAY_Longs
    dWeights.Create eGDARRAY_Doubles
    lFlags.Create eGDARRAY_Longs
    dVolWeights.Create eGDARRAY_Doubles

    If SymbolID <> 0 Then
        hstrName = gdCreateArray(eGDARRAY_gdString)
        hstrDesc = gdCreateArray(eGDARRAY_gdString)
        
        If SU_GetCompositeInfo(g.SU, SymbolID, hstrName, hstrDesc, dDivisor, lComponents.ArrayHandle, dWeights.ArrayHandle, lFlags.ArrayHandle, dVolDivisor, dVolWeights.ArrayHandle) <> 0 Then
            SU_GetCompositeInf = True
            strName = gdGetStr(hstrName)
            strDescription = gdGetStr(hstrDesc)
        End If
        
        gdDestroyArray hstrName
        gdDestroyArray hstrDesc
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.SU_GetCompositeInf", eGDRaiseError_Raise
    
End Function

Public Function SU_SetCompositeInf(SymbolID As Long, ByVal strName$, ByVal strDescription$, ByVal dDivisor As Double, lComponents As cGdArray, dWeights As cGdArray, lFlags As cGdArray, ByVal dVolDivisor#, dVolWeights As cGdArray) As Boolean
On Error GoTo ErrSection:

    Dim hstrName As Long, hstrDesc As Long
    
    hstrName = gdCreateArray(eGDARRAY_gdString)
    hstrDesc = gdCreateArray(eGDARRAY_gdString)
    gdSetStr hstrName, 0, strName
    gdSetStr hstrDesc, 0, strDescription
    
    If SU_SetCompositeInfo(g.SU, SymbolID, hstrName, hstrDesc, dDivisor, lComponents.ArrayHandle, dWeights.ArrayHandle, lFlags.ArrayHandle, dVolDivisor, dVolWeights.ArrayHandle) <> 0 Then
        SU_SetCompositeInf = True
    Else
        SU_SetCompositeInf = False
    End If
    
    SyncCodebaseCaches
    
ErrExit:
    gdDestroyArray hstrName
    gdDestroyArray hstrDesc
    Exit Function
    
ErrSection:
    gdDestroyArray hstrName
    gdDestroyArray hstrDesc
    RaiseError "mDMDll.SU_SetCompositeInf", eGDRaiseError_Raise
    
End Function

' See if there is newer data on the last CD used for installing
' (compared to what the current data was installed from)
Public Function NewerCdDataExists() As Boolean
On Error GoTo ErrSection:

    Dim dCD#, dCurr#, strInf$, strFile$, strKey$

If NewFullTick Then Exit Function

    ' check for flag file (for Genesis install)
    strFile = "c:\GenesisInstall.flg"
    If FileExist(strFile) Then
        ' get Starter.GZP location from the flag file
        strInf = FileToString(strFile, , True)
        KillFile strFile
        strFile = Trim(strInf)
        If FileExist(strFile) Then
            dCD = FileDate(strFile)
            If dCD > 0 Then
                ' and replace the registry settings (as if installed from specified location)
                strKey = "Software\Genesis Financial Data Services\Data\Starter.GZP"
                SetRegistryValue rkLocalMachine, strKey, "Path", strFile
                strInf = Format(dCD, "yyyy/mm/dd HH:NN:SS")
                strInf = Replace(strInf, "/", "\")
                SetRegistryValue rkLocalMachine, strKey, "Date", strInf
            End If
        End If
    End If

    ' get date/time of Starter.GZP on the CD used for last installation
    strInf = GetCDDataInf
    dCD = Val(Parse(strInf, vbTab, 2))
    
    ' get date/time of Starter.GZP that current data was from
    strInf = FileToString(DataPath & "Starter.INF", , True)
    dCurr = Val(Parse(strInf, vbTab, 2))
    
    ' see if newer data on the CD (allow for a rounding fudge factor)
    If dCD > dCurr + 1# / 1440# Then
        NewerCdDataExists = True
    Else
        NewerCdDataExists = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.NewerCDDataExists", eGDRaiseError_Raise
    
End Function

' String:  PathOfStarterGZPonCD + Tab + DateTimeOfStarterGZP
Public Function GetCDDataInf() As String
On Error GoTo ErrSection:

    Dim d#, Sec#, strDate$, strTime$, strKey$, strFileName$
    
    ' get path and name of Starter.GZP on the CD
    strKey = "Software\Genesis Financial Data Services\Data\Starter.GZP"
    strFileName = GetRegistryValue(rkLocalMachine, strKey, "Path", "")
    
    ' get date/time of it
    strDate = GetRegistryValue(rkLocalMachine, strKey, "Date", "")
    If Len(Trim(strDate)) > 0 Then
        strTime = Parse(strDate, " ", 2)
        strDate = Parse(strDate, " ", 1)
        ' get date
        d = Val(Parse(strDate, "\", 1)) * 10000
        d = d + Val(Parse(strDate, "\", 2)) * 100
        d = d + Val(Parse(strDate, "\", 3))
        d = JulFromLong(d)
        ' get # seconds from midnight
        Sec = Val(Parse(strTime, ":", 1)) * 60 * 60
        Sec = Sec + Val(Parse(strTime, ":", 2)) * 60
        Sec = Sec + Val(Parse(strTime, ":", 3))
        ' combine
        d = d + Sec / (60# * 60 * 24)
    End If
    
    GetCDDataInf = Trim(strFileName) & vbTab & Str(d)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.GetCDDataInf", eGDRaiseError_Raise
    
End Function

Public Function DM_GetOptionSnap(ByVal strOptionSymbol As String, ByVal strUnderlyingSecType As String, iOptionMonth As Integer, iOptionYear As Integer, _
                                dOptionStrike As Double, bOptionIsCall As Boolean, OptionBars As cGdBars, lUnderlyingSymbolID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lSymbol&, lSecType&
    Dim lIsCall As Long
    
    lSymbol = gdCreateArray(eGDARRAY_gdString)
    lSecType = gdCreateArray(eGDARRAY_gdString)
    gdSetStr lSymbol, 0, strOptionSymbol
    gdSetStr lSecType, 0, strUnderlyingSecType
    
    If DM_GetOptSnap1(g.DMS, lSymbol, lSecType, iOptionMonth, iOptionYear, dOptionStrike, lIsCall, OptionBars.BarsHandle, lUnderlyingSymbolID) <> 0 Then
        bOptionIsCall = (lIsCall <> 0)
        DM_GetOptionSnap = True
    End If
    
ErrExit:
    gdDestroyArray lSymbol
    gdDestroyArray lSecType
    Exit Function
    
ErrSection:
    gdDestroyArray lSymbol
    gdDestroyArray lSecType
    RaiseError "mDMDll.DM_GetOptionSnap", eGDRaiseError_Raise
    
End Function

Public Function DM_GetOptionSnapNew(ByVal strOptionSymbol As String, ByVal strUnderlyingSecType As String, iOptionDay As Integer, iOptionMonth As Integer, iOptionYear As Integer, _
                                dOptionStrike As Double, bOptionIsCall As Boolean, OptionBars As cGdBars, lUnderlyingSymbolID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lSymbol&, lSecType&
    Dim lIsCall As Long
    Dim bReturn As Boolean              ' Return value for the function
    
    lSymbol = gdCreateArray(eGDARRAY_gdString)
    lSecType = gdCreateArray(eGDARRAY_gdString)
    gdSetStr lSymbol, 0, strOptionSymbol
    gdSetStr lSecType, 0, strUnderlyingSecType
    
    bReturn = False
    If DM_GetOptSnap1Sym(g.DMS, lSymbol, lSecType, iOptionDay, iOptionMonth, iOptionYear, dOptionStrike, lIsCall, OptionBars.BarsHandle, lUnderlyingSymbolID) <> 0 Then
        bOptionIsCall = (lIsCall <> 0)
        bReturn = True
    End If
    
    DM_GetOptionSnapNew = bReturn
    
ErrExit:
    gdDestroyArray lSymbol
    gdDestroyArray lSecType
    Exit Function
    
ErrSection:
    gdDestroyArray lSymbol
    gdDestroyArray lSecType
    RaiseError "mDMDll.DM_GetOptionSnapNew", eGDRaiseError_Raise
    
End Function

' Gets current "authorization string" from the registry
Public Sub GetAuthorizationStringFromRegistry()
On Error GoTo ErrSection:

    Dim hAuth&, nDays&, strAuth$, iPos&
    Dim strAllowEnable As String
    Dim strExtra As String
    
    ' get from registry
    hAuth = gdCreateString(0)
    If SU_GetUpdatingString(hAuth, nDays) Then
        If nDays >= 0 And nDays < 32 Then
            strAuth = UCase(StripStr(gdGetStr(hAuth), " "))
        End If
    End If
    gdDestroyArray hAuth
    
    ' cleanup
    If Len(strAuth) = 0 Then
        ' assign default (stripped-down)
        ''If ExtremeCharts >= 1 Then ' CAN'T CALL THIS HERE -- CAUSES RECURSIVE LOOP!!
        ' (TLB: the above used to cause a recursive loop, but now the flags aren't necessarily set yet)
        If GetSourceCode = "R1U" Or FileDate(App.Path & "\..\Eta\Eta.exe") > DateSerial(2005, 10, 1) Then
            'strAuth = ",DEFAULT,I,S"
            strAuth = ",DEFAULT,I,S,SD" '3/15/2012: to add US stocks by default???
        Else
            'strAuth = ",DEFAULT,I,S,F"
            strAuth = ",DEFAULT,I,F,S,SD" '3/15/2012: to add US stocks by default???
        End If
    End If
    
    If Left(strAuth, 1) <> "," Then strAuth = "," & strAuth
    If Right(strAuth, 1) <> "," Then strAuth = strAuth & ","
    
    ' 05/13/2013 DAJ: If a Genesis user has the Enable.FLG, take the contents of that Enable.FLG file
    ' and append it to the enablements string...
    If FileExist("C:\Common\Files.EXE") Then
        strAllowEnable = DecryptFromHex(GetProvidedProperty("AllowEnable"))
        
        If (IsIDE = True) Or (InStr("," & strAllowEnable & ",", "," & Str(RI_GetLastDataServiceID \ 1000) & ",") <> 0) Then
            If FileExist(AddSlash(App.Path) & "Enable.FLG") Then
                strExtra = FileToString(AddSlash(App.Path) & "Enable.FLG")
                If Len(strExtra) > 0 Then
                    strAuth = strAuth & strExtra & ","
                End If
            End If
        End If
    End If
    
    If g.strAuthorizationString <> strAuth Then
        g.strAuthorizationString = strAuth
        
        ' set TradeNav Level
        If InStr(strAuth, ",PLAT,") > 0 Or InStr(strAuth, ",SNV,") > 0 Then
            g.eTradeNavLevel = eTN6_Platinum
        ElseIf InStr(strAuth, ",PROF,") > 0 Or InStr(strAuth, ",RTGPLAT,") > 0 Then
            g.eTradeNavLevel = eTN5_Professional
        ElseIf InStr(strAuth, ",GOLD,") > 0 Then
            g.eTradeNavLevel = eTN4_Gold
        ElseIf InStr(strAuth, ",STAN,") > 0 Or InStr(strAuth, ",PRO,") > 0 Then
            g.eTradeNavLevel = eTN3_Standard
        ElseIf InStr(strAuth, ",LITE,") > 0 Or InStr(strAuth, ",RTGGOLD,") > 0 Then
            g.eTradeNavLevel = eTN2_Lite
        ElseIf g.lLCD > 0 Then '(don't assume Silver if hasn't connected yet)
            g.eTradeNavLevel = eTN1_Silver
        Else
            g.eTradeNavLevel = eTN0_Unknown
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    gdDestroyArray hAuth
    RaiseError "mDMDll.GetAuthorizationStringFromRegistry", eGDRaiseError_Raise
    
End Sub

Public Function DM_Install(ByVal pstrMasterDir$, ByVal pstrCdRoot$, ByVal pstrDescFile$, _
                ByVal pstrAuthString$, Optional ByVal plMsgWindow& = 0, Optional ByVal plMainWindow& = 0) As Long
On Error GoTo ErrSection:

    Dim lMaster As Long
    Dim lCdRoot As Long
    Dim lDescFile As Long
    Dim lAuthString As Long
    
    If Right(pstrMasterDir, 1) = "\" And Right(pstrMasterDir, 2) <> ":\" Then
        pstrMasterDir = Left(pstrMasterDir, Len(pstrMasterDir) - 1)
    End If
    
    lMaster = gdCreateArray(eGDARRAY_gdString)
    lCdRoot = gdCreateArray(eGDARRAY_gdString)
    lDescFile = gdCreateArray(eGDARRAY_gdString)
    lAuthString = gdCreateArray(eGDARRAY_gdString)
    
    gdSetStr lMaster, 0, pstrMasterDir
    gdSetStr lCdRoot, 0, pstrCdRoot
    gdSetStr lDescFile, 0, pstrDescFile
    gdSetStr lAuthString, 0, pstrAuthString
    
    DM_Init False
    DM_Construct g.DMS
    
    'DM_Install = DM_Setup2(g.DMS, lMaster, lCdRoot, lDescFile, lAuthString, 0, plMsgWindow)
    DM_Install = DM_Setup3(g.DMS, lMaster, lCdRoot, lDescFile, lAuthString, 0, plMsgWindow, plMainWindow)
    
ErrExit:
    gdDestroyArray lMaster
    gdDestroyArray lCdRoot
    gdDestroyArray lDescFile
    gdDestroyArray lAuthString
    Exit Function
    
ErrSection:
    gdDestroyArray lMaster
    gdDestroyArray lCdRoot
    gdDestroyArray lDescFile
    gdDestroyArray lAuthString
    RaiseError "mDMDll.DM_Install", eGDRaiseError_Raise
    
End Function

Public Function DM_InstallStatus(ByVal plStatusHandle&, pstrStatus$, pStatusPtr As DM_Status) As Boolean
On Error GoTo ErrSection:

    Dim hStatus As Long
    Dim rc As Byte
    
    hStatus = gdCreateArray(eGDARRAY_gdString)
    rc = DM_SetupStatus(g.DMS, plStatusHandle, hStatus, pStatusPtr)
    If rc <> 0 Then
        pstrStatus = gdGetStr(hStatus)
        DM_InstallStatus = True
    End If
    
ErrExit:
    gdDestroyArray hStatus
    Exit Function
    
ErrSection:
    gdDestroyArray hStatus
    RaiseError "mDMDll.DM_InstallStatus", eGDRaiseError_Raise
    
End Function

' To sync up the separate codebase caches after data has been
' written, we need to clear and restart the optimizations.
Public Sub SyncCodebaseCaches()
On Error GoTo ErrSection:

    Dim cb&, rc&
    
    TblFlush
    
    ' get cb4 ptr for DataMgr
    cb = DM_ActiveCodeBase(g.DMS)
    
    ' clear everybody's optimization
    If cb <> 0 Then
        rc = code4optSuspend(cb)
    End If
    If cb4Ptr <> cb Then
        rc = code4optSuspend(cb4Ptr)
    End If
    
    ' now start optimization again
    If cb <> 0 Then
        rc = code4optStart(cb)
    End If
    If cb4Ptr <> cb Then
        rc = code4optStart(cb4Ptr)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mDMDll.SyncCodebaseCaches", eGDRaiseError_Raise
    
End Sub

Public Function SU_GetSymID(ByVal strSymbol$) As Long
On Error GoTo ErrSection

    Dim aSymbol As New cGdArray
    
    aSymbol.Create eGDARRAY_gdString
    aSymbol(0) = strSymbol
    SU_GetSymID = SU_GetSymbolID(g.SU, aSymbol.ArrayHandle)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.SU_GetSymID", eGDRaiseError_Raise
    
End Function

Public Function GetSymbolID(ByVal Symbol As Variant) As Long
On Error GoTo ErrSection:

    Dim nSymbolID&, hSymbol&, strPrefix$, i&, dValue#
    
    If VarType(Symbol) <> vbString Then
        nSymbolID = CLng(Symbol)
    ElseIf Len(Symbol) > 0 And Left(Symbol, 1) <> "*" Then
        ' strip off prefix (sector/subsector)
        i = InStr(Symbol, ":")
        If i > 0 Then
            strPrefix = UCase(Trim(Left(Symbol, i)))
            Symbol = Trim(Mid(Symbol, i + 1))
        End If
        
        ' first try the symbol pool (in memory)
        nSymbolID = g.SymbolPool.SymbolIDforSymbol(Symbol)
        If nSymbolID = 0 Then
            ' else try the DataManager (database lookup)
            hSymbol = gdCreateString(12)
            gdSetStr hSymbol, 0, Symbol
            nSymbolID = SU_GetSymbolID(g.SU, hSymbol)
            gdDestroyString hSymbol
        End If
        
        ' get sector/subsector
        If Len(strPrefix) > 0 And nSymbolID <> 0 Then
            Select Case strPrefix
            Case "SECTOR:"
                nSymbolID = GetSectorID(nSymbolID, False)
            Case "SUBSECTOR:"
                nSymbolID = GetSectorID(nSymbolID, True)
            End Select
        End If
    End If
    
    GetSymbolID = nSymbolID

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.GetSymbolID", eGDRaiseError_Raise
    
End Function

Public Function GetSymbol(ByVal SymbolID As Variant) As String
On Error GoTo ErrSection:

    Dim strSymbol As String, nSymbolID&, i&
    
    If VarType(SymbolID) <> vbString Then
        nSymbolID = SymbolID
    ElseIf InStr(SymbolID, ":") = 0 Then
        strSymbol = Str(SymbolID)
        ' strip off the month in parentheses (if exists)
        i = InStr(strSymbol, "(")
        If i > 0 Then
            strSymbol = Trim(Left(strSymbol, i - 1))
        End If
    Else
        ' for sectors or subsectors
        nSymbolID = GetSymbolID(SymbolID)
    End If
    
    If nSymbolID <> 0 Then
        strSymbol = g.SymbolPool.SymbolForID(nSymbolID)
        If Len(strSymbol) = 0 Then
            strSymbol = SU_GetSymbol(nSymbolID)
        End If
    End If
    
    GetSymbol = strSymbol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.GetSymbol", eGDRaiseError_Raise
    
End Function

' to get the Sector or Subsector ID for a symbol
Public Function GetSectorID(ByVal nSymbolID&, ByVal bGetSubsector As Boolean) As Long
On Error GoTo ErrSection:

    Dim dValue#, lDate&, lKindID&, strSymbol$

    ' TLB 10/24/2013: due to an old bug in the data, first verify that this is a U.S. stock (or a sector)
    strSymbol = GetSymbol(nSymbolID)
    If InStr(strSymbol, "@") > 0 Then
        nSymbolID = 0
    ElseIf SecurityType(strSymbol) <> "S" Then
        ' except ok for sectors/subsectors
        If Left(strSymbol, 2) <> "$-" Then
            nSymbolID = 0
        End If
    End If
    
    If nSymbolID > 0 Then
        If bGetSubsector Then
            lKindID = 163
        Else
            lKindID = 162
        End If
        If DM_GetSnap1(g.DMS, nSymbolID, lKindID, dValue, lDate) Then
            GetSectorID = dValue ' SymbolID of the sector or subsector
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.GetSectorID"
    
End Function

Public Function SetBarProperties(Bars As cGdBars, ByVal Symbol As Variant, Optional ByVal bJustDefaults As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID of the symbol to get info for
    Dim bSuccess As Boolean             ' Was the symbol found?
    Dim lMarketID As Long               ' Market ID to get information from gMarkets
    Dim strValue As String
    Dim rs As Recordset                 ' Recordset into the database
    Dim strLookupSymbol As String       ' Symbol to lookup in the overrides table
    Dim OptionSymbol As cOptionSymbol
                  
    Bars.SecurityType = Chr(0)
    lSymbolID = GetSymbolID(Symbol)
    If lSymbolID <> 0 Then
        
        ' if no data yet, more efficient to use cache of Bars Properties
        If Bars.Size = 0 Then
            
        End If
                  
        If Not bSuccess Then
            Bars.Prop(eBARS_TickMove) = 0 ' this is a flag to cause LoadSymbolInfo to refresh
            If DM_LoadSymbolInfo(g.DMS, lSymbolID, Bars.BarsHandle) Then
                bSuccess = True
            End If
        End If
    End If
    Bars.Prop(eBARS_hDataMgr) = g.DMS.DM_Handle
    
    ' TLB: TradeNav uses this property for custom purposes,
    ' so we should clear it whenever changing bar properties
    Bars.Prop(eBARS_CustomString) = ""
        
    If Not bSuccess Then
        Bars.Prop(eBARS_SymbolID) = lSymbolID
        If VarType(Symbol) = vbString Then
            Bars.Prop(eBARS_Symbol) = Symbol
            strLookupSymbol = Symbol
        Else
            Bars.Prop(eBARS_Symbol) = ""
            strLookupSymbol = ""
        End If
        ' clear some properties
        Bars.Prop(eBARS_Desc) = ""
        
        ' DAJ 07/21/2010: Don't assume that an external symbol is an option just because
        ' it has spaces in it (don't want to set defaults here)...
        If (InStr(strLookupSymbol, " ") > 0) And (Left(strLookupSymbol, 1) <> "*") Then
            Set OptionSymbol = New cOptionSymbol
            OptionSymbol.FromGenesis strLookupSymbol
            If OptionSymbol.IsFutureOption Then
                bSuccess = FutureOptionProperties(Bars, Symbol)
            Else
                ' default properties for an option (stocks for now)
                Bars.Prop(eBARS_MinMoveInTicks) = 1
                Bars.Prop(eBARS_TickMove) = 0.01
                Bars.Prop(eBARS_TickValue) = 1
                
                Bars.Prop(eBARS_StartTime) = HHMMtoMinutes("0930")
                Bars.Prop(eBARS_DefaultStartTime) = HHMMtoMinutes("0930")
                Bars.Prop(eBARS_EndTime) = HHMMtoMinutes("1600")
                Bars.Prop(eBARS_DefaultEndTime) = HHMMtoMinutes("1600")
                Bars.Prop(eBARS_CrossoverTime) = HHMMtoMinutes("1630")
                Bars.Prop(eBARS_ExchangeTimeZoneInf) = "NY"
            End If
            Bars.Prop(eBARS_Desc) = OptionSymbol.ToDisplay(1)
        Else
            Bars.Prop(eBARS_MinMoveInTicks) = 0
            Bars.Prop(eBARS_TickMove) = 0
            Bars.Prop(eBARS_TickValue) = 0
        End If
    Else
        ' don't know if we'll need to do this in the future:
        Select Case Bars.SecurityType
            Case "S"
                Bars.Prop(eBARS_MarketSymbol) = "!"
            Case "I"
                Bars.Prop(eBARS_MarketSymbol) = "$"
            Case "M"
                Bars.Prop(eBARS_MarketSymbol) = "~"
            Case Else
                ''Bars.Prop(eBARS_MarketSymbol) = Bars.Prop(eBARS_BaseSymbol)
        End Select
    
        strLookupSymbol = Bars.Prop(eBARS_Symbol)
    End If
    
    If Len(Trim(strLookupSymbol)) > 0 And Not bJustDefaults Then
        ' If not just loading defaults, get overrides out of the database...
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarketInfo] " & _
                    "WHERE [Symbol]='" & strLookupSymbol & "' AND [SymbolID]=" & Bars.Prop(eBARS_SymbolID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            strValue = Trim(NullChk(rs!Value))
            Select Case rs!DataType
                Case MarketType(eMarketType_Desc)
                    If Len(strValue) > 0 Then
                        Bars.Prop(eBARS_Desc) = strValue
                    End If
                Case MarketType(eMarketType_SecurityType)
                    Bars.SecurityType = strValue
                Case MarketType(eMarketType_TickMove)
                    Bars.Prop(eBARS_TickMove) = Val(strValue)
                Case MarketType(eMarketType_TickValue)
                    Bars.Prop(eBARS_TickValue) = Val(strValue)
                Case MarketType(eMarketType_MinMoveInTicks)
                    Bars.Prop(eBARS_MinMoveInTicks) = Val(strValue)
                Case MarketType(eMarketType_Margin)
                    Bars.Prop(eBARS_Margin) = Val(strValue)
            End Select
            
            rs.MoveNext
        Loop
    End If
      
    SetBarProperties = bSuccess

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.SetBarProperties", eGDRaiseError_Raise
    
End Function

Public Function DM_LoadAuth(ByVal strAuthString As String) As Boolean
On Error GoTo ErrSection:

    Dim rc As Byte
    Dim hstrAuthString As Long
    
    hstrAuthString = gdCreateArray(eGDARRAY_gdString)
    gdSetStr hstrAuthString, 0, strAuthString

    If DM_LoadAuthorization(g.DMS, hstrAuthString) <> 0 Then
        DM_LoadAuth = True
    End If
    
ErrExit:
    gdDestroyArray hstrAuthString
    Exit Function
    
ErrSection:
    gdDestroyArray hstrAuthString
    RaiseError "mDMDll.DM_LoadAuth", eGDRaiseError_Raise
    
End Function

Public Function DM_GetPurchased(Optional ByVal bUseDefault As Boolean = True) As String
On Error GoTo ErrSection:

    Dim rc As Byte
    Dim hstrAuthString As Long
    
    hstrAuthString = gdCreateArray(eGDARRAY_gdString)
    
    If DM_GetPurchaseString(hstrAuthString) <> 0 Then
        DM_GetPurchased = gdGetStr(hstrAuthString)
    End If
    If Len(DM_GetPurchased) = 0 And bUseDefault = True Then
        'If FileExist(App.Path & "\Install.flg") Then
        '    DM_GetPurchased = FileToString(App.Path & "\Install.flg")
        'Else
            DM_GetPurchased = "*F:0-0,*S:0-0,*I:0-0,*SO:0-0,*IO:0-0"
        'End If
    End If
    
ErrExit:
    gdDestroyArray hstrAuthString
    Exit Function
    
ErrSection:
    gdDestroyArray hstrAuthString
    RaiseError "mDMDll.DM_GetPurchased", eGDRaiseError_Raise
    
End Function

Public Function GetRollsTable(ByVal Symbol As Variant) As cGdTable
On Error GoTo ErrSection:

    Dim nSymbolID&
    Dim Table As New cGdTable
    
    Table.CreateField eGDARRAY_Longs, 0, "SymbolID"
    Table.CreateField eGDARRAY_Longs, 1, "Date"
    Table.CreateField eGDARRAY_Doubles, 2, "Delta"
    Table.NumRecords = 0
    
    If VarType(Symbol) = vbString Then
        ' rolls are only valid for a continuous contract (so we can quickly rule out the others)
        If InStr(Symbol, "-0") > 0 Then
            nSymbolID = GetSymbolID(Symbol)
        End If
    Else
        nSymbolID = GetSymbolID(Symbol)
    End If
    
    If nSymbolID <> 0 Then
        If SU_GetRollInfo(g.SU, nSymbolID, Table.FieldArrayHandle(0), Table.FieldArrayHandle(1), Table.FieldArrayHandle(2)) Then
            Table.NumRecords = gdGetSize(Table.FieldArrayHandle(0))
        Else
            Table.NumRecords = 0
        End If
    End If

ErrExit:
    Set GetRollsTable = Table
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.GetRollsTable"
End Function

Public Function GetSplitsTable(ByVal Symbol As Variant) As cGdTable
On Error GoTo ErrSection:

    Dim nSymbolID&
    Dim Table As New cGdTable
    
    Table.CreateField eGDARRAY_Longs, 0, "Date"
    Table.CreateField eGDARRAY_Longs, 1, "NewShares"
    Table.CreateField eGDARRAY_Longs, 2, "OldShares"
    Table.NumRecords = 0
    
    nSymbolID = GetSymbolID(Symbol)
    If nSymbolID <> 0 Then
        If SU_GetSplitInfo2(g.SU, nSymbolID, Table.FieldArrayHandle(0), _
                Table.FieldArrayHandle(1), Table.FieldArrayHandle(2)) Then
            Table.NumRecords = gdGetSize(Table.FieldArrayHandle(0))
        Else
            Table.NumRecords = 0
        End If
    End If

ErrExit:
    Set GetSplitsTable = Table
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.GetSplitsTable"
End Function

'// -- DIVIDENDS Info -- consolidated dividend info
'//  GetDividendInfo Info -- return the combined cash or stock dividend information for a given symbol
'//  DivInfo -- gdTable where each record in the table has the following fields
'//      0:  dividend kind, one of the following data kinds:
'//          Cash:   DIV_DIST    Total dividend distribution (typically a sum of the other types for a particular date)
'//                  DIV_CASH    Regular cash distribution (for ex: typical quarterly dividend)
'//                  DIV_CASHEQ  Cash equivalent distribution
'//                  DIV_SPEC    Special cash distribution
'//          Stock:  DIV_STOCK   Stock Dividend (for ex: 0.20 = 20%, 0.05 = 5%, 0.0395837 = weird dividend)
'//                                  (these are like splits, divide price by 1.2, 1.05, or 1.0395837)
'//                  DIV_CONSOL  Stock Consolidation (these actually reduce the amount of stock and increase
'//                                  the price.  For instance, 0.96 would divide price by 0.96,
'//                                  effectively increasing the price by 4%.  Kind of a weird reverse split)
'//      1:  value -- depending on dividend kind, a value for the dividend
'//              cash dividend: amount of cash dividend
'//              stock dividend: amount of stock dividend (0.20 = 20%, 0.05 = 5%, etc)
'//              stock consolidation: ratio of consolidation (between 0.0 and 1.0)
'//      2:  execution date -- date the given dividend is attributed to a stock (and the stock's owner)
'//              this is generally when it affects the stock price
'//      3:  pay dates -- date the dividend is paid (to the owner of record as of the execution date)
'//   Flags: 0 = kDivCashTotal   return only total cash dividends (DIV_DIST)
'//          1 = kDivCashAll     return all cash dividends (DIV_DIST, DIV_CASH, DIV_CASHEQ, DIV_SPEC)
'//          2 = kDivStock       return only stock-related dividends (DIV_STOCK, DIV_CONSOL)
'//          3 = kDivCombine     return stock and total cash dividends (DIV_DIST, DIV_STOCK, DIV_CONSOL)
'//          4 = kDivAll         return all dividend data kinds (all 6 kinds shown above)
'//   Types returned by the kDivCombine flag (for displaying each dividend event):
'//          524 = DIV_DIST (all of the cash-type dividends)
'//          529 = DIV_STOCK (stock dividend, which acts like a split)
'//          530 = DIV_CONSOL (stock consolidation, sort of like a reverse split)
Public Function GetDividendsTable(ByVal Symbol As Variant, ByVal bSplitAdjusted As Boolean, _
            Optional ByVal iFlags As Integer = 0, Optional ByVal nFromDate As Long = 0) As cGdTable
On Error GoTo ErrSection:

    Dim nSymbolID&
    Dim d#, dMult#, iDivRec&, iSplitRec&, nDate&, nSplitDate&, nType&
    Dim Table As New cGdTable
    Dim Splits As cGdTable
    
    Table.NumRecords = 0
    
    If AllowDivAndMF Then
        nSymbolID = GetSymbolID(Symbol)
        If nSymbolID <> 0 Then
            'iFlags = 3 ' combined (cash and stock dividends)
            'iFlags = 0 ' all the cash-type dividends
            If DM_GetDividendInfo(g.DMS, nSymbolID, nFromDate, 0, Table.TableHandle, iFlags) = 0 Then
                Table.NumRecords = 0
            ElseIf bSplitAdjusted And Table.NumRecords > 0 Then
                Set Splits = GetSplitsTable(nSymbolID)
                If Splits.NumRecords > 0 Then
                    iSplitRec = Splits.NumRecords - 1
                    dMult = 1
                    For iDivRec = Table.NumRecords - 1 To 0 Step -1
                        ' calculate the split-adjust factor back to the date of this dividend
                        nDate = Table.Num(2, iDivRec)
                        Do While iSplitRec >= 0 And nDate < Splits.Num(0, iSplitRec)
                            If Splits.Num(1, iSplitRec) > 0 And Splits.Num(2, iSplitRec) > 0 Then
                                dMult = dMult * Splits.Num(1, iSplitRec) / Splits.Num(2, iSplitRec)
                            End If
                            iSplitRec = iSplitRec - 1
                        Loop
                        ' should we only split-adjust for cash-type dividends?
                        'nType = Table.Num(0, iDivRec)
                        If dMult > 0 And dMult <> 1 Then 'And nType = 524 Then
                            Table.Num(1, iDivRec) = Table.Num(1, iDivRec) / dMult
                        End If
                    Next
                End If
            End If
        End If
    End If

ErrExit:
    Set GetDividendsTable = Table
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.GetDividendsTable"
End Function

Public Function TranslateSymbol(strSymbol$, strFeed$, strSecType$, strFeedSymbol$, strFeedExchange$, _
        dMult#, nCrossoverTime&, nGmtOffset&, Optional ByVal bToGenesisSymbol As Boolean = False, _
        Optional aCompSymbols As cGdArray = Nothing, Optional aCompExchanges As cGdArray = Nothing) As Boolean

On Error GoTo ErrSection:

    Dim strContract$, strGenSymbol$
    Dim hSymbol&, hFeed&, hSecType&, hFeedSymbol&, hFeedExchange&
    Dim nSymbolID&, i&
    Dim TimeInfo As TradeTimeInfo
    Dim tRolls As cGdTable
    Dim bSuccess As Boolean
    
    dMult = 1
    nCrossoverTime = 0
    nGmtOffset = 0
    
    hSymbol = gdCreateString(0)
    hFeed = gdCreateString(0)
    hSecType = gdCreateString(0)
    hFeedSymbol = gdCreateString(0)
    hFeedExchange = gdCreateString(0)

    gdSetStr hFeed, 0, UCase(Left(strFeed, 1))
    If bToGenesisSymbol Then
        ' look up Genesis symbol
        gdSetStr hSecType, 0, UCase(strSecType)
        gdSetStr hFeedSymbol, 0, strFeedSymbol
        gdSetStr hFeedExchange, 0, strFeedExchange
        strSymbol = ""
        If SU_Feed2Gen(g.SU, hFeedSymbol, hFeedExchange, hFeed, hSecType, hSymbol, nSymbolID, dMult) Then
            strSymbol = gdGetStr(hSymbol)
            bSuccess = True
        End If
    Else
        ' look up Feed symbol
        If InStr(strSymbol, "-05") > 0 Or InStr(strSymbol, "-06") > 0 _
                Or InStr(strSymbol, "-08") > 0 Or InStr(strSymbol, "-09") > 0 Then
            ' for a continuous contract, look up current contract
            Set tRolls = GetRollsTable(strSymbol)
            If tRolls.NumRecords > 0 Then
                nSymbolID = tRolls.Num(0, tRolls.NumRecords - 1)
                strGenSymbol = UCase(Trim(SU_GetSymbol(nSymbolID)))
            End If
            Set tRolls = Nothing
        End If
        If Len(strGenSymbol) = 0 Then
            strGenSymbol = UCase(Trim(strSymbol))
        End If
        gdSetStr hSymbol, 0, strGenSymbol
        nSymbolID = 0
        strSecType = ""
        strFeedSymbol = ""
        strFeedExchange = ""
        If SU_Gen2Feed(g.SU, hSymbol, nSymbolID, hFeed, hSecType, hFeedSymbol, hFeedExchange, dMult) Then
            bSuccess = True
            strSecType = gdGetStr(hSecType)
            strFeedSymbol = gdGetStr(hFeedSymbol)
            strFeedExchange = gdGetStr(hFeedExchange)
            If nSymbolID = 0 Then
                nSymbolID = SU_GetSymbolID(g.SU, hSymbol)
            End If
            'TLB: now need to do this for stocks since could be foreign
            ''If Left(strSecType, 1) <> "S" Then
                If SU_GetTimeInfo(g.SU, nSymbolID, Date, TimeInfo) Then
                    nCrossoverTime = TimeInfo.iFeedCrossover
                    If TimeInfo.cFeedTime <> "N" Then
                        nGmtOffset = TimeInfo.iLocalToGmtOffset
                    End If
                End If
            ''End If
        End If
    End If
    
    ' also get the feed's composite symbols that make up this symbol (if requested)
    If bSuccess And Not aCompSymbols Is Nothing And Not aCompExchanges Is Nothing Then
        aCompSymbols.Create eGDARRAY_Strings
        aCompExchanges.Create eGDARRAY_Strings
        ' (only for futures and only if valid symbol)
        If strSecType = "F" And Len(strFeedSymbol) > 2 Then
            ' must first strip off contract
            gdSetStr hFeedSymbol, 0, Left(strFeedSymbol, Len(strFeedSymbol) - 2)
            gdSetStr hFeedExchange, 0, strFeedExchange
            If SU_GetComponentSymbols(g.SU, Asc(UCase(strFeed)), hFeedSymbol, hFeedExchange, aCompSymbols.ArrayHandle, aCompExchanges.ArrayHandle) Then
                ' then append contract to each component symbol
                For i = 0 To aCompSymbols.Size - 1
                    strContract = Right(strFeedSymbol, 3)
                    If Not IsDigit(strContract, 2) Then
                        strContract = Mid(strContract, 2)
                    End If
                    If Not bToGenesisSymbol And Mid(aCompSymbols(i), 3, 1) = ":" And IsDigit(aCompSymbols(i), 2) And Len(strContract) < 3 Then
                        strContract = Left(strContract, 1) & Mid(Parse(strGenSymbol, "-", 2), 3, 2)
                    End If
                    aCompSymbols(i) = aCompSymbols(i) & strContract
                Next
            End If
        End If
    End If

    TranslateSymbol = bSuccess

ErrExit:
    gdDestroyString hSymbol
    gdDestroyString hFeed
    gdDestroyString hSecType
    gdDestroyString hFeedSymbol
    gdDestroyString hFeedExchange
    Exit Function
    
ErrSection:
    gdDestroyString hSymbol
    gdDestroyString hFeed
    gdDestroyString hSecType
    gdDestroyString hFeedSymbol
    gdDestroyString hFeedExchange
    RaiseError "mDMDll.TranslateSymbol", eGDRaiseError_Raise
    
End Function

Public Function SU_DeleteComposite(ByVal lSymbolID&, ByVal strSymbol$) As Boolean
On Error GoTo ErrSection:

    Dim alTemp As New cGdArray, adTemp As New cGdArray

    alTemp.Create eGDARRAY_Longs
    adTemp.Create eGDARRAY_Doubles
    alTemp.Clear
    adTemp.Clear
    
    SU_DeleteComposite = SU_SetCompositeInf(lSymbolID, strSymbol, "", 0#, alTemp, adTemp, alTemp, 0#, adTemp)
    
    Set alTemp = Nothing
    Set adTemp = Nothing
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.SU_DeleteComposite", eGDRaiseError_Raise
    
End Function

#If 0 Then
' TLB 12/29/2004: these time zone functions were not working consistently enough
' -- we should now use the "ConvertTimeZone" function in mGenesis

Public Function SU_LocalTimeZone() As String
On Error GoTo ErrSection:

    Dim hTimeZone As Long
    
    hTimeZone = gdCreateString(0)
    
    If SU_GetLocalTimeZone(g.SU, hTimeZone) <> 0 Then
        SU_LocalTimeZone = gdGetStr(hTimeZone)
    End If
    
ErrExit:
    gdDestroyString hTimeZone
    Exit Function
    
ErrSection:
    gdDestroyString hTimeZone
    RaiseError "mDMDll.SU_LocalTimeZone", eGDRaiseError_Raise
    
End Function

Public Function SU_NewYorkTimeZone() As String
On Error GoTo ErrSection:

    Dim hTimeZone As Long
    
    hTimeZone = gdCreateString(0)
    
    If SU_GetNYTimeZone(g.SU, hTimeZone) <> 0 Then
        SU_NewYorkTimeZone = gdGetStr(hTimeZone)
    End If
    
ErrExit:
    gdDestroyString hTimeZone
    Exit Function
    
ErrSection:
    gdDestroyString hTimeZone
    RaiseError "mDMDll.SU_NewYorkTimeZone", eGDRaiseError_Raise
    
End Function

Public Function SU_GetLocalTime(ByVal dDateTime#, ByVal strOtherTimeZone$) As Double
On Error GoTo ErrSection:

    Dim hTimeZone As Long
    Dim dLocalDateTime As Double
    
    hTimeZone = gdCreateString(0)
    gdSetStr hTimeZone, 0, strOtherTimeZone
    
    If SU_GetLocalDateTime(g.SU, dDateTime, hTimeZone, dLocalDateTime) <> 0 Then
        SU_GetLocalTime = dLocalDateTime
    End If
    
ErrExit:
    gdDestroyString hTimeZone
    Exit Function
    
ErrSection:
    gdDestroyString hTimeZone
    RaiseError "mDMDll.SU_GetLocalTime", eGDRaiseError_Raise
    
End Function

Public Function SU_GetGMTTime(ByVal dDateTime#, ByVal strTimeZone$) As Double
On Error GoTo ErrSection:

    Dim hTimeZone As Long
    Dim dGMT As Double
    
    hTimeZone = gdCreateString(0)
    gdSetStr hTimeZone, 0, strTimeZone
    
    If SU_GetGMTDateTime(g.SU, dDateTime, hTimeZone, dGMT) <> 0 Then
        SU_GetGMTTime = dGMT
    End If
    
ErrExit:
    gdDestroyString hTimeZone
    Exit Function
    
ErrSection:
    gdDestroyString hTimeZone
    RaiseError "mDMDll.SU_GetGMTTime", eGDRaiseError_Raise
    
End Function

Public Function SU_GetTimeFromGMT(ByVal dGMTDateTime#, ByVal strTimeZone$) As Double
On Error GoTo ErrSection:

    Dim hTimeZone As Long
    Dim dDateTime As Double
    
    hTimeZone = gdCreateString(0)
    gdSetStr hTimeZone, 0, strTimeZone
    
    If SU_GetDateTimeFromGMT(g.SU, dGMTDateTime, hTimeZone, dDateTime) <> 0 Then
        SU_GetTimeFromGMT = dDateTime
    End If
    
ErrExit:
    gdDestroyString hTimeZone
    Exit Function
    
ErrSection:
    gdDestroyString hTimeZone
    RaiseError "mDMDll.SU_GetTimeFromGMT", eGDRaiseError_Raise
    
End Function
#End If

Public Function DM_GetSnapFromHistory(ByVal lSymbolID&, ByVal strKind$, ByVal lDate&, dValue#) As Boolean
On Error GoTo ErrSection:

    Dim alDates As cGdArray
    Dim adValues As cGdArray
    Dim gdsField As New cGdArray
    Dim lFieldID As Long
    
    Set alDates = New cGdArray
    alDates.Create eGDARRAY_Longs
    Set adValues = New cGdArray
    adValues.Create eGDARRAY_Doubles
    gdsField.Create eGDARRAY_gdString
    gdsField(0) = Trim(strKind)
    
    If DM_GetDataKindID(g.DMS, gdsField.ArrayHandle, lFieldID) <> 0 Then
        If DM_GetDataHist(g.DMS, lSymbolID, lFieldID, lDate - 10, lDate, alDates.ArrayHandle, adValues.ArrayHandle, 1) <> 0 Then
            If adValues.Size > 0 Then
                dValue = adValues(adValues.Size - 1)
                DM_GetSnapFromHistory = True
            End If
        End If
    End If
    
ErrExit:
    Set alDates = Nothing
    Set adValues = Nothing
    Set gdsField = Nothing
    Exit Function
    
ErrSection:
    Set alDates = Nothing
    Set adValues = Nothing
    Set gdsField = Nothing
    RaiseError "mDMDll.DM_GetSnapFromHistory", eGDRaiseError_Raise
    
End Function

Public Function DM_GetHistory(ByVal lSymbolID&, ByVal strKind$, ByVal lStartDate&, ByVal lEndDate&, alDates As cGdArray, adValues As cGdArray) As Boolean
On Error GoTo ErrSection:

    Dim gdsField As cGdArray
    Dim lFieldID As Long
    
    Set gdsField = New cGdArray
    gdsField.Create eGDARRAY_gdString
    gdsField(0) = Trim(strKind)

    If DM_GetDataKindID(g.DMS, gdsField.ArrayHandle, lFieldID) <> 0 Then
        If DM_GetDataHist(g.DMS, lSymbolID, lFieldID, lStartDate, lEndDate, alDates.ArrayHandle, adValues.ArrayHandle, 1) <> 0 Then
            DM_GetHistory = True
        End If
    End If
    
ErrExit:
    Set gdsField = Nothing
    Exit Function
    
ErrSection:
    Set gdsField = Nothing
    RaiseError "mDMDll.DM_GetHistory", eGDRaiseError_Raise
    
End Function

' returns path of Data area (where updating area is)
Public Function DataPath() As String
On Error GoTo ErrSection:

    Static strPath As String

    If Len(strPath) = 0 Then
        ' get path for data mgr
        strPath = FileToString(App.Path & "\DATAPATH.DM", , True)
        If Len(strPath) = 0 Then
            ' default
            strPath = App.Path & "\Data\"
        Else
            ' get data path from file
            If InStr(strPath, ":") = 0 And Left(strPath, 1) <> "\" Then
                strPath = App.Path & "\" & strPath
            End If
            strPath = AddSlash(Trim(strPath))
        End If
    End If
    DataPath = strPath

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DataPath", eGDRaiseError_Raise
    
End Function

Public Function DM_GetKindIDFromName(ByVal strName As String) As Long
On Error GoTo ErrSection:

    Dim hName As Long
    Dim lKindID As Long
    
    hName = gdCreateString(0)
    gdSetStr hName, 0, strName
    
    If DM_GetDataKindID(g.DMS, hName, lKindID) <> 0 Then
        DM_GetKindIDFromName = lKindID
    End If
    
ErrExit:
    gdDestroyArray hName
    Exit Function
    
ErrSection:
    gdDestroyArray hName
    RaiseError "mDMDll.DM_GetKindIDFromName", eGDRaiseError_Raise
    
End Function

Public Sub FixBarVolumes(Bars As cGdBars)
On Error GoTo ErrSection:

    Dim hArray&, nBytes&, nPtr&

    ' clear volume for intraday indices
    If Bars.SecurityType = "I" Then
        If Bars.IsIntraday And (Bars.ArrayMask And eBARS_Vol) Then
            hArray = Bars.ArrayHandle(eBARS_Vol)
            If gdIsConstantValue(hArray) = 0 Then
                Select Case Chr(gdGetType(hArray))
                Case "D"
                    nBytes = 8
                Case "F", "L"
                    nBytes = 4
                Case Else
                    nBytes = 0
                End Select
                nBytes = gdGetSize(hArray) * nBytes
                If nBytes > 0 Then
                    nPtr = gdGetDataPtr(hArray)
                    ZeroMemory ByVal nPtr, nBytes
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mDMDll.FixBarVolumes", eGDRaiseError_Raise
    
End Sub

Public Sub ClearSnapshotData()

    Dim rc&
    rc = DM_ClearSnapshot(g.DMS)

End Sub

Public Function DM_GetBadTicks(ByVal Symbol As Variant, _
        Optional ByVal nFirstDate As Long = 0, _
        Optional ByVal nLastDate As Long = 0) As cGdTable

    Dim nSymbolID&
    Dim BadTicks As New cGdTable

    nSymbolID = GetSymbolID(Symbol)
    If DM_GetBadTickTable(g.DMS, nSymbolID, nFirstDate, nLastDate, BadTicks.TableHandle) = 0 Then
        BadTicks.NumRecords = 0
    End If

    Set DM_GetBadTicks = BadTicks

End Function

Public Function DM_PutBadTicks(ByVal Symbol As Variant, ByVal BadTickTable As cGdTable) As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long
    
    lSymbolID = GetSymbolID(Symbol)
    
    If DM_PutBadTickTable(g.DMS, lSymbolID, BadTickTable.TableHandle) <> 0 Then
        DM_PutBadTicks = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.DM_PutBadTicks", eGDRaiseError_Raise
    
End Function

Public Function SU_GetFutureOptionExp(ByVal Symbol As Variant, lExpDate As Long) As Boolean
On Error GoTo ErrSection:

    Dim hSymbol As Long
    Dim strSymbol As String
    
    strSymbol = GetSymbol(Symbol)
    hSymbol = gdCreateString(0)
    gdSetStr hSymbol, 0, strSymbol
    
    If SU_GetFOExpDate(g.SU, hSymbol, lExpDate) <> 0 Then
        SU_GetFutureOptionExp = True
    End If

ErrExit:
    gdDestroyArray hSymbol
    Exit Function
    
ErrSection:
    gdDestroyArray hSymbol
    RaiseError "mDmDll.SU_GetFutureOptionExp", eGDRaiseError_Raise
    
End Function

Public Function SU_GetFutureOptionStrike(ByVal Symbol As Variant, dStrike As Double) As Boolean
On Error GoTo ErrSection:

    Dim hSymbol As Long
    Dim strSymbol As String
    
    strSymbol = GetSymbol(Symbol)
    hSymbol = gdCreateString(0)
    gdSetStr hSymbol, 0, strSymbol
    
    If SU_GetFOStrike(g.SU, hSymbol, dStrike) <> 0 Then
        SU_GetFutureOptionStrike = True
    End If

ErrExit:
    gdDestroyArray hSymbol
    Exit Function
    
ErrSection:
    gdDestroyArray hSymbol
    RaiseError "mDmDll.SU_GetFutureOptionStrike", eGDRaiseError_Raise
    
End Function

Public Function DoFutOpts() As Boolean
    DoFutOpts = True 'FileExist(AddSlash(App.Path) & "FutOpt.FLG")
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SU_GetMarkets
'' Description: Get the list of market contracts (-0) from the symbol universe
'' Inputs:      Array to store contracts in, Flags
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SU_GetMarkets(astrContracts As cGdArray, Optional ByVal lFlags As Long = 0&) As Boolean
On Error GoTo ErrSection:

    ' Symbol;ID;Description
    If SU_GetMarketContracts(g.SU, astrContracts.ArrayHandle, lFlags) <> 0 Then
        SU_GetMarkets = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.SU_GetMarkets", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SU_GetContracts
'' Description: Get the list of contracts for a base symbol ID
'' Inputs:      Symbol ID, Array to store contracts in, Flags
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SU_GetContracts(ByVal lSymbolID As Long, astrContracts As cGdArray, Optional ByVal lFlags As Long = 0&) As Boolean
On Error GoTo ErrSection:

    ' Symbol;ID
    If SU_GetContractList(g.SU, lSymbolID, astrContracts.ArrayHandle, lFlags) <> 0 Then
        SU_GetContracts = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.SU_GetContracts", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DM_GetSnapSelected
'' Description: Get a selected fundamental information for a symbol
'' Inputs:      Symbol ID, Array of Data Kinds, Array of Values, Array of Dates
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DM_GetSnapSelected(ByVal lSymbolID As Long, alDataKindIDs As cGdArray, adValues As cGdArray, alDates As cGdArray) As Boolean
On Error GoTo ErrSection:

    If DM_GetSnapSelectedData(g.DMS, lSymbolID, alDataKindIDs.ArrayHandle, adValues.ArrayHandle, alDates.ArrayHandle) <> 0 Then
        DM_GetSnapSelected = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.DM_GetSnapSelected", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SU_GetGroupChildren
'' Description: Get the children of a symbol in a given family of sector groups
'' Inputs:      Symbol ID, Array of Children IDs, Family ID
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SU_GetGroupChildren(ByVal vSymbol As Variant, alChildrenIDs As cGdArray, Optional ByVal nFamily As Integer = 0) As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID for the symbol passed in
    
    lSymbolID = GetSymbolID(vSymbol)

    If SU_GetChildren(g.SU, lSymbolID, alChildrenIDs.ArrayHandle, nFamily) <> 0 Then
        SU_GetGroupChildren = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.SU_GetGroupChildren", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SU_GetGroupParent
'' Description: Get the parent of a sector group member for a given family
'' Inputs:      Symbol ID, Family ID
'' Returns:     Parent ID if applicable, Zero otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SU_GetGroupParent(ByVal vSymbol As Variant, Optional ByVal nFamily As Integer = 0) As Long
On Error GoTo ErrSection:

    Dim lParentID As Long               ' Parent ID for the symbol passed in
    Dim lSymbolID As Long               ' Symbol ID for the symbol passed in
    
    lSymbolID = GetSymbolID(vSymbol)
    
    SU_GetGroupParent = -1&
    If SU_GetParent(g.SU, lSymbolID, lParentID, nFamily) <> 0 Then
        SU_GetGroupParent = lParentID
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.SU_GetGroupParent", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SU_GetGroupSiblings
'' Description: Given a Symbol in a sector group, get its siblings in the group
'' Inputs:      Symbol ID, Array of Siblings, Family ID
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SU_GetGroupSiblings(ByVal vSymbol As Variant, alSiblingIDs As cGdArray, Optional nFamily As Integer = 0) As Boolean
On Error GoTo ErrSection:

    Dim lParentID As Long               ' Parent ID for the symbol passed in
    Dim lSymbolID As Long               ' Symbol ID for the symbol passed in
    
    lSymbolID = GetSymbolID(vSymbol)
    lParentID = SU_GetGroupParent(lSymbolID, nFamily)
    SU_GetGroupSiblings = SU_GetGroupChildren(lParentID, alSiblingIDs, nFamily)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.SU_GetGroupSiblings", eGDRaiseError_Raise
End Function

Public Function LoadHolidays(Bars As cGdBars, Optional ByVal nFromDate As Long = 0) As Boolean
On Error GoTo ErrSection:

    Dim hHolidayTable&, nSymbolID&, strSymbol$, strPit$, n&
    
    hHolidayTable = gdGetBarsNumProp(Bars.BarsHandle, 99)
    If hHolidayTable <> 0 Then
        nSymbolID = Bars.Prop(eBARS_SymbolID)
        If nSymbolID <> 0 Then
            ' TLB 10/11/2013: for a "Combined" symbol, load holidays for the Pit instead
            If Bars.SecurityType = "F" Then
                strSymbol = Bars.Prop(eBARS_Symbol)
                If ConvertFutureSymbol(strSymbol, eCombinedSymbol) = strSymbol Then
                    strPit = ConvertFutureSymbol(strSymbol, ePitSymbol)
                    n = GetSymbolID(strPit)
                    If n <> 0 Then
                        nSymbolID = n
                    End If
                End If
            End If
            If SU_GetHolidayTable(g.SU, nSymbolID, hHolidayTable, nFromDate) Then
                LoadHolidays = True
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.LoadHolidays", eGDRaiseError_Raise
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SU_GetTickInfoWithFormat
'' Description: Get the tick information including display format for a symbol
'' Inputs:      Symbol, Tick Value, Tick Move, Min Move In Ticks, Format
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SU_GetTickInfoWithFormat(ByVal vSymbol As Variant, dTickValue As Double, dTickMove As Double, dMinMoveInTicks As Double, strFormat As String) As Boolean
On Error GoTo ErrSection:

    Dim hStrFormat As Long              ' Handle to a string with the display format
    Dim lSymbolID As Long               ' Symbol ID of the symbol (or ID) passed in

    lSymbolID = GetSymbolID(vSymbol)
    hStrFormat = gdCreateArray(eGDARRAY_gdString)
    
    If SU_GetTickInfoFmt(g.SU, lSymbolID, dTickValue, dTickMove, dMinMoveInTicks, hStrFormat) <> 0 Then
        strFormat = gdGetStr(hStrFormat)
        SU_GetTickInfoWithFormat = True
    Else
        SU_GetTickInfoWithFormat = False
    End If

ErrExit:
    gdDestroyArray hStrFormat
    Exit Function
    
ErrSection:
    gdDestroyArray hStrFormat
    RaiseError "mDmDll.SU_GetTickInfoWithFormat", eGDRaiseError_Raise
    
End Function

' Returns symbol for the next contract -- blank if none exists
' e.g. GetNextContract("SP-200612") returns "SP-200703"
Public Function GetNextContract(ByVal strSymbol$) As String
On Error GoTo ErrSection:

    Dim i&, nContract&
    
    ' look for next contract forward
    nContract = Val(Parse(strSymbol, "-", 2))
    If nContract > 190000 Then
        ' look at next 12 months
        For i = 1 To 12
            nContract = nContract + 1
            If nContract Mod 100 = 13 Then
                nContract = nContract + 88
            End If
            strSymbol = Parse(strSymbol, "-", 1) & "-" & Str(nContract)
            If GetSymbolID(strSymbol) > 0 Then
                GetNextContract = strSymbol
                Exit For
            End If
        Next
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.GetNextContract", eGDRaiseError_Raise
End Function

' returns the date of the earliest tick data available for a symbol
Public Function GetFirstTickDate(ByVal Symbol) As Long
On Error GoTo ErrSection:

    Dim nUpdateDate&, nNearestDate&, nBufLen&, nSymbolID&, strSymbol$, nContract&, i&
    Dim Rolls As cGdTable
    
    ' need to see if a continuous contract
    nSymbolID = GetSymbolID(Symbol)
    strSymbol = GetSymbol(nSymbolID)
    nContract = Val(Parse(strSymbol, "-", 2))
    If nContract > 0 And nContract < 1000 Then
        ' for a continuous contract, must check each individual contract from roll file
        Set Rolls = GetRollsTable(nSymbolID)
        For i = 0 To Rolls.NumRecords - 1
            nSymbolID = Rolls(0, i)
            If DM_GetTickRawInfo(g.DMS, nSymbolID, 1, nUpdateDate, nBufLen, nNearestDate) <> 0 Then
                If nNearestDate > 0 Then
                    ' TLB 2/9/09: if earliest date for a continuous contract is before the roll date,
                    ' then must bump it up to the roll date
                    If nNearestDate < Rolls(1, i) Then
                        nNearestDate = Rolls(1, i)
                    End If
                    GetFirstTickDate = nNearestDate
                    Exit For
                End If
            End If
        Next
        Set Rolls = Nothing
    Else
        ' else just make the call
        If DM_GetTickRawInfo(g.DMS, nSymbolID, 1, nUpdateDate, nBufLen, nNearestDate) <> 0 Then
            If nNearestDate > 0 Then
                GetFirstTickDate = nNearestDate
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.GetFirstTickDate"
End Function

' used when installing more data history to link in the new tables
' - returns False if an error during the linking process
Public Function UpdateDBConfig() As Boolean
On Error GoTo ErrSection:

    Dim aString As New cGdArray
    aString.Create eGDARRAY_gdString, 0
    aString(0) = DataPath
    ChangePath App.Path ' TLB: need to do this right before the next call
    If DM_UpdateDBConfig(g.DMS, aString.ArrayHandle) <> 0 Then
        UpdateDBConfig = True
    End If
    Set aString = Nothing
    
    ' close and reopen the Data Manager again to make sure new links are included
    DM_Init False
    DM_Init True
    
    ' reset the IsFullTickDB flag
    IsFullTickDB True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.UpdateDBConfig"
End Function

' returns True if the database in the Data folder is the newer Full-Tick type of data install
' (check the FileTypes string, 1 char per file, to see if there is an "F" type)
Public Function IsFullTickDB(Optional ByVal bReset As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strFileTypes$
    Dim aFileTypes As New cGdArray, aSecTypes As New cGdArray, aAccess As New cGdArray, aPaths As New cGdArray
    Static iIsFullTickDB As Integer
       
    If bReset Or (iIsFullTickDB = 0) Then
        iIsFullTickDB = -1 ' assume false
        If FileDate(App.Path & "\DMDLL.DLL") >= DateSerial(2007, 7, 22) Then
            aSecTypes.Create eGDARRAY_Strings
            aAccess.Create eGDARRAY_Longs
            aPaths.Create eGDARRAY_Strings
            aFileTypes.Create eGDARRAY_gdString, 0
            
            If DM_GetDBInfo(g.DMS, aSecTypes.ArrayHandle, aFileTypes.ArrayHandle, aAccess.ArrayHandle, aPaths.ArrayHandle) <> 0 Then
                strFileTypes = gdGetStr(aFileTypes.ArrayHandle)
                If InStr(UCase(strFileTypes), "F") > 0 Then
                    iIsFullTickDB = 1 ' true
                End If
            End If
            
            Set aFileTypes = Nothing
            Set aSecTypes = Nothing
            Set aAccess = Nothing
            Set aPaths = Nothing
        End If
    End If
    If iIsFullTickDB > 0 Then
        IsFullTickDB = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.IsFullTickDB"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BaseSymbolForSymbol
'' Description: Determine the base symbol for the symbol passed in
'' Inputs:      Symbol
'' Returns:     Base Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BaseSymbolForSymbol(ByVal vSymbolOrSymbolID As Variant) As String
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol for the symbol passed in
    Dim strReturn As String             ' Return value for the function
    
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    strReturn = strSymbol
    
    If (InStr(strSymbol, "$") = 0) Then
        If InStr(strSymbol, "-") <> 0 Then
            strReturn = Parse(strSymbol, "-", 1)
        End If
    End If
    
    BaseSymbolForSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.BaseSymbolForSymbol"
    
End Function

Public Function BarPeriodError(ByVal strSymbol$, ByVal nPeriodicity&) As String
On Error GoTo ErrSection:

    Dim strErr$
    Dim nPeriodType As eBarsPeriodType

    nPeriodType = GetPeriodType(nPeriodicity)
    If App.Major >= 4 Then
        ' Intraday bars other than minute bars requires full ticks
        If nPeriodType > ePRD_EachTick And nPeriodType < ePRD_Days And nPeriodType <> ePRD_Minutes And nPeriodType <> ePRD_SMP Then
            If Not IsFullTickDB Then
                strErr = "This type of Bar Period requires the FULL TICK DATABASE (contact Genesis Sales)"
            ElseIf SecurityType(strSymbol) = "S" Then
                If Not HasModule("SFT") Then
                    strErr = "This type of Bar Period for stocks requires the daily downloading" & vbCrLf & "of 'Stock Full Ticks' (please contact Genesis Sales)"
                End If
            End If
        End If
        
        ' Breakout bars requires Standard level
        If nPeriodType = ePRD_IntBreakout Or nPeriodType = ePRD_EodBreakout Then
            If Not HasLevel(eTN3_Standard, False) Then
                If Not HasModule("ROCK*", True) Then
                    strErr = "This type of Bar Period requires at least the 'STANDARD' version (contact Genesis Sales)"
                End If
            End If
        End If
        
        ' Don't allow vol-per-bar or trades-per-bar for Index or non-PFG Forex symbols
        If Left(strSymbol, 1) = "$" And Right(strSymbol, 4) <> "@PFG" Then
            If nPeriodType = ePRD_IntVol Or nPeriodType = ePRD_EodVol Then
                strErr = "Volume bars are not valid"
            ElseIf nPeriodType = ePRD_Ticks Then
                strErr = "Trades per bar is not valid"
            End If
            If Len(strErr) > 0 Then
                If IsForex(strSymbol) Then
                    strErr = strErr & " for the composite Forex symbols"
                Else
                    strErr = strErr & " for Index symbols"
                End If
                strErr = strErr & vbCrLf & "(since the data does not consist of actual trades)"
            End If
        End If
        
        ' TLB 5/10/2011: don't allow intraday ASX data unless has broker override
        If nPeriodType < ePRD_Days Then
            If Right(strSymbol, 4) = "@ASX" Then
                If Not g.RealTime Is Nothing Then
                    If Not g.RealTime.IsBrokerSymbol(strSymbol) Then
                        strErr = "Intraday data not enabled for ASX symbols"
                    End If
                End If
            End If
        End If
        
        ' TLB 7/5/2011: for spreads, only time-based bar periods are allowed
        If InStr(strSymbol, ";") > 0 And Len(strErr) = 0 Then
            If nPeriodType < ePRD_Days Or nPeriodType > ePRD_Years Then
                If nPeriodType <> ePRD_Minutes Then
                    If nPeriodType = ePRD_IntBreakout And HasModule("ROCK*") Then
                        ' TLB: ok for now, due to backwards-compatibility
                    Else
                        strErr = "Spreads require a time-based bar period"
                    End If
                End If
            End If
        End If
    End If
    
    If Len(strErr) > 0 Then
        BarPeriodError = strErr
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.BarPeriodError"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FutureOptionProperties
'' Description: Get the Bars properties for a future option if possible
'' Inputs:      Bars, Symbol
'' Returns:     True if succeeded, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FutureOptionProperties(Bars As cGdBars, ByVal Symbol As Variant) As Boolean
On Error GoTo ErrSection:

    Static astrInfo As cGdArray         ' Array of the base symbol information
    Dim lPos As Long                    ' Position in the array
    Dim astrFields As cGdArray          ' Array of fields
    Dim strBaseSymbol As String         ' Base symbol
    Dim lAdjustment As Long             ' Time adjustment to perform
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    
    If astrInfo Is Nothing Then
        Set astrInfo = New cGdArray
        astrInfo.Create eGDARRAY_Strings
        astrInfo.FromFile AddSlash(App.Path) & "Provided\FutOptMkt.TXT"
    End If
    
    If astrInfo.Size > 0 Then
        strBaseSymbol = Parse(GetSymbol(Symbol), "-", 1)
        If astrInfo.BinarySearch(strBaseSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
            Set astrFields = New cGdArray
            astrFields.SplitFields astrInfo(lPos)
            
            lAdjustment = CLng(Val(astrFields(9)))
            
            Bars.Prop(eBARS_MinMoveInTicks) = Val(astrFields(3))
            Bars.Prop(eBARS_TickMove) = Val(astrFields(2))
            Bars.Prop(eBARS_TickValue) = Val(astrFields(1))
            Bars.Prop(eBARS_StartTime) = HHMMtoMinutes(astrFields(6), lAdjustment)
            Bars.Prop(eBARS_DefaultStartTime) = HHMMtoMinutes(astrFields(6), lAdjustment)
            Bars.Prop(eBARS_EndTime) = HHMMtoMinutes(astrFields(7), lAdjustment)
            Bars.Prop(eBARS_DefaultEndTime) = HHMMtoMinutes(astrFields(7), lAdjustment)
            Bars.Prop(eBARS_CrossoverTime) = HHMMtoMinutes(astrFields(8), lAdjustment)
            If Len(astrFields(10)) = 0 Then
                Bars.Prop(eBARS_ExchangeTimeZoneInf) = "NY"
            Else
                Bars.Prop(eBARS_ExchangeTimeZoneInf) = astrFields(10)
            End If
            
            Bars.PriceThresholds = astrFields(11)
            Bars.SecondaryMinMoves = astrFields(12)
            
            bReturn = True
        Else
            Bars.Prop(eBARS_MinMoveInTicks) = 1
            Bars.Prop(eBARS_TickMove) = 0.01
            Bars.Prop(eBARS_TickValue) = 1
            Bars.Prop(eBARS_ExchangeTimeZoneInf) = "NY"
            
            Bars.PriceThresholds = ""
            Bars.SecondaryMinMoves = ""
        End If
    End If
    
    FutureOptionProperties = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.FutureOptionProperties"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HHMMtoMinutes
'' Description: Convert an HHMM time to minutes from midnight
'' Inputs:      HHMM, Adjustment
'' Returns:     Minutes from Midnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HHMMtoMinutes(ByVal strHHMM As String, Optional ByVal lAdjustment As Long = 0&) As Long
On Error GoTo ErrSection:

    Dim lHour As Long                   ' Hour part of the time
    Dim lMinutes As Long                ' Minutes part of the time
    Dim lReturn As Long                 ' Return value for the function
    
    If Len(strHHMM) = 4 Then
        lHour = CLng(Val(Left(strHHMM, 2)))
        lMinutes = CLng(Val(Right(strHHMM, 2)))
        
        lReturn = (lHour * 60) + lMinutes + lAdjustment
        If lReturn >= 2400 Then
            lReturn = lReturn - 2400
        End If
    End If
    
    HHMMtoMinutes = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.HHMMtoMinutes"
    
End Function

' When OptNav asks for price data, get the data and send it back
Public Sub GetPriceDataForOptNav(ByVal strMsg$)
On Error GoTo ErrSection:

    Dim i&, nBars&, nFromDate&
    Dim Args As New cGdArray
    Dim Bars As New cGdBars

    'strMsg = ES-067;D;200      -> symbol   ;  D=Daily, W=Weekly ; nnn = total bars from present
    Args.SplitFields strMsg, ";"
    
    nBars = Val(Args(2))
    If nBars = 0 Then nBars = 99999
    nFromDate = Int(Date - nBars * 1.5 - 7)
    If nFromDate < 0 Then nFromDate = 0
    DM_GetBars Bars, Args(0), Args(1), nFromDate
    i = Bars.Size - nBars
    If i < 0 Then i = 0
    Do While i < Bars.Size
        strMsg = strMsg & vbTab & Str(Bars(eBARS_DateTime, i)) & ";" _
            & Str(Round(Bars(eBARS_Open, i), 6)) & ";" & Str(Round(Bars(eBARS_High, i), 6)) & ";" _
            & Str(Round(Bars(eBARS_Low, i), 6)) & ";" & Str(Round(Bars(eBARS_Close, i), 6)) & ";" _
            & Str(Bars(eBARS_Vol, i))
        i = i + 1
    Loop
    FileFromString "c:\test.txt", Replace(strMsg, vbTab, vbCrLf)
    
    SendMessageToOptNav eGDOptNav_PriceData, strMsg, True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mDmDll.GetPriceDataForOptNav"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LiveContracts
'' Description: Build a list of contracts for the base symbol of the given
''              symbol that are live (or would have been live on the given date)
'' Inputs:      Symbol, Date
'' Returns:     List of Contracts
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LiveContracts(ByVal strSymbol As String, Optional ByVal lDate As Long = 999999) As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    Dim astrContracts As cGdArray       ' Contracts for the given symbol
    Dim strMarketSymbol As String       ' Market symbol for the given symbol
    Dim lMarketSymbolID As Long         ' Market symbol ID
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPoolRec As Long                ' Record number for the symbol in the symbol pool
    Dim strFrontMonth As String         ' Front month of the 56 contract for the date
    Dim bFound As Boolean               ' Has the front month been found?
    Dim strContract As String           ' Current contract in the array
    Dim lLDD As Long                    ' Last daily download date
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    If SecurityType(strSymbol, True) = "F" Then
        Set astrContracts = New cGdArray
        astrContracts.Create eGDARRAY_Strings
        
        strMarketSymbol = Parse(strSymbol, "-", 1) & "-0"
        lMarketSymbolID = GetSymbolID(strMarketSymbol)
        
        If SU_GetContracts(lMarketSymbolID, astrContracts) Then
            strFrontMonth = RollSymbolForDate(Parse(strSymbol, "-", 1) & "-056", lDate)
            lLDD = LastDailyDownload
            bFound = False
            
            For lIndex = 0 To astrContracts.Size - 1
                strContract = Parse(astrContracts(lIndex), ";", 1)
                
                If bFound = False Then
                    If strContract = strFrontMonth Then
                        bFound = True
                    End If
                End If
                If bFound = True Then
                    If lDate >= lLDD Then
                        astrReturn.Add strContract
                    Else
                        lPoolRec = g.SymbolPool.PoolRecForSymbol(strContract)
                        If lPoolRec <> 0 Then
                            If g.SymbolPool.LastDate(lPoolRec) >= lDate Then
                                astrReturn.Add strContract
                            End If
                        End If
                    End If
                End If
            Next lIndex
        End If
    Else
        astrReturn.Add strSymbol
    End If
    
    Set LiveContracts = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDmDll.LiveContracts"
    
End Function

' GetCalendarSpreadData: Returns difference in daily close between given symbol and next contract out (or +2, +3, etc)
'   symbolData -- contains continuous futures symbol and date range
'   values -- returns the difference between given symbol and further out contract (see numContractsOut)
'   numContractsOut -- specifies the number of contracts out to take difference of
'   spreadFlags -- bit flags can be combined:
'     0 => return currentContract less furtherOutContract (typically negative)
'     1 => return furtherOutContract less currentContract (typically positive)
'     2 => return the values back-adjusted (to smooth out contract rolls)
'   ptrSymbolIDs -- returns the symbolID of "furtherOutContract" for each element of the values (ignore if NULL)
' returns true if successful, false otherwise (non-futures symbol, non-continuous symbol, etc)
Public Function DM_CalendarSpread(aResults As cGdArray, ByVal Bars As cGdBars, ByVal iNumContractsOut&, _
                    Optional ByVal bBackAdjusted As Boolean = False, Optional aSymbolIds As cGdArray = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim iFlags&, hSymbolIDs&

    ' if contracts out is negative, then set flags to return opposite direction
    If iNumContractsOut < 0 Then
        iFlags = 1
        iNumContractsOut = Abs(iNumContractsOut)
    End If
    If bBackAdjusted Then
        iFlags = iFlags Or 2
    End If
    
    ' get ptr to the SymbolIDs array if the array was passed (correctly)
    If Not aSymbolIds Is Nothing Then
        If aSymbolIds.ArrayType = eGDARRAY_Longs Then
            hSymbolIDs = aSymbolIds.ArrayHandle
        End If
    End If

    'aResults.Create eGDARRAY_Doubles, Bars.Size
    If DM_GetCalendarSpreadData2(g.DMS, Bars.BarsHandle, aResults.ArrayHandle, iNumContractsOut, iFlags, hSymbolIDs) <> 0 Then
        DM_CalendarSpread = True
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDMDll.DM_CalendarSpread", eGDRaiseError_Raise
End Function

