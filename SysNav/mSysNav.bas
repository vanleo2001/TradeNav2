Attribute VB_Name = "mSysNav"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mSysNav.bas
'' Description: Helper functions and declarations for system navigator stuff
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Date         Author      Description
'' 04/17/2009   DAJ         Fix pyramiding information after a library import
'' 03/11/2010   DAJ         Mods to MarketsInExpressions routine (#5580)
'' 05/17/2010   DAJ         Moved routines from frmRule for global use
'' 10/08/2010   DAJ         Added functions to determine Next Bar functions
'' 06/16/2011   DAJ         Added code for the Highlight Bar Reporter
'' 09/30/2011   DAJ         Added code for capturing a report instead of showing it
'' 10/10/2011   DAJ         When doing a RefreshReverify, only load engine functions once
'' 11/15/2011   DAJ         Renamed the strategy basket stuff
'' 12/09/2011   DAJ         Optionally include Next Bar Open in Next Bar References
'' 05/01/2013   DAJ         Shadow Trading
'' 05/14/2013   DAJ         Optionally allow loading a guru basket if not the owner
'' 05/15/2013   DAJ         Load basket regardless of ownership when determining max units
'' 05/24/2013   DAJ         Speed enhancements
'' 07/30/2013   DAJ         Don't allow exact match for enablement for shadow basket
'' 08/05/2013   DAJ         Don't allow user to save strategy basket with an existing name
'' 08/13/2013   DAJ         Fix for looking up if trading item is deleted
'' 09/05/2013   DAJ         Don't bother to check for duplicate basket names if no records in table
'' 10/04/2013   DAJ         Don't load basket items when checking for guru items
'' 12/29/2014   DAJ         Make loading data optional in MarketsInExpressions
'' 03/03/2015   DAJ         Added the StrategyBasketItemIdForKey function
'' 06/02/2015   DAJ         Added the kSN_BASKETLIMIT constant
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_UseAdvForRunExpressions As Boolean

' Columns in the gdTable holding library information
Public Enum etblLibrary
    etblLib_ID = 0
    etblLib_Name
    etblLib_Type
    etblLib_Path
    etblLib_VBProcAddr
    etblLib_LastModified
    etblLib_NumFields
End Enum

' Columns in the gdTable holding function information
Public Enum etblFunction
    etblFunction_ID = 0                 ' Long
    etblFunction_LibID                  ' Long
    etblFunction_Name                   ' String
    etblFunction_NameCoded              ' String
    etblFunction_Implementation         ' Long
    etblFunction_CodedText              ' String
    etblFunction_ReturnType             ' Long
    etblFunction_LateCalculating        ' Tiny
    etblFunction_UsesOpenNextBar        ' Tiny
    etblFunction_UsesHLCNextBar         ' Tiny
    etblFunction_CategoryID             ' Long
    etblFunction_LastModified           ' Double
    etblFunction_Usage                  ' Tiny
    etblFunction_TradeSenseUsage        ' String
    etblFunction_Description            ' String
    etblFunction_SecurityLevel          ' Long
    etblFunction_Password               ' String
    etblFunction_CannotDelete           ' Tiny
    etblFunction_Reverify               ' Tiny
    etblFunction_NumFields
End Enum

' Columns in the gdTable holding function parameter information
Public Enum etblFunctionParm
    etblFunctionParm_ID = 0
    etblFunctionParm_FunctionID
    etblFunctionParm_Sequence
    etblFunctionParm_Name
    etblFunctionParm_Type
    etblFunctionParm_FromValue
    etblFunctionParm_ToValue
    etblFunctionParm_Required
    etblFunctionParm_Default
    etblFunctionParm_NumFields
End Enum

' Columns in the gdTable holding rule information
Public Enum etblRule
    etblRule_RuleID = 0                 ' Long
    etblRule_RuleName                   ' String
    etblRule_RuleType                   ' Long
    etblRule_BuySell                    ' Tiny
    etblRule_LibraryID                  ' Long
    etblRule_LastModified               ' Double
    etblRule_PreviewRTF                 ' String
    etblRule_SecurityLevel              ' Long
    etblRule_Password                   ' String
    etblRule_CannotDelete               ' Tiny
    etblRule_SystemNumber               ' Long
    etblRule_Reverify                   ' Tiny
    etblRule_CategoryID                 ' Long
    etblRule_NumFields
End Enum

Public Enum eTradesHeader
    eTradesHeader_SystemNumber = 0
    eTradesHeader_SystemName
    eTradesHeader_BarTimeFrame
    eTradesHeader_StartDate
    eTradesHeader_EndDate
    eTradesHeader_TotalBars
    eTradesHeader_Expenses
    eTradesHeader_Symbol
    eTradesHeader_TickMove
    eTradesHeader_TickValue
    eTradesHeader_MinMoveInTicks
    eTradesHeader_Margin
    eTradesHeader_SecurityType
    eTradesHeader_SessionStart
    eTradesHeader_SessionEnd
    eTradesHeader_TimeZoneInfo
    eTradesHeader_LongStopLoss
    eTradesHeader_ShortStopLoss
End Enum

Public Enum eMarketTypes
    eMarketType_Desc = 0
    eMarketType_SecurityType
    eMarketType_TickMove
    eMarketType_TickValue
    eMarketType_MinMoveInTicks
    eMarketType_Margin
End Enum

Public Enum eGDEquityFilterMode
    eGDEquityFilterMode_BelowMa = 0
    eGDEquityFilterMode_MaDown
End Enum

Public Enum eGDTakeNextTradeValue
    eGDTakeNextTrade_No = 0
    eGDTakeNextTrade_Yes
    eGDTakeNextTrade_NotEnoughData
    eGDTakeNextTrade_NoEquityFilter
End Enum

Public Enum eGDTransactionField
    eGDTransactionField_Date = 0
    eGDTransactionField_Action
    eGDTransactionField_Quantity
    eGDTransactionField_Price
    eGDTransactionField_Position
    eGDTransactionField_Rule
    eGDTransactionField_AvgEntry
    eGDTransactionField_Link
    eGDTransactionField_Profit
    eGDTransactionField_UnfilteredEquity
    eGDTransactionField_EquityMovAvg
    eGDTransactionField_Skip
    eGDTransactionField_FilteredEquity
End Enum

Global Const gExitSignal = 1
Global Const gEntrySignal = 0
Global Const gUserErr = vbObjectError + 1000

'Parameter Types
Global Const kSN_RetNumericConstant = 1
Global Const kSN_RetText = 2
Global Const kSN_RetTrueFalse = 3
Global Const kSN_RetNumeric = 4
Global Const kSN_RetBars = 5
Global Const kSN_RetTrueFalseConstant = 6
Global Const kSN_RetTrades = 7
Global Const kSN_RetTextSeries = 8

'Implementation types
Global Const kSN_BuiltIn = 1
Global Const kSN_Custom = 2
Global Const kSN_Internal = 3

'Source types
Global Const kSN_MM = 1
Global Const kSN_System = 2
Global Const kSN_Both = 3

'Used to determine the type of rule in tblRules
Global Const kSN_RESERVED_STOP_RULE = 2
Global Const kSN_SYSTEMRULE = 0

'Default user library is always ID
Global Const kSN_UserLibrary = 8

'System Study Modes
Global Const kSN_SystemStudies = "S"

'Constants used in Menu/Shortcut navigation
Global Const kSN_Systems = "kSystems"
Global Const kSN_Rules = "kRules"
Global Const kSN_Functions = "kFunctions"
Global Const kSN_Libraries = "kLibraries"
Global Const kSN_LibraryItems = "kLibraryItems"

'Trade types
Global Const kSN_Navigator = "System Navigator"

'Optimization Status Codes
Global Const kSN_OPTIMIZATION_IN_PROGRESS = 0
Global Const kSN_OPTIMIZATION_COMPLETED = 1
Global Const kSN_OPTIMIZATION_ERROR = 2

Global Const kSN_OPTIMIZATION_RUN_ABORTED = 65557 ' from engine
Global Const kSN_MAX_ITERATION_ERROR = 65587

Global Const kSN_MAX_ITERATIONS = 1000000000

Global Const kSN_MIN_ASPERCENT = 5

Global Const kSN_BASKETLIMIT = 99

'Source types
Global Const C_MM = 1
Global Const C_System = 2
Global Const C_Both = 3

Public Type TradeStruct
    bLong As Boolean
    lEntryDate As Long
    dEntryPrice As Double
    strEntry As String
    lExitDate As Long
    dExitPrice As Double
    strExit As String
    dProfit As Double
End Type

' Mark Jurik library
Private Declare Function JMAUT Lib "JRS_UT.dll" (ByVal dSeries#, ByVal dSmooth#, ByVal dPhase#, pdOutput#, ByVal lDestroy&, piSeriesID&, ByVal lSameBar&) As Long
Private Declare Function RSXUT Lib "JRS_UT.dll" (ByVal dSeries#, ByVal dSmooth#, pdOutput#, ByVal lDestroy&, piSeriesID&, ByVal lSameBar&) As Long
Private Declare Function VELUT Lib "JRS_UT.dll" (ByVal dSeries#, ByVal lDepth&, pdOutput#, ByVal lDestroy&, piSeriesID&, ByVal lSameBar&) As Long
Private Declare Function DMXUT Lib "JRS_UT.dll" (ByVal dHigh#, ByVal dLow#, ByVal dClose#, pdOutBipolar#, pdOutPlus#, pdOutMinus#, ByVal dLength#, ByVal lDestroy&, piSeriesID&, ByVal lSameBar&) As Long

' TAS Indicator library
Private Declare Function TAS_IndicatorInit Lib "TASIndicators.dll" (ByVal strPgm$, ByVal strFunc$, ByVal strSymbol$) As Long
Private Declare Function TAS_IndicatorSetParameter Lib "TASIndicators.dll" (ByVal nIndId&, ByVal nParmID&, ByVal dParmValue#) As Long
Private Declare Function TAS_IndicatorSetBar Lib "TASIndicators.dll" (ByVal nIndId&, ByVal nBar&, ByVal nYYYYMMDD&, ByVal nHHMM&, ByVal dOpen#, ByVal dHigh#, ByVal dLow#, ByVal dClose#, ByVal dVol#, ByVal dOI#) As Long
Private Declare Function TAS_IndicatorSetBarNoCalc Lib "TASIndicators.dll" (ByVal nIndId&, ByVal nBar&, ByVal nYYYYMMDD&, ByVal nHHMM&, ByVal dOpen#, ByVal dHigh#, ByVal dLow#, ByVal dClose#, ByVal dVol#, ByVal dOI#) As Long
Private Declare Function TAS_IndicatorValue Lib "TASIndicators.dll" (ByVal nIndId&, ByVal nReturnID&) As Double


' Pass "RunSysNavEngine" a gdStringArray of parms,
' a callback function, and an optional gdStringArray for timings.
' e.g. i = RunSysNavEngine(gdStrParms, AddressOf EngineCallback, gdStrTimes)
''Private Declare Function RunSysNavEngine Lib "NavEngine.dll" (ByVal hStrParms As Long, ByVal CallbackFunction As Long, ByVal hStrTimes As Long) As Long
Private Declare Function InitSysNavEngine Lib "NavEngine.dll" (ByVal hStrParms As Long, ByVal CallbackFunction As Long, ByVal hStrTimes As Long) As Long
Private Declare Function ExitSysNavEngine Lib "NavEngine.dll" (ByVal hStrParms As Long, ByVal CallbackFunction As Long, ByVal hStrTimes As Long) As Long

Private Declare Function InitSysNavEngineAdv Lib "NavEngineAdv.dll" Alias "InitSysNavEngine" (ByVal hStrParms As Long, ByVal CallbackFunction As Long, ByVal hStrTimes As Long) As Long
Private Declare Function ExitSysNavEngineAdv Lib "NavEngineAdv.dll" Alias "ExitSysNavEngine" (ByVal hStrParms As Long, ByVal CallbackFunction As Long, ByVal hStrTimes As Long) As Long


'StrParms (string array):
'  0:  Name of the Function Set to use
'  1:  Last Bar Good flag - assumed true unless 'false'
'  2:  AlignMarketTraded - if 'true', the results of all expressions
'      are aligned to market one; default is 'true'
'StrBars - string array of parm names of the bars, e.g., Market1
'BarsArray - array of bars handles (same size as StrBars)
'ExpArray - the string array of expressions to execute
'ExpResults - array of results matching the expressions
Public Declare Function ExecuteExpressions Lib "NavEngine.dll" _
    (ByVal hStrParms&, ByVal hStrBars&, ByVal hBarsArray&, _
    ByVal hExpArray&, ByVal hExpResults&, ByVal hMinBarsReq&, ByVal hStrTimes&) As Long

'StrParms (string array):
'  0:  Expression set name (filled in if left empty)
'  1:  Name of the Function Set to use
'  2:  AlignMarketTraded - if 'true', the results of all expressions
'      are aligned to market one; default is 'true'
'StrBars - string array of parm names of the bars, e.g., Market1
'ExpArray - the string array of expressions to execute
Private Declare Function InitExpressionsOLD Lib "NavEngine.dll" Alias "InitExpressions" _
    (ByVal hStrParms&, ByVal hStrBars&, _
    ByVal hExpArray&, ByVal hStrTimes&) As Long
Private Declare Function InitExpressionsNEW Lib "NavEngineAdv.dll" Alias "InitExpressions" _
    (ByVal hStrParms&, ByVal hStrBars&, _
    ByVal hExpArray&, ByVal hStrTimes&) As Long

'StrParms (string array):
'  0:  Expression set name
'  1:  Last Bar Good flag - assumed true unless 'false'
'StrBars - string array of parm names of the bars, e.g., Market1
'BarsArray - array of bars handles (same size as StrBars)
'ExpResults - array of results matching the expressions
Private Declare Function RunExpressionsOLD Lib "NavEngine.dll" Alias "RunExpressions" _
    (ByVal hStrParms&, ByVal hStrBars&, ByVal hBarsArray&, _
     ByVal hExpResults&, ByVal hMinBarsReq&, ByVal hStrTimes&) As Long
Private Declare Function RunExpressionsNEW Lib "NavEngineAdv.dll" Alias "RunExpressions" _
    (ByVal hStrParms&, ByVal hStrBars&, ByVal hBarsArray&, _
     ByVal hExpResults&, ByVal hMinBarsReq&, ByVal hStrTimes&) As Long

'StrParms (string array):
'  0:  Name of Expression Set to clear; clears all expression set
'      if null list item or null string passed
Private Declare Function ClearExpressionsOLD Lib "NavEngine.dll" Alias "ClearExpressions" _
    (ByVal hStrParms&, ByVal hStrTimes&) As Long
Private Declare Function ClearExpressionsNEW Lib "NavEngineAdv.dll" Alias "ClearExpressions" _
    (ByVal hStrParms&, ByVal hStrTimes&) As Long


'//////////////////////////////////////////////////////////////////////////////
' Last Bar Expression Sets (so can recalc only the last bar)
'
' Start a last bar set of expressions
' paParms:
'  0:  Expression set name (filled in if left empty)
'  1:  Function Set name (default is 'Default')
'  2:  AlignMarketTraded - if 'true', the results of all expressions
'      are aligned to market one; default is 'true'
'  3:  Last Bar Good flag - assumed true unless 'false'
' paParmNames - string array of parm names of the bars, e.g., Market1
' paParmList - array of bars
' paExpressions - the expressions to execute
' paResults - array of results matching the expressions
'SYSNAV_API long EXPORTDLL StartLastBarSet(      gdArrayStr*         paParms,
'                                                gdArrayStr*         paParmNames,
'                                                gdParmList*         paParmList,
'                                                gdArrayStr*         paExpressions,
'                                                gdArrayResults*     paResults,
'                                                gdArrayL*           paMinBarsReq = NULL,
'                                                gdArrayStr*         paTimings = NULL);
Public Declare Function StartLastBarSet Lib "NavEngineAdv.dll" _
    (ByVal hStrParms&, ByVal hStrBars&, ByVal hBarsArray&, _
     ByVal hExpArray&, ByVal hExpResults&, ByVal hMinBarsReq&, ByVal hStrDrawingCommands&) As Long
'
' Last bar expression caculations from the given bar list forward
' paParms:
'  0:  Expression set name
'  1:  Last Bar Good flag - assumed true unless 'false'
' paFromBarIndexes - array of bar numbers from which to redo / continue calculations
'SYSNAV_API long EXPORTDLL ContinueLastBarSet(   gdArrayStr*         paParms,
'                                                gdArrayL*           paFromBarIndexs,
'                                                gdArrayResults*     paResults,
'                                                gdArrayL*           paMinBarsReq = NULL,
'                                                gdArrayStr*         paTimings = NULL);
'Public Declare Function ContinueLastBarSet Lib "NavEngineAdv.dll" _
    (ByVal hStrParms&, ByVal hFromBarIndex&, ByVal hExpResults&, ByVal hMinBarsReq&, ByVal hStrTimes&) As Long
Public Declare Function ContinueLastBarSet Lib "NavEngineAdv.dll" _
    (ByVal hStrParms&, ByVal hFromBarIndex&, ByVal hExpResults&, ByVal hMinBarsReq&, ByVal hStrDrawingCommands&) As Long
'
' Clear the bar by bar expression set(s) to end a series of runs
'  paParms:
'   0:  Name of Expression Set to clear; clears all expression set
'       if null list item or null string passed
'SYSNAV_API long EXPORTDLL ClearLastBarSet(      gdArrayStr*         paParms,
'                                                gdArrayStr*         paTimings = NULL);
Public Declare Function ClearLastBarSet Lib "NavEngineAdv.dll" _
    (ByVal hStrParms&, ByVal hStrTimes&) As Long
'//////////////////////////////////////////////////////////////////////////////


' Execute the given system, optimizing within the parameter values given
' or doing a next bar analysis
' The step call back is used to deliver the results of each optimization
' parameter set. The first call back has null results and the step number
' is the total number of runs to test all parameter combinations
' pxTradeStrings is used to store the trades for a given run
' pxRunValueStrings are the parameter values and/or rule alternation flags
' for a given run.
' pxTradeStrings and pxRunValueStrings are passed to the call back function
' at the end of each run. The call back function must not rely on these
' value remaining the same after returning, as these arrays are reused for
' each run.
' pxCnslNextBarStrings and pxRuleNextBarStrings contain the next bar reports (if any)
'
' paParms:
'   0: Function Set Name
'   1: Up/down method (how to assume whether high or low of bar hits first)
'      'Omega' or 'Genesis', if missing or not one of these values, Omega is assumed
'   2: Start Trading Date, yyyymmddhhnnss
'   3: Stop Trading Date,  yyyymmddhhnnss (if nextbar, then the date/time for the
'                                          next bar report)
'   4: Last bar good data - OHLC, that is, O means the open is good, H high good, etc
'   5: Run Option: 'nextbar'   do a next bar analysis (given an existing position,
'                              if supplied, or from start trading date)
'                  'backtest'  backtest only with default parameter values
'                  'optimize'  using the parameter values in the database,
'                  'alternate' the rules with the alternate flags on
'                  'both'      or anything else does parameter optimization
'                              and rule alternation
'                  'parmrun'   run system with given parm values and rule alternation
'                              in pxRunValueStrings
'   6: Max Rules Off (if alternating rules)
'   7: Single Entry Rule ID (textual number); if supplied, only the given entry rule
'      plus the exit rules are loaded and used in the specified execution of the system
'   8: Existing Position: not implemented; ignored
'
Public Declare Function ExecuteSystem Lib "NavEngine.dll" _
    (ByVal hStrParms&, ByVal hTblSystem&, ByVal hTblSystemRule&, ByVal hTblSystemParm&, _
    ByVal hStrBarParms&, ByVal hBarsArray&, _
    ByVal hStrTrades&, ByVal hStrConsolidated&, ByVal hStrByRule&, _
    ByVal hStrParmValues&, ByVal CallbackFunction&, ByVal hStrTimes&) As Long

Public Declare Function ExecuteSystemAdv Lib "NavEngineAdv.dll" Alias "ExecuteSystem" _
    (ByVal hStrParms&, ByVal hTblSystem&, ByVal hTblSystemRule&, ByVal hTblSystemParm&, _
    ByVal hStrBarParms&, ByVal hBarsArray&, _
    ByVal hStrTrades&, ByVal hStrConsolidated&, ByVal hStrByRule&, _
    ByVal hStrParmValues&, ByVal CallbackFunction&, ByVal hStrTimes&) As Long

Private Declare Function LoadFunctionSet Lib "NavEngine.dll" _
    (ByVal hstrFunctionSetName&, ByVal hTblLibrary&, ByVal hTblFunction&, _
    ByVal hTblFunctionParms&) As Long

Private Declare Function LoadFunctionSetAdv Lib "NavEngineAdv.dll" Alias "LoadFunctionSet" _
    (ByVal hstrFunctionSetName&, ByVal hTblLibrary&, ByVal hTblFunction&, _
    ByVal hTblFunctionParms&) As Long

Private Declare Function UnloadFunctionSet Lib "NavEngine.dll" _
    (ByVal hstrFunctionSetName&) As Long

Private Declare Function UnloadFunctionSetAdv Lib "NavEngineAdv.dll" Alias "UnloadFunctionSet" _
    (ByVal hstrFunctionSetName&) As Long

' TLB: this function does not appear to be used
''Private Declare Function RefreshFunctionSet Lib "NavEngine.dll" _
    (ByVal hstrFunctionSetName&, ByVal hTblLibrary&, ByVal hTblFunction&, _
    ByVal hTblFunctionParms&) As Long


Private Type mPrivate
    TASResults As cGdTree ' collection of cGdTables (one results table for each of their nIndID's)
End Type
Private m As mPrivate

Public Function LibraryField(ByVal LibFld As etblLibrary) As Long
    LibraryField = LibFld
End Function
Public Function FunctionField(ByVal FuncFld As etblFunction) As Long
    FunctionField = FuncFld
End Function
Public Function FuncParmField(ByVal FuncParmFld As etblFunctionParm) As Long
    FuncParmField = FuncParmFld
End Function
Public Function RuleField(ByVal RuleFld As etblRule) As Long
    RuleField = RuleFld
End Function
Public Function TradesHdrField(ByVal HdrField As eTradesHeader) As Long
    TradesHdrField = HdrField
End Function
Public Function MarketType(ByVal MktType As eMarketTypes) As Long
    MarketType = MktType
End Function

Public Function Flex2Bool(ByVal iChecked As Integer) As Boolean
On Error GoTo ErrSection:

    Flex2Bool = (iChecked = flexChecked)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.Flex2Bool", eGDRaiseError_Raise
    
End Function

Public Function Bool2Flex(ByVal bBool As Boolean) As Integer
On Error GoTo ErrSection:

    If bBool = True Then
        Bool2Flex = flexChecked
    Else
        Bool2Flex = flexUnchecked
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.Bool2Flex", eGDRaiseError_Raise
    
End Function

' Return FunctionID for function with the coded name
Public Function GetFunctionIDFromCodedName(pCodedName As String) As Long
On Error GoTo ErrSection:

    Dim X&

    GetFunctionIDFromCodedName = 0

    For X = 1 To g.Functions.Count
        If UCase(pCodedName) = UCase(g.Functions.Item(X).CodedName) Then
            GetFunctionIDFromCodedName = g.Functions.Item(X).FunctionID
            Exit For
        End If
    Next X

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.GetFunctionIDFromCodedName", eGDRaiseError_Raise
    
End Function

' Calls the DLL function.
Public Function InitEngine(ByVal bLoadEngine As Boolean, Optional strErrMsg As String) As Long
On Error GoTo ErrSection:

    Dim rc&, strNavDB$
    Dim strArray As New cGdArray
    Static bEngineLoaded As Boolean
    
    strErrMsg = ""
    If bLoadEngine Then
        If bEngineLoaded Then Exit Function
    Else
        ' 1/10/2002 - Bill says that we should run exit no matter what (DAJ)
        'If Not bEngineLoaded Then Exit Function
    End If
    
    ChangePath App.Path 'in case need to load DLL
    
    strNavDB = App.Path & "\Libraries.MDB"
    
    ' Put parms into string array to pass to DLL function
    strArray(0) = "" 'Trim(strNavDB)
    strArray(1) = "" 'DbPassword
    strArray(2) = Str(kSN_MAX_ITERATIONS) 'max # iterations when optimizing
    strArray(3) = "" 'FieldEncryptKey
    
    ' Run the system
    If bLoadEngine Then
        Screen.MousePointer = 11
        
        ' Set the Ignore flags appropriately
        SetIgnoreFlags
        
        If HasNewNavEngine Then
            rc = InitSysNavEngineAdv(strArray.ArrayHandle, AddressOf RunVbFunctionCallback, 0&)
        End If
        'If FileLength(App.Path & "\Provided\SystemAdv.flg") < 3 Then
            rc = InitSysNavEngine(strArray.ArrayHandle, AddressOf RunVbFunctionCallback, 0&)
        'End If
        Screen.MousePointer = 0
        If rc = 0 Then
            rc = LoadEngineFunctions
            If rc = 0 Then
                bEngineLoaded = True
            Else
                ''Err.Raise vbObjectError + 1000, , "Error loading engine functions"
                InfBox "Error loading engine functions: " & CStr(rc), "!", , "Error"
            End If
        Else
            strErrMsg = strArray(strArray.Size - 1)
            If Len(Trim(strErrMsg)) = 0 Then strErrMsg = "InitEngine Failed"
            strErrMsg = "Error " & CStr(rc) & ": " & strErrMsg
            InfBox strErrMsg, "!", , "InitEngine Error"
            ''Err.Raise vbObjectError + 1000, , strErrMsg
        End If
    Else
        rc = UnloadEngineFunctions
        If HasNewNavEngine Then
            rc = ExitSysNavEngineAdv(strArray.ArrayHandle, 0&, 0&)
        End If
        'If FileLength(App.Path & "\Provided\SystemAdv.flg") < 3 Then
            rc = ExitSysNavEngine(strArray.ArrayHandle, 0&, 0&)
        'End If
        If rc = 0 Then bEngineLoaded = False
    End If
    
    g.bDirtyFunctionLibrary = False

    InitEngine = rc

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.InitEngine", eGDRaiseError_Raise
    
End Function

' callback passed to engine in order to run the
' VB/Com type of DLL functions
Public Function RunVbFunctionCallback(ByVal hStrParms As Long, ByVal hArgs As Long) As Long
On Error GoTo CallbackExit
    
    Dim rc&, strFunction$
    Dim Func As Object
    
    rc = -998 'can't load function
    strFunction = gdGetStr(hStrParms, 0)
    
    Select Case UCase(strFunction)
        Case "GETBARSDATA"
            If hArgs <> 0 Then
                rc = Engine_GetBarsData(hArgs)
            End If
    
        Case "BUILTIN2.GETDMDATACURRENT"
''gdStartProfile 790
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_GetDmSnapshot(hArgs)
            End If
''gdStopProfile 790
            
        Case "BUILTIN2.GETDMDATAHISTORY"
''gdStartProfile 791
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_GetDmHistory(hArgs)
            End If
''gdStopProfile 791
            
        Case "BUILTIN2.GETTEXTDATA"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_GetTextData(hArgs)
            End If
            
        Case "BUILTIN2.GETSYMBOLID"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_GetSymbolID(hArgs)
            End If
            
        Case "BUILTIN2.HASMODULE"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_HasModule(hArgs)
            End If
            
        Case "BUILTIN2.ISFOREX"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_IsForex(hArgs)
            End If
        
        Case "BUILTIN2.PERCENTCOMPLETE"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_PercentComplete(hArgs)
            End If
        
        Case "BUILTIN2.BACKADJUSTMENT"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_BackAdjustment(hArgs)
            End If
        
        Case "BUILTIN2.RATEOFCHANGE2"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_RateOfChange(hArgs)
            End If
        
        Case "BUILTIN2.CALENDARSPREAD"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_CalendarSpread(hArgs)
            End If
        
        Case "BUILTIN2.AVGWEEKDAYMINUTEPRICEDIFF"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_AvgWeekdayMinutePriceDiff(hArgs)
            End If
        
        Case "BUILTIN2.TAS_RESULT"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = TAS_Results(hArgs)
            End If
        
        Case "BUILTIN2.TAS_NAVIGATOR"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = TAS_CalcIndicator(hArgs, "COMBO")
            End If
        
        Case "BUILTIN2.TAS_RATIO"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = TAS_CalcIndicator(hArgs, "RATIO")
            End If
        
        Case "BUILTIN2.TAS_STATICS"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = TAS_CalcIndicator(hArgs, "STATICPCL")
            End If
        
        Case "BUILTIN2.TAS_FLOATERS"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = TAS_CalcIndicator(hArgs, "FLOATPCL")
            End If
        
        Case "BUILTIN2.TAS_BOXES"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = TAS_CalcIndicator(hArgs, "SWINGRSI")
            End If
        
        Case "BUILTIN2.TAS_VEGA"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = TAS_CalcIndicator(hArgs, "VEGA")
            End If
        
        Case "BUILTIN2.TAS_MARKETMAP"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = TAS_CalcIndicator(hArgs, "MARKETMAP")
            End If
        
        Case "BUILTIN2.TAS_RESULT"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = TAS_Results(hArgs)
            End If
        
        Case "BUILTIN2.JURIK_JMA"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Jurik_JMA(hArgs)
            End If
        
        Case "BUILTIN2.JURIK_RSX"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Jurik_RSX(hArgs)
            End If
            
        Case "BUILTIN2.JURIK_VEL"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Jurik_VEL(hArgs)
            End If
            
        Case "BUILTIN2.JURIK_DMX"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Jurik_DMX(hArgs)
            End If
        
        Case "BUILTIN2.JURIK_DMXPLUS"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Jurik_DMX(hArgs, 1)
            End If
        
        Case "BUILTIN2.JURIK_DMXMINUS"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Jurik_DMX(hArgs, -1)
            End If
            
        Case "BUILTIN2.PREDICTIONLABSPATH"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_PredLabs(hArgs)
            End If
        
        Case "BUILTIN2.POWERZONESDATA"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_PowerZonesData(hArgs)
            End If
        
        Case "BUILTIN2.FRACTZENTICKRANGE"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_FractZenRange(hArgs)
            End If
        
        Case "BUILTIN2.FZSYMBOLPROPERTY"
            If hArgs = 1 Then
                rc = 0
            Else
                rc = Engine_FZSymbolProperty(hArgs)
            End If
        
        Case Else
''gdStartProfile 792
            Set Func = CreateObject(Trim(strFunction))
            If Func Is Nothing Then GoTo CallbackExit
            If hArgs = 1 Then
                'engine just seeing if this function exists
                rc = 0
            Else
                rc = -997 'can't run function
                ''rc = Func.Run(hArgs)
                rc = Func.NavigatorEntryPoint(hArgs)
            End If
''gdStopProfile 792
    End Select
       
CallbackExit:
    Set Func = Nothing
    RunVbFunctionCallback = rc

End Function

' to retrieve bars data for the engine
Private Function Engine_GetBarsData(ByVal hArgs&) As Long
On Error GoTo Error_GetBarsData
    
    Dim rc&, i&, j&, nError&, nSymbolID&, strErrMsg$, dTime#, dElapsed#
    Dim dFromDate As Double, dToDate As Double
    Dim hBars As Long
    Dim Bars As New cGdBars, MinuteBars As cGdBars
    Dim bMinutized As Boolean
    
    nError = 0
    ' get gdBars arg
    If gdGetArgAsHandle(hArgs, 1, hBars) = 0 Then
        nError = 1
    ElseIf hBars = 0 Then
        nError = 1
    ' get other args
    ElseIf gdGetArgAsNumber(hArgs, 2, dFromDate) = 0 Then
        nError = 2
    ElseIf gdGetArgAsNumber(hArgs, 3, dToDate) = 0 Then
        nError = 3
    End If
    
    'Check for arguments error
    If nError <> 0 Then
        Engine_GetBarsData = nError Or &H200 '(parm match error)
        strErrMsg = "Error in GetBarsData" _
            & ":  Data type mismatch for arg " & CStr(nError)
        Exit Function
    End If

    ' get bars data for symbol
    dElapsed = gdTickCount
    Bars.SetBarsHandle hBars, False
    If Bars.Prop(eBARS_Periodicity) = ePRD_EachTick + 2 Then ' special flag to ignore full ticks
        bMinutized = True
    End If
    nSymbolID = Bars.Prop(eBARS_SymbolID)
    If nSymbolID = 0 Then
        nError = -900
    ElseIf DM_GetBars(Bars, nSymbolID, Bars.Prop(eBARS_Periodicity), dFromDate, dToDate) Then
        If g.RealTime.Active And dToDate > LastDailyDownload Then
            ' append today's data (since LDD)
            If Not bMinutized Then
                g.RealTime.SpliceBars Bars
            Else
                ' TLB 10/19/2010: we'll just pseudo-minutize today's data (ignoring volumes)
                Set MinuteBars = Bars.MakeCopy(True)
                MinuteBars.ArrayMask = eBARS_Intraday
                MinuteBars.Prop(eBARS_Periodicity) = ePRD_Minutes + 1
                g.RealTime.SpliceBars MinuteBars
                j = Bars.Size
                Bars.Size = j + MinuteBars.Size * 4 + 5
                For i = 0 To MinuteBars.Size - 1
                    ' subtract 1 second from the MinuteBar time
                    dTime = MinuteBars(eBARS_DateTime, i) - 1 / 86400#
                    ' and make sure the time is greater than previous "tick"
                    If dTime > Bars(eBARS_DateTime, j - 1) Then
                        Bars(eBARS_DateTime, j) = dTime
                        Bars(eBARS_Close, j) = MinuteBars(eBARS_Open, i)
                        If MinuteBars(eBARS_High, i) = MinuteBars(eBARS_Low, i) Then
                            ' for a flat-line bar, only need to check the open
                            j = j + 1
                        Else
                            ' if Close > Open, assume O->L->H->C
                            If MinuteBars(eBARS_Close, i) >= MinuteBars(eBARS_Open, i) Then
                                Bars(eBARS_DateTime, j + 1) = dTime
                                Bars(eBARS_Close, j + 1) = MinuteBars(eBARS_Low, i)
                                Bars(eBARS_DateTime, j + 2) = dTime
                                Bars(eBARS_Close, j + 2) = MinuteBars(eBARS_High, i)
                            Else ' else assume O->H->L->C
                                Bars(eBARS_DateTime, j + 1) = dTime
                                Bars(eBARS_Close, j + 1) = MinuteBars(eBARS_High, i)
                                Bars(eBARS_DateTime, j + 2) = dTime
                                Bars(eBARS_Close, j + 2) = MinuteBars(eBARS_Low, i)
                            End If
                            Bars(eBARS_DateTime, j + 3) = dTime
                            Bars(eBARS_Close, j + 3) = MinuteBars(eBARS_Close, i)
                            j = j + 4
                        End If
                    End If
                Next
                Bars.Size = j
                Set MinuteBars = Nothing
            End If
        End If
        dElapsed = Int(gdTickCount - dElapsed + 0.5)
        DebugLog Bars.Prop(eBARS_Symbol) & " " & Str(dFromDate) & "-" & Str(dToDate) & ", " & Str(Bars.Size) & " bars, ms = " & Str(dElapsed)
    Else
        nError = -901
    End If
    
    Engine_GetBarsData = nError
    Exit Function
    
Error_GetBarsData:
    Engine_GetBarsData = -999 '(unexpected error)
    Exit Function

End Function

Private Function Engine_GetSymbolID(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim rc&, nError&, nSymbolID&
    Dim strSymbol As String, strText As String
    Dim hDataMgr As Long, hBars As Long, hResults As Long, hString As Long
    Dim strErrMsg As String
    
    nError = 0
    ' get gdArray arg
    If gdGetArgAsHandle(hArgs, 1, hResults) = 0 Then
        nError = 1
    ' get gdBars arg
    ElseIf gdGetArgAsHandle(hArgs, 2, hBars) = 0 Then
        nError = 2
    ' get gdString arg
    ElseIf gdGetArgAsHandle(hArgs, 3, hString) = 0 Then
        nError = 3
    End If
    
    ' check for arguments error
    If nError <> 0 Then
        Engine_GetSymbolID = nError Or &H200 '(parm match error)
        strErrMsg = "Error in GetSymbolID" _
            & ":  Data type mismatch for arg " & CStr(nError)
        Exit Function
    End If
    
    'default: return null
    gdSetNum hResults, 0, gdNullValue(hResults)
    
If gdGetBarsNumProp(hBars, eBARS_SymbolID) = 11936 Then
    hBars = hBars
End If
    
    
    
    ' the text argument could be any of the following:
    '  "" -- get symbol ID for Market1
    '  "SECTOR" or "SUBSECTOR" -- get symbol ID for the sector or subsector of Market1
    '  "IBM" -- get symbol ID for IBM
    '  "SECTOR:IBM" or "SUBSECTOR:IBM" -- get symbol ID for the sector or subsector of IBM
    strText = UCase(Trim(gdGetStr(hString)))
    If strText = "SECTOR" Or strText = "SUBSECTOR" Then
        strText = strText & ":"
    End If
    If InStr(strText, ":") > 0 Then
        strSymbol = Parse(strText, ":", 2)
    Else
        strSymbol = strText
        strText = ""
    End If
   
    If Len(strSymbol) = 0 Then
        ' if empty text, get symbol ID for Market1
        nSymbolID = gdGetBarsNumProp(hBars, eBARS_SymbolID)
    Else
        ' get symbol ID for the symbol passed as text
        nSymbolID = GetSymbolID(strSymbol)
    End If

    If Len(strText) > 0 Then
        'check if Sector or Subsector of DataKind (if so,
        ' SymbolID and strDataKind will be changed)
        CheckForSectorDataKind nSymbolID, strText
    End If
    
    ' return SymbolID
    gdSetNum hResults, 0, nSymbolID
    
    Engine_GetSymbolID = nError
    Exit Function
    
ErrSection:
    Engine_GetSymbolID = -999 '(unexpected error)
    Exit Function
End Function

Private Function Engine_GetDmSnapshot(ByVal hArgs&) As Long
On Error GoTo Error_GetDmSnapshot
    
    Dim rc&, nError&, nSymbolID&
    Dim strBase$
    Dim lBarDate As Long
    Dim dValue As Double, lActiveDate As Long
    Dim hDataMgr As Long, hBars As Long, hResults As Long, hString As Long
    Dim DivTable As cGdTable
    
    Dim strErrMsg As String
    
    'Get each argument (from object passed by engine)
    Dim strDataKind As String
    Dim nDataKindID As Long
    Dim nMaxFillDays As Long
    
    Dim alDataIDs As New cGdArray
    Dim adValues As New cGdArray
    Dim alDates As New cGdArray
    Dim lIndex As Long
    
    Dim iLifetime As Integer
        
''gdStartProfile 770
    nError = 0
    ' get gdArray arg
    If gdGetArgAsHandle(hArgs, 1, hResults) = 0 Then
        nError = 1
    ' get gdBars arg
    ElseIf gdGetArgAsHandle(hArgs, 2, hBars) = 0 Then
        nError = 2
    ' get gdString arg
    ElseIf gdGetArgAsHandle(hArgs, 3, hString) = 0 Then
        nError = 3
    ' get numeric arg
    ElseIf gdGetArgAsNumber(hArgs, 4, dValue) = 0 Then
        nError = 4
    End If
''gdStopProfile 770
''gdStartProfile 771
    'Check for arguments error
    If nError <> 0 Then
        Engine_GetDmSnapshot = nError Or &H200 '(parm match error)
        strErrMsg = "Error in GetDmSnapshot" _
            & ":  Data type mismatch for arg " & CStr(nError)
        Exit Function
    End If
    nMaxFillDays = dValue
    strDataKind = gdGetStr(hString)
''gdStopProfile 771

''gdStartProfile 772

    'default: return null
    gdSetNum hResults, 0, gdNullValue(hResults)
   
    'get SymbolID
    nSymbolID = gdGetBarsNumProp(hBars, eBARS_SymbolID)
    ' TLB 10/2/2015: added this to make it like GetDataHistory has always been (i.e. converting to the primary symbol)
    If UCase(Chr(gdGetBarsNumProp(hBars, eBARS_SecurityType))) = "F" Then
        ' for futures, lookup the primary base symbol
        ' (the first symbol in record of SymbolMap.csv file)
        hString = gdGetBarsStrProp(hBars, eBARS_BaseSymbol)
        strBase = PrimaryFutureBase(gdGetStr(hString))
        gdDestroyString hString
        nSymbolID = GetSymbolID(strBase & "-057")
    End If
    If nSymbolID = 283 Then 'AA
        rc = rc
    End If
    
    ' TLB 2/18/2015: special case for Next Dividend Date
    If UCase(strDataKind) = "NEXTDIVDATE" Then
        strDataKind = ""
        Set DivTable = GetDividendsTable(nSymbolID, False, 0, LastDailyDownload)
        If DivTable.NumRecords > 0 Then
            dValue = DivTable(2, DivTable.NumRecords - 1)
            If dValue > LastDailyDownload Then
                gdSetNum hResults, 0, dValue
            End If
        End If
        Set DivTable = Nothing
    End If
    
    'check first if DataKindID got passed directly (e.g. "34")
    nDataKindID = Val(strDataKind)
''gdStopProfile 772
    If nDataKindID = 0 And Len(strDataKind) > 0 Then
        'check if Sector or Subsector of DataKind (if so,
        ' SymbolID and strDataKind will be changed)
''gdStartProfile 773
        CheckForSectorDataKind nSymbolID, strDataKind
''gdStopProfile 773
''gdStartProfile 774
        hString = gdCreateString(Len(strDataKind))
        gdSetStr hString, 0, strDataKind
        If DM_GetDataKindID(g.DMS, hString, nDataKindID) = 0 Then
            nDataKindID = 0
        End If
        gdDestroyString hString
''gdStopProfile 774
    End If
    
    'get data for this symbol
    If nDataKindID <> 0 Then
        ' Default is to get the lifetime value from the database
        If nMaxFillDays <= -99999 Then
''gdStartProfile 775
            If DM_GetDataKindLifetime(g.DMS, nDataKindID, iLifetime) <> 0 Then
                nMaxFillDays = iLifetime
            Else
                nMaxFillDays = 0
            End If
''gdStopProfile 775
        End If
''gdStartProfile 776
        If DM_GetAllSnapData(nSymbolID, alDataIDs, adValues, alDates) Then
''gdStopProfile 776
''gdStartProfile 777
            'see if DataKind is found
            If alDataIDs.BinarySearch(nDataKindID, lIndex) Then
                dValue = adValues(lIndex)
                lActiveDate = alDates(lIndex)
                
                'see if date is "in range"
                lBarDate = Int(gdBarsData(hBars, eBARS_DateTime, gdGetSize(hBars) - 1))
                If lBarDate > 0 And lActiveDate > 0 Then
                    If lBarDate = lActiveDate Then
                        'exact date match
                        gdSetNum hResults, 0, dValue
                    ElseIf nMaxFillDays > 0 And lBarDate > lActiveDate Then
                        'post-fill data (up to "n" days after value)
                        If lBarDate <= lActiveDate + nMaxFillDays Then
                            gdSetNum hResults, 0, dValue
                        End If
                    ElseIf nMaxFillDays < 0 And lBarDate < lActiveDate Then
                        'pre-fill data (up to "n" days prior to value)
                        If lBarDate >= lActiveDate + nMaxFillDays Then
                            gdSetNum hResults, 0, dValue
                        End If
                    End If
                End If
            End If
        End If
''gdStopProfile 777
''gdStopProfile 776
    End If
    
    Engine_GetDmSnapshot = nError
    Exit Function
    
Error_GetDmSnapshot:
    Engine_GetDmSnapshot = -999 '(unexpected error)
    Exit Function

End Function

'Get "associated data" for a symbol from the Data Manager.
' Returns an array of values with a one-to-one correspondence
' with the bars (i.e. lined up with the dates of the bars).
'If MaxFillDays = 0: values only for dates with data (otherwise Null).
'If MaxFillDays > 0: post-fills values (up to "n" days after previous value).
'If MaxFillDays < 0: pre-fills values (up to "n" days before next value).
Private Function Engine_GetDmHistory(ByVal hArgs&) As Long
On Error GoTo Error_GetDmHistory
    
    Dim rc&, nError&, nSymbolID&, dResult#, nFromBar&
    Dim dCurDate#, dBarValue#, dStartDate#
    Dim aDates As New cGdArray, aValues As New cGdArray
    Dim nDataPos As Long, nBar As Long
    Dim hDataMgr As Long
    Dim hResults&, hValues&, hDates&
    
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    Args.ArgsHandle = hArgs
    
    'Get each argument (from object passed by engine)
    Dim Results As New cGdArray
    Dim Bars As New cGdBars
    Dim strDataKind As String
    Dim gdsDataKind As New cGdArray
    Dim nDataKindID As Long
    Dim nMaxFillDays As Long, nMaxFill As Long
    Dim bUseActiveDates As Boolean
    Dim bIsCot As Boolean
    Dim strBase As String
    
    Dim iLifetime As Integer
          
    Args.GetArg Results
    Args.GetArg Bars
    Args.GetArg gdsDataKind 'nDataKindID
    Args.GetArg nMaxFillDays
    
    ' get nFromBar (greater than 0 when just need recalc for last bar)
    nFromBar = Args.FromBar
    If nFromBar = -2 Then Exit Function
    If nFromBar < 0 Then nFromBar = 0

    ''bUseActiveDates = True
    
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_GetDmHistory = Args.Error
        strErrMsg = "Error in GetDmHistory" _
            & ":  " & Args.ErrorMessage
        Exit Function
    End If
        
    'default: return null
    aDates.Create eGDARRAY_Longs
    aValues.Create eGDARRAY_Doubles
   
    'get SymbolID
    nSymbolID = Bars.Prop(eBARS_SymbolID)
    If Bars.SecurityType = "F" Then
        ' for futures, lookup the primary base symbol
        ' (the first symbol in record of SymbolMap.csv file)
        strBase = PrimaryFutureBase(Bars.Prop(eBARS_BaseSymbol))
        nSymbolID = GetSymbolID(strBase & "-057")
    End If
    If nSymbolID = 283 Then 'AA
        rc = rc
    End If
           
    'check first if DataKindID got passed directly (e.g. "34")
    strDataKind = gdsDataKind(0)
    nDataKindID = Val(strDataKind)
    If nDataKindID = 0 Then
        'check if Sector or Subsector of DataKind (if so,
        ' SymbolID and strDataKind will be changed)
        CheckForSectorDataKind nSymbolID, strDataKind
        gdsDataKind(0) = strDataKind
        If UCase(Left(strDataKind, 4)) = "COT_" Then
            bIsCot = True
        End If
        ' get nDataKindID
        If DM_GetDataKindID(g.DMS, gdsDataKind.ArrayHandle, nDataKindID) = 0 Then
            nDataKindID = 0
        End If
    End If
    
    'get data for this symbol
    If nDataKindID <> 0 Then
        ' Default is to get the lifetime value from the database
        If nMaxFillDays <= -99999 Then
            If DM_GetDataKindLifetime(g.DMS, nDataKindID, iLifetime) <> 0 Then
                nMaxFillDays = iLifetime
            Else
                nMaxFillDays = 0
            End If
        End If
        
        ' get data
        dStartDate = 0
        If nFromBar > 0 Then
            dStartDate = Int(Bars(eBARS_DateTime, nFromBar)) - 5
            If nMaxFillDays > 0 Then
                dStartDate = dStartDate - nMaxFillDays
            End If
        End If
        If DM_GetDataHist(g.DMS, nSymbolID, nDataKindID, dStartDate, 99999, _
                aDates.ArrayHandle, aValues.ArrayHandle, bUseActiveDates) Then
            rc = rc
        End If
    End If
    
    If nError = 0 And aValues.Size > 0 Then
        ' use array handles directly in loop (for efficiency)
        hResults = Results.ArrayHandle
        hValues = aValues.ArrayHandle
        hDates = aDates.ArrayHandle
        
        'set value for each day
        nDataPos = 0
        nBar = nFromBar
        dBarValue = gdNullValue(hResults)
        For dCurDate = Int(Bars(eBARS_DateTime, nFromBar)) To 999999
            ' can skip weekends
            If IsWeekday(dCurDate) Then
                'find all bars (could be intraday) where the bar value goes
                '(put bar value in all bars prior to the new current date)
                Do While True
                    If nBar >= Bars.Size Then
                        Exit For '(done: no more bars)
                    End If
                    If Int(Bars(eBARS_DateTime, nBar)) >= dCurDate Then
                        Exit Do '(no more prior bars)
                    End If
                    gdSetNum hResults, nBar, dBarValue
                    nBar = nBar + 1
                    'if new bar has different date, null the bar value
                    If Int(Bars(eBARS_DateTime, nBar)) <> Int(Bars(eBARS_DateTime, nBar - 1)) Then
                        dBarValue = gdNullValue(hResults)
                    End If
                Loop
                
                'find place in array where text date >= cur date
                Do While nDataPos < gdGetSize(hDates)
                    If gdGetNum(hDates, nDataPos) >= dCurDate Then Exit Do
                    nDataPos = nDataPos + 1
                Loop
            
                'set the bar value if valid for this day
                If dCurDate = gdGetNum(hDates, nDataPos) Then
                    'exact date match
                    dBarValue = gdGetNum(hValues, nDataPos)
                ElseIf nMaxFillDays > 0 And nDataPos > 0 Then
                    'post-fill data (up to "n" days after previous value)
                    nMaxFill = nMaxFillDays
                    If bIsCot = True And dCurDate < 33880 And _
                            nMaxFillDays >= 6 And nMaxFillDays < 70 Then
                        nMaxFill = 70 '(to go across big gaps from old COT data - before Oct'92)
                    End If
                    If dCurDate <= gdGetNum(hDates, nDataPos - 1) + nMaxFill Then
                        dBarValue = gdGetNum(hValues, nDataPos - 1)
                    End If
                ElseIf nMaxFillDays < 0 And nDataPos < gdGetSize(hDates) Then
                    'pre-fill data (up to "n" days prior to next value)
                    If dCurDate >= gdGetNum(hDates, nDataPos) + nMaxFillDays Then
                        dBarValue = gdGetNum(hValues, nDataPos)
                    End If
                End If
            End If
        Next
    End If
    
    Engine_GetDmHistory = nError
    
    Exit Function
    
Error_GetDmHistory:
    Engine_GetDmHistory = -999 '(unexpected error)
    Exit Function

End Function

Private Sub CheckForSectorDataKind(nSymbolID&, strDataKind$)
On Error GoTo ErrSection:

    strDataKind = Trim(UCase(strDataKind))
    If Left(strDataKind, 7) = "SECTOR:" Then
        strDataKind = Trim(Mid(strDataKind, 8))
        nSymbolID = GetSectorID(nSymbolID, False)
    ElseIf Left(strDataKind, 10) = "SUBSECTOR:" Then
        strDataKind = Trim(Mid(strDataKind, 11))
        nSymbolID = GetSectorID(nSymbolID, True)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.CheckForSectorDataKind", eGDRaiseError_Raise
    
End Sub

Private Function Engine_GetTextData(hArgs As Long) As Long
    'Setup Args and error-handling
    On Error GoTo RunError
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    Args.ArgsHandle = hArgs
    
    'Declare all your "arguments"
    Dim Results As New cGdArray
    Dim Bars As New cGdBars
    Dim strFileName As String
    Dim nDataColumn As Long
    Dim dMaxFillDays As Double

    'Get each argument (from object passed by engine)
    Args.GetArg Results
    Args.GetArg Bars
    Args.GetArg strFileName
    Args.GetArg nDataColumn
    Args.GetArg dMaxFillDays
    
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_GetTextData = Args.Error
        strErrMsg = "Error in Engine_GetTextData" _
            & ":  " & Args.ErrorMessage
        Exit Function
    End If
        
    'Now call your custom function
    '(it should return 0 for success,
    ' or a negative error number)
    Engine_GetTextData = GetTextData(Results, Bars, _
            strFileName, nDataColumn, dMaxFillDays)
    Exit Function
    
RunError:
    Engine_GetTextData = -999 '(unexpected error)
    Exit Function
End Function

'Get "associated data" for a symbol from the Data Manager.
' Returns an array of values with a one-to-one correspondence
' with the bars (i.e. lined up with the dates of the bars).
'If MaxFillDays = 0: values only for dates with data (otherwise Null).
'If MaxFillDays > 0: post-fills values (up to "n" days after previous value).
'If MaxFillDays < 0: pre-fills values (up to "n" days before next value).
Private Function GetTextData(Results As cGdArray, _
        Bars As cGdBars, ByVal strFileName As String, _
        ByVal nDataColumn As Long, _
        Optional ByVal dMaxFillDays As Double = 0) As Long

    Dim nBar&, rc&, nDataPos&, dCurDate#, dPrevDate#, dBarValue#, dTimeOffset#
    Static bTextDataIsIntraday As Boolean
    Static dTextDates() As Double, dTextValues() As Double
    Static nPrevDataColumn&, strPrevFilename$, dPrevFiledate#
    
    ' init variables
    If nDataColumn < 1 Then nDataColumn = 1 '(default)
    
    If Left(strFileName, 1) <> "\" And InStr(strFileName, ":") = 0 Then
        strFileName = AddSlash(App.Path) & strFileName
    End If
    
    ' if using same file as previous call and it hasn't changed,
    ' then don't need to reload it
    If strFileName <> strPrevFilename Or nDataColumn <> nPrevDataColumn Then
        dPrevFiledate = 0
    ElseIf FileDate(strFileName) <> dPrevFiledate Then
        dPrevFiledate = 0
    End If
    If dPrevFiledate = 0 Then
        nPrevDataColumn = nDataColumn
        strPrevFilename = strFileName
        dPrevFiledate = FileDate(strFileName)
        
        ' load dates and values from file into arrays
        rc = LoadValuesFromFile(strFileName, nDataColumn, _
                Results.NullValue, dTextDates(), dTextValues(), bTextDataIsIntraday)
    End If
    If rc < 0 Then
        'return error
        GetTextData = rc
        dPrevFiledate = 0
        Exit Function
    End If

#If 1 Then
    If bTextDataIsIntraday And Not Bars.IsIntraday Then
        'dTimeOffset = Bars.Prop(eBARS_EndTime) / 1440#
    End If

    ' set value for each bar date
    nDataPos = 0
    dPrevDate = -999999
    For nBar = 0 To Bars.Size - 1
        dBarValue = Results.NullValue
        dCurDate = Bars(eBARS_DateTime, nBar)
''dCurDate = Bars.SessionDate(nBar)
        If dCurDate > -99999 Then
If IsIDE Then
    If dCurDate = DateSerial(2014, 12, 10) Then
        nBar = nBar
    End If
End If
            
            dCurDate = gdFixDateTime(dCurDate + dTimeOffset)
            ' TLB 6/19/2008: if intraday times, need to subtract
            ' 1/10th second to allow for any rounding issues
            If dCurDate <> Int(dCurDate) Then
                dCurDate = dCurDate - 0.000001
            End If
            
            ' find first text date >= cur date
            Do While nDataPos <= UBound(dTextDates)
                If dTextDates(nDataPos) >= dCurDate Then Exit Do
                nDataPos = nDataPos + 1
            Loop
        
            ' set the bar value if valid for this day
            If dCurDate > dPrevDate And dCurDate <= dTextDates(nDataPos) Then
''If dCurDate = dTextDates(nDataPos) Then
                ' "exact" date match if between last bar date and this one
                dBarValue = dTextValues(nDataPos)
            ElseIf dMaxFillDays > 0 And nDataPos > 0 Then
                ' post-fill data (up to "n" days after previous value)
                If dCurDate <= dTextDates(nDataPos - 1) + dMaxFillDays Then
                    dBarValue = dTextValues(nDataPos - 1)
                End If
            ElseIf dMaxFillDays < 0 Then
                ' pre-fill data (up to "n" days prior to next value)
                If dCurDate >= dTextDates(nDataPos) + dMaxFillDays Then
                    dBarValue = dTextValues(nDataPos)
                End If
            End If
            dPrevDate = dCurDate
        End If
        Results.Num(nBar) = dBarValue
    Next
#Else
    'set value for each day
    nDataPos = 0
    nBar = 0
    dBarValue = Results.NullValue
    For dCurDate = Int(Bars(eBARS_DateTime, 0)) To 999999
        ' can skip weekends
        If IsWeekday(dCurDate) Then
            'find all bars (could be intraday) where the bar value goes
            '(put bar value in all bars prior to the new current date)
            Do While True
                If nBar >= Bars.Size Then
                    Exit For '(done: no more bars)
                End If
                If Int(Bars(eBARS_DateTime, nBar)) >= dCurDate Then
                    Exit Do '(no more prior bars)
                End If
                Results.Num(nBar) = dBarValue
                nBar = nBar + 1
                'if new bar has different date, null the bar value
                If Int(Bars(eBARS_DateTime, nBar)) <> Int(Bars(eBARS_DateTime, nBar - 1)) Then
                    dBarValue = Results.NullValue
                End If
            Loop
            
            'find place in array where text date >= cur date
            Do While nDataPos <= UBound(dTextDates)
                If dTextDates(nDataPos) >= dCurDate Then Exit Do
                nDataPos = nDataPos + 1
            Loop
        
            'set the bar value if valid for this day
            If dCurDate = dTextDates(nDataPos) Then
                'exact date match
                dBarValue = dTextValues(nDataPos)
            ElseIf nMaxFillDays > 0 And nDataPos > 0 Then
                'post-fill data (up to "n" days after previous value)
                If dCurDate <= dTextDates(nDataPos - 1) + nMaxFillDays Then
                    dBarValue = dTextValues(nDataPos - 1)
                End If
            ElseIf nMaxFillDays < 0 Then
                'pre-fill data (up to "n" days prior to next value)
                If dCurDate >= dTextDates(nDataPos) + nMaxFillDays Then
                    dBarValue = dTextValues(nDataPos)
                End If
            End If
        End If
    Next
#End If

    GetTextData = 0 'success
End Function

Private Function LoadValuesFromFile(ByVal strFileName$, _
        ByVal nDataColumn&, ByVal dNullValue As Double, _
        dDates() As Double, dValues() As Double, bIsIntraday As Boolean) As Long

    Dim nYear&, nMonth&, nDay&, nPos&, iLine&
    Dim strLine$, strDate$, strDelim$
    Dim dDate As Double, dValue As Double
    Dim saLines As New cGdArray, saFields As New cGdArray
    
    'init arrays
    ReDim dDates(0) As Double
    ReDim dValues(0) As Double
    'set first entry to null value with earliest possible date
    dDates(0) = DateSerial(1899, 12, 31)
    dValues(0) = dNullValue
    nPos = 0
    
    'read data file
    saLines.FromFile strFileName
    If saLines.Size = 0 Then
        LoadValuesFromFile = -1 'error: file does not exist
        Exit Function
    End If

    'read each line of file
    bIsIntraday = False
    For iLine = 0 To saLines.Size - 1
        'ignore lines that don't start with a digit
        strLine = Trim(saLines(iLine))
        If Len(strLine) > 0 And Left(strLine, 1) >= "0" _
                And Left(strLine, 1) <= "9" Then
            
            ' increment and resize arrays if necessary
            nPos = nPos + 1
            If nPos > UBound(dDates) Then
                ReDim Preserve dDates(UBound(dDates) + 1000) As Double
                ReDim Preserve dValues(UBound(dDates)) As Double
            End If
            
            'see if need to detect the field delimiter
            If Len(strDelim) = 0 Then
                If InStr(strLine, vbTab) > 0 Then
                    strDelim = vbTab
                ElseIf InStr(strLine, ",") > 0 Then
                    strDelim = ","
                ElseIf InStr(strLine, " ") > 0 Then
                    strDelim = " "
                Else
                    LoadValuesFromFile = -2 'error: no delimiter
                    Exit Function
                End If
            End If
                
            'parse line into fields
            saFields.SplitFields strLine, strDelim
            
            'see if a valid date
            strDate = Trim(saFields(0))
            If InStr(strDate, "/") > 0 Or InStr(strDate, "-") > 0 Then
                    ''Or InStr(strDate, ".") > 0 Then
                'get date from formatted string
                If Right(strDate, 1) = ":" Then
                    strDate = Trim(Left(strDate, Len(strDate) - 1))
                End If
                If Not IsDate(strDate) Then
                    LoadValuesFromFile = -3 'error: invalid date
                    Exit Function
                End If
                dDate = CVDate(strDate)
            ElseIf ValOfText(strDate) > 10000000 Then
                'get date as YYYYMMDD
                nDay = Int(ValOfText(strDate))
                nYear = nDay \ 10000
                nDay = nDay Mod 10000
                nMonth = nDay \ 100
                nDay = nDay Mod 100
                If nYear < 1000 Or nYear > 2999 _
                        Or nMonth < 1 Or nMonth > 12 _
                        Or nDay < 1 Or nDay > 31 Then
                    LoadValuesFromFile = -3 'error: invalid date
                    Exit Function
                End If
                dDate = DateSerial(nYear, nMonth, nDay)
            Else
                'date is already julian
                dDate = ValOfText(strDate)
            End If
            If dDate = Int(dDate) Then
                dDates(nPos) = dDate
            Else
                bIsIntraday = True
                dDates(nPos) = gdFixDateTime(dDate)
            End If
        
            'make sure dates in sequence
            If dDates(nPos) <= dDates(nPos - 1) Then
                LoadValuesFromFile = -4 'error: dates out of sequence
                Exit Function
            End If
            
            'get value in specified column
            dValues(nPos) = dNullValue
            If nDataColumn < saFields.Size Then
                strLine = Trim(saFields(nDataColumn))
                If Len(strLine) > 0 Then
                    dValues(nPos) = Val(strLine)
                End If
            End If
        End If
    Next

    ' set last entry to null value with latest possible date
    nPos = nPos + 1
    ReDim Preserve dDates(nPos) As Double
    ReDim Preserve dValues(nPos) As Double
    dDates(nPos) = DateSerial(2999, 12, 31)
    dValues(nPos) = dNullValue

    LoadValuesFromFile = 0 'success
End Function

' Returns the % complete for each bar:
' - if bar is complete, result is 100
' - if bar has not yet started, result is Null
' - if bar is Daily or MinutesPerBar, result is % of seconds into the current bar
' - if bar is Ticks or Vol per bar, result is % of ticks or vol into the current bar
' - this function does not work for any other bar period (returns Null for all bars)
Private Function Engine_PercentComplete(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim i&, nError&, nSessionDate&, nPPB&, nFromBar&, nLastBar&, hResults&
    Dim dResult#, dStart#, dEnd#, dCrossOver#, dLastTick#, dDiff#
    Dim nSegment&, dSegmentLength#
    Dim strErrMsg As String
    
    'Get each argument (from object passed by engine)
    Dim Args As New cGdArgs
    Dim Results As New cGdArray
    Dim Bars As New cGdBars
    Dim bWeightedDailyMode As Boolean
             
    Args.ArgsHandle = hArgs
    Args.GetArg Results
    Args.GetArg Bars
    Args.GetArg bWeightedDailyMode
    
    ' get nFromBar (greater than 0 when just need recalc for last bar)
    nFromBar = Args.FromBar
    If nFromBar = -2 Then Exit Function
    If nFromBar < 0 Then nFromBar = 0
   
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_PercentComplete = Args.Error
        strErrMsg = "Error in PercentComplete" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' get last bar with good data
    nLastBar = -1
    For i = Bars.Size - 1 To 0 Step -1
        If Bars(eBARS_Close, i) <> kNullData Then
            nLastBar = i
            Exit For
        End If
    Next
    
'Enum eBarsPeriodType
'    ' Intraday types
'    ePRD_EachTick = &H1000000   '  16777216
'    ePRD_Ticks = &H2000000      '  33554432
'    ePRD_Minutes = &H3000000    '  50331648
'    ePRD_IntBreakout = &H5000000 ' 83886080
'    ePRD_IntRenko = &H6000000   ' 100663296
'    ePRD_IntKagi = &H7000000    ' 117440512
'    ePRD_IntPF = &H8000000      ' 134217728
'    ePRD_IntVol = &H9000000     ' 150994944
'    ' End-of-day types
'    ePRD_Days = &H11000000      ' 285212672
'    ePRD_Weeks = &H12000000     ' 301989888
'    ePRD_Months = &H13000000    ' 318767104
'    ePRD_Quarters = &H14000000  ' 335544320
'    ePRD_Years = &H15000000     ' 352321536
'    ePRD_EodRenko = &H16000000  ' 369098752
'    ePRD_EodKagi = &H17000000   ' 385875968
'    ePRD_EodPF = &H18000000     ' 402653184
'    ePRD_EodVol = &H19000000    ' 419430400
'    ePRD_EodBreakout = &H20000000 '536870912
'End Enum
    dResult = kNullData
    nPPB = Bars.Prop(eBARS_PeriodsPerBar)
    If nLastBar >= 0 And nPPB > 0 Then
        ' get the crossover time
        nSessionDate = Bars.SessionDate(nLastBar)
        dCrossOver = Bars.Prop(eBARS_CrossoverTime)
        If dCrossOver = 0 Then dCrossOver = 1439
        dCrossOver = nSessionDate + dCrossOver / 1440#
        ' get the time of the last tick
        dLastTick = nSessionDate + Bars.Prop(eBARS_LastTickTime) / 1440#
        If dLastTick > dCrossOver Then dLastTick = dLastTick - 1
        
        Select Case Bars.Prop(eBARS_PeriodType)
        Case ePRD_Ticks
            dResult = (Bars(eBARS_UpTicks, nLastBar) + Bars(eBARS_DownTicks, nLastBar)) _
                        / Bars.Prop(eBARS_PeriodsPerBar)
        Case ePRD_IntVol
            dResult = Bars(eBARS_Vol, nLastBar) / Bars.Prop(eBARS_PeriodsPerBar)
            
        Case ePRD_Minutes
            ' get total time for this bar
            dDiff = nPPB / 1440# ' (is usually the specified minutes per bar)
            If dDiff > Bars(eBARS_DateTime, nLastBar) - Bars(eBARS_DateTime, nLastBar - 1) Then
                ' but an exception if smaller (e.g. last bar of the day)
                dDiff = Bars(eBARS_DateTime, nLastBar) - Bars(eBARS_DateTime, nLastBar - 1)
            End If
            If dDiff > 0 Then
                ' %Complete = (LastTick - Start) / (End - Start)
                dStart = Bars(eBARS_DateTime, nLastBar) - dDiff
                dResult = (dLastTick - dStart) / dDiff
            End If
        
        Case ePRD_Days
            If nPPB = 1 Then
                If Not bWeightedDailyMode Then
                    ' get total time for this bar
                    dDiff = (Bars.Prop(eBARS_EndTime) - Bars.Prop(eBARS_StartTime)) / 1440#
                    If dDiff < 0 Then dDiff = dDiff + 1
                    If dDiff > 0 Then
                        ' %Complete = (LastTick - Start) / (End - Start)
                        dStart = nSessionDate + Bars.Prop(eBARS_StartTime) / 1440#
                        If dStart > dCrossOver Then dStart = dStart - 1
                        dResult = (dLastTick - dStart) / dDiff
                    End If
                Else
                    ' do a "weighted completion" based on each 30-minute segment of the day
                    ' (but ignore if tick from previous night or before market opens)
                    dStart = Bars.Prop(eBARS_StartTime)
                    If dStart < 1 Or dStart > Bars.Prop(eBARS_EndTime) Or _
                        (dStart > Bars.Prop(eBARS_CrossoverTime) And Bars.Prop(eBARS_CrossoverTime) > 0) Then
                            dStart = 570 ' (if overnight-type session, default to start of stock market)
                    End If
                    dEnd = Bars.Prop(eBARS_EndTime)
                    If dEnd < dStart Then
                        dEnd = 960 ' (default to end of stock market)
                    End If
                    dDiff = Bars.Prop(eBARS_LastTickTime) - dStart ' (minutes since start of day)
                    If Int(dLastTick) = Int(dCrossOver) And dDiff > 0 Then
                        ' calculate length of each segment (there are 13 segments
                        ' per day -- which makes 30 minutes per segment for stocks)
                        dSegmentLength = (dEnd - dStart) / 13
                        ' calculate which segment we're in now (0 through 12)
                        nSegment = Int(dDiff / dSegmentLength)
                        ' calculate how far through this segment we are now (e.g. 1/10th into it)
                        dDiff = (dDiff - nSegment * dSegmentLength) / dSegmentLength
                        ' calculate % complete based on how far through the current segment
                        Select Case nSegment
                        Case 0
                            dResult = 0 + 0.16 * dDiff
                        Case 1
                            dResult = 0.16 + 0.12 * dDiff
                        Case 2
                            dResult = 0.28 + 0.09 * dDiff
                        Case 3
                            dResult = 0.37 + 0.07 * dDiff
                        Case 4
                            dResult = 0.44 + 0.06 * dDiff
                        Case 5
                            dResult = 0.5 + 0.05 * dDiff
                        Case 6
                            dResult = 0.55 + 0.03 * dDiff
                        Case 7
                            dResult = 0.58 + 0.05 * dDiff
                        Case 8
                            dResult = 0.63 + 0.04 * dDiff
                        Case 9
                            dResult = 0.67 + 0.05 * dDiff
                        Case 10
                            dResult = 0.72 + 0.07 * dDiff
                        Case 11
                            dResult = 0.79 + 0.09 * dDiff
                        Case 12
                            dResult = 0.88 + 0.12 * dDiff
                        Case Else ' (if at or beyond the session end)
                            dResult = 1
                        End Select
                    End If
                End If
            End If
        
        Case ePRD_Weeks, ePRD_Months, ePRD_Quarters, ePRD_Years
            ' for periods > Daily, just use LastDailyDownload to calc % of weekdays complete in the period
            If nPPB = 1 Then
                dStart = 0  ' # of weekdays completed in the period
                dEnd = 0    ' total # of weekdays in the period
                ' start from end of period and work backwards ...
                For nSessionDate = Bars.SessionDate(nLastBar) To 0 Step -1
                    ' count it if a weekday
                    If IsWeekday(nSessionDate) Then
                        dEnd = dEnd + 1 ' total # of weekdays in the period
                        If nSessionDate <= LastDailyDownload Then
                            dStart = dStart + 1 ' # of weekdays completed
                        End If
                    End If
                    ' stop when get to beginning of the period
                    Select Case Bars.Prop(eBARS_PeriodType)
                    Case ePRD_Weeks
                        If Weekday(nSessionDate) = vbMonday Then
                            Exit For
                        End If
                    Case ePRD_Months
                        If Month(nSessionDate) <> Month(nSessionDate - 1) Then
                            Exit For
                        End If
                    Case ePRD_Quarters
                        If Int((Month(nSessionDate) - 1) / 3) <> Int((Month(nSessionDate - 1) - 1) / 3) Then
                            Exit For
                        End If
                    Case ePRD_Years
                        If Year(nSessionDate) <> Year(nSessionDate - 1) Then
                            Exit For
                        End If
                    End Select
                Next
                ' % complete = # weekdays completed / total # weekdays
                If dEnd > 0 Then
                    dResult = dStart / dEnd
                End If
            End If
        End Select
        
        If dResult > kNullData Then
            ' but bar considered complete if that session's daily download already done
            If Bars.SessionDate(nLastBar) <= LastDailyDownload Or dResult > 1 Then
                dResult = 1
            End If
        End If
    End If
    
    ' every bar before the last bar is complete
    hResults = Results.ArrayHandle
    For i = nFromBar To Bars.Size - 1
        If i < nLastBar Then
            gdSetNum hResults, i, 100
        ElseIf i > nLastBar Or dResult < 0 Then
            gdSetNum hResults, i, kNullData
        Else
            gdSetNum hResults, i, dResult * 100
        End If
    Next
    
    Engine_PercentComplete = nError
    Exit Function
    
ErrSection:
    Engine_PercentComplete = -999 '(unexpected error)
    Exit Function
End Function

' return BackAdjust amount for each bar (e.g. 057 minus 067)
Private Function Engine_BackAdjustment(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim i&, nError&, hResults&, nFromBar&
    Dim iFromBar&, iToBar&, iRec&, nRollDate&, dCumDelta#
    Dim strErrMsg As String
    Dim Table As cGdTable
    
    'Get each argument (from object passed by engine)
    Dim Args As New cGdArgs
    Dim Results As New cGdArray
    Dim Bars As New cGdBars
             
    Args.ArgsHandle = hArgs
    
    ' get nFromBar (greater than 0 when just need recalc for last bar)
    nFromBar = Args.FromBar
    If nFromBar = -2 Then
        Args.InstanceMemPtr = 0
        Exit Function
    End If
    If nFromBar < 0 Then nFromBar = 0
   
    Args.GetArg Results
    Args.GetArg Bars
   
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_BackAdjustment = Args.Error
        strErrMsg = "Error in BackAdjustment" & ":  " & Args.ErrorMessage
        Exit Function
    End If
       
    hResults = Results.ArrayHandle
    If Bars.SecurityType <> "F" Then
        ' just set to zero if not a Future
        gdMakeConstantValue hResults, 0, Bars.Size
    Else
        If nFromBar = 0 Then
            ' start at end of roll table and start working backwards
            Set Table = GetRollsTable(Bars.Prop(eBARS_Symbol))
            dCumDelta = 0
            iToBar = Bars.Size
            For iRec = Table.NumRecords - 1 To 0 Step -1
                ' find bar # of roll date
                nRollDate = Table.Num(1, iRec)
                iFromBar = Bars.FindDateTime(nRollDate)
                ' back up until session date is before roll date (e.g. if intraday data)
                For i = iFromBar - 1 To 0 Step -1
                    If Bars.SessionDate(i) < nRollDate Then
                        Exit For
                    End If
                    iFromBar = i
                Next
                ' set results for this chunk of dates
                For i = iFromBar To iToBar - 1
                    gdSetNum hResults, i, dCumDelta
                Next
                If iFromBar = 0 Then
                    Exit For
                End If
                ' then setup for next set of deltas
                iToBar = iFromBar
                dCumDelta = dCumDelta + Table.Num(2, iRec)
            Next
        End If
        
        ' and set ending bars to null (after end of data)
        For i = Bars.Size - 1 To nFromBar Step -1
            If Bars(eBARS_Close, i) = kNullData Then
                gdSetNum hResults, i, kNullData
            ElseIf nFromBar > 0 Then
                gdSetNum hResults, i, 0
            Else
                Exit For
            End If
        Next
    End If
    
    Engine_BackAdjustment = nError
    Exit Function
    
ErrSection:
    Engine_BackAdjustment = -999 '(unexpected error)
    Exit Function
End Function


Private Function Engine_RateOfChange(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim i&, nError&, nSessionDate&, nFromBar&, nLastBar&, hResults&, hBars&
    Dim dResult#
    Dim strErrMsg As String
    
    'Get each argument (from object passed by engine)
    Dim Args As New cGdArgs
    Dim Results As New cGdArray
    Dim Bars As New cGdBars
             
    Args.ArgsHandle = hArgs
    Args.GetArg Results
    Args.GetArg Bars
    
    ' get nFromBar (greater than 0 when just need recalc for last bar)
    nFromBar = Args.FromBar
    If nFromBar = -2 Then Exit Function
    If nFromBar < 0 Then nFromBar = 0
   
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_RateOfChange = Args.Error
        strErrMsg = "Error in RateOfChange" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
'for Futures = 100 * (Close of "xx-067" - Close.1 of "xx-067") / Close.1 of "xx-057"
'for others (only going back while prices > 0) = 100 * (Close - Close.1) / Close.1
#If 0 Then
    ' get last bar with good data
    nLastBar = -1
    For i = Bars.Size - 1 To 0 Step -1
        If Bars(eBARS_Close, i) <> kNullData Then
            nLastBar = i
            Exit For
        End If
    Next
    
    dResult = kNullData
'nPPB = Bars.Prop(eBARS_PeriodsPerBar)
    
    ' every bar before the last bar is complete
    hResults = Results.ArrayHandle
    hBars = Bars.BarsHandle
    For i = nFromBar To Bars.Size - 1
        'dResult = gdgetnum(
        If i < nLastBar Then
            gdSetNum hResults, i, 100
        ElseIf i > nLastBar Or dResult < 0 Then
            gdSetNum hResults, i, kNullData
        Else
            gdSetNum hResults, i, dResult * 100
        End If
    Next
#End If

    Engine_RateOfChange = nError
    Exit Function
    
ErrSection:
    Engine_RateOfChange = -999 '(unexpected error)
    Exit Function
End Function

Private Function Engine_CalendarSpread(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim i&, d#, nError&, nFromBar&, hResults&, hBars&, hSpread&, hDates1&, hDates2&
    Dim dResult#, iNumContractsOut&, strSymbol$, strTemp$, nFromDate&, nToDate&, nBar2&
    Dim bBackAdjusted As Boolean
    Dim strErrMsg As String
    
    'Get each argument (from object passed by engine)
    Dim Args As New cGdArgs
    Dim Results As New cGdArray
    Dim Spread As New cGdArray
    Dim Bars As New cGdBars
    Dim Bars2 As New cGdBars
             
    Args.ArgsHandle = hArgs
    Args.GetArg Results
    Args.GetArg Bars
    Args.GetArg iNumContractsOut
    If Args.Count >= 4 Then
        Args.GetArg bBackAdjusted
    End If
    
    ' get nFromBar (greater than 0 when just need recalc for last bar)
    nFromBar = Args.FromBar
    If nFromBar = -2 Then Exit Function
    If nFromBar < 0 Then nFromBar = 0
 
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_CalendarSpread = Args.Error
        strErrMsg = "Error in CalendarSpread" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    
    'Set Bars2 = Bars.MakeCopy
    'If nFromBar > 0 Then
        ' if last bar recalc, then only hand the last part of the bars (more efficient)
    '    Bars2.DeleteFirstBars nFromBar
    'End If
    
    ' load pit symbol (which has all the settles)
    nFromDate = Bars(eBARS_DateTime, nFromBar)
    nToDate = Bars(eBARS_DateTime, Bars.Size - 1)
    If nFromDate > 0 And Bars.Prop(eBARS_SecurityType) = Asc("F") And Not Bars.IsIntraday Then
        strSymbol = Bars.Prop(eBARS_Symbol)
        strTemp = ConvertFutureSymbol(strSymbol, ePitSymbol)
        If Len(strTemp) > 0 Then
            strSymbol = strTemp
        End If
        DM_GetBars Bars2, strSymbol, Bars.Prop(eBARS_Periodicity), nFromDate, nToDate
        If Bars2.Size > 0 Then
            Spread.Create eGDARRAY_Doubles, 0
            d = gdTickCount
            If DM_CalendarSpread(Spread, Bars2, iNumContractsOut, bBackAdjusted) Then
                ' need to align results by date of Bars
                hResults = Results.ArrayHandle
                hSpread = Spread.ArrayHandle
                hDates1 = Bars.ArrayHandle(eBARS_DateTime)
                hDates2 = Bars2.ArrayHandle(eBARS_DateTime)
                nBar2 = 0
                For i = nFromBar To Bars.Size - 1
                    Do While nBar2 + 1 < Bars2.Size
                        'If Bars2(eBARS_DateTime, nBar2 + 1) > Bars(eBARS_DateTime, i) Then
                        If gdGetNum(hDates2, nBar2 + 1) > gdGetNum(hDates1, i) Then
                            Exit Do
                        End If
                        nBar2 = nBar2 + 1
                    Loop
                    gdSetNum hResults, i, gdGetNum(hSpread, nBar2)
                Next
            End If
            'frmTest.AddList "msTot = " & Str(gdTickCount - d)
        End If
    End If

    Set Bars = Nothing
    Set Bars2 = Nothing
    Set Spread = Nothing
    Set Results = Nothing
    Set Args = Nothing
    
    Engine_CalendarSpread = nError
    Exit Function
    
ErrSection:
    Engine_CalendarSpread = -999 '(unexpected error)
    Exit Function
End Function

Private Function Engine_AvgWeekdayMinutePriceDiff(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim i&, n&, d#, nMinutesPerBar&, nMinuteOfWeek&
    Dim nError&, nFromBar&, hResults&, hBars&, nYearsBack&
    Dim dResult#, dDiff#, dTime#
    Dim s$, strSymbol$, strFile$, strErrMsg$
    Dim hMinuteOfWeek&, hFile&
    Dim TypeByte As Byte
    
    'Get each argument (from object passed by engine)
    Dim Args As New cGdArgs
    Dim Results As New cGdArray
    Dim Bars As New cGdBars
                
    Args.ArgsHandle = hArgs
    ' get nFromBar (greater than 0 when just need recalc for last bar)
    hMinuteOfWeek = Args.InstanceMemPtr
    nFromBar = Args.FromBar
    
    If nFromBar = -2 Then
        If hMinuteOfWeek <> 0 Then
            gdDestroyArray hMinuteOfWeek
        End If
        Args.InstanceMemPtr = 0
        Exit Function
    End If
    If nFromBar < 0 Then nFromBar = 0
    
    Args.GetArg Results
    Args.GetArg Bars
    Args.GetArg nYearsBack

    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_AvgWeekdayMinutePriceDiff = Args.Error
        strErrMsg = "Error in AvgWeekdayMinutePriceDiff" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
'AvgWeekdayMinutePriceDiff{
'- only valid for minute bars of futures
'- get electronic base symbol?
'- read minute data from file (store in a static array for efficiency while streaming)
'- for each bar of data:
'    - get start and end minute of bar (start = EndMin - MinPerBar)
'    - price diff = sum of price diffs for each minute (from weekday minute price diffs)
'    - get cumulative price diff?
    
    If Bars.Prop(eBARS_PeriodType) = ePRD_Minutes Then
        nMinutesPerBar = Bars.Prop(eBARS_PeriodsPerBar)
        If nMinutesPerBar > 0 Then 'And SecurityType(Bars, True) = "F" Then
            
            ' if the data array (the avg price diff of each weekday minute for this symbol)
            ' has not been read from the file yet, then get it now
            If hMinuteOfWeek = 0 Then
                If SecurityType(Bars, True) = "F" Then
                    strSymbol = ConvertFutureSymbol(Bars.Prop(eBARS_Symbol), eElectronicSymbol)
                    strSymbol = Parse(strSymbol, "-", 1) & "-"
                Else
                    strSymbol = Bars.Prop(eBARS_Symbol) & "!"
                End If
                If Len(strSymbol) > 0 Then
                    strFile = App.Path & "\WIA\" & strSymbol & Str(nYearsBack) & "yr.gda"
                    If FileExist(strFile) Then
                        hMinuteOfWeek = gdCreateArray(eGDARRAY_Doubles, 1440 * 7, 0)
                        hFile = gdFileOpen(strFile, "rb")
                        If hFile <> 0 Then
                            If gdFileBinaryIO(hFile, TypeByte, 1, False) = 1 Then
                                gdSerializeArray hMinuteOfWeek, hFile, False
                            End If
                            gdFileClose hFile
                            hFile = 0
                        End If
                    End If
                End If
            End If
            
            ' now use the data array to return the avg price diff of each weekday minute for this symbol
            If hMinuteOfWeek <> 0 Then
                hResults = Results.ArrayHandle
                For i = nFromBar To Bars.Size - 1
                    ' calc # of minutes into the week (since midnight Sat)
                    dTime = Bars(eBARS_DateTime, i)
                    nMinuteOfWeek = Round((dTime - Int(dTime)) * 1440) + (Weekday(Int(dTime)) - 1) * 1440
                    If nMinuteOfWeek <= 0 Or nMinuteOfWeek >= 1440 * 7 Then
                        Exit For
                    End If
                    ' sum the price diffs for each minute in this bar
                    dDiff = 0
                    For n = nMinuteOfWeek - nMinutesPerBar + 1 To nMinuteOfWeek
                        dDiff = dDiff + gdGetNum(hMinuteOfWeek, n)
                    Next
                    gdSetNum hResults, i, dDiff
                Next
            End If
            
        End If
    End If
   
    ' store the data array so we don't have to keep reading it from the file during streaming
    Args.InstanceMemPtr = hMinuteOfWeek

    Set Bars = Nothing
    Set Results = Nothing
    Set Args = Nothing
    
    Engine_AvgWeekdayMinutePriceDiff = nError
    Exit Function
    
ErrSection:
    Engine_AvgWeekdayMinutePriceDiff = -999 '(unexpected error)
    Exit Function
End Function

Private Function Engine_HasModule(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim nError&
    Dim strErrMsg As String
    
    'Get each argument (from object passed by engine)
    Dim Args As New cGdArgs
    Dim Results As New cGdArray
    Dim bIncludeSourceCode As Boolean
    Dim strText As String
                
    Args.ArgsHandle = hArgs
    Args.GetArg Results
    Args.GetArg strText
    Args.GetArg bIncludeSourceCode

    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_HasModule = Args.Error
        strErrMsg = "Error in HasModule" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' return single true/false if has enablement code
    gdSetNum Results.ArrayHandle, 0, Abs(HasModule(strText, bIncludeSourceCode))
        
    Set Args = Nothing
    Set Results = Nothing
    Engine_HasModule = nError
    Exit Function
    
ErrSection:
    Engine_HasModule = -999 '(unexpected error)
    Exit Function
End Function

Private Function Engine_IsForex(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim nError&
    Dim strErrMsg As String
    Dim strSymbol$
    
    'Get each argument (from object passed by engine)
    Dim Args As New cGdArgs
    Dim Results As New cGdArray
    Dim Bars As New cGdBars
                
    Args.ArgsHandle = hArgs
    Args.GetArg Results
    Args.GetArg Bars

    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_IsForex = Args.Error
        strErrMsg = "Error in IsForex" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' return single true/false if symbol is Forex
    strSymbol = Bars.Prop(eBARS_Symbol)
    gdSetNum Results.ArrayHandle, 0, Abs(IsForex(strSymbol))
        
    Set Args = Nothing
    Set Results = Nothing
    Set Bars = Nothing
    Engine_IsForex = nError
    Exit Function
    
ErrSection:
    Engine_IsForex = -999 '(unexpected error)
    Exit Function
End Function

Private Function Engine_FractZenRange(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim i&, nError&, nSessionDate&, nPrevSession&, nFromBar&, nLastBar&, hResults&, hBars&
    Dim dResult#, dBarDate#
    Dim strErrMsg$, strSymbol$
    
    'Get each argument (from object passed by engine)
    Dim Args As New cGdArgs
    Dim Results As New cGdArray
    Dim Bars As New cGdBars
             
    Args.ArgsHandle = hArgs
    Args.GetArg Results
    Args.GetArg Bars
    
    ' get nFromBar (greater than 0 when just need recalc for last bar)
    nFromBar = Args.FromBar
    If nFromBar = -2 Then Exit Function
    If nFromBar < 0 Then nFromBar = 0
   
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_FractZenRange = Args.Error
        strErrMsg = "Error in Engine_FractZenRange" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' get last bar with good data
    nLastBar = -1
    For i = Bars.Size - 1 To 0 Step -1
        If Bars(eBARS_Close, i) <> kNullData Then
            nLastBar = i
            Exit For
        End If
    Next
    
    strSymbol = ""
    nPrevSession = 0
    dResult = kNullData
    
    ' just return Nulls if not tick breakout bars
    If Bars.Prop(eBARS_PeriodType) = ePRD_IntBreakout Then
        If Bars.Prop(eBARS_FractZen) = 0 Then
            ' if not "merged FractZen" bars, just return tick breakout range (a constant)
            dResult = Bars.Prop(eBARS_PeriodsPerBar)
            strSymbol = ""
        Else
            ' else need to get tick range for each session
            strSymbol = Bars.Prop(eBARS_Symbol)
        End If
    End If
    
    ' every bar before the last bar is complete
    hResults = Results.ArrayHandle
    hBars = Bars.BarsHandle
    For i = nFromBar To Bars.Size - 1
        If Len(strSymbol) > 0 Then
            dBarDate = gdBarsData(hBars, eBARS_DateTime, i)
            nSessionDate = gdBarsSessionDate(hBars, dBarDate, 0)
            If nSessionDate <> nPrevSession Then
                nPrevSession = nSessionDate
                If i > nLastBar Then
                    dResult = kNullData
                Else
                    dResult = g.FractZen.GetFractZenRange(strSymbol, nSessionDate)
                End If
            End If
        End If
        If dResult <= 0 Then
            gdSetNum hResults, i, kNullData
        Else
            gdSetNum hResults, i, dResult
        End If
    Next

    Engine_FractZenRange = nError
    Exit Function
    
ErrSection:
    Engine_FractZenRange = -999 '(unexpected error)
    Exit Function
End Function

Private Function Engine_FZSymbolProperty(ByVal hArgs&) As Long
On Error GoTo ErrSection:
    
    Dim nFromBar&, strSymbol$, dResult#
    Dim strErrMsg As String
    Static strIniFile$
       
    'Get each argument (from object passed by engine)
    Dim Args As New cGdArgs
    Dim Results As New cGdArray
    Dim Bars As New cGdBars
    Dim strProp As String
             
    Args.ArgsHandle = hArgs
    Args.GetArg Results
    Args.GetArg Bars
    Args.GetArg strProp
    
    ' get nFromBar (greater than 0 when just need recalc for last bar)
    nFromBar = Args.FromBar
    If nFromBar = -2 Then Exit Function
    If nFromBar < 0 Then nFromBar = 0
   
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_FZSymbolProperty = Args.Error
        strErrMsg = "Error in Engine_FZSymbolProperty" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' don't need to keep checking if already set
    If nFromBar <= 0 Or Results(0) = kNullData Then
        dResult = kNullData
        If Len(strProp) > 0 Then
            ' get the Prop Value for this symbol
            If Len(strIniFile) = 0 Then
                strIniFile = App.Path & "\Provided\FractZenTest.INI" ' override just for John's machine
                If Not FileExist(strIniFile) Then
                    strIniFile = App.Path & "\Provided\FractZen.INI"
                End If
            End If
            strSymbol = g.FractZen.FzSymbol(Bars.Prop(eBARS_Symbol))
            dResult = GetIniFileProperty(strSymbol, kNullData, strProp, strIniFile)
            If dResult = kNullData Then
                ' if not exist for the symbol, then get default value
                dResult = GetIniFileProperty(strProp, kNullData, "Defaults", strIniFile)
            End If
        End If
        Results(0) = dResult
    End If
    
    Engine_FZSymbolProperty = 0
    Exit Function
    
ErrSection:
    Engine_FZSymbolProperty = -999 '(unexpected error)
    Exit Function
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetIgnoreFlags
'' Description: Walk through the Librarys table and set the Ignore flags
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetIgnoreFlags()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset for the Librarys table
    Dim t As TableDef                   ' TableDef for the Librarys table
    Dim bIgnore As Boolean              ' Whether or not to ignore this Library
    Dim bTransactionStarted As Boolean
    Dim strMod$, dExpire#
    Dim strPath As String
    Dim lPos As Long
    
    Dim bMessageShown As Boolean
    Static bMessagesAlreadyShown As Boolean
    
    Set t = g.dbNav.TableDefs("tblLibrarys")
    With t
        ' Only do this if the Ignore field exists in the Librarys table
        If ItemExists(.Fields, "Ignore") Then
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys];", dbOpenDynaset)
            ValidateCheckSums rs, "tblLibrarys"
            rs.MoveFirst
            Do While Not rs.EOF
                bIgnore = False
                
                ' See if the required mod is in the authorization string
                strMod = FixRequiredMod(NullChk(rs!RequiredMod))
                If Not HasModule(strMod) Then
                    bIgnore = True
                End If
                
                ' Ignore "LW Functions" for BetterTrades
                If UCase(rs!LibraryName) = "LW FUNCTIONS" Then
                    If ExtremeCharts >= 1 And Not HasModule("BAT") Then
                        bIgnore = True
                    End If
                End If
                    
                ' See if we are past the expiration date (if there is one)
                dExpire = NullChk(rs!Expiration, 0#)
                If dExpire > 0 Then
                    If Date > dExpire Then bIgnore = True
                End If
                
                ' Make sure that the DLL exists...
                If Not bIgnore And Len(rs!Path) > 0 And UCase(rs!Path) <> "BUILTIN2.DLL" _
                        And UCase(rs!Path) <> "BUILTIN.DLL" And Not rs!BuiltIn Then
                    ' 3rd-party DLL's should be in the "LibraryDLLs" folder
                    strPath = App.Path
                    lPos = At(strPath, "\", -1)
                    If lPos > 0 Then
                        strPath = Left(strPath, lPos) & "LibraryDLLs\"
                    Else
                        strPath = ""
                    End If
                    
                    If Not FileExist(AddSlash(strPath) & rs!Path) Then
                        bIgnore = True
                        If Not bMessagesAlreadyShown Then
                            bMessageShown = True
                            InfBox "The " & UCase(rs!LibraryName) & " library could not be loaded because " & UCase(rs!Path) & "|currently does not exist in|" & strPath, "!", , "Library Error"
                        End If
                    ElseIf UCase(rs!Path) = "PLANET.DLL" Then
                        ' make sure the Swedish Ephemeral DLL exists
                        If Not FileExist(App.Path & "\JPLdata\*.se1") Then
                            bIgnore = True
                            If Not bMessagesAlreadyShown Then
                                bMessageShown = True
                                InfBox "The Planetary library could not be loaded because the required planetary tables could not be found.  Please download the special file 'PLANETS' in order to enable this module.", "!", , "Library Error"
                            End If
                        End If
                    End If
                End If
                
                ' Make sure that the CheckSum is valid
                If rs!CheckSum = 0.5 Then bIgnore = True
                
                ' Reset the Ignore flag
                If rs!Ignore <> bIgnore Then
                    If Not bTransactionStarted Then
                        g.WrkJet.BeginTrans
                        bTransactionStarted = True
                    End If
                    rs.Edit
                    rs!Ignore = bIgnore
                    rs!CheckSum = BuildCheckSum(rs, "tblLibrarys")
                    rs.Update
                    KillRulesFile '(force a rebuild of the rules table)
                End If
                
                rs.MoveNext
            Loop
            If bTransactionStarted Then
                ' make sure to flush so Engine will be able to "see" the new flag settings
                bTransactionStarted = False
                g.WrkJet.CommitTrans dbForceOSFlush
                
                ' and also need to reload our functions
                Set g.Functions = Nothing
                g.Functions.Load
                FilterFunctions
            End If
        End If
    End With
    
    ' set flag so will only show the messages once
    If bMessageShown Then bMessagesAlreadyShown = True
    
ErrExit:
    Set t = Nothing
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    If bTransactionStarted Then g.WrkJet.Rollback
    Set t = Nothing
    Set rs = Nothing

End Sub

' Callback function passed to the engine.
Public Function EngineCallback(ByVal nIteration As Long, _
    ByVal hStrTrades As Long, ByVal hStrParmValues As Long) As Long
On Error GoTo ErrSection:
    
    Dim i&, s$
    Dim aTrades As New cGdArray, aParmValues As New cGdArray
    Dim nOptMode As eGDOptMode
    Static nTotalIterations&
    
    If nIteration < 0 Then Exit Function 'mostly for debugging the engine
    
    'Initialization: Pass total iterations and input values
    If hStrTrades = 0 Then
        nTotalIterations = nIteration
        aParmValues.Create eGDARRAY_Strings
        gdCopy aParmValues.ArrayHandle, hStrParmValues
        If nTotalIterations > 1 Then
            ''aParmValues.ToFile App.Path & "\parms1.chk" 'debugging
            
            If g.CurrentSystem.HighlightBarReport Then
                nOptMode = eGDOptMode_HighlightBarReport
            Else
                nOptMode = eGDOptMode_Optimization
            End If
            
            frmOptimizer.Init nTotalIterations, aParmValues, nOptMode
        End If
    
    'Pass current iteration, trades, and input values...
    ElseIf nTotalIterations > 1 Then
        
        aTrades.Create eGDARRAY_Strings
        aParmValues.Create eGDARRAY_Strings
        gdCopy aTrades.ArrayHandle, hStrTrades
        gdCopy aParmValues.ArrayHandle, hStrParmValues
        EngineCallback = frmOptimizer.Add(nIteration, aTrades, aParmValues)
        
    End If
    
    Set aTrades = Nothing
    Set aParmValues = Nothing
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.EngineCallback", eGDRaiseError_Raise
    
End Function

' "L",#1982-05-06#,392.300000,"New Rule",#1982-05-07#,393.150000,"New Exit Rule",162.5
Private Function FormatTrade$(ByVal strID$, Trades() As TradeStruct, ByVal Num%, ByVal rule_len%)
On Error GoTo ErrSection:

    Dim strPos$, strEntry$, strExit$, strRtrn$
    
    If Num >= 1 And Num <= UBound(Trades) Then
        strEntry = Trades(Num).strEntry
        strEntry = Pad(strEntry, rule_len, "L")
        strExit = Trades(Num).strExit
        strExit = Pad(strExit, rule_len, "L")
        If Trades(Num).bLong Then
            strPos = "L"
        Else
            strPos = "S"
        End If
        'NumStr$(ByVal Numb#, ByVal wid%, Optional ByVal dec% = 0)
        strRtrn = strID & NumStr(Num, 4) _
            & ": " & strPos & NumStr(Trades(Num).dProfit, 10, 2) _
            & " | " & Right(Str(Trades(Num).lEntryDate), 6) _
            & NumStr(Trades(Num).dEntryPrice, 0, 2) & " " & strEntry _
            & " | " & Right(Str(Trades(Num).lExitDate), 6) _
            & NumStr(Trades(Num).dExitPrice, 10, 2) & " " & strExit
' "%s%4d: %c %9.2f | %06d %10.4f %s | %06d %10.4f %s",
' strID.Ptr(), num+1, pos, trd[num].dProfit,
' trd[num].lEntryDate % 1000000, trd[num].dEntryPrice, strEntry.Ptr(),
' trd[num].lExitDate % 1000000, trd[num].dExitPrice, strExit.Ptr());
    End If
    
    FormatTrade = strRtrn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.FormatTrade", eGDRaiseError_Raise
    
End Function

Public Sub TradeDiff(lst As ListBox, ByVal strTSfile$, ByVal strSNfile$, _
    Optional ByVal rule_len% = 9)
On Error GoTo ErrSection:

    Dim i%, t%, iSN%, iTS%, lDate&, bExitLine As Boolean
    Dim strTemp$, strLine$, sa$(), cr$, iNumDiff&
    Dim trdSN() As TradeStruct, trdTS() As TradeStruct
    Dim Trade As TradeStruct
    Dim mult#, iPos%
    
    cr = Chr(13) & Chr(10)
    mult = 32#
    
    ' read SysNav trades
    i = FileToArray(strSNfile, sa())
    For i = 1 To UBound(sa)
        strLine = sa(i)
        If Len(strLine) > 0 Then
'"L",#1982-05-06#,392.300000,"New Rule",#1982-05-07#,393.150000,"New Exit Rule",162.5
            strTemp = Parse(strLine, ",", 1)
            strTemp = UCase(Trim(StripStr(strTemp, Chr(34))))
            If strTemp = "L" Then
                Trade.bLong = True
            Else
                Trade.bLong = False
            End If

            strTemp = Parse(strLine, ",", 2)
            Trade.lEntryDate = Val(StripStr(strTemp, "-#/"))
            Trade.dEntryPrice = Val(Parse(strLine, ",", 3))
            strTemp = Parse(strLine, ",", 4)
            Trade.strEntry = Trim(StripStr(strTemp, Chr(34)))

            strTemp = Parse(strLine, ",", 5)
            Trade.lExitDate = Val(StripStr(strTemp, "-#/"))
            Trade.dExitPrice = Val(Parse(strLine, ",", 6))
            strTemp = Parse(strLine, ",", 7)
            Trade.strExit = Trim(StripStr(strTemp, Chr(34)))

            Trade.dProfit = Val(Parse(strLine, ",", 8))
            
            If iSN Mod 100 = 0 Then ReDim Preserve trdSN(iSN + 100) As TradeStruct
            iSN = iSN + 1
            trdSN(iSN) = Trade
        End If
    Next
    ReDim Preserve trdSN(iSN) As TradeStruct

    ' read TradeStation trades
    i = FileToArray(strTSfile, sa())
    For i = 1 To UBound(sa)
        strLine = sa(i)
        If Len(strLine) > 0 And Left(strLine, 1) >= "0" And Left(strLine, 1) <= "9" Then
' 1Date,2Time,3Type,4Contracts,5Price,6Signal,(7Profit,8Cum)
' 05/06/82      Buy 1     392.30
' 05/07/82      LExit   1     393.15        $    425.00 $    425.00
            bExitLine = False
            strTemp = UCase(Trim(Parse(strLine, Chr(9), 3)))
            If strTemp = "BUY" Then
                Trade.bLong = True
            ElseIf strTemp = "SELL" Then
                Trade.bLong = False
            Else
                bExitLine = True
            End If

            strTemp = Parse(strLine, Chr(9), 1)
            lDate = Val(Parse(strTemp, "/", 3)) * 10000 _
                + Val(Parse(strTemp, "/", 1)) * 100 _
                + Val(Parse(strTemp, "/", 2))
            If lDate < 200000 Then
                lDate = lDate + 20000000
            ElseIf lDate < 999999 Then
                lDate = lDate + 19000000
            End If
            strTemp = Parse(strLine, Chr(9), 6)
            If Not bExitLine Then
                Trade.lEntryDate = lDate
                Trade.strEntry = strTemp
                strTemp = Parse(strLine, Chr(9), 5)
                iPos = InStr(strTemp, "^")
                If iPos > 0 Then
                    Trade.dEntryPrice = Val(Left(strTemp, iPos - 1)) _
                        + Val(Mid(strTemp, iPos + 1)) / mult
                Else
                    Trade.dEntryPrice = Val(strTemp)
                End If
            Else
                Trade.lExitDate = lDate
                Trade.strExit = strTemp
                strTemp = Parse(strLine, Chr(9), 5)
                iPos = InStr(strTemp, "^")
                If iPos > 0 Then
                    Trade.dExitPrice = Val(Left(strTemp, iPos - 1)) _
                        + Val(Mid(strTemp, iPos + 1)) / mult
                Else
                    Trade.dExitPrice = Val(strTemp)
                End If
                
                Trade.dProfit = Val(Mid(Parse(strLine, Chr(9), 7), 2))
                
                If iTS Mod 100 = 0 Then ReDim Preserve trdTS(iTS + 100) As TradeStruct
                iTS = iTS + 1
                trdTS(iTS) = Trade
            End If
        End If
    Next
    ReDim Preserve trdTS(iTS) As TradeStruct

    ' Compare trades
    lst.Clear
    strTemp = "SysNav:" & Str(iSN) & " trades in " & strSNfile
    lst.AddItem strTemp
    strTemp = "TrdSta:" & Str(iTS) & " trades in " & strTSfile
    lst.AddItem strTemp
    iSN = 1
  If 0 Then
    lst.AddItem FormatTrade("SN", trdSN, 1, rule_len)
    lst.AddItem FormatTrade("TS", trdTS, 1, rule_len)
  Else
    For iTS = 1 To UBound(trdTS)
        Do While iSN <= UBound(trdSN)
            If trdSN(iSN).lEntryDate = trdTS(iTS).lEntryDate Then
                ' same trade
                If trdSN(iSN).lExitDate <> trdTS(iTS).lExitDate _
                        Or trdSN(iSN).bLong <> trdTS(iTS).bLong _
                        Or trdSN(iSN).dEntryPrice <> trdTS(iTS).dEntryPrice _
                        Or trdSN(iSN).dExitPrice <> trdTS(iTS).dExitPrice Then
                    lst.AddItem ""
                    lst.AddItem "Trades not matching ..."
                    lst.AddItem FormatTrade("TS", trdTS, iTS, rule_len)
                    lst.AddItem FormatTrade("SN", trdSN, iSN, rule_len)
                    iNumDiff = iNumDiff + 1
                End If
                iSN = iSN + 1
                Exit Do
            ElseIf trdSN(iSN).lEntryDate > trdTS(iTS).lEntryDate Then
                ' an extra TS trade
                lst.AddItem ""
                lst.AddItem "Extra TS trade ..."
                lst.AddItem FormatTrade("xT", trdTS, iTS, rule_len)
                iNumDiff = iNumDiff + 1
                Exit Do
            End If
            ' an extra SN trade
            lst.AddItem ""
            lst.AddItem "Extra SN trade ..."
            lst.AddItem FormatTrade("xS", trdSN, iSN, rule_len)
            iNumDiff = iNumDiff + 1
            iSN = iSN + 1
        Loop
    Next
    Do While iSN <= UBound(trdSN)
        ' an extra SN trade
        lst.AddItem ""
        lst.AddItem "Extra SN trade ..."
        lst.AddItem FormatTrade("xS", trdSN, iSN, rule_len)
        iNumDiff = iNumDiff + 1
        iSN = iSN + 1
    Loop
  End If
  
  lst.AddItem ""
  lst.AddItem "# of differences:" & Str(iNumDiff)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.TradeDiff", eGDRaiseError_Raise
    
End Sub

#If 0 Then
Public Function FormatNum(ByVal pValue As Variant) As Variant
On Error GoTo ErrSection:

    Dim F       As String
    
    If VarType(pValue) = vbString Then
        pValue = Val(pValue)
    End If
    If Val(pValue) >= 1000 Or Val(pValue) <= -1000 Then
        F = F & "#,###"
    End If
    If InStr(1, pValue, ".") > 0 Then
        F = F & ".00"
    End If
    FormatNum = Format(Val(pValue), F)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.FormatNum", eGDRaiseError_Raise
    
End Function
#End If

Public Sub TestCompiledFunctionCT()
On Error GoTo ErrSection:

    Dim Expr    As cExpression
    Dim Functions   As cFunctions
    Dim mFunction   As cFunction
    Dim X           As Long
    
    Set Functions = New cFunctions
    Functions.Load
    
    Set Expr = New cExpression
    With Expr
        .Functions = g.Functions
        .PortfolioNavigator = False
    End With
    
    For X = 1 To Functions.Count
        Set mFunction = Functions.Item(X)
        If mFunction.ImplementationTypeID = 1 Then
            mFunction.CodedText = Expr.BuiltinCodedText(mFunction.CodedName, _
                      mFunction.DataTypeID, _
                      mFunction.Inputs)
            mFunction.Save
        End If
    Next X
    
ErrExit:
    Set Functions = Nothing
    Set Expr = Nothing
    Exit Sub
    
ErrSection:
    Set Functions = Nothing
    Set Expr = Nothing
    RaiseError "mSysNav.TestCompiledFunctionCT", eGDRaiseError_Raise
    
End Sub

'Converts an input value from display format to a number
Public Function ConvertInputValue(pDefault As Variant, pInputTypeID As Long) As Variant
On Error GoTo ErrSection:

    Select Case pInputTypeID
        Case kSN_RetNumericConstant, kSN_RetTrueFalse, kSN_RetNumeric, kSN_RetTrueFalseConstant
            'If InStr(1, pDefault, ".") > 0 Then
            If ValOfText(pDefault) <> Int(ValOfText(pDefault)) Then
                ConvertInputValue = ValOfText(pDefault)
            Else
                ConvertInputValue = CLng(ValOfText(pDefault))
            End If
        Case Else
            ConvertInputValue = pDefault
    End Select
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.ConvertInputValue", eGDRaiseError_Raise
    
End Function

'Create an Outlook style title in the first row (row 0 of grid)
Public Sub OutLookTitle(pTitle As String, pGrid As VSFlexGrid)
On Error GoTo ErrSection:

    With pGrid
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = pTitle
        .Cell(flexcpBackColor, 0, 0) = &H808080         'Dark Gray
        .Cell(flexcpFontBold, 0, 0) = True
        .Cell(flexcpFontName, 0, 0) = "Times New Roman"
        .Cell(flexcpFontSize, 0, 0) = 12
        .Cell(flexcpForeColor, 0, 0) = vbWhite
        .RowHeight(0) = .CellHeight * 1.5
        .MergeCells = flexMergeRestrictRows
        .MergeRow(0) = True
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.OutLookTitle", eGDRaiseError_Raise
    
End Sub

Public Sub InitFunctions()
On Error GoTo ErrSection:

    Set g.Functions = New cFunctions
    g.Functions.Load
    FilterFunctions
    
    ' set this flag so we know to reload coded
    ' text functions next time we run the engine
    g.bDirtyFunctionLibrary = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.InitFunctions", eGDRaiseError_Raise
    
End Sub

Public Function FunctionFound(pID As Long) As Boolean
On Error GoTo ErrSection:
    
    Dim pFunction       As cFunction
    FunctionFound = True
    Set pFunction = g.Functions.Item(CStr(pID))

ErrExit:
    Exit Function

ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Or Err.Number = 9 Then
        FunctionFound = False
        Resume Next
    Else
        RaiseError "mSysNav.FunctionFound", eGDRaiseError_Raise
    End If

End Function

Public Sub ShowMsg(Optional ByVal lErrNum& = 0, Optional ByVal strSource$ = "", _
                    Optional ByVal strDesc$ = "")
    
    Dim RetVal As Variant
    
    If lErrNum = 0 Then
        lErrNum = Err.Number
        strSource = Err.Source
        strDesc = Err.Description
    End If
    
    RetVal = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
    
    If FormIsLoaded("frmSplash") Then Unload frmSplash

    If lErrNum < 0 Then
        Replace strDesc, vbCrLf, "|"
        InfBox strDesc, , , "Error", , , , , , , , eGDAlign_Left
    Else
        Replace strDesc, vbCrLf, "|"
        InfBox "An unexpected error occurred.||Please report the following: " & _
            "|Source:  " & strSource & _
            "|Message: " & strDesc, , , "Error", , , , , , , , eGDAlign_Left
    End If

End Sub

Public Sub ReSizeMDIChildForm(frm As Form, ctl As Control)
On Error GoTo ErrSection:

    frm.Width = ctl.Left + ctl.Width + frm.Width - frm.ScaleWidth
    frm.Height = ctl.Top + ctl.Height + frm.Height - frm.ScaleHeight

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.ReSizeMDIChildForm", eGDRaiseError_Raise
    
End Sub

Public Function FunctionParmFound(fparms As cInputs, pID As Long) As Boolean
On Error GoTo ErrSection:

    Dim FunctionParm As cInput
    FunctionParmFound = True
    Set FunctionParm = fparms.Item(CStr(pID))

ErrExit:
    Exit Function

ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Or Err.Number = 9 Then
        FunctionParmFound = False
        Resume Next
    Else
        RaiseError "mSysNav.FunctionParmFound", eGDRaiseError_Raise
    End If

End Function

Public Sub HighlightText(pRtf As RichTextBox, pText As String, pStart As Integer, pBold As Boolean, _
        pColor As Long, pItalics As Boolean)
On Error GoTo ErrSection:

    With pRtf
        .Find pText, pStart
        .SelLength = Len(pText)
        .SelBold = pBold
        .SelColor = pColor
        .SelItalic = pItalics
        .SelLength = 0
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.HighlightText", eGDRaiseError_Raise
    
End Sub

' Asks the user if they want to rename or copy when the name has changed
Public Function RenameOrCopy(strType$) As String
On Error GoTo ErrSection:

    Dim strMsg$
    
    strMsg = "Name of the " & strType & " has changed.||Do you wish to rename the existing " & strType & ",| or create a copy with the new name?"
    RenameOrCopy = AskBox("i=? ; b=+Copy|-Rename ; h=Copy or Rename ; " & strMsg)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.RenameOrCopy", eGDRaiseError_Raise
    
End Function

'Function type Description
Public Function ImplementationTypeDesc(pID As Byte) As String
On Error GoTo ErrSection:

    If pID = kSN_BuiltIn Then
        ImplementationTypeDesc = "Built-in"
    Else
        ImplementationTypeDesc = "Custom"
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.ImplementationTypeDesc", eGDRaiseError_Raise
    
End Function

'Security Description
Public Function SecurityDesc(pID As Variant) As String
On Error GoTo ErrSection:

    Select Case pID
        Case 0: SecurityDesc = "Can Edit/Can View"
        Case 1: SecurityDesc = "No Edit/Can View"
        Case 2: SecurityDesc = "No Edit/No View"
        Case 3: SecurityDesc = "No Access"
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.SecurityDesc", eGDRaiseError_Raise
    
End Function

'Unhide all columns in a grid
Public Sub HideCols(pGrid As VSFlexGrid, pCols As Integer, pHide As Boolean)
On Error GoTo ErrSection:

    Dim X   As Integer
    
    For X = 0 To pCols - 1
        pGrid.ColHidden(X) = pHide
    Next X

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.HideCols", eGDRaiseError_Raise
    
End Sub

'Mass change (insert blank line into existing PreviewRTF) for all rules
Public Sub InsertPar()
On Error GoTo ErrSection:

    Dim cRule   As cRule
    Dim cRules  As cRules
    Dim wrkText As String
    Dim newText As String
    Dim X       As Long
    Dim pos     As Long
    Dim WrkJet  As Workspace
    Dim dbNav   As Database
    Dim cb      As cCommonBridge
    Dim RuleID  As Long
    
    Set WrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbNav = WrkJet.OpenDatabase(App.Path & "\Libraries.MDB", False)
    
    'Pass connection to Access throught bridges...
    Set cb = New cCommonBridge
    cb.dbNavRef = dbNav
    
    Set cRules = New cRules
    cRules.Load
    For X = 1 To cRules.Count
        Set cRule = New cRule
        cRule.RuleID = cRules.Item(X).RuleID
        cRule.Load
        wrkText = cRule.CondFillWords
        pos = InStr(1, wrkText, "THEN")
        
        If Mid(wrkText, pos - 5, 4) <> "\par" Then
            newText = Left(wrkText, pos - 1) & _
                " \par " & _
                Right(wrkText, Len(wrkText) - pos + 1)
                
            cRule.CondFillWords = newText
            cRule.Save
        End If
    Next X
    
ErrExit:
    Set cRules = Nothing
    Set cRule = Nothing
    Exit Sub
    
ErrSection:
    Set cRules = Nothing
    Set cRule = Nothing
    RaiseError "mSysNav.InsertPar", eGDRaiseError_Raise
    
End Sub

Public Function GetRuleType(ByVal hTable As Long, ByVal lIndex As Long, Optional ByVal bForcedExit As Boolean = False) As String
On Error GoTo ErrSection:

    Dim lType As Long
    Dim bBuySell As Boolean
    
    lType = gdGetTableNum(hTable, RuleField(etblRule_RuleType), lIndex)
    bBuySell = CBool(gdGetTableNum(hTable, RuleField(etblRule_BuySell), lIndex))

    'If (rs!LateCondition = True) Or (rs!LateAction = True) Or bForcedExit Then
    If lType = 1 Or bForcedExit = True Then
        If bBuySell Then
            GetRuleType = "Short Exit"
        Else
            GetRuleType = "Long Exit"
        End If
    Else
        If bBuySell Then
            GetRuleType = "Long Entry"
        Else
            GetRuleType = "Short Entry"
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.GetRuleType", eGDRaiseError_Raise
    
End Function

'This routine builds a gdTable out of tblSettings and tblSettingsOptions
'and saves the table in the application root as "gdSettings.dat" and "gdSettingsO.dat".
'These tables are loaded in this class at startup and are used for loading
'the settings grid appropriate.
'Any time a change is done to report options or settings on a maintenance form,
'This method must be called by the developer.  This file must then be
'distributed to the end user.
Public Sub BuildSettingsFile()
On Error GoTo ErrSection:

    Dim RetVal          As Variant
    Dim X               As Integer
    Dim Settings        As Recordset
    Dim gds             As cGdTable
    
    ChangePath App.Path
    
    Set gds = New cGdTable
    RetVal = gds.CreateField(eGDARRAY_Longs, 0, "AppID")
    RetVal = gds.CreateField(eGDARRAY_Longs, 1, "SettingID")
    RetVal = gds.CreateField(eGDARRAY_Strings, 2, "SettingName")
    RetVal = gds.CreateField(eGDARRAY_Strings, 3, "LabelName")
    RetVal = gds.CreateField(eGDARRAY_Strings, 4, "Desc")
    RetVal = gds.CreateField(eGDARRAY_Strings, 5, "Type")
    RetVal = gds.CreateField(eGDARRAY_Longs, 6, "DecPos")
    RetVal = gds.CreateField(eGDARRAY_Strings, 7, "Formatting")
    RetVal = gds.CreateField(eGDARRAY_Doubles, 8, "ValFrom")
    RetVal = gds.CreateField(eGDARRAY_Doubles, 9, "ValTo")
    RetVal = gds.CreateField(eGDARRAY_Longs, 10, "Length")
    RetVal = gds.CreateField(eGDARRAY_TinyInts, 11, "CanEdit")
    RetVal = gds.CreateField(eGDARRAY_TinyInts, 12, "Required")
    RetVal = gds.CreateField(eGDARRAY_TinyInts, 13, "ShowEdit")
    RetVal = gds.CreateField(eGDARRAY_TinyInts, 14, "ShowAdd")
    RetVal = gds.CreateField(eGDARRAY_Strings, 15, "Default")
    RetVal = gds.CreateField(eGDARRAY_Strings, 16, "Group")
    RetVal = gds.CreateField(eGDARRAY_TinyInts, 17, "Global")
    RetVal = gds.CreateField(eGDARRAY_TinyInts, 18, "Save")
    
    'Load "Settings Options" table
    Set Settings = g.dbNav.OpenRecordset("Select * from [tblSettings] " & _
        "Order by [AppID],[SettingID];", dbOpenSnapshot)
    X = 0
    Do Until Settings.EOF
        X = X + 1
        gds.Item(0, X) = Settings!AppID
        gds.Item(1, X) = Settings!SettingID
        gds.Item(2, X) = Settings!SettingName
        gds.Item(3, X) = Settings!LabelName
        gds.Item(4, X) = Settings!Desc
        gds.Item(5, X) = Settings!Type
        gds.Item(6, X) = Settings!DecPos
        gds.Item(7, X) = NullChk(Settings!Formatting, "")
        gds.Item(8, X) = Settings!ValFrom
        gds.Item(9, X) = Settings!ValTo
        gds.Item(10, X) = Settings!Length
        gds.Item(11, X) = Abs(Settings!CanEdit)
        gds.Item(12, X) = Abs(Settings!Required)
        gds.Item(13, X) = Abs(Settings!ShowEdit)
        gds.Item(14, X) = Abs(Settings!ShowAdd)
        gds.Item(15, X) = NullChk(Settings!Default, 0)
        gds.Item(16, X) = NullChk(Settings!Group, "")
        gds.Item(17, X) = Abs(NullChk(Settings!Global, 0))
        gds.Item(18, X) = Abs(NullChk(Settings!Save, 0))
        Settings.MoveNext
    Loop
    
    'Save to disk
    KillFile AddSlash(App.Path) & "gdSettings.dat"
    gds.Serialize AddSlash(App.Path) & "gdSettings.dat", True
    
ErrExit:
    Set Settings = Nothing
    Set gds = Nothing
    Exit Sub
    
ErrSection:
    Set Settings = Nothing
    Set gds = Nothing
    RaiseError "mSysNav.BuildSettingsFile", eGDRaiseError_Raise
    
End Sub

Public Function SystemIDForName(ByVal strSystemName As String) As Long
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
                "WHERE [SystemName]='" & strSystemName & "';", dbOpenDynaset)
    ValidateCheckSums rs, "tblSystems"
    
    If Not rs.EOF Then
        If rs!CheckSum <> 0.5 Then SystemIDForName = rs!SystemNumber
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.SystemIDForName", eGDRaiseError_Raise
    
End Function

Public Function SystemNameForID(ByVal lSystemID As Long) As String
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
                "WHERE [SystemNumber]=" & lSystemID & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblSystems"
    
    If Not rs.EOF Then
        If rs!CheckSum <> 0.5 Then SystemNameForID = rs!SystemName
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.SystemNameForID", eGDRaiseError_Raise
    
End Function

Public Function BasketIDForName(ByVal strBasketName As String) As Long
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] " & _
                "WHERE [Name]='" & strBasketName & "';", dbOpenDynaset)
    ValidateCheckSums rs, "tblStrategyBaskets"
    
    If Not rs.EOF Then
        If rs!CheckSum <> 0.5 Then BasketIDForName = rs!StrategyBasketID
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.BasketIDForName", eGDRaiseError_Raise
    
End Function

Public Function BasketNameForID(ByVal lBasketID As Long) As String
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] " & _
                "WHERE [StrategyBasketID]=" & lBasketID & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblStrategyBaskets"
    
    If Not rs.EOF Then
        If rs!CheckSum <> 0.5 Then BasketNameForID = rs!Name
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.BasketNameForID", eGDRaiseError_Raise
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateEditor
'' Description: Try to activate a form of a certian type and certain ID
'' Inputs:      Form Name, ID loaded on Form
'' Returns:     True if Activated, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ActivateEditor(ByVal strFormName As String, ByVal strID As String, _
            Optional frmReturn As Form) As Boolean
On Error GoTo ErrSection:

    Dim i&
    Dim frm As Form
    
    DoEvents '(to give a chance for editors to unload after toolbox is hidden)
    
    ' step backwards so unloading won't mess up our index
    For i = Forms.Count - 1 To 0 Step -1
        Set frm = Forms(i)
        If frm.Name = strFormName Then
            ' first check if not unloaded all the way
            If Not frm.Visible Then
                Unload frm
            ElseIf frm.ID = strID Then
                ' if already loaded and visible, just show again
                frm.SetFocus
                If frm.WindowState = vbMinimized Then frm.WindowState = vbNormal
                If Not IsMissing(frmReturn) Then Set frmReturn = frm
                ActivateEditor = True
                Exit For
            End If
        End If
    Next

ErrExit:
    Set frm = Nothing
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.ActivateEditor", eGDRaiseError_Raise
    
End Function

Public Sub LoadLibrariesTable()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lProcAddr As Long               ' Callback function address
    Dim strDLL As String
    Dim lPos As Long

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] WHERE [Ignore]=0 ORDER BY [LibraryID];", dbOpenDynaset)
    ValidateCheckSums rs, "tblLibrarys"
    
    Set g.tblLibrary = New cGdTable
    With g.tblLibrary
        .CreateField eGDARRAY_Longs, LibraryField(etblLib_ID)
        .CreateField eGDARRAY_Strings, LibraryField(etblLib_Name)
        .CreateField eGDARRAY_Longs, LibraryField(etblLib_Type)
        .CreateField eGDARRAY_Strings, LibraryField(etblLib_Path)
        .CreateField eGDARRAY_Longs, LibraryField(etblLib_VBProcAddr)
        .CreateField eGDARRAY_Doubles, LibraryField(etblLib_LastModified)
        
        Do While Not rs.EOF
            If rs!CheckSum <> 0.5 Then
                If rs!LibraryType = 1 Then
                    lProcAddr = FunctionPtrToLong(AddressOf RunVbFunctionCallback)
                Else
                    lProcAddr = 0&
                End If
                
                strDLL = Trim(NullChk(rs!Path))
                If Len(strDLL) > 0 Then
                    If UCase(strDLL) <> "BUILTIN.DLL" And Not rs!BuiltIn Then
                        ' 3rd-party DLL's should be in the "LibraryDLLs" folder
                        strDLL = App.Path
                        lPos = At(strDLL, "\", -1)
                        If lPos > 0 Then
                            strDLL = Left(strDLL, lPos) & "LibraryDLLs\" & Trim(rs!Path)
                        Else
                            strDLL = ""
                        End If
                    End If
                End If
                
                .AddRecord Str(rs!LibraryID) & vbTab & rs!LibraryName & vbTab & rs!LibraryType & _
                        vbTab & strDLL & vbTab & Str(lProcAddr) & vbTab & Str(rs!LastModified)
            End If
            rs.MoveNext
        Loop
    End With

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.LoadLibrariesTable", eGDRaiseError_Raise
    
End Sub

Public Sub LoadFunctionsTable()
On Error GoTo ErrSection:

    Dim iReturnType As Integer
    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] ORDER BY [FunctionID];", dbOpenDynaset)
    ValidateCheckSums rs, "tblFunctions"
    
    Set g.tblFunction = New cGdTable
    With g.tblFunction
        .CreateField eGDARRAY_Longs, FunctionField(etblFunction_ID)
        .CreateField eGDARRAY_Longs, FunctionField(etblFunction_LibID)
        .CreateField eGDARRAY_Strings, FunctionField(etblFunction_Name)
        .CreateField eGDARRAY_Strings, FunctionField(etblFunction_NameCoded)
        .CreateField eGDARRAY_Longs, FunctionField(etblFunction_Implementation)
        .CreateField eGDARRAY_Strings, FunctionField(etblFunction_CodedText)
        .CreateField eGDARRAY_Longs, FunctionField(etblFunction_ReturnType)
        .CreateField eGDARRAY_TinyInts, FunctionField(etblFunction_LateCalculating), , 0
        .CreateField eGDARRAY_TinyInts, FunctionField(etblFunction_UsesOpenNextBar), , 0
        .CreateField eGDARRAY_TinyInts, FunctionField(etblFunction_UsesHLCNextBar), , 0
        .CreateField eGDARRAY_Longs, FunctionField(etblFunction_CategoryID)
        .CreateField eGDARRAY_Doubles, FunctionField(etblFunction_LastModified)
        .CreateField eGDARRAY_TinyInts, FunctionField(etblFunction_Usage)
        .CreateField eGDARRAY_Strings, FunctionField(etblFunction_TradeSenseUsage)
        .CreateField eGDARRAY_Strings, FunctionField(etblFunction_Description)
        .CreateField eGDARRAY_Longs, FunctionField(etblFunction_SecurityLevel)
        .CreateField eGDARRAY_Strings, FunctionField(etblFunction_Password)
        .CreateField eGDARRAY_TinyInts, FunctionField(etblFunction_CannotDelete), , 0
        .CreateField eGDARRAY_TinyInts, FunctionField(etblFunction_Reverify)
        
        Do While Not rs.EOF
            If rs!CheckSum <> 0.5 Then
                iReturnType = rs!ReturnTypeID
#If 0 Then '(guess we don't really need to do this after all!)
                ' change return type of engine functions to "arrays"
                If rs!ImplementationTypeID = 3 Then '(an internal engine function)
                    Select Case iReturnType
                    Case 1 'constant number
                        iReturnType = 4 'array of numbers
                    Case 2 'constant string
                        iReturnType = 8 'array of strings
                    Case 6 'constant boolean
                        iReturnType = 3 'array of booleans
                    End Select
                End If
#End If

                If HasModule(NullChk(rs!RequiredMod)) Or IsIDE Then
                    .AddRecord Str(rs!FunctionID) & vbTab & Str(rs!LibraryID) & vbTab & rs!FunctionName & _
                            vbTab & rs!CodedName & vbTab & Str(rs!ImplementationTypeID) & _
                            vbTab & DecryptField(rs!CodedText) & vbTab & Str(iReturnType) & _
                            vbTab & CLng(rs!LateCalculating) & vbTab & CLng(rs!UsesOpenNextBar) & _
                            vbTab & CLng(rs!UsesHLCNextBar) & vbTab & CLng(rs!FunctionCategoryID) & _
                            vbTab & Str(CDbl(rs!LastModified)) & vbTab & rs!Usage & _
                            vbTab & rs!TradeSenseUsage & vbTab & rs!Description & _
                            vbTab & CLng(rs!SecurityLevel) & vbTab & DecryptField(rs!Password) & _
                            vbTab & CLng(rs!CannotDelete) & vbTab & CLng(rs!Reverify)
                End If
            End If
            
            rs.MoveNext
        Loop
    End With

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.LoadFunctionsTable", eGDRaiseError_Raise
    
End Sub

Public Sub LoadFunctionParmsTable()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionParms] ORDER BY [FunctionID], [ParmNbr];", dbOpenDynaset)
    
    Set g.tblFunctionParm = New cGdTable
    With g.tblFunctionParm
        .CreateField eGDARRAY_Longs, FuncParmField(etblFunctionParm_ID)
        .CreateField eGDARRAY_Longs, FuncParmField(etblFunctionParm_FunctionID)
        .CreateField eGDARRAY_TinyInts, FuncParmField(etblFunctionParm_Sequence)
        .CreateField eGDARRAY_Strings, FuncParmField(etblFunctionParm_Name)
        .CreateField eGDARRAY_Longs, FuncParmField(etblFunctionParm_Type)
        .CreateField eGDARRAY_Doubles, FuncParmField(etblFunctionParm_FromValue)
        .CreateField eGDARRAY_Doubles, FuncParmField(etblFunctionParm_ToValue)
        .CreateField eGDARRAY_TinyInts, FuncParmField(etblFunctionParm_Required)
        .CreateField eGDARRAY_Strings, FuncParmField(etblFunctionParm_Default)
        
        Do While Not rs.EOF
            .AddRecord Str(rs!ParmID) & vbTab & Str(rs!FunctionID) & vbTab & rs!ParmNbr & _
                    vbTab & rs!ParmText & vbTab & Str(rs!ParmTypeID) & vbTab & _
                    Str(rs!FromValue) & vbTab & Str(rs!ToValue) & vbTab & CLng(rs!Required) & _
                    vbTab & rs!DefaultValue
            rs.MoveNext
        Loop
    End With

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.LoadFunctionParmsTable", eGDRaiseError_Raise
    
End Sub

Private Function LoadEngineFunctionSet(ByVal strFunctionSetName) As Long
On Error GoTo ErrSection:

    Dim i&, strText$, strSave$
    Dim hstrFunctionSetName As Long
    
    hstrFunctionSetName = gdCreateString(0)
    gdSetStr hstrFunctionSetName, 0, strFunctionSetName
    
    If HasNewNavEngine Then
#If 0 Then ' TLB 4/18/2011: NOW OBSOLETE
        ' TLB: for now, have the new NavEngineAdv.DLL make calls into BuiltInAdv.DLL,
        ' while the old NavEngine.DLL will still make calls into BuiltIn.DLL
        For i = 0 To g.tblLibrary.NumRecords - 1
            Select Case UCase(FileBase(g.tblLibrary(3, i)))
            Case "BUILTIN"
                ''g.tblLibrary(3, i) = "BuiltInAdv.DLL"
            Case "OCEANTHEORY"
                strText = FilePath(g.tblLibrary(3, i)) & "OceanTheoryAdv.DLL"
                If FileExist(strText) Then
                    g.tblLibrary(3, i) = strText
                    strSave = "OCEAN"
                End If
            End Select
        Next
        If strSave = "OCEAN" Then
            strSave = ""
            For i = g.tblFunction.NumRecords - 1 To 0 Step -1
                strText = g.tblFunction(3, i)
                If strText = "STX2" Then
                    'strText = "NMA_DLL (STX_Prices (Smooth) , NMM_MaxLB (STX_Prices (Smooth) , 50 , 15) , True , Round (TEMA (STX2_DLL (NST_DLL (0.5 , NMM_MaxLB (STX_Prices (Smooth) , 50 , 15)) , NMM_MaxLB (Close , 50 , 15) , Automatic , Dial) , Smooth , True) , 2))"
                    'strSave = "~01007NMA_DLL ~16001( ~01010STX_Prices ~16001( ~07007Market1 ~22001, ~05006Smooth ~17001) ~22001, ~01009NMM_MaxLB ~16001( ~01010STX_Prices ~16001( ~07007Market1 ~22001, ~05006Smooth ~17001) ~22001, ~1300250 ~22001, ~1300215 ~17001) ~22001, ~03004True ~16001( ~17001) ~22001, ~01005Round ~16001( ~01004TEMA ~16001( ~01008STX2_DLL ~16001( ~01007NST_DLL ~16001( ~07007Market1 ~22001, ~130030.5 ~22001, ~01009NMM_MaxLB ~16001( ~01010STX_Prices ~16001( ~07007Market1 ~22001, ~05006Smooth ~17001) ~22001, ~1300250 ~22001, ~1300215 ~17001) ~17001) ~22001, ~01009NMM_MaxLB ~16001( ~01005Close ~16001( ~07007Market1 ~17001) ~22001, ~1300250 ~22001, ~1300215 ~17001) ~22001, ~06009Automatic ~22001, ~05004Dial ~17001) ~22001, ~05006Smooth ~22001, ~03004True ~16001( ~17001) ~17001) ~22001, ~130012 ~17001) ~17001)"
                    strSave = g.tblFunction(5, i) ' save the STX2 coded text
                ElseIf strText = "STX" Then
                    If Len(strSave) > 0 Then ' replace STX coded text with the STX2
                        strText = g.tblFunction(5, i)
                        g.tblFunction(5, i) = strSave
                        strSave = strText
                    End If
                    Exit For
                End If
            Next
        End If
#End If

        LoadEngineFunctionSet = LoadFunctionSetAdv(hstrFunctionSetName, g.tblLibrary.TableHandle, _
            g.tblFunction.TableHandle, g.tblFunctionParm.TableHandle)
            
#If 0 Then ' TLB 4/18/2011: NOW OBSOLETE
        If Len(strSave) > 0 Then
            For i = g.tblFunction.NumRecords - 1 To 0 Step -1
                If g.tblFunction(3, i) = "STX" Then
                    g.tblFunction(5, i) = strSave
                    Exit For
                End If
            Next
        End If
        For i = 0 To g.tblLibrary.NumRecords - 1
            Select Case UCase(FileBase(g.tblLibrary(3, i)))
            Case "BUILTINADV"
                g.tblLibrary(3, i) = "BuiltIn.DLL"
            Case "OCEANTHEORYADV"
                g.tblLibrary(3, i) = FilePath(g.tblLibrary(3, i)) & "OceanTheory.DLL"
            End Select
        Next
#End If
    End If
    
    'If FileLength(App.Path & "\Provided\SystemAdv.flg") < 3 Then
        LoadEngineFunctionSet = LoadFunctionSet(hstrFunctionSetName, g.tblLibrary.TableHandle, _
            g.tblFunction.TableHandle, g.tblFunctionParm.TableHandle)
    'End If
            
ErrExit:
    gdDestroyString hstrFunctionSetName
    Exit Function
    
ErrSection:
    gdDestroyString hstrFunctionSetName
    RaiseError "mSysNav.LoadEngineFunctionSet"
End Function

Public Function LoadEngineFunctions(Optional ByVal strFunctionSetName As String = "") As Long
On Error GoTo ErrSection:

    LoadLibrariesTable
    LoadFunctionsTable
    LoadFunctionParmsTable
    LoadFunctionCategories
    LoadRulesTable
    
    LoadEngineFunctions = LoadEngineFunctionSet(strFunctionSetName)
            
    UpdateVisibleCharts eRedo6_ReloadInd

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.LoadEngineFunctions"
End Function

Public Function UnloadEngineFunctions(Optional ByVal strFunctionSetName As String = "") As Long
On Error GoTo ErrSection:

    Dim hstrFunctionSetName As Long
    Dim strFile$
    Dim mb As New cMemBuffer, mbKey As New cMemBuffer
    
    If g.bUnloading And Not g.tblRule Is Nothing Then
        If g.tblRule.NumRecords > 0 Then
            ' serialize and encrypt the rules table (for faster loading)
            strFile = App.Path & "\Rules.tbl"
            g.tblRule.Serialize strFile, True
            mb.FromFile strFile
            mbKey.Buffer = "Key-for-Rules.tbl"
            gdEncrypt True, mb, mbKey
            KillFile strFile
            strFile = App.Path & "\RulesEnc.tbl"
            mb.ToFile strFile
        End If
    End If
    
    hstrFunctionSetName = gdCreateString(0)
    gdSetStr hstrFunctionSetName, 0, strFunctionSetName
    
    If HasNewNavEngine Then
        UnloadEngineFunctions = UnloadFunctionSetAdv(hstrFunctionSetName)
    End If
    UnloadEngineFunctions = UnloadFunctionSet(hstrFunctionSetName)
    
    Set g.tblFunction = Nothing
    Set g.tblFunctionParm = Nothing
    Set g.tblLibrary = Nothing
    Set g.astrFunctionCategory = Nothing
    Set g.tblRule = Nothing

ErrExit:
    gdDestroyString hstrFunctionSetName
    Exit Function
    
ErrSection:
    gdDestroyString hstrFunctionSetName
    RaiseError "mSysNav.UnloadEngineFunctions", eGDRaiseError_Raise
    
End Function

Public Sub RefreshFunction(F As cFunction, Optional ByVal strFunctionSetName As String = "", Optional ByVal bReloadEngineFunctions As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Proper Index into the table
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim bFound As Boolean               ' Whether the function was found
    
    bFound = g.tblFunction.FieldArray(FunctionField(etblFunction_ID), False).BinarySearch(F.FunctionID, lIndex)
    If bFound Then
        g.tblFunction.RemoveRecords lIndex
        
        bFound = g.tblFunctionParm.FieldArray(FuncParmField(etblFunctionParm_FunctionID), False).BinarySearch(F.FunctionID, lIndex2)
        Do While g.tblFunctionParm(FuncParmField(etblFunctionParm_FunctionID), lIndex2) = F.FunctionID
            g.tblFunctionParm.RemoveRecords lIndex2
        Loop
    End If
    
    If HasModule(F.RequiredMod) Or IsIDE Then
        g.tblFunction.AddRecord Str(F.FunctionID) & vbTab & Str(F.LibraryID) & vbTab & F.FunctionName & _
                vbTab & F.CodedName & vbTab & Str(F.ImplementationTypeID) & vbTab & _
                F.CodedText & vbTab & Str(F.ReturnTypeID) & vbTab & CLng(F.LateCalculating) & _
                vbTab & CLng(F.UsesOpenNextBar) & vbTab & CLng(F.UsesNextBarHLC) & _
                vbTab & Str(F.FunctionCategoryID) & vbTab & Str(CDbl(F.LastModified)) & vbTab & _
                F.Usage & vbTab & F.TradeSenseUsage & vbTab & F.Description & vbTab & _
                F.SecurityLevel & vbTab & F.Password & vbTab & CLng(F.CannotDelete) & vbTab & _
                CLng(F.Reverify), lIndex, vbTab
        
        For lIndex2 = 1 To F.Inputs.Count
            RefreshFunctionParm F.FunctionID, F.Inputs.Item(lIndex2)
        Next lIndex2
    End If
    
    If bReloadEngineFunctions Then
        LoadEngineFunctionSet strFunctionSetName
    End If
    
    If HasModule(F.RequiredMod) Or IsIDE Then
        If g.Functions.Found(CStr(F.FunctionID)) Then
            g.Functions.ReloadFunction CStr(F.FunctionID)
        Else
            g.Functions.Add F, CStr(F.FunctionID)
        End If
    ElseIf g.Functions.Found(CStr(F.FunctionID)) Then
        g.Functions.Remove CStr(F.FunctionID)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.RefreshFunction"
End Sub

Public Sub DeleteFunction(ByVal lFunctionID As Long, Optional ByVal strFunctionSetName As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Proper Index into the table
    Dim bFound As Boolean               ' Whether the function was found

    bFound = g.tblFunction.FieldArray(FunctionField(etblFunction_ID), False).BinarySearch(lFunctionID, lIndex)
    If bFound Then
        g.tblFunction.RemoveRecords lIndex
        
        bFound = g.tblFunctionParm.FieldArray(FuncParmField(etblFunctionParm_FunctionID), False).BinarySearch(lFunctionID, lIndex)
        Do While g.tblFunctionParm(FuncParmField(etblFunctionParm_FunctionID), lIndex) = lFunctionID
            g.tblFunctionParm.RemoveRecords lIndex
        Loop
    End If
    
    LoadEngineFunctionSet strFunctionSetName
            
    g.Functions.Remove CStr(lFunctionID)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mSysNav.DeleteFunction"
End Sub

Public Sub DeleteLibrary(ByVal lLibraryID As Long, Optional ByVal strFunctionSetName As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Proper Index into the table
    Dim bFound As Boolean               ' Whether the function was found
    
    bFound = g.tblLibrary.FieldArray(LibraryField(etblLib_ID), False).BinarySearch(lLibraryID, lIndex)
    If bFound Then
        g.tblLibrary.RemoveRecords lIndex
        
        ' Remove any function from the table that were in this library...
        For lIndex = g.tblFunction.NumRecords - 1 To 0 Step -1
            If g.tblFunction.Num(FunctionField(etblFunction_LibID), lIndex) = lLibraryID Then
                g.tblFunction.RemoveRecords lIndex
            End If
        Next lIndex
        
        ' Remove any rules from the table that were in this library...
        For lIndex = g.tblRule.NumRecords - 1 To 0 Step -1
            If g.tblRule.Num(RuleField(etblRule_LibraryID), lIndex) = lLibraryID Then
                g.tblRule.RemoveRecords lIndex
            End If
        Next lIndex
        
        LoadEngineFunctionSet strFunctionSetName
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mSysNav.DeleteLibrary"
End Sub

Public Sub RefreshFunctionParm(ByVal lFunctionID As Long, Parm As cInput)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Proper Index into the table
    Dim bFound As Boolean               ' Whether the function was found
    
    bFound = g.tblFunctionParm.FieldArray(FuncParmField(etblFunctionParm_FunctionID), False).BinarySearch(lFunctionID, lIndex)
    Do While g.tblFunctionParm(FuncParmField(etblFunctionParm_ID), lIndex) <> Parm.ParmID
        If g.tblFunctionParm(FuncParmField(etblFunctionParm_FunctionID), lIndex) <> lFunctionID Then
            bFound = False
            Exit Do
        End If
        lIndex = lIndex + 1
    Loop
    
    If bFound Then
        g.tblFunctionParm.SetRecord Str(Parm.ParmID) & vbTab & Str(lFunctionID) & vbTab & _
                Parm.ParmSeq & vbTab & Parm.ParmName & vbTab & Str(Parm.ParmTypeID) & _
                vbTab & Str(Parm.FromValue) & vbTab & Str(Parm.ToValue) & vbTab & Parm.Required & _
                vbTab & Parm.DefaultValue, lIndex, vbTab
    Else
        g.tblFunctionParm.AddRecord Str(Parm.ParmID) & vbTab & Str(lFunctionID) & vbTab & _
                Parm.ParmSeq & vbTab & Parm.ParmName & vbTab & Str(Parm.ParmTypeID) & _
                vbTab & Str(Parm.FromValue) & vbTab & Str(Parm.ToValue) & vbTab & Parm.Required & _
                vbTab & Parm.DefaultValue, lIndex, vbTab
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.RefreshFunctionParm", eGDRaiseError_Raise
    
End Sub

Public Function DefaultRuleNumber(ByVal strRuleName As String) As Long
On Error GoTo ErrSection:

    Dim strTemp As String
    
    strTemp = Parse(strRuleName, " ", 3)
    strTemp = Right(strTemp, Len(strTemp) - 1)
    
    DefaultRuleNumber = CLng(Val(strTemp))

ErrExit:
    Exit Function
    
ErrSection:
    If Err.Number = 13 Or Err.Number = 6 Then
        DefaultRuleNumber = 0
        Resume ErrExit
    End If
    RaiseError "mSysNav.DefaultRuleNumber", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsDefaultRuleName
'' Description: Figure out if this is one of our made up default rule names
'' Inputs:      Rule Name
'' Returns:     True if default rule name, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsDefaultRuleName(ByVal strRuleName As String) As Boolean
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray
    
    astrFields.Create eGDARRAY_Strings
    astrFields.SplitFields strRuleName, " "
    
    If astrFields.Size = 3 Then
        If UCase(astrFields(0)) = "LONG" Or UCase(astrFields(0)) = "SHORT" Then
            If UCase(astrFields(1)) = "ENTRY" Or UCase(astrFields(1)) = "EXIT" Then
                If Len(astrFields(2)) = 4 And Left(astrFields(2), 1) = "#" Then
                    If IsNumeric(Right(astrFields(2), Len(astrFields(2)) - 1)) Then
                        IsDefaultRuleName = True
                    End If
                End If
            End If
        End If
    End If
    
ErrExit:
    Set astrFields = Nothing
    Exit Function
    
ErrSection:
    Set astrFields = Nothing
    RaiseError "mSysNav.IsDefaultRuleName", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SyncSystemInfo
'' Description: Get the correct Name and ID pair for a System
'' Inputs:      System Name, System ID
'' Returns:     True if the System was found in some way, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SyncSystemInfo(strSystemName As String, lSystemID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lTempID As Long
    Dim strTempName As String
    
    ' Try to get the ID for the given name and the Name for the given ID...
    SyncSystemInfo = True
    lTempID = SystemIDForName(strSystemName)
    If lTempID > 0 Then
        lSystemID = lTempID
    Else
        strTempName = Trim(SystemNameForID(lSystemID))
        If Len(strTempName) <> 0 Then
            strSystemName = strTempName
        Else
            SyncSystemInfo = False
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.SyncSystemInfo", eGDRaiseError_Raise
    
End Function

Public Function IsEditor(ByVal strFormName As String) As Boolean
On Error GoTo ErrSection:

    Dim strEditorNames As String
    
    strEditorNames = ",frmSymbolGroup,frmCriteria,frmFilter,frmFunctionMgr," & _
                        "frmFunctionMgrCT,frmRule,frmSystemManager,frmStrategyBasket,"
    
    If InStr(UCase(strEditorNames), "," & UCase(strFormName) & ",") > 0 Then
        IsEditor = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IsEditor", eGDRaiseError_Raise
    
End Function

Public Sub LoadFunctionCategories()
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionCategories];", dbOpenDynaset)
    
    Set g.astrFunctionCategory = New cGdArray
    g.astrFunctionCategory.Create eGDARRAY_Strings
    
    Do While Not rs.EOF
        g.astrFunctionCategory(rs!FunctionCategoryID) = rs!FunctionCategory
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.LoadFunctionCategories", eGDRaiseError_Raise
    
End Sub

Public Sub LoadRulesTable()
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim bAdd As Boolean, bLoaded As Boolean
    Dim aSystems As New cGdArray
    Dim lIndex&, lSystemID&, hRules&
    Dim strSystem$, strFile$, strPreview$
    Dim mb As New cMemBuffer, mbKey As New cMemBuffer
    
    Set g.tblRule = New cGdTable
    
    ' when starting, load from serialized rules table if it exists
    strFile = App.Path & "\RulesEnc.tbl" ' (encrypted file)
    If g.bStarting And FileExist(strFile) Then
        If mb.FromFile(strFile) Then
            mbKey.Buffer = "Key-for-Rules.tbl"
            gdEncrypt False, mb, mbKey
            strFile = App.Path & "\Rules.tbl" ' (unencrypted file)
            mb.ToFile strFile
            If g.tblRule.Serialize(strFile, False) Then
                bLoaded = True
            End If
        End If
    End If
    KillRulesFile
    
    If Not bLoaded Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] ORDER BY [RuleID];", dbOpenDynaset)
        ValidateCheckSums rs, "tblRules"
        
        With g.tblRule
            .CreateField eGDARRAY_Longs, RuleField(etblRule_RuleID)
            .CreateField eGDARRAY_Strings, RuleField(etblRule_RuleName)
            .CreateField eGDARRAY_Longs, RuleField(etblRule_RuleType)
            .CreateField eGDARRAY_TinyInts, RuleField(etblRule_BuySell)
            .CreateField eGDARRAY_Longs, RuleField(etblRule_LibraryID)
            .CreateField eGDARRAY_Doubles, RuleField(etblRule_LastModified)
            .CreateField eGDARRAY_Strings, RuleField(etblRule_PreviewRTF)
            .CreateField eGDARRAY_Longs, RuleField(etblRule_SecurityLevel)
            .CreateField eGDARRAY_Strings, RuleField(etblRule_Password)
            .CreateField eGDARRAY_TinyInts, RuleField(etblRule_CannotDelete)
            .CreateField eGDARRAY_Longs, RuleField(etblRule_SystemNumber)
            .CreateField eGDARRAY_TinyInts, RuleField(etblRule_Reverify)
            .CreateField eGDARRAY_Longs, RuleField(etblRule_CategoryID)
            
            Do While Not rs.EOF
                If rs!CheckSum <> 0.5 Then
                    bAdd = True
                    
                    ' Delete any local rules that may be left over from a system being
                    ' deleted (DAJ: 03/24/2003)...
                    If rs!SystemNumber <> 0 Then
                        Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] WHERE [SystemNumber]=" & rs!SystemNumber & ";", dbOpenDynaset)
                        If rs2.EOF And rs2.BOF Then
                            bAdd = False
                            rs.Delete
                        Else
                            ' If the local rule is either in a different library than the
                            ' system or it has a different security level or password than
                            ' the system, then set the local rule information to the system
                            ' information (DAJ: 03/24/2003)...
                            If rs!LibraryID <> rs2!LibraryID Or rs!SecurityLevel <> rs2!SecurityLevel _
                                    Or DecryptField(rs!Password) <> DecryptField(rs2!Password) Then
                                rs.Edit
                                rs!LibraryID = rs2!LibraryID
                                rs!SecurityLevel = rs2!SecurityLevel
                                rs!Password = rs2!Password
                                rs!CheckSum = BuildCheckSum(rs, "tblRules")
                                rs.Update
                            End If
                        End If
                    End If
                    
                    ' DAJ 11/3/2004: Fix any rules that may be in a strategy but still
                    ' have a Category ID assigned to them.
                    ' TLB 7/14/2014: And check for tab in PreviewRTF (this shouldn't happen,
                    ' but it did somehow happen in Chris Johnston's libraries.mdb!)
                    strPreview = DecryptField(rs!PreviewRTF)
                    If InStr(strPreview, vbTab) > 0 Or (rs!SystemNumber <> 0 And rs!CategoryID <> 0) Then
                        rs.Edit
                        If rs!SystemNumber <> 0 Then
                            rs!CategoryID = 0
                        End If
                        If InStr(strPreview, vbTab) > 0 Then
                            ' don't know how the tab go there, but just replace it with a space
                            strPreview = Replace(strPreview, vbTab, " ")
                            EncryptField rs!PreviewRTF, strPreview
                        End If
                        rs!CheckSum = BuildCheckSum(rs, "tblRules")
                        rs.Update
                    End If
        
                    If bAdd Then
                        .AddRecord Str(rs!RuleID) & vbTab & rs!Name & _
                            vbTab & Str(rs!RuleType) & vbTab & Str(CLng(rs!BuySell)) & vbTab & Str(rs!LibraryID) & _
                            vbTab & Str(CDbl(rs!LastModified)) & vbTab & strPreview & _
                            vbTab & Str(rs!SecurityLevel) & vbTab & DecryptField(rs!Password) & _
                            vbTab & Str(CLng(rs!CannotDelete)) & vbTab & Str(rs!SystemNumber) & _
                            vbTab & Str(CLng(rs!Reverify)) & vbTab & Str(CLng(NullChk(rs!CategoryID, 0&)))
                    End If
                End If
            
                rs.MoveNext
            Loop
        End With
    End If
    
    ' TLB and DAJ 7/27/2012: need to remove any orphaned rules from invalid systems
    hRules = g.tblRule.TableHandle
    For lIndex = g.tblRule.NumRecords - 1 To 0 Step -1
        lSystemID = gdGetTableNum(hRules, RuleField(etblRule_SystemNumber), lIndex)
        If lSystemID <> 0 Then
            strSystem = aSystems(lSystemID)
            If Len(strSystem) = 0 Then
                strSystem = SystemNameForID(lSystemID)
                aSystems(lSystemID) = strSystem
            End If
            If Len(strSystem) = 0 Then
                ' this rule must be from an invalid system, so remove it
                g.tblRule.RemoveRecords lIndex, 1
            End If
        End If
    Next
    
    Set aSystems = Nothing
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.LoadRulesTable", eGDRaiseError_Raise
    
End Sub

Public Sub DeleteRule(ByVal lRuleID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Proper Index into the table

    If g.tblRule.FieldArray(RuleField(etblRule_RuleID), False).BinarySearch(lRuleID, lIndex) Then
        g.tblRule.RemoveRecords lIndex
    End If
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mSysNav.DeleteRule", eGDRaiseError_Raise
    
End Sub

Public Sub RefreshRule(r As cRule)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Proper Index into the table
    
    If g.tblRule.FieldArray(RuleField(etblRule_RuleID), False).BinarySearch(r.RuleID, lIndex) Then
        g.tblRule.RemoveRecords lIndex
    End If
    
    g.tblRule.AddRecord Str(r.RuleID) & vbTab & r.Name & vbTab & r.RuleType & vbTab & _
            CLng(r.BuySell) & vbTab & Str(r.LibraryID) & vbTab & Str(CDbl(r.LastModified)) & vbTab & _
            r.CondFillWords & vbTab & r.SecurityLevel & vbTab & r.Password & vbTab & _
            CLng(r.CannotDelete) & vbTab & r.SystemNumber & vbTab & CLng(r.Reverify) & vbTab & Str(r.CategoryID), _
            lIndex, vbTab
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.RefreshRule", eGDRaiseError_Raise
    
End Sub

Public Function GetMarketInfo(ByVal strSymbol As String, Bars As cGdBars, Optional ByVal bJustDefaults As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim lMarketID As Long               ' Market ID to get information from gMarkets
    Dim strMarketSymbol As String       ' Symbol to look up in the gMarkets
    Dim rs As Recordset                 ' Recordset into the database
    
    ' Try to get the Data Manager properties for the given symbol...
    If SetBarProperties(Bars, strSymbol) Then
        strMarketSymbol = Bars.Prop(eBARS_MarketSymbol)
    Else
        strMarketSymbol = strSymbol
    End If
    
    ' Try to get the MarketID into gMarkets for the given symbol...
    'g.Markets.GetID strMarketSymbol, lMarketID
    
    ' Add the Margin information to the Bars...
    'If g.Markets.Found(lMarketID) Then
    '    Bars.Prop(eBARS_Margin) = g.Markets.Item(CStr(lMarketID)).Margin
    'Else
    '    Bars.Prop(eBARS_Margin) = 0
    'End If
    
    ' If not just loading defaults, get overrides out of the database...
    If Not bJustDefaults Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarketInfo] " & _
                    "WHERE [Symbol]='" & strSymbol & "' AND [SymbolID]=" & Bars.Prop(eBARS_SymbolID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            Select Case rs!DataType
                Case MarketType(eMarketType_Desc)
                    Bars.Prop(eBARS_Desc) = rs!Value
                Case MarketType(eMarketType_SecurityType)
                    Bars.Prop(eBARS_SecurityType) = Asc(rs!Value)
                Case MarketType(eMarketType_TickMove)
                    Bars.Prop(eBARS_TickMove) = Val(rs!Value)
                Case MarketType(eMarketType_TickValue)
                    Bars.Prop(eBARS_TickValue) = Val(rs!Value)
                Case MarketType(eMarketType_MinMoveInTicks)
                    Bars.Prop(eBARS_MinMoveInTicks) = Val(rs!Value)
                Case MarketType(eMarketType_Margin)
                    Bars.Prop(eBARS_Margin) = Val(rs!Value)
            End Select
            
            rs.MoveNext
        Loop
    End If
    
    GetMarketInfo = Bars.Prop(eBARS_SymbolID)

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.GetMarketInfo", eGDRaiseError_Raise
    
End Function

Public Function RuleCopy(ByVal lRuleID As Long, ByVal lOldSystemID As Long, ByVal lNewSystemID As Long) As Long
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database
    Dim rs2 As Recordset                ' Recordset from the database
    Dim rs3 As Recordset                ' Recordset from the database
    Dim Rule As New cRule               ' Rule object for copying
    Dim alOldParmIDs As New cGdArray    ' Old Parameter ID's
    Dim alNewParmIDs As New cGdArray    ' New Parameter ID's
    Dim astrParmNames As New cGdArray   ' Parameter Names
    Dim alParmTypeIDs As New cGdArray   ' Parameter Types
    Dim lIndex As Long                  ' Index into a for loop
    Dim bLinkedInputs As Boolean        ' Does the new system link inputs?
    Dim bNewSystem As Boolean           ' Are we changing systems?
    Dim NewRule As New cRule
    
    ' Create the arrays...
    alOldParmIDs.Create eGDARRAY_Longs
    alNewParmIDs.Create eGDARRAY_Longs
    astrParmNames.Create eGDARRAY_Strings
    alParmTypeIDs.Create eGDARRAY_Longs
    
    ' Copy the Rule and it's Parms...
    With Rule
        .RuleID = lRuleID
        .Load
        For lIndex = 1 To .Inputs.Count
            alOldParmIDs(lIndex - 1) = .Inputs.Item(lIndex).ParmID
            astrParmNames(lIndex - 1) = .Inputs.Item(lIndex).ParmName
            alParmTypeIDs(lIndex - 1) = .Inputs.Item(lIndex).ParmTypeID
        Next lIndex
        .RuleID = 0
        
        If lNewSystemID > 0 Then
            .SystemNumber = lNewSystemID
        ElseIf lNewSystemID = 0 Then
            .SystemNumber = -2&
        Else
            .SystemNumber = 0&
        End If
        
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] WHERE [SystemNumber]=" & lNewSystemID & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            .LibraryID = rs!LibraryID
        Else
            .LibraryID = kSN_UserLibrary
        End If
        
        .Save
        For lIndex = 1 To .Inputs.Count
            alNewParmIDs(lIndex - 1) = .Inputs.Item(lIndex).ParmID
        Next lIndex
        RuleCopy = .RuleID
    End With
    
    ' If we are making a "favorite" then we don't need to do the system stuff...
    If lNewSystemID <> -1 Then
        ' If we are changing systems, see if the Linked Inputs flag is turned on for the
        ' new system...
        bLinkedInputs = False
        bNewSystem = (lNewSystemID <> lOldSystemID)
        If bNewSystem Then
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] WHERE [SystemNumber]=" & lNewSystemID & ";", dbOpenDynaset)
            If Not rs.EOF Then
                bLinkedInputs = CBool(rs!LinkInputs)
            End If
        End If
        
        ' Copy the System Rule record...
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemRules] WHERE [SystemNumber]=" & lOldSystemID & " AND [RuleID]=" & lRuleID & ";", dbOpenDynaset)
        If Not rs.EOF Then
            Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemRules];", dbOpenDynaset)
            rs2.AddNew
            rs2!SystemNumber = lNewSystemID
            rs2!RuleID = Rule.RuleID
            rs2!Seq = rs!Seq
            rs2!Selected = rs!Selected
            rs2!Alternate = rs!Alternate
            rs2!RuleUse = rs!RuleUse
            rs2!LastModifiedKnown = rs!LastModifiedKnown
            rs2!UnitsID = rs!UnitsID
            rs2!LinkedRules = rs!LinkedRules
            rs2!ExitOnEntryBar = rs!ExitOnEntryBar
            rs2!ExitBasedOnEachTrade = rs!ExitBasedOnEachTrade
            rs2!NumberContracts = rs!NumberContracts
            rs2!AsPercentOfPosition = rs!AsPercentOfPosition
            rs2!CheckSum = BuildCheckSum(rs2, "tblSystemRules")
            rs2.Update
        
            ' Copy the System Parms...
            For lIndex = 0 To alOldParmIDs.Size - 1
                ' System Parm
                If alParmTypeIDs(lIndex) <> 5 Then
                    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemParms] WHERE [SystemNumber]=" & lOldSystemID & " AND [ParmID]=" & alOldParmIDs(lIndex) & ";", dbOpenDynaset)
                    If Not rs.EOF Then
                        Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemParms];", dbOpenDynaset)
                        rs2.AddNew
                        rs2!SystemNumber = lNewSystemID
                        rs2!ParmID = alNewParmIDs(lIndex)
                        If bLinkedInputs And bNewSystem Then
                            Set rs3 = g.dbNav.OpenRecordset("SELECT tblRuleParms.ParmName, tblSystemParms.* " & _
                                            "FROM tblRuleParms INNER JOIN tblSystemParms ON tblRuleParms.ParmID = tblSystemParms.ParmID " & _
                                            "WHERE [SystemNumber]=" & lNewSystemID & " AND [ParmName]='" & astrParmNames(lIndex) & "';", dbOpenDynaset)
                            ' Need to use the values from the input in the system with the
                            ' same name since the linked inputs flag is turned on for this system
                            If Not rs3.EOF Then
                                rs2!Value = rs3!Value
                                rs2!IfOptimize = rs3!IfOptimize
                                rs2!OptFromValue = rs3!OptFromValue
                                rs2!OptToValue = rs3!OptToValue
                                rs2!OptStepValue = rs3!OptStepValue
                                rs2!OptListID = rs3!OptListID
                                
                            ' Even though the linked inputs flag is turned on for the new system,
                            ' there are no inputs by this parm name, so we can retain the values
                            Else
                                rs2!Value = rs!Value
                                rs2!IfOptimize = rs!IfOptimize
                                rs2!OptFromValue = rs!OptFromValue
                                rs2!OptToValue = rs!OptToValue
                                rs2!OptStepValue = rs!OptStepValue
                                rs2!OptListID = rs!OptListID
                            End If
                            
                        ' We are either going to the same system, or the linked inputs flag is
                        ' turned off, so we can retain the values
                        Else
                            rs2!Value = rs!Value
                            rs2!IfOptimize = rs!IfOptimize
                            rs2!OptFromValue = rs!OptFromValue
                            rs2!OptToValue = rs!OptToValue
                            rs2!OptStepValue = rs!OptStepValue
                            rs2!OptListID = rs!OptListID
                        End If
                        rs2.Update
                    End If
                
                ' System Security
                Else
                    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemSecurities] WHERE [SystemNumber]=" & lOldSystemID & " AND [ParmID]=" & alOldParmIDs(lIndex) & ";", dbOpenDynaset)
                    If Not rs.EOF Then
                        Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemSecurities];", dbOpenDynaset)
                        rs2.AddNew
                        rs2!SystemNumber = lNewSystemID
                        rs2!ParmID = alNewParmIDs(lIndex)
                        If bNewSystem Then
                            Set rs3 = g.dbNav.OpenRecordset("SELECT tblRuleParms.ParmName, tblSystemParms.* " & _
                                            "FROM tblRuleParms INNER JOIN tblSystemParms ON tblRuleParms.ParmID = tblSystemParms.ParmID " & _
                                            "WHERE [SystemNumber]=" & lNewSystemID & " AND [ParmName]='" & astrParmNames(lIndex) & "';", dbOpenDynaset)
                            ' This Security exists in the new system, so we need to use the
                            ' values from the new system
                            If Not rs3.EOF Then
                                rs2!Path = rs3!Path
                                rs2!Symbol = rs3!Symbol
                                rs2!MarketSymbol = rs3!MarketSymbol
                                rs2!Periodicity = rs3!Periodicity
                                rs2!Format = rs3!Format
                                rs2!SecurityType = rs3!SecurityType
                                rs2!SecurityName = rs3!SecurityName
                                
                            ' This security does not exist in the new system, so retain the
                            ' values from the old system
                            Else
                                rs2!Path = rs!Path
                                rs2!Symbol = rs!Symbol
                                rs2!MarketSymbol = rs!MarketSymbol
                                rs2!Periodicity = rs!Periodicity
                                rs2!Format = rs!Format
                                rs2!SecurityType = rs!SecurityType
                                rs2!SecurityName = rs!SecurityName
                            End If
                        
                        ' We are going to the same system, so we need to retain the values
                        Else
                            rs2!Path = rs!Path
                            rs2!Symbol = rs!Symbol
                            rs2!MarketSymbol = rs!MarketSymbol
                            rs2!Periodicity = rs!Periodicity
                            rs2!Format = rs!Format
                            rs2!SecurityType = rs!SecurityType
                            rs2!SecurityName = rs!SecurityName
                        End If
                        rs2.Update
                    End If
                End If
            Next lIndex
        
        ' Going from a "Favorite" to a local rule, so we need to create the
        ' system information
        ElseIf lOldSystemID = -1 Then
            Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemRules];", dbOpenDynaset)
            rs2.AddNew
            rs2!SystemNumber = lNewSystemID
            rs2!RuleID = Rule.RuleID
            rs2!Seq = 0
            rs2!Selected = True
            rs2!Alternate = False
            rs2!RuleUse = Rule.RuleType
            rs2!LastModifiedKnown = Rule.LastModified
            rs2!UnitsID = 0
            rs2!LinkedRules = ""
            rs2!ExitOnEntryBar = Rule.ExitOnEntryBar
            rs2!ExitBasedOnEachTrade = Rule.ExitBasedOnEachTrade
            rs2!NumberContracts = Rule.NumberContracts
            rs2!AsPercentOfPosition = Rule.AsPercentOfPosition
            rs2!CheckSum = BuildCheckSum(rs2, "tblSystemRules")
            rs2.Update
            
            For lIndex = 0 To alOldParmIDs.Size - 1
                If alParmTypeIDs(lIndex) <> 5 Then
                    Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemParms];", dbOpenDynaset)
                    rs2.AddNew
                    rs2!SystemNumber = lNewSystemID
                    rs2!ParmID = alNewParmIDs(lIndex)
                    If bLinkedInputs Then
                        Set rs3 = g.dbNav.OpenRecordset("SELECT tblRuleParms.ParmName, tblSystemParms.* " & _
                                        "FROM tblRuleParms INNER JOIN tblSystemParms ON tblRuleParms.ParmID = tblSystemParms.ParmID " & _
                                        "WHERE [SystemNumber]=" & lNewSystemID & " AND [ParmName]='" & astrParmNames(lIndex) & "';", dbOpenDynaset)
                        ' Need to use the values from the input in the system with the
                        ' same name since the linked inputs flag is turned on for this system
                        If Not rs3.EOF Then
                            rs2!Value = rs3!Value
                            rs2!IfOptimize = rs3!IfOptimize
                            rs2!OptFromValue = rs3!OptFromValue
                            rs2!OptToValue = rs3!OptToValue
                            rs2!OptStepValue = rs3!OptStepValue
                            rs2!OptListID = rs3!OptListID
                            
                        ' Even though the linked inputs flag is turned on for the new system,
                        ' there are no inputs by this parm name, so we can retain the values
                        Else
                            rs2!Value = Rule.Inputs.Item(lIndex + 1).Value
                            rs2!IfOptimize = False
                            rs2!OptFromValue = 0
                            rs2!OptToValue = 0
                            rs2!OptStepValue = 0
                            rs2!OptListID = 0
                        End If
                    
                    ' Linked inputs flag is turned off, so we can retain the values
                    Else
                        rs2!Value = Rule.Inputs.Item(lIndex + 1).Value
                        rs2!IfOptimize = False
                        rs2!OptFromValue = 0
                        rs2!OptToValue = 0
                        rs2!OptStepValue = 0
                        rs2!OptListID = 0
                    End If
                    rs2.Update
                
                Else
                    Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemSecurities];", dbOpenDynaset)
                    rs2.AddNew
                    rs2!SystemNumber = lNewSystemID
                    rs2!ParmID = alNewParmIDs(lIndex)
                    If bNewSystem Then
                        Set rs3 = g.dbNav.OpenRecordset("SELECT tblRuleParms.ParmName, tblSystemParms.* " & _
                                        "FROM tblRuleParms INNER JOIN tblSystemParms ON tblRuleParms.ParmID = tblSystemParms.ParmID " & _
                                        "WHERE [SystemNumber]=" & lNewSystemID & " AND [ParmName]='" & astrParmNames(lIndex) & "';", dbOpenDynaset)
                        ' This Security exists in the new system, so we need to use the
                        ' values from the new system
                        If Not rs3.EOF Then
                            rs2!Path = rs3!Path
                            rs2!Symbol = rs3!Symbol
                            rs2!MarketSymbol = rs3!MarketSymbol
                            rs2!Periodicity = rs3!Periodicity
                            rs2!Format = rs3!Format
                            rs2!SecurityType = rs3!SecurityType
                            rs2!SecurityName = rs3!SecurityName
                            
                        ' This security does not exist in the new system, so retain the
                        ' values from the old system
                        Else
                            rs2!Path = ""
                            rs2!Symbol = ""
                            rs2!MarketSymbol = ""
                            rs2!Periodicity = ""
                            rs2!Format = ""
                            rs2!SecurityType = ""
                            rs2!SecurityName = ""
                        End If
                    
                    ' We are going to the same system, so we need to retain the values
                    Else
                        rs2!Path = ""
                        rs2!Symbol = ""
                        rs2!MarketSymbol = ""
                        rs2!Periodicity = ""
                        rs2!Format = ""
                        rs2!SecurityType = ""
                        rs2!SecurityName = ""
                    End If
                    rs2.Update
                End If
            Next lIndex
        End If
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.RuleCopy", eGDRaiseError_Raise
    
End Function

Public Function NextSystemID() As Long
On Error GoTo ErrSection

    Static lNextSystemID As Long
    
    lNextSystemID = lNextSystemID - 1
    NextSystemID = lNextSystemID

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.NextSystemID", eGDRaiseError_Raise
    
End Function

Public Sub AddRuleToFavorites(ByVal Rule As cRule)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim NewRule As New cRule            ' Copy of the rule passed in
    Dim strNewName As String            ' New name for the rule
    Dim lRuleNum As Long                ' Temporary rule number
    Dim lRuleID As Long                 ' New ID for the rule
    Dim bOverwrite As Boolean           ' Overwrite the existing favorite?
    
    bOverwrite = False
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] WHERE [SystemNumber]=0 " & _
                    "ORDER BY [Name];", dbOpenDynaset)
    rs.FindFirst "[Name]='" & Rule.Name & "'"
    If rs.NoMatch Then
        lRuleID = -1&
        strNewName = Rule.Name
    Else
        If InfBox(UCase(Rule.Name) & "|already exists as a favorite.||Would you like to overwrite it?", "?", "+Yes|-No", "Confirmation") = "Y" Then
            lRuleID = rs!RuleID
            strNewName = Rule.Name
            bOverwrite = True
        Else
            lRuleNum = 2&
            strNewName = Rule.Name
            Do While Not rs.NoMatch
                strNewName = Trim(Rule.Name) & " #" & Format(lRuleNum, "00")
                lRuleNum = lRuleNum + 1
                rs.FindFirst "[Name]='" & strNewName & "'"
            Loop
            
            lRuleID = -1&
        End If
    End If

    Set NewRule = Rule.MakeCopy(lRuleID, 0&)
    NewRule.Name = strNewName
    If lRuleID = -1& Then
        NewRule.LibraryID = kSN_UserLibrary
        NewRule.SecurityLevel = 0
        NewRule.Password = ""
        NewRule.CannotDelete = False
        If NewRule.RuleType = 0 Then
            NewRule.CategoryID = RuleCategoryIDFromName("Other Entries")
        Else
            NewRule.CategoryID = RuleCategoryIDFromName("Other Exits")
        End If
    End If
    NewRule.SaveWithSystemInfo
    RefreshRule NewRule
    
    If bOverwrite Then
        frmUpdateLocals.ShowMe NewRule
    End If

ErrExit:
    Set NewRule = Nothing
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set NewRule = Nothing
    Set rs = Nothing
    RaiseError "mSysNav.AddRuleToFavorites", eGDRaiseError_Raise
    
End Sub

Public Function ValidateMarketInfo(ByVal Bars As cGdBars, Optional ByVal bShowMsg As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim dTemp As Double

    dTemp = Bars.Prop(eBARS_TickMove)
    If dTemp <= 0 Or dTemp > 1000000 Then
        Err.Raise vbObjectError + 1000, , _
            "Tick Move must be greater than 0 and less than 1,000,000"
    End If
        
    dTemp = Bars.Prop(eBARS_TickValue)
    If dTemp <= 0 Or dTemp > 1000000 Then
        Err.Raise vbObjectError + 1000, , _
            "Tick Value must be greater than 0 and less than 1,000,000"
    End If
    
    dTemp = Bars.Prop(eBARS_MinMoveInTicks)
    If dTemp <= 0 Or dTemp > 1000000 Then
        Err.Raise vbObjectError + 1000, , _
            "Min Move must be greater than 0 and less than 1,000,000"
    End If
    
    dTemp = Bars.Prop(eBARS_Margin)
    If dTemp < 0 Or dTemp > 1000000 Then
        Err.Raise vbObjectError + 1000, , "Margin must be between 0 and 1,000,000"
    End If
    
    ValidateMarketInfo = True
    
ErrExit:
    Exit Function
    
ErrSection:
    If bShowMsg Then
        RaiseError "mSysNav.ValidateMarketInfo", eGDRaiseError_Raise
    Else
        Err.Clear
    End If
    
End Function

Public Sub RefreshReverify(Optional ByVal nSpreadFuncId& = 0)
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim Rule As cRule
    Dim Func As cFunction
    Dim bReloadEngineFunctions As Boolean

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] WHERE [Reverify]=True;", dbOpenDynaset)
    Do While Not rs.EOF
        Set Rule = New cRule
        Rule.LoadWithSystemInfo rs!RuleID
        RefreshRule Rule
        
        rs.MoveNext
    Loop

    bReloadEngineFunctions = False
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] WHERE [Reverify]=True;", dbOpenDynaset)
    Do While Not rs.EOF
        Set Func = New cFunction
        Func.FunctionID = rs!FunctionID
        Func.Load
        RefreshFunction Func, , False
        bReloadEngineFunctions = True
        
        rs.MoveNext
    Loop

    If bReloadEngineFunctions Then
        LoadEngineFunctionSet ""
    End If

    ' now that all functions and rules that were effected have been reverified,
    ' update the charts to reflect the changes to indicators or systems
    UpdateVisibleCharts eRedo6_ReloadInd, , nSpreadFuncId
    
ErrExit:
    Set Func = Nothing
    Set Rule = Nothing
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set Func = Nothing
    Set Rule = Nothing
    Set rs = Nothing
    RaiseError "mSysNav.RefreshReverify", eGDRaiseError_Raise
    
End Sub

Public Function FunctionIDFromName(ByVal strName As String) As Long
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                "WHERE [FunctionName]='" & strName & "';", dbOpenDynaset)
    If Not rs.EOF Then FunctionIDFromName = rs!FunctionID

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.FunctionIDFromName", eGDRaiseError_Raise
    
End Function

Public Function FunctionNameFromID(ByVal lID As Long) As String
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                "WHERE [FunctionID]=" & Str(lID) & ";", dbOpenDynaset)
    If Not rs.EOF Then FunctionNameFromID = rs!FunctionName

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.FunctionNameFromID", eGDRaiseError_Raise
    
End Function

Public Function RuleRefs(ByVal hRefs As Long) As cGdArray
On Error GoTo ErrSection:

    Dim alRefs As New cGdArray
    Dim lIndex As Long
    Dim alFunctions As New cGdArray
    
    alFunctions.Create eGDARRAY_Longs
    If alRefs.CopyFromHandle(hRefs) Then
        For lIndex = 0 To alRefs.Size - 1
            FunctionRefs alRefs(lIndex), alFunctions
        Next lIndex
    End If
    
    Set RuleRefs = alFunctions

ErrExit:
    Set alRefs = Nothing
    Set alFunctions = Nothing
    Exit Function
    
ErrSection:
    Set alRefs = Nothing
    Set alFunctions = Nothing
    RaiseError "mSysNav.RuleRefs", eGDRaiseError_Raise
    
End Function

Public Function FunctionRefs(ByVal lFunctionID As Long, Optional alRefs As cGdArray = Nothing) As cGdArray
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    ' Create the array if it does not exist yet...
    If alRefs Is Nothing Then
        Set alRefs = New cGdArray
        alRefs.Create eGDARRAY_Longs
    End If
    
    ' Add the Function ID passed in and re-sort the array...
    alRefs.Add lFunctionID
    alRefs.Sort eGdSort_DeleteDuplicates
    
    ' Find all of the Functions that this Function uses...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionRefs] " & _
                "WHERE [FunctionID]=" & Str(lFunctionID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        Do While Not rs.EOF
            ' If we have not already added this Function (and it's references),
            ' find the Functions that it uses...
            If alRefs.BinarySearch(rs!FunctionIDRef) = False Then
                FunctionRefs rs!FunctionIDRef, alRefs
            End If
            
            rs.MoveNext
        Loop
    End If
    
    ' Return the array of references...
    Set FunctionRefs = alRefs

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.FunctionRefs", eGDRaiseError_Raise
    
End Function

Public Function IsBooleanExpression(ByVal strExpression As String) As Boolean
On Error GoTo ErrSection:

    Dim Expr As cExpression

    Set Expr = New cExpression
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        .ValidateFunctionRule strExpression
        
        If .FunctionReturnType = kSN_RetTrueFalse Or .FunctionReturnType = kSN_RetTrueFalseConstant Then
            IsBooleanExpression = True
        Else
            IsBooleanExpression = False
        End If
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IsBooleanExpression", eGDRaiseError_Raise
    
End Function

Public Sub ImportLibrary()
On Error GoTo ErrSection:

    Dim LibMgrBridge As cLibManagerBridge   ' Bridge to the library manager DLL
    Dim astrQbtFiles As cGdArray            ' Quote board tab files
    Dim lIndex As Long                      ' Index into a for loop
    Dim bNewQbtFiles As Boolean             ' Are there new qbt files?
    Dim strQbtFileName As String        ' QBT file name with path
    Dim strDonFileName As String        ' Done file to compare with
    Dim strQbtDate As String            ' Date/Time stamp for the QBT file
    Dim strDonDate As String            ' Date/Time stamp for the DON file
       
    If g.RealTime.Active Then
        InfBox "You cannot import a library while streaming is on.", "!", , "Import Library Error"
    Else
        Set LibMgrBridge = GetLibMgrBridge
        With LibMgrBridge
            .ShowImporter
            
            If .ImportOK Then
                InfBox "Please wait while reloading functions and rules...", , , "Reloading", True
                Screen.MousePointer = vbHourglass
                
                ' Clean up any bogus pyramiding information (but don't reload the rules table
                ' in here because it will be done right after this)...
                FixPyramidInfo False
                
                ' Reload the Function and Rule tables in memory...
                LoadEngineFunctions
                LoadRulesTable
                Set g.Functions = New cFunctions
                g.Functions.Load
                FilterFunctions
                
                g.bDirtyLibrariesMDB = True
                
                ' Check to see if a new quote board tab file has been imported...
                bNewQbtFiles = False
                Set astrQbtFiles = New cGdArray
                astrQbtFiles.GetMatchingFiles AddSlash(App.Path) & "QBT\*.QBT", False
                For lIndex = 0 To astrQbtFiles.Size - 1
                    strQbtFileName = AddSlash(App.Path) & "QBT\" & astrQbtFiles(lIndex)
                    strDonFileName = AddSlash(App.Path) & "QBT\" & FileBase(astrQbtFiles(lIndex)) & ".DON"
            
                    strQbtDate = FileToString(strQbtFileName, , True)
                    strDonDate = FileToString(strDonFileName, , True)
                    
                    If Val(strQbtDate) > Val(strDonDate) Then
                        bNewQbtFiles = True
                    Else
                        FileCopy strDonFileName, strQbtFileName, True
                    End If
                Next lIndex
                
                ' If a quote board tab file was imported, the user will need to restart Trade Navigator
                ' in order for this to take effect...
                If bNewQbtFiles Then
                    If InfBox("A new quote board tab file has been imported.  In order for these changes to take effect, you must restart Trade Navigator.||Would you like to do that now?", "?", "+Yes|-No", "Quote Board Tab Import") = "Y" Then
                        frmMain.tmrMain.Tag = "QUIT"
                    End If
                End If
                
                mSysNav.CreateGuruAutoTradeItems
            End If
        End With
    End If
    
ErrExit:
    Screen.MousePointer = vbDefault
    InfBox ""
    Set LibMgrBridge = Nothing
    
    ' If the RestoreMDB.FLG file exists after attempting an import, we need to
    ' shut down the program so that when they start it back up, we can restore
    ' the old database
    If FileExist(AddSlash(App.Path) & "RestoreMDB.FLG") Then
        frmMain.tmrMain.Tag = "QUIT"
    End If
    
    Exit Sub
    
ErrSection:
    Screen.MousePointer = vbDefault
    InfBox ""
    Set LibMgrBridge = Nothing
    
    ' If the RestoreMDB.FLG file exists after attempting an import, we need to
    ' shut down the program so that when they start it back up, we can restore
    ' the old database
    If FileExist(AddSlash(App.Path) & "RestoreMDB.FLG") Then
        frmMain.tmrMain.Tag = "QUIT"
    End If
    
    RaiseError "mSysNav.ImportLibrary", eGDRaiseError_Raise
    
End Sub

Public Function GetLibMgrBridge() As cLibManagerBridge
On Error GoTo ErrSection:

    Dim LibMgrBridge As New cLibManagerBridge
    With LibMgrBridge
        .CalledFrom = SystemNavigator
        .AppPath = App.Path
        .dbNavRef = g.dbNav
        .ImageList = frmMain.img16.ListImages
        .dbNavPassword = DbPassword
        .CustomerID = g.lLCD
        .HonestDate = RI_HonestDate
        .Help = g.Help
        .ShowShadow = (IsIDE Or HasModule("SHADOW"))
        
        ' set owner so Modal Lib forms won't get "lost" when clicking back from another app
        ' (but can't set it while in IDE since it complains about a mismatch)
        If Not IsIDE Then
            .OwnerForm = frmMain
        End If
    End With

ErrExit:
    Set GetLibMgrBridge = LibMgrBridge
    Exit Function

ErrSection:
    RaiseError "mSysNav.GetLibMgrBridge", eGDRaiseError_Raise
    
End Function

Public Function FilterFunctions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    If IsIDE = False Then
        For lIndex = g.Functions.Count To 1 Step -1
            If HasModule(g.Functions.Item(lIndex).RequiredMod) = False Then
                g.Functions.Remove lIndex
            End If
        Next lIndex
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.FilterFunctions", eGDRaiseError_Raise
    
End Function

Public Function CalcFirstDate(ByVal lNumBars As Long) As Long
On Error GoTo ErrSection:

    Dim lFromDate As Long
    Dim lLastDateOfData As Long
    Dim lLongDate As Long
    Dim lMonth As Long
    Dim lYear As Long
    Dim lDay As Long

    lLastDateOfData = LastDailyDownload
    lFromDate = lLastDateOfData - Int(lNumBars * 1.46 + 0.5) - 2
    lLongDate = JulToLong(lFromDate, 1)
    lYear = lLongDate / 10000&
    lMonth = (lLongDate / 100) Mod 100
    lDay = lLongDate Mod 100
    
    If lDay > 5 Then
        CalcFirstDate = JulFromLong((lYear * 10000) + (lMonth * 100) + 1)
    Else
        CalcFirstDate = JulFromLong((lYear * 10000) + ((lMonth - 1) * 100) + 1)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.CalcFirstDate", eGDRaiseError_Raise
    
End Function

' call this before RunExpressions with the BarNames and Expressions
' call this after RunExpressions without the BarNames to clear the expressions
' (first parm of astrParms should be the Name of the expression set -- can be empty)
Public Function SetupExpressions(astrParms As cGdArray, _
                    Optional astrBarNames As cGdArray = Nothing, _
                    Optional astrExpressions As cGdArray = Nothing, _
                    Optional strError As String = "") As Boolean
On Error GoTo ErrSection:

    Dim rc&, i&, strCodedText$

    m_UseAdvForRunExpressions = (FileLength(App.Path & "\Provided\UseAdvForRun.flg") > 3)

    ' TLB 1/31/2007: do NOT allow inadvertantly passing an empty string
    ' since will clear everything
    If Len(Trim(astrParms(0))) = 0 Then Exit Function

    ' always clear first to make sure everything has been reset
    astrParms(1) = "" '(function set name)
    astrParms.Size = 2
    If m_UseAdvForRunExpressions Then
        rc = ClearExpressionsNEW(astrParms.ArrayHandle, ByVal 0&)
    Else
        rc = ClearExpressionsOLD(astrParms.ArrayHandle, ByVal 0&)
    End If
    
    If Not astrBarNames Is Nothing Then
        ' TLB 9/20/2012: rename the "LW Sentiment" functions
        For i = 0 To astrExpressions.Size - 1
            strCodedText = astrExpressions(i)
            If InStr(strCodedText, "LWSentiment") > 0 Then
                strCodedText = Replace(strCodedText, "LW Sentiment", "TN Consensus")
                astrExpressions(i) = strCodedText
            End If
        Next
    
        ' now init the expression set
        astrParms(1) = "" '(function set name)
        astrParms.Size = 2
        If m_UseAdvForRunExpressions Then
            rc = InitExpressionsNEW(astrParms.ArrayHandle, _
                astrBarNames.ArrayHandle, astrExpressions.ArrayHandle, ByVal 0&)
        Else
            rc = InitExpressionsOLD(astrParms.ArrayHandle, _
                astrBarNames.ArrayHandle, astrExpressions.ArrayHandle, ByVal 0&)
        End If
    End If
    
    If rc = 0 Then
        SetupExpressions = True
    Else
        strError = astrParms(astrParms.Size - 1)
    End If
    
    astrParms.Size = 1

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.SetupExpressions", eGDRaiseError_Raise

End Function

'StrParms (string array):
'  0:  Expression set name
'  1:  Last Bar Good flag - assumed true unless 'false'
'StrBars - string array of parm names of the bars, e.g., Market1
'BarsArray - array of bars handles (same size as StrBars)
'ExpResults - array of results matching the expressions
Public Function RunExpressions(ByVal hStrParms&, ByVal hStrBars&, ByVal hBarsArray&, _
        ByVal hExpResults&, Optional ByVal hMinBarsReq& = 0, Optional ByVal hStrTimes& = 0) As Long
On Error GoTo ErrSection:

    Dim rc&
    
    If m_UseAdvForRunExpressions Then
        rc = RunExpressionsNEW(hStrParms, hStrBars, hBarsArray, hExpResults, hMinBarsReq, hStrTimes)
    Else
        rc = RunExpressionsOLD(hStrParms, hStrBars, hBarsArray, hExpResults, hMinBarsReq, hStrTimes)
    End If

ErrExit:
    RunExpressions = rc
    Exit Function

ErrSection:
    RaiseError "mSysNav.RunExpressions"
End Function

Public Sub ShowMergedReports(RptBridge As cRptBridge, ByVal strReportName As String, _
            ByVal bPyramid As Boolean, ByVal hTradeFiles As Long, _
            Optional ByVal hTblRptRules As Long = 0&, Optional ByVal bHideTdoReports As Boolean = False, _
            Optional ByVal strCaptureFile As String = "")
On Error GoTo ErrSection:

    With RptBridge
        .AppPath = App.Path
        .AppName = "System Navigator"
        .DB = g.dbNav
        .DefaultBeginBalance = 100000
        .PortOrSystemName = strReportName
        .ShowInLocalTime = g.bShowInLocalTimeZone
        
        ' TLB 5/17/2011: more stuff to pass over
        .SetIrxBars GetIrxBarsHandle
        .SetAppBackColor GetAppBackColor
        .AltGridRowColor = ALT_GRID_ROW_COLOR
        
        .ChartHwnd = 0
        .Pyramiding = bPyramid
        .ImportMultipleWithHandles hTradeFiles, hTblRptRules
        .MainForm = frmMain
        .ImageList = frmMain.img16.ListImages
        .Help = g.Help
        .HideTdoReports = bHideTdoReports
        
        If Len(strCaptureFile) = 0 Then
            .Show True
        Else
            .CaptureReport 0, strCaptureFile
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.ShowMergedReports", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RuleCategoryFromID
'' Description: Return the Rule Category Name given a Category ID
'' Inputs:      Category ID
'' Returns:     Category Name if Found, Blank otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RuleCategoryFromID(ByVal lCategoryID As Long) As String
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database
    
    RuleCategoryFromID = ""
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleCategories] " & _
                "WHERE [CategoryID]=" & Str(lCategoryID) & ";", dbOpenDynaset)
    If Not (rs.EOF And rs.BOF) Then
        RuleCategoryFromID = rs!CategoryName
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.RuleCategoryFromID", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RuleCategoryIDFromName
'' Description: Return the Rule Category ID given a Category Name
'' Inputs:      Category Name
'' Returns:     Category ID if Found, Zero otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RuleCategoryIDFromName(ByVal strName As String) As Long
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database
    
    RuleCategoryIDFromName = 0
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleCategories] " & _
                "WHERE [CategoryName]='" & strName & "';", dbOpenDynaset)
    If Not (rs.EOF And rs.BOF) Then
        RuleCategoryIDFromName = rs!CategoryID
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "mSysNav.RuleCategoryIDFromName", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixSystemSecuritySymbolIDs
'' Description: If all of the Symbol ID's in the System Securities table are
''              zero, then walk through and set all of them that we can
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FixSystemSecuritySymbolIDs()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lSymbolID As Long               ' Symbol ID for the given symbol
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemSecurities] WHERE [SymbolID]<>0;", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemSecurities];", dbOpenDynaset)
        Do While Not rs.EOF
            lSymbolID = GetSymbolID(rs!Symbol)
            If lSymbolID <> 0 Then
                rs.Edit
                rs!SymbolID = lSymbolID
                rs.Update
            End If
        
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.FixSystemSecuritySymbolIDs", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixStrategyBasketFiles
'' Description: Add the Symbol ID to any MRG file and rename it to a SB file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FixStrategyBasketFiles()
On Error GoTo ErrSection:

    Dim astrFiles As New cGdArray       ' Array of MRG files in the custom folder
    Dim astrFile As New cGdArray        ' File to input/output
    Dim astrLine As New cGdArray        ' Split out records from a line in the file
    Dim lFile As Long                   ' Index into a for loop
    Dim lLine As Long                   ' Index into a for loop
    
    astrFiles.Create eGDARRAY_Strings
    astrFiles.GetMatchingFiles AddSlash(App.Path) & "Custom\*.MRG", True
    
    For lFile = 0 To astrFiles.Size - 1
        astrFile.FromFile astrFiles(lFile)
        For lLine = 1 To astrFile.Size - 1
            astrLine.SplitFields astrFile(lLine), vbTab
            If Len(astrLine(11)) = 0 And Len(astrLine(3)) = 0 Then
                astrLine(11) = Str(GetSymbolID(astrLine(2)))
            End If
            astrFile(lLine) = astrLine.JoinFields(vbTab)
        Next lLine
        
        astrFile.ToFile astrFiles(lFile)
        RenameFile astrFiles(lFile), Replace(astrFiles(lFile), ".MRG", ".SB")
    Next lFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.FixStrategyBasketFiles", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsBooleanCriteria
'' Description: Determine whether the criteria with the given id is boolean
'' Inputs:      ID of the Criteria
'' Returns:     True if a Boolean Criteria, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsBooleanCriteria(ByVal strCriteriaID As String) As Boolean
On Error GoTo ErrSection:

    Dim Criteria As New cCriteria       ' Temporary Criteria object
    
    IsBooleanCriteria = False
    If Len(strCriteriaID) > 0 Then
        Set Criteria = g.SymbolPool.Criterias(strCriteriaID)
        If Not Criteria Is Nothing Then
            IsBooleanCriteria = Criteria.IsBoolean
        End If
    End If

ErrExit:
    Set Criteria = Nothing
    Exit Function
    
ErrSection:
    Set Criteria = Nothing
    RaiseError "mSysNav.IsBooleanCriteria", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MarketsInExpressions
'' Description: Determine the secondary markets that are in the given
''              expressions and then load them up from the first date
'' Inputs:      Expressions, Start Date, Whether to include Snaphot Data,
''              Bar Names, Bars Collection, Default Period, Default Symbol,
''              Invalid secondary period?, Load Bars?
'' Returns:     True if Secondary Markets, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MarketsInExpressions(astrExpressions As cGdArray, ByVal dStartDate As Double, ByVal bIncludeSnapshotData As Boolean, _
                    astrBarNames As cGdArray, LoadedBars As cGdTree, ByVal strDefaultPeriod As String, _
                    Optional ByVal strDefaultSymbol As String = "", Optional bInvalidSecondaryPeriod As Boolean = False, _
                    Optional ByVal bLoadBars As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lIndex3 As Long                 ' Index into a for loop
    Dim astrTokens() As String          ' List of tokens from the coded text
    Dim lTokenType As Long              ' Type of the token
    Dim lTokenLength As Long            ' Length of the token
    Dim strTokenValue As String         ' Value of the token
    Dim strNewToken As String           ' Value of token to replace
    Dim bFound As Boolean               ' Was the market found in the array?
    Dim strSymbol As String             ' Symbol to load
    Dim strPeriod As String             ' Period of the symbol to load
    Dim Bars As New cGdBars             ' Temporary Bars collection
    Dim bRebuild As Boolean             ' Rebuild expression?
    Dim strToken As String
    Dim strType As String
    
    ' Default for the list of BarNames ...
    If astrBarNames.Size = 0 Then
        astrBarNames.Add "Market1"
        'astrBarNames.Add "Daily"
        astrBarNames.Add "Weekly"
        'astrBarNames.Add "Monthly"
    End If
    
    bInvalidSecondaryPeriod = False
    
    ' Get a list of names for secondary markets...
    For lIndex = 0 To astrExpressions.Size - 1
        bRebuild = False
        strToken = astrExpressions(lIndex)
        If InStr(strToken, "~") > 0 Then
            astrTokens = Split(strToken, "~")
            For lIndex2 = LBound(astrTokens) To UBound(astrTokens)
                strToken = Trim(astrTokens(lIndex2))
                If Left(strToken, 2) = "07" Then
                    lTokenType = Val(Left(strToken, 2))
                    lTokenLength = Val(Mid(strToken, 3, 3))
                    strTokenValue = Mid(strToken, 6, lTokenLength)
                                       
                    If lTokenType = 7 And UCase(strTokenValue) <> "MARKET1" Then
                        ' TLB 8/21/2009: rebuild a new token (starting with a copy of the old)
                        strNewToken = UCase(Trim(strTokenValue))
                        strNewToken = Replace(strNewToken, ", ", ",")
                        strNewToken = Replace(strNewToken, " ,", ",")
                        If Left(strNewToken, 1) = Chr(34) And Right(strNewToken, 1) = Chr(34) Then
                            strSymbol = Parse(Replace(strNewToken, Chr(34), ""), ",", 1)
                            strPeriod = Parse(Replace(strNewToken, Chr(34), ""), ",", 2)
                            ' TLB 4/15/2011: allow for symbols like "-057", "-099", "-201103", etc
                            ' to be used -- we will just prepend it with the base of the default symbol
                            If Left(strSymbol, 1) = "-" And Len(strDefaultSymbol) > 0 Then
                                ' if a future, replace just the contract portion of the symbol
                                If SecurityType(strDefaultSymbol) = "F" Then
                                    ' TLB 12/9/2014: allow for specifying a particular flavor if ends with a letter
                                    ' - e.g. "-057p" means the pit 57, so "-057p" for "CL3-067" would become "CL-057"
                                    ' - or "-c" means the combined, so "-c" for "CL3-067" would become "CL2-067"
                                    strType = UCase(Right(strSymbol, 1))
                                    If IsAlpha(strType) Then
                                        strSymbol = Left(strSymbol, Len(strSymbol) - 1)
                                    Else
                                        strType = ""
                                    End If
                                    ' replace the contract portion (if exists)
                                    If Len(strSymbol) > 1 Then
                                        strSymbol = Parse(strDefaultSymbol, "-", 1) & strSymbol
                                    Else
                                        strSymbol = strDefaultSymbol
                                    End If
                                    ' convert to specified type
                                    Select Case strType
                                    Case "P" ' pit
                                        strSymbol = ConvertFutureSymbol(strSymbol, ePitSymbol)
                                    Case "C" ' combined
                                        strSymbol = ConvertFutureSymbol(strSymbol, eCombinedSymbol)
                                    Case "E" ' electronic
                                        strSymbol = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
                                    Case "S", "D" ' synthetic or day
                                        strSymbol = ConvertFutureSymbol(strSymbol, eSyntheticSymbol)
                                    End Select
                                    strNewToken = Chr(34) & strSymbol & "," & strPeriod & Chr(34)
                                Else ' else just replace with the default symbol
                                    strSymbol = strDefaultSymbol
                                End If
                            End If
                            If Len(strPeriod) = 0 Then
                                strNewToken = Chr(34) & strSymbol & "," & strDefaultPeriod & Chr(34)
                            ElseIf (Not IsIntraday(GetPeriodicity(strDefaultPeriod))) And (IsIntraday(GetPeriodicity(strPeriod))) Then
                                bInvalidSecondaryPeriod = True
                            End If
                        Else
                            Select Case UCase(strNewToken)
                                Case "GC"
                                    strNewToken = Chr(34) & "GC-067," & strDefaultPeriod & Chr(34)
                                Case "TQ"
                                    strNewToken = Chr(34) & "TQ-067," & strDefaultPeriod & Chr(34)
                                Case "DX"
                                    strNewToken = Chr(34) & "DX-067," & strDefaultPeriod & Chr(34)
                                'TLB: can't do Daily,Weekly this way since we don't know the symbol
                                'Case "DAILY"
                                '    strNewToken = Chr(34) & ",Daily" & Chr(34)
                                'Case "WEEKLY"
                                '    strNewToken = Chr(34) & ",Weekly" & Chr(34)
                            End Select
                        End If
                        strNewToken = UCase(Trim(strNewToken))
                        
                        ' if the rebuilt token has changed at all, then replace it and rebuild the expression
                        If strNewToken <> strTokenValue Then
                            strTokenValue = strNewToken
                            astrTokens(lIndex2) = Format(lTokenType, "00") & Format(Len(strTokenValue), "000") & strTokenValue & " "
                            bRebuild = True
                        End If
                        
                        ' then add this "Symbol,BarPeriod" to the bar names (if not already there)
                        bFound = False
                        For lIndex3 = 0 To astrBarNames.Size - 1
                            If UCase(astrBarNames(lIndex3)) = UCase(strTokenValue) Then
                                bFound = True
                                Exit For
                            End If
                        Next lIndex3
                        If bFound = False Then
                            astrBarNames.Add strTokenValue
                        End If
                    End If
                End If
            Next lIndex2
        
            If bRebuild = True Then
                astrExpressions(lIndex) = Join(astrTokens, "~")
            End If
        End If
    Next lIndex
    
    ' Load up the Bars Collection...
    If Not LoadedBars Is Nothing Then
        For lIndex = 0 To astrBarNames.Size - 1
            If Left(astrBarNames(lIndex), 1) = Chr(34) And Right(astrBarNames(lIndex), 1) = Chr(34) Then
                strSymbol = Parse(Replace(astrBarNames(lIndex), Chr(34), ""), ",", 1)
                strPeriod = Parse(Replace(astrBarNames(lIndex), Chr(34), ""), ",", 2)
                
                ' DAJ 02/19/2010: If the symbol is blank (e.g. ",30 Min"), use the default symbol passed in...
                If Len(strSymbol) = 0 Then
                    strSymbol = strDefaultSymbol
                End If
                If Len(strPeriod) = 0 Then
                    If Len(strDefaultPeriod) = 0 Then
                        strPeriod = "Daily"
                    Else
                        strPeriod = strDefaultPeriod
                    End If
                End If
                
                ' DAJ 02/19/2010: Tim nor I can think of a valid reason to allow a non-intraday default period
                ' with an intraday secondary market, so we will just use a blank set of bars here...
                Set Bars = New cGdBars
                If (bLoadBars = False) Or ((Not IsIntraday(GetPeriodicity(strDefaultPeriod))) And (IsIntraday(GetPeriodicity(strPeriod)))) Then
                    SetBarProperties Bars, strSymbol
                    Bars.Size = 0
                Else
                    DM_GetBars Bars, strSymbol, strPeriod, dStartDate, , , , , bIncludeSnapshotData
                End If
                Set LoadedBars(lIndex + 1) = Bars
                'DebugLog astrBarNames(lIndex) & " (Symbol=" & Bars.Prop(eBARS_Symbol) & ", Period = " & Bars.Prop(eBARS_PeriodicityStr) & ", NumBars=" & Str(Bars.Size) & ")"
            End If
        Next lIndex
    End If
    
    MarketsInExpressions = (astrBarNames.Size > 0)

ErrExit:
    Set Bars = Nothing
    Exit Function
    
ErrSection:
    Set Bars = Nothing
    RaiseError "mSysNav.MarketsInExpressions", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PeriodStr
'' Description: Run a period through the Bars to get a consistent period name
'' Inputs:      Period to convert
'' Returns:     Bars Period
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PeriodStr(ByVal strPeriod As String) As String
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary Bars object
    
    Bars.Prop(eBARS_PeriodicityStr) = strPeriod
    PeriodStr = Bars.Prop(eBARS_PeriodicityStr)

ErrExit:
    Set Bars = Nothing
    Exit Function
    
ErrSection:
    Set Bars = Nothing
    RaiseError "mSysNav.PeriodStr", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildTradesHeader
'' Description: Build a header for a trade-by-trade file
'' Inputs:      System Number, System Name, Time Frame, From Date, To Date,
''              Expenses, Symbol
'' Returns:     Trade Header Line
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildTradesHeader(ByVal lSystemNumber As Long, ByVal strSystemName As String, _
                ByVal strTimeFrame As String, ByVal dFromDate As Double, ByVal dToDate As Double, _
                ByVal dExpenses As Double, ByVal strSymbol As String, Optional ByVal lTotalBars As Long = 0&) As String
On Error GoTo ErrSection:

    Dim astrHeader As New cGdArray      ' Array of header information
    Dim Bars As New cGdBars             ' Temporary Bars object
    
    Set astrHeader = New cGdArray
    astrHeader.Create eGDARRAY_Strings
    
    If dToDate = 0# Then dToDate = Date
    
    SetBarProperties Bars, strSymbol

    astrHeader(TradesHdrField(eTradesHeader_SystemNumber)) = Str(lSystemNumber)
    astrHeader(TradesHdrField(eTradesHeader_SystemName)) = strSystemName
    astrHeader(TradesHdrField(eTradesHeader_BarTimeFrame)) = strTimeFrame
    astrHeader(TradesHdrField(eTradesHeader_StartDate)) = Str(Int(dFromDate))
    astrHeader(TradesHdrField(eTradesHeader_EndDate)) = Str(Int(dToDate))
    astrHeader(TradesHdrField(eTradesHeader_TotalBars)) = Str(lTotalBars)
    astrHeader(TradesHdrField(eTradesHeader_Expenses)) = Str(dExpenses)
    astrHeader(TradesHdrField(eTradesHeader_Symbol)) = Bars.Prop(eBARS_Symbol)
    astrHeader(TradesHdrField(eTradesHeader_TickMove)) = Str(Bars.Prop(eBARS_TickMove))
    astrHeader(TradesHdrField(eTradesHeader_TickValue)) = Str(Bars.Prop(eBARS_TickValue))
    astrHeader(TradesHdrField(eTradesHeader_MinMoveInTicks)) = Str(Bars.Prop(eBARS_MinMoveInTicks))
    astrHeader(TradesHdrField(eTradesHeader_Margin)) = Str(Bars.Prop(eBARS_Margin))
    astrHeader(TradesHdrField(eTradesHeader_SecurityType)) = Bars.SecurityType
    astrHeader(TradesHdrField(eTradesHeader_SessionStart)) = Str(Bars.Prop(eBARS_StartTime))
    astrHeader(TradesHdrField(eTradesHeader_SessionEnd)) = Str(Bars.Prop(eBARS_EndTime))
    
    BuildTradesHeader = astrHeader.JoinFields(vbTab)

ErrExit:
    Set astrHeader = Nothing
    Set Bars = Nothing
    Exit Function
    
ErrSection:
    Set astrHeader = Nothing
    Set Bars = Nothing
    RaiseError "mSysNav.BuildTradesHeader", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshLibrary
'' Description: Refresh a library and its functions and rules in memory
'' Inputs:      Library ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshLibrary(ByVal lLibraryID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Did we find the Library?
    Dim lRecord As Long                 ' Record to change in the table
    Dim rs As Recordset                 ' Recordset into the database
    Dim lProcAddr As Long               ' Callback function address
    Dim strDLL As String                ' Path and FileName of the DLL
    Dim lPos As Long
    Dim F As New cFunction              ' Function object
    Dim r As New cRule                  ' Rule object
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] WHERE [LibraryID]=" & Str(lLibraryID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        For lIndex = 0 To g.tblLibrary.NumRecords - 1
            If g.tblLibrary.Num(LibraryField(etblLib_ID), lIndex) = lLibraryID Then
                bFound = True
                lRecord = lIndex
                Exit For
            End If
        Next lIndex
        If bFound = False Then lRecord = g.tblLibrary.NumRecords
        
        If rs!LibraryType = 1 Then
            lProcAddr = FunctionPtrToLong(AddressOf RunVbFunctionCallback)
        Else
            lProcAddr = 0&
        End If
        
        strDLL = Trim(NullChk(rs!Path))
        If Len(strDLL) > 0 Then
            If UCase(strDLL) <> "BUILTIN.DLL" And Not rs!BuiltIn Then
                ' 3rd-party DLL's should be in the "LibraryDLLs" folder
                strDLL = App.Path
                lPos = At(strDLL, "\", -1)
                If lPos > 0 Then
                    strDLL = Left(strDLL, lPos) & "LibraryDLLs\" & Trim(rs!Path)
                Else
                    strDLL = ""
                End If
            End If
        End If
        
        g.tblLibrary.SetRecord Str(rs!LibraryID) & vbTab & rs!LibraryName & vbTab & rs!LibraryType & _
                vbTab & strDLL & vbTab & Str(lProcAddr) & vbTab & Str(rs!LastModified)
            
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] WHERE [LibraryID]=" & Str(lLibraryID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            Set F = New cFunction
            F.FunctionID = rs!FunctionID
            F.Load
            RefreshFunction F
            
            rs.MoveNext
        Loop
        
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] WHERE [LibraryID]=" & Str(lLibraryID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            Set r = New cRule
            r.RuleID = rs!RuleID
            r.Load
            RefreshRule r
            
            rs.MoveNext
        Loop
    End If
    
ErrExit:
    Set rs = Nothing
    Set F = Nothing
    Set r = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    Set F = Nothing
    Set r = Nothing
    RaiseError "mSysNav.RefreshLibrary", eGDRaiseError_Raise
    
End Sub

Public Function TASDllExists() As Boolean
On Error Resume Next

    Dim strFile$, d#

    ' don't need to keep checking if know it already exists
    Static bExists As Boolean
    If Not bExists Then
        If HasModule("TASIND") Then
            strFile = App.Path & "\TASIndicators.DLL"
            bExists = FileExist(strFile)
            If bExists Then
                ' make a dummy call just to make sure the DLL gets loaded while we're set to the right folder
                ChangePath App.Path ' to insure library will load from app.path
                d = TAS_IndicatorValue(0, 0)
                
                ' their newer DLL requires their Authenticator to be running,
                ' so we should start it up if it's not already running
                If FileDate(strFile) > DateSerial(2014, 7, 1) Then
                    ' TLB 10/20/2015: now using their newer authenticator
                    strFile = App.Path & "\TASLaunchPad.EXE"
                    If FileExist(strFile) Then
                        If KillProcess("TAS Launch Pad", True) = 0 Then
                            RunProcess strFile
                        End If
                    Else
                        ' and I'm guessing the original authenticator should now be obsolete?
                        strFile = App.Path & "\TASAuthServer.EXE"
                        If FileExist(strFile) Then
                            If KillProcess("TAS Authenticator", True) = 0 Then
                                RunProcess strFile
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    TASDllExists = bExists
    
    If m.TASResults Is Nothing Then
        Set m.TASResults = New cGdTree
    End If
    
End Function

Private Sub TASDeleteResults(ByVal nIndId&)

    Dim strKey$
    Dim t As cGdTable

    If nIndId > 0 And Not m.TASResults Is Nothing Then
        strKey = "TAS" & Str(nIndId)
        If m.TASResults.Exists(strKey) Then
            Set t = m.TASResults(strKey)
            t.Clear
            m.TASResults.Remove strKey
            Set t = Nothing
        End If
    End If

End Sub

Private Function TASResultsTable(ByVal nIndId&, Optional ByVal nNumResultsToCreate& = 0) As cGdTable

    Dim strKey$, i&
    Dim tResults As cGdTable
    
    If m.TASResults Is Nothing Then
        Set m.TASResults = New cGdTree
    End If

    If nIndId > 0 Then
        strKey = "TAS" & Str(nIndId)
        If nNumResultsToCreate = 0 Then
            ' get existing table
            If m.TASResults.Exists(strKey) Then
                Set tResults = m.TASResults(strKey)
            End If
        Else
            ' create new table
            Set tResults = New cGdTable
            For i = 0 To nNumResultsToCreate - 1
                tResults.CreateField eGDARRAY_Doubles, i
            Next
            If m.TASResults.Exists(strKey) Then
                m.TASResults.Remove strKey
            End If
            m.TASResults.Add tResults, strKey
        End If
    End If
    
    Set TASResultsTable = tResults

End Function

Public Function TASGetResultsTable(ByVal nIndId&) As cGdTable
    Set TASGetResultsTable = TASResultsTable(nIndId)
End Function

Private Sub TASPasswordError()

    ' only show the message once
    Static bAlreadyShown As Boolean
    If Not bAlreadyShown Then
        bAlreadyShown = True
        InfBox "Installation error with TAS_Indicators.DLL -- please contact TAS technical support.", "!", , "DLL Error"
    End If
    
End Sub

Private Function TAS_Results(hArgs As Long) As Long
    
    'Setup Args and error-handling
    On Error GoTo RunError
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    
    'Declare all your "arguments"
    Dim i&, s$, d#, rc&, hArrayR&, nSize&, nFromBar&, nResultID&, nIndId&
    Dim aIndID As New cGdArray, aResults As New cGdArray
    Dim tResults As cGdTable
    
    ' get nFromBar (greater than 0 when just need recalc for last bar),
    ' and if -2 then just destroy any memory allocated for this instance
    Args.ArgsHandle = hArgs
    nFromBar = Args.FromBar
    If nFromBar = -2 Then
        Args.InstanceMemPtr = 0
        Exit Function
    End If
    If nFromBar < 0 Then nFromBar = 0
    
    'Get each argument (from object passed by engine)
    Args.GetArg aResults
    Args.GetArg aIndID
    Args.GetArg d
    nResultID = d
    
    If Not aIndID.IsConstantValue Then
        nResultID = -1 ' is invalid
    Else
        nIndId = aIndID(aIndID.Size - 1)
        If nIndId <= 0 Then
            nIndId = aIndID(0)
        End If
        hArrayR = 0
        Set tResults = TASResultsTable(nIndId)
        If Not tResults Is Nothing Then
            If nResultID < tResults.NumFields Then
                hArrayR = tResults.FieldArrayHandle(nResultID)
            End If
        End If
    End If

    'Check for arguments error
    If Args.Error <> 0 Then
        TAS_Results = Args.Error
        strErrMsg = "Error in TAS_Results" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' check for valid inputs
    If nResultID < 0 Or Not TASDllExists Then
        TAS_Results = -901
        Exit Function
    End If
       
    If hArrayR = 0 Then
        'If rc = -1 Then TASPasswordError
        TAS_Results = -999 '(unexpected error)
    Else
        ' initialize things
        TAS_Results = 0 ' success
        gdCopy aResults.ArrayHandle, hArrayR
    End If
    
    Exit Function
    
RunError:
    TAS_Results = -999 '(unexpected error)
    Exit Function
End Function

#If 0 Then
Public Function TAS_GetResult(ByVal nTASIndID&, ByVal nResultID&, ByVal nBar&) As Double

    Dim dResult#
    Dim tResults As cGdTable
        
    dResult = kNullData
    Set tResults = TASResultsTable(nTASIndID)
    If Not tResults Is Nothing Then
        dResult = tResults.Num(nResultID, nBar)
        Set tResults = Nothing
    End If
    
    TAS_GetResult = dResult

End Function
#End If

' returns the IndicatorID (except "RATIO")
Private Function TAS_CalcIndicator(hArgs As Long, ByVal strIndName$) As Long
    
    'Setup Args and error-handling
    On Error GoTo RunError
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    
    'Declare all your "arguments"
    Dim i&, s$, d#, rc&, nSize&, nFromBar&, nLastBar&, iNumResults&, nIgnoreBeforeBar&
    Dim dInput#, dResult#, dFirstNonzero#, dCompare#, dElapsedTime#
    Dim bGood As Boolean, bIgnoreLastDataBar As Boolean
    Dim Bars As New cGdBars, aResults As New cGdArray
    Dim tResults As cGdTable
    Dim aParms As New cGdArray
    
    Dim nIndId&, nBar&, nYYYYMMDD&, nHHMM&, dOpen#, dHigh#, dLow#, dClose#, dVol#, dOI#
    Dim dLength#, dMinSignal#, dMapLength#, dRangeBars#, dAvgBars#
    Dim dRawLen#, dSmoothed#, dMABars#, dSLength#, dSDevs#, dSATRs#, dPLen#
    Dim iColor&, bPeak As Boolean
    Dim dSignalStrength#, dBoxesLookback#, dBoxesBack#, dATRLookback#
    Dim dCompThresh#, dLongTrades#, dShortTrades#, dATRStop#, dATRProfits#, dATRLength#, dStopsOnTouch#
    Dim nPlot4Col&, nPlot2&, nPlot3&, nPlot4&, nPlot31&, nArrow&
    Dim dNumBins#, dBarsBackStart#, dBarsBackEnd#
    
    ' get nFromBar (greater than 0 when just need recalc for last bar),
    ' and if -2 then just destroy any memory allocated for this instance
    Args.ArgsHandle = hArgs
    nFromBar = Args.FromBar
    If nFromBar = -2 Then
        If Args.InstanceMemPtr <> 0 Then
            nIndId = gdGetNum(Args.InstanceMemPtr, 0)
            If IsIDE Then frmTest.AddList "Destroy " & strIndName & " " & Str(nIndId)
            TASDeleteResults nIndId
            gdDestroyArray Args.InstanceMemPtr
            Args.InstanceMemPtr = 0
        End If
        Exit Function
    End If
    If nFromBar < 0 Then nFromBar = 0
    ' allocate array to hold stuff for this instance
    If Args.InstanceMemPtr = 0 Then
        Args.InstanceMemPtr = gdCreateArray(eGDARRAY_Doubles, 0, 0)
    End If
    
    'Get each argument (from object passed by engine)
    bIgnoreLastDataBar = False
    nIgnoreBeforeBar = 0
    aParms.Create eGDARRAY_Doubles
    Args.GetArg aResults
    Args.GetArg Bars
    Select Case strIndName
    Case "STATICPCL"
        iNumResults = 10
    
    Case "FLOATPCL"
        Args.GetArg dMinSignal ' 1
        Args.GetArg dLength ' 8
        aParms.Add dLength
        aParms.Add dMinSignal
        iNumResults = 10
    
    Case "SWINGRSI" ' Boxes
        Args.GetArg dMinSignal ' 2
        Args.GetArg dLength ' 7
        Args.GetArg dMapLength ' 7
        aParms.Add dLength
        aParms.Add dMinSignal
        aParms.Add dMapLength
        iNumResults = 3
    
    Case "COMBO" ' Navigator
        Args.GetArg dRawLen '= 25
        Args.GetArg dSmoothed '= 13
        Args.GetArg dMABars '= 5
        Args.GetArg dSLength '= 10
        Args.GetArg dSDevs '= 2
        Args.GetArg dSATRs '= 1.5
        Args.GetArg dPLen '= 30
        aParms.Add dRawLen
        aParms.Add dSmoothed
        aParms.Add 1#
        aParms.Add dMABars
        aParms.Add dSLength
        aParms.Add dSDevs
        aParms.Add dSATRs
        aParms.Add dPLen
        iNumResults = 6 '4
        nIgnoreBeforeBar = dRawLen
        bIgnoreLastDataBar = True
    
    Case "VEGA"
        Args.GetArg dSignalStrength ' 2
        Args.GetArg dBoxesLookback ' 7
        Args.GetArg dBoxesBack ' 3
        Args.GetArg dATRLookback ' 7
        Args.GetArg dATRStop ' 1.15
        Args.GetArg dATRProfits ' 0
        Args.GetArg dATRLength ' 27
        Args.GetArg dStopsOnTouch ' 0
        Args.GetArg dCompThresh ' 100
        Args.GetArg dLongTrades ' 1
        Args.GetArg dShortTrades ' 1
        aParms.Add dSignalStrength
        aParms.Add dBoxesLookback
        aParms.Add 1#  ' dDontChange
        aParms.Add dATRLookback
        aParms.Add dBoxesBack
        aParms.Add dCompThresh
        aParms.Add dLongTrades
        aParms.Add dShortTrades
        aParms.Add dATRStop
        aParms.Add dATRProfits
        aParms.Add dATRLength
        aParms.Add dStopsOnTouch
        iNumResults = 17
        nIgnoreBeforeBar = 30
    
    Case "RATIO"
        Args.GetArg dRangeBars ' 10
        Args.GetArg dAvgBars ' 3
        'Args.GetArg dMABars ' 3
        dMABars = 3
        aParms.Add dRangeBars
        aParms.Add dAvgBars
        aParms.Add dMABars
        iNumResults = 2
        nIgnoreBeforeBar = dRangeBars
        bIgnoreLastDataBar = True
        
    Case "MARKETMAP"
        Args.GetArg dBarsBackStart
        Args.GetArg dBarsBackEnd
        Args.GetArg dNumBins
        
        ' this indicator is handled completely different from the rest
        If nFromBar = 0 Then
            Set tResults = TAS_CalcMarketMap(Bars, dBarsBackStart, dBarsBackEnd, dNumBins)
            s = "TASMarketMap" & vbTab & Bars.Prop(eBARS_Symbol) & vbTab
            For i = 0 To tResults.NumRecords - 1
                Args.AddDrawingCommand s & tResults.GetRecord(i, ";") 'X1, Y1, X2, Y2, Color
            Next
        End If
        Set tResults = Nothing
        Exit Function
    End Select
             
    'Check for arguments error
    If Args.Error <> 0 Then
        TAS_CalcIndicator = Args.Error
        strErrMsg = "Error in " & strIndName & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' check for valid inputs
    If Not TASDllExists Then
        TAS_CalcIndicator = -901
        Exit Function
    End If
    
    ' get stuff from Instance Array
    Set tResults = Nothing
    If gdGetSize(Args.InstanceMemPtr) > 0 And nFromBar > 0 Then
        ' get nIndID
        nIndId = gdGetNum(Args.InstanceMemPtr, 0)
        Set tResults = TASResultsTable(nIndId)
        ' but if nFromBar < last bar we called SetBar for, then we must start over
        If nFromBar < gdGetNum(Args.InstanceMemPtr, 1) Then
            ' but when new ticks for same bar then nFromBar is always set to prior bar,
            ' and we can just start from the current bar so won't have to reinit their function
            ' (since results for prior bar are already stored in the results table)
            If nFromBar = gdGetNum(Args.InstanceMemPtr, 1) - 1 And nFromBar > 10 Then
                nFromBar = gdGetNum(Args.InstanceMemPtr, 1)
            Else
                nFromBar = 0
            End If
        End If
    End If
    If nFromBar <= 0 Or tResults Is Nothing Then
        nFromBar = 0
        nIndId = 0
        Set tResults = Nothing
        gdSetSize Args.InstanceMemPtr, 0, 0
    End If
            
    ' determine last bar to calculate
    nSize = Bars.Size
    For nBar = Bars.Size - 1 To 0 Step -1
        If Bars(eBARS_Close, nBar) <> kNullData Then
            If bIgnoreLastDataBar Then 'And Bars.SessionDate(nBar) > LastDailyDownload Then
                nLastBar = nBar - 1
            Else
                nLastBar = nBar
            End If
            Exit For
        End If
    Next
    
    ' init the function (only if need to)
    If nIndId <= 0 Then
        s = Bars.Prop(eBARS_Symbol)
        dElapsedTime = gdTickCount
        nIndId = TAS_IndicatorInit("eSignal", strIndName, s)
        If IsIDE Then
            dElapsedTime = gdTickCount - dElapsedTime
            frmTest.AddList "Init " & strIndName & " for " & s & " = " & Str(nIndId) & " (" & Str(Int(dElapsedTime)) & " ms)"
        End If
    End If
    If nIndId <= 0 Then
        'If rc = -1 Then TASPasswordError
        TAS_CalcIndicator = -999 '(unexpected error)
    Else
        ' initialize things
        gdSetNum Args.InstanceMemPtr, 0, nIndId ' store nIndID so won't have to keep reinitializing during streaming
        TAS_CalcIndicator = 0 ' success
        bGood = False
        
        ' RATIO is the exception: just store the actual value in the Results array
        ' (instead of the IndicatorID like the rest of them)
        If strIndName = "RATIO" Then
            aResults.Size = nSize
        Else
            aResults.MakeConstantArray nIndId, nSize
        End If
        
        If nFromBar = 0 Then
            ' set all the Indicator parameters
            For i = 0 To aParms.Size - 1
                rc = TAS_IndicatorSetParameter(nIndId, i, aParms(i))
            Next
            ' create the Results table
            Set tResults = TASResultsTable(nIndId, iNumResults)
            If aResults.IsConstantValue Then
                ' to store actual return values (rows = each bar, columns = each return value for each bar)
                tResults.NumRecords = nSize
            Else
                ' don't need to store anything if returning the actual values (e.g. RATIO)
                tResults.NumRecords = 0
            End If
        Else
            bGood = True
        End If
        
        ' call TAS function for each valid bar
        If nFromBar = 0 Then
            dElapsedTime = gdTickCount
        Else
            dElapsedTime = 0
        End If
        For nBar = nFromBar To nSize - 1
            dResult = kNullData
            dClose = Bars(eBARS_Close, nBar)
            If dClose <> kNullData And nBar <= nLastBar Then
                If 0 Then
                    d = Bars(eBARS_DateTime, nBar)
                    nYYYYMMDD = JulToLong(Int(d), 1)
                    nHHMM = Hour(d) * 100 + Minute(d)
                Else
                    ' just pass session date for date (and time isn't used)
                    d = Bars.SessionDate(nBar)
                    nYYYYMMDD = JulToLong(d, 1)
                    nHHMM = 0
                End If
                dOpen = Bars(eBARS_Open, nBar)
                dHigh = Bars(eBARS_High, nBar)
                dLow = Bars(eBARS_Low, nBar)
                dVol = Bars(eBARS_Vol, nBar)
                If dVol <= 0 Then dVol = 1
                dOI = Bars(eBARS_OI, nBar)
                If dOI < 0 Then dOI = 0
                
                gdSetNum Args.InstanceMemPtr, 1, nBar ' store last bar# for SetBar (since can't go backwards)
                i = TAS_IndicatorSetBar(nIndId, nBar, nYYYYMMDD, nHHMM, dOpen, dHigh, dLow, dClose, dVol, dOI)
                
                ' the initial results are not yet "good" until the first non-zero value has changed
                If Not bGood Then
                    If strIndName = "VEGA" Then
                        d = Round(TAS_IndicatorValue(nIndId, 12), 10)
                    Else
                        d = Round(TAS_IndicatorValue(nIndId, 0), 10)
                    End If
                    If dFirstNonzero = 0 Then
                        dFirstNonzero = d
                    ElseIf d <> dFirstNonzero Then
                        bGood = True
                    End If
                End If
                If bGood And nBar >= nIgnoreBeforeBar Then
                    Select Case strIndName
                    Case "RATIO"
                        ' RATIO is the exception: just store the actual value in the Results array
                        ' (instead of the IndicatorID like the rest of them)
                        dResult = TAS_IndicatorValue(nIndId, 0)
                        aResults(nBar) = dResult
                    
                    Case "COMBO"
                        bPeak = False
                        'For i = tResults.NumFields - 1 To 0 Step -1
                        For i = 3 To 0 Step -1
                            dResult = TAS_IndicatorValue(nIndId, i)
                            If i = 3 And dResult <> 0 Then
                                bPeak = True
                            ElseIf i = 2 Then
                                If dResult <> 0 Then dResult = 0.00001 ' sideways
                            ElseIf i = 0 Then
                                ' TLB 1/7/2015: set green or red dot based on whether above/below the MA
                                If dResult > tResults.Num(1, nBar) Then
                                    tResults.Num(4, nBar) = dResult ' green dot
                                ElseIf dResult < tResults.Num(1, nBar) Then
                                    tResults.Num(5, nBar) = dResult ' red dot
                                End If
                                ' value will be compared with 2 bars ago
                                dCompare = tResults.Num(0, nBar - 2)
                                If dCompare = kNullData Then
                                    dCompare = dResult
                                End If
                                ' a hack for coloring: set the 10th digit after decimal point to color#
                                iColor = 0
                                If bPeak Then
                                    iColor = 5
                                ElseIf dResult >= 0 Then
                                    If dResult >= dCompare Then
                                        iColor = 1
                                    Else
                                        iColor = 2
                                    End If
                                Else
                                    If dResult >= dCompare Then
                                        iColor = 3
                                    Else
                                        iColor = 4
                                    End If
                                End If
                                dResult = Round(dResult, 9)
                                If dResult < 0 Then
                                    dResult = dResult - 0.0000000001 * iColor
                                Else
                                    dResult = dResult + 0.0000000001 * iColor
                                End If
                            End If
                            tResults.Num(i, nBar) = dResult
                        Next
                        
                    Case "VEGA"
                        For i = 0 To tResults.NumFields - 1
                            If i < 16 Then
                                dResult = TAS_IndicatorValue(nIndId, i)
                                Select Case i
                                Case 2
                                    nPlot3 = kNullData
                                Case 6
                                    nPlot2 = dResult
                                Case 7
                                    nPlot3 = dResult
                                Case 8
                                    nPlot31 = dResult
                                Case 9
                                    nPlot4 = dResult
                                Case 12
                                    nPlot4Col = dResult
                                Case 14
                                    nArrow = dResult
                                End Select
                            Else
                                ' determine color for OHLC bar
                                dResult = 2 'RGB(128, 128, 128) ' gray (default)
                                Select Case nPlot4Col
                                Case 1
                                    dResult = 1 'RGB(255, 165, 0)
                                Case 2
                                    'dResult = 2 'RGB(128, 128, 128) ' gray
                                    tResults.Num(7, nBar) = kNullData
                                Case 3
                                    If nArrow = 1 Then
                                        dResult = 4 'RGB(0, 192, 0) ' lime
                                    ElseIf nArrow = 2 Then
                                        dResult = 5 'vbRed
                                    End If
                                Case 4
                                    dResult = 4 'RGB(0, 192, 0) ' lime
                                Case 5
                                    dResult = 5 'vbRed
                                Case 6
                                    dResult = 6 'vbMagenta
                                End Select
                            End If
                            tResults.Num(i, nBar) = dResult
                        Next
                        
                    Case Else
                        For i = 0 To tResults.NumFields - 1
                            dResult = TAS_IndicatorValue(nIndId, i)
                            tResults.Num(i, nBar) = dResult
                        Next
                    End Select
                End If
            End If
        Next
        If IsIDE And dElapsedTime > 0 Then
            dElapsedTime = gdTickCount - dElapsedTime
            s = Bars.Prop(eBARS_Symbol)
            frmTest.AddList "Calc " & strIndName & " for " & s & " = " & Str(nIndId) & " (" & Str(Int(dElapsedTime)) & " ms)"
        End If
    End If
    
    Exit Function
    
RunError:
    TAS_CalcIndicator = -999 '(unexpected error)
    Exit Function
End Function

Public Function TAS_CalcMarketMap(Bars As cGdBars, ByVal nBarsBackStart&, ByVal nBarsBackEnd&, ByVal nNumBins&) As cGdTable

    Dim s$, i&, d#, rc&, nStartBar&, nEndBar&, nColor&
    Dim nIndId&, nBar&, nYYYYMMDD&, nHHMM&, dOpen#, dHigh#, dLow#, dClose#, dVol#, dOI#
    Dim hMin#, hMax#, dP#, X1#, X2#, hW#, mW#, tg#, Y#, wID#, b1#, b2#, ff#
    Dim bReverseDirection As Boolean
    Dim aLines As New cGdTable
        
    ' get start and end bar#'s
    ' NOTE: #BarsBackStart could be either > or < #BarsBackEnd (display just flips)
    nStartBar = -1
    For nBar = Bars.Size - 1 To 0 Step -1
        ' find last data bar
        If Bars(eBARS_Close, nBar) <> kNullData Then
            If nBarsBackStart < nBarsBackEnd Then
                nEndBar = nBar - nBarsBackStart
                bReverseDirection = True
            Else
                nEndBar = nBar - nBarsBackEnd
            End If
            nEndBar = nEndBar - 1 ' since their code actually uses data from the previous bar
            For nEndBar = nEndBar To 0 Step -1
                If Bars(eBARS_Close, nEndBar) <> kNullData Then
                    nStartBar = nEndBar - Abs(nBarsBackStart - nBarsBackEnd)
                    Exit For
                End If
            Next
            Exit For
        End If
    Next

    ' initialize their indicator (if parms are all good)
    nIndId = 0
    If nNumBins > 0 And nStartBar >= 0 And nEndBar > nStartBar And nBarsBackStart >= 0 And nBarsBackEnd >= 0 Then
        s = Bars.Prop(eBARS_Symbol)
        d = gdTickCount
        nIndId = TAS_IndicatorInit("eSignal", "MARKETMAP", s)
        If IsIDE Then
            d = gdTickCount - d
            frmTest.AddList "Init MarketMap" & " for " & s & " = " & Str(nIndId) & " (" & Str(Int(d)) & " ms)"
        End If
    End If
    
    If nIndId > 0 Then
        b1 = Abs(nBarsBackStart - nBarsBackEnd)
        b2 = 0
            
        ' init their variables (just like they do in their eSignal code)
        If b1 < 0 Then b1 = 0
        If b2 < 0 Then b2 = 0
        If b1 = b2 Then b2 = b1 + 1
        mW = Abs(b1 - b2)
        If mW < 1 Then mW = 1
        If nNumBins < 1 Then
            nNumBins = 1
        ElseIf nNumBins > 200 Then
            nNumBins = 200 ' max allowed
        End If
        rc = TAS_IndicatorSetParameter(nIndId, 0, nNumBins)
        rc = TAS_IndicatorSetParameter(nIndId, 1, b1)
        rc = TAS_IndicatorSetParameter(nIndId, 2, b2)
        
        ' create Table of lines to draw
        aLines.CreateField eGDARRAY_Doubles, 0, "X1"
        aLines.CreateField eGDARRAY_Doubles, 1, "Y1"
        aLines.CreateField eGDARRAY_Doubles, 2, "X2"
        aLines.CreateField eGDARRAY_Doubles, 3, "Y2"
        aLines.CreateField eGDARRAY_Longs, 4, "Color"
        aLines.NumRecords = nNumBins
    
        ' call TAS function for each valid bar
        For nBar = nStartBar To nEndBar
            dClose = Bars(eBARS_Close, nBar)
            If dClose <> kNullData Then
                If 0 Then
                    d = Bars(eBARS_DateTime, nBar)
                    nYYYYMMDD = JulToLong(Int(d), 1)
                    nHHMM = Hour(d) * 100 + Minute(d)
                Else
                    ' just pass session date for date (and time isn't used)
                    d = Bars.SessionDate(nBar)
                    nYYYYMMDD = JulToLong(d, 1)
                    nHHMM = 0
                End If
                dOpen = Bars(eBARS_Open, nBar)
                dHigh = Bars(eBARS_High, nBar)
                dLow = Bars(eBARS_Low, nBar)
                dVol = Bars(eBARS_Vol, nBar)
                If dVol <= 0 Then dVol = 1
                dOI = Bars(eBARS_OI, nBar)
                If dOI < 0 Then dOI = 0
                
                If nBar = nEndBar Then
                    i = TAS_IndicatorSetBar(nIndId, nBar, nYYYYMMDD, nHHMM, dOpen, dHigh, dLow, dClose, dVol, dOI)
                Else
                    i = TAS_IndicatorSetBarNoCalc(nIndId, nBar, nYYYYMMDD, nHHMM, dOpen, dHigh, dLow, dClose, dVol, dOI)
                End If
            End If
        Next
        
        ' get their info for each line to draw
        hMin = TAS_IndicatorValue(nIndId, 0) ' lowest low
        hMax = TAS_IndicatorValue(nIndId, 1) ' highest high
        dP = (hMax - hMin) / nNumBins ' distance between each bin
        For i = 1 To nNumBins
            hW = TAS_IndicatorValue(nIndId, 100 + i - 1)
            tg = TAS_IndicatorValue(nIndId, 500 + i - 1)
            Y = hMin + (dP * (i - 1))
            wID = Round(hW * mW) ' length of line as # of bars
            If wID < 1 Then wID = 1
            
            If tg = 1 Then
                nColor = RGB(0, 192, 0) ' vbGreen
            ElseIf tg = 2 Then
                nColor = RGB(210, 200, 125) 'khaki
            Else
                ff = Round(hW * 255)
                If ff < 0 Then
                    ff = 0
                ElseIf ff > 255 Then
                    ff = 255
                End If
                nColor = RGB(ff, 0, 255 - ff)
            End If
            
            If bReverseDirection Then
                ' displays from right to left
                'X1 = -(b1 + wID)
                'X2 = -b1
                X2 = nEndBar
                X1 = X2 - wID
            Else
                ' displays from left to right
                'X1 = -b1
                'X2 = -(b1 - wID)
                X1 = nStartBar
                X2 = X1 + wID
            End If
            
            ' draw line from x1,y to x2,y of nColor
            ' (add 1 to X for drawing since their code actually uses data from the previous bar)
            aLines.Num(0, i - 1) = X1 + 1
            aLines.Num(1, i - 1) = Y
            aLines.Num(2, i - 1) = X2 + 1
            aLines.Num(3, i - 1) = Y
            aLines.Num(4, i - 1) = nColor
        Next
    End If

    Set TAS_CalcMarketMap = aLines
    
End Function

Public Function JurikDllExists() As Boolean

    ' don't need to keep checking if know it already exists
    Static bExists As Boolean
    If Not bExists Then
        bExists = FileExist(WinSysPath & "JRS_UT.DLL")
    End If
    JurikDllExists = bExists
    
End Function

Private Sub JurikPasswordError()

    ' only show the message once
    Static bAlreadyShown As Boolean
    If Not bAlreadyShown Then
        bAlreadyShown = True
        InfBox "Password/Installation error with JRS_UT.DLL -- please contact Jurik Research technical support.", "!", , "JRS_UT.DLL Error"
    End If
    
End Sub

Private Function Jurik_JMA(hArgs As Long) As Long
    
    'Setup Args and error-handling
    On Error GoTo RunError
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    Args.ArgsHandle = hArgs
    
    'Declare all your "arguments"
    Dim i&, rc&, hArrayI&, hArrayR&, nSize&
    Dim iSeriesID&, dSmooth#, dPhase#, dInput#, dResult#, bGood As Boolean
    Dim aInput As New cGdArray, aResults As New cGdArray

    'Get each argument (from object passed by engine)
    Args.GetArg aResults
    Args.GetArg aInput
    Args.GetArg dSmooth ' "length": positive values
    Args.GetArg dPhase  ' range: -100 to +100
    
    'Check for arguments error
    If Args.Error <> 0 Then
        Jurik_JMA = Args.Error
        strErrMsg = "Error in Jurik_JMA" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' check for valid inputs
    If dSmooth <= 0 Or dPhase < -100 Or dPhase > 100 Or Not JurikDllExists Then
        Jurik_JMA = -901
        Exit Function
    End If
    
    ' initialize things
    Jurik_JMA = 0 ' success
    iSeriesID = 0
    nSize = aInput.Size
    hArrayI = aInput.ArrayHandle
    hArrayR = aResults.ArrayHandle
    bGood = False
    aResults.Size = nSize
    
    ' call Jurik function for each valid bar
    For i = 0 To nSize - 1
        dInput = gdGetNum(hArrayI, i)
        If dInput <> kNullData Then
            dResult = kNullData
            rc = JMAUT(dInput, dSmooth, dPhase, dResult, 0, iSeriesID, 0)
            If rc <> 0 Then
                If rc = -1 Then JurikPasswordError
                Jurik_JMA = -999 '(unexpected error)
                Exit For
            Else
                ' the initial results are not yet "good" until they differ from the input
                If dResult <> dInput Then bGood = True
                If bGood Then
                    gdSetNum hArrayR, i, dResult  'aResults(i) = dResult
                End If
            End If
        End If
    Next
    
    ' need to call it this way in order to destroy internal memory
    rc = JMAUT(0, 0, 0, 0, 1, iSeriesID, 0)
    Exit Function
    
RunError:
    Jurik_JMA = -999 '(unexpected error)
    Exit Function
End Function

Private Function Jurik_VEL(hArgs As Long) As Long
    
    'Setup Args and error-handling
    On Error GoTo RunError
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    Args.ArgsHandle = hArgs
    
    'Declare all your "arguments"
    Dim i&, rc&, hArrayI&, hArrayR&, nSize&
    Dim iSeriesID&, nDepth&, dInput#, dResult#, bGood As Boolean
    Dim aInput As New cGdArray, aResults As New cGdArray

    'Get each argument (from object passed by engine)
    Args.GetArg aResults
    Args.GetArg aInput
    Args.GetArg nDepth ' size of moving window
    
    'Check for arguments error
    If Args.Error <> 0 Then
        Jurik_VEL = Args.Error
        strErrMsg = "Error in Jurik_VEL" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' check for valid inputs
    If nDepth <= 0 Or Not JurikDllExists Then
        Jurik_VEL = -901
        Exit Function
    End If
    
    ' initialize things
    Jurik_VEL = 0 ' success
    iSeriesID = 0
    nSize = aInput.Size
    hArrayI = aInput.ArrayHandle
    hArrayR = aResults.ArrayHandle
    bGood = False
    aResults.Size = nSize
    
    ' call Jurik function for each valid bar
    For i = 0 To nSize - 1
        dInput = gdGetNum(hArrayI, i)
        If dInput <> kNullData Then
            dResult = kNullData
            rc = VELUT(dInput, nDepth, dResult, 0, iSeriesID, 0)
            If rc <> 0 Then
                If rc = -1 Then JurikPasswordError
                Jurik_VEL = -999 '(unexpected error)
                Exit For
            Else
                ' the initial results are not yet "good" until they are non-zero
                If dResult <> 0 Then bGood = True
                If bGood Then
                    gdSetNum hArrayR, i, dResult  'aResults(i) = dResult
                End If
            End If
        End If
    Next
    
    ' need to call it this way in order to destroy internal memory
    rc = VELUT(0, 0, 0, 1, iSeriesID, 0)
    Exit Function
    
RunError:
    Jurik_VEL = -999 '(unexpected error)
    Exit Function
End Function

Private Function Jurik_RSX(hArgs As Long) As Long
    
    'Setup Args and error-handling
    On Error GoTo RunError
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    Args.ArgsHandle = hArgs
    
    'Declare all your "arguments"
    Dim i&, rc&, hArrayI&, hArrayR&, nSize&
    Dim iSeriesID&, dSmooth#, dInput#, dResult#, bGood As Boolean
    Dim aInput As New cGdArray, aResults As New cGdArray

    'Get each argument (from object passed by engine)
    Args.GetArg aResults
    Args.GetArg aInput
    Args.GetArg dSmooth ' "length": positive values
    
    'Check for arguments error
    If Args.Error <> 0 Then
        Jurik_RSX = Args.Error
        strErrMsg = "Error in Jurik_RSX" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' check for valid inputs
    If dSmooth <= 0 Or Not JurikDllExists Then
        Jurik_RSX = -901
        Exit Function
    End If
    
    ' initialize things
    Jurik_RSX = 0 ' success
    iSeriesID = 0
    nSize = aInput.Size
    hArrayI = aInput.ArrayHandle
    hArrayR = aResults.ArrayHandle
    bGood = False
    aResults.Size = nSize
    
    ' call Jurik function for each valid bar
    For i = 0 To nSize - 1
        dInput = gdGetNum(hArrayI, i)
        If dInput <> kNullData Then
            dResult = kNullData
            rc = RSXUT(dInput, dSmooth, dResult, 0, iSeriesID, 0)
            If rc <> 0 Then
                If rc = -1 Then JurikPasswordError
                Jurik_RSX = -999 '(unexpected error)
                Exit For
            Else
                ' the initial results are not yet "good" until they differ from 50
                If dResult <> 50 Then bGood = True
                If bGood Then
                    gdSetNum hArrayR, i, dResult  'aResults(i) = dResult
                End If
            End If
        End If
    Next
    
    ' need to call it this way in order to destroy internal memory
    rc = RSXUT(0, 0, 0, 1, iSeriesID, 0)
    Exit Function
    
RunError:
    Jurik_RSX = -999 '(unexpected error)
    Exit Function
End Function

' iMode: 0 = DMX,  1 = DMX Plus,  -1 = DMX Minus
Private Function Jurik_DMX(hArgs As Long, Optional ByVal iMode% = 0) As Long
    
    'Setup Args and error-handling
    On Error GoTo RunError
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    Args.ArgsHandle = hArgs
    
    'Declare all your "arguments"
    Dim i&, rc&, hArrayH&, hArrayL&, hArrayC&, hArrayR&, nSize&
    Dim iSeriesID&, dLength#, dHigh#, dLow#, dClose#, dResult#, d1#, d2#, bGood As Boolean
    Dim Bars As New cGdBars, aResults As New cGdArray

    'Get each argument (from object passed by engine)
    Args.GetArg aResults
    Args.GetArg Bars
    Args.GetArg dLength ' "length": positive values
    
    'Check for arguments error
    If Args.Error <> 0 Then
        Jurik_DMX = Args.Error
        strErrMsg = "Error in Jurik_DMX" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' check for valid inputs
    If dLength <= 0 Or Not JurikDllExists Then
        Jurik_DMX = -901
        Exit Function
    End If
    
    ' initialize things
    Jurik_DMX = 0 ' success
    iSeriesID = 0
    nSize = Bars.Size
    hArrayH = Bars.ArrayHandle(eBARS_High)
    hArrayL = Bars.ArrayHandle(eBARS_Low)
    hArrayC = Bars.ArrayHandle(eBARS_Close)
    hArrayR = aResults.ArrayHandle
    bGood = False
    aResults.Size = nSize
    
    ' call Jurik function for each valid bar
    For i = 0 To nSize - 1
        dHigh = gdGetNum(hArrayH, i)
        dLow = gdGetNum(hArrayL, i)
        dClose = gdGetNum(hArrayC, i)
        If dHigh <> kNullData And dLow <> kNullData And dClose <> kNullData Then
            dResult = kNullData
            If iMode > 0 Then
                rc = DMXUT(dHigh, dLow, dClose, d1, dResult, d2, dLength, 0, iSeriesID, 0)
            ElseIf iMode < 0 Then
                rc = DMXUT(dHigh, dLow, dClose, d1, d2, dResult, dLength, 0, iSeriesID, 0)
            Else
                rc = DMXUT(dHigh, dLow, dClose, dResult, d1, d2, dLength, 0, iSeriesID, 0)
            End If
            If rc <> 0 Then
                If rc = -1 Then JurikPasswordError
                Jurik_DMX = -999 '(unexpected error)
                Exit For
            Else
                ' the initial results are not yet "good" until they are non-zero
                If dResult <> 0 Then bGood = True
                If bGood Then
                    gdSetNum hArrayR, i, dResult  'aResults(i) = dResult
                End If
            End If
        End If
    Next
    
    ' need to call it this way in order to destroy internal memory
    rc = DMXUT(0, 0, 0, 0, 0, 0, 0, 1, iSeriesID, 0)
    Exit Function
    
RunError:
    Jurik_DMX = -999 '(unexpected error)
    Exit Function
End Function

Private Function Engine_PredLabs(hArgs As Long) As Long
    
    'Setup Args and error-handling
    On Error GoTo RunError
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    Args.ArgsHandle = hArgs
    
    'Declare all your "arguments"
    Dim i&, j&, hArrayC&, hArrayR&, nSessionDate&, d#, dOffSet#, iOffsetCount&
    Dim strFileName$
    Dim Bars As New cGdBars, aResults As New cGdArray

    'Get each argument (from object passed by engine)
    Args.GetArg aResults
    Args.GetArg Bars
    
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_PredLabs = Args.Error
        strErrMsg = "Error in Engine_PredLabs" & ":  " & Args.ErrorMessage
        Exit Function
    End If
    
    ' initialize things
    Engine_PredLabs = 0 ' success
    
    ' load Pred Labs data from file
    strFileName = PrimaryFutureBase(Bars.Prop(eBARS_BaseSymbol))
    strFileName = AddSlash(App.Path) & "PredLabs\" & strFileName & ".dat"
    If GetTextData(aResults, Bars, strFileName, 1, 0) = 0 Then
        ' go back to the first bar where PredLabs data exists
        hArrayC = Bars.ArrayHandle(eBARS_Close)
        hArrayR = aResults.ArrayHandle
        For i = aResults.Size - 1 To 1 Step -1
            If gdGetNum(hArrayR, i) <> kNullData Then
                If gdGetNum(hArrayR, i - 1) = kNullData Then
                    ' if not an index, then auto-adjust the price offset
                    ' (average of the difference in the completed bar's closes prior until 9:45am)
                    dOffSet = 0
                    iOffsetCount = 0
                    If Bars.SecurityType <> "I" Then
                        For j = i To Bars.Size - 1
                            ' exit if bar has not yet completed or if past 9:45am
                            If gdGetNum(hArrayC, j + 1) = kNullData Then Exit For
                            If iOffsetCount > 0 Then '(but make sure at least 1 bar gets used)
                                d = Bars(eBARS_DateTime, j)
                                d = d - Int(d)
                                If d > 585 / 1440# Then Exit For
                            End If
                            dOffSet = dOffSet + gdGetNum(hArrayC, j) - gdGetNum(hArrayR, j)
                            iOffsetCount = iOffsetCount + 1
                        Next
                        If iOffsetCount > 1 Then dOffSet = dOffSet / iOffsetCount
                    End If
                    If gdGetNum(hArrayC, i + 1) = kNullData Then
                        nSessionDate = 0 ' if 2nd bar for day has not started yet, then null the results
                    Else
                        nSessionDate = Bars.SessionDate(i) ' results only valid for this session
                    End If
                    ' then walk forward through this session
                    For j = i To Bars.Size - 1
                        d = gdGetNum(hArrayR, j)
                        If d <> kNullData Then
                            ' set to null if not the same session
                            If Bars.SessionDate(j) <> nSessionDate Then
                                gdSetNum hArrayR, j, kNullData
                            Else ' else offset the data
                                gdSetNum hArrayR, j, d + dOffSet
                            End If
                        End If
                    Next
                    Exit For
                End If
            End If
        Next
    End If
    
    Exit Function
    
RunError:
    Engine_PredLabs = -999 '(unexpected error)
    Exit Function
End Function

' PowerZones(Market1, ZoneID, UserID, Password, ZoneAdjust)
Private Function Engine_PowerZonesData(hArgs As Long) As Long
    
    'Setup Args and error-handling
    On Error GoTo RunError
    Dim strErrMsg As String
    Dim Args As New cGdArgs
    Args.ArgsHandle = hArgs
    
    'Declare all your "arguments"
    Dim iZone&, iBar&, iRec&, dDate#, dPivotDate#, dVal#, dAdjust#, nFromBar&
    Dim strSymbol$, strZoneFile$
    Dim Bars As New cGdBars, aResults As New cGdArray
    Static tPivotData As cGdTable, strPrevArgs$

    If g.bUnloading Then Exit Function

    ' get nFromBar (greater than 0 when just need recalc for last bar)
    nFromBar = Args.FromBar
    If nFromBar = -2 Then Exit Function
    If nFromBar < 0 Then nFromBar = 0

    'Get each argument (from object passed by engine)
    Args.GetArg aResults
    Args.GetArg Bars
    Args.GetArg iZone
    Args.GetArg strZoneFile
    Args.GetArg dAdjust
    
    'Check for arguments error
    If Args.Error <> 0 Then
        Engine_PowerZonesData = Args.Error
        strErrMsg = "Error in Engine_PowerZonesData" & ":  " & Args.ErrorMessage
        Exit Function
    End If
       
    ' initialize things
    Engine_PowerZonesData = 0 ' success
    
    ' get the pivot data for this symbol
    strSymbol = Bars.Prop(eBARS_Symbol)
    
#If 1 Then
    strZoneFile = LCase(Trim(strZoneFile))
    If Len(strZoneFile) = 0 Or Left(strZoneFile, 1) = "(" Then
        strZoneFile = frmPowerZones.ZoneFileForSymbol(strSymbol)
    End If
    Set tPivotData = frmPowerZones.GetZoneData(strZoneFile)
#Else
    If strPrevArgs <> strSymbol & vbTab & strUser Or (tPivotData Is Nothing) Then
        strPrevArgs = strSymbol & vbTab & strUser
        Set tPivotData = GetPivotFarmData(strSymbol, strUser)
        nFromBar = 0
    End If
#End If

    ' align pivot data with session dates
    ' (more efficient to start at end and work backwards)
    If tPivotData.NumRecords > 0 And tPivotData.NumFields > 1 And iZone >= 1 Then
        iRec = tPivotData.NumRecords - 1
        For iBar = Bars.Size - 1 To nFromBar Step -1
            dVal = kNullData
            dDate = Bars.SessionDate(iBar)
            If dDate > 0 Then
                For iRec = iRec To 0 Step -1
                    dPivotDate = tPivotData.Num(0, iRec)
                    If dPivotDate = dDate Then
                        ' if session date matches, get value for the specified zone#
                        dVal = tPivotData.Num(iZone, iRec)
                        If dVal <> kNullData Then
                            dVal = dVal + dAdjust
                        End If
                        Exit For
                    ElseIf dPivotDate < dDate Then
                        Exit For
                    End If
                Next
            End If
            aResults.Num(iBar) = dVal
        Next
    End If
    
    Exit Function
    
RunError:
    Engine_PowerZonesData = -999 '(unexpected error)
    Exit Function
End Function

'http://www.tothetick.com/fetchcsv.php?mkt=sandp500&user=aamar@tickstrike.com
'Store pivot data into gdTable (rows sorted by SessionDate, fields are prices in increasing sequence):
' SessionDate, From1, To1, From2, To2, ..., From12, To12
Public Function GetPivotFarmData(ByVal strSymbol$, ByVal strUser$) As cGdTable

    Dim d#, dPrev#, iRec&, iZone&, iFirstGoodRec&, iLine&
    Dim s$, strWeb$, strZoneFile$
    Dim aLines As New cGdArray
    Dim aFields As New cGdArray
    Dim aData As New cGdTable
    
    If UCase(strUser) = "DEBUG" And IsIDE Then
        strUser = "aamar@tickstrike.com" ' a valid user ID for our testing purposes
    End If
    
    ' determine the "zone file"
    If IsForex(strSymbol) Then
        strZoneFile = LCase(StripStr(strSymbol, "$-"))
    ElseIf SecurityType(strSymbol) = "F" Then
        strZoneFile = PrimaryFutureBase(strSymbol)
        Select Case Parse(strZoneFile, "-", 1)
        Case "SP", "ES"
            strZoneFile = "sandp500"
        Case "DJ", "YM"
            strZoneFile = "dowjones"
        Case "ND", "NQ"
            strZoneFile = "nasdaq"
        Case "TF"
            strZoneFile = "russell"
        Case "GC", "XK", "ZG", "QO"
            strZoneFile = "gold"
        Case "CL", "QM"
            strZoneFile = "crudeoil"
        Case Else
            strZoneFile = ""
        End Select
    End If
    
    ' append random number as arg just to override any browser web-page-caching
    strWeb = "http://www.tothetick.com/fetchcsv.php?mkt=" & strZoneFile & "&user=" & strUser _
        & "&rand=" & Str(RandomNum(1, 9999))
    s = GetWebPageData(strWeb, 5)
    'FileFromString "c:\temp\Test1.txt", s
    s = Replace(s, vbCrLf, vbTab)
    s = Replace(s, vbLf, vbTab)
    s = Replace(s, vbCr, vbTab)
    aLines.SplitFields s, vbTab
    ' remove blank lines
    For iLine = aLines.Size - 1 To 0 Step -1
        If Len(Trim(aLines(iLine))) = 0 Then
            aLines.Remove iLine
        End If
    Next
    
    aData.Clear
    If aLines.Size >= 10 Then
        ' first line is dates: 15/04/2011,18/04/2011,19/04/2011,...,06/06/2013,07/06/2013,10/06/2013
        ' store dates in first column of table
        aFields.SplitFields aLines(0), ","
        aData.CreateField eGDARRAY_Longs, 0
        aData.NumRecords = aFields.Size
        iFirstGoodRec = 0
        dPrev = 0
        For iRec = 0 To aFields.Size - 1
            s = Trim(aFields(iRec))
            If Len(s) = 10 And Mid(s, 3, 1) = "/" And Mid(s, 6, 1) = "/" Then
                ' convert from DD/MM/YYYY to YYYYMMDD, then to Julian
                s = Right(s, 4) & Mid(s, 4, 2) & Left(s, 2)
                d = DateOf(Val(s))
                If d < 25000 Then
                    iFirstGoodRec = iRec + 1 ' bad date
                Else
                    aData.Num(0, iRec) = d
                    If d <= dPrev Then
                        iFirstGoodRec = iRec  ' date of last record must have been bad
                    End If
                    dPrev = d
                End If
            End If
        Next
        For iRec = 0 To iFirstGoodRec - 1
            aData.Num(0, iRec) = kNullData
        Next
    
        ' now parse out the values for each zone (table fields) at each date (table record)
        iZone = 0
        For iLine = aLines.Size - 1 To 1 Step -1
            s = aLines(iLine)
            If Len(s) > 10 And InStr(s, ",") > 1 Then
                ' add 2 columns for each zone
                iZone = iZone + 1
                aFields.SplitFields s, ","
                aData.CreateField eGDARRAY_Doubles, , "From" & Str(iZone)
                aData.CreateField eGDARRAY_Doubles, , "To" & Str(iZone)
                For iRec = iFirstGoodRec To aFields.Size - 1
                    s = aFields(iRec)
                    If InStr(s, "-") > 1 Then
                        d = Val(Parse(s, "-", 1))
                        aData.Num(aData.NumFields - 2, iRec) = d
                        d = Val(Parse(s, "-", 2))
                        aData.Num(aData.NumFields - 1, iRec) = d
                    End If
                Next
            End If
        Next
    End If
    
    If 0 Then
        s = aData.ToString
        FileFromString "c:\temp\Test2.txt", s
    End If
    Set GetPivotFarmData = aData

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPyramidInfo
'' Description: Fix up bogus pyramid information for rules and system rules
'' Inputs:      Reload the rules table?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FixPyramidInfo(Optional ByVal bReloadRules As Boolean = True)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bFixed As Boolean               ' Have we fixed anything?
    
    bFixed = False
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                "WHERE ([ExitBasedOnEachTrade]=-1) OR (([NumberContracts]<" & Str(kSN_MIN_ASPERCENT) & ") AND ([AsPercentOfPosition]=-1)) OR ([NumberContracts]=0) " & _
                "ORDER BY [Name];", dbOpenDynaset)
    Do While Not rs.EOF
        StatusMsg "Fixing Rule: " & rs!Name
        
        rs.Edit
        
        If rs!ExitBasedOnEachTrade = True Then
            rs!ExitBasedOnEachTrade = False
        End If
        If (rs!NumberContracts < kSN_MIN_ASPERCENT) And (rs!AsPercentOfPosition = True) Then
            rs!AsPercentOfPosition = False
        End If
        If rs!NumberContracts = 0 Then
            rs!NumberContracts = 1
        End If
        
        rs!CheckSum = BuildCheckSum(rs, "tblRules")
        rs.Update
        bFixed = True
        
        rs.MoveNext
    Loop
    
    Set rs = g.dbNav.OpenRecordset("SELECT tblSystemRules.*,tblRules.Name " & _
                "FROM [tblSystemRules] INNER JOIN [tblRules] ON tblSystemRules.RuleID=tblRules.RuleID " & _
                "WHERE (tblSystemRules.ExitBasedOnEachTrade=-1) OR ((tblSystemRules.NumberContracts<" & Str(kSN_MIN_ASPERCENT) & ") AND (tblSystemRules.AsPercentOfPosition=-1)) OR (tblSystemRules.NumberContracts=0) " & _
                "ORDER BY tblRules.Name;", dbOpenDynaset)
    Do While Not rs.EOF
        StatusMsg "Fixing Strategy Rule: " & rs!Name
        
        rs.Edit
        
        If rs!ExitBasedOnEachTrade = True Then
            rs!ExitBasedOnEachTrade = False
        End If
        If (rs!NumberContracts < kSN_MIN_ASPERCENT) And (rs!AsPercentOfPosition = True) Then
            rs!AsPercentOfPosition = False
        End If
        If rs!NumberContracts = 0 Then
            rs!NumberContracts = 1
        End If
        
        rs!CheckSum = BuildCheckSum(rs, "tblSystemRules")
        rs.Update
        bFixed = True
        
        rs.MoveNext
    Loop
    
    If (bFixed = True) And (bReloadRules = True) Then
        StatusMsg "Reloading Rules..."
        KillRulesFile
        LoadRulesTable
    End If
    
    StatusMsg "Done"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.FixPyramidInfo"
    
End Sub

' To force a rebuild of the rules table
Public Sub KillRulesFile()
    KillFile App.Path & "\Rules.tbl"
    KillFile App.Path & "\RulesEnc.tbl"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowFunctionMgr
'' Description: Show the appropriate Function Manager for the given information
'' Inputs:      Function ID, Function Name, Whether Found or not
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowFunctionMgr(ByVal lFunctionID As Long, ByVal strFunctionName As String, bFound As Boolean)
On Error GoTo ErrSection:
    
    If bFound Then
        If g.Functions.Item(CStr(lFunctionID)).ImplementationTypeID = kSN_Custom Then
            frmFunctionMgrCT.ShowMe lFunctionID
        Else
            frmFunctionMgr.ShowMe lFunctionID
        End If
    Else
        frmFunctionMgrCT.ShowMe 0, , strFunctionName
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mSysNav.ShowFunctionMgr"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPeriodInMarkets
'' Description: Fix the Period in "Of" expressions surrounded by quotes
'' Inputs:      Expression
'' Returns:     Fixed expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FixPeriodInMarkets(ByVal strExpression As String) As String
On Error GoTo ErrSection:

    Dim astrTokens As New cGdArray      ' Array of space delimited tokens
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol of the market variable
    Dim strPeriod As String             ' Period of the market variable
    Dim Bars As New cGdBars             ' Temporary Bars structure
    
    If Len(strExpression) > 0 Then
        astrTokens.SplitFields strExpression, " "
        For lIndex = 0 To astrTokens.Size - 1
            If UCase(astrTokens(lIndex)) = "OF" Then
                If lIndex + 1 < astrTokens.Size Then
                    If Left(astrTokens(lIndex + 1), 1) = Chr(34) And Right(astrTokens(lIndex + 1), 1) = Chr(34) Then
                        strSymbol = Parse(Replace(astrTokens(lIndex + 1), Chr(34), ""), ",", 1)
                        strPeriod = Parse(Replace(astrTokens(lIndex + 1), Chr(34), ""), ",", 2)
                        
                        If Len(strPeriod) > 0 Then
                            Bars.Prop(eBARS_PeriodicityStr) = strPeriod
                            strPeriod = Bars.Prop(eBARS_PeriodicityStr)
                            
                            astrTokens(lIndex + 1) = Chr(34) & strSymbol & "," & strPeriod & Chr(34)
                        Else
                            astrTokens(lIndex + 1) = Chr(34) & strSymbol & Chr(34)
                        End If
                    End If
                End If
            End If
        Next lIndex
    End If
    
    FixPeriodInMarkets = astrTokens.JoinFields(" ")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.FixPeriodInMarkets"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsValidMarket
'' Description: Validate any "Symbol,Period" markets
'' Inputs:      Market
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsValidMarket(ByVal strMarket As String) As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    Dim strSymbol As String             ' Symbol of the given market
    Dim strPeriod As String             ' Period of the given market
    Dim Bars As New cGdBars             ' Temporary Bars structure
    
    bReturn = False
    If Left(strMarket, 1) = Chr(34) And Right(strMarket, 1) = Chr(34) Then
        strSymbol = Parse(Replace(strMarket, Chr(34), ""), ",", 1)
        strPeriod = Parse(Replace(strMarket, Chr(34), ""), ",", 2)
        
        If Len(strSymbol) > 0 And Len(strPeriod) > 0 Then
            Bars.Prop(eBARS_PeriodicityStr) = strPeriod
            If Bars.Prop(eBARS_Periodicity) < ePRD_Days Then
                DM_GetBars Bars, strSymbol, strPeriod, LastDailyDownload - 5
            Else
                DM_GetBars Bars, strSymbol, strPeriod
            End If
            bReturn = (Bars.Size > 0)
        ElseIf Len(strSymbol) > 0 Then
            DM_GetBars Bars, strSymbol, "Daily"
            bReturn = (Bars.Size > 0)
        ElseIf Len(strPeriod) > 0 Then
            bReturn = True
        End If
    Else
        bReturn = True
    End If
    
    IsValidMarket = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IsValidMarket"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsGenesisLibrary
'' Description: Determine if the library ID given is a Genesis library
'' Inputs:      Library ID
'' Returns:     True if Genesis Library, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsGenesisLibrary(ByVal lLibraryID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    For lIndex = 0 To g.tblLibrary.NumRecords - 1
        If g.tblLibrary.Num(LibraryField(etblLib_ID), lIndex) = lLibraryID Then
            bReturn = (UCase(g.tblLibrary.Item(LibraryField(etblLib_Name), lIndex)) = "GENESIS SYSTEM FUNCTIONS")
            Exit For
        End If
    Next lIndex
    
    IsGenesisLibrary = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IsGenesisLibrary"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnglishFromCoded
'' Description: Build an English string from the Coded Text string
'' Inputs:      Coded Text
'' Returns:     English
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnglishFromCoded(ByVal strCodedText As String) As String
On Error GoTo ErrSection:

    Dim astrTokens As New cGdArray      ' Array of tokens from the coded text string
    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return value for the function
    
    astrTokens.SplitFields strCodedText, "~"
    For lIndex = 0 To astrTokens.Size - 1
        astrTokens(lIndex) = Mid(Trim(astrTokens(lIndex)), 6)
    Next lIndex
    
    strReturn = astrTokens.JoinFields(" ")
    strReturn = Replace(strReturn, " ( ", "(")
    strReturn = Replace(strReturn, " )", ")")
    
    EnglishFromCoded = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.EnglishFromCoded"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextBarFunctionIds
'' Description: Build a string of function IDs for the Next Bar functions
'' Inputs:      Include Next Bar Open?
'' Returns:     String of ID's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextBarFunctionIds(Optional ByVal bIncludeNextBarOpen As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    strReturn = ""
    For lIndex = 1 To g.Functions.Count
        If UCase(Left(g.Functions.Item(lIndex).FunctionName, 9)) = "NEXT BAR " Then
            If IsGenesisLibrary(g.Functions.Item(lIndex).LibraryID) Then
                If (bIncludeNextBarOpen = True) Or (UCase(g.Functions.Item(lIndex).FunctionName) <> "NEXT BAR OPEN") Then
                    strReturn = strReturn & "," & g.Functions.Item(lIndex).FunctionID & ","
                End If
            End If
        End If
    Next lIndex
    
    NextBarFunctionIds = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.NextBarFunctionIds"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UsesNextBarFunctions
'' Description: Does the given function ID reference any next bar functions?
'' Inputs:      Function ID
'' Returns:     True if reference Next Bar functions, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UsesNextBarFunctions(ByVal lFunctionID As Long, Optional ByVal bIncludeNextBarOpen As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim alRefs As cGdArray              ' Function references for the given ID
    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim strNextBarFuncs As String       ' Next bar function ID's
    
    bReturn = False
    strNextBarFuncs = NextBarFunctionIds(bIncludeNextBarOpen)
    Set alRefs = FunctionRefs(lFunctionID)
    
    For lIndex = 0 To alRefs.Size - 1
        If InStr(strNextBarFuncs, "," & Str(alRefs(lIndex)) & ",") <> 0 Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    UsesNextBarFunctions = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.UsesNextBarFunctions"
    
End Function

Public Function ExecuteExpressionSet(astrParms As cGdArray, astrBarNames As cGdArray, aArrayOfBarHandles As cGdArray, _
    astrExpressions As cGdArray, aArrayOfResultHandles As cGdArray, Optional ByVal hMinBarsReq& = 0, Optional ByVal hStrTimes& = 0) As Long
On Error GoTo ErrSection:

    Dim rc&
    Dim aClearParms As New cGdArray
    Static iCount&
    
    ' string array should just have 1 item: a unique ID (using the static counter)
    iCount = iCount + 1
    If iCount > 1000000000 Then iCount = 1
    astrParms.Size = 1
    astrParms(0) = "ExpSet-" & Str(iCount)
    aClearParms.Create eGDARRAY_Strings, 1
    aClearParms(0) = astrParms(0)
    
    ' run the expressions
    rc = StartLastBarSet(astrParms.ArrayHandle, astrBarNames.ArrayHandle, aArrayOfBarHandles.ArrayHandle, _
                astrExpressions.ArrayHandle, aArrayOfResultHandles.ArrayHandle, hMinBarsReq, hStrTimes)
    
    ' and clear the expression set
    ClearLastBarSet aClearParms.ArrayHandle, ByVal 0&
    
ErrExit:
    Set aClearParms = Nothing
    ExecuteExpressionSet = rc
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.ExecuteExpressionSet"
End Function

' All the $ amounts in the trades array will be multiplied by the specified #
' NOTE: for Futures, the multiplier essentially = # of contracts
Public Sub MultiplyTrades(astrTrades As cGdArray, Optional ByVal dMultiplier As Double = 1)
On Error GoTo ErrSection:

    Dim iTrade&, iFld&, d#
    Dim astrFlds As New cGdArray

    ' just return if nothing to do
    If dMultiplier = 1 Or astrTrades.Size <= 1 Then Exit Sub

    ' parse out the header and get the Margin
    astrFlds.SplitFields astrTrades(0), vbTab
    iFld = 11
    d = Val(astrFlds(iFld))
    If d > 0 Then
        ' TLB: just testing this as an idea (not sure if we'll ever actually use it) ...
        If dMultiplier < 0 Then
            ' in this mode, calculate # contracts that can be
            ' traded for the specified $ amount that was passed in
            dMultiplier = Int(Abs(dMultiplier) / d)
            If dMultiplier < 1 Then dMultiplier = 1
        End If
        astrFlds(iFld) = Str(d * dMultiplier)
        astrTrades(0) = astrFlds.JoinFields(vbTab)
    End If

    If dMultiplier <> 1 Then
        ' now do every trade
        For iTrade = 1 To astrTrades.Size - 1
            astrFlds.SplitFields astrTrades(iTrade), vbTab
            For iFld = 7 To 9 ' the $ profit fields
                d = Val(astrFlds(iFld))
                If d <> kNullData Then
                    astrFlds(iFld) = Format(d * dMultiplier, "#0.00")
                End If
            Next
            iFld = 18 ' and # stock shares
            d = Val(astrFlds(iFld))
            If d > 0 Then
                astrFlds(iFld) = Format(d * dMultiplier, "#0")
            End If
            astrTrades(iTrade) = astrFlds.JoinFields(vbTab)
        Next
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.MultiplyTrades"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildRule
'' Description: Builds one expression out of the given parameters
'' Inputs:      Condition, Order Price, With Limit Price, Buy, Order Type
'' Returns:     Expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildRuleFromParts(ByVal strEnglishCondition As String, ByVal strEnglishPrice As String, ByVal strEnglishWithLimit As String, ByVal bBuy As Boolean, ByVal nOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:
    
    Dim strRule As String               ' Expression to return from the function
    Dim lPos As Long                    ' Position in the string
    Dim astrRule As cGdArray            ' Array of expressions for the rule
    Dim lIndex As Long                  ' Index into a for loop
    
    ' IF Condition THEN Action...
    strRule = Trim(strEnglishCondition)
    lPos = InStr(UCase(strRule), "IF")
    If Left(UCase(strRule), 2) <> "IF" And Left(strRule, 1) <> "'" Then
        If lPos <> 0 Then
            strRule = strEnglishCondition
        ElseIf InStr(strRule, ":=") = 0 Then
            strRule = "IF " & strEnglishCondition
        Else
            Set astrRule = New cGdArray
            astrRule.SplitFields strRule, vbLf
            For lIndex = 0 To astrRule.Size - 1
                If InStr(astrRule(lIndex), ":=") = 0 Then
                    astrRule(lIndex) = "IF " & astrRule(lIndex)
                    Exit For
                End If
            Next lIndex
            
            strRule = astrRule.JoinFields(vbCrLf)
        End If
    End If
    
    If bBuy = True Then
        strRule = strRule & " THEN " & vbCrLf & vbTab & "BUY ("
    Else
        strRule = strRule & " THEN " & vbCrLf & vbTab & "SELL ("
    End If
    
    ' Order Price...
    If Len(strEnglishPrice) > 0 Then
        If (nOrderType <> eTT_OrderType_Market) And (nOrderType <> eTT_OrderType_MarketOnClose) Then
            strRule = strRule & strEnglishPrice & ", "
        End If
    End If
    
    ' Order type...
    Select Case nOrderType
        Case eTT_OrderType_Stop
            strRule = strRule & Chr(34) & "Stop" & Chr(34)
        Case eTT_OrderType_Limit
            strRule = strRule & Chr(34) & "Limit" & Chr(34)
        Case eTT_OrderType_Market
            strRule = strRule & "Close, " & Chr(34) & "Market" & Chr(34)
        Case eTT_OrderType_StopWithLimit
            strRule = strRule & Chr(34) & "Stop with Limit" & Chr(34)
    End Select
    
    ' With Limit Price
    If Len(strEnglishWithLimit) > 0 Then
        If (nOrderType <> eTT_OrderType_Market) And (nOrderType <> eTT_OrderType_MarketOnClose) Then
            strRule = strRule & ", " & strEnglishWithLimit
        End If
    End If
    
    BuildRuleFromParts = strRule & ")" & vbCrLf & "ENDIF"
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mSysNav.BuildRuleFromParts"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadStrategiesRecordset
'' Description: Load a recordset of strategies
'' Inputs:      Sort by Strategy Name?
'' Returns:     Recordset
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadStrategiesRecordset(Optional ByVal bSortByStrategyName As Boolean = False) As Recordset
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strSql As String                ' SQL statement for the recordset
    
    strSql = "SELECT tblSystems.*, tblLibrarys.* FROM tblLibrarys INNER JOIN tblSystems ON tblLibrarys.LibraryID = tblSystems.LibraryID"
    If bSortByStrategyName Then
        strSql = strSql & " ORDER BY [SystemName]"
    End If
    strSql = strSql & ";"
    
    Set rs = g.dbNav.OpenRecordset(strSql, dbOpenDynaset)
    ValidateCheckSums rs, "tblSystems"
    ValidateCheckSums rs, "tblLibrarys"
    
    Set LoadStrategiesRecordset = rs

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.LoadStrategiesRecordset"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadStrategyBasketsRecordset
'' Description: Load a recordset of strategy basket(s)
'' Inputs:      Strategy Basket ID, Database, Include Libraries?
'' Returns:     Recordset
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadStrategyBasketsRecordset(Optional ByVal lStrategyBasketID As Long = -1&, Optional DB As Database = Nothing, Optional ByVal bIncludeLibraries As Boolean = True) As Recordset
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strSql As String                ' SQL statement for the recordset
    
    If DB Is Nothing Then
        Set DB = g.dbNav
    End If
    
    If bIncludeLibraries = True Then
        strSql = "SELECT tblStrategyBaskets.*, tblLibrarys.* FROM tblLibrarys INNER JOIN tblStrategyBaskets ON tblLibrarys.LibraryID = tblStrategyBaskets.LibraryID"
    Else
        strSql = "SELECT * FROM [tblStrategyBaskets]"
    End If
    If lStrategyBasketID <> -1& Then
        strSql = strSql & " WHERE [StrategyBasketID] = " & Str(lStrategyBasketID)
    End If
    strSql = strSql & ";"
    
    Set rs = DB.OpenRecordset(strSql, dbOpenDynaset)
    ValidateCheckSums rs, "tblStrategyBaskets"
    
    If bIncludeLibraries Then
        ValidateCheckSums rs, "tblLibrarys"
    End If
    
    Set LoadStrategyBasketsRecordset = rs

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.LoadStrategyBasketsRecordset"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOwnerOfLibraryRecordset
'' Description: Determine if the current user is marked as the owner on the
''              given library recordset
'' Inputs:      Library Recordset
'' Returns:     True if owner, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsOwnerOfLibraryRecordset(ByVal rs As Recordset) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            If ItemExists(rs.Fields, "Owners") Then
                bReturn = (InStr("," & rs!Owners & ",", "," & Str(g.lLCD) & ",") <> 0)
            End If
        End If
    End If
    
    IsOwnerOfLibraryRecordset = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IsOwnerOfLibraryRecordset"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IncludeStrategiesFromRecordset
'' Description: Include the strategy in the given recordset record?
'' Inputs:      Recordset, Include Hidden System if IDE?
'' Returns:     True if include, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IncludeStrategiesFromRecordset(ByVal rs As Recordset, Optional ByVal bIncludeHiddenIfIde As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not (rs.BOF And rs.EOF) Then
        If rs!Ignore = 0 Then
            If (rs![tblSystems.CheckSum] <> 0.5) And (rs![tblLibrarys.CheckSum] <> 0.5) Then
                If ((rs![tblSystems.SecurityLevel] <> 3) And (rs![tblSystems.IsGuru] = 0)) Or ((bIncludeHiddenIfIde = True) And (IsIDE = True)) Then
                    bReturn = True
                Else
                    bReturn = IsOwnerOfLibraryRecordset(rs)
                End If
            End If
        End If
    End If
    
    IncludeStrategiesFromRecordset = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IncludeStrategiesFromRecordset"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IncludeStrategyBasketsFromRecordset
'' Description: Include the strategy baskets in the given recordset record?
'' Inputs:      Recordset, Include Hidden Strategy Basket if IDE?, Include
''              even if not the owner?
'' Returns:     True if include, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IncludeStrategyBasketsFromRecordset(ByVal rs As Recordset, Optional ByVal bIncludeHiddenIfIde As Boolean = False, Optional ByVal bIncludeGuruIfNotOwner As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not (rs.BOF And rs.EOF) Then
        If rs!Ignore = 0 Then
            If (rs![tblStrategyBaskets.CheckSum] <> 0.5) And (rs![tblLibrarys.CheckSum] <> 0.5) Then
                If (bIncludeHiddenIfIde = True) And (IsIDE = True) Then
                    bReturn = True
                ElseIf rs![tblStrategyBaskets.IsGuru] <> 0 Then
                    If bIncludeGuruIfNotOwner Then
                        bReturn = True
                    Else
                        bReturn = IsOwnerOfLibraryRecordset(rs)
                    End If
                Else
                    bReturn = (rs![tblStrategyBaskets.SecurityLevel] <> 3)
                End If
            End If
        End If
    End If
    
    IncludeStrategyBasketsFromRecordset = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IncludeStrategyBasketsFromRecordset"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOwnerOfGuruObject
'' Description: Determine if the current user is marked as the owner of guru
''              objects in the given library
'' Inputs:      Library ID
'' Returns:     True if owner, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsOwnerOfGuruObject(ByVal lLibraryID As Long) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] WHERE [LibraryID]=" & Str(lLibraryID) & ";", dbOpenDynaset)
    IsOwnerOfGuruObject = IsOwnerOfLibraryRecordset(rs)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IsOwnerOfGuruObject"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOwnerOfGuruBasket
'' Description: Determine if the current user is marked as the owner of the
''              given guru strategy basket
'' Inputs:      Strategy Basket ID
'' Returns:     True if owner, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsOwnerOfGuruBasket(ByVal lStrategyBasketID As Long) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * " & _
                "FROM [tblStrategyBaskets] INNER JOIN [tblLibrarys] ON tblStrategyBaskets.LibraryID=tblLibrarys.LibraryID " & _
                "WHERE [StrategyBasketID]=" & Str(lStrategyBasketID) & ";", dbOpenDynaset)
    IsOwnerOfGuruBasket = IsOwnerOfLibraryRecordset(rs)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IsOwnerOfGuruBasket"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOwnerOfGuruStrategy
'' Description: Determine if the current user is marked as the owner of the
''              given guru strategy
'' Inputs:      Strategy ID
'' Returns:     True if owner, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsOwnerOfGuruStrategy(ByVal lStrategyID As Long) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * " & _
                "FROM [tblSystems] INNER JOIN [tblLibrarys] ON tblSystems.LibraryID=tblLibrarys.LibraryID " & _
                "WHERE [SystemNumber]=" & Str(lStrategyID) & ";", dbOpenDynaset)
    IsOwnerOfGuruStrategy = IsOwnerOfLibraryRecordset(rs)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.IsOwnerOfGuruStrategy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MaxContractsForEnablement
'' Description: Maximum number of contracts for the given enablement
'' Inputs:      Enablement, Is Guru?
'' Returns:     Max Contracts ( 999999 if no limit )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MaxContractsForEnablement(ByVal strEnablement As String, ByVal bIsGuru As Boolean) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strEnablements As String        ' Enablement string
    Dim astrEnablements As cGdArray     ' Array of enablements
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrEnablement As cGdArray      ' Enablement code broken out into fields
    
    lReturn = 0
    If Len(strEnablement) = 0 Then
        lReturn = Abs(kNullData)
    ElseIf (bIsGuru = False) And (HasModule(strEnablement) = True) Then
        lReturn = Abs(kNullData)
    ElseIf (bIsGuru = True) Then
        strEnablements = g.strAuthorizationString
        If Len(strEnablements) > 0 Then
            Set astrEnablements = New cGdArray
            astrEnablements.SplitFields strEnablements, ","
            
            strEnablement = strEnablement & "_"
            For lIndex = 0 To astrEnablements.Size - 1
                If Left(astrEnablements(lIndex), Len(strEnablement)) = strEnablement Then
                    Set astrEnablement = New cGdArray
                    astrEnablement.SplitFields astrEnablements(lIndex), "_"
                    
                    If astrEnablement.Size = 3 Then
                        lReturn = CLng(Val(astrEnablement(2)))
                    End If
                    
                    Exit For
                End If
            Next lIndex
        End If
    End If
    
    DebugLog "MaxContractsForEnablement(" & strEnablement & ", " & Str(bIsGuru) & ") = " & Str(lReturn)
    MaxContractsForEnablement = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.MaxContractsForEnablement"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MaxUnitsForAutoTrade
'' Description: Maximum number of units allowed for an automated trading item
'' Inputs:      Automated Trading Item
'' Returns:     Max Units Allowed ( 999999 if no limit )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MaxUnitsForAutoTrade(ByVal TradeItem As cAutoTradeItem) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = Abs(kNullData)
    If TradeItem.StrategyBasketID > 0& Then
        lReturn = MaxUnitsForStrategyBasketID(TradeItem.StrategyBasketID)
    End If
    
    MaxUnitsForAutoTrade = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.MaxUnitsForAutoTrade"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MaxUnitsForAutoTradeID
'' Description: Maximum number of units allowed for an automated trading item ID
'' Inputs:      Automated Trading Item ID
'' Returns:     Max Units Allowed ( 999999 if no limit )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MaxUnitsForAutoTradeID(ByVal lAutoTradeItemID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    lReturn = Abs(kNullData)
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM " & _
                "[tblAutoTradingItem] INNER JOIN [tblStrategyBaskets] ON tblAutoTradingItem.StrategyBasketID=tblStrategyBaskets.StrategyBasketID " & _
                "WHERE [TradingItemID]=" & Str(lAutoTradeItemID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        lReturn = MaxContractsForEnablement(rs!RequiredMod, rs!IsGuru)
    End If
    
    MaxUnitsForAutoTradeID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.MaxUnitsForAutoTradeID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MaxUnitsForStrategyBasketID
'' Description: Maximum number of units allowed for a strategy basket ID
'' Inputs:      Strategy Basket ID
'' Returns:     Max Units Allowed ( 999999 if no limit )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MaxUnitsForStrategyBasketID(ByVal lStrategyBasketID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    lReturn = Abs(kNullData)
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] WHERE [StrategyBasketID]=" & Str(lStrategyBasketID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        lReturn = MaxContractsForEnablement(rs!RequiredMod, rs!IsGuru)
    End If
    
    MaxUnitsForStrategyBasketID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.MaxUnitsForStrategyBasketID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateGuruAutoTradeItems
'' Description: Create automated trading items for guru baskets if they don't
''              currently exist, but the user is enabled for them
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateGuruAutoTradeItems()
On Error GoTo ErrSection:

    Dim Baskets As cStrategyBaskets     ' Collection of strategy baskets
    Dim lIndex As Long                  ' Index into a for loop
    Dim Basket As cStrategyBasket       ' Strategy basket object
    Dim lMaxUnits As Long               ' Max units allowed for the basket
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lParentID As Long               ' Parent ID
    Dim basketItem As cStrategyBasketItem ' Strategy basket item
    Dim Bars As cGdBars                 ' Bars object
    Dim lCount As Long                  ' Counter variable
    Dim lAccountID As Long              ' First SimStream account ID
    Dim ParentTradeItem As cAutoTradeItem ' Parent auto trade item
    
    Set Baskets = New cStrategyBaskets
    
    ' Load all of the baskets without loading the items...
    Baskets.LoadDb True, True, False
    
    For lIndex = 1 To Baskets.Count
        Set Basket = Baskets(lIndex)
        
        If (Basket.IsGuru = True) Then
            ' Reload the basket here because we did not load all the items on the big load above...
            Basket.LoadDb Basket.ID, True, True
            
            lMaxUnits = MaxContractsForEnablement(Basket.RequiredModule, Basket.IsGuru)
            If lMaxUnits > 0 Then
                If g.TradingItems.IsStrategyBasketInAutoTradeItem(Basket.ID, ParentTradeItem) = False Then
                    lAccountID = mTradeTracker.FirstSimStreamAccountID
                    
                    Set TradeItem = New cAutoTradeItem
                    With TradeItem
                        If Len(Basket.Name) > 50 Then
                            .Name = Left(Basket.Name, 50)
                        Else
                            .Name = Basket.Name
                        End If
                        .AccountID = lAccountID
                        .Deleted = False
                        
                        .QtyNextEntry = lMaxUnits
                        .ConfirmOrders = False
                        .ConfirmTimeout = 10&
                        
                        .ParentID = -1&
                        .StrategyBasketID = Basket.ID
                        .StrategyBasketItemID = 0&
                        .StrategyBasketLastModified = Basket.LastModified
                        .Overrides = ""
                        .StrategyBasketItemKey = ""
                        .Save
                        
                        lParentID = .AutoTradeItemID
                    End With
                    
                    g.TradingItems.Add TradeItem
                    
                    Set ParentTradeItem = TradeItem
                End If
                
                If g.TradingItems.NumStrategyBasketItemsInAutoTradeItem(Basket.ID) = 0 Then
                    lCount = 1
                    For lIndex2 = 1 To Basket.Items.Count
                        Set basketItem = Basket.Items(lIndex2)
                        
                        If Len(basketItem.Symbol) > 0 Then
                            Set Bars = New cGdBars
                            Set TradeItem = New cAutoTradeItem
                            With TradeItem
                                If Len(Basket.Name) > 45 Then
                                    .Name = Left(Basket.Name, 45) & " #" & Format(lCount, "000")
                                Else
                                    .Name = Basket.Name & " #" & Format(lCount, "000")
                                End If
                                lCount = lCount + 1&
                                
                                .AccountID = ParentTradeItem.AccountID ' lAccountID
                                .SymbolOrSymbolID = basketItem.SymbolOrSymbolID
                                .BarPeriod = basketItem.Period
                                .StrategyID = basketItem.StrategyID
                                .StrategyName = basketItem.StrategyName
                                .ConfirmOrders = False
                                .ConfirmTimeout = 10&
                                .Deleted = False
                                
                                .ParentID = ParentTradeItem.AutoTradeItemID ' lParentID
                                .QtyNextEntry = lMaxUnits * basketItem.ContractMultiplier
                                .StrategyBasketID = Basket.ID
                                .StrategyBasketLastModified = Basket.LastModified
                                .StrategyBasketItemID = basketItem.ID
                                .StrategyBasketItemMult = basketItem.ContractMultiplier
                                .Overrides = basketItem.Overrides
                                .StrategyBasketItemKey = basketItem.Key
                                
                                SetBarProperties Bars, basketItem.Symbol
                                .MinutesBefore = 3&
                                .OnCloseTimeExch = (Bars.Prop(eBARS_EndTime) - 3#) / 1440#
                                
                                .Save
                            End With
                            g.TradingItems.Add TradeItem
                        End If
                    Next lIndex2
                End If
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.CreateGuruAutoTradeItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportStrategyBaskets
'' Description: Import the existing strategy baskets from files into the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportStrategyBaskets()
On Error GoTo ErrSection:

    Dim astrFiles As cGdArray           ' List of files to import
    Dim lIndex As Long                  ' Index into a for loop
    Dim Basket As cStrategyBasket       ' Strategy basket object
    
    Set astrFiles = New cGdArray
    astrFiles.GetMatchingFiles AddSlash(App.Path) & "Custom\*.SB"
    For lIndex = 0 To astrFiles.Size - 1
        Set Basket = New cStrategyBasket
        If Basket.LoadFile(astrFiles(lIndex)) Then
            Basket.SaveDb
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.ImportStrategyBaskets"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoTradeItemIsDeleted
'' Description: Determine if the given automated trading item is deleted
'' Inputs:      Auto Trade Item ID
'' Returns:     True if Deleted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AutoTradeItemIsDeleted(ByVal lAutoTradeItemID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    bReturn = False
    If g.TradingItems.Exists(Str(lAutoTradeItemID)) Then
        bReturn = g.TradingItems(Str(lAutoTradeItemID)).Deleted
    Else
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] WHERE [TradingItemID]=" & Str(lAutoTradeItemID) & ";", dbOpenDynaset)
        If (rs.BOF And rs.EOF) Then
            bReturn = True
        Else
            bReturn = rs!Deleted
        End If
    End If
    
    AutoTradeItemIsDeleted = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.AutoTradeItemIsDeleted"
    
End Function

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixDuplicateBasketNames
'' Description: Get rid of duplicate strategy basket names out of the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FixDuplicateBasketNames()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim Baskets As cGdTree              ' Collection of baskets
    Dim strInfo As String               ' Basket information
    Dim lStrategyBasketID As Long       ' Strategy Basket ID
    Dim lLibraryID As Long              ' Library ID
    Dim strNewName As String            ' New name for the basket
    
    Set Baskets = New cGdTree

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets];", dbOpenDynaset)
    Do While Not rs.EOF
        If Baskets.Exists(rs!Name) Then
            strInfo = Baskets(rs!Name)
            lStrategyBasketID = CLng(Val(Parse(strInfo, vbTab, 1)))
            lLibraryID = CLng(Val(Parse(strInfo, vbTab, 2)))
            
            If rs!LibraryID = lLibraryID Then
                Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] WHERE [StrategyBasketID]=" & Str(rs!StrategyBasketID) & ";", dbOpenDynaset)
                Do While Not rs2.EOF
                    rs2.Edit
                    rs2!StrategyBasketID = lStrategyBasketID
                    rs2.Update
                    
                    rs2.MoveNext
                Loop
                
                rs.Delete
            Else
                strNewName = "_" & rs!Name
                Do While Baskets.Exists(strNewName)
                    strNewName = "_" & strNewName
                Loop
                
                If rs!LibraryID = kSN_UserLibrary Then
                    rs.Edit
                    rs!Name = strNewName
                    rs.Update
                    
                    strInfo = Str(rs!StrategyBasketID) & vbTab & Str(rs!LibraryID)
                    Baskets.Add strInfo, rs!Name
                    
                ElseIf lLibraryID = kSN_UserLibrary Then
                    Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] WHERE [StrategyBasketID]=" & Str(lStrategyBasketID) & ";", dbOpenDynaset)
                    If Not (rs2.BOF And rs2.EOF) Then
                        rs2.Edit
                        rs2!Name = strNewName
                        rs2.Update
                        
                        strInfo = Str(rs2!StrategyBasketID) & vbTab & Str(rs2!LibraryID)
                        Baskets.Add strInfo, rs2!Name
                    End If
                    
                    strInfo = Str(rs!StrategyBasketID) & vbTab & Str(rs!LibraryID)
                    Baskets(rs!Name) = strInfo
                Else
                End If
            End If
        Else
            strInfo = Str(rs!StrategyBasketID) & vbTab & Str(rs!LibraryID)
            Baskets.Add strInfo, rs!Name
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.FixDuplicateBasketNames"
    
End Sub
#Else
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixDuplicateBasketNames
'' Description: Get rid of duplicate strategy basket names out of the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FixDuplicateBasketNames()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim Baskets As cGdTree              ' Collection of baskets
    Dim strInfo As String               ' Basket information
    Dim lStrategyBasketID As Long       ' Strategy Basket ID
    Dim lLibraryID As Long              ' Library ID
    Dim strNewName As String            ' New name for the basket
    
    Set Baskets = New cGdTree

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets];", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        Do While Not rs.BOF
            If Baskets.Exists(rs!Name) Then
                strInfo = Baskets(rs!Name)
                lStrategyBasketID = CLng(Val(Parse(strInfo, vbTab, 1)))
                lLibraryID = CLng(Val(Parse(strInfo, vbTab, 2)))
                
                If (rs!LibraryID = lLibraryID) And (lLibraryID <> kSN_UserLibrary) Then
                    Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] WHERE [StrategyBasketID]=" & Str(rs!StrategyBasketID) & ";", dbOpenDynaset)
                    Do While Not rs2.EOF
                        rs2.Edit
                        rs2!StrategyBasketID = lStrategyBasketID
                        rs2.Update
                        
                        rs2.MoveNext
                    Loop
                    
                    rs.Delete
                End If
            Else
                strInfo = Str(rs!StrategyBasketID) & vbTab & Str(rs!LibraryID)
                Baskets.Add strInfo, rs!Name
            End If
            
            rs.MovePrevious
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mSysNav.FixDuplicateBasketNames"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StrategyBasketItemIdForKey
'' Description: Get the Strategy Basket Item for the given key
'' Inputs:      Strategy Basket ID, Strategy Basket Item Key
'' Returns:     Strategy Basket Item ID ( 0 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StrategyBasketItemIdForKey(ByVal lStrategyBasketID As Long, ByVal strStrategyBasketItemKey As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim astrKey As cGdArray             ' Key broken out into fields
    Dim astrFields As cGdArray          ' Expressions for the query
    Dim rs As Recordset                 ' Recordset into the database
    
    lReturn = 0&
    
    If Len(strStrategyBasketItemKey) > 0 Then
        Set astrKey = New cGdArray
        astrKey.SplitFields strStrategyBasketItemKey, "|"
        
        If astrKey.Size = 4 Then
            Set astrFields = New cGdArray
            astrFields.Create eGDARRAY_Strings, 5
            astrFields(0) = "[StrategyBasketID]=" & Str(lStrategyBasketID)
            astrFields(1) = "[SystemNumber]=" & astrKey(0)
            astrFields(2) = "[SymbolGroupID]='" & astrKey(1) & "'"
            If IsNumeric(astrKey(2)) Then
                astrFields(3) = "[SymbolID]=" & astrKey(2)
            Else
                astrFields(3) = "[Symbol]='" & astrKey(2) & "'"
            End If
            astrFields(4) = "[BarPeriod]='" & astrKey(3) & "'"
            
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBasketItems] WHERE " & astrFields.JoinFields(" AND ") & ";", dbOpenDynaset)
            If Not (rs.BOF And rs.EOF) Then
                lReturn = rs!StrategyBasketItemID
            End If
        End If
    End If
    
    StrategyBasketItemIdForKey = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mSysNav.StrategyBasketItemIdForKey"
    
End Function
