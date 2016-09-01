Attribute VB_Name = "mComEval"
Option Explicit

Global Const gDS_Trades = 1
Global Const gDS_SystemInfo = 2

Global Const gTokenLen = 6
Global Const gEntrySignal = 0
Global Const gExitSignal = 1

'All possible Phrase Types that can appear in Rules
Global Const gPT_Nbr = 1
Global Const gPT_Add = 2
Global Const gPT_Sub = 3
Global Const gPT_Mult = 4
Global Const gPT_Div = 5
Global Const gPT_LeftPar = 6
Global Const gPT_RightPar = 7
Global Const gPT_GT = 8
Global Const gPT_GE = 9
Global Const gPT_LT = 10
Global Const gPT_LE = 11
Global Const gPT_NE = 12
Global Const gPT_EQ = 13
Global Const gPT_And = 14
Global Const gPT_Or = 15
Global Const gPT_Not = 16
Global Const gPT_Parm = 17
Global Const gPT_Text = 22
Global Const gPT_FLParen = 23
Global Const gPT_FRParen = 24
Global Const gPT_Offset = 25
Global Const gPT_Comma = 26
Global Const gPT_FInternal = 27
Global Const gPT_FCompiled = 28
Global Const gPT_FCompiledAction = 29
Global Const gPT_FTradeSense = 30
Global Const gPT_If = 31
Global Const gPT_DoubleQuote = 32
Global Const gPT_Error = 33
Global Const gPT_OuterParens = 34
Global Const gPT_Then = 35
Global Const gPT_Comment = 36
Global Const gPT_Enter = 37
Global Const gPT_Of = 38
Global Const gPT_Else = 41
Global Const gPT_ElseIf = 42
Global Const gPT_EndIf = 43
Global Const gPT_Tab = 44
Global Const gPT_EnterFormatting = 45
Global Const gPT_DoUntil = 48
Global Const gPT_EndDo = 49
Global Const gPT_BracketComment = 50

'Tokens that point to internal arrays
Global Const gPT_Trades = 39
Global Const gPT_Bars = 40
Global Const gPT_Portfolio = 46
Global Const gPT_Systems = 47

'Constants used to access correct gdArray in Trades class  (cTrades)
'Each number represents the gdArray.  ei entr_Position is array 1, etc.
Public Enum en_Trades
    entr_TradeNbr = 1
    entr_Position = 2
    entr_SignalType = 3
    entr_TradeDate = 4
    entr_Price = 5
    entr_Signal = 6
    entr_EntryRuleID = 7
    entr_ExitRuleID = 8
    entr_Units = 9
    entr_Profit = 10
    entr_TotalProfit = 11
    entr_AccountBalance = 12
    entr_TradeSymbol = 13
    entr_MaxProfit = 14
    entr_MaxLoss = 15
    entr_BarsInTrade = 16
    entr_Skip = 17
    entr_SkipRpt = 18
    entr_TradeDayOfWeek = 19
    entr_TradeDayOfMonth = 20
    entr_TradeDayOfYear = 21
    entr_SysNbr = 22
    entr_Allocation = 23
    entr_Rank = 24
    entr_EquityAvail = 25
    entr_Conflict = 26
    entr_OpenTrade = 27
    entr_Link = 28
    entr_EntryExitPtr = 29
    entr_SortKey = 30
    entr_TestNumeric = 31
    entr_TestString = 32
    entr_OpenTradesTotal = 33
    entr_Msg = 34
    entr_SignalsTot = 35
End Enum
Global Const gTradeArrays = 35

'Constants used to access correct gdArray in Portfolio class  (cPortfolio)
'Each number represents the gdArray.  ei enM_BeginBalance is array 1, etc.
Public Enum enm_Portfolio
    enm_BeginBalance = 0
    enm_WinPct = 1
    enm_PLRatio = 2
    enm_Drawdown = 3
    enm_LargestLoss = 4
    enm_WinAvg = 5
    enm_CL = 6
End Enum

Public Enum ens_Systems
    ens_SystemNumber = 0
    ens_SystemName = 1
    ens_SystemClass = 2
    ens_BarTimeFrame = 3
    ens_Margin = 4
    ens_WinPct = 5
    ens_PLRatio = 6
    ens_Drawdown = 7
    ens_LargestLoss = 8
    ens_WinAvg = 9
    ens_CL = 10
End Enum

Global Const G_Null = -999999999

'ParmList array elements reserved (in all DLL functions)
Global Const gReservedParms = 4    'Number of reserved parms pass to DLL funcs
Global Const gIndResults = 1       'Element 1: Address to Results array
Global Const gIndCurItem = 2       'Element 2: Current Item Nbr
Global Const gIndErrorNbr = 3      'Element 3: Error number
Global Const gIndErrorMsg = 4      'Element 4: Address to 1 element array
                                   '           storing Error Message
