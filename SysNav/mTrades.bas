Attribute VB_Name = "mTrades"
Option Compare Text
Option Explicit

'Trades structures

'Header (row 1 in trades text file)
Public Enum enth_Trades
    enth_SystemNumber = 0
    enth_SystemName
    enth_BarTimeFrame
    enth_IntraDaySystem
    enth_SystemClass
    enth_StartDate
    enth_EndDate
    enth_TotalBars
    enth_Expenses
    enth_Symbol
    enth_SymbolKey
    enth_TickMove
    enth_TickValue
    enth_TickMinMove
    enth_Margin
    enth_SecurityType
    enth_DefaultUnits
    enth_SessionStart
    enth_SessionEnd
    enth_TimeZoneInfo
    enth_LongStopLoss
    enth_ShortStopLoss
    enth_Cols
End Enum

'Detail rows (one row per entry, one row per exit)
Public Enum entd_Trades
    entd_TradeNbr = 1
    entd_Position
    entd_SignalType
    entd_TradeDate
    entd_Price
    entd_RuleID
    entd_Units
    entd_Profit
    entd_TotalProfit
    entd_AccountBalance
    entd_Equity
    entd_EquityMA
    entd_SkipEqFilter
    entd_FilteredEquity
    entd_UnfilteredEquity
    entd_SymbolIndex
    entd_MaxProfit
    entd_MaxLoss
    entd_BarsInTrade
    entd_Skip
    entd_SkipRpt
    entd_TradeDayOfWeek
    entd_TradeDayOfMonth
    entd_TradeDayOfYear
    entd_SysNbr
    entd_Allocation
    entd_Rank
    entd_EquityAvail
    entd_Conflict
    entd_OpenTrade
    entd_Link
    entd_EntryExitPtr
    entd_SortKey
    entd_TestNumeric
    entd_TestString
    entd_OpenTradesTotal
    entd_Msg
    entd_SignalsTot
    entd_Show
    entd_SignalIndex
    entd_HeaderIndex
    entd_SortKey2
    entd_NumShares
    entd_Cols
End Enum

'Systems structure (This fills gdTable (cSystems class) in Reports.dll)
Public Enum ensy_Systems
    ensy_SystemNumber = 0
    ensy_SystemName
    ensy_Symbol
    ensy_TickMove
    ensy_TickValue
    ensy_MinMoveInTicks
    ensy_DefaultUnits
    ensy_Cols
End Enum
