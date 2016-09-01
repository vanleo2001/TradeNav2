VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStatsSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
'This routine calculates performance statistics.  It performs its calcs
'base on the trades class passed in.  The beginning balance in row 0
'must be set by the caller.  Also, the unit values for each trades must
'be set in the caller.

'Properties used to return calculated summary stats
Private Type mPrivate
    dBeginBalance As Double
    Trades As cTrades
    dMinProfit As Double
    
    'Field addresses from m.Trades
    hPosition As Long
    hSignalType As Long
    hSkip As Long
    hSkipRpt As Long
    hProfit As Long
    hTotalProfit As Long
    hEquity As Long
    hEquityMA As Long
    hSkipEqFilter As Long
    hFilteredEquity As Long
    hUnfilteredEquity As Long
    hTradeDate As Long
    hSignalsTot As Long
    hOpenTradesTot As Long
    hAccountBalance As Long
    hEquityAvail As Long
    hUnits As Long
    hBarsInTrade As Long
    hMaxProfit As Long
    hMaxLoss As Long
    hSorted As Long
    hEntryExitPtr As Long
    hHeaderIndex As Long
    hRuleID As Long
    
    ' Work fields for calculating statistics
    dAccountBalance As Double
    dEquity As Double
    dMaxDrawdownClosed As Double
    lMaxDrawdownClosedLoc As Long
    dMaxDrawdownIntra As Double
    dMaxDrawdownPct As Double
    dEquityPeak As Double
    dLossLargest As Double
    lLossConsec As Long
    lLossRuns As Long
    lLossConsecMax As Long
    lLosses As Long
    dLossAmt As Double
    dPeakToValleyDrawdown As Double
    dWinLargest As Double
    lWinConsec As Long
    lWinConsecMax As Long
    lWins As Long
    lWinRuns As Long
    dWinAmt As Double
    lTotalBars As Long
    dWinRunUp As Double
    dWinRunDown As Double
    dLossRunUp As Double
    dLossRunDown As Double
    dMaxAccountReq As Double
    
    dIntraDrawdown As Double
    bNewEquityPeak As Boolean
    dNetProfit As Double
    dProfit As Double
    lTotalTrades As Long
    lWinBarsInTrades As Long
    lLossBarsInTrades As Long
    dUnits As Double
    dMaxDrawdownInWinner As Double
    dMaxPeakInLoser As Double
    dTotDrawdownInWinners As Double
    dTotPeakInLosers As Double
    
    dOptimalF As Double
    
    dOpenEquity As Double
    lCurrentRun As Long
    
    dStartDate As Double
    dEndDate As Double
    dYearsInSystem As Double
    
    lMovAvgPeriod As Long               ' Period for the equity moving average
    strMAType As String                 ' Type of moving average (Simple, Exponential)
    bEquityFilterOn As Boolean          ' Is the Equity filter on?
    nEquityFilterMode As eGDEquityFilterMode ' Mode of the equity filter
    nTakeNextTrade As eGDTakeNextTradeValue ' Should the user take the next trade based on the equity filter?
    
    'Constants used to reference columns in Monthly Returns table
    tblRetMths As cGdTable
    MthRetResults As gdArrayStatistics
    
    'TLB 2/24/2011: Sharpe Ratio stuff
    dMonthlyProfitAvg As Double
    dMonthlyProfitStdDev As Double
    dSharpeRatioMonthly As Double
End Type
Private m As mPrivate

Private Const mMR_TradeDate = 0
Private Const mMR_AccBal = 1
Private Const mMR_NumTrades = 2
Private Const mMR_ReturnPct = 3
Private Const mMR_Profit = 4

Public Sub CopyTradeHandles(pTradesDtlHandle As Long, pTradesHdrHandle As Long)
    
    If m.Trades Is Nothing Then
        Set m.Trades = New cTrades
    End If
    
    m.Trades.CopyTableFromHandle pTradesDtlHandle
    m.Trades.CopyHdrFromHandle pTradesHdrHandle

End Sub

Public Sub SetTradeHandles(pDtlHandle As Long, pHdrHandle As Long)
    
    If m.Trades Is Nothing Then
        Set m.Trades = New cTrades
    End If
    
    m.Trades.SetTableFromHandle pDtlHandle, False
    m.Trades.SetHdrFromHandle pHdrHandle, False

End Sub

Public Sub SetSortHandle(ByVal hSortHandle As Long)
    m.Trades.SortHandle = hSortHandle
End Sub

Property Get MinProfit() As Double
    MinProfit = m.dMinProfit
End Property
Property Let MinProfit(pData As Double)
    m.dMinProfit = pData
End Property

Public Property Get StartDate() As Double
    StartDate = m.dStartDate
End Property
Public Property Let StartDate(ByVal dStartDate As Double)
    m.dStartDate = dStartDate
End Property

Public Property Get EndDate() As Double
    EndDate = m.dEndDate
End Property
Public Property Let EndDate(ByVal dEndDate As Double)
    m.dEndDate = dEndDate
End Property

Public Property Get BeginBalance() As Double
    BeginBalance = m.dBeginBalance
End Property
Public Property Let BeginBalance(ByVal dBalance As Double)
    m.dBeginBalance = dBalance
End Property

Public Property Get MovAvgPeriod() As Long
    MovAvgPeriod = m.lMovAvgPeriod
End Property
Public Property Let MovAvgPeriod(ByVal lPeriod As Long)
    m.lMovAvgPeriod = lPeriod
End Property

Public Property Get MovAvgType() As String
    MovAvgType = m.strMAType
End Property
Public Property Let MovAvgType(ByVal strType As String)
    m.strMAType = strType
End Property

Public Property Get EquityFilterOn() As Boolean
    EquityFilterOn = m.bEquityFilterOn
End Property
Public Property Let EquityFilterOn(ByVal bEquityFilterOn As Boolean)
    m.bEquityFilterOn = bEquityFilterOn
End Property

Public Property Get EquityFilterMode() As eGDEquityFilterMode
    EquityFilterMode = m.nEquityFilterMode
End Property
Public Property Let EquityFilterMode(ByVal nEquityFilterMode As eGDEquityFilterMode)
    m.nEquityFilterMode = nEquityFilterMode
End Property

Public Property Get TakeNextTrade() As eGDTakeNextTradeValue
    TakeNextTrade = m.nTakeNextTrade
End Property

'Overall stats
Property Get EndingBalance() As Double
    EndingBalance = m.dAccountBalance
End Property
Property Get TotTrades() As Long
    TotTrades = m.lTotalTrades
End Property
Property Get MaxDrawdownClosed() As Double
    MaxDrawdownClosed = m.dMaxDrawdownClosed
End Property
Property Get MaxDrawdownLocation() As Long
    MaxDrawdownLocation = m.lMaxDrawdownClosedLoc
End Property
Property Get MaxDrawdownIntra() As Double
    MaxDrawdownIntra = m.dMaxDrawdownIntra
End Property
Property Get MaxDrawdownPct() As Double
    MaxDrawdownPct = m.dMaxDrawdownPct
End Property
Property Get PeakToValleyDrawdown() As Double
    PeakToValleyDrawdown = Abs(m.dPeakToValleyDrawdown)
End Property
Property Get EquityPeak() As Double
    EquityPeak = m.dEquityPeak
End Property
Property Get AccountSizeReq() As Double
    AccountSizeReq = m.dMaxAccountReq
End Property
Property Get OpenEquity() As Double
    OpenEquity = m.dOpenEquity
End Property
Property Get CurrentRun() As Long
    CurrentRun = m.lCurrentRun
End Property

Property Get OptimalF() As Double
    OptimalF = m.dOptimalF
End Property

' >=  2 Winners tend to follow losers, losers follow winners
' <= -2 Winners tend to follow winners, losers follow losers
Property Get ZScore() As Double
On Error GoTo ErrSection:

    Dim dDivisor As Double
    Dim n As Double
    Dim X As Double
    Dim R As Double
    Dim xWins As Double
    Dim xLosses As Double
    
    ZScore = 0
    
    'X = Wins * Losses * 2      'overflow (Wins and Losses are Long values)
    xWins = Wins
    xLosses = Losses
    X = xWins * xLosses * 2
    n = TotTrades
    
    ZScore = 0
    If n - 1 > 0 Then dDivisor = (((X * (X - n)) / (n - 1)) ^ (1 / 2))
    If dDivisor <> 0 Then
        ZScore = (n * ((WinRuns + LossRuns) - 0.5) - X) / dDivisor
    Else
        ZScore = 0
    End If
    
ErrExit:
    Exit Property

ErrSection:
    If Err.Number = 6 Then Resume ErrExit
    RaiseError "cStatsSummary.ZScore.Get", eGDRaiseError_Raise, frmReports.AppPath

End Property

Property Get PLRatio() As Double
    
    Dim dLossAvg As Double
    
    dLossAvg = LossAvg
    If dLossAvg <> 0 Then
        PLRatio = Abs(WinAvg / dLossAvg)
    Else
        PLRatio = WinAvg
    End If

End Property

Property Get ReturnPctMM() As Double
    
    If m.dBeginBalance > 0 Then
        ReturnPctMM = (m.dAccountBalance - m.dBeginBalance) / m.dBeginBalance
    End If

End Property

'Property Get ROIPct() As Double
Property Get ReturnPct() As Double
On Error GoTo ErrSection:

    If m.dMaxAccountReq <> 0 Then
        ReturnPct = TotalNetProfit / m.dMaxAccountReq
    Else
        ReturnPct = 0#
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cStatsSummary.ReturnPct", eGDRaiseError_Raise
    
End Property

Property Get MonthlyProfitAvg() As Double
    MonthlyProfitAvg = m.dMonthlyProfitAvg
End Property
Property Get MonthlyProfitStdDev() As Double
    MonthlyProfitStdDev = m.dMonthlyProfitStdDev
End Property
Property Get SharpeRatioMonthly() As Double
    SharpeRatioMonthly = m.dSharpeRatioMonthly
End Property
Property Get SharpeRatioAnnualized() As Double
    SharpeRatioAnnualized = m.dSharpeRatioMonthly * Sqr(12)
End Property
Property Get CalmarRatio() As Double
    If m.dMaxDrawdownIntra < 0 And m.dMonthlyProfitAvg > 0 Then
        CalmarRatio = (m.dMonthlyProfitAvg * 12) / Abs(m.dMaxDrawdownIntra)
    End If
End Property

' TLB 3/4/2011: we should just deal with the dollar returns rather than %'s
' (since %'s don't make any sense when testing 1 contract per trade).
' So:  Monthly Sharpe Ratio = (Avg Monthly Profit - Avg Riskfree Monthly Profit) / StdDev of Monthly Profit
'   where Avg Riskfree Monthly Profit = Max Acct Required * Avg Tbill% / 12
' and: Annualized Sharpe Ratio = Monthly Sharpe Ratio * sqrt(12)
Public Sub CalcSharpeRatio()
On Error GoTo ErrSection:
     
    Dim s$, iBar&, iFirst&, nCount&, dIRX#
    Dim MonthlyProfitStats As gdArrayStatistics
    
    m.dMonthlyProfitAvg = 0
    m.dMonthlyProfitStdDev = 0
    m.dSharpeRatioMonthly = 0
    
If IsIDE Then
    s = m.tblRetMths.ToString(vbCrLf, vbTab)
    FileFromString "c:\dvlp\months.txt", s, True
End If
        
    ' Calculate standard statistics on Monthly Profits array
    ' (if enough of a sample for stats to be valid)
    If m.tblRetMths.NumRecords >= 3 Then
        gdCalcStatistics m.tblRetMths.FieldArrayHandle(mMR_Profit), MonthlyProfitStats, True, 0, m.tblRetMths.NumRecords - 1
        m.dMonthlyProfitAvg = MonthlyProfitStats.Average
        m.dMonthlyProfitStdDev = MonthlyProfitStats.StdDev
        If m.dMonthlyProfitStdDev > 0 And m.tblRetMths.NumRecords >= 12 Then
            ' get average $IRX (t-bills rate)
            dIRX = 0
            If Not g.IrxBars Is Nothing Then
                If g.IrxBars.Size > 0 Then
                    nCount = 0
                    iFirst = g.IrxBars.FindDateTime(m.tblRetMths(mMR_TradeDate, 0))
                    iBar = g.IrxBars.FindDateTime(m.tblRetMths(mMR_TradeDate, m.tblRetMths.NumRecords - 1))
                    For iBar = iBar To 0 Step -1
                        If g.IrxBars(eBARS_Close, iBar) > 0 Then
                            dIRX = dIRX + g.IrxBars(eBARS_Close, iBar)
                            nCount = nCount + 1
                        End If
                        If iBar <= iFirst And nCount > 0 Then
                            Exit For
                        End If
                    Next
                    If nCount > 1 Then
                        dIRX = dIRX / nCount ' convert to average
                    End If
                    dIRX = dIRX / 1200 ' convert to a monthly interest rate
                End If
            End If
            ' subtract the avg monthly risk-free profit from the avg monthly system profit
            ' and divide that by the std dev of the monthly profits
            m.dSharpeRatioMonthly = (m.dMonthlyProfitAvg - m.dMaxAccountReq * dIRX) / m.dMonthlyProfitStdDev
        End If
    End If
        
ErrExit:
    Exit Sub

ErrSection:
    If Err.Number = 6 Then Resume ErrExit
    RaiseError "cStatsSummary.CalcSharpeRatio", eGDRaiseError_Raise, frmReports.AppPath

End Sub

' TLB 2/24/2011: this older original code does NOT seem to be working correctly.
'
'Calculates the SharpeRatio.  A measure of risk-adjusted return of an investment.
'Values of 2.0 or greater are considered good (especially for futures trading
'systems).  3.0 or greater outstanding.
Property Get SharpeRatio(Optional ByVal Years As Integer = 0, Optional ByVal RiskFreeRate As Single = 0.04) As Single
On Error GoTo ErrSection:

    Dim dPortfolioReturn As Double
    Dim dPortfolioStd As Double
    Dim MthRetResults As gdArrayStatistics
    Dim lIndex As Long
    Dim lCount As Long
    Dim lMonths As Long
    Dim lFromRec As Long
    
    SharpeRatio = 0
    
    ' Take monthly average returns * 12 (to annualize)
    dPortfolioReturn = CompoundedAvgAnnualROR(Years)
    
    ' Calculate Monthly standard deviation lIndex Sqr(12) to annualize
    If Years = 0 Then
        dPortfolioStd = MthReturnStd() * Sqr(12)
    Else
        ' Calculate returns and std for recent n years ("Years" parm)
        ' Search back for first monthly return record...
        lMonths = Years * 12
        lFromRec = 1
        For lIndex = m.tblRetMths.NumRecords - 1 To 1 Step -1
            lCount = lCount + 1
            If lCount > lMonths Then
                lFromRec = lIndex
                Exit For
            End If
        Next lIndex
        
        ' Calculate standard statistics on Monthly returns array
        gdCalcStatistics m.tblRetMths.FieldArrayHandle(mMR_ReturnPct), _
            MthRetResults, 0, lFromRec, m.tblRetMths.NumRecords - 1
        dPortfolioStd = MthRetResults.StdDev * Sqr(12)
    End If
    
    
    If dPortfolioStd <> 0 Then
        SharpeRatio = (dPortfolioReturn - RiskFreeRate) / dPortfolioStd
    Else
        SharpeRatio = 0
    End If
    
ErrExit:
    Exit Property

ErrSection:
    If Err.Number = 6 Then Resume ErrExit
    RaiseError "cStatsSummary.SharpeRatio.Get", eGDRaiseError_Raise, frmReports.AppPath

End Property

'This is the rate of return which, if compounded over the iYears covered by
'the performance history, would yield the cumulative gain.  For example, if
'an account goes from 1,000 to 1,500 over 2 years, the compounded Avg annual
'rate would be 22.5 return/year.  Yr1: 1,000 * .225 = 1,250,
'1,250 * .225 = 1,500.
Public Function CompoundedAvgAnnualROR(Optional iYears As Integer = 0) As Single
On Error GoTo ErrSection:
    
    Dim lStartDatePos As Long
    Dim dReturnPct As Double
    Dim dYears As Double
    Dim dStartDate As Double
    Dim lIndex As Long
    
    CompoundedAvgAnnualROR = 0
    
    'Calculate for entire time period
    If iYears = 0 Then
        If ReturnPctMM > 0 Then
            CompoundedAvgAnnualROR = ((1 + ReturnPctMM) ^ (1 / m.dYearsInSystem)) - 1
        End If
    Else
        'Calculate exact time in years from most recent date back n years.
        dStartDate = DateAdd("yyyy", iYears * -1, m.Trades.Num(m.Trades.NumRecords - 1, entd_TradeDate))
        lStartDatePos = 0
        For lIndex = m.Trades.NumRecords - 1 To 0 Step -1
            If gdGetNum(m.hTradeDate, lIndex) < dStartDate Then
                lStartDatePos = lIndex + 1
                Exit For
            End If
        Next lIndex
        dYears = (m.Trades.Num(m.Trades.NumRecords - 1, entd_TradeDate) - _
               m.Trades.Num(lStartDatePos, entd_TradeDate)) / 365.25
        
        'Calculate cumulative return pct...
        dReturnPct = (m.Trades.Num(m.Trades.NumRecords - 1, entd_AccountBalance) - m.Trades.Num(lStartDatePos, entd_AccountBalance)) / m.Trades.Num(lStartDatePos, entd_AccountBalance)
        
        If dReturnPct > 0 Then
            CompoundedAvgAnnualROR = ((1 + dReturnPct) ^ (1 / dYears)) - 1
        End If
    End If
    
ErrExit:
    Exit Function

ErrSection:
    If Err.Number = 6 Then Resume ErrExit
    RaiseError "cStatsSummary.CompoundedAvgAnnualROR", eGDRaiseError_Raise, frmReports.AppPath

End Function

'Avg annual rate of return over n years divided by the
'Avg max annual drawdown in those n years + 10%
Property Get SterlingRatioCust(Optional ByVal iYears As Integer = 0) As Single
On Error GoTo ErrSection:
    
    Dim dPV As Double
    
    SterlingRatioCust = 0
    
    'Years = 0 means take the entire test period
    If iYears = 0 Then
        dPV = PeakToValleyDrawdown() + 0.1
    Else
        dPV = DrawdownPVAvg(iYears) + 0.1
    End If
    
    If dPV <> 0 Then
        If iYears = 0 Then
            SterlingRatioCust = CompoundedAvgAnnualROR / dPV
        Else
            SterlingRatioCust = CompoundedAvgAnnualROR(iYears) / dPV
        End If
    Else
        SterlingRatioCust = 0
    End If
    
ErrExit:
    Exit Property

ErrSection:
    If Err.Number = 6 Then Resume ErrExit
    RaiseError "cStatsSummary.SterlingRatioCust", eGDRaiseError_Raise, frmReports.AppPath

End Property

'Calculates the average drawdown (Peak to valley) over a given number of
'years.  Each period is broken into years and the drawdown calculated for
'each years.  The average of the drawdown of all years is then calced.
Private Function DrawdownPVAvg(Optional ByVal Years As Integer = 0) As Single
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim EquityPeak As Double
    Dim Valley As Double
    Dim ValleyMth As Double
    Dim StartDatePos As Long
    Dim Yrs As Single
    Dim StartYear As Long
    Dim Depth(3) As Single
    Dim LastRec As Long
    Dim R As Long
    Dim StartDate As Double
    Dim YearChange As Double
    Dim total As Double
    
    DrawdownPVAvg = 0
    LastRec = m.Trades.NumRecords - 1
    
    'Go back 3 years from the last trade in the trades array
    StartDate = DateAdd("yyyy", -1 * Years, gdGetNum(m.hTradeDate, LastRec))
    StartDatePos = 0
    For X = m.Trades.NumRecords - 1 To 0 Step -1
        If gdGetNum(m.hTradeDate, X) < StartDate Then
            StartDatePos = X + 1
            Exit For
        End If
    Next X
    
    R = 1
    YearChange = DateAdd("yyyy", 1, gdGetNum(m.hTradeDate, StartDatePos))
    For X = StartDatePos To LastRec
    
        'Index for keeping track of Depth% (peak to valley dips)
        If gdGetNum(m.hTradeDate, X) > YearChange Then
            R = R + 1
            YearChange = DateAdd("yyyy", 1, gdGetNum(m.hTradeDate, X))
        End If
        
        'Check for new equity peak when a loss occurs.
        If gdGetNum(m.hAccountBalance, X) > EquityPeak Then
        
            'Record peak to valley percentage change and distance
            'in mths from start to end
            If Valley < EquityPeak Then
                If (EquityPeak - Valley) / EquityPeak > Depth(R) Then
                    Depth(R) = (EquityPeak - Valley) / EquityPeak
                End If
            End If
            
            'Initialize a new drawdown period
            Valley = gdGetNum(m.hAccountBalance, X)
            EquityPeak = gdGetNum(m.hAccountBalance, X)
        Else
            If gdGetNum(m.hAccountBalance, X) < Valley Then
                Valley = gdGetNum(m.hAccountBalance, X)
            End If
        End If
        
    Next X
    
    For X = 1 To R
        total = total + Depth(X)
    Next X
    If R > 0 Then
        DrawdownPVAvg = total / R
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cStatsSummary.DrawdownPVAvg", eGDRaiseError_Raise, g.strAppPath

End Function

'*** NOT USED ****
'Avg annual rate of return past 3 calendar years divided by the
'Avg max annual drawdown in those 3 years + 10%
Property Get SterlingRatio()
On Error GoTo ErrSection:

    Dim PV          As Double
    
    PV = PeakToValleyDrawdown() - 0.1
    If PV <> 0 Then
        SterlingRatio = CompoundedAvgAnnualROR() / PV
    Else
        SterlingRatio = 0
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStatsSummary.SterlingRatio", eGDRaiseError_Raise, g.strAppPath

End Property

'This is called to calculate stats for trade mgt section on money mgt rpt
Public Sub CalcTradeMgt(pPctConflicts As Single, pAvgAlloc As Single, _
    pAvgOpenPos As Single, pPctSkipped As Single, pMaxAlloc As Single, _
    pMaxOpen As Single)
On Error GoTo ErrSection:
    
    Dim ConflictSignals As Long
    Dim TotTrades As Long
    Dim OpenPosTot As Long
    Dim OpenPosCnt As Long
    Dim AllocAmt As Single
    Dim AllocCnt As Long
    Dim TotSkipped As Long
    Dim TotEntries As Long
    Dim lIndex As Long
    
    For lIndex = 1 To m.Trades.NumRecords - 1
        If gdGetNum(m.hTradeDate, lIndex) >= m.dStartDate And gdGetNum(m.hTradeDate, lIndex) <= m.dEndDate Then
            If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                If gdGetNum(m.hSkip, lIndex) = 0 And gdGetNum(m.hSkipRpt, lIndex) = 0 Then
                    If gdGetNum(m.hSignalsTot, lIndex) > 0 Then
                        ConflictSignals = ConflictSignals + 1
                    End If
                    If gdGetNum(m.hOpenTradesTot, lIndex) > 0 Then
                        OpenPosTot = OpenPosTot + gdGetNum(m.hOpenTradesTot, lIndex)
                        OpenPosCnt = OpenPosCnt + 1
                        If gdGetNum(m.hOpenTradesTot, lIndex) > pMaxOpen Then
                            pMaxOpen = gdGetNum(m.hOpenTradesTot, lIndex)
                        End If
                    End If
                    AllocAmt = AllocAmt + ((gdGetNum(m.hAccountBalance, lIndex) - gdGetNum(m.hEquityAvail, lIndex)) / gdGetNum(m.hAccountBalance, lIndex))
                    AllocCnt = AllocCnt + 1
                    If ((gdGetNum(m.hAccountBalance, lIndex) - gdGetNum(m.hEquityAvail, lIndex)) / gdGetNum(m.hAccountBalance, lIndex)) > pMaxAlloc Then
                        pMaxAlloc = (gdGetNum(m.hAccountBalance, lIndex) - gdGetNum(m.hEquityAvail, lIndex)) / gdGetNum(m.hAccountBalance, lIndex)
                    End If
                Else
                    TotSkipped = TotSkipped + 1
                End If
                TotEntries = TotEntries + 1
            End If
        End If
    Next lIndex
    
    'Pct conflict signals
    If TotEntries > 0 Then
        pPctConflicts = ConflictSignals / TotEntries
    Else
        pPctConflicts = 0
    End If
    
    'Pct skipped signals
    If TotEntries > 0 Then
        pPctSkipped = TotSkipped / TotEntries
    Else
        pPctSkipped = 0
    End If
    
    'Avg open positions
    If OpenPosCnt > 0 Then
        pAvgOpenPos = OpenPosTot / OpenPosCnt
    Else
        pAvgOpenPos = 0
    End If
    
    'Avg equity available pct
    If AllocCnt > 0 Then
        pAvgAlloc = AllocAmt / AllocCnt
    Else
        pAvgAlloc = 0
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.CalcTradeMgt", eGDRaiseError_Raise, g.strAppPath

End Sub

Property Get TotalNetProfit() As Double
    TotalNetProfit = m.dWinAmt + m.dLossAmt
End Property

Property Get CPCIndex() As Double
    CPCIndex = PLRatio * WinPct * ProfitFactor
End Property

Property Get Expectancy() As Double
    
    If LossAvg <> 0 Then
        Expectancy = AvgTrade / Abs(LossAvg)
    Else
        Expectancy = 0
    End If

End Property

Property Get KellyRatio() As Double
    
    If PLRatio <> 0 Then
        KellyRatio = ((PLRatio + 1) * WinPct - 1) / PLRatio
    Else
        KellyRatio = 0
    End If

End Property

Property Get AvgTrade() As Double
    
    If m.lTotalTrades <> 0 Then
        AvgTrade = (m.dWinAmt + m.dLossAmt) / m.lTotalTrades
    Else
        AvgTrade = 0
    End If

End Property

Property Get ProfitFactor() As Double
    
    ' Make sure that the Profit Factor is zero if there are no Winning trades
    ' or if there are no trades at all (DAJ: 04/25/2003)
    If m.dWinAmt = 0 Then
        ProfitFactor = 0
        
    ' If there were winners but no loseres, set the Profit Factor to 999 since
    ' it is essentially infinite (DAJ: 04/25/2003)
    ElseIf m.dLossAmt = 0 Then
        ProfitFactor = 999
    
    Else
        ProfitFactor = Abs(m.dWinAmt / m.dLossAmt)
    End If

End Property

Property Get AvgBarsInTrade() As Double
    
    If m.lTotalTrades <> 0 Then
        AvgBarsInTrade = (m.lWinBarsInTrades + m.lLossBarsInTrades) / m.lTotalTrades
    Else
        AvgBarsInTrade = 0
    End If

End Property

Property Get AvgTradesPerYear() As Double
    
    If m.dYearsInSystem <> 0 Then
        AvgTradesPerYear = m.lTotalTrades / m.dYearsInSystem
    Else
        AvgTradesPerYear = 0
    End If
    
End Property

Property Get PctInMarket() As Double
    
    If m.lTotalBars <> 0 Then
        PctInMarket = (m.lWinBarsInTrades + m.lLossBarsInTrades) / m.lTotalBars
    Else
        PctInMarket = 0
    End If

End Property

' Properties returning winning stats
Property Get Wins() As Long
    Wins = m.lWins
End Property

Property Get WinPct() As Double
    
    If m.lWins + m.lLosses > 0 Then
        WinPct = m.lWins / (m.lWins + m.lLosses)
    Else
        WinPct = 0
    End If

End Property

Property Get WinAvg() As Double
    
    If m.lWins > 0 Then
        WinAvg = m.dWinAmt / m.lWins
    Else
        WinAvg = 0
    End If

End Property

Property Get WinLargest() As Double
    WinLargest = m.dWinLargest
End Property

Property Get WinTotal() As Double
    WinTotal = m.dWinAmt
End Property

Property Get WinConsec() As Long
    WinConsec = m.lWinConsecMax
End Property

Property Get WinRuns() As Long
    WinRuns = m.lWinRuns
End Property

Property Get WinAvgBars() As Double
    
    If m.lWins <> 0 Then
        WinAvgBars = m.lWinBarsInTrades / m.lWins
    Else
        WinAvgBars = 0
    End If

End Property

Property Get WinMaxDrawdown() As Double
    
    If m.dMaxDrawdownInWinner = 999999999 Then
        WinMaxDrawdown = 0
    Else
        WinMaxDrawdown = m.dMaxDrawdownInWinner
    End If

End Property

Property Get WinAvgDrawdown() As Double
    
    If m.lWins > 0 Then
        WinAvgDrawdown = m.dTotDrawdownInWinners / m.lWins
    Else
        WinAvgDrawdown = 0
    End If

End Property

' Properties returning losing stats
Property Get Losses() As Long
    Losses = m.lLosses
End Property

Property Get LossPct() As Double
    
    If m.lLosses + m.lWins > 0 Then
        LossPct = m.lLosses / (m.lLosses + m.lWins)
    Else
        LossPct = 0
    End If

End Property

Property Get LossAvg() As Double
    
    If m.lLosses > 0 Then
        LossAvg = m.dLossAmt / m.lLosses
    Else
        LossAvg = 0
    End If

End Property

Property Get LossLargest() As Double
    LossLargest = m.dLossLargest
End Property

Property Get LossTotal() As Double
    LossTotal = m.dLossAmt
End Property

Property Get LossConsec() As Long
    LossConsec = m.lLossConsecMax
End Property

Property Get LossRuns() As Long
    LossRuns = m.lLossRuns
End Property

Property Get LossAvgBars() As Double
    
    If m.lLosses <> 0 Then
        LossAvgBars = m.lLossBarsInTrades / m.lLosses
    Else
        LossAvgBars = 0
    End If

End Property

Property Get LossMaxPeak() As Double
    
    If m.dMaxPeakInLoser = -999999999 Then
        LossMaxPeak = 0
    Else
        LossMaxPeak = m.dMaxPeakInLoser
    End If

End Property

Property Get LossAvgPeak() As Double
    
    If m.lLosses > 0 Then
        LossAvgPeak = m.dTotPeakInLosers / m.lLosses
    Else
        LossAvgPeak = 0
    End If

End Property

Property Get WinRunUpAvg() As Double
    
    If m.lWins > 0 Then
        WinRunUpAvg = m.dWinRunUp / m.lWins
    Else
        WinRunUpAvg = 0
    End If

End Property

Property Get WinRunDownAvg() As Double
    
    If m.lWins > 0 Then
        WinRunDownAvg = m.dWinRunDown / m.lWins
    Else
        WinRunDownAvg = 0
    End If

End Property

Property Get LossRunUpAvg() As Double
    
    If m.lLosses > 0 Then
        LossRunUpAvg = m.dLossRunUp / m.lLosses
    Else
        LossRunUpAvg = 0
    End If

End Property

Property Get LossRunDownAvg() As Double
    
    If m.lLosses > 0 Then
        LossRunDownAvg = m.dLossRunDown / m.lLosses
    Else
        LossRunDownAvg = 0
    End If

End Property

Property Get MthReturnValuesHandle() As Long
    MthReturnValuesHandle = m.tblRetMths.TableHandle
End Property

Property Get AvgYearlyReturn() As Double
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim dYearAgoBalance As Double
    Dim lYear As Long
    Dim lYearReturn As Double
    Dim lYears As Long
    Dim dYearReturnTotal As Double

    dYearAgoBalance = gdGetNum(m.hAccountBalance, 0)
    
    For lIndex = 1 To m.tblRetMths.NumRecords - 1
        If Year(CDate(m.tblRetMths(mMR_TradeDate, lIndex))) > lYear Then
            If dYearAgoBalance > 0 Then
                lYearReturn = (m.tblRetMths(mMR_AccBal, lIndex) - dYearAgoBalance) / dYearAgoBalance
            End If
            lYears = lYears + 1
            dYearAgoBalance = m.tblRetMths(mMR_AccBal, lIndex)
            dYearReturnTotal = dYearReturnTotal + lYearReturn
            lYear = Year(CDate(m.tblRetMths(mMR_TradeDate, lIndex)))
        End If
    Next lIndex
    
    If lYears <> 0 Then
        AvgYearlyReturn = dYearReturnTotal / lYears
    Else
        AvgYearlyReturn = 0
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.AvgYearlyReturn.Get", eGDRaiseError_Raise, frmReports.AppPath

End Property

Property Get MthReturnMax() As Double
    
    Dim a As cGdArray
    
    Set a = m.tblRetMths.FieldArray(mMR_ReturnPct, True)
    MthReturnMax = a.MaxValue

End Property

Property Get MthReturnMin() As Double
    
    Dim a As cGdArray
    
    Set a = m.tblRetMths.FieldArray(mMR_ReturnPct, True)
    MthReturnMin = a.MinValue

End Property

Property Get MthReturnStd() As Single
    MthReturnStd = m.MthRetResults.StdDev
End Property

Property Get MthReturnAvg() As Double
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim dTotal As Double
    
    For lIndex = 1 To m.tblRetMths.NumRecords - 1
        dTotal = dTotal + m.tblRetMths.Num(mMR_ReturnPct, lIndex)
    Next lIndex
    
    If m.tblRetMths.NumRecords > 0 Then
        MthReturnAvg = dTotal / m.tblRetMths.NumRecords
    Else
        MthReturnAvg = 0
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.MthReturnAvg", eGDRaiseError_Raise, frmReports.AppPath

End Property

Property Get MthReturnWinAvg() As Double
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lCount As Long
    Dim dTotal As Double
    
    For lIndex = 1 To m.tblRetMths.NumRecords - 1
        If m.tblRetMths.Num(mMR_ReturnPct, lIndex) > 0 Then
            lCount = lCount + 1
            dTotal = dTotal + m.tblRetMths.Num(mMR_ReturnPct, lIndex)
        End If
    Next lIndex
    
    If lCount > 0 Then
        MthReturnWinAvg = dTotal / lCount
    Else
        MthReturnWinAvg = 0
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStatsSummary.MthReturnWinAvg", eGDRaiseError_Raise, frmReports.AppPath

End Property

Property Get MthReturnLossAvg() As Double
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lCount As Long
    Dim dTotal As Double
    
    For lIndex = 1 To m.tblRetMths.NumRecords - 1
        If m.tblRetMths.Num(mMR_ReturnPct, lIndex) < 0 Then
            lCount = lCount + 1
            dTotal = dTotal + m.tblRetMths.Num(mMR_ReturnPct, lIndex)
        End If
    Next lIndex
    
    If lCount > 0 Then
        MthReturnLossAvg = dTotal / lCount
    Else
        MthReturnLossAvg = 0
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.MthReturnLossAvg", eGDRaiseError_Raise, frmReports.AppPath

End Property

Property Get MthReturnPctOfWinMths() As Double
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lCount As Long
    Dim dTotal As Double
    
    For lIndex = 1 To m.tblRetMths.NumRecords - 1
        If m.tblRetMths.Num(mMR_ReturnPct, lIndex) > 0 Then
            lCount = lCount + 1
        End If
        If m.tblRetMths.Num(mMR_NumTrades, lIndex) > 0 Then
            dTotal = dTotal + 1
        End If
    Next lIndex
    
    If dTotal > 0 Then
        MthReturnPctOfWinMths = lCount / dTotal
    Else
        MthReturnPctOfWinMths = 0
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.MthReturnPctOfWinMths", eGDRaiseError_Raise, frmReports.AppPath

End Property

'Driver for calculating statistics on a series of trades
Public Sub Calculate(Optional ByVal bCalcMthReturns As Boolean = False, Optional ByVal bCalcOptimalf As Boolean = True, Optional ByVal bCalcMovingAverage As Boolean = True, Optional ByVal bUseSkipEqFilter As Boolean = False)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    m.dNetProfit = 0
    m.dEquityPeak = m.dBeginBalance
    m.dPeakToValleyDrawdown = 0
    m.dIntraDrawdown = 0
    m.dMaxDrawdownClosed = 0
    m.dMaxDrawdownIntra = 0
    m.dMaxAccountReq = 0
    
    'Retrieve Field address from Trades class
    With m.Trades
        m.hPosition = .FieldHandle(entd_Position)
        m.hSignalType = .FieldHandle(entd_SignalType)
        m.hSkip = .FieldHandle(entd_Skip)
        m.hSkipRpt = .FieldHandle(entd_SkipRpt)
        m.hProfit = .FieldHandle(entd_Profit)
        m.hTotalProfit = .FieldHandle(entd_TotalProfit)
        m.hEquity = .FieldHandle(entd_Equity)
        m.hEquityMA = .FieldHandle(entd_EquityMA)
        m.hSkipEqFilter = .FieldHandle(entd_SkipEqFilter)
        m.hFilteredEquity = .FieldHandle(entd_FilteredEquity)
        m.hUnfilteredEquity = .FieldHandle(entd_UnfilteredEquity)
        m.hTradeDate = .FieldHandle(entd_TradeDate)
        m.hSignalsTot = .FieldHandle(entd_SignalsTot)
        m.hOpenTradesTot = .FieldHandle(entd_OpenTradesTotal)
        m.hAccountBalance = .FieldHandle(entd_AccountBalance)
        m.hEquityAvail = .FieldHandle(entd_EquityAvail)
        m.hUnits = .FieldHandle(entd_Units)
        m.hBarsInTrade = .FieldHandle(entd_BarsInTrade)
        m.hMaxProfit = .FieldHandle(entd_MaxProfit)
        m.hMaxLoss = .FieldHandle(entd_MaxLoss)
        m.hEntryExitPtr = .FieldHandle(entd_EntryExitPtr)
        m.hHeaderIndex = .FieldHandle(entd_HeaderIndex)
        m.hRuleID = .FieldHandle(entd_RuleID)
        m.hSorted = .SortHandle
    End With
    
    m.lTotalBars = 0&
    For lIndex = 1 To m.Trades.HeaderNumRecords - 1
        m.lTotalBars = m.lTotalBars + m.Trades.NumHdr(lIndex, enth_TotalBars)
    Next lIndex
    
    m.lTotalTrades = 0
    m.lWins = 0: m.dWinAmt = 0: m.lLosses = 0: m.dLossAmt = 0
    m.dWinLargest = 0: m.lWinConsecMax = 0
    m.dLossLargest = 0: m.lLossConsecMax = 0
    m.lWinConsec = 0: m.lLossConsec = 0
    
    'Begin balance is always in row 0
    If gdGetNum(m.hAccountBalance, 0) < 0 Then
        'gdSetNum m.hAccountBalance, 0, 0
    End If
    m.dBeginBalance = gdGetNum(m.hAccountBalance, 0)
    
    m.dAccountBalance = m.dBeginBalance
    m.dEquity = 0#
    m.lWinBarsInTrades = 0
    m.lLossBarsInTrades = 0
    m.dMaxDrawdownInWinner = 999999999
    m.dMaxPeakInLoser = -999999999
    m.dTotDrawdownInWinners = 0
    m.dTotPeakInLosers = 0
    m.dWinRunUp = 0: m.dWinRunDown = 0: m.dLossRunUp = 0: m.dLossRunDown = 0
    
    ProcessTradeResults bUseSkipEqFilter
    
    'Retrieve the end of month balances
    If bCalcMthReturns Then CalcEndOfMthBalances
    m.dYearsInSystem = (m.dEndDate - m.dStartDate) / 365.25
    If bCalcOptimalf Then m.dOptimalF = CalculateOptimalF
    
    If bCalcMovingAverage Then
        EquityMovingAverage
        
        If m.bEquityFilterOn Then
            Calculate bCalcMthReturns, bCalcOptimalf, False, True
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    If Err.Number = 6 Then
        Resume ErrExit
    Else
        RaiseError "cStatsSummary.Calculate", eGDRaiseError_Raise, frmReports.AppPath
    End If

End Sub

Private Sub ProcessTradeResults(ByVal bUseSkipEqFilter As Boolean)
On Error GoTo ErrSection:
    
    Dim lLast As Long
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim dMaxAccount As Double
    Dim dDrawPct As Double
    Dim alOpen As New cGdArray
    
    If m.Trades Is Nothing Then Err.Raise gUserErr, , "No Trades Found"
    
    ' If the Last trade is still open, ignore the Last Entry and Exit for
    ' summary calculations (DAJ: 03/26/2003)...
    lLast = m.Trades.NumRecords - 1
    m.dOpenEquity = 0#
    alOpen.Create eGDARRAY_Longs
    
    For lIndex2 = 1 To m.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        If gdGetNum(m.hSkip, lIndex) = 0 And gdGetNum(m.hSkipRpt, lIndex) = 0 Then
            If gdGetNum(m.hSignalType, lIndex) = gExitSignal Then
                If m.Trades.Item(lIndex, entd_RuleID) = 0 Then
                    gdSetNum m.hSkip, lIndex, 1#
                    gdSetNum m.hSkip, gdGetNum(m.hEntryExitPtr, lIndex), 1#
                    
                    m.dOpenEquity = m.dOpenEquity + gdGetNum(m.hProfit, lIndex)
                    
                    alOpen.Add gdGetNum(m.hEntryExitPtr, lIndex)
                    alOpen.Add lIndex
                End If
            End If
        End If
    Next lIndex2

    If lLast <= 0 Then Err.Raise gUserErr, , "No Trades Found"
    
    'Clear static variables
    CheckDrawdown True, 0
    
    'Close #1
    'Open "c:\my data\portfolionavigator\reports1.txt" For Output As #1
    'Print #1, "lIndex" & vbTab & "AccBal" & vbTab; "Skip"
    dMaxAccount = 0#
    For lIndex2 = 1 To lLast
        lIndex = gdGetNum(m.hSorted, lIndex2)
        m.dUnits = gdGetNum(m.hUnits, lIndex) / m.Trades.ItemHdr(gdGetNum(m.hHeaderIndex, lIndex), enth_DefaultUnits)
        
        'Calling routine must be the skip and skiprpt flags...
        If (m.dUnits <> 0) And (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipRpt, lIndex) = 0) And ((bUseSkipEqFilter = False) Or (gdGetNum(m.hSkipEqFilter, lIndex) = 0)) Then
            If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                AddToTotals lIndex
                CheckConsec
                CheckLargest
                CheckMaxDrawPeaksInTrades lIndex
                m.lTotalTrades = m.lTotalTrades + 1
            Else
                m.dAccountBalance = m.dAccountBalance + (m.dUnits * gdGetNum(m.hProfit, lIndex))
                m.dEquity = m.dEquity + (m.dUnits * gdGetNum(m.hProfit, lIndex))
                gdSetNum m.hTotalProfit, lIndex, m.dUnits * gdGetNum(m.hProfit, lIndex)
                gdSetNum m.hAccountBalance, lIndex, m.dAccountBalance
                gdSetNum m.hEquity, lIndex, m.dEquity
                CheckDrawdown False, lIndex
                
                If m.dAccountBalance > dMaxAccount Then
                    dMaxAccount = m.dAccountBalance
                Else
                    If m.dAccountBalance > 0 Then
                        If dMaxAccount <> 0 Then
                            dDrawPct = (dMaxAccount - m.dAccountBalance) / dMaxAccount
                        Else
                            dDrawPct = 0#
                        End If
                        If dDrawPct > m.dMaxDrawdownPct Then m.dMaxDrawdownPct = dDrawPct
                    End If
                End If
            End If
        Else
            gdSetNum m.hTotalProfit, lIndex, 0
            gdSetNum m.hAccountBalance, lIndex, m.dAccountBalance
            gdSetNum m.hEquity, lIndex, m.dEquity
        End If
        
        'Print #1, lIndex & vbTab & gdGetNum(m.hAccountBalance, lIndex) & vbTab; .Skip(lIndex)
    Next lIndex2
    
#If 0 Then
    ' If we have an open equity trade, then add the open equity to the equity array...
    If lLast = m.Trades.NumRecords - 3 Then
        If gdGetNum(m.hEquity, m.Trades.NumRecords - 1) = -999999 Then
            m.dEquity = m.dEquity + (m.dUnits * m.dOpenEquity)
            gdSetNum m.hEquity, m.Trades.NumRecords - 2, m.dEquity
            gdSetNum m.hEquity, m.Trades.NumRecords - 1, m.dEquity
        End If
    End If
#End If

    For lIndex = 0 To alOpen.Size - 1
        gdSetNum m.hSkip, alOpen(lIndex), 0#
        gdSetNum m.hSkipRpt, alOpen(lIndex), 0#
    Next lIndex
    
    'Close #1

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.ProcessTradeResults", eGDRaiseError_Raise, frmReports.AppPath

End Sub

Private Sub AddToTotals(ByVal lIndex As Long)
On Error GoTo ErrSection:
    
    m.dUnits = gdGetNum(m.hUnits, lIndex) / m.Trades.ItemHdr(gdGetNum(m.hHeaderIndex, lIndex), enth_DefaultUnits)
    m.dProfit = gdGetNum(m.hProfit, lIndex)
    gdSetNum m.hTotalProfit, lIndex, 0
    gdSetNum m.hAccountBalance, lIndex, m.dAccountBalance
    gdSetNum m.hEquity, lIndex, m.dEquity
    
    If m.dProfit > m.dMinProfit Then
        m.lWins = m.lWins + 1
        m.dWinAmt = m.dWinAmt + (m.dProfit * m.dUnits)
        m.lWinBarsInTrades = m.lWinBarsInTrades + gdGetNum(m.hBarsInTrade, lIndex)
        m.dWinRunUp = m.dWinRunUp + (gdGetNum(m.hMaxProfit, lIndex) * m.dUnits)
        m.dWinRunDown = m.dWinRunDown + (gdGetNum(m.hMaxLoss, lIndex) * m.dUnits)
    Else
        m.lLosses = m.lLosses + 1
        m.dLossAmt = m.dLossAmt + (m.dProfit * m.dUnits)
        m.lLossBarsInTrades = m.lLossBarsInTrades + gdGetNum(m.hBarsInTrade, lIndex)
        m.dLossRunUp = m.dLossRunUp + (gdGetNum(m.hMaxProfit, lIndex) * m.dUnits)
        m.dLossRunDown = m.dLossRunDown + (gdGetNum(m.hMaxLoss, lIndex) * m.dUnits)
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.AddToTotals", eGDRaiseError_Raise, frmReports.AppPath

End Sub

Private Sub CheckConsec()
On Error GoTo ErrSection:

    If m.dProfit > m.dMinProfit Then
        m.lWinConsec = m.lWinConsec + 1
        If m.lWinConsec > m.lWinConsecMax Then
            m.lWinConsecMax = m.lWinConsec
        End If
        If m.lWinConsec = 1 Then
            m.lWinRuns = m.lWinRuns + 1
        End If
        m.lLossConsec = 0
        m.lCurrentRun = m.lWinConsec
    Else
        m.lLossConsec = m.lLossConsec + 1
        If m.lLossConsec > m.lLossConsecMax Then
            m.lLossConsecMax = m.lLossConsec
        End If
        If m.lLossConsec = 1 Then
            m.lLossRuns = m.lLossRuns + 1
        End If
        m.lWinConsec = 0
        m.lCurrentRun = m.lLossConsec * -1
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.CheckConsec", eGDRaiseError_Raise, frmReports.AppPath

End Sub

Private Sub CheckLargest()
On Error GoTo ErrSection:

    'MinProfit check is always at the single unit level...
    If m.dProfit > m.dMinProfit Then
        If (m.dProfit * m.dUnits) > m.dWinLargest Then
            m.dWinLargest = m.dProfit * m.dUnits
        End If
    Else
        If (m.dProfit * m.dUnits) < m.dLossLargest Then
            m.dLossLargest = m.dProfit * m.dUnits
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStatsSummary.CheckLargest", eGDRaiseError_Raise, frmReports.AppPath

End Sub

'Keep track of the drawdown using money management
Private Sub CheckDrawdown(ByVal bInitialize As Boolean, ByVal lIndex As Long)
On Error GoTo ErrSection:
    
    Static sdMaxDrawdownClosed As Double
    Static sdMaxDrawdownIntra As Double
    Static sdPrevDrawdown As Double
    Static sdClosedDrawdown As Double
    Dim dPeakToValleyPct As Double
    Dim dProfit As Double
    
    If bInitialize Then
        sdMaxDrawdownClosed = 0
        sdMaxDrawdownIntra = 0
        sdPrevDrawdown = 0
        sdClosedDrawdown = 0
        CheckDrawdown1Unit True, lIndex
        Exit Sub
    End If
    
    'Check for new equity peak when a loss occurs.
    If m.dProfit < m.dMinProfit Then
        m.bNewEquityPeak = False
    Else
        If m.dAccountBalance > m.dEquityPeak Then
            m.bNewEquityPeak = True
            m.dEquityPeak = m.dAccountBalance
        Else
            m.bNewEquityPeak = False
        End If
    End If
    
    If Not m.bNewEquityPeak Then
        sdPrevDrawdown = sdClosedDrawdown
        sdClosedDrawdown = m.dAccountBalance - m.dEquityPeak
        If sdClosedDrawdown < sdMaxDrawdownClosed Then
            sdMaxDrawdownClosed = sdClosedDrawdown
        End If
        
        If m.dEquityPeak > 0 Then
            dPeakToValleyPct = sdClosedDrawdown / m.dEquityPeak
            If dPeakToValleyPct < m.dPeakToValleyDrawdown Then
                m.dPeakToValleyDrawdown = dPeakToValleyPct
            End If
        End If
    End If
    
    'Check Max Intra-Day drawdown
    If lIndex = 1 Then
        m.dIntraDrawdown = gdGetNum(m.hMaxLoss, lIndex) * m.dUnits
    Else
        m.dIntraDrawdown = sdPrevDrawdown + (gdGetNum(m.hMaxLoss, lIndex) * m.dUnits)
    End If
    If m.dIntraDrawdown < sdMaxDrawdownIntra Then
        sdMaxDrawdownIntra = m.dIntraDrawdown
    End If
    
    CheckDrawdown1Unit False, lIndex
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.CheckDrawdown", eGDRaiseError_Raise, frmReports.AppPath

End Sub

'Keep track of the drawdown (always for 1 unit)
Private Sub CheckDrawdown1Unit(ByVal bInitialize As Boolean, ByVal lIndex As Long)
On Error GoTo ErrSection:
    
    Static sdEquityPeak As Double
    Static sdPrevDrawdown As Double
    Static sdClosedDrawdown As Double
    Static sdNetProfit As Double
    Static sdLargestMargin As Double
    
    Dim dProfit As Double
    Dim dMargin As Double
    Dim dAccountReq As Double
    Dim lEntryIndex As Long
    Dim lNumShares As Long
    
    If bInitialize Then
        sdEquityPeak = 0
        sdPrevDrawdown = 0
        sdClosedDrawdown = 0
        sdNetProfit = 0
        sdLargestMargin = 0
        Exit Sub
    End If
    
    ' TLB 3/5/2011: for strategy baskets, we should keep track of at least
    ' the largest margin being traded for purpose of max acct required
    ' (Note: it's hard to know if we can assume multiple margins or not)
    lNumShares = m.Trades.Num(lIndex, entd_NumShares)
    If lNumShares > 0 Then
        ' for stocks, "margin" is simply the cost to enter this trade
        lEntryIndex = m.Trades.Num(lIndex, entd_EntryExitPtr)
        dMargin = m.Trades.Num(lEntryIndex, entd_Price) * lNumShares
    Else
        ' for futures, get margin from header
        dMargin = m.Trades.NumHdr(gdGetNum(m.hHeaderIndex, lIndex), enth_Margin)
    End If
    If dMargin > sdLargestMargin Then
        sdLargestMargin = dMargin
    Else
        dMargin = sdLargestMargin
    End If
    
    dProfit = gdGetNum(m.hProfit, lIndex)
    sdNetProfit = sdNetProfit + (dProfit * m.dUnits)
    
    'Check for new equity peak when loss occurs
    If dProfit < m.dMinProfit Then
        m.bNewEquityPeak = False
    Else
        If sdNetProfit > sdEquityPeak Then
            m.bNewEquityPeak = True
            sdEquityPeak = sdNetProfit
        Else
            m.bNewEquityPeak = False
        End If
    End If
    
    If Not m.bNewEquityPeak Then
    
        sdPrevDrawdown = sdClosedDrawdown
        sdClosedDrawdown = sdNetProfit - sdEquityPeak
        
        If sdClosedDrawdown < m.dMaxDrawdownClosed Then
            m.dMaxDrawdownClosed = sdClosedDrawdown
            m.lMaxDrawdownClosedLoc = lIndex
        End If
        
    End If
    
    'Check Max Intra-Day drawdown
    If lIndex = 1 Then
        m.dIntraDrawdown = gdGetNum(m.hMaxLoss, lIndex) * m.dUnits
    Else
        m.dIntraDrawdown = sdPrevDrawdown + (gdGetNum(m.hMaxLoss, lIndex) * m.dUnits)
    End If
    If m.dIntraDrawdown < m.dMaxDrawdownIntra Then
        m.dMaxDrawdownIntra = m.dIntraDrawdown
    End If
    
    ' Account Required = Intra Trade Drawdown + Margin
    dAccountReq = Abs(m.dIntraDrawdown) + (dMargin * m.dUnits)
    If dAccountReq > m.dMaxAccountReq Then
        m.dMaxAccountReq = dAccountReq
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStatsSummary.CheckDrawdown1Unit", eGDRaiseError_Raise, frmReports.AppPath

End Sub

' Keep track of maximum dips/peaks in winning and losing trades
Private Sub CheckMaxDrawPeaksInTrades(ByVal lIndex As Long)
On Error GoTo ErrSection:

    'Keep track of largest drawdown during winners
    If gdGetNum(m.hProfit, lIndex) > m.dMinProfit Then
        If (gdGetNum(m.hMaxLoss, lIndex) * m.dUnits) < m.dMaxDrawdownInWinner Then
            m.dMaxDrawdownInWinner = gdGetNum(m.hMaxLoss, lIndex) * m.dUnits
        End If
        m.dTotDrawdownInWinners = m.dTotDrawdownInWinners + (gdGetNum(m.hMaxLoss, lIndex) * m.dUnits)
    Else
        If (gdGetNum(m.hMaxProfit, lIndex) * m.dUnits) > m.dMaxPeakInLoser Then
            m.dMaxPeakInLoser = gdGetNum(m.hMaxProfit, lIndex) * m.dUnits
        End If
        m.dTotPeakInLosers = m.dTotPeakInLosers + (gdGetNum(m.hMaxProfit, lIndex) * m.dUnits)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStatsSummary.CheckMaxDrawPeaksInTrades", eGDRaiseError_Raise, frmReports.AppPath

End Sub

' TLB 2/24/2011: this routine was not working very well (esp. for Strategy Baskets),
' so rewritten (new routine below) -- original code is now obsolete ...
'
'This routine creates a table of end of month balances. The table has the following fields
' - Date:   Date of latest trade during month
' - AccBal: The account balance as of this date
' - TradesOccured: 1=If at least one trade was exited during mth, 0=no activity
Public Sub CalcEndOfMthBalances_OLD()
On Error GoTo ErrSection:
    
    Dim TradeActivityFound  As Byte
    Dim X As Long
    Dim R As Long
    Dim RetVal As Long
    Dim curMonth As Byte
    Dim NextMonth As Byte
    Dim n As Integer
    Dim aTradeDate As Long
    Dim aAccountBalance As Long
    Dim aSignalType As Long
    Dim aSkip As Long
    Dim aSkipRpt As Long
    
    'Save global Field addresses from Trades class
    With m.Trades
        aTradeDate = .FieldHandle(entd_TradeDate)
        aAccountBalance = .FieldHandle(entd_AccountBalance)
        aSignalType = .FieldHandle(entd_SignalType)
        aSkip = .FieldHandle(entd_Skip)
        aSkipRpt = .FieldHandle(entd_SkipRpt)
    End With
    
    'Grab the end of month balances only and build a gdArray (m.tblRetMths)
    Set m.tblRetMths = New cGdTable
    RetVal = m.tblRetMths.CreateField(eGDARRAY_Doubles, mMR_TradeDate, "Date")
    RetVal = m.tblRetMths.CreateField(eGDARRAY_Doubles, mMR_AccBal, "AccBal")
    RetVal = m.tblRetMths.CreateField(eGDARRAY_Longs, mMR_NumTrades, "TradesOccured")
    RetVal = m.tblRetMths.CreateField(eGDARRAY_Floats, mMR_ReturnPct, "ReturnPct")
    m.tblRetMths.Num(mMR_TradeDate, 0) = gdGetNum(aTradeDate, 1)
    m.tblRetMths.Num(mMR_AccBal, 0) = gdGetNum(aAccountBalance, 0)
    
    
    '---------------------------------------------------------------
    ' Build a gdTable of end of month account balances.  Take the
    ' last exit signals to occur during the month and save the account
    ' balance as of that trade
    Dim Mths As Integer
    Dim curYear As Long
    Dim Y As Long
    Dim NextTrade As Long
    Dim AccBal As Double
    Dim MonthDate As Double
    Dim PriorTo1st As Boolean
    
    Mths = DateDiff("m", gdGetNum(aTradeDate, 1), gdGetNum(aTradeDate, m.Trades.NumRecords - 1))
    curMonth = 1
    curYear = Year(gdGetNum(aTradeDate, 1))
    NextTrade = 1
    PriorTo1st = True
    
    'Build array of ending account balance values
    For X = 1 To Mths
        
        'Calculate the current month/Year
        If curMonth = 13 Then
            curMonth = 1
            curYear = curYear + 1
        End If
        TradeActivityFound = 0
        MonthDate = CDate(curMonth & "/01/" & curYear)
        AccBal = gdGetNum(aAccountBalance, NextTrade - 1)
        
        'Look for and analyze trades in current month/year
        For Y = NextTrade To m.Trades.NumRecords - 1
            If gdGetNum(aSignalType, Y) = 1 And gdGetNum(aSkip, Y) = 0 And _
                gdGetNum(aSkipRpt, Y) = 0 Then 'Look at exits only
                If curMonth = Month(gdGetNum(aTradeDate, Y)) And _
                   curYear = Year(gdGetNum(aTradeDate, Y)) Then
                    AccBal = gdGetNum(aAccountBalance, Y)
                    MonthDate = gdGetNum(aTradeDate, Y)
                    TradeActivityFound = 1
                    PriorTo1st = False
                Else
                    Exit For    'Month changed...
                End If
            End If
        Next Y
        If Not PriorTo1st Then
            NextTrade = Y
        End If
        
        R = R + 1
        m.tblRetMths.Num(mMR_TradeDate, R) = MonthDate
        m.tblRetMths.Num(mMR_AccBal, R) = AccBal
        m.tblRetMths.Num(mMR_NumTrades, R) = TradeActivityFound
        
        curMonth = curMonth + 1
    Next X
   
'** Was not working for Scaling/pyramiding type systems.
If 0 Then
    'Build array of ending account balance values
    For X = 1 To m.Trades.NumRecords - 2
        curMonth = Month(gdGetNum(aTradeDate, X))
        NextMonth = Month(gdGetNum(aTradeDate, X + 1))
        
        'Check to see if an Exit trade occurred.  If so, then there
        'was activity in this period
        If gdGetNum(aSignalType, X) = 1 Then
            TradeActivityFound = 1
        End If
        
        'Month has changed
        If curMonth <> NextMonth Then
            If curMonth = 12 Then
                R = R + 1
                m.tblRetMths.Num(mMR_TradeDate, R) = gdGetNum(aTradeDate, X)
                m.tblRetMths.Num(mMR_AccBal, R) = gdGetNum(aAccountBalance, X)
                m.tblRetMths.Num(mMR_NumTrades, R) = TradeActivityFound
                TradeActivityFound = 0
                curMonth = 1
            End If
            If curMonth < NextMonth Then
                n = 0
                Do Until curMonth = NextMonth
                    n = n + 1
                    R = R + 1
                    m.tblRetMths.Num(mMR_TradeDate, R) = DateAdd("m", n, gdGetNum(aTradeDate, X))
                    m.tblRetMths.Num(mMR_AccBal, R) = gdGetNum(aAccountBalance, X)
                    m.tblRetMths.Num(mMR_NumTrades, R) = TradeActivityFound
                    TradeActivityFound = 0
                    curMonth = curMonth + 1
                Loop
                
            End If
        End If
    Next X
End If

    'Calculate return pct for each month
    For X = 1 To m.tblRetMths.NumRecords - 1
        If m.tblRetMths.Num(mMR_AccBal, X - 1) <> 0 Then
            m.tblRetMths.Num(mMR_ReturnPct, X) = (m.tblRetMths.Num(mMR_AccBal, X) - m.tblRetMths.Num(mMR_AccBal, X - 1)) / m.tblRetMths.Num(mMR_AccBal, X - 1)
        Else
            m.tblRetMths.Num(mMR_AccBal, X) = 0
        End If
    Next X
        
    'Calculate standard statistics on Monthly returns array
    gdCalcStatistics m.tblRetMths.FieldArrayHandle(mMR_ReturnPct), _
        m.MthRetResults, 0, 0, -1
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.CalcEndOfMthBalances_OLD", eGDRaiseError_Raise, frmReports.AppPath

End Sub

' TLB 2/24/2011: with this new calculation method, the trade sequence doesn't matter
' (which may be why the original code was not working properly for strategy baskets)
'
'This routine creates a table of end of month balances. The table has the following fields
' - Date:   Date of latest trade during month
' - AccBal: The account balance as of this date
' - NumTrades: 0=no activity
' - Profit: profit for month
Public Sub CalcEndOfMthBalances()
On Error GoTo ErrSection:
       
    Dim i&, nDate&, nMonth&, nFirstMonth&
    Dim dProfit#, dBalance#, dPrevBalance#
    Dim aTradeDate As Long
    Dim aProfit As Long
    Dim aSignalType As Long
    Dim aSkip As Long
    Dim aSkipRpt As Long
    Dim aRuleID As Long
    
If 1 Then
    CalcEndOfPeriodBalances
    Exit Sub
End If
    
    'Save global Field addresses from Trades class
    With m.Trades
        aTradeDate = .FieldHandle(entd_TradeDate)
        aProfit = .FieldHandle(entd_TotalProfit)
        aSignalType = .FieldHandle(entd_SignalType)
        aSkip = .FieldHandle(entd_Skip)
        aSkipRpt = .FieldHandle(entd_SkipRpt)
        aRuleID = .FieldHandle(entd_RuleID)
    End With
    
    'Grab the end of month balances only and build a gdArray (m.tblRetMths)
    Set m.tblRetMths = New cGdTable
    i = m.tblRetMths.CreateField(eGDARRAY_Doubles, mMR_TradeDate, "Date")
    i = m.tblRetMths.CreateField(eGDARRAY_Doubles, mMR_AccBal, "AccBal")
    i = m.tblRetMths.CreateField(eGDARRAY_Longs, mMR_NumTrades, "NumTrades")
    i = m.tblRetMths.CreateField(eGDARRAY_Floats, mMR_ReturnPct, "ReturnPct")
    i = m.tblRetMths.CreateField(eGDARRAY_Doubles, mMR_Profit, "Profit")
    
    ' get earliest trade date (so know where to start monthly records)
    nDate = gdMinValue(aTradeDate, 0, gdGetSize(aTradeDate))
    nFirstMonth = Year(nDate) * 12 + Month(nDate)
    For i = 0 To m.Trades.NumRecords - 1
        ' only look at the exits for closed trades which are valid (based on filters)
        If gdGetNum(aSignalType, i) = 1 And gdGetNum(aSkip, i) = 0 _
                And gdGetNum(aSkipRpt, i) = 0 And gdGetNum(aRuleID, i) <> 0 Then
            ' get index into the monthly array (nMonth)
            nDate = gdGetNum(aTradeDate, i)
            nMonth = (Year(nDate) * 12 + Month(nDate)) - nFirstMonth
            ' keep track of latest trade date in the month
            If nDate > m.tblRetMths.Num(mMR_TradeDate, nMonth) Then
                If m.tblRetMths.Num(mMR_TradeDate, nMonth) <= 0 Then
                    ' init for a new month
                    m.tblRetMths.Num(mMR_NumTrades, nMonth) = 0
                    m.tblRetMths.Num(mMR_Profit, nMonth) = 0
                End If
                m.tblRetMths.Num(mMR_TradeDate, nMonth) = nDate
            End If
            dProfit = gdGetNum(aProfit, i)
            m.tblRetMths.Num(mMR_Profit, nMonth) = m.tblRetMths.Num(mMR_Profit, nMonth) + dProfit
            m.tblRetMths.Num(mMR_NumTrades, nMonth) = m.tblRetMths.Num(mMR_NumTrades, nMonth) + 1
        End If
    Next
    
    ' remove any empty beginning records
    Do While m.tblRetMths.Num(mMR_NumTrades, 0) <= 0 And m.tblRetMths.NumRecords > 0
        m.tblRetMths.RemoveRecords 0
        nFirstMonth = nFirstMonth + 1
    Loop

    ' fix other items for each month
    dBalance = m.Trades.Num(0, entd_AccountBalance)
    If dBalance <= 0 Then
        dBalance = 100000 ' default starting balance
    End If
    For i = 0 To m.tblRetMths.NumRecords - 1
        ' make sure this month had been initialized
        If m.tblRetMths.Num(mMR_NumTrades, i) <= 0 Then
            nMonth = nFirstMonth + i
            nDate = DateSerial(Int(nMonth / 12), nMonth Mod 12, 1)
            m.tblRetMths.Num(mMR_TradeDate, i) = nDate
            m.tblRetMths.Num(mMR_NumTrades, i) = 0
            m.tblRetMths.Num(mMR_Profit, i) = 0
        End If
        
        dPrevBalance = dBalance
        dProfit = m.tblRetMths.Num(mMR_Profit, i)
        dBalance = dBalance + dProfit
        m.tblRetMths.Num(mMR_AccBal, i) = dBalance
        If dPrevBalance > 0 Then
            m.tblRetMths.Num(mMR_ReturnPct, i) = (dBalance - dPrevBalance) / dPrevBalance
        Else
            m.tblRetMths.Num(mMR_ReturnPct, i) = 0
        End If
    Next
        
    'Calculate standard statistics on Monthly returns array
    gdCalcStatistics m.tblRetMths.FieldArrayHandle(mMR_ReturnPct), m.MthRetResults, 0, 0, -1
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.CalcEndOfMthBalances", eGDRaiseError_Raise, frmReports.AppPath
End Sub

' TLB 2/24/2011: with this new calculation method, the trade sequence doesn't matter
' (which may be why the original code was not working properly for strategy baskets)
'
' This routine creates a table of end of period balances. The table has the following fields
' - Date:  Date of latest trade during period (month, year, day, etc)
' - AccBal: The account balance as of this date
' - NumTrades: 0=no activity
' - Profit: profit for period
Public Sub CalcEndOfPeriodBalances()
On Error GoTo ErrSection:
       
    Dim iTrade&, iRec&, iRecOffset&, nDate&
    Dim dProfit#, dBalance#, dPrevBalance#
    Dim aTradeDate&, aProfit&, aSignalType&, aSkip&, aSkipRpt&, aRuleID&
    
    Dim iMode% ' 0=year, 1=qtr, 2=month, 3=week, 4=day
    iMode = 2
    
    'Save global Field addresses from Trades class
    With m.Trades
        aTradeDate = .FieldHandle(entd_TradeDate)
        aProfit = .FieldHandle(entd_TotalProfit)
        aSignalType = .FieldHandle(entd_SignalType)
        aSkip = .FieldHandle(entd_Skip)
        aSkipRpt = .FieldHandle(entd_SkipRpt)
        aRuleID = .FieldHandle(entd_RuleID)
    End With
    
    ' build a table (m.tblRetMths) with end-of-period balances
    Set m.tblRetMths = New cGdTable
    m.tblRetMths.CreateField eGDARRAY_Longs, mMR_TradeDate, "Date"
    m.tblRetMths.CreateField eGDARRAY_Doubles, mMR_AccBal, "AccBal"
    m.tblRetMths.CreateField eGDARRAY_Longs, mMR_NumTrades, "NumTrades"
    m.tblRetMths.CreateField eGDARRAY_Floats, mMR_ReturnPct, "ReturnPct"
    m.tblRetMths.CreateField eGDARRAY_Doubles, mMR_Profit, "Profit"
    
    ' get earliest trade date (so know where to start monthly records)
    nDate = Int(gdMinValue(aTradeDate, 0, gdGetSize(aTradeDate)))
    Select Case iMode
    Case 0 ' Yearly
        iRecOffset = Year(nDate)
    Case 1 ' Qtrly
        iRecOffset = Year(nDate) * 4 + Int((Month(nDate) - 1) / 3)
    Case 2 ' Monthly
        iRecOffset = Year(nDate) * 12 + Month(nDate) - 1
    Case 3 ' Weekly
        iRecOffset = Int((nDate - 1) / 7)
    Case 4 ' Daily
        If Weekday(nDate) = vbSunday Then
            nDate = nDate + 1
        ElseIf Weekday(nDate) = vbSaturday Then
            nDate = nDate - 1
        End If
        iRecOffset = nDate
    End Select
    
    For iTrade = 0 To m.Trades.NumRecords - 1
        ' only look at the exits for closed trades which are valid (based on filters)
        If gdGetNum(aSignalType, iTrade) = 1 And gdGetNum(aSkip, iTrade) = 0 _
                And gdGetNum(aSkipRpt, iTrade) = 0 And gdGetNum(aRuleID, iTrade) <> 0 Then
            
            ' get index into the period array (iRec)
            nDate = Int(gdGetNum(aTradeDate, iTrade))
            Select Case iMode
            Case 0 ' Yearly
                iRec = Year(nDate)
            Case 1 ' Qtrly
                iRec = Year(nDate) * 4 + Int((Month(nDate) - 1) / 3)
            Case 2 ' Monthly
                iRec = Year(nDate) * 12 + Month(nDate) - 1
            Case 3 ' Weekly
                iRec = Int((nDate - 1) / 7)
            Case 4 ' Daily
                If Weekday(nDate) = vbSunday Then
                    nDate = nDate + 1
                ElseIf Weekday(nDate) = vbSaturday Then
                    nDate = nDate - 1
                End If
                iRec = nDate
            End Select
            iRec = iRec - iRecOffset
            
            ' keep track of latest trade date in the period
            If nDate > m.tblRetMths.Num(mMR_TradeDate, iRec) Then
                If m.tblRetMths.Num(mMR_TradeDate, iRec) <= 0 Then
                    ' init for a new period
                    m.tblRetMths.Num(mMR_NumTrades, iRec) = 0
                    m.tblRetMths.Num(mMR_Profit, iRec) = 0
                End If
                m.tblRetMths.Num(mMR_TradeDate, iRec) = nDate
            End If
            dProfit = gdGetNum(aProfit, iTrade)
            m.tblRetMths.Num(mMR_Profit, iRec) = m.tblRetMths.Num(mMR_Profit, iRec) + dProfit
            m.tblRetMths.Num(mMR_NumTrades, iRec) = m.tblRetMths.Num(mMR_NumTrades, iRec) + 1
        End If
    Next
    
    ' remove any empty beginning records
    Do While m.tblRetMths.Num(mMR_NumTrades, 0) <= 0 And m.tblRetMths.NumRecords > 0
        m.tblRetMths.RemoveRecords 0
        iRecOffset = iRecOffset + 1
    Loop

    ' fix other items for each period
    dBalance = m.Trades.Num(0, entd_AccountBalance)
    If dBalance <= 0 Then
        dBalance = 100000 ' default starting balance
    End If
    For iRec = 0 To m.tblRetMths.NumRecords - 1
        ' make sure this period had been initialized
        If m.tblRetMths.Num(mMR_NumTrades, iRec) <= 0 Then
            ' convert to first date of period
            nDate = iRecOffset + iRec
            Select Case iMode
            Case 0 ' Yearly
                nDate = DateSerial(nDate, 1, 1)
            Case 1 ' Qtrly
                nDate = DateSerial(Int(nDate / 4), (nDate Mod 4) * 3 + 1, 1)
            Case 2 ' Monthly
                nDate = DateSerial(Int(nDate / 12), (nDate Mod 12) + 1, 1)
            Case 3 ' Weekly
                nDate = nDate * 7 + 2 '= Monday
            Case 4 ' Daily
                If Not IsWeekday(nDate) Then
                    nDate = -1 ' to skip over weekends
                End If
            End Select
            If nDate >= 0 Then
                m.tblRetMths.Num(mMR_TradeDate, iRec) = nDate
                m.tblRetMths.Num(mMR_NumTrades, iRec) = 0
                m.tblRetMths.Num(mMR_Profit, iRec) = 0
            End If
        Else
            nDate = 0 ' so won't skip over it
        End If
        
        If nDate >= 0 Then
            dPrevBalance = dBalance
            dProfit = m.tblRetMths.Num(mMR_Profit, iRec)
            dBalance = dBalance + dProfit
            m.tblRetMths.Num(mMR_AccBal, iRec) = dBalance
            If dPrevBalance > 0 Then
                m.tblRetMths.Num(mMR_ReturnPct, iRec) = (dBalance - dPrevBalance) / dPrevBalance
            Else
                m.tblRetMths.Num(mMR_ReturnPct, iRec) = 0
            End If
        End If
    Next
        
    'Calculate standard statistics on Monthly returns array
    gdCalcStatistics m.tblRetMths.FieldArrayHandle(mMR_ReturnPct), m.MthRetResults, 0, 0, -1
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStatsSummary.CalcEndOfPeriodBalances", eGDRaiseError_Raise, frmReports.AppPath
End Sub


Private Function CalculateTWR(ByVal F As Double) As Double
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lTrade As Long                  ' Index into a for loop
    Dim TWR As Double                   ' TWR value to return
    Dim lLast As Long                   ' Last trade to process (ignore open positions)
    
    'If m.Trades.Item(m.Trades.NumRecords - 1, entd_RuleID) = 0 Then
    '    lLast = m.Trades.NumRecords - 3
    'Else
        lLast = m.Trades.NumRecords - 1
    'End If
    
    TWR = 1#
    For lIndex = 1 To lLast
        lTrade = gdGetNum(m.hSorted, lIndex)
        
        If gdGetNum(m.hSignalType, lTrade) = gExitSignal Then
            If gdGetNum(m.hSkip, lTrade) = 0 And gdGetNum(m.hSkipRpt, lTrade) = 0 Then
                TWR = TWR * (1 + (F * gdGetNum(m.hProfit, lTrade) / Abs(m.dLossLargest)))
            End If
        End If
    Next lIndex
    
    CalculateTWR = TWR

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStatsSummary.CalculateTWR", eGDRaiseError_Raise, frmReports.AppPath
    
End Function

Private Function CalculateOptimalF() As Double
On Error GoTo ErrSection:

    Dim F As Double                     ' Optimal f value that we calculated
    Dim TWR As Double                   ' Calculated TWR value
    Dim dAbove As Double                ' Temporary TWR value
    Dim dBelow As Double                ' Temporary TWR value
    
    If m.lLosses = 0& Or m.dLossLargest = 0# Or m.lWins = 0& Or m.dWinLargest = 0# Then
        CalculateOptimalF = 0#
    Else
        F = Round(KellyRatio, 2)
        TWR = CalculateTWR(F)
        dAbove = CalculateTWR(F + 0.01)
        dBelow = CalculateTWR(F - 0.01)
        
        If dAbove > TWR Then
            Do
                F = F + 0.01
                If F = 1 Then Exit Do
                TWR = CalculateTWR(F)
                dAbove = CalculateTWR(F + 0.01)
            Loop While dAbove > TWR
        ElseIf dBelow > TWR Then
            Do
                F = F - 0.01
                If F = 0 Then Exit Do
                TWR = CalculateTWR(F)
                dBelow = CalculateTWR(F - 0.01)
            Loop While dBelow > TWR
        End If
        
        CalculateOptimalF = F
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStatsSummary.CalculateOptimalF", eGDRaiseError_Raise, frmReports.AppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EquityMovingAverage
'' Description: Calculate and store the moving average of the equity
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EquityMovingAverage()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim dRollingSum As Double           ' Rolling sum value
    Dim dValue As Double                ' Value to store into the array
    Dim lNum As Long                    ' Number of periods for moving average
    Dim dPct As Double                  ' Percentage for a exponential MA
    Dim lCounter As Long                ' Counter variable
    Dim adEquity As New cGdArray        ' Rolling queue of equities
    Dim dPrevMa As Double               ' Previous moving average value
    Dim lExit As Long                   ' Index for the exit in the table
    Dim dFilteredEquity As Double       ' Filtered equity
    
    lCounter = 0&
    dRollingSum = 0#
    dFilteredEquity = 0#
    lNum = m.lMovAvgPeriod
    dPct = 2# / (lNum + 1)
    adEquity.Create eGDARRAY_Doubles, lNum + 1
    m.nTakeNextTrade = eGDTakeNextTrade_NotEnoughData
    
    For lIndex2 = 1 To m.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipRpt, lIndex) = 0) Then
            If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                lExit = gdGetNum(m.hEntryExitPtr, lIndex)
                
                gdSetNum m.hEquityMA, lIndex, dValue
                gdSetNum m.hFilteredEquity, lIndex, dFilteredEquity
                gdSetNum m.hUnfilteredEquity, lIndex, gdGetNum(m.hEquity, lIndex)
                
                If m.nTakeNextTrade = eGDTakeNextTrade_No Then
                    gdSetNum m.hSkipEqFilter, lIndex, 1
                    gdSetNum m.hSkipEqFilter, lExit, 1
                Else
                    gdSetNum m.hSkipEqFilter, lIndex, 0
                    gdSetNum m.hSkipEqFilter, lExit, 0
                End If
            Else
                lCounter = lCounter + 1
                
                If (gdGetNum(m.hRuleID, lIndex) <> 0) Then
                    dPrevMa = dValue
                    If UCase(m.strMAType) = "SIMPLE" Then
                        dRollingSum = dRollingSum + gdGetNum(m.hEquity, lIndex)
                        If lCounter > lNum Then
                            dRollingSum = dRollingSum - adEquity(1)
                            adEquity.Remove 1
                            adEquity(lNum) = gdGetNum(m.hEquity, lIndex)
                        Else
                            adEquity(lCounter) = gdGetNum(m.hEquity, lIndex)
                        End If
                        
                        If lCounter >= lNum Then
                            dValue = dRollingSum / CDbl(lNum)
                        Else
                            dValue = -999999#
                        End If
                    Else
                        dValue = (gdGetNum(m.hEquity, lIndex) * dPct) + (dValue * (1 - dPct))
                    End If
                            
                    If (gdGetNum(m.hSkipEqFilter, lIndex) = 0) Then
                        dFilteredEquity = dFilteredEquity + (m.dUnits * gdGetNum(m.hProfit, lIndex))
                    End If
                End If
                
                If m.bEquityFilterOn Then
                    If dValue = -999999# Then
                        m.nTakeNextTrade = eGDTakeNextTrade_NotEnoughData
                    Else
                        Select Case m.nEquityFilterMode
                            Case eGDEquityFilterMode_BelowMa
                                If (gdGetNum(m.hEquity, lIndex) >= dValue) Then
                                    m.nTakeNextTrade = eGDTakeNextTrade_Yes
                                Else
                                    m.nTakeNextTrade = eGDTakeNextTrade_No
                                End If
                            
                            Case eGDEquityFilterMode_MaDown
                                If (dValue >= dPrevMa) Then
                                    m.nTakeNextTrade = eGDTakeNextTrade_Yes
                                Else
                                    m.nTakeNextTrade = eGDTakeNextTrade_No
                                End If
                                
                        End Select
                    End If
                Else
                    m.nTakeNextTrade = eGDTakeNextTrade_NoEquityFilter
                End If
            
                gdSetNum m.hEquityMA, lIndex, dValue
                gdSetNum m.hFilteredEquity, lIndex, dFilteredEquity
                gdSetNum m.hUnfilteredEquity, lIndex, gdGetNum(m.hEquity, lIndex)
            End If
        End If
    Next lIndex2
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStatsSummary.EquityMovingAverage", eGDRaiseError_Raise, frmReports.AppPath
    
End Sub

