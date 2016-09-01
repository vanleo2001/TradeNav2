VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Type mPrivate
    ' Property used to return calculated summary stats
    cBeginBalance As Currency
    cTradeToTradeDrawdown As Currency
    aTrades As cGdArray
    bUseGdTrades As Boolean
    cMinProfit As Currency
    cMargin As Currency

    ' Work fields for calculating statistics
    cAccountBalance As Currency
    cMaxDrawdownClosed As Currency
    cMaxDrawdownIntra As Currency
    cEndingBalance As Currency
    cEquityPeak As Currency
    cLossLargest As Currency
    lLossConsec As Long
    lLossRuns As Long
    lLossConsecMax As Long
    lLosses As Long
    cLossAmt As Currency
    dPeakToValleyDrawdown As Double
    cWinLargest As Currency
    lWinConsec As Long
    lWinConsecMax As Long
    lWins As Long
    lWinRuns As Long
    cWinAmt As Currency

    'Used for calculating average yearly return
    cLastBalance As Currency
    dYrReturnTotal As Double
    lSaveYear As Long
    lYears As Long

    dDDPct As Double
    cClosedDrawdown As Currency
    cIntraDrawdown As Currency
    cPrevDrawdown As Currency
    bNewEquityPeak As Boolean
    cNetProfit As Currency
    cProfit As Currency

    X As Long
    lTotalTrades As Long
    lWinBarsInTrades As Long
    lLossBarsInTrades As Long
    dUnits As Double
    lBarsInTrade As Long
    cOpenPosLoss As Currency
    bSkip As Boolean
    dExitDate As Date
    lTotalBars As Long
    
    dStartDate As Double
    dEndDate As Double
    dYearsInSystem As Double
End Type
Private m As mPrivate

Property Let BeginBalance(ByVal pData As Currency)
    m.cBeginBalance = pData
End Property
Property Let Margin(ByVal pData As Currency)
    m.cMargin = pData
End Property
Property Let MinProfit(pData As Currency)
    m.cMinProfit = pData
End Property

'Overall stats
Property Get EndingBalance() As Currency
    EndingBalance = m.cAccountBalance
End Property
Property Get TotTrades() As Long
    TotTrades = m.lTotalTrades
End Property
Property Get MaxDrawdownClosed() As Currency
    MaxDrawdownClosed = m.cMaxDrawdownClosed
End Property
Property Get MaxDrawdownIntra() As Currency
    MaxDrawdownIntra = m.cMaxDrawdownIntra
End Property
Property Get PeakToValleyDrawdown() As Double
    PeakToValleyDrawdown = Abs(m.dPeakToValleyDrawdown)
End Property
Property Get AvgYearlyReturn() As Double
    If m.dYearsInSystem <> 0 Then
        AvgYearlyReturn = m.dYrReturnTotal / m.dYearsInSystem
    Else
        AvgYearlyReturn = 0
    End If
End Property
Property Get EquityPeak() As Currency
    EquityPeak = m.cEquityPeak
End Property
Property Get AccountSizeReq() As Currency
    AccountSizeReq = Abs(m.cMaxDrawdownIntra - m.cMargin)
End Property

'f = ((PLRatio+1)*WinPct-1)/PLRatio
Property Get OptimalF() As Double
    If PLRatio > 0 Then
        OptimalF = ((PLRatio + 1) * WinPct - 1) / PLRatio
    Else
        OptimalF = 0
    End If
End Property

' >=  2 Winners tend to follow losers, losers follow winners
' <= -2 Winners tend to follow winners, losers follow losers
Property Get ZScore() As Double
On Error GoTo ErrSection:

    Dim Divisor As Double
    Dim n       As Double
    Dim X       As Double
    Dim r       As Double
    Dim xWins   As Double
    Dim xLosses As Double
    
    'X = Wins * Losses * 2      'overflow (Wins and Losses are integer values)
    xWins = Wins
    xLosses = Losses
    X = xWins * xLosses * 2
    n = TotTrades
    
    ZScore = 0
    Divisor = (((X * (X - n)) / (n - 1)) ^ (1 / 2))
    If Divisor <> 0 Then
        ZScore = (n * ((WinRuns + LossRuns) - 0.5) - X) / Divisor
    Else
        ZScore = 0
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cStats.ZScore.Get", eGDRaiseError_Raise
    
End Property

Property Get PLRatio() As Double
On Error GoTo ErrSection:

    Dim LA As Double
    
    LA = LossAvg
    If LA <> 0 Then
        PLRatio = Abs(WinAvg / LA)
    Else
        PLRatio = WinAvg
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cStats.PLRatio.Get", eGDRaiseError_Raise
    
End Property

Property Get ReturnPct() As Double
On Error GoTo ErrSection:

    Dim MaxLoss     As Single
    
''    If m.cBeginBalance <> 0 Then
''        ReturnPct = (m.cAccountBalance - m.cBeginBalance) / m.cBeginBalance
''    Else
        'If Max High/Low exist on trade records then calculate max intraday
        'drawdown.  Use it to calculated Return Pct.  This is the default mode
        'for system tests (no money management).
        If Abs(m.cMaxDrawdownIntra) + m.cMargin <> 0 Then
            ReturnPct = (m.cWinAmt + m.cLossAmt) / (Abs(m.cMaxDrawdownIntra) + m.cMargin)
        Else
            'Use drawdown closed if no MaxIntradrawdown specified
            If Abs(m.cMaxDrawdownClosed) + m.cMargin <> 0 Then
                ReturnPct = (m.cWinAmt + m.cLossAmt) / (Abs(m.cMaxDrawdownClosed) + m.cMargin)
            Else
                ReturnPct = 0
            End If
        End If
''    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cStats.ReturnPct.Get", eGDRaiseError_Raise
    
End Property

Property Get TotalNetProfit() As Currency
    TotalNetProfit = m.cWinAmt + m.cLossAmt
End Property

Property Get KellyRatio() As Double
    If PLRatio <> 0 Then
        KellyRatio = ((PLRatio + 1) * WinPct - 1) / PLRatio
    Else
        KellyRatio = 0
    End If
End Property

Property Get AvgTrade() As Currency
    If m.lTotalTrades <> 0 Then
        AvgTrade = (m.cWinAmt + m.cLossAmt) / m.lTotalTrades
    Else
        AvgTrade = 0
    End If
End Property

Property Get ProfitFactor() As Double
    If m.cWinAmt = 0 Then
        ProfitFactor = 0
    ElseIf m.cLossAmt <> 0 Then
        ProfitFactor = Abs(m.cWinAmt / m.cLossAmt)
    Else
        ProfitFactor = 999
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
        WinAvg = m.cWinAmt / m.lWins
    Else
        WinAvg = 0
    End If
End Property

Property Get WinLargest() As Currency
    WinLargest = m.cWinLargest
End Property
Property Get WinTotal() As Currency
    WinTotal = m.cWinAmt
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
        LossAvg = m.cLossAmt / m.lLosses
    Else
        LossAvg = 0
    End If
End Property

Property Get LossLargest() As Currency
    LossLargest = m.cLossLargest
End Property
Property Get LossTotal() As Currency
    LossTotal = m.cLossAmt
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

Public Sub gdCalculate(pgdTrades As cGdArray)
On Error GoTo ErrSection:
    
    Dim bStocks As Boolean
    Dim dEntryCost As Double
    Dim aFields As cGdArray
    
    Set aFields = New cGdArray
    
    m.cNetProfit = 0
    m.cEquityPeak = m.cBeginBalance
    m.dPeakToValleyDrawdown = 0
    m.cClosedDrawdown = 0
    m.cIntraDrawdown = 0
    m.cMaxDrawdownClosed = 0
    m.cMaxDrawdownIntra = 0
    
    'Used to calculate Avg year return
    m.cLastBalance = m.cBeginBalance
    m.dYrReturnTotal = 0
    
    'Extract header info
    aFields.SplitFields pgdTrades(0), vbTab
    m.lSaveYear = Year(CVDate(Val(aFields(3))))
    m.lTotalBars = Val(aFields(5))
    m.cMargin = Val(aFields(11))
    m.dStartDate = ValOfText(aFields(3))
    m.dEndDate = ValOfText(aFields(4))
    m.dYearsInSystem = (m.dEndDate - m.dStartDate) / 365.25
    
    m.lYears = 0
    
    m.lTotalTrades = 0
    m.lWins = 0: m.cWinAmt = 0: m.lLosses = 0: m.cLossAmt = 0
    m.lWinRuns = 0: m.lLossRuns = 0
    m.cWinLargest = 0: m.lWinConsecMax = 0: m.lWinConsec = 0
    m.cLossLargest = 0: m.lLossConsecMax = 0: m.lLossConsec = 0
    m.cAccountBalance = m.cBeginBalance
    m.lWinBarsInTrades = 0
    m.lLossBarsInTrades = 0
    
    ' TLB 7/25/2013: use different method for stocks
    If UCase(aFields(12)) = "S" Then
        bStocks = True
        m.cMargin = 0
    End If
    
    'Initialize Statics variables (for yearly runs with drawdown)
    CheckDrawdown True

    'Skip header record (record 0)
    For m.X = 1 To pgdTrades.Size - 1
    
        'Extract trade info
        aFields.SplitFields pgdTrades(m.X), vbTab
        
        m.cProfit = Val(aFields(7))
        m.dUnits = 1
        m.dExitDate = CVDate(Val(aFields(4)))
        m.cOpenPosLoss = Val(aFields(9))
        m.lBarsInTrade = Val(aFields(10))
        
        ' Skip the last trade if still open 11/4/2003 DAJ...
        If Len(aFields(6)) = 0 Then
            m.bSkip = True
        Else
            m.bSkip = False
        End If
            
        If m.dUnits <> 0 And Not m.bSkip Then
            If bStocks Then
                ' TLB 7/25/2013: for stocks, the "margin" is just the largest entry cost
                dEntryCost = Val(aFields(2)) * Val(aFields(18))
                If dEntryCost > m.cMargin Then
                    m.cMargin = dEntryCost
                End If
            End If
            AddToTotals
            CheckConsec
            CheckLargest
            CheckDrawdown False
            m.lTotalTrades = m.lTotalTrades + 1
        End If
        
    Next m.X
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStats.gdCalculate", eGDRaiseError_Raise
    
End Sub

Private Sub AddToTotals()
On Error GoTo ErrSection:

    Dim YrReturn            As Double

    m.cAccountBalance = m.cAccountBalance + (m.dUnits * m.cProfit)
    
    If m.cProfit > m.cMinProfit Then
        m.lWins = m.lWins + 1
        m.cWinAmt = m.cWinAmt + m.cProfit
        m.lWinBarsInTrades = m.lWinBarsInTrades + m.lBarsInTrade
    Else
        m.lLosses = m.lLosses + 1
        m.cLossAmt = m.cLossAmt + m.cProfit
        m.lLossBarsInTrades = m.lLossBarsInTrades + m.lBarsInTrade
    End If
    
    'If year changes, determine the return for that year
    If Year(m.dExitDate) > m.lSaveYear Then
        If m.cLastBalance > 0 Then
            YrReturn = (m.cAccountBalance - m.cLastBalance) / m.cLastBalance
        End If
        m.lYears = m.lYears + 1
        m.cLastBalance = m.cAccountBalance
        m.dYrReturnTotal = m.dYrReturnTotal + YrReturn
        m.lSaveYear = Year(m.dExitDate)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStats.AddToTotals", eGDRaiseError_Raise
    
End Sub

Private Sub CheckConsec()
On Error GoTo ErrSection:

    If m.cProfit > m.cMinProfit Then
        m.lWinConsec = m.lWinConsec + 1
        If m.lWinConsec > m.lWinConsecMax Then
            m.lWinConsecMax = m.lWinConsec
        End If
        ''If m.lLossConsec > 0 Then
        ''    m.lLossRuns = m.lLossRuns + 1
        ''End If
        If m.lWinConsec = 1 Then
            m.lWinRuns = m.lWinRuns + 1
        End If
        m.lLossConsec = 0
    Else
        m.lLossConsec = m.lLossConsec + 1
        If m.lLossConsec > m.lLossConsecMax Then
            m.lLossConsecMax = m.lLossConsec
        End If
        ''If m.lWinConsec > 0 Then
        ''    m.lWinRuns = m.lWinRuns + 1
        ''End If
        If m.lLossConsec = 1 Then
            m.lLossRuns = m.lLossRuns + 1
        End If
        m.lWinConsec = 0
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStats.CheckConsec", eGDRaiseError_Raise
    
End Sub

Private Sub CheckLargest()
On Error GoTo ErrSection:

    If m.cProfit > m.cMinProfit Then
        If m.cProfit > m.cWinLargest Then
            m.cWinLargest = m.cProfit
        End If
    Else
        If m.cProfit < m.cLossLargest Then
            m.cLossLargest = m.cProfit
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStats.CheckLargest", eGDRaiseError_Raise
    
End Sub

'Keep track of the drawdown using money management
Private Sub CheckDrawdown(pInitialize As Boolean)
On Error GoTo ErrSection:

    Static MaxDrawdownClosed        As Currency
    Static MaxDrawdownIntra         As Currency
    Static PrevDrawdown             As Currency
    Static CloseDrawdown            As Currency
    Dim PeakToValleyPct             As Double
    
    'This corrects problem when calling multiple times for yearly stats
    If pInitialize Then
        MaxDrawdownClosed = 0
        MaxDrawdownIntra = 0
        PrevDrawdown = 0
        CloseDrawdown = 0
        CheckDrawdown1Unit True
        Exit Sub
    End If
    
    ' Check for new equity peak
    If m.cProfit < m.cMinProfit Then
        m.bNewEquityPeak = False
    Else
        If m.cAccountBalance > m.cEquityPeak Then
            m.bNewEquityPeak = True
            m.cEquityPeak = m.cAccountBalance
        Else
            m.bNewEquityPeak = False
        End If
    End If
    
    If Not m.bNewEquityPeak Then
    
        PrevDrawdown = m.cClosedDrawdown
        m.cClosedDrawdown = m.cAccountBalance - m.cEquityPeak
        If m.cClosedDrawdown < MaxDrawdownClosed Then
            MaxDrawdownClosed = m.cClosedDrawdown
        End If
        
        If m.cEquityPeak > 0 Then
            PeakToValleyPct = m.cClosedDrawdown / m.cEquityPeak
            If PeakToValleyPct < m.dPeakToValleyDrawdown Then
                m.dPeakToValleyDrawdown = PeakToValleyPct
            End If
        End If
    End If
    
    'Check Max Intra-Day drawdown
    If m.X = 1 Then
        m.cIntraDrawdown = m.cOpenPosLoss
    Else
        m.cIntraDrawdown = (PrevDrawdown + m.cOpenPosLoss)
    End If
    If m.cIntraDrawdown < MaxDrawdownIntra Then
        MaxDrawdownIntra = m.cIntraDrawdown
    End If
    
    CheckDrawdown1Unit False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStats.CheckDrawdown", eGDRaiseError_Raise
    
End Sub

'Keep track of the drawdown (always for 1 unit)
Private Sub CheckDrawdown1Unit(pInitialize As Boolean)
On Error GoTo ErrSection:

    Static EquityPeak       As Currency
    Static PrevDrawdown     As Currency
    Static ClosedDrawdown   As Currency
    
    If pInitialize Then
        EquityPeak = 0
        PrevDrawdown = 0
        ClosedDrawdown = 0
        Exit Sub
    End If
    
    m.cNetProfit = m.cNetProfit + m.cProfit
    
    ' Check for new equity peak
    If m.cProfit < m.cMinProfit Then
        m.bNewEquityPeak = False
    Else
        If m.cNetProfit > EquityPeak Then
            m.bNewEquityPeak = True
            EquityPeak = m.cNetProfit
        Else
            m.bNewEquityPeak = False
        End If
    End If
    
    If Not m.bNewEquityPeak Then
    
        PrevDrawdown = ClosedDrawdown
        ClosedDrawdown = m.cNetProfit - EquityPeak
        
        If ClosedDrawdown < m.cMaxDrawdownClosed Then
            m.cMaxDrawdownClosed = ClosedDrawdown
        End If
        
    End If
    
    'Check Max Intra-Day drawdown
    If m.X = 1 Then
        m.cIntraDrawdown = m.cOpenPosLoss
    Else
        m.cIntraDrawdown = (PrevDrawdown + m.cOpenPosLoss)
    End If
    If m.cIntraDrawdown < m.cMaxDrawdownIntra Then
        m.cMaxDrawdownIntra = m.cIntraDrawdown
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStats.CheckDrawdown1Unit", eGDRaiseError_Raise
    
End Sub

