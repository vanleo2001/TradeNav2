VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMonteCarlo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum eStatRow
    eNumTrades = 0
    eAvgTrade
    eWinPerc
    ePayoutRatio
    eAvgWin
    eAvgLoss
    eProfitFactor
    eCPCIndex
    eExpectancy
    eKellyRatio
    eAvgProfit
    eAvgDD
    eProfitDD
    eSTAT_ROWS
End Enum

Private Type mPrivate
    aTrades As cGdArray
    aRisks As cGdArray

    ' basic trade stats for systems
    dWinPercent As Double
    dAvgWin As Double
    dAvgLoss As Double
    dAvgTrade As Double
    
    ' parameters for simulations
    dStartBalance As Double
    
    ' for avg performance stats
    dWon As Double
    dLost As Double
    nNumWins As Long
    nNumLosses As Long
    dAvgProfit As Double
    dAvgDrawdown As Double
    
    bRunning As Boolean ' flag when in progress
End Type
Private m As mPrivate


Private Sub Class_Initialize()
    
    Randomize
    Set m.aTrades = New cGdArray
    m.aTrades.Create eGDARRAY_Doubles, 0
    Set m.aRisks = New cGdArray
    m.aRisks.Create eGDARRAY_Doubles, 0
    
End Sub

Private Sub Class_Terminate()
    
    Set m.aTrades = Nothing
    Set m.aRisks = Nothing
    
End Sub

Private Sub CalcTradeStats()

    Dim i&, iWins&, iLosses&, dTrade#

    m.dAvgLoss = 0
    m.dAvgWin = 0
    m.dAvgTrade = 0
    m.dWinPercent = 0
    
    If m.aTrades.Size > 0 Then
        ' sum up the wins and losses
        For i = 0 To m.aTrades.Size - 1
            dTrade = m.aTrades.Num(i)
            If dTrade > 0 Then
                iWins = iWins + 1
                m.dAvgWin = m.dAvgWin + dTrade
            Else
                iLosses = iLosses + 1
                m.dAvgLoss = m.dAvgLoss + dTrade
            End If
        Next
        m.dAvgTrade = m.dAvgWin + m.dAvgLoss
        ' then calc the averages
        If iWins > 0 Then
            m.dAvgWin = m.dAvgWin / iWins
        End If
        If iLosses > 0 Then
            m.dAvgLoss = m.dAvgLoss / iLosses
        End If
        If iWins + iLosses > 0 Then
            m.dAvgTrade = m.dAvgTrade / (iWins + iLosses)
            m.dWinPercent = iWins / (iWins + iLosses)
        End If
    End If

End Sub

Public Function SetTrades(ByVal tTrades As cGdTable) As Boolean

    m.aTrades.Size = 0
    m.aRisks.Size = 0
    If Not tTrades Is Nothing Then
        If tTrades.FieldType(0) = eGDARRAY_Doubles Then
            Set m.aTrades = tTrades.FieldArray(0, True)
            SetTrades = True
        End If
        If tTrades.NumFields > 1 Then
            Set m.aRisks = tTrades.FieldArray(1, True)
        End If
    End If
    
    CalcTradeStats
    
End Function

Public Function ReadTrades(ByVal strFile$) As Boolean

    Dim fh%, d#, strLine$, strChk$, bHeaderDone As Boolean, bSuccess As Boolean
    
    On Error GoTo ErrExit:
    
    m.aTrades.Create eGDARRAY_Doubles, 0
    CalcTradeStats ' to clear stats
    
    If FileExist(strFile) Then
        fh = FreeFile
        Open strFile For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strLine
            strChk = Left(strLine, 2)
            If strChk = "S" & vbTab Or strChk = "L" & vbTab Then
                d = Val(Parse(strLine, vbTab, 8))
                m.aTrades.Add d
            End If
        Loop
        Close #fh
    End If
    bSuccess = True

ErrExit:
    If Not bSuccess Then
        m.aTrades.Size = 0
    End If
    CalcTradeStats
    Exit Function
    
End Function


Private Sub InitRiskGrid(fgRisks As VSFlexGrid)

    If fgRisks Is Nothing Then Exit Sub
                
    With fgRisks
        .Rows = 2
        SetupGrid fgRisks, eGridMode_Grid
        
        .BackColorFixed = &HD1E7ED
        .ExtendLastCol = True
        .SelectionMode = flexSelectionFree
        .ExplorerBar = flexExNone
        '.BackColorBkg = Me.BackColor
        .ScrollBars = flexScrollBarVertical
        
        .Cols = 4
        If 0 Then
            ' without "% Failures"
            .ColHidden(3) = True
            .Width = 4800
        Else
            ' with "% Failures"
            .ColHidden(3) = False
            .Width = 6000
        End If
        .FixedRows = 1 '2
        .Rows = .FixedRows
        .ColAlignment(0) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .TextMatrix(0, 0) = "Risk%"
        If .FixedRows > 1 Then
            .MergeCells = flexMergeFixedOnly
            .MergeRow(0) = True
            .MergeCol(0) = True
            .TextMatrix(1, 0) = "Risk%"
            .TextMatrix(0, 1) = "System"
            '.Cell(flexcpForeColor, 0, 1, 0, 3) = lblSystemA.ForeColor
            .Cell(flexcpFontSize, 0, 1, 0, .Cols - 1) = 10
        Else
            .Cell(flexcpFontSize, 0, 0) = 10
        End If
        
        .TextMatrix(0, 2) = .TextMatrix(0, 1)
        .TextMatrix(0, 3) = .TextMatrix(0, 2)
        .RowHeight(0) = 400 'Int(.RowHeight(1) * 1.25)
        
        .TextMatrix(.FixedRows - 1, 1) = "Avg Ending Balance"
        .TextMatrix(.FixedRows - 1, 2) = "Avg Drawdown%"
        .TextMatrix(.FixedRows - 1, 3) = "% Failures" '"% Bankrupts"
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
        '.Select 0, 1, 2, 3
        '.CellBorder RGB(0, 0, 128), 2, 2, 2, 2, 0, 0
        '.Select 0, 0
    End With

End Sub

Private Sub DisplayStats(ReportForm As frmReports, ByVal nNumTrades&)

    Dim i&, d#
    Dim dAvgTrade#, dAvgWin#, dAvgLoss#, dWinPerc#, dPayoutRatio#, dProfitFactor#, dCPC#, dExpectancy#, dKelly#
  
    With ReportForm.fgDDStats
        .Rows = eSTAT_ROWS
        .ScrollBars = flexScrollBarNone
        .ColWidth(0) = .ClientWidth / 2
        '.RowHeight(.Rows - 1) = .RowHeight(0) * 1.25
        .Cell(flexcpFontBold, 0, 1, .Rows - 1, 1) = True
        .Cell(flexcpFontBold, .Rows - 3, 0, .Rows - 1, 0) = True
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .Cell(flexcpFontSize, .Rows - 3, 1, .Rows - 2, 1) = 12
        .Cell(flexcpFontSize, .Rows - 1, 1) = 16
        
        .TextMatrix(eNumTrades, 0) = "Trades per Year"
        .TextMatrix(eAvgTrade, 0) = "Average Trade"
        .TextMatrix(eWinPerc, 0) = "% Winning Trades"
        .TextMatrix(ePayoutRatio, 0) = "Payout Ratio"
        .TextMatrix(eAvgWin, 0) = "Average $ Win"
        .TextMatrix(eAvgLoss, 0) = "Average $ Loss"
        .TextMatrix(eProfitFactor, 0) = "Profit Factor"
        .TextMatrix(eCPCIndex, 0) = "CPC Index"
        .TextMatrix(eExpectancy, 0) = "Expectancy%"
        .TextMatrix(eKellyRatio, 0) = "Kelly Ratio"
        .TextMatrix(eAvgProfit, 0) = "Average Annual" & vbCrLf & "PROFIT"
        .TextMatrix(eAvgDD, 0) = "Average Annual" & vbCrLf & "DRAWDOWN"
        .TextMatrix(eProfitDD, 0) = "Average Annual" & vbCrLf & "Profit / Drawdown" & vbCrLf & "RATIO"
        If m.nNumLosses > 0 And m.nNumWins > 0 Then
            .Visible = True
        
            dAvgTrade = (m.dWon + m.dLost) / (m.nNumWins + m.nNumLosses)
            dWinPerc = m.nNumWins / (m.nNumWins + m.nNumLosses)
            dAvgWin = m.dWon / m.nNumWins
            dAvgLoss = Abs(m.dLost / m.nNumLosses)
            dPayoutRatio = dAvgWin / dAvgLoss
            dProfitFactor = Abs(m.dWon / m.dLost)
            dCPC = (dProfitFactor ^ 2) * (1 - dWinPerc)
            dExpectancy = (1 + dAvgWin / dAvgLoss) * dWinPerc - 1
            dKelly = dWinPerc - (1 - dWinPerc) / dPayoutRatio
            
            .TextMatrix(eNumTrades, 1) = Format(nNumTrades, "#0")
            .TextMatrix(eAvgTrade, 1) = Format(dAvgTrade, "$#,##0.00")
            .TextMatrix(eWinPerc, 1) = Format(dWinPerc, "#0.00%")
            .TextMatrix(eAvgWin, 1) = Format(dAvgWin, "$#,##0.00")
            .TextMatrix(eAvgLoss, 1) = Format(dAvgLoss, "$#,##0.00")
            .TextMatrix(ePayoutRatio, 1) = Format(dPayoutRatio, "#0.000")
            .TextMatrix(eProfitFactor, 1) = Format(dProfitFactor, "#0.000")
            .TextMatrix(eCPCIndex, 1) = Format(dCPC, "#0.000")
            .TextMatrix(eExpectancy, 1) = Format(dExpectancy, "#0.00%")
            .TextMatrix(eKellyRatio, 1) = Format(dKelly, "#0.00%")
            .TextMatrix(eAvgProfit, 1) = Format(Abs(m.dAvgProfit), "$#,##0")
            .TextMatrix(eAvgDD, 1) = Format(Abs(m.dAvgDrawdown), "$#,##0")
            If m.dAvgDrawdown <> 0 Then
                .TextMatrix(eProfitDD, 1) = Format(Abs(m.dAvgProfit / m.dAvgDrawdown), "#0.00")
            End If
        Else
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = ""
            Next
        End If
        
        ' make the grid the correct height (sum of the rows)
        d = 0
        For i = 0 To .Rows - 1
            If i = .Rows - 1 Then
                .RowHeight(i) = 700
            ElseIf i >= .Rows - 3 Then
                .RowHeight(i) = 450
            Else
                .RowHeight(i) = 300
            End If
            d = d + .RowHeight(i)
        Next
        .Height = d + (.Height - .ClientHeight)
        '.AutoSize 0, 1, True
    End With

End Sub

Private Function GetPercentile(aArray As cGdArray, ByVal dPercentile#) As Double

    Dim i&, d#
    
    If dPercentile <= 0 Then
        d = 0
    ElseIf dPercentile >= 100 Then
        d = aArray.Size - 1
    Else
        d = CDbl(aArray.Size - 1) * dPercentile / 100#
    End If
    
    If d - Int(d) = 0.5 Then
        ' get average of the 2 values
        i = Int(d)
        GetPercentile = (aArray.Num(i) + aArray.Num(i + 1)) / 2
    Else
        i = Round(d)
        GetPercentile = aArray.Num(i)
    End If
    
End Function


Public Sub RunPerformance(ReportForm As frmReports, ByVal nNumRuns&, ByVal nNumTrades&)

    Dim i&, d#, iMid&, iRun&, dProfit#, dDD#, iRefreshRows&, iColor&
    Dim dSumProfit As Double
    Dim dSumDD As Double
    Dim dMaxDD As Double
    Dim aMaxDrawdowns As New cGdArray

    If m.bRunning Or m.dAvgTrade = 0 Then Exit Sub
    m.bRunning = True
    
    aMaxDrawdowns.Create eGDARRAY_Doubles, 0
    
    m.dWon = 0
    m.dLost = 0
    m.nNumLosses = 0
    m.nNumWins = 0
    
    'DisplayStats
    
    With ReportForm
        'Screen.MousePointer = vbHourglass
        .pbRuns.Value = 0
        .pbRuns.Visible = True
        .pbRuns.ZOrder
        '.cmdPerformance.Enabled = False
        '.cmdRisks.Enabled = False
        '.lblPerfNote.Visible = False
        '.lblNumSims = ""
        '.lblNumSims.Visible = True
        .Refresh
    End With
    
    If nNumRuns >= 2500 Then
        iRefreshRows = 250
    Else
        iRefreshRows = 100
    End If
    For iRun = 1 To nNumRuns
        TestRun nNumTrades, 0, dProfit, dDD
        aMaxDrawdowns.Add dDD
        dSumProfit = dSumProfit + dProfit
        dSumDD = dSumDD + dDD
        If dDD > dMaxDD Then
            dMaxDD = dDD
        End If
        m.dAvgProfit = dSumProfit / iRun
        m.dAvgDrawdown = dSumDD / iRun
        If iRun Mod iRefreshRows = 0 Or iRun = nNumRuns Then
            ' update progress bar and labels
            With ReportForm
                .pbRuns.Value = Int(100# * iRun / nNumRuns + 0.5)
                DisplayStats ReportForm, nNumTrades
                '.lblNumSims = "Avg of " & Format(iRun, "#,##0") & " simulations"
                '.lblAvgProfit = Format(dSumProfit / iRun, "$#,##0")
                '.lblAvgDrawdown = Format(dSumDD / iRun, "$#,##0")
                '.lblAvgRatio = Format(dSumProfit / dSumDD, "#0.0")
            End With
            
            aMaxDrawdowns.Sort
            With ReportForm.fgDD
                '.ColHidden(2) = True
                '.Width = 3000
                .ColWidth(0) = 1300
                .ColWidth(1) = 1450
                .FixedRows = 1
                .Rows = 10
                .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
                .Cell(flexcpFontBold, 5, 0, 5, .Cols - 1) = True
                .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 10
                .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
                .ColAlignment(0) = flexAlignCenterCenter
                .ColAlignment(1) = flexAlignCenterCenter
                .HighLight = flexHighlightNever
                .ScrollBars = flexScrollBarNone
                
                .TextMatrix(0, 0) = "Percentile"
                .TextMatrix(0, 1) = "Drawdown"
                .TextMatrix(0, 2) = "Expect this ..." '"Profit / DD"
                .TextMatrix(1, 0) = "Smallest"
                .TextMatrix(2, 0) = "1st"
                .TextMatrix(3, 0) = "10th"
                .TextMatrix(4, 0) = "25th"
                .TextMatrix(5, 0) = "50th"
                .TextMatrix(6, 0) = "75th"
                .TextMatrix(7, 0) = "90th"
                .TextMatrix(8, 0) = "99th"
                .TextMatrix(9, 0) = "Largest"
                .TextMatrix(5, 2) = " most years"
                .TextMatrix(6, 2) = " once every 4 years"
                .TextMatrix(7, 2) = " once every 10 years"
                .TextMatrix(8, 2) = " once every 100 years"
                .TextMatrix(9, 2) = " worst-case scenario"
                
                .TextMatrix(1, 1) = Format(GetPercentile(aMaxDrawdowns, 0), "$#,##0")
                .TextMatrix(2, 1) = Format(GetPercentile(aMaxDrawdowns, 1), "$#,##0")
                .TextMatrix(3, 1) = Format(GetPercentile(aMaxDrawdowns, 10), "$#,##0")
                .TextMatrix(4, 1) = Format(GetPercentile(aMaxDrawdowns, 25), "$#,##0")
                .TextMatrix(5, 1) = Format(GetPercentile(aMaxDrawdowns, 50), "$#,##0")
                .TextMatrix(6, 1) = Format(GetPercentile(aMaxDrawdowns, 75), "$#,##0")
                .TextMatrix(7, 1) = Format(GetPercentile(aMaxDrawdowns, 90), "$#,##0")
                .TextMatrix(8, 1) = Format(GetPercentile(aMaxDrawdowns, 99), "$#,##0")
                .TextMatrix(9, 1) = Format(GetPercentile(aMaxDrawdowns, 100), "$#,##0")
                
                '.AutoSize 0, .Cols - 1, True
                
                d = 0
                For i = 0 To .Rows - 1
                    If i = 0 Then
                        .RowHeight(i) = 400
                    Else
                        .RowHeight(i) = 350
                        iColor = GradientColor(100 * Abs(5 - i) / 4, RGB(240, 240, 255), RGB(192, 192, 255))
                        .Cell(flexcpBackColor, i, 1) = iColor
                    End If
                    d = d + .RowHeight(i)
                Next
                .Height = d + (.Height - .ClientHeight)
                .Visible = True
            End With
            DoEvents
        End If
        If Not m.bRunning Then Exit For ' user aborting
    Next
    
    With ReportForm
        .pbRuns.Visible = False
        '.cmdPerformance.Enabled = True
        '.cmdRisks.Enabled = True
        '.lblPerfNote.Visible = True
        '.lblNumSims.Visible = False
        'Screen.MousePointer = vbDefault
    End With
    m.bRunning = False
    
End Sub


'Public Sub RunRisks(fgRisks As VSFlexGrid, ByVal nNumRuns&, ByVal nNumTrades&, ByVal dStartBalance#)
Public Sub RunRisks(ReportForm As frmReports, ByVal nNumRuns&, ByVal nNumTrades&, ByVal dStartBalance#)

    Dim bDone As Boolean
    Dim i&, iRow&, iMid&, s$
    Dim dRisk#, iRun&, iFailures&, dAvgEndBalance#, dEndBalance#
    Dim dDDP#, dAvgDDP#, dMaxDDP# ' drawdown percentages
    Dim aEndBalances As New cGdArray, aDrawdowns As New cGdArray

    m.bRunning = True
    
    With ReportForm.fgRisks
        InitRiskGrid ReportForm.fgRisks
        iRow = .FixedRows
        .Rows = iRow
        .Visible = True
    End With
    ReportForm.fraColors.Visible = True
    ReportForm.txtFailures.Visible = True
    ReportForm.Refresh
    
    m.dStartBalance = dStartBalance

    ' NOTE: when doing money management analysis, the "normal" averaging doesn't work
    ' very well due to the worst-case scenarios getting stopped out at zero (as opposed
    ' to going negative) -- so the much better "averaging" method is to use the MEDIAN
    ' instead (which requires storing all of them and sorting when done).
    aEndBalances.Create eGDARRAY_Floats, nNumRuns, 0
    aDrawdowns.Create eGDARRAY_Floats, nNumRuns, 0
    
    dRisk = 0
    Do While m.dAvgTrade > 0 And Not bDone
        If dRisk < 10 Then
            dRisk = dRisk + 0.5
        ElseIf dRisk < 50 Then
            dRisk = dRisk + 1
        Else
            dRisk = dRisk + 5
        End If
        If dRisk > 95 Then Exit Do
        
        dAvgEndBalance = 0
        dAvgDDP = 0
        dMaxDDP = 0
        iFailures = 0
        aEndBalances.Clear False
        aDrawdowns.Clear False
        For iRun = 1 To nNumRuns
            ' run a simulation for this risk%
            If TestRun(nNumTrades, dRisk / 100#, dEndBalance, dDDP) = False Then
                'iBankrupts = iBankrupts + 1
                dEndBalance = 0
                dDDP = 1 ' bankrupt = 100% drawdown
            End If
            ' check if failed to make money
            If dEndBalance <= dStartBalance Then
                iFailures = iFailures + 1
            End If
            aEndBalances.Num(iRun - 1) = dEndBalance
            aDrawdowns.Num(iRun - 1) = dDDP
            If dDDP > dMaxDDP Then
                dMaxDDP = dDDP
            End If
            If Not m.bRunning Then Exit For ' user aborting
        Next
        If Not m.bRunning Then Exit Do ' user aborting
        
        ' sort results and find the median (do the average of 6 middle spots)
        aEndBalances.Sort
        aDrawdowns.Sort
        dAvgEndBalance = 0
        dAvgDDP = 0
        iMid = Int(nNumRuns / 2#)
        For i = iMid - 3 To iMid + 2
            dAvgEndBalance = dAvgEndBalance + aEndBalances.Num(i)
            dAvgDDP = dAvgDDP + aDrawdowns.Num(i)
        Next
        dAvgEndBalance = dAvgEndBalance / 6
        dAvgDDP = dAvgDDP / 6
        
        ' output to grid row: AvgEndBalance, AvgDD%, MaxDD%, #Bankrupts
        With ReportForm.fgRisks
            .MergeCells = flexMergeFree
            .Rows = iRow + 1
            s = Str(dRisk)
            If Left(s, 1) = "." Or Left(s, 1) = "," Then
                s = "0" & s
            End If
            .TextMatrix(iRow, 0) = s & "%"
            '.Select 0, 1, .Rows - 1, 1
            '.CellBorder RGB(0, 0, 0), 2, 0, 0, 0, 0, 0
            '.Select 0, 4, .Rows - 1, 4
            '.CellBorder RGB(0, 0, 0), 2, 0, 0, 0, 0, 0
            .Select 0, 0, 0, 0
            .Cell(flexcpFontBold, iRow, 2, iRow, 2) = True
            .Cell(flexcpFontBold, iRow, 0, iRow, 0) = True
            
            If dAvgEndBalance = dStartBalance And dAvgDDP = 0 Then
                .MergeRow(iRow) = True
                s = " starting balance too low"
                .TextMatrix(iRow, 1) = s
                .TextMatrix(iRow, 2) = s
                .TextMatrix(iRow, 3) = "100%"
                .Select iRow, 1, iRow, 2
                .CellAlignment = flexAlignCenterCenter
                .Cell(flexcpBackColor, iRow, 1, iRow, 3) = RGB(224, 224, 224)
            Else
                .MergeRow(iRow) = False
                .TextMatrix(iRow, 1) = Format(dAvgEndBalance, "$#,##0")
                .TextMatrix(iRow, 2) = Format(dAvgDDP, "#0.00%")
                .TextMatrix(iRow, 3) = Format(CDbl(iFailures) / nNumRuns, "#0.0%")
                '.TextMatrix(iRow, 4) = Format(dMaxDDP, "#0.00%")
                i = Int(dAvgDDP * 10)
                If i > 9 Then
                    i = 9
                End If
                .Cell(flexcpBackColor, iRow, 1, iRow, 3) = ReportForm.lblColor(i).BackColor
            End If
        End With
        DoEvents
        
        ' don't bother continuing once get past a certain point
        If dAvgEndBalance <= 0 Or dAvgDDP > 0.98 Then
            bDone = True
        End If
        
        If bDone Then
            Exit Do
        End If
        iRow = iRow + 1
    Loop
    
    m.bRunning = False

End Sub


Private Function TestRun(ByVal nNumTrades&, ByVal dRiskPercent#, dBalance#, dMaxDD#) As Boolean

    Dim i&, iTrade&, dRandom#, dPeak#, dTrade#, dNumRisk#, iShuffleTrade&, dRiskAmt#
    Dim aShuffled As New cGdArray
    Dim bUseBacktest As Boolean, bUseShuffleMethod As Boolean
        
    If m.aTrades.Size > 0 Then
        bUseBacktest = True
'bUseShuffleMethod = True
        If bUseShuffleMethod Then
            Set aShuffled = m.aTrades.MakeCopy
            gdShuffle aShuffled.ArrayHandle, 0, aShuffled.Size - 1
        End If
    End If
        
    If dRiskPercent > 0 Then
        dBalance = m.dStartBalance
    Else
        dBalance = 0
    End If
    dPeak = dBalance
    dMaxDD = 0
    
    For iTrade = 1 To nNumTrades
        ' random trade
        If Not bUseBacktest Then
            ' randomly determine win or loss
            dRandom = Rnd 'dRandom = gdRandomNumber(0, 999999999) / 1000000000#
            If dRandom < m.dWinPercent Then
                dTrade = m.dAvgWin
            Else
                dTrade = -Abs(m.dAvgLoss)
            End If
        ElseIf bUseShuffleMethod Then
            ' get a shuffled trade from the backtest
            dTrade = aShuffled.Num(iShuffleTrade)
            iShuffleTrade = iShuffleTrade + 1
            If iShuffleTrade >= aShuffled.Size Then
                iShuffleTrade = 0
            End If
        Else
            ' get a random trade from the backtest
            i = gdRandomNumber(0, m.aTrades.Size - 1)
            dTrade = m.aTrades.Num(i)
            dRiskAmt = m.aRisks.Num(i)
            If dRiskAmt = m.aRisks.NullValue Then
                dRiskAmt = 0
            End If
        End If
        
        If dRiskPercent = 0 Then
            dBalance = dBalance + dTrade
            ' calc $MaxDD if 1 contract per trade
            If dBalance > dPeak Then
                dPeak = dBalance
            ElseIf dPeak - dBalance > dMaxDD Then
                dMaxDD = dPeak - dBalance
            End If
            ' accumulate for the average performance stats
            If dTrade > 0 Then
                m.dWon = m.dWon + dTrade
                m.nNumWins = m.nNumWins + 1
            Else
                m.dLost = m.dLost + dTrade
                m.nNumLosses = m.nNumLosses + 1
            End If
        Else
            ' for Money Management analysis, determine
            ' # contracts to risk for this trade
            If dRiskAmt = 0 Then
                dRiskAmt = m.dAvgLoss
            End If
            If dRiskAmt = 0 Then
                ' if can't determine a $RiskAmt, then no way to determine # of contracts to trade!
                dNumRisk = 0
            Else
                dNumRisk = Int(dBalance * dRiskPercent / Abs(dRiskAmt))
            End If
            If dNumRisk < 1 Then dNumRisk = 0 '1
            dBalance = dBalance + dTrade * dNumRisk
            If dBalance <= 0 Then
                ' BANKRUPT !!
                dBalance = 0
                dMaxDD = 1 ' bankrupt = 100% drawdown
                TestRun = False
                Exit Function
            End If
            
            ' calc MaxDD% if money management
            If dBalance > dPeak Then
                dPeak = dBalance
            ElseIf dPeak > 0 Then
                If (dPeak - dBalance) / dPeak > dMaxDD Then
                    dMaxDD = (dPeak - dBalance) / dPeak
                End If
            End If
                        
            ' allow for user to abort a MM run
            If iTrade Mod 1000 = 0 Then
'                DoEvents
            End If
            If Not m.bRunning Then Exit For ' user aborting
        End If
    Next
    
    TestRun = True

End Function

