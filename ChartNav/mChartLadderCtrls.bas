Attribute VB_Name = "mChartLadderCtrls"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mChartLadderCtrls.bas
'' Description: Shared code for populating controls on various chart, price ladder
''              and depth of market forms (include config forms)
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial, Suite 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/01/2008   MJM         Created
'' 01/21/2009   DAJ         Only give the wrong side of the market warning if the
''                          trade settings say to give the warning
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 04/15/2013   DAJ         Only allow auto exits, TSOG, and automated trading if enabled for streaming
'' 04/16/2014   DAJ         Fix for submitting TradeSense order group via favorites ( Pete Laverde )
'' 03/02/2016   DAJ         Fixed ShortDisplayNumber function for millions
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'short/long frame background for order bar (Glen does not want users to change this)
Public Const kFrameLive = &H80FFFF
Public Const kFrameShort = &HA0A0FF
Public Const kFrameLong = &HFFA0A0
'default colors for coloring bid/ask controls
Public Const kBidColor = 16761024 '16769248
Public Const kAskColor = 12632319 '14737663
'default columns for account bar
Public Const kABarCols = "Position|Avg Entry|Open Equity|Acct Balance|Daily P/L|# Contracts"
'configurable order bar defaults (0=no checkbox,1=checked,2=unchecked)
Public Const kOrdWizardDefaults = "POS;1|OE;1|BM;2|SM;2|OT;2|BC;2|SC;2|CLR;0|QTY;2|AJ;2|AE;2|PR;2|RV;2|CA;2|FL;2|BA;2|SA;2|BB;2|SB;2|"
Public Const kOrdBarDefaults = "POS;1|OE;1|BM;1|SM;1|CLR;0|QTY;2|OT;2|BC;1|SC;1|BRORD;1|CO;0|AJ;2|AE;1|TSO;2|RV;2|CA;1|FL;1|PR;2|BA;2|SA;2|BB;2|SB;2"
Public Const kOrdBarDisconnect = "POS;1|OE;1|BM;2|SM;2|CLR;0|QTY;2|OT;2|BC;2|SC;2|BRORD;2|AJ;2|AE;2|RV;2|CA;2|FL;2|PR;2|BA;2|SA;2|BB;2|SB;2|"

'default/minimum quantity for @IB & @CNX Forex
Public Const kIBCNXMinQty = 10000

Public Enum eGDOrderBarCfg
    eGDOrderBarCfg_ButtonShow = 0
    eGDOrderBarCfg_ButtonCode
    eGDOrderBarCfg_ButtonDescript
End Enum

'[begin] shared cons, enum & type for frmChart & frmChart2
Public Const kMinMoveCount As Long = 8
Public Const kMinAnnotSize As Long = 5
Public Const kMinTriChannSize As Long = 15
Public Const kMinOrderMove As Long = 3
Public Const kSextantGif = "tradenav.gif"
Public Const kWaterPng = "TNWaterMark.png"

Public Const kOrdBarWidth = 1440
Public Const kOrdWizBtnLeft = 400
Public Const kPriceWizLeft = 495        'left of price label for wizard prompt when "call" & "put" labels visible
Public Const kPriceWizWidth = 675       'width of price label for wizard prompt when "call" & "put" labels visible
Public Const kPFPBarWidth = 2580        '2490        'width for pattern for profit bar

Public Enum enumScaleFlag
    eScaleY_Max = 1
    eScaleY_Min = 2
    eScaleX_MoreLessBars = 3
    eScale_Arrow = 4                    'red arrrow is drawn (arrow visible)
    eScale_Arrow_Area = 5               'area where red arrow would be drawn but isn't (arrow not visible)
    eScale_Unhandled = 6
End Enum

Public Enum enumCmdMode
    eCmdMode_Off = 0
    eCmdMode_Symbol = 1
    eCmdMode_BarPeriod = 2
End Enum

Public Enum enumTipType
    eTiptype_None = 0
    eTiptype_Same = 1
    eTipType_Trade = 2
    eTipType_Annot = 3
    eTipType_OptGraph = 4
End Enum

Public Enum enumOneClickOrder
    eClickOrder_None = 0         'not a one click order
    eClickOrder_BuyMkt = 1
    eClickOrder_BuyBid = 2
    eClickOrder_BuyAsk = 3
    eClickOrder_SellMkt = 4
    eClickOrder_SellBid = 5
    eClickOrder_SellAsk = 6
    eClickOrder_BuyCall = 7
    eClickOrder_BuyPut = 8
    eClickOrder_SellCall = 9
    eClickOrder_SellPut = 10
End Enum

Public Enum enumDetachStatus
    eNotDetached = 0
    eDetachInProg = 1
    eAttachInProg = 2
    eDetached = 3
End Enum

Public Enum eOrderBarMode
    eOrdBarMode_Undefined = -1
    eOrdBarMode_Order = 0
    eOrdBarMode_Wizard = 1
    eOrdBarMode_PFP = 2
    eOrdBarMode_BrokerDisconnect = 3
End Enum

Public Enum ePfpResetMode
    ePfpReset_GridInd = 1
    ePfpReset_GridPfp = 2
    ePfpReset_Forecastbars = 3
    ePfpReset_PpfAnnot = 4
    ePfpReset_PpfPattern = 5
    ePfpReset_PpfAnnotInd = 6
    ePfpReset_ClearAll = 7
End Enum

Public Enum eSeasonalCtrlsStatus
    eSeasonCtrlStatus_Unpopulated
    eSeasonCtrlStatus_Updated
    eSeasonCtrlStatus_Restore
End Enum

Public Enum eSeasonalCtrlType
    eSeasonalCtrl_AvgTrendColor
    eSeasonalCtrl_BullTrendColor
    eSeasonalCtrl_BearTrendColor
    eSeasonalCtrl_CurrCycleColor
    eSeasonalCtrl_OtherColorFrom
    eSeasonalCtrl_OtherColorTo
    eSeasonalCtrl_CycleNum
    eSeasonalCtrl_CycleType
    eSeasonalCtrl_BarType
    eSeasonalCtrl_FromDate
    eSeasonalCtrl_AvgTrendCheckBox
    eSeasonalCtrl_AvgTrendStyle
    eSeasonalCtrl_BullTrendCheckBox
    eSeasonalCtrl_BullTrendStyle
    eSeasonalCtrl_BearTrendCheckBox
    eSeasonalCtrl_BearTrendStyle
    eSeasonalCtrl_CurrCycleCheckBox
    eSeasonalCtrl_CurrCycleStyle
    eSeasonalCtrl_SeasonalGrid
    eSeasonalCtrl_OverlayCycles
    eSeasonalCtrl_ShowCycles
    eSeasonalCtrl_OverlayTrends
End Enum

Public Enum eChartFlexCtrlIndex
    eFlexGridIdx_AcctBar                  'available when order bar is on (both options & non-options trading)
    eFlexGridIdx_OrdWizard                'available when trading options on order bar
    eFlexGridIdx_Seasonal                 'seasonal chart
    eFlexGridIdx_PfpInd                   'pattern for profit list of indicators
    eFlexGridIdx_PfpHits                  'pattern for profil list of hits
End Enum

Public Type frmChartPrivateType
    WindowLink As New cWindowLink

    eCmdMode As enumCmdMode
    eImgSrv As eImgSrvState
    
    iAutoSize As Integer

    bTopMost As Boolean

    MouseDown As ChartCoordinates
    MouseLast As ChartCoordinates

    Chart As cChart
    dLastScrollTime As Double
    bLockValuesDisplay As Boolean
    bPrevOffChart As Boolean
    bEditing As Boolean
    
    aTabs As cGdArray

    'for when user is manipulating annotations, indicators or scale areas
    nActiveOrderID As Long
    nActiveOrderLoc As Long         '1=triangle,2=X,3=connecting line
    nActiveIndIdx As Long
    nActiveAnnotIdx As Long
    nActiveAnnotPt As Long
    nAnnotClickTime As Long
    bAnnotCreated As Boolean
    bNewPattern As Boolean
    bNewPatternMoving As Boolean
    bIgnoreShift As Boolean
    bChartMoveInProg As Boolean     'flag to handle cursor moving outside client area during a move
    bScreenCaptured As Boolean
    nObjectMoving As Long           'mouse move counter
    nScrollChange As Long           'for setting scroll position when in ChartMove mode
    
    'flag to determine which scale area(X or Y) to manipulate and how
    eScaleFlag As enumScaleFlag
    
    'for when user is moving pane separators
    nActiveTopPane As Long
    nActiveBtmPane As Long
        
    'variables for Dinapoli Retracement/Expansion
    'identifies whether the first point clicked is closer to high or low of a bar
    '-1=none, 1=hi, 2=low
    '3=close (ignored for dinapoli focus point)
    '4=hi dynamic, 5=low dynamic
    nFocusHiLo As Long
    nPointCount As Long    'keeps count of # of points selected for dinapoli annotations
    
    'for picture box control
    epbCursor As enumCursor
    eCrossHairOn As enumCursor
    
    'for setting properties of new annotations
    AnnotOptions As cAnnotation
    
    nScaleStartPixel As Long
    nBarsToRight As Long
    nHsbMaxSave As Long
        
    'game mode
    bGameMode As Boolean
    bGameOrderMoving As Boolean
    oGameMode As cGameMode  'class object for controlling game
    eReplayModeSave As eGDReplayMode
    
    'account bar
    aABarColHeader As New cGdArray
    
    'order bar
    eOrdBarMode As eOrderBarMode
    
    'ID of split pane & label in split pane that was hit
    nSplitPaneHittestID As Long
    nSplitPaneLabelID As Double
    
    oToolTip As cToolTip
    strDetachedPlacment As String       'window placment when detached
    strNormalPlacement As String        ' window placement when not maximized or minimized (as fixed # twips)
    strRatioPlacement As String         ' window placement when not maximized or minimized (as ratio of MDI client area)

    bSettingAutoExit As Boolean
    bDrawToolJustCleared As Boolean
    
    eDetachStatus As enumDetachStatus
    
    bToolbarWrap As Boolean
    oBtnMouseLast As cPicBoxButton          'button object that mouse was last in
    nBtnsPerRow As Long                     'number of buttons per row when toolbar is wrapped
    
    aTbButtons As New cGdArray              'array of button objects for non-drawing toolbars
    aTbButtonsDraw As New cGdArray          'array of button objects for drawing toolbar
    
    bTradeContinuous As Boolean
'    bAllowDetach As Boolean                'JM 07-01-2013 - per Glen: allow detach for everyone
    bSkipFocusFix As Boolean
    bWindowLinkInitDone As Boolean

    'variables of options wizard bar
    bAllowOptWizard As Boolean
    bAllowRiskGraph As Boolean
    bResetOptWizardSpace As Boolean
    bDrawToolSelected As Boolean            'flag to indicate a draw tool was selected from the more-buttons form -5180
    bOptNavLoaded As Boolean                'flag to fix 6060, 6066
    strPrevRiskGraph As String
    
    'for bracket order
    oBracketOrdOne As cPtOrder
    oBracketOrdTwo As cPtOrder
    
    'for pattern for profit
    oPatternProfit As cPatternProfit
    bDropdownPFP As Boolean
    
    'for Schiff pitchfork
    bSchiffFork As Boolean
    
    'for seasonal controls
    eSeasonalCtrlsState As eSeasonalCtrlsStatus
    
    'for dynamic flex grids
    bFlexOrdBar As Boolean
    bFlexSeasonal As Boolean
    bFlexPFP As Boolean

    Quantity As cPriceEditor            ' Quantity editor
    lPreset1 As Long                    ' Quantity preset 1
    lPreset2 As Long                    ' Quantity preset 2
    lPreset3 As Long                    ' Quantity preset 3
End Type
'[end] shared const & enum for frmChart & frmChart2

'this sub used by price ladder & chart
Public Sub ResetGridBar(fg As VSFlexGrid, aHeader As cGdArray, nSumWidth As Long)
On Error GoTo ErrSection:

    Dim i&

    nSumWidth = 0
    With fg         'fgQuoteBar
        .Redraw = flexRDNone
        SetupGrid fg, eGridMode_Grid
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionFree
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarHorizontal
        .ExtendLastCol = True
        .AutoSizeMode = flexAutoSizeColWidth
        .FixedCols = 0
        .FixedRows = 0
        .Rows = 2
        .Cols = aHeader.Size
        'set max row height
        .RowHeightMax = .RowHeight(0)
        .RowHeight(1) = .RowHeightMax
        'headers
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = aHeader(i)
            If "Acct Balance" = aHeader(i) Then
                .ColFormat(i) = "$#,##0.00"
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = ALT_GRID_ROW_COLOR
        .AutoSize 0, .Cols - 1, True
        .Redraw = flexRDBuffered
        For i = 0 To .Cols - 1
            nSumWidth = nSumWidth + .ColWidth(i)
        Next
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartLadderCtrls.ResetGridBar", eGDRaiseError_Raise

End Sub

'this sub used by price ladder & chart
Public Sub GridBarHeader(fg As VSFlexGrid, aHeader As cGdArray, nSumWidth&, ByVal strText$)
On Error GoTo ErrSection:

    Dim i&
    Dim bReset As Boolean
    Dim aNewHeader As New cGdArray

    If Len(strText) < 1 Then
        Exit Sub
    End If

    aNewHeader.SplitFields strText, "|"
    
    If aHeader.Size = 0 Then
        For i = 0 To aNewHeader.Size - 1
            aHeader(i) = aNewHeader(i)
        Next
    ElseIf aHeader.Size <> aNewHeader.Size Then
        bReset = True
    ElseIf aHeader.Size = aNewHeader.Size Then
        For i = 0 To aHeader.Size - 1
            If aHeader(i) <> aNewHeader(i) Then
                aHeader(i) = aNewHeader(i)
                bReset = True
            End If
        Next
    End If
    
    If bReset Then
        aHeader.Size = 0
        fg.Cols = 0
        For i = 0 To aNewHeader.Size - 1
            aHeader(i) = aNewHeader(i)
        Next
        ResetGridBar fg, aHeader, nSumWidth
    End If
    
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.GridBarHeader"
    
End Sub

'this sub used by price ladder & chart
Public Sub UpdateAccountBar(fg As VSFlexGrid, strPos$, strPosQty$, strOpenEq$, _
     strAvgEntry$, strSessionPL$, strSessionQty$, nTradeAcctID&, _
     Optional ByVal strSecurityType As String = "")
On Error GoTo ErrSection:

    Dim Account As cPtAccount
    Dim i&, dCurrValue#, strText$
    
    If fg.Rows < 2 Then Exit Sub
    
    Set Account = g.Broker.Account(nTradeAcctID)
    If Not Account Is Nothing Then
        dCurrValue = Account.CurrentClosedBalance
    End If
    
    With fg
        .Redraw = flexRDNone
        For i = 0 To .Cols - 1
            strText = .TextMatrix(0, i)
            If InStr(strText, "#") > 0 Then
                .TextMatrix(1, i) = strSessionQty
                If Len(strSecurityType) > 0 Then
                    If InStr(strSecurityType, "S") > 0 Then
                        If InStr(UCase(strText), "SHARES") = 0 Then
                            .TextMatrix(0, i) = "# Shares"
                        End If
                    ElseIf InStr(UCase(strText), "CONTRACTS") = 0 Then
                        .TextMatrix(0, i) = "# Contracts"
                    End If
                End If
            Else
                Select Case strText
                    Case "Position"
                        If InStr(UCase(strPos), "LONG") > 0 Then
                            .TextMatrix(1, i) = strPos & " " & strPosQty
                            .Cell(flexcpForeColor, 1, i) = g.ChartGlobals.nLongColor
                        ElseIf InStr(UCase(strPos), "SHORT") > 0 Then
                            .TextMatrix(1, i) = strPos & " " & strPosQty
                            .Cell(flexcpForeColor, 1, i) = g.ChartGlobals.nShortColor
                        Else
                            .TextMatrix(1, i) = "Flat"
                            .Cell(flexcpForeColor, 1, i) = RGB(1, 1, 1)
                        End If
                    Case "Open Equity"
                        .TextMatrix(1, i) = strOpenEq
                        If InStr(UCase(strOpenEq), "-") > 0 Or InStr(UCase(strOpenEq), "(") > 0 Then
                            .Cell(flexcpForeColor, 1, i) = g.ChartGlobals.nLossColor
                        Else
                            .Cell(flexcpForeColor, 1, i) = g.ChartGlobals.nWinColor
                        End If
                    Case "Avg Entry"
                        If UCase(strPos) = "FLAT" Or Len(strPos) = 0 Then
                            .TextMatrix(1, i) = ""
                        Else
                            .TextMatrix(1, i) = strAvgEntry 'm.TickBars.PriceDisplay(m.dAvgEntry)
                        End If
                    Case "Acct Balance"
                        If dCurrValue < 0 Then
                            .Cell(flexcpForeColor, 1, i) = g.ChartGlobals.nLossColor
                        ElseIf dCurrValue > 0 Then
                            .Cell(flexcpForeColor, 1, i) = g.ChartGlobals.nWinColor
                        Else
                            .Cell(flexcpForeColor, 1, i) = vbBlack
                        End If
                        If dCurrValue = 0 Then
                            .TextMatrix(1, i) = ""
                        Else
                            .TextMatrix(1, i) = CStr(dCurrValue)
                        End If
                    Case "Daily P/L"
                        .TextMatrix(1, i) = strSessionPL
                        If InStr(strSessionPL, "-") Or InStr(strSessionPL, "(") Then
                            .Cell(flexcpForeColor, 1, i) = g.ChartGlobals.nLossColor
                        Else
                            .Cell(flexcpForeColor, 1, i) = g.ChartGlobals.nWinColor
                        End If
                End Select
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.UpdateAccountBar"
    
End Sub

Public Sub ParseOrdButtonString(ByVal strToParse$, nShow As Long, strCode As String, strDescription As String)
On Error GoTo ErrSection:
    
    nShow = Val(Parse(strToParse, ";", 2))
    strCode = Parse(strToParse, ";", 1)
    
'Constant       Value   Description
'flexNoCheckbox  0   The cell has no check box. This is the default setting.
'flexChecked     1   The cell has a check box that is checked.
'flexUnchecked   2   The cell has a check box that is not checked.
    
    If nShow <> 1 And nShow <> 2 Then nShow = 1
    
    If Len(strCode) = 0 Then
        strDescription = "ERROR: blank code"
    Else
        Select Case strCode
            Case "POS"
                strDescription = "Position"
            Case "OE"
                strDescription = "Open Equity"
            Case "BM"
                strDescription = "Buy Market"
            Case "SM"
                strDescription = "Sell Market"
            Case "CLR"
                strDescription = "Quantity"
                nShow = 0   'disallow hiding quantity controls (don't show check box)
            Case "QTY"
                strDescription = "Preset Quantities"
            Case "OT"
                strDescription = "Order Type"
            Case "BC"
                strDescription = "Buy Chart"
            Case "SC"
                strDescription = "Sell Chart"
            Case "AJ"
                strDescription = "Auto Journal"
            Case "AE"
                strDescription = "Auto Exits"
            Case "RV"
                strDescription = "Reverse"
            Case "CA"
                strDescription = "Cancel All"
            Case "FL"
                strDescription = "Flatten"
            Case "PR"
                strDescription = "Prices"
            Case "BA"
                strDescription = "Buy Ask"
            Case "SA"
                strDescription = "Sell Ask"
            Case "BB"
                strDescription = "Buy Bid"
            Case "SB"
                strDescription = "Sell Bid"
            Case "BRORD"
                strDescription = "Bracket Order"
            Case "CO"
                strDescription = "Confirm Orders"
                nShow = 0
            Case "TSO"
                strDescription = "Tradesense Groups"
            Case Else
                strDescription = "ERROR: " & strCode
        End Select
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.ParseOrdButtonString"

End Sub

Public Function OrdBarCtrlFromCode(aCodes As cGdArray, frm As Form, ByVal iIndex&) As Control
On Error GoTo ErrSection:

    Dim oNextControl As Control
    Dim strText As String
    
    Dim bVisible As Boolean
    Dim bBrokerViewMode As Boolean
    Dim bFavoritesBySym As Boolean
    
    Dim strSymbol$
    
'POS=Position, BM=BuyMarket, SM=SellMarket, CLR=quantity controls, QTY=preset quantities, OT=OrderType
'BC=BuyChart, SC=SellChart, BR=BracketOrders(old name for AutoExits), CA=CancelAll, FL=Flatten
'PR=bid/ask/market prices, BA=BuyAsk, SA=SellAsk, BB=BuyBid, SB=SellBid
'POS|BM|SM|CLR|QTY|OT|BC|SC|BR|RV|CA|FL|PR|BA|SA|BB|SB
    
    If aCodes Is Nothing Or frm Is Nothing Then Exit Function
    
    If IsFrmChart(frm) Then
        strSymbol = frm.Chart.Symbol
    ElseIf TypeOf frm Is frmTickDistribution Then
        strSymbol = g.SymbolPool.SymbolForID(frm.SymbolID)
        bBrokerViewMode = frm.BrokerViewMode
    Else
        Exit Function
    End If
    
    If Not IsFrmChart(frm) And Not TypeOf frm Is frmTickDistribution Then Exit Function

    If iIndex >= 0 And iIndex < aCodes.Size Then
        strText = aCodes(iIndex)
        If InStr(strText, ";2") = 0 Then bVisible = True
        Select Case strText
            Case "POS", "POS;1", "POS;2"
                Set oNextControl = frm.lblTradePos
                frm.lblTradePos.Visible = bVisible
            Case "OE", "OE;1", "OE;2"
                Set oNextControl = frm.lblEquity
                frm.lblEquity.Visible = bVisible
            Case "BM", "BM;1", "BM;2"
                Set oNextControl = frm.cmdBuyMarket
            Case "SM", "SM;1", "SM;2"
                Set oNextControl = frm.cmdSellMarket
            Case "CLR", "CLR;0", "CLR;1", "CLR;2"
                Set oNextControl = frm.cmdClearQty
                bVisible = True
                frm.vscrQty.Visible = True
                frm.txtTradeQty.Visible = True
            Case "QTY", "QTY;1", "QTY;2"
                Set oNextControl = frm.cmdQty1
                frm.cmdQty2.Visible = bVisible
                frm.cmdQty3.Visible = bVisible
            Case "OT", "OT;1", "OT;2"
                Set oNextControl = frm.lblOrderType
                frm.cboOrderType.Visible = bVisible
            Case "BC", "BC;1", "BC;2"
                Set oNextControl = frm.vseBuyChart
            Case "SC", "SC;1", "SC;2"
                Set oNextControl = frm.vseSellChart
            Case "AJ", "AJ;1", "AJ;2"
                Set oNextControl = frm.chkAutoJournal
                If bBrokerViewMode Then bVisible = False
            Case "BR", "AE;1", "AE;2"
                Set oNextControl = frm.chkAutoExit
                If bBrokerViewMode Then bVisible = False
                frm.lblAutoExit.Visible = bVisible
                'favorite exits buttons are shown only if they have been assigned
                If Len(ExitFavoritesAssigned(strSymbol)) > 0 Then
                    frm.fraExitFavorites.Visible = bVisible
                Else
                    frm.fraExitFavorites.Visible = False
                End If
            Case "RV", "RV;1", "RV;2"
                Set oNextControl = frm.cmdReverse
                If bBrokerViewMode Then bVisible = False
            Case "CA", "CA;1", "CA;2"
                Set oNextControl = frm.cmdCancelAll
                If bBrokerViewMode Then bVisible = False
            Case "FL", "FL;1", "FL;2"
                Set oNextControl = frm.cmdBailout
            Case "PR", "PR;1", "PR;2"
                Set oNextControl = frm.fraPrices
            Case "BA", "BA;1", "BA;2"
                Set oNextControl = frm.cmdBuyAsk
                If Not g.RealTime.Active Then bVisible = False
            Case "SA", "SA;1", "SA;2"
                Set oNextControl = frm.cmdSellAsk
                If Not g.RealTime.Active Then bVisible = False
            Case "BB", "BB;1", "BB;2"
                Set oNextControl = frm.cmdBuyBid
                If Not g.RealTime.Active Then bVisible = False
            Case "SB", "SB;1", "SB;2"
                Set oNextControl = frm.cmdSellBid
                If Not g.RealTime.Active Then bVisible = False
            Case "BRORD", "BRORD;1", "BRORD;2"
                Set oNextControl = frm.vseBracketOrder
                If bBrokerViewMode Then bVisible = False
            Case "CO", "CO;0", "CO;1", "CO;2"
                Set oNextControl = frm.chkConfirmOrder
                If bBrokerViewMode Then bVisible = False
            Case "TSO", "TSO;0", "TSO;1", "TSO;2"
                Set oNextControl = frm.fraTSO
        End Select

        If Not oNextControl Is Nothing Then oNextControl.Visible = bVisible
    End If
    
    Set OrdBarCtrlFromCode = oNextControl

ErrExit:
    Exit Function

ErrSection:
    DebugLog "mChartLadderCtrls.OrdBarCtrlFromCode processing " & strText
    RaiseError "mChartLadderCtrls.OrdBarCtrlFromCode"

End Function

Public Sub InitQBarGrid(fg As VSFlexGrid, aQBarHeaders As cGdArray)
On Error GoTo ErrSection:

    Dim i&, j&, strText$
    Dim aSorted As New cGdArray
    
    If fg Is Nothing Or aQBarHeaders Is Nothing Then Exit Sub
                
    With fg
        .Redraw = flexRDNone
        SetupGrid fg, eGridMode_Grid
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .Rows = 12
        .Cols = 2
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical
        .ColWidth(0) = 1000
        'headers
        .TextMatrix(0, 0) = "Show"
        .TextMatrix(0, 1) = "Data"
        
        .TextMatrix(1, 1) = "Symbol"
        .TextMatrix(2, 1) = "Bid"
        .TextMatrix(3, 1) = "Bid Size"
        .TextMatrix(4, 1) = "Ask"
        .TextMatrix(5, 1) = "Ask Size"
        .TextMatrix(6, 1) = "Trade"
        .TextMatrix(7, 1) = "Trade Size"
        .TextMatrix(8, 1) = "Open"
        .TextMatrix(9, 1) = "High"
        .TextMatrix(10, 1) = "Low"
        .TextMatrix(11, 1) = "Close"
        
        'set check boxes
        For i = 0 To aQBarHeaders.Size - 1
            aSorted.Add aQBarHeaders(i)
        Next
        If aSorted.Size > 0 Then
            aSorted.Sort
            For i = .FixedRows To .Rows - 1
                strText = .TextMatrix(i, 1)
                If aSorted.BinarySearch(strText, j) Then
                    If aSorted(j) = strText Then
                        .Cell(flexcpChecked, i, 0) = 1
                    Else
                        .Cell(flexcpChecked, i, 0) = 2
                    End If
                Else
                    .Cell(flexcpChecked, i, 0) = 2
                End If
            Next
        End If
        
        .Cell(flexcpAlignment, 0, 0, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 1, .Rows - 1, 1) = flexAlignLeftCenter
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexAlignCenterCenter
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartLadderCtrls.InitQBarGrid", eGDRaiseError_Raise

End Sub

Public Sub InitAccountGrid(fg As VSFlexGrid, aAcctHeaders As cGdArray, ByVal strSecType$)
On Error GoTo ErrSection:

    Dim i&, j&, strText$
    Dim aSorted As New cGdArray
    
    If fg Is Nothing Or aAcctHeaders Is Nothing Then Exit Sub
        
    With fg
        .Redraw = flexRDNone
        SetupGrid fg, eGridMode_Grid
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .Rows = 7
        .Cols = 2
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarNone
        .ColWidth(0) = 1000
        'headers
        .TextMatrix(0, 0) = "Show"
        .TextMatrix(0, 1) = "Data"
        
        .TextMatrix(1, 1) = "Position"
        .TextMatrix(2, 1) = "Avg Entry"
        .TextMatrix(3, 1) = "Open Equity"
        .TextMatrix(4, 1) = "Acct Balance"
        .TextMatrix(5, 1) = "Daily P/L"
        If strSecType = "S" Then
            .TextMatrix(6, 1) = "# Shares"
        Else
            .TextMatrix(6, 1) = "# Contracts"
        End If
        
        'set check boxes
        If Not aAcctHeaders Is Nothing Then
            For i = 0 To aAcctHeaders.Size - 1
                aSorted.Add aAcctHeaders(i)
            Next
            If aSorted.Size > 0 Then
                aSorted.Sort
                For i = .FixedRows To .Rows - 1
                    strText = .TextMatrix(i, 1)
                    If aSorted.BinarySearch(strText, j) Then
                        If aSorted(j) = strText Then
                            .Cell(flexcpChecked, i, 0) = 1
                        Else
                            .Cell(flexcpChecked, i, 0) = 2
                        End If
                    ElseIf InStr(strText, "#") And InStr(aSorted(j), "#") Then
                        .Cell(flexcpChecked, i, 0) = 1      'precautionary
                    Else
                        .Cell(flexcpChecked, i, 0) = 2
                    End If
                Next
            End If
        End If
        
        .Cell(flexcpAlignment, 0, 0, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 1, .Rows - 1, 1) = flexAlignLeftCenter
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexAlignCenterCenter
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        .Height = .RowHeight(0) * .Rows + .RowHeight(0) / 3
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartLadderCtrls.InitAccountGrid", eGDRaiseError_Raise

End Sub

Public Sub InitOrderButtonsGrid(fg As VSFlexGrid, ByVal strButtons$)
On Error GoTo ErrSection:

    Dim i&
    Dim aButtons As New cGdArray
    
    Dim nShow As Long
    Dim strCode As String
    Dim strDescription As String
    
    If InStr(strButtons, ";") = 0 Then strButtons = kOrdBarDefaults
    aButtons.SplitFields strButtons, "|"
    
    With fg
        .Redraw = flexRDNone
        
        SetupGrid fg, eGridMode_Grid
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionByRow
        .FixedRows = 1
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical
        .Cols = 3
        .Rows = 1
        .ColWidth(0) = 600
        .ColHidden(eGDOrderBarCfg_ButtonCode) = True
        'headers
        .TextMatrix(0, eGDOrderBarCfg_ButtonShow) = "Show"
        .TextMatrix(0, eGDOrderBarCfg_ButtonCode) = "Code"
        .TextMatrix(0, eGDOrderBarCfg_ButtonDescript) = "Button"
        
        For i = 0 To aButtons.Size - 1
            ParseOrdButtonString aButtons(i), nShow, strCode, strDescription
            If InStr(strCode, "BC") = 0 And InStr(strCode, "SC") = 0 And InStr(strCode, "PR") = 0 Then
                .Rows = .Rows + 1
                .Cell(flexcpChecked, .Rows - 1, eGDOrderBarCfg_ButtonShow) = nShow
                .TextMatrix(.Rows - 1, eGDOrderBarCfg_ButtonCode) = strCode
                .TextMatrix(.Rows - 1, eGDOrderBarCfg_ButtonDescript) = strDescription
            End If
        Next
        
        .Cell(flexcpPictureAlignment, .FixedRows, eGDOrderBarCfg_ButtonShow, .Rows - 1, eGDOrderBarCfg_ButtonShow) = flexAlignCenterCenter
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.InitOrderButtonsGrid"

End Sub

Private Sub RemoveChartControls(aCtrlCodes As cGdArray)
On Error GoTo ErrSection:

    Dim i&, iSize&, hArray&
    
    If aCtrlCodes Is Nothing Then Exit Sub
    
    hArray = aCtrlCodes.ArrayHandle
    iSize = aCtrlCodes.Size

    For i = iSize - 1 To 0 Step -1
        If InStr(gdGetStr(hArray, i), "BC") <> 0 Then
            gdDeleteItems hArray, i, 1          'buy chart
        ElseIf InStr(gdGetStr(hArray, i), "SC") <> 0 Then
            gdDeleteItems hArray, i, 1          'sell chart
        ElseIf InStr(gdGetStr(hArray, i), "PR") <> 0 Then
            gdDeleteItems hArray, i, 1          'prices on chart's order bar (market,bid,ask)
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartLadderCtrls.RemoveChartControls"

End Sub

Private Function ControlIndex(aCtrlCodes As cGdArray, ByVal strCtlName$) As Long
On Error GoTo ErrSection:

    Dim i&, iSize&, hArray&
    
    If aCtrlCodes Is Nothing Then Exit Function
    
    ControlIndex = -1
    
    hArray = aCtrlCodes.ArrayHandle
    iSize = aCtrlCodes.Size

    For i = 0 To iSize
        If InStr(gdGetStr(hArray, i), strCtlName) <> 0 Then
            ControlIndex = i
            Exit For
        End If
    Next

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartLadderCtrls.ControlIndex"

End Function

Public Function ConvertOrdBarButtons(ByVal strButtons$, Optional frm As Form) As String
On Error GoTo ErrSection:

    Dim strText$, strDefault$, strReturn$
    Dim i&, iButtons&, iPos&, txtLen&
    Dim aButtons As New cGdArray
    Dim aNewDefaults As New cGdArray
    
    Dim bOldFormat As Boolean
    Dim bLadder As Boolean
    
    strReturn = strButtons
    aButtons.SplitFields strButtons, "|"
    aNewDefaults.SplitFields kOrdBarDefaults, "|"
    
    If Not frm Is Nothing Then
        If TypeOf frm Is frmTickDistribution Then
            RemoveChartControls aButtons
            RemoveChartControls aNewDefaults
            bLadder = True
        End If
    End If
    
    iPos = InStr(strButtons, "TSO")
    If iPos > 0 Then
        If bLadder Then
            'remove from ladder order bar
            If InStr(strButtons, "|TSO;1") > 0 Then
                strReturn = Replace(strButtons, "|TSO;1", "")
            ElseIf InStr(strButtons, "|TSO;2") > 0 Then
                strReturn = Replace(strButtons, "|TSO;2", "")
            ElseIf InStr(strButtons, "|TSO;0") > 0 Then
                strReturn = Replace(strButtons, "|TSO;0", "")
            ElseIf InStr(strButtons, "|TSO") > 0 Then
                strReturn = Replace(strButtons, "|TSO", "")
            End If
        End If
        ConvertOrdBarButtons = strReturn
        GoTo ErrExit          'already in new format, no need to convert
    ElseIf Not bLadder Then
        'add to chart order bar
        If InStr(strButtons, "|AE;1") > 0 Then
            strReturn = Replace(strButtons, "|AE;1", "|AE;1|TSO;2")
        ElseIf InStr(strButtons, "|AE;2") > 0 Then
            strReturn = Replace(strButtons, "|AE;2", "|AE;2|TSO;2")
        ElseIf InStr(strButtons, "|AE;0") > 0 Then
            strReturn = Replace(strButtons, "|AE;0", "|AE;0|TSO;2")
        ElseIf InStr(strButtons, "|AE") > 0 Then
            strReturn = Replace(strButtons, "|AE", "|AE|TSO")
        End If
    End If
    
    iPos = InStr(strButtons, "CO")   'confirm order
    If iPos > 0 Then
        ConvertOrdBarButtons = strReturn
        GoTo ErrExit          'already in new format, no need to convert
    End If
    
    iPos = InStr(strButtons, "AJ")  'auto journal
    If iPos > 0 Then
        'Conversion History:
        '1. original did not have AJ button
        '2. AJ button added
        '3. BRORD button added
        '4. Confirm Order checkbox added (09-21-2010)
        
        'if already have AJ button then just add BRORD, CO buttons if needed
    Else
        iPos = InStr(strButtons, ";")
        If iPos > 0 Then
            If 0 = InStr(strButtons, "OE") Then aButtons.Add "OE;1", 2      'has format with checkbox flag, but does not have OE button
            strReturn = aButtons.JoinFields("|")
        Else
            'has original, old format
            strReturn = Replace(strReturn, "|", ";1|")
            strReturn = Replace(strReturn, "BR", "AE")
            bOldFormat = True
        End If
        
        i = -1
        If Not bOldFormat Then
            'has all buttons except AJ, locate AE button then add above
            i = ControlIndex(aButtons, "AE")
            If i >= 0 Then aButtons.Add "AJ;1", i
        End If
        
        If i = -1 Then
            aButtons.Size = 0
            For i = 0 To aNewDefaults.Size - 1
                strDefault = aNewDefaults(i)
                iPos = InStr(strReturn, strDefault)
                If iPos = 0 Then
                    If strDefault = "POS;1" Or strDefault = "OE;1" Or strDefault = "CLR;0" Or _
                       strDefault = "OT;1" Or strDefault = "AJ;1" Then
                        'do nothing, these are new items that must be added
                    ElseIf bOldFormat Then
                        'original format: 0=off, 1=on (no check box involved)
                        'new format: 0=no checkbox (user cannot turn on/off), 1=checkBox checked (on), 2=checkBox unchecked (off)
                        strDefault = Replace(strDefault, "1", "2")
                    End If
                    aButtons.Add strDefault
                Else
                    strText = Mid(strReturn, iPos, Len(strDefault))
                    aButtons.Add strText
                End If
            Next
        End If
    End If
    
    
    If bLadder Then
        'locate CLR button & add BRORD button before it
        iPos = InStr(strButtons, "BRORD")
        If iPos <= 0 Then
            i = -1
            i = ControlIndex(aButtons, "CLR")
            If i >= 0 Then aButtons.Add "BRORD;1", i
        End If
    
        'locate AJ button & add CO button before it
        iPos = InStr(strButtons, "CO")
        If iPos <= 0 Then
            i = -1
            i = ControlIndex(aButtons, "AJ")
            If i >= 0 Then aButtons.Add "CO;0", i
        End If
    Else
        'locate AJ button & add BRORD button before it
        iPos = InStr(strButtons, "BRORD")
        If iPos <= 0 Then
            i = -1
            i = ControlIndex(aButtons, "AJ")
            If i >= 0 Then aButtons.Add "BRORD;1", i
        End If
    
        'locate AJ button & add CO button before it
        iPos = InStr(strButtons, "CO")
        If iPos <= 0 Then
            i = -1
            i = ControlIndex(aButtons, "AJ")
            If i >= 0 Then aButtons.Add "CO;0", i
        End If
    End If
        
    If aNewDefaults.Size = aButtons.Size Then
        strReturn = aButtons.JoinFields("|")
    Else
        strReturn = kOrdBarDefaults      'something went wrong, just use new defaults
    End If
    
    ConvertOrdBarButtons = strReturn
    
ErrExit:
    Set aButtons = Nothing
    Set aNewDefaults = Nothing
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.ConvertOrdBarButtons"
    Resume ErrExit

End Function

Public Function ParseGridCtrl(fg As VSFlexGrid) As String
On Error GoTo ErrSection:
'Used by Price Ladder & Market Depth config forms for QuoteBar and AccountBar grids

    Dim i As Long
    Dim strText As String

    If Not fg Is Nothing Then
        With fg
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 0) = 1 Then
                    If Len(strText) > 0 Then
                        strText = strText & "|" & .TextMatrix(i, 1)
                    Else
                        strText = .TextMatrix(i, 1)
                    End If
                End If
            Next
        End With
    End If

    ParseGridCtrl = strText

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.ProcessQuoteBarGrid"
    
End Function

Public Function ParseOrderButtonsGrid(fg As VSFlexGrid) As String
On Error GoTo ErrSection:
'Used by Price Ladder & Market Depth config forms for order bar buttons grid

    Dim strSave$, strCode$, i&
    
    With fg
        For i = .FixedRows To .Rows - 1
            strCode = .TextMatrix(i, eGDOrderBarCfg_ButtonCode)
            If strCode = "CLR" Then
                strSave = strSave & "CLR;0|"
            Else
                strSave = strSave & strCode & ";"
                strSave = strSave & .Cell(flexcpChecked, i, eGDOrderBarCfg_ButtonShow) & "|"
            End If
        Next
    End With
    
    ParseOrderButtonsGrid = strSave
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.SaveOrderButtonsInfo"

End Function

Public Function OkayToExecute(Order As cPtOrder, ByVal dPrice#, _
    Optional ByVal bBracketOrder As Boolean = False, _
    Optional frm As Form = Nothing) As Boolean
On Error GoTo ErrSection:
'Note: dPrice passed in should be rounded to minimum move

    Dim bExecute As Boolean
    Dim oBracketOrderOne As cPtOrder
    Dim dBracketPrice#, dOrderPrice#, dCurrPrice#, strReturn$

    If Order Is Nothing Then Exit Function
    If bBracketOrder And frm Is Nothing Then Exit Function
    
    bExecute = True
    dOrderPrice = Order.OrderPrice(False)
    strReturn = "Y"
    
    Select Case Order.OrderType
        Case eTT_OrderType_Limit
            If g.Broker.WarnLimitWrongSide = True Then
                If Order.Buy Then
                    If dOrderPrice > dPrice Then
                        If bBracketOrder Then
                            strReturn = "N"
                            InfBox "The Limit Price for a Buy order should be less than the current market price.", "I", , "Order Confirmation"
                        Else
                            strReturn = InfBox("The Limit Price for a Buy order should be less than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                        End If
                    End If
                ElseIf dOrderPrice < dPrice Then
                    If bBracketOrder Then
                        strReturn = "N"
                        InfBox "The Limit Price for a Sell order should be greater than the current market price.", "I", , "Order Confirmation"
                    Else
                        strReturn = InfBox("The Limit Price for a Sell order should be greater than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                    End If
                End If
            End If
            
        Case eTT_OrderType_Stop
            If g.Broker.WarnStopWrongSide = True Then
                If Order.Buy Then
                    If dOrderPrice < dPrice Then
                        If bBracketOrder Then
                            strReturn = "N"
                            InfBox "The Stop Price for a Buy order should be greater than the current market price.", "I", , "Order Confirmation"
                        Else
                            strReturn = InfBox("The Stop Price for a Buy order should be greater than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                        End If
                    End If
                ElseIf dOrderPrice > dPrice Then
                    If bBracketOrder Then
                        strReturn = "N"
                        InfBox "The Stop Price for a Sell order should be less than the current market price.", "I", , "Order Confirmation"
                    Else
                        strReturn = InfBox("The Stop Price for a Sell order should be less than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                    End If
                End If
            End If
            
        Case eTT_OrderType_MIT
            If g.Broker.WarnStopWrongSide = True Then
                If Order.Buy Then
                    If dOrderPrice > dPrice Then
                        If bBracketOrder Then
                            strReturn = "N"
                            InfBox "The Market if Touched Price for a Buy order should be less than the current market price.", "I", , "Order Confirmation"
                        Else
                            strReturn = InfBox("The Market if Touched Price for a Buy order should be less than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                        End If
                    End If
                ElseIf dOrderPrice < dPrice Then
                    If bBracketOrder Then
                        strReturn = "N"
                        InfBox "The Market if Touched Price for a Sell order should be greater than the current market price.", "I", , "Order Confirmation"
                    Else
                        strReturn = InfBox("The Market if Touched Price for a Sell order should be greater than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                    End If
                End If
            End If
            
    End Select
    
    If strReturn = "N" Then
        bExecute = False
    ElseIf bBracketOrder Then
        Dim X&, Y&
        
        Set oBracketOrderOne = frm.BracketOrderOne
        If oBracketOrderOne Is Nothing Then
            'first side of bracket order - don't need to do anything
        Else
            dBracketPrice = oBracketOrderOne.OrderPrice(False)
            If (dBracketPrice > dPrice And dOrderPrice > dPrice) Or (dBracketPrice < dPrice And dOrderPrice < dPrice) Then
                If TypeOf frm Is frmTickDistribution Then
                    X = -1
                    Y = -1
                Else
                    frm.GetPromptLocation X, Y
                End If
                strReturn = InfBox("Bracket orders cannot be on same side of the market.", "I", "+Ok|-Cancel", "Bracket Order", , , , , , , , , , , X, Y)
                If strReturn = "C" Then
                    CancelOrder oBracketOrderOne
                    frm.ClearBuySellButtons True
                End If
                bExecute = False
            ElseIf oBracketOrderOne.Quantity <> Order.Quantity Then
                If TypeOf frm Is frmTickDistribution Then
                    X = -1
                    Y = -1
                Else
                    frm.GetPromptLocation X, Y
                End If
                strReturn = InfBox("Bracket orders must have the same quantiy. Please choose quantity to use.", "I", "+" & Str(oBracketOrderOne.Quantity) & _
                    "|" & Str(Order.Quantity) & "|Cancel", "Bracket Order", , , , , , , , , , , X, Y)
                If strReturn = "C" Then
                    CancelOrder oBracketOrderOne
                    frm.ClearBuySellButtons True
                    bExecute = False
                Else
                    oBracketOrderOne.Quantity = Int(ValOfText(strReturn))
                    Order.Quantity = Int(ValOfText(strReturn))
                End If
            End If
        End If
    End If
    
    OkayToExecute = bExecute

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.OkayToExecute"

End Function

Public Sub AttachChart(frm As Form)
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    
    Dim i&, iWindowState&, nShowTradesSave&, strPlacement$
    Dim bLocked As Boolean
    Dim bTimersSave As Boolean
    
    Dim frmNew As Form              'new form
    Dim ChartNew As cChart          'chart to use with new form
           
    If IsFrmChartMDI(frm) Then Exit Sub
    
    If bInProgress Then
        InfBox "Please wait until previous chart finish attaching then try again.", "I", , "Attach chart"
        Exit Sub
    End If

    bInProgress = True
    
    bTimersSave = ChartTimers
    ChartTimers = False
    
    If g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        iWindowState = vbMaximized
    Else
        iWindowState = g.ChartGlobals.frmActiveNonDetached.WindowState
    End If
        
    frm.Hide
    frm.SkipFocusFix = True
    frm.PfpReset ePfpReset_ClearAll
    frm.DetachStatus = eAttachInProg
    nShowTradesSave = frm.Chart.ShowTrades    '4980
    strPlacement = frm.GetNormalPlacement()
    
    'since the form will not be unloaded right away, do this here to free up some GDI resources now
    Set frm.pbTbBack(0).Picture = Nothing
    Set frm.pbTbBackDraw(0).Picture = Nothing
    Set frm.imgTbBack(0).Picture = Nothing
    Set frm.imgTbBackDraw(0).Picture = Nothing
    
    
    Set ChartNew = frm.Chart
    'get a new chart object for old form so old form can unload properly
    frm.ClearChartObject
    i = frm.Chart.ChartForeColor
        
    'set old form's timer tag so timer will unload old form
    'do not unload this form here because mouse-down event needs to finish
    frm.tmr.Tag = "UNLOAD_NOW"
    
    'since the chart that is getting attached is going away and was frmMain's active form
    'must activate last-known, last-active non-deatched form
    If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        ActiveChartFormSet g.ChartGlobals.frmActiveNonDetached      '5210
        SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_ACTIVATE, 1, 0
        DoEvents
    End If
    
    'initialize objects for new form
    Set frmNew = New frmChart                  'must be this type for a new non-detached chart

    frmNew.SetChartObject ChartNew              'do first so form_load does not exit
    frmNew.DetachStatus = eAttachInProg         'do this to trigger form_load & to prevent "events" from triggering
    frmNew.Caption = frmNew.Caption             'to force load if not loaded already
    
    frmNew.WindowLink.SymbolColor = frm.WindowLink.SymbolColor           '5132
    frmNew.WindowLink.PeriodColor = frm.WindowLink.PeriodColor
    
    frmNew.CopyPlacements frm
    frmNew.GameMode = frm.GameMode
    If frmNew.GameMode Is Nothing Then
        frmNew.Chart.ShowTrades = 0
    Else
        frmNew.GameMode.FormOwnerChange frmNew
        frmNew.IsInGameMode = frm.IsInGameMode
    End If
    
    If frm.OrderBarMode = eOrdBarMode_Order Then
        frmNew.InitQuantityEditor
        frmNew.txtTradeQty.Text = frm.txtTradeQty.Text      '6230
    End If
    
    Set frmNew.Chart.Form = frmNew
    Set frmNew.Chart.tbToolbar = frmMain.tbToolbar
    Set g.ChartGlobals.frmActiveNonDetached = frmNew
    
    If iWindowState = vbMaximized Then
        frmNew.WindowState = vbMaximized       '4882
    Else
        SetFormPlacement frmNew, strPlacement, "P"
    End If
    
    frmNew.DetachStatus = eNotDetached
    ActiveChartFormSet frmNew
    
    frmNew.Chart.ShowTrades = nShowTradesSave
    frmNew.Chart.SeasonalIndClear
    If frm.OrderBarMode = eOrdBarMode_Wizard Then frmNew.tmr.Tag = "ToggleOrderbarMode"       '4992
        
    If IsAtLeastVista Then
        bLocked = LockWindowUpdate(frmMain.hWnd)
    Else
        bLocked = LockWindowUpdate(GetDesktopWindow())
    End If
        
    ShowForm frmNew
    
    If bLocked Then LockWindowUpdate 0
    
    frmNew.SetAutoExit      '5251
    If Len(frmNew.Chart.SpreadSymbols) > 0 Then FormResize frmMain      '5244

ErrExit:
    ChartTimers = bTimersSave
    bInProgress = False
    Exit Sub

ErrSection:
    ChartTimers = bTimersSave
    bInProgress = False
    RaiseError "mChartLadderCtrls.AttachChart"

End Sub

Public Sub DetachChart(frm As Form)
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    
    Dim i&, iWindowState&, nShowTradesSave&, strPlacement$
    Dim bLocked As Boolean
    Dim bTimersSave As Boolean
    Dim bMaximizeNewForm As Boolean
    
    Dim frmNew As Form              'new form
    Dim ChartNew As cChart          'chart to use with new form
        
    If Not IsFrmChartMDI(frm) Then Exit Sub
    
    If g.ChartGlobals.nDetached = 0 Then        'this count is > 0 during chart page restore
        If bInProgress Then
            InfBox "Please wait until previous chart finish detaching then try again.", "I", , "Detach chart"
            Exit Sub
'        ElseIf NonDetachCount <= 1 Then
'            InfBox "You cannot detach all charts.", "I", , "Detach chart"           '5003
'            Exit Sub
        End If
    ElseIf FormIsLoaded("frmPositionConfirm") Then
        Exit Sub        '5120
    End If
    
    bInProgress = True

    bTimersSave = ChartTimers
    ChartTimers = False
    
    iWindowState = frm.WindowState
            
    frm.Hide
    frm.SkipFocusFix = True
    frm.PfpReset ePfpReset_ClearAll
    frm.DetachStatus = eDetachInProg
    nShowTradesSave = frm.Chart.ShowTrades    '4980
    strPlacement = frm.GetDetachedPlacement()
    
    'since the form will not be unloaded right away, do this here to free up some GDI resources now
    Set frm.pbTbBack(0).Picture = Nothing
    Set frm.pbTbBackDraw(0).Picture = Nothing
    Set frm.imgTbBack(0).Picture = Nothing
    Set frm.imgTbBackDraw(0).Picture = Nothing
        
    Set ChartNew = frm.Chart
    'get a new chart object for old form so old form can unload properly
    frm.ClearChartObject
    i = frm.Chart.ChartForeColor
    
    If frm.tmr.Tag = "DETACH_NOW" And Len(frm.Tag) > 0 Then
        If Val(frm.Tag) = vbMaximized Then bMaximizeNewForm = True
        frm.Tag = ""
    ElseIf frm.tmr.Tag <> "DETACH_NOW" Then
        'since old chart is getting detached & unloaded, must reset non-detached active form pointer
        UpdateNonDetached frm, iWindowState
    End If
        
    'set old form's timer tag so timer will unload old form
    'do not unload this form here because mouse-down event needs to finish
    frm.tmr.Tag = "UNLOAD_NOW"
    
    'initialize objects for new form
    Set frmNew = New frmChart2                  'must be this type for a new detached chart

    frmNew.SetChartObject ChartNew              'do first so form_load does not exit
    frmNew.DetachStatus = eDetachInProg         'do this to trigger form_load & to prevent "events" from triggering
    frmNew.Caption = ""
    
    frmNew.WindowLink.SymbolColor = frm.WindowLink.SymbolColor           '5132
    frmNew.WindowLink.PeriodColor = frm.WindowLink.PeriodColor

    frmNew.tmr.Enabled = False
    frmNew.CopyPlacements frm
    frmNew.GameMode = frm.GameMode
    If frmNew.GameMode Is Nothing Then
        frmNew.Chart.ShowTrades = 0
    Else
        frmNew.GameMode.FormOwnerChange frmNew
        frmNew.IsInGameMode = frm.IsInGameMode
    End If
    
    If frm.OrderBarMode = eOrdBarMode_Order Then
        frmNew.InitQuantityEditor
        frmNew.txtTradeQty.Text = frm.txtTradeQty.Text      '6230
    End If
    
    Set frmNew.Chart.Form = frmNew
    If frmNew.Chart.ShowToolbar = 1 Then
        ToolbarInit2 frmNew, frmNew.TbButtonsArray(kTbGeneral), , kTbGeneral
        ToolbarInit2 frmNew, frmNew.TbButtonsArray(kTbDraw), , kTbDraw, , g.vbeTbAlignDraw
        If g.vbeTbAlignDraw = vbAlignTop Or g.vbeTbAlignDraw = vbAlignBottom Then
            frmNew.pbTbBackDraw(0).align = vbAlignTop
        Else
            frmNew.pbTbBackDraw(0).align = vbAlignRight
        End If
    End If
            
    If Len(strPlacement) > 0 Then
        Dim l&, t&, w&, h&
        l = Parse(strPlacement, ";", 1)
        t = Parse(strPlacement, ";", 2)
        i = geIsPointOnScreen(l / Screen.TwipsPerPixelX, t / Screen.TwipsPerPixelY)
        If i = 1 Then
            'SetFormPlacement Me, m.strDetachedPlacment, "P"
            w = Parse(strPlacement, ";", 3)
            h = Parse(strPlacement, ";", 4)
            frmNew.Move l, t, w, h
        Else
            frmNew.Move 10, 10, g.ChartGlobals.frmActiveNonDetached.Width, g.ChartGlobals.frmActiveNonDetached.Height
        End If
    Else
        frmNew.Move 10, 10, frm.Width, frm.Height
    End If
            
    If iWindowState = vbMaximized Then
        'when the detachment is complete, the system sets all child forms
        'within the main app to window state vbNormal for whatever reason
        'this code resets the the MDI child forms to vbMaximized if they were
        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then g.ChartGlobals.frmActiveNonDetached.WindowState = vbMaximized
    End If
    
    ActiveChartFormSet frmNew                  '4882
    
    frmNew.DetachStatus = eDetached
    frmNew.Chart.ShowTrades = nShowTradesSave
    frmNew.Chart.SeasonalIndClear
    If frm.OrderBarMode = eOrdBarMode_Wizard Then frmNew.tmr.Tag = "ToggleOrderbarMode"       '4992
        
    If IsAtLeastVista Then
        bLocked = LockWindowUpdate(frmMain.hWnd)
    Else
        bLocked = LockWindowUpdate(GetDesktopWindow())
    End If
    
    If frmNew.Chart.ShowToolbar <> 0 Then
        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
            If Not g.ChartGlobals.frmActiveNonDetached.Chart Is Nothing Then
                g.ChartGlobals.frmActiveNonDetached.Chart.SyncToolbar True
            End If
        End If
    End If
    
    If bMaximizeNewForm Then frmNew.WindowState = vbMaximized       '5516
    ShowForm frmNew, eForm_Nonmodal, frmMain
    
    If iWindowState = vbMaximized Then
        FormResize frmMain      'to get system menu back to MDI child charts
    End If
    
    If bLocked Then LockWindowUpdate 0
    
    frmNew.SetAutoExit      '5251
    
    FormResize frmNew
    
    If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        g.ChartGlobals.frmActiveNonDetached.SetChartTabs        '7004 - must do AFTER window update is unlocked
    End If
    
    
ErrExit:
    bInProgress = False
    ChartTimers = bTimersSave
    Exit Sub

ErrSection:
    bInProgress = False
    ChartTimers = bTimersSave
    RaiseError "mChartLadderCtrls.DetachChart"

End Sub

Private Sub UpdateNonDetached(frmPrev As Form, ByVal iWindowState&)
On Error GoTo ErrSection:

    Dim i&
    Dim frmActive As Form
    
    Set frmActive = frmMain.ActiveForm
    If Not frmActive Is Nothing Then
        If IsFrmChartMDI(frmActive) And Not (frmActive Is frmPrev) Then
            frmActive.WindowState = iWindowState
            Set g.ChartGlobals.frmActiveNonDetached = frmActive
            If iWindowState = vbMaximized Then FormResize frmActive
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.UpdateNonDetached"

End Sub

Public Function NonDetachCount() As Long
On Error GoTo ErrSection:

    Dim i&, iCount&
    
    For i = 0 To Forms.Count - 1
        If IsFrmChartMDI(Forms(i)) Then
            If Forms(i).tmr.Tag <> "UNLOAD_NOW" And Forms(i).tmr.Tag <> "UNLOADING" Then
                iCount = iCount + 1
            End If
        End If
    Next
    
    NonDetachCount = iCount
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.NonDetachCount"

End Function

Public Function WizardGridLegInfo(frm As Form, ByVal bIncludeOrdType As Boolean) As String
On Error GoTo ErrSection:

    Dim i&, iQty&, iBuyColor&, dPrice#
    Dim strBaseSym$, strSymbol$, strOptSym$, strQty$, strPrice$, strOrderType$
    Dim strText$, strMsg$
    Dim bBaseSym As Boolean
    
    Dim Chart As cChart

    If frm Is Nothing Then Exit Function
    If Not IsFrmChart(frm) Then Exit Function
    
    Set Chart = frm.Chart
    If Chart Is Nothing Then Exit Function

    strBaseSym = Parse(Chart.Symbol, "-", 1)
    strSymbol = RollSymbolForDate(Chart.Symbol, Chart.Bars(eBARS_DateTime, Chart.Bars.Size - 1))
    iBuyColor = frm.vseBuyWizard.BackColor
    
    With frm.fgChartFlex(eFlexGridIdx_OrdWizard)
    
        For i = .FixedRows To .Rows - 1
            strText = .TextMatrix(i, 1)
            iQty = ValOfText(.TextMatrix(i, 2))
            strOptSym = .TextMatrix(i, 3)
            If strOptSym = strBaseSym Then
                strOptSym = strSymbol
                bBaseSym = True
            Else
                bBaseSym = False
            End If
            
            If Len(strText) > 0 Then
                strOrderType = ""
                If bBaseSym Then
                    If bIncludeOrdType Then strOrderType = ";" & .TextMatrix(i, 4)
                    strPrice = Parse(strText, " ", 2)
                    
                    'If InStr(strPrice, "^") <> 0 Then
                    'JM 06-08-2009: For now, always do this so price will go to Options Navigator without regional settings
                    dPrice = Chart.Bars.PriceFromString(strPrice)     'Bernie does not want ^ in price string
                    strPrice = Str(dPrice)
                Else
                    If bIncludeOrdType Then strOrderType = ";L"
                    strPrice = "0"
                End If
                
                If .Cell(flexcpBackColor, i, 1) = iBuyColor Then
                    strQty = Str(iQty)
                Else
                    strQty = Str(-iQty)
                End If
                
                If Len(strMsg) = 0 Then
                    If bBaseSym Then
                        strMsg = strSymbol & ";" & strQty & ";" & strPrice & strOrderType
                    ElseIf bIncludeOrdType Then
                        strMsg = strSymbol & ";0;0;L" & vbTab & strOptSym & ";" & strQty & ";" & strPrice & strOrderType
                    Else
                        strMsg = strSymbol & ";0;0" & vbTab & strOptSym & ";" & strQty & ";" & strPrice & strOrderType
                    End If
                Else
                    strMsg = strMsg & vbTab & strOptSym & ";" & strQty & ";" & strPrice & strOrderType
                End If
            End If
        Next
    
    End With
    
    WizardGridLegInfo = strMsg

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.WizardLegInfo"

End Function

Public Sub SetNextChartActive(frmClosing As Form, ByVal strSymbol$)
On Error GoTo ErrSection:

    Dim i&
    Dim frm As Form
    Dim frmNonDetached As frmChart
    Dim frmDetached As frmChart2
      
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    
    If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        If frmClosing.DetachStatus = eDetached Then
            If frmClosing Is ActiveChart Then           '6201
                If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                    ActiveChartFormSet g.ChartGlobals.frmActiveNonDetached
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        ElseIf Not frmClosing Is g.ChartGlobals.frmActiveNonDetached Then
            Exit Sub
        End If
    End If
    
    For i = 0 To Forms.Count - 1
        Set frm = Forms(i)
        If Not frm Is frmClosing Then
            If IsFrmChart(frm) Then
                If TypeOf frm Is frmChart Then
                    If Len(frm.Tag) = 0 And Len(frm.tmr.Tag) = 0 And frm.DetachStatus = eNotDetached Then
                        Set frmNonDetached = frm
                        Exit For
                    End If
                ElseIf frmNonDetached Is Nothing Then
                    If Len(frm.Tag) = 0 And Len(frm.tmr.Tag) = 0 And frm.DetachStatus = eDetached Then
                        Set frmDetached = frm
                    End If
                End If
            End If
        End If
    Next
    
    If Not frmNonDetached Is Nothing Then
        Set frm = Screen.ActiveForm
        If frm Is Nothing Then
            frmNonDetached.SkipFocusFix = True
            ActiveChartFormSet frmNonDetached
        ElseIf IsFrmChart(frm) Then
            If InStr(frm.tmr.Tag, "UNLOAD") = 0 Then        'check for "UNLOAD", "UNLOAD NOW" & "UNLOADING" tag
                Set g.ChartGlobals.frmActiveNonDetached = frmNonDetached
                'a just detached chart has the focus; return focus to it
                MoveFocus frm.pbChart
            Else
                frmNonDetached.SkipFocusFix = True
                ActiveChartFormSet frmNonDetached
                MoveFocus frmNonDetached.pbChart
            End If
        Else
            frmNonDetached.SkipFocusFix = True
            ActiveChartFormSet frmNonDetached
            MoveFocus frmNonDetached.pbChart
        End If
    ElseIf Not frmDetached Is Nothing Then
        frmDetached.SkipFocusFix = True
        Set g.ChartGlobals.frmActiveNonDetached = Nothing
        ActiveChartFormSet frmDetached
        MoveFocus frmDetached.pbChart
    Else
        Set g.ChartGlobals.frmActiveNonDetached = Nothing
        ActiveChartFormSet Nothing
    
'JM 08-22-2011: this seens to be a work around for issue 5206 rather than a real fix
'   don't think it is needed anymore, leave awhile then remove if all ok.
'        InfBox "There must be at least one chart open.", "I", , "Chart Information"      '5206
'
'        Set frm = New frmChart        'new chart is always non-detached
'        If Len(strSymbol) > 0 Then
'            frm.Chart.SetSymbol strSymbol
'        Else
'            frm.Chart.SetSymbol g.SymbolPool.SymbolIDforSymbol("$DJIA")
'        End If
'        frm.WindowState = vbMaximized
'        frm.Show
    End If
    

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.SetNextChartActive"

End Sub


Public Function GetSymbolPitType(ByVal strSymbol$) As eFutureSymbolType
On Error GoTo ErrSection:

    Dim strSymConverted$
    Dim eType As eFutureSymbolType
    
    eType = ePrimarySymbol
    
    If SecurityType(strSymbol) <> "F" Then GoTo ErrExit    '6171
    
    strSymConverted = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
    If strSymbol = strSymConverted Then
        eType = eElectronicSymbol
    Else
        strSymConverted = ConvertFutureSymbol(strSymbol, ePitSymbol)
        If strSymbol = strSymConverted Then
            eType = ePitSymbol
        Else
            strSymConverted = ConvertFutureSymbol(strSymbol, eCombinedSymbol)
            If strSymbol = strSymConverted Then
                eType = eCombinedSymbol
            Else
                strSymConverted = ConvertFutureSymbol(strSymbol, eSyntheticSymbol)
                If strSymbol = strSymConverted Then eType = eSyntheticSymbol
            End If
        End If
    End If

ErrExit:
    GetSymbolPitType = eType
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.GetSymbolPitType"

End Function

Public Sub CheckOpenOrderPos(frm As Form)
On Error Resume Next:

    Static bHasOrderPos As Boolean
    
    Dim bSyncBtn As Boolean
    
    If g.bStarting Or g.bLoadingChartPage Or g.bUnloading Then Exit Sub
    
    ' 03/17/2010 DAJ: Use this call now to see if there are any visible working orders or
    ' open positions...
    bSyncBtn = (g.ConsoleForms.NumVisible(eGDConsoleForm_Summary) > 0)
        
    If bHasOrderPos <> bSyncBtn Then
        bHasOrderPos = bSyncBtn
        SyncTradeTrackerBtn frmMain
    End If
    
End Sub

Public Sub LoadAnnotPenstyle(cbo As ctlUniComboImageXP)
On Error GoTo ErrSection:
    
    With cbo
        .AddItem "Default"
        .ItemData(.ListCount - 1) = eANNOT_Default
        .AddItem "Thin"
        .ItemData(.ListCount - 1) = eANNOT_Thin
        .AddItem "Medium"
        .ItemData(.ListCount - 1) = eANNOT_Medium
        .AddItem "Thick"
        .ItemData(.ListCount - 1) = eANNOT_Thick
        .AddItem "Dashed (Large)"
        .ItemData(.ListCount - 1) = eANNOT_DashLg
        .AddItem "Dashed (Small)"
        .ItemData(.ListCount - 1) = eANNOT_DashSm
        .AddItem "Dash Dot"
        .ItemData(.ListCount - 1) = eANNOT_DashDot
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mChartLadderCtrls.LoadAnnotPenstyle", eGDRaiseError_Raise
    
End Sub

Public Sub SetAnnotPenstyleCombo(cbo As ctlUniComboImageXP, nValue&)
On Error GoTo ErrSection:

    Dim i&, nMatch&

    For i = 0 To cbo.ListCount - 1
        If nValue = cbo.ItemData(i) Then
            nMatch = i
            Exit For
        End If
    Next
    If nMatch >= 0 And nMatch < cbo.ListCount Then
        cbo.ListIndex = nMatch
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.SetAnnotPenstyleCombo", eGDRaiseError_Raise

End Sub

Public Function ShowRithmic(ByVal TradeAccountID&) As Boolean
On Error GoTo ErrSection:

    ShowRithmic = g.Broker.IsRithmicBroker(g.Broker.AccountTypeForID(TradeAccountID)) And (g.Broker.ConnectionStatusForAccount(TradeAccountID) = eGDConnectionStatus_Connected)


ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.ShowRithmic"

End Function

Public Sub ExitFavoriteBtnClick(frm As Form, ctrl As vsElastic, nButton As Integer, _
    ByVal strFavorite$, ByVal globalArrayIdx&)
On Error GoTo ErrSection:

    Dim strBaseSym$, strStragegy$
    
    Dim oExit As cExitStrategy
    Dim oSymExits As cSymExitFavorites
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    
    If frm Is Nothing Then Exit Sub
    If ctrl Is Nothing Then Exit Sub
    If globalArrayIdx < 0 Or globalArrayIdx > 3 Then Exit Sub
    If strFavorite <> "A" And strFavorite <> "B" And strFavorite <> "C" And strFavorite <> "D" Then Exit Sub

    If g.ChartGlobals.eAEFMode = eAEFMode_Symbol Then
        strBaseSym = ExitFavoritesBaseSym(frm)
        Set oSymExits = g.ChartGlobals.treeExitFavorites.Item(strBaseSym)
    End If

    If nButton = vbRightButton Then
        If ctrl.Appearance = apInset Then frm.chkAutoExit.Value = vbUnchecked
    Else
        ExitCtrlAppearance frm, ctrl, ""
        ExitFavoriteSelect frm, strFavorite
        
        If oSymExits Is Nothing Then
            strStragegy = g.ChartGlobals.aExitFavorites(globalArrayIdx).StrategyName
        ElseIf Not oSymExits.ExitObjectGet(strBaseSym, strFavorite) Is Nothing Then
            strStragegy = oSymExits.ExitObjectGet(strBaseSym, strFavorite).StrategyName
        End If
    
        'double check that favorite was successfully activated
        If frm.lblAutoExit.Caption <> strStragegy Then
            ExitCtrlAppearance frm, Nothing, ""
        End If
    End If

ErrExit:
    frm.chkAutoExit.Enabled = True
    frm.lblAutoExit.Enabled = True
    Exit Sub

ErrSection:
    frm.chkAutoExit.Enabled = True
    frm.lblAutoExit.Enabled = True
    RaiseError "mChartLadderCtrls.ExitFavoriteBtnClick"

End Sub

Private Sub ExitFavoriteSelect(frm As Form, ByVal strFavorite$)
On Error GoTo ErrSection:

    Dim nSymID&, nAccountID&, strBaseSym$
    
    Dim AutoExit As cExitStrategy
    Dim oSymExits As cSymExitFavorites
    Dim strSource As String

    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If frm Is Nothing Then Exit Sub
    
    If IsFrmChart(frm) Then
        nAccountID = frm.Chart.TradeAccountID
        nSymID = frm.Chart.TradeSymbolID
        strSource = "Chart"
    ElseIf TypeOf frm Is frmTickDistribution Then
        nAccountID = frm.TradeAccountID
        nSymID = frm.SymbolID
        strSource = "Price Ladder"
    Else
        Exit Sub
    End If
    
    If g.ChartGlobals.eAEFMode = eAEFMode_Symbol Then
        strBaseSym = ExitFavoritesBaseSym(frm)
        Set oSymExits = g.ChartGlobals.treeExitFavorites.Item(strBaseSym)
    End If
    
    Select Case strFavorite
        Case "A"
            If oSymExits Is Nothing Then
                Set AutoExit = g.ChartGlobals.aExitFavorites(0)
            Else
                Set AutoExit = oSymExits.ExitObjectGet(strBaseSym, "A")
            End If
        Case "B"
            If oSymExits Is Nothing Then
                Set AutoExit = g.ChartGlobals.aExitFavorites(1)
            Else
                Set AutoExit = oSymExits.ExitObjectGet(strBaseSym, "B")
            End If
        Case "C"
            If oSymExits Is Nothing Then
                Set AutoExit = g.ChartGlobals.aExitFavorites(2)
            Else
                Set AutoExit = oSymExits.ExitObjectGet(strBaseSym, "C")
            End If
        Case "D"
            If oSymExits Is Nothing Then
                Set AutoExit = g.ChartGlobals.aExitFavorites(3)
            Else
                Set AutoExit = oSymExits.ExitObjectGet(strBaseSym, "D")
            End If
    End Select
    
    'precautionary checks, theoretically should be valid
    If AutoExit Is Nothing Then Exit Sub
    If Len(AutoExit.FileName) = 0 Then Exit Sub
    
    'do this to give user visual feedback that something is happening and to prevent the
    'form from turning off the AppInset appearance of button (see SetAutoExit in forms)
    frm.chkAutoExit.Enabled = False
    frm.lblAutoExit.Enabled = False
    
    If frm.chkAutoExit.Value = vbChecked Then
        If AutoExit.StrategyName = frm.lblAutoExit.Caption Then GoTo ErrExit
    End If
    
    If CanActivateAutomatedItem(nAccountID, nSymID, "Auto Exit", strSource) Then
        SelectXOS nAccountID, nSymID, AutoExit.FileName
    End If

ErrExit:
    frm.chkAutoExit.Enabled = True
    frm.lblAutoExit.Enabled = True
    Exit Sub

ErrSection:
    frm.chkAutoExit.Enabled = True
    frm.lblAutoExit.Enabled = True
    RaiseError "mChartLadderCtrls.ExitFavoriteSelect"

End Sub

Public Sub ExitCtrlAppearance(frm As Form, SelectedCtrl As vsElastic, ByVal strExitName$, _
    Optional ByVal bSync As Boolean = False)
On Error GoTo ErrSection:

    Dim strCaption$, strBaseSym$
    
    Dim strA$, strB$, strC$, strD$
    
    Dim AutoExit As cExitStrategy
    Dim oSymExits As cSymExitFavorites
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If frm Is Nothing Then Exit Sub
    
    If g.ChartGlobals.eAEFMode = eAEFMode_Symbol Then
        strBaseSym = ExitFavoritesBaseSym(frm)
        Set oSymExits = g.ChartGlobals.treeExitFavorites.Item(strBaseSym)
    End If
    
    If oSymExits Is Nothing Then
        strA = g.ChartGlobals.aExitFavorites(0).StrategyName
        strB = g.ChartGlobals.aExitFavorites(1).StrategyName
        strC = g.ChartGlobals.aExitFavorites(2).StrategyName
        strD = g.ChartGlobals.aExitFavorites(3).StrategyName
    Else
        If Not oSymExits.ExitObjectGet(strBaseSym, "A") Is Nothing Then
            strA = oSymExits.ExitObjectGet(strBaseSym, "A").StrategyName
        End If
        If Not oSymExits.ExitObjectGet(strBaseSym, "B") Is Nothing Then
            strB = oSymExits.ExitObjectGet(strBaseSym, "B").StrategyName
        End If
        If Not oSymExits.ExitObjectGet(strBaseSym, "C") Is Nothing Then
            strC = oSymExits.ExitObjectGet(strBaseSym, "C").StrategyName
        End If
        If Not oSymExits.ExitObjectGet(strBaseSym, "D") Is Nothing Then
            strD = oSymExits.ExitObjectGet(strBaseSym, "D").StrategyName
        End If
    End If
    
    'set tooltip
    frm.vseExitA.ToolTipText = strA
    frm.vseExitB.ToolTipText = strB
    frm.vseExitC.ToolTipText = strC
    frm.vseExitD.ToolTipText = strD
    
    'set appearance of favorite buttons
    strCaption = frm.lblAutoExit.Caption
    If strCaption = "None" Then
        If frm.vseExitA.Appearance <> ap3D Then frm.vseExitA.Appearance = ap3D
        If frm.vseExitB.Appearance <> ap3D Then frm.vseExitB.Appearance = ap3D
        If frm.vseExitC.Appearance <> ap3D Then frm.vseExitC.Appearance = ap3D
        If frm.vseExitD.Appearance <> ap3D Then frm.vseExitD.Appearance = ap3D
    ElseIf bSync Or Len(strExitName) > 0 Then
        'see if any button has matching strategy as autoexit label
        If strCaption <> "None" Then
            If strA = strCaption Then
                If frm.vseExitA.Appearance <> apInset Then frm.vseExitA.Appearance = apInset
            Else
                If frm.vseExitA.Appearance <> ap3D Then frm.vseExitA.Appearance = ap3D
            End If
            
            If strB = strCaption Then
                If frm.vseExitB.Appearance <> apInset Then frm.vseExitB.Appearance = apInset
            Else
                If frm.vseExitB.Appearance <> ap3D Then frm.vseExitB.Appearance = ap3D
            End If
            
            If strC = strCaption Then
                If frm.vseExitC.Appearance <> apInset Then frm.vseExitC.Appearance = apInset
            Else
                If frm.vseExitC.Appearance <> ap3D Then frm.vseExitC.Appearance = ap3D
            End If
            
            If strD = strCaption Then
                If frm.vseExitD.Appearance <> apInset Then frm.vseExitD.Appearance = apInset
            Else
                If frm.vseExitD.Appearance <> ap3D Then frm.vseExitD.Appearance = ap3D
            End If
        End If
    
    ElseIf Not SelectedCtrl Is Nothing Then
    
        If frm.vseExitA Is SelectedCtrl Then
            If frm.vseExitA.Appearance <> apInset Then frm.vseExitA.Appearance = apInset
        ElseIf frm.vseExitA.Appearance <> ap3D Then
            frm.vseExitA.Appearance = ap3D
        End If
        
        If frm.vseExitB Is SelectedCtrl Then
            If frm.vseExitB.Appearance <> apInset Then frm.vseExitB.Appearance = apInset
        ElseIf frm.vseExitB.Appearance <> ap3D Then
            frm.vseExitB.Appearance = ap3D
        End If
    
        If frm.vseExitC Is SelectedCtrl Then
            If frm.vseExitC.Appearance <> apInset Then frm.vseExitC.Appearance = apInset
        ElseIf frm.vseExitC.Appearance <> ap3D Then
            frm.vseExitC.Appearance = ap3D
        End If
    
        If frm.vseExitD Is SelectedCtrl Then
            If frm.vseExitD.Appearance <> apInset Then frm.vseExitD.Appearance = apInset
        ElseIf frm.vseExitD.Appearance <> ap3D Then
            frm.vseExitD.Appearance = ap3D
        End If
    End If
    
    If TypeOf frm Is frmTickDistribution Then
        strCaption = frm.lblTradePos.Caption
    Else
        strCaption = frm.lblTradePos.Text
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.ExitCtrlAppearance"

End Sub

Public Sub ExitFavoritesCheck(frm As Form, ByVal strPos$)
On Error Resume Next

    Dim i&, strBaseSym$
    
    Dim oExit As cExitStrategy
    Dim oSymExits As cSymExitFavorites
    
    Dim bContinue As Boolean
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If frm Is Nothing Then Exit Sub
    
    If FormIsLoaded("frmOrderStrategies") Then
        'don't want to check form's visible property first cause that will cause form to load if it is not loaded
        bContinue = Not frmOrderStrategies.Visible
    Else
        bContinue = True
    End If
    
    'check exit favorite buttons in case trade position changed
    If frm.fraExitFavorites.Visible Then
        If bContinue Then
            If strPos = "Flat" Then
                If g.ChartGlobals.eAEFMode = eAEFMode_Symbol Then
                    strBaseSym = ExitFavoritesBaseSym(frm)
                    Set oSymExits = g.ChartGlobals.treeExitFavorites.Item(strBaseSym)
                End If
                
                'A button
                If oSymExits Is Nothing Then
                    i = Len(g.ChartGlobals.aExitFavorites(0).FileName)
                Else
                    Set oExit = oSymExits.ExitObjectGet(strBaseSym, "A")
                    If oExit Is Nothing Then
                        i = 0
                    Else
                        i = Len(oExit.FileName)
                    End If
                End If
                If i = 0 And frm.vseExitA.Enabled Then
                    frm.vseExitA.Enabled = False
                ElseIf i > 0 And Not frm.vseExitA.Enabled Then
                    frm.vseExitA.Enabled = True
                End If
                
                'B button
                If oSymExits Is Nothing Then
                    i = Len(g.ChartGlobals.aExitFavorites(1).FileName)
                Else
                    Set oExit = oSymExits.ExitObjectGet(strBaseSym, "B")
                    If oExit Is Nothing Then
                        i = 0
                    Else
                        i = Len(oExit.FileName)
                    End If
                End If
                If i = 0 And frm.vseExitB.Enabled Then
                    frm.vseExitB.Enabled = False
                ElseIf i > 0 And Not frm.vseExitB.Enabled Then
                    frm.vseExitB.Enabled = True
                End If
                
                'C button
                If oSymExits Is Nothing Then
                    i = Len(g.ChartGlobals.aExitFavorites(2).FileName)
                Else
                    Set oExit = oSymExits.ExitObjectGet(strBaseSym, "C")
                    If oExit Is Nothing Then
                        i = 0
                    Else
                        i = Len(oExit.FileName)
                    End If
                End If
                If i = 0 And frm.vseExitC.Enabled Then
                    frm.vseExitC.Enabled = False
                ElseIf i > 0 And Not frm.vseExitC.Enabled Then
                    frm.vseExitC.Enabled = True
                End If
                
                'D button
                If oSymExits Is Nothing Then
                    i = Len(g.ChartGlobals.aExitFavorites(3).FileName)
                Else
                    Set oExit = oSymExits.ExitObjectGet(strBaseSym, "D")
                    If oExit Is Nothing Then
                        i = 0
                    Else
                        i = Len(oExit.FileName)
                    End If
                End If
                If i = 0 And frm.vseExitD.Enabled Then
                    frm.vseExitD.Enabled = False
                ElseIf i > 0 And Not frm.vseExitD.Enabled Then
                    frm.vseExitD.Enabled = True
                End If
            Else
                If frm.vseExitA.Enabled Then frm.vseExitA.Enabled = False
                If frm.vseExitB.Enabled Then frm.vseExitB.Enabled = False
                If frm.vseExitC.Enabled Then frm.vseExitC.Enabled = False
                If frm.vseExitD.Enabled Then frm.vseExitD.Enabled = False
            End If
        End If
    End If

End Sub

' If memory is too low, drop chart pages from cache
Public Sub CheckPageCacheSize()

    On Error Resume Next
    
    If g.ChartPageCache Is Nothing Then Exit Sub
    
    ' until "Available RAM" (i.e. TotalPhysical - Committed) is at least 200 megs
    Do While g.ChartPageCache.Count > 10 Or PhysicalRAM(True) < 400
        ' but only if we have a page to remove
        If g.ChartPageCache.Count = 0 Then
            Exit Do         'JM 10-05-2011: without this code, users were caught in infinite loop that caused chart page change failure (this is the infbox never went away bug)
        End If
        g.ChartPageCache.Remove 1
        DoEvents
    Loop

End Sub

Public Function CachePageSave(ByVal strPageName$) As cGdArray
On Error GoTo ErrSection:

    Dim i&
    Dim bMax As Boolean

    Dim frm As Form
    Dim Page As cPageCache
    Dim Charts As cGdTree
'    Dim aChartForms As New cGdArray

    If g.ChartPageCache Is Nothing Then Set g.ChartPageCache = New cGdTree

    CheckPageCacheSize
    
    Set Page = New cPageCache
    Page.PageName = strPageName
    
    Page.ActiveIdxReset
    
    If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        If g.ChartGlobals.frmActiveNonDetached.WindowState = vbMaximized Then
            bMax = True
        End If
    End If
    
    'walk through forms and
    '1. add chart objects to new page cache
    '2. remove detached & minimized charts from forms collection since they cannot be reused
    '3. return array of non-detached forms for reuse
    
'    Dim bProfileChart As Boolean        'JM 04-30-2013: for now, don't cache profile charts (issues on restore)
    
    'For i = Forms.Count - 1 To 0 Step -1
        'Set frm = Forms(i)
    Set Charts = GetChartsInZorder
    For i = Charts.Count To 1 Step -1
        Set frm = Charts(i)
        If TypeOf frm Is frmChart Then
'            If Not bProfileChart Then
'                bProfileChart = frm.Chart.IsProfileChart
'                If Not bProfileChart Then Page.PageCacheAdd frm.Chart
'            End If
            Page.PageCacheAdd frm.Chart
            If frm.Tag = "BOGUS" Then
                Unload frm
                Set frm = Nothing
                'Unload Forms(i)
            ElseIf frm.WindowState = vbMinimized Then
                Unload frm
                Set frm = Nothing
                'Unload Forms(i)             '6659, 6662
            Else
                frm.SetNormalPlacement ""
'                aChartForms.Add frm
            End If
        ElseIf TypeOf frm Is frmChart2 Then
'            If Not bProfileChart Then
'                bProfileChart = frm.Chart.IsProfileChart
'                If Not bProfileChart Then Page.PageCacheAdd frm.Chart
'            End If
            Page.PageCacheAdd frm.Chart
            Unload frm
            Set frm = Nothing
            'Unload Forms(i)
        End If
    Next
    Set Charts = Nothing
    
'    If Not bProfileChart Then
'        'move the active chart to end so will become the active chart when restored
'        Page.MoveActiveToEnd bMax
'
'        g.ChartPageCache.Add Page, strPageName
'    End If


    'move the active chart to end so will become the active chart when restored
    Page.MoveActiveToEnd bMax
    g.ChartPageCache.Add Page, strPageName
    
    CheckPageCacheSize
    
    If g.ChartPageCache.Count = 0 Then
        Dim dRam#
        dRam = PhysicalRAM(True)
        
        DebugLog "Cannot cache chart page. Low memory: " & Str(dRam)
    End If
    
    'JM 07-29-2013 - special build for Steve Craig to narrow issue with extra charts when switching
    '   from pae with lots of charts (intraday.gzp) to page with 2 charts (quote.gzp)
    Set CachePageSave = Nothing         'aChartForms

ErrExit:
    Set frm = Nothing
    Set Page = Nothing
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.CachePageSave"

End Function

Public Function CachePageRestore(ByRef aReuseable As cGdArray, ByVal strPageName$) As Boolean
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    
    Dim i&, iFrmIndex&, iChartsCount&, iWindowState&
    
    Dim bSuccess As Boolean

    Dim frm As frmChart
    Dim frmActive As frmChart
    Dim Page As cPageCache
    Dim ChartCached As cChartCache
    Dim Chart As cChart
    
    Dim bResizeTimerSave As Boolean

    If bInProgress Then Exit Function
    
    If g.ChartPageCache Is Nothing Then Exit Function
    Set Page = g.ChartPageCache(strPageName)
    If Page Is Nothing Then Exit Function

    'precautionary checks, theoretically both of these will always be true
    If Page.PageName <> strPageName Then Exit Function
    
    iChartsCount = Page.ChartsCount
    If iChartsCount <= 0 Then Exit Function
    
    g.bPageHasEWILabels = False
    g.bLoadingChartPage = True
    bInProgress = True
    
    ChartTimers = False
    bResizeTimerSave = frmMain.tmrAutoResize.Enabled
    frmMain.tmrAutoResize.Enabled = False
    
    i = 1
    iFrmIndex = -1
    If Not aReuseable Is Nothing Then
        If aReuseable.Size > 0 Then
            iFrmIndex = aReuseable.Size - 1
        End If
    End If
    
    If iFrmIndex = -1 Then
        iFrmIndex = Forms.Count - 1
        'attach each chart object in cache to an existing form
        Do While i <= iChartsCount
            Set ChartCached = Page.CachedObject(i)
            If Not ChartCached Is Nothing Then
                Set Chart = ChartCached.CacheChartGet
                If Not Chart Is Nothing Then
                    Set frm = GetAvailChartForm(iFrmIndex)
                    
                    If frm Is Nothing Then
                        Set frm = New frmChart
                    Else
                        If frm.WindowState = vbMaximized Then
                            '12-05-2012: if form was never in normal mode (i.e. create new chart page)
                            'then gets re-used for a chart that is in normal mode then the chart
                            'that reused this form is mis-positioned (i.e. restored out of place)
                            'Tim reported this using Needham Test.GZP
                            frm.WindowState = vbNormal
                        End If
                        frm.Hide
                        frm.SkipFocusFix = True
                        frm.PfpReset ePfpReset_ClearAll
                        frm.Chart.SeasonalIndClear
                        frm.Chart.ClearChartForReuse
                        frm.Tag = ""
                        frm.tmr.Tag = ""
                    End If
                
                    ChartCached.CacheChartRestore frm
                    If frm.OrderBarMode = eOrdBarMode_Wizard Then frm.tmr.Tag = "ToggleOrderbarMode"       '4992
                
                    If frm.DetachStatus = eDetached Then
                        frm.tmr.Tag = "DETACH_NOW"
                        'a new window is always shown as windowstate normal
                        'save the state now so we know to maximize it later if necessary
                        frm.Tag = ChartCached.CacheChartWindowState
                        g.ChartGlobals.nDetached = g.ChartGlobals.nDetached + 1
                    Else
                        Set frmActive = frm
                        iWindowState = ChartCached.CacheChartWindowState
                        ShowForm frm
                        frm.ZOrder
                    End If
                End If
            End If
            
            i = i + 1
        Loop
        Set frm = Nothing
    
        'unload any existing chart form that was not reused
        For i = iFrmIndex To 0 Step -1
            If IsFrmChart(Forms(i)) Then
                Unload Forms(i)
            End If
        Next
    
    Else
        Do While i <= iChartsCount
            Set ChartCached = Page.CachedObject(i)
            
            If Not ChartCached Is Nothing Then
                Set Chart = ChartCached.CacheChartGet
                If Not Chart Is Nothing Then
                
                    If iFrmIndex >= 0 Then
                        Set frm = aReuseable(iFrmIndex)
                    
                        frm.Hide
                        frm.SkipFocusFix = True
                        frm.PfpReset ePfpReset_ClearAll
                        frm.Chart.SeasonalIndClear
                        frm.Chart.ClearChartForReuse
                        frm.Tag = ""
                        frm.tmr.Tag = ""
                    
                        iFrmIndex = iFrmIndex - 1
                    Else
                        Set frm = New frmChart
                    End If
                    
                    ChartCached.CacheChartRestore frm
                    If frm.OrderBarMode = eOrdBarMode_Wizard Then frm.tmr.Tag = "ToggleOrderbarMode"       '4992
                    
                    If frm.DetachStatus = eDetached Then
                        frm.tmr.Tag = "DETACH_NOW"
                        'a new window is always shown as windowstate normal
                        'save the state now so we know to maximize it later if necessary
                        frm.Tag = ChartCached.CacheChartWindowState
                        g.ChartGlobals.nDetached = g.ChartGlobals.nDetached + 1
                    Else
                        Set frmActive = frm
                        ShowForm frm
                    End If
                End If
            End If
            
            i = i + 1
        Loop
        Set frm = Nothing
    
        'unload any existing chart form that was not reused
        For i = iFrmIndex To 0 Step -1
            Unload aReuseable(i)
        Next
    
    End If
    
    If Not frmActive Is Nothing Then
        Set g.ChartGlobals.frmActiveNonDetached = frmActive
        frmMain.SetWindowLink frmActive
        
        If Page.ChartsMaximized Then
            frmActive.WindowState = vbMaximized
        Else
            frmActive.WindowState = vbNormal
        End If
    End If
    
    g.bLoadingChartPage = False
    
    UpdateVisibleCharts -1
    InfBox
    DoEvents
    
    
    If Not frmActive Is Nothing Then
        ActiveChartFormSet frmActive
    End If
    Set frmActive = Nothing
    
    Page.PageCacheReleaseObjects
    Set Page = Nothing
    g.ChartPageCache.Remove strPageName
    
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
    
    bSuccess = True
    frmMain.tmrAutoResize.Enabled = bResizeTimerSave
    
ErrExit:
    CachePageRestore = bSuccess
    bInProgress = False
    
    Exit Function

ErrSection:
    bInProgress = False
    RaiseError "mChartLadderCtrls.CachePageRestore"

End Function

Private Function GetAvailChartForm(ByRef Index&) As frmChart
On Error GoTo ErrSection:

    Dim i&
    Dim frmReturn As frmChart

    If Index < Forms.Count And Index >= 0 Then
        For i = Index To 0 Step -1
            If TypeOf Forms(i) Is frmChart Then
                If Forms(i).Tag = "BOGUS" Then
                    Unload Forms(i)
                Else
                    Set frmReturn = Forms(i)
                    Exit For
                End If
            ElseIf TypeOf Forms(i) Is frmChart2 Then
                Unload Forms(i)         'cannot reuse detached chart forms
            End If
        Next
    End If
    
    Index = i - 1
    Set GetAvailChartForm = frmReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.GetAvailChartForm"

End Function


Private Sub OHLCString(Bars As cGdBars, ByVal Index&, _
    ByRef strOpen$, ByRef strHigh$, ByRef strLow$, ByRef strClose$)
On Error GoTo ErrSection:

    Dim d#
    
    'initialize return variables
    strOpen = ""
    strHigh = ""
    strLow = ""
    strClose = ""
    
    If Bars Is Nothing Then Exit Sub
    
    d = Bars(eBARS_Open, Index)
    If d <> kNullData Then strOpen = CStr(d)

    d = Bars(eBARS_High, Index)
    If d <> kNullData Then strHigh = CStr(d)

    d = Bars(eBARS_Low, Index)
    If d <> kNullData Then strLow = CStr(d)

    d = Bars(eBARS_Close, Index)
    If d <> kNullData Then strClose = CStr(d)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.OHLCString"

End Sub

Public Sub DataToClipboard(frm As Form, ByVal bIndCopy As Boolean)
On Error GoTo ErrSection:

    Dim i&, j&, d#
    Dim bHasData As Boolean
    Dim strData$, strTemp$
    Dim strOpen$, strHigh$, strLow$, strClose$
    
    Dim Chart As cChart
    Dim Bars As cGdBars
    Dim Tree As cGdTree
    
    Dim Pane As cPane
    Dim Ind As cIndicator
    
    Dim aData As cGdArray
    Dim aIndicators As cGdArray
    
    If frm Is Nothing Then Exit Sub
    If Not IsFrmChart(frm) Then Exit Sub
    
    Set Chart = frm.Chart
    If Chart Is Nothing Then Exit Sub
    
    Set Bars = Chart.Bars
    If Bars Is Nothing Then Exit Sub
    If Bars.Size <= 0 Then Exit Sub
    
    If bIndCopy Then Set Tree = Chart.Tree
    
    Set aData = New cGdArray
    Set aIndicators = New cGdArray
    
    strData = "Date" & vbTab & "Open" & vbTab & "High" & vbTab & "Low" & vbTab & "Close" & vbTab & "Volume" & vbTab & "Open Int."
    
    If bIndCopy And Not Tree Is Nothing Then
        For i = 1 To Tree.Count
            If TypeOf Tree(i) Is cPane Then
                Set Pane = Tree(i)
            ElseIf TypeOf Tree(i) Is cIndicator Then
                Set Ind = Tree(i)
                If Ind.isPriceInd <> 1 Then
                    If Not Pane Is Nothing Then
                        If Pane.Display And Ind.Display Then
                            If Ind.DataType = eINDIC_BarData Then
                                If Ind.DisplayType < 0 Then
                                    If Ind.DisplayType = eINDIC_HL Then
                                        strData = strData & vbTab & Ind.Name & " High"
                                        strData = strData & vbTab & Ind.Name & " Low"
                                    ElseIf Ind.DisplayType = eINDIC_HLC Then
                                        strData = strData & vbTab & Ind.Name & " High"
                                        strData = strData & vbTab & Ind.Name & " Low"
                                        strData = strData & vbTab & Ind.Name & " Close"
                                    Else
                                        strData = strData & vbTab & Ind.Name & " Open"
                                        strData = strData & vbTab & Ind.Name & " High"
                                        strData = strData & vbTab & Ind.Name & " Low"
                                        strData = strData & vbTab & Ind.Name & " Close"
                                    End If
                                Else
                                    strData = strData & vbTab & Ind.Name & " Close"
                                End If
                            Else
                                strData = strData & vbTab & Ind.Name
                            End If
                            aIndicators.Add Ind
                        End If
                    End If
                End If
            End If
        Next
    End If
    Set Ind = Nothing
    
    aData.Add strData
    
    Screen.MousePointer = vbHourglass
    For i = 0 To Bars.Size - 1
        strTemp = DateFormat(Bars(eBARS_DateTime, i), MM_DD_YYYY)
        strData = Right(strTemp, 4) & "/" & Left(strTemp, 5)
        If Bars.IsIntraday Then
            strData = strData & " " & DateFormat(Bars(eBARS_DateTime, i), NO_DATE, HH_MM_SS)
        End If
        
        If Bars(eBARS_Close, i) = kNullData Then
            strData = strData & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & " "
        Else
            OHLCString Bars, i, strOpen, strHigh, strLow, strClose
            strData = strData & vbTab & strOpen & vbTab & strHigh & vbTab & strLow & vbTab & strClose
        End If
            
        d = Bars(eBARS_Vol, i)
        If d = kNullData Then
            strData = strData & vbTab & " "
        Else
            strData = strData & vbTab & d
        End If
        
        d = Bars(eBARS_OI, i)
        If d = kNullData Then
            strData = strData & vbTab & " "
        Else
            strData = strData & vbTab & d
        End If
        
        For j = 0 To aIndicators.Size - 1
            Set Ind = aIndicators(j)
            
            Select Case Ind.DataType
                Case eINDIC_Constant
                    strTemp = Ind.Parm(0)
                    If Len(strTemp) <= 0 Then
                        strData = strData & vbTab & " "
                    Else
                        strData = strData & vbTab & strTemp
                    End If
                
                Case eINDIC_Array
                    If Ind.Data(i) = kNullData Then
                        strData = strData & vbTab & " "
                    Else
                        strData = strData & vbTab & Ind.Data(i)
                    End If
                
                Case eINDIC_BarData
                    OHLCString Ind.Bars, i, strOpen, strHigh, strLow, strClose
                    If Ind.DisplayType < 0 Then
                        If Ind.DisplayType = eINDIC_HL Then
                            strData = strData & vbTab & strHigh & vbTab & strLow
                        ElseIf Ind.DisplayType = eINDIC_HLC Then
                            strData = strData & vbTab & strHigh & vbTab & strLow & vbTab & strClose
                        Else
                            strData = strData & vbTab & strOpen & vbTab & strHigh & vbTab & strLow & vbTab & strClose
                        End If
                    Else
                        strData = strData & vbTab & strClose
                    End If
                    
                Case eINDIC_BooleanArray
                    If Ind.Data(i) = 0 Then
                        strData = strData & vbTab & "0"
                    Else
                        strData = strData & vbTab & "1"
                    End If
            End Select
        Next
            
        ' check if this record has data (i.e. something other than just a date/time)
        bHasData = False
        If Bars(eBARS_Close, i) <> kNullData Then
            bHasData = True
        Else
            ' start checking beyond the date/time
            j = InStr(strData, vbTab)
            If j > 0 Then
                For j = j To Len(strData)
                    strTemp = Mid(strData, j, 1)
                    If strTemp <> " " And strTemp <> vbTab Then
                        bHasData = True
                        Exit For
                    End If
                Next
            End If
        End If
            
        If bHasData Then
            aData.Add strData
        End If
    Next
    
    For i = 0 To aIndicators.Size - 1
        Set aIndicators(i) = Nothing
    Next
    
    Set Chart = Nothing
    Set Bars = Nothing
    Set Tree = Nothing
    Set Pane = Nothing
    Set Ind = Nothing
    Set aIndicators = Nothing
   
    ' TLB 11/17/2011: the following section needed to be moved to above the ErrExit section
    ' since it was erroring when the string or clipboard memory was being exceeded ...
    Clipboard.Clear
    strData = aData.JoinFields(vbCrLf)
    Set aData = Nothing
    Clipboard.SetText strData
    
    Screen.MousePointer = 0
    strData = "You can now paste the data into |another application by selecting |'Edit-Paste'  (or hit 'Ctrl-V')."
    InfBox "i=i ; h=Copy chart data ; " + strData

ErrExit:
    Set Chart = Nothing
    Set Bars = Nothing
    Set Tree = Nothing
    Set Pane = Nothing
    Set Ind = Nothing
    Set aIndicators = Nothing
    Set aData = Nothing

    Exit Sub

ErrSection:
    Screen.MousePointer = 0
    ' give a little nicer message if ran out of memory when trying to paste into the clipboard
    If Err.Number = 7 Or Err.Number = 14 Then
        InfBox "Not enough clipboard memory| for this amount of data.", "e", , "ERROR"
    Else
        RaiseError "mChartLadderCtrls.DataToClipboard"
    End If
    Resume ErrExit
End Sub

Private Function ExitFavoritesBaseSym(frm As Form) As String
On Error GoTo ErrSection

    Dim strBaseSym$
    
    If IsFrmChart(frm) Then
        strBaseSym = frm.Chart.Symbol
    ElseIf TypeOf frm Is frmTickDistribution Then
        strBaseSym = g.SymbolPool.SymbolForID(frm.SymbolID)
    Else
        Exit Function
    End If
    
    If Len(strBaseSym) = 0 Then Exit Function

    strBaseSym = BaseForAutoExitFavorites(strBaseSym)
    
    ExitFavoritesBaseSym = strBaseSym

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.ExitFavoritesBaseSym"

End Function

Public Sub TSOGrpFavoritesEdit(frm As Form, ByVal bEdit As Boolean)
On Error GoTo ErrSection:

    Dim frmActiveGrps As Form
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If Not IsFrmChart(frm) Then Exit Sub
    
    If HasLevel(eTN4_Gold, True, "TradeSense Orders") Then
        If FormIsLoaded("frmActiveTsOrderGroups") Then
            Set frmActiveGrps = frmActiveTsOrderGroups
            If frmActiveGrps.Visible = False Then
                mGenesis.ShowForm frmActiveGrps, eForm_Nonmodal, frmMain
            End If
            If bEdit Then frmTradeSenseOrderGroups.ShowMe
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.TSOGrpFavoritesEdit"

End Sub

Public Sub TSOGrpFavoritesBtnClick(frm As Form, ctrl As vsElastic, ByVal globalArrayIdx&)
On Error GoTo ErrSection:
    
#If 1 Then
    Dim strTsogFilename As String       ' TradeSense order group filename
    Dim tsoGrp As cTradeSenseOrderGroup ' TradeSense order group object
    
    If (g.bUnloading = False) And (g.bLoadingChartPage = False) Then
        If IsFrmChart(frm) = True Then
            If (globalArrayIdx >= 0) And (globalArrayIdx <= 3) Then
                If HasLevel(eTN4_Gold, True, "TradeSense Orders") Then
                    strTsogFilename = Parse(g.ChartGlobals.astrTsogFavorites(globalArrayIdx), "|", 1)
                    If Len(strTsogFilename) = 0 Then
                        TSOGrpFavoritesEdit frm, True
                    ElseIf ctrl.Appearance = apInset Then
                        TSOGrpFavoritesEdit frm, False
                    Else
                        Set tsoGrp = New cTradeSenseOrderGroup
                        
                        tsoGrp.FromFile strTsogFilename, (InStr(strTsogFilename, "Custom\") > 0)
                        frmTradeSenseOrderGroups.HandleTradeSenseWrapper tsoGrp, frm.Chart.Symbol, frm.Chart.TradeAccountID
                    End If
                End If
            End If
        End If
    End If
#Else
    Dim frmActiveGrps As Form
    Dim tsoGrp As cTradeSenseOrderGroup
    Dim bUnloadWhenDone As Boolean
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If Not IsFrmChart(frm) Then Exit Sub
    
    If globalArrayIdx < 0 Or globalArrayIdx > 3 Then Exit Sub
    
    If HasLevel(eTN4_Gold, True, "TradeSense Orders") Then
        Set tsoGrp = g.ChartGlobals.aTSOGrpFavorites(globalArrayIdx)
        If tsoGrp Is Nothing Then Exit Sub
        
        If Len(tsoGrp.ID) = 0 Then
            TSOGrpFavoritesEdit frm, True
        ElseIf ctrl.Appearance = apInset Then
            TSOGrpFavoritesEdit frm, False
        Else
            frmTradeSenseOrderGroups.HandleTradeSenseWrapper tsoGrp, frm.Chart.Symbol, frm.Chart.TradeAccountID
        End If
    End If
#End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.TSOGrpFavoritesBtnClick"

End Sub

Public Sub TSOGrpFavoritesCheck(frm As Form)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol from the chart
    Dim strAccountID As String          ' Account ID from the chart
    Dim strTsogFilename As String       ' TradeSense order group filename
    Dim strName As String               ' TradeSense order group name
    Dim strID As String                 ' TradeSense order group ID
    Dim vseButton As vsElastic          ' Appropriate button from the form
    Dim bActive As Boolean              ' Is the button active?
    Dim strKey As String                ' Key into the active TradeSense order group collection
    
#If 1 Then
    If (g.bUnloading = False) And (g.bLoadingChartPage = False) And (Not g.ChartGlobals.astrTsogFavorites Is Nothing) Then
        If IsFrmChart(frm) Then
            If frm.fraTSO.Visible = True Then
                strSymbol = frm.Chart.Symbol
                strAccountID = Str(frm.Chart.TradeAccountID)
                
                For lIndex = 0 To 3
                    Select Case lIndex
                        Case 0:
                            Set vseButton = frm.vseTSO1
                        Case 1:
                            Set vseButton = frm.vseTSO2
                        Case 2:
                            Set vseButton = frm.vseTSO3
                        Case 3:
                            Set vseButton = frm.vseTSO4
                    End Select
                    
                    strName = ""
                    bActive = False
                    
                    strTsogFilename = Parse(g.ChartGlobals.astrTsogFavorites(lIndex), "|", 1)
                    If Len(strTsogFilename) > 0 Then
                        strID = Parse(g.ChartGlobals.astrTsogFavorites(lIndex), "|", 2)
                        strName = Parse(g.ChartGlobals.astrTsogFavorites(lIndex), "|", 3)
                        strKey = strSymbol & vbTab & strAccountID & vbTab & strID
                        bActive = g.TsoGroups.Exists(strKey)
                    End If
                        
                    vseButton.ToolTipText = strName
                    vseButton.Font.Bold = (Len(strName) > 0)
                    
                    If (bActive = True) And (vseButton.Appearance <> apInset) Then
                        vseButton.Appearance = apInset
                    ElseIf (bActive = False) And (vseButton.Appearance = apInset) Then
                        vseButton.Appearance = ap3D
                    End If
                Next lIndex
            End If
        End If
    End If
#Else
    Dim i&, strKey$, strID$, strSymbol$, strTradeAcct$
        
    Dim bActive1 As Boolean
    Dim bActive2 As Boolean
    Dim bActive3 As Boolean
    Dim bActive4 As Boolean
    
    Dim TsoGrp1 As cTradeSenseOrderGroup
    Dim TsoGrp2 As cTradeSenseOrderGroup
    Dim TsoGrp3 As cTradeSenseOrderGroup
    Dim TsoGrp4 As cTradeSenseOrderGroup
    
    Dim TsoActive As cActiveTsOrderGroup
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    
    If Not IsFrmChart(frm) Then Exit Sub
    If Not frm.fraTSO.Visible Then Exit Sub
    
    If g.ChartGlobals.aTSOGrpFavorites Is Nothing Then Exit Sub

    'set button font to bold if assigned to a tradesense order group
    Set TsoGrp1 = g.ChartGlobals.aTSOGrpFavorites(0)
    If Not TsoGrp1 Is Nothing Then
        frm.vseTSO1.ToolTipText = TsoGrp1.Name
        If Len(TsoGrp1.Name) > 0 Then
            frm.vseTSO1.Font.Bold = True
        Else
            frm.vseTSO1.Font.Bold = False
        End If
    End If
    
    Set TsoGrp2 = g.ChartGlobals.aTSOGrpFavorites(1)
    If Not TsoGrp2 Is Nothing Then
        frm.vseTSO2.ToolTipText = TsoGrp2.Name
        If Len(TsoGrp2.Name) > 0 Then
            frm.vseTSO2.Font.Bold = True
        Else
            frm.vseTSO2.Font.Bold = False
        End If
    End If

    Set TsoGrp3 = g.ChartGlobals.aTSOGrpFavorites(2)
    If Not TsoGrp3 Is Nothing Then
        frm.vseTSO3.ToolTipText = TsoGrp3.Name
        If Len(TsoGrp3.Name) > 0 Then
            frm.vseTSO3.Font.Bold = True
        Else
            frm.vseTSO3.Font.Bold = False
        End If
    End If

    Set TsoGrp4 = g.ChartGlobals.aTSOGrpFavorites(3)
    If Not TsoGrp4 Is Nothing Then
        frm.vseTSO4.ToolTipText = TsoGrp4.Name
        If Len(TsoGrp4.Name) > 0 Then
            frm.vseTSO4.Font.Bold = True
        Else
            frm.vseTSO4.Font.Bold = False
        End If
    End If

    'set button appearance to depressed if assigned group is active for this chart
    strSymbol = frm.Chart.Symbol
    strTradeAcct = Str(frm.Chart.TradeAccountID)
    
    For i = 1 To g.TsoGroups.Count
        Set TsoActive = g.TsoGroups(i)
        If Not TsoActive Is Nothing Then
            strKey = TsoActive.Key
            strID = Parse(strKey, vbTab, 3)
            
            If strSymbol = Parse(strKey, vbTab, 1) Then
                If strTradeAcct = Parse(strKey, vbTab, 2) Then
                    If Len(TsoGrp1.ID) > 0 And TsoGrp1.ID = strID Then
                        bActive1 = True
                        If frm.vseTSO1.Appearance <> apInset Then frm.vseTSO1.Appearance = apInset
                    End If
                    If Len(TsoGrp2.ID) > 0 And TsoGrp2.ID = strID Then
                        bActive2 = True
                        If frm.vseTSO2.Appearance <> apInset Then frm.vseTSO2.Appearance = apInset
                    End If
                    If Len(TsoGrp3.ID) > 0 And TsoGrp3.ID = strID Then
                        bActive3 = True
                        If frm.vseTSO3.Appearance <> apInset Then frm.vseTSO3.Appearance = apInset
                    End If
                    If Len(TsoGrp4.ID) > 0 And TsoGrp4.ID = strID Then
                        bActive4 = True
                        If frm.vseTSO4.Appearance <> apInset Then frm.vseTSO4.Appearance = apInset
                    End If
                End If
            End If
        End If
    Next
    
    'clear button depressed appearance if was previously active, but no longer acter
    If Not bActive1 And frm.vseTSO1.Appearance = apInset Then
        frm.vseTSO1.Appearance = ap3D
    End If
    If Not bActive2 And frm.vseTSO2.Appearance = apInset Then
        frm.vseTSO2.Appearance = ap3D
    End If
    If Not bActive3 And frm.vseTSO3.Appearance = apInset Then
        frm.vseTSO3.Appearance = ap3D
    End If
    If Not bActive4 And frm.vseTSO4.Appearance = apInset Then
        frm.vseTSO4.Appearance = ap3D
    End If
#End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.TSOGrpFavoritesCheck"

End Sub

Public Function CachePageRemove(ByVal strPageName$) As Long
On Error GoTo ErrSection:

    Dim i&
    Dim Page As cPageCache
    
    If g.ChartPageCache Is Nothing Then Exit Function
    
    For i = 1 To g.ChartPageCache.Count
        Set Page = g.ChartPageCache(i)
        If Page.PageName = strPageName Then
            If g.ChartPageCache.Remove(i) Then
                CachePageRemove = 1
                Exit For
            End If
        End If
    Next

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mChartLadderCtrls.CachePageRemove"

End Function

Public Sub StartStopTimeLabel(Chart As cChart, lblStartStopTimes As ctlUniLabelXP, lblStartStopInfo2 As ctlUniLabelXP)
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim dStart#, dEnd#
    Dim strStart$, strZone$
    
    If Not Chart Is Nothing Then Set Bars = Chart.Bars
        
    If Chart Is Nothing Or Bars Is Nothing Then
        lblStartStopTimes.Caption = ""
        lblStartStopInfo2.Caption = ""
        Exit Sub
    End If
    
    dStart = Bars.Prop(eBARS_StartTime) / 1440#
    
    If dStart = 0 Then
        strStart = "00:00"
    Else
        strStart = DateFormat(dStart, NO_DATE, HH_MM, NO_AMPM)
    End If
    
    dEnd = Bars.Prop(eBARS_EndTime) / 1440#
    
    
    lblStartStopTimes.Caption = strStart & " - " & DateFormat(dEnd, NO_DATE, HH_MM, NO_AMPM)
    
    strZone = Chart.Bars.Prop(eBARS_ExchangeTimeZoneInf)
    If strZone <> "NY" And strZone <> "GMT" Then strZone = "Exchange"
    lblStartStopInfo2.Caption = "( times are in " & strZone & " time )"
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mChartLadderCtrls.StartStopTimeLabel"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShortDisplayNumber
'' Description: Build a short display version of the given number
'' Inputs:      Number
'' Returns:     Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShortDisplayNumber(ByVal lNumber As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = Str(lNumber)
    If Len(strReturn) > 6 Then
        ' 03/02/2016 DAJ: Changed from checking for five zeroes to checking for six zeroes...
        If Right(strReturn, 6) = "000000" Then
            strReturn = Left(strReturn, Len(strReturn) - 6) & "M"
        End If
    End If
    If Len(strReturn) > 3 Then
        If Right(strReturn, 3) = "000" Then
            strReturn = Left(strReturn, Len(strReturn) - 3) & "k"
        End If
    End If
    
    ShortDisplayNumber = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mChartLadderCtrls.ShortDisplayNumber"
    
End Function
