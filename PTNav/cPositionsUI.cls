VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPositionsUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPositionsUI.cls
'' Description: Wrapper class for a positions grid
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/15/2010   DAJ         Fixed the grid column information persistence
'' 03/17/2010   DAJ         Update the number visible items for the toolbar captions
'' 03/17/2010   DAJ         Removed HasOpenPositions call
'' 03/17/2010   DAJ         Toggle show/hide of auto trade item column
'' 03/18/2010   DAJ         Don't count a position for a hidden account as visible
'' 10/26/2010   DAJ         Efficiency enhancements
'' 05/11/2011   DAJ         Utilize CheckSimTradeStatus, EnableCheckStatusMenu
'' 06/28/2011   DAJ         Setup clickable cells like hyperlinks
'' 09/23/2011   DAJ         Show date journals form instead of old journals form
'' 10/04/2011   DAJ         Call the ShowJournals function instead of calling the form direct
'' 10/29/2012   DAJ         Employ Tim's new fix for grid scrolling vs. streaming issue
'' 11/28/2012   DAJ         Speed enhancements for the Trade Console
'' 12/06/2012   DAJ         Fix for open equity on accounts not updating all the time
'' 02/20/2013   DAJ         Added "Actual Performance" menu item
'' 03/21/2013   DAJ         Don't clear updated colors on a hidden row
'' 05/08/2013   DAJ         Log the last modified value when positions are updated
'' 02/06/2014   DAJ         Only add symbols to the stream if the position is visible
'' 03/12/2014   DAJ         Need to add symbols to the stream if "Total" line and in a position
'' 05/28/2014   DAJ         Display blank for null open profit
'' 06/26/2014   DAJ         Display the last for the last price on options instead of bid/ask
'' 08/05/2014   DAJ         After user scrolls the grid, call refresh prices
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/09/2014   DAJ         Fix for price/open equity not updating immediately after scroll
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kMenuPrefix = "P:"

Private WithEvents m_fgPositions As VSFlexGrid
Attribute m_fgPositions.VB_VarHelpID = -1
Private WithEvents m_tmrMenu As Timer
Attribute m_tmrMenu.VB_VarHelpID = -1
Private WithEvents m_mnuFlatten As Menu
Attribute m_mnuFlatten.VB_VarHelpID = -1
Private WithEvents m_mnuReverse As Menu
Attribute m_mnuReverse.VB_VarHelpID = -1
Private WithEvents m_mnuManageXOS As Menu
Attribute m_mnuManageXOS.VB_VarHelpID = -1
Private WithEvents m_mnuSelectXOS As Menu
Attribute m_mnuSelectXOS.VB_VarHelpID = -1
Private WithEvents m_mnuRemoveXOS As Menu
Attribute m_mnuRemoveXOS.VB_VarHelpID = -1
Private WithEvents m_mnuPrint As Menu
Attribute m_mnuPrint.VB_VarHelpID = -1
Private WithEvents m_mnuActualPerformance As Menu
Attribute m_mnuActualPerformance.VB_VarHelpID = -1
Private WithEvents m_mnuTradeHistory As Menu
Attribute m_mnuTradeHistory.VB_VarHelpID = -1
Private WithEvents m_mnuSettings As Menu
Attribute m_mnuSettings.VB_VarHelpID = -1
Private WithEvents m_mnuCheckStatus As Menu
Attribute m_mnuCheckStatus.VB_VarHelpID = -1
Private WithEvents m_mnuViewJournals As Menu
Attribute m_mnuViewJournals.VB_VarHelpID = -1
Private WithEvents m_mnuAutoSizeColumns As Menu
Attribute m_mnuAutoSizeColumns.VB_VarHelpID = -1
Private WithEvents m_mnuDefaultColumns As Menu
Attribute m_mnuDefaultColumns.VB_VarHelpID = -1

Private Enum eGDCols
    eGDCol_PositionID = 0
    eGDCol_AccountID
    eGDCol_AccountType
    eGDCol_SymbolID
    eGDCol_Remove
    
    eGDCol_Symbol
    eGDCol_Account
    eGDCol_AutoTradeItem
    eGDCol_Position
    eGDCol_Flatten
    
    eGDCol_AvgEntry
    eGDCol_CurrentPrice
    eGDCol_OpenProfit
    eGDCol_OrderStrategy
    eGDCol_LastTraded
    eGDCol_SessionDate
    eGDCol_SessionQuantity
    eGDCol_SessionProfit
    eGDCol_NumCols
End Enum

Private Type mPrivate
    strSource As String                 ' Name for the calling form
    TCGrid As cTradeConsoleGrid         ' Trade console grid object for generic routines
    tmrRealTime As Timer                ' Realtime timer control
    frm As Form                         ' Form that the grid is on
    SnapshotPositions As cGdTree        ' Collection of snapshot positions collections by broker
    mnuPositions As Object              ' Positions menu
    
    bSummary As Boolean                 ' Is this a summary grid?
    strSummaryDisplay As String         ' Summary display string
    strDisplay As String                ' Non-summary display string
    
    LastPrice As cGdTree                ' Collection of last known prices
    PositionMatch As cGdTree            ' Collection of position match information
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

Public Property Get Summary() As Boolean
    Summary = m.bSummary
End Property
Public Property Let Summary(ByVal bSummary As Boolean)
    If bSummary <> m.bSummary Then
        m.bSummary = bSummary
        If m.bSummary Then
            m.TCGrid.Display = m.strSummaryDisplay
        Else
            m.TCGrid.Display = m.strDisplay
        End If
        ApplyDisplayString
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object
'' Inputs:      Source, Working Orders UI, Summary?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal strSource As String, UI As cPositionsControls, ByVal bSummary As Boolean)
On Error GoTo ErrSection:

    m.strSource = strSource
    m.bSummary = bSummary
    
    With UI
        Set m.frm = .frm
        Set m_fgPositions = .fgGrid
        Set m_tmrMenu = .tmrMenu
        Set m.tmrRealTime = .tmrRealTime
        Set m.mnuPositions = .mnuPositions
        Set m_mnuFlatten = .mnuFlatten
        Set m_mnuReverse = .mnuReverse
        Set m_mnuSelectXOS = .mnuSelectXOS
        Set m_mnuManageXOS = .mnuManageXOS
        Set m_mnuRemoveXOS = .mnuRemoveXOS
        Set m_mnuPrint = .mnuPrint
        Set m_mnuActualPerformance = .mnuActualPerformance
        Set m_mnuTradeHistory = .mnuTradeHistory
        Set m_mnuSettings = .mnuSettings
        Set m_mnuCheckStatus = .mnuCheckStatus
        Set m_mnuViewJournals = .mnuViewJournals
        Set m_mnuAutoSizeColumns = .mnuAutoSizeColumns
        Set m_mnuDefaultColumns = .mnuDefaultColumns
    End With

    InitGrid
    
    m.strSummaryDisplay = FixDisplay(GetIniFileProperty("SummaryPositionsDisplay", "", "TTSummary", g.strIniFile), True)
    m.strDisplay = FixDisplay(GetIniFileProperty("PositionsDisplay", "", "TTSummary", g.strIniFile), False)
    
    Set m.TCGrid = New cTradeConsoleGrid
    If m.bSummary Then
        m.TCGrid.Init m_fgPositions, m.tmrRealTime, m.strSummaryDisplay
    Else
        m.TCGrid.Init m_fgPositions, m.tmrRealTime, m.strDisplay
    End If
    
    ApplyDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Clear the updated colors on the grid if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearUpdatedColors()
On Error GoTo ErrSection:

    Dim iSaveRedraw As Integer          ' Current state of the redraw for the grid
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim bStillColor As Boolean          ' Do we still need the cell colored?
    Dim dTickCount As Double            ' Tick count that had been stored in the cell

    With m_fgPositions
        iSaveRedraw = .Redraw
        .Redraw = flexRDNone
        
        lCol = PositionsCol(eGDCol_CurrentPrice)
        For lRow = .FixedRows To .Rows - 1
            If g.bUnloading = False Then
                If .RowHidden(lRow) = False Then
                    If .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UpdateColor Then
                        bStillColor = False
                        If m.tmrRealTime.Enabled Then
                            If .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UpdateColor Then
                                ' see if has been more than 1 second since colored
                                dTickCount = .Cell(flexcpData, lRow, lCol)
                                dTickCount = gdTickCount - dTickCount
                                If dTickCount >= 0 And dTickCount <= 1000 Then
                                    bStillColor = True
                                Else
                                    .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UnchColor
                                End If
                            End If
                        End If
                        
                        ' color symbol cell only if a cell was still colored
                        If Not bStillColor Then
                            .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UnchColor
                        End If
                    End If
                End If
            Else
                Exit For
            End If
        Next
        
        .Redraw = iSaveRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPositionsUI.ClearUpdatedColors"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPrices
'' Description: Refresh the prices in the grid as appropriate
'' Inputs:      Symbol, Price, Bid, Ask
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPrices(ByVal vSymbolOrSymbolID As Variant, ByVal dPrice As Double, ByVal dBid As Double, ByVal dAsk As Double)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strValue As String              ' String representation of the value
    Dim AcctPos As cAccountPosition     ' Temporary account position object
    Dim dCurrent As Double              ' Current price to use for open profit
    Dim Bars As New cGdBars             ' Bars object
    Dim dOpenProfit As Double           ' Open profit
    
    SetBarProperties Bars, vSymbolOrSymbolID
    
    With m_fgPositions
        For lIndex = .FixedRows To .Rows - 1
            Set AcctPos = PositionForRow(lIndex)
            If Not AcctPos Is Nothing Then
                ' DAJ 08/06/2007: Only need to updated prices on visible rows...
                If (.RowHidden(lIndex) = False) Or (AcctPos.AutoTradeItemID = -1) Then
                    If (g.Broker.PositionMatch(AcctPos.AccountID, AcctPos.SymbolOrSymbolID) = True) Then
                        If AcctPos.SymbolOrSymbolID = vSymbolOrSymbolID Then
                            dCurrent = AcctPos.CurrentPrice(dPrice, dBid, dAsk)
                        
                            strValue = ""
                            If dPrice <> kNullData Then
                                ' DAJ 06/26/2014: Pete and Brady have both found that the current price
                                ' on the positions grid is different than everywhere else when looking at
                                ' options.  So, display the current price even if we use bid/ask for open equity...
                                strValue = Bars.PriceDisplay(dPrice)
                            End If
                            If dCurrent <> kNullData Then
                                'strValue = Bars.PriceDisplay(dCurrent)
                                dOpenProfit = AcctPos.OpenProfit(dCurrent)
                                If dOpenProfit <> kNullData Then
                                    .TextMatrix(lIndex, PositionsCol(eGDCol_OpenProfit)) = dOpenProfit
                                    m.TCGrid.ColorCell lIndex, PositionsCol(eGDCol_OpenProfit)
                                Else
                                    .TextMatrix(lIndex, PositionsCol(eGDCol_OpenProfit)) = ""
                                End If
                                
                                g.Broker.UpdatePrice AcctPos.Broker, AcctPos.Key, dCurrent
                            Else
                                .TextMatrix(lIndex, PositionsCol(eGDCol_OpenProfit)) = ""
                            End If
                            
                            m.TCGrid.ChangeCell lIndex, PositionsCol(eGDCol_CurrentPrice), strValue
                        End If
                    End If
                End If
            Else
                .TextMatrix(lIndex, PositionsCol(eGDCol_CurrentPrice)) = ""
                .TextMatrix(lIndex, PositionsCol(eGDCol_OpenProfit)) = ""
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.RefreshPrices"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPrices2
'' Description: Refresh the prices in the grid as appropriate
'' Inputs:      Update Broker Price?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPrices2(ByVal bUpdateBrokerPrice As Boolean)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim AcctPos As cAccountPosition     ' Temporary account position object
    Dim dCurrent As Double              ' Current price to use for open profit
    Dim Bars As cGdBars                 ' Bars object
    Dim bPositionMatch As Boolean       ' Is the symbol in a position mismatch?
    Dim BarsColl As cGdTree             ' Collection of Bars objects
    Dim bBarsFound As Boolean           ' Bars were found in collection?
    Dim dLastPrice As Double            ' Last price known
    Dim lFrom As Long                   ' From value for the for loop
    Dim lTo As Long                     ' To value for the for loop
    Dim bContinue As Boolean            ' Continue?
    Dim strKey As String                ' Key into the position match collection
    Dim lTopRow As Long                 ' Top visible row in the grid
    Dim lBottomRow As Long              ' Bottom visible row in the grid
    Dim dOpenProfit As Double           ' Open profit
    Dim bVisibleRow As Boolean          ' Is the row visible on the screen?

    Set BarsColl = New cGdTree
    
    With m_fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lFrom = .FixedRows
        lTo = .Rows - 1
        lTopRow = .TopRow
        lBottomRow = .BottomRow
        
        For lIndex = lFrom To lTo
            Set AcctPos = PositionForRow(lIndex)
            If Not AcctPos Is Nothing Then
                ' DAJ 03/12/2014: Back on 02/06/2014, I checked in a fix commenting out the
                ' "Or (AcctPos.AutoTradeItemID = -1)" portion because we had a customer with
                ' a ton of symbols with fill summaries and we figured that we didn't have to
                ' subscribe to a symbol on the stream unless it was visible.  Unfortunately,
                ' I am using this to update the account open equity/current value as well, so
                ' if only the manual line is visible, we aren't updating the account values...
                'bContinue = (.RowHidden(lIndex) = False) ' Or (AcctPos.AutoTradeItemID = -1)
                bContinue = (.RowHidden(lIndex) = False) Or ((AcctPos.AutoTradeItemID = -1) And (AcctPos.CurrentPositionSnapshot <> 0))
                If bContinue Then
                    bVisibleRow = ((lIndex >= lTopRow) And (lIndex <= lBottomRow))
                    
                    If BarsColl.Exists(AcctPos.Symbol) Then
                        bBarsFound = True
                        Set Bars = BarsColl(AcctPos.Symbol)
                    Else
                        bBarsFound = False
                        Set Bars = frmTTSummary.GetBars(AcctPos.SymbolOrSymbolID)
                        BarsColl.Add Bars, AcctPos.Symbol
                    End If
                
                    strKey = Str(AcctPos.AccountID) & "|" & Str(AcctPos.SymbolOrSymbolID)
                    If m.PositionMatch.Exists(strKey) Then
                        bPositionMatch = m.PositionMatch(strKey)
                    Else
                        bPositionMatch = g.Broker.PositionMatch(AcctPos.AccountID, AcctPos.SymbolOrSymbolID)
                        m.PositionMatch.Add bPositionMatch, strKey
                    End If
                
                    If bPositionMatch Then
                        If m.LastPrice.Exists(AcctPos.Key) Then
                            dLastPrice = m.LastPrice(AcctPos.Key)
                        Else
                            dLastPrice = kNullData
                            m.LastPrice.Add dLastPrice, AcctPos.Key
                        End If
                                                
                        ' DAJ 06/26/2014: Pete and Brady have both found that the current price
                        ' on the positions grid is different than everywhere else when looking at
                        ' options.  So, display the current price even if we use bid/ask for open equity...
                        If bVisibleRow = True Then
                            dCurrent = Bars(eBARS_Close, Bars.Size - 1)
                            If dCurrent <> kNullData Then
                                If dCurrent <> dLastPrice Then
                                    m.TCGrid.ChangeCell lIndex, PositionsCol(eGDCol_CurrentPrice), Bars.PriceDisplay(dCurrent)
                                    m.LastPrice(AcctPos.Key) = dCurrent
                                End If
                            End If
                        End If
                        
                        dCurrent = AcctPos.CurrentPrice(Bars(eBARS_Close, Bars.Size - 1), Bars(eBARS_Bid, Bars.Size - 1), Bars(eBARS_Ask, Bars.Size - 1))
                        
                        If dCurrent <> kNullData Then
                            If dCurrent <> dLastPrice Then
                                If (bUpdateBrokerPrice = True) Then
                                    g.Broker.UpdatePrice AcctPos.Broker, AcctPos.Key, dCurrent
                                End If
                                'm.LastPrice(AcctPos.Key) = dCurrent
                                
                                If bVisibleRow = True Then
                                    dOpenProfit = AcctPos.OpenProfit(dCurrent)
                                    If dOpenProfit <> kNullData Then
                                        .TextMatrix(lIndex, PositionsCol(eGDCol_OpenProfit)) = dOpenProfit
                                        m.TCGrid.ColorCell lIndex, PositionsCol(eGDCol_OpenProfit)
                                    Else
                                        .TextMatrix(lIndex, PositionsCol(eGDCol_OpenProfit)) = ""
                                    End If
                                    
                                    'm.TCGrid.ChangeCell lIndex, PositionsCol(eGDCol_CurrentPrice), Bars.PriceDisplay(dCurrent)
                                End If
                            End If
                        Else
                            .TextMatrix(lIndex, PositionsCol(eGDCol_OpenProfit)) = ""
                        End If
                    End If
                End If
            Else
                .TextMatrix(lIndex, PositionsCol(eGDCol_CurrentPrice)) = ""
                .TextMatrix(lIndex, PositionsCol(eGDCol_OpenProfit)) = ""
            End If
        Next lIndex
                
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.RefreshPrices2"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeTradeItemName
'' Description: Change the trade item name in the grid
'' Inputs:      Old Name, New Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeTradeItemName(ByVal strOldName As String, ByVal strNewName As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With m_fgPositions
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, PositionsCol(eGDCol_AutoTradeItem)) = strOldName Then
                .TextMatrix(lIndex, PositionsCol(eGDCol_AutoTradeItem)) = strNewName
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.ChangeTradeItemName"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Update
'' Description: Update the data
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Update(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim SnapshotPositions As cAccountPositions ' Copy of the snapshot positions
    Dim bUpdate As Boolean              ' Update the grid?

    bUpdate = False
    Set SnapshotPositions = g.Broker.FillSummaryForBroker(nBroker)
    If Not SnapshotPositions Is Nothing Then
        If m.SnapshotPositions(nBroker) Is Nothing Then
            bUpdate = True
        ElseIf SnapshotPositions.LastModified > m.SnapshotPositions(nBroker).LastModified Then
            bUpdate = True
        End If
    End If
    
    If bUpdate Then
        ' Get the latest collections for the broker...
        Set m.SnapshotPositions(nBroker) = SnapshotPositions
        g.Broker.BrokerDebug nBroker, vbTab & vbTab & "Positions Last Modified = " & Str(SnapshotPositions.LastModified)
        
        With m_fgPositions
            .Redraw = flexRDNone
            
            ' Update the orders grid...
            ResetRemovePositionFlag nBroker
            RefreshBrokerPositions nBroker, False
            RemoveFlaggedPositions nBroker
        
            .Redraw = flexRDBuffered
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.Update"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterPositionsGrid
'' Description: Hide positions that we do not want showing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FilterPositionsGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lFirstChild As Long             ' First child row for the row
    Dim lLastChild As Long              ' Last child row for the row
    Dim AcctPos As cAccountPosition     ' Account Position object
    Dim lParent As Long                 ' Parent row
    Dim lNumVisible As Long             ' Number of children visible
    Dim lVisibleChild As Long           ' Row of the visible child
    Dim lNumExistingStrat As Long       ' Number of existing strategies
    Dim bMismatch As Boolean            ' Are we in a position mismatch?
    Dim lNumPositions As Long           ' Number of symbols in a position
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    lNumPositions = 0&
    With m_fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .BackColor = ALT_GRID_ROW_COLOR
        If g.nColorTheme = kDarkThemeColor Or g.nColorTheme = vbWhite Then
            .BackColorAlternate = g.nColorTheme
        Else
            .BackColorAlternate = &H80000005        'system constant for window background
        End If
        
        ' 1) Restore tree structure...
        For lIndex = .FixedRows To .Rows - 1
            Set AcctPos = .RowData(lIndex)
            If (AcctPos.AutoTradeItemID = -1&) Then
                .RowOutlineLevel(lIndex) = 1
            Else
                .RowOutlineLevel(lIndex) = 2
            End If
        Next lIndex
        
        ' 2) Remove automated trading items that no longer exist and are flat as well as
        '    expired contracts...
        For lIndex = .Rows - 1 To .FixedRows Step -1
            Set AcctPos = .RowData(lIndex)
            
            If IsExpiredContract(AcctPos.SymbolOrSymbolID) = True Then
                .RemoveItem lIndex
            ElseIf (AcctPos.AutoTradeItemID > 0) Then
                If (g.TradingItems.Exists(Str(AcctPos.AutoTradeItemID)) = False) Then
                    If AcctPos.CurrentPositionSnapshot = 0& Then
                        .RemoveItem lIndex
                    End If
                End If
            End If
        Next lIndex
        
        ' 3) Show/Hide appropriate rows...
        For lIndex = .FixedRows To .Rows - 1
            Set AcctPos = .RowData(lIndex)
            lParent = .GetNodeRow(lIndex, flexNTParent)
            
            bMismatch = (g.Broker.PositionMatch(AcctPos.AccountID, AcctPos.SymbolOrSymbolID) = False)
            
            If g.Broker.HideAccount(AcctPos.AccountID) = True Then
                .RowHidden(lIndex) = True
            ElseIf (AcctPos.CurrentPositionSnapshot = 0&) And (bMismatch = False) Then
                .RowHidden(lIndex) = True
            Else
                .RowHidden(lIndex) = False
                If lParent <> -1& Then
                    .RowHidden(lParent) = False
                End If
            End If
        
            If .RowHidden(lIndex) = False Then
                If (AcctPos.AutoTradeItemID = -1&) And (AcctPos.CurrentPositionSnapshot <> 0&) Then
                    lNumPositions = lNumPositions + 1&
                End If
            End If
        Next lIndex
        
        ' 4) Determine if parent should be shown or if entire tree should be shown...
        For lIndex = .FixedRows To .Rows - 1
            If (.RowHidden(lIndex) = False) Then
                lFirstChild = .GetNodeRow(lIndex, flexNTFirstChild)
                lLastChild = .GetNodeRow(lIndex, flexNTLastChild)
                If (lFirstChild <> -1&) Then
                    lNumVisible = 0&
                    lNumExistingStrat = 0&
                    For lIndex2 = lFirstChild To lLastChild
                        If .RowHidden(lIndex2) = False Then
                            lNumVisible = lNumVisible + 1&
                            lVisibleChild = lIndex2
                        End If
                        
                        ' Count up the automated trading strategies if either it is in a position
                        ' or it still exists...
                        If .RowData(lIndex2).AutoTradeItemID > 0 Then
                            If (.RowHidden(lIndex2) = False) Or (g.TradingItems.Exists(Str(.RowData(lIndex2).AutoTradeItemID)) = True) Then
                                lNumExistingStrat = lNumExistingStrat + 1&
                            End If
                        End If
                    Next lIndex2
                    
                    ' If there are no children visible, hide the parent...
                    If lNumVisible = 0& Then
                        .RowHidden(lIndex) = True
                        
                    ' Otherwise, if there is only one child visible and either it is the
                    ' only child or none of the strategies exist anymore, show the child
                    ' and hide the parent...
                    ElseIf (lNumVisible = 1&) And ((lFirstChild = lLastChild) Or (lNumExistingStrat = 0&)) Then
                        .RowHidden(lIndex) = True
                        
                        .TextMatrix(lVisibleChild, PositionsCol(eGDCol_Symbol)) = .TextMatrix(lIndex, PositionsCol(eGDCol_Symbol))
                        .TextMatrix(lVisibleChild, PositionsCol(eGDCol_Account)) = .TextMatrix(lIndex, PositionsCol(eGDCol_Account))
                        .RowOutlineLevel(lVisibleChild) = 1
                        
                    ' Otherwise, if there is one visible child, show the entire tree...
                    ElseIf lNumVisible = 1& Then
                        .RowHidden(lIndex) = False
                        For lIndex2 = lFirstChild To lLastChild
                            .RowHidden(lIndex2) = False
                        Next lIndex2
                        
                    ' Otherwise, show the parent...
                    Else
                        .RowHidden(lIndex) = False
                    End If
                End If
            End If
        Next lIndex

        .Redraw = nRedraw
    End With

    ShowAutoTradeColumns
    SetBackColors m_fgPositions
    m.TCGrid.AutoSizeGrid
    
    If Not g.ConsoleForms Is Nothing Then
        g.ConsoleForms.NumVisible(eGDConsoleForm_Positions) = lNumPositions
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.FilterPositionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyDisplayString
'' Description: Apply the display string from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ApplyDisplayString()
On Error GoTo ErrSection:

    m.TCGrid.ApplyDisplayString eGDCol_Flatten + 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.ApplyDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionForSymbol
'' Description: Get the position for the given symbol, account, auto trade item
'' Inputs:      Symbol, Account, Auto Trade Item
'' Returns:     Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionForSymbol(ByVal vSymbolOrSymbolID As Variant, ByVal lAccountID As Long, ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim AcctPos As cAccountPosition     ' Account Position object
    Dim lReturn As Long                 ' Return value for the function

    With m_fgPositions
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cAccountPosition Then
                Set AcctPos = .RowData(lIndex)
                If (AcctPos.SymbolOrSymbolID = vSymbolOrSymbolID) And (AcctPos.AccountID = lAccountID) And (AcctPos.AutoTradeItemID = lAutoTradeItemID) Then
                    lReturn = AcctPos.CurrentPosition
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    PositionForSymbol = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPositionsUI.PositionForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateConsoleSettings
'' Description: Update the console settings from the configuration form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateConsoleSettings()
On Error GoTo ErrSection:

    ' Update the font on all of the grids...
    FontFromString m_fgPositions.Font, g.Broker.GridFont
    m_fgPositions.Font = m_fgPositions.Font
    
    ' Either show or hide the cents on all of the currency columns...
    If g.Broker.ShowCents Then
        m_fgPositions.ColFormat(PositionsCol(eGDCol_OpenProfit)) = "$#,##0.00"
        m.TCGrid.AutoSizeGrid
    Else
        m_fgPositions.ColFormat(PositionsCol(eGDCol_OpenProfit)) = "$#,##0"
        m.TCGrid.AutoSizeGrid
    End If
    
    ' Update the columns displayed on the grids on the summary tab...
    If m.bSummary Then
        m.strSummaryDisplay = FixDisplay(GetIniFileProperty("SummaryPositionsDisplay", "", "TTSummary", g.strIniFile), True)
        m.TCGrid.Display = m.strSummaryDisplay
        ApplyDisplayString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.UpdateConsoleSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedPosition
'' Description: Grab the selected position in the grid
'' Inputs:      None
'' Returns:     Position (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectedPosition() As cAccountPosition
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Account position from the grid

    Set AcctPos = Nothing
    With m_fgPositions
        If m.TCGrid.ValidRow(.Row) Then
            If TypeOf .RowData(.Row) Is cAccountPosition Then
                Set AcctPos = .RowData(.Row)
            End If
        End If
    End With
    
    Set SelectedPosition = AcctPos

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPositionsUI.SelectedPosition"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionForRow
'' Description: Grab the selected position in the grid for the given row
'' Inputs:      Row
'' Returns:     Position (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionForRow(ByVal Row As Long) As cAccountPosition
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Account position from the grid

    Set AcctPos = Nothing
    With m_fgPositions
        If m.TCGrid.ValidRow(Row) Then
            If TypeOf .RowData(Row) Is cAccountPosition Then
                Set AcctPos = .RowData(Row)
            End If
        End If
    End With
    
    Set PositionForRow = AcctPos

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPositionsUI.PositionForRow"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview form for this grid
'' Inputs:      Arguments passed in from PrintMe
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    m.TCGrid.GenerateReport vArgs, "Positions"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    With m_fgPositions
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .BackColor = ALT_GRID_ROW_COLOR
        .BackColorAlternate = &H80000005 'fgSummaryOrders.BackColor
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbButtonFace
        .Editable = flexEDNone
        .ExplorerBar = flexExMove
        .ExtendLastCol = False
        .HighLight = flexHighlightNever
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        .FrozenCols = GDCol(eGDCol_Flatten) + 1
        .OutlineCol = GDCol(eGDCol_Symbol)
                
        .TextMatrix(0, GDCol(eGDCol_PositionID)) = "Position ID"
        .TextMatrix(0, GDCol(eGDCol_AccountID)) = "Account ID"
        .TextMatrix(0, GDCol(eGDCol_AccountType)) = "Account Type"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_SymbolID)) = "Symbol ID"
        .TextMatrix(0, GDCol(eGDCol_Remove)) = "Remove"
        .TextMatrix(0, GDCol(eGDCol_Position)) = "Position"
        .TextMatrix(0, GDCol(eGDCol_Flatten)) = "F"
        .TextMatrix(0, GDCol(eGDCol_AvgEntry)) = "Avg Entry"
        .TextMatrix(0, GDCol(eGDCol_CurrentPrice)) = "Current Price"
        .TextMatrix(0, GDCol(eGDCol_OpenProfit)) = "Open Equity"
        .TextMatrix(0, GDCol(eGDCol_Account)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_AutoTradeItem)) = "Auto Trade Item"
        .TextMatrix(0, GDCol(eGDCol_OrderStrategy)) = "Auto Exit"
        .TextMatrix(0, GDCol(eGDCol_LastTraded)) = "Last Traded"
        .TextMatrix(0, GDCol(eGDCol_SessionDate)) = "Session Date"
        .TextMatrix(0, GDCol(eGDCol_SessionQuantity)) = "Session Qty"
        .TextMatrix(0, GDCol(eGDCol_SessionProfit)) = "Session Profit"
        
        .ColHidden(GDCol(eGDCol_PositionID)) = True
        .ColHidden(GDCol(eGDCol_AccountID)) = True
        .ColHidden(GDCol(eGDCol_AccountType)) = True
        .ColHidden(GDCol(eGDCol_SymbolID)) = True
        .ColHidden(GDCol(eGDCol_Remove)) = True
        
        .ColDataType(GDCol(eGDCol_Remove)) = flexDTBoolean
        
        If g.Broker.ShowCents Then
            .ColFormat(GDCol(eGDCol_OpenProfit)) = "$#,##0.00"
            .ColFormat(GDCol(eGDCol_SessionProfit)) = "$#,##0.00"
        Else
            .ColFormat(GDCol(eGDCol_OpenProfit)) = "$#,##0"
            .ColFormat(GDCol(eGDCol_SessionProfit)) = "$#,##0"
        End If
        .ColFormat(GDCol(eGDCol_LastTraded)) = DateFormat("FORMAT", MM_DD_YYYY, HH_MM_SS)
        .ColFormat(GDCol(eGDCol_SessionDate)) = DateFormat("FORMAT", MM_DD_YYYY)
        
        .ColAlignment(GDCol(eGDCol_AvgEntry)) = flexAlignRightTop
        .ColAlignment(GDCol(eGDCol_CurrentPrice)) = flexAlignRightTop
        .ColAlignment(GDCol(eGDCol_Account)) = flexAlignLeftTop
        .ColAlignment(GDCol(eGDCol_LastTraded)) = flexAlignCenterTop
        .ColAlignment(GDCol(eGDCol_SessionDate)) = flexAlignCenterTop
        
        .Cell(flexcpFontBold, 0, GDCol(eGDCol_Position)) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadPositionsGrid
'' Description: Load the positions grid with the account positions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadPositionsGrid()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim Bars As cGdBars                 ' Bars object
    Dim nBroker As eTT_AccountType      ' Account type for the account
    Dim AcctPos As cAccountPosition     ' Account position object

    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions];", dbOpenDynaset)
    With m_fgPositions
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
                
        Do While Not rs.EOF
            Set AcctPos = New cAccountPosition
            If AcctPos.Load(rs!AccountPositionID, rs) Then
                If g.Broker.PositionVerify(AcctPos.Broker) = False Then
                    PositionToGrid AcctPos
                End If
            End If
            
            rs.MoveNext
        Loop
        
        FilterPositionsGrid
        m.TCGrid.AutoSizeGrid
        
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.LoadAccountPositionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionToGrid
'' Description: Create a line in the grid for the given position
'' Inputs:      Position, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PositionToGrid(AcctPos As cAccountPosition, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim Bars As cGdBars                 ' Bars object
    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim bMismatch As Boolean            ' Is the symbol in a position mismatch?
    Dim dCurrent As Double              ' Current price for open profit
    Dim strKey As String                ' Key into the position match collection
    Dim dOpenProfit As Double           ' Open profit
    
    With m_fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
    
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
            
            For lIndex = .FixedRows To .Rows - 2
                If .RowData(lIndex).SortKey > AcctPos.SortKey Then
                    .RowPosition(lRow) = lIndex
                    lRow = lIndex
                    Exit For
                End If
            Next lIndex
        End If
        
        bMismatch = (g.Broker.PositionMatch(AcctPos.AccountID, AcctPos.SymbolOrSymbolID) = False)
        strKey = Str(AcctPos.AccountID) & "|" & Str(AcctPos.SymbolOrSymbolID)
        If m.PositionMatch.Exists(strKey) Then
            m.PositionMatch(strKey) = Not bMismatch
        Else
            m.PositionMatch.Add (Not bMismatch), strKey
        End If
    
        .RowData(lRow) = AcctPos
        
        .TextMatrix(lRow, PositionsCol(eGDCol_Account)) = g.Broker.AccountNameForID(AcctPos.AccountID)
        .TextMatrix(lRow, PositionsCol(eGDCol_AccountID)) = Str(AcctPos.AccountID)
        .TextMatrix(lRow, PositionsCol(eGDCol_AccountType)) = Str(AcctPos.Broker)
        CheckedCell(m_fgPositions, lRow, PositionsCol(eGDCol_Remove)) = False
        
        Select Case AcctPos.AutoTradeItemID
            Case Is > 0
                .TextMatrix(lRow, PositionsCol(eGDCol_AutoTradeItem)) = AutoTradeItemNameForID(AcctPos.AutoTradeItemID)
            Case 0
                .TextMatrix(lRow, PositionsCol(eGDCol_AutoTradeItem)) = "Manual"
            Case -1
                .TextMatrix(lRow, PositionsCol(eGDCol_AutoTradeItem)) = "Total"
        End Select
        If (AcctPos.CurrentPositionSnapshot <> 0) And (bMismatch = False) Then
            .TextMatrix(lRow, PositionsCol(eGDCol_AvgEntry)) = AcctPos.AverageEntrySnapshotString
        Else
            .TextMatrix(lRow, PositionsCol(eGDCol_AvgEntry)) = ""
        End If
        If AcctPos.LastTradedSnapshot > 0 Then
            .TextMatrix(lRow, PositionsCol(eGDCol_LastTraded)) = ConvertBrokerDate(AcctPos.LastTradedSnapshot, g.Broker.AccountTypeForID(AcctPos.AccountID), AcctPos.Symbol, g.bShowInLocalTimeZone)
        Else
            .TextMatrix(lRow, PositionsCol(eGDCol_LastTraded)) = ""
        End If
        If AcctPos.AutoTradeItemID = 0& Then
            .TextMatrix(lRow, PositionsCol(eGDCol_OrderStrategy)) = g.OrderStrategies.ExitForAccountAndSymbol(AcctPos.AccountID, AcctPos.SymbolOrSymbolID)
        Else
            .TextMatrix(lRow, PositionsCol(eGDCol_OrderStrategy)) = ""
        End If
        If bMismatch = False Then
            .TextMatrix(lRow, PositionsCol(eGDCol_Position)) = AcctPos.CurrentPositionSnapshotString
        Else
            .TextMatrix(lRow, PositionsCol(eGDCol_Position)) = "Mismatch"
        End If
        .TextMatrix(lRow, PositionsCol(eGDCol_Flatten)) = "F"
        .Cell(flexcpFontUnderline, lRow, PositionsCol(eGDCol_Flatten)) = True
        If g.nColorTheme = kDarkThemeColor Then
            .Cell(flexcpForeColor, lRow, PositionsCol(eGDCol_Flatten)) = vbCyan
        Else
            .Cell(flexcpForeColor, lRow, PositionsCol(eGDCol_Flatten)) = vbBlue
        End If
        .TextMatrix(lRow, PositionsCol(eGDCol_PositionID)) = Str(AcctPos.AccountPositionID)
        .TextMatrix(lRow, PositionsCol(eGDCol_Symbol)) = AcctPos.Symbol
        .TextMatrix(lRow, PositionsCol(eGDCol_SymbolID)) = Str(AcctPos.SymbolID)
        
        If AcctPos.SessionDateSnapshot = 0& Then
            .TextMatrix(lRow, PositionsCol(eGDCol_SessionDate)) = ""
            .TextMatrix(lRow, PositionsCol(eGDCol_SessionQuantity)) = ""
            .TextMatrix(lRow, PositionsCol(eGDCol_SessionProfit)) = ""
        Else
            .TextMatrix(lRow, PositionsCol(eGDCol_SessionDate)) = AcctPos.SessionDateSnapshot
            .TextMatrix(lRow, PositionsCol(eGDCol_SessionQuantity)) = Format(AcctPos.SessionQuantitySnapshot, "#,##0")
            .TextMatrix(lRow, PositionsCol(eGDCol_SessionProfit)) = AcctPos.SessionProfitSnapshot
            m.TCGrid.ColorCell lRow, PositionsCol(eGDCol_SessionProfit)
        End If
        
        If (AcctPos.CurrentPositionSnapshot <> 0) And (bMismatch = False) Then
            Set Bars = GetTradeBars(AcctPos.SymbolOrSymbolID)
            If Not Bars Is Nothing Then
                dCurrent = AcctPos.CurrentPrice(Bars(eBARS_Close, Bars.Size - 1), Bars(eBARS_Bid, Bars.Size - 1), Bars(eBARS_Ask, Bars.Size - 1))
                
                .TextMatrix(lRow, PositionsCol(eGDCol_CurrentPrice)) = Bars.PriceDisplay(dCurrent)
                If AcctPos.CurrentPositionSnapshot <> 0 Then
                    dOpenProfit = AcctPos.OpenProfit(dCurrent)
                    If dOpenProfit <> kNullData Then
                        .TextMatrix(lRow, PositionsCol(eGDCol_OpenProfit)) = dOpenProfit
                        m.TCGrid.ColorCell lRow, PositionsCol(eGDCol_OpenProfit)
                    Else
                        .TextMatrix(lRow, PositionsCol(eGDCol_OpenProfit)) = ""
                    End If
                Else
                    .TextMatrix(lRow, PositionsCol(eGDCol_OpenProfit)) = ""
                End If
            Else
                .TextMatrix(lRow, PositionsCol(eGDCol_CurrentPrice)) = ""
                .TextMatrix(lRow, PositionsCol(eGDCol_OpenProfit)) = ""
            End If
        Else
            .TextMatrix(lRow, PositionsCol(eGDCol_CurrentPrice)) = ""
            .TextMatrix(lRow, PositionsCol(eGDCol_OpenProfit)) = ""
        End If
        
        If (AcctPos.CurrentPositionSnapshot = 0) Or (bMismatch = True) Then
            .Cell(flexcpForeColor, lRow, PositionsCol(eGDCol_Symbol), lRow, PositionsCol(eGDCol_Position)) = .ForeColor
            .Cell(flexcpFontBold, lRow, PositionsCol(eGDCol_Symbol), lRow, PositionsCol(eGDCol_Position)) = False
        ElseIf AcctPos.CurrentPositionSnapshot > 0 Then
            .Cell(flexcpForeColor, lRow, PositionsCol(eGDCol_Symbol), lRow, PositionsCol(eGDCol_Position)) = g.ChartGlobals.nLongColor
            .Cell(flexcpFontBold, lRow, PositionsCol(eGDCol_Symbol), lRow, PositionsCol(eGDCol_Position)) = True
        ElseIf AcctPos.CurrentPositionSnapshot < 0 Then
            .Cell(flexcpForeColor, lRow, PositionsCol(eGDCol_Symbol), lRow, PositionsCol(eGDCol_Position)) = g.ChartGlobals.nShortColor
            .Cell(flexcpFontBold, lRow, PositionsCol(eGDCol_Symbol), lRow, PositionsCol(eGDCol_Position)) = True
        End If
        
        .IsSubtotal(lRow) = True
        If AcctPos.AutoTradeItemID = -1& Then
            .RowOutlineLevel(lRow) = 1
        Else
            .RowOutlineLevel(lRow) = 2
            .TextMatrix(lRow, PositionsCol(eGDCol_Symbol)) = ""
            .TextMatrix(lRow, PositionsCol(eGDCol_Account)) = ""
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.PositionToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetRemovePositionFlag
'' Description: Set the remove position flag on all positions for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetRemovePositionFlag(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With m_fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If .RowData(lIndex).Broker = nBroker Then
                CheckedCell(m_fgPositions, lIndex, PositionsCol(eGDCol_Remove)) = True
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.ResetRemovePositionFlag"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFlaggedPositions
'' Description: Remove rows with the remove flag set for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveFlaggedPositions(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With m_fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If .RowData(lIndex).Broker = nBroker Then
                If CheckedCell(m_fgPositions, lIndex, PositionsCol(eGDCol_Remove)) = True Then
                    .RemoveItem lIndex
                End If
            End If
        Next lIndex
        
        FilterPositionsGrid
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.RemoveFlaggedPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshBrokerPositions
'' Description: Refresh positions in the grid for the given broker from the collection
'' Inputs:      Broker, Filter Grid?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshBrokerPositions(ByVal nBroker As eTT_AccountType, Optional ByVal bFilterGrid As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    Dim Positions As cAccountPositions  ' Collection of account positions
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    If Not m.SnapshotPositions(nBroker) Is Nothing Then
        Set Positions = m.SnapshotPositions(nBroker).MakeCopy
        
        With m_fgPositions
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            For lIndex = 1 To Positions.Count
                lRow = -1&
                For lIndex2 = .FixedRows To .Rows - 1
                    If .RowData(lIndex2).AccountPositionID = Positions(lIndex).AccountPositionID Then
                        lRow = lIndex2
                        Exit For
                    End If
                Next lIndex2
                
                If IsExpiredContract(Positions(lIndex).SymbolOrSymbolID) = False Then
                    PositionToGrid Positions(lIndex), lRow
                End If
            Next lIndex
            
            .Redraw = nRedraw
        End With
        
        If bFilterGrid Then
            FilterPositionsGrid
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.RefreshBrokerPositions"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionsCol
'' Description: Determine where the given column is really sitting right now
'' Inputs:      Column Number
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PositionsCol(ByVal Col As eGDCols) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    If Col = eGDCol_NumCols Then
        lReturn = Col
    Else
        lReturn = m.TCGrid.ColNumFromColData(Col)
    End If
    
    PositionsCol = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPositionsUI.PositionsCol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SmartColumns
'' Description: Determine what columns should be shown/hidden based on the
''              current screen width
'' Inputs:      None
'' Returns:     Display string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SmartColumns() As String
On Error GoTo ErrSection:

    Dim astrDisplay As New cGdArray     ' Array of field display information

    Select Case Screen.Width / Screen.TwipsPerPixelX
        Case Is >= 1600
            'strReturn = "1;Avg Entry;1;7,1;Current Price;1;8,1;Open Equity;1;9,1;Auto Exit;1;10,1;Last Traded;1;11,1;Session Date;1;12,1;Session Qty;1;13,1;Session Profit;1;14"
            astrDisplay.Add "0;Position ID;1;0;0;0"
            astrDisplay.Add "0;Account ID;1;1;0;0"
            astrDisplay.Add "0;Account Type;1;2;0;0"
            astrDisplay.Add "0;Symbol ID;1;3;0;0"
            astrDisplay.Add "0;Remove;1;4;0;0"
            astrDisplay.Add "1;Symbol;1;5;0;0"
            astrDisplay.Add "1;Account;1;6;0;0"
            astrDisplay.Add "1;Auto Trade Item;1;7;0;0"
            astrDisplay.Add "1;Position;1;8;0;0"
            astrDisplay.Add "1;F;1;9;0;0"
            astrDisplay.Add "1;Avg Entry;1;10;0;1"
            astrDisplay.Add "1;Current Price;1;11;0;1"
            astrDisplay.Add "1;Open Equity;1;12;0;1"
            astrDisplay.Add "1;Auto Exit;1;13;0;1"
            astrDisplay.Add "1;Last Traded;1;14;0;1"
            astrDisplay.Add "0;Session Date;1;15;0;1"
            astrDisplay.Add "0;Session Qty;1;16;0;1"
            astrDisplay.Add "0;Session Profit;1;17;0;1"
        
        Case Is >= 1280
            'strReturn = "1;Avg Entry;1;7,1;Current Price;1;8,1;Open Equity;1;9,1;Auto Exit;1;10,1;Last Traded;1;11,1;Session Date;1;12,1;Session Qty;1;13,1;Session Profit;1;14"
            astrDisplay.Add "0;Position ID;1;0;0;0"
            astrDisplay.Add "0;Account ID;1;1;0;0"
            astrDisplay.Add "0;Account Type;1;2;0;0"
            astrDisplay.Add "0;Symbol ID;1;3;0;0"
            astrDisplay.Add "0;Remove;1;4;0;0"
            astrDisplay.Add "1;Symbol;1;5;0;0"
            astrDisplay.Add "1;Account;1;6;0;0"
            astrDisplay.Add "1;Auto Trade Item;1;7;0;0"
            astrDisplay.Add "1;Position;1;8;0;0"
            astrDisplay.Add "1;F;1;9;0;0"
            astrDisplay.Add "1;Avg Entry;1;10;0;1"
            astrDisplay.Add "1;Current Price;1;11;0;1"
            astrDisplay.Add "1;Open Equity;1;12;0;1"
            astrDisplay.Add "1;Auto Exit;1;13;0;1"
            astrDisplay.Add "0;Last Traded;1;14;0;1"
            astrDisplay.Add "0;Session Date;1;15;0;1"
            astrDisplay.Add "0;Session Qty;1;16;0;1"
            astrDisplay.Add "0;Session Profit;1;17;0;1"
        
        Case Is >= 1152
            'strReturn = "0;Avg Entry;1;7,1;Current Price;1;8,1;Open Equity;1;9,1;Auto Exit;1;10,1;Last Traded;1;11,1;Session Date;1;12,1;Session Qty;1;13,1;Session Profit;1;14"
            astrDisplay.Add "0;Position ID;1;0;0;0"
            astrDisplay.Add "0;Account ID;1;1;0;0"
            astrDisplay.Add "0;Account Type;1;2;0;0"
            astrDisplay.Add "0;Symbol ID;1;3;0;0"
            astrDisplay.Add "0;Remove;1;4;0;0"
            astrDisplay.Add "1;Symbol;1;5;0;0"
            astrDisplay.Add "1;Account;1;6;0;0"
            astrDisplay.Add "1;Auto Trade Item;1;7;0;0"
            astrDisplay.Add "1;Position;1;8;0;0"
            astrDisplay.Add "1;F;1;9;0;0"
            astrDisplay.Add "0;Avg Entry;1;10;0;1"
            astrDisplay.Add "1;Current Price;1;11;0;1"
            astrDisplay.Add "1;Open Equity;1;12;0;1"
            astrDisplay.Add "1;Auto Exit;1;13;0;1"
            astrDisplay.Add "0;Last Traded;1;14;0;1"
            astrDisplay.Add "0;Session Date;1;15;0;1"
            astrDisplay.Add "0;Session Qty;1;16;0;1"
            astrDisplay.Add "0;Session Profit;1;17;0;1"
        
        Case Else
            'strReturn = "0;Avg Entry;1;7,1;Current Price;1;8,1;Open Equity;1;9,1;Auto Exit;1;10,1;Last Traded;1;11,1;Session Date;1;12,1;Session Qty;1;13,1;Session Profit;1;14"
            astrDisplay.Add "0;Position ID;1;0;0;0"
            astrDisplay.Add "0;Account ID;1;1;0;0"
            astrDisplay.Add "0;Account Type;1;2;0;0"
            astrDisplay.Add "0;Symbol ID;1;3;0;0"
            astrDisplay.Add "0;Remove;1;4;0;0"
            astrDisplay.Add "1;Symbol;1;5;0;0"
            astrDisplay.Add "1;Account;1;6;0;0"
            astrDisplay.Add "1;Auto Trade Item;1;7;0;0"
            astrDisplay.Add "1;Position;1;8;0;0"
            astrDisplay.Add "1;F;1;9;0;0"
            astrDisplay.Add "0;Avg Entry;1;10;0;1"
            astrDisplay.Add "1;Current Price;1;11;0;1"
            astrDisplay.Add "1;Open Equity;1;12;0;1"
            astrDisplay.Add "1;Auto Exit;1;13;0;1"
            astrDisplay.Add "0;Last Traded;1;14;0;1"
            astrDisplay.Add "0;Session Date;1;15;0;1"
            astrDisplay.Add "0;Session Qty;1;16;0;1"
            astrDisplay.Add "0;Session Profit;1;17;0;1"
    
    End Select
    
    SmartColumns = "1|" & astrDisplay.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPositionsUI.SmartColumns"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColumnExceptions
'' Description: Build a list of columns not included in the display string
'' Inputs:      None
'' Returns:     Exception String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ColumnExceptions() As String
On Error GoTo ErrSection:

    Dim astrColumns As cGdArray         ' Array of column numbers to join
    
    Set astrColumns = New cGdArray
    astrColumns.Create eGDARRAY_Strings
    
    astrColumns.Add Str(PositionsCol(eGDCol_PositionID))
    astrColumns.Add Str(PositionsCol(eGDCol_AccountID))
    astrColumns.Add Str(PositionsCol(eGDCol_SymbolID))
    astrColumns.Add Str(PositionsCol(eGDCol_Symbol))
    astrColumns.Add Str(PositionsCol(eGDCol_Account))
    astrColumns.Add Str(PositionsCol(eGDCol_AutoTradeItem))
    astrColumns.Add Str(PositionsCol(eGDCol_Position))
    astrColumns.Add Str(PositionsCol(eGDCol_Flatten))
    astrColumns.Add Str(PositionsCol(eGDCol_AccountType))
    astrColumns.Add Str(PositionsCol(eGDCol_Remove))
    
    ColumnExceptions = "," & astrColumns.JoinFields(",") & ","

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPositionsUI.ColumnExceptions"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Setup the popup menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bMismatch As Boolean            ' Is the symbol in a position mismatch?
    Dim AcctPos As cAccountPosition     ' Account position object
        
    Set AcctPos = SelectedPosition
    If Not AcctPos Is Nothing Then
        bMismatch = (g.Broker.PositionMatch(AcctPos.AccountID, AcctPos.SymbolOrSymbolID) = False)
    
        m_mnuFlatten.Enabled = ((AcctPos.CurrentPositionSnapshot <> 0&) Or (bMismatch = True))
        m_mnuReverse.Enabled = ((AcctPos.AutoTradeItemID = 0&) And ((AcctPos.CurrentPositionSnapshot <> 0&) Or (bMismatch = True)))
        
        m_mnuManageXOS.Enabled = (AcctPos.AutoTradeItemID = 0&)
        m_mnuSelectXOS.Enabled = (AcctPos.AutoTradeItemID = 0&)
        m_mnuRemoveXOS.Enabled = (AcctPos.AutoTradeItemID = 0&) And (Len(m_fgPositions.TextMatrix(m_fgPositions.Row, PositionsCol(eGDCol_OrderStrategy))) > 0)
        
        m_mnuActualPerformance.Enabled = True
    Else
        m_mnuFlatten.Enabled = False
        m_mnuReverse.Enabled = False
        
        m_mnuManageXOS.Enabled = True
        m_mnuSelectXOS.Enabled = False
        m_mnuRemoveXOS.Enabled = False
    
        m_mnuActualPerformance.Enabled = False
    End If
    
    m_mnuCheckStatus.Enabled = g.Broker.EnableCheckStatusMenu
    m_mnuDefaultColumns.Enabled = (m.TCGrid.Display <> DefaultDisplay)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowAutoTradeColumns
'' Description: Show the auto trade columns as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowAutoTradeColumns()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was there an auto trade item found?
    Dim Position As cAccountPosition    ' Position object
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With m_fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        bFound = False
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then
                Set Position = PositionForRow(lIndex)
                If Not Position Is Nothing Then
                    If Position.AutoTradeItemID > 0 Then
                        bFound = True
                        Exit For
                    End If
                End If
            End If
        Next lIndex
        
        .ColHidden(PositionsCol(eGDCol_AutoTradeItem)) = Not bFound
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.ShowAutoTradeColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_AfterMoveColumn
'' Description: After a user moves a column, resave the display string
'' Inputs:      Column moved, Position moved to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_AfterMoveColumn"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_AfterRowColChange
'' Description: Enable/Disable controls as appropriate as the user moves
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With m_fgPositions
        'if no horiz scrollbars, we need to "manually scroll"
        'by forcing the cell to show when changing columns
        If .Redraw <> flexRDNone And .ScrollBars <> flexScrollBarBoth Then
            If OldCol <> NewCol Then
                If (NewRow >= 0) And (NewCol >= 0) Then
                    .ShowCell NewRow, NewCol
                End If
            End If
        End If
    End With

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_AfterScroll
'' Description: Handle the user scrolling the grid
'' Inputs:      Old Top Row, Old Left Columns, New Top Row, New Left Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error GoTo ErrSection:

    ' DAJ 08/05/2014: Since we are only updating the open equity on visible
    ' rows, when the user scrolls, the open equity won't update on newly visible
    ' rows after the next tick.  Instead, do a refresh prices here to get all
    ' visible rows updated again...
    If NewTopRow <> OldTopRow Then
        RefreshPrices2 True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_AfterScroll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_AfterSort
'' Description: Set the background colors after a sort
'' Inputs:      Column of Sort, Order of Sort
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors m_fgPositions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_AfterUserResize
'' Description: Rebuild the display string after a user resize
'' Inputs:      Row and Column of resize
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_AfterUserResize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_BeforeCollapse
'' Description: Don't allow the user to expand or collapse the tree
'' Inputs:      Row, State, Whether to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
On Error GoTo ErrSection:

    Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_BeforeCollapse"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_BeforeMouseDown
'' Description: Bring up the Pop-Up menu when the user right clicks
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    m.TCGrid.SelectRow m_fgPositions.MouseRow

    If Button = vbRightButton Then
        EnableControls
        m.frm.PopupMenu m.mnuPositions
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_BeforeMoveColumn
'' Description: Make sure that the frozen columns stay where they are
'' Inputs:      Column to move, Position to move it to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        With m_fgPositions
            If Col < .FrozenCols Then
                Position = Col
            ElseIf Position < .FrozenCols Then
                Position = .FrozenCols
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_BeforeMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_BeforeScroll
'' Description: Make sure left col stays the same if no horizontal
'' Inputs:      Old Top Row, Old Left Col, New Top Row, New Left Col, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    GridScrollCheck m_fgPositions, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_Click
'' Description: When the user clicks in the "F" column, attempt to flatten
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_Click()
On Error GoTo ErrSection:

    Dim lMouseCol As Long               ' Current mouse column in the grid
    Dim lMouseRow As Long               ' Current mouse row in the grid
    
    With m_fgPositions
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        If (lMouseCol = PositionsCol(eGDCol_Flatten)) Then
            If m.TCGrid.SelectRow(lMouseRow) Then
                FlattenPositionFromGrid m_fgPositions, m.strSource
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_DblClick
'' Description: If the user double clicks on a valid position, bring up the
''              order form defaulting to close out the selected position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_DblClick()
On Error GoTo ErrSection:

    With m_fgPositions
        If m.TCGrid.ValidRow(.Row) Then
            SetActiveChartSymbol .TextMatrix(.Row, PositionsCol(eGDCol_Symbol))
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_KeyUp
'' Description: Delete allows the user to close the position
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            With m_fgPositions
                If m.TCGrid.ValidRow(.Row) Then
                    FlattenPositionFromGrid m_fgPositions, m.strSource
                End If
            End With
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_fgPositions_KeyUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgPositions_MouseMove
'' Description: Set the tooltip text based on where the cursor is
'' Inputs:      Mouse Button Down, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgPositions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseCol As Long               ' Column that the mouse is over
    Dim lMouseRow As Long               ' Row that the mouse is over
    Dim strTooltipText As String        ' Tooltip text
    
    strTooltipText = ""
    With m_fgPositions
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        If m.TCGrid.ValidRow(lMouseRow) Then
            If lMouseCol = PositionsCol(eGDCol_Flatten) Then
                strTooltipText = "Click here to Flatten the position"
            End If
        End If
        
        .ToolTipText = strTooltipText
    End With
    
    m.TCGrid.SetMouseCursor m.frm

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuActualPerformance_Click
'' Description: Allow the user to view the performance reports on their trades
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuActualPerformance_Click()
On Error GoTo ErrSection:

    StartMenuTimer "ACTUALPERFORMANCE"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuActualPerformance_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuAutoSizeColumns_Click
'' Description: Allow the user to auto size the columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuAutoSizeColumns_Click()
On Error GoTo ErrSection:

    m.TCGrid.AutoSizeGrid True
    BuildDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuAutoSizeColumns_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuCheckStatus_Click
'' Description: Check the status of any open orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuCheckStatus_Click()
On Error GoTo ErrSection:

    g.Broker.CheckTradeServerOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuCheckStatus_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuDefaultColumns_Click
'' Description: Allow the user to revert back to the default column information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuDefaultColumns_Click()
On Error GoTo ErrSection:

    If m.bSummary Then
        m.strSummaryDisplay = DefaultDisplay
        m.TCGrid.Display = m.strSummaryDisplay
    Else
        m.strDisplay = DefaultDisplay
        m.TCGrid.Display = m.strDisplay
    End If
    
    ApplyDisplayString
    BuildDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuDefaultColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuFlatten_Click
'' Description: Allow the user to close the selected position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuFlatten_Click()
On Error GoTo ErrSection:

    FlattenPositionFromGrid m_fgPositions, m.strSource

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuFlatten_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuManageXOS_Click
'' Description: Allow the user to manage exit order strategies
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuManageXOS_Click()
On Error GoTo ErrSection:

    StartMenuTimer "MANAGEXOS"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuManageXOS_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuPrint_Click
'' Description: Allow the user to print the orders, positions, and accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuPrint_Click()
On Error GoTo ErrSection:

    m.frm.PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuRemoveXOS_Click
'' Description: Allow the user to remove the exit order strategy from the position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuRemoveXOS_Click()
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Selected position object
    
    Set AcctPos = SelectedPosition
    If Not AcctPos Is Nothing Then
        g.OrderStrategies.DeactivateExit AcctPos.AccountID, AcctPos.SymbolOrSymbolID, True, "Turned off from " & m.strSource
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuRemoveXOS_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuReverse_Click
'' Description: Allow the user to close the selected position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuReverse_Click()
On Error GoTo ErrSection:

    ReversePositionFromGrid m_fgPositions, m.strSource

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuReverse_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSelectXOS_Click
'' Description: Allow the user to select an exit order strategy for the position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuSelectXOS_Click()
On Error GoTo ErrSection:

    StartMenuTimer "SELECTXOS"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuSelectXOS_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSettings_Click
'' Description: Allow the user to view the Trade Console settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuSettings_Click()
On Error GoTo ErrSection:

    StartMenuTimer "SETTINGS"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuSettings_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuTradeHistory_Click
'' Description: Allow the user to view trade history
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuTradeHistory_Click()
On Error GoTo ErrSection:

    StartMenuTimer "TRADEHISTORY"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuTradeHistory_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuViewJournals_Click
'' Description: Allow the user to view their journal entries
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuViewJournals_Click()
On Error GoTo ErrSection:

    StartMenuTimer "JOURNALS"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_mnuViewJournals_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_tmrMenu_Timer
'' Description: Do actions that require a modal dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Action to perform
    Dim Position As cAccountPosition    ' Account position object

    strTag = m_tmrMenu.Tag
    If Len(strTag) > 2 Then
        If Left(strTag, 2) = kMenuPrefix Then
            strTag = Mid(strTag, 3)
            m_tmrMenu.Tag = ""
            m_tmrMenu.Enabled = False
            
            Set Position = SelectedPosition
            
            Select Case UCase(strTag)
                Case "ACTUALPERFORMANCE"
                    ShowTradeFilter Position.AccountID, Position.SymbolOrSymbolID, Position.AutoTradeItemID
                
                Case "JOURNALS"
                    g.TnJournal.ShowJournals
                    
                Case "MANAGEXOS"
                    frmOrderStrategies.ShowMeManage
                    
                Case "SELECTXOS"
                    If Not Position Is Nothing Then
                        SelectXOS Position.AccountID, Position.SymbolOrSymbolID
                    End If
                    
                Case "SETTINGS"
                    frmTTSummaryCfg.ShowMe
                    
                Case "TRADEHISTORY"
                    If Position Is Nothing Then
                        frmTTAccounts.ShowMe True
                    Else
                        frmTTPositions.ShowMe Position.AccountID, Position.Broker
                    End If
                    
            End Select
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.m_tmrMenu_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DefaultDisplay
'' Description: Determine the default display
'' Inputs:      None
'' Returns:     Default Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DefaultDisplay() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrDisplay As New cGdArray     ' Array of field display information
    
    If m.bSummary Then
        strReturn = SmartColumns
    End If
    
    If Len(strReturn) = 0 Then
        astrDisplay.Add "0;Position ID;1;0;0;0"
        astrDisplay.Add "0;Account ID;1;1;0;0"
        astrDisplay.Add "0;Account Type;1;2;0;0"
        astrDisplay.Add "0;Symbol ID;1;3;0;0"
        astrDisplay.Add "0;Remove;1;4;0;0"
        astrDisplay.Add "1;Symbol;1;5;0;0"
        astrDisplay.Add "1;Account;1;6;0;0"
        astrDisplay.Add "1;Auto Trade Item;1;7;0;0"
        astrDisplay.Add "1;Position;1;8;0;0"
        astrDisplay.Add "1;F;1;9;0;0"
        astrDisplay.Add "1;Avg Entry;1;10;0;1"
        astrDisplay.Add "1;Current Price;1;11;0;1"
        astrDisplay.Add "1;Open Equity;1;12;0;1"
        astrDisplay.Add "1;Auto Exit;1;13;0;1"
        astrDisplay.Add "1;Last Traded;1;14;0;1"
        astrDisplay.Add "1;Session Date;1;15;0;1"
        astrDisplay.Add "1;Session Qty;1;16;0;1"
        astrDisplay.Add "1;Session Profit;1;17;0;1"

        'strReturn = "1;Avg Entry;1;7,1;Current Price;1;8,1;Open Equity;1;9,1;Auto Exit;1;10,1;Last Traded;1;11,1;Session Date;1;12,1;Session Qty;1;13,1;Session Profit;1;14"
        strReturn = "1|" & astrDisplay.JoinFields(",")
    End If

    DefaultDisplay = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPositionsUI.DefaultDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixDisplay
'' Description: Fix the given display string
'' Inputs:      Display, Summary?
'' Returns:     Fixed Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixDisplay(ByVal strDisplay As String, ByVal bSummary As Boolean) As String
On Error GoTo ErrSection:

    Dim astrDisplay As cGdArray         ' Display string split out into an array
    Dim astrLine As cGdArray            ' Line of information split out into an array
    Dim lIndex As Long                  ' Index into a for loop
    
    If Len(strDisplay) = 0 Then
        strDisplay = DefaultDisplay
    Else
        strDisplay = Replace(strDisplay, ";Open Position;", ";Position;")
        If InStr(strDisplay, ";Auto Exit;") = 0 Then
            strDisplay = strDisplay & ",1;Auto Exit;1;" & Str(GDCol(eGDCol_OrderStrategy))
        End If
        If InStr(strDisplay, ";Last Traded;") = 0 Then
            strDisplay = strDisplay & ",1;Last Traded;1;" & Str(GDCol(eGDCol_LastTraded))
        End If
        If InStr(strDisplay, ";Session Date;") = 0 Then
            strDisplay = strDisplay & ",1;Session Date;1;" & Str(GDCol(eGDCol_SessionDate))
        End If
        If InStr(strDisplay, ";Session Qty;") = 0 Then
            strDisplay = strDisplay & ",1;Session Qty;1;" & Str(GDCol(eGDCol_SessionQuantity))
        End If
        If InStr(strDisplay, ";Session Profit;") = 0 Then
            strDisplay = strDisplay & ",1;Session Profit;1;" & Str(GDCol(eGDCol_SessionProfit))
        End If
    End If
    
    FixDisplay = strDisplay

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPositionsUI.FixDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildDisplayString
'' Description: Tell the grid class to build the display string
'' Inputs:      Force Sum Calc?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildDisplayString(Optional ByVal bForceSumCalc As Boolean = False)
On Error GoTo ErrSection:

    m.TCGrid.BuildDisplayString ColumnExceptions, bForceSumCalc
    If m.bSummary Then
        m.strSummaryDisplay = m.TCGrid.Display
        SetIniFileProperty "SummaryPositionsDisplay", m.strSummaryDisplay, "TTSummary", g.strIniFile
    Else
        m.strDisplay = m.TCGrid.Display
        SetIniFileProperty "PositionsDisplay", m.strDisplay, "TTSummary", g.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.BuildDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartMenuTimer
'' Description: Start the menu timer with the given command
'' Inputs:      Command
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StartMenuTimer(ByVal strCommand As String)
On Error GoTo ErrSection:

    m_tmrMenu.Tag = kMenuPrefix & strCommand
    m_tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    Set m.SnapshotPositions = New cGdTree

    For lIndex = 1 To kNumBrokers
        Set m.SnapshotPositions(lIndex) = Nothing
    Next lIndex
    
    Set m.LastPrice = New cGdTree
    Set m.PositionMatch = New cGdTree
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    BuildDisplayString
    Set m.SnapshotPositions = Nothing
    Set m.LastPrice = Nothing
    Set m.PositionMatch = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsUI.Class_Terminate"
    
End Sub
