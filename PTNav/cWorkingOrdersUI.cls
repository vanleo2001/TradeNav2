VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWorkingOrdersUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cWorkingOrdersUI.cls
'' Description: Wrapper class for a working orders grid
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
'' 03/17/2010   DAJ         Removed HasWorkingOrders call
'' 03/17/2010   DAJ         Toggle show/hide of auto trade item column
'' 04/15/2010   DAJ         Ensure that the Manage XOS menu item is always enabled
'' 06/03/2010   DAJ         Changes for new TradeSense Order Groups
'' 06/15/2010   DAJ         Removed TradeSense order group menu item
'' 09/13/2010   DAJ         Show TradeSense order groups in working orders grids
'' 10/26/2010   DAJ         Added the StatusCol property
'' 10/26/2010   DAJ         Efficiency enhancements
'' 05/11/2011   DAJ         Utilize CheckSimTradeStatus, EnableCheckStatusMenu
'' 06/28/2011   DAJ         Setup clickable cells like hyperlinks
'' 09/23/2011   DAJ         Show date journals form instead of old journals form
'' 10/04/2011   DAJ         Call the ShowJournals function instead of calling the form direct
'' 02/14/2012   DAJ         Added multi-leg order support
'' 04/02/2012   DAJ         Set the RowOutlineLevel for TradeSense order groups
'' 08/17/2012   DAJ         Allowed sorting on the Trade Console orders grids
'' 10/10/2012   DAJ         Added Time-In-Force column
'' 10/29/2012   DAJ         Employ Tim's new fix for grid scrolling vs. streaming issue
'' 11/28/2012   DAJ         Speed enhancements for the Trade Console
'' 03/21/2013   DAJ         Don't clear updated colors on a hidden row
'' 07/30/2013   DAJ         Data Pending Order status for Conditional Orders
'' 08/01/2013   DAJ         Change to whether or not order is considered conditional
'' 10/10/2013   DAJ         Moved conditional orders add to mTradeTracker.RefreshOrder
'' 03/05/2014   DAJ         Removed unused reference to cLotColumn
'' 05/16/2014   DAJ         Fix for bid/ask/last not updating on TradeSense order group row
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kMenuPrefix = "O:"

Private WithEvents m_fgWorkingOrders As VSFlexGrid
Attribute m_fgWorkingOrders.VB_VarHelpID = -1
Private WithEvents m_tmrMenu As Timer
Attribute m_tmrMenu.VB_VarHelpID = -1
Private WithEvents m_mnuBuy As Menu
Attribute m_mnuBuy.VB_VarHelpID = -1
Private WithEvents m_mnuSell As Menu
Attribute m_mnuSell.VB_VarHelpID = -1
Private WithEvents m_mnuOrderGroups As Menu
Attribute m_mnuOrderGroups.VB_VarHelpID = -1
Private WithEvents m_mnuEditOrder As Menu
Attribute m_mnuEditOrder.VB_VarHelpID = -1
Private WithEvents m_mnuCancelOrder As Menu
Attribute m_mnuCancelOrder.VB_VarHelpID = -1
Private WithEvents m_mnuParkOrder As Menu
Attribute m_mnuParkOrder.VB_VarHelpID = -1
Private WithEvents m_mnuSubmitOrder As Menu
Attribute m_mnuSubmitOrder.VB_VarHelpID = -1
Private WithEvents m_mnuSubmitAll As Menu
Attribute m_mnuSubmitAll.VB_VarHelpID = -1
Private WithEvents m_mnuOrderHistory As Menu
Attribute m_mnuOrderHistory.VB_VarHelpID = -1
Private WithEvents m_mnuNewJournal As Menu
Attribute m_mnuNewJournal.VB_VarHelpID = -1
Private WithEvents m_mnuManageXOS As Menu
Attribute m_mnuManageXOS.VB_VarHelpID = -1
Private WithEvents m_mnuSelectXOS As Menu
Attribute m_mnuSelectXOS.VB_VarHelpID = -1
Private WithEvents m_mnuRemoveXOS As Menu
Attribute m_mnuRemoveXOS.VB_VarHelpID = -1
Private WithEvents m_mnuPrint As Menu
Attribute m_mnuPrint.VB_VarHelpID = -1
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
    eGDCol_OrderID = 0
    eGDCol_AccountID
    eGDCol_AccountType
    eGDCol_Symbol
    eGDCol_SymbolID
    eGDCol_Remove
    eGDCol_SortKey
    eGDCol_AscSortKey
    eGDCol_DescSortKey
    eGDCol_OutlineLevel
    
    eGDCol_OrderText
    eGDCol_Cancel
    eGDCol_Modify
    eGDCol_NumFilled
    eGDCol_Status
    eGDCol_CurrentPrice
    eGDCol_CurrentBid
    eGDCol_CurrentAsk
    eGDCol_Date
    eGDCol_Account
    eGDCol_AutoTradeItem
    eGDCol_BrokerID
    eGDCol_LinkStatus
    eGDCol_TIF
    eGDCol_NumCols
End Enum

Private Type mPrivate
    strSource As String                 ' Name for the calling form
    TCGrid As cTradeConsoleGrid         ' Trade console grid object for generic routines
    tmrRealTime As Timer                ' Realtime timer control
    frm As Form                         ' Form that the grid is on
    SnapshotOrders As cGdTree           ' Collection of snapshot order collections by broker
    mnuOrders As Object                 ' Orders menu
    mnuOrderGroup As Object             ' Order group menu
    
    bSummary As Boolean                 ' Is this a summary grid?
    strSummaryDisplay As String         ' Summary display string
    strDisplay As String                ' Non-summary display string
    lSortedCol As Long                  ' Sorted column
    nSortedDir As SortSettings          ' Sort direction for the column
    
    dGroupsLastModified As Double       ' Last time the active TradeSense order groups were changed
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

Public Property Get StatusCol() As Long
    StatusCol = OrdersCol(eGDCol_Status)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object
'' Inputs:      Source, Working Orders UI, Summary?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal strSource As String, UI As cWorkingOrdersControls, ByVal bSummary As Boolean)
On Error GoTo ErrSection:

    m.strSource = strSource
    With UI
        Set m.frm = .frm
        Set m_fgWorkingOrders = .fgGrid
        Set m.tmrRealTime = .tmrRealTime
        Set m_tmrMenu = .tmrMenu
        Set m.mnuOrders = .mnuOrders
        Set m_mnuBuy = .mnuBuy
        Set m_mnuSell = .mnuSell
        Set m_mnuOrderGroups = .mnuOrderGroups
        Set m.mnuOrderGroup = .mnuOrderGroup
        Set m_mnuEditOrder = .mnuEditOrder
        Set m_mnuCancelOrder = .mnuCancelOrder
        Set m_mnuParkOrder = .mnuParkOrder
        Set m_mnuSubmitOrder = .mnuSubmitOrder
        Set m_mnuSubmitAll = .mnuSubmitAll
        Set m_mnuOrderHistory = .mnuOrderHistory
        Set m_mnuNewJournal = .mnuNewJournal
        Set m_mnuManageXOS = .mnuManageXOS
        Set m_mnuSelectXOS = .mnuSelectXOS
        Set m_mnuRemoveXOS = .mnuRemoveXOS
        Set m_mnuPrint = .mnuPrint
        Set m_mnuTradeHistory = .mnuTradeHistory
        Set m_mnuSettings = .mnuSettings
        Set m_mnuCheckStatus = .mnuCheckStatus
        Set m_mnuViewJournals = .mnuViewJournals
        Set m_mnuAutoSizeColumns = .mnuAutoSizeColumns
        Set m_mnuDefaultColumns = .mnuDefaultColumns
    End With
    
    InitGrid
    
    m.bSummary = bSummary
    m.strSummaryDisplay = FixDisplay(GetIniFileProperty("SummaryOrdersDisplay", "", "TTSummary", g.strIniFile), True)
    m.strDisplay = FixDisplay(GetIniFileProperty("OrdersDisplay", "", "TTSummary", g.strIniFile), False)
    
    Set m.TCGrid = New cTradeConsoleGrid
    If m.bSummary Then
        m.TCGrid.Init m_fgWorkingOrders, m.tmrRealTime, m.strSummaryDisplay
    Else
        m.TCGrid.Init m_fgWorkingOrders, m.tmrRealTime, m.strDisplay
    End If
    
    ApplyDisplayString
    
    If OrdersCol(eGDCol_SortKey) <> GDCol(eGDCol_SortKey) Then
        m_fgWorkingOrders.ColPosition(OrdersCol(eGDCol_SortKey)) = GDCol(eGDCol_SortKey)
    End If
    If OrdersCol(eGDCol_AscSortKey) <> GDCol(eGDCol_AscSortKey) Then
        m_fgWorkingOrders.ColPosition(OrdersCol(eGDCol_AscSortKey)) = GDCol(eGDCol_AscSortKey)
    End If
    If OrdersCol(eGDCol_DescSortKey) <> GDCol(eGDCol_DescSortKey) Then
        m_fgWorkingOrders.ColPosition(OrdersCol(eGDCol_DescSortKey)) = GDCol(eGDCol_DescSortKey)
    End If
    If OrdersCol(eGDCol_OutlineLevel) <> GDCol(eGDCol_OutlineLevel) Then
        m_fgWorkingOrders.ColPosition(OrdersCol(eGDCol_OutlineLevel)) = GDCol(eGDCol_OutlineLevel)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Clear the updated colors in the cells of the grid
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

    With m_fgWorkingOrders
        iSaveRedraw = .Redraw
        
        .Redraw = flexRDNone
        For lRow = .FixedRows To .Rows - 1
            If g.bUnloading = False Then
                If .RowHidden(lRow) = False Then
                    If .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UpdateColor Then
                        bStillColor = False
                        If m.tmrRealTime.Enabled Then
                            For lCol = OrdersCol(eGDCol_CurrentPrice) To OrdersCol(eGDCol_CurrentAsk)
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
                            Next lCol
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
    RaiseError "cWorkingOrdersUI.ClearUpdatedColors"
    
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
    Dim Bars As New cGdBars             ' Bars object
    
    SetBarProperties Bars, vSymbolOrSymbolID
    
    With m_fgWorkingOrders
        For lIndex = .FixedRows To .Rows - 1
            ' DAJ 08/06/2007: Only need to update visible rows...
            If .RowHidden(lIndex) = False Then
                If SymbolForRow(lIndex) = vSymbolOrSymbolID Then
                    strValue = ""
                    If dPrice <> kNullData Then
                        strValue = Bars.PriceDisplay(dPrice, True)
                    End If
                    m.TCGrid.ChangeCell lIndex, OrdersCol(eGDCol_CurrentPrice), strValue
                
                    strValue = ""
                    If dBid <> kNullData Then
                        strValue = Bars.PriceDisplay(dBid, True)
                    End If
                    m.TCGrid.ChangeCell lIndex, OrdersCol(eGDCol_CurrentBid), strValue
                
                    strValue = ""
                    If dAsk <> kNullData Then
                        strValue = Bars.PriceDisplay(dAsk, True)
                    End If
                    m.TCGrid.ChangeCell lIndex, OrdersCol(eGDCol_CurrentAsk), strValue
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.RefreshPrices"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPrices2
'' Description: Refresh the prices in the grid as appropriate
'' Inputs:      Symbol, Price, Bid, Ask
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPrices2(ByVal vSymbolOrSymbolID As Variant, ByVal strPrice As String, ByVal strBid As String, ByVal strAsk As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With m_fgWorkingOrders
        For lIndex = .TopRow To .BottomRow
            If m.TCGrid.ValidRow(lIndex) Then
                ' DAJ 08/06/2007: Only need to update visible rows...
                If .RowHidden(lIndex) = False Then
                    If SymbolForRow(lIndex) = vSymbolOrSymbolID Then
                        m.TCGrid.ChangeCell lIndex, OrdersCol(eGDCol_CurrentPrice), strPrice
                        m.TCGrid.ChangeCell lIndex, OrdersCol(eGDCol_CurrentBid), strBid
                        m.TCGrid.ChangeCell lIndex, OrdersCol(eGDCol_CurrentAsk), strAsk
                    End If
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.RefreshPrices2"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetGridRowForOrderID
'' Description: Set the current grid row for the given order ID
'' Inputs:      Order ID
'' Returns:     Grid Row
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetGridRowForOrderID(ByVal lOrderID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order object
    
    lReturn = -1&
    With m_fgWorkingOrders
        For lIndex = .FixedRows To .Rows - 1
            Set Order = OrderForRow(lIndex)
            If Not Order Is Nothing Then
                If Order.OrderID = lOrderID Then
                    lReturn = lIndex
                    Exit For
                End If
            End If
        Next lIndex
    
        If m.TCGrid.ValidRow(lReturn) Then
            .RowSel = lReturn
            .Row = lReturn
        End If
    End With
    
    SetGridRowForOrderID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.SetGridRowForOrderID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasParkedOrders
'' Description: Are there parked orders in the open orders grid?
'' Inputs:      None
'' Returns:     True if Parked Orders exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasParkedOrders() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order object
    Dim bReturn As Boolean              ' Return value from the function
    
    With m_fgWorkingOrders
        bReturn = False
        For lIndex = .FixedRows To .Rows - 1
            Set Order = OrderForRow(lIndex)
            If Not Order Is Nothing Then
                If Order.Status = eTT_OrderStatus_Parked Then
                    bReturn = True
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    HasParkedOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.HasParkedOrders"
    
End Function

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

    With m_fgWorkingOrders
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, OrdersCol(eGDCol_AutoTradeItem)) = strOldName Then
                .TextMatrix(lIndex, OrdersCol(eGDCol_AutoTradeItem)) = strNewName
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.ChangeTradeItemName"
    
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

    ' Get the latest collections for the broker...
    Set m.SnapshotOrders(nBroker) = g.Broker.SnapshotOrdersForBroker(nBroker)
    
    With m_fgWorkingOrders
        .Redraw = flexRDNone
        
        ' Update the orders grid...
        ResetRemoveOrderFlag nBroker
        RefreshBrokerOrders nBroker, False
        RemoveFlaggedOrders nBroker
        
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.Update"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateTsOrders
'' Description: Update the TradeSense order groups in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateTsOrders()
On Error GoTo ErrSection:

    If g.TsoGroups.LastModified > m.dGroupsLastModified Then
        With m_fgWorkingOrders
            .Redraw = flexRDNone
            
            ResetRemoveTsOrderFlag
            RefreshTsOrders False
            RemoveFlaggedTsOrders
            
            .Redraw = flexRDBuffered
        End With
        
        m.dGroupsLastModified = g.TsoGroups.LastModified
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.UpdateTsOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterOrdersGrid
'' Description: Hide orders that we do not want showing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FilterOrdersGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID for the row in the grid
    Dim lNumVisible As Long             ' Number of visible orders
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With m_fgWorkingOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lNumVisible = 0&
        For lIndex = .FixedRows To .Rows - 1
            lAccountID = CLng(Val(.TextMatrix(lIndex, OrdersCol(eGDCol_AccountID))))
            .RowHidden(lIndex) = g.Broker.HideAccount(lAccountID)
            If (.RowHidden(lIndex) = False) And (.RowOutlineLevel(lIndex) = 1) Then
                lNumVisible = lNumVisible + 1&
            End If
        Next lIndex
        
        SortOnCol
        SetBackColors m_fgWorkingOrders
        ShowAutoTradeColumns
        m.TCGrid.AutoSizeGrid
        
        .Redraw = nRedraw
    End With
    
    If Not g.ConsoleForms Is Nothing Then
        g.ConsoleForms.NumVisible(eGDConsoleForm_OpenOrders) = lNumVisible
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.FilterOrdersGrid"
    
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

    m.TCGrid.ApplyDisplayString eGDCol_Modify + 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.ApplyDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedOrder
'' Description: Order attached to the selected row in the grid
'' Inputs:      None
'' Returns:     Order if valid row, Nothing otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectedOrder() As cPtOrder
On Error GoTo ErrSection:

    Set SelectedOrder = OrderForRow(m_fgWorkingOrders.Row)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.SelectedOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForRow
'' Description: Order attached to the given row in the grid
'' Inputs:      Row
'' Returns:     Order if valid row, Nothing otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderForRow(ByVal Row As Long) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to return from the function

    Set Order = Nothing
    With m_fgWorkingOrders
        If m.TCGrid.ValidRow(Row) Then
            If TypeOf .RowData(Row) Is cPtOrder Then
                Set Order = .RowData(Row)
            End If
        End If
    End With
    
    Set OrderForRow = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.OrderForRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedTsOrder
'' Description: TradeSense Order attached to the selected row in the grid
'' Inputs:      None
'' Returns:     TradeSense Order if valid row, Nothing otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectedTsOrder() As cActiveTsOrderGroup
On Error GoTo ErrSection:

    Set SelectedTsOrder = TsOrderForRow(m_fgWorkingOrders.Row)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.SelectedTsOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TsOrderForRow
'' Description: TradeSense Order attached to the given row in the grid
'' Inputs:      Row
'' Returns:     TradeSense Order if valid row, Nothing otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TsOrderForRow(ByVal Row As Long) As cActiveTsOrderGroup
On Error GoTo ErrSection:

    Dim tsOrder As cActiveTsOrderGroup  ' Order to return from the function

    Set tsOrder = Nothing
    With m_fgWorkingOrders
        If m.TCGrid.ValidRow(Row) Then
            If TypeOf .RowData(Row) Is cActiveTsOrderGroup Then
                Set tsOrder = .RowData(Row)
            End If
        End If
    End With
    
    Set TsOrderForRow = tsOrder

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.TsOrderForRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolForRow
'' Description: Symbol for the order attached to the given row in the grid
'' Inputs:      Row
'' Returns:     Symbol or Symbol ID if valid row, Blank otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SymbolForRow(ByVal Row As Long) As Variant
On Error GoTo ErrSection:

    Dim vReturn As Variant              ' Return value for the function
    Dim strSymbol As String             ' Symbol from the grid
    Dim lSymbolID As Long               ' Symbol ID from the grid
    
    vReturn = ""
    If m.TCGrid.ValidRow(Row) Then
        If TypeOf m_fgWorkingOrders.RowData(Row) Is cPtOrder Then
            If m_fgWorkingOrders.RowData(Row).NumberOfLegs = 1 Then
                vReturn = m_fgWorkingOrders.RowData(Row).SymbolOrSymbolID
            Else
                vReturn = m_fgWorkingOrders.RowData(Row).SpreadSymbol
            End If
        ElseIf TypeOf m_fgWorkingOrders.RowData(Row) Is cOrderLeg Then
            vReturn = m_fgWorkingOrders.RowData(Row).SymbolOrSymbolID
        ElseIf TypeOf m_fgWorkingOrders.RowData(Row) Is cActiveTsOrderGroup Then
            vReturn = m_fgWorkingOrders.RowData(Row).SymbolOrSymbolID
        End If
        
        'lSymbolID = CLng(Val(m_fgWorkingOrders.TextMatrix(Row, OrdersCol(eGDCol_SymbolID))))
        'strSymbol = m_fgWorkingOrders.TextMatrix(Row, OrdersCol(eGDCol_Symbol))
        
        'If lSymbolID = 0 Then
        '    vReturn = strSymbol
        'Else
        '    vReturn = lSymbolID
        'End If
    End If
    
    SymbolForRow = vReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.SymbolForRow"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectOrderGroup
'' Description: The user has chosed the selected order group
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SelectOrderGroup(ByVal Index As Long)
On Error GoTo ErrSection:

    Dim lOrderGroupID As Long           ' Order Group ID to load
    Dim rs As Recordset                 ' Recordset into the database
    Dim Order As New cPtOrder           ' Temporary order object
    Dim Orders As New cGdTree           ' Collection of orders
    Dim strMessage As String            ' Message to display to the user
    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return from an infbox

    If Index = 0 Then
        StartMenuTimer "ORDERGROUPS"
    Else
        lOrderGroupID = CLng(Val(m.mnuOrderGroup(Index).Tag))
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroupItems] WHERE [OrderGroupID]=" & Str(lOrderGroupID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            Set Order = frmOrderGroup.OrderFromOrderText(rs!OrderText)
            Orders.Add Order
            
            rs.MoveNext
        Loop
        
        If Orders.Count > 0 Then
            strMessage = "You are about to place all of the|orders in the order group:|" & m.mnuOrderGroup(Index).Caption & "|"
            strReturn = InfBox(strMessage, "", "+Submit|Park|-Cancel", m.mnuOrderGroup(Index).Caption)
            If strReturn <> "C" Then
                For lIndex = 1 To Orders.Count
                    If Len(Orders(lIndex).GenesisOrderID) = 0 Then
                        Orders(lIndex).GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Orders(lIndex).AccountID))
                    End If
                    Orders(lIndex).Status = eTT_OrderStatus_Parked
                    Orders(lIndex).Save
                    
                    g.Broker.AddOrder Orders(lIndex)
                    OrderCallback Orders(lIndex)
                Next lIndex
                
                If strReturn = "S" Then SubmitMultipleOrders Orders
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.SelectOrderGroup"
    
End Sub

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
    FontFromString m_fgWorkingOrders.Font, g.Broker.GridFont
    m_fgWorkingOrders.Font = m_fgWorkingOrders.Font
    
    ' Update the columns displayed on the grids on the summary tab...
    If m.bSummary Then
        m.strSummaryDisplay = FixDisplay(GetIniFileProperty("SummaryOrdersDisplay", "", "TTSummary", g.strIniFile), True)
        m.TCGrid.Display = m.strSummaryDisplay
        ApplyDisplayString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.UpdateConsoleSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Setup the popup menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EnableControls()
On Error GoTo ErrSection:

    Dim bValidRow As Boolean            ' Valid row in the grid?
    Dim Order As cPtOrder               ' Currently selected order
    Dim tsOrder As cActiveTsOrderGroup  ' Currently selected TradeSense order
    Dim strAutoExit As String           ' Auto exit for the order

    bValidRow = m.TCGrid.CurrentRowValid
    
    m_mnuEditOrder.Enabled = bValidRow
    m_mnuCancelOrder.Enabled = bValidRow
    m_mnuParkOrder.Enabled = bValidRow
    m_mnuSubmitAll.Enabled = HasParkedOrders
    m_mnuOrderHistory.Enabled = bValidRow
    m_mnuNewJournal.Enabled = bValidRow
    m_mnuManageXOS.Enabled = True
    
    Set Order = SelectedOrder
    Set tsOrder = SelectedTsOrder
    
    If Not tsOrder Is Nothing Then
        m_mnuOrderGroups.Visible = False
        m_mnuEditOrder.Visible = False
        m_mnuCancelOrder.Caption = "Cancel TradeSense Order Group"
        m_mnuParkOrder.Caption = "Park TradeSense Order Group"
        m_mnuParkOrder.Enabled = tsOrder.Submitted
        m_mnuSubmitOrder.Caption = "Submit TradeSense Order Group"
        m_mnuSubmitOrder.Enabled = Not tsOrder.Submitted
        m_mnuSubmitAll.Visible = False
        m_mnuOrderHistory.Visible = False
        m_mnuNewJournal.Visible = False
        m_mnuManageXOS.Caption = "Manage TradeSense Order Groups"
        m_mnuSelectXOS.Visible = False
        m_mnuRemoveXOS.Visible = False
    ElseIf Not Order Is Nothing Then
        m_mnuOrderGroups.Visible = True
        m_mnuEditOrder.Visible = True
        m_mnuEditOrder.Enabled = (Order.NumberOfLegs = 1)
        m_mnuCancelOrder.Caption = "Cancel Order"
        m_mnuParkOrder.Caption = "Park Order"
        m_mnuSubmitOrder.Caption = "Submit Order"
        m_mnuSubmitAll.Visible = True
        m_mnuOrderHistory.Visible = True
        m_mnuNewJournal.Visible = True
        m_mnuManageXOS.Caption = "Manage Exit Order Strategies"
        m_mnuSelectXOS.Visible = True
        m_mnuRemoveXOS.Visible = True
        
        strAutoExit = g.OrderStrategies.ExitForAccountAndSymbol(Order.AccountID, Order.SymbolOrSymbolID)
        
        m_mnuSubmitOrder.Enabled = bValidRow And ((Order.Status = eTT_OrderStatus_Parked) Or (Order.Status = eTT_OrderStatus_BrokerParked))
        m_mnuSelectXOS.Enabled = (Order.AutoTradeItemID = 0&)
        m_mnuRemoveXOS.Enabled = (Order.AutoTradeItemID = 0&) And (Len(strAutoExit) > 0)
    Else
        m_mnuEditOrder.Enabled = False
        m_mnuCancelOrder.Enabled = False
        m_mnuParkOrder.Enabled = False
        m_mnuOrderHistory.Enabled = False
        m_mnuNewJournal.Enabled = False
        m_mnuSubmitOrder.Enabled = False
        m_mnuSelectXOS.Enabled = False
        m_mnuRemoveXOS.Enabled = False
    End If
    
    m_mnuCheckStatus.Enabled = g.Broker.EnableCheckStatusMenu
    m_mnuDefaultColumns.Enabled = (m.TCGrid.Display <> DefaultDisplay)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.EnableControls"

End Sub

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

    m.TCGrid.GenerateReport vArgs, "Working Orders"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.GenerateReport"
    
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

    With m_fgWorkingOrders
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbButtonFace
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = False
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .HighLight = flexHighlightNever
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        .FrozenCols = GDCol(eGDCol_Modify) + 1
        .OutlineCol = GDCol(eGDCol_Symbol)
                
        .TextMatrix(0, GDCol(eGDCol_OrderID)) = "Order ID"
        .TextMatrix(0, GDCol(eGDCol_AccountID)) = "Account ID"
        .TextMatrix(0, GDCol(eGDCol_AccountType)) = "Account Type"
        .TextMatrix(0, GDCol(eGDCol_SymbolID)) = "Symbol ID"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Remove)) = "Remove"
        .TextMatrix(0, GDCol(eGDCol_SortKey)) = "Sort Key"
        .TextMatrix(0, GDCol(eGDCol_AscSortKey)) = "Asc Sort Key"
        .TextMatrix(0, GDCol(eGDCol_DescSortKey)) = "Desc Sort Key"
        .TextMatrix(0, GDCol(eGDCol_OutlineLevel)) = "Outline Level"
        
        .TextMatrix(0, GDCol(eGDCol_Date)) = "Date"
        .TextMatrix(0, GDCol(eGDCol_OrderText)) = "Working Order"
        .Cell(flexcpFontBold, 0, GDCol(eGDCol_OrderText)) = True
        .TextMatrix(0, GDCol(eGDCol_Cancel)) = "X"
        .TextMatrix(0, GDCol(eGDCol_Modify)) = "M"
        .TextMatrix(0, GDCol(eGDCol_NumFilled)) = "Filled"
        .TextMatrix(0, GDCol(eGDCol_CurrentPrice)) = "Current Price"
        .TextMatrix(0, GDCol(eGDCol_CurrentBid)) = "Bid"
        .TextMatrix(0, GDCol(eGDCol_CurrentAsk)) = "Ask"
        .TextMatrix(0, GDCol(eGDCol_Status)) = "Status"
        .TextMatrix(0, GDCol(eGDCol_Account)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_AutoTradeItem)) = "Auto Trade Item"
        .TextMatrix(0, GDCol(eGDCol_BrokerID)) = "Broker ID"
        .TextMatrix(0, GDCol(eGDCol_LinkStatus)) = "Link"
        .TextMatrix(0, GDCol(eGDCol_TIF)) = "TIF"
        
        .ColHidden(GDCol(eGDCol_OrderID)) = True
        .ColHidden(GDCol(eGDCol_AccountID)) = True
        .ColHidden(GDCol(eGDCol_AccountType)) = True
        .ColHidden(GDCol(eGDCol_SymbolID)) = True
        .ColHidden(GDCol(eGDCol_Remove)) = True
        .ColHidden(GDCol(eGDCol_SortKey)) = True
        .ColHidden(GDCol(eGDCol_AscSortKey)) = True
        .ColHidden(GDCol(eGDCol_DescSortKey)) = True
        .ColHidden(GDCol(eGDCol_OutlineLevel)) = True
                                
        .ColFormat(GDCol(eGDCol_Date)) = DateAndTime("Format")
        
        .ColDataType(GDCol(eGDCol_Remove)) = flexDTBoolean
        .ColDataType(GDCol(eGDCol_Date)) = flexDTDate
        
        .ColAlignment(GDCol(eGDCol_Date)) = flexAlignCenterTop
        .ColAlignment(GDCol(eGDCol_CurrentPrice)) = flexAlignRightTop
        .ColAlignment(GDCol(eGDCol_CurrentBid)) = flexAlignRightTop
        .ColAlignment(GDCol(eGDCol_CurrentAsk)) = flexAlignRightTop
        .ColAlignment(GDCol(eGDCol_Account)) = flexAlignLeftTop
        .ColAlignment(GDCol(eGDCol_BrokerID)) = flexAlignLeftTop
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.InitOrdersGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrdersCol
'' Description: Determine where the given column is really sitting right now
'' Inputs:      Column Number
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrdersCol(ByVal Col As eGDCols) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    If Col = eGDCol_NumCols Then
        lReturn = Col
    Else
        lReturn = m.TCGrid.ColNumFromColData(Col)
    End If
    
    OrdersCol = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.OrdersCol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetSymbolForRow
'' Description: Set the symbol for the given row
'' Inputs:      Row, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetSymbolForRow(ByVal lRow As Long, ByVal vSymbolOrSymbolID As Variant, Optional ByVal strSpreadSymbol As String = "")
On Error GoTo ErrSection:

    Dim Bars As cGdBars                 ' Bars structure to get latest value
    Dim SpreadData As cSpreadData       ' Spread data structure to get latest value
    Dim nRedraw As Long                 ' Current state of the grid's redraw

    With m_fgWorkingOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .TextMatrix(lRow, OrdersCol(eGDCol_SymbolID)) = Str(GetSymbolID(vSymbolOrSymbolID))
        .TextMatrix(lRow, OrdersCol(eGDCol_Symbol)) = GetSymbol(vSymbolOrSymbolID)
        
        If Len(strSpreadSymbol) = 0 Then
            Set Bars = GetTradeBars(vSymbolOrSymbolID)
            If Not Bars Is Nothing Then
                If Bars(eBARS_Close, Bars.Size - 1) <> kNullData Then
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentPrice)) = Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
                Else
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentPrice)) = ""
                End If
                If Bars(eBARS_Bid, Bars.Size - 1) <> kNullData Then
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentBid)) = Bars.PriceDisplay(Bars(eBARS_Bid, Bars.Size - 1))
                Else
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentBid)) = ""
                End If
                If Bars(eBARS_Ask, Bars.Size - 1) <> kNullData Then
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentAsk)) = Bars.PriceDisplay(Bars(eBARS_Ask, Bars.Size - 1))
                Else
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentAsk)) = ""
                End If
            Else
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentPrice)) = ""
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentBid)) = ""
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentAsk)) = ""
            End If
        Else
            Set SpreadData = frmTTSummary.GetSpreadData(strSpreadSymbol)
            If Not SpreadData Is Nothing Then
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentPrice)) = ""
                If SpreadData.Bid <> kNullData Then
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentBid)) = Str(SpreadData.Bid)
                Else
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentBid)) = ""
                End If
                If SpreadData.Ask <> kNullData Then
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentAsk)) = Str(SpreadData.Ask)
                Else
                    .TextMatrix(lRow, OrdersCol(eGDCol_CurrentAsk)) = ""
                End If
            Else
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentPrice)) = ""
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentBid)) = ""
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentAsk)) = ""
            End If
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.SetSymbolForRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToGrid
'' Description: Update the given order in the grid
'' Inputs:      Order, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderToGrid(Order As cPtOrder, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Bars As cGdBars                 ' Bars structure to get latest value
    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lNumLegs As Long                ' Number of legs for the order

    With m_fgWorkingOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lNumLegs = Order.NumberOfLegs
        
        If lRow = -1& Then
            If lNumLegs = 1 Then
                .Rows = .Rows + 1
                lRow = .Rows - 1
            Else
                .Rows = .Rows + (lNumLegs + 1)
                lRow = .Rows - (lNumLegs + 1)
            End If
        End If
        
        .RowData(lRow) = Order.MakeCopy
        
        ' DAJ 10/10/2013: Moved this to mTradeTracker.RefreshOrder.  Since this is coming off a copy
        ' of the original order, the Bars Handle is different than the one that got added to the
        ' g.RealTime.TickBuffers, so the data wasn't updating...
        'If ((Order.IsConditional(True) = True) And ((Order.Status = eTT_OrderStatus_TriggerPending) Or (Order.Status = eTT_OrderStatus_DataPending))) Or _
        '   (((Order.TrailAmount <> 0) Or (Order.ExpireTime <> 0)) And (IsOpenOrder(Order.Status) = True)) Then
        '    g.CondOrders(Str(Order.OrderID)) = Order
        'ElseIf g.CondOrders.Exists(Str(Order.OrderID)) And (Order.Status <> eTT_OrderStatus_TriggerPending) And (Order.TrailAmount = 0) Then
        '    g.CondOrders.Remove Str(Order.OrderID)
        'End If
        
        .TextMatrix(lRow, OrdersCol(eGDCol_OrderID)) = Str(Order.OrderID)
        .TextMatrix(lRow, OrdersCol(eGDCol_AccountID)) = Str(Order.AccountID)
        .TextMatrix(lRow, OrdersCol(eGDCol_AccountType)) = Str(Order.Broker)
        .TextMatrix(lRow, OrdersCol(eGDCol_Date)) = ConvertBrokerDate(Order.OrderDate, g.Broker.AccountTypeForID(Order.AccountID), Order.Symbol, g.bShowInLocalTimeZone)
        .TextMatrix(lRow, OrdersCol(eGDCol_SortKey)) = ""
        .TextMatrix(lRow, OrdersCol(eGDCol_AscSortKey)) = "O" & Str(Order.OrderID) & Chr(31)
        .TextMatrix(lRow, OrdersCol(eGDCol_DescSortKey)) = "O" & Str(Order.OrderID) & "}"
        .TextMatrix(lRow, OrdersCol(eGDCol_OutlineLevel)) = "1"
        
        If lNumLegs = 1 Then
            SetSymbolForRow lRow, Order.SymbolOrSymbolID
        Else
            SetSymbolForRow lRow, Order.UnderlyingSymbolOrSymbolID, Order.SpreadSymbol
        End If
        
        .TextMatrix(lRow, OrdersCol(eGDCol_OrderText)) = Order.OrderText(False)
        .TextMatrix(lRow, OrdersCol(eGDCol_Cancel)) = "X"
        .Cell(flexcpFontUnderline, lRow, GDCol(eGDCol_Cancel)) = True
        If g.nColorTheme = kDarkThemeColor Then
            .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Cancel)) = vbCyan
        Else
            .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Cancel)) = vbBlue
        End If
        .TextMatrix(lRow, OrdersCol(eGDCol_Modify)) = "M"
        .Cell(flexcpFontUnderline, lRow, GDCol(eGDCol_Modify)) = True
        If Order.NumberOfLegs = 1 Then
            If g.nColorTheme = kDarkThemeColor Then
                .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Modify)) = vbCyan
            Else
                .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Modify)) = vbBlue
            End If
        Else
            .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Modify)) = RGB(128, 128, 128)
        End If
        .TextMatrix(lRow, OrdersCol(eGDCol_NumFilled)) = Str(Order.FillQuantity)
                    
        If (Order.Status = eTT_OrderStatus_Parked) Or (lNumLegs > 1) Then
            .Cell(flexcpForeColor, lRow, OrdersCol(eGDCol_Symbol), lRow, OrdersCol(eGDCol_OrderText)) = .ForeColor
        ElseIf Order.Buy Then
            .Cell(flexcpForeColor, lRow, OrdersCol(eGDCol_Symbol), lRow, OrdersCol(eGDCol_OrderText)) = g.ChartGlobals.nLongColor
        Else
            .Cell(flexcpForeColor, lRow, OrdersCol(eGDCol_Symbol), lRow, OrdersCol(eGDCol_OrderText)) = g.ChartGlobals.nShortColor
        End If
        
        .TextMatrix(lRow, OrdersCol(eGDCol_Status)) = OrderStatus(Order.Status)
        .TextMatrix(lRow, OrdersCol(eGDCol_Account)) = g.Broker.AccountNameForID(Order.AccountID)
        .TextMatrix(lRow, OrdersCol(eGDCol_BrokerID)) = Str(Order.BrokerID)
        .TextMatrix(lRow, OrdersCol(eGDCol_LinkStatus)) = Order.LinkStatus
        .TextMatrix(lRow, OrdersCol(eGDCol_TIF)) = TimeInForce(Order.TimeInForce)
        
        .TextMatrix(lRow, OrdersCol(eGDCol_AutoTradeItem)) = g.TradingItems.NameForID(Order.AutoTradeItemID)
        
        .IsSubtotal(lRow) = True
        .RowOutlineLevel(lRow) = 1
        CheckedCell(m_fgWorkingOrders, lRow, OrdersCol(eGDCol_Remove)) = False
        
        If lNumLegs > 1 Then
            For lIndex = 1 To lNumLegs
                .RowData(lRow + lIndex) = Order.OrderLegs(lIndex)
                
                .TextMatrix(lRow + lIndex, OrdersCol(eGDCol_AccountID)) = Str(Order.AccountID)
                
                .TextMatrix(lRow + lIndex, OrdersCol(eGDCol_SortKey)) = ""
                .TextMatrix(lRow + lIndex, OrdersCol(eGDCol_AscSortKey)) = "O" & Str(Order.OrderID) & "_" & Format(lIndex, "000")
                .TextMatrix(lRow + lIndex, OrdersCol(eGDCol_DescSortKey)) = "O" & Str(Order.OrderID) & "_" & Format(lNumLegs + 1 - lIndex, "000")
                .TextMatrix(lRow + lIndex, OrdersCol(eGDCol_OutlineLevel)) = "2"
                
                .IsSubtotal(lRow + lIndex) = True
                .RowOutlineLevel(lRow + lIndex) = 2
                .TextMatrix(lRow + lIndex, OrdersCol(eGDCol_OrderText)) = Order.OrderLegs(lIndex).ToString(False)
                SetSymbolForRow lRow + lIndex, Order.OrderLegs(lIndex).SymbolOrSymbolID
                
                CheckedCell(m_fgWorkingOrders, lRow + lIndex, OrdersCol(eGDCol_Remove)) = False
            
                If Order.OrderLegs(lIndex).IsBuy Then
                    .Cell(flexcpForeColor, lRow + lIndex, OrdersCol(eGDCol_Symbol), lRow + lIndex, OrdersCol(eGDCol_OrderText)) = g.ChartGlobals.nLongColor
                Else
                    .Cell(flexcpForeColor, lRow + lIndex, OrdersCol(eGDCol_Symbol), lRow + lIndex, OrdersCol(eGDCol_OrderText)) = g.ChartGlobals.nShortColor
                End If
            Next lIndex
        End If
               
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.OrderToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TsOrderToGrid
'' Description: Update the given TradeSense order in the grid
'' Inputs:      TradeSense Order, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TsOrderToGrid(tsOrder As cActiveTsOrderGroup, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Bars As cGdBars                 ' Bars structure to get latest value
    Dim nRedraw As Long                 ' Current state of the grid's redraw

    With m_fgWorkingOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = tsOrder
        
        .TextMatrix(lRow, OrdersCol(eGDCol_OrderID)) = Str(tsOrder.ID)
        .TextMatrix(lRow, OrdersCol(eGDCol_AccountID)) = Str(tsOrder.AccountID)
        .TextMatrix(lRow, OrdersCol(eGDCol_AccountType)) = Str(tsOrder.Broker)
        .TextMatrix(lRow, OrdersCol(eGDCol_Date)) = ""
        .TextMatrix(lRow, OrdersCol(eGDCol_SymbolID)) = Str(GetSymbolID(tsOrder.SymbolOrSymbolID))
        .TextMatrix(lRow, OrdersCol(eGDCol_Symbol)) = GetSymbol(tsOrder.SymbolOrSymbolID)
        .TextMatrix(lRow, OrdersCol(eGDCol_SortKey)) = ""
        .TextMatrix(lRow, OrdersCol(eGDCol_AscSortKey)) = "TS" & Str(tsOrder.ID) & Chr(31)
        .TextMatrix(lRow, OrdersCol(eGDCol_DescSortKey)) = "TS" & Str(tsOrder.ID) & "}"
        .TextMatrix(lRow, OrdersCol(eGDCol_OutlineLevel)) = "1"
        
        .TextMatrix(lRow, OrdersCol(eGDCol_OrderText)) = tsOrder.tsOrderGroup.Name
        .TextMatrix(lRow, OrdersCol(eGDCol_Cancel)) = "X"
        .TextMatrix(lRow, OrdersCol(eGDCol_Modify)) = ""
        .TextMatrix(lRow, OrdersCol(eGDCol_NumFilled)) = ""
                    
        .Cell(flexcpForeColor, lRow, OrdersCol(eGDCol_Symbol), lRow, OrdersCol(eGDCol_OrderText)) = .ForeColor
        
        Set Bars = GetTradeBars(tsOrder.SymbolOrSymbolID)
        If Not Bars Is Nothing Then
            If Bars(eBARS_Close, Bars.Size - 1) <> kNullData Then
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentPrice)) = Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
            Else
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentPrice)) = ""
            End If
            If Bars(eBARS_Bid, Bars.Size - 1) <> kNullData Then
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentBid)) = Bars.PriceDisplay(Bars(eBARS_Bid, Bars.Size - 1))
            Else
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentBid)) = ""
            End If
            If Bars(eBARS_Ask, Bars.Size - 1) <> kNullData Then
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentAsk)) = Bars.PriceDisplay(Bars(eBARS_Ask, Bars.Size - 1))
            Else
                .TextMatrix(lRow, OrdersCol(eGDCol_CurrentAsk)) = ""
            End If
        Else
            .TextMatrix(lRow, OrdersCol(eGDCol_CurrentPrice)) = ""
            .TextMatrix(lRow, OrdersCol(eGDCol_CurrentBid)) = ""
            .TextMatrix(lRow, OrdersCol(eGDCol_CurrentAsk)) = ""
        End If
        
        If tsOrder.Submitted Then
            .TextMatrix(lRow, OrdersCol(eGDCol_Status)) = "Working"
        Else
            .TextMatrix(lRow, OrdersCol(eGDCol_Status)) = "Parked"
        End If
        .TextMatrix(lRow, OrdersCol(eGDCol_Account)) = g.Broker.AccountNameForID(tsOrder.AccountID)
        .TextMatrix(lRow, OrdersCol(eGDCol_BrokerID)) = ""
        .TextMatrix(lRow, OrdersCol(eGDCol_LinkStatus)) = ""
        .TextMatrix(lRow, OrdersCol(eGDCol_TIF)) = ""
        .TextMatrix(lRow, OrdersCol(eGDCol_AutoTradeItem)) = ""
        
        .IsSubtotal(lRow) = True
        .RowOutlineLevel(lRow) = 1
        CheckedCell(m_fgWorkingOrders, lRow, OrdersCol(eGDCol_Remove)) = False
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.TsOrderToGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetRemoveOrderFlag
'' Description: Set the remove order flag on all orders for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetRemoveOrderFlag(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim nBrokerForRow As eTT_AccountType ' Broker for the given row
    Dim lAccountID As Long              ' Account ID for the row
    
    With m_fgWorkingOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            nBrokerForRow = -1&
            If TypeOf .RowData(lIndex) Is cPtOrder Then
                nBrokerForRow = .RowData(lIndex).Broker
            ElseIf TypeOf .RowData(lIndex) Is cOrderLeg Then
                lAccountID = CLng(Val(.TextMatrix(lIndex, OrdersCol(eGDCol_AccountID))))
                nBrokerForRow = g.Broker.AccountTypeForID(lAccountID)
            End If
        
            If nBrokerForRow = nBroker Then
                CheckedCell(m_fgWorkingOrders, lIndex, OrdersCol(eGDCol_Remove)) = True
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.ResetRemoveOrderFlag"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFlaggedOrders
'' Description: Remove rows with the remove flag set for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveFlaggedOrders(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order object
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim nBrokerForRow As eTT_AccountType ' Broker for the given row
    Dim lAccountID As Long              ' Account ID for the row
    
    With m_fgWorkingOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .Rows - 1 To .FixedRows Step -1
            nBrokerForRow = -1&
            If TypeOf .RowData(lIndex) Is cPtOrder Then
                nBrokerForRow = .RowData(lIndex).Broker
            ElseIf TypeOf .RowData(lIndex) Is cOrderLeg Then
                lAccountID = CLng(Val(.TextMatrix(lIndex, OrdersCol(eGDCol_AccountID))))
                nBrokerForRow = g.Broker.AccountTypeForID(lAccountID)
            End If
        
            If nBrokerForRow = nBroker Then
                If CheckedCell(m_fgWorkingOrders, lIndex, OrdersCol(eGDCol_Remove)) = True Then
                    If TypeOf .RowData(lIndex) Is cPtOrder Then
                        Set Order = .RowData(lIndex)
                        
                        If g.CondOrders.Exists(Str(Order.OrderID)) Then
                            g.CondOrders.Remove Str(Order.OrderID)
                        End If
                    End If
                    
                    .RemoveItem lIndex
                End If
            End If
        Next lIndex
        
        FilterOrdersGrid
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.RemoveFlaggedOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshBrokerOrders
'' Description: Refresh orders in the grid for the given broker from the collection
'' Inputs:      Broker, Filter Grid?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshBrokerOrders(ByVal nBroker As eTT_AccountType, Optional ByVal bFilter As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    Dim Orders As cPtOrders             ' Collection of orders for the broker
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim bOpenOrder As Boolean           ' Is the order an open order?
    
    If Not m.SnapshotOrders(nBroker) Is Nothing Then
        Set Orders = m.SnapshotOrders(nBroker).MakeCopy
        With m_fgWorkingOrders
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            For lIndex = 1 To Orders.Count
                bOpenOrder = IsOpenOrder(Orders(lIndex).Status)
                If bOpenOrder Then
                    lRow = -1&
                    For lIndex2 = .FixedRows To .Rows - 1
                        If TypeOf .RowData(lIndex2) Is cPtOrder Then
                            If .RowData(lIndex2).OrderID = Orders(lIndex).OrderID Then
                                lRow = lIndex2
                                Exit For
                            End If
                        End If
                    Next lIndex2
                    
                    OrderToGrid Orders(lIndex), lRow
                
                ElseIf g.CondOrders.Exists(Str(Orders(lIndex).OrderID)) Then
                    g.CondOrders.Remove Str(Orders(lIndex).OrderID)
                End If
            Next lIndex
            
            If bFilter Then
                FilterOrdersGrid
            End If
            
            .Redraw = nRedraw
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.RefreshBrokerOrders"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetRemoveTsOrderFlag
'' Description: Set the remove flag on all TradeSense order groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetRemoveTsOrderFlag()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With m_fgWorkingOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cActiveTsOrderGroup Then
                CheckedCell(m_fgWorkingOrders, lIndex, OrdersCol(eGDCol_Remove)) = True
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.ResetRemoveTsOrderFlag"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFlaggedTsOrders
'' Description: Remove TradeSense order group rows with the remove flag set
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveFlaggedTsOrders()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With m_fgWorkingOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If TypeOf .RowData(lIndex) Is cActiveTsOrderGroup Then
                If CheckedCell(m_fgWorkingOrders, lIndex, OrdersCol(eGDCol_Remove)) = True Then
                    .RemoveItem lIndex
                End If
            End If
        Next lIndex
        
        FilterOrdersGrid
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.RemoveFlaggedTsOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTsOrders
'' Description: Refresh the TradeSense order groups in the grid
'' Inputs:      Filter Grid?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshTsOrders(Optional ByVal bFilter As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    Dim tsOrder As cActiveTsOrderGroup  ' TradeSense order group from the grid
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With m_fgWorkingOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 1 To g.TsoGroups.Count
            lRow = -1&
            
            For lIndex2 = .FixedRows To .Rows - 1
                Set tsOrder = TsOrderForRow(lIndex2)
                If Not tsOrder Is Nothing Then
                    If tsOrder.Key = g.TsoGroups(lIndex).Key Then
                        lRow = lIndex2
                        Exit For
                    End If
                End If
            Next lIndex2
        
            TsOrderToGrid g.TsoGroups(lIndex), lRow
        Next lIndex
        
        If bFilter Then
            FilterOrdersGrid
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.RefreshTsOrders"
    
End Sub

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
    
    astrColumns.Add Str(OrdersCol(eGDCol_OrderID))
    astrColumns.Add Str(OrdersCol(eGDCol_AccountID))
    astrColumns.Add Str(OrdersCol(eGDCol_SymbolID))
    astrColumns.Add Str(OrdersCol(eGDCol_Symbol))
    astrColumns.Add Str(OrdersCol(eGDCol_OrderText))
    astrColumns.Add Str(OrdersCol(eGDCol_Cancel))
    astrColumns.Add Str(OrdersCol(eGDCol_Modify))
    astrColumns.Add Str(OrdersCol(eGDCol_AccountType))
    astrColumns.Add Str(OrdersCol(eGDCol_Remove))
    astrColumns.Add Str(OrdersCol(eGDCol_SortKey))
    astrColumns.Add Str(OrdersCol(eGDCol_AscSortKey))
    astrColumns.Add Str(OrdersCol(eGDCol_DescSortKey))
    astrColumns.Add Str(OrdersCol(eGDCol_OutlineLevel))
    
    ColumnExceptions = "," & astrColumns.JoinFields(",") & ","

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.ColumnExceptions"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrderGroupMenu
'' Description: Load the order group menu with the available order groups
'' Inputs:      Order Group menu
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadOrderGroupMenu()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCounter As Long                ' Counter variable
    
    lCounter = 1&
    
    For lIndex = m.mnuOrderGroup.UBound To 1 Step -1
        Unload m.mnuOrderGroup(lIndex)
    Next lIndex
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroups];", dbOpenDynaset)
    Do While Not rs.EOF
        Load m.mnuOrderGroup(lCounter)
        m.mnuOrderGroup(lCounter).Caption = rs!Name
        m.mnuOrderGroup(lCounter).Tag = rs!OrderGroupID
        m.mnuOrderGroup(lCounter).Visible = True
        
        lCounter = lCounter + 1
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.LoadOrderGroupMenu"
    
End Sub

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
            'strReturn = "1;Filled;1;5,1;Status;1;6,1;Current Price;1;7,1;Bid;1;8,1;Ask;1;9,1;Date;1;10,1;Account;1;11,1;Auto Trade Item;1;12,0;Broker ID;1;13"
            astrDisplay.Add "0;Order ID;1;0;0;0"
            astrDisplay.Add "0;Account ID;1;1;0;0"
            astrDisplay.Add "0;Account Type;1;2;0;0"
            astrDisplay.Add "1;Symbol;1;3;0;0"
            astrDisplay.Add "0;Symbol ID;1;4;0;0"
            astrDisplay.Add "0;Remove;1;5;0;0"
            astrDisplay.Add "1;Working Order;1;6;0;0"
            astrDisplay.Add "1;X;1;7;0;0"
            astrDisplay.Add "1;M;1;8;0;0"
            astrDisplay.Add "1;Filled;1;9;0;1"
            astrDisplay.Add "1;Status;1;10;0;1"
            astrDisplay.Add "1;Current Price;1;11;0;1"
            astrDisplay.Add "1;Bid;1;12;0;1"
            astrDisplay.Add "1;Ask;1;13;0;1"
            astrDisplay.Add "0;Date;1;14;0;1"
            astrDisplay.Add "1;Account;1;15;0;1"
            astrDisplay.Add "0;Auto Trade Item;1;16;0;1"
            astrDisplay.Add "1;Broker ID;1;17;0;1"
            astrDisplay.Add "1;Link;1;18;0;1"
        
        Case Is >= 1280
            'strReturn = "1;Filled;1;5,1;Status;1;6,1;Current Price;1;7,0;Bid;1;8,0;Ask;1;9,0;Date;1;10,1;Account;1;11,1;Auto Trade Item;1;12,0;Broker ID;1;13"
            astrDisplay.Add "0;Order ID;1;0;0;0"
            astrDisplay.Add "0;Account ID;1;1;0;0"
            astrDisplay.Add "0;Account Type;1;2;0;0"
            astrDisplay.Add "1;Symbol;1;3;0;0"
            astrDisplay.Add "0;Symbol ID;1;4;0;0"
            astrDisplay.Add "0;Remove;1;5;0;0"
            astrDisplay.Add "1;Working Order;1;6;0;0"
            astrDisplay.Add "1;X;1;7;0;0"
            astrDisplay.Add "1;M;1;8;0;0"
            astrDisplay.Add "1;Filled;1;9;0;1"
            astrDisplay.Add "1;Status;1;10;0;1"
            astrDisplay.Add "1;Current Price;1;11;0;1"
            astrDisplay.Add "0;Bid;1;12;0;1"
            astrDisplay.Add "0;Ask;1;13;0;1"
            astrDisplay.Add "0;Date;1;14;0;1"
            astrDisplay.Add "1;Account;1;15;0;1"
            astrDisplay.Add "0;Auto Trade Item;1;16;0;1"
            astrDisplay.Add "1;Broker ID;1;17;0;1"
            astrDisplay.Add "1;Link;1;18;0;1"
        
        Case Is >= 1152
            'strReturn = "1;Filled;1;5,1;Status;1;6,1;Current Price;1;7,0;Bid;1;8,0;Ask;1;9,0;Date;1;10,0;Account;1;11,1;Auto Trade Item;1;12,0;Broker ID;1;13"
            astrDisplay.Add "0;Order ID;1;0;0;0"
            astrDisplay.Add "0;Account ID;1;1;0;0"
            astrDisplay.Add "0;Account Type;1;2;0;0"
            astrDisplay.Add "1;Symbol;1;3;0;0"
            astrDisplay.Add "0;Symbol ID;1;4;0;0"
            astrDisplay.Add "0;Remove;1;5;0;0"
            astrDisplay.Add "1;Working Order;1;6;0;0"
            astrDisplay.Add "1;X;1;7;0;0"
            astrDisplay.Add "1;M;1;8;0;0"
            astrDisplay.Add "1;Filled;1;9;0;1"
            astrDisplay.Add "1;Status;1;10;0;1"
            astrDisplay.Add "1;Current Price;1;11;0;1"
            astrDisplay.Add "0;Bid;1;12;0;1"
            astrDisplay.Add "0;Ask;1;13;0;1"
            astrDisplay.Add "0;Date;1;14;0;1"
            astrDisplay.Add "0;Account;1;15;0;1"
            astrDisplay.Add "0;Auto Trade Item;1;16;0;1"
            astrDisplay.Add "1;Broker ID;1;17;0;1"
            astrDisplay.Add "1;Link;1;18;0;1"
        
        Case Else
            'strReturn = "1;Filled;1;5,1;Status;1;6,0;Current Price;1;7,0;Bid;1;8,0;Ask;1;9,0;Date;1;10,0;Account;1;11,1;Auto Trade Item;1;12,0;Broker ID;1;13"
            astrDisplay.Add "0;Order ID;1;0;0;0"
            astrDisplay.Add "0;Account ID;1;1;0;0"
            astrDisplay.Add "0;Account Type;1;2;0;0"
            astrDisplay.Add "1;Symbol;1;3;0;0"
            astrDisplay.Add "0;Symbol ID;1;4;0;0"
            astrDisplay.Add "0;Remove;1;5;0;0"
            astrDisplay.Add "1;Working Order;1;6;0;0"
            astrDisplay.Add "1;X;1;7;0;0"
            astrDisplay.Add "1;M;1;8;0;0"
            astrDisplay.Add "1;Filled;1;9;0;1"
            astrDisplay.Add "1;Status;1;10;0;1"
            astrDisplay.Add "0;Current Price;1;11;0;1"
            astrDisplay.Add "0;Bid;1;12;0;1"
            astrDisplay.Add "0;Ask;1;13;0;1"
            astrDisplay.Add "0;Date;1;14;0;1"
            astrDisplay.Add "0;Account;1;15;0;1"
            astrDisplay.Add "0;Auto Trade Item;1;16;0;1"
            astrDisplay.Add "1;Broker ID;1;17;0;1"
            astrDisplay.Add "0;Link;1;18;0;1"
    
    End Select
    
    SmartColumns = "1|" & astrDisplay.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.SmartColumns"
    
End Function

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
        astrDisplay.Add "0;Order ID;1;0;0;0"
        astrDisplay.Add "0;Account ID;1;1;0;0"
        astrDisplay.Add "0;Account Type;1;2;0;0"
        astrDisplay.Add "1;Symbol;1;3;0;0"
        astrDisplay.Add "0;Symbol ID;1;4;0;0"
        astrDisplay.Add "0;Remove;1;5;0;0"
        astrDisplay.Add "1;Working Order;1;6;0;0"
        astrDisplay.Add "1;X;1;7;0;0"
        astrDisplay.Add "1;M;1;8;0;0"
        astrDisplay.Add "1;Filled;1;9;0;1"
        astrDisplay.Add "1;Status;1;10;0;1"
        astrDisplay.Add "1;Current Price;1;11;0;1"
        astrDisplay.Add "1;Bid;1;12;0;1"
        astrDisplay.Add "1;Ask;1;13;0;1"
        astrDisplay.Add "1;Date;1;14;0;1"
        astrDisplay.Add "1;Account;1;15;0;1"
        astrDisplay.Add "1;Auto Trade Item;1;16;0;1"
        astrDisplay.Add "1;Broker ID;1;17;0;1"
        astrDisplay.Add "1;Link;1;18;0;1"
        
        'strReturn = "1;Filled;1;5,1;Status;1;6,1;Current Price;1;7,1;Bid;1;8,1;Ask;1;9,1;Date;1;10,1;Account;1;11,1;Auto Trade Item;1;12,1;Broker ID;1;13"
        strReturn = "1|" & astrDisplay.JoinFields(",")
    End If

    DefaultDisplay = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.DefaultDisplay"
    
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
    End If
    
    FixDisplay = strDisplay

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cWorkingOrdersUI.FixDisplay"
    
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
        SetIniFileProperty "SummaryOrdersDisplay", m.strSummaryDisplay, "TTSummary", g.strIniFile
    Else
        m.strDisplay = m.TCGrid.Display
        SetIniFileProperty "OrdersDisplay", m.strDisplay, "TTSummary", g.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.BuildDisplayString"
    
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
    Dim Order As cPtOrder               ' Order object
    
    With m_fgWorkingOrders
        bFound = False
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then
                Set Order = OrderForRow(lIndex)
                If Not Order Is Nothing Then
                    If Order.AutoTradeItemID > 0 Then
                        bFound = True
                        Exit For
                    End If
                End If
            End If
        Next lIndex
        
        .ColHidden(OrdersCol(eGDCol_AutoTradeItem)) = Not bFound
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.ShowAutoTradeColumns"
    
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
    RaiseError "cWorkingOrdersUI.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_BeforeSort
'' Description: Handle the user sorting a column
'' Inputs:      Column, Sort Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SortOnCol Col, Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_BeforeSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_AfterMoveColumn
'' Description: After a user moves a column, resave the display string
'' Inputs:      Column moved, Position moved to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_AfterMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_AfterRowColChange
'' Description: After the user moves in the grid, enable/disable controls
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With m_fgWorkingOrders
        'if no horiz scrollbars, we need to "manually scroll"
        'by forcing the cell to show when changing columns
        If .Redraw <> flexRDNone And .ScrollBars <> flexScrollBarBoth Then
            If (OldCol <> NewCol) Then
                If (NewCol >= 0) And (NewRow >= 0) Then
                    .ShowCell NewRow, NewCol
                End If
            End If
        End If
    End With

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_AfterSort
'' Description: Set the background colors after a sort
'' Inputs:      Column of Sort, Order of Sort
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors m_fgWorkingOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_AfterUserResize
'' Description: Rebuild the display string after a user resize
'' Inputs:      Row and Column of resize
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_AfterUserResize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_BeforeMouseDown
'' Description: Bring up the Pop-Up menu when the user right clicks
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    m.TCGrid.SelectRow m_fgWorkingOrders.MouseRow

    If Button = vbRightButton Then
        LoadOrderGroupMenu
        EnableControls
        m.frm.PopupMenu m.mnuOrders
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_BeforeMoveColumn
'' Description: Make sure that the frozen columns stay where they are
'' Inputs:      Column to move, Position to move it to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        With m_fgWorkingOrders
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
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_BeforeMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_BeforeScroll
'' Description: Make sure left col stays the same if no horizontal
'' Inputs:      Old Top Row, Old Left Col, New Top Row, New Left Col, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    GridScrollCheck m_fgWorkingOrders, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_Click
'' Description: When the user clicks in the "X" column, attempt to cancel the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_Click()
On Error GoTo ErrSection:

    Dim lMouseCol As Long               ' Current mouse column in the grid
    Dim lMouseRow As Long               ' Current mouse row in the grid
    
    With m_fgWorkingOrders
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        If (lMouseCol = OrdersCol(eGDCol_Cancel)) Then
            If m.TCGrid.SelectRow(lMouseRow) Then
                CancelItemForRow
            End If
        ElseIf (lMouseCol = OrdersCol(eGDCol_Modify)) Then
            If m.TCGrid.SelectRow(lMouseRow) Then
                ModifyItemForRow
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_Compare
'' Description: Perform a comparison for the two rows for sorting purposes
'' Inputs:      Row 1, Row 2, Compare Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
On Error GoTo ErrSection:

    Dim strRow1 As String
    Dim strRow2 As String
    
    strRow1 = m_fgWorkingOrders.TextMatrix(Row1, OrdersCol(eGDCol_SortKey))
    strRow2 = m_fgWorkingOrders.TextMatrix(Row2, OrdersCol(eGDCol_SortKey))
    
    If strRow1 = strRow2 Then
        Cmp = 0
    ElseIf strRow1 < strRow2 Then
        Cmp = -1
    Else
        Cmp = 1
    End If
    
    If m.nSortedDir = flexSortStringDescending Then
        Cmp = Cmp * -1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_Compare"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_DblClick
'' Description: On a double click, change the active chart symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_DblClick()
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    
    Set Order = SelectedOrder
    If Not Order Is Nothing Then
        SetActiveChartSymbol Order.SymbolOrSymbolID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_DblClick"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_KeyUp
'' Description: Delete cancels an order, Insert creates a new one
'' Inputs:      Code of the Key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            If m.TCGrid.ValidRow(m_fgWorkingOrders.Row) Then
                CancelItemForRow
            End If
        
        Case vbKeyInsert
            CreateOrder "", 0&, , , , m.strSource
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgWorkingOrders_KeyUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgWorkingOrders_MouseMove
'' Description: Set the tooltip text based on where the cursor is
'' Inputs:      Mouse Button Down, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgWorkingOrders_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseCol As Long               ' Column that the mouse is over
    Dim lMouseRow As Long               ' Row that the mouse is over
    Dim strTooltipText As String        ' Tooltip text
    
    strTooltipText = ""
    With m_fgWorkingOrders
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        If m.TCGrid.ValidRow(lMouseRow) Then
            If .MouseCol = OrdersCol(eGDCol_Cancel) Then
                If TypeOf .RowData(lMouseRow) Is cPtOrder Then
                    strTooltipText = "Click here to Cancel the order"
                ElseIf TypeOf .RowData(lMouseRow) Is cActiveTsOrderGroup Then
                    strTooltipText = "Click here to Cancel the TradeSense order group"
                End If
            ElseIf .MouseCol = OrdersCol(eGDCol_Modify) Then
                strTooltipText = "Click here to Modify the order"
            End If
        End If
        
        If .ToolTipText <> strTooltipText Then
            .ToolTipText = strTooltipText
        End If
    End With
    
    m.TCGrid.SetMouseCursor m.frm

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
    RaiseError "cWorkingOrdersUI.m_mnuAutoSizeColumns_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuBuy_Click
'' Description: Allow the user to Buy a security
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuBuy_Click()
On Error GoTo ErrSection:

    CreateOrder "", 0&, 1, , , m.strSource

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuBuy_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuCancelOrder_Click
'' Description: Allow the user to cancel the selected order from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuCancelOrder_Click()
On Error GoTo ErrSection:

    CancelItemForRow

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuCancelOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuCheckStatus_Click
'' Description: Check the status of any open orders from the trade server
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
    RaiseError "cWorkingOrdersUI.m_mnuCheckStatus_Click"

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
    RaiseError "cWorkingOrdersUI.m_mnuDefaultColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuEditOrder_Click
'' Description: Allow the user to edit the selected order from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuEditOrder_Click()
On Error GoTo ErrSection:

    ModifyItemForRow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuEditOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuManageXOS_Click
'' Description: Allow the user to manage exit order strategies
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuManageXOS_Click()
On Error GoTo ErrSection:

    If SelectedTsOrder Is Nothing Then
        StartMenuTimer "MANAGEXOS"
    Else
        StartMenuTimer "MANAGETSOG"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuManageXOS_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuNewJournal_Click
'' Description: Allow the user to view the journal entries for this order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuNewJournal_Click()
On Error GoTo ErrSection:

    StartMenuTimer "JOURNAL"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuNewJournal_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuOrderGroups_Click
'' Description: Load up the order group menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuOrderGroups_Click()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuOrderGroups_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuOrderHistory_Click
'' Description: Allow the user to view the history for the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuOrderHistory_Click()
On Error GoTo ErrSection:

    StartMenuTimer "ORDERHISTORY"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuOrderHistory_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuParkOrder_Click
'' Description: Allow the user to park the selected order in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuParkOrder_Click()
On Error GoTo ErrSection:

    ParkItemForRow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "m_mnuParkOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuPrint_Click
'' Description: Allow the user to print the form
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
    RaiseError "cWorkingOrdersUI.m_mnuPrint_Click"

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

    Dim Order As cPtOrder               ' Selected order object
    
    Set Order = SelectedOrder
    If Not Order Is Nothing Then
        g.OrderStrategies.DeactivateExit Order.AccountID, Order.SymbolOrSymbolID, True, "Turned off from " & m.strSource
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuRemoveXOS_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSelectXOS_Click
'' Description: Allow the user to select an exit strategy order
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
    RaiseError "cWorkingOrdersUI.m_mnuSelectXOS_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSell_Click
'' Description: Allow the user to Sell a security
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuSell_Click()
On Error GoTo ErrSection:

    CreateOrder "", 0&, 0, , , m.strSource

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuSell_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSettings_Click
'' Description: Allow the user to view and edit their console settings
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
    RaiseError "cWorkingOrdersUI.m_mnuSettings_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSubmitAll_Click
'' Description: Allow the user to submit all of the parked orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuSubmitAll_Click()
On Error GoTo ErrSection:

    SubmitAllOrdersFromGrid m_fgWorkingOrders, m.strSource

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuSubmitAll_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSubmitOrder_Click
'' Description: Allow the user to submit the selected parked order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuSubmitOrder_Click()
On Error GoTo ErrSection:

    SubmitItemForRow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_mnuSubmitOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuTradeHistory_Click
'' Description: Allow the user to view trade history for the given account
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
    RaiseError "cWorkingOrdersUI.m_mnuTradeHistory_Click"
    
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
    RaiseError "cWorkingOrdersUI.m_mnuViewJournals_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_tmrMenu_Timer
'' Description: Perform a specified action when the timer goes off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Action to perform
    Dim Order As cPtOrder               ' Selected order from the grid

    strTag = m_tmrMenu.Tag
    If Len(strTag) > 2 Then
        If Left(strTag, 2) = kMenuPrefix Then
            strTag = Mid(strTag, 3)
            m_tmrMenu.Tag = ""
            m_tmrMenu.Enabled = False
            
            Set Order = SelectedOrder
            
            Select Case UCase(strTag)
                Case "JOURNAL"
                    If Not Order Is Nothing Then
                        g.TnJournal.ShowOrderJournal Order
                    End If
                                
                Case "JOURNALS"
                    g.TnJournal.ShowJournals
                    
                Case "MANAGEXOS"
                    frmOrderStrategies.ShowMeManage
                    
                Case "MANAGETSOG"
                    frmTradeSenseOrderGroups.ShowMe
                    
                Case "ORDERGROUPS"
                    frmOrderGroups.ShowMe
                    
                Case "ORDERHISTORY"
                    If Not Order Is Nothing Then
                        frmOrderHistory.ShowMe Order
                    End If
                    
                Case "SELECTXOS"
                    If Not Order Is Nothing Then
                        SelectXOS Order.AccountID, Order.SymbolOrSymbolID
                    End If
                    
                Case "SETTINGS"
                    frmTTSummaryCfg.ShowMe
                    
                Case "TRADEHISTORY"
                    If Order Is Nothing Then
                        frmTTAccounts.ShowMe True
                    Else
                        frmTTPositions.ShowMe Order.AccountID, Order.Broker
                    End If
                    
            End Select
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_tmrMenu_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelItemForRow
'' Description: Cancel the item in the selected row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelItemForRow()
On Error GoTo ErrSection:

    With m_fgWorkingOrders
        If TypeOf .RowData(.Row) Is cPtOrder Then
            CancelOrderFromGrid m_fgWorkingOrders, m.strSource, True
        ElseIf TypeOf .RowData(.Row) Is cActiveTsOrderGroup Then
            g.TsoGroups.CancelSubmittedGroup .RowData(.Row), "User cancelled from " & m.strSource
        End If
    End With
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.CancelItemForRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyItemForRow
'' Description: Modify the item in the selected row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ModifyItemForRow()
On Error GoTo ErrSection:

    With m_fgWorkingOrders
        If TypeOf .RowData(.Row) Is cPtOrder Then
            EditOrderFromGrid m_fgWorkingOrders, m.strSource
        ElseIf TypeOf .RowData(.Row) Is cActiveTsOrderGroup Then
        End If
    End With
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.ModifyItemForRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkItemForRow
'' Description: Park the item in the selected row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParkItemForRow()
On Error GoTo ErrSection:

    With m_fgWorkingOrders
        If TypeOf .RowData(.Row) Is cPtOrder Then
            ParkOrderFromGrid m_fgWorkingOrders, m.strSource
        ElseIf TypeOf .RowData(.Row) Is cActiveTsOrderGroup Then
            g.TsoGroups.ParkSubmittedGroup .RowData(.Row), "User parking group from " & m.strSource
        End If
    End With
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.ParkItemForRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitItemForRow
'' Description: Submit the item in the selected row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitItemForRow()
On Error GoTo ErrSection:

    With m_fgWorkingOrders
        If TypeOf .RowData(.Row) Is cPtOrder Then
            SubmitOrderFromGrid m_fgWorkingOrders, m.strSource
        ElseIf TypeOf .RowData(.Row) Is cActiveTsOrderGroup Then
            g.TsoGroups.SubmitParkedGroup .RowData(.Row)
        End If
    End With
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.SubmitItemForRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortOnCol
'' Description: Sort the grid for the given column number and order
'' Inputs:      Column, Sort Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SortOnCol(Optional ByVal lCol As Long = kNullData, Optional ByVal nOrder As SortSettings = kNullData)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim bAscending As Boolean           ' Do we want to sort ascending?
    Dim lIndex As Long                  ' Index into a for loop
    Dim strValue As String              ' Value for the column being sorted
    Dim strFormat As String             ' Format for the column

    If lCol = kNullData Then
        If m.lSortedCol = -1& Then
            lCol = OrdersCol(eGDCol_Symbol)
        Else
            lCol = m.lSortedCol
        End If
    End If
    
    strFormat = ""
    Select Case m_fgWorkingOrders.ColData(lCol)
        Case GDCol(eGDCol_CurrentPrice), GDCol(eGDCol_CurrentBid), GDCol(eGDCol_CurrentAsk), GDCol(eGDCol_NumFilled)
            strFormat = "Number"
        Case GDCol(eGDCol_Date)
            strFormat = "Date"
    End Select
    
    If nOrder = kNullData Then
        If m.nSortedDir = -1& Then
            nOrder = flexSortStringAscending
        Else
            nOrder = m.nSortedDir
        End If
    End If

    If (nOrder = flexSortGenericAscending) Or (nOrder = flexSortNumericAscending) Or (nOrder = flexSortStringAscending) Or (nOrder = flexSortStringNoCaseAscending) Then
        bAscending = True
    Else
        bAscending = False
    End If

    With m_fgWorkingOrders
        If .Rows > .FixedRows Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            strValue = ""
            For lIndex = .FixedRows To .Rows - 1
                If .RowOutlineLevel(lIndex) = 1 Then
                    Select Case UCase(strFormat)
                        Case "NUMBER", "DATE"
                            strValue = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.0000000"), 50, "R")
                        Case "CURRENCY"
                            strValue = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.00"), 50, "R")
                        Case Else
                            strValue = Pad(.TextMatrix(lIndex, lCol), 50, "L")
                    End Select
                End If
                
                If bAscending Then
                    .TextMatrix(lIndex, OrdersCol(eGDCol_SortKey)) = strValue & "_" & .TextMatrix(lIndex, OrdersCol(eGDCol_AscSortKey))
                Else
                    .TextMatrix(lIndex, OrdersCol(eGDCol_SortKey)) = strValue & "_" & .TextMatrix(lIndex, OrdersCol(eGDCol_DescSortKey))
                End If
                
                .RowOutlineLevel(lIndex) = 0
                .IsSubtotal(lIndex) = False
            Next lIndex
            
            .Select .FixedRows, OrdersCol(eGDCol_SortKey), .Rows - 1, OrdersCol(eGDCol_SortKey)
            If bAscending Then
                m.nSortedDir = flexSortStringAscending
            Else
                m.nSortedDir = flexSortStringDescending
            End If
            .Sort = flexSortCustom
            .Select .FixedRows, 0
            
            For lIndex = .FixedRows To .Rows - 1
                .RowOutlineLevel(lIndex) = CLng(Val(.TextMatrix(lIndex, OrdersCol(eGDCol_OutlineLevel))))
                .IsSubtotal(lIndex) = True
            Next lIndex
            
            If m.lSortedCol > -1& Then
                .Cell(flexcpPicture, 0, m.lSortedCol) = Nothing
            End If
            If bAscending Then
                .Cell(flexcpPicture, 0, lCol) = Picture16("kSortedUpRight")
            Else
                .Cell(flexcpPicture, 0, lCol) = Picture16("kSortedDownRight")
            End If
            
            .Cell(flexcpPictureAlignment, 0, lCol) = flexPicAlignRightTop
            .PicturesOver = True
            
            m.lSortedCol = lCol
            
            .AutoSize 0, .Cols - 1, False, 75
            .Redraw = nRedraw
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.SortOnCol"
    
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

    Set m.SnapshotOrders = New cGdTree

    For lIndex = 1 To kNumBrokers
        Set m.SnapshotOrders(lIndex) = Nothing
    Next lIndex
    
    m.dGroupsLastModified = kNullData

    m.lSortedCol = -1&
    m.nSortedDir = -1&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.Class_Initialize"
    
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
    Set m.SnapshotOrders = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.Class_Terminate"
    
End Sub
