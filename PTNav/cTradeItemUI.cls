VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeItemUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeItemUI.cls
'' Description: Wrapper class for an auto trade item grid
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
'' 06/16/2010   DAJ         Added Click here... row
'' 06/18/2010   DAJ         Draw Active Check Box for new trading item (#5794)
'' 08/05/2010   DAJ         Update all rows in the grid on a refresh trading item (#5842, #5843)
'' 09/16/2010   DAJ         Handle the Manage row after a sort (#5917)
'' 10/01/2010   DAJ         Changed "Open Pos" to "Actual Pos", remove qty on Strategy Pos (#5959)
'' 06/28/2011   DAJ         Setup clickable cells like hyperlinks
'' 07/15/2011   DAJ         Enhancements for allowing auto trading on continuous contracts
'' 07/21/2011   DAJ         Changed automated trading item enable
'' 07/22/2011   DAJ         Handle user changing symbol or account on an auto trade item
'' 07/25/2011   DAJ         Enhancements for deleting an auto trade item
'' 08/02/2011   DAJ         Allow user to change position by reassigning fills
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 06/28/2012   DAJ         Added contract month on position columns
'' 08/07/2012   DAJ         Only show number of enabled trade items on the trade console
'' 08/08/2012   DAJ         Don't call EditCell on the manage row
'' 08/14/2012   DAJ         Display 'Pending Data' in strategy position if waiting for data
'' 11/13/2012   DAJ         Log in the automated trading log when user chooses to Flatten
'' 12/19/2012   DAJ         Fix for changing the on-close time from the grid
'' 02/20/2013   DAJ         Added "Actual Performance" menu item, renamed "Strategy Performance"
'' 03/14/2013   DAJ         Show trade report filter dialog from right-click
'' 04/03/2013   DAJ         Automated Strategy Baskets
'' 04/15/2013   DAJ         Only allow auto exits, TSOG, and automated trading if enabled for streaming
'' 04/18/2013   DAJ         Only check for streaming if user activates automated basket
'' 05/14/2013   DAJ         Load a guru basket even if not the owner
'' 05/15/2013   DAJ         Validate user entered quantity next entry
'' 05/22/2013   DAJ         When loading the grid, load the baskets collection once for speed
'' 05/22/2013   DAJ         Only load baskets if we need to when loading the grid
'' 05/28/2013   DAJ         Do a pre-flight check of the broker status when activating
'' 06/04/2013   DAJ         Persist expanded/collapsed settings
'' 06/07/2013   DAJ         Ucase the sort key to make the sort case insensitive
'' 06/12/2013   DAJ         Symbol and quantity validation for automated trading items
'' 03/05/2014   DAJ         Removed unused reference to cLotColumn
'' 04/04/2014   DAJ         Allow automated trading for pyramiding systems
'' 05/06/2014   DAJ         Show FractZen period; Handle data being behind on an item
'' 06/11/2014   DAJ         Dump the automated trading items grid to a file if it changes
'' 07/11/2014   DAJ         Added the Max Quantity column
'' 10/15/2014   DAJ         Allow user to disable all trade items with a Ctrl-click
'' 01/23/2015   DAJ         Don't allow user to move the Active column
'' 01/23/2015   DAJ         Disable All menu item; Force click in check box
'' 05/12/2015   DAJ         Fix for "Object not found..." error in TradeItemToGrid
'' 08/05/2015   DAJ         Don't allow closing time to be outside of market hours
'' 01/14/2016   DAJ         Fix for new symbol group/strategy basket/automated trading item getting sorted incorrectly
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kMenuPrefix = "T:"
Private Const kManageRowText = "Click here to create a new Automated Trading Item"

Private WithEvents m_fgTradeItems As VSFlexGrid
Attribute m_fgTradeItems.VB_VarHelpID = -1
Private WithEvents m_tmrMenu As Timer
Attribute m_tmrMenu.VB_VarHelpID = -1
Private WithEvents m_mnuDisableAll As Menu
Attribute m_mnuDisableAll.VB_VarHelpID = -1
Private WithEvents m_mnuFlatten As Menu
Attribute m_mnuFlatten.VB_VarHelpID = -1
Private WithEvents m_mnuEnterPosition As Menu
Attribute m_mnuEnterPosition.VB_VarHelpID = -1
Private WithEvents m_mnuChangePosition As Menu
Attribute m_mnuChangePosition.VB_VarHelpID = -1
Private WithEvents m_mnuNewTradeItem As Menu
Attribute m_mnuNewTradeItem.VB_VarHelpID = -1
Private WithEvents m_mnuEditTradeItem As Menu
Attribute m_mnuEditTradeItem.VB_VarHelpID = -1
Private WithEvents m_mnuDeleteTradeItem As Menu
Attribute m_mnuDeleteTradeItem.VB_VarHelpID = -1
Private WithEvents m_mnuRollContracts As Menu
Attribute m_mnuRollContracts.VB_VarHelpID = -1
Private WithEvents m_mnuEditStrategy As Menu
Attribute m_mnuEditStrategy.VB_VarHelpID = -1
Private WithEvents m_mnuStrategyPerformance As Menu
Attribute m_mnuStrategyPerformance.VB_VarHelpID = -1
Private WithEvents m_mnuActualPerformance As Menu
Attribute m_mnuActualPerformance.VB_VarHelpID = -1
Private WithEvents m_mnuNextBarReport As Menu
Attribute m_mnuNextBarReport.VB_VarHelpID = -1
Private WithEvents m_mnuShowChart As Menu
Attribute m_mnuShowChart.VB_VarHelpID = -1
Private WithEvents m_mnuPrint As Menu
Attribute m_mnuPrint.VB_VarHelpID = -1
Private WithEvents m_mnuTradeHistory As Menu
Attribute m_mnuTradeHistory.VB_VarHelpID = -1
Private WithEvents m_mnuAlerts As Menu
Attribute m_mnuAlerts.VB_VarHelpID = -1
Private WithEvents m_mnuAutoSizeColumns As Menu
Attribute m_mnuAutoSizeColumns.VB_VarHelpID = -1
Private WithEvents m_mnuDefaultColumns As Menu
Attribute m_mnuDefaultColumns.VB_VarHelpID = -1

Private Enum eGDCols
    eGDCol_Active = 0
    eGDCol_Name
    eGDCol_Strategy
    eGDCol_Symbol
    eGDCol_Period
    eGDCol_Account
    eGDCol_StrategyPos
    eGDCol_OpenPos
    eGDCol_QtyNextEntry
    eGDCol_MaxQty
    eGDCol_OpenOrders
    eGDCol_OnCloseTime
    
    eGDCol_ItemBasketLastModified
    eGDCol_BasketLastModified
    eGDCol_Information
    eGDCol_SymbolError
    eGDCol_QuantityError
    
    eGDCol_SortKey
    eGDCol_AscSortKey
    eGDCol_DescSortKey
    eGDCol_OutlineLevel
    
    eGDCol_NumCols
End Enum

Private Type mPrivate
    strSource As String                 ' Name for the calling form
    TCGrid As cTradeConsoleGrid         ' Trade console grid object for generic routines
    tmrRealTime As Timer                ' Realtime timer control
    frm As Form                         ' Form that the grid is on
    mnuTradeItems As Object             ' Trade items menu

    lSortedCol As Long                  ' Sorted column
    nSortedDir As SortSettings          ' Sort direction for the column
    
    bLoadingExpanded As Boolean         ' Are we currently loading the expanded settings?
    strGrid As String                   ' Last known state of the grid
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object
'' Inputs:      Source, Trade Items UI
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal strSource As String, UI As cTradeItemContols)
On Error GoTo ErrSection:

    Dim strDisplay As String            ' Display string from the INI file
    
    m.strSource = strSource
    
    With UI
        Set m.frm = .frm
        Set m_fgTradeItems = .fgGrid
        Set m.tmrRealTime = .tmrRealTime
        Set m_tmrMenu = .tmrMenu
        Set m.mnuTradeItems = .mnuTradeItems
        Set m_mnuDisableAll = .mnuDisableAll
        Set m_mnuFlatten = .mnuFlatten
        Set m_mnuEnterPosition = .mnuEnterPosition
        Set m_mnuChangePosition = .mnuChangePosition
        Set m_mnuNewTradeItem = .mnuNewTradeItem
        Set m_mnuEditTradeItem = .mnuEditTradeItem
        Set m_mnuDeleteTradeItem = .mnuDeleteTradeItem
        Set m_mnuRollContracts = .mnuRollContract
        Set m_mnuEditStrategy = .mnuEditStrategy
        Set m_mnuStrategyPerformance = .mnuStrategyPerformance
        Set m_mnuActualPerformance = .mnuActualPerformance
        Set m_mnuNextBarReport = .mnuNextBarReport
        Set m_mnuShowChart = .mnuShowChart
        Set m_mnuPrint = .mnuPrint
        Set m_mnuTradeHistory = .mnuTradeHistory
        Set m_mnuAlerts = .mnuAlerts
        Set m_mnuAutoSizeColumns = .mnuAutoSizeColumns
        Set m_mnuDefaultColumns = .mnuDefaultColumns
    End With

    InitGrid
    
    strDisplay = FixDisplay(GetIniFileProperty("TradeItemsDisplay", "", "TTSummary", g.strIniFile))
    
    Set m.TCGrid = New cTradeConsoleGrid
    m.TCGrid.Init m_fgTradeItems, m.tmrRealTime, strDisplay
    
    If Not g.TradingItems Is Nothing Then
        g.TradingItems.FixAccountPositionIds
    End If
    
    LoadGrid
    
    ApplyDisplayString
    
    If InStr(strDisplay, "Max Qty") = 0 Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_MaxQty)) = TradeItemCol(eGDCol_QtyNextEntry) + 1
    End If
    
    If TradeItemCol(eGDCol_ItemBasketLastModified) <> GDCol(eGDCol_ItemBasketLastModified) Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_ItemBasketLastModified)) = GDCol(eGDCol_ItemBasketLastModified)
    End If
    If TradeItemCol(eGDCol_BasketLastModified) <> GDCol(eGDCol_BasketLastModified) Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_BasketLastModified)) = GDCol(eGDCol_BasketLastModified)
    End If
    If TradeItemCol(eGDCol_Information) <> GDCol(eGDCol_Information) Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_Information)) = GDCol(eGDCol_Information)
    End If
    If TradeItemCol(eGDCol_SymbolError) <> GDCol(eGDCol_SymbolError) Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_SymbolError)) = GDCol(eGDCol_SymbolError)
    End If
    If TradeItemCol(eGDCol_QuantityError) <> GDCol(eGDCol_QuantityError) Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_QuantityError)) = GDCol(eGDCol_QuantityError)
    End If
    If TradeItemCol(eGDCol_SortKey) <> GDCol(eGDCol_SortKey) Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_SortKey)) = GDCol(eGDCol_SortKey)
    End If
    If TradeItemCol(eGDCol_AscSortKey) <> GDCol(eGDCol_AscSortKey) Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_AscSortKey)) = GDCol(eGDCol_AscSortKey)
    End If
    If TradeItemCol(eGDCol_DescSortKey) <> GDCol(eGDCol_DescSortKey) Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_DescSortKey)) = GDCol(eGDCol_DescSortKey)
    End If
    If TradeItemCol(eGDCol_OutlineLevel) <> GDCol(eGDCol_OutlineLevel) Then
        m_fgTradeItems.ColPosition(TradeItemCol(eGDCol_OutlineLevel)) = GDCol(eGDCol_OutlineLevel)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.Init"
    
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

    m.TCGrid.ApplyDisplayString GDCol(eGDCol_Active) + 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.ApplyDisplayString"
    
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
    FontFromString m_fgTradeItems.Font, g.Broker.GridFont
    m_fgTradeItems.Font = m_fgTradeItems.Font
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.UpdateConsoleSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedTradeItem
'' Description: Grab the selected trade item in the grid
'' Inputs:      None
'' Returns:     Trade Item (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectedTradeItem() As cAutoTradeItem
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Trade item from the grid

    Set TradeItem = Nothing
    With m_fgTradeItems
        If m.TCGrid.ValidRow(.Row) Then
            If TypeOf .RowData(.Row) Is cAutoTradeItem Then
                Set TradeItem = .RowData(.Row)
            End If
        End If
    End With
    
    Set SelectedTradeItem = TradeItem

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.SelectedTradeItem"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeItemForRow
'' Description: Grab the trade item in the grid on the given row
'' Inputs:      Row
'' Returns:     Trade Item (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradeItemForRow(ByVal Row As Long) As cAutoTradeItem
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Trade item from the grid

    Set TradeItem = Nothing
    With m_fgTradeItems
        If m.TCGrid.ValidRow(Row) Then
            If TypeOf .RowData(Row) Is cAutoTradeItem Then
                Set TradeItem = .RowData(Row)
            End If
        End If
    End With
    
    Set TradeItemForRow = TradeItem

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.TradeItemForRow"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterTradeItemsGrid
'' Description: Hide auto trade items that we do not want showing
'' Inputs:      Save Expanded Settings?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FilterTradeItemsGrid(Optional ByVal bSaveExpanded As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Trade Item object
    Dim lAccountID As Long              ' Account ID for the row in the grid
    Dim lNumVisible As Long             ' Number of visible items
    Dim bHideMaxQty As Boolean          ' Hide the max quantity column?
    
    With m_fgTradeItems
        .Redraw = flexRDNone
        
        lNumVisible = 0&
        bHideMaxQty = True
        
        For lIndex = .FixedRows To .Rows - 1
            Set TradeItem = TradeItemForRow(lIndex)
            If Not TradeItem Is Nothing Then
                .RowHidden(lIndex) = g.Broker.HideAccount(TradeItem.AccountID)
                If .RowHidden(lIndex) = False Then
                    If CheckedCell(m_fgTradeItems, lIndex, TradeItemCol(eGDCol_Active)) Then
                        lNumVisible = lNumVisible + 1&
                    End If
                End If
                
                If .MergeRow(lIndex) = False Then
                    If Len(.TextMatrix(lIndex, TradeItemCol(eGDCol_MaxQty))) > 0 Then
                        bHideMaxQty = False
                    End If
                End If
            End If
        Next lIndex
        
        .ColHidden(TradeItemCol(eGDCol_MaxQty)) = bHideMaxQty
        
        SortOnCol
        LoadExpandedSettings
        HandleManageRow
        SetBackColors m_fgTradeItems
        m.TCGrid.AutoSizeGrid
        
        If bSaveExpanded Then
            SaveExpandedSettings
        End If
        
        .Redraw = flexRDBuffered
    End With
    
    If Not g.ConsoleForms Is Nothing Then
        g.ConsoleForms.NumVisible(eGDConsoleForm_AutoTrading) = g.TradingItems.NumActive ' = lNumVisible
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.FilterTradeItemsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTradeItem
'' Description: Refreh the grid for the given trading item
'' Inputs:      Trading Item ID, Deleted?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshTradeItem(ByVal lAutoTradingItemID As Long, Optional ByVal bDeleted As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Temporary trade item
    
    With m_fgTradeItems
        For lIndex = .FixedRows To .Rows - 1
            Set TradeItem = TradeItemForRow(lIndex)
            If Not TradeItem Is Nothing Then
                If TradeItem.AutoTradeItemID = lAutoTradingItemID Then
                    If bDeleted Then
                        .RemoveItem lIndex
                    Else
                        TradeItemToGrid TradeItem, lIndex
                    End If
                
                    Exit For
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.RefreshTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditTradeItem
'' Description: Allow the user to edit the automated trading item of the row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EditTradeItem()
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated Trading Item to edit
    Dim Original As New cAutoTradeItem  ' Original version of automated trading item
    Dim bActive As Boolean              ' Is the automated trading item active?
    Dim lIndex As Long                  ' Index into a for loop
    Dim lParent As Long                 ' Parent row in the grid
    Dim BasketItems As cGdTree          ' Strategy basket trade items
    Dim lRow As Long                    ' Row in the grid
    
    With m_fgTradeItems
        If .RowOutlineLevel(.Row) = 1 Then
            .Row = .GetNodeRow(.Row, flexNTParent)
        End If
        
        Set TradeItem = SelectedTradeItem
        If Not TradeItem Is Nothing Then
            Select Case TradeItem.ParentID
                Case -1:
                    Original.Load TradeItem.AutoTradeItemID, False
                    bActive = CheckedCell(m_fgTradeItems, .Row, TradeItemCol(eGDCol_Active))

                    If frmAutoTradeItem.ShowMe(TradeItem, (bActive Or (ChildInPosition(.Row) <> 0&)), BasketItems) = True Then
                        UpdateTradeItem TradeItem, Original, .Row
                        
                        ' Walk through existing children and delete if they no longer exist...
                        lRow = .GetNodeRow(.Row, flexNTLastChild)
                        If lRow <> -1& Then
                            For lIndex = lRow To .Row + 1 Step -1
                                Set Original = .RowData(lIndex)
                                If BasketItems.Exists(Str(Original.AutoTradeItemID)) = False Then
                                    RemoveTradeItemForRow lIndex
                                End If
                            Next lIndex
                        End If
                        
                        ' Walk through new children and update or create as necessary...
                        For lIndex = 1 To BasketItems.Count
                            lRow = RowForTradeItem(BasketItems(lIndex))
                            If lRow = -1& Then
                                Set Original = BasketItems(lIndex)
                            Else
                                Set Original = .RowData(lRow)
                            End If
                            
                            UpdateTradeItem BasketItems(lIndex), Original, lRow
                        Next lIndex
                        
                        m.TCGrid.AutoSizeGrid
                        FilterTradeItemsGrid
                    End If
                
                Case 0:
                    Original.Load TradeItem.AutoTradeItemID, False
                    bActive = CheckedCell(m_fgTradeItems, .Row, TradeItemCol(eGDCol_Active))
                    
                    If frmAutoTradeItem.ShowMe(TradeItem, (bActive Or (TradeItem.CurrentPosition <> 0&))) = True Then
                        UpdateTradeItem TradeItem, Original, .Row
                        
                        m.TCGrid.AutoSizeGrid
                        FilterTradeItemsGrid
                    End If
            
            End Select
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.EditTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewTradeItem
'' Description: Allow the user to create an automated trading item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub NewTradeItem()
On Error GoTo ErrSection:

    Dim TradeItem As New cAutoTradeItem ' Automated Trading Item to edit
    Dim BasketItems As cGdTree          ' Collection of basket items for the given trade item
    Dim lIndex As Long                  ' Index into a for loop
    
    With m_fgTradeItems
        If frmAutoTradeItem.ShowMe(TradeItem, False, BasketItems) = True Then
            ' Do a load here so that the strategy gets loaded...
            If TradeItem.Load(TradeItem.AutoTradeItemID) = True Then
                g.TradingItems.Add TradeItem
                TradeItemToGrid TradeItem
                
                For lIndex = 1 To BasketItems.Count
                    Set TradeItem = BasketItems(lIndex)
                    If TradeItem.Load(TradeItem.AutoTradeItemID) = True Then
                        g.TradingItems.Add TradeItem
                        TradeItemToGrid TradeItem
                    End If
                Next lIndex
                
                m.TCGrid.AutoSizeGrid
                FilterTradeItemsGrid
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.NewTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveTradeItem
'' Description: Allow the user to remove an automated trading item
'' Inputs:      Confirm?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveTradeItem()
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated Trading Item to edit
    Dim bActive As Boolean              ' Is the automated trading item active?
    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return value from the InfBox
    Dim lChildRow As Long               ' Child row for the selected row
    
    Set TradeItem = SelectedTradeItem
    If Not TradeItem Is Nothing Then
        With m_fgTradeItems
            bActive = CheckedCell(m_fgTradeItems, .Row, TradeItemCol(eGDCol_Active))
            
            If bActive Then
                InfBox "You cannot remove an automated trading item while it is still active.", "!", , "Error"
            ElseIf TradeItem.CurrentPosition <> 0& Then
                InfBox "You cannot remove an automated trading item while it is still in a position.", "!", , "Error"
            ElseIf TradeItem.IsGuru = True Then
                InfBox "You cannot remove a Shadow Trading automated trading item", "!", , "Error"
            Else
                If TradeItem.ParentID = 0& Then
                    strReturn = InfBox("Are you sure you wish to remove|" & TradeItem.Name & "?", "?", "+Yes|-No", "Confirmation")
                Else
                    strReturn = InfBox("Are you sure you wish to remove||" & TradeItem.Name & "||and all of the automated trading|items associated with it?", "?", "+Yes|-No", "Confirmation")
                End If
                
                If strReturn = "Y" Then
                    If TradeItem.ParentID < 0& Then
                        lChildRow = .GetNodeRow(.Row, flexNTLastChild)
                        Do While lChildRow <> -1&
                            RemoveTradeItemForRow lChildRow
                            lChildRow = .GetNodeRow(.Row, flexNTLastChild)
                        Loop
                    End If
                    RemoveTradeItemForRow .Row
                    
                    FilterTradeItemsGrid
                End If
            End If
        End With
    
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.RemoveTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisableTradeItem
'' Description: Clean up and stop the selected trading item
'' Inputs:      Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisableTradeItem(Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated Trading Item to edit
    
    Set TradeItem = SelectedTradeItem
    If Not TradeItem Is Nothing Then
        TradeItem.Disable , , strReason
    
        ' Blank out the current strategy position...
        m_fgTradeItems.TextMatrix(m_fgTradeItems.Row, TradeItemCol(eGDCol_StrategyPos)) = ""
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.DisableTradeItem"
    
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

    m.TCGrid.GenerateReport vArgs, "Automated Trading Items"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls on the form as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EnableControls()
On Error GoTo ErrSection:

    Dim bValidRow As Boolean            ' Is the current row a valid row?
    Dim TradeItem As cAutoTradeItem     ' Currently selected item
    Dim bIsOwner As Boolean             ' Is the current user an owner of the object?
    
    bValidRow = m.TCGrid.CurrentRowValid
    Set TradeItem = SelectedTradeItem
    
    Enable m_mnuDisableAll, g.TradingItems.HasActiveAutoTradeItems
    
    If TradeItem Is Nothing Then
        Enable m_mnuFlatten, False
        Enable m_mnuEnterPosition, False
        Enable m_mnuChangePosition, False
        Enable m_mnuEditStrategy, False
        Enable m_mnuShowChart, False
        Enable m_mnuRollContracts, False
        Enable m_mnuDeleteTradeItem, False
        
        bIsOwner = False
    Else
        If TradeItem.ParentID >= 0& Then
            Enable m_mnuFlatten, (TradeItem.Active And (TradeItem.CurrentPosition <> 0))
            Enable m_mnuEnterPosition, (TradeItem.Active And (TradeItem.CurrentPosition = 0))
            Enable m_mnuChangePosition, (TradeItem.ParentID >= 0&)
                
            m_mnuEditStrategy.Caption = "Edit Strategy"
            Enable m_mnuEditStrategy, (TradeItem.StrategyID > 0)
            Enable m_mnuShowChart, (Len(TradeItem.Symbol) > 0)
            Enable m_mnuRollContracts, AllowRollContracts
        
            Enable m_mnuDeleteTradeItem, (TradeItem.ParentID = 0&)
        Else
            Enable m_mnuFlatten, False
            Enable m_mnuEnterPosition, False
            Enable m_mnuChangePosition, False
            
            m_mnuEditStrategy.Caption = "Edit Strategy Basket"
            Enable m_mnuEditStrategy, (TradeItem.StrategyBasketID > 0)
            Enable m_mnuShowChart, False
            Enable m_mnuRollContracts, False
        
            Enable m_mnuDeleteTradeItem, True
        End If
        
        bIsOwner = TradeItem.IsGuruOwner
    End If
        
    Enable m_mnuEditTradeItem, bValidRow
    
    Enable m_mnuStrategyPerformance, bValidRow
    Enable m_mnuActualPerformance, bValidRow
    Enable m_mnuNextBarReport, bValidRow And bIsOwner
    Enable m_mnuTradeHistory, bValidRow
    
    m_mnuDefaultColumns.Enabled = (m.TCGrid.Display <> DefaultDisplay)
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketChanged
'' Description: Notification that the given strategy basket has changed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BasketChanged(ByVal Basket As cStrategyBasket)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Trade item from the grid
    
    With m_fgTradeItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If IsManageRow(lIndex) = False Then
                If TypeOf .RowData(lIndex) Is cAutoTradeItem Then
                    Set TradeItem = .RowData(lIndex)
                    If TradeItem.StrategyBasketID = Basket.ID Then
                        If TradeItem.StrategyBasketLastModified <> Basket.LastModified Then
                            .TextMatrix(lIndex, TradeItemCol(eGDCol_BasketLastModified)) = Str(Basket.LastModified)
                        End If
                        
                        CheckLastModified lIndex
                    End If
                End If
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.BasketChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpGridIfDifferent
'' Description: Dump the contents of the grid to a file if it has changed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DumpGridIfDifferent()
On Error GoTo ErrSection:

    Dim strGrid As String               ' Contents of the grid
    Dim strFile As String
    Static bInProgress As Boolean       ' Are we already in progress?
    Static dLastTime As Double          ' Last time we checked
    
    If bInProgress = False Then
        bInProgress = True
        
        If gdTickCount > dLastTime + 1000 Then
            strFile = AddSlash(App.Path) & "AutoTradeItems.TXT"
            'If FileExist(strFile) Then
                strGrid = GridToString
                If strGrid <> m.strGrid Then
                    FileFromString strFile, strGrid
                    m.strGrid = strGrid
                End If
            'End If
            
            dLastTime = gdTickCount
        End If
        
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "cTradeItemUI.DumpGridIfDifferent"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the trade items grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    With m_fgTradeItems
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbButtonFace
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = False
        .MergeCells = flexMergeFree
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
        
        .TextMatrix(0, GDCol(eGDCol_Active)) = "Active"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_Strategy)) = "Strategy"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Period)) = "Period"
        .TextMatrix(0, GDCol(eGDCol_Account)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_StrategyPos)) = "Strategy Pos"
        .TextMatrix(0, GDCol(eGDCol_OpenPos)) = "Actual Pos"
        .TextMatrix(0, GDCol(eGDCol_QtyNextEntry)) = "Qty Next Entry"
        .TextMatrix(0, GDCol(eGDCol_MaxQty)) = "Max Qty"
        .TextMatrix(0, GDCol(eGDCol_OpenOrders)) = "Open Orders"
        .TextMatrix(0, GDCol(eGDCol_OnCloseTime)) = "On Close Time"
        .TextMatrix(0, GDCol(eGDCol_ItemBasketLastModified)) = "Item Last Mod"
        .TextMatrix(0, GDCol(eGDCol_BasketLastModified)) = "Basket Last Mod"
        .TextMatrix(0, GDCol(eGDCol_Information)) = "Information"
        .TextMatrix(0, GDCol(eGDCol_SymbolError)) = "Symbol Error"
        .TextMatrix(0, GDCol(eGDCol_QuantityError)) = "Quantity Error"
        .TextMatrix(0, GDCol(eGDCol_SortKey)) = "Sort Key"
        .TextMatrix(0, GDCol(eGDCol_AscSortKey)) = "Asc Sort Key"
        .TextMatrix(0, GDCol(eGDCol_DescSortKey)) = "Desc Sort Key"
        .TextMatrix(0, GDCol(eGDCol_OutlineLevel)) = "Outline Level"
        
        .ColHidden(GDCol(eGDCol_ItemBasketLastModified)) = True
        .ColHidden(GDCol(eGDCol_BasketLastModified)) = True
        .ColHidden(GDCol(eGDCol_Information)) = True
        .ColHidden(GDCol(eGDCol_SymbolError)) = True
        .ColHidden(GDCol(eGDCol_QuantityError)) = True
        .ColHidden(GDCol(eGDCol_SortKey)) = True
        .ColHidden(GDCol(eGDCol_AscSortKey)) = True
        .ColHidden(GDCol(eGDCol_DescSortKey)) = True
        .ColHidden(GDCol(eGDCol_OutlineLevel)) = True

        .ColAlignment(GDCol(eGDCol_Account)) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the trade items grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Trading item
    Dim lAcctPosID As Long              ' Account Position ID
    Dim Baskets As cStrategyBaskets     ' Strategy basket collection

    Set Baskets = Nothing

    With m_fgTradeItems
        .Redraw = flexRDNone
        
        If Not g.TradingItems Is Nothing Then
            For lIndex = 1 To g.TradingItems.Count
                Set TradeItem = g.TradingItems.Item(lIndex)
                If (TradeItem.AccountPositionID = 0) And (TradeItem.ParentID <> -1&) Then
                    TradeItem.AccountPositionID = g.Broker.CreateFillSummaryForAutoTrade(TradeItem)
                    TradeItem.Save
                End If
                
                If TradeItem.StrategyBasketID > 0& Then
                    If Baskets Is Nothing Then
                        Set Baskets = New cStrategyBaskets
                        Baskets.LoadDb True, True
                    End If
                End If
                
                TradeItemToGrid TradeItem, , Baskets
                
                .Cell(flexcpFontUnderline, .Rows - 1, 0, .Rows - 1, .Cols - 1) = False
                '.Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpForeColor, 0, 0)
            Next lIndex
        End If
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        FilterTradeItemsGrid False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.LoadGrid"
    
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

    Dim astrDisplay As New cGdArray     ' Array of field display information
    
    astrDisplay.Add "1;Active;1;0;0;1"
    astrDisplay.Add "1;Name;1;1;0;1"
    astrDisplay.Add "1;Strategy;1;2;0;1"
    astrDisplay.Add "1;Symbol;1;3;0;1"
    astrDisplay.Add "1;Period;1;4;0;1"
    astrDisplay.Add "1;Account;1;5;0;1"
    astrDisplay.Add "1;Strategy Pos;1;6;0;1"
    astrDisplay.Add "1;Open Pos;1;7;0;1"
    astrDisplay.Add "1;Qty Next Entry;1;8;0;1"
    astrDisplay.Add "1;Open Orders;1;9;0;1"
    astrDisplay.Add "1;On Close Time;1;10;0;1"
            
    'DefaultDisplay = "1;Active;1;0,1;Name;1;1,1;Strategy;1;2,1;Symbol;1;3,1;Period;1;4,1;Account;1;5,1;Strategy Pos;1;6,1;Open Pos;1;7,1;Qty Next Entry;1;8,1;Open Orders;1;9,1;On Close Time;1;10"
    DefaultDisplay = "1|" & astrDisplay.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.DefaultDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixDisplay
'' Description: Fix the given display string
'' Inputs:      Display
'' Returns:     Fixed Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixDisplay(ByVal strDisplay As String) As String
On Error GoTo ErrSection:

    If Len(strDisplay) = 0 Then
        strDisplay = DefaultDisplay
    End If
    
    FixDisplay = strDisplay

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.FixDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowChartForTradingItem
'' Description: Show the appropriate chart for an automated trading item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowChartForTradingItem()
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated Trading Item to edit
    Dim bFound As Boolean               ' Does this chart already exist?
    Dim lIndex As Long                  ' Index into a for loop
    Dim frm As Form                     ' Chart form
    
    With m_fgTradeItems
        Set TradeItem = SelectedTradeItem
        If Not TradeItem Is Nothing Then
            bFound = False
            
            For lIndex = 0 To Forms.Count - 1
                If IsFrmChart(Forms(lIndex)) Then
                    Set frm = Forms(lIndex)
                    If frm.Chart.SymbolID = TradeItem.SymbolID And frm.Chart.Bars.Prop(eBARS_PeriodicityStr) = TradeItem.BarPeriod Then
                        If frm.Chart.SystemID = TradeItem.StrategyID Then
                            ShowForm frm
                            bFound = True
                            Exit For
                        End If
                    End If
                End If
            Next lIndex
            
            If bFound = False Then
                Set frm = New frmChart              'new chart is always non-detached
                With frm
                    .Chart.SetSymbol TradeItem.SymbolID
                    .Chart.ChangeBarPeriod TradeItem.BarPeriod
                    .Chart.SystemID = TradeItem.StrategyID
                    .Chart.ShowTrades = True
                    .Chart.GenerateChart
                    ShowForm frm
                End With
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.ShowChartForTradingItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowRollContracts
'' Description: Do we allow rolling for this automated trading item?
'' Inputs:      None
'' Returns:     True if Allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllowRollContracts() As Boolean
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated Trading Item to edit
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    Set TradeItem = SelectedTradeItem
    If Not TradeItem Is Nothing Then
        If SecurityType(TradeItem.Symbol) = "F" Then
            bReturn = (InStr(TradeItem.Symbol, "-0") = 0)
        End If
    End If
    
    AllowRollContracts = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.AllowRollContracts"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RollContracts
'' Description: Allow the user to roll into another contract
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RollContracts()
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated Trading Item to edit
    Dim Original As New cAutoTradeItem  ' Original version of Automated Trading Item
    
    Set TradeItem = SelectedTradeItem
    If Not TradeItem Is Nothing Then
        Original.Load TradeItem.AutoTradeItemID, False
        TradeItem.RollContracts
        RefreshTradeItem TradeItem.AutoTradeItemID
        
        g.Broker.CreateFillSummaryForAutoTrade TradeItem
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.RollContracts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeItemCol
'' Description: Determine where the given column is really sitting right now
'' Inputs:      Column Number
'' Returns:     Current Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TradeItemCol(ByVal Col As eGDCols) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    If Col = eGDCol_NumCols Then
        lReturn = Col
    Else
        lReturn = m.TCGrid.ColNumFromColData(Col)
    End If
    
    TradeItemCol = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.TradeItemCol"
    
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
    
    astrColumns.Add Str(TradeItemCol(eGDCol_Active))
    astrColumns.Add Str(TradeItemCol(eGDCol_ItemBasketLastModified))
    astrColumns.Add Str(TradeItemCol(eGDCol_BasketLastModified))
    astrColumns.Add Str(TradeItemCol(eGDCol_Information))
    astrColumns.Add Str(TradeItemCol(eGDCol_SymbolError))
    astrColumns.Add Str(TradeItemCol(eGDCol_QuantityError))
    astrColumns.Add Str(TradeItemCol(eGDCol_SortKey))
    astrColumns.Add Str(TradeItemCol(eGDCol_AscSortKey))
    astrColumns.Add Str(TradeItemCol(eGDCol_DescSortKey))
    astrColumns.Add Str(TradeItemCol(eGDCol_OutlineLevel))
    
    ColumnExceptions = "," & astrColumns.JoinFields(",") & ","

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.ColumnExceptions"
    
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
    SetIniFileProperty "TradeItemsDisplay", m.TCGrid.Display, "TTSummary", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.BuildDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_AfterCollapse
'' Description: Handle the user expanding or collapsing a tree node
'' Inputs:      Row, State
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
On Error GoTo ErrSection:

    If m.bLoadingExpanded = False Then
        SaveExpandedSettings
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_AfterCollapse"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_AfterEdit
'' Description: Handle the user's changes appropriately
'' Inputs:      Row and Column of Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Temporary automated trading object
    Dim lNewQty As Long                 ' New Quantity of the Next Entry
    Dim lOldQty As Long                 ' Old Quantity of the Next Entry

    If m.frm.Visible Then
        m.TCGrid.SelectRow Row
    
        Select Case Col
            Case TradeItemCol(eGDCol_Active)
                ToggleTradeItem Row
            
            Case TradeItemCol(eGDCol_QtyNextEntry)
                Set TradeItem = SelectedTradeItem
                If Not TradeItem Is Nothing Then
                    lNewQty = CLng(Val(m_fgTradeItems.TextMatrix(Row, Col)))
                    lOldQty = TradeItem.QtyNextEntry
                    If lNewQty <> lOldQty Then
                        TradeItem.QtyNextEntry = lNewQty
                        TradeItem.DumpDebug "User changed Quantity Next Entry from Trade Console.  Old = " & Str(lOldQty) & "; New = " & Str(lNewQty)
                        TradeItem.Save
                        
                        If CheckedCell(m_fgTradeItems, Row, TradeItemCol(eGDCol_Active)) = True Then
                            If lNewQty = 0 Then
                                CheckedCell(m_fgTradeItems, Row, TradeItemCol(eGDCol_Active)) = False
                            Else
                                TradeItem.ChangeEntryQuantity lOldQty, lNewQty
                            End If
                        End If
                    End If
                End If
                
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_AfterMoveColumn
'' Description: After a user moves a column, resave the display string
'' Inputs:      Column moved, Position moved to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_AfterMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_AfterRowColChange
'' Description: If QtyNextEntry, go into edit mode
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Static bEditingCell As Boolean      ' Are we currently editing the cell?

    If (NewRow < m_fgTradeItems.Rows - 1) And (NewCol = TradeItemCol(eGDCol_QtyNextEntry)) Then
        If m_fgTradeItems.MergeRow(NewRow) = False Then
            If bEditingCell = False Then
                bEditingCell = True
                m_fgTradeItems.EditCell
                bEditingCell = False
            End If
        End If
    End If
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_AfterSort
'' Description: Set the background colors after a sort
'' Inputs:      Column of Sort, Order of Sort
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    HandleManageRow
    SetBackColors m_fgTradeItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_AfterUserResize
'' Description: Rebuild the display string after a user resize
'' Inputs:      Row and Column of resize
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_AfterUserResize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_BeforeEdit
'' Description: Only allow editing of Active and Qty Next Entry
'' Inputs:      Row and Column of Edited Cell, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If m.TCGrid.SelectRow(Row) Then
        If IsManageRow(Row) = True Then
            Cancel = True
        Else
            Select Case Col
                Case TradeItemCol(eGDCol_Active)
                    m_fgTradeItems.ComboList = ""
                
                Case TradeItemCol(eGDCol_OpenPos)
                    m_fgTradeItems.ComboList = "..."
                
                Case TradeItemCol(eGDCol_QtyNextEntry)
                    m_fgTradeItems.ComboList = ""
                
                Case TradeItemCol(eGDCol_OnCloseTime)
                    m_fgTradeItems.ComboList = "..."
                
                Case Else
                    Cancel = True
            End Select
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTradeItems_BeforeMouseDown
'' Description: If the user clicks the right button, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Location of Mouse, Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row the user clicked in
    Dim lMouseCol As Long               ' Column the user clicked in

    lMouseRow = m_fgTradeItems.MouseRow
    lMouseCol = m_fgTradeItems.MouseCol

    If Button = vbRightButton Then
        m.TCGrid.SelectRow lMouseRow
        EnableControls
        m.frm.PopupMenu m.mnuTradeItems
    
    ElseIf m.TCGrid.ValidRow(lMouseRow) Then
        If IsManageRow(lMouseRow) = True Then
            StartMenuTimer "NEWTRADEITEM"
        
        ElseIf lMouseCol = TradeItemCol(eGDCol_Active) Then
            If Shift = vbCtrlMask Then
                If mFlexGrid.CheckedCell(m_fgTradeItems, lMouseRow, TradeItemCol(eGDCol_Active)) = True Then
                    DisableAll "User Ctrl-Clicked in Grid"
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_BeforeMoveColumn
'' Description: Make sure that the frozen columns stay where they are
'' Inputs:      Column to move, Position to move it to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        With m_fgTradeItems
            If Col <= eGDCol_Active Then
                Position = Col
            ElseIf Position <= eGDCol_Active Then
                Position = eGDCol_Active + 1
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_BeforeMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_BeforeSort
'' Description: Handle the user sorting a column
'' Inputs:      Column, Sort Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SortOnCol Col, Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cWorkingOrdersUI.m_fgTradeItems_BeforeSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_CellButtonClick
'' Description: Allow the user to edit the on close time
'' Inputs:      Row and Column of Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim TradeItem As cAutoTradeItem     ' Temporary trading item object
    Dim dSessionEnd As Double           ' Session ending time for the symbol
    Dim dSessionStart As Double         ' Session starting time for the symbol
    Dim dEndTime As Double              ' User specified ending time
    Dim lMinutesBefore As Long          ' Minutes before closing time to check
    Dim Bars As New cGdBars             ' Temporary bars structure
    Dim dOnCloseTimeExch As Double      ' On-Close time in exchange time
    
    With m_fgTradeItems
        Set TradeItem = TradeItemForRow(Row)
        If Not TradeItem Is Nothing Then
            If Col = TradeItemCol(eGDCol_OnCloseTime) Then
                SetBarProperties Bars, TradeItem.Symbol
                dSessionEnd = ConvertTimeZone(Bars.Prop(eBARS_DefaultEndTime) / 1440#, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                dSessionStart = ConvertTimeZone(Bars.Prop(eBARS_DefaultStartTime) / 1440#, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                lMinutesBefore = TradeItem.MinutesBefore
                dEndTime = TradeItem.OnCloseTime + (lMinutesBefore / 1440#)
                
                If frmOnCloseTime.ShowMe(dEndTime, lMinutesBefore, dSessionEnd, dSessionStart) = True Then
                    'TradeItem.OnCloseTime = dEndTime - (lMinutesBefore / 1440#)
                    TradeItem.MinutesBefore = lMinutesBefore
                    
                    dOnCloseTimeExch = ConvertTimeZone(dEndTime - (lMinutesBefore / 1440#), "", Bars.Prop(eBARS_ExchangeTimeZoneInf))
                    dOnCloseTimeExch = dOnCloseTimeExch - Int(dOnCloseTimeExch)
                    TradeItem.OnCloseTimeExch = dOnCloseTimeExch
                    
                    TradeItem.Save
                    
                    RefreshTradeItem TradeItem.AutoTradeItemID
                End If
            ElseIf Col = TradeItemCol(eGDCol_OpenPos) Then
                StartMenuTimer "REASSIGNFILLS"
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_CellButtonClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_Compare
'' Description: Perform a comparison for the two rows for sorting purposes
'' Inputs:      Row 1, Row 2, Compare Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
On Error GoTo ErrSection:

    Dim strRow1 As String
    Dim strRow2 As String
    
    strRow1 = m_fgTradeItems.TextMatrix(Row1, TradeItemCol(eGDCol_SortKey))
    strRow2 = m_fgTradeItems.TextMatrix(Row2, TradeItemCol(eGDCol_SortKey))
    
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
    RaiseError "cTradeItemsUI.m_fgTradeItems_Compare"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_DblClick
'' Description: Allow the user to edit an item by double clicking on it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_DblClick()
On Error GoTo ErrSection:

    With m_fgTradeItems
        If m.TCGrid.SelectRow(.MouseRow) Then
            If .MouseCol = TradeItemCol(eGDCol_Symbol) Then
                ShowChartForTradingItem
            Else
                EditTradeItem
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_KeyUp
'' Description: Delete removes a trade item, Insert creates a new one
'' Inputs:      Code of the Key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            If m.TCGrid.ValidRow(m_fgTradeItems.Row) Then
                If m_fgTradeItems.RowOutlineLevel(m_fgTradeItems.Row) = 0 Then
                    StartMenuTimer "DELETETRADEITEM"
                End If
            End If
        
        Case vbKeyInsert
            StartMenuTimer "NEWTRADEITEM"
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgWorkingOrders_KeyUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_MouseMove
'' Description: Set the cursor appropriately
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim lMouseCol As Long               ' Mouse column in the grid
    Dim strNewText As String            ' New tooltip text

    lMouseRow = m_fgTradeItems.MouseRow
    lMouseCol = m_fgTradeItems.MouseCol
    m.TCGrid.SetMouseCursor m.frm
    
    strNewText = ""
    If m.TCGrid.ValidRow(lMouseRow) Then
        If Len(m_fgTradeItems.TextMatrix(lMouseRow, TradeItemCol(eGDCol_Information))) > 0 Then
            strNewText = m_fgTradeItems.TextMatrix(lMouseRow, TradeItemCol(eGDCol_Information))
        ElseIf Len(m_fgTradeItems.TextMatrix(lMouseRow, TradeItemCol(eGDCol_SymbolError))) > 0 Then
            If lMouseCol = TradeItemCol(eGDCol_Symbol) Then
                strNewText = Replace(Replace(m_fgTradeItems.TextMatrix(lMouseRow, TradeItemCol(eGDCol_SymbolError)), "||", " "), "|", " ")
            End If
        ElseIf Len(m_fgTradeItems.TextMatrix(lMouseRow, TradeItemCol(eGDCol_QuantityError))) > 0 Then
            If lMouseCol = TradeItemCol(eGDCol_QtyNextEntry) Then
                strNewText = Replace(Replace(m_fgTradeItems.TextMatrix(lMouseRow, TradeItemCol(eGDCol_QuantityError)), "||", " "), "|", " ")
            End If
        End If
    ElseIf Len(m_fgTradeItems.ToolTipText) > 0 Then
        strNewText = ""
    End If
    
    If strNewText <> m_fgTradeItems.ToolTipText Then
        m_fgTradeItems.ToolTipText = strNewText
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgTradeItems_ValidateEdit
'' Description: Make sure that realtime is running if going active
'' Inputs:      Row and Column of Edited Cell, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgTradeItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated trading item from the grid
    Dim strError As String              ' Error

    Select Case Col
        'Case TradeItemCol(eGDCol_Active)
            'If Not g.RealTime.IsServerActive(True) Then
                'Cancel = True
                'InfBox "You cannot activate an automated trading item unless streaming is active", "!", , "Automated Trading Error"
            'ElseIf g.RealTime.SymbolDelay(fgTradeItems.TextMatrix(Row, TradeItemCol(eGDCol_Symbol))) <> 0 Then
                'Cancel = True
                'InfBox "You cannot run an automated trading item on a delayed symbol", "!", , "Automated Trading Error"
            'End If
            
        Case TradeItemCol(eGDCol_QtyNextEntry)
            If m_fgTradeItems.MergeRow(Row) = False Then
                If TypeOf m_fgTradeItems.RowData(Row) Is cAutoTradeItem Then
                    Set TradeItem = m_fgTradeItems.RowData(Row)
                    strError = AutomatedQuantityError(TradeItem, m_fgTradeItems.EditText, m_fgTradeItems.TextMatrix(Row, TradeItemCol(eGDCol_SymbolError)))
                    If Len(strError) > 0 Then
                        InfBox strError, "!", , "Error"
                        Cancel = True
                    ElseIf (TradeItem.Pyramid = True) And (TradeItem.CurrentPosition <> 0) Then
                        InfBox "You cannot change the quantity of the next entry while you are in a position in a pyramiding system", "!", , "Error"
                        Cancel = True
                    End If
                End If
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_fgTradeItems_ValidateEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuActualPerformance_Click
'' Description: Allow the user to view the performance reports for the trades
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
    RaiseError "cTradeItemUI.m_mnuActualPerformance_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuAlerts_Click
'' Description: Allow the user to manage alerts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuAlerts_Click()
On Error GoTo ErrSection:
    
    StartMenuTimer "ALERTS"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuAlerts_Click"
    
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
    RaiseError "cTradeItemUI.m_mnuAutoSizeColumns_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuChangePosition_Click
'' Description: Allow the user to change the current auto trade position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuChangePosition_Click()
On Error GoTo ErrSection:

    StartMenuTimer "REASSIGNFILLS"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuChangePosition_Click"
    
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

    m.TCGrid.Display = DefaultDisplay
    ApplyDisplayString
    BuildDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuDefaultColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuDeleteTradeItem_Click
'' Description: Allow the user to remove an automated trading item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuDeleteTradeItem_Click()
On Error GoTo ErrSection:

    StartMenuTimer "DELETETRADEITEM"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuDeleteTradeItem_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuDisableAll_Click
'' Description: Allow the user to disable all active automated trading items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuDisableAll_Click()
On Error GoTo ErrSection:

    StartMenuTimer "DISABLEALL"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuDisableAll_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuEditStrategy_Click
'' Description: Allow the user to edit a strategy from a trading item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuEditStrategy_Click()
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated Trading Item to edit
    
    Set TradeItem = SelectedTradeItem
    If Not TradeItem Is Nothing Then
        If TradeItem.ParentID >= 0& Then
            EditStrategy TradeItem
        Else
            If TradeItem.IsGuruOwner Then
                EditBasket TradeItem
            Else
                InfBox "You are not authorized to edit this strategy", "!", , "Error"
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuEditStrategy_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuEditTradeItem_Click
'' Description: Allow the user to edit an automated trading item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuEditTradeItem_Click()
On Error GoTo ErrSection:

    StartMenuTimer "EDITTRADEITEM"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuEditTradeItem_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuEnterPosition_Click
'' Description: Allow the user to enter into the current auto trade position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuEnterPosition_Click()
On Error GoTo ErrSection:

    StartMenuTimer "ENTERPOSITION"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuEnterPosition_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuFlatten_Click
'' Description: Allow the user to close the current auto trade position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuFlatten_Click()
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Auto trade item object
    
    Set TradeItem = SelectedTradeItem
    If Not TradeItem Is Nothing Then
        If TradeItem.Active Then
            TradeItem.DumpDebug "User chose to Flatten from the Trading Items grid"
            FlattenForSymbol TradeItem.AccountID, TradeItem.SymbolOrSymbolID, TradeItem.AutoTradeItemID
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuFlatten_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuNewTradeItem_Click
'' Description: Allow the user to create a new automated trading item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuNewTradeItem_Click()
On Error GoTo ErrSection:

    StartMenuTimer "NEWTRADEITEM"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuNewTradeItem_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuNextBarReport_Click
'' Description: Allow the user to view the next bar report for a strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuNextBarReport_Click()
On Error GoTo ErrSection:

    StartMenuTimer "NEXTBARREPORT"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuNextBarReport_Click"
    
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
    RaiseError "cTradeItemUI.m_mnuPrint_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuStrategyPerformance_Click
'' Description: Allow the user to view the performance reports for a strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuStrategyPerformance_Click()
On Error GoTo ErrSection:

    StartMenuTimer "STRATEGYPERFORMANCE"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuStrategyPerformance_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuRollContracts_Click
'' Description: Allow the user to roll into another contract
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuRollContracts_Click()
On Error GoTo ErrSection:

    StartMenuTimer "ROLLCONTRACTS"

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTradeItemUI.m_mnuRollContracts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuShowChart_Click
'' Description: Allow the user to show a chart with the strategy/symbol/period
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuShowChart_Click()
On Error GoTo ErrSection:

    ShowChartForTradingItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_mnuShowChart_Click"
    
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
    RaiseError "cTradeItemUI.m_mnuTradeHistory_Click"
    
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
    Dim TradeItem As cAutoTradeItem     ' Trading item from the grid

    strTag = m_tmrMenu.Tag
    If Len(strTag) > 2 Then
        If Left(strTag, 2) = kMenuPrefix Then
            strTag = Mid(strTag, 3)
            m_tmrMenu.Tag = ""
            m_tmrMenu.Enabled = False
            
            Set TradeItem = SelectedTradeItem
            
            Select Case UCase(strTag)
                Case "ACTUALPERFORMANCE"
                    If Not TradeItem Is Nothing Then
                        frmTradeReportFilter.ShowForTradeItem TradeItem, True
                    End If
                
                Case "ALERTS"
                    frmAlertsSetup.ShowMe
                    
                Case "DELETETRADEITEM"
                    RemoveTradeItem
                    
                Case "DISABLEALL"
                    DisableAll "User clicked on right-click option"
                    
                Case "EDITTRADEITEM"
                    EditTradeItem
                    
                Case "ENTERPOSITION"
                    If Not TradeItem Is Nothing Then
                        If TradeItem.Active Then
                            TradeItem.EnterPosition
                        End If
                    End If
                    
                Case "NEWTRADEITEM"
                    NewTradeItem
                    
                Case "NEXTBARREPORT"
                    If Not TradeItem Is Nothing Then
                        If TradeItem.IsGuruOwner Then
                            TradeItem.NextBarReport
                        End If
                    End If
                    
                Case "REASSIGNFILLS"
                    ReassignFills TradeItem
                    
                Case "STRATEGYPERFORMANCE"
                    If Not TradeItem Is Nothing Then
                        TradeItem.PerformanceReport
                    End If
                    
                Case "ROLLCONTRACTS"
                    RollContracts
                    
                Case "TRADEHISTORY"
                    If TradeItem Is Nothing Then
                        frmTTAccounts.ShowMe True
                    Else
                        frmTTPositions.ShowMe TradeItem.AccountID, g.Broker.AccountTypeForID(TradeItem.AccountID)
                    End If
                    
            End Select
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.m_tmrMenu_Timer"
    
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
    RaiseError "cTradeItemUI.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleManageRow
'' Description: Make sure that the "Manage Row" is the last row in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleManageRow()
On Error GoTo ErrSection:

    Dim lManageRow As Long              ' Row in the grid containing the manage row
    Dim lIndex As Long                  ' Index into a for loop
    
    With m_fgTradeItems
        lManageRow = -1&
        For lIndex = .FixedRows To .Rows - 1
            If IsManageRow(lIndex) Then
                lManageRow = lIndex
                Exit For
            End If
        Next lIndex
        
        If lManageRow = -1& Then
            .Rows = .Rows + 1
            .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Active)) = flexNoCheckbox
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = kManageRowText
            .Cell(flexcpFontUnderline, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            If g.nColorTheme = kDarkThemeColor Then
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbCyan
            Else
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
            End If
            .MergeRow(.Rows - 1) = True
        ElseIf lManageRow <> .Rows - 1 Then
            .RowPosition(lIndex) = .Rows - 1
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.HandleManageRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeItemToGrid
'' Description: Update the grid with the given automated trading item
'' Inputs:      Trade Item, Row
'' Returns:     Row
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TradeItemToGrid(TradeItem As cAutoTradeItem, Optional ByVal lRow As Long = -1&, Optional ByVal Baskets As cStrategyBaskets = Nothing) As Long
On Error GoTo ErrSection:
    
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim strStrategyMonth As String      ' Month of the contract the strategy is in
    Dim strPositionMonth As String      ' Month of the contract the position is in
    Dim strStrategyPos As String        ' Strategy position string
    Dim strActualPos As String          ' Actual position string
    Dim bMismatch As Boolean            ' Position mismatch
    Dim lChildIndex As Long             ' Index of the child for the parent
    Dim Basket As cStrategyBasket       ' Strategy basket
    Dim strSymbolError As String        ' Symbol error
    Dim strQtyError As String           ' Quantity error
    Dim lMaxUnits As Long               ' Maximum units for automated trading item
    Dim dBasketLastModified As Double   ' Last modified date for the basket
    Dim lParentRow As Long              ' Row for the parent trading item

    With m_fgTradeItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
            
            ' DAJ 01/14/2016: Tim ran into an issue where he added a symbol to a symbol group that was in
            ' a strategy basket that was in an automated trading item.  When he saved everything up the
            ' chain, the new automated trading item got sorted in the wrong part of the grid.  We need to
            ' make sure here that any child starts out as a child row of the parent row...
            If TradeItem.ParentID > 0& Then
                lParentRow = RowForTradeItemId(TradeItem.ParentID)
                If lParentRow <> -1& Then
                    .RowPosition(lRow) = lParentRow + 1&
                    lRow = lParentRow + 1&
                End If
            End If
        End If
        
        If TradeItem.ParentID <> -1& Then
            If TradeItem.WaitingForData Then
                strStrategyPos = "Pending Data"
                strActualPos = ""
                bMismatch = True
            Else
                If TradeItem.Active Then
                    strStrategyPos = TradeItem.StrategyPosString
                Else
                    strStrategyPos = ""
                End If
                strActualPos = PositionToString(TradeItem.CurrentPosition, True)
                If (Len(strStrategyPos) > 0) And (Len(strActualPos) > 0) Then
                    bMismatch = (UCase(Parse(strStrategyPos, " ", 1)) <> UCase(Parse(strActualPos, " ", 1)))
                End If
                
                If SecurityType(TradeItem.SymbolOrSymbolID, True) = "F" Then
                    If (Len(strStrategyPos) > 0) And (UCase(strStrategyPos) <> "FLAT") Then
                        strStrategyMonth = GetSymbol(TradeItem.TradeSymbolOrID)
                        If Len(strStrategyMonth) > 0 Then
                            strStrategyMonth = MonthName(CLng(Val(Right(strStrategyMonth, 2))), True)
                            strStrategyPos = strStrategyPos & " (" & strStrategyMonth & ")"
                        End If
                    End If
                    If (Len(strActualPos) > 0) And (UCase(strActualPos) <> "FLAT") Then
                        strPositionMonth = GetSymbol(TradeItem.PositionSymbolOrID)
                        If Len(strPositionMonth) > 0 Then
                            strPositionMonth = MonthName(CLng(Val(Right(strPositionMonth, 2))), True)
                            strActualPos = strActualPos & " (" & strPositionMonth & ")"
                        End If
                    End If
                    
                    If (Len(strStrategyMonth) > 0) And (Len(strPositionMonth) > 0) Then
                        If strStrategyMonth <> strPositionMonth Then
                            bMismatch = True
                        End If
                    End If
                End If
            End If
        End If
        
        Set Basket = New cStrategyBasket
        If TradeItem.StrategyBasketID > 0& Then
            If Baskets Is Nothing Then
                Basket.LoadDb TradeItem.StrategyBasketID, , True
            Else
                Set Basket = Baskets("ID=" & Str(TradeItem.StrategyBasketID))
            End If
        End If
        
        ' 05/12/2015 DAJ: Customer had a trial of the FractZen that expired.  With that trial expiration,
        ' the library was set to be ignored.  Because of that, the basket is not in the baskets collection
        ' and, therefore, is nothing here.  Set the basket last modified to the one that we had stored in
        ' the automated trading item so that it is not colored orange.
        If Basket Is Nothing Then
            dBasketLastModified = TradeItem.StrategyBasketLastModified
        Else
            dBasketLastModified = Basket.LastModified
        End If
        
        .RowData(lRow) = TradeItem
    
        CheckedCell(m_fgTradeItems, lRow, TradeItemCol(eGDCol_Active)) = TradeItem.Active
        .Cell(flexcpPictureAlignment, lRow, TradeItemCol(eGDCol_Active)) = flexPicAlignCenterCenter
        
        If TradeItem.ParentID = -1& Then
            .Cell(flexcpText, lRow, TradeItemCol(eGDCol_Name), lRow, .Cols - 1) = TradeItem.Name
            .RowOutlineLevel(lRow) = 0&
            
            .TextMatrix(lRow, TradeItemCol(eGDCol_ItemBasketLastModified)) = Str(TradeItem.StrategyBasketLastModified)
            .TextMatrix(lRow, TradeItemCol(eGDCol_BasketLastModified)) = Str(dBasketLastModified)
        
            .TextMatrix(lRow, TradeItemCol(eGDCol_SortKey)) = ""
            .TextMatrix(lRow, TradeItemCol(eGDCol_AscSortKey)) = "T" & Str(TradeItem.AutoTradeItemID) & Chr(31)
            .TextMatrix(lRow, TradeItemCol(eGDCol_DescSortKey)) = "T" & Str(TradeItem.AutoTradeItemID) & "}"
            .TextMatrix(lRow, TradeItemCol(eGDCol_OutlineLevel)) = "0"
            
            .MergeRow(lRow) = True
        Else
            .TextMatrix(lRow, TradeItemCol(eGDCol_Name)) = TradeItem.Name
            .TextMatrix(lRow, TradeItemCol(eGDCol_Strategy)) = TradeItem.StrategyName
            .TextMatrix(lRow, TradeItemCol(eGDCol_Symbol)) = TradeItem.Symbol
            If (UCase(TradeItem.BarPeriod) = "AUTO BREAKOUT") Or (UCase(TradeItem.BarPeriod) = "FRACTZEN") Then
                If InStr(TradeItem.ActualBarPeriod, " ") <> 0 Then
                    .TextMatrix(lRow, TradeItemCol(eGDCol_Period)) = TradeItem.BarPeriod & " (" & Parse(TradeItem.ActualBarPeriod, " ", 1) & ")"
                Else
                    .TextMatrix(lRow, TradeItemCol(eGDCol_Period)) = TradeItem.BarPeriod
                End If
            Else
                .TextMatrix(lRow, TradeItemCol(eGDCol_Period)) = TradeItem.BarPeriod
            End If
            .TextMatrix(lRow, TradeItemCol(eGDCol_Account)) = g.Broker.AccountNameForID(TradeItem.AccountID)
            .TextMatrix(lRow, TradeItemCol(eGDCol_StrategyPos)) = strStrategyPos
            .TextMatrix(lRow, TradeItemCol(eGDCol_OpenPos)) = strActualPos
            .TextMatrix(lRow, TradeItemCol(eGDCol_QtyNextEntry)) = Str(TradeItem.QtyNextEntry)
            .TextMatrix(lRow, TradeItemCol(eGDCol_OpenOrders)) = Str(g.Broker.HasWorkingOrdersForAutoTrade(TradeItem))
            .TextMatrix(lRow, TradeItemCol(eGDCol_OnCloseTime)) = Format(TradeItem.OnCloseTime, "hh:mm")
            
            If TradeItem.ParentID = 0& Then
                .RowOutlineLevel(lRow) = 0&
            
                .TextMatrix(lRow, TradeItemCol(eGDCol_ItemBasketLastModified)) = ""
                .TextMatrix(lRow, TradeItemCol(eGDCol_BasketLastModified)) = ""
        
                .TextMatrix(lRow, TradeItemCol(eGDCol_SortKey)) = ""
                .TextMatrix(lRow, TradeItemCol(eGDCol_AscSortKey)) = "T" & Str(TradeItem.AutoTradeItemID) & Chr(31)
                .TextMatrix(lRow, TradeItemCol(eGDCol_DescSortKey)) = "T" & Str(TradeItem.AutoTradeItemID) & "}"
                .TextMatrix(lRow, TradeItemCol(eGDCol_OutlineLevel)) = "0"
            Else
                .RowOutlineLevel(lRow) = 1&
            
                lChildIndex = CLng(Val(Parse(TradeItem.Name, "#", 2)))
            
                .TextMatrix(lRow, TradeItemCol(eGDCol_ItemBasketLastModified)) = Str(TradeItem.StrategyBasketLastModified)
                .TextMatrix(lRow, TradeItemCol(eGDCol_BasketLastModified)) = Str(dBasketLastModified)
        
                .TextMatrix(lRow, TradeItemCol(eGDCol_SortKey)) = ""
                .TextMatrix(lRow, TradeItemCol(eGDCol_AscSortKey)) = "T" & Str(TradeItem.ParentID) & "_" ' & Format(lChildIndex, "000")
                .TextMatrix(lRow, TradeItemCol(eGDCol_DescSortKey)) = "T" & Str(TradeItem.ParentID) & "_" ' & Format(lChildIndex, "000")
                .TextMatrix(lRow, TradeItemCol(eGDCol_OutlineLevel)) = "1"
            End If
            
            lMaxUnits = mSysNav.MaxUnitsForAutoTrade(TradeItem)
            If lMaxUnits = Abs(kNullData) Then
                .TextMatrix(lRow, TradeItemCol(eGDCol_MaxQty)) = ""
            Else
                .TextMatrix(lRow, TradeItemCol(eGDCol_MaxQty)) = Str(lMaxUnits * TradeItem.StrategyBasketItemMult)
            End If
            
            If bMismatch Then
                .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_StrategyPos), lRow, TradeItemCol(eGDCol_OpenPos)) = vbRed
            Else
                .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_StrategyPos), lRow, TradeItemCol(eGDCol_OpenPos)) = vbBlack
            End If
            
            If TradeItem.DataBehind Then
                .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_Name)) = vbRed
                .TextMatrix(lRow, TradeItemCol(eGDCol_Information)) = "Data is behind"
            Else
                .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_Name)) = .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_Strategy))
                .TextMatrix(lRow, TradeItemCol(eGDCol_Information)) = ""
            End If
            
            strSymbolError = AutomatedSymbolError(TradeItem.AccountID, TradeItem.SymbolOrSymbolID, "Automated Trading", True)
            If Len(strSymbolError) > 0 Then
                .TextMatrix(lRow, TradeItemCol(eGDCol_SymbolError)) = strSymbolError
                .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_Symbol)) = vbRed
            Else
                .TextMatrix(lRow, TradeItemCol(eGDCol_SymbolError)) = ""
                .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_Symbol)) = .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_Strategy))
            End If
            
            strQtyError = AutomatedQuantityError(TradeItem, .TextMatrix(lRow, TradeItemCol(eGDCol_QtyNextEntry)), strSymbolError)
            If Len(strQtyError) > 0 Then
                .TextMatrix(lRow, TradeItemCol(eGDCol_QuantityError)) = strQtyError
                .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_QtyNextEntry)) = vbRed
            Else
                .TextMatrix(lRow, TradeItemCol(eGDCol_QuantityError)) = ""
                .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_QtyNextEntry)) = .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_Strategy))
            End If
                        
            .MergeRow(lRow) = False
        End If
        
        CheckLastModified lRow
        
        .IsSubtotal(lRow) = True
        
        .Redraw = nRedraw
    End With
    
    TradeItemToGrid = lRow

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.TradeItemToGrid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReassignFills
'' Description: Reassign the fills between manual and automated trading items
'' Inputs:      Trade Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReassignFills(TradeItem As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim bChanged As Boolean             ' Did some fills get reassigned?

    If Not TradeItem Is Nothing Then
        If frmReassignFills.ShowMe(TradeItem, bChanged) Then
            If bChanged Then
                'If TradeItem.Active Then
                    TradeItem.RefreshPosition
                'Else
                '    TradeItem.CurrentPosition = g.Broker.CurrentPosition(TradeItem.AccountID, TradeItem.PositionSymbolOrID, TradeItem.AutoTradeItemID)
                '    TradeItem.Save
                '
                '    RefreshTradeItem TradeItem.AutoTradeItemID
                'End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.ReassignFills"
    
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
    Dim strValue2 As String             ' Value for the child

    If lCol = kNullData Then
        If m.lSortedCol = -1& Then
            lCol = TradeItemCol(eGDCol_Name)
        Else
            lCol = m.lSortedCol
        End If
    End If
    
    strFormat = ""
    Select Case m_fgTradeItems.ColData(lCol)
        Case GDCol(eGDCol_Active)
            strFormat = "Bool"
        Case GDCol(eGDCol_QtyNextEntry), GDCol(eGDCol_OpenOrders)
            strFormat = "Int"
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

    With m_fgTradeItems
        If .Rows > .FixedRows Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            strValue = ""
            strValue2 = ""
            For lIndex = .FixedRows To .Rows - 1
                If .RowOutlineLevel(lIndex) = 0 Then
                    Select Case UCase(strFormat)
                        Case "BOOL"
                            strValue = Pad(Format(Val(CheckedCell(m_fgTradeItems, lIndex, lCol)), "0000000"), 50, "R")
                        Case "INT"
                            strValue = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "0000000"), 50, "R")
                        Case "NUMBER", "DATE"
                            strValue = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.0000000"), 50, "R")
                        Case "CURRENCY"
                            strValue = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.00"), 50, "R")
                        Case Else
                            strValue = Pad(.TextMatrix(lIndex, lCol), 50, "L")
                    End Select
                Else
                    Select Case UCase(strFormat)
                        Case "BOOL"
                            strValue2 = Pad(Format(Val(CheckedCell(m_fgTradeItems, lIndex, lCol)), "0000000"), 50, "R")
                        Case "INT"
                            strValue2 = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "0000000"), 50, "R")
                        Case "NUMBER", "DATE"
                            strValue2 = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.0000000"), 50, "R")
                        Case "CURRENCY"
                            strValue2 = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.00"), 50, "R")
                        Case Else
                            strValue2 = Pad(.TextMatrix(lIndex, lCol), 50, "L")
                    End Select
                End If
                
                If bAscending Then
                    .TextMatrix(lIndex, TradeItemCol(eGDCol_SortKey)) = UCase(strValue & "_" & .TextMatrix(lIndex, TradeItemCol(eGDCol_AscSortKey)) & strValue2)
                Else
                    .TextMatrix(lIndex, TradeItemCol(eGDCol_SortKey)) = UCase(strValue & "_" & .TextMatrix(lIndex, TradeItemCol(eGDCol_DescSortKey)) & strValue2)
                End If
                
                .RowOutlineLevel(lIndex) = 0
                .IsSubtotal(lIndex) = False
            Next lIndex
            
            .Select .FixedRows, TradeItemCol(eGDCol_SortKey), .Rows - 1, TradeItemCol(eGDCol_SortKey)
            If bAscending Then
                m.nSortedDir = flexSortStringAscending
            Else
                m.nSortedDir = flexSortStringDescending
            End If
            .Sort = flexSortCustom
            .Select .FixedRows, 0
            
            For lIndex = .FixedRows To .Rows - 1
                .RowOutlineLevel(lIndex) = CLng(Val(.TextMatrix(lIndex, TradeItemCol(eGDCol_OutlineLevel))))
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
    RaiseError "cTradeItemUI.SortOnCol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsManageRow
'' Description: Is the given row the manage row?
'' Inputs:      Row
'' Returns:     True if manage row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsManageRow(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean

    bReturn = False
    With m_fgTradeItems
        If m.TCGrid.ValidRow(lRow) Then
            bReturn = (.MergeRow(lRow) = True) And (.TextMatrix(lRow, 0) = kManageRowText)
        End If
    End With
    
    IsManageRow = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.IsManageRow"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChildInPosition
'' Description: Determine if any children of the given row are in a position
'' Inputs:      Row
'' Returns:     True if children in a position, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChildInPosition(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lChildRow As Long               ' Child row for the given row
    Dim TradeItem As cAutoTradeItem     ' Automated trading item for the row
    
    bReturn = False
    With m_fgTradeItems
        lChildRow = .GetNodeRow(lRow, flexNTFirstChild)
        Do While lChildRow <> -1&
            Set TradeItem = .RowData(lChildRow)
            
            If TradeItem.CurrentPosition <> 0& Then
                bReturn = True
                Exit Do
            End If
            
            lChildRow = .GetNodeRow(lChildRow, flexNTNextSibling)
        Loop
    End With
    
    ChildInPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.ChildInPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditStrategy
'' Description: Edit the strategy for an automated trading item
'' Inputs:      Trade Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditStrategy(ByVal TradeItem As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim SystemMgr As frmSystemManager   ' New System Manager form
    
    If Not TradeItem Is Nothing Then
        If TradeItem.StrategyID > 0 Then
            Set SystemMgr = New frmSystemManager
            SystemMgr.ShowMe TradeItem.StrategyID, , False
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.EditStrategy"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditBasket
'' Description: Edit the strategy basket for an automated trading item
'' Inputs:      Trade Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditBasket(ByVal TradeItem As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim BasketMgr As frmStrategyBasket  ' New strategy basket form
    
    If Not TradeItem Is Nothing Then
        If TradeItem.StrategyBasketID > 0 Then
            Set BasketMgr = New frmStrategyBasket
            frmStrategyBasket.ShowMe TradeItem.StrategyBasketID
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.EditBasket"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateTradeItem
'' Description: Update the given automated trading item in the grid
'' Inputs:      Trade Item, Original, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateTradeItem(TradeItem As cAutoTradeItem, ByVal Original As cAutoTradeItem, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim lAcctPosID As Long              ' Account Position ID

    If lRow = -1& Then
        lRow = RowForTradeItem(TradeItem)
    End If

    If TradeItem.ParentID >= 0& Then
        lAcctPosID = g.Broker.CreateFillSummaryForAutoTrade(TradeItem)
        If lAcctPosID <> TradeItem.AccountPositionID Then
            TradeItem.AccountPositionID = lAcctPosID
            TradeItem.Save
            g.TradingItems.Add TradeItem
        End If
    End If
    
    lRow = TradeItemToGrid(TradeItem, lRow)
    
    If TradeItem.ParentID >= 0& Then
        If Original.QtyNextEntry <> TradeItem.QtyNextEntry Then
            TradeItem.DumpDebug "User changed Quantity Next Entry from Dialog.  Old = " & Str(Original.QtyNextEntry) & "; New = " & Str(TradeItem.QtyNextEntry)
            If CheckedCell(m_fgTradeItems, lRow, TradeItemCol(eGDCol_Active)) Then
                TradeItem.ChangeEntryQuantity Original.QtyNextEntry, TradeItem.QtyNextEntry
            End If
        End If
    End If
    
    If TradeItem.Name <> Original.Name Then
        'm.SummaryOrders.ChangeTradeItemName Original.Name, TradeItem.Name
        'm.Orders.ChangeTradeItemName Original.Name, TradeItem.Name
        
        'm.SummaryPositions.ChangeTradeItemName Original.Name, TradeItem.Name
        'm.Positions.ChangeTradeItemName Original.Name, TradeItem.Name
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTradeItemUI.UpdateTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowForTradeItem
'' Description: Determine the row for the given trade item
'' Inputs:      Trade Item
'' Returns:     Row ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowForTradeItem(ByVal TradeItem As cAutoTradeItem) As Long
On Error GoTo ErrSection:

    RowForTradeItem = RowForTradeItemId(TradeItem.AutoTradeItemID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.RowForTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowForTradeItemId
'' Description: Determine the row for the given trade item ID
'' Inputs:      Trade Item ID
'' Returns:     Row ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowForTradeItemId(ByVal lAutoTradeItemID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim RowItem As cAutoTradeItem       ' Trade item from the row
    
    lReturn = -1&
    With m_fgTradeItems
        For lIndex = .FixedRows To .Rows - 1
            If IsManageRow(lIndex) = False Then
                Set RowItem = .RowData(lIndex)
                If RowItem.AutoTradeItemID = lAutoTradeItemID Then
                    lReturn = lIndex
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    RowForTradeItemId = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.RowForTradeItemId"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveTradeItemForRow
'' Description: Remove the trade item for the given row in the grid
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveTradeItemForRow(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings in the grid
    
    With m_fgTradeItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If m.TCGrid.ValidRow(lRow) Then
            g.TradingItems.Delete .RowData(lRow).AutoTradeItemID
            .RemoveItem lRow
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.RemoveTradeItemForRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveExpandedSettings
'' Description: Save the Expanded/Collapsed settings for the parent rows in grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveExpandedSettings()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid
    Dim lChildRow As Long               ' Row of the first child
    Dim astrSettings As cGdArray        ' Array of settings
    Dim TradeItem As cAutoTradeItem     ' Automated trading item for the row
    
    Set astrSettings = New cGdArray
    astrSettings.Create eGDARRAY_Strings
    
    With m_fgTradeItems
        If .Rows > .FixedRows Then
            lRow = .FixedRows
            Do While lRow <> -1&
                lChildRow = .GetNodeRow(lRow, flexNTFirstChild)
                If lChildRow <> -1& Then
                    Set TradeItem = TradeItemForRow(lRow)
                    If Not TradeItem Is Nothing Then
                        astrSettings.Add Str(TradeItem.AutoTradeItemID) & ":" & Str(.IsCollapsed(lRow))
                    End If
                End If
                
                lRow = .GetNodeRow(lRow, flexNTNextSibling)
            Loop
        End If
    End With
    
    astrSettings.Sort
    
    SetIniFileProperty "TradeItemsExpanded", astrSettings.JoinFields("|"), "TTSummary", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.SaveExpandedSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadExpandedSettings
'' Description: Load the Expanded/Collapsed settings for the parent rows in grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadExpandedSettings()
On Error GoTo ErrSection:

    Dim strSettings As String           ' Settings from the INI file
    Dim astrSettings As cGdArray        ' Array of settings
    Dim lRow As Long                    ' Row in the grid
    Dim lChildRow As Long               ' Row of the first child
    Dim TradeItem As cAutoTradeItem     ' Automated trading item for the row
    Dim lPos As Long                    ' Position in the array
    
    m.bLoadingExpanded = True
    
    strSettings = GetIniFileProperty("TradeItemsExpanded", "", "TTSummary", g.strIniFile)
    If Len(strSettings) > 0 Then
        Set astrSettings = New cGdArray
        astrSettings.SplitFields strSettings, "|"
        
        With m_fgTradeItems
            If .Rows > .FixedRows Then
                lRow = .FixedRows
                Do While lRow <> -1&
                    lChildRow = .GetNodeRow(lRow, flexNTFirstChild)
                    If lChildRow <> -1& Then
                        Set TradeItem = TradeItemForRow(lRow)
                        If Not TradeItem Is Nothing Then
                            If astrSettings.BinarySearch(Str(TradeItem.AutoTradeItemID) & ":", lPos, eGdSort_MatchUsingSearchStringLength) Then
                                .IsCollapsed(lRow) = CLng(Val(Parse(astrSettings(lPos), ":", 2)))
                            Else
                                .IsCollapsed(lRow) = flexOutlineExpanded
                            End If
                        End If
                    End If
                    
                    lRow = .GetNodeRow(lRow, flexNTNextSibling)
                Loop
            End If
        End With
    End If
    
    m.bLoadingExpanded = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.LoadExpandedSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckLastModified
'' Description: Check to see if the basket has been modified
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckLastModified(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With m_fgTradeItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If .TextMatrix(lRow, TradeItemCol(eGDCol_ItemBasketLastModified)) <> .TextMatrix(lRow, TradeItemCol(eGDCol_BasketLastModified)) Then
            .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_Name)) = RGB(255, 153, 0)
            .TextMatrix(lRow, TradeItemCol(eGDCol_Information)) = "The strategy basket that this automated trading item uses has changed.  Edit the automated trading item to resync."
        Else
            .Cell(flexcpForeColor, lRow, TradeItemCol(eGDCol_Name)) = vbBlack
            .TextMatrix(lRow, TradeItemCol(eGDCol_Information)) = ""
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.CheckLastModified"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridToString
'' Description: Convert the data in the grid to a string
'' Inputs:      None
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GridToString() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lRow As Long                    ' Index into a for loop
    Dim astrRows As cGdArray            ' Array of strings for rows in grid
    Dim strRow As String                ' Row in the grid converted to a string
    
    With m_fgTradeItems
        If (.Rows = .FixedRows) Or ((.Rows = .FixedRows + 1) And (IsManageRow(.FixedRows) = True)) Then
            strReturn = "No rows in grid"
        Else
            Set astrRows = New cGdArray
            astrRows.Create eGDARRAY_Strings
            
            For lRow = .FixedRows To .Rows - 1
                strRow = GridRowToString(lRow)
                If Len(strRow) > 0 Then
                    astrRows.Add strRow
                End If
            Next lRow
            
            strReturn = astrRows.JoinFields(vbCrLf)
        End If
    End With
    
    GridToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.GridToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridRowToString
'' Description: Convert the data in the given row to a string
'' Inputs:      Row
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GridRowToString(ByVal lRow As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrFields As cGdArray          ' Fields from the grid
    Dim lCol As Long                    ' Index into a for loop
    
    strReturn = ""
    If IsManageRow(lRow) = False Then
        With m_fgTradeItems
            Set astrFields = New cGdArray
            astrFields.Create eGDARRAY_Strings
            
            If CheckedCell(m_fgTradeItems, lRow, TradeItemCol(eGDCol_Active)) = True Then
                astrFields(0) = "On"
            Else
                astrFields(0) = "Off"
            End If
            
            If .MergeRow(lRow) = True Then
                astrFields(1) = .TextMatrix(lRow, TradeItemCol(eGDCol_Name))
            Else
                For lCol = TradeItemCol(eGDCol_Active) + 1 To .Cols - 1
                    If .ColHidden(lCol) = False Then
                        astrFields.Add .TextMatrix(lRow, lCol)
                    End If
                Next lCol
            End If
            
            strReturn = astrFields.JoinFields(vbTab)
        End With
    End If
    
    GridRowToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeItemUI.GridRowToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleTradeItem
'' Description: Toggle the trade item for the given row on/off
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ToggleTradeItem(ByVal Row As Long)
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Temporary automated trading object
    Dim bContinue As Boolean            ' Continue?
    Dim strSymbolError As String        ' Symbol error
    Dim strQtyError As String           ' Quantity error

    If m.TCGrid.ValidRow(Row) = True Then
        m.TCGrid.SelectRow Row
        
        'mFlexGrid.CheckedCell(m_fgTradeItems, Row, TradeItemCol(eGDCol_Active)) = Not mFlexGrid.CheckedCell(m_fgTradeItems, Row, TradeItemCol(eGDCol_Active))

        Set TradeItem = SelectedTradeItem
        If Not TradeItem Is Nothing Then
            strSymbolError = m_fgTradeItems.TextMatrix(Row, TradeItemCol(eGDCol_SymbolError))
            strQtyError = m_fgTradeItems.TextMatrix(Row, TradeItemCol(eGDCol_QuantityError))

            If mFlexGrid.CheckedCell(m_fgTradeItems, Row, TradeItemCol(eGDCol_Active)) = True Then
                If HasModule("RTG,RTE") = False Then
                    bContinue = False
                    InfBox "You cannot activate an automated trading item because you are not enabled for streaming", "!", , "Automated Trading Error"
                
                ElseIf g.Broker.ConnectionStatusForAccount(TradeItem.AccountID) <> eGDConnectionStatus_Connected Then
                    bContinue = False
                    InfBox "You cannot activate the automated trading item because you are not connected to the broker", "!", , "Automated Trading Error"
                
                ElseIf Len(TradeItem.Symbol) = 0 Then
                    If mSysNav.MaxUnitsForAutoTrade(TradeItem) = 0 Then
                        bContinue = False
                        InfBox "You are not enabled to Shadow Trade this automated trading item", "!", , "Automated Trading Error"
                    Else
                        bContinue = True
                    End If
                    
                ElseIf Len(strSymbolError) > 0 Then
                    bContinue = False
                    InfBox strSymbolError, "!", , "Automated Trading Error"
                    
                ElseIf Len(strQtyError) > 0 Then
                    bContinue = False
                    InfBox strQtyError, "!", , "Automated Trading Error"
                    
                Else
                    bContinue = mTradeTracker.CanActivateAutomatedItem(TradeItem.AccountID, TradeItem.SymbolOrSymbolID, "Automated Trading Item", "Trading Items Grid")
                End If
                
                If bContinue = True Then
                    If frmPositionConfirm.ShowMeForTradeItem(TradeItem) = True Then
                        RefreshTradeItem TradeItem.AutoTradeItemID
                    Else
                        CheckedCell(m_fgTradeItems, Row, TradeItemCol(eGDCol_Active)) = False
                    End If
                Else
                    CheckedCell(m_fgTradeItems, Row, TradeItemCol(eGDCol_Active)) = False
                End If
            Else
                If g.TradingItems.Disable(TradeItem.AutoTradeItemID, , "Check box unchecked in trading items form") = False Then
                    CheckedCell(m_fgTradeItems, Row, TradeItemCol(eGDCol_Active)) = True
                Else
                    If TradeItem.ParentID > -1& Then
                        ' Blank out the current strategy position...
                        m_fgTradeItems.TextMatrix(Row, TradeItemCol(eGDCol_StrategyPos)) = ""
                    End If
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.ToggleTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisableAll
'' Description: Allow the user to disable all active automated trading items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DisableAll(ByVal strReason As String)
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to show the user
    Dim strReturn As String             ' Return value from the InfBox

    strMessage = "You are about to disable all active|automated trading items.||Do you want to continue?|"
    strReturn = InfBox(strMessage, "?", "+Yes|-No", "Confirmation")
    
    If strReturn = "Y" Then
        g.TradingItems.DisableTradeItems strReason, False, True, True, False, False
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTradeItemUI.DisableAll"
    
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

    m.bLoadingExpanded = False
    m.strGrid = ""
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.Class_Initialize"
    
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
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemUI.Class_Terminate"
    
End Sub

