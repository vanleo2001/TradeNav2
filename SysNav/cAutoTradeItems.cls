VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoTradeItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAutoTradeItems.cls
'' Description: Collection of automated trading items
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/04/2009   DAJ         Added DisableTradeItems call
'' 05/06/2009   DAJ         Fail on enable if no data in the bars
'' 03/11/2010   DAJ         Added Active property, moved RefreshUI
'' 03/18/2010   DAJ         Update trade console summary when activate/deactivate
'' 03/18/2010   DAJ         Check for g.ConsoleForms being nothing
'' 07/15/2011   DAJ         Enhancements for allowing auto trading on continuous contracts
'' 07/21/2011   DAJ         Pass account position on enable
'' 07/25/2011   DAJ         Delete items for account, remove item if account doesn't exist
'' 07/28/2011   DAJ         Fix delete function
'' 08/19/2011   DAJ         Added ReloadData function
'' 09/08/2011   DAJ         Handle deactivation/reactivation around a loss of stream
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 01/26/2012   DAJ         Added the RefreshPosition function
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 08/07/2012   DAJ         Only show number of enabled trade items on the trade console
'' 12/11/2012   DAJ         Handle broker status in automated trading item
'' 01/07/2013   DAJ         Fix for number of active auto trading items being doubled
'' 01/24/2013   DAJ         Call DoEvents in ReloadData loop to let things breathe
'' 04/03/2013   DAJ         Automated Strategy Baskets
'' 05/01/2013   DAJ         Shadow Trading
'' 05/13/2013   DAJ         Allow activation on a delayed symbol if SimStream or SimReplay account
'' 05/24/2013   DAJ         Speed enhancements
'' 05/29/2013   DAJ         Added queues for NewBarChecks and OnCloseChecks
'' 06/06/2013   DAJ         Fix cound of active automated trading item in toolbar
'' 06/07/2013   DAJ         Expose the check to see if a parent should be active
'' 08/16/2013   DAJ         Don't load deleted automated trading items
'' 09/06/2013   DAJ         Tweak to the "DoActionChecks" routine
'' 03/03/2014   DAJ         Log when cannot activate trade item
'' 03/04/2014   DAJ         Beefed up logging for parking/resubmitting items when stream goes down
'' 03/18/2014   DAJ         Placed a DoEvents in the ResubmitPreviouslyActive loop
'' 05/01/2014   DAJ         Queue up fills going to automated trading items
'' 05/06/2014   DAJ         Pass fill time to automated trading item in FillCheck
'' 05/29/2014   DAJ         Added IsStrategyBasketItemInAutoTradeItem and NumStrategyBasketItemsInAutoTradeItem
'' 06/06/2014   DAJ         Handle the enablements changing for Trade Clone items
'' 06/09/2014   DAJ         Modified questions for disabling an automated basket
'' 08/19/2014   DAJ         Expose Strategy Basket Item Inputs
'' 09/24/2014   DAJ         Do the DoEvents in the ResubmitPreviouslyActive loop for everybody
'' 10/14/2014   DAJ         Added optional arguments to DisableTradeItems
'' 10/15/2014   DAJ         Added ClosePosition to DisableTradeItems
'' 10/24/2014   DAJ         Fill Display
'' 10/27/2014   DAJ         Pass in fill to FillCheck instead of individual fields
'' 10/28/2014   DAJ         Pass order into the automated trading item for a FillCheck
'' 11/10/2014   DAJ         Fix orphaned automated trading items on load
'' 11/11/2014   DAJ         Expose access to the previously active collection
'' 11/24/2014   DAJ         Don't do a new bar check if we did a fill or on-close check for same item
'' 01/21/2015   DAJ         Added additional logging in SubmitPreviouslyActive routine
'' 02/18/2015   DAJ         Added the FillCheckQueued function
'' 03/06/2015   DAJ         Fixed the FillCheckQueued function
'' 05/05/2015   DAJ         When loading, delete auto trade items for which the basket no longer exists
'' 01/08/2016   DAJ         Added support for CurrentBarChecks
'' 01/08/2016   DAJ         Added AccountUpdated function
'' 03/29/2016   DAJ         Don't remove item from previously submitted list if enable failed
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    AutoTradeItems As cGdTree           ' Collection of automated trading items
    
    astrLastActive As cGdArray          ' Array of items that were last active
    dLastDisabled As Double             ' Time the items were disabled
    lNumActive As Long                  ' Number of items that are active
    
    alNewBarChecks As cGdArray          ' Array of items that need a next bar check done
    alOnCloseChecks As cGdArray         ' Array of items that need an on-close check done
    alCurrentBarChecks As cGdArray      ' Array of items that need a current bar check done
    Fills As cGdTree                    ' Collection of fills waiting to be handled
    Orders As cGdTree                   ' Collection of orders for the fills in the collection
End Type
Private m As mPrivate

Public Property Get Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.AutoTradeItems.Exists(vIndexOrKey)
End Property

Public Property Get Item(ByVal vIndexOrKey As Variant) As cAutoTradeItem
Attribute Item.VB_UserMemId = 0
    If Exists(vIndexOrKey) Then
        Set Item = m.AutoTradeItems.Item(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, TradeItem As cAutoTradeItem)
    If Exists(vIndexOrKey) Then
        Set m.AutoTradeItems(vIndexOrKey) = TradeItem
    Else
        m.AutoTradeItems.Add TradeItem, Str(TradeItem.AutoTradeItemID)
    End If
End Property

Public Property Get Count() As Long
    Count = m.AutoTradeItems.Count
End Property

Public Property Get NumActive() As Long
    NumActive = m.lNumActive
End Property
Public Property Let NumActive(ByVal lNumActive As Long)
    If m.lNumActive <> lNumActive Then
        m.lNumActive = lNumActive
        If Not g.ConsoleForms Is Nothing Then
            g.ConsoleForms.NumVisible(eGDConsoleForm_AutoTrading) = lNumActive
            DebugLog "Number of active trading items changed to " & Str(lNumActive)
        End If
    End If
End Property

Public Sub Add(TradeItem As cAutoTradeItem)
    If Exists(Str(TradeItem.AutoTradeItemID)) Then
        Set m.AutoTradeItems(Str(TradeItem.AutoTradeItemID)) = TradeItem
    Else
        m.AutoTradeItems.Add TradeItem, Str(TradeItem.AutoTradeItemID)
    End If
End Sub

Public Sub Remove(ByVal vIndexOrKey As Variant)
    If Exists(vIndexOrKey) Then
        m.AutoTradeItems.Remove vIndexOrKey
    End If
End Sub

Public Function Active(ByVal vIndexOrKey As Variant) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Exists(vIndexOrKey) Then
        bReturn = Item(vIndexOrKey).Active
    End If
    
    Active = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.Active"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the automated trading items out of the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim TradeItem As cAutoTradeItem     ' Automated trading item object
    Dim dLoadTime As Double             ' Time it took to load
    Dim strReason As String             ' Reason we changed the parent ID
    Dim astrDeleted As cGdArray         ' List of items deleted
    
    Set astrDeleted = New cGdArray
    astrDeleted.Create eGDARRAY_Strings
    
    dLoadTime = gdTickCount
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] WHERE [Deleted]=0;", dbOpenDynaset)
    Do While Not rs.EOF
        Set TradeItem = New cAutoTradeItem
        ' TLB 4/30/2012: don't need to load bars during startup
        ' (the bars will be loaded as each item gets activated -- after streaming starts)
        If TradeItem.Load(rs!TradingItemID, False, rs) Then
            If StrategyBasketExists(rs!StrategyBasketID) Then
                ' DAJ 08/16/2013: Making this check up in the recordset now...
                'If TradeItem.Deleted = False Then
                    ' DAJ 11/10/2014: We have had several issues that have caused "orphaned" automated trading
                    ' items, so when we load the item, if the parent is deleted or no longer a parent, clear
                    ' the parent ID on this item...
                    If TradeItem.ParentID > 0 Then
                        strReason = ""
                        
                        Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] WHERE [TradingItemID]=" & Str(TradeItem.ParentID) & ";", dbOpenDynaset)
                        If rs2.BOF And rs2.EOF Then
                            strReason = "Parent no longer exists"
                        ElseIf rs2!ParentID <> -1 Then
                            strReason = "Parent is not a parent"
                        ElseIf rs2!Deleted = -1 Then
                            strReason = "Parent is marked as deleted"
                        End If
                        
                        If Len(strReason) > 0 Then
                            TradeItem.DumpDebug "Changing ParentID from " & Str(TradeItem.ParentID) & " to 0 because '" & strReason & "'"
                            TradeItem.ParentID = 0
                            TradeItem.Save
                        End If
                    End If
                    
                    If g.Broker.AccountExists(TradeItem.AccountID) Then
                        m.AutoTradeItems.Add TradeItem, Str(TradeItem.AutoTradeItemID)
                    Else
                        TradeItem.Delete
                    End If
                'End If
            Else
                astrDeleted.Add TradeItem.Name
                TradeItem.Delete
            End If
        End If
        
        rs.MoveNext
    Loop
    
    mSysNav.CreateGuruAutoTradeItems
    
    If astrDeleted.Size > 0 Then
        InfBox "The following automated trading items have|been deleted because the underlying|strategy basket no longer exists:||" & astrDeleted.JoinFields("|") & "|", "", , "Automated Trading Items"
    End If
    
    dLoadTime = gdTickCount - dLoadTime

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Enable
'' Description: Enable the given automated trading item if possible
'' Inputs:      Auto Trade Item ID, Believe Position?, Account Position object,
''              Show Message?
'' Returns:     True if able to enable, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Enable(ByVal lAutoTradeItemID As Long, Optional ByVal bBelieveCurrentPosition As Boolean = False, Optional ByVal AcctPos As cAccountPosition = Nothing, Optional ByVal bShowMessage As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated trading item to attempt to enable
    Dim bReturn As Boolean              ' Return value from the function
    Dim ParentItem As cAutoTradeItem    ' Parent item for an automated strategy basket
    
    bReturn = False
    Set TradeItem = Item(Str(lAutoTradeItemID))
    If Not TradeItem Is Nothing Then
        If TradeItem.Active = False Then
            If CanActivate(TradeItem, bShowMessage) Then
                If TradeItem.Load(lAutoTradeItemID) Then
                    Item(Str(lAutoTradeItemID)) = TradeItem
                    bReturn = TradeItem.Enable(bBelieveCurrentPosition, AcctPos)
                    'If bReturn = True Then
                    '    NumActive = NumActive + 1&
                    'End If
                Else
                    DebugLog "Not enabling " & Str(lAutoTradeItemID) & " because TradeItem not loaded"
                End If
            End If
        Else
            DebugLog "Not enabling " & Str(lAutoTradeItemID) & " because TradeItem already active"
            bReturn = True
        End If
        
        If bReturn = False Then
            TradeItem.RefreshUI
        End If
        
        If (bReturn = True) And (TradeItem.ParentID > 0&) Then
            Set ParentItem = Item(Str(TradeItem.ParentID))
            If Not ParentItem Is Nothing Then
                If ParentItem.Active = False Then
                    If ParentItem.Enable Then
                        ParentItem.RefreshUI
                    End If
                End If
            End If
        End If
                
        ' 03/18/2010 DAJ: Do this so that we can update the color on the Auto Trading
        ' caption on the summary toolbar...
        'If Not g.ConsoleForms Is Nothing Then
        '    g.ConsoleForms.NumVisible(eGDConsoleForm_AutoTrading) = NumActive
        'End If
    Else
        DebugLog "Not enabling " & Str(lAutoTradeItemID) & " because TradeItem is Nothing"
    End If
    
    Enable = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.Enable"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disable
'' Description: Disable the given automated trading item if possible
'' Inputs:      Auto Trade Item ID, Confirm Position Close?, Reason,
''              Disabled from code?, Force Refresh?
'' Returns:     True if able to disable, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Disable(ByVal lAutoTradeItemID As Long, Optional ByVal bConfirmPositionClose As Boolean = True, Optional ByVal strReason As String = "", Optional ByVal bDisabledFromCode As Boolean = False, Optional ByVal bForceRefresh As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated trading item to attempt to enable
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    Set TradeItem = Item(Str(lAutoTradeItemID))
    If Not TradeItem Is Nothing Then
        If TradeItem.ParentID = -1& Then
            bReturn = DisableBasket(TradeItem, bConfirmPositionClose, , strReason, bDisabledFromCode)
        Else
            bReturn = DisableTradeItem(TradeItem, bConfirmPositionClose, , strReason, bDisabledFromCode)
        End If
    End If
    
    Disable = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.Disable"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderCallback
'' Description: Send the given order to the appropriate automated trading item
'' Inputs:      Auto Trade Item ID, Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderCallback(ByVal lAutoTradeItemID As Long, ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    If Exists(Str(lAutoTradeItemID)) Then
        Item(Str(lAutoTradeItemID)).OrderStatusChange Order
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.OrderCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasActiveAutoTradeItems
'' Description: Determine if there are active auto trading items
'' Inputs:      None
'' Returns:     True if Active Items, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasActiveAutoTradeItems() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    For lIndex = 1 To m.AutoTradeItems.Count
        If Item(lIndex).Active Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    HasActiveAutoTradeItems = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.HasActiveAutoTradeItems"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsStrategyAutoTrading
'' Description: Determine if the given strategy is in an active auto-trade
'' Inputs:      Strategy ID
'' Returns:     True if Auto-Trading, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsStrategyAutoTrading(ByVal lStrategyID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    For lIndex = 1 To m.AutoTradeItems.Count
        If (Item(lIndex).StrategyID = lStrategyID) Then
            If Item(lIndex).Active Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    IsStrategyAutoTrading = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.IsStrategyAutoTrading"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsStrategyBasketInAutoTradeItem
'' Description: Determine if the given strategy basket is in an auto-trade item
'' Inputs:      Strategy Basket ID
'' Returns:     True if Auto-Trading, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsStrategyBasketInAutoTradeItem(ByVal lBasketID As Long, Optional BasketTradeItem As cAutoTradeItem) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        If (TradeItem.StrategyBasketID = lBasketID) Then
            Set BasketTradeItem = TradeItem
            
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    IsStrategyBasketInAutoTradeItem = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.IsStrategyBasketInAutoTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsStrategyBasketItemInAutoTradeItem
'' Description: Determine if the given strategy basket item is in an auto-trade item
'' Inputs:      Strategy Basket Item ID
'' Returns:     True if Auto-Trading, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsStrategyBasketItemInAutoTradeItem(ByVal lBasketItemID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        If (TradeItem.StrategyBasketItemID = lBasketItemID) Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    IsStrategyBasketItemInAutoTradeItem = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.IsStrategyBasketItemInAutoTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumStrategyBasketItemsInAutoTradeItem
'' Description: Determine the number of basket items in an auto-trade item for a basket
'' Inputs:      Strategy Basket ID
'' Returns:     True if Auto-Trading, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NumStrategyBasketItemsInAutoTradeItem(ByVal lBasketID As Long) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim lReturn As Long                 ' Return value from the function
    
    lReturn = 0&
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        If (TradeItem.StrategyBasketID = lBasketID) And (TradeItem.ParentID <> -1&) Then
            lReturn = lReturn + 1&
            Exit For
        End If
    Next lIndex
    
    NumStrategyBasketItemsInAutoTradeItem = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.NumStrategyBasketItemsInAutoTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsStrategyBasketAutoTrading
'' Description: Determine if the given strategy basket is in an active auto-trade
'' Inputs:      Strategy Basket ID
'' Returns:     True if Auto-Trading, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsStrategyBasketAutoTrading(ByVal lBasketID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        If (TradeItem.StrategyBasketID = lBasketID) Then
            If TradeItem.Active Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    IsStrategyBasketAutoTrading = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.IsStrategyBasketAutoTrading"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsStrategyBasketInPosition
'' Description: Determine if the given strategy basket is in an automated
''              trading item that is in a position
'' Inputs:      Strategy Basket ID
'' Returns:     True if in position, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsStrategyBasketInPosition(ByVal lBasketID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        If (TradeItem.StrategyBasketID = lBasketID) Then
            If (TradeItem.ParentID > 0&) Then
                If (TradeItem.CurrentPosition <> 0&) Then
                    bReturn = True
                    Exit For
                End If
            End If
        End If
    Next lIndex
    
    IsStrategyBasketInPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.IsStrategyBasketInPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasNonExistentItemsInPosition
'' Description: Determine if there are nonexistent items in a position
'' Inputs:      Basket ID, List of Item ID's
'' Returns:     True if Nonexistent items in a position, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasNonExistentItemsInPosition(ByVal lBasketID As Long, ByVal alItemIds As cGdArray) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        If (TradeItem.StrategyBasketID = lBasketID) Then
            If alItemIds.BinarySearch(TradeItem.StrategyBasketItemID) = False Then
                If (TradeItem.CurrentPosition <> 0&) Then
                    bReturn = True
                    Exit For
                End If
            End If
        End If
    Next lIndex
    
    HasNonExistentItemsInPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.HasNonExistentItemsInPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTradingItemActive
'' Description: Determine whether the given trading item is active
'' Inputs:      Trading Item ID
'' Returns:     True if Active, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTradingItemActive(ByVal lAutoTradingItemID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    If Exists(Str(lAutoTradingItemID)) Then
        bReturn = Item(Str(lAutoTradingItemID)).Active
    End If
    
    IsTradingItemActive = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.IsTradingItemActive"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisableTradeItems
'' Description: Disable all the trading items
'' Inputs:      Reason, Disabled from code?, Confirm Position Close?, Close
''              Position?, Skip the orders question?, Add to Last Active collection?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisableTradeItems(Optional ByVal strReason As String = "", Optional ByVal bDisabledFromCode As Boolean = False, Optional ByVal bConfirmPositionClose As Boolean = False, Optional ByVal bClosePosition As Boolean = False, Optional ByVal bSkipQuestion As Boolean = True, Optional ByVal bAddToLastActive As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Item out of the collection
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        If TradeItem.Active Then
            If TradeItem.ParentID >= 0& Then
                DisableTradeItem TradeItem, bConfirmPositionClose, bClosePosition, strReason, bDisabledFromCode, False, bSkipQuestion
                
                If bAddToLastActive = True Then
                    DebugLog "Adding Trading Item to list of previously active: '" & m.AutoTradeItems.Key(lIndex) & "'"
                    m.astrLastActive.Add m.AutoTradeItems.Key(lIndex)
                End If
            End If
        End If
    Next lIndex
    
    If m.astrLastActive.Size > 0 Then
        m.dLastDisabled = CurrentTime
        m.astrLastActive.Sort
        DebugLog "Last Disabled for Automated Trading Items set to: " & DateFormat(m.dLastDisabled, MM_DD_YYYY, HH_MM_SS)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.DisableTradeItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisableTradeItemsForSymbol
'' Description: Disable trading items for the given account and symbol
'' Inputs:      Account, Symbol, Reason, Disabled from code?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisableTradeItemsForSymbol(ByVal vAccountIDorNumber As Variant, ByVal vSymbolOrSymbolID As Variant, Optional ByVal strReason As String = "", Optional ByVal bDisabledFromCode As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID passed in
    Dim strSymbol As String             ' Symbol passed in
    Dim TradeItem As cAutoTradeItem     ' Automated trading item from the collection
    
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    If lSymbolID = 0 Then
        vSymbolOrSymbolID = strSymbol
    Else
        vSymbolOrSymbolID = lSymbolID
    End If
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        
        If TradeItem.Active Then
            If TradeItem.ParentID >= 0& Then
                If TradeItem.AccountID = g.Broker.GetAccountID(vAccountIDorNumber) Then
                    If ConvertToTradeSymbol(TradeItem.SymbolOrSymbolID) = vSymbolOrSymbolID Then
                        DisableTradeItem TradeItem, False, False, strReason, bDisabledFromCode, False, True
                    End If
                End If
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAutoTradeItems.DisableTradeItemsForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPosition
'' Description: Retrieve the current position for the given auto trade item
'' Inputs:      Automated Trading Item ID
'' Returns:     Current Position for that item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentPosition(ByVal lAutoTradingItemID As Long, Optional ByVal vSymbolOrSymbolID As Variant = "", Optional ByVal lAccountID As Long = 0&) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value from the function
    Dim lSymbolID As Long               ' Symbol ID for symbol passed in
    Dim TradeItem As cAutoTradeItem     ' Trading item
    
    lReturn = 0&
    Set TradeItem = Item(Str(lAutoTradingItemID))
    If Not TradeItem Is Nothing Then
        If vSymbolOrSymbolID <> "" Then
            lSymbolID = GetSymbolID(vSymbolOrSymbolID)
            If lSymbolID = 0& Then
                vSymbolOrSymbolID = GetSymbol(vSymbolOrSymbolID)
            Else
                vSymbolOrSymbolID = lSymbolID
            End If
        End If
        
        If (vSymbolOrSymbolID = "") Or (vSymbolOrSymbolID = ConvertToTradeSymbol(TradeItem.SymbolOrSymbolID)) Then
            If (lAccountID = 0&) Or (lAccountID = TradeItem.AccountID) Then
                lReturn = TradeItem.CurrentPosition
            End If
        End If
    End If
    
    CurrentPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.CurrentPosition"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NameForID
'' Description: Determine the name for the strategy with the given ID
'' Inputs:      Auto Trade ID
'' Returns:     Auto Trade Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NameForID(ByVal lAutoTradingItemID As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If lAutoTradingItemID > 0& Then
        If Exists(Str(lAutoTradingItemID)) Then
            strReturn = Item(Str(lAutoTradingItemID)).Name
        Else
            strReturn = "No Longer Exists"
        End If
    End If
    
    NameForID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.NameForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixAccountPositionIds
'' Description: Fix any account position ID's that are still zero
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FixAccountPositionIds()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim AcctPos As cAccountPosition     ' Account position object
    Dim vTradeSymbolOrID As Variant     ' Trade Symbol or ID for the Trade Item symbol
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        If TradeItem.AccountPositionID = 0 Then
            vTradeSymbolOrID = ConvertToTradeSymbol(TradeItem.SymbolOrSymbolID, CurrentTime("", TradeItem.Symbol))
            Set AcctPos = g.Broker.FillSummary(TradeItem.AccountID, vTradeSymbolOrID, TradeItem.AutoTradeItemID)
            If Not AcctPos Is Nothing Then
                TradeItem.AccountPositionID = AcctPos.AccountPositionID
                TradeItem.Save
                
                Item(lIndex) = TradeItem
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.FixAccountPositionIds"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete the automated trading item with the given ID
'' Inputs:      Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete(ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    If Exists(Str(lAutoTradeItemID)) Then
        Item(Str(lAutoTradeItemID)).Delete
        g.Alerts.RemoveAlertsForAutoTradeItem lAutoTradeItemID
    
        Remove Str(lAutoTradeItemID)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.Delete"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteForAccount
'' Description: Delete any automated trading items for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteForAccount(ByVal vAccountNumberOrID As Variant)
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID for the given account
    Dim lIndex As Long                  ' Index into a for loop
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    
    For lIndex = Count To 1 Step -1
        If Item(lIndex).AccountID = lAccountID Then
            Item(lIndex).Delete
            g.Alerts.RemoveAlertsForAutoTradeItem Item(lIndex).AutoTradeItemID
            
            If FormIsLoaded("frmTradeItems") Then
                frmTradeItems.RefreshTradeItem Item(lIndex).AutoTradeItemID, True
            End If
        
            Remove lIndex
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.DeleteForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteForBasket
'' Description: Delete any automated trading items for the strategy basket ID
'' Inputs:      Strategy Basket ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteForBasket(ByVal lStrategyBasketID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = Count To 1 Step -1
        If Item(lIndex).StrategyBasketID = lStrategyBasketID Then
            Item(lIndex).Delete
            g.Alerts.RemoveAlertsForAutoTradeItem Item(lIndex).AutoTradeItemID
            
            If FormIsLoaded("frmTradeItems") Then
                frmTradeItems.RefreshTradeItem Item(lIndex).AutoTradeItemID, True
            End If
        
            Remove lIndex
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.DeleteForBasket"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReloadData
'' Description: Tell each active automated trading item to reload its data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReloadData()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To Count
        If Item(lIndex).Active Then
            Item(lIndex).ReloadData
        End If
        
        ' DAJ 01/24/2013: Call DoEvents every five times through the loop to let
        ' other things happen.  Noticed in Brady's logs that on a machine with 28
        ' trading items, the stream and everything gets stopped for over a minute
        ' while this loop is happening...
        If lIndex Mod 5 = 0 Then
            DoEvents
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.ReloadData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Tell each item to update its bars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBars()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    If m.astrLastActive.Size > 0 Then
        SubmitPreviouslyActive
    End If
    
    For lIndex = 1 To Count
        Item(lIndex).UpdateBars
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.UpdateBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPosition
'' Description: Refresh the position for the given automated trading item ID
'' Inputs:      Automated Trading Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPosition(ByVal lAutoTradeID As Long)
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated trading item object
    
    If Exists(Str(lAutoTradeID)) Then
        Item(Str(lAutoTradeID)).RefreshPosition
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.RefreshPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectionStatusChanged
'' Description: Notification that the connection status for the given broker
''              has just changed
'' Inputs:      Broker, Connection Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ConnectionStatusChanged(ByVal nBroker As eTT_AccountType, ByVal nNewStatus As eGDConnectionStatus)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Auto trade item out of the collection
    
    For lIndex = 1 To m.AutoTradeItems.Count
        Set TradeItem = Item(lIndex)
        If TradeItem.Broker = nBroker Then
            TradeItem.ConnectionStatusChanged nBroker, nNewStatus
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.ConnectionStatusChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParentID
'' Description: Get the parent ID for the given automated trading item ID
'' Inputs:      Automated Trading Item ID
'' Returns:     Parent ID ( Null if none )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParentID(ByVal lAutoTradeItemID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = kNullData
    If Exists(Str(lAutoTradeItemID)) Then
        If Item(Str(lAutoTradeItemID)).ParentID > 0& Then
            lReturn = Item(Str(lAutoTradeItemID)).ParentID
        End If
    End If
    
    ParentID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.ParentID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NameExistsInDb
'' Description: Does the given name exist in the database?
'' Inputs:      Name to Check, Ignore ID
'' Returns:     True if Exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NameExistsInDb(ByVal strAutoTradeName As String, Optional ByVal lIgnoreID As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bReturn As Boolean              ' Return value for the function
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] WHERE [Name]='" & strAutoTradeName & "';", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        bReturn = False
    Else
        bReturn = (rs!TradingItemID <> lIgnoreID)
    End If
    
    NameExistsInDb = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.NameExistsInDb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDeletedItem
'' Description: Get a deleted item with the given information if possible
'' Inputs:      Parent ID, Key
'' Returns:     Auto Trade Item ( Nothing if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDeletedItem(ByVal lParentID As Long, ByVal strKey As String) As cAutoTradeItem
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Trade item to return from the function
    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem] " & _
                "WHERE [ParentID]=" & Str(lParentID) & " AND [StrategyBasketItemKey]='" & strKey & "' AND [Deleted]=-1;", dbOpenDynaset)
    If rs.BOF And rs.EOF Then
        Set TradeItem = Nothing
    Else
        Set TradeItem = New cAutoTradeItem
        If TradeItem.Load(rs!TradingItemID, False, rs) = False Then
            Set TradeItem = Nothing
        End If
    End If
    
    Set GetDeletedItem = TradeItem

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.GetDeletedItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddNewBarCheck
'' Description: Add an item to the list that needs a new bar check
'' Inputs:      Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddNewBarCheck(ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' It the ID already in the array?
    
    bFound = False
    For lIndex = 0 To m.alNewBarChecks.Size - 1
        If m.alNewBarChecks(lIndex) = lAutoTradeItemID Then
            bFound = True
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        m.alNewBarChecks.Add lAutoTradeItemID, 0&
        frmOnlineBroker.tmrAutoTradeAction.Enabled = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.AddNewBarCheck"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOnCloseCheck
'' Description: Add an item to the list that needs an on-close check
'' Inputs:      Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddOnCloseCheck(ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' It the ID already in the array?
    
    bFound = False
    For lIndex = 0 To m.alOnCloseChecks.Size - 1
        If m.alOnCloseChecks(lIndex) = lAutoTradeItemID Then
            bFound = True
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        m.alOnCloseChecks.Add lAutoTradeItemID, 0&
        frmOnlineBroker.tmrAutoTradeAction.Enabled = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.AddOnCloseCheck"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddCurrentBarCheck
'' Description: Add an item to the list that needs a current bar check
'' Inputs:      Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddCurrentBarCheck(ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' It the ID already in the array?
    
    bFound = False
    For lIndex = 0 To m.alCurrentBarChecks.Size - 1
        If m.alCurrentBarChecks(lIndex) = lAutoTradeItemID Then
            bFound = True
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        m.alCurrentBarChecks.Add lAutoTradeItemID, 0&
        frmOnlineBroker.tmrAutoTradeAction.Enabled = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.AddCurrentBarCheck"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFillCheck
'' Description: Add an item to the list that needs a fill check
'' Inputs:      Fill, Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddFillCheck(ByVal Fill As cPtFill, Order As cPtOrder)
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated trading item for the fill
    Dim strFillDisplay As String        ' Fill display
    Dim strKey As String                ' Key into the collection
    
    If Exists(Str(Fill.AutoTradingItemID)) Then
        Set TradeItem = Item(Str(Fill.AutoTradingItemID))
        
        strFillDisplay = mTradeTracker.FillDisplay(Fill, True, False, True, True, True, False, False)
        
        strKey = Str(Fill.FillID)
        If m.Fills.Exists(strKey) = False Then
            TradeItem.DumpDebug "FillCheck added to queue for '" & strFillDisplay & "'"
            m.Fills.Add Fill, strKey
            frmOnlineBroker.tmrAutoTradeAction.Enabled = True
        Else
            TradeItem.DumpDebug "FillCheck already in queue for '" & strFillDisplay & "'"
        End If
        
        strKey = Str(Order.OrderID)
        If m.Orders.Exists(strKey) = True Then
            Set m.Orders(strKey) = Order
        Else
            m.Orders.Add Order, strKey
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.AddFillCheck"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoActionChecks
'' Description: Do the appropriate action checks
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoActionChecks()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastItem As Long               ' Last item in the array
    Dim lTradeItemID As Long            ' Trading item ID
    Dim lSize As Long                   ' Original size of the array
    Dim Fill As cPtFill                 ' Fill from the collection
    Dim Order As cPtOrder               ' Order from the collection
    Dim strKey As String                ' Key into the orders collection
    Dim alSkipNewBar As cGdArray        ' Array of items for which we want to skip the new bar check
    Dim alSkipCurrentBar As cGdArray    ' Array of items for which we want to skip the current bar check
    Dim lPos As Long                    ' Position in the array
    
    Set alSkipNewBar = New cGdArray
    alSkipNewBar.Create eGDARRAY_Longs
    
    Set alSkipCurrentBar = New cGdArray
    alSkipCurrentBar.Create eGDARRAY_Longs
    
    Do While m.Fills.Count > 0
        Set Fill = m.Fills(1)
        m.Fills.Remove 1
        
        strKey = Str(Fill.OrderID)
        If m.Orders.Exists(strKey) Then
            Set Order = m.Orders(strKey)
        Else
            Set Order = Nothing
        End If
        
        If m.Fills.Count = 0 Then
            m.Orders.Clear
        End If
        
        DoFillCheck Fill, Order
        
        ' Add the automated trading item identifier to the list of items for which
        ' we want to skip the new bar check if it is in the queue...
        If alSkipNewBar.BinarySearch(Fill.AutoTradingItemID, lPos) = False Then
            alSkipNewBar.Add Fill.AutoTradingItemID, lPos
        End If
        
        ' Add the automated trading item identifier to the list of items for which
        ' we want to skip the current bar check if it is in the queue...
        If alSkipCurrentBar.BinarySearch(Fill.AutoTradingItemID, lPos) = False Then
            alSkipCurrentBar.Add Fill.AutoTradingItemID, lPos
        End If
    Loop
    
    lSize = m.alOnCloseChecks.Size
    For lIndex = lSize - 1 To 0 Step -1
        lLastItem = m.alOnCloseChecks.Size - 1
        lTradeItemID = m.alOnCloseChecks(lLastItem)
        
        m.alOnCloseChecks.Remove lLastItem
        If Exists(Str(lTradeItemID)) Then
            Item(Str(lTradeItemID)).OnCloseCheck
        End If
        
        ' Add the automated trading item identifier to the list of items for which
        ' we want to skip the new bar check if it is in the queue...
        If alSkipNewBar.BinarySearch(lTradeItemID, lPos) = False Then
            alSkipNewBar.Add lTradeItemID, lPos
        End If
        
        ' Add the automated trading item identifier to the list of items for which
        ' we want to skip the current bar check if it is in the queue...
        If alSkipCurrentBar.BinarySearch(lTradeItemID, lPos) = False Then
            alSkipCurrentBar.Add lTradeItemID, lPos
        End If
    Next lIndex

    lSize = m.alNewBarChecks.Size
    For lIndex = lSize - 1 To 0 Step -1
        lLastItem = m.alNewBarChecks.Size - 1
        lTradeItemID = m.alNewBarChecks(lLastItem)
        
        m.alNewBarChecks.Remove lLastItem
        If Exists(Str(lTradeItemID)) Then
            If alSkipNewBar.BinarySearch(lTradeItemID) = True Then
                Item(Str(lTradeItemID)).DumpDebug "NewBarCheck not done because we did a fill check or an on-close check for this item"
            Else
                Item(Str(lTradeItemID)).NewBarCheck
            End If
        End If
        
        ' Add the automated trading item identifier to the list of items for which
        ' we want to skip the current bar check if it is in the queue...
        If alSkipCurrentBar.BinarySearch(lTradeItemID, lPos) = False Then
            alSkipCurrentBar.Add lTradeItemID, lPos
        End If
    Next lIndex

    lSize = m.alCurrentBarChecks.Size
    For lIndex = lSize - 1 To 0 Step -1
        lLastItem = m.alCurrentBarChecks.Size - 1
        lTradeItemID = m.alCurrentBarChecks(lLastItem)
        
        m.alCurrentBarChecks.Remove lLastItem
        If Exists(Str(lTradeItemID)) Then
            If alSkipCurrentBar.BinarySearch(lTradeItemID) = True Then
                Item(Str(lTradeItemID)).DumpDebug "CurrentBarCheck not done because we did a fill check, on-close check, or new bar check for this item"
            Else
                Item(Str(lTradeItemID)).CurrentBarCheck
            End If
        End If
    Next lIndex
    
    If m.Fills.Count + m.alNewBarChecks.Size + m.alOnCloseChecks.Size + m.alCurrentBarChecks.Size = 0 Then
        frmOnlineBroker.tmrAutoTradeAction.Enabled = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.DoActionChecks"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckParent
'' Description: Check the parent to see if all of its children are active
'' Inputs:      Parent ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckParent(ByVal lParentAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Dim ParentItem As cAutoTradeItem    ' Parent item if part of a basket

    If lParentAutoTradeItemID > 0& Then
        Set ParentItem = Item(Str(lParentAutoTradeItemID))
        If ParentItem.Active = True Then
            If ParentHasActiveChildren(ParentItem.AutoTradeItemID) = False Then
                ParentItem.Disable
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.CheckParent"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnablementsChanges
'' Description: Notification that the enablements have changed for the user
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EnablementsChanged()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Trade item object
    Dim lMaxUnits As Long               ' Max units for the trade item
    Dim lMaxQuantity As Long            ' Max quantity for the individual basket item
    Dim lOldQuantity As Long            ' Old quantity next entry
    
    For lIndex = 1 To Count
        Set TradeItem = Item(lIndex)
        lMaxUnits = mSysNav.MaxUnitsForAutoTrade(TradeItem)
        If (lMaxUnits <> Abs(kNullData)) And (TradeItem.StrategyBasketItemMult <> 0) Then
            lMaxQuantity = lMaxUnits * TradeItem.StrategyBasketItemMult
            If TradeItem.QtyNextEntry > lMaxQuantity Then
                lOldQuantity = TradeItem.QtyNextEntry
                TradeItem.QtyNextEntry = lMaxQuantity
                TradeItem.DumpDebug "Quantity Next Entry changed from " & Str(lOldQuantity) & " to " & Str(lMaxQuantity) & " because enablements changed"
                TradeItem.Save
                
                Item(lIndex) = TradeItem
                
                TradeItem.RefreshUI
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.EnablementsChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetItemForBasketItem
'' Description: Get the automated trading item for the given information
'' Inputs:      Parent ID, Strategy Basket Item ID
'' Returns:     Auto Trade Item ( Nothing if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetItemForBasketItem(ByVal lParentID As Long, ByVal lStrategyBasketItemID As Long) As cAutoTradeItem
On Error GoTo ErrSection:

    Dim ReturnItem As cAutoTradeItem    ' Trade item object to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Trade item object
    
    Set ReturnItem = Nothing
    For lIndex = 1 To Count
        Set TradeItem = Item(lIndex)
        If (TradeItem.ParentID = lParentID) And (TradeItem.StrategyBasketItemID = lStrategyBasketItemID) Then
            Set ReturnItem = TradeItem
            Exit For
        End If
    Next lIndex
    
    Set GetItemForBasketItem = ReturnItem

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.GetItemForBasketItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PreviouslyActive
'' Description: Is the given ID in the previously active collection AND it has
''              been less then five minutes since it was put there?
'' Inputs:      Item ID
'' Returns:     True if in collection less then five minutes, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PreviouslyActive(ByVal lAutoTradeItemID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If m.astrLastActive.BinarySearch(Str(lAutoTradeItemID)) = True Then
        If CurrentTime <= (m.dLastDisabled + (5# / 1440#)) Then
            bReturn = True
        End If
    End If
    
    PreviouslyActive = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.PreviouslyActive"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCheckQueued
'' Description: Determine if a fill check is queued for the given automated
''              trading item ID
'' Inputs:      Automated Trading Item ID
'' Returns:     True if queued, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillCheckQueued(ByVal lAutoTradeItemID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill from the collection
    
    bReturn = False
    For lIndex = 1 To m.Fills.Count
        Set Fill = m.Fills(lIndex)
        If Fill.AutoTradingItemID = lAutoTradeItemID Then
            bReturn = True
            Exit For
        End If
    Next lIndex

    FillCheckQueued = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.FillCheckQueued"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountUpdated
'' Description: Notification that the given account has been updated
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AccountUpdated(ByVal lAccountID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    If FormIsLoaded("frmTradeItems") Then
        For lIndex = 1 To Count
            If Item(lIndex).AccountID = lAccountID Then
                frmTradeItems.RefreshTradeItem Item(lIndex).AutoTradeItemID
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.AccountUpdated"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitPreviouslyActive
'' Description: Attempt to submit previously active auto trade items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitPreviouslyActive()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bClear As Boolean               ' Clear information?
    Dim strReturn As String             ' Return value from the user question
    Dim strReason As String             ' Reason we are clearing out the array
    Dim bEnabled As Boolean             ' Was the item successfully enabled?
    Static bAlreadyAsked As Boolean     ' Have we already asked the user?
    Static iLevel As Integer            ' Level of the check we have made it to

    bClear = False
    If m.astrLastActive.Size > 0 Then
        If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
            If g.RealTime.Reconnecting = False Then
                If g.Broker.PositionsToVerify = False Then
                    iLevel = 0
                    
                    If CurrentTime <= (m.dLastDisabled + (5# / 1440#)) Then
                        If bAlreadyAsked = False Then
                            strReturn = InfBox("Would you like to reactivate automated trading items that were previously active?", "?", "+Yes|-No", "Reactivate Automated Trading Items", , 20)
                        End If
                        
                        If (strReturn = "Y") Or (bAlreadyAsked = True) Then
                            For lIndex = m.astrLastActive.Size - 1 To 0 Step -1
                                DebugLog "Submitting Previously Active Automated Trading Item: " & m.astrLastActive(lIndex) & " ( " & Str(lIndex) & " )"
                                If m.AutoTradeItems.Exists(m.astrLastActive(lIndex)) Then
                                    If Active(m.astrLastActive(lIndex)) = False Then
                                        DebugLog vbTab & "Enabling Previously Active Automated Trading Item: " & m.astrLastActive(lIndex) & " ( " & Str(lIndex) & " )"
                                        bEnabled = Enable(CLng(Val(m.astrLastActive(lIndex))), True, , False)
                                    Else
                                        DebugLog vbTab & "Not Enabling Previously Active Automated Trading Item: " & m.astrLastActive(lIndex) & " ( " & Str(lIndex) & " ) because it is already active"
                                    End If
                                    
                                    If bEnabled = True Then
                                        DebugLog vbTab & "Removing Previously Active Automated Trading Item: " & m.astrLastActive(lIndex) & " ( " & Str(lIndex) & " )"
                                        m.astrLastActive.Remove lIndex
                                    End If
                                Else
                                    DebugLog vbTab & "Not Enabling Previously Active Automated Trading Item: " & m.astrLastActive(lIndex) & " ( " & Str(lIndex) & " ) because it was not found"
                                End If
                                
                                ' DAJ 03/18/2014: Pete and Tim are finding that quite often on Sunday after the
                                ' stream is restarted, not all auto trade items reactivate.  In looking at Pete's
                                ' logs, it is because g.RealTime.IsServerActive is false.  By doing a DoEvents here
                                ' and allowing the system to "breathe", we may take care of that...
                                ' DAJ 09/24/2014: Tim ran into this again on one of his machines that doesn't have
                                ' the Files.EXE program.  Since more people are going to have a lot of automated
                                ' trading items, we will do the DoEvents here no matter what...
                                'If FileExist("C:\Common\Files.EXE") Then
                                    DoEvents
                                'End If
                            Next lIndex
                            
                            If m.astrLastActive.Size > 0 Then
                                bAlreadyAsked = True
                            Else
                                strReason = "Done"
                                bClear = True
                            End If
                        Else
                            strReason = "User chose not to reactivate items"
                            bClear = True
                        End If
                    Else
                        strReason = "Took too long to reconnect to streaming: Current = " & DateFormat(CurrentTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & "; Last Disabled = " & DateFormat(m.dLastDisabled, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                        bClear = True
                    End If
                ElseIf iLevel < 3 Then
                    DebugLog "Didn't submit previously active because there are still positions to verify"
                    iLevel = 3
                End If
            
                If bClear Then
                    DebugLog "Clearing Last Active array because '" & strReason & "'"
                    m.astrLastActive.Clear
                    m.dLastDisabled = kNullData
                    bAlreadyAsked = False
                End If
            ElseIf iLevel < 2 Then
                DebugLog "Didn't submit previously active because streaming still reconnecting"
                iLevel = 2
            End If
        ElseIf iLevel < 1 Then
            DebugLog "Didn't submit previously active because streaming not connected"
            iLevel = 1
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.SubmitPreviouslyActive"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanActivate
'' Description: Can the given trading item be activated?
'' Inputs:      Trade Item, Show Message?
'' Returns:     True if can activate, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CanActivate(ByVal TradeItem As cAutoTradeItem, Optional ByVal bShowMessage As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strMessage As String            ' Message to display to the user
    Dim lDelay As Long                  ' Delay
    
    bReturn = True
    strMessage = ""
    lDelay = g.RealTime.SymbolDelay(TradeItem.Symbol)
    
    If Not g.RealTime.IsServerActive(True) Then
        strMessage = "You cannot activate|" & TradeItem.Name & "|because streaming is not active"
        bReturn = False
    ElseIf (lDelay > 0) And (TradeItem.Broker <> eTT_AccountType_SimStream) And (TradeItem.Broker <> eTT_AccountType_SimReplay) Then
        strMessage = "You cannot activate|" & TradeItem.Name & "|because " & TradeItem.Symbol & " is a delayed symbol"
        bReturn = False
    ElseIf g.Broker.ConnectionStatusForAccount(TradeItem.AccountID) <> eGDConnectionStatus_Connected Then
        strMessage = "You cannot activate|" & TradeItem.Name & "|because you are not connected to the account"
        bReturn = False
    ElseIf g.Broker.PositionVerify(TradeItem.Broker) = True Then
        strMessage = "You cannot activate|" & TradeItem.Name & "|because positions are not verified yet"
        bReturn = False
    Else
        bReturn = AllowLiveAutoTrading(g.Broker.AccountTypeForID(TradeItem.AccountID), bShowMessage)
    End If
    
    If (Len(strMessage) > 0) Then
        TradeItem.DumpDebug Replace(strMessage, "|", " ") & " ( ShowMessage = " & Str(bShowMessage) & "; Delay = " & Str(lDelay) & " )"
        DebugLog "CanActivate failed for " & Str(TradeItem.AutoTradeItemID) & ": " & Replace(strMessage, "|", " ") & " ( ShowMessage = " & Str(bShowMessage) & "; Delay = " & Str(lDelay) & " )"
        
        If (bShowMessage = True) Then
            frmOnlineBroker.AddDialogMessage strMessage, "!", , "Automated Trading Error"
        End If
    End If
        
    CanActivate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.CanActivate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisableBasket
'' Description: Disable the given automated trading basket if possible
'' Inputs:      Auto Trade Basket, Confirm Position Close?, Close Position,
''              Reason, Disabled from code?, Skip Question?
'' Returns:     True if able to disable, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DisableBasket(TradeItem As cAutoTradeItem, Optional ByVal bConfirmPositionClose As Boolean = True, Optional ByVal bClosePosition As Boolean = True, Optional ByVal strReason As String = "", Optional ByVal bDisabledFromCode As Boolean = False, Optional bSkipQuestion As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim ChildItem As cAutoTradeItem     ' Child item for an automated strategy basket
    Dim bChildrenActive As Boolean      ' Are there still children active?
    Dim strReturn As String             ' Return from the user question
    Dim lItemsInPosition As Long        ' Number of basket items that are in a position
    
    bReturn = False
    If Not TradeItem Is Nothing Then
        If TradeItem.ParentID = -1& Then
            bReturn = True
            If TradeItem.Active = True Then
                If bSkipQuestion Then
                    strReturn = "T"
                Else
                    strReturn = InfBox("Turning off '" & TradeItem.Name & "' will cancel any open orders that have been generated for all items in the basket.||Do you want to continue?|", "?", "+Turn Off|-Leave On", "Disable Automated Basket")
                    TradeItem.DumpDebug "User answered '" & strReturn & "' to turning off the basket"
                End If
                
                If strReturn = "T" Then
                    lItemsInPosition = 0&
                    For lIndex = 1 To Count
                        Set ChildItem = Item(lIndex)
                        If ChildItem.ParentID = TradeItem.AutoTradeItemID Then
                            If ChildItem.Active = True Then
                                If ChildItem.CurrentPosition <> 0& Then
                                    lItemsInPosition = lItemsInPosition + 1&
                                End If
                            End If
                        End If
                    Next lIndex
                    
                    If lItemsInPosition > 1& Then
                        strReturn = InfBox("More than one of the automated trading items|in this basket are in a position.||What do you want to do with these positions?|", , "+-Keep All|Exit All|Ask per Item", "Disable Automated Basket")
                        TradeItem.DumpDebug "User answered '" & strReturn & "' to what to do with the positions in the basket"
                        
                        Select Case strReturn
                            Case "K"
                                bConfirmPositionClose = False
                                bClosePosition = False
                            Case "E"
                                bConfirmPositionClose = False
                                bClosePosition = True
                            Case "A"
                                bConfirmPositionClose = True
                                bClosePosition = True
                        End Select
                    End If
                    
                    For lIndex = 1 To Count
                        Set ChildItem = Item(lIndex)
                        If ChildItem.ParentID = TradeItem.AutoTradeItemID Then
                            If ChildItem.Active = True Then
                                If DisableTradeItem(ChildItem, bConfirmPositionClose, bClosePosition, strReason, bDisabledFromCode, True, True) = False Then
                                    bReturn = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next lIndex
                    
                    If bReturn = True Then
                        TradeItem.Disable
                    End If
                Else
                    bReturn = False
                End If
            End If
        End If
    End If
    
    DisableBasket = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.DisableBasket"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisableTradeItem
'' Description: Disable the given automated trading item if possible
'' Inputs:      Auto Trade Item, Confirm Position Close?, Close Position?,
''              Reason, Disabled from code?, From Basket?, Skip Question?
'' Returns:     True if able to disable, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DisableTradeItem(TradeItem As cAutoTradeItem, Optional ByVal bConfirmPositionClose As Boolean = True, Optional ByVal bClosePosition As Boolean = True, Optional ByVal strReason As String = "", Optional ByVal bDisabledFromCode As Boolean = False, Optional ByVal bFromBasket As Boolean = False, Optional bSkipQuestion As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim bDisable As Boolean             ' Do the user still want to disable?
    
    bReturn = False
    If Not TradeItem Is Nothing Then
        If TradeItem.ParentID > -1& Then
            If TradeItem.Active = True Then
                bDisable = True
                If bSkipQuestion = False Then
                    If g.Broker.HasWorkingOrdersForAutoTrade(TradeItem) Then
                        If InfBox("Turning off '" & TradeItem.Name & "' will cancel any open orders that have been generated.||Do you want to continue?|", "?", "+Turn Off|-Leave On", "Warning") = "L" Then
                            bDisable = False
                        End If
                    End If
                End If
                
                If bDisable Then
                    TradeItem.Disable bConfirmPositionClose, bClosePosition, strReason, bDisabledFromCode
                    'NumActive = NumActive - 1&
                    bReturn = True
                End If
            Else
                bReturn = True
            End If
            
            If (bReturn = True) And (bFromBasket = False) And (TradeItem.ParentID > 0&) Then
                CheckParent TradeItem.ParentID
            End If
        End If
    End If
    
    DisableTradeItem = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.DisableTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParentHasActiveChildren
'' Description: Does the given parent have active children?
'' Inputs:      Parent ID
'' Returns:     True if any child is working, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ParentHasActiveChildren(ByVal lParentID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim ChildItem As cAutoTradeItem     ' Child item for the given parent
    
    bReturn = False
    If lParentID > 0& Then
        For lIndex = 1 To Count
            Set ChildItem = Item(lIndex)
            If ChildItem.ParentID = lParentID Then
                If ChildItem.Active = True Then
                    bReturn = True
                    Exit For
                End If
            End If
        Next lIndex
    End If
    
    ParentHasActiveChildren = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.ParentHasActiveChildren"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoFillCheck
'' Description: Send the given fill to the appropriate automated trading item
'' Inputs:      Fill, Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoFillCheck(ByVal Fill As cPtFill, ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lAutoTradeItemID As Long        ' Automated trading item ID
    Dim strKey As String                ' Key into the orders collection

    lAutoTradeItemID = Fill.AutoTradingItemID
    If Exists(Str(lAutoTradeItemID)) Then
        Item(Str(lAutoTradeItemID)).FillCheck Fill, Order
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.DoFillCheck"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StrategyBasketExists
'' Description: Determine if the given strategy basket exists in the database
'' Inputs:      Strategy Basket ID
'' Returns:     True if exists in the database, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function StrategyBasketExists(ByVal lStrategyBasketID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    bReturn = True
    If lStrategyBasketID > 0 Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] WHERE StrategyBasketID=" & Str(lStrategyBasketID) & ";", dbOpenDynaset)
        bReturn = Not (rs.BOF And rs.EOF)
    End If

    StrategyBasketExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAutoTradeItems.StrategyBasketExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.AutoTradeItems = New cGdTree
    Set m.astrLastActive = New cGdArray
    m.astrLastActive.Create eGDARRAY_Strings
    
    Set m.alNewBarChecks = New cGdArray
    m.alNewBarChecks.Create eGDARRAY_Longs
    
    Set m.alOnCloseChecks = New cGdArray
    m.alOnCloseChecks.Create eGDARRAY_Longs
    
    Set m.alCurrentBarChecks = New cGdArray
    m.alCurrentBarChecks.Create eGDARRAY_Longs
    
    Set m.Fills = New cGdTree
    Set m.Orders = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.AutoTradeItems = Nothing
    Set m.astrLastActive = Nothing
    Set m.alNewBarChecks = Nothing
    Set m.alOnCloseChecks = Nothing
    Set m.alCurrentBarChecks = Nothing
    Set m.Fills = Nothing
    Set m.Orders = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAutoTradeItems.Class_Terminate"
    
End Sub
