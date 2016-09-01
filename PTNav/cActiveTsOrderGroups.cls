VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActiveTsOrderGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cActiveTsOrderGroups.cls
'' Description: Class that handles a collection of active trade sense order groups
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/24/2010   DAJ         Modified display code for the displayed orders
'' 06/30/2010   DAJ         Fix for using TradeSense order groups through PFG,
''                          Option to park all/cancel all at close
'' 07/01/2010   DAJ         Fixes for initial park and cancelling a parked group
'' 07/05/2010   DAJ         Fixed with block errors, open/close for broker
'' 07/09/2010   DAJ         Fixes regarding resubmitting parked group
'' 08/09/2010   DAJ         Added OCO for TSOGs with same symbol/account
'' 08/19/2010   DAJ         Save groups on state change for recovery purposes
'' 08/23/2010   DAJ         New method for deactivating active TradeSense groups
'' 09/15/2010   DAJ         Fix for "With block..." error in CancelAll (#5923/#5890)
'' 09/23/2010   DAJ         Option to not cancel parked items ona CancelAll (#5921)
'' 10/08/2010   DAJ         Added OrderExistsInGroup function
'' 10/20/2010   DAJ         Continuous Loop
'' 10/26/2010   DAJ         Implemented LastModified property
'' 11/04/2010   DAJ         Try to reactivate active groups if drop/reconnect to stream
'' 11/15/2010   DAJ         Only OCO groups in same session, Don't count TSOG position for auto exits
'' 03/17/2011   DAJ         Flatten other group(s) with same symbol if in opposite position
'' 05/18/2011   DAJ         Added custom start/stop time for Market1
'' 08/10/2011   DAJ         Ask user to flatten if cancel a TSOG in a position
'' 10/24/2011   DAJ         When cancelling an item, ensure it still exists before removing
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
'' 10/03/2012   DAJ         Lot size for forex symbols in TradeSense order groups
'' 10/10/2012   DAJ         If activation errors out, need to remove from the collection
'' 11/01/2012   DAJ         Change "flatten other side" message to an alert popup
'' 01/22/2013   DAJ         Added the DeleteForAccount function
'' 02/01/2013   DAJ         Log when a group doesn't get added because it is already there
'' 02/13/2013   DAJ         Made FlattenOpposite private, Change IgnoreIndex to IgnoreKey in
''                          FlattenOpposite and CancelForSymbol
'' 05/21/2013   DAJ         Changed the confirmations when cancelling a TSOG
'' 06/26/2014   DAJ         Roll position on a TradeSense order group
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Groups As cGdTree                   ' Collection of Active Trade Sense Order Groups
    bLoading As Boolean                 ' Are we currently loading the collection?
    dLastModified As Double             ' Last time the collection was modified
    
    strLastActive As String             ' String of last active group information
    dLastParked As Double               ' Time the last active groups were parked
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cActiveTsOrderGroup
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    Dim ActiveGp As cActiveTsOrderGroup ' Item to return
    
    Set ActiveGp = Nothing
    If Exists(vIndexOrKey) Then
        Set ActiveGp = m.Groups(vIndexOrKey)
    End If
    
    Set Item = ActiveGp

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.Item.Get"
    
End Property

Public Property Get LastModified() As Double
    LastModified = m.dLastModified
End Property

Public Function Count() As Long
    Count = m.Groups.Count
End Function

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.Groups.Exists(vIndexOrKey)
End Function

Public Sub UpdateLastModified()
    m.dLastModified = CurrentTime
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Remove
'' Description: Remove the group from the collection
'' Inputs:      Index or Key, Reason
'' Returns:     True if removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Remove(ByVal vIndexOrKey As Variant, Optional ByVal strReason As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim ActiveGroup As cActiveTsOrderGroup

    Set ActiveGroup = m.Groups(vIndexOrKey)
    If FormIsLoaded("frmActiveTsOrderGroups") Then
        frmActiveTsOrderGroups.RemoveGroup ActiveGroup
    End If
    
    bReturn = m.Groups.Remove(vIndexOrKey)
    If bReturn Then
        With ActiveGroup
            If Len(strReason) > 0 Then
                g.Broker.BrokerDebug .Broker, "'" & .tsOrderGroup.Name & "; " & .Symbol & "; " & Str(.AccountID) & "; " & Str(.Quantity) & "' removed from collection because '" & strReason & "'"
            Else
                g.Broker.BrokerDebug .Broker, "'" & .tsOrderGroup.Name & "; " & .Symbol & "; " & Str(.AccountID) & "; " & Str(.Quantity) & "' removed from collection"
            End If
        End With
        
        Save
    End If
    
    UpdateLastModified
    Remove = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.Remove"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load up any saved active Trade Sense order groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Information read in from a file
    Dim lIndex As Long                  ' Index into a for loop

    If astrFile.FromFile(AddSlash(App.Path) & "Custom\ATSOG.sav") Then
        m.bLoading = True
        For lIndex = 0 To astrFile.Size - 1
            AddFromString astrFile(lIndex), False
        Next lIndex
        m.bLoading = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save any saved active Trade Sense order groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Information read in from a file
    Dim lIndex As Long                  ' Index into a for loop
    
    astrFile.Create eGDARRAY_Strings
    
    For lIndex = 1 To m.Groups.Count
        astrFile.Add Item(lIndex).ToString
    Next lIndex
    
    astrFile.ToFile AddSlash(App.Path) & "Custom\ATSOG.sav"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitGroup
'' Description: Attempt to submit a group with the given information
'' Inputs:      Trade Sense Order Group, Symbol, Account, Quantity,
''              Loop?, Loop Expiration, Custom Start Time, Custom Stop Time,
''              Lot Size
'' Returns:     True if submitted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SubmitGroup(ByVal tsoGrp As cTradeSenseOrderGroup, ByVal strSymbol As String, ByVal lAccountID As Long, ByVal lQuantity As Long, ByVal bLoop As Boolean, ByVal dLoopExp As Double, Optional ByVal dCustomStartTime As Double = kNullData, Optional ByVal dCustomStopTime As Double = kNullData, Optional ByVal lLotSize As Long = 1&) As Boolean
On Error GoTo ErrSection:

    SubmitGroup = (AddFromInfo(tsoGrp, strSymbol, lAccountID, lQuantity, True, bLoop, dLoopExp, , , dCustomStartTime, dCustomStopTime, lLotSize) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.SubmitGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Submit
'' Description: Attempt to submit the given active TradeSense group
'' Inputs:      Active TradeSense Order Group
'' Returns:     True if submitted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Submit(ByVal tsoGroup As cActiveTsOrderGroup)
On Error GoTo ErrSection:

    Submit = (AddToCollection(tsoGroup, True) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.Submit"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkGroup
'' Description: Attempt to park a group with the given information
'' Inputs:      Trade Sense Order Group, Symbol, Account, Quantity,
''              Loop?, Loop Expiration, Custom Start Time, Custom Stop Time,
''              Lot Size
'' Returns:     True if park, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParkGroup(ByVal tsoGrp As cTradeSenseOrderGroup, ByVal strSymbol As String, ByVal lAccountID As Long, ByVal lQuantity As Long, ByVal bLoop As Boolean, ByVal dLoopExp As Double, Optional ByVal dCustomStartTime As Double = kNullData, Optional ByVal dCustomStopTime As Double = kNullData, Optional ByVal lLotSize As Long = 1&) As Boolean
On Error GoTo ErrSection:

    ParkGroup = (AddFromInfo(tsoGrp, strSymbol, lAccountID, lQuantity, False, bLoop, dLoopExp, , , dCustomStartTime, dCustomStopTime, lLotSize) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.ParkGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Park
'' Description: Attempt to park the given active TradeSense group
'' Inputs:      Active TradeSense Order Group
'' Returns:     True if submitted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Park(ByVal tsoGroup As cActiveTsOrderGroup)
On Error GoTo ErrSection:

    Park = (AddToCollection(tsoGroup, False) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.Park"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Tell each active group to update their bars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBars()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim grp As cActiveTsOrderGroup      ' Order group
    
    If Len(m.strLastActive) > 0 Then
        SubmitPreviouslyActive
    End If
    
    For lIndex = m.Groups.Count To 1 Step -1
        Set grp = Item(lIndex)
        If Not grp Is Nothing Then
            If grp.ActivationError Then
                Remove lIndex, "activation error"
            Else
                If grp.AllClosed Then
                    If grp.DoReactivate Then
                        grp.Activate True
                    Else
                        Remove lIndex, "All Closed in UpdateBars"
                    End If
                ElseIf grp.Submitted = True Then
                    grp.UpdateBars
                End If
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.UpdateBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Refresh the given order in the appropriate group
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshOrder(ByVal Order As cPtOrder, Optional ByVal lOldOrderID As Long = -1&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim grp As cActiveTsOrderGroup      ' Order group
    
    For lIndex = 1 To m.Groups.Count
        Set grp = Item(lIndex)
        If Not grp Is Nothing Then
            If grp.RefreshOrder(Order, lOldOrderID) Then
                If grp.AllClosed Then
                    If grp.DoReactivate Then
                        grp.Activate True
                    Else
                        Remove lIndex, "All Closed in RefreshOrder"
                    End If
                Else
                    If grp.AllParked Then
                        grp.Submitted = False
                    End If
                    
                    Save
                End If
                
                Exit For
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.RefreshOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCallback
'' Description: Handle the fill for one of the orders in the appropriate group
'' Inputs:      Fill, Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillCallback(ByVal Fill As cPtFill, ByVal FillOrder As cPtOrder)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bEnteredPosition As Boolean     ' Did the fill enter the order group into a position?
    
    For lIndex = 1 To m.Groups.Count
        If Item(lIndex).FillCallback(Fill, FillOrder, bEnteredPosition) Then
            If (bEnteredPosition = True) And (Item(lIndex).FlattenOpposite = True) Then
                FlattenOpposite Item(lIndex).SymbolOrSymbolID, Item(lIndex).AccountID, Item(lIndex).Position, "TradeSense order group '" & Item(lIndex).tsOrderGroup.Name & "' entered a position in the opposite direction", Item(lIndex).Key
            End If
            If CancelForSymbol(Fill.SymbolOrSymbolID, Fill.AccountID, "Group OCO", Item(lIndex).Key, , Item(lIndex).SessionActivated) Then
                frmOnlineBroker.AddDialogMessage "All other TradeSense order groups for '" & Fill.Symbol & "' in account '" & g.Broker.AccountNameForID(Fill.AccountID) & "' have been cancelled because a fill was received for one", "i", , "TradeSense Order Groups"
            End If
            
            Save
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.FillCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelSubmittedGroup
'' Description: Cancel the given group if it exists
'' Inputs:      Group, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelSubmittedGroup(ByVal tsoGroup As cActiveTsOrderGroup, ByVal strReason As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFlatten As Boolean             ' Flatten the position?
    Dim strText As String               ' Text to display in the confirmation box
    Dim strReturn As String             ' Return from the InfBox
    
    bFlatten = False
    For lIndex = 1 To m.Groups.Count
        If Item(lIndex).Key = tsoGroup.Key Then
            Set tsoGroup = Item(lIndex)
            
            strText = "You are about to close the|TradeSense order group||'" & tsoGroup.tsOrderGroup.Name & "'|for '" & tsoGroup.Symbol & "' in account '" & g.Broker.AccountNameForID(tsoGroup.AccountID) & "'"
            
            If Item(lIndex).Position <> 0 Then
                'bFlatten = (InfBox("Do you want to flatten the position created by '" & Item(lIndex).tsOrderGroup.Name & "'?", "?", "+Yes|-No", "Flatten Position") = "Y")
                strReturn = InfBox(strText & "||Do you want to close the group and flatten your position, just close the group, or do nothing?|", "?", "+Flatten/Close|Close Only|-Do Nothing", "TradeSense Order Group Confirmation")
            Else
                strReturn = InfBox(strText & "||Do you want to close the group or do nothing?|", "?", "+Close Group|-Do Nothing", "TradeSense Order Group Confirmation")
            End If
            
            ' If the user chose to cancel the item and the item still exists, close it...
            If strReturn <> "D" Then
                If lIndex <= m.Groups.Count Then
                    If Item(lIndex).Key = tsoGroup.Key Then
                        CancelItem lIndex, strReason, (strReturn = "F")
                        Save
                    End If
                End If
            End If
            
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.CancelSubmittedGroup"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkSubmittedGroup
'' Description: Park the given group if it exists
'' Inputs:      Group, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkSubmittedGroup(ByVal tsoGroup As cActiveTsOrderGroup, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Groups.Count
        If Item(lIndex).Key = tsoGroup.Key Then
            ParkItem lIndex, strReason
            Save
            
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.ParkSubmittedGroup"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitParkedGroup
'' Description: Submit the given group if it exists
'' Inputs:      Group
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SubmitParkedGroup(ByVal tsoGroup As cActiveTsOrderGroup)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim grp As cActiveTsOrderGroup      ' Order group
    
    For lIndex = 1 To m.Groups.Count
        Set grp = Item(lIndex)
        If Not grp Is Nothing Then
            If grp.Key = tsoGroup.Key Then
                grp.Activate
                Save
                
                Exit For
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.SubmitParkedGroup"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelForSymbol
'' Description: Cancel all groups with the given symbol and account
'' Inputs:      Symbol, Account, Reason, Ignore Index, Ignore Parked?
'' Returns:     True if items cancelled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelForSymbol(ByVal vSymbolOrSymbolID As Variant, ByVal lAccountID As Long, ByVal strReason As String, Optional ByVal strIgnoreKey As String = "", Optional ByVal bIgnoreParked As Boolean = True, Optional ByVal lOnlyInSession = kNullData) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID
    Dim strSymbol As String             ' Symbol
    
    bReturn = False
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    If lSymbolID = 0 Then
        vSymbolOrSymbolID = strSymbol
    Else
        vSymbolOrSymbolID = lSymbolID
    End If
    
    For lIndex = m.Groups.Count To 1 Step -1
        If (ConvertToTradeSymbol(Item(lIndex).SymbolOrSymbolID, CurrentTime("", "", True)) = vSymbolOrSymbolID) Then
            If (Item(lIndex).AccountID = lAccountID) Then
                If Item(lIndex).Key <> strIgnoreKey Then
                    If ((Item(lIndex).Submitted = True) And (Item(lIndex).SessionActivated = lOnlyInSession)) Or (bIgnoreParked = False) Then
                        CancelItem lIndex, strReason
                        bReturn = True
                    End If
                End If
            End If
        End If
    Next lIndex
    
    CancelForSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.CancelGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingGroups
'' Description: Determine if there are working (non-parked) groups
'' Inputs:      Broker
'' Returns:     True if Working groups, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingGroups(Optional ByVal nBroker As eTT_AccountType = -1&) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    For lIndex = 1 To m.Groups.Count
        If Item(lIndex).Submitted Then
            If (nBroker = -1&) Or (nBroker = Item(lIndex).Broker) Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    HasWorkingGroups = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.HasWorkingGroups"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingGroupsForSymbol
'' Description: Determine if there are working groups for the given info
'' Inputs:      Symbol, Account
'' Returns:     True if Working groups, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingGroupsForSymbol(ByVal vSymbolOrSymbolID As Variant, ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID
    Dim strSymbol As String             ' Symbol
    Dim bReturn As Boolean              ' Return value for the function
    
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    If lSymbolID = 0 Then
        vSymbolOrSymbolID = strSymbol
    Else
        vSymbolOrSymbolID = lSymbolID
    End If
    
    bReturn = False
    For lIndex = 1 To m.Groups.Count
        If (Item(lIndex).SymbolOrSymbolID = vSymbolOrSymbolID) And (Item(lIndex).AccountID = lAccountID) Then
            If Item(lIndex).Submitted Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    HasWorkingGroupsForSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.HasWorkingGroupsForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasPendingOrdersForSymbol
'' Description: Determine if there are pending orders for the given info
'' Inputs:      Symbol, Account
'' Returns:     True if pending orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasPendingOrdersForSymbol(ByVal vSymbolOrSymbolID As Variant, ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID
    Dim strSymbol As String             ' Symbol
    Dim bReturn As Boolean              ' Return value for the function
    
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    If lSymbolID = 0 Then
        vSymbolOrSymbolID = strSymbol
    Else
        vSymbolOrSymbolID = lSymbolID
    End If
    
    bReturn = False
    For lIndex = 1 To m.Groups.Count
        If (Item(lIndex).SymbolOrSymbolID = vSymbolOrSymbolID) And (Item(lIndex).AccountID = lAccountID) Then
            If Item(lIndex).Submitted Then
                If Item(lIndex).HasPendingOrders Then
                    bReturn = True
                    Exit For
                End If
            End If
        End If
    Next lIndex
    
    HasPendingOrdersForSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.HasPendingOrdersForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAllActive
'' Description: Cancel all active TradeSense order groups
'' Inputs:      Reason, Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelAllActive(ByVal strReason As String, Optional ByVal nBroker As eTT_AccountType = -1&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = Count To 1 Step -1
        If Item(lIndex).Submitted Then
            If (nBroker = -1&) Or (nBroker = Item(lIndex).Broker) Then
                CancelItem lIndex, strReason
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.CancelAllActive"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkAllActive
'' Description: Park all active TradeSense order groups
'' Inputs:      Broker, Reason
'' Returns:     Items Parked
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParkAllActive(Optional ByVal nBroker As eTT_AccountType = -1&, Optional ByVal strReason As String = "") As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrParked As New cGdArray      ' Array of items that were parked
    
    astrParked.Create eGDARRAY_Strings
    For lIndex = Count To 1 Step -1
        If Item(lIndex).Submitted Then
            If (nBroker = -1&) Or (nBroker = Item(lIndex).Broker) Then
                astrParked.Add Item(lIndex).Key
                ParkItem lIndex, strReason
            End If
        End If
    Next lIndex
    
    ParkAllActive = astrParked.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.ParkAllActive"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeactivateGroups
'' Description: Deactivate all working groups
'' Inputs:      Reason, Wait until Done?, Message Timeout
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeactivateGroups(Optional ByVal strReason As String = "", Optional ByVal bWaitUntilDone As Boolean = True, Optional ByVal lMsgTimeout As Long = 0&)
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout for waiting until done
    Dim strInfo As String               ' Saved order group information

    If HasWorkingGroups Then
        If InfBox("Would you like to park or cancel your active TradeSense order groups?", "?", "+Park|-Cancel", "Deactivate TradeSense Order Groups", , lMsgTimeout) = "P" Then
            strInfo = ParkAllActive(, "Deactivating groups because " & strReason)
            If g.RealTime.Reconnecting Then
                m.strLastActive = strInfo
                m.dLastParked = CurrentTime
            End If
        Else
            CancelAllActive strReason
        End If
        
        If bWaitUntilDone Then
            lTimeOut = 0&
            Do While (HasWorkingGroups = True) And (lTimeOut < 30&)
                Sleep 1, False, True
                lTimeOut = lTimeOut + 1&
            Loop
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.DeactivateGroups"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderExistsInGroup
'' Description: Does this order exist in an active TradeSense order group?
'' Inputs:      Order
'' Returns:     True if group has order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderExistsInGroup(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop

    bReturn = False
    For lIndex = 1 To m.Groups.Count
        If Item(lIndex).HasOrder(Order) Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    OrderExistsInGroup = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.OrderExistsInGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TotalPosition
'' Description: Calculate the total group position for the given account/symbol
'' Inputs:      Account, Symbol
'' Returns:     Total Position of Groups
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TotalPosition(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim dCurrentTime As Double          ' Current time
    
    lReturn = 0&
    For lIndex = 1 To m.Groups.Count
        If Item(lIndex).AccountID = lAccountID Then
            dCurrentTime = CurrentTime("", GetSymbol(vSymbolOrSymbolID), True)
            If ConvertToTradeSymbol(Item(lIndex).SymbolOrSymbolID, dCurrentTime) = ConvertToTradeSymbol(vSymbolOrSymbolID, dCurrentTime) Then
                lReturn = lReturn + Item(lIndex).Position
            End If
        End If
    Next lIndex
    
    TotalPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.TotalPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelRequested
'' Description: The given order is being cancelled some other place in Trade Navigator
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelRequested(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    For lIndex = 1 To m.Groups.Count
        If Item(lIndex).HasOrder(Order) Then
            Item(lIndex).CancelRequested Order
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.CancelRequested"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteForAccount
'' Description: Delete any TradeSense order groups for the given account
'' Inputs:      Account, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteForAccount(ByVal vAccountNumberOrID As Variant, ByVal strReason As String)
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID for the given account
    Dim lIndex As Long                  ' Index into a for loop
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    For lIndex = Count To 1 Step -1
        If Item(lIndex).AccountID = lAccountID Then
            Remove lIndex, strReason
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.DeleteForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReloadData
'' Description: Reload the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReloadData()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To Count
        Item(lIndex).ReloadData
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.ReloadData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitPreviouslyActive
'' Description: Attempt to submit previously active order groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitPreviouslyActive()
On Error GoTo ErrSection:

    Dim astrKeys As cGdArray            ' Keys of order groups to resubmit
    Dim lIndex As Long                  ' Index into a for loop
    Dim bClear As Boolean               ' Clear information?
    Dim tsoGroup As cActiveTsOrderGroup ' Active TradeSense order group
    Dim strReturn As String             ' Return value from the user question
    Static bAlreadyAsked As Boolean     ' Have we already asked the user?

    bClear = False
    If Len(m.strLastActive) > 0 Then
        If (g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected) And (g.RealTime.Reconnecting = False) Then
            If CurrentTime <= (m.dLastParked + (5# / 1440#)) Then
                If bAlreadyAsked = False Then
                    strReturn = InfBox("Would you like to resubmit groups that were previously active?", "?", "+Yes|-No", "Reactivate TradeSense Order Groups", , 20)
                End If
                
                If (strReturn = "Y") Or (bAlreadyAsked = True) Then
                    Set astrKeys = New cGdArray
                    astrKeys.SplitFields m.strLastActive, "|"
                    
                    For lIndex = astrKeys.Size - 1 To 0 Step -1
                        If m.Groups.Exists(astrKeys(lIndex)) Then
                            Set tsoGroup = m.Groups(astrKeys(lIndex))
                            If tsoGroup.CanActivate Then
                                SubmitParkedGroup tsoGroup
                                astrKeys.Remove lIndex
                            End If
                        End If
                    Next lIndex
                    
                    If astrKeys.Size > 0 Then
                        m.strLastActive = astrKeys.JoinFields("|")
                        bAlreadyAsked = True
                    Else
                        bClear = True
                    End If
                Else
                    bClear = True
                End If
            Else
                bClear = True
            End If
            
            If bClear Then
                m.strLastActive = ""
                m.dLastParked = kNullData
                bAlreadyAsked = False
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.SubmitPreviouslyActive"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFromInfo
'' Description: Add the given Active Trade Sense Order Group to the collection
'' Inputs:      Trade Sense Order Group, Symbol, Account, Quantity, Submit,
''              Loop?, Loop Expiration, Order Status, Display Order Status,
''              Custom Start Time, Custom Stop Time, Lot Size
'' Returns:     Index
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddFromInfo(ByVal tsoGrp As cTradeSenseOrderGroup, ByVal strSymbol As String, ByVal lAccountID As Long, ByVal lQuantity As Long, ByVal bSubmit As Boolean, ByVal bLoop As Boolean, ByVal dLoopExp As Double, Optional ByVal strOrderStatus As String = "", Optional ByVal strDisplayOrderStatus As String = "", Optional ByVal dCustomStartTime As Double = kNullData, Optional ByVal dCustomStopTime As Double = kNullData, Optional ByVal lLotSize As Long = 1&) As Long
On Error GoTo ErrSection:

    Dim ActiveGp As cActiveTsOrderGroup ' Active Trade Sense order group
    
    Set ActiveGp = New cActiveTsOrderGroup
    ActiveGp.tsOrderGroup = tsoGrp
    ActiveGp.SymbolOrSymbolID = strSymbol
    ActiveGp.AccountID = lAccountID
    ActiveGp.Quantity = lQuantity
    ActiveGp.Submitted = bSubmit
    ActiveGp.AutoReactivate = bLoop
    ActiveGp.AutoExpiration = dLoopExp
    ActiveGp.CustomStartTime = dCustomStartTime
    ActiveGp.CustomStopTime = dCustomStopTime
    ActiveGp.LotSize = lLotSize
    
    If Len(strOrderStatus) > 0 Then
        ActiveGp.StatusForOrders = strOrderStatus
    End If
    If Len(strDisplayOrderStatus) > 0 Then
        ActiveGp.DisplayStatusForOrders = strDisplayOrderStatus
    End If
    
    AddFromInfo = AddToCollection(ActiveGp, bSubmit)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.AddFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFromString
'' Description: Add the given Active Trade Sense Order Group to the collection
'' Inputs:      Information String, Submit?
'' Returns:     Index
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddFromString(ByVal strInfo As String, ByVal bSubmit As Boolean) As Long
On Error GoTo ErrSection:

    Dim ActiveGp As cActiveTsOrderGroup ' Active Trade Sense order group
    
    Set ActiveGp = New cActiveTsOrderGroup
    ActiveGp.FromString strInfo
    
    AddFromString = AddToCollection(ActiveGp, bSubmit)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.AddFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddToCollection
'' Description: Add the given Active Trade Sense Order Group to the collection
'' Inputs:      Active Trade Sense Order Group, Submit?
'' Returns:     Index
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddToCollection(ActiveGroup As cActiveTsOrderGroup, ByVal bSubmit As Boolean) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

    lReturn = -1&
    If Exists(ActiveGroup.Key) = False Then
        With ActiveGroup
            If bSubmit Then
                g.Broker.BrokerDebug .Broker, "'" & .tsOrderGroup.Name & "; " & .Symbol & "; " & Str(.AccountID) & "; " & Str(.Quantity) & "' added to collection for submit"
            Else
                g.Broker.BrokerDebug .Broker, "'" & .tsOrderGroup.Name & "; " & .Symbol & "; " & Str(.AccountID) & "; " & Str(.Quantity) & "' added to collection for park"
            End If
        End With
        
        lReturn = m.Groups.Add(ActiveGroup, ActiveGroup.Key)
        UpdateLastModified
        
        m.Groups(lReturn).ID = lReturn
        
        If bSubmit Then
            If Not m.Groups(lReturn).Activate Then
                m.Groups(lReturn).ParkAll
            End If
        Else
            m.Groups(lReturn).ParkAll
        End If
        
        If m.bLoading = False Then
            Save
        End If
    Else
        g.Broker.BrokerDebug ActiveGroup.Broker, "Could not activate '" & ActiveGroup.tsOrderGroup.Name & "' on '" & ActiveGroup.Symbol & "' in account '" & g.Broker.AccountNameForID(ActiveGroup.AccountID) & "' because it is already submitted or parked"
    End If
    
    AddToCollection = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.AddToCollection"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelItem
'' Description: Cancel the given item
'' Inputs:      Index, Reason, Flatten?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelItem(ByVal lIndex As Long, ByVal strReason As String, Optional ByVal bFlatten As Boolean = False)
On Error GoTo ErrSection:

    If Item(lIndex).Submitted Then
        If bFlatten Then
            g.Broker.BrokerDebug Item(lIndex).Broker, "Flattening Active TradeSense Order Group: " & strReason
            Item(lIndex).Flatten
        Else
            g.Broker.BrokerDebug Item(lIndex).Broker, "Cancelling Active TradeSense Order Group: " & strReason
            Item(lIndex).CancelAll strReason
        End If
        
        If Exists(lIndex) Then
            If Item(lIndex).AllClosed Then
                Remove lIndex, "Cancel Item"
            End If
        End If
    Else
        g.Broker.BrokerDebug Item(lIndex).Broker, "Removing Parked TradeSense Order Group: " & strReason
        If Item(lIndex).AllParked Then
            Remove lIndex, "Cancel Item"
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.CancelItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkItem
'' Description: Park the given item
'' Inputs:      Index, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParkItem(ByVal lIndex As Long, Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    If Item(lIndex).Submitted Then
        Item(lIndex).ParkAll strReason
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.ParkItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenOpposite
'' Description: Flatten order groups for the given symbol/account that are in
''              the opposite position
'' Inputs:      Symbol, Account, Position, Reason, Ignore Index, Ignore Parked?
'' Returns:     True if any items flattened, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FlattenOpposite(ByVal vSymbolOrSymbolID As Variant, ByVal lAccountID As Long, ByVal lPosition As Long, ByVal strReason As String, Optional ByVal strIgnoreKey As String = "", Optional ByVal bIgnoreParked As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID
    Dim strSymbol As String             ' Symbol
    
    g.Broker.BrokerDebug g.Broker.AccountTypeForID(lAccountID), "FlattenOpposite(" & Str(vSymbolOrSymbolID) & ", " & Str(lAccountID) & ", " & Str(lPosition) & ", '" & strReason & "', '" & strIgnoreKey & "', " & Str(bIgnoreParked) & ")"
    
    bReturn = False
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    If lSymbolID = 0 Then
        vSymbolOrSymbolID = strSymbol
    Else
        vSymbolOrSymbolID = lSymbolID
    End If
    vSymbolOrSymbolID = ConvertToTradeSymbol(vSymbolOrSymbolID, CurrentTime("", "", True))
    
    For lIndex = m.Groups.Count To 1 Step -1
        If (ConvertToTradeSymbol(Item(lIndex).SymbolOrSymbolID, CurrentTime("", "", True)) = vSymbolOrSymbolID) Then
            If (Item(lIndex).AccountID = lAccountID) Then
                If Item(lIndex).Key <> strIgnoreKey Then
                    If (Item(lIndex).Submitted = True) Or (bIgnoreParked = False) Then
                        If (Item(lIndex).Position <> 0) And (Sgn(Item(lIndex).Position) <> Sgn(lPosition)) Then
                            'frmOnlineBroker.AddDialogMessage "The '" & Item(lIndex).tsOrderGroup.Name & "' TradeSense order group for '" & Item(lIndex).Symbol & "' in account '" & g.Broker.AccountNameForID(lAccountID) & "' is being flattened because another TradeSense order group has entered into an opposite position for the same symbol", "i", , "TradeSense Order Groups"
                            ShowAlertPopup "The '" & Item(lIndex).tsOrderGroup.Name & "' TradeSense order group for '" & Item(lIndex).Symbol & "' in account '" & g.Broker.AccountNameForID(lAccountID) & "' is being flattened because another TradeSense order group has entered into an opposite position for the same symbol", "TradeSense Order Groups", vbCenter
                            
                            Item(lIndex).Flatten
                            bReturn = True
                        End If
                    End If
                End If
            End If
        End If
    Next lIndex
    
    FlattenOpposite = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.FlattenOpposite"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Groups = New cGdTree
    m.bLoading = False
    Load
    
    m.strLastActive = ""
    m.dLastParked = kNullData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.Class_Initialize"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Save
    Set m.Groups = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroups.Class_Terminate"

End Sub
