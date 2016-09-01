VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerInfo.cls
'' Description: Class to hold broker information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/06/2009   DAJ         Display "Mismatch" for position if in a mismatch
'' 03/27/2009   DAJ         Pay attention to symbols in order and fill refreshes
'' 06/09/2009   DAJ         Added ability to set connection status
'' 08/21/2009   DAJ         Make chanes for Parking/Submitting/Cancelling OCOs
'' 09/01/2009   DAJ         Use new Parked order status
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 12/21/2009   DAJ         Optionally don't count parked quick fills as working
'' 12/22/2009   DAJ         Fix for conditional AND trigger by orders
'' 01/04/2009   DAJ         Added IncludeAutoExitOrders flag to HasWorkingQuickFill
'' 03/11/2010   DAJ         Use global collections
'' 09/13/2010   DAJ         Added code for Rithmic
'' 10/26/2010   DAJ         Update LastModified time for positions/accounts when change
'' 03/07/2011   DAJ         Changed Option Nav calls, Added AddPositionFromObject
'' 03/10/2011   DAJ         Added OrderForAccessID call
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 07/15/2011   DAJ         Return ID from CreateFillSummaryForAutoTrade
'' 07/21/2011   DAJ         Further tweaks for automated trading
'' 07/22/2011   DAJ         Only create fill summaries for this broker
'' 07/29/2011   DAJ         Enhanced logging in AddOrder
'' 08/02/2011   DAJ         In HasWorkingOrders, optionally include auto trade orders
'' 08/15/2011   DAJ         Fixed DateIsSnapshot when no orders or fills
'' 10/20/2011   DAJ         Dump BrokerInfo to log
'' 10/21/2011   DAJ         If previous auto trade id passed into AddFill, recalc that position
'' 10/21/2011   DAJ         Expose LoadCollections
'' 11/11/2011   DAJ         Fix for sending connection status for account to Option Nav
'' 12/09/2011   DAJ         Added Expiration Date to Account Positions
'' 05/15/2012   DAJ         Added RenameOptionNavGroup
'' 08/03/2012   DAJ         Remove Alaron
'' 09/26/2012   DAJ         Added the ChangeGenesisIdForOrder routine
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 03/22/2013   DAJ         Added the OpenPositionsForAccount call
'' 04/16/2013   DAJ         Added the RefreshOrder call
'' 04/17/2013   DAJ         Added direction filter to OrdersForSymbol call
'' 05/01/2013   DAJ         Shadow Trading
'' 07/30/2013   DAJ         Data Pending order status for a conditional order
'' 09/11/2013   DAJ         Added logging for when the LastModified changes for AccountPositions
'' 02/05/2014   DAJ         Fixes for deleting an account
'' 05/19/2014   DAJ         Log earliest session date before removing orders or fills
'' 05/28/2014   DAJ         Don't include null open profit when calculating account open equity
'' 10/24/2014   DAJ         Fill Display
'' 11/07/2014   DAJ         Added UsePreviousBrokerId parameter to AddOrder function
'' 12/10/2014   DAJ         Determine snapshot/history line per account
'' 01/07/2015   DAJ         Added the ClearRefreshFlags function
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kPendingOrderTimeout = 30#

Private Type mPrivate
    nBroker As eTT_AccountType          ' Broker
    strBrokerPath As String             ' Broker path

    dLastChanged As Double              ' Tick count when the last thing changed
    astrLastChanged As cGdArray         ' Tick count when account/symbol changed
    
    astrConnectedAccounts As cGdArray   ' Array of connected accounts from the broker
    astrPositions As cGdArray           ' Array of positions from the broker
    astrOvernight As cGdArray           ' Array of overnight positions/overrides
    astrPositionMatch As cGdArray       ' Array of Account/Symbol/Position Match
    
    EarliestOrderDate As cEarliestDate  ' Earliest order date
    EarliestFillDate As cEarliestDate   ' Earliest fill date
    
    bRefreshingAccounts As Boolean      ' Are we currently refreshing the accounts?
    bRefreshingOrders As Boolean        ' Are we currently refreshing the orders?
    bRefreshingFills As Boolean         ' Are we currently refreshing the fills?
    bRefreshingPositions As Boolean     ' Are we currently refreshing the positions?
    
    Accounts As cPtAccounts             ' Collection of accounts for this broker
    FillSummaries As cAccountPositions  ' Collection of account positions for this broker
    SnapshotFills As cPtFills           ' Collection of snapshot fills for this broker
    SnapshotOrders As cPtOrders         ' Collection of snapshot orders for this broker
    PendingOrders As cGdTree            ' Collection of pending orders for this broker
    
    astrOrderSymbols As cGdArray        ' Array of symbols that came through with orders
    bOrderRefreshHasOnlyTF As Boolean   ' Does the order refresh have only TF in it?
    astrFillSymbols As cGdArray         ' Array of symbols that came through with fills
    bFillRefreshHasOnlyTF As Boolean    ' Does the fill refresh have only TF in it?
End Type
Private m As mPrivate

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property
Public Property Let Broker(ByVal nBroker As eTT_AccountType)
    m.nBroker = nBroker
End Property

Public Property Get BrokerPath() As String
    BrokerPath = m.strBrokerPath
End Property
Public Property Let BrokerPath(ByVal strBrokerPath As String)
    m.strBrokerPath = strBrokerPath
End Property

Public Property Get LastChanged() As Double
    LastChanged = m.dLastChanged
End Property
Public Property Get ConnectedAccounts() As cGdArray
    Set ConnectedAccounts = m.astrConnectedAccounts
End Property

Public Property Get Accounts() As cPtAccounts
    Set Accounts = m.Accounts
End Property
Public Property Get Orders() As cPtOrders
    Set Orders = m.SnapshotOrders
End Property
Public Property Get Fills() As cPtFills
    Set Fills = m.SnapshotFills
End Property
Public Property Get Positions() As cGdArray
    Set Positions = m.astrPositions
End Property
Public Property Get FillSummary() As cAccountPositions
    Set FillSummary = m.FillSummaries
End Property
Public Property Get PendingOrders() As cGdTree
    Set PendingOrders = m.PendingOrders
End Property

Public Property Get OrderRefreshHasOnlyTF() As Boolean
    OrderRefreshHasOnlyTF = m.bOrderRefreshHasOnlyTF
End Property
Public Property Get FillRefreshHasOnlyTF() As Boolean
    FillRefreshHasOnlyTF = m.bFillRefreshHasOnlyTF
End Property

Public Property Get Refreshing() As Boolean
    Refreshing = m.bRefreshingAccounts Or m.bRefreshingOrders Or m.bRefreshingFills Or m.bRefreshingPositions
End Property

Public Property Get RefreshingAccounts() As Boolean
    RefreshingAccounts = m.bRefreshingAccounts
End Property
Public Property Let RefreshingAccounts(ByVal bRefreshingAccounts As Boolean)
    m.bRefreshingAccounts = bRefreshingAccounts
End Property

Public Property Get RefreshingOrders() As Boolean
    RefreshingOrders = m.bRefreshingOrders
End Property
Public Property Let RefreshingOrders(ByVal bRefreshingOrders As Boolean)
On Error GoTo ErrSection:

    If (bRefreshingOrders = True) And (m.bRefreshingOrders = False) Then
        ClearOrders
        m.astrOrderSymbols.Clear
        m.bOrderRefreshHasOnlyTF = True
    ElseIf (bRefreshingOrders = False) And (m.bRefreshingOrders = True) Then
        If m.astrOrderSymbols.Size = 0 Then
            m.bOrderRefreshHasOnlyTF = False
        End If
        
        m.EarliestOrderDate.DumpDebug m.nBroker, "BInfo: RemoveSnapshotOrders before"
        m.SnapshotOrders.RemoveSnapshotOrders m.EarliestOrderDate, m.nBroker
    End If
    
    m.bRefreshingOrders = bRefreshingOrders
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cBrokerInfo.RefreshingOrders"
    
End Property

Public Property Get RefreshingFills() As Boolean
    RefreshingFills = m.bRefreshingFills
End Property
Public Property Let RefreshingFills(ByVal bRefreshingFills As Boolean)
On Error GoTo ErrSection:

    If (bRefreshingFills = True) And (m.bRefreshingFills = False) Then
        ClearFills
        m.astrFillSymbols.Clear
        m.bFillRefreshHasOnlyTF = True
    ElseIf (m.bRefreshingFills = True) And (bRefreshingFills = False) Then
        If m.astrFillSymbols.Size = 0 Then
            m.bFillRefreshHasOnlyTF = False
        End If
        
        m.EarliestFillDate.DumpDebug m.nBroker, "BInfo: RemoveSnapshotFills before"
        If m.SnapshotFills.RemoveSnapshotFills(m.EarliestFillDate, m.astrFillSymbols) Then
            RebuildFillSummaries "", "", True
            If FormIsLoaded("frmTTPositions") Then
                frmTTPositions.Reload m.nBroker
            End If
            If IncludeBrokerToOptionNav(m.nBroker) Then
                GetHistoricalFillsForOptionNav
            End If
        Else
            RebuildFillSummaries "", "", False
        End If
    End If
    m.bRefreshingFills = bRefreshingFills
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cBrokerInfo.RefreshingFills.Let"
    
End Property

Public Property Get RefreshingPositions() As Boolean
    RefreshingPositions = m.bRefreshingPositions
End Property
Public Property Let RefreshingPositions(ByVal bRefreshingPositions As Boolean)
On Error GoTo ErrSection:

    If (m.bRefreshingPositions = True) And (bRefreshingPositions = False) Then
        RebuildFillSummaries
    End If
    m.bRefreshingPositions = bRefreshingPositions
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.RefreshingPositions.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolLastChanged
'' Description: Determine the last time anything changed for account/symbol pair
'' Inputs:      Account, Symbol
'' Returns:     Tick Count when last changed
''
'' Fields:      Account, Symbol, Last Changed
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get SymbolLastChanged(ByVal strAccount As String, ByVal strSymbol As String) As Double
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim dReturn As Double               ' Value to return from the function
    
    dReturn = 0#
    If m.astrLastChanged.BinarySearch(strAccount & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        dReturn = Val(Parse(m.astrLastChanged(lPos), vbTab, 3))
    End If
    
    SymbolLastChanged = dReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.SymbolLastChanged.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolLastChanged
'' Description: Set the last time anything changed for account/symbol pair
'' Inputs:      Account, Symbol, Tick Count when last changed
'' Returns:     None
''
'' Fields:      Account, Symbol, Last Changed
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let SymbolLastChanged(ByVal strAccount As String, ByVal strSymbol As String, ByVal dLastChanged As Double)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    
    If m.astrLastChanged.BinarySearch(strAccount & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        m.astrLastChanged(lPos) = strAccount & vbTab & strSymbol & vbTab & Str(dLastChanged)
    Else
        m.astrLastChanged.Add strAccount & vbTab & strSymbol & vbTab & Str(dLastChanged), lPos
    End If
    m.dLastChanged = dLastChanged

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.SymbolLastChanged.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OvernightPosition
'' Description: Pass back the overnight position for an account and symbol
'' Inputs:      Account, Symbol, At ID
'' Returns:     Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get OvernightPosition(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long) As Long
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    
    If m.astrOvernight.BinarySearch(strAccount & vbTab & strSymbol & vbTab & Str(lAtID) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
        m.astrOvernight.Add strAccount & vbTab & strSymbol & vbTab & Str(lAtID) & vbTab & "0", lPos
    End If
    
    OvernightPosition = CLng(Val(Parse(m.astrOvernight(lPos), vbTab, 4)))

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.OvernightPosition.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OvernightPosition
'' Description: Pass back the overnight position for an account and symbol
'' Inputs:      Account, Symbol, At ID
'' Returns:     Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let OvernightPosition(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, ByVal lOvernightPosition As Long)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim bChanged As Boolean             ' Did the overnight position change?
    
    If m.astrOvernight.BinarySearch(strAccount & vbTab & strSymbol & vbTab & Str(lAtID) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        If CLng(Val(Parse(m.astrOvernight(lPos), vbTab, 4))) <> lOvernightPosition Then
            bChanged = True
            m.astrOvernight(lPos) = strAccount & vbTab & strSymbol & vbTab & Str(lAtID) & vbTab & Str(lOvernightPosition)
        Else
            bChanged = False
        End If
    Else
        bChanged = True
        m.astrOvernight.Add strAccount & vbTab & strSymbol & vbTab & Str(lAtID) & vbTab & Str(lOvernightPosition), lPos
    End If
    
    If bChanged Then
        m.FillSummaries.SetDirtyForSymbol strAccount, strSymbol, lAtID
        m.FillSummaries.UpdateLastModified
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( OvernightPosition )"
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.OvernightPosition.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionMatch
'' Description: Return whether position matches for given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     True if position match, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get PositionMatch(ByVal strAccountNumber As String, ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If m.astrPositionMatch.BinarySearch(strAccountNumber & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        bReturn = ((Parse(m.astrPositionMatch(lPos), vbTab, 3) = "1") And (Parse(m.astrPositionMatch(lPos), vbTab, 4) = "1"))
    End If
    
    PositionMatch = bReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.PositionMatch.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConsistentBroker
'' Description: Return whether broker data consistent for given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     True if broker data consistent, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ConsistentBroker(ByVal strAccountNumber As String, ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If m.astrPositionMatch.BinarySearch(strAccountNumber & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        bReturn = (Parse(m.astrPositionMatch(lPos), vbTab, 3) = "1")
    End If
    
    ConsistentBroker = bReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.ConsistentBroker.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConsistentBroker
'' Description: Set whether broker data consistent for given account and symbol
'' Inputs:      Account, Symbol, Consistent Data?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ConsistentBroker(ByVal strAccountNumber As String, ByVal strSymbol As String, ByVal bConsistentBroker As Boolean)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim strConsistentBroker As String   ' String version of consistent broker

    If bConsistentBroker = True Then strConsistentBroker = "1" Else strConsistentBroker = "0"

    If m.astrPositionMatch.BinarySearch(strAccountNumber & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        m.astrPositionMatch(lPos) = strAccountNumber & vbTab & strSymbol & vbTab & strConsistentBroker & vbTab & Parse(m.astrPositionMatch(lPos), vbTab, 4)
    Else
        m.astrPositionMatch.Add strAccountNumber & vbTab & strSymbol & vbTab & strConsistentBroker & vbTab & "1", lPos
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.ConsistentBroker.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedMatch
'' Description: Return whether carried position matches for given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     True if carried position matches, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get CarriedMatch(ByVal strAccountNumber As String, ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If m.astrPositionMatch.BinarySearch(strAccountNumber & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        bReturn = (Parse(m.astrPositionMatch(lPos), vbTab, 4) = "1")
    End If
    
    CarriedMatch = bReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.CarriedMatch.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedMatch
'' Description: Set whether carried position matches for given account and symbol
'' Inputs:      Account, Symbol, Carried Position Matches?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let CarriedMatch(ByVal strAccountNumber As String, ByVal strSymbol As String, ByVal bCarriedMatch As Boolean)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim strCarriedMatch As String       ' String version of carried position match

    If bCarriedMatch = True Then strCarriedMatch = "1" Else strCarriedMatch = "0"

    If m.astrPositionMatch.BinarySearch(strAccountNumber & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        m.astrPositionMatch(lPos) = strAccountNumber & vbTab & strSymbol & vbTab & Parse(m.astrPositionMatch(lPos), vbTab, 3) & vbTab & strCarriedMatch
    Else
        m.astrPositionMatch.Add strAccountNumber & vbTab & strSymbol & vbTab & "1" & vbTab & strCarriedMatch, lPos
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBrokerInfo.CarriedMatch.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize this broker info object with the given information
'' Inputs:      Broker, Broker Path
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal nBroker As eTT_AccountType, ByVal strBrokerPath As String)
On Error GoTo ErrSection:
    
    m.nBroker = nBroker
    m.strBrokerPath = strBrokerPath
    
    LoadCollections
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCollections
'' Description: Load all of the collections
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadCollections()
On Error GoTo ErrSection:

    m.Accounts.LoadForBroker m.nBroker
    m.FillSummaries.LoadForBroker m.nBroker
    m.SnapshotFills.LoadSnapshotFillsForBroker m.nBroker
    m.SnapshotOrders.LoadSnapshotOrdersForBroker m.nBroker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.LoadCollections"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrdersForSymbol
'' Description: Retrieve the orders for the given account and symbol
'' Inputs:      Account, Symbol, At ID, Only Working Orders?, Direction
'' Returns:     Orders for that account and symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrdersForSymbol(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal bOnlyWorking As Boolean = False, Optional ByVal nDirection As eGDFilterDirection = eGDFilterDirection_All) As cPtOrders
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Orders As cPtOrders             ' Collection of orders to return
    
    Set Orders = New cPtOrders
    For lIndex = 1 To m.SnapshotOrders.Count
        With m.SnapshotOrders(lIndex)
            If .AccountID = g.Broker.AccountIDForNumber(strAccount) Then
                If (.Symbol = strSymbol) And (.AutoTradeItemID = lAtID) Then
                    If (bOnlyWorking = False) Or (IsOpenOrder(.Status) = True) Then
                        If nDirection = eGDFilterDirection_All Then
                            Orders.Add .MakeCopy
                        ElseIf (nDirection = eGDFilterDirection_Longs) And (.Buy = True) Then
                            Orders.Add .MakeCopy
                        ElseIf (nDirection = eGDFilterDirection_Shorts) And (.Buy = False) Then
                            Orders.Add .MakeCopy
                        End If
                    End If
                End If
            End If
        End With
    Next lIndex
    
    Set OrdersForSymbol = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.OrdersForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrdersForAccount
'' Description: Retrieve the orders for the given account
'' Inputs:      Account, Only Working Orders?
'' Returns:     Orders for the account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrdersForAccount(ByVal strAccount As String, Optional ByVal bOnlyWorking As Boolean = False) As cPtOrders
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Orders As cPtOrders             ' Collection of orders to return
    
    Set Orders = New cPtOrders
    For lIndex = 1 To m.SnapshotOrders.Count
        With m.SnapshotOrders(lIndex)
            If .AccountID = g.Broker.AccountIDForNumber(strAccount) Then
                If (bOnlyWorking = False) Or (IsOpenOrder(.Status) = True) Then
                    Orders.Add .MakeCopy
                End If
            End If
        End With
    Next lIndex
    
    Set OrdersForAccount = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.OrdersForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasPendingOrders
'' Description: Are there working orders?
'' Inputs:      None
'' Returns:     True if Working Orders Exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasPendingOrders(Optional ByVal bOnlyQuickFill As Boolean = False) As Boolean
On Error GoTo ErrSection:

    HasPendingOrders = (m.PendingOrders.Count > 0)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.HasPendingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasPendingOrdersForSymbol
'' Description: Are there pending orders for the given account and symbol?
'' Inputs:      Account, Symbol, At ID
'' Returns:     True if Symbols Exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasPendingOrdersForSymbol(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    For lIndex = 1 To m.PendingOrders.Count
        With m.PendingOrders(lIndex)
            If .AccountID = g.Broker.AccountIDForNumber(strAccount) Then
                If (.Symbol = strSymbol) And (.AutoTradeItemID = lAtID) Then
                    bReturn = True
                    Exit For
                End If
            End If
        End With
    Next lIndex
    
    HasPendingOrdersForSymbol = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.HasPendingOrdersForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasOldPendingOrders
'' Description: Check to see if there are old pending orders
'' Inputs:      List of Pending Order IDs, Are any "Sent"?
'' Returns:     True if Old Pending Orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasOldPendingOrders(Optional astrOrders As cGdArray = Nothing, Optional bHasSent As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim dCurrentTime As Double          ' Current date/time in broker time
    Dim lTimeOut As Long                ' Timeout for pending orders
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    bHasSent = False
    
    For lIndex = m.PendingOrders.Count To 1 Step -1
        dCurrentTime = ConvertToBrokerDate(CurrentTime, m.nBroker, m.PendingOrders(lIndex).Symbol)
        lTimeOut = kPendingOrderTimeout
        
        If m.PendingOrders(lIndex).StatusDate < (dCurrentTime - (lTimeOut / 86400#)) Then
            g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "Pending Order Exists: Current Time = " & DateFormat(dCurrentTime, MM_DD_YY, HH_MM_SS, AMPM_UPPER) & "; Status Date = " & DateFormat(m.PendingOrders(lIndex).StatusDate, MM_DD_YY, HH_MM_SS, AMPM_UPPER) & "; Order = " & m.PendingOrders(lIndex).OrderText & " (" & m.PendingOrders(lIndex).GenesisOrderID & ", " & m.PendingOrders(lIndex).BrokerID & ")"
            bReturn = True
            
            If m.PendingOrders(lIndex).Status = eTT_OrderStatus_Sent Then
                bHasSent = True
            End If
            
            If astrOrders Is Nothing Then
                Exit For
            Else
                astrOrders.Add m.PendingOrders(lIndex).BrokerID
            End If
        End If
    Next lIndex
    
    HasOldPendingOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.HasOldPendingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingOrders
'' Description: Are there working orders?
'' Inputs:      Include Parked Orders?, Include Auto Exit Orders?, Include
''              Automated Trading Orders?
'' Returns:     True if Working Orders Exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingOrders(Optional ByVal bIncludeParkedOrders As Boolean = True, Optional ByVal bIncludeAutoExitOrders As Boolean = True, Optional ByVal bIncludeAutoTradeOrders As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    For lIndex = 1 To m.SnapshotOrders.Count
        If (IsOpenOrder(m.SnapshotOrders(lIndex).Status) = True) Then
            If (bIncludeParkedOrders = True) Or (m.SnapshotOrders(lIndex).Status <> eTT_OrderStatus_Parked) Then
                If (bIncludeAutoExitOrders = True) Or (m.SnapshotOrders(lIndex).IsAutoExit = False) Then
                    If (bIncludeAutoTradeOrders = True) Or (m.SnapshotOrders(lIndex).AutoTradeItemID = 0) Then
                        bReturn = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next lIndex
    
    HasWorkingOrders = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.HasWorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingOrdersForSymbol
'' Description: Are there working orders for the given account and symbol?
'' Inputs:      Account, Symbol, At ID, Skip Order ID
'' Returns:     True if Symbols Exist, False otherwise
''
'' Order:       Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date, At ID, Session Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingOrdersForSymbol(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal lSkipOrderID As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    For lIndex = 1 To m.SnapshotOrders.Count
        With m.SnapshotOrders(lIndex)
            If .AccountID = g.Broker.AccountIDForNumber(strAccount) Then
                If (.Symbol = strSymbol) And (.AutoTradeItemID = lAtID) Then
                    If (IsOpenOrder(.Status) = True) And (.OrderID <> lSkipOrderID) Then
                        bReturn = True
                        Exit For
                    End If
                End If
            End If
        End With
    Next lIndex
    
    HasWorkingOrdersForSymbol = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.HasWorkingOrdersForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingOrdersForAutoTrade
'' Description: Are there working orders for the given auto trade item?
'' Inputs:      At ID
'' Returns:     Number of working orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingOrdersForAutoTrade(ByVal lAtID As Long) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value from the function
    
    lReturn = 0&
    For lIndex = 1 To m.SnapshotOrders.Count
        With m.SnapshotOrders(lIndex)
            If .AutoTradeItemID = lAtID Then
                If (IsOpenOrder(.Status) = True) Then
                    lReturn = lReturn + 1&
                End If
            End If
        End With
    Next lIndex
    
    HasWorkingOrdersForAutoTrade = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.HasWorkingOrdersForAutoTrade"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderPriceHitForAutoTrade
'' Description: Has an order price hit for an auto trade item?
'' Inputs:      Auto Trade Item ID, Current Price
'' Returns:     True if Order Price Hit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderPriceHitForAutoTrade(ByVal lAutoTradeItemID As Long, ByVal dPrice As Double) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 1 To m.SnapshotOrders.Count
        If m.SnapshotOrders(lIndex).AutoTradeItemID = lAutoTradeItemID Then
            If m.SnapshotOrders(lIndex).OrderPriceHit(dPrice) Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    OrderPriceHitForAutoTrade = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.OrderPriceHitForAutoTrade"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrimaryOrdersForSymbol
'' Description: Make a collection of the non-triggered open orders for a symbol
'' Inputs:      Symbol, Account ID, Auto Trade ID, Price
'' Returns:     Collection of primary orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrimaryOrdersForSymbol(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal dPrice As Double = kNullData) As cGdTree
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Orders As New cGdTree           ' Orders collection to return
    Dim lAccountID As Long              ' Account ID
    
    lAccountID = g.Broker.AccountIDForNumber(strAccount)
    
    For lIndex = 1 To m.SnapshotOrders.Count
        If m.SnapshotOrders(lIndex).Symbol = strSymbol Then
            If m.SnapshotOrders(lIndex).IsPrimaryOrder Then
                If (m.SnapshotOrders(lIndex).AutoTradeItemID = lAtID) And (m.SnapshotOrders(lIndex).AccountID = lAccountID) Then
                    If dPrice = kNullData Then
                        Orders.Add m.SnapshotOrders(lIndex), Str(m.SnapshotOrders(lIndex).OrderID)
                    ElseIf dPrice = m.SnapshotOrders(lIndex).OrderPrice(True) Then
                        Orders.Add m.SnapshotOrders(lIndex), Str(m.SnapshotOrders(lIndex).OrderID)
                    End If
                End If
            End If
        End If
    Next lIndex
    
    Set PrimaryOrdersForSymbol = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.PrimaryOrdersForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TriggeredOrdersForOrderID
'' Description: Compile a collection of the orders to be triggered by an order ID
'' Inputs:      Order ID
'' Returns:     Collection of triggered orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TriggeredOrdersForOrderID(ByVal lOrderID As Long) As cGdTree
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Orders As New cGdTree           ' Collection to return from the function
    
    For lIndex = 1 To m.SnapshotOrders.Count
        If (m.SnapshotOrders(lIndex).TriggerOrderID = lOrderID) Then ' And (m.SnapshotOrders(lIndex).IsConditional = False) Then
            If (m.SnapshotOrders(lIndex).Status = eTT_OrderStatus_TriggerPending) Or (m.SnapshotOrders(lIndex).Status = eTT_OrderStatus_DataPending) Then
                Orders.Add m.SnapshotOrders(lIndex), Str(m.SnapshotOrders(lIndex).OrderID)
            End If
        End If
    Next lIndex
    
    Set TriggeredOrdersForOrderID = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.TriggeredOrdersForOrderID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrdersForOrderID
'' Description: Compile a collection of the orders to be cancelled by an order
'' Inputs:      Order ID
'' Returns:     Collection of cancel orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrdersForOrderID(ByVal lOrderID As Long) As cGdTree
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Orders As New cGdTree           ' Collection to return from the function
    
    For lIndex = 1 To m.SnapshotOrders.Count
        If (m.SnapshotOrders(lIndex).CancelOrderID = lOrderID) Then
            Orders.Add m.SnapshotOrders(lIndex), Str(m.SnapshotOrders(lIndex).OrderID)
        End If
    Next lIndex
    
    Set CancelOrdersForOrderID = Orders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.CancelOrdersForOrderID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillsForSymbol
'' Description: Retrieve the fills for the given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Fills for that account and symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillsForSymbol(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long) As cPtFills
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fills As cPtFills               ' Fills collection to return
    
    Set Fills = New cPtFills
    For lIndex = 1 To m.SnapshotFills.Count
        With m.SnapshotFills(lIndex)
            If g.Broker.AccountIDForNumber(strAccount) = .AccountID Then
                If (.Symbol = strSymbol) And (.AutoTradingItemID = lAtID) Then
                    Fills.Add .MakeCopy
                End If
            End If
        End With
    Next lIndex
    
    Set FillsForSymbol = Fills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.FillsForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillsForAccount
'' Description: Retrieve the fills for the given account
'' Inputs:      Account
'' Returns:     Fills for the account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillsForAccount(ByVal strAccount As String) As cPtFills
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fills As cPtFills               ' Fills collection to return
    
    Set Fills = New cPtFills
    For lIndex = 1 To m.SnapshotFills.Count
        With m.SnapshotFills(lIndex)
            If g.Broker.AccountIDForNumber(strAccount) = .AccountID Then
                Fills.Add .MakeCopy
            End If
        End With
    Next lIndex
    
    Set FillsForAccount = Fills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.FillsForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionForSymbol
'' Description: Retrieve the position for the given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Position for that account and symbol
''
'' Fields:      Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionForSymbol(ByVal strAccount As String, ByVal strSymbol As String) As String
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim strReturn As String             ' Return string
    
    If m.astrPositions.BinarySearch(strAccount & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        strReturn = m.astrPositions(lPos)
    Else
        strReturn = ""
    End If
    
    PositionForSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.PositionForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSummaryForSymbol
'' Description: Retrieve the fill summary info for the given account and symbol
'' Inputs:      Account, Symbol, Auto Trade ID
'' Returns:     Fill Summary for that account and symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillSummaryForSymbol(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long) As cAccountPosition
On Error GoTo ErrSection:

    Set FillSummaryForSymbol = m.FillSummaries(m.FillSummaries.Key(strAccount, strSymbol, lAtID))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.FillSummaryForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSummaryForID
'' Description: Retrieve the fill summary object for the given ID
'' Inputs:      Fill Summary ID
'' Returns:     Fill Summary for that ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillSummaryForID(ByVal lFillSummaryID As Long) As cAccountPosition
On Error GoTo ErrSection:

    Dim AcctPos As cAccountPosition     ' Fill summary to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    Set AcctPos = Nothing
    For lIndex = 1 To m.FillSummaries.Count
        If m.FillSummaries(lIndex).AccountPositionID = lFillSummaryID Then
            Set AcctPos = m.FillSummaries(lIndex)
            Exit For
        End If
    Next lIndex

    Set FillSummaryForID = AcctPos

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.FillSummaryForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SnapshotFillsForSymbol
'' Description: Retrieve the snapshot fills for the given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Snapshot Fills for that account and symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SnapshotFillsForSymbol(ByVal strAccount As String, ByVal strSymbol As String) As cPtFills
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fills As cPtFills               ' Collection of fills to return
    
    Set Fills = New cPtFills
    
    For lIndex = 1 To m.SnapshotFills.Count
        If (m.SnapshotFills(lIndex).AccountID = g.Broker.AccountIDForNumber(strAccount)) And (m.SnapshotFills(lIndex).Symbol = strSymbol) Then
            Fills.Add m.SnapshotFills(lIndex)
        End If
    Next lIndex
    
    Set SnapshotFillsForSymbol = Fills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.SnapshotFillsForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSummaries
'' Description: Retrieve the fill summaries
'' Inputs:      Only Connected?
'' Returns:     Fill Summaries
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillSummaries(Optional ByVal bOnlyConnected As Boolean = False) As cAccountPositions
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim FillSumms As cAccountPositions  ' Collection to return
    Dim strAccount As String            ' Account number for the account ID
    
    Set FillSumms = New cAccountPositions
    For lIndex = 1 To m.FillSummaries.Count
        strAccount = g.Broker.AccountNumberForID(m.FillSummaries(lIndex).AccountID)
        If (bOnlyConnected = False) Or (AccountExists(strAccount) = True) Then
            FillSumms.Add m.FillSummaries(lIndex)
        End If
    Next lIndex

    Set FillSummaries = FillSumms

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.FillSummaries"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSummariesForAccount
'' Description: Retrieve the fill summary info for the given account
'' Inputs:      Account
'' Returns:     Fill Summary for the account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillSummariesForAccount(ByVal lAccountID As Long) As cAccountPositions
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim FillSumms As cAccountPositions  ' Collection to return
    
    Set FillSumms = New cAccountPositions
    For lIndex = 1 To m.FillSummaries.Count
        If m.FillSummaries(lIndex).AccountID = lAccountID Then
            FillSumms.Add m.FillSummaries(lIndex)
        End If
    Next lIndex

    Set FillSummariesForAccount = FillSumms

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.FillSummariesForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenPositionsForAccount
'' Description: Retrieve the open positions for the given account
'' Inputs:      Account, Only Total?
'' Returns:     Open positions for the account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenPositionsForAccount(ByVal lAccountID As Long, Optional ByVal bOnlyTotal As Boolean = False) As cAccountPositions
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim FillSumms As cAccountPositions  ' Collection to return
    Dim bInclude As Boolean             ' Include in collection?
    
    Set FillSumms = New cAccountPositions
    For lIndex = 1 To m.FillSummaries.Count
        bInclude = False
        If m.FillSummaries(lIndex).AccountID = lAccountID Then
            bInclude = (m.FillSummaries(lIndex).CurrentPositionSnapshot <> 0)
            If (bOnlyTotal = True) And (bInclude = True) Then
                bInclude = (m.FillSummaries(lIndex).AutoTradeItemID = -1&)
            End If
        End If
        
        If bInclude Then
            FillSumms.Add m.FillSummaries(lIndex)
        End If
    Next lIndex

    Set OpenPositionsForAccount = FillSumms

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.OpenPositionsForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearAccounts
'' Description: Clear the accounts array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearAccounts()
On Error GoTo ErrSection:

    m.astrConnectedAccounts.Clear

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.ClearAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearOrders
'' Description: Clear the orders array
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearOrders(Optional ByVal strOnlyAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccount As String            ' Account number for the ID

    For lIndex = 1 To m.SnapshotOrders.Count
        strAccount = g.Broker.AccountNumberForID(m.SnapshotOrders(lIndex).AccountID)
        If (Len(strOnlyAccount) = 0) Or (strAccount = strOnlyAccount) Then
            If AccountExists(strAccount) Then
                m.SnapshotOrders(lIndex).Refreshed = False
            Else
                m.SnapshotOrders(lIndex).Refreshed = True
            End If
        End If
    Next lIndex

    m.EarliestOrderDate.Clear

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.ClearOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearFills
'' Description: Clear the fills array
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearFills(Optional ByVal strOnlyAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccount As String            ' Account number for the ID

    For lIndex = 1 To m.SnapshotFills.Count
        strAccount = g.Broker.AccountNumberForID(m.SnapshotFills(lIndex).AccountID)
        If (Len(strOnlyAccount) = 0) Or (strAccount = strOnlyAccount) Then
            If AccountExists(strAccount) Then
                m.SnapshotFills(lIndex).Refreshed = False
            Else
                m.SnapshotFills(lIndex).Refreshed = True
            End If
        End If
    Next lIndex

    m.EarliestFillDate.Clear

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.ClearFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearPositions
'' Description: Clear the positions array
'' Inputs:      Account
'' Returns:     None
''
'' Fields:      Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearPositions(Optional ByVal strOnlyAccount As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrPosition As cGdArray        ' Position information split from string
    Dim strAccount As String            ' Account number for the ID
    
    Set astrPosition = New cGdArray
    For lIndex = 0 To m.astrPositions.Size - 1
        astrPosition.SplitFields m.astrPositions(lIndex), vbTab
        strAccount = astrPosition(0)
        If (Len(strOnlyAccount) = 0) Or (strAccount = strOnlyAccount) Then
            If AccountExists(astrPosition(0)) Then
                astrPosition(2) = "0"
                astrPosition(3) = ""
                astrPosition(4) = Str(OvernightPosition(astrPosition(0), astrPosition(1), -1&))
                m.astrPositions(lIndex) = astrPosition.JoinFields(vbTab)
                
                g.Broker.BrokerDebug m.nBroker, "Position: " & m.astrPositions(lIndex)
                
                SymbolLastChanged(astrPosition(0), astrPosition(1)) = gdTickCount
            End If
        End If
    Next lIndex
        
    CreateForAutoTradeItems
    CreateForAutoExits

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.ClearPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearFillSummary
'' Description: Clear the fill summary array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearFillSummary()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.ClearFillSummary"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddAccount
'' Description: Add an item to the accounts array
'' Inputs:      Account String, Connected?, Reload Account?
'' Returns:     None
''
'' Fields:      Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddAccount(ByVal strAccountString As String, Optional ByVal bConnected As Boolean = True, Optional ByVal bReloadAccount As Boolean = False)
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account number to add
    Dim lPos As Long                    ' Position in the array
    Dim Account As cPtAccount           ' Account object
    
    strAccount = Parse(strAccountString, vbTab, 1)
    
    If bConnected = True Then
        If m.astrConnectedAccounts.BinarySearch(strAccount & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
            m.astrConnectedAccounts(lPos) = strAccount & vbTab
        Else
            m.astrConnectedAccounts.Add strAccount & vbTab, lPos
        End If
    End If
    
    If (m.Accounts.Exists(strAccount) = False) Or (bReloadAccount = True) Then
        Set Account = New cPtAccount
        If Account.Load(g.Broker.AccountIDForNumber(strAccount)) = True Then
            AddAccountObject Account
        End If
    End If
    
    m.dLastChanged = gdTickCount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.AddAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddAccountObject
'' Description: Add an item to the accounts collection
'' Inputs:      Account, Connection Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddAccountObject(Account As cPtAccount, Optional ByVal nStatus As eGDConnectionStatus = -1&)
On Error GoTo ErrSection:

    Dim nFillMatch As eTT_FillMatchMode ' Previous fill match mode for the account
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the array

    If m.Accounts.Exists(Account.AccountNumber) Then
        nFillMatch = m.Accounts(Account.AccountNumber).FillMatchMode
        m.Accounts(Account.AccountNumber) = Account
        
        If nFillMatch <> Account.FillMatchMode Then
            InfBox "Recalculating historical profit and loss with new fill matching.  Please wait...", , , "Recalculating History", True
            
            For lIndex = 1 To m.FillSummaries.Count
                m.FillSummaries(lIndex).FillMatchMode = Account.FillMatchMode
                m.FillSummaries(lIndex).Save
                m.FillSummaries(lIndex).RecalculateHistory
                
                SetDirtyChartTrades m.FillSummaries(lIndex).SymbolOrSymbolID, m.FillSummaries(lIndex).AccountID
            Next lIndex
            m.FillSummaries.UpdateLastModified
            g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( AddAccountObject )"
            
            RebuildAccountSummary Account.AccountID
            If FormIsLoaded("frmTTPositions") Then frmTTPositions.Reload m.nBroker
            
            InfBox ""
        End If
    Else
        m.Accounts.Add Account
    End If
    
    If nStatus = eGDConnectionStatus_Connected Then
        If m.astrConnectedAccounts.BinarySearch(Account.AccountNumber & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
            m.astrConnectedAccounts(lPos) = Account.AccountNumber & vbTab
        Else
            m.astrConnectedAccounts.Add Account.AccountNumber & vbTab, lPos
        End If
    End If
    
    m.dLastChanged = gdTickCount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.AddAccountObject"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountExists
'' Description: Does the given account number exist?
'' Inputs:      Account Number
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountExists(ByVal strAccountNumber As String) As Boolean
On Error GoTo ErrSection:

    AccountExists = m.astrConnectedAccounts.BinarySearch(strAccountNumber & vbTab, , eGdSort_MatchUsingSearchStringLength)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.AccountExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Account
'' Description: Return the account object for the given account information
'' Inputs:      Account Number or ID
'' Returns:     Account if exists, Nothing otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Account(ByVal vAccountNumberOrID As Variant) As cPtAccount
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account number
    
    strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
    If m.Accounts.Exists(strAccount) Then
        Set Account = m.Accounts(strAccount)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.Account"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrder
'' Description: Add an item to the orders array
'' Inputs:      Order String, Refresh?, Snapshot Order?, Use Previous Broker ID?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddOrder(Order As cPtOrder, ByVal bRefresh As Boolean, Optional ByVal bSnapshotOrder As Boolean = True, Optional ByVal bUsePreviousBrokerId As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the array
    Dim strAccount As String            ' Account number
    Dim strBaseSymbol As String         ' Base symbol for the order
    Dim lNewIndex As Long               ' New index for the order
    
    If Len(Order.Symbol) > 0 Then
        lPos = m.SnapshotOrders.IndexForOrder(Order, bUsePreviousBrokerId)
        
        If bSnapshotOrder = True Then
            If (Order.SessionDate <> 0) And (Order.Expiration <> 0) Then
                m.EarliestOrderDate.Update Order.AccountID, Order.SessionDate
            End If
            
            Order.Refreshed = True
        End If
        
        If lPos = -1& Then
            lNewIndex = m.SnapshotOrders.Add(Order.MakeCopy)
            g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: " & Order.OrderText(True, True, True) & " added to position " & Str(lNewIndex)
        Else
            m.SnapshotOrders(lPos) = Order.MakeCopy
            m.SnapshotOrders.Key(lPos) = Order.BrokerID
            g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: " & Order.OrderText(True, True, True) & " updated at position " & Str(lPos)
        End If
        
        If OrderIsPending(Order) Then
            If g.Broker.IsPitSymbol(Order.AccountID, Order.SymbolOrSymbolID) = False Then
                If Order.Broker <> eTT_AccountType_SimBroker Then
                    If m.PendingOrders.Exists(Str(Order.OrderID)) = False Then
                        m.PendingOrders.Add Order, Str(Order.OrderID)
                    End If
                End If
            End If
        Else
            If m.PendingOrders.Exists(Str(Order.OrderID)) = True Then
                m.PendingOrders.Remove Str(Order.OrderID)
            End If
        End If
        
        ' If the base symbol is not yet in the array of order symbols, add it...
        If bRefresh = True Then
            strBaseSymbol = Parse(Order.Symbol, "-", 1)
            If m.astrOrderSymbols.BinarySearch(strBaseSymbol, lPos) = False Then
                m.astrOrderSymbols.Add strBaseSymbol, lPos
                
                If m.bOrderRefreshHasOnlyTF = True Then
                    If UCase(strBaseSymbol) <> "TF" Then
                        m.bOrderRefreshHasOnlyTF = False
                    End If
                End If
            End If
        End If
        
        strAccount = g.Broker.AccountNumberForID(Order.AccountID)
        
        SymbolLastChanged(strAccount, Order.Symbol) = gdTickCount
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.AddOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Refresh an order if it exists in the collection
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshOrder(Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    
    If Len(Order.Symbol) > 0 Then
        lPos = m.SnapshotOrders.IndexForOrder(Order)
        
        If lPos <> -1& Then
            Order.Refreshed = m.SnapshotOrders(lPos).Refreshed
            m.SnapshotOrders(lPos) = Order.MakeCopy
        End If
        
        SymbolLastChanged(g.Broker.AccountNumberForID(Order.AccountID), Order.Symbol) = gdTickCount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RefreshOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFill
'' Description: Add an item to the fills array
'' Inputs:      Fill String, Refresh?, Recalculate?, Snapshot Fill?, Previous AT ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddFill(Fill As cPtFill, ByVal bRefresh As Boolean, Optional ByVal bRecalculateSnapshot As Boolean = True, Optional ByVal bSnapshotFill As Boolean = True, Optional ByVal lPrevAutoTradeID = kNullData)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim strAccount As String            ' Account number for the ID
    Dim strBaseSymbol As String         ' Base symbol for the fill
    
    If (Len(Fill.Symbol) > 0) And (Fill.IsSpread = False) Then
        lPos = m.SnapshotFills.IndexForFill(Fill)
        
        If bSnapshotFill = True Then
            If Fill.SessionDate <> 0 Then
                m.EarliestFillDate.Update Fill.AccountID, Fill.SessionDate
            End If
            
            Fill.Refreshed = True
        End If
        
        If lPos = -1& Then
            m.SnapshotFills.Add Fill.MakeCopy
        Else
            m.SnapshotFills(lPos) = Fill.MakeCopy
        End If
        
        ' If the base symbol is not yet in the array of order symbols, add it...
        If bRefresh = True Then
            strBaseSymbol = Parse(Fill.Symbol, "-", 1)
            If m.astrFillSymbols.BinarySearch(strBaseSymbol, lPos) = False Then
                m.astrFillSymbols.Add strBaseSymbol, lPos
                
                If m.bFillRefreshHasOnlyTF = True Then
                    If UCase(strBaseSymbol) <> "TF" Then
                        m.bFillRefreshHasOnlyTF = False
                    End If
                End If
            End If
        End If
        
        strAccount = g.Broker.AccountNumberForID(Fill.AccountID)
        SymbolLastChanged(strAccount, Fill.Symbol) = gdTickCount
        
        If bRecalculateSnapshot Then
            If (bRefresh = False) And ((m.bRefreshingFills = False) And (m.bRefreshingPositions = False)) Then
                ' If this is a brand new fill that came outside of the refresh, just add it to the snapshot
                ' fills and recalculate the position accordingly...
                If lPos = -1& Then
                    'm.FillSummaries.RecalculateSnapshotForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, Fill.AutoTradingItemID
                    m.FillSummaries.AddSnapshotFillForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, Fill.AutoTradingItemID, Fill
                    'm.FillSummaries.RecalculateSnapshotForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, -1&
                    m.FillSummaries.AddSnapshotFillForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, -1&, Fill
                
                ' Otherwise, if we get a fill we already know about outside of a refresh, do a full
                ' recalculate on the snapshot instead of adding it to the bottom (because it is already there)...
                Else
                    If lPrevAutoTradeID <> kNullData Then
                        m.FillSummaries.RecalculateSnapshotForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, lPrevAutoTradeID
                    End If
                    m.FillSummaries.RecalculateSnapshotForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, Fill.AutoTradingItemID
                    m.FillSummaries.RecalculateSnapshotForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, -1&
                End If
                               
                RebuildAccountSummary Fill.AccountID
                SetDirtyChartTrades Fill.SymbolOrSymbolID, Fill.AccountID
            Else
                m.FillSummaries.SetDirtyForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, Fill.AutoTradingItemID
                m.FillSummaries.SetDirtyForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, -1&
            End If
            m.FillSummaries.UpdateLastModified
            g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( AddFill )"
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.AddFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshFill
'' Description: Refresh a fill if it exists in the collection
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshFill(Fill As cPtFill)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    
    If Len(Fill.Symbol) > 0 Then
        lPos = m.SnapshotFills.IndexForFill(Fill)
        
        If lPos <> -1& Then
            Fill.Refreshed = m.SnapshotFills(lPos).Refreshed
            m.SnapshotFills(lPos) = Fill.MakeCopy
        End If
        
        SymbolLastChanged(g.Broker.AccountNumberForID(Fill.AccountID), Fill.Symbol) = gdTickCount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RefreshFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPosition
'' Description: Add an item to the positions array
'' Inputs:      Position String, Refresh?
'' Returns:     None
''
'' Fields:      Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddPosition(ByVal strPositionString As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim astrPosition As cGdArray        ' Position information split from string
    
    Set astrPosition = New cGdArray
    astrPosition.SplitFields strPositionString, vbTab
    
    If Len(astrPosition(1)) > 0 Then
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BI: Position Passed in: " & strPositionString
        If Len(astrPosition(4)) = 0 Then
            astrPosition(4) = Str(OvernightPosition(astrPosition(0), astrPosition(1), -1&))
            strPositionString = astrPosition.JoinFields(vbTab)
            g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BI: Fixed Position: " & strPositionString
        End If
        
        If m.astrPositions.BinarySearch(astrPosition(0) & vbTab & astrPosition(1) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
            m.astrPositions(lPos) = strPositionString
        Else
            m.astrPositions.Add strPositionString, lPos
        End If
                
        OvernightPosition(astrPosition(0), astrPosition(1), -1&) = CLng(Val(astrPosition(4)))
        
        If (bRefresh = False) And ((m.bRefreshingPositions = False) And (m.bRefreshingFills = False)) Then
            m.FillSummaries.RecalculateSnapshotForSymbol astrPosition(0), astrPosition(1), -1&
            
            RebuildAccountSummary astrPosition(0)
            SetDirtyChartTrades astrPosition(1), g.Broker.AccountIDForNumber(astrPosition(0))
        Else
            m.FillSummaries.SetDirtyForSymbol astrPosition(0), astrPosition(1), -1&
        End If
        m.FillSummaries.UpdateLastModified
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( AddPosition )"
        
        SymbolLastChanged(astrPosition(0), astrPosition(1)) = gdTickCount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.AddPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPositionFromObject
'' Description: Add an item to the positions array
'' Inputs:      Position Object, Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddPositionFromObject(ByVal Position As cPtPosition, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    AddPosition Position.ToString, bRefresh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.AddPositionFromObject"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveAccount
'' Description: Remove an item from the accounts array
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveAccount(ByVal Account As cPtAccount)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim lAccountID As Long              ' Account ID
    Dim strAccountNumber As String      ' Account Number
    
    lAccountID = Account.AccountID
    strAccountNumber = Account.AccountNumber
    
    g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: RemoveAccount ( " & Str(lAccountID) & ", " & strAccountNumber & " )"

    If m.astrConnectedAccounts.BinarySearch(strAccountNumber & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        m.astrConnectedAccounts.Remove lPos
    End If
    
    m.Accounts.Remove strAccountNumber
    m.FillSummaries.RemoveForAccount lAccountID
    m.SnapshotFills.RemoveForAccount lAccountID
    m.SnapshotOrders.RemoveForAccount lAccountID
    
    RemovePendingOrdersForAccount lAccountID
    RemovePositionsForAccount strAccountNumber
    
    m.dLastChanged = gdTickCount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RemoveAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveOrder
'' Description: Remove an item from the orders array
'' Inputs:      Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveOrder(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    
    lPos = m.SnapshotOrders.IndexForOrder(Order)
    If lPos > -1& Then
        m.SnapshotOrders.Remove lPos
        SymbolLastChanged(g.Broker.AccountNumberForID(Order.AccountID), Order.Symbol) = gdTickCount
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RemoveOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFill
'' Description: Remove an item from the fills array
'' Inputs:      Fill, Recalculate?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFill(Fill As cPtFill, Optional ByVal bRecalculateSnapshot As Boolean = True)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim strAccount As String            ' Account number for the given account ID
    
    lPos = m.SnapshotFills.IndexForFill(Fill)
    If lPos > -1& Then
        strAccount = g.Broker.AccountNumberForID(Fill.AccountID)
        m.SnapshotFills.Remove lPos
        
        If bRecalculateSnapshot Then
            m.FillSummaries.RecalculateSnapshotForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, Fill.AutoTradingItemID
            m.FillSummaries.RecalculateSnapshotForSymbol Fill.AccountID, Fill.SymbolOrSymbolID, -1&
            
            RebuildAccountSummary Fill.AccountID
            SetDirtyChartTrades Fill.SymbolOrSymbolID, Fill.AccountID
        End If
        m.FillSummaries.UpdateLastModified
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( RemoveFill )"
    
        SymbolLastChanged(strAccount, Fill.Symbol) = gdTickCount
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RemoveFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemovePosition
'' Description: Remove an item from the positions array
'' Inputs:      Account, Symbol
'' Returns:     None
''
'' Fields:      Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemovePosition(ByVal strAccount As String, ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    
    If m.astrPositions.BinarySearch(strAccount & vbTab & strSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        m.astrPositions.Remove lPos
    
        m.FillSummaries.RecalculateSnapshotForSymbol strAccount, strSymbol, -1&
        m.FillSummaries.UpdateLastModified
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( RemovePosition )"
        
        RebuildAccountSummary strAccount
        SetDirtyChartTrades strSymbol, g.Broker.AccountIDForNumber(strAccount)
    End If
    
    SymbolLastChanged(strAccount, strSymbol) = gdTickCount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RemovePosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RebuildFillSummaries
'' Description: Rebuild Fill Summaries for all items in the unique list
'' Inputs:      Account, Symbol, Rebuild History?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RebuildFillSummaries(Optional ByVal strAccount As String = "", Optional ByVal strSymbol = "", Optional ByVal bRebuildHistory As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bDidOne As Boolean              ' Did we recalculate any fill summaries?
    
    bDidOne = False

    For lIndex = 1 To m.FillSummaries.Count
        If (Len(strAccount) = 0) Or (g.Broker.AccountNumberForID(m.FillSummaries(lIndex).AccountID) = strAccount) Then
            If (Len(strSymbol) = 0) Or (m.FillSummaries(lIndex).Symbol = strSymbol) Then
                If bRebuildHistory Then
                    m.FillSummaries(lIndex).RecalculateHistory
                    m.FillSummaries(lIndex).Dirty = False
                ElseIf m.FillSummaries(lIndex).Dirty = True Then
                    m.FillSummaries(lIndex).RecalculateSnapshot
                    m.FillSummaries(lIndex).Dirty = False
                End If
                
                SetDirtyChartTrades m.FillSummaries(lIndex).SymbolOrSymbolID, m.FillSummaries(lIndex).AccountID
                SymbolLastChanged(g.Broker.AccountNumberForID(m.FillSummaries(lIndex).AccountID), m.FillSummaries(lIndex).Symbol) = gdTickCount
                
                bDidOne = True
            End If
        End If
    Next lIndex
    m.FillSummaries.UpdateLastModified
    g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( RebuildFillSummaries )"
    
    If bDidOne Then
        RebuildAccountSummaries
        m.dLastChanged = gdTickCount
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RebuildFillSummaries"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RebuildFillSummaryForSymbol
'' Description: Rebuild Fill Summaries for the given information
'' Inputs:      Account, Symbol, Auto Trade ID, Rebuild History?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RebuildFillSummaryForSymbol(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant, ByVal lAtID As Long, Optional ByVal bRebuildHistory As Boolean = False)
On Error GoTo ErrSection:

    ' Rebuild the entire fill summary history for the given information...
    If bRebuildHistory Then
        ' Had to move the total (-1) line first because when creating a fake fill and rebuilding
        ' history, the non-category info for the fill doesn't calculate at all the other way...
        If lAtID <> -1& Then
            m.FillSummaries.RecalculateHistoryForSymbol vAccountNumberOrID, vSymbolOrSymbolID, -1&
        End If
        m.FillSummaries.RecalculateHistoryForSymbol vAccountNumberOrID, vSymbolOrSymbolID, lAtID
        
    ' Rebuild only the snapshot fill summary for the given information...
    Else
        ' Had to move the total (-1) line first because when creating a fake fill and rebuilding
        ' history, the non-category info for the fill doesn't calculate at all the other way...
        If lAtID <> -1& Then
            m.FillSummaries.RecalculateSnapshotForSymbol vAccountNumberOrID, vSymbolOrSymbolID, -1&
        End If
        m.FillSummaries.RecalculateSnapshotForSymbol vAccountNumberOrID, vSymbolOrSymbolID, lAtID
    End If
    m.FillSummaries.UpdateLastModified
    g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( RebuildFillSummaryForSymbol )"
    
    ' Tell the charts to reload the trades...
    SetDirtyChartTrades vSymbolOrSymbolID, g.Broker.GetAccountID(vAccountNumberOrID)
    
    ' Tell the account to recalculate its closed profit...
    RebuildAccountSummary vAccountNumberOrID
    m.dLastChanged = gdTickCount
    
    ' If an appropriate Trade Tracker form is loaded, tell it to reload...
    If FormIsLoaded("frmTTPositions") Then frmTTPositions.Reload m.nBroker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RebuildFillSummaryForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RebuildFillSummaryForAccount
'' Description: Rebuild Fill Summaries for the given account
'' Inputs:      Account, Rebuild History?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RebuildFillSummaryForAccount(ByVal vAccountNumberOrID As Variant, Optional ByVal bRebuildHistory As Boolean = False)
On Error GoTo ErrSection:

    ' Rebuild the entire fill summary history for the given information...
    If bRebuildHistory Then
        m.FillSummaries.RecalculateHistoryForAccount vAccountNumberOrID
        
    ' Rebuild only the snapshot fill summary for the given information...
    Else
        m.FillSummaries.RecalculateSnapshotForAccount vAccountNumberOrID
    End If
    m.FillSummaries.UpdateLastModified
    g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( RebuildFillSummaryForAccount )"
    
    ' Tell the charts to reload the trades...
    SetDirtyChartTrades "", g.Broker.GetAccountID(vAccountNumberOrID)
    
    ' Tell the account to recalculate its closed profit...
    RebuildAccountSummary vAccountNumberOrID
    m.dLastChanged = gdTickCount
    
    ' If an appropriate Trade Tracker form is loaded, tell it to reload...
    If FormIsLoaded("frmTTPositions") Then frmTTPositions.Reload m.nBroker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RebuildFillSummaryForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RebuildAccountSummaries
'' Description: Build account summaries for all accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RebuildAccountSummaries()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Accounts.Count
        RebuildAccountSummary m.Accounts(lIndex).AccountID
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RebuildAccountSummaries"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTradeLine
'' Description: Get the trade line for the given account position and trade number
'' Inputs:      Account Position ID, Trade Number
'' Returns:     Trade Line (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTradeLine(ByVal lAccountPositionID As Long, ByVal lTradeNumber As Long) As cTradeLine
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeLine As cTradeLine         ' Trade line object to return
    
    Set TradeLine = Nothing
    For lIndex = 1 To m.FillSummaries.Count
        If m.FillSummaries(lIndex).AccountPositionID = lAccountPositionID Then
            Set TradeLine = m.FillSummaries(lIndex).GetTradeLine(lTradeNumber)
            Exit For
        End If
    Next lIndex
    
    Set GetTradeLine = TradeLine

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.GetTradeLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTradeLine
'' Description: Refresh the given trade line in the correct fill summary
'' Inputs:      Trade Line
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshTradeLine(ByVal TradeLine As cTradeLine)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.FillSummaries.Count
        If m.FillSummaries(lIndex).AccountPositionID = TradeLine.AccountPositionID Then
            m.FillSummaries(lIndex).RefreshTradeLine TradeLine
            m.FillSummaries.UpdateLastModified
            g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: LastModified changed to " & Str(m.FillSummaries.LastModified) & " ( RefreshTradeLine )"
            SymbolLastChanged(g.Broker.AccountNameForID(m.FillSummaries(lIndex).AccountID), m.FillSummaries(lIndex).Symbol) = gdTickCount
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RefreshTradeLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingOcoOrders
'' Description: Are there working OCO orders?
'' Inputs:      None
'' Returns:     True if Working OCO Orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingOcoOrders() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order object
    
    bReturn = False
    For lIndex = 1 To m.SnapshotOrders.Count
        Set Order = m.SnapshotOrders(lIndex)
        If Not Order Is Nothing Then
            If (IsOpenOrder(Order.Status) = True) And (Order.Status <> eTT_OrderStatus_Parked) Then
                If Order.CancelOrderID <> 0 Then
                    bReturn = True
                    Exit For
                End If
            End If
        End If
    Next lIndex
    
    HasWorkingOcoOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.HasWorkingOcoOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelWorkingOcoOrders
'' Description: Cancel all working OCO orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelWorkingOcoOrders()
On Error GoTo ErrSection:

    Dim Orders As New cPtOrders         ' Collection of orders to cancel
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order object
    
    For lIndex = 1 To m.SnapshotOrders.Count
        Set Order = m.SnapshotOrders(lIndex)
        If Not Order Is Nothing Then
            If (IsOpenOrder(Order.Status) = True) And (Order.Status = eTT_OrderStatus_Parked) Then
                If Order.CancelOrderID <> 0 Then
                    Orders.Add Order
                End If
            End If
        End If
    Next lIndex
    
    For lIndex = 1 To Orders.Count
        CancelOrderFromOrder Orders(lIndex), "Broker Info", False
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.CancelWorkingOcoOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameOptionNavGroup
'' Description: Rename the Option Nav group with the given ID
'' Inputs:      Group ID, New Group Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RenameOptionNavGroup(ByVal lGroupID As Long, ByVal strNewGroupName As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.SnapshotOrders.Count
        If m.SnapshotOrders(lIndex).GroupID = lGroupID Then
            m.SnapshotOrders(lIndex).GroupName = strNewGroupName
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RenameOptionNavGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeGenesisIdForOrder
'' Description: Change the Genesis ID for an order
'' Inputs:      Order, New Genesis ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeGenesisIdForOrder(Order As cPtOrder, ByVal strNewGenesisID As String)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of the order in the collection
    Dim strPrevGenesisID As String      ' Previous Genesis Order ID
    
    lPos = m.SnapshotOrders.IndexForOrder(Order)
    If lPos <> -1& Then
        strPrevGenesisID = Order.GenesisOrderID
        Order.GenesisOrderID = strNewGenesisID
        
        m.SnapshotOrders(lPos) = Order.MakeCopy
        m.SnapshotOrders.Key(lPos) = Order.BrokerID
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: " & Order.OrderText(True, True, True) & " updated at position " & Str(lPos) & " ( Changed GenesisOrderID from '" & strPrevGenesisID & "' to '" & strNewGenesisID & "')"
    Else
        Order.GenesisOrderID = strNewGenesisID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.ChangeGenesisIdForOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearRefreshFlags
'' Description: Clear the refresh flags
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearRefreshFlags()
On Error GoTo ErrSection:

    If m.bRefreshingAccounts = True Then
        m.bRefreshingAccounts = False
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: Cleared RefreshingAccounts flag"
    End If

    If m.bRefreshingOrders = True Then
        m.bRefreshingOrders = False
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: Cleared RefreshingOrders flag"
    End If

    If m.bRefreshingFills = True Then
        m.bRefreshingFills = False
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: Cleared RefreshingFills flag"
    End If

    If m.bRefreshingPositions = True Then
        m.bRefreshingPositions = False
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & "BInfo: Cleared RefreshingPositions flag"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.ClearRefreshFlags"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RebuildAccountSummary
'' Description: Rebuild the summary information for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RebuildAccountSummary(ByVal vAccountNumberOrID As Variant)
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID
    Dim strAccount As String            ' Account number
    Dim lIndex As Long                  ' Index into a for loop
    Dim dClosedProfit As Double         ' Closed profit
    Dim dOpenProfit As Double           ' Open profit
    Dim dTotalFees As Double            ' Total fees
    Dim dFillSumOpenProfit As Double    ' Open profit from the fill summary
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
    
    dClosedProfit = 0#
    dOpenProfit = 0#
    dTotalFees = 0#
    For lIndex = 1 To m.FillSummaries.Count
        If (m.FillSummaries(lIndex).AccountID = lAccountID) And (m.FillSummaries(lIndex).AutoTradeItemID = -1&) Then
            dClosedProfit = dClosedProfit + m.FillSummaries(lIndex).ClosedProfitSnapshot
            dFillSumOpenProfit = m.FillSummaries(lIndex).OpenProfit
            If dFillSumOpenProfit <> kNullData Then
                dOpenProfit = dOpenProfit + dFillSumOpenProfit
            End If
            dTotalFees = dTotalFees + m.FillSummaries(lIndex).TotalFeesTotal
        End If
    Next lIndex
    
    If m.Accounts.Exists(strAccount) Then
        With m.Accounts(strAccount)
            .ClosedProfit = dClosedProfit
            .OpenProfit = dOpenProfit
            .TotalFees = dTotalFees
            
            .Save
        End With
        m.Accounts.UpdateLastModified
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RebuildAccountSummary"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Order
'' Description: Retrieve the order with the given identifiers
'' Inputs:      Broker ID, Genesis ID
'' Returns:     Order (blank if not found)
''
'' Order:       Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
''              Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
''              Order Date, At ID, Session Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Order(ByVal strBrokerID As String, Optional ByVal strGenesisID As String = "") As cPtOrder
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the array
    
    lPos = m.SnapshotOrders.IndexForOrderIDs(strBrokerID, strGenesisID)
    If lPos > -1& Then
        Set Order = m.SnapshotOrders(lPos)
    Else
        Set Order = Nothing
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.Order"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPosition
'' Description: Pass back the current position for an account, symbol, and
''              auto trade id
'' Inputs:      Account, Symbol, Auto Trade ID
'' Returns:     Current Position
''
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentPosition(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long) As Long
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the fill summary collection
    Dim lReturn As Long                 ' Return value from the function
    
    lReturn = 0&
    strKey = m.FillSummaries.Key(strAccount, strSymbol, lAtID)
    If m.FillSummaries.Exists(strKey) Then
        lReturn = m.FillSummaries(strKey).CurrentPositionSnapshot
    End If
    
    CurrentPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.CurrentPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerPosition
'' Description: Pass back the current broker position for an account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Current Broker Position
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerPosition(ByVal strAccount As String, ByVal strSymbol As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value from the function
    Dim strPosition As String           ' Position information
    
    lReturn = 0&
    strPosition = PositionForSymbol(strAccount, strSymbol)
    If Len(strPosition) > 0 Then
        lReturn = CLng(Val(Parse(strPosition, vbTab, 3)))
    End If
    
    BrokerPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.BrokerPosition"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenEquity
'' Description: Determine the open equity for the given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Open Equity
''
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenEquity(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long) As Double
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the fill summary collection
    Dim dReturn As Double               ' Return value from the function
    
    dReturn = 0#
    strKey = m.FillSummaries.Key(strAccount, strSymbol, lAtID)
    If m.FillSummaries.Exists(strKey) Then
        dReturn = m.FillSummaries(strKey).OpenProfit
    End If
    
    OpenEquity = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.OpenEquity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionString
'' Description: Get the position string for the given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     Position String for Chart and Ladder
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
'' Return:      Direction, Qty, OpenEquity, AvgEntry Display, SessionQty,
''              SessionProfit, AvgEntry Actual
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionString(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of values to join for return string
    Dim strKey As String                ' Key into the fill summary collection
    
    Set astrReturn = New cGdArray
    
    If (ConsistentBroker(strAccount, strSymbol) = True) And (CarriedMatch(strAccount, strSymbol) = True) Then
        strKey = m.FillSummaries.Key(strAccount, strSymbol, lAtID)
        If m.FillSummaries.Exists(strKey) Then
            With m.FillSummaries(strKey)
                Select Case .CurrentPositionSnapshot
                    Case Is > 0
                        astrReturn(0) = "Long"
                        astrReturn(1) = Str(.CurrentPositionSnapshot)
                        astrReturn(2) = Format(.OpenProfit, "$#,##0.00")
                    Case Is < 0
                        astrReturn(0) = "Short"
                        astrReturn(1) = Str(Abs(.CurrentPositionSnapshot))
                        astrReturn(2) = Format(.OpenProfit, "$#,##0.00")
                    Case Else
                        astrReturn(0) = "Flat"
                        astrReturn(1) = ""
                        astrReturn(2) = ""
                End Select
                astrReturn(3) = .AverageEntrySnapshotString
                astrReturn(4) = Format(.SessionQuantitySnapshot, "#,##0")
                astrReturn(5) = Format(.SessionProfitSnapshot, "$#,##0.00")
                astrReturn(6) = Str(.AverageEntrySnapshot)
            End With
        End If
    Else
        astrReturn(0) = "Mismatch"
        astrReturn(1) = ""
        astrReturn(2) = ""
        astrReturn(3) = ""
        astrReturn(4) = ""
        astrReturn(5) = ""
        astrReturn(6) = ""
    End If
        
    PositionString = astrReturn.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.PositionString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAutoTradeItems
'' Description: Refresh the fill summaries with auto trade items
'' Inputs:      None
'' Returns:     None
''
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAutoTradeItems()
On Error GoTo ErrSection:

    CreateForAutoTradeItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RefreshAutoTradeItems"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAutoExits
'' Description: Refresh the fill summaries with auto exits
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAutoExits()
On Error GoTo ErrSection:

    CreateForAutoExits

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RefreshAutoExits"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradesForSymbol
'' Description: Return trade-by-trade file for the given information
'' Inputs:      Account, Symbol, Auto Trade ID, Period
'' Returns:     Trade-by-Trade file (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TradesForSymbol(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, Optional ByVal strPeriod = "Daily") As cGdArray
On Error GoTo ErrSection:
    
    Dim strKey As String                ' Key into the fill summary collection
    Dim FillSummary As cAccountPosition ' Fill summary object
    Dim strHeader As String             ' Header for the trade array
    Dim astrReturn As cGdArray          ' Array to return from the function
    
    strKey = m.FillSummaries.Key(strAccount, strSymbol, lAtID)
    If m.FillSummaries.Exists(strKey) Then
        Set FillSummary = m.FillSummaries(strKey)
        strHeader = BuildTradesHeader(0&, "Trades for " & strSymbol & " in account " & strAccount, strPeriod, FillSummary.StartFillTime, FillSummary.LastTradedSnapshot, 0#, strSymbol)
        Set astrReturn = FillSummary.TradeLines.ToArray
        astrReturn.Add strHeader, 0&
    Else
        Set astrReturn = Nothing
    End If
    
    Set TradesForSymbol = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.TradesForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionMatch
'' Description: Do we have a position match for the given symbol?
'' Inputs:      Account, Symbol
'' Returns:     True if Positions Match, False otherwise
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionMatch2(ByVal strAccount As String, ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim AcctPos As cAccountPosition     ' Fill summary
    Dim strPosition As String           ' Position information
    Dim lFillSummPos As Long            ' Calculated position
    Dim lPosition As Long               ' Position from the broker

    bReturn = True
    Set AcctPos = FillSummaryForSymbol(strAccount, strSymbol, -1&)
    strPosition = PositionForSymbol(strAccount, strSymbol)
    
    If (Not AcctPos Is Nothing) Then
        lFillSummPos = AcctPos.CurrentPositionSnapshot
        If Len(strPosition) > 0 Then
            lPosition = CLng(Val(Parse(strPosition, vbTab, 3)))
        Else
            lPosition = 0&
        End If
        bReturn = (lFillSummPos = lPosition)
    End If
    
    PositionMatch2 = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.PositionMatch"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateFillSummaryForAutoExit
'' Description: Create position and fill summary entries for the auto exit
'' Inputs:      Auto Exit
'' Returns:     None
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateFillSummaryForAutoExit(ByVal AutoExit As cActiveExit)
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account number for the auto trade item
    Dim strSymbol As String             ' Symbol for the auto trade item
    Dim strToAdd As String              ' String to add to an array
    
    strAccount = g.Broker.AccountNumberForID(AutoExit.AccountID)
    strSymbol = ConvertToTradeSymbol(GetSymbol(AutoExit.SymbolOrSymbolID))
            
    If AccountExists(strAccount) Then
        If Len(PositionForSymbol(strAccount, strSymbol)) = 0 Then
            strToAdd = strAccount & vbTab & strSymbol & vbTab & "0" & vbTab & "0" & vbTab & ""
            AddPosition strToAdd, True
        End If
        
        m.FillSummaries.CreateForSymbol strAccount, strSymbol, -1&
        m.FillSummaries.CreateForSymbol strAccount, strSymbol, 0&
        
        SymbolLastChanged(strAccount, strSymbol) = gdTickCount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.CreateFillSummaryForAutoExit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateFillSummaryForAutoTrade
'' Description: Create position and fill summary entries for the auto trade
'' Inputs:      Auto Trade
'' Returns:     Fill Summary ID
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateFillSummaryForAutoTrade(ByVal AutoTrade As cAutoTradeItem) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strAccount As String            ' Account number for the auto trade item
    Dim strSymbol As String             ' Symbol for the auto trade item
    Dim strToAdd As String              ' String to add to an array
    
    lReturn = 0&
    If AutoTrade.ParentID <> -1& Then
        strAccount = g.Broker.AccountNumberForID(AutoTrade.AccountID)
        strSymbol = GetSymbol(AutoTrade.TradeSymbolOrID)
    
        m.FillSummaries.CreateForSymbol strAccount, strSymbol, -1&
        m.FillSummaries.CreateForSymbol strAccount, strSymbol, 0&
        lReturn = m.FillSummaries.CreateForSymbol(strAccount, strSymbol, AutoTrade.AutoTradeItemID)
        
        If AccountExists(strAccount) Then
            If Len(PositionForSymbol(strAccount, strSymbol)) = 0 Then
                strToAdd = strAccount & vbTab & strSymbol & vbTab & "0" & vbTab & "0" & vbTab & ""
                AddPosition strToAdd, True
            End If
            
            SymbolLastChanged(strAccount, strSymbol) = gdTickCount
        End If
    End If
    
    CreateFillSummaryForAutoTrade = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.CreateFillSummaryForAutoTrade"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFillSummaryForAutoTrade
'' Description: Remove position and fill summary entries for the auto trade
'' Inputs:      Auto Trade
'' Returns:     None
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFillSummaryForAutoTrade(ByVal AutoTrade As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account number for the auto trade item
    Dim strSymbol As String             ' Symbol for the auto trade item
    Dim strToAdd As String              ' String to add to an array
    
    strAccount = g.Broker.AccountNumberForID(AutoTrade.AccountID)
    strSymbol = ConvertToTradeSymbol(GetSymbol(AutoTrade.SymbolOrSymbolID))
            
    If AccountExists(strAccount) Then
        'If Len(PositionForSymbol(strAccount, strSymbol)) = 0 Then
            'strToAdd = strAccount & vbTab & strSymbol & vbTab & "0" & vbTab & "0" & vbTab & ""
            'AddPosition strToAdd, True
        'End If
        
        'm.FillSummaries.CreateForSymbol strAccount, strSymbol, -1&
        'm.FillSummaries.CreateForSymbol strAccount, strSymbol, 0&
        'm.FillSummaries.CreateForSymbol strAccount, strSymbol, AutoTrade.AutoTradeItemID
        
        'SymbolLastChanged(strAccount, strSymbol) = gdTickCount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.RemoveFillSummaryForAutoTrade"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateIsSnapshot
'' Description: Determine if the given date is within the snapshot date range
'' Inputs:      Session Date
'' Returns:     True if Snapshot, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateIsSnapshot(ByVal lSessionDate As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lEarliestOrderDate As Long      ' Earliest order date
    Dim lEarliestFillDate As Long       ' Earliest fill date
    
    lEarliestOrderDate = m.EarliestOrderDate.EarliestDate
    lEarliestFillDate = m.EarliestFillDate.EarliestDate
    
    If (lEarliestFillDate = Abs(kNullData)) And (lEarliestOrderDate = Abs(kNullData)) Then
        bReturn = (lSessionDate >= Int(ConvertToBrokerDate(CurrentTime, m.nBroker, "IBM")))
    Else
        bReturn = (lSessionDate >= lEarliestFillDate) Or (lSessionDate >= lEarliestOrderDate)
    End If
    
    DateIsSnapshot = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.DateIsSnapshot"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateIsSnapshotForAccount
'' Description: Determine if the given date is within the snapshot date range
'' Inputs:      Session Date, Account
'' Returns:     True if Snapshot, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateIsSnapshotForAccount(ByVal lSessionDate As Long, ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lEarliestOrderDate As Long      ' Earliest order date
    Dim lEarliestFillDate As Long       ' Earliest fill date
    
    bReturn = False
    lEarliestOrderDate = m.EarliestOrderDate(lAccountID)
    lEarliestFillDate = m.EarliestFillDate(lAccountID)
    
    If (lEarliestOrderDate = Abs(kNullData)) And (lEarliestFillDate = Abs(kNullData)) Then
        bReturn = (lSessionDate >= Int(ConvertToBrokerDate(CurrentTime, m.nBroker, "IBM")))
    Else
        bReturn = (lSessionDate >= lEarliestFillDate) Or (lSessionDate >= lEarliestOrderDate)
    End If
        
    DateIsSnapshotForAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.DateIsSnapshotForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdatePrice
'' Description: Update the price for open profit for the given fill summary
'' Inputs:      Fill Summary Key, New Price
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdatePrice(ByVal strFillSummKey As String, ByVal dPrice As Double)
On Error GoTo ErrSection:

    Dim FillSumm As cAccountPosition    ' Account Position object

    If m.FillSummaries.Exists(strFillSummKey) Then
        Set FillSumm = m.FillSummaries(strFillSummKey)
        If FillSumm.CurrentPositionSnapshot <> 0& Then
            FillSumm.OpenProfit dPrice
            If FillSumm.AutoTradeItemID = -1& Then
                RebuildAccountSummary FillSumm.AccountID
            End If
            
            ''m.dLastChanged = gdTickCount
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.UpdatePrice"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForGenesisID
'' Description: Retrieve the order with the given Genesis Order ID and account
'' Inputs:      Account, Genesis Order ID
'' Returns:     Order if found, Nothing if not
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderForGenesisID(ByVal vAccountNumberOrID As Variant, ByVal strGenesisOrderID As String) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to return for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID
    
    Set Order = Nothing
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    
    For lIndex = 1 To m.SnapshotOrders.Count
        If (m.SnapshotOrders(lIndex).AccountID = lAccountID) And (m.SnapshotOrders(lIndex).GenesisOrderID = strGenesisOrderID) Then
            Set Order = m.SnapshotOrders(lIndex)
            Exit For
        End If
    Next lIndex
    
    Set OrderForGenesisID = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.OrderForGenesisID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForAccessID
'' Description: Retrieve the order with the given Access ID and account
'' Inputs:      Account, Access ID
'' Returns:     Order if found, Nothing if not
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderForAccessID(ByVal vAccountNumberOrID As Variant, ByVal lAccessID As Long) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to return for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID
    
    Set Order = Nothing
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    
    For lIndex = 1 To m.SnapshotOrders.Count
        If (m.SnapshotOrders(lIndex).AccountID = lAccountID) And (m.SnapshotOrders(lIndex).OrderID = lAccessID) Then
            Set Order = m.SnapshotOrders(lIndex)
            Exit For
        End If
    Next lIndex
    
    Set OrderForAccessID = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.OrderForAccessID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetConnectionStatus
'' Description: Set the connection status for all of the accounts
'' Inputs:      Connection Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetConnectionStatus(ByVal nStatus As eGDConnectionStatus)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nPrev As eGDConnectionStatus    ' Previous connection status
    Dim Account As cPtAccount           ' Account object
    
    For lIndex = 1 To m.Accounts.Count
        Set Account = m.Accounts(lIndex)
        
        nPrev = Account.ConnectionStatus
        Account.ConnectionStatus = nStatus
        m.Accounts.UpdateLastModified
        
        If nPrev <> nStatus Then
            SendAccountToOptionNav Account, False, nStatus
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.SetConnectionStatus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetConnectionStatusForAccount
'' Description: Set the connection status for the specified account
'' Inputs:      Connection Status, Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetConnectionStatusForAccount(ByVal nStatus As eGDConnectionStatus, ByVal vAccountNumberOrID As Variant)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccountNumber As String      ' Account number
    Dim nPrev As eGDConnectionStatus    ' Previous connection status
    Dim Account As cPtAccount           ' Account object
    
    strAccountNumber = g.Broker.GetAccountNumber(vAccountNumberOrID)
    If m.Accounts.Exists(strAccountNumber) = True Then
        Set Account = m.Accounts(strAccountNumber)
        
        nPrev = Account.ConnectionStatus
        Account.ConnectionStatus = nStatus
        m.Accounts.UpdateLastModified
        
        If nPrev <> nStatus Then
            SendAccountToOptionNav Account, False, nStatus
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.SetConnectionStatusForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetConnectionStatusForUser
'' Description: Set the connection status for accounts for specified user
'' Inputs:      Connection Status, User Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetConnectionStatusForUser(ByVal nStatus As eGDConnectionStatus, ByVal strUserName As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nPrev As eGDConnectionStatus    ' Previous connection status
    Dim Account As cPtAccount           ' Account object
    
    For lIndex = 1 To m.Accounts.Count
        If m.Accounts(lIndex).UserName = strUserName Then
            Set Account = m.Accounts(lIndex)
            
            nPrev = Account.ConnectionStatus
            Account.ConnectionStatus = nStatus
            m.Accounts.UpdateLastModified
        
            If nPrev <> nStatus Then
                SendAccountToOptionNav Account, False, nStatus
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.SetConnectionStatusForUser"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAltBrokerID
'' Description: Get the alternate broker ID given the broker ID
'' Inputs:      Broker ID
'' Returns:     Alternate Broker ID (Blank if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAltBrokerID(ByVal strBrokerID As String) As String
On Error GoTo ErrSection:

    GetAltBrokerID = m.SnapshotOrders.GetAltBrokerID(strBrokerID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.GetAltBrokerID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDatabaseIdForBrokerId
'' Description: Get the database ID for the order with the given the broker ID
'' Inputs:      Broker ID
'' Returns:     Database ID (Zero if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDatabaseIdForBrokerId(ByVal strBrokerID As String) As Long

    GetDatabaseIdForBrokerId = m.SnapshotOrders.GetDatabaseIdForBrokerId(strBrokerID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.GetDatabaseIdForBrokerId"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendToLog
'' Description: Send the local information for this object to a log
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendToLog()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    g.Broker.BrokerDebug m.nBroker, "==============BROKER INFO (" & g.Broker.BrokerName(m.nBroker) & ")=============="
    g.Broker.BrokerDebug m.nBroker, "--------------ACCOUNTS--------------"
    For lIndex = 1 To m.Accounts.Count
        g.Broker.BrokerDebug m.nBroker, vbTab & Str(lIndex) & vbTab & m.Accounts(lIndex).AccountNumber
    Next lIndex
    
    g.Broker.BrokerDebug m.nBroker, "--------------SNAPSHOT ORDERS--------------"
    For lIndex = 1 To m.SnapshotOrders.Count
        g.Broker.BrokerDebug m.nBroker, vbTab & Str(lIndex) & vbTab & m.SnapshotOrders(lIndex).OrderText(True, True, True)
    Next lIndex

    g.Broker.BrokerDebug m.nBroker, "--------------SNAPSHOT FILLS--------------"
    For lIndex = 1 To m.SnapshotFills.Count
        g.Broker.BrokerDebug m.nBroker, vbTab & Str(lIndex) & vbTab & mTradeTracker.FillDisplay(m.SnapshotFills(lIndex), True, False, True, True, True, False, False)
    Next lIndex
    
    g.Broker.BrokerDebug m.nBroker, "--------------POSITIONS--------------"
    For lIndex = 1 To m.FillSummaries.Count
        g.Broker.BrokerDebug m.nBroker, vbTab & Str(lIndex) & vbTab & m.FillSummaries(lIndex).ToFillSummary
    Next lIndex
    
    g.Broker.BrokerDebug m.nBroker, "============================================================="

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.SendToLog"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetExpirationDateForSymbol
'' Description: Set the expiration date for the given symbol
'' Inputs:      Symbol, Expiration Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetExpirationDateForSymbol(ByVal vSymbolOrSymbolID As Variant, ByVal lExpirationDate As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.FillSummaries.Count
        If m.FillSummaries(lIndex).SymbolOrSymbolID = vSymbolOrSymbolID Then
            m.FillSummaries(lIndex).ExpirationDate = lExpirationDate
            m.FillSummaries(lIndex).Save
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.SetExpirationDateForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateForAutoTradeItems
'' Description: Create position and fill summary entries for the auto trading
''              items for the accounts if they don't already exist
'' Inputs:      None
'' Returns:     None
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateForAutoTradeItems()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccountNumber As String      ' Account number for the auto trade item
    Dim strSymbol As String             ' Symbol for the auto trade item
    Dim lAutoTradeItemID As Long        ' Auto trade item ID
    Dim strToAdd As String              ' String to add to an array
    
    With g.TradingItems
        For lIndex = 1 To .Count
#If 0 Then
            strAccountNumber = g.Broker.AccountNumberForID(.Item(lIndex).AccountID)
            strSymbol = ConvertToTradeSymbol(GetSymbol(.Item(lIndex).SymbolOrSymbolID))
            lAutoTradeItemID = .Item(lIndex).AutoTradeItemID
            
            If AccountExists(strAccountNumber) Then
                If Len(PositionForSymbol(strAccountNumber, strSymbol)) = 0 Then
                    strToAdd = strAccountNumber & vbTab & strSymbol & vbTab & "0" & vbTab & "0" & vbTab & ""
                    AddPosition strToAdd, True
                End If
                
                m.FillSummaries.CreateForSymbol strAccountNumber, strSymbol, -1&
                m.FillSummaries.CreateForSymbol strAccountNumber, strSymbol, 0&
                m.FillSummaries.CreateForSymbol strAccountNumber, strSymbol, lAutoTradeItemID
                
                SymbolLastChanged(strAccountNumber, strSymbol) = gdTickCount
            End If
#End If
            If .Item(lIndex).Broker = m.nBroker Then
                CreateFillSummaryForAutoTrade .Item(lIndex)
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.CreateForAutoTradeItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateForAutoExits
'' Description: Create position and fill summary entries for the auto exits
''              for the accounts if they don't already exist
'' Inputs:      None
'' Returns:     None
''
'' Position:    Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateForAutoExits()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccountNumber As String      ' Account number for the auto trade item
    Dim strSymbol As String             ' Symbol for the auto trade item
    Dim strToAdd As String              ' String to add to an array
    
    With g.OrderStrategies
        For lIndex = 1 To .Count
            strAccountNumber = g.Broker.AccountNumberForID(.Item(lIndex).AccountID)
            strSymbol = ConvertToTradeSymbol(GetSymbol(.Item(lIndex).SymbolOrSymbolID))
            
            If AccountExists(strAccountNumber) Then
                If Len(PositionForSymbol(strAccountNumber, strSymbol)) = 0 Then
                    strToAdd = strAccountNumber & vbTab & strSymbol & vbTab & "0" & vbTab & "0" & vbTab & ""
                    AddPosition strToAdd, True
                End If
                
                m.FillSummaries.CreateForSymbol strAccountNumber, strSymbol, -1&
                m.FillSummaries.CreateForSymbol strAccountNumber, strSymbol, 0&
                
                SymbolLastChanged(strAccountNumber, strSymbol) = gdTickCount
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.CreateForAutoExits"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the broker information structure
'' Inputs:      None
'' Returns:     Copy of the broker information structure
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cBrokerInfo
On Error GoTo ErrSection:

    Dim aCopy As New cBrokerInfo        ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.MakeCopy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for copying purposes
'' Inputs:      Private members to copy from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    
    Set m.astrConnectedAccounts = mCopyFrom.astrConnectedAccounts.MakeCopy
    Set m.astrLastChanged = mCopyFrom.astrLastChanged.MakeCopy
    Set m.astrPositions = mCopyFrom.astrPositions.MakeCopy
    
    Set m.Accounts = mCopyFrom.Accounts.MakeCopy
    Set m.FillSummaries = mCopyFrom.FillSummaries.MakeCopy
    Set m.SnapshotFills = mCopyFrom.SnapshotFills.MakeCopy
    Set m.SnapshotOrders = mCopyFrom.SnapshotOrders.MakeCopy
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemovePendingOrdersForAccount
'' Description: Remove any pending orders for the given account
'' Inputs:      Account ID
'' Returns:     Number of items deleted
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemovePendingOrdersForAccount(ByVal lAccountID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Pending order
    
    lReturn = 0&
    For lIndex = m.PendingOrders.Count To 1 Step -1&
        Set Order = m.PendingOrders(lIndex)
        If Order.AccountID = lAccountID Then
            m.PendingOrders.Remove lIndex
            lReturn = lReturn + 1&
        End If
    Next lIndex
    
    RemovePendingOrdersForAccount = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.RemovePendingOrdersForAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemovePositionsForAccount
'' Description: Remove any positions for the given account
'' Inputs:      Account Number
'' Returns:     Number of items deleted
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemovePositionsForAccount(ByVal strAccountNumber As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim strKey As String                ' Key into the array
    Dim lLength As Long                 ' Length of the string
    
    lReturn = 0&
    strKey = strAccountNumber & vbTab
    lLength = Len(strKey)
    
    For lIndex = m.astrPositions.Size - 1 To 0 Step -1&
        If Left(m.astrPositions(lIndex), lLength) = strKey Then
            m.astrPositions.Remove lIndex
            lReturn = lReturn + 1&
        End If
    Next lIndex
    
    For lIndex = m.astrOvernight.Size - 1 To 0 Step -1&
        If Left(m.astrOvernight(lIndex), lLength) = strKey Then
            m.astrOvernight.Remove lIndex
            lReturn = lReturn + 1&
        End If
    Next lIndex
    
    RemovePositionsForAccount = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerInfo.RemovePositionsForAccount"
    
End Function

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

    Set m.astrLastChanged = New cGdArray
    m.astrLastChanged.Create eGDARRAY_Strings
    
    Set m.astrConnectedAccounts = New cGdArray
    m.astrConnectedAccounts.Create eGDARRAY_Strings
    
    Set m.astrPositions = New cGdArray
    m.astrPositions.Create eGDARRAY_Strings
    
    Set m.astrOvernight = New cGdArray
    m.astrOvernight.Create eGDARRAY_Strings
    
    Set m.astrPositionMatch = New cGdArray
    m.astrPositionMatch.Create eGDARRAY_Strings
    
    Set m.EarliestOrderDate = New cEarliestDate
    Set m.EarliestFillDate = New cEarliestDate
    
    Set m.Accounts = New cPtAccounts
    Set m.FillSummaries = New cAccountPositions
    Set m.SnapshotFills = New cPtFills
    Set m.SnapshotOrders = New cPtOrders
    Set m.PendingOrders = New cGdTree
    
    Set m.astrOrderSymbols = New cGdArray
    m.astrOrderSymbols.Create eGDARRAY_Strings
    Set m.astrFillSymbols = New cGdArray
    m.astrFillSymbols.Create eGDARRAY_Strings
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerInfo.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Destroy the class members when the object is destroyed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.astrLastChanged = Nothing
    Set m.astrConnectedAccounts = Nothing
    Set m.astrPositions = Nothing
    Set m.astrOvernight = Nothing
    Set m.astrPositionMatch = Nothing
    
    Set m.EarliestOrderDate = Nothing
    Set m.EarliestFillDate = Nothing
    
    Set m.Accounts = Nothing
    Set m.FillSummaries = Nothing
    Set m.SnapshotFills = Nothing
    Set m.SnapshotOrders = Nothing
    Set m.PendingOrders = Nothing
    
    Set m.astrOrderSymbols = Nothing
    Set m.astrFillSymbols = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cBrokerInfo.Class_Terminate"
    
End Sub
