VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOrderStrategies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOrderStrategies.cls
'' Description: Collection of active order strategies
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 12/01/2009   DAJ         Enhancements for Auto Exits held at broker
'' 01/04/2009   DAJ         Enhancements for activating/deactivating
'' 03/11/2010   DAJ         When asked for name, check if it has been removed
'' 07/21/2010   DAJ         Added message timeout to DeactivateExits
'' 09/22/2010   DAJ         Don't show deactivate exits message if all broker held (#5929)
'' 04/06/2011   DAJ         On a reassignment, do an immediate delete if no orders to be cancelled
'' 04/26/2011   DAJ         Make sure when exit gets removed from collection to delete from DB
'' 01/30/2012   DAJ         User configure timeout on auto exit reject
'' 07/30/2013   DAJ         Added the OrderExistsInStrategy function
'' 09/10/2014   DAJ         Consolidate the delete account code
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Strategies As cGdTree               ' Collection of order strategies
    
    bActivatingExits As Boolean         ' Are we in the process of activating exits?
    nRejectOption As eGDOrderRejectOption ' User option for how to handle a rejected order
    lRejectTimeout As Long              ' Timeout value on a question after a reject
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cActiveExit
    If m.Strategies.Exists(vIndexOrKey) Then
        Set Item = m.Strategies(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property

Public Property Get Count() As Long
    Count = m.Strategies.Count
End Property

Public Property Get RejectOption() As eGDOrderRejectOption
    RejectOption = m.nRejectOption
End Property
Public Property Let RejectOption(ByVal nRejectOption As eGDOrderRejectOption)
    If nRejectOption <> m.nRejectOption Then
        g.Broker.BrokerDebugAll "Reject Option Changed from '" & OrderRejectOptionString(m.nRejectOption) & "' to '" & OrderRejectOptionString(nRejectOption) & "'"
        m.nRejectOption = nRejectOption
        SetIniFileProperty "RejectOption", nRejectOption, "OrderStrategies", g.strIniFile
    End If
End Property

Public Property Get RejectTimeout() As Long
    RejectTimeout = m.lRejectTimeout
End Property
Public Property Let RejectTimeout(ByVal lRejectTimeout As Long)
    If lRejectTimeout <> m.lRejectTimeout Then
        g.Broker.BrokerDebugAll "Reject Timeout Changed from '" & Str(m.lRejectTimeout) & "' to '" & Str(lRejectTimeout) & "'"
        m.lRejectTimeout = lRejectTimeout
        SetIniFileProperty "RejectTimeout", lRejectTimeout, "OrderStrategies", g.strIniFile
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the active order strategies into the collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim ActiveExit As New cActiveExit   ' Active exit object
    
    ' Load the reject option from the INI file...
    m.nRejectOption = GetIniFileProperty("RejectOption", eGDOrderRejectOption_Ask, "OrderStrategies", g.strIniFile)
    g.Broker.BrokerDebugAll "Reject Option Loaded as '" & OrderRejectOptionString(m.nRejectOption) & "'"
    m.lRejectTimeout = GetIniFileProperty("RejectTimeout", 120&, "OrderStrategies", g.strIniFile)
    g.Broker.BrokerDebugAll "Reject Timeout Loaded as " & Str(m.lRejectTimeout)
    
    m.Strategies.Clear
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblActiveExits];", dbOpenDynaset)
    Do While Not rs.EOF
        Set ActiveExit = New cActiveExit
        If ActiveExit.LoadFromRecordset(rs) Then
            Set rs2 = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                        "WHERE [AccountID]=" & Str(ActiveExit.AccountID) & ";", dbOpenDynaset)
            If (rs2.BOF And rs2.EOF) Then
                rs.Delete
            Else
                AddExit ActiveExit
                g.Broker.CreateFillSummaryForAutoExit ActiveExit
            End If
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddExit
'' Description: Add the given active exit to the collection
'' Inputs:      Active Exit Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddExit(ActiveExit As cActiveExit)
On Error GoTo ErrSection:

    m.Strategies.Add ActiveExit, ActiveExit.Key
    RefreshAutoExitOnForms ActiveExit.AccountID, ActiveExit.SymbolOrSymbolID
    g.Broker.RefreshAutoExits g.Broker.AccountTypeForID(ActiveExit.AccountID)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.AddExit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveExit
'' Description: Remove the given active exit from the collection
'' Inputs:      Active Exit Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveExit(ActiveExit As cActiveExit)
On Error GoTo ErrSection:

    ' Don't remove any exits while in the process of activating, because that
    ' can cause items to go away that are in the process of trying to be activated
    ' thereby causing an "Object variable or With block variable not set" error...
    If m.bActivatingExits = False Then
        If m.Strategies.Exists(ActiveExit.Key) Then
            Item(ActiveExit.Key).Delete
            m.Strategies.Remove ActiveExit.Key
            RefreshAutoExitOnForms ActiveExit.AccountID, ActiveExit.SymbolOrSymbolID
            g.Broker.RefreshAutoExits g.Broker.AccountTypeForID(ActiveExit.AccountID)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.RemoveExit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteForAccount
'' Description: Remove the exits for the given account ID
'' Inputs:      Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteForAccount(ByVal lAccountID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = m.Strategies.Count To 1 Step -1
        If Item(lIndex).AccountID = lAccountID Then
            RemoveExit Item(lIndex)
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.DeleteForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExitExists
'' Description: Determine if the given active exit exists in the collection
'' Inputs:      Active Exit Object
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExitExists(ActiveExit As cActiveExit) As Boolean
On Error GoTo ErrSection:

    ExitExists = m.Strategies.Exists(ActiveExit.Key)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.ExitExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Call UpdateBars for each of the items in the collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBars()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).Active Then
            Item(lIndex).UpdateBars
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.UpdateBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderCallback
'' Description: Send the given order to the appropriate items in the collection
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderCallback(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If m.Strategies(lIndex).Active Then
            If m.Strategies(lIndex).HasOrder(Order.OrderID) Then
                m.Strategies(lIndex).OrderCallback Order
            End If
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.OrderCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCallback
'' Description: Send the given fill to the appropriate items in the collection
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillCallback(ByVal Fill As cPtFill, ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If m.Strategies(lIndex).Active Then
            If (m.Strategies(lIndex).AccountID = Fill.AccountID) And (m.Strategies(lIndex).SymbolOrSymbolID = Fill.SymbolOrSymbolID) Then
                Item(lIndex).FillCallback Fill, Order
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.FillCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Refresh the given order in the appropriate items in the collection
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshOrder(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If m.Strategies(lIndex).Active Then
            If m.Strategies(lIndex).HasOrder(Order.OrderID) Then
                Item(lIndex).RefreshOrder Order
            End If
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.RefreshOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPosition
'' Description: Refresh the position for the given account and symbol
'' Inputs:      Account, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPosition(ByVal vAccountNumberOrID As Variant, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID
    
    lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).Active Then
            If (Item(lIndex).AccountID = lAccountID) And (Item(lIndex).SymbolOrSymbolID = vSymbolOrSymbolID) Then
                Item(lIndex).RefreshPosition
                Exit For
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.RefreshPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIDChanged
'' Description: Notify the appropriate item that an order ID may have changed
'' Inputs:      Old Order ID, New Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderIDChanged(ByVal lOldOrderID As Long, ByVal lNewOrderID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If m.Strategies(lIndex).Active Then
            If m.Strategies(lIndex).HasOrder(lOldOrderID) Then
                m.Strategies(lIndex).OrderIDChanged lOldOrderID, lNewOrderID
            End If
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.OrderIDChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderExistsInStrategy
'' Description: Does the given order go to one of the order strategies?
'' Inputs:      Order, Strategy Name ( out )
'' Returns:     True if goes to an order strategy, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderExistsInStrategy(ByVal Order As cPtOrder, Optional strName As String) As Boolean
On Error GoTo ErrSection:

    OrderExistsInStrategy = OrderIdExistsInStrategy(Order.OrderID, strName)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.OrderExistsInStrategy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIdExistsInStrategy
'' Description: Does the given order ID go to one of the order strategies?
'' Inputs:      Order, Strategy Name ( out )
'' Returns:     True if goes to an order strategy, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderIdExistsInStrategy(ByVal lOrderID As Long, Optional strName As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    If lOrderID > 0 Then
        For lIndex = 1 To m.Strategies.Count
            If Item(lIndex).HasOrder(lOrderID) = True Then
                strName = Item(lIndex).Name
                bReturn = True
                Exit For
            End If
        Next lIndex
    End If
    
    OrderIdExistsInStrategy = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.OrderIdExistsInStrategy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExitForAccountAndSymbol
'' Description: Find the currently active exit order strategy for the account
''              and symbol that are passed in
'' Inputs:      Account ID, Symbol or Symbol ID, File Name?
'' Returns:     Exit Order Strategy Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExitForAccountAndSymbol(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, Optional ByVal bFileName As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    Dim strReturn As String             ' Return value from the function
    Dim vTradeSymbol As Variant         ' Trade symbol for the symbol passed in
    
    vTradeSymbol = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    If GetSymbolID(vTradeSymbol) <> 0 Then vTradeSymbol = GetSymbolID(vTradeSymbol) Else vTradeSymbol = GetSymbol(vTradeSymbol)
    strKey = "Exit" & vbTab & Str(lAccountID) & vbTab & Str(vTradeSymbol)
    strReturn = ""
    
    If m.Strategies.Exists(strKey) Then
        If Item(strKey).ToDelete = False Then
            If bFileName Then
                strReturn = Item(strKey).ExitStrategy.FileName
            Else
                strReturn = Item(strKey).ExitStrategy.StrategyName
            End If
        End If
    End If
    
    ExitForAccountAndSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.ExitForAccountAndSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExitForAccountAndSymbol
'' Description: Find the currently active exit order strategy for the account
''              and symbol that are passed in
'' Inputs:      Account ID, Symbol or Symbol ID
'' Returns:     Exit Order Strategy Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExitObjectForAccountAndSymbol(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant) As cActiveExit
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    Dim AutoExit As cActiveExit         ' Return object from the function
    Dim vTradeSymbol As Variant         ' Trade symbol for the symbol passed in
    Dim lSymbolID As Long               ' Symbol ID for symbol passed in
    
    vTradeSymbol = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    lSymbolID = GetSymbolID(vTradeSymbol)
    If lSymbolID = 0 Then
        strKey = "Exit" & vbTab & Str(lAccountID) & vbTab & Str(vTradeSymbol)
    Else
        strKey = "Exit" & vbTab & Str(lAccountID) & vbTab & Str(lSymbolID)
    End If
    
    Set AutoExit = Nothing
    If m.Strategies.Exists(strKey) Then
        Set AutoExit = Item(strKey)
    End If
    
    Set ExitObjectForAccountAndSymbol = AutoExit

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.ExitObjectForAccountAndSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateExit
'' Description: Activate the given exit strategy for the given account and symbol
'' Inputs:      Account ID, Symbol or Symbol ID, Strategy, Current Info
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ActivateExit(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, ByVal strStrategy As String, Optional ByVal strCurrentInfo As String = "") As Boolean
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    Dim AutoExit As cActiveExit         ' New active exit item
    Dim bContinue As Boolean            ' Do we need to create a new item?
    Dim lTimeOut As Long                ' Timeout variable
    Dim vTradeSymbol As Variant         ' Trade symbol for the symbol passed in
    Dim bExists As Boolean              ' Does the exit strategy exist?
    
    vTradeSymbol = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    If GetSymbolID(vTradeSymbol) <> 0 Then vTradeSymbol = GetSymbolID(vTradeSymbol) Else vTradeSymbol = GetSymbol(vTradeSymbol)
    strKey = "Exit" & vbTab & Str(lAccountID) & vbTab & Str(vTradeSymbol)
    
    If m.Strategies.Exists(strKey) Then
        If UCase(Item(strKey).ExitStrategyFile) = UCase(strStrategy) Then
            bContinue = Not Item(strKey).Active
            bExists = True
        Else
            DeactivateExit lAccountID, vTradeSymbol, True, "Switching Auto Exit Strategy"
            
            lTimeOut = 0&
            Do While m.Strategies.Exists(strKey)
                Sleep 1#
                lTimeOut = lTimeOut + 1&
                
                If m.Strategies.Exists(strKey) Then
                    If (Item(strKey).HasWorkingOrders = False) Or (lTimeOut >= 5&) Then Exit Do
                End If
            Loop
            
            If m.Strategies.Exists(strKey) Then
                bContinue = Not Item(strKey).HasWorkingOrders
            Else
                bContinue = True
            End If
            bExists = False
        End If
    Else
        bContinue = True
        bExists = False
    End If
    
    If bContinue Then
        If bExists = False Then
            Set AutoExit = New cActiveExit
            With AutoExit
                .AccountID = lAccountID
                .SymbolOrSymbolID = vTradeSymbol
                .ExitStrategyFile = strStrategy
                .Save
                
                ' If we cannot load the strategy, that means that the user didn't want to because
                ' of the "too close to trail" warning, so in that case, we don't want to continue...
                bContinue = .LoadStrategy
                If bContinue = False Then
                    .Delete
                Else
                    AutoExit.LoadBars
                    AddExit AutoExit
                                
                    ' Let the frmTTSummary.tmrBroker take over and reload the grid...
                    g.Broker.SymbolLastChanged(lAccountID, vTradeSymbol) = gdTickCount
                End If
            End With
        Else
            Set AutoExit = Item(strKey)
        End If
        
        If bContinue Then
            If AutoExit.Activate(strCurrentInfo) = False Then
                DeactivateExit lAccountID, vTradeSymbol, True, "Activation Failed due to one or more orders being on wrong side of market"
            End If
        End If
    End If
    
    ActivateExit = bContinue
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.ActivateExit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateExits
'' Description: Activate whatever exits can be activated
'' Inputs:      Only if Waiting on Data?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ActivateExits(Optional ByVal bOnlyIfWaitingOnData As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    If g.bUnloading = False Then
        m.bActivatingExits = True
        For lIndex = 1 To m.Strategies.Count
            If Item(lIndex).Active = False Then
                If (bOnlyIfWaitingOnData = False) Or (Item(lIndex).WaitingOnData = True) Then
                    If Item(lIndex).Activate = False Then
                        Item(lIndex).Deactivate True, "Activation Failed due to one or more orders being on wrong side of market"
                    End If
                End If
            End If
        Next lIndex
        m.bActivatingExits = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.ActivateExits"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateExitsForBroker
'' Description: Activate whatever exits can be activated for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ActivateExitsForBroker(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    If g.bUnloading = False Then
        m.bActivatingExits = True
        For lIndex = 1 To m.Strategies.Count
            If g.Broker.AccountTypeForID(Item(lIndex).AccountID) = nBroker Then
                If Item(lIndex).Activate = False Then
                    Item(lIndex).Deactivate True, "Activation Failed due to one or more orders being on wrong side of market"
                End If
            End If
        Next lIndex
        m.bActivatingExits = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.ActivateExitsForBroker"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeactivateExits
'' Description: Deactivate whatever exits can be deactivated
'' Inputs:      Reason, Ignore Broker Held?, Wait until Done?, Message Timeout
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeactivateExits(Optional ByVal strReason As String = "", Optional ByVal bIgnoreBrokerHeld As Boolean = False, Optional ByVal bWaitUntilDone As Boolean = True, Optional ByVal lMsgTimeout As Long = 0&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lTimeOut As Long                ' Timeout value
    
    If HasWorkingExits(bIgnoreBrokerHeld) Then
        InfBox "All of your auto exits are being disabled.  Any working orders associated with any auto exit will be cancelled at this time.", "i", , "Auto Exit Deactivation Warning", , lMsgTimeout
    End If
    
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).Active = True Then
            If (bIgnoreBrokerHeld = False) Or (Item(lIndex).HoldAtBroker = False) Then
                Item(lIndex).Deactivate False, strReason
            End If
        End If
    Next lIndex
    
    If (bWaitUntilDone = True) Then
        lTimeOut = 0&
        Do While (HasDeactivatingExits = True) And (lTimeOut < 30)
            Sleep 1, False, True
            lTimeOut = lTimeOut + 1&
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.DeactivateExits"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeactivateExit
'' Description: Deactivate the exit order strategy for the given account and symbol
'' Inputs:      Account ID, Symbol or Symbol ID, Manual Deactivate?, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeactivateExit(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant, Optional ByVal bManualDeactivate As Boolean = False, Optional strReason As String = "", Optional ByVal bNotIfBrokerHeld As Boolean = False)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    Dim vTradeSymbol As Variant         ' Trade symbol for the symbol passed in
    
    vTradeSymbol = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    If GetSymbolID(vTradeSymbol) <> 0 Then vTradeSymbol = GetSymbolID(vTradeSymbol) Else vTradeSymbol = GetSymbol(vTradeSymbol)
    strKey = "Exit" & vbTab & Str(lAccountID) & vbTab & Str(vTradeSymbol)
    
    If m.Strategies.Exists(strKey) Then
        If (bNotIfBrokerHeld = False) Or (Item(strKey).HoldAtBroker = False) Then
            Item(strKey).Deactivate bManualDeactivate, strReason
            
            ' DAJ 04/06/2011: If the active flag is still on, that means that we are waiting for
            ' orders to be cancelled, so we want to queue up the delete.  Otherwise, do the delete
            ' immediately...
            If Item(strKey).Active Then
                Item(strKey).ToDelete = True
            Else
                RemoveExit Item(strKey)
            End If
        
            g.Broker.SymbolLastChanged(lAccountID, vTradeSymbol) = gdTickCount
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.DeactivateExit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeactivateExitsForBroker
'' Description: Deactivate all of the active exits for the given broker
'' Inputs:      Broker, Reason, Not if Held at Broker?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeactivateExitsForBroker(ByVal nBroker As eTT_AccountType, Optional ByVal strReason As String = "", Optional ByVal bNotIfBrokerHeld As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If g.Broker.AccountTypeForID(Item(lIndex).AccountID) = nBroker Then
            If (bNotIfBrokerHeld = False) Or (Item(lIndex).HoldAtBroker = False) Then
                Item(lIndex).Deactivate False, strReason
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.DeactivateExitsForBroker"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshExitStrategy
'' Description: Refresh any active exits that have the given exit strategy
'' Inputs:      Exit Strategy
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshExitStrategy(ByVal strExitStrategyFile As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If (Item(lIndex).ExitStrategyFile = strExitStrategyFile) Then
            Item(lIndex).RefreshStrategy
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.RefreshExitStrategy"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StrategyRenamed
'' Description: Tell active exits with the given strategy name that it has been
''              renamed
'' Inputs:      Old Exit Strategy File, New Strategy File
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExitStrategyRenamed(ByVal strOldExitStrategyFile As String, ByVal strNewExitStrategyFile As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If (Item(lIndex).ExitStrategyFile = strOldExitStrategyFile) Then
            Item(lIndex).StrategyRenamed strNewExitStrategyFile
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.ExitStrategyRenamed"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Flatten
'' Description: Tell an order strategy to flatten itself
'' Inputs:      Account ID, Symbol or Symbol ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Flatten(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    Dim vTradeSymbol As Variant         ' Trade symbol for the symbol passed in
    
    vTradeSymbol = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    strKey = "Exit" & vbTab & Str(lAccountID) & vbTab & Str(vTradeSymbol)
    
    If m.Strategies.Exists(strKey) Then
        Item(strKey).Flatten
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.Flatten"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAllOrders
'' Description: Tell an order strategy to cancel all orders
'' Inputs:      Account ID, Symbol or Symbol ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelAllOrders(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    Dim vTradeSymbol As Variant         ' Trade symbol for the symbol passed in
    
    vTradeSymbol = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    strKey = "Exit" & vbTab & Str(lAccountID) & vbTab & Str(vTradeSymbol)
    
    If m.Strategies.Exists(strKey) Then
        Item(strKey).CancelAllOrders
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.CancelAllOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingExits
'' Description: Are there any working exits?
'' Inputs:      Ignore Broker Held?
'' Returns:     True if working exits, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingExits(Optional ByVal bIgnoreBrokerHeld As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).HasWorkingOrders(, bIgnoreBrokerHeld) Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    HasWorkingExits = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.HasWorkingExits"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingExitsForBroker
'' Description: Are there working exits for the given broker?
'' Inputs:      Broker, Not All Held at Broker?
'' Returns:     True if working exits, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingExitsForBroker(ByVal nBroker As eTT_AccountType, Optional bNotAllHeldAtBroker As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    bNotAllHeldAtBroker = False
    
    For lIndex = 1 To m.Strategies.Count
        If g.Broker.AccountTypeForID(Item(lIndex).AccountID) = nBroker Then
            If Item(lIndex).HasWorkingOrders Then
                If bNotAllHeldAtBroker = False Then
                    bNotAllHeldAtBroker = (Item(lIndex).HoldAtBroker = False)
                End If
                bReturn = True
                
                ' 11/13/2009 DAJ: Need to walk through all of the auto exits for the given
                ' broker type to find out if all are held at the broker...
                'Exit For
            End If
        End If
    Next lIndex
    
    HasWorkingExitsForBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.HasWorkingExitsForBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteInactiveExits
'' Description: Delete any inactive exits that have no working orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteInactiveExits()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = m.Strategies.Count To 1 Step -1
        If Item(lIndex).ToDelete = True Then
            If Item(lIndex).HasWorkingOrders = False Then
                RemoveExit Item(lIndex)
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.DeleteInactiveExits"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateSessionDate
'' Description: Update the session date for any auto exits with the given symbol
'' Inputs:      Symbol or Symbol ID, Session Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateSessionDate(ByVal vSymbolOrSymbolID As Variant, ByVal lSessionDate As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim vTradeSymbolOrID As Variant     ' Trading
    Dim lSymbolID As Long               ' Symbol ID passed in
    
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    If lSymbolID = 0 Then
        vTradeSymbolOrID = ConvertToTradeSymbol(vSymbolOrSymbolID, Int(CurrentTime("", "", True)))
    Else
        vTradeSymbolOrID = ConvertToTradeSymbol(lSymbolID, Int(CurrentTime("", "", True)))
    End If
    
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).SymbolOrSymbolID = vTradeSymbolOrID Then
            Item(lIndex).UpdateSessionDate lSessionDate
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.UpdateSessionDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelRequested
'' Description: Cancel was requested by Trade Navigator for the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelRequested(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).HasOrder(Order.OrderID) Then
            Item(lIndex).CancelRequested Order
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.CancelRequested"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckModifyOrders
'' Description: Check to see if any orders need to (and can be) modified
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckModifyOrders()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).Active Then
            Item(lIndex).CheckModifyOrders
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.CheckModifyOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasDeactivatingExits
'' Description: Are there deactivating exits in the collection?
'' Inputs:      None
'' Returns:     True if deactivating exits, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasDeactivatingExits() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).Deactivating Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    HasDeactivatingExits = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.HasDeactivatingExits"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasDeactivatingExitsForBroker
'' Description: Are there deactivating exits in the collection for the given broker?
'' Inputs:      Broker
'' Returns:     True if deactivating exits, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasDeactivatingExitsForBroker(ByVal nBroker As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    For lIndex = 1 To m.Strategies.Count
        If g.Broker.AccountTypeForID(Item(lIndex).AccountID) = nBroker Then
            If Item(lIndex).Deactivating Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    HasDeactivatingExitsForBroker = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.HasDeactivatingExitsForBroker"
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAutoExitOnForms
'' Description: Refresh appropriate forms as to an auto exit change
'' Inputs:      Account ID, Symbol or Symbol ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshAutoExitOnForms(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To Forms.Count - 1
        If IsFrmChart(Forms(lIndex)) Then
            If (Forms(lIndex).TradeAccountID = lAccountID) And (ConvertToTradeSymbol(Forms(lIndex).SymbolOrSymbolID, Int(CurrentTime("", "", True))) = vSymbolOrSymbolID) Then
                Forms(lIndex).SetAutoExit
            End If
        ElseIf TypeOf Forms(lIndex) Is frmTickDistribution Then
            If (Forms(lIndex).TradeAccountID = lAccountID) And (ConvertToTradeSymbol(Forms(lIndex).SymbolOrSymbolID, Int(CurrentTime("", "", True))) = vSymbolOrSymbolID) Then
                Forms(lIndex).SetAutoExit
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.RefreshAutoExitOnForms"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasActiveExits
'' Description: Are there active exits in the collection?
'' Inputs:      None
'' Returns:     True if active exits, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasActiveExits() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).Active = True Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    HasActiveExits = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.HasActiveExits"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasInactiveExits
'' Description: Are there inactive exits in the collection?
'' Inputs:      None
'' Returns:     True if inactive exits, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasInactiveExits() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    For lIndex = 1 To m.Strategies.Count
        If Item(lIndex).Active = False Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    HasInactiveExits = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderStrategies.HasInactiveExits"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize members of the class when the object is created
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Strategies = New cGdTree
    m.bActivatingExits = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate members of the class when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Strategies = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderStrategies.Class_Terminate"
    
End Sub
