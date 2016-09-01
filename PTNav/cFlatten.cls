VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlatten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cFlatten.cls
'' Description: Object that handles flattening a position
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/11/2010   DAJ         Use global Order Strategies collection
'' 09/13/2011   DAJ         Clear InfBox if flatten check proves already flat
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strAccountNumber As String          ' Account Number
    lAccountID As Long                  ' Account ID
    strSymbol As String                 ' Symbol
    lSymbolID As Long                   ' Symbol ID
    strTradeSymbol As String            ' Trade Symbol
    lTradeSymbolID As Long              ' Trade Symbol ID
    lAutoTradeItemID As Long            ' Automated trading item ID
    lSkipOrderID As Long                ' Do not cancel this order
    lOldPosition As Long                ' Old position the user was in
    
    nBroker As eTT_AccountType          ' Broker for the given account
    
    bDoneWorkingCheck As Boolean        ' Have we done the working orders check?
    dWorkingCheckStart As Double        ' Working order check starting time
    dWorkingCheckTimeOut As Double      ' Time out value for the working order check
    bDoneMarketOrder As Boolean         ' Have we done the market order?
    dMarketOrderStart As Double         ' Market order starting time
    nOperation As eGDFlattenQueueOperations ' Operation to perform
    strOperation As String              ' String representation of the operation to perform
    bAutoExitOnly As Boolean            ' Cancel auto exit orders only?
End Type
Private m As mPrivate

Public Property Get AccountNumber() As String
    AccountNumber = m.strAccountNumber
End Property
Public Property Let AccountNumber(ByVal strAccountNumber As String)
    m.strAccountNumber = strAccountNumber
    m.lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
    m.nBroker = g.Broker.AccountTypeForNumber(strAccountNumber)
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
    m.strAccountNumber = g.Broker.AccountNumberForID(lAccountID)
    m.nBroker = g.Broker.AccountTypeForID(lAccountID)
End Property

Public Property Get AccountNumberOrID() As Variant
    If m.lAccountID = 0& Then
        AccountNumberOrID = m.strAccountNumber
    Else
        AccountNumberOrID = m.lAccountID
    End If
End Property
Public Property Let AccountNumberOrID(ByVal vAccountNumberOrID As Variant)
    m.lAccountID = g.Broker.GetAccountID(vAccountNumberOrID)
    m.strAccountNumber = g.Broker.GetAccountNumber(vAccountNumberOrID)
    m.nBroker = g.Broker.AccountTypeForID(m.lAccountID)
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    m.strSymbol = strSymbol
    m.lSymbolID = GetSymbolID(strSymbol)
    m.strTradeSymbol = ConvertToTradeSymbol(strSymbol)
    m.lTradeSymbolID = GetSymbolID(m.strTradeSymbol)
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Public Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymbolID = lSymbolID
    m.strSymbol = GetSymbol(lSymbolID)
    m.lTradeSymbolID = ConvertToTradeSymbol(lSymbolID)
    m.strTradeSymbol = GetSymbol(m.lTradeSymbolID)
End Property

Public Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0& Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
    m.lTradeSymbolID = GetSymbolID(ConvertToTradeSymbol(vSymbolOrSymbolID))
    m.strTradeSymbol = GetSymbol(ConvertToTradeSymbol(vSymbolOrSymbolID))
End Property

Public Property Get TradeSymbolOrSymbolID() As Variant
    If m.lTradeSymbolID = 0& Then
        TradeSymbolOrSymbolID = m.strTradeSymbol
    Else
        TradeSymbolOrSymbolID = m.lTradeSymbolID
    End If
End Property

Public Property Get AutoTradeItemID() As Long
    AutoTradeItemID = m.lAutoTradeItemID
End Property
Public Property Let AutoTradeItemID(ByVal lAutoTradeItemID As Long)
    m.lAutoTradeItemID = lAutoTradeItemID
End Property

Public Property Get SkipOrderID() As Long
    SkipOrderID = m.lSkipOrderID
End Property
Public Property Let SkipOrderID(ByVal lSkipOrderID As Long)
    m.lSkipOrderID = lSkipOrderID
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get Key() As String
    Key = m.strAccountNumber & vbTab & m.strSymbol & vbTab & Str(m.lAutoTradeItemID) & vbTab & Str(m.nOperation)
End Property

Public Property Get Operation() As eGDFlattenQueueOperations
    Operation = m.nOperation
End Property
Public Property Let Operation(ByVal nOperation As eGDFlattenQueueOperations)
    m.nOperation = nOperation
    Select Case m.nOperation
        Case eGDFlattenQueueOperation_CancelAll
            m.strOperation = "Cancel All"
        Case eGDFlattenQueueOperation_Flatten
            m.strOperation = "Flatten"
        Case eGDFlattenQueueOperation_Reverse
            m.strOperation = "Reverse"
    End Select
End Property

Public Property Get AutoExitOnly() As Boolean
    AutoExitOnly = m.bAutoExitOnly
End Property
Public Property Let AutoExitOnly(ByVal bAutoExitOnly As Boolean)
    m.bAutoExitOnly = bAutoExitOnly
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Flatten
'' Description: Attempt to flatten the item
'' Inputs:      None
'' Returns:     True if done, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Flatten() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    ' We need to do the working order check...
    If m.bDoneWorkingCheck = False Then
        CheckForWorkingOrders
        bReturn = False
    
    ' The working order check timed out...
    ElseIf m.dWorkingCheckStart = -1# Then
        bReturn = True
    
    ' We need to submit the market order...
    ElseIf m.bDoneMarketOrder = False Then
        CheckForFlatPosition
        bReturn = False
    
    ' The market order timed out...
    ElseIf m.dMarketOrderStart = -1# Then
        bReturn = True
        
    ' Flatten is done...
    Else
        bReturn = True
    End If
    
    Flatten = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFlatten.Flatten"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reverse
'' Description: Attempt to reverse the item
'' Inputs:      None
'' Returns:     True if done, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Reverse() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    ' We need to do the working order check...
    If m.bDoneWorkingCheck = False Then
        CheckForWorkingOrders
        bReturn = False
    
    ' The working order check timed out...
    ElseIf m.dWorkingCheckStart = -1# Then
        bReturn = True
    
    ' We need to submit the market order...
    ElseIf m.bDoneMarketOrder = False Then
        CheckForReversedPosition
        bReturn = False
    
    ' The market order timed out...
    ElseIf m.dMarketOrderStart = -1# Then
        bReturn = True
        
    ' Flatten is done...
    Else
        bReturn = True
    End If
    
    Reverse = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFlatten.Reverse"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAll
'' Description: Attempt to cancel all orders for the given information
'' Inputs:      None
'' Returns:     True if done, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelAll() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    ' We need to do the working order check...
    If m.bDoneWorkingCheck = False Then
        CheckForWorkingOrders
        bReturn = False
    
    ' The working order check timed out...
    ElseIf m.dWorkingCheckStart = -1# Then
        bReturn = True
    
    ' Cancel all is done...
    Else
        bReturn = True
    End If
    
    CancelAll = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFlatten.CancelAll"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAllAutoExit
'' Description: Attempt to cancel all auto exit orders for the given information
'' Inputs:      None
'' Returns:     True if done, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelAllAutoExit() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    ' We need to do the working order check...
    If m.bDoneWorkingCheck = False Then
        CheckForAutoExitWorkingOrders
        bReturn = False
    
    ' The working order check timed out...
    ElseIf m.dWorkingCheckStart = -1# Then
        bReturn = True
    
    ' Cancel all is done...
    Else
        bReturn = True
    End If
    
    CancelAllAutoExit = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFlatten.CancelAllAutoExit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckForWorkingOrders
'' Description: Check for working orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckForWorkingOrders()
On Error GoTo ErrSection:

    Dim bHasWorkingOrders As Boolean    ' Do we have working orders?
    Dim strOperation As String          ' Operation to be done

    If m.dWorkingCheckStart = 0# Then
        g.Broker.BrokerDebug m.nBroker, m.strOperation & "(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Beginning Working Order Check"
        If frmMain.Enabled Then
            InfBox "Cancelling all working orders for " & m.strTradeSymbol & " in account " & m.strAccountNumber & ".||Waiting for verification from the exchange that the orders are cancelled.||Please Wait...", , , m.strOperation & " Position", True
        End If
        m.dWorkingCheckStart = gdTickCount
        
        g.Broker.CancelWorkingOrders AccountNumberOrID, TradeSymbolOrSymbolID, m.lAutoTradeItemID, m.lSkipOrderID
        If g.Broker.HasPendingOrders(AccountNumberOrID, TradeSymbolOrSymbolID, m.lAutoTradeItemID) Then
            m.dWorkingCheckTimeOut = 20000#
        Else
            m.dWorkingCheckTimeOut = 10000#
        End If
    Else
        bHasWorkingOrders = g.Broker.HasWorkingOrders(AccountNumberOrID, TradeSymbolOrSymbolID, m.lAutoTradeItemID, m.lSkipOrderID)
        If bHasWorkingOrders Then
            g.Broker.CancelWorkingOrders AccountNumberOrID, TradeSymbolOrSymbolID, m.lAutoTradeItemID, m.lSkipOrderID
        End If
        
        If (bHasWorkingOrders = False) Then
            g.Broker.BrokerDebug m.nBroker, m.strOperation & "(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): No Working Orders"
            If m.nOperation = eGDFlattenQueueOperation_CancelAll Then
                InfBox ""
            End If
            m.bDoneWorkingCheck = True
        ElseIf (gdTickCount > (m.dWorkingCheckStart + m.dWorkingCheckTimeOut)) Then
            g.Broker.BrokerDebug m.nBroker, m.strOperation & "(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Working Order Check Timed Out"
            InfBox "Timed out while waiting for working orders to be cancelled for " & m.strTradeSymbol & " in account '" & g.Broker.AccountNameForID(m.lAccountID) & "'", "!", , m.strOperation & " for " & m.strTradeSymbol
            m.bDoneWorkingCheck = True
            m.dWorkingCheckStart = -1#
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlatten.CheckForWorkingOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckForAutoExitWorkingOrders
'' Description: Check for auto exit working orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckForAutoExitWorkingOrders()
On Error GoTo ErrSection:

    Dim bHasWorkingOrders As Boolean    ' Do we have working orders?
    Dim AutoExit As cActiveExit         ' Auto exit object
    
    Set AutoExit = g.OrderStrategies.ExitObjectForAccountAndSymbol(m.lAccountID, SymbolOrSymbolID)
    If AutoExit Is Nothing Then
        g.Broker.BrokerDebug m.nBroker, "Auto Exit is Nothing"
        m.bDoneWorkingCheck = True
        InfBox ""
    Else
        If m.dWorkingCheckStart = 0# Then
            g.Broker.BrokerDebug m.nBroker, m.strOperation & "(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Beginning Auto Exit Working Order Check"
            If frmMain.Enabled Then
                InfBox "Cancelling all auto exit working orders for " & m.strTradeSymbol & " in account " & m.strAccountNumber & ".||Waiting for verification from the exchange that the orders are cancelled.||Please Wait...", , , m.strOperation & " Position", True
            End If
            m.dWorkingCheckStart = gdTickCount
            
            AutoExit.CancelWorkingOrders m.lSkipOrderID
            If AutoExit.HasPendingOrders Then
                m.dWorkingCheckTimeOut = 20000#
            Else
                m.dWorkingCheckTimeOut = 10000#
            End If
        Else
            bHasWorkingOrders = AutoExit.HasWorkingOrders(m.lSkipOrderID)
            If bHasWorkingOrders Then
                AutoExit.CancelWorkingOrders m.lSkipOrderID
            End If
            
            If (bHasWorkingOrders = False) Then
                g.Broker.BrokerDebug m.nBroker, m.strOperation & "(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): No Auto Exit Working Orders"
                If m.nOperation = eGDFlattenQueueOperation_CancelAll Then
                    InfBox ""
                End If
                m.bDoneWorkingCheck = True
            ElseIf (gdTickCount > (m.dWorkingCheckStart + m.dWorkingCheckTimeOut)) Then
                g.Broker.BrokerDebug m.nBroker, m.strOperation & "(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Auto Exit Working Order Check Timed Out"
                InfBox "Timed out while waiting for auto exit working orders to be cancelled for " & m.strTradeSymbol & " in account '" & g.Broker.AccountNameForID(m.lAccountID) & "'", "!", , m.strOperation & " for " & m.strTradeSymbol
                m.bDoneWorkingCheck = True
                m.dWorkingCheckStart = -1#
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlatten.CheckForAutoExitWorkingOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckForFlatPosition
'' Description: Check for flat position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckForFlatPosition()
On Error GoTo ErrSection:

    Dim lPosition As Long               ' Current Position
    Dim AutoExit As cActiveExit         ' Active exit item
    
    If m.dMarketOrderStart = 0# Then
        If g.Broker.IsPitSymbol(AccountNumberOrID, TradeSymbolOrSymbolID) Then
            g.Broker.BrokerDebug m.nBroker, "Flatten(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Aborting because this is a pit session symbol"
            InfBox "Trade Navigator has cancelled your working orders, but cannot flatten a position on a pit-session symbol.", , "!", "Flatten for " & m.strTradeSymbol
            m.bDoneMarketOrder = True
            m.dMarketOrderStart = -1#
        ElseIf g.Broker.PositionMatch(AccountNumberOrID, TradeSymbolOrSymbolID) = False Then
            g.Broker.BrokerDebug m.nBroker, "Flatten(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Aborting because of a position mismatch"
            InfBox "Because of inconsistent position data from the broker for " & m.strTradeSymbol & " in account '" & g.Broker.AccountNameForID(m.lAccountID) & "', we cannot flatten your position at this time.||PLEASE CALL YOUR BROKER AND VERIFY YOUR POSITION IN THIS ACCOUNT.|", "!", , "Flatten for " & m.strTradeSymbol
            m.bDoneMarketOrder = True
            m.dMarketOrderStart = -1#
        ElseIf g.Broker.FlattenPosition(AccountNumberOrID, TradeSymbolOrSymbolID, m.lAutoTradeItemID) = False Then
            InfBox ""
            m.bDoneMarketOrder = True
            m.dMarketOrderStart = -1#
        Else
            If frmMain.Enabled Then
                InfBox "Flattening position for " & m.strTradeSymbol & " in account " & m.strAccountNumber & ".||Waiting for verification from the exchange that the market order has been filled.||Please Wait...|", , , "Flatten Position", True
            End If
            m.dMarketOrderStart = gdTickCount
        End If
    Else
        lPosition = g.Broker.CurrentPosition(AccountNumberOrID, TradeSymbolOrSymbolID, m.lAutoTradeItemID)
        
        If (lPosition = 0) Then
            g.Broker.BrokerDebug m.nBroker, "Flatten(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Flat"
            InfBox ""
            m.bDoneMarketOrder = True
        ElseIf (gdTickCount > (m.dMarketOrderStart + 10000#)) Then
            g.Broker.BrokerDebug m.nBroker, "Flatten(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Flatten Attempt Timed Out"
            InfBox "Timed out while attempting to Flatten position for " & m.strTradeSymbol & " in account '" & g.Broker.AccountNameForID(m.lAccountID) & "'", "!", , "Flatten for " & m.strTradeSymbol
            m.bDoneMarketOrder = True
            m.dMarketOrderStart = -1#
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlatten.CheckForFlatPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckForReversedPosition
'' Description: Check for reversed position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckForReversedPosition()
On Error GoTo ErrSection:

    Dim lPosition As Long               ' Current Position
    Dim AutoExit As cActiveExit         ' Active exit item
    
    If m.dMarketOrderStart = 0# Then
        If g.Broker.IsPitSymbol(AccountNumberOrID, TradeSymbolOrSymbolID) Then
            g.Broker.BrokerDebug m.nBroker, "Reverse(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Aborting because this is a pit session symbol"
            InfBox "Trade Navigator has cancelled your working orders, but cannot reverse a position on a pit-session symbol.", , "!", "Reverse for " & m.strTradeSymbol
            m.bDoneMarketOrder = True
            m.dMarketOrderStart = -1#
        ElseIf g.Broker.PositionMatch(AccountNumberOrID, TradeSymbolOrSymbolID) = False Then
            g.Broker.BrokerDebug m.nBroker, "Reverse(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Aborting because of a position mismatch"
            InfBox "Because of inconsistent position data from the broker for " & m.strTradeSymbol & " in account '" & g.Broker.AccountNameForID(m.lAccountID) & "', we cannot reverse your position at this time.||PLEASE CALL YOUR BROKER AND VERIFY YOUR POSITION IN THIS ACCOUNT.|", "!", , "Reverse for " & m.strTradeSymbol
            m.bDoneMarketOrder = True
            m.dMarketOrderStart = -1#
        ElseIf g.Broker.ReversePosition(AccountNumberOrID, TradeSymbolOrSymbolID, m.lAutoTradeItemID) = False Then
            InfBox ""
            m.bDoneMarketOrder = True
            m.dMarketOrderStart = -1#
        Else
            m.lOldPosition = g.Broker.CurrentPosition(AccountNumberOrID, TradeSymbolOrSymbolID, m.lAutoTradeItemID)
            If frmMain.Enabled Then
                InfBox "Reversing position for " & m.strTradeSymbol & " in account " & m.strAccountNumber & ".||Waiting for verification from the exchange that the market order has been filled.||Please Wait...|", , , "Reverse Position", True
            End If
            m.dMarketOrderStart = gdTickCount
        End If
    Else
        lPosition = g.Broker.CurrentPosition(AccountNumberOrID, TradeSymbolOrSymbolID, m.lAutoTradeItemID)
        
        If (lPosition = m.lOldPosition * -1&) Then
            g.Broker.BrokerDebug m.nBroker, "Reverse(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Flat"
            InfBox ""
            m.bDoneMarketOrder = True
        ElseIf (gdTickCount > (m.dMarketOrderStart + 10000#)) Then
            g.Broker.BrokerDebug m.nBroker, "Reverse(" & m.strAccountNumber & ", " & m.strTradeSymbol & ", " & Str(m.lAutoTradeItemID) & "): Reverse Attempt Timed Out"
            InfBox "Timed out while attempting to Reverse position for " & m.strTradeSymbol & " in account '" & g.Broker.AccountNameForID(m.lAccountID) & "'", "!", , "Reverse for " & m.strTradeSymbol
            m.bDoneMarketOrder = True
            m.dMarketOrderStart = -1#
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlatten.CheckForReversedPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Intialize class members when object gets initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.lSkipOrderID = -1&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlatten.Class_Initialize"
    
End Sub
