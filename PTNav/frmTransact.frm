VERSION 5.00
Object = "{9F97659F-2075-40AD-8155-E7C6B3936B34}#1.0#0"; "YesTradeEngineServer.dll"
Begin VB.Form frmTransact 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   300
   End
   Begin YESTRADEENGINESERVERLibCtl.YesTradeEngine TransactEngine 
      Height          =   2895
      Left            =   900
      TabIndex        =   0
      Top             =   180
      Width           =   2895
      _cx             =   1184754
      _cy             =   1184754
   End
End
Attribute VB_Name = "frmTransact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTransact.frm
'' Description: Object and routines to talk to the Transact servers
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kWaitSeconds = 10         ' Number of seconds to wait for response

Private Type mPrivate
    strAccountNumber As String          ' Account number for this account
    
    bConnected As Boolean               ' Is this form connected to Transact?
    bAccountSet As Boolean              ' Is the account set for this connection?
    
    nOrderInit As Long
    nTradeInit As Long
    nOrderCount As Long
    nTradeCount As Long
    
    SubscribedSyms As cGdTree           ' Collection of Subscribed contracts
End Type
Private m As mPrivate

Public Property Get AccountNumber() As String
    AccountNumber = m.strAccountNumber
End Property
Public Property Let AccountNumber(ByVal strAccountNumber As String)
    m.strAccountNumber = strAccountNumber
End Property

Public Property Get Orders() As YesTypes.OrderCollection
    Set Orders = TransactEngine.Orders(Account, Empty)
End Property

Public Property Get Trades() As YesTypes.OrderCollection
    Set Trades = TransactEngine.Trades(Account, Empty, Empty)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Connect
'' Description: Attempt to connect to the Transact servers
'' Inputs:      None
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Connect() As YesTypes.enumErrors
On Error GoTo ErrSection:

    DumpDebug "frmTransact.Connect(" & m.strAccountNumber & ")"

    If m.bConnected = False Then
        frmTTSummary.TransactStatus(m.strAccountNumber) = eGDBrokerConnect_InProcess
        TransactEngine.HostName = g.Transact.IPAddress
        Connect = TransactEngine.Logon(g.Transact.UserName, g.Transact.Password)
        If Connect <> YesTypes.enumErrors.ok Then
            frmTTSummary.TransactStatus(m.strAccountNumber) = eGDBrokerConnect_Disconnected
        End If
    Else
        Connect = YesTypes.enumErrors.ok
    End If

    DumpDebug "End frmTransact.Connect(" & m.strAccountNumber & ")"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransact.Connect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnsetAccount
'' Description: Unset the account for this connection
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UnsetAccount() As Boolean
On Error GoTo ErrSection:

    Dim lCounter As Long                ' Counter variable for the loop
    
    DumpDebug "frmTransact.UnsetAccount(" & m.strAccountNumber & ")"
    If m.bAccountSet = True Then TransactEngine.UnsetAccount
    Do While (m.bAccountSet = True) And (lCounter < kWaitSeconds)
        Sleep 1#
        lCounter = lCounter + 1
    Loop
    
    UnsetAccount = Not m.bAccountSet

    DumpDebug "End frmTransact.UnsetAccount(" & m.strAccountNumber & ")"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransact.UnsetAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Unsubscribe
'' Description: Unsubscribe to the symbols that were subscribed to
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Unsubscribe() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lCounter As Long                ' Counter variable in a loop
    
    DumpDebug "frmTransact.UnsubscribeAccount(" & m.strAccountNumber & ")"
    
    If Not m.SubscribedSyms Is Nothing Then
        For lIndex = m.SubscribedSyms.Count To 1 Step -1
            TransactEngine.Unsubscribe m.SubscribedSyms(lIndex)
        Next lIndex
        
        Do While (m.SubscribedSyms.Count > 0) And (lCounter < kWaitSeconds)
            Sleep 1#
        Loop
        
        Unsubscribe = (m.SubscribedSyms.Count = 0)
    Else
        Unsubscribe = True
    End If
    
    DumpDebug "End frmTransact.UnsetAccount(" & m.strAccountNumber & ")"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransact.Unsubscribe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Disconnect
'' Description: Attempt to disconnect from the Transact servers
'' Inputs:      None
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Disconnect() As Boolean
On Error GoTo ErrSection:

    Dim lCounter As Long                ' Counter variable in a loop

    DumpDebug "frmTransact.Disconnect(" & m.strAccountNumber & ")"

    If m.bConnected = True Then TransactEngine.LogOff
    
    Do While (m.bConnected = True) And (lCounter < kWaitSeconds)
        Sleep 1#
    Loop
    Disconnect = Not m.bConnected

    DumpDebug "End frmTransact.Disconnect(" & m.strAccountNumber & ")"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransact.Disconnect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ErrorString
'' Description: Get the error string from the engine for the given error code
'' Inputs:      Error Code
'' Returns:     Error String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ErrorString(ByVal nError As YesTypes.enumErrors) As String
On Error GoTo ErrSection:

    ErrorString = TransactEngine.GetErrorString(nError)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransact.ErrorString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Subscribe
'' Description: Subscribe to a specific Genesis symbol if not already
'' Inputs:      Genesis symbol
'' Returns:     True if can subscribe, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Subscribe(ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    If m.SubscribedSyms.Exists(strGenesisSymbol) = False Then
        If TransactEngine.Subscribe(g.Transact.TransactContract(strGenesisSymbol)) = YesTypes.enumErrors.ok Then
            Subscribe = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransact.Subscribe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Initialize
'' Description: Initialize things when the form is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Initialize()
On Error GoTo ErrSection:

    m.strAccountNumber = ""
    m.bConnected = False
    m.bAccountSet = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.Form_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize things when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Set m.SubscribedSyms = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up things when the form is unloaded
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.SubscribedSyms = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTransact.Form_Unload"
    
End Sub

Private Sub tmrInit_Timer()

    Dim Acct As YesTypes.Account

    If g.bUnloading Then
        tmrInit.Enabled = False
        Exit Sub
    End If

    Set Acct = Account
    If m.nOrderInit <= m.nOrderCount Then
If IsIDE Then StatusMsg "Order " & Str(m.nOrderInit)
        g.Transact.HandleOrderCallback "GetOrders", TransactEngine.Orders(Acct, Empty).Item(m.nOrderInit), YesTypes.enumErrors.ok
        m.nOrderInit = m.nOrderInit + 1
    ElseIf m.nTradeInit <= m.nTradeCount Then
If IsIDE And (m.nTradeInit Mod 100 = 0) Then StatusMsg "Trade " & Str(m.nTradeInit)
        g.Transact.HandleOrderCallback "GetTrades", TransactEngine.Trades(Acct).Item(m.nTradeInit), YesTypes.enumErrors.ok
        m.nTradeInit = m.nTradeInit + 1
    Else
If IsIDE Then StatusMsg "done " & Str(m.nTradeCount + m.nOrderCount)
        tmrInit.Enabled = False
        UpdateVisibleCharts eRedo1_Scrolled
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_CanceledOrder
'' Description: Callback notifying us that an order has been cancelled
'' Inputs:      Transact Order, Error Code
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_CanceledOrder(ByVal CanceledOrder As YESTRADEENGINESERVERLibCtl.IOrder, ByVal errorCode As YESTRADEENGINESERVERLibCtl.enumErrors)
On Error GoTo ErrSection:

    g.Transact.HandleOrderCallback "CancelOrder", CanceledOrder, errorCode

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_CanceledOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_ChangedOrder
'' Description: Callback notifying us that an order has changed
'' Inputs:      Transact Order, Error Code
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_ChangedOrder(ByVal updatedOrder As YESTRADEENGINESERVERLibCtl.IOrder, ByVal errorCode As YESTRADEENGINESERVERLibCtl.enumErrors)
On Error GoTo ErrSection:

    g.Transact.HandleOrderCallback "ChangedOrder", updatedOrder, errorCode

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_ChangedOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_Error
'' Description: Callback notifying us that an error has occurred
'' Inputs:      Error Code
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_Error(ByVal Error As YESTRADEENGINESERVERLibCtl.enumErrors)
On Error GoTo ErrSection:

    DumpDebug "frmTransact.Error(" & m.strAccountNumber & ", " & ErrorString(Error) & ")"

    If Error = YesTypes.enumErrors.errConnectionLost Then
        frmTTSummary.TransactStatus(m.strAccountNumber) = eGDBrokerConnect_Disconnected
        m.bConnected = False
    Else
        InfBox TransactEngine.GetErrorString(Error), "!", , "Transact Error"
    End If
    
    DumpDebug "End frmTransact.Error(" & m.strAccountNumber & ")"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_Error"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_FilledOrder
'' Description: Callback notifying us that an order has been filled
'' Inputs:      Transact Order, Position, Average Fill Price, Error Code
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_FilledOrder(ByVal FilledOrder As YESTRADEENGINESERVERLibCtl.IOrder, ByVal Position As Long, ByVal averagePrice As String, ByVal errorCode As YESTRADEENGINESERVERLibCtl.enumErrors)
On Error GoTo ErrSection:

    DumpDebug "TransactEngine_FilledOrder --> " & g.Transact.GenesisSymbol(FilledOrder.contract) & ", " & Str(Position) & ", " & averagePrice & ", " & TransactEngine.GetErrorString(errorCode)
    g.Transact.HandleOrderCallback "FilledOrder", FilledOrder, errorCode

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_FilledOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_LoggedOn
'' Description: Callback notifying us that the user is logged on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_LoggedOn()
On Error GoTo ErrSection:

gdStartProfile 900

    Dim lIndex As Long                  ' Index into a for loop
    Dim nReturn As YesTypes.enumErrors  ' Return from the SetAccount call

    DumpDebug "frmTransact.LoggedON(" & m.strAccountNumber & ")"

    m.bConnected = True
    frmTTSummary.TransactStatus(m.strAccountNumber) = eGDBrokerConnect_Connected
    GetAccounts
    nReturn = TransactEngine.SetAccount(Account)
    
    DumpDebug "End frmTransact.LoggedON(" & m.strAccountNumber & ")"

gdStopProfile 900

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_LoggedOn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_OrderExpired
'' Description: Callback notifying us that an order has expired
'' Inputs:      Transact Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_OrderExpired(ByVal expiredOrder As YESTRADEENGINESERVERLibCtl.IOrder)
On Error GoTo ErrSection:

    g.Transact.HandleOrderCallback "OrderExpired", expiredOrder, ok

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_OrderExpired"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_PriceChanged
'' Description: Callback notifying us that a price has changed
'' Inputs:      Changed Price
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_PriceChanged(ByVal changedPrice As YESTRADEENGINESERVERLibCtl.IPrice)
On Error GoTo ErrSection:

#If 0 Then
    Dim strGenesisSymbol As String      ' Genesis symbol
    
    g.Transact.UpdatePrice changedPrice
    
    strGenesisSymbol = g.Transact.GenesisSymbol(changedPrice.contract)
    If FormIsLoaded("frmTTPositions") Then
        frmTTPositions.RefreshTransactPrices strGenesisSymbol, changedPrice
    End If
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_PriceChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_SentOrder
'' Description: Callback notifying us that an order has been sent
'' Inputs:      Transact Order, Error Code
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_SentOrder(ByVal SentOrder As YESTRADEENGINESERVERLibCtl.IOrder, ByVal errorCode As YESTRADEENGINESERVERLibCtl.enumErrors)
On Error GoTo ErrSection:

    g.Transact.HandleOrderCallback "SentOrder", SentOrder, errorCode

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_SentOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_SetAccountResponse
'' Description: Callback notifying us that the account has been set
'' Inputs:      Account ID, Error Code
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_SetAccountResponse(ByVal AccountID As Long, ByVal errorCode As YESTRADEENGINESERVERLibCtl.enumErrors)
On Error GoTo ErrSection:

    Dim Acct As YesTypes.Account

gdStartProfile 920

    DumpDebug "frmTransact.AccountSet(" & m.strAccountNumber & ")"

    m.bAccountSet = True
       
    GetContracts
    
    Set Acct = Account
    m.nOrderCount = TransactEngine.Orders(Acct, Empty).Count
    m.nTradeCount = TransactEngine.Trades(Acct, Empty).Count
    m.nOrderInit = 1
    m.nTradeInit = 1
    tmrInit.Enabled = True
    'GetOrders
    'GetTrades
    
    DumpDebug "End frmTransact.AccountSet(" & m.strAccountNumber & ")"

gdStopProfile 920

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_SetAccountResponse"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_Subscribed
'' Description: Callback notifying us that a contract has been subscribed to
'' Inputs:      Contract, Error Code
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_Subscribed(ByVal subbedContract As YESTRADEENGINESERVERLibCtl.IContract, ByVal errorCode As YESTRADEENGINESERVERLibCtl.enumErrors)
On Error GoTo ErrSection:

    DumpDebug "TransactEngine_Subscribed --> " & subbedContract.Name
    m.SubscribedSyms.Add subbedContract, g.Transact.GenesisSymbol(subbedContract)
    If FormIsLoaded("frmTTPositions") Then
        frmTTPositions.AddToQuotes g.Transact.GenesisSymbol(subbedContract), subbedContract.Name
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_Subscribed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_UnsetAccountResponse
'' Description: Callback notifying us the account has been unset
'' Inputs:      Contract, Error Codes
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_UnsetAccountResponse(ByVal errorCode As YESTRADEENGINESERVERLibCtl.enumErrors)
On Error GoTo ErrSection:

    DumpDebug "frmTransact.AccountUnset(" & m.strAccountNumber & ")"

    m.bAccountSet = False
    'TransactEngine.LogOff

    DumpDebug "End frmTransact.AccountUnset(" & m.strAccountNumber & ")"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_UnsetAccountResponse"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransactEngine_Unsubscribed
'' Description: Callback notifying us that a contract has been unsubscribed
'' Inputs:      Contract, Error Code
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransactEngine_Unsubscribed(ByVal unsubbedContract As YESTRADEENGINESERVERLibCtl.IContract, ByVal errorCode As YESTRADEENGINESERVERLibCtl.enumErrors)
On Error GoTo ErrSection:

    DumpDebug "TransactEngine_Unsubscribed --> " & unsubbedContract.Name
    m.SubscribedSyms.Remove g.Transact.GenesisSymbol(unsubbedContract)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.TransactEngine_Unsubscribed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Request all account information from the Transact server
'' Inputs:      None
'' Returns:     True if new account added, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetAccounts() As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID for the account
    Dim bNew As Boolean                 ' Is this a new account?
    Dim bReturn As Boolean              ' Return value from the function
    Dim rs2 As Recordset                ' Recordset into the database

gdStartProfile 910
gdStartProfile 911

    bReturn = False
    'InfBox "Refreshing Transact Account Information...", , , , True
    DumpDebug "Transact Accounts...(" & Str(TransactEngine.Accounts.Count) & ")"
    For lIndex = 1 To TransactEngine.Accounts.Count
        DumpDebug Str(TransactEngine.Accounts(lIndex).AccountID) & vbTab & TransactEngine.Accounts(lIndex).Name
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountNumber]='" & Str(TransactEngine.Accounts(lIndex).AccountID) & "';", dbOpenDynaset)
        If rs.BOF And rs.EOF Then
            rs.AddNew
            rs!AccountNumber = Str(TransactEngine.Accounts(lIndex).AccountID)
            rs!AccountType = eTT_AccountType_Transact
            rs!StartingDate = Date
            rs!SecTypeMask = 1
            rs!Name = TransactEngine.Accounts(lIndex).Name
            lAccountID = rs!AccountID
            rs.Update
            
            g.Transact.AddForm Str(TransactEngine.Accounts(lIndex).AccountID)
            If FormIsLoaded("frmTTSummary") Then
                frmTTSummary.RefreshAccount lAccountID
                frmTTSummary.TransactStatus(Str(TransactEngine.Accounts(lIndex).AccountID)) = eGDBrokerConnect_Connected
            End If
            
            bReturn = True
        End If
    Next lIndex
        
gdStopProfile 911
gdStartProfile 912

    If bReturn = True Then
        Set rs2 = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                "WHERE [AccountNumber]='TransactPlaceholder';", dbOpenDynaset)
        If Not (rs2.BOF And rs2.EOF) Then
            lAccountID = rs2!AccountID
            rs2.Delete
            If FormIsLoaded("frmTTSummary") Then frmTTSummary.RefreshAccount lAccountID
        End If
    End If
    
    GetAccounts = bReturn

gdStopProfile 912
gdStopProfile 900

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransact.GetAccounts"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContracts
'' Description: Request available contracts for an account
'' Inputs:      Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetContracts()
On Error GoTo ErrSection:

    Dim Acct As YesTypes.Account        ' Account to get contracts for
    Dim lContract As Long               ' Index into a for loop
    Dim strGenesisSymbol As String      ' Genesis symbol for the Transact contract
    Dim nReturn As YesTypes.enumErrors  ' Return from subscribe call
    
gdStartProfile 930
    
    DumpDebug "frmTransact.GetContracts(" & m.strAccountNumber & ")"

    Set Acct = Account
    If Not Acct Is Nothing Then
        For lContract = 1 To Acct.AccountContracts.Count
            With Acct.AccountContracts(lContract).contract
                DumpDebug Str(.ID) & ";" & .Symbol & ";" & Str(.Month) & ";" & Str(.Year) & ";" & Str(.ExchangeID) & ";" & Str(.TickValue) & ";" & Str(Acct.AccountContracts(lContract).PreviousClose) & ";" & Str(Acct.AccountContracts(lContract).PreviousSettle) & ";" & Str(.priceformat)
                
                strGenesisSymbol = g.Transact.GenesisSymbol(Acct.AccountContracts(lContract).contract)
                If Len(strGenesisSymbol) > 0 Then
                    If Not g.Transact.Contracts.Exists(strGenesisSymbol) Then
                        g.Transact.Contracts.Add Acct.AccountContracts(lContract).contract, strGenesisSymbol
                        
                        nReturn = TransactEngine.Subscribe(Acct.AccountContracts(lContract).contract)
                        DumpDebug "Subscribing to " & Acct.AccountContracts(lContract).contract.Name & ": " & Str(nReturn)
                        If nReturn = YesTypes.enumErrors.ok Then g.RealTime.AddBrokerRtSymbol strGenesisSymbol, "Transact"
                    End If
                End If
            End With
        Next lContract
    End If
    
    DumpDebug "End frmTransact.GetContracts(" & m.strAccountNumber & ")"

gdStopProfile 930

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.GetContracts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Request orders for an account
'' Inputs:      Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetOrders()
On Error GoTo ErrSection:

    Dim Acct As YesTypes.Account        ' Account to get contracts for
    Dim lOrder As Long                  ' Index into a for loop
    
gdStartProfile 940
    
    DumpDebug "frmTransact.GetOrders(" & m.strAccountNumber & ")"

    Set Acct = Account
    DumpDebug vbTab & "Number for " & m.strAccountNumber & " = " & TransactEngine.Orders(Acct, Empty).Count
    For lOrder = 1 To TransactEngine.Orders(Acct, Empty).Count
        g.Transact.HandleOrderCallback "GetOrders", TransactEngine.Orders(Acct, Empty).Item(lOrder), YesTypes.enumErrors.ok
        DoEvents
    Next lOrder
    
    If TransactEngine.Orders(Acct, Empty).Count > 0 Then
        UpdateVisibleCharts eRedo1_Scrolled ', Order.SymbolID
    End If
    
    DumpDebug "End frmTransact.GetOrders(" & m.strAccountNumber & ")"

gdStopProfile 940

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTrades
'' Description: Request trades for an account
'' Inputs:      Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetTrades()
On Error GoTo ErrSection:

    Dim Acct As YesTypes.Account        ' Account to get contracts for
    Dim lTrade As Long                  ' Index into a for loop
    
gdStartProfile 950
    
    DumpDebug "frmTransact.GetTrades(" & m.strAccountNumber & ")"

    Set Acct = Account
    DumpDebug vbTab & "Number for " & m.strAccountNumber & " = " & TransactEngine.Trades(Acct).Count
    For lTrade = 1 To TransactEngine.Trades(Acct).Count
        g.Transact.HandleOrderCallback "GetTrades", TransactEngine.Trades(Acct).Item(lTrade), YesTypes.enumErrors.ok
        DoEvents
    Next lTrade
    
    If TransactEngine.Trades(Acct).Count > 0 Then
        UpdateVisibleCharts eRedo1_Scrolled ', Order.SymbolID
    End If
    
    DumpDebug "End frmTransact.GetTrades(" & m.strAccountNumber & ")"

gdStopProfile 950

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.GetTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Account
'' Description: Get the account object for the account number
'' Inputs:      None
'' Returns:     Account Object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Account() As YesTypes.Account
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With TransactEngine
        For lIndex = 1 To .Accounts.Count
            If Str(.Accounts(lIndex).AccountID) = m.strAccountNumber Then
                Set Account = .Accounts(lIndex)
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransact.Account"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send the given string to the test form and the debug log
'' Inputs:      String to Send
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strDebug As String)
On Error GoTo ErrSection:

    frmTest2.AddList strDebug
    DebugLog strDebug

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransact.DumpDebug"
    
End Sub

