VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCattleBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cCattleBridge.cls
'' Description: Bridge for the Cattle DLL
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/07/2014   DAJ         Created
'' 03/17/2014   DAJ         Renamed Turnkey to Cattle for admin stuff
'' 05/22/2014   DAJ         Renamed cTurnkey to cCattle; Renamed g.Turnkey to g.Cattle
'' 05/30/2014   DAJ         Added LoadAccountsCombo and CattleFillExists; Utilized
''                          new accounts object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Property Get AppBridge() As Object
    AppBridge = g.AppBridge.AppBridge
End Property
Public Property Let AppBridge(TheAppBridge As Object)
    g.AppBridge.AppBridge = TheAppBridge
End Property

Public Property Get AppPath() As String
    AppPath = g.strAppPath
End Property
Public Property Let AppPath(ByVal strAppPath As String)
    g.strAppPath = strAppPath
End Property

Public Property Get IniFile() As String
    IniFile = g.strIniFile
End Property
Public Property Let IniFile(ByVal strIniFile As String)
    g.strIniFile = strIniFile
End Property

Public Property Get DataServiceID() As Long
    DataServiceID = g.lDataServiceID
End Property
Public Property Let DataServiceID(ByVal lDataServiceID As Long)
    g.lDataServiceID = lDataServiceID
End Property

Public Property Get MainForm() As Object
    Set MainForm = g.frmMain
End Property
Public Property Let MainForm(frmMain As Object)
    Set g.frmMain = frmMain
End Property

Public Property Get AltGridRowColor() As Long
    AltGridRowColor = ALT_GRID_ROW_COLOR
End Property
Public Property Let AltGridRowColor(ByVal lAltGridRowColor As Long)
    ALT_GRID_ROW_COLOR = lAltGridRowColor
End Property

Public Property Get StreamInterval() As Long
    StreamInterval = g.lStreamInterval
End Property
Public Property Let StreamInterval(ByVal lStreamInterval As Long)
    g.lStreamInterval = lStreamInterval
End Property

Public Property Get StreamActive() As Boolean
    StreamActive = g.bStreamActive
End Property
Public Property Let StreamActive(ByVal bStreamActive As Boolean)
    g.bStreamActive = bStreamActive
End Property

Public Property Get StreamTimerEnabled() As Boolean
    If CattleFormLoaded Then
        StreamTimerEnabled = frmLots.tmrRealtime.Enabled
    End If
End Property
Public Property Let StreamTimerEnabled(bEnabled As Boolean)
    If CattleFormLoaded Then
        frmLots.tmrRealtime.Enabled = bEnabled
    End If
End Property

Public Property Get Help() As Object
    Set Help = g.Help
End Property
Public Property Let Help(AppHelp As Object)
    Set g.Help = AppHelp
End Property

Public Property Get ProductName() As String
    ProductName = g.Cattle.ProductName
End Property

Public Sub UpdateFeedYardCustomer()
    g.Cattle.UpdateFeedYardCustomer
End Sub

Public Sub SetUpNewOrder(ByVal strGenesisOrderID As String, ByVal strFeedYardLotID As String)
    If FormIsLoaded("frmLots") Then
        frmLots.SetUpNewOrder strGenesisOrderID, strFeedYardLotID
    End If
End Sub

Public Sub LoadLotsCombo(cboLots As Object, Optional ByVal strFeedYardLotID As String = "", Optional ByVal strInitialEntry As String = "")
    If Not g.Cattle Is Nothing Then
        g.Cattle.LoadLotsCombo cboLots, strFeedYardLotID, strInitialEntry
    End If
End Sub

Public Sub LoadAccountsCombo(cboAccounts As Object, Optional ByVal strAccountID As String = "")
    If Not g.Cattle Is Nothing Then
        g.Cattle.Accounts.LoadAccountsCombo cboAccounts, strAccountID
    End If
End Sub

Public Function Position(ByVal strFeedYardLotID As String, ByVal strAccountNumber As String, ByVal nBroker As eTT_AccountType, ByVal strGenesisSymbol As String) As Long
    If Not g.Cattle Is Nothing Then
        Position = g.Cattle.Position(strFeedYardLotID, strAccountNumber, nBroker, strGenesisSymbol)
    End If
End Function

Public Function IsCattleUser() As Boolean
    If Not g.Cattle Is Nothing Then
        IsCattleUser = g.Cattle.IsCattleUser
    End If
End Function

Public Function IsTurnkeyUser() As Boolean
    If Not g.Cattle Is Nothing Then
        IsTurnkeyUser = g.Cattle.IsTurnkeyUser
    End If
End Function

Public Function IsTurnkeyAdminUser() As Boolean
    If Not g.Cattle Is Nothing Then
        IsTurnkeyAdminUser = g.Cattle.IsTurnkeyAdminUser
    End If
End Function

Public Sub ShowCattleForm(ByVal bTurnkey As Boolean)
    frmLots.ShowMe bTurnkey
End Sub

Public Sub ShowCattleAdminForm()
    frmCattleAdmin.ShowMe
End Sub

Public Function CattleFormLoaded() As Boolean
    CattleFormLoaded = FormIsLoaded("frmLots")
End Function

Public Sub Broker_Order(Order As cBrokerMessage)
    If CattleFormLoaded Then
        mCattle.AddBrokerAccountNumberToOrder Order
        frmLots.Broker_Order Order
    End If
End Sub

Public Sub Broker_Fill(Fill As cBrokerMessage)
    If CattleFormLoaded Then
        mCattle.AddBrokerAccountNumberToFill Fill
        frmLots.Broker_Fill Fill
    End If
End Sub

Public Function ConfirmOrder(ByVal strBrokerAccountNumber As String, ByVal nBroker As eTT_AccountType) As Boolean
    If (CattleFormLoaded = True) And (Not g.Cattle Is Nothing) Then
        ConfirmOrder = g.Cattle.Accounts.AccountIdForBrokerNumber(strBrokerAccountNumber, nBroker)
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CattleFillExists
'' Description: Does a fill exist with the given Broker Fill ID?
'' Inputs:      Broker Fill ID, Account ID
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CattleFillExists(ByVal strBrokerFillID As String, ByVal strAccountID As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strBroker As String             ' Broker for the given account ID
    
    bReturn = False
    If (Not g.Cattle Is Nothing) And (CattleFormLoaded = True) Then
        strBroker = g.Cattle.Accounts.BrokerForAccountID(strAccountID)
        bReturn = frmLots.CattleFillExists(strBrokerFillID, strBroker)
    End If
    
    CattleFillExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleBridge.CattleFillExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the bridge and global variables for the DLL
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init()
On Error GoTo ErrSection:

    Load frmCattleAM
    Set g.Cattle = New cCattle
    Set g.BrokerEnums = New cBrokerEnums
    Set g.CattleEnums = New cCattleEnums
    Set g.CattleKeyValue = New cCattleKeyValue

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattleBridge.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Shutdown
'' Description: Clean up the bridge and global variables for the DLL
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Shutdown()
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.Disconnect
        g.Cattle.UnloadApp False
        Set g.Cattle = Nothing
    End If
    Set g.BrokerEnums = Nothing
    Set g.CattleEnums = Nothing
    Set g.CattleKeyValue = Nothing
    Unload frmCattleAM

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattleBridge.Shutdown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set g.AppBridge = New cCattleTn

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattleBridge.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Unload frmCattleAM

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattleBridge.Class_Terminate"
    
End Sub
