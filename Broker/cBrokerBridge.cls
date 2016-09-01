VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerBridge.cls
'' Description: Bridge for the Broker DLL
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
'' 09/11/2014   DAJ         Added database objects; Added Core Bridge
'' 10/24/2014   DAJ         Core Application functions for DLL's; Database Objects
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Property Get AppBridge() As Object
    Set AppBridge = g.AppBridge.AppBridge
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

Public Property Get TnCore() As Object
    Set TnCore = g.TnCore.CoreBridge
End Property
Public Property Let TnCore(TheTnCore As Object)
    g.TnCore.CoreBridge = TheTnCore
End Property

Public Property Get TradingDatabase() As Database
    Set TradingDatabase = g.TradeTrackerDB.Database
End Property
Public Property Let TradingDatabase(TradeTrackerDatabase As Database)
    g.TradeTrackerDB.Database = TradeTrackerDatabase
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsSimulatedAccount
'' Description: Determine if the given account type is a simulated account
'' Inputs:      Account Type
'' Returns:     True if a simulated account, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsSimulatedAccount(ByVal nAccountType As eTT_AccountType) As Boolean
    IsSimulatedAccount = mBroker.IsSimulatedAccount(nAccountType)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TypeOfAccount
'' Description: Determine the type of the given account
'' Inputs:      Account Type
'' Returns:     Type of Account (Simulated, Broker Live, Broker Demo)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TypeOfAccount(ByVal nAccountType As eTT_AccountType) As eGDTypeOfAccount
    TypeOfAccount = mBroker.TypeOfAccount(nAccountType)
End Function

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

    Set g.AppBridge = New cBrokerTn
    Set g.BrokerDB = New cBrokerDb
    Set g.TnCore = New cCoreTn
    Set g.TradeTrackerDB = New cTradeTrackerDb

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cBrokerBridge.Class_Initialize"
    
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

    Set g.BrokerDB = Nothing
    Set g.TradeTrackerDB = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cBrokerBridge.Class_Terminate"
    
End Sub
