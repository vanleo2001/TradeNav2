VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerEnums.cls
'' Description: Enumerations for the Broker DLL
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
'' 08/22/2014   DAJ         Added E-Trade to the account type enum
'' 10/24/2014   DAJ         Core Application functions for DLL's; Moved TypeOfAccount
''                          and FillMatchMode enums out of NavSuite into NavBroker.DLL
'' 10/31/2014   DAJ         Added Market-On-Open order type
'' 09/14/2015   DAJ         Added Tradier broker
'' 03/18/2016   DAJ         Added TD Ameritrade broker
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eGDConnectionStatus
    eGDConnectionStatus_Disconnected = 0
    eGDConnectionStatus_Disconnecting
    eGDConnectionStatus_Connecting
    eGDConnectionStatus_Connected
End Enum

Public Enum eTT_AccountType
    eTT_AccountType_Standard = 0
    eTT_AccountType_PATS = 1
    eTT_AccountType_SimTrade = 2
    'eTT_AccountType_Photon = 3
    eTT_AccountType_IntBrokers = 4
    'eTT_AccountType_LindWaldock = 5
    'eTT_AccountType_ManLondon = 6
    'eTT_AccountType_ManChicago = 7
    'eTT_AccountType_Alaron = 8
    eTT_AccountType_TransAct = 9
    'eTT_AccountType_PFG = 10
    'eTT_AccountType_FXCM = 11
    eTT_AccountType_TT = 12
    eTT_AccountType_AdvFut = 13
    'eTT_AccountType_Gain = 14
    'eTT_AccountType_ManExpress = 15
    'eTT_AccountType_Rosenthal = 16
    'eTT_AccountType_Robbins = 17
    eTT_AccountType_Ideal = 18
    'eTT_AccountType_Cadent = 19
    'eTT_AccountType_Lotus = 20
    eTT_AccountType_Rithmic = 21
    eTT_AccountType_Vision = 22
    eTT_AccountType_Optimus = 23
    eTT_AccountType_OpVest = 24
    eTT_AccountType_ZenFire = 25
    eTT_AccountType_Oec = 26
    'eTT_AccountType_OptionsXpress = 27
    eTT_AccountType_SimBroker = 28
    eTT_AccountType_SimStream = 29
    eTT_AccountType_SimReplay = 30
    eTT_AccountType_CQG = 31
    eTT_AccountType_AmpCqg = 32
    eTT_AccountType_RjoCqg = 33
    eTT_AccountType_RjoPats = 34
    'eTT_AccountType_OptionsHouse = 35
    eTT_AccountType_Gft = 36
    eTT_AccountType_CtgCqg = 37
    eTT_AccountType_CtgPats = 38
    'eTT_AccountType_CtgPfg = 39
    'eTT_AccountType_FintecPfg = 40
    eTT_AccountType_AlpariCurrenex = 41
    eTT_AccountType_AlpariPats = 42
    eTT_AccountType_KnightCurrenex = 43
    eTT_AccountType_KnightCqg = 44
    eTT_AccountType_ZanerCqg = 45
    eTT_AccountType_ZanerPats = 46
    eTT_AccountType_ZanerRithmic = 47
    eTT_AccountType_ZanerZenFire = 48
    eTT_AccountType_AlpariZenFire = 49
    eTT_AccountType_RobbinsCqg = 50
    eTT_AccountType_RcgPats = 51
    eTT_AccountType_DemoPats = 52
    eTT_AccountType_RjoHkPats = 53
    eTT_AccountType_BornPats = 54
    eTT_AccountType_ZanerCurrenex = 55
    eTT_AccountType_Currenex = 56
    eTT_AccountType_FxddCurrenex = 57
    eTT_AccountType_VanKarCurrenex = 58
    eTT_AccountType_VisionCqg = 59
    eTT_AccountType_FptOec = 60
    eTT_AccountType_FptCqg = 61
    eTT_AccountType_Etrade = 62
    eTT_AccountType_Tradier = 63
    eTT_AccountType_Ameritrade = 64
End Enum

Public Enum eTT_OrderStatus
    eTT_OrderStatus_Open = 0
    eTT_OrderStatus_Partial = 1
    eTT_OrderStatus_Filled = 2
    eTT_OrderStatus_Cancelled = 3
    eTT_OrderStatus_Queued = 4
    eTT_OrderStatus_Sent = 5
    eTT_OrderStatus_Working = 6
    eTT_OrderStatus_Rejected = 7
    eTT_OrderStatus_BalCancelled = 8
    eTT_OrderStatus_CancelPending = 9
    eTT_OrderStatus_AmendPending = 10
    eTT_OrderStatus_UnconfirmedFilled = 11
    eTT_OrderStatus_UnconfirmedPartial = 12
    eTT_OrderStatus_Held = 13
    eTT_OrderStatus_CancelHeld = 14
    eTT_OrderStatus_Error = 15
    eTT_OrderStatus_Amended = 16
    eTT_OrderStatus_Expired = 17
    eTT_OrderStatus_Frozen = 18
    eTT_OrderStatus_ParkPending = 19
    eTT_OrderStatus_TriggerPending = 20
    eTT_OrderStatus_Approved = 21
    eTT_OrderStatus_BrokerParked = 22
    eTT_OrderStatus_OverFilled = 23
    eTT_OrderStatus_Parked = 24
    eTT_OrderStatus_PreSubmitted = 25
    eTT_OrderStatus_Inactive = 26
    eTT_OrderStatus_Suspended = 27
    eTT_OrderStatus_DataPending = 28
End Enum

Public Enum eTT_OrderType
    eTT_OrderType_Market = 0
    eTT_OrderType_Stop = 1
    eTT_OrderType_Limit = 2
    eTT_OrderType_StopWithLimit = 3
    
    eTT_OrderType_MarketOnClose = 4
    eTT_OrderType_StopCloseOnly = 5
    eTT_OrderType_LimitCloseOnly = 6
    eTT_OrderType_StopWithLimitCloseOnly = 7
    
    eTT_OrderType_MIT = 8
    eTT_OrderType_MarketOnOpen = 9
    
    eTT_OrderType_Adjustment = 10
End Enum

Public Enum eTT_TimeInForce
    eTT_TimeInForce_Day = 0
    eTT_TimeInForce_GTC = 1
    eTT_TimeInForce_GTD = 2
End Enum

Public Enum eGDTypeOfAccount
    eGDTypeOfAccount_Simulated = 0
    eGDTypeOfAccount_BrokerLive = 1
    eGDTypeOfAccount_BrokerDemo = 2
End Enum

Public Enum eTT_FillMatchMode
    eTT_FillMatchMode_Fifo = 0
    eTT_FillMatchMode_Lifo = 1
End Enum

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectionStatusToString
'' Description: Determine the displayable string for the given connection status
'' Inputs:      Connection Status
'' Returns:     Display Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConnectionStatusToString(ByVal nConnectionStatus As eGDConnectionStatus) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nConnectionStatus
        Case eGDConnectionStatus_Disconnected
            strReturn = "Disconnected"
        Case eGDConnectionStatus_Disconnecting
            strReturn = "Disconnecting"
        Case eGDConnectionStatus_Connecting
            strReturn = "Connecting"
        Case eGDConnectionStatus_Connected
            strReturn = "Connected"
    End Select
    
    ConnectionStatusToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.ConnectionStatusToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderStatusToString
'' Description: Determine the string description of the given order status
'' Inputs:      Order Status
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderStatusToString(ByVal nStatus As eTT_OrderStatus) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nStatus
        Case eTT_OrderStatus_Open
            strReturn = "New"
        Case eTT_OrderStatus_Partial
            strReturn = "Partial Fill"
        Case eTT_OrderStatus_Filled
            strReturn = "Filled"
        Case eTT_OrderStatus_Cancelled
            strReturn = "Cancelled"
        Case eTT_OrderStatus_Queued
            strReturn = "Queued"
        Case eTT_OrderStatus_Sent
            strReturn = "Sent"
        Case eTT_OrderStatus_Working
            strReturn = "Working"
        Case eTT_OrderStatus_Rejected
            strReturn = "Rejected"
        Case eTT_OrderStatus_BalCancelled
            strReturn = "Balance Cancelled"
        Case eTT_OrderStatus_CancelPending
            strReturn = "Cancel Pending"
        Case eTT_OrderStatus_AmendPending
            strReturn = "Amend Pending"
        Case eTT_OrderStatus_UnconfirmedFilled
            strReturn = "Unconfirmed Filled"
        Case eTT_OrderStatus_UnconfirmedPartial
            strReturn = "Unconfirmed Partial Fill"
        Case eTT_OrderStatus_Held
            strReturn = "Held"
        Case eTT_OrderStatus_CancelHeld
            strReturn = "Cancel Held"
        Case eTT_OrderStatus_Error
            strReturn = "Error"
        Case eTT_OrderStatus_Amended
            strReturn = "Amended"
        Case eTT_OrderStatus_Expired
            strReturn = "Expired"
        Case eTT_OrderStatus_Frozen
            strReturn = "Frozen"
        Case eTT_OrderStatus_ParkPending
            strReturn = "Park Pending"
        Case eTT_OrderStatus_TriggerPending
            strReturn = "Pending Trigger"
        Case eTT_OrderStatus_Approved
            strReturn = "Approved"
        Case eTT_OrderStatus_BrokerParked
            strReturn = "Parked at Broker"
        Case eTT_OrderStatus_OverFilled
            strReturn = "Over Filled"
        Case eTT_OrderStatus_Parked
            strReturn = "Parked"
        Case eTT_OrderStatus_PreSubmitted
            strReturn = "PreSubmitted"
        Case eTT_OrderStatus_Inactive
            strReturn = "Inactive"
        Case eTT_OrderStatus_Suspended
            strReturn = "Suspended"
        Case eTT_OrderStatus_DataPending
            strReturn = "Pending Data"
    End Select
    
    OrderStatusToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.OrderStatusToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderStatusFromString
'' Description: Determine the order status given the string description
'' Inputs:      String description
'' Returns:     Order Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderStatusFromString(ByVal strOrderStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderStatus      ' Return value for the function

    Select Case UCase(strOrderStatus)
        Case "OPEN", "NEW"
            nReturn = eTT_OrderStatus_Open
        Case "PARTIAL FILL"
            nReturn = eTT_OrderStatus_Partial
        Case "FILLED"
            nReturn = eTT_OrderStatus_Filled
        Case "CANCELLED"
            nReturn = eTT_OrderStatus_Cancelled
        Case "QUEUED"
            nReturn = eTT_OrderStatus_Queued
        Case "SENT"
            nReturn = eTT_OrderStatus_Sent
        Case "WORKING"
            nReturn = eTT_OrderStatus_Working
        Case "REJECTED"
            nReturn = eTT_OrderStatus_Rejected
        Case "BALANCE CANCELLED"
            nReturn = eTT_OrderStatus_BalCancelled
        Case "CANCEL PENDING"
            nReturn = eTT_OrderStatus_CancelPending
        Case "AMEND PENDING"
            nReturn = eTT_OrderStatus_AmendPending
        Case "UNCONFIRMED FILLED"
            nReturn = eTT_OrderStatus_UnconfirmedFilled
        Case "UNCONFIRMED PARTIAL FILL"
            nReturn = eTT_OrderStatus_UnconfirmedPartial
        Case "HELD"
            nReturn = eTT_OrderStatus_Held
        Case "CANCEL HELD"
            nReturn = eTT_OrderStatus_CancelHeld
        Case "ERROR"
            nReturn = eTT_OrderStatus_Error
        Case "AMENDED"
            nReturn = eTT_OrderStatus_Amended
        Case "EXPIRED"
            nReturn = eTT_OrderStatus_Expired
        Case "FROZEN"
            nReturn = eTT_OrderStatus_Frozen
        Case "PARK PENDING"
            nReturn = eTT_OrderStatus_ParkPending
        Case "PENDING TRIGGER"
            nReturn = eTT_OrderStatus_TriggerPending
        Case "APPROVED"
            nReturn = eTT_OrderStatus_Approved
        Case "PARKED AT BROKER"
            nReturn = eTT_OrderStatus_BrokerParked
        Case "OVER FILLED"
            nReturn = eTT_OrderStatus_OverFilled
        Case "PARKED"
            nReturn = eTT_OrderStatus_Parked
        Case "PRESUBMITTED"
            nReturn = eTT_OrderStatus_PreSubmitted
        Case "INACTIVE"
            nReturn = eTT_OrderStatus_Inactive
        Case "SUSPENDED"
            nReturn = eTT_OrderStatus_Suspended
        Case "PENDING DATA"
            nReturn = eTT_OrderStatus_DataPending
    End Select
    
    OrderStatusFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.OrderStatusFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeToString
'' Description: Determine the string description of the given order type
'' Inputs:      Order Type
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderTypeToString(ByVal eOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case eOrderType
        Case eTT_OrderType_Market
            strReturn = "Market"
        Case eTT_OrderType_Stop
            strReturn = "Stop"
        Case eTT_OrderType_Limit
            strReturn = "Limit"
        Case eTT_OrderType_StopWithLimit
            strReturn = "Stop With Limit"
        
        Case eTT_OrderType_MarketOnClose
            strReturn = "Market On Close"
        Case eTT_OrderType_StopCloseOnly
            strReturn = "Stop Close Only"
        Case eTT_OrderType_LimitCloseOnly
            strReturn = "Limit Close Only"
        Case eTT_OrderType_StopWithLimitCloseOnly
            strReturn = "Stop With Limit Close Only"
        
        Case eTT_OrderType_MIT
            strReturn = "MIT"
        Case eTT_OrderType_MarketOnOpen
            strReturn = "Market On Open"
        
        Case eTT_OrderType_Adjustment
            strReturn = "Other"
    End Select
    
    OrderTypeToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.OrderTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeFromString
'' Description: Determine the order type given the string description
'' Inputs:      String description
'' Returns:     Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderTypeFromString(ByVal strOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderType        ' Return value for the function

    Select Case UCase(strOrderType)
        Case "MARKET"
            nReturn = eTT_OrderType_Market
        Case "STOP"
            nReturn = eTT_OrderType_Stop
        Case "LIMIT"
            nReturn = eTT_OrderType_Limit
        Case "STOP WITH LIMIT"
            nReturn = eTT_OrderType_StopWithLimit
        
        Case "MARKET ON CLOSE"
            nReturn = eTT_OrderType_MarketOnClose
        Case "STOP CLOSE ONLY"
            nReturn = eTT_OrderType_StopCloseOnly
        Case "LIMIT ON CLOSE"
            nReturn = eTT_OrderType_LimitCloseOnly
        Case "STOP WITH LIMIT CLOSE ONLY"
            nReturn = eTT_OrderType_StopWithLimitCloseOnly
        
        Case "MIT"
            nReturn = eTT_OrderType_MIT
        Case "MARKET ON OPEN"
            nReturn = eTT_OrderType_MarketOnOpen
    End Select
    
    OrderTypeFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.OrderTypeFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeInForceToString
'' Description: Determine the string description of the given time in force
'' Inputs:      Time In Force
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimeInForceToString(ByVal nTimeInForce As eTT_TimeInForce) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nTimeInForce
        Case eTT_TimeInForce_Day
            strReturn = "Day"
        Case eTT_TimeInForce_GTC
            strReturn = "GTC"
        Case eTT_TimeInForce_GTD
            strReturn = "GTD"
    End Select
    
    TimeInForceToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.TimeInForceToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeInForceFromString
'' Description: Determine the time in force for the given description
'' Inputs:      String description
'' Returns:     Time in Force
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimeInForceFromString(ByVal strTimeInForce As String) As eTT_TimeInForce
On Error GoTo ErrSection:

    Dim nReturn As eTT_TimeInForce      ' Return value for the function
    
    Select Case UCase(strTimeInForce)
        Case "DAY"
            nReturn = eTT_TimeInForce_Day
        Case "GTC"
            nReturn = eTT_TimeInForce_GTC
        Case "GTD"
            nReturn = eTT_TimeInForce_GTD
    End Select
    
    TimeInForceFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.TimeInForceFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TypeOfAccountToString
'' Description: Determine the string description of the given type of account
'' Inputs:      Type of Account
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TypeOfAccountToString(ByVal nTypeOfAccount As eGDTypeOfAccount) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nTypeOfAccount
        Case eGDTypeOfAccount_Simulated
            strReturn = "eGDTypeOfAccount_Simulated"
        Case eGDTypeOfAccount_BrokerLive
            strReturn = "eGDTypeOfAccount_BrokerLive"
        Case eGDTypeOfAccount_BrokerDemo
            strReturn = "eGDTypeOfAccount_BrokerDemo"
        Case Else
            strReturn = Str(nTypeOfAccount)
    End Select

    TypeOfAccountToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.TypeOfAccountToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TypeOfAccountFromString
'' Description: Determine the type of account from the given string
'' Inputs:      String description
'' Returns:     Type of Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TypeOfAccountFromString(ByVal strTypeOfAccount As String) As eGDTypeOfAccount
On Error GoTo ErrSection:

    Dim nReturn As eGDTypeOfAccount     ' Return value for the function
    
    Select Case UCase(strTypeOfAccount)
        Case UCase("eGDTypeOfAccount_Simulated")
            nReturn = eGDTypeOfAccount_Simulated
        Case UCase("eGDTypeOfAccount_BrokerLive")
            nReturn = eGDTypeOfAccount_BrokerLive
        Case UCase("eGDTypeOfAccount_BrokerDemo")
            nReturn = eGDTypeOfAccount_BrokerDemo
        Case Else
            nReturn = -1&
    End Select
    
    TypeOfAccountFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.TypeOfAccountFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillMatchToString
'' Description: Determine the string description of the given fill match mode
'' Inputs:      Fill Match Mode
'' Returns:     String description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillMatchModeToString(ByVal nFillMatchMode As eTT_FillMatchMode) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nFillMatchMode
        Case eTT_FillMatchMode_Fifo
            strReturn = "eTT_FillMatchMode_Fifo"
        Case eTT_FillMatchMode_Lifo
            strReturn = "eTT_FillMatchMode_Lifo"
        Case Else
            strReturn = Str(nFillMatchMode)
    End Select
    
    FillMatchModeToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.FillMatchModeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillMatchModeFromString
'' Description: Determine the fill match mode from the given string
'' Inputs:      String description
'' Returns:     Fill Match Mode
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillMatchModeFromString(ByVal strFillMatchMode As String) As eTT_FillMatchMode
On Error GoTo ErrSection:

    Dim nReturn As eTT_FillMatchMode    ' Return value for the function
    
    Select Case UCase(strFillMatchMode)
        Case UCase("eTT_FillMatchMode_Fifo")
            nReturn = eTT_FillMatchMode_Fifo
        Case UCase("eTT_FillMatchMode_Lifo")
            nReturn = eTT_FillMatchMode_Lifo
        Case Else
            nReturn = -1&
    End Select
    
    FillMatchModeFromString = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerEnums.FillMatchModeFromString"
    
End Function
