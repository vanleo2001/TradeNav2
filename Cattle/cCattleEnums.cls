VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCattleEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cCattleEnums.cls
'' Description: Enumerations for the Cattle DLL
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
'' 03/17/2014   DAJ         Added the FeedYard Source enum; CustomerTypeToString method
'' 04/15/2014   DAJ         Added GetFeedYardCustomers and FeedYardCustomer
'' 05/15/2014   DAJ         Implemented DeletePosition callback from server
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eGDCattleMessage
    eGDCattleMessage_Heartbeat = 0
    eGDCattleMessage_Connect = 1
    eGDCattleMessage_ConnectionStatus = 2
    eGDCattleMessage_Disconnect = 3
    eGDCattleMessage_GetFeedYards = 5
    eGDCattleMessage_FeedYard = 6
    eGDCattleMessage_GetCustomers = 7
    eGDCattleMessage_Customer = 8
    eGDCattleMessage_GetLots = 9
    eGDCattleMessage_Lot = 10
    eGDCattleMessage_GetOrders = 11
    eGDCattleMessage_Order = 12
    eGDCattleMessage_GetAssociatedFills = 13
    eGDCattleMessage_AssociatedFill = 14
    eGDCattleMessage_GetPositions = 15
    eGDCattleMessage_Position = 16
    eGDCattleMessage_GetAccounts = 17
    eGDCattleMessage_Account = 18
    eGDCattleMessage_UpdateAccount = 19
    eGDCattleMessage_UpdateOrder = 21
    eGDCattleMessage_UpdateFill = 23
    eGDCattleMessage_UpdatePosition = 25
    eGDCattleMessage_GetTrades = 27
    eGDCattleMessage_Trade = 28
    eGDCattleMessage_DeleteTrade = 30
    eGDCattleMessage_DeleteOrder = 32
    eGDCattleMessage_DeleteAssociatedFill = 34
    eGDCattleMessage_AppLoaded = 36
    eGDCattleMessage_UnloadApp = 37
    eGDCattleMessage_AppUnloaded = 38
    eGDCattleMessage_GetAllLotColumns = 39
    eGDCattleMessage_LotColumns = 40
    eGDCattleMessage_GetVisibleLotColumns = 41
    eGDCattleMessage_VisibleLotColumns = 42
    eGDCattleMessage_RemoveAccount = 43
    eGDCattleMessage_GetAllFills = 45
    eGDCattleMessage_Fill = 46
    eGDCattleMessage_AssociateFill = 47
    eGDCattleMessage_GetVisibleLots = 49
    eGDCattleMessage_VisibleLots = 50
    eGDCattleMessage_GetLotColumnCategories = 51
    eGDCattleMessage_LotColumnCategories = 52
    eGDCattleMessage_GetLotContentDetails = 53
    eGDCattleMessage_LotContentDetails = 54
    eGDCattleMessage_GenesisCustomerInfo = 56
    eGDCattleMessage_UpdateLotContentDetails = 57
    eGDCattleMessage_AddFeedYards = 59
    eGDCattleMessage_AddLots = 61
    eGDCattleMessage_AddCustomers = 63
    eGDCattleMessage_GetLotColumnSubCategories = 65
    eGDCattleMessage_LotColumnSubCategories = 66
    eGDCattleMessage_GetDetailOptions = 67
    eGDCattleMessage_DetailOptions = 68
    eGDCattleMessage_GetRations = 69
    eGDCattleMessage_Ration = 70
    eGDCattleMessage_GetIngredients = 71
    eGDCattleMessage_Ingredient = 72
    eGDCattleMessage_UpdateRation = 73
    eGDCattleMessage_UpdateIngredient = 75
    eGDCattleMessage_GetFeedyardCustomers = 77
    eGDCattleMessage_FeedyardCustomer = 78
    eGDCattleMessage_DeletePosition = 80
    
    eGDCattleMessage_UpdateVisibleFeedYards = 1007
    eGDCattleMessage_UpdateVisibleCustomers = 1009
    eGDCattleMessage_GetGenesisCustomers = 1011
    eGDCattleMessage_GenesisCustomers = 1012
    eGDCattleMessage_GetAllFeedYards = 1013
    eGDCattleMessage_AllFeedYards = 1014
    eGDCattleMessage_GetAllCustomers = 1015
    eGDCattleMessage_AllCustomers = 1016
    eGDCattleMessage_GetVisibleFeedYards = 1017
    eGDCattleMessage_VisibleFeedYards = 1018
    eGDCattleMessage_GetVisibleCustomers = 1019
    eGDCattleMessage_VisibleCustomers = 1020
    eGDCattleMessage_UpdateGenesisCustomer = 1021
    eGDCattleMessage_GetAllLotColumnsAdmin = 1023
    eGDCattleMessage_LotColumnsAdmin = 1024
    eGDCattleMessage_GetVisibleLotColumnsAdmin = 1025
    eGDCattleMessage_VisibleLotColumnsAdmin = 1026
    eGDCattleMessage_UpdateVisibleLotColumns = 1027
    eGDCattleMessage_GetDefaultVisibleLotColumnsAdmin = 1029
    eGDCattleMessage_DefaultVisibleLotColumnsAdmin = 1030
    eGDCattleMessage_UpdateDefaultVisibleLotColumns = 1031
End Enum

Public Enum eGDCattleCustomerType
    eGDCattleCustomerType_TurnkeyCustomer = 0
    eGDCattleCustomerType_Broker = 1
    eGDCattleCustomerType_Admin = 2
    eGDCattleCustomerType_TurnkeyFeedYard = 3
    eGDCattleCustomerType_CattleNavCustomer = 4
    eGDCattleCustomerType_CattleNavFeedYard = 5
    eGDCattleCustomerType_EitherFeedYard = 6
End Enum

Public Enum eGDCattleFeedYardSource
    eGDCattleFeedYardSource_Turnkey = 1
    eGDCattleFeedYardSource_Manual = 2
    eGDCattleFeedYardSource_ViewTrak = 3
End Enum

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CattleMessageTypeToString
'' Description: Return a string representation of the Cattle message enumeration
'' Inputs:      Cattle Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CattleMessageTypeToString(ByVal nCattleMessage As eGDCattleMessage) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nCattleMessage
        Case eGDCattleMessage_Heartbeat
            strReturn = "eGDCattleMessage_Heartbeat"
        Case eGDCattleMessage_Connect
            strReturn = "eGDCattleMessage_Connect"
        Case eGDCattleMessage_ConnectionStatus
            strReturn = "eGDCattleMessage_ConnectionStatus"
        Case eGDCattleMessage_Disconnect
            strReturn = "eGDCattleMessage_Disconnect"
        Case eGDCattleMessage_GetFeedYards
            strReturn = "eGDCattleMessage_GetFeedYards"
        Case eGDCattleMessage_FeedYard
            strReturn = "eGDCattleMessage_FeedYard"
        Case eGDCattleMessage_GetCustomers
            strReturn = "eGDCattleMessage_GetCustomers"
        Case eGDCattleMessage_Customer
            strReturn = "eGDCattleMessage_Customer"
        Case eGDCattleMessage_GetLots
            strReturn = "eGDCattleMessage_GetLots"
        Case eGDCattleMessage_Lot
            strReturn = "eGDCattleMessage_Lot"
        Case eGDCattleMessage_GetOrders
            strReturn = "eGDCattleMessage_GetOrders"
        Case eGDCattleMessage_Order
            strReturn = "eGDCattleMessage_Order"
        Case eGDCattleMessage_GetAssociatedFills
            strReturn = "eGDCattleMessage_GetAssociatedFills"
        Case eGDCattleMessage_AssociatedFill
            strReturn = "eGDCattleMessage_AssociatedFill"
        Case eGDCattleMessage_GetPositions
            strReturn = "eGDCattleMessage_GetPositions"
        Case eGDCattleMessage_Position
            strReturn = "eGDCattleMessage_Position"
        Case eGDCattleMessage_GetAccounts
            strReturn = "eGDCattleMessage_GetAccounts"
        Case eGDCattleMessage_Account
            strReturn = "eGDCattleMessage_Account"
        Case eGDCattleMessage_UpdateAccount
            strReturn = "eGDCattleMessage_UpdateAccount"
        Case eGDCattleMessage_UpdateOrder
            strReturn = "eGDCattleMessage_UpdateOrder"
        Case eGDCattleMessage_UpdateFill
            strReturn = "eGDCattleMessage_UpdateFill"
        Case eGDCattleMessage_UpdatePosition
            strReturn = "eGDCattleMessage_UpdatePosition"
        Case eGDCattleMessage_GetTrades
            strReturn = "eGDCattleMessage_GetTrades"
        Case eGDCattleMessage_Trade
            strReturn = "eGDCattleMessage_Trade"
        Case eGDCattleMessage_DeleteTrade
            strReturn = "eGDCattleMessage_DeleteTrade"
        Case eGDCattleMessage_DeleteOrder
            strReturn = "eGDCattleMessage_DeleteOrder"
        Case eGDCattleMessage_DeleteAssociatedFill
            strReturn = "eGDCattleMessage_DeleteAssociatedFill"
        Case eGDCattleMessage_AppLoaded
            strReturn = "eGDCattleMessage_AppLoaded"
        Case eGDCattleMessage_UnloadApp
            strReturn = "eGDCattleMessage_UnloadApp"
        Case eGDCattleMessage_AppUnloaded
            strReturn = "eGDCattleMessage_AppUnloaded"
        Case eGDCattleMessage_GetAllLotColumns
            strReturn = "eGDCattleMessage_GetAllLotColumns"
        Case eGDCattleMessage_LotColumns
            strReturn = "eGDCattleMessage_LotColumns"
        Case eGDCattleMessage_GetVisibleLotColumns
            strReturn = "eGDCattleMessage_GetVisibleLotColumns"
        Case eGDCattleMessage_VisibleLotColumns
            strReturn = "eGDCattleMessage_VisibleLotColumns"
        Case eGDCattleMessage_RemoveAccount
            strReturn = "eGDCattleMessage_RemoveAccount"
        Case eGDCattleMessage_GetAllFills
            strReturn = "eGDCattleMessage_GetAllFills"
        Case eGDCattleMessage_Fill
            strReturn = "eGDCattleMessage_Fill"
        Case eGDCattleMessage_AssociateFill
            strReturn = "eGDCattleMessage_AssociateFill"
        Case eGDCattleMessage_GetVisibleLots
            strReturn = "eGDCattleMessage_GetVisibleLots"
        Case eGDCattleMessage_VisibleLots
            strReturn = "eGDCattleMessage_VisibleLots"
        Case eGDCattleMessage_GetLotColumnCategories
            strReturn = "eGDCattleMessage_GetLotColumnCategories"
        Case eGDCattleMessage_LotColumnCategories
            strReturn = "eGDCattleMessage_LotColumnCategories"
        Case eGDCattleMessage_GetLotContentDetails
            strReturn = "eGDCattleMessage_GetLotContentDetails"
        Case eGDCattleMessage_LotContentDetails
            strReturn = "eGDCattleMessage_LotContentDetails"
        Case eGDCattleMessage_GenesisCustomerInfo
            strReturn = "eGDCattleMessage_GenesisCustomerInfo"
        Case eGDCattleMessage_UpdateLotContentDetails
            strReturn = "eGDCattleMessage_UpdateLotContentDetails"
        Case eGDCattleMessage_AddFeedYards
            strReturn = "eGDCattleMessage_AddFeedYards"
        Case eGDCattleMessage_AddLots
            strReturn = "eGDCattleMessage_AddLots"
        Case eGDCattleMessage_AddCustomers
            strReturn = "eGDCattleMessage_AddCustomers"
        Case eGDCattleMessage_GetLotColumnSubCategories
            strReturn = "eGDCattleMessage_GetLotColumnSubCategories"
        Case eGDCattleMessage_LotColumnSubCategories
            strReturn = "eGDCattleMessage_LotColumnSubCategories"
        Case eGDCattleMessage_GetDetailOptions
            strReturn = "eGDCattleMessage_GetDetailOptions"
        Case eGDCattleMessage_DetailOptions
            strReturn = "eGDCattleMessage_DetailOptions"
        Case eGDCattleMessage_GetRations
            strReturn = "eGDCattleMessage_GetRations"
        Case eGDCattleMessage_Ration
            strReturn = "eGDCattleMessage_Ration"
        Case eGDCattleMessage_GetIngredients
            strReturn = "eGDCattleMessage_GetIngredients"
        Case eGDCattleMessage_Ingredient
            strReturn = "eGDCattleMessage_Ingredient"
        Case eGDCattleMessage_UpdateRation
            strReturn = "eGDCattleMessage_UpdateRation"
        Case eGDCattleMessage_UpdateIngredient
            strReturn = "eGDCattleMessage_UpdateIngredient"
        Case eGDCattleMessage_GetFeedyardCustomers
            strReturn = "eGDCattleMessage_GetFeedyardCustomers"
        Case eGDCattleMessage_FeedyardCustomer
            strReturn = "eGDCattleMessage_FeedyardCustomer"
        Case eGDCattleMessage_DeletePosition
            strReturn = "eGDCattleMessage_DeletePosition"
            
        Case eGDCattleMessage_UpdateVisibleFeedYards
            strReturn = "eGDCattleMessage_UpdateVisibleFeedYards"
        Case eGDCattleMessage_UpdateVisibleCustomers
            strReturn = "eGDCattleMessage_UpdateVisibleCustomers"
        Case eGDCattleMessage_GetGenesisCustomers
            strReturn = "eGDCattleMessage_GetGenesisCustomers"
        Case eGDCattleMessage_GenesisCustomers
            strReturn = "eGDCattleMessage_GenesisCustomers"
        Case eGDCattleMessage_GetAllFeedYards
            strReturn = "eGDCattleMessage_GetAllFeedYards"
        Case eGDCattleMessage_AllFeedYards
            strReturn = "eGDCattleMessage_AllFeedYards"
        Case eGDCattleMessage_GetAllCustomers
            strReturn = "eGDCattleMessage_GetAllCustomers"
        Case eGDCattleMessage_AllCustomers
            strReturn = "eGDCattleMessage_AllCustomers"
        Case eGDCattleMessage_GetVisibleFeedYards
            strReturn = "eGDCattleMessage_GetVisibleFeedYards"
        Case eGDCattleMessage_VisibleFeedYards
            strReturn = "eGDCattleMessage_VisibleFeedYards"
        Case eGDCattleMessage_GetVisibleCustomers
            strReturn = "eGDCattleMessage_GetVisibleCustomers"
        Case eGDCattleMessage_VisibleCustomers
            strReturn = "eGDCattleMessage_VisibleCustomers"
        Case eGDCattleMessage_UpdateGenesisCustomer
            strReturn = "eGDCattleMessage_UpdateGenesisCustomer"
        Case eGDCattleMessage_GetAllLotColumnsAdmin
            strReturn = "eGDCattleMessage_GetAllLotColumnsAdmin"
        Case eGDCattleMessage_LotColumnsAdmin
            strReturn = "eGDCattleMessage_LotColumnsAdmin"
        Case eGDCattleMessage_GetVisibleLotColumnsAdmin
            strReturn = "eGDCattleMessage_GetVisibleLotColumnsAdmin"
        Case eGDCattleMessage_VisibleLotColumnsAdmin
            strReturn = "eGDCattleMessage_VisibleLotColumnsAdmin"
        Case eGDCattleMessage_UpdateVisibleLotColumns
            strReturn = "eGDCattleMessage_UpdateVisibleLotColumns"
        Case eGDCattleMessage_GetDefaultVisibleLotColumnsAdmin
            strReturn = "eGDCattleMessage_GetDefaultVisibleLotColumnsAdmin"
        Case eGDCattleMessage_DefaultVisibleLotColumnsAdmin
            strReturn = "eGDCattleMessage_DefaultVisibleLotColumnsAdmin"
        Case eGDCattleMessage_UpdateDefaultVisibleLotColumns
            strReturn = "eGDCattleMessage_UpdateDefaultVisibleLotColumns"
            
        Case Else
            strReturn = Str(nCattleMessage)
    End Select
    
    CattleMessageTypeToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleEnums.CattleMessageTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CattleCustomerTypeToString
'' Description: Return a string representation of the Customer type enum
'' Inputs:      Cattle Customer Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CattleCustomerTypeToString(ByVal nCustomerType As eGDCattleCustomerType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nCustomerType
        Case eGDCattleCustomerType_TurnkeyCustomer
            strReturn = "eGDCattleCustomerType_TurnkeyCustomer"
        Case eGDCattleCustomerType_Broker
            strReturn = "eGDCattleCustomerType_Broker"
        Case eGDCattleCustomerType_Admin
            strReturn = "eGDCattleCustomerType_Admin"
        Case eGDCattleCustomerType_TurnkeyFeedYard
            strReturn = "eGDCattleCustomerType_TurnkeyFeedYard"
        Case eGDCattleCustomerType_CattleNavCustomer
            strReturn = "eGDCattleCustomerType_CattleNavCustomer"
        Case eGDCattleCustomerType_CattleNavFeedYard
            strReturn = "eGDCattleCustomerType_CattleNavFeedYard"
        Case eGDCattleCustomerType_EitherFeedYard
            strReturn = "eGDCattleCustomerType_EitherFeedYard"
            
        Case Else
            strReturn = Str(nCustomerType)
    End Select
    
    CattleCustomerTypeToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleEnums.CattleCustomerTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CattleFeedYardSourceToString
'' Description: Return a string representation of the Feed Yard source enum
'' Inputs:      Cattle Feed Yard Source
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CattleFeedYardSourceToString(ByVal nFeedYardSource As eGDCattleFeedYardSource) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Select Case nFeedYardSource
        Case eGDCattleFeedYardSource_Turnkey
            strReturn = "eGDCattleFeedYardSource_Turnkey"
        Case eGDCattleFeedYardSource_Manual
            strReturn = "eGDCattleFeedYardSource_Manual"
        Case eGDCattleFeedYardSource_ViewTrak
            strReturn = "eGDCattleFeedYardSource_ViewTrak"
            
        Case Else
            strReturn = Str(nFeedYardSource)
    End Select
    
    CattleFeedYardSourceToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleEnums.CattleFeedYardSourceToString"
    
End Function

