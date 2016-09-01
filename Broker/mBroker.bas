Attribute VB_Name = "mBroker"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mBroker.bas
'' Description: Main module for the Broker DLL
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
'' 09/11/2014   DAJ         Added Database objects; Removed form functions; Added Account functions
'' 10/24/2014   DAJ         Core Application functions for DLL's; Database Objects;
''                          Move account objects out of NavSuite into NavBroker.DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Type gGlobal
    AppBridge As cBrokerTn              ' Application bridge from Broker DLL to Trade Navigator
    TnCore As cCoreTn                   ' Application bridge for core function calls into Trade Navigator
    strAppPath As String                ' Main application path
    strIniFile As String                ' Main application INI file
    
    TradeTrackerDB As cTradeTrackerDb   ' Trade Tracker database wrapper
    BrokerDB As cBrokerDb               ' Interaction between broker stuff and Trade Tracker database
End Type
Global g As gGlobal

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsSimulatedAccount
'' Description: Determine if the given account type is a simulated account
'' Inputs:      Account Type
'' Returns:     True if a simulated account, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsSimulatedAccount(ByVal nAccountType As eTT_AccountType) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    Select Case nAccountType
        Case eTT_AccountType_SimBroker, eTT_AccountType_SimReplay, eTT_AccountType_SimStream
            bReturn = True
        Case Else
            bReturn = False
    End Select
    
    IsSimulatedAccount = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "mBroker.IsSimulatedAccount"
    
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
On Error GoTo ErrSection:

    Dim nReturn As eGDTypeOfAccount     ' Return value for the function
    
    If IsSimulatedAccount(nAccountType) = True Then
        nReturn = eGDTypeOfAccount_Simulated
    ElseIf nAccountType = eTT_AccountType_DemoPats Then
        nReturn = eGDTypeOfAccount_BrokerDemo
    Else
        nReturn = eGDTypeOfAccount_BrokerLive
    End If
    
    TypeOfAccount = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "mBroker.TypeOfAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMsg
'' Description: Show an error message
'' Inputs:      Form to Save
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMsg(Optional ByVal lErrNum& = 0, Optional ByVal strSource$ = "", Optional ByVal strDesc$ = "")
    
    Dim RetVal As Variant
    
    If lErrNum = 0 Then
        lErrNum = Err.Number
        strSource = Err.Source
        strDesc = Err.Description
    End If
    
    RetVal = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
    
    If lErrNum < 0 Then
        Replace strDesc, vbCrLf, "|"
        InfBox strDesc, , , "Error", , , , , , , , eGDAlign_Left
    Else
        Replace strDesc, vbCrLf, "|"
        InfBox "An unexpected error occurred.||Please report the following: " & _
            "|Source:  " & strSource & _
            "|Message: " & strDesc, , , "Error", , , , , , , , eGDAlign_Left
    End If

End Sub


