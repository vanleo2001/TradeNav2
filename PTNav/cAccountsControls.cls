VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccountsControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAccountsControls.cls
'' Description: Wrapper class for the controls necessary for accounts
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/07/2011   DAJ         Added Change Password menu item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public frm As Form                      ' Form that the accounts are on
    
Public fgGrid As VSFlexGrid             ' Grid that contains the accounts
    
Public tmrRealTime As Timer             ' Realtime timer on the form
Public tmrMenu As Timer                 ' Timer to perform selected menu action
    
Public mnuAccounts As Menu              ' Popup menu to display in the grid
Public mnuConnect As Menu               ' Menu item to allow user to connect to broker
Public mnuDisconnect As Menu            ' Menu item to allow user to disconnect from the broker
Public mnuSwitchAccounts As Menu        ' Menu item to allow user to switch broker accounts
Public mnuSwitchAccountsMode As Menu    ' Menu item to allow user to switch broker accounts mode
Public mnuConnectInfo As Menu           ' Menu item to allow user to view connection information
Public mnuChangePassword As Menu        ' Menu item to allow user to change their broker password
Public mnuRefresh As Menu               ' Menu item to allow user to refresh their account
Public mnuViewActivity As Menu          ' Menu item to allow user to view broker activity
Public mnuBrokerView As Menu            ' Menu item to allow user to show the broker view
Public mnuViewOnline As Menu            ' Menu item to allow user to view account online
Public mnuVerifyPositions As Menu       ' Menu item to allow user to verify their positions
Public mnuAccountDetails As Menu        ' Menu item to allow user to view account details

Public mnuSep1 As Menu                  ' Separator
Public mnuNewAccount As Menu            ' Menu item to allow user to create a new account
Public mnuEditAccount As Menu           ' Menu item to allow user to edit the selected account
Public mnuDeleteAccount As Menu         ' Menu item to allow user to delete the selected account
Public mnuReports As Menu               ' Menu item to allow user to view performance reports
Public mnuSep2 As Menu                  ' Separator

Public mnuPrint As Menu                 ' Menu item to allow user to print the grid
Public mnuTradeHistory As Menu          ' Menu item to allow user to view trade history
Public mnuSettings As Menu              ' Menu item to allow user to change settings
Public mnuCheckStatus As Menu           ' Menu item to allow user to check sim status
Public mnuViewJournals As Menu          ' Menu item to allow user to view journals
Public mnuAutoSizeColumns As Menu       ' Menu item to allow user to auto size the columns
Public mnuDefaultColumns As Menu        ' Menu item to allow user to revert to default column order

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearMembers
'' Description: Set all of the class members to nothing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearMembers()
On Error GoTo ErrSection:

    Set frm = Nothing
    
    Set fgGrid = Nothing
    
    Set tmrRealTime = Nothing
    Set tmrMenu = Nothing
    
    Set mnuAccounts = Nothing
    Set mnuConnect = Nothing
    Set mnuDisconnect = Nothing
    Set mnuSwitchAccounts = Nothing
    Set mnuSwitchAccountsMode = Nothing
    Set mnuConnectInfo = Nothing
    Set mnuChangePassword = Nothing
    Set mnuRefresh = Nothing
    Set mnuViewActivity = Nothing
    Set mnuBrokerView = Nothing
    Set mnuViewOnline = Nothing
    Set mnuVerifyPositions = Nothing
    Set mnuAccountDetails = Nothing
    
    Set mnuSep1 = Nothing
    Set mnuNewAccount = Nothing
    Set mnuEditAccount = Nothing
    Set mnuDeleteAccount = Nothing
    Set mnuReports = Nothing
    Set mnuSep2 = Nothing
    
    Set mnuPrint = Nothing
    Set mnuTradeHistory = Nothing
    Set mnuSettings = Nothing
    Set mnuCheckStatus = Nothing
    Set mnuViewJournals = Nothing
    
    Set mnuAutoSizeColumns = Nothing
    Set mnuDefaultColumns = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsControls.ClearMembers"
    
End Sub

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

    ClearMembers

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsControls.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    ClearMembers

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsControls.Class_Terminate"
    
End Sub

