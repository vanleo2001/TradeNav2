VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWorkingOrdersControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cWorkingOrdersControls.cls
'' Description: Wrapper class for the controls necessary for working orders
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 06/03/2010   DAJ         Changes for new Trade Sense Order Groups
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public frm As Form                      ' Form that the working orders are on
    
Public fgGrid As VSFlexGrid             ' Grid that contains the working orders
    
Public tmrRealtime As Timer             ' Realtime timer on the form
Public tmrMenu As Timer                 ' Timer to perform selected menu action
    
Public mnuOrders As Menu                ' Popup menu to display in the grid
Public mnuBuy As Menu                   ' Menu item to allow user to Buy a security
Public mnuSell As Menu                  ' Menu item to allow user to Sell a security
Public mnuTsOrderGroups As Menu         ' Menu item to allow user to manage Trade Sense order groups
Public mnuOrderGroups As Menu           ' Menu item to allow user to see order groups
Public mnuOrderGroup As Object          ' Menu item array for order groups
Public mnuEditOrder As Menu             ' Menu item to allow user to edit an order
Public mnuCancelOrder As Menu           ' Menu item to allow user to cancel an order
Public mnuParkOrder As Menu             ' Menu item to allow user to park an order
Public mnuSubmitOrder As Menu           ' Menu item to allow user to submit a parked order
Public mnuSubmitAll As Menu             ' Menu item to allow user to submit all parked orders
Public mnuOrderHistory As Menu          ' Menu item to allow user to view order history
Public mnuNewJournal As Menu            ' Menu item to allow user to create a new journal
Public mnuManageXOS As Menu             ' Menu item to allow user to manage auto exits
Public mnuSelectXOS As Menu             ' Menu item to allow user to select auto exits
Public mnuRemoveXOS As Menu             ' Menu item to allow user to remove auto exit
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
    
    Set tmrRealtime = Nothing
    Set tmrMenu = Nothing
    
    Set mnuOrders = Nothing
    Set mnuBuy = Nothing
    Set mnuSell = Nothing
    Set mnuTsOrderGroups = Nothing
    Set mnuOrderGroups = Nothing
    Set mnuOrderGroup = Nothing
    Set mnuEditOrder = Nothing
    Set mnuCancelOrder = Nothing
    Set mnuParkOrder = Nothing
    Set mnuSubmitOrder = Nothing
    Set mnuSubmitAll = Nothing
    Set mnuOrderHistory = Nothing
    Set mnuNewJournal = Nothing
    Set mnuManageXOS = Nothing
    Set mnuSelectXOS = Nothing
    Set mnuRemoveXOS = Nothing
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
    RaiseError "cWorkingOrdersControls.ClearMembers"
    
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
    RaiseError "cWorkingOrdersControls.Class_Initialize"
    
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
    RaiseError "cWorkingOrdersControls.Class_Terminate"
    
End Sub

