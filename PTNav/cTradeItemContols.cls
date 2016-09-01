VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeItemContols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeItemControls.cls
'' Description: Wrapper class for the controls necessary for auto trading items
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/02/2011   DAJ         Added the change position menu item
'' 02/20/2013   DAJ         Added "Actual Performance" menu item, renamed "Strategy Performance"
'' 01/23/2015   DAJ         Disable All menu item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public frm As Form                      ' Form that the accounts are on

Public fgGrid As VSFlexGrid             ' Grid that contains the accounts

Public tmrRealTime As Timer             ' Realtime timer on the form
Public tmrMenu As Timer                 ' Timer to perform selected menu action

Public mnuTradeItems As Menu            ' Popup menu to display in the grid
Public mnuDisableAll As Menu            ' Menu item to allow user to disable all active automated trading items
Public mnuFlatten As Menu               ' Menu item to allow user to flatten a position
Public mnuEnterPosition As Menu         ' Menu item to allow user to enter a position
Public mnuChangePosition As Menu        ' Menu item to allow user to change a position
Public mnuNewTradeItem As Menu          ' Menu item to allow user to create a new auto trade item
Public mnuEditTradeItem As Menu         ' Menu item to allow user to edit an auto trade item
Public mnuDeleteTradeItem As Menu       ' Menu item to allow user to delete an auto trade item
Public mnuRollContract As Menu          ' Menu item to allow user to roll the contract
Public mnuEditStrategy As Menu          ' Menu item to allow user to edit the strategy
Public mnuStrategyPerformance As Menu   ' Menu item to allow user to view strategy performance reports
Public mnuActualPerformance As Menu     ' Menu item to allow user to view actual performance reports
Public mnuNextBarReport As Menu         ' Menu item to allow user to view the next bar report
Public mnuShowChart As Menu             ' Menu item to allow user to show information on chart
Public mnuPrint As Menu                 ' Menu item to allow user to print the grid
Public mnuTradeHistory As Menu          ' Menu item to allow user to view trade history
Public mnuAlerts As Menu                ' Menu item to allow user to view and setup alerts
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

    Set mnuDisableAll = Nothing
    Set mnuTradeItems = Nothing
    Set mnuFlatten = Nothing
    Set mnuEnterPosition = Nothing
    Set mnuChangePosition = Nothing
    Set mnuNewTradeItem = Nothing
    Set mnuEditTradeItem = Nothing
    Set mnuDeleteTradeItem = Nothing
    Set mnuRollContract = Nothing
    Set mnuEditStrategy = Nothing
    Set mnuStrategyPerformance = Nothing
    Set mnuActualPerformance = Nothing
    Set mnuNextBarReport = Nothing
    Set mnuShowChart = Nothing
    Set mnuPrint = Nothing
    Set mnuTradeHistory = Nothing
    Set mnuAlerts = Nothing
    
    Set mnuAutoSizeColumns = Nothing
    Set mnuDefaultColumns = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeItemControls.ClearMembers"
    
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
    RaiseError "cTradeItemControls.Class_Initialize"
    
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
    RaiseError "cTradeItemControls.Class_Terminate"
    
End Sub
