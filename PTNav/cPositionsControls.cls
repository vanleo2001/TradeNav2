VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPositionsControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPositionsControls.cls
'' Description: Wrapper class for the controls necessary for positions
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/20/2013   DAJ         Added "Actual Performance" menu item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public frm As Form                      ' Form that the positions are on
    
Public fgGrid As VSFlexGrid             ' Grid that contains the positions
    
Public tmrRealTime As Timer             ' Realtime timer on the form
Public tmrMenu As Timer                 ' Timer to perform selected menu action
    
Public mnuPositions As Menu             ' Popup menu to display in the grid
Public mnuFlatten As Menu               ' Menu item to allow user to flatten position
Public mnuReverse As Menu               ' Menu item to allow user to reverse position
Public mnuManageXOS As Menu             ' Menu item to allow user to manage auto exits
Public mnuSelectXOS As Menu             ' Menu item to allow user to select auto exits
Public mnuRemoveXOS As Menu             ' Menu item to allow user to remove auto exit
Public mnuPrint As Menu                 ' Menu item to allow user to print the grid
Public mnuActualPerformance As Menu     ' Menu item to allow user to view actual performance reports
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
    
    Set mnuPositions = Nothing
    Set mnuFlatten = Nothing
    Set mnuReverse = Nothing
    Set mnuManageXOS = Nothing
    Set mnuSelectXOS = Nothing
    Set mnuRemoveXOS = Nothing
    Set mnuPrint = Nothing
    Set mnuActualPerformance = Nothing
    Set mnuTradeHistory = Nothing
    Set mnuSettings = Nothing
    Set mnuCheckStatus = Nothing
    Set mnuViewJournals = Nothing
    
    Set mnuAutoSizeColumns = Nothing
    Set mnuDefaultColumns = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPositionsControls.ClearMembers"
    
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
    RaiseError "cPositionsControls.Class_Initialize"
    
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
    RaiseError "cPositionsControls.Class_Terminate"
    
End Sub
