VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActivityLogControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cActivityLogControls.cls
'' Description: Controls for use with the activity log
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public frm As Form                      ' Form that the activity log is on
Public fgGrid As VSFlexGrid             ' Grid that contains the activity log
Public tmrRealtime As Timer             ' Realtime timer on the form
Public tmrMenu As Timer                 ' Timer to perform selected menu action
    
Public mnuActivityLog As Menu           ' Popup menu to display in the grid
Public mnuPrint As Menu                 ' Menu item to allow user to print the grid
Public mnuTradeHistory As Menu          ' Menu item to allow user to view trade history
Public mnuSettings As Menu              ' Menu item to allow user to change settings
Public mnuCheckStatus As Menu           ' Menu item to allow user to check sim status
Public mnuViewJournals As Menu          ' Menu item to allow user to view journals

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
    
    Set mnuActivityLog = Nothing
    Set mnuPrint = Nothing
    Set mnuTradeHistory = Nothing
    Set mnuSettings = Nothing
    Set mnuCheckStatus = Nothing
    Set mnuViewJournals = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLogControls.ClearMembers"
    
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
    RaiseError "cActivityLogControls.Class_Initialize"
    
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
    RaiseError "cActivityLogControls.Class_Terminate"
    
End Sub
