Attribute VB_Name = "mJournal"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mJournal.bas
'' Description: Module for journal stuff
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/08/2014   DAJ         Added bridge object for NavCore
'' 10/24/2014   DAJ         Core Application functions for DLL's; Trade Tracker Database object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eGDJournalImageTypes
    eGDJournalImageType_Chart = 0
    eGDJournalImageType_SummaryReport
    eGDJournalImageType_OptionNavOrder
End Enum

Public Enum eGDJournalCategoryTypes
    eGDJournalCategoryType_Note = -1
    eGDJournalCategoryType_MoneyCode = 0
    eGDJournalCategoryType_CustomChecklist = 1
End Enum

'RH
Public Enum eStyleColorTypes
    'form
    eForm_Background = 0
    
    'frame
    eFrame_Background = 1
    eFrame_Border = 2
    
    'button
    eButton_Background = 3
    eButton_Border = 4
    eButton_Text = 5
    
    'checkbox
    eCheck_Border = 6
    eCheck_Background = 7
    eCheck_Forecolor = 8
    
    'flexgrid
    eGrid_Background = 9
    
End Enum

Type gGlobal

    'RH
    Styler As New cStyler
    
    AppBridge As cJournalTn             ' Application bridge from Journal DLL to Trade Navigator
    TnCore As cCoreTn                   ' Application bridge for core function calls into Trade Navigator
    strAppPath As String                ' Main application path
    strIniFile As String                ' Main application INI file
    frmMain As Form                     ' Main form
    bAppIsIde As Boolean                ' Is the main application being run in the IDE?
    
    CoreBridge As cCoreBridge           ' Bridge to the Core DLL
    
    TradeTrackerDB As cTradeTrackerDb   ' Trade Tracker database wrapper
    JournalDB As cJournalDatabase       ' Journal database
    
    JournalCategories As cJournalCategories ' Collection of journal categories
End Type
Global g As gGlobal

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
