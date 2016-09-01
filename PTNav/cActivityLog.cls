VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActivityLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cActivityLog.cls
'' Description: Class to handle the activity log
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/13/2010   DAJ         Added code for Rithmic
'' 11/01/2010   DAJ         Added Optimus, OpVest, and Vision (Rithmic Brokers)
'' 12/10/2010   DAJ         Moved the ActivityLogSource stuff over to cBrokerDispatch
'' 05/11/2011   DAJ         Utilize CheckSimTradeStatus, EnableCheckStatusMenu, IsLiveAccount
'' 09/23/2011   DAJ         Show date journals form instead of old journals form
'' 10/04/2011   DAJ         Call the ShowJournals function instead of calling the form direct
'' 06/19/2014   DAJ         Auto-size the grid after loading
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kMenuPrefix = "L:"

Private WithEvents m_fgActivityLog As VSFlexGrid
Attribute m_fgActivityLog.VB_VarHelpID = -1
Private WithEvents m_tmrMenu As Timer
Attribute m_tmrMenu.VB_VarHelpID = -1
Private WithEvents m_mnuPrint As Menu
Attribute m_mnuPrint.VB_VarHelpID = -1
Private WithEvents m_mnuTradeHistory As Menu
Attribute m_mnuTradeHistory.VB_VarHelpID = -1
Private WithEvents m_mnuSettings As Menu
Attribute m_mnuSettings.VB_VarHelpID = -1
Private WithEvents m_mnuCheckStatus As Menu
Attribute m_mnuCheckStatus.VB_VarHelpID = -1
Private WithEvents m_mnuViewJournals As Menu
Attribute m_mnuViewJournals.VB_VarHelpID = -1

Private Enum eGDActivityLogCols
    eGDActivityLogCol_Date = 0
    eGDActivityLogCol_DateDisplay
    eGDActivityLogCol_Source
    eGDActivityLogCol_Message
    eGDActivityLogCol_NumCols
End Enum

Private Type mPrivate
    lAccountID As Long                  ' Account ID
    
    TCGrid As cTradeConsoleGrid         ' Trade console grid object for generic routines
    tmrRealTime As Timer                ' Realtime timer control
    frm As Form                         ' Form that the grid is on
    mnuActivityLog As Object            ' Activity Log menu
End Type
Private m As mPrivate

Private Function ActivityLogCol(ByVal Col As eGDActivityLogCols)
    ActivityLogCol = Col
End Function

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the class
'' Inputs:      Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(UI As cActivityLogControls, Optional ByVal lAccountID = -1&)
On Error GoTo ErrSection:

    m.lAccountID = lAccountID
    With UI
        Set m.frm = .frm
        Set m_fgActivityLog = .fgGrid
        Set m.tmrRealTime = .tmrRealTime
        Set m_tmrMenu = .tmrMenu
        
        Set m.mnuActivityLog = .mnuActivityLog
        Set m_mnuPrint = .mnuPrint
        Set m_mnuTradeHistory = .mnuTradeHistory
        Set m_mnuSettings = .mnuSettings
        Set m_mnuCheckStatus = .mnuCheckStatus
        Set m_mnuViewJournals = .mnuViewJournals
    End With
    
    Set m.TCGrid = New cTradeConsoleGrid
    m.TCGrid.Init m_fgActivityLog, m.tmrRealTime, ""
    
    InitGrid
    LoadGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add an activity log entry to the activity log grid
'' Inputs:      Source, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add(ByVal strSource As String, ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim bAdd As Boolean                 ' Add to grid?
    
    bAdd = True
    If m.lAccountID > -1& Then
        bAdd = IncludeMessage(strSource)
    End If
    
    If bAdd Then
        With m_fgActivityLog
            .Redraw = flexRDNone
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, ActivityLogCol(eGDActivityLogCol_Date)) = Str(Date)
            .TextMatrix(.Rows - 1, ActivityLogCol(eGDActivityLogCol_DateDisplay)) = DateFormat(Now, MM_DD_YYYY, HH_MM_SS)
            .TextMatrix(.Rows - 1, ActivityLogCol(eGDActivityLogCol_Source)) = strSource
            .TextMatrix(.Rows - 1, ActivityLogCol(eGDActivityLogCol_Message)) = strMessage
            
            .AutoSize 0, .Cols - 1
            .ShowCell .Rows - 1, 0
            .Redraw = flexRDBuffered
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.Add"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveGrid
'' Description: Save the activity log grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveGrid()
On Error GoTo ErrSection:

    Dim astrLogFile As New cGdArray     ' Array of saved log file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strDate As String               ' Date for which to save information
    
    astrLogFile.Create eGDARRAY_Strings
    
    With m_fgActivityLog
        If .Rows > .FixedRows Then
            strDate = .TextMatrix(.Rows - 1, ActivityLogCol(eGDActivityLogCol_Date))
        End If
        
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, ActivityLogCol(eGDActivityLogCol_Date)) = strDate Then
                astrLogFile.Add .TextMatrix(lIndex, ActivityLogCol(eGDActivityLogCol_Date)) & vbTab & _
                                .TextMatrix(lIndex, ActivityLogCol(eGDActivityLogCol_DateDisplay)) & vbTab & _
                                .TextMatrix(lIndex, ActivityLogCol(eGDActivityLogCol_Source)) & vbTab & _
                                .TextMatrix(lIndex, ActivityLogCol(eGDActivityLogCol_Message))
            End If
        Next lIndex
        
        astrLogFile.ToFile AddSlash(App.Path) & "ActivityLog.LOG"
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.SaveGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview form for this grid
'' Inputs:      Arguments passed in from PrintMe
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    m.TCGrid.GenerateReport vArgs, "Activity Log"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Setup the popup menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    If Not m_mnuCheckStatus Is Nothing Then
        m_mnuCheckStatus.Enabled = g.Broker.EnableCheckStatusMenu
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    With m_fgActivityLog
        .Redraw = flexRDNone
        
        SetupGrid m_fgActivityLog, eGridMode_List
        .Cols = ActivityLogCol(eGDActivityLogCol_NumCols)
        .FixedCols = 0
        .Rows = 0
        
        .ColHidden(ActivityLogCol(eGDActivityLogCol_Date)) = True
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the activity log grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim astrLogFile As New cGdArray     ' Array of saved log file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSource As String             ' Source of the item
    
    If astrLogFile.FromFile(AddSlash(App.Path) & "ActivityLog.LOG") Then
        With m_fgActivityLog
            .Redraw = flexRDNone
            
            For lIndex = 0 To astrLogFile.Size - 1
                strSource = Parse(astrLogFile(lIndex), vbTab, 3)
                
                If m.lAccountID = -1& Then
                    .AddItem astrLogFile(lIndex)
                ElseIf IncludeMessage(strSource) Then
                    .AddItem astrLogFile(lIndex)
                End If
            Next lIndex
            
            .AutoSize 0, .Cols - 1
            .Redraw = flexRDBuffered
        
            If .Rows > .FixedRows Then
                .ShowCell .Rows - 1, 1
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IncludeMessage
'' Description: Determine whether to show a message with the given source and
''              account ID
'' Inputs:      Source
'' Returns:     True if include, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IncludeMessage(ByVal strSource As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim nBroker As eTT_AccountType      ' Account type for the given account

    nBroker = g.Broker.AccountTypeForID(m.lAccountID)
    
    If UCase(strSource) = "GENESIS" Or UCase(strSource) = "DEMO" Then
        bReturn = Not g.Broker.IsLiveAccount(nBroker)
    Else
        bReturn = (UCase(strSource) = UCase(g.Broker.ActivityLogSource(nBroker)))
    End If

    IncludeMessage = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActivityLog.IncludeMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartMenuTimer
'' Description: Start the menu timer with the given command
'' Inputs:      Command
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StartMenuTimer(ByVal strCommand As String)
On Error GoTo ErrSection:

    m_tmrMenu.Tag = kMenuPrefix & strCommand
    m_tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgActivityLog_BeforeMouseDown
'' Description: Bring up the Pop-Up menu when the user right clicks
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgActivityLog_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        EnableControls
        m.frm.PopupMenu m.mnuActivityLog
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.m_fgActivityLog_BeforeMouseDown"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuCheckStatus_Click
'' Description: Check the status of any open orders from the trade server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuCheckStatus_Click()
On Error GoTo ErrSection:

    g.Broker.CheckTradeServerOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.m_mnuCheckStatus_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuPrint_Click
'' Description: Allow the user to print the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuPrint_Click()
On Error GoTo ErrSection:

    m.frm.PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.m_mnuPrint_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSettings_Click
'' Description: Allow the user to view and edit their console settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuSettings_Click()
On Error GoTo ErrSection:

    StartMenuTimer "SETTINGS"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.m_mnuSettings_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuTradeHistory_Click
'' Description: Allow the user to view trade history for the given account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuTradeHistory_Click()
On Error GoTo ErrSection:

    StartMenuTimer "TRADEHISTORY"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.m_mnuTradeHistory_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuViewJournals_Click
'' Description: Allow the user to view their journal entries
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuViewJournals_Click()
On Error GoTo ErrSection:

    StartMenuTimer "JOURNALS"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.m_mnuViewJournals_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_tmrMenu_Timer
'' Description: Perform a specified action when the timer goes off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Action to perform

    strTag = m_tmrMenu.Tag
    If Len(strTag) > 2 Then
        If Left(strTag, 2) = kMenuPrefix Then
            strTag = Mid(strTag, 3)
            m_tmrMenu.Tag = ""
            m_tmrMenu.Enabled = False
            
            Select Case UCase(strTag)
                Case "JOURNALS"
                    g.TnJournal.ShowJournals
                    
                Case "SETTINGS"
                    frmTTSummaryCfg.ShowMe
                    
                Case "TRADEHISTORY"
                    frmTTAccounts.ShowMe True
                    
            End Select
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.m_tmrMenu_Timer"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActivityLog.Class_Terminate"
    
End Sub

