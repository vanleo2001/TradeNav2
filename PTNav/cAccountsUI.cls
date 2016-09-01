VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccountsUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAccountsUI.cls
'' Description: Wrapper class for an accounts grid
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/15/2010   DAJ         Fixed the grid column information persistence
'' 03/16/2010   DAJ         Fixed blank accounts after delete, no delete last sim
'' 03/17/2010   DAJ         Update the number visible items for the toolbar captions
'' 05/25/2010   DAJ         Additional check to avoid error in RefreshPrices
'' 09/13/2010   DAJ         Added code for Rithmic
'' 10/26/2010   DAJ         Efficiency enhancements
'' 11/01/2010   DAJ         Added Optimus, OpVest, and Vision (Rithmic Brokers)
'' 12/10/2010   DAJ         Added Zen-Fire (Rithmic Broker)
'' 12/15/2010   DAJ         Added ShowAccountOnline for Xpress
'' 03/07/2011   DAJ         Added Change Password menu item, Reworked EnableControls
'' 05/11/2011   DAJ         Utilize CheckSimTradeStatus, EnableCheckStatusMenu, IsLiveAccount
'' 06/21/2011   DAJ         Separate out Simulated trading types, SetPictureForAccount enhancements
'' 06/24/2011   DAJ         Fixed "click here" lines
'' 06/28/2011   DAJ         Setup clickable cells like hyperlinks
'' 07/25/2011   DAJ         On account delete, remove auto trade items for that account
'' 08/10/2011   DAJ         Add a "click here to create new account" line to the grid
'' 09/23/2011   DAJ         Show date journals form instead of old journals form
'' 10/04/2011   DAJ         Call the ShowJournals function instead of calling the form direct
'' 10/21/2011   DAJ         On delete of account, tell broker info object to reload
'' 04/05/2012   DAJ         Added broker mode code
'' 08/03/2012   DAJ         Remove Alaron
'' 10/29/2012   DAJ         Employ Tim's new fix for grid scrolling vs. streaming issue
'' 01/22/2013   DAJ         Delete any active TradeSense order groups for account being deleted
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/10/2014   DAJ         Consolidate the delete account code
'' 07/30/2015   DAJ         Don't call TransActSimulatedAccount unless it is necessary
'' 01/28/2016   DAJ         Moved delete account code to cBrokerDispatch
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kMenuPrefix = "A:"

Private WithEvents m_fgAccounts As VSFlexGrid
Attribute m_fgAccounts.VB_VarHelpID = -1
Private WithEvents m_tmrMenu As Timer
Attribute m_tmrMenu.VB_VarHelpID = -1
Private WithEvents m_mnuConnect As Menu
Attribute m_mnuConnect.VB_VarHelpID = -1
Private WithEvents m_mnuDisconnect As Menu
Attribute m_mnuDisconnect.VB_VarHelpID = -1
Private WithEvents m_mnuSwitchAccounts As Menu
Attribute m_mnuSwitchAccounts.VB_VarHelpID = -1
Private WithEvents m_mnuSwitchAccountsMode As Menu
Attribute m_mnuSwitchAccountsMode.VB_VarHelpID = -1
Private WithEvents m_mnuConnectInfo As Menu
Attribute m_mnuConnectInfo.VB_VarHelpID = -1
Private WithEvents m_mnuChangePassword As Menu
Attribute m_mnuChangePassword.VB_VarHelpID = -1
Private WithEvents m_mnuRefresh As Menu
Attribute m_mnuRefresh.VB_VarHelpID = -1
Private WithEvents m_mnuViewActivity As Menu
Attribute m_mnuViewActivity.VB_VarHelpID = -1
Private WithEvents m_mnuBrokerView As Menu
Attribute m_mnuBrokerView.VB_VarHelpID = -1
Private WithEvents m_mnuViewOnline As Menu
Attribute m_mnuViewOnline.VB_VarHelpID = -1
Private WithEvents m_mnuVerifyPositions As Menu
Attribute m_mnuVerifyPositions.VB_VarHelpID = -1
Private WithEvents m_mnuAccountDetails As Menu
Attribute m_mnuAccountDetails.VB_VarHelpID = -1
Private WithEvents m_mnuNewAccount As Menu
Attribute m_mnuNewAccount.VB_VarHelpID = -1
Private WithEvents m_mnuEditAccount As Menu
Attribute m_mnuEditAccount.VB_VarHelpID = -1
Private WithEvents m_mnuDeleteAccount As Menu
Attribute m_mnuDeleteAccount.VB_VarHelpID = -1
Private WithEvents m_mnuReports As Menu
Attribute m_mnuReports.VB_VarHelpID = -1
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
Private WithEvents m_mnuAutoSizeColumns As Menu
Attribute m_mnuAutoSizeColumns.VB_VarHelpID = -1
Private WithEvents m_mnuDefaultColumns As Menu
Attribute m_mnuDefaultColumns.VB_VarHelpID = -1

Private Enum eGDCols
    eGDCol_AccountID = 0
    eGDCol_AccountType
    eGDCol_Remove
    
    eGDCol_Connected
    eGDCol_Name
    eGDCol_Balance
    eGDCol_OpenEquity
    eGDCol_CurrentValue
    eGDCol_NumCols
End Enum

Private Type mPrivate
    strSource As String                 ' Name for the calling form
    TCGrid As cTradeConsoleGrid         ' Trade console grid object for generic routines
    tmrRealTime As Timer                ' Realtime timer control
    frm As Form                         ' Form that the grid is on
    SnapshotAccounts As cGdTree         ' Collection of snapshot accounts collections by broker
    mnuAccounts As Object               ' Accounts menu
    mnuSep1 As Menu                     ' Menu Separator
    mnuSep2 As Menu                     ' Menu Separator
    
    bSummary As Boolean                 ' Is this a summary grid?
    strSummaryDisplay As String         ' Summary display string
    strDisplay As String                ' Non-summary display string
    Status As cGdTree                   ' Collection of connection status'
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

Public Property Get Summary() As Boolean
    Summary = m.bSummary
End Property
Public Property Let Summary(ByVal bSummary As Boolean)
    If bSummary <> m.bSummary Then
        m.bSummary = bSummary
        If m.bSummary Then
            m.TCGrid.Display = m.strSummaryDisplay
        Else
            m.TCGrid.Display = m.strDisplay
        End If
        ApplyDisplayString
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object
'' Inputs:      Source, Working Orders UI, Summary?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal strSource As String, UI As cAccountsControls, ByVal bSummary As Boolean)
On Error GoTo ErrSection:

    m.strSource = strSource
    m.bSummary = bSummary
    
    With UI
        Set m.frm = .frm
        Set m_fgAccounts = .fgGrid
        Set m.tmrRealTime = .tmrRealTime
        Set m_tmrMenu = .tmrMenu
        Set m.mnuAccounts = .mnuAccounts
        Set m_mnuConnect = .mnuConnect
        Set m_mnuDisconnect = .mnuDisconnect
        Set m_mnuSwitchAccounts = .mnuSwitchAccounts
        Set m_mnuSwitchAccountsMode = .mnuSwitchAccountsMode
        Set m_mnuConnectInfo = .mnuConnectInfo
        Set m_mnuChangePassword = .mnuChangePassword
        Set m_mnuRefresh = .mnuRefresh
        Set m_mnuViewActivity = .mnuViewActivity
        Set m_mnuBrokerView = .mnuBrokerView
        Set m_mnuViewOnline = .mnuViewOnline
        Set m_mnuVerifyPositions = .mnuVerifyPositions
        Set m_mnuAccountDetails = .mnuAccountDetails
        Set m.mnuSep1 = .mnuSep1
        Set m_mnuNewAccount = .mnuNewAccount
        Set m_mnuEditAccount = .mnuEditAccount
        Set m_mnuDeleteAccount = .mnuDeleteAccount
        Set m.mnuSep2 = .mnuSep2
        Set m_mnuReports = .mnuReports
        Set m_mnuPrint = .mnuPrint
        Set m_mnuTradeHistory = .mnuTradeHistory
        Set m_mnuSettings = .mnuSettings
        Set m_mnuCheckStatus = .mnuCheckStatus
        Set m_mnuViewJournals = .mnuViewJournals
        Set m_mnuAutoSizeColumns = .mnuAutoSizeColumns
        Set m_mnuDefaultColumns = .mnuDefaultColumns
    End With

    InitGrid
    
    m.strSummaryDisplay = FixDisplay(GetIniFileProperty("SummaryAccountsDisplay", "", "TTSummary", g.strIniFile), True)
    m.strDisplay = FixDisplay(GetIniFileProperty("AccountsDisplay", "", "TTSummary", g.strIniFile), False)
    
    Set m.TCGrid = New cTradeConsoleGrid
    If m.bSummary Then
        m.TCGrid.Init m_fgAccounts, m.tmrRealTime, m.strSummaryDisplay
    Else
        m.TCGrid.Init m_fgAccounts, m.tmrRealTime, m.strDisplay
    End If
    
    ApplyDisplayString
    
    AddNewAccountLine

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.Init"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPrices
'' Description: Refresh the prices in the grid as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPrices()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim OldAccount As cPtAccount        ' Current account in the grid
    Dim NewAccount As cPtAccount        ' New account copy
    
    With m_fgAccounts
        For lIndex = .FixedRows To .Rows - 1
            ' DAJ 08/06/2007: Only need to update visible rows...
            If .RowHidden(lIndex) = False Then
                Set OldAccount = AccountForRow(lIndex)
                If Not OldAccount Is Nothing Then
                    If OldAccount.AccountID > 0 Then
                        ' DAJ 05/25/2010: The MakeCopy causes an error if the Account
                        ' function returns Nothing...
                        If Not g.Broker.Account(OldAccount.AccountID) Is Nothing Then
                            Set NewAccount = g.Broker.Account(OldAccount.AccountID).MakeCopy
                            If Not NewAccount Is Nothing Then
                                If NewAccount.AccountID > 0 Then
                                    AccountToGrid NewAccount, lIndex
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.RefreshPrices"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Update
'' Description: Update the data
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Update(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim Accounts As cPtAccounts         ' Accounts collection
    Dim bUpdate As Boolean              ' Do we need to update the UI?

    bUpdate = False
    Set Accounts = g.Broker.SnapshotAccountsForBroker(nBroker)
    If Not Accounts Is Nothing Then
        If m.SnapshotAccounts(nBroker) Is Nothing Then
            bUpdate = True
        ElseIf Accounts.LastModified > m.SnapshotAccounts(nBroker).LastModified Then
            bUpdate = True
        End If
    End If
    
    If bUpdate Then
        ' Get the latest collections for the broker...
        Set m.SnapshotAccounts(nBroker) = Accounts
        
        With m_fgAccounts
            .Redraw = flexRDNone
            
            ' Update the orders grid...
            ResetRemoveAccountFlag nBroker
            RefreshBrokerAccounts nBroker, False
            RemoveFlaggedAccounts nBroker
            
            .Redraw = flexRDBuffered
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.Update"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyDisplayString
'' Description: Apply the display string from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ApplyDisplayString()
On Error GoTo ErrSection:

    m.TCGrid.ApplyDisplayString eGDCol_Name

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.ApplyDisplayString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterAccountsGrid
'' Description: Hide accounts that we do not want showing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FilterAccountsGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Account As cPtAccount           ' Account object
    Dim lNumVisible As Long             ' Number of visible accounts
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With m_fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lNumVisible = 0&
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cPtAccount Then
                Set Account = .RowData(lIndex)
                
                If (Account.AccountType = eTT_AccountType_TransAct) And (UCase(Account.UserName) = "SIMUSER") Then
                    .RowHidden(lIndex) = True
                ElseIf (Account.AccountType <> -1&) Then
                    .RowHidden(lIndex) = g.Broker.HideBroker(Account.AccountType, Account.AccountID, True)
                    If (.RowHidden(lIndex) = False) And (Account.AccountID > 0) Then
                        lNumVisible = lNumVisible + 1&
                    End If
                End If
            End If
        Next lIndex
        
        .ColHidden(AccountCol(eGDCol_Connected)) = False ' Not HasVisibleBrokerAccounts
        SetBackColors m_fgAccounts
        m.TCGrid.AutoSizeGrid
        
        If m.TCGrid.CurrentRowValid Then
            .ShowCell .Row, AccountCol(eGDCol_Connected)
        ElseIf .Rows > .FixedRows Then
            .ShowCell .FixedRows, AccountCol(eGDCol_Connected)
        End If
        
        .Redraw = nRedraw
    End With
    
    If Not g.ConsoleForms Is Nothing Then
        g.ConsoleForms.NumVisible(eGDConsoleForm_Accounts) = lNumVisible
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.FilterAccountsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAccountPictures
'' Description: Set the account pictures for all rows in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAccountPictures()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With m_fgAccounts
        For lIndex = .FixedRows To .Rows - 1
            SetPictureForAccount lIndex
        Next lIndex
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.SetAccountPictures"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateConsoleSettings
'' Description: Update the console settings from the configuration form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateConsoleSettings()
On Error GoTo ErrSection:

    ' Update the font on all of the grids...
    FontFromString m_fgAccounts.Font, g.Broker.GridFont
    m_fgAccounts.Font = m_fgAccounts.Font
    
    ' Either show or hide the cents on all of the currency columns...
    If g.Broker.ShowCents Then
        m_fgAccounts.ColFormat(AccountCol(eGDCol_Balance)) = "$#,##0.00"
        m_fgAccounts.ColFormat(AccountCol(eGDCol_OpenEquity)) = "$#,##0.00"
        m_fgAccounts.ColFormat(AccountCol(eGDCol_CurrentValue)) = "$#,##0.00"
        m.TCGrid.AutoSizeGrid
    Else
        m_fgAccounts.ColFormat(AccountCol(eGDCol_Balance)) = "$#,##0"
        m_fgAccounts.ColFormat(AccountCol(eGDCol_OpenEquity)) = "$#,##0"
        m_fgAccounts.ColFormat(AccountCol(eGDCol_CurrentValue)) = "$#,##0"
        m.TCGrid.AutoSizeGrid
    End If
    
    ' Update the columns displayed on the grids on the summary tab...
    If m.bSummary Then
        m.strSummaryDisplay = FixDisplay(GetIniFileProperty("SummaryAccountsDisplay", "", "TTSummary", g.strIniFile), True)
        m.TCGrid.Display = m.strSummaryDisplay
        ApplyDisplayString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.UpdateConsoleSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedAccount
'' Description: Grab the selected account in the grid
'' Inputs:      None
'' Returns:     Account (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectedAccount() As cPtAccount
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account from the grid

    Set Account = Nothing
    With m_fgAccounts
        If m.TCGrid.ValidRow(.Row) Then
            If TypeOf .RowData(.Row) Is cPtAccount Then
                Set Account = .RowData(.Row)
            End If
        End If
    End With
    
    Set SelectedAccount = Account

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.SelectedAccount"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountForRow
'' Description: Grab the selected account in the grid for the given row
'' Inputs:      Row
'' Returns:     Account (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountForRow(ByVal Row As Long) As cPtAccount
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account from the grid

    Set Account = Nothing
    With m_fgAccounts
        If m.TCGrid.ValidRow(Row) Then
            If TypeOf .RowData(Row) Is cPtAccount Then
                Set Account = .RowData(Row)
            End If
        End If
    End With
    
    Set AccountForRow = Account

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.AccountForRow"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountForID
'' Description: Grab the selected account in the grid for the given ID
'' Inputs:      ID
'' Returns:     Account (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountForID(ByVal lAccountID As Long) As cPtAccount
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account from the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim RetAccount As cPtAccount        ' Account to return

    Set RetAccount = Nothing
    With m_fgAccounts
        For lIndex = .FixedRows To .Rows - 1
            Set Account = AccountForRow(lIndex)
            If Not Account Is Nothing Then
                If Account.AccountID = lAccountID Then
                    Set RetAccount = Account
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    Set AccountForID = RetAccount

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.AccountForID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TempBrokerAccount
'' Description: Create a temporary broker account for purposes of connection
'' Inputs:      Broker Type, Broker User
'' Returns:     Temporary Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TempBrokerAccount(ByVal nBroker As eTT_AccountType, ByVal bBrokerUser As Boolean)
On Error GoTo ErrSection:

    Dim Account As New cPtAccount       ' Account to return from the function
    Dim lNumAccounts As Long            ' Number of accounts for this broker
    Dim lIndex As Long                  ' Index into a for loop
    Dim lTmpIndex As Long               ' Index for the temp account
    Dim astrLogins As cGdArray          ' Array of possible logins
    Dim bIsBroker As Boolean            ' Is this user a broker?
    Dim strBrokerName As String         ' Broker name
    
    If (nBroker <> eTT_AccountType_SimStream) And (nBroker <> eTT_AccountType_SimReplay) Then
        lTmpIndex = -1&
        With m_fgAccounts
            For lIndex = .FixedRows To .Rows - 1
                If (.RowData(lIndex).AccountType = nBroker) Then
                    If (.RowData(lIndex).AccountID > 0) Then
                        ' DAJ 07/30/2015: Split out this if statement because the call to TransActSimulatedAccount will cause the TransAct object to be
                        ' created for non-TransAct clients which will cause the DLL's to be registered which will cause Yes log files to be generated...
                        'If ((.RowData(lIndex).AccountNumber <> kTransActSimUserAccount) And (.RowData(lIndex).AccountNumber <> kTransActOldSimUserAccount)) Or (nBroker <> eTT_AccountType_TransAct) Then
                        If nBroker <> eTT_AccountType_TransAct Then
                            lNumAccounts = lNumAccounts + 1
                        ElseIf TransActSimulatedAccount(.RowData(lIndex).AccountNumber) = False Then
                            lNumAccounts = lNumAccounts + 1
                        End If
                    Else
                        lTmpIndex = lIndex
                    End If
                End If
            Next lIndex
        End With
        
        If bBrokerUser = True Then
            bIsBroker = g.Broker.IsBroker(nBroker)
            If ((lNumAccounts = 0) And (lTmpIndex = -1&)) Or (bIsBroker = True) Then
                With Account
                    If bIsBroker Then
                        strBrokerName = g.Broker.BrokerName(nBroker)
                        
                        .AccountID = -10&
                        .AccountNumber = strBrokerName & " Broker Management"
                        .AccountType = nBroker
                        .Name = .AccountNumber
                        
                        If (lTmpIndex <> -1&) Then m_fgAccounts.RemoveItem lTmpIndex
                    Else
                        .AccountID = -9&
                        .AccountNumber = "Click here to setup " & g.Broker.BrokerName(nBroker) & " account"
                        .AccountType = nBroker
                        .Name = .AccountNumber
                    End If
                End With
                
                AddClickHereLine Account
            ElseIf (lNumAccounts > 0) And (lTmpIndex > 0) Then
                m_fgAccounts.RemoveItem lTmpIndex
            End If
        Else
            With m_fgAccounts
                For lIndex = .Rows - 1 To .FixedRows Step -1
                    If (.RowData(lIndex).AccountType = nBroker) Then
                        If (.RowData(lIndex).AccountID > 0) Then
                            .RowHidden(lIndex) = True
                        Else
                            .RemoveItem lIndex
                        End If
                    End If
                Next lIndex
            End With
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.TempBrokerAccount"
    
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

    m.TCGrid.GenerateReport vArgs, "Accounts"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    With m_fgAccounts
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbButtonFace
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = False
        .MergeCells = flexMergeFree
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .HighLight = flexHighlightNever
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
                
        .TextMatrix(0, GDCol(eGDCol_AccountID)) = "Account ID"
        .TextMatrix(0, GDCol(eGDCol_AccountType)) = "Account Type"
        .TextMatrix(0, GDCol(eGDCol_Remove)) = "Remove"
        .TextMatrix(0, GDCol(eGDCol_Connected)) = "On"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_Balance)) = "Balance"
        .TextMatrix(0, GDCol(eGDCol_OpenEquity)) = "Open Equity"
        .TextMatrix(0, GDCol(eGDCol_CurrentValue)) = "Current Value"
        
        If g.Broker.ShowCents Then
            .ColFormat(GDCol(eGDCol_Balance)) = "$#,##0.00"
            .ColFormat(GDCol(eGDCol_OpenEquity)) = "$#,##0.00"
            .ColFormat(GDCol(eGDCol_CurrentValue)) = "$#,##0.00"
        Else
            .ColFormat(GDCol(eGDCol_Balance)) = "$#,##0"
            .ColFormat(GDCol(eGDCol_OpenEquity)) = "$#,##0"
            .ColFormat(GDCol(eGDCol_CurrentValue)) = "$#,##0"
        End If
        
        .ColHidden(GDCol(eGDCol_AccountID)) = True
        .ColHidden(GDCol(eGDCol_AccountType)) = True
        .ColHidden(GDCol(eGDCol_Remove)) = True
        
        .ColAlignment(GDCol(eGDCol_Name)) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetRemoveAccountFlag
'' Description: Set the remove account flag on all accounts for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetRemoveAccountFlag(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With m_fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If .RowData(lIndex).AccountType = nBroker Then
                CheckedCell(m_fgAccounts, lIndex, AccountCol(eGDCol_Remove)) = True
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.ResetRemoveAccountFlag"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFlaggedAccounts
'' Description: Remove rows with the remove flag set for the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveFlaggedAccounts(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With m_fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If .RowData(lIndex).AccountType = nBroker Then
                If CheckedCell(m_fgAccounts, lIndex, AccountCol(eGDCol_Remove)) = True Then
                    .RemoveItem lIndex
                End If
            End If
        Next lIndex
        
        TempBrokerAccount nBroker, Not g.Broker.HideBroker(nBroker)
        
        .Col = AccountCol(eGDCol_Name)
        .Sort = flexSortStringAscending
        AfterAccountSort False
        
        FilterAccountsGrid
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.RemoveFlaggedAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshBrokerAccounts
'' Description: Refresh accounts in the grid for the given broker from the collection
'' Inputs:      Broker, Filter grid?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshBrokerAccounts(ByVal nBroker As eTT_AccountType, Optional ByVal bFilter As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    Dim Accounts As cPtAccounts         ' Accounts collection for the broker
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    If Not m.SnapshotAccounts(nBroker) Is Nothing Then
        Set Accounts = m.SnapshotAccounts(nBroker).MakeCopy
        With m_fgAccounts
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            For lIndex = 1 To Accounts.Count
                lRow = -1&
                For lIndex2 = .FixedRows To .Rows - 1
                    If .RowData(lIndex2).AccountID = Accounts(lIndex).AccountID Then
                        lRow = lIndex2
                        Exit For
                    End If
                Next lIndex2
                AccountToGrid Accounts(lIndex), lRow
            Next lIndex
        
            If bFilter Then
                TempBrokerAccount nBroker, Not g.Broker.HideBroker(nBroker)
                
                .Col = AccountCol(eGDCol_Name)
                .Sort = flexSortStringAscending
                AfterAccountSort False
                
                FilterAccountsGrid
            End If
            
            .Redraw = nRedraw
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.RefreshBrokerAccounts"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AfterAccountSort
'' Description: Move the "Click Here" lines to the bottom after a sort
'' Inputs:      Set Back Colors?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AfterAccountSort(Optional ByVal bSetBackColors As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lCounter As Long                ' Counter variable
    
    With m_fgAccounts
        lCounter = 1&
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If (.RowData(lIndex).AccountID = -9&) Or (.RowData(lIndex).AccountID = -10&) Then
                .RowPosition(lIndex) = .Rows - lCounter
                lCounter = lCounter + 1&
            End If
        Next lIndex
    End With
    
    If bSetBackColors Then
        SetBackColors m_fgAccounts
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.AfterAccountSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountCol
'' Description: Determine where the given column is really sitting right now
'' Inputs:      Column Number
'' Returns:     Current Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AccountCol(ByVal Col As eGDCols) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    If Col = eGDCol_NumCols Then
        lReturn = Col
    Else
        lReturn = m.TCGrid.ColNumFromColData(Col)
    End If
    
    AccountCol = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.AccountCol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountToGrid
'' Description: Create a line in the grid for the given account
'' Inputs:      Account, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AccountToGrid(Account As cPtAccount, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    With m_fgAccounts
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = Account
        
        .Cell(flexcpFontUnderline, lRow, GDCol(eGDCol_Connected), lRow, GDCol(eGDCol_CurrentValue)) = False
        .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Connected), lRow, GDCol(eGDCol_CurrentValue)) = .Cell(flexcpForeColor, 0, 0)
        
        .TextMatrix(lRow, AccountCol(eGDCol_AccountID)) = Str(Account.AccountID)
        .TextMatrix(lRow, AccountCol(eGDCol_AccountType)) = Str(Account.AccountType)
        .TextMatrix(lRow, AccountCol(eGDCol_Name)) = Account.Name
        .TextMatrix(lRow, AccountCol(eGDCol_Balance)) = Account.CurrentBalance
        .TextMatrix(lRow, AccountCol(eGDCol_OpenEquity)) = Account.OpenProfit
        .TextMatrix(lRow, AccountCol(eGDCol_CurrentValue)) = Account.CurrentValue
        
        CheckedCell(m_fgAccounts, lRow, AccountCol(eGDCol_Remove)) = False
        
        m.TCGrid.ColorCell lRow, AccountCol(eGDCol_Balance)
        m.TCGrid.ColorCell lRow, AccountCol(eGDCol_OpenEquity)
        m.TCGrid.ColorCell lRow, AccountCol(eGDCol_CurrentValue)
        
        .Cell(flexcpFontUnderline, lRow, GDCol(eGDCol_Connected)) = True
        
        SetPictureForAccount lRow
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.AccountToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasVisibleBrokerAccounts
'' Description: Walk the accounts grid looking for broker accounts that are
''              not hidden
'' Inputs:      None
'' Returns:     True if Visible Broker Accounts, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasVisibleBrokerAccounts() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Account As cPtAccount           ' Temporary account variable
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    With m_fgAccounts
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then
                Set Account = AccountForRow(lIndex)
                If Not Account Is Nothing Then
                    If g.Broker.IsLiveAccount(Account.AccountType) Then
                        bReturn = True
                        Exit For
                    End If
                End If
            End If
        Next lIndex
    End With
    
    HasVisibleBrokerAccounts = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.HasVisibleBrokerAccounts"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPictureForAccount
'' Description: Set the connection picture for the given account
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPictureForAccount(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account from the grid
    Dim bChanged As Boolean             ' Has the status changed?
    Dim nStatus As eGDConnectionStatus  ' Connection status
    Dim strKey As String                ' Key into the collection
    
    With m_fgAccounts
        Set Account = AccountForRow(lRow)
        If Not Account Is Nothing Then
            If (Account.AccountID <> -9&) And (Account.AccountID <> -10&) Then
                strKey = Str(Account.AccountID)
                nStatus = g.Broker.ConnectionStatusForAccount(Account.AccountID)
                
                If m.Status.Exists(strKey) = False Then
                    bChanged = True
                ElseIf m.Status(strKey) <> nStatus Then
                    bChanged = True
                Else
                    bChanged = False
                End If
                
                If bChanged Then
                    If m.Status.Exists(strKey) = False Then
                        m.Status.Add nStatus, strKey
                    Else
                        m.Status(strKey) = nStatus
                    End If
                    
                    Select Case nStatus
                        Case eGDConnectionStatus_Disconnected
                            .Cell(flexcpPicture, lRow, AccountCol(eGDCol_Connected)) = frmOnlineBroker.imgRed
                        Case eGDConnectionStatus_Disconnecting, eGDConnectionStatus_Connecting
                            .Cell(flexcpPicture, lRow, AccountCol(eGDCol_Connected)) = frmOnlineBroker.imgYellow
                        Case eGDConnectionStatus_Connected
                            .Cell(flexcpPicture, lRow, AccountCol(eGDCol_Connected)) = frmOnlineBroker.imgGreen
                    End Select
                    
                    .Cell(flexcpPictureAlignment, lRow, AccountCol(eGDCol_Connected)) = flexAlignCenterCenter
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.SetPictureForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountTooltip
'' Description: Figure out the appropriate tooltip for the accounts grid
'' Inputs:      Row and Column of the Mouse
'' Returns:     Account Tooltip
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AccountTooltip(ByVal Row As Long, ByVal Col As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim Account As cPtAccount           ' Account object
    
    strReturn = ""
    
    Set Account = AccountForRow(Row)
    If Not Account Is Nothing Then
        If (Account.AccountID = -9&) Then
            If Account.AccountType = -1& Then
                strReturn = "Click here to create a new account"
            Else
                strReturn = "Click here to enter information for your live " & g.Broker.BrokerName(Account.AccountType) & " account"
            End If
        ElseIf (Account.AccountID = -10&) And (Col <> AccountCol(eGDCol_Connected)) Then
            strReturn = "Click here to bring up the " & Account.Name & " screen"
        ElseIf Col = AccountCol(eGDCol_Connected) Then
            strReturn = g.Broker.TooltipForAccount(Account.AccountID)
        End If
    End If
    
    AccountTooltip = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.AccountTooltip"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColumnExceptions
'' Description: Build a list of columns not included in the display string
'' Inputs:      None
'' Returns:     Exception String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ColumnExceptions() As String
On Error GoTo ErrSection:

    Dim astrColumns As cGdArray         ' Array of column numbers to join
    
    Set astrColumns = New cGdArray
    astrColumns.Create eGDARRAY_Strings
    
    astrColumns.Add Str(AccountCol(eGDCol_AccountID))
    astrColumns.Add Str(AccountCol(eGDCol_AccountType))
    astrColumns.Add Str(AccountCol(eGDCol_Connected))
    astrColumns.Add Str(AccountCol(eGDCol_Remove))
    
    ColumnExceptions = "," & astrColumns.JoinFields(",") & ","

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.ColumnExceptions"
    
End Function

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

    Dim Account As cPtAccount           ' Account object from the grid
    Dim bNotStreamReplay As Boolean     ' Are we in stream replay?
    Dim bNotBrokerMgr As Boolean        ' Is this the broker management line?
    Dim nStatus As eGDConnectionStatus  ' Current connection status for the broker
    
    bNotStreamReplay = (g.nReplaySession = 0)
    
    With m_fgAccounts
        Set Account = SelectedAccount
        If Account Is Nothing Then
            m_mnuConnect.Caption = "Connect"
            m_mnuConnect.Enabled = False
            m_mnuConnect.Visible = True
            m_mnuDisconnect.Caption = "Disconnect"
            m_mnuDisconnect.Enabled = False
            m_mnuDisconnect.Visible = True
            m_mnuSwitchAccounts.Caption = "Switch Accounts"
            m_mnuSwitchAccounts.Enabled = False
            m_mnuSwitchAccounts.Visible = False
            m_mnuSwitchAccountsMode.Visible = False
            m_mnuConnectInfo.Caption = "Connection Information"
            m_mnuConnectInfo.Enabled = False
            m_mnuConnectInfo.Visible = True
            m_mnuChangePassword.Caption = "Change Password"
            m_mnuChangePassword.Enabled = False
            m_mnuChangePassword.Visible = False
            m_mnuRefresh.Caption = "Refresh Account"
            m_mnuRefresh.Enabled = False
            m_mnuRefresh.Visible = True
            m_mnuViewActivity.Caption = "View Activity"
            m_mnuViewActivity.Enabled = False
            m_mnuViewActivity.Visible = True
            m_mnuBrokerView.Caption = "Broker View"
            m_mnuBrokerView.Enabled = False
            m_mnuBrokerView.Visible = False
            m_mnuViewOnline.Caption = "View Account Online"
            m_mnuViewOnline.Enabled = False
            m_mnuViewOnline.Visible = False
            m_mnuVerifyPositions.Enabled = False
            m_mnuVerifyPositions.Visible = False
            m_mnuAccountDetails.Enabled = False
            m_mnuAccountDetails.Visible = False
        
            m.mnuSep1.Visible = bNotStreamReplay And bNotBrokerMgr
            m_mnuNewAccount.Caption = "New Account"
            m_mnuNewAccount.Enabled = True
            m_mnuNewAccount.Visible = bNotStreamReplay And bNotBrokerMgr
            m_mnuEditAccount.Caption = "Edit Account"
            m_mnuEditAccount.Enabled = False
            m_mnuEditAccount.Visible = False
            m_mnuDeleteAccount.Caption = "Delete Account"
            m_mnuDeleteAccount.Enabled = False
            m_mnuDeleteAccount.Visible = False
            m.mnuSep2.Visible = bNotStreamReplay
            m_mnuReports.Visible = True
            m_mnuReports.Enabled = False
            
            If Not m_mnuCheckStatus Is Nothing Then
                m_mnuCheckStatus.Enabled = False
                m_mnuCheckStatus.Visible = False
            End If
        Else
            bNotBrokerMgr = (Account.AccountID <> -10&)
            nStatus = g.Broker.ConnectionStatusForBroker(Account.AccountType)
            
            m_mnuConnect.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_Connect)
            m_mnuConnect.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_Connect)
            m_mnuConnect.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_Connect)
            m_mnuDisconnect.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_Disconnect)
            m_mnuDisconnect.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_Disconnect)
            m_mnuDisconnect.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_Disconnect)
            m_mnuSwitchAccounts.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_SwitchAccounts)
            m_mnuSwitchAccounts.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_SwitchAccounts)
            m_mnuSwitchAccounts.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_SwitchAccounts)
            m_mnuSwitchAccountsMode.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_SwitchAccountsMode)
            m_mnuSwitchAccountsMode.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_SwitchAccountsMode)
            m_mnuSwitchAccountsMode.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_SwitchAccountsMode)
            m_mnuConnectInfo.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_ConnectInfo)
            m_mnuConnectInfo.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_ConnectInfo)
            m_mnuConnectInfo.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_ConnectInfo)
            m_mnuChangePassword.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_ChangePassword)
            m_mnuChangePassword.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_ChangePassword)
            m_mnuChangePassword.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_ChangePassword)
            m_mnuRefresh.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_Refresh)
            m_mnuRefresh.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_Refresh)
            m_mnuRefresh.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_Refresh)
            m_mnuViewActivity.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_ViewActivity)
            m_mnuViewActivity.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_ViewActivity)
            m_mnuViewActivity.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_ViewActivity)
            m_mnuBrokerView.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_BrokerView)
            m_mnuBrokerView.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_BrokerView)
            m_mnuBrokerView.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_BrokerView)
            m_mnuViewOnline.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_ViewOnline)
            m_mnuViewOnline.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_ViewOnline)
            m_mnuViewOnline.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_ViewOnline)
            m_mnuVerifyPositions.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_VerifyPositions)
            m_mnuVerifyPositions.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_VerifyPositions)
            m_mnuVerifyPositions.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_VerifyPositions)
            m_mnuAccountDetails.Caption = g.Broker.TradingMenuItemName(Account.AccountType, eGDTradingMenu_AccountDetails)
            m_mnuAccountDetails.Enabled = g.Broker.EnableTradingMenuItem(Account.AccountType, eGDTradingMenu_AccountDetails)
            m_mnuAccountDetails.Visible = g.Broker.ShowTradingMenuItem(Account.AccountType, eGDTradingMenu_AccountDetails)
            
            m.mnuSep1.Visible = bNotStreamReplay And bNotBrokerMgr
            m_mnuNewAccount.Caption = "New Account"
            m_mnuNewAccount.Enabled = True
            m_mnuNewAccount.Visible = bNotStreamReplay And bNotBrokerMgr
            m_mnuEditAccount.Caption = "Edit '" & Account.Name & "'"
            m_mnuEditAccount.Enabled = True
            m_mnuEditAccount.Visible = bNotStreamReplay And (Account.AccountID > 0) ' And bNotBrokerMgr
            m_mnuDeleteAccount.Caption = "Delete '" & Account.Name & "'"
            m_mnuDeleteAccount.Enabled = True
            m_mnuDeleteAccount.Visible = bNotStreamReplay And (Account.AccountID > 0) ' And bNotBrokerMgr
            m.mnuSep2.Visible = bNotStreamReplay
            m_mnuReports.Visible = True '(Account.AccountID > 0)
            m_mnuReports.Enabled = True
            
            If Not m_mnuCheckStatus Is Nothing Then
                m_mnuCheckStatus.Enabled = False
                m_mnuCheckStatus.Visible = False
            End If
        End If
    End With
    
    m_mnuDefaultColumns.Enabled = (m.TCGrid.Display <> DefaultDisplay)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SmartColumns
'' Description: Determine what columns should be shown/hidden based on the
''              current screen width
'' Inputs:      None
'' Returns:     Display string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SmartColumns() As String
On Error GoTo ErrSection:

    Dim astrDisplay As New cGdArray     ' Array of field display information

    Select Case Screen.Width / Screen.TwipsPerPixelX
        Case Is >= 1600
            'strReturn = "1;Account;1;3,1;Balance;1;4,1;Open Equity;1;5,1;Current Value;1;6"
            astrDisplay.Add "0;Account ID;1;0;0;0"
            astrDisplay.Add "0;Account Type;1;1;0;0"
            astrDisplay.Add "0;Remove;1;2;0;0"
            astrDisplay.Add "1;On;1;3;0;0"
            astrDisplay.Add "1;Account;1;4;0;1"
            astrDisplay.Add "1;Balance;1;5;0;1"
            astrDisplay.Add "1;Open Equity;1;6;0;1"
            astrDisplay.Add "1;Current Value;1;7;0;1"
        
        Case Is >= 1280
            'strReturn = "1;Account;1;3,1;Balance;1;4,1;Open Equity;1;5,0;Current Value;1;6"
            astrDisplay.Add "0;Account ID;1;0;0;0"
            astrDisplay.Add "0;Account Type;1;1;0;0"
            astrDisplay.Add "0;Remove;1;2;0;0"
            astrDisplay.Add "1;On;1;3;0;0"
            astrDisplay.Add "1;Account;1;4;0;1"
            astrDisplay.Add "1;Balance;1;5;0;1"
            astrDisplay.Add "1;Open Equity;1;6;0;1"
            astrDisplay.Add "0;Current Value;1;7;0;1"
        
        Case Is >= 1152
            'strReturn = "1;Account;1;3,1;Balance;1;4,1;Open Equity;1;5,0;Current Value;1;6"
            astrDisplay.Add "0;Account ID;1;0;0;0"
            astrDisplay.Add "0;Account Type;1;1;0;0"
            astrDisplay.Add "0;Remove;1;2;0;0"
            astrDisplay.Add "1;On;1;3;0;0"
            astrDisplay.Add "1;Account;1;4;0;1"
            astrDisplay.Add "1;Balance;1;5;0;1"
            astrDisplay.Add "1;Open Equity;1;6;0;1"
            astrDisplay.Add "0;Current Value;1;7;0;1"
        
        Case Else
            'strReturn = "1;Account;1;3,1;Balance;1;4,1;Open Equity;1;5,0;Current Value;1;6"
            astrDisplay.Add "0;Account ID;1;0;0;0"
            astrDisplay.Add "0;Account Type;1;1;0;0"
            astrDisplay.Add "0;Remove;1;2;0;0"
            astrDisplay.Add "1;On;1;3;0;0"
            astrDisplay.Add "1;Account;1;4;0;1"
            astrDisplay.Add "1;Balance;1;5;0;1"
            astrDisplay.Add "1;Open Equity;1;6;0;1"
            astrDisplay.Add "0;Current Value;1;7;0;1"
    
    End Select
    
    SmartColumns = "1|" & astrDisplay.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.SmartColumns"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_AfterMoveColumn
'' Description: After a user moves a column, resave the display string
'' Inputs:      Column moved, Position moved to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_fgAccounts_AfterMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_AfterRowColChange
'' Description: Make sure to show cell after a row/col change
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With m_fgAccounts
        'if no horiz scrollbars, we need to "manually scroll"
        'by forcing the cell to show when changing columns
        If .Redraw <> flexRDNone And .ScrollBars <> flexScrollBarBoth Then
            If OldCol <> NewCol Then
                If (NewRow >= 0) And (NewCol >= 0) Then
                    .ShowCell NewRow, NewCol
                End If
            End If
        End If
    End With
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_fgAccounts_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_AfterSort
'' Description: Set the background colors after a sort
'' Inputs:      Column of Sort, Order of Sort
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    AfterAccountSort
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_fgAccounts_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_AfterUserResize
'' Description: Rebuild the display string after a user resize
'' Inputs:      Row and Column of resize
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        BuildDisplayString True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_fgAccounts_AfterUserResize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_BeforeMouseDown
'' Description: Bring up the Pop-Up menu when the user right clicks
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    m.TCGrid.SelectRow m_fgAccounts.MouseRow
    
    If Button = vbRightButton Then
        EnableControls
        m.frm.PopupMenu m.mnuAccounts
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_fgAccounts_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_BeforeMoveColumn
'' Description: Make sure that the frozen columns stay where they are
'' Inputs:      Column to move, Position to move it to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.frm.Visible Then
        With m_fgAccounts
            If Col <= eGDCol_Connected Then
                Position = Col
            ElseIf Position <= eGDCol_Connected Then
                Position = eGDCol_Connected + 1
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_fgAccounts_BeforeMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_BeforeScroll
'' Description: Make sure left col stays the same if no horizontal
'' Inputs:      Old Top Row, Old Left Col, New Top Row, New Left Col, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    GridScrollCheck m_fgAccounts, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_Click
'' Description: If the user clicks on the connect column, toggle the connection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim lMouseCol As Long               ' Mouse column in the grid
    Dim Account As cPtAccount           ' Account object
    
    With m_fgAccounts
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        Set Account = AccountForRow(lMouseRow)
        If Not Account Is Nothing Then
            If (Account.AccountID = -9&) Then
                If Account.AccountType = -1& Then
                    frmTTPositions.ShowMe 0&
                ElseIf Account.AccountType = eTT_AccountType_SimBroker Then
                    frmTTPositions.ShowMe 0&, eTT_AccountType_SimBroker
                Else
                    g.Broker.ShowBrokerConnectionInfo Account.AccountType, True
                End If
            ElseIf (Account.AccountID = -10&) And (lMouseCol <> AccountCol(eGDCol_Connected)) Then
                g.Broker.ShowBrokerView Account.AccountType
            ElseIf (lMouseCol = AccountCol(eGDCol_Connected)) Then
                If (Account.AccountType = eTT_AccountType_SimStream) And (g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected) Then
                    InfBox "To disconnect from this account,|you will need to stop streaming", "i", , "Disconnect"
                Else
                    g.Broker.ToggleConnectionForAccount Account.AccountID, Account.UserName, "User Disconnecting from " & m.strSource & " Accounts Grid"
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_fgAccounts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_DblClick
'' Description: If the user double clicks on an account, bring it up in TT
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim lAccountID As Long              ' Account ID to edit
    
    With m_fgAccounts
        lAccountID = -1&
        lMouseRow = .MouseRow
        
        If m.TCGrid.ValidRow(lMouseRow) Then
            lAccountID = Val(.TextMatrix(lMouseRow, AccountCol(eGDCol_AccountID)))
        End If
    End With
    
    If lAccountID > 0 Then
        frmTTPositions.ShowMe lAccountID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_fgAccounts_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_KeyUp
'' Description: Delete cancels an order, Insert creates a new one
'' Inputs:      Code of the Key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account object

    Select Case KeyCode
        Case vbKeyDelete
            If m.TCGrid.ValidRow(m_fgAccounts.Row) Then
                DeleteAccount
            End If
        
        Case vbKeyInsert
            frmTTPositions.ShowMe 0&
            
'        Case vbKeyReturn
'            Set Account = SelectedAccount
'            If Not Account Is Nothing Then
'                frmTTPositions.ShowMe Account.AccountID, Account.AccountType
'            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_fgAccounts_KeyUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgAccounts_MouseMove
'' Description: Set the tooltip text based on where the cursor is
'' Inputs:      Mouse Button Down, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgAccounts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    m_fgAccounts.ToolTipText = AccountTooltip(m_fgAccounts.MouseRow, m_fgAccounts.MouseCol)
    m.TCGrid.SetMouseCursor m.frm

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuAccountDetails_Click
'' Description: Allow the user to view extra account information for the
''              selected account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuAccountDetails_Click()
On Error GoTo ErrSection:

    StartMenuTimer "ACCOUNTDETAILS"
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAccountsUI.m_mnuAccountDetails_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuAutoSizeColumns_Click
'' Description: Allow the user to auto size the columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuAutoSizeColumns_Click()
On Error GoTo ErrSection:

    m.TCGrid.AutoSizeGrid True
    BuildDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuAutoSizeColumns_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuAccountBrokerView_Click
'' Description: Allow the user to view the broker screen
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuBrokerView_Click()
On Error GoTo ErrSection:

    StartMenuTimer "BROKERVIEW"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuBrokerView_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuChangePassword_Click
'' Description: Allow the user to change their broker password
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuChangePassword_Click()
On Error GoTo ErrSection:

    StartMenuTimer "CHANGEPASSWORD"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuChangePassword_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuCheckStatus_Click
'' Description: Check the status of any open orders
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
    RaiseError "cAccountsUI.m_mnuCheckStatus_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuConnect_Click
'' Description: Allow the user to connect to an online broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuConnect_Click()
On Error GoTo ErrSection:

    StartMenuTimer "CONNECT"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuConnect_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuConnectInfo_Click
'' Description: Allow the user to edit their online broker connection settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuConnectInfo_Click()
On Error GoTo ErrSection:

    StartMenuTimer "CONNECTINFO"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuConnectInfo_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuDefaultColumns_Click
'' Description: Allow the user to revert back to the default column information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuDefaultColumns_Click()
On Error GoTo ErrSection:

    If m.bSummary Then
        m.strSummaryDisplay = DefaultDisplay
        m.TCGrid.Display = m.strSummaryDisplay
    Else
        m.strDisplay = DefaultDisplay
        m.TCGrid.Display = m.strDisplay
    End If
    
    ApplyDisplayString
    BuildDisplayString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuDefaultColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuDeleteAccount_Click
'' Description: Allow the user to delete an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuDeleteAccount_Click()
On Error GoTo ErrSection:

    DeleteAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuDeleteAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuDisconnect_Click
'' Description: Allow the user to disconnect from an online broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuDisconnect_Click()
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account from the grid

    Set Account = SelectedAccount
    If Not Account Is Nothing Then
        g.Broker.Disconnect Account.AccountType, "User disconnecting from " & m.strSource & " Menu"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuDisconnect_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuEditAccount_Click
'' Description: Allow the user to edit their account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuEditAccount_Click()
On Error GoTo ErrSection:

    StartMenuTimer "EDITACCOUNT"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuEditAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuNewAccount_Click
'' Description: Allow the user to add a new account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuNewAccount_Click()
On Error GoTo ErrSection:

    StartMenuTimer "NEWACCOUNT"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuNewAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuPrint_Click
'' Description: Allow the user to print the orders, positions, and accounts
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
    RaiseError "cAccountsUI.m_mnuPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuRefresh_Click
'' Description: Allow the user to refresh an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuRefresh_Click()
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account from the grid

    Set Account = SelectedAccount
    If Not Account Is Nothing Then
        g.Broker.Refresh Account.AccountType, True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuRefresh_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuReports_Click
'' Description: Allow the user to view performance reports on trades
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuReports_Click()
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account from the grid
    Dim lAccountID As Long              ' Account ID

    With m_fgAccounts
        lAccountID = -1&
        
        If m.TCGrid.ValidRow(.Row) Then
            If TypeOf .RowData(.Row) Is cPtAccount Then
                Set Account = .RowData(.Row)
                If Not Account Is Nothing Then
                    lAccountID = Account.AccountID
                End If
            End If
        End If
        
        ShowTradeFilter lAccountID
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuReports_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSettings_Click
'' Description: Allow the user to view the Trade Console settings
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
    RaiseError "cAccountsUI.m_mnuSettings_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSwitchAccounts_Click
'' Description: Allow the user to switch accounts for a broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuSwitchAccounts_Click()
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account from the grid

    Set Account = SelectedAccount
    If Not Account Is Nothing Then
        g.Broker.SwitchBrokerAccount Account.AccountType, Account.UserName, Account.AccountNumber
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuSwitchAccounts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuSwitchAccountsMode_Click
'' Description: Allow the user to switch login modes for a broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuSwitchAccountsMode_Click()
On Error GoTo ErrSection:

    If Not g.Transact Is Nothing Then
        Select Case g.Transact.LoginMode
            Case eGDTransActLoginMode_Live
                g.Transact.SwitchLoginModes eGDTransActLoginMode_SimLive
            Case eGDTransActLoginMode_SimLive
                g.Transact.SwitchLoginModes eGDTransActLoginMode_Live
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuSwitchAccountsMode_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuTradeHistory_Click
'' Description: Allow the user to view trade history
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
    RaiseError "cAccountsUI.m_mnuTradeHistory_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuVerifyPositions_Click
'' Description: Allow the user to verify their positions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuVerifyPositions_Click()
On Error GoTo ErrSection:

    StartMenuTimer "VERIFYPOSITIONS"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuVerifyPositions_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuViewActivity_Click
'' Description: Allow the user to view the activity view
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuViewActivity_Click()
On Error GoTo ErrSection:

    StartMenuTimer "VIEWACTIVITY"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuViewActivity_Click"
    
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
    RaiseError "cAccountsUI.m_mnuViewJournals_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_mnuViewOnline_Click
'' Description: Allow the user to view their account online
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_mnuViewOnline_Click()
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Selected account
    
    Set Account = SelectedAccount
    If Not Account Is Nothing Then
        g.Broker.ShowAccountOnline Account.AccountType
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_mnuViewOnline_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_tmrMenu_Timer
'' Description: Do actions that require a modal dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Action to perform
    Dim Account As cPtAccount           ' Selected account

    strTag = m_tmrMenu.Tag
    If Len(strTag) > 2 Then
        If Left(strTag, 2) = kMenuPrefix Then
            strTag = Mid(strTag, 3)
            m_tmrMenu.Tag = ""
            m_tmrMenu.Enabled = False
            
            Set Account = SelectedAccount
            
            Select Case UCase(strTag)
                Case "ACCOUNTDETAILS"
                    If Not Account Is Nothing Then
                        g.Broker.ShowAccountInfoForm Account.AccountNumber
                    End If
                    
                Case "BROKERVIEW"
                    If Not Account Is Nothing Then
                        g.Broker.ShowBrokerView Account.AccountType
                    End If
                    
                Case "CHANGEPASSWORD"
                    If Not Account Is Nothing Then
                        g.Broker.ChangeBrokerPassword Account.AccountType
                    End If
                    
                Case "CONNECT"
                    If Not Account Is Nothing Then
                        g.Broker.Connect Account.AccountType
                    End If
                    
                Case "CONNECTINFO"
                    If Not Account Is Nothing Then
                        g.Broker.ShowBrokerConnectionInfo Account.AccountType, (Account.AccountID = -9&)
                    End If
                    
                Case "EDITACCOUNT"
                    If Not Account Is Nothing Then
                        frmTTPositions.ShowMe Account.AccountID, Account.AccountType
                    End If
                
                Case "JOURNALS"
                    g.TnJournal.ShowJournals
                    
                Case "NEWACCOUNT"
                    frmTTPositions.ShowMe 0&
                    
                Case "SETTINGS"
                    frmTTSummaryCfg.ShowMe
                    
                Case "TRADEHISTORY"
                    If Account Is Nothing Then
                        frmTTAccounts.ShowMe True
                    Else
                        frmTTPositions.ShowMe Account.AccountID, Account.AccountType
                    End If
                    
                Case "VERIFYPOSITIONS"
                    If Not Account Is Nothing Then
                        g.Broker.VerifyPositions Account.AccountType
                    End If
                    
                Case "VIEWACTIVITY"
                    If Not Account Is Nothing Then
                        g.Broker.ShowActivityView Account.AccountType
                    End If
                    
            End Select
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.m_tmrMenu_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DefaultDisplay
'' Description: Determine the default display
'' Inputs:      None
'' Returns:     Default Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DefaultDisplay() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrDisplay As New cGdArray     ' Array of field display information
    
    If m.bSummary Then
        strReturn = SmartColumns
    End If
    
    If Len(strReturn) = 0 Then
        astrDisplay.Add "0;Account ID;1;0;0;0"
        astrDisplay.Add "0;Account Type;1;1;0;0"
        astrDisplay.Add "0;Remove;1;2;0;0"
        astrDisplay.Add "1;On;1;3;0;0"
        astrDisplay.Add "1;Account;1;4;0;1"
        astrDisplay.Add "1;Balance;1;5;0;1"
        astrDisplay.Add "1;Open Equity;1;6;0;1"
        astrDisplay.Add "1;Current Value;1;7;0;1"

        'strReturn = "1;Account;1;3,1;Balance;1;4,1;Open Equity;1;5,1;Current Value;1;6"
        strReturn = "1|" & astrDisplay.JoinFields(",")
    End If

    DefaultDisplay = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.DefaultDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixDisplay
'' Description: Fix the given display string
'' Inputs:      Display, Summary?
'' Returns:     Fixed Display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixDisplay(ByVal strDisplay As String, ByVal bSummary As Boolean) As String
On Error GoTo ErrSection:

    Dim astrDisplay As cGdArray         ' Display string split out into an array
    Dim astrLine As cGdArray            ' Line of information split out into an array
    Dim lIndex As Long                  ' Index into a for loop

    If Len(strDisplay) = 0 Then
        strDisplay = DefaultDisplay
    End If
    
    FixDisplay = strDisplay

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountsUI.FixDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildDisplayString
'' Description: Tell the grid class to build the display string
'' Inputs:      Force Sum Calc?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildDisplayString(Optional ByVal bForceSumCalc As Boolean = False)
On Error GoTo ErrSection:

    m.TCGrid.BuildDisplayString ColumnExceptions, bForceSumCalc
    If m.bSummary Then
        m.strSummaryDisplay = m.TCGrid.Display
        SetIniFileProperty "SummaryAccountsDisplay", m.strSummaryDisplay, "TTSummary", g.strIniFile
    Else
        m.strDisplay = m.TCGrid.Display
        SetIniFileProperty "AccountsDisplay", m.strDisplay, "TTSummary", g.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.BuildDisplayString"
    
End Sub

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
    RaiseError "cAccountsUI.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteAccounts
'' Description: Delete the selected account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteAccount()
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account from the grid

    Set Account = SelectedAccount
    If Not Account Is Nothing Then
        g.Broker.DeleteAccount Account
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.DeleteAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddNewAccountLine
'' Description: Add the "Click here to create new account" line to the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddNewAccountLine()
On Error GoTo ErrSection:

    Dim Account As cPtAccount           ' Account object
    
    Set Account = New cPtAccount
    With Account
        .AccountID = -9&
        .AccountNumber = "Click here to create new account"
        .AccountType = -1&
        .Name = .AccountNumber
    End With
    
    AddClickHereLine Account

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.AddNewAccountLine"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddClickHereLine
'' Description: Add the "Click here ..." line to the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddClickHereLine(ByVal Account As cPtAccount)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    With m_fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        
        .RowData(.Rows - 1) = Account
        .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_Connected), .Rows - 1, GDCol(eGDCol_CurrentValue)) = Account.Name
        .Cell(flexcpFontUnderline, .Rows - 1, GDCol(eGDCol_Connected), .Rows - 1, GDCol(eGDCol_CurrentValue)) = True
        If g.nColorTheme = kDarkThemeColor Then
            .Cell(flexcpForeColor, .Rows - 1, GDCol(eGDCol_Connected), .Rows - 1, GDCol(eGDCol_CurrentValue)) = vbCyan
        Else
            .Cell(flexcpForeColor, .Rows - 1, GDCol(eGDCol_Connected), .Rows - 1, GDCol(eGDCol_CurrentValue)) = vbBlue
        End If
        .TextMatrix(.Rows - 1, AccountCol(eGDCol_AccountID)) = Str(Account.AccountID)
        If Account.AccountID = -10& Then
            .TextMatrix(.Rows - 1, AccountCol(eGDCol_AccountType)) = Str(Account.AccountType)
        End If
        SetPictureForAccount .Rows - 1
        
        .MergeRow(.Rows - 1) = True
        
        .Col = AccountCol(eGDCol_Name)
        .Sort = flexSortStringAscending
        AfterAccountSort
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.AddClickHereLine"
    
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

    Dim lIndex As Long                  ' Index into a for loop

    Set m.SnapshotAccounts = New cGdTree
    For lIndex = 1 To kNumBrokers
        Set m.SnapshotAccounts(lIndex) = Nothing
    Next lIndex
    
    Set m.Status = New cGdTree
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.Class_Initialize"
    
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

    BuildDisplayString
    
    Set m.SnapshotAccounts = Nothing
    Set m.Status = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountsUI.Class_Terminate"
    
End Sub
