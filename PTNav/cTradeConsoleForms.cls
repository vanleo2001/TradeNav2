VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeConsoleForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeConsoleForms.cls
'' Description: Class to control all of the trade console forms
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/17/2010   DAJ         Added number visible items for the toolbar captions
'' 03/17/2010   DAJ         Toggle show/hide of auto trade item column in orders, positions
'' 04/08/2010   DAJ         Changed the default summary height when undocked
'' 06/15/2010   DAJ         Added TradeSense orders as new Trade Console form
'' 11/23/2015   DAJ         Added Refresh function
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kDefaultFullHeight = 1950
Private Const kDefaultUndockedFull = 6000
Private Const kDockedToolbarHeight = 705 '675
Private Const kUndockedToolbarHeight = 795
Private Const kMinHeight = 800

Public Enum eGDConsoleForms
    eGDConsoleForm_Summary = 0
    eGDConsoleForm_OpenOrders
    eGDConsoleForm_Positions
    eGDConsoleForm_Accounts
    eGDConsoleForm_AutoTrading
    eGDConsoleForm_ActivityLog
    eGDConsoleForm_TodaysFills
    eGDConsoleForm_TradeSenseOrders
    eGDConsoleForm_NumForms
End Enum

Private Type mPrivate
    aFormShown As cGdArray              ' Array of flags to show form
    alNumVisible As cGdArray            ' Array of number of visible items
End Type
Private m As mPrivate

Private Function ConsoleForm(ByVal nConsoleForm As eGDConsoleForms) As Long
    ConsoleForm = nConsoleForm
End Function

Public Property Get ShowForm(ByVal nForm As eGDConsoleForms) As Boolean
    ShowForm = (m.aFormShown(nForm) <> 0)
End Property
Public Property Let ShowForm(ByVal nForm As eGDConsoleForms, ByVal bShow As Boolean)
    m.aFormShown(nForm) = Abs(bShow)
    SetIniFileProperty "Show" & Str(nForm), m.aFormShown(nForm), "ConsoleForms", g.strIniFile
    ShowTheForm nForm
End Property

Public Property Get NumVisible(ByVal nForm As eGDConsoleForms) As Long
    NumVisible = m.alNumVisible(nForm)
End Property
Public Property Let NumVisible(ByVal nForm As eGDConsoleForms, ByVal lNumVisible As Long)
On Error GoTo ErrSection:

    m.alNumVisible(nForm) = lNumVisible
    
    If FormIsLoaded("frmTTSummary") Then
        frmTTSummary.UpdateToolbarCaption nForm, lNumVisible
    End If

    Select Case nForm
        Case eGDConsoleForm_OpenOrders
            NumVisible(eGDConsoleForm_Summary) = NumVisible(eGDConsoleForm_Positions) + lNumVisible
        Case eGDConsoleForm_Positions
            NumVisible(eGDConsoleForm_Summary) = NumVisible(eGDConsoleForm_OpenOrders) + lNumVisible
        Case eGDConsoleForm_AutoTrading
            If FormIsLoaded("frmTTSummary") Then
                FilterGrids False
            End If
    End Select
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeConsoleForms.NumVisible.Let"
    
End Property

Public Property Get SummaryHeight() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the property
    Static nDockedToolbarHeight&

    If frmMain.DockPro.State("frmTTSummary") = DPDocked Then
        If ShowForm(eGDConsoleForm_Summary) = True Then
            lReturn = GetIniFileProperty("SummaryHeightDocked", kDefaultFullHeight, "ConsoleForms", g.strIniFile)
            If lReturn < kMinHeight Then
                lReturn = kDefaultFullHeight
            End If
        Else
            lReturn = kDockedToolbarHeight
        
            ' TLB 12/18/2015: when docked, the "ScaleHeight" method isn't working consistently (e.g. even between the IDE
            ' and compiled on the same machine!), so we're now trying a different approach ...
            ' Height = the window title height (by calling GetSystemMetrics) + sheridan toolbar height + a few pixels
            lReturn = GetSystemMetrics(4) + frmTTSummary.tbToolbar.GetDockHeight(ssDockedTop)
            If IsAtLeastVista Then
                lReturn = lReturn + 3 ' for some reason, don't want the extra pixels for XP
            End If
            lReturn = lReturn * Screen.TwipsPerPixelY
            If lReturn > nDockedToolbarHeight Then
                ' but for some reason,  must also ONLY change it when it gets bigger (never smaller),
                ' otherwise the initial behavior gets all screwed up after first startup (until you toggle a couple times)
                DebugLog "TTSummaryDockedToolbarHeight = " & Str(nDockedToolbarHeight)
                nDockedToolbarHeight = lReturn
            End If
            lReturn = nDockedToolbarHeight
        End If
    ElseIf frmMain.DockPro.State("frmTTSummary") = DPUndocked Then
        If ShowForm(eGDConsoleForm_Summary) = True Then
            lReturn = GetIniFileProperty("SummaryHeightUndocked", kDefaultUndockedFull, "ConsoleForms", g.strIniFile)
            If lReturn < kMinHeight Then
                lReturn = kDefaultUndockedFull
            End If
        Else
            'lReturn = kUndockedToolbarHeight
            lReturn = frmTTSummary.Height - frmTTSummary.ScaleHeight - 1 * Screen.TwipsPerPixelY  '15
        End If
    End If
    
    SummaryHeight = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeConsoleForms.SummaryHeight.Get"
    
End Property
Public Property Let SummaryHeight(ByVal lSummaryHeight As Long)
On Error GoTo ErrSection:

    If ShowForm(eGDConsoleForm_Summary) Then
        If frmMain.DockPro.State("frmTTSummary") = DPDocked Then
            SetIniFileProperty "SummaryHeightDocked", lSummaryHeight, "ConsoleForms", g.strIniFile
        ElseIf frmMain.DockPro.State("frmTTSummary") = DPUndocked Then
            SetIniFileProperty "SummaryHeightUndocked", lSummaryHeight, "ConsoleForms", g.strIniFile
        End If
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeConsoleForms.SummaryHeight.Let"

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowForms
'' Description: Show/Hide the forms as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowForms()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    For lIndex = 0 To ConsoleForm(eGDConsoleForm_NumForms) - 1
        ShowTheForm lIndex
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleForms.ShowForms"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SummaryChangedState
'' Description: The Summary form has just changed it's dockable state
'' Inputs:      Previous State
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SummaryChangedState(ByVal nPreviousState As DockState)
On Error GoTo ErrSection:

    Dim bShowForm As Boolean            ' Is the form shown?
    Dim bMoreThanOne As Boolean         ' Is there more than one docked?

    bMoreThanOne = False
    bShowForm = ShowForm(eGDConsoleForm_Summary)
    
    If frmMain.DockPro.State("frmTTSummary") = DPDocked Then
        If frmMain.DockPro.AlignWhenDocked("frmTTSummary") = HAlignTop Then
            If frmMain.DockPro.DockedCount(HAlignTop) > 1 Then
                bMoreThanOne = True
            End If
        ElseIf frmMain.DockPro.AlignWhenDocked("frmTTSummary") = HAlignBottom Then
            If frmMain.DockPro.DockedCount(HAlignBottom) > 1 Then
                bMoreThanOne = True
            End If
        End If
    End If
    
    If bMoreThanOne Then
        ShowForm(eGDConsoleForm_Summary) = True
    Else
        ShowForm(eGDConsoleForm_Summary) = bShowForm
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTradeConsoleForms.SummaryChangedState"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrids
'' Description: Re-Filter all of the grids
'' Inputs:      Do TradeItems?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FilterGrids(Optional ByVal bDoTradeItems As Boolean = True)
On Error GoTo ErrSection:

    If FormIsLoaded("frmTTSummary") Then
        FixFormControls frmTTSummary, ALT_GRID_ROW_COLOR
        frmTTSummary.FilterGrids
    End If
    If FormIsLoaded("frmWorkingOrders") Then
        FixFormControls frmWorkingOrders, ALT_GRID_ROW_COLOR
        frmWorkingOrders.FilterGrid
    End If
    If FormIsLoaded("frmOpenPositions") Then
        FixFormControls frmOpenPositions, ALT_GRID_ROW_COLOR
        frmOpenPositions.FilterGrid
    End If
    If FormIsLoaded("frmAccounts") Then
        FixFormControls frmAccounts, ALT_GRID_ROW_COLOR
        frmAccounts.FilterGrid
    End If
    If FormIsLoaded("frmTradeItems") And bDoTradeItems Then
        FixFormControls frmTradeItems, ALT_GRID_ROW_COLOR
        frmTradeItems.FilterGrid
    End If
    If FormIsLoaded("frmActiveTsOrderGroups") Then
        FixFormControls frmActiveTsOrderGroups, ALT_GRID_ROW_COLOR
        frmActiveTsOrderGroups.FilterGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleForms.FilterGrids"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Tell the Trade Console to update all of the buttons
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Refresh()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If FormIsLoaded("frmTTSummary") Then
        For lIndex = 0 To ConsoleForm(eGDConsoleForm_NumForms) - 1
            frmTTSummary.UpdateToolbarCaption lIndex, NumVisible(lIndex)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleForms.Refresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowTheForm
'' Description: Show/Hide the given form as appropriate
'' Inputs:      Form
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowTheForm(ByVal nForm As eGDConsoleForms)
On Error GoTo ErrSection:

    Dim bShow As Boolean                ' Should the form be shown?
    Dim frm As Form                     ' The form to operate on
    Dim lHeight As Long                 ' Height of the from

    bShow = ShowForm(nForm)
    If nForm = eGDConsoleForm_Summary Then
        With frmMain.DockPro
            If .State("frmTTSummary") = DPDocked Then
                If .AlignWhenDocked("frmTTSummary") = HAlignTop Then
                    .TopEdgeHeight = SummaryHeight
                ElseIf .AlignWhenDocked("frmTTSummary") = HAlignBottom Then
                    .BottomEdgeHeight = SummaryHeight
                End If
            ElseIf .State("frmTTSummary") = DPUndocked Then
                frmTTSummary.Height = SummaryHeight
            End If
        End With
    Else
        Set frm = TheForm(nForm)
        If Not frm Is Nothing Then
            If bShow And ConsoleVisible Then
                If frm.Visible = False Then
                    mGenesis.ShowForm frm, eForm_Nonmodal, frmMain, , ALT_GRID_ROW_COLOR
                End If
            Else
                frm.Visible = False
            End If
        End If
    End If
        
    If FormIsLoaded("frmTTSummary") Then
        frmTTSummary.FormShown nForm, bShow
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleForms.ShowTheForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TheForm
'' Description: Get the actual form that was requested
'' Inputs:      Form to Retrieve
'' Returns:     Form (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TheForm(ByVal nForm As eGDConsoleForms) As Form
On Error GoTo ErrSection:

    Dim frmReturn As Form               ' Return value for the function
    
    Set frmReturn = Nothing
    Select Case nForm
        Case eGDConsoleForm_OpenOrders
            If FormIsLoaded("frmWorkingOrders") Then
                Set frmReturn = frmWorkingOrders
            End If
        
        Case eGDConsoleForm_Positions
            If FormIsLoaded("frmOpenPositions") Then
                Set frmReturn = frmOpenPositions
            End If
        
        Case eGDConsoleForm_Accounts
            If FormIsLoaded("frmAccounts") Then
                Set frmReturn = frmAccounts
            End If
        
        Case eGDConsoleForm_AutoTrading
            If FormIsLoaded("frmTradeItems") Then
                Set frmReturn = frmTradeItems
            End If
        
        Case eGDConsoleForm_ActivityLog
            If FormIsLoaded("frmActivityLog") Then
                Set frmReturn = frmActivityLog
            End If
        
        Case eGDConsoleForm_TodaysFills
            If FormIsLoaded("frmTodaysFills") Then
                Set frmReturn = frmTodaysFills
            End If
            
        Case eGDConsoleForm_TradeSenseOrders
            If FormIsLoaded("frmActiveTsOrderGroups") Then
                Set frmReturn = frmActiveTsOrderGroups
            End If
        
    End Select
    
    Set TheForm = frmReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeConsoleForms.TheForm"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConsoleVisible
'' Description: Is the Trade Console form visible?
'' Inputs:      None
'' Returns:     True if visible, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ConsoleVisible() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If FormIsLoaded("frmTTSummary") Then
        bReturn = frmTTSummary.Visible
    Else
        bReturn = False
    End If
    
    ConsoleVisible = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeConsoleForms.ConsoleVisible"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    ' Create the form shown array...
    Set m.aFormShown = New cGdArray
    m.aFormShown.Create eGDARRAY_TinyInts, ConsoleForm(eGDConsoleForm_NumForms)
    
    ' Create the number visible array...
    Set m.alNumVisible = New cGdArray
    m.alNumVisible.Create eGDARRAY_Longs, ConsoleForm(eGDConsoleForm_NumForms)
    
    ' Populate the form shown array from the INI file -- default to not shown...
    For lIndex = 0 To ConsoleForm(eGDConsoleForm_NumForms) - 1
        m.aFormShown(lIndex) = GetIniFileProperty("Show" & Str(lIndex), 0, "ConsoleForms", g.strIniFile)
        m.alNumVisible(lIndex) = 0&
    Next lIndex
    
    ' Load the forms...
    Load frmWorkingOrders
    Load frmOpenPositions
    Load frmAccounts
    Load frmTradeItems
    Load frmActivityLog
    Load frmTodaysFills
    Load frmActiveTsOrderGroups

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleForms.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    ' Destroy the arrays...
    Set m.aFormShown = Nothing
    Set m.alNumVisible = Nothing
    
    ' Unload the forms...
    Unload frmWorkingOrders
    Unload frmOpenPositions
    Unload frmAccounts
    Unload frmTradeItems
    Unload frmActivityLog
    Unload frmTodaysFills
    Unload frmActiveTsOrderGroups
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleForms.Class_Terminate"
    
End Sub
