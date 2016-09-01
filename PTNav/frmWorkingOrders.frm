VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmWorkingOrders 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMenu 
      Left            =   4140
      Top             =   1620
   End
   Begin VB.Timer tmrBrokers 
      Left            =   4140
      Top             =   2100
   End
   Begin VB.Timer tmrRealtime 
      Left            =   4140
      Top             =   2580
   End
   Begin VSFlex7LCtl.VSFlexGrid fgWorkingOrders 
      Height          =   2895
      Left            =   900
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _cx             =   5106
      _cy             =   5106
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Menu mnuOrders 
      Caption         =   "Orders"
      Begin VB.Menu mnuBuy 
         Caption         =   "BUY a Security"
      End
      Begin VB.Menu mnuSell 
         Caption         =   "SELL a Security"
      End
      Begin VB.Menu mnuOrderGroups 
         Caption         =   "Order Groups"
         Begin VB.Menu mnuOrderGroup 
            Caption         =   "<Manage>"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditOrder 
         Caption         =   "Edit Order"
      End
      Begin VB.Menu mnuCancelOrder 
         Caption         =   "Cancel Order"
      End
      Begin VB.Menu mnuParkOrder 
         Caption         =   "Park Order"
      End
      Begin VB.Menu mnuSubmitOrder 
         Caption         =   "Submit Order"
      End
      Begin VB.Menu mnuSubmitAll 
         Caption         =   "Submit All Parked Orders"
      End
      Begin VB.Menu mnuOrderHistory 
         Caption         =   "Order History"
      End
      Begin VB.Menu mnuNewJournal 
         Caption         =   "New Journal for Order"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuManageXOS 
         Caption         =   "Manage Exit Order Strategies"
      End
      Begin VB.Menu mnuSelectXOS 
         Caption         =   "Select Exit Order Strategy"
      End
      Begin VB.Menu mnuRemoveXOS 
         Caption         =   "Remove Exit Order Strategy"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuTradeHistory 
         Caption         =   "Trade History"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuCheckStatus 
         Caption         =   "Check Status"
      End
      Begin VB.Menu mnuViewJournals 
         Caption         =   "View Journals"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoSizeColumns 
         Caption         =   "Auto Size Columns"
      End
      Begin VB.Menu mnuDefaultColumns 
         Caption         =   "Default Columns"
      End
   End
End
Attribute VB_Name = "frmWorkingOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmWorkingOrders.cls
'' Description: Form to show a working orders grid
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/16/2010   DAJ         Fixed grid size and default startup position/size
'' 06/03/2010   DAJ         Changes for new TradeSense Order Groups
'' 06/15/2010   DAJ         Removed the TradeSense Order group menu item
'' 09/13/2010   DAJ         Show TradeSense order groups in working orders grids
'' 06/28/2011   DAJ         Setup clickable cells like hyperlinks
'' 11/28/2012   DAJ         Speed enhancements for the Trade Console
'' 01/07/2013   DAJ         Profiling for trade stuff ( for Brady and Tim )
'' 01/08/2013   DAJ         Only refresh prices if form is visible
'' 06/24/2013   DAJ         Timer Logging
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    WorkingOrdersUI As cWorkingOrdersUI ' Working orders user interface object
    adLastChanged As cGdArray           ' Array of Last Changed information by broker
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow as outside caller to print the grid information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection:

    PrintMe = frmPrintPreview.ShowMe("TNV WorkingOrders", Me, , , , , , True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmWorkingOrders.PrintMe"
    
End Function

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

    m.WorkingOrdersUI.GenerateReport vArgs
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoBrokerTimer
'' Description: Update broker information when the timer goes off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoBrokerTimer()
On Error GoTo ErrSection:

gdResetProfiles 650, 659
gdStartProfile 650
gdStartProfile 651
    Dim lIndex As Long                  ' Index into a for loop
    Dim adBrokers As cGdArray           ' Array of last changed information by broker
    Dim bUpdate As Boolean              ' Update the order?

gdStopProfile 651

    If g.bUnloading = False Then
gdStartProfile 652
        Set adBrokers = g.Broker.LastChangedForAll
gdStopProfile 652
        If Not adBrokers Is Nothing Then
            For lIndex = 1 To adBrokers.Size - 1
gdStartProfile 653
                bUpdate = (m.adLastChanged(lIndex) < adBrokers(lIndex))
gdStopProfile 653
                If m.adLastChanged(lIndex) < adBrokers(lIndex) Then
gdStartProfile 654
                    m.WorkingOrdersUI.Update lIndex
gdStopProfile 654
gdStartProfile 655
                    m.adLastChanged(lIndex) = adBrokers(lIndex)
gdStopProfile 655
                End If
            Next lIndex
        End If
        
gdStartProfile 656
        m.WorkingOrdersUI.UpdateTsOrders
gdStopProfile 656
    End If

gdStopProfile 650

If frmTTSummary.DumpProfile Then
    DebugLog "=================" & vbCrLf & gdGetProfiles(650, 659, vbCrLf) & vbCrLf & "================="
End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.DoBrokerTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisableTimers
'' Description: Disable all of the timers on the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisableTimers()
On Error GoTo ErrSection:

    tmrRealtime.Enabled = False
    tmrBrokers.Enabled = False
    tmrMenu.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.DisableTimers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPrices
'' Description: Refresh the prices in the grids with the info in the Bars
'' Inputs:      Symbol, Price, Bid, Ask
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPrices(ByVal vSymbolOrSymbolID As Variant, ByVal dPrice As Double, ByVal dBid As Double, ByVal dAsk As Double)
On Error GoTo ErrSection:

    If Visible Then
        m.WorkingOrdersUI.RefreshPrices vSymbolOrSymbolID, dPrice, dBid, dAsk
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.RefreshPrices"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPrices2
'' Description: Refresh the prices in the grids with the info in the Bars
'' Inputs:      Symbol, Price, Bid, Ask
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPrices2(ByVal vSymbolOrSymbolID As Variant, ByVal strPrice As String, ByVal strBid As String, ByVal strAsk As String)
On Error GoTo ErrSection:

    If Visible Then
        m.WorkingOrdersUI.RefreshPrices2 vSymbolOrSymbolID, strPrice, strBid, strAsk
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.RefreshPrices2"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Clear the updated colors on both grids if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearUpdatedColors()
On Error GoTo ErrSection:

    m.WorkingOrdersUI.ClearUpdatedColors

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.ClearUpdatedColors"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrid
'' Description: Filter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FilterGrid()
On Error GoTo ErrSection:

    m.WorkingOrdersUI.FilterOrdersGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.FilterGrid"
    
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
    
    m.WorkingOrdersUI.UpdateConsoleSettings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.UpdateConsoleSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the member variables when form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim UI As cWorkingOrdersControls    ' Working order controls object
    Dim strPlacement As String          ' Placement string from the ini file
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("frmWorkingOrders", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        Move 225, 2265, 15720, 3600
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    Caption = "Open Orders (right-click on grid to see options)"
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)

    Set UI = New cWorkingOrdersControls
    With UI
        Set .frm = Me
        Set .fgGrid = fgWorkingOrders
        Set .tmrRealtime = tmrRealtime
        Set .tmrMenu = tmrMenu
        Set .mnuOrders = mnuOrders
        Set .mnuBuy = mnuBuy
        Set .mnuSell = mnuSell
        Set .mnuOrderGroups = mnuOrderGroups
        Set .mnuOrderGroup = mnuOrderGroup
        Set .mnuEditOrder = mnuEditOrder
        Set .mnuCancelOrder = mnuCancelOrder
        Set .mnuParkOrder = mnuParkOrder
        Set .mnuSubmitOrder = mnuSubmitOrder
        Set .mnuSubmitAll = mnuSubmitAll
        Set .mnuOrderHistory = mnuOrderHistory
        Set .mnuNewJournal = mnuNewJournal
        Set .mnuManageXOS = mnuManageXOS
        Set .mnuSelectXOS = mnuSelectXOS
        Set .mnuRemoveXOS = mnuRemoveXOS
        Set .mnuPrint = mnuPrint
        Set .mnuTradeHistory = mnuTradeHistory
        Set .mnuSettings = mnuSettings
        Set .mnuCheckStatus = mnuCheckStatus
        Set .mnuViewJournals = mnuViewJournals
        Set .mnuAutoSizeColumns = mnuAutoSizeColumns
        Set .mnuDefaultColumns = mnuDefaultColumns
    End With

    Set m.WorkingOrdersUI = New cWorkingOrdersUI
    m.WorkingOrdersUI.Init "Working Orders", UI, False
    
    Set m.adLastChanged = New cGdArray
    m.adLastChanged.Create eGDARRAY_Doubles, kNumBrokers
    
    tmrBrokers.Interval = 1000
    tmrBrokers.Enabled = True
    
    tmrMenu.Interval = 100
    tmrMenu.Enabled = False
    
    mnuOrders.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_MouseMove
'' Description: If the mouse cursor has been set somewhere else, reset it
'' Inputs:      Button pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Me.MousePointer = vbCustom Then
        Me.MousePointer = vbDefault
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    With fgWorkingOrders
        .Move 0, 0, ScaleWidth, ScaleHeight
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, re-attach the grid
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = vbFormControlMenu Then
        If Not g.ConsoleForms Is Nothing Then
            g.ConsoleForms.ShowForm(eGDConsoleForm_OpenOrders) = False
        End If
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up member variables when form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    DisableTimers

    SetIniFileProperty "frmWorkingOrders", GetFormPlacement(Me), "Placement", g.strIniFile
    
    Set m.WorkingOrdersUI = Nothing
    Set m.adLastChanged = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.Form_Unload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrderGroup_Click
'' Description: The user has chosen an order group
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrderGroup_Click(Index As Integer)
On Error GoTo ErrSection:

    m.WorkingOrdersUI.SelectOrderGroup Index

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.mnuOrderGroup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrBrokers_Timer
'' Description: Update broker information when the timer goes off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrBrokers_Timer()
On Error GoTo ErrSection:

    TimerStart "frmWorkingOrders.tmrBrokers"
    DoBrokerTimer
    TimerEnd "frmWorkingOrders.tmrBrokers", tmrBrokers.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.tmrBrokers_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrRealtime_Timer
'' Description: Update data when the timer goes off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrRealTime_Timer()
On Error GoTo ErrSection:

    TimerStart "frmWorkingOrders.tmrRealTime"
    TimerEnd "frmWorkingOrders.tmrRealTime", tmrRealtime.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmWorkingOrders.tmrRealtime_Timer"
    
End Sub

