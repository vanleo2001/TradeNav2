VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTradeItems 
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
   Begin VB.Timer tmrRealtime 
      Left            =   4080
      Top             =   2520
   End
   Begin VB.Timer tmrMenu 
      Left            =   4080
      Top             =   2040
   End
   Begin VSFlex7LCtl.VSFlexGrid fgTradeItems 
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
   Begin VB.Menu mnuTradeItems 
      Caption         =   "Trade Items"
      Begin VB.Menu mnuDisableAll 
         Caption         =   "Disable All"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlatten 
         Caption         =   "Flatten Position"
      End
      Begin VB.Menu mnuEnterPosition 
         Caption         =   "Enter Position"
      End
      Begin VB.Menu mnuChangePosition 
         Caption         =   "Change Position"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewTradingItem 
         Caption         =   "New Trading Item"
      End
      Begin VB.Menu mnuEditTradingItem 
         Caption         =   "Edit Trading Item"
      End
      Begin VB.Menu mnuDeleteTradingItem 
         Caption         =   "Delete Trading Item"
      End
      Begin VB.Menu mnuRollContracts 
         Caption         =   "Roll to another Contract"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStrategy 
         Caption         =   "Edit Strategy"
      End
      Begin VB.Menu mnuStrategyPerformance 
         Caption         =   "Strategy Performance"
      End
      Begin VB.Menu mnuActualPerformance 
         Caption         =   "Actual Performance"
      End
      Begin VB.Menu mnuNextBarReport 
         Caption         =   "Next Bar Report"
      End
      Begin VB.Menu mnuShowChart 
         Caption         =   "Show Chart"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuAccountHistory 
         Caption         =   "View Account History"
      End
      Begin VB.Menu mnuAlerts 
         Caption         =   "Alerts"
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
Attribute VB_Name = "frmTradeItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTradeItems.cls
'' Description: Form to show an automated trading items
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
'' 06/28/2011   DAJ         Setup clickable cells like hyperlinks
'' 07/25/2011   DAJ         Allow for delete of auto trade item from outside
'' 08/02/2011   DAJ         Added the change position menu item
'' 02/20/2013   DAJ         Added "Actual Performance" menu item
'' 04/03/2013   DAJ         Automated Strategy Baskets
'' 06/11/2014   DAJ         Dump the automated trading items grid to a file if it changes
'' 01/23/2015   DAJ         Disable All menu item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    TradeItemUI As cTradeItemUI         ' Trading Items user interface object
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

    PrintMe = frmPrintPreview.ShowMe("TNV TradeItems", Me, , , , , , True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeItems.PrintMe"
    
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
    
    m.TradeItemUI.GenerateReport vArgs
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.GenerateReport"

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
    tmrMenu.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.DisableTimers"
    
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

    m.TradeItemUI.FilterTradeItemsGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.FilterGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTradeItem
'' Description: Refresh the trade item
'' Inputs:      Trade Item ID, Deleted?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshTradeItem(ByVal lTradeItemID As Long, Optional ByVal bDeleted As Boolean = False)
On Error GoTo ErrSection:

    m.TradeItemUI.RefreshTradeItem lTradeItemID, bDeleted

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.RefreshTradeItem"
    
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
    
    m.TradeItemUI.UpdateConsoleSettings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.UpdateConsoleSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketChanged
'' Description: Notification that the given strategy basket has changed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BasketChanged(ByVal Basket As cStrategyBasket)
On Error GoTo ErrSection:

    m.TradeItemUI.BasketChanged Basket

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.BasketChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpGridIfDifferent
'' Description: Dump the contents of the grid to a file if it has changed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DumpGridIfDifferent()
On Error GoTo ErrSection:

    m.TradeItemUI.DumpGridIfDifferent

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.DumpGridIfDifferent"
    
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

    Dim UI As cTradeItemContols         ' Trading item controls
    Dim strPlacement As String          ' Placement string from the ini file
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("frmTradeItems", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        Move 1185, 3630, 15720, 3600
    Else
        SetFormPlacement Me, strPlacement
    End If
        
    Caption = "Automated Trading Items (right-click on grid to see options)"
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    
    Set UI = New cTradeItemContols
    With UI
        Set .frm = Me
        
        Set .fgGrid = fgTradeItems
        Set .tmrMenu = tmrMenu
        Set .tmrRealtime = tmrRealtime
        
        Set .mnuTradeItems = mnuTradeItems
        Set .mnuDisableAll = mnuDisableAll
        Set .mnuFlatten = mnuFlatten
        Set .mnuEnterPosition = mnuEnterPosition
        Set .mnuChangePosition = mnuChangePosition
        Set .mnuNewTradeItem = mnuNewTradingItem
        Set .mnuEditTradeItem = mnuEditTradingItem
        Set .mnuDeleteTradeItem = mnuDeleteTradingItem
        Set .mnuRollContract = mnuRollContracts
        Set .mnuEditStrategy = mnuEditStrategy
        Set .mnuStrategyPerformance = mnuStrategyPerformance
        Set .mnuActualPerformance = mnuActualPerformance
        Set .mnuNextBarReport = mnuNextBarReport
        Set .mnuShowChart = mnuShowChart
        Set .mnuPrint = mnuPrint
        Set .mnuTradeHistory = mnuAccountHistory
        Set .mnuAlerts = mnuAlerts
        Set .mnuAutoSizeColumns = mnuAutoSizeColumns
        Set .mnuDefaultColumns = mnuDefaultColumns
    End With
    
    Set m.TradeItemUI = New cTradeItemUI
    m.TradeItemUI.Init "Trade Items", UI
    
    tmrMenu.Interval = 100
    tmrMenu.Enabled = False
    
    mnuTradeItems.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.Form_Load"
    
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

    With fgTradeItems
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
            g.ConsoleForms.ShowForm(eGDConsoleForm_AutoTrading) = False
        End If
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.Form_QueryUnload"
    
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
    
    SetIniFileProperty "frmTradeItems", GetFormPlacement(Me), "Placement", g.strIniFile
    
    Set m.TradeItemUI = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeItems.Form_Unload"

End Sub

