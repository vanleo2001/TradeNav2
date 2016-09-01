VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTodaysFills 
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
      Top             =   1440
   End
   Begin VB.Timer tmrBrokers 
      Enabled         =   0   'False
      Left            =   4140
      Top             =   2520
   End
   Begin VB.Timer tmrRealtime 
      Enabled         =   0   'False
      Left            =   4140
      Top             =   1980
   End
   Begin VSFlex7LCtl.VSFlexGrid fgTodaysFills 
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
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
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
      Begin VB.Menu mnuSep1 
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
Attribute VB_Name = "frmTodaysFills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTodaysFills.cls
'' Description: Form that holds a grid with todays fills
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
'' 06/24/2013   DAJ         Timer Logging
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    TodaysFills As cTodaysFillsUI       ' Object to handle today's fills
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

    PrintMe = frmPrintPreview.ShowMe("TNV TodaysFills", Me, , , , , , True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTodaysFills.PrintMe"
    
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
    
    m.TodaysFills.GenerateReport vArgs

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTodaysFills.GenerateReport"

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

    Dim lIndex As Long                  ' Index into a for loop
    Dim adBrokers As cGdArray           ' Array of last changed information by broker

    If g.bUnloading = False Then
        Set adBrokers = g.Broker.LastChangedForAll
        If Not adBrokers Is Nothing Then
            For lIndex = 1 To adBrokers.Size - 1
                If m.adLastChanged(lIndex) < adBrokers(lIndex) Then
                    m.TodaysFills.Update lIndex
                    m.adLastChanged(lIndex) = adBrokers(lIndex)
                End If
            Next lIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTodaysFills.DoBrokerTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisableTimers
'' Description: Disable the timers on the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisableTimers()
On Error GoTo ErrSection:

    tmrRealtime.Enabled = False
    tmrBrokers.Enabled = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTodaysFills.DisableTimers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize members when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement string from the ini file
    Dim UI As cTodaysFillsControls      ' Controls for the todays fills object
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("frmTodaysFills", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        Move 1800, 4560, 15720, 3600
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    Caption = "Today's Fills (right-click on grid to see options)"
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    
    Set UI = New cTodaysFillsControls
    With UI
        Set .frm = Me
        Set .fgGrid = fgTodaysFills
        Set .tmrRealtime = tmrRealtime
        Set .tmrMenu = tmrMenu
        
        Set .mnuTodaysFills = mnuPopUp
        Set .mnuPrint = mnuPrint
        Set .mnuTradeHistory = mnuTradeHistory
        Set .mnuSettings = mnuSettings
        Set .mnuCheckStatus = mnuCheckStatus
        Set .mnuViewJournals = mnuViewJournals
        Set .mnuAutoSizeColumns = mnuAutoSizeColumns
        Set .mnuDefaultColumns = mnuDefaultColumns
    End With

    Set m.TodaysFills = New cTodaysFillsUI
    m.TodaysFills.Init "Todays Fills", UI
    
    Set m.adLastChanged = New cGdArray
    m.adLastChanged.Create eGDARRAY_Doubles

    tmrBrokers.Interval = 1000
    tmrBrokers.Enabled = True
    
    tmrMenu.Interval = 100
    tmrMenu.Enabled = False
    
    mnuPopUp.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTodaysFills.Form_Load"

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

    With fgTodaysFills
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
            g.ConsoleForms.ShowForm(eGDConsoleForm_TodaysFills) = False
        End If
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTodaysFills.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up members when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    DisableTimers

    SetIniFileProperty "frmTodaysFills", GetFormPlacement(Me), "Placement", g.strIniFile
    
    Set m.TodaysFills = Nothing
    Set m.adLastChanged = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTodaysFills.Form_Unload"
    
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

    TimerStart "frmTodaysFills.tmrBrokers_Timer"
    DoBrokerTimer
    TimerEnd "frmTodaysFills.tmrBrokers_Timer", tmrBrokers.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTodaysFills.tmrBrokers_Timer"
    
End Sub

