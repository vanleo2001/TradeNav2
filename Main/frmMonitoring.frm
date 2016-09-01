VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmMonitoring 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Left            =   120
      Top             =   240
   End
   Begin VSFlex7LCtl.VSFlexGrid fgMonitor 
      Height          =   2895
      Left            =   900
      TabIndex        =   0
      Top             =   180
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
End
Attribute VB_Name = "frmMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmMonitoring.frm
'' Description: Perform automatic downloads to make sure processes are up and
''              running
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strMonitorPath As String            ' Path for the monitoring files

    dNextQuoteBoardRefresh As Double    ' Time of the next quote board refresh
    dNextOptionChainRefresh As Double   ' Time of the next option chain
    dNextDailyDownload As Double        ' Time of the next daily download
    dNextCurrentSessionUpdate As Double ' Time of the next current session update
    
    astrHolidays As cGdArray            ' List of domestic holidays
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_On = 0
    eGDCol_Name
    eGDCol_Interval
    eGDCol_Start
    eGDCol_End
    eGDCol_Status
    eGDCol_NumCols
End Enum

Private Enum eGDRows
    eGDRow_QuoteBoard = 1
    eGDRow_CurrentSession
    eGDRow_DailyDownload
    eGDRow_OptionChain
    eGDRow_NumRows
End Enum

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Function GDRow(ByVal Row As eGDRows) As Long
    GDRow = Row
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe()
On Error GoTo ErrSection:

    ' Set up the path...
    m.strMonitorPath = AddSlash(GetIniFileProperty("Path", "K:\Monitor\", "Monitoring", g.strIniFile))
    
    ' Put out the Umbrella.RUN file...
    FileFromString m.strMonitorPath & "Umbrella.RUN", CStr(CDbl(Now))

    ' Load the Holidays table...
    Set m.astrHolidays = New cGdArray
    m.astrHolidays.Create eGDARRAY_Strings
    m.astrHolidays.FromFile m.strMonitorPath & "Holiday.DAT"
    If m.astrHolidays.Size > 0 Then m.astrHolidays.Sort

    ' Initialize and Load the grid...
    fgMonitor.Redraw = flexRDNone
    InitGrid
    LoadGrid
    fgMonitor.Redraw = flexRDBuffered
    
    ' Calculate the Next Download times for each type...
    CalcNext eGDRow_QuoteBoard
    CalcNext eGDRow_OptionChain
    CalcNext eGDRow_CurrentSession
    CalcNext eGDRow_DailyDownload
    
    ' Set up the timer...
    tmr.interval = 60000
    tmr.Enabled = True

    ' Show the form...
    ShowForm Me, eForm_Nonmodal, frmMain

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.ShowMe", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgMonitor_AfterEdit
'' Description: Recalculate next refresh based on the new information
'' Inputs:      Row and Column of the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgMonitor_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    CalcNext Row

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.fgMonitor.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgMonitor_AfterRowColChange
'' Description: Go into edit mode on the new cell if appropriate
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgMonitor_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If NewCol <> GDCol(eGDCol_On) Then fgMonitor.EditCell

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.fgMonitor.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgMonitor_BeforeEdit
'' Description: Only allow the user to edit certain cells
'' Inputs:      Row and Column of Edit, Whether to Cancelt the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgMonitor_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col = GDCol(eGDCol_Name) Or Col = GDCol(eGDCol_Status) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.fgMonitor.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show the help if the user presses F1
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText As String               ' Text from an ini file property

    Caption = "Trade Navigator Monitoring"
    Icon = Picture16(ToolbarIcon("kGreenLight"))
    
    CenterTheForm Me
    strText = GetIniFileProperty("Monitoring", "", "Placement", g.strIniFile)
    If strText <> "" Then SetFormPlacement Me, strText, "LHTW"
    
    tmr.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Only let the form unload under certain conditions
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = vbFormCode Or UnloadMode = vbFormControlMenu Then
        Cancel = True
        Beep
    Else
        tmr.Enabled = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.Form.QueryUnload", eGDRaiseError_Show

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, 5400, 2450) Then Exit Sub
    
    With fgMonitor
        .Move 0, 0, ScaleWidth, ScaleHeight
    End With

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

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgMonitor
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' General settings
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_On)) = "On"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Activity"
        .TextMatrix(0, GDCol(eGDCol_Interval)) = "Interval"
        .TextMatrix(0, GDCol(eGDCol_Start)) = "Start"
        .TextMatrix(0, GDCol(eGDCol_End)) = "End"
        .TextMatrix(0, GDCol(eGDCol_Status)) = "Status"
        
        .ColDataType(GDCol(eGDCol_On)) = flexDTBoolean
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.InitGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim strItem As String               ' Item out of the INI file
    
    With fgMonitor
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = GDRow(eGDRow_NumRows)
        
        strItem = GetIniFileProperty("QuoteBoard", "True;10;0;1439", "Monitoring", g.strIniFile)
        CheckedCell(fgMonitor, GDRow(eGDRow_QuoteBoard), GDCol(eGDCol_On)) = CBool(Parse(strItem, ";", 1))
        .TextMatrix(GDRow(eGDRow_QuoteBoard), GDCol(eGDCol_Name)) = "Quote Board Refresh"
        .TextMatrix(GDRow(eGDRow_QuoteBoard), GDCol(eGDCol_Interval)) = Parse(strItem, ";", 2)
        .TextMatrix(GDRow(eGDRow_QuoteBoard), GDCol(eGDCol_Start)) = Format(CLng(Parse(strItem, ";", 3)) / 1440, "HH:MM AM/PM")
        .TextMatrix(GDRow(eGDRow_QuoteBoard), GDCol(eGDCol_End)) = Format(CLng(Parse(strItem, ";", 4)) / 1440, "HH:MM AM/PM")
        
        strItem = GetIniFileProperty("OptionChain", "True;10;2;1439", "Monitoring", g.strIniFile)
        CheckedCell(fgMonitor, GDRow(eGDRow_OptionChain), GDCol(eGDCol_On)) = CBool(Parse(strItem, ";", 1))
        .TextMatrix(GDRow(eGDRow_OptionChain), GDCol(eGDCol_Name)) = "Option Chain"
        .TextMatrix(GDRow(eGDRow_OptionChain), GDCol(eGDCol_Interval)) = CStr(Parse(strItem, ";", 2))
        .TextMatrix(GDRow(eGDRow_OptionChain), GDCol(eGDCol_Start)) = Format(CLng(Parse(strItem, ";", 3)) / 1440, "HH:MM AM/PM")
        .TextMatrix(GDRow(eGDRow_OptionChain), GDCol(eGDCol_End)) = Format(CLng(Parse(strItem, ";", 4)) / 1440, "HH:MM AM/PM")
        
        strItem = GetIniFileProperty("CSU", "True;10;335;935", "Monitoring", g.strIniFile)
        CheckedCell(fgMonitor, GDRow(eGDRow_CurrentSession), GDCol(eGDCol_On)) = CBool(Parse(strItem, ";", 1))
        .TextMatrix(GDRow(eGDRow_CurrentSession), GDCol(eGDCol_Name)) = "Current Session Update"
        .TextMatrix(GDRow(eGDRow_CurrentSession), GDCol(eGDCol_Interval)) = CStr(Parse(strItem, ";", 2))
        .TextMatrix(GDRow(eGDRow_CurrentSession), GDCol(eGDCol_Start)) = Format(CLng(Parse(strItem, ";", 3)) / 1440, "HH:MM AM/PM")
        .TextMatrix(GDRow(eGDRow_CurrentSession), GDCol(eGDCol_End)) = Format(CLng(Parse(strItem, ";", 4)) / 1440, "HH:MM AM/PM")
        
        strItem = GetIniFileProperty("DailyDownload", "True;10;7;1439", "Monitoring", g.strIniFile)
        CheckedCell(fgMonitor, GDRow(eGDRow_DailyDownload), GDCol(eGDCol_On)) = CBool(Parse(strItem, ";", 1))
        .TextMatrix(GDRow(eGDRow_DailyDownload), GDCol(eGDCol_Name)) = "Daily Download"
        .TextMatrix(GDRow(eGDRow_DailyDownload), GDCol(eGDCol_Interval)) = CStr(Parse(strItem, ";", 2))
        .TextMatrix(GDRow(eGDRow_DailyDownload), GDCol(eGDCol_Start)) = Format(CLng(Parse(strItem, ";", 3)) / 1440, "HH:MM AM/PM")
        .TextMatrix(GDRow(eGDRow_DailyDownload), GDCol(eGDCol_End)) = Format(CLng(Parse(strItem, ";", 4)) / 1440, "HH:MM AM/PM")
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.LoadGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save information and cleanup when the form is unloaded
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim strItem As String               ' Item to store to the INI file

    tmr.Enabled = False

    SetIniFileProperty "Monitoring", GetFormPlacement(Me), "Placement", g.strIniFile
    
    strItem = CStr(CheckedCell(fgMonitor, 1, GDCol(eGDCol_On))) & ";" & CStr(fgMonitor.TextMatrix(1, GDCol(eGDCol_Interval))) & ";" & CStr(TimeFromString(1, GDCol(eGDCol_Start))) & ";" & CStr(TimeFromString(1, GDCol(eGDCol_End)))
    SetIniFileProperty "QuoteBoard", strItem, "Monitoring", g.strIniFile
    strItem = CStr(CheckedCell(fgMonitor, 2, GDCol(eGDCol_On))) & ";" & CStr(fgMonitor.TextMatrix(2, GDCol(eGDCol_Interval))) & ";" & CStr(TimeFromString(2, GDCol(eGDCol_Start))) & ";" & CStr(TimeFromString(2, GDCol(eGDCol_End)))
    SetIniFileProperty "CSU", strItem, "Monitoring", g.strIniFile
    strItem = CStr(CheckedCell(fgMonitor, 3, GDCol(eGDCol_On))) & ";" & CStr(fgMonitor.TextMatrix(3, GDCol(eGDCol_Interval))) & ";" & CStr(TimeFromString(3, GDCol(eGDCol_Start))) & ";" & CStr(TimeFromString(3, GDCol(eGDCol_End)))
    SetIniFileProperty "DailyDownload", strItem, "Monitoring", g.strIniFile
    strItem = CStr(CheckedCell(fgMonitor, 4, GDCol(eGDCol_On))) & ";" & CStr(fgMonitor.TextMatrix(4, GDCol(eGDCol_Interval))) & ";" & CStr(TimeFromString(4, GDCol(eGDCol_Start))) & ";" & CStr(TimeFromString(4, GDCol(eGDCol_End)))
    SetIniFileProperty "OptionChain", strItem, "Monitoring", g.strIniFile
    
    ' Delete the Umbrella.RUN file...
    KillFile m.strMonitorPath & "Umbrella.RUN"
    
ErrExit:
    Set m.astrHolidays = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmr_Timer
'' Description: Perform necessary actions at necessary times
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmr_Timer()
On Error GoTo ErrSection:

    Dim dNow As Double                  ' Current date and time
    Dim strKey As String                ' Key in the registry
    Static bInProgress As Boolean       ' Are we currently working on something?
    
    ' Get the current date and time...
    dNow = Now
    
    ' Reset the registry for a "random" Tweedle...
    strKey = "Software\Genesis Financial Data Services\GClient"
    SetRegistryValue rkLocalMachine, strKey, "IPLastTime", 0&
    
    ' See if we need to do a Current Session Update now...
    ' (Important to have this before the Quote Board Refresh, because the Current
    ' Session Update does a Quote Board Refresh also)
    If m.dNextCurrentSessionUpdate <> 0# And dNow >= m.dNextCurrentSessionUpdate Then
        If Not bInProgress Then
DebugLog "***** Doing CSU: " & Format(Now, "HH:MM:SS") & " *****"
            bInProgress = True
            CheckCSU
            bInProgress = False
DebugLog "***** Done Doing CSU: " & Format(Now, "HH:MM:SS") & " *****"
        End If
        CalcNext eGDRow_CurrentSession
        CalcNext eGDRow_QuoteBoard
    End If
    
    ' See if we need to do a Quote Board refresh now...
    If m.dNextQuoteBoardRefresh <> 0# And dNow >= m.dNextQuoteBoardRefresh Then
        If Not bInProgress Then
DebugLog "***** Doing QB: " & Format(Now, "HH:MM:SS") & " *****"
            bInProgress = True
            CheckQuoteBoard
            bInProgress = False
DebugLog "***** Done Doing QB: " & Format(Now, "HH:MM:SS") & " *****"
        End If
        CalcNext eGDRow_QuoteBoard
    End If

    ' See if we need to do a Daily Download now...
    If m.dNextDailyDownload <> 0# And dNow >= m.dNextDailyDownload Then
        If Not bInProgress Then
DebugLog "***** Doing Daily Download: " & Format(Now, "HH:MM:SS") & " *****"
            bInProgress = True
            CheckDailyDownload
            bInProgress = False
DebugLog "***** Done Doing Daily Download: " & Format(Now, "HH:MM:SS") & " *****"
        End If
        CalcNext eGDRow_DailyDownload
    End If
    
    ' See if we need to do an Option Chain now...
    If m.dNextOptionChainRefresh <> 0# And dNow >= m.dNextOptionChainRefresh Then
        If Not bInProgress Then
DebugLog "***** Doing OC: " & Format(Now, "HH:MM:SS") & " *****"
            bInProgress = True
            CheckOptionChain
            bInProgress = False
DebugLog "***** Done Doing OC: " & Format(Now, "HH:MM:SS") & " *****"
        End If
        CalcNext eGDRow_OptionChain
    End If
    
    ' Update the Umbrella.RUN file...
    FileFromString m.strMonitorPath & "Umbrella.RUN", CStr(CDbl(Now))
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.tmr.Timer", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeFromString
'' Description: Figure out the time from the string in the grid
'' Inputs:      Row and Column of the cell
'' Returns:     Time in that cell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TimeFromString(ByVal lRow As Long, ByVal lCol As Long) As Long
On Error GoTo ErrSection:

    Dim strTemp As String
    Dim strTime As String
    Dim lReturn As Long
    
    strTemp = fgMonitor.TextMatrix(lRow, lCol)
    strTime = Parse(strTemp, " ", 1)
    
    lReturn = CLng(Parse(strTime, ":", 1))
    If lReturn = 12& Then
        If UCase(Parse(strTemp, " ", 2)) = "AM" Then lReturn = 0&
    Else
        If UCase(Parse(strTemp, " ", 2)) = "PM" Then lReturn = lReturn + 12
    End If
    lReturn = lReturn * 60
    lReturn = lReturn + CLng(Parse(strTime, ":", 2))
    
    TimeFromString = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmMonitoring.TimeFromString", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcNext
'' Description: Calculate the time for the next event for the given row
'' Inputs:      Row to Calculate
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcNext(ByVal Row As eGDRows)
On Error GoTo ErrSection:

    Dim dStartTime As Double            ' Time to start refreshing quote board
    Dim dEndTime As Double              ' Time to stop refreshing quote board
    Dim dNextTime As Double             ' Time of next interval
    Dim lInterval As Long               ' Interval to refresh quote board
    Dim dCurrentLocal As Double         ' Current Time Locally

    If CheckedCell(fgMonitor, Row, GDCol(eGDCol_On)) Then
        dCurrentLocal = Now
        
        ' Get Start and End Times from the Grid
        lInterval = CLng(fgMonitor.TextMatrix(Row, GDCol(eGDCol_Interval)))
        dStartTime = CDbl(Int(dCurrentLocal)) + (TimeFromString(Row, GDCol(eGDCol_Start)) / 1440#)
        dEndTime = CDbl(Int(dCurrentLocal)) + (TimeFromString(Row, GDCol(eGDCol_End)) / 1440#)
        
        ' If crosses midnight, bump end time to tomorrow
        If dEndTime < dStartTime Then dEndTime = dEndTime + 1
        
        ' Make sure not on a weekend
        ' (if starting after 4:30pm, then really want to run Sun-Thu)
        Do While Not IsWeekday(dStartTime + (7.5 / 24#))
            dStartTime = dStartTime + 1#
            dEndTime = dEndTime + 1#
        Loop
        
        ' Figure out next start time
        Do While dCurrentLocal > dEndTime '(make sure EndTime is in the future)
            Do
                dStartTime = dStartTime + 1#
                dEndTime = dEndTime + 1#
            Loop While Not IsWeekday(dStartTime + 7.5 / 24#)
        Loop
        dNextTime = dStartTime
        Do While dCurrentLocal > dNextTime
            dNextTime = dNextTime + (lInterval / 1440#)
            If dNextTime > dEndTime Then
                ' if past end time, go to beginning of next day
                Do
                    dStartTime = dStartTime + 1#
                    dEndTime = dEndTime + 1#
                Loop While Not IsWeekday(dStartTime + (7.5 / 24#))
                dNextTime = dStartTime
                Exit Do
            End If
        Loop
    End If

    Select Case Row
        Case GDRow(eGDRow_QuoteBoard)
            m.dNextQuoteBoardRefresh = dNextTime
frmTest2.AddList "QuoteBoard: " & Format(dNextTime, "MM/DD/YYYY HH:MM AM/PM")
        
        Case GDRow(eGDRow_OptionChain)
            m.dNextOptionChainRefresh = dNextTime
frmTest2.AddList "OptionChain: " & Format(dNextTime, "MM/DD/YYYY HH:MM AM/PM")
        
        Case GDRow(eGDRow_CurrentSession)
            m.dNextCurrentSessionUpdate = dNextTime
frmTest2.AddList "CurrentSession: " & Format(dNextTime, "MM/DD/YYYY HH:MM AM/PM")
        
        Case GDRow(eGDRow_DailyDownload)
            m.dNextDailyDownload = dNextTime
frmTest2.AddList "DailyDownload: " & Format(dNextTime, "MM/DD/YYYY HH:MM AM/PM")
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.CalcNext", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckQuoteBoard
'' Description: Do a check on the quote board refresh
'' Inputs:      Do a Download?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckQuoteBoard(Optional ByVal bDownload As Boolean = True)
On Error GoTo ErrSection:

    Dim lCheckDate As Long              ' Date to check for
    Dim strCheckDate As String          ' String version of date to check
    Dim Bars As New cGdBars             ' Bars to get data in
    Dim dTime As Double                 ' Current time
    Dim dNow As Double                  ' Current date and time
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRetry As Boolean               ' Whether to retry the download
    Dim bRetryStream As Boolean         ' Whether to retry the download
    Dim StreamBars As New cGdBars       ' Version of the bars to stream
    Dim dStartTime As Double            ' Start time of timeout period
    
DebugLog vbTab & "Check Quote Board"

    ' Check Date is Today unless we are after 5pm, then it is tomorrow...
    dNow = Now
    lCheckDate = Date
    dTime = Time
    If Weekday(lCheckDate) = vbFriday Then
        If Time > (15# / 24#) Then lCheckDate = lCheckDate + 1
    Else
        If Time > (17# / 24#) Then lCheckDate = lCheckDate + 1
    End If
    strCheckDate = Format(lCheckDate, "YYYYMMDD")
    
    ' If the check date is a weekday and not a holiday then do the check...
    If IsWeekday(lCheckDate) Then ' And Not m.astrHolidays.BinarySearch(strCheckDate) Then
        bRetry = False
        For lIndex = 1 To 2
DebugLog vbTab & "Try " & Str(lIndex)
            If bDownload Then
                ' Clear the snapshot area...
DebugLog vbTab & "Clear Snapshot Area"
                ClearSnapshotData
                
                ' Request the Quote Board Refresh...
                Set MsgForm = frmStatus
                'g.RealTime.RefreshSymbolList True
DebugLog vbTab & "Request Quote Board Refresh and Turn on Stream"
                TurnOnStream
                Set MsgForm = Nothing
            End If
            
DebugLog vbTab & "Check $EUR-USD End-of-Day Data"
            ' Check $EUR-USD
            DM_GetBars Bars, "$EUR-USD", , lCheckDate
            If bDownload Then
                Set StreamBars = Bars.MakeCopy
                g.RealTime.AddTickBuffer StreamBars
                g.RealTime.SpliceBars StreamBars
            End If
            If Bars.Size = 0 Then
                ' Error - No Daily Bar for $EUR-USD
                OutputError lIndex & "11", "No Daily Bar for $EUR-USD (QB)"
                bRetry = True
            Else
DebugLog vbTab & "Check $EUR-USD Tick-By-Tick Data"
                DM_GetBars Bars, "$EUR-USD", ePRD_EachTick, lCheckDate
                If Bars.Size = 0 Then
                    ' Error - No Tick Data for $EUR-USD
                    OutputError lIndex & "12", "No Tick Data for $EUR-USD (QB)"
                    bRetry = True
                Else
                    If ConvertTimeZone(Bars(eBARS_DateTime, Bars.Size - 1), "GMT", "") < (dNow - (5# / 1440#)) Then
                        ' Error - Tick Data too old
                        OutputError lIndex & "13", "Old Tick Data for $EUR-USD (QB)"
                        bRetry = True
                    End If
                End If
            End If
            
            If Not m.astrHolidays.BinarySearch(strCheckDate) Then
                ' Check ES-067...
DebugLog vbTab & "Check ES-067 End-of-Day Data"
                DM_GetBars Bars, "ES-067", ePRD_Days + 1, lCheckDate
                If Bars.Size = 0 Then
                    ' Error - No Daily Bar for ES-067
                    OutputError lIndex & "11", "No Daily Bar for ES-067 (QB)"
                    bRetry = True
                Else
DebugLog vbTab & "Check ES-067 Tick-By-Tick Data"
                    DM_GetBars Bars, "ES-067", ePRD_EachTick, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Tick Data for ES-067
                        OutputError lIndex & "12", "No Tick Data for ES-067 (QB)"
                        bRetry = True
                    Else
                        ' Do different check for between 5:00am and 2:15pm...
                        If dTime >= (300# / 1440#) And dTime <= (855# / 1440#) Then
                            If ConvertTimeZone(Bars(eBARS_DateTime, Bars.Size - 1), "NY", "") < (dNow - (5# / 1440#)) Then
                                ' Error - Tick Data too old
                                OutputError lIndex & "13", "Old Tick Data for ES-067 (QB)"
                                bRetry = True
                            End If
                        Else
                            If ConvertTimeZone(Bars(eBARS_DateTime, Bars.Size - 1), "NY", "") < (dNow - (90# / 1440#)) Then
                                ' Error - Tick Data too old
                                OutputError lIndex & "13", "Old Tick Data for ES-067 (QB)"
                                bRetry = True
                            End If
                        End If
                    End If
                End If
                
                ' Check $DJIA (Between 7:35am and 2:00pm)...
                If dTime >= (455# / 1440#) And dTime <= (840# / 1440#) And (Not bRetry) Then
DebugLog vbTab & "Check $DJIA End-of-Day Data"
                    DM_GetBars Bars, "$DJIA", ePRD_Days + 1, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Daily Bar for $DJIA
                        OutputError lIndex & "11", "No Daily Bar for $DJIA (QB)"
                        bRetry = True
                    Else
DebugLog vbTab & "Check $DJIA Tick-By-Tick Data"
                        DM_GetBars Bars, "$DJIA", ePRD_EachTick, lCheckDate
                        If Bars.Size = 0 Then
                            ' Error - No Tick Data for $DJIA
                            OutputError lIndex & "12", "No Tick Data for $DJIA (QB)"
                            bRetry = True
                        ElseIf ConvertTimeZone(Bars(eBARS_DateTime, Bars.Size - 1), "NY", "") < (dNow - (5# / 1440#)) Then
                            ' Error - Tick Data too old
                            OutputError lIndex & "13", "Old Tick Data for $DJIA (QB)"
                            bRetry = True
                        End If
                    End If
                End If
                
                ' Check IBM (Between 7:55am and 2:00pm)...
                If dTime >= (475# / 1440#) And dTime <= (840# / 1440#) And (Not bRetry) Then
DebugLog vbTab & "Check IBM End-of-Day Data"
                    DM_GetBars Bars, "IBM", ePRD_Days + 1, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Daily Bar for IBM
                        OutputError lIndex & "11", "No Daily Bar for IBM (QB)"
                        bRetry = True
                    Else
DebugLog vbTab & "Check IBM Tick-By-Tick Data"
                        DM_GetBars Bars, "IBM", ePRD_EachTick, lCheckDate
                        If Bars.Size = 0 Then
                            ' Error - No Tick Data for IBM
                            OutputError lIndex & "12", "No Tick Data for IBM (QB)"
                            bRetry = True
                        ElseIf ConvertTimeZone(Bars(eBARS_DateTime, Bars.Size - 1), "NY", "") < (dNow - (25# / 1440#)) Then
                            ' Error - Tick Data too old
                            OutputError lIndex & "13", "Old Tick Data for IBM (QB)"
                            bRetry = True
                        End If
                    End If
                End If
                
                ' Check SP-067 (Between 7:35am and 2:20pm)...
                If dTime >= (455# / 1440#) And dTime <= (860# / 1440#) And (Not bRetry) Then
DebugLog vbTab & "Check SP-067 End-of-Day Data"
                    DM_GetBars Bars, "SP-067", ePRD_Days + 1, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Daily Bar for SP-067
                        OutputError lIndex & "11", "No Daily Bar for SP-067 (QB)"
                        bRetry = True
                    Else
DebugLog vbTab & "Check SP-067 Tick-By-Tick Data"
                        DM_GetBars Bars, "SP-067", ePRD_EachTick, lCheckDate
                        If Bars.Size = 0 Then
                            ' Error - No Tick Data for SP-067
                            OutputError lIndex & "12", "No Tick Data for SP-067 (QB)"
                            bRetry = True
                        ElseIf ConvertTimeZone(Bars(eBARS_DateTime, Bars.Size - 1), "NY", "") < (dNow - (5# / 1440#)) Then
                            ' Error - Tick Data too old
                            OutputError lIndex & "13", "Old Tick Data for SP-067 (QB)"
                            bRetry = True
                        End If
                    End If
                End If
            End If
            
            ' Check to make sure stream is active...
            If bDownload = True Then
DebugLog vbTab & "Check Stream"
                DoEvents
                If g.RealTime.IsServerActive(True) = False Then
                    OutputError lIndex & "14", "Not connected to stream"
                    bRetryStream = True
                End If
                
                ' Check to make sure we can get streaming data...
                If Not bRetryStream Then
                    dStartTime = gdTickCount
                    Do While True
                        DoEvents
                        If g.RealTime.UpdateBars(StreamBars) = True Then
                            Exit Do
                        End If
                        If gdTickCount - dStartTime > 30000 Then
                            OutputError lIndex & "15", "No realtime data for " & StreamBars.Prop(eBARS_Symbol)
                            bRetryStream = True
                            Exit Do
                        End If
                    Loop
                End If
                
                Set StreamBars = Nothing
                g.RealTime.Init False
                Sleep 5
            End If
            
            If (Not bDownload) Or ((Not bRetry) And (Not bRetryStream)) Then Exit For
        Next lIndex
    End If
    
    If (Not bRetry) And (Not bRetryStream) Then
        fgMonitor.TextMatrix(GDRow(eGDRow_QuoteBoard), GDCol(eGDCol_Status)) = "OK"
    End If
    
DebugLog vbTab & "Done Check Quote Board"

ErrExit:
    Set Bars = Nothing
    Exit Sub
    
ErrSection:
    Set Bars = Nothing
    RaiseError "frmMonitoring.CheckQuoteBoard", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckCSU
'' Description: Do a check on the current session update
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckCSU()
On Error GoTo ErrSection:

    Dim lCheckDate As Long              ' Date to check for
    Dim strCheckDate As String          ' String version of date to check
    Dim Bars As New cGdBars             ' Bars to get data in
    Dim dTime As Double                 ' Current time
    Dim dNow As Double                  ' Current date and time
    Dim astrFiles As New cGdArray       ' Array of matching files
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRetry As Boolean               ' Whether to retry the download
    
DebugLog vbTab & "Check Current Session Update"
    
    ' Check Date is Today unless we are after 5pm, then it is tomorrow...
    dNow = Now
    lCheckDate = Date
    dTime = Time
    strCheckDate = Format(lCheckDate, "YYYYMMDD")
    
    ' If the checkdate is a weekday and not a holiday then do the check...
    If IsWeekday(lCheckDate) And Not m.astrHolidays.BinarySearch(strCheckDate) Then
        If dTime >= (340# / 1440#) And dTime <= (940# / 1440#) Then
            bRetry = False
            For lIndex = 1 To 2
DebugLog vbTab & "Try " & Str(lIndex)
                
                ' Delete the files in FTP\Backup...
                KillFile AddSlash(App.Path) & "FTP\Backup\Today*.*"
            
DebugLog vbTab & "Clear Snapshot Data"
                ' Clear the snapshot area...
                ClearSnapshotData
                
DebugLog vbTab & "Request Current Session Update"
                ' Request the Current Session Update...
                Set MsgForm = frmStatus
                frmDownload.optCurrentSession = True
                frmDownload.DownloadData
                Set MsgForm = Nothing
                
DebugLog vbTab & "Check Files"
                ' Check to see if files exist in FTP\Backup again...
                astrFiles.GetMatchingFiles AddSlash(App.Path) & "FTP\Backup\Today*.GZP"
                If astrFiles.Size = 0 Then
                    ' Error - Files were not downloaded
                    OutputError lIndex & "20", "Files were not downloaded for " & strCheckDate & " (CS)"
                    bRetry = True
                End If
            
                ' Check NQ-067...
                If Not bRetry Then
DebugLog vbTab & "Check NQ-067"
                    DM_GetBars Bars, "NQ-067", ePRD_Days + 1, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Daily Bar for NQ-067
                        OutputError lIndex & "21", "No Daily Bar for NQ-067 (CS)"
                        bRetry = True
                    End If
                End If
                
                ' Check $COMPQ (After 8:10am)
                If dTime > (490# / 1440#) And (Not bRetry) Then
DebugLog vbTab & "Check $COMPQ"
                    DM_GetBars Bars, "$COMPQ", ePRD_Days + 1, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Daily Bar for $COMPQ
                        OutputError lIndex & "21", "No Daily Bar for $COMPQ (CS)"
                        bRetry = True
                    End If
                End If
                
                ' Check MSFT (After 8:10am)
                If dTime > (490# / 1440#) And (Not bRetry) Then
DebugLog vbTab & "Check MSFT"
                    DM_GetBars Bars, "MSFT", ePRD_Days + 1, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Daily Bar for MSFT
                        OutputError lIndex & "21", "No Daily Bar for MSFT (CS)"
                        bRetry = True
                    End If
                End If
                
                ' Check the Quote Board stuff since it piggy backed along...
                CheckQuoteBoard False
                
                If (Not bRetry) Then Exit For
            Next lIndex
        End If
    End If

    If Not bRetry Then
        fgMonitor.TextMatrix(GDRow(eGDRow_CurrentSession), GDCol(eGDCol_Status)) = "OK"
    End If
    
DebugLog vbTab & "Done Check Current Session Update"
    
ErrExit:
    Set astrFiles = Nothing
    Set Bars = Nothing
    Exit Sub
    
ErrSection:
    Set astrFiles = Nothing
    Set Bars = Nothing
    RaiseError "frmMonitoring.CheckCSU", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckDailyDownload
'' Description: Do a check on the daily download
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckDailyDownload()
On Error GoTo ErrSection:

    Dim lCheckDate As Long              ' Date to check for
    Dim strCheckDate As String          ' String version of date to check
    Dim Bars As New cGdBars             ' Bars to get data in
    Dim dTime As Double                 ' Current time
    Dim dNow As Double                  ' Current date and time
    Dim astrFiles As New cGdArray       ' Array of matching files
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRetry As Boolean               ' Whether or not to retry
    
DebugLog vbTab & "Check Daily Download"
    
    ' Check Date is Today unless we are after 5pm, then it is tomorrow...
    dNow = Now
    lCheckDate = Date
    dTime = Time
    If dTime < (1020# / 1440#) Or Not IsWeekday(lCheckDate) Then
        Do
            lCheckDate = lCheckDate - 1&
            strCheckDate = Format(lCheckDate, "YYYYMMDD")
            If IsWeekday(lCheckDate) Then Exit Do
        Loop
    End If
    
    For lIndex = 1 To 2
DebugLog vbTab & "Try " & Str(lIndex)
    
        ' Delete files for the checkdate in FTP\Backup...
        KillFile AddSlash(App.Path) & "FTP\Backup\" & Right(strCheckDate, 6) & "*.*"
        
DebugLog vbTab & "Clear Snapshot Data"
        ' Clear the snapshot area...
        ClearSnapshotData
        
DebugLog vbTab & "Request Daily Download"
        ' Request the Daily Download...
        Set MsgForm = frmStatus
        frmDownload.dtpFromDate = lCheckDate
        frmDownload.optDaily = True
        frmDownload.DownloadData
        Set MsgForm = Nothing
        
DebugLog vbTab & "Check Files"
        ' Check to see if files exist in FTP\Backup again...
        astrFiles.GetMatchingFiles AddSlash(App.Path) & "FTP\Backup\" & Right(strCheckDate, 6) & "*.GZP"
        If astrFiles.Size = 0 Then
            ' Error - Files were not downloaded
            OutputError lIndex & "30", "Files were not downloaded for " & strCheckDate & " (DD)"
            bRetry = True
        End If
    
        ' Check the data as long as the check date is not a holiday...
        If Not m.astrHolidays.BinarySearch(strCheckDate) Then
            ' Check ES-067...
            If Not bRetry Then
DebugLog vbTab & "Check ES-067 End-of-Day Data"
                DM_GetBars Bars, "ES-067", ePRD_Days + 1, lCheckDate, lCheckDate
                If Bars.Size = 0 Then
                    ' Error - No Daily Bars for ES-067
                    OutputError lIndex & "31", "No Daily Data for ES-067 (DD)"
                    bRetry = True
                Else
DebugLog vbTab & "Check ES-067 Tick-By-Tick Data"
                    DM_GetBars Bars, "ES-067", ePRD_EachTick, lCheckDate, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Tick Data for ES-067
                        OutputError lIndex & "32", "No Tick Data for ES-067 (DD)"
                        bRetry = True
                    End If
                End If
            End If
            
            ' Check $DJIA...
            If Not bRetry Then
DebugLog vbTab & "Check $DJIA End-of-Day Data"
                DM_GetBars Bars, "$DJIA", ePRD_Days + 1, lCheckDate, lCheckDate
                If Bars.Size = 0 Then
                    ' Error - No Daily Bars for $DJIA
                    OutputError lIndex & "31", "No Daily Data for $DJIA (DD)"
                    bRetry = True
                Else
DebugLog vbTab & "Check $DJIA Tick-By-Tick Data"
                    DM_GetBars Bars, "$DJIA", ePRD_EachTick, lCheckDate, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Tick Data for $DJIA
                        OutputError lIndex & "32", "No Tick Data for $DJIA (DD)"
                        bRetry = True
                    End If
                End If
            End If
            
            ' Check IBM...
            If Not bRetry Then
DebugLog vbTab & "Check IBM End-of-Day Data"
                DM_GetBars Bars, "IBM", ePRD_Days + 1, lCheckDate, lCheckDate
                If Bars.Size = 0 Then
                    ' Error - No Daily Bars for IBM
                    OutputError lIndex & "31", "No Daily Data for IBM (DD)"
                    bRetry = True
                Else
DebugLog vbTab & "Check IBM Tick-By-Tick Data"
                    DM_GetBars Bars, "IBM", ePRD_EachTick, lCheckDate, lCheckDate
                    If Bars.Size = 0 Then
                        ' Error - No Tick Data for IBM
                        OutputError lIndex & "32", "No Tick Data for IBM (DD)"
                        bRetry = True
                    End If
                End If
            End If
        End If
        
        If (Not bRetry) Then Exit For
    Next lIndex
    
    If Not bRetry Then
        fgMonitor.TextMatrix(GDRow(eGDRow_DailyDownload), GDCol(eGDCol_Status)) = "OK"
    End If
    
DebugLog vbTab & "Done Check Daily Download"
    
ErrExit:
    Set Bars = Nothing
    Set astrFiles = Nothing
    Exit Sub
    
ErrSection:
    Set Bars = Nothing
    Set astrFiles = Nothing
    RaiseError "frmMonitoring.CheckDailyDownload", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckOptionChain
'' Description: Do a check on the option chain
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckOptionChain()
On Error GoTo ErrSection:

    Dim lCheckDate As Long              ' Date to check for
    Dim strCheckDate As String          ' String version of date to check
    Dim Bars As New cGdBars             ' Bars to get data in
    Dim dTime As Double                 ' Current time
    Dim dNow As Double                  ' Current date and time
    Dim lIndex As Long                  ' Index into a for loop
    Dim bRetry As Boolean               ' Whether or not to retry
    
DebugLog vbTab & "Check Option Chain"
    
    ' Check Date is Today...
    dNow = Now
    lCheckDate = Date
    dTime = Time
    strCheckDate = Format(lCheckDate, "YYYYMMDD")
    
    For lIndex = 1 To 2
DebugLog vbTab & "Try " & Str(lIndex)
    
DebugLog vbTab & "Clear Snapshot Data"
        ' Clear the snapshot area...
        ClearSnapshotData
        
DebugLog vbTab & "Request Option Chain for SP-067"
        ' Check Option Chain for SP-067...
        Set MsgForm = frmStatus
        If Not frmOptionChain.ShowMe(g.SymbolPool.SymbolIDforSymbol("SP-067"), True) Then
            ' Error - Did not get option data for SP-067
            OutputError lIndex & "44", "No Option Chain for SP-067 (OC)"
            bRetry = True
        End If
        frmOptionChain.cmdClose_Click
        
        ' Check Option Chain for $OEX...
        If (Not bRetry) Then
DebugLog vbTab & "Request Option Chain for $OEX"
            If Not frmOptionChain.ShowMe(g.SymbolPool.SymbolIDforSymbol("$OEX"), True) Then
                ' Error - Did not get option data for $OEX
                OutputError lIndex & "44", "No Option Chain for $OEX (OC)"
                bRetry = True
            End If
            frmOptionChain.cmdClose_Click
        End If
    
        ' Check Option Chain for IBM...
        If (Not bRetry) Then
DebugLog vbTab & "Request Option Chain for IBM"
            If Not frmOptionChain.ShowMe(g.SymbolPool.SymbolIDforSymbol("IBM"), True) Then
                ' Error - Did not get option data for IBM
                OutputError lIndex & "44", "No Option Chain for IBM (OC)"
                bRetry = True
            End If
            frmOptionChain.cmdClose_Click
        End If
        Set MsgForm = Nothing
        
        If (Not bRetry) Then Exit For
    Next lIndex

    If Not bRetry Then
        fgMonitor.TextMatrix(GDRow(eGDRow_OptionChain), GDCol(eGDCol_Status)) = "OK"
    End If
    
DebugLog vbTab & "Done Check Option Chain"
    
ErrExit:
    Set Bars = Nothing
    Exit Sub
    
ErrSection:
    Set Bars = Nothing
    RaiseError "frmMonitoring.CheckOptionChain", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OutputError
'' Description: Output the error as deemed appropriate
'' Inputs:      Error to output, Message to output
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OutputError(ByVal strError As String, ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim astrOut As New cGdArray         ' Array to output as the error file
    Dim strFilename As String           ' Name and Path of the file
    Dim lRow As Long                    ' Corresponding row in the grid
    Dim lIPLastSuccess As Long          ' Last successful IP index from registry
    Dim strIPAddresses As String        ' IP Addresses from the registry
    Dim strKey As String                ' Key into the registry
    Dim strIPAddress As String          ' Last IP Address tried
    
    ' Get IP Address from the registry...
    strKey = "Software\Genesis Financial Data Services\GClient"
#If 0 Then
    lIPLastSuccess = GetRegistryValue(rkLocalMachine, strKey, "IPLastSuccess", 0&)
    strIPAddresses = GetRegistryValue(rkLocalMachine, strKey, "IPAddresses", "")
    strIPAddress = Parse(strIPAddresses, "~", lIPLastSuccess + 2)
#Else
    strIPAddress = GetRegistryValue(rkLocalMachine, strKey, "HostLastAttempt", "")
#End If
    
    ' Put together the file name and path...
    strFilename = m.strMonitorPath & "Umbrella.ERR"
    
    ' Build the file to output...
    astrOut.Create eGDARRAY_Strings
    astrOut.Add "Umbrella." & strError & " | " & strMessage & " (IP=" & strIPAddress & ")"
    astrOut.ToFile strFilename, True

    lRow = CLng(Mid(strError, 2, 1))
    fgMonitor.TextMatrix(lRow, GDCol(eGDCol_Status)) = strMessage
    
ErrExit:
    Set astrOut = Nothing
    Exit Sub
    
ErrSection:
    Set astrOut = Nothing
    RaiseError "frmMonitoring.OutputError", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TurnOnStream
'' Description: Turn on the stream and wait for everything to get done
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TurnOnStream()
On Error GoTo ErrSection:

    Dim dStartTime As Double            ' Starting time of the check

    g.RealTime.Init True
    
    dStartTime = gdTickCount
    Do While True
        Sleep 0.5
        If (ProcessIsBusy(True) = True) Or (gdTickCount - dStartTime > 5000) Then
            Exit Do
        End If
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMonitoring.TurnOnStream", eGDRaiseError_Raise
    
End Sub
