VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Begin VB.Form frmAlertMessages 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Alert Messages"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7LCtl.VSFlexGrid fgMessages 
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1815
      _cx             =   3201
      _cy             =   2778
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
   Begin gdOCX.gdSelectColor gdColor 
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      CustomColor     =   255
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuEditAlert 
         Caption         =   "Edit this alert"
      End
      Begin VB.Menu mnuRemoveMessage 
         Caption         =   "Delete this message"
      End
      Begin VB.Menu mnuSetColor 
         Caption         =   "Set color for this message"
      End
      Begin VB.Menu mnuClearColor 
         Caption         =   "Clear color for this message"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearColorAll 
         Caption         =   "Clear color for all messages"
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "Delete all messages"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadHistory 
         Caption         =   "Show messages from history file"
      End
      Begin VB.Menu mnuUnloadHistory 
         Caption         =   "Hide messages from history file"
      End
      Begin VB.Menu mnuDeleteHistory 
         Caption         =   "Delete messages from history file"
      End
      Begin VB.Menu mnuDaysToKeep 
         Caption         =   "Days to keep messages in history file"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStopSound 
         Caption         =   "Stop sound"
      End
      Begin VB.Menu mnuAlertsSetup 
         Caption         =   "View setup for all alerts"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "Test Message"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAlertMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kDefaultHSDays = 10

' Columns in the grid
Private Enum eMsgCols
    eMsgCols_DateString = 0
    eMsgCols_AlertType
    eMsgCols_Symbol
    eMsgCols_AlertText
    eMsgCols_AlertKey
    eMsgCols_DateDouble
    eMsgCols_TimeZone
    eMsgCols_Color
    eMsgCols_NumCols
End Enum

Private Type mPrivate
    astrHistory As cGdArray             'Array of alert messages read from file
    astrAlertMessages As cGdArray       'Array of alert messages not yet saved to file
    
    iMouseRowDown As Long
    iMouseDownX As Single
    iMouseDownY As Single
    
    iNumDays As Long                    'Number of days to keep history before removing
    iSortOrder As Long
    
    bHistoryExist As Boolean
    bHistoryChanged As Boolean
    bMessageChanged As Boolean
End Type

Private m As mPrivate

Private Function GDCol(ByVal nCol As eMsgCols) As Long
    GDCol = nCol
End Function

Public Sub ShowMe()
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    Dim strMsg As String
                
    If bInProgress Then Exit Sub
    
    If frmMain.Enabled = False Then
        g.bShowAlertMsgForm = True
        GoTo ErrExit
    End If
    
    bInProgress = True

    Set m.astrHistory = New cGdArray
    m.astrHistory.Create eGDARRAY_Strings
    
    m.bHistoryChanged = False
    m.bMessageChanged = False
    m.bHistoryExist = FileExist(AddSlash(App.Path) & kstrHistoryFile)
    
    InitGrid
    LoadGrid
    
    If frmMain.Enabled = False Then
        g.bShowAlertMsgForm = True
    Else
        strMsg = "none"
        ShowForm Me, eForm_Nonmodal, frmMain, , ALT_GRID_ROW_COLOR
    End If
    
    bInProgress = False
    
ErrExit:
    Exit Sub

ErrSection:
    bInProgress = False
    RaiseError "frmAlertMessages.ShowMe"
    
End Sub

Private Sub ShowPopup()
On Error GoTo ErrSection:

    Dim Alert As cAlert
    Dim strKey As String
    Dim iColor As Long
    
    Dim bCanEdit As Boolean
    Dim bCanRemove As Boolean
        
    With fgMessages
        If m.iMouseRowDown >= .FixedRows Then
            bCanRemove = True
            strKey = .TextMatrix(m.iMouseRowDown, eMsgCols_AlertKey)
            If Len(strKey) <> 0 Then
                Set Alert = g.Alerts(strKey)
                If Not Alert Is Nothing Then bCanEdit = True
            End If
        End If
        If m.iMouseRowDown >= .FixedRows And m.iMouseRowDown < .Rows Then
            mnuSetColor.Visible = True
            iColor = .Cell(flexcpBackColor, m.iMouseRowDown)
            If iColor = .BackColor Or iColor = .BackColorAlternate Or iColor = 0 Then
                mnuClearColor.Visible = False
            Else
                mnuClearColor.Visible = True
            End If
        Else
            mnuSetColor.Visible = False
            mnuClearColor.Visible = False
        End If
    End With
        
    mnuRemoveMessage.Visible = bCanRemove
    mnuEditAlert.Visible = bCanEdit
    
    If m.astrHistory.Size > 0 Then
        mnuLoadHistory.Visible = False
        mnuUnloadHistory.Visible = True
    Else
        mnuLoadHistory.Visible = m.bHistoryExist
        mnuUnloadHistory.Visible = False
    End If
    
    mnuDeleteHistory.Visible = m.bHistoryExist
            
    Me.PopupMenu mnuPopUp

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.ShowPopup"

End Sub

Private Sub fgMessages_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    m.iSortOrder = Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.fgMessages_AfterSort"

End Sub

Private Sub fgMessages_DblClick()
On Error Resume Next:


    If fgMessages.Row >= fgMessages.FixedRows And fgMessages.Row < fgMessages.Rows Then
        m.iMouseRowDown = fgMessages.Row
        mnuEditAlert_Click
    End If

End Sub

Private Sub fgMessages_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    geTaskBarNotify 0, 0           'stop sound
    
    m.iMouseDownX = X
    m.iMouseDownY = Y
    
    If gdColor.Visible Then gdColor_ColorClicked
    
    If Button = vbRightButton Then
        m.iMouseRowDown = fgMessages.MouseRow
        ShowPopup
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.fgMessages_MouseDown"

End Sub

Private Sub fgMessages_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    With fgMessages
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .ToolTipText = .TextMatrix(.MouseRow, 2)                    '6101
        Else
            .ToolTipText = ""
        End If
    End With

End Sub

Private Sub fgMessages_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    geTaskBarNotify 0, 0           'stop sound

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.fgMessages_MouseUp"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement from the ini file
        
    g.Styler.StyleForm Me
    
    mnuPopUp.Visible = False
    m.iNumDays = GetIniFileProperty("AlertsSetup", kDefaultHSDays, "HSDays", g.strIniFile)
    If m.iNumDays <= 0 Then m.iNumDays = kDefaultHSDays
        
    strPlacement = GetIniFileProperty("AlertMessages", "", "Placement", g.strIniFile)
    If Len(strPlacement) > 0 Then
        SetFormPlacement Me, strPlacement, "LHTW"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.Form_Load"

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    geTaskBarNotify 0, 0       'stop sound
    m.iMouseRowDown = -1
    ShowPopup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.Form_MouseDown"

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    geTaskBarNotify 0, 0           'stop sound

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.Form_MouseUp"

End Sub

Private Sub Form_Resize()
On Error Resume Next

    With fgMessages
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End With

End Sub

Private Sub DeleteHistory()
On Error GoTo ErrSection:

    KillFile AddSlash(App.Path) & kstrHistoryFile, True
    m.bHistoryChanged = False
    m.bHistoryExist = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.DeleteHistory"

End Sub

Private Sub UnloadHistory()
On Error GoTo ErrSection:

    Dim i&

    If m.astrHistory.Size = 0 Then Exit Sub
    
    m.astrHistory.Size = 0
    
    If m.bHistoryChanged Then
        With fgMessages
            For i = .FixedRows To .Rows - 1
                If Len(.TextMatrix(i, eMsgCols_AlertKey)) = 0 Then
                    m.astrHistory.Add .TextMatrix(i, eMsgCols_DateDouble) & vbTab & .TextMatrix(i, eMsgCols_AlertType) _
                    & vbTab & .TextMatrix(i, eMsgCols_AlertText) & vbTab & .TextMatrix(i, eMsgCols_Color) & vbTab & .TextMatrix(i, eMsgCols_Symbol)
                End If
            Next
        End With
        If m.astrHistory.Size > 0 Then
            m.astrHistory.ToFile AddSlash(App.Path) & kstrHistoryFile
            m.astrHistory.Size = 0
        End If
        m.bHistoryChanged = False
    End If
    
    LoadGrid
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.UnloadHistory"

End Sub

Private Sub LoadHistory()
On Error GoTo ErrSection:

    Dim i&, j&
    Dim dDateTime#, dDisplayDate#, strText$
    Dim strType$, strMsg$, strTimeZone$, strColor$
    Dim strSymbol As String             ' Symbol for the alert
    
    If Not m.bHistoryExist Then Exit Sub        'precautionary
    
    ' Load messages from file...
    m.astrHistory.Size = 0
    m.astrHistory.FromFile AddSlash(App.Path) & kstrHistoryFile
    
    For i = m.astrHistory.Size - 1 To 0 Step -1
        If Int(Val(Parse(m.astrHistory(i), vbTab, 1))) < (Int(Val(Date)) - m.iNumDays) Then
            m.astrHistory.Remove i
            m.bHistoryChanged = True
        End If
    Next
        
    If m.astrHistory.Size = 0 Then
        KillFile AddSlash(App.Path) & kstrHistoryFile, True
        strText = "There are no messages in the history file less than " & Str(m.iNumDays) & " days old."
        InfBox strText, "I"
        m.bHistoryExist = False
        Exit Sub
    End If
            
    With fgMessages
        .Redraw = flexRDNone
        For i = 0 To m.astrHistory.Size - 1
            .Rows = .Rows + 1
            j = .Rows - 1
            
            strText = m.astrHistory(i)
            dDateTime = Val(Parse(strText, vbTab, 1))
            dDisplayDate = dDateTime
            strType = Parse(strText, vbTab, 2)
            strMsg = Parse(strText, vbTab, 3)
            strTimeZone = Parse(strText, vbTab, 4)
            strColor = Parse(strText, vbTab, 5)
            strSymbol = Parse(strText, vbTab, 6)
            
            If g.bShowInLocalTimeZone Then dDisplayDate = ConvertTimeZone(dDateTime, strTimeZone, "")
            
            .TextMatrix(j, eMsgCols_DateString) = DateFormat(dDisplayDate, MM_DD_YYYY, HH_MM_SS)
            .TextMatrix(j, eMsgCols_AlertType) = strType
            .TextMatrix(j, eMsgCols_Symbol) = strSymbol
            .TextMatrix(j, eMsgCols_AlertText) = strMsg
            .TextMatrix(j, eMsgCols_AlertKey) = ""                  'blank to indicate message came from history file
            .TextMatrix(j, eMsgCols_DateDouble) = dDateTime
            .TextMatrix(j, eMsgCols_TimeZone) = strTimeZone
            .TextMatrix(j, eMsgCols_Color) = strColor
            
            If Len(strColor) > 0 Then
                .Cell(flexcpBackColor, j, 0, j, 5) = Val(strColor)
            End If
            .Cell(flexcpPicture, j, eMsgCols_DateString) = Picture16(ToolbarIcon("kFile"))
        Next
        
        .AutoSize 0, .Cols - 1, False, 75
        .Col = 0
        .Sort = m.iSortOrder
                
        .Redraw = flexRDBuffered
    End With
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.LoadHistory"

End Sub

Private Sub SetCellIcon()
On Error GoTo ErrSection:

    Dim Alert As cAlert
    Dim aTemp As New cGdArray
    Dim i&, dDateTime#, strKey$
    
    If fgMessages.Rows <= fgMessages.FixedRows Then Exit Sub
    
    With fgMessages
        .Redraw = flexRDNone
        
        .Col = 3
        .Sort = m.iSortOrder
        
        strKey = .TextMatrix(.FixedRows, eMsgCols_AlertKey)
        For i = .FixedRows + 1 To .Rows - 1
            If Len(strKey) = 0 Then Exit For
            
            If strKey <> .TextMatrix(i, eMsgCols_AlertKey) Then
                Set Alert = g.Alerts(strKey)
                If Not Alert Is Nothing Then
                    If Alert.Active Then
                        .Cell(flexcpPicture, i - 1, 1) = Picture16(ToolbarIcon("ID_Alerts"))
                    Else
                        .Cell(flexcpPicture, i - 1, 1) = Picture16(ToolbarIcon("kGrayBell"))
                    End If
                End If
            End If
            strKey = .TextMatrix(i, eMsgCols_AlertKey)
        Next
        .Col = 0
        .Sort = m.iSortOrder
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.SetCellIcon"

End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim Alert As cAlert
    Dim i&, j&
    Dim dDateTime#, dDisplayDate#, strText$
    Dim strType$, strMsg$, strKey$, strColor$
    Dim strSymbol As String             ' Symbol for the alert
        
    If m.astrAlertMessages Is Nothing Then
        Set m.astrAlertMessages = g.Alerts.AlertMsgArray
    End If
    
    If m.astrAlertMessages Is Nothing Then
        Set m.astrAlertMessages = New cGdArray
        m.astrAlertMessages.Create eGDARRAY_Strings
    Else
        With fgMessages
            .Redraw = flexRDNone
            
            .Rows = .FixedRows
            For i = 0 To m.astrAlertMessages.Size - 1
                strText = m.astrAlertMessages(i)
                dDateTime = Val(Parse(strText, vbTab, 1))
                dDisplayDate = dDateTime
                strType = Parse(strText, vbTab, 2)
                strMsg = Parse(strText, vbTab, 3)
                strKey = Parse(strText, vbTab, 4)
                strColor = Parse(strText, vbTab, 6)
                strSymbol = Parse(strText, vbTab, 7)
                
                'need to do this in case alert got removed since last time form was shown
                Set Alert = g.Alerts(strKey)
                
                If Not Alert Is Nothing Then
                    If dDateTime > 0 And Len(strType) > 0 And Len(strMsg) > 0 And Len(strKey) > 1 Then
                        .Rows = .Rows + 1
                        j = .Rows - 1
                        
                        If g.bShowInLocalTimeZone Then
                            dDisplayDate = ConvertTimeZone(dDateTime, Alert.LastCheckedTimeZone, "")
                        End If
                        
                        .TextMatrix(j, eMsgCols_DateString) = DateFormat(dDisplayDate, MM_DD_YYYY, HH_MM_SS)
                        .TextMatrix(j, eMsgCols_AlertType) = strType
                        .TextMatrix(j, eMsgCols_Symbol) = strSymbol
                        .TextMatrix(j, eMsgCols_AlertText) = strMsg
                        .TextMatrix(j, eMsgCols_AlertKey) = strKey
                        .TextMatrix(j, eMsgCols_DateDouble) = dDateTime
                        .TextMatrix(j, eMsgCols_TimeZone) = Alert.LastCheckedTimeZone
                        .TextMatrix(j, eMsgCols_Color) = strColor
                        If Len(strColor) > 0 Then
                            .Cell(flexcpBackColor, j, 0, j, 5) = Val(strColor)
                        End If
                    End If
                End If
            Next
            
            .AutoSize 0, .Cols - 1, False, 75
            .Col = 0
            .Sort = m.iSortOrder
            
            .Redraw = flexRDBuffered
        End With
        
        g.Alerts.AlertMsgArray = Nothing    'do this so array cannot be modified while this form is visible
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.LoadGrid"

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgMessages
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .WordWrap = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .MergeCells = flexMergeFree
        
        .Rows = 1
        .Cols = GDCol(eMsgCols_NumCols)
        .FixedRows = 1
        .FixedCols = 0
        
        .ColDataType(eMsgCols_DateString) = flexDTDate
        .ColDataType(eMsgCols_DateDouble) = flexDTDouble
               
        .TextMatrix(0, eMsgCols_DateString) = "Date"
        .TextMatrix(0, eMsgCols_AlertType) = "Type"
        .TextMatrix(0, eMsgCols_Symbol) = "Symbol"
        .TextMatrix(0, eMsgCols_AlertText) = "Alert"
        .TextMatrix(0, eMsgCols_AlertKey) = "Key"               'hidden col containing alert key or blank if message came from history file
        .TextMatrix(0, eMsgCols_DateDouble) = "DateTimeDbl"     'hidden col to store datetime as double
        .TextMatrix(0, eMsgCols_TimeZone) = "TimeZone"          'hidden col containing time zone of dateTime value
        .TextMatrix(0, eMsgCols_Color) = "Color"                'hidden col to store color for row
        
        .ColHidden(eMsgCols_AlertKey) = True
        .ColHidden(eMsgCols_DateDouble) = True
        .ColHidden(eMsgCols_TimeZone) = True
        .ColHidden(eMsgCols_Color) = True
        
        .ColAlignment(eMsgCols_DateString) = flexAlignCenterTop
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
                                
        .Redraw = flexRDBuffered
    End With
    
    m.iSortOrder = flexSortGenericDescending

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.InitGrid"

End Sub

Private Sub CheckChanges()
On Error GoTo ErrSection:

    Dim i&
    
    If Not m.bHistoryChanged And Not m.bMessageChanged Then Exit Sub

    m.astrHistory.Size = 0
    If m.bMessageChanged Then m.astrAlertMessages.Size = 0
    
    With fgMessages
        For i = .FixedRows To .Rows - 1
            If Len(.TextMatrix(i, eMsgCols_AlertKey)) = 0 Then
                If m.bHistoryChanged Then
                    'dateTime \t alertType \t alertText \t alert time zone \t color
                    m.astrHistory.Add .TextMatrix(i, eMsgCols_DateDouble) & vbTab & .TextMatrix(i, eMsgCols_AlertType) _
                        & vbTab & .TextMatrix(i, eMsgCols_AlertText) & vbTab & .TextMatrix(i, eMsgCols_TimeZone) _
                        & vbTab & .TextMatrix(i, eMsgCols_Color) & vbTab & .TextMatrix(i, eMsgCols_Symbol)
                End If
            ElseIf m.bMessageChanged Then
                'dateTime \t alertType \t alertText \t alertKey \t alert time zone \t color
                m.astrAlertMessages.Add .TextMatrix(i, eMsgCols_DateDouble) & vbTab & .TextMatrix(i, eMsgCols_AlertType) _
                    & vbTab & .TextMatrix(i, eMsgCols_AlertText) & vbTab & .TextMatrix(i, eMsgCols_AlertKey) _
                    & vbTab & .TextMatrix(i, eMsgCols_TimeZone) & vbTab & .TextMatrix(i, eMsgCols_Color) & vbTab & .TextMatrix(i, eMsgCols_Symbol)
            End If
        Next
    End With
    
    If m.bHistoryChanged Then
        If m.astrHistory.Size > 0 Then
            m.astrHistory.ToFile AddSlash(App.Path) & kstrHistoryFile
            m.astrHistory.Size = 0
        End If
    End If

    m.bHistoryChanged = False
    m.bMessageChanged = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.CheckChanges"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim aTemp As cGdArray
    
    geTaskBarNotify 0, 0       'stop sound
    CheckChanges
    
    If Not m.astrAlertMessages Is Nothing Then
        'precautionary check (array in alerts collection should be nothing when this form is loaded)
        Set aTemp = g.Alerts.AlertMsgArray
        If Not aTemp Is Nothing Then aTemp.Size = 0
        'replace array in alerts collection with this one
        g.Alerts.AlertMsgArray = m.astrAlertMessages
    End If

    SetIniFileProperty "AlertsSetup", m.iNumDays, "HSDays", g.strIniFile
    SetIniFileProperty "AlertMessages", GetFormPlacement(Me), "Placement", g.strIniFile
    
    Set m.astrHistory = Nothing
    Set m.astrAlertMessages = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.Form_Unload"

End Sub

Private Sub mnuAlertsSetup_Click()
On Error GoTo ErrSection:

    If FormIsLoaded("frmAlertsSetup") Then
        frmAlertsSetup.SetFocus
    Else
        frmAlertsSetup.ShowMe
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuAlertsSetup_Click"

End Sub

Private Sub mnuClearColor_Click()
On Error GoTo ErrSection:

    Dim i&
    
    With fgMessages
        If m.iMouseRowDown >= .FixedRows And m.iMouseRowDown < .Rows Then
'            If m.iMouseRowDown = .FixedRows Then
'                i = .FixedRows + 1
'            Else
'                i = m.iMouseRowDown - 1
'            End If
            If .Cell(flexcpBackColor, m.iMouseRowDown) = .BackColor Then
                .Cell(flexcpBackColor, m.iMouseRowDown, 0, m.iMouseRowDown, .Cols - 1) = .BackColorAlternate
            Else
                .Cell(flexcpBackColor, m.iMouseRowDown, 0, m.iMouseRowDown, .Cols - 1) = .BackColor
            End If
            .TextMatrix(m.iMouseRowDown, eMsgCols_Color) = ""
            If Len(.TextMatrix(m.iMouseRowDown, eMsgCols_AlertKey)) = 0 Then
                m.bHistoryChanged = True
            Else
                m.bMessageChanged = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuClearColor_Click"

End Sub

Private Sub mnuClearColorAll_Click()
On Error GoTo ErrSection:

    Dim i&

    With fgMessages
        For i = .FixedRows To .Rows - 1 Step 2
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = .BackColor
            .TextMatrix(i, eMsgCols_Color) = ""
            If i + 1 < .Rows Then
                .Cell(flexcpBackColor, i + 1, 0, i + 1, .Cols - 1) = .BackColorAlternate
                .TextMatrix(i + 1, eMsgCols_Color) = ""
            End If
        Next
    End With
    
    m.bHistoryChanged = True
    m.bMessageChanged = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuClearColorAll_Click"

End Sub

Private Sub mnuClose_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuClose_Click"

End Sub

Private Sub mnuDaysToKeep_Click()
On Error GoTo ErrSection:

    Dim s$

    s = InfBox("Number of days to keep messages in history file", "?", , "Alert History", , , , , , "s", Str(m.iNumDays))
    If Len(s) > 0 Then
        m.iNumDays = Abs(Int(Val(s)))
        If m.iNumDays = 0 Then m.iNumDays = 10
    Else
        m.iNumDays = 10
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuDaysToKeep_Click"

End Sub

Private Sub mnuDeleteAll_Click()
On Error GoTo ErrSection:

    DeleteHistory
    m.astrAlertMessages.Size = 0
    LoadGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuDeleteAll_Click"

End Sub

Private Sub mnuDeleteHistory_Click()
On Error GoTo ErrSection:

    DeleteHistory

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuDeleteHistory_Click"

End Sub

Private Sub mnuEditAlert_Click()
On Error GoTo ErrSection:

    Dim Alert As cAlert
    Dim strKey As String
    Dim eType As eGDAlertType
    
    strKey = fgMessages.TextMatrix(m.iMouseRowDown, eMsgCols_AlertKey)
    
    If Len(strKey) >= 2 Then
        Set Alert = g.Alerts(strKey)
        If Not Alert Is Nothing Then
            eType = Alert.AlertType
            If frmAlerts.ShowMe(Alert, eType) = True Then
                If eType = eGDAlertType_Time Then
                    Alert.CalcNextTriggerTime
                ElseIf eType = eGDAlertType_QuoteBoard Then
                    frmQuotes.DisplayAlert Alert, False
                    Alert.CheckAlert , , , True
                ElseIf eType = eGDAlertType_Annot Then
                    'If Not Alert.Annotation Is Nothing Then Alert.Annotation.CheckAnnotAlert True
                ElseIf eType = eGDAlertType_Chart Then
                    If Not Alert.Indicator Is Nothing Then Alert.Indicator.CheckIndAlert
                ElseIf Alert.CheckFromCheckAlerts Then
                    Alert.CheckAlert , , , True
                End If
            End If
        End If
    End If
    
    Set Alert = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuEditAlert_Click"

End Sub

Private Sub mnuLoadHistory_Click()
On Error GoTo ErrSection:

    LoadHistory

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuLoadHistory_Click"

End Sub

Public Sub AddAlertMessage(ByVal dDate As Double, ByVal strMessage As String, Alert As cAlert, Optional ByVal strSymbol As String = "")
On Error GoTo ErrSection:

    Dim strKey$, strAlertType$, j&, iColor&
    Dim strDate$, dDisplayDate#
    
    If Alert Is Nothing Then Exit Sub
    
    strAlertType = Alert.AlertTypeText
    strKey = Alert.AlertKey
    
    If dDate <= 0 Then Exit Sub
    If Len(strAlertType) = 0 Then Exit Sub
    If Len(strMessage) = 0 Then Exit Sub
    If Len(strKey) = 0 Then Exit Sub
    
    dDisplayDate = dDate
    iColor = Alert.MessageColor
    
    m.astrAlertMessages.Add dDate & vbTab & strAlertType & vbTab & strMessage & vbTab & strKey _
        & vbTab & Alert.LastCheckedTimeZone & vbTab & iColor & vbTab & strSymbol
    
    With fgMessages
        .Redraw = flexRDNone
        .Rows = .Rows + 1
        j = .Rows - 1
        
        If g.bShowInLocalTimeZone Then dDisplayDate = ConvertTimeZone(dDate, Alert.LastCheckedTimeZone, "")
        
        .TextMatrix(j, eMsgCols_DateString) = DateFormat(dDisplayDate, MM_DD_YYYY, HH_MM_SS)
        .TextMatrix(j, eMsgCols_AlertType) = strAlertType
        .TextMatrix(j, eMsgCols_Symbol) = strSymbol
        .TextMatrix(j, eMsgCols_AlertText) = strMessage
        .TextMatrix(j, eMsgCols_AlertKey) = strKey
        .TextMatrix(j, eMsgCols_DateDouble) = dDate
        .TextMatrix(j, eMsgCols_TimeZone) = Alert.LastCheckedTimeZone
        
        If iColor > 0 Then
            .Cell(flexcpBackColor, j, 0, j, 5) = iColor
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Col = 0
        .Sort = m.iSortOrder

        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.AddAlertMessage"

End Sub

Private Sub mnuRemoveMessage_Click()
On Error GoTo ErrSection:

    Dim tbTemp As New cGdTable
    Dim strKey$, i&

    With fgMessages
        If m.iMouseRowDown >= .FixedRows And m.iMouseRowDown < .Rows Then
            If Len(.TextMatrix(m.iMouseRowDown, eMsgCols_AlertKey)) = 0 Then
                m.bHistoryChanged = True
            Else
                m.bMessageChanged = True
            End If
            .RemoveItem m.iMouseRowDown
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuRemoveMessage_Click"

End Sub

Private Sub gdColor_Changed()
On Error Resume Next
    
    Dim iColor&
    
    gdColor.Visible = False
    
    If gdColor.Color = 0 Then
        iColor = 1          '0 is reserved number for grid
    Else
        iColor = gdColor.Color
    End If
    
    With fgMessages
        If m.iMouseRowDown >= .FixedRows And m.iMouseRowDown < .Rows Then
            .Cell(flexcpBackColor, m.iMouseRowDown, 0, m.iMouseRowDown, .Cols - 1) = iColor
            .TextMatrix(m.iMouseRowDown, eMsgCols_Color) = iColor
            If Len(.TextMatrix(m.iMouseRowDown, eMsgCols_AlertKey)) = 0 Then
                m.bHistoryChanged = True
            Else
                m.bMessageChanged = True
            End If
        End If
    End With
    

End Sub

Private Sub gdColor_ColorClicked()
On Error Resume Next

    If gdColor.DropDownVisible Then gdColor.UserControl_Click
    gdColor.Visible = False
    
End Sub

Private Sub mnuSetColor_Click()
On Error GoTo ErrSection:

    If m.iMouseRowDown >= fgMessages.FixedRows And m.iMouseRowDown < fgMessages.Rows Then
        gdColor.Move m.iMouseDownX, m.iMouseDownY
        gdColor.ZOrder
        gdColor.Visible = True
        gdColor.UserControl_Click
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuSetColor_Click"

End Sub

Private Sub mnuStopSound_Click()
On Error GoTo ErrSection:

    geTaskBarNotify 0, 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuStopSound_Click"

End Sub

Private Sub mnuTest_Click()
On Error GoTo ErrSection:

    'geTaskBarNotify 0, "MSFT <= 29.09"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuTest_Click"

End Sub

Private Sub mnuUnloadHistory_Click()
On Error GoTo ErrSection:

    UnloadHistory

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.mnuUnloadHistory_Click"

End Sub

Public Sub AlertRemoved(Alert As cAlert)
On Error GoTo ErrSection:

    Dim i&, iRow&, strKey$
    
    strKey = Alert.AlertKey
    
    With fgMessages
        For i = .FixedRows To .Rows - 1
            If strKey = .TextMatrix(i, eMsgCols_AlertKey) Then
                iRow = i
                Exit For
            End If
        Next
        If iRow >= .FixedRows And iRow < .Rows Then
            .Cell(flexcpPicture, iRow, eMsgCols_AlertType) = ""
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.AlertRemoved"

End Sub

Public Sub AlertActiveChanged(Alert As cAlert)
On Error GoTo ErrSection:

    Dim i&, iRow&, strKey$

'    If Alert Is Nothing Then Exit Sub
'
'    strKey = Alert.AlertKey
'
'    With fgMessages
'        For i = .FixedRows To .Rows - 1
'            If strKey = .TextMatrix(i, 3) Then
'                iRow = i
'                Exit For
'            End If
'        Next
'    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertMessages.AlertActiveChanged"

End Sub

