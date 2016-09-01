VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{16DE7640-1A28-11D6-B28B-0080C7A5F099}#1.7#0"; "gdDockable.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msInet.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00F2F2F2&
   Caption         =   "Trade Navigator"
   ClientHeight    =   5955
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11085
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   Begin VB.Timer tmrGridScrollPressed 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3840
      Top             =   5100
   End
   Begin VB.Timer tmrPlaySound 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3240
      Top             =   5100
   End
   Begin VB.PictureBox pbTbBackDraw 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   11085
      TabIndex        =   3
      Top             =   1215
      Visible         =   0   'False
      Width           =   11085
      Begin VB.Image imgTbBackDraw 
         Height          =   240
         Index           =   0
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   60
         Visible         =   0   'False
         Width           =   6300
      End
   End
   Begin VB.PictureBox pbTbBack 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   415
      Index           =   0
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   11085
      TabIndex        =   1
      Top             =   795
      Visible         =   0   'False
      Width           =   11085
      Begin HexUniControls.ctlUniComboBoxXP cboBarPeriod 
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   -15
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tip             =   "frmMain.frx":9C92
         Sorted          =   0   'False
         HScroll         =   0   'False
         Style           =   2
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
         ButtonWidth     =   17
         Locked          =   0   'False
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
         TrapTab         =   0   'False
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":9CB2
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin VB.PictureBox pbNotUsed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   4
         Top             =   75
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgTbBack 
         Height          =   585
         Index           =   0
         Left            =   4515
         Stretch         =   -1  'True
         Top             =   -75
         Width           =   4650
      End
   End
   Begin VB.Timer tmrAutoResize 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   810
      Top             =   4545
   End
   Begin VB.Timer tmrSymbol 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2610
      Top             =   5085
   End
   Begin VB.Timer tmrPredLabs 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2010
      Top             =   5085
   End
   Begin InetCtlsObjects.Inet INet 
      Left            =   5520
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin gdOCX.gdAppMail apmNews 
      Left            =   2730
      Top             =   3090
      _ExtentX        =   953
      _ExtentY        =   847
      ControlName     =   "NewsTN"
   End
   Begin VB.Timer tmrCheckBuySellButtons 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1410
      Top             =   5085
   End
   Begin VB.Timer tmrQuickStart 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   810
      Top             =   5085
   End
   Begin VB.Timer tmrWindowLink 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   210
      Top             =   5085
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1890
      Top             =   3030
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin gdDockable.DockablePro DockPro 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      Top             =   0
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   1402
      Persistant      =   0   'False
   End
   Begin gdOCX.gdAppMail apmRTClient 
      Left            =   1050
      Top             =   3090
      _ExtentX        =   953
      _ExtentY        =   847
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   7065
      Top             =   4875
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      MenuAnimations  =   3
      ToolBarsCount   =   5
      ToolsCount      =   289
      PersonalizedMenus=   0
      Style           =   0
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmMain.frx":9CCE
      ToolBars        =   "frmMain.frx":BB5F2
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   3090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   210
      Top             =   4545
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5700
      Visible         =   0   'False
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    nInitialTop As Long
    iPrevNonminimizedState As Integer
    strNormalPlacement As String
    bShouldBeSeen As Boolean
    bReadyToUnload As Boolean
    dWaitToCheckRT As Double
    
    nSymbolLinkColor As Long
    nSymbolLinkID As Long
    nPeriodLinkColor As Long
    nPeriodLinkID As Long
    
    oBtnMouseLast As cPicBoxButton          'button object that mouse was last in
    nBtnsPerRow As Long                     'number of buttons per row when toolbar is wrapped
    
    bToolbarWrap As Boolean
    strLastToolID As String                 'ID of last tool that was clicked
    aTbButtons As New cGdArray              'array of button objects for non-drawing toolbars
    aTbButtonsDraw As New cGdArray          'array of button objects for drawing toolbar
    
    aSoundToPlay As New cGdArray
End Type
Private m As mPrivate

Private Sub apmNews_MessageReceived(msg As gdOCX.gdAppMailMsg)
On Error GoTo ErrSection:

    Dim s$, i&
    Dim a As cGdArray
    Dim frm As frmSymbolSelector

    Select Case msg.MsgType
    Case -1: ' from the WebShell program
        s = msg.Message
        If Left(s, 4) = "+CP" & vbTab Then
            ' load the shared chart page
            LoadSharedChartPage s
            ' close down the Shared Chart Pages window
            apmNews.CreateMessage msg.FromControlName, 1, ""
        ElseIf Not IsIDE Then
            'symbol group from web stock screener
            mMain.CreateGroupFromScreener s
        End If
        
    Case 1:
        If 1 Then ' frmMain.Enabled Then
            Set frm = New frmSymbolSelector
            s = msg.Message
            If Len(Trim(s)) = 0 Then s = "Symbol Selector"
            i = Len(frm.Caption)
            Set a = frm.ShowMe(, , , s, , , True)
            s = ""
            If Not a Is Nothing Then
                For i = 0 To a.Size - 1
                    If Len(s) = 0 Then
                        s = a(i)
                    Else
                        s = s & vbTab & a(i)
                    End If
                Next
            End If
            Set frm = Nothing
        Else
            s = "."
        End If
        apmNews.CreateMessage msg.FromControlName, 2, s
    Case Else
        apmNews.CreateMessage msg.FromControlName, msg.MsgType + 1000, msg.Message
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.apmNews_MessageReceived", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub apmRTClient_MessageReceived(msg As gdOCX.gdAppMailMsg)
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean
    Static aLogs As New cGdArray
    Static d#, p&, dDiff#
    
    Dim strText$
    Static iLogging&, dLastCounted#, nCount&, nIndices&, nBidAsk&, nTrades&
    If iLogging = 0 Then
        If FileExist(App.Path & "\RtMsg.log") Then
            iLogging = 1
            FileFromString App.Path & "\RtMsg.log", "GenesisRT message counts ...", True, False
        Else
            iLogging = -1
        End If
    End If
    If iLogging > 0 Then
        If dLastCounted = 0 Then
            dLastCounted = gdTickCount
        ElseIf gdTickCount > dLastCounted + 15000 Then
            strText = "AppMails/sec = " & Str(Int(nCount / 15)) & ", Indices = " & Str(Int(nIndices / 15)) _
                & ", BidsAsks = " & Str(Int(nBidAsk / 15)) & ", Trades = " & Str(Int(nTrades / 15))
            FileFromString App.Path & "\RtMsg.log", strText, True, True
            nCount = 0
            nIndices = 0
            nBidAsk = 0
            nTrades = 0
            dLastCounted = gdTickCount
        End If
        nCount = nCount + 1
        Select Case msg.MsgType
        Case 20
            nTrades = nTrades + 1
        Case 25, 26
            nBidAsk = nBidAsk + 1
        End Select
        If Left(msg.Message, 1) = "$" Then
            nIndices = nIndices + 1
        End If
    End If
    
    If 0 Then 'IsIDE Then
        If Not bAlreadyDone Then
            bAlreadyDone = True
            KillFile App.Path & "\rtmsg.log"
            gdResetProfiles 400, 499
        End If
        
        If msg.MsgType > 0 And msg.MsgType <= 99 Then
            p = 400 + msg.MsgType
            gdStartProfile p
        Else
            p = 0
        End If
        If d = 0 Then d = gdTickCount
    End If
    
    g.RealTime.RTMessage msg
    
    If 0 Then 'IsIDE Then
        dDiff = d
        d = gdTickCount(False)
        dDiff = d - dDiff
        If p > 0 Then gdStopProfile p
    
        aLogs.Add "MSG: " & Str(msg.MsgNumber) & vbTab & Format(g.RealTime.FeedTime, "hh:mm:ss") & vbTab _
            & "#" & Str(msg.MsgType) & vbTab & Format(dDiff, "#0.00") & " ms" & vbTab _
            & Str(apmRTClient.InboxCount) & vbTab & Str(apmRTClient.OutboxCount)
            
        If aLogs.Size > 100 Then
            aLogs.Add gdGetProfiles(401, 499, ", ")
            aLogs.ToFile App.Path & "\rtmsg.log", True
            aLogs.Size = 0
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.apmRTClient_MessageReceived", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboBarPeriod_Click()
On Error GoTo ErrSection:

    If Not ActiveChart Is Nothing Then
        If Not ActiveChart.Chart Is Nothing Then
            If ActiveChart.Chart.TypeOfChart = eTypeChart_Seasonal Then
                InfBox "Please use Seasonal Sidebar on chart.", "I", , "Seasonal Chart"
            Else
                BarPeriodClick Me, m.oBtnMouseLast, False
                ActiveChart.Chart.ChangeBarPeriod cboBarPeriod.Text
                MoveFocus ActiveChart.pbChart
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMain.cboBarPeriod_Click"

End Sub

Private Sub cboBarPeriod_DropDown()
On Error GoTo ErrSection:

    If Not ActiveChart Is Nothing Then
        If ActiveChart.DetachStatus = eDetached Then ActiveChart.SkipFocusFix = True    '5322
    End If
    
    BarPeriodClick Me, m.oBtnMouseLast, True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMain.cboBarPeriod_Dropdown"

End Sub

Private Sub cboBarPeriod_GotFocus()
On Error GoTo ErrSection:

    If Not ActiveChart Is Nothing Then
        If ActiveChart.DetachStatus = eDetached Then
            ActiveChart.SkipFocusFix = True       '4883
            cboBarPeriod.SelLength = Len(cboBarPeriod.Text)     '5914
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMain.cboBarPeriod_GotFocus"

End Sub

Private Sub cboBarPeriod_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    If KeyCode = vbKeyReturn Then cboBarPeriod_Click        '5166

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMain.cboBarPeriod_KeyUp"

End Sub

Private Sub DockPro_Error(Message As String)
On Error GoTo ErrSection:

    Debug.Print Message

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.DockPro.Error", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Public Sub DockPro_ShortcutKeyDown(KeyCode As Integer, Shift As Integer, FormName As String)
On Error GoTo ErrSection:

    Dim strID$
    ' check for Ctrl-?
    If Shift = 2 Then
        Select Case KeyCode
            Case Asc("A")
                strID = "ID_Tile"
            Case Asc("G")
                strID = "ID_SymbolGrid"
            Case Asc("H")
                strID = "ID_HotKeys"
            Case Asc("I")
                strID = "ID_ChartOnOff"
            Case Asc("N")
                strID = "ID_Chart"
            Case Asc("O")
                strID = "ID_Chain"
            Case Asc("P")
                strID = "ID_Settings"
            Case Asc("Q")
                strID = "ID_Quote"
            Case Asc("R")
                strID = "ID_Replay"
            Case Asc("S")
                strID = "ID_Snapshot"
            Case Asc("T")
                strID = "ID_Toolbox"
            Case Asc("U")
                strID = "ID_Download"
            Case Asc("W")
                strID = "ID_ChartData"
            Case Asc("Z")
                strID = "ID_CustomizeToolbar"
            Case vbKeyF1
                strID = "ID_Test1"
            Case vbKeyF2
                strID = "ID_Test2"
            Case vbKeyF10
                strID = "ID_ImageServer"
            Case vbKeyPageUp
                strID = "-"
            Case vbKeyPageDown
                strID = "+"
        End Select
        If Len(strID) > 0 Then
            KeyCode = 0
            If strID = "+" Or strID = "-" Then
                LoadChartPage strID
            Else
                With tbToolbar
                    If .Tools(strID).Type = ssTypeStateButton Then
                        If .Tools(strID).State = ssUnchecked Then
                            .Tools(strID).State = ssChecked
                        Else
                            .Tools(strID).State = ssUnchecked
                        End If
                    End If
                    ToolBarClick .Tools(strID), Me
                End With
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.DockPro.ShortcutKeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub imgTbBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Set m.oBtnMouseLast = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_LBUTTONDOWN, Button, Index, X, Y, False)
    
End Sub

Private Sub imgTbBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Set m.oBtnMouseLast = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_LBUTTONUP, Button, Index, X, Y, False)

End Sub

Private Sub imgTbBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim oButton As cPicBoxButton

    Set oButton = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_MOUSEMOVE, Button, Index, X, Y, False)
    If Not oButton Is m.oBtnMouseLast Then
        ClearLastMouseButton Me, m.oBtnMouseLast
        Set m.oBtnMouseLast = oButton
        If Not m.oBtnMouseLast Is Nothing Then
            'Note: just setting the tooltiptext of the image control will result in the tooltip
            'always showing on the primary monitor so use our tooltip object instead (5128)
            m.oBtnMouseLast.BtnToolTipShow Me, pbTbBack(Index)
        End If
    End If

End Sub

Private Sub imgTbBackDraw_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Set m.oBtnMouseLast = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_LBUTTONDOWN, Button, Index, X, Y, True)

End Sub

Private Sub imgTbBackDraw_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    Set m.oBtnMouseLast = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_LBUTTONUP, Button, Index, X, Y, True)

End Sub

Private Sub imgTbBackDraw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim oButton As cPicBoxButton

    Set oButton = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_MOUSEMOVE, Button, Index, X, Y, True)
    If Not oButton Is m.oBtnMouseLast Then
        ClearLastMouseButton Me, m.oBtnMouseLast
        Set m.oBtnMouseLast = oButton
        If Not m.oBtnMouseLast Is Nothing Then
            m.oBtnMouseLast.BtnToolTipShow Me, pbTbBackDraw(Index)
        End If
    End If
        
End Sub

Private Sub MDIForm_Activate()
On Error GoTo ErrSection:

    Dim frm As Form, i&, s$
    Static bAlreadyDone As Boolean

    If Not bAlreadyDone Then
        bAlreadyDone = True
        
        'frmSymbolGrid.Show
        
        'tmrMain.Enabled = True
    End If
    
#If 0 Then
    ' Must do this since when non-child form loses focus,
    ' the FormX child activates but then immediately
    ' deactivates (must be a bug), so this activates it again.
    Set frm = MDIActiveForm
    If IsMDIChild(frm) Then
        'On Error Resume Next
        MoveFocus frm
    End If
#End If

    'UnloadEditors
    If FormIsLoaded("frmEditAnnot") Then Unload frmEditAnnot
    
    'JM 07-01-2009 - don't do this with new toolbar that can act on detached chart
    'If FormIsLoaded("frmTemplatePage") Then Unload frmTemplatePage
       
    ''AutoHideStatusForm

    'ToolbarSync Me, False '(just to reset the toolbar)
    
    g.dLastMouseActivity = gdTickCount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub MDIForm_DblClick()
On Error GoTo ErrSection:

    Dim i&, s$
    Dim frm As Form

    ' check for a "shortcut" action in the DblClick.flg file
    s = Trim(FileToString(App.Path & "\DblClick.flg", , True))
    Select Case UCase(Parse(s, vbTab, 1))
    
    Case "S" ' load this Strategy
        s = Parse(s, vbTab, 2)
        ' find ID for the Strategy Name
        i = SystemIDForName(s)
        If i > 0 Then
            If Not ActivateEditor("frmSystemManager", i) Then
                Set frm = New frmSystemManager
                frm.ShowMe i, , False
            End If
        End If
    
    Case "B" ' load this Basket
        s = Parse(s, vbTab, 2)
        ' find ID for the Basket Name
        i = BasketIDForName(s)
        If i > 0 Then
            If Not ActivateEditor("frmStrategyBasket", i) Then
                Set frm = New frmStrategyBasket
                frm.ShowMe i
            End If
        End If
    
    Case "TEST3"
        Set frm = New frmTest3
        frm.ShowMe
    
    End Select
    
ErrExit:
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub MDIForm_Deactivate()
On Error GoTo ErrSection:

    Dim frm As Form
'    Set frm = Me.ActiveForm
'    Set frm = ActiveChart
'    If Not frm Is Nothing Then
'        If IsFrmChart(frm) Then
'            ' if in the middle of drawing a new annotation, then delete it
'            If Len(g.strActiveDraw) > 0 Then
'                frm.ClearAnnotFlags True
'            Else
'                ' otherwise just clear the annotation flags
'                frm.ClearAnnotFlags False
'            End If
'        End If
'        Set frm = Nothing
'    End If

    g.dLastMouseActivity = gdTickCount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form.Deactivate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub MDIForm_Load()
On Error GoTo ErrSection:

    Dim i&, strText$

    i = GetIniFileProperty("LargeButtons", False, "Toolbars", g.strIniFile)
    pbTbBack(0).BorderStyle = 0
        
'    Set g.RealTime = New cRealTime
    cboBarPeriod.Text = ""
    
    ' don't auto-show child forms when they're first loaded
    Me.AutoShowChildren = False
    Me.Icon = Picture16(ToolbarIcon("ID_About"))
    
    ' set main form's icon and caption
    SetMainCaption
            
    ' get "normal" size and location
    m.strNormalPlacement = GetIniFileProperty("MDI_Placement", "", "Forms", g.strIniFile)
    If Len(m.strNormalPlacement) > 0 Then
        SetFormPlacement Me, m.strNormalPlacement
    Else
        Me.Height = Screen.Height * 0.9
        Me.Width = Screen.Width * 0.9
        CenterTheForm Me
        m.strNormalPlacement = GetFormPlacement(Me)
    End If
    ' (but put out-of-sight for now)
    m.nInitialTop = Me.Top
    If Is9598orMe Then
        Me.Top = -Me.Height - 10000
    Else
        Me.Top = -16000 * Screen.TwipsPerPixelY '(pretty close to the max negative value allowed)
    End If
                        
    'tbToolbar.SaveConfiguration App.Path & "\Toolbar.cfg"

'    ToolbarReset

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    'check if mouse is still over last highlighted toolbar button
    CheckHighlightedToolButton

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Dim i&, strText$, nDate&
    Dim dTimeout As Date

    ' if app being shutdown because of Windows shutting down, skip the MDB compact
    ' (since a partially finished compact may be a cause of MDB corruption)
    If UnloadMode > 1 Then
        g.bSkipMdbCompact = True
    ElseIf ProcessIsBusy Then
        ' otherwise if a process is busy, then don't shutdown yet
        ' (e.g. we want to wait until after fully initialized)
        Cancel = True
        Me.tbToolbar.Redraw = True ' TLB: need to set this back on if exit is cancelled
        Exit Sub
    End If
        
    ' prompt to run Archive or allow user to cancel the shutdown
    If Not g.bSkipMdbCompact And Not m.bReadyToUnload And Len(g.strRunWhenExit) = 0 And FileExist(App.Path & "\TNArchive.exe") Then
        i = GetIniFileProperty("ArchivePrompt", 0, "General", g.strIniFile)
        ' TLB 8/15/2012: but if on an NVS server, then always prompt upon exiting
        If FileExist(App.Path & "\..\Restart.bat") Or HasModule("NVS*") Then
            i = 0 ' to force the prompt regardless of the setting
        End If
        If i >= 0 Then
            ' see if date of last archive is at least "X" days ago
            nDate = 0
            If i > 0 Then
                strText = GetIniFileProperty("LastBackupFile", "", "Main", App.Path & "\TNArchive.INI")
                If Len(strText) > 0 Then
                    nDate = Int(FileDate(App.Path & "\Archive\" & strText)) + i
                End If
            End If
            If Date >= nDate Then
                If WindowState = vbMinimized Then
                    If m.iPrevNonminimizedState = vbMinimized Then ' (shouldn't be true, but just in case)
                        WindowState = 0
                    Else
                        WindowState = m.iPrevNonminimizedState
                    End If
                End If
                strText = "It is recommended that you use the Archive program periodically to backup your work|(all settings, charts, library items, etc.) ...|"
                Select Case InfBox(strText, "?", "+EXIT Now|Archive|-Cancel", "Exit " & g.strTitle)
                Case "C"
                    Cancel = True
                    Me.tbToolbar.Redraw = True ' TLB: need to set this back on if exit is cancelled
                    Exit Sub
                Case "A"
                    g.strRunWhenExit = Chr(34) & App.Path & "\TNArchive.exe" & Chr(34) '& " " & Chr(34) & g.strTitle & Chr(34)
                End Select
            End If
        End If
    End If

    If Not g.Broker.DisconnectFromAll("TradeNav shutting down: " & Str(UnloadMode), True) Then
    'ElseIf frmTTSummary.HasActiveAutoTradeItems Then
        Cancel = True
        'InfBox "You cannot shut down Trade Navigator until you disable all auto trading items", "!", , "Shut Down Error"
    ElseIf Not m.bReadyToUnload Then
        On Error Resume Next
        
        ' set cancel flag now -- will let timer call unload again when ready flag is set
        Cancel = True
        
        ' Walk through all of the Editors and attempt to close them...
        For i = Forms.Count - 1 To 0 Step -1
            If IsEditor(Forms(i).Name) Then
                ' if editor is not visible, it just didn't get
                ' unloaded all the way so kill it now
                If Not Forms(i).Visible Then
                    Unload Forms(i)
                ElseIf Forms(i).AskToSave Then
                    Exit Sub
                Else
                    Unload Forms(i)
                End If
            End If
        Next i
    
        tmrMain.Enabled = False ' so won't start anything else yet
        tmrWindowLink.Enabled = False
        tmrAutoResize.Enabled = False
        StatusMsg "*** UNLOADING ***", vbRed
        InfBox
        Me.tbToolbar.Redraw = False
        g.bUnloading = True
        Screen.MousePointer = vbHourglass
        'InfBox "w=NOWAIT ; Unloading ..."
        'SetFormTopmost frmAsk, True
        DoEvents
        
        ' save defaults to Ini file
        SetIniFileProperty "MDI_Placement", m.strNormalPlacement, "Forms", g.strIniFile
        If WindowState <> vbMinimized Then
            i = WindowState
        ElseIf m.iPrevNonminimizedState <> vbMinimized Then
            i = m.iPrevNonminimizedState
        Else
            i = 0
        End If
        SetIniFileProperty "MDI_State", i, "Forms", g.strIniFile
        With tbToolbar
            .Redraw = False
            ToolbarSavePositions
            
            ' Save cursor type
            strText = "ID_CursorArrow"
            If .Tools(strText).State <> ssChecked Then
                strText = "ID_CursorVertLine"
                If .Tools(strText).State <> ssChecked Then
                    strText = "ID_CursorHorizLine"
                    If .Tools(strText).State <> ssChecked Then
                        strText = "ID_CursorCrosshairs"
                    End If
                End If
            End If
            SetIniFileProperty "CursorType", strText, "Charting", g.strIniFile
                
            SetIniFileProperty "RealTime", .Tools("ID_RealTime").State, "Toolbars", g.strIniFile
        End With
        
        ' so entire app disappears all at once
        LockWindowUpdate Me.hWnd
        
        ' make sure test forms are unloaded
        If FormIsLoaded("frmTest") Then Unload frmTest
        If FormIsLoaded("frmTest2") Then Unload frmTest2
                
        ' save current charts
        SaveCharts
        
        ' save other visible forms
        SaveVisibleForms
    
        ' Inactivate real-time hookup
        If g.RealTime.Active Then
            g.RealTime.Init False, "TradeNav Unloading"
        End If
        
        ' set flag so timer will unload app after RT inactive
        m.bReadyToUnload = True
        tmrMain.Enabled = True
            
        ' close down any WebReport forms that might still be open
        For i = 1 To g.nLastWebReportID
            apmNews.CreateMessage "TNWebReport" & Str(i), 1, "", , True
        Next
            
#If 0 Then
            ' wait till not busy, or hit 5-second timeout
            dTimeout = DateAdd("s", 5, Now)
            Do While frmQuotes.IsBusy
                If Now > dTimeout Then
                    DebugLog "QueryUnload timed out while waiting for frmQuotes to be free"
                    Exit Do
                End If
                Sleep 0.25
            Loop
            Sleep 0.25
        End If
#End If
    
    End If
    
    If Not Cancel Then
        apmRTClient.Unload
        apmNews.Unload
        tmrQuickStart.Enabled = False
        tmrPredLabs.Enabled = False
        ChartTimers = False
        frmTTSummary.DisableTimers
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next

    Dim ws%, i&
    Static bInProgress As Boolean, bAlreadyDone As Boolean
         
    If bInProgress Or g.bUnloading Then Exit Sub
    bInProgress = True

    ws = WindowState
    If ws = 0 Then
        If m.bShouldBeSeen Then
            ' make sure can be seen (in screen "range")
            If Me.Top < 0 And m.nInitialTop > 0 Then
                Me.Top = m.nInitialTop
            End If
            m.nInitialTop = 0
            If Not bAlreadyDone Then
                bAlreadyDone = True
                MoveFormOnScreen Me
            End If
            m.strNormalPlacement = GetFormPlacement(Me)
        End If
    End If
        
    ' relocate "icons" for minimized charts
    If ws <> vbMinimized Then
        Me.Arrange vbArrangeIcons
        m.iPrevNonminimizedState = ws
    End If
    
    bInProgress = False

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim i&
    
    ChartTimers = False
    frmTTSummary.DisableTimers
    tmrMain.Enabled = False
    'g.RealTime.Init False
    DoEvents

    CleanupWhenExit
    'apmRTClient.Unload 'TLB 12/1/05: Dave ran into an anomaly that indicates we may not want to do this here
    Screen.MousePointer = 0
    
    'notify grapheng to release gdiplus objects for toolbar (must do last)
    geShutdownAll
    
    ' only create the flag if there have not been any errors raised
    If RaiseError("", eGDRaiseError_HasHadErrors) = False Then
        FileFromString App.Path & "\SkipReg.flg", "Normal shutdown, so no need to re-run RegFiles"
    End If
    
ErrExit:
    'End  'TLB: DO NOT DO THIS!!! (causes app to sometimes crash when shutting down)
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub tbToolbar_CtlClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    ToolBarClick Tool, Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.tbToolbar.CtlClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub tbToolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
'JM: 02-17-2009. A part of this subroutine is copied to frmChart. Make sure to check code in frmChart when changing this code.

    Dim bCustomizeToolbar As Boolean
    Dim Tool As SSTool
    Dim Toolbar As SSToolBar
    Dim frm As Form
    Dim Chart As cChart
    Dim Annot As cAnnotation
    Dim eType As eAnnotType
    Dim bCallClick As Boolean
                
    ' if right click on a menu item, let's try to do the action while leaving the menu up
    ' (e.g. for "more bars", "less bars")
    If Button = 2 Then
        ' if toolbar is nothing, then the tool should be from the menu toolbar
        Set Toolbar = tbToolbar.ToolBarFromPosition(X, Y)
        If Toolbar Is Nothing Then
            Set Tool = tbToolbar.ToolFromPosition(X, Y)
            If Not Tool Is Nothing Then
                If Tool.Type = ssTypeButton Or Tool.Type = ssTypeStateButton Then
                    ToolBarClick Tool, Me
                End If
            Else
                bCustomizeToolbar = True
            End If
        ElseIf Toolbar.Name = kTbDraw Then
            Set Tool = tbToolbar.ToolFromPosition(X, Y)
            If Not Tool Is Nothing Then
                Set Annot = New cAnnotation
                eType = Annot.AnnotTypeFromToolID(Tool.ID)
                If eType = eANNOT_UndefinedType Then
                    ' if the button is a state button that is not in a group and is currently down,
                    ' then toggle it back up
                    If Len(Trim(Tool.Group)) = 0 Then
                        If Tool.State = ssChecked Then
                            Tool.State = ssUnchecked
                        End If
                    End If
                    'If IsIDE Then StatusMsg Tool.ID
                Else
                    Set frm = ActiveChart
                    If Not frm Is Nothing Then
                        Set Chart = frm.Chart
                        If Not Chart Is Nothing Then
                            Chart.RemoveAnnots True, eType
                            Chart.GenerateChart eRedo1_Scrolled
                        End If
                    End If
                End If
            End If
        Else
            bCustomizeToolbar = True
        End If
    ElseIf FormIsLoaded("frmTemplatePage") Then
        ' if already loaded, see if we are switching modes (between templates and pages)
        Set Tool = tbToolbar.ToolFromPosition(X, Y)
        If Not Tool Is Nothing Then
            If Tool.ID = "ID_Pages" Then
                If frmTemplatePage.FormMode <> eMode_Pages Then
                    bCallClick = True
                End If
            ElseIf Tool.ID = "ID_Templates" Then
                If frmTemplatePage.FormMode <> eMode_Templates Then
                    bCallClick = True
                End If
            End If
        End If
        ' call DoEvents to allow the LostFocus to trigger and unload the form,
        ' but doublecheck and unload form if still loaded
        DoEvents
        If FormIsLoaded("frmTemplatePage") Then
            Unload frmTemplatePage
        End If
        If bCallClick Then
            ' to now reload and switch to other mode
            Tool.State = ssChecked
        End If
    ElseIf Not ActiveChart Is Nothing Then
        If ActiveChart.DetachStatus = eDetached Then
            ActiveChart.SkipFocusFix = True         '4883 - allows menu dropdown to work
        End If
    End If
    Set Tool = Nothing
    Set Toolbar = Nothing
    Set frm = Nothing
    Set Chart = Nothing
    Set Annot = Nothing
    
    ' customize toolbar if right-click on anything but the drawing toolbar
    If bCustomizeToolbar Then
        frmToolbar.ShowMe
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.tbToolbar.MouseUp", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub tbToolbar_OnCustomize(ByVal Cancel As ActiveToolBars.SSReturnBoolean)
    
    ' bring up our customization form instead
    Cancel = True
    frmToolbar.ShowMe

End Sub

Private Sub tbToolbar_ToolBarModified(ByVal change As ActiveToolBars.Constants_Modified, ByVal Toolbar As ActiveToolBars.SSToolBar, ByVal Tool As ActiveToolBars.SSTool)

    Dim Toolbar2 As SSToolBar

    On Error Resume Next
    If Toolbar.Style = ssStandard Then
        Select Case change
        ' make sure a toolbar doesn't get docked above the menu
        Case ssToolBarEndDrag
            If Toolbar.DockedStatus = ssDockedTop And Toolbar.DockedRow = 1 Then
                tbToolbar.ToolBars("Menu").DockedRow = 1
            End If
        
        ' set order when open new toolbar at top
        Case ssToolBarOpened
            If Toolbar.DockedStatus = ssDockedTop And Toolbar.Style <> ssMenuBar Then
                For Each Toolbar2 In tbToolbar.ToolBars
                    With Toolbar2
                        If .Style = ssMenuBar Then
                            .DockedRow = 1
                        ElseIf .DockedStatus = ssDockedTop And .Visible Then
                            .DockedRow = 2
                            Select Case .Name
                            Case "General"
                                .DockedColumn = 1
                            Case kTbWindows
                                .DockedColumn = 2
                            Case kTbChartSettings
                                .DockedColumn = 3
                            Case kTbDraw
                                .DockedColumn = 4
                            End Select
                        End If
                    End With
                Next
            End If
        End Select
    End If

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    ToolBarClick Tool, Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub tbToolbar_ToolDropDown(ByVal Tool As ActiveToolBars.SSTool, ByVal ScreenX As Single, ByVal ScreenY As Single)
On Error GoTo ErrSection:

    ToolBarClick Tool, Me, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.tbToolbar.ToolDropDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub tbToolbar_ToolKeyDown(ByVal Tool As ActiveToolBars.SSTool, ByVal KeyCode As Integer, ByVal Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = 13 Then
        If Tool.Type = ssTypeComboBox Or Tool.Type = ssTypeEdit Then
            ToolBarClick Tool, Me
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.tbToolbar.ToolKeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub tmrAutoResize_Timer()
On Error Resume Next

    Dim i&, iRows&, iHeight&
    Static nPrevWidth&, nPrevHeight&, bInProgress As Boolean
    
    TimerStart "frmMain.tmrAutoResize"
    If Not bInProgress And Not g.bUnloading And Not g.bLoadingChartPage Then
        bInProgress = True
        
        ' store current size and location (since form could have
        ' been moved and the resize event doesn't trigger for that)
        If Me.WindowState = 0 Then
            m.strNormalPlacement = GetFormPlacement(Me)
        End If
    
        If Not ActiveChart Is Nothing Then
            ' check if MDI client size has changed (e.g. if a form just docked or undocked)
            Do While Me.ScaleWidth <> nPrevWidth Or Me.ScaleHeight <> nPrevHeight
                tmrAutoResize.Enabled = False ' disable now so each chart will not yet store new ratios
                nPrevWidth = Me.ScaleWidth
                nPrevHeight = Me.ScaleHeight
                If ActiveChart.WindowState = 0 Then
                    For i = 0 To Forms.Count - 1
                        If IsFrmChartMDI(Forms(i)) Then
                            If Forms(i).DetachStatus = eNotDetached Then
                                ' reposition the chart form to it's current ratios
                                Forms(i).SetRatioPlacement Forms(i).GetRatioPlacement
                            End If
                        End If
                    Next
                    ' arrange minimized chart icons
                    Me.Arrange vbArrangeIcons
                End If
                If Not FormIsLoaded("frmToolbar") Then      '5191
                    'resize tool bar
                    m.nBtnsPerRow = ToolbarResize2(Me, pbTbBack, imgTbBack, m.aTbButtons, m.bToolbarWrap)
                    ToolbarResize2 Me, pbTbBackDraw, imgTbBackDraw, m.aTbButtonsDraw, m.bToolbarWrap
                    Me.cboBarPeriod.SelLength = 0
                End If
                 ' wait and loop back to see if still resizing
                DoEvents
            Loop
            If Not tmrAutoResize.Enabled And Not g.bUnloading Then
                tmrAutoResize.Enabled = True ' reenable now so each chart can store ratios
            End If
        End If
        
        bInProgress = False
    End If
    TimerEnd "frmMain.tmrAutoResize", tmrAutoResize.Interval

End Sub

Private Sub tmrCheckBuySellButtons_Timer()
On Error GoTo ErrSection:

    Dim rc&, iHwnd&, X&, Y&
    
    TimerStart "frmMain.tmrCheckBuySellButtons"
    iHwnd = ValOfText(Me.tmrCheckBuySellButtons.Tag)
    If iHwnd = 0 Then
        ClearAllBuySellBtns
    ElseIf geIsCursorInWnd(iHwnd, X, Y) <> 1 Then
        ClearAllBuySellBtns
    End If
    TimerEnd "frmMain.tmrCheckBuySellButtons", tmrCheckBuySellButtons.Interval
        
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.tmrCheckBuySellButtons_Timer"

End Sub

Private Sub tmrGridScrollPressed_Timer()
On Error Resume Next

    TimerStart "frmMain.tmrGridScrollPressed"
    If Not MouseIsPressed Then
        tmrGridScrollPressed.Enabled = False
        GridScrollCheck Nothing, 0, 0, 0, 0, False
    End If
    TimerEnd "frmMain.tmrGridScrollPressed", tmrGridScrollPressed.Interval

End Sub

Private Sub tmrMain_Timer()
On Error GoTo ErrSection:

TimerStart "frmMain.tmrMain"

    Dim i&, n&, s$, d#, strTag$, bFirstTime As Boolean
    Dim bShowMsg As Boolean, bNeedDailyDownload As Boolean, bStartRealtime As Boolean, bNewUser As Boolean
    Dim strKey As String, strTemp$
    Dim lValue As Long
    Dim strReturn As String             ' Return from in an infbox
    Dim astrFile As New cGdArray        ' New Way file for configuration
    Dim lStartRange As Long             ' Starting number of days
    Dim lEndRange As Long               ' Ending number of days
    
    Static bTradeSenseAlreadyKilled As Boolean
    Static bInProgress As Boolean
    Static dLastAlertCheck As Double
    Static dLastHeapCompact As Double
    Static dNextRealTimeReconnect As Double
    Static dStreamRecordStart#, dStreamRecordEnd#, strStreamRecordMask$
    Static dLastAutoExport As Double
       
    If g.bUnloading Then
        ' check if waiting for RT to disable before shutting down
        If m.bReadyToUnload And Not bInProgress Then
            If Not g.RealTime.Active Then
                ' now we can call unload for good
                tmrMain.Enabled = False
                DoEvents
                Unload Me
                DoEvents
                End
            End If
        End If
        Exit Sub
    End If
        
    'check if mouse is still over last highlighted toolbar button
    CheckHighlightedToolButton False

    strTag = tmrMain.Tag
    If Len(strTag) > 0 Then
        tmrMain.Tag = ""
        If strTag = "QUIT" Then
            Unload Me
            Exit Sub
        ElseIf Left(strTag, 5) = "SLEEP" Then
            Sleep Val(Parse(strTag, " ", 2))
            Exit Sub
        ElseIf Left(strTag, 9) = "ACTIVATE " Then
            n = Val(Parse(strTag, " ", 2))
            For i = 0 To Forms.Count - 1
                If Forms(i).hWnd = n And n <> 0 Then
                    LockWindowUpdate frmMain.hWnd
                    MoveFocus Forms(i)
                    ActiveChartFormSet Forms(i)     '6029
                    'this tag is set from mMain when user chooses a chart from drop down menu
                    'if the chosen chart is minimized then set flag for chart's timer will restore it (to be nice)
                    If Forms(i).WindowState = vbMinimized Then Forms(i).tmr.Tag = "RESTORE_NOW"
                    DoEvents
                    Exit For
                End If
            Next
            LockWindowUpdate 0
            Exit Sub
        ElseIf UCase(strTag) = "UNLOCKWINDOWUPDATE" Then
            LockWindowUpdate 0
            Exit Sub
        ElseIf UCase(strTag) = "DOWNLOADHELP" Or UCase(strTag) = "UPGRADE" Then
            If Me.Enabled And Not ProcessIsBusy(True) Then
                If FormIsLoaded("frmDownload") Then Unload frmDownload
                Set MsgForm = frmStatus
                frmDownload.optSpecialFile = True
                If UCase(strTag) = "UPGRADE" Then
                    frmDownload.txtSpecialFile = "Upgrade"
                Else
                    frmDownload.txtSpecialFile = "Help"
                End If
                frmDownload.DownloadData
                Set MsgForm = Nothing
                Exit Sub
            End If
        End If
    End If
    
    If bInProgress Then Exit Sub
    bInProgress = True

    ' try compacting the memory heap once every 10 minutes
    ' (not supported for 95, 98 or ME)
    If Not Is9598orMe Then
        If gdTickCount > dLastHeapCompact + 600000 Then
            dLastHeapCompact = gdTickCount
            i = HeapCompact(GetProcessHeap, 0)
            DebugLog Format(Date, "YYYYMMDD") & ", HeapCompact = " & Format(i, "#,##0") + g.RealTime.DumpTickBufferInfo
        End If
    End If

    If Not g.RealTime.Active Then
        dNextRealTimeReconnect = 0 ' clear it when realtime inactive
    ElseIf dNextRealTimeReconnect = 0 Then
        ' calculate next time to do a weekly reconnect for realtime (random time on Sunday morning in NY)
        ' - this keeps the TedNEd's from getting too huge
        ' - and it forces a reset after any DST changes have happened around the world
        dNextRealTimeReconnect = Int(ConvertTimeZone(Now, "", "NY")) ' date in NY right now
        Do ' go to the next Sunday (next week if today is already Sunday in NY)
            dNextRealTimeReconnect = dNextRealTimeReconnect + 1
        Loop While Weekday(dNextRealTimeReconnect) <> vbSunday
        dNextRealTimeReconnect = dNextRealTimeReconnect + RandomNum(22200, 41400) / 86400# ' add random partial day (between 6:10am-11:30am NY)
        dNextRealTimeReconnect = ConvertTimeZone(dNextRealTimeReconnect, "NY", "") ' convert back to local time zone
    End If
    
If 0 Then
    'dNextRealTimeReconnect = Now + 15 / 86400#
End If

    If g.bStarting Then
        ' First time:
        bFirstTime = True
        DoEvents
        bShowMsg = True
        If FormIsLoaded("frmSplash") Then
            Unload frmSplash
        End If
        If FormIsLoaded("frmTTSummary") Then frmTTSummary.Form_Resize
        StartupLog "------ Timer -------"
ChartTimers = True
        
        ' check if auto-export
        If FileExist(App.Path & "\AutoExport.flg") Then
            dLastAutoExport = 1
        End If
        
        ' check if stream recording
        s = Trim(FileToString(App.Path & "\RTS\Record.flg", , True))
        If Len(s) > 0 Then
            i = Val(Parse(s, vbTab, 1))
            i = Int(i / 100) * 60 + (i Mod 100)
            dStreamRecordStart = i / 1440#
            i = Val(Parse(s, vbTab, 2))
            i = Int(i / 100) * 60 + (i Mod 100)
            dStreamRecordEnd = i / 1440#
            strStreamRecordMask = Parse(s, vbTab, 3)
            If Len(strStreamRecordMask) > 0 And Right(strStreamRecordMask, 1) <> "*" Then
                strStreamRecordMask = strStreamRecordMask & "*"
            End If
        End If
        
        ' CUSTOMER INFO -- AUTO-SUBSCRIBE
        strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
        ' TLB 12/7/06: try checking last known Cust ID (from last successful connection)
        ' instead of what they last typed in -- this will allow for a mistyped Cust ID
        ' or password to get corrected the next time TradeNav gets started
        If g.lLCD = 0 Then 'If NewUser Then
            ' create new customer account
            bNewUser = True
            If frmNewAccount.ShowMe = False Then
                g.bStarting = False
                bInProgress = False
                Unload Me
                Exit Sub
            End If
        End If
        ' DATA INSTALL
        If g.SymbolPool.NumRecords = 0 And Not NewUser Then
            If InstallData = False Then
                g.bStarting = False
                bInProgress = False
                Unload Me
                Exit Sub
            End If
        End If
                
        ' if we don't have the true enablement codes (i.e. still the default), try to get them now
        AskForActivate
        
        ' check for a newer file of News, Msg, WhatsNew, etc.
        ' (in case just after an upgrade)
        CheckForSpecialDownloadFiles App.Path & "\ftp"
        
        ' do ETA charts before real-time hookup
        CheckForDoItNow
        CheckForGraphNow '(temp. leave for backward-compat)
        
'frmStartConnect.ShowMe
               
        ' check if need to do a daily download
        bNeedDailyDownload = NeedDailyUpdate
        
        ' create SimTrade account (and broker accounts?)
        SetupInitialAccounts
        'If Not bNeedDailyDownload Then
            SetupBrokerLayout
        'End If
        
        ' restore price ladders, etc.
        RestoreVisibleForms
        
        ' check if want to start real-time
        StartupLog "RT and QB initialization"
        If (HasModule("RTG") Or GetIniFileProperty("RealTime", 0, "Toolbars", g.strIniFile) <> 0) Then
            If Not bNeedDailyDownload And Not FileExist(App.Path & "\ImageServer.flg") Then
                If Not g.FtpDownloader.DownloaderIsRunning Then
                    If InfBox("Do you wish to connect now to the |Data Streaming server?", "?", "+Streaming|-Not Now", "Data Streaming Connection") = "S" Then
                        bStartRealtime = True
                    End If
                End If
            End If
        End If
        g.RealTime.Init bStartRealtime
        StartupLog "RT and QB initialized"
        
        ' need this AFTER first real-time init,
        ' but BEFORE things below ...
        g.bStarting = False
        
        ' let's reset the toolbar in case enablements have changed during data install, etc
        ' TLB/MJM 3/10/2015: for a new user, call ToolbarReset with bReset = True so will forces changes to displayed toolbars
        ToolbarReset bNewUser
        
        ' update all charts now (will now run strategies on charts)
        UpdateVisibleCharts eRedo1_Scrolled  'eRedo5_RecalcInd
        
        'JM 08-18-2009: for some reason if active chart is a maximized spread chart, the tabs do not restore correctly at startup
        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
            If g.ChartGlobals.frmActiveNonDetached.WindowState = vbMaximized Then
                If Len(g.ChartGlobals.frmActiveNonDetached.Chart.SpreadSymbols) > 0 Then
                    LockWindowUpdate frmMain.hWnd
                    g.ChartGlobals.frmActiveNonDetached.WindowState = vbNormal
                    g.ChartGlobals.frmActiveNonDetached.WindowState = vbMaximized
                    LockWindowUpdate 0
                End If
            End If
        End If

        ' Warn of expiring modules if applicable
        ''ExpiringModuleWarning
        ExpiringDataPkgWarning True

        ' See if need to recalc criteria
        If g.SymbolPool.NumRecords > 0 And Not bNeedDailyDownload And Not bStartRealtime Then
            DoEvents
            If Not WindowStateX(Me) = wsMinimized Then
                CheckCriteria True
            End If
        End If
        ' calc times of next auto-downloads
        CalcNextTryTime
        
        CalcNextQuoteRefresh
        
        InfBox '(to clear no-wait msg)
        StatusMsg "To see charting Hot-keys and Tips, hit 'Ctrl-H'"
        
        ' if running Image Server
        If FileExist(App.Path & "\ImageServer.flg") Then
            frmImageServer.ShowMe True
        End If
        
        ' ask to do a daily download (if has had a valid connection)
        If bNeedDailyDownload And (RI_GetLastDataServiceID > 0) Then
            If Me.Enabled And Not ProcessIsBusy(True) Then
                If InfBox("Daily Update files should be ready for download.||Would you like to download them now?", "?", "+Update|-Not Now", "Daily Update") <> "N" Then
                    g.dNextDownloadTry = Now
                End If
            End If
        End If
        
        apmNews.Active = True
    End If '(end of if g.bStarting)

    ' check if ETA charts to-do
    CheckForDoItNow
    
    CheckForGraphNow '(temp. leave for backward-compatibility)
    
    ' if in a tradesense editor ...
    If Not g.ActiveEditor Is Nothing Then
        On Error Resume Next
        ' check to see if user went to a different app
        If GetActiveWindow() = 0 Then
            ' if so, then kill any "on top" tradesense windows
            If Not bTradeSenseAlreadyKilled Then
                g.ActiveEditor.RemoveTradeSense
                bTradeSenseAlreadyKilled = True ' (don't need to keep killing)
            End If
        Else
            bTradeSenseAlreadyKilled = False
        End If
    End If
    
    If Not bFirstTime And Not g.bUnloading Then
    
        ' load chart page here instead of from chart form
        ' (since chart form may need to be unloaded during page load)
        If strTag = "LoadChartPage +" Then
            LoadChartPage "+"
        ElseIf strTag = "LoadChartPage -" Then
            LoadChartPage "-"
        ElseIf strTag = "UpdateVisibleCharts" Then
            UpdateVisibleCharts -1
        End If
        
        ' show alert messages form if an alert happened while a modal form had been up
        If g.bShowAlertMsgForm Then
            If Me.Enabled Then
                g.bShowAlertMsgForm = False
                If Not frmAlertMessages.Visible Then
                    frmAlertMessages.ShowMe        '5878
                End If
            End If
        End If
            
        ' check if should start an auto Daily Download
        If Now >= g.dNextDownloadTry And g.dNextDownloadTry > 0 And g.nReplaySession = 0 Then
            If Me.Enabled And Not ProcessIsBusy(True) Then
                ' for a daily download, force a reload of the form
                If FormIsLoaded("frmDownload") Then Unload frmDownload
                Set MsgForm = frmStatus
                frmDownload.optDaily = True
                frmDownload.DownloadData
                Set MsgForm = Nothing
            End If
        End If
        
        ' check if RealTime server has just activated
        If strTag = "REALTIME" Then
            g.RealTime.RefreshSymbolList True
            ' and check again if need to do a daily download (don't ask this time)
            If NeedDailyUpdate And (g.nReplaySession = 0) Then
                g.dNextDownloadTry = Now
            End If
        ' else if realtime needs to be reconnected
        ElseIf strTag = "RECONNECT" Then
            g.RealTime.Reconnect
            If tmrMain.Tag = "RECONNECT" Then tmrMain.Tag = ""
        ' else if realtime is active, see if any symbols have been added to the list
        ElseIf g.RealTime.Active And gdTickCount > m.dWaitToCheckRT Then
            If dNextRealTimeReconnect > 0 And Now > dNextRealTimeReconnect Then
                If Me.Enabled And Not ProcessIsBusy(True) Then
                    ' also make sure they're not actively working with the program right now
                    If (gdTickCount - g.dLastMouseActivity) / 60000# > 1 Then
                        dNextRealTimeReconnect = 0 '(clear so will recalculate again after reconnecting)
                        g.RealTime.Reconnect 0, True ' (also compact the TradeTracker during the Sun morning reconnect)
                    End If
                End If
            ElseIf g.RealTime.IsServerActive(True) Then
                ' (unless the focus is on the symbol grid -- so won't keep requesting data
                '  while changing symbols by moving up and down the symbol grid)
                If Not Screen.ActiveForm Is frmSymbolGrid Then
                    g.RealTime.RefreshSymbolList
                End If
            End If
        End If
    
        ' check for auto qb-refresh
        If Now >= g.dNextQuoteBoardRefresh And g.dNextQuoteBoardRefresh > 0 Then
            If Me.Enabled And Not ProcessIsBusy(True) Then
                Set MsgForm = frmStatus
                If g.RealTime.Active Then
                    g.dLastQuoteBoardRefresh = Now
                    g.RealTime.RefreshSymbolList 2
                Else
                    g.RealTime.RefreshSymbolList True
                End If
                Set MsgForm = Nothing
                CalcNextQuoteRefresh
                
                ' for debugging: to store off QB data
                If FileExist(App.Path & "\QB-data\*.*") Then
                    strReturn = App.Path & "\ftp\reqdata.gzp"
                    If FileExist(strReturn) Then
                        On Error Resume Next
                        FileCopy strReturn, App.Path & "\QB-data\" & Format(Now, "yyyymmddhhnn.gzp")
                    End If
                End If
            End If
        End If
            
        If Not g.bStarting And Not g.bUnloading Then
            ' check alerts every 15 seconds in case there are time alerts...
            If gdTickCount > dLastAlertCheck + 15000 Then
                g.Alerts.CheckAlerts False
                dLastAlertCheck = gdTickCount + 15000
                ' and remove any obsolete tick buffers
                If g.RealTime.Active Then
                    g.RealTime.RemoveObsoleteTickBuffers
                End If
                
                ' make sure TradeNavStartup is still running
                If IsAtLeastVista And Not IsIDE Then
                    If KillProcess("TradeNavStartup", True) = 0 Then
                        If FileExist(App.Path & "\TradeNavStartup.exe") Then
                            ' (but don't need to restart TradeNav in this case)
                            FileFromString App.Path & "\TradeNavStartup.Run", "c:\nonexistant-program-1234.exe", True
                            ShellExecute ByVal 0&, "open", App.Path & "\TradeNavStartup.exe", "", App.Path, 0&
                            DebugLog "Re-started TradeNavStartup.exe"
                        End If
                    End If
                End If
                frmQuotes.WebPageCheck
                FtpUploadCheck
                
                ' this is just for testing the weekly auto-reconnect feature
                If g.RealTime.Active And dNextRealTimeReconnect > 0 Then
                    s = App.Path & "\Reconnect.Now"
                    If FileExist(s) Then
                        KillFile s
                        g.dLastMouseActivity = Date - 1
                        dNextRealTimeReconnect = Now
                    End If
                End If
            End If
            
            ' auto-export (custom feature for Bollinger: re-export each minute)
            If dLastAutoExport > 0 And g.RealTime.Active Then
                d = Int(g.RealTime.FeedTime * 1440)
                If d > dLastAutoExport Then
                    dLastAutoExport = d
                    ExportData True
                End If
            End If
            
            ' check if an FTP Data Download is ready to install
            FtpInstallCheck
            
            CheckForTradeNavMessages
        End If
    
        ' check for auto stream-recording (for the recording machine, not for customers)
        If dStreamRecordEnd > 0 Then
            If Time > dStreamRecordEnd Then
                If g.RealTime.Active Then
                    g.RealTime.Init False, "StreamRecord Ending"
                End If
            ElseIf Time > dStreamRecordStart Then
                If Not g.RealTime.Active Then
                    If IsWeekday(Date) Then
                        g.RealTime.Init True
                    End If
                ' for first 15 minutes, keep trying to re-connect to a specific site (if mask exists)
                ElseIf Time < dStreamRecordStart + 15 / 1440# And g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
                    s = g.RealTime.SplitterIP
                    If Len(s) > 0 And Len(strStreamRecordMask) > 0 Then
                        If Not s Like strStreamRecordMask Then
                            g.RealTime.Init False, "Not the preferred data site"
                        End If
                    End If
                End If
            End If
        End If
    
        LoadAppBkImage
    End If
    
    If IsIDE Then
        'StatusMsg Str(Int(gdTickCount - g.dLastMouseActivity))
    End If
    
    TimerEnd "frmMain.tmrMain", tmrMain.Interval

ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "frmMain.tmrMain_Timer"
    Resume ErrExit
    
End Sub

Private Sub CheckForGraphNow()
On Error Resume Next

    Dim strGraphNow$, strText$, nSymbolID&, i&, strPrd$
    
    strGraphNow = AddSlash(App.Path) + "GRAPH.NOW"
    If Not FileExist(strGraphNow) Then Exit Sub
    strText = Trim(FileToString(strGraphNow, , True))
    KillFile strGraphNow
    strGraphNow = strText
    If Len(strGraphNow) = 0 Then Exit Sub
    
    nSymbolID = g.SymbolPool.SymbolIDforSymbol(strGraphNow)
    If nSymbolID <> 0 Then
        'MoveFocus frmMain
            
        If Me.WindowState <> 1 Then Me.WindowState = 1
        'WindowStateX(frmMain) = wsMinimized

#If 0 Then
        If Not FormIsLoaded("frmChart2") Then Load frmChart2
        With frmChart2
            .Chart.SetSymbol nSymbolID, True
            ShowForm frmChart2
            .TopMost = True
        End With
#End If
    End If

    'MsgBox graph_now + Str(Len(graph_now))

End Sub

Private Sub CheckForDoItNow()
On Error Resume Next

    Dim strDoItNow$, strCmd$, strText$, nSymbolID&, i&, iCmd&, hWnd&, dDate#
    Dim strSymbol$, nPeriodicity&
    Dim aCmds As New cGdArray
    Static bSnapshotAlreadyShown As Boolean
    
    strDoItNow = AddSlash(App.Path) + "DoIt.NOW"
    If Not FileExist(strDoItNow) Then Exit Sub
    aCmds.FromFile strDoItNow
    KillFile strDoItNow
    
    For iCmd = 0 To aCmds.Size - 1
        'parse:  CMD=Text
        strText = Trim(aCmds(iCmd))
        strCmd = Trim(Parse(strText, "=", 1))
        i = InStr(strText, "=")
        If i > 0 Then
            strText = Trim(Mid(strText, i + 1))
        Else
            strText = ""
        End If
        
        Select Case UCase(strCmd)
        Case "SETTINGS"
            If UCase(strText) = "ACCOUNT" Then
                If WindowStateX(Me) = wsMinimized Then
                    WindowStateX(Me) = wsMaximized
                End If
                frmConfig.ShowMe eAccountTab
            End If
            
        Case "CHART"
            nSymbolID = g.SymbolPool.SymbolIDforSymbol(Parse(strText, "|", 1))
            If nSymbolID <> 0 Then
#If 1 Then
                ' NEW method: show chart and snapshot within TradeNav
                If Not ActiveChart Is Nothing Then
                    ActiveChart.Chart.SetSymbol nSymbolID, True
                End If
                Sleep
                If Me.WindowState = 1 Then
                    ShowWindow Me.hWnd, SW_RESTORE
                End If
                ' need to set form topmost in order to display over ETA
                SetForegroundWindow Me.hWnd
                SetFormTopmost frmMain, True
                ' just show snapshot once -- if they have closed
                ' the window we don't want to keep reshowing it
                If frmSnapshot.Visible Or Not bSnapshotAlreadyShown Then
                    frmSnapshot.ShowMe nSymbolID
                    bSnapshotAlreadyShown = True
                End If
                ' once shown, release topmost status
                Sleep 0.1
                SetFormTopmost frmMain, False
                
#Else
                ' OLD method: show Chart2 by itself
                WindowStateX(frmMain) = wsMinimized
        
                If Not FormIsLoaded("frmChart2") Then Load frmChart2
                With frmChart2
                    .Chart.SetSymbol nSymbolID, True
                    ShowForm frmChart2
                    .TopMost = True
                End With
#End If
            End If
            
        Case "CHARTDATE"
            hWnd = Val(Parse(strText, vbTab, 1))
            dDate = RoundToSecond(Val(Parse(strText, vbTab, 2)))
            If dDate > 0 And dDate < 199999 Then
                strSymbol = UCase(Parse(strText, vbTab, 3))
                nPeriodicity = GetPeriodicity(Parse(strText, vbTab, 4))
                For i = Forms.Count - 1 To 0 Step -1
                    If IsFrmChart(Forms(i)) Then
                        If hWnd = 0 Or hWnd = -1 Then
                            ' do for any chart with same symbol and bar period
                            If Forms(i).Chart.Symbol = strSymbol Then
                                If Forms(i).Chart.Bars.Prop(eBARS_Periodicity) = nPeriodicity Then
                                    Forms(i).CenterTheDate dDate, Val(Parse(strText, vbTab, 7))
                                End If
                            End If
                        ElseIf Forms(i).hWnd = hWnd Then
                            ' only for chart with trades from this system
                            Forms(i).CenterTheDate dDate, Val(Parse(strText, vbTab, 7))
                            Exit For
                        End If
                    End If
                Next
            End If
        End Select
    Next

End Sub

Public Sub InitialShow(ByVal bHidden As Boolean)
On Error GoTo ErrSection:

    Dim ws%, strText$
    Dim frm As Form
       
    m.bShouldBeSeen = True
    ws = Val(GetIniFileProperty("MDI_State", 2, "Forms", g.strIniFile))
    If ws <> 2 Then ws = 0
    If bHidden Then
        WindowState = 1
    ElseIf ws = 2 Then
        ' seem to need to call resize then maximize in order to get the
        ' maximized MDI parent to initially maximize in the correct monitor
        Me.Visible = False
        MDIForm_Resize
        WindowState = ws
        Me.Visible = True
    Else
        ' call resize to trigger moving onto visible part of screen
        MDIForm_Resize
    End If
    m.iPrevNonminimizedState = ws
    tmrAutoResize.Enabled = True
    DoEvents
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.InitialShow", eGDRaiseError_Raise
    
End Sub

' Sets flag to wait for specified # seconds before checking for new symbols
' - to put it on hold, should pass # bigger than what you think you'll need
' - then pass 0 when done (to take it off hold)
Public Sub SuspendNewSymbolCheck(Optional ByVal nSecondsToWait As Long = 0)
On Error GoTo ErrSection:

    If nSecondsToWait <= 0 Then
        ' clear flag
        m.dWaitToCheckRT = 0
    Else
        ' set flag to when to start checking again
        m.dWaitToCheckRT = gdTickCount + nSecondsToWait * 1000#
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.SuspendNewSymbolCheck", eGDRaiseError_Raise
End Sub

' TBL: this acts as a queue to play multiple sounds which have been triggered
Private Sub tmrPlaySound_Timer()
On Error Resume Next

    Dim strSoundFile$
    Static bInProgress As Boolean

    TimerStart "frmMain.tmrPlaySound"
    If g.bUnloading Then
        m.aSoundToPlay.Size = 0
        tmrPlaySound.Enabled = False
        Exit Sub
    End If
    
    If bInProgress Then Exit Sub
    bInProgress = True

    Do While m.aSoundToPlay.Size > 0
        strSoundFile = m.aSoundToPlay(0)
        m.aSoundToPlay.Remove 0, 1
        ' we can't tell it to "wait until finished" because it blocks our entire process (e.g. qb/charts stop refreshing)
        mGenesis.PlaySoundFile strSoundFile ', True
        ' so instead we'll just wait a couple seconds between sounds which have been queued up
        Sleep 3
    Loop
    tmrPlaySound.Enabled = False
    bInProgress = False
    TimerEnd "frmMain.tmrPlaySound", tmrPlaySound.Interval

End Sub

Private Sub tmrPredLabs_Timer()
On Error GoTo ErrSection:

    Dim i&, d#, iMSM As Long ' minutes since midnight
    Static dNextRefresh As Double
    Static bInProgress As Boolean
    
    TimerStart "frmMain.tmrPredLabs"
    If bInProgress Or (g.nReplaySession <> 0) Or g.bUnloading Or (g.RealTime Is Nothing) Then Exit Sub
    bInProgress = True
    
    ' when realtime is running, auto-refresh the Prediction Labs data
    ' every 5 minutes during the day (at 30 seconds prior to the even
    ' 5-minute marks -- e.g. 9:34:30, 9:39:30, 9:44:30, ... 16:14:30)
    If g.RealTime.Active Then
        If g.RealTime.FeedTime > dNextRefresh Then
            ' get the data (if between 9:30 and 16:30)
            dNextRefresh = g.RealTime.FeedTime
            iMSM = Hour(dNextRefresh) * 60 + Minute(dNextRefresh)
            If iMSM >= 570 And iMSM <= 990 Then
                i = GetPredictionLabsData
                DebugLog "GetPredictionLabsData = " & Str(i) & " at " & Format(g.RealTime.FeedTime, "hh:mm:ss")
            End If
        
            ' set next refresh to 30 seconds before the next even 5-minute mark
            iMSM = iMSM + 2
            Do While iMSM Mod 5 <> 0
                iMSM = iMSM + 1
            Loop
            dNextRefresh = Int(dNextRefresh) + (iMSM - 0.5) / 1440#
        End If
    Else
        ' when realtime is not running, this is just a one-time call (so turn timer back off)
        tmrPredLabs.Enabled = False
        i = GetPredictionLabsData
        DebugLog "GetPredictionLabsData = " & Str(i)
    End If
    TimerEnd "frmMain.tmrPredLabs", tmrPredLabs.Interval

ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "frmMain.tmrPredLabs_Timer", eGDRaiseError_Show
End Sub

Private Sub tmrQuickStart_Timer()
On Error Resume Next
    
    Dim strText$

    TimerStart "frmMain.tmrQuickStart"
    tmrQuickStart.Enabled = False
    If ExtremeCharts = 0 Then
        strText = "www.TradeNavigator.com/QuickStart.asp?S=*"
        strText = FixURL(GetProvidedProperty("QuickStartWeb", strText))
        RunProcess InternetBrowser, strText
    End If
    TimerEnd "frmMain.tmrQuickStart", tmrQuickStart.Interval

End Sub

Private Sub tmrSymbol_Timer()
On Error GoTo ErrSection:
    
    Dim i&, strSymbols$, strCaption$, strCenter$, strGroup$, strSymbol$
    Dim aSymbols As cGdArray
    Dim frm As New frmSymbolSelector

    TimerStart "frmMain.tmrSymbol"
    tmrSymbol.Enabled = False
    
    strCaption = Parse(tmrSymbol.Tag, vbTab, 1)
    If Len(Trim(strCaption)) = 0 Then strCaption = "Symbol Selector"
    strCenter = Parse(tmrSymbol.Tag, vbTab, 2)
    If Len(Trim(strCenter)) = 0 Then strCenter = " "
    strGroup = Parse(tmrSymbol.Tag, vbTab, 3)
    
    Set frm = New frmSymbolSelector
    i = frm.Width ' just to make sure form is loaded
    Set aSymbols = frm.ShowMe(, False, , strCaption, , , False, strGroup, , strCenter, True)
    
    strSymbols = ""
    If Not aSymbols Is Nothing Then
        For i = 0 To aSymbols.Size - 1
            strSymbol = aSymbols(i)
            If SecurityType(strSymbol) = "F" Then
                strSymbol = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
                If Len(strSymbol) = 0 Then
                    strSymbol = aSymbols(i)
                End If
                strSymbol = RollSymbolForDate(Parse(strSymbol, "-", 1) & "-067")
            End If
            If Len(strSymbols) = 0 Then
                strSymbols = strSymbol
            Else
                strSymbols = strSymbols & vbTab & strSymbol
            End If
        Next
    End If
    
    SendMessageToOptNav eGDOptNav_Symbols, strSymbols
    TimerEnd "frmMain.tmrSymbol", tmrSymbol.Interval
    
ErrExit:
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.tmrSymbol_Timer", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub tmrWindowLink_Timer()
On Error GoTo ErrSection:

    Dim iForm&, nColor&, iPass&, nActiveChartColor&
    Dim bCheckIt As Boolean, bClearSymbol As Boolean
    Dim frm As Form, frmActive As Form
    Dim bSymGridActive As Boolean
    
    Dim iWindowState&
    
    TimerStart "frmMain.tmrWindowLink"
    If Screen.ActiveForm Is frmSymbolGrid Then
        bSymGridActive = True
    End If
    
    If m.nSymbolLinkID <> 0 Then
        'get window state of non-detached chart
        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
            iWindowState = g.ChartGlobals.frmActiveNonDetached.WindowState
        End If
        ' get color of the active chart
        Set frmActive = ActiveChart
        If Not frmActive Is Nothing Then
            nActiveChartColor = frmActive.WindowLink.SymbolColor
            ' treat unlinked active chart same as black
            If nActiveChartColor = 0 Then nActiveChartColor = 1
            ' if linking blacks, change it to color of active chart
            If m.nSymbolLinkColor = 1 Then m.nSymbolLinkColor = nActiveChartColor
        End If
        
        For iPass = 1 To 3
            For iForm = 0 To Forms.Count - 1
                Set frm = Forms(iForm)
                bCheckIt = False
                bClearSymbol = False
                Select Case iPass
                Case 1
                    If Not IsFrmChart(frm) Then
                        bCheckIt = True
                        If TypeOf frm Is frmTickDistribution Then
'                            bClearSymbol = True
                        End If
                    End If
                Case 2
                    If IsFrmChart(frm) Then
                        If Not frm.Chart.Bars.IsIntraday Then bCheckIt = True
                    End If
                Case 3
                    If IsFrmChart(frm) Then
                        If frm.Chart.Bars.IsIntraday Then bCheckIt = True
                    ElseIf TypeOf frm Is frmTickDistribution Then
'                        bCheckIt = True
                    End If
                End Select
                    
                If bCheckIt Then
                    nColor = -1
                    On Error Resume Next
                    nColor = frm.WindowLink.SymbolColor
                    On Error GoTo ErrSection:
                    ' treat black same as the active chart color
                    If nColor = 1 Then
                        nColor = nActiveChartColor
                    ElseIf frm Is frmActive Then
                        ' treat unlinked active chart same as black
                        If nColor = 0 Then nColor = 1
                        If IsFrmChart(frm) Then
                            'spread charts have symbol ID = 0, attempting to link active spread chart
                            'will cause form to load the linked symbol ID and wipe out the spread
                            If Len(frm.Chart.SpreadSymbols) > 0 Or Len(frm.Chart.ExternalData) > 0 Then
                                nColor = 0     'aardvark 3391 and 4595
                            End If
                        End If
                    End If
                    If nColor = m.nSymbolLinkColor Then
                        If frm.SymbolID <> m.nSymbolLinkID Then
                            ' change symbol for this form
                            If bClearSymbol Then
                                frm.SymbolID = 0
                            Else
                                frm.SymbolID = m.nSymbolLinkID
                                Set frm = Nothing
                                Set frmActive = Nothing
                                ' then exit sub and allow timeslice before changing for another chart
                                GoTo ErrExit        'Exit Sub
                            End If
                        End If
                    End If
                End If
            Next
        Next
    End If
    
    If m.nPeriodLinkID <> 0 Then
        For iForm = 0 To Forms.Count - 1
            Set frm = Forms(iForm)
            If IsFrmChart(frm) Then
                nColor = -1
                On Error Resume Next
                nColor = frm.WindowLink.PeriodColor
                On Error GoTo ErrSection:
                If nColor = m.nPeriodLinkColor Then
                    If frm.Periodicity <> m.nPeriodLinkID Then
                        ' change period for this form
                        frm.Periodicity = m.nPeriodLinkID
                        Set frm = Nothing
                        Set frmActive = Nothing
                        ' then exit sub and allow timeslice before changing for another chart
                        GoTo ErrExit    'Exit Sub
                    End If
                End If
            End If
        Next
    End If
    
    If iWindowState = vbMaximized Then
        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
            If g.ChartGlobals.frmActiveNonDetached.WindowLink.SymbolColor > 0 Or g.ChartGlobals.frmActiveNonDetached.WindowLink.PeriodColor > 0 Then
                g.ChartGlobals.frmActiveNonDetached.SetChartTabs True       '4885, 5452
            Else
                bSymGridActive = False
            End If
        End If
    End If
    
    ' if got here, nothing needed to be changed
    tmrWindowLink.Enabled = False
    TimerEnd "frmMain.tmrWindowLink", tmrWindowLink.Interval

ErrExit:
    Set frm = Nothing
    Set frmActive = Nothing
    If iWindowState = vbMaximized Then
        If bSymGridActive Then
            MoveFocus frmSymbolGrid.fgVirtual
        End If
    End If
    
    Exit Sub
    
ErrSection:
    tmrWindowLink.Enabled = False
    RaiseError "frmMain.tmrWindowLink", eGDRaiseError_Show
End Sub

Public Sub SetWindowLink(frm As Form, Optional ByVal eMode As eLinkMode = eLink_Symbol)
On Error GoTo ErrSection:

    Dim nID&, nColor&
    
    If SubclassingEnabled Then
        If eMode = eLink_Period Then
            nID = frm.Periodicity
            nColor = frm.WindowLink.PeriodColor
            If nID <> 0 And nColor <> 0 Then
                m.nPeriodLinkID = nID
                m.nPeriodLinkColor = nColor
                tmrWindowLink.Enabled = True
            End If
        Else
            nID = frm.SymbolID
            nColor = frm.WindowLink.SymbolColor
            If (nColor = 0) And frm Is ActiveChart Then
                ' treat unlinked active chart same as black
                nColor = 1
            End If
            If nID <> 0 And nColor <> 0 Then
                m.nSymbolLinkID = nID
                m.nSymbolLinkColor = nColor
                tmrWindowLink.Enabled = True
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.SetWindowLink", eGDRaiseError_Raise
End Sub

Public Function ToolbarVisible(ByVal strToolbar$)
On Error Resume Next

    If strToolbar = kTbDraw Then
        ToolbarVisible = pbTbBackDraw(0).Visible
    Else
        ToolbarVisible = True
    End If

End Function

Public Sub ToolBarBtnSizeGet(ByVal strToolbar$, X&, Y&)
'returns width, height in pixels
On Error GoTo ErrSection:

'    If strToolbar = kTbDraw Then
'        If ValOfText(Parse(pbTbBack(0).Tag, ";", 1)) = kBtnLargeIco Then
'            X = kBtnLargeIco
'            Y = kBtnLargeIco
'        Else
'            X = 22
'            Y = 22
'        End If
'    Else
        X = ValOfText(Parse(pbTbBack(0).Tag, ";", 1))
        Y = ValOfText(Parse(pbTbBack(0).Tag, ";", 2))
'    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMain.ToolBarBtnSizeGet"

End Sub

Public Sub ToolBarBtnSizeSet(ByVal strToolbar$, X&, Y&)
On Error GoTo ErrSection:

'    If strToolbar <> kTbDraw Then
        'sets width, height in pixels
        pbTbBack(0).Tag = Str(X) & ";" & Str(Y)
'    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMain.ToolBarBtnSizeSet"

End Sub

Public Function ToolBarWrapGet(ByVal strToolbar$) As Boolean
On Error GoTo ErrSection:

    ToolBarWrapGet = m.bToolbarWrap

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmMain.ToolBarWrapGet"

End Function

Public Sub ToolBarWrapSet(ByVal strToolbar$, ByVal bWrap As Boolean)
On Error GoTo ErrSection:

    m.bToolbarWrap = bWrap

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMain.ToolBarWrapSet"

End Sub

Public Property Let LastClickedToolID(ByVal strID$)
On Error GoTo ErrSection:
    
    m.strLastToolID = strID

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMain.LastClickedToolID"

End Property

Public Property Get TbButtonsArray(ByVal strToolbar$) As cGdArray
On Error GoTo ErrSection:

    If strToolbar = kTbDraw Then
        Set TbButtonsArray = m.aTbButtonsDraw
    Else
        Set TbButtonsArray = m.aTbButtons
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMain.TbButtonsArray"

End Property

Public Property Get LastMouseButton() As cPicBoxButton
On Error GoTo ErrSection:

    Set LastMouseButton = m.oBtnMouseLast

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmMain.LastMouseButton"

End Property

'check if mouse is still over last highlighted toolbar button
Public Sub CheckHighlightedToolButton(Optional ByVal bFromAMouseMoveEvent As Boolean = True)
On Error Resume Next
    
    If Not m.oBtnMouseLast Is Nothing Then
        If m.oBtnMouseLast.ToolBarName = kTbDraw Then
            If m.oBtnMouseLast.CursorCheckClear(Me, m.aTbButtonsDraw) Then Set m.oBtnMouseLast = Nothing
        Else
            If m.oBtnMouseLast.CursorCheckClear(Me, m.aTbButtons) Then Set m.oBtnMouseLast = Nothing
        End If
    End If
    
    If bFromAMouseMoveEvent Then
        g.dLastMouseActivity = gdTickCount
    End If

End Sub

' Queues up a sound to play
Public Sub PlaySound(ByVal strSoundFile$)
On Error Resume Next

    m.aSoundToPlay.Add strSoundFile
    tmrPlaySound.Enabled = True

End Sub

'Public Function img16() As ImageList
Public Function img16() As Object
    Set img16 = g.CoreBridge.Image16
End Function

Public Function ImageList1() As Object
    Set ImageList1 = g.CoreBridge.ImageList1
End Function

Public Function ImageList2() As Object
    Set ImageList2 = g.CoreBridge.ImageList2
End Function

