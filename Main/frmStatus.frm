VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmStatus 
   Caption         =   "Processing Status"
   ClientHeight    =   1230
   ClientLeft      =   -99945
   ClientTop       =   300
   ClientWidth     =   4140
   Icon            =   "frmStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrRecalc 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1380
      Top             =   720
   End
   Begin HexUniControls.ctlUniListBoxXP lstStatus 
      Height          =   450
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   4125
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      TrapTab         =   0   'False
      Tip             =   "frmStatus.frx":0442
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmStatus.frx":0462
      ManualStart     =   0   'False
      Columns         =   0
      RightToLeft     =   0   'False
   End
   Begin VB.PictureBox picMoveUp 
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   180
      Picture         =   "frmStatus.frx":047E
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Hide Details"
      Top             =   900
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picMoveDown 
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   615
      Picture         =   "frmStatus.frx":0788
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Show Details"
      Top             =   150
      Visible         =   0   'False
      Width           =   120
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdDetails 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   15
      Width           =   795
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frmStatus.frx":0A92
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmStatus.frx":0AC8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmStatus.frx":0AE8
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdClose 
      Height          =   315
      Left            =   3300
      TabIndex        =   3
      Top             =   15
      Width           =   795
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frmStatus.frx":0B04
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmStatus.frx":0B30
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmStatus.frx":0B50
      RightToLeft     =   0   'False
   End
   Begin vsOcx6LibCtl.vsElastic vsProgress 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   15
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   3
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   8421504
      ForeColor       =   16777215
      FloodColor      =   16711680
      ForeColorDisabled=   -2147483631
      Caption         =   "Uploading: 50%"
      Align           =   0
      Appearance      =   3
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   1
      FloodPercent    =   50
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin HexUniControls.ctlUniTextBoxXP txtHwnd 
      Height          =   315
      Left            =   420
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmStatus.frx":0B6C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmStatus.frx":0B96
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStatus.frx":0BB6
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmStatus.frm
'' Description: Shows the user the status of their download
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Enum eStatusStep
    eStatus_Initialized = 0
    eStatus_Running = 1
    eStatus_Aborting = 2
    eStatus_Aborted = 3
    eStatus_Error = 4
    eStatus_Completed = 5
End Enum

Private Type mPrivate
    bShowDetails As Boolean
    eStatus As eStatusStep
    dWhenStartedAborting As Double
    bBusyFlag As Boolean
    hMsgHwnd As Long
    nLastKnownStatusCode As Long
End Type
Private m As mPrivate

Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    If ProcessIsBusy(True) Then
        Status = eStatus_Aborting
    Else
        'Unload Me
        'DockState(Me) = DPHidden
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.cmdClose.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdDetails_Click()
On Error GoTo ErrSection:

    If m.bShowDetails Then
        SetIniFileProperty "StatusHeight", Me.Height, "General", g.strIniFile
    End If
    
    ShowDetails (Not m.bShowDetails)
    
    MoveFocus lstStatus
    If Not m.bShowDetails Then
    '    MoveFocus cmdClose
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.cmdDetails.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, reset the toolbar and the window
''              list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    'ToolbarSync Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, let everyone know that it is loaded
''              and center it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlace$, w&, i&

    picMoveUp.Move picMoveDown.Left, picMoveDown.Top
       
    g.Styler.StyleForm Me
    
    w = 4200
    strPlace = GetIniFileProperty("Status", "", "Placement", g.strIniFile)
    If strPlace = "" Then
        Me.Move Screen.Width - w - 1050, 0, w, 1455
    Else
        SetFormPlacement Me, strPlace ', "LT"
    End If

    'ShowDetails True '(for now)
    ShowDetails GetIniFileProperty("StatusDetails", True, "General", g.strIniFile)

        'CenterTheForm Me
    If 0 Then
        CenterTheForm Me
    Else
        'Me.FormX1.Dock otxLeftEdge, 1, 1, 25, otxPosEndRel
    End If
    Status = eStatus_Initialized

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        If Not ProcessIsBusy Then
            Me.Hide
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim h As Long

    h = cmdDetails.Height + Screen.TwipsPerPixelY * 2 '+ Me.Height - Me.ScaleHeight
    If m.bShowDetails Then
        h = h + 660
        'If Me.Height < h + 660 Then
            'Me.Height = h + 660
            'Exit Sub
        'End If
    Else
        If Me.Height <> h Then
            'Me.Height = h
            'Exit Sub
        End If
    End If
    
    If LimitFormSize(Me, cmdClose.Width * 3, h) Then Exit Sub
    If Not m.bShowDetails Then
        If Me.ScaleHeight <> h Then
            Me.Height = h + Me.Height - Me.ScaleHeight
            Exit Sub
        End If
    End If
    
    cmdClose.Left = Me.ScaleWidth - cmdClose.Width
    vsProgress.Width = cmdClose.Left - vsProgress.Left - 60
    With lstStatus
        If m.bShowDetails Then
            .Move .Left, .Top, Me.ScaleWidth - .Left, Me.ScaleHeight - .Top
        End If
    End With

    'AutoSizeChart
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, let everyone know that it is no
''              longer loaded, and update the window list
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    'ToolbarSync Me, False
    ''frmMain.DockPro.RemoveForm Me.Name

    If m.hMsgHwnd <> 0 Then
        DestroyWindow m.hMsgHwnd
        m.hMsgHwnd = 0
    End If

    SetIniFileProperty "Status", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "StatusDetails", m.bShowDetails, "General", g.strIniFile
    If m.bShowDetails Then
        SetIniFileProperty "StatusHeight", Me.Height, "General", g.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub picMoveDown_Click()
    cmdDetails_Click
End Sub

Private Sub picMoveUp_Click()
    cmdDetails_Click
End Sub

Private Sub tmrRecalc_Timer()
On Error GoTo ErrSection:

    Dim bIncludeSnapshot As Boolean

    tmrRecalc.Enabled = False
    Status = eStatus_Initialized
    If Val(tmrRecalc.Tag) <> 0 Then bIncludeSnapshot = True
    tmrRecalc.Tag = ""
    If g.SymbolPool.RecalcDirtyCriteria(bIncludeSnapshot) Then
        AddDetail "Finished"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.tmrRecalc.Timer", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtHwnd_Change
'' Description: Change the status depending on who called the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtHwnd_Change()
On Error GoTo ErrSection:

    ProcessStatusMsg txtHwnd.Text

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.txtHwnd.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Sub UpdateProgress(ByVal strCaption$, Optional ByVal nPercent% = -1, Optional ByVal bShowPercent As Boolean = True)
On Error GoTo ErrSection:

    ' to avoid computer going into hibernation/sleep mode while downloading/distributing/etc is in progress
    DoNotHibernateNow

    If m.eStatus = eStatus_Aborting Or m.eStatus = eStatus_Aborted Then
        If InStr(UCase(strCaption), "ABORT") = 0 Then
            Exit Sub
        End If
    ElseIf m.eStatus = eStatus_Error Then
        If InStr(UCase(strCaption), "ERROR") = 0 Then
            Exit Sub
        End If
    End If

    ' fix percent
    If nPercent < 0 Then
        nPercent = 0
        bShowPercent = False
    ElseIf nPercent > 100 Then
        nPercent = 100
    End If
    
    ' show percent with caption?
    If bShowPercent Then
        If strCaption <> "" Then
            strCaption = Trim(strCaption) & ": "
        End If
        strCaption = strCaption & Format(nPercent, "#0") & "%"
    End If
    
    ' set progress bar
    With vsProgress
        If .FloodPercent <> nPercent Or .Caption <> strCaption Then
            .FloodPercent = nPercent
            .Caption = strCaption
            .Refresh
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.UpdateProgress", eGDRaiseError_Raise
    
End Sub

Public Sub AddDetail(ByVal strText$)
On Error GoTo ErrSection:

    ''If DockState(Me) = eHidden Then ShowMe
    
    'vsProgress.Caption = strText
   
    With lstStatus
        .AddItem DateFormat(Now, NO_DATE, H_MM_SS, AP_LOWER) & " - " & strText
        .ListIndex = .ListCount - 1
'RH commented out         .Refresh
    End With

    On Error Resume Next
    If UCase(strText) = "FINISHED" Then
        UpdateProgress "Finished"
        ' auto-hide if successfully completed when not showing details
        If m.eStatus = eStatus_Completed And m.bShowDetails = False Then
            Me.Hide
        End If
    ElseIf Not Me.Visible Then
        ShowMe 'ShowForm Me, , frmMain
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.AddDetail", eGDRaiseError_Raise
    
End Sub

Public Property Get Status() As eStatusStep
On Error GoTo ErrSection:

    ' make like a "timeout" for aborting
    If m.eStatus = eStatus_Aborting And m.dWhenStartedAborting > 0 Then
        ' if more than 15 seconds, just set to aborted
        If gdTickCount - m.dWhenStartedAborting > 15000# Then
            '(FYI: must call "Me.Status" to run the "Let" routine)
            Me.Status = eStatus_Aborted
            KillProcess "Lil'Fred"
        End If
    End If
    Status = m.eStatus

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmStatus.Status.Get", eGDRaiseError_Raise
    
End Property

Public Property Let Status(ByVal eNewStatus As eStatusStep)
On Error GoTo ErrSection:

    ' don't go to aborting if already aborting or aborted
    If eNewStatus = eStatus_Aborting And (m.eStatus = eStatus_Aborting Or m.eStatus = eStatus_Aborted) Then
        Exit Property
    End If
    
    m.dWhenStartedAborting = 0
    If m.eStatus = eStatus_Aborting And eNewStatus > eStatus_Initialized Then
        eNewStatus = eStatus_Aborted
    End If
    m.eStatus = eNewStatus
    Select Case m.eStatus
        Case eStatus_Initialized
            lstStatus.Clear
            UpdateProgress ""
            cmdClose.Caption = "&Close"
            cmdClose.Enabled = True
            m.nLastKnownStatusCode = -1

        Case eStatus_Running
            'make sure toolbar set back to active
            frmMain.tbToolbar.Redraw = True
            cmdClose.Caption = "&ABORT"
            cmdClose.Enabled = True

        Case eStatus_Aborting
            If cmdClose.Enabled Then
                UpdateProgress "ABORTING..."
                cmdClose.Enabled = False
                m.dWhenStartedAborting = gdTickCount
                'put file so GCLIENT knows to abort
                FileFromString App.Path & "\Gclient.can", "Abort"
            End If

        Case eStatus_Aborted
            UpdateProgress "ABORTED"
            AddDetail "ABORTED"
            cmdClose.Caption = "&Close"
            cmdClose.Enabled = True

        Case eStatus_Error
            UpdateProgress "ERROR"
            cmdClose.Caption = "&Close"
            cmdClose.Enabled = True
        
        Case eStatus_Completed
            'UpdateProgress ""
            cmdClose.Caption = "&Close"
            cmdClose.Enabled = True
            
    End Select

    If IsBusy Then
        Me.Icon = Picture16(ToolbarIcon("kProcessingOn"))
    Else
        Me.Icon = Picture16(ToolbarIcon("kProcessingOff"))
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmStatus.Status.Let", eGDRaiseError_Raise
    
End Property

' Parse and process status messages coming back from
' downloading, distributions, etc.
Public Sub ProcessStatusMsg(ByVal strText$)
On Error GoTo ErrSection:

    Dim nPercent&, strMsg$, bDownload As Boolean
    Static bDownloading As Boolean

    'DebugLog strText
    nPercent = Val(Parse(strText, vbTab, 2))
    m.nLastKnownStatusCode = Val(Parse(strText, vbTab, 6))
    Select Case m.nLastKnownStatusCode
        Case 50: ' Display message in "detail" list
            strMsg = Trim(Parse(strText, vbTab, 1))
            If Len(strMsg) > 0 Then AddDetail strMsg
        Case 100: ' Initializing
            UpdateProgress "Connecting ..."
        Case 110:
            UpdateProgress "Initializing ..."
        Case 150:
            UpdateProgress "Uploading", nPercent
        Case 200: ' Processing
            UpdateProgress "Processing Request ...", nPercent, False
        Case 210:
            UpdateProgress "Distributing", nPercent
        Case 250:
            If Not bDownloading Then AddDetail "Downloading Data"
            bDownload = True
            UpdateProgress "Downloading", nPercent
        Case 300:
            UpdateProgress "Unpacking", nPercent
        Case 310:
            UpdateProgress "Final Updating", nPercent
        Case 400: ' Finished
            Status = eStatus_Completed
        Case 500: ' Error
            strMsg = Trim(Parse(strText, vbTab, 1))
            If Len(strMsg) > 0 Then AddDetail "Error: " & strMsg
            Status = eStatus_Error
        Case 600: ' Aborted
            Status = eStatus_Aborted
    End Select
    bDownloading = bDownload
    
    'DoEvents 'TLB 12/1/05: commented this out since I don't think we need it?

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.ProcessStatusMsg", eGDRaiseError_Raise
    
End Sub

' to show the Status window in the "right" spot
Public Sub ShowMe(Optional ByVal bForceActivate As Boolean = False)
On Error Resume Next

    Dim h&, t&, l&, w&
    Dim frmActive As Form
    Static bAlreadyShown As Boolean
    
    ' check if a modal form is active
    If Not frmMain.Enabled Then Exit Sub
    
    On Error Resume Next
    
    If Not Me.Visible Then
        ' first time we must do "ShowForm" to set the owner
        If bForceActivate Or Not bAlreadyShown Then
            If Not bForceActivate Then
                Set frmActive = Screen.ActiveForm
            End If
            ShowForm Me, , frmMain
            If Not frmActive Is Nothing Then
                MoveFocus frmActive
            End If
            bAlreadyShown = True
        Else
            ' after first time, just set visible so form won't grab focus from other apps
            ' (owner set the first time is retained while the form is hidden)
            ShowWindow Me.hWnd, SW_SHOWNA '(use API call so won't get focus from this app either)
            'Me.Visible = True
        End If
    End If
    frmMain.tbToolbar.Tools("ID_ProcessingStatus").Enabled = True

End Sub

Public Property Get IsBusy() As Boolean
On Error GoTo ErrSection:

    Dim eStatus As eStatusStep

    ' get status through property (so will allow timeout to work)
    eStatus = Status
    If m.bBusyFlag Or eStatus = eStatus_Aborting Or eStatus = eStatus_Running Then
        IsBusy = True
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmStatus.IsBusy.Get", eGDRaiseError_Raise
    
End Property

Public Property Let IsBusy(ByVal bNewValue As Boolean)
On Error GoTo ErrSection:

    m.bBusyFlag = bNewValue
    If IsBusy Then
        Me.Icon = Picture16(ToolbarIcon("kProcessingOn"))
    Else
        Me.Icon = Picture16(ToolbarIcon("kProcessingOff"))
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmStatus.IsBusy.Let", eGDRaiseError_Raise
    
End Property

Public Sub ShowDetails(ByVal bShowDetails As Boolean)
On Error GoTo ErrSection:

    Dim h&

    m.bShowDetails = bShowDetails
    If m.bShowDetails Then
        'lstStatus.Visible = True
        picMoveUp.Visible = True
        picMoveDown.Visible = False
        cmdDetails.ToolTipText = picMoveUp.ToolTipText
        Me.Height = GetIniFileProperty("StatusHeight", 100, "General", g.strIniFile)
    Else
        'lstStatus.Visible = False
        picMoveUp.Visible = False
        picMoveDown.Visible = True
        cmdDetails.ToolTipText = picMoveDown.ToolTipText
        FormResize Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStatus.ShowDetails", eGDRaiseError_Raise
End Sub

Public Sub SetTitle(ByVal strTitle$)
    If Len(strTitle) = 0 Then strTitle = "Processing Status"
    Me.Caption = strTitle
End Sub

Public Property Get MsgHwnd() As Long
On Error GoTo ErrSection:
    
    Dim i&, strWindowClass$
    Dim wcex As WNDCLASSEX
    
    If m.hMsgHwnd = 0 Then
        With wcex
            'setup window class
            strWindowClass = "GclientMessageWindow"
            .cbSize = Len(wcex)
            .lpszClassName = strWindowClass
            .hInstance = App.hInstance
            .lpfnWndProc = FunctionPtrToLong(AddressOf GclientCallbackProc)
            'make sure window class is registered
            i = RegisterClassEx(wcex)
        End With
        m.hMsgHwnd = CreateWindowEx(0, strWindowClass, strWindowClass, 0, 0, 0, 0, 0, 0, 0, App.hInstance, 0)
    End If
    
ErrExit:
    MsgHwnd = m.hMsgHwnd
    Exit Property
    
ErrSection:
    RaiseError "frmStatus.MsgHwnd"
End Property

Public Property Get LastKnownStatusCode() As Long
    LastKnownStatusCode = m.nLastKnownStatusCode
End Property


