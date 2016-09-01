VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmReplay 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stream Replay"
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniButtonImageXP cmdBack 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frmReplay.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmReplay.frx":002E
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmReplay.frx":0070
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdStop 
      Height          =   300
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   315
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frmReplay.frx":008C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmReplay.frx":00BA
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmReplay.frx":00F0
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdForward 
      Height          =   300
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   375
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
      Caption         =   "frmReplay.frx":010C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmReplay.frx":0140
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmReplay.frx":0188
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdPlay 
      Height          =   420
      Left            =   4860
      TabIndex        =   6
      Top             =   60
      Width           =   435
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
      Caption         =   "frmReplay.frx":01A4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmReplay.frx":01D2
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmReplay.frx":0250
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdSettings 
      Height          =   360
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   1275
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
      Caption         =   "frmReplay.frx":026C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmReplay.frx":02A8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmReplay.frx":02C8
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   555
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   1995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmReplay.frx":02E4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmReplay.frx":0310
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmReplay.frx":0330
      RightToLeft     =   0   'False
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   285
         Left            =   600
         TabIndex        =   1
         ToolTipText     =   "Select replay speed (hotkey: Up/Down arrows)"
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   1
         Min             =   -1
         Max             =   3
         TextPosition    =   1
      End
      Begin HexUniControls.ctlUniLabelXP lblSpeed 
         Height          =   195
         Index           =   1
         Left            =   900
         Top             =   300
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmReplay.frx":034C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmReplay.frx":0370
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmReplay.frx":0390
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSpeed 
         Height          =   195
         Index           =   0
         Left            =   600
         Top             =   300
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmReplay.frx":03AC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmReplay.frx":03D2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmReplay.frx":03F2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSpeed 
         Height          =   195
         Index           =   4
         Left            =   1620
         Top             =   300
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmReplay.frx":040E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmReplay.frx":0432
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmReplay.frx":0452
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSpeed 
         Height          =   195
         Index           =   3
         Left            =   1380
         Top             =   300
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmReplay.frx":046E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmReplay.frx":0492
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmReplay.frx":04B2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSpeed 
         Height          =   195
         Index           =   2
         Left            =   1140
         Top             =   300
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmReplay.frx":04CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmReplay.frx":04F2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmReplay.frx":0512
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   435
         Left            =   0
         Top             =   30
         Width           =   615
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmReplay.frx":052E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmReplay.frx":0568
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmReplay.frx":0588
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblStatus 
      Height          =   300
      Left            =   1800
      Top             =   120
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmReplay.frx":05A4
      BackColor       =   8454016
      ForeColor       =   8388608
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmReplay.frx":05E8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmReplay.frx":0608
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmReplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum eReplayStatus
    eReplay_Inactive = 0
    eReplay_Paused = 1
    eReplay_Loading = 2
    eReplay_Playing = 3
End Enum

Private Type mPrivate
    dPlayTime As Double
    dMinutesBump As Long
    
    eStatus As eReplayStatus
    dStartTime As Double
    dFeedTime As Double
End Type
Private m As mPrivate

Private Sub cmdBack_Click()
On Error GoTo ErrSection:
   
    Dim dTime#

    SetSpeed True ' pause

#If 0 Then
    ' back up to nearest half-hour
    dTime = dtTime.Value
    dTime = dTime - Int(dTime)
    dtTime.Value = (Int(dTime * 1440 / m.dMinutesBump - 0.000001) + 0) / (1440 / m.dMinutesBump)
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.cmdBack_Click"
End Sub

Private Sub cmdForward_Click()
On Error GoTo ErrSection:

    Dim dTime#

'    SetSpeed True

    ' go forward to nearest half-hour
    dTime = g.RealTime.FeedTime
    If dTime > 0 Then
        dTime = (Int(dTime * 1440 / m.dMinutesBump + 0.000001) + 1) / (1440 / m.dMinutesBump)
        Play dTime
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.cmdForward_Click"
End Sub

Private Sub cmdPlay_Click()
On Error GoTo ErrSection:

    Play

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.cmdPlay_Click"
End Sub

Private Sub cmdSettings_Click()
On Error GoTo ErrSection:

    SetSpeed True
    frmGameModeCfg.ShowMe m.dPlayTime

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.cmdSettings_Click"
End Sub

Private Sub cmdStop_Click()
On Error GoTo ErrSection:

    If g.RealTime.Active Then
        If g.TradingItems.HasActiveAutoTradeItems Then
            g.TradingItems.DisableTradeItems "User stopped replay"
        End If
        
        Me.Hide
        g.RealTime.Init False
        FixButtons
        Unload Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.cmdStop_Click"
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_Replay"), , True)
    'Me.Move frmMain.Left + frmMain.Width - Me.Width, frmMain.Top
    'Me.Move frmMain.Left + (frmMain.Width - Me.Width) / 2, frmMain.Top - Me.Height / 2 + 60
    
    m.dMinutesBump = 5
    lblStatus.Caption = ""
    FixButtons
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.Form_Load"
End Sub

Public Sub ShowMe()
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean
    If Not bAlreadyDone Then
        'bAlreadyDone = True
        'Me.Move frmMain.Left + frmMain.Width - Me.Width, frmMain.Top
        'Me.Move frmMain.Left + (frmMain.Width - Me.Width) / 2, frmMain.Top - Me.Height / 2 + 60
        Me.Move frmMain.Left + frmMain.Width - Me.Width, frmMain.Top - Me.Height / 2 + 60
    End If
   
    ShowForm Me, eForm_Nonmodal, frmMain
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.ShowMe"
End Sub

Public Sub Play(Optional ByVal dStartTime# = 0)
On Error GoTo ErrSection:

    Dim d#, strFile$, strZipFile$
          
    ''Me.Caption = "Stream Replay for " & DateFormat(dtDate.Value)
    ''dStart = RoundToMinute(dtDate.Value + ConvertTimeZone(dtTime.Value, "", "NY"))
    
    ' if user has backed up the time or gone to a new day,
    ' then we need to stop and restart the replay
    If g.RealTime.Active Then
        d = RoundToMinute(g.RealTime.FeedTime)
        If d <= 0 Or (dStartTime > 0 And dStartTime < d) Or (Int(d) <> Int(dStartTime) And dStartTime > 0) Then
            g.RealTime.Init False
            DoEvents
        End If
    End If

    If Not g.RealTime.Active Then
        If dStartTime <= 0 Then
            frmGameModeCfg.ShowMe m.dPlayTime
            Exit Sub
        End If
        m.dPlayTime = dStartTime
        strFile = Format(dStartTime, "YYYYMMDD") & ".RTS"
        If Not FileExist(App.Path & "\RTS\" & strFile) Then
            ' Download the recorded file
            If InfBox("The recorded data for this session needs to be downloaded (this may take a few minutes).", "i", "+Download|-Cancel", "Streaming Replay") = "C" Then
                Exit Sub
            End If
            If ProcessIsBusy Then Exit Sub
            If Not frmMain.Enabled Then Exit Sub
            If FormIsLoaded("frmDownload") Then Unload frmDownload
            Set MsgForm = frmStatus
            frmDownload.optSpecialFile = True
            strZipFile = "RT" & Mid(strFile, 3, 6) & ".gzp"
            lblStatus = "Downloading ..."
            FixButtons
            frmDownload.txtSpecialFile = strZipFile
            frmDownload.DownloadData
            Set MsgForm = Nothing
            ' Move the file
            If Not FileExist(App.Path & "\FTP\" & strFile) Then
                InfBox "This session is not available.|Please select a different date.", "!", , "Streaming Replay"
                Exit Sub
            End If
            ' the "Name ... As" seems to cause problems sometimes, so just copy the file instead
            lblStatus = "Storing data ..."
            DoEvents
            FileCopy App.Path & "\FTP\" & strFile, App.Path & "\RTS\"
            DoEvents
        End If
    
        lblStatus.Caption = "Loading Data ..."
        SetSpeed
        FixButtons
        ' start-up stream from specified point
        g.RealTime.SetReplayTime dStartTime
        g.RealTime.Init True
    ElseIf dStartTime > 0 Then
        m.dPlayTime = dStartTime
        g.RealTime.SetReplayTime dStartTime
        SetSpeed
    ElseIf g.RealTime.ReplaySpeed = 0 Then
        SetSpeed
    Else
        SetSpeed True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.Play"
End Sub

Public Sub SetSpeed(Optional ByVal bPause As Boolean = False)
On Error GoTo ErrSection:

    Dim dSpeed#
    If bPause Then
        dSpeed = 0
    Else
        dSpeed = 2 ^ (sldSpeed.Value)
    End If
    g.RealTime.ReplaySpeed = dSpeed
    FixButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.SetSpeed"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.Form_QueryUnload"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    If Not g.bUnloading And Not g.RealTime Is Nothing Then
        If g.RealTime.Active Then
            If g.TradingItems.HasActiveAutoTradeItems Then
                g.TradingItems.DisableTradeItems "Replay form is unloading"
            End If
            Me.Visible = False
            g.RealTime.Init False
        End If
    End If
    g.nReplaySession = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.Form_Unload"
End Sub

Private Sub lblSpeed_Click(Index As Integer)
On Error GoTo ErrSection:

    sldSpeed.Value = sldSpeed.Min + Index

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.lblSpeed_Click"
End Sub

Private Sub sldSpeed_Change()

    SetSpeed

End Sub

Private Sub sldSpeed_Click()

    SetSpeed

End Sub

Public Sub FixButtons()
On Error GoTo ErrSection:

    Dim nColor&

    If g.RealTime Is Nothing Then Exit Sub

    nColor = &H40FFFF ' yellow
    If Not g.RealTime.Active Then
        ' stopped
        cmdStop.Enabled = False
        cmdPlay.Enabled = True
        'RH commented out cmdPlay.Picture = Picture16("kPlay")
    ElseIf g.RealTime.ReplaySpeed = 0 Then
        ' paused
        cmdStop.Enabled = True
        cmdPlay.Enabled = True
        'RH commented out cmdPlay.Picture = Picture16("kPlay")
        cmdPlay.ToolTipText = "Play the recorded stream"
    Else
        ' playing
        cmdStop.Enabled = True
        cmdPlay.Enabled = True
        'RH commented out cmdPlay.Picture = Picture16("kPause")
        cmdPlay.ToolTipText = "Pause the recorded stream"
        If IsDigit(lblStatus.Caption, 1) Then
            nColor = &H80FF80 ' green
        End If
    End If
    
    lblStatus.BackColor = nColor
    
    If nColor = &H40FFFF Then
        g.SimTradeReplay.Broker.HandleConnectionInfo eGDConnectionStatus_Connecting, "", ""
    Else
        g.SimTradeReplay.Broker.HandleConnectionInfo eGDConnectionStatus_Connected, "", ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.FixButtons"
End Sub

Public Sub UpdateTime(ByVal dTime#)
On Error GoTo ErrSection:

    Dim s$, dTimeTZ#, bReload As Boolean
    Static dPrevTime#

    If dTime > 0 Then
        dTime = RoundToMinute(dTime)
        If g.bShowInLocalTimeZone Then
            dTimeTZ = ConvertTimeZone(dTime, "NY", "")
        Else
            dTimeTZ = dTime
        End If
        If dTime < m.dPlayTime Then
            ' must be reading data before ready to play
            s = "Reading " & DateFormat(dTimeTZ, NO_DATE, H_MM, AMPM_LOWER)
        Else
            ' playing the data
            'check for word loading in status label is fix for issue 5327
            If InStr(lblStatus, "Loading") > 0 Or (dTime = m.dPlayTime And dTime > dPrevTime And g.RealTime.Active) Then
                ' just caught up, so reload all forms
                bReload = True
            End If
            m.dPlayTime = dTime
            s = DateFormat(dTimeTZ, MM_DD_YYYY, H_MM, AMPM_LOWER)
        End If
        If s <> lblStatus.Caption Then
            lblStatus.Caption = s
            FixButtons
            lblStatus.Refresh
        End If
    End If
    dPrevTime = dTime

    If bReload Then
        g.RealTime.RefreshAllFormData True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReplay.UpdateTime"
End Sub

Public Property Get IsPlaying() As Boolean
            
    If lblStatus.BackColor = &H80FF80 Then  ' green
        IsPlaying = True
    End If
    
End Property

Public Property Get Status() As eReplayStatus
    Status = m.eStatus
End Property

Private Property Let Status(ByVal eStatus As eReplayStatus)
On Error GoTo ErrSection:

    If eStatus <> m.eStatus And Not g.RealTime Is Nothing Then
        Select Case eStatus
            Case eReplay_Inactive
            
            Case eReplay_Loading
            
            Case eReplay_Paused
                
            Case eReplay_Playing
            
        End Select
        m.eStatus = eStatus
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmReplay.LetStatus"
End Property

