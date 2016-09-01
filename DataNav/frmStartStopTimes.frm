VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmStartStopTimes 
   Caption         =   "Start/End Times"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdateRect 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   255
      Top             =   2640
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1222
      TabIndex        =   9
      Top             =   2535
      Width           =   2235
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
      Caption         =   "frmStartStopTimes.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStartStopTimes.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStartStopTimes.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   60
         Width           =   975
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
         Caption         =   "frmStartStopTimes.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStartStopTimes.frx":0094
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStartStopTimes.frx":00B4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   975
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
         Caption         =   "frmStartStopTimes.frx":00D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStartStopTimes.frx":00F4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStartStopTimes.frx":0114
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame3 
      Height          =   675
      Left            =   202
      TabIndex        =   6
      Top             =   1845
      Width           =   4260
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
      Caption         =   "frmStartStopTimes.frx":0130
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStartStopTimes.frx":015C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStartStopTimes.frx":017C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP lblRectError 
         Height          =   210
         Left            =   1080
         Top             =   75
         Visible         =   0   'False
         Width           =   1785
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
         Caption         =   "frmStartStopTimes.frx":0198
         BackColor       =   -2147483633
         ForeColor       =   0
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmStartStopTimes.frx":01DC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStartStopTimes.frx":01FC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblEndTime 
         Height          =   240
         Left            =   3300
         Top             =   400
         Width           =   1050
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
         Caption         =   "frmStartStopTimes.frx":0218
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmStartStopTimes.frx":024A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStartStopTimes.frx":026A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStartTime 
         Height          =   240
         Left            =   0
         Top             =   400
         Width           =   1050
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
         Caption         =   "frmStartStopTimes.frx":0286
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmStartStopTimes.frx":02BA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStartStopTimes.frx":02DA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Shape shpRect 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   930
         Top             =   60
         Width           =   2220
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   412
         X2              =   412
         Y1              =   345
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   3832
         X2              =   3832
         Y1              =   345
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   412
         X2              =   3832
         Y1              =   160
         Y2              =   175
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraComboBoxes 
      Height          =   435
      Left            =   202
      TabIndex        =   3
      Top             =   1230
      Width           =   4260
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
      Caption         =   "frmStartStopTimes.frx":02F6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStartStopTimes.frx":0330
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStartStopTimes.frx":0350
      RightToLeft     =   0   'False
      Begin gdOCX.gdSelectDate gdStartTime 
         Height          =   315
         Left            =   930
         TabIndex        =   7
         Top             =   60
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         ShowDayOfWeek   =   0   'False
         ShowCalendar    =   0   'False
         ShowPM          =   2
         ShowDate        =   0
         ShowTime        =   2
         MinDate         =   0
         MaxDate         =   0.99999
         Value           =   0
      End
      Begin gdOCX.gdSelectDate gdEndTime 
         Height          =   315
         Left            =   2820
         TabIndex        =   8
         Top             =   60
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         ShowDayOfWeek   =   0   'False
         ShowCalendar    =   0   'False
         ShowPM          =   2
         ShowDate        =   0
         ShowTime        =   2
         MinDate         =   0
         MaxDate         =   0.99999
         Value           =   0
      End
      Begin HexUniControls.ctlUniLabelXP lblCboEnd 
         Height          =   255
         Left            =   2325
         Top             =   90
         Width           =   810
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
         Caption         =   "frmStartStopTimes.frx":036C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmStartStopTimes.frx":0394
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStartStopTimes.frx":03B4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblCboStart 
         Height          =   255
         Left            =   330
         Top             =   90
         Width           =   810
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
         Caption         =   "frmStartStopTimes.frx":03D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmStartStopTimes.frx":03FC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStartStopTimes.frx":041C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   615
      Left            =   202
      TabIndex        =   0
      Top             =   495
      Width           =   4260
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
      Caption         =   "frmStartStopTimes.frx":0438
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStartStopTimes.frx":0464
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStartStopTimes.frx":0484
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optCustom 
         Height          =   255
         Left            =   30
         TabIndex        =   2
         Top             =   330
         Width           =   4095
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
         Caption         =   "frmStartStopTimes.frx":04A0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmStartStopTimes.frx":0518
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmStartStopTimes.frx":0538
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optDefault 
         Height          =   315
         Left            =   30
         TabIndex        =   1
         Top             =   -15
         Width           =   3765
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
         Caption         =   "frmStartStopTimes.frx":0554
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmStartStopTimes.frx":05CE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmStartStopTimes.frx":05EE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblTimeZoneAlert 
      Height          =   315
      Left            =   90
      Top             =   120
      Width           =   4260
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
      Caption         =   "frmStartStopTimes.frx":060A
      BackColor       =   -2147483633
      ForeColor       =   192
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmStartStopTimes.frx":065E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStartStopTimes.frx":067E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmStartStopTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
        
    InBars As cGdBars           'bars object passed in
    
    dDefaultStart As Double
    dDefaultEnd As Double
    
    dStartSel As Double         'value user selected (minutes from midnight)
    dEndSel As Double           'value user selected (minutes from midnight)
    
    nTotalMinutes As Long       'total minutes of normal session
    dTwipsPerMinute As Double
    
    bOK As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:  ShowMe
''
'' Parameters:
''     In/Out:  Bars        - used on input to get symbol information
''                            eBARS_StartTime and eBARS_EndTime changed on exit
''
'' Function return value:   true if user clicked okay AND (start or end time were changed)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Bars As cGdBars, Optional Chart As cChart) As Boolean
On Error GoTo ErrSection:
        
    m.bOK = False
        
    If Bars Is Nothing Then
        InfBox "Invalid parameter. Bars object is null.", "E", , "Start/End Times"
    ElseIf Bars.Prop(eBARS_SymbolID) <= 0 Then
        InfBox "Session times cannot be modified for this symbol.", "E", , "Start/End Times"
    Else
        Set m.InBars = Bars
        InitControls
        If Chart Is Nothing Then
            CenterTheForm Me
        Else
            CenterFormOnChart Me, Chart                 '6499
        End If
        ShowForm Me, eForm_ActModal, frmMain
    End If
        
    If m.bOK Then
        If m.dStartSel = Bars.Prop(eBARS_StartTime) And m.dEndSel = Bars.Prop(eBARS_EndTime) Then
            m.bOK = False       'user did not change start/End times, return false
        Else
            Bars.Prop(eBARS_StartTime) = m.dStartSel
            Bars.Prop(eBARS_EndTime) = m.dEndSel
        End If
    End If

    Set m.InBars = Nothing
    Unload Me
    ShowMe = m.bOK

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmStartStopTimes.ShowMe"
    
End Function

Private Sub cmdCancel_Click()
    m.bOK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Dim strErr$
    
    m.dStartSel = gdStartTime.Value
    m.dStartSel = Round(m.dStartSel * 1440#)    'convert to minutes from midnight
    
    m.dEndSel = gdEndTime.Value
    m.dEndSel = Round(m.dEndSel * 1440#)
    
    If ValidTimes(m.dStartSel, m.dEndSel, True, strErr) Then
        m.bOK = True
        Me.Hide
    Else
        m.dEndSel = gdEndTime.Value                    'restore to time values instead of minutes from midnight
        m.dStartSel = gdStartTime.Value
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStartStopTimes.cmdOK_Click"
    
End Sub

Private Sub InitControls()
On Error GoTo ErrSection:

    Dim dTime#, strZone$, dDefaultStart#, dDefaultEnd#
    Dim DefaultBars As cGdBars
            
    strZone = m.InBars.Prop(eBARS_ExchangeTimeZoneInf)
    If strZone <> "NY" And strZone <> "GMT" Then strZone = "Exchange"
    lblTimeZoneAlert.Caption = "Times are in " & strZone & " time"
    
    m.dDefaultStart = m.InBars.Prop(eBARS_DefaultStartTime)
    m.dDefaultEnd = m.InBars.Prop(eBARS_DefaultEndTime)
    
    If m.dDefaultStart = 0 And m.dDefaultEnd = 0 Then
        Set DefaultBars = New cGdBars
        SetBarProperties DefaultBars, m.InBars.Prop(eBARS_SymbolID), True
        m.dDefaultStart = DefaultBars.Prop(eBARS_StartTime)
        m.dDefaultEnd = DefaultBars.Prop(eBARS_EndTime)
        Set DefaultBars = Nothing
    End If
    
    m.dStartSel = m.InBars.Prop(eBARS_StartTime) / 1440#
    m.dEndSel = m.InBars.Prop(eBARS_EndTime) / 1440#
    
    If m.dDefaultStart > 0 And m.dDefaultEnd > 0 Then
        dDefaultStart = m.dDefaultStart / 1440#
        dDefaultEnd = m.dDefaultEnd / 1440#
        lblStartTime.Caption = DateFormat(dDefaultStart, NO_DATE, HH_MM, NO_AMPM)
        lblEndTime.Caption = DateFormat(dDefaultEnd, NO_DATE, HH_MM, NO_AMPM)
        
        gdStartTime.Value = m.dStartSel
        gdEndTime.Value = m.dEndSel
        
        If m.dStartSel = dDefaultStart And m.dEndSel = dDefaultEnd Then
            optDefault.Value = True
        Else
            optDefault.Value = False
        End If
        
        If m.dDefaultStart > m.dDefaultEnd Then
            m.nTotalMinutes = 1440# - m.dDefaultStart + m.dDefaultEnd
        Else
            m.nTotalMinutes = m.dDefaultEnd - m.dDefaultStart
        End If
        
        If m.nTotalMinutes > 0 Then m.dTwipsPerMinute = (Line1.X2 - Line1.X1) / m.nTotalMinutes

        lblRectError.Left = Line1.X1 + 20
        lblRectError.Width = Line1.X2 - Line1.X1 + 20
        
        UpdateRect True
    Else
        InfBox "Unable to obtain default start/end times.", "E", "OK", "Start/End Times"
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStartStopTimes.InitControls"
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStartStopTimes.Form_Load"
    
End Sub

Private Sub gdStartTime_Changed()
On Error GoTo ErrSection:

    m.dStartSel = gdStartTime.Value

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStartStopTimes.gdStartTime_Changed"
    
End Sub

Private Sub gdEndTime_Changed()
On Error GoTo ErrSection:

    m.dEndSel = gdEndTime.Value
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStartStopTimes.gdEndTime_Changed"
    
End Sub

Private Sub optCustom_Click()
On Error GoTo ErrSection:

    EnableComboFrame True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStartStopTimes.optCustom_Click"
    
End Sub

Private Sub optDefault_Click()
On Error GoTo ErrSection:

    EnableComboFrame False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStartStopTimes.optDefault_Click"
    
End Sub

Private Sub EnableComboFrame(ByVal bEnable As Boolean)
On Error GoTo ErrSection:
    
    If bEnable = True Then
        gdStartTime.Value = m.dStartSel
        gdEndTime.Value = m.dEndSel
    Else
        gdStartTime.Value = m.dDefaultStart / 1440#
        gdEndTime.Value = m.dDefaultEnd / 1440#
    End If
    
    fraComboBoxes.Enabled = bEnable
    gdStartTime.Enabled = bEnable
    gdEndTime.Enabled = bEnable
    lblCboStart.Enabled = bEnable
    lblCboEnd.Enabled = bEnable

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStartStopTimes.EnableComboFrame"
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Validation Rules:
' 1. midnight < custom_end <= default end               (must be true)
'
' 2. if default start > default end
'       (this is over night symbol)
'       custom_start < custom_end OR >= default start   (this is okay)
'    else
'       (this is not over night symbol)
'       default start <= custom start <= custom end     (this is okay)
'    end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidTimes(ByVal dStart#, ByVal dEnd#, ByVal bShowErr As Boolean, strErr As String) As Boolean
On Error GoTo ErrSection:
    
    Dim bValid As Boolean
    Dim bOverNight As Boolean
    
    If m.dDefaultStart > m.dDefaultEnd Then bOverNight = True
    
    If dStart = m.dDefaultStart And dEnd = m.dDefaultEnd Then
        bValid = True
    ElseIf dEnd > 0 And dEnd <= m.dDefaultEnd Then
        If bOverNight Then
            If dStart < dEnd Or dStart >= m.dDefaultStart Then
                bValid = True
            ElseIf dStart > dEnd Then
                strErr = "Start time must be at or after " & DateFormat(m.dDefaultStart / 1440#, NO_DATE, HH_MM, NO_AMPM)
            ElseIf dStart = dEnd Then
                strErr = "Start time must be earlier than end time"
            Else
                strErr = "Start time must be between " & _
                         DateFormat(m.dDefaultStart / 1440#, NO_DATE, HH_MM) & " - " & _
                         DateFormat(m.dDefaultEnd / 1440#, NO_DATE, HH_MM)
            End If
        ElseIf m.dDefaultStart <= dStart And dStart < dEnd Then
            bValid = True
        ElseIf m.dDefaultStart = m.dDefaultEnd Then
            strErr = "Start time must be earlier than end time."
        Else
            strErr = "Start time must be between " & _
                     DateFormat(m.dDefaultStart / 1440#, NO_DATE, HH_MM) & " - " & _
                     DateFormat(m.dDefaultEnd / 1440#, NO_DATE, HH_MM)
        End If
    ElseIf dEnd > m.dDefaultEnd Then
        'failed rule 1
        strErr = "End time cannot be after " & DateFormat(m.dDefaultEnd / 1440#, NO_DATE, HH_MM, NO_AMPM)
    ElseIf dEnd = 0 Then
        strErr = "End time must be after mindnight"
    Else
        strErr = "End must be between 00:01-" & DateFormat(m.dDefaultEnd / 1440#, NO_DATE, HH_MM)
    End If
    
    If bShowErr Then
        If Len(strErr) > 0 Then InfBox strErr, "E", "OK", "Start/End Times"
    End If
        
    cmdOK.Enabled = bValid
    
    ValidTimes = bValid
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmStartStopTimes.ValidTimes"
    
End Function

Private Sub UpdateRect(Optional ByVal bReset As Boolean = False)
On Error GoTo ErrSection:

    Static bInprog As Boolean
    Static dPrevStart#, dPrevEnd#              'minutes from midnight
    
    Dim dStart#, dEnd#, dMinutes#, strErr$
    
    If bInprog Then Exit Sub
    
    bInprog = True
    tmrUpdateRect.Enabled = False
    
    If m.dTwipsPerMinute > 0# Then
        dStart = RoundNum(gdStartTime.Value * 1440#)
        dEnd = RoundNum(gdEndTime.Value * 1440#)
        If dStart <> dPrevStart Or dEnd <> dPrevEnd Or bReset Then
            If ValidTimes(dStart, dEnd, False, strErr) Then
                If dStart > m.dDefaultStart Then
                    shpRect.Left = Line1.X1 + (dStart - m.dDefaultStart) * m.dTwipsPerMinute
                    shpRect.Width = (Line1.X2 - Line1.X1) - ((dStart - m.dDefaultStart) * m.dTwipsPerMinute)
                ElseIf dStart < m.dDefaultStart Then
                    'this should only be true for overnight symbols
                    dMinutes = (1440# - m.dDefaultStart) + dStart    'minutes until midnight + minutes from midnight
                    shpRect.Left = Line1.X1 + dMinutes * m.dTwipsPerMinute
                    shpRect.Width = (Line1.X2 - Line1.X1) - (dMinutes * m.dTwipsPerMinute)
                Else
                    shpRect.Left = Line1.X1
                    shpRect.Width = Line1.X2 - Line1.X1
                End If
                If dEnd < m.dDefaultEnd Then
                    dMinutes = m.dDefaultEnd - dEnd
                    shpRect.Width = shpRect.Width - dMinutes * m.dTwipsPerMinute
                End If
                shpRect.FillColor = vbBlue
                lblRectError.Enabled = False
                lblRectError.Visible = False
            Else
                shpRect.Left = Line1.X1
                shpRect.Width = Line1.X2 - Line1.X1
                shpRect.FillColor = vbRed
                lblRectError.Enabled = True
                lblRectError.Visible = True
                lblRectError.Caption = strErr
            End If
        End If
        
        dPrevStart = dStart
        dPrevEnd = dEnd
    End If
    
    bInprog = False
    tmrUpdateRect.Enabled = True
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStartStopTimes.UpdateRect"
    
End Sub

Private Sub tmrUpdateRect_Timer()
    UpdateRect
End Sub

