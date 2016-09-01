VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmGameTargetLoss 
   Caption         =   " Replay Auto Exits"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   375
      Left            =   90
      TabIndex        =   9
      Top             =   786
      Width           =   4995
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
      Caption         =   "frmGameTargetLoss.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmGameTargetLoss.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmGameTargetLoss.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtTargetPoints 
         Height          =   285
         Left            =   1500
         TabIndex        =   0
         Top             =   75
         Width           =   915
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmGameTargetLoss.frx":0068
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmGameTargetLoss.frx":0090
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":00B0
      End
      Begin HexUniControls.ctlUniCheckXP chkTarget 
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   90
         Width           =   1935
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
         Caption         =   "frmGameTargetLoss.frx":00CC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmGameTargetLoss.frx":0106
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":0126
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtTargetDollars 
         Height          =   285
         Left            =   3420
         TabIndex        =   8
         Top             =   75
         Width           =   915
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmGameTargetLoss.frx":0142
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmGameTargetLoss.frx":016A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":018A
      End
      Begin HexUniControls.ctlUniLabelXP lblTargetPoints 
         Height          =   255
         Left            =   2460
         Top             =   90
         Width           =   1155
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
         Caption         =   "frmGameTargetLoss.frx":01A6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameTargetLoss.frx":01DC
         Style           =   0
         Enabled         =   0   'False
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":01FC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTargetDollars 
         Height          =   255
         Left            =   4440
         Top             =   90
         Width           =   675
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
         Caption         =   "frmGameTargetLoss.frx":0218
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameTargetLoss.frx":0246
         Style           =   0
         Enabled         =   0   'False
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":0266
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame2 
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   1238
      Width           =   4995
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
      Caption         =   "frmGameTargetLoss.frx":0282
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmGameTargetLoss.frx":02AE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmGameTargetLoss.frx":02CE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtStopPoints 
         Height          =   285
         Left            =   1500
         TabIndex        =   5
         Top             =   75
         Width           =   915
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmGameTargetLoss.frx":02EA
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmGameTargetLoss.frx":0312
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":0332
      End
      Begin HexUniControls.ctlUniCheckXP chkStop 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   90
         Width           =   1935
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
         Caption         =   "frmGameTargetLoss.frx":034E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmGameTargetLoss.frx":0380
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":03A0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtStopDollars 
         Height          =   285
         Left            =   3420
         TabIndex        =   4
         Top             =   75
         Width           =   915
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmGameTargetLoss.frx":03BC
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmGameTargetLoss.frx":03E4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":0404
      End
      Begin HexUniControls.ctlUniLabelXP lblStopPoints 
         Height          =   255
         Left            =   2460
         Top             =   90
         Width           =   1155
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
         Caption         =   "frmGameTargetLoss.frx":0420
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameTargetLoss.frx":0456
         Style           =   0
         Enabled         =   0   'False
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":0476
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStopDollars 
         Height          =   255
         Left            =   4440
         Top             =   90
         Width           =   675
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
         Caption         =   "frmGameTargetLoss.frx":0492
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameTargetLoss.frx":04C0
         Style           =   0
         Enabled         =   0   'False
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameTargetLoss.frx":04E0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2677
      TabIndex        =   2
      Top             =   1853
      Width           =   1095
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
      Caption         =   "frmGameTargetLoss.frx":04FC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmGameTargetLoss.frx":052A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmGameTargetLoss.frx":054A
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1447
      TabIndex        =   1
      Top             =   1853
      Width           =   1095
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
      Caption         =   "frmGameTargetLoss.frx":0566
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmGameTargetLoss.frx":058C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmGameTargetLoss.frx":05AC
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   555
      Left            =   90
      Top             =   66
      Width           =   4995
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
      Caption         =   "frmGameTargetLoss.frx":05C8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmGameTargetLoss.frx":068A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmGameTargetLoss.frx":06AA
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmGameTargetLoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    frmGame As Form
    oGameMode As cGameMode
    Bars As cGdBars
    FilledOrder As cPtOrder
    
    dMinMove As Double
    dTargetDollars As Double
    dTargetPoints As Double
    dStopDollars As Double
    dStopPoints As Double
    
    bStocks As Boolean
    bTDChanged As Boolean
    bTPChanged As Boolean
    bSDChanged As Boolean
    bSPChanged As Boolean
End Type
Private m As mPrivate

Public Sub ShowMe(GameModeObj As cGameMode, FormObj As Form, Order As cPtOrder)
    
    Set m.oGameMode = GameModeObj
    Set m.frmGame = FormObj
    Set m.Bars = m.frmGame.Chart.Bars
    Set m.FilledOrder = Order
    
    If Not m.Bars Is Nothing Then
        m.dMinMove = m.Bars.MinMove(m.oGameMode.GameDataTime)
        If m.Bars.Prop(eBARS_SecurityType) = 83 Then m.bStocks = True
    End If
        
    InitControls
        
    CenterTheForm Me
    ShowForm Me, eForm_Modal

End Sub

Private Sub chkStop_Click()

    Dim bEnable As Boolean
    
    bEnable = chkStop.Value

    txtStopDollars.Enabled = bEnable
    lblStopDollars.Enabled = bEnable
    
    txtStopPoints.Enabled = bEnable
    lblStopPoints.Enabled = bEnable

End Sub

Private Sub chkStop_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub chkTarget_Click()

    Dim bEnable As Boolean
    
    bEnable = chkTarget.Value

    txtTargetDollars.Enabled = bEnable
    lblTargetDollars.Enabled = bEnable
    
    txtTargetPoints.Enabled = bEnable
    lblTargetPoints.Enabled = bEnable

End Sub

Private Sub chkTarget_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub cmdOK_Click()
            
    Dim bRemoveStop As Boolean
    Dim bRemoveTarget As Boolean
            
    FixValues
    
    If chkTarget.Value = 1 Then
        m.oGameMode.SetTargetPrice m.dTargetPoints, m.dTargetDollars, m.FilledOrder
    ElseIf m.oGameMode.HasStop Then
        bRemoveTarget = True
    End If

    If chkStop.Value = 1 Then
        m.oGameMode.SetStopPrice m.dStopPoints, m.dStopDollars, m.FilledOrder
    ElseIf m.oGameMode.HasTarget Then
        bRemoveStop = True
    End If
    
    m.oGameMode.RemoveAutoExits bRemoveStop, bRemoveTarget
    m.oGameMode.HasTarget = chkTarget.Value
    m.oGameMode.HasStop = chkStop.Value
    
    m.frmGame.Chart.GenerateChart eRedo1_Scrolled
    
    Unload Me
    
End Sub

Private Sub InitControls()
                        
    If m.oGameMode.TargetDollars = 0 Then
        m.dTargetDollars = m.oGameMode.AutoTargetDollars
        m.dTargetPoints = m.oGameMode.AutoTargetPoints
    Else
        m.dTargetDollars = m.oGameMode.TargetDollars
        m.dTargetPoints = m.oGameMode.TargetPoints
    End If
    
    If m.oGameMode.StopDollars = 0 Then
        m.dStopDollars = m.oGameMode.AutoStopDollars
        m.dStopPoints = m.oGameMode.AutoStopPoints
    Else
        m.dStopDollars = m.oGameMode.StopDollars
        m.dStopPoints = m.oGameMode.StopPoints
    End If
    
    If Not m.Bars Is Nothing Then
        txtTargetPoints = m.Bars.PriceDisplay(m.dTargetPoints)
        txtStopPoints = m.Bars.PriceDisplay(m.dStopPoints)
    Else
        txtTargetPoints = Str(m.dTargetPoints)
        txtStopPoints = Str(m.dStopPoints)
    End If
            
    txtTargetDollars.Visible = Not m.bStocks
    txtStopDollars.Visible = Not m.bStocks
    lblTargetDollars.Visible = Not m.bStocks
    lblStopDollars.Visible = Not m.bStocks
    
    If m.bStocks Then
        lblTargetPoints = "points"
        lblStopPoints = "points"
    Else
        lblTargetPoints = "points   ="
        lblStopPoints = "points   ="
        txtTargetDollars = Str(m.dTargetDollars)
        txtStopDollars = Str(m.dStopDollars)
    End If
    
    chkTarget.Value = m.oGameMode.HasTarget
    chkStop.Value = m.oGameMode.HasStop
    
    m.bTDChanged = False
    m.bTPChanged = False
    m.bSDChanged = False
    m.bTPChanged = False
    
End Sub

Private Sub FixValues()
           
    If m.Bars Is Nothing Then Exit Sub
    
    If m.bTDChanged Or m.bTPChanged Then FixTarget
    If m.bSDChanged Or m.bSPChanged Then FixStop
    
End Sub

Private Sub FixStop()

    Dim dTicks#, dDollars#, dPoints#
    Dim dNewDollars#, dNewPoints#
        
    dDollars = ValOfText(txtStopDollars)
    dPoints = m.Bars.PriceFromString(txtStopPoints)
    
    If m.bSDChanged Then
        'calculate ticks
        dTicks = dDollars / m.Bars.Prop(eBARS_TickValue)
    Else
        dNewPoints = dPoints
        'calculate ticks
        dTicks = dNewPoints / m.Bars.Prop(eBARS_TickMove)
    End If
    
    If dTicks > 0 Then
        dTicks = RoundNum(dTicks, 0)
        'convert to points
        dNewPoints = dTicks * m.Bars.Prop(eBARS_TickMove)
        'make dollar values match points
        dNewDollars = dTicks * m.Bars.Prop(eBARS_TickValue)
        m.dStopDollars = dNewDollars
        m.dStopPoints = dNewPoints
        txtStopDollars = Str(m.dStopDollars)
        txtStopPoints = m.Bars.PriceDisplay(m.dStopPoints)
    End If
    
    m.bSDChanged = False
    m.bSPChanged = False

End Sub

Private Sub FixTarget()

    Dim dTicks#, dDollars#, dPoints#
    Dim dNewDollars#, dNewPoints#
        
    dDollars = ValOfText(txtTargetDollars)
    dPoints = m.Bars.PriceFromString(txtTargetPoints)
    
    If m.bTDChanged Then
        'calculate ticks
        dTicks = dDollars / m.Bars.Prop(eBARS_TickValue)
    Else
        dNewPoints = dPoints
        'calculate ticks
        dTicks = dNewPoints / m.Bars.Prop(eBARS_TickMove)
    End If
    
    If dTicks > 0 Then
        dTicks = RoundNum(dTicks, 0)
        'convert to points
        dNewPoints = dTicks * m.Bars.Prop(eBARS_TickMove)
        'make dollar values match points
        dNewDollars = dTicks * m.Bars.Prop(eBARS_TickValue)
        m.dTargetDollars = dNewDollars
        m.dTargetPoints = dNewPoints
        txtTargetDollars = Str(m.dTargetDollars)
        txtTargetPoints = m.Bars.PriceDisplay(m.dTargetPoints)
    End If
    
    m.bTDChanged = False
    m.bTPChanged = False

End Sub

Private Sub cmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub Form_Load()
    Me.Icon = Picture16(ToolbarIcon("ID_Replay"))
    
    g.Styler.StyleForm Me
End Sub

Private Sub txtStopDollars_Change()
    If ValOfText(txtStopDollars) <> m.dStopDollars Then
        m.bSDChanged = True
    End If
End Sub

Private Sub txtStopDollars_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub txtStopPoints_Change()
    If m.Bars.PriceFromString(txtStopPoints) <> m.dStopPoints Then
        m.bSPChanged = True
    End If
End Sub

Private Sub txtStopDollars_LostFocus()
    FixValues
End Sub

Private Sub txtStopPoints_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub txtStopPoints_LostFocus()
    FixValues
End Sub

Private Sub txtTargetDollars_Change()
    If m.dTargetDollars <> ValOfText(txtTargetDollars) Then
        m.bTDChanged = True
    End If
End Sub

Private Sub txtTargetDollars_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub txtTargetPoints_Change()
    If m.Bars.PriceFromString(txtTargetPoints) <> m.dTargetPoints Then
        m.bTPChanged = True
    End If
End Sub

Private Sub txtTargetDollars_LostFocus()
    FixValues
End Sub

Private Sub txtTargetPoints_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub txtTargetPoints_LostFocus()
    FixValues
End Sub

Private Sub ShowHelp(KeyCode As Integer)
On Error Resume Next:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Nothing
    End If

End Sub

