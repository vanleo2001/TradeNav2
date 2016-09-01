VERSION 5.00
Begin VB.Form frmPatternProfitOpt 
   Caption         =   "Optimization Parameters"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
Begin HexUniControls.ctlUniButtonImageXP cmdClose
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3907
      TabIndex        =   14
      Top             =   870
      Width           =   975
   End
Begin HexUniControls.ctlUniFrameWL fraCorrelationFit
VistaStyle      =   0   'False
      Caption         =   "Correlation Fit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   652
      TabIndex        =   7
      Top             =   1350
      Width           =   4230
Begin HexUniControls.ctlUniTextBoxXP txtMinHits
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1905
         TabIndex        =   9
         Text            =   "10"
         Top             =   345
         Width           =   495
      End
Begin HexUniControls.ctlUniTextBoxXP txtMinCorr
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2025
         TabIndex        =   8
         Text            =   "80"
         Top             =   705
         Width           =   375
      End
Begin HexUniControls.ctlUniLabelXP Label3
         Appearance      =   0  'Flat
         Caption         =   "&Ignore when less than"
         Height          =   255
         Left            =   225
         TabIndex        =   13
         Top             =   375
         Width           =   1695
      End
Begin HexUniControls.ctlUniLabelXP Label4
         Appearance      =   0  'Flat
         Caption         =   "matches found."
         Height          =   255
         Left            =   2505
         TabIndex        =   12
         Top             =   375
         Width           =   1215
      End
Begin HexUniControls.ctlUniLabelXP Label6
         Appearance      =   0  'Flat
         Caption         =   "&Do not test for less than"
         Height          =   255
         Left            =   225
         TabIndex        =   11
         Top             =   735
         Width           =   1815
      End
Begin HexUniControls.ctlUniLabelXP Label7
         Appearance      =   0  'Flat
         Caption         =   "percent fit."
         Height          =   255
         Left            =   2505
         TabIndex        =   10
         Top             =   735
         Width           =   1155
      End
   End
Begin HexUniControls.ctlUniButtonImageXP cmdOK
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3907
      TabIndex        =   6
      Top             =   390
      Width           =   975
   End
Begin HexUniControls.ctlUniFrameWL fraPtrnLen
VistaStyle      =   0   'False
      Caption         =   "Length of Pattern"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   667
      TabIndex        =   0
      Top             =   345
      Width           =   2955
Begin HexUniControls.ctlUniTextBoxXP txtMinBars
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1035
         TabIndex        =   2
         Text            =   "2"
         Top             =   330
         Width           =   375
      End
Begin HexUniControls.ctlUniTextBoxXP txtMaxBars
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1830
         TabIndex        =   1
         Text            =   "6"
         Top             =   330
         Width           =   375
      End
Begin HexUniControls.ctlUniLabelXP Label5
         Appearance      =   0  'Flat
         Caption         =   "bars."
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
Begin HexUniControls.ctlUniLabelXP Label2
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "to"
         Height          =   255
         Left            =   1470
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
Begin HexUniControls.ctlUniLabelXP Label1
         Appearance      =   0  'Flat
         Caption         =   "&Test from"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmPatternProfitOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    frm As Form
    oPFP As cPatternProfit
    bOK As Boolean
End Type

Private m As mPrivate

Public Function ShowMe(frm As Form) As Boolean
On Error GoTo ErrSection:
    
    If frm Is Nothing Then Exit Function
    
    Set m.frm = frm
    
    m.bOK = False
    
    If TypeOf frm Is frmPatternProfit Then
        txtMinBars.Text = frm.OptimizePtrnLen
        txtMaxBars.Text = frm.OptimizeMaxBars
        txtMinCorr.Text = frm.OptimizeLowestCorr
        txtMinHits.Text = frm.OptimizeMinHits
        Me.Icon = Picture16(ToolbarIcon("ID_PatternProfit"), , True)
    ElseIf IsFrmChart(frm) Then
        Set m.oPFP = frm.PatternProfitObj
        If m.oPFP Is Nothing Then GoTo ErrExit
        
        txtMinBars.Text = m.oPFP.OptimizeMinBars
        txtMaxBars.Text = m.oPFP.OptimizeMaxBars
        txtMinCorr.Text = m.oPFP.OptimizeLowestCorr
        txtMinHits.Text = m.oPFP.OptimizeMinHits
        Me.Icon = Picture16(ToolbarIcon("ID_IndAnalyst"), , True)     '5888
    Else
        GoTo ErrExit            'theoretically should never get here
    End If
    
    CenterTheForm Me
    ShowForm Me, eForm_Modal, frmMain
    
    ShowMe = m.bOK

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPatternProfitOpt.ShowMe"

End Function

Private Sub cmdClose_Click()
On Error GoTo ErrSection:
    
    If IsFrmChart(m.frm) Then m.frm.cmdOptimizePFP.Enabled = True
    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfitOpt.cmdClose_Click"

End Sub


Private Function ValidateParams() As Boolean
On Error GoTo ErrSection:

    Dim iMin&, iMax&, iMinHits&, iMinCorr&
    Dim bOkay As Boolean
    
    iMin = Val(txtMinBars.Text)
    iMax = Val(txtMaxBars.Text)
    
    iMinHits = Val(txtMinHits.Text)
    iMinCorr = Val(txtMinCorr.Text)
    
    If iMin <= 0 Or iMin > iMax Then
        InfBox "'Test from' bars must be greater than zero and less than 'test to' bars.", "I"
    ElseIf iMax <= 0 Or iMax > 260 Then
        InfBox "'Test to' bars must be greater than zero and less than or equal to 260.", "I"
    ElseIf iMinHits < 0 Then
        InfBox "Matches to ignore must be greater than zero.", "I"
    ElseIf iMinCorr <= 0 Or iMinCorr > 100 Then
        InfBox "Percent correlation must be between 1 and 100."
    Else
        bOkay = True
    End If
    
    ValidateParams = bOkay

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPatternProfitOpt.ValidateParams"

End Function

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If ValidateParams Then
        Me.Hide
        DoEvents
        If TypeOf m.frm Is frmPatternProfit Then
            If Not m.frm Is Nothing Then m.frm.Optimize
        ElseIf Not m.oPFP Is Nothing Then
            m.oPFP.OptimizeMinBars = Val(frmPatternProfitOpt.txtMinBars)
            m.oPFP.OptimizeMaxBars = Val(frmPatternProfitOpt.txtMaxBars)
            m.oPFP.OptimizeMinHits = Val(frmPatternProfitOpt.txtMinHits)
            m.oPFP.OptimizeLowestCorr = Val(frmPatternProfitOpt.txtMinCorr)
        End If
        m.bOK = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfitOpt.cmdOK_Click"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.frm = Nothing
    Set m.oPFP = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfitOpt.Form_Unload"

End Sub

