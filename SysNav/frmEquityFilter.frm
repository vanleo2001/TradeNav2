VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.0#0"; "HexUniControls42.ocx"
Begin VB.Form frmEquityFilter 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraEquityFilter 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Width           =   4575
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
      Caption         =   "frmEquityFilter.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEquityFilter.frx":003C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEquityFilter.frx":005C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optIgnore 
         Height          =   220
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEquityFilter.frx":0078
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEquityFilter.frx":00C4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":00E4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optTakeAll 
         Height          =   220
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEquityFilter.frx":0100
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEquityFilter.frx":0142
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":0162
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraIgnoreTrades 
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   840
         Width           =   4035
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
         Caption         =   "frmEquityFilter.frx":017E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEquityFilter.frx":01AA
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":01CA
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optFalling 
            Height          =   220
            Left            =   0
            TabIndex        =   12
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmEquityFilter.frx":01E6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEquityFilter.frx":0242
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEquityFilter.frx":0262
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optBelow 
            Height          =   220
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmEquityFilter.frx":027E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmEquityFilter.frx":0304
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEquityFilter.frx":0324
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniLabelXP lblNote 
         Height          =   1215
         Left            =   180
         Top             =   1440
         Width           =   4215
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
         Caption         =   "frmEquityFilter.frx":0340
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEquityFilter.frx":05EE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":060E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraMovingAverage 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
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
      Caption         =   "frmEquityFilter.frx":062A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEquityFilter.frx":067C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEquityFilter.frx":069C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtMovingAveragePeriod 
         Height          =   315
         Left            =   2820
         TabIndex        =   4
         Top             =   270
         Width           =   615
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEquityFilter.frx":06B8
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
         Alignment       =   2
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmEquityFilter.frx":06DC
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":06FC
      End
      Begin HexUniControls.ctlUniComboImageXP cboMovingAverageType 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   270
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   -2147483630
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
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
         Tip             =   "frmEquityFilter.frx":0718
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":0738
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdScrollBar sbMovingAveragePeriod 
         Height          =   360
         Left            =   3420
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   247
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
      Begin HexUniControls.ctlUniLabelXP lblTrades 
         Height          =   195
         Left            =   3720
         Top             =   330
         Width           =   555
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
         Caption         =   "frmEquityFilter.frx":0754
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEquityFilter.frx":0780
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":07A0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblMovingAveragePeriod 
         Height          =   255
         Left            =   2280
         Top             =   300
         Width           =   555
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
         Caption         =   "frmEquityFilter.frx":07BC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEquityFilter.frx":07EC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":080C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblMovingAverageType 
         Height          =   255
         Left            =   180
         Top             =   300
         Width           =   555
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
         Caption         =   "frmEquityFilter.frx":0828
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEquityFilter.frx":0854
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":0874
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1095
      Left            =   4860
      TabIndex        =   1
      Top             =   120
      Width           =   1215
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
      Caption         =   "frmEquityFilter.frx":0890
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEquityFilter.frx":08BC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEquityFilter.frx":08DC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   1215
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
         Caption         =   "frmEquityFilter.frx":08F8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEquityFilter.frx":0926
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":0946
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1215
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
         Caption         =   "frmEquityFilter.frx":0962
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEquityFilter.frx":0988
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEquityFilter.frx":09A8
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmEquityFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmEquityFilter.frm
'' Description: Allow the user to modify equity filter options
''
'' Author:      Genesis Financial Data Services
''              425 Wind Chime Pl
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    
    MovingAveragePeriod As cPriceEditor ' Price editor class for moving average period
    
    EquityFilter As cEquityFilter       ' Equity filter information
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(EquityFilter As cEquityFilter, imgIcon As Picture) As Boolean
On Error GoTo ErrSection:
    
    Icon = imgIcon
    Caption = "Equity Filter Options"
    
    Set m.EquityFilter = EquityFilter
    
    SetMovingAverageTypeCombo EquityFilter.MovingAverageType
    Set m.MovingAveragePeriod = New cPriceEditor
    m.MovingAveragePeriod.Init sbMovingAveragePeriod, txtMovingAveragePeriod, Nothing, EquityFilter.MovingAveragePeriod, 1
    If m.EquityFilter.EquityFilterOn = False Then
        optTakeAll.Value = True
        optIgnore.Value = False
    Else
        optTakeAll.Value = False
        optIgnore.Value = True
    End If
    If EquityFilter.EquityFilterMode = eGDEquityFilterMode_BelowMa Then
        optBelow.Value = True
        optFalling.Value = False
    Else
        optBelow.Value = False
        optFalling.Value = True
    End If
    
    EnableControls

    ShowForm Me, eForm_Modal
    
    If m.bOK Then
        EquityFilter.MovingAverageType = cboMovingAverageType.Text
        EquityFilter.MovingAveragePeriod = m.MovingAveragePeriod.Price
        EquityFilter.EquityFilterOn = optIgnore.Value
        If optBelow.Value = True Then
            EquityFilter.EquityFilterMode = eGDEquityFilterMode_BelowMa
        Else
            EquityFilter.EquityFilterMode = eGDEquityFilterMode_MaDown
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmEquityFilter.ShowMe", , g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the ShowMe to unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEquityFilter.cmdCancel_Click", , g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allow the ShowMe to unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEquityFilter.cmdOK_Click", , g.strAppPath

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

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    cboMovingAverageType.AddItem "Simple"
    cboMovingAverageType.AddItem "Exponential"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEquityFilter.Form_Load", , g.strAppPath
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, allow the ShowMe to unload form
'' Inputs:      Cancel Unload?, Unload Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEquityFilter.Form_QueryUnload", , g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optIgnore_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optIgnore_Click()
On Error GoTo ErrSection:

    If Visible Then
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEquityFilter.optIgnore_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTakeAll_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTakeAll_Click()
On Error GoTo ErrSection:

    If Visible Then
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEquityFilter.optTakeAll_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetMovingAverageTypeCombo
'' Description: Attempt to set the combo to the given value
'' Inputs:      Moving Average Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetMovingAverageTypeCombo(ByVal strMovingAverageType As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found in the combo?
    
    bFound = False
    With cboMovingAverageType
        For lIndex = 0 To .ListCount - 1
            If UCase(.List(lIndex)) = UCase(strMovingAverageType) Then
                .ListIndex = lIndex
                bFound = True
            End If
        Next lIndex
    End With
    
    If bFound = False Then
        cboMovingAverageType.ListIndex = 0&
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEquityFilter.SetMovingAverageTypeCombo", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bEnableIgnore As Boolean        ' Enable ignore filter controls?
    
    bEnableIgnore = optIgnore.Value
    
    Enable optBelow, bEnableIgnore
    Enable optFalling, bEnableIgnore
    Enable lblNote, bEnableIgnore

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEquityFilter.EnableControls", , g.strAppPath
    
End Sub


