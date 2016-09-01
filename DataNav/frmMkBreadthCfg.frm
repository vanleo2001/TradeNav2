VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmMkBreadthCfg 
   Caption         =   "Market Breadth Settings"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   375
      Left            =   3128
      TabIndex        =   6
      Top             =   2895
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
      Caption         =   "frmMkBreadthCfg.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMkBreadthCfg.frx":002C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMkBreadthCfg.frx":004C
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Height          =   375
      Left            =   1688
      TabIndex        =   5
      Top             =   2895
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
      Caption         =   "frmMkBreadthCfg.frx":0068
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMkBreadthCfg.frx":008C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMkBreadthCfg.frx":00AC
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraSummary 
      Height          =   2535
      Left            =   98
      TabIndex        =   0
      Top             =   186
      Width           =   5715
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
      Caption         =   "frmMkBreadthCfg.frx":00C8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMkBreadthCfg.frx":0100
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMkBreadthCfg.frx":0120
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraValuesStyle 
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   5475
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
         Caption         =   "frmMkBreadthCfg.frx":013C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMkBreadthCfg.frx":01A4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadthCfg.frx":01C4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optValueStyle 
            Height          =   220
            Index           =   2
            Left            =   4320
            TabIndex        =   1
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmMkBreadthCfg.frx":01E0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMkBreadthCfg.frx":0208
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMkBreadthCfg.frx":0228
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optValueStyle 
            Height          =   220
            Index           =   1
            Left            =   2460
            TabIndex        =   2
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
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
            Caption         =   "frmMkBreadthCfg.frx":0244
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMkBreadthCfg.frx":027E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMkBreadthCfg.frx":029E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optValueStyle 
            Height          =   220
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "frmMkBreadthCfg.frx":02BA
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMkBreadthCfg.frx":0308
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMkBreadthCfg.frx":0328
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin gdOCX.gdSelectColor gdAdvancedColor 
         Height          =   315
         Left            =   3300
         TabIndex        =   3
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdDeclinedColor 
         Height          =   315
         Left            =   3300
         TabIndex        =   4
         Top             =   840
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   795
         Top             =   960
         Width           =   3375
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
         Caption         =   "frmMkBreadthCfg.frx":0344
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMkBreadthCfg.frx":03A4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadthCfg.frx":03C4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   795
         Top             =   480
         Width           =   3375
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
         Caption         =   "frmMkBreadthCfg.frx":03E0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMkBreadthCfg.frx":0440
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadthCfg.frx":0460
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmMkBreadthCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    oData As cMkBreadth
End Type

Private m As mPrivate

Public Sub ShowMe(oMkData As cMkBreadth)
    
    If Not oMkData Is Nothing Then
        Set m.oData = oMkData
        gdAdvancedColor.Color = m.oData.AdvancedColor
        gdDeclinedColor.Color = m.oData.DeclinedColor
        optValueStyle(m.oData.ValueStyle).Value = True
    End If
        
    CenterTheForm Me
    ShowForm Me, True

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim i&
    
    m.oData.AdvancedColor = gdAdvancedColor.Color
    m.oData.DeclinedColor = gdDeclinedColor.Color
        
    For i = 0 To 2
        If optValueStyle(i).Value = True Then
            m.oData.ValueStyle = i
            Exit For
        End If
    Next
    
    Unload Me
        
End Sub

