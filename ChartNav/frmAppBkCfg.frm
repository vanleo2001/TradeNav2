VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAppBkCfg 
   Caption         =   "Application Background"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   540
      Left            =   1955
      TabIndex        =   9
      Top             =   2040
      Width           =   2100
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
      Caption         =   "frmAppBkCfg.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAppBkCfg.frx":0034
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAppBkCfg.frx":0054
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   15
         TabIndex        =   11
         Top             =   120
         Width           =   915
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
         Caption         =   "frmAppBkCfg.frx":0070
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAppBkCfg.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   1095
         TabIndex        =   12
         Top             =   120
         Width           =   915
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
         Caption         =   "frmAppBkCfg.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAppBkCfg.frx":0100
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":0120
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraLogo 
      Height          =   1935
      Left            =   3105
      TabIndex        =   4
      Top             =   120
      Width           =   2775
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
      Caption         =   "frmAppBkCfg.frx":013C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAppBkCfg.frx":0164
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAppBkCfg.frx":0184
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtLogoSize 
         Height          =   375
         Left            =   1875
         TabIndex        =   5
         Top             =   1335
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAppBkCfg.frx":01A0
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
         Tip             =   "frmAppBkCfg.frx":01CA
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":01EA
      End
      Begin gdOCX.gdSelectColor gdLogoColor 
         Height          =   375
         Left            =   1620
         TabIndex        =   6
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Color           =   12632256
         CustomColor     =   12632256
      End
      Begin HexUniControls.ctlUniRadioXP optCustomColor 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   780
         Width           =   1575
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
         Caption         =   "frmAppBkCfg.frx":0206
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAppBkCfg.frx":0240
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":0260
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optDarkShadow 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   510
         Width           =   1575
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
         Caption         =   "frmAppBkCfg.frx":027C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAppBkCfg.frx":02B2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":02D2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLightShadow 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1575
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
         Caption         =   "frmAppBkCfg.frx":02EE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAppBkCfg.frx":0326
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":0346
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   255
         Left            =   240
         Top             =   1500
         Width           =   1575
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
         Caption         =   "frmAppBkCfg.frx":0362
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAppBkCfg.frx":03A6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":03C6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   240
         Top             =   1260
         Width           =   1575
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
         Caption         =   "frmAppBkCfg.frx":03E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAppBkCfg.frx":0420
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":0440
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraDots 
      Height          =   1935
      Left            =   146
      TabIndex        =   0
      Top             =   120
      Width           =   2775
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
      Caption         =   "frmAppBkCfg.frx":045C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAppBkCfg.frx":048E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAppBkCfg.frx":04AE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtDotSpacing 
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   1320
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAppBkCfg.frx":04CA
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
         Tip             =   "frmAppBkCfg.frx":04F4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":0514
      End
      Begin gdOCX.gdSelectColor gdDotColor 
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Color           =   0
         CustomColor     =   0
      End
      Begin HexUniControls.ctlUniRadioXP optDotLarge 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   855
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
         Caption         =   "frmAppBkCfg.frx":0530
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAppBkCfg.frx":055A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":057A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optDotSmall 
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   855
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
         Caption         =   "frmAppBkCfg.frx":0596
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAppBkCfg.frx":05C0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":05E0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   240
         Top             =   1380
         Width           =   2295
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
         Caption         =   "frmAppBkCfg.frx":05FC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAppBkCfg.frx":0644
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":0664
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   255
         Left            =   240
         Top             =   420
         Width           =   735
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
         Caption         =   "frmAppBkCfg.frx":0680
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAppBkCfg.frx":06AC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":06CC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   240
         Top             =   900
         Width           =   855
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
         Caption         =   "frmAppBkCfg.frx":06E8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAppBkCfg.frx":0712
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAppBkCfg.frx":0732
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmAppBkCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kRegWidth = 6135
Private Const kBetterTradesWidth = 3230
Private Const kRegLeft = 1955
Private Const kBetterTradesLeft = 500

Private Type mPrivate
    nDotColor As Long
    nLogoColor As Long
    nDotSize As Long
    nLogoSize As Long
    nDotSpacing As Long
End Type

Private m As mPrivate

Public Sub ShowMe()

    m.nDotColor = ValOfText(GetIniFileProperty("DotColor", 0, "AppBitmap", g.strIniFile))
    m.nLogoColor = ValOfText(GetIniFileProperty("LogoColor", 0, "AppBitmap", g.strIniFile))
    m.nDotSize = ValOfText(GetIniFileProperty("DotSize", 1, "AppBitmap", g.strIniFile))
    m.nLogoSize = ValOfText(GetIniFileProperty("LogoSize", 40, "AppBitmap", g.strIniFile))
    m.nDotSpacing = ValOfText(GetIniFileProperty("DotPixSpace", 40, "AppBitmap", g.strIniFile))
    
'form dimension
    If ExtremeCharts >= 1 Then
        Me.Width = kBetterTradesWidth
        fraButtons.Left = kBetterTradesLeft
    Else
        Me.Width = kRegWidth
        fraButtons.Left = kRegLeft
    End If
    
'colors
    If m.nDotColor < 0 Then m.nDotColor = 0
    gdDotColor.Color = m.nDotColor
    
    If m.nLogoColor = -1 Then
        optLightShadow.Value = True
        gdLogoColor.Visible = False
    ElseIf m.nLogoColor > 0 Then
        optCustomColor = True
        gdLogoColor.Visible = True
        gdLogoColor.Color = m.nLogoColor
    Else
        optDarkShadow.Value = True
        gdLogoColor.Visible = False
    End If
    
'sizes
    If m.nDotSize = 1 Then
        optDotLarge.Value = True
    Else
        optDotSmall.Value = True
    End If
    
    If m.nLogoSize > 155 Then
        txtLogoSize.Text = 155
        m.nLogoSize = 155
    ElseIf m.nLogoSize >= 0 Then
        txtLogoSize.Text = m.nLogoSize
    Else
        txtLogoSize.Text = 0
        m.nLogoSize = 0
    End If
    
'dots spacing
    If m.nDotSpacing >= 0 Then
        txtDotSpacing.Text = m.nDotSpacing
    Else
        txtDotSpacing.Text = 0
        m.nDotSpacing = 0
    End If

    CenterTheForm Me
    ShowForm Me, True

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim iDotSizeNew&, iDotColorNew&
    Dim iLogoSizeNew&, iLogoColorNew&
    Dim iSpacingNew&, i&, j&, s$
    
    iSpacingNew = Int(ValOfText(Trim(txtDotSpacing.Text)))
    iDotColorNew = gdDotColor.Color
    If optDotLarge.Value = True Then iDotSizeNew = 1
    
    iLogoSizeNew = Int(ValOfText(Trim(txtLogoSize.Text)))
    If optLightShadow.Value = True Then
        iLogoColorNew = -1
    ElseIf optCustomColor.Value = True Then
        iLogoColorNew = gdLogoColor.Color
    Else
        iLogoColorNew = 0
    End If
    
    If iDotSizeNew <> m.nDotSize Or iDotColorNew <> m.nDotColor Or _
        iLogoSizeNew <> m.nLogoSize Or iLogoColorNew <> m.nLogoColor Or _
        iSpacingNew <> m.nDotSpacing Then
        
        SetIniFileProperty "DotSize", iDotSizeNew, "AppBitmap", g.strIniFile
        SetIniFileProperty "DotColor", iDotColorNew, "AppBitmap", g.strIniFile
        SetIniFileProperty "LogoSize", iLogoSizeNew, "AppBitmap", g.strIniFile
        SetIniFileProperty "LogoColor", iLogoColorNew, "AppBitmap", g.strIniFile
        SetIniFileProperty "DotPixSpace", iSpacingNew, "AppBitmap", g.strIniFile
        
        If g.ChartGlobals.bSnapToDots And iLogoSizeNew <> m.nLogoSize Then
            SetIniFileProperty "LogoSizeExplicit", 1, "AppBitmap", g.strIniFile
        End If
        
        LoadAppBkImage True
    End If
    
ErrExit:
    Unload Me
    Exit Sub

ErrSection:
    RaiseError "frmAppBkCfg.cmdOk_Click"
    
End Sub

Private Sub Form_Load()

    Me.Icon = Picture16("kBlank")
    CenterTheForm Me
    
    g.Styler.StyleForm Me

End Sub

Private Sub optCustomColor_Click()
    gdLogoColor.Visible = True
End Sub

Private Sub optDarkShadow_Click()
    gdLogoColor.Visible = False
End Sub

Private Sub optLightShadow_Click()
    gdLogoColor.Visible = False
End Sub

