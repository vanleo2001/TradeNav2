VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmChartOrdBar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trading Account Settings"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniCheckXP chkHighlightEquity 
      Height          =   255
      Left            =   631
      TabIndex        =   30
      Top             =   1155
      Width           =   2355
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
      Caption         =   "frmChartOrdBar.frx":0000
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmChartOrdBar.frx":004A
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":006A
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkHighlightPos 
      Height          =   255
      Left            =   631
      TabIndex        =   28
      Top             =   845
      Width           =   2355
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
      Caption         =   "frmChartOrdBar.frx":0086
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmChartOrdBar.frx":00DA
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":00FA
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkAvgEntryLine 
      Height          =   255
      Left            =   623
      TabIndex        =   16
      Top             =   1515
      Width           =   2205
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
      Caption         =   "frmChartOrdBar.frx":0116
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmChartOrdBar.frx":0164
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":0184
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraAvgEntryLine 
      Height          =   1860
      Left            =   623
      TabIndex        =   17
      Top             =   1560
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
      Caption         =   "frmChartOrdBar.frx":01A0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmChartOrdBar.frx":01DE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":01FE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraAvgEntryTextPos 
         Height          =   1080
         Left            =   195
         TabIndex        =   20
         Top             =   720
         Width           =   4200
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
         Caption         =   "frmChartOrdBar.frx":021A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmChartOrdBar.frx":0252
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmChartOrdBar.frx":0272
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdFont 
            Height          =   360
            Left            =   2235
            TabIndex        =   25
            Top             =   570
            Width           =   1560
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
            Caption         =   "frmChartOrdBar.frx":028E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmChartOrdBar.frx":02B6
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmChartOrdBar.frx":02D6
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTextOnRight 
            Height          =   255
            Left            =   2235
            TabIndex        =   24
            Top             =   225
            Width           =   1530
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
            Caption         =   "frmChartOrdBar.frx":02F2
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmChartOrdBar.frx":0336
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmChartOrdBar.frx":0356
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optTextPosition 
            Height          =   285
            Index           =   1
            Left            =   270
            TabIndex        =   23
            Top             =   488
            Width           =   1470
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
            Caption         =   "frmChartOrdBar.frx":0372
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmChartOrdBar.frx":03B2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmChartOrdBar.frx":03D2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optTextPosition 
            Height          =   285
            Index           =   2
            Left            =   270
            TabIndex        =   22
            Top             =   750
            Width           =   1470
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
            Caption         =   "frmChartOrdBar.frx":03EE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmChartOrdBar.frx":0422
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmChartOrdBar.frx":0442
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optTextPosition 
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   21
            Top             =   225
            Width           =   1470
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
            Caption         =   "frmChartOrdBar.frx":045E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmChartOrdBar.frx":0492
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmChartOrdBar.frx":04B2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniComboImageXP cboLineStyle 
         Height          =   315
         Left            =   2115
         TabIndex        =   19
         Top             =   315
         Width           =   2280
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
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
         Tip             =   "frmChartOrdBar.frx":04CE
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmChartOrdBar.frx":04EE
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor gdColor 
         Height          =   315
         Left            =   195
         TabIndex        =   18
         Top             =   315
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         CustomColor     =   255
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkOrderLine 
      Height          =   255
      Left            =   631
      TabIndex        =   15
      Top             =   525
      Width           =   3930
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
      Caption         =   "frmChartOrdBar.frx":050A
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmChartOrdBar.frx":057A
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":059A
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboLimitPeriod 
      Height          =   315
      Left            =   3638
      TabIndex        =   14
      Top             =   3585
      Width           =   1815
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
      Tip             =   "frmChartOrdBar.frx":05B6
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":05D6
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtLimitTrades 
      Height          =   315
      Left            =   2798
      TabIndex        =   13
      Top             =   3585
      Width           =   555
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmChartOrdBar.frx":05F2
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
      Tip             =   "frmChartOrdBar.frx":0616
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":0636
   End
   Begin HexUniControls.ctlUniCheckXP chkLimitTrades 
      Height          =   255
      Left            =   413
      TabIndex        =   12
      Top             =   3615
      Width           =   2535
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
      Caption         =   "frmChartOrdBar.frx":0652
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmChartOrdBar.frx":06A4
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":06C4
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkShowAcctBar 
      Height          =   255
      Left            =   3016
      TabIndex        =   11
      Top             =   195
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
      Caption         =   "frmChartOrdBar.frx":06E0
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmChartOrdBar.frx":0720
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":0740
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkShowOrdBar 
      Height          =   255
      Left            =   631
      TabIndex        =   10
      Top             =   195
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
      Caption         =   "frmChartOrdBar.frx":075C
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmChartOrdBar.frx":0798
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":07B8
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtQty3 
      Height          =   315
      Left            =   4223
      TabIndex        =   8
      Top             =   4050
      Width           =   555
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmChartOrdBar.frx":07D4
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
      Tip             =   "frmChartOrdBar.frx":07F6
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":0816
   End
   Begin HexUniControls.ctlUniTextBoxXP txtQty2 
      Height          =   315
      Left            =   3623
      TabIndex        =   7
      Top             =   4050
      Width           =   555
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmChartOrdBar.frx":0832
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
      Tip             =   "frmChartOrdBar.frx":0854
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":0874
   End
   Begin HexUniControls.ctlUniTextBoxXP txtQty1 
      Height          =   315
      Left            =   2963
      TabIndex        =   6
      Top             =   4050
      Width           =   555
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmChartOrdBar.frx":0890
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
      Tip             =   "frmChartOrdBar.frx":08B2
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":08D2
   End
   Begin HexUniControls.ctlUniFrameWL fraShowButtons 
      Height          =   3795
      Left            =   53
      TabIndex        =   5
      Top             =   4515
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
      Caption         =   "frmChartOrdBar.frx":08EE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmChartOrdBar.frx":0946
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":0966
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fgOrderButtons 
         Height          =   3345
         Left            =   180
         TabIndex        =   27
         Top             =   330
         Width           =   2355
         _cx             =   4154
         _cy             =   5900
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
   End
   Begin HexUniControls.ctlUniFrameWL fraAcctBarFields 
      Height          =   3795
      Left            =   2933
      TabIndex        =   0
      Top             =   4515
      Width           =   2895
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
      Caption         =   "frmChartOrdBar.frx":0982
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmChartOrdBar.frx":09DC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":09FC
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fgAcctCols 
         Height          =   3345
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2355
         _cx             =   4154
         _cy             =   5900
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
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   923
      TabIndex        =   2
      Top             =   8400
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
      Caption         =   "frmChartOrdBar.frx":0A18
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmChartOrdBar.frx":0A38
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":0A58
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSaveDefaults 
         Height          =   375
         Left            =   1167
         TabIndex        =   26
         Top             =   60
         Width           =   1605
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
         Caption         =   "frmChartOrdBar.frx":0A74
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmChartOrdBar.frx":0AB4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmChartOrdBar.frx":0AD4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   100
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
         Caption         =   "frmChartOrdBar.frx":0AF0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmChartOrdBar.frx":0B14
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmChartOrdBar.frx":0B34
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   2865
         TabIndex        =   3
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
         Caption         =   "frmChartOrdBar.frx":0B50
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmChartOrdBar.frx":0B7C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmChartOrdBar.frx":0B9C
         RightToLeft     =   0   'False
      End
   End
   Begin gdOCX.gdSelectColor gdHighlightPos 
      Height          =   315
      Left            =   3016
      TabIndex        =   29
      Top             =   815
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      CustomColor     =   255
   End
   Begin gdOCX.gdSelectColor gdHighlightEquity 
      Height          =   315
      Left            =   3016
      TabIndex        =   9
      Top             =   1125
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      CustomColor     =   255
   End
   Begin HexUniControls.ctlUniLabelXP Label13 
      Height          =   255
      Left            =   1118
      Top             =   4080
      Width           =   1695
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
      Caption         =   "frmChartOrdBar.frx":0BB8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmChartOrdBar.frx":0C06
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOrdBar.frx":0C26
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmChartOrdBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmChartOrdBar.frm
'' Description: Form to allow the user to change settings for the chart order bar
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    frm As Form
    iTextAlign As Long
End Type

Private m As mPrivate

Public Sub ShowMe(frm As Form)
On Error GoTo ErrSection:

    Dim i&
    Dim lPreset1 As Long                ' First order quantity preset
    Dim lPreset2 As Long                ' Second order quantity preset
    Dim lPreset3 As Long                ' Third order quantity preset
    
    Set m.frm = frm
    
    If frm Is Nothing Then Exit Sub
    If frm.Chart Is Nothing Then Exit Sub

    chkShowAcctBar.Value = Abs(m.frm.Chart.ShowAccountBar)
    chkShowOrdBar.Value = Abs(m.frm.vseOrderBar.Visible)
    
    g.Broker.GetQuantityPresets m.frm.TradeAccountID, frm.SymbolOrSymbolID, lPreset1, lPreset2, lPreset3
    txtQty1.Text = Str(lPreset1)
    txtQty2.Text = Str(lPreset2)
    txtQty3.Text = Str(lPreset3)
    
    chkOrderLine.Value = frm.Chart.ShowOrderHorzLine

'average entry line
'1=color, 2=pen style, 3=textPosition, 4=font, 5=font size, 6=font style, 7=font underline
    chkAvgEntryLine.Value = frm.Chart.ShowAvgEntryLine
    gdColor.Color = ValOfText(frm.Chart.AvgEntryProp(1))
    LoadPenStyles ValOfText(frm.Chart.AvgEntryProp(2))
    
'highlight position checkbox & color
    i = m.frm.Chart.HighlightPos
    If i = -1 Then
        chkHighlightPos.Value = vbUnchecked
        gdHighlightPos.Color = vbCyan
    Else
        chkHighlightPos.Value = vbChecked
        gdHighlightPos.Color = i
    End If

'highlight equity checkbox & color
    i = m.frm.Chart.HighlightEquity
    If i = -1 Then
        chkHighlightEquity.Value = vbUnchecked
        gdHighlightEquity.Color = vbCyan
    Else
        chkHighlightEquity.Value = vbChecked
        gdHighlightEquity.Color = i
    End If
    
'e_topLeft, e_topRight, e_topCtr        0,1,2 (text below line)
'e_btmLeft, e_btmRight, e_btmCtr        3,4,5 (text above line)
'e_ctrLeft, e_ctrRight, e_ctrCtr        6,7,8 (text centered on line)
    i = ValOfText(frm.Chart.AvgEntryProp(3))
    Select Case i
        Case 3, 4, 5
            optTextPosition(0).Value = True     'above line
        Case 6, 7, 8, 9
            optTextPosition(1).Value = True     'centered on line
        Case 0, 1, 2
            optTextPosition(2).Value = True     'below line
        Case Else
            optTextPosition(1).Value = True     'centered on line
    End Select
    
    If i = 1 Or i = 4 Or i = 7 Then
        chkTextOnRight.Value = vbChecked
    Else
        chkTextOnRight.Value = vbUnchecked
    End If
    m.iTextAlign = i
    
    InitOrderButtonsGrid
    InitAccountGrid
    InitCboLimitPeriod
    
    CenterTheForm Me
    
    ShowForm Me, eForm_Modal, , , ALT_GRID_ROW_COLOR

    Exit Sub
    
ErrSection:
    RaiseError "frmChartOrdBar.ShowMe"
    
End Sub

Private Sub chkShowAcctBar_Click()
On Error Resume Next

    If Not m.frm Is Nothing Then
        If Not m.frm.Chart Is Nothing Then
            If m.frm.Chart.ShowTrades = 2 And chkShowOrdBar.Value <> 0 Then
                m.frm.ToggleAccountBar          'aardvark 4115
            End If
        End If
    End If

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:
    
    Unload Me
    
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOrdBar.cmdCancel_Click"
    
End Sub

Private Sub InitOrderButtonsGrid()

    Dim i&, strText$
    Dim aButtons As New cGdArray
    
    Dim nShow As Long
    Dim strCode As String
    Dim strDescription As String

    strText = m.frm.Chart.OrdBarCtrls
    
    If InStr(strText, ";") = 0 Then strText = kOrdBarDefaults
    aButtons.SplitFields strText, "|"
    
    With fgOrderButtons
        .Redraw = flexRDNone
        
        SetupGrid fgOrderButtons, eGridMode_Grid
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionByRow
        .FixedRows = 1
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical
        .Cols = 3
        .Rows = 1
        .ColWidth(0) = 600
        .ColHidden(eGDOrderBarCfg_ButtonCode) = True
        'headers
        .TextMatrix(0, eGDOrderBarCfg_ButtonShow) = "Show"
        .TextMatrix(0, eGDOrderBarCfg_ButtonCode) = "Code"
        .TextMatrix(0, eGDOrderBarCfg_ButtonDescript) = "Button"
        
        For i = 0 To aButtons.Size - 1
            .Rows = .Rows + 1
            ParseOrdButtonString aButtons(i), nShow, strCode, strDescription
            .Cell(flexcpChecked, .Rows - 1, eGDOrderBarCfg_ButtonShow) = nShow
            .TextMatrix(.Rows - 1, eGDOrderBarCfg_ButtonCode) = strCode
            .TextMatrix(.Rows - 1, eGDOrderBarCfg_ButtonDescript) = strDescription
        Next
        
        .Cell(flexcpPictureAlignment, .FixedRows, eGDOrderBarCfg_ButtonShow, .Rows - 1, eGDOrderBarCfg_ButtonShow) = flexAlignCenterCenter
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        .Redraw = flexRDBuffered
    End With

End Sub

Private Sub InitAccountGrid()
On Error GoTo ErrSection:

    Dim i&, j&, strText$
    Dim aAcctHeaders As cGdArray
    Dim aSorted As New cGdArray
    
    Set aAcctHeaders = m.frm.AcctBarHeader
    
    With fgAcctCols
        .Redraw = flexRDNone
        SetupGrid Me.fgAcctCols, eGridMode_Grid
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .Rows = 7
        .Cols = 2
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarNone
        .ColWidth(0) = 1000
        'headers
        .TextMatrix(0, 0) = "Show"
        .TextMatrix(0, 1) = "Data"
        
        .TextMatrix(1, 1) = "Position"
        .TextMatrix(2, 1) = "Avg Entry"
        .TextMatrix(3, 1) = "Open Equity"
        .TextMatrix(4, 1) = "Acct Balance"
        .TextMatrix(5, 1) = "Daily P/L"
        If SecurityType(m.frm.Chart.Symbol) = "S" Then
            .TextMatrix(6, 1) = "# Shares"
        Else
            .TextMatrix(6, 1) = "# Contracts"
        End If
        
        'set check boxes
        If Not aAcctHeaders Is Nothing Then
            For i = 0 To aAcctHeaders.Size - 1
                aSorted.Add aAcctHeaders(i)
            Next
            If aSorted.Size > 0 Then
                aSorted.Sort
                For i = .FixedRows To .Rows - 1
                    strText = .TextMatrix(i, 1)
                    If aSorted.BinarySearch(strText, j) Then
                        If aSorted(j) = strText Then
                            .Cell(flexcpChecked, i, 0) = 1
                        Else
                            .Cell(flexcpChecked, i, 0) = 2
                        End If
                    ElseIf InStr(strText, "#") And InStr(aSorted(0), "#") Then
                        'note: this check works for now because an ascending sort will put # as first array element
                        .Cell(flexcpChecked, i, 0) = 1
                    Else
                        .Cell(flexcpChecked, i, 0) = 2
                    End If
                Next
            End If
        End If
        
        .Cell(flexcpAlignment, 0, 0, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 1, .Rows - 1, 1) = flexAlignLeftCenter
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexAlignCenterCenter
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistributionCfg.InitAccountGrid", eGDRaiseError_Raise

End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:

    Dim nStyle&
    
    'set font currently in use
    '1=color, 2=pen style, 3=textPosition, 4=font, 5=font size, 6=font style, 7=font underline
    Me.Font.Name = m.frm.Chart.AvgEntryProp(4)
    Me.Font.Size = ValOfText(m.frm.Chart.AvgEntryProp(5))
    Me.Font.Underline = ValOfText(m.frm.Chart.AvgEntryProp(7))
    nStyle = ValOfText(m.frm.Chart.AvgEntryProp(6))
    Select Case nStyle
        Case 0:
            Me.Font.Italic = False
            Me.Font.Bold = False
        Case 1:
            Me.Font.Italic = False
            Me.Font.Bold = True
        Case 2:
            Me.Font.Italic = True
            Me.Font.Bold = False
        Case 3:
            Me.Font.Italic = True
            Me.Font.Bold = True
    End Select
    
    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        m.frm.Chart.AvgEntryProp(4) = Me.Font.Name
        m.frm.Chart.AvgEntryProp(5) = Me.Font.Size
        If Me.Font.Underline = True Then
            m.frm.Chart.AvgEntryProp(7) = 1
        Else
            m.frm.Chart.AvgEntryProp(7) = 0
        End If
        
        'style - 0=reg,1=bold,2=italic,3=bold italic
        nStyle = 0
        If Me.Font.Bold = True Then
            If Me.Font.Italic = True Then
                nStyle = 3
            Else
                nStyle = 1
            End If
        ElseIf Me.Font.Italic = True Then
            nStyle = 2
        End If
        
        m.frm.Chart.AvgEntryProp(6) = nStyle
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOrdBar.cmdFont_Click"

End Sub

Private Sub SaveAvgEntryInfo()
On Error GoTo ErrSection:
    
    Dim i&

    m.frm.Chart.ShowAvgEntryLine = chkAvgEntryLine.Value

    i = ValOfText(m.frm.Chart.AvgEntryProp(1))
    If i <> gdColor.Color Then
        m.frm.Chart.AvgEntryProp(1) = Str(gdColor.Color)
    End If
    i = ValOfText(m.frm.Chart.AvgEntryProp(2))
    If i <> cboLineStyle.ListIndex Then
        m.frm.Chart.AvgEntryProp(2) = cboLineStyle.ListIndex
    End If
    
    'text position
    'e_topLeft, e_topRight, e_topCtr        0,1,2 (text below line)
    'e_btmLeft, e_btmRight, e_btmCtr        3,4,5 (text above line)
    'e_ctrLeft, e_ctrRight, e_ctrCtr        6,7,8 (text centered on line)
    If chkTextOnRight.Value = vbChecked Then
        If optTextPosition(0).Value = True Then
            i = 4       'above to right
        ElseIf optTextPosition(2).Value = True Then
            i = 1       'below to right
        Else
            i = 7       'centered to right
        End If
    Else
        If optTextPosition(0).Value = True Then
            i = 3       'above to left
        ElseIf optTextPosition(2).Value = True Then
            i = 0       'below to left
        Else
            i = 6       'centered to left
        End If
    End If
    If i <> m.iTextAlign Then
        m.frm.Chart.AvgEntryProp(3) = i
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartOrdBar.SaveAvgEntryInfo"

End Sub

Private Sub SaveLimitTradesInfo()
On Error GoTo ErrSection:

    Dim strText$
    
    If chkLimitTrades.Value = 1 Then
        strText = Int(ValOfText(txtLimitTrades.Text)) & "|"
        If cboLimitPeriod.ListIndex = 0 Then
            strText = strText & "T"
        ElseIf cboLimitPeriod.ListIndex = 1 Then
            strText = strText & "B"
        ElseIf cboLimitPeriod.ListIndex = 2 Then
            strText = strText & "D"
        ElseIf cboLimitPeriod.ListIndex = 3 Then
            strText = strText & "M"
        Else
            strText = strText & "T"
        End If
    Else
        strText = ""
    End If
    m.frm.Chart.LimitTrades = strText

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartOrdBar.SaveLimitTradesInfo"
    
End Sub

Private Sub SavePresetQtyInfo()
On Error GoTo ErrSection:

    g.Broker.SetQuantityPresets m.frm.TradeAccountID, m.frm.SymbolOrSymbolID, Int(ValOfText(txtQty1.Text)), Int(ValOfText(txtQty2.Text)), Int(ValOfText(txtQty3.Text))

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartOrdBar.SavePresetQtyInfo"

End Sub

Private Sub SaveAcctBarInfo()
On Error GoTo ErrSection:
    
    Dim i&, strText$
    
    With fgAcctCols
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = 1 Then
                If Len(strText) > 0 Then
                    strText = strText & "|" & .TextMatrix(i, 1)
                Else
                    strText = .TextMatrix(i, 1)
                End If
            End If
        Next
    End With

    m.frm.SetAcctBarHeader strText
    m.frm.Chart.AcctBarCols = strText

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartOrdBar.SaveAcctBarInfo"

End Sub

Private Sub SaveOrderButtonsInfo()
On Error GoTo ErrSection:

    Dim strSave$, strCode$, i&
    
    With fgOrderButtons
        For i = .FixedRows To .Rows - 1
            strCode = .TextMatrix(i, eGDOrderBarCfg_ButtonCode)
            If strCode = "CLR" Then
                strSave = strSave & "CLR;0|"
            Else
                strSave = strSave & strCode & ";"
                strSave = strSave & .Cell(flexcpChecked, i, eGDOrderBarCfg_ButtonShow) & "|"
            End If
        Next
    End With
    
    m.frm.Chart.OrdBarCtrls = strSave
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartOrdBar.SaveOrderButtonsInfo"

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    If VerifyQuantityPresets = True Then
        SaveSettings
        Unload Me
    End If
        
    Exit Sub

ErrSection:
    RaiseError "frmChartOrdBar.cmdOK_Click"
End Sub

Private Sub InitCboLimitPeriod()
On Error GoTo ErrSection:

    Dim strText$

    cboLimitPeriod.Clear
    cboLimitPeriod.AddItem "Trades"
    cboLimitPeriod.AddItem "Bars"
    cboLimitPeriod.AddItem "Days"
    cboLimitPeriod.AddItem "Minutes"
    
    strText = m.frm.Chart.LimitTrades
    If Len(strText) > 0 Then
        txtLimitTrades.Text = Parse(strText, "|", 1)
        If InStr(strText, "T") <> 0 Then
            cboLimitPeriod.ListIndex = 0
        ElseIf InStr(strText, "B") <> 0 Then
            cboLimitPeriod.ListIndex = 1
        ElseIf InStr(strText, "D") <> 0 Then
            cboLimitPeriod.ListIndex = 2
        ElseIf InStr(strText, "M") <> 0 Then
            cboLimitPeriod.ListIndex = 3
        Else
            cboLimitPeriod.ListIndex = 0
        End If
        chkLimitTrades.Value = 1
    Else
        txtLimitTrades.Text = "10"
        cboLimitPeriod.ListIndex = 0
        chkLimitTrades.Value = 0
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOrdBar.InitCboLimitPeriod"
    
End Sub

Private Sub cmdSaveDefaults_Click()
On Error GoTo ErrSection:
    
    SaveSettings True
    Unload Me
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartOrdBar.cmdSaveDefaults_Click"

End Sub

Private Sub fgOrderButtons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    Dim r&
    
    With fgOrderButtons
        If .Row >= .FixedRows And .Row < .Rows Then
            r = .Row
            .Select r, 0, r, .Cols - 1
            .Refresh
            .DragRow r
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartOrdBar.fgOrderButtons_MouseDown"

End Sub

Private Sub Form_Load()
    Me.Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me
    
End Sub

Private Sub LoadPenStyles(ByVal nStyle&)
On Error GoTo ErrSection:
    
    With cboLineStyle
        .AddItem "Default"
        .AddItem "Thin"
        .AddItem "Medium"
        .AddItem "Thick"
        .AddItem "Dashed (Large)"
        .AddItem "Dashed (Small)"
        .AddItem "Dash Dot"
        If nStyle >= 0 And nStyle < cboLineStyle.ListCount Then
            .ListIndex = nStyle
        End If
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOrdBar.LoadPenStyles", eGDRaiseError_Raise
    
End Sub

Private Sub SaveSettings(Optional ByVal bSaveAsDefaults As Boolean = False)
On Error GoTo ErrSection:

    If m.frm Is Nothing Then Exit Sub
    If m.frm.Chart Is Nothing Then Exit Sub

    m.frm.Chart.ShowOrderHorzLine = chkOrderLine.Value
    m.frm.Chart.ShowAccountBar = -1 * chkShowAcctBar.Value
    
    If chkShowOrdBar.Value = 0 Then
        m.frm.Chart.ShowTrades = 0
    Else
        m.frm.Chart.ShowTrades = 2
    End If
    
    If chkHighlightPos.Value = vbChecked Then
        m.frm.Chart.HighlightPos = gdHighlightPos.Color
    Else
        m.frm.Chart.HighlightPos = -1
    End If
    
    If chkHighlightEquity.Value = vbChecked Then
        m.frm.Chart.HighlightEquity = gdHighlightEquity.Color
    Else
        m.frm.Chart.HighlightEquity = -1
    End If
    
    SavePresetQtyInfo
    SaveAvgEntryInfo
    SaveLimitTradesInfo
    SaveAcctBarInfo
    SaveOrderButtonsInfo
    
    If bSaveAsDefaults Then m.frm.Chart.OrdBarSaveAsDefaults

    m.frm.Chart.GenerateChart eRedo1_Scrolled       '4243
    FormResize m.frm

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOrdBar.SaveSettings", eGDRaiseError_Raise
    
End Sub

Private Function VerifyQuantityPresets() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lPreset As Long                 ' Preset value
    
    bReturn = True
    
    lPreset = Int(Val(txtQty1.Text))
    If g.Broker.ValidQuantity(m.frm.TradeAccountID, m.frm.SymbolOrSymbolID, lPreset) = False Then
        MoveFocus txtQty1
        InfBox "The first quantity preset is an invalid quantity", "!", , "Quantity Preset Error"
        bReturn = False
    End If
    
    If bReturn = True Then
        lPreset = Int(Val(txtQty2.Text))
        If g.Broker.ValidQuantity(m.frm.TradeAccountID, m.frm.SymbolOrSymbolID, lPreset) = False Then
            MoveFocus txtQty2
            InfBox "The second quantity preset is an invalid quantity", "!", , "Quantity Preset Error"
            bReturn = False
        End If
    End If
    
    If bReturn = True Then
        lPreset = Int(Val(txtQty3.Text))
        If g.Broker.ValidQuantity(m.frm.TradeAccountID, m.frm.SymbolOrSymbolID, lPreset) = False Then
            MoveFocus txtQty3
            InfBox "The third quantity preset is an invalid quantity", "!", , "Quantity Preset Error"
            bReturn = False
        End If
    End If
    
    VerifyQuantityPresets = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmChartOrdBar.VerifyQuantityPresets"
    
End Function

Private Sub txtQty1_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOrdBar.txtQty1_GotFocus"
    
End Sub

Private Sub txtQty2_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty2

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOrdBar.txtQty2_GotFocus"
    
End Sub

Private Sub txtQty3_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty3

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOrdBar.txtQty3_GotFocus"
    
End Sub

