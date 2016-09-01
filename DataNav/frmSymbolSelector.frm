VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSymbolSelector 
   Caption         =   "Symbol Selector"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   555
      Left            =   120
      TabIndex        =   14
      Top             =   5700
      Width           =   5595
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
      Caption         =   "frmSymbolSelector.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolSelector.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolSelector.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkAllCharts 
         Height          =   225
         Left            =   105
         TabIndex        =   4
         Top             =   225
         Visible         =   0   'False
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
         Caption         =   "frmSymbolSelector.frx":005C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolSelector.frx":009C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":00BC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdHardDrive 
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   120
         Width           =   1155
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
         Caption         =   "frmSymbolSelector.frx":00D8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolSelector.frx":010E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":01A2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   120
         Width           =   1035
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
         Caption         =   "frmSymbolSelector.frx":01BE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolSelector.frx":01EA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":020A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   120
         Width           =   1035
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
         Caption         =   "frmSymbolSelector.frx":0226
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolSelector.frx":024C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":026C
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraOption 
      Height          =   3135
      Left            =   180
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   5535
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
      Caption         =   "frmSymbolSelector.frx":0288
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolSelector.frx":02E2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolSelector.frx":0302
      RightToLeft     =   0   'False
      Begin gdOCX.gdSelectDate dtMonth 
         Height          =   315
         Left            =   2820
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
         ShowCalendar    =   0   'False
         ShowDate        =   2
         Value           =   40193
      End
      Begin HexUniControls.ctlUniRadioXP optPut 
         Height          =   220
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   675
         _ExtentX        =   1191
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
         Caption         =   "frmSymbolSelector.frx":031E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolSelector.frx":0346
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":0366
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtStrike 
         Height          =   315
         Left            =   1200
         TabIndex        =   26
         Top             =   720
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmSymbolSelector.frx":0382
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
         Tip             =   "frmSymbolSelector.frx":03A2
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":03C2
      End
      Begin HexUniControls.ctlUniRadioXP optCall 
         Height          =   220
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   675
         _ExtentX        =   1191
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
         Caption         =   "frmSymbolSelector.frx":03DE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolSelector.frx":0408
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":0428
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin MSComCtl2.MonthView mvExp 
         Height          =   2370
         Left            =   2700
         TabIndex        =   34
         Top             =   600
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   122028033
         TitleBackColor  =   16035718
         CurrentDate     =   40193
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   255
         Left            =   60
         Top             =   60
         Visible         =   0   'False
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
         Caption         =   "frmSymbolSelector.frx":0444
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolSelector.frx":049E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":04BE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label9 
         Height          =   255
         Left            =   180
         Top             =   360
         Width           =   975
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
         Caption         =   "frmSymbolSelector.frx":04DA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolSelector.frx":0512
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":0532
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblExp 
         Height          =   255
         Left            =   2820
         Top             =   360
         Width           =   2475
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
         Caption         =   "frmSymbolSelector.frx":054E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolSelector.frx":05B2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":05D2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label11 
         Height          =   255
         Left            =   1320
         Top             =   360
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
         Caption         =   "frmSymbolSelector.frx":05EE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolSelector.frx":0628
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":0648
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   255
         Left            =   900
         Top             =   780
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
         Caption         =   "frmSymbolSelector.frx":0664
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolSelector.frx":0688
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":06A8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   2340
         Top             =   780
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
         Caption         =   "frmSymbolSelector.frx":06C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolSelector.frx":06EA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":070A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblNote 
         Height          =   915
         Left            =   360
         Top             =   1260
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
         Caption         =   "frmSymbolSelector.frx":0726
         BackColor       =   -2147483633
         ForeColor       =   12582912
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolSelector.frx":0818
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":0838
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFutures 
      Height          =   1380
      Left            =   180
      TabIndex        =   18
      Top             =   4200
      Width           =   5535
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
      Caption         =   "frmSymbolSelector.frx":0854
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolSelector.frx":0880
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolSelector.frx":08A0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraContinuous 
         Height          =   1305
         Left            =   2820
         TabIndex        =   28
         Top             =   60
         Width           =   2655
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
         Caption         =   "frmSymbolSelector.frx":08BC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSymbolSelector.frx":08DE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":08FE
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkRoll 
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   15
            Width           =   1815
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
            Caption         =   "frmSymbolSelector.frx":091A
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":0964
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":0A2E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP opt57 
            Height          =   220
            Left            =   420
            TabIndex        =   32
            Top             =   270
            Width           =   600
            _ExtentX        =   1058
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
            Caption         =   "frmSymbolSelector.frx":0A4A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   12582912
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":0A70
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":0AF4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP opt56 
            Height          =   220
            Left            =   420
            TabIndex        =   31
            Top             =   510
            Width           =   600
            _ExtentX        =   1058
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
            Caption         =   "frmSymbolSelector.frx":0B10
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   12582912
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":0B36
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":0B9E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP opt55 
            Height          =   220
            Left            =   420
            TabIndex        =   30
            Top             =   750
            Width           =   600
            _ExtentX        =   1058
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
            Caption         =   "frmSymbolSelector.frx":0BBA
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   12582912
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":0BE0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":0C74
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkBackAdjust 
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   1020
            Width           =   2055
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
            Caption         =   "frmSymbolSelector.frx":0C90
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":0CDE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":0D80
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lbl57 
            Height          =   195
            Left            =   1080
            Top             =   270
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
            Caption         =   "frmSymbolSelector.frx":0D9C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSymbolSelector.frx":0DD8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":0DF8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lbl55 
            Height          =   195
            Left            =   1080
            Top             =   750
            Width           =   1455
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
            Caption         =   "frmSymbolSelector.frx":0E14
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSymbolSelector.frx":0E5E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":0E7E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lbl56 
            Height          =   195
            Left            =   1080
            Top             =   510
            Width           =   1335
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
            Caption         =   "frmSymbolSelector.frx":0E9A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSymbolSelector.frx":0ED4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":0EF4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSession 
         Height          =   1305
         Left            =   0
         TabIndex        =   19
         Top             =   60
         Width           =   2655
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
         Caption         =   "frmSymbolSelector.frx":0F10
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSymbolSelector.frx":0F50
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":0F70
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optPit 
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   270
            Width           =   720
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
            Caption         =   "frmSymbolSelector.frx":0F8C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   12582912
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":0FB2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":101A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optElectronic 
            Height          =   220
            Left            =   240
            TabIndex        =   22
            Top             =   510
            Width           =   720
            _ExtentX        =   1270
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
            Caption         =   "frmSymbolSelector.frx":1036
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   12582912
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":105C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":10D2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optElectronicDay 
            Height          =   220
            Left            =   240
            TabIndex        =   21
            Top             =   750
            Width           =   720
            _ExtentX        =   1270
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
            Caption         =   "frmSymbolSelector.frx":10EE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   12582912
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":1114
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":11F0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optCombined 
            Height          =   220
            Left            =   240
            TabIndex        =   20
            Top             =   990
            Width           =   720
            _ExtentX        =   1270
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
            Caption         =   "frmSymbolSelector.frx":120C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   12582912
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":1232
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":12CC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPit 
            Height          =   195
            Left            =   1020
            Top             =   270
            Width           =   1095
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
            Caption         =   "frmSymbolSelector.frx":12E8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSymbolSelector.frx":1324
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":1344
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblElectronic 
            Height          =   195
            Left            =   1020
            Top             =   510
            Width           =   1515
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
            Caption         =   "frmSymbolSelector.frx":1360
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSymbolSelector.frx":13AA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":13CA
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblElectronicDay 
            Height          =   195
            Left            =   1020
            Top             =   750
            Width           =   1515
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
            Caption         =   "frmSymbolSelector.frx":13E6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSymbolSelector.frx":1430
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":1450
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblCombined 
            Height          =   195
            Left            =   1020
            Top             =   990
            Width           =   1395
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
            Caption         =   "frmSymbolSelector.frx":146C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSymbolSelector.frx":14B2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":14D2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgStrikes 
      Height          =   1215
      Left            =   2760
      TabIndex        =   17
      Top             =   2340
      Width           =   2895
      _cx             =   5106
      _cy             =   2143
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
   Begin VSFlex7LCtl.VSFlexGrid fgMonths 
      Height          =   1215
      Left            =   300
      TabIndex        =   16
      Top             =   2340
      Width           =   2895
      _cx             =   5106
      _cy             =   2143
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
   Begin VB.Timer tmrSortCol 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   180
      Top             =   3600
   End
   Begin VSFlex7LCtl.VSFlexGrid fgSymbols 
      Height          =   1575
      Left            =   180
      TabIndex        =   10
      Top             =   1620
      Width           =   5535
      _cx             =   9763
      _cy             =   2778
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
   Begin HexUniControls.ctlUniFrameWL fraFilter 
      Height          =   1455
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   5535
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
      Caption         =   "frmSymbolSelector.frx":14EE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolSelector.frx":1520
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolSelector.frx":1540
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkDefault 
         Height          =   315
         Left            =   3600
         TabIndex        =   9
         Top             =   1020
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
         Caption         =   "frmSymbolSelector.frx":155C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolSelector.frx":159C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":15BC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraFind 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   5115
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
         Caption         =   "frmSymbolSelector.frx":15D8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSymbolSelector.frx":1604
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":1624
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optDescBeginsWith 
            Height          =   255
            Left            =   1980
            TabIndex        =   2
            Top             =   0
            Width           =   2055
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
            Caption         =   "frmSymbolSelector.frx":1640
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":1690
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":16B0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optSymbolBeginsWith 
            Height          =   255
            Left            =   60
            TabIndex        =   1
            Top             =   0
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
            Caption         =   "frmSymbolSelector.frx":16CC
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmSymbolSelector.frx":1712
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":1732
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optContains 
            Height          =   255
            Left            =   4080
            TabIndex        =   3
            Top             =   0
            Width           =   975
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
            Caption         =   "frmSymbolSelector.frx":174E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmSymbolSelector.frx":1780
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSymbolSelector.frx":17A0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin MSComctlLib.ImageCombo cboFilters 
         Height          =   330
         Left            =   840
         TabIndex        =   8
         Top             =   1020
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "ImageCombo1"
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdFind 
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   570
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
         Caption         =   "frmSymbolSelector.frx":17BC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolSelector.frx":17E4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":1804
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtFilterText 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   600
         Width           =   3255
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmSymbolSelector.frx":1820
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
         Tip             =   "frmSymbolSelector.frx":1840
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":1860
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   240
         Top             =   630
         Width           =   495
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
         Caption         =   "frmSymbolSelector.frx":187C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolSelector.frx":18A8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":18C8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   240
         Top             =   1080
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
         Caption         =   "frmSymbolSelector.frx":18E4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolSelector.frx":1914
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolSelector.frx":1934
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmSymbolSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSymbolSelector.frm
'' Description: Allows the user to easily select one or more symbols given
''              certain filters and sorting options
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 08/21/01  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Enum eDirection
    edirDescending = -1
    edirToggle = 0
    edirAscending = 1
End Enum

Private Enum eGDMonthCols
    eGDMonthCol_Letter = 0
    eGDMonthCol_Desc
    eGDMonthCol_NumCols
End Enum

Private Enum eGDStrikeCols
    eGDStrikeCol_Letter = 0
    eGDStrikeCol_Desc
    eGDStrikeCol_NumCols
End Enum

Private Type mPrivate
    bSaved As Boolean
    strHardDrive As String
    SymbolGrid As cSymbolGrid
    lFieldNum As Long
    lSortedCol As Long
    aFilter As cGdArray
    bDescending As Boolean
    bAllowOptions As Boolean
    bSkipRowChangeEvent As Boolean
    astrComboIDs As cGdArray
    strSymbolMap As String
    strCaption As String
    
    bFromUser As Boolean
    bSkipSaveSettings As Boolean
    bUseOptionWizard As Boolean
End Type
Private m As mPrivate

Private Function MonthCol(ByVal Col As eGDMonthCols) As Long
    MonthCol = Col
End Function

Private Function StrikeCol(ByVal Col As eGDStrikeCols) As Long
    StrikeCol = Col
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboFilters_Click
'' Description: When the Filters combo box changes, refilter the symbol grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboFilters_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If Screen.MousePointer = vbHourglass Then Exit Sub
    Filter
    MoveFocus txtFilterText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.cboFilters.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkBackAdjust_Click
'' Description: Change the symbol labels as the back adjust box changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkBackAdjust_Click()
On Error GoTo ErrSection:

    If chkBackAdjust.Value = vbChecked Then
        opt55.Caption = "065"
        opt56.Caption = "066"
        opt57.Caption = "067"
    Else
        opt55.Caption = "055"
        opt56.Caption = "056"
        opt57.Caption = "057"
    End If
        
    If Me.Visible And m.bFromUser Then
        ChangeSymbol
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.chkBackAdjust.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkRoll_Click
'' Description: Change the symbol based on the roll check box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkRoll_Click()
On Error GoTo ErrSection:

    If Me.Visible And m.bFromUser Then
        If (Not opt55) And (Not opt56) And (Not opt57) And (chkRoll = vbChecked) Then
            opt57.Value = True
            chkBackAdjust.Value = vbChecked
        End If
        
        ChangeSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.chkRoll.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on the cancel button, set the saved to false
''              and hide the form to give control back to ShowMe
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    If Screen.MousePointer = vbHourglass Then Exit Sub
    m.bSaved = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdFind_Click
'' Description: If the user clicks on find, refilter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdFind_Click()
On Error GoTo ErrSection:

    If Screen.MousePointer = vbHourglass Then Exit Sub
    Filter
    MoveFocus txtFilterText
    cmdFind.Default = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.cmdFind.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdHardDrive_Click
'' Description: Allow the user to select a CSI, MetaStock, or GenTick file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdHardDrive_Click()
On Error GoTo ErrSection:

    Dim strInfo As String               ' Information back from DataMan form
    
    strInfo = frmDataMan.ShowMe
    If Len(strInfo) > 0 Then
        m.strHardDrive = strInfo
        m.bSaved = True
        Me.Hide
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSymbolSelector.cmdHardDrive.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOk_Click
'' Description: If the user clicks on the OK button, set the saved to true and
''              hide the form to give control back to ShowMe
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If Screen.MousePointer = vbHourglass Then Exit Sub
    m.bSaved = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub dtMonth_Changed()
On Error GoTo ErrSection:

    FixExpDate

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.dtMonth_Changed"
    Resume ErrExit
End Sub

Private Sub dtMonth_GotFocus()

    DoEvents
    MoveFocus txtStrike

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgMonths_AfterRowColChange
'' Description: Display the new month after the user changes rows
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgMonths_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim strText As String               ' String from the text box
    Dim lSpace As Long                  ' Position of the space in the string
    Dim strBase As String               ' Base symbol
    Dim strOption As String             ' Option part of the symbol
    Dim strMonth As String              ' Month that the user selected

    If m.bSkipRowChangeEvent Then Exit Sub

    If (NewRow <> OldRow) And (Me.Visible = True) Then
        strText = UCase(txtFilterText.Text)
        strBase = Parse(strText, " ", 1)
        strOption = Parse(strText, " ", 2)
        
        If NewRow >= fgMonths.FixedRows Then
            strMonth = fgMonths.TextMatrix(NewRow, MonthCol(eGDMonthCol_Letter))
            
            Select Case Len(strOption)
                Case 0, 1
                    txtFilterText.Text = strBase & " " & strMonth
                
                Case 2
                    txtFilterText.Text = strBase & " " & strMonth & Right(strOption, 1)
            
            End Select
            
            EnableControls
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.fgMonths.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrikes_AfterRowColChange
'' Description: Display the new strike code after the user changes rows
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrikes_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim strText As String               ' String from the text box
    Dim lSpace As Long                  ' Position of the space in the string
    Dim strBase As String               ' Base symbol
    Dim strOption As String             ' Option part of the symbol
    Dim strStrike As String             ' Strike that the user selected

    If m.bSkipRowChangeEvent Then Exit Sub

    If (NewRow <> OldRow) And (Me.Visible = True) Then
        strText = UCase(txtFilterText.Text)
        strBase = Parse(strText, " ", 1)
        strOption = Parse(strText, " ", 2)
        
        If NewRow >= fgStrikes.FixedRows Then
            strStrike = fgStrikes.TextMatrix(NewRow, StrikeCol(eGDStrikeCol_Letter))
            
            Select Case Len(strOption)
                Case 1, 2
                    txtFilterText.Text = strBase & " " & Left(strOption, 1) & strStrike
            
            End Select
            
            EnableControls
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.fgStrikes.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_AfterRowColChange
'' Description: Enable/Disable/Set controls based on the new symbol
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If OldRow <> NewRow Then
        EnableControls
    End If
    
    If ActiveControl Is fgSymbols And optSymbolBeginsWith And NewRow >= fgSymbols.FixedRows Then
        txtFilterText.Text = fgSymbols.TextMatrix(NewRow, kSymbolCol)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.fgSymbols.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_BeforeMouseDown
'' Description: If the user clicks the left mouse button in the fixed row
''              of the grid, start the timer to allow sorting
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Mouse location, Whether
''              or not to cancel the mouse press
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    If Screen.MousePointer = vbHourglass Then
        Cancel = True
        Exit Sub
    End If
    
    If Button = vbLeftButton And fgSymbols.MouseRow = 0 Then
        tmrSortCol.Enabled = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.fgSymbols.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgSymbols_DblClick()
On Error GoTo ErrSection:
    
    If Screen.MousePointer = vbHourglass Then Exit Sub
    m.bSaved = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.fgSymbols.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_KeyPress
'' Description: If the user presses an ASCII key in the grid, send it to the
''              Filter text box instead
'' Inputs:      ASCII number of the key that was pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii > 32 And KeyAscii < 128 Then
        MoveFocus txtFilterText
        txtFilterText.SelStart = 9999
        SendKeys Chr(KeyAscii)
        KeyAscii = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.fgSymbols.KeyPress", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgSymbols_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgSymbols
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form gets activated, make sure that the FilterText
''              text box gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    If optPut Or optCall Then
        MoveFocus txtStrike
    Else
        MoveFocus txtFilterText
    End If
    If fgSymbols.Row >= fgSymbols.FixedRows Then
        fgSymbols.ShowCell fgSymbols.Row, kSymbolCol
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim i&

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    ElseIf Shift = 2 And UCase(Chr(KeyCode)) = "A" Then
        ' Ctrl-A: hot-key to select all
        With fgSymbols
            If .Rows < 5000 And .AllowSelection Then
                For i = .FixedRows To .Rows - 1
                    .IsSelected(i) = True
                Next
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center it and set the icon
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement of the form from the ini file

    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("frmSymbolSelector", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    Me.Icon = Picture16(ToolbarIcon("kSelect"))
    cmdOK.Default = True
    Set m.astrComboIDs = New cGdArray
    
    optPit.Caption = ""
    optElectronic.Caption = ""
    optElectronicDay.Caption = ""
    optCombined.Caption = ""
    
    chkDefault.Value = vbUnchecked
    
    'opt55.Caption = ""
    'opt56.Caption = ""
    'opt57.Caption = ""
    
    m.strSymbolMap = FileToString(AddSlash(App.Path) & "Info\SymbolMap.CSV")
    m.strSymbolMap = "|," & Replace(m.strSymbolMap, vbCrLf, ",|,") & ",|"
    
    'chkAllCharts.Value = GetIniFileProperty("AllCharts", vbUnchecked, "SymbolSelector", g.strIniFile)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Shows the form and if the user has clicked on OK, returns the
''              array of symbols that were chosen
'' Inputs:      Text to start with
'' Returns:     Array of symbols on OK, empty array otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal strSymbol As String = "", _
                        Optional ByVal bAllowMultiSelect As Boolean = True, _
                        Optional ByVal bShowFilters As Boolean = True, _
                        Optional ByVal strCaption As String = "Symbol Selector", _
                        Optional ByVal bAllowHardDrive As Boolean = False, _
                        Optional ByVal bSelectFilterText As Boolean = True, _
                        Optional ByVal bAllowOptions As Boolean = False, _
                        Optional ByVal strOnlyThisGroup As String = "", _
                        Optional ByVal bChangeAllCharts As Boolean = False, _
                        Optional ByVal strCenterPoint As String = "", _
                        Optional ByVal bFromOptNav As Boolean = False) As cGdArray
On Error GoTo ErrSection:

    Dim aSymbols As New cGdArray        ' Array of symbols to pass back
    Dim lIndex As Long                  ' Index for a for loop
    Dim frm As Form                     ' Index for a for loop
    Dim strFilter As String             ' Default filter to start with
    Dim strText As String
    Dim OptionSymbol As New cOptionSymbol

    ' don't do this if "busy" (e.g. causes a lock-up if try to do
    ' this while in the middle of running a strategy on a chart)
    If Screen.MousePointer = vbHourglass Then
        ' wait a second and check again
        Sleep 1
        If Screen.MousePointer = vbHourglass Then
            Beep
            Set ShowMe = aSymbols
            Unload Me
            Exit Function
        End If
    End If

    strSymbol = Trim(UCase(strSymbol))
    m.bUseOptionWizard = True ' FileExist(App.Path & "\NewOptWizard.flg")

    If Not HasGold(False, , False) Then
        bAllowHardDrive = False
    End If
    
    m.bAllowOptions = bAllowOptions
    tmrSortCol.Enabled = False
    
    aSymbols.Create eGDARRAY_Strings
    
    m.lFieldNum = g.SymbolPool.ArrayTable.NumFields
    
    ' Initialize the controls on the form
    Set m.SymbolGrid = Nothing
    m.lSortedCol = kSymbolCol
    optSymbolBeginsWith.Value = True
    cmdFind.Enabled = False
    txtFilterText.Text = strSymbol
    If bSelectFilterText = True Then SelectAll txtFilterText
    cboFilters.ImageList = frmMain.img16
    cboFilters.Locked = True
    LoadCombo bShowFilters
    strOnlyThisGroup = Trim(UCase(strOnlyThisGroup))
    If Len(strOnlyThisGroup) > 0 Then
        If InStr(strOnlyThisGroup, ":") = 0 Then
            Select Case Parse(strOnlyThisGroup, ".", 2)
            Case "GRP"
                strOnlyThisGroup = "GRP:" & strOnlyThisGroup
            Case "FIL"
                strOnlyThisGroup = "FIL:" & strOnlyThisGroup
            Case "SCN"
                strOnlyThisGroup = "DSV:" & strOnlyThisGroup
            End Select
        End If
        If m.astrComboIDs.BinarySearch(UCase(strOnlyThisGroup)) = True Then
            cboFilters.ComboItems(UCase(strOnlyThisGroup)).Selected = True
            cboFilters.Enabled = False
        End If
    Else
        strFilter = GetIniFileProperty("SymbolGroup", "", "SymbolSelector", g.strIniFile)
        If Len(strFilter) > 0 Then
            If m.astrComboIDs.BinarySearch(UCase(strFilter)) = True Then
                cboFilters.ComboItems(UCase(strFilter)).Selected = True
                chkDefault.Value = vbChecked
            End If
        End If
    End If
    
    If Not bAllowHardDrive Then
        cmdHardDrive.Visible = False
        'fraButtons.Width = cmdCancel.Left + cmdCancel.Width
    End If
    
    m.strCaption = strCaption
    Me.Caption = strCaption
    
    Set m.SymbolGrid = New cSymbolGrid
    m.SymbolGrid.InitGrid fgSymbols, "GRP:_FLAGS_.GRP|INF:Symbol|INF:Description", bAllowMultiSelect, False
    fgSymbols.ColWidth(kSymbolCol) = 1440 ' so things like $EUR-USD@FXCM can be seen
    Filter 'True
        
    fgSymbols.Editable = flexEDNone
    'Sort edirAscending '(this messes up the selection of the default symbol)
   
    m.bSaved = False
    m.strHardDrive = ""
    
    If bSelectFilterText = False Then
        MoveFocus txtFilterText
        SendKeys strSymbol
    End If
    
    InitMonthsGrid
    InitStrikesGrid
    
    If m.bUseOptionWizard And m.bAllowOptions And InStr(strSymbol, " ") > 0 Then
        With OptionSymbol
            .FromGenesis strSymbol
            If .Strike > 0 And .Month > 0 Then
                If .IsPut Then
                    optPut = True
                Else
                    optCall = True
                End If
                txtStrike.Text = Str(.Strike)
                If .IsFutureOption Then
                    dtMonth.Value = DateSerial(.Year, .Month, 1)
                Else
                    mvExp.Value = DateSerial(.Year, .Month, .Day)
                    FixExpDate ' run this now to set default day of month
                    ' now set date again since default day can be overridden
                    mvExp.Value = DateSerial(.Year, .Month, .Day)
                End If
                txtFilterText.Text = .ToGenesis
            End If
        End With
    End If
    
    EnableControls
    If Len(strSymbol) > 0 And m.bAllowOptions = True Then
        txtFilterText_Change
    End If
    
    If bChangeAllCharts And NumberOfCharts > 1 Then
        'chkAllCharts.Visible = True
    ElseIf bAllowHardDrive Then
        'chkAllCharts.Visible = False
    Else
        'chkAllCharts.Visible = False
    End If
    
    If Len(strCenterPoint) > 0 Then
        If InStr(strCenterPoint, ",") > 0 Then
            Me.Move Parse(strCenterPoint, ",", 1) - Me.Width / 2, _
                    Parse(strCenterPoint, ",", 2) - Me.Height / 2
        End If
        If bFromOptNav Then
            chkDefault.Visible = True
        Else
            chkDefault.Visible = False
        End If
        m.bSkipSaveSettings = True
    Else
        m.bSkipSaveSettings = False
        chkDefault.Visible = True
    End If
    
    fgSymbols.ShowCell fgSymbols.Row, kSymbolCol
    If bFromOptNav Then
        ' if from OptNav, need to show this on top as an "act modal" without an owner
        SetFormTopmost Me, True
        ShowForm Me, eForm_ActModal, , , ALT_GRID_ROW_COLOR
        SetFormTopmost Me, False
    Else
        ShowForm Me, eForm_Modal, , , ALT_GRID_ROW_COLOR
    End If
    
    If m.bSaved Then
        If Len(m.strHardDrive) > 0 Then
            aSymbols.Add m.strHardDrive
        ' if choosing an option symbol (with a space, but must be using "Symbol Begins With")
        ElseIf m.bAllowOptions And optSymbolBeginsWith And InStr(Trim(txtFilterText.Text), " ") > 0 Then
            aSymbols.Add UCase(Trim(txtFilterText.Text))
        Else
            With fgSymbols
                For lIndex = 0 To .SelectedRows - 1
                    ' Make sure not to include the fixed row
                    If .SelectedRow(lIndex) >= .FixedRows Then
                        aSymbols.Add .Cell(flexcpText, .SelectedRow(lIndex), kSymbolCol)
                    End If
                Next lIndex
            End With
            ' check for special cases
            If optSymbolBeginsWith And aSymbols.Size < 2 Then
                strText = UCase(Trim(txtFilterText))
                ' check if typed in a SymbolID
                If IsDigit(strText, 1) Then
                    aSymbols(0) = GetSymbol(Val(strText))
                ' check if typing in a symbol which isn't in the selected filter
                ' (if the first part of the grid row doesn't match what was typed, then use what was typed)
                ElseIf UCase(cboFilters.Text) <> "ALL SYMBOLS" And Left(UCase(aSymbols(0)), Len(strText)) <> strText Then
                    If Right(strText, 1) = "-" Then strText = strText & "067"
                    aSymbols(0) = strText
                End If
            End If
        End If
        
        If bChangeAllCharts And NumberOfCharts > 1 And chkAllCharts.Value = vbChecked _
                    And InStr(aSymbols(0), "|") = 0 And aSymbols.Size = 1 Then
            ' change active form first
            Set frm = ActiveChart
            If Not frm Is Nothing Then
                frm.Chart.SetSymbol aSymbols(0), True
                DoEvents
            End If
            ' then change others
            For Each frm In Forms
                'If frm.Name = "frmChart" Then      'JM 06-04-2009: original code; leave awhile then remove if all okay
                If IsFrmChart(frm) Then
                    frm.Chart.SetSymbol aSymbols(0), True
                    DoEvents
                End If
            Next frm
            Set frm = ActiveChart
            If Not frm Is Nothing Then
                frm.SetChartTabs
            End If
        End If
    End If
    
    Set ShowMe = aSymbols

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmSymbolSelector.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user exits the form via the control menu, set the saved
''              to false, cancel the unload, and hide the form to give control
''              back to ShowMe
'' Inputs:      Whether or not to cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = vbFormControlMenu Then
        m.bSaved = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: As the user resizes the form, resize controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Static lMinWidth&
    Dim lGridWidth As Long
    
    If lMinWidth = 0 Then lMinWidth = fraFilter.Width + fraFilter.Left * 2
    If LimitFormSize(Me, lMinWidth, fraFilter.Height + fraOption.Height + fraButtons.Height * 2) Then Exit Sub
    
    With fraButtons
        .Move (Me.ScaleWidth - .Width) / 2, Me.ScaleHeight - .Height
    End With
    
    With fraFutures
        .Move fraFilter.Left, fraButtons.Top - .Height
    End With
    
    With fgSymbols
        If fraFutures.Visible Then
            .Move .Left, .Top, Me.ScaleWidth - .Left * 2, fraFutures.Top - .Top
        Else
            .Move .Left, .Top, Me.ScaleWidth - .Left * 2, fraButtons.Top - .Top
        End If
    End With
    
    lGridWidth = (fgSymbols.Width - fgSymbols.Left) / 2
    With fgMonths
        .Move fgSymbols.Left, fgSymbols.Top, lGridWidth, fgSymbols.Height
    End With
    
    With fgStrikes
        .Move fgMonths.Width + (fgMonths.Left * 2), fgMonths.Top, lGridWidth, fgMonths.Height
    End With
    
    With fraFilter
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, .Height
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, clear the field that we created out
''              of the pool
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    If Not m.bSkipSaveSettings Then
        'SetIniFileProperty "AllCharts", chkAllCharts.Value, "SymbolSelector", g.strIniFile
        SetIniFileProperty "frmSymbolSelector", GetFormPlacement(Me), "Placement", g.strIniFile
        
        If chkDefault.Value = vbChecked Then
            SetIniFileProperty "SymbolGroup", cboFilters.SelectedItem.Key, "SymbolSelector", g.strIniFile
        Else
            SetIniFileProperty "SymbolGroup", "", "SymbolSelector", g.strIniFile
        End If
    End If
    
    tmrSortCol.Enabled = False
    g.SymbolPool.ArrayTable.ClearField m.lFieldNum
    Set m.aFilter = Nothing
    Set m.astrComboIDs = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lbl55_Click
'' Description: Make this act as if the user clicked on the 55 option button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lbl55_Click()
On Error GoTo ErrSection:

    opt55.Value = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.lbl55.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lbl56_Click
'' Description: Make this act as if the user clicked on the 56 option button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lbl56_Click()
On Error GoTo ErrSection:

    opt56.Value = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.lbl56.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lbl57_Click
'' Description: Make this act as if the user clicked on the 57 option button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lbl57_Click()
On Error GoTo ErrSection:

    opt57.Value = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.lbl57.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lblCombined_Click
'' Description: Make this act as if the user clicked on the pit+electronic option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblCombined_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        optCombined.Value = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSymbolSelector.lblCombined.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lblElectronic_Click
'' Description: Make this act as if the user clicked on the electronic option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblElectronic_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        optElectronic.Value = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSymbolSelector.lblElectronic.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lblElectronicDay_Click
'' Description: Make this act as if the user clicked on the electronic RTH option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblElectronicDay_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        optElectronicDay.Value = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSymbolSelector.lblElectronicDay.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lblPit_Click
'' Description: Make this act as if the user clicked on the pit option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblPit_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        optPit.Value = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSymbolSelector.lblPit.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mvExp_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
On Error GoTo ErrSection:

    FixExpDate

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.mvExp_SelChange"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    opt55_Click
'' Description: Change the symbol if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub opt55_Click()
On Error GoTo ErrSection:

    If Me.Visible And m.bFromUser Then
        If chkRoll.Value <> vbChecked Then chkRoll.Value = vbChecked
        ChangeSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.opt55.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    opt56_Click
'' Description: Change the symbol if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub opt56_Click()
On Error GoTo ErrSection:

    If Me.Visible And m.bFromUser Then
        If chkRoll.Value <> vbChecked Then chkRoll.Value = vbChecked
        ChangeSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.opt56.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    opt57_Click
'' Description: Change the symbol if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub opt57_Click()
On Error GoTo ErrSection:

    If Me.Visible And m.bFromUser Then
        If chkRoll.Value <> vbChecked Then chkRoll.Value = vbChecked
        ChangeSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.opt57.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optCall_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.optCall_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optCombined_Click
'' Description: Change the symbol to the pit+electronic session symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optCombined_Click()
On Error GoTo ErrSection:

    If Me.Visible And m.bFromUser Then
        ChangeSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.optCombined.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optContains_Click
'' Description: If the user clicks on the contains option, refilter the grid
''              and enable the Find button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optContains_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If Screen.MousePointer = vbHourglass Then Exit Sub
    
    ' TLB 6/22/2016: gets bogged down when click on "Contains" before ready, so wait until they hit "Find"
    'If txtFilterText.Text <> "" Then Filter
    
    cmdFind.Enabled = True
    
    MoveFocus txtFilterText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.optContains.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDescBeginsWith_Click
'' Description: If the user clicks on the "Description Begins with" option,
''              sort the grid on the description column and goto the first
''              description that begins with the text in the Filter text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDescBeginsWith_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If Screen.MousePointer = vbHourglass Then Exit Sub
    cmdFind.Enabled = False
    cmdOK.Default = True
    
    m.lSortedCol = 2
    Filter
    Sort edirAscending
    
    txtFilterText_Change
    
    MoveFocus txtFilterText
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.optDescBeginsWith.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optElectronic_Click
'' Description: Change the symbol to the electronic-only session symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optElectronic_Click()
On Error GoTo ErrSection:

    If Me.Visible And m.bFromUser Then
        ChangeSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.optElectronic.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optElectronicDay_Click
'' Description: Change the symbol to the electronic day session symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optElectronicDay_Click()
On Error GoTo ErrSection:

    If Me.Visible And m.bFromUser Then
        ChangeSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.optElectronicDay.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optPit_Click
'' Description: Change the symbol to the pit-only session symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optPit_Click()
On Error GoTo ErrSection:

    If Me.Visible And m.bFromUser Then
        ChangeSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.optPit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optPut_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.optPut_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSymbolBeginsWith_Click
'' Description: If the user clicks on the "Symbol Begins with" option, sort
''              the grid on the symbol column and goto the first symbol that
''              begins with the text in the Filter text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSymbolBeginsWith_Click()
On Error GoTo ErrSection:

    Dim nRow&

    If Not Me.Visible Then Exit Sub
    If Screen.MousePointer = vbHourglass Then Exit Sub
    cmdFind.Enabled = False
    cmdOK.Default = True
    
    ' TLB 1/15/2010: if coming back to this after searching by desc, save the found symbol
    nRow = fgSymbols.Row
    If nRow >= fgSymbols.FixedRows And nRow < fgSymbols.Rows Then
        txtFilterText.Text = fgSymbols.TextMatrix(nRow, kSymbolCol)
    End If
    
    m.lSortedCol = kSymbolCol
    Filter
    Sort edirAscending
    
    txtFilterText_Change
    
    MoveFocus txtFilterText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.optSymbolBeginsWith.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFilterText_Change
'' Description: As the user types characters in the text box, if the Begins
''              With option is selected, go to the first match in the grid for
''              what the user is typing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFilterText_Change()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in grid to go to
    Dim strText As String               ' Text from the filter text box
    Dim strBase As String               ' Base symbol (if an option)
    Dim strOption As String             ' Option part of the symbol
    Dim nContract As Long
    
    If m.SymbolGrid Is Nothing Then Exit Sub
    
    If Not optContains Then
        m.bSkipRowChangeEvent = True
        If m.bDescending = True Then Sort edirAscending
        strText = txtFilterText.Text
        strBase = Parse(strText, " ", 1)
        strOption = Parse(strText, " ", 2)
        nContract = ContractFromCode(strOption)
        
        ' default to continuous contract
        If Right(strText, 1) = "-" Then
            strText = strText & "067"
        ' else convert if using letter codes for month-year (e.g. "U4")
        ElseIf nContract > 0 Then
            strText = strBase & "-" & Str(nContract)
            txtFilterText.Text = UCase(strText)
            txtFilterText.SelStart = Len(strText)
            Exit Sub
        End If
        
        ' TLB 4/27/2011: allow entering the SymbolID
        If IsDigit(strText, 1) Then
            strText = GetSymbol(Val(strText))
            Me.Caption = strText
        ElseIf Me.Caption <> m.strCaption Then
            Me.Caption = m.strCaption ' (just to make sure caption is restored)
        End If
        
        With fgSymbols
            lRow = m.SymbolGrid.Search(strText)
            If lRow <> -1& Then
                .Row = lRow
                .RowSel = lRow
                .ShowCell lRow, m.lSortedCol
            End If
        End With
        
        If m.bAllowOptions And optSymbolBeginsWith And Not m.bUseOptionWizard Then
            If InStr(strBase, "-") <> 0 Then
                ' Future option (to be done later?)
            Else
                ' Stock option
                With fgMonths
                    If Len(strOption) >= 1 Then
                        lRow = Asc(UCase(Left(strOption, 1))) - Asc("A") + 1
                    Else
                        lRow = -1
                    End If
                    If lRow >= .FixedRows And lRow < .Rows Then
                        .Row = lRow
                        .RowSel = lRow
                        .ShowCell lRow, MonthCol(eGDMonthCol_Letter)
                    Else
                        If .SelectedRows > 0 Then
                            .IsSelected(.SelectedRow(0)) = False
                        End If
                        .Row = -1
                        .ShowCell 1, MonthCol(eGDMonthCol_Letter)
                    End If
                End With
                        
                With fgStrikes
                    If Len(strOption) >= 2 Then
                        lRow = Asc(UCase(Mid(strOption, 2, 1))) - Asc("A") + 1
                    Else
                        lRow = -1
                    End If
                    If lRow >= .FixedRows And lRow < .Rows Then
                        .Row = lRow
                        .RowSel = lRow
                        .ShowCell lRow, MonthCol(eGDMonthCol_Letter)
                    Else
                        If .SelectedRows > 0 Then
                            .IsSelected(.SelectedRow(0)) = False
                        End If
                        .Row = -1
                        .ShowCell 1, MonthCol(eGDMonthCol_Letter)
                    End If
                End With
            End If
        End If
        m.bSkipRowChangeEvent = False
    End If
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.txtFilterText_Change"
    Resume ErrExit
    
End Sub

Private Sub txtFilterText_Click(Button As Integer)
On Error GoTo ErrSection:

    LockOptionPart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.txtFilterText_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFilterText_GotFocus
'' Description: When the filter text box gets the focus, make the Find button
''              the default if it is enabled
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFilterText_GotFocus()
On Error GoTo ErrSection:

    If optPut Or optCall Then
        LockOptionPart
        txtFilterText.SelLength = 0
    Else
        If cmdFind.Enabled = True Then cmdFind.Default = True
        SelectAll txtFilterText
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.txtFilterText.GotFocus", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load up the filters combo box with the symbol groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo(Optional ByVal bShowFilters As Boolean = False)
On Error Resume Next

    Dim lIndex As Long                  ' Index for a for loop
    Dim strID As String                 ' Symbol pool ID for the field
    Dim strType As String               ' Type of thing (i.e. Filter, Criteria, etc)
    Dim strPicture As String            ' Picture to use in the combo box
    Dim strSelID As String              ' ID of the currently selected item
    Dim bSelExists As Boolean           ' Old selection still exists
    Dim iSortStart As Long              ' Where to start the sort
    Dim strItem As String               ' Item to add to the combo box
    Dim aItems As New cGdArray          ' Items to add to the combo box
    Dim obj As Object                   ' Symbol Pool Object
    Dim bScans As Boolean               ' Are we doing scans?
   
    bScans = ScansEnabled
        
    If cboFilters.ComboItems.Count > 0 Then
        strSelID = cboFilters.SelectedItem.Key
        cboFilters.ComboItems.Clear
    End If
    
    ' get list of items to put into combo list
    With g.SymbolPool
        For lIndex = 0 To .ArrayTable.NumFields - 1
            strID = .FieldID(lIndex)
            If Len(strID) = 0 Then
                strType = "" '???
            Else
                strType = Left(strID, 3)
                strPicture = ""
                Set obj = .PoolObject(strID)
                Select Case UCase(strType)
                    Case "GRP"
                        If strID <> "GRP:_FLAGS_.GRP" Then
                            strPicture = ToolbarIcon("ID_SymbolGroups")
                        End If
                    Case "FIL"
                        If bScans And bShowFilters Then
                            strPicture = ToolbarIcon("ID_Filters")
                        End If
                End Select
                If Len(strPicture) > 0 Then
                    If obj.IsActive = True Then
                        If strID = strSelID Then
                            bSelExists = True
                        End If
                        
                        If iSortStart = 0 And lIndex >= g.SymbolPool.OtherFieldsStart Then
                            iSortStart = aItems.Size
                        End If
                        
                        aItems.Add .ArrayTable.FieldName(lIndex) & vbTab _
                                & strID & vbTab & strPicture
                                
                        m.astrComboIDs.Add strID
                    End If
                End If
            End If
        Next
    End With
    If iSortStart > 0 Then
        aItems.Sort eGdSort_IgnoreCase, iSortStart
    End If
    m.astrComboIDs.Sort

    For lIndex = 0 To aItems.Size - 1
        strItem = aItems(lIndex)
        cboFilters.ComboItems.Add , Parse(strItem, vbTab, 2), _
            Parse(strItem, vbTab, 1), Parse(strItem, vbTab, 3)
    Next


    If bSelExists Then
        cboFilters.ComboItems(strSelID).Selected = True
    Else
        cboFilters.ComboItems(1).Selected = True
    End If

    cboFilters.Refresh

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Filter
'' Description: Filter the grid based on the currently selected symbol group
''              in the filters combo box and the filter in the text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Filter() 'Optional ByVal bInitialSort As Boolean = False)
On Error GoTo ErrSection:

    Dim lFilterID As Long               ' ID for the currently selected filter
    Dim lIndex As Long                  ' Index for a for loop
    Dim strID As String                 ' ID for the new pool filter
    Dim lFindLen As Long                ' Length of the filter from the text box
    
    With g.SymbolPool
        lFilterID = .FieldNumForID(cboFilters.SelectedItem.Key)
        Set m.aFilter = .ArrayTable.FieldArray(lFilterID, True)
        
        If optContains.Value = True And txtFilterText.Text <> "" Then
            Screen.MousePointer = vbHourglass
            For lIndex = 0 To m.aFilter.Size - 1
                If m.aFilter(lIndex) = 1 Then
                    If InStr(UCase(.Symbol(lIndex)), UCase(txtFilterText.Text)) = 0 And _
                            InStr(UCase(.Desc(lIndex)), UCase(txtFilterText.Text)) = 0 Then
                        m.aFilter(lIndex) = 0
                    End If
                End If
            Next lIndex
            Screen.MousePointer = vbDefault
        End If
        
        strID = "SymbolPickerFilterID" & Str(m.lFieldNum)
        .FieldID(m.lFieldNum) = strID
        .ArrayTable.AttachField m.aFilter, m.lFieldNum, strID
    End With
    
    With m.SymbolGrid
        .FilterID = strID
        'If bInitialSort Then
            ' initial sort was sometimes not coming up correct,
            ' so this should help
            '.SortOnCol kSymbolCol, 1, 0
        'Else
            .SortOnCol -1&, 0, 0
        'End If
    End With
    
    If optContains.Value = False And txtFilterText.Text <> "" Then
        txtFilterText_Change
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.Filter", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Sort
'' Description: Sort the grid based on column and direction
'' Inputs:      Direction of sort
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Sort(Optional ByVal Direction As eDirection = edirToggle)
On Error GoTo ErrSection:

    Select Case Direction
        Case edirAscending
            m.bDescending = False
        Case edirDescending
            m.bDescending = True
        Case edirToggle
            m.bDescending = Not m.bDescending
    End Select
    
    m.SymbolGrid.SortOnCol m.lSortedCol, Direction, 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.Sort", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrSortCol_Timer
'' Description: After a user has clicked on a column to sort it, wait until
''              they let go of the button, then sort
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrSortCol_Timer()
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Column that user clicked on
    
    If Not MouseIsPressed Then
        tmrSortCol.Enabled = False
        lCol = fgSymbols.MouseCol
        If fgSymbols.MouseRow = 0 And lCol >= 0 And lCol < fgSymbols.Cols Then
            m.lSortedCol = lCol
            Sort edirToggle
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.tmrSortCol.Timer", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitMonthsGrid
'' Description: Initialize the months grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitMonthsGrid()
On Error GoTo ErrSection:

    With fgMonths
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .HighLight = flexHighlightAlways
        .ScrollBars = flexScrollBarVertical
        
        .Rows = 1
        .FixedRows = 1
        .Cols = MonthCol(eGDMonthCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, MonthCol(eGDMonthCol_Letter)) = "Month"
        .TextMatrix(0, MonthCol(eGDMonthCol_Desc)) = "Expiration - Put/Call"
        .Cell(flexcpFontBold, 0, MonthCol(eGDMonthCol_Letter)) = True
        
        .ColAlignment(MonthCol(eGDMonthCol_Letter)) = flexAlignLeftTop
        
        .AddItem "A" & vbTab & "January - Call"
        .AddItem "B" & vbTab & "February - Call"
        .AddItem "C" & vbTab & "March - Call"
        .AddItem "D" & vbTab & "April - Call"
        .AddItem "E" & vbTab & "May - Call"
        .AddItem "F" & vbTab & "June - Call"
        .AddItem "G" & vbTab & "July - Call"
        .AddItem "H" & vbTab & "August - Call"
        .AddItem "I" & vbTab & "September - Call"
        .AddItem "J" & vbTab & "October - Call"
        .AddItem "K" & vbTab & "November - Call"
        .AddItem "L" & vbTab & "December - Call"
        .AddItem "M" & vbTab & "January - Put"
        .AddItem "N" & vbTab & "February - Put"
        .AddItem "O" & vbTab & "March - Put"
        .AddItem "P" & vbTab & "April - Put"
        .AddItem "Q" & vbTab & "May - Put"
        .AddItem "R" & vbTab & "June - Put"
        .AddItem "S" & vbTab & "July - Put"
        .AddItem "T" & vbTab & "August - Put"
        .AddItem "U" & vbTab & "September - Put"
        .AddItem "V" & vbTab & "October - Put"
        .AddItem "W" & vbTab & "November - Put"
        .AddItem "X" & vbTab & "December - Put"
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.InitMonthsGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitStrikesGrid
'' Description: Initialize the strike price grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitStrikesGrid()
On Error GoTo ErrSection:

    With fgStrikes
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .HighLight = flexHighlightAlways
        .ScrollBars = flexScrollBarVertical
        
        .Rows = 1
        .FixedRows = 1
        .Cols = StrikeCol(eGDStrikeCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, StrikeCol(eGDStrikeCol_Letter)) = "Strike"
        .TextMatrix(0, StrikeCol(eGDStrikeCol_Desc)) = "Option Strike Price"
        .Cell(flexcpFontBold, 0, StrikeCol(eGDStrikeCol_Letter)) = True
        
        .ColAlignment(StrikeCol(eGDStrikeCol_Desc)) = flexAlignLeftTop
        
        .AddItem "A" & vbTab & " 5,  105,  205,  ..."
        .AddItem "B" & vbTab & "10,  110,  210,  ..."
        .AddItem "C" & vbTab & "15,  115,  215,  ..."
        .AddItem "D" & vbTab & "20,  120,  220,  ..."
        .AddItem "E" & vbTab & "25,  125,  225,  ..."
        .AddItem "F" & vbTab & "30,  130,  230,  ..."
        .AddItem "G" & vbTab & "35,  135,  235,  ..."
        .AddItem "H" & vbTab & "40,  140,  240,  ..."
        .AddItem "I" & vbTab & "45,  145,  245,  ..."
        .AddItem "J" & vbTab & "50,  150,  250,  ..."
        .AddItem "K" & vbTab & "55,  155,  255,  ..."
        .AddItem "L" & vbTab & "60,  160,  260,  ..."
        .AddItem "M" & vbTab & "65,  165,  265,  ..."
        .AddItem "N" & vbTab & "70,  170,  270,  ..."
        .AddItem "O" & vbTab & "75,  175,  275,  ..."
        .AddItem "P" & vbTab & "80,  180,  280,  ..."
        .AddItem "Q" & vbTab & "85,  185,  285,  ..."
        .AddItem "R" & vbTab & "90,  190,  290,  ..."
        .AddItem "S" & vbTab & "95,  195,  295,  ..."
        .AddItem "T" & vbTab & "100,  200,  300,  ..."
        .AddItem "U" & vbTab & " 7.5,  37.5,  67.5,  ..."
        .AddItem "V" & vbTab & "12.5,  42.5,  72.5,  ..."
        .AddItem "W" & vbTab & "17.5,  47.5,  77.5,  ..."
        .AddItem "X" & vbTab & "22.5,  52.5,  82.5,  ..."
        .AddItem "Y" & vbTab & "27.5,  57.5,  87.5,  ..."
        .AddItem "Z" & vbTab & "32.5,  62.5,  92.5,  ..."
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.InitStrikesGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable/Show/Hide controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim strSymbol$, strOption$, iPos&
    Dim bFuturesFrame As Boolean, iOptionType As Integer
    Dim bValid As Boolean
   
    bValid = True
    
    ' see if it's an option
    If optSymbolBeginsWith And m.bAllowOptions And InStr(txtFilterText.Text, " ") > 0 Then
        ' for an option, get symbol from the text box
        strSymbol = UCase(Trim(txtFilterText.Text))
        iPos = InStr(strSymbol, " ")
        If iPos > 0 Then
            strOption = Trim(Mid(strSymbol, iPos))
            strSymbol = Trim(Left(strSymbol, iPos))
        End If
        ' show the options grids if it's a stock option
        If SecurityType(strSymbol) = "F" Then
            ' future option
            If Len(strOption) > 9 Or InStr("PC", UCase(Left(strOption, 1))) = 0 Or InStr(strOption, " ") > 0 Then
                bValid = False
            End If
            If m.bUseOptionWizard Then
                iOptionType = -1
            End If
        Else
            ' stock or index option
            If Len(strOption) < 2 Then '' <> 2 Then
                bValid = False
            End If
            iOptionType = 1
        End If
    ElseIf fgSymbols.Row >= fgSymbols.FixedRows Then
        ' if it's not an option, get the symbol from the symbol grid
        strSymbol = fgSymbols.TextMatrix(fgSymbols.Row, kSymbolCol)
        ' show the futures stuff if it's a future
        If SecurityType(strSymbol) = "F" Then
            bFuturesFrame = True
        End If
    End If
    
    ' Stock Options grids
    If iOptionType <> 0 Then
        bFuturesFrame = False
        fgSymbols.Visible = False
        If m.bUseOptionWizard Then
            fgMonths.Visible = False
            fgStrikes.Visible = False
            fraOption.Visible = True
            If optPut Or optCall Then
                If Not txtStrike.Enabled Then
                    txtStrike.Enabled = True
                    txtStrike.BackColor = txtFilterText.BackColor
                    If iOptionType < 0 Then
                        ' future option
                        fraOption.Caption = "For a Future Option symbol ..."
                        lblExp.Caption = "Select the Option Month:"
                        mvExp.Visible = False
                        dtMonth.Visible = True
                        If Right(strSymbol, 1) = "-" Then
                            strSymbol = strSymbol & "067"
                        End If
                        strSymbol = RollSymbolForDate(strSymbol)
                        iPos = Val(Parse(strSymbol, "-", 2))
                        If iPos < 200001 Or iPos > 290001 Then
                            iPos = Year(Date) * 100 + Month(Date)
                        End If
                        dtMonth.YYYYMMDD = iPos * 100 + 1
                    Else
                        ' stock option
                        fraOption.Caption = "For a Stock Option symbol ..."
                        lblExp.Caption = "Select the Option Expiration Date:"
                        mvExp.Value = 0
                        mvExp.Visible = True
                        dtMonth.Visible = False
                    End If
                    lblExp.Visible = True
                    lblNote.Visible = True
                End If
                FixExpDate
            Else
                txtStrike.Enabled = False
                txtStrike.BackColor = optCall.BackColor
                mvExp.Visible = False
                dtMonth.Visible = False
                lblExp.Visible = False
                lblNote.Visible = False
            End If
        Else
            fraOption.Visible = False
            fgMonths.Visible = True
            If Len(strOption) = 0 Then
                fgStrikes.Visible = False
            Else
                fgStrikes.Visible = True
            End If
        End If
    Else
        fgSymbols.Visible = True
        fgMonths.Visible = False
        fgStrikes.Visible = False
        fraOption.Visible = False
        optCall = False
        optPut = False
    End If
    
    ' Futures grids
    If bFuturesFrame Then
        fraFutures.Visible = True
        Form_Resize
        If fgSymbols.Row >= fgSymbols.FixedRows Then
            fgSymbols.ShowCell fgSymbols.Row, kSymbolCol
        End If
        FillContinuousInfo
        FillSessionInfo
    Else
        fraFutures.Visible = False
        Form_Resize
        If fgSymbols.Row >= fgSymbols.FixedRows Then
            fgSymbols.ShowCell fgSymbols.Row, kSymbolCol
        End If
    End If

    cmdOK.Enabled = bValid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSessionInfo
'' Description: Fill the session information frame
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillSessionInfo()
On Error GoTo ErrSection:

    Dim strBaseSymbol As String         ' Base symbol of the current symbol in the grid
    Dim lPos As Long                    ' Position of symbol in the lookup string
    Dim lStart As Long                  ' Start of "record" in the lookup string
    Dim lEnd As Long                    ' End of "record" in the lookup string
    Dim strRecord As String             ' "Record" of the base symbol
    
    m.bFromUser = False
    
    strBaseSymbol = Parse(fgSymbols.TextMatrix(fgSymbols.Row, kSymbolCol), "-", 1)
    lPos = InStr(m.strSymbolMap, "," & strBaseSymbol & ",")
    If lPos > 0 Then
        lStart = InStrRev(m.strSymbolMap, "|", lPos)
        lEnd = InStr(lPos, m.strSymbolMap, "|")
        strRecord = Mid(m.strSymbolMap, lStart + 1, (lEnd - 1) - (lStart + 1) + 1)
    Else
        strRecord = ""
    End If
    
    optPit.Caption = Parse(strRecord, ",", 2)
    Enable optPit, Len(optPit.Caption) > 0
    Enable lblPit, Len(optPit.Caption) > 0
    optPit.Value = (optPit.Caption = strBaseSymbol)
    
    optElectronic.Caption = Parse(strRecord, ",", 3)
    Enable optElectronic, Len(optElectronic.Caption) > 0
    Enable lblElectronic, Len(optElectronic.Caption) > 0
    optElectronic.Value = (optElectronic.Caption = strBaseSymbol)
    
    optElectronicDay.Caption = Parse(strRecord, ",", 4)
    Enable optElectronicDay, Len(optElectronicDay.Caption) > 0
    Enable lblElectronicDay, Len(optElectronicDay.Caption) > 0
    optElectronicDay.Value = (optElectronicDay.Caption = strBaseSymbol)
    
    optCombined.Caption = Parse(strRecord, ",", 5)
    Enable optCombined, Len(optCombined.Caption) > 0
    Enable lblCombined, Len(optCombined.Caption) > 0
    optCombined.Value = (optCombined.Caption = strBaseSymbol)

ErrExit:
    m.bFromUser = True
    Exit Sub
    
ErrSection:
    m.bFromUser = True
    RaiseError "frmSymbolSelector.FillSessionInfo", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillContinuousInfo
'' Description: Fill the continuous contract information frame
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillContinuousInfo()
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol from the filter text box
        
    m.bFromUser = False
    strSymbol = fgSymbols.TextMatrix(fgSymbols.Row, kSymbolCol)
    
    If InStr(strSymbol, "-05") <> 0 Or InStr(strSymbol, "-06") <> 0 Then
        chkRoll.Value = vbChecked
        Select Case Parse(strSymbol, "-", 2)
            Case "055"
                chkBackAdjust.Value = vbUnchecked
                opt55.Value = True
                
            Case "056"
                chkBackAdjust.Value = vbUnchecked
                opt56.Value = True
                
            Case "057"
                chkBackAdjust.Value = vbUnchecked
                opt57.Value = True
                
            Case "065"
                chkBackAdjust.Value = vbChecked
                opt55.Value = True
                
            Case "066"
                chkBackAdjust.Value = vbChecked
                opt56.Value = True
                
            Case "067"
                chkBackAdjust.Value = vbChecked
                opt57.Value = True
                
        End Select
    ElseIf InStr(strSymbol, "-0") <> 0 Then
        chkRoll.Value = vbChecked
        opt55.Value = False
        opt56.Value = False
        opt57.Value = False
        chkBackAdjust.Value = vbUnchecked
    Else
        chkRoll.Value = vbUnchecked
        opt55.Value = False
        opt56.Value = False
        opt57.Value = False
        chkBackAdjust.Value = vbUnchecked
    End If

ErrExit:
    m.bFromUser = True
    Exit Sub
    
ErrSection:
    m.bFromUser = True
    RaiseError "frmSymbolSelector.FillContinuousInfo", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeSymbol
'' Description: Change the symbol based on the information in the frames
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeSymbol()
On Error GoTo ErrSection:

    Dim strBaseSymbol As String         ' Base symbol of the current future
    Dim strSymbol As String             ' Symbol to display in the text box
    Dim lRow As Long                    ' Row in the grid for the new symbol
    Dim strDesc As String               ' Description for the symbol
    
    Select Case True
        Case optPit
            strBaseSymbol = optPit.Caption
        Case optElectronic
            strBaseSymbol = optElectronic.Caption
        Case optElectronicDay
            strBaseSymbol = optElectronicDay.Caption
        Case optCombined
            strBaseSymbol = optCombined.Caption
    End Select
    If Len(strBaseSymbol) = 0 Then
        strBaseSymbol = Parse(fgSymbols.TextMatrix(fgSymbols.Row, kSymbolCol), "-", 1)
    End If
    
    If chkRoll.Value = vbChecked Then
        Select Case True
            Case opt55
                If chkBackAdjust.Value = vbChecked Then
                    strSymbol = strBaseSymbol & "-065"
                Else
                    strSymbol = strBaseSymbol & "-055"
                End If
            
            Case opt56
                If chkBackAdjust.Value = vbChecked Then
                    strSymbol = strBaseSymbol & "-066"
                Else
                    strSymbol = strBaseSymbol & "-056"
                End If
            
            Case opt57
                If chkBackAdjust.Value = vbChecked Then
                    strSymbol = strBaseSymbol & "-067"
                Else
                    strSymbol = strBaseSymbol & "-057"
                End If
            
            Case Else
                ' TLB 2/8/2011: else try to append the same continuous contract (e.g. -099, -082, etc)
                strSymbol = Parse(txtFilterText.Text, "-", 2)
                If Len(strSymbol) <> 3 Then
                    strSymbol = "067"
                End If
                strSymbol = strBaseSymbol & "-" & strSymbol
        End Select
    Else
        strSymbol = RollSymbolForDate(strBaseSymbol & "-057", Date)
    End If
        
    If m.lSortedCol = kSymbolCol Then
        lRow = m.SymbolGrid.Search(strSymbol)
    Else
        strDesc = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(strSymbol))
        lRow = m.SymbolGrid.Search(strDesc)
    End If
    
    If lRow <> -1 Then
        fgSymbols.Row = lRow
        fgSymbols.RowSel = lRow
        fgSymbols.ShowCell lRow, kSymbolCol
        If optSymbolBeginsWith.Value = True Then
            txtFilterText.Text = fgSymbols.TextMatrix(fgSymbols.Row, kSymbolCol)
        End If
    End If
    
    MoveFocus fgSymbols
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.ChangeSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumberOfCharts
'' Description: Get the number of charts that are currently up
'' Inputs:      None
'' Returns:     Number of charts
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NumberOfCharts() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Number of charts
    Dim frm As Form                     ' Index into a for loop
    
    lReturn = 0&
    For Each frm In Forms
        'If frm.Name = "frmChart" Then      'JM 06-04-2009: original code; leave awhile then remove if all okay
        If IsFrmChart(frm) Then
            lReturn = lReturn + 1
        End If
    Next frm
    
    NumberOfCharts = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSymbolSelector.NumberOfCharts", eGDRaiseError_Raise
    
End Function

' return contract for month-year code (e.g. "U4" or "4U" returns 200409)
Private Function ContractFromCode(ByVal strCode$) As Long
On Error GoTo ErrSection:
            
    Dim nMonth&, nYear&
            
    If Len(strCode) = 2 Then
        ' swap if digit is first
        If IsDigit(strCode, 1) Then
            strCode = Right(strCode, 1) & Left(strCode, 1)
        End If
        ' check for month codes: FGHJKMNQUVXZ
        If IsDigit(Right(strCode, 1)) Then
            Select Case UCase(Left(strCode, 1))
            Case "F": nMonth = 1
            Case "G": nMonth = 2
            Case "H": nMonth = 3
            Case "J": nMonth = 4
            Case "K": nMonth = 5
            Case "M": nMonth = 6
            Case "N": nMonth = 7
            Case "Q": nMonth = 8
            Case "U": nMonth = 9
            Case "V": nMonth = 10
            Case "X": nMonth = 11
            Case "Z": nMonth = 12
            End Select
            ' get year (range = last year to 8 years into future)
            If nMonth > 0 Then
                nYear = Val(Right(strCode, 1)) + 2000
                While nYear < Year(Date) - 1
                    nYear = nYear + 10
                Wend
            End If
        End If
    End If
    ContractFromCode = nYear * 100 + nMonth

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSymbolSelector.ContractFromCode", eGDRaiseError_Raise
End Function

Private Sub BuildOptionSymbol()
On Error GoTo ErrSection:

    Dim strStrike$
    Dim OptSym As New cOptionSymbol

    If Not m.bUseOptionWizard Then Exit Sub
    If optDescBeginsWith Or optContains Then Exit Sub
    If Not optCall And Not optPut Then Exit Sub
    
#If 1 Then
    With OptSym
        .BaseSymbol = Parse(txtFilterText.Text, " ", 1)
        .IsPut = optPut.Value
        strStrike = Replace(txtStrike.Text, ",", ".") ' to override regional settings
        .Strike = Val(strStrike)
        If .IsFutureOption Then
            .Year = dtMonth.Year
            .Month = dtMonth.Month
        Else
            .Year = mvExp.Year
            .Month = mvExp.Month
            .Day = mvExp.Day
        End If
        txtFilterText.Text = .ToGenesis
    End With
    
#Else
    Dim strSymbol$, strExp$, strMonths$
    Dim bFuture As Boolean
    
    ' get base symbol
    strSymbol = UCase(Parse(txtFilterText.Text, " ", 1))
    If SecurityType(strSymbol) = "F" Then
        bFuture = True
        strExp = Left(Str(dtMonth.YYYYMMDD), 6)
    ElseIf 1 Then
        strExp = Str(mvExp.Year * 10000& + mvExp.Month * 100 + mvExp.Day)
    Else
        ' exp month: need to override regional settings and always use the English months
        strMonths = "  JanFebMarAprMayJunJulAugSepOctNovDec"
        strExp = Mid(strMonths, mvExp.Month * 3, 3)
        strExp = Format(mvExp.Year, "0000") & UCase(strExp) & Format(mvExp.Day, "00")
    End If
    
    ' format the strike price
    strStrike = Format(Val(txtStrike), "#.########")
    strStrike = Replace(strStrike, ",", ".") ' (need to override regional settings)
    If Right(strStrike, 1) = "." Then strStrike = Left(strStrike, Len(strStrike) - 1)
    If Left(strStrike, 1) = "." Then strStrike = "0" & strStrike
    If optCall Then
        strStrike = "C" & strStrike
    ElseIf optPut Then
        strStrike = "P" & strStrike
    Else
        strStrike = ""
    End If
    
    If bFuture Then
        strSymbol = Parse(strSymbol, "-", 1) & "-" & strExp & " " & strStrike
    ElseIf Len(strExp) = 8 Then
        strSymbol = SymbolDisplay(strSymbol & " " & strExp & " " & strStrike)
    Else
        strSymbol = strSymbol & " " & strExp & " " & strStrike
    End If
    txtFilterText.Text = strSymbol
#End If

    MoveFocus txtStrike

ErrExit:
    Set OptSym = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.BuildOptionSymbol"
End Sub

Public Sub LockOptionPart(Optional KeyCode As Integer = 0)
On Error GoTo ErrSection:

    Dim i&
    
    If Not m.bUseOptionWizard Then Exit Sub
    If optDescBeginsWith Or optContains Then Exit Sub
    If Not optCall And Not optPut Then Exit Sub
    
    With txtFilterText
        i = InStr(.Text, " ")
        If i > 1 Then
            If .SelStart >= i Then
                .SelStart = i - 1
                .SelLength = 0
                KeyCode = 0
            ElseIf .SelStart = i - 1 And KeyCode = vbKeyDelete Then
                .Text = Trim(Left(.Text, i))
                .SelStart = Len(.Text)
                .SelLength = 0
                optCall = False
                optPut = False
                KeyCode = 0
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.LockOptionPart"
End Sub

Private Sub FixExpDate()
On Error GoTo ErrSection:
    
    Dim nDate&
    Static nPrevDate&
    
    If SecurityType(Parse(txtFilterText.Text, " ", 1)) <> "F" Then
        With mvExp
            nDate = .Value
            If nDate = 0 Then
                ' if date was set to 0, then initialize to the next exp date from now
                nDate = GetDateFromRule(Year(Date), Month(Date), "3F") + 1
                If Date > nDate Then
                    nDate = nDate + 30
                End If
                nPrevDate = 0
            End If
            If Year(nDate) * 100 + Month(nDate) <> Year(nPrevDate) * 100 + Month(nPrevDate) Then
                nDate = GetDateFromRule(Year(nDate), Month(nDate), "3F") + 1
                .Value = nDate
            End If
            On Error Resume Next
            .DayBold(nPrevDate) = False
            .DayBold(nDate) = True
            nPrevDate = nDate
            If Year(nDate) * 100 + Month(nDate) <> Year(Date) * 100 + Month(Date) Then
                .ShowToday = False
            Else
                .ShowToday = True
            End If
        End With
    End If

    BuildOptionSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.FixExpDate"
End Sub

Private Sub ProcessKey(KeyCode As Integer)
On Error GoTo ErrSection:

    Dim i&, n&

    Select Case KeyCode
    Case Asc("c"), Asc("C")
        KeyCode = 0
        optCall = True
        'BuildSymbol
    Case Asc("p"), Asc("P")
        KeyCode = 0
        optPut = True
        'BuildSymbol
    Case vbKeyUp, vbKeyPageUp
        i = 1
    Case vbKeyDown, vbKeyPageDown
        i = -1
    End Select
    If i <> 0 Then
        KeyCode = 0
        If SecurityType(Parse(txtFilterText.Text, " ", 1)) = "F" Then
            With dtMonth
                n = DateSerial(.Year, .Month, 15)
                n = n + 20 * i
                .Value = n
            End With
        Else
            With mvExp
                n = DateSerial(.Year, .Month, 15)
                n = n + 20 * i
                .Value = n
            End With
        End If
        FixExpDate
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.ProcessKey"
End Sub

Private Function UsingOptWizard() As Boolean
On Error GoTo ErrSection:
    
    If m.bUseOptionWizard Then
        If optSymbolBeginsWith Then
            If optCall Or optPut Then
                UsingOptWizard = True
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSymbolSelector.UsingOptWizard"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFilterText_KeyDown
'' Description: Allow the scroll keys to go to the grid
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFilterText_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LockOptionPart KeyCode
    
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
            KeyCode = 0
            MoveFocus fgSymbols
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.txtFilterText_KeyDown"
    Resume ErrExit
End Sub

Private Sub txtFilterText_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LockOptionPart KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.txtFilterText_KeyPress", 0
    Resume ErrExit
End Sub

Private Sub txtFilterText_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LockOptionPart KeyCode

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.txtFilterText_KeyUp"
    Resume ErrExit
End Sub

Private Sub txtStrike_Change()
On Error GoTo ErrSection:

    BuildOptionSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.txtStrike_Change"
    Resume ErrExit
End Sub

Private Sub txtStrike_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    ProcessKey KeyCode

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolSelector.txtStrike_KeyDown"
    Resume ErrExit
End Sub

