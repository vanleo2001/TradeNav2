VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmActiveTsOrderGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Group Settings"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraQuantityOptions 
      Height          =   795
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   3915
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
      Caption         =   "frmActiveTsOrderGroup.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmActiveTsOrderGroup.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsOrderGroup.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraQuantity 
         Height          =   375
         Left            =   1020
         TabIndex        =   7
         Top             =   0
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
         Caption         =   "frmActiveTsOrderGroup.frx":005C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmActiveTsOrderGroup.frx":0088
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":00A8
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtQty 
            Height          =   315
            Left            =   0
            TabIndex        =   8
            Top             =   23
            Width           =   780
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmActiveTsOrderGroup.frx":00C4
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
            Tip             =   "frmActiveTsOrderGroup.frx":00EE
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmActiveTsOrderGroup.frx":010E
         End
         Begin gdOCX.gdScrollBar sbQty 
            Height          =   360
            Left            =   780
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraLots 
         Height          =   375
         Left            =   1020
         TabIndex        =   11
         Top             =   420
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
         Caption         =   "frmActiveTsOrderGroup.frx":012A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmActiveTsOrderGroup.frx":0156
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":0176
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtLotSize 
            Height          =   315
            Left            =   1860
            TabIndex        =   15
            Top             =   23
            Width           =   780
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmActiveTsOrderGroup.frx":0192
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
            Tip             =   "frmActiveTsOrderGroup.frx":01BC
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmActiveTsOrderGroup.frx":01DC
         End
         Begin HexUniControls.ctlUniTextBoxXP txtLots 
            Height          =   315
            Left            =   0
            TabIndex        =   12
            Top             =   23
            Width           =   780
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmActiveTsOrderGroup.frx":01F8
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
            Tip             =   "frmActiveTsOrderGroup.frx":0222
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmActiveTsOrderGroup.frx":0242
         End
         Begin gdOCX.gdScrollBar sbLots 
            Height          =   360
            Left            =   780
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
         Begin gdOCX.gdScrollBar sbLotSize 
            Height          =   360
            Left            =   2640
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
         Begin HexUniControls.ctlUniLabelXP lblLotSize 
            Height          =   195
            Left            =   1140
            Top             =   83
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
            Caption         =   "frmActiveTsOrderGroup.frx":025E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmActiveTsOrderGroup.frx":0292
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmActiveTsOrderGroup.frx":02B2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniRadioXP optLots 
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "frmActiveTsOrderGroup.frx":02CE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmActiveTsOrderGroup.frx":02FA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":031A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optQuantity 
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   90
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmActiveTsOrderGroup.frx":0336
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmActiveTsOrderGroup.frx":036A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":038A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraCustomTimes 
      Height          =   1455
      Left            =   60
      TabIndex        =   18
      Top             =   3180
      Width           =   3975
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
      Caption         =   "frmActiveTsOrderGroup.frx":03A6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmActiveTsOrderGroup.frx":03F4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsOrderGroup.frx":0414
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkStopTime 
         Height          =   220
         Left            =   180
         TabIndex        =   22
         Top             =   1140
         Width           =   2235
         _ExtentX        =   3942
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
         Caption         =   "frmActiveTsOrderGroup.frx":0430
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmActiveTsOrderGroup.frx":0486
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":04A6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkStartTime 
         Height          =   220
         Left            =   180
         TabIndex        =   20
         Top             =   780
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "frmActiveTsOrderGroup.frx":04C2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmActiveTsOrderGroup.frx":051A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":053A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdStartTime 
         Height          =   315
         Left            =   2520
         TabIndex        =   21
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         ShowDayOfWeek   =   0   'False
         ShowCalendar    =   0   'False
         ShowDate        =   0
         ShowTime        =   2
         MinDate         =   0
         MaxDate         =   0.99999
         Value           =   0
      End
      Begin gdOCX.gdSelectDate gdStopTime 
         Height          =   315
         Left            =   2520
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         ShowDayOfWeek   =   0   'False
         ShowCalendar    =   0   'False
         ShowDate        =   0
         ShowTime        =   2
         MinDate         =   0
         MaxDate         =   0.99999
         Value           =   0
      End
      Begin HexUniControls.ctlUniLabelXP lblCustomTimes 
         Height          =   375
         Left            =   180
         Top             =   240
         Width           =   3675
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
         Caption         =   "frmActiveTsOrderGroup.frx":0556
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmActiveTsOrderGroup.frx":0644
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":0664
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraLoop 
      Height          =   675
      Left            =   60
      TabIndex        =   24
      Top             =   4800
      Width           =   3975
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
      Caption         =   "frmActiveTsOrderGroup.frx":0680
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmActiveTsOrderGroup.frx":06C6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsOrderGroup.frx":06E6
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkLoop 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   270
         Width           =   1875
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
         Caption         =   "frmActiveTsOrderGroup.frx":0702
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmActiveTsOrderGroup.frx":0750
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":0770
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdLoopExp 
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         ShowDayOfWeek   =   0   'False
         ShowPM          =   1
         ShowDate        =   0
         ShowTime        =   2
         MinDate         =   0
         MaxDate         =   0.99999
         Value           =   0
      End
      Begin HexUniControls.ctlUniLabelXP lblEastern 
         Height          =   195
         Left            =   3300
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
         Caption         =   "frmActiveTsOrderGroup.frx":078C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmActiveTsOrderGroup.frx":07BA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":07DA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   780
      TabIndex        =   3
      Top             =   5640
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
      Caption         =   "frmActiveTsOrderGroup.frx":07F6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmActiveTsOrderGroup.frx":0822
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsOrderGroup.frx":0842
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   1320
         TabIndex        =   14
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
         Caption         =   "frmActiveTsOrderGroup.frx":085E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmActiveTsOrderGroup.frx":088C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":08AC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   19
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
         Caption         =   "frmActiveTsOrderGroup.frx":08C8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmActiveTsOrderGroup.frx":08EE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsOrderGroup.frx":090E
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboAccounts 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Top             =   540
      Width           =   2295
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
      Tip             =   "frmActiveTsOrderGroup.frx":092A
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsOrderGroup.frx":094A
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
      Height          =   255
      Left            =   2130
      TabIndex        =   2
      Top             =   150
      Width           =   240
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
      Caption         =   "frmActiveTsOrderGroup.frx":0966
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmActiveTsOrderGroup.frx":0998
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsOrderGroup.frx":09B8
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1440
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmActiveTsOrderGroup.frx":09D4
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
      Tip             =   "frmActiveTsOrderGroup.frx":0A08
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsOrderGroup.frx":0A28
   End
   Begin VSFlex7LCtl.VSFlexGrid fgInputs 
      Height          =   1035
      Left            =   60
      TabIndex        =   17
      Top             =   1980
      Width           =   3915
      _cx             =   6906
      _cy             =   1826
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
   Begin HexUniControls.ctlUniLabelXP lblSymbol 
      Height          =   195
      Left            =   120
      Top             =   180
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
      Caption         =   "frmActiveTsOrderGroup.frx":0A44
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmActiveTsOrderGroup.frx":0A74
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsOrderGroup.frx":0A94
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAccount 
      Height          =   195
      Left            =   120
      Top             =   600
      Width           =   795
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
      Caption         =   "frmActiveTsOrderGroup.frx":0AB0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmActiveTsOrderGroup.frx":0AE2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsOrderGroup.frx":0B02
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmActiveTsOrderGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmActiveTsOrderGroup.frm
'' Description: Form that allows user to setup an active trade sense order group
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/14/2010   DAJ         Turned the Sorted property off on the combo box
'' 06/15/2010   DAJ         Changed icon for the form
'' 06/21/2010   DAJ         Don't convert symbol after symbol lookup (#5813)
'' 07/15/2010   DAJ         Added capabilities for inputs
'' 08/12/2010   DAJ         Fix defaults when chart not on order bar
'' 10/20/2010   DAJ         Continuous Loop
'' 11/04/2010   DAJ         Don't show continuous loop warning unless box checked
'' 11/17/2010   DAJ         Changed ShowLoop to use IB3X code instead of flag file
'' 12/08/2010   DAJ         Allow live trading with enablement code
'' 01/13/2011   DAJ         When determining loop expiration, force it to be weekday
'' 04/25/2011   DAJ         Don't allow TradeSense order groups for options or stocks
'' 05/18/2011   DAJ         Added custom start/stop time for Market1
'' 07/13/2011   DAJ         Allow for true continuous looping if ProjectX
'' 10/17/2011   DAJ         Utilize the auto breakout for TradeSense order groups function
'' 10/03/2012   DAJ         Lot size for forex symbols in TradeSense order groups
'' 10/05/2012   DAJ         Tweaked defaults for quantity/lot size, disable lot size when literal
'' 10/10/2012   DAJ         Changed UI for the quantity/lot size
'' 10/15/2012   DAJ         Fix for quantity controls not being set correctly second time in
'' 01/18/2013   DAJ         Don't allow automated trading for spreads
'' 02/01/2013   DAJ         Don't allow OK if the TSOG/Account/Symbol already exists
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 03/14/2013   DAJ         Changes related to moving Genesis Forex over to literal quantities
'' 07/30/2013   DAJ         Fix for expiration date not being used for TradeSense order group looping ( Rick Freeman )
'' 08/02/2013   DAJ         When setting expiration date, allow use of streaming replay time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDInputsCols
    eGDInputsCol_Name = 0
    eGDInputsCol_OrderNum
    eGDInputsCol_Value
    eGDInputsCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on the OK button?
    
    Quantity As cPriceEditor            ' Editor for quantity
    Lots As cPriceEditor                ' Editor for the number of lots
    LotSize As cPriceEditor             ' Editor for the lot size
    
    strLastSecurityType As String       ' Last security type used
    bShowExpiration As Boolean          ' Show the expiration on the looping?
    
    tsOrderGrp As cTradeSenseOrderGroup ' TradeSense order group object
End Type
Private m As mPrivate

Private Property Get Broker() As eTT_AccountType
    Broker = g.Broker.AccountTypeForID(SelectedAccountID)
End Property

Private Function InputsCol(ByVal nCol As eGDInputsCols) As Long
    InputsCol = nCol
End Function

Private Function ShowLoop() As Boolean
    ShowLoop = HasModule("TSOGLOOP") Or IsIDE
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      TradeSense Order Group, Symbol, Account, Quantity, Lot Size, Inputs,
''              Loop?, Loop Expiration, Custom Start Time, Custom Stop Time
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal tsOrderGroup As cTradeSenseOrderGroup, strSymbol As String, lAccountID As Long, lQuantity As Long, lLotSize As Long, Inputs As cTradeSenseOrderInputs, bLoop As Boolean, dLoopExpiration As Double, dCustomStartTime As Double, dCustomStopTime As Double) As Boolean
On Error GoTo ErrSection:

    Set m.tsOrderGrp = tsOrderGroup

    fgInputs.Visible = (m.tsOrderGrp.Inputs.Count > 0)
    InitInputsGrid
    LoadInputsGrid

    SetControls strSymbol, lAccountID, lQuantity, lLotSize
    
    Form_Resize
    EnableControls
    
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK Then
        strSymbol = txtSymbol.Text
        lAccountID = SelectedAccountID
        
        If optQuantity.Value = True Then
            lQuantity = m.Quantity.Price
            lLotSize = 1&
        ElseIf optLots.Value = True Then
            lQuantity = m.Lots.Price
            lLotSize = m.LotSize.Price
        End If
        
        InputsFromGrid Inputs
        If ShowLoop Then
            bLoop = CheckBoxValue(chkLoop)
            If m.bShowExpiration Then
                dLoopExpiration = Int(CurrentTime("NY", , True)) + Val(gdLoopExp.Value)
                If (bLoop = True) And (dLoopExpiration < g.RealTime.FeedTime) Then
                    InfBox "This will continue to reactivate until the next business date at the expiration time given", "i", , "TradeSense Order Group"
                    dLoopExpiration = dLoopExpiration + 1#
                    Do While Not IsWeekday(dLoopExpiration)
                        dLoopExpiration = dLoopExpiration + 1#
                    Loop
                End If
            Else
                dLoopExpiration = kNullData
            End If
        Else
            bLoop = False
            dLoopExpiration = kNullData
        End If
        If chkStartTime.Value = vbChecked Then
            dCustomStartTime = Round(gdStartTime.Value * 1440#)
        Else
            dCustomStartTime = kNullData
        End If
        If chkStopTime.Value = vbChecked Then
            dCustomStopTime = Round(gdStopTime.Value * 1440#)
        Else
            dCustomStopTime = kNullData
        End If
            
        ' Save off the values to be used as possible defaults the next time...
        SaveLastValues
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmActiveTsOrderGroup.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the user to cancel the dialog
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
    RaiseError "frmActiveTsOrderGroup.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: Allow the user to select a different symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookup_Click()
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.cmdLookup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allow the user to OK the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol from the text box
    Dim lSelectedAccountID As Long      ' Selected account ID
    Dim strKey As String                ' Active TradeSense Order Group key
    
    lSelectedAccountID = SelectedAccountID
    strSymbol = Trim(txtSymbol.Text)
    strKey = strSymbol & vbTab & Str(lSelectedAccountID) & vbTab & m.tsOrderGrp.ID
    
    MoveFocus cmdOK
    If (TypeOfAccount(lSelectedAccountID) = eGDTypeOfAccount_BrokerLive) And (HasModule("TSOGLIVE") = False) Then
        InfBox "You are not authorized to submit or park these orders for a live account", "!", , "Warning"
    ElseIf (TypeOfAccount(lSelectedAccountID) <> eGDTypeOfAccount_Simulated) And (SecurityType(strSymbol) = "S") And (HasModule("AUTOSTK") = False) Then
        MoveFocus txtSymbol
        InfBox "You are not authorized to submit or park these orders for a stock in a live account", "!", , "Warning"
    ElseIf g.TsoGroups.Exists(strKey) Then
        InfBox "You cannot activate|'" & m.tsOrderGrp.Name & "'|on '" & strSymbol & "' in account '" & g.Broker.AccountNameForID(lSelectedAccountID) & "'|because it is already submitted or parked", "!", , "TradeSense Order Groups"
    ElseIf CanActivateAutomatedItem(lSelectedAccountID, txtSymbol.Text, "TradeSense Order Group", "Active TradeSense Order Group") Then
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgInputs_AfterEdit
'' Description: If inputs are linked, update all values
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strName As String               ' Name for the value just changed
    Dim strValue As String              ' Value just entered
    Static bInProgress As Boolean       ' Are we currently updating the grid?

    If Not bInProgress Then
        If Visible Then
            bInProgress = True
            If m.tsOrderGrp.LinkInputs Then
                With fgInputs
                    strName = .TextMatrix(Row, InputsCol(eGDInputsCol_Name))
                    strValue = .TextMatrix(Row, InputsCol(eGDInputsCol_Value))
                    For lIndex = .FixedRows To .Rows - 1
                        If (lIndex <> Row) And (UCase(.TextMatrix(lIndex, InputsCol(eGDInputsCol_Name))) = UCase(strName)) Then
                            .TextMatrix(lIndex, InputsCol(eGDInputsCol_Value)) = strValue
                        End If
                    Next lIndex
                End With
            End If
            bInProgress = False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.fgInputs_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgInputs_BeforeEdit
'' Description: Only allow the user to edit the value column
'' Inputs:      Row, Column, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim tsInput As cTradeSenseOrderInput ' Order input object

    If Col <> InputsCol(eGDInputsCol_Value) Then
        Cancel = True
    Else
        With fgInputs
            If TypeOf .RowData(Row) Is cTradeSenseOrderInput Then
                Set tsInput = .RowData(Row)
                                
                If (tsInput.ParmType = kSN_RetTrueFalse) Or (tsInput.ParmType = kSN_RetTrueFalseConstant) Then
                    .ComboList = "True|False"
                Else
                    .ComboList = ""
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.fgInputs_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize things about the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Icon = Picture16(ToolbarIcon("kTradeSenseOrders"))
    
    fgInputs.Visible = ShowAdvancedTSOG
    fraLoop.Visible = ShowLoop
    
    If g.FractZen.AllowTSOG Then
        chkLoop.Caption = "&Reactivate group"
        m.bShowExpiration = False
        gdLoopExp.Visible = False
        lblEastern.Visible = False
    Else
        chkLoop.Caption = "&Reactivate group until"
        m.bShowExpiration = True
        gdLoopExp.Visible = True
        lblEastern.Visible = True
    End If
    
    Set m.Quantity = New cPriceEditor
    Set m.Lots = New cPriceEditor
    Set m.LotSize = New cPriceEditor
    
    m.strLastSecurityType = ""
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, allow ShowMe to unload the form
'' Inputs:      Cancel Unload?, Mode of the Unload
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
    RaiseError "frmActiveTsOrderGroup.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lHeightDiff As Long             ' Difference between the scale height and the height
    Dim bShowLoop As Boolean            ' Show the loop frame?
    Dim lHeight As Long                 ' Height for the window
    Dim lTop As Long                    ' Where to start the top-most moveable control

    lTop = fgInputs.Top
    lHeightDiff = Height - ScaleHeight
    bShowLoop = ShowLoop
    
    If m.tsOrderGrp.Inputs.Count = 0 Then
        If bShowLoop Then
            lHeight = lTop + fraCustomTimes.Height + fraLoop.Height + fraButtons.Height + (60 * 3) + lHeightDiff
        Else
            lHeight = lTop + fraCustomTimes.Height + fraButtons.Height + (60 * 2) + lHeightDiff
        End If
        
        If Height <> lHeight Then
            Height = lHeight
        End If
        
        With fraCustomTimes
            .Move .Left, lTop
        End With
    
        If bShowLoop Then
            With fraLoop
                .Move .Left, fraCustomTimes.Top + fraCustomTimes.Height + 60
            End With
        End If
    Else
        If bShowLoop Then
            lHeight = lTop + fgInputs.Height + fraCustomTimes.Height + fraLoop.Height + fraButtons.Height + (60 * 4) + lHeightDiff
        Else
            lHeight = lTop + fgInputs.Height + fraCustomTimes.Height + fraButtons.Height + (60 * 3) + lHeightDiff
        End If
        
        If Height <> lHeight Then
            Height = lHeight
        End If
        
        With fgInputs
            .Move .Left, .Top, ScaleWidth - (.Left * 2), .Height
        End With
        
        With fraCustomTimes
            .Move .Left, fgInputs.Top + fgInputs.Height + 60
        End With
    
        If bShowLoop Then
            With fraLoop
                .Move .Left, fraCustomTimes.Top + fraCustomTimes.Height + 60
            End With
        End If
    End If

    With fraButtons
        .Move (ScaleWidth / 2) - (.Width / 2), ScaleHeight - .Height - 60
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.Quantity = Nothing
    Set m.Lots = Nothing
    Set m.LotSize = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLots_Click
'' Description: Handle the user choosing to do lots instead of literal quantities
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLots_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.optLots_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optQuantity_Click
'' Description: Handle the user choosing to do literal quantities instead of lots
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optQuantity_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.optQuantity_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_Click
'' Description: If the user clicks on the text box, bring up the symbol selector
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.txtSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_GotFocus
'' Description: Select the text of the text box when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.txtSymbol_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_KeyPress
'' Description: If the user types in the text box, bring up the symbol selector
'' Inputs:      Ascii Value of Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LookupSymbol KeyAscii
    KeyAscii = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.txtSymbol_KeyPress", 0
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupSymbol
'' Description: Lookup a symbol for the user to trade
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LookupSymbol(Optional ByVal KeyAscii As Long = 0&)
On Error GoTo ErrSection:

    Dim astrSymbol As New cGdArray      ' Array to get lookup symbol from
    Dim strSymbol As String
    
    If KeyAscii = 0& Then
        Set astrSymbol = frmSymbolSelector.ShowMe(txtSymbol.Text, False, True, "Symbol to Buy/Sell")
    Else
        Set astrSymbol = frmSymbolSelector.ShowMe(Chr(KeyAscii), False, True, "Symbol to Buy/Sell", , False)
    End If
    If astrSymbol.Size > 0 Then
        strSymbol = astrSymbol(0) ' ConvertToTradeSymbol(astrSymbol(0), Date)
        
        If strSymbol <> UCase(Trim(txtSymbol.Text)) Then
            If ValidAutomatedSymbol(SelectedAccountID, strSymbol, "TradeSense Order Group", "Active TradeSense Order Group") Then
                txtSymbol.Text = strSymbol
                
                SetQuantityControls
                SetTimeLimits
                SetTimeValues
                
                EnableControls
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.LookupSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable and Show/Hide controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bEnableLotControls As Boolean   ' Enable the lot controls?
    
    bEnableLotControls = AllowLotControls
    
    optLots.Enabled = (bEnableLotControls = True)
    fraLots.Enabled = (bEnableLotControls = True) And (optLots.Value = True)
    txtLots.Enabled = (bEnableLotControls = True) And (optLots.Value = True)
    txtLotSize.Enabled = (bEnableLotControls = True) And (optLots.Value = True)
    
    fraQuantity.Enabled = (optQuantity.Value = True)
    txtQty.Enabled = (optQuantity.Value = True)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetControls
'' Description: Set the controls as appropriate
'' Inputs:      Symbol, Account, Quantity
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetControls(ByVal strSymbol As Variant, ByVal lAccountID As Long, ByVal lQuantity As Long, ByVal lLotSize As Long)
On Error GoTo ErrSection:

    Dim Chart As Form                   ' Active chart form - 6049 (active chart can be frmChart or frmChart2 so need to use generic "form"
    Dim strChartSymbol As String        ' Chart symbol
    Dim strLastSymbol As String         ' Last symbol used
    Dim dLoopExp As Double              ' Loop expiration
    
    Set Chart = ActiveChart
    
    If Len(strSymbol) = 0 Then
        strLastSymbol = GetIniFileProperty("LastSymbol", "", "TsOrderGroup", g.strIniFile)
        
        If Not Chart Is Nothing Then
            strChartSymbol = GetSymbol(Chart.SymbolOrSymbolID)
            If IsSpreadSymbol(strChartSymbol) Then
                txtSymbol.Text = strLastSymbol
            Else
                txtSymbol.Text = strChartSymbol
            End If
        Else
            txtSymbol.Text = strLastSymbol
        End If
    Else
        txtSymbol.Text = strSymbol
    End If
    
    If lAccountID <= 0 Then
        lAccountID = GetIniFileProperty("LastAccount", -1&, "TsOrderGroup", g.strIniFile)
        If Not Chart Is Nothing Then
            If Chart.vseOrderBar.Visible Then
                lAccountID = Chart.TradeAccountID
            End If
        End If
    End If
    PopulateAccountsCbo cboAccounts, lAccountID
    
    dLoopExp = GetIniFileProperty("LastLoopExp", 0#, "TsOrderGroup", g.strIniFile)
    gdLoopExp.Value = dLoopExp
    
    If lLotSize = 1& Then
        SetQuantityControls lQuantity, -1&, lLotSize
    Else
        SetQuantityControls -1&, lQuantity, lLotSize
    End If
    SetTimeLimits
    SetTimeValues
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.SetControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitInputsGrid
'' Description: Initialize the inputs grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitInputsGrid()
On Error GoTo ErrSection:

    With fgInputs
        .Redraw = flexRDNone
        
        SetupGrid fgInputs, eGridMode_Grid
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        
        .Cols = InputsCol(eGDInputsCol_NumCols)
        .FixedCols = 0
        .Rows = 1
        .FixedRows = 1
        
        .TextMatrix(0, InputsCol(eGDInputsCol_Name)) = "Input Name"
        .TextMatrix(0, InputsCol(eGDInputsCol_OrderNum)) = "Order#"
        .TextMatrix(0, InputsCol(eGDInputsCol_Value)) = "Value"
        
        .ColHidden(InputsCol(eGDInputsCol_OrderNum)) = m.tsOrderGrp.LinkInputs
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.InitInputsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadInputsGrid
'' Description: Load the inputs grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadInputsGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsInput As cTradeSenseOrderInput ' TradeSense order input object

    With fgInputs
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To m.tsOrderGrp.Inputs.Count
            Set tsInput = m.tsOrderGrp.Inputs(lIndex)
            
            .Rows = .Rows + 1
            
            .RowData(.Rows - 1) = tsInput.MakeCopy
            .TextMatrix(.Rows - 1, InputsCol(eGDInputsCol_Name)) = tsInput.Name
            .TextMatrix(.Rows - 1, InputsCol(eGDInputsCol_OrderNum)) = Str(tsInput.OrderNumber)
            .TextMatrix(.Rows - 1, InputsCol(eGDInputsCol_Value)) = tsInput.DefaultValue
        Next lIndex
        
        FilterInputsGrid
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.LoadInputsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterInputsGrid
'' Description: Filter the inputs grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterInputsGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw settings
    Dim strPrevName As String           ' Previous input name in the grid
    
    With fgInputs
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Col = InputsCol(eGDInputsCol_OrderNum)
        .Sort = flexSortGenericAscending
        .Col = InputsCol(eGDInputsCol_Name)
        .Sort = flexSortGenericAscending
        
        For lIndex = .FixedRows To .Rows - 1
            If m.tsOrderGrp.LinkInputs = False Then
                .RowHidden(lIndex) = False
            Else
                If UCase(strPrevName) <> UCase(.TextMatrix(lIndex, InputsCol(eGDInputsCol_Name))) Then
                    .RowHidden(lIndex) = False
                    strPrevName = .TextMatrix(lIndex, InputsCol(eGDInputsCol_Name))
                Else
                    .RowHidden(lIndex) = True
                End If
            End If
        Next lIndex
        
        SetBackColors fgInputs
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.FilterInputsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InputsFromGrid
'' Description: Extract the inputs from the grid
'' Inputs:      Inputs
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InputsFromGrid(Inputs As cTradeSenseOrderInputs)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim tsInput As cTradeSenseOrderInput ' TradeSense order input object
    
    If Inputs Is Nothing Then
        Set Inputs = New cTradeSenseOrderInputs
    End If
    
    With fgInputs
        Inputs.Clear
        Inputs.ForGroups = True
        
        For lRow = .FixedRows To .Rows - 1
            If TypeOf .RowData(lRow) Is cTradeSenseOrderInput Then
                Set tsInput = .RowData(lRow)
                tsInput.Value = .TextMatrix(lRow, InputsCol(eGDInputsCol_Value))
                Inputs.Add tsInput
            End If
        Next lRow
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.InputsFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedAccountID
'' Description: Get the currently selected account ID from the combo box
'' Inputs:      None
'' Returns:     Account ID (-1 if none selected)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedAccountID() As Long
On Error GoTo ErrSection:
    
    Dim lReturn As Long                 ' Return value for the function

    lReturn = -1&
    If cboAccounts.ListIndex >= 0 Then
        lReturn = cboAccounts.ItemData(cboAccounts.ListIndex)
    End If
        
    SelectedAccountID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.SelectedAccountID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetQuantityControls
'' Description: Set the quantity controls given the current symbol/account
'' Inputs:      Quantity, Lots, Lot Size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetQuantityControls(Optional ByVal lQuantity As Long = -1&, Optional ByVal lLots As Long = -1&, Optional ByVal lLotSize As Long = -1&)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Selected symbol
    Dim strSecurityType As String       ' Security type for the selected symbol
    Dim Chart As Form                   ' Active chart form
    Dim lLastQuantity As Long           ' Last value for quantity that the user used
    Dim lLastLots As Long               ' Last value for lots that the user used
    Dim lLastLotSize As Long            ' Last value for lot size that the user used
    Dim lChartQuantity As Long          ' Current quantity on the active chart
    Dim strChartSecType As String       ' Current security type of the symbol on the active chart
    
    strSymbol = Trim(txtSymbol.Text)
    strSecurityType = g.Broker.TradeSecType(strSymbol)
    If strSecurityType <> m.strLastSecurityType Then
        m.strLastSecurityType = strSecurityType
        
        Set Chart = ActiveChart
        
        lLastQuantity = GetIniFileProperty("LastQuantity" & strSecurityType, -1&, "TsOrderGroup", g.strIniFile)
        If lLastQuantity = -1& Then
            lLastQuantity = GetIniFileProperty("LastQuantity", -1&, "TsOrderGroup", g.strIniFile)
        End If
        
        lLastLots = GetIniFileProperty("LastLots" & strSecurityType, -1&, "TsOrderGroup", g.strIniFile)
        lLastLotSize = GetIniFileProperty("LastLotSize" & strSecurityType, -1&, "TsOrderGroup", g.strIniFile)
        
        lChartQuantity = -1&
        strChartSecType = ""
        If Not Chart Is Nothing Then
            If Chart.vseOrderBar.Visible Then
                strChartSecType = g.Broker.TradeSecType(Chart.Chart.Symbol)
                If strChartSecType = strSecurityType Then
                    lChartQuantity = CLng(Val(Chart.txtTradeQty.Text))
                End If
            End If
        End If
        
        If lQuantity = -1& Then
            If lLastQuantity = -1& Then
                If lChartQuantity = -1& Then
                    lQuantity = 1&
                Else
                    lQuantity = lChartQuantity
                End If
            Else
                lQuantity = lLastQuantity
            End If
        End If
        If lLotSize = -1& Then
            If lLastLotSize = -1& Then
                If strSecurityType = "S" Then
                    lLotSize = 100&
                Else
                    lLotSize = g.Broker.MinimumLotSize(SelectedAccountID, txtSymbol.Text)
                End If
            Else
                lLotSize = lLastLotSize
            End If
        End If
        If lLots = -1& Then
            If lLastLots = -1& Then
                If lChartQuantity = -1& Then
                    lLots = 1&
                Else
                    lLots = lChartQuantity / lLastLotSize
                    If lLots < 1& Then
                        lLots = 1&
                    End If
                End If
            Else
                lLots = lLastLots
            End If
        End If
        
        g.Broker.InitQuantityEditor m.Quantity, sbQty, txtQty, SelectedAccountID, txtSymbol.Text, lQuantity
        m.Lots.Init sbLots, txtLots, Nothing, lLots, 1&
        g.Broker.InitQuantityEditor m.LotSize, sbLotSize, txtLotSize, SelectedAccountID, txtSymbol.Text, lLotSize
        
        If AllowLotControls Then
            optQuantity.Value = GetIniFileProperty("LastLiteral", False, "TsOrderGroup", g.strIniFile)
            optLots.Value = Not optQuantity.Value
        Else
            optQuantity.Value = True
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.SetQuantityControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTimeLimits
'' Description: Set the min/max time limits for the specified symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTimeLimits()
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol from the UI
    Dim dStartTime As Double            ' Start time for the symbol
    Dim dStopTime As Double             ' Stop time for the symbol
    Dim Bars As cGdBars                 ' Bars to get the properties from
    
    strSymbol = Trim(txtSymbol.Text)
    If Len(strSymbol) > 0 Then
        chkStartTime.Enabled = True
        gdStartTime.Enabled = True
        chkStopTime.Enabled = True
        gdStopTime.Enabled = True
        
        Set Bars = New cGdBars
        SetBarProperties Bars, strSymbol
        
        dStartTime = Bars.Prop(eBARS_DefaultStartTime) / 1440#
        dStopTime = Bars.Prop(eBARS_DefaultEndTime) / 1440#
        
        gdStartTime.MinDate = dStartTime
        gdStartTime.MaxDate = dStopTime
        gdStopTime.MinDate = dStartTime
        gdStopTime.MaxDate = dStopTime
    Else
        chkStartTime.Enabled = False
        gdStartTime.Enabled = False
        chkStopTime.Enabled = False
        gdStopTime.Enabled = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.SetTimeLimits"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTimeValues
'' Description: Set the time values for the specified symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTimeValues()
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol from the UI
    Dim strTimes As String              ' Previous values for the symbol from the INI file
    Dim astrTimes As cGdArray           ' Array of time information split out
    
    strSymbol = Trim(txtSymbol.Text)
    If Len(strSymbol) > 0 Then
        strTimes = GetIniFileProperty(strSymbol, "", "TsOrderGroup", g.strIniFile)
        If Len(strTimes) = 0 Then
            gdStartTime.Value = gdStartTime.MinDate
            chkStartTime.Value = vbUnchecked
            gdStopTime.Value = gdStopTime.MaxDate
            chkStopTime.Value = vbUnchecked
        Else
            Set astrTimes = New cGdArray
            astrTimes.SplitFields strTimes, ","
            
            CheckBoxValue(chkStartTime) = (CLng(Val(astrTimes(0))) = 1)
            gdStartTime.Value = Val(astrTimes(1))
            CheckBoxValue(chkStopTime) = (CLng(Val(astrTimes(2))) = 1)
            gdStopTime.Value = Val(astrTimes(3))
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.SetTimeValues"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveLastValues
'' Description: Save the last values for the next time the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveLastValues()
On Error GoTo ErrSection:

    Dim astrTimes As cGdArray           ' Array of time information
    Dim strSymbol As String             ' Symbol
    Dim strSecurityType As String       ' Security type for the symbol

    strSymbol = Trim(txtSymbol.Text)
    strSecurityType = g.Broker.TradeSecType(strSymbol)
    
    SetIniFileProperty "LastSymbol", strSymbol, "TsOrderGroup", g.strIniFile
    SetIniFileProperty "LastAccount", SelectedAccountID, "TsOrderGroup", g.strIniFile
    SetIniFileProperty "LastQuantity" & strSecurityType, m.Quantity.Price, "TsOrderGroup", g.strIniFile
    If AllowLotControls Then
        SetIniFileProperty "LastLots" & strSecurityType, m.Lots.Price, "TsOrderGroup", g.strIniFile
        SetIniFileProperty "LastLotSize" & strSecurityType, m.LotSize.Price, "TsOrderGroup", g.strIniFile
        SetIniFileProperty "LastLiteral", optQuantity.Value, "TsOrderGroup", g.strIniFile
    End If
    SetIniFileProperty "LastLoopExp", Str(Val(gdLoopExp.Value)), "TsOrderGroup", g.strIniFile
    
    Set astrTimes = New cGdArray
    astrTimes.Create eGDARRAY_Strings, 4
    astrTimes(0) = Str(Abs(CLng(CheckBoxValue(chkStartTime))))
    astrTimes(1) = Str(gdStartTime.Value)
    astrTimes(2) = Str(Abs(CLng(CheckBoxValue(chkStopTime))))
    astrTimes(3) = Str(gdStopTime.Value)
    SetIniFileProperty strSymbol, astrTimes.JoinFields(","), "TsOrderGroup", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.SaveLastValues"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowLotControls
'' Description: Determine whether to allow the lot controls or not
'' Inputs:      None
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllowLotControls() As Boolean
On Error GoTo ErrSection:

    Dim strSecType As String            ' Security type for the selected symbol
    
    strSecType = g.Broker.TradeSecType(Trim(txtSymbol.Text))
    
    AllowLotControls = ((strSecType = "S") Or (m.LotSize.Min > 1))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroup.AllowLotControls"
    
End Function

