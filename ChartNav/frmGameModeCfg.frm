VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmGameModeCfg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replay Settings"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   3900
   Begin HexUniControls.ctlUniFrameWL fraInstant 
      Height          =   4815
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
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
      Caption         =   "frmGameModeCfg.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmGameModeCfg.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmGameModeCfg.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraSymbol 
         Height          =   1335
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   3555
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
         Caption         =   "frmGameModeCfg.frx":0068
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmGameModeCfg.frx":0094
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":00B4
         RightToLeft     =   0   'False
         Begin MSComctlLib.ImageCombo cboFilters 
            Height          =   330
            Left            =   420
            TabIndex        =   23
            Top             =   840
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "ImageCombo1"
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSelectSymbol 
            Height          =   315
            Left            =   3060
            TabIndex        =   20
            Top             =   210
            Width           =   315
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
            Caption         =   "frmGameModeCfg.frx":00D0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmGameModeCfg.frx":00F6
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":0116
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
            Height          =   315
            Left            =   1560
            TabIndex        =   21
            Top             =   210
            Width           =   1695
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmGameModeCfg.frx":0132
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
            Tip             =   "frmGameModeCfg.frx":0152
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":0172
         End
         Begin HexUniControls.ctlUniRadioXP optRandomSym 
            Height          =   255
            Left            =   180
            TabIndex        =   19
            Top             =   600
            Width           =   3015
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
            Caption         =   "frmGameModeCfg.frx":018E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmGameModeCfg.frx":01EE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":020E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optSelectSymbol 
            Height          =   255
            Left            =   180
            TabIndex        =   22
            Top             =   240
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
            Caption         =   "frmGameModeCfg.frx":022A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmGameModeCfg.frx":0266
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":0286
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraOptions 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   1575
         Width           =   3555
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
         Caption         =   "frmGameModeCfg.frx":02A2
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmGameModeCfg.frx":02D0
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":02F0
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboImageXP cboStrategy 
            Height          =   315
            Left            =   1200
            TabIndex        =   13
            Top             =   1020
            Width           =   2175
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
            Tip             =   "frmGameModeCfg.frx":030C
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":032C
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboBarsPeriod 
            Height          =   315
            Left            =   1200
            TabIndex        =   12
            Top             =   630
            Width           =   2175
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
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
            Tip             =   "frmGameModeCfg.frx":0348
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   2
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":0368
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkReplayAll 
            Height          =   255
            Left            =   180
            TabIndex        =   11
            Top             =   1440
            Width           =   3135
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
            Caption         =   "frmGameModeCfg.frx":0384
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmGameModeCfg.frx":03D6
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":03F6
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectDate gdDate 
            Height          =   315
            Left            =   1200
            TabIndex        =   14
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            AllowWeekends   =   0   'False
            MaxDate         =   42611
            MaxDateIsToday  =   -1  'True
            Value           =   37991
         End
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   255
            Left            =   180
            Top             =   300
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
            Caption         =   "frmGameModeCfg.frx":0412
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmGameModeCfg.frx":0448
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":0468
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label2 
            Height          =   255
            Left            =   180
            Top             =   660
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
            Caption         =   "frmGameModeCfg.frx":0484
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmGameModeCfg.frx":04BC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":04DC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   255
            Left            =   180
            Top             =   1020
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
            Caption         =   "frmGameModeCfg.frx":04F8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmGameModeCfg.frx":052A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":054A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraDataReplay 
         Height          =   1275
         Left            =   120
         TabIndex        =   5
         Top             =   3510
         Width           =   3555
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
         Caption         =   "frmGameModeCfg.frx":0566
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmGameModeCfg.frx":05C0
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":05E0
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtMinutes 
            Height          =   285
            Left            =   1140
            TabIndex        =   6
            Top             =   600
            Width           =   495
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmGameModeCfg.frx":05FC
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
            Tip             =   "frmGameModeCfg.frx":061E
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":063E
         End
         Begin HexUniControls.ctlUniRadioXP optMinutes 
            Height          =   255
            Left            =   180
            TabIndex        =   9
            Top             =   600
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
            Caption         =   "frmGameModeCfg.frx":065A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmGameModeCfg.frx":068A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":06AA
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optDay 
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   900
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
            Caption         =   "frmGameModeCfg.frx":06C6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmGameModeCfg.frx":0708
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":0728
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optAuto 
            Height          =   255
            Left            =   180
            TabIndex        =   7
            Top             =   300
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
            Caption         =   "frmGameModeCfg.frx":0744
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmGameModeCfg.frx":07AE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmGameModeCfg.frx":07CE
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   435
      Left            =   180
      TabIndex        =   2
      Top             =   6360
      Width           =   3555
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
      Caption         =   "frmGameModeCfg.frx":07EA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmGameModeCfg.frx":0816
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmGameModeCfg.frx":0836
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   0
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
         Caption         =   "frmGameModeCfg.frx":0852
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmGameModeCfg.frx":0880
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":08A0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdStartGame 
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   0
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
         Caption         =   "frmGameModeCfg.frx":08BC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmGameModeCfg.frx":08F6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0916
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdResults 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   0
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
         Caption         =   "frmGameModeCfg.frx":0932
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmGameModeCfg.frx":0962
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0982
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraStreaming 
      Height          =   4695
      Left            =   120
      TabIndex        =   27
      Top             =   1500
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
      Caption         =   "frmGameModeCfg.frx":099E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmGameModeCfg.frx":09DC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmGameModeCfg.frx":09FC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniListBoxXP lstSymbols 
         Height          =   2985
         Left            =   2040
         TabIndex        =   15
         Top             =   1140
         Width           =   1515
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
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
         TrapTab         =   0   'False
         Tip             =   "frmGameModeCfg.frx":0A18
         MultiSelect     =   0
         Sorted          =   0   'False
         HScroll         =   0   'False
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
         RoundedBorders  =   0   'False
         SelectorStyle   =   -1
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0A38
         ManualStart     =   0   'False
         Columns         =   0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniListBoxXP lstSessions 
         Height          =   3570
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1755
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
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
         TrapTab         =   0   'False
         Tip             =   "frmGameModeCfg.frx":0A54
         MultiSelect     =   0
         Sorted          =   0   'False
         HScroll         =   0   'False
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
         RoundedBorders  =   0   'False
         SelectorStyle   =   -1
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0A74
         ManualStart     =   0   'False
         Columns         =   0
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate dtDate 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         AllowWeekends   =   0   'False
         MaxDate         =   42611
         MaxDateIsToday  =   -1  'True
         Value           =   39034
      End
      Begin gdOCX.gdSelectDate dtTime 
         Height          =   315
         Left            =   2370
         TabIndex        =   17
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         ShowDayOfWeek   =   0   'False
         ShowDate        =   0
         ShowTime        =   2
         MinDate         =   0
         MaxDate         =   0.99999
         Value           =   0.506944444444444
      End
      Begin HexUniControls.ctlUniLabelXP Label11 
         Height          =   435
         Left            =   1980
         Top             =   4170
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
         Caption         =   "frmGameModeCfg.frx":0A90
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameModeCfg.frx":0B0A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0B2A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label10 
         Height          =   435
         Left            =   2100
         Top             =   720
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
         Caption         =   "frmGameModeCfg.frx":0B46
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameModeCfg.frx":0BAE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0BCE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   195
         Left            =   2340
         Top             =   0
         Width           =   1035
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
         Caption         =   "frmGameModeCfg.frx":0BEA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameModeCfg.frx":0C20
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0C40
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label6 
         Height          =   255
         Left            =   120
         Top             =   720
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
         Caption         =   "frmGameModeCfg.frx":0C5C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameModeCfg.frx":0CA4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0CC4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraMode 
      Height          =   1275
      Left            =   120
      TabIndex        =   24
      Top             =   120
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
      Caption         =   "frmGameModeCfg.frx":0CE0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmGameModeCfg.frx":0D16
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmGameModeCfg.frx":0D36
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optInstant 
         Height          =   220
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
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
         Caption         =   "frmGameModeCfg.frx":0D52
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmGameModeCfg.frx":0D90
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0DB0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optStreaming 
         Height          =   225
         Left            =   180
         TabIndex        =   25
         Top             =   780
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "frmGameModeCfg.frx":0DCC
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmGameModeCfg.frx":0E0E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0E2E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label9 
         Height          =   195
         Left            =   1860
         Top             =   300
         Width           =   1635
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
         Caption         =   "frmGameModeCfg.frx":0E4A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameModeCfg.frx":0E96
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0EB6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label8 
         Height          =   195
         Left            =   660
         Top             =   480
         Width           =   2715
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
         Caption         =   "frmGameModeCfg.frx":0ED2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameModeCfg.frx":0F36
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0F56
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label7 
         Height          =   195
         Left            =   2100
         Top             =   780
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
         Caption         =   "frmGameModeCfg.frx":0F72
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameModeCfg.frx":0FBA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":0FDA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   195
         Left            =   660
         Top             =   960
         Width           =   2715
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
         Caption         =   "frmGameModeCfg.frx":0FF6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmGameModeCfg.frx":1064
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmGameModeCfg.frx":1084
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmGameModeCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmGameModeCfg.frm
'' Description: Allow the user to configure game mode
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/01/2013   DAJ         Changed code for loading strategies
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kCmdButtonsTop = 4980
Private Const kFormHeight = 5955

Private Type mPrivate
    tbStrategyInfo As New cGdTable  'fields:0=system name,1=system ID,2=lib ID
    oGameModeObj As New cGameMode
    strRandom As String             'values are: "random symbol", "user select", or symbol group string
    strBarsPeriod As String         '1 minute ... yearly
    strReplaySpeed As String        'auto, minutes, day
    nReplayMinutes As Long
    dStartDate As Double
    bAbortSearch As Boolean
    nEarliestDate As Long   ' earliest date available (from Symbols.RTS file)
End Type

Private m As mPrivate

Private Sub cboBarsPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub cboFilters_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub cboStrategy_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub chkReplayAll_Click()
    m.oGameModeObj.ReplayAll = chkReplayAll.Value
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub cmdResults_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub cmdSelectSymbol_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub cmdStartGame_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub dtDate_Changed()
On Error GoTo ErrSection:

    SetStreamDate dtDate.Value

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmGameModeCfg.dtDate_Changed"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub lstSessions_Click()
On Error GoTo ErrSection:

    SetStreamDate

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmGameModeCfg.lstSessions_Click"
End Sub

Private Sub lstSymbols_Click()
    On Error Resume Next
    lstSymbols.ListIndex = -1
End Sub

Private Sub optAuto_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub optDay_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub optInstant_Click()
    FixFrames
End Sub

Private Sub optMinutes_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub optRandomSym_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub optSelectSymbol_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub optStreaming_Click()
    
    If CanDoStreamReplay(True) Then
        FixFrames
    Else
        optInstant = True
    End If
    
End Sub

Private Sub txtMinutes_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowHelp KeyCode
End Sub

Private Sub txtSymbol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ShowHelp KeyCode
    Else
        ShowSymSelector (UCase(Chr(KeyCode)))
    End If
End Sub

Private Sub ShowHelp(KeyCode As Integer)
On Error Resume Next:
    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Nothing
    End If
End Sub

Public Sub ShowMe(Optional ByVal dStreamDate As Double = 0)
On Error GoTo ErrSection:

    Dim i&, strMsg$, bStreaming As Boolean
    Dim bNew As Boolean

    If g.ChartGlobals.nGameInProg > 0 Then
        bNew = PromptStartNew()
        DoEvents
        If bNew Then
            Set m.oGameModeObj = New cGameMode
        Else
            Unload Me
            Exit Sub
        End If
    End If
    
    m.oGameModeObj.ResetGameMode
        
    'read saved info from INI
    m.strRandom = GetIniFileProperty("Random", "random symbol", "Game Mode", g.strIniFile)
    m.strBarsPeriod = GetIniFileProperty("BarsPeriod", "Daily", "Game Mode", g.strIniFile)
    m.strReplaySpeed = GetIniFileProperty("ReplaySpeed", "auto", "Game Mode", g.strIniFile)
    m.nReplayMinutes = GetIniFileProperty("ReplayMinutes", 5, "Game Mode", g.strIniFile)
    m.dStartDate = GetIniFileProperty("ReplayStart", DateSerial(2003, 12, 1), "Game Mode", g.strIniFile)
    m.oGameModeObj.ReplayAll = GetIniFileProperty("ReplayAll", 0, "Game Mode", g.strIniFile)
    bStreaming = GetIniFileProperty("StreamingMode", False, "Game Mode", g.strIniFile)
    
    'symbol info
    SetPlaySymbol
    'symbol groups info
    cboFilters.ImageList = frmMain.img16
    cboFilters.Locked = True
    LoadCombo
    
    'symbol option
    If m.strRandom = "user select" Then
        optSelectSymbol.Value = True
    Else
        optRandomSym.Value = True
    End If
    EnableSelectSym optSelectSymbol.Value

    m.oGameModeObj.GameDataTime = m.dStartDate
    gdDate.Value = Int(m.oGameModeObj.GameDataTime)
    
    ' Data replay options:
    ' if has intraday data ...
    If HasModule("IT") Or HasModule("FT") Or HasModule("ST") Then
        ' if can do streaming replay (T_SRP or has bought Gold and has streaming) ...
        If CanDoStreamReplay(False) Then
            SetStreamDate dStreamDate
            If dStreamDate > 0 Then
                optStreaming.Value = True
                If g.bShowInLocalTimeZone Then
                    dStreamDate = ConvertTimeZone(dStreamDate, "NY", "")
                End If
                dtTime.Value = dStreamDate - Int(dStreamDate)
            Else
                If bStreaming Then
                    optStreaming.Value = True
                Else
                    optInstant.Value = True
                End If
                If g.bShowInLocalTimeZone Then
                    dtTime.Value = ConvertTimeZone(570 / 1440#, "NY", "")
                Else
                    dtTime.Value = 570 / 1440#
                End If
            End If
        Else
            optInstant.Value = True
            'optStreaming.Enabled = False
        End If
        
        LoadStreamFiles
        txtMinutes = Str(m.nReplayMinutes)
        If m.strReplaySpeed = "minutes" Then
            m.oGameModeObj.GameAutoInterval = False
            optMinutes = True
        ElseIf m.strReplaySpeed = "day" Then
            m.oGameModeObj.GameAutoInterval = False
            optDay = True
        Else
            m.oGameModeObj.GameAutoInterval = True
            optAuto = True
        End If
    Else
        ' if no intraday data, then can only do Instant one-day-at-a-time
        optInstant.Value = False
        fraMode.Visible = False
        fraDataReplay.Visible = False
        fraInstant.Top = 60
        fraButtons.Top = fraInstant.Top + fraOptions.Top + fraOptions.Height + 240
        Me.Height = fraButtons.Top + fraButtons.Height + (Me.Height - Me.ScaleHeight)
        optDay = True
    End If
    fraStreaming.Move fraMode.Left, fraMode.Top + fraMode.Height + 120
    FixFrames
    
    'set results & replay all buttons
    cmdResults.Enabled = m.oGameModeObj.HasResultFiles
    If g.RealTime.Active Then
        chkReplayAll.Value = 0
        chkReplayAll.Enabled = False
    Else
        chkReplayAll.Value = Abs(m.oGameModeObj.ReplayAll)
        chkReplayAll.Enabled = True
    End If
    
    LoadBarsPeriodCbo
    LoadStrategyCbo
        
    CenterTheForm Me
    'ShowForm Me, False, frmMain
    ShowForm Me, True, frmMain
    
    Exit Sub

ErrSection:
    RaiseError "frmGameModeCfg.ShowMe"

End Sub

Private Sub ShowSymSelector(ByVal strChar)
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbol(s) back from the symbol selector
    
    If Len(strChar) = 0 Then
        Set astrSymbols = frmSymbolSelector.ShowMe("", False)
    Else
        Set astrSymbols = frmSymbolSelector.ShowMe(strChar, False, , , , False, True)
    End If
    
    If astrSymbols.Size > 0 Then
        txtSymbol.Text = astrSymbols(0)
    End If
    
    Set astrSymbols = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "frmGameModeCfg.ShowSymSelector"
    
End Sub

Private Sub cmdCancel_Click()
    If cmdCancel.Caption = "Abort" Then
        m.bAbortSearch = True
    Else
        Unload Me
    End If
End Sub

Private Sub cmdResults_Click()
On Error GoTo ErrSection:
    
    ' need to hide this modal window before showing the non-modal results window
    Me.Hide
    If Not m.oGameModeObj Is Nothing Then
        DoEvents
        m.oGameModeObj.ShowGameResult
        DoEvents
    End If
    Unload Me
    
    Exit Sub

ErrSection:
    RaiseError "frmGameModeCfg.cmdResults_Click"
End Sub

Private Sub cmdSelectSymbol_Click()
    ShowSymSelector ""
End Sub

Private Sub cmdStartGame_Click()
On Error GoTo ErrSection:

    Dim frm As Form
    Dim strSymbol$, nStrategyId&
    Dim dBarsDate#, strInf$
'    Dim dDateNow#, iMaxDays&
                    
    Dim strOtherSym$, i&
    Dim bExit As Boolean
                    
    
    If optStreaming Then
        If dtDate.Value < m.nEarliestDate Then
            InfBox "The earliest date available for streaming replay| is: " & DateFormat(m.nEarliestDate), "e", , "Streaming Replay"
            Exit Sub
        End If
        SetIniFileProperty "StreamingMode", True, "Game Mode", g.strIniFile
        Me.Hide
        DoEvents
        frmReplay.ShowMe
        dBarsDate = dtDate.Value + dtTime.Value
        If g.bShowInLocalTimeZone Then
            dBarsDate = ConvertTimeZone(dBarsDate, "", "NY")
        End If
        frmReplay.Play dBarsDate
        Unload Me
        Exit Sub
    End If
    SetIniFileProperty "StreamingMode", False, "Game Mode", g.strIniFile
        
    cmdStartGame.Enabled = False
    DoEvents
    
    If optSelectSymbol.Value = True Then
        strSymbol = txtSymbol.Text
        If Not EnoughData(strSymbol, True) Then
            cmdStartGame.Enabled = True
            Exit Sub
        End If
    ElseIf optRandomSym.Value = True Then
        cmdCancel.Caption = "Abort"
        cmdCancel.SetFocus
        DoEvents
        strSymbol = RandomSymbol()
    End If
    
    If chkReplayAll.Value = 1 Then
        For i = 0 To Forms.Count - 1
            If bExit Then Exit For
            If IsFrmChart(Forms(i)) Then
                Set frm = Forms(i)
                strOtherSym = frm.Chart.Symbol
                If Len(frm.Chart.SpreadSymbols) = 0 Then
                    If Not EnoughData(strOtherSym, True) Then bExit = True
                Else
                    strInf = "Cannot replay spread chart: " & strOtherSym & "."
                    InfBox strInf, "I", , "Instant Replay"
                    bExit = True
                End If
            End If
        Next
    End If
    
    If bExit Then
        cmdStartGame.Enabled = True
        Exit Sub
    End If
    
    m.dStartDate = DateSerial(gdDate.Year, gdDate.Month, gdDate.Day)
    m.nReplayMinutes = Val(txtMinutes.Text)
    m.strBarsPeriod = cboBarsPeriod.Text
    SaveSettings
       
    cmdStartGame.Enabled = True
    cmdCancel.Caption = "Cancel"
    
    If strSymbol = "Insufficient Data" Or strSymbol = "Try Again" Then
        DoEvents
        Exit Sub
    Else
        Me.Hide
        DoEvents
    End If
        
    If Len(strSymbol) > 0 Then
        If g.RealTime.Active Then
            m.oGameModeObj.ReplayAll = False
        Else
            m.oGameModeObj.ReplayAll = chkReplayAll.Value
        End If
        m.oGameModeObj.GameDataTime = DateSerial(gdDate.Year, gdDate.Month, gdDate.Day) 'gdDate.Value
        m.oGameModeObj.GameAutoInterval = optAuto
        If optMinutes Then
            m.oGameModeObj.GameInterval = m.nReplayMinutes
        ElseIf optDay Then
            m.oGameModeObj.GameInterval = 1440
        Else
            m.oGameModeObj.GameInterval = 0
        End If
        
        If optRandomSym.Value = True Then
            m.oGameModeObj.GameRandomSym = True
        Else
            m.oGameModeObj.GameRandomSym = False
        End If
        Set frm = New frmChart                  'instant replay always starts as non-detached chart
        frm.IsInGameMode = True                 '4696 (replay bar not shown when maximized)
        frm.Chart.Template = "Replay"
        frm.Chart.SetSymbol strSymbol
                
'JM: 02-09-2009 - This is inefficient way to fix issue 3734. Leave awhile then remove if all okay.
'        dDateNow = Now              'aardvark 3734 fix (3735 is duplicate issue)
'        iMaxDays = Int((dDateNow - m.oGameModeObj.GameDataTime) * 5 / 7) + 3
'        If iMaxDays > 0 And iMaxDays < 999 Then
'            frm.Chart.MaxIntradayDays = iMaxDays
'        End If
        
        frm.Chart.ChangeBarPeriod cboBarsPeriod.Text, False, True
        nStrategyId = cboStrategy.ItemData(cboStrategy.ListIndex)
        m.oGameModeObj.GameStrategyID = nStrategyId
        m.oGameModeObj.GameStrategyName = cboStrategy.Text
        If frm.Chart.SystemID <> nStrategyId Then
            frm.Chart.SystemID = nStrategyId
        End If
        frm.Chart.ShowTrades = 0            '4311
        SetBarProperties frm.Chart.Bars, frm.Chart.Symbol, True     '4843 - cause of bug: StartTime & EndTime properties were missing
        
        frm.Chart.Bars.Prop(eBARS_Periodicity) = frm.Chart.Periodicity      '5943
        
        
If 0 Then
        'need to do initial load to get data into bars and chart object
        frm.Chart.GenerateChart eRedo9_ReloadData
        
'JM: 02-09-2009 - Issue 4570 is not because of "new way of loading data".
'   Actual cause is due to inefficient fixe for issue 3734 above.
        If frm.Chart.IsPartiallyLoaded Then
            InfBox "Loading data for Instant Replay ...", "t", , "Please wait", True    '4570
            Do While frm.Chart.IsPartiallyLoaded
                ' while data is still only partially loaded, call again to load more data
                DoEvents
                frm.Chart.GenerateChart eRedo5_RecalcInd
            Loop
            InfBox
        End If
                        
'JM: 02-09-2009 - This is inefficient way to fix issue 2635. Better fix is in subroutine EnoughData.
        'work-around fix for aardvark 2635
        dBarsDate = frm.Chart.Bars(eBARS_DateTime, 0)
        If Int(dBarsDate) > m.dStartDate Then
            strInf = InfBox("The start date has been adjusted to " & DateFormat(dBarsDate) & "." _
                & vbCrLf & "Do you wish to continue?", "I", "Continue|Quit", "Instant Replay" _
                , , , , , , , , eGDAlign_Center)
            If strInf = "C" Then
                m.oGameModeObj.GameDataTime = dBarsDate
                m.dStartDate = dBarsDate
                SaveSettings
            Else
                Unload frm
                Unload Me
                Exit Sub
            End If
        End If
End If
        
        If m.oGameModeObj.InitGame(frm) Then
            frm.hsb.Value = frm.hsb.Max
            ' hide symbol grid if MDI client area not wide enough
            If frmMain.ScaleWidth < 16000 Then
                If DockState(frmSymbolGrid) = eDocked Then
                    DockState(frmSymbolGrid) = eHidden
                    frmMain.tbToolbar.Tools("ID_SymbolGrid").State = ssUnchecked
                End If
            End If
            If m.oGameModeObj.ReplayAll Then
                g.ChartGlobals.nGameInProg = 2
            Else
                g.ChartGlobals.nGameInProg = 1
            End If
            ShowForm frm
            If m.oGameModeObj.GameStrategyID = 0 Or m.oGameModeObj.GameStrategyID = kGameModeSysID Then
                frm.Chart.ShowTrades = 2
            Else
                frm.Chart.ShowTrades = 1
            End If
            'frm.tmrGameMode.Enabled = True
            frm.EnableGameControls
            frm.SetChartTabs
        Else
            InfBox "Replay initialization failed.", "I", , "Instant Replay"
            Unload frm
        End If
    
    End If
    
    Set frm = Nothing
    Unload Me
    
    Exit Sub

ErrSection:
    RaiseError "frmGameModeCfg.cmdStartGame_Click"
    
End Sub

Private Sub Form_Load()
    Me.Icon = Picture16(ToolbarIcon("ID_Replay"), , True)
    
    g.Styler.StyleForm Me
End Sub

Private Sub optAuto_Click()
    m.strReplaySpeed = "auto"
    m.oGameModeObj.GameAutoInterval = True
End Sub

Private Sub optDay_Click()
    m.strReplaySpeed = "day"
    m.oGameModeObj.GameAutoInterval = False
End Sub

Private Sub optMinutes_Click()
    m.strReplaySpeed = "minutes"
    m.oGameModeObj.GameAutoInterval = False
End Sub

Private Sub optRandomSym_Click()
    m.strRandom = "random symbol"
    EnableSelectSym optSelectSymbol.Value
End Sub

Private Sub optSelectSymbol_Click()
    m.strRandom = "user select"
    EnableSelectSym optSelectSymbol.Value
End Sub

Private Sub LoadStrategyCbo()
On Error GoTo ErrSection:

    Dim i&, j&, strSysNum$
    Dim aIdx As cGdArray
    
    cboStrategy.Clear
    
    If cboStrategy.ListCount = 0 Then
        
        If m.tbStrategyInfo.NumRecords = 0 Then
            InitStrategyList
        End If
        
        Set aIdx = m.tbStrategyInfo.CreateSortedIndex(0, eGdSort_Default)
        For i = 0 To m.tbStrategyInfo.NumRecords - 1
            strSysNum = m.tbStrategyInfo(1, aIdx(i))
            If IsDigit(strSysNum, 1) Then
                cboStrategy.AddItem m.tbStrategyInfo(0, aIdx(i))
                cboStrategy.ItemData(i) = Val(strSysNum)
                If "< Game Mode >" = m.tbStrategyInfo(0, aIdx(i)) Then
                    j = i       'aardvark 2652 (default to game mode)
                End If
            Else
                DebugLog "LoadStrategyCbo strSysNum=" & strSysNum & ", Name=" & m.tbStrategyInfo(0, aIdx(i))
            End If
        Next
        
    End If
    
    cboStrategy.ListIndex = j

    Exit Sub
    
ErrSection:
    RaiseError "frmGameModeCfg.LoadStrategyCbo"
    
End Sub

Private Sub LoadBarsPeriodCbo()
On Error GoTo ErrSection:
    
    Dim i&, nDailyIdx&
    Dim bIntraday As Boolean
    
    bIntraday = HasModule("IT") Or HasModule("FT") Or HasModule("ST")
    
    With cboBarsPeriod
        If .ListCount = 0 Then
            If bIntraday Then
                .AddItem "1 minute"
                .AddItem "5 minute"
                .AddItem "10 minute"
                .AddItem "15 minute"
                .AddItem "30 minute"
                .AddItem "60 minute"
            End If
            .AddItem "Daily"
            nDailyIdx = .ListCount - 1
            .AddItem "Weekly"
            .AddItem "Monthly"
            .AddItem "Quarterly"
            .AddItem "Yearly"
        End If

        For i = 0 To .ListCount
            If m.strBarsPeriod = .List(i) Then
                Exit For
            End If
        Next
        
        If i < .ListCount Then
            .ListIndex = i
        Else
            .ListIndex = nDailyIdx
        End If
        
    End With
    
    Exit Sub
    
ErrSection:
    RaiseError "frmGameModeCfg.LoadBarsPeriodCbo"

End Sub

Private Sub SetPlaySymbol()
On Error GoTo ErrSection:

    Dim frm As Form
    
    Set frm = ActiveChart
    
    txtSymbol = ""
    
    If Not frm Is Nothing Then
        If IsFrmChart(frm) Then
            txtSymbol = frm.Chart.Symbol
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmGameModeCfg.SetPlaySymbol"

End Sub

Private Sub InitStrategyList()
On Error GoTo ErrSection:

    Dim rsSystems As Recordset
    Dim strGameMode$
    
    Set rsSystems = mSysNav.LoadStrategiesRecordset

    m.tbStrategyInfo.Clear
    strGameMode = "< Game Mode >" & vbTab & "0" & vbTab & "0"   'aardvark 2927 fix
    m.tbStrategyInfo.AddRecord strGameMode, -1, vbTab
    
    If Not (rsSystems.BOF And rsSystems.EOF) Then
        rsSystems.MoveFirst
        
        Do While Not rsSystems.EOF
            If mSysNav.IncludeStrategiesFromRecordset(rsSystems) Then
                m.tbStrategyInfo.AddRecord rsSystems!SystemName & vbTab & Str(rsSystems!SystemNumber) & vbTab & Str(rsSystems![tblSystems.LibraryID]), -1, vbTab
            End If
            
            rsSystems.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmGameModeCfg.InitStrategyList"

End Sub

Private Sub EnableSelectSym(ByVal bEnable As Boolean)
On Error Resume Next

    txtSymbol.Enabled = bEnable
    cmdSelectSymbol.Enabled = bEnable
    cboFilters.Enabled = Not bEnable

End Sub

Private Function RandomSymbol() As String
On Error GoTo ErrSection:

    Dim nFilterID&, i&, j&, s$
    Dim strSymbol$, strSymGroup$
    Dim nPointerSave&
    
    Dim aAllSyms As cGdArray
    Dim tbGroup As New cGdTable
    
    Dim dEarliest#, dStart#, dDate#
    Dim bIsIntraday As Boolean
    
    tbGroup.CreateField eGDARRAY_Strings, 0, "Symbol"
    tbGroup.CreateField eGDARRAY_Longs, 1, "SymID"
            
    strSymGroup = cboFilters.SelectedItem.Key
    dEarliest = Now
    If InStr(cboBarsPeriod.Text, "minute") Then
        bIsIntraday = True
    Else
        bIsIntraday = False
    End If

    dStart = DateSerial(gdDate.Year, gdDate.Month, gdDate.Day)
        
    nPointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    m.bAbortSearch = False
    
    With g.SymbolPool
        nFilterID = .FieldNumForID(strSymGroup)
        Set aAllSyms = .ArrayTable.FieldArray(nFilterID, True)
        If Not aAllSyms Is Nothing Then
            For i = 0 To aAllSyms.Size - 1
                DoEvents
                If m.bAbortSearch Then
                    Exit For
                ElseIf aAllSyms(i) = 1 Then
                    If bIsIntraday Then
                        dDate = .TickFirstDate(.Symbol(i))
                    Else
                        dDate = .EodFirstDate(.Symbol(i))
                    End If
                    If dDate > 0 And dDate <= dStart Then
                        tbGroup.AddRecord ""
                        tbGroup(0, tbGroup.NumRecords - 1) = .Symbol(i)
                        tbGroup(1, tbGroup.NumRecords - 1) = .SymbolID(i)
                    End If
                    If dDate > 0 And dEarliest > dDate Then dEarliest = dDate
                End If
            Next
        End If
    End With
    
    Screen.MousePointer = nPointerSave
        
    Set aAllSyms = Nothing
    
    If m.bAbortSearch Then
        RandomSymbol = "Try Again"
        Exit Function
    Else
        RandomSymbol = ""
    End If
                
    If dEarliest > dStart Then
        s = "The earliest available data for this symbol group is " & DateFormat(dEarliest) & ". "
        s = s & "Please choose a different start date or symbol group."
        InfBox s, "I", , "Instant Replay (Random Symbol Group)"
        RandomSymbol = "Insufficient Data"
        Exit Function
    ElseIf tbGroup.NumRecords = 0 Then
        InfBox "Instant Replay encountered an internal error and cannot continue." _
               & "You may try again with a different symbol group or start date.", _
               "I", "Ok", "Instant Replay (Internal Error)"
        Exit Function
    End If
    
    For j = 1 To 50
        i = gdRandomNumber(0, tbGroup.NumRecords - 1)
        strSymbol = tbGroup(0, i)
        If EnoughData(strSymbol, False) Then
            Exit For
        Else
            strSymbol = "Insufficient Data"
        End If
    Next
    
    If strSymbol = "Insufficient Data" Then
        InfBox "Unable to find a suitable symbol." & vbCrLf & _
               "Please try a different symbol group or start date.", "I", , "Instant Replay"
    End If
    
    RandomSymbol = strSymbol

    Exit Function
    
ErrSection:
    RaiseError "frmGameModeCfg.RandomSymbol"

End Function

Private Sub SaveSettings()
On Error GoTo ErrSection:

    If optRandomSym.Value = True Then
        m.strRandom = cboFilters.SelectedItem.Key
    Else
        m.strRandom = "user select"
    End If
    'saved info to INI
    SetIniFileProperty "Random", m.strRandom, "Game Mode", g.strIniFile
    SetIniFileProperty "BarsPeriod", m.strBarsPeriod, "Game Mode", g.strIniFile
    SetIniFileProperty "ReplaySpeed", m.strReplaySpeed, "Game Mode", g.strIniFile
    SetIniFileProperty "ReplayMinutes", m.nReplayMinutes, "Game Mode", g.strIniFile
    SetIniFileProperty "ReplayStart", m.dStartDate, "Game Mode", g.strIniFile
    SetIniFileProperty "ReplayAll", m.oGameModeObj.ReplayAll, "Game Mode", g.strIniFile

    Exit Sub
    
ErrSection:
    RaiseError "frmGameModeCfg.SaveSettings"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load up the filters combo box with the symbol groups
''              This sub is virtually an exact duplicate of the one in frmSymbolSelector
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo(Optional ByVal bShowFilters As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim strID As String                 ' Symbol pool ID for the field
    Dim strType As String               ' Type of thing (i.e. Filter, Criteria, etc)
    Dim strPicture As String            ' Picture to use in the combo box
    Dim bRandom As Boolean
    Dim iSortStart As Long              ' Where to start the sort
    Dim strItem As String               ' Item to add to the combo box
    Dim aItems As New cGdArray          ' Items to add to the combo box
    Dim obj As Object                   ' Symbol Pool Object
    Dim bScans As Boolean               ' Are we doing scans?
       
    bScans = ScansEnabled
        
    If cboFilters.ComboItems.Count > 0 Then
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
                        If iSortStart = 0 And lIndex >= g.SymbolPool.OtherFieldsStart Then
                            iSortStart = aItems.Size
                        End If
                        
                        aItems.Add .ArrayTable.FieldName(lIndex) & vbTab _
                                & strID & vbTab & strPicture
                   End If
                End If
            End If
        Next
    End With
    If iSortStart > 0 Then
        aItems.Sort eGdSort_IgnoreCase, iSortStart
    End If

    Dim str1$, Str2$, str3$

    For lIndex = 0 To aItems.Size - 1
        strItem = aItems(lIndex)
        str1 = Parse(strItem, vbTab, 1)
        Str2 = Parse(strItem, vbTab, 2)
        str3 = Parse(strItem, vbTab, 3)
        If Len(str1) > 0 And Len(Str2) > 0 And Len(str3) > 0 Then
            cboFilters.ComboItems.Add , Str2, str1, str3
        End If
        If m.strRandom = Str2 Then bRandom = True
    Next

    ' Set combo to first item (All Symbols), then try to set
    ' to a better default (but don't error if not exists)
    On Error Resume Next
    cboFilters.ComboItems(1).Selected = True
    If bRandom Then
        cboFilters.ComboItems(m.strRandom).Selected = True
    ElseIf HasModule("F") Then
        cboFilters.ComboItems("GRP:CONT067.GRP").Selected = True
    Else
        cboFilters.ComboItems("GRP:SP500.GRP").Selected = True
    End If

    cboFilters.Refresh
    
    Exit Sub

ErrSection:
    RaiseError "frmGameModeCfg.LoadCombo"
    
End Sub

Private Function EnoughData(ByVal strSymbol, ByVal bInfMsg As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim dFirst#, dStart#, iCount&
    Dim bEnough As Boolean
    Dim strMsg$
    Dim Bars As New cGdBars
    
    bEnough = True
    dStart = DateSerial(gdDate.Year, gdDate.Month, gdDate.Day)
    
    With g.SymbolPool
        Bars.Prop(eBARS_PeriodicityStr) = cboBarsPeriod.Text
        If Bars.IsIntraday Then
            ' and go forward until there are at least 10 ticks in a day
            ' (to avoid issues with sparsely traded contracts)     -aardvark 5943
            ' TLB/JM 3/6/2015: doing this check for stocks requires Full Ticks, so only do
            ' this check for Futures (which is probably all it was intended for originally)
            dFirst = .TickFirstDate(strSymbol)
            If dFirst > 0 And Bars.SecurityType = "F" Then
                Do While True
                    DM_GetBars Bars, strSymbol, "EachTick", dFirst, dFirst
                    
                    If Bars.Size >= 10 Or iCount > 364 Then Exit Do
                    
                    iCount = iCount + 1
                    dFirst = dFirst + 1
                Loop
            End If
        
            If iCount > 364 Then
                DebugLog "Instant Replay enough data check failed after 365 attemps. Symbol=" & Bars.Prop(eBARS_Symbol) & " Periodicity=" & Bars.Prop(eBARS_PeriodicityStr)
                dFirst = 0
            End If
        Else
            ' TLB: at minimum, should start at least 7 days after the very first data (2635)
            dFirst = .EodFirstDate(strSymbol)
            If dFirst > 0 Then dFirst = dFirst + 7
        End If
    End With

    If dFirst <= 0 Then
        bEnough = False
        If bInfMsg Then
            strMsg = "There is not enough data for " & strSymbol & "."
            strMsg = strMsg & vbCrLf & "Please select a different symbol."
            InfBox strMsg, "I", , "Instant Replay"
        End If
    ElseIf dFirst > dStart Then
        bEnough = False
        If bInfMsg Then
            strMsg = "The earliest available data for " & strSymbol & " is " & DateFormat(dFirst) & ". "
            strMsg = strMsg & "Do you wish to use this date?"
            strMsg = InfBox(strMsg, "I", "Yes|No", "Instant Replay")
            If Left(strMsg, 1) = "Y" Then
                bEnough = True
                gdDate.Value = dFirst + 1
            End If
        End If
    ElseIf Int(dStart) = Int(dFirst) Then
        'game object backs game start time up one session
        gdDate.Value = dFirst + 1
    End If

    EnoughData = bEnough

    Exit Function
    
ErrSection:
    RaiseError "frmGameModeCfg.EnoughData"

End Function

Private Function PromptStartNew() As Boolean
On Error GoTo ErrSection:

    Dim s$, i&

    Dim bNew As Boolean
    
    s = "Would you like to quit the current Instant Replay and start a new one?"
    s = InfBox(s, "?", "+Start New|-Cancel", "Instant Replay")
    
    If s = "S" Then
        'find form currently in instant replay mode
        For i = 0 To Forms.Count - 1
            If IsFrmChart(Forms(i)) Then
                If Forms(i).IsInGameMode Then
                    If Forms(i).GameReplayMode <> eGDReplayMode_Sync Then
                        Forms(i).ClearReplaySync
                        DoEvents
                        Forms(i).GameMode.ShowResultsFlag = False
                        DoEvents
                        Unload Forms(i)
                        DoEvents
                        bNew = True
                        DoEvents
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    PromptStartNew = bNew
    
    Exit Function

ErrSection:
    RaiseError "frmGameModeCfg.PromptStartNew"

End Function


Private Sub FixFrames()

    On Error Resume Next
    If optStreaming.Value Then
        fraStreaming.Visible = True
        fraInstant.Visible = False
        cmdResults.Visible = False
    Else
        fraStreaming.Visible = False
        fraInstant.Visible = True
        cmdResults.Visible = True
    End If
    fraButtons.ZOrder

End Sub

Private Sub LoadStreamFiles()
On Error GoTo ErrSection:

    Dim i&, d#, s$
    Dim aFiles As New cGdArray

    lstSessions.Clear
    aFiles.GetMatchingFiles App.Path & "\RTS\*.rts", False
    aFiles.Sort eGdSort_IgnoreCase Or eGdSort_Descending
    For i = 0 To aFiles.Size - 1
        s = Parse(aFiles(i), ".", 1)
        If Len(s) >= 8 Then
            d = Val(s)
            d = DateSerial(Int(d / 10000), Int(d / 100) Mod 100, d Mod 100)
            s = DateFormat(d) & Format(d, "  ddd")
            lstSessions.AddItem s
            lstSessions.ItemData(lstSessions.ListCount - 1) = d
        End If
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmGameModeCfg.LoadStreamFiles"
End Sub

Private Sub SetStreamDate(Optional ByVal dDate As Double = 0)
On Error GoTo ErrSection:

    Dim i&, n&, d#
    Dim aSymbols As New cGdArray
    Static bInProgress As Boolean

    If bInProgress Then Exit Sub
    bInProgress = True
    
    With lstSessions
        ' if valid date is not passed in, then get from List
        If dDate <= 0 Then
            If .ListIndex >= 0 Then
                dDate = .ItemData(.ListIndex)
            Else
                dDate = LastDailyDownload
            End If
        End If
        dDate = Int(dDate)
        ''If dDate > LastDailyDownload Then dDate = LastDailyDownload
        Do While Not IsWeekday(dDate)
            dDate = dDate - 1
        Loop
        
        ' set date control
        If dtDate.Value <> dDate Then
            dtDate.Value = dDate
        End If
    
        ' select item in list (if a match)
        If .ListIndex >= 0 Then
            If .ItemData(.ListIndex) <> dDate Then
                .ListIndex = -1
            End If
        End If
        If .ListIndex < 0 Then
            For i = 0 To .ListCount - 1
                If .ItemData(i) = dDate Then
                    .ListIndex = i
                    Exit For
                End If
            Next
        End If
    End With
    
    ' get available symbols for this date
    lstSymbols.Clear
    If dDate > 0 Then
        aSymbols.FromFile App.Path & "\Info\Symbols.RTS"
        aSymbols.Sort eGdSort_IgnoreCase
        d = Val(Parse(aSymbols(0), vbTab, 1))
        If d > 0 Then m.nEarliestDate = JulFromLong(d)
        For i = aSymbols.Size - 1 To 0 Step -1
            d = Val(Parse(aSymbols(i), vbTab, 1))
            d = JulFromLong(d)
            If d <= dDate Then
                aSymbols.SplitFields Parse(aSymbols(i), vbTab, 2), ","
                'aSymbols.Sort eGdSort_IgnoreCase Or eGdSort_DeleteDuplicates
                For n = 0 To aSymbols.Size - 1
                    lstSymbols.AddItem aSymbols(n)
                Next
                Exit For
            End If
        Next
    End If

ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "frmGameModeCfg.SetStreamDate"
End Sub

Private Function CanDoStreamReplay(ByVal bShowMsg As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim strMsg$
   
    ' if can do streaming replay (T_SRP or has bought Gold and has streaming) ...
    If HasModule("T_SRP") Or (HasModule("GOLD") And HasModule("RTG")) Then
        ' but not while realtime streaming is active
        If g.RealTime.Active And g.nReplaySession = 0 Then
            strMsg = "You must first stop the realtime streaming before trying to start a replay session."
        End If
    Else
        strMsg = "Your account is not currently enabled for streaming replay.  Please contact Genesis sales for more information."
    End If
    
    If Len(strMsg) = 0 Then
        CanDoStreamReplay = True
    ElseIf bShowMsg Then
        InfBox strMsg, "e", , "Invalid"
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmGameModeCfg.CanDoStreamReplay"
End Function

