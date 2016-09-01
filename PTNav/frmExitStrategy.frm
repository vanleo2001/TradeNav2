VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmExitStrategy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraOptions 
      Height          =   1275
      Left            =   180
      TabIndex        =   0
      Top             =   6840
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
      Caption         =   "frmExitStrategy.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExitStrategy.frx":0030
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExitStrategy.frx":0050
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkMarketIfWrongSide 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   900
         Width           =   5415
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
         Caption         =   "frmExitStrategy.frx":006C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExitStrategy.frx":0118
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":0138
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkCancelOpposite 
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   5355
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
         Caption         =   "frmExitStrategy.frx":0154
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExitStrategy.frx":01F6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":0216
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optQuarterTicks 
         Height          =   220
         Left            =   4320
         TabIndex        =   18
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
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
         Caption         =   "frmExitStrategy.frx":0232
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExitStrategy.frx":026C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":028C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optHalfTicks 
         Height          =   220
         Left            =   3060
         TabIndex        =   19
         Top             =   300
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "frmExitStrategy.frx":02A8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExitStrategy.frx":02DC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":02FC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optFullTicks 
         Height          =   220
         Left            =   1860
         TabIndex        =   24
         Top             =   300
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "frmExitStrategy.frx":0318
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmExitStrategy.frx":034C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":036C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblShowValues 
         Height          =   195
         Left            =   240
         Top             =   300
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
         Caption         =   "frmExitStrategy.frx":0388
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExitStrategy.frx":03D0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":03F0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtDescription 
      Height          =   555
      Left            =   1140
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmExitStrategy.frx":040C
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
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmExitStrategy.frx":042C
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExitStrategy.frx":044C
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3375
      Left            =   6060
      TabIndex        =   25
      Top             =   60
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
      Caption         =   "frmExitStrategy.frx":0468
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExitStrategy.frx":0494
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExitStrategy.frx":04B4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraEditorButtons 
         Height          =   2115
         Left            =   0
         TabIndex        =   28
         Top             =   1260
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
         Caption         =   "frmExitStrategy.frx":04D0
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmExitStrategy.frx":04FC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":051C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdRename 
            Height          =   495
            Left            =   0
            TabIndex        =   29
            Top             =   1080
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
            Caption         =   "frmExitStrategy.frx":0538
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":0566
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0586
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSaveAs 
            Height          =   495
            Left            =   0
            TabIndex        =   34
            Top             =   540
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
            Caption         =   "frmExitStrategy.frx":05A2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":05D2
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":05F2
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSave 
            Height          =   495
            Left            =   0
            TabIndex        =   35
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
            Caption         =   "frmExitStrategy.frx":060E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":0638
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0658
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdExit 
            Height          =   495
            Left            =   0
            TabIndex        =   38
            Top             =   1620
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
            Caption         =   "frmExitStrategy.frx":0674
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":069E
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":06BE
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraOKCancel 
         Height          =   1035
         Left            =   0
         TabIndex        =   39
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
         Caption         =   "frmExitStrategy.frx":06DA
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmExitStrategy.frx":0706
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":0726
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
            Height          =   495
            Left            =   0
            TabIndex        =   42
            Top             =   540
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
            Caption         =   "frmExitStrategy.frx":0742
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":076E
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":078E
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdOK 
            Height          =   495
            Left            =   0
            TabIndex        =   45
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
            Caption         =   "frmExitStrategy.frx":07AA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":07CE
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":07EE
            RightToLeft     =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraProfitTargets 
      Height          =   2895
      Left            =   180
      TabIndex        =   2
      Top             =   840
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
      Caption         =   "frmExitStrategy.frx":080A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExitStrategy.frx":084A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExitStrategy.frx":086A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraTsProfit 
         Height          =   2295
         Left            =   900
         TabIndex        =   40
         Top             =   1020
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
         Caption         =   "frmExitStrategy.frx":0886
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmExitStrategy.frx":08B2
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":08D2
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdClearShortProfit 
            Height          =   255
            Left            =   3180
            TabIndex        =   54
            Top             =   1380
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":08EE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":0918
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0938
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdClearLongProfit 
            Height          =   255
            Left            =   3180
            TabIndex        =   49
            Top             =   420
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":0954
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":097E
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":099E
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtTargetQtyTs 
            Height          =   285
            Left            =   2460
            TabIndex        =   43
            Top             =   45
            Width           =   495
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frmExitStrategy.frx":09BA
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
            Tip             =   "frmExitStrategy.frx":09DC
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":09FC
         End
         Begin HexUniControls.ctlUniCheckXP chkEntirePositionTs 
            Height          =   195
            Left            =   0
            TabIndex        =   41
            Top             =   90
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
            Caption         =   "frmExitStrategy.frx":0A18
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":0A60
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0A80
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdCopyShortProfit 
            Height          =   255
            Left            =   2340
            TabIndex        =   53
            Top             =   1380
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":0A9C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":0AC4
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0AE4
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdCopyLongProfit 
            Height          =   255
            Left            =   2340
            TabIndex        =   48
            Top             =   420
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":0B00
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":0B28
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0B48
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdEditLongProfit 
            Height          =   255
            Left            =   1500
            TabIndex        =   47
            Top             =   420
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":0B64
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":0B8C
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0BAC
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdEditShortProfit 
            Height          =   255
            Left            =   1500
            TabIndex        =   52
            Top             =   1380
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":0BC8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":0BF0
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0C10
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRichTextBoxXP rtbTsLongProfit 
            Height          =   615
            Left            =   0
            TabIndex        =   50
            Top             =   720
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   1085
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmExitStrategy.frx":0C2C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   -1
            MultiLine       =   -1  'True
            Alignment       =   0
            ScrollBars      =   3
            PasswordChar    =   ""
            TrapTab         =   0   'False
            RaiseChangeEvent=   -1  'True
            RaiseUpdateEvent=   0   'False
            RaiseSelChangeEvent=   -1  'True
            Tip             =   "frmExitStrategy.frx":0C4C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0C6C
            ViewMode        =   0
            TextModeText    =   2
            TextModeUndoLevel=   8
            TextModeCodePage=   32
            AutoURLDetect   =   0   'False
            FileName        =   ""
            VerticalLayout  =   0   'False
            OnlyNumbers     =   0   'False
            NoIME           =   0   'False
            SelfIME         =   0   'False
            LanguageOptions =   150
            RaiseRequestResizeEvent=   0   'False
            RaiseMsgFilterEvent=   0   'False
            SubClassPaintMessage=   0   'False
            TabSize         =   4
            TypographyOptions=   0
            BlockAutoCopy   =   0   'False
            BlockAutoCut    =   0   'False
            BlockAutoPaste  =   0   'False
            BlockAutoUndo   =   0   'False
         End
         Begin HexUniControls.ctlUniRichTextBoxXP rtbTsShortProfit 
            Height          =   615
            Left            =   0
            TabIndex        =   55
            Top             =   1680
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   1085
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmExitStrategy.frx":0C88
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   -1
            MultiLine       =   -1  'True
            Alignment       =   0
            ScrollBars      =   3
            PasswordChar    =   ""
            TrapTab         =   0   'False
            RaiseChangeEvent=   -1  'True
            RaiseUpdateEvent=   0   'False
            RaiseSelChangeEvent=   -1  'True
            Tip             =   "frmExitStrategy.frx":0CA8
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0CC8
            ViewMode        =   0
            TextModeText    =   2
            TextModeUndoLevel=   8
            TextModeCodePage=   32
            AutoURLDetect   =   0   'False
            FileName        =   ""
            VerticalLayout  =   0   'False
            OnlyNumbers     =   0   'False
            NoIME           =   0   'False
            SelfIME         =   0   'False
            LanguageOptions =   150
            RaiseRequestResizeEvent=   0   'False
            RaiseMsgFilterEvent=   0   'False
            SubClassPaintMessage=   0   'False
            TabSize         =   4
            TypographyOptions=   0
            BlockAutoCopy   =   0   'False
            BlockAutoCut    =   0   'False
            BlockAutoPaste  =   0   'False
            BlockAutoUndo   =   0   'False
         End
         Begin gdOCX.gdScrollBar sbTargetQtyTs 
            Height          =   360
            Left            =   2940
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   0
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
         Begin HexUniControls.ctlUniLabelXP lblTargetQtyTs 
            Height          =   195
            Left            =   1740
            Top             =   90
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
            Caption         =   "frmExitStrategy.frx":0CE4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmExitStrategy.frx":0D18
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0D38
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblTargetTsLots 
            Height          =   195
            Left            =   3240
            Top             =   90
            Width           =   375
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
            Caption         =   "frmExitStrategy.frx":0D54
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmExitStrategy.frx":0D80
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0DA0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblTsLongProfit 
            Height          =   195
            Left            =   0
            Top             =   450
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
            Caption         =   "frmExitStrategy.frx":0DBC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmExitStrategy.frx":0E02
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0E22
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblTsShortProfit 
            Height          =   195
            Left            =   0
            Top             =   1410
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
            Caption         =   "frmExitStrategy.frx":0E3E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmExitStrategy.frx":0E86
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0EA6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraProfitTargetOptions 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
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
         Caption         =   "frmExitStrategy.frx":0EC2
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmExitStrategy.frx":0EEE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":0F0E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optStandardProfit 
            Height          =   255
            Left            =   0
            TabIndex        =   4
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
            Caption         =   "frmExitStrategy.frx":0F2A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmExitStrategy.frx":0F5A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0F7A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optTsProfit 
            Height          =   255
            Left            =   1080
            TabIndex        =   5
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
            Caption         =   "frmExitStrategy.frx":0F96
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":0FCA
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":0FEA
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraNonTsProfit 
         Height          =   2175
         Left            =   180
         TabIndex        =   6
         Top             =   540
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
         Caption         =   "frmExitStrategy.frx":1006
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmExitStrategy.frx":1026
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":1046
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraLots 
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   5415
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
            Caption         =   "frmExitStrategy.frx":1062
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmExitStrategy.frx":108E
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":10AE
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkEntirePosition 
               Height          =   220
               Left            =   3240
               TabIndex        =   10
               Top             =   0
               Width           =   2115
               _ExtentX        =   3731
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
               Caption         =   "frmExitStrategy.frx":10CA
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmExitStrategy.frx":1112
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":1132
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optEqualLots 
               Height          =   220
               Left            =   60
               TabIndex        =   8
               Top             =   0
               Width           =   1215
               _ExtentX        =   2143
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
               Caption         =   "frmExitStrategy.frx":114E
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmExitStrategy.frx":1182
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":11A2
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optSpecifyLots 
               Height          =   220
               Left            =   1320
               TabIndex        =   9
               Top             =   0
               Width           =   1215
               _ExtentX        =   2143
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
               Caption         =   "frmExitStrategy.frx":11BE
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmExitStrategy.frx":11F6
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":1216
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraProfitTarget3 
            Height          =   375
            Left            =   60
            TabIndex        =   11
            Top             =   480
            Width           =   5355
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
            Caption         =   "frmExitStrategy.frx":1232
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmExitStrategy.frx":125E
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":127E
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraProfitTarget3Qty 
               Height          =   375
               Left            =   3180
               TabIndex        =   46
               Top             =   0
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
               Caption         =   "frmExitStrategy.frx":129A
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmExitStrategy.frx":12D6
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":12F6
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtTargetQty3 
                  Height          =   285
                  Left            =   720
                  TabIndex        =   17
                  Top             =   38
                  Width           =   495
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   0   'False
                  Locked          =   0   'False
                  Text            =   "frmExitStrategy.frx":1312
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
                  Tip             =   "frmExitStrategy.frx":1334
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":1354
               End
               Begin gdOCX.gdScrollBar sbTargetQty3 
                  Height          =   360
                  Left            =   1200
                  TabIndex        =   51
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   635
               End
               Begin HexUniControls.ctlUniLabelXP lblTarget3Lots 
                  Height          =   195
                  Left            =   1500
                  Top             =   120
                  Width           =   375
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
                  Caption         =   "frmExitStrategy.frx":1370
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":139C
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":13BC
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblTargetQty3 
                  Height          =   195
                  Left            =   0
                  Top             =   120
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
                  Caption         =   "frmExitStrategy.frx":13D8
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":140C
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":142C
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniTextBoxXP txtTargetTicks3 
               Height          =   285
               Left            =   1380
               TabIndex        =   13
               Top             =   38
               Width           =   915
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Locked          =   0   'False
               Text            =   "frmExitStrategy.frx":1448
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
               Tip             =   "frmExitStrategy.frx":1470
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":1490
            End
            Begin HexUniControls.ctlUniCheckXP chkProfitTarget3 
               Height          =   220
               Left            =   0
               TabIndex        =   12
               Top             =   120
               Width           =   1395
               _ExtentX        =   2461
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
               Caption         =   "frmExitStrategy.frx":14AC
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmExitStrategy.frx":14EE
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":150E
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdScrollBar sbTargetTicks3 
               Height          =   360
               Left            =   2280
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   0
               Width           =   210
               _ExtentX        =   370
               _ExtentY        =   635
            End
            Begin HexUniControls.ctlUniLabelXP lblProfitTarget3Pos 
               Height          =   195
               Left            =   4560
               Top             =   90
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
               Caption         =   "frmExitStrategy.frx":152A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmExitStrategy.frx":156A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":158A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTicks3 
               Height          =   195
               Left            =   2520
               Top             =   120
               Width           =   375
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
               Caption         =   "frmExitStrategy.frx":15A6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmExitStrategy.frx":15D0
               Style           =   0
               Enabled         =   0   'False
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":15F0
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraProfitTarget2 
            Height          =   375
            Left            =   60
            TabIndex        =   20
            Top             =   900
            Width           =   5355
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
            Caption         =   "frmExitStrategy.frx":160C
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmExitStrategy.frx":1638
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":1658
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraProfitTarget2Qty 
               Height          =   375
               Left            =   3180
               TabIndex        =   64
               Top             =   0
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
               Caption         =   "frmExitStrategy.frx":1674
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmExitStrategy.frx":16A0
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":16C0
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtTargetQty2 
                  Height          =   285
                  Left            =   720
                  TabIndex        =   26
                  Top             =   38
                  Width           =   495
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   0   'False
                  Locked          =   0   'False
                  Text            =   "frmExitStrategy.frx":16DC
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
                  Tip             =   "frmExitStrategy.frx":16FE
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":171E
               End
               Begin gdOCX.gdScrollBar sbTargetQty2 
                  Height          =   360
                  Left            =   1200
                  TabIndex        =   27
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   635
               End
               Begin HexUniControls.ctlUniLabelXP lblTarget2Lots 
                  Height          =   195
                  Left            =   1500
                  Top             =   120
                  Width           =   375
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
                  Caption         =   "frmExitStrategy.frx":173A
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":1766
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":1786
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblTargetQty2 
                  Height          =   195
                  Left            =   0
                  Top             =   120
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
                  Caption         =   "frmExitStrategy.frx":17A2
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":17D6
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":17F6
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniTextBoxXP txtTargetTicks2 
               Height          =   285
               Left            =   1380
               TabIndex        =   22
               Top             =   38
               Width           =   915
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Locked          =   0   'False
               Text            =   "frmExitStrategy.frx":1812
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
               Tip             =   "frmExitStrategy.frx":183A
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":185A
            End
            Begin HexUniControls.ctlUniCheckXP chkProfitTarget2 
               Height          =   220
               Left            =   0
               TabIndex        =   21
               Top             =   120
               Width           =   1395
               _ExtentX        =   2461
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
               Caption         =   "frmExitStrategy.frx":1876
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmExitStrategy.frx":18B8
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":18D8
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdScrollBar sbTargetTicks2 
               Height          =   360
               Left            =   2280
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   0
               Width           =   210
               _ExtentX        =   370
               _ExtentY        =   635
            End
            Begin HexUniControls.ctlUniLabelXP lblProfitTarget2Pos 
               Height          =   195
               Left            =   4560
               Top             =   90
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
               Caption         =   "frmExitStrategy.frx":18F4
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmExitStrategy.frx":1934
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":1954
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTicks2 
               Height          =   195
               Left            =   2520
               Top             =   120
               Width           =   435
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
               Caption         =   "frmExitStrategy.frx":1970
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmExitStrategy.frx":199A
               Style           =   0
               Enabled         =   0   'False
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":19BA
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraProfitTarget1 
            Height          =   375
            Left            =   60
            TabIndex        =   30
            Top             =   1320
            Width           =   5355
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
            Caption         =   "frmExitStrategy.frx":19D6
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmExitStrategy.frx":1A02
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":1A22
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraProfitTarget1Qty 
               Height          =   375
               Left            =   3180
               TabIndex        =   67
               Top             =   0
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
               Caption         =   "frmExitStrategy.frx":1A3E
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmExitStrategy.frx":1A6A
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":1A8A
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtTargetQty1 
                  Height          =   285
                  Left            =   720
                  TabIndex        =   36
                  Top             =   38
                  Width           =   495
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   0   'False
                  Locked          =   0   'False
                  Text            =   "frmExitStrategy.frx":1AA6
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
                  Tip             =   "frmExitStrategy.frx":1AC8
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":1AE8
               End
               Begin gdOCX.gdScrollBar sbTargetQty1 
                  Height          =   360
                  Left            =   1200
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   635
               End
               Begin HexUniControls.ctlUniLabelXP lblTarget1Lots 
                  Height          =   195
                  Left            =   1500
                  Top             =   120
                  Width           =   375
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
                  Caption         =   "frmExitStrategy.frx":1B04
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":1B30
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":1B50
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblTargetQty1 
                  Height          =   195
                  Left            =   0
                  Top             =   120
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
                  Caption         =   "frmExitStrategy.frx":1B6C
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":1BA2
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":1BC2
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniTextBoxXP txtTargetTicks1 
               Height          =   285
               Left            =   1380
               TabIndex        =   32
               Top             =   60
               Width           =   915
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Locked          =   0   'False
               Text            =   "frmExitStrategy.frx":1BDE
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
               Tip             =   "frmExitStrategy.frx":1C06
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":1C26
            End
            Begin HexUniControls.ctlUniCheckXP chkProfitTarget1 
               Height          =   220
               Left            =   0
               TabIndex        =   31
               Top             =   120
               Width           =   1395
               _ExtentX        =   2461
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
               Caption         =   "frmExitStrategy.frx":1C42
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmExitStrategy.frx":1C84
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":1CA4
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdScrollBar sbTargetTicks1 
               Height          =   360
               Left            =   2280
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   0
               Width           =   210
               _ExtentX        =   370
               _ExtentY        =   635
            End
            Begin HexUniControls.ctlUniLabelXP lblProfitTarget1Pos 
               Height          =   195
               Left            =   4560
               Top             =   90
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
               Caption         =   "frmExitStrategy.frx":1CC0
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmExitStrategy.frx":1D00
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":1D20
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTicks1 
               Height          =   195
               Left            =   2520
               Top             =   120
               Width           =   375
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
               Caption         =   "frmExitStrategy.frx":1D3C
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmExitStrategy.frx":1D66
               Style           =   0
               Enabled         =   0   'False
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":1D86
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraStopLoss 
      Height          =   2895
      Left            =   180
      TabIndex        =   56
      Top             =   3840
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
      Caption         =   "frmExitStrategy.frx":1DA2
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExitStrategy.frx":1DD4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExitStrategy.frx":1DF4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraTsStop 
         Height          =   2175
         Left            =   600
         TabIndex        =   89
         Top             =   840
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
         Caption         =   "frmExitStrategy.frx":1E10
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmExitStrategy.frx":1E3C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":1E5C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdClearShortStop 
            Height          =   255
            Left            =   3240
            TabIndex        =   72
            Top             =   1140
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":1E78
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":1EA2
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":1EC2
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdClearLongStop 
            Height          =   255
            Left            =   3240
            TabIndex        =   75
            Top             =   60
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":1EDE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":1F08
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":1F28
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdCopyShortStop 
            Height          =   255
            Left            =   2400
            TabIndex        =   78
            Top             =   1140
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":1F44
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":1F6C
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":1F8C
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdCopyLongStop 
            Height          =   255
            Left            =   2400
            TabIndex        =   80
            Top             =   60
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":1FA8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":1FD0
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":1FF0
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdEditLongStop 
            Height          =   255
            Left            =   1560
            TabIndex        =   91
            Top             =   60
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":200C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":2034
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":2054
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdEditShortStop 
            Height          =   255
            Left            =   1560
            TabIndex        =   83
            Top             =   1140
            Width           =   795
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
            Caption         =   "frmExitStrategy.frx":2070
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":2098
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":20B8
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRichTextBoxXP rtbTsLongStop 
            Height          =   615
            Left            =   60
            TabIndex        =   88
            Top             =   360
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   1085
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmExitStrategy.frx":20D4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   -1
            MultiLine       =   -1  'True
            Alignment       =   0
            ScrollBars      =   3
            PasswordChar    =   ""
            TrapTab         =   0   'False
            RaiseChangeEvent=   -1  'True
            RaiseUpdateEvent=   0   'False
            RaiseSelChangeEvent=   -1  'True
            Tip             =   "frmExitStrategy.frx":20F4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":2114
            ViewMode        =   0
            TextModeText    =   2
            TextModeUndoLevel=   8
            TextModeCodePage=   32
            AutoURLDetect   =   0   'False
            FileName        =   ""
            VerticalLayout  =   0   'False
            OnlyNumbers     =   0   'False
            NoIME           =   0   'False
            SelfIME         =   0   'False
            LanguageOptions =   150
            RaiseRequestResizeEvent=   0   'False
            RaiseMsgFilterEvent=   0   'False
            SubClassPaintMessage=   0   'False
            TabSize         =   4
            TypographyOptions=   0
            BlockAutoCopy   =   0   'False
            BlockAutoCut    =   0   'False
            BlockAutoPaste  =   0   'False
            BlockAutoUndo   =   0   'False
         End
         Begin HexUniControls.ctlUniRichTextBoxXP rtbTsShortStop 
            Height          =   615
            Left            =   60
            TabIndex        =   90
            Top             =   1440
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   1085
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmExitStrategy.frx":2130
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   -1
            MultiLine       =   -1  'True
            Alignment       =   0
            ScrollBars      =   3
            PasswordChar    =   ""
            TrapTab         =   0   'False
            RaiseChangeEvent=   -1  'True
            RaiseUpdateEvent=   0   'False
            RaiseSelChangeEvent=   -1  'True
            Tip             =   "frmExitStrategy.frx":2150
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":2170
            ViewMode        =   0
            TextModeText    =   2
            TextModeUndoLevel=   8
            TextModeCodePage=   32
            AutoURLDetect   =   0   'False
            FileName        =   ""
            VerticalLayout  =   0   'False
            OnlyNumbers     =   0   'False
            NoIME           =   0   'False
            SelfIME         =   0   'False
            LanguageOptions =   150
            RaiseRequestResizeEvent=   0   'False
            RaiseMsgFilterEvent=   0   'False
            SubClassPaintMessage=   0   'False
            TabSize         =   4
            TypographyOptions=   0
            BlockAutoCopy   =   0   'False
            BlockAutoCut    =   0   'False
            BlockAutoPaste  =   0   'False
            BlockAutoUndo   =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblTsLongStop 
            Height          =   195
            Left            =   60
            Top             =   90
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
            Caption         =   "frmExitStrategy.frx":218C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmExitStrategy.frx":21CA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":21EA
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblTsShortStop 
            Height          =   195
            Left            =   60
            Top             =   1170
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
            Caption         =   "frmExitStrategy.frx":2206
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmExitStrategy.frx":2246
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":2266
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraStopLossOptions 
         Height          =   255
         Left            =   180
         TabIndex        =   57
         Top             =   300
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
         Caption         =   "frmExitStrategy.frx":2282
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmExitStrategy.frx":22AE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":22CE
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optTsStop 
            Height          =   255
            Left            =   3960
            TabIndex        =   62
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
            Caption         =   "frmExitStrategy.frx":22EA
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":231E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":233E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optBreakEven 
            Height          =   255
            Left            =   2700
            TabIndex        =   61
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
            Caption         =   "frmExitStrategy.frx":235A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":238E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":23AE
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optTrailStop 
            Height          =   255
            Left            =   1740
            TabIndex        =   60
            Top             =   0
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
            Caption         =   "frmExitStrategy.frx":23CA
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":23FA
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":241A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optFixedStop 
            Height          =   255
            Left            =   900
            TabIndex        =   59
            Top             =   0
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
            Caption         =   "frmExitStrategy.frx":2436
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmExitStrategy.frx":2460
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":2480
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optNoStop 
            Height          =   255
            Left            =   0
            TabIndex        =   58
            Top             =   0
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
            Caption         =   "frmExitStrategy.frx":249C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmExitStrategy.frx":24C4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":24E4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraNonTsStop 
         Height          =   2175
         Left            =   180
         TabIndex        =   63
         Top             =   600
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
         Caption         =   "frmExitStrategy.frx":2500
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmExitStrategy.frx":252C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExitStrategy.frx":254C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraWithLimit 
            Height          =   435
            Left            =   420
            TabIndex        =   68
            Top             =   390
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
            Caption         =   "frmExitStrategy.frx":2568
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmExitStrategy.frx":2594
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":25B4
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtWithLimitTicks 
               Height          =   285
               Left            =   1620
               TabIndex        =   70
               Top             =   75
               Width           =   915
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Locked          =   0   'False
               Text            =   "frmExitStrategy.frx":25D0
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
               Tip             =   "frmExitStrategy.frx":25F8
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":2618
            End
            Begin HexUniControls.ctlUniCheckXP chkWithLimit 
               Height          =   195
               Left            =   0
               TabIndex        =   69
               Top             =   120
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
               Caption         =   "frmExitStrategy.frx":2634
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmExitStrategy.frx":267A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":269A
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdScrollBar sbWithLimitTicks 
               Height          =   360
               Left            =   2520
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   30
               Width           =   210
               _ExtentX        =   370
               _ExtentY        =   635
            End
            Begin HexUniControls.ctlUniLabelXP lblWithLimitTicks 
               Height          =   195
               Left            =   2820
               Top             =   120
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
               Caption         =   "frmExitStrategy.frx":26B6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmExitStrategy.frx":270A
               Style           =   0
               Enabled         =   0   'False
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":272A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniTextBoxXP txtStopTicks 
            Height          =   285
            Left            =   2100
            TabIndex        =   65
            Top             =   45
            Width           =   915
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frmExitStrategy.frx":2746
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
            Tip             =   "frmExitStrategy.frx":276E
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":278E
         End
         Begin HexUniControls.ctlUniFrameWL fraBreakEven 
            Height          =   1215
            Left            =   60
            TabIndex        =   73
            Top             =   930
            Width           =   5415
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
            Caption         =   "frmExitStrategy.frx":27AA
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmExitStrategy.frx":27E0
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":2800
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraTrailBE 
               Height          =   435
               Left            =   360
               TabIndex        =   84
               Top             =   780
               Width           =   5295
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
               Caption         =   "frmExitStrategy.frx":281C
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmExitStrategy.frx":2848
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":2868
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtTrailBETicks 
                  Height          =   285
                  Left            =   2280
                  TabIndex        =   86
                  Top             =   75
                  Width           =   915
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   0   'False
                  Locked          =   0   'False
                  Text            =   "frmExitStrategy.frx":2884
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
                  Tip             =   "frmExitStrategy.frx":28AC
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":28CC
               End
               Begin HexUniControls.ctlUniCheckXP chkTrailBE 
                  Height          =   195
                  Left            =   0
                  TabIndex        =   85
                  Top             =   120
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
                  Caption         =   "frmExitStrategy.frx":28E8
                  Enabled         =   -1  'True
                  Align           =   0
                  CheckBackColor  =   -2147483643
                  CheckForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmExitStrategy.frx":2942
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":2962
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin gdOCX.gdScrollBar sbTrailBETicks 
                  Height          =   360
                  Left            =   3180
                  TabIndex        =   87
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   635
               End
               Begin HexUniControls.ctlUniLabelXP lblTrailBETicks 
                  Height          =   195
                  Left            =   3480
                  Top             =   120
                  Width           =   435
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
                  Caption         =   "frmExitStrategy.frx":297E
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":29A8
                  Style           =   0
                  Enabled         =   0   'False
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":29C8
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniFrameWL fraAfter 
               Height          =   375
               Left            =   60
               TabIndex        =   74
               Top             =   0
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
               Caption         =   "frmExitStrategy.frx":29E4
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmExitStrategy.frx":2A10
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":2A30
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtAfterTicks 
                  Height          =   285
                  Left            =   1740
                  TabIndex        =   76
                  Top             =   45
                  Width           =   915
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   0   'False
                  Locked          =   0   'False
                  Text            =   "frmExitStrategy.frx":2A4C
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
                  Tip             =   "frmExitStrategy.frx":2A74
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":2A94
               End
               Begin gdOCX.gdScrollBar sbAfterTicks 
                  Height          =   360
                  Left            =   2640
                  TabIndex        =   77
                  TabStop         =   0   'False
                  Top             =   7
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   635
               End
               Begin HexUniControls.ctlUniLabelXP lblAfter 
                  Height          =   255
                  Left            =   0
                  Top             =   60
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
                  Caption         =   "frmExitStrategy.frx":2AB0
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":2AFC
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":2B1C
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblAfterTicks 
                  Height          =   195
                  Left            =   2940
                  Top             =   90
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
                  Caption         =   "frmExitStrategy.frx":2B38
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":2B9C
                  Style           =   0
                  Enabled         =   0   'False
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":2BBC
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniFrameWL fraMoveWhen 
               Height          =   375
               Left            =   360
               TabIndex        =   79
               Top             =   420
               Width           =   5175
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
               Caption         =   "frmExitStrategy.frx":2BD8
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmExitStrategy.frx":2C0E
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmExitStrategy.frx":2C2E
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtMoveTicks 
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   81
                  Top             =   45
                  Width           =   915
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   0   'False
                  Locked          =   0   'False
                  Text            =   "frmExitStrategy.frx":2C4A
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
                  Tip             =   "frmExitStrategy.frx":2C72
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":2C92
               End
               Begin gdOCX.gdScrollBar sbMoveTicks 
                  Height          =   360
                  Left            =   2460
                  TabIndex        =   82
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   635
               End
               Begin HexUniControls.ctlUniLabelXP lblMoveTo 
                  Height          =   255
                  Left            =   0
                  Top             =   60
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
                  Caption         =   "frmExitStrategy.frx":2CAE
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":2CF4
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":2D14
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblMoveTicks 
                  Height          =   195
                  Left            =   2760
                  Top             =   90
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
                  Caption         =   "frmExitStrategy.frx":2D30
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmExitStrategy.frx":2D92
                  Style           =   0
                  Enabled         =   0   'False
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmExitStrategy.frx":2DB2
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
         End
         Begin gdOCX.gdScrollBar sbStopTicks 
            Height          =   360
            Left            =   3000
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   0
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
         Begin HexUniControls.ctlUniLabelXP lblStopLoss 
            Height          =   195
            Left            =   60
            Top             =   90
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
            Caption         =   "frmExitStrategy.frx":2DCE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmExitStrategy.frx":2E28
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":2E48
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblStopTicks 
            Height          =   195
            Left            =   3300
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
            Caption         =   "frmExitStrategy.frx":2E64
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmExitStrategy.frx":2E98
            Style           =   0
            Enabled         =   0   'False
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmExitStrategy.frx":2EB8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblDescription 
      Height          =   255
      Left            =   180
      Top             =   150
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
      Caption         =   "frmExitStrategy.frx":2ED4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmExitStrategy.frx":2F0C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExitStrategy.frx":2F2C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuCopy 
      Caption         =   "Copy"
      Begin VB.Menu mnuLongProfitTarget 
         Caption         =   "Long Profit Target"
      End
      Begin VB.Menu mnuShortProfitTarget 
         Caption         =   "Short Profit Target"
      End
      Begin VB.Menu mnuLongStopLoss 
         Caption         =   "Long Stop Loss"
      End
      Begin VB.Menu mnuShortStopLoss 
         Caption         =   "Short Stop Loss"
      End
   End
End
Attribute VB_Name = "frmExitStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmExitStrategy.frm
'' Description: Allow the user to specify and manage their exit strategy for an
''              order or a position
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/11/2010   DAJ         Use the global OrderStrategies collection
'' 05/17/2010   DAJ         Added support for TradeSense orders
'' 06/14/2010   DAJ         Don't error if only Profit Target Ts order specified (#5774)
'' 06/14/2010   DAJ         Added ability to clear out Trade Sense expressions (#5773)
'' 08/04/2010   DAJ         Added flag file for DanielCode/TradeSense Orders/Groups
'' 12/01/2010   DAJ         Require Gold for TradeSense auto exits instead of flag file
'' 04/17/2013   DAJ         Flatten if Stop on wrong side, Cancel manual orders on opposite side
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDExitType
    eGDExitType_LongProfit = 0
    eGDExitType_ShortProfit
    eGDExitType_LongStop
    eGDExitType_ShortStop
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK or Cancel?
    bSkipEnable As Boolean              ' Do we want to temporarily skip the enable/disable?
    bSaveAs As Boolean                  ' Has the user done a successful save as?
    bDirty As Boolean                   ' Is the form dirty?
    bLoadRtfs As Boolean                ' Do we need to load the RTF boxes?
    
    ExitStrategy As cExitStrategy       ' Local version of the exit strategy
    
    TargetTicks1 As cPriceEditor        ' Spinner control for the first profit target tick amount
    TargetQty1 As cPriceEditor          ' Spinner control for the first profit target quantity
    TargetTicks2 As cPriceEditor        ' Spinner control for the second profit target tick amount
    TargetQty2 As cPriceEditor          ' Spinner control for the second profit target quantity
    TargetTicks3 As cPriceEditor        ' Spinner control for the third profit target tick amount
    TargetQty3 As cPriceEditor          ' Spinner control for the third profit target quantity
    TargetQtyTs As cPriceEditor         ' Spinner control for the Trade Sense profit target quantity
    StopLossTicks As cPriceEditor       ' Spinner control for the stop loss tick amount
    WithLimitTicks As cPriceEditor      ' Spinner control for the with limit tick amount
    AfterTicks As cPriceEditor          ' Spinner control for the after tick amount
    MoveToTicks As cPriceEditor         ' Spinner control for the move to tick amount
    TrailTicks As cPriceEditor          ' Spinner control for the trailing stop tick amount
    
    tsoProfitLong As cTradeSenseOrder   ' Working copy of the long profit trade sense order
    tsoProfitShort As cTradeSenseOrder  ' Working copy of the short profit trade sense order
    tsoStopLong As cTradeSenseOrder     ' Working copy of the long stop trade sense order
    tsoStopShort As cTradeSenseOrder    ' Working copy of the short stop trade sense order
    
    nPrevStop As eGDStopLossType        ' Previously chosen stop loss type
End Type
Private m As mPrivate

Private Property Get Dirty() As Boolean
    Dirty = m.bDirty
End Property
Private Property Let Dirty(ByVal bDirty As Boolean)
    m.bDirty = bDirty
    EnableControls
End Property

Private Function ExitType(ByVal nExitType As eGDExitType) As Long
    ExitType = nExitType
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Exit Strategy, Is it Live?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ExitStrategy As cExitStrategy, ByVal bLiveVersion As Boolean, Optional bSaveAs As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Set m.ExitStrategy = ExitStrategy
    
    If bLiveVersion Then
        fraEditorButtons.Visible = False
        fraOKCancel.Top = 0
        fraOKCancel.Visible = True
    Else
        fraOKCancel.Visible = False
        fraEditorButtons.Top = 0
        fraEditorButtons.Visible = True
    End If
    
    fraProfitTargetOptions.Visible = True
    optTsStop.Visible = True
    
    SetUpSpinners
    LoadControlsFromObject
    
    SetEditorCaption Me, "Exit Order Strategy", ExitStrategy.StrategyName
    
    Dirty = False
    ShowForm Me, eForm_Modal, frmMain
    
    Set ExitStrategy = m.ExitStrategy
    bSaveAs = m.bSaveAs
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmExitStrategy.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCancelOpposite_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCancelOpposite_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.chkCancelOpposite_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkEntirePosition_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkEntirePosition_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.chkEntirePosition_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkEntirePositionTs_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkEntirePositionTs_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.chkEntirePositionTs_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkMarketIfWrongSide_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkMarketIfWrongSide_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.chkMarketIfWrongSide_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkProfitTarget1_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkProfitTarget1_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.chkProfitTarget1_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkProfitTarget2_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkProfitTarget2_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.chkProfitTarget2_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkProfitTarget3_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkProfitTarget3_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.chkProfitTarget3_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkTrailBE_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkTrailBE_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.chkTrailBE_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkWithLimit_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkWithLimit_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.chkWithLimit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow ShowMe to unload the form without submitting orders
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
    RaiseError "frmExitStrategy.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClearLongProfit_Click
'' Description: Allow the user to clear the expression
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearLongProfit_Click()
On Error GoTo ErrSection:

    ClearTradeSenseOrder eGDExitType_LongProfit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdClearLongProfit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClearLongStop_Click
'' Description: Allow the user to clear the expression
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearLongStop_Click()
On Error GoTo ErrSection:

    ClearTradeSenseOrder eGDExitType_LongStop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdClearLongStop_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClearShortProfit_Click
'' Description: Allow the user to clear the expression
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearShortProfit_Click()
On Error GoTo ErrSection:

    ClearTradeSenseOrder eGDExitType_ShortProfit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdClearShortProfit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClearShortStop_Click
'' Description: Allow the user to clear the expression
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearShortStop_Click()
On Error GoTo ErrSection:

    ClearTradeSenseOrder eGDExitType_ShortStop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdClearShortStop_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCopyLongProfit_Click
'' Description: Allow the user to copy one of the other Trade Sense orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCopyLongProfit_Click()
On Error GoTo ErrSection:

    ShowCopyMenu cmdCopyLongProfit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdCopyLongProfit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCopyLongStop_Click
'' Description: Allow the user to copy one of the other Trade Sense orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCopyLongStop_Click()
On Error GoTo ErrSection:

    ShowCopyMenu cmdCopyLongStop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdCopyLongStop_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCopyShortProfit_Click
'' Description: Allow the user to copy one of the other Trade Sense orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCopyShortProfit_Click()
On Error GoTo ErrSection:

    ShowCopyMenu cmdCopyShortProfit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdCopyShortProfit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCopyShortStop_Click
'' Description: Allow the user to copy one of the other Trade Sense orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCopyShortStop_Click()
On Error GoTo ErrSection:

    ShowCopyMenu cmdCopyShortStop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdCopyShortStop_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditLongProfit_Click
'' Description: Allow the user to edit the long profit Trade Sense order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditLongProfit_Click()
On Error GoTo ErrSection:

    EditTradeSenseOrder eGDExitType_LongProfit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdEditLongProfit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditLongStop_Click
'' Description: Allow the user to edit the long stop Trade Sense order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditLongStop_Click()
On Error GoTo ErrSection:

    EditTradeSenseOrder eGDExitType_LongStop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdEditLongStop_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditShortProfit_Click
'' Description: Allow the user to edit the short profit Trade Sense order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditShortProfit_Click()
On Error GoTo ErrSection:

    EditTradeSenseOrder eGDExitType_ShortProfit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdEditShortProfit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditShortStop_Click
'' Description: Allow the user to edit the short stop Trade Sense order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditShortStop_Click()
On Error GoTo ErrSection:

    EditTradeSenseOrder eGDExitType_ShortStop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdEditShortStop_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdExit_Click
'' Description: Allow the user to exit the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExit_Click()
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the message box to the user

    If Dirty Then
        strReturn = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", Caption)
        Select Case strReturn
            Case "Y"
                Save
                If Dirty Then Exit Sub
                
            Case "N"
            
            Case "C"
                Exit Sub
        
        End Select
    End If

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdExit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Verify form information and let ShowMe unload and submit orders
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
    RaiseError "frmExitStrategy.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRename_Click
'' Description: Allow the user to rename the exit order strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRename_Click()
On Error GoTo ErrSection:

    Dim strName As String               ' User supplied name
    Dim bValidName As Boolean           ' Is the supplied name valid?
    Dim strOldName As String            ' Old strategy name

    If VerifyControls Then
        Do
            strName = InfBox("Please supply a new name for the exit order strategy:", , , "Exit Order Strategy Name", , , , , , "string", m.ExitStrategy.StrategyName)
            If Len(strName) > 0 Then bValidName = VerifyName(strName)
        Loop While (bValidName = False) And (Len(strName) > 0)
        
        If Len(strName) > 0 Then
            strOldName = m.ExitStrategy.FileName
            If strName <> m.ExitStrategy.StrategyName Then
                KillFile AddSlash(App.Path) & m.ExitStrategy.FileName, True
            End If
            m.ExitStrategy.StrategyName = strName
            SaveControlsToObject
            m.ExitStrategy.Save
            
            SetEditorCaption Me, "Exit Order Strategy", m.ExitStrategy.StrategyName
            Dirty = False
        
            ' Refresh any active exits that are using this strategy...
            If strOldName <> m.ExitStrategy.FileName Then
                If Not g.OrderStrategies Is Nothing Then
                    g.OrderStrategies.ExitStrategyRenamed strOldName, m.ExitStrategy.FileName
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdRename_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSave_Click
'' Description: Allow the user to save the exit order strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSave_Click()
On Error GoTo ErrSection:

    Save

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdSave_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSaveAs_Click
'' Description: Allow the user to save a copy of the exit order strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSaveAs_Click()
On Error GoTo ErrSection:

    Dim strName As String               ' User supplied name
    Dim bValidName As Boolean           ' Is the supplied name valid?
    Dim NewExit As New cExitStrategy    ' New copy of the exit strategy

    If VerifyControls Then
        Do
            strName = InfBox("Please supply a name for the new exit order strategy:", , , "Exit Order Strategy Name", , , , , , "string", "Copy of " & m.ExitStrategy.StrategyName)
            If Len(strName) > 0 Then bValidName = VerifyName(strName, True)
        Loop While (bValidName = False) And (Len(strName) > 0)
        
        If Len(strName) > 0 Then
            Set NewExit = m.ExitStrategy.MakeCopy
            Set m.ExitStrategy = NewExit
            m.ExitStrategy.StrategyName = strName
            m.ExitStrategy.Provided = False
            SaveControlsToObject
            m.ExitStrategy.Save
            m.bSaveAs = True
            
            SetEditorCaption Me, "Exit Order Strategy", m.ExitStrategy.StrategyName
            Dirty = False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.cmdSaveAs_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Perform tasks when the form is activated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    If m.bLoadRtfs = True Then
        m.bLoadRtfs = False
        
        ' For some reason, the RTF doesn't work at load time, so we need to
        ' do it the first time the form is activated...
        rtbTsLongProfit.TextRTF = m.tsoProfitLong.PreviewRTF
        rtbTsShortProfit.TextRTF = m.tsoProfitShort.PreviewRTF
        rtbTsLongStop.TextRTF = m.tsoStopLong.PreviewRTF
        rtbTsShortStop.TextRTF = m.tsoStopShort.PreviewRTF
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize and setup the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Caption = "Exit Order Strategy"
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    lblProfitTarget1Pos.Left = fraProfitTarget1Qty.Left
    lblProfitTarget1Pos.Width = chkEntirePosition.Width
    lblProfitTarget2Pos.Left = fraProfitTarget2Qty.Left
    lblProfitTarget2Pos.Width = chkEntirePosition.Width
    lblProfitTarget3Pos.Left = fraProfitTarget3Qty.Left
    lblProfitTarget3Pos.Width = chkEntirePosition.Width
    
    m.bSaveAs = False
    m.bLoadRtfs = False
    
    mnuCopy.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, allow ShowMe to unload the form
'' Inputs:      Cancel the Unload?, Mode of the Unload
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
    RaiseError "frmExitStrategy.Form_QueryUnload"
    
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

    Dim lFrameWidth As Long             ' Frame width

    If LimitFormSize(Me, 7410, 7545) = False Then
        lFrameWidth = fraButtons.Left - 240
        
        With fraOptions
            .Move 120, ScaleHeight - .Height - 120, lFrameWidth
        End With
        
        With fraStopLoss
            .Move 120, fraOptions.Top - .Height - 120, lFrameWidth
        End With
        
        With fraProfitTargets
            .Move 120, .Top, lFrameWidth, fraStopLoss.Top - .Top - 120
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up after the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.ExitStrategy = Nothing
    
    Set m.TargetTicks1 = Nothing
    Set m.TargetQty1 = Nothing
    Set m.TargetQty2 = Nothing
    Set m.TargetQty3 = Nothing
    Set m.TargetQtyTs = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmExitStrategy.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuLongProfitTarget_Click
'' Description: Copy the appropriate Trade Sense order to this one
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuLongProfitTarget_Click()
On Error GoTo ErrSection:

    CopyTradeSenseOrder eGDExitType_LongProfit, CLng(Val(mnuCopy.Tag))
    mnuCopy.Tag = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.mnuLongProfitTarget_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuLongStopLoss_Click
'' Description: Copy the appropriate Trade Sense order to this one
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuLongStopLoss_Click()
On Error GoTo ErrSection:

    CopyTradeSenseOrder eGDExitType_LongStop, CLng(Val(mnuCopy.Tag))
    mnuCopy.Tag = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.mnuLongStopLoss_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuShortProfitTarget_Click
'' Description: Copy the appropriate Trade Sense order to this one
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuShortProfitTarget_Click()
On Error GoTo ErrSection:

    CopyTradeSenseOrder eGDExitType_ShortProfit, CLng(Val(mnuCopy.Tag))
    mnuCopy.Tag = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.mnuShortProfitTarget_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuShortStopLoss_Click
'' Description: Copy the appropriate Trade Sense order to this one
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuShortStopLoss_Click()
On Error GoTo ErrSection:

    CopyTradeSenseOrder eGDExitType_ShortStop, CLng(Val(mnuCopy.Tag))
    mnuCopy.Tag = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.mnuShortStopLoss_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optBreakEven_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optBreakEven_Click()
On Error GoTo ErrSection:

    lblStopLoss.Caption = "Place Initial Stop Order at a"
    lblStopTicks.Caption = "tick loss."
    txtStopTicks.Left = 2220
    sbStopTicks.Left = 3120
    lblStopTicks.Left = 3420
    
    m.nPrevStop = eGDStopLossType_BreakEven
    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optBreakEven_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optEqualLots_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optEqualLots_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optEqualLots_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optFixedStop_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optFixedStop_Click()
On Error GoTo ErrSection:

    lblStopLoss.Caption = "Place Fixed Stop Order at a"
    lblStopTicks.Caption = "tick loss."
    txtStopTicks.Left = 2220
    sbStopTicks.Left = 3120
    lblStopTicks.Left = 3420
    
    m.nPrevStop = eGDStopLossType_Fixed
    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optFixedStop_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optFullTicks_Click
'' Description: Change the spinners as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optFullTicks_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
        ChangeSpinners
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optFullTicks_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optHalfTicks_Click
'' Description: Change the spinners as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optHalfTicks_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
        ChangeSpinners
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optHalfTicks_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optNoStop_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optNoStop_Click()
On Error GoTo ErrSection:

    lblStopLoss.Caption = "Place Fixed Stop Order at a"
    lblStopTicks.Caption = "tick loss."
    txtStopTicks.Left = 2220
    sbStopTicks.Left = 3120
    lblStopTicks.Left = 3420
    
    m.nPrevStop = eGDStopLossType_None
    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optNoStop_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optQuarterTicks_Click
'' Description: Change the spinners as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optQuarterTicks_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
        ChangeSpinners
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optQuarterTicks_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSpecifyLots_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSpecifyLots_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optSpecifyLots_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optStandardProfit_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optStandardProfit_Click()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optStandardProfit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTrailStop_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTrailStop_Click()
On Error GoTo ErrSection:

    lblStopLoss.Caption = "Place a Stop Order to trail the market by"
    lblStopTicks.Caption = "ticks."
    txtStopTicks.Left = 3060
    sbStopTicks.Left = 3960
    lblStopTicks.Left = 4260
    
    m.nPrevStop = eGDStopLossType_Trail
    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optTrailStop_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTsProfit_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTsProfit_Click()
On Error GoTo ErrSection:

    If Visible Then
        If HasLevel(eTN4_Gold, True, "TradeSense Auto Exits") = True Then
            Dirty = True
        Else
            optStandardProfit.Value = True
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optTsProfit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTsStop_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTsStop_Click()
On Error GoTo ErrSection:

    If Visible Then
        If HasLevel(eTN4_Gold, True, "TradeSense Auto Exits") = True Then
            Dirty = True
        Else
            Select Case m.nPrevStop
                Case eGDStopLossType_None
                    optNoStop.Value = True
                Case eGDStopLossType_Fixed
                    optFixedStop.Value = True
                Case eGDStopLossType_Trail
                    optTrailStop.Value = True
                Case eGDStopLossType_BreakEven
                    optBreakEven.Value = True
            End Select
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.optTsStop_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAfterTicks_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAfterTicks_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtAfterTicks_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDescription_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDescription_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtDescription_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtMoveTicks_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtMoveTicks_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtMoveTicks_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStopTicks_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStopTicks_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtStopTicks_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetQty1_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetQty1_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtTargetQty1_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetQty2_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetQty2_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtTargetQty2_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetQty3_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetQty3_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtTargetQty3_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetQtyTs_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetQtyTs_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtTargetQtyTs_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetTicks1_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetTicks1_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtTargetTicks1_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetTicks2_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetTicks2_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtTargetTicks2_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetTicks3_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetTicks3_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtTargetTicks3_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTrailBETicks_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTrailBETicks_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtTrailBETicks_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtWithLimitTicks_Change
'' Description: Once the control is changed, make sure to dirty the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtWithLimitTicks_Change()
On Error GoTo ErrSection:

    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.txtWithLimitTicks_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpSpinners
'' Description: Set up and initialize the price editor controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpSpinners()
On Error GoTo ErrSection:

    Dim dMinMove As Double              ' Min move from the exit strategy
    
    Select Case m.ExitStrategy.TickMode
        Case 0
            dMinMove = 1
        Case 1
            dMinMove = 0.5
        Case 2
            dMinMove = 0.25
    End Select

    Set m.TargetTicks1 = New cPriceEditor
    m.TargetTicks1.Init sbTargetTicks1, txtTargetTicks1, Nothing, 1, 1, , , , dMinMove
    Set m.TargetQty1 = New cPriceEditor
    m.TargetQty1.Init sbTargetQty1, txtTargetQty1, Nothing, 1, 1

    Set m.TargetTicks2 = New cPriceEditor
    m.TargetTicks2.Init sbTargetTicks2, txtTargetTicks2, Nothing, 1, 1, , , , dMinMove
    Set m.TargetQty2 = New cPriceEditor
    m.TargetQty2.Init sbTargetQty2, txtTargetQty2, Nothing, 1, 1

    Set m.TargetTicks3 = New cPriceEditor
    m.TargetTicks3.Init sbTargetTicks3, txtTargetTicks3, Nothing, 1, 1, , , , dMinMove
    Set m.TargetQty3 = New cPriceEditor
    m.TargetQty3.Init sbTargetQty3, txtTargetQty3, Nothing, 1, 1
    
    Set m.TargetQtyTs = New cPriceEditor
    m.TargetQtyTs.Init sbTargetQtyTs, txtTargetQtyTs, Nothing, 1, 1
    
    Set m.StopLossTicks = New cPriceEditor
    If FileExist(AddSlash(App.Path) & "AllowNegStop.FLG") Then
        m.StopLossTicks.Init sbStopTicks, txtStopTicks, Nothing, 1, -999999, , , , dMinMove
    Else
        m.StopLossTicks.Init sbStopTicks, txtStopTicks, Nothing, 1, 1, , , , dMinMove
    End If
    Set m.WithLimitTicks = New cPriceEditor
    m.WithLimitTicks.Init sbWithLimitTicks, txtWithLimitTicks, Nothing, 1, 1, , , , dMinMove
    Set m.AfterTicks = New cPriceEditor
    m.AfterTicks.Init sbAfterTicks, txtAfterTicks, Nothing, 1, 1, , , , dMinMove
    Set m.MoveToTicks = New cPriceEditor
    m.MoveToTicks.Init sbMoveTicks, txtMoveTicks, Nothing, 0, -999999, , , True, dMinMove
    Set m.TrailTicks = New cPriceEditor
    m.TrailTicks.Init sbTrailBETicks, txtTrailBETicks, Nothing, 1, 1, , , , dMinMove

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.SetUpSpinners"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bHasLongProfitTS As Boolean     ' Is there a Trade Sense long profit?
    Dim bHasShortProfitTS As Boolean    ' Is there a Trade Sense short profit?
    Dim bHasLongStopTS As Boolean       ' Is there a Trade Sense long stop?
    Dim bHasShortStopTS As Boolean      ' Is there a Trade Sense short stop?

    If m.bSkipEnable = False Then
        ' Profit Target 2 can only be utilized if Profit Target 1 is turned on...
        If chkProfitTarget1.Value = vbUnchecked Then
            chkProfitTarget2.Value = vbUnchecked
        End If
   
        ' Profit Target 3 can only be utilized if Profit Targets 1 & 2 are turned on...
        If chkProfitTarget2.Value = vbUnchecked Then
            chkProfitTarget3.Value = vbUnchecked
        End If
        
        ShowQuantityControls
        ShowTradeSenseControls
        
        ' Don't allow user to do break even controls unless break even is turned on...
        If optBreakEven.Value = False Then
            chkTrailBE.Value = vbUnchecked
        End If
        fraBreakEven.Visible = optBreakEven.Value
        
        ' Don't allow the stop limit option if the stop loss is turned off...
        If optNoStop.Value = True Then
            chkWithLimit.Value = vbUnchecked
        End If
        
        ' Enable/Disable controls for the Target Profit 3...
        Enable txtTargetTicks3, (chkProfitTarget3.Value = vbChecked)
        Enable sbTargetTicks3, (chkProfitTarget3.Value = vbChecked)
        Enable lblTicks3, (chkProfitTarget3.Value = vbChecked)
        Enable lblTargetQty3, (chkProfitTarget3.Value = vbChecked) And (optSpecifyLots.Value = True)
        Enable txtTargetQty3, (chkProfitTarget3.Value = vbChecked) And (optSpecifyLots.Value = True)
        Enable sbTargetQty3, (chkProfitTarget3.Value = vbChecked) And (optSpecifyLots.Value = True)
        
        ' Enable/Disable controls for the Target Profit 2...
        Enable txtTargetTicks2, (chkProfitTarget2.Value = vbChecked)
        Enable sbTargetTicks2, (chkProfitTarget2.Value = vbChecked)
        Enable lblTicks2, (chkProfitTarget2.Value = vbChecked)
        Enable lblTargetQty2, (chkProfitTarget2.Value = vbChecked) And (optSpecifyLots.Value = True)
        Enable txtTargetQty2, (chkProfitTarget2.Value = vbChecked) And (optSpecifyLots.Value = True)
        Enable sbTargetQty2, (chkProfitTarget2.Value = vbChecked) And (optSpecifyLots.Value = True)
        
        ' Enable/Disable controls for the Target Profit 1...
        Enable txtTargetTicks1, (chkProfitTarget1.Value = vbChecked)
        Enable sbTargetTicks1, (chkProfitTarget1.Value = vbChecked)
        Enable lblTicks1, (chkProfitTarget1.Value = vbChecked)
        Enable lblTargetQty1, (chkProfitTarget1.Value = vbChecked) And (optSpecifyLots.Value = True)
        Enable txtTargetQty1, (chkProfitTarget1.Value = vbChecked) And (optSpecifyLots.Value = True)
        Enable sbTargetQty1, (chkProfitTarget1.Value = vbChecked) And (optSpecifyLots.Value = True)
        
        ' Enable/Disable controls for the Trade Sense Profit Target...
        Enable lblTargetQtyTs, Not CheckBoxValue(chkEntirePositionTs)
        Enable txtTargetQtyTs, Not CheckBoxValue(chkEntirePositionTs)
        Enable sbTargetQtyTs, Not CheckBoxValue(chkEntirePositionTs)
        
        ' Enable/Disable controls for the With Limit Stop option...
        Enable txtWithLimitTicks, (chkWithLimit.Value = vbChecked)
        Enable sbWithLimitTicks, (chkWithLimit.Value = vbChecked)
        Enable lblWithLimitTicks, (chkWithLimit.Value = vbChecked)
        
        ' Enable/Disable controls for the Break Even Move Stop option...
        Enable lblMoveTo, optBreakEven
        Enable txtMoveTicks, optBreakEven
        Enable sbMoveTicks, optBreakEven
        Enable lblMoveTicks, optBreakEven
        
        ' Enable/Disable controls for the Break Even TrailBE Stop option...
        Enable txtTrailBETicks, (chkTrailBE.Value = vbChecked)
        Enable sbTrailBETicks, (chkTrailBE.Value = vbChecked)
        Enable lblTrailBETicks, (chkTrailBE.Value = vbChecked)
        
        ' Enable/Disable Stop Loss controls as long as user chooses to do a stop loss...
        Enable lblStopLoss, (optNoStop.Value = False)
        Enable txtStopTicks, (optNoStop.Value = False)
        Enable sbStopTicks, (optNoStop.Value = False)
        Enable lblStopTicks, (optNoStop.Value = False)
        Enable chkWithLimit, (optNoStop.Value = False)
        
        ' Enable/Disable the Break Even after X Ticks controls...
        Enable txtAfterTicks, (optBreakEven.Value = True)
        Enable sbAfterTicks, (optBreakEven.Value = True)
        Enable lblAfterTicks, (optBreakEven.Value = True)
        
        ' Don't allow the user to save a provided strategy...
        Enable cmdSave, Not m.ExitStrategy.Provided
        Enable cmdSaveAs, (Len(m.ExitStrategy.StrategyName) > 0)
        Enable cmdRename, (Not m.ExitStrategy.Provided) And (Len(m.ExitStrategy.StrategyName) > 0)
        
        ' Only enable the entire position check box if specify lots (if equal lots then force to
        ' entire position)...
        If optEqualLots.Value = True Then
            chkEntirePosition.Value = vbChecked
            chkEntirePosition.Enabled = False
        Else
            chkEntirePosition.Enabled = True
        End If
        
        If Not m.tsoProfitLong Is Nothing Then
            bHasLongProfitTS = (Len(m.tsoProfitLong.ConditionCoded) > 0)
        End If
        If Not m.tsoProfitShort Is Nothing Then
            bHasShortProfitTS = (Len(m.tsoProfitShort.ConditionCoded) > 0)
        End If
        If Not m.tsoStopLong Is Nothing Then
            bHasLongStopTS = (Len(m.tsoStopLong.ConditionCoded) > 0)
        End If
        If Not m.tsoStopShort Is Nothing Then
            bHasShortStopTS = (Len(m.tsoStopShort.ConditionCoded) > 0)
        End If
        
        Enable cmdCopyLongProfit, (bHasShortProfitTS Or bHasLongStopTS Or bHasShortStopTS)
        Enable cmdCopyShortProfit, (bHasLongProfitTS Or bHasLongStopTS Or bHasShortStopTS)
        Enable cmdCopyLongStop, (bHasLongProfitTS Or bHasShortProfitTS Or bHasShortStopTS)
        Enable cmdCopyShortStop, (bHasLongProfitTS Or bHasShortProfitTS Or bHasLongStopTS)
        
        Enable cmdClearLongProfit, bHasLongProfitTS
        Enable cmdClearShortProfit, bHasShortProfitTS
        Enable cmdClearLongStop, bHasLongStopTS
        Enable cmdClearShortStop, bHasShortStopTS
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadControlsFromObject
'' Description: Load up the controls from the local exit strategy object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadControlsFromObject()
On Error GoTo ErrSection:

    m.bSkipEnable = True
    
    With m.ExitStrategy
        txtDescription.Text = .Description
        optSpecifyLots.Value = .SpecifyLots
        If .ExitEntirePosition Then chkEntirePosition.Value = vbChecked Else chkEntirePosition.Value = vbUnchecked
        optEqualLots.Value = Not .SpecifyLots
        If .UseTarget3 Then chkProfitTarget3.Value = vbChecked Else chkProfitTarget3.Value = vbUnchecked
        m.TargetTicks3.Price = .Target3Ticks
        m.TargetQty3.Price = .Target3Quantity
        If .UseTarget2 Then chkProfitTarget2.Value = vbChecked Else chkProfitTarget2.Value = vbUnchecked
        m.TargetTicks2.Price = .Target2Ticks
        m.TargetQty2.Price = .Target2Quantity
        If .UseTarget1 Then chkProfitTarget1.Value = vbChecked Else chkProfitTarget1.Value = vbUnchecked
        m.TargetTicks1.Price = .Target1Ticks
        m.TargetQty1.Price = .Target1Quantity
        
        Select Case .StopLossType
            Case eGDStopLossType_None
                optNoStop.Value = True
                optFixedStop.Value = False
                optBreakEven.Value = False
                optTrailStop.Value = False
                optTsStop.Value = False
                
            Case eGDStopLossType_Fixed
                optNoStop.Value = False
                optFixedStop.Value = True
                optBreakEven.Value = False
                optTrailStop.Value = False
                optTsStop.Value = False
                
            Case eGDStopLossType_BreakEven
                optNoStop.Value = False
                optFixedStop.Value = False
                optBreakEven.Value = True
                optTrailStop.Value = False
                optTsStop.Value = False
            
            Case eGDStopLossType_Trail
                optNoStop.Value = False
                optFixedStop.Value = False
                optBreakEven.Value = False
                optTrailStop.Value = True
                optTsStop.Value = False
                
            Case eGDStopLossType_TradeSense
                optNoStop.Value = False
                optFixedStop.Value = False
                optBreakEven.Value = False
                optTrailStop.Value = False
                optTsStop.Value = True
        End Select
        m.nPrevStop = .StopLossType
        
        m.StopLossTicks = .StopLossTicks
        If .UseWithLimit Then chkWithLimit.Value = vbChecked Else chkWithLimit.Value = vbUnchecked
        m.WithLimitTicks = .WithLimitTicks
        m.AfterTicks = .AfterTicks
        m.MoveToTicks = .MoveToTicks
        If .UseTrail Then chkTrailBE.Value = vbChecked Else chkTrailBE.Value = vbUnchecked
        m.TrailTicks = .TrailTicks
        
        Select Case .TickMode
            Case 0
                optFullTicks.Value = True
            Case 1
                optHalfTicks.Value = True
            Case 2
                optQuarterTicks.Value = True
        End Select
    
        ' 04/14/2010 DAJ: New fields for the new Trade Sense auto exit orders...
        Select Case .ProfitTargetType
            Case eGDProfitTargetType_Standard
                optStandardProfit.Value = True
                optTsProfit.Value = False
                
            Case eGDProfitTargetType_TradeSense
                optStandardProfit.Value = False
                optTsProfit.Value = True
                
                CheckBoxValue(chkEntirePositionTs) = .ExitEntirePosition
                m.TargetQtyTs.Price = .Target1Quantity
        End Select
        
        Set m.tsoProfitLong = .TsProfitLong.MakeCopy
        Set m.tsoProfitShort = .TsProfitShort.MakeCopy
        Set m.tsoStopLong = .TsStopLong.MakeCopy
        Set m.tsoStopShort = .TsStopShort.MakeCopy
        m.bLoadRtfs = True
        
        ' 04/16/2013 DAJ: New fields for how to handle situations...
        CheckBoxValue(chkCancelOpposite) = .CancelOpposite
        CheckBoxValue(chkMarketIfWrongSide) = .MarketIfWrongSide
    End With
    
    m.bSkipEnable = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.LoadControlsFromObject"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveControlsToObject
'' Description: Save the controls to the local exit strategy object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveControlsToObject()
On Error GoTo ErrSection:

    With m.ExitStrategy
        .Description = Trim(txtDescription.Text)
        
        Select Case True
            Case optStandardProfit.Value = True
                .ProfitTargetType = eGDProfitTargetType_Standard
                .SpecifyLots = optSpecifyLots.Value
                .ExitEntirePosition = (chkEntirePosition.Value = vbChecked)
                .UseTarget3 = (chkProfitTarget3.Value = vbChecked)
                .Target3Ticks = m.TargetTicks3.Price
                .Target3Quantity = m.TargetQty3.Price
                .UseTarget2 = (chkProfitTarget2.Value = vbChecked)
                .Target2Ticks = m.TargetTicks2.Price
                .Target2Quantity = m.TargetQty2.Price
                .UseTarget1 = (chkProfitTarget1.Value = vbChecked)
                .Target1Ticks = m.TargetTicks1.Price
                .Target1Quantity = m.TargetQty1.Price
            
            Case optTsProfit.Value = True
                .ProfitTargetType = eGDProfitTargetType_TradeSense
                .UseTarget1 = True
                .ExitEntirePosition = CheckBoxValue(chkEntirePositionTs)
                .Target1Quantity = m.TargetQtyTs.Price
                .UseTarget2 = False
                .UseTarget3 = False
        
        End Select
        
        .TsProfitLong = m.tsoProfitLong.MakeCopy
        .TsProfitShort = m.tsoProfitShort.MakeCopy
        
        Select Case True
            Case optNoStop.Value = True
                .StopLossType = eGDStopLossType_None
            Case optFixedStop.Value = True
                .StopLossType = eGDStopLossType_Fixed
            Case optBreakEven.Value = True
                .StopLossType = eGDStopLossType_BreakEven
            Case optTrailStop.Value = True
                .StopLossType = eGDStopLossType_Trail
            Case optTsStop.Value = True
                .StopLossType = eGDStopLossType_TradeSense
        End Select
        
        .StopLossTicks = m.StopLossTicks
        .UseWithLimit = (chkWithLimit.Value = vbChecked)
        .WithLimitTicks = m.WithLimitTicks
        .AfterTicks = m.AfterTicks
        .MoveToTicks = m.MoveToTicks
        .UseTrail = (chkTrailBE.Value = vbChecked)
        .TrailTicks = m.TrailTicks
    
        .TsStopLong = m.tsoStopLong.MakeCopy
        .TsStopShort = m.tsoStopShort.MakeCopy
        
        Select Case True
            Case optFullTicks
                .TickMode = 0
            Case optHalfTicks
                .TickMode = 1
            Case optQuarterTicks
                .TickMode = 2
        End Select
        
        .CancelOpposite = CheckBoxValue(chkCancelOpposite)
        .MarketIfWrongSide = CheckBoxValue(chkMarketIfWrongSide)
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.SaveControlsToObject"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyName
'' Description: Verify the name that the user supplied
'' Inputs:      Name
'' Returns:     True if name is OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyName(ByVal strName As String, Optional ByVal bCheckSameName As Boolean = False) As Boolean
On Error GoTo ErrSection:

    If (Len(strName) <= 0) Or (Len(strName) > 50) Then
        InfBox "Strategy Name must be between 1 and 50 characters in length", "!", , "Exit Strategy Error"
        Exit Function
    End If

    If StripStr(strName, ":\/*?|><" & Chr(34)) <> strName Then
        InfBox "Strategy Name must not contain :,\,/,*,?,|,<,>, or " & Chr(34), "!", , "Exit Strategy Error"
        Exit Function
    End If
    
    If bCheckSameName And (strName = m.ExitStrategy.StrategyName) Then
        InfBox "Strategy Name must be different than the one that is currently open.", "!", , "Exit Strategy Error"
        Exit Function
    End If
    
    If (strName <> m.ExitStrategy.StrategyName) Then
        If StrategyNameExists(strName) Then
            InfBox "Strategy Name must be unique", "!", , "Exit Strategy Error"
            Exit Function
        End If
    End If
    
    VerifyName = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExitStrategy.VerifyName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyControls
'' Description: Verify the values in the controls that the user has entered
'' Inputs:      None
'' Returns:     True if everything is OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyControls() As Boolean
On Error GoTo ErrSection:

    If (ProfitSelected = False) And (StopSelected = False) Then
        InfBox "You must supply at least one profit target or a stop loss", "!", , "Exit Strategy Error"
        Exit Function
    End If
    
    If (chkProfitTarget2.Value = vbChecked) Then
        If m.TargetTicks2 <= m.TargetTicks1 Then
            MoveFocus txtTargetTicks2
            InfBox "The specified number of ticks for the second profit target must be greater than the specified number of ticks for the first profit target", "!", , "Exit Strategy Error"
            Exit Function
        End If
    End If

    If (chkProfitTarget3.Value = vbChecked) Then
        If m.TargetTicks3 <= m.TargetTicks2 Then
            MoveFocus txtTargetTicks3
            InfBox "The specified number of ticks for the third profit target must be greater than the specified number of ticks for the second profit target", "!", , "Exit Strategy Error"
            Exit Function
        End If
    End If
    
    ' If the user has chosen a trailing stop, check the StopLossTicks value to make sure it is greater than
    ' three ticks.  If not, warn the user that this might be too tight...
    If optTrailStop.Value = True Then
        If m.StopLossTicks <= 3 Then
            MoveFocus txtStopTicks
        
            If InfBox("Setting a trailing stop this close to the market can have undesired results including getting filled immediately and order rejection.||Do you want to continue?|", "?", "+Yes|-No", "Exit Strategy Warning") = "N" Then
                Exit Function
            End If
        End If
    
    ' If the user has chosen a break even stop with a trail, check the TrailTicks value to make sure it is
    ' greater than three ticks.  If not, warn the user that this might be too tight...
    ElseIf (optBreakEven.Value = True) And (chkTrailBE.Value = vbChecked) Then
        If m.TrailTicks <= 3 Then
            MoveFocus txtTrailBETicks
        
            If InfBox("Setting a trailing stop this close to the market can have undesired results including getting filled immediately and order rejection.||Do you want to continue?|", "?", "+Yes|-No", "Exit Strategy Warning") = "N" Then
                Exit Function
            End If
        End If
    End If
    
    VerifyControls = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExitStrategy.VerifyControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StrategyNameExists
'' Description: Determine if the given strategy name already exists
'' Inputs:      Strategy Name
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function StrategyNameExists(ByVal strStrategyName As String) As Boolean
On Error GoTo ErrSection:

    If FileExist(AddSlash(App.Path) & "Custom\" & strStrategyName & ".XOS") Then
        StrategyNameExists = True
    ElseIf FileExist(AddSlash(App.Path) & "Provided\" & strStrategyName & ".XOS") Then
        StrategyNameExists = True
    Else
        StrategyNameExists = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExitStrategy.StrategyNameExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Allow the user to save the exit order strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save()
On Error GoTo ErrSection:

    Dim strName As String               ' User supplied name
    Dim bValidName As Boolean           ' Is the supplied name valid?

    If VerifyControls Then
        If Len(m.ExitStrategy.StrategyName) = 0 Then
            Do
                strName = InfBox("Please supply a name for the exit order strategy:", , , "Exit Order Strategy Name", , , , , , "string")
                If Len(strName) > 0 Then bValidName = VerifyName(strName)
            Loop While (bValidName = False) And (Len(strName) > 0)
        Else
            strName = m.ExitStrategy.StrategyName
        End If
    
        If Len(strName) > 0 Then
            m.ExitStrategy.StrategyName = strName
            SaveControlsToObject
            m.ExitStrategy.Save
            
            SetEditorCaption Me, "Exit Order Strategy", m.ExitStrategy.StrategyName
            Dirty = False
        
            ' Refresh any active exits that are using this strategy...
            If Not g.OrderStrategies Is Nothing Then
                g.OrderStrategies.RefreshExitStrategy m.ExitStrategy.FileName
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowQuantityControls
'' Description: Show/Hide the profit target quantity controls as applicable
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowQuantityControls()
On Error GoTo ErrSection:

    ' Equal Lots are selected...
    If optEqualLots.Value = True Then
        ' No profit targets checked...
        If chkProfitTarget1.Value = vbUnchecked Then
            fraProfitTarget1Qty.Visible = False
            lblProfitTarget1Pos.Visible = False
            fraProfitTarget2Qty.Visible = False
            lblProfitTarget2Pos.Visible = False
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = False
        
        ' Only the first profit target is checked...
        ElseIf chkProfitTarget2.Value = vbUnchecked Then
            fraProfitTarget1Qty.Visible = False
            lblProfitTarget1Pos.Visible = True
            lblProfitTarget1Pos.Caption = "Exit Entire Position"
            fraProfitTarget2Qty.Visible = False
            lblProfitTarget2Pos.Visible = False
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = False
        
        ' Only the first and second profit targets are checked...
        ElseIf chkProfitTarget3.Value = vbUnchecked Then
            fraProfitTarget1Qty.Visible = False
            lblProfitTarget1Pos.Visible = True
            lblProfitTarget1Pos.Caption = "Exit One Half of the Position"
            fraProfitTarget2Qty.Visible = False
            lblProfitTarget2Pos.Visible = True
            lblProfitTarget2Pos.Caption = "Exit One Half of the Position"
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = False
        
        ' All three of the profit targets are checked...
        Else
            fraProfitTarget1Qty.Visible = False
            lblProfitTarget1Pos.Visible = True
            lblProfitTarget1Pos.Caption = "Exit One Third of the Position"
            fraProfitTarget2Qty.Visible = False
            lblProfitTarget2Pos.Visible = True
            lblProfitTarget2Pos.Caption = "Exit One Third of the Position"
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = True
            lblProfitTarget3Pos.Caption = "Exit One Third of the Position"
        End If
    
    ' Specify Lots and exit entire position...
    ElseIf chkEntirePosition.Value = vbChecked Then
        ' No profit targets checked...
        If chkProfitTarget1.Value = vbUnchecked Then
            fraProfitTarget1Qty.Visible = False
            lblProfitTarget1Pos.Visible = False
            fraProfitTarget2Qty.Visible = False
            lblProfitTarget2Pos.Visible = False
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = False
        
        ' Only the first profit target is checked...
        ElseIf chkProfitTarget2.Value = vbUnchecked Then
            fraProfitTarget1Qty.Visible = False
            lblProfitTarget1Pos.Visible = True
            lblProfitTarget1Pos.Caption = "Exit Entire Position"
            fraProfitTarget2Qty.Visible = False
            lblProfitTarget2Pos.Visible = False
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = False
        
        ' Only the first and second profit targets are checked...
        ElseIf chkProfitTarget3.Value = vbUnchecked Then
            fraProfitTarget1Qty.Visible = True
            lblProfitTarget1Pos.Visible = False
            lblProfitTarget1Pos.Caption = ""
            fraProfitTarget2Qty.Visible = False
            lblProfitTarget2Pos.Visible = True
            lblProfitTarget2Pos.Caption = "Exit Remainder of the Position"
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = False
        
        ' All three of the profit targets are checked...
        Else
            fraProfitTarget1Qty.Visible = True
            lblProfitTarget1Pos.Visible = False
            lblProfitTarget1Pos.Caption = ""
            fraProfitTarget2Qty.Visible = True
            lblProfitTarget2Pos.Visible = False
            lblProfitTarget2Pos.Caption = ""
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = True
            lblProfitTarget3Pos.Caption = "Exit Remainder of the Position"
        End If
    
    ' Specify Lots and take values literally...
    Else
        ' No profit targets checked...
        If chkProfitTarget1.Value = vbUnchecked Then
            fraProfitTarget1Qty.Visible = False
            lblProfitTarget1Pos.Visible = False
            fraProfitTarget2Qty.Visible = False
            lblProfitTarget2Pos.Visible = False
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = False
        
        ' Only the first profit target is checked...
        ElseIf chkProfitTarget2.Value = vbUnchecked Then
            fraProfitTarget1Qty.Visible = True
            lblProfitTarget1Pos.Visible = False
            lblProfitTarget1Pos.Caption = ""
            fraProfitTarget2Qty.Visible = False
            lblProfitTarget2Pos.Visible = False
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = False
        
        ' Only the first and second profit targets are checked...
        ElseIf chkProfitTarget3.Value = vbUnchecked Then
            fraProfitTarget1Qty.Visible = True
            lblProfitTarget1Pos.Visible = False
            lblProfitTarget1Pos.Caption = ""
            fraProfitTarget2Qty.Visible = True
            lblProfitTarget2Pos.Visible = False
            lblProfitTarget2Pos.Caption = ""
            fraProfitTarget3Qty.Visible = False
            lblProfitTarget3Pos.Visible = False
        
        ' All three of the profit targets are checked...
        Else
            fraProfitTarget1Qty.Visible = True
            lblProfitTarget1Pos.Visible = False
            lblProfitTarget1Pos.Caption = ""
            fraProfitTarget2Qty.Visible = True
            lblProfitTarget2Pos.Visible = False
            lblProfitTarget2Pos.Caption = ""
            fraProfitTarget3Qty.Visible = True
            lblProfitTarget3Pos.Visible = False
            lblProfitTarget3Pos.Caption = ""
        End If
    End If

    lblTargetQtyTs.Visible = (chkEntirePositionTs.Value = vbUnchecked)
    txtTargetQtyTs.Visible = (chkEntirePositionTs.Value = vbUnchecked)
    sbTargetQtyTs.Visible = (chkEntirePositionTs.Value = vbUnchecked)
    lblTargetTsLots.Visible = (chkEntirePositionTs.Value = vbUnchecked)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.ShowQuantityControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowTradeSenseControls
'' Description: Show/Hide the Trade Sense controls as applicable
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowTradeSenseControls()
On Error GoTo ErrSection:

    If optTsStop.Value = True Then
        fraTsStop.Move fraNonTsStop.Left, fraNonTsStop.Top
        fraTsStop.Visible = True
        fraNonTsStop.Visible = False
    Else
        fraTsStop.Visible = False
        fraNonTsStop.Visible = True
    End If
    
    If optTsProfit.Value = True Then
        fraTsProfit.Move fraNonTsProfit.Left, fraNonTsProfit.Top
        fraTsProfit.Visible = True
        fraNonTsProfit.Visible = False
    Else
        fraTsProfit.Visible = False
        fraNonTsProfit.Visible = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.ShowTradeSenseControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeSpinners
'' Description: Change the min move of the spinners as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeSpinners()
On Error GoTo ErrSection:

    Dim dMinMove As Double              ' Min move to set the price editors to

    Select Case True
        Case optFullTicks
            dMinMove = 1#
        Case optHalfTicks
            dMinMove = 0.5
        Case optQuarterTicks
            dMinMove = 0.25
    End Select

    m.TargetTicks1.ChangeMinMove dMinMove
    m.TargetTicks2.ChangeMinMove dMinMove
    m.TargetTicks3.ChangeMinMove dMinMove
    m.StopLossTicks.ChangeMinMove dMinMove
    m.WithLimitTicks.ChangeMinMove dMinMove
    m.AfterTicks.ChangeMinMove dMinMove
    m.MoveToTicks.ChangeMinMove dMinMove
    m.TrailTicks.ChangeMinMove dMinMove

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.ChangeSpinners"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowCopyMenu
'' Description: Show the copy menu for the given button
'' Inputs:      Button
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowCopyMenu(cmdCopy As CommandButton)
On Error GoTo ErrSection:

    Dim lLeft As Long                   ' Left position of the menu
    Dim lTop As Long                    ' Top position of the menu

    Select Case cmdCopy.Name
        Case "cmdCopyLongProfit"
            mnuCopy.Tag = Str(ExitType(eGDExitType_LongProfit))
            lTop = fraProfitTargets.Top + fraTsProfit.Top + cmdCopy.Top + cmdCopy.Height
            lLeft = fraProfitTargets.Left + fraTsProfit.Left + cmdCopy.Left
            
            mnuLongProfitTarget.Visible = False
            mnuShortProfitTarget.Visible = True
            mnuLongStopLoss.Visible = True
            mnuShortStopLoss.Visible = True
            
            Enable mnuShortProfitTarget, (Len(m.tsoProfitShort.ConditionCoded) > 0)
            Enable mnuLongStopLoss, (Len(m.tsoStopLong.ConditionCoded) > 0)
            Enable mnuShortStopLoss, (Len(m.tsoStopShort.ConditionCoded) > 0)
        
        Case "cmdCopyShortProfit"
            mnuCopy.Tag = Str(ExitType(eGDExitType_ShortProfit))
            lTop = fraProfitTargets.Top + fraTsProfit.Top + cmdCopy.Top + cmdCopy.Height
            lLeft = fraProfitTargets.Left + fraTsProfit.Left + cmdCopy.Left
            
            mnuLongProfitTarget.Visible = True
            mnuShortProfitTarget.Visible = False
            mnuLongStopLoss.Visible = True
            mnuShortStopLoss.Visible = True
            
            Enable mnuLongProfitTarget, (Len(m.tsoProfitLong.ConditionCoded) > 0)
            Enable mnuLongStopLoss, (Len(m.tsoStopLong.ConditionCoded) > 0)
            Enable mnuShortStopLoss, (Len(m.tsoStopShort.ConditionCoded) > 0)
        
        Case "cmdCopyLongStop"
            mnuCopy.Tag = Str(ExitType(eGDExitType_LongStop))
            lTop = fraStopLoss.Top + fraTsStop.Top + cmdCopy.Top + cmdCopy.Height
            lLeft = fraStopLoss.Left + fraTsStop.Left + cmdCopy.Left
            
            mnuLongProfitTarget.Visible = True
            mnuShortProfitTarget.Visible = True
            mnuLongStopLoss.Visible = False
            mnuShortStopLoss.Visible = True
            
            Enable mnuLongProfitTarget, (Len(m.tsoProfitLong.ConditionCoded) > 0)
            Enable mnuShortProfitTarget, (Len(m.tsoProfitShort.ConditionCoded) > 0)
            Enable mnuShortStopLoss, (Len(m.tsoStopShort.ConditionCoded) > 0)
        
        Case "cmdCopyShortStop"
            mnuCopy.Tag = Str(ExitType(eGDExitType_ShortStop))
            lTop = fraStopLoss.Top + fraTsStop.Top + cmdCopy.Top + cmdCopy.Height
            lLeft = fraStopLoss.Left + fraTsStop.Left + cmdCopy.Left
            
            mnuLongProfitTarget.Visible = True
            mnuShortProfitTarget.Visible = True
            mnuLongStopLoss.Visible = True
            mnuShortStopLoss.Visible = False
            
            Enable mnuLongProfitTarget, (Len(m.tsoProfitLong.ConditionCoded) > 0)
            Enable mnuShortProfitTarget, (Len(m.tsoProfitShort.ConditionCoded) > 0)
            Enable mnuLongStopLoss, (Len(m.tsoStopLong.ConditionCoded) > 0)
    End Select
    
    PopupMenu mnuCopy, , lLeft, lTop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.ShowCopyMenu"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditTradeSenseOrder
'' Description: Edit the Trade Sense order
'' Inputs:      Exit Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditTradeSenseOrder(ByVal nExitType As eGDExitType)
On Error GoTo ErrSection:

    Select Case nExitType
        Case ExitType(eGDExitType_LongProfit)
            If frmTradeSenseOrder.ShowMe(m.tsoProfitLong, eGDOrderAction_LongExit, eTT_OrderType_Limit, False) = True Then
                rtbTsLongProfit.TextRTF = m.tsoProfitLong.PreviewRTF
            End If
            
        Case ExitType(eGDExitType_LongStop)
            If frmTradeSenseOrder.ShowMe(m.tsoStopLong, eGDOrderAction_LongExit, eTT_OrderType_Stop, False) = True Then
                rtbTsLongStop.TextRTF = m.tsoStopLong.PreviewRTF
            End If
            
        Case ExitType(eGDExitType_ShortProfit)
            If frmTradeSenseOrder.ShowMe(m.tsoProfitShort, eGDOrderAction_ShortExit, eTT_OrderType_Limit, False) = True Then
                rtbTsShortProfit.TextRTF = m.tsoProfitShort.PreviewRTF
            End If
            
        Case ExitType(eGDExitType_ShortStop)
            If frmTradeSenseOrder.ShowMe(m.tsoStopShort, eGDOrderAction_ShortExit, eTT_OrderType_Stop, False) = True Then
                rtbTsShortStop.TextRTF = m.tsoStopShort.PreviewRTF
            End If
    
    End Select
    
    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.EditTradeSenseOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyTradeSenseOrder
'' Description: Copy the Trade Sense order
'' Inputs:      Source, Destination
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CopyTradeSenseOrder(ByVal nSource As eGDExitType, ByVal nDestination As eGDExitType)
On Error GoTo ErrSection:

    Dim tsoSource As cTradeSenseOrder   ' Source order
    
    Select Case nSource
        Case ExitType(eGDExitType_LongProfit)
            Set tsoSource = m.tsoProfitLong.MakeCopy
        Case ExitType(eGDExitType_ShortProfit)
            Set tsoSource = m.tsoProfitShort.MakeCopy
        Case ExitType(eGDExitType_LongStop)
            Set tsoSource = m.tsoStopLong.MakeCopy
        Case ExitType(eGDExitType_ShortStop)
            Set tsoSource = m.tsoStopShort.MakeCopy
    End Select
    
    Select Case nDestination
        Case ExitType(eGDExitType_LongProfit)
            Set m.tsoProfitLong = tsoSource.MakeCopy
            m.tsoProfitLong.Buy = False
            rtbTsLongProfit.TextRTF = m.tsoProfitLong.PreviewRTF
            
        Case ExitType(eGDExitType_LongStop)
            Set m.tsoStopLong = tsoSource.MakeCopy
            m.tsoStopLong.Buy = False
            rtbTsLongStop.TextRTF = m.tsoStopLong.PreviewRTF
            
        Case ExitType(eGDExitType_ShortProfit)
            Set m.tsoProfitShort = tsoSource.MakeCopy
            m.tsoProfitShort.Buy = True
            rtbTsShortProfit.TextRTF = m.tsoProfitShort.PreviewRTF
            
        Case ExitType(eGDExitType_ShortStop)
            Set m.tsoStopShort = tsoSource.MakeCopy
            m.tsoStopShort.Buy = True
            rtbTsShortStop.TextRTF = m.tsoStopShort.PreviewRTF
    
    End Select
    
    EditTradeSenseOrder nDestination
    
    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.CopyTradeSenseOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearTradeSenseOrder
'' Description: Clear the Trade Sense order
'' Inputs:      Exit Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearTradeSenseOrder(ByVal nExitType As eGDExitType)
On Error GoTo ErrSection:

    Select Case nExitType
        Case ExitType(eGDExitType_LongProfit)
            Set m.tsoProfitLong = New cTradeSenseOrder
            rtbTsLongProfit.TextRTF = ""
            
        Case ExitType(eGDExitType_LongStop)
            Set m.tsoStopLong = New cTradeSenseOrder
            rtbTsLongStop.TextRTF = ""
            
        Case ExitType(eGDExitType_ShortProfit)
            Set m.tsoProfitShort = New cTradeSenseOrder
            rtbTsShortProfit.TextRTF = ""
            
        Case ExitType(eGDExitType_ShortStop)
            Set m.tsoStopShort = New cTradeSenseOrder
            rtbTsShortStop.TextRTF = ""
    
    End Select
    
    Dirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExitStrategy.ClearTradeSenseOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ProfitSelected
'' Description: Has the user selected a profit target?
'' Inputs:      None
'' Returns:     True if profit selected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ProfitSelected() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If optStandardProfit.Value = True Then
        bReturn = (chkProfitTarget1.Value = vbChecked) Or (chkProfitTarget2.Value = vbChecked) Or (chkProfitTarget3.Value = vbChecked)
    Else
        bReturn = (Len(rtbTsLongProfit.Text) > 0) Or (Len(rtbTsShortProfit.Text) > 0)
    End If
    
    ProfitSelected = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExitStrategy.ProfitSelected"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StopSelected
'' Description: Has the user selected a stop loss?
'' Inputs:      None
'' Returns:     True if stop selected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function StopSelected() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If optNoStop.Value = True Then
        bReturn = False
    ElseIf optTsStop.Value = True Then
        bReturn = (Len(rtbTsLongStop.Text) > 0) Or (Len(rtbTsShortStop.Text) > 0)
    Else
        bReturn = True
    End If
    
    StopSelected = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExitStrategy.StopSelected"
    
End Function


