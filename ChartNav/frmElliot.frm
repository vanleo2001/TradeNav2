VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmElliot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elliott Wave palette (courtesy of EWI)"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraLabelOptions 
      Height          =   7215
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   6660
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
      Caption         =   "frmElliot.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmElliot.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmElliot.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraBottom 
         Height          =   1110
         Left            =   45
         TabIndex        =   86
         Top             =   6000
         Width           =   5190
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
         Caption         =   "frmElliot.frx":005C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmElliot.frx":007C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmElliot.frx":009C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkPreIndicator 
            Height          =   220
            Left            =   75
            TabIndex        =   89
            Top             =   0
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frmElliot.frx":00B8
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":0104
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":0124
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkMultiChart 
            Height          =   220
            Left            =   75
            TabIndex        =   88
            Top             =   275
            Width           =   2025
            _ExtentX        =   3572
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
            Caption         =   "frmElliot.frx":0140
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":0184
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":01A4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSaveDefaults 
            Height          =   450
            Left            =   300
            TabIndex        =   87
            Top             =   600
            Width           =   1515
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
            Caption         =   "frmElliot.frx":01C0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":0202
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":0222
            RightToLeft     =   0   'False
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1125
            Left            =   2220
            Picture         =   "frmElliot.frx":023E
            Top             =   0
            Width           =   2985
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraPresetStyles 
         Height          =   4950
         Left            =   45
         TabIndex        =   44
         Top             =   908
         Width           =   5190
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
         Caption         =   "frmElliot.frx":1179
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmElliot.frx":1199
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmElliot.frx":11B9
         RightToLeft     =   0   'False
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   3885
            Picture         =   "frmElliot.frx":11D5
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   48
            Top             =   1238
            Width           =   255
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   4170
            Picture         =   "frmElliot.frx":1617
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   0
            Top             =   2273
            Width           =   255
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3885
            Picture         =   "frmElliot.frx":1761
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   65
            Top             =   2273
            Width           =   255
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2505
            Picture         =   "frmElliot.frx":18AB
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   66
            Top             =   1583
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2220
            Picture         =   "frmElliot.frx":19F5
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   67
            Top             =   1583
            Width           =   255
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   765
            Picture         =   "frmElliot.frx":1B3F
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   68
            Top             =   1928
            Width           =   255
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   495
            Picture         =   "frmElliot.frx":1C89
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   93
            Top             =   1928
            Width           =   255
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   4170
            Picture         =   "frmElliot.frx":1DD3
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   96
            Top             =   1238
            Width           =   255
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3885
            Picture         =   "frmElliot.frx":1F1D
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   101
            Top             =   1238
            Width           =   255
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   220
            Index           =   12
            Left            =   240
            TabIndex        =   91
            Top             =   2730
            Width           =   1320
            _ExtentX        =   2328
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
            Caption         =   "frmElliot.frx":2067
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":209D
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":20BD
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraText 
            Height          =   2085
            Left            =   90
            TabIndex        =   85
            Top             =   2730
            Width           =   4980
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
            Caption         =   "frmElliot.frx":20D9
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmElliot.frx":20F9
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2119
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkBorder 
               Height          =   315
               Left            =   3840
               TabIndex        =   100
               Top             =   1705
               Width           =   960
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
               Caption         =   "frmElliot.frx":2135
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmElliot.frx":2161
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":2181
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optJustify 
               Height          =   220
               Index           =   2
               Left            =   3840
               TabIndex        =   99
               Top             =   1380
               Width           =   960
               _ExtentX        =   1693
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
               Caption         =   "frmElliot.frx":219D
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmElliot.frx":21C9
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":21E9
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optJustify 
               Height          =   220
               Index           =   1
               Left            =   3840
               TabIndex        =   98
               Top             =   1125
               Width           =   735
               _ExtentX        =   1296
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
               Caption         =   "frmElliot.frx":2205
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmElliot.frx":222F
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":224F
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optJustify 
               Height          =   220
               Index           =   0
               Left            =   3840
               TabIndex        =   97
               Top             =   870
               Width           =   735
               _ExtentX        =   1296
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
               Caption         =   "frmElliot.frx":226B
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmElliot.frx":2293
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":22B3
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkCopyright 
               Height          =   315
               Left            =   3240
               TabIndex        =   95
               Top             =   270
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
               Caption         =   "frmElliot.frx":22CF
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmElliot.frx":2311
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":2331
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniComboBoxXP cboYear 
               Height          =   315
               Left            =   1920
               TabIndex        =   94
               Top             =   270
               Width           =   1215
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
               Tip             =   "frmElliot.frx":234D
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
               MouseIcon       =   "frmElliot.frx":236D
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               ManualStart     =   0   'False
               MaxLength       =   0
               RightToLeft     =   0   'False
               LeftMargin      =   0
               RightMargin     =   0
               SelectOnFocus   =   0   'False
            End
            Begin HexUniControls.ctlUniComboBoxXP cboMonth 
               Height          =   315
               Left            =   480
               TabIndex        =   92
               Top             =   270
               Width           =   1335
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
               Tip             =   "frmElliot.frx":2389
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
               MouseIcon       =   "frmElliot.frx":23A9
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               ManualStart     =   0   'False
               MaxLength       =   0
               RightToLeft     =   0   'False
               LeftMargin      =   0
               RightMargin     =   0
               SelectOnFocus   =   0   'False
            End
            Begin HexUniControls.ctlUniRichTextBoxXP rtfText 
               Height          =   1275
               Left            =   120
               TabIndex        =   90
               Top             =   720
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   2249
               BackColor       =   -2147483643
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmElliot.frx":23C5
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
               ScrollBars      =   2
               PasswordChar    =   ""
               TrapTab         =   0   'False
               RaiseChangeEvent=   -1  'True
               RaiseUpdateEvent=   0   'False
               RaiseSelChangeEvent=   -1  'True
               Tip             =   "frmElliot.frx":23E5
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":2405
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
            Begin HexUniControls.ctlUniLabelXP Label1 
               Height          =   255
               Left            =   3840
               Top             =   660
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
               Caption         =   "frmElliot.frx":2421
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmElliot.frx":2451
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":2471
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblCopyright 
               Height          =   315
               Left            =   120
               Top             =   270
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
               Caption         =   "frmElliot.frx":248D
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmElliot.frx":24AF
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":24CF
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   11
            Left            =   4560
            TabIndex        =   81
            Top             =   2258
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":24EB
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
            Tip             =   "frmElliot.frx":250D
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":252D
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   10
            Left            =   4560
            TabIndex        =   80
            Top             =   1913
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":2549
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
            Tip             =   "frmElliot.frx":256B
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":258B
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   9
            Left            =   4560
            TabIndex        =   79
            Top             =   1568
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":25A7
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
            Tip             =   "frmElliot.frx":25C9
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":25E9
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   8
            Left            =   4560
            TabIndex        =   78
            Top             =   1223
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":2605
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
            Tip             =   "frmElliot.frx":2627
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2647
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   7
            Left            =   2895
            TabIndex        =   77
            Top             =   2258
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":2663
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
            Tip             =   "frmElliot.frx":2685
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":26A5
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   6
            Left            =   2895
            TabIndex        =   76
            Top             =   1913
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":26C1
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
            Tip             =   "frmElliot.frx":26E3
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2703
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   5
            Left            =   2895
            TabIndex        =   75
            Top             =   1568
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":271F
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
            Tip             =   "frmElliot.frx":2741
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2761
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   4
            Left            =   2895
            TabIndex        =   74
            Top             =   1223
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":277D
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
            Tip             =   "frmElliot.frx":279F
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":27BF
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   3
            Left            =   1170
            TabIndex        =   73
            Top             =   2258
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":27DB
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
            Tip             =   "frmElliot.frx":27FD
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":281D
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   2
            Left            =   1170
            TabIndex        =   72
            Top             =   1913
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":2839
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
            Tip             =   "frmElliot.frx":285B
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":287B
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   1
            Left            =   1170
            TabIndex        =   71
            Top             =   1568
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":2897
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
            Tip             =   "frmElliot.frx":28B9
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":28D9
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCustomFont 
            Height          =   285
            Index           =   0
            Left            =   1170
            TabIndex        =   70
            Top             =   1223
            Width           =   355
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmElliot.frx":28F5
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
            Tip             =   "frmElliot.frx":2917
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2937
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   10
            Left            =   3630
            TabIndex        =   60
            Top             =   1875
            Width           =   990
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
            Caption         =   "frmElliot.frx":2953
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   255
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":2981
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":29A1
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   0
            Left            =   240
            TabIndex        =   59
            Top             =   1185
            Width           =   990
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
            Caption         =   "frmElliot.frx":29BD
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   8388736
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":29E7
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2A07
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   1
            Left            =   240
            TabIndex        =   58
            Top             =   1530
            Width           =   990
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
            Caption         =   "frmElliot.frx":2A23
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   32768
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":2A51
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2A71
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   6
            Left            =   1965
            TabIndex        =   56
            Top             =   1875
            Width           =   990
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
            Caption         =   "frmElliot.frx":2A8D
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   255
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":2AB7
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2AD7
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   9
            Left            =   3630
            TabIndex        =   55
            Top             =   1530
            Width           =   990
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
            Caption         =   "frmElliot.frx":2AF3
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   32768
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":2B1F
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2B3F
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   3
            Left            =   240
            TabIndex        =   54
            Top             =   2220
            Width           =   990
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
            Caption         =   "frmElliot.frx":2B5B
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   16711680
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":2B85
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":2BA5
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   4170
            Picture         =   "frmElliot.frx":2BC1
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   53
            Top             =   2273
            Width           =   255
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   3885
            Picture         =   "frmElliot.frx":3003
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   52
            Top             =   2273
            Width           =   255
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2505
            Picture         =   "frmElliot.frx":3445
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   51
            Top             =   1583
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2220
            Picture         =   "frmElliot.frx":3887
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   50
            Top             =   1583
            Width           =   255
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   4170
            Picture         =   "frmElliot.frx":3CC9
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   49
            Top             =   1238
            Width           =   255
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   765
            Picture         =   "frmElliot.frx":410B
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   47
            Top             =   1928
            Width           =   255
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   495
            Picture         =   "frmElliot.frx":454D
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   46
            Top             =   1928
            Width           =   255
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   7
            Left            =   1965
            TabIndex        =   45
            Top             =   2220
            Width           =   990
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
            Caption         =   "frmElliot.frx":498F
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   16711680
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":49BD
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":49DD
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   8
            Left            =   3630
            TabIndex        =   63
            Top             =   1185
            Width           =   990
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
            Caption         =   "frmElliot.frx":49F9
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":4A35
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":4A55
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   11
            Left            =   3630
            TabIndex        =   64
            Top             =   2220
            Width           =   990
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
            Caption         =   "frmElliot.frx":4A71
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":4AAD
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":4ACD
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   5
            Left            =   1965
            TabIndex        =   62
            Top             =   1530
            Width           =   990
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
            Caption         =   "frmElliot.frx":4AE9
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":4B25
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":4B45
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   2
            Left            =   240
            TabIndex        =   61
            Top             =   1875
            Width           =   990
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
            Caption         =   "frmElliot.frx":4B61
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":4B9D
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":4BBD
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optStylePreset 
            Height          =   360
            Index           =   4
            Left            =   1965
            TabIndex        =   57
            Top             =   1185
            Width           =   990
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
            Caption         =   "frmElliot.frx":4BD9
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   8388736
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":4C07
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":4C27
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraCustomFont 
            Height          =   285
            Left            =   600
            TabIndex        =   82
            Top             =   585
            Width           =   3855
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
            Caption         =   "frmElliot.frx":4C43
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmElliot.frx":4C63
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":4C83
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optFontPreset 
               Height          =   255
               Left            =   240
               TabIndex        =   84
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
               Caption         =   "frmElliot.frx":4C9F
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmElliot.frx":4CE1
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":4D01
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optFontCustom 
               Height          =   255
               Left            =   2025
               TabIndex        =   83
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
               Caption         =   "frmElliot.frx":4D1D
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmElliot.frx":4D5F
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmElliot.frx":4D7F
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniLabelXP lblInfoPreset 
            Height          =   435
            Left            =   510
            Top             =   120
            Width           =   4095
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
            Caption         =   "frmElliot.frx":4D9B
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmElliot.frx":4E97
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":4EB7
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPresetSize14 
            Height          =   210
            Left            =   3630
            Top             =   930
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
            Caption         =   "frmElliot.frx":4ED3
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmElliot.frx":4F0B
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":4F2B
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPresetSize12 
            Height          =   210
            Left            =   1965
            Top             =   930
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
            Caption         =   "frmElliot.frx":4F47
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmElliot.frx":4F7F
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":4F9F
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPresetSize10 
            Height          =   210
            Left            =   240
            Top             =   930
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
            Caption         =   "frmElliot.frx":4FBB
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmElliot.frx":4FF3
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":5013
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdFont 
         Height          =   345
         Left            =   4440
         TabIndex        =   43
         Top             =   45
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
         Caption         =   "frmElliot.frx":502F
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmElliot.frx":5057
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmElliot.frx":5077
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraAlphaCbo 
         Height          =   390
         Left            =   0
         TabIndex        =   22
         Top             =   368
         Width           =   570
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
         Caption         =   "frmElliot.frx":5093
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmElliot.frx":50B3
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmElliot.frx":50D3
         RightToLeft     =   0   'False
         Begin VB.PictureBox picAlphaCbo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            Picture         =   "frmElliot.frx":50EF
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   24
            Top             =   120
            Width           =   255
         End
         Begin VB.PictureBox picAlphaCboArrow 
            BorderStyle     =   0  'None
            FillColor       =   &H80000012&
            ForeColor       =   &H80000013&
            Height          =   60
            Left            =   345
            Picture         =   "frmElliot.frx":5531
            ScaleHeight     =   60
            ScaleWidth      =   120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   210
            Width           =   120
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdAlphaCbo 
            Height          =   255
            Left            =   285
            TabIndex        =   25
            Top             =   120
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
            Caption         =   "frmElliot.frx":583B
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":5871
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":5891
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraAlphaDrop 
         Height          =   1710
         Left            =   6000
         TabIndex        =   36
         Top             =   675
         Visible         =   0   'False
         Width           =   540
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
         Caption         =   "frmElliot.frx":58AD
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Tip             =   "frmElliot.frx":58CD
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmElliot.frx":58ED
         RightToLeft     =   0   'False
         Begin VB.PictureBox picAlphaDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   75
            Picture         =   "frmElliot.frx":5909
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   42
            Top             =   1140
            Width           =   420
         End
         Begin VB.PictureBox picAlphaDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   75
            Picture         =   "frmElliot.frx":5D4B
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   41
            Top             =   120
            Width           =   420
         End
         Begin VB.PictureBox picAlphaDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   75
            Picture         =   "frmElliot.frx":618D
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   40
            Top             =   1410
            Width           =   420
         End
         Begin VB.PictureBox picAlphaDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   75
            Picture         =   "frmElliot.frx":65CF
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   39
            Top             =   360
            Width           =   420
         End
         Begin VB.PictureBox picAlphaDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   75
            Picture         =   "frmElliot.frx":6A11
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   38
            Top             =   630
            Width           =   420
         End
         Begin VB.PictureBox picAlphaDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   75
            Picture         =   "frmElliot.frx":6E53
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   37
            Top             =   885
            Width           =   420
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraNumDrop 
         Height          =   2520
         Left            =   5640
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   540
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
         Caption         =   "frmElliot.frx":7295
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Tip             =   "frmElliot.frx":72B5
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmElliot.frx":72D5
         RightToLeft     =   0   'False
         Begin VB.PictureBox picNumDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   75
            Picture         =   "frmElliot.frx":72F1
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   35
            Top             =   897
            Width           =   420
         End
         Begin VB.PictureBox picNumDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   75
            Picture         =   "frmElliot.frx":7733
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   34
            Top             =   633
            Width           =   420
         End
         Begin VB.PictureBox picNumDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   75
            Picture         =   "frmElliot.frx":7B75
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   33
            Top             =   369
            Width           =   420
         End
         Begin VB.PictureBox picNumDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   75
            Picture         =   "frmElliot.frx":7FB7
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   32
            Top             =   1425
            Width           =   420
         End
         Begin VB.PictureBox picNumDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   75
            Picture         =   "frmElliot.frx":83F9
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   31
            Top             =   105
            Width           =   420
         End
         Begin VB.PictureBox picNumDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   75
            Picture         =   "frmElliot.frx":883B
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   30
            Top             =   1161
            Width           =   420
         End
         Begin VB.PictureBox picNumDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   75
            Picture         =   "frmElliot.frx":8C7D
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   29
            Top             =   1689
            Width           =   420
         End
         Begin VB.PictureBox picNumDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   75
            Picture         =   "frmElliot.frx":90BF
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   28
            Top             =   1953
            Width           =   420
         End
         Begin VB.PictureBox picNumDrop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   75
            Picture         =   "frmElliot.frx":9501
            ScaleHeight     =   255
            ScaleWidth      =   420
            TabIndex        =   27
            Top             =   2220
            Width           =   420
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraNumCbo 
         Height          =   390
         Left            =   0
         TabIndex        =   18
         Top             =   13
         Width           =   570
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
         Caption         =   "frmElliot.frx":9943
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmElliot.frx":9963
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmElliot.frx":9983
         RightToLeft     =   0   'False
         Begin VB.PictureBox picNumCboArrow 
            BorderStyle     =   0  'None
            FillColor       =   &H80000012&
            ForeColor       =   &H80000013&
            Height          =   60
            Left            =   345
            Picture         =   "frmElliot.frx":999F
            ScaleHeight     =   60
            ScaleWidth      =   120
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   210
            Width           =   120
         End
         Begin VB.PictureBox picNumCbo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            Picture         =   "frmElliot.frx":9CA9
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   19
            Top             =   120
            Width           =   255
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdNumCbo 
            Height          =   255
            Left            =   285
            TabIndex        =   21
            Top             =   120
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
            Caption         =   "frmElliot.frx":A0EB
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmElliot.frx":A11D
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmElliot.frx":A13D
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraNumButtons 
         Height          =   465
         Left            =   568
         TabIndex        =   2
         Top             =   0
         Width           =   2000
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
         Caption         =   "frmElliot.frx":A159
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmElliot.frx":A179
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmElliot.frx":A199
         RightToLeft     =   0   'False
         Begin VB.PictureBox picNumButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   0
            Left            =   15
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   7
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picNumButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   1
            Left            =   400
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   6
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picNumButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   2
            Left            =   790
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   5
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picNumButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   3
            Left            =   1190
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   4
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picNumButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   4
            Left            =   1580
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   3
            Top             =   105
            Width           =   375
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraAlphaButtons 
         Height          =   465
         Left            =   568
         TabIndex        =   8
         Top             =   358
         Width           =   3565
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
         Caption         =   "frmElliot.frx":A1B5
         Enabled         =   0   'False
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmElliot.frx":A1D5
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmElliot.frx":A1F5
         RightToLeft     =   0   'False
         Begin VB.PictureBox picAlphaButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   0
            Left            =   15
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   17
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picAlphaButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   1
            Left            =   400
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   16
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picAlphaButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   2
            Left            =   790
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   15
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picAlphaButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   3
            Left            =   1190
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   14
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picAlphaButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   4
            Left            =   1580
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   13
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picAlphaButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   5
            Left            =   1970
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   12
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picAlphaButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   6
            Left            =   2360
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   11
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picAlphaButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   7
            Left            =   2750
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   10
            Top             =   105
            Width           =   375
         End
         Begin VB.PictureBox picAlphaButton 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   325
            Index           =   8
            Left            =   3150
            ScaleHeight     =   330
            ScaleWidth      =   375
            TabIndex        =   9
            Top             =   105
            Width           =   375
         End
      End
      Begin gdOCX.gdSelectColor gdColor 
         Height          =   345
         Left            =   4440
         TabIndex        =   69
         Top             =   443
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
         CustomColor     =   255
      End
   End
   Begin MSComctlLib.ImageList ImgListAlpha 
      Left            =   6180
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":A211
            Key             =   "kLowerA"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":A52B
            Key             =   "kLowerB"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":A845
            Key             =   "kLowerC"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":AB5F
            Key             =   "kLowerD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":AE79
            Key             =   "kLowerE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":B193
            Key             =   "kLowerX"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":B4AD
            Key             =   "kLowerW"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":B7C7
            Key             =   "kLowerY"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":BAE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":BDFB
            Key             =   "kCapA"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":C115
            Key             =   "kCapB"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":C42F
            Key             =   "kCapC"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":C749
            Key             =   "kCapD"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":CA63
            Key             =   "kCapE"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":CD7D
            Key             =   "kCapX"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":D097
            Key             =   "kCapW"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":D3B1
            Key             =   "kCapY"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":D6CB
            Key             =   "kCapZ"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListNum 
      Left            =   5490
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":D9E5
            Key             =   "kOne"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":DCFF
            Key             =   "kTwo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":E019
            Key             =   "kThree"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":E333
            Key             =   "kFour"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":E64D
            Key             =   "kFive"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":E967
            Key             =   "kOneCapR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":EC81
            Key             =   "kTwoCapR"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":EF9B
            Key             =   "kThreeCapR"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":F2B5
            Key             =   "kFourCapR"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":F5CF
            Key             =   "kFiveCapR"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":F8E9
            Key             =   "kOneR"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":FC03
            Key             =   "kTwoR"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":FF1D
            Key             =   "kThreeR"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":10237
            Key             =   "kFourR"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmElliot.frx":10551
            Key             =   "kFiveR"
         EndProperty
      EndProperty
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   4755
      Top             =   105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   9
      DisplayContextMenu=   0   'False
      Tools           =   "frmElliot.frx":1086B
      ToolBars        =   "frmElliot.frx":161F7
   End
   Begin HexUniControls.ctlUniLabelXP lblInfoPrompt 
      Height          =   435
      Left            =   1170
      Top             =   37
      Width           =   2835
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
      Caption         =   "frmElliot.frx":1642E
      BackColor       =   -2147483633
      ForeColor       =   -2147483635
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmElliot.frx":164E6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmElliot.frx":16506
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmElliot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kPurple = 12583104
Private Const kLightGreen = 49152
Private Const kDarkGreen = 32768                  'RGB(0,128,0)      - forest green (aardvark 6897)
Private Const kHeightNew = 4200
Private Const kHeightEdit = 1815

'colors for end user palette
Private Const kColorPink = 16711935
Private Const kColorOlive = 32896
Private Const kColorOrange = 33023
Private Const kColorCyan = 12632064

Private Enum eNumStyle
    eNumStyle_Normal = 0
    eNumStyle_Normal_Paren = 1
    eNumStyle_Normal_Circle = 2
    eNumStyle_LowerR = 3
    eNumStyle_LowerR_Paren = 4
    eNumStyle_LowerR_Circle = 5
    eNumStyle_UpperR = 6
    eNumStyle_UpperR_Paren = 7
    eNumStyle_UpperR_Circle = 8
End Enum

Private Enum eAlphaStyle
    eAlphaStyle_Lower = 0
    eAlphaStyle_Lower_Paren = 1
    eAlphaStyle_Lower_Circle = 2
    eAlphaStyle_Upper = 3
    eAlphaStyle_Upper_Paren = 4
    eAlphaStyle_Upper_Circle = 5
End Enum

Private Type mPrivate
    Annot As cAnnotation
    AnnotDefaults As cAnnotation
    
    eNumberStyle As eNumStyle
    eCharStyle As eAlphaStyle
    eCommon As eCommonStyle
    eImage As eStockImage
    
    strLabel As String
    strFont As String
    nFontSize As Long
    nItalic As Long
    nBold As Long
    
    nNumberBorderIdx As Long        'index of number buttons
    nAlphaBorderIdx As Long
    nRepeatStateSave As Long
    
    bEditExisting As Boolean
    bWasMultiChart As Boolean
    bInitInprog As Boolean
    bEndUserPallette As Boolean
End Type

Private m As mPrivate

Public Sub ShowMe(Annot As cAnnotation, Optional ByVal bEndUserPallette As Boolean = False)
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    
    Dim i&, j&, idx&
    Dim eMode As eShowFormMode
    Dim ThisChart As cChart
    Dim bShowDelete As Boolean
    Dim bGMP As Boolean
    Dim bHide As Boolean
    
    Dim picboxIdx As Integer
    
    If bInProgress Then Exit Sub
    
    bInProgress = True
        
    m.nNumberBorderIdx = -1
    m.nAlphaBorderIdx = -1
    m.eCommon = eCommon_None
    m.nRepeatStateSave = -1
        
    ClearBorder
    
    If Annot Is Nothing Then
        m.bEditExisting = False
        m.bEndUserPallette = bEndUserPallette
        Set m.Annot = New cAnnotation
        m.Annot.CreateNew ActiveChart().Chart, eANNOT_ElliotLabel, 1, 0, 0, 0, 0, , , , , True
        
        InitFromAnnot
        bGMP = m.Annot.AllowGMP
        
        Set m.AnnotDefaults = m.Annot
        Set m.Annot = Nothing
        
        Set ThisChart = ActiveChart().Chart
        
        eMode = eForm_Nonmodal
        lblInfoPrompt.Visible = True
        
        If frmMain.tbToolbar.Tools("ID_RepeatDraw").State = ssUnchecked Then
            m.nRepeatStateSave = ssUnchecked
            frmMain.tbToolbar.Tools("ID_RepeatDraw").State = ssChecked
        End If
        
    Else
        m.bEditExisting = True
        
        Set m.Annot = Annot
        m.strLabel = m.Annot.Text
        
        InitFromAnnot
        bGMP = m.Annot.AllowGMP
        
        If Annot.AllowEWI Or Annot.AllowGMP Then
            m.bEndUserPallette = False
        ElseIf Annot.IsEndUserEWI And HasModule("EWL") Then
            m.bEndUserPallette = True
        End If
        
        Set ThisChart = Annot.AnnotChart
        
        eMode = eForm_ActModal
        lblInfoPrompt.Visible = False
        Me.cmdSaveDefaults.Visible = False
        Me.cmdSaveDefaults.Visible = False
        
        picboxIdx = IndexForChar(m.Annot.Text)      '6915
        
        If picboxIdx < 0 Or picboxIdx > picAlphaButton.Count Then
            picboxIdx = IndexForNum(m.Annot.Text)
            If picboxIdx >= 0 And picboxIdx < picNumButton.Count Then
                picNumButton_Click picboxIdx
            End If
        ElseIf picboxIdx >= 0 And picboxIdx < picAlphaButton.Count Then
            picAlphaButton_Click picboxIdx
        End If
        
    End If
    
    ' TLB 12/23/2013: now Ewave no longer wants the custom colors, and also allow the custom color/font
    'If m.bEndUserPallette Then
    If 0 Then
        optStylePreset(0).ForeColor = kColorPink
        optStylePreset(4).ForeColor = kColorPink
        optStylePreset(1).ForeColor = kColorOlive
        optStylePreset(9).ForeColor = kColorOlive
        optStylePreset(6).ForeColor = kColorOrange
        optStylePreset(10).ForeColor = kColorOrange
        optStylePreset(3).ForeColor = kColorCyan
        optStylePreset(7).ForeColor = kColorCyan
        
        gdColor.Visible = False
        gdColor.Enabled = False
        cmdFont.Visible = False
        cmdFont.Enabled = False
        
        m.strFont = "Times New Roman"
    End If
    
    fraText.ZOrder
    optStylePreset(12).ZOrder
    
    For i = 0 To 1
        If i = 0 Then
            bHide = True ' Not m.bEndUserPallette
        Else
            bHide = False ' m.bEndUserPallette
        End If
        
        Picture1(i).Visible = bHide
        Picture1(i).Enabled = bHide
        Picture2(i).Visible = bHide
        Picture2(i).Enabled = bHide
        Picture3(i).Visible = bHide
        Picture3(i).Enabled = bHide
        Picture4(i).Visible = bHide
        Picture4(i).Enabled = bHide
        Picture5(i).Visible = bHide
        Picture5(i).Enabled = bHide
        Picture6(i).Visible = bHide
        Picture6(i).Enabled = bHide
        Picture7(i).Visible = bHide
        Picture7(i).Enabled = bHide
        Picture8(i).Visible = bHide
        Picture8(i).Enabled = bHide
        
        Picture1(i).BackColor = g.nColorTheme
        Picture2(i).BackColor = g.nColorTheme
        Picture3(i).BackColor = g.nColorTheme
        Picture4(i).BackColor = g.nColorTheme
        Picture5(i).BackColor = g.nColorTheme
        Picture6(i).BackColor = g.nColorTheme
        Picture7(i).BackColor = g.nColorTheme
        Picture8(i).BackColor = g.nColorTheme
    Next

    Me.Width = 5490
    
    With fraLabelOptions
'RH commented out         .BorderStyle = 0
        .Width = Me.Width
        If eMode = eForm_ActModal Then
            .Move .Left, lblInfoPrompt.Top
            If bGMP Then
                'editing annot with gmp.flg file present
                If Len(m.Annot.Text) > 5 Then       '6903
                    'editing copyright label
                    'hide the preset style radio buttons & show custom text frame
                    For i = 0 To 12
                        optStylePreset(i).Visible = False
                    Next
                
                    fraPresetStyles.Top = fraPresetStyles.Top + 45
                    fraPresetStyles.Height = fraText.Height + 180 ' 215
                    'RH commented out fraText.BorderStyle = 0
                    fraText.Top = 120 'lblInfoPreset.Top
                    fraBottom.Top = fraPresetStyles.Top + fraPresetStyles.Height + 150
                    'Me.Height = 4750
                    
                    Me.fraAlphaButtons.Enabled = False
                    Me.fraNumButtons.Enabled = False
                    Me.fraAlphaCbo.Enabled = False
                    Me.fraNumCbo.Enabled = False
                    
                    optStylePreset(12).Value = True
                Else
                    'editing not copyright label
                    'hide custom text frame and show preset style buttons
                    optStylePreset(12).Visible = False
                    fraText.Visible = False
                    fraText.Enabled = False
                    fraPresetStyles.Height = fraPresetStyles.Height - fraText.Height
                    fraBottom.Top = fraPresetStyles.Top + fraPresetStyles.Height + 150
                    'Me.Height = 5275
                    
                End If
            Else
                'editing annot without gmp.flg file present
                'Me.Height = 5065
                fraPresetStyles.Height = optStylePreset(3).Top + optStylePreset(3).Height + 90
                fraBottom.Top = fraPresetStyles.Top + fraPresetStyles.Height + 135
                
                optStylePreset(12).Visible = False
            End If
            
            bShowDelete = True
        ElseIf bGMP Then
            'adding new annot with gmp.flg file present
            'Me.Height = 7800
        Else
            'adding new annot without gmp.flg file present
            'Me.Height = 5555
            fraPresetStyles.Height = optStylePreset(3).Top + optStylePreset(3).Height + 90
            fraBottom.Top = fraPresetStyles.Top + fraPresetStyles.Height + 135
        End If
        
        Me.Height = fraBottom.Top + fraBottom.Height + (Me.Height - Me.ScaleHeight) + fraLabelOptions.Top + 45
    End With
    
    With fraText
        If m.bEndUserPallette Then
            .Visible = bGMP
            .Enabled = bGMP
        End If
    End With
    
    tbToolbar.Tools("ID_DeleteIcon").Visible = bShowDelete
    
    bShowDelete = Not bShowDelete
    tbToolbar.Tools("ID_ArrowLine").Visible = bShowDelete
    tbToolbar.Tools("ID_Trendline").Visible = bShowDelete
    tbToolbar.Tools("ID_TrendChannel").Visible = bShowDelete
    tbToolbar.Tools("ID_RegressionLine").Visible = bShowDelete
    tbToolbar.Tools("ID_Fibonacci").Visible = bShowDelete
    tbToolbar.Tools("ID_DNExpansion").Visible = bShowDelete
    
    CenterFormOnChart Me, ThisChart
    ToggleTopMost True
    ShowForm Me, eMode
        
    bInProgress = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.ShowMe"

End Sub

Private Sub cboMonth_Click()
On Error GoTo ErrSection:

    If Not m.bInitInprog Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                ToggleTopMost True
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If

    If Me.Visible Then
        SetIniFileProperty "EWYearMonth", Val(cboYear.Text) * 100 + cboMonth.ListIndex + 1, "", g.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.cboMonth_Click"

End Sub

Private Sub cboYear_Click()
On Error GoTo ErrSection

    If Not m.bInitInprog Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                ToggleTopMost True
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If

    If Me.Visible Then
        SetIniFileProperty "EWYearMonth", Val(cboYear.Text) * 100 + cboMonth.ListIndex + 1, "", g.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.cboYear_Click"

End Sub

Private Sub chkBorder_Click()
On Error GoTo ErrSection

    If Not m.bInitInprog Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                ToggleTopMost True
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.chkBorder_Click"

End Sub

Private Sub chkCopyright_Click()
On Error GoTo ErrSection

    If Not m.bInitInprog Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                ToggleTopMost True
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.chkCopyright_Click"

End Sub

Private Sub chkPreIndicator_Click()
On Error GoTo ErrSection:

    If Not m.bInitInprog Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmElliot.chkPreIndicator_Click"

End Sub

Private Sub cmdAlphaCbo_Click()
On Error GoTo ErrSection:

    Dim bVisible As Boolean
    
    If fraNumDrop.Visible Then fraNumDrop.Visible = False
    
    bVisible = Not fraAlphaDrop.Visible
    
    If bVisible Then
        fraAlphaDrop.ZOrder
        fraAlphaCbo.ZOrder
    Else
        fraAlphaDrop.ZOrder 1
    End If
    
    fraAlphaDrop.Visible = bVisible
    
    Exit Sub
            
ErrSection:
    RaiseError "frmElliot.cmdAlphaCbo_Click"

End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:

    If fraAlphaDrop.Visible Then fraAlphaDrop.Visible = False
    If fraNumDrop.Visible Then fraNumDrop.Visible = False
    
    RestoreDefaultFontInfo
    
    Me.Font.Name = m.strFont
    Me.Font.Size = m.nFontSize
    Me.FontItalic = m.nItalic * -1
    Me.Font.Bold = m.nBold * -1

    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        m.strFont = Me.Font.Name
        m.nFontSize = Me.Font.Size
        m.nItalic = Abs(Me.FontItalic)
        m.nBold = Abs(Me.Font.Bold)
    End If

    If Not m.Annot Is Nothing Then
        If Not m.Annot.AnnotChart Is Nothing Then
            GetSettings m.Annot
            m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
        End If
    End If
    
    If fraText.Visible Then
        If Me.optStylePreset(eCommon_CustomText).Value = True Then
            rtfText.Font.Name = m.strFont
            rtfText.Font.Size = m.nFontSize
            rtfText.Font.Bold = m.nBold
            rtfText.Font.Italic = m.nItalic
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmElliot.cmdFont_Click"

End Sub

Private Sub cmdNumCbo_Click()
On Error GoTo ErrSection:

    Dim bVisible As Boolean
    
    If fraAlphaDrop.Visible Then fraAlphaDrop.Visible = False
    
    bVisible = Not fraNumDrop.Visible
    
    If bVisible Then
        fraNumDrop.ZOrder
        fraNumCbo.ZOrder
    Else
        fraNumDrop.ZOrder 1
    End If
    
    fraNumDrop.Visible = bVisible
    
    Exit Sub

ErrSection:
    RaiseError "frmElliot.cmdNumCbo_Click"

End Sub

Private Sub cmdSaveDefaults_Click()
On Error GoTo ErrSection:

    Dim bTemp As Boolean
    
    If m.Annot Is Nothing Then
        bTemp = True
        Set m.Annot = New cAnnotation
        m.Annot.CreateNew ActiveChart().Chart, eANNOT_ElliotLabel, 1, 0, 0, 0, 0, , , , , True
    End If
    
    GetSettings m.Annot
    
    m.Annot.Text = ""
    
    If fraText.Visible And optStylePreset(eCommon_CustomText).Value = True Then
        'omit copyright line as this is a separate property
        m.Annot.Prop("CustomText") = rtfText.Text
        m.Annot.Text = "FakeText"
    End If
    
    m.Annot.SaveDefaults
    
    m.Annot.Text = ""
    
    Set m.AnnotDefaults = m.Annot
    
    If bTemp Then Set m.Annot = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "frmElliot.cmdSaveDefaults_Click"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText$
    
    g.Styler.StyleForm Me

    fraAlphaDrop.Move fraAlphaCbo.Left, fraAlphaCbo.Top + fraAlphaCbo.Height - 80
    fraNumDrop.Move fraNumCbo.Left, fraNumCbo.Top + fraNumCbo.Height - 90
    
    With tbToolbar
        ' if FIB then give tool as "DiNapoli Expansion", else if has Gold then
        ' give tool as "Fibonacci Extension", else don't give tool at all
        If HasModule("FIB") Then
            .Tools("ID_DNExpansion").Visible = True
        ElseIf HasGold(False) Then
            .Tools("ID_DNExpansion").Visible = True
            strText = "Fibonacci Extension"
            If .Tools("ID_DNExpansion").Name <> strText Then
                .Tools("ID_DNExpansion").ToolTipText = strText
                .Tools("ID_DNExpansion").ChangeAll ssChangeAllName, strText
            End If
        Else
            .Tools("ID_DNExpansion").Visible = False
        End If
    End With
    
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim Chart As cChart
    
    If g.bUnloading Then Exit Sub
    
    If Not m.Annot Is Nothing Then
        Set Chart = m.Annot.AnnotChart
    ElseIf Not ActiveChart() Is Nothing Then
        Set Chart = ActiveChart().Chart
    End If
    
    If ssUnchecked = m.nRepeatStateSave Then
        frmMain.tbToolbar.Tools("ID_RepeatDraw").State = ssUnchecked
    End If
    
    Me.Hide     '4937
    
    If Not Chart Is Nothing Then
        If Not m.Annot Is Nothing Then Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart
        Chart.SetCursor
        If IsEWIActiveDraw() Or g.strActiveDraw = "" Then
            ToolbarSetCursorGroup Chart.tbToolbar, False
            If Not Chart.Form Is Nothing Then Chart.Form.SyncDrawTools      '5141
        Else
            'user clicked on another drawtool
            ToolbarSetCursorGroup Chart.tbToolbar, True
        End If
    End If
        
    Set m.Annot = Nothing
    m.strLabel = ""
    
    Exit Sub

ErrSection:
    RaiseError "frmElliot.Form_Unload"

End Sub

Private Sub gdColor_Changed()
On Error GoTo ErrSection:

    If Not m.bInitInprog Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                ToggleTopMost True
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmElliot.gdColor_Changed"

End Sub

Private Sub gdColor_GotFocus()
On Error Resume Next
    
    ToggleTopMost False
    If fraAlphaDrop.Visible Then fraAlphaDrop.Visible = False
    If fraNumDrop.Visible Then fraNumDrop.Visible = False

End Sub

Private Sub optFontCustom_Click()
On Error GoTo ErrSection:

    AdjustPresetLabels 1

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmElliot.optFontCustom_Click"

End Sub

Private Sub optFontPreset_Click()
On Error GoTo ErrSection:

    AdjustPresetLabels 0

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmElliot.optFontPreset_Click"

End Sub

Private Sub optJustify_Click(Index As Integer)
On Error GoTo ErrSection

    If Not m.bInitInprog Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                ToggleTopMost True
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmElliot.optJustify_Click"

End Sub

Private Sub optStylePreset_Click(Index As Integer)
On Error GoTo ErrSection:

    RestoreDefaultFontInfo
    
    Select Case Index
        Case eCommon_Miniscule
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorPink
            'Else
                gdColor.Color = kPurple
            'End If
            picAlphaDrop_Click eAlphaStyle_Upper
            picNumDrop_Click eNumStyle_Normal
            m.eCommon = eCommon_Miniscule
            m.nFontSize = 10
        
        Case eCommon_Submicro
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorOlive
            'Else
                gdColor.Color = kDarkGreen
            'End If
            picAlphaDrop_Click eAlphaStyle_Upper_Paren
            picNumDrop_Click eNumStyle_Normal_Paren
            m.eCommon = eCommon_Submicro
            m.nFontSize = 10
    
        Case eCommon_Micro
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorOrange
            'Else
                gdColor.Color = vbRed
            'End If
            picAlphaDrop_Click eAlphaStyle_Upper_Circle
            picNumDrop_Click eNumStyle_Normal_Circle
            m.eCommon = eCommon_Micro
            m.nFontSize = 10
        
        Case eCommon_Subminuette
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorCyan
            'Else
                gdColor.Color = vbBlue
            'End If
            picAlphaDrop_Click eAlphaStyle_Lower
            picNumDrop_Click eNumStyle_LowerR
            m.eCommon = eCommon_Subminuette
            m.nFontSize = 10
        
        Case eCommon_Minuette
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorPink
            'Else
                gdColor.Color = kPurple
            'End If
            picAlphaDrop_Click eAlphaStyle_Lower_Paren
            picNumDrop_Click eNumStyle_LowerR_Paren
            m.eCommon = eCommon_Minuette
            m.nFontSize = 12
        
        Case eCommon_Minute
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorOlive
            'Else
                gdColor.Color = kDarkGreen
            'End If
            picAlphaDrop_Click eAlphaStyle_Lower_Circle
            picNumDrop_Click eNumStyle_LowerR_Circle
            m.eCommon = eCommon_Minute
            m.nFontSize = 12
        
        Case eCommon_Minor
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorOrange
            'Else
                gdColor.Color = vbRed
            'End If
            picAlphaDrop_Click eAlphaStyle_Upper
            picNumDrop_Click eNumStyle_Normal
            m.eCommon = eCommon_Minor
            m.nFontSize = 12
        
        Case eCommon_Intermediate
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorCyan
            'Else
                gdColor.Color = vbBlue
            'End If
            picAlphaDrop_Click eAlphaStyle_Upper_Paren
            picNumDrop_Click eNumStyle_Normal_Paren
            m.eCommon = eCommon_Intermediate
            m.nFontSize = 12
        
        Case eCommon_Primary
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorPink
            'Else
                gdColor.Color = kPurple
            'End If
            picAlphaDrop_Click eAlphaStyle_Upper_Circle
            picNumDrop_Click eNumStyle_Normal_Circle
            m.eCommon = eCommon_Primary
            m.nFontSize = 14
        
        Case eCommon_Cycle
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorOlive
            'Else
                gdColor.Color = kDarkGreen
            'End If
            picAlphaDrop_Click eAlphaStyle_Lower
            picNumDrop_Click eNumStyle_UpperR
            m.eCommon = eCommon_Cycle
            m.nFontSize = 14
        
        Case eCommon_Supercycle
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorOrange
            'Else
                gdColor.Color = vbRed
            'End If
            picAlphaDrop_Click eAlphaStyle_Lower_Paren
            picNumDrop_Click eNumStyle_UpperR_Paren
            m.eCommon = eCommon_Supercycle
            m.nFontSize = 14
        
        Case eCommon_GrandSupercycle
            'If m.bEndUserPallette Then
            '    gdColor.Color = kColorCyan
            'Else
                gdColor.Color = vbBlue
            'End If
            picAlphaDrop_Click eAlphaStyle_Lower_Circle
            picNumDrop_Click eNumStyle_UpperR_Circle
            m.eCommon = eCommon_GrandSupercycle
            m.nFontSize = 14
        
        Case eCommon_CustomText
            If Not m.AnnotDefaults Is Nothing And Not m.bEditExisting Then
                gdColor.Color = m.AnnotDefaults.Color
            End If
            rtfText.Font.Name = m.strFont
            rtfText.Font.Size = m.nFontSize
            rtfText.Font.Bold = m.nBold
            rtfText.Font.Italic = m.nItalic
    
    End Select
    
    If m.bEndUserPallette Then m.strFont = "Times New Roman"
    
    If Index <> eCommon_CustomText Then
        If optFontCustom.Visible And optFontPreset.Visible Then
            If Not optFontCustom.Value And Not optFontPreset.Value Then
                optFontPreset.Value = True    'editor brought up to edit existing
            ElseIf optFontCustom.Visible And optFontCustom.Value Then
                m.nFontSize = ValOfText(txtCustomFont(Index).Text)      'have both flag files
            End If
        ElseIf txtCustomFont(0).Visible Then
            m.nFontSize = ValOfText(txtCustomFont(Index).Text)  'have only gmp.flg file
        End If
    End If

    If Not m.Annot Is Nothing Then
        If Not m.Annot.AnnotChart Is Nothing Then
            GetSettings m.Annot
            m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.picAlphaButton_Click"

End Sub

Private Sub picAlphaButton_Click(Index As Integer)
On Error GoTo ErrSection:

    ClearBorder
    picAlphaButton(Index).BorderStyle = 1
    
    SetAlphaString Index

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.picAlphaButton_Click"

End Sub

Private Sub picAlphaCbo_Click()
On Error Resume Next

    cmdAlphaCbo_Click
End Sub

Private Sub picAlphaCboArrow_Click()
    cmdAlphaCbo_Click
End Sub

Private Sub picAlphaDrop_Click(Index As Integer)
On Error GoTo ErrSection:

    ClearBorder
    picAlphaCbo.Picture = picAlphaDrop(Index).Picture
    fraAlphaDrop.Visible = False
    fraAlphaDrop.ZOrder 1
    
    Select Case Index
        Case 0
            AlphaButtonPrint "a", "b", "c", "d", "e", "x", "w", "y", "z"
            m.eCharStyle = eAlphaStyle_Lower
        Case 1
            AlphaButtonPrint "(a)", "(b)", "(c)", "(d)", "(e)", "(x)", "(w)", "(y)", "(z)"
            m.eCharStyle = eAlphaStyle_Lower_Paren
        Case 2
            AlphaButtonPic "kLowerA"
            m.eCharStyle = eAlphaStyle_Lower_Circle
        Case 3
            AlphaButtonPrint "A", "B", "C", "D", "E", "X", "W", "Y", "Z"
            m.eCharStyle = eAlphaStyle_Upper
        Case 4
            AlphaButtonPrint "(A)", "(B)", "(C)", "(D)", "(E)", "(X)", "(W)", "(Y)", "(Z)"
            m.eCharStyle = eAlphaStyle_Upper_Paren
        Case 5
            AlphaButtonPic "kCapA"
            m.eCharStyle = eAlphaStyle_Upper_Circle
    End Select
    
    If Len(m.strLabel) = 0 Then
        picAlphaButton(0).BorderStyle = 1
        SetAlphaString 0
        MoveFocus picAlphaCboArrow
    Else
        SetButtonBorder m.strLabel, True
    End If

    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.picAlphaDrop_Click"

End Sub

Private Sub picNumButton_Click(Index As Integer)
On Error GoTo ErrSection:

    ClearBorder
    picNumButton(Index).BorderStyle = 1
    SetNumericString Index
    
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.picNumButton_Click"

End Sub

Private Sub picNumCbo_Click()
On Error Resume Next

    cmdNumCbo_Click
    
End Sub

Private Sub picNumCboArrow_Click()
On Error Resume Next

    cmdNumCbo_Click

End Sub

Private Sub picNumDrop_Click(Index As Integer)
On Error GoTo ErrSection:

    ClearBorder
    picNumCbo.Picture = picNumDrop(Index).Picture
    fraNumDrop.Visible = False
    fraNumDrop.ZOrder 1
    
    If tbToolbar.Tools("ID_ElliotLabels").State = ssUnchecked Then tbToolbar.Tools("ID_ElliotLabels").State = ssChecked

    Select Case Index
        Case 0
            NumButtonPrint "1", "2", "3", "4", "5"
            m.eNumberStyle = eNumStyle_Normal
        Case 1
            NumButtonPrint "(1)", "(2)", "(3)", "(4)", "(5)"
            m.eNumberStyle = eNumStyle_Normal_Paren
        Case 2
            NumButtonPic "kOne"
            m.eNumberStyle = eNumStyle_Normal_Circle
        Case 3
            NumButtonPrint "i", "ii", "iii", "iv", "v"
            m.eNumberStyle = eNumStyle_LowerR
        Case 4
            NumButtonPrint "(i)", "(ii)", "(iii)", "(iv)", "(v)"
            m.eNumberStyle = eNumStyle_LowerR_Paren
        Case 5
            NumButtonPic "kOneR"
            m.eNumberStyle = eNumStyle_LowerR_Circle
        Case 6
            NumButtonPrint "I", "II", "III", "IV", "V"
            m.eNumberStyle = eNumStyle_UpperR
        Case 7
            NumButtonPrint "(I)", "(II)", "(III)", "(IV)", "(V)"
            m.eNumberStyle = eNumStyle_UpperR_Paren
        Case 8
            NumButtonPic "kOneCapR"
            m.eNumberStyle = eNumStyle_UpperR_Circle
    End Select

    If Len(m.strLabel) = 0 Then
        picNumButton(0).BorderStyle = 1
        SetNumericString 0
        MoveFocus picNumCboArrow
    Else
        SetButtonBorder m.strLabel, True
    End If

    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.picNumDrop_Click"

End Sub

Private Sub NumButtonPic(ByVal strKeyStart$)
On Error GoTo ErrSection:

    Dim i&, j&
    
    j = ImgListNum.ListImages(strKeyStart).Index
    
    If j > 0 And j + 4 <= ImgListNum.ListImages.Count Then
        For i = 0 To 4
            picNumButton(i).Cls
            picNumButton(i).Picture = ImgListNum.ListImages(j).Picture
            j = j + 1
        Next
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.NumButtonPic"

End Sub

Private Sub NumButtonPrint(ByVal strChar$, ByVal strChar1$, _
    ByVal strChar2$, ByVal strChar3$, ByVal strChar4$)
On Error GoTo ErrSection:

    Dim i&
    
    For i = 0 To 4
        picNumButton(i).Cls
        picNumButton(i).Picture = Picture16("kBlank")
    Next
    
    picNumButton(0).Print strChar
    picNumButton(1).Print strChar1
    picNumButton(2).Print strChar2
    picNumButton(3).Print strChar3
    picNumButton(4).Print strChar4

    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.NumButtonPrint"

End Sub

Private Sub AlphaButtonPic(ByVal strKeyStart$)
On Error GoTo ErrSection:

    Dim i&, j&
    
    j = ImgListAlpha.ListImages(strKeyStart).Index
    
    If j > 0 And j + 8 <= ImgListAlpha.ListImages.Count Then
        For i = 0 To 8
            picAlphaButton(i).Cls
            picAlphaButton(i).Picture = ImgListAlpha.ListImages(j).Picture
            j = j + 1
        Next
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.AlphaButtonPic"

End Sub

Private Sub AlphaButtonPrint(ByVal strChar$, ByVal strChar1$, _
    ByVal strChar2$, ByVal strChar3$, ByVal strChar4$, ByVal strChar5$, _
    ByVal strChar6$, ByVal strChar7$, ByVal strChar8$)
On Error GoTo ErrSection:

    Dim i&
    
    For i = 0 To 8
        picAlphaButton(i).Cls
        picAlphaButton(i).Picture = Picture16("kBlank")
    Next
    
    picAlphaButton(0).Print strChar
    picAlphaButton(1).Print strChar1
    picAlphaButton(2).Print strChar2
    picAlphaButton(3).Print strChar3
    picAlphaButton(4).Print strChar4
    picAlphaButton(5).Print strChar5
    picAlphaButton(6).Print strChar6
    picAlphaButton(7).Print strChar7
    picAlphaButton(8).Print strChar8

    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.AlphaButtonPrint"

End Sub

Private Sub ClearBorder()
On Error GoTo ErrSection:

    Dim i&
    
    If fraAlphaDrop.Visible Then fraAlphaDrop.Visible = False
    If fraNumDrop.Visible Then fraNumDrop.Visible = False
    
    For i = 0 To 4
        picNumButton(i).BorderStyle = 0
    Next
    
    For i = 0 To 8
        picAlphaButton(i).BorderStyle = 0
    Next
    
    If Len(g.strActiveDraw) = 0 Then
        If HasModule("EWL") Then
            g.strActiveDraw = "ID_ElliotEndUser"        '6988
            frmMain.tbToolbar.Tools("ID_ElliotEndUser").State = ssChecked
        Else
            g.strActiveDraw = "ID_ElliotLabels"
            frmMain.tbToolbar.Tools("ID_ElliotLabels").State = ssChecked
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.ClearBorder"

End Sub

Private Sub SetNumericString(ByVal nIdx As Long)
On Error GoTo ErrSection:

    m.strLabel = ""
    m.eImage = eCNI_Ascii

    If tbToolbar.Tools("ID_ElliotLabels").State = ssUnchecked Then tbToolbar.Tools("ID_ElliotLabels").State = ssChecked
    
    If m.eNumberStyle = eNumStyle_Normal Or m.eNumberStyle = eNumStyle_Normal_Paren Or _
        m.eNumberStyle = eNumStyle_Normal_Circle Then
        m.strLabel = Str(nIdx + 1)
    ElseIf m.eNumberStyle = eNumStyle_LowerR Or m.eNumberStyle = eNumStyle_LowerR_Paren Or _
        m.eNumberStyle = eNumStyle_LowerR_Circle Then
        Select Case nIdx
            Case 0
                m.strLabel = "i"
            Case 1
                m.strLabel = "ii"
            Case 2
                m.strLabel = "iii"
            Case 3
                m.strLabel = "iv"
            Case 4
                m.strLabel = "v"
        End Select
    Else
        Select Case nIdx
            Case 0
                m.strLabel = "I"
            Case 1
                m.strLabel = "II"
            Case 2
                m.strLabel = "III"
            Case 3
                m.strLabel = "IV"
            Case 4
                m.strLabel = "V"
        End Select
    End If

    If m.eNumberStyle = eNumStyle_Normal_Paren Or m.eNumberStyle = eNumStyle_LowerR_Paren Or _
        m.eNumberStyle = eNumStyle_UpperR_Paren Then
        m.strLabel = "(" & m.strLabel & ")"
    ElseIf m.eNumberStyle = eNumStyle_Normal_Circle Or m.eNumberStyle = eNumStyle_LowerR_Circle Or _
        m.eNumberStyle = eNumStyle_UpperR_Circle Then
        m.eImage = eCNI_Circle
    End If
    
    If Not m.bInitInprog Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If

    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.SetNumericString"

End Sub

Private Sub SetAlphaString(ByVal nIdx As Long)
On Error GoTo ErrSection:

    m.strLabel = ""
    m.eImage = eCNI_Ascii
    
    If tbToolbar.Tools("ID_ElliotLabels").State = ssUnchecked Then tbToolbar.Tools("ID_ElliotLabels").State = ssChecked
    
    If m.eCharStyle = eAlphaStyle_Lower Or m.eCharStyle = eAlphaStyle_Lower_Paren Or _
        m.eCharStyle = eAlphaStyle_Lower_Circle Then
            Select Case nIdx
                Case 0
                    m.strLabel = "a"
                Case 1
                    m.strLabel = "b"
                Case 2
                    m.strLabel = "c"
                Case 3
                    m.strLabel = "d"
                Case 4
                    m.strLabel = "e"
                Case 5
                    m.strLabel = "x"
                Case 6
                    m.strLabel = "w"
                Case 7
                    m.strLabel = "y"
                Case 8
                    m.strLabel = "z"
            End Select
            If m.eCharStyle = eAlphaStyle_Lower_Paren Then
                m.strLabel = "(" & m.strLabel & ")"
            ElseIf m.eCharStyle = eAlphaStyle_Lower_Circle Then
                m.eImage = eCNI_Circle
            End If
    Else
        Select Case nIdx
            Case 0
                m.strLabel = "A"
            Case 1
                m.strLabel = "B"
            Case 2
                m.strLabel = "C"
            Case 3
                m.strLabel = "D"
            Case 4
                m.strLabel = "E"
            Case 5
                m.strLabel = "X"
            Case 6
                m.strLabel = "W"
            Case 7
                m.strLabel = "Y"
            Case 8
                m.strLabel = "Z"
        End Select
        If m.eCharStyle = eAlphaStyle_Upper_Paren Then
            m.strLabel = "(" & m.strLabel & ")"
        ElseIf m.eCharStyle = eAlphaStyle_Upper_Circle Then
            m.eImage = eCNI_Circle
        End If
    End If
    
    If Not m.bInitInprog Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.SetAlphaString"

End Sub

Public Sub GetSettings(Annot As cAnnotation)
On Error GoTo ErrSection:

    Dim i&, sttTemp$, strSizes$
    Dim bSave As Boolean
    Dim Chart As cChart
    
    If Annot Is Nothing Then Exit Sub

    bSave = Annot.AllowEWIGMP Or Annot.AllowGMP
    
    If bSave Then
        For i = eCommon_Miniscule To eCommon_GrandSupercycle
            sttTemp = Str(Int(ValOfText(txtCustomFont(i).Text)))
            If Len(sttTemp) < 1 Then sttTemp = "10"
            strSizes = strSizes & sttTemp & "|"
        Next
    End If

    With Annot
        .Text = m.strLabel
        .Color = gdColor.Color
        .geTextAlign = 8
        
        .Prop("ImageType") = m.eImage
        .Prop("ImageSize") = 999999
        .Prop("ImageStyle") = 0
        
        'style - 0=reg,1=bold,2=italic,3=bold italic
        i = 0
        If m.nBold = 1 Then
            If m.nItalic = 1 Then
                i = 3
            Else
                i = 1
            End If
        ElseIf m.nItalic = 1 Then
            i = 2
        End If
        
        .Prop("FontName") = m.strFont
        .Prop("FontSize") = m.nFontSize
        .Prop("FontUnderline") = 0
        .Prop("FontStyle") = i
        
        If bSave Then
            .Prop("Border") = chkBorder.Value
            .Prop("CustomFontSizes") = strSizes
            .Prop("UseCopyright") = chkCopyright.Value
            If optJustify(0) Then
                .Prop("TextJustify") = "0"
            ElseIf optJustify(1) Then
                .Prop("TextJustify") = "1"
            ElseIf optJustify(2) Then
                .Prop("TextJustify") = "2"
            End If
            
            If optStylePreset(eCommon_CustomText).Value = True Then
                .Text = ""
                If chkCopyright.Value = 1 Then
                    sttTemp = lblCopyright.Caption & " " & cboMonth.Text & " " & cboYear.Text & vbCrLf & rtfText.Text
                Else
                    sttTemp = rtfText.Text
                End If
                .Text = sttTemp
            End If
            
            If Annot.AllowEWIGMP Then
                .Prop("UseCustomSizes") = optFontCustom.Value
            Else
                .Prop("UseCustomSizes") = 1
            End If
        End If
        
        .Prop("NumberStyle") = m.eNumberStyle
        .Prop("AlphaStyle") = m.eCharStyle
        
        .PreIndicator = chkPreIndicator.Value
        .MultiChartFlag = -1 * chkMultiChart.Value
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.GetSettings"

End Sub

Private Sub InitFromAnnot()
On Error GoTo ErrSection:

    Dim strChar$, nStyle&, i&, j&, iYear&, iMonth&, strText$
    Dim bParen As Boolean
    Dim bNumeric As Boolean
    Dim aFontSizes As cGdArray
    
    If m.Annot Is Nothing Then Exit Sub     'precautionary
    
    'set flag so certain code in click event of some controls will not execute
    m.bInitInprog = True
    
    With m.Annot
        m.eImage = Val(.Prop("ImageType"))
        m.strFont = .Prop("FontName")
        m.nFontSize = ValOfText(.Prop("FontSize"))
        nStyle = ValOfText(.Prop("FontStyle"))
        
        gdColor.Color = .Color
        
        m.bWasMultiChart = .MultiChartFlag
        chkMultiChart.Value = Abs(m.bWasMultiChart)
        chkPreIndicator.Value = .PreIndicator
    End With
    
    'style - 0=reg,1=bold,2=italic,3=bold italic
    If nStyle = 1 Then
        m.nBold = 1
    ElseIf nStyle = 2 Then
        m.nItalic = 1
    ElseIf nStyle = 3 Then
        m.nBold = 1
        m.nItalic = 1
    End If

'''''''''''''''''''''''''''''''''''
'settings for GMP palette (begin)
'''''''''''''''''''''''''''''''''''
    
    If m.Annot.AllowEWIGMP Or m.Annot.AllowGMP Then
        Set aFontSizes = New cGdArray
        aFontSizes.SplitFields m.Annot.Prop("CustomFontSizes"), "|"
    End If
    
    If Not m.Annot.AllowEWIGMP Then
        lblInfoPreset.Top = 270
        fraCustomFont.Top = 630
    End If
    
    'show radio buttons for preset / custom sizes if EWI_GMP (i.e. both flag files)
    optFontPreset.Visible = m.Annot.AllowEWIGMP
    optFontCustom.Visible = m.Annot.AllowEWIGMP
    optFontPreset.Enabled = m.Annot.AllowEWIGMP
    optFontCustom.Enabled = m.Annot.AllowEWIGMP
    
    If m.Annot.AllowEWIGMP Or m.Annot.AllowGMP Then
        For i = eCommon_Miniscule To eCommon_GrandSupercycle
            If aFontSizes.Size > i Then
                txtCustomFont(i).Text = aFontSizes(i)
            End If
        Next
    End If
    
    'do these last of the GMP controls so text fields can be toggle correctly
    If m.bEditExisting Then
        optFontCustom.Value = False
        optFontPreset.Value = False
        
        If m.Annot.AllowGMP And Not m.Annot.AllowEWIGMP Then
            AdjustPresetLabels 1
        Else
            AdjustPresetLabels 0
        End If
        
    ElseIf m.Annot.AllowEWIGMP Or m.Annot.AllowGMP Then
        If m.Annot.AllowEWIGMP Then
            optFontCustom.Value = Val(m.Annot.Prop("UseCustomSizes"))
        Else
            optFontCustom.Value = True
        End If
        optFontPreset.Value = Not optFontCustom.Value
    Else
        optFontPreset.Value = True          'to trigger hiding of text controls
        optFontCustom.Value = False
    End If
    
    If m.Annot.AllowGMP Then
        chkBorder.Value = Val(m.Annot.Prop("Border"))
        
        'TLB: default to the month they last used, unless it's in the past
        iMonth = GetIniFileProperty("EWYearMonth", 0, "", g.strIniFile)
        If iMonth < Year(Date) * 100 + Month(Date) Then
            iMonth = Year(Date) * 100 + Month(Date)
        End If
        iYear = Int(iMonth / 100)
        iMonth = iMonth Mod 100
        
        cboMonth.Clear
        For i = 1 To 12
            cboMonth.AddItem MonthName(i, False)
        Next
        cboMonth.ListIndex = iMonth - 1
        
        cboYear.Clear
        For i = 2000 To iYear + 5
            cboYear.AddItem Str(i)
            If i = iYear Then
                cboYear.ListIndex = cboYear.ListCount - 1
            End If
        Next
        
        'custom text
        i = Val(m.Annot.Prop("TextJustify"))
        If i >= 0 And i <= 2 Then
            optJustify(i).Value = True
        Else
            optJustify(1).Value = True
        End If
        chkCopyright.Value = Val(m.Annot.Prop("UseCopyright"))
        
        
        rtfText.Font.Name = m.strFont
        rtfText.Font.Size = m.nFontSize
        rtfText.Font.Bold = m.nBold
        rtfText.Font.Italic = m.nItalic
        
        strText = m.Annot.Prop("CustomText")
        If Len(strText) = 0 Then
            ' TLB 4/29/2015: "Data courtesy of TradeNav" taken out per Glen
            rtfText.Text = "Elliott Wave International" & vbCrLf & "www.elliottwave.com" '& vbCrLf & "Data courtesy TradeNavigator.com"
        Else
            rtfText.Text = Replace(strText, "~", vbCrLf)
        End If
    End If
'''''''''''''''''''''''''''''''''''
'settings for GMP palette (end)
'''''''''''''''''''''''''''''''''''
    
    If Left(m.strLabel, 1) = "(" Then
        bParen = True
        strChar = Mid(m.strLabel, 2, Len(m.strLabel) - 2)
    Else
        strChar = m.strLabel
    End If
        
    If Len(strChar) <= 0 Then
        Dim nIdx As Integer
        nIdx = m.eNumberStyle
        picNumDrop_Click nIdx
        nIdx = m.eCharStyle
        picAlphaDrop_Click nIdx
        
        m.bInitInprog = False
        
        Exit Sub
    End If
        
    Select Case strChar
        Case "1", "2", "3", "4", "5"
            If bParen Then
                m.eNumberStyle = eNumStyle_Normal_Paren
                NumButtonPrint "(1)", "(2)", "(3)", "(4)", "(5)"
            ElseIf m.eImage = eCNI_Circle Then
                m.eNumberStyle = eNumStyle_Normal_Circle
                NumButtonPic "kOne"
            Else
                m.eNumberStyle = eNumStyle_Normal
                NumButtonPrint "1", "2", "3", "4", "5"
            End If
            bNumeric = True
        Case "i", "ii", "iii", "iv", "v"
            If bParen Then
                m.eNumberStyle = eNumStyle_LowerR_Paren
                NumButtonPrint "(i)", "(ii)", "(iii)", "(iv)", "(v)"
            ElseIf m.eImage = eCNI_Circle Then
                m.eNumberStyle = eNumStyle_LowerR_Circle
                NumButtonPic "kOneR"
            Else
                m.eNumberStyle = eNumStyle_LowerR
                NumButtonPrint "i", "ii", "iii", "iv", "v"
            End If
            bNumeric = True
        Case "I", "II", "III", "IV", "V"
            If bParen Then
                m.eNumberStyle = eNumStyle_UpperR_Paren
                NumButtonPrint "(I)", "(II)", "(III)", "(IV)", "(V)"
            ElseIf m.eImage = eCNI_Circle Then
                m.eNumberStyle = eNumStyle_UpperR_Circle
                NumButtonPic "kOneCapR"
            Else
                m.eNumberStyle = eNumStyle_UpperR
                NumButtonPrint "I", "II", "III", "IV", "V"
            End If
            bNumeric = True
        Case "a", "b", "c", "d", "e", "x", "w", "y", "z"
            If bParen Then
                AlphaButtonPrint "(a)", "(b)", "(c)", "(d)", "(e)", "(x)", "(w)", "(y)", "(z)"
                m.eCharStyle = eAlphaStyle_Lower_Paren
            ElseIf m.eImage = eCNI_Circle Then
                AlphaButtonPic "kLowerA"
                m.eCharStyle = eAlphaStyle_Lower_Circle
            Else
                AlphaButtonPrint "a", "b", "c", "d", "e", "x", "w", "y", "z"
                m.eCharStyle = eAlphaStyle_Lower
            End If
        Case "A", "B", "C", "D", "E", "X", "W", "Y", "Z"
            If bParen Then
                AlphaButtonPrint "(A)", "(B)", "(C)", "(D)", "(E)", "(X)", "(W)", "(Y)", "(Z)"
                m.eCharStyle = eAlphaStyle_Upper_Paren
            ElseIf m.eImage = eCNI_Circle Then
                AlphaButtonPic "kCapA"
                m.eCharStyle = eAlphaStyle_Upper_Circle
            Else
                AlphaButtonPrint "A", "B", "C", "D", "E", "X", "W", "Y", "Z"
                m.eCharStyle = eAlphaStyle_Upper
            End If
    End Select
    
    If bNumeric Then
        If m.eCommon = eCommon_None Then
            m.eCharStyle = eAlphaStyle_Lower
            AlphaButtonPrint "a", "b", "c", "d", "e", "x", "w", "y", "z"
        End If
    Else
        If m.eCommon = eCommon_None Then
            m.eNumberStyle = eNumStyle_Normal
            NumButtonPrint "1", "2", "3", "4", "5"
        End If
    End If
    
    SetButtonBorder strChar, False
    
    If m.eCommon = eCommon_None Then
        picAlphaCbo.Picture = picAlphaDrop(m.eCharStyle)
        picNumCbo.Picture = picNumDrop(m.eNumberStyle)
    End If
    
    m.bInitInprog = False
    
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.InitFromAnnot"

End Sub

Private Sub SetButtonBorder(ByVal strCharIn$, ByVal bUpdateLabel As Boolean)
On Error GoTo ErrSection:

    Dim strChar$, nIdx As Integer
    Dim bNumeric As Boolean

    nIdx = -1
    strChar = strCharIn
    
    If Left(strChar, 1) = "(" Then
        strChar = Mid(strChar, 2, Len(strChar) - 2)
    End If

    Select Case strChar
        Case "1", "i", "I"
            nIdx = 0
            bNumeric = True
        Case "2", "ii", "II"
            nIdx = 1
            bNumeric = True
        Case "3", "iii", "III"
            nIdx = 2
            bNumeric = True
        Case "4", "iv", "IV"
            nIdx = 3
            bNumeric = True
        Case "5", "v", "V"
            nIdx = 4
            bNumeric = True
        Case "a", "A"
            nIdx = 0
        Case "b", "B"
            nIdx = 1
        Case "c", "C"
            nIdx = 2
        Case "d", "D"
            nIdx = 3
        Case "e", "E"
            nIdx = 4
        Case "x", "X"
            nIdx = 5
        Case "w", "W"
            nIdx = 6
        Case "y", "Y"
            nIdx = 7
        Case "z", "Z"
            nIdx = 8
    End Select
    
    If nIdx >= 0 Then
        If bNumeric Then
            If bUpdateLabel Then
                picNumButton_Click nIdx
            Else
                picNumButton(nIdx).BorderStyle = 1
            End If
        Else
            If bUpdateLabel Then
                picAlphaButton_Click nIdx
            Else
                picAlphaButton(nIdx).BorderStyle = 1
            End If
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.SetButtonBorder"
    
End Sub

Private Sub tbToolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim Chart As cChart
    Dim Annot As cAnnotation
    Dim eType As eAnnotType
    Dim Tool As SSTool
    
    If Button = vbRightButton Then
        Set Tool = tbToolbar.ToolFromPosition(X, Y)
        If Not Tool Is Nothing Then
            If Not ActiveChart Is Nothing Then
                Set Chart = ActiveChart.Chart
                If Not Chart Is Nothing Then
                    Set Annot = New cAnnotation
                    eType = Annot.AnnotTypeFromToolID(Tool.ID)
                    Chart.RemoveAnnots True, eType
                    Chart.GenerateChart eRedo1_Scrolled
                End If
            End If
        End If
    End If
    
    Set Chart = Nothing
    Set Annot = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.tbToolbar_MouseDown"

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim Chart As cChart
        
    'EWI requested that the palette stays open when the move-tool is in use.
    'When the move-tool is in use the cursor is not pencil and main toolbar
    'is set to the move-tool button so just make sure to resync if necessary.
    If Not ActiveChart Is Nothing Then
        If ActiveChart.pbCursor <> eCursor_Pencil Then
            ToolbarSetCursorGroup frmMain.tbToolbar, True, "ID_ElliotLabels"
            ActiveChart.pbCursor = eCursor_Pencil
        End If
    End If
    
    If Tool.ID = "ID_DeleteIcon" Then
        If Not m.Annot Is Nothing Then
            Set Chart = m.Annot.AnnotChart
            If Not Chart Is Nothing Then
                m.Annot.geRemoveAnnotation (Chart.geChartObj)
                Chart.Annots.Remove m.Annot.geAnnId
                Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart
            End If
        End If
        Unload Me
    ElseIf Tool.ID = "ID_Close" Then
        If Not m.Annot Is Nothing Then
            If Not m.Annot.AnnotChart Is Nothing Then
                GetSettings m.Annot
                m.Annot.AnnotChart.GenerateChart eRedo1_Scrolled
            End If
        End If
        Unload Me
    Else
        g.strActiveDraw = Tool.ID
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.tbToolbar_ToolClick"

End Sub

Private Function IsEWIActiveDraw() As Boolean
On Error GoTo ErrSection:
    
    If Not FormIsLoaded("frmElliot") Then Exit Function
    
    If g.strActiveDraw = "ID_Trendline" Or g.strActiveDraw = "ID_TrendChannel" Or _
       g.strActiveDraw = "ID_ElliotLabels" Or g.strActiveDraw = "ID_ArrowLine" Or _
       g.strActiveDraw = "ID_RegressionLine" Or g.strActiveDraw = "Fibonacci" Or _
       g.strActiveDraw = "ID_DNExpansion" Or g.strActiveDraw = "ID_ElliotEndUser" Then
    
        IsEWIActiveDraw = True
        
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmElliot.IsEWIActiveDraw"

End Function

Private Sub ToggleTopMost(ByVal bTopMost As Boolean)
On Error GoTo ErrSection:

    If Not ActiveChart Is Nothing Then
        If ActiveChart.DetachStatus = eDetached Then
            SetFormTopmost Me, bTopMost             '6441
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.ToggleTopMost", eGDRaiseError_Raise

End Sub

Private Sub txtCustomFont_Change(Index As Integer)
On Error GoTo ErrSection:

    If optStylePreset(Index).Value Then
        m.nFontSize = Int(ValOfText(txtCustomFont(Index).Text))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.txtCustomFont_Change", eGDRaiseError_Raise

End Sub

Private Sub AdjustPresetLabels(ByVal nUseCustom As Long)
On Error GoTo ErrSection:

    Dim i&
    Dim optIndex As Integer

    optIndex = -1
    
    If 0 = nUseCustom Then
        lblPresetSize10.Caption = "Font size 10"
        lblPresetSize12.Caption = "Font size 12"
        lblPresetSize14.Caption = "Font size 14"
        
        'set to values as placed in form design
        lblPresetSize10.Left = 240
        lblPresetSize12.Left = 1965
        lblPresetSize14.Left = 3690
        
        lblPresetSize14.Width = lblPresetSize12.Width
    
        For i = eCommon_Miniscule To eCommon_GrandSupercycle
            txtCustomFont(i).Visible = False
            txtCustomFont(i).Enabled = False
            
            If optStylePreset(i).Value = True Then
                optIndex = i
            End If
        Next
    Else
        lblPresetSize10.Caption = "Font size"
        lblPresetSize12.Caption = "Font size"
        lblPresetSize14.Caption = "Font size"
        
        lblPresetSize10.Left = lblPresetSize10.Left + 645
        lblPresetSize12.Left = lblPresetSize12.Left + 645
        lblPresetSize14.Left = lblPresetSize14.Left + 645
        
        lblPresetSize14.Width = lblPresetSize14.Width - 150
    
        For i = eCommon_Miniscule To eCommon_GrandSupercycle
            txtCustomFont(i).Visible = True
            txtCustomFont(i).Enabled = True
        
            If optStylePreset(i).Value = True Then
                optIndex = i
            End If
        Next
    End If
    
    If optIndex <> -1 Then
        optStylePreset_Click optIndex       '6890
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmElliot.AdjustPresetLabels", eGDRaiseError_Raise

End Sub

Private Sub RestoreDefaultFontInfo()
On Error GoTo ErrSection

    Dim nStyle&
    
    If Not m.AnnotDefaults Is Nothing Then
        With m.AnnotDefaults
            m.strFont = .Prop("FontName")
            nStyle = 1
            
            If fraText.Visible Then
                If Me.optStylePreset(eCommon_CustomText).Value = True Then
                    m.nFontSize = ValOfText(.Prop("FontSize"))
                    nStyle = ValOfText(.Prop("FontStyle"))
                End If
            End If
        End With
    
        'style - 0=reg,1=bold,2=italic,3=bold italic
        If nStyle = 1 Then
            m.nBold = 1
            m.nItalic = 0
        ElseIf nStyle = 2 Then
            m.nBold = 0
            m.nItalic = 1
        ElseIf nStyle = 3 Then
            m.nBold = 1
            m.nItalic = 1
        Else
            m.nBold = 0
            m.nItalic = 0
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmElliot.RestoreDefaultFontInfo", eGDRaiseError_Raise

End Sub

Private Function IndexForNum(ByVal Char As String) As Integer
On Error GoTo ErrExit

    Dim i As Integer
    
    i = -1
    
    Select Case Char
        Case "1", "(1)", "i", "(i)"
            i = 0
        Case "2", "(2)", "ii", "(ii)"
            i = 1
        Case "3", "(3)", "iii", "(iii)"
            i = 2
        Case "4", "(4)", "iv", "(iv)"
            i = 3
        Case "5", "(5)", "v", "(v)"
            i = 4
    End Select
    
    IndexForNum = i

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmElliot.IndexForNum", eGDRaiseError_Raise

End Function

Private Function IndexForChar(ByVal Char As String) As Integer
On Error GoTo ErrExit

    Dim i As Integer
    
    i = -1
    
    Select Case Char
        Case "A", "a", "(A)", "(a)"
            i = 0
        Case "B", "b", "(B)", "(b)"
            i = 1
        Case "C", "c", "(C)", "(c)"
            i = 2
        Case "D", "d", "(D)", "(d)"
            i = 3
        Case "E", "e", "(E)", "(e)"
            i = 4
        Case "X", "x", "(X)", "(x)"
            i = 5
        Case "W", "w", "(W)", "(w)"
            i = 6
        Case "Y", "y", "(Y)", "(y)"
            i = 7
        Case "Z", "z", "(Z)", "(z)"
            i = 8
    End Select
    
    IndexForChar = i

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmElliot.IndexForChar", eGDRaiseError_Raise

End Function



