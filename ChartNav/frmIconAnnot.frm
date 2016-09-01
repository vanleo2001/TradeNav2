VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmIconAnnot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Icon Palette"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3855
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniButtonImageXP cmdDeleteIcon 
      Height          =   375
      Left            =   2220
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   855
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
      Caption         =   "frmIconAnnot.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmIconAnnot.frx":002E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmIconAnnot.frx":004E
      RightToLeft     =   0   'False
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2730
      Top             =   75
   End
   Begin HexUniControls.ctlUniFrameWL fraSavedIcons 
      Height          =   2385
      Left            =   75
      TabIndex        =   54
      Top             =   4455
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
      Caption         =   "frmIconAnnot.frx":006A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmIconAnnot.frx":00B2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIconAnnot.frx":00D2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDown 
         Height          =   315
         Left            =   2370
         TabIndex        =   51
         Top             =   1935
         Width           =   1170
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmIconAnnot.frx":00EE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":0122
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":0142
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdUp 
         Height          =   315
         Left            =   2370
         TabIndex        =   63
         Top             =   1575
         Width           =   1170
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmIconAnnot.frx":015E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":018E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":01AE
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fgSaved 
         Height          =   2010
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   2085
         _cx             =   3678
         _cy             =   3545
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
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   315
         Left            =   2370
         TabIndex        =   57
         Top             =   585
         Width           =   1170
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmIconAnnot.frx":01CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":01F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":0218
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRename 
         Height          =   315
         Left            =   2370
         TabIndex        =   56
         Top             =   960
         Width           =   1170
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmIconAnnot.frx":0234
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":0262
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":0282
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSaveAs 
         Height          =   315
         Left            =   2370
         TabIndex        =   55
         Top             =   210
         Width           =   1170
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
         Caption         =   "frmIconAnnot.frx":029E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":02C8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":02E8
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraIconSettings 
      Height          =   3855
      Left            =   60
      TabIndex        =   2
      Top             =   540
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
      Caption         =   "frmIconAnnot.frx":0304
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmIconAnnot.frx":033E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIconAnnot.frx":035E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdFont 
         Height          =   330
         Left            =   2850
         TabIndex        =   53
         Top             =   2370
         Width           =   750
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
         Caption         =   "frmIconAnnot.frx":037A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":03A4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":03C4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraSize 
         Height          =   255
         Left            =   1080
         TabIndex        =   58
         Top             =   360
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
         Caption         =   "frmIconAnnot.frx":03E0
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmIconAnnot.frx":040E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":042E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optLarge 
            Height          =   220
            Left            =   1740
            TabIndex        =   61
            Top             =   0
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
            Caption         =   "frmIconAnnot.frx":044A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmIconAnnot.frx":0474
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmIconAnnot.frx":0494
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optMedium 
            Height          =   220
            Left            =   780
            TabIndex        =   60
            Top             =   0
            Width           =   915
            _ExtentX        =   1614
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
            Caption         =   "frmIconAnnot.frx":04B0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmIconAnnot.frx":04DC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmIconAnnot.frx":04FC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optSmall 
            Height          =   220
            Left            =   60
            TabIndex        =   59
            Top             =   0
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
            Caption         =   "frmIconAnnot.frx":0518
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmIconAnnot.frx":0542
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmIconAnnot.frx":0562
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniCheckXP chkAutoIncrement 
         Height          =   220
         Left            =   600
         TabIndex        =   52
         Top             =   1860
         Width           =   2595
         _ExtentX        =   4577
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
         Caption         =   "frmIconAnnot.frx":057E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":05D6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":05F6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCenter 
         Height          =   220
         Left            =   855
         TabIndex        =   50
         Top             =   2505
         Width           =   795
         _ExtentX        =   1402
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
         Caption         =   "frmIconAnnot.frx":0612
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmIconAnnot.frx":063E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":065E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optRight 
         Height          =   220
         Left            =   1695
         TabIndex        =   49
         Top             =   2520
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
         Caption         =   "frmIconAnnot.frx":067A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":06A4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":06C4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLeft 
         Height          =   220
         Left            =   120
         TabIndex        =   48
         Top             =   2505
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
         Caption         =   "frmIconAnnot.frx":06E0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":0708
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":0728
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtCustom 
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   2790
         Width           =   3480
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmIconAnnot.frx":0744
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
         Tip             =   "frmIconAnnot.frx":0764
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":0784
      End
      Begin HexUniControls.ctlUniCheckXP chkMultiChart 
         Height          =   220
         Left            =   120
         TabIndex        =   45
         Top             =   3510
         Width           =   3495
         _ExtentX        =   6165
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
         Caption         =   "frmIconAnnot.frx":07A0
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":07E4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":0804
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkPreIndicator 
         Height          =   220
         Left            =   120
         TabIndex        =   44
         Top             =   3270
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
         Caption         =   "frmIconAnnot.frx":0820
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmIconAnnot.frx":086C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":088C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fg 
         Height          =   1095
         Left            =   540
         TabIndex        =   3
         Top             =   720
         Width           =   2595
         _cx             =   4577
         _cy             =   1931
         _ConvInfo       =   1
         Appearance      =   0
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
         BackColorSel    =   8421504
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483648
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483648
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   39
            Left            =   2295
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   43
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   38
            Left            =   2040
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   42
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   37
            Left            =   1785
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   41
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   36
            Left            =   1530
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   40
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   35
            Left            =   1275
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   39
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   34
            Left            =   1020
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   38
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   33
            Left            =   765
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   37
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   32
            Left            =   510
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   36
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   31
            Left            =   255
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   35
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   30
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   34
            Top             =   765
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            Picture         =   "frmIconAnnot.frx":08A8
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   33
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   1275
            Picture         =   "frmIconAnnot.frx":0BB2
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   32
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   0
            Picture         =   "frmIconAnnot.frx":0EBC
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   31
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   1275
            Picture         =   "frmIconAnnot.frx":11C6
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   30
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   29
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   25
            Left            =   1275
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   28
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   255
            Picture         =   "frmIconAnnot.frx":14D0
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   27
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   1530
            Picture         =   "frmIconAnnot.frx":17DA
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   26
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   255
            Picture         =   "frmIconAnnot.frx":1AE4
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   25
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   1530
            Picture         =   "frmIconAnnot.frx":1DEE
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   24
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   21
            Left            =   255
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   23
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   26
            Left            =   1530
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   22
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   510
            Picture         =   "frmIconAnnot.frx":20F8
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   21
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   1785
            Picture         =   "frmIconAnnot.frx":2402
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   20
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   510
            Picture         =   "frmIconAnnot.frx":270C
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   19
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   1785
            Picture         =   "frmIconAnnot.frx":2A16
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   18
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   22
            Left            =   510
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   17
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   27
            Left            =   1785
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   16
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   765
            Picture         =   "frmIconAnnot.frx":2D20
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   15
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   2040
            Picture         =   "frmIconAnnot.frx":302A
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   14
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   765
            Picture         =   "frmIconAnnot.frx":3334
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   13
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   2040
            Picture         =   "frmIconAnnot.frx":363E
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   12
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   23
            Left            =   765
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   11
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   28
            Left            =   2040
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   10
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1020
            Picture         =   "frmIconAnnot.frx":3948
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   9
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   2295
            Picture         =   "frmIconAnnot.frx":3C52
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   8
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   1020
            Picture         =   "frmIconAnnot.frx":3F5C
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   7
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   2295
            Picture         =   "frmIconAnnot.frx":4266
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   6
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   24
            Left            =   1020
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   5
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   29
            Left            =   2295
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   4
            Top             =   510
            Width           =   255
         End
      End
      Begin gdOCX.gdSelectColor clrColor 
         Height          =   315
         Left            =   120
         TabIndex        =   47
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Left            =   120
         Top             =   2250
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
         Caption         =   "frmIconAnnot.frx":4570
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmIconAnnot.frx":45B2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmIconAnnot.frx":45D2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdClose 
      Height          =   375
      Left            =   2940
      TabIndex        =   0
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
      Caption         =   "frmIconAnnot.frx":45EE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmIconAnnot.frx":461A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmIconAnnot.frx":463A
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblInfoPrompt 
      Height          =   435
      Left            =   60
      Top             =   60
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
      Caption         =   "frmIconAnnot.frx":4656
      BackColor       =   -2147483633
      ForeColor       =   -2147483635
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmIconAnnot.frx":470C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmIconAnnot.frx":472C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmIconAnnot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'- saved icons: message, show icon from editor, "custom", sorted?

Private Const kMax = 39
Private Const kLastPicIndex = 39
Private Const kRows = 4
Private Const kCols = 10
Private Const kBoxSize = 235
Private Const kBoxSpacing = 8
Private Const kMultiChartLabel = "Show on all charts for "

Private Enum ePicIdx
    ePic_ArrowNorth = 0
    ePic_ArrowSouth
    ePic_ArrowEast
    ePic_ArrowWest
    ePic_ArrowNorthEast
    ePic_ArrowSouthWest
    ePic_ArrowSouthEast
    ePic_ArrowNorthWest
    ePic_Plus
    ePic_Cross
    'shapes
    ePic_TriUp_Filled
    ePic_TriDown_Filled
    ePic_Circle_Filled
    ePic_Square_Filled
    ePic_Diamond_Filled
    ePic_TriUp_Hollow
    ePic_TriDown_Hollow
    ePic_Circle_Hollow
    ePic_Square_Hollow
    ePic_Diamond_Hollow
    'digits
    ePic_Zero
    ePic_One
    ePic_Two
    ePic_Three
    ePic_Four
    ePic_Five
    ePic_Six
    ePic_Seven
    ePic_Eight
    ePic_Nine
End Enum

Private Type mPrivate
    aIconList As cGdArray
    aCustomChars As cGdArray
    Chart As cChart
    Annot As cAnnotation
        
    eImage As eStockImage
    eSize As eImageSize
    eStyle As eImageStyle
    eDir As eImageDir
    eImgLabelAlign As eImageLabelAlign
    
    nSelItem As Long
    nColor As Long
    nPreIndicator As Long
    bMultiChart As Boolean
    bWasMultiChart As Boolean
    
    strText As String
    strChar As String
    
    strFont As String
    strFontSize As String
    bBold As Boolean
    bUnderline As Boolean
    bItalic As Boolean
    
    strPicAlphaSolid As String
    strPicAlphaHollow As String
        
    oToolTip As cToolTip
End Type
Private m As mPrivate

Public Sub ShowMe(Chart As cChart, _
    Optional ByVal Annot As cAnnotation = Nothing)
On Error GoTo ErrSection:
    
    Dim eShowMode As eShowFormMode
    
    Set m.Annot = Annot
    Set m.Chart = Chart
        
    InitGridSaved
    LoadIconList
            
    If m.Annot Is Nothing Then
        InitSettings
        cmdClose.Caption = "&Close"
        cmdClose.Left = lblInfoPrompt.Left + lblInfoPrompt.Width
        lblInfoPrompt.Visible = True
        cmdDeleteIcon.Visible = False
        chkAutoIncrement.Enabled = True
        eShowMode = eForm_Nonmodal
    Else
        cmdClose.Caption = "&OK"
        cmdClose.Left = 750
        lblInfoPrompt.Visible = False
        cmdDeleteIcon.Top = cmdClose.Top
        cmdDeleteIcon.Visible = True
        chkAutoIncrement.Enabled = False
        LoadAnnotSettings
        eShowMode = eForm_Modal
    End If
    
    If Not m.Chart Is Nothing Then
        chkMultiChart.Caption = kMultiChartLabel & m.Chart.Symbol
    End If
    
    CenterFormOnChart Me, m.Chart
    ToggleTopMost True
    ShowForm Me, eShowMode
        
    Exit Sub
    
ErrSection:
     RaiseError Me.Name & ".ShowMe"

End Sub

Private Sub chkMultiChart_Click()
On Error GoTo ErrSection:
    
    If chkMultiChart.Enabled Then
        m.bMultiChart = -1 * chkMultiChart.Value
    End If
    UpdateAnnot
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".chkMultiChart_Click"

End Sub

Private Sub chkMultiChart_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:
    
    HandleKeyDown KeyCode

End Sub

Private Sub chkPreIndicator_Click()
On Error GoTo ErrSection:

    m.nPreIndicator = chkPreIndicator.Value
    UpdateAnnot
    
    Exit Sub
    
ErrSection:
     RaiseError Me.Name & ".chkPreIndicator_Click"

End Sub

Private Sub chkPreIndicator_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:
    
    HandleKeyDown KeyCode

End Sub

Private Sub clrColor_Changed()
On Error GoTo ErrSection:

    ToggleTopMost True
    m.nColor = clrColor.Color
    UpdateAnnot
    
    Exit Sub
    
ErrSection:
     RaiseError Me.Name & ".clrColor_Changed"

End Sub

Private Sub clrColor_DropDown()
On Error Resume Next:

    ToggleTopMost False
    tmr.Enabled = True

End Sub

Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Unload Me
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".cmdClose_Click"

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:

    HandleKeyDown KeyCode

End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Dim i&
    
    i = fgSaved.Row
    
    If i >= 0 And i < fgSaved.Rows Then
        i = FindItemInArray(fgSaved.TextMatrix(i, 0))
        If i >= 0 Then
            m.aIconList.Remove i
            m.aIconList.ToFile g.strAppPath & "\custom\IconAnnot.Cfg"
            fgSaved.RemoveItem fgSaved.Row
        End If
    End If

    Exit Sub
    
ErrSection:
     RaiseError Me.Name & ".cmdDelete_Click"

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:
    
    HandleKeyDown KeyCode

End Sub

Private Sub cmdDeleteIcon_Click()
On Error GoTo ErrSection:

    If Not m.Annot Is Nothing And Not m.Chart Is Nothing Then
        m.Annot.geRemoveAnnotation (m.Chart.geChartObj)
        m.Chart.Annots.Remove m.Annot.geAnnId
        m.Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart
    End If
    Unload Me
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".cmdDeleteIcon_Click"

End Sub

Private Sub cmdDeleteIcon_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:
    
    HandleKeyDown KeyCode

End Sub

Private Sub cmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:

    HandleKeyDown KeyCode

End Sub

Private Sub cmdDown_Click()
On Error GoTo ErrSection:

    Dim i&, iRow&, strName$, strIconInfo$
    
    With fgSaved
        iRow = .Row
        If iRow > 0 And iRow < .Rows - 1 Then
            .Redraw = flexRDNone
            strName = .TextMatrix(iRow, 0)
            i = FindItemInArray(.TextMatrix(iRow, 0))
            If i >= 0 And i + 1 < m.aIconList.Size Then
                strIconInfo = m.aIconList(i)
                m.aIconList.Remove i
                m.aIconList.Add strIconInfo, i + 1
                m.aIconList.ToFile g.strAppPath & "\custom\IconAnnot.Cfg"
                .RemoveItem iRow
                .AddItem strName, iRow + 1
                .Row = iRow + 1
                .ShowCell .Row, 0
                If .Row = .Rows - 1 Then cmdDown.Enabled = False
                cmdUp.Enabled = True
            End If
            .Redraw = flexRDBuffered
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmIconAnnot.cmdDown_Click"

End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:
    
    Me.Font.Name = m.strFont
    Me.Font.Size = Val(m.strFontSize)
    Me.Font.Bold = m.bBold
    Me.Font.Italic = m.bItalic
    Me.Font.Underline = m.bItalic
    
    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        m.strFont = Me.Font.Name
        m.strFontSize = Str(Me.Font.Size)
        m.bBold = Me.Font.Bold
        m.bItalic = Me.Font.Italic
        m.bUnderline = Me.Font.Underline
        
        UpdateAnnot
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIconAnnot.cmdFont.Click", eGDRaiseError_Show
    
End Sub

Private Sub cmdRename_Click()
On Error GoTo ErrSection:

    Dim i&, strName$
    Dim aFields As New cGdArray
    Dim iRow&
    
    iRow = fgSaved.Row
    If iRow >= 0 And iRow < fgSaved.Rows Then
        strName = fgSaved.TextMatrix(iRow, 0)
    Else
        InfBox "Please select an item from the list for renaming.", "I", , "Rename Icon Configuration"
        Exit Sub
    End If
      
    strName = InfBox("Enter new name:", "?", , "Rename Icon Configuration", , , , , , "string", strName, eGDAlign_Left)
    strName = Trim(strName)
    
    If Len(strName) > 0 Then
        i = FindItemInArray(fgSaved.TextMatrix(iRow, 0))
        If i >= 0 Then
            aFields.SplitFields (m.aIconList(i)), vbTab
            aFields(0) = strName
            m.aIconList(i) = aFields.JoinFields(vbTab)
            m.aIconList.ToFile g.strAppPath & "\custom\IconAnnot.Cfg"
            fgSaved.TextMatrix(iRow, 0) = strName
        End If
    End If

    Exit Sub
    
ErrSection:
     RaiseError Me.Name & ".cmdRename_Click"

End Sub

Private Sub cmdRename_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:
    
    HandleKeyDown KeyCode

End Sub

Private Sub cmdSaveAs_Click()
On Error GoTo ErrSection:
    
    Dim i&, strName$
    
    If m.Annot Is Nothing Then
        Set m.Annot = New cAnnotation
        m.Annot.CreateNew m.Chart, eANNOT_Icon, 1, 0, 0, 0, 0, , , , , True
        m.Annot.Prop("Custom") = 0
    End If
        
    i = fgSaved.Row
    If i > 0 And i < fgSaved.Rows Then
        strName = fgSaved.TextMatrix(i, 0)
    End If
    
    UpdateAnnot
    strName = m.Annot.SaveIconToString(strName)
    If Len(strName) > 0 Then
        LoadIconList strName
        fgSaved.Row = i
    End If
    
    Exit Sub
    
ErrSection:
     RaiseError Me.Name & ".cmdSaveAs_Click"

End Sub

Private Sub cmdSaveAs_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:
    
    HandleKeyDown KeyCode

End Sub

Private Sub cmdUp_Click()
On Error GoTo ErrSection:

    Dim i&, iRow&, strName$, strIconInfo$
    
    With fgSaved
        iRow = .Row
        If iRow > 1 And iRow < .Rows Then
            .Redraw = flexRDNone
            strName = .TextMatrix(iRow, 0)
            i = FindItemInArray(.TextMatrix(iRow, 0))
            If i > 0 And i < m.aIconList.Size And i - 1 >= 0 Then
                strIconInfo = m.aIconList(i)
                m.aIconList.Remove i
                m.aIconList.Add strIconInfo, i - 1
                m.aIconList.ToFile g.strAppPath & "\custom\IconAnnot.Cfg"
                .RemoveItem iRow
                .AddItem strName, iRow - 1
                .Row = iRow - 1
                If .Row = 1 Then
                    cmdUp.Enabled = False
                    .ShowCell 0, 0      'so user will not have to scroll to see they are at the top
                Else
                    .ShowCell .Row, 0
                End If
            End If
            .Redraw = flexRDBuffered
        End If
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmIconAnnot.cmdDown_Click"

End Sub

Private Sub fg_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
On Error GoTo ErrSection:

    Dim nIdx&
    
    nIdx = GridToPicIdx(NewRowSel, NewColSel)
    PicIdxToImage nIdx, m.eImage, m.eStyle, m.eDir, m.strChar
    UpdateAnnot
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".fg_AfterSelChange"

End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:
    
    If KeyCode = vbKeyF1 Then g.Help.ShowF1Help Nothing

End Sub

Private Sub fgSaved_Click()
On Error GoTo ErrSection:

    Dim bEnable As Boolean, i&
    
    i = fgSaved.Row
    If i > 0 And i < fgSaved.Rows Then
        cmdSaveAs.Caption = "&Save"
        bEnable = True
        If cmdClose.Caption = "&OK" And Not m.Annot Is Nothing Then
            m.Annot.Prop("AsciiChar") = ""
            m.Annot.SetIconFromString m.aIconList, i - 1
            LoadAnnotSettings
            If Not m.Chart Is Nothing Then
                m.Chart.GenerateChart eRedo1_Scrolled   'redraw the annotation
            End If
        Else
            Set m.Annot = Nothing
            Set m.Annot = New cAnnotation
            m.Annot.CreateNew m.Chart, eANNOT_Icon, 1, 0, 0, 0, 0, , , , , True
            m.Annot.SetIconFromString m.aIconList, i - 1
            LoadAnnotSettings
        End If
        
        If i > 1 Then
            cmdUp.Enabled = True
        Else
            cmdUp.Enabled = False
        End If
        If i < fgSaved.Rows - 1 Then
            cmdDown.Enabled = True
        Else
            cmdDown.Enabled = False
        End If
        
    Else
        cmdSaveAs.Caption = "&New"
        bEnable = False
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    
    cmdDelete.Enabled = bEnable
    cmdRename.Enabled = bEnable

ErrExit:
    Exit Sub
    
ErrSection:
     RaiseError "frmIconAnnot.fgSaved_Click"
     
End Sub

Private Sub fgSaved_KeyUp(KeyCode As Integer, Shift As Integer)
    fgSaved_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:
    
    HandleKeyDown KeyCode

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim s$, i&

    Me.Icon = Picture16(ToolbarIcon("ID_Icon"))
    
    g.Styler.StyleForm Me
        
    s = "Click an icon to select it. "
    s = s & "Double click icons in the last row "
    s = s & "to set custom characters."
    
    If m.oToolTip Is Nothing Then Set m.oToolTip = New cToolTip
    m.oToolTip.Create Me
    
    For i = 0 To ePic_Nine
        m.oToolTip.AddTool pic(i), s
    Next
    
    s = "Double click to set custom characters."
    For i = 30 To 39
        m.oToolTip.AddTool pic(i), s
    Next
    
    InitIconGrid
    
    s = GetIniFileProperty(Me.Name, "", "Placement", g.strIniFile)
    If Len(s) > 0 Then
        SetFormPlacement Me, s, "LT"
    Else
        CenterTheForm Me
    End If
    
    Exit Sub
    
ErrSection:
     RaiseError Me.Name & ".Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   
    Dim s$
    
    s = m.aCustomChars.JoinFields(",")
    
    Set m.aIconList = Nothing
    Set m.aCustomChars = Nothing
    Set m.oToolTip = Nothing
    tmr.Enabled = False
    m.nSelItem = -1
    
    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "CustomIconChars", s, "Charting", g.strIniFile
    
    Me.Hide
    
    If Not g.bUnloading And Not m.Chart Is Nothing Then
        If cmdClose.Caption = "&OK" Then
            If Not m.Annot Is Nothing And Not m.Chart Is Nothing Then
                m.Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart
            End If
        End If
        m.Chart.SetCursor
        If g.strActiveDraw = "" Or g.strActiveDraw = "ID_Icon" Then
            ToolbarSetCursorGroup m.Chart.tbToolbar, False
            If Not m.Chart.Form Is Nothing Then m.Chart.Form.SyncDrawTools      '5141
        Else
            'user clicked on another drawtool
            ToolbarSetCursorGroup m.Chart.tbToolbar, True
        End If
    End If

End Sub

Private Function FindItemInArray(ByVal strName$) As Long
On Error GoTo ErrSection:

    Dim i&, nIndex&
    Dim aFields As New cGdArray
    
    nIndex = -1
    
    For i = 0 To m.aIconList.Size
        aFields.SplitFields m.aIconList(i), vbTab
        If aFields(0) = strName Then
            nIndex = i
            Exit For
        End If
    Next

    FindItemInArray = nIndex

    Exit Function
    
ErrSection:
     RaiseError Me.Name & ".FindItemInArray"

End Function

Private Sub LoadGridSaved(Optional ByVal strName$ = "")
On Error GoTo ErrSection:

    Dim i&, j&
    Dim aFields As New cGdArray
    
    j = -1
    
    With fgSaved
        .Rows = 1
        .TextMatrix(0, 0) = "<New>"
        If m.aIconList.Size > 0 Then
            For i = 0 To m.aIconList.Size - 1
                .Rows = .Rows + 1
                aFields.SplitFields m.aIconList(i), vbTab
                .TextMatrix(.Rows - 1, 0) = aFields(0)
                If strName = aFields(0) Then j = i
            Next
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIconAnnot.LoadGridSaved"

End Sub

Private Sub InitIconGrid()
On Error GoTo ErrSection:
    
    Dim nBoxSize&, nBoxSpacing&, nCellSize&, i&, s$
    Dim Point As POINTAPI

    s = GetIniFileProperty("CustomIconChars", "*,?,A,B,C,D,E,F,G,H", "Charting", g.strIniFile)
    ' defaults
    nBoxSize = RoundTwips(kBoxSize)
    nBoxSpacing = RoundTwips(kBoxSpacing)
    nCellSize = nBoxSize + nBoxSpacing * 2

    fg.Rows = kRows
    fg.Cols = kCols
    fg.ColWidth(-1) = nCellSize
    fg.RowHeight(-1) = nCellSize
    fg.Move fg.Left, fg.Top, _
        nCellSize * kCols + 2 * Screen.TwipsPerPixelX, _
        nCellSize * fg.Rows + 2 * Screen.TwipsPerPixelY
        
    'position pic boxes
    For i = 0 To kMax
        pic(i).Move (i Mod kCols) * nCellSize + nBoxSpacing, _
            (i \ kCols) * nCellSize + nBoxSpacing, _
            nBoxSize, nBoxSize
    Next
    
    'numbers
    For i = 20 To 29
        pic(i).ScaleMode = 3
        pic(i).Font.Bold = True
        pic(i).Font.Size = 10
        geDrawText pic(i).hDC, 2, -1, 0, Str(i - 20)
    Next
    
    'custom characters
    If m.aCustomChars Is Nothing Then
        Set m.aCustomChars = New cGdArray
    End If
    
    m.aCustomChars.Size = 0
    If Len(s) > 0 Then m.aCustomChars.SplitFields s, ","
    If m.aCustomChars.Size <> 10 Then
        'want to make sure size is exactly 10 in case INI file got corrupted
        m.aCustomChars.Size = 0
        m.aCustomChars(0) = "*"
        m.aCustomChars(1) = "?"
        m.aCustomChars(2) = "A"
        m.aCustomChars(3) = "B"
        m.aCustomChars(4) = "C"
        m.aCustomChars(5) = "D"
        m.aCustomChars(6) = "E"
        m.aCustomChars(7) = "F"
        m.aCustomChars(8) = "G"
        m.aCustomChars(9) = "H"
    End If
    
    For i = 30 To 39
        pic(i).ScaleMode = 3
        pic(i).Font.Bold = True
        pic(i).Font.Size = 10
    Next
    geDrawText pic(30).hDC, 2, -1, 0, m.aCustomChars(0)
    geDrawText pic(31).hDC, 2, -1, 0, m.aCustomChars(1)
    geDrawText pic(32).hDC, 2, -1, 0, m.aCustomChars(2)
    geDrawText pic(33).hDC, 2, -1, 0, m.aCustomChars(3)
    geDrawText pic(34).hDC, 2, -1, 0, m.aCustomChars(4)
    geDrawText pic(35).hDC, 2, -1, 0, m.aCustomChars(5)
    geDrawText pic(36).hDC, 2, -1, 0, m.aCustomChars(6)
    geDrawText pic(37).hDC, 2, -1, 0, m.aCustomChars(7)
    geDrawText pic(38).hDC, 2, -1, 0, m.aCustomChars(8)
    geDrawText pic(39).hDC, 2, -1, 0, m.aCustomChars(9)
                
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".InitIconGrid"

End Sub

Private Sub optCenter_Click()
On Error Resume Next:

    m.eImgLabelAlign = eImgLblAlign_Center
    UpdateAnnot

End Sub

Private Sub optCenter_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:

    HandleKeyDown KeyCode

End Sub

Private Sub optLarge_Click()
On Error Resume Next:

    m.eSize = eImgLarge
    UpdateAnnot

End Sub

Private Sub optLarge_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:

    HandleKeyDown KeyCode

End Sub

Private Sub optLeft_Click()
On Error Resume Next:

    m.eImgLabelAlign = eImgLblAlign_Left
    UpdateAnnot

End Sub

Private Sub optLeft_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:
    
    HandleKeyDown KeyCode

End Sub

Private Sub optMedium_Click()
On Error Resume Next:

    m.eSize = eImgMedium
    UpdateAnnot

End Sub

Private Sub optMedium_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:

    HandleKeyDown KeyCode

End Sub

Private Sub optRight_Click()
On Error Resume Next:

    m.eImgLabelAlign = eImgLblAlign_Right
    UpdateAnnot

End Sub

Private Sub optRight_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:

    HandleKeyDown KeyCode

End Sub

Private Sub optSmall_Click()
On Error Resume Next:

    m.eSize = eImgSmall
    UpdateAnnot

End Sub

Private Sub optSmall_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next:

    HandleKeyDown KeyCode

End Sub

Private Sub pic_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim nRow&, nCol&
    
    PicIdxToGrid Index, nRow, nCol
    
    If nRow >= 0 And nRow < fg.Rows And nCol >= 0 And nCol < fg.Cols Then
        fg.Select nRow, nCol
    End If
    
    PicIdxToImage Index, m.eImage, m.eStyle, m.eDir, m.strChar
    
'    If Index > ePic_Diamond_Hollow Then
'        If optSmall.Value = True Then
'            optSmall.Value = False
'            optMedium.Value = True
'        End If
'        optSmall.Enabled = False
'    Else
'        optSmall.Enabled = True
'    End If
    
    UpdateAnnot
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".pic_Click"

End Sub

Private Sub InitSettings()
On Error GoTo ErrSection:

    m.eImage = eCNI_Arrow
    m.eSize = eImgMedium
    m.eStyle = eImgHollow
    m.eDir = eCNI_North
    m.eImgLabelAlign = eImgLblAlign_Center
    m.nPreIndicator = 0
    m.bMultiChart = False
    m.bWasMultiChart = False
    m.nColor = clrColor.Color
    m.strText = ""
    
    m.strFont = "Arial"
    m.strFontSize = "8"
    m.bBold = False
    m.bItalic = False
    m.bUnderline = False
    
    fg.Row = 0
    fg.Col = 0
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".InitSettings"

End Sub

Private Sub LoadAnnotSettings()
On Error GoTo ErrSection:

    Dim nIdx&, nRow&, nCol&, i&
    
    Dim bShowMutltichart As Boolean

    m.eImage = Val(m.Annot.Prop("ImageType"))
    m.eSize = Val(m.Annot.Prop("ImageSize"))
    m.eStyle = Val(m.Annot.Prop("ImageStyle"))
    m.eDir = Val(m.Annot.Prop("ImageDir"))
    m.strChar = m.Annot.Prop("AsciiChar")
    
    m.eImgLabelAlign = m.Annot.geTextAlign
    m.nPreIndicator = m.Annot.PreIndicator
    
    If lblInfoPrompt.Visible Then
        bShowMutltichart = True
    ElseIf Not m.Chart Is Nothing Then
        If Not m.Chart.Tree(m.Annot.gePaneId) Is Nothing Then
            '5514 - show on all chart is only for annotations in Price Pane
            bShowMutltichart = m.Chart.Tree(m.Annot.gePaneId).PricePaneFlag
        End If
    End If
    
    chkMultiChart.Enabled = bShowMutltichart
    chkMultiChart.Visible = bShowMutltichart
    If bShowMutltichart Then
        m.bMultiChart = m.Annot.MultiChartFlag
    Else
        m.bMultiChart = False
    End If
    
    'backwards compatibility:
    '   original implementation has solid/hollow circles around letters/numbers
    If Len(m.strChar) > 0 Then
        m.eImage = eCNI_Ascii
        m.eStyle = eImgHollow
    End If
    
    'set font currently in use
    m.strFont = m.Annot.Prop("FontName")
    m.strFontSize = m.Annot.Prop("FontSize")
    m.bUnderline = Val(m.Annot.Prop("FontUnderline"))
    i = Val(m.Annot.Prop("FontStyle"))
    Select Case i
        Case 0:
            m.bItalic = False
            m.bBold = False
        Case 1:
            m.bItalic = False
            m.bBold = True
        Case 2:
            m.bItalic = True
            m.bBold = False
        Case 3:
            m.bItalic = True
            m.bBold = True
    End Select
    
    'was multichart flag tells chart object whether to remove annots from other charts
    m.bWasMultiChart = m.bMultiChart
    
    m.nColor = m.Annot.Color
    m.strText = m.Annot.Text
    
    nIdx = ImageToPicIdx(m.eImage, m.eStyle, m.eDir, m.strChar)
    PicIdxToGrid nIdx, nRow, nCol
    'set the pic box to match annot's char
    If nIdx > 29 Then
        pic(nIdx).Cls
        geDrawText pic(nIdx).hDC, 2, -1, 0, m.strChar
        i = nIdx - 30
        If i >= 0 And i < m.aCustomChars.Size Then
            m.aCustomChars(i) = m.strChar
        End If
    End If
        
    clrColor.Color = m.Annot.Color
    chkPreIndicator.Value = m.nPreIndicator
    chkMultiChart.Value = Abs(m.bMultiChart)
    If Not m.Chart Is Nothing Then
        chkMultiChart.Caption = kMultiChartLabel & m.Chart.Symbol
    End If
    
    Select Case m.eSize
        Case eImgSmall:
            optSmall.Value = True
        Case eImgMedium:
            optMedium.Value = True
        Case eImgLarge:
            optLarge.Value = True
        Case Else:
            optSmall.Value = True
    End Select
    
    'disable "small" for circled numbers/alpha
'    If nIdx > 19 Then
'        optSmall.Enabled = False
'    Else
'        optSmall.Enabled = True
'    End If
        
    'set icon selection
    If nRow >= 0 And nRow < fg.Rows And nCol >= 0 And nCol < fg.Cols Then
        fg.Select nRow, nCol
    End If
    
    'set label info
    txtCustom.Text = m.strText
        
    'set alignment info
    If m.eImgLabelAlign = eImgLblAlign_Auto Or m.eImgLabelAlign = eImgLblAlign_Left Then
        optLeft.Value = True
    ElseIf m.eImgLabelAlign = eImgLblAlign_Center Then
        optCenter.Value = True
    ElseIf m.eImgLabelAlign = eImgLblAlign_Right Then
        optRight.Value = True
    Else
        optCenter.Value = True
    End If
    
    Exit Sub
    
ErrSection:
     RaiseError Me.Name & ".LoadAnnotSettings"

End Sub

Private Function AsciiToPicIdx(ByVal eStyle As eImageStyle, _
    ByVal strAsciiChar$) As Long
On Error GoTo ErrSection:

    Dim eIdx As ePicIdx, i&
    
    Select Case strAsciiChar
        Case "0"
            eIdx = ePic_Zero
        Case "1"
            eIdx = ePic_One
        Case "2"
            eIdx = ePic_Two
        Case "3"
            eIdx = ePic_Three
        Case "4"
            eIdx = ePic_Four
        Case "5"
            eIdx = ePic_Five
        Case "6"
            eIdx = ePic_Six
        Case "7"
            eIdx = ePic_Seven
        Case "8"
            eIdx = ePic_Eight
        Case "9"
            eIdx = ePic_Nine
        Case Else
            For i = 0 To m.aCustomChars.Size - 1
                If m.aCustomChars(i) = strAsciiChar Then
                    eIdx = 30 + i
                    Exit For
                End If
            Next
            If eIdx > kLastPicIndex Then
                eIdx = kLastPicIndex
            ElseIf eIdx < 30 Then
                eIdx = 30
            End If
    End Select
    
    AsciiToPicIdx = eIdx
    
    Exit Function

ErrSection:
     RaiseError Me.Name & ".AsciiToPicIdx"

End Function

Private Function ImageToPicIdx(ByVal eBase As eStockImage, _
    ByVal eStyle As eImageStyle, ByVal eDir As eImageDir, _
    ByVal strAsciiChar As String) As Long
On Error GoTo ErrSection:

    Dim ePicIndex As ePicIdx
    
    ePicIndex = ePic_ArrowNorth
    
    If Len(strAsciiChar) > 0 Then
        ImageToPicIdx = AsciiToPicIdx(eStyle, strAsciiChar)
        Exit Function
    End If
    
    Select Case eBase
        Case eCNI_Arrow
            Select Case eDir
                Case eCNI_North
                    ePicIndex = ePic_ArrowNorth
                Case eCNI_South
                    ePicIndex = ePic_ArrowSouth
                Case eCNI_East
                    ePicIndex = ePic_ArrowEast
                Case eCNI_West
                    ePicIndex = ePic_ArrowWest
                Case eCNI_NorthEast
                    ePicIndex = ePic_ArrowNorthEast
                Case eCNI_SouthEast
                    ePicIndex = ePic_ArrowSouthEast
                Case eCNI_NorthWest
                    ePicIndex = ePic_ArrowNorthWest
                Case eCNI_SouthWest
                    ePicIndex = ePic_ArrowSouthWest
                Case Else
                    ePicIndex = ePic_ArrowNorth
            End Select
        Case eCNI_Plus
            ePicIndex = ePic_Plus
        Case eCNI_Cross
            ePicIndex = ePic_Cross
        Case eCNI_Circle
            If eStyle = eImgFilled Then
                ePicIndex = ePic_Circle_Filled
            Else
                ePicIndex = ePic_Circle_Hollow
            End If
        Case eCNI_Square
            If eStyle = eImgFilled Then
                ePicIndex = ePic_Square_Filled
            Else
                ePicIndex = ePic_Square_Hollow
            End If
        Case eCNI_Diamond
            If eStyle = eImgFilled Then
                ePicIndex = ePic_Diamond_Filled
            Else
                ePicIndex = ePic_Diamond_Hollow
            End If
        Case eCNI_Triangle
            If eDir = eCNI_North Then
                If eStyle = eImgFilled Then
                    ePicIndex = ePic_TriUp_Filled
                Else
                    ePicIndex = ePic_TriUp_Hollow
                End If
            Else
                If eStyle = eImgFilled Then
                    ePicIndex = ePic_TriDown_Filled
                Else
                    ePicIndex = ePic_TriDown_Hollow
                End If
            End If
        Case Else
            ePicIndex = ePic_ArrowNorth
    End Select
    
    ImageToPicIdx = ePicIndex
    
    Exit Function

ErrSection:
     RaiseError Me.Name & ".ImageToPicIdx"

End Function

Private Sub PicIdxToGrid(ByVal nIdx&, nRow&, nCol&)
On Error GoTo ErrSection:

    Dim nRemainder&
    
    If nIdx < 0 Or nIdx > kLastPicIndex Then
        nRow = -1
        nCol = -1
    Else
        nRow = Int(nIdx / 10)
        nCol = (nIdx Mod 10)
    End If
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".PicIdxToGrid"

End Sub

Private Function GridToPicIdx(ByVal nRow&, nCol&) As Long
On Error GoTo ErrSection:

    GridToPicIdx = nRow * 10 + nCol

    Exit Function

ErrSection:
     RaiseError Me.Name & ".GridToPicIdx"

End Function

Private Sub UpdateAnnot()
On Error GoTo ErrSection:

    Dim nStyle&

    If m.Annot Is Nothing Then Exit Sub
    
    With m.Annot
        .Color = m.nColor
        .Text = m.strText
        .PreIndicator = m.nPreIndicator
        .MultiChartFlag = m.bMultiChart
        
        .Prop("ImageType") = m.eImage
        .Prop("ImageSize") = m.eSize
        .Prop("ImageStyle") = m.eStyle
        .Prop("ImageDir") = m.eDir
        .Prop("AsciiChar") = m.strChar
        
        .Prop("FontName") = m.strFont
        .Prop("FontSize") = m.strFontSize
        .Prop("FontUnderline") = m.bUnderline
        'style - 0=reg,1=bold,2=italic,3=bold italic
        nStyle = 0
        If m.bBold = True Then
            If m.bItalic = True Then
                nStyle = 3
            Else
                nStyle = 1
            End If
        ElseIf m.bItalic = True Then
            nStyle = 2
        End If
        
        .Prop("FontStyle") = nStyle
        
        .geTextAlign = m.eImgLabelAlign
    End With
    
    If cmdClose.Caption = "&OK" And Not m.Chart Is Nothing Then
        m.Chart.GenerateChart eRedo1_Scrolled
    End If

    Exit Sub

ErrSection:
     RaiseError Me.Name & ".UpdateAnnot"
     
End Sub

Private Sub PicIdxToImage(ByVal nIdx&, eBase As eStockImage, _
    eStyle As eImageStyle, eDir As eImageDir, strChar As String)
On Error GoTo ErrSection:
    
    Dim ePic As ePicIdx
    Dim nRow&, nCol&, i&
    
    'custom characters
    If nIdx > 29 Then
        eBase = eCNI_Ascii
        eStyle = eImgHollow
        i = nIdx - 30
        If i >= 0 And i < m.aCustomChars.Size Then
            strChar = m.aCustomChars(i)
        End If
        Exit Sub
    End If
    
    ePic = nIdx
    PicIdxToGrid nIdx, nRow, nCol
    
    If ePic = ePic_TriUp_Filled Or ePic = ePic_TriDown_Filled Or _
       ePic = ePic_Circle_Filled Or ePic = ePic_Square_Filled Or _
       ePic = ePic_Diamond_Filled Then
        eStyle = eImgFilled
    Else
        eStyle = eImgHollow
    End If
    
    If ePic < ePic_Zero Then
        strChar = ""
    Else
        eBase = eCNI_Ascii
    End If
    
    Select Case ePic
        Case ePic_ArrowNorth, ePic_ArrowSouth, ePic_ArrowEast, ePic_ArrowWest, _
            ePic_ArrowNorthEast, ePic_ArrowNorthWest, ePic_ArrowSouthEast, ePic_ArrowSouthWest
            eBase = eCNI_Arrow
            eDir = ArrowDir(ePic)
        
        Case ePic_Plus
            eBase = eCNI_Plus
        
        Case ePic_Cross
            eBase = eCNI_Cross
        
        Case ePic_TriUp_Filled, ePic_TriUp_Hollow
            eBase = eCNI_Triangle
            eDir = eCNI_North
        
        Case ePic_TriDown_Hollow, ePic_TriDown_Filled
            eBase = eCNI_Triangle
            eDir = eCNI_South
        
        Case ePic_Circle_Filled, ePic_Circle_Hollow
            eBase = eCNI_Circle
            
        Case ePic_Square_Filled, ePic_Square_Hollow
            eBase = eCNI_Square
        
        Case ePic_Diamond_Filled, ePic_Diamond_Hollow
            eBase = eCNI_Diamond

        Case ePic_Zero, _
             ePic_One, _
             ePic_Two, _
             ePic_Three, _
             ePic_Four, _
             ePic_Five, _
             ePic_Six, _
             ePic_Seven, _
             ePic_Eight, _
             ePic_Nine
             
            strChar = Str(nCol)
    End Select
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".PicIdxToImage"

End Sub

Private Function ArrowDir(ByVal eIndex As ePicIdx) As eImageDir
On Error GoTo ErrSection:
    
    Dim eDir As eImageDir
    
    Select Case eIndex
        Case ePic_ArrowNorth
            eDir = eCNI_North
        Case ePic_ArrowSouth
            eDir = eCNI_South
        Case ePic_ArrowEast
            eDir = eCNI_East
        Case ePic_ArrowWest
            eDir = eCNI_West
        Case ePic_ArrowNorthEast
            eDir = eCNI_NorthEast
        Case ePic_ArrowNorthWest
            eDir = eCNI_NorthWest
        Case ePic_ArrowSouthEast
            eDir = eCNI_SouthEast
        Case ePic_ArrowSouthWest
            eDir = eCNI_SouthWest
        Case Else
            eDir = eCNI_North
    End Select

    ArrowDir = eDir

    Exit Function

ErrSection:
     RaiseError Me.Name & ".ArrowDir"

End Function

Public Property Let chartObj(Chart As cChart)
On Error GoTo ErrSection:

    Dim strSymPrev$

    If Not m.Chart Is Nothing Then strSymPrev = m.Chart.Symbol
    
    Set m.Chart = Nothing
    Set m.Chart = Chart
    
    If m.Chart.Symbol <> strSymPrev Then
        chkMultiChart.Caption = kMultiChartLabel & m.Chart.Symbol
    End If
    
    Exit Property

ErrSection:
     RaiseError Me.Name & ".ChartObj"

End Property

Private Sub LoadIconList(Optional ByVal strName$ = "")
On Error GoTo ErrSection:
        
    Dim strFile$, s$, i&, j&
    Dim aTemp As New cGdArray
    Dim aFields As New cGdArray
    Dim bFound As Boolean
    Dim bAdded As Boolean
    
    If m.aIconList Is Nothing Then Set m.aIconList = New cGdArray
    
    m.aIconList.Size = 0
    m.aIconList.FromFile g.strAppPath & "\custom\IconAnnot.cfg"
    
    'check for upgrade add icon file
    strFile = g.strAppPath & "\provided\IconAnnot.add"
    If FileExist(strFile) Then
        aTemp.FromFile strFile
        For i = 0 To aTemp.Size - 1
            bFound = False
            aFields.SplitFields aTemp(i), vbTab
            s = aFields(0)
            For j = 0 To m.aIconList.Size - 1
                aFields.SplitFields m.aIconList(j), vbTab
                If s = aFields(0) Then
                    bFound = True
                    Exit For
                End If
            Next
            If Not bFound Then
                m.aIconList.Add aTemp(i)
                bAdded = True
            End If
        Next
        If bAdded Then m.aIconList.ToFile g.strAppPath & "\custom\IconAnnot.cfg"
        KillFile strFile
    End If
    
    LoadGridSaved strName

    Exit Sub

ErrSection:
     RaiseError Me.Name & ".LoadIconList"

End Sub

Private Sub pic_DblClick(Index As Integer)
On Error GoTo ErrSection:

    Dim rtrn$, strAscii$, i&
    
    If Index > 29 Then
        rtrn = InfBox("Please enter a single character ...", "?", , "Custom character", , , , , , "s", strAscii)
    
        If Len(rtrn) > 0 Then
            strAscii = Left(rtrn, 1)
            If Len(strAscii) = 1 Then
                pic(Index).Cls
                geDrawText pic(Index).hDC, 2, -1, 0, strAscii
                m.strChar = strAscii
                i = Index - 30
                If i >= 0 And i < m.aCustomChars.Size Then
                    m.aCustomChars(i) = strAscii
                End If
                UpdateAnnot
            End If
        End If
    End If

    Exit Sub

ErrSection:
     RaiseError Me.Name & ".pic_DblClick"

End Sub

Private Sub pic_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
        
    Dim nRow&, nCol&
    
    If KeyCode = vbKeyLeft Or _
       KeyCode = vbKeyRight Or _
       KeyCode = vbKeyUp Or _
       KeyCode = vbKeyDown Then
        PicIdxToGrid Index, nRow, nCol
        If nRow >= 0 And nRow < fg.Rows And nCol >= 0 And nCol < fg.Cols Then
            fg.Row = nRow
            fg.Col = nCol
        End If
    End If
    
    HandleKeyDown KeyCode
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".pic_KeyDown"

End Sub

Private Sub tmr_Timer()
On Error GoTo ErrSection:

    m.nColor = clrColor.Color
    UpdateAnnot
    tmr.Enabled = False
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".tmr_Timer"

End Sub

Private Sub txtCustom_Change()
On Error Resume Next:

    If txtCustom.Text <> m.strText Then
        m.strText = txtCustom.Text
        UpdateAnnot
    End If

End Sub

Public Sub GetSettings(Annot As cAnnotation)
On Error GoTo ErrSection:

    Set m.Annot = Annot
    UpdateAnnot
    Set m.Annot = Nothing
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".GetSettings"

End Sub

Private Sub HandleKeyDown(KeyCode As Integer)
On Error GoTo ErrSection:
        
    Dim nRow&, nCol&
    
    nRow = fg.Row
    nCol = fg.Col
        
    If KeyCode = vbKeyLeft And nCol > 0 Then
        nCol = nCol - 1
    ElseIf KeyCode = vbKeyRight And fg.Col < fg.Cols - 1 Then
        nCol = fg.Col + 1
    ElseIf KeyCode = vbKeyDown And fg.Row < fg.Rows - 1 Then
        nRow = fg.Row + 1
    ElseIf KeyCode = vbKeyUp And fg.Row > 0 Then
        nRow = fg.Row - 1
    ElseIf KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Nothing
    End If
    
    fg.Select nRow, nCol
    fg.SetFocus
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".HandleKeyDown"

End Sub

Public Sub NextIcon()
On Error GoTo ErrSection:

    Dim nIdx As Integer
    
    If chkAutoIncrement.Value = 1 Then
        nIdx = ImageToPicIdx(m.eImage, m.eStyle, m.eDir, m.strChar)
        If nIdx < kMax Then
            nIdx = nIdx + 1
        Else
            nIdx = 0
        End If
        pic_Click nIdx
    End If
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".NextIcon"

End Sub

Private Sub InitGridSaved()
On Error GoTo ErrSection:

    With fgSaved
        .Redraw = flexRDNone
        SetupGrid Me.fgSaved, eGridMode_Grid
        .ExplorerBar = flexExNone
        .FixedCols = 0
        .Editable = flexEDNone
        .FixedRows = 0
        .Rows = 0
        .Cols = 1
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIconAnnot.InitGridSaved", eGDRaiseError_Raise

End Sub

Public Sub ToggleMultichart(ByVal IsPricePane As Boolean)
On Error GoTo ErrSection:

    chkMultiChart.Enabled = IsPricePane
    If IsPricePane Then
        chkMultiChart.Value = Abs(m.bMultiChart)
    Else
        chkMultiChart.Value = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIconAnnot.ToggleMultichart", eGDRaiseError_Raise

End Sub

Private Sub ToggleTopMost(ByVal bTopMost As Boolean)
On Error GoTo ErrSection:

    If Not m.Chart Is Nothing Then
        If Not m.Chart.Form Is Nothing Then
            If m.Chart.Form.DetachStatus = eDetached Then
                SetFormTopmost Me, bTopMost             '6441
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmIconAnnot.ToggleTopMost", eGDRaiseError_Raise

End Sub

