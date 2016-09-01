VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTradeReportFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniCheckXP chkCalcPNL 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4860
      Width           =   7275
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
      Caption         =   "frmTradeReportFilter.frx":0000
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmTradeReportFilter.frx":00EA
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmTradeReportFilter.frx":010A
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraTradeFilter 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   7455
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
      Caption         =   "frmTradeReportFilter.frx":0126
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeReportFilter.frx":0178
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeReportFilter.frx":0198
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraRealSim 
         Height          =   555
         Left            =   480
         TabIndex        =   22
         Top             =   840
         Width           =   5775
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
         Caption         =   "frmTradeReportFilter.frx":01B4
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTradeReportFilter.frx":01F6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":0216
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optSim 
            Height          =   195
            Left            =   4140
            TabIndex        =   25
            Top             =   240
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
            Caption         =   "frmTradeReportFilter.frx":0232
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTradeReportFilter.frx":0264
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTradeReportFilter.frx":0284
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optReal 
            Height          =   195
            Left            =   2040
            TabIndex        =   24
            Top             =   240
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
            Caption         =   "frmTradeReportFilter.frx":02A0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTradeReportFilter.frx":02D4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTradeReportFilter.frx":02F4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optAllFlags 
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   240
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
            Caption         =   "frmTradeReportFilter.frx":0310
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmTradeReportFilter.frx":0338
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTradeReportFilter.frx":0358
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdManageRules 
         Height          =   495
         Left            =   6060
         TabIndex        =   21
         Top             =   360
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
         Caption         =   "frmTradeReportFilter.frx":0374
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":03AE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":03CE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkExitRule 
         Height          =   220
         Left            =   120
         TabIndex        =   19
         Top             =   720
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
         Caption         =   "frmTradeReportFilter.frx":03EA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":0420
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":0440
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboExitRules 
         Height          =   315
         Left            =   1380
         TabIndex        =   20
         Top             =   660
         Width           =   4575
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
         Tip             =   "frmTradeReportFilter.frx":045C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":047C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkEntryRule 
         Height          =   220
         Left            =   120
         TabIndex        =   17
         Top             =   300
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
         Caption         =   "frmTradeReportFilter.frx":0498
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":04D0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":04F0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboEntryRules 
         Height          =   315
         Left            =   1380
         TabIndex        =   18
         Top             =   240
         Width           =   4575
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
         Tip             =   "frmTradeReportFilter.frx":050C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":052C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraGeneral 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6075
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
      Caption         =   "frmTradeReportFilter.frx":0548
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeReportFilter.frx":0588
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeReportFilter.frx":05A8
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fgAccounts 
         Height          =   1215
         Left            =   1380
         TabIndex        =   6
         Top             =   660
         Width           =   4515
         _cx             =   7964
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
      Begin HexUniControls.ctlUniComboImageXP cboCategory 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Top             =   2340
         Width           =   4515
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
         Tip             =   "frmTradeReportFilter.frx":05C4
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":05E4
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkCategory 
         Height          =   220
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "frmTradeReportFilter.frx":0600
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":0634
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":0654
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraLongShort 
         Height          =   555
         Left            =   120
         TabIndex        =   12
         Top             =   2700
         Width           =   5775
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
         Caption         =   "frmTradeReportFilter.frx":0670
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTradeReportFilter.frx":06A4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":06C4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optBoth 
            Height          =   220
            Left            =   240
            TabIndex        =   13
            Top             =   240
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
            Caption         =   "frmTradeReportFilter.frx":06E0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmTradeReportFilter.frx":070A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTradeReportFilter.frx":072A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optLongs 
            Height          =   220
            Left            =   2040
            TabIndex        =   14
            Top             =   240
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
            Caption         =   "frmTradeReportFilter.frx":0746
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTradeReportFilter.frx":077C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTradeReportFilter.frx":079C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optShorts 
            Height          =   220
            Left            =   4140
            TabIndex        =   15
            Top             =   240
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
            Caption         =   "frmTradeReportFilter.frx":07B8
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTradeReportFilter.frx":07F0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTradeReportFilter.frx":0810
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniCheckXP chkAccount 
         Height          =   220
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "frmTradeReportFilter.frx":082C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":085E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":087E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDateRange 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   270
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
         Caption         =   "frmTradeReportFilter.frx":089A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":08D2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":08F2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSymbol 
         Height          =   220
         Left            =   120
         TabIndex        =   7
         Top             =   1980
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "frmTradeReportFilter.frx":090E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":093E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":095E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdFromDate 
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
      End
      Begin gdOCX.gdSelectDate gdToDate 
         Height          =   315
         Left            =   3780
         TabIndex        =   4
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
         Height          =   255
         Left            =   2550
         TabIndex        =   9
         Top             =   1950
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
         Caption         =   "frmTradeReportFilter.frx":097A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":09AC
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":09CC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   315
         Left            =   1380
         TabIndex        =   8
         Top             =   1920
         Width           =   1440
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTradeReportFilter.frx":09E8
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
         Tip             =   "frmTradeReportFilter.frx":0A1C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":0A3C
      End
      Begin HexUniControls.ctlUniLabelXP lblTo 
         Height          =   255
         Left            =   3540
         Top             =   270
         Width           =   195
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
         Caption         =   "frmTradeReportFilter.frx":0A58
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeReportFilter.frx":0A7C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":0A9C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1815
      Left            =   6360
      TabIndex        =   26
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
      Caption         =   "frmTradeReportFilter.frx":0AB8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeReportFilter.frx":0AE4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeReportFilter.frx":0B04
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdTrades 
         Height          =   495
         Left            =   0
         TabIndex        =   29
         Top             =   1320
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
         Caption         =   "frmTradeReportFilter.frx":0B20
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":0B5A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":0B7A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   0
         TabIndex        =   28
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
         Caption         =   "frmTradeReportFilter.frx":0B96
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":0BC4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":0BE4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   27
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
         Caption         =   "frmTradeReportFilter.frx":0C00
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeReportFilter.frx":0C26
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeReportFilter.frx":0C46
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmTradeReportFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTradeReportFilter.frm
'' Description: Allow the user to filter their trades on certain criteria
''              before calling the performance report on the filtered trades
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/21/2009   DAJ         Now storing the Max Profit/Max Loss to the database
'' 01/30/2009   DAJ         When calculating ending date for loading Bars for
''                          Max Profit/Max Loss, if the position is open, need to
''                          load up to the end of data
'' 05/04/2009   DAJ         Fixed bug that was causing full data to be loaded
'' 09/30/2011   DAJ         Added code for capturing a report image instead of showing it
'' 02/20/2013   DAJ         Added automated trading item filter, utilize settings object
'' 04/03/2013   DAJ         Automated Strategy Baskets
'' 06/19/2013   DAJ         Fix for date filter
'' 01/23/2014   DAJ         If filtered by account, send account name as report name
'' 06/20/2014   DAJ         If user selects continuous contract, show all trades for that base symbol
'' 05/20/2015   DAJ         Allow multiple accounts for the trade report filter
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    bShowTrades As Boolean              ' Show the trades form?
    bShowForm As Boolean                ' Did we show the form?
    bSettingsMode As Boolean            ' Are we in settings mode?
    
    RptBridge As cRptBridge             ' Reports bridge object for showing reports
    TradeRules As cTradeRules           ' Trade rules object
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Settings, Show Trades Button?, Filename for capture, Show form?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal Settings As cTradeFilterSettings = Nothing, Optional ByVal bShowTradesButton As Boolean = True, Optional ByVal strCaptureFile As String = "", Optional ByVal bShowForm As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim strControls As String           ' Current settings from the controls
    Dim lAccountID As Long              ' Account ID
    
    m.bShowForm = bShowForm
    m.bSettingsMode = False
    LoadCategoryCombo
    
    If Settings Is Nothing Then
        Set Settings = New cTradeFilterSettings
        Settings.LoadFromIni
    End If
    
    InitGrid
    LoadGrid Settings
    
    SetControls Settings

    ' DAJ 09/06/2007: Need to clear out the reports bridge (and therefore the reports form
    ' if it is showing) to avoid some modality issues...
    Set m.RptBridge = Nothing
    
    chkAccount.Enabled = True
    fgAccounts.Enabled = True
    chkCalcPNL.Visible = True
    cmdTrades.Visible = bShowTradesButton
    EnableControls
    
    If bShowForm Then
        ShowForm Me, eForm_ActModal, frmMain, , ALT_GRID_ROW_COLOR
    End If
    
    If (m.bOK = True) Or (bShowForm = False) Then
        ShowReports strCaptureFile
    ElseIf m.bShowTrades = True Then
        lAccountID = CLng(Val(fgAccounts.TextMatrix(fgAccounts.Row, 3)))

        If lAccountID > 0 Then
            frmTTPositions.ShowMe lAccountID, , eGDTradeTrackerTab_Trades, GetControls
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTradeReportFilter.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowForTradeItem
'' Description: Show the form for the given trade item
'' Inputs:      Trade Item, Show Form?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowForTradeItem(ByVal TradeItem As cAutoTradeItem, Optional ByVal bShowForm As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim Settings As cTradeFilterSettings ' Trade filter settings
    
    Set Settings = New cTradeFilterSettings
    Settings.UseAutoTrade = True
    Settings.AutoTradeID = TradeItem.AutoTradeItemID
    
    ShowForTradeItem = ShowMe(Settings, , , bShowForm)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeReportFilter.ShowForTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowForSettings
'' Description: Show the form in settings only mode
'' Inputs:      Settings
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowForSettings(Settings As cTradeFilterSettings) As Boolean
On Error GoTo ErrSection:

    m.bSettingsMode = True
    
    InitGrid
    LoadGrid Settings
    
    LoadCategoryCombo
    SetControls Settings

    cmdTrades.Visible = False
    chkCalcPNL.Visible = False
    chkAccount.Enabled = False
    fgAccounts.Enabled = False
    
    ShowForm Me, eForm_ActModal, frmMain, , ALT_GRID_ROW_COLOR
    
    If m.bOK Then
        Set Settings = GetControls
    End If
    
    ShowForSettings = m.bOK

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeReportFilter.ShowForSettings"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAccount_Click
'' Description: Enable/Disable controls when the user changes this control
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAccount_Click()
On Error GoTo ErrSection:

    If Visible Then
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.chkAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the ShowMe to unload the form without bringing up reports
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    m.bShowTrades = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: Allow the user to lookup a symbol with the symbol selector
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookup_Click()
On Error GoTo ErrSection:

    LookupSymbol
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.cmdLookup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdManageRules_Click
'' Description: Allow the user to manage their custom trade rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdManageRules_Click()
On Error GoTo ErrSection:

    Dim lEntryRuleID As Long            ' Entry rule ID in the combo
    Dim lExitRuleID As Long             ' Exit rule ID in the combo

    ' Only need to reload everything if the user changed something...
    If frmTradeRules.ShowMe Then
        ' Save off the currently selected items in the combo boxes...
        If cboEntryRules.ListIndex = -1& Then
            lEntryRuleID = -1&
        Else
            lEntryRuleID = cboEntryRules.ItemData(cboEntryRules.ListIndex)
        End If
        If cboExitRules.ListIndex = -1& Then
            lExitRuleID = -1&
        Else
            lExitRuleID = cboExitRules.ItemData(cboExitRules.ListIndex)
        End If
    
        ' Reload the trade rules collection and combo boxes...
        m.TradeRules.Load
        m.TradeRules.LoadEntryCombo cboEntryRules
        m.TradeRules.LoadExitCombo cboExitRules
        
        ' Set the combo boxes to what they were selected to before if possible...
        If lEntryRuleID = -1& Then
            cboEntryRules.ListIndex = -1&
        Else
            m.TradeRules.SetEntryCombo lEntryRuleID
        End If
        If lExitRuleID = -1& Then
            cboExitRules.ListIndex = -1&
        Else
            m.TradeRules.SetExitCombo lExitRuleID
        End If
        
        ' If the Trade Tracker is loaded, refresh the trade rules over there as well...
        If FormIsLoaded("frmTTPositions") Then
            frmTTPositions.RefreshTradeRules
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.cmdManagerRules_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allow the ShowMe to unload the form and bring up the reports
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If (chkAccount.Value = vbChecked) And (SelectedAccounts = 0) Then
        chkAccount.Value = vbUnchecked
    End If

    m.bOK = HasLevel(eTN3_Standard, True)
    m.bShowTrades = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdTrades_Click
'' Description: Allow the user to go to the trade tracker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTrades_Click()
On Error GoTo ErrSection:
    
    Dim lAccountID As Long              ' Account ID
    
    If mFlexGrid.ValidGridRow(fgAccounts) Then
        lAccountID = CLng(Val(fgAccounts.TextMatrix(fgAccounts.Row, 3)))

        If lAccountID > 0& Then
            m.bOK = False
            m.bShowTrades = True
            Hide
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.cmdTrades_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_BeforeMouseDown
'' Description: Notification that the mouse has been pressed in the grid
'' Inputs:      Button, Shift/Ctrl/Alt status, X Location of the click,
''              Y Location of the click, Cancel the click?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseCol As Long               ' Column in the grid that the mouse is in
    Dim lMouseRow As Long               ' Row in the grid that the mouse is in

    If Button = vbLeftButton Then
        With fgAccounts
            lMouseCol = .MouseCol
            lMouseRow = .MouseRow
                        
            If mFlexGrid.ValidGridRow(fgAccounts, lMouseRow) = True Then
                .Row = lMouseRow
                
                If lMouseCol = 0 Then
                    CheckedCell(fgAccounts, lMouseRow, 0) = Not CheckedCell(fgAccounts, lMouseRow, 0)
                    
                    SetAccountsCheckBox
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.fgAccounts_BeforeMouseDown"
    
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

    Me.Icon = Picture16(ToolbarIcon("ID_TradeFilter"), , True)       '6412
    
    g.Styler.StyleForm Me
    
    If IsWoodiesVersion Then
        Caption = "Woodies Trade Report Filter"
        fraTradeFilter.Caption = "Woodies CCI Club Filters:"
    Else
        Caption = "Trade Report Filter"
        fraTradeFilter.Caption = "Trade Filters:"
    End If
    
    ' Hide the real vs sim flag stuff for now since we are handling it in the
    ' performance reports...
    fraRealSim.Visible = False
    ''fraTradeFilter.Height = 1155 '1875
    ''Height = 4350 '5070
    CenterTheForm Me
    
    Set m.TradeRules = New cTradeRules
    m.TradeRules.Load
    
    m.TradeRules.LoadEntryCombo cboEntryRules
    m.TradeRules.LoadExitCombo cboExitRules

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Let the ShowMe unload the form if the user clicks on the 'X'
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        m.bShowTrades = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Terminate
'' Description: Terminate the module objects when the form is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Terminate()
On Error GoTo ErrSection:

    Set m.RptBridge = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.Form_Terminate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save settings and clean up when the form goes away
'' Inputs:      Cancel Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim Settings As cTradeFilterSettings ' Trade report filter settings
    
    If (m.bShowForm = True) And (m.bSettingsMode = False) Then
        Set Settings = GetControls
        Settings.SaveToIni
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_Click
'' Description: Allow the user to lookup a symbol with the symbol selector
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    LookupSymbol
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.txtSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_KeyPress
'' Description: Allow the user to lookup a symbol with the symbol selector
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LookupSymbol KeyAscii
    KeyAscii = 0
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.txtSymbol_KeyPress", 0
    
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
        Set astrSymbol = frmSymbolSelector.ShowMe(txtSymbol.Text, False, True, "Symbol to show trades for", , , True)
    Else
        Set astrSymbol = frmSymbolSelector.ShowMe(Chr(KeyAscii), False, True, "Symbol to show trades for", False, False, True)
    End If
    
    If astrSymbol.Size > 0 Then
        ' DAJ 06/20/2014: No longer convert continuous contracts -- if the user selects a
        ' continuous, we will show all trades for that base symbol...
        'strSymbol = ConvertToTradeSymbol(astrSymbol(0), Date)
        strSymbol = ConvertToTradeSymbol(astrSymbol(0))
        If strSymbol <> UCase(Trim(txtSymbol.Text)) Then
            txtSymbol.Text = strSymbol
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.LookupSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCategoryCombo
'' Description: Load up the automated trading items combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCategoryCombo()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Auto trade item
    
    With cboCategory
        .Clear
        
        .AddItem "Manual Trades"
        .ItemData(.NewIndex) = 0&
        
        For lIndex = 1 To g.TradingItems.Count
            Set TradeItem = g.TradingItems(lIndex)
            
            If g.Broker.HideAccount(TradeItem.AccountID) = False Then
                .AddItem TradeItem.Name
                .ItemData(.NewIndex) = TradeItem.AutoTradeItemID
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeReportFilter.LoadCategoryCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetControls
'' Description: Set the controls based on the given settings string
'' Inputs:      Settings
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetControls(ByVal Settings As cTradeFilterSettings)
On Error GoTo ErrSection:

    If Not Settings Is Nothing Then
        CheckBoxValue(chkDateRange) = Settings.UseDateRange
        gdFromDate.Value = Settings.FromDate
        gdToDate.Value = Settings.ToDate
        CheckBoxValue(chkAccount) = Settings.UseAccount
        CheckBoxValue(chkSymbol) = Settings.UseSymbol
        txtSymbol.Text = Settings.Symbol
        Select Case Settings.Direction
            Case eGDFilterDirection_All
                optBoth.Value = True
            Case eGDFilterDirection_Longs
                optLongs.Value = True
            Case eGDFilterDirection_Shorts
                optShorts.Value = True
        End Select
        CheckBoxValue(chkEntryRule) = Settings.UseEntryRule
        m.TradeRules.SetEntryCombo Settings.EntryRuleID
        CheckBoxValue(chkExitRule) = Settings.UseExitRule
        m.TradeRules.SetExitCombo Settings.ExitRuleID
        Select Case Settings.TradeType
            Case eGDFilterTradeType_All
                optAllFlags.Value = True
            Case eGDFilterTradeType_Real
                optReal.Value = True
            Case eGDFilterTradeType_Sim
                optSim.Value = True
        End Select
        CheckBoxValue(chkCalcPNL) = Settings.CalcPnl
        CheckBoxValue(chkCategory) = Settings.UseAutoTrade
        If SelectComboByItemData(cboCategory, Settings.AutoTradeID) = False Then
            cboCategory.ListIndex = 0&
        End If
    Else
        chkDateRange.Value = vbUnchecked
        gdFromDate.Value = Date
        gdToDate.Value = Date
        chkAccount.Value = vbUnchecked
        chkSymbol.Value = vbUnchecked
        txtSymbol.Text = ""
        optBoth.Value = True
        chkEntryRule.Value = vbUnchecked
        m.TradeRules.SetEntryCombo -1
        chkExitRule.Value = vbUnchecked
        m.TradeRules.SetExitCombo -1
        optAllFlags.Value = True
        chkCalcPNL.Value = vbUnchecked
        chkCategory.Value = vbUnchecked
        cboCategory.ListIndex = 0&
    End If
    
    SetAccountsCheckBox

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.SetControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetControls
'' Description: Get the current control settings
'' Inputs:      None
'' Returns:     Settings
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetControls() As cTradeFilterSettings
On Error GoTo ErrSection:

    Dim Settings As cTradeFilterSettings ' Trade report filter settings
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Settings = New cTradeFilterSettings
    Settings.UseDateRange = CheckBoxValue(chkDateRange)
    Settings.FromDate = gdFromDate.Value
    Settings.ToDate = gdToDate.Value
    Settings.UseAccount = CheckBoxValue(chkAccount)
    
    Settings.AccountIds.Clear
    With fgAccounts
        For lIndex = .FixedRows To .Rows - 1
            If CheckedCell(fgAccounts, lIndex, 0) = True Then
                Settings.AccountIds.Add CLng(Val(.TextMatrix(lIndex, 3)))
            End If
        Next lIndex
    End With
    
    Settings.UseSymbol = CheckBoxValue(chkSymbol)
    Settings.Symbol = txtSymbol.Text
    Select Case True
        Case optBoth.Value
            Settings.Direction = eGDFilterDirection_All
        Case optLongs.Value
            Settings.Direction = eGDFilterDirection_Longs
        Case optShorts.Value
            Settings.Direction = eGDFilterDirection_Shorts
    End Select
    Settings.UseEntryRule = CheckBoxValue(chkEntryRule)
    Settings.EntryRuleID = cboEntryRules.ItemData(cboEntryRules.ListIndex)
    Settings.UseExitRule = CheckBoxValue(chkExitRule)
    Settings.ExitRuleID = cboExitRules.ItemData(cboExitRules.ListIndex)
    Select Case True
        Case optAllFlags.Value
            Settings.TradeType = eGDFilterTradeType_All
        Case optReal.Value
            Settings.TradeType = eGDFilterTradeType_Real
        Case optSim.Value
            Settings.TradeType = eGDFilterTradeType_Sim
    End Select
    Settings.CalcPnl = CheckBoxValue(chkCalcPNL)
    Settings.UseAutoTrade = CheckBoxValue(chkCategory)
    Settings.AutoTradeID = cboCategory.ItemData(cboCategory.ListIndex)
    
    Set GetControls = Settings

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeReportFilter.GetControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowReports
'' Description: Dump the trades files and call the reports
'' Inputs:      Filename for Capture
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowReports(Optional ByVal strCaptureFile As String = "")
On Error GoTo ErrSection:

    Dim Positions As cAccountPositions  ' Collection of account positions
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim astrFiles As cGdArray           ' List of files that have been created
    Dim astrTrades As cGdArray          ' Trades for the current account position
    Dim bContinue As Boolean            ' Do we want to continue?
    Dim TradeLines As cTradeLines       ' Trade lines object
    Dim bKeep As Boolean                ' Keep the trade line?
    Dim strHeader As String             ' Header line for the trades file
    Dim strFileName As String           ' Filename for the trades file
    Dim dFromDate As Double             ' From date
    Dim dToDate As Double               ' To date
    Dim RptBridge As cRptBridge         ' Reports bridge
    Dim dLastKnownPrice As Double       ' Last known price for a symbol
    Dim dLastKnownTime As Double        ' Date/time for the last known price for a symbol
    Dim tblRules As cGdTable            ' Table of "Rules" and "Rule ID's"
    Dim rs As Recordset                 ' Recordset into the database
    Dim BarsColl As New cGdTree         ' Collection of bars
    Dim Bars As cGdBars                 ' Bars object
    Dim dStartDataDate As Double        ' Starting data date
    Dim dEndDataDate As Double          ' End data date
    Dim lStartBar As Long               ' Starting bar
    Dim astrDataDates As cGdArray       ' Array of data date information
    Dim lPos As Long                    ' Position in an array
    Dim strSymbol As String             ' Symbol to load
    Dim bCalcPnl As Boolean
    Dim lMaxDays As Long
    Dim strMsg As String
    Dim Settings As cTradeFilterSettings ' Trade Report Filter Settings
    Dim strReportName As String         ' Title for the reports
       
    InfBox "Generating report.  Please wait...", , , "Trade Report", True
    
    Set astrFiles = New cGdArray
    astrFiles.Create eGDARRAY_Strings
    Set Positions = New cAccountPositions
    Set tblRules = New cGdTable
        
    tblRules.CreateField eGDARRAY_Longs, 0, "RuleID"
    tblRules.CreateField eGDARRAY_Shorts, 1, "OrderTypeID"
    tblRules.CreateField eGDARRAY_Strings, 2, "RuleName"
    
    Positions.Load
    
    Set Settings = GetControls
    
    bCalcPnl = Settings.CalcPnl
    If bCalcPnl Then
        ' Get the starting and ending dates to load for each unique symbol...
        lMaxDays = 0
        Set astrDataDates = New cGdArray
        For lIndex = 1 To Positions.Count
            bContinue = True
            If Settings.UseAccount Then
                bContinue = Settings.IncludeAccount(Positions(lIndex).AccountID)
            End If
            If (bContinue = True) And (Settings.UseSymbol) Then
                ' DAJ 06/20/2014: If the user selects a continuous contract, show all trades
                ' for the base symbol, else show all trades for the exact symbol...
                If InStr(Settings.Symbol, "-0") = 0 Then
                    bContinue = (Positions(lIndex).Symbol = Settings.Symbol)
                Else
                    bContinue = (Parse(Positions(lIndex).Symbol, "-", 1) = Parse(Settings.Symbol, "-", 1))
                End If
            End If
            If (bContinue = True) And (Settings.UseAutoTrade) Then
                bContinue = (Positions(lIndex).AutoTradeItemID = Settings.AutoTradeID)
                If bContinue = False Then
                    bContinue = (g.TradingItems.ParentID(Positions(lIndex).AutoTradeItemID) = Settings.AutoTradeID)
                End If
            End If
            
            dStartDataDate = Positions(lIndex).StartFillTime
            If (bContinue = True) And (dStartDataDate > 0) Then
                ' Make sure that if there is an open position to load data up to the
                ' end of the data, otherwise load up to the last trade date...
                If (Positions(lIndex).CurrentPositionSnapshot <> 0) Or (Positions(lIndex).CurrentPosition <> 0) Then
                    dEndDataDate = Date + 1
                Else
                    dEndDataDate = MaxDouble(Positions(lIndex).LastTradedSnapshot, Positions(lIndex).LastTraded)
                End If
                
                ' check against date range
                If Settings.UseDateRange Then
                    If (dStartDataDate > Settings.ToDate) Or (dEndDataDate < Settings.FromDate) Then
                        bContinue = False
                    Else
                        If dStartDataDate < Settings.FromDate Then
                            dStartDataDate = Settings.FromDate
                        End If
                        If dEndDataDate > Settings.ToDate Then
                            dEndDataDate = Settings.ToDate
                        End If
                    End If
                End If
                
                If astrDataDates.BinarySearch(Positions(lIndex).Symbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                    dStartDataDate = MinDouble(Val(Parse(astrDataDates(lPos), vbTab, 2)), dStartDataDate)
                    dEndDataDate = MaxDouble(Val(Parse(astrDataDates(lPos), vbTab, 3)), dEndDataDate)
                    astrDataDates(lPos) = Positions(lIndex).Symbol & vbTab & Str(dStartDataDate) & vbTab & Str(dEndDataDate)
                Else
                    astrDataDates.Add Positions(lIndex).Symbol & vbTab & Str(dStartDataDate) & vbTab & Str(dEndDataDate)
                End If
                
                If dEndDataDate - dStartDataDate > lMaxDays Then
                    lMaxDays = Int(dEndDataDate - dStartDataDate) + 1
                End If
            End If
        Next lIndex
    
        ' make sure to warn them if this could take a REALLY long time!
        ' (e.g. if more than 1 symbol and a date range of over a month or so)
        If astrDataDates.Size > 1 And lMaxDays > 45 Then
            strMsg = "Calulating the max profit/loss within each trade may take a REALLY LONG time to load months of intraday data for multiple symbols.  Do you really need to calculate the max profit/loss within each trade?"
            If InfBox(strMsg, "!", "Yes|+-No", "WARNING") = "N" Then
                bCalcPnl = False
            End If
            
            InfBox "Generating report.  Please wait...", , , "Trade Report", True
        End If
    End If
    
    If bCalcPnl Then
        ' Load the bars for each unique symbols...
        For lIndex = 0 To astrDataDates.Size - 1
            strSymbol = Parse(astrDataDates(lIndex), vbTab, 1)
            dStartDataDate = Val(Parse(astrDataDates(lIndex), vbTab, 2))
            dEndDataDate = Val(Parse(astrDataDates(lIndex), vbTab, 3))
        
            Set Bars = New cGdBars
            If DM_GetBars(Bars, strSymbol, "1 minute", dStartDataDate, dEndDataDate) Then
                BarsColl.Add Bars, strSymbol
            End If
        Next lIndex
    End If
    
    For lIndex = 1 To Positions.Count
        If (g.Broker.HideAccount(Positions(lIndex).AccountID) = False) And (Positions(lIndex).AutoTradeItemID >= 0) Then
            bContinue = True
            If Settings.UseAccount Then
                bContinue = Settings.IncludeAccount(Positions(lIndex).AccountID)
            End If
            If (bContinue = True) And (Settings.UseSymbol) Then
                ' DAJ 06/20/2014: If the user selects a continuous contract, show all trades
                ' for the base symbol, else show all trades for the exact symbol...
                If InStr(Settings.Symbol, "-0") = 0 Then
                    bContinue = (Positions(lIndex).Symbol = Settings.Symbol)
                Else
                    bContinue = (Parse(Positions(lIndex).Symbol, "-", 1) = Parse(Settings.Symbol, "-", 1))
                End If
            End If
            If (bContinue = True) And (Settings.UseAutoTrade) Then
                bContinue = (Positions(lIndex).AutoTradeItemID = Settings.AutoTradeID)
                If bContinue = False Then
                    bContinue = (g.TradingItems.ParentID(Positions(lIndex).AutoTradeItemID) = Settings.AutoTradeID)
                End If
            End If
            
            If bContinue = True Then
                StatusMsg "Creating Trade Report for " & Positions(lIndex).Symbol & " in account " & g.Broker.AccountNameForID(Positions(lIndex).AccountID)
                
                If Not bCalcPnl Then
                    Set Bars = Nothing
                ElseIf BarsColl.Exists(Positions(lIndex).Symbol) Then
                    Set Bars = BarsColl(Positions(lIndex).Symbol)
                Else
                    Set Bars = New cGdBars
                    SetBarProperties Bars, Positions(lIndex).Symbol
                End If
                
                dFromDate = 0#
                dToDate = 0#
                
                Set TradeLines = Positions(lIndex).TradeLines.MakeCopy
                For lIndex2 = TradeLines.Count To 1 Step -1
                    bKeep = True
                    If Settings.UseDateRange Then
                        bKeep = ((Int(TradeLines(lIndex2).ExitTime) >= Settings.FromDate) And (Int(TradeLines(lIndex2).ExitTime) <= Settings.ToDate))
                    End If
                    If (bKeep = True) And (Settings.Direction = eGDFilterDirection_Longs) Then
                        bKeep = (TradeLines(lIndex2).Direction = "L")
                    End If
                    If (bKeep = True) And (Settings.Direction = eGDFilterDirection_Shorts) Then
                        bKeep = (TradeLines(lIndex2).Direction = "S")
                    End If
                    If (bKeep = True) And (Settings.UseEntryRule) Then
                        bKeep = (TradeLines(lIndex2).EntryRuleID = Settings.EntryRuleID)
                    End If
                    If (bKeep = True) And (Settings.UseExitRule) Then
                        bKeep = (TradeLines(lIndex2).ExitRuleID = Settings.ExitRuleID)
                    End If
                    'If (bKeep = True) And (Settings.TradeType = eGDFilterTradeType_Real) Then
                    'End If
                    'If (bKeep = True) And (Settings.TradeType = eGDFilterTradeType_Sim) Then
                    'End If
                    
                    If bKeep = False Then
                        TradeLines.Remove lIndex2
                    Else
                        If bCalcPnl Then
                            TradeLines(lIndex2).CalcMaxPNL Bars
                            TradeLines(lIndex2).Save
                        Else
                            TradeLines(lIndex2).ClearMaxPNL
                        End If
                        
                        If (dFromDate = 0#) Or (TradeLines(lIndex2).EntryTime < dFromDate) Then
                            dFromDate = TradeLines(lIndex2).EntryTime
                        End If
                        If (dFromDate = 0#) Or (TradeLines(lIndex2).ExitTime < dFromDate) Then
                            dFromDate = TradeLines(lIndex2).ExitTime
                        End If
                        If (dToDate = 0#) Or (TradeLines(lIndex2).EntryTime > dToDate) Then
                            dToDate = TradeLines(lIndex2).EntryTime
                        End If
                        If (dToDate = 0#) Or (TradeLines(lIndex2).ExitTime > dToDate) Then
                            dToDate = TradeLines(lIndex2).ExitTime
                        End If
                        
                        If TradeLines(lIndex2).IsOpen Then
                            dLastKnownPrice = g.RealTime.LastKnownPrice(Positions(lIndex).SymbolOrSymbolID, 0, True, dLastKnownTime)
                            If dLastKnownPrice <> kNullData Then
                                TradeLines(lIndex2).OpenProfit dLastKnownPrice, dLastKnownTime
                            End If
                        End If
                    End If
                Next lIndex2
                
                If TradeLines.Count > 0 Then
                    strHeader = BuildTradesHeader(Positions(lIndex).AccountPositionID, "Manual Trades for " & Positions(lIndex).Symbol & " in " & g.Broker.AccountNameForID(Positions(lIndex).AccountID), "1 Minute", dFromDate, dToDate, 0#, Positions(lIndex).Symbol)
                    For lIndex2 = 1 To TradeLines.Count
                        TradeLines(lIndex2).TradeNumber = lIndex2
                    Next lIndex2
                    Set astrTrades = TradeLines.ToArray
                    astrTrades.Add strHeader, 0&
                    
                    strFileName = AddSlash(App.Path) & "Trades\S_" & Str(Positions(lIndex).AccountID) & "_" & Str(Positions(lIndex).SymbolOrSymbolID) & "_" & Str(Positions(lIndex).AutoTradeItemID) & ".TXT"
                    astrFiles.Add strFileName
                    astrTrades.ToFile strFileName
                End If
            End If
        End If
    Next lIndex
    
    InfBox ""
        
    If astrFiles.Size > 0 Then
        If Not m.RptBridge Is Nothing Then
            Set m.RptBridge = Nothing
        End If
        Set m.RptBridge = New cRptBridge
        
        If (Settings.UseAccount = True) And (Settings.AccountIds.Size = 1) Then
            strReportName = g.Broker.AccountNameForID(Settings.AccountIds(0))
        Else
            strReportName = "My Trades"
        End If
        
        ShowMergedReports m.RptBridge, strReportName, False, astrFiles.ArrayHandle, m.TradeRules.RulesTable.TableHandle, True, strCaptureFile
    Else
        InfBox "No reports to show", "!", , "Error"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.ShowReports"

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

    Enable cmdTrades, (mFlexGrid.ValidGridRow(fgAccounts))
    Enable chkCategory, (cboCategory.ListCount > 1)
    Enable cboCategory, (cboCategory.ListCount > 1)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    With fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        mGenesis.SetupGrid fgAccounts, eGridMode_List
        
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .FixedCols = 0
        .Cols = 4
        .FixedRows = 0
        .Rows = 0
        
        .ColDataType(0) = flexDTBoolean
        .ColHidden(3) = True
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeReportFilter.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Initialize the grid
'' Inputs:      Settings
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(ByVal Settings As cTradeFilterSettings)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)

    With fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        Do While Not rs.EOF
            If g.Broker.HideAccount(rs!AccountID) = False Then
                .Rows = .Rows + 1
                
                If Settings Is Nothing Then
                    CheckedCell(fgAccounts, .Rows - 1, 0) = False
                Else
                    CheckedCell(fgAccounts, .Rows - 1, 0) = Settings.IncludeAccount(rs!AccountID)
                End If
                .TextMatrix(.Rows - 1, 1) = rs!Name
                .TextMatrix(.Rows - 1, 2) = g.Broker.BrokerName(rs!AccountType)
                .TextMatrix(.Rows - 1, 3) = Str(rs!AccountID)
            End If
            
            rs.MoveNext
        Loop
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            
            .Col = 1
            .Sort = flexSortGenericAscending
            
            mFlexGrid.SetBackColors fgAccounts
        End If
        
        .AutoSize 0, .Cols - 1, False
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAccountsCheckBox
'' Description: Set the accounts check box based on accounts that are checked
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAccountsCheckBox()
On Error GoTo ErrSection:

    CheckBoxValue(chkAccount) = (SelectedAccounts > 0)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeReportFilter.SetAccountsCheckBox"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedAccounts
'' Description: Determine the number of accounts that are turned on
'' Inputs:      None
'' Returns:     Number of accounts that are turned on
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedAccounts() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = 0&
    With fgAccounts
        For lIndex = .FixedRows To .Rows - 1
            If CheckedCell(fgAccounts, lIndex, 0) = True Then
                lReturn = lReturn + 1&
            End If
        Next lIndex
    End With

    SelectedAccounts = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeReportFilter.SelectedAccounts"
    
End Function

