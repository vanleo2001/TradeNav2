VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{C0F09D2D-1125-11D7-8DA9-0004757A4B66}#1.9#0"; "NavTradeSenseOCXV3.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTradeSenseOrder 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraAdvanced 
      Height          =   1515
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   8715
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
      Caption         =   "frmTradeSenseOrder.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeSenseOrder.frx":0030
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrder.frx":0050
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtRequiredMod 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   180
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTradeSenseOrder.frx":006C
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
         Tip             =   "frmTradeSenseOrder.frx":008C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":00AC
      End
      Begin HexUniControls.ctlUniCheckXP chkAllowInputs 
         Height          =   220
         Left            =   120
         TabIndex        =   9
         Top             =   540
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
         Caption         =   "frmTradeSenseOrder.frx":00C8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":0102
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0122
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fgInputs 
         Height          =   855
         Left            =   1440
         TabIndex        =   11
         Top             =   540
         Width           =   6975
         _cx             =   12303
         _cy             =   1508
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
      Begin HexUniControls.ctlUniLabelXP lblRequiredMod 
         Height          =   195
         Left            =   120
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
         Caption         =   "frmTradeSenseOrder.frx":013E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":017E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":019E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   1920
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "frmTradeSenseOrder.frx":01BA
      ToolBars        =   "frmTradeSenseOrder.frx":02ED
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   2460
      TabIndex        =   14
      Top             =   7560
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
      Caption         =   "frmTradeSenseOrder.frx":043E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeSenseOrder.frx":046A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrder.frx":048A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdVerify 
         Height          =   495
         Left            =   1260
         TabIndex        =   17
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
         Caption         =   "frmTradeSenseOrder.frx":04A6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":04D4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":04F4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   2520
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
         Caption         =   "frmTradeSenseOrder.frx":0510
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":053E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":055E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   22
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
         Caption         =   "frmTradeSenseOrder.frx":057A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":05A0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":05C0
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraInformation 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   4800
      Width           =   8175
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
      Caption         =   "frmTradeSenseOrder.frx":05DC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeSenseOrder.frx":0612
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrder.frx":0632
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraNumDays 
         Height          =   615
         Left            =   2760
         TabIndex        =   24
         Top             =   240
         Width           =   5235
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
         Caption         =   "frmTradeSenseOrder.frx":064E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTradeSenseOrder.frx":067A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":069A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optAutoDetect 
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   300
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
            Caption         =   "frmTradeSenseOrder.frx":06B6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmTradeSenseOrder.frx":070A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrder.frx":072A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optOverride 
            Height          =   255
            Left            =   2400
            TabIndex        =   27
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
            Caption         =   "frmTradeSenseOrder.frx":0746
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTradeSenseOrder.frx":0786
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrder.frx":07A6
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtOverride 
            Height          =   315
            Left            =   4140
            TabIndex        =   25
            Top             =   180
            Width           =   1095
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTradeSenseOrder.frx":07C2
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
            Tip             =   "frmTradeSenseOrder.frx":07E4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrder.frx":0804
         End
         Begin HexUniControls.ctlUniTextBoxXP txtNumBars 
            Height          =   315
            Left            =   4080
            TabIndex        =   28
            Top             =   60
            Width           =   1095
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTradeSenseOrder.frx":0820
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
            Tip             =   "frmTradeSenseOrder.frx":0842
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrder.frx":0862
         End
         Begin HexUniControls.ctlUniLabelXP lblNumBars 
            Height          =   195
            Left            =   0
            Top             =   60
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
            Caption         =   "frmTradeSenseOrder.frx":087E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTradeSenseOrder.frx":0912
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrder.frx":0932
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniComboBoxXP cboBarPeriod 
         Height          =   315
         Left            =   1020
         TabIndex        =   23
         Top             =   300
         Width           =   1455
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
         Tip             =   "frmTradeSenseOrder.frx":094E
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
         MouseIcon       =   "frmTradeSenseOrder.frx":096E
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBarPeriod 
         Height          =   195
         Left            =   120
         Top             =   360
         Width           =   915
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
         Caption         =   "frmTradeSenseOrder.frx":098A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":09C2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":09E2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraOrder 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   8835
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
      Caption         =   "frmTradeSenseOrder.frx":09FE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeSenseOrder.frx":0A28
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrder.frx":0A48
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkWrongSide 
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   1920
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
         Caption         =   "frmTradeSenseOrder.frx":0A64
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":0AEE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0B0E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkExpireDay 
         Height          =   375
         Left            =   180
         TabIndex        =   15
         Top             =   1500
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
         Caption         =   "frmTradeSenseOrder.frx":0B2A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":0BC0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0BE0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtExitPct 
         Height          =   315
         Left            =   540
         TabIndex        =   12
         Top             =   1140
         Width           =   480
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTradeSenseOrder.frx":0BFC
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
         Tip             =   "frmTradeSenseOrder.frx":0C22
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0C42
      End
      Begin HexUniControls.ctlUniComboImageXP cboOrderType 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   1455
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
         Tip             =   "frmTradeSenseOrder.frx":0C5E
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0C7E
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboAction 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1455
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
         Tip             =   "frmTradeSenseOrder.frx":0C9A
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0CBA
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin NavTradeSenseV3.Editor tsPrice 
         Height          =   855
         Left            =   2760
         TabIndex        =   18
         Top             =   390
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1508
      End
      Begin NavTradeSenseV3.Editor tsWithLimit 
         Height          =   750
         Left            =   2760
         TabIndex        =   20
         Top             =   1560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1323
      End
      Begin gdOCX.gdScrollBar sbExitPct 
         Height          =   360
         Left            =   1020
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1110
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
      Begin HexUniControls.ctlUniLabelXP lblExitPct 
         Height          =   195
         Left            =   1320
         Top             =   1200
         Width           =   915
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
         Caption         =   "frmTradeSenseOrder.frx":0CD6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":0D10
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0D30
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblExit 
         Height          =   195
         Left            =   180
         Top             =   1200
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
         Caption         =   "frmTradeSenseOrder.frx":0D4C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":0D76
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0D96
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblWithLimit 
         Height          =   255
         Left            =   2760
         Top             =   1320
         Width           =   5385
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
         Caption         =   "frmTradeSenseOrder.frx":0DB2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":0DF6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0E16
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPrice 
         Height          =   240
         Left            =   2760
         Top             =   180
         Width           =   5385
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
         Caption         =   "frmTradeSenseOrder.frx":0E32
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":0E6C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0E8C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOrderType 
         Height          =   195
         Left            =   240
         Top             =   660
         Width           =   915
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
         Caption         =   "frmTradeSenseOrder.frx":0EA8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":0EE0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0F00
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAction 
         Height          =   195
         Left            =   240
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
         Caption         =   "frmTradeSenseOrder.frx":0F1C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":0F4C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":0F6C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraCondition 
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8835
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
      Caption         =   "frmTradeSenseOrder.frx":0F88
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeSenseOrder.frx":0FBA
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrder.frx":0FDA
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboEvaluate 
         Height          =   315
         Left            =   5520
         TabIndex        =   4
         Top             =   420
         Width           =   1455
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
         Tip             =   "frmTradeSenseOrder.frx":0FF6
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":1016
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCondition 
         Height          =   220
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "frmTradeSenseOrder.frx":1032
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":10AC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":10CC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optNoCondition 
         Height          =   220
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
         Caption         =   "frmTradeSenseOrder.frx":10E8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":1122
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":1142
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin NavTradeSenseV3.Editor tsCondition 
         Height          =   1095
         Left            =   420
         TabIndex        =   5
         Top             =   780
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   1931
      End
      Begin HexUniControls.ctlUniLabelXP lblEvaluate 
         Height          =   195
         Left            =   4080
         Top             =   480
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
         Caption         =   "frmTradeSenseOrder.frx":115E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTradeSenseOrder.frx":11A6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrder.frx":11C6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmTradeSenseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTradeSenseOrder.frm
'' Description: Form that handles a Trade Sense order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 06/03/2010   DAJ         Changes for new Trade Sense Order Groups
'' 06/14/2010   DAJ         Changed the caption for the condition per Pete (#5780)
'' 06/17/2010   DAJ         Changed filenames to ID instead of name
'' 07/15/2010   DAJ         Added capabilities for inputs
'' 07/26/2010   DAJ         Fixed Exit Pct control when new order
'' 08/19/2010   DAJ         Added capability for TradeSense order expire at end of session
'' 08/23/2010   DAJ         Added required module flag for TradeSense orders/groups
'' 09/16/2010   DAJ         Allow the Enter Key in the condition editor for macros
'' 09/16/2010   DAJ         Set Dirty flag for Check Intra Bar check box
'' 09/29/2010   DAJ         Swapped Action/Order Type, moved Evaluate to top
'' 10/01/2010   DAJ         Fixed dirty flag on evaluate combo
'' 10/08/2010   DAJ         Don't allow Next Bar functions in expressions
'' 10/11/2010   DAJ         Use "Close" for Market order instead of "Next Bar Open"
'' 12/09/2010   DAJ         More descriptive labels for price editors
'' 06/08/2011   DAJ         Don't allow Intra-Bar evaluation if no condition
'' 06/09/2011   DAJ         Implemented auto-breakout bar period
'' 10/17/2011   DAJ         Added the auto breakout for TradeSense order groups function
'' 12/09/2011   DAJ         Allow Next Bar Open if evaluating on each new bar
'' 05/01/2012   DAJ         Allow criteria and charting functions in the editors
'' 11/28/2012   DAJ         Submit Market if Stop on wrong side of market flag
'' 01/16/2014   DAJ         Don't allow user to type in FractZen as bar type if not authorized ( #6948 )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kErrorCaption = "Trade Sense Order Error"

' Coded text tokens
Private Const kIfToken = 24
Private Const kThenToken = 35
Private Const kEnterToken = 80
Private Const kLeftParenToken = 16
Private Const kRightParenToken = 17
Private Const kCommaToken = 22

Private Enum eGDSaveCmd
    eGDSaveCmd_Save = 0
    eGDSaveCmd_SaveAs
    eGDSaveCmd_Rename
End Enum

Private Enum eGDEvaluate
    eGDEvaluate_EachBar = 0
    eGDEvaluate_IntraBar
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    bForAutoExit As Boolean             ' Is this use for an auto exit order?
    
    ExitPct As cPriceEditor             ' Price editor to handle the exit percent
    
    ListLoading As cListLoading         ' Lists of stuff for TradeSense
    lateCalc As cLateCalculating        ' Late calculating object
    
    strCodedCondition As String         ' Condition in coded text
    strCodedPrice As String             ' Price in coded text
    strCodedWithLimit As String         ' With Limit price in coded text
    strFormattedCondition As String     ' Condition in formatted text
    strFormattedPrice As String         ' Price in formatted text
    strFormattedWithLimit As String     ' With Limit price in formatted text
    
    tsOrder As cTradeSenseOrder         ' Trade Sense order object
    bDirty As Boolean                   ' Has the user made changes?
    bVerified As Boolean                ' Has the expression been verified?
    
    Inputs As cTradeSenseOrderInputs    ' Collection of inputs
End Type
Private m As mPrivate

Private Property Get Dirty() As Boolean
    Dirty = m.bDirty
End Property
Private Property Let Dirty(ByVal bDirty As Boolean)
    m.bDirty = bDirty
    EnableControls
End Property

Private Property Get Verified() As Boolean
    Verified = m.bVerified
End Property
Private Property Let Verified(ByVal bVerified As Boolean)
    m.bVerified = bVerified
    EnableControls
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup the controls and show the form
'' Inputs:      Trade Sense Order, Order Action, Order Type, Allow Opposite Type?,
''              For Auto Exit?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(TradeSenseOrder As cTradeSenseOrder, Optional ByVal nOrderAction As eGDOrderAction = -1&, Optional ByVal nOrderType As eTT_OrderType = -1&, Optional ByVal bAllowOppositeOrderType As Boolean = True, Optional ByVal bForAutoExit As Boolean = True) As Boolean
On Error GoTo ErrSection:

    m.bForAutoExit = bForAutoExit
    Set m.tsOrder = TradeSenseOrder
    
    If m.bForAutoExit Then
        Caption = "Trade Sense Order"
    Else
        SetEditorCaption Me, "Trade Sense Order", m.tsOrder.Name
    End If
    
    InitInputsGrid
    LoadActionCombo nOrderAction
    LoadOrderTypeCombo nOrderType, bAllowOppositeOrderType
    
    ControlsFromObject TradeSenseOrder
    
    m.bDirty = False
    m.bVerified = True
    
    EnableControls
    If (optCondition.Value = True) And (Len(Trim(tsCondition.Text)) = 0) Then
        MoveFocus tsCondition
    End If

    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK Then
        If m.bForAutoExit Then
            ObjectFromControls TradeSenseOrder
        End If
    End If
    
    If m.bForAutoExit = False Then
        Set TradeSenseOrder = m.tsOrder
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTradeSenseOrder.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAction_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAction_Click()
On Error GoTo ErrSection:

    If Visible Then
        SetPriceEditorLabels
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.cboAction_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboBarPeriod_Change
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboBarPeriod_Change()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.cboBarPeriod_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboBarPeriod_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboBarPeriod_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.cboBarPeriod_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboBarPeriod_LostFocus
'' Description: Fix the bar period if necessary after the focus moves
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboBarPeriod_LostFocus()
On Error GoTo ErrSection:

    Dim strPeriod As String             ' Period string

    If Visible Then
        strPeriod = FixPeriod(cboBarPeriod.Text)
        
        If UCase(strPeriod) = "FRACTZEN" And Not g.FractZen.AllowTSOG Then
            strPeriod = "Daily"
            InfBox "You are not authorized to use FractZen bars", "!", , kErrorCaption
            MoveFocus cboBarPeriod
        End If
        
        If strPeriod <> cboBarPeriod.Text Then
            cboBarPeriod.Text = strPeriod
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.cboBarPeriod_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboEvaluate_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboEvaluate_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.cboEvaluate_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboOrderType_Click
'' Description: Handle the case when the user changes the order type value
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboOrderType_Click()
On Error GoTo ErrSection:

    Dim nOrderType As eTT_OrderType     ' Order type value
    
    If cboOrderType.ListIndex <> -1 Then
        nOrderType = cboOrderType.ItemData(cboOrderType.ListIndex)
        
        Select Case nOrderType
            Case eTT_OrderType_Market
                lblPrice.Visible = False
                tsPrice.Visible = False
                lblWithLimit.Visible = False
                tsWithLimit.Visible = False
                
            Case eTT_OrderType_Limit
                lblPrice.Visible = True
                tsPrice.Visible = True
                lblWithLimit.Visible = False
                tsWithLimit.Visible = False
                
                tsPrice.Height = fraOrder.Height - tsPrice.Top - 120
            
            Case eTT_OrderType_Stop
                lblPrice.Visible = True
                tsPrice.Visible = True
                lblWithLimit.Visible = False
                tsWithLimit.Visible = False
                
                tsPrice.Height = fraOrder.Height - tsPrice.Top - 120
            
            Case eTT_OrderType_StopWithLimit
                lblPrice.Visible = True
                tsPrice.Visible = True
                lblWithLimit.Visible = True
                tsWithLimit.Visible = True
                
                tsPrice.Height = fraOrder.Height - lblWithLimit.Top - tsPrice.Top
            
        End Select
    End If
    
    If Visible Then
        SetPriceEditorLabels
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.cboOrderType_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAllowInputs_Click
'' Description: Set the order as dirty when user changes this setting
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAllowInputs_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.chkAllowInputs_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkExpireDay_Click
'' Description: Set the order as dirty when user changes this setting
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkExpireDay_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.chkExpireDay_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkWrongSide_Click
'' Description: Set the order as dirty when user changes this setting
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkWrongSide_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.chkWrongSide_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Handle the case when the user clicks on the Cancel button
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
    RaiseError "frmTradeSenseOrder.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Handle the case when the user clicks on the OK button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If VerifyExpression Then
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdVerify_Click
'' Description: Handle the case when the user clicks on the Verify button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdVerify_Click()
On Error GoTo ErrSection:

    VerifyExpression

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.cmdVerify_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgInputs_AfterEdit
'' Description: Update the inputs array after the user changes a value
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    With fgInputs
        If m.Inputs.Exists(.TextMatrix(Row, 0)) Then
            m.Inputs(.TextMatrix(Row, 0)).DefaultValue = .TextMatrix(Row, 1)
            Dirty = True
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.fgInputs_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgInputs_BeforeEdit
'' Description: Only allow the user to edit the default value column
'' Inputs:      Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim tsInput As cTradeSenseOrderInput ' Order input object

    If Col <> 1 Then
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
    RaiseError "frmTradeSenseOrder.fgInputs_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Set the focus to the condition when the form is activated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    ' Load internally generated TradeSense lists (Symbols, etc.)...
    Set m.ListLoading = New cListLoading
    m.ListLoading.Load
    
    ''MoveFocus tsCondition

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.Form_Activate"
    
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

    Dim strPlacement As String          ' Placement of the form
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("frmTradeSenseOrder", "", "Placement", g.strIniFile)
    If Len(strPlacement) > 0 Then
        SetFormPlacement Me, strPlacement, "LTHW"
    Else
        CenterTheForm Me
    End If
    
    Me.Icon = Picture16(ToolbarIcon("ID_Rules"), , True)
    
    ' Load internally generated TradeSense lists (Symbols, etc.)...
    Set m.ListLoading = New cListLoading
    m.ListLoading.Load
    
    Set m.lateCalc = New cLateCalculating
    
    InitializeEditor tsCondition, False
    InitializeEditor tsPrice, True
    InitializeEditor tsWithLimit, True

    With cboBarPeriod
        .AddItem "Daily"
        .ListIndex = .NewIndex
        
        .AddItem "60 Minute"
        .AddItem "30 Minute"
        .AddItem "10 Minute"
        .AddItem "5 Minute"
        
        If g.FractZen.AllowTSOG Then
            .AddItem "FractZen" '"Auto Breakout"
        End If
    End With
    
    With cboEvaluate
        .AddItem "Each New Bar"
        .ListIndex = .NewIndex
        .ItemData(.NewIndex) = eGDEvaluate_EachBar
        
        .AddItem "Intra-Bar"
        .ItemData(.NewIndex) = eGDEvaluate_IntraBar
    End With

    With tbToolbar
        .Tools("ID_Verify").Picture = Picture16(ToolbarIcon("kVerify"))
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_Exit").Picture = Picture16(ToolbarIcon("kCancel"))
    End With
    
    Set m.Inputs = New cTradeSenseOrderInputs
    m.Inputs.ForGroups = False

    Set m.ExitPct = New cPriceEditor
    txtExitPct.Text = "100"
    m.ExitPct.Init sbExitPct, txtExitPct, Nothing, 100, 1, 100
    
    txtOverride.Move txtNumBars.Left, txtNumBars.Top, txtNumBars.Width, txtNumBars.Height

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Handle the case when the user closes form with the context menu
'' Inputs:      Cancel Unload?, Unload Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        
        If m.bForAutoExit Then
            m.bOK = False
            Hide
        Else
            ExitForm
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinScaleWidth As Long          ' Minimum allowed scale width
    Dim lMinScaleHeight As Long         ' Minimum allowed scale height
    Dim lVertSpace As Long              ' Vertical space between controls
    Dim lHorzSpace As Long              ' Horizontal space between controls
    
    lVertSpace = 120
    lHorzSpace = 120
    
    If Not ShowAdvancedTSOG Or m.bForAutoExit Then
        lMinScaleWidth = 8415
        lMinScaleHeight = 6890 ' 5700
    Else
        lMinScaleWidth = 8415
        lMinScaleHeight = 8190 ' 7000
    End If

    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        If m.bForAutoExit Then
            With fraButtons
                .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height - lVertSpace
            End With
            
            With fraInformation
                .Move lHorzSpace, fraButtons.Top - .Height - lVertSpace, ScaleWidth - (lHorzSpace * 2)
            End With
        ElseIf ShowAdvancedTSOG Then
            With fraAdvanced
                .Move lHorzSpace, ScaleHeight - .Height - lVertSpace, ScaleWidth - (lHorzSpace * 2)
            End With
            With fgInputs
                .Move .Left, .Top, fraAdvanced.Width - .Left - lHorzSpace, fraAdvanced.Height - .Top - lVertSpace
            End With
            
            With fraInformation
                .Move lHorzSpace, fraAdvanced.Top - .Height - lVertSpace, ScaleWidth - (lHorzSpace * 2)
            End With
        Else
            With fraInformation
                .Move lHorzSpace, ScaleHeight - .Height - lVertSpace, ScaleWidth - (lHorzSpace * 2)
            End With
        End If
            
        With fraOrder
            .Move lHorzSpace, fraInformation.Top - .Height - lVertSpace, ScaleWidth - (lHorzSpace * 2)
            
            With tsPrice
                .Move .Left, .Top, fraOrder.Width - .Left - lHorzSpace
            End With
            With tsWithLimit
                .Move .Left, .Top, fraOrder.Width - .Left - lHorzSpace
            End With
        End With
        
        With fraCondition
            .Move lHorzSpace, lVertSpace, ScaleWidth - (lHorzSpace * 2), fraOrder.Top - (lVertSpace * 2)
            
            With tsCondition
                .Move .Left, .Top, fraCondition.Width - .Left - lHorzSpace, fraCondition.Height - .Top - lVertSpace
            End With
        End With
    End If

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

    SetIniFileProperty "frmTradeSenseOrder", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAutoDetect_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAutoDetect_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrder.optAutoDetect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optCondition_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optCondition_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
        If Len(Trim(tsCondition.Text)) = 0 Then
            MoveFocus tsCondition
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.optCondition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optNoCondition_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optNoCondition_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.optNoCondition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optOverride_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optOverride_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrder.optOverride"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    sbExitPct_Change
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub sbExitPct_Change()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.sbExitPct_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle a user action on the toolbar
'' Inputs:      Tool
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Select Case UCase(Tool.ID)
        Case "ID_VERIFY"
            VerifyExpression
        Case "ID_SAVE"
            Save eGDSaveCmd_Save
        Case "ID_SAVEAS"
            Save eGDSaveCmd_SaveAs
        Case "ID_RENAME"
            Save eGDSaveCmd_Rename
        Case "ID_EXIT"
            ExitForm
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tbToolbar_ToolClick"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_Change
'' Description: Handle the user changing the condition expression
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_Change()
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of the IF in the text

    lPos = PositionOfIf
    If lPos <> 0 Then
        ' If the user tries to type in an assignment expression after the IF, put up an error message...
        If InStr(lPos, tsCondition.Text, ":=") <> 0 Then
            Tag = "tsCondition"
            InfBox "You cannot have an assignment operator after the IF.", "!", , kErrorCaption
            tsCondition.Text = Left(tsCondition.Text, lPos - 1) & Replace(tsCondition.Text, ":=", "", lPos)
        End If
    End If
    
    Dirty = True
    Verified = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsCondition_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_EditFunction
'' Description: Allow the user to edit a function from the expression
'' Inputs:      Function ID, Function Name, Found
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
On Error GoTo ErrSection:

    ShowFunctionMgr FunctionID, FunctionName, Found

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsCondition_EditFunction"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_GotFocus
'' Description: Set up the editor when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_GotFocus()
On Error GoTo ErrSection:

    If Visible Then
        Set g.ActiveEditor = tsCondition
        InitializeEditor tsCondition, False
        
        If Len(Trim(tsCondition.Text)) = 0 Then
            tsCondition.Text = ""
            SendKeys " "
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrder.tsCondition_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_LostFocus
'' Description: Clean up when the editor loses focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_LostFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = Nothing
    tsCondition.RemoveTradeSense
    
    If UCase(Trim(tsCondition.Text)) = "IF TRUE" Then
        optNoCondition.Value = True
        tsCondition.Text = ""
    ElseIf (Len(Trim(tsCondition.Text)) > 0) And (optCondition.Value = False) Then
        optCondition.Value = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsConditon_LostFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_NewFunction
'' Description: Allow the user to create a new function from the editor
'' Inputs:      Category
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_NewFunction(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    Dim frmFuncMgr As frmFunctionMgrCT  ' Function manager form
    
    Set frmFuncMgr = New frmFunctionMgrCT
    frmFuncMgr.ShowMe 0&, , , lCategoryID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "tsCondition_NewFunction"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsPrice_Change
'' Description: Handle the user changing the condition expression
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsPrice_Change()
On Error GoTo ErrSection:

    ' If the user tries to type in an assignment expression into the Limit price
    ' expression, put up an error message...
    If InStr(tsPrice.Text, ":=") <> 0 Then
        Tag = "tsPrice"
        InfBox "You cannot have an assignment operator in this expression.", "!", , kErrorCaption
        tsPrice.Text = Replace(tsPrice.Text, ":=", "")
        If Len(tsPrice.Text) > 0 Then
            tsPrice.SelStart = Len(tsPrice.Text)
        End If
    End If
    
    Dirty = True
    Verified = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsPrice_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsPrice_EditFunction
'' Description: Allow the user to edit a function from the expression
'' Inputs:      Function ID, Function Name, Found
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsPrice_EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
On Error GoTo ErrSection:

    ShowFunctionMgr FunctionID, FunctionName, Found

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsPrice_EditFunction"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsPrice_GotFocus
'' Description: Set up the editor when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsPrice_GotFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = tsPrice
    InitializeEditor tsPrice, True
    
    If Len(Trim(tsPrice.Text)) = 0 Then
        tsPrice.Text = ""
        SendKeys " "
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsPrice_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsPrice_LostFocus
'' Description: Clean up when the editor loses focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsPrice_LostFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = Nothing
    tsPrice.RemoveTradeSense
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsPrice_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsPrice_NewFunction
'' Description: Allow the user to create a new function from the editor
'' Inputs:      Category
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsPrice_NewFunction(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    Dim frmFuncMgr As frmFunctionMgrCT  ' Function manager form
    
    Set frmFuncMgr = New frmFunctionMgrCT
    frmFuncMgr.ShowMe 0&, , , lCategoryID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "tsPrice_NewFunction"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsWithLimit_Change
'' Description: Handle the user changing the condition expression
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsWithLimit_Change()
On Error GoTo ErrSection:

    ' If the user tries to type in an assignment expression into the Limit price
    ' expression, put up an error message...
    If InStr(tsWithLimit.Text, ":=") <> 0 Then
        Tag = "tsWithLimit"
        InfBox "You cannot have an assignment operator in this expression.", "!", , kErrorCaption
        tsWithLimit.Text = Replace(tsWithLimit.Text, ":=", "")
        If Len(tsWithLimit.Text) > 0 Then
            tsWithLimit.SelStart = Len(tsWithLimit.Text)
        End If
    End If
    
    Dirty = True
    Verified = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsWithLimit_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsWithLimit_EditFunction
'' Description: Allow the user to edit a function from the expression
'' Inputs:      Function ID, Function Name, Found
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsWithLimit_EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
On Error GoTo ErrSection:

    ShowFunctionMgr FunctionID, FunctionName, Found

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsWithLimit_EditFunction"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsWithLimit_GotFocus
'' Description: Set up the editor when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsWithLimit_GotFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = tsWithLimit
    InitializeEditor tsWithLimit, True
    
    If Len(Trim(tsWithLimit.Text)) = 0 Then
        tsWithLimit.Text = ""
        SendKeys " "
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsWithLimit_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsWithLimit_LostFocus
'' Description: Clean up when the editor loses focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsWithLimit_LostFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = Nothing
    tsWithLimit.RemoveTradeSense
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.tsWithLimit_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsWithLimit_NewFunction
'' Description: Allow the user to create a new function from the editor
'' Inputs:      Category
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsWithLimit_NewFunction(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    Dim frmFuncMgr As frmFunctionMgrCT  ' Function manager form
    
    Set frmFuncMgr = New frmFunctionMgrCT
    frmFuncMgr.ShowMe 0&, , , lCategoryID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "tsWithLimit_NewFunction"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtExitPct_Change
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtExitPct_Change()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.txtExitPct_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtNumBars_Change
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtNumBars_Change()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.txtNumBars_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtOverride_Change
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtOverride_Change()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.txtOverride_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtRequiredMod_Change
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtRequiredMod_Change()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.txtRequiredMod_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadActionCombo
'' Description: Load up the order action combo box
'' Inputs:      Order Action
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadActionCombo(ByVal nOrderAction As eGDOrderAction)
On Error GoTo ErrSection:

    cboAction.Clear
    
    cboAction.AddItem "Long Entry"
    cboAction.ItemData(cboAction.NewIndex) = eGDOrderAction_LongEntry
    If nOrderAction = eGDOrderAction_LongEntry Then
        cboAction.ListIndex = cboAction.NewIndex
    End If
    
    cboAction.AddItem "Long Exit"
    cboAction.ItemData(cboAction.NewIndex) = eGDOrderAction_LongExit
    If nOrderAction = eGDOrderAction_LongExit Then
        cboAction.ListIndex = cboAction.NewIndex
    End If
    
    cboAction.AddItem "Short Entry"
    cboAction.ItemData(cboAction.NewIndex) = eGDOrderAction_ShortEntry
    If nOrderAction = eGDOrderAction_ShortEntry Then
        cboAction.ListIndex = cboAction.NewIndex
    End If
    
    cboAction.AddItem "Short Exit"
    cboAction.ItemData(cboAction.NewIndex) = eGDOrderAction_ShortExit
    If nOrderAction = eGDOrderAction_ShortExit Then
        cboAction.ListIndex = cboAction.NewIndex
    End If
    
    If nOrderAction <> -1& Then
        cboAction.Enabled = False
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrder.LoadActionCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectAction
'' Description: Select the action in the combo box based on the paramaters
'' Inputs:      Buy?, Enter?
'' Returns:     Index in the combo (-1 if not found, -2 if disabled)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectAction(ByVal bBuy As Boolean, ByVal bEnter As Boolean) As Long
On Error GoTo ErrSection:

    Dim nOrderAction As eGDOrderAction  ' Order action
    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = -2&
    If cboAction.Enabled = True Then
        If bBuy = True Then
            If bEnter = True Then
                nOrderAction = eGDOrderAction_LongEntry
            Else
                nOrderAction = eGDOrderAction_ShortExit
            End If
        Else
            If bEnter = True Then
                nOrderAction = eGDOrderAction_ShortEntry
            Else
                nOrderAction = eGDOrderAction_LongExit
            End If
        End If
        
        lReturn = -1&
        For lIndex = 0 To cboAction.ListCount - 1
            If cboAction.ItemData(lIndex) = nOrderAction Then
                cboAction.ListIndex = lIndex
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End If
    
    SelectAction = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.SelectAction"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrderTypeCombo
'' Description: Load up the order type combo box
'' Inputs:      Order Type, Allow Opposite Type?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadOrderTypeCombo(ByVal nOrderType As eTT_OrderType, ByVal bAllowOppositeOrderType As Boolean)
On Error GoTo ErrSection:

    cboOrderType.Clear
    
    cboOrderType.AddItem "Market"
    cboOrderType.ItemData(cboOrderType.NewIndex) = eTT_OrderType_Market
    If nOrderType = eTT_OrderType_Market Then
        cboOrderType.ListIndex = cboOrderType.NewIndex
    End If
    
    If ((nOrderType <> eTT_OrderType_Stop) And (nOrderType <> eTT_OrderType_StopWithLimit)) Or (bAllowOppositeOrderType = True) Then
        cboOrderType.AddItem "Limit"
        cboOrderType.ItemData(cboOrderType.NewIndex) = eTT_OrderType_Limit
        If nOrderType = eTT_OrderType_Limit Then
            cboOrderType.ListIndex = cboOrderType.NewIndex
        End If
    End If
    
    If (nOrderType <> eTT_OrderType_Limit) Or (bAllowOppositeOrderType = True) Then
        cboOrderType.AddItem "Stop"
        cboOrderType.ItemData(cboOrderType.NewIndex) = eTT_OrderType_Stop
        If nOrderType = eTT_OrderType_Stop Then
            cboOrderType.ListIndex = cboOrderType.NewIndex
        End If
        
        cboOrderType.AddItem "Stop with Limit"
        cboOrderType.ItemData(cboOrderType.NewIndex) = eTT_OrderType_StopWithLimit
        If nOrderType = eTT_OrderType_StopWithLimit Then
            cboOrderType.ListIndex = cboOrderType.NewIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.LoadOrderTypeCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectOrderType
'' Description: Select the order type in the combo box based on the paramaters
'' Inputs:      Order Type
'' Returns:     Index in the combo
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectOrderType(ByVal nOrderType As eTT_OrderType) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = -1&
    For lIndex = 0 To cboOrderType.ListCount - 1
        If cboOrderType.ItemData(lIndex) = nOrderType Then
            cboOrderType.ListIndex = lIndex
            lReturn = lIndex
            Exit For
        End If
    Next lIndex
    
    SelectOrderType = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.SelectOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ControlsFromObject
'' Description: Fill in the controls from the given object
'' Inputs:      Trade Sense Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ControlsFromObject(TradeSenseOrder As cTradeSenseOrder)
On Error GoTo ErrSection:

    With TradeSenseOrder
        If Len(.ConditionCoded) > 0 Then
            If Trim(.ConditionCoded) = "~24002IF ~03004True ~16001( ~17001)" Then
                optNoCondition.Value = True
                tsCondition.Text = ""
            Else
                optCondition.Value = True
                tsCondition.TextRTF = .ConditionRTF
            End If
            tsPrice.TextRTF = .PriceRTF
            tsWithLimit.TextRTF = .WithLimitRTF
        
            SelectAction .Buy, .Entry
            m.ExitPct.Price = .ExitPercent
        
            SelectOrderType .OrderType
            SetPriceEditorLabels
            
            cboBarPeriod.Text = FixPeriod(.BarPeriod)
            optAutoDetect.Value = Not .Override
            optOverride.Value = .Override
            txtNumBars.Text = Str(.NumBarsAuto)
            txtOverride.Text = Str(.NumBarsOverride)
            If .IntraBar Then
                cboEvaluate.ListIndex = 1
            Else
                cboEvaluate.ListIndex = 0
            End If
            
            If m.bForAutoExit Then
                CheckBoxValue(chkExpireDay) = False
            Else
                CheckBoxValue(chkExpireDay) = .ExpireDay
            End If
            CheckBoxValue(chkAllowInputs) = .AllowInputs
            txtRequiredMod = .RequiredMod
            
            CheckBoxValue(chkWrongSide) = .MarketIfWrongSide
            
            Set m.Inputs = .Inputs
            LoadInputsGrid
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.ControlsFromObject"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ObjectFromControls
'' Description: Fill in the object from the controls
'' Inputs:      Trade Sense Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ObjectFromControls(TradeSenseOrder As cTradeSenseOrder)
On Error GoTo ErrSection:

    Dim nAction As eGDOrderAction       ' Order action from the combo box
    Dim lIndex As Long                  ' Index into a for loop

    With TradeSenseOrder
        .ConditionFormatted = m.strFormattedCondition
        .ConditionCoded = m.strCodedCondition
        .PriceFormatted = m.strFormattedPrice
        .PriceCoded = m.strCodedPrice
        .WithLimitFormatted = m.strFormattedWithLimit
        .WithLimitCoded = m.strCodedWithLimit
        
        .Buy = IsBuy
        .Entry = IsEntry
        
        If cboOrderType.ListIndex > -1& Then
            .OrderType = cboOrderType.ItemData(cboOrderType.ListIndex)
        End If
        
        .ExitPercent = CLng(Val(txtExitPct.Text))
        
        .BarPeriod = cboBarPeriod.Text
        .Override = optOverride.Value
        .NumBarsAuto = CLng(Val(txtNumBars.Text))
        .NumBarsOverride = CLng(Val(txtOverride.Text))
        .IntraBar = (cboEvaluate.ItemData(cboEvaluate.ListIndex) = eGDEvaluate_IntraBar)
        
        .ExpireDay = CheckBoxValue(chkExpireDay)
        .AllowInputs = CheckBoxValue(chkAllowInputs)
        .Inputs = m.Inputs
        .RequiredMod = Trim(txtRequiredMod.Text)
        
        If chkWrongSide.Visible Then
            .MarketIfWrongSide = CheckBoxValue(chkWrongSide)
        Else
            .MarketIfWrongSide = False
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.ObjectFromControls"
    
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
    
    Dim bIsExit As Boolean              ' Is this currently set to an exit?
    
    tbToolbar.Visible = Not m.bForAutoExit
    fraButtons.Visible = m.bForAutoExit
    fraAdvanced.Visible = (Not m.bForAutoExit) And (ShowAdvancedTSOG = True)

    chkExpireDay.Visible = Not m.bForAutoExit
    txtNumBars.Visible = (optAutoDetect.Value = True)
    txtOverride.Visible = (optOverride.Value = True)
    chkWrongSide.Visible = ShowMarketWrongSide
    
    tbToolbar.Tools("ID_Save").Enabled = Dirty
    tbToolbar.Tools("ID_Verify").Enabled = Not Verified
    
    bIsExit = Not IsEntry
    Enable lblExit, bIsExit
    Enable sbExitPct, bIsExit
    Enable txtExitPct, bIsExit
    Enable lblExitPct, bIsExit
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitializeEditor
'' Description: Initialize the given editor
'' Inputs:      Editor
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitializeEditor(tsEditor As Editor, ByVal bDisableEnterKey As Boolean)
On Error GoTo ErrSection:

    With tsEditor
        .AppPath = App.Path
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = bDisableEnterKey
        ' DAJ 05/01/2012: After conversations with Pete and Tim, we decided that we
        ' can allow criteria and charting functions for TradeSense orders...
        .Usage = 14 ' 2
        .TurnOnEditing
        .Refresh
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.InitializeEditor"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateEditor
'' Description: Update the given editor with the given expression
'' Inputs:      Editor, Expression
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateEditor(tsEditor As Editor, ByVal strExpression As String)
On Error GoTo ErrSection:

    Dim Rule As New cRule               ' Rule object for building RTF text
    
    With tsEditor
        .TurnOffEditing
        If Len(strExpression) > 0 Then
            .TextRTF = Rule.GetRTF(strExpression)
        Else
            .TextRTF = ""
        End If
        .ExprIsFormatted = True
        .TurnOnEditing
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.UpdateEditor"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionOfIf
'' Description: Position of the IF in the editor text
'' Inputs:      None
'' Returns:     Position of IF if there, else zero
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PositionOfIf() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lPos As Long                    ' Position of IF in the string
    Dim strText As String               ' Text to look through
    Dim strCharBefore As String         ' Character before
    Dim strCharAfter As String          ' Character after
    
    lReturn = 0&
    strText = UCase(" " & tsCondition.Text & " ")
    
    lPos = InStr(strText, "IF")
    Do While lPos > 0
        strCharBefore = Mid(strText, lPos - 1, 1)
        strCharAfter = Mid(strText, lPos + 2, 1)
        
        If (strCharBefore = " ") Or (strCharBefore = vbCr) Or (strCharBefore = vbLf) Then
            If (strCharAfter = " ") Or (strCharAfter = vbCr) Or (strCharAfter = vbLf) Then
                lReturn = lPos
                Exit Do
            End If
        End If
    
        lPos = InStr(lPos + 1, strText, "IF")
    Loop
    
    PositionOfIf = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.PositionOfIf"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsBuy
'' Description: Determine if the Trade Sense order is to buy or to sell
'' Inputs:      None
'' Returns:     True if Buy, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsBuy() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nAction As eGDOrderAction       ' Order action
    
    bReturn = True
    If cboAction.ListIndex > -1& Then
        nAction = cboAction.ItemData(cboAction.ListIndex)
        bReturn = (nAction = eGDOrderAction_LongEntry) Or (nAction = eGDOrderAction_ShortExit)
    End If
    
    IsBuy = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.IsBuy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsEntry
'' Description: Determine if the Trade Sense order is an entry or an exit
'' Inputs:      None
'' Returns:     True if Entry, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsEntry() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nAction As eGDOrderAction       ' Order action
    
    bReturn = True
    If cboAction.ListIndex > -1& Then
        nAction = cboAction.ItemData(cboAction.ListIndex)
        bReturn = (nAction = eGDOrderAction_LongEntry) Or (nAction = eGDOrderAction_ShortEntry)
    End If
    
    IsEntry = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.IsEntry"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildRule
'' Description: Builds one expression out of the given parameters
'' Inputs:      Condition, Buy, Order Price, Order Type, With Limit Price
'' Returns:     Expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildRule(ByVal strCondition As String, ByVal bBuy As Boolean, ByVal strPrice As String, ByVal nOrderType As eTT_OrderType, ByVal strWithLimit As String) As String
On Error GoTo ErrSection:
    
    Dim strRule As String               ' Expression to return from the function
    Dim lPos As Long                    ' Position in the string
    Dim astrRule As cGdArray            ' Array of expressions for the rule
    Dim lIndex As Long                  ' Index into a for loop
    
    ' IF Condition THEN Action...
    strRule = Trim(strCondition)
    lPos = InStr(UCase(strRule), "IF")
    If Left(UCase(strRule), 2) <> "IF" And Left(strRule, 1) <> "'" Then
        If lPos <> 0 Then
            strRule = strCondition
        ElseIf InStr(strRule, ":=") = 0 Then
            strRule = "IF " & strCondition
        Else
            Set astrRule = New cGdArray
            astrRule.SplitFields strRule, vbLf
            For lIndex = 0 To astrRule.Size - 1
                If InStr(astrRule(lIndex), ":=") = 0 Then
                    astrRule(lIndex) = "IF " & astrRule(lIndex)
                    Exit For
                End If
            Next lIndex
            
            strRule = astrRule.JoinFields(vbCrLf)
        End If
    End If
    
    If bBuy = True Then
        strRule = strRule & " THEN " & vbCrLf & vbTab & "BUY ("
    Else
        strRule = strRule & " THEN " & vbCrLf & vbTab & "SELL ("
    End If
    
    ' Order Price...
    If Len(strPrice) > 0 Then
        If (nOrderType <> eTT_OrderType_Market) And (nOrderType <> eTT_OrderType_MarketOnClose) Then
            strRule = strRule & strPrice & ", "
        End If
    End If
    
    ' Order type...
    Select Case nOrderType
        Case eTT_OrderType_Stop
            strRule = strRule & Chr(34) & "Stop" & Chr(34)
        Case eTT_OrderType_Limit
            strRule = strRule & Chr(34) & "Limit" & Chr(34)
        Case eTT_OrderType_Market
            strRule = strRule & "Close, " & Chr(34) & "Market" & Chr(34)
        Case eTT_OrderType_StopWithLimit
            strRule = strRule & Chr(34) & "Stop with Limit" & Chr(34)
    End Select
    
    ' With Limit Price
    If Len(strWithLimit) > 0 Then
        If (nOrderType <> eTT_OrderType_Market) And (nOrderType <> eTT_OrderType_MarketOnClose) Then
            strRule = strRule & ", " & strWithLimit
        End If
    End If
    
    BuildRule = strRule & ")" & vbCrLf & "ENDIF"
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmTradeSenseOrder.BuildRule"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SplitRule
'' Description: Split the big expression into pieces
'' Inputs:      Expression, Buy, Order Type, Condition, Price, With Limit Price
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SplitRule(ByVal strExpression As String, bBuy As Boolean, nOrderType As eTT_OrderType, strCondition As String, strPrice As String, strWithLimit As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIfPos As Long                  ' Position of the 'IF' token in the expression
    Dim lThenPos As Long                ' Position of the 'THEN' token in the expression
    Dim lEnterPos As Long               ' Position of the Enter token in the expression
    Dim lPricePos As Long               ' Position of the price expression in the big expression
    Dim lCommaPos As Long               ' Position of the next comma in the expression
    Dim lCommaPos2 As Long              ' Position of another comma in the expression
    Dim lTypePos As Long                ' Position of the order type in the expression
    Dim lRightParenPos As Long          ' Position of the right paren in the expression
    Dim strOrderType As String          ' Order type in the expression
    Dim lParens As Long                 ' Number of open parens
    Dim lCommas As Long                 ' Number of commas
    Dim lCurPos As Long                 ' Current position in the string
    
    ' First, look for the 'IF', 'THEN', and Enter tokens in the expression...
    bReturn = False
    lIfPos = InStr(strExpression, "~" & Format(kIfToken, "00"))
    lThenPos = InStr(lIfPos + 1, strExpression, "~" & Format(kThenToken, "00"))
    lEnterPos = InStr(lThenPos + 1, strExpression, "~" & Format(kEnterToken, "00"))
    
    If (lIfPos > 0) And (lThenPos > 0) And (lEnterPos > 0) Then
        strCondition = Left(strExpression, lThenPos - 1)
        bBuy = (InStr(lEnterPos + 1, UCase(strExpression), "~01003BUY") > 0)
        
        ' Look for the left paren that starts the price section...
        lPricePos = InStr(lEnterPos + 1, strExpression, "~" & Format(kLeftParenToken, "00"))
        If lPricePos > 0 Then
            ' Move past the left paren...
            lPricePos = InStr(lPricePos + 1, strExpression, "~")
            
            ' If the price is set to Market1, skip it...
            If UCase(Mid(strExpression, lPricePos + 6, 7)) = "MARKET1" Then
                lPricePos = InStr(lPricePos + 21, strExpression, "~")
            End If
            
            lParens = 1
            lCommas = 0
            lCurPos = lPricePos + 1
            Do Until (lCurPos > Len(strExpression)) Or (lCurPos = 0)
                lCurPos = InStr(lCurPos, strExpression, "~")
                If lCurPos > 0 Then
                    Select Case CLng(Val(Mid(strExpression, lCurPos + 1, 2)))
                        Case kLeftParenToken
                            lParens = lParens + 1
                            
                        Case kRightParenToken
                            lParens = lParens - 1
                            If lParens = 0 Then
                                lRightParenPos = lCurPos
                            End If
                            
                        Case kCommaToken
                            If lParens = 1 Then
                                lCommas = lCommas + 1
                                If lCommas = 1 Then
                                    lCommaPos = lCurPos
                                Else
                                    lCommaPos2 = lCurPos
                                End If
                            End If
                            
                    End Select
                    
                    lCurPos = lCurPos + 1
                End If
            Loop
            
            strPrice = Mid(strExpression, lPricePos, lCommaPos - lPricePos)
            
            lTypePos = InStr(lCommaPos + 1, strExpression, "~")
            If lCommaPos2 > 0 Then
                strOrderType = Mid(strExpression, lTypePos, lCommaPos2 - lTypePos)
            Else
                strOrderType = Mid(strExpression, lTypePos, lRightParenPos - lTypePos)
            End If
            
            strOrderType = Mid(strOrderType, 7, Val(Mid(strOrderType, 4, 3)))
            strOrderType = Replace(strOrderType, Chr(34), "")
            
            Select Case UCase(strOrderType)
                Case "MARKET"
                    nOrderType = eTT_OrderType_Market
                Case "STOP"
                    nOrderType = eTT_OrderType_Stop
                Case "LIMIT"
                    nOrderType = eTT_OrderType_Limit
                Case "STOP WITH LIMIT"
                    nOrderType = eTT_OrderType_StopWithLimit
                Case "MARKET ON CLOSE"
                    nOrderType = eTT_OrderType_MarketOnClose
                Case "STOP CLOSE ONLY"
                    nOrderType = eTT_OrderType_StopCloseOnly
                Case "LIMIT CLOSE ONLY"
                    nOrderType = eTT_OrderType_LimitCloseOnly
                Case "STOP WITH LIMIT CLOSE ONLY"
                    nOrderType = eTT_OrderType_StopWithLimitCloseOnly
                Case Else
                    nOrderType = eTT_OrderType_Market
            End Select
            
            If lCommaPos2 > 0 Then
                lPricePos = InStr(lCommaPos2 + 1, strExpression, "~")
                strWithLimit = Mid(strExpression, lPricePos, lRightParenPos - lPricePos)
            End If
            
            bReturn = True
        End If
    End If
    
    SplitRule = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.SplitRule"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyControls
'' Description: Verify the controls
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyControls() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (optCondition.Value = True) Then
        If (Len(Trim(tsCondition.Text)) = 0) Or (UCase(Trim(tsCondition.Text)) = "IF") Or (UCase(Trim(tsCondition.Text)) = "IF TRUE") Then
            optNoCondition.Value = True
            tsCondition.Text = ""
        End If
    End If
    
    If (tsPrice.Visible = True) And (Len(Trim(tsPrice.Text)) = 0) Then
        InfBox "You must enter an order price expression", "!", , kErrorCaption
        MoveFocus tsPrice
    ElseIf (tsPrice.Visible = True) And (IsBooleanExpression(tsPrice.Text) = True) Then
        InfBox "The order price expression cannot result in a True/False value", "!", , kErrorCaption
        MoveFocus tsPrice
    ElseIf (tsWithLimit.Visible = True) And (Len(Trim(tsWithLimit.Text)) = 0) Then
        InfBox "You must enter a with limit price expression", "!", , kErrorCaption
        MoveFocus tsWithLimit
    ElseIf (tsWithLimit.Visible = True) And (IsBooleanExpression(tsWithLimit.Text) = True) Then
        InfBox "The with limit price expression cannot result in a True/False value", "!", , kErrorCaption
        MoveFocus tsWithLimit
    ElseIf cboAction.ListIndex = -1& Then
        InfBox "You must select an order action", "!", , kErrorCaption
        MoveFocus cboAction
    ElseIf cboOrderType.ListIndex = -1& Then
        InfBox "You must select an order type", "!", , kErrorCaption
        MoveFocus cboOrderType
    ElseIf (optNoCondition.Value = True) And (cboOrderType.ItemData(cboOrderType.ListIndex) = eTT_OrderType_Market) Then
        InfBox "You cannot select a Market order unless you specify a condition.", "!", , kErrorCaption
        MoveFocus cboOrderType
    ElseIf (optNoCondition.Value = True) And (cboEvaluate.ItemData(cboEvaluate.ListIndex) = eGDEvaluate_IntraBar) Then
        InfBox "You cannot evaluate a condition intra-bar unless you specify a condition.", "!", , kErrorCaption
        MoveFocus cboEvaluate
    Else
        bReturn = True
    End If
    
    VerifyControls = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.VerifyControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyExpression
'' Description: Verify the expression
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyExpression() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim Expr As cExpression             ' Expression object for verifying
    Dim nOrderType As eTT_OrderType     ' Order type
    Dim lIndex As Long                  ' Index into a for loop
    Dim bContinue As Boolean            ' Do we want to continue?
    Dim bBuy As Boolean                 ' Buy/Sell
    Dim strCondition As String          ' Condition part of the expression
    Dim strPrice As String              ' Price part of the expression
    Dim strWithLimit As String          ' With Limit part of the expression
    Dim strFormatted As String          ' Formatted expression
    Dim strRule As String               ' Rule built from the pieces

    bReturn = False
    If VerifyControls = True Then
        If optCondition Then
            tsCondition.Text = FixPeriodInMarkets(tsCondition.Text)
        End If
        tsPrice.Text = FixPeriodInMarkets(tsPrice.Text)
        tsWithLimit.Text = FixPeriodInMarkets(tsWithLimit.Text)
        
        Set Expr = New cExpression
        With Expr
            .PortfolioNavigator = False
            .Functions = g.Functions
            
            nOrderType = cboOrderType.ItemData(cboOrderType.ListIndex)
            If optCondition Then
                strRule = BuildRule(tsCondition.Text, IsBuy, tsPrice.Text, nOrderType, tsWithLimit.Text)
            Else
                strRule = BuildRule("If True", IsBuy, tsPrice.Text, nOrderType, tsWithLimit.Text)
            End If
            .ValidateRule strRule
            
            strFormatted = .EditText
            
            bContinue = CheckExpression(strFormatted, .Inputs, .GetFIDs)
            If SplitRule(strFormatted, bBuy, nOrderType, strCondition, strPrice, strWithLimit) Then
                If Trim(strCondition) = "~24002IF ~03004True" Then
                    UpdateEditor tsCondition, ""
                Else
                    UpdateEditor tsCondition, strCondition
                End If
                UpdateEditor tsPrice, strPrice
                UpdateEditor tsWithLimit, strWithLimit
                
                m.strFormattedCondition = strCondition
                m.strFormattedPrice = strPrice
                m.strFormattedWithLimit = strWithLimit
                
                ' Only handle the coded text if all is OK with the expression...
                If bContinue = True Then
                    If chkAllowInputs.Value = vbChecked Then
                        UpdateInputs .Inputs
                        LoadInputsGrid
                    End If
                    If SplitRule(.CodedText, bBuy, nOrderType, strCondition, strPrice, strWithLimit) Then
                        m.strCodedCondition = strCondition
                        m.strCodedPrice = strPrice
                        m.strCodedWithLimit = strWithLimit
                        
                        AutoDetect
                        Verified = True
                        bReturn = True
                    End If
                End If
            End If
        End With
    End If
    
    VerifyExpression = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.VerifyExpression"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetect
'' Description: Attempt to auto detect the number of bars required
'' Inputs:      None
'' Returns:     Number of Bars required (-1 if cannot calculate)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AutoDetect() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim AD As New cAutoDetect           ' Auto detection object
    Dim strCodedText As String          ' Coded text expression
    Dim lCondition As Long              ' Number of bars required for the condition
    Dim lPrice As Long                  ' Number of bars required for the price
    Dim lWithLimit As Long              ' Number of bars required for the with limit price
    Dim bAllGood As Boolean             ' Did all of the expressions evaluate?
    Dim strPeriod As String             ' Period to use for the auto detect
    
    If UCase(cboBarPeriod.Text) = "AUTO BREAKOUT" Or UCase(cboBarPeriod.Text) = "FRACTZEN" Then
        strPeriod = GetPeriodStr("10b")
    Else
        strPeriod = cboBarPeriod.Text
    End If
    
    strCodedText = m.lateCalc.FixExpression(m.strCodedCondition)
    lCondition = AD.AutoDetect(strCodedText, , strPeriod)
    If tsPrice.Visible Then
        strCodedText = m.lateCalc.FixExpression(m.strCodedPrice)
        lPrice = AD.AutoDetect(strCodedText, , strPeriod)
    Else
        lPrice = 1
    End If
    If tsWithLimit.Visible Then
        strCodedText = m.lateCalc.FixExpression(m.strCodedWithLimit)
        lWithLimit = AD.AutoDetect(strCodedText, , strPeriod)
    Else
        lWithLimit = 1
    End If
   
    lReturn = -1
    bAllGood = True
    
    If lCondition <= 0 Then
        bAllGood = False
    ElseIf lCondition > lReturn Then
        lReturn = lCondition
    End If
    If lPrice <= 0 Then
        bAllGood = False
    ElseIf lPrice > lReturn Then
        lReturn = lPrice
    End If
    If lWithLimit <= 0 Then
        bAllGood = False
    ElseIf lWithLimit > lReturn Then
        lReturn = lWithLimit
    End If
        
    txtNumBars.Text = Str(lReturn)
    
    If (optOverride.Value = True) And (ValOfText(txtOverride.Text) < lReturn) And (lReturn > 0) Then
        InfBox "Trade Navigator has determined that your Trade Sense order needs at least " & _
            Trim(CStr(lReturn)) & " bars to run properly.  " & _
            "The value has been set accordingly.", _
            "i", , "Trade Sense Order"
        optAutoDetect = True
        txtOverride.Text = Str(lReturn)
    End If
    
    If (bAllGood = False) And ((optAutoDetect.Value = True) Or (ValOfText(txtOverride.Text) <= 0)) Then
        InfBox "Trade Navigator could not automatically determine how many bars are needed to calculate " & _
                " the Trade Sense order.  Please specify an override for the number of necessary bars.", _
                "!", , kErrorCaption
        optOverride = True
        
        If lReturn > ValOfText(txtOverride.Text) Then
            txtOverride.Text = Str(lReturn)
        End If
        
        MoveFocus txtOverride
    End If
    
    AutoDetect = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.AutoDetect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckRefs
'' Description: Check the function references for indirect late calculating
'' Inputs:      Function Refs Handle
'' Returns:     Bad function names
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckRefs(ByVal hFunctionRefs As Long, astrNextBar As cGdArray) As cGdArray
On Error GoTo ErrSection:

    Dim alFunctions As New cGdArray     ' Array of function refrences for the expression
    Dim astrReturn As New cGdArray      ' Array of bad function names
    Dim lIndex As Long                  ' Index into a for loop
    Dim fnFunction As cFunction         ' Function object
    Dim bIncludeNextBarOpen As Boolean  ' Include the next bar open in the check?
    
    ' 11/08/2011 DAJ: We want to allow Next Bar Open if the user has chosen to evaluate the
    ' expression on each new bar, so we don't want to include it in the Next Bar Functions
    ' reference in that case...
    bIncludeNextBarOpen = (cboEvaluate.ItemData(cboEvaluate.ListIndex) = eGDEvaluate_IntraBar)
    
    astrReturn.Create eGDARRAY_Strings
    If alFunctions.CopyFromHandle(hFunctionRefs) Then
        For lIndex = 0 To alFunctions.Size - 1
            Set fnFunction = g.Functions.Item(Str(alFunctions(lIndex)))
            If Not fnFunction Is Nothing Then
                If fnFunction.LateCalculating = True Then
                    If IsGenesisLibrary(fnFunction.LibraryID) = False Then
                        astrReturn.Add fnFunction.FunctionName
                    End If
                ElseIf UsesNextBarFunctions(fnFunction.FunctionID, bIncludeNextBarOpen) Then
                    astrNextBar.Add fnFunction.FunctionName
                End If
            End If
        Next lIndex
    End If
    
    Set CheckRefs = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.CheckRefs"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetFunctionsAsError
'' Description: Change the expression to set the given functions as an error
'' Inputs:      Expression, Names
'' Returns:     New Expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetFunctionsAsError(ByVal strExpression As String, ByVal astrNames As cGdArray) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strToken As String              ' Token to look up
    Dim strError As String              ' Token to replace with
    Dim strReturn As String             ' Return value for the function
    
    strReturn = strExpression
    For lIndex = 0 To astrNames.Size - 1
        strError = "~18" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
        
        strToken = "~01" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
        strReturn = Replace(strReturn, strToken, strError)
        strToken = "~02" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
        strReturn = Replace(strReturn, strToken, strError)
        strToken = "~03" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
        strReturn = Replace(strReturn, strToken, strError)
        strToken = "~04" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
        strReturn = Replace(strReturn, strToken, strError)
        strToken = "~30" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
        strReturn = Replace(strReturn, strToken, strError)
    Next lIndex
    
    SetFunctionsAsError = strReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmTradeSenseOrder.SetFunctionsAsError"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetParmsAsError
'' Description: Change the expression to set the given parameters as an error
'' Inputs:      Expression, Names, Markets?
'' Returns:     New Expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetParmsAsError(ByVal strExpression As String, ByVal astrNames As cGdArray, ByVal bMarket As Boolean) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strToken As String              ' Token to look up
    Dim strError As String              ' Token to replace with
    Dim strReturn As String             ' Return value for the function
    
    strReturn = strExpression
    For lIndex = 0 To astrNames.Size - 1
        strError = "~18" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
        
        If bMarket Then
            strToken = "~07" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
            strReturn = Replace(strReturn, strToken, strError)
        Else
            strToken = "~05" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
            strReturn = Replace(strReturn, strToken, strError)
            strToken = "~06" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
            strReturn = Replace(strReturn, strToken, strError)
            strToken = "~25" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
            strReturn = Replace(strReturn, strToken, strError)
            strToken = "~27" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
            strReturn = Replace(strReturn, strToken, strError)
            strToken = "~28" & Format(Len(astrNames(lIndex)), "000") & astrNames(lIndex)
            strReturn = Replace(strReturn, strToken, strError)
        End If
    Next lIndex
    
    SetParmsAsError = strReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmTradeSenseOrder.SetParmsAsError"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckExpression
'' Description: Check the expression for indirect late calculating refrences,
''              non-market inputs, or invalid market inputs
'' Inputs:      Expression, Inputs, Function References
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckExpression(strExpression As String, ByVal expressionParms As cInputs, ByVal hFunctionRefs As Long) As Boolean
On Error GoTo ErrSection:

    Dim astrLateCalc As New cGdArray    ' Functions that use late calculating function
    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrNonMarkets As New cGdArray  ' Non-market inputs
    Dim astrBadMarkets As New cGdArray  ' Bad market inputs
    Dim astrNextBar As New cGdArray     ' Functions that require "next bar"

    bReturn = True
    
    Set astrLateCalc = CheckRefs(hFunctionRefs, astrNextBar)
    If astrLateCalc.Size > 0 Then
        strExpression = SetFunctionsAsError(strExpression, astrLateCalc)
        InfBox "You cannot use functions that use late calculating functions.  You must reference these late calculating functions directly", "!", , kErrorCaption
        bReturn = False
    ElseIf astrNextBar.Size > 0 Then
        strExpression = SetFunctionsAsError(strExpression, astrNextBar)
        If cboEvaluate.ItemData(cboEvaluate.ListIndex) = eGDEvaluate_IntraBar Then
            InfBox "You cannot use functions that reference|next bar functions when evaluating intra-bar", "!", , kErrorCaption
        Else
            InfBox "You cannot use functions that reference|next bar functions other than Next Bar Open|when evaluating on each new bar", "!", , kErrorCaption
        End If
        bReturn = False
    End If
    
    If bReturn = True Then
        If Not expressionParms Is Nothing Then
            For lIndex = 1 To expressionParms.Count
                If expressionParms.Item(lIndex).ParmTypeID <> kSN_RetBars Then
                    astrNonMarkets.Add expressionParms.Item(lIndex).ParmName
                ElseIf IsValidMarket(expressionParms.Item(lIndex).ParmName) = False Then
                    astrBadMarkets.Add expressionParms.Item(lIndex).ParmName
                End If
            Next lIndex
        End If
        
        If (astrNonMarkets.Size > 0) And (chkAllowInputs.Value = vbUnchecked) Then
            strExpression = SetParmsAsError(strExpression, astrNonMarkets, False)
            InfBox "Inputs are not allowed in expression", "!", , kErrorCaption
            bReturn = False
        ElseIf astrBadMarkets.Size > 0 Then
            strExpression = SetParmsAsError(strExpression, astrBadMarkets, True)
            If astrBadMarkets.Size = 1 Then
                InfBox "No data can be loaded for '" & astrBadMarkets(0) & "'", "!", , kErrorCaption
            Else
                InfBox "No data could be loaded for markets", "!", , kErrorCaption
            End If
            bReturn = False
        End If
    End If
    
    CheckExpression = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.CheckExpression"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Allow the user to save the Trade Sense order group
'' Inputs:      Save Command
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save(ByVal nSaveCmd As eGDSaveCmd)
On Error GoTo ErrSection:

    Dim strText As String               ' Text for the ask box
    Dim strHeader As String             ' Header for the ask box
    Dim strNewName As String            ' New name for the Trade Sense Order Group
    Dim strID As String                 ' ID for the new name
    Dim Orders As cTradeSenseOrders     ' Collection of existing TradeSense orders

    If VerifyExpression Then
        If Len(m.tsOrder.Name) = 0 Then
            strText = "Save the current Trade Sense Order as..."
            strHeader = "Save"
            strNewName = AskBox("h=" & strHeader & " ; i=? ; g=string ; d=" & m.tsOrder.Name & " ; " & strText)
        ElseIf nSaveCmd = eGDSaveCmd_SaveAs Then
            strText = "Save a copy of the current Trade Sense Order as..."
            strHeader = "Save As"
            strNewName = AskBox("h=" & strHeader & " ; i=? ; g=string ; d=" & m.tsOrder.Name & " ; " & strText)
        ElseIf nSaveCmd = eGDSaveCmd_Rename Then
            strText = "Rename the current Trade Sense Order as..."
            strHeader = "Rename"
            strNewName = AskBox("h=" & strHeader & " ; i=? ; g=string ; d=" & m.tsOrder.Name & " ; " & strText)
        Else
            strNewName = m.tsOrder.Name
        End If
        
        Set Orders = New cTradeSenseOrders
        Orders.Load
        
        Do While (Len(strNewName) > 0) And (strNewName <> m.tsOrder.Name)
            strID = Orders.IdForName(strNewName)
            
            If (Len(strID) > 0) And (strID <> m.tsOrder.ID) Then
                InfBox "'" & strNewName & "' already exists.  Please select a new name", "!", , "Save Error"
            ElseIf IsValidFileBase(strNewName, False) = False Then
                InfBox "'" & strNewName & "' is not a valid name.  Please select a new name", "!", , "Save Error"
            Else
                Exit Do
            End If
            
            strNewName = AskBox("h=" & strHeader & " ; i=? ; g=string ; d=" & m.tsOrder.Name & " ; " & strText)
        Loop
        
        If Len(strNewName) > 0 Then
            ObjectFromControls m.tsOrder
            If strNewName <> m.tsOrder.Name Then
                If nSaveCmd = eGDSaveCmd_SaveAs Then
                    m.tsOrder.ClearID
                End If
                m.tsOrder.Name = strNewName
                SetEditorCaption Me, "Trade Sense Order", m.tsOrder.Name
            End If
            m.tsOrder.ToFile
            Dirty = False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExitForm
'' Description: Exit the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExitForm()
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from an InfBox
    Dim bHide As Boolean                ' Hide the form and allow ShowMe to unload it?

    bHide = True
    If Dirty Then
        strReturn = InfBox("Do you want to save your changes?||Clicking No will undo any changes you have made to this Trade Sense order.|", "?", "+Yes|No|-Cancel", Caption)
        Select Case strReturn
            Case "C"
                bHide = False
                
            Case "Y"
                Save eGDSaveCmd_Save
                bHide = True
                m.bOK = True
                
            Case "N"
                bHide = True
                m.bOK = False
                
        End Select
    Else
        m.bOK = True
    End If
    
    If bHide Then
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.ExitForm"
    
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
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = False
        
        .Cols = 2
        .FixedCols = 0
        .Rows = 1
        .FixedRows = 1
        
        .TextMatrix(0, 0) = "Input Name"
        .TextMatrix(0, 1) = "Default Value"
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.InitInputsGrid"
    
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

    With fgInputs
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To m.Inputs.Count
            .Rows = .Rows + 1
            
            .RowData(.Rows - 1) = m.Inputs(lIndex)
            .TextMatrix(.Rows - 1, 0) = m.Inputs(lIndex).Name
            .TextMatrix(.Rows - 1, 1) = m.Inputs(lIndex).DefaultValue
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.LoadInputsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateInputs
'' Description: Update the inputs array
'' Inputs:      Inputs collection
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateInputs(ByVal Parms As cInputs)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrNonMarkets As cGdArray      ' Array of non-market parms
    Dim tsInput As cTradeSenseOrderInput ' Order input
    Dim strParmName As String           ' Parameter name

    Set astrNonMarkets = New cGdArray
    astrNonMarkets.Create eGDARRAY_Strings

    If Not Parms Is Nothing Then
        For lIndex = 1 To Parms.Count
            If Parms.Item(lIndex).ParmTypeID <> kSN_RetBars Then
                astrNonMarkets.Add Parms.Item(lIndex).ParmName & ";" & Str(Parms.Item(lIndex).ParmTypeID)
            End If
        Next lIndex
    End If
    
    If astrNonMarkets.Size = 0 Then
        m.Inputs.Clear
    Else
        astrNonMarkets.Sort
        
        For lIndex = 0 To astrNonMarkets.Size - 1
            strParmName = Parse(astrNonMarkets(lIndex), ";", 1)
            If m.Inputs.Exists(strParmName) = False Then
                Set tsInput = New cTradeSenseOrderInput
                tsInput.Name = strParmName
                tsInput.ParmType = CLng(Val(Parse(astrNonMarkets(lIndex), ";", 2)))
                m.Inputs.Add tsInput
            End If
        Next lIndex
        
        For lIndex = m.Inputs.Count To 1 Step -1
            If astrNonMarkets.BinarySearch(m.Inputs(lIndex).Name & ";", , eGdSort_MatchUsingSearchStringLength) = False Then
                m.Inputs.Remove lIndex
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.UpdateInputs"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPriceEditorLabels
'' Description: Set the editor price labels based on Buy/Sell and order type
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPriceEditorLabels()
On Error GoTo ErrSection:
    
    Dim nOrderType As eTT_OrderType     ' Order type
    Dim bBuy As Boolean                 ' Buy/Sell
    Dim nAction As eGDOrderAction       ' Order action
    
    If (cboAction.ListIndex = -1&) Or (cboOrderType.ListIndex = -1&) Then
        lblPrice.Caption = "Order &Price:"
        lblWithLimit.Caption = "With &Limit Price:"
    Else
        nOrderType = cboOrderType.ItemData(cboOrderType.ListIndex)
        nAction = cboAction.ItemData(cboAction.ListIndex)
        bBuy = IsBuy
        
        Select Case nOrderType
            Case eTT_OrderType_Market
                lblPrice.Caption = "Order &Price:"
                lblWithLimit.Caption = "With &Limit Price:"
                
            Case eTT_OrderType_Limit
                If bBuy Then
                    lblPrice.Caption = "Limit &Price: Buy if price gets down to..."
                Else
                    lblPrice.Caption = "Limit &Price: Sell if price gets up to..."
                End If
            
            Case eTT_OrderType_Stop
                If bBuy Then
                    lblPrice.Caption = "Stop &Price: Buy if price gets up to..."
                Else
                    lblPrice.Caption = "Stop &Price: Sell if price gets down to..."
                End If
            
            Case eTT_OrderType_StopWithLimit
                If bBuy Then
                    lblPrice.Caption = "Stop &Price: Buy if price gets up to..."
                    lblWithLimit.Caption = "With &Limit Price: and if price is at or below..."
                Else
                    lblPrice.Caption = "Stop &Price: Sell if price gets down to..."
                    lblWithLimit.Caption = "With &Limit Price: and if price is at or above..."
                End If
            
        End Select
    End If
    
    chkWrongSide.Visible = ShowMarketWrongSide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrder.SetPriceEditorLabels"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPeriod
'' Description: Fix the bar period
'' Inputs:      Bar Period
'' Returns:     Fixed Bar Period
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixPeriod(ByVal strPeriod As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If UCase(strPeriod) = "AUTO BREAKOUT" Or UCase(strPeriod) = "FRACTZEN" Then
        strReturn = "FractZen" 'strPeriod
    Else
        strReturn = GetPeriodStr(strPeriod)
    End If
    
    FixPeriod = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.FixPeriod"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMarketWrongSide
'' Description: Determine whether to show the "wrong side" check box
'' Inputs:      None
'' Returns:     True to Show, False to Hide
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ShowMarketWrongSide() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nOrderType As eTT_OrderType     ' Order type
    Dim nAction As eGDOrderAction       ' Order action
    
    bReturn = False
    If m.bForAutoExit = False Then
        If (cboAction.ListIndex > -1&) And (cboOrderType.ListIndex > -1&) Then
            nOrderType = cboOrderType.ItemData(cboOrderType.ListIndex)
            nAction = cboAction.ItemData(cboAction.ListIndex)
            
            If (nOrderType = eTT_OrderType_Stop) Or (nOrderType = eTT_OrderType_StopWithLimit) Then
                bReturn = ((nAction = eGDOrderAction_LongExit) Or (nAction = eGDOrderAction_ShortExit))
            End If
        End If
    End If
    
    ShowMarketWrongSide = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrder.ShowMarketWrongSide"
    
End Function

