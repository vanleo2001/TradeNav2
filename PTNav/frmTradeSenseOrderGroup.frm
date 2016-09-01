VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTradeSenseOrderGroup 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraAdvanced 
      Height          =   2535
      Left            =   180
      TabIndex        =   4
      Top             =   3780
      Width           =   5655
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
      Caption         =   "frmTradeSenseOrderGroup.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeSenseOrderGroup.frx":0030
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrderGroup.frx":0050
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkFlattenOpposite 
         Height          =   400
         Left            =   180
         TabIndex        =   0
         Top             =   1980
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTradeSenseOrderGroup.frx":006C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroup.frx":019A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroup.frx":01BA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraGenesis 
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   1560
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
         Caption         =   "frmTradeSenseOrderGroup.frx":01D6
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTradeSenseOrderGroup.frx":0202
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroup.frx":0222
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkAllowManual 
            Height          =   220
            Left            =   3240
            TabIndex        =   16
            Top             =   60
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
            Caption         =   "frmTradeSenseOrderGroup.frx":023E
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTradeSenseOrderGroup.frx":028E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrderGroup.frx":02AE
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtRequiredMod 
            Height          =   315
            Left            =   1320
            TabIndex        =   15
            Top             =   0
            Width           =   1215
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTradeSenseOrderGroup.frx":02CA
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
            Tip             =   "frmTradeSenseOrderGroup.frx":02EA
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrderGroup.frx":030A
         End
         Begin HexUniControls.ctlUniLabelXP lblRequiredMod 
            Height          =   195
            Left            =   0
            Top             =   60
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
            Caption         =   "frmTradeSenseOrderGroup.frx":0326
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTradeSenseOrderGroup.frx":0368
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTradeSenseOrderGroup.frx":0388
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniCheckXP chkLinkInputs 
         Height          =   220
         Left            =   120
         TabIndex        =   5
         Top             =   240
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
         Caption         =   "frmTradeSenseOrderGroup.frx":03A4
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroup.frx":0402
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroup.frx":0422
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fgInputs 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   5415
         _cx             =   9551
         _cy             =   1720
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
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   5820
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "frmTradeSenseOrderGroup.frx":043E
      ToolBars        =   "frmTradeSenseOrderGroup.frx":056F
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2115
      Left            =   5880
      TabIndex        =   7
      Top             =   1080
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
      Caption         =   "frmTradeSenseOrderGroup.frx":067B
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeSenseOrderGroup.frx":06A7
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrderGroup.frx":06C7
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdNewOrder 
         Height          =   315
         Left            =   0
         TabIndex        =   8
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
         Caption         =   "frmTradeSenseOrderGroup.frx":06E3
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroup.frx":0717
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroup.frx":0737
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOco 
         Height          =   315
         Left            =   0
         TabIndex        =   13
         Top             =   1800
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
         Caption         =   "frmTradeSenseOrderGroup.frx":0753
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroup.frx":077B
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroup.frx":079B
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOto 
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   1440
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
         Caption         =   "frmTradeSenseOrderGroup.frx":07B7
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroup.frx":07DF
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroup.frx":07FF
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemoveOrder 
         Height          =   315
         Left            =   0
         TabIndex        =   11
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
         Caption         =   "frmTradeSenseOrderGroup.frx":081B
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroup.frx":0855
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroup.frx":0875
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEditOrder 
         Height          =   315
         Left            =   0
         TabIndex        =   10
         Top             =   720
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
         Caption         =   "frmTradeSenseOrderGroup.frx":0891
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroup.frx":08C7
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroup.frx":08E7
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddOrder 
         Height          =   315
         Left            =   0
         TabIndex        =   9
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
         Caption         =   "frmTradeSenseOrderGroup.frx":0903
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrderGroup.frx":0937
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrderGroup.frx":0957
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtbPreview 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2940
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1296
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTradeSenseOrderGroup.frx":0973
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
      Tip             =   "frmTradeSenseOrderGroup.frx":0993
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrderGroup.frx":09B3
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
   Begin VSFlex7LCtl.VSFlexGrid fgOrders 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5595
      _cx             =   9869
      _cy             =   2990
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
   Begin HexUniControls.ctlUniTextBoxXP txtDescription 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5595
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTradeSenseOrderGroup.frx":09CF
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
      Tip             =   "frmTradeSenseOrderGroup.frx":09EF
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrderGroup.frx":0A0F
   End
   Begin HexUniControls.ctlUniLabelXP lblDescription 
      Height          =   195
      Left            =   120
      Top             =   120
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
      Caption         =   "frmTradeSenseOrderGroup.frx":0A2B
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTradeSenseOrderGroup.frx":0A65
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrderGroup.frx":0A85
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuNewOrder 
         Caption         =   "New Order"
      End
      Begin VB.Menu mnuAddOrder 
         Caption         =   "Add Order"
      End
      Begin VB.Menu mnuEditOrder 
         Caption         =   "Edit Order"
      End
      Begin VB.Menu mnuRemoveOrder 
         Caption         =   "Remove Order"
      End
      Begin VB.Menu mnuOto 
         Caption         =   "OTO"
      End
      Begin VB.Menu mnuOco 
         Caption         =   "OCO"
      End
   End
End
Attribute VB_Name = "frmTradeSenseOrderGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTradeSenseOrderGroup.frm
'' Description: Form that handles a Trade Sense order group
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/15/2010   DAJ         Changed the form icon
'' 06/16/2010   DAJ         Allow for adding same order multiple times (#5800)
'' 06/17/2010   DAJ         Changed filenames to ID instead of name
'' 06/21/2010   DAJ         Fixes for OTOs/OCOs/Tree (#5800)
'' 06/22/2010   DAJ         Further fixed the OTOs (#5800), Rename issue (#5801)
'' 06/28/2010   DAJ         Added Use flag for turning off orders
'' 07/15/2010   DAJ         Added capabilities for inputs
'' 07/21/2010   DAJ         Fixed bug with OTO's
'' 08/11/2010   DAJ         Return wheter or not a Save As or Rename occurred
'' 08/12/2010   DAJ         Possible fixes for TN not shutting down correctly
'' 08/23/2010   DAJ         Added required module flag for TradeSense orders/groups
'' 09/16/2010   DAJ         Enable controls in ShowMe, Make copies of orders
'' 09/17/2010   DAJ         Make sure Dirty flag is off before showing form (#5932)
'' 09/30/2010   DAJ         Warn if group only has exit orders
'' 10/11/2010   DAJ         Save default value overrides for inputs
'' 10/11/2010   DAJ         Don't update links on order if user edited it (#5956)
'' 11/16/2010   DAJ         Added allow manual submission flag
'' 03/17/2011   DAJ         Added flatten opposite property
'' 04/21/2001   DAJ         Added vertical scroll bar to the preview control
'' 10/27/2011   DAJ         Utilized new Order Trigger Order form
'' 12/09/2011   DAJ         Added Print capability
'' 10/29/2012   DAJ         Make sure if printing to file to dump the grid ( #6745 )
'' 12/19/2012   DAJ         If user does "Save As" with an order, add copy to TSOG ( #6763 )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Number = 0
    eGDCol_Use
    eGDCol_Name
    eGDCol_Action
    eGDCol_OCO
    eGDCol_OTO
    eGDCol_OrderNumber
    eGDCol_NumCols
End Enum

Private Enum eGDInputsCols
    eGDInputsCol_Name = 0
    eGDInputsCol_OrderNum
    eGDInputsCol_Default
    eGDInputsCol_NumCols
End Enum

Private Enum eGDSaveCmd
    eGDSaveCmd_Save = 0
    eGDSaveCmd_SaveAs
    eGDSaveCmd_Rename
End Enum

Private Type mPrivate
    bDirty As Boolean                   ' Has the user made changes?
    tsoGroup As cTradeSenseOrderGroup   ' Trade Sense order group object
    
    lNextNumber As Long                 ' Next Order number
    Inputs As cTradeSenseOrderInputs    ' Collection of inputs
    bSaveAsOrRename As Boolean          ' Was a Save As or Rename performed?
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

Private Function InputsCol(ByVal nCol As eGDInputsCols) As Long
    InputsCol = nCol
End Function

Private Property Get Dirty() As Boolean
    Dirty = m.bDirty
End Property
Private Property Let Dirty(ByVal bDirty As Boolean)
    m.bDirty = bDirty
    EnableControls
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Load controls and show the form
'' Inputs:      Trade Sense Order Group, Save As or Rename done?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal tsoGroup As cTradeSenseOrderGroup, Optional bSaveAsOrRenameDone As Boolean)
On Error GoTo ErrSection:

    Set m.tsoGroup = tsoGroup
       
    InitGrid
    InitInputsGrid
    ObjectToControls
    Dirty = False

    m.bSaveAsOrRename = False
    ShowForm Me, eForm_Modal, frmMain
    bSaveAsOrRenameDone = m.bSaveAsOrRename

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmTradeSenseOrderGroup.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the journals for the selected day
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "TradeSenseOrderGroup", Me, 0
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.PrintMe"
            
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the print preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrder As cTradeSenseOrder     ' TradeSense order object
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .FontUnderline = False
        .TextAlign = taCenterMiddle
        If Len(m.tsoGroup.Name) = 0 Then
            .Text = "New TradeSense Order Group"
        Else
            .Text = "TradeSense Order Group: " & m.tsoGroup.Name
        End If
        .Font.Bold = False
        .Font.Size = 12
        .TextAlign = taLeftMiddle
        
        .Text = vbLf & vbLf
        If Len(Trim(txtDescription.Text)) > 0 Then
            .Text = "Description: " & txtDescription.Text
            .Text = vbLf & vbLf
        End If
        
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgOrders
        Else
            .RenderControl = fgOrders.hWnd
        End If
        
        .Paragraph = ""
        .Text = vbLf & vbLf
        .Text = "TradeSense Orders:" & vbLf
    
        For lIndex = fgOrders.FixedRows To fgOrders.Rows - 1
            Set tsOrder = OrderForRow(lIndex)
            If Not tsOrder Is Nothing Then
                .Text = Str(lIndex) & ": " & tsOrder.Name & vbLf
                .Text = tsOrder.ToolTip & vbLf & vbLf
            End If
        Next lIndex
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAllowManual_Click
'' Description: Allow the user to determine if the group can be manually submitted
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAllowManual_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.chkAllowManual_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkFlattenOpposite_Click
'' Description: Allow the user to determine if other groups get flattened
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkFlattenOpposite_Click()
On Error GoTo ErrSection:

    If Visible Then
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.chkFlattenOpposite_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkLinkInputs_Click
'' Description: Allow the user to toggle the linking of input values
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkLinkInputs_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterInputsGrid
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.chkLinkInputs_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddOrder_Click
'' Description: Allow the user to add an existing order to the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddOrder_Click()
On Error GoTo ErrSection:

    AddOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.cmdAddOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditOrder_Click
'' Description: Allow the user to edit an existing order in the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditOrder_Click()
On Error GoTo ErrSection:

    EditOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.cmdEditOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewOrder_Click
'' Description: Allow the user to add a new order to the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewOrder_Click()
On Error GoTo ErrSection:

    NewOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.cmdNewOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOco_Click
'' Description: Allow the user to setup an order-cancel-order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOco_Click()
On Error GoTo ErrSection:

    SetupOCO

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrderGroup.cmdOco_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOto_Click
'' Description: Allow the user to setup an order-trigger-order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOto_Click()
On Error GoTo ErrSection:

    SetupOTO

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrderGroup.cmdOto_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemoveOrder_Click
'' Description: Allow the user to remove an existing order in the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemoveOrder_Click()
On Error GoTo ErrSection:

    RemoveOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.cmdRemoveOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgInputs_AfterEdit
'' Description: If inputs are linked, change default on all inputs of same name
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Static bInProgress As Boolean       ' Are we currently changing other inputs?
    Dim lIndex As Long                  ' Index into a for loop
    Dim strInputName As String          ' Input name
    Dim strValue As String              ' Value
    Dim tsInput As cTradeSenseOrderInput ' Input object

    If (Visible = True) And (bInProgress = False) Then
        bInProgress = True
        
        With fgInputs
            strInputName = .TextMatrix(Row, InputsCol(eGDInputsCol_Name))
            strValue = .TextMatrix(Row, InputsCol(eGDInputsCol_Default))
        
            If CheckBoxValue(chkLinkInputs) = True Then
                m.Inputs.DefaultValueForName(strInputName) = strValue
                
                For lIndex = .FixedRows To .Rows - 1
                    If lIndex <> Row Then
                        If UCase(.TextMatrix(lIndex, InputsCol(eGDInputsCol_Name))) = UCase(strInputName) Then
                            .TextMatrix(lIndex, InputsCol(eGDInputsCol_Default)) = strValue
                        End If
                    End If
                Next lIndex
            ElseIf Not .RowData(Row) Is Nothing Then
                If TypeOf .RowData(Row) Is cTradeSenseOrderInput Then
                    Set tsInput = .RowData(Row)
                    m.Inputs(tsInput.Key(True)).DefaultValue = strValue
                End If
            End If
        End With
        
        Dirty = True
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.fgInputs_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgInputs_BeforeEdit
'' Description: Only allow the user to edit the Default Value column
'' Inputs:      Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim tsInput As cTradeSenseOrderInput ' Order input object

    If Col <> InputsCol(eGDInputsCol_Default) Then
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
    RaiseError "frmTradeSenseOrderGroup.fgInputs_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_AfterEdit
'' Description: Make sure all orders in a subtree get the Use flag from the parent
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim bUse As Boolean                 ' Use flag on the current line
    Static bInProgress As Boolean       ' Are we currently setting the use flag?

    If (bInProgress = False) And (Visible = True) Then
        bInProgress = True
        
        If Col = GDCol(eGDCol_Use) Then
            bUse = CheckedCell(fgOrders, Row, GDCol(eGDCol_Use))
            
            UpdateDescendants Row, bUse
            UpdateAncestors Row, bUse
            UpdateUseFlag Row, bUse
            
            Dirty = True
        End If
        
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.fgOrders_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_AfterMoveRow
'' Description: Fix order numbers and such after a row has been moved
'' Inputs:      Row, Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_AfterMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.fgOrders_AfterMoveRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_AfterRowColChange
'' Description: Enable/Disable controls as the selection changes in the grid
'' Inputs:      Old Row, Old Column, New Row, New Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    UpdatePreview
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.fgOrders_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_BeforeEdit
'' Description: Only allow the user to edit the "Use" column
'' Inputs:      Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> GDCol(eGDCol_Use) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.fgOrders_BeforeEdit"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_BeforeMouseDown
'' Description: Allow the user to bring up the popup menu with a right-click
'' Inputs:      Mouse Button, Shift/Ctrl/Alt Status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    fgOrders.Row = fgOrders.MouseRow
    
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.fgOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_DblClick
'' Description: Allow the user to edit an order with a double click on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_DblClick()
On Error GoTo ErrSection:

    EditOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.fgOrders_DblClick"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_KeyDown
'' Description: Allow the user to add and delete orders with the keyboard
'' Inputs:      Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyInsert Then
        AddOrder
    ElseIf KeyCode = vbKeyDelete Then
        RemoveOrder
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.fgOrders_KeyDown"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_KeyPress
'' Description: Allow the user to edit orders with the keyboard
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        EditOrder
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.fgOrders_KeyPress"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: The first time the form is shown, update the preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean      ' Have we already been called?
    
    If bAlreadyDone = False Then
        bAlreadyDone = True
        
        ' The first time that the form is shown, we need to update the preview again
        ' because the coloring doesn't seem to work while the form is invisible...
        UpdatePreview
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement of the form from the INI file
    Dim bGenesisUser As Boolean         ' Genesis user
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("frmTradeSenseOrderGroup", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If

    With tbToolbar
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("kPrint"))
        .Tools("ID_Exit").Picture = Picture16(ToolbarIcon("kCancel"))
    End With
    
    rtbPreview.Locked = True
    rtbPreview.BackColor = &H80000000
    
    Icon = Picture16(ToolbarIcon("kTradeSenseOrders"))
    
    bGenesisUser = FileExist("C:\Common\Files.EXE")
    fraAdvanced.Visible = ShowAdvancedTSOG Or bGenesisUser
    fraGenesis.Visible = bGenesisUser
    
    mnuPopUp.Visible = False
    
    Set m.Inputs = New cTradeSenseOrderInputs
    m.Inputs.ForGroups = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.Form_Load"
    
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
        ExitForm
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.Form_QueryUnload"
    
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

    Dim lMinScaleWidth As Long          ' Minimum allowed scale width
    Dim lMinScaleHeight As Long         ' Minimum allowed scale height
    Dim lVertSpace As Long              ' Vertical space between controls
    Dim lHorzSpace As Long              ' Horizontal space between controls
    Dim lGridTop As Long                ' Grid top
    
    lVertSpace = 120
    lHorzSpace = 120
    
    lMinScaleWidth = (fraButtons.Width * 5) + (lHorzSpace * 2)
    If fraAdvanced.Visible Then
        lMinScaleHeight = lblDescription.Height + txtDescription.Height + fraButtons.Height + rtbPreview.Height + fraAdvanced.Height + (lVertSpace * 5)
    Else
        lMinScaleHeight = lblDescription.Height + txtDescription.Height + fraButtons.Height + rtbPreview.Height + (lVertSpace * 4)
    End If

    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        With lblDescription
            .Move lHorzSpace, lVertSpace
        End With
        
        With txtDescription
            .Move lHorzSpace, lblDescription.Height + lVertSpace, ScaleWidth - (lHorzSpace * 2)
        End With
        
        With fgOrders
            lGridTop = lblDescription.Height + txtDescription.Height + (lVertSpace * 2)
            If fraAdvanced.Visible Then
                .Move lHorzSpace, lGridTop, ScaleWidth - fraButtons.Width - (lHorzSpace * 3), ScaleHeight - lGridTop - rtbPreview.Height - fraAdvanced.Height - (lVertSpace * 2)
            Else
                .Move lHorzSpace, lGridTop, ScaleWidth - fraButtons.Width - (lHorzSpace * 3), ScaleHeight - lGridTop - rtbPreview.Height - (lVertSpace * 2)
            End If
        End With
        
        With fraButtons
            .Move ScaleWidth - .Width - lHorzSpace, lGridTop
        End With
        
        If fraAdvanced.Visible Then
            With rtbPreview
                .Move lHorzSpace, fgOrders.Top + fgOrders.Height + lVertSpace, ScaleWidth - (lHorzSpace * 2)
            End With
            
            With fraAdvanced
                .Move lHorzSpace, ScaleHeight - .Height - lVertSpace, ScaleWidth - (lHorzSpace * 2)
            End With
            
            With fgInputs
                .Move lHorzSpace, .Top, fraAdvanced.Width - (lHorzSpace * 2)
            End With
        Else
            With rtbPreview
                .Move lHorzSpace, ScaleHeight - .Height - lVertSpace, ScaleWidth - (lHorzSpace * 2)
            End With
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmTradeSenseOrderGroup", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAddOrder_Click
'' Description: Allow the user to add an existing order to the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddOrder_Click()
On Error GoTo ErrSection:

    AddOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.mnuAddOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditOrder_Click
'' Description: Allow the user to edit an existing order in the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditOrder_Click()
On Error GoTo ErrSection:

    EditOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.mnuEditOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNewOrder_Click
'' Description: Allow the user to add a new order to the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNewOrder_Click()
On Error GoTo ErrSection:

    NewOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.mnuNewOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRemoveOrder_Click
'' Description: Allow the user to remove an existing order in the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRemoveOrder_Click()
On Error GoTo ErrSection:

    RemoveOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.mnuRemoveOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOco_Click
'' Description: Allow the user to setup an order-cancel-order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOco_Click()
On Error GoTo ErrSection:

    SetupOCO

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrderGroup.mnuOco_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOto_Click
'' Description: Allow the user to setup an order-trigger-order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOto_Click()
On Error GoTo ErrSection:

    SetupOTO

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrderGroup.mnuOto_Click"

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
        Case "ID_SAVE"
            Save eGDSaveCmd_Save
        Case "ID_SAVEAS"
            Save eGDSaveCmd_SaveAs
        Case "ID_RENAME"
            Save eGDSaveCmd_Rename
        Case "ID_PRINT"
            PrintMe
        Case "ID_EXIT"
            ExitForm
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.tbToolbar_ToolClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDescription_Change
'' Description: Set the dirty flag when user changes the description
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDescription_Change()
On Error GoTo ErrSection:

    Dirty = (txtDescription.Text <> m.tsoGroup.Description)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.txtDescription_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtRequiredMod_Change
'' Description: Set the dirty flag when user changes the required module
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtRequiredMod_Change()
On Error GoTo ErrSection:

    Dirty = (txtRequiredMod.Text <> m.tsoGroup.RequiredMod)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.txtRequiredMod_Change"
    
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

    With fgOrders
        .Redraw = flexRDNone
        
        SetupGrid fgOrders, eGridMode_Tree
        .OutlineBar = flexOutlineBarSimpleLeaf
        .Editable = flexEDKbdMouse
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Number)) = "#"
        .TextMatrix(0, GDCol(eGDCol_Use)) = "Use"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_Action)) = "Action"
        .TextMatrix(0, GDCol(eGDCol_OCO)) = "OCO"
        .TextMatrix(0, GDCol(eGDCol_OTO)) = "OTO"
        .TextMatrix(0, GDCol(eGDCol_OrderNumber)) = "Actual #"
        
        .ColDataType(GDCol(eGDCol_Use)) = flexDTBoolean
        
        .ColHidden(GDCol(eGDCol_OTO)) = Not IsIDE
        .ColHidden(GDCol(eGDCol_OrderNumber)) = Not IsIDE
        
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrderGroup.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ObjectToControls
'' Description: Load up the controls from the object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ObjectToControls()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrders As cGdTree             ' Collection of orders
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    Dim lMaxOrderNum As Long            ' Max order number

    With m.tsoGroup
        SetEditorCaption Me, "Trade Sense Order Template", .Name
        txtDescription.Text = .Description
        
        Set tsOrders = .Orders
        
        With fgOrders
            .Redraw = flexRDNone
            
            For lIndex = 1 To tsOrders.Count
                Set tsOrder = tsOrders(lIndex).MakeCopy
                OrderToGrid tsOrder, , False
            Next lIndex
            
            .AutoSize 0, .Cols - 1, False, 75
            .Redraw = flexRDBuffered
        End With
    
        CheckBoxValue(chkLinkInputs) = .LinkInputs
        Set m.Inputs = .Inputs
        LoadInputsGrid
        txtRequiredMod.Text = .RequiredMod
        CheckBoxValue(chkAllowManual) = .AllowManualSubmission
        CheckBoxValue(chkFlattenOpposite) = .FlattenOpposite
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.ObjectToControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ControlsToObject
'' Description: Load up the object from the controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ControlsToObject()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    m.tsoGroup.Description = Trim(txtDescription.Text)
        
    m.tsoGroup.ClearOrders
    With fgOrders
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cTradeSenseOrder Then
                m.tsoGroup.AddOrder .RowData(lIndex)
            End If
        Next lIndex
    End With
    
    m.tsoGroup.LinkInputs = CheckBoxValue(chkLinkInputs)
    m.tsoGroup.RequiredMod = Trim(txtRequiredMod.Text)
    m.tsoGroup.AllowManualSubmission = CheckBoxValue(chkAllowManual)
    m.tsoGroup.FlattenOpposite = CheckBoxValue(chkFlattenOpposite)
    
    For lIndex = 1 To m.Inputs.Count
        m.tsoGroup.Inputs(m.Inputs(lIndex).Key(True)).DefaultValue = m.Inputs(lIndex).DefaultValue
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.ControlsToObject"

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

    Dim bValidRow As Boolean            ' Is the currently selected row in the grid valid?
    Dim bTwoRows As Boolean             ' Is there at least two orders in the grid?

    ' Don't enable the Save button unless the user has made changes
    ' or this is a new order group...
    tbToolbar.Tools("ID_Save").Enabled = Dirty
    
    ' Only enable the Edit and Remove Order buttons if there is an order
    ' selected in the grid...
    bValidRow = ValidRowSelected
    Enable cmdEditOrder, bValidRow
    Enable cmdRemoveOrder, bValidRow
    
    ' Only enable the OTO and OCO buttons if there is an order selected in
    ' the grid and there are at least two orders in the grid...
    bTwoRows = (fgOrders.Rows >= fgOrders.FixedRows + 2)
    Enable cmdOto, bValidRow And bTwoRows
    Enable cmdOco, bValidRow And bTwoRows
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRowSelected
'' Description: Is the currently selected row in the grid valid?
'' Inputs:      None
'' Returns:     True if Valid Row selected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidRowSelected() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    With fgOrders
        If .Rows > .FixedRows Then
            bReturn = (.RowSel >= .FixedRows) And (.RowSel < .Rows)
        End If
    End With
    
    ValidRowSelected = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.ValidRowSelected"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedOrder
'' Description: Get the currently selected order in the grid
'' Inputs:      None
'' Returns:     Selected Order in the Grid (Nothing if not valid row)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedOrder() As cTradeSenseOrder
On Error GoTo ErrSection:

    Set SelectedOrder = OrderForRow(fgOrders.Row)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.SelectedOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToGrid
'' Description: Fill in the row in the grid with the order
'' Inputs:      Order, Row, Auto Size Grid?, Update Links?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderToGrid(tsOrder As cTradeSenseOrder, Optional ByVal lRow As Long = -1&, Optional ByVal bAutoSize As Boolean = True, Optional ByVal bUpdateLinks As Boolean = True)
On Error GoTo ErrSection:

    With fgOrders
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        If tsOrder.OrderNumber = 0 Then
            tsOrder.OrderNumber = lRow
        End If
        .RowData(lRow) = tsOrder
        
        .IsSubtotal(lRow) = True
        .TextMatrix(lRow, GDCol(eGDCol_Number)) = Str(lRow)
        CheckedCell(fgOrders, lRow, GDCol(eGDCol_Use)) = tsOrder.Use
        .TextMatrix(lRow, GDCol(eGDCol_Name)) = tsOrder.Name
        .TextMatrix(lRow, GDCol(eGDCol_Action)) = tsOrder.Action
        
        If bUpdateLinks = True Then
            If tsOrder.OCOs.Size = 0 Then
                .TextMatrix(lRow, GDCol(eGDCol_OCO)) = ""
            Else
                .TextMatrix(lRow, GDCol(eGDCol_OCO)) = tsOrder.OCOs.JoinFields(",")
            End If
            If tsOrder.OTO = 0 Then
                .TextMatrix(lRow, GDCol(eGDCol_OTO)) = ""
            Else
                .TextMatrix(lRow, GDCol(eGDCol_OTO)) = Str(tsOrder.OTO)
            End If
            .TextMatrix(lRow, GDCol(eGDCol_OrderNumber)) = Str(tsOrder.OrderNumber)
        
            If tsOrder.OTO = 0 Then
                .RowOutlineLevel(lRow) = 0
            Else
                .RowOutlineLevel(lRow) = .RowOutlineLevel(Abs(tsOrder.OTO)) + 1
            End If
        End If
        
        ' If the current row is the one selected, temporarily remove selection and
        ' put it back to allow the RowColChange event to take over...
        If .Row = lRow Then
            .Row = -1&
            .Row = lRow
        End If
        
        If bAutoSize Then
            .AutoSize 0, .Cols - 1, False, 75
        End If
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.OrderToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewOrder
'' Description: Allow the user to add a new order to the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewOrder()
On Error GoTo ErrSection:

    Dim tsOrder As New cTradeSenseOrder ' Trade Sense order for adding
    
    If frmTradeSenseOrder.ShowMe(tsOrder, , , , False) = True Then
        OrderToGrid tsOrder
        AddInputsForOrder tsOrder
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.NewOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrder
'' Description: Allow the user to add an existing order to the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddOrder()
On Error GoTo ErrSection:

    Dim tsOrder As New cTradeSenseOrder ' Trade Sense order for adding
    
    If frmTradeSenseOrders.ShowMe(tsOrder) = True Then
        OrderToGrid tsOrder
        AddInputsForOrder tsOrder
        Dirty = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.AddOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditOrder
'' Description: Allow the user to edit an existing order in the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditOrder()
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' TradeSense order for editing
    Dim strOldId As String              ' Old TradeSense order ID
    
    Set tsOrder = SelectedOrder.MakeCopy
    If Not tsOrder Is Nothing Then
        strOldId = tsOrder.ID
        
        If frmTradeSenseOrder.ShowMe(tsOrder, , , , False) = True Then
            If tsOrder.ID = strOldId Then
                UpdateOrder tsOrder, False
                UpdateInputsForOrder tsOrder
            Else
                tsOrder.OrderNumber = 0
                OrderToGrid tsOrder
                AddInputsForOrder tsOrder
                Dirty = True
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.EditOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveOrder
'' Description: Allow the user to remove an existing order in the group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveOrder(Optional ByVal bAskUser As Boolean = True)
On Error GoTo ErrSection:

    Dim strResponse As String           ' User resonse to the question
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order for editing
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Row where the order was
    Dim bRemoved As Boolean             ' Were any inputs removed?
    Dim bChanged As Boolean             ' Were any inputs changed?
    
    Set tsOrder = SelectedOrder
    If Not tsOrder Is Nothing Then
        If bAskUser Then
            strResponse = InfBox("Are you sure that you want to remove '" & tsOrder.Name & "'?", "?", "+Yes|-No", "Order Remove Confirmation")
        Else
            strResponse = "Y"
        End If
        
        If strResponse = "Y" Then
            bRemoved = False
            bChanged = False
            
            With fgOrders
                lRow = .Row
                
                For lIndex = .FixedRows To .Rows - 1
                    Set tsOrder = OrderForRow(lIndex)
                    If Not tsOrder Is Nothing Then
                        If tsOrder.RemoveOCO(lRow) Then
                            OrderToGrid tsOrder, lIndex
                        ElseIf tsOrder.OTO = lRow Then
                            tsOrder.OTO = 0
                            OrderToGrid tsOrder, lIndex
                        End If
                    End If
                Next lIndex
                
                .RemoveItem lRow
                If RemoveInputsForOrder(lRow, False) Then
                    bRemoved = True
                End If
                
                ResetOrderNumbers
                ResetTree
            End With
            
            If bRemoved Or bChanged Then
                LoadInputsGrid
            End If
            
            Dirty = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.RemoveOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetupOTO
'' Description: Allow the user to setup an order-trigger-order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupOTO()
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    Dim tsOto As cTradeSenseOrder       ' Trade Sense order to trigger off of
    Dim lLastChild As Long              ' Last child of the OTO order
    Dim lDestination As Long            ' Destination for the row
    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrders As cTradeSenseOrders   ' Collection of TradeSense orders
    Dim bChanged As Boolean             ' Has any OTO information changed?
    Dim ordersTree As cGdTree           ' Tree of orders
    Dim lRow As Long                    ' Index into a for loop
    
    Set tsOrder = SelectedOrder
    If Not tsOrder Is Nothing Then
        Set tsOrders = OrdersForOTO(tsOrder)
        If frmOrderTriggerOrder.ShowMe(tsOrder, tsOrders) Then
            bChanged = False
            
            With fgOrders
                ' 1) Reassign OTO numbers...
                For lIndex = .FixedRows To .Rows - 1
                    Set tsOrder = .RowData(lIndex)
                    If tsOrders.Exists(Str(tsOrder.OrderNumber)) Then
                        If tsOrder.OTO <> tsOrders(Str(tsOrder.OrderNumber)).OTO Then
                            tsOrder.OTO = tsOrders(Str(tsOrder.OrderNumber)).OTO
                            .RowData(lIndex) = tsOrder
                            bChanged = True
                        End If
                    End If
                Next lIndex
                
                If bChanged Then
                    ' 2) Rebuild tree...
                    Set ordersTree = New cGdTree
                    ordersTree.Add Nothing, "0"
                    For lIndex = .FixedRows To .Rows - 1
                        Set tsOrder = .RowData(lIndex)
                        If ordersTree.Exists(Str(tsOrder.OTO)) Then
                            ordersTree.Add tsOrder, Str(tsOrder.OrderNumber), Str(tsOrder.OTO), eTREE_LastChild
                        End If
                    Next lIndex
                    For lIndex = .FixedRows To .Rows - 1
                        Set tsOrder = .RowData(lIndex)
                        If (ordersTree.Exists(Str(tsOrder.OTO)) = True) And (ordersTree.Exists(Str(tsOrder.OrderNumber)) = False) Then
                            ordersTree.Add tsOrder, Str(tsOrder.OrderNumber), Str(tsOrder.OTO), eTREE_LastChild
                        End If
                    Next lIndex
                    
                    ' 3) Rebuild grid
                    .Redraw = flexRDNone
                    .Rows = .FixedRows
                    For lIndex = 2 To ordersTree.Count
                        OrderToGrid ordersTree(lIndex), , False, False
                    Next lIndex
                
                    ResetOrderNumbers
                    ResetTree
                    
                    .AutoSize 0, .Cols - 1, False, 75
                    .Redraw = flexRDBuffered
                    
                    Dirty = True
                End If
            End With
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.SetupOTO"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetupOCO
'' Description: Allow the user to setup an order-cancel-order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupOCO()
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    Dim tsOcos As cTradeSenseOrders     ' Trade Sense orders to oco with
    Dim tsOco2 As cTradeSenseOrder      ' Trade Sense order that oco is currently oco'd with
    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrders As cTradeSenseOrders   ' Collection of TradeSense orders
    
    Set tsOrder = SelectedOrder
    If Not tsOrder Is Nothing Then
        Set tsOrders = Orders(tsOrder)
        If frmTradeSenseOrders.ShowMeOCO(tsOcos, tsOrders, tsOrder.OCOs.Size > 0, True) = True Then
            If tsOcos Is Nothing Then
                For lIndex = 0 To tsOrder.OCOs.Size - 1
                    Set tsOco2 = OrderForRow(tsOrder.OCOs(lIndex))
                    tsOco2.RemoveOCO tsOrder.OrderNumber
                    OrderToGrid tsOco2, tsOco2.OrderNumber
                Next lIndex
                                
                tsOrder.OCOs.Clear
                OrderToGrid tsOrder, tsOrder.OrderNumber
            Else
                For lIndex = 1 To tsOcos.Count
                    If tsOcos(lIndex).IsOCO(tsOrder.OrderNumber) = False Then
                        tsOcos(lIndex).OCOs.Add tsOrder.OrderNumber
                        OrderToGrid tsOcos(lIndex), tsOcos(lIndex).OrderNumber
                    End If
                    
                    If tsOrder.IsOCO(tsOcos(lIndex).OrderNumber) = False Then
                        tsOrder.OCOs.Add tsOcos(lIndex).OrderNumber
                    End If
                Next lIndex
                
                OrderToGrid tsOrder, tsOrder.OrderNumber
            End If
            
            Dirty = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.SetupOCO"
    
End Sub

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
    Dim Grps As cTradeSenseOrderGroups  ' Existing TradeSense order groups
    Dim strID As String                 ' ID for the chosen name

    If Len(m.tsoGroup.Name) = 0 Then
        strText = "Save the current Trade Sense Order Group as..."
        strHeader = "Save"
        strNewName = AskBox("h=" & strHeader & " ; i=? ; g=string ; d=" & m.tsoGroup.Name & " ; " & strText)
    ElseIf nSaveCmd = eGDSaveCmd_SaveAs Then
        strText = "Save a copy of the current Trade Sense Order Group as..."
        strHeader = "Save As"
        strNewName = AskBox("h=" & strHeader & " ; i=? ; g=string ; d=" & "Copy of " & m.tsoGroup.Name & " ; " & strText)
    ElseIf nSaveCmd = eGDSaveCmd_Rename Then
        strText = "Rename the current Trade Sense Order Group as..."
        strHeader = "Rename"
        strNewName = AskBox("h=" & strHeader & " ; i=? ; g=string ; d=" & m.tsoGroup.Name & " ; " & strText)
    Else
        strNewName = m.tsoGroup.Name
    End If
    
    Set Grps = New cTradeSenseOrderGroups
    Grps.Load
    
    Do While (Len(strNewName) > 0) And (strNewName <> m.tsoGroup.Name)
        strID = Grps.IdForName(strNewName)
        
        If (Len(strID) > 0) And (strID <> m.tsoGroup.ID) Then
            InfBox "'" & strNewName & "' already exists.  Please select a new name", "!", , "Save Error"
        ElseIf IsValidFileBase(strNewName, False) = False Then
            InfBox "'" & strNewName & "' is not a valid name.  Please select a new name", "!", , "Save Error"
        Else
            Exit Do
        End If
        
        strNewName = AskBox("h=" & strHeader & " ; i=? ; g=string ; d=" & m.tsoGroup.Name & " ; " & strText)
    Loop
    
    If Len(strNewName) > 0 Then
        ControlsToObject
        If strNewName <> m.tsoGroup.Name Then
            If nSaveCmd = eGDSaveCmd_SaveAs Then
                m.tsoGroup.ClearID
            End If
            m.tsoGroup.Name = strNewName
            SetEditorCaption Me, "Trade Sense Order Template", m.tsoGroup.Name
            
            m.bSaveAsOrRename = True
        End If
        m.tsoGroup.ToFile
        
        If m.tsoGroup.OnlyContainsExits Then
            InfBox "This TradeSense order group only contains exits.  If you submit this group, nothing will happen.", "i", , "Warning"
        End If
        
        Dirty = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.Save"
    
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
        strReturn = InfBox("Do you want to save your changes?||Clicking No will undo any changes you have made to this Trade Sense order group.|", "?", "+Yes|No|-Cancel", Caption)
        Select Case strReturn
            Case "C"
                bHide = False
                
            Case "Y"
                bHide = True
                Save eGDSaveCmd_Save
                
            Case "N"
                bHide = True
                
        End Select
    End If
    
    If bHide Then
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.ExitForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Orders
'' Description: Build a collection from the orders in the grid
'' Inputs:      None
'' Returns:     Orders Collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Orders(Optional ByVal Order As cTradeSenseOrder = Nothing) As cTradeSenseOrders
On Error GoTo ErrSection:

    Dim tsOrders As cTradeSenseOrders   ' Collection of Trade Sense orders
    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    
    Set tsOrders = New cTradeSenseOrders
    With fgOrders
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cTradeSenseOrder Then
                Set tsOrder = .RowData(lIndex)
                If Order Is Nothing Then
                    tsOrders.Add tsOrder, Str(tsOrder.OrderNumber)
                ElseIf tsOrder.OrderNumber <> Order.OrderNumber Then
                    'If (tsOrder.OrderNumber <> Order.OTO) And (tsOrder.OTO <> Order.OrderNumber) Then
                    If (IsDescendant(tsOrder, Order) = False) And (IsAncestor(tsOrder, Order) = False) Then
                        If Not Order.IsOCO(tsOrder.OrderNumber) Then
                            tsOrders.Add tsOrder, Str(tsOrder.OrderNumber)
                        End If
                    End If
                End If
            End If
        Next lIndex
    End With
    
    Set Orders = tsOrders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.Orders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrdersForOTO
'' Description: Build a collection from the orders than can be OTO for given order
'' Inputs:      Order
'' Returns:     Orders Collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrdersForOTO(ByVal Order As cTradeSenseOrder) As cTradeSenseOrders
On Error GoTo ErrSection:

    Dim tsOrders As cTradeSenseOrders   ' Collection of Trade Sense orders
    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    
    Set tsOrders = New cTradeSenseOrders
    With fgOrders
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cTradeSenseOrder Then
                Set tsOrder = .RowData(lIndex)
                If tsOrder.OrderNumber <> Order.OrderNumber Then
                    If (tsOrder.OTO = 0) Or (tsOrder.OTO = Order.OrderNumber) Then
                        If IsAncestor(tsOrder, Order) = False Then
                            If Not Order.IsOCO(tsOrder.OrderNumber) Then
                                tsOrders.Add tsOrder.MakeCopy, Str(tsOrder.OrderNumber)
                            End If
                        End If
                    End If
                End If
            End If
        Next lIndex
    End With
    
    Set OrdersForOTO = tsOrders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.OrdersForOTO"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeOrderNumber
'' Description: Change an order number
'' Inputs:      Old Number, New Number, Start Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeOrderNumber(ByVal lOldNumber As Long, ByVal lNewNumber As Long, Optional ByVal lStartIndex As Long = -1&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object

    With fgOrders
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cTradeSenseOrder Then
                Set tsOrder = OrderForRow(lIndex)
                If Not tsOrder Is Nothing Then
                    If (tsOrder.OrderNumber = lOldNumber) And (lIndex >= lStartIndex) Then
                        tsOrder.OrderNumber = lNewNumber
                        OrderToGrid tsOrder, lIndex
                    ElseIf tsOrder.ChangeOCO(lOldNumber, lNewNumber * -1&) Then
                        OrderToGrid tsOrder, lIndex
                    ElseIf tsOrder.OTO = lOldNumber Then
                        tsOrder.OTO = lNewNumber * -1&
                        OrderToGrid tsOrder, lIndex
                    ElseIf (tsOrder.OrderNumber <> lIndex) And (lStartIndex = -1&) Then
                        tsOrder.OrderNumber = lIndex
                        OrderToGrid tsOrder, lIndex
                    End If
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.ChangeOrderNumber"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForRow
'' Description: Get the order for the given row
'' Inputs:      Row
'' Returns:     Order (Nothing if not valid)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderForRow(ByVal lRow As Long) As cTradeSenseOrder
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' Order to return from the function
    
    Set tsOrder = Nothing
    If (lRow >= fgOrders.FixedRows) And (lRow < fgOrders.Rows) Then
        If TypeOf fgOrders.RowData(lRow) Is cTradeSenseOrder Then
            Set tsOrder = fgOrders.RowData(lRow)
        End If
    End If
    
    Set OrderForRow = tsOrder

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.OrderForRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdatePreview
'' Description: Update the preview if there is an order selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdatePreview()
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' Currently selected order in the grid
    
    Set tsOrder = SelectedOrder
    If Not tsOrder Is Nothing Then
        rtbPreview.TextRTF = tsOrder.PreviewRTF
    Else
        rtbPreview.Text = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.UpdatePreview"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateOrder
'' Description: Update the order in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateOrder(ByVal tsOrder As cTradeSenseOrder, Optional ByVal bUpdateLinks As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim GridOrder As cTradeSenseOrder   ' Order from the grid
    
    With fgOrders
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            Set GridOrder = OrderForRow(lIndex)
            If Not GridOrder Is Nothing Then
                If GridOrder.ID = tsOrder.ID Then
                    OrderToGrid tsOrder, lIndex, False, bUpdateLinks
                End If
            End If
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.UpdateOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetOrderNumbers
'' Description: Reset the order numbers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetOrderNumbers()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim GridOrder As cTradeSenseOrder   ' Order from the grid
    Dim bChanged As Boolean             ' Did any inputs change?
    Dim lPrevNumber As Long             ' Previous order number
    Dim Inputs As cTradeSenseOrderInputs ' Order Inputs
    Dim lIndex2 As Long                 ' Index into a for loop

    bChanged = False
    With fgOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        Set Inputs = m.Inputs.MakeCopy
        m.Inputs.Clear
        
        For lIndex = .FixedRows To .Rows - 1
            Set GridOrder = OrderForRow(lIndex)
            If Not GridOrder Is Nothing Then
                lPrevNumber = GridOrder.OrderNumber
                If lIndex <> GridOrder.OrderNumber Then
                    ChangeOrderNumber GridOrder.OrderNumber, lIndex, lIndex
                End If
                    
                For lIndex2 = Inputs.Count To 1 Step -1
                    If Inputs(lIndex2).OrderNumber = lPrevNumber Then
                        Inputs(lIndex2).OrderNumber = lIndex
                        m.Inputs.Add Inputs(lIndex2)
                        Inputs.Remove lIndex2
                        bChanged = True
                    End If
                Next lIndex2
            End If
        Next lIndex
        
        For lIndex = .FixedRows To .Rows - 1
            Set GridOrder = OrderForRow(lIndex)
            If Not GridOrder Is Nothing Then
                If GridOrder.OTO < 0 Then
                    GridOrder.OTO = GridOrder.OTO * -1&
                End If
                GridOrder.FixNegativeOCO
                
                OrderToGrid GridOrder, lIndex, False
            End If
        Next lIndex
        
        If bChanged Then
            LoadInputsGrid
        End If
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrderGroup.ResetOrderNumbers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetTree
'' Description: Reset the tree levels
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetTree()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim GridOrder As cTradeSenseOrder   ' Order from the grid

    With fgOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            Set GridOrder = OrderForRow(lIndex)
            If Not GridOrder Is Nothing Then
                If GridOrder.OTO = 0 Then
                    .RowOutlineLevel(lIndex) = 0
                Else
                    .RowOutlineLevel(lIndex) = .RowOutlineLevel(GridOrder.OTO) + 1
                End If
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.ResetTree"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsDescendant
'' Description: Determine if the given descendant is a descendant of the ancestor
'' Inputs:      Descendant, Ancestor
'' Returns:     True if descendant, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsDescendant(ByVal tsDescendant As cTradeSenseOrder, ByVal tsAncestor As cTradeSenseOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim GridOrder As cTradeSenseOrder   ' Order from the grid
    
    bReturn = False
    Set GridOrder = tsDescendant
    Do
        If GridOrder.OTO = 0 Then
            Exit Do
        ElseIf GridOrder.OTO = tsAncestor.OrderNumber Then
            bReturn = True
            Exit Do
        Else
            Set GridOrder = OrderForRow(GridOrder.OTO)
        End If
    Loop While True
    
    IsDescendant = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.IsDescendant"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsAncestor
'' Description: Determine if the given ancestor is an ancestor of the descendant
'' Inputs:      Ancestor, Descendant
'' Returns:     True if ancestor, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsAncestor(ByVal tsAncestor As cTradeSenseOrder, ByVal tsDescendant As cTradeSenseOrder) As Boolean
On Error GoTo ErrSection:

    IsAncestor = IsDescendant(tsDescendant, tsAncestor)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.IsAncestor"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateUseFlag
'' Description: Update the use flag on the given row
'' Inputs:      Row, Use
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateUseFlag(ByVal lRow As Long, ByVal bUse As Boolean)
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' TradeSense order object
    
    Set tsOrder = OrderForRow(lRow)
    If Not tsOrder Is Nothing Then
        tsOrder.Use = bUse
        fgOrders.RowData(lRow) = tsOrder
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.UpdateUseFlag"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UdpateDescendants
'' Description: Set the use flag on the descendants for the given parent row
'' Inputs:      Parent Row, Use
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateDescendants(ByVal lParentRow As Long, ByVal bUse As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lNextSibling As Long            ' Next sibling for this order
    
    With fgOrders
        lNextSibling = .GetNodeRow(lParentRow, flexNTNextSibling)
        If (lNextSibling = -1&) And (.RowOutlineLevel(lParentRow) = 0) Then
            lNextSibling = .Rows
        End If
        
        If lNextSibling > lParentRow + 1 Then
            For lIndex = lParentRow + 1 To lNextSibling - 1
                CheckedCell(fgOrders, lIndex, GDCol(eGDCol_Use)) = bUse
                UpdateUseFlag lIndex, bUse
            Next lIndex
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.UdpateDescendants"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateAncestors
'' Description: Set the use flag on the ancestors for the given child row
'' Inputs:      Child Row, Use
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateAncestors(ByVal lChildRow As Long, ByVal bUse As Boolean)
On Error GoTo ErrSection:

    Dim lParentRow As Long              ' Parent row
    
    If bUse = True Then
        With fgOrders
            lParentRow = .GetNodeRow(lChildRow, flexNTParent)
            Do While lParentRow <> -1&
                CheckedCell(fgOrders, lParentRow, GDCol(eGDCol_Use)) = bUse
                UpdateUseFlag lParentRow, bUse
                lParentRow = .GetNodeRow(lParentRow, flexNTParent)
            Loop
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.UpdateAncestors"
    
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
        .ExtendLastCol = False
        
        .Cols = InputsCol(eGDInputsCol_NumCols)
        .FixedCols = 0
        .Rows = 1
        .FixedRows = 1
        
        .TextMatrix(0, InputsCol(eGDInputsCol_Name)) = "Input Name"
        .TextMatrix(0, InputsCol(eGDInputsCol_OrderNum)) = "Order#"
        .TextMatrix(0, InputsCol(eGDInputsCol_Default)) = "Default Value"
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.InitInputsGrid"
    
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
    Dim strPrevName As String           ' Previous selection before reloading grid
    Dim strPrevOrdNum As String         ' Previous selection before reloading grid

    With fgInputs
        .Redraw = flexRDNone
        
        strPrevName = ""
        strPrevOrdNum = ""
        If (.Row >= .FixedRows) And (.Row < .Rows) Then
            If CheckBoxValue(chkLinkInputs) = True Then
                strPrevName = .TextMatrix(.Row, InputsCol(eGDInputsCol_Name))
            Else
                strPrevName = .TextMatrix(.Row, InputsCol(eGDInputsCol_Name))
                strPrevOrdNum = .TextMatrix(.Row, InputsCol(eGDInputsCol_OrderNum))
            End If
        End If
        
        .Rows = .FixedRows
        For lIndex = 1 To m.Inputs.Count
            .Rows = .Rows + 1
            
            .RowData(.Rows - 1) = m.Inputs(lIndex)
            .TextMatrix(.Rows - 1, InputsCol(eGDInputsCol_Name)) = m.Inputs(lIndex).Name
            .TextMatrix(.Rows - 1, InputsCol(eGDInputsCol_OrderNum)) = m.Inputs(lIndex).OrderNumber
            .TextMatrix(.Rows - 1, InputsCol(eGDInputsCol_Default)) = m.Inputs(lIndex).DefaultValue
        Next lIndex
        
        FilterInputsGrid
        SelectInput strPrevName, strPrevOrdNum
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.LoadInputsGrid"
    
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
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim strLastName As String           ' Last input name from the grid

    With fgInputs
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Sort by order number, then input name...
        .Col = InputsCol(eGDInputsCol_OrderNum)
        .Sort = flexSortGenericAscending
        .Col = InputsCol(eGDInputsCol_Name)
        .Sort = flexSortGenericAscending
        
        For lIndex = .FixedRows To .Rows - 1
            If CheckBoxValue(chkLinkInputs) = False Then
                .RowHidden(lIndex) = False
            Else
                .RowHidden(lIndex) = (UCase(.TextMatrix(lIndex, InputsCol(eGDInputsCol_Name))) = UCase(strLastName))
                strLastName = .TextMatrix(lIndex, InputsCol(eGDInputsCol_Name))
            End If
        Next lIndex
        
        .ColHidden(InputsCol(eGDInputsCol_OrderNum)) = CheckBoxValue(chkLinkInputs)

        SetBackColors fgInputs
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.FilterInputsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectInput
'' Description: Select the given input in the grid
'' Inputs:      Name, Order Number
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SelectInput(ByVal strName As String, ByVal strOrderNumber As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    If Len(strName) > 0 Then
        With fgInputs
            For lIndex = .FixedRows To .Rows - 1
                If (CheckBoxValue(chkLinkInputs) = True) Or (Len(strOrderNumber) = 0) Then
                    If UCase(.TextMatrix(lIndex, InputsCol(eGDInputsCol_Name))) = UCase(strName) Then
                        .Row = lIndex
                        Exit For
                    End If
                Else
                    If UCase(.TextMatrix(lIndex, InputsCol(eGDInputsCol_Name))) = UCase(strName) Then
                        If UCase(.TextMatrix(lIndex, InputsCol(eGDInputsCol_OrderNum))) = UCase(strOrderNumber) Then
                            .Row = lIndex
                            Exit For
                        End If
                    End If
                End If
            Next lIndex
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.SelectInput"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddInputsForOrder
'' Description: Add inputs for the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddInputsForOrder(tsOrder As cTradeSenseOrder)
On Error GoTo ErrSection:

    Dim Inputs As cTradeSenseOrderInputs ' Inputs collection from the order
    Dim tsInput As cTradeSenseOrderInput ' TradeSense order input object
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the array
    Dim strValue As String              ' Value for the inputs
    Dim bAdded As Boolean               ' Did we add new inputs?

    Set Inputs = tsOrder.Inputs
    If Inputs.Count > 0 Then
        bAdded = False
        
        For lIndex = 1 To Inputs.Count
            Set tsInput = Inputs(lIndex).MakeCopy
            
            tsInput.OrderNumber = tsOrder.OrderNumber
            
            If m.Inputs.Exists(tsInput.Key(True)) = False Then
                strValue = m.Inputs.DefaultValueForName(tsInput.Name)
                If Len(strValue) > 0 Then
                    tsInput.DefaultValue = strValue
                End If
                
                m.Inputs.Add tsInput
                bAdded = True
            End If
        Next lIndex
        
        If bAdded = True Then
            LoadInputsGrid
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.AddInputsForOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveInputsForOrder
'' Description: Remove inputs for the given order number
'' Inputs:      Order Number, Reload Grid?
'' Returns:     True if Inputs Removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemoveInputsForOrder(ByVal lOrderNumber As Long, Optional ByVal bReloadGrid As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Were any inputs removed?
    
    bRemoved = False
    For lIndex = m.Inputs.Count To 1 Step -1
        If m.Inputs(lIndex).OrderNumber = lOrderNumber Then
            m.Inputs.Remove lIndex
            bRemoved = True
        End If
    Next lIndex
    
    If (bRemoved = True) And (bReloadGrid = True) Then
        LoadInputsGrid
    End If
    
    RemoveInputsForOrder = bRemoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.RemoveInputsForOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateInputsForOrder
'' Description: Add inputs for the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateInputsForOrder(tsOrder As cTradeSenseOrder)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bChanged As Boolean             ' Were any inputs changed?
    Dim tsInput As cTradeSenseOrderInput ' Order input object
    
    bChanged = False
    For lIndex = m.Inputs.Count To 1 Step -1
        If m.Inputs(lIndex).OrderNumber = tsOrder.OrderNumber Then
            If tsOrder.Inputs.Exists(m.Inputs(lIndex).Key(False)) = False Then
                m.Inputs.Remove lIndex
                bChanged = True
            End If
        End If
    Next lIndex
    
    For lIndex = 1 To tsOrder.Inputs.Count
        Set tsInput = tsOrder.Inputs(lIndex).MakeCopy
        
        tsInput.OrderNumber = tsOrder.OrderNumber
        If m.Inputs.Exists(tsInput.Key(True)) = False Then
            m.Inputs.Add tsInput
            bChanged = True
        End If
    Next lIndex
    
    If bChanged Then
        Dirty = True
        LoadInputsGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.UpdateInputsForOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeOrderNumberForInputs
'' Description: Change order number on inputs for the given order number
'' Inputs:      Old Order Number, New Order Number, Reload Grid?
'' Returns:     True if Inputs Changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChangeOrderNumberForInputs(ByVal lOldOrderNumber As Long, ByVal lNewOrderNumber As Long, Optional ByVal bReloadGrid As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bChanged As Boolean             ' Were any inputs changed?
    
    bChanged = False
    For lIndex = 1 To m.Inputs.Count
        If m.Inputs(lIndex).OrderNumber = lOldOrderNumber Then
            m.Inputs(lIndex).OrderNumber = lNewOrderNumber
            m.Inputs.UpdateKey lIndex
            bChanged = True
        End If
    Next lIndex
    
    If bChanged = True Then
        If bReloadGrid = True Then
            LoadInputsGrid
        End If
    End If
    
    ChangeOrderNumberForInputs = bChanged

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrderGroup.ChangeOrderNumberForInputs"
    
End Function

