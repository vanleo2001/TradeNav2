VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmStrategyBasketItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraData 
      Height          =   1875
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   3255
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
      Caption         =   "frmStrategyBasketItem.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStrategyBasketItem.frx":0028
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyBasketItem.frx":0048
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboBoxXP cboPeriod 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Top             =   270
         Width           =   1575
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
         Tip             =   "frmStrategyBasketItem.frx":0064
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
         MouseIcon       =   "frmStrategyBasketItem.frx":0084
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraToDate 
         Height          =   675
         Left            =   240
         TabIndex        =   13
         Top             =   1140
         Width           =   2955
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
         Caption         =   "frmStrategyBasketItem.frx":00A0
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmStrategyBasketItem.frx":00CC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasketItem.frx":00EC
         RightToLeft     =   0   'False
         Begin gdOCX.gdSelectDate gdToDate 
            Height          =   315
            Left            =   600
            TabIndex        =   16
            Top             =   0
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
         End
         Begin HexUniControls.ctlUniRadioXP optToDate 
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   30
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
            Caption         =   "frmStrategyBasketItem.frx":0108
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmStrategyBasketItem.frx":013A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmStrategyBasketItem.frx":015A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optToEnd 
            Height          =   255
            Left            =   360
            TabIndex        =   17
            Top             =   360
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
            Caption         =   "frmStrategyBasketItem.frx":0176
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmStrategyBasketItem.frx":01B2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmStrategyBasketItem.frx":01D2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblTo 
            Height          =   255
            Left            =   0
            Top             =   60
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
            Caption         =   "frmStrategyBasketItem.frx":01EE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmStrategyBasketItem.frx":0214
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmStrategyBasketItem.frx":0234
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin gdOCX.gdSelectDate gdFromDate 
         Height          =   315
         Left            =   840
         TabIndex        =   12
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
      End
      Begin HexUniControls.ctlUniLabelXP lblStepThree 
         Height          =   255
         Left            =   240
         Top             =   300
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
         Caption         =   "frmStrategyBasketItem.frx":0250
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmStrategyBasketItem.frx":027E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasketItem.frx":029E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFrom 
         Height          =   255
         Left            =   240
         Top             =   750
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
         Caption         =   "frmStrategyBasketItem.frx":02BA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmStrategyBasketItem.frx":02E4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasketItem.frx":0304
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraSymbol 
      Height          =   1395
      Left            =   120
      TabIndex        =   2
      Top             =   600
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
      Caption         =   "frmStrategyBasketItem.frx":0320
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStrategyBasketItem.frx":036C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyBasketItem.frx":038C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optSymbolGroup 
         Height          =   225
         Left            =   180
         TabIndex        =   3
         Top             =   360
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
         Caption         =   "frmStrategyBasketItem.frx":03A8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmStrategyBasketItem.frx":03E2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasketItem.frx":0402
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optSymbol 
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   840
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
         Caption         =   "frmStrategyBasketItem.frx":041E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmStrategyBasketItem.frx":044C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasketItem.frx":046C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   -1  'True
         Text            =   "frmStrategyBasketItem.frx":0488
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
         Tip             =   "frmStrategyBasketItem.frx":04B2
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasketItem.frx":04D2
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
         Height          =   315
         Left            =   3540
         TabIndex        =   7
         Top             =   840
         Width           =   915
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
         Caption         =   "frmStrategyBasketItem.frx":04EE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStrategyBasketItem.frx":051C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasketItem.frx":053C
         RightToLeft     =   0   'False
      End
      Begin MSComctlLib.ImageCombo cboFilters 
         Height          =   330
         Left            =   1620
         TabIndex        =   4
         Top             =   360
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "ImageCombo1"
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgInputs 
      Height          =   1155
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   8775
      _cx             =   15478
      _cy             =   2037
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
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   435
      Left            =   3188
      TabIndex        =   0
      Top             =   3480
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
      Caption         =   "frmStrategyBasketItem.frx":0558
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStrategyBasketItem.frx":0584
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyBasketItem.frx":05A4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   435
         Left            =   1440
         TabIndex        =   9
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
         Caption         =   "frmStrategyBasketItem.frx":05C0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStrategyBasketItem.frx":05EE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasketItem.frx":060E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   435
         Left            =   0
         TabIndex        =   11
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
         Caption         =   "frmStrategyBasketItem.frx":062A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStrategyBasketItem.frx":0650
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasketItem.frx":0670
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkSplit 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3540
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
      Caption         =   "frmStrategyBasketItem.frx":068C
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmStrategyBasketItem.frx":0712
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmStrategyBasketItem.frx":0732
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboSystems 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   150
      Width           =   4635
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
      Tip             =   "frmStrategyBasketItem.frx":074E
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmStrategyBasketItem.frx":076E
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblStepOne 
      Height          =   255
      Left            =   120
      Top             =   180
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
      Caption         =   "frmStrategyBasketItem.frx":078A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmStrategyBasketItem.frx":07BC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyBasketItem.frx":07DC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmStrategyBasketItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmStrategyBasketItem.frm
'' Description: Allow the user to create a strategy basket item
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 11/15/2011   DAJ         Renamed the Strategy Basket stuff
'' 02/16/2012   DAJ         Moved some code to objects
'' 04/03/2013   DAJ         Move Strategy Baskets into the database
'' 05/01/2013   DAJ         Shadow Trading
'' 08/02/2013   DAJ         Force the 'ToDate' to be after the 'FromDate'
'' 05/05/2014   DAJ         Allow FractZen bars for strategy baskets
'' 05/05/2014   DAJ         Default to 'Daily' if it was FractZen, but they aren't enabled
'' 08/19/2014   DAJ         Expose Strategy Basket Item Inputs
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kErrorCaption = "Strategy Basket Item Error"

Private Type mPrivate
    bOK As Boolean
    astrSystemInfo As cGdArray
    lOutlineLevel As Long
    
    basketItem As cStrategyBasketItem
    ParmsGrid As cStrategyBasketItemParmGrid
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the Form
'' Inputs:      Basket Item, Outline level
'' Returns:     True on OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(basketItem As cStrategyBasketItem, ByVal lOutlineLevel As Long) As Boolean
On Error GoTo ErrSection:

    m.lOutlineLevel = lOutlineLevel
    Set m.basketItem = basketItem.MakeCopy(False)
    
    'JM 12-18-2015: need to call this here because the grids are getting loaded before showing the form
    FixFormControls Me, ALT_GRID_ROW_COLOR
    
    Set m.ParmsGrid = New cStrategyBasketItemParmGrid
    m.ParmsGrid.InitForStrategyBasketItem fgInputs, basketItem
    
    LoadStrategiesCombo
    LoadFiltersCombo
    
    BasketItemToControls
    EnableControls
    
    ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR

    If m.bOK Then
        BasketItemFromControls
        Set basketItem = m.basketItem
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmStrategyBasketItem.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboPeriod_Validate
'' Description: Validate what the user selected as a period
'' Inputs:      Whether to Cancel the change
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboPeriod_Validate(Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strPeriod As String             ' Period string

    If Visible Then
        strPeriod = FixPeriod(cboPeriod.Text)
        
        If UCase(strPeriod) = "FRACTZEN" And Not g.FractZen.Allowed Then
            strPeriod = "Daily"
            InfBox "You are not authorized to use FractZen bars", "!", , kErrorCaption
            MoveFocus cboPeriod
        End If
        
        If strPeriod <> cboPeriod.Text Then
            cboPeriod.Text = strPeriod
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.cboPeriod_Validate"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboSystems_Click
'' Description: When the user changes the system, change the dates and period
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSystems_Click()
On Error GoTo ErrSection:

    ChangeStrategy cboSystems.ListIndex

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmStrategyBasketItem.cboSystems_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.cmdCancel_Click"

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

    Lookup
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.cmdLookup_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Unload the form and Save
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    'Dim lFieldNum As Long               ' Field number for the selected group

    ' If they chose to run the strategy on a symbol, make sure that they entered a symbol...
    If (optSymbol.Value = True) And (Len(Trim(txtSymbol.Text)) = 0) Then
        InfBox "Please enter in a Symbol", "!", , kErrorCaption
        MoveFocus txtSymbol
    ElseIf (optToDate.Value = True) And (gdToDate.Value <= gdFromDate.Value) Then
        InfBox "The to date must be after the from date", "!", , kErrorCaption
        MoveFocus gdToDate
    Else
        ' If they chose a symbol group or filter, make sure that there are 500 or less symbols in it...
        'If optSymbolGroup Then
        '    lFieldNum = g.SymbolPool.FieldNumForID(cboFilters.SelectedItem.Key)
        '    If g.SymbolPool.ArrayTable.FieldArray(lFieldNum).CountOf(1) > 500 Then
        '        InfBox "You cannot select a symbol group or filter with more than 500 symbols in it.", "!", , kErrorCaption
        '        MoveFocus cboFilters
        '        GoTo ErrExit
        '    End If
        'End If
    
        m.bOK = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.cmdOK_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Allow the form to capture an F1 to bring up the help
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.Form_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize and Place the Form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    ' Set the form caption
    Caption = "Strategy Basket Item"
    Me.Icon = Picture16(ToolbarIcon("ID_StrategyBaskets"), , True)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me
    
    With cboPeriod
        .AddItem "5 Minute"
        .AddItem "10 Minute"
        .AddItem "30 Minute"
        .AddItem "60 Minute"
        .AddItem "Daily"
        .AddItem "Weekly"
        .AddItem "Monthly"
        .AddItem "Quarterly"
        .AddItem "Yearly"
    
        If g.FractZen.Allowed Then
            .AddItem "FractZen" '"Auto Breakout"
        End If
    End With
    cboPeriod.Text = "Daily"
    
    cboFilters.ImageList = frmMain.img16
    cboFilters.Locked = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: When the user hits the 'X', unload the form without saving
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdFromDate_Changed
'' Description: Limit the ToDate according to the FromDate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdFromDate_Changed()
On Error GoTo ErrSection:

    gdToDate.MinDate = gdFromDate.Value

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.gdFromDate_Changed"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSymbol_Click
'' Description: Enable/Disable Controls based on the click
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSymbol_Click()
On Error GoTo ErrSection:

    EnableControls
    
    If Visible = True And Len(txtSymbol.Text) = 0 Then
        Lookup
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.optSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSymbolGroup_Click
'' Description: Enable/Disable Controls based on the click
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSymbolGroup_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.optSymbolGroup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optToDate_Click
'' Description: Enable/Disable controls based on the click
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optToDate_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.optToDate_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optToEnd_Click
'' Description: Enable/Disable controls based on the click
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optToEnd_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.optToEnd_Click"

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

    Lookup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.txtSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load up the filters combo box with the symbol groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFiltersCombo()
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
                        If bScans Then
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
                        
                        aItems.Add .ArrayTable.FieldName(lIndex) & vbTab & strID & vbTab & strPicture
                    End If
                End If
            End If
        Next
    End With
    If iSortStart > 0 Then
        aItems.Sort eGdSort_IgnoreCase, iSortStart
    End If

    For lIndex = 0 To aItems.Size - 1
        strItem = aItems(lIndex)
        cboFilters.ComboItems.Add , Parse(strItem, vbTab, 2), Parse(strItem, vbTab, 1), Parse(strItem, vbTab, 3)
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
'' Function:    EnableControls
'' Description: Enable/Disable controls on the form as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bEnableSystemAndSymbol As Boolean

    bEnableSystemAndSymbol = (m.lOutlineLevel = 0)
    
    Enable cboSystems, bEnableSystemAndSymbol
    Enable optSymbolGroup, bEnableSystemAndSymbol
    Enable cboFilters, bEnableSystemAndSymbol And optSymbolGroup
    Enable optSymbol, bEnableSystemAndSymbol
    Enable txtSymbol, bEnableSystemAndSymbol And optSymbol
    Enable cmdLookup, bEnableSystemAndSymbol And optSymbol
    Enable cboPeriod, bEnableSystemAndSymbol
    Enable gdFromDate, bEnableSystemAndSymbol
    Enable optToDate, bEnableSystemAndSymbol
    Enable gdToDate, bEnableSystemAndSymbol And optToDate
    Enable optToEnd, bEnableSystemAndSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Lookup
'' Description: Allow the user to lookup a symbol with the symbol selector
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Lookup()
On Error GoTo ErrSection:

    Dim astrSymbols As cGdArray         ' Symbol back from the symbol selector
    
    Set astrSymbols = frmSymbolSelector.ShowMe(txtSymbol.Text, False, True, , True)
    If astrSymbols.Size > 0 Then
        txtSymbol.Text = astrSymbols(0)
    End If

ErrExit:
    Set astrSymbols = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.Lookup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadStrategiesCombo
'' Description: Load the strategies combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadStrategiesCombo()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set m.astrSystemInfo = New cGdArray
    m.astrSystemInfo.Create eGDARRAY_Strings
        
    Set rs = mSysNav.LoadStrategiesRecordset(True)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            If mSysNav.IncludeStrategiesFromRecordset(rs) Then
                cboSystems.AddItem rs!SystemName
                cboSystems.ItemData(cboSystems.NewIndex) = rs!SystemNumber
                
                m.astrSystemInfo.Add rs!SystemName & vbTab & rs!BarTimeFrame & vbTab & CStr(CLng(DateOf(rs!FromDate))) & vbTab & CStr(CLng(DateOf(rs!ToDate))) & vbTab & CStr(rs!ToEndOfData)
            End If
            
            rs.MoveNext
        Loop
        
        cboSystems.ListIndex = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.LoadStrategiesCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeStrategy
'' Description: Change the strategy information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeStrategy(ByVal lIndex As Long)
On Error GoTo ErrSection:

    Dim astrSystemInfo As cGdArray      ' Array of system information
    
    If (lIndex >= 0) And (lIndex < m.astrSystemInfo.Size) Then
        Set astrSystemInfo = New cGdArray
        astrSystemInfo.SplitFields m.astrSystemInfo(lIndex), vbTab
        
        cboPeriod.Text = astrSystemInfo(1)
        gdFromDate = ValOfText(astrSystemInfo(2))
        gdToDate = ValOfText(astrSystemInfo(3))
        optToEnd = CBool(astrSystemInfo(4))
        optToDate = Not optToEnd
        
        If Visible Then
            m.ParmsGrid.ChangeStrategy cboSystems.ItemData(cboSystems.ListIndex)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.ChangeStrategy"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketItemToControls
'' Description: Set the controls based on the basket item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BasketItemToControls()
On Error GoTo ErrSection:

    mGenesis.SelectComboByItemData cboSystems, m.basketItem.StrategyID
    If (UCase(m.basketItem.Period) = "FRACTZEN") And (g.FractZen.Allowed = False) Then
        mGenesis.SelectComboByText cboPeriod, "Daily"
    ElseIf mGenesis.SelectComboByText(cboPeriod, m.basketItem.Period) = False Then
        cboPeriod.Text = m.basketItem.Period
    End If
    
    If Len(m.basketItem.SymbolGroupID) > 0 Then
        cboFilters.ComboItems(m.basketItem.SymbolGroupID).Selected = True
    End If
    optSymbol.Value = ((m.basketItem.SymbolID <> 0) Or (Len(m.basketItem.Symbol) > 0))
    optSymbolGroup.Value = Not optSymbol.Value
    txtSymbol.Text = m.basketItem.Symbol
        
    If m.basketItem.FromDate <> 0 Then
        gdFromDate.Value = m.basketItem.FromDate
    End If
    If m.basketItem.ToDate <> 0 Then
        gdToDate.Value = m.basketItem.ToDate
        If m.basketItem.ToEndOfData Then
            optToEnd.Value = True
        Else
            optToDate.Value = True
        End If
    End If
    
    'CheckBoxValue chkSplit, m.BasketItem.Split
    ' We don't want to handle the unsplit stuff right now...
    chkSplit.Value = vbChecked
    chkSplit.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.BaskteItemToControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketItemFromControls
'' Description: Set the basket item based on the controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BasketItemFromControls()
On Error GoTo ErrSection:

    m.basketItem.StrategyID = cboSystems.ItemData(cboSystems.ListIndex)
    m.basketItem.StrategyName = cboSystems.Text
    If (m.lOutlineLevel = 0) And (optSymbol.Value = True) Then
        m.basketItem.SymbolGroupID = ""
        m.basketItem.SymbolGroupName = ""
    Else
        m.basketItem.SymbolGroupID = cboFilters.SelectedItem.Key
        m.basketItem.SymbolGroupName = cboFilters.SelectedItem.Text
    End If
    If (m.lOutlineLevel = 0) And (optSymbolGroup.Value = True) Then
        m.basketItem.Symbol = ""
        m.basketItem.SymbolID = 0&
    Else
        m.basketItem.Symbol = Trim(txtSymbol.Text)
        m.basketItem.SymbolID = GetSymbolID(m.basketItem.Symbol)
    End If
    'm.basketItem.Period = PeriodStr(cboPeriod.Text)
    m.basketItem.Period = FixPeriod(cboPeriod.Text)
    m.basketItem.FromDate = gdFromDate.Value
    m.basketItem.ToDate = gdToDate.Value
    m.basketItem.ToEndOfData = optToEnd.Value
    m.basketItem.Split = CheckBoxValue(chkSplit)
    
    m.ParmsGrid.BasketItemParmsFromGrid m.basketItem.Parms
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasketItem.BasketItemFromControls"
    
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

    If (UCase(strPeriod) = "AUTO BREAKOUT") Or (UCase(strPeriod) = "FRACTZEN") Then
        strReturn = "FractZen" 'strPeriod
    Else
        strReturn = GetPeriodStr(strPeriod)
    End If
    
    FixPeriod = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasketItem.FixPeriod"
    
End Function

