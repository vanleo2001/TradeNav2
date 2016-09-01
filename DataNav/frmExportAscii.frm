VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmExportAscii 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ASCII Exporting Options"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraFiles 
      Height          =   1215
      Left            =   3480
      TabIndex        =   14
      Top             =   2520
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
      Caption         =   "frmExportAscii.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportAscii.frx":003E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportAscii.frx":005E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optMultipleFiles 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
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
         Caption         =   "frmExportAscii.frx":007A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmExportAscii.frx":00B6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":00D6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtFileName 
         Height          =   285
         Left            =   600
         TabIndex        =   17
         Top             =   720
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmExportAscii.frx":00F2
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
         Tip             =   "frmExportAscii.frx":0112
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":0132
      End
      Begin HexUniControls.ctlUniRadioXP optSingleFile 
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   480
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
         Caption         =   "frmExportAscii.frx":014E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportAscii.frx":0184
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":01A4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboDecimal 
      Height          =   315
      Left            =   4920
      TabIndex        =   13
      Top             =   2010
      Width           =   1215
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
      Tip             =   "frmExportAscii.frx":01C0
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmExportAscii.frx":01E0
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboExtension 
      Height          =   315
      Left            =   4680
      TabIndex        =   11
      Top             =   1530
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
      Tip             =   "frmExportAscii.frx":01FC
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
      MouseIcon       =   "frmExportAscii.frx":021C
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdDown 
      Height          =   315
      Left            =   1920
      TabIndex        =   20
      Top             =   4080
      Width           =   1095
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
      Caption         =   "frmExportAscii.frx":0238
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmExportAscii.frx":026C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmExportAscii.frx":028C
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdUp 
      Height          =   315
      Left            =   480
      TabIndex        =   19
      Top             =   4080
      Width           =   1095
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
      Caption         =   "frmExportAscii.frx":02A8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmExportAscii.frx":02D8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmExportAscii.frx":02F8
      RightToLeft     =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid fgFields 
      Height          =   2535
      Left            =   360
      TabIndex        =   18
      Top             =   1440
      Width           =   2775
      _cx             =   4895
      _cy             =   4471
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
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   375
      Left            =   405
      TabIndex        =   0
      Top             =   240
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
      Caption         =   "frmExportAscii.frx":0314
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportAscii.frx":0340
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportAscii.frx":0360
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboDelimiter 
         Height          =   315
         Left            =   4080
         TabIndex        =   4
         Top             =   0
         Width           =   1695
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
         Tip             =   "frmExportAscii.frx":037C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":039C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboDateFormat 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   1695
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
         Tip             =   "frmExportAscii.frx":03B8
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
         MouseIcon       =   "frmExportAscii.frx":03D8
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   3240
         Top             =   30
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
         Caption         =   "frmExportAscii.frx":03F4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportAscii.frx":0428
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":0448
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   0
         Top             =   30
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
         Caption         =   "frmExportAscii.frx":0464
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportAscii.frx":049C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":04BC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraHeader 
      Height          =   615
      Left            =   405
      TabIndex        =   5
      Top             =   720
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
      Caption         =   "frmExportAscii.frx":04D8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportAscii.frx":0504
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportAscii.frx":0524
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtHeader 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   3375
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmExportAscii.frx":0540
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
         Tip             =   "frmExportAscii.frx":0560
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":0580
      End
      Begin HexUniControls.ctlUniRadioXP optCustom 
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   240
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
         Caption         =   "frmExportAscii.frx":059C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportAscii.frx":05C8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":05E8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optColHeaders 
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   0
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
         Caption         =   "frmExportAscii.frx":0604
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmExportAscii.frx":0640
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":0660
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkHeaderLine 
         Height          =   255
         Left            =   0
         TabIndex        =   6
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
         Caption         =   "frmExportAscii.frx":067C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportAscii.frx":06B2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":06D2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   435
      Left            =   945
      TabIndex        =   1
      Top             =   4620
      Width           =   4695
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
      Caption         =   "frmExportAscii.frx":06EE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportAscii.frx":071A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportAscii.frx":073A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDefault 
         Height          =   435
         Left            =   3360
         TabIndex        =   3
         Top             =   0
         Width           =   1335
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
         Caption         =   "frmExportAscii.frx":0756
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExportAscii.frx":0794
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":07B4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   1680
         TabIndex        =   10
         Top             =   0
         Width           =   1335
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
         Caption         =   "frmExportAscii.frx":07D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExportAscii.frx":07FE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":081E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   435
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1335
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
         Caption         =   "frmExportAscii.frx":083A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExportAscii.frx":0860
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExportAscii.frx":0880
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label4 
      Height          =   255
      Left            =   3480
      Top             =   2040
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
      Caption         =   "frmExportAscii.frx":089C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmExportAscii.frx":08E0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportAscii.frx":0900
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label3 
      Height          =   255
      Left            =   3480
      Top             =   1560
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
      Caption         =   "frmExportAscii.frx":091C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmExportAscii.frx":095A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportAscii.frx":097A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmExportAscii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmExportAscii.frm
'' Description: Allows the user to custom define the ASCII export options
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 08/10/01  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Enum eGDCols
    eGDCol_Active = 0
    eGDCol_Name = 1
End Enum
Private Const kNumCols = 2

Private Type mPrivate
    bOK As Boolean
End Type

Private m As mPrivate

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the fields grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgFields
        .Redraw = flexRDNone
        SetupGrid fgFields, eGridMode_Grid
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow Or flexExMoveRows
        .GridLines = flexGridNone
        .GridLinesFixed = flexGridInset
        
        .Rows = 1
        .FixedRows = 1
        .Cols = kNumCols
        .FixedCols = 0
        
        .Cell(flexcpText, 0, GDCol(eGDCol_Active)) = "Show"
        .Cell(flexcpText, 0, GDCol(eGDCol_Name)) = "Name"
        
        .ColDataType(GDCol(eGDCol_Active)) = flexDTBoolean
        
        .AutoSize 0
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.InitGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboExtension_Change
'' Description: If the user enters in an extension longer than 3 characters,
''              warn them and change back to first option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboExtension_Change()
On Error GoTo ErrSection:

    If Len(cboExtension.Text) > 3 Then
        InfBox "Extension cannot be longer than 3 characters", "!", , "Error"
        cboExtension.Text = Mid(cboExtension.Text, 1, 3)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.cboExtension.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkHeaderLine_Click
'' Description: If the turns header line on, enable the option buttons, else
''              if the user turns header line off, disable the option buttons
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkHeaderLine_Click()
On Error GoTo ErrSection:

    If chkHeaderLine.Value = vbChecked Then
        optColHeaders.Enabled = True
        optCustom.Enabled = True
        txtHeader.Enabled = (optCustom.Value = True)
    Else
        optColHeaders.Enabled = False
        optCustom.Enabled = False
        txtHeader.Enabled = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.chkHeaderLine.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on the cancel button, hide the form and
''              let the ShowMe take over
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
    RaiseError "frmExportAscii.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDefault_Click
'' Description: If the user clicks on the Defaults button, reset all of the
''              controls back to the default
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDefault_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    cboDateFormat.ListIndex = 0
    cboDecimal.ListIndex = 0
    cboDelimiter.ListIndex = 0
    cboExtension.ListIndex = 0
    
    chkHeaderLine.Value = vbUnchecked
    chkHeaderLine_Click
    optMultipleFiles.Value = True
    optMultipleFiles_Click
    optColHeaders.Value = True
    optColHeaders_Click
    
    With fgFields
        For lIndex = .FixedRows To .Rows - 1
            CheckedCell(fgFields, lIndex, GDCol(eGDCol_Active)) = True
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.cmdDefault.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDown_Click
'' Description: If the user clicks on the "Move Down" button, move the selected
''              row down one row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDown_Click()
On Error GoTo ErrSection:

    With fgFields
        If .RowSel > .FixedRows - 1 And .RowSel < .Rows - 1 Then
            .RowPosition(.RowSel) = .RowSel + 1
            .Row = .RowSel + 1
            .RowSel = .Row
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.cmdDown.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on the OK button, hide the form and let
''              ShowMe take over
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim bPrices As Boolean, bFields As Boolean
    Dim lRow&, strField$
    Dim strDate As String               ' Temporary formatted version of the date

    ' see what type of fields have been selected
    With fgFields
        For lRow = .FixedRows To .Rows - 1
            If CheckedCell(fgFields, lRow, GDCol(eGDCol_Active)) Then
                bFields = True
                strField = UCase(.Cell(flexcpText, lRow, GDCol(eGDCol_Name)))
                If strField <> "SYMBOL" And strField <> "DESCRIPTION" Then
                    bPrices = True
                End If
            End If
        Next lRow
    End With
    
    ' verify stuff
    If Not bFields Then
        InfBox "Please specify one or more data fields to export.", "e", , "Error"
        Exit Sub
    ElseIf Not bPrices Then
        ' if no prices, then must go to a single file
        optSingleFile = True
    End If
    
    If optSingleFile Then
        If Len(Trim(txtFileName)) = 0 Then
            InfBox "Please specify a filename.", "e", , "Error"
            MoveFocus txtFileName
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    strDate = Format(Date, cboDateFormat.Text)
    'If IsDate(CVDate(Format(Date, cboDateFormat.Text))) = False Then
    If strDate = cboDateFormat.Text Then
        InfBox "Please specify a valid date format", "e", , "Error"
        MoveFocus cboDateFormat
        Exit Sub
    End If
    On Error GoTo ErrSection

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdUp_Click
'' Description: If the user clicks on the "Move Up" button, move the selected
''              row up one row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdUp_Click()
On Error GoTo ErrSection:

    With fgFields
        If .RowSel > .FixedRows Then
            .RowPosition(.RowSel) = .RowSel - 1
            .Row = .RowSel - 1
            .RowSel = .Row
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.cmdUp.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_AfterMoveRow
'' Description: After the user has moved a row, make sure that it is selected
'' Inputs:      Row moved, New Position of row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_AfterMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    With fgFields
        .Row = Position
        .RowSel = Position
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.fgFields.AfterMoveRow", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_BeforeEdit
'' Description: Only allow the user to edit the "Active" field
'' Inputs:      Row and Column being edited, Whether or not to cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> GDCol(eGDCol_Active) Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.fgFields.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_BeforeMouseDown
'' Description: Mark the row as a drag row to allow the user to move it where
''              they want it
'' Inputs:      Mouse button pressed, Shift/Ctrl/Alt status, Location of click,
''              Whether or not to cancel the operation
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of new row
    Dim lRow As Long                    ' Row being moved
    
    With fgFields
        lRow = .MouseRow
        .Row = lRow
        .RowSel = lRow
        
        .Refresh
        lPos = .DragRow(lRow)
        If lPos <> lRow Then
            Cancel = True
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.fgFields.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_Click
'' Description: Make sure to select the row that the user clicks on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_Click()
On Error GoTo ErrSection:

    With fgFields
        .Row = .MouseRow
        .RowSel = .Row
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.fgFields.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_AfterRowColChange
'' Description: Enable/Disable the buttons according to what row the user is
''              currently on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With fgFields
        cmdUp.Enabled = NewRow > .FixedRows
        cmdDown.Enabled = NewRow > .FixedRows - 1 And NewRow < .Rows - 1
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.fgFields.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, load up the combo boxes and set the
''              defaults
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_ExportData"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    With cboDateFormat
        .AddItem UCase(DateFormat("Format", MM_DD_YYYY))
        .AddItem UCase(DateFormat("Format", MM_DD_YY))
        .AddItem UCase(DateFormat("Format", M_D_YY))
        .AddItem "YYYYMMDD"
        .AddItem "YYMMDD"
        .ListIndex = 0
    End With
    
    With cboDelimiter
        .AddItem "Comma"
        .AddItem "Tab"
        .AddItem "Space"
        .AddItem "Pipe (|)"
        .AddItem "Semicolon (;)"
        .ListIndex = 0
    End With
    
    With cboExtension
        .AddItem "CSV"
        .AddItem "TXT"
        .ListIndex = 0
    End With
    
    With cboDecimal
        .AddItem "Decimal (.)"
        .AddItem "Comma (,)"
        .ListIndex = 0
    End With
    
    chkHeaderLine.Value = vbUnchecked
    chkHeaderLine_Click
    optColHeaders.Value = True
    optColHeaders_Click
    optMultipleFiles.Value = True
    optMultipleFiles_Click
    
    InitGrid
    With fgFields
        .AddItem vbChecked & vbTab & "Symbol"
        .AddItem vbUnchecked & vbTab & "Description"
        .AddItem vbChecked & vbTab & "Date"
        .AddItem vbChecked & vbTab & "Open"
        .AddItem vbChecked & vbTab & "High"
        .AddItem vbChecked & vbTab & "Low"
        .AddItem vbChecked & vbTab & "Close"
        .AddItem vbChecked & vbTab & "Total Volume"
        .AddItem vbChecked & vbTab & "Total Open Interest"
        .AddItem vbChecked & vbTab & "Contract Volume"
        .AddItem vbChecked & vbTab & "Contract Open Interest"
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user closes the form with the control menu, cancel the
''              unload, hide the form, and let the ShowMe take over from there
'' Inputs:      Whether or not to cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optColHeaders_Click
'' Description: If the user clicks the Column Headers option button, disable
''              the Custom text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optColHeaders_Click()
On Error GoTo ErrSection:

    txtHeader.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.optColHeaders.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optCustom_Click
'' Description: If the user clicks the Custom option button, enable the Custom
''              text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optCustom_Click()
On Error GoTo ErrSection:

    txtHeader.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.optCustom.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optMultipleFiles_Click
'' Description: If the user clicks on the multiple files option, disable the
''              filename text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optMultipleFiles_Click()
On Error GoTo ErrSection:

    txtFileName.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.optMultipleFiles.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSingleFile_Click
'' Description: If the user clicks on the single file option, enable the
''              filename text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSingleFile_Click()
On Error GoTo ErrSection:

    txtFileName.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportAscii.optSingleFile.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Show the form and pass back the appropriate information if the
''              user clicks on OK
'' Inputs:      ExportGroup object
'' Returns:     True if user clicked on OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ExportGroup As cExportGroup) As Boolean
On Error GoTo ErrSection:

    Dim strField As String              ' Field from the fields string
    Dim lRow As Long                    ' Index into a for loop
    Dim lPosition As Long               ' Position to move the row to
    Dim bFound As Boolean               ' Was the field found in the grid?
    Dim lField As Long
    Dim strTemp As String

    With ExportGroup
        If Len(.DateFormat) > 0 Then
            cboDateFormat.Text = .DateFormat
            If Len(.Delimiter) > 0 Then cboDelimiter.Text = .Delimiter
            If .ShowHeader Then chkHeaderLine.Value = vbChecked Else chkHeaderLine.Value = vbUnchecked
            optCustom.Value = .CustomHeader
            txtHeader.Text = .HeaderLine
            cboDecimal.Text = .DecimalChar
            cboExtension.Text = .Extension
            optSingleFile.Value = .SingleFile
            txtFileName.Text = .FileName
            
            With fgFields
                lField = 1
                strField = Parse(ExportGroup.Fields, "|", lField)
                lPosition = .FixedRows
                Do While strField <> ""
                    For lRow = .FixedRows To .Rows - 1
                        If .Cell(flexcpText, lRow, GDCol(eGDCol_Name)) = Parse(strField, ";", 2) Then
                            .Cell(flexcpChecked, lRow, GDCol(eGDCol_Active)) = ValOfText(Parse(strField, ";", 1))
                            .RowPosition(lRow) = lPosition
                            Exit For
                        End If
                    Next lRow
                    lPosition = lPosition + 1
                    lField = lField + 1
                    strField = Parse(ExportGroup.Fields, "|", lField)
                Loop
            End With
        End If
    End With

    ShowForm Me, True
    
    If m.bOK = True Then
        With ExportGroup
            .DateFormat = cboDateFormat.Text
            .Delimiter = cboDelimiter.Text
            .ShowHeader = (chkHeaderLine.Value = vbChecked)
            .CustomHeader = optCustom.Value
            .HeaderLine = txtHeader.Text
            .DecimalChar = cboDecimal.Text
            .Extension = cboExtension.Text
            .SingleFile = optSingleFile.Value
            
            strTemp = FileExt(txtFileName.Text)
            If strTemp <> "" Then .Extension = strTemp
            strTemp = FilePath(txtFileName.Text)
            If strTemp <> "" Then .Path = strTemp
            .FileName = FileBase(txtFileName.Text)
            
            With fgFields
                strField = ""
                For lRow = .FixedRows To .Rows - 1
                    strField = strField & .Cell(flexcpChecked, lRow, GDCol(eGDCol_Active)) & _
                        ";" & .Cell(flexcpText, lRow, GDCol(eGDCol_Name)) & "|"
                Next lRow
                strField = Mid(strField, 1, Len(strField) - 1)
            End With
            .Fields = strField
        End With
    End If
    
    ShowMe = m.bOK
    Unload Me
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExportAscii.ShowMe", eGDRaiseError_Raise

End Function

