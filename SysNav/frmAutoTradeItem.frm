VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAutoTradeItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMenu 
      Left            =   8160
      Top             =   5220
   End
   Begin VSFlex7LCtl.VSFlexGrid fgBasketItems 
      Height          =   1635
      Left            =   120
      TabIndex        =   22
      Top             =   6240
      Width           =   4815
      _cx             =   8493
      _cy             =   2884
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
   Begin HexUniControls.ctlUniTextBoxXP txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   105
      Width           =   4095
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmAutoTradeItem.frx":0000
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
      Tip             =   "frmAutoTradeItem.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoTradeItem.frx":0040
   End
   Begin HexUniControls.ctlUniFrameWL fraData 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   1980
      Width           =   8415
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
      Caption         =   "frmAutoTradeItem.frx":005C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAutoTradeItem.frx":00BC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoTradeItem.frx":00DC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   840
         Width           =   615
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
         Caption         =   "frmAutoTradeItem.frx":00F8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0122
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0142
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
         Height          =   240
         Left            =   2280
         TabIndex        =   16
         Top             =   410
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
         Caption         =   "frmAutoTradeItem.frx":015E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0190
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":01B0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboBoxXP cboBarPeriod 
         Height          =   315
         Left            =   3420
         TabIndex        =   18
         Top             =   360
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
         Tip             =   "frmAutoTradeItem.frx":01CC
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
         MouseIcon       =   "frmAutoTradeItem.frx":01EC
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   375
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAutoTradeItem.frx":0208
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
         Tip             =   "frmAutoTradeItem.frx":0228
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0248
      End
      Begin HexUniControls.ctlUniLabelXP lblOnClose 
         Height          =   255
         Left            =   240
         Top             =   840
         Width           =   2775
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
         Caption         =   "frmAutoTradeItem.frx":0264
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAutoTradeItem.frx":02CE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":02EE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOnCloseTime 
         Height          =   255
         Left            =   3120
         Top             =   840
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
         Caption         =   "frmAutoTradeItem.frx":030A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0336
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0356
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSymbol 
         Height          =   255
         Left            =   240
         Top             =   390
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
         Caption         =   "frmAutoTradeItem.frx":0372
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAutoTradeItem.frx":03A2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":03C2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBarPeriod 
         Height          =   255
         Left            =   2880
         Top             =   390
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
         Caption         =   "frmAutoTradeItem.frx":03DE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAutoTradeItem.frx":040E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":042E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraStrategy 
      Height          =   1275
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   8415
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
      Caption         =   "frmAutoTradeItem.frx":044A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAutoTradeItem.frx":04A8
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoTradeItem.frx":04C8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdReset 
         Height          =   375
         Left            =   7140
         TabIndex        =   12
         Top             =   780
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
         Caption         =   "frmAutoTradeItem.frx":04E4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":051A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":053A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtNextEntry 
         Height          =   285
         Left            =   6120
         TabIndex        =   10
         Top             =   810
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAutoTradeItem.frx":0556
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
         Tip             =   "frmAutoTradeItem.frx":0576
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0596
      End
      Begin HexUniControls.ctlUniRadioXP optBasket 
         Height          =   220
         Left            =   420
         TabIndex        =   5
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
         Caption         =   "frmAutoTradeItem.frx":05B2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":05E2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0602
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optStrategy 
         Height          =   220
         Left            =   420
         TabIndex        =   3
         Top             =   240
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
         Caption         =   "frmAutoTradeItem.frx":061E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0652
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0672
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboBaskets 
         Height          =   315
         Left            =   6540
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
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
         Tip             =   "frmAutoTradeItem.frx":068E
         Sorted          =   -1  'True
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":06AE
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboStrategies 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   330
         Width           =   4935
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
         Tip             =   "frmAutoTradeItem.frx":06CA
         Sorted          =   -1  'True
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":06EA
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   1020
         TabIndex        =   8
         Top             =   810
         Width           =   3135
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
         Tip             =   "frmAutoTradeItem.frx":0706
         Sorted          =   -1  'True
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0726
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdScrollBar sbQty 
         Height          =   360
         Left            =   6840
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   750
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
      Begin HexUniControls.ctlUniLabelXP lblQtyNextEntry 
         Height          =   255
         Left            =   4380
         Top             =   840
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
         Caption         =   "frmAutoTradeItem.frx":0742
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0794
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":07B4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAccount 
         Height          =   255
         Left            =   240
         Top             =   840
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
         Caption         =   "frmAutoTradeItem.frx":07D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0802
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0822
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraOptions 
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   8475
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
      Caption         =   "frmAutoTradeItem.frx":083E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAutoTradeItem.frx":0890
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoTradeItem.frx":08B0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkExitAtEndOfDay 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   1020
         Width           =   7875
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
         Caption         =   "frmAutoTradeItem.frx":08CC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":094A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":096A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtTimeout 
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   645
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAutoTradeItem.frx":0986
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
         Tip             =   "frmAutoTradeItem.frx":09A6
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":09C6
      End
      Begin HexUniControls.ctlUniCheckXP chkConfirm 
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   300
         Width           =   8115
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
         Caption         =   "frmAutoTradeItem.frx":09E2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0ADA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0AFA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdScrollBar sbTimeout 
         Height          =   360
         Left            =   3840
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
      Begin HexUniControls.ctlUniLabelXP lblSeconds 
         Height          =   255
         Left            =   4140
         Top             =   660
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
         Caption         =   "frmAutoTradeItem.frx":0B16
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0B44
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0B64
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTimeout 
         Height          =   255
         Left            =   720
         Top             =   660
         Width           =   2175
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
         Caption         =   "frmAutoTradeItem.frx":0B80
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0BDE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0BFE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   2340
      TabIndex        =   14
      Top             =   5100
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
      Caption         =   "frmAutoTradeItem.frx":0C1A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAutoTradeItem.frx":0C46
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoTradeItem.frx":0C66
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdReports 
         Height          =   495
         Left            =   1380
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
         Caption         =   "frmAutoTradeItem.frx":0C82
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0CB2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0CD2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   2760
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
         Caption         =   "frmAutoTradeItem.frx":0CEE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0D1C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0D3C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   20
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
         Caption         =   "frmAutoTradeItem.frx":0D58
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAutoTradeItem.frx":0D7E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAutoTradeItem.frx":0D9E
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Left            =   120
      Top             =   120
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
      Caption         =   "frmAutoTradeItem.frx":0DBA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAutoTradeItem.frx":0DE6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAutoTradeItem.frx":0E06
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuChangeInputs 
         Caption         =   "Change Inputs"
      End
   End
End
Attribute VB_Name = "frmAutoTradeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAutoTradeItem.frm
'' Description: Allow the user to edit or create an Automated Trading Item
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/07/2009   DAJ         Allow PFG Forex for auto trading with flag file
'' 03/11/2010   DAJ         Use global populate accounts combo
'' 09/27/2010   DAJ         Allow automated trading for Genesis forex on sim account with flag file
'' 04/25/2011   DAJ         Don't allow automated trading for options or stocks
'' 05/11/2011   DAJ         Utilize IsLiveAccount function
'' 06/15/2011   DAJ         Allow 55/65/57/67 contracts on simulated account with flag file
'' 09/20/2011   DAJ         Open up 55/65/57/67 contracts for sim, live with flag file
'' 10/31/2012   DAJ         Put in allowances for IB and CNX forex
'' 12/17/2012   DAJ         Store off and use OnClose time in exchange time
'' 01/18/2013   DAJ         Don't allow automated trading for spreads
'' 04/03/2013   DAJ         Automated Strategy Baskets
'' 04/12/2013   DAJ         Fix for multiplier on basket items not always working ( #6811 )
'' 05/01/2013   DAJ         Shadow Trading
'' 05/14/2013   DAJ         Load a guru basket even if not the owner
'' 05/15/2013   DAJ         Load all baskets regardless of ownership, Set max units in SetupForm
'' 05/28/2013   DAJ         Only reset quantity editor if security type changed
'' 06/12/2013   DAJ         Symbol and quantity validation for automated trading items
'' 07/23/2013   DAJ         Don't allow basket to be automated if it contains a filter
'' 07/30/2013   DAJ         Don't allow exact match for enablement for shadow basket
'' 08/15/2013   DAJ         Allow user to reset the quantities on all basket items
'' 04/04/2014   DAJ         Allow automated trading for pyramiding systems
'' 04/23/2014   DAJ         Allow FractZen bars for automated trading
'' 06/26/2014   DAJ         Changed default on-close time from three minutes before close to one minute before
'' 08/19/2014   DAJ         Expose Strategy Basket Item Inputs
'' 08/22/2014   DAJ         Allow users to auto-trade continuous contracts on live account
'' 11/10/2014   DAJ         Only save the basket items if the use basket option is chosen
'' 12/03/2014   DAJ         Change the symbol and quantity for forex when account changes
'' 01/05/2015   DAJ         Changed default on-close time from one minute before close to two minutes before
'' 03/31/2015   DAJ         Changed default for confirmation box to unchecked
'' 05/26/2015   DAJ         Fix for miscalculating on-close time for New Zealand
'' 06/02/2015   DAJ         Don't allow for more than kSN_BASKETLIMIT strategy basket items
'' 08/05/2015   DAJ         Don't allow closing time to be outside of market hours
'' 10/06/2015   DAJ         Always exit at end of day flag for automated trading
'' 10/07/2015   DAJ         Confirmation dialog when turn on ExitAtEndOfDay flag for basket
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Strategy = 0
    eGDCol_Symbol
    eGDCol_Period
    eGDCol_Multiplier
    eGDCol_QtyNextEntry
    eGDCol_Name
    eGDCol_OnCloseTime
    eGDCol_ClosingTime
    eGDCol_MinutesBefore
    eGDCol_SortKey
    eGDCol_BasketItemID
    eGDCol_Overrides
    eGDCol_BasketItemKey
    eGDCol_OutlineLevel
    eGDCol_Units
    eGDCol_SymbolError
    eGDCol_QuantityError
    eGDCol_HasExposed
    eGDCol_BasketItemSymbol
    eGDCol_BasketItemMult
    eGDCol_NumCols
End Enum

Private Const kExtendedCol = eGDCol_Strategy
Private Const kErrorCaption = "Automated Trading Error"

Private Type mPrivate
    bActive As Boolean                  ' Is the automated trading object active?
    bOK As Boolean                      ' Did the user click on OK?
    Qty As cPriceEditor                 ' Quantity editor
    TimeOut As cPriceEditor             ' Timeout editor
    
    lMinutesBefore As Long              ' Minutes before close to check on-close orders
    dOnCloseTime As Double              ' Time to check On Close orders
    
    lBasketID As Long                   ' Basket currently loaded in the grid
    dBasketLastModified As Double       ' Date/Time the selected strategy basket was last modified
    
    nPrevColWidth As Long               ' Previous column width
    
    TradeItem As cAutoTradeItem         ' Automated Trading Item Object
    BasketTradeItems As cGdTree         ' Collection of automated trading items for the basket
    astrChildNames As cGdArray          ' Array of names of child items
    strName As String                   ' Name from the text box
    lMaxUnits As Long                   ' Maximum number of units allowed for the user
    strPreviousSymbol As String         ' Previous symbol selected
    lNumBasketItems As Long             ' Number of basket items
End Type
Private m As mPrivate

Private Property Get GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Property

Private Property Get AccountID() As Long
    If cboAccounts.ListIndex >= 0 Then
        AccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
    Else
        AccountID = -1&
    End If
End Property
Private Property Let AccountID(ByVal lAccountID As Long)
    SelectComboByItemData cboAccounts, lAccountID
End Property

Private Property Get BasketID() As Long
    If cboBaskets.ListIndex >= 0 Then
        BasketID = cboBaskets.ItemData(cboBaskets.ListIndex)
    Else
        BasketID = -1&
    End If
End Property
Private Property Let BasketID(ByVal lBasketID As Long)
    SelectComboByItemData cboBaskets, lBasketID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Auto Trade Item (in/out), Is it Active?, Basket Items (out)
'' Returns:     True if user clicked on OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(TradeItem As cAutoTradeItem, ByVal bActive As Boolean, Optional BasketItems As cGdTree) As Boolean
On Error GoTo ErrSection:

    Set m.TradeItem = TradeItem
    
    InitBasketItemsGrid

    LoadBasketsCombo TradeItem.StrategyBasketID
    LoadStrategiesCombo
    PopulateAccountsCbo cboAccounts, TradeItem.AccountID, True
    m.bActive = bActive
    
    Set m.TimeOut = New cPriceEditor
    m.TimeOut.Init sbTimeout, txtTimeout, Nothing, , 5, 60, True
    
    m.lMinutesBefore = TradeItem.MinutesBefore
    LoadBasketTradeItems
    SetupForm
    
    EnableControls
    ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR
    
    If m.bOK Then
        Save
        Set TradeItem = m.TradeItem
        Set BasketItems = BasketItemsFromGrid
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmAutoTradeItem.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccounts_Click
'' Description: Update symbol after the user changes accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    If (Visible = True) Then
        If (optStrategy.Value = True) Then
            If (Len(txtSymbol.Text) > 0) Then
                UpdateSymbol Trim(txtSymbol.Text)
            End If
        ElseIf (optBasket.Value = True) Then
            If ChangeForex = True Then
                ChangeBasketItemQuantities True
            End If
            
            ValidateBasketSymbols
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.cboAccounts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboBarPeriod_Validate
'' Description: Change the user's entry to our version of the period string
'' Inputs:      Whether to Cancel the Update
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboBarPeriod_Validate(Cancel As Boolean)
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
    RaiseError "frmAutoTradeItem.cboBarPeriod.Validate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboBaskets_Click
'' Description: Handle the user changing the strategy basket
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboBaskets_Click()
On Error GoTo ErrSection:

    If Visible Then
        If optBasket.Value = False Then
            optBasket.Value = True
        Else
            LoadBasketItemsGrid
            InitQuantityEditor
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.cboBaskets_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkExitAtEndOfDay_Click
'' Description: Notification that the user has clicked on the ExitAtEndOfDay check box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkExitAtEndOfDay_Click()
On Error GoTo ErrSection:

    If Visible Then
        If CheckBoxValue(chkExitAtEndOfDay) = True Then
            If optBasket.Value = True Then
                If InfBox("Turning this option on will exit ALL open|positions in ALL of the basket items|at each appropriate On-Close time.||Do you want to continue?|", "?", "+Yes|-No", "Confirmation") = "N" Then
                    CheckBoxValue(chkExitAtEndOfDay) = False
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.chkExitAtEndOfDay_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving information
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
    RaiseError "frmAutoTradeItem.cmdCancel", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to edit the on-close time
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    Dim dSessionEnd As Double           ' Session ending time for the symbol
    Dim dClosingTime As Double          ' User specified closing time
    Dim lMinutesBefore As Long          ' Minutes before closing time to check
    
    lMinutesBefore = m.lMinutesBefore
    dClosingTime = m.dOnCloseTime + (lMinutesBefore / 1440#)
    
    If ChangeOnCloseTime(txtSymbol.Text, dClosingTime, lMinutesBefore) Then
        m.dOnCloseTime = dClosingTime - (lMinutesBefore / 1440#)
        lblOnCloseTime.Caption = DateFormat(m.dOnCloseTime, NO_DATE, HH_MM)
        m.lMinutesBefore = lMinutesBefore
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: Allow the user to change symbols with the symbol selector
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookup_Click()
On Error GoTo ErrSection:

    ChangeSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.cmdLookup.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Verify User Information, then save information and unload
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim strError As String              ' Error message to display to the user
    Dim bContinue As Boolean            ' Continue?
    Dim bHasInvalidQuantities As Boolean ' Do we have invalid quantities?
    
    ' Make sure the length of the name is between 1 and 50 characters...
    If Len(Trim(txtName.Text)) = 0 Then
        MoveFocus txtName
        InfBox "Please enter in a name for the|Automated Trading Item", "!", , "Error"
    ElseIf Len(Trim(txtName.Text)) > 50 Then
        MoveFocus txtName
        InfBox "The Automated Trading Item name must be less than 50 characters in length", "!", , "Error"
    ElseIf InStr(txtName.Text, "|") <> 0 Then
        MoveFocus txtName
        InfBox "The Automated Trading Item name cannot have a pipe (|) character in it", "!", , "Error"
    ElseIf g.TradingItems.NameExistsInDb(Trim(txtName.Text), m.TradeItem.AutoTradeItemID) Then
        MoveFocus txtName
        InfBox "The Automated Trading Item name|must be unique", "!", , "Error"
    ElseIf m.lNumBasketItems > kSN_BASKETLIMIT Then
        InfBox "You cannot set up an automated trading item with more than " & Str(kSN_BASKETLIMIT) & " strategy basket items", "!", , "Error"
    Else
        bContinue = True
        
        ' Make sure the user has selected a symbol...
        If optStrategy.Value = True Then
            strError = AutomatedSymbolError(AccountID, Trim(txtSymbol.Text), "Automated Trading", True)
            If Len(strError) > 0 Then
                MoveFocus txtSymbol
                InfBox strError, "!", , "Error"
                ChangeSymbol
                
                bContinue = False
            End If
        Else
            If ValidateBasketSymbols(bHasInvalidQuantities) = 0 Then
                InfBox "You must have at least one valid symbol for automated trading", "!", , "Error"
                bContinue = False
            End If
            If bHasInvalidQuantities Then
                InfBox "You have one or more items with|an invalid quantity", "!", , "Error"
                bContinue = False
            End If
        End If
    
        If bContinue Then
            m.bOK = True
            Me.Hide
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdReports_Click
'' Description: Allow the user to view performance reports for the current info
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdReports_Click()
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated trading item from the controls
    
    Set TradeItem = TradeItemFromControls
    TradeItem.PerformanceReport

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.cmdReports_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdReset_Click
'' Description: Allow the user to reset the quantities of the basket items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdReset_Click()
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from an InfBox
    
    strReturn = InfBox("This will reset the next entry quantities|on all items in your basket.||Do you want to continue?", "?", "+Yes|-No", "Quantity Reset")
    If strReturn = "Y" Then
        ChangeBasketItemQuantities True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.cmdReset_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgBasketItems_AfterRowColChange
'' Description: Handle the user changing cells in the grid
'' Inputs:      Old Row, Old Column, New Row, New Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgBasketItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If (NewCol = GDCol(eGDCol_QtyNextEntry)) Or (NewCol = GDCol(eGDCol_Name)) Then
        If (NewRow >= fgBasketItems.FixedRows) And (NewRow < fgBasketItems.Rows) Then
            If fgBasketItems.MergeRow(NewRow) = False Then
                fgBasketItems.EditCell
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.fgBasketItems_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgBasketItems_BeforeEdit
'' Description: Only allow the user to edit the quantity next entry column
'' Inputs:      Row, Column, Cancel the Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgBasketItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col = GDCol(eGDCol_OnCloseTime) Then
        If Len(fgBasketItems.TextMatrix(Row, GDCol(eGDCol_Period))) = 0 Then
            fgBasketItems.ComboList = ""
            Cancel = True
        Else
            fgBasketItems.ComboList = "..."
        End If
    Else
        fgBasketItems.ComboList = ""
        
        If (Col <> GDCol(eGDCol_QtyNextEntry)) And (Col <> GDCol(eGDCol_Name)) Then
            Cancel = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.fgBasketItems_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgBasketItems_CellButtonClick
'' Description: Handle a cell button click on the grid
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgBasketItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim dClosingTime As Double          ' Closing time
    Dim lMinutesBefore As Long          ' Minutes before
    
    dClosingTime = Val(fgBasketItems.TextMatrix(Row, GDCol(eGDCol_ClosingTime)))
    lMinutesBefore = CLng(Val(fgBasketItems.TextMatrix(Row, GDCol(eGDCol_MinutesBefore))))

    If ChangeOnCloseTime(fgBasketItems.TextMatrix(Row, GDCol(eGDCol_Symbol)), dClosingTime, lMinutesBefore) Then
        fgBasketItems.TextMatrix(Row, GDCol(eGDCol_OnCloseTime)) = Format(dClosingTime - (lMinutesBefore / 1440#), "HH:MM")
        fgBasketItems.TextMatrix(Row, GDCol(eGDCol_ClosingTime)) = Str(dClosingTime)
        fgBasketItems.TextMatrix(Row, GDCol(eGDCol_MinutesBefore)) = Str(lMinutesBefore)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.fgBasketItems_CellButtonClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgBasketItems_MouseDown
'' Description: Handle the user clicking on the grid
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Location of mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgBasketItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    
    With fgBasketItems
        lMouseRow = .MouseRow
        
        If Button = vbRightButton Then
            If ValidGridRow(fgBasketItems, lMouseRow) = True Then
                .Row = lMouseRow
                
                Enable mnuChangeInputs, (.TextMatrix(lMouseRow, GDCol(eGDCol_HasExposed)) <> "0")
                
                PopupMenu mnuPopUp
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.fgBasketItems_MouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgBasketItems_MouseMove
'' Description: Handle the user moving the mouse over the grid
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Location of mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgBasketItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Static strText As String            ' Previous tool tip text
    Dim lMouseRow As Long               ' Row in the grid that the mouse is over
    Dim lMouseCol As Long               ' Column in the grid that the mouse is over
    Dim strNewText As String            ' New tool tip text
    
    With fgBasketItems
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        strNewText = ""
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            If lMouseCol = GDCol(eGDCol_Symbol) Then
                strNewText = .TextMatrix(lMouseRow, GDCol(eGDCol_SymbolError))
            ElseIf lMouseCol = GDCol(eGDCol_QtyNextEntry) Then
                strNewText = .TextMatrix(lMouseRow, GDCol(eGDCol_QuantityError))
            End If
            strNewText = Replace(Replace(strNewText, "||", " "), "|", " ")
        End If
        
        If strNewText <> strText Then
            strText = strNewText
            .ToolTipText = strText
        End If
    End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgBasketItems_ValidateEdit
'' Description: Validate the name the user typed in
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgBasketItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strNewName As String            ' New name the user chose
    Dim strOldName As String            ' Old name from the grid
    Dim lPos As Long                    ' Position of name in the array
    Dim TradeItem As cAutoTradeItem     ' Automated trading item from the grid
    Dim strError As String              ' Error

    If Col = GDCol(eGDCol_Name) Then
        strNewName = fgBasketItems.EditText
        strOldName = fgBasketItems.TextMatrix(Row, Col)
        
        If Len(strNewName) = 0 Then
            Cancel = True
        ElseIf strNewName <> strOldName Then
            Do While NameExists(strNewName)
                strNewName = InfBox("That name already exists.  Please select a unique name", , "+OK|-Cancel", "Automated Trading Item Name", , , , , , strNewName)
                If Len(strNewName) = 0 Then
                    Cancel = True
                    Exit Do
                End If
            Loop
            
            If Len(strNewName) > 0 Then
                fgBasketItems.EditText = strNewName
                
                If m.astrChildNames.BinarySearch(strOldName, lPos) = True Then
                    m.astrChildNames.Remove lPos
                End If
                If m.astrChildNames.BinarySearch(strNewName, lPos) = False Then
                    m.astrChildNames.Add strNewName, lPos
                End If
            End If
        End If
    ElseIf Col = GDCol(eGDCol_QtyNextEntry) Then
        Set TradeItem = fgBasketItems.RowData(Row)
        strError = AutomatedQuantityError(TradeItem, fgBasketItems.EditText, fgBasketItems.TextMatrix(Row, GDCol(eGDCol_SymbolError)))
        If Len(strError) > 0 Then
            InfBox strError, "!", , "Error"
            Cancel = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.fgBasketItems_ValidateEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Caption = "Automated Trading Item"
    Me.Icon = Picture16(ToolbarIcon("ID_Strategies"), , True)
    
    CenterTheForm Me
    
    g.Styler.StyleForm Me

    With cboBarPeriod
        .AddItem "5 Minute"
        .AddItem "10 Minute"
        .AddItem "15 Minute"
        .AddItem "30 Minute"
        .AddItem "60 Minute"
        .AddItem "Daily"
        .AddItem "Weekly"
        .AddItem "Monthly"
        .AddItem "Quarterly"
        .AddItem "Yearly"
    
        If g.FractZen.AllowTSOG Then
            .AddItem "FractZen" '"Auto Breakout"
        End If
    End With
    cboBarPeriod.Text = "Daily"
    
    chkConfirm.Value = vbUnchecked
    
    m.lBasketID = 0&
    m.dBasketLastModified = 0#
    Set m.BasketTradeItems = New cGdTree
    Set m.astrChildNames = New cGdArray
    m.astrChildNames.Create eGDARRAY_Strings

    Set m.Qty = New cPriceEditor
    
    cmdReports.Visible = IsIDE
    
    mnuPopUp.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the 'X', hide the form and cancel unload
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadStrategiesCombo
'' Description: Load the Strategies combo box from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadStrategiesCombo()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = mSysNav.LoadStrategiesRecordset(True)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            If mSysNav.IncludeStrategiesFromRecordset(rs) Then
                cboStrategies.AddItem rs!SystemName
                cboStrategies.ItemData(cboStrategies.NewIndex) = rs!SystemNumber
            End If
            rs.MoveNext
        Loop
        
        If cboStrategies.ListCount > 0 Then
            cboStrategies.ListIndex = 0
        End If
    End If

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmAutoTradeItem.LoadStrategiesCombo", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBasketsCombo
'' Description: Load the strategy basket combo box from the database
'' Inputs:      Strategy Basket ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadBasketsCombo(ByVal lStrategyBasketID As Long)
On Error GoTo ErrSection:

    Dim Baskets As cStrategyBaskets ' Collection of strategy baskets
    Dim Basket As cStrategyBasket   ' Strategy basket object
    Dim lIndex As Long              ' Index into a for loop
    Dim lMaxUnits As Long           ' Max units allowed for the basket
    
    Set Baskets = New cStrategyBaskets
    Baskets.LoadDb True, True
    
    For lIndex = 1 To Baskets.Count
        Set Basket = Baskets(lIndex)
        
        If (Basket.ID = lStrategyBasketID) Or (Basket.HasFilter = False) Then
            lMaxUnits = mSysNav.MaxContractsForEnablement(Basket.RequiredModule, Basket.IsGuru)
            If ((lMaxUnits > 0) And (Basket.IsGuru = False)) Or (Basket.ID = m.TradeItem.StrategyBasketID) Then
                cboBaskets.AddItem Basket.Name
                cboBaskets.ItemData(cboBaskets.NewIndex) = Basket.ID
            End If
        End If
    Next lIndex
        
    If cboBaskets.ListCount > 0 Then
        cboBaskets.ListIndex = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.LoadBasketsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeSymbol
'' Description: Change the symbol using the symbol selector
'' Inputs:      Optional starting character
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeSymbol(Optional ByVal strText As String = "")
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Return from the symbol selector
    Dim dSessionEnd As Double           ' Session end time
    Dim strPreviousSymbol As String     ' Previous symbol
    
    strPreviousSymbol = txtSymbol.Text
    
    MoveFocus cmdLookup
    If Len(strText) = 0 Then
        Set astrSymbols = frmSymbolSelector.ShowMe(txtSymbol.Text, False, , "Select Symbol to Trade")
    Else
        Set astrSymbols = frmSymbolSelector.ShowMe(strText, False, , "Select Symbol to Trade", , False)
    End If
    If astrSymbols.Size > 0 Then
        'txtSymbol.Text = RollSymbolForDate(astrSymbols(0), Date)
        UpdateSymbol astrSymbols(0)
    End If
    MoveFocus txtSymbol
    
    ' Does this need to happen here since it happens in UpdateSymbol???
    'InitQuantityEditor
    
    m.dOnCloseTime = DefaultOnCloseTime(txtSymbol.Text, dSessionEnd, m.lMinutesBefore)
    lblOnCloseTime.Caption = DateFormat(m.dOnCloseTime, NO_DATE, HH_MM)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.ChangeSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and size controls appropriately as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lSpace As Long                  ' Spacing between controls
    Dim lDiff As Long                   ' Difference between Height and ScaleHeight

    lSpace = 165
    lDiff = Height - ScaleHeight

    If optBasket.Value = True Then
        fraData.Visible = False
        fgBasketItems.Visible = True
        
        With fgBasketItems
            .Move .Left, fraData.Top, fraData.Width, fraData.Height * 2
        End With
        With fraOptions
            .Move .Left, fgBasketItems.Top + fgBasketItems.Height + lSpace
        End With
    Else
        fraData.Visible = True
        fgBasketItems.Visible = False
        
        With fraOptions
            .Move .Left, fraData.Top + fraData.Height + lSpace
        End With
    End If
    
    With fraButtons
        .Move (ScaleWidth / 2) - (.Width / 2), fraOptions.Top + fraOptions.Height + lSpace
    End With
    
    Move Left, Top, Width, fraButtons.Top + fraButtons.Height + lSpace + lDiff
    
    ExtendCustomColumn

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeInputs_Click
'' Description: Handle the user selecting the Change Inputs menu item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeInputs_Click()
On Error GoTo ErrSection:

    tmrMenu.Interval = 10
    tmrMenu.Tag = "ChangeInputs" & ";" & Str(fgBasketItems.Row)
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.mnuChangeInputs_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optBasket_Click
'' Description: Change the view based on the chosen option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optBasket_Click()
On Error GoTo ErrSection:

    If Visible Then
        ChangeView
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.optBasket_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optStrategy_Click
'' Description: Change the view based on the chosen option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optStrategy_Click()
On Error GoTo ErrSection:

    If Visible Then
        ChangeView
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.optStrategy_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMenu_Timer
'' Description: Handle the menu action the user chose
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strAction As String             ' Action to perform
    Dim lRow As Long                    ' Row in the grid

    tmrMenu.Enabled = False
    
    strAction = Parse(tmrMenu.Tag, ";", 1)
    lRow = CLng(Val(Parse(tmrMenu.Tag, ";", 2)))
    tmrMenu.Tag = ""
    
    Select Case UCase(strAction)
        Case "CHANGEINPUTS"
            ChangeInputs lRow
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.tmrMenu_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_GotFocus
'' Description: Upon getting the focus, select all the text in the text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_GotFocus()
On Error GoTo ErrSection:

    m.strName = Trim(txtName.Text)
    SelectAll txtName

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAutoTradeItem.txtName.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_LostFocus
'' Description: Upon losing the focus, see if the user changed names
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_LostFocus()
On Error GoTo ErrSection:

    If Trim(txtName.Text) <> m.strName Then
        m.astrChildNames.Clear
        FixChildNames m.strName
        m.strName = Trim(txtName.Text)
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAutoTradeItem.txtName_LostFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtNextEntry_Change
'' Description: Handle the user changing the quantity of the next entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtNextEntry_Change()
On Error GoTo ErrSection:

    If Visible Then
        ChangeBasketItemQuantities False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.txtNextEntry_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_Click
'' Description: Allow the user to change symbols using the symbol selector
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    ChangeSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.txtSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_GotFocus
'' Description: Upon getting the focus, select all the text in the text box
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
    RaiseError "frmAutoTradeItem.txtSymbol.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_KeyPress
'' Description: Allow the user to change symbols using the symbol selector
'' Inputs:      Ascii Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    ChangeSymbol Chr(KeyAscii)
    KeyAscii = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.txtSymbol.KeyPress", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpForm
'' Description: Set up the controls on the form from the auto trade object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupForm()
On Error GoTo ErrSection:

    With m.TradeItem
        txtName.Text = .Name
        m.strName = .Name
        SelectComboByItemData cboAccounts, .AccountID
    
        If .StrategyBasketID = 0& Then
            optStrategy.Value = True
            SelectComboByItemData cboStrategies, .StrategyID
            txtSymbol.Text = GetSymbol(.SymbolOrSymbolID)
            m.strPreviousSymbol = txtSymbol.Text
            cboBarPeriod.Text = .BarPeriod
            m.dOnCloseTime = .OnCloseTime
            lblOnCloseTime.Caption = DateFormat(m.dOnCloseTime, NO_DATE, HH_MM)
        Else
            optBasket.Value = True
            SelectComboByItemData cboBaskets, .StrategyBasketID
            m.lMaxUnits = mSysNav.MaxUnitsForAutoTrade(m.TradeItem)
        End If
        
        InitQuantityEditor
        
        CheckBoxValue(chkConfirm) = .ConfirmOrders
        m.TimeOut.Price = .ConfirmTimeout
        CheckBoxValue(chkExitAtEndOfDay) = .ExitAtEndOfDay
        
        ChangeView
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.SetUpForm", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeItemFromControls
'' Description: Set the trade item properties from the controls
'' Inputs:      None
'' Returns:     Auto Trade Item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TradeItemFromControls() As cAutoTradeItem
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated trading item to return
    Dim Bars As cGdBars                 ' Bars object
    Dim dOnCloseTimeExch As Double      ' On-Close time in exchange time
    
    Set TradeItem = New cAutoTradeItem
    With TradeItem
        .AutoTradeItemID = m.TradeItem.AutoTradeItemID
        .Name = Trim(txtName.Text)
        .AccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
        .Deleted = False
        
        .QtyNextEntry = m.Qty.Price
        .ConfirmOrders = (chkConfirm.Value = vbChecked)
        .ConfirmTimeout = m.TimeOut.Price
        .ExitAtEndOfDay = CheckBoxValue(chkExitAtEndOfDay)
        
        If optStrategy.Value = True Then
            .StrategyID = cboStrategies.ItemData(cboStrategies.ListIndex)
            .StrategyName = cboStrategies.Text
            .SymbolOrSymbolID = Trim(txtSymbol.Text)
            '.BarPeriod = PeriodStr(cboBarPeriod.Text)
            .BarPeriod = FixPeriod(cboBarPeriod.Text)
        
            Set Bars = New cGdBars
            SetBarProperties Bars, .SymbolOrSymbolID
            
            dOnCloseTimeExch = ConvertTimeZone(Date + m.dOnCloseTime, "", Bars.Prop(eBARS_ExchangeTimeZoneInf))
            dOnCloseTimeExch = dOnCloseTimeExch - Int(dOnCloseTimeExch)
        
            .OnCloseTimeExch = dOnCloseTimeExch
            .MinutesBefore = m.lMinutesBefore
            
            .ParentID = 0&
            .StrategyBasketID = 0&
            .StrategyBasketItemID = 0&
            .StrategyBasketLastModified = 0#
            .Overrides = ""
            .StrategyBasketItemKey = ""
        Else
            .ParentID = -1&
            .StrategyBasketID = m.lBasketID
            .StrategyBasketItemID = 0&
            .StrategyBasketLastModified = m.dBasketLastModified
            .Overrides = ""
            .StrategyBasketItemKey = ""
        End If
    End With
    
    Set TradeItemFromControls = TradeItem

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.TradeItemFromControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the form data to the automated trading item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save()
On Error GoTo ErrSection:

    Set m.TradeItem = TradeItemFromControls

    m.TradeItem.Save
    g.TradingItems.Add m.TradeItem
    SaveBasketItems m.TradeItem
    
    g.bDirtyLibrariesMDB = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.Save", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveBasketItems
'' Description: Save the basket items
'' Inputs:      Parent ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveBasketItems(ByVal ParentItem As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim ChildItem As cAutoTradeItem     ' Child item to be saved
    Dim lMinutesBefore As Long          ' Minutes before value from the grid
    Dim dLocalClose As Double           ' Local Closing time from the grid
    Dim dLocalOnClose As Double         ' Local on-close time
    Dim dExchOnClose As Double          ' Exchange on-close time
    Dim Bars As cGdBars                 ' Bars object
    Dim lCount As Long                  ' Counter variable for the name
    Dim alIds As cGdArray               ' Array of trade item ID's that we have saved
    
    lCount = 0&
    Set alIds = New cGdArray
    alIds.Create eGDARRAY_Longs
    
    With fgBasketItems
        If optBasket.Value = True Then
            For lIndex = .FixedRows To .Rows - 1
                If .MergeRow(lIndex) = False Then
                    lCount = lCount + 1&
                    lMinutesBefore = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_MinutesBefore))))
                    dLocalClose = Val(.TextMatrix(lIndex, GDCol(eGDCol_ClosingTime)))
                    dLocalOnClose = dLocalClose - (lMinutesBefore / 1440#)
                    
                    Set ChildItem = .RowData(lIndex)
                    With ChildItem
                        Set Bars = New cGdBars
                        SetBarProperties Bars, .SymbolOrSymbolID
                        
                        dExchOnClose = ConvertTimeZone(dLocalOnClose, "", Bars.Prop(eBARS_ExchangeTimeZoneInf))
                        
                        .AccountID = ParentItem.AccountID
                        .ConfirmOrders = ParentItem.ConfirmOrders
                        .ConfirmTimeout = ParentItem.ConfirmTimeout
                        .ExitAtEndOfDay = ParentItem.ExitAtEndOfDay
                        .Deleted = False
                        .MinutesBefore = lMinutesBefore
                        '.Name = ParentItem.Name & " #" & Str(lCount)
                        .Name = fgBasketItems.TextMatrix(lIndex, GDCol(eGDCol_Name))
                        .OnCloseTimeExch = dExchOnClose
                        .ParentID = ParentItem.AutoTradeItemID
                        .QtyNextEntry = CLng(Val(fgBasketItems.TextMatrix(lIndex, GDCol(eGDCol_QtyNextEntry))))
                        .StrategyBasketID = ParentItem.StrategyBasketID
                        .StrategyBasketLastModified = ParentItem.StrategyBasketLastModified
                        .Overrides = fgBasketItems.TextMatrix(lIndex, GDCol(eGDCol_Overrides))
                        .StrategyBasketItemKey = fgBasketItems.TextMatrix(lIndex, GDCol(eGDCol_BasketItemKey))
                    End With
                    .RowData(lIndex) = ChildItem
                    
                    ChildItem.Save
                    g.TradingItems.Add ChildItem
                    
                    alIds.Add ChildItem.AutoTradeItemID
                End If
            Next lIndex
        End If
        
        alIds.Sort
        
        For lIndex = m.BasketTradeItems.Count To 1 Step -1
            Set ChildItem = m.BasketTradeItems(lIndex)
            If alIds.BinarySearch(ChildItem.AutoTradeItemID) = False Then
                g.TradingItems.Delete ChildItem.AutoTradeItemID
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.SaveBasketItems"
    
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

    Dim bIsGuru As Boolean              ' Is the auto trade item for a guru basket?
    Dim bEnable As Boolean              ' Enable the controls?

    bIsGuru = m.TradeItem.IsGuru

    Enable lblName, Not m.bActive
    Enable txtName, Not m.bActive
    Enable optStrategy, ((Not m.bActive) And (cboStrategies.ListCount > 0) And (Not bIsGuru))
    Enable cboStrategies, optStrategy.Enabled
    Enable optBasket, ((Not m.bActive) And (cboBaskets.ListCount > 0) And (Not bIsGuru))
    Enable cboBaskets, optBasket.Enabled
    Enable lblAccount, Not m.bActive
    Enable cboAccounts, Not m.bActive
    Enable lblSymbol, Not m.bActive
    Enable txtSymbol, Not m.bActive
    Enable cmdLookup, Not m.bActive
    Enable lblBarPeriod, Not m.bActive
    Enable cboBarPeriod, Not m.bActive
    
    ' 04/04/2014 DAJ: Don't allow user to change the quantity of the next entry if
    ' they are pyramiding and in a position...
    bEnable = (m.TradeItem.Pyramid = False) Or (m.TradeItem.CurrentPosition = 0)
    Enable lblQtyNextEntry, bEnable
    Enable txtNextEntry, bEnable
    Enable sbQty, bEnable
    Enable cmdReset, bEnable

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowableSymbol
'' Description: Determine if the symbol is allowable
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllowableSymbol(ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strType As String               ' Type of continuous contract
    
    If InStr(strSymbol, "-0") = 0 Then
        bReturn = True
    Else
        strType = Parse(strSymbol, "-", 2)
        bReturn = ((strType = "055") Or (strType = "065") Or (strType = "057") Or (strType = "067"))
    End If
    
    AllowableSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.AllowableSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateSymbol
'' Description: Update the symbol text box
'' Inputs:      New Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateSymbol(ByVal strNewSymbol As String)
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Currently selected account ID
    Dim nBroker As eTT_AccountType      ' Broker for the given account ID
    
    lAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
    nBroker = g.Broker.AccountTypeForID(lAccountID)
    
    If ValidAutomatedSymbol(lAccountID, strNewSymbol, "Automated Trading Item", "Auto Trade Item") Then
        ' DAJ 08/22/2014: We are confident enough in the rolling of continuous contracts with
        ' automated trading items that we are willing to remove the need for the flag file now...
        'If ((g.Broker.IsLiveAccount(nBroker) = False) Or (FileExist(AddSlash(App.Path) & "AutoContLive.FLG")) And AllowableSymbol(strNewSymbol)) Then
        If AllowableSymbol(strNewSymbol) Then
            txtSymbol.Text = strNewSymbol
        Else
            txtSymbol.Text = RollSymbolForDate(strNewSymbol, Date)
        End If
    End If
    
    If txtSymbol.Text <> m.strPreviousSymbol Then
        If SecurityType(txtSymbol.Text) <> SecurityType(m.strPreviousSymbol) Then
            InitQuantityEditor
        End If
        
        m.strPreviousSymbol = txtSymbol.Text
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.UpdateSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeView
'' Description: Change the view
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeView()
On Error GoTo ErrSection:

    If optBasket.Value = True Then
        cboBaskets.Move cboStrategies.Left, cboStrategies.Top, cboStrategies.Width
        cboStrategies.Visible = False
        cboBaskets.Visible = True
        lblQtyNextEntry.Caption = "&Units for the Basket:"
        LoadBasketItemsGrid
        InitQuantityEditor
        cmdReset.Visible = True
    Else
        cboStrategies.Visible = True
        cboBaskets.Visible = False
        lblQtyNextEntry.Caption = "&Quantity for Next Entry:"
        cmdReset.Visible = False
        
        m.lNumBasketItems = 1
    End If
    Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.ChangeView"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitBasketItemsGrid
'' Description: Initialize the basket items grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitBasketItemsGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgBasketItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .HighLight = flexHighlightNever
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Strategy)) = "Strategy"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Period)) = "Period"
        .TextMatrix(0, GDCol(eGDCol_Multiplier)) = "Basket Qty"
        .TextMatrix(0, GDCol(eGDCol_QtyNextEntry)) = "Qty Next Entry"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_OnCloseTime)) = "On Close"
        .TextMatrix(0, GDCol(eGDCol_ClosingTime)) = "Closing"
        .TextMatrix(0, GDCol(eGDCol_MinutesBefore)) = "Minutes Before"
        .TextMatrix(0, GDCol(eGDCol_SortKey)) = "Sort Key"
        .TextMatrix(0, GDCol(eGDCol_BasketItemID)) = "Basket Item ID"
        .TextMatrix(0, GDCol(eGDCol_Overrides)) = "Overrides"
        .TextMatrix(0, GDCol(eGDCol_BasketItemKey)) = "Basket Item Key"
        .TextMatrix(0, GDCol(eGDCol_OutlineLevel)) = "Outline Level"
        .TextMatrix(0, GDCol(eGDCol_Units)) = "Units"
        .TextMatrix(0, GDCol(eGDCol_SymbolError)) = "Symbol Error"
        .TextMatrix(0, GDCol(eGDCol_QuantityError)) = "Quantity Error"
        .TextMatrix(0, GDCol(eGDCol_HasExposed)) = "Has Exposed"
        .TextMatrix(0, GDCol(eGDCol_BasketItemSymbol)) = "Item Symbol"
        .TextMatrix(0, GDCol(eGDCol_BasketItemMult)) = "Item Mult"
        
        .ColHidden(GDCol(eGDCol_ClosingTime)) = True
        .ColHidden(GDCol(eGDCol_MinutesBefore)) = True
        .ColHidden(GDCol(eGDCol_SortKey)) = True
        .ColHidden(GDCol(eGDCol_BasketItemID)) = True
        .ColHidden(GDCol(eGDCol_Overrides)) = True
        .ColHidden(GDCol(eGDCol_BasketItemKey)) = True
        .ColHidden(GDCol(eGDCol_OutlineLevel)) = True
        .ColHidden(GDCol(eGDCol_Units)) = True
        .ColHidden(GDCol(eGDCol_SymbolError)) = True
        .ColHidden(GDCol(eGDCol_QuantityError)) = True
        .ColHidden(GDCol(eGDCol_HasExposed)) = True
        .ColHidden(GDCol(eGDCol_BasketItemSymbol)) = True
        .ColHidden(GDCol(eGDCol_BasketItemMult)) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAutoTradeItem.InitBasketItemsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBasketItemsGrid
'' Description: Load the basket items grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadBasketItemsGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lBasketID As Long               ' Basket ID selected
    Dim Basket As cStrategyBasket       ' Strategy basket selected
    Dim lIndex As Long                  ' Index into a for loop
    Dim lNumBasketItems As Long         ' Number of items in the basket
    
    If (optBasket.Value = True) And (cboBaskets.ListIndex >= 0) Then
        lBasketID = cboBaskets.ItemData(cboBaskets.ListIndex)
        
        If lBasketID <> m.lBasketID Then
            m.lBasketID = lBasketID
            
            Set Basket = New cStrategyBasket
            If Basket.LoadDb(lBasketID, , True) Then
                m.lMaxUnits = mSysNav.MaxContractsForEnablement(Basket.RequiredModule, Basket.IsGuru)
                m.dBasketLastModified = Basket.LastModified
                m.astrChildNames.Clear
                
                With fgBasketItems
                    nRedraw = .Redraw
                    .Redraw = flexRDNone
                    
                    .Rows = .FixedRows
                    For lIndex = 1 To Basket.Items.Count
                        BasketItemToGrid Basket.Items(lIndex)
                    Next lIndex
                    
                    FixChildNames
                    
                    ToggleOutlineLevel False
                    .Col = GDCol(eGDCol_SortKey)
                    .Sort = flexSortGenericAscending
                    ToggleOutlineLevel True
                    
                    If ChangeForex = True Then
                        ChangeBasketItemQuantities True
                    End If
                    
                    .AutoSize 0, .Cols - 1, False, 75
                    ExtendCustomColumn
                    
                    ValidateBasketSymbols
                    
                    CountBasketItems
                    
                    .Redraw = nRedraw
                End With
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.LoadBasketItemsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketItemToGrid
'' Description: Add the given strategy basket item to the grid
'' Inputs:      Strategy basket item, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BasketItemToGrid(ByVal Item As cStrategyBasketItem, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim dClosingTime As Double          ' Closing time for the symbol
    Dim TradeItem As cAutoTradeItem     ' Automated trading item for the basket item
    Dim lPos As Long                    ' Position of the name in the array
    Dim DeletedItem As cAutoTradeItem   ' Deleted automated trading item

    With fgBasketItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        If (Len(Item.SymbolGroupID) > 0) And (Len(Item.Symbol) > 0) Then
            .RowOutlineLevel(lRow) = 1
        Else
            .RowOutlineLevel(lRow) = 0
        End If
        
        If m.TradeItem.AutoTradeItemID = 0& Then
            Set TradeItem = AutoTradeItemFromBasketItem(Item, Nothing)
        ElseIf (m.TradeItem.StrategyBasketID = BasketID) And (m.BasketTradeItems.Exists("ID=" & Str(Item.ID)) = True) Then
            Set TradeItem = AutoTradeItemFromBasketItem(Item, m.BasketTradeItems("ID=" & Str(Item.ID)))
        Else
            Set DeletedItem = g.TradingItems.GetDeletedItem(m.TradeItem.AutoTradeItemID, Item.Key)
            Set TradeItem = AutoTradeItemFromBasketItem(Item, DeletedItem)
        End If
        
        .RowData(lRow) = TradeItem
        If Len(Item.Symbol) > 0 Then
            dClosingTime = TradeItem.OnCloseTime + (TradeItem.MinutesBefore / 1440#)
            
            If Len(TradeItem.Name) > 0 Then
                If m.astrChildNames.BinarySearch(TradeItem.Name, lPos) = False Then
                    m.astrChildNames.Add TradeItem.Name, lPos
                End If
            End If
            
            .TextMatrix(lRow, GDCol(eGDCol_Strategy)) = TradeItem.StrategyName
            .TextMatrix(lRow, GDCol(eGDCol_Symbol)) = TradeItem.Symbol
            .TextMatrix(lRow, GDCol(eGDCol_Period)) = TradeItem.BarPeriod
            .TextMatrix(lRow, GDCol(eGDCol_Multiplier)) = Str(TradeItem.StrategyBasketItemMult)
            .TextMatrix(lRow, GDCol(eGDCol_QtyNextEntry)) = Str(TradeItem.QtyNextEntry)
            .TextMatrix(lRow, GDCol(eGDCol_Name)) = TradeItem.Name
            .TextMatrix(lRow, GDCol(eGDCol_OnCloseTime)) = Format(TradeItem.OnCloseTime, "HH:MM")
            .TextMatrix(lRow, GDCol(eGDCol_ClosingTime)) = Str(dClosingTime)
            .TextMatrix(lRow, GDCol(eGDCol_MinutesBefore)) = Str(TradeItem.MinutesBefore)
            .TextMatrix(lRow, GDCol(eGDCol_Units)) = Str(m.Qty.Price)
            
            .Cell(flexcpFontBold, lRow, GDCol(eGDCol_QtyNextEntry)) = True
            
            ' 05/15/2013 DAJ:  If I don't do this here, I get a 'l' looking character after
            ' the text in the cell through the IDE.  I don't seem to get that if I comment
            ' out the bold line above or if I do this...
            .Cell(flexcpPicture, lRow, GDCol(eGDCol_QtyNextEntry)) = Nothing
            
            .MergeRow(lRow) = False
        Else
            .Cell(flexcpText, lRow, 0, lRow, .Cols - 1) = Item.StrategyName
            .Cell(flexcpFontBold, lRow, GDCol(eGDCol_QtyNextEntry)) = False
            .MergeRow(lRow) = True
        End If
        
        .TextMatrix(lRow, GDCol(eGDCol_SortKey)) = Pad(Item.StrategyName, 50, "L") & Pad(Item.SymbolGroupName, 50, "L") & Pad(Item.Symbol, 50, "L")
        .TextMatrix(lRow, GDCol(eGDCol_BasketItemID)) = Str(Item.ID)
        .TextMatrix(lRow, GDCol(eGDCol_Overrides)) = FixOverrides(Item.Overrides, TradeItem.Overrides)
        .TextMatrix(lRow, GDCol(eGDCol_BasketItemKey)) = Item.Key
        .TextMatrix(lRow, GDCol(eGDCol_OutlineLevel)) = Str(.RowOutlineLevel(lRow))
        .TextMatrix(lRow, GDCol(eGDCol_HasExposed)) = Str(Item.HasExposedParameters)
        .TextMatrix(lRow, GDCol(eGDCol_BasketItemSymbol)) = Item.Symbol
        .TextMatrix(lRow, GDCol(eGDCol_BasketItemMult)) = Str(Item.ContractMultiplier)
        
        .IsSubtotal(lRow) = True
                
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.BasketItemToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoTradeItemFromBasketItem
'' Description: Build an automated trading item from a strategy basket item
'' Inputs:      Strategy Basket Item, New Item?
'' Returns:     Automated Trading Item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AutoTradeItemFromBasketItem(ByVal basketItem As cStrategyBasketItem, ByVal AutoTradeItem As cAutoTradeItem) As cAutoTradeItem
On Error GoTo ErrSection:

    Dim lMinutesBefore As Long          ' Minutes before session end for on-close time
    Dim dClosingTime As Double          ' Closing time for the symbol
    Dim dLocalOnClose As Double         ' Local on-close time
    Dim dExchOnClose As Double          ' Exchange on-close time
    Dim Bars As cGdBars                 ' Bars object
    Dim lCounter As Long                ' Counter for the name
    Dim lMinLotSize As Long             ' Minimum lot size
    
    If AutoTradeItem Is Nothing Then
        Set AutoTradeItem = New cAutoTradeItem
        With AutoTradeItem
            Set Bars = New cGdBars
            SetBarProperties Bars, basketItem.Symbol
            
            lMinutesBefore = m.lMinutesBefore
            dLocalOnClose = DefaultOnCloseTime(basketItem.Symbol, dClosingTime, lMinutesBefore)
            dExchOnClose = ConvertTimeZone(dLocalOnClose, "", Bars.Prop(eBARS_ExchangeTimeZoneInf))
            dExchOnClose = dExchOnClose - Int(dExchOnClose)
            
            .AccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
            .ConfirmOrders = CheckBoxValue(chkConfirm)
            .ConfirmTimeout = CLng(Val(txtTimeout.Text))
            .MinutesBefore = lMinutesBefore
            .OnCloseTimeExch = dExchOnClose
            
            lMinLotSize = g.Broker.MinimumLotSize(.AccountID, basketItem.SymbolOrSymbolID)
            .QtyNextEntry = basketItem.ContractMultiplier * m.Qty.Price * lMinLotSize
        End With
    End If
    
    With AutoTradeItem
        If InStr(.Name, "|") > 0 Then
            .Name = Replace(.Name, "|", "")
            If NameExists(.Name) = True Then
                .Name = ""
            End If
        End If
        
        .StrategyBasketID = basketItem.StrategyBasketID
        .StrategyBasketItemID = basketItem.ID
        .StrategyBasketItemMult = basketItem.ContractMultiplier
        .StrategyID = basketItem.StrategyID
        .StrategyName = basketItem.StrategyName
        .SymbolOrSymbolID = basketItem.SymbolOrSymbolID
        .BarPeriod = basketItem.Period
        
        If m.TradeItem.QtyNextEntry > 0& Then
            If m.Qty.Price <> m.TradeItem.QtyNextEntry Then
                If .QtyNextEntry = basketItem.ContractMultiplier * m.TradeItem.QtyNextEntry Then
                    .QtyNextEntry = basketItem.ContractMultiplier * m.Qty.Price
                End If
            End If
        End If
    End With
    
    Set AutoTradeItemFromBasketItem = AutoTradeItem

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.AutoTradeItemFromGrid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LocalSessionStart
'' Description: Determine the session start time for a symbol in local time
'' Inputs:      Symbol
'' Returns:     Session Start
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LocalSessionStart(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary bars structure

    SetBarProperties Bars, strSymbol
    LocalSessionStart = ConvertTimeZone(Date + (Bars.Prop(eBARS_DefaultStartTime) / 1440#), Bars.Prop(eBARS_ExchangeTimeZoneInf), "")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.LocalSessionStart"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LocalSessionEnd
'' Description: Determine the session end time for a symbol in local time
'' Inputs:      Symbol
'' Returns:     Session End
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LocalSessionEnd(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary bars structure

    SetBarProperties Bars, strSymbol
    LocalSessionEnd = ConvertTimeZone(Date + (Bars.Prop(eBARS_DefaultEndTime) / 1440#), Bars.Prop(eBARS_ExchangeTimeZoneInf), "")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.LocalSessionEnd"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DefaultOnCloseTime
'' Description: Determine the on-close time for a symbol
'' Inputs:      Symbol, Closing Time (out), Minutes Before (in/out)
'' Returns:     On Close Time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DefaultOnCloseTime(ByVal strSymbol As String, dLocalClosingTime As Double, lMinutesBefore As Long) As Double
On Error GoTo ErrSection:

    dLocalClosingTime = LocalSessionEnd(strSymbol)
    If lMinutesBefore = 0& Then
        ' DAJ 06/26/2014: Tim is thinking that three minutes before the close is too
        ' soon to do the on-close check, so we are changing the default to one minute before...
        ' DAJ 01/05/2015: When you have a bunch of automated trading items that close at the
        ' same time, they sometimes take the whole minute, so back this off to two minutes...
        lMinutesBefore = 2& ' 1& ' 3&
    End If
    DefaultOnCloseTime = LocalSessionEnd(strSymbol) - (lMinutesBefore / 1440#)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.DefaultOnCloseTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeOnCloseTime
'' Description: Allow the user to change the on-close time for a symbol
'' Inputs:      Symbol, Closing Time (in/out), Minutes Before (in/out)
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChangeOnCloseTime(ByVal strSymbol As String, dClosingTime As Double, lMinutesBefore As Long) As Boolean
On Error GoTo ErrSection:

    ChangeOnCloseTime = frmOnCloseTime.ShowMe(dClosingTime, lMinutesBefore, LocalSessionEnd(strSymbol), LocalSessionStart(strSymbol))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.ChangeOnCloseTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBasketTradeItems
'' Description: Load up the trading items for the basket if applicable
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadBasketTradeItems()
On Error GoTo ErrSection:

    Dim TradeItem As cAutoTradeItem     ' Automated trading item for the basket item
    Dim lIndex As Long                  ' Index into a for loop

    Set m.BasketTradeItems = New cGdTree
    If m.TradeItem.ParentID = -1& Then
        For lIndex = 1 To g.TradingItems.Count
            If g.TradingItems(lIndex).ParentID = m.TradeItem.AutoTradeItemID Then
                Set TradeItem = g.TradingItems(lIndex)
                m.BasketTradeItems.Add TradeItem, "ID=" & Str(TradeItem.StrategyBasketItemID)
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.LoadBasketTradeItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeBasketItemQuantities
'' Description: Change the basket item quantities based on the basket units
'' Inputs:      Force Quantity Change?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeBasketItemQuantities(ByVal bForceChange As Boolean)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim Item As cAutoTradeItem          ' Automated trading item
    Dim lNewMult As Long                ' New multiplier
    Dim lOldMult As Long                ' Old multiplier
    Dim lMinLotSize As Long             ' Minimum lot size
    
    With fgBasketItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If .MergeRow(lIndex) = False Then
                Set Item = .RowData(lIndex)
                
                lMinLotSize = g.Broker.MinimumLotSize(AccountID, Item.SymbolOrSymbolID)
                lNewMult = Item.StrategyBasketItemMult * m.Qty.Price * lMinLotSize
                lOldMult = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_Multiplier)))) * CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_Units))))
                
                If (bForceChange = True) Or (CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_QtyNextEntry)))) = lOldMult * lMinLotSize) Then
                    .TextMatrix(lIndex, GDCol(eGDCol_QtyNextEntry)) = Str(lNewMult)
                End If
                .TextMatrix(lIndex, GDCol(eGDCol_Units)) = Str(m.Qty.Price)
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.ChangeBasketItemQuantities"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate custom extended column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn(Optional ByVal lResizeCol As Long = -1)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lTotal As Long                  ' Total width
    Dim lDiff As Long                   ' Difference in column width

    With fgBasketItems
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If lResizeCol >= kExtendedCol Then
            .Redraw = flexRDNone
            lDiff = .ColWidth(lResizeCol) - m.nPrevColWidth
            For lIndex = lResizeCol + 1 To .Cols - 1
                If Not .ColHidden(lIndex) Then
                    .ColWidth(lIndex) = .ColWidth(lIndex) - lDiff
                    Exit For
                End If
            Next
            m.nPrevColWidth = 0
        End If
        
        ' size the custom extended column in order to fill the client width
        .ColHidden(kExtendedCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        lTotal = 0
        For lIndex = 0 To .Cols - 1
            If Not .ColHidden(lIndex) Then
                lTotal = lTotal + .ColWidth(lIndex)
            End If
        Next
        lTotal = .ClientWidth - lTotal
        If lTotal > 0 Then .ColWidth(kExtendedCol) = lTotal
        .ColHidden(kExtendedCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.ExtendCustomColumn"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitQuantityEditor
'' Description: Initialize the quantity editor according to the selected
''              account and symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitQuantityEditor()
On Error GoTo ErrSection:

    Dim dPrice As Double                ' Price to set the editor to

    If optBasket Then
        If m.TradeItem.QtyNextEntry = 0 Then
            If m.lMaxUnits = 999999 Then
                dPrice = 1#
            Else
                dPrice = m.lMaxUnits
            End If
        Else
            dPrice = m.TradeItem.QtyNextEntry
        End If
        
        m.Qty.Init sbQty, txtNextEntry, Nothing, dPrice, 0&, m.lMaxUnits, , True, 1&
        ChangeBasketItemQuantities False
    Else
        If m.TradeItem.QtyNextEntry = 0 Then
            dPrice = kNullData
        ElseIf SecurityType(m.TradeItem.Symbol) <> SecurityType(txtSymbol.Text) Then
            dPrice = kNullData
        Else
            dPrice = m.TradeItem.QtyNextEntry
        End If
        
        g.Broker.InitQuantityEditor m.Qty, sbQty, txtNextEntry, AccountID, txtSymbol.Text, dPrice, True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.InitQuantityEditor"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketItemsFromGrid
'' Description: Get the basket items from the grid
'' Inputs:      None
'' Returns:     Basket Items
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BasketItemsFromGrid() As cGdTree
On Error GoTo ErrSection:

    Dim BasketItems As cGdTree          ' Collection of strategy basket items
    Dim lIndex As Long                  ' Index into a for loop
    Dim ChildItem As cAutoTradeItem     ' Child item to be saved
    
    Set BasketItems = New cGdTree
    If optBasket.Value = True Then
        With fgBasketItems
            For lIndex = .FixedRows To .Rows - 1
                If .MergeRow(lIndex) = False Then
                    Set ChildItem = .RowData(lIndex)
                    BasketItems.Add ChildItem, Str(ChildItem.AutoTradeItemID)
                End If
            Next lIndex
        End With
    End If
    
    Set BasketItemsFromGrid = BasketItems

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.BasketItemsFromGrid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleOutlineLevel
'' Description: Turn the outline level on or off
'' Inputs:      On or Off?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ToggleOutlineLevel(ByVal bOutlineLevelOn As Boolean)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    
    With fgBasketItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If bOutlineLevelOn Then
                .RowOutlineLevel(lIndex) = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_OutlineLevel))))
                .IsSubtotal(lIndex) = True
            Else
                .RowOutlineLevel(lIndex) = 0
                .IsSubtotal(lIndex) = False
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.ToggleOutlineLevel"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixChildNames
'' Description: Fix any empty child names in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixChildNames(Optional ByVal strOldParentName As String = "")
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings in the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim strParentName As String         ' Name of the parent
    Dim strOldName As String            ' Old name from the grid
    Dim strNewName As String            ' New name to use for the child
    Dim lCounter As Long                ' Counter for the name
    Dim lPos As Long                    ' Position of the name in the array
    Dim bFix As Boolean                 ' Fix the name?
    
    With fgBasketItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If Len(Trim(txtName.Text)) > 0 Then
            strParentName = Trim(txtName.Text)
        Else
            strParentName = "AT"
        End If
        If Len(strOldParentName) = 0 Then
            strOldParentName = strParentName
        End If
        
        lCounter = 1&
        
        For lIndex = .FixedRows To .Rows - 1
            strOldName = .TextMatrix(lIndex, GDCol(eGDCol_Name))
            
            bFix = False
            If Len(strOldName) = 0 Then
                bFix = True
            ElseIf strParentName <> strOldParentName Then
                bFix = IsDefaultChildName(strOldName, strOldParentName)
            End If
            
            If bFix = True Then
                strNewName = strParentName & " #" & Str(lCounter)
                Do While m.astrChildNames.BinarySearch(strNewName, lPos) = True
                    lCounter = lCounter + 1&
                    strNewName = strParentName & " #" & Str(lCounter)
                Loop
                
                .TextMatrix(lIndex, GDCol(eGDCol_Name)) = strNewName
                m.astrChildNames.Add strNewName, lPos
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.FixChildNames"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsDefaultChildName
'' Description: Is the given child name a default name?
'' Inputs:      Child Name, Parent Name
'' Returns:     True if default, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsDefaultChildName(ByVal strChildName As String, ByVal strParentName As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lParentLength As Long           ' Length of the parent
    
    bReturn = False
    lParentLength = Len(strParentName) + 2
    
    If Len(strChildName) > lParentLength Then
        If Left(strChildName, lParentLength) = strParentName & " #" Then
            If Val(Parse(strChildName, "#", 2)) > 0 Then
                bReturn = True
            End If
        End If
    End If
    
    IsDefaultChildName = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.IsDefaultChildName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NameExists
'' Description: Does the given name exist?
'' Inputs:      Name to Check
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NameExists(ByVal strNameToCheck As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Trim(txtName.Text) = strNameToCheck Then
        bReturn = True
    ElseIf m.astrChildNames.BinarySearch(strNameToCheck) = True Then
        bReturn = True
    Else
        bReturn = g.TradingItems.NameExistsInDb(strNameToCheck)
    End If
    
    NameExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.NameExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidateBasketSymbols
'' Description: Validate the symbols in a basket
'' Inputs:      Has invalid quantities ( out )
'' Returns:     Number of valid symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateBasketSymbols(Optional bHasInvalidQuantities As Boolean) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim bValid As Boolean               ' Is the row valid?
    Dim lIndex As Long                  ' Index into a for loop
    Dim strError As String              ' Symbol error
    Dim TradeItem As cAutoTradeItem     ' Automated trading item from the grid
    
    lReturn = 0&
    bHasInvalidQuantities = False
    
    With fgBasketItems
        For lIndex = .FixedRows To .Rows - 1
            If .MergeRow(lIndex) = False Then
                Set TradeItem = .RowData(lIndex)
                                
                strError = AutomatedSymbolError(AccountID, TradeItem.Symbol, "Automated Trading", True)
                If Len(strError) > 0 Then
                    .TextMatrix(lIndex, GDCol(eGDCol_SymbolError)) = strError
                    .TextMatrix(lIndex, GDCol(eGDCol_QtyNextEntry)) = "0"
                    .Cell(flexcpForeColor, lIndex, GDCol(eGDCol_Symbol)) = vbRed
                    bValid = False
                Else
                    .TextMatrix(lIndex, GDCol(eGDCol_SymbolError)) = ""
                    .Cell(flexcpForeColor, lIndex, GDCol(eGDCol_Symbol)) = .Cell(flexcpForeColor, lIndex, GDCol(eGDCol_SymbolError))
                    bValid = True
                End If
                
                strError = AutomatedQuantityError(TradeItem, .TextMatrix(lIndex, GDCol(eGDCol_QtyNextEntry)), strError)
                If Len(strError) > 0 Then
                    .TextMatrix(lIndex, GDCol(eGDCol_QuantityError)) = strError
                    .Cell(flexcpForeColor, lIndex, GDCol(eGDCol_QtyNextEntry)) = vbRed
                    bValid = False
                    bHasInvalidQuantities = True
                Else
                    .TextMatrix(lIndex, GDCol(eGDCol_QuantityError)) = ""
                    .Cell(flexcpForeColor, lIndex, GDCol(eGDCol_QtyNextEntry)) = .Cell(flexcpForeColor, lIndex, GDCol(eGDCol_QuantityError))
                End If
                
                If bValid Then
                    lReturn = lReturn + 1&
                End If
            End If
        Next lIndex
    End With
    
    ValidateBasketSymbols = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.ValidateBasketSymbols"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeForex
'' Description: Change Forex symbols and quantities if applicable
'' Inputs:      None
'' Returns:     True if changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChangeForex() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lRow As Long                    ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Automated trading item from the grid
    Dim strBasketItemSymbol As String   ' Basket item symbol
    Dim strTradeSymbol As String        ' Symbol to use for trading
    Dim lBasketItemMult As Long         ' Basket item multiplier
    Dim lMult As Long                   ' Multiplier to use for the item
    Dim nBroker As eTT_AccountType      ' Account type for the selected account
    Dim lContractSize As Long           ' Contract size for the symbol/account
    
    bReturn = False
    With fgBasketItems
        For lRow = .FixedRows To .Rows - 1
            If .MergeRow(lRow) = False Then
                strBasketItemSymbol = .TextMatrix(lRow, GDCol(eGDCol_BasketItemSymbol))
                If (IsForex(strBasketItemSymbol) = True) And (InStr(strBasketItemSymbol, "@") = 0) Then
                    Set TradeItem = .RowData(lRow)
                    lBasketItemMult = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_BasketItemMult))))
                    lContractSize = 1&
                    
                    nBroker = g.Broker.AccountTypeForID(AccountID)
                    If g.Broker.IsOecBroker(nBroker) = True Then
                        strTradeSymbol = strBasketItemSymbol & "@OEC"
                        lContractSize = g.Broker.ContractSize(strTradeSymbol, AccountID)
                        If lContractSize <= 0 Then
                            lContractSize = 1&
                        End If
                    ElseIf g.Broker.IsIbBroker(nBroker) = True Then
                        strTradeSymbol = strBasketItemSymbol & "@IB"
                    ElseIf g.Broker.IsCurrenexBroker(nBroker) = True Then
                        strTradeSymbol = strBasketItemSymbol & "@CNX"
                    Else 'If g.Broker.IsLiveAccount(nBroker) = False Then
                        strTradeSymbol = strBasketItemSymbol
                    End If
                    
                    lMult = lBasketItemMult / lContractSize
                    
                    If (strTradeSymbol <> TradeItem.Symbol) Or (lMult <> TradeItem.StrategyBasketItemMult) Then
                        TradeItem.SymbolOrSymbolID = strTradeSymbol
                        TradeItem.StrategyBasketItemMult = lMult
                        
                        .RowData(lRow) = TradeItem
                        
                        .TextMatrix(lRow, GDCol(eGDCol_Symbol)) = strTradeSymbol
                        
                        bReturn = True
                    End If
                End If
            End If
        Next lRow
    End With
    
    ChangeForex = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.ChangeForex"
    
End Function

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
    RaiseError "frmAutoTradeItem.FixPeriod"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeInputs
'' Description: Allow the user to change exposed inputs for the given row
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeInputs(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim pt As POINTAPI                  ' Point on the screen
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim lChildRow As Long               ' Child row in the grid
    Dim strOverrides As String          ' Overrides for the automated trading item

    If mFlexGrid.ValidGridRow(fgBasketItems, lRow) = True Then
        If TypeOf fgBasketItems.RowData(lRow) Is cAutoTradeItem Then
            Set TradeItem = fgBasketItems.RowData(lRow)
            
            pt = mFlexGrid.PointFromCell(fgBasketItems, lRow, GDCol(eGDCol_Symbol))
            If frmAutoTradeItemOverrides.ShowMe(pt.X, pt.Y, TradeItem) = True Then
                fgBasketItems.RowData(lRow) = TradeItem
                fgBasketItems.TextMatrix(lRow, GDCol(eGDCol_Overrides)) = TradeItem.Overrides
                
                If fgBasketItems.RowOutlineLevel(lRow) = 0 Then
                    strOverrides = TradeItem.Overrides
                    lChildRow = fgBasketItems.GetNodeRow(lRow, flexNTFirstChild)
                    Do While lChildRow <> -1&
                        Set TradeItem = fgBasketItems.RowData(lChildRow)
                        TradeItem.Overrides = strOverrides
                        fgBasketItems.RowData(lChildRow) = TradeItem
                        fgBasketItems.TextMatrix(lChildRow, GDCol(eGDCol_Overrides)) = TradeItem.Overrides
                        
                        lChildRow = fgBasketItems.GetNodeRow(lChildRow, flexNTNextSibling)
                    Loop
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.ChangeInputs"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixOverrides
'' Description: Fix the automated trading overrides
'' Inputs:      Basket Item Overrides, Automated Trading Item Overrides
'' Returns:     Fixed Automated Trading Item Overrides
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixOverrides(ByVal strBasketItemOverrides As String, ByVal strAutoTradeItemOverrides As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrBasketOverrides As cGdArray ' Overrides for the strategy basket item
    Dim astrAutoTradeOverrides As cGdArray ' Overrides for the automated trading item
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrBasketOverrides = New cGdArray
    astrBasketOverrides.SplitFields strBasketItemOverrides, ","
    astrBasketOverrides.Sort
    
    Set astrAutoTradeOverrides = New cGdArray
    astrAutoTradeOverrides.SplitFields Parse(strAutoTradeItemOverrides, "|", 2), ","
    astrAutoTradeOverrides.Sort
    
    For lIndex = astrAutoTradeOverrides.Size - 1 To 0 Step -1
        If astrBasketOverrides.BinarySearch(Parse(astrAutoTradeOverrides(lIndex), "=", 1) & "=", , eGdSort_MatchUsingSearchStringLength) = False Then
            astrAutoTradeOverrides.Remove lIndex
        End If
    Next lIndex

    If astrAutoTradeOverrides.Size = 0 Then
        strReturn = astrBasketOverrides.JoinFields(",")
    Else
        strReturn = astrBasketOverrides.JoinFields(",") & "|" & astrAutoTradeOverrides.JoinFields(",")
    End If

    FixOverrides = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAutoTradeItem.FixOverrides"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CountBasketItems
'' Description: Count the strategy basket items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CountBasketItems()
On Error GoTo ErrSection:

    Dim lCount As Long                  ' Count of basket items
    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeItem As cAutoTradeItem     ' Automated trading item for the basket item
    Dim lNumRecords As Long             ' Number of records in the pool for the group
    Dim basketItem As cStrategyBasketItem ' Strategy basket item
    Dim alAddWarnRow As cGdArray        ' Rows to add a warning row

    Set alAddWarnRow = New cGdArray
    alAddWarnRow.Create eGDARRAY_Longs

    With fgBasketItems
        lCount = 0
        
        For lIndex = .FixedRows To .Rows - 1
            If .RowOutlineLevel(lIndex) = 0 Then
                If TypeOf .RowData(lIndex) Is cAutoTradeItem Then
                    Set TradeItem = .RowData(lIndex)
                    If Len(TradeItem.Symbol) > 0 Then
                        lCount = lCount + 1
                    ElseIf .GetNodeRow(lIndex, flexNTFirstChild) = -1 Then
                        Set basketItem = New cStrategyBasketItem
                        If basketItem.LoadDb(TradeItem.StrategyBasketItemID) Then
                            lNumRecords = g.SymbolPool.NumberRecordsForID(basketItem.SymbolGroupID)
                            lCount = lCount + lNumRecords
                            
                            If lNumRecords > kSN_BASKETLIMIT Then
                                alAddWarnRow.Add lIndex
                            End If
                        End If
                    End If
                End If
            Else
                lCount = lCount + 1
            End If
        Next lIndex
        
        m.lNumBasketItems = lCount
        
        For lIndex = alAddWarnRow.Size - 1 To 0 Step -1
            .Rows = .Rows + 1
            .MergeRow(.Rows - 1) = True
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "There are more than " & Str(kSN_BASKETLIMIT) & " symbols in this symbol group"
            .RowPosition(.Rows - 1) = alAddWarnRow(lIndex) + 1
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAutoTradeItem.CountBasketItems"
    
End Sub

