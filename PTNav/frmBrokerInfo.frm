VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmBrokerInfo 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPbo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6060
      Picture         =   "frmBrokerInfo.frx":0000
      ScaleHeight     =   210
      ScaleWidth      =   1830
      TabIndex        =   1
      Top             =   5400
      Width           =   1830
   End
   Begin VB.PictureBox picRithmic 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      Picture         =   "frmBrokerInfo.frx":050E
      ScaleHeight     =   315
      ScaleWidth      =   2025
      TabIndex        =   2
      Top             =   5340
      Width           =   2025
   End
   Begin VB.Timer tmrMenu 
      Left            =   7020
      Top             =   4860
   End
   Begin VB.Timer tmrUpdate 
      Left            =   7500
      Top             =   4860
   End
   Begin HexUniControls.ctlUniFrameWL fraConnection 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
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
      Caption         =   "frmBrokerInfo.frx":079D
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBrokerInfo.frx":07C9
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerInfo.frx":07E9
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP lblConnection 
         Height          =   195
         Left            =   1560
         Top             =   60
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
         Caption         =   "frmBrokerInfo.frx":0805
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBrokerInfo.frx":0831
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":0851
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblConnectionStatus 
         Height          =   195
         Left            =   0
         Top             =   60
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
         Caption         =   "frmBrokerInfo.frx":086D
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBrokerInfo.frx":08B1
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":08D1
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraPositions 
      Height          =   1035
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   6375
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
      Caption         =   "frmBrokerInfo.frx":08ED
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBrokerInfo.frx":0919
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerInfo.frx":0939
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkShowFlat 
         Height          =   195
         Left            =   960
         TabIndex        =   14
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
         Caption         =   "frmBrokerInfo.frx":0955
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBrokerInfo.frx":099D
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":09BD
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fgPositions 
         Height          =   795
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   6375
         _cx             =   11245
         _cy             =   1402
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
      Begin HexUniControls.ctlUniLabelXP lblPositions 
         Height          =   255
         Left            =   0
         Top             =   0
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
         Caption         =   "frmBrokerInfo.frx":09D9
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBrokerInfo.frx":0A0D
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":0A2D
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFills 
      Height          =   1035
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   6375
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
      Caption         =   "frmBrokerInfo.frx":0A49
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBrokerInfo.frx":0A75
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerInfo.frx":0A95
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fgFills 
         Height          =   795
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   6375
         _cx             =   11245
         _cy             =   1402
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
      Begin HexUniControls.ctlUniLabelXP lblFills 
         Height          =   255
         Left            =   0
         Top             =   0
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
         Caption         =   "frmBrokerInfo.frx":0AB1
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBrokerInfo.frx":0ADD
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":0AFD
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraClosedOrders 
      Height          =   1035
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   6375
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
      Caption         =   "frmBrokerInfo.frx":0B19
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBrokerInfo.frx":0B45
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerInfo.frx":0B65
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fgClosedOrders 
         Height          =   795
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   6375
         _cx             =   11245
         _cy             =   1402
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
      Begin HexUniControls.ctlUniLabelXP lblClosedOrders 
         Height          =   255
         Left            =   0
         Top             =   0
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
         Caption         =   "frmBrokerInfo.frx":0B81
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBrokerInfo.frx":0BBD
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":0BDD
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraOpenOrders 
      Height          =   1035
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   6375
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
      Caption         =   "frmBrokerInfo.frx":0BF9
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBrokerInfo.frx":0C25
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerInfo.frx":0C45
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fgOpenOrders 
         Height          =   795
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   6375
         _cx             =   11245
         _cy             =   1402
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
      Begin HexUniControls.ctlUniLabelXP lblOpenOrders 
         Height          =   255
         Left            =   0
         Top             =   0
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
         Caption         =   "frmBrokerInfo.frx":0C61
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBrokerInfo.frx":0C99
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":0CB9
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2415
      Left            =   6660
      TabIndex        =   16
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
      Caption         =   "frmBrokerInfo.frx":0CD5
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBrokerInfo.frx":0D01
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerInfo.frx":0D21
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDisconnect 
         Height          =   495
         Left            =   0
         TabIndex        =   4
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
         Caption         =   "frmBrokerInfo.frx":0D3D
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBrokerInfo.frx":0D73
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":0D93
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdConnect 
         Height          =   495
         Left            =   0
         TabIndex        =   7
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
         Caption         =   "frmBrokerInfo.frx":0DAF
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBrokerInfo.frx":0DDF
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":0DFF
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPrint 
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   1860
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
         Caption         =   "frmBrokerInfo.frx":0E1B
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBrokerInfo.frx":0E47
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":0E67
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRefresh 
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   1200
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
         Caption         =   "frmBrokerInfo.frx":0E83
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBrokerInfo.frx":0EB3
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBrokerInfo.frx":0ED3
         RightToLeft     =   0   'False
      End
   End
   Begin VB.Menu mnuOpenOrders 
      Caption         =   "Open Orders"
      Begin VB.Menu mnuEditOrder 
         Caption         =   "Edit Order"
      End
      Begin VB.Menu mnuCancelOrder 
         Caption         =   "Cancel Order"
      End
      Begin VB.Menu mnuParkOrder 
         Caption         =   "Park Order"
      End
      Begin VB.Menu mnuSubmitOrder 
         Caption         =   "Submit Order"
      End
      Begin VB.Menu mnuSubmitAll 
         Caption         =   "Submit All Parked Orders"
      End
      Begin VB.Menu mnuOrderSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrderHistory 
         Caption         =   "Order History"
      End
      Begin VB.Menu mnuOrderJournal 
         Caption         =   "New Journal for Order"
      End
      Begin VB.Menu mnuViewJournals 
         Caption         =   "View Journals"
      End
   End
   Begin VB.Menu mnuPositions 
      Caption         =   "Positions"
      Begin VB.Menu mnuFlatten 
         Caption         =   "Flatten"
      End
      Begin VB.Menu mnuReverse 
         Caption         =   "Reverse"
      End
   End
End
Attribute VB_Name = "frmBrokerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmBrokerInfo.frm
'' Description: Show information that comes directly from the online broker
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/06/2009   DAJ         Display "Mismatch" for position if in a mismatch
'' 06/02/2009   DAJ         Fix error when position in a future option
'' 08/21/2009   DAJ         Set UserCancel flag on CancelOrder call
'' 09/01/2009   DAJ         Use new Parked order status
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 03/11/2010   DAJ         Use global TradingItems collection
'' 09/24/2010   DAJ         Added some artwork to be shown for Rithmic
'' 10/05/2010   DAJ         Changed the Rithmic image
'' 10/26/2010   DAJ         Changed interval for the update timer
'' 10/27/2010   DAJ         More mods to the Rithmic image
'' 09/23/2011   DAJ         Show date journals form instead of old journals form
'' 10/04/2011   DAJ         Call the ShowJournals function instead of calling the form direct
'' 06/24/2013   DAJ         Timer Logging
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    nBroker As eTT_AccountType          ' Broker for this instance of the form
    BrokerInfo As cBrokerInfo           ' Broker information structure
End Type
Private m As mPrivate

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim bRithmicBroker As Boolean       ' Is this a Rithmic broker?

    m.nBroker = nBroker

    tmrUpdate.Interval = 100
    
    Caption = "Activity for " & g.Broker.BrokerName(m.nBroker)
    
    InitOpenOrdersGrid
    InitClosedOrdersGrid
    InitFillsGrid
    InitPositionsGrid
    
    bRithmicBroker = g.Broker.IsRithmicBroker(nBroker)
    picRithmic.Visible = bRithmicBroker
    picPbo.Visible = bRithmicBroker

    tmrUpdate.Enabled = True
    ShowForm Me, eForm_Nonmodal, frmMain, , ALT_GRID_ROW_COLOR
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Send the grid to the Print Preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    PrintMe = frmPrintPreview.ShowMe("TNV BrokerInfo", Me, , , , 0.75, 0.75)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerInfo.PrintMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the Print Preview
'' Inputs:      Arguments into the Print Preview
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim strDefaults As String
    Dim lRow As Long
    Dim lCol As Long
    Dim strText As String
    

    strDefaults = "GRP:CONT067.GRP;7;14;3;3;3;3;-1;2;22;156;GC-067;TQ-067"
    strDefaults = GetIniFileProperty("Defaults", strDefaults, "Defaults", AddSlash(App.Path) & "CotRpt.INI")

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader 10
        
        .TextAlign = taCenterMiddle
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .Text = Caption
        .Font.Size = 12
        .FontUnderline = False
        .Font.Bold = False
        .TextAlign = taLeftMiddle
        
        .Text = vbLf & "Open Orders:" & vbLf
        fgOpenOrders.ExtendLastCol = False
        If frmPrintPreview.GoingToFile Then
            With fgOpenOrders
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If Not .ColHidden(lCol) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        Else
            .RenderControl = fgOpenOrders.hWnd
        End If
        fgOpenOrders.ExtendLastCol = True
        
        .Text = vbLf & "Closed Orders:" & vbLf
        fgClosedOrders.ExtendLastCol = False
        If frmPrintPreview.GoingToFile Then
            With fgClosedOrders
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If Not .ColHidden(lCol) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        Else
            .RenderControl = fgClosedOrders.hWnd
        End If
        fgClosedOrders.ExtendLastCol = True
        
        .Text = vbLf & "Fills:" & vbLf
        fgFills.ExtendLastCol = False
        If frmPrintPreview.GoingToFile Then
            With fgFills
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If Not .ColHidden(lCol) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        Else
            .RenderControl = fgFills.hWnd
        End If
        fgFills.ExtendLastCol = True
        
        .Text = vbLf & "Positions:" & vbLf
        fgPositions.ExtendLastCol = False
        If frmPrintPreview.GoingToFile Then
            With fgPositions
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If Not .ColHidden(lCol) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        Else
            .RenderControl = fgPositions.hWnd
        End If
        fgPositions.ExtendLastCol = True
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowFlat_Click
'' Description: Toggle the filter on the positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowFlat_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterPositions
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.chkShowFlat_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdConnect_Click
'' Description: Make an attempt to connect to the appropriate broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdConnect_Click()
On Error GoTo ErrSection:

    g.Broker.Connect m.nBroker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.cmdConnect_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDisconnect_Click
'' Description: Make an attempt to disconnect from the appropriate broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDisconnect_Click()
On Error GoTo ErrSection:

    g.Broker.Disconnect m.nBroker, "User disconnected from Broker Activity form"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.cmdDisconnect_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: Print out the information on the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.cmdPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRefresh_Click
'' Description: Refresh the appropriate thing based on the broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRefresh_Click()
On Error GoTo ErrSection:

    g.Broker.Refresh m.nBroker, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.cmdRefresh_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgClosedOrders_BeforeMouseDown
'' Description: If the user right clicks, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgClosedOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    
    With fgClosedOrders
        lMouseRow = .MouseRow
        
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            If Button = vbRightButton Then
                mnuEditOrder.Enabled = False
                mnuCancelOrder.Enabled = False
                mnuParkOrder.Enabled = False
                mnuSubmitOrder.Enabled = False
                mnuSubmitAll.Enabled = False
                mnuOrderHistory.Enabled = True
                mnuOrderJournal.Enabled = True
                
                mnuOpenOrders.Tag = "Closed"
                PopupMenu mnuOpenOrders
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.fgClosedOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOpenOrders_BeforeMouseDown
'' Description: Show the popup menu if necessary
'' Inputs:      Mouse Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOpenOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row in the grid that the user clicked on
    Dim Order As cPtOrder               ' Order object

    With fgOpenOrders
        lMouseRow = .MouseRow
        
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            If Button = vbRightButton Then
                Set Order = fgOpenOrders.RowData(fgOpenOrders.Row)
                
                mnuEditOrder.Enabled = True
                mnuCancelOrder.Enabled = True
                mnuParkOrder.Enabled = IsOpenOrder(Order.Status, False)
                mnuSubmitOrder.Enabled = ((Order.Status = eTT_OrderStatus_Open) Or (Order.Status = eTT_OrderStatus_Parked))
                mnuSubmitAll.Enabled = HasParkedOrders
                mnuOrderHistory.Enabled = True
                mnuOrderJournal.Enabled = True
                
                mnuOpenOrders.Tag = "Open"
                PopupMenu mnuOpenOrders
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.fgOpenOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_BeforeMouseDown
'' Description: Show the popup menu if necessary
'' Inputs:      Mouse Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row in the grid that the user clicked on
    Dim AcctPos As cAccountPosition     ' Account position object
    Dim bMismatch As Boolean            ' Is the symbol currently in a position mismatch?

    If Button = vbRightButton Then
        lMouseRow = fgPositions.MouseRow
        If (lMouseRow >= fgPositions.FixedRows) And (lMouseRow < fgPositions.Rows) Then
            fgPositions.Row = lMouseRow
            fgPositions.RowSel = lMouseRow
            
            Set AcctPos = fgPositions.RowData(lMouseRow)
            bMismatch = (g.Broker.PositionMatch(AcctPos.AccountID, AcctPos.SymbolOrSymbolID) = False)
            
            mnuFlatten.Enabled = ((AcctPos.CurrentPositionSnapshot <> 0&) Or (bMismatch = True))
            mnuReverse.Enabled = (AcctPos.AutoTradeItemID = 0&) And ((AcctPos.CurrentPositionSnapshot <> 0&) Or (bMismatch = True))
            
            PopupMenu mnuPositions
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.fgPositions_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement string from the ini file
    
    strPlacement = GetIniFileProperty("frmBrokerInfo", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    
    g.Styler.StyleForm Me
    
    ' Hide the popup menus until we need them...
    mnuOpenOrders.Visible = False
    mnuPositions.Visible = False
    
    Set m.BrokerInfo = New cBrokerInfo
    
    lblConnection.Caption = ""

    chkShowFlat.Value = GetIniFileProperty("ShowFlat", vbUnchecked, "ActivityView", g.strIniFile)
    
    tmrMenu.Interval = 100
    tmrMenu.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.Form_Load"
    
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

    Dim lFrameHeight As Long            ' Height of each grid
    Dim lFrameWidth As Long             ' Width of each grid
    Dim lMinScaleHeight As Long         ' Minimum scale height for the form
    
    If picRithmic.Visible Then
        lMinScaleHeight = (1035 * 4) + fraConnection.Height + picRithmic.Height + (60 * 7)
    Else
        lMinScaleHeight = (1035 * 4) + fraConnection.Height + (60 * 6)
    End If

    If Not LimitFormSize(Me, 7950, lMinScaleHeight) Then
        With fraButtons
            .Move ScaleWidth - .Width - 60, 60
        End With
        
        If picRithmic.Visible Then
            lFrameHeight = (ScaleHeight - fraConnection.Height - picRithmic.Height - (60 * 7)) / 4
        Else
            lFrameHeight = (ScaleHeight - fraConnection.Height - (60 * 6)) / 4
        End If
        lFrameWidth = ScaleWidth - fraButtons.Width - 180
        
        With fraConnection
            .Move 60, 60
        End With
        
        With picRithmic
            .Move 60, ScaleHeight - .Height - 60
        End With
        
        With picPbo
            .Move ScaleWidth - .Width - 60, ScaleHeight - .Height - 60
        End With
        
        With fraOpenOrders
            .Move 60, fraConnection.Top + fraConnection.Height + 60, lFrameWidth, lFrameHeight
        End With
        
        With fgOpenOrders
            .Move 0, .Top, fraOpenOrders.Width, fraOpenOrders.Height - lblOpenOrders.Height
        End With
        
        With fraClosedOrders
            .Move 60, fraOpenOrders.Top + fraOpenOrders.Height + 60, lFrameWidth, lFrameHeight
        End With
    
        With fgClosedOrders
            .Move 0, .Top, fraClosedOrders.Width, fraClosedOrders.Height - lblClosedOrders.Height
        End With
        
        With fraFills
            .Move 60, fraClosedOrders.Top + fraClosedOrders.Height + 60, lFrameWidth, lFrameHeight
        End With
    
        With fgFills
            .Move 0, .Top, fraFills.Width, fraFills.Height - lblFills.Height
        End With
        
        With fraPositions
            .Move 60, fraFills.Top + fraFills.Height + 60, lFrameWidth, lFrameHeight
        End With
    
        With fgPositions
            .Move 0, .Top, fraPositions.Width, fraPositions.Height - lblPositions.Height
        End With
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save settings and clean up when form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    tmrUpdate.Enabled = False
    tmrMenu.Enabled = False
    
    Set m.BrokerInfo = Nothing
    
    SetIniFileProperty "frmBrokerInfo", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "ShowFlat", chkShowFlat.Value, "ActivityView", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCancelOrder_Click
'' Description: Attempt to Cancel the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCancelOrder_Click()
On Error GoTo ErrSection:

    CancelOrderFromGrid fgOpenOrders, "Activity View", True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuCancelOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditOrder_Click
'' Description: Allow the user to edit the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditOrder_Click()
On Error GoTo ErrSection:

    EditOrderFromGrid fgOpenOrders, "Activity View"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuEditOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuFlatten_Click
'' Description: Attempt to flatten the user for the selected symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuFlatten_Click()
On Error GoTo ErrSection:

    FlattenPositionFromGrid fgPositions, "Activity View"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuFlatten_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrderHistory_Click
'' Description: Allow the user to view the history for an order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrderHistory_Click()
On Error GoTo ErrSection:

    If mnuOpenOrders.Tag = "Open" Then
        With fgOpenOrders
            If .Row >= .FixedRows And .Row < .Rows Then
                frmOrderHistory.ShowMe .RowData(.Row)
            End If
        End With
    Else
        With fgClosedOrders
            If .Row >= .FixedRows And .Row < .Rows Then
                frmOrderHistory.ShowMe .RowData(.Row)
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuOrderHistory_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrderJournal_Click
'' Description: Allow the user to view the journal for an order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrderJournal_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "JOURNAL"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuOrderJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuParkOrder_Click
'' Description: Allow the user to park the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuParkOrder_Click()
On Error GoTo ErrSection:

    ParkOrderFromGrid fgOpenOrders, "Activity View"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuParkOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuReverse_Click
'' Description: Attempt to Reverse the user for the selected symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuReverse_Click()
On Error GoTo ErrSection:

    ReversePositionFromGrid fgPositions, "Activity View"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuReverse_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSubmitAll_Click
'' Description: Allow the user to submit all parked orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSubmitAll_Click()
On Error GoTo ErrSection:

    SubmitAllOrdersFromGrid fgOpenOrders, "Activity View"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuSubmitAll_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSubmitOrder_Click
'' Description: Allow the user to submit the selected order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSubmitOrder_Click()
On Error GoTo ErrSection:

    SubmitOrderFromGrid fgOpenOrders, "Activity View"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuSubmitOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuViewJournals_Click
'' Description: Allow the user to view their journals
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuViewJournals_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "JOURNALS"
    tmrMenu.Enabled = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.mnuViewJournals_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMenu_Timer
'' Description: Perform the appropriate menu item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Tag of the timer
    
    TimerStart "frmBrokerInfo.tmrMenu"
    strTag = tmrMenu.Tag
    tmrMenu.Tag = ""
    tmrMenu.Enabled = False
    
    Select Case UCase(strTag)
        Case "JOURNAL"
            If mnuOpenOrders.Tag = "Open" Then
                With fgOpenOrders
                    If (.Row >= .FixedRows) And (.Row < .Rows) Then
                        If TypeOf .RowData(.Row) Is cPtOrder Then
                            g.TnJournal.ShowOrderJournal .RowData(.Row)
                        End If
                    End If
                End With
            Else
                With fgClosedOrders
                    If (.Row >= .FixedRows) And (.Row < .Rows) Then
                        If TypeOf .RowData(.Row) Is cPtOrder Then
                            g.TnJournal.ShowOrderJournal .RowData(.Row)
                        End If
                    End If
                End With
            End If
                
        Case "JOURNALS"
            g.TnJournal.ShowJournals
            
    End Select
    TimerEnd "frmBrokerInfo.tmrMenu", tmrMenu.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.tmrMenu_Timer"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrUpdate_Timer
'' Description: Update the form if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrUpdate_Timer()
On Error GoTo ErrSection:

    Dim dLastChanged As Double          ' Last time broker information changed

    TimerStart "frmBrokerInfo.tmrUpdate"
    UpdateConnectionStatus m.nBroker, g.Broker.ConnectionStatusForBroker(m.nBroker)
    
    dLastChanged = g.Broker.LastChanged(m.nBroker)
    If (dLastChanged > m.BrokerInfo.LastChanged) Then
        If (g.Broker.Refreshing(m.nBroker) = False) Then
            Set m.BrokerInfo = g.Broker.BrokerInfo(m.nBroker).MakeCopy
            UpdateGridsFromBrokerInfo
        End If
    ElseIf (dLastChanged = -1#) Then
        fgOpenOrders.Rows = fgOpenOrders.FixedRows
        fgClosedOrders.Rows = fgClosedOrders.FixedRows
        fgFills.Rows = fgFills.FixedRows
        fgPositions.Rows = fgPositions.FixedRows
    End If
    TimerEnd "frmBrokerInfo.tmrUpdate", tmrUpdate.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.tmrUpdate_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitOpenOrdersGrid
'' Description: Initialize the Open Orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitOpenOrdersGrid()
On Error GoTo ErrSection:

    With fgOpenOrders
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        
        .Cols = 16
        .TextMatrix(0, 0) = "Date"
        .TextMatrix(0, 1) = "Account"
        .TextMatrix(0, 2) = "B/S"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Symbol"
        .TextMatrix(0, 5) = "Limit"
        .TextMatrix(0, 6) = "Stop"
        .TextMatrix(0, 7) = "Status"
        .TextMatrix(0, 8) = "Order ID"
        .TextMatrix(0, 9) = "Original ID"
        .TextMatrix(0, 10) = "Original Qty"
        .TextMatrix(0, 11) = "Remaining Qty"
        .TextMatrix(0, 12) = "Genesis ID"
        .TextMatrix(0, 13) = "At ID"
        .TextMatrix(0, 14) = "Link"
        
        .ColFormat(0) = DateFormat("Format", MM_DD_YY, HH_MM_SS, AMPM_UPPER)
        
        .ColAlignment(0) = flexAlignCenterTop
        .ColAlignment(1) = flexAlignLeftTop
        .ColAlignment(5) = flexAlignRightTop
        .ColAlignment(6) = flexAlignRightTop
        .ColAlignment(8) = flexAlignLeftTop
        .ColAlignment(9) = flexAlignLeftTop
        .ColAlignment(13) = flexAlignLeftTop
        
        .ColHidden(10) = True
        .ColHidden(12) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.InitOpenOrdersGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitClosedOrdersGrid
'' Description: Initialize the Closed Orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitClosedOrdersGrid()
On Error GoTo ErrSection:

    With fgClosedOrders
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        
        .Cols = 16
        .TextMatrix(0, 0) = "Date"
        .TextMatrix(0, 1) = "Account"
        .TextMatrix(0, 2) = "B/S"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Symbol"
        .TextMatrix(0, 5) = "Limit"
        .TextMatrix(0, 6) = "Stop"
        .TextMatrix(0, 7) = "Status"
        .TextMatrix(0, 8) = "Order ID"
        .TextMatrix(0, 9) = "Original ID"
        .TextMatrix(0, 10) = "Original Qty"
        .TextMatrix(0, 11) = "Remaining Qty"
        .TextMatrix(0, 12) = "Genesis ID"
        .TextMatrix(0, 13) = "At ID"
        .TextMatrix(0, 14) = "Link"
        
        .ColFormat(0) = DateFormat("Format", MM_DD_YY, HH_MM_SS, AMPM_UPPER)
        
        .ColAlignment(0) = flexAlignCenterTop
        .ColAlignment(1) = flexAlignLeftTop
        .ColAlignment(5) = flexAlignRightTop
        .ColAlignment(6) = flexAlignRightTop
        .ColAlignment(8) = flexAlignLeftTop
        .ColAlignment(9) = flexAlignLeftTop
        .ColAlignment(13) = flexAlignLeftTop
        
        .ColHidden(10) = True
        .ColHidden(12) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.InitClosedOrdersGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitFillsGrid
'' Description: Initialize the Open Orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFillsGrid()
On Error GoTo ErrSection:

    With fgFills
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        
        .Cols = 10
        .TextMatrix(0, 0) = "Time"
        .TextMatrix(0, 1) = "Account"
        .TextMatrix(0, 2) = "B/S"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Symbol"
        .TextMatrix(0, 5) = "Price"
        .TextMatrix(0, 6) = "Fill ID"
        .TextMatrix(0, 7) = "Order ID"
        .TextMatrix(0, 8) = "At ID"
                        
        .ColFormat(0) = DateFormat("Format", MM_DD_YY, HH_MM_SS, AMPM_UPPER)
        
        .ColAlignment(0) = flexAlignCenterTop
        .ColAlignment(1) = flexAlignLeftTop
        .ColAlignment(5) = flexAlignRightTop
        .ColAlignment(6) = flexAlignLeftTop
        .ColAlignment(7) = flexAlignLeftTop
        .ColAlignment(8) = flexAlignLeftTop
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.InitFillsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitPositionsGrid
'' Description: Initialize the Open Orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitPositionsGrid()
On Error GoTo ErrSection:

    With fgPositions
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        
        .Cols = 13
        .TextMatrix(0, 0) = "Account"
        .TextMatrix(0, 1) = "Symbol"
        .TextMatrix(0, 2) = "At ID"
        .TextMatrix(0, 3) = "Source"
        .TextMatrix(0, 4) = "Position"
        .TextMatrix(0, 5) = "Average Entry"
        .TextMatrix(0, 6) = "Overnight Pos"
        .TextMatrix(0, 7) = "Buys"
        .TextMatrix(0, 8) = "Sells"
        .TextMatrix(0, 9) = "Total"
        .TextMatrix(0, 10) = "Session Profit"
        .TextMatrix(0, 11) = "Closed Profit"
        
        .ColHidden(2) = True
        
        .ColFormat(10) = "$#,##0.00"
        .ColFormat(11) = "$#,##0.00"
        
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(2) = flexAlignLeftTop
        .ColAlignment(5) = flexAlignRightTop
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.InitPositionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateConnectionStatus
'' Description: Update the connection status for the broker
'' Inputs:      Broker, Connection Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateConnectionStatus(ByVal nBroker As eTT_AccountType, ByVal nStatus As eGDConnectionStatus)
On Error GoTo ErrSection:

    If m.nBroker = nBroker Then
        Select Case nStatus
            Case eGDConnectionStatus_Disconnected
                lblConnection.Caption = "Disconnected"
                cmdConnect.Enabled = True
                cmdDisconnect.Enabled = False
                
            Case eGDConnectionStatus_Disconnecting
                lblConnection.Caption = "Disconnecting"
                cmdConnect.Enabled = False
                cmdDisconnect.Enabled = False
            
            Case eGDConnectionStatus_Connecting
                lblConnection.Caption = "Connecting"
                cmdConnect.Enabled = False
                cmdDisconnect.Enabled = False
            
            Case eGDConnectionStatus_Connected
                lblConnection.Caption = "Connected"
                cmdConnect.Enabled = False
                cmdDisconnect.Enabled = True
        
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.UpdateConnectionStatus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateGridsFromBrokerInfo
'' Description: Update the grids from the new broker information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateGridsFromBrokerInfo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lPos As Long                    ' Position of item in the grid
    Dim astrInfo As New cGdArray        ' Array of information split out from string
    Dim lParent As Long                 ' Parent record
    Dim lLastChild As Long              ' Last child of the parent
    Dim bMismatch As Boolean            ' Is the symbol currently in a position mismatch?
    
    Dim Orders As cPtOrders             ' Collection of orders
    Dim Order As cPtOrder               ' Individual order
    Dim Fills As cPtFills               ' Collection of fills
    Dim Fill As cPtFill                 ' Individual fill
    Dim FillSumms As cAccountPositions  ' Collection of account positions
    Dim FillSum As cAccountPosition     ' Individual account position
    Dim strAccount As String            ' Account number
    Dim Bars As New cGdBars             ' Bars object
    
    fgOpenOrders.Redraw = flexRDNone
    fgClosedOrders.Redraw = flexRDNone
    fgOpenOrders.Rows = fgOpenOrders.FixedRows
    fgClosedOrders.Rows = fgClosedOrders.FixedRows
    
    Set Orders = m.BrokerInfo.Orders
    For lIndex = 1 To Orders.Count
        ' Order Fields: Date, Account, B/S, Qty, Symbol, Limit, Stop, Status, Order ID
        '               Original ID, Original Qty, Remaining Qty, Genesis ID, At ID
        ' Order Record: Broker ID, Genesis ID, Original ID, Account, Symbol, Type, B/S,
        '               Quantity, Limit, Stop, Expiration, Status, Qty Remaining,
        '               Order Date, At ID
        Set Order = Orders(lIndex)
        
        If IsOpenOrder(Order.Status) Then
            With fgOpenOrders
                .Rows = .Rows + 1
                
                .RowData(.Rows - 1) = Order
                .TextMatrix(.Rows - 1, 0) = ConvertBrokerDate(Order.OrderDate, m.nBroker, Order.Symbol, g.bShowInLocalTimeZone)
                .TextMatrix(.Rows - 1, 1) = g.Broker.AccountNumberForID(Order.AccountID)
                If Order.Buy = False Then .TextMatrix(.Rows - 1, 2) = "Sell" Else .TextMatrix(.Rows - 1, 2) = "Buy"
                .TextMatrix(.Rows - 1, 3) = Str(Order.Quantity)
                .TextMatrix(.Rows - 1, 4) = Order.Symbol
                .TextMatrix(.Rows - 1, 5) = Order.LimitPriceString
                .TextMatrix(.Rows - 1, 6) = Order.StopPriceString
                .TextMatrix(.Rows - 1, 7) = OrderStatus(Order.Status)
                .TextMatrix(.Rows - 1, 8) = Order.BrokerID
                .TextMatrix(.Rows - 1, 9) = Order.PreviousBrokerID
                .TextMatrix(.Rows - 1, 10) = ""
                .TextMatrix(.Rows - 1, 11) = Str(Order.RemainingQuantity)
                .TextMatrix(.Rows - 1, 12) = Order.GenesisOrderID
                .TextMatrix(.Rows - 1, 13) = AutoTradeItemNameForID(Order.AutoTradeItemID)
                .TextMatrix(.Rows - 1, 14) = Order.LinkStatus
            End With
        Else
            With fgClosedOrders
                .Rows = .Rows + 1
                
                .RowData(.Rows - 1) = Order
                .TextMatrix(.Rows - 1, 0) = ConvertBrokerDate(Order.OrderDate, m.nBroker, Order.Symbol, g.bShowInLocalTimeZone)
                .TextMatrix(.Rows - 1, 1) = g.Broker.AccountNumberForID(Order.AccountID)
                If Order.Buy = False Then .TextMatrix(.Rows - 1, 2) = "Sell" Else .TextMatrix(.Rows - 1, 2) = "Buy"
                .TextMatrix(.Rows - 1, 3) = Str(Order.Quantity)
                .TextMatrix(.Rows - 1, 4) = Order.Symbol
                .TextMatrix(.Rows - 1, 5) = Order.LimitPriceString
                .TextMatrix(.Rows - 1, 6) = Order.StopPriceString
                .TextMatrix(.Rows - 1, 7) = OrderStatus(Order.Status)
                .TextMatrix(.Rows - 1, 8) = Order.BrokerID
                .TextMatrix(.Rows - 1, 9) = Order.PreviousBrokerID
                .TextMatrix(.Rows - 1, 10) = ""
                .TextMatrix(.Rows - 1, 11) = Str(Order.RemainingQuantity)
                .TextMatrix(.Rows - 1, 12) = Order.GenesisOrderID
                .TextMatrix(.Rows - 1, 13) = AutoTradeItemNameForID(Order.AutoTradeItemID)
                .TextMatrix(.Rows - 1, 14) = Order.LinkStatus
            End With
        End If
    Next lIndex
    
    With fgOpenOrders
        If .Rows > .FixedRows Then
            .Col = 0
            .Sort = flexSortGenericDescending
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With
    
    With fgClosedOrders
        If .Rows > .FixedRows Then
            .Col = 0
            .Sort = flexSortGenericDescending
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With
    
    With fgFills
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        Set Fills = m.BrokerInfo.Fills
        For lIndex = 1 To Fills.Count
            Set Fill = Fills(lIndex)
            
            ' Fill Fields: Time, Account, B/S, Quantity, Symbol, Price, Fill ID, Order ID
            '              At ID
            ' Fill Record: Broker ID, Genesis ID, Fill ID, Account, Symbol, Fill Date,
            '              B/S, Fill Quantity, Fill Price, At ID
            .Rows = .Rows + 1
            
            .RowData(.Rows - 1) = Fill
            .TextMatrix(.Rows - 1, 0) = ConvertBrokerDate(Fill.FillDate, m.nBroker, Fill.Symbol, g.bShowInLocalTimeZone)
            .TextMatrix(.Rows - 1, 1) = g.Broker.AccountNumberForID(Fill.AccountID)
            If Fill.Buy = False Then .TextMatrix(.Rows - 1, 2) = "Sell" Else .TextMatrix(.Rows - 1, 2) = "Buy"
            .TextMatrix(.Rows - 1, 3) = Str(Fill.Quantity)
            .TextMatrix(.Rows - 1, 4) = Fill.Symbol
            .TextMatrix(.Rows - 1, 5) = Fill.PriceString
            .TextMatrix(.Rows - 1, 6) = Fill.BrokerID
            .TextMatrix(.Rows - 1, 7) = Fill.BrokerOrderID
            .TextMatrix(.Rows - 1, 8) = AutoTradeItemNameForID(Fill.AutoTradingItemID)
        Next lIndex
        
        If .Rows > .FixedRows Then
            .Col = 0
            .Sort = flexSortGenericDescending
        End If

        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With
    
    With fgPositions
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 0 To m.BrokerInfo.Positions.Size - 1
            ' Position Fields: Account, Symbol, At ID, At ID, Position, Average Entry,
            '                  Overnight Position, Buys, Sells, Total, Closed
            ' Position Record: Account, Symbol, Position, Average Entry, Overnight Position
            astrInfo.Clear
            astrInfo.SplitFields m.BrokerInfo.Positions(lIndex), vbTab
            
            SetBarProperties Bars, astrInfo(1)
            
            bMismatch = (g.Broker.PositionMatch(astrInfo(0), astrInfo(1)) = False)

            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = astrInfo(0)
            .TextMatrix(.Rows - 1, 1) = astrInfo(1)
            .TextMatrix(.Rows - 1, 2) = "-1"
            .TextMatrix(.Rows - 1, 3) = "Total"
            If bMismatch = True Then
                .TextMatrix(.Rows - 1, 4) = "Mismatch"
                .TextMatrix(.Rows - 1, 5) = ""
                .TextMatrix(.Rows - 1, 6) = ""
            Else
                .TextMatrix(.Rows - 1, 4) = astrInfo(2)
                If Val(astrInfo(3)) = 0 Then
                    .TextMatrix(.Rows - 1, 5) = ""
                Else
                    .TextMatrix(.Rows - 1, 5) = Bars.PriceDisplay(Val(astrInfo(3)))
                End If
                .TextMatrix(.Rows - 1, 6) = astrInfo(4)
            End If
            
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = 1
        Next lIndex
        
        Set FillSumms = m.BrokerInfo.FillSummary
        For lIndex = 1 To FillSumms.Count
            Set FillSum = FillSumms(lIndex)
            strAccount = g.Broker.AccountNumberForID(FillSum.AccountID)
            bMismatch = (g.Broker.PositionMatch(strAccount, FillSum.SymbolOrSymbolID) = False)
            
            ' Fill Summary Record: Account, Symbol, At ID, Buys, Sells, Net, Total, Price Sum,
            '                      Entries, Closed Profit, Average Entry, Initial Fill Price,
            '                      Initial Fill Date, Session Date, Last Traded, Overnight
            lPos = -1&
            lParent = -1&
            For lIndex2 = .FixedRows To .Rows - 1
                If (.TextMatrix(lIndex2, 0) = strAccount) And (.TextMatrix(lIndex2, 1) = FillSum.Symbol) Then
                    lParent = lIndex2
                End If
                
                If (.TextMatrix(lIndex2, 0) = strAccount) And (.TextMatrix(lIndex2, 1) = FillSum.Symbol) And (.TextMatrix(lIndex2, 2) = Str(FillSum.AutoTradeItemID)) Then
                    lPos = lIndex2
                    Exit For
                End If
            Next lIndex2
            
            If lPos = -1& Then
                .Rows = .Rows + 1&
                If lParent <> -1& Then
                    lLastChild = .GetNodeRow(lParent, flexNTLastChild)
                    If lLastChild = -1& Then
                        .RowPosition(.Rows - 1) = lParent + 1
                        lPos = lParent + 1
                    ElseIf (lLastChild <> .Rows - 1) Then
                        .RowPosition(.Rows - 1) = lLastChild + 1
                        lPos = lLastChild + 1
                    Else
                        lPos = .Rows - 1
                    End If
                Else
                    lPos = .Rows - 1&
                End If
                                
                .TextMatrix(lPos, 2) = Str(FillSum.AutoTradeItemID)
                If FillSum.AutoTradeItemID = 0& Then
                    .TextMatrix(lPos, 3) = "Manual"
                ElseIf FillSum.AutoTradeItemID = -1& Then
                    .TextMatrix(lPos, 3) = "Total"
                Else
                    .TextMatrix(lPos, 3) = AutoTradeItemNameForID(FillSum.AutoTradeItemID)
                End If
                
                If bMismatch = True Then
                    .TextMatrix(lPos, 4) = "Mismatch"
                    .TextMatrix(lPos, 5) = ""
                Else
                    .TextMatrix(lPos, 4) = Str(FillSum.CurrentPositionSnapshot)
                    .TextMatrix(lPos, 5) = FillSum.AverageEntrySnapshotString
                End If
                        
                .IsSubtotal(lPos) = True
                If FillSum.AutoTradeItemID = -1& Then
                    .TextMatrix(lPos, 0) = strAccount
                    .TextMatrix(lPos, 1) = FillSum.Symbol
                    .RowOutlineLevel(lPos) = 1
                Else
                    .TextMatrix(lPos, 0) = ""
                    .TextMatrix(lPos, 1) = ""
                    .RowOutlineLevel(lPos) = 2
                End If
            End If
            
            .RowData(lPos) = FillSum
            If bMismatch = True Then
                .TextMatrix(lPos, 6) = ""
                .TextMatrix(lPos, 7) = ""
                .TextMatrix(lPos, 8) = ""
                .TextMatrix(lPos, 9) = ""
                .TextMatrix(lPos, 10) = ""
                .TextMatrix(lPos, 11) = ""
            Else
                .TextMatrix(lPos, 6) = Str(FillSum.CurrentPosition)
                .TextMatrix(lPos, 7) = Str(FillSum.NumBuysSnapshot)
                .TextMatrix(lPos, 8) = Str(FillSum.NumSellsSnapshot)
                .TextMatrix(lPos, 9) = Str(FillSum.NumTotalSnapshot)
                .TextMatrix(lPos, 10) = FillSum.SessionProfitSnapshot
                .TextMatrix(lPos, 11) = FillSum.ClosedProfitSnapshot
            End If
            
            ColorCell fgPositions, lPos, 10
            ColorCell fgPositions, lPos, 11
        Next lIndex
        
        FilterPositions
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.UpdateGridsFromBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorCell
'' Description: Color the given cell red or green depending on sign
'' Inputs:      Grid, Row and Column to color
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorCell(Grid As VSFlexGrid, ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim dValue As Double                ' Value of the cell

    dValue = ValOfText(Grid.TextMatrix(Row, Col))
    
    If dValue < 0 Then
        Grid.Cell(flexcpForeColor, Row, Col) = vbRed
    ElseIf dValue = 0 Then
        Grid.Cell(flexcpForeColor, Row, Col) = vbBlack
    Else
        Grid.Cell(flexcpForeColor, Row, Col) = QBColor(2)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.ColorCell"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasParkedOrders
'' Description: Are there parked SimTrade orders to submit?
'' Inputs:      None
'' Returns:     True if Parked SimTrade orders exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasParkedOrders() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    With fgOpenOrders
        For lIndex = .FixedRows To .Rows - 1
            If .RowData(lIndex).Status = eTT_OrderStatus_Parked Then
                bReturn = True
                Exit For
            End If
        Next lIndex
    End With
    
    HasParkedOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerInfo.HasParkedOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterPositions
'' Description: Filter the account positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterPositions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim bAllFlat As Boolean             ' Are all children flat?
    Dim AcctPos As cAccountPosition     ' Account position object
    Dim bMismatch As Boolean            ' Is the symbol currently in a mismatch?
    
    With fgPositions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If chkShowFlat.Value = vbChecked Then
                If .RowOutlineLevel(lIndex) = 1 Then
                    If .GetNodeRow(lIndex, flexNTFirstChild) = -1& Then
                        .RowHidden(lIndex) = True
                    Else
                        Set AcctPos = .RowData(lIndex)
                        .RowHidden(lIndex) = IsExpiredContract(AcctPos.SymbolOrSymbolID)
                        For lIndex2 = .GetNodeRow(lIndex, flexNTFirstChild) To .GetNodeRow(lIndex, flexNTLastChild)
                            Set AcctPos = .RowData(lIndex2)
                            bMismatch = (g.Broker.PositionMatch(AcctPos.AccountID, AcctPos.SymbolOrSymbolID) = False)
                            
                            If IsExpiredContract(AcctPos.SymbolOrSymbolID) Then
                                .RowHidden(lIndex2) = True
                            ElseIf (g.TradingItems.Exists(Str(AcctPos.AutoTradeItemID)) = False) And (AcctPos.AutoTradeItemID > 0) Then
                                .RowHidden(lIndex2) = (AcctPos.CurrentPositionSnapshot = 0&)
                            Else
                                .RowHidden(lIndex2) = False
                            End If
                        Next lIndex2
                    End If
                End If
            Else
                If .RowOutlineLevel(lIndex) = 1 Then
                    bAllFlat = True
                    If .GetNodeRow(lIndex, flexNTFirstChild) <> -1& Then
                        For lIndex2 = .GetNodeRow(lIndex, flexNTFirstChild) To .GetNodeRow(lIndex, flexNTLastChild)
                            Set AcctPos = .RowData(lIndex2)
                            bMismatch = (g.Broker.PositionMatch(AcctPos.AccountID, AcctPos.SymbolOrSymbolID) = False)
                            
                            If ((AcctPos.CurrentPositionSnapshot <> 0&) Or (bMismatch = True)) Then
                                bAllFlat = False
                                Exit For
                            End If
                        Next lIndex2
                    End If
                    
                    Set AcctPos = .RowData(lIndex)
                    .RowHidden(lIndex) = (bAllFlat = True) Or (IsExpiredContract(AcctPos.SymbolOrSymbolID) = True)
                    If .GetNodeRow(lIndex, flexNTFirstChild) <> -1& Then
                        For lIndex2 = .GetNodeRow(lIndex, flexNTFirstChild) To .GetNodeRow(lIndex, flexNTLastChild)
                            Set AcctPos = .RowData(lIndex)
                            If bAllFlat Then
                                .RowHidden(lIndex2) = True
                            ElseIf IsExpiredContract(AcctPos.SymbolOrSymbolID) Then
                                .RowHidden(lIndex2) = True
                            ElseIf (g.TradingItems.Exists(Str(AcctPos.AutoTradeItemID)) = False) And (AcctPos.AutoTradeItemID > 0) Then
                                .RowHidden(lIndex2) = (AcctPos.CurrentPositionSnapshot = 0&)
                            Else
                                .RowHidden(lIndex2) = False
                            End If
                        Next lIndex2
                    End If
                End If
            End If
        Next lIndex
        
        SetBackColors fgPositions
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerInfo.FilterPositions"
    
End Sub


