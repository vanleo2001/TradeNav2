VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmReports 
   Caption         =   "Reports"
   ClientHeight    =   8535
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15375
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMDIChild"
   ScaleHeight     =   8535
   ScaleWidth      =   15375
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActiveToolBars.SSActiveToolBars Toolbar1 
      Left            =   1380
      Top             =   5940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   9
      DisplayContextMenu=   0   'False
      Tools           =   "frmReports.frx":0000
      ToolBars        =   "frmReports.frx":4DA1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":4F30
            Key             =   "kOpen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":5076
            Key             =   "kClose"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":51D4
            Key             =   "kForm"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":52D6
            Key             =   "kReport"
         EndProperty
      EndProperty
   End
   Begin vsOcx6LibCtl.vsElastic vsElastic1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   15055
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      Appearance      =   0
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   6
      GridCols        =   4
      _GridInfo       =   $"frmReports.frx":5508
      Begin RichTextLib.RichTextBox txtPreview 
         Height          =   1740
         Left            =   90
         TabIndex        =   4
         Top             =   6705
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   3069
         _Version        =   393217
         BackColor       =   -2147483644
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmReports.frx":55A2
      End
      Begin VSFlex7LCtl.VSFlexGrid vsSettings 
         Height          =   2595
         Left            =   90
         TabIndex        =   2
         Top             =   4050
         Width           =   3075
         _cx             =   5424
         _cy             =   4577
         _ConvInfo       =   1
         Appearance      =   2
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
         BackColorFixed  =   13756397
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
         ScrollBars      =   2
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
      Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
         Height          =   3900
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   6879
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   0
         FrontTabForeColor=   -2147483630
         Caption         =   "&Analysis|&Money Mgt"
         Align           =   0
         Appearance      =   1
         CurrTab         =   0
         FirstTab        =   0
         Style           =   4
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         Begin VSFlex7LCtl.VSFlexGrid vsRpts 
            Height          =   3525
            Index           =   0
            Left            =   45
            TabIndex        =   6
            Top             =   330
            Width           =   2985
            _cx             =   5265
            _cy             =   6218
            _ConvInfo       =   1
            Appearance      =   2
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
            BackColorFixed  =   13756397
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
         Begin VSFlex7LCtl.VSFlexGrid vsRpts 
            Height          =   3525
            Index           =   1
            Left            =   3720
            TabIndex        =   7
            Top             =   330
            Width           =   2985
            _cx             =   5265
            _cy             =   6218
            _ConvInfo       =   1
            Appearance      =   2
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
            BackColorFixed  =   13756397
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
      Begin vsOcx6LibCtl.vsIndexTab vsTabs 
         Height          =   8055
         Left            =   3225
         TabIndex        =   1
         Top             =   390
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   14208
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   0
         FrontTabForeColor=   -2147483630
         Caption         =   "&Report|&Report|&Chart|&Chart|Monte Carlo Analysis"
         Align           =   0
         Appearance      =   1
         CurrTab         =   4
         FirstTab        =   0
         Style           =   4
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         Begin HexUniControls.ctlUniFrameWL fraMonteCarlo 
            Height          =   7680
            Left            =   45
            TabIndex        =   37
            Top             =   330
            Width           =   11970
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
            Caption         =   "frmReports.frx":5628
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmReports.frx":5662
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmReports.frx":5682
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraShuffled 
               Height          =   2415
               Left            =   0
               TabIndex        =   42
               Top             =   3300
               Visible         =   0   'False
               Width           =   10275
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
               Caption         =   "frmReports.frx":569E
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmReports.frx":56D4
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":56F4
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtShuffled 
                  Height          =   285
                  Left            =   4680
                  TabIndex        =   46
                  Top             =   660
                  Width           =   495
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmReports.frx":5710
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
                  Tip             =   "frmReports.frx":5734
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":5754
               End
               Begin HexUniControls.ctlUniLabelXP Label3 
                  Height          =   255
                  Left            =   5280
                  Top             =   690
                  Width           =   4515
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
                  Caption         =   "frmReports.frx":5770
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":580C
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":582C
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label2 
                  Height          =   255
                  Left            =   4380
                  Top             =   690
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
                  Caption         =   "frmReports.frx":5848
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":586E
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":588E
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label1 
                  Height          =   255
                  Left            =   4380
                  Top             =   420
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
                  Caption         =   "frmReports.frx":58AA
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":596C
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":598C
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniFrameWL fraDrawdown 
               Height          =   5955
               Left            =   0
               TabIndex        =   49
               Top             =   60
               Visible         =   0   'False
               Width           =   11715
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
               Caption         =   "frmReports.frx":59A8
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmReports.frx":59DE
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":59FE
               RightToLeft     =   0   'False
               Begin VSFlex7LCtl.VSFlexGrid fgDD 
                  Height          =   3015
                  Left            =   6840
                  TabIndex        =   10
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   4620
                  _cx             =   8149
                  _cy             =   5318
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
                  BackColorFixed  =   13756397
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
                  Rows            =   10
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
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
               Begin VSFlex7LCtl.VSFlexGrid fgDDStats 
                  Height          =   3195
                  Left            =   2940
                  TabIndex        =   11
                  Top             =   1800
                  Visible         =   0   'False
                  Width           =   3600
                  _cx             =   6350
                  _cy             =   5636
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
                  BackColorFixed  =   13756397
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
                  FocusRect       =   0
                  HighLight       =   0
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   13
                  Cols            =   2
                  FixedRows       =   0
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
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
               End
               Begin HexUniControls.ctlUniTextBoxXP txtDrawdownSims 
                  Height          =   285
                  Left            =   8160
                  TabIndex        =   12
                  Top             =   360
                  Width           =   795
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmReports.frx":5A1A
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
                  Tip             =   "frmReports.frx":5A44
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":5A64
               End
               Begin HexUniControls.ctlUniTextBoxXP txtDrawdownTrades 
                  Height          =   285
                  Left            =   4620
                  TabIndex        =   50
                  Top             =   660
                  Width           =   675
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   0   'False
                  Locked          =   0   'False
                  Text            =   "frmReports.frx":5A80
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
                  Tip             =   "frmReports.frx":5AA8
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":5AC8
               End
               Begin MSComctlLib.ProgressBar pbRuns 
                  Height          =   300
                  Left            =   3360
                  TabIndex        =   13
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   6495
                  _ExtentX        =   11456
                  _ExtentY        =   529
                  _Version        =   393216
                  Appearance      =   1
               End
               Begin HexUniControls.ctlUniLabelXP Label15 
                  Height          =   2235
                  Left            =   360
                  Top             =   1920
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
                  Caption         =   "frmReports.frx":5AE4
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   2
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":5C5E
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":5C7E
                  RightToLeft     =   0   'False
                  WordWrap        =   -1  'True
               End
               Begin HexUniControls.ctlUniLabelXP Label10 
                  Height          =   255
                  Left            =   9000
                  Top             =   420
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
                  Caption         =   "frmReports.frx":5C9A
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":5CD0
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":5CF0
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label6 
                  Height          =   255
                  Left            =   4380
                  Top             =   420
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
                  Caption         =   "frmReports.frx":5D0C
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":5D90
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":5DB0
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label5 
                  Height          =   255
                  Left            =   4380
                  Top             =   690
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
                  Caption         =   "frmReports.frx":5DCC
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":5DF0
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":5E10
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label4 
                  Height          =   255
                  Left            =   5340
                  Top             =   690
                  Width           =   4575
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
                  Caption         =   "frmReports.frx":5E2C
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":5ECC
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":5EEC
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniFrameWL fraRisk 
               Height          =   6975
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Visible         =   0   'False
               Width           =   10275
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
               Caption         =   "frmReports.frx":5F08
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmReports.frx":5F36
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":5F56
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniComboImageXP cboRiskYears 
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   30
                  Top             =   930
                  Width           =   675
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
                  Tip             =   "frmReports.frx":5F72
                  Sorted          =   0   'False
                  HScroll         =   0   'False
                  RoundedBorders  =   -1  'True
                  IconDim         =   16
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":5F92
                  DropDownOnTextClick=   -1  'True
                  DropDownWidth   =   -1
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniFrameWL fraColors 
                  Height          =   3495
                  Left            =   360
                  TabIndex        =   33
                  Top             =   1560
                  Visible         =   0   'False
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
                  Caption         =   "frmReports.frx":5FAE
                  Enabled         =   -1  'True
                  ForeColor       =   -2147483642
                  BackColor       =   -2147483633
                  Tip             =   "frmReports.frx":5FE0
                  VistaStyle      =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6000
                  RightToLeft     =   0   'False
                  Begin HexUniControls.ctlUniLabelXP Label13 
                     Height          =   495
                     Left            =   180
                     Top             =   300
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
                     Caption         =   "frmReports.frx":601C
                     BackColor       =   -2147483633
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   0
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":6086
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":60A6
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   255
                     Index           =   0
                     Left            =   720
                     Top             =   840
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
                     Caption         =   "frmReports.frx":60C2
                     BackColor       =   65280
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":60EC
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":610C
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   240
                     Index           =   1
                     Left            =   720
                     Top             =   1080
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
                     Caption         =   "frmReports.frx":6128
                     BackColor       =   8454016
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":6154
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":6174
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   240
                     Index           =   2
                     Left            =   720
                     Top             =   1320
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
                     Caption         =   "frmReports.frx":6190
                     BackColor       =   12648384
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":61BC
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":61DC
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   240
                     Index           =   3
                     Left            =   720
                     Top             =   1560
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
                     Caption         =   "frmReports.frx":61F8
                     BackColor       =   12648447
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":6224
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":6244
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   240
                     Index           =   4
                     Left            =   720
                     Top             =   1800
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
                     Caption         =   "frmReports.frx":6260
                     BackColor       =   8454143
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":628C
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":62AC
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   240
                     Index           =   5
                     Left            =   720
                     Top             =   2040
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
                     Caption         =   "frmReports.frx":62C8
                     BackColor       =   65535
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":62F4
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":6314
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   240
                     Index           =   6
                     Left            =   720
                     Top             =   2280
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
                     Caption         =   "frmReports.frx":6330
                     BackColor       =   12640511
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":635C
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":637C
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   240
                     Index           =   7
                     Left            =   720
                     Top             =   2520
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
                     Caption         =   "frmReports.frx":6398
                     BackColor       =   12632319
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":63C4
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":63E4
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   240
                     Index           =   8
                     Left            =   720
                     Top             =   2760
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
                     Caption         =   "frmReports.frx":6400
                     BackColor       =   8421631
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":642C
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":644C
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
                  Begin HexUniControls.ctlUniLabelXP lblColor 
                     Height          =   240
                     Index           =   9
                     Left            =   720
                     Top             =   3000
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
                     Caption         =   "frmReports.frx":6468
                     BackColor       =   255
                     ForeColor       =   -2147483630
                     Alignment       =   2
                     VAlignment      =   0
                     BackStyle       =   1
                     BorderStyle     =   0
                     AutoSize        =   0   'False
                     Tip             =   "frmReports.frx":6492
                     Style           =   0
                     Enabled         =   -1  'True
                     Margin          =   0
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmReports.frx":64B2
                     RightToLeft     =   0   'False
                     WordWrap        =   0   'False
                  End
               End
               Begin HexUniControls.ctlUniTextBoxXP txtRiskTrades 
                  Height          =   285
                  Left            =   9420
                  TabIndex        =   34
                  Top             =   780
                  Visible         =   0   'False
                  Width           =   675
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmReports.frx":64CE
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
                  Tip             =   "frmReports.frx":64F4
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6514
               End
               Begin HexUniControls.ctlUniTextBoxXP txtRiskStart 
                  Height          =   285
                  Left            =   4980
                  TabIndex        =   35
                  Top             =   930
                  Width           =   1095
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmReports.frx":6530
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
                  Tip             =   "frmReports.frx":6560
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6580
               End
               Begin HexUniControls.ctlUniTextBoxXP txtRiskSims 
                  Height          =   285
                  Left            =   4680
                  TabIndex        =   36
                  Top             =   600
                  Width           =   735
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmReports.frx":659C
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
                  Tip             =   "frmReports.frx":65C4
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":65E4
               End
               Begin VSFlex7LCtl.VSFlexGrid fgRisks 
                  Height          =   6255
                  Left            =   3360
                  TabIndex        =   43
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   6000
                  _cx             =   10583
                  _cy             =   11033
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
                  Cols            =   3
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
                  Editable        =   2
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
               Begin HexUniControls.ctlUniLabelXP Label12 
                  Height          =   255
                  Left            =   4380
                  Top             =   960
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
                  Caption         =   "frmReports.frx":6600
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":6632
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6652
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP txtFailures 
                  Height          =   615
                  Left            =   420
                  Top             =   5280
                  Visible         =   0   'False
                  Width           =   2115
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
                  Caption         =   "frmReports.frx":666E
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":6740
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6760
                  RightToLeft     =   0   'False
                  WordWrap        =   -1  'True
               End
               Begin HexUniControls.ctlUniLabelXP lblRiskTrades 
                  Height          =   255
                  Left            =   7920
                  Top             =   960
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
                  Caption         =   "frmReports.frx":677C
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":67A8
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":67C8
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label11 
                  Height          =   255
                  Left            =   5460
                  Top             =   660
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
                  Caption         =   "frmReports.frx":67E4
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":686A
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":688A
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label9 
                  Height          =   255
                  Left            =   6120
                  Top             =   960
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
                  Caption         =   "frmReports.frx":68A6
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":68E2
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6902
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label8 
                  Height          =   255
                  Left            =   4380
                  Top             =   660
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
                  Caption         =   "frmReports.frx":691E
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":6944
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6964
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label7 
                  Height          =   255
                  Left            =   4380
                  Top             =   360
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
                  Caption         =   "frmReports.frx":6980
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":6A2C
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6A4C
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdCompare 
               Height          =   975
               Left            =   720
               TabIndex        =   45
               Top             =   4260
               Visible         =   0   'False
               Width           =   1455
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
               Caption         =   "frmReports.frx":6A68
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmReports.frx":6AC0
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":6AE0
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniFrameWL fraType 
               Height          =   1155
               Left            =   180
               TabIndex        =   38
               Top             =   60
               Width           =   3015
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
               Caption         =   "frmReports.frx":6AFC
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmReports.frx":6B6A
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":6B8A
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniRadioXP optRisk 
                  Height          =   255
                  Left            =   180
                  TabIndex        =   41
                  Top             =   780
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
                  Caption         =   "frmReports.frx":6BA6
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmReports.frx":6BE2
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6C02
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optDrawdown 
                  Height          =   255
                  Left            =   180
                  TabIndex        =   40
                  Top             =   540
                  Width           =   2715
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
                  Caption         =   "frmReports.frx":6C1E
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmReports.frx":6C6E
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6C8E
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optShuffled 
                  Height          =   255
                  Left            =   180
                  TabIndex        =   39
                  Top             =   300
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
                  Caption         =   "frmReports.frx":6CAA
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmReports.frx":6CE8
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6D08
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdMonteCarlo 
               Default         =   -1  'True
               Height          =   615
               Left            =   3360
               TabIndex        =   44
               Top             =   360
               Width           =   855
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
               Caption         =   "frmReports.frx":6D24
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmReports.frx":6D4C
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":6D6C
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraImplements 
            Height          =   7680
            Left            =   -13215
            TabIndex        =   18
            Top             =   330
            Width           =   11970
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
            Caption         =   "frmReports.frx":6D88
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmReports.frx":6DB4
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmReports.frx":6DD4
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraViewOptions 
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   600
               Width           =   3015
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
               Caption         =   "frmReports.frx":6DF0
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmReports.frx":6E1C
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":6E3C
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniRadioXP optMonthly 
                  Height          =   255
                  Left            =   960
                  TabIndex        =   26
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
                  Caption         =   "frmReports.frx":6E58
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmReports.frx":6E88
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6EA8
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optYearly 
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   27
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
                  Caption         =   "frmReports.frx":6EC4
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmReports.frx":6EF2
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6F12
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optTrades 
                  Height          =   255
                  Left            =   0
                  TabIndex        =   25
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
                  Caption         =   "frmReports.frx":6F2E
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   -1  'True
                  Tip             =   "frmReports.frx":6F5C
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":6F7C
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
            Begin vsOcx6LibCtl.vsElastic txtReportName 
               Height          =   555
               Index           =   1
               Left            =   0
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   0
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   979
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   2
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483636
               ForeColor       =   -2147483634
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "Performance Summary"
               Align           =   0
               Appearance      =   2
               AutoSizeChildren=   5
               BorderWidth     =   6
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   1
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   0   'False
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   0
               GridCols        =   0
               _GridInfo       =   ""
            End
            Begin VSFlex7LCtl.VSFlexGrid vsImplements 
               Height          =   3855
               Left            =   0
               TabIndex        =   22
               Top             =   1080
               Width           =   7410
               _cx             =   13070
               _cy             =   6800
               _ConvInfo       =   1
               Appearance      =   2
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
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   1
               HighLight       =   0
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   0
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   ""
               ScrollTrack     =   -1  'True
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   0   'False
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   1
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
               BackColorFrozen =   -2147483643
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
            Begin RichTextLib.RichTextBox rtfMessages 
               Height          =   810
               Left            =   0
               TabIndex        =   23
               Top             =   5040
               Width           =   7395
               _ExtentX        =   13044
               _ExtentY        =   1429
               _Version        =   393217
               ReadOnly        =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmReports.frx":6F98
            End
            Begin VSFlex7LCtl.VSFlexGrid fgMonthly 
               Height          =   1815
               Left            =   240
               TabIndex        =   28
               Top             =   3000
               Width           =   2535
               _cx             =   4471
               _cy             =   3201
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
            Begin VSFlex7LCtl.VSFlexGrid fgYearly 
               Height          =   1815
               Left            =   3360
               TabIndex        =   29
               Top             =   2880
               Width           =   2535
               _cx             =   4471
               _cy             =   3201
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
         Begin HexUniControls.ctlUniFrameWL fraReports 
            Height          =   7680
            Left            =   -13515
            TabIndex        =   17
            Top             =   330
            Width           =   11970
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
            Caption         =   "frmReports.frx":701A
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmReports.frx":7046
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmReports.frx":7066
            RightToLeft     =   0   'False
            Begin vsOcx6LibCtl.vsElastic txtReportName 
               Height          =   570
               Index           =   0
               Left            =   0
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   0
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   1005
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   2
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483636
               ForeColor       =   -2147483634
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "Performance Summary"
               Align           =   0
               Appearance      =   2
               AutoSizeChildren=   5
               BorderWidth     =   6
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   1
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   0   'False
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   0
               GridCols        =   0
               _GridInfo       =   ""
            End
            Begin VSFlex7LCtl.VSFlexGrid vsReport 
               Height          =   5190
               Left            =   0
               TabIndex        =   20
               Top             =   630
               Width           =   7410
               _cx             =   13070
               _cy             =   9155
               _ConvInfo       =   1
               Appearance      =   2
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
               BackColorFixed  =   12632256
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   1
               HighLight       =   0
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   0
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   ""
               ScrollTrack     =   -1  'True
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   0   'False
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   1
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
         Begin HexUniControls.ctlUniFrameWL fraPieChart 
            Height          =   7680
            Left            =   -12615
            TabIndex        =   9
            Top             =   330
            Width           =   11970
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
            Caption         =   "frmReports.frx":7082
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmReports.frx":70AE
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmReports.frx":70CE
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraBarChart 
            Height          =   7680
            Left            =   -12915
            TabIndex        =   8
            Top             =   330
            Width           =   11970
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
            Caption         =   "frmReports.frx":70EA
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmReports.frx":7116
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmReports.frx":7136
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraLegend 
               Height          =   555
               Left            =   300
               TabIndex        =   32
               Top             =   4740
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
               Caption         =   "frmReports.frx":7152
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmReports.frx":717E
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":719E
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniLabelXP lblLegendMaxIntra 
                  Height          =   255
                  Left            =   5640
                  Top             =   240
                  Width           =   1755
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
                  Caption         =   "frmReports.frx":71BA
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":7204
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":7224
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblLegendFiltered 
                  Height          =   255
                  Left            =   4140
                  Top             =   240
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
                  Caption         =   "frmReports.frx":7240
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":727E
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":729E
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblLegendMovingAverage 
                  Height          =   255
                  Left            =   2040
                  Top             =   240
                  Width           =   1755
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
                  Caption         =   "frmReports.frx":72BA
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":7304
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":7324
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblLegendUnfiltered 
                  Height          =   255
                  Left            =   420
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
                  Caption         =   "frmReports.frx":7340
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmReports.frx":7382
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmReports.frx":73A2
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin VB.Image imgLegendMaxIntra 
                  Height          =   240
                  Left            =   5340
                  Picture         =   "frmReports.frx":73BE
                  Top             =   247
                  Width           =   240
               End
               Begin VB.Image imgLegendFiltered 
                  Height          =   240
                  Left            =   3840
                  Picture         =   "frmReports.frx":7508
                  Top             =   247
                  Width           =   240
               End
               Begin VB.Image imgLegendMovingAverage 
                  Height          =   240
                  Left            =   1740
                  Picture         =   "frmReports.frx":7652
                  Top             =   247
                  Width           =   240
               End
               Begin VB.Image imgLegendUnfiltered 
                  Height          =   240
                  Left            =   120
                  Picture         =   "frmReports.frx":779C
                  Top             =   247
                  Width           =   240
               End
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdLessBars 
               Height          =   435
               Left            =   4620
               TabIndex        =   15
               Top             =   5880
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
               Caption         =   "frmReports.frx":78E6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmReports.frx":7918
               Style           =   1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":7938
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdMoreBars 
               Height          =   435
               Left            =   1980
               TabIndex        =   16
               Top             =   5460
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
               Caption         =   "frmReports.frx":7954
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmReports.frx":7986
               Style           =   1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":79A6
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblNoMMTrades 
               Height          =   675
               Left            =   2820
               Top             =   5160
               Visible         =   0   'False
               Width           =   4875
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
               Caption         =   "frmReports.frx":79C2
               BackColor       =   16777215
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmReports.frx":7A76
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":7A96
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblNoTrades 
               Height          =   255
               Left            =   2400
               Top             =   5640
               Visible         =   0   'False
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
               Caption         =   "frmReports.frx":7AB2
               BackColor       =   16777215
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmReports.frx":7AF6
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":7B16
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblValues 
               Height          =   435
               Left            =   180
               Top             =   5460
               Width           =   1755
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
               Caption         =   "frmReports.frx":7B32
               BackColor       =   16777215
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmReports.frx":7B5E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmReports.frx":7B7E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
      End
      Begin vsOcx6LibCtl.vsElastic elMsg 
         Height          =   285
         Left            =   4875
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   390
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   255
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   0
         WordWrap        =   0   'False
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
      End
      Begin HexUniControls.ctlUniLabelXP lblEquityFilter 
         Height          =   240
         Left            =   3570
         Top             =   90
         Width           =   11715
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
         Caption         =   "frmReports.frx":7B9A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmReports.frx":7C02
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmReports.frx":7C22
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Image imgStopGo 
         Height          =   240
         Left            =   3225
         Top             =   90
         Width           =   285
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP Corner 
      Height          =   375
      Left            =   9750
      TabIndex        =   3
      Top             =   6255
      Visible         =   0   'False
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
      Caption         =   "frmReports.frx":7C3E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmReports.frx":7C6A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmReports.frx":7C8A
      RightToLeft     =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuExportChart 
         Caption         =   "Export Chart"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Enum eRptGridType
    eRptGridType_Normal = 0
    eRptGridType_Implements = 1
End Enum

Private Enum eRptChartType
    eRptChartType_Bar = 0
    eRptChartType_Pie = 1
    eRptChartType_None = 2
End Enum

Private Enum eRptGridRows
    eRptGridRow_ReportName = 0
    eRptGridRow_ReportID = 1
    eRptGridRow_LeftMargin = 2
    eRptGridRow_RightMargin = 3
    eRptGridRow_TopMargin = 4
    eRptGridRow_BottomMargin = 5
    eRptGridRow_Orientation = 6
    eRptGridRow_FiltersAppliedFirstTime = 7
    eRptGridRow_ClassName = 8
    eRptGridRow_GridType = 9
    eRptGridRow_ChartType = 10
    eRptGridRow_SettingsFirstRow = 11
    eRptGridRow_SettingsLastRow = 12
    eRptGridRow_Order = 13
End Enum
Private Const kRptGridRows = 14

Private Const kDefaultReportID = 1
Private Const kDefaultReportIDPN = 30
Private Const kNoMoneyMgt = 0
Private Const kMoneyMgt = 1
Private Const kRegPathFormLoc = "SOFTWARE\Genesis Financial Data Services\Portfolio Navigator\FormLoc"

Private Type mPrivate
    bPyramiding As Boolean
    dbNav As Database
    Systems As cSystemsList
    Settings As cSettings
    ''gSettingsFile As cSettingsFile
    strAppName As String
    dDefaultBeginBalance As Double
    dFromDate As Double
    dToDate As Double
    strPortOrSystemName As String
    strAppPath As String
    MainForm As Object
    Images As ListImages
    Trades As cTrades
    bHideTdoReports As Boolean

    ReportObj As Object
    lCurrentGrid As Long
    ReportGrid As VSFlexGrid
    dBeginBalance As Double
    SettingsFile As cSettingsFile
    
    lReportID As Long
    strReportName As String
    strClassName As String
    ChartType As eRptChartType
    GridType As eRptGridType
    lOrder As Long
    strTabName As String
    strShowInSysNav As String
    
    dLeftMargin As Double
    dRightMargin As Double
    dTopMargin As Double
    dBottomMargin As Double
    strOrientation As String
    
    aOTSkip As cGdArray
    aOTRptSkip As cGdArray
    aOTOpenPos As cGdArray
    aOTUnits As cGdArray
    aOTTotProfit As cGdArray
    aOTAccBal As cGdArray
    aOTAvail As cGdArray
    hSkip As Long
    hRptSkip As Long
    hOpenPos As Long
    hUnits As Long
    hTotProfit As Long
    hAccBal As Long
    hAvail As Long
    
    lCustomCol As Long                  ' Custom Extend Column
    lMinColWidth As Long                ' Minimum Column Width for Extend Column
    lPrevColWidth As Long               ' Used for Extend custom column
    dLastX As Double
    dLastY As Double
    
    SystemsListHandle As Long
    TradesDtlHandle As Long
    TradesHdrHandle As Long
    
    hChartHwnd As Long
    
    Help As Object
    bIsLoaded As Boolean
    strTradesPath As String
    bAscending As Boolean
    
    lMovAvgPeriod As Long               ' Period for the equity moving average
    strMAType As String                 ' Type of moving average (Simple, Exponential)
    EquityFilter As cEquityFilter       ' Equity Filter options
    nTakeNextTrade As eGDTakeNextTradeValue ' Should we take the next trade?
    
    bShowViewOptions As Boolean         ' Show the view options for reports?
    nViewOption As eGDReportViewOptions ' Report viewing option
    
    dDelayedStartDate As Double         ' to delay the start of MM trading
    dAvgTradesPerYear As Double
    
    PrevRptSummary As cRptSummary       ' for MonteCarlo shuffle
End Type
Private m As mPrivate

Private Const mClass = "frmReports "

Private Function RptRow(ByVal lRow As eRptGridRows) As Long
    RptRow = lRow
End Function

Public Property Get Pyramiding() As Boolean
    Pyramiding = m.bPyramiding
End Property
Public Property Let Pyramiding(ByVal bPyramiding As Boolean)
    m.bPyramiding = bPyramiding
End Property
Public Property Get DB() As Database
    Set DB = m.dbNav
End Property
Public Property Let DB(dbNav As Database)
    Set m.dbNav = dbNav
End Property
Public Property Get Systems() As cSystemsList
    Set Systems = m.Systems
End Property
Public Property Let Systems(pData As cSystemsList)
    Set m.Systems = pData
End Property
Public Property Get Settings() As cSettings
    Set Settings = m.Settings
End Property
Public Property Let Settings(pData As cSettings)
    Set m.Settings = pData
End Property
Public Property Get SettingsFile() As cSettingsFile
    Set SettingsFile = m.SettingsFile
End Property
Public Property Let SettingsFile(pData As cSettingsFile)
    Set m.SettingsFile = pData
End Property
Public Property Get AppName() As String
    AppName = m.strAppName
End Property
Public Property Let AppName(ByVal strAppName As String)
    m.strAppName = strAppName
End Property
Public Property Get AppPath() As String
    AppPath = m.strAppPath
End Property
Public Property Let AppPath(ByVal strAppPath As String)
    m.strAppPath = strAppPath
End Property
Public Property Get DefaultBeginBalance() As Double
    DefaultBeginBalance = m.dDefaultBeginBalance
End Property
Public Property Let DefaultBeginBalance(ByVal dDefaultBeginBalance As Double)
    m.dDefaultBeginBalance = dDefaultBeginBalance
End Property
Public Property Get FromDate() As Double
    FromDate = m.dFromDate
End Property
Public Property Let FromDate(ByVal dFromDate As Double)
    m.dFromDate = dFromDate
End Property
Public Property Get ToDate() As Double
    ToDate = m.dToDate
End Property
Public Property Let ToDate(ByVal dToDate As Double)
    m.dToDate = dToDate
End Property
Public Property Get SystemName() As String
    SystemName = m.strPortOrSystemName
End Property
Public Property Let SystemName(ByVal strSystemName As String)
    m.strPortOrSystemName = strSystemName
End Property
Public Property Get MainForm() As Object
    Set MainForm = m.MainForm
End Property
Public Property Let MainForm(pData As Object)
    Set m.MainForm = pData
End Property
Public Property Get Images() As ListImages
    Set Images = m.Images
End Property
Public Property Let Images(pData As ListImages)
    Set m.Images = pData
End Property
Public Property Get Trades() As cTrades
    Set Trades = m.Trades
End Property
Public Property Let Trades(pData As cTrades)
    Set m.Trades = pData
End Property
Public Property Let SystemsListHandle(ByVal pData As Long)
    m.SystemsListHandle = pData
    Set m.Systems = New cSystemsList
    m.Systems.CopyTableFromHandle pData
End Property
Public Sub TradesHandles(ByVal pTradesDtlHandle As Long, ByVal pTradesHdrHandle As Long)
    m.TradesDtlHandle = pTradesDtlHandle
    m.TradesHdrHandle = pTradesHdrHandle
    Set m.Trades = New cTrades
    m.Trades.CopyTableFromHandle pTradesDtlHandle
    m.Trades.CopyHdrFromHandle pTradesHdrHandle
End Sub

Public Property Get CustomColumn() As Long
    CustomColumn = m.lCustomCol
End Property
Public Property Get MinColWidth() As Long
    MinColWidth = m.lMinColWidth
End Property
Public Property Let CustomColumn(ByVal pData As Long)
    m.lCustomCol = pData
End Property
Public Property Let MinColWidth(ByVal pData As Long)
    m.lMinColWidth = pData
End Property

Public Property Get ChartHwnd() As Long
    ChartHwnd = m.hChartHwnd
End Property
Public Property Let ChartHwnd(ByVal hChartHwnd As Long)
    m.hChartHwnd = hChartHwnd
End Property
Public Property Get Help() As Object
    Set Help = m.Help
End Property
Public Property Let Help(HelpObj As Object)
    Set m.Help = HelpObj
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = m.bIsLoaded
End Property
Public Property Get ReportID() As Long
    ReportID = m.lReportID
End Property

Public Property Get ShowViewOptions() As Boolean
    ShowViewOptions = m.bShowViewOptions
End Property
Public Property Let ShowViewOptions(ByVal bShowViewOptions As Boolean)
    m.bShowViewOptions = bShowViewOptions
    Form_Resize
End Property

Public Property Get ViewOption() As eGDReportViewOptions
    Select Case True
        Case optTrades
            ViewOption = eGDReportViewOption_Trades
        Case optMonthly
            ViewOption = eGDReportViewOption_Monthly
        Case optYearly
            ViewOption = eGDReportViewOption_Yearly
    End Select
End Property

Public Property Get HideTdoReports() As Boolean
    HideTdoReports = m.bHideTdoReports
End Property
Public Property Let HideTdoReports(ByVal bHideTdoReports As Boolean)
    m.bHideTdoReports = bHideTdoReports
End Property

Public Property Get MovingAveragePeriod() As Long
    MovingAveragePeriod = m.lMovAvgPeriod
End Property

Public Property Get MovingAverageType() As String
    MovingAverageType = m.strMAType
End Property

Public Property Get EquityFilter() As cEquityFilter
    Set EquityFilter = m.EquityFilter
End Property

Private Sub BarChart_DblClick()
On Error GoTo ErrSection:

    Dim lRow As Long

    If BarChart.PlottingMethod = GPM_LINE Or UCase(m.strClassName) = "CRPTTRADES2" Then
        If m.dLastX <> -99999 And m.dLastY <> -99999 Then
            lRow = m.ReportObj.RowFromPoint(m.dLastX, m.dLastY)
            If lRow > 0 Then
                vsTabs.CurrTab = eRptGridType_Implements
        
                With vsImplements
                    .Row = lRow
                    .RowSel = lRow
                    .ShowCell lRow, 0
                End With
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.BarChart.DblClick", eGDRaiseError_Show, m.strAppPath

End Sub

Private Sub BarChart_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim lpPoint As POINTSTRUCT
    Dim rc As Long

    If Shift = vbCtrlMask And KeyCode = vbKeyC Then
        ' capture as a metafile (to fit 8.5"x11" page with 1" margins)
        lpPoint.X = 2540 * 9    ' 9 inches wide (landscape)
        lpPoint.Y = 2540 * 6.5  ' 6.5 inches high
        rc = PEcopymetatoclipboard(BarChart.hObject, lpPoint)
        'rc = PEcopymetatofile(m.Peg, lpPoint, App.Path + "\Chart.WMF")
        If rc Then
            InfBox "i=i ; h=Process Graph Image ; You can now paste the graph into |another application by selecting |'Edit-Paste'  (or hit 'Ctrl-V')."
        Else
            InfBox "i=[] ; h=Export Image ; Image could not be exported!"
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.BarChart.KeyDown", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

Private Sub BarChart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        mnuExportChart.Visible = True
        mnuChangeFont.Visible = False
        
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.BarChart.MouseDown", eGDRaiseError_Show, m.strAppPath
    
End Sub

Private Sub cboRiskYears_Change()

    SetRiskTrades

End Sub

Private Sub cboRiskYears_Click()

    SetRiskTrades

End Sub

Private Sub cmdCompare_Click()

    Dim tTrades As cGdTable
    Set tTrades = GetProfits
    frmMonteCarlo.ShowMe GetProfits.FieldArray(0, True), , m.dAvgTradesPerYear

End Sub

Private Sub cmdLessBars_Click()
On Error GoTo ErrSection:

    If UCase(m.strClassName) = "CRPTTRADES2" Then
        m.ReportObj.LessBars
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.cmdLessBars.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub cmdMonteCarlo_Click()
On Error GoTo ErrSection:

    Dim i&, d#, s$, iRuns&, iTrades&
    Dim MonteCarlo As cMonteCarlo

    If 1 Then
        With vsRpts(0)
            If .Row <> .FixedRows Then
                .Row = .FixedRows
            End If
        End With
        If vsTabs.CurrTab <> 4 Then
            vsTabs.CurrTab = 4
        End If
    End If
            
    If optShuffled.Value Then
        i = MinMaxInteger(ValOfText(txtShuffled), 1, 99, 10)
        txtShuffled.Text = Str(i)
        If UCase(m.strClassName) = "CRPTSUMMARY" Then
            ' if from RptSummary object, use it for shuffling
            m.ReportObj.ShowShuffledTrades i
        ElseIf Not m.PrevRptSummary Is Nothing Then
            ' if from MM tabs, use the previous RptSummary object for shuffling
            m.PrevRptSummary.ShowShuffledTrades i
        End If
        BarChart2.Visible = True
    Else
        Set MonteCarlo = New cMonteCarlo
        ' get table of Trades and RiskPerTrade, and send to MonteCarlo
        MonteCarlo.SetTrades GetProfits
If IsIDE Then
        'MonteCarlo.ReadTrades "c:\dvlp\Batting800.txt"
End If
        If optRisk.Value Then
            iRuns = MinMaxInteger(ValOfText(txtRiskSims), 10, 9999, 1000)
            txtRiskSims = Str(iRuns)
            SetRiskTrades
            iTrades = Val(lblRiskTrades.Tag)
            d = MinMaxInteger(ValOfText(txtRiskStart), 1000, 1000000, 50000)
            txtRiskStart = Format(d, "$#,##0")
            
            MonteCarlo.RunRisks Me, iRuns, iTrades, d
        Else
            iRuns = MinMaxInteger(ValOfText(txtDrawdownSims), 10, 99999, 5000)
            txtDrawdownSims = Str(iRuns) ' Format(iRuns, "#,##0")
            iTrades = ValOfText(txtDrawdownTrades)
            txtDrawdownTrades = Str(iTrades)
        
            MonteCarlo.RunPerformance Me, iRuns, iTrades
        End If
        Set MonteCarlo = Nothing
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.cmdMonteCarlo_Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
End Sub

Private Sub cmdMoreBars_Click()
On Error GoTo ErrSection:

    If UCase(m.strClassName) = "CRPTTRADES2" Then
        m.ReportObj.MoreBars
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.cmdMoreBars.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub elImplements_ResizeChildren()
On Error Resume Next

    If Not m.ReportObj Is Nothing Then m.ReportObj.Resize

End Sub

Private Sub Form_Activate()
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean
    
    If bAlreadyDone = False Then
        Form_Resize
        bAlreadyDone = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.Form.Activate", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        'If Not m.Help Is Nothing Then m.Help.ShowF1Help Me
        If Not g.Help Is Nothing Then g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.Form.KeyDown", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If UnloadMode = 0 Then
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.Form.QueryUnload", eGDRaiseError_Show, m.strAppPath

End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim i&

    If m.lCustomCol <> -1 Then ExtendCustomColumn vsReport
    
    vsTabs.Refresh
    
    ' Reports Tab...
    With txtReportName(0)
        .Move 60, 60, fraReports.Width - 120
    End With
    With vsReport
        .Move 60, txtReportName(0).Height + 120, fraReports.Width - 120, fraReports.Height - txtReportName(0).Height - 180
    End With
    
    ' Implements Tab...
    With txtReportName(1)
        .Move 60, 60, fraImplements.Width - 120
    End With
    With fraViewOptions
        .Move 60, txtReportName(1).Height + 120
        .Visible = m.bShowViewOptions
    End With
    With vsImplements
        If m.bShowViewOptions = True Then
            .Move 60, fraViewOptions.Height + txtReportName(1).Height + 180, fraImplements.Width - 120, fraImplements.Height - txtReportName(1).Height - fraViewOptions.Height - 240
        Else
            .Move 60, txtReportName(1).Height + 120, fraImplements.Width - 120, fraImplements.Height - txtReportName(1).Height - 180
        End If
        
        fgMonthly.Move .Left, .Top, .Width, .Height
        fgYearly.Move .Left, .Top, .Width, .Height
    End With
    
    ' Bar Chart Tab...
    With lblValues
        .Move 0, fraBarChart.Height - .Height, fraBarChart.Width
    End With
    With cmdLessBars
        '.Move fraBarChart.Width - .Width * 2, fraBarChart.Height - .Height
        .Move vsTabs.ClientWidth - .Width * 2, vsTabs.ClientHeight - .Height
    End With
    With cmdMoreBars
        '.Move fraBarChart.Width - .Width, fraBarChart.Height - .Height
        .Move vsTabs.ClientWidth - .Width, vsTabs.ClientHeight - .Height
    End With
    With fraLegend
        .Move 0, lblValues.Top - .Height, fraBarChart.Width
    End With
    If UCase(m.strClassName) = "CRPTSUMMARY" Then
        With BarChart
            .Move 0, 0, fraBarChart.Width, fraBarChart.Height - lblValues.Height - fraLegend.Height
        End With
    Else
        With BarChart
            .Move 0, 0, fraBarChart.Width, fraBarChart.Height - lblValues.Height
        End With
    End If
    With lblNoTrades
        .Move (fraBarChart.Width - lblNoTrades.Width) / 2, (fraBarChart.Height - lblNoTrades.Height) / 2
    End With
    With lblNoMMTrades
        .Move (fraBarChart.Width - lblNoMMTrades.Width) / 2, (fraBarChart.Height - lblNoMMTrades.Height) / 2
    End With
    
    ' Pie Chart Tab...
    With PieChart
        .Move 0, 0, fraPieChart.Width, fraPieChart.Height
    End With
    
    With fraShuffled
        .Move 0, 0, fraMonteCarlo.Width, fraMonteCarlo.Height
    End With
    With fraDrawdown
        .Move 0, 0, fraMonteCarlo.Width, fraMonteCarlo.Height
    End With
    With fraRisk
        .Move 0, 0, fraMonteCarlo.Width, fraMonteCarlo.Height
    End With
    With BarChart2
        .Move 0, .Top, fraShuffled.Width, fraShuffled.Height - .Top
    End With
    With fgRisks
        .Height = fraRisk.Height - .Top - 300
        i = (fraRisk.Height - fraColors.Height) / 2
        If i < .Top Then i = .Top
        'fraColors.Top = i
    End With
    
    
    'With elMsg
    '    '.Move .Left, .Top, fraBarChart.Width - (.Left - fraBarChart.Left)
    '    .Move .Left, .Top, vsTabs.Width - .Left
    'End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lblEquityFilter_Click
'' Description: Allow the user to edit the equity filter information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblEquityFilter_Click()
On Error GoTo ErrSection:

    EditEquityFilter

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.lblEquityFilter_Click", , g.strAppPath

End Sub

Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the Redraw
    Dim lAutoSizeMode As Long           ' Current status of the Auto Size Mode

    If CommonDialogFont(CommonDialog1, vsReport.Font) Then
        With vsReport
            lRedraw = .Redraw
            lAutoSizeMode = .AutoSizeMode
            
            .Redraw = flexRDNone
            .Font = .Font '(this is required to trigger the grid to reset itself!)
            
            'If bResizeColumns Then
                .AutoSizeMode = flexAutoSizeColWidth
                .AutoSize 0, .Cols - 1, , 75
            'End If
            
            .AutoSizeMode = lAutoSizeMode
            .Redraw = lRedraw
        End With
        With vsImplements
            lRedraw = .Redraw
            lAutoSizeMode = .AutoSizeMode
            
            .Redraw = flexRDNone
            .Font = vsReport.Font
            .Font = .Font '(this is required to trigger the grid to reset itself!)
            
            'If bResizeColumns Then
                .AutoSizeMode = flexAutoSizeColWidth
                .AutoSize 0, .Cols - 1, , 75
            'End If
            
            .AutoSizeMode = lAutoSizeMode
            .Redraw = lRedraw
        End With
        
        If UCase(m.strClassName) = "CRPTSUMMARY" Then m.ReportObj.Run Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.mnuChangeFont.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub mnuExportChart_Click()
On Error GoTo ErrSection:

    Select Case vsTabs.CurrTab
        Case 2
            BarChart.PEactions = 6
            
        Case 3
            PieChart.PEactions = 6
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.mnuExportChart.Click", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optMonthly_Click
'' Description: Allow the user to view the monthly money management trades report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optMonthly_Click()
On Error GoTo ErrSection:

    If fgMonthly.Rows = fgMonthly.FixedRows + 1 Then
        m.ReportObj.CalculateMonthly
    End If

    vsImplements.Visible = False
    fgMonthly.Visible = True
    fgYearly.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.optMonthly.Click", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub


Private Sub optDrawdown_Click()
    
    cmdMonteCarlo.Enabled = True
    fraShuffled.Visible = False
    fraDrawdown.Visible = True
    fraRisk.Visible = False
    cmdCompare.Visible = True

End Sub

Private Sub optRisk_Click()

    SetRiskTrades
    cmdMonteCarlo.Enabled = True
    fraShuffled.Visible = False
    fraDrawdown.Visible = False
    fraRisk.Visible = True
    cmdCompare.Visible = False ' True

End Sub

Private Sub optShuffled_Click()

    cmdMonteCarlo.Enabled = True
    fraShuffled.Visible = True
    fraDrawdown.Visible = False
    fraRisk.Visible = False
    cmdCompare.Visible = False

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTrades_Click
'' Description: Allow the user to view the money management trades report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTrades_Click()
On Error GoTo ErrSection:

    vsImplements.Visible = True
    fgMonthly.Visible = False
    fgYearly.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.optTrades.Click", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optYearly_Click
'' Description: Allow the user to view the yearly money management trades report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optYearly_Click()
On Error GoTo ErrSection:

    If fgYearly.Rows = fgYearly.FixedRows + 1 Then
        m.ReportObj.CalculateYearly
    End If

    vsImplements.Visible = False
    fgMonthly.Visible = False
    fgYearly.Visible = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.optYearly.Click", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

Private Sub PieChart_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim lpPoint As POINTSTRUCT
    Dim rc As Long

    If Shift = vbCtrlMask And KeyCode = vbKeyC Then
        ' capture as a metafile (to fit 8.5"x11" page with 1" margins)
        lpPoint.X = 2540 * 9    ' 9 inches wide (landscape)
        lpPoint.Y = 2540 * 6.5  ' 6.5 inches high
        rc = PEcopymetatoclipboard(PieChart.hObject, lpPoint)
        'rc = PEcopymetatofile(m.Peg, lpPoint, App.Path + "\Chart.WMF")
        If rc Then
            InfBox "i=i ; h=Process Graph Image ; You can now paste the graph into |another application by selecting |'Edit-Paste'  (or hit 'Ctrl-V')."
        Else
            InfBox "i=[] ; h=Export Image ; Image could not be exported!"
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.PieChart.KeyDown", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

Private Sub PieChart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        mnuExportChart.Visible = True
        mnuChangeFont.Visible = False
        
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.PieChart.MouseDown", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

Private Sub Toolbar1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:
    
    Dim i&, s$
    
    Select Case Tool.ID
        Case "ID_Print"
            PrintMe
        
        Case "ID_Fields"
            m.ReportObj.ChangeFields
            
        Case "ID_Export"
            'ExportToTradeIT AddSlash(App.Path) & "TradeIT.TXT"
            Select Case vsTabs.CurrTab
                Case 0, 1
                    ExportToCSV
            
                Case 2
                    BarChart.PEactions = 6
                    
                Case 3
                    PieChart.PEactions = 6
                    
            End Select
            
        Case "ID_EquityFilter"
            EditEquityFilter
            
        Case "ID_Leave"
            Unload Me
        
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.Toolbar1.ToolClick", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit:

End Sub

Public Sub PrintMe()
On Error GoTo ErrSection:
    
    If vsTabs.CurrTab = 2 Then
        PEnset BarChart, 2978, True '(to make dialog modal)
        BarChart.PEactions = 8
    Else
        If vsTabs.CurrTab = 3 Then
            PEnset PieChart, 2978, True '(to make dialog modal)
            PieChart.PEactions = 8
        Else
            If vsTabs.CurrTab = 0 Then
                Set m.ReportGrid = vsReport
            Else
                Select Case True
                    Case optTrades
                        Set m.ReportGrid = vsImplements
                    Case optMonthly
                        Set m.ReportGrid = fgMonthly
                    Case optYearly
                        Set m.ReportGrid = fgYearly
                End Select
            End If
            If UCase(m.strAppName) = "SYSTEM NAVIGATOR" Then
                frmPrintPreview.ShowMe "SNV Reports", Me
            Else
                GenerateReport2
            End If
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.PrintMe", eGDRaiseError_Raise, m.strAppPath

End Sub

Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lIndex As Long, lTemp As Long
    Dim X   As Long
    Dim F As String, h As String
    Dim l(5) As String
    Dim svExtend        As Long
    Dim lFrozen         As Long
    
    Dim lCol As Long
    Dim lRow As Long
    Dim strText As String, strSymbols As String
    Dim aSymbols As New cGdArray
    
    ' get sorted list of unique symbols
    For lIndex = 0 To gdGetSize(m.Trades.SymbolHandle) - 1
        aSymbols.Add m.Trades.Symbol(lIndex)
    Next
    aSymbols.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
    strSymbols = aSymbols(0)
    For lIndex = 1 To aSymbols.Size - 1
        If lIndex > 50 Then
            strSymbols = strSymbols & " ..."
            Exit For
        End If
        strSymbols = strSymbols & ", " & aSymbols(lIndex)
    Next
    
    With frmPrintPreview.vp
        .StartDoc
        
        'Report Heading
        .HdrFontName = "Times New Roman"
        .HdrFontSize = 14
        If UCase(m.strAppName) = "SYSTEM NAVIGATOR" Then
            DoPrintHeader
        Else
            .Header = "|" & m.strAppName & vbCrLf & "Genesis Financial Data Services - (800) 808-DATA - www.TradeNavigator.com"
        End If
        .Footer = "|Page %d|"
        
        .FontName = "Arial"
        .FontSize = 14
        .FontBold = False
        .FontItalic = False
        .TextAlign = taLeftTop
        .TableBorder = tbAll
        .PageBorder = pbTopBottom
        .PenStyle = psSolid
        .BrushStyle = bsSolid
        .PenWidth = 2
        .PenColor = 0
        .BrushColor = 0
        .TextColor = 0
        .Columns = 1
        
        'Report Heading and date/time...
        .TextColor = 0
        .FontBold = True
        .FontSize = 14
        .TextAlign = taCenterMiddle
        .Text = m.strReportName & " Report"
        .FontBold = False
        .FontItalic = True
        .FontSize = 8
        .Paragraph = ""
        .TextAlign = taRightMiddle
        .Text = Format(Now(), "mmm d, yyyy  hh:mm:ss")
        .Paragraph = ""
        .Paragraph = ""
        
        'Blank line after printing heading, then draw a line under title
        .TextColor = 0
        .FontSize = 10
        .FontBold = False
        .FontItalic = False
        .SpaceAfter = 0
        
        'Summary section describing system
        .TextAlign = taLeftTop
        .TextColor = 0
        F = "<+1700|<+8000;"
        h = ""
        .TableBorder = tbNone
        l(1) = "Name:|" & m.strPortOrSystemName & ";"
        If aSymbols.Size > 1 Then
            l(2) = "Symbols:|" & strSymbols & ";"
        Else
            l(2) = "Symbol:|" & strSymbols & ";"
        End If
        l(3) = m.EquityFilter.RptPrintString
        Select Case m.nTakeNextTrade
            Case eGDTakeNextTrade_No
                l(4) = "Take Next Trade:|No"
            Case eGDTakeNextTrade_Yes
                l(4) = "Take Next Trade:|Yes"
            Case eGDTakeNextTrade_NoEquityFilter
                l(4) = "Take Next Trade:|No Equity Filter Applied"
            Case eGDTakeNextTrade_NotEnoughData
                l(4) = "Take Next Trade:|Not Enough Trades"
        End Select
        
        For X = 1 To UBound(l)
            If X = 2 And aSymbols.Size >= 8 Then
                .FontSize = 8
            Else
                .FontSize = 10
            End If
            .AddTable F, h, l(X)
        Next X
                
        .EndTable
        
        '-------------------------------------------------------------
        'Report filters and end of header section
        
        .Paragraph = ""
        .FontSize = 8
        .FontBold = False
        .FontItalic = False
        .Text = m.Settings.Text
        
        'Draw line between header information and report
        .Paragraph = ""
        .Paragraph = ""
        .DrawLine .MarginLeft, .CurrentY, _
                  .PageWidth - .MarginRight, .CurrentY
        .Paragraph = ""
        .Paragraph = ""
        
        ' Turn off the Frozen Columns and the Extend Last column to make
        ' the printing look better
        With m.ReportGrid
            .Redraw = flexRDNone
            svExtend = .ExtendLastCol
            .ExtendLastCol = False
            lFrozen = .FrozenCols
            .FrozenCols = 0
            .Redraw = flexRDBuffered
        End With
        
        'Print report
        If Not frmPrintPreview.GoingToFile Then
            .RenderControl = m.ReportGrid.hWnd
        Else
            frmPrintPreview.GridToTable m.ReportGrid
#If 0 Then
            With m.ReportGrid
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If .ColHidden(lCol) = False Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
#End If
        End If
        
        ' Restore the Frozen Columns and the Extend Last Column
        With m.ReportGrid
            .Redraw = flexRDNone
            .ExtendLastCol = svExtend
            .FrozenCols = lFrozen
            .Redraw = flexRDDirect
        End With
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.GenerateReport", eGDRaiseError_Raise, m.strAppPath
    
End Sub

Private Sub GenerateReport2()
On Error GoTo ErrSection:
    
'-------------------------------------------------
' Set layout data

    frmPrint.CallingForm = Me
    With frmPrint.vp
        .MarginBottom = m.dBottomMargin * 1440
        .MarginLeft = m.dLeftMargin * 1440
        .MarginRight = m.dRightMargin * 1440
        .MarginTop = m.dTopMargin * 1440
        If m.strOrientation = "Landscape" Then
            .Orientation = orLandscape
        Else
            .Orientation = orPortrait
        End If
    End With
    
    'Preview the report
    RunReport
    ShowForm frmPrint, True

'--------------------------------------------------
' Save layout data

    'Convert current layout settings back to inches before saving
    With frmPrint.vp
        If .MarginTop > 0 Then m.dTopMargin = .MarginTop / 1440
        If .MarginBottom > 0 Then m.dBottomMargin = .MarginBottom / 1440
        If .MarginLeft > 0 Then m.dLeftMargin = .MarginLeft / 1440
        If .MarginRight > 0 Then m.dRightMargin = .MarginRight / 1440
        If .Orientation = orPortrait Then
            m.strOrientation = "Portrait"
        Else
            m.strOrientation = "Landscape"
        End If
    End With
    
    'Save current report settings back to grid (hidden page layout stuff)
    'and then update the gd master table (gdSettings) using "UpdategdSettings"
    With m.Settings
        .SetItem "TopMargin", ValOfText(m.dTopMargin)
        .SetItem "BottomMargin", ValOfText(m.dBottomMargin)
        .SetItem "LeftMargin", ValOfText(m.dLeftMargin)
        .SetItem "RightMargin", ValOfText(m.dRightMargin)
        .SetItem "Orientation", m.strOrientation
        .Save
    End With
    
    'Update Report level settings
    With vsRpts(m.lCurrentGrid)
        .TextMatrix(.Row, RptRow(eRptGridRow_TopMargin)) = m.dTopMargin
        .TextMatrix(.Row, RptRow(eRptGridRow_BottomMargin)) = m.dBottomMargin
        .TextMatrix(.Row, RptRow(eRptGridRow_LeftMargin)) = m.dLeftMargin
        .TextMatrix(.Row, RptRow(eRptGridRow_RightMargin)) = m.dRightMargin
        .TextMatrix(.Row, RptRow(eRptGridRow_Orientation)) = m.strOrientation
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.GenerateReport2", eGDRaiseError_Raise, m.strAppPath

End Sub

Public Sub RunReport()
On Error GoTo ErrSection:
    
    Dim lIndex As Long, lTemp As Long
    Dim X   As Long
    Dim F As String, h As String
    Dim l(3) As String
    Dim svExtend        As Long
    Dim lFrozen         As Long
    
    'With frmPrintPreview.vp
    With frmPrint.vp
        .StartDoc
        
        'Report Heading
        .HdrFontName = "Times New Roman"
        .HdrFontSize = 14
        .Header = "|" & m.strAppName & vbCrLf & "Genesis Financial Data Services - (800) 808-DATA - www.TradeNavigator.com"
        .Footer = "|Page %d|"
        
        .FontName = "Arial"
        .FontSize = 14
        .FontBold = False
        .FontItalic = False
        .TextAlign = taLeftTop
        .TableBorder = tbAll
        .PageBorder = pbTopBottom
        .PenStyle = psSolid
        .BrushStyle = bsSolid
        .PenWidth = 2
        .PenColor = 0
        .BrushColor = 0
        .TextColor = 0
        .Columns = 1
        
        'Report Heading and date/time...
        .TextColor = 0
        .FontBold = True
        .FontSize = 14
        .TextAlign = taCenterMiddle
        .Text = m.strReportName & " Report"
        .FontBold = False
        .FontItalic = True
        .FontSize = 8
        .Paragraph = ""
        .TextAlign = taRightMiddle
        .Text = Format(Now(), "mmm d, yyyy  hh:mm:ss")
        .Paragraph = ""
        .Paragraph = ""
        
        'Blank line after printing heading, then draw a line under title
        .TextColor = 0
        .FontSize = 10
        .FontBold = False
        .FontItalic = False
        .SpaceAfter = 0
        
        'Summary section describing system
        .TextAlign = taLeftTop
        .TextColor = 0
        F = "<+1500|<+8000;"
        h = ""
        .TableBorder = tbNone
        l(1) = "Name:|" & m.strPortOrSystemName & ";"
        l(2) = m.EquityFilter.EnglishString(True, True)
        
        For X = 1 To UBound(l)
            .AddTable F, h, l(X)
        Next X
                
        .EndTable
        
        '-------------------------------------------------------------
        'Report filters and end of header section
        
        .Paragraph = ""
        .FontSize = 8
        .FontBold = False
        .FontItalic = False
        .Text = m.Settings.Text
        
        'Draw line between header information and report
        .Paragraph = ""
        .Paragraph = ""
        .DrawLine .MarginLeft, .CurrentY, _
                  .PageWidth - .MarginRight, .CurrentY
        .Paragraph = ""
        .Paragraph = ""
        
        ' Turn off the Frozen Columns and the Extend Last column to make
        ' the printing look better
        With m.ReportGrid
            .Redraw = flexRDNone
            svExtend = .ExtendLastCol
            .ExtendLastCol = False
            lFrozen = .FrozenCols
            .FrozenCols = 0
            .Redraw = flexRDBuffered
        End With
        
        'Print report
        .RenderControl = m.ReportGrid.hWnd
        
        ' Restore the Frozen Columns and the Extend Last Column
        With m.ReportGrid
            .Redraw = flexRDNone
            .ExtendLastCol = svExtend
            .FrozenCols = lFrozen
            .Redraw = flexRDDirect
        End With
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.RunReport", eGDRaiseError_Raise, m.strAppPath

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim w           As String
    Dim X           As Long
    Dim RetVal      As Long
    Dim dFileDate   As Double
    Dim strFont As String               ' Font string from the ini file
    Dim strCaption As String
    Dim d As Double
        
    strCaption = "Reports: " & m.strPortOrSystemName
    mnuPopUp.Visible = False
    
    'FixFonts Me
    FixFormControls Me
    
    cmdCompare.ZOrder
    fraType.ZOrder
    cmdMonteCarlo.ZOrder
    'RH commented out fraShuffled.BorderStyle = 0
    'RH commented out fraDrawdown.BorderStyle = 0
    'RH commented out fraRisk.BorderStyle = 0
   
    On Error Resume Next
    Me.Icon = m.Images("kPerformance").Picture
    If UCase(m.strAppName) = "SYSTEM NAVIGATOR" Then
        With Toolbar1
            .Tools("ID_EquityFilter").Picture = m.Images("kDollarLine").Picture
            .Tools("ID_Leave").Picture = m.Images("kCancel").Picture
            .Tools("ID_Leave").ChangeAll ssChangeAllName, "&Cancel"
            .Tools("ID_Print").Picture = m.Images("kPrint").Picture
            .Tools("ID_Shuffle").Picture = m.Images("kCards").Picture
            .Tools("ID_MonteCarlo").Picture = m.Images("kDice").Picture
            .Tools("ID_Export").Visible = False
            
            .Tools("ID_Shuffle").Visible = False
            .Tools("ID_MonteCarlo").Visible = False
        End With
        If FileExist(m.strTradesPath) Then
            dFileDate = FileDate(m.strTradesPath)
        Else
            dFileDate = Now
        End If
        
        strCaption = m.strPortOrSystemName & "  [Performance Reports]  -  Run " _
            & DateFormat(dFileDate) & " " & Format(dFileDate, "h:mm:ss AM/PM")
        If gdGetSize(m.Trades.SymbolHandle) = 1 Then
            strCaption = strCaption & " on " & m.Trades.Symbol(0)
        End If
    End If
    Me.Caption = strCaption
    On Error GoTo ErrSection:
    
    vsTabs.CurrTab = 0

    w = GetRegistryValue(rkLocalMachine, kRegPathFormLoc, "frmReports", "")
    If w = "" Then
        ReSizeMDIChildForm Me, Corner
        CenterTheForm Me
    Else
        SetFormPlacement Me, w, "LHTW"
    End If
    
    ' This helps to eliminate a Gray flicker while resizing the form
    ' 8/26/2002 DAJ
    fraBarChart.BackColor = BarChart.GraphBackColor
    lblValues.BackColor = BarChart.GraphBackColor
    lblValues.Caption = ""
    
    fraLegend.BackColor = BarChart.GraphBackColor
    lblLegendUnfiltered.BackColor = BarChart.GraphBackColor
    'lblLegendUnfiltered.ForeColor = vbRed
    lblLegendMovingAverage.BackColor = BarChart.GraphBackColor
    'lblLegendMovingAverage.ForeColor = vbBlue
    lblLegendFiltered.BackColor = BarChart.GraphBackColor
    'lblLegendFiltered.ForeColor = QBColor(2)
    lblLegendMaxIntra.BackColor = BarChart.GraphBackColor
    'lblLegendMaxIntra.ForeColor = vbBlack
    
    BarChart.AllowPopup = False
    PieChart.AllowPopup = False
        
    strFont = GetIniFileProperty("Reports", "", "Fonts", AddSlash(g.strAppPath) & "ChartNavigator.INI")
    If Len(strFont) > 0 Then
        FontFromString vsReport.Font, strFont
        FontFromString vsImplements.Font, strFont
    End If
    
    ' Get equity filter options from the INI file...
    m.lMovAvgPeriod = GetIniFileProperty("MaPeriod", 0&, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI")
    If m.lMovAvgPeriod = 0& Then
        m.lMovAvgPeriod = GetIniFileProperty("MovAvg", 30&, "Global", AddSlash(g.strAppPath) & "Reports.INI")
        SetIniFileProperty "MaPeriod", m.lMovAvgPeriod, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI"
    End If
    m.strMAType = GetIniFileProperty("MaType", "", "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI")
    If Len(m.strMAType) = 0 Then
        m.strMAType = GetIniFileProperty("MovAvgType", "Simple", "Global", AddSlash(g.strAppPath) & "Reports.INI")
        SetIniFileProperty "MaType", m.strMAType, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI"
    End If
    
    Set m.EquityFilter = New cEquityFilter
    EquityFilter.MovingAverageType = m.strMAType
    EquityFilter.MovingAveragePeriod = m.lMovAvgPeriod
    EquityFilter.EquityFilterOn = GetIniFileProperty("FilterOn", False, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI")
    EquityFilter.EquityFilterMode = GetIniFileProperty("FilterType", 0&, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI")
    lblEquityFilter.Caption = EquityFilter.EnglishString(m.nTakeNextTrade)
        
    'Make copy of Skip flag array...
    Set m.aOTSkip = New cGdArray
    RetVal = m.aOTSkip.Create(eGDARRAY_TinyInts)
    m.hSkip = m.Trades.FieldHandle(entd_Skip)
    gdCopy m.aOTSkip.ArrayHandle, m.hSkip
    
    'Make copy of Skip Rpt flag array...
    Set m.aOTRptSkip = New cGdArray
    RetVal = m.aOTRptSkip.Create(eGDARRAY_TinyInts)
    m.hRptSkip = m.Trades.FieldHandle(entd_SkipRpt)
    gdCopy m.aOTRptSkip.ArrayHandle, m.hRptSkip
    
    'Make copy of Skip Rpt flag array...
    Set m.aOTOpenPos = New cGdArray
    RetVal = m.aOTOpenPos.Create(eGDARRAY_Longs)
    m.hOpenPos = m.Trades.FieldHandle(entd_OpenTrade)
    gdCopy m.aOTOpenPos.ArrayHandle, m.hOpenPos
    
    'Make copy of units array...
    Set m.aOTUnits = New cGdArray
    RetVal = m.aOTUnits.Create(eGDARRAY_Doubles)
    m.hUnits = m.Trades.FieldHandle(entd_Units)
    gdCopy m.aOTUnits.ArrayHandle, m.hUnits
    
    'Make copy of TotProfit array...
    Set m.aOTTotProfit = New cGdArray
    RetVal = m.aOTTotProfit.Create(eGDARRAY_Doubles)
    m.hTotProfit = m.Trades.FieldHandle(entd_TotalProfit)
    gdCopy m.aOTTotProfit.ArrayHandle, m.hTotProfit

    'Make copy of AccBal array...
    Set m.aOTAccBal = New cGdArray
    RetVal = m.aOTAccBal.Create(eGDARRAY_Doubles)
    m.hAccBal = m.Trades.FieldHandle(entd_AccountBalance)
    gdCopy m.aOTAccBal.ArrayHandle, m.hAccBal
    
    'Make copy of AccBal array...
    Set m.aOTAvail = New cGdArray
    RetVal = m.aOTAvail.Create(eGDARRAY_Doubles)
    m.hAvail = m.Trades.FieldHandle(entd_EquityAvail)
    gdCopy m.aOTAvail.ArrayHandle, m.hAvail

    'Save beginning balance and update Settings later
    If m.Trades.Num(0, entd_AccountBalance) > 0 Then
        m.dBeginBalance = m.Trades.Num(0, entd_AccountBalance)
    Else
        m.dBeginBalance = m.dDefaultBeginBalance
    End If
    
    ' Default to the money management trades being shown...
    optTrades.Value = True
    
    With cboRiskYears
        For X = 1 To 30
            .AddItem Str(X)
        Next
        .ListIndex = 0
    End With
    
    ' get Monte Carlo settings
    d = Round(GetIniFileProperty("RiskYears", 0&, "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI"))
    If d > 0 And d < cboRiskYears.ListCount Then
        cboRiskYears.ListIndex = d - 1
    End If
    d = GetIniFileProperty("RiskStart", 0&, "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI")
    If d > 0 And d < 10000000 Then
        txtRiskStart = Format(d, "$#,##0")
    End If
    d = GetIniFileProperty("RiskSims", 0&, "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI")
    If d > 0 And d < 1000000 Then
        txtRiskSims = Str(Round(d))
    End If
    d = GetIniFileProperty("DrawdownSims", 0&, "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI")
    If d > 0 And d < 1000000 Then
        txtDrawdownSims = Str(Round(d))
    End If
    d = GetIniFileProperty("NumShuffled", 0&, "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI")
    If d > 0 And d < 100 Then
        txtShuffled = Str(Round(d))
    End If
    
    'Check to see if user owns Portfolio Navigator.  If valid ownership
    'then show all money management reports on the money mgt tabl.
    'Otherwise, only enable 3 reports: Williams, Risk/ratio, Optimal f.
    'Gray out the other reports and display a message if they are clicked
    'to purchase Portfolio Navigator if you want to use these.
    '===
    '===
    
    LoadReports
    
    m.bIsLoaded = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.Form.Load", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub LoadReports()
On Error GoTo ErrSection:
    
    Dim X               As Long
    Dim RetVal          As Variant
    Dim Continue        As Boolean
    Dim svAppID         As Long
    Dim bShowTime       As Boolean
    
    Set m.SettingsFile = New cSettingsFile
    With m.SettingsFile
        .Path = m.strAppPath
        .FileName = "gdSettings.dat"
        .IsStrategy = Not m.bHideTdoReports
        .Load
    End With
    
    'Instance settings class to manage options grid
    Set m.Settings = New cSettings
    With m.Settings
        .SettingsFile = m.SettingsFile
        .DB = m.dbNav
        .vsGrid = vsSettings
        .PreviewRtf = txtPreview
        .NewItem = False
    End With
    
    'Load money mgt reports
    With vsRpts(kMoneyMgt)
        InitReportsGrid kMoneyMgt
        .Redraw = flexRDNone
        
        'Loop through settings of each report and pick out report related
        'settings and save in grid...
        For X = 1 To m.SettingsFile.NumRecords - 1
            svAppID = m.SettingsFile.Num(ensgd_AppID, X)
            Continue = False
            m.strTabName = ""
            Do Until svAppID <> m.SettingsFile.Num(ensgd_AppID, X)
                Select Case m.SettingsFile.Item(ensgd_SettingName, X)
                    Case "ReportName": m.strReportName = m.SettingsFile.Item(ensgd_Default, X)
                    Case "Order": m.lOrder = m.SettingsFile.Item(ensgd_Default, X)
                    Case "ClassName": m.strClassName = m.SettingsFile.Item(ensgd_Default, X)
                    Case "LeftMargin": m.dLeftMargin = m.SettingsFile.Num(ensgd_Default, X)
                    Case "RightMargin": m.dRightMargin = m.SettingsFile.Num(ensgd_Default, X)
                    Case "TopMargin": m.dTopMargin = m.SettingsFile.Num(ensgd_Default, X)
                    Case "BottomMargin": m.dBottomMargin = m.SettingsFile.Num(ensgd_Default, X)
                    Case "Orientation": m.strOrientation = m.SettingsFile.Item(ensgd_Default, X)
                    Case "GridType": m.GridType = m.SettingsFile.Num(ensgd_Default, X)
                    Case "ChartType": m.ChartType = m.SettingsFile.Num(ensgd_Default, X)
                    Case "TabName": m.strTabName = m.SettingsFile.Item(ensgd_Default, X)
                    Case "ShowInSysNav": m.strShowInSysNav = m.SettingsFile.Item(ensgd_Default, X)
                End Select
                X = X + 1
            Loop
            X = X - 1
            
            'Add report to grid
            If m.strTabName = "Money Mgt" Then
                Continue = True
                If m.strAppName = "System Navigator" Then
                    If m.strShowInSysNav = "No" And (Not gbForceMM) Then
                        Continue = False
                    End If
                End If
                
                If Continue Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_ReportID)) = svAppID
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_ReportName)) = m.strReportName
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_ClassName)) = m.strClassName
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_LeftMargin)) = m.dLeftMargin
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_RightMargin)) = m.dRightMargin
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_TopMargin)) = m.dTopMargin
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_BottomMargin)) = m.dBottomMargin
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_Orientation)) = m.strOrientation
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_FiltersAppliedFirstTime)) = 0
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_GridType)) = m.GridType
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_ChartType)) = m.ChartType
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_Order)) = m.lOrder
                    m.SettingsFile.SetItem svAppID, "FromDate", m.dFromDate, ensgd_Value
                    m.SettingsFile.SetItem svAppID, "ToDate", m.dToDate, ensgd_Value
                    'm.SettingsFile.SetItem svAppID, "BeginBalance", m.dBeginBalance, ensgd_Value
                    m.SettingsFile.SetItem svAppID, "Margin", m.Trades.ItemHdr(1, enth_Margin), ensgd_Value
                    If m.Trades.ItemHdr(1, enth_IntraDaySystem) = 0 Then
                        m.SettingsFile.SetItem svAppID, "ShowTime", "No", ensgd_Value
                    Else
                        m.SettingsFile.SetItem svAppID, "ShowTime", "Yes", ensgd_Value
                    End If
                    m.SettingsFile.SetItem svAppID, "ShowTime", False, ensgd_ShowEdit
                End If
            End If
        Next X
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDDirect
    End With
    
    'Load performance reports from gdTable settings
    With vsRpts(kNoMoneyMgt)
        InitReportsGrid kNoMoneyMgt
        .Redraw = flexRDNone
        
        'Loop through settings of each report and pick out report related
        'settings and save in grid...
        For X = 1 To m.SettingsFile.NumRecords - 1
            svAppID = m.SettingsFile.Num(ensgd_AppID, X)
            Continue = False
            m.strTabName = ""
            Do Until svAppID <> m.SettingsFile.Num(ensgd_AppID, X)
                Select Case m.SettingsFile.Item(ensgd_SettingName, X)
                    Case "ReportName": m.strReportName = m.SettingsFile.Item(ensgd_Default, X)
                    Case "Order": m.lOrder = m.SettingsFile.Item(ensgd_Default, X)
                    Case "ClassName": m.strClassName = m.SettingsFile.Item(ensgd_Default, X)
                    Case "LeftMargin": m.dLeftMargin = m.SettingsFile.Num(ensgd_Default, X)
                    Case "RightMargin": m.dRightMargin = m.SettingsFile.Num(ensgd_Default, X)
                    Case "TopMargin": m.dTopMargin = m.SettingsFile.Num(ensgd_Default, X)
                    Case "BottomMargin": m.dBottomMargin = m.SettingsFile.Num(ensgd_Default, X)
                    Case "Orientation": m.strOrientation = m.SettingsFile.Item(ensgd_Default, X)
                    Case "GridType": m.GridType = m.SettingsFile.Num(ensgd_Default, X)
                    Case "ChartType": m.ChartType = m.SettingsFile.Num(ensgd_Default, X)
                    Case "TabName": m.strTabName = m.SettingsFile.Item(ensgd_Default, X)
                    Case "ShowInSysNav": m.strShowInSysNav = m.SettingsFile.Item(ensgd_Default, X)
                End Select
                X = X + 1
            Loop
            X = X - 1
            
            'Add report to grid
            If m.strTabName = "Analysis" Then
                Continue = True
                If m.strAppName = "System Navigator" Then
                    If m.strShowInSysNav = "No" And (Not gbForceMM) Then
                        Continue = False
                    End If
                End If
                
                If Continue Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_ReportID)) = svAppID
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_ReportName)) = m.strReportName
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_ClassName)) = m.strClassName
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_LeftMargin)) = m.dLeftMargin
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_RightMargin)) = m.dRightMargin
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_TopMargin)) = m.dTopMargin
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_BottomMargin)) = m.dBottomMargin
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_Orientation)) = m.strOrientation
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_FiltersAppliedFirstTime)) = 0
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_GridType)) = m.GridType
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_ChartType)) = m.ChartType
                    .TextMatrix(.Rows - 1, RptRow(eRptGridRow_Order)) = m.lOrder
                    m.SettingsFile.SetItem svAppID, "FromDate", m.dFromDate, ensgd_Value
                    m.SettingsFile.SetItem svAppID, "ToDate", m.dToDate, ensgd_Value
                    'm.SettingsFile.SetItem svAppID, "BeginBalance", m.dBeginBalance, ensgd_Value
                    m.SettingsFile.SetItem svAppID, "Margin", m.Trades.ItemHdr(1, enth_Margin), ensgd_Value
                    If m.Trades.ItemHdr(1, enth_IntraDaySystem) = 0 Then
                        m.SettingsFile.SetItem svAppID, "ShowTime", "No", ensgd_Value
                    Else
                        m.SettingsFile.SetItem svAppID, "ShowTime", "Yes", ensgd_Value
                    End If
                    m.SettingsFile.SetItem svAppID, "ShowTime", False, ensgd_ShowEdit
                    
                    If UCase(Left(m.strReportName, 6)) = "BY TDO" Then
                        .RowHidden(.Rows - 1) = m.bHideTdoReports
                    End If
                End If
            End If
        Next X
                
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDDirect
    End With
    
    'Default the report to the performance report summary
    If m.strAppName = "Portfolio Navigator" Then
        m.lReportID = kDefaultReportIDPN
        vsIndexTab1.CurrTab = 1
        RetVal = vsRpts(kMoneyMgt).CellTop
    Else
        m.lReportID = kDefaultReportID
        RetVal = vsRpts(kNoMoneyMgt).CellTop
    End If

    'Show default report
    vsIndexTab1_Click
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.LoadReports", eGDRaiseError_Raise, m.strAppPath

End Sub

'This is the driver for showing a selected report
Private Sub ProcessRpt()
On Error GoTo ErrSection:
    
    Dim X As Long
    
    'Validations
    If m.Trades Is Nothing Then
        Err.Raise gUserErr, , "No trade history was found.  Reports cannot be viewed."
    End If
    If m.Trades.NumRecords - 1 <= 0 Then
        Err.Raise gUserErr, , "No trade history was found.  Reports cannot be viewed."
    End If
    If m.Trades.NumRecords = 3 And m.Trades.Item(2, entd_RuleID) = 0 Then
        Err.Raise gUserErr, , "There was only one entry with no exit.  Reports cannot be viewed."
    End If
    If Len(m.strClassName) = 0 Then
        Err.Raise gUserErr, , "Please select a report"
    End If
    
    Select Case m.strClassName
        Case "cRptSummary"
            Set m.ReportObj = New cRptSummary
        Case "cRptSummaryBy"
            Set m.ReportObj = New cRptSummaryBy
        Case "cRptMMStudy"
            Set m.ReportObj = New cRptMMStudy
            m.ReportObj.DelayedStartDate = m.dDelayedStartDate
        Case "cRptAllocations"
            Set m.ReportObj = New cRptAllocations
        Case "cRptConsecStudy"
            Set m.ReportObj = New cRptConsecStudy
        'Case "cRptFixedFractional"
        '    Set m.ReportObj = New cRptFixedFractional
        'Case "cRptFixedRatio"
        '    Set m.ReportObj = New cRptFixedRatio
        Case "cRptLTStudy"
            Set m.ReportObj = New cRptLTStudy
            m.bAscending = True
        Case "cRptMMByYear"
            Set m.ReportObj = New cRptMMByYear
        Case "cRptMMDDAnalysis"
            Set m.ReportObj = New cRptMMDDAnalysis
        Case "cRptMMRtnDtl"
            Set m.ReportObj = New cRptMMRtnDtl
        Case "cRptMMRtnSum"
            Set m.ReportObj = New cRptMMRtnSum
        Case "cRptMMSummary"
            Set m.ReportObj = New cRptMMSummary
        Case "cRptMMSummary2"
            Set m.ReportObj = New cRptMMSummary2
        Case "cRptMMTrades"
            Set m.ReportObj = New cRptMMTrades
        'Case "cRptOptimalf"
        '    Set m.ReportObj = New cRptOptimalf
        Case "cRptSystems"
            Set m.ReportObj = New cRptSystems
        Case "cRptTrades2"
            Set m.ReportObj = New cRptTrades2
        'Case "cRptWilliamsRiskRatio"
        '    Set m.ReportObj = New cRptWilliamsRiskRatio
        Case "cRptWinStreakStudy"
            Set m.ReportObj = New cRptWinStreakStudy
        Case Else
            Set m.ReportObj = CreateObject(m.strClassName)
    End Select
    If m.ReportObj Is Nothing Then
        Err.Raise gUserErr, , "Report could not be created.  Class: '" & m.strClassName & "' was not found."
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Restore trade arrays...
    gdCopy m.hSkip, m.aOTSkip.ArrayHandle
    gdCopy m.hRptSkip, m.aOTRptSkip.ArrayHandle
    gdCopy m.hOpenPos, m.aOTOpenPos.ArrayHandle
    gdCopy m.hUnits, m.aOTUnits.ArrayHandle
    gdCopy m.hTotProfit, m.aOTTotProfit.ArrayHandle
    gdCopy m.hAccBal, m.aOTAccBal.ArrayHandle
    gdCopy m.hAvail, m.aOTAvail.ArrayHandle
    
    'If called from Portfolio navigator, ensure that the skip flags are
    'set off for Non-money management reporting (they will always be off
    'when called from System Navigator)
    If vsIndexTab1.CurrTab = 0 Then
        If m.strAppName = "Portfolio Navigator" Then
            For X = 0 To m.Trades.NumRecords - 1
                gdSetNum m.hSkip, X, 0
                gdSetNum m.hRptSkip, X, 0
            Next X
        End If
    End If
    
    'Set active tab and grid report title
    SetTabs
    SetReportTitle
    
    'Default the Messages elastic invisible and autofill the trades grid
    'elMessages.Visible = False
    'elImplements.Align = asFill
    rtfMessages.Visible = False
    
    'Update Settings gdTable with changes from vsSettings
    m.Settings.Save
    
    'Pass references to grid to report object
    'Select Case m.GridType
    '    Case eRptGridType_Normal
            'Set mvsgrid = vsReport
    '    Case eRptGridType_Implements
            'Set mvsgrid = vsImplements
            'Set mRtfMsg = rtfMessages
            'Set melMsg = elMessages
            'Set melGrid = elImplements
    'End Select
    
    'Pass references to chart to report object
    'Select Case m.ChartType
    '    Case eRptChartType_Bar: Set mChart = BarChart.Object
    '    Case eRptChartType_Pie: Set mChart = PieChart.Object
    '    Case eRptChartType_None
    'End Select
    
    'Make appropriate tab active...
    If vsTabs.CurrTab = 0 Or vsTabs.CurrTab = 1 Then
        If m.GridType = eRptGridType_Normal Then
            vsTabs.CurrTab = 0
            MoveFocus vsReport
        Else
            vsTabs.CurrTab = 1
            MoveFocus vsImplements
        End If
    Else
        If m.ChartType = eRptChartType_Bar Then
            vsTabs.CurrTab = 2
            MoveFocus BarChart
        Else
            If m.ChartType = eRptChartType_Pie Then
                vsTabs.CurrTab = 3
                MoveFocus PieChart
            Else
                If m.GridType = eRptGridType_Normal Then
                    vsTabs.CurrTab = 0
                    MoveFocus vsReport
                Else
                    vsTabs.CurrTab = 1
                    MoveFocus vsImplements
                End If
            End If
        End If
    End If
    
    'Run/Show report
    elMsg.Caption = ""
    m.lCustomCol = -1
    m.ReportObj.Run Me
    
#If 0 Then
    Select Case m.GridType
        Case eRptGridType_Normal
            GridToFile vsReport, AddSlash(g.strAppPath) & m.strClassName & ".TXT"
            
        Case eRptGridType_Implements
            GridToFile vsImplements, AddSlash(g.strAppPath) & m.strClassName & ".TXT"
            
    End Select
#End If

    Screen.MousePointer = vbDefault
    'mReloadTrades = False
    
    If UCase(m.strClassName) = "CRPTSUMMARY" Then
        Set m.PrevRptSummary = m.ReportObj
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.ProcessRpt", eGDRaiseError_Raise, m.strAppPath

End Sub

Private Sub SetTabs()
On Error GoTo ErrSection:

    vsTabs.TabVisible(0) = False
    vsTabs.TabVisible(1) = False
    vsTabs.TabVisible(2) = False
    vsTabs.TabVisible(3) = False
    
    Select Case m.GridType
        Case eRptGridType_Normal: vsTabs.TabVisible(0) = True
        Case eRptGridType_Implements: vsTabs.TabVisible(1) = True
    End Select
    Select Case m.ChartType
        Case eRptChartType_Bar: vsTabs.TabVisible(2) = True
        Case eRptChartType_Pie: vsTabs.TabVisible(3) = True
        Case eRptChartType_None
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.SetTabs", eGDRaiseError_Raise, m.strAppPath

End Sub

Private Sub SetReportTitle()
On Error GoTo ErrSection:

    Select Case m.GridType
        Case eRptGridType_Normal: txtReportName(0).Caption = m.strReportName
        Case eRptGridType_Implements: txtReportName(1).Caption = m.strReportName
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.SetReportTitle", eGDRaiseError_Raise, m.strAppPath

End Sub

Private Sub InitReportsGrid(pIndex As Integer)
On Error GoTo ErrSection:
    
    Dim X As Integer
    
    'Clear and reinitialize grid...
    With vsRpts(pIndex)
        .Redraw = False
        .Clear
        .AllowBigSelection = False
        .AllowSelection = True
        .TabBehavior = flexTabCells
        .Editable = True
        .FontName = "Arial"
        .FontSize = 8
        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .ScrollTips = True
        .ScrollTrack = True
        .Ellipsis = flexEllipsisEnd
        .Cols = kRptGridRows
        .Rows = 0
        .FixedCols = 0
        .FixedRows = 0
        For X = 0 To .Cols - 1
            .ColHidden(X) = True
        Next X
        .ColHidden(RptRow(eRptGridRow_ReportName)) = False
        .Redraw = flexRDDirect
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.InitReportsGrid", eGDRaiseError_Raise, m.strAppPath
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    Dim w As String

    SetIniFileProperty "Reports", FontToString(vsReport.Font), "Fonts", AddSlash(g.strAppPath) & "ChartNavigator.INI"
    
    ' Only save the placement and size if non-minimized
    If WindowState <> vbMinimized Then
        w = GetFormPlacement(Me)
        SetRegistryValue rkLocalMachine, kRegPathFormLoc, "frmReports", w
    End If
    
    m.SettingsFile.Serialize
    
    ' save Monte Carlo settings
    SetIniFileProperty "RiskYears", cboRiskYears.ListIndex + 1, "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI"
    SetIniFileProperty "RiskStart", ValOfText(txtRiskStart.Text), "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI"
    SetIniFileProperty "RiskSims", ValOfText(txtRiskSims.Text), "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI"
    SetIniFileProperty "DrawdownSims", ValOfText(txtDrawdownSims.Text), "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI"
    SetIniFileProperty "NumShuffled", ValOfText(txtShuffled.Text), "MonteCarlo", AddSlash(g.strAppPath) & "Reports.INI"
        
    m.bIsLoaded = False
    
ErrExit:
    Set m.Trades = Nothing
    Set m.Systems = Nothing
    Set m.Settings = Nothing
    Set m.SettingsFile = Nothing
    Set m.Images = Nothing
    Set m.ReportObj = Nothing
    Set m.PrevRptSummary = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmReports.Form.Unload", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub BarChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    Dim nA As Long
    Dim nX As Long
    Dim nY As Long
    Dim fX As Double
    Dim fY As Double
    Dim pt As POINTSTRUCT
    Dim t As Long
    Dim R As Rect

    '** get last mouse location within control **'
    t = PEvget(BarChart, PEP_ptLASTMOUSEMOVE, pt)

    '** test to see if this is within grid area **'
    t = PEvget(BarChart, PEP_rectGRAPH, R)
    If pt.X > R.Left And pt.X < R.Right And pt.Y > R.Top And pt.Y < R.Bottom Then
        nA = 0      'Initialize axis, non-zero only if using MultiAxesSubsets
        nX = pt.X   'Initialize nX and nY with mouse location
        nY = pt.Y
        t = PEconvpixeltograph(BarChart, nA, nX, nY, fX, fY, 0, 0, 0)
        m.dLastX = fX
        m.dLastY = fY
        m.ReportObj.DisplayValues fX, fY, lblValues
    Else
        m.dLastX = -99999
        m.dLastY = -99999
        lblValues.Caption = ""
        lblValues.Refresh
    End If
  
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.BarChart.MouseMove", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsElastic5_ResizeChildren(Index As Integer)
On Error Resume Next

    If Not m.ReportObj Is Nothing Then m.ReportObj.Resize

End Sub

Private Sub vsElastic1_ResizeChildren()
On Error Resume Next

    ' Need to do this so that the original report gets sized correctly...
    Form_Resize

End Sub

Private Sub vsImplements_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error Resume Next

    Select Case UCase(m.strClassName)
        Case "CRPTTRADES2", "CRPTMMSTUDY"
            m.ReportObj.SaveCols
    End Select

End Sub

Private Sub vsImplements_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

    Select Case UCase(m.strClassName)
        Case "CRPTTRADES2", "CRPTMMSTUDY"
            m.ReportObj.SaveCols
    End Select

End Sub

Private Sub vsImplements_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:
    
    If Button = vbLeftButton Then
        With vsImplements
            If Not .FlexDataSource Is Nothing And .MouseRow = 0 Then
                m.ReportObj.SortOnCol .MouseCol
            End If
        End With
    Else
        mnuExportChart.Visible = False
        PopupMenu mnuPopUp ', , X, Y
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsImplements.BeforeMouseDown", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsImplements_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    With vsImplements
        If Col < 15 Then
            Position = Col
        ElseIf Position < 15 Then
            Position = 15
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsImplements_BeforeMoveColumn"
    
End Sub

Private Sub vsImplements_DblClick()
On Error Resume Next
    
    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    lMouseRow = vsImplements.MouseRow
    lMouseCol = vsImplements.MouseCol
    
    If Not vsImplements.FlexDataSource Is Nothing Then
        If lMouseRow >= vsImplements.FixedRows And lMouseRow < vsImplements.Rows Then
            m.ReportObj.DblClick lMouseRow, lMouseCol
            If m.strClassName = "cRptMMStudy" Then
                ' for MM: delay the start date for MM trading until where they just dbl-clicked
                m.dDelayedStartDate = m.ReportObj.DelayedStartDate
                ProcessRpt
            End If
        End If
    End If

End Sub

Private Sub vsIndexTab1_Click()
On Error GoTo ErrSection:
    
    If vsIndexTab1.CurrTab = 0 Then
        With vsRpts(0)
            .Redraw = flexRDNone
            .Select 0, RptRow(eRptGridRow_Order)
            .Sort = flexSortGenericAscending
            .Redraw = flexRDBuffered
            .Row = 0
        End With
        vsRpts_AfterRowColChange 0, 0, 0, 0, 0
    Else
        With vsRpts(1)
            .Redraw = flexRDNone
            .Select 0, RptRow(eRptGridRow_Order)
            .Sort = flexSortGenericAscending
            .Redraw = flexRDBuffered
            .Row = 0
        End With
        vsRpts_AfterRowColChange 1, 0, 0, 0, 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsIndexTab1.Click", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsReport_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error Resume Next

    If UCase(m.strClassName) = "CRPTSUMMARYBY" Then m.ReportObj.SaveCols

End Sub

Private Sub vsReport_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsReport.AfterSort", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsReport_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        With vsReport
            'If Not .FlexDataSource Is Nothing And .MouseRow = 0 Then
            '    m.ReportObj.SortOnCol .MouseCol
            'End If
            mnuExportChart.Visible = False
            PopupMenu mnuPopUp ', , X, Y
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsReport.BeforeMouseDown", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsReport_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    With vsReport
        If Col < .FrozenCols Then
            Position = Col
        ElseIf Position < .FrozenCols Then
            Position = .FrozenCols
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsReport_BeforeMoveColumn"
    
End Sub

Private Sub vsReport_BeforeScrollTip(ByVal Row As Long)
On Error Resume Next
    
    m.ReportObj.BeforeScrollTip Row

End Sub

Private Sub vsReport_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error Resume Next
    
    m.ReportObj.AfterScroll OldTopRow, OldLeftCol, NewTopRow, NewLeftCol

End Sub

Private Sub vsReport_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsReport.BeforeEdit", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsImplements_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsImplements.BeforeEdit", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsImplements_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error Resume Next
    
    m.ReportObj.AfterScroll OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Me

End Sub

Private Sub vsImplements_BeforeScrollTip(ByVal Row As Long)
On Error Resume Next
    
    m.ReportObj.BeforeScrollTip Row

End Sub

Private Sub vsImplements_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    'If elMessages.Visible Then
    '    rtfMessages.Text = ""
    '    m.ReportObj.RowColChange
    'End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.vsImplements.RowColChange", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsReport_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    If m.strClassName = "cRptLTStudy" Then
        If Col = 0 Then
            m.bAscending = Not m.bAscending
            Order = flexSortCustom
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsReport.BeforeSort", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

Private Sub vsReport_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
On Error GoTo ErrSection:

    Dim strRow1 As String
    Dim strRow2 As String
    Dim dValue1 As Double
    Dim dValue2 As Double
    
    If m.strClassName = "cRptLTStudy" Then
        strRow1 = Right(vsReport.TextMatrix(Row1, 0), Len(vsReport.TextMatrix(Row1, 0)) - 10)
        strRow2 = Right(vsReport.TextMatrix(Row2, 0), Len(vsReport.TextMatrix(Row2, 0)) - 10)
        dValue1 = ValOfText(strRow1)
        dValue2 = ValOfText(strRow2)
        
        If dValue1 = dValue2 Then
            Cmp = 0
        ElseIf dValue1 > dValue2 Then
            Cmp = 1
        Else
            Cmp = -1
        End If
        
        If m.bAscending = False Then Cmp = -Cmp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsReport.Compare", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    Static nPrevRow As Long, nPrevCol As Long
    Dim nRow As Long, nCol As Long, strTip$
    
    If UCase(m.strClassName) = "CRPTSUMMARY" Then
        With vsReport
            ' only change tooltip if row or column has changed
            nRow = .MouseRow
            nCol = .MouseCol
            If nRow <> nPrevRow Or nCol <> nPrevCol Then
                nPrevRow = nRow
                nPrevCol = nCol
                .ToolTipText = m.ReportObj.ToolTip(nRow, nCol)
            End If
        End With
    End If

End Sub

Private Sub vsRpts_Click(Index As Integer)
On Error GoTo ErrSection:
    
''    vsRpts_AfterRowColChange Index
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.vsRpts.Click", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsRpts_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    LockWindowUpdate Me.hWnd
    
    m.lCurrentGrid = Index
    With vsRpts(Index)
        m.lReportID = CLng(ValOfText(.TextMatrix(.Row, RptRow(eRptGridRow_ReportID))))
        m.strReportName = .TextMatrix(.Row, RptRow(eRptGridRow_ReportName))
        m.strClassName = .TextMatrix(.Row, RptRow(eRptGridRow_ClassName))
        m.dLeftMargin = ValOfText(.TextMatrix(.Row, RptRow(eRptGridRow_LeftMargin)))
        m.dRightMargin = ValOfText(.TextMatrix(.Row, RptRow(eRptGridRow_RightMargin)))
        m.dTopMargin = ValOfText(.TextMatrix(.Row, RptRow(eRptGridRow_TopMargin)))
        m.dBottomMargin = ValOfText(.TextMatrix(.Row, RptRow(eRptGridRow_BottomMargin)))
        m.strOrientation = .TextMatrix(.Row, RptRow(eRptGridRow_Orientation))
        m.ChartType = CLng(ValOfText(.TextMatrix(.Row, RptRow(eRptGridRow_ChartType))))
        m.GridType = CLng(ValOfText(.TextMatrix(.Row, RptRow(eRptGridRow_GridType))))
        m.lOrder = CLng(ValOfText(.TextMatrix(.Row, RptRow(eRptGridRow_Order))))
    End With
    
    Select Case UCase(m.strClassName)
        Case "CRPTSUMMARY"
            Toolbar1.Tools("ID_Fields").Enabled = False
            Toolbar1.Tools("ID_Export").Visible = False
            cmdMoreBars.Visible = False
            cmdLessBars.Visible = False
            fraLegend.Visible = True
        
        Case "CRPTSUMMARYBY", "CRPTMMSTUDY"
            Toolbar1.Tools("ID_Fields").Enabled = True
            Toolbar1.Tools("ID_Export").Visible = True 'False
            cmdMoreBars.Visible = False
            cmdLessBars.Visible = False
            fraLegend.Visible = False

        Case "CRPTTRADES2"
            Toolbar1.Tools("ID_Fields").Enabled = True
            Toolbar1.Tools("ID_Export").Visible = True
            cmdMoreBars.Caption = ""
            'RH commented out cmdMoreBars.Picture = m.Images("kMoreBars").Picture
            cmdLessBars.Caption = ""
            'RH commented out cmdLessBars.Picture = m.Images("kLessBars").Picture
    
            cmdMoreBars.Visible = True
            cmdLessBars.Visible = True
            fraLegend.Visible = False
        
        Case Else
            Toolbar1.Tools("ID_Fields").Enabled = False
            Toolbar1.Tools("ID_Export").Visible = False
            cmdMoreBars.Visible = False
            cmdLessBars.Visible = False
            fraLegend.Visible = False
    
    End Select
    
    If m.lReportID <> 0 Then
        'Load the report options for this report
        m.Settings.Load m.lReportID
        
        ProcessRpt
        
        If Index = kMoneyMgt Then
            With vsImplements
                .Row = .Rows - 1
                .RowSel = .Rows - 1
                .ShowCell .Row, .Col
            End With
        End If
    End If
    
    ' Call a form resize so that the chart can be resized based on if the legend frame
    ' is visible or not...
    Form_Resize
    
ErrExit:
    LockWindowUpdate 0
    Exit Sub

ErrSection:
    RaiseError "frmReports.vsRpts.AfterRowColChange", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsSettings_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim strSettingName As String
    Dim strSettingType As String

    m.Settings.AfterEdit Row, Col
    
    'Custom:  If settings are changed which require trades to be reloaded
    'then set on flag
    'reloaded in the report class.
    strSettingName = m.Settings.Item(Str(Row), ensg_Name)
    Select Case strSettingName
        Case "FromDate", "ToDate", "TradeTypes", "IgnoretradesLE", _
             "IgnoreTradesGE", "IgnoreBigWins", "IgnoreBigLosses"
             'mReloadTrades = True
    End Select
    
    'If setting updated was a list box of combo then the report was already
    'generated in the ChangeEdit event.  Don't do it again.
    strSettingType = m.Settings.Item(Str(Row), ensg_Type)
    Select Case strSettingType
        Case "List", "ListAdd", "TableLookup"
            If strSettingName = "Delta" And vsSettings.TextMatrix(Row, Col) = "0" Then
                vsSettings.TextMatrix(Row, Col) = "Drawdown / 2"
                ProcessRpt
            ElseIf ValOfText(vsSettings.TextMatrix(Row, Col)) <> 0 Then
                ProcessRpt
            End If
        
        Case Else
            ProcessRpt
    
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.vsSettings.AfterEdit", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

Private Sub vsSettings_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    m.Settings.BeforeEdit Row, Col, Cancel

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "vsSettings.BeforeEdit", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit:

End Sub

Private Sub vsSettings_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim strValue As String              ' Current value to compare against
    
    strValue = vsSettings.TextMatrix(Row, Col)
    m.Settings.CellButtonClick Row, Col, Me
    If strValue <> vsSettings.TextMatrix(Row, Col) Then ProcessRpt

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "vsSettings.CellButtonClick", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit:

End Sub

Private Sub vsSettings_ChangeEdit()
On Error GoTo ErrSection:
    
    Dim strSettingName As String
    Dim strSettingType As String
    Dim bCancel As Boolean
    
    m.Settings.ChangeEdit bCancel
    
    If Not bCancel Then
        'Custom:  If settings are changed which require trades to be reloaded
        'then set on flag
        strSettingName = m.Settings.Item(Str(vsSettings.Row), ensg_Name)
        strSettingType = m.Settings.Item(Str(vsSettings.Row), ensg_Type)
        
        Select Case strSettingType
            Case "List", "TableLookup", "Boolean", "ListAdd"
                Select Case strSettingName
                    Case "FromDate", "ToDate", "TradeTypes", "IgnoreTradesLE", _
                         "IgnoreTradesGE", "IgnoreBigWins", "IgnoreBigLosses"
                         'mReloadTrades = True
                End Select
            
                ProcessRpt
        End Select
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.vsSettings.ChangeEdit", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit

End Sub

'Update the status bar message for the current option
Private Sub vsSettings_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    m.Settings.RowColChange

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.vsSettings.RowColChange", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit:

End Sub

Private Sub vsSettings_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsSettings.ComboCloseUp", eGDRaiseError_Show, m.strAppPath
    Resume ErrExit
    
End Sub

Private Sub vsSettings_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    m.Settings.ValidateEdit Row, Col, Cancel

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.vsSettings.ValidateEdit", eGDRaiseError_Show, m.strAppPath
    Cancel = True
    Resume ErrExit:

End Sub

Private Sub vsReport_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

    Dim lWidth As Long                  ' Amount to adjust the column width
    Dim lIndex As Long                  ' Index into a for loop
    
    ' if column being resized is the extended column,
    ' then make the next column bigger (instead of adjusting
    ' the extended column)
    If m.lCustomCol <> -1 Then ExtendCustomColumn vsReport
    If UCase(m.strClassName) = "CRPTSUMMARYBY" Then m.ReportObj.SaveCols

End Sub

Private Sub vsReport_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next

End Sub

Public Sub ImportTrades(ByVal strTradesPath$, ByVal lSystemNumber&, ByVal strSystemName$, _
                Optional ByVal hTblRules As Long = 0&, Optional ByVal hTblMarkets As Long = 0&)
On Error GoTo ErrSection:

    Dim dToDate As Double

    'Load trades...
    Set m.Trades = New cTrades
    m.strTradesPath = strTradesPath
    If hTblRules = 0& Or hTblMarkets = 0& Then
        m.Trades.Import strTradesPath, lSystemNumber, False, m.dbNav, m.dbNav
    Else
        m.Trades.ImportWithHandles strTradesPath, lSystemNumber, False, hTblRules, hTblMarkets
    End If

    If m.Trades.NumRecords - 1 <= 0 Then
        Err.Raise vbObjectError + 1000, , "No trade history was found.  Reports cannot be viewed."
    End If

    'Add system to Report Systems collection (System Navigator will always add 1 system only).
    'Note: this code assumes "System" has been instanced..
    If m.Systems Is Nothing Then Set m.Systems = New cSystemsList
    With m.Systems
        .Num(1, ensy_SystemNumber) = lSystemNumber
        .Item(1, ensy_SystemName) = strSystemName
        .Item(1, ensy_Symbol) = m.Trades.ItemHdr(1, enth_Symbol)
        .Num(1, ensy_TickMove) = m.Trades.ItemHdr(1, enth_TickMove)
        .Num(1, ensy_TickValue) = m.Trades.ItemHdr(1, enth_TickValue)
        .Num(1, ensy_MinMoveInTicks) = m.Trades.ItemHdr(1, enth_TickMinMove)
        .Num(1, ensy_DefaultUnits) = m.Trades.ItemHdr(1, enth_DefaultUnits)
    End With

    'Set todate to 11:59pm if time not specified
    dToDate = m.Trades.Num(m.Trades.NumRecords - 1, entd_TradeDate)
    If InStr(1, dToDate, ".") <= 0 Then dToDate = dToDate + 0.9999
    m.dToDate = dToDate
    m.dFromDate = m.Trades.Num(1, entd_TradeDate)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.ImportTrades", eGDRaiseError_Raise, m.strAppPath

End Sub

Public Sub ImportMultipleTradeFiles(ByVal hFiles&, Optional ByVal hTblRules As Long = 0&)
On Error GoTo ErrSection:

    Dim dToDate As Double
    Dim lIndex As Long

    ' Load trades...
    Set m.Trades = New cTrades
    'm.strTradesPath = strTradesPath
    If hTblRules = 0& Then
        m.Trades.ImportMultiple hFiles, m.dbNav, m.dbNav
    Else
        m.Trades.ImportMultipleWithHandles hFiles, hTblRules
    End If

    If m.Trades.NumRecords - 1 <= 0 Then
        Err.Raise vbObjectError + 1000, , "No trade history was found.  Reports cannot be viewed."
    End If

    ' Add systems to Report Systems collection...
    If m.Systems Is Nothing Then Set m.Systems = New cSystemsList
    With m.Systems
        .NumRecords = m.Trades.HeaderNumRecords
        
        For lIndex = 1 To .NumRecords - 1
            .Num(lIndex, ensy_SystemNumber) = m.Trades.HeaderNum(enth_SystemNumber, lIndex)
            .Item(lIndex, ensy_SystemName) = m.Trades.HeaderStr(enth_SystemName, lIndex)
            .Item(lIndex, ensy_Symbol) = m.Trades.HeaderStr(enth_Symbol, lIndex)
            .Num(lIndex, ensy_TickMove) = m.Trades.HeaderNum(enth_TickMove, lIndex)
            .Num(lIndex, ensy_TickValue) = m.Trades.HeaderNum(enth_TickValue, lIndex)
            .Num(lIndex, ensy_MinMoveInTicks) = m.Trades.HeaderNum(enth_TickMinMove, lIndex)
            .Num(lIndex, ensy_DefaultUnits) = m.Trades.HeaderNum(enth_DefaultUnits, lIndex)
        Next lIndex
    End With

    'Set todate to 11:59pm if time not specified
    dToDate = m.Trades.Num(gdGetNum(m.Trades.SortHandle, m.Trades.NumRecords - 1), entd_TradeDate)
    If dToDate = Int(dToDate) Then dToDate = dToDate + 0.9999
    m.dToDate = dToDate
    m.dFromDate = m.Trades.Num(gdGetNum(m.Trades.SortHandle, 1), entd_TradeDate)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.ImportMultipleTradeFiles", eGDRaiseError_Raise, m.strAppPath

End Sub

'Force this report to run in single unit mode.
Public Sub SetToSingleUnit(Optional pUseBeginBalanceRowZero As Boolean = True)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                  ' Index into a for loop
    Dim lSysArrayIndex As Long          ' Index into the Systems Array
    Dim lOpenPositions As Long          ' Number of open positions
    
    ' Headers for Arrays in the Trades table
    Dim hSysNbr As Long
    Dim hUnits As Long
    Dim hProfit As Long
    Dim hTotalProfit As Long
    Dim hOpenTradesTotal As Long
    Dim hSkip As Long
    Dim hSkipRpt As Long
    Dim hAccountBalance As Long
    Dim hSignalType As Long
    Dim hSorted As Long
    
    hSysNbr = m.Trades.FieldHandle(entd_SysNbr)
    hUnits = m.Trades.FieldHandle(entd_Units)
    hProfit = m.Trades.FieldHandle(entd_Profit)
    hTotalProfit = m.Trades.FieldHandle(entd_TotalProfit)
    hAccountBalance = m.Trades.FieldHandle(entd_AccountBalance)
    hOpenTradesTotal = m.Trades.FieldHandle(entd_OpenTradesTotal)
    hSkip = m.Trades.FieldHandle(entd_Skip)
    hSkipRpt = m.Trades.FieldHandle(entd_SkipRpt)
    hSignalType = m.Trades.FieldHandle(entd_SignalType)
    hSorted = m.Trades.SortHandle
    
    For lIndex2 = 1 To m.Trades.NumRecords - 1
        lIndex = gdGetNum(hSorted, lIndex2)
    
        lSysArrayIndex = SystemArrayInd(gdGetNum(hSysNbr, lIndex))
        If lSysArrayIndex > 0 Then
            gdSetNum hUnits, lIndex, m.Systems.Num(lSysArrayIndex, ensy_DefaultUnits)
        End If
        
        gdSetNum hTotalProfit, lIndex, gdGetNum(hProfit, lIndex)
        
        'Ignore first row since it has beginning balance.
        If lIndex = 1 Then
            If pUseBeginBalanceRowZero Then
                gdSetNum hAccountBalance, lIndex, gdGetNum(hAccountBalance, 0)
            Else
                gdSetNum hAccountBalance, lIndex, 0
            End If
            
            If gdGetNum(hSkip, lIndex) = 0 And gdGetNum(hSkipRpt, lIndex) = 0 Then
                lOpenPositions = 1
                gdSetNum hOpenTradesTotal, lIndex, 1
            Else
                gdSetNum hOpenTradesTotal, lIndex, 0
            End If
        Else
            If gdGetNum(hUnits, lIndex) > 0 And _
               gdGetNum(hSkip, lIndex) = 0 And _
               gdGetNum(hSkipRpt, lIndex) = 0 Then
                If gdGetNum(hSignalType, lIndex) = gEntrySignal Then
                    lOpenPositions = lOpenPositions + 1
                    gdSetNum hAccountBalance, lIndex, gdGetNum(hAccountBalance, lIndex - 1)
                    gdSetNum hOpenTradesTotal, lIndex, lOpenPositions
                Else
                    lOpenPositions = lOpenPositions - 1
                    gdSetNum hAccountBalance, lIndex, _
                        gdGetNum(hAccountBalance, lIndex - 1) + _
                        gdGetNum(hTotalProfit, lIndex)
                    gdSetNum hOpenTradesTotal, lIndex, lOpenPositions
                End If
                
            Else
                gdSetNum hOpenTradesTotal, lIndex, lOpenPositions
                gdSetNum hAccountBalance, lIndex, gdGetNum(hAccountBalance, lIndex - 1)
            End If
        End If
        
    Next lIndex2
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.SetToSingleUnit", eGDRaiseError_Raise, m.strAppPath

End Sub

'This routine find a specific setting value
Public Function OptionValue(ByVal strSettingName As String) As Variant
On Error GoTo ErrSection:
    
    OptionValue = m.Settings.Item(strSettingName)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmReports.OptionValue", eGDRaiseError_Raise, m.strAppPath

End Function

'Call this to set Filter trades based on settings designed to reloaded
'the trades by setting the skip flags.
Public Sub FilterTrades()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim bContinue As Boolean
    Dim alLargest As cGdArray
    Dim alSmallest As cGdArray
    Dim dLargest As Double
    Dim dSmallest As Double
    Dim bPos As Byte
    Dim dEntryDate As Double
    Dim dProfit As Double
    Dim dAccountBalance As Double
    
    Dim adProfit As cGdArray
    Dim alProfitIndex As cGdArray
    
    ' Handles to arrays in the Trades table
    Dim hSignalType As Long
    Dim hPosition As Long
    Dim hEntryExitPtr As Long
    Dim hProfit As Long
    Dim hTradeDate As Long
    Dim hSkip As Long
    Dim hShow As Long
    
    'Filter variables
    Dim FilterTradesType As String
    Dim FilterPos As Byte
    Dim FilterFromDate As Double
    Dim FilterToDate As Double
    Dim FilterIgnoreLE As Double
    Dim FilterIgnoreGE As Double
    Dim FilterIgnoreBigWins As Integer
    Dim FilterIgnoreBigLosses As Integer
    
    If m.Trades.NumRecords - 1 = 0 Then Exit Sub
    
    'Get Filter options
    FilterTradesType = OptionValue("TradeTypes")
    If FilterTradesType = "All Trades" Then
        FilterPos = 2
    Else
        If FilterTradesType = "Long Trades only" Then
            FilterPos = 1
        Else
            FilterPos = 0
        End If
    End If
    FilterFromDate = OptionValue("FromDate")
    FilterToDate = OptionValue("ToDate")
    FilterIgnoreLE = OptionValue("IgnoreTradesLE")
    FilterIgnoreGE = OptionValue("IgnoreTradesGE")
    FilterIgnoreBigWins = OptionValue("IgnoreBigWins")
    FilterIgnoreBigLosses = OptionValue("IgnoreBigLosses")
    
    'Get Trade handles
    hSignalType = m.Trades.FieldHandle(entd_SignalType)
    hPosition = m.Trades.FieldHandle(entd_Position)
    hEntryExitPtr = m.Trades.FieldHandle(entd_EntryExitPtr)
    hProfit = m.Trades.FieldHandle(entd_Profit)
    hTradeDate = m.Trades.FieldHandle(entd_TradeDate)
    hSkip = m.Trades.FieldHandle(entd_Skip)
    hShow = m.Trades.FieldHandle(entd_Show)
    
    'Allocate room to store alLargest and alSmallest trades
    Set alLargest = New cGdArray
    Set alSmallest = New cGdArray
    alLargest.Create eGDARRAY_Longs
    alSmallest.Create eGDARRAY_Longs
    dLargest = -999999
    dSmallest = 999999
    
    Set adProfit = New cGdArray
    Set alProfitIndex = New cGdArray
    adProfit.Create eGDARRAY_Doubles
    alProfitIndex.Create eGDARRAY_Longs
        
    For lIndex = 1 To m.Trades.NumRecords - 1
        
        'Continue for Entry types only...
        If gdGetNum(hSignalType, lIndex) = gEntrySignal Then
        
            bContinue = True
            bPos = gdGetNum(hPosition, lIndex)
            dEntryDate = gdGetNum(hTradeDate, lIndex)
            dProfit = gdGetNum(hProfit, gdGetNum(hEntryExitPtr, lIndex))
            
            'Trades to load filter
            If FilterPos <> bPos And FilterPos <> 2 Then
                bContinue = False
            End If
            
            'Entry Date filter
            If FilterFromDate <> 0 Then
                If dEntryDate < FilterFromDate Or _
                   dEntryDate > FilterToDate Then
                    bContinue = False
                End If
            End If
                 
            'Profit extremes filter...
            If dProfit < FilterIgnoreLE Or dProfit > FilterIgnoreGE Then
                bContinue = False
            End If
            
            'Keep track of position of largest winners and losers
            'If FilterIgnoreBigWins > 0 Then
            '    If dProfit >= dLargest Then
            '        dLargest = dProfit
            '        alLargest.Add lIndex
            '    End If
            'End If
            
            'If FilterIgnoreBigLosses > 0 Then
            '    If dProfit <= dSmallest Then
            '        dSmallest = dProfit
            '        alSmallest.Add lIndex
            '    End If
            'End If
            
            adProfit.Add dProfit
            alProfitIndex.Add lIndex
         
            'If this is called from the engine, some skip flags may already
            'be on.  Don't set these to 0 (Skip(lIndex)<>1)
            If bContinue And gdGetNum(hSkip, lIndex) <> 1 Then
                gdSetNum hSkip, lIndex, 0
                gdSetNum hShow, lIndex, 1
                gdSetNum hSkip, gdGetNum(hEntryExitPtr, lIndex), 0
                gdSetNum hShow, gdGetNum(hEntryExitPtr, lIndex), 1
            Else
                gdSetNum hSkip, lIndex, 1
                gdSetNum hShow, lIndex, 0
                gdSetNum hSkip, gdGetNum(hEntryExitPtr, lIndex), 1
                gdSetNum hShow, gdGetNum(hEntryExitPtr, lIndex), 0
            End If
        End If
    Next lIndex
    
    gdSortAsIndex alLargest.ArrayHandle, adProfit.ArrayHandle, 1, eGdSort_Default, 0, adProfit.Size - 1
    
    'Set the skip flag for the Maximum and Minimum profit trades
    Dim Ind As Long
    For lIndex = 1 To FilterIgnoreBigWins
        'If alLargest.Size - 1 = lIndex Then Exit For
        'Ind = alLargest.Num(alLargest.Size - 1 - lIndex + 1)
        Ind = alProfitIndex(alLargest(alLargest.Size - lIndex))
        gdSetNum hSkip, Ind, 1
        gdSetNum hSkip, gdGetNum(hEntryExitPtr, Ind), 1
        gdSetNum hShow, Ind, 0
        gdSetNum hShow, gdGetNum(hEntryExitPtr, Ind), 0
    Next lIndex
    For lIndex = 1 To FilterIgnoreBigLosses
        'If alSmallest.Size - 1 = lIndex Then Exit For
        'Ind = alSmallest.Num(alSmallest.Size - 1 - lIndex + 1)
        Ind = alProfitIndex(alLargest(lIndex - 1))
        gdSetNum hSkip, Ind, 1
        gdSetNum hSkip, gdGetNum(hEntryExitPtr, Ind), 1
        gdSetNum hShow, Ind, 0
        gdSetNum hShow, gdGetNum(hEntryExitPtr, Ind), 0
    Next lIndex
    
    'Recalculate Account Balances
    m.Trades.CalcAccBal
    
ErrExit:
    Set alLargest = Nothing
    Set alSmallest = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmReports.FilterTrades", eGDRaiseError_Raise, m.strAppPath

End Sub

Public Function SystemArrayInd(ByVal lSystemNumber As Long) As Long
On Error GoTo ErrSection:

    Dim lIndex As Integer               ' Index into a for loop
    
    For lIndex = 0 To m.Systems.NumRecords - 1
        If m.Systems.Num(lIndex, ensy_SystemNumber) = lSystemNumber Then
            SystemArrayInd = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmReports.SystemArrayInd", eGDRaiseError_Raise, m.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExtendCustomColumn(vsGrid As VSFlexGrid)
On Error GoTo ErrSection:
    
    Dim lTotal As Long                  ' New width of the extended column
    Dim lIndex As Long                  ' Index into a for loop
        
    With vsGrid
        .ColHidden(m.lCustomCol) = True
        .Redraw = flexRDBuffered '(so .ClientWidth will be correct)
        .Redraw = flexRDNone
        lTotal = 0 * Screen.TwipsPerPixelX
        For lIndex = 0 To .Cols - 1
            If Not .ColHidden(lIndex) Then
                lTotal = lTotal + .ColWidth(lIndex)
            End If
        Next
        lTotal = .ClientWidth - lTotal
        .ColHidden(m.lCustomCol) = False
        If lTotal > 0 Then
            If lTotal >= m.lMinColWidth Then
                .ColWidth(m.lCustomCol) = lTotal
            Else
                .ColWidth(m.lCustomCol) = m.lMinColWidth
            End If
        Else
            .ColWidth(m.lCustomCol) = m.lMinColWidth
        End If
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.ExtendCustomColumn", eGDRaiseError_Raise, m.strAppPath

End Sub

Private Sub vsTabs_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    ' For some reason, the frame on the BarChart tab is not being resized
    ' automatically unless we do a Form_Resize here.  11/8/2002 DAJ
    Form_Resize
    
    'If NewTab = 2 Or NewTab = 3 Then
    '    Toolbar1.Tools("ID_Export").Visible = True
    'Else
    '    Toolbar1.Tools("ID_Export").Visible = False
    'End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.vsTabs.Switch", eGDRaiseError_Show, m.strAppPath
End Sub

Public Sub ShowDateInChart(ByVal dDate As Double, ByVal dPrice As Double, ByVal lHeader As Long)
On Error GoTo ErrSection:

    Dim strLineToAdd As String
    Dim astrFile As New cGdArray
    Dim strFileName As String
    
    '(TLB 11/10/2003: now do whether from a chart or not)
    'If m.hChartHwnd <> -1 Then
        strFileName = AddSlash(m.strAppPath) & "Doit.NOW"
        
        ' If the periodicity is Days or greater, we don't need the time...
        'If GetPeriodicity(m.Trades.ItemHdr(1, enth_BarTimeFrame)) >= ePRD_Days + 1 Then
        '    dDate = Int(dDate)
        'End If
        
        If m.Trades.HeaderNum(enth_IntraDaySystem, lHeader) = 0 Then
            ' if not an intraday system, just pass the date
            dDate = Int(dDate)
        ElseIf g.bShowInLocalTimeZone Then
            ' if converted to local time, convert back
            dDate = ConvertTimeZone(dDate, "", m.Trades.HeaderStr(enth_TimeZoneInfo, lHeader))
        End If
        
        strLineToAdd = "CHARTDATE=" & Str(m.hChartHwnd) & vbTab & Str(dDate) & vbTab & _
                m.Trades.ItemHdr(lHeader, enth_Symbol) & vbTab & _
                m.Trades.ItemHdr(lHeader, enth_BarTimeFrame) & vbTab & _
                m.strPortOrSystemName & vbTab & DateFormat(dDate) & vbTab & _
                Str(dPrice)
                
        astrFile.Create eGDARRAY_Strings
        astrFile.Add strLineToAdd
        astrFile.ToFile strFileName, True
    'End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.ShowDateInChart", eGDRaiseError_Raise, m.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportToTradeIT
'' Description: Export the trades file to a format to be read by TradeIT!
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExportToTradeIT(ByVal strFileName As String)
On Error GoTo ErrSection:

    Dim astrOutput As New cGdArray      ' Array for output purposes
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim hSorted As Long                 ' Handle to the sorted array in the trades
    Dim astrTrade As New cGdArray       ' String to add to the output array
    Dim dEntryDate As Double            ' Entry Date for the Trade
    Dim dEntryPrice As Double           ' Entry Price for the Trade
    Dim dExitDate As Double             ' Exit Date for the Trade
    Dim dExitPrice As Double            ' Exit Price for the Trade
    Dim lHeld As Long                   ' Number of days the position was held
    
    Screen.MousePointer = vbHourglass
    
    astrOutput.Create eGDARRAY_Strings
    astrTrade.Create eGDARRAY_Strings
    
    With m.Trades
        hSorted = .SortHandle
        
        astrOutput.Add "Ticker,Held,Entry Date,Entry Price,Exit Date,Exit Price,Profit,DrawDn%,Max Pft%,AROI%,Profit%,SPX%"
        
        For lIndex2 = 1 To m.Trades.NumRecords - 1
            lIndex = gdGetNum(hSorted, lIndex2)
            
            If .DataNum(entd_SignalType, lIndex) = gEntrySignal Then
                dEntryDate = .DataNum(entd_TradeDate, lIndex)
                dEntryPrice = .DataNum(entd_Price, lIndex)
                dExitDate = .DataNum(entd_TradeDate, .DataNum(entd_EntryExitPtr, lIndex))
                dExitPrice = .DataNum(entd_Price, .DataNum(entd_EntryExitPtr, lIndex))
                lHeld = Int(dExitDate) - Int(dEntryDate)
            
                astrTrade.Clear
                
                astrTrade.Add .Symbol(.DataNum(entd_SymbolIndex, lIndex))
                astrTrade.Add Str(lHeld)
                astrTrade.Add Format(dEntryDate, "MM/DD/YY")
                astrTrade.Add Str(dEntryPrice)
                astrTrade.Add Format(dExitDate, "MM/DD/YY")
                astrTrade.Add Str(dExitPrice)
                astrTrade.Add Str(dExitPrice - dEntryPrice)
                
                astrOutput.Add astrTrade.JoinFields(",")
            End If
        Next lIndex2
    End With
    
    astrOutput.ToFile strFileName

ErrExit:
    Screen.MousePointer = vbDefault
    Set astrOutput = Nothing
    Exit Sub
    
ErrSection:
    Screen.MousePointer = vbDefault
    Set astrOutput = Nothing
    RaiseError "frmReports.ExportToTradeIT", eGDRaiseError_Raise, m.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportToCSV
'' Description: Export the current grid to CSV format
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExportToCSV()
On Error GoTo ErrSection:

    Dim strFileName As String           ' Filename to export the grid to
    
    strFileName = CommonDialogFile(CommonDialog1, True, "TXT Files (*.TXT)|*.txt", AddSlash(m.strAppPath))
    If Len(strFileName) > 0 Then
        m.ReportObj.ExportToCSV strFileName
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmReports.ExportToCSV", eGDRaiseError_Raise, m.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditEquityFilter
'' Description: Allow the user to edit the equity filter information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditEquityFilter()
On Error GoTo ErrSection:

    If frmEquityFilter.ShowMe(m.EquityFilter, m.Images("kDollarLine").Picture) Then
        SetIniFileProperty "MaPeriod", m.EquityFilter.MovingAveragePeriod, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI"
        SetIniFileProperty "MaType", m.EquityFilter.MovingAverageType, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI"
        If m.EquityFilter.EquityFilterOn Then
            SetIniFileProperty "FilterOn", True, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI"
        Else
            SetIniFileProperty "FilterOn", False, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI"
        End If
        SetIniFileProperty "FilterType", m.EquityFilter.EquityFilterMode, "EquityFilter", AddSlash(g.strAppPath) & "Reports.INI"
        
        m.lMovAvgPeriod = m.EquityFilter.MovingAveragePeriod
        m.strMAType = m.EquityFilter.MovingAverageType
                
        ProcessRpt
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.EditEquityFilter", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TakeNextTrade
'' Description: Should the user take the next trade based on the equity filter?
'' Inputs:      Take Next Trade?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TakeNextTrade(ByVal nTakeNextTrade As eGDTakeNextTradeValue)
On Error GoTo ErrSection:

    m.nTakeNextTrade = nTakeNextTrade
    Select Case nTakeNextTrade
        Case eGDTakeNextTrade_Yes
            imgStopGo.Picture = m.Images("kEquityGo").Picture
        Case eGDTakeNextTrade_No
            imgStopGo.Picture = m.Images("kEquityStop").Picture
        Case Else
            imgStopGo.Picture = m.Images("kBlank").Picture
    End Select

    lblEquityFilter.Caption = m.EquityFilter.EnglishString(m.nTakeNextTrade)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmReports.TakeNextTrade", , g.strAppPath
    
End Sub

Private Function MinMaxInteger(ByVal dValue#, ByVal iMin&, ByVal iMax&, Optional ByVal iDefault& = 0) As Long

    dValue = Round(dValue)
    If iDefault > 0 And dValue <= 0 Then
        dValue = iDefault
    ElseIf dValue < iMin Then
        dValue = iMin
    ElseIf dValue > iMax Then
        dValue = iMax
    End If
    MinMaxInteger = dValue

End Function

' returns gdTable of Profits and RiskAmt (columns) for each trade (rows)
Private Function GetProfits() As cGdTable

    Dim iNumRecs&, iNumProfits&, iRec&, iTrade&, iHdrIndex&
    Dim d#, dProfit#, dRisk#
    Dim hSorted&, hSkip&, hSkipRpt&, hSignalType&, hTotalProfit&, hHeaderIndex&
    Dim tProfits As New cGdTable
    Dim aMaxLosses As New cGdArray
    Dim aAvgLosses As New cGdArray
    Dim aLossCounts As New cGdArray

    With m.Trades
        iNumRecs = .NumRecords
        hSignalType = .FieldHandle(entd_SignalType)
        hSkip = .FieldHandle(entd_Skip)
        hSkipRpt = .FieldHandle(entd_SkipRpt)
        'hEntryExitPtr = .FieldHandle(entd_EntryExitPtr)
        'hTradeDate = .FieldHandle(entd_TradeDate)
        'hMaxLoss = .FieldHandle(entd_MaxLoss)
        hTotalProfit = .FieldHandle(entd_TotalProfit)
        'hTradeNbr = .FieldHandle(entd_TradeNbr)
        'hUnfilteredEquity = .FieldHandle(entd_UnfilteredEquity)
        'hEquityMA = .FieldHandle(entd_EquityMA)
        'hFilteredEquity = .FieldHandle(entd_FilteredEquity)
        hHeaderIndex = .FieldHandle(entd_HeaderIndex)
        hSorted = .SortHandle
    End With
    
    ' create table of about the right size
    tProfits.CreateField eGDARRAY_Doubles, 0, "Profits", 0
    tProfits.CreateField eGDARRAY_Doubles, 1, "RiskAmt", 0
    tProfits.NumRecords = iNumRecs / 2
    
    ' create arrays to help calc largest and average losses for each system
    aMaxLosses.Create eGDARRAY_Doubles, 0, 0
    aAvgLosses.Create eGDARRAY_Doubles, 0, 0
    aLossCounts.Create eGDARRAY_Longs, 0, 0
    
    ' get profit for each trade
    For iRec = 1 To iNumRecs - 1
        iTrade = gdGetNum(hSorted, iRec)
        If gdGetNum(hSkip, iTrade) = 0 And gdGetNum(hSkipRpt, iTrade) = 0 Then
            If gdGetNum(hSignalType, iTrade) = 1 Then
                dProfit = gdGetNum(hTotalProfit, iTrade)
                iHdrIndex = gdGetNum(hHeaderIndex, iTrade)
                
                ' store largest and avg loss for each system
                If dProfit < 0 Then
                    If dProfit < aMaxLosses.Num(iHdrIndex) Then
                        aMaxLosses.Num(iHdrIndex) = dProfit
                    End If
                    aAvgLosses.Num(iHdrIndex) = aAvgLosses.Num(iHdrIndex) + dProfit
                    aLossCounts.Num(iHdrIndex) = aLossCounts.Num(iHdrIndex) + 1
                End If
                
                ' store $Profit, and HdrIndex for now (will be replaced with $RiskAmt later)
                tProfits.Num(0, iNumProfits) = dProfit
                tProfits.Num(1, iNumProfits) = iHdrIndex
                iNumProfits = iNumProfits + 1
            End If
        End If
    Next
    tProfits.NumRecords = iNumProfits ' exact size
    
    ' now calc the average loss for each system
    For iRec = 0 To aAvgLosses.Size - 1
        d = aLossCounts.Num(iRec)
        If d > 1 Then
            aAvgLosses.Num(iRec) = aAvgLosses.Num(iRec) / d
        End If
    Next
    
    ' now determine the actual $RiskAmt for each trade
    For iRec = 0 To tProfits.NumRecords - 1
        dRisk = 0
        iHdrIndex = tProfits.Num(1, iRec)
        
        ' get $Risk based on for this system (use StopLoss if exists)
        d = gdGetTableNum(m.Trades.HdrTableHandle, enth_LongStopLoss, iHdrIndex)
        If d > 0 And -d < dRisk Then
            dRisk = -d
        End If
        d = gdGetTableNum(m.Trades.HdrTableHandle, enth_ShortStopLoss, iHdrIndex)
        If d > 0 And -d < dRisk Then
            dRisk = -d
        End If

        ' if there is no identifiable StopLoss, then just use MaxLoss for this system
        If dRisk = 0 Then
            dRisk = aMaxLosses.Num(iHdrIndex)
        ' but if a StopLoss, make sure it's not "smaller" than the AvgLoss
        ' (remember both are negative values)
        ElseIf dRisk > aAvgLosses.Num(iHdrIndex) Then
            dRisk = aAvgLosses.Num(iHdrIndex)
        End If
        
        ' store the $RiskAmt for each trade
        tProfits.Num(1, iRec) = dRisk
    Next
    
    Set aMaxLosses = Nothing
    Set aAvgLosses = Nothing
    Set aLossCounts = Nothing
    
    Set GetProfits = tProfits

End Function

Public Property Get AvgTradesPerYear() As Double
    AvgTradesPerYear = m.dAvgTradesPerYear
End Property

Public Property Let AvgTradesPerYear(ByVal dAvgTradesPerYear As Double)
    m.dAvgTradesPerYear = dAvgTradesPerYear
    Me.txtDrawdownTrades.Text = Str(Round(m.dAvgTradesPerYear))
End Property

Private Sub SetRiskTrades()

    Dim n&, Y&
    Y = Val(cboRiskYears.Text)
    n = Round(m.dAvgTradesPerYear * Y)
    If Y > 1 Then
        lblRiskTrades.Caption = "years (" & Str(n) & " trades)."
    Else
        lblRiskTrades.Caption = "year (" & Str(n) & " trades)."
    End If
    lblRiskTrades.Tag = Str(n)

End Sub

