VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmToolbox 
   Caption         =   "Toolbox"
   ClientHeight    =   5730
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10575
   Icon            =   "frmToolbox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3675
      Left            =   9300
      TabIndex        =   9
      Top             =   480
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
      Caption         =   "frmToolbox.frx":014A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmToolbox.frx":0176
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmToolbox.frx":0196
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   435
         Left            =   0
         TabIndex        =   8
         Top             =   1080
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":01B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":01DE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":01FE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdImport 
         Height          =   435
         Left            =   0
         TabIndex        =   6
         Top             =   1320
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":021A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":0248
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":0268
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   435
         Left            =   0
         TabIndex        =   5
         Top             =   3180
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":0284
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":02B2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":02D2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   2640
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":02EE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":0318
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":0338
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   435
         Left            =   0
         TabIndex        =   3
         Top             =   2280
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":0354
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":037C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":039C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddCopy 
         Height          =   435
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":03B8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":03EA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":040A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   435
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":0426
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":044E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":046E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdInfo 
         Height          =   435
         Left            =   0
         TabIndex        =   7
         Top             =   1620
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":048A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":04B4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":04D4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNewDLL 
         Height          =   435
         Left            =   0
         TabIndex        =   10
         Top             =   1920
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":04F0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":0520
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":0540
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRename 
         Height          =   435
         Left            =   0
         TabIndex        =   43
         Top             =   780
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":055C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":058A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":05AA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdExport 
         Height          =   435
         Left            =   0
         TabIndex        =   46
         Top             =   2880
         Width           =   1155
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
         Caption         =   "frmToolbox.frx":05C6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbox.frx":05F4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":0614
         RightToLeft     =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab vsTypeTabs 
      Height          =   5355
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9446
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
      FrontTabForeColor=   -2147483635
      Caption         =   "Symbol &Groups|Cri&teria|F&ilters|&Functions|&Rules|&Strategies|Strategy &Baskets|&Libraries"
      Align           =   0
      Appearance      =   1
      CurrTab         =   7
      FirstTab        =   2
      Style           =   3
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
      Picture(0)      =   "frmToolbox.frx":0630
      Picture(1)      =   "frmToolbox.frx":0BCA
      Picture(2)      =   "frmToolbox.frx":1164
      Picture(3)      =   "frmToolbox.frx":16FE
      Picture(4)      =   "frmToolbox.frx":1C98
      Picture(5)      =   "frmToolbox.frx":1DF2
      Picture(6)      =   "frmToolbox.frx":238C
      Picture(7)      =   "frmToolbox.frx":2726
      Begin HexUniControls.ctlUniFrameWL fraStrategyBaskets 
         Height          =   4935
         Left            =   -9690
         TabIndex        =   44
         Top             =   375
         Width           =   9045
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
         Caption         =   "frmToolbox.frx":2CC0
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmToolbox.frx":2CEC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":2D0C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRichTextBoxXP txtPreview 
            Height          =   915
            Index           =   6
            Left            =   120
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   3840
            Visible         =   0   'False
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1614
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmToolbox.frx":2D28
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
            Tip             =   "frmToolbox.frx":2D48
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":2D68
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
         Begin VSFlex7LCtl.VSFlexGrid fgStrategyBaskets 
            Height          =   4275
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   6555
            _cx             =   11562
            _cy             =   7541
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
         Begin HexUniControls.ctlUniLabelXP Label8 
            Height          =   255
            Left            =   120
            Top             =   60
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
            Caption         =   "frmToolbox.frx":2D84
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmToolbox.frx":2E60
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":2E80
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraLibraries 
         Height          =   4935
         Left            =   45
         TabIndex        =   34
         Top             =   375
         Width           =   9045
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
         Caption         =   "frmToolbox.frx":2E9C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmToolbox.frx":2EC8
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":2EE8
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRichTextBoxXP txtPreview 
            Height          =   915
            Index           =   7
            Left            =   120
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   3840
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   1614
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmToolbox.frx":2F04
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
            Tip             =   "frmToolbox.frx":2F24
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":2F44
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
         Begin VSFlex7LCtl.VSFlexGrid fgLibraries 
            Height          =   3315
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   6435
            _cx             =   11351
            _cy             =   5847
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
         Begin HexUniControls.ctlUniLabelXP Label7 
            Height          =   255
            Left            =   120
            Top             =   60
            Width           =   6855
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
            Caption         =   "frmToolbox.frx":2F60
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmToolbox.frx":3006
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":3026
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSystems 
         Height          =   4935
         Left            =   -9990
         TabIndex        =   31
         Top             =   375
         Width           =   9045
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
         Caption         =   "frmToolbox.frx":3042
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmToolbox.frx":306E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":308E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRichTextBoxXP txtPreview 
            Height          =   915
            Index           =   5
            Left            =   120
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   3840
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   1614
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmToolbox.frx":30AA
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
            Tip             =   "frmToolbox.frx":30CA
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":30EA
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
         Begin VSFlex7LCtl.VSFlexGrid fgSystems 
            Height          =   3315
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   6435
            _cx             =   11351
            _cy             =   5847
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
         Begin HexUniControls.ctlUniLabelXP Label6 
            Height          =   255
            Left            =   120
            Top             =   60
            Width           =   6855
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
            Caption         =   "frmToolbox.frx":3106
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmToolbox.frx":31D4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":31F4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraRules 
         Height          =   4935
         Left            =   -10290
         TabIndex        =   17
         Top             =   375
         Width           =   9045
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
         Caption         =   "frmToolbox.frx":3210
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmToolbox.frx":323C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":325C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraRuleFilter 
            Height          =   1050
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Width           =   6495
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
            Caption         =   "frmToolbox.frx":3278
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmToolbox.frx":32BA
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":32DA
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraRuleFilters 
               Height          =   315
               Left            =   240
               TabIndex        =   25
               Top             =   600
               Width           =   4815
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
               Caption         =   "frmToolbox.frx":32F6
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmToolbox.frx":3322
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmToolbox.frx":3342
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniComboImageXP cboLibrary 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   28
                  Top             =   0
                  Width           =   2295
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
                  Tip             =   "frmToolbox.frx":335E
                  Sorted          =   0   'False
                  HScroll         =   0   'False
                  RoundedBorders  =   -1  'True
                  IconDim         =   16
                  MousePointer    =   0
                  MouseIcon       =   "frmToolbox.frx":337E
                  DropDownOnTextClick=   -1  'True
                  DropDownWidth   =   -1
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniCheckXP chkLibrary 
                  Height          =   255
                  Left            =   0
                  TabIndex        =   27
                  Top             =   30
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
                  Caption         =   "frmToolbox.frx":339A
                  Enabled         =   -1  'True
                  Align           =   0
                  CheckBackColor  =   -2147483643
                  CheckForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmToolbox.frx":33CA
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmToolbox.frx":33EA
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniCheckXP chkFavorites 
                  Height          =   255
                  Left            =   3360
                  TabIndex        =   26
                  Top             =   30
                  Width           =   1635
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
                  Caption         =   "frmToolbox.frx":3406
                  Enabled         =   -1  'True
                  Align           =   0
                  CheckBackColor  =   -2147483643
                  CheckForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmToolbox.frx":3442
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmToolbox.frx":34F2
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
            Begin HexUniControls.ctlUniFrameWL fraRuleTypes 
               Height          =   255
               Left            =   240
               TabIndex        =   19
               Top             =   240
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
               Caption         =   "frmToolbox.frx":350E
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmToolbox.frx":353A
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmToolbox.frx":355A
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniCheckXP chkTreeRules 
                  Height          =   255
                  Left            =   0
                  TabIndex        =   47
                  Top             =   0
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
                  Caption         =   "frmToolbox.frx":3576
                  Enabled         =   -1  'True
                  Align           =   0
                  CheckBackColor  =   -2147483643
                  CheckForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmToolbox.frx":359E
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmToolbox.frx":35BE
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optShortExit 
                  Height          =   255
                  Left            =   4080
                  TabIndex        =   24
                  Top             =   0
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
                  Caption         =   "frmToolbox.frx":35DA
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmToolbox.frx":360E
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmToolbox.frx":362E
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optLongExit 
                  Height          =   255
                  Left            =   2940
                  TabIndex        =   23
                  Top             =   0
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
                  Caption         =   "frmToolbox.frx":364A
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmToolbox.frx":367C
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmToolbox.frx":369C
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optShort 
                  Height          =   255
                  Left            =   2100
                  TabIndex        =   22
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
                  Caption         =   "frmToolbox.frx":36B8
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmToolbox.frx":36E2
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmToolbox.frx":3702
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optLong 
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   21
                  Top             =   0
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
                  Caption         =   "frmToolbox.frx":371E
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmToolbox.frx":3746
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmToolbox.frx":3766
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optAll 
                  Height          =   255
                  Left            =   720
                  TabIndex        =   20
                  Top             =   0
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
                  Caption         =   "frmToolbox.frx":3782
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   -1  'True
                  Tip             =   "frmToolbox.frx":37A8
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmToolbox.frx":37C8
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
         End
         Begin HexUniControls.ctlUniRichTextBoxXP txtPreview 
            Height          =   915
            Index           =   4
            Left            =   120
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   3900
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   1614
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmToolbox.frx":37E4
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
            Tip             =   "frmToolbox.frx":3804
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":3824
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
         Begin VSFlex7LCtl.VSFlexGrid fgRules 
            Height          =   2295
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   6465
            _cx             =   11404
            _cy             =   4048
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
         Begin HexUniControls.ctlUniLabelXP Label5 
            Height          =   255
            Left            =   120
            Top             =   60
            Width           =   8535
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
            Caption         =   "frmToolbox.frx":3840
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmToolbox.frx":3930
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":3950
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraFunctions 
         Height          =   4935
         Left            =   -10590
         TabIndex        =   14
         Top             =   375
         Width           =   9045
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
         Caption         =   "frmToolbox.frx":396C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmToolbox.frx":39A4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":39C4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRichTextBoxXP txtPreview 
            Height          =   915
            Index           =   3
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   3840
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   1614
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmToolbox.frx":39E0
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
            Tip             =   "frmToolbox.frx":3A00
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":3A20
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
         Begin VSFlex7LCtl.VSFlexGrid fgFunctions 
            Height          =   3075
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   6435
            _cx             =   11351
            _cy             =   5424
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
         Begin HexUniControls.ctlUniFrameWL fraFunctionFilter 
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   300
            Width           =   2595
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
            Caption         =   "frmToolbox.frx":3A3C
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmToolbox.frx":3A68
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":3A88
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkTreeFunctions 
               Height          =   220
               Left            =   0
               TabIndex        =   49
               Top             =   0
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
               Caption         =   "frmToolbox.frx":3AA4
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmToolbox.frx":3ACC
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmToolbox.frx":3AEC
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkFuncFav 
               Height          =   220
               Left            =   780
               TabIndex        =   50
               Top             =   0
               Width           =   1755
               _ExtentX        =   3096
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
               Caption         =   "frmToolbox.frx":3B08
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmToolbox.frx":3B50
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmToolbox.frx":3B70
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   255
            Left            =   120
            Top             =   60
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
            Caption         =   "frmToolbox.frx":3B8C
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmToolbox.frx":3C70
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":3C90
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraFilters 
         Height          =   4935
         Left            =   -10890
         TabIndex        =   13
         Top             =   375
         Width           =   9045
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
         Caption         =   "frmToolbox.frx":3CAC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmToolbox.frx":3CD8
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":3CF8
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraFilterMsg 
            Height          =   2955
            Left            =   600
            TabIndex        =   51
            Top             =   960
            Width           =   5055
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
            Caption         =   "frmToolbox.frx":3D14
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmToolbox.frx":3D8A
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":3DAA
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniLabelXP Label2 
               Height          =   975
               Left            =   480
               Top             =   1680
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
               Caption         =   "frmToolbox.frx":3DC6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmToolbox.frx":3F3C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmToolbox.frx":3F5C
               RightToLeft     =   0   'False
               WordWrap        =   -1  'True
            End
            Begin HexUniControls.ctlUniLabelXP Label1 
               Height          =   555
               Left            =   360
               Top             =   960
               Width           =   4335
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
               Caption         =   "frmToolbox.frx":3F78
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmToolbox.frx":4050
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmToolbox.frx":4070
               RightToLeft     =   0   'False
               WordWrap        =   -1  'True
            End
            Begin HexUniControls.ctlUniLabelXP lblFilters 
               Height          =   615
               Left            =   360
               Top             =   480
               Width           =   4335
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
               Caption         =   "frmToolbox.frx":408C
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmToolbox.frx":416A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmToolbox.frx":418A
               RightToLeft     =   0   'False
               WordWrap        =   -1  'True
            End
         End
         Begin HexUniControls.ctlUniCheckXP chkFilters 
            Height          =   255
            Left            =   180
            TabIndex        =   52
            Top             =   60
            Width           =   7935
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
            Caption         =   "frmToolbox.frx":41A6
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            Pressed         =   0   'False
            Tip             =   "frmToolbox.frx":4284
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":42A4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgFilters 
            Height          =   2115
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   6435
            _cx             =   11351
            _cy             =   3731
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
         Begin HexUniControls.ctlUniRichTextBoxXP txtPreview 
            Height          =   915
            Index           =   2
            Left            =   120
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   3900
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   1614
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmToolbox.frx":42C0
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
            Tip             =   "frmToolbox.frx":42E0
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":4300
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
      End
      Begin HexUniControls.ctlUniFrameWL fraCriteria 
         Height          =   4935
         Left            =   -11190
         TabIndex        =   12
         Top             =   375
         Width           =   9045
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
         Caption         =   "frmToolbox.frx":431C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmToolbox.frx":4348
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":4368
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRichTextBoxXP txtPreview 
            Height          =   915
            Index           =   1
            Left            =   120
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   3840
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   1614
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmToolbox.frx":4384
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
            Tip             =   "frmToolbox.frx":43A4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":43C4
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
         Begin VSFlex7LCtl.VSFlexGrid fgCriteria 
            Height          =   3315
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   6435
            _cx             =   11351
            _cy             =   5847
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
         Begin HexUniControls.ctlUniLabelXP lblCriteria 
            Height          =   255
            Left            =   120
            Top             =   60
            Width           =   9300
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
            Caption         =   "frmToolbox.frx":43E0
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmToolbox.frx":44E0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":4500
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSymbolGroups 
         Height          =   4935
         Left            =   -11490
         TabIndex        =   11
         Top             =   375
         Width           =   9045
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
         Caption         =   "frmToolbox.frx":451C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmToolbox.frx":4548
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmToolbox.frx":4568
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgGroups 
            Height          =   3315
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   6435
            _cx             =   11351
            _cy             =   5847
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
         Begin HexUniControls.ctlUniRichTextBoxXP txtPreview 
            Height          =   915
            Index           =   0
            Left            =   120
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   3900
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   1614
            BackColor       =   -2147483648
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmToolbox.frx":4584
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
            Tip             =   "frmToolbox.frx":45A4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":45C4
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
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   255
            Left            =   120
            Top             =   60
            Width           =   6855
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
            Caption         =   "frmToolbox.frx":45E0
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmToolbox.frx":466A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbox.frx":468A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   8160
      Top             =   4560
      Width           =   675
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuAddCopy 
         Caption         =   "Add Copy"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuNewDLL 
         Caption         =   "New DLL"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Info"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuDependencies 
         Caption         =   "Dependencies"
      End
      Begin VB.Menu mnuExportAll 
         Caption         =   "Export All"
      End
      Begin VB.Menu mnuRenameFile 
         Caption         =   "Rename File"
      End
      Begin VB.Menu mnuExportList 
         Caption         =   "Export List"
      End
      Begin VB.Menu mnuCreateAutoTrade 
         Caption         =   "Create Auto Trade"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmToolbox.frm
'' Description: Allows the user to select one or more Strategies, Rules, or
''              Functions
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/13/2009   DAJ         Made sure not to load a symbol group with a blank ID
''                          and check for valid group when changing rows in grid
'' 01/28/2009   DAJ         Put in additional debugging for Criteria/Filter issue
'' 02/02/2009   DAJ         Fix for Type Mismatch in ValidateEdit of Filters and
''                          Criteria grids when EditText is blank
'' 04/22/2009   DAJ         Make sure to call FixPyramidInfo after a Library Import
'' 09/20/2011   DAJ         Added strategy list export for in house people
'' 11/15/2011   DAJ         Renamed the Strategy Basket stuff
'' 04/03/2013   DAJ         Move Strategy Baskets into the database
'' 05/01/2013   DAJ         Shadow Trading
'' 07/23/2013   DAJ         Show baskets if has module OR is IDE
'' 08/12/2013   DAJ         Show baskets if has module OR is IDE OR is owner of guru basket
'' 08/18/2014   DAJ         Added Last Modified column to Strategy Basket grid
'' 08/18/2014   DAJ         Don't show seconds on Last Modified time in Strategy Basket grid
'' 05/05/2015   DAJ         Don't let user delete basket under certain conditions; confirm in others
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Mode to call the form in (Select or Add)
Public Enum eAddFormMode
    eAddMode_Select = 0
    eAddMode_Add = 1
    eAddMode_List = 2
End Enum

' Columns in the Symbol Group Grid
Private Enum eSymbolGroupColumns
    eGGC_Name = 0
    eGGC_Description = 1
    eGGC_GroupID = 2
End Enum
Private Const kGroupsGridCols = 3

' Columns in the Criteria Grid
Private Enum eCriteriaGridColumns
    eCGC_Active = 0
    eCGC_Name = 1
    eCGC_Description = 2
    eCGC_NumDays = 3
    eCGC_CodedText = 4
    eCGC_CriteriaID = 5
End Enum
Private Const kCriteriaGridCols = 6

' Columns in the Filters Grid
Private Enum eFiltersGridColumns
    eFIC_Active = 0
    eFIC_Name = 1
    eFIC_Description = 2
    eFIC_FilterID = 3
End Enum
Private Const kFiltersGridCols = 4

' Columns in the Systems Grid
Private Enum eSystemGridColumns
    eSGC_SystemDesc = 0
    eSGC_LibraryName = 1
    eSGC_Developer = 2
    eSGC_LastModified = 3
    eSGC_UnVerified = 4
    eSGC_Preview = 5
    eSGC_SystemNumber = 6
    eSGC_TradesPath = 7
    eSGC_SecurityLevel = 8
    eSGC_Password = 9
    eSGC_CannotDelete = 10
End Enum
Private Const kSystemGridCols = 11

' Columns in the Rules Grid
Private Enum eRulesGridColumns
    eRGC_RuleName = 0
    eRGC_RuleType
    eRGC_SystemName
    eRGC_LibraryName
    eRGC_CategoryID
    eRGC_LastModified
    eRGC_Preview
    eRGC_RuleID
    eRGC_SecurityLevel
    eRGC_Password
    eRGC_CannotDelete
    eRGC_SystemNumber
    eRGC_Reverify
    eRGC_TreeSortKey
    eRGC_TreeLevel
    eRGC_NumCols
End Enum

' Columns in the Functions Grid
Private Enum eFunctionGridColumns
    eFGC_FunctionName = 0
    eFGC_LibraryName
    eFGC_Category
    eFGC_LastModified
    eFGC_ImplType
    eFGC_Usage
    eFGC_Preview
    eFGC_FunctionID
    eFGC_SecurityLevel
    eFGC_Password
    eFGC_CannotDelete
    eFGC_Reverify
    eFGC_Favorites
    eFGC_TreeSortKey
    eFGC_TreeLevel
    eFGC_CodedName
    eFGC_NumCols
End Enum
Private Const kFunctionGridCols = 12

' Columns in the Libraries Grid
Private Enum eLibraryGridColumns
    eLGC_LibraryName = 0
    eLGC_Author = 1
    eLGC_LastModified = 2
    eLGC_Preview = 3
    eLGC_LibraryID = 4
    eLGC_SecurityLevel = 5
    eLGC_Password = 6
    eLGC_CannotDelete = 7
End Enum
Private Const kLibraryGridCols = 8

Private Enum eStrategyBasketColumns
    eSBC_Name = 0
    eSBC_LastModified = 1
    eSBC_Description = 2
End Enum
Private Const kStrategyBasketGridCols = 3

' Tabs on the form
Public Enum eAddFormTabs
    eTab_SymbolGroups = 0
    eTab_Criteria = 1
    eTab_Filters = 2
    eTab_Functions = 3
    eTab_Rules = 4
    eTab_Systems = 5
    eTab_StrategyBaskets = 6 '7
    eTab_Libraries = 7 '6
    eTab_NumTabs
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Was the form Cancelled or OK'd?
    Mode As eAddFormMode                ' What mode is the form called in?
    lSystemNumber As Long               ' System Number called from
    lLibraryID As Long                  ' Library ID of the System
    alReturnIds As cGdArray             ' Array of IDs to return
    SystemRules As cRules               ' Rules in the System passed in
    RulesToAdd As cRules                 ' Rules to add to the system
    strInitialSelect As String          ' Row to initially select (optional)
    strPrevSort As String
    lUsedInStrategiesRow As Long
    
    alLongest As cGdArray
    abAutoSize As cGdArray
    lPrevColWidth As Long       ' used when resizing columns
End Type
Private m As mPrivate

Private Function Tabs(ByVal peTab As eAddFormTabs) As Long
    Tabs = peTab
End Function
Private Function GGCol(ByVal lColumn As eSymbolGroupColumns) As Long
    GGCol = lColumn
End Function
Private Function CGCol(ByVal lColumn As eCriteriaGridColumns) As Long
    CGCol = lColumn
End Function
Private Function FiCol(ByVal lColumn As eFiltersGridColumns) As Long
    FiCol = lColumn
End Function
Private Function RGCol(ByVal eCol As eRulesGridColumns) As Long
    RGCol = eCol
End Function
Private Function SGCol(ByVal eCol As eSystemGridColumns) As Long
    SGCol = eCol
End Function
Private Function FGCol(ByVal eCol As eFunctionGridColumns) As Long
    FGCol = eCol
End Function
Private Function LGCol(ByVal eCol As eLibraryGridColumns) As Long
    LGCol = eCol
End Function
Private Function SBCol(ByVal eCol As eStrategyBasketColumns) As Long
    SBCol = eCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboLibrary_Click
'' Description: If the user changes the Library filter, filter the rules grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboLibrary_Click()
On Error GoTo ErrSection:

    FilterRulesGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cboLibrary.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkFilters_Click()

    If Me.Visible Then
        ScansEnabled = (chkFilters.Value <> 0)
        FixFilterDisplay
        If chkFilters.Value <> 0 Then
            MoveFocus fgFilters
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkFuncFav_Click
'' Description: Filter the functions on the favorites
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkFuncFav_Click()
On Error GoTo ErrSection:

    FilterFunctionsGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.chkFuncFav.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkLibrary_Click
'' Description: If the user chooses to Filter on the Library, Filter the Grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkLibrary_Click()
On Error GoTo ErrSection:

    FilterRulesGrid
    cboLibrary.Enabled = (chkLibrary = vbChecked)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.chkLibrary.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkFavorites_Click
'' Description: Filter the grid as appropriate on the Show Local check box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkFavorites_Click()
On Error GoTo ErrSection:

    fgRules.Redraw = flexRDNone
    FilterRulesGrid
    fgRules.Redraw = flexRDBuffered

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.chkFavorites.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkTreeFunctions_Click
'' Description: Allow the user to choose between Tree and Grid for Functions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkTreeFunctions_Click()
On Error GoTo ErrSection:

    If Me.Visible Then ChangeFunctionsView

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.chkTreeFunctions.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkTreeRules_Click
'' Description: Allow the user to choose between Tree and Grid views for Rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkTreeRules_Click()
On Error GoTo ErrSection:
    
    If Me.Visible Then ChangeRulesView

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.chkTreeRules.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: Return the IDs of the items selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim lRow As Long                    ' Current row working on from grid
    Dim Rule As New cRule               ' Working version of a Rule

    Set m.alReturnIds = New cGdArray
    m.alReturnIds.Create eGDARRAY_Longs

    With fgRules
        For lIndex = 0 To .SelectedRows - 1
            lRow = .SelectedRow(lIndex)
            If Not .RowHidden(lRow) Then
                If .TextMatrix(lRow, RGCol(eRGC_SystemNumber)) = "0" Then
                    m.alReturnIds.Add CLng(.TextMatrix(lRow, RGCol(eRGC_RuleID)))
                Else
                    If InfBox("This will make a local copy of|" & .TextMatrix(lRow, RGCol(eRGC_RuleName)) & _
                                ".||Do you want to do this?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                        Set Rule = New cRule
                        With Rule
                            .RuleID = CLng(fgRules.TextMatrix(lRow, RGCol(eRGC_RuleID)))
                            .Load
                            
                            If g.Security.CanEdit(.SecurityLevel, .Password, fgRules.TextMatrix(lRow, RGCol(eRGC_RuleName))) Then
                                .RuleID = 0
                                .SystemNumber = -2 'm.lSystemNumber
                                .LibraryID = m.lLibraryID
                                .Save
                                m.alReturnIds.Add .RuleID
                            End If
                        End With
                    End If
                End If
            End If
        Next lIndex
    End With
    
    m.bOK = True

ErrExit:
    Set Rule = Nothing
    Me.Hide
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdAdd.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddCopy_Click
'' Description: Return the IDs of the local copies added
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddCopy_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim lRow As Long                    ' Current row working on from grid
    Dim Rule As New cRule               ' Working version of a Rule
    Dim lSysNbr As Long                 ' Current System Number of the Rule

    'Set m.alReturnIds = New cGdArray
    'm.alReturnIds.Create eGDARRAY_Longs
    Set m.RulesToAdd = New cRules

    With fgRules
        For lIndex = 0 To .SelectedRows - 1
            lRow = .SelectedRow(lIndex)
            If Not .RowHidden(lRow) Then
                Set Rule = New cRule
                With Rule
                    .RuleID = CLng(fgRules.TextMatrix(lRow, RGCol(eRGC_RuleID)))
                    .LoadWithSystemInfo .RuleID
                    
                    If g.Security.CanEdit(.SecurityLevel, .Password, fgRules.TextMatrix(lRow, RGCol(eRGC_RuleName))) Then
                        m.RulesToAdd.Add .RuleID, Rule
                    End If
                End With
            End If
        Next lIndex
    End With
    
    m.bOK = True
    
ErrExit:
    Set Rule = Nothing
    Me.Hide
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdAddCopy.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on the Cancel button, unload the form
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
    RaiseError "frmToolbox.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow the user to delete the currently selected item(s)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Dim lID As Long                     ' ID of item to delete
    Dim System As New cSystem           ' System to delete
    Dim Rule As New cRule               ' Rule to delete
    Dim Func As New cFunction           ' Function to delete
    Dim Library As cLibrary             ' Library to delete
    Dim Basket As cStrategyBasket       ' Strategy basket object
    Dim LMB As cLibManagerBridge        ' Library Manager Bridge
    Dim lRow As Long                    ' Row selected in the grid
    Dim strID As String                 ' ID of the item from the pool
    Dim SymbolGroup As cSymbolGroup     ' Temporary Symbol Group object
    Dim strMsg As String                ' Message to display to the user
    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return value from the infbox to user
    Dim iPos As Long
    Dim bContinue As Boolean            ' Continue with deleting?
    Dim bOnlyOne As Boolean             ' Only one selected?

    Select Case vsTypeTabs.CurrTab
        Case Tabs(eTab_SymbolGroups)
            With fgGroups
                If .SelectedRows = 1 Then
                    strMsg = "Are you sure you want to delete:|" & .TextMatrix(.SelectedRow(0), GGCol(eGGC_Name))
                    strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Symbol Group")
                ElseIf .SelectedRows > 0 Then
                    strMsg = "Are you sure you want to delete the selected Symbol Groups?"
                    strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Symbol Groups")
                End If
            
                If strReturn = "D" Then
                    lRow = .SelectedRow(0)
                    For lIndex = .SelectedRows - 1 To 0 Step -1
                        If .RowHidden(.SelectedRow(lIndex)) = False Then
                            strID = .TextMatrix(.SelectedRow(lIndex), GGCol(eGGC_GroupID))
                            KillFile AddSlash(App.Path) & "Custom\" & strID
                            With g.SymbolPool
                                If .SymbolGroups.Exists(strID) Then
                                    Set SymbolGroup = .SymbolGroups(strID)
                                    If SymbolGroup.IsIndex Then
                                        SU_DeleteComposite SymbolGroup.SymbolID, UCase("#" & SymbolGroup.Name)
                                        .RemoveCustomIndex SymbolGroup.SymbolID
                                        frmSymbolGrid.RefreshGrid
                                    End If
                                    .SymbolGroups.Remove strID
                                    .RemoveOrphanedArraysFromTable
                                End If
                            End With
                            
                            .RemoveItem .SelectedRow(lIndex)
                        End If
                    Next lIndex
                    
                    If lRow >= .FixedRows And lRow < .Rows Then
                        .Row = lRow
                        .RowSel = lRow
                    ElseIf .Rows > .FixedRows Then
                        .Row = .FixedRows
                        .RowSel = .FixedRows
                    End If
                    
                    frmSymbolGrid.LoadCombo
                End If
            End With
        
        Case Tabs(eTab_Criteria)
            With fgCriteria
                If .SelectedRows = 1 Then
                    strMsg = "Are you sure you want to delete:|" & .TextMatrix(.SelectedRow(0), CGCol(eCGC_Name))
                    strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Criteria")
                ElseIf .SelectedRows > 0 Then
                    strMsg = "Are you sure you want to delete the selected Criteria?"
                    strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Criteria")
                End If
                
                If strReturn = "D" Then
                    lRow = .SelectedRow(0)
                    For lIndex = .SelectedRows - 1 To 0 Step -1
                        If .RowHidden(.SelectedRow(lIndex)) = False Then
                            strID = .TextMatrix(.SelectedRow(lIndex), CGCol(eCGC_CriteriaID))
                            KillFile AddSlash(App.Path) & "Custom\" & strID
                            With g.SymbolPool
                                If .Criterias.Exists(strID) Then
                                    .Criterias.Remove strID
                                    .RemoveOrphanedArraysFromTable
                                End If
                            End With
                            
                            .RemoveItem .SelectedRow(lIndex)
                        End If
                    Next lIndex
                    
                    If lRow >= .FixedRows And lRow < .Rows Then
                        .Row = lRow
                        .RowSel = lRow
                    ElseIf .Rows > .FixedRows Then
                        .Row = .FixedRows
                        .RowSel = .FixedRows
                    End If
                    
                    frmSymbolGrid.LoadCombo
                End If
            End With
        
        Case Tabs(eTab_Filters)
            With fgFilters
                If .SelectedRows = 1 Then
                    strMsg = "Are you sure you want to delete:|" & .TextMatrix(.SelectedRow(0), FiCol(eFIC_Name))
                    strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Filter")
                ElseIf .SelectedRows > 0 Then
                    strMsg = "Are you sure you want to delete the selected Filters?"
                    strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Filters")
                End If
                
                If strReturn = "D" Then
                    lRow = .SelectedRow(0)
                    For lIndex = .SelectedRows - 1 To 0 Step -1
                        If .RowHidden(.SelectedRow(lIndex)) = False Then
                            strID = .TextMatrix(.SelectedRow(lIndex), FiCol(eFIC_FilterID))
                            KillFile AddSlash(App.Path) & "Custom\" & strID
                            With g.SymbolPool
                                If .Filters.Exists(strID) Then
                                    .Filters.Remove strID
                                    .RemoveOrphanedArraysFromTable
                                End If
                            End With
                            .RemoveItem .SelectedRow(lIndex)
                        End If
                    Next lIndex
            
                    If lRow >= .FixedRows And lRow < .Rows Then
                        .Row = lRow
                        .RowSel = lRow
                    ElseIf .Rows > .FixedRows Then
                        .Row = .FixedRows
                        .RowSel = .FixedRows
                    End If
                    
                    frmSymbolGrid.LoadCombo
                End If
            End With
        
        Case Tabs(eTab_Systems)
            If HasPlatinum(True) Then
                With fgSystems
                    If .SelectedRows = 1 Then
                        strMsg = "Are you sure you want to delete:|" & .TextMatrix(.SelectedRow(0), SGCol(eSGC_SystemDesc))
                        strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Strategy")
                    ElseIf .SelectedRows > 0 Then
                        strMsg = "Are you sure you want to delete the selected Strategies?"
                        strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Strategies")
                    End If
                    
                    If strReturn = "D" Then
                        lRow = .SelectedRow(0)
                        For lIndex = .SelectedRows - 1 To 0 Step -1
                            If .RowHidden(.SelectedRow(lIndex)) = False Then
                                lID = Val(.TextMatrix(.SelectedRow(lIndex), SGCol(eSGC_SystemNumber)))
                                
                                If CheckedCell(fgSystems, .SelectedRow(lIndex), SGCol(eSGC_CannotDelete)) = True Then
                                    InfBox .TextMatrix(.SelectedRow(lIndex), SGCol(eSGC_SystemDesc)) & " has been marked as protected by the developer and cannot be deleted", "!", , "Delete Error"
                                ElseIf g.Security.CanRemove("Strategy", _
                                        .Cell(flexcpValue, .SelectedRow(lIndex), SGCol(eSGC_SecurityLevel)), _
                                        .TextMatrix(.SelectedRow(lIndex), SGCol(eSGC_Password)), _
                                        .TextMatrix(.SelectedRow(lIndex), SGCol(eSGC_CannotDelete)), _
                                        .TextMatrix(.SelectedRow(lIndex), SGCol(eSGC_SystemDesc))) Then
                        
                                    Set System = New cSystem
                                    With System
                                        .SystemNumber = lID
                                        .Delete
                                    End With
                                    Set System = Nothing
                                    
                                    .RemoveItem .SelectedRow(lIndex)
                                End If
                            End If
                        Next lIndex
                            
                        ' Make sure that the next time the user goes to the
                        ' rules tab that it gets reloaded (DAJ: 03/24/2003)...
                        fgRules.Rows = fgRules.FixedRows
                            
                        If lRow >= .FixedRows And lRow < .Rows Then
                            .Row = lRow
                            .RowSel = lRow
                        ElseIf .Rows > .FixedRows Then
                            .Row = .FixedRows
                            .RowSel = .FixedRows
                        End If
                    End If
                End With
            End If
        
        Case Tabs(eTab_Rules)
            If 0 And IsIDE Then
                ' for finding certain kinds of rules ...
                With fgRules
                    For lRow = .FixedRows To .Rows - 1
                        .RowHidden(lRow) = True
                        strMsg = UCase(.TextMatrix(lRow, RGCol(eRGC_Preview)))
                        'iPos = InStr(strMsg, "~24005THEN")
                        iPos = InStr(strMsg, "~01013NEXT BAR OPEN ~21002OF")
                        If iPos > 0 Then
                            strMsg = Trim(Mid(strMsg, iPos + 10))
                            'If InStr(strMsg, "*") > 0 And InStr(strMsg, "TICK MOVE") = 0 Then
                            If InStr(strMsg, "AND") > 0 Or InStr(strMsg, "OR") > 0 Then
                                .RowHidden(lRow) = False
                            End If
                        End If
                    Next
                End With
            ElseIf HasPlatinum(True) Then
                With fgRules
                    If .SelectedRows = 1 Then
                        strMsg = "Are you sure you want to delete:|" & .TextMatrix(.SelectedRow(0), RGCol(eRGC_RuleName))
                        strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Rule")
                    ElseIf .SelectedRows > 0 Then
                        strMsg = "Are you sure you want to delete the selected Rules?"
                        strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Rules")
                    End If
                    
                    If strReturn = "D" Then
                        lRow = .SelectedRow(0)
                        For lIndex = .SelectedRows - 1 To 0 Step -1
                            If .RowHidden(.SelectedRow(lIndex)) = False Then
                                lID = Val(.TextMatrix(.SelectedRow(lIndex), RGCol(eRGC_RuleID)))
                                If lID > 0 Then
                                    If CheckedCell(fgRules, .SelectedRow(lIndex), RGCol(eRGC_CannotDelete)) = True Then
                                        InfBox .TextMatrix(.SelectedRow(lIndex), RGCol(eRGC_RuleName)) & " has been marked as protected by the developer and cannot be deleted", "!", , "Delete Error"
                                    ElseIf g.Security.CanRemove("Rule", _
                                            .Cell(flexcpValue, .SelectedRow(lIndex), RGCol(eRGC_SecurityLevel)), _
                                            .TextMatrix(.SelectedRow(lIndex), RGCol(eRGC_Password)), _
                                            .TextMatrix(.SelectedRow(lIndex), RGCol(eRGC_CannotDelete)), _
                                            .TextMatrix(.SelectedRow(lIndex), RGCol(eRGC_RuleName))) Then
                            
                                        Set Rule = New cRule
                                        With Rule
                                            .RuleID = lID
                                            .Delete
                                        End With
                                        Set Rule = Nothing
                                    
                                        ' Delete the rule from the global table...
                                        DeleteRule lID
                                        
                                        .RemoveItem .SelectedRow(lIndex)
                                    End If
                                End If
                            End If
                        Next lIndex
                        
                        If lRow >= .FixedRows And lRow < .Rows Then
                            .Row = lRow
                            .RowSel = lRow
                        ElseIf .Rows > .FixedRows Then
                            .Row = .FixedRows
                            .RowSel = .FixedRows
                        End If
                    End If
                End With
            End If
        
        Case Tabs(eTab_Functions)
            If HasGold(True) Then
                With fgFunctions
                    If .SelectedRows = 1 Then
                        strMsg = "Are you sure you want to delete:|" & .TextMatrix(.SelectedRow(0), FGCol(eFGC_FunctionName))
                        strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Function")
                    ElseIf .SelectedRows > 0 Then
                        strMsg = "Are you sure you want to delete the selected Functions?"
                        strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Functions")
                    End If
                    
                    If strReturn = "D" Then
                        lRow = .SelectedRow(0)
                        For lIndex = .SelectedRows - 1 To 0 Step -1
                            If .RowHidden(.SelectedRow(lIndex)) = False Then
                                lID = Val(.TextMatrix(.SelectedRow(lIndex), FGCol(eFGC_FunctionID)))
                                If lID > 0 Then
                                    If CheckedCell(fgFunctions, .SelectedRow(lIndex), FGCol(eFGC_CannotDelete)) = True Then
                                        InfBox .TextMatrix(.SelectedRow(lIndex), FGCol(eFGC_FunctionName)) & " has been marked as protected by the developer and cannot be deleted", "!", , "Delete Error"
                                    ElseIf g.Security.CanRemove("Function", _
                                            .Cell(flexcpValue, .SelectedRow(lIndex), FGCol(eFGC_SecurityLevel)), _
                                            .TextMatrix(.SelectedRow(lIndex), FGCol(eFGC_Password)), _
                                            .TextMatrix(.SelectedRow(lIndex), FGCol(eFGC_CannotDelete)), _
                                            .TextMatrix(.SelectedRow(lIndex), FGCol(eFGC_FunctionName))) Then
                            
                                        Set Func = New cFunction
                                        With Func
                                            .FunctionID = lID
                                            .Delete
                                        End With
                                        Set Func = Nothing
                                        
                                        ' Delete the function from the engine tables and the
                                        ' global function collection
                                        DeleteFunction lID
                                        
                                        .RemoveItem .SelectedRow(lIndex)
                                    End If
                                End If
                            End If
                        Next lIndex
                            
                        If lRow >= .FixedRows And lRow < .Rows Then
                            .Row = lRow
                            .RowSel = lRow
                        ElseIf .Rows > .FixedRows Then
                            .Row = .FixedRows
                            .RowSel = .FixedRows
                        End If
                    End If
                End With
            End If
        
        Case Tabs(eTab_Libraries)
            With fgLibraries
                If .SelectedRows = 1 Then
                    strMsg = "Are you sure you want to delete:|" & .TextMatrix(.SelectedRow(0), LGCol(eLGC_LibraryName))
                    strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Library")
                ElseIf .SelectedRows > 0 Then
                    strMsg = "Are you sure you want to delete the selected Libraries?"
                    strReturn = InfBox(strMsg, "?", "+Delete|-Cancel", "Delete Libraries")
                End If
            
                If strReturn = "D" Then
                    lRow = .SelectedRow(0)
                    Set LMB = GetLibMgrBridge
                    
                    For lIndex = .SelectedRows - 1 To 0 Step -1
                        If .RowHidden(.SelectedRow(lIndex)) = False Then
                            lID = Val(.TextMatrix(.SelectedRow(lIndex), LGCol(eLGC_LibraryID)))
                        
                            'Make sure user is authorized
                            If CheckedCell(fgLibraries, .SelectedRow(lIndex), LGCol(eLGC_CannotDelete)) Then
                                InfBox .TextMatrix(.SelectedRow(lIndex), LGCol(eLGC_LibraryName)) & " is a required library and cannot be deleted.", "!", , "Delete Error"
                            Else
                                Screen.MousePointer = vbHourglass
                                
                                Set Library = New cLibrary
                                With Library
                                    .LibraryID = lID
                                    .Delete
                                End With
                                
                                ' Delete the library from the global engine tables
                                DeleteLibrary lID
                                
                                ' Reload the Function and Rule tables
                                'LoadEngineFunctions
                    
                                ' Trigger a reload the next time they switch to these tabs...
                                fgFunctions.Rows = fgFunctions.FixedRows
                                fgRules.Rows = fgRules.FixedRows
                                fgSystems.Rows = fgSystems.FixedRows
                                
                                .RemoveItem .SelectedRow(lIndex)
                            End If
                        End If
                    Next lIndex
                        
                    If lRow >= .FixedRows And lRow < .Rows Then
                        .Row = lRow
                        .RowSel = lRow
                    ElseIf .Rows > .FixedRows Then
                        .Row = .FixedRows
                        .RowSel = .FixedRows
                    End If
                    Screen.MousePointer = vbDefault
                End If
            End With
            
        Case Tabs(eTab_StrategyBaskets)
            If HasPlatinum(True) Then
                With fgStrategyBaskets
                    If .SelectedRows = 1 Then
                        bOnlyOne = True
                        bContinue = CanDeleteStrategyBasket(.RowData(.SelectedRow(0)), True)
                    ElseIf .SelectedRows > 0 Then
                        bOnlyOne = False
                        bContinue = (InfBox("Are you sure you want to delete the selected Strategy Baskets?", "?", "Delete|+-Cancel", "Delete Strategy Baskets") = "D")
                    End If
                    
                    If bContinue Then
                        lRow = .SelectedRow(0)
                        For lIndex = .SelectedRows - 1 To 0 Step -1
                            If .RowHidden(.SelectedRow(lIndex)) = False Then
                                Set Basket = .RowData(.SelectedRow(lIndex))
                                
                                If bOnlyOne = True Then
                                    bContinue = True
                                Else
                                    bContinue = CanDeleteStrategyBasket(Basket, False)
                                End If
                                
                                If bContinue = True Then
                                    Basket.DeleteDb
                                    .RemoveItem .SelectedRow(lIndex)
                                    
                                    g.TradingItems.DeleteForBasket Basket.ID
                                End If
                            End If
                        Next lIndex
                        
                        If lRow >= .FixedRows And lRow < .Rows Then
                            .Row = lRow
                            .RowSel = lRow
                        ElseIf .Rows > .FixedRows Then
                            .Row = .FixedRows
                            .RowSel = .FixedRows
                        End If
                    End If
                End With
            End If
            
    End Select

ErrExit:
    Set System = Nothing
    Set Rule = Nothing
    Set Func = Nothing
    Set Library = Nothing
    Set LMB = Nothing
    Set SymbolGroup = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdDelete.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to edit the item they have selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    Dim lID As Long                     ' ID of the item to edit
    Dim LibMgrBridge As cLibManagerBridge
    Dim frm As Form
    Dim obj As Object
    Dim strID As String
    Dim lSystemID As Long
    Dim frm2 As frmSystemManager
    Dim bReload As Boolean
    Dim bEdit As Boolean                ' Does the user want to edit the function?
    Dim Basket As cStrategyBasket       ' Selected strategy basket

    MoveFocus cmdEdit

    Select Case vsTypeTabs.CurrTab
        Case Tabs(eTab_SymbolGroups)
            With fgGroups
                strID = Trim(.TextMatrix(.RowSel, GGCol(eGGC_GroupID)))
            End With
            If strID <> "" Then
                Me.Hide
                If Not ActivateEditor("frmSymbolGroup", strID) Then
                    Set frm = New frmSymbolGroup
                    frm.ShowMe AddSlash(App.Path) & "Custom\", strID
                End If
            End If
            
        Case Tabs(eTab_Criteria)
            With fgCriteria
                strID = Trim(.TextMatrix(.RowSel, CGCol(eCGC_CriteriaID)))
            End With
            If strID <> "" Then
                Me.Hide
                If Not ActivateEditor("frmCriteria", strID) Then
                    Set frm = New frmCriteria
                    frm.ShowMe AddSlash(App.Path) & "Custom\", strID
                End If
            End If
        
        Case Tabs(eTab_Filters)
            With fgFilters
                strID = Trim(.TextMatrix(.RowSel, FiCol(eFIC_FilterID)))
            End With
            If strID <> "" Then
                Me.Hide
                If Not ActivateEditor("frmFilter", strID) Then
                    Set frm = New frmFilter
                    frm.ShowMe AddSlash(App.Path) & "Custom\", strID
                End If
            End If
        
        Case Tabs(eTab_Systems)
            bEdit = False
            If HasModule("JDMP") Then ' first check for non-gold allowances (to skip upgrade message)
                bEdit = True
            ElseIf HasGold(True) Then ' then give upgrade message if still not allowed
                bEdit = True
            End If
            If bEdit Then
                With fgSystems
                    If .RowSel >= .FixedRows And .RowSel < .Rows Then
                        lID = CLng(.TextMatrix(.RowSel, SGCol(eSGC_SystemNumber)))
                        Me.Hide
                        If Not ActivateEditor("frmSystemManager", lID) Then
                            Set frm = New frmSystemManager
                            frm.ShowMe lID, , False
                        End If
                    End If
                End With
            End If

        Case Tabs(eTab_Rules)
            If HasPlatinum(True) Then
                With fgRules
                    If .RowSel >= .FixedRows And .RowSel < .Rows Then
                        lID = CLng(.TextMatrix(.RowSel, RGCol(eRGC_RuleID)))
                        lSystemID = CLng(.TextMatrix(.RowSel, RGCol(eRGC_SystemNumber)))
                        Me.Hide
                        If Not ActivateEditor("frmRule", lID) Then
                            If lSystemID <> 0 Then
                                If Not ActivateEditor("frmSystemManager", lSystemID, frm2) Then
                                    Set frm2 = New frmSystemManager
                                    If frm2.ShowMe(lSystemID, , False) Then
                                        frm2.EditRule lID
                                    End If
                                Else
                                    frm2.EditRule lID
                                End If
                            Else
                                Set frm = New frmRule
                                frm.ShowMe Name, lID
                            End If
                        End If
                    End If
                End With
            End If
        
        Case Tabs(eTab_Functions)
            If HasGold(True) Then
                With fgFunctions
                    bEdit = True
                    If (.SelectedRows > 1) And (IsIDE = True) Then
                        If InfBox("You are about to resave all of the selected functions.|Do you want to continue?|", "?", "+Yes|-No", "Resave Functions") = "Y" Then
                            ResaveFunctions
                            bEdit = False
                        End If
                    End If
                    
                    If bEdit = True Then
                        If .RowSel >= .FixedRows And .RowSel < .Rows Then
                            lID = CLng(.TextMatrix(.RowSel, FGCol(eFGC_FunctionID)))
                            If .TextMatrix(.RowSel, FGCol(eFGC_ImplType)) = "2" Then
                                Me.Hide
                                If Not ActivateEditor("frmFunctionMgrCT", lID) Then
                                    Set frm = New frmFunctionMgrCT
                                    frm.ShowMe lID
                                End If
                            ElseIf cmdNewDLL.Visible Then
                                Me.Hide
                                If Not ActivateEditor("frmFunctionMgr", lID) Then
                                    Set frm = New frmFunctionMgr
                                    frm.ShowMe lID
                                End If
                            Else
                                InfBox "This is a compiled DLL function|(there is no TradeSense code to view/edit).", "i", , "Edit Function"
                            End If
                        End If
                    End If
                End With
            End If
        
        Case Tabs(eTab_Libraries)
            With fgLibraries
                If .RowSel >= .FixedRows And .RowSel < .Rows Then
                    lID = CLng(.TextMatrix(.RowSel, LGCol(eLGC_LibraryID)))
                    
                    'TLB: let's keep the toolbox up when editing a library
                    ' (since library manager is modal anyway)
                    'Me.Hide
            
                    Set LibMgrBridge = GetLibMgrBridge
            
                    'Make sure user is authorized
                    'If not, then show Viewer form only
                    If g.Security.CanEdit( _
                            .Cell(flexcpValue, .RowSel, LGCol(eLGC_SecurityLevel)), _
                            .TextMatrix(.RowSel, LGCol(eLGC_Password))) Then
                        LibMgrBridge.EditLibrary lID
                        bReload = LibMgrBridge.Saved
                    Else
                        LibMgrBridge.ViewLibrary lID
                        bReload = False
                    End If
                End If

                If bReload Then
                    g.bDirtyLibrariesMDB = True
                    Screen.MousePointer = vbHourglass
                    StatusMsg "Reloading Libraries ...", vbRed
                    
                    ' Reload the libraries grid...
                    m.strInitialSelect = Str(lID)
                    LoadLibrariesGrid
                    
                    ' Reload the Function and Rule tables in memory...
                    'LoadEngineFunctions
                    RefreshLibrary lID
                    RefreshLibrary kSN_UserLibrary
                    
                    ' Trigger a reload the next time they switch to these tabs...
                    fgFunctions.Rows = fgFunctions.FixedRows
                    fgRules.Rows = fgRules.FixedRows
                    fgSystems.Rows = fgSystems.FixedRows
                    
                    Screen.MousePointer = vbDefault
                    StatusMsg
                End If
            End With
            
        Case Tabs(eTab_StrategyBaskets)
            If HasPlatinum(True) Then
                With fgStrategyBaskets
                    If (.RowSel >= .FixedRows) And (.RowSel < .Rows) Then
                        Set Basket = .RowData(.RowSel)
                        
                        strID = Str(Basket.ID)
                        Me.Hide
                        If Not ActivateEditor("frmStrategyBasket", strID) Then
                            Set frm = New frmStrategyBasket
                            frm.ShowMe Basket.ID
                        End If
                    End If
                End With
            End If
            
    End Select

ErrExit:
    Set obj = Nothing
    Set LibMgrBridge = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdExport_Click
'' Description: Export the library containing the selected item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExport_Click()
On Error GoTo ErrSection:

    Export
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdExport.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdImport_Click
'' Description: Allow the user to import a library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdImport_Click()
On Error GoTo ErrSection:

    Dim LibMgrBridge As cLibManagerBridge
       
    Set LibMgrBridge = GetLibMgrBridge
    With LibMgrBridge
        .ShowImporter
        
        If .ImportOK Then
            InfBox "Please wait while reloading functions and rules...", , , "Reloading", True
            Screen.MousePointer = vbHourglass
                
            ' Clean up any bogus pyramiding information (but don't reload the rules table
            ' in here because it will be done right after this)...
            FixPyramidInfo False
            
            ' Reload the libraries grid...
            LoadLibrariesGrid
            
            ' Reload the Function and Rule tables in memory...
            LoadEngineFunctions
            LoadRulesTable
            Set g.Functions = New cFunctions
            g.Functions.Load
            FilterFunctions
            
            ' Trigger a reload the next time they switch to these tabs...
            fgFunctions.Rows = fgFunctions.FixedRows
            fgRules.Rows = fgRules.FixedRows
            fgSystems.Rows = fgSystems.FixedRows
            
            mSysNav.CreateGuruAutoTradeItems
        End If
    End With
    
ErrExit:
    Screen.MousePointer = vbDefault
    InfBox ""
    Set LibMgrBridge = Nothing
    
    ' If the RestoreMDB.FLG file exists after attempting an import, we need to
    ' shut down the program so that when they start it back up, we can restore
    ' the old database
    If FileExist(AddSlash(App.Path) & "RestoreMDB.FLG") Then
        Me.Hide
    End If
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdImport.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdInfo_Click
'' Description: Allow the user to view information on a library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdInfo_Click()
On Error GoTo ErrSection:

    Dim LibMgrBridge As cLibManagerBridge
    Dim lID As Long
    
    With fgLibraries
        If .RowSel >= .FixedRows And .RowSel < .Rows Then
            lID = CLng(.TextMatrix(.RowSel, LGCol(eLGC_LibraryID)))
    
            Set LibMgrBridge = GetLibMgrBridge
            LibMgrBridge.ViewLibrary lID
        End If
    End With

ErrExit:
    Set LibMgrBridge = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdInfo.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNew_Click
'' Description: Allow the user to create a New Item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    Dim LibMgrBridge As cLibManagerBridge
    Dim frm As Form
    Dim frmActiveChart As Form
    Dim obj As Object
    Dim strCBPrompt$
    Dim lLibraryID As Long              ' Library ID for the new library (if saved)

    Select Case vsTypeTabs.CurrTab
        Case Tabs(eTab_SymbolGroups)
            Me.Hide
            Set frm = New frmSymbolGroup
            frm.ShowMe AddSlash(App.Path) & "Custom\", ""
        
        Case Tabs(eTab_Criteria)
            If gdNumMatchingFiles(AddSlash(App.Path) & "Custom\Cus0*.SCN") >= 1 Then
                If Not HasGold(True, "Creating more custom Criteria") Then
                    Exit Sub
                End If
            End If
            Me.Hide
            Set frm = ActiveChart
            If IsFrmChart(frm) Then Set frmActiveChart = frm
            
            strCBPrompt = GetIniFileProperty("UseCondBuilderNewCriteria", "", "DontAsk", g.strIniFile)
            If Len(strCBPrompt) = 0 Then
                strCBPrompt = "Would you like to build the criteria's condition|using indicators from the active chart?"
                strCBPrompt = InfBox(strCBPrompt, "?", "+Yes|-No", "Condition Builder", , , , , , , , , True)
                If InStr(strCBPrompt, "-") > 0 Then
                    ' don't ask anymore, store for future use
                    Call SetIniFileProperty("UseCondBuilderNewRule", "N", "DontAsk", g.strIniFile)
                End If
            End If
            
            If UCase(Left(strCBPrompt, 1)) = "Y" And Not frmActiveChart Is Nothing Then
                frmConditionBuilder.ShowMe frmActiveChart.Chart, , eType_Criteria
            Else
                Set frm = New frmCriteria
                frm.ShowMe AddSlash(App.Path) & "Custom\", ""
            End If
            
        Case Tabs(eTab_Filters)
            If gdNumMatchingFiles(AddSlash(App.Path) & "Custom\Cus0*.FIL") >= 5 Then
                If Not HasGold(True, "Creating more custom Filters") Then
                    Exit Sub
                End If
            End If
            Me.Hide
            Set frm = New frmFilter
            frm.ShowMe AddSlash(App.Path) & "Custom\", ""
            
        Case Tabs(eTab_Systems)
            If HasPlatinum(True) Then
                Me.Hide
                Set frm = New frmSystemManager
                frm.ShowMe NextSystemID, , False, , True
            End If

        Case Tabs(eTab_Rules)
            If HasPlatinum(True) Then
                Me.Hide
                Set frm = ActiveChart
                If IsFrmChart(frm) Then Set frmActiveChart = frm
                
                strCBPrompt = GetIniFileProperty("UseCondBuilderNewRule", "", "DontAsk", g.strIniFile)
                If Len(strCBPrompt) = 0 Then
                    strCBPrompt = "Would you like to build the rule's condition|using indicators from the active chart?"
                    strCBPrompt = InfBox(strCBPrompt, "?", "+Yes|-No", "Condition Builder", , , , , , , , , True)
                    If InStr(strCBPrompt, "-") > 0 Then
                        ' don't ask anymore, store for future use
                        Call SetIniFileProperty("UseCondBuilderNewRule", "N", "DontAsk", g.strIniFile)
                    End If
                End If
                
                If UCase(Left(strCBPrompt, 1)) = "Y" And Not frmActiveChart Is Nothing Then
                    frmConditionBuilder.ShowMe frmActiveChart.Chart, , eType_Rule
                Else
                    Set frm = New frmRule
                    frm.ShowMe Name
                End If
            End If
        
        Case Tabs(eTab_Functions)
            If HasGold(True) Then
                Me.Hide
                Set frm = New frmFunctionMgrCT
                frm.ShowMe 0
            End If
        
        Case Tabs(eTab_Libraries)
            If HasGold(True) Then
                Me.Hide
                Set LibMgrBridge = GetLibMgrBridge
                lLibraryID = LibMgrBridge.CreateNewLibrary
                If lLibraryID > 0 Then
                    Screen.MousePointer = vbHourglass
                    
                    ' Reload the libraries grid...
                    LoadLibrariesGrid
                                        
                    ' Reload the Function and Rule tables in memory...
                    'LoadEngineFunctions
                    RefreshLibrary lLibraryID
                    RefreshLibrary kSN_UserLibrary
                    
                    ' Trigger a reload the next time they switch to these tabs...
                    fgFunctions.Rows = fgFunctions.FixedRows
                    fgRules.Rows = fgRules.FixedRows
                    fgSystems.Rows = fgSystems.FixedRows
                    
                    Screen.MousePointer = vbDefault
                End If
            End If
            
        Case Tabs(eTab_StrategyBaskets)
            If HasPlatinum(True) Then
                Me.Hide
                Set frm = New frmStrategyBasket
                frm.ShowMe
            End If
            
    End Select

ErrExit:
    Set obj = Nothing
    Set LibMgrBridge = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdNew.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewDLL_Click
'' Description: Allow the user to create a New DLL Function
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewDLL_Click()
On Error GoTo ErrSection:

    Dim frm As Form
    
    If HasPlatinum(True) Then
        Me.Hide
        Set frm = New frmFunctionMgr
        frm.ShowMe 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdNewDLL.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRename_Click
'' Description: Allow the user to rename a Symbol Group, Criteria, or Filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRename_Click()
On Error GoTo ErrSection:

    Rename

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.cmdRename_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgCriteria_AfterRowColChange
'' Description: As the user changes rows in the grid, change the Preview
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgCriteria_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Preview fgCriteria, eTab_Criteria

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgCriteria.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgCriteria_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgCriteria.Tag = "fgCriteria" & vbTab & Str(Col) & vbTab & Str(Order)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgCriteria.AfterSort", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgCriteria_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    ExtendCustomColumn fgCriteria, Col
    SaveCols eTab_Criteria

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgCriteria.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgCriteria_BeforeEdit
'' Description: Only allow the user to edit the Active column
'' Inputs:      Row and Column of cell being edited, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgCriteria_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> CGCol(eCGC_Active) Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgCriteria.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgCriteria_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size of column (in case is after the extended column)
    m.lPrevColWidth = fgCriteria.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgCriteria.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgCriteria_DblClick
'' Description: When the user double clicks on a Criteria, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgCriteria_DblClick()
On Error GoTo ErrSection:

    With fgCriteria
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
            cmdEdit_Click
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgCriteria.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        cmdDelete_Click
    ElseIf KeyCode = vbKeyInsert Then
        cmdNew_Click
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgCriteria.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgCriteria_KeyPress
'' Description: When the user presses Enter on a Criteria, edit it
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgCriteria_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        cmdEdit_Click
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgCriteria.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgCriteria_KeyUp(KeyCode As Integer, Shift As Integer)

    ' to select row as type letters
    If KeyCode >= 32 And Shift = 0 Then
        On Error Resume Next
        fgCriteria.Row = fgCriteria.Row
    End If

End Sub

Private Sub fgCriteria_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgCriteria
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows And Button = vbRightButton Then
            .RowSel = lMouseRow
            If .SelectedRows <= 1 Then .Row = lMouseRow
            
            SetUpPopup False
            PopupMenu mnuPopUp
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgCriteria.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgCriteria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgCriteria
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgCriteria_ValidateEdit
'' Description: Make sure that the user wants to do this
'' Inputs:      Row and Column of field to edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgCriteria_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim astrInactiveCriteria As New cGdArray
    Dim Filter As New cFilter
    Dim Criteria As New cCriteria
    Dim strMessage As String
    Dim lIndex As Long
    Dim bAsked As Boolean
    Dim strReturn As String

    With fgCriteria
        Set Criteria = g.SymbolPool.Criterias(.TextMatrix(Row, CGCol(eCGC_CriteriaID)))
        
        If Val(.EditText) = flexUnchecked Then
            bAsked = False
            For Each Filter In g.SymbolPool.Filters
                If Filter.IsActive Then
                    If Filter.CriteriaInFilter(Criteria.ID) = True Then
                        If bAsked = False Then
                            strReturn = AskBox("h=Criteria ; i=? ; b=+Yes|-No ; " & _
                                "There are active filters that are using||" & Criteria.Name & _
                                "||Deactivating this criteria will also" & _
                                "|deactivate the filters that use this criteria.||" & _
                                "Are you sure you want to do this?|")
                            bAsked = True
                            If strReturn = "N" Then
                                Cancel = True
                                Exit For
                            End If
                        End If
                        Filter.IsActive = False
                        Filter.ToFile
                    End If
                End If
            Next Filter
        End If
    
        If Not Cancel Then
            Criteria.IsActive = (CLng(Val(.EditText)) = flexChecked)
            Criteria.ToFile
        End If
    End With

ErrExit:
    Set astrInactiveCriteria = Nothing
    Set Filter = Nothing
    Set Criteria = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgCriteria.ValidateEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilters_AfterEdit
'' Description: Disable the edit button if the filter is deactivated
'' Inputs:      Row and Column being edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilters_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Enable cmdEdit, CheckedCell(fgFilters, Row, FiCol(eFIC_Active))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "fgFilters.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilters_AfterRowColChange
'' Description: As the user changes rows in the grid, change the Preview
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilters_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Preview fgFilters, eTab_Filters
    Enable cmdEdit, CheckedCell(fgFilters, NewRow, FiCol(eFIC_Active))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFilters.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgFilters_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgFilters.Tag = "fgFilters" & vbTab & Str(Col) & vbTab & Str(Order)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFilters.AfterSort", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgFilters_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    ExtendCustomColumn fgFilters, Col
    SaveCols eTab_Filters

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFilters.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilters_BeforeEdit
'' Description: Only allow the user to edit the Active column
'' Inputs:      Row and Column of cell being edited, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilters_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> FiCol(eFIC_Active) Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFilters.BeforeEdit", eGDRaiseError_Show
        Resume ErrExit
    
End Sub

Private Sub fgFilters_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size of column (in case is after the extended column)
    m.lPrevColWidth = fgFilters.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFilters.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilters_DblClick
'' Description: When the user double clicks on a Filter, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilters_DblClick()
On Error GoTo ErrSection:

    With fgFilters
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
            cmdEdit_Click
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFilters.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgFilters_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        cmdDelete_Click
    ElseIf KeyCode = vbKeyInsert Then
        cmdNew_Click
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgFilters.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilters_KeyPress
'' Description: When the user presses Enter on a Filter, edit it
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilters_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        cmdEdit_Click
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFilters.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgFilters_KeyUp(KeyCode As Integer, Shift As Integer)

    ' to select row as type letters
    If KeyCode >= 32 And Shift = 0 Then
        On Error Resume Next
        fgFilters.Row = fgFilters.Row
    End If

End Sub

Private Sub fgFilters_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgFilters
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows And Button = vbRightButton Then
            .RowSel = lMouseRow
            If .SelectedRows <= 1 Then .Row = lMouseRow
            
            SetUpPopup False
            PopupMenu mnuPopUp
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgFilters.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgFilters_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgFilters

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilters_ValidateEdit
'' Description: Make sure that the user wants to do this
'' Inputs:      Row and Column of field to edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilters_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim astrInactiveCriteria As New cGdArray
    Dim Filter As New cFilter
    Dim Criteria As New cCriteria
    Dim strMessage As String
    Dim lIndex As Long
    
    With fgFilters
        Set Filter = g.SymbolPool.Filters(.TextMatrix(Row, FiCol(eFIC_FilterID)))
        
        If Val(.EditText) = flexChecked Then
            Set astrInactiveCriteria = Filter.InactiveCriteria
            If astrInactiveCriteria.Size > 0 Then
                If astrInactiveCriteria.Size = 1 Then
                    strMessage = "The following criteria used by this filter|is currently inactive:||"
                Else
                    strMessage = "The following criteria used by this filter|are currently inactive:||"
                End If
                For lIndex = 0 To astrInactiveCriteria.Size - 1
                    strMessage = strMessage & Parse(astrInactiveCriteria(lIndex), "|", 2) & "|"
                Next lIndex
                If astrInactiveCriteria.Size = 1 Then
                    strMessage = strMessage & "|This criteria will be activated by activating the filter.|"
                Else
                    strMessage = strMessage & "|These criteria will be activated by activating the filter.|"
                End If
                
                If InfBox(strMessage, "!", "+OK|-Cancel", "Warning") = "C" Then
                    Cancel = True
                Else
                    For lIndex = 0 To astrInactiveCriteria.Size - 1
                        Set Criteria = g.SymbolPool.Criterias(Parse(astrInactiveCriteria(lIndex), "|", 1))
                        Criteria.IsActive = True
                        Criteria.ToFile
                    Next lIndex
                End If
            End If
        End If
        
        If Not Cancel Then
            Filter.IsActive = (CLng(Val(.EditText)) = flexChecked)
            Filter.ToFile
        End If
    End With

ErrExit:
    Set astrInactiveCriteria = Nothing
    Set Filter = Nothing
    Set Criteria = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFilters.ValidateEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFunctions_AfterCollapse
'' Description: Make sure that the background colors are set on a expand/collapse
'' Inputs:      Row expanded/collapsed, Whether expanded or collapsed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFunctions_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
On Error GoTo ErrSection:

    SetBackColors fgFunctions
    ExtendCustomColumn fgFunctions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFunctions.AfterCollapse", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgFunctions_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFunctions.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgFunctions_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgFunctions.Tag = "fgFunctions" & vbTab & Str(Col) & vbTab & Str(Order)
    SetBackColors fgFunctions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFunctions.AfterSort", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgFunctions_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    ExtendCustomColumn fgFunctions, Col
    SaveCols eTab_Functions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFunctions.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgFunctions_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> FGCol(eFGC_Favorites) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFunctions.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgFunctions_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size of column (in case is after the extended column)
    m.lPrevColWidth = fgFunctions.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFunctions.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFunctions_DblClick
'' Description: When the user double clicks on a Function, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFunctions_DblClick()
On Error GoTo ErrSection:

    With fgFunctions
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
        
            If .TextMatrix(.Row, FGCol(eFGC_TreeLevel)) = "1" Then
                If m.Mode = eAddMode_Select Then
                    cmdEdit_Click
                ElseIf m.Mode = eAddMode_List Then
                    cmdCancel_Click
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFunctions.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFunctions_KeyDown
'' Description: Handle user keystrokes appropriately
'' Inputs:      Code of the Key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFunctions_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    Select Case KeyCode
        Case vbKeyA
            If Shift And vbCtrlMask Then
                For lIndex = fgFunctions.FixedRows To fgFunctions.Rows - 1
                    fgFunctions.IsSelected(lIndex) = True
                Next lIndex
            End If
    
        Case vbKeyDelete
            cmdDelete_Click
    
        Case vbKeyInsert
            cmdNew_Click
    
        Case vbKeyRight
            If chkTreeFunctions.Value = vbChecked Then
                With fgFunctions
                    KeyCode = 0
                    If .RowOutlineLevel(.Row) = 0 Then
                        If .IsCollapsed(.Row) = flexOutlineExpanded Then
                            If .RowOutlineLevel(.Row + 1) >= .RowOutlineLevel(.Row) Then
                                .Row = .Row + 1
                                .ShowCell .Row, FGCol(eFGC_FunctionName)
                            End If
                        Else
                            .IsCollapsed(.Row) = flexOutlineExpanded
                        End If
                    ElseIf .Row + 1 < .Rows Then
                        If .RowOutlineLevel(.Row + 1) >= .RowOutlineLevel(.Row) Then
                            .Row = .Row + 1
                            .ShowCell .Row, FGCol(eFGC_FunctionName)
                        End If
                    End If
                End With
            End If
        
        Case vbKeyLeft
            If chkTreeFunctions.Value = vbChecked Then
                With fgFunctions
                    KeyCode = 0
                    If .Row >= .FixedRows And .Row < .Rows Then
                        If .RowOutlineLevel(.Row) = 1 Then
                            If .GetNodeRow(.Row, flexNTParent) <> -1 Then
                                .Row = .GetNodeRow(.Row, flexNTParent)
                                .ShowCell .Row, FGCol(eFGC_FunctionName)
                            End If
                        ElseIf .IsCollapsed(.Row) = flexOutlineExpanded Then
                            .IsCollapsed(.Row) = flexOutlineCollapsed
                        ElseIf .GetNodeRow(.Row, flexNTParent) <> -1 Then
                            .Row = .GetNodeRow(.Row, flexNTParent)
                            .ShowCell .Row, FGCol(eFGC_FunctionName)
                        End If
                    End If
                End With
            End If
    
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgFunctions.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFunctions_KeyPress
'' Description: When the user presses Enter on a Function, edit it
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFunctions_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        If m.Mode = eAddMode_Select Then
            cmdEdit_Click
        ElseIf m.Mode = eAddMode_List Then
            cmdCancel_Click
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFunctions.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgFunctions_KeyUp(KeyCode As Integer, Shift As Integer)

    ' to select row as type letters
    If KeyCode >= 32 And Shift = 0 Then
        On Error Resume Next
        fgFunctions.Row = fgFunctions.Row
    End If

End Sub

Private Sub fgFunctions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgFunctions
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows And Button = vbRightButton Then
            .RowSel = lMouseRow
            If .SelectedRows <= 1 Then .Row = lMouseRow
            
            SetUpPopup True
            PopupMenu mnuPopUp
            
            ' Need to let the Popup menu unload before calling a modal form
            ' because only one popup menu can be displayed at a time...
            If mnuPopUp.Tag = "Dependencies" Then mnuDependencies_Click
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgFunctions.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFunctions_AfterRowColChange
'' Description: When the user changes rows in the grid, change the preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFunctions_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    Preview fgFunctions, Tabs(eTab_Functions)
    SetButtons vsTypeTabs.CurrTab

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgFunctions.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgFunctions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    GridTooltip fgFunctions
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_AfterRowColChange
'' Description: As the user changes rows in the grid, change the Preview
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim PoolObject As cSymbolGroup      ' Object out of the pool

    With fgGroups
        Set PoolObject = g.SymbolPool.PoolObject("GRP:" & .TextMatrix(NewRow, GGCol(eGGC_GroupID)))
        
        ' Have seen one case on a customer's machine where the symbol group ID was blank and
        ' therefore the pool object was Nothing.  Put a check here to alleviate any "Object
        ' or With block variable not set" errors (01/13/2009 DAJ)...
        If Not PoolObject Is Nothing Then
            If PoolObject.GroupType = eGROUP_QuoteList Then
                cmdRename.Enabled = False
            Else
                cmdRename.Enabled = True
            End If
        Else
            cmdRename.Enabled = False
        End If
    End With
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgGroups_AfterRowColChange"

End Sub

Private Sub fgGroups_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgGroups.Tag = "fgGroups" & vbTab & Str(Col) & vbTab & Str(Order)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgGroups.AfterSort", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgGroups_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    ExtendCustomColumn fgGroups, Col
    SaveCols eTab_SymbolGroups

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgGroups.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgGroups_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size of column (in case is after the extended column)
    m.lPrevColWidth = fgGroups.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgGroups.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_DblClick
'' Description: When the user double clicks on a Symbol Group, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_DblClick()
On Error GoTo ErrSection:

    With fgGroups
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
    
            If m.Mode = eAddMode_Select Then
                cmdEdit_Click
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgGroups.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgGroups_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        cmdDelete_Click
    ElseIf KeyCode = vbKeyInsert Then
        cmdNew_Click
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgGroups.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_KeyPress
'' Description: When the user presses Enter on a Symbol Group, edit it
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        cmdEdit_Click
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgGroups.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgGroups_KeyUp(KeyCode As Integer, Shift As Integer)

    ' to select row as type letters
    If KeyCode >= 32 And Shift = 0 Then
        On Error Resume Next
        fgGroups.Row = fgGroups.Row
    End If

End Sub

Private Sub fgGroups_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgGroups
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows And Button = vbRightButton Then
            .RowSel = lMouseRow
            If .SelectedRows <= 1 Then .Row = lMouseRow
            
            SetUpPopup False
            PopupMenu mnuPopUp
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgGroups.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgGroups_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgGroups
    
End Sub

Private Sub fgLibraries_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgLibraries.Tag = "fgLibraries" & vbTab & Str(Col) & vbTab & Str(Order)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgLibraries.AfterSort", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgLibraries_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    ExtendCustomColumn fgLibraries, Col
    SaveCols eTab_Libraries

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgLibraries.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgLibraries_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size of column (in case is after the extended column)
    m.lPrevColWidth = fgLibraries.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgLibraries.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLibraries_DblClick
'' Description: When the user double clicks on a Function, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLibraries_DblClick()
On Error GoTo ErrSection:

    With fgLibraries
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
    
            If m.Mode = eAddMode_Select Then
                cmdEdit_Click
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgLibraries.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgLibraries_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        cmdDelete_Click
    ElseIf KeyCode = vbKeyInsert Then
        cmdNew_Click
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgLibraries.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLibraries_KeyPress
'' Description: When the user presses Enter on a Function, edit it
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLibraries_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        If m.Mode = eAddMode_Select Then
            cmdEdit_Click
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgLibraries.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgLibraries_KeyUp(KeyCode As Integer, Shift As Integer)

    ' to select row as type letters
    If KeyCode >= 32 And Shift = 0 Then
        On Error Resume Next
        fgLibraries.Row = fgLibraries.Row
    End If

End Sub

Private Sub fgLibraries_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgLibraries
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows And Button = vbRightButton Then
            .RowSel = lMouseRow
            If .SelectedRows <= 1 Then .Row = lMouseRow
            
            SetUpPopup False
            PopupMenu mnuPopUp
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgLibraries.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgLibraries_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    GridTooltip fgLibraries
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgRules_AfterCollapse
'' Description: After the user collapses/expands a row, recolor the grid rows
'' Inputs:      Row expanded/collapsed, Whether Expanded or Collapsed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgRules_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
On Error GoTo ErrSection:

    FilterRulesGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgRules.AfterCollapse", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgRules_AfterSort
'' Description: After the user sorts on a column, reset the background colors
'' Inputs:      Column sorted, Order sorted in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgRules_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgRules.Tag = "fgRules" & vbTab & Str(Col) & vbTab & Str(Order)
    SetBackColors fgRules

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgRules.AfterSort", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgRules_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    ExtendCustomColumn fgRules, Col
    SaveCols eTab_Rules

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgRules.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgRules_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size of column (in case is after the extended column)
    m.lPrevColWidth = fgRules.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgRules.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgRules_DblClick
'' Description: If the user double clicks on an entry select it for the mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgRules_DblClick()
On Error GoTo ErrSection:
    
    Dim lMouseRow As Long

    With fgRules
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            If .TextMatrix(lMouseRow, RGCol(eRGC_TreeLevel)) = 2 Then
                If m.Mode = eAddMode_Add Then
                    ' If the rule is a "Shared" rule, then add it to the system as shared
                    If .TextMatrix(lMouseRow, RGCol(eRGC_SystemNumber)) = "0" Then
                        cmdAddCopy_Click
                    
                    ' If the rule is a "Local" rule, add a local copy to the system
                    Else
                        cmdAddCopy_Click
                    End If
                Else
                    With fgRules
                        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                            .Row = lMouseRow
                            .RowSel = .Row
                            
                            cmdEdit_Click
                        End If
                    End With
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgRules.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgRules_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        cmdDelete_Click
    ElseIf KeyCode = vbKeyInsert Then
        cmdNew_Click
    ElseIf Shift And vbCtrlMask Then
        Select Case UCase(Chr(KeyCode))
        Case "D"
            KeyCode = 0
            mnuDependencies_Click
        End Select
    ElseIf chkTreeRules.Value = vbChecked Then
        With fgRules
            Select Case KeyCode
                Case vbKeyRight
                    KeyCode = 0
                    If .RowOutlineLevel(.Row) < 2 Then
                        If .IsCollapsed(.Row) = flexOutlineExpanded Then
                            If .RowOutlineLevel(.Row + 1) >= .RowOutlineLevel(.Row) Then
                                .Row = .Row + 1
                                .ShowCell .Row, RGCol(eRGC_RuleName)
                            End If
                        Else
                            .IsCollapsed(.Row) = flexOutlineExpanded
                        End If
                    ElseIf .Row + 1 < .Rows Then
                        If .RowOutlineLevel(.Row + 1) >= .RowOutlineLevel(.Row) Then
                            .Row = .Row + 1
                            .ShowCell .Row, RGCol(eRGC_RuleName)
                        End If
                    End If
                
                Case vbKeyLeft
                    KeyCode = 0
                    If .Row >= .FixedRows And .Row < .Rows Then
                        If .RowOutlineLevel(.Row) = 2 Then
                            If .GetNodeRow(.Row, flexNTParent) <> -1 Then
                                .Row = .GetNodeRow(.Row, flexNTParent)
                                .ShowCell .Row, RGCol(eRGC_RuleName)
                            End If
                        ElseIf .IsCollapsed(.Row) = flexOutlineExpanded Then
                            .IsCollapsed(.Row) = flexOutlineCollapsed
                        ElseIf .GetNodeRow(.Row, flexNTParent) <> -1 Then
                            .Row = .GetNodeRow(.Row, flexNTParent)
                            .ShowCell .Row, RGCol(eRGC_RuleName)
                        End If
                    End If
            
            End Select
        End With
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgRules.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgRules_KeyPress
'' Description: If the user presses Enter on an entry select it for the mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgRules_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        With fgRules
            If m.Mode = eAddMode_Add Then
                ' If the rule is a "Shared" rule, then add it to the system as shared
                If .TextMatrix(.RowSel, RGCol(eRGC_SystemNumber)) = "0" Then
                    cmdAddCopy_Click
                
                ' If the rule is a "Local" rule, add a local copy to the system
                Else
                    cmdAddCopy_Click
                End If
            ElseIf m.Mode = eAddMode_Select Then
                cmdEdit_Click
            Else
                cmdCancel_Click
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgRules.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgRules_KeyUp(KeyCode As Integer, Shift As Integer)

    ' to select row as type letters
    If KeyCode >= 32 And Shift = 0 Then
        On Error Resume Next
        fgRules.Row = fgRules.Row
    End If

End Sub

Private Sub fgRules_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgRules
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows And Button = vbRightButton Then
            .RowSel = lMouseRow
            If .SelectedRows <= 1 Then .Row = lMouseRow
            
            SetUpPopup True
            PopupMenu mnuPopUp
            
            ' Need to let the Popup menu unload before calling a modal form
            ' because only one popup menu can be displayed at a time...
            If mnuPopUp.Tag = "Dependencies" Then mnuDependencies_Click
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgRules.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgRules_AfterRowColChange
'' Description: When the user changes rows in the grid, change the preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgRules_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Preview fgRules, eTab_Rules
    SetButtons vsTypeTabs.CurrTab

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgRules.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgRules_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgRules
    
End Sub

Private Sub fgStrategyBaskets_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgStrategyBaskets.Tag = "fgStrategyBaskets" & vbTab & Str(Col) & vbTab & Str(Order)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgStrategyBaskets.AfterSort", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgStrategyBaskets_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    SaveCols eTab_StrategyBaskets

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgStrategyBaskets.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgStrategyBaskets_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size of column (in case is after the extended column)
    m.lPrevColWidth = fgStrategyBaskets.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgStrategyBaskets.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBaskets_DblClick
'' Description: When the user double clicks on a System Run, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBaskets_DblClick()
On Error GoTo ErrSection:

    With fgStrategyBaskets
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
            cmdEdit_Click
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgStrategyBaskets.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgStrategyBaskets_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        cmdDelete_Click
    ElseIf KeyCode = vbKeyInsert Then
        cmdNew_Click
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgStrategyBaskets.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBaskets_KeyPress
'' Description: When the user hits Enter on a System Run, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBaskets_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:
    
    If KeyAscii = vbKeyReturn Then
        cmdEdit_Click
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgStrategyBaskets.KeyPress", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgStrategyBaskets_KeyUp(KeyCode As Integer, Shift As Integer)

    ' to select row as type letters
    If KeyCode >= 32 And Shift = 0 Then
        On Error Resume Next
        fgStrategyBaskets.Row = fgStrategyBaskets.Row
    End If

End Sub

Private Sub fgStrategyBaskets_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgStrategyBaskets
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows And Button = vbRightButton Then
            .RowSel = lMouseRow
            If .SelectedRows <= 1 Then .Row = lMouseRow
            
            SetUpPopup False
            PopupMenu mnuPopUp
        
            ' Need to let the Popup menu unload before calling a modal form
            ' because only one popup menu can be displayed at a time...
            Select Case mnuPopUp.Tag
                Case "New":  mnuNew_Click
                Case "Edit": mnuEdit_Click
            End Select
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgStrategyBaskets.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgStrategyBaskets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgStrategyBaskets

End Sub

Private Sub fgSystems_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    fgSystems.Tag = "fgSystems" & vbTab & Str(Col) & vbTab & Str(Order)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgSystems.AfterSort", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgSystems_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    ExtendCustomColumn fgSystems, Col
    SaveCols eTab_Systems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgSystems.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgSystems_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size of column (in case is after the extended column)
    m.lPrevColWidth = fgSystems.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgSystems.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSystems_DblClick
'' Description: When the user double clicks on a system, edit that system
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSystems_DblClick()
On Error GoTo ErrSection:
    
    With fgSystems
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
            cmdEdit_Click
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgSystems.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgSystems_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        cmdDelete_Click
    ElseIf KeyCode = vbKeyInsert Then
        cmdNew_Click
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgSystems.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSystems_KeyPress
'' Description: When the user hits Enter on a system, edit that system
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSystems_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        cmdEdit_Click
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgSystems.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgSystems_KeyUp(KeyCode As Integer, Shift As Integer)

    ' to select row as type letters
    If KeyCode >= 32 And Shift = 0 Then
        On Error Resume Next
        fgSystems.Row = fgSystems.Row
    End If

End Sub

Private Sub fgSystems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgSystems
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows And Button = vbRightButton Then
            .RowSel = lMouseRow
            If .SelectedRows <= 1 Then .Row = lMouseRow
            
            SetUpPopup True
            PopupMenu mnuPopUp
        
            ' Need to let the Popup menu unload before calling a modal form
            ' because only one popup menu can be displayed at a time...
            If mnuPopUp.Tag = "Dependencies" Then mnuDependencies_Click
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.fgSystems.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSystems_AfterRowColChange
'' Description: When the user changes rows in the grid, change the preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSystems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Preview fgSystems, eTab_Systems
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgSystems.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLibraries_AfterRowColChange
'' Description: When the user changes rows in the grid, change the preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLibraries_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Preview fgLibraries, eTab_Libraries
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.fgLibraries.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgSystems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgSystems
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, set the focus appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    vsTypeTabs_Click
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.Form.Activate", eGDRaiseError_Show
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
    RaiseError "frmToolbox.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, set the icon and location
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strTemp As String
    Dim rs As Recordset
    Dim strFont As String
    Dim lIndex As Long

    Me.Icon = Picture16(ToolbarIcon("ID_Toolbox"), , True)
    
    g.Styler.StyleForm Me
    
    cmdCancel.Cancel = True
    
    With vsTypeTabs
        .FirstTab = 0
        .TabPicture(Tabs(eTab_SymbolGroups)) = Picture16(ToolbarIcon("ID_SymbolGroups"))
        .TabPicture(Tabs(eTab_Criteria)) = Picture16(ToolbarIcon("ID_Criteria"))
        .TabPicture(Tabs(eTab_Filters)) = Picture16(ToolbarIcon("ID_Filters"))
        .TabPicture(Tabs(eTab_Functions)) = Picture16(ToolbarIcon("ID_Functions"))
        .TabPicture(Tabs(eTab_Rules)) = Picture16(ToolbarIcon("ID_Rules"))
        .TabPicture(Tabs(eTab_Systems)) = Picture16(ToolbarIcon("ID_Strategies"))
        .TabPicture(Tabs(eTab_StrategyBaskets)) = Picture16(ToolbarIcon("ID_StrategyBaskets"))
        .TabPicture(Tabs(eTab_Libraries)) = Picture16(ToolbarIcon("ID_Libraries"))
    End With
    
    ' set form location
    If Screen.Width <= 12200 Then
        ' if low-res screen (800x600), trim "Strategy Baskets" to "Baskets" so tabs will fit better
        strTemp = Replace(vsTypeTabs.Caption, "Strategy ", "")
        vsTypeTabs.Caption = strTemp
    End If
    strTemp = GetIniFileProperty("FormLoc", "", "Toolbox", g.strIniFile)
    If Len(strTemp) = 0 Then
        If Screen.Width <= 12200 Then
            Me.Width = 11880
        Else
            Me.Width = 13000
        End If
        CenterTheForm Me
    Else
        SetFormPlacement Me, strTemp, "WH"
        CenterTheForm Me
    End If
    
    fgRules.Cols = RGCol(eRGC_NumCols)
    fgFunctions.Cols = FGCol(eFGC_NumCols)
    
    ' Fill the library combo box
    ''Set rs = g.dbNav.OpenRecordset("qryLibrarysByName", dbOpenSnapshot)
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] WHERE [Ignore]=0;", dbOpenDynaset)
    ValidateCheckSums rs, "tblLibrarys"
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        If rs!Ignore = 0 And rs!CheckSum <> 0.5 Then
            cboLibrary.AddItem rs!LibraryName
        End If
        rs.MoveNext
    Loop
    cboLibrary.ListIndex = 0
    
    strTemp = GetIniFileProperty("RulesFilter", "", "Toolbox", g.strIniFile)
    If strTemp <> "" Then
        ''optAll = Parse(strTemp, ";", 1)
        ''optLong = Parse(strTemp, ";", 2)
        ''optLongExit = Parse(strTemp, ";", 3)
        ''optShort = Parse(strTemp, ";", 4)
        ''optShortExit = Parse(strTemp, ";", 5)
        optAll = True
        chkLibrary = Parse(strTemp, ";", 6)
        On Error Resume Next
        cboLibrary.Text = Parse(strTemp, ";", 7)
        On Error GoTo ErrSection:
        ''If Parse(strTemp, ";", 9) = "0" Then
            chkFavorites = vbUnchecked
        ''Else
        ''    chkFavorites = vbChecked '(default)
        ''End If
        If Parse(strTemp, ";", 10) = "0" Then
            chkTreeRules = vbUnchecked
        Else
            chkTreeRules = vbChecked '(default)
        End If
    End If
    
    strTemp = GetIniFileProperty("FunctionsFilter", "", "Toolbox", g.strIniFile)
    If Len(strTemp) > 0 Then
        If Parse(strTemp, ";", 1) = "0" Then
            chkTreeFunctions.Value = vbUnchecked
        Else
            chkTreeFunctions.Value = vbChecked
        End If
        
        If Parse(strTemp, ";", 2) = "0" Then
            chkFuncFav.Value = vbUnchecked
        Else
            chkFuncFav.Value = vbChecked
        End If
    End If
    
    mnuPopUp.Visible = False
    strFont = GetIniFileProperty("Toolbox", "", "Fonts", g.strIniFile)
    If strFont <> "" Then
        FontFromString fgGroups.Font, strFont
        FontFromString fgCriteria.Font, strFont
        FontFromString fgFilters.Font, strFont
        FontFromString fgFunctions.Font, strFont
        FontFromString fgRules.Font, strFont
        FontFromString fgSystems.Font, strFont
        FontFromString fgLibraries.Font, strFont
        FontFromString fgStrategyBaskets.Font, strFont
    End If
    
    Set m.abAutoSize = New cGdArray
    m.abAutoSize.Create eGDARRAY_TinyInts, Tabs(eTab_NumTabs)
    For lIndex = 0 To Tabs(eTab_NumTabs) - 1
        m.abAutoSize(lIndex) = 1
    Next lIndex
    
    chkFavorites.Visible = False
    fraRuleFilters.Width = cboLibrary.Left + cboLibrary.Width + 120
    
    chkTreeRules.ToolTipText = "Toggle between Tree View and Grid View"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(Optional ByVal lStartTab As eAddFormTabs = -1, Optional ByVal strInitialSelect As String = "")
On Error GoTo ErrSection:

    Static nPrevTab&, nPrevRow&, nPrevTopRow&
    
    If lStartTab < 0 Then lStartTab = nPrevTab
    If lStartTab < 0 Then lStartTab = eTab_SymbolGroups

    Screen.MousePointer = vbHourglass
    m.Mode = eAddMode_Select

    If g.nColorTheme = kDarkThemeColor Or g.nColorTheme = vbWhite Then
        'JM 12-18-2015: need to call this here because the grids are getting loaded before showing the form
        FixFormControls Me, ALT_GRID_ROW_COLOR
    End If

    With vsTypeTabs
        .TabVisible(Tabs(eTab_SymbolGroups)) = True
        InitSymbolGroupGrid
        
        .TabVisible(Tabs(eTab_Criteria)) = True
        InitCriteriaGrid
        
        .TabVisible(Tabs(eTab_Filters)) = True
        InitFiltersGrid
    
        ' Load the Systems Grid
        .TabVisible(Tabs(eTab_Systems)) = (ExtremeCharts <> 1)
        InitSystemsGrid
        
        ' Load the Rules Grid
        .TabVisible(Tabs(eTab_Rules)) = (ExtremeCharts <> 1)
        Set m.SystemRules = New cRules
        InitRulesGrid
        
        ' Load the Functions Grid
        .TabVisible(Tabs(eTab_Functions)) = True
        InitFunctionsGrid
        
        ' Load the Libraries Grid
        .TabVisible(Tabs(eTab_Libraries)) = True
        InitLibrariesGrid
        
        .TabVisible(Tabs(eTab_StrategyBaskets)) = (ExtremeCharts <> 1)
        InitStrategyBasketsGrid
        
        FixFilterDisplay
        
        Me.Caption = "Select item to edit ..."
        If Len(strInitialSelect) > 0 Then
            m.strInitialSelect = strInitialSelect
              nPrevRow = 0
        End If
        m.strInitialSelect = UCase(Trim(m.strInitialSelect))
        .CurrTab = -1 '(do this so "Switch" will get called)
        .CurrTab = lStartTab
        m.strInitialSelect = ""
               
        ' show form (unless was cancelled)
        Screen.MousePointer = vbDefault
        If .CurrTab < 0 Then ' <> lStartTab Then
            Unload Me
            Exit Sub '(switching to specified tab must have been cancelled)
        End If
        ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR
        
        ' save current tab and row for next time
        nPrevTab = .CurrTab
        
        Select Case .CurrTab
        Case Tabs(eTab_SymbolGroups)
            With Me.fgGroups
                If .Row >= .FixedRows Then
                    m.strInitialSelect = .TextMatrix(.Row, GGCol(eGGC_GroupID))
                End If
                m.strPrevSort = .Tag
            End With
        Case Tabs(eTab_Criteria)
            With Me.fgCriteria
                If .Row >= .FixedRows Then
                    m.strInitialSelect = .TextMatrix(.Row, CGCol(eCGC_CriteriaID))
                End If
                m.strPrevSort = .Tag
            End With
        Case Tabs(eTab_Filters)
            With Me.fgFilters
                If .Row >= .FixedRows Then
                    m.strInitialSelect = .TextMatrix(.Row, FiCol(eFIC_FilterID))
                End If
                m.strPrevSort = .Tag
            End With
        Case Tabs(eTab_Functions)
            With Me.fgFunctions
                If .Row >= .FixedRows Then
                    m.strInitialSelect = .TextMatrix(.Row, FGCol(eFGC_FunctionID))
                End If
                m.strPrevSort = .Tag
            End With
        Case Tabs(eTab_Rules)
            With Me.fgRules
                If .Row >= .FixedRows Then
                    m.strInitialSelect = .TextMatrix(.Row, RGCol(eRGC_RuleID))
                End If
                m.strPrevSort = .Tag
            End With
        Case Tabs(eTab_Systems)
            With Me.fgSystems
                If .Row >= .FixedRows Then
                    m.strInitialSelect = .TextMatrix(.Row, SGCol(eSGC_SystemNumber))
                End If
                m.strPrevSort = .Tag
            End With
        Case Tabs(eTab_Libraries)
            With Me.fgLibraries
                If .Row >= .FixedRows Then
                    m.strInitialSelect = .TextMatrix(.Row, LGCol(eLGC_LibraryID))
                End If
                m.strPrevSort = .Tag
            End With
        Case Tabs(eTab_StrategyBaskets)
            With Me.fgStrategyBaskets
                If .Row >= .FixedRows Then
                    m.strInitialSelect = .TextMatrix(.Row, SBCol(eSBC_Name))
                End If
                m.strPrevSort = .Tag
            End With
        End Select
        
    End With
    
ErrExit:
    Screen.MousePointer = vbDefault
    Unload Me
    
    If FileExist(AddSlash(App.Path) & "RestoreMDB.FLG") Then
        frmMain.tmrMain.Tag = "QUIT"
    End If
    Exit Sub
    
ErrSection:
    Screen.MousePointer = vbDefault
    RaiseError "frmToolbox.ShowMe", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, unload the form in Cancel mode
'' Inputs:      Whether or not to Cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If UnloadMode = vbFormControlMenu Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls on the form as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next
        
    Dim lWidth As Long, lSpace As Long, lTop As Long
    
    If LimitFormSize(Me, fraButtons.Width * 6, fraButtons.Height * 1.5) Then Exit Sub
   
    With fraButtons
        .Move Me.ScaleWidth - .Width '- vsTypeTabs.Left
    End With
    
    With vsTypeTabs
        .Move .Left, .Top, Me.ScaleWidth - fraButtons.Width - (.Left * 3), _
                Me.ScaleHeight - (.Top * 2)
        .Refresh
    End With
    
    lSpace = 120
    lWidth = fraSymbolGroups.Width - lSpace * 2
    
    ' Symbol Group Tab
    With txtPreview(Tabs(eTab_SymbolGroups))
        .Visible = False
        '.Move lSpace, fraSymbolGroups.Height - .Height - lSpace, lWidth
    End With
    With fgGroups
        .Move lSpace, .Top, lWidth, fraSymbolGroups.Height - .Top - lSpace
        ExtendCustomColumn fgGroups
    End With
    
    ' Criteria Tab
    With txtPreview(Tabs(eTab_Criteria))
        .Move lSpace, fraCriteria.Height - .Height - lSpace, lWidth
    End With
    With fgCriteria
        .Move lSpace, .Top, lWidth, txtPreview(Tabs(eTab_Criteria)).Top - .Top - lSpace
        ExtendCustomColumn fgCriteria
    End With
    
    ' Filters Tab
    With txtPreview(Tabs(eTab_Filters))
        .Visible = False
        '.Move lSpace, fraFilters.Height - .Height - lSpace, lWidth
    End With
    With fgFilters
        .Move lSpace, .Top, lWidth, fraFilters.Height - .Top - lSpace
        ExtendCustomColumn fgFilters
    End With
    ''CenterTheControl fraFilterMsg, fraFilters
    
    ' System Runs Tab
    With fgStrategyBaskets
        .Move lSpace, .Top, lWidth, fraStrategyBaskets.Height - .Top - lSpace
    End With
    
    ' Systems Tab
    With txtPreview(Tabs(eTab_Systems))
        .Move lSpace, fraSystems.Height - .Height - lSpace, lWidth
    End With
    With fgSystems
        .Move lSpace, .Top, lWidth, txtPreview(Tabs(eTab_Systems)).Top - .Top - lSpace
        ExtendCustomColumn fgSystems
    End With
    
    ' Rules Tab
    With fraRuleFilter
        If lWidth > lSpace * 2 + fraRuleTypes.Width + fraRuleFilters.Width Then
            .Move lSpace, .Top, lWidth, fraRuleFilters.Height * 2
            fraRuleFilters.Move fraRuleTypes.Left + fraRuleTypes.Width, fraRuleTypes.Top - 30
        Else
            .Move lSpace, .Top, lWidth, fraRuleFilters.Height * 3
            fraRuleFilters.Move fraRuleTypes.Left, fraRuleTypes.Top + fraRuleTypes.Height + 30
        End If
    End With
    With txtPreview(Tabs(eTab_Rules))
        .Move lSpace, fraRules.Height - .Height - lSpace, lWidth
    End With
    With fgRules
        lTop = fraRuleFilter.Top + fraRuleFilter.Height + lSpace
        .Move lSpace, lTop, lWidth, txtPreview(Tabs(eTab_Rules)).Top - lTop - lSpace
        ExtendCustomColumn fgRules
    End With

    ' Functions Tab
    With txtPreview(Tabs(eTab_Functions))
        .Move lSpace, fraFunctions.Height - .Height - lSpace, lWidth
    End With
    If Label4.Width + fraFunctionFilter.Width + (Label4.Left * 3) < fraFunctions.Width Then
        With fgFunctions
            .Move lSpace, 360, lWidth, txtPreview(Tabs(eTab_Functions)).Top - 360 - lSpace
            ExtendCustomColumn fgFunctions
        End With
        With fraFunctionFilter
            .Move Label4.Width + (Label4.Left * 2), Label4.Top + 30
        End With
    Else
        With fgFunctions
            .Move lSpace, 600, lWidth, txtPreview(Tabs(eTab_Functions)).Top - 600 - lSpace
            ExtendCustomColumn fgFunctions
        End With
        With fraFunctionFilter
            .Move Label4.Left, 330
        End With
    End If
    
    ' Libraries Tab
    With txtPreview(Tabs(eTab_Libraries))
        .Move lSpace, fraLibraries.Height - .Height - lSpace, lWidth
    End With
    With fgLibraries
        .Move lSpace, .Top, lWidth, txtPreview(Tabs(eTab_Libraries)).Top - .Top - lSpace
        ExtendCustomColumn fgLibraries
    End With
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterRulesGrid
'' Description: Hide/Show Rules according to the filters the user chose
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterRulesGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    Dim lRow As Long                    ' Index into a for loop
    
    With fgRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lRow = .FixedRows To .Rows - 1
            Select Case True
                Case optAll
                    .RowHidden(lRow) = False
                Case optLong
                    If .TextMatrix(lRow, RGCol(eRGC_RuleType)) = "Long Entry" Then
                        .RowHidden(lRow) = False
                    Else
                        .RowHidden(lRow) = True
                    End If
                Case optLongExit
                    If .TextMatrix(lRow, RGCol(eRGC_RuleType)) = "Long Exit" Then
                        .RowHidden(lRow) = False
                    Else
                        .RowHidden(lRow) = True
                    End If
                Case optShort
                    If .TextMatrix(lRow, RGCol(eRGC_RuleType)) = "Short Entry" Then
                        .RowHidden(lRow) = False
                    Else
                        .RowHidden(lRow) = True
                    End If
                Case optShortExit
                    If .TextMatrix(lRow, RGCol(eRGC_RuleType)) = "Short Exit" Then
                        .RowHidden(lRow) = False
                    Else
                        .RowHidden(lRow) = True
                    End If
            End Select
            
            If chkLibrary = vbChecked Then
                If .TextMatrix(lRow, RGCol(eRGC_LibraryName)) <> cboLibrary.Text Then
                    If .RowHidden(lRow) = False Then .RowHidden(lRow) = True
                End If
            End If
            
            If chkFavorites = vbChecked Then
                If Val(.TextMatrix(lRow, RGCol(eRGC_SystemNumber))) <> 0 Then
                    If .RowHidden(lRow) = False Then .RowHidden(lRow) = True
                End If
            End If
            
            If chkTreeRules.Value = vbChecked Then
                If Val(.TextMatrix(lRow, RGCol(eRGC_TreeLevel))) < 2 Then
                    .RowHidden(lRow) = False
                End If
                .RowOutlineLevel(lRow) = Val(.TextMatrix(lRow, RGCol(eRGC_TreeLevel)))
                .IsSubtotal(lRow) = True
                If Val(.TextMatrix(lRow, RGCol(eRGC_TreeLevel))) > 0 Then
                    If .IsCollapsed(.GetNodeRow(lRow, flexNTParent)) Then
                        .RowHidden(lRow) = True
                    End If
                End If
            Else
                If Val(.TextMatrix(lRow, RGCol(eRGC_TreeLevel))) < 2 Then
                    .RowHidden(lRow) = True
                End If
                .RowOutlineLevel(lRow) = 0
                .IsSubtotal(lRow) = False
            End If
            
            If .TextMatrix(lRow, RGCol(eRGC_RuleName)) = "RULES USED IN STRATEGIES" Then
                m.lUsedInStrategiesRow = lRow
            End If
        Next lRow
        
        ' Select the first visible row
        For lRow = .FixedRows To .Rows - 1
            If .RowHidden(lRow) = False Then
                .Row = lRow
                .RowSel = lRow
                Exit For
            End If
        Next lRow
        
        '.AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn fgRules
        SetBackColors fgRules
        .Redraw = lRedraw
    End With
    
    SetButtons vsTypeTabs.CurrTab

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.FilterRulesGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, do any necessary persistence
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim strTemp As String               ' Temporary string variable
    
    strTemp = optAll & ";" & optLong & ";" & optLongExit & ";" & optShort & ";" & optShortExit
    strTemp = strTemp & ";" & chkLibrary & ";" & cboLibrary.Text
    strTemp = strTemp & ";0;" & Str(chkFavorites) & ";" & Str(chkTreeRules)
    SetIniFileProperty "RulesFilter", strTemp, "Toolbox", g.strIniFile
    
    strTemp = Str(chkTreeFunctions.Value) & ";" & Str(chkFuncFav.Value)
    SetIniFileProperty "FunctionsFilter", strTemp, "Toolbox", g.strIniFile
    
    SetIniFileProperty "FormLoc", GetFormPlacement(Me), "Toolbox", g.strIniFile
    SetIniFileProperty "Toolbox", FontToString(fgGroups.Font), "Fonts", g.strIniFile
    
    SaveFunctionFavorites

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuAdd_Click()
    cmdAdd_Click
End Sub

Private Sub mnuAddCopy_Click()
    cmdAddCopy_Click
End Sub

Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    Dim strFont As String

    If ChangeGridFont(fgGroups) Then
        fgCriteria.Font = fgGroups.Font
        fgCriteria.Font = fgCriteria.Font
        
        fgFilters.Font = fgGroups.Font
        fgFilters.Font = fgFilters.Font
        
        fgFunctions.Font = fgGroups.Font
        fgFunctions.Font = fgFunctions.Font
        
        fgRules.Font = fgGroups.Font
        fgRules.Font = fgRules.Font
        
        fgSystems.Font = fgGroups.Font
        fgSystems.Font = fgSystems.Font
        
        fgLibraries.Font = fgGroups.Font
        fgLibraries.Font = fgLibraries.Font
        
        fgStrategyBaskets.Font = fgGroups.Font
        fgStrategyBaskets.Font = fgStrategyBaskets.Font
        
        Form_Resize
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCreateAutoTrade_Click
'' Description: Allow the user to create an automated trading item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCreateAutoTrade_Click()
On Error GoTo ErrSection:

    Dim lID As Long                     ' ID off of the current row in the appropriate grid
    Dim Basket As cStrategyBasket       ' Strategy basket item
    Dim TradeItem As cAutoTradeItem     ' Automated trading item

    Set TradeItem = New cAutoTradeItem
    
    Select Case vsTypeTabs.CurrTab
        Case Tabs(eTab_Systems)
            With fgSystems
                If (.RowSel >= .FixedRows) And (.RowSel < .Rows) Then
                    lID = CLng(.TextMatrix(.RowSel, SGCol(eSGC_SystemNumber)))
                    TradeItem.StrategyID = lID
                End If
            End With
        
        Case Tabs(eTab_StrategyBaskets)
            With fgStrategyBaskets
                If (.RowSel >= .FixedRows) And (.RowSel < .Rows) Then
                    Set Basket = .RowData(.RowSel)
                    lID = Basket.ID
                    TradeItem.StrategyBasketID = lID
                End If
            End With
        
    End Select
    
    If lID > 0 Then
        frmAutoTradeItem.ShowMe TradeItem, False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.mnuCreateAutoTrade_Click"
    
End Sub

Private Sub mnuDelete_Click()
    cmdDelete_Click
End Sub

Private Sub mnuDependencies_Click()
On Error GoTo ErrSection:

    Dim lID As Long
    Dim strName As String
    
    If mnuPopUp.Tag = "" Then
        mnuPopUp.Tag = "Dependencies"
    Else
        Select Case vsTypeTabs.CurrTab
            Case Tabs(eTab_Functions)
                With fgFunctions
                    lID = CLng(.TextMatrix(.RowSel, FGCol(eFGC_FunctionID)))
                    strName = .TextMatrix(.RowSel, FGCol(eFGC_FunctionName))
                    If g.Security.CanEdit(.TextMatrix(.RowSel, FGCol(eFGC_SecurityLevel)), .TextMatrix(.RowSel, FGCol(eFGC_Password)), strName) Then
                        frmDepends.ShowMe eDepends_Function, lID, strName
                    End If
                End With
            
            Case Tabs(eTab_Rules)
                With fgRules
                    lID = CLng(.TextMatrix(.RowSel, RGCol(eRGC_RuleID)))
                    strName = .TextMatrix(.RowSel, RGCol(eRGC_RuleName))
                    If g.Security.CanEdit(.TextMatrix(.RowSel, RGCol(eRGC_SecurityLevel)), .TextMatrix(.RowSel, RGCol(eRGC_Password)), strName) Then
                        frmDepends.ShowMe eDepends_Rule, lID, strName
                    End If
                End With
            
            Case Tabs(eTab_Systems)
                With fgSystems
                    lID = CLng(.TextMatrix(.RowSel, SGCol(eSGC_SystemNumber)))
                    strName = .TextMatrix(.RowSel, SGCol(eSGC_SystemDesc))
                    If g.Security.CanEdit(.TextMatrix(.RowSel, SGCol(eSGC_SecurityLevel)), .TextMatrix(.RowSel, SGCol(eSGC_Password)), strName) Then
                        frmDepends.ShowMe eDepends_System, lID, strName
                    End If
                End With
        
        End Select
        mnuPopUp.Tag = ""
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.mnuDependencies.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuEdit_Click()
On Error GoTo ErrSection:

    If mnuPopUp.Tag = "" And vsTypeTabs.CurrTab = Tabs(eTab_StrategyBaskets) Then
        mnuPopUp.Tag = "Edit"
    Else
        cmdEdit_Click
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.mnuEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuExport_Click
'' Description: Export the library containing the selected item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuExport_Click()
On Error GoTo ErrSection:

    Export

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.mnuExport.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuExportAll_Click()
On Error GoTo ErrSection:

    Dim astrLibraryIDs As New cGdArray  ' Array of Library ID's to export
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLibraryID As Long              ' Library ID to export
    Dim LibMgrBridge As cLibManagerBridge
    
    If astrLibraryIDs.FromFile(AddSlash(App.Path) & "ExportAll.TXT") = False Then
        astrLibraryIDs.Size = 0
        With fgLibraries
            For lIndex = .FixedRows To .Rows - 1
                astrLibraryIDs.Add .TextMatrix(lIndex, LGCol(eLGC_LibraryID))
            Next lIndex
        End With
    End If
    
    Set LibMgrBridge = GetLibMgrBridge
    
    For lIndex = 0 To astrLibraryIDs.Size - 1
        lLibraryID = CLng(ValOfText(astrLibraryIDs(lIndex)))
        'LibMgrBridge.ShowPackager lLibraryID
        'If LibMgrBridge.ExportOK = False Then Exit For
        If LibMgrBridge.ExportAuto(lLibraryID) = False Then Exit For
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.mnuExportAll.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuExportList_Click
'' Description: Allow the user to export a list of strategies
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuExportList_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFile As cGdArray            ' Array to dump to a file
    
    Set astrFile = New cGdArray
    astrFile.Create eGDARRAY_Strings
    
    With fgSystems
        For lIndex = .FixedRows To .Rows - 1
            astrFile.Add .TextMatrix(lIndex, SGCol(eSGC_SystemDesc)) & vbTab & .TextMatrix(lIndex, SGCol(eSGC_LibraryName))
        Next lIndex
    End With
    
    astrFile.ToFile AddSlash(App.Path) & "Strategies.LST"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.frmExportList_Click"
    
End Sub

Private Sub mnuImport_Click()
    cmdImport_Click
End Sub

Private Sub mnuInfo_Click()
    cmdInfo_Click
End Sub

Private Sub mnuNew_Click()
On Error GoTo ErrSection:

    If mnuPopUp.Tag = "" And vsTypeTabs.CurrTab = Tabs(eTab_StrategyBaskets) Then
        mnuPopUp.Tag = "New"
    Else
        cmdNew_Click
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.mnuNew.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuNewDLL_Click()
    cmdNewDLL_Click
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRename_Click
'' Description: Allow the user to rename a Symbol Group, Criteria, or Filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRename_Click()
On Error GoTo ErrSection:

    Rename

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.mnuRename_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRenameFile_Click
'' Description: Allow the user to rename a Criteria file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRenameFile_Click()
On Error GoTo ErrSection:

    RenameFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.mnuRenameFile_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAll_Click
'' Description: Show all of the rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAll_Click()
On Error GoTo ErrSection:

    FilterRulesGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.optAll.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLong_Click
'' Description: Show all of the Long Entry rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLong_Click()
On Error GoTo ErrSection:

    FilterRulesGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.optLong.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLongExit_Click
'' Description: Show all of the Long Exit rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLongExit_Click()
On Error GoTo ErrSection:

    FilterRulesGrid
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.optLongExit.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optShort_Click
'' Description: Show all of the Short Entry rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optShort_Click()
On Error GoTo ErrSection:

    FilterRulesGrid
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.optShort.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optShortExit_Click
'' Description: Show all of the Short Exit rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optShortExit_Click()
On Error GoTo ErrSection:

    FilterRulesGrid
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.optShortExit.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Preview
'' Description: Update the preview for the current selection
'' Inputs:      Grid that is currently active
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Preview(Grid As VSFlexGrid, lIndex As eAddFormTabs)
On Error GoTo ErrSection:
    
    Dim Rule As New cRule               ' Temporary Rule to get RTF Coded Text
    Dim lSecurityCol As Long            ' Column that has the security information
    Dim strPreview As String            ' String to display in the Preview box
    
    With Grid
        txtPreview(lIndex).Text = ""
    
        If .RowSel >= .FixedRows Then
            Select Case lIndex
            Case Tabs(eTab_Criteria)
                If Rule Is Nothing Then Set Rule = New cRule
                txtPreview(Tabs(eTab_Criteria)).TextRTF = Rule.GetRTF(.TextMatrix(.RowSel, CGCol(eCGC_CodedText)))
            
            Case eTab_Systems
                lSecurityCol = SGCol(eSGC_SecurityLevel)
                strPreview = .TextMatrix(.RowSel, SGCol(eSGC_Preview))
                If strPreview = "N/A" Then
                    strPreview = ""
                End If
                With txtPreview(lIndex)
                    .Text = strPreview
                    If Len(.Text) > 0 Then
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .SelColor = vbBlack
                        .SelBold = False
                        
                        .SelStart = 0
                        .SelLength = InStr(1, .Text, ": ")
                        If .SelLength > 0 Then
                            .SelBold = True
                        Else
                            .SelBold = False
                        End If
                        .SelLength = 0
                    End If
                End With
                
            Case eTab_Rules
                If Rule Is Nothing Then Set Rule = New cRule
                lSecurityCol = RGCol(eRGC_SecurityLevel)
                strPreview = Rule.GetRTF(.TextMatrix(.RowSel, RGCol(eRGC_Preview)))
                If g.Security.CanPreview(.Cell(flexcpValue, .RowSel, lSecurityCol)) Then
                    txtPreview(lIndex).TextRTF = strPreview
                Else
                    txtPreview(lIndex).SelColor = vbBlack
                    txtPreview(lIndex).Text = "Not authorized to view"
                End If
                
            Case eTab_Functions
                lSecurityCol = FGCol(eFGC_SecurityLevel)
                strPreview = .TextMatrix(.RowSel, FGCol(eFGC_Preview))
                With txtPreview(lIndex)
                    .Text = strPreview
                    If Len(.Text) > 0 And InStr(.Text, ":") > 0 Then
                        .SelLength = Len(.Text)
                        .SelColor = vbBlack
                        .SelStart = InStr(.Text, ": ") + 1
                        .SelLength = InStr(.Text, Chr(13)) - InStr(.Text, ": ") - 2
                        If .SelLength > 0 Then
                            .SelBold = True
                        Else
                            .SelBold = False
                        End If
                        .SelLength = 0
                    End If
                End With
                
            Case eTab_Libraries
                lSecurityCol = LGCol(eLGC_SecurityLevel)
                strPreview = .TextMatrix(.RowSel, LGCol(eLGC_Preview))
                With txtPreview(lIndex)
                    .Text = strPreview
                    If Len(.Text) > 0 Then
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .SelColor = vbBlack
                        .SelLength = 0
                    End If
                End With
            End Select
        End If
    End With
    
ErrExit:
    Set Rule = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.Preview", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitRulesGrid
'' Description: Initialize the rules grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitRulesGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    
    With fgRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .ExplorerBar = flexExSortShow 'AndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AutoSearch = flexSearchFromTop
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .ScrollTips = False
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = RGCol(eRGC_NumCols)
        .FrozenCols = 1
        
        'Column headings
        .TextMatrix(0, RGCol(eRGC_RuleName)) = "Rule"
        .Cell(flexcpFontBold, 0, RGCol(eRGC_RuleName)) = True
        .TextMatrix(0, RGCol(eRGC_RuleType)) = "Rule Type"
        .TextMatrix(0, RGCol(eRGC_SystemName)) = "Category"
        .TextMatrix(0, RGCol(eRGC_LibraryName)) = "Library"
        .TextMatrix(0, RGCol(eRGC_CategoryID)) = "Category ID"
        .TextMatrix(0, RGCol(eRGC_LastModified)) = "Last Modified"
        .TextMatrix(0, RGCol(eRGC_Reverify)) = "Verified"
        .TextMatrix(0, RGCol(eRGC_Preview)) = "Preview"
        .TextMatrix(0, RGCol(eRGC_RuleID)) = "Rule ID"
        .TextMatrix(0, RGCol(eRGC_SecurityLevel)) = "Security Level"
        .TextMatrix(0, RGCol(eRGC_Password)) = "Password"
        .TextMatrix(0, RGCol(eRGC_CannotDelete)) = "Cannot Delete"
        .TextMatrix(0, RGCol(eRGC_SystemNumber)) = "Strategy Number"
        .TextMatrix(0, RGCol(eRGC_TreeSortKey)) = "Tree Sort Key"
        .TextMatrix(0, RGCol(eRGC_TreeLevel)) = "Tree Level"
        
        .ColHidden(RGCol(eRGC_CategoryID)) = True
        .ColHidden(RGCol(eRGC_Preview)) = True
        .ColHidden(RGCol(eRGC_RuleID)) = True
        .ColHidden(RGCol(eRGC_SecurityLevel)) = True
        .ColHidden(RGCol(eRGC_Password)) = True
        .ColHidden(RGCol(eRGC_CannotDelete)) = True
        .ColHidden(RGCol(eRGC_SystemNumber)) = True
        .ColHidden(RGCol(eRGC_TreeSortKey)) = True
        .ColHidden(RGCol(eRGC_TreeLevel)) = True
        
        .ColDataType(RGCol(eRGC_CannotDelete)) = flexDTBoolean
        .ColDataType(RGCol(eRGC_Reverify)) = flexDTBoolean
        
        .ColAlignment(RGCol(eRGC_LastModified)) = flexAlignLeftCenter
        .ColFormat(RGCol(eRGC_LastModified)) = DateAndTime("Format")
        
        SetUpColumns eTab_Rules
        ExtendCustomColumn fgRules
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.InitRulesGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadRulesGrid
'' Description: Load the rules grid from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadRulesGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    Dim lScreenPointer As Long          ' Current state of the screen pointer
    Dim lRow As Long
    Dim lIndex As Long
    Dim aSystems As New cGdArray
    Dim strSystem As String
    Dim hRules As Long
    Dim hLibraries As Long
    Dim lLibraryIndex As Long
    Dim lSystemID As Long
    Dim lCategoryID As Long
    Dim astrCategories As New cGdArray
    Dim lPos As Long
    Dim rs As Recordset
    Dim strDesc As String
    Dim bSkip As Boolean
    
    lScreenPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    astrCategories.Create eGDARRAY_Strings
    
    With fgRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows + g.tblRule.NumRecords + 2
            
        ' Save off the table handles for the low level calls...
        hRules = g.tblRule.TableHandle
        hLibraries = g.tblLibrary.TableHandle
        
        lRow = .FixedRows
        .TextMatrix(lRow, RGCol(eRGC_RuleName)) = "BUILDING BLOCK RULES"
        .TextMatrix(lRow, RGCol(eRGC_RuleID)) = "0"
        .TextMatrix(lRow, RGCol(eRGC_TreeLevel)) = "0"
        .TextMatrix(lRow, RGCol(eRGC_TreeSortKey)) = Pad("a", 256, "L") & "0" & Pad("STANDARD RULES", 50, "L")
        .Cell(flexcpFontBold, lRow, RGCol(eRGC_RuleName)) = True
        lRow = lRow + 1
        .TextMatrix(lRow, RGCol(eRGC_RuleName)) = "RULES USED IN STRATEGIES"
        .TextMatrix(lRow, RGCol(eRGC_RuleID)) = "0"
        .TextMatrix(lRow, RGCol(eRGC_TreeLevel)) = "0"
        .TextMatrix(lRow, RGCol(eRGC_TreeSortKey)) = Pad("z", 256, "L") & "0" & Pad("RULES USED IN STRATEGIES", 50, "L")
        .Cell(flexcpFontBold, lRow, RGCol(eRGC_RuleName)) = True
        lRow = lRow + 1
            
        For lIndex = 0 To g.tblRule.NumRecords - 1
            bSkip = False
            If Not g.tblLibrary.FieldArray(LibraryField(etblLib_ID), False).BinarySearch(gdGetTableNum(hRules, RuleField(etblRule_LibraryID), lIndex), lLibraryIndex) Then
                bSkip = True
            ElseIf gdGetTableNum(hRules, RuleField(etblRule_SecurityLevel), lIndex) = 3 Then
                bSkip = True ' hide any hidden rules (e.g. Highlight Bar Reporter strategy)
            End If
            
            If Not bSkip Then
                .TextMatrix(lRow, RGCol(eRGC_RuleName)) = gdGetTableString(hRules, RuleField(etblRule_RuleName), lIndex)
                .TextMatrix(lRow, RGCol(eRGC_LibraryName)) = gdGetTableString(hLibraries, LibraryField(etblLib_Name), lLibraryIndex)
                .TextMatrix(lRow, RGCol(eRGC_RuleType)) = GetRuleType(hRules, lIndex)
                .TextMatrix(lRow, RGCol(eRGC_LastModified)) = CStr(gdGetTableNum(hRules, RuleField(etblRule_LastModified), lIndex))
                .TextMatrix(lRow, RGCol(eRGC_Preview)) = gdGetTableString(hRules, RuleField(etblRule_PreviewRTF), lIndex)
                .TextMatrix(lRow, RGCol(eRGC_RuleID)) = gdGetTableNum(hRules, RuleField(etblRule_RuleID), lIndex)
                .TextMatrix(lRow, RGCol(eRGC_SecurityLevel)) = gdGetTableNum(hRules, RuleField(etblRule_SecurityLevel), lIndex)
                .TextMatrix(lRow, RGCol(eRGC_Password)) = gdGetTableString(hRules, RuleField(etblRule_Password), lIndex)
                CheckedCell(fgRules, lRow, RGCol(eRGC_CannotDelete)) = CBool(gdGetTableString(hRules, RuleField(etblRule_CannotDelete), lIndex))
                .TextMatrix(lRow, RGCol(eRGC_SystemNumber)) = gdGetTableNum(hRules, RuleField(etblRule_SystemNumber), lIndex)
                CheckedCell(fgRules, lRow, RGCol(eRGC_Reverify)) = Not CBool(gdGetTableNum(hRules, RuleField(etblRule_Reverify), lIndex))

                lCategoryID = gdGetTableNum(hRules, RuleField(etblRule_CategoryID), lIndex)
                .TextMatrix(lRow, RGCol(eRGC_CategoryID)) = Str(lCategoryID)
                
                lSystemID = gdGetTableNum(hRules, RuleField(etblRule_SystemNumber), lIndex)
                If lSystemID = 0 Then
                    If lCategoryID > 0 Then
                        .TextMatrix(lRow, RGCol(eRGC_SystemName)) = RuleCategoryFromID(lCategoryID)
                    Else
                        .TextMatrix(lRow, RGCol(eRGC_SystemName)) = ""
                    End If
                Else
                    strSystem = aSystems(lSystemID)
                    If Len(strSystem) = 0 Then
                        strSystem = SystemNameForID(lSystemID)
                        aSystems(lSystemID) = strSystem
                    End If
                    .TextMatrix(lRow, RGCol(eRGC_SystemName)) = strSystem
                    If Len(strSystem) = 0 Then
                        bSkip = True ' this rule must be from an invalid system
                    End If
                End If
                
                .TextMatrix(lRow, RGCol(eRGC_TreeLevel)) = "2"
                If lCategoryID = 0 Then
                    .TextMatrix(lRow, RGCol(eRGC_TreeSortKey)) = Pad("z" & .TextMatrix(lRow, RGCol(eRGC_SystemName)), 256, "L") & .TextMatrix(lRow, RGCol(eRGC_TreeLevel)) & Pad(.TextMatrix(lRow, RGCol(eRGC_RuleName)), 50, "L")
                    
                    If Not astrCategories.BinarySearch("z" & .TextMatrix(lRow, RGCol(eRGC_SystemName)), lPos) Then
                        astrCategories.Add "z" & .TextMatrix(lRow, RGCol(eRGC_SystemName)), lPos
                    End If
                Else
                    .TextMatrix(lRow, RGCol(eRGC_TreeSortKey)) = Pad("a" & .TextMatrix(lRow, RGCol(eRGC_SystemName)), 256, "L") & .TextMatrix(lRow, RGCol(eRGC_TreeLevel)) & Pad(.TextMatrix(lRow, RGCol(eRGC_RuleName)), 50, "L")
                    
                    If Not astrCategories.BinarySearch("a" & .TextMatrix(lRow, RGCol(eRGC_SystemName)), lPos) Then
                        astrCategories.Add "a" & .TextMatrix(lRow, RGCol(eRGC_SystemName)), lPos
                    End If
                End If

                If gdGetTableNum(hRules, RuleField(etblRule_Reverify), lIndex) <> 0 Then
                    .Cell(flexcpForeColor, lRow, RGCol(eRGC_RuleName)) = vbRed
                Else
                    .Cell(flexcpForeColor, lRow, RGCol(eRGC_RuleName)) = .Cell(flexcpForeColor, 0, 0)
                End If
                
                If Not bSkip Then
                    ' Only AutoSize after the first ten rows...
                    'If lRow = 10 Then .AutoSize 0, .Cols - 1, False, 75
                    If lRow = 10 Then SetUpColumns eTab_Rules
                    
                    lRow = lRow + 1
                End If
            End If
        Next lIndex
        
        .Rows = lRow + astrCategories.Size
        
        For lIndex = 0 To astrCategories.Size - 1
            .TextMatrix(lRow, RGCol(eRGC_RuleName)) = Mid(astrCategories(lIndex), 2)
            .TextMatrix(lRow, RGCol(eRGC_RuleID)) = "0"
            .TextMatrix(lRow, RGCol(eRGC_TreeLevel)) = "1"
            .TextMatrix(lRow, RGCol(eRGC_TreeSortKey)) = Pad(astrCategories(lIndex), 256, "L") & "1" & Pad(Mid(astrCategories(lIndex), 2), 50, "L")
            .Cell(flexcpFontBold, lRow, RGCol(eRGC_RuleName)) = True
            lRow = lRow + 1
        Next lIndex
        
        ' Sort the Rule Name ascending
        .Col = RGCol(eRGC_RuleName)
        .Sort = flexSortGenericAscending
        DoPrevSort fgRules
        
        'fgRules.ColHidden(RGCol(eRGC_SystemName)) = (chkFavorites = vbChecked)
        cboLibrary.Enabled = (chkLibrary = vbChecked)
        'FilterRulesGrid
        
        ChangeRulesView
        
        ' Select a row
        lRow = 0
        If Len(m.strInitialSelect) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If Trim(UCase(.TextMatrix(lIndex, RGCol(eRGC_RuleID)))) = m.strInitialSelect Then
                    lRow = lIndex
                    Exit For
                End If
            Next
        End If

        If lRow > 0 Then
            .Row = lRow
            .RowSel = lRow
            .ShowCell lRow, RGCol(eRGC_RuleName)
        ElseIf .Rows > .FixedRows Then
            ' Select the first row (if there are any)
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
               
        .Redraw = lRedraw
    End With

ErrExit:
    Screen.MousePointer = lScreenPointer
    Set aSystems = Nothing
    Exit Sub
    
ErrSection:
    Screen.MousePointer = lScreenPointer
    RaiseError "frmToolbox.LoadRulesGrid", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitSystemsGrid
'' Description: Initialize the Systems grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitSystemsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    
    With fgSystems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .ExplorerBar = flexExSortShow 'AndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AutoSearch = flexSearchFromTop
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .ScrollTips = False
        .FixedCols = 0
        .FixedRows = 1
        .FrozenCols = 1
        .Rows = 1
        .Cols = kSystemGridCols
        
        'Column headings
        .TextMatrix(0, SGCol(eSGC_SystemDesc)) = "Strategy"
        .Cell(flexcpFontBold, 0, SGCol(eSGC_SystemDesc)) = True
        .TextMatrix(0, SGCol(eSGC_LibraryName)) = "Library"
        .TextMatrix(0, SGCol(eSGC_Developer)) = "Developer"
        .TextMatrix(0, SGCol(eSGC_LastModified)) = "Last Modified"
        .TextMatrix(0, SGCol(eSGC_UnVerified)) = "Verified"
        .TextMatrix(0, SGCol(eSGC_Preview)) = "Preview"
        .TextMatrix(0, SGCol(eSGC_SystemNumber)) = "Strategy Number"
        .TextMatrix(0, SGCol(eSGC_TradesPath)) = "Trades Path"
        .TextMatrix(0, SGCol(eSGC_SecurityLevel)) = "Security Level"
        .TextMatrix(0, SGCol(eSGC_Password)) = "Password"
        .TextMatrix(0, SGCol(eSGC_CannotDelete)) = "Cannot Delete"
        
        .ColHidden(SGCol(eSGC_Preview)) = True
        .ColHidden(SGCol(eSGC_SystemNumber)) = True
        .ColHidden(SGCol(eSGC_TradesPath)) = True
        .ColHidden(SGCol(eSGC_SecurityLevel)) = True
        .ColHidden(SGCol(eSGC_Password)) = True
        .ColHidden(SGCol(eSGC_CannotDelete)) = True
        
        .ColDataType(SGCol(eSGC_CannotDelete)) = flexDTBoolean
        ''.ColDataType(SGCol(eSGC_LastModified)) = flexDTDate
        .ColDataType(SGCol(eSGC_UnVerified)) = flexDTBoolean
        
        .ColAlignment(SGCol(eSGC_LastModified)) = flexAlignLeftCenter
        .ColFormat(SGCol(eSGC_LastModified)) = DateAndTime("Format")
        
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_Systems
        ExtendCustomColumn fgSystems
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.InitSystemsGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSystemsGrid
'' Description: Load the systems grid from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSystemsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lRow As Long
    Dim lIndex As Long
    Dim rsSystems As Recordset
    Dim lScreenPointer As Long
    
    lScreenPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    With fgSystems
        lRedraw = .Redraw
        .Redraw = flexRDNone
    
        .Rows = .FixedRows
        
        Set rsSystems = mSysNav.LoadStrategiesRecordset
        If Not (rsSystems.BOF And rsSystems.EOF) Then
            rsSystems.MoveFirst
            Do While Not rsSystems.EOF
                'If rsSystems!Ignore = 0 And rsSystems![tblSystems.CheckSum] <> 0.5 And rsSystems![tblLibrarys.CheckSum] <> 0.5 _
                        And (rsSystems![tblSystems.SecurityLevel] <> 3 Or IsIDE) Then
                If mSysNav.IncludeStrategiesFromRecordset(rsSystems, True) Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, SGCol(eSGC_SystemDesc)) = rsSystems!SystemName
                    .TextMatrix(.Rows - 1, SGCol(eSGC_LibraryName)) = rsSystems!LibraryName
                    .TextMatrix(.Rows - 1, SGCol(eSGC_Developer)) = rsSystems!Developer
                    .TextMatrix(.Rows - 1, SGCol(eSGC_LastModified)) = CStr(CDbl(rsSystems![tblSystems.LastModified]))
                    CheckedCell(fgSystems, .Rows - 1, SGCol(eSGC_UnVerified)) = Not rsSystems!Reverify
                    .TextMatrix(.Rows - 1, SGCol(eSGC_Preview)) = "Strategy Notes:  " & rsSystems!Notes
                    .TextMatrix(.Rows - 1, SGCol(eSGC_SystemNumber)) = rsSystems!SystemNumber
                    .TextMatrix(.Rows - 1, SGCol(eSGC_TradesPath)) = App.Path & "\TRADES\S" & rsSystems!SystemNumber & ".TXT"
                    .TextMatrix(.Rows - 1, SGCol(eSGC_SecurityLevel)) = rsSystems![tblSystems.SecurityLevel]
                    .TextMatrix(.Rows - 1, SGCol(eSGC_Password)) = DecryptField(rsSystems![tblSystems.Password])
                    CheckedCell(fgSystems, .Rows - 1, SGCol(eSGC_CannotDelete)) = rsSystems![tblSystems.CannotDelete]
                    
                    ' No longer want to color systems that need to be reverified in red
                    ' DAJ: 06/20/2003
                    'If rsSystems!Reverify Then
                    '    .Cell(flexcpForeColor, .Rows - 1, SGCol(eSGC_SystemDesc)) = vbRed
                    'Else
                        .Cell(flexcpForeColor, .Rows - 1, SGCol(eSGC_SystemDesc)) = .Cell(flexcpForeColor, 0, 0)
                    'End If
                End If
            
                rsSystems.MoveNext
            Loop
        End If
        
        ' Auto Size the columns
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_Systems
        ExtendCustomColumn fgSystems
        
        ' Sort the System Name ascending
        .Col = SGCol(eSGC_SystemDesc)
        .Sort = flexSortGenericAscending
        DoPrevSort fgSystems
        
        ' Select a row
        lRow = 0
        If Len(m.strInitialSelect) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If Trim(UCase(.TextMatrix(lIndex, SGCol(eSGC_SystemNumber)))) = m.strInitialSelect Then
                    lRow = lIndex
                    Exit For
                End If
            Next
        End If
        If lRow > 0 Then
            .Row = lRow
            .RowSel = lRow
            .ShowCell lRow, SGCol(eSGC_SystemDesc)
        ElseIf .Rows > .FixedRows Then
            ' Select the first row (if there are any)
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Screen.MousePointer = lScreenPointer
    Set rsSystems = Nothing
    Exit Sub
    
ErrSection:
    Screen.MousePointer = lScreenPointer
    RaiseError "frmToolbox.LoadSystemsGrid", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitSymbolGroupGrid
'' Description: Initialize the Symbol Group grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitSymbolGroupGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    
    With fgGroups
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .ExplorerBar = flexExSortShow 'AndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AutoSearch = flexSearchFromTop
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .ScrollTips = False
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = kGroupsGridCols
        
        .TextMatrix(0, GGCol(eGGC_Name)) = "Symbol Group"
        .Cell(flexcpFontBold, 0, GGCol(eGGC_Name)) = True
        .TextMatrix(0, GGCol(eGGC_Description)) = "Description"
        .TextMatrix(0, GGCol(eGGC_GroupID)) = "Group ID"
        
        .ColHidden(GGCol(eGGC_GroupID)) = True
        
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_SymbolGroups
        ExtendCustomColumn fgGroups
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.InitSymbolGroupGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitCriteriaGrid
'' Description: Initialize the Criteria grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitCriteriaGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    
    With fgCriteria
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow 'AndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AutoSearch = flexSearchFromTop
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .ScrollTips = False
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = kCriteriaGridCols
        
        .TextMatrix(0, CGCol(eCGC_Active)) = "Active"
        .TextMatrix(0, CGCol(eCGC_Name)) = "Criteria"
        .Cell(flexcpFontBold, 0, CGCol(eCGC_Name)) = True
        .TextMatrix(0, CGCol(eCGC_Description)) = "Description"
        .TextMatrix(0, CGCol(eCGC_NumDays)) = "Num Days"
        .TextMatrix(0, CGCol(eCGC_CodedText)) = "Coded Text"
        .TextMatrix(0, CGCol(eCGC_CriteriaID)) = "CriteriaID"
        
        .ColHidden(CGCol(eCGC_CodedText)) = True
        .ColHidden(CGCol(eCGC_CriteriaID)) = True
        
        .ColDataType(CGCol(eCGC_Active)) = flexDTBoolean
        
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_Criteria
        ExtendCustomColumn fgCriteria
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.InitCriteriaGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitFiltersGrid
'' Description: Initialize the Filters grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFiltersGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    
    With fgFilters
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow 'AndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AutoSearch = flexSearchFromTop
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .ScrollTips = False
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = kFiltersGridCols
        
        .TextMatrix(0, FiCol(eFIC_Active)) = "Active"
        .TextMatrix(0, FiCol(eFIC_Name)) = "Filter"
        .Cell(flexcpFontBold, 0, FiCol(eFIC_Name)) = True
        .TextMatrix(0, FiCol(eFIC_Description)) = "Description"
        .TextMatrix(0, FiCol(eFIC_FilterID)) = "FilterID"
        
        .ColHidden(FiCol(eFIC_FilterID)) = True
        
        .ColDataType(FiCol(eFIC_Active)) = flexDTBoolean
        
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_Filters
        ExtendCustomColumn fgFilters
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.InitFiltersGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitStrategyBasketsGrid
'' Description: Initialize the Filters grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitStrategyBasketsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    
    With fgStrategyBaskets
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow 'AndMove
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AutoSearch = flexSearchFromTop
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .ScrollTips = False
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = kStrategyBasketGridCols
        
        .TextMatrix(0, SBCol(eSBC_Name)) = "Basket Name"
        .Cell(flexcpFontBold, 0, SBCol(eSBC_Name)) = True
        .TextMatrix(0, SBCol(eSBC_LastModified)) = "Last Modified"
        .TextMatrix(0, SBCol(eSBC_Description)) = "Description"
        
        .ColFormat(SBCol(eSBC_LastModified)) = DateFormat("Format", MM_DD_YYYY, HH_MM, AMPM_UPPER)
        
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_StrategyBaskets
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.InitStrategyBasketsGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitFunctionsGrid
'' Description: Initialize the Functions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFunctionsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    
    With fgFunctions
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .ExplorerBar = flexExSortShow 'AndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AutoSearch = flexSearchFromTop
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .ScrollTips = False
        .Editable = flexEDKbdMouse
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = FGCol(eFGC_NumCols)
        
        'Column headings
        .TextMatrix(0, FGCol(eFGC_FunctionName)) = "Function"
        .Cell(flexcpFontBold, 0, FGCol(eFGC_FunctionName)) = True
        .TextMatrix(0, FGCol(eFGC_LibraryName)) = "Library"
        .TextMatrix(0, FGCol(eFGC_Category)) = "Category"
        .TextMatrix(0, FGCol(eFGC_LastModified)) = "Last Modified"
        .TextMatrix(0, FGCol(eFGC_Reverify)) = "Verified"
        .TextMatrix(0, FGCol(eFGC_ImplType)) = "Implementation Type"
        .TextMatrix(0, FGCol(eFGC_Preview)) = "Preview"
        .TextMatrix(0, FGCol(eFGC_FunctionID)) = "Function ID"
        .TextMatrix(0, FGCol(eFGC_SecurityLevel)) = "Security Level"
        .TextMatrix(0, FGCol(eFGC_Usage)) = "Usage"
        .TextMatrix(0, FGCol(eFGC_Password)) = "Password"
        .TextMatrix(0, FGCol(eFGC_CannotDelete)) = "Cannot Delete"
        .TextMatrix(0, FGCol(eFGC_Favorites)) = "Favorites"
        
        .ColHidden(FGCol(eFGC_ImplType)) = True
        .ColHidden(FGCol(eFGC_Preview)) = True
        .ColHidden(FGCol(eFGC_FunctionID)) = True
        .ColHidden(FGCol(eFGC_SecurityLevel)) = True
        .ColHidden(FGCol(eFGC_Usage)) = True
        .ColHidden(FGCol(eFGC_Password)) = True
        .ColHidden(FGCol(eFGC_CannotDelete)) = True
        .ColHidden(FGCol(eFGC_TreeSortKey)) = True
        .ColHidden(FGCol(eFGC_TreeLevel)) = True
        .ColHidden(FGCol(eFGC_CodedName)) = True
        
        .ColDataType(FGCol(eFGC_CannotDelete)) = flexDTBoolean
        .ColDataType(FGCol(eFGC_Reverify)) = flexDTBoolean
        .ColDataType(FGCol(eFGC_Favorites)) = flexDTBoolean
        
        .ColAlignment(FGCol(eFGC_LastModified)) = flexAlignLeftCenter
        .ColFormat(FGCol(eFGC_LastModified)) = DateAndTime("Format")
        
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_Functions
        ExtendCustomColumn fgFunctions
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.InitFunctionsGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSymbolGroupGrid
'' Description: Load the Symbol Group Grid from the Symbol Pool
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSymbolGroupGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lRow As Long
    Dim lScreenPointer As Long
    Dim SymbolGroups As Object
    Dim lIndex As Long
    Dim SymbolGroup As cSymbolGroup
    
    lScreenPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    With fgGroups
        lRedraw = .Redraw
        .Redraw = flexRDNone
    
        .Rows = .FixedRows
        Set SymbolGroups = g.SymbolPool.SymbolGroups
        For lIndex = 1 To SymbolGroups.Count
            Set SymbolGroup = SymbolGroups(lIndex)
            
            ' Have seen one case on a customer's machine where the symbol group ID must
            ' be blank -- put a check here not to load it in the grid if it is (01/13/2009 DAJ)...
            If Len(SymbolGroup.ID) > 0 Then
                If SymbolGroup.Custom And SymbolGroup.GroupType = eGROUP_Normal And HasModule(SymbolGroup.Required) Then
                    .Rows = .Rows + 1
                    lRow = .Rows - 1
                    .TextMatrix(lRow, GGCol(eGGC_GroupID)) = SymbolGroup.ID
                    .TextMatrix(lRow, GGCol(eGGC_Name)) = SymbolGroup.Name
                    .TextMatrix(lRow, GGCol(eGGC_Description)) = SymbolGroup.Desc
                End If
            End If
        Next lIndex

        ' Auto Size the columns
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_SymbolGroups
        ExtendCustomColumn fgGroups
        
        ' Sort the System Name ascending
        .Col = GGCol(eGGC_Name)
        .Sort = flexSortGenericAscending
        DoPrevSort fgGroups
        
        ' Select a row
        lRow = 0
        If Len(m.strInitialSelect) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If Trim(UCase(.TextMatrix(lIndex, GGCol(eGGC_GroupID)))) = m.strInitialSelect Then
                    lRow = lIndex
                    Exit For
                End If
            Next
        End If
        If lRow > 0 Then
            .Row = lRow
            .RowSel = lRow
            .ShowCell lRow, GGCol(eGGC_Name)
        ElseIf .Rows > .FixedRows Then
            ' Select the first row (if there are any)
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Screen.MousePointer = lScreenPointer
    Set SymbolGroup = Nothing
    Set SymbolGroups = Nothing
    Exit Sub
    
ErrSection:
    Screen.MousePointer = lScreenPointer
    RaiseError "frmToolbox.LoadSymbolGroupGrid", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFiltersGrid
'' Description: Load the Filters Grid from the Symbol Pool
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFiltersGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lRow As Long
    Dim lScreenPointer As Long
    Dim Filters As Object
    Dim Filter As cFilter
    Dim lIndex As Long
    
    lScreenPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    With fgFilters
        lRedraw = .Redraw
        .Redraw = flexRDNone
    
        .Rows = .FixedRows
        
        Set Filters = g.SymbolPool.Filters
        For lIndex = 1 To Filters.Count
            Set Filter = Filters(lIndex)
            If Filter.Custom Then
                If HasModule(Filter.Required) Then
                    .Rows = .Rows + 1
                    lRow = .Rows - 1
                    .TextMatrix(lRow, FiCol(eFIC_FilterID)) = Filter.ID
                    CheckedCell(fgFilters, lRow, FiCol(eFIC_Active)) = Filter.IsActive
                    .TextMatrix(lRow, FiCol(eFIC_Name)) = Filter.Name
                    .TextMatrix(lRow, FiCol(eFIC_Description)) = Filter.Desc
                End If
            End If
        Next lIndex

        ' Auto Size the columns
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_Filters
        ExtendCustomColumn fgFilters
        
        ' Sort the System Name ascending
        .Col = FiCol(eFIC_Name)
        .Sort = flexSortGenericAscending
        DoPrevSort fgFilters
        
        ' Select a row
        lRow = 0
        If Len(m.strInitialSelect) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If Trim(UCase(.TextMatrix(lIndex, FiCol(eFIC_FilterID)))) = m.strInitialSelect Then
                    lRow = lIndex
                    Exit For
                End If
            Next
        End If
        If lRow > 0 Then
            .Row = lRow
            .RowSel = lRow
            .ShowCell lRow, FiCol(eFIC_Name)
        ElseIf .Rows > .FixedRows Then
            ' Select the first row (if there are any)
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Screen.MousePointer = lScreenPointer
    Set Filter = Nothing
    Set Filters = Nothing
    Exit Sub
    
ErrSection:
    Screen.MousePointer = lScreenPointer
    RaiseError "frmToolbox.LoadFiltersGrid", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCriteriaGrid
'' Description: Load the Criteria Grid from the Symbol Pool
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCriteriaGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lRow As Long
    Dim lScreenPointer As Long
    Dim Criterias As Object
    Dim Criteria As cCriteria
    Dim lIndex As Long
    Dim Expr As New cExpression
    
    lScreenPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    With fgCriteria
        lRedraw = .Redraw
        .Redraw = flexRDNone
    
        .Rows = .FixedRows
        
        Set Criterias = g.SymbolPool.Criterias
        For lIndex = 1 To Criterias.Count
            Set Criteria = Criterias(lIndex)
            If Criteria.Custom Then
                If HasModule(Criteria.Required) Then
                    .Rows = .Rows + 1
                    lRow = .Rows - 1
                    CheckedCell(fgCriteria, lRow, CGCol(eCGC_Active)) = Criteria.IsActive
                    .TextMatrix(lRow, CGCol(eCGC_CriteriaID)) = Criteria.ID
                    .TextMatrix(lRow, CGCol(eCGC_Name)) = Criteria.Name
                    .TextMatrix(lRow, CGCol(eCGC_Description)) = Criteria.Desc
                    If Criteria.IsWeekly Then
                        .TextMatrix(lRow, CGCol(eCGC_NumDays)) = Trim(CStr(Criteria.NumDays * 5))
                    Else
                        .TextMatrix(lRow, CGCol(eCGC_NumDays)) = Trim(CStr(Criteria.NumDays))
                    End If
                    
                    If Criteria.FormattedText <> "" Then
                        .TextMatrix(lRow, CGCol(eCGC_CodedText)) = Criteria.FormattedText
                    Else
                        .TextMatrix(lRow, CGCol(eCGC_CodedText)) = Criteria.CodedText
                    End If
                End If
            End If
        Next lIndex

        ' Auto Size the columns
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_Criteria
        ExtendCustomColumn fgCriteria
        
        ' Sort the System Name ascending
        .Col = CGCol(eCGC_Name)
        .Sort = flexSortGenericAscending
        DoPrevSort fgCriteria
        
        ' Select a row
        lRow = 0
        If Len(m.strInitialSelect) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If Trim(UCase(.TextMatrix(lIndex, CGCol(eCGC_CriteriaID)))) = m.strInitialSelect Then
                    lRow = lIndex
                    Exit For
                End If
            Next
        End If
        If lRow > 0 Then
            .Row = lRow
            .RowSel = lRow
            .ShowCell lRow, CGCol(eCGC_Name)
        ElseIf .Rows > .FixedRows Then
            ' Select the first row (if there are any)
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Screen.MousePointer = lScreenPointer
    Set Criteria = Nothing
    Set Criterias = Nothing
    Exit Sub
    
ErrSection:
    Screen.MousePointer = lScreenPointer
    RaiseError "frmToolbox.LoadCriteriaGrid", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFunctionsGrid
'' Description: Load the Functions grid from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFunctionsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lRow As Long
    Dim lIndex As Long
    Dim lScreenPointer As Long
    Dim lLibraryIndex As Long
    Dim hFunctions As Long
    Dim hLibraries As Long
    Dim hFuncCat As Long
    Dim astrFavorites As New cGdArray
    Dim astrCategories As New cGdArray
    Dim lPos As Long
    
    lScreenPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    Set astrFavorites = GetFunctionFavorites
    astrCategories.Create eGDARRAY_Strings
    
    With fgFunctions
        lRedraw = .Redraw
        .Redraw = flexRDNone
   
        lRow = .FixedRows
        .Rows = .FixedRows + g.Functions.Count
        
        ' Store off the table handle to use with the low level calls...
        hFunctions = g.tblFunction.TableHandle
        hLibraries = g.tblLibrary.TableHandle
        hFuncCat = g.astrFunctionCategory.ArrayHandle
        
        For lIndex = 0 To g.tblFunction.NumRecords - 1
            If Not g.tblLibrary.FieldArray(LibraryField(etblLib_ID), False).BinarySearch(gdGetTableNum(hFunctions, FunctionField(etblFunction_LibID), lIndex), lLibraryIndex) Then
                lLibraryIndex = -1
            End If

            If lLibraryIndex <> -1 Then
                If (gdGetTableNum(hFunctions, FunctionField(etblFunction_Implementation), lIndex) <= 2 _
                        And gdGetTableNum(hFunctions, FunctionField(etblFunction_SecurityLevel), lIndex) <= 2) _
                            Or IsIDE Then
                    .TextMatrix(lRow, FGCol(eFGC_FunctionName)) = gdGetTableString(hFunctions, FunctionField(etblFunction_Name), lIndex)
                    .TextMatrix(lRow, FGCol(eFGC_CodedName)) = gdGetTableString(hFunctions, FunctionField(etblFunction_NameCoded), lIndex)
                    .TextMatrix(lRow, FGCol(eFGC_LibraryName)) = gdGetTableString(hLibraries, LibraryField(etblLib_Name), lLibraryIndex)
                    .TextMatrix(lRow, FGCol(eFGC_Category)) = gdGetStr(hFuncCat, gdGetTableNum(hFunctions, FunctionField(etblFunction_CategoryID), lIndex))
                    .TextMatrix(lRow, FGCol(eFGC_LastModified)) = CStr(gdGetTableNum(hFunctions, FunctionField(etblFunction_LastModified), lIndex))
                    .TextMatrix(lRow, FGCol(eFGC_Preview)) = "Usage: " & gdGetTableString(hFunctions, FunctionField(etblFunction_TradeSenseUsage), lIndex) & vbCrLf & _
                         "Description: " & gdGetTableString(hFunctions, FunctionField(etblFunction_Description), lIndex)
                    .TextMatrix(lRow, FGCol(eFGC_FunctionID)) = gdGetTableNum(hFunctions, FunctionField(etblFunction_ID), lIndex)
                    .TextMatrix(lRow, FGCol(eFGC_ImplType)) = gdGetTableNum(hFunctions, FunctionField(etblFunction_Implementation), lIndex)
                    .TextMatrix(lRow, FGCol(eFGC_Usage)) = gdGetTableNum(hFunctions, FunctionField(etblFunction_Usage), lIndex)
                    .TextMatrix(lRow, FGCol(eFGC_SecurityLevel)) = gdGetTableNum(hFunctions, FunctionField(etblFunction_SecurityLevel), lIndex)
                    .TextMatrix(lRow, FGCol(eFGC_Password)) = gdGetTableString(hFunctions, FunctionField(etblFunction_Password), lIndex)
                    CheckedCell(fgFunctions, lRow, FGCol(eFGC_CannotDelete)) = CBool(gdGetTableNum(hFunctions, FunctionField(etblFunction_CannotDelete), lIndex))
                    CheckedCell(fgFunctions, lRow, FGCol(eFGC_Reverify)) = Not CBool(gdGetTableNum(hFunctions, FunctionField(etblFunction_Reverify), lIndex))
                    
                    If Not CheckedCell(fgFunctions, lRow, FGCol(eFGC_Reverify)) Then
                        .Cell(flexcpForeColor, lRow, FGCol(eFGC_FunctionName)) = vbRed
                    Else
                        .Cell(flexcpForeColor, lRow, FGCol(eFGC_FunctionName)) = .Cell(flexcpForeColor, 0, 0)
                    End If
                    
                    If astrFavorites.BinarySearch(.TextMatrix(lRow, FGCol(eFGC_CodedName))) = True Then
                        CheckedCell(fgFunctions, lRow, FGCol(eFGC_Favorites)) = True
                    Else
                        CheckedCell(fgFunctions, lRow, FGCol(eFGC_Favorites)) = False
                    End If
                    
                    If astrCategories.BinarySearch(.TextMatrix(lRow, FGCol(eFGC_Category)), lPos) = False Then
                        astrCategories.Add .TextMatrix(lRow, FGCol(eFGC_Category)), lPos
                    End If
                    
                    .TextMatrix(lRow, FGCol(eFGC_TreeSortKey)) = Pad(.TextMatrix(lRow, FGCol(eFGC_Category)), 30, "L") & Pad(.TextMatrix(lRow, FGCol(eFGC_FunctionName)), 50, "L")
                    .TextMatrix(lRow, FGCol(eFGC_TreeLevel)) = "1"
                                        
                    'If lRow = 10 Then .AutoSize 0, .Cols - 1, False, 75
                    If lRow = 10 Then SetUpColumns eTab_Functions
                    
                    lRow = lRow + 1
                End If
            End If
        Next lIndex
        
        .Rows = lRow + astrCategories.Size
        
        For lIndex = 0 To astrCategories.Size - 1
            .TextMatrix(lRow, FGCol(eFGC_FunctionName)) = astrCategories(lIndex)
            .TextMatrix(lRow, FGCol(eFGC_TreeSortKey)) = Pad(astrCategories(lIndex), 30, "L") & Pad("", 50, "L")
            .TextMatrix(lRow, FGCol(eFGC_TreeLevel)) = "0"
            
            lRow = lRow + 1
        Next lIndex
        
        ' Auto Size the columns
        '.AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn fgFunctions
        
        ' Sort the Function Name ascending
        .Col = FGCol(eFGC_FunctionName)
        .Sort = flexSortGenericAscending
        DoPrevSort fgFunctions
        
        ChangeFunctionsView
        
        ' Select a row
        lRow = 0
        If Len(m.strInitialSelect) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If Trim(UCase(.TextMatrix(lIndex, FGCol(eFGC_FunctionID)))) = m.strInitialSelect Then
                    lRow = lIndex
                    Exit For
                End If
            Next
        End If

        If lRow > 0 Then
            .Row = lRow
            .RowSel = lRow
            .ShowCell lRow, FGCol(eFGC_FunctionName)
        ElseIf .Rows > .FixedRows Then
            ' Select the first row (if there are any)
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Screen.MousePointer = lScreenPointer
    Exit Sub
    
ErrSection:
    Screen.MousePointer = lScreenPointer
    RaiseError "frmToolbox.LoadFunctionsGrid", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitLibrariesGrid
'' Description: Initialize the Functions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitLibrariesGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current status of the grid redraw
    
    With fgLibraries
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .ExplorerBar = flexExSortShow 'AndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AutoSearch = flexSearchFromTop
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .ScrollTips = False
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = kLibraryGridCols
        
        'Column headings
        .TextMatrix(0, LGCol(eLGC_LibraryName)) = "Library"
        .Cell(flexcpFontBold, 0, LGCol(eLGC_LibraryName)) = True
        .TextMatrix(0, LGCol(eLGC_Author)) = "Author"
        .TextMatrix(0, LGCol(eLGC_LastModified)) = "Last Modified"
        .TextMatrix(0, LGCol(eLGC_Preview)) = "Preview"
        .TextMatrix(0, LGCol(eLGC_LibraryID)) = "Library ID"
        .TextMatrix(0, LGCol(eLGC_SecurityLevel)) = "Security Level"
        .TextMatrix(0, LGCol(eLGC_Password)) = "Password"
        .TextMatrix(0, LGCol(eLGC_CannotDelete)) = "Cannot Delete"
        
        .ColHidden(LGCol(eLGC_Preview)) = True
        .ColHidden(LGCol(eLGC_LibraryID)) = True
        .ColHidden(LGCol(eLGC_SecurityLevel)) = True
        .ColDataType(LGCol(eLGC_SecurityLevel)) = flexDTShort
        .ColHidden(LGCol(eLGC_Password)) = True
        .ColHidden(LGCol(eLGC_CannotDelete)) = True
        
        .ColAlignment(LGCol(eLGC_LastModified)) = flexAlignLeftCenter
        .ColFormat(LGCol(eLGC_LastModified)) = DateAndTime("Format")
        
        .ColDataType(LGCol(eLGC_CannotDelete)) = flexDTBoolean
        ''.ColDataType(LGCol(eLGC_LastModified)) = flexDTDate
        
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_Libraries
        ExtendCustomColumn fgLibraries
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.InitLibrariesGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadLibrariesGrid
'' Description: Load the Functions grid from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadLibrariesGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lRow As Long
    Dim lIndex As Long
    Dim rsLibraries As Recordset
    Dim lScreenPointer As Long
    
    lScreenPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    With fgLibraries
        lRedraw = .Redraw
        .Redraw = flexRDNone
    
        .Rows = .FixedRows
        Set rsLibraries = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
                "WHERE [SecurityLevel]<>3;", dbOpenDynaset)
        ValidateCheckSums rsLibraries, "tblLibrarys"
                
        If Not (rsLibraries.BOF And rsLibraries.EOF) Then
            rsLibraries.MoveFirst
            Do While Not rsLibraries.EOF
                If rsLibraries!Ignore = 0 And rsLibraries!CheckSum <> 0.5 Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, LGCol(eLGC_LibraryName)) = rsLibraries!LibraryName
                    .TextMatrix(.Rows - 1, LGCol(eLGC_Author)) = NullChk(rsLibraries!Author)
                    .TextMatrix(.Rows - 1, LGCol(eLGC_LastModified)) = CStr(CDbl(rsLibraries!LastModified))
                    .TextMatrix(.Rows - 1, LGCol(eLGC_Preview)) = NullChk(rsLibraries!LibraryDesc)
                    .TextMatrix(.Rows - 1, LGCol(eLGC_LibraryID)) = rsLibraries!LibraryID
                    .TextMatrix(.Rows - 1, LGCol(eLGC_SecurityLevel)) = rsLibraries!SecurityLevel
                    .TextMatrix(.Rows - 1, LGCol(eLGC_Password)) = DecryptField(rsLibraries!Password)
                    CheckedCell(fgLibraries, .Rows - 1, LGCol(eLGC_CannotDelete)) = rsLibraries!CannotDelete
                End If
            
                rsLibraries.MoveNext
            Loop
        End If
        
        ' Auto Size the columns
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_Libraries
        ExtendCustomColumn fgLibraries
        
        ' Sort the System Name ascending
        .Col = FGCol(eFGC_FunctionName)
        .Sort = flexSortGenericAscending
        DoPrevSort fgLibraries
        
        ' Select a row
        lRow = 0
        If Len(m.strInitialSelect) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If Trim(UCase(.TextMatrix(lIndex, LGCol(eLGC_LibraryID)))) = m.strInitialSelect Then
                    lRow = lIndex
                    Exit For
                End If
            Next
        End If

        If lRow > 0 Then
            .Row = lRow
            .RowSel = lRow
            .ShowCell lRow, LGCol(eLGC_LibraryName)
        ElseIf .Rows > .FixedRows Then
            ' Select the first row (if there are any)
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Screen.MousePointer = lScreenPointer
    Set rsLibraries = Nothing
    Exit Sub
    
ErrSection:
    Screen.MousePointer = lScreenPointer
    RaiseError "frmToolbox.LoadLibrariesGrid", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadStrategyBasketsGrid
'' Description: Load the System Runs Grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadStrategyBasketsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long
    Dim Baskets As cStrategyBaskets     ' Strategy basket collection
    Dim Basket As cStrategyBasket       ' Strategy basket item
    Dim bInclude As Boolean             ' Include the basket in the grid?
    
    With fgStrategyBaskets
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Clear the grid first
        .Rows = .FixedRows
        
        ' TLB 7/5/2013: it is MUCH faster to load the baskets without the basket items
        ' (and we don't need the items loaded for just showing all the baskets in the toolbox grid)
        Set Baskets = New cStrategyBaskets
        Baskets.LoadDb , , False
        
        For lIndex = 1 To Baskets.Count
            Set Basket = Baskets(lIndex)
            
            bInclude = False
            If (HasModule(Basket.RequiredModule & "*") = True) Then
                bInclude = True
            ElseIf (IsIDE = True) Then
                bInclude = True
            ElseIf (Basket.IsGuru = True) And (IsOwnerOfGuruObject(Basket.LibraryID) = True) Then
                bInclude = True
            End If
            
            If bInclude = True Then
                .Rows = .Rows + 1
                
                .RowData(.Rows - 1) = Basket
                .TextMatrix(.Rows - 1, SBCol(eSBC_Name)) = Basket.Name
                .TextMatrix(.Rows - 1, SBCol(eSBC_LastModified)) = Basket.LastModified
                .TextMatrix(.Rows - 1, SBCol(eSBC_Description)) = Basket.Description
            End If
        Next lIndex
                
        ' Sort the Name ascending
        .Col = SBCol(eSBC_Name)
        .Sort = flexSortGenericAscending
        DoPrevSort fgStrategyBaskets
                
        ' Select a row
        lRow = 0
        If Len(m.strInitialSelect) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If Trim(UCase(.TextMatrix(lIndex, SBCol(eSBC_Name)))) = m.strInitialSelect Then
                    lRow = lIndex
                    Exit For
                End If
            Next
        End If
        If lRow > 0 Then
            .Row = lRow
            .RowSel = lRow
            .ShowCell lRow, SBCol(eSBC_Name)
        ElseIf .Rows > .FixedRows Then
            ' Select the first row (if there are any)
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        '.AutoSize 0, .Cols - 1, False, 75
        SetUpColumns eTab_StrategyBaskets
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.LoadStrategyBasketsGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetButtons
'' Description: Show/Hide the appropriate buttons in the appropriate place
''              depending on the Mode and the Current Tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetButtons(ByVal lTab As Long)
On Error GoTo ErrSection:

    Static lDiff As Long

    If lDiff = 0 Then lDiff = cmdAddCopy.Top - cmdAdd.Top

    LockWindowUpdate Me.hWnd
    Select Case m.Mode
        Case eAddMode_Add
            cmdAdd.Visible = False
            cmdAddCopy.Visible = True
            cmdAddCopy.Top = cmdAdd.Top
            cmdCancel.Visible = True
            cmdCancel.Top = cmdAddCopy.Top + lDiff
            cmdNew.Visible = False 'True
            cmdNewDLL.Visible = False
            cmdEdit.Visible = False 'True
            cmdDelete.Visible = False 'True
            cmdImport.Visible = False
            cmdInfo.Visible = False
            cmdRename.Visible = False
            cmdExport.Visible = False
            
        Case eAddMode_Select
            cmdAdd.Visible = False
            cmdAddCopy.Visible = False
            cmdNew.Visible = True
            cmdEdit.Visible = True
            cmdDelete.Visible = True
            cmdNew.Top = cmdAdd.Top
            cmdEdit.Top = cmdAddCopy.Top
            cmdDelete.Top = cmdEdit.Top + lDiff
            If lTab = eTab_Libraries Then
                cmdImport.Visible = True
                cmdInfo.Visible = True
                cmdExport.Visible = True
                cmdRename.Visible = False
                cmdImport.Top = cmdDelete.Top + lDiff
                cmdExport.Top = cmdImport.Top + lDiff
                cmdInfo.Top = cmdExport.Top + lDiff
                cmdCancel.Top = cmdInfo.Top + lDiff
            Else
                cmdImport.Visible = False
                cmdInfo.Visible = False
                cmdCancel.Top = cmdDelete.Top + lDiff
                If lTab = Tabs(eTab_Functions) Then
                    cmdNewDLL.Visible = CBool(DirExist(AddSlash(App.Path) & "..\SDK"))
                    cmdRename.Visible = False
                    cmdNewDLL.Top = cmdNew.Top + lDiff
                    If cmdNewDLL.Visible Then
                        cmdEdit.Top = cmdNewDLL.Top + lDiff
                    Else
                        cmdEdit.Top = cmdNew.Top + lDiff
                    End If
                    cmdDelete.Top = cmdEdit.Top + lDiff
                    cmdExport.Visible = True
                    cmdExport.Top = cmdDelete.Top + lDiff
                    cmdCancel.Top = cmdExport.Top + lDiff
                ElseIf lTab = Tabs(eTab_Rules) Or lTab = Tabs(eTab_Systems) Then
                    cmdNewDLL.Visible = False
                    cmdRename.Visible = False
                    cmdExport.Visible = True
                    cmdEdit.Top = cmdNew.Top + lDiff
                    cmdDelete.Top = cmdEdit.Top + lDiff
                    cmdExport.Top = cmdDelete.Top + lDiff
                    cmdCancel.Top = cmdExport.Top + lDiff
                ElseIf lTab = Tabs(eTab_SymbolGroups) Or lTab = Tabs(eTab_Criteria) Or lTab = Tabs(eTab_Filters) Then
                    cmdRename.Visible = True
                    cmdExport.Visible = False
                    cmdRename.Top = cmdEdit.Top + lDiff
                    cmdDelete.Top = cmdRename.Top + lDiff
                    cmdCancel.Top = cmdDelete.Top + lDiff
                Else
                    cmdNewDLL.Visible = False
                    cmdRename.Visible = False
                    cmdExport.Visible = False
                    cmdEdit.Top = cmdNew.Top + lDiff
                    cmdDelete.Top = cmdEdit.Top + lDiff
                    cmdCancel.Top = cmdDelete.Top + lDiff
                End If
            End If
            
        Case eAddMode_List
            cmdAdd.Visible = False
            cmdAddCopy.Visible = False
            cmdCancel.Visible = True
            cmdCancel.Top = cmdAdd.Top
            cmdCancel.Caption = "&Close"
            cmdNew.Visible = False 'True
            cmdNewDLL.Value = False
            cmdEdit.Visible = False 'True
            cmdDelete.Visible = False 'True
            cmdImport.Visible = False
            cmdInfo.Visible = False
            cmdRename.Visible = False
            cmdExport.Visible = False
                                        
    End Select
    
    If lTab = eTab_Systems Or lTab = eTab_StrategyBaskets Then
        cmdEdit.Caption = "&Edit/Run"
    Else
        cmdEdit.Caption = "&Edit"
    End If
    
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    
    Select Case lTab
        Case Tabs(eTab_Functions)
            cmdEdit.Enabled = fgFunctions.TextMatrix(fgFunctions.RowSel, FGCol(eFGC_TreeLevel)) = "1"
            cmdDelete.Enabled = fgFunctions.TextMatrix(fgFunctions.RowSel, FGCol(eFGC_TreeLevel)) = "1"
    
        Case Tabs(eTab_Rules)
            cmdEdit.Enabled = fgRules.TextMatrix(fgRules.RowSel, RGCol(eRGC_TreeLevel)) = "2"
            cmdDelete.Enabled = fgRules.TextMatrix(fgRules.RowSel, RGCol(eRGC_TreeLevel)) = "2"
    
    End Select
    
    LockWindowUpdate 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.SetButtons", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsTypeTabs_Click
'' Description: When the user changes tabs, set the focus to the current grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsTypeTabs_Click()
On Error GoTo ErrSection:

    Select Case vsTypeTabs.CurrTab
        Case Tabs(eTab_SymbolGroups)
            MoveFocus fgGroups
        Case Tabs(eTab_Criteria)
            MoveFocus fgCriteria
        Case Tabs(eTab_Filters)
            MoveFocus fgFilters
        Case Tabs(eTab_Functions)
            MoveFocus fgFunctions
        Case Tabs(eTab_Rules)
            MoveFocus fgRules
        Case Tabs(eTab_Systems)
            MoveFocus fgSystems
        Case Tabs(eTab_Libraries)
            MoveFocus fgLibraries
        Case Tabs(eTab_StrategyBaskets)
            MoveFocus fgStrategyBaskets
    End Select
    SetButtons vsTypeTabs.CurrTab

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.vsTypeTabs.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsTypeTabs_Switch
'' Description: When the user changes tabs, set the buttons appropriately
'' Inputs:      Old Tab, New Tab, Whether or not to Cancel the Switch
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsTypeTabs_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:
   
    'If NewTab = Tabs(eTab_Criteria) Or NewTab = Tabs(eTab_Filters) Then
    If 0 Then
        If Not ScansEnabled Then
            If InfBox("Filters are currently turned off.|Would you like to turn them on?", "?", "+Yes|-No", "Criteria/Filters") = "Y" Then
                ScansEnabled = True
                FixFilterDisplay
            ElseIf OldTab < 0 Then ' = Tabs(eTab_Criteria) Or OldTab = Tabs(eTab_Filters) Then
                Cancel = True
                vsTypeTabs.CurrTab = Tabs(eTab_SymbolGroups)
            Else
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    
    If NewTab <> OldTab Then
        'If OldTab = Tabs(eTab_Filters) Then Enable cmdEdit, True
        Enable cmdEdit
        
        Select Case NewTab
            Case Tabs(eTab_SymbolGroups)
                If fgGroups.Rows = fgGroups.FixedRows Then LoadSymbolGroupGrid
            Case Tabs(eTab_Criteria)
                If fgCriteria.Rows = fgCriteria.FixedRows Then LoadCriteriaGrid
                FixFilterDisplay
            Case Tabs(eTab_Filters)
                If fgFilters.Rows = fgFilters.FixedRows Then LoadFiltersGrid
                FixFilterDisplay
            Case Tabs(eTab_Functions)
                If fgFunctions.Rows = fgFunctions.FixedRows Then LoadFunctionsGrid
            Case Tabs(eTab_Rules)
                If fgRules.Rows = fgRules.FixedRows Then LoadRulesGrid
            Case Tabs(eTab_Systems)
                If fgSystems.Rows = fgSystems.FixedRows Then LoadSystemsGrid
            Case Tabs(eTab_Libraries)
                If fgLibraries.Rows = fgLibraries.FixedRows Then LoadLibrariesGrid
            Case Tabs(eTab_StrategyBaskets)
                If fgStrategyBaskets.Rows = fgStrategyBaskets.FixedRows Then LoadStrategyBasketsGrid
        End Select
        SetButtons NewTab
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.vsTypeTabs.Switch", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowAddRules
'' Description: Show the form in the "Add Rules to System" Mode
'' Inputs:      System ID, Library ID of the System
'' Returns:     Array of Rules to Add (Size = 0 if Cancelled)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowAddRules(ByVal lSystemNumber As Long, ByVal lLibraryID As Long, SystemRules As cRules, _
    Optional ByVal lFilter As Long = -1) As cRules
On Error GoTo ErrSection:

    ' Set up the module level variables
    m.Mode = eAddMode_Add
    m.lSystemNumber = lSystemNumber
    m.lLibraryID = lLibraryID
    Set m.SystemRules = SystemRules
  
    ' Initialize and Load the Rules Grid
    InitRulesGrid
    LoadRulesGrid
    
    ' Filter the Rules Grid
    fgRules.ColHidden(RGCol(eRGC_SystemName)) = (chkFavorites = vbChecked)
    cboLibrary.Enabled = (chkLibrary = vbChecked)
    
    Select Case lFilter         'for strategy assistant
        Case 0:
            optLong.Value = True
        Case 1:
            optLongExit.Value = True
        Case 2:
            optShort.Value = True
        Case 3:
            optShortExit.Value = True
    End Select
    
    FilterRulesGrid
    If chkTreeRules.Value = vbChecked Then
        fgRules.Redraw = flexRDNone
        fgRules.Outline 1
        fgRules.IsCollapsed(m.lUsedInStrategiesRow) = flexOutlineCollapsed
        SetBackColors fgRules
        fgRules.Redraw = flexRDBuffered
    End If
            
    Me.Caption = "Add Rules to a Strategy ..."
    ShowOnlyTab Tabs(eTab_Rules)
    SetButtons Tabs(eTab_Rules)
    ShowForm Me, True, , , ALT_GRID_ROW_COLOR
    
    If Not m.bOK Then
        'Set m.alReturnIds = New cGdArray
        'm.alReturnIds.Create eGDARRAY_Longs
        'm.alReturnIds.Size = 0
        Set m.RulesToAdd = Nothing
    End If
    'Set ShowAddRules = m.alReturnIds
    Set ShowAddRules = m.RulesToAdd
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    RaiseError "frmToolbox.ShowAddRules", eGDRaiseError_Show
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowFunctionList
'' Description: Show the list of functions only with only a Close button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowFunctionList()
On Error GoTo ErrSection:

    m.Mode = eAddMode_List
    
    ' Initialize and Load the Functions Grid
    InitFunctionsGrid
    LoadFunctionsGrid
    
    Me.Caption = "Functions"
    ShowOnlyTab Tabs(eTab_Functions)
    SetButtons Tabs(eTab_Functions)
    ShowForm Me, True
    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.ShowFunctionList", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn(Grid As VSFlexGrid, Optional ByVal lResizeCol As Long = -1)
On Error GoTo ErrSection:
    
    Dim lExtCol As Long                 ' Extended column
    Dim lTotal As Long                  ' New width of the extended column
    Dim lIndex As Long                  ' Index into a for loop
    Dim lWidthDiff As Long
    
    ' set which column is the custom extended column
    If Grid Is fgGroups Then
        lExtCol = GGCol(eGGC_Description)
    ElseIf Grid Is fgCriteria Then
        lExtCol = CGCol(eCGC_Description)
    ElseIf Grid Is fgFilters Then
        lExtCol = FiCol(eFIC_Description)
    ElseIf Grid Is fgFunctions Then
        lExtCol = FGCol(eFGC_FunctionName)
    ElseIf Grid Is fgRules Then
        lExtCol = RGCol(eRGC_RuleName)
    ElseIf Grid Is fgSystems Then
        lExtCol = SGCol(eSGC_SystemDesc)
    ElseIf Grid Is fgLibraries Then
        lExtCol = LGCol(eLGC_LibraryName)
    End If
    
    With Grid
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If lResizeCol >= lExtCol Then
            .Redraw = flexRDNone
            lWidthDiff = .ColWidth(lResizeCol) - m.lPrevColWidth
            For lIndex = lResizeCol + 1 To .Cols - 1
                If Not .ColHidden(lIndex) Then
                    .ColWidth(lIndex) = .ColWidth(lIndex) - lWidthDiff
                    Exit For
                End If
            Next
            m.lPrevColWidth = 0
        End If
        
        .ColHidden(lExtCol) = True
        .Redraw = flexRDBuffered '(so .ClientWidth will be correct)
        .Redraw = flexRDNone
        lTotal = 0 * Screen.TwipsPerPixelX
        For lIndex = 0 To .Cols - 1
            If Not .ColHidden(lIndex) Then
                lTotal = lTotal + .ColWidth(lIndex)
            End If
        Next
        lTotal = .ClientWidth - lTotal
        If lTotal > 0 Then .ColWidth(lExtCol) = lTotal
        .ColHidden(lExtCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.ExtendCustomColumn", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

Private Sub ShowOnlyTab(ByVal nTab)

    Dim n&
    
    With vsTypeTabs
        If nTab >= 0 And nTab < .NumTabs Then
            For n = 0 To .NumTabs - 1
                If n = nTab Then
                    .TabVisible(n) = True
                    .CurrTab = n
                Else
                    .TabVisible(n) = False
                End If
            Next
        End If
    End With

End Sub

Private Function GetCurrentGrid() As VSFlexGrid

    Select Case vsTypeTabs.CurrTab
    Case Tabs(eTab_SymbolGroups)
        Set GetCurrentGrid = Me.fgGroups
    Case Tabs(eTab_Criteria)
        Set GetCurrentGrid = Me.fgCriteria
    Case Tabs(eTab_Filters)
        Set GetCurrentGrid = Me.fgFilters
    Case Tabs(eTab_Functions)
        Set GetCurrentGrid = Me.fgFunctions
    Case Tabs(eTab_Rules)
        Set GetCurrentGrid = Me.fgRules
    Case Tabs(eTab_Systems)
        Set GetCurrentGrid = Me.fgSystems
    Case Tabs(eTab_Libraries)
        Set GetCurrentGrid = Me.fgLibraries
    Case Tabs(eTab_StrategyBaskets)
        Set GetCurrentGrid = Me.fgStrategyBaskets
    End Select

End Function

Private Sub SetUpPopup(ByVal bShowDependencies As Boolean)
On Error GoTo ErrSection:

    Dim Basket As cStrategyBasket       ' Strategy basket item

    mnuAdd.Visible = cmdAdd.Visible
    mnuAddCopy.Visible = cmdAddCopy.Visible
    mnuChangeFont.Visible = True
    mnuDelete.Visible = cmdDelete.Visible
    mnuDependencies.Visible = bShowDependencies
    mnuEdit.Visible = cmdEdit.Visible
    mnuImport.Visible = cmdImport.Visible
    mnuInfo.Visible = cmdInfo.Visible
    mnuNew.Visible = cmdNew.Visible
    mnuNewDLL.Visible = cmdNewDLL.Visible
    mnuRename.Visible = cmdRename.Visible
    mnuExport.Visible = cmdExport.Visible
    mnuExportList.Visible = (FileExist("C:\Common\Files.EXE") And (vsTypeTabs.CurrTab = Tabs(eTab_Systems)))
    mnuCreateAutoTrade.Visible = ((vsTypeTabs.CurrTab = Tabs(eTab_Systems)) Or (vsTypeTabs.CurrTab = Tabs(eTab_StrategyBaskets)))
    
    mnuEdit.Enabled = True
    mnuDelete.Enabled = True
    mnuDependencies.Enabled = True
    
    ' We are not going to allow the user to edit a DLL function if they don't
    ' have the SDK...
    Select Case vsTypeTabs.CurrTab
        Case Tabs(eTab_Functions)
            If fgFunctions.TextMatrix(fgFunctions.RowSel, RGCol(eFGC_TreeLevel)) = "1" Then
                If fgFunctions.TextMatrix(fgFunctions.RowSel, FGCol(eFGC_ImplType)) <> "2" Then
                    mnuEdit.Enabled = DirExist(AddSlash(App.Path) & "..\SDK")
                Else
                    mnuEdit.Enabled = True
                End If
            Else
                mnuEdit.Enabled = False
            End If
            mnuDelete.Enabled = fgFunctions.TextMatrix(fgFunctions.RowSel, RGCol(eFGC_TreeLevel)) = "1"
            mnuDependencies.Enabled = fgFunctions.TextMatrix(fgFunctions.RowSel, RGCol(eFGC_TreeLevel)) = "1"
            
        Case Tabs(eTab_Rules)
            mnuEdit.Enabled = fgRules.TextMatrix(fgRules.RowSel, RGCol(eRGC_TreeLevel)) = "2"
            mnuDelete.Enabled = fgRules.TextMatrix(fgRules.RowSel, RGCol(eRGC_TreeLevel)) = "2"
            mnuDependencies.Enabled = fgRules.TextMatrix(fgRules.RowSel, RGCol(eRGC_TreeLevel)) = "2"
    
        Case Tabs(eTab_StrategyBaskets)
            Set Basket = fgStrategyBaskets.RowData(fgStrategyBaskets.RowSel)
            mnuCreateAutoTrade.Enabled = Not Basket.IsGuru
    
    End Select
    
    mnuExportAll.Visible = (vsTypeTabs.CurrTab = Tabs(eTab_Libraries)) And IsIDE
    
    If vsTypeTabs.CurrTab = Tabs(eTab_Criteria) Then
        mnuRenameFile.Visible = True
        mnuRenameFile.Enabled = ((fgCriteria.Row >= fgCriteria.FixedRows) And (fgCriteria.Row < fgCriteria.Rows))
    Else
        mnuRenameFile.Visible = False
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.SetUpPopup", eGDRaiseError_Raise
   
End Sub

Private Sub DoPrevSort(fg As VSFlexGrid)
On Error GoTo ErrSection:

    With fg
        If Parse(m.strPrevSort, vbTab, 1) = .Name Then
            .Col = Val(Parse(m.strPrevSort, vbTab, 2))
            If Val(Parse(m.strPrevSort, vbTab, 3)) = 2 Then
                .Sort = flexSortGenericDescending
            Else
                .Sort = flexSortGenericAscending
            End If
            .Tag = m.strPrevSort
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.DoPrevSort", eGDRaiseError_Raise
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpColumns
'' Description: Set up the column order/width/visibility according to spec
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpColumns(ByVal lTab As eAddFormTabs)
On Error GoTo ErrSection:

    Dim strFields As String             ' Fields from the ini file
    Dim strIniFile As String            ' File path and name of the ini file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrFields As New cGdArray      ' Array of field information
    Dim strColName As String            ' Column Name
    Dim strHidden As String             ' Is Column Hidden?
    Dim lColWidth As Long               ' Width of the column
    Dim lTotWidth As Long               ' Total width from the fields string
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lColPos As Long                 ' Column position
    Dim strPropName As String           ' Property name in the INI file
    Dim vsGrid As VSFlexGrid            ' Grid to work with
    
    Select Case lTab
        Case Tabs(eTab_SymbolGroups)
            strPropName = "SymbolGroupFields"
            Set vsGrid = fgGroups
            
        Case Tabs(eTab_Criteria)
            strPropName = "CriteriaFields"
            Set vsGrid = fgCriteria
            
        Case Tabs(eTab_Filters)
            strPropName = "FilterFields"
            Set vsGrid = fgFilters
            
        Case Tabs(eTab_Functions)
            strPropName = "FunctionFields"
            Set vsGrid = fgFunctions
            
        Case Tabs(eTab_Rules)
            strPropName = "RuleFields"
            Set vsGrid = fgRules
            
        Case Tabs(eTab_Systems)
            strPropName = "SystemFields"
            Set vsGrid = fgSystems
            
        Case Tabs(eTab_StrategyBaskets)
            strPropName = "SystemRunFields"
            Set vsGrid = fgStrategyBaskets
            
        Case Tabs(eTab_Libraries)
            strPropName = "LibraryFields"
            Set vsGrid = fgLibraries
            
    End Select
    
    strFields = GetIniFileProperty(strPropName, "", "Toolbox", g.strIniFile)
    lTotWidth = 0&
    
    With vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        astrFields.SplitFields strFields, ","
        For lIndex = 0 To astrFields.Size - 1
            strColName = Parse(astrFields(lIndex), ";", 1)
            strHidden = Parse(astrFields(lIndex), ";", 2)
            lColWidth = CLng(ValOfText(Parse(astrFields(lIndex), ";", 3)))
            
            For lCol = 0 To .Cols - 1
                If UCase(.TextMatrix(0, lCol)) = UCase(strColName) Then
                    .ColPosition(lCol) = lCol
                    lTotWidth = lTotWidth + lColWidth
                    If strHidden = "-1" Then
                        .ColHidden(lCol) = True
                    Else
                        .ColHidden(lCol) = False
                    End If
                    .ColWidth(lCol) = lColWidth
                    
                    Exit For
                End If
            Next lCol
        Next lIndex
        
        If lTotWidth = 0& Then
            .AutoSize 0, .Cols - 1, False, 75
            m.abAutoSize(lTab) = 1
        Else
            m.abAutoSize(lTab) = 0
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Set astrFields = Nothing
    Exit Sub
    
ErrSection:
    Set astrFields = Nothing
    RaiseError "frmToolbox.SetUpColumns", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveCols
'' Description: Save the column information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveCols(ByVal lTab As eAddFormTabs)
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of field information
    Dim lIndex As Long                  ' Index into a for loop
    Dim strPropName As String           ' Property name in the INI file
    Dim vsGrid As VSFlexGrid            ' Grid to work with
    Dim strFields As String             ' Fields to save to ini file

    Select Case lTab
        Case Tabs(eTab_SymbolGroups)
            strPropName = "SymbolGroupFields"
            Set vsGrid = fgGroups
            
        Case Tabs(eTab_Criteria)
            strPropName = "CriteriaFields"
            Set vsGrid = fgCriteria
            
        Case Tabs(eTab_Filters)
            strPropName = "FilterFields"
            Set vsGrid = fgFilters
            
        Case Tabs(eTab_Functions)
            strPropName = "FunctionFields"
            Set vsGrid = fgFunctions
            
        Case Tabs(eTab_Rules)
            strPropName = "RuleFields"
            Set vsGrid = fgRules
            
        Case Tabs(eTab_Systems)
            strPropName = "SystemFields"
            Set vsGrid = fgSystems
            
        Case Tabs(eTab_StrategyBaskets)
            strPropName = "SystemRunFields"
            Set vsGrid = fgStrategyBaskets
            
        Case Tabs(eTab_Libraries)
            strPropName = "LibraryFields"
            Set vsGrid = fgLibraries
            
    End Select
    
    astrFields.Create eGDARRAY_Strings
    For lIndex = 0 To vsGrid.Cols - 1
        If vsGrid.ColHidden(lIndex) = True Then
            astrFields.Add vsGrid.TextMatrix(0, lIndex) & ";-1;" & Str(vsGrid.ColWidth(lIndex))
        Else
            astrFields.Add vsGrid.TextMatrix(0, lIndex) & ";0;" & Str(vsGrid.ColWidth(lIndex))
        End If
    Next lIndex
    
    strFields = astrFields.JoinFields(",")
    SetIniFileProperty strPropName, strFields, "Toolbox", g.strIniFile
    m.abAutoSize(lTab) = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.SaveCols", eGDRaiseError_Raise
    
End Sub

Private Sub SaveFunctionFavorites()
On Error GoTo ErrSection:

    Dim astrFavorites As New cGdArray
    Dim lIndex As Long
    
    astrFavorites.Create eGDARRAY_Strings
    
    With fgFunctions
        If .Rows > .FixedRows Then
            For lIndex = .FixedRows To .Rows - 1
                If CheckedCell(fgFunctions, lIndex, FGCol(eFGC_Favorites)) = True Then
                    astrFavorites.Add .TextMatrix(lIndex, FGCol(eFGC_CodedName))
                End If
            Next lIndex
            
            ' TLB 2/24/05: temporary fix so won't write a 0-byte file
            ' (should later find out why we were writing a 0-byte file and fix that)
            If astrFavorites.Size > 0 Then
                astrFavorites.Sort
                astrFavorites.ToFile AddSlash(App.Path) & "Custom\Function.FAV", , , False
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.SaveFunctionFavorites", eGDRaiseError_Raise
    
End Sub

Private Sub ChangeRulesView()
On Error GoTo ErrSection:

    With fgRules
        .Redraw = flexRDNone
        If chkTreeRules.Value = vbChecked Then
            .Col = RGCol(eRGC_TreeSortKey)
            .Sort = flexSortStringAscending
            .OutlineBar = flexOutlineBarSimpleLeaf
            .OutlineCol = RGCol(eRGC_RuleName)
            .ColHidden(RGCol(eRGC_SystemName)) = True
        Else
            .OutlineBar = flexOutlineBarNone
            .ColHidden(RGCol(eRGC_SystemName)) = False
        End If
        FilterRulesGrid
        
        .ShowCell .RowSel, RGCol(eRGC_RuleName)
        .Redraw = flexRDBuffered
    End With
    
    MoveFocus fgRules

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.ChangeRulesView", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterFunctionsGrid
'' Description: Filter the functions grid as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterFunctionsGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgFunctions
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            .RowHidden(lIndex) = False
            
            If chkTreeFunctions.Value = vbChecked Then
                If .TextMatrix(lIndex, FGCol(eFGC_TreeLevel)) = "0" Then
                    .RowHidden(lIndex) = False
                End If
                
                .RowOutlineLevel(lIndex) = Val(.TextMatrix(lIndex, FGCol(eFGC_TreeLevel)))
                .IsSubtotal(lIndex) = True
            Else
                If .TextMatrix(lIndex, FGCol(eFGC_TreeLevel)) = "0" Then
                    .RowHidden(lIndex) = True
                End If
                
                .RowOutlineLevel(lIndex) = 0
                .IsSubtotal(lIndex) = False
            End If
            
            If chkFuncFav.Value = vbChecked Then
                If CheckedCell(fgFunctions, lIndex, FGCol(eFGC_Favorites)) = False Then
                    If chkTreeFunctions.Value = vbUnchecked Or .TextMatrix(lIndex, FGCol(eFGC_TreeLevel)) = "1" Then
                        If .RowHidden(lIndex) = False Then .RowHidden(lIndex) = True
                    End If
                End If
            End If
        Next lIndex
        
        SetBackColors fgFunctions
        
        ExtendCustomColumn fgFunctions
        .Redraw = flexRDBuffered
    End With
    
    SetButtons vsTypeTabs.CurrTab

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.FilterFunctionsGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeFunctionsView
'' Description: Allow the user to choose between Tree and Grid for Functions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeFunctionsView()
On Error GoTo ErrSection:

    With fgFunctions
        .Redraw = flexRDNone
        If chkTreeFunctions.Value = vbChecked Then
            .Col = FGCol(eFGC_TreeSortKey)
            .Sort = flexSortStringAscending
            .OutlineBar = flexOutlineBarSimpleLeaf
            .OutlineCol = FGCol(eFGC_FunctionName)
            .ColHidden(FGCol(eFGC_Category)) = True
        Else
            .OutlineBar = flexOutlineBarNone
            .ColHidden(FGCol(eFGC_Category)) = False
        End If
        FilterFunctionsGrid
        .ShowCell .RowSel, FGCol(eFGC_FunctionName)
        .Redraw = flexRDBuffered
    End With
    
    MoveFocus fgFunctions

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbox.ChangeFunctionsView", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportLibrary
'' Description: Export the given library
'' Inputs:      Library Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExportLibrary(ByVal strLibraryName As String)
On Error GoTo ErrSection:

    Dim lLibraryID As Long              ' ID of the library passed in
    Dim strPassword As String           ' Password of the library passed in
    Dim nSecurity As Byte               ' Security Level of the library passed in
    Dim rs As Recordset                 ' Recordset into the database
    Dim LMB As New cLibManagerBridge    ' Bridge to the Library Manager
    Dim astrDepends As New cGdArray     ' Dependency array
    Dim strItems As String              ' Items to add to the library
    Dim strName As String               ' Name of item to add to the library
    Dim strID As String                 ' ID of the item to export
    Dim bReload As Boolean              ' Do we need to reload libraries?
    
    bReload = False
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] WHERE [LibraryName]='" & strLibraryName & "';", dbOpenDynaset)
    If rs.BOF And rs.EOF Then
        Err.Raise vbObjectError + 1000, , "Trade Navigator could not export|" & strLibraryName & "|because it does not exist."
    Else
        lLibraryID = rs!LibraryID
        strPassword = DecryptField(rs!Password)
        nSecurity = rs!SecurityLevel
        
        If lLibraryID = kSN_UserLibrary Then
            If InfBox("The item you chose to Export is currently in the User Library.  Trade Navigator will need to create a library for this item before it can do the Export.", "!", "+OK|-Cancel", "Export Warning") = "O" Then
                Select Case vsTypeTabs.CurrTab
                    Case Tabs(eTab_Functions)
                        With fgFunctions
                            strID = .TextMatrix(.SelectedRow(0), FGCol(eFGC_FunctionID))
                            strItems = "Function," & strID
                            strName = .TextMatrix(.SelectedRow(0), FGCol(eFGC_FunctionName))
                        End With
                    
                    Case Tabs(eTab_Rules)
                        With fgRules
                            strID = .TextMatrix(.SelectedRow(0), RGCol(eRGC_RuleID))
                            strItems = "Rule," & strID
                            strName = .TextMatrix(.SelectedRow(0), RGCol(eRGC_RuleName))
                        End With
                    
                    Case Tabs(eTab_Systems)
                        With fgSystems
                            strID = .TextMatrix(.SelectedRow(0), SGCol(eSGC_SystemNumber))
                            strItems = "System," & strID
                            strName = .TextMatrix(.SelectedRow(0), SGCol(eSGC_SystemDesc))
                        End With
                End Select
                
                If Len(strItems) > 0 Then
                    Set LMB = GetLibMgrBridge
                    If LMB.CreateLibraryFromItems(strName, strItems, lLibraryID) = True Then
                        'Export
                        LMB.ShowPackager lLibraryID
                        
                        bReload = True
                    End If
                End If
            End If
        Else
            If g.Security.CanEdit(nSecurity, strPassword, strLibraryName) = True Then
                Set LMB = GetLibMgrBridge
                LMB.ShowPackager lLibraryID
                bReload = LMB.Reload
            End If
        End If
    End If
    
    If bReload Then
        Screen.MousePointer = vbHourglass
        StatusMsg "Reloading Libraries ...", vbRed
        
        ' Reload the libraries grid...
        LoadLibrariesGrid
        
        ' Reload the Function and Rule tables in memory...
        'LoadEngineFunctions
        RefreshLibrary lLibraryID
        
        ' Trigger a reload the next time they switch to these tabs...
        Select Case vsTypeTabs.CurrTab
            Case Tabs(eTab_Functions)
                fgRules.Rows = fgRules.FixedRows
                fgSystems.Rows = fgSystems.FixedRows
                With fgFunctions
                    .Redraw = flexRDNone
                    .Rows = .FixedRows
                    m.strInitialSelect = strID
                    LoadFunctionsGrid
                    .Redraw = flexRDBuffered
                End With
                
            Case Tabs(eTab_Rules)
                fgFunctions.Rows = fgFunctions.FixedRows
                fgSystems.Rows = fgSystems.FixedRows
                With fgRules
                    .Redraw = flexRDNone
                    .Rows = .FixedRows
                    m.strInitialSelect = strID
                    LoadRulesGrid
                    .Redraw = flexRDBuffered
                End With
                
            Case Tabs(eTab_Systems)
                fgFunctions.Rows = fgFunctions.FixedRows
                fgRules.Rows = fgRules.FixedRows
                With fgSystems
                    .Redraw = flexRDNone
                    .Rows = .FixedRows
                    m.strInitialSelect = strID
                    LoadSystemsGrid
                    .Redraw = flexRDBuffered
                End With
                
        End Select
        
        Screen.MousePointer = vbDefault
        StatusMsg
    End If

ErrExit:
    Set LMB = Nothing
    Set rs = Nothing
    Set astrDepends = Nothing
    Exit Sub
    
ErrSection:
    Set LMB = Nothing
    Set rs = Nothing
    Set astrDepends = Nothing
    RaiseError "frmToolbox.ExportLibrary", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Export
'' Description: Export the library that the selected item is in
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Export()
On Error GoTo ErrSection:

    Dim strLibraryName As String        ' Library name from the grid
    
    Select Case vsTypeTabs.CurrTab
        Case Tabs(eTab_Functions)
            With fgFunctions
                If .SelectedRows > 0 Then
                    strLibraryName = .TextMatrix(.SelectedRow(0), FGCol(eFGC_LibraryName))
                End If
            End With
            
        Case Tabs(eTab_Rules)
            With fgRules
                If .SelectedRows > 0 Then
                    strLibraryName = .TextMatrix(.SelectedRow(0), RGCol(eRGC_LibraryName))
                End If
            End With
        
        Case Tabs(eTab_Systems)
            With fgSystems
                If .SelectedRows > 0 Then
                    strLibraryName = .TextMatrix(.SelectedRow(0), SGCol(eSGC_LibraryName))
                End If
            End With
            
        Case Tabs(eTab_Libraries)
            With fgLibraries
                If .SelectedRows > 0 Then
                    strLibraryName = .TextMatrix(.SelectedRow(0), LGCol(eLGC_LibraryName))
                End If
            End With
            
    End Select
    
    If Len(strLibraryName) > 0 Then
        ExportLibrary strLibraryName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.Export", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Rename
'' Description: Allow the user to rename a Symbol Group, Criteria, or Filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Rename()
On Error GoTo ErrSection:

    Dim obj As Object
    Dim strReturn As String
    Dim strText As String
    
    Select Case vsTypeTabs.CurrTab
        Case Tabs(eTab_SymbolGroups)
            strText = "Rename current Symbol Group as ..."
            With fgGroups
                Set obj = g.SymbolPool.PoolObject("GRP:" & .TextMatrix(.Row, GGCol(eGGC_GroupID)))
                If Not obj Is Nothing Then
                    obj.FromFile AddSlash(App.Path) & "Custom", obj.ID, True
                    obj.Rename
                    .TextMatrix(.Row, GGCol(eGGC_Name)) = obj.Name
                    Set obj = Nothing
                End If
                
#If 0 Then
                strReturn = AskBox("h=Rename ; i=? ; g=string ; d=" & .TextMatrix(.Row, GGCol(eGGC_Name)) & " ; " & strText)
                If strReturn <> "" And strReturn <> .TextMatrix(.Row, GGCol(eGGC_Name)) Then
                    Set obj = g.SymbolPool.PoolObject("GRP:" & .TextMatrix(.Row, GGCol(eGGC_GroupID)))
                    If Not obj Is Nothing Then
                        obj.FromFile AddSlash(App.Path) & "Custom", obj.ID, True
                        obj.Name = strReturn
                        obj.ToFile
                        obj.AddToPool ' to replace the name in the fields table
                        frmSymbolGrid.RefreshGrid
                        .TextMatrix(.Row, GGCol(eGGC_Name)) = strReturn
                        Set obj = Nothing
                    End If
                End If
#End If
            End With
            
        Case Tabs(eTab_Criteria)
            strText = "Rename current Criteria as..."
            With fgCriteria
                strReturn = AskBox("h=Rename ; i=? ; g=string ; d=" & .TextMatrix(.Row, CGCol(eCGC_Name)) & " ; " & strText)
                If strReturn <> "" And strReturn <> .TextMatrix(.Row, CGCol(eCGC_Name)) Then
                    Set obj = g.SymbolPool.PoolObject("SCN:" & .TextMatrix(.Row, CGCol(eCGC_CriteriaID)))
                    If Not obj Is Nothing Then
                        obj.Name = strReturn
                        obj.ToFile
                        obj.AddToPool ' to replace the name in the fields table
                        frmSymbolGrid.RefreshGrid
                        .TextMatrix(.Row, CGCol(eCGC_Name)) = strReturn
                        Set obj = Nothing
                    End If
                End If
            End With
        
        Case Tabs(eTab_Filters)
            strText = "Rename current Filter as..."
            With fgFilters
                strReturn = AskBox("h=Rename ; i=? ; g=string ; d=" & .TextMatrix(.Row, FiCol(eFIC_Name)) & " ; " & strText)
                If strReturn <> "" And strReturn <> .TextMatrix(.Row, FiCol(eFIC_Name)) Then
                    Set obj = g.SymbolPool.PoolObject("FIL:" & .TextMatrix(.Row, FiCol(eFIC_FilterID)))
                    If Not obj Is Nothing Then
                        obj.Name = strReturn
                        obj.ToFile
                        obj.AddToPool ' to replace the name in the fields table
                        frmSymbolGrid.RefreshGrid
                        .TextMatrix(.Row, FiCol(eFIC_Name)) = strReturn
                        Set obj = Nothing
                    End If
                End If
            End With
    
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.Rename"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameFile
'' Description: Allow the user to rename a Criteria file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RenameFile()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strName As String               ' Name of the criteria

    With fgCriteria
        If (.Row >= .FixedRows) And (.Row < .Rows) Then
            strName = .TextMatrix(.Row, CGCol(eCGC_Name))
            If RenameCriteriaFile(.TextMatrix(.Row, CGCol(eCGC_CriteriaID)), strName) Then
                LoadCriteriaGrid
                
                For lIndex = .FixedRows To .Rows - 1
                    If UCase(.TextMatrix(lIndex, CGCol(eCGC_Name))) = UCase(strName) Then
                        .Row = lIndex
                        .RowSel = lIndex
                        .ShowCell lIndex, 0
                        
                        Exit For
                    End If
                Next lIndex
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.RenameFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpGridRow
'' Description: Dump all of the columns from a row in the grid to the Debug Log
'' Inputs:      Grid, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpGridRow(Grid As VSFlexGrid, ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Column in the grid
    Dim strToDump As String             ' String to dump to the log

    With Grid
        strToDump = vbTab & .TextMatrix(lRow, 0)
        For lCol = 1 To .Cols - 1
            strToDump = strToDump & "|" & .TextMatrix(lRow, lCol)
        Next lCol
    End With
    
    DebugLog strToDump

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.DumpGridRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResaveFunctions
'' Description: Resave (and reverify) the selected functions
'' Inputs:      None
'' Returns:     True if all successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ResaveFunctions() As Boolean
On Error GoTo ErrSection:

    Dim alFunctionIDs As cGdArray       ' Array of function IDs to send
    Dim lIndex As Long                  ' Index into a for loop
    Dim lID As Long                     ' Function ID from selected function
    Dim lImplementationType As Long     ' Implementation Type for the selected function
    
    Set alFunctionIDs = New cGdArray
    alFunctionIDs.Create eGDARRAY_Longs
    
    With fgFunctions
        For lIndex = 0 To .SelectedRows - 1
            If .RowHidden(.SelectedRow(lIndex)) = False Then
                lID = CLng(Val(.TextMatrix(.SelectedRow(lIndex), FGCol(eFGC_FunctionID))))
                lImplementationType = CLng(Val(.TextMatrix(.SelectedRow(lIndex), FGCol(eFGC_ImplType))))
                
                If (lID > 0) And (lImplementationType = 2) Then
                    alFunctionIDs.Add lID
                End If
            End If
        Next lIndex
    End With
    
    Me.Hide
    
    ResaveFunctions = frmFunctionMgrCT.ResaveFunctions(alFunctionIDs)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmToolbox.ResaveFunctions"
    
End Function

Private Sub FixFilterDisplay()
On Error GoTo ErrSection:

    Dim bScansEnabled As Boolean
    
    bScansEnabled = ScansEnabled
    
    If bScansEnabled Then
        ' show Active columns
        'fgFilters.ColHidden(FiCol(eFIC_Active)) = False
        fraFilterMsg.Visible = False
        fgFilters.Visible = True
        lblCriteria.Caption = "CRITERIA: active criteria are calculated for all symbols after end-of-day downloads (so can be used in filters) ..."
        If fgCriteria.ColHidden(CGCol(eCGC_Active)) <> False Then
            fgCriteria.ColHidden(CGCol(eCGC_Active)) = False
            ExtendCustomColumn fgCriteria
        End If
    Else
        ' hide Active columns
        'fgFilters.ColHidden(FiCol(eFIC_Active)) = True
        fraFilterMsg.Visible = True
        fgFilters.Visible = False
        lblCriteria.Caption = "CRITERIA: conditions and values which can be used in filters, quote board fields & alerts, snapshot window, etc."
        If fgCriteria.ColHidden(CGCol(eCGC_Active)) <> True Then
            fgCriteria.ColHidden(CGCol(eCGC_Active)) = True
            ExtendCustomColumn fgCriteria
        End If
    End If
    
    If chkFilters.Value <> Abs(bScansEnabled) Then
        chkFilters.Value = Abs(bScansEnabled)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmToolbox.FixFilterDisplay"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanDeleteStrategyBasket
'' Description: Determine if a strategy basket can be deleted
'' Inputs:      Strategy Basket, Confirm with user?
'' Returns:     True if basket can be deleted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CanDeleteStrategyBasket(ByVal Basket As cStrategyBasket, Optional ByVal bConfirm As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strMessage As String            ' Message to display to the user
    
    bReturn = True
    strMessage = ""
    
    If g.TradingItems.IsStrategyBasketInPosition(Basket.ID) Then
        strMessage = "Cannot delete '" & Basket.Name & "' because it is used in an automated trading item that is in a position"
    ElseIf g.TradingItems.IsStrategyBasketAutoTrading(Basket.ID) Then
        strMessage = "Cannot delete '" & Basket.Name & "' because it is used in an active automated trading item"
    End If
    
    If Len(strMessage) > 0 Then
        InfBox strMessage, "!", , "Delete Error"
        bReturn = False
    ElseIf g.TradingItems.IsStrategyBasketInAutoTradeItem(Basket.ID) Then
        bReturn = (InfBox("'" & Basket.Name & "' is contained in one or more|automated trading items.  Deleting this|basket will result in those automated|trading items being deleted as well.||Do you want to continue?|", "?", "Delete|+-Cancel", "Confirmation") = "D")
    ElseIf bConfirm = True Then
        bReturn = (InfBox("Are you sure you want to delete '" & Basket.Name & "'?", "?", "Delete|+-Cancel", "Confirmation") = "D")
    End If
    
    CanDeleteStrategyBasket = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmToolbox.CanDeleteStrategyBasket"
    
End Function

