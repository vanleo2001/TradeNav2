VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmChartCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chart Settings:"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12090
   ForeColor       =   &H00000000&
   Icon            =   "ChartCfg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   Begin vsOcx6LibCtl.vsElastic vseGeneral 
      Height          =   6165
      Left            =   2640
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   10874
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
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
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
      Begin HexUniControls.ctlUniFrameWL fraMisc 
         Height          =   1980
         Left            =   120
         TabIndex        =   88
         Top             =   4080
         Width           =   3030
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
         Caption         =   "ChartCfg.frx":0442
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":047E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":049E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdShowBidAsk 
            Height          =   315
            Left            =   1410
            TabIndex        =   3
            Top             =   240
            Width           =   1425
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
            Caption         =   "ChartCfg.frx":04BA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":04F2
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0512
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSplitPaneCfg 
            Height          =   315
            Left            =   2100
            TabIndex        =   10
            Top             =   1545
            Width           =   825
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
            Caption         =   "ChartCfg.frx":052E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":0560
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0580
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSplitPane 
            Height          =   225
            Left            =   105
            TabIndex        =   22
            Top             =   1590
            Width           =   2690
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
            Caption         =   "ChartCfg.frx":059C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":05E4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0604
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkChartTips 
            Height          =   255
            Left            =   105
            TabIndex        =   28
            Top             =   600
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
            Caption         =   "ChartCfg.frx":0620
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":0680
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":06A0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSplitsRolls 
            Height          =   255
            Left            =   105
            TabIndex        =   30
            Top             =   1200
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
            Caption         =   "ChartCfg.frx":06BC
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":0700
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":07B0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTips 
            Height          =   255
            Left            =   105
            TabIndex        =   36
            Top             =   900
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
            Caption         =   "ChartCfg.frx":07CC
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":082E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":084E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdFont 
            Height          =   315
            Left            =   105
            TabIndex        =   108
            Top             =   240
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
            Caption         =   "ChartCfg.frx":086A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":08A0
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":08C0
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor clrSplitsRolls 
            Height          =   315
            Left            =   1785
            TabIndex        =   89
            Top             =   1170
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   556
            CustomColor     =   255
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraDefaults 
         Height          =   2070
         Left            =   120
         TabIndex        =   78
         Top             =   1995
         Width           =   3030
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
         Caption         =   "ChartCfg.frx":08DC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":0918
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":0938
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboBoxXP cboAnnotDefault 
            Height          =   315
            Left            =   1155
            TabIndex        =   123
            Top             =   1680
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
            Tip             =   "ChartCfg.frx":0954
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0974
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboIndLabelDefault 
            Height          =   315
            Left            =   1155
            TabIndex        =   98
            Top             =   1320
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
            Tip             =   "ChartCfg.frx":0990
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":09B0
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboBarsDefault 
            Height          =   315
            Left            =   1155
            TabIndex        =   49
            Top             =   240
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
            Tip             =   "ChartCfg.frx":09CC
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":09EC
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboHorzDefault 
            Height          =   315
            Left            =   1155
            TabIndex        =   53
            Top             =   960
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
            Tip             =   "ChartCfg.frx":0A08
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0A28
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboLineDefault 
            Height          =   315
            Left            =   1155
            TabIndex        =   51
            Top             =   600
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
            Tip             =   "ChartCfg.frx":0A44
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0A64
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAnnotDefault 
            Height          =   195
            Left            =   240
            Top             =   1740
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
            Caption         =   "ChartCfg.frx":0A80
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":0AB6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0AD6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBarsDefault 
            Height          =   195
            Left            =   255
            Top             =   300
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
            Caption         =   "ChartCfg.frx":0AF2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":0B26
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0B46
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblLineDefault 
            Height          =   195
            Left            =   255
            Top             =   660
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
            Caption         =   "ChartCfg.frx":0B62
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":0B96
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0BB6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblHorzDefault 
            Height          =   195
            Left            =   255
            Top             =   1020
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
            Caption         =   "ChartCfg.frx":0BD2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":0C08
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0C28
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblIndLabelDefault 
            Height          =   195
            Left            =   255
            Top             =   1380
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
            Caption         =   "ChartCfg.frx":0C44
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":0C7A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0C9A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraColors 
         Height          =   1875
         Left            =   120
         TabIndex        =   38
         Top             =   60
         Width           =   3030
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
         Caption         =   "ChartCfg.frx":0CB6
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":0D0C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":0D2C
         RightToLeft     =   0   'False
         Begin gdOCX.gdSelectColor clrChartBack 
            Height          =   315
            Left            =   1920
            TabIndex        =   39
            Top             =   870
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor clrChartFore 
            Height          =   315
            Left            =   900
            TabIndex        =   44
            Top             =   870
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor clrBorderBack 
            Height          =   315
            Left            =   1920
            TabIndex        =   45
            Top             =   480
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor clrBorderFore 
            Height          =   315
            Left            =   900
            TabIndex        =   46
            Top             =   480
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniCheckXP chkCustomColors 
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   1560
            Width           =   2475
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
            Caption         =   "ChartCfg.frx":0D48
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":0DA6
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0DC6
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor clrChartGradient 
            Height          =   315
            Left            =   1920
            TabIndex        =   40
            Top             =   1245
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniCheckXP chkGradient 
            Height          =   225
            Left            =   240
            TabIndex        =   41
            Top             =   1320
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
            Caption         =   "ChartCfg.frx":0DE2
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":0E28
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0E48
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label20 
            Height          =   255
            Left            =   225
            Top             =   900
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
            Caption         =   "ChartCfg.frx":0E64
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":0E90
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0EB0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label19 
            Height          =   255
            Left            =   225
            Top             =   540
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
            Caption         =   "ChartCfg.frx":0ECC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":0EFA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0F1A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label18 
            Height          =   255
            Left            =   1980
            Top             =   240
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
            Caption         =   "ChartCfg.frx":0F36
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":0F68
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0F88
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label17 
            Height          =   255
            Left            =   960
            Top             =   240
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
            Caption         =   "ChartCfg.frx":0FA4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":0FD6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":0FF6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin VB.Image imgGradient 
            Height          =   600
            Left            =   2820
            Picture         =   "ChartCfg.frx":1012
            Top             =   900
            Width           =   225
         End
         Begin VB.Image imgGradientWhite 
            Height          =   375
            Left            =   2835
            Picture         =   "ChartCfg.frx":177C
            Top             =   1020
            Visible         =   0   'False
            Width           =   105
         End
      End
      Begin HexUniControls.ctlUniLabelXP Label26 
         Height          =   195
         Left            =   60
         Top             =   60
         Visible         =   0   'False
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
         Caption         =   "ChartCfg.frx":18CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "ChartCfg.frx":1932
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":1952
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsElastic vseXaxis 
      Height          =   5640
      Left            =   1140
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9948
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
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
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
      Begin HexUniControls.ctlUniFrameWL fraDates 
         Height          =   2955
         Left            =   120
         TabIndex        =   66
         Top             =   2460
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":196E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":19A8
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":19C8
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdStartStop 
            Height          =   315
            Left            =   1875
            TabIndex        =   42
            Top             =   2520
            Width           =   900
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
            Caption         =   "ChartCfg.frx":19E4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":1A10
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1A30
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectDate dtToDate 
            Height          =   315
            Left            =   600
            TabIndex        =   80
            Top             =   840
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            AllowWeekends   =   0   'False
            Value           =   37274
         End
         Begin HexUniControls.ctlUniTextBoxXP txtMaxDays 
            Height          =   285
            Left            =   2220
            TabIndex        =   68
            Top             =   1530
            Width           =   495
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":1A4C
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
            Tip             =   "ChartCfg.frx":1A70
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1A90
         End
         Begin gdOCX.gdSelectDate dtFromDate 
            Height          =   315
            Left            =   600
            TabIndex        =   69
            Top             =   300
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            AllowWeekends   =   0   'False
            MaxDate         =   42611
            MaxDateIsToday  =   -1  'True
            Value           =   2
         End
         Begin HexUniControls.ctlUniRadioXP optEndOfData 
            Height          =   240
            Left            =   300
            TabIndex        =   82
            Top             =   1170
            Width           =   1515
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
            Caption         =   "ChartCfg.frx":1AAC
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "ChartCfg.frx":1AE4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1B04
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optToDate 
            Height          =   240
            Left            =   300
            TabIndex        =   81
            Top             =   900
            Width           =   2415
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
            Caption         =   "ChartCfg.frx":1B20
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":1B52
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1B72
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblStartStopInfo2 
            Height          =   195
            Left            =   230
            Top             =   2250
            Width           =   2220
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
            Caption         =   "ChartCfg.frx":1B8E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":1BE4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1C04
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblStartStopTimes 
            Height          =   255
            Left            =   230
            Top             =   2565
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
            Caption         =   "ChartCfg.frx":1C20
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":1C5A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1C7A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblStartStopInfo1 
            Height          =   195
            Left            =   180
            Top             =   2055
            Width           =   2625
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
            Caption         =   "ChartCfg.frx":1C96
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":1CF8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1D18
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label25 
            Height          =   195
            Left            =   180
            Top             =   660
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
            Caption         =   "ChartCfg.frx":1D34
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":1D5A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1D7A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblMaxDays1 
            Height          =   195
            Left            =   180
            Top             =   1500
            Width           =   1995
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
            Caption         =   "ChartCfg.frx":1D96
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":1DEE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1E0E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   195
            Left            =   180
            Top             =   360
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
            Caption         =   "ChartCfg.frx":1E2A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":1E54
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1E74
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblMaxDays2 
            Height          =   195
            Left            =   300
            Top             =   1680
            Width           =   1935
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
            Caption         =   "ChartCfg.frx":1E90
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":1EE6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1F06
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraXaxis 
         Height          =   2175
         Left            =   120
         TabIndex        =   70
         Top             =   60
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":1F22
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":1F56
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":1F76
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtBarWidth 
            Height          =   285
            Left            =   1560
            TabIndex        =   43
            Top             =   1080
            Visible         =   0   'False
            Width           =   435
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":1F92
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "ChartCfg.frx":1FB4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":1FD4
         End
         Begin HexUniControls.ctlUniComboBoxXP cboVertGrid 
            Height          =   315
            Left            =   1560
            TabIndex        =   109
            Top             =   720
            Width           =   1155
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
            Tip             =   "ChartCfg.frx":1FF0
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2010
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkEmptyBars 
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1380
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
            Caption         =   "ChartCfg.frx":202C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":208A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":20AA
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkHorzGrid 
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   2055
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
            Caption         =   "ChartCfg.frx":20C6
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":2110
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2130
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboBarPeriod 
            Height          =   315
            Left            =   1200
            TabIndex        =   72
            Top             =   360
            Width           =   1515
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
            Tip             =   "ChartCfg.frx":214C
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
            MouseIcon       =   "ChartCfg.frx":216C
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdBarPeriod 
            Height          =   315
            Left            =   120
            TabIndex        =   71
            Top             =   360
            Width           =   975
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
            Caption         =   "ChartCfg.frx":2188
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":21BE
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":21DE
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtForecastBars 
            Height          =   285
            Left            =   2280
            TabIndex        =   74
            Top             =   1710
            Width           =   435
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":21FA
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
            Tip             =   "ChartCfg.frx":221E
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":223E
         End
         Begin HexUniControls.ctlUniLabelXP Label22 
            Height          =   255
            Left            =   120
            Top             =   780
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
            Caption         =   "ChartCfg.frx":225A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":22A2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":22C2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label21 
            Height          =   195
            Left            =   120
            Top             =   1740
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
            Caption         =   "ChartCfg.frx":22DE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":2338
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2358
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin vsOcx6LibCtl.vsElastic vsePane 
      Height          =   6375
      Left            =   7965
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   225
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   11245
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
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
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
      Begin HexUniControls.ctlUniFrameWL fraLinearLog 
         Height          =   675
         Left            =   120
         TabIndex        =   48
         Top             =   0
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":2374
         Enabled         =   -1  'True
         ForeColor       =   8388608
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":23C2
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":23E2
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optScalePercent 
            Height          =   405
            Left            =   1830
            TabIndex        =   50
            Top             =   225
            Width           =   960
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
            Caption         =   "ChartCfg.frx":23FE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":243A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":245A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optScaleLog 
            Height          =   405
            Left            =   975
            TabIndex        =   52
            Top             =   225
            Width           =   960
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
            Caption         =   "ChartCfg.frx":2476
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":24A8
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":24C8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optScaleLinear 
            Height          =   405
            Left            =   60
            TabIndex        =   57
            Top             =   225
            Width           =   960
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
            Caption         =   "ChartCfg.frx":24E4
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":2510
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2530
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraPaneScale 
         Height          =   2640
         Left            =   120
         TabIndex        =   25
         Top             =   780
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":254C
         Enabled         =   -1  'True
         ForeColor       =   8388608
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":259E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":25BE
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optAutoScalePrice 
            Height          =   255
            Left            =   60
            TabIndex        =   59
            Top             =   540
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
            Caption         =   "ChartCfg.frx":25DA
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":263E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":265E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtMin 
            Height          =   285
            Left            =   1440
            TabIndex        =   29
            Top             =   1365
            Width           =   735
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":267A
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
            Tip             =   "ChartCfg.frx":269C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":26BC
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPerBar 
            Height          =   285
            Left            =   1080
            TabIndex        =   115
            Top             =   2220
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":26D8
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "ChartCfg.frx":26F8
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2718
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPtsOrTicks 
            Height          =   285
            Left            =   420
            TabIndex        =   113
            Top             =   1920
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":2734
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "ChartCfg.frx":2754
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2774
         End
         Begin HexUniControls.ctlUniRadioXP optSquareScale 
            Height          =   255
            Left            =   60
            TabIndex        =   112
            Top             =   1650
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
            Caption         =   "ChartCfg.frx":2790
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "ChartCfg.frx":27C8
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":27E8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtMax 
            Height          =   285
            Left            =   1440
            TabIndex        =   31
            Top             =   1050
            Width           =   735
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":2804
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
            Tip             =   "ChartCfg.frx":282A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":284A
         End
         Begin HexUniControls.ctlUniRadioXP optManualScale 
            Height          =   255
            Left            =   60
            TabIndex        =   27
            Top             =   810
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
            Caption         =   "ChartCfg.frx":2866
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":28B2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":28D2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optAutoScale 
            Height          =   255
            Left            =   60
            TabIndex        =   26
            Top             =   270
            Width           =   2415
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
            Caption         =   "ChartCfg.frx":28EE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":294C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":296C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraDummy 
            Height          =   435
            Left            =   1020
            TabIndex        =   117
            Top             =   1770
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
            Caption         =   "ChartCfg.frx":2988
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "ChartCfg.frx":29A8
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":29C8
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optPtsPerBar 
               Height          =   220
               Left            =   60
               TabIndex        =   119
               Top             =   180
               Width           =   855
               _ExtentX        =   1508
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
               Caption         =   "ChartCfg.frx":29E4
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "ChartCfg.frx":2A10
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "ChartCfg.frx":2A30
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optTicksPerBar 
               Height          =   220
               Left            =   900
               TabIndex        =   118
               Top             =   180
               Width           =   855
               _ExtentX        =   1508
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
               Caption         =   "ChartCfg.frx":2A4C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "ChartCfg.frx":2A76
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "ChartCfg.frx":2A96
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniLabelXP lblPer 
            Height          =   195
            Left            =   420
            Top             =   2250
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
            Caption         =   "ChartCfg.frx":2AB2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":2AD8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2AF8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblMaxValue 
            Height          =   195
            Left            =   540
            Top             =   1080
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
            Caption         =   "ChartCfg.frx":2B14
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":2B4A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2B6A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblMinValue 
            Height          =   195
            Left            =   540
            Top             =   1380
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
            Caption         =   "ChartCfg.frx":2B86
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":2BBC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2BDC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBars 
            Height          =   195
            Left            =   1740
            Top             =   2250
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
            Caption         =   "ChartCfg.frx":2BF8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":2C24
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2C44
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniCheckXP chkPriceTopMost 
         Height          =   225
         Left            =   135
         TabIndex        =   62
         Top             =   5505
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":2C60
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":2CC6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":2CE6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkYscaleLabelAll 
         Height          =   225
         Left            =   135
         TabIndex        =   63
         Top             =   5205
         Width           =   2385
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
         Caption         =   "ChartCfg.frx":2D02
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":2D56
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":2D76
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraPaneDisplay 
         Height          =   1275
         Left            =   120
         TabIndex        =   32
         Top             =   3525
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":2D92
         Enabled         =   -1  'True
         ForeColor       =   8388608
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":2DCE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":2DEE
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtDecimals 
            Height          =   285
            Left            =   1290
            TabIndex        =   37
            Top             =   540
            Width           =   375
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":2E0A
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
            Tip             =   "ChartCfg.frx":2E2C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2E4C
         End
         Begin HexUniControls.ctlUniRadioXP optDisplayFormat 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   35
            Top             =   570
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
            Caption         =   "ChartCfg.frx":2E68
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":2E98
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2EB8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optDisplayFormat 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   240
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
            Caption         =   "ChartCfg.frx":2ED4
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "ChartCfg.frx":2F08
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2F28
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optDisplayFormat 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   33
            Top             =   900
            Width           =   2415
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
            Caption         =   "ChartCfg.frx":2F44
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":2F90
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":2FB0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label16 
            Height          =   255
            Left            =   1740
            Top             =   570
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
            Caption         =   "ChartCfg.frx":2FCC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":2FFC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":301C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniCheckXP chkHideSeparator 
         Height          =   225
         Left            =   135
         TabIndex        =   85
         Top             =   4905
         Width           =   2385
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
         Caption         =   "ChartCfg.frx":3038
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":308A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":30AA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsElastic vseLinkedInputs 
      Height          =   5355
      Left            =   7050
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   465
      Visible         =   0   'False
      Width           =   3150
      _ExtentX        =   5556
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
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
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
      Begin HexUniControls.ctlUniFrameWL fraLinkedInputs 
         Height          =   5115
         Left            =   120
         TabIndex        =   75
         Top             =   60
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":30C6
         Enabled         =   -1  'True
         ForeColor       =   8388608
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":3100
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":3120
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgLinkedInputs 
            Height          =   2760
            Left            =   345
            TabIndex        =   77
            Top             =   1860
            Width           =   2205
            _cx             =   3889
            _cy             =   4868
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
         Begin HexUniControls.ctlUniLabelXP lblLinkedInputUsage 
            Height          =   1305
            Left            =   120
            Top             =   300
            Width           =   2640
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
            Caption         =   "ChartCfg.frx":313C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":32EA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":330A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin vsOcx6LibCtl.vsElastic vsElastic1 
      Height          =   30
      Left            =   6705
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   4020
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
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
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
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
   Begin vsOcx6LibCtl.vsElastic vseBars 
      Height          =   5595
      Left            =   540
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9869
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
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
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
      Begin HexUniControls.ctlUniFrameWL fraBars 
         Height          =   2235
         Left            =   120
         TabIndex        =   147
         Top             =   3360
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":3326
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":335A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":337A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkRemoveOvernightGap 
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   1875
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
            Caption         =   "ChartCfg.frx":3396
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":33DE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":33FE
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdMarketInfo 
            Height          =   315
            Left            =   240
            TabIndex        =   156
            Top             =   1320
            Width           =   2415
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
            Caption         =   "ChartCfg.frx":341A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":3472
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3492
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSymbol 
            Height          =   435
            Left            =   240
            TabIndex        =   149
            Top             =   300
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
            Caption         =   "ChartCfg.frx":34AE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":34DC
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":34FC
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkUnsplit 
            Height          =   255
            Left            =   240
            TabIndex        =   148
            Top             =   1560
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
            Caption         =   "ChartCfg.frx":3518
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":355E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":357E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSymbol 
            Height          =   195
            Left            =   1260
            Top             =   420
            Width           =   1155
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
            Caption         =   "ChartCfg.frx":359A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":35C6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":35E6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDesc 
            Height          =   435
            Left            =   240
            Top             =   840
            Width           =   2475
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
            Caption         =   "ChartCfg.frx":3602
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":3692
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":36B2
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraAppearanceBars 
         Height          =   3195
         Left            =   120
         TabIndex        =   131
         Top             =   60
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":36CE
         Enabled         =   -1  'True
         ForeColor       =   8388608
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":3702
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":3722
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkFlip 
            Height          =   220
            Left            =   105
            TabIndex        =   84
            Top             =   2910
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "ChartCfg.frx":373E
            Enabled         =   0   'False
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":37A4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3820
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor gdTrueRange 
            Height          =   315
            Left            =   2040
            TabIndex        =   132
            ToolTipText     =   "Color for the extended portion of the true high or true low"
            Top             =   1440
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniCheckXP chkTrueRange 
            Height          =   195
            Left            =   240
            TabIndex        =   137
            Top             =   1500
            Width           =   1905
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
            Caption         =   "ChartCfg.frx":383C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":3880
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3938
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboBarsLabelMode 
            Height          =   315
            Left            =   840
            TabIndex        =   136
            Top             =   2235
            Width           =   1800
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
            Tip             =   "ChartCfg.frx":3954
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
            MouseIcon       =   "ChartCfg.frx":3974
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboLineTypesBars 
            Height          =   315
            Left            =   840
            TabIndex        =   135
            Top             =   1845
            Width           =   1800
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
            Tip             =   "ChartCfg.frx":3990
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":39B0
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboOHLCType 
            Height          =   315
            Left            =   840
            TabIndex        =   134
            Top             =   300
            Width           =   1800
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
            Tip             =   "ChartCfg.frx":39CC
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":39EC
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkOverlayed 
            Height          =   220
            Index           =   0
            Left            =   105
            TabIndex        =   133
            Top             =   2640
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "ChartCfg.frx":3A08
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":3A6C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3B2E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor gdColorBars 
            Height          =   315
            Left            =   840
            TabIndex        =   138
            Top             =   690
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdColorDown 
            Height          =   315
            Left            =   2040
            TabIndex        =   139
            ToolTipText     =   "Color to use if close is less than the open"
            Top             =   1080
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdColorUp 
            Height          =   315
            Left            =   840
            TabIndex        =   140
            ToolTipText     =   "Color to use if close is greater than the open"
            Top             =   1080
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniCheckXP chkUpDownColors 
            Height          =   300
            Left            =   240
            TabIndex        =   141
            Top             =   1080
            Width           =   1245
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
            Caption         =   "ChartCfg.frx":3B4A
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":3B70
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3BF4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblUpColor 
            Height          =   225
            Left            =   2160
            Top             =   120
            Width           =   435
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
            Caption         =   "ChartCfg.frx":3C10
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":3C36
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3CBA
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblWidthBars 
            Height          =   255
            Left            =   240
            Top             =   1905
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
            Caption         =   "ChartCfg.frx":3CD6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":3D02
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3D22
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label34 
            Height          =   255
            Left            =   240
            Top             =   2265
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
            Caption         =   "ChartCfg.frx":3D3E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":3D6A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3D8A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label35 
            Height          =   255
            Left            =   240
            Top             =   360
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
            Caption         =   "ChartCfg.frx":3DA6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":3DD0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3DF0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label37 
            Height          =   255
            Left            =   240
            Top             =   750
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
            Caption         =   "ChartCfg.frx":3E0C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":3E38
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3E58
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDownColor 
            Height          =   225
            Left            =   1530
            Top             =   1125
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
            Caption         =   "ChartCfg.frx":3E74
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":3E9E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":3F22
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin vsOcx6LibCtl.vsElastic vseTipInput 
      Height          =   240
      Left            =   5760
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   6600
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   423
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
      BackColor       =   8421504
      ForeColor       =   -2147483630
      FloodColor      =   12648447
      ForeColorDisabled=   -2147483631
      Caption         =   "vseTipX"
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   100
      CaptionPos      =   3
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
      Begin HexUniControls.ctlUniLabelXP lblTipInput 
         Height          =   135
         Left            =   165
         Top             =   60
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
         Caption         =   "ChartCfg.frx":3F3E
         BackColor       =   -2147483624
         ForeColor       =   -2147483625
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "ChartCfg.frx":3F6A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":3F8A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   390
      Left            =   5535
      TabIndex        =   1
      Top             =   6150
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
      Caption         =   "ChartCfg.frx":3FA6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "ChartCfg.frx":3FCC
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "ChartCfg.frx":3FEC
      RightToLeft     =   0   'False
   End
   Begin VB.Timer tmrChartCfg 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5700
      Top             =   5400
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   390
      Left            =   4635
      TabIndex        =   2
      Top             =   6150
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
      Caption         =   "ChartCfg.frx":4008
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "ChartCfg.frx":4036
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "ChartCfg.frx":409A
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP Corner 
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
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
      Caption         =   "ChartCfg.frx":40B6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "ChartCfg.frx":40D8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "ChartCfg.frx":40F8
      RightToLeft     =   0   'False
   End
   Begin vsOcx6LibCtl.vsElastic vseAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6180
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
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
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
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
      Begin HexUniControls.ctlUniButtonImageXP cmdTemplate 
         Height          =   330
         Left            =   3300
         TabIndex        =   87
         Top             =   0
         Width           =   1035
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
         Caption         =   "ChartCfg.frx":4114
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":4148
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":41A8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSaveStudy 
         Height          =   330
         Left            =   2160
         TabIndex        =   86
         Top             =   0
         Width           =   1080
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
         Caption         =   "ChartCfg.frx":41C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":41FA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":421A
         RightToLeft     =   0   'False
      End
      Begin VB.PictureBox picAdd 
         BorderStyle     =   0  'None
         Height          =   60
         Left            =   480
         Picture         =   "ChartCfg.frx":4236
         ScaleHeight     =   60
         ScaleWidth      =   120
         TabIndex        =   54
         Top             =   180
         Width           =   120
      End
      Begin VB.PictureBox picMoveDown 
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   1845
         Picture         =   "ChartCfg.frx":4540
         ScaleHeight     =   165
         ScaleWidth      =   150
         TabIndex        =   12
         ToolTipText     =   "Move item down"
         Top             =   90
         Width           =   150
      End
      Begin VB.PictureBox picMoveUp 
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   1605
         Picture         =   "ChartCfg.frx":484A
         ScaleHeight     =   165
         ScaleWidth      =   150
         TabIndex        =   11
         ToolTipText     =   "Move item up"
         Top             =   75
         Width           =   150
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   330
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   660
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
         Caption         =   "ChartCfg.frx":4B54
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":4B82
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":4BA2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
         Height          =   330
         Left            =   720
         TabIndex        =   8
         Top             =   0
         Width           =   780
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
         Caption         =   "ChartCfg.frx":4BBE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":4BEC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":4C0C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdMoveUp 
         Height          =   330
         Left            =   1560
         TabIndex        =   7
         Top             =   0
         Width           =   252
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
         Caption         =   "ChartCfg.frx":4C28
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":4C5A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":4C92
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdMoveDown 
         Height          =   330
         Left            =   1800
         TabIndex        =   6
         Top             =   0
         Width           =   252
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
         Caption         =   "ChartCfg.frx":4CAE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":4CE4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":4D20
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   180
         Left            =   1800
         Top             =   60
         Visible         =   0   'False
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
         Caption         =   "ChartCfg.frx":4D3C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "ChartCfg.frx":4D66
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":4D86
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin MSComctlLib.ImageList imgColors 
      Left            =   5880
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartCfg.frx":4DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartCfg.frx":50BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartCfg.frx":53D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex7LCtl.VSFlexGrid fgSettings 
      Height          =   5880
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   3000
      _cx             =   5292
      _cy             =   10372
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
      BackColorAlternate=   16777215
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
   Begin HexUniControls.ctlUniFrameWL fraTemplate 
      Height          =   972
      Left            =   3120
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   2412
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
      Caption         =   "ChartCfg.frx":56F0
      Enabled         =   -1  'True
      ForeColor       =   8388608
      BackColor       =   -2147483633
      Tip             =   "ChartCfg.frx":5732
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "ChartCfg.frx":5752
      RightToLeft     =   0   'False
      Begin vsOcx6LibCtl.vsElastic vseSaveAs 
         Height          =   195
         Left            =   1620
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   660
         Width           =   620
         _ExtentX        =   1085
         _ExtentY        =   344
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
         ForeColor       =   8388608
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   "Save As"
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   3
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
      Begin vsOcx6LibCtl.vsElastic vseSave 
         Height          =   195
         Left            =   180
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   660
         Width           =   1212
         _ExtentX        =   2143
         _ExtentY        =   344
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
         ForeColor       =   8388608
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   "Save Template"
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   3
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
      Begin HexUniControls.ctlUniButtonImageXP cmdSaveAs 
         Height          =   270
         Left            =   1560
         TabIndex        =   16
         Top             =   630
         Width           =   732
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
         Caption         =   "ChartCfg.frx":576E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":57A0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":57C0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSave 
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   1332
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
         Caption         =   "ChartCfg.frx":57DC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "ChartCfg.frx":580A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":582A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboBoxXP cboTemplate 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2175
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   8388608
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
         Tip             =   "ChartCfg.frx":5846
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
         MouseIcon       =   "ChartCfg.frx":5866
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsElastic vseSystem 
      Height          =   5580
      Left            =   3360
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9843
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
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
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
      Begin HexUniControls.ctlUniFrameWL fraTrades 
         Height          =   1095
         Left            =   120
         TabIndex        =   177
         Top             =   60
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":5882
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":58D4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":58F4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdTradeSettings 
            Height          =   375
            Left            =   1800
            TabIndex        =   93
            Top             =   710
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
            Caption         =   "ChartCfg.frx":5910
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":5940
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5960
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optTradesStrategy 
            Height          =   255
            Left            =   180
            TabIndex        =   94
            Top             =   520
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
            Caption         =   "ChartCfg.frx":597C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":59AC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":59CC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optTradesAccount 
            Height          =   255
            Left            =   180
            TabIndex        =   179
            Top             =   770
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
            Caption         =   "ChartCfg.frx":59E8
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":5A26
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5A46
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optTradesNone 
            Height          =   255
            Left            =   180
            TabIndex        =   178
            Top             =   270
            Width           =   2475
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
            Caption         =   "ChartCfg.frx":5A62
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":5A8A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5AAA
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraProfitLines 
         Height          =   1725
         Left            =   120
         TabIndex        =   124
         Top             =   3540
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":5AC6
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":5B0A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":5B2A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optLineBox 
            Height          =   195
            Index           =   3
            Left            =   1260
            TabIndex        =   170
            Top             =   540
            Width           =   1275
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
            Caption         =   "ChartCfg.frx":5B46
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":5B82
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5BA2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optLineBox 
            Height          =   195
            Index           =   2
            Left            =   1260
            TabIndex        =   169
            Top             =   300
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
            Caption         =   "ChartCfg.frx":5BBE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":5BFA
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5C1A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optLineBox 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   167
            Top             =   540
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
            Caption         =   "ChartCfg.frx":5C36
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":5C60
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5C80
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optLineBox 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   166
            Top             =   300
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
            Caption         =   "ChartCfg.frx":5C9C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":5CC4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5CE4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboProfitLines 
            Height          =   315
            Left            =   525
            TabIndex        =   128
            Top             =   850
            Width           =   2235
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
            Tip             =   "ChartCfg.frx":5D00
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5D20
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin gdOCX.gdSelectColor clrLoss 
            Height          =   315
            Left            =   1980
            TabIndex        =   125
            Top             =   1275
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   556
            Color           =   16711680
            CustomColor     =   16711680
         End
         Begin gdOCX.gdSelectColor clrWin 
            Height          =   315
            Left            =   540
            TabIndex        =   126
            Top             =   1270
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label10 
            Height          =   195
            Left            =   1440
            Top             =   1335
            Width           =   600
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
            Caption         =   "ChartCfg.frx":5D3C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":5D68
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5D88
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label8 
            Height          =   195
            Left            =   120
            Top             =   910
            Width           =   600
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
            Caption         =   "ChartCfg.frx":5DA4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":5DCE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5DEE
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label9 
            Height          =   195
            Left            =   120
            Top             =   1330
            Width           =   600
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
            Caption         =   "ChartCfg.frx":5E0A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":5E32
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5E52
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSystem 
         Height          =   2175
         Left            =   120
         TabIndex        =   91
         Top             =   1260
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":5E6E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":5EAE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":5ECE
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdSelectSystem 
            Height          =   375
            Left            =   120
            TabIndex        =   101
            Top             =   330
            Width           =   1275
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
            Caption         =   "ChartCfg.frx":5EEA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":5F28
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5F48
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdEditSystem 
            Height          =   375
            Left            =   1500
            TabIndex        =   92
            Top             =   330
            Width           =   1275
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
            Caption         =   "ChartCfg.frx":5F64
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":5FA0
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5FC0
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor clrLong 
            Height          =   315
            Left            =   600
            TabIndex        =   95
            Top             =   1200
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   556
            Color           =   16711680
            CustomColor     =   16711680
         End
         Begin gdOCX.gdSelectColor clrShort 
            Height          =   315
            Left            =   1980
            TabIndex        =   97
            Top             =   1200
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniComboBoxXP cboAccounts 
            Height          =   315
            Left            =   240
            TabIndex        =   96
            Top             =   590
            Width           =   2415
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
            Tip             =   "ChartCfg.frx":5FDC
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":5FFC
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAccounts 
            Height          =   255
            Left            =   240
            Top             =   320
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
            Caption         =   "ChartCfg.frx":6018
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":6056
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6076
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSystemName 
            Height          =   405
            Left            =   120
            Top             =   780
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
            Caption         =   "ChartCfg.frx":6092
            BackColor       =   -2147483633
            ForeColor       =   8388608
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":613E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":615E
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
         Begin HexUniControls.ctlUniLabelXP Label31 
            Height          =   195
            Left            =   180
            Top             =   1260
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
            Caption         =   "ChartCfg.frx":617A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":61A2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":61C2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label30 
            Height          =   195
            Left            =   1560
            Top             =   1260
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
            Caption         =   "ChartCfg.frx":61DE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":6208
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6228
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label6 
            Height          =   435
            Left            =   180
            Top             =   1620
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
            Caption         =   "ChartCfg.frx":6244
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":62F8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6318
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin vsOcx6LibCtl.vsElastic vseIndicator 
      Height          =   5925
      Left            =   3180
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   3115
      _ExtentX        =   5503
      _ExtentY        =   10451
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
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
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
      Begin HexUniControls.ctlUniFrameWL fraArtPyramids 
         Height          =   2715
         Left            =   255
         TabIndex        =   99
         Top             =   1305
         Visible         =   0   'False
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":6334
         Enabled         =   -1  'True
         ForeColor       =   8388608
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":636C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":638C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkVolume_VA 
            Height          =   225
            Index           =   4
            Left            =   195
            TabIndex        =   100
            Top             =   1365
            Visible         =   0   'False
            Width           =   2275
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
            Caption         =   "ChartCfg.frx":63A8
            Enabled         =   0   'False
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":63F0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6410
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkVolume_VA 
            Height          =   225
            Index           =   3
            Left            =   195
            TabIndex        =   102
            Top             =   1365
            Visible         =   0   'False
            Width           =   2625
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
            Caption         =   "ChartCfg.frx":642C
            Enabled         =   0   'False
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":6484
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":64A4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkVolume_VA 
            Height          =   225
            Index           =   2
            Left            =   195
            TabIndex        =   106
            Top             =   1380
            Visible         =   0   'False
            Width           =   2275
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
            Caption         =   "ChartCfg.frx":64C0
            Enabled         =   0   'False
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":6506
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6526
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkVolume_VA 
            Height          =   225
            Index           =   1
            Left            =   195
            TabIndex        =   107
            Top             =   1350
            Visible         =   0   'False
            Width           =   2275
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
            Caption         =   "ChartCfg.frx":6542
            Enabled         =   0   'False
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":6584
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":65A4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPercentVolume_VA 
            Height          =   315
            Left            =   2190
            TabIndex        =   110
            Top             =   1720
            Visible         =   0   'False
            Width           =   570
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":65C0
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "ChartCfg.frx":65E4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6604
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPoints 
            Height          =   315
            Left            =   2190
            TabIndex        =   114
            Top             =   2150
            Visible         =   0   'False
            Width           =   570
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":6620
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "ChartCfg.frx":6642
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6662
         End
         Begin HexUniControls.ctlUniCheckXP chkHorzLetters 
            Height          =   195
            Left            =   465
            TabIndex        =   116
            Top             =   2340
            Width           =   2520
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
            Caption         =   "ChartCfg.frx":667E
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":66D0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":66F0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkArtChartLabel 
            Height          =   195
            Left            =   195
            TabIndex        =   122
            Top             =   2395
            Width           =   2520
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
            Caption         =   "ChartCfg.frx":670C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":6764
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6784
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdPyramidFont 
            Height          =   315
            Left            =   195
            TabIndex        =   127
            Top             =   1950
            Width           =   960
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
            Caption         =   "ChartCfg.frx":67A0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":67C8
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":67E8
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboPyramidStyle 
            Height          =   315
            Left            =   1320
            TabIndex        =   129
            Top             =   1320
            Width           =   1500
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
            Tip             =   "ChartCfg.frx":6804
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6824
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin gdOCX.gdSelectColor gdPyramidColorUp 
            Height          =   315
            Left            =   1860
            TabIndex        =   142
            Top             =   195
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdPyramidColorDown 
            Height          =   315
            Left            =   1860
            TabIndex        =   143
            Top             =   547
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdPyramidColorP 
            Height          =   315
            Left            =   1860
            TabIndex        =   144
            Top             =   900
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdPyramidColorLabel 
            Height          =   315
            Left            =   1320
            TabIndex        =   145
            Top             =   1950
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniCheckXP chkVolume_VA 
            Height          =   225
            Index           =   0
            Left            =   195
            TabIndex        =   146
            Top             =   1320
            Visible         =   0   'False
            Width           =   2275
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
            Caption         =   "ChartCfg.frx":6840
            Enabled         =   0   'False
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":6880
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":68A0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor gdColorVolume_POC 
            Height          =   315
            Left            =   1620
            TabIndex        =   150
            Top             =   900
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            Enabled         =   0   'False
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdColorVolume_VA 
            Height          =   315
            Left            =   1620
            TabIndex        =   151
            Top             =   647
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            Enabled         =   0   'False
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label5 
            Height          =   275
            Left            =   195
            Top             =   1775
            Visible         =   0   'False
            Width           =   1650
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
            Caption         =   "ChartCfg.frx":68BC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":68F8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6918
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblArtLabels 
            Height          =   255
            Left            =   195
            Top             =   1725
            Width           =   1785
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
            Caption         =   "ChartCfg.frx":6934
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":697E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":699E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblArtColorDown 
            Height          =   255
            Left            =   195
            Top             =   577
            Width           =   1350
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
            Caption         =   "ChartCfg.frx":69BA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":69F4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6A14
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblArtColorPending 
            Height          =   255
            Left            =   195
            Top             =   930
            Width           =   1350
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
            Caption         =   "ChartCfg.frx":6A30
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":6A74
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6A94
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblArtStyle 
            Height          =   255
            Left            =   195
            Top             =   1350
            Width           =   1350
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
            Caption         =   "ChartCfg.frx":6AB0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":6AEC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6B0C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblArtColorUp 
            Height          =   255
            Left            =   195
            Top             =   225
            Width           =   1350
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
            Caption         =   "ChartCfg.frx":6B28
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":6B5E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6B7E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraAppearance 
         Height          =   2900
         Left            =   120
         TabIndex        =   20
         Top             =   60
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":6B9A
         Enabled         =   -1  'True
         ForeColor       =   8388608
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":6BCE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":6BEE
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboBoxXP cboIndLabelMode 
            Height          =   315
            Left            =   900
            TabIndex        =   165
            Top             =   1470
            Width           =   1572
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
            Tip             =   "ChartCfg.frx":6C0A
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6C2A
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtBaseLineY 
            Height          =   285
            Left            =   360
            TabIndex        =   152
            Top             =   1440
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":6C46
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "ChartCfg.frx":6C68
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6C88
         End
         Begin HexUniControls.ctlUniComboBoxXP cboMarkerSize 
            Height          =   315
            Left            =   900
            TabIndex        =   155
            Top             =   1860
            Width           =   1572
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
            Tip             =   "ChartCfg.frx":6CA4
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6CC4
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtColorSeperator 
            Height          =   285
            Left            =   180
            TabIndex        =   173
            Top             =   1620
            Width           =   675
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":6CE0
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "ChartCfg.frx":6D02
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6D22
         End
         Begin HexUniControls.ctlUniCheckXP chkBiColorBars 
            Height          =   195
            Left            =   120
            TabIndex        =   172
            Top             =   2400
            Width           =   1920
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
            Caption         =   "ChartCfg.frx":6D3E
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":6D82
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6E44
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkColorPrice 
            Height          =   195
            Left            =   120
            TabIndex        =   111
            Top             =   1880
            Width           =   2340
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
            Caption         =   "ChartCfg.frx":6E60
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":6EB2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6ED2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor gdFillColor 
            Height          =   315
            Left            =   1980
            TabIndex        =   160
            Top             =   2235
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniTextBoxXP txtBarsRight 
            Height          =   315
            Left            =   2097
            TabIndex        =   164
            Top             =   1320
            Width           =   375
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":6EEE
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
            Tip             =   "ChartCfg.frx":6F10
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6F30
         End
         Begin HexUniControls.ctlUniComboBoxXP cboBoxPenStyle 
            Height          =   315
            Left            =   1680
            TabIndex        =   163
            Top             =   1680
            Width           =   1572
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
            Tip             =   "ChartCfg.frx":6F4C
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6F6C
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optBoxFill 
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   162
            Top             =   2295
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
            Caption         =   "ChartCfg.frx":6F88
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":6FBC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":6FDC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optBoxFill 
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   161
            Top             =   2295
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
            Caption         =   "ChartCfg.frx":6FF8
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":7026
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7046
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtBarsLeft 
            Height          =   315
            Left            =   2340
            TabIndex        =   159
            Top             =   900
            Width           =   375
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":7062
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
            Tip             =   "ChartCfg.frx":7084
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":70A4
         End
         Begin HexUniControls.ctlUniCheckXP chkOverlayed 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   153
            Top             =   2100
            Width           =   2700
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
            Caption         =   "ChartCfg.frx":70C0
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":7124
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":71E6
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboMarkerLoc 
            Height          =   315
            Left            =   1560
            TabIndex        =   120
            Top             =   1080
            Width           =   915
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
            Tip             =   "ChartCfg.frx":7202
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
            MouseIcon       =   "ChartCfg.frx":7222
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboType 
            Height          =   315
            Left            =   900
            TabIndex        =   61
            Top             =   690
            Width           =   1572
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
            Tip             =   "ChartCfg.frx":723E
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":725E
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboLineTypes 
            Height          =   315
            Left            =   900
            TabIndex        =   60
            Top             =   1080
            Width           =   1572
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
            Tip             =   "ChartCfg.frx":727A
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":729A
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin gdOCX.gdSelectColor gdColor 
            Height          =   315
            Left            =   900
            TabIndex        =   21
            Top             =   300
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniComboBoxXP cboHighlightBars 
            Height          =   315
            Left            =   1260
            TabIndex        =   64
            Top             =   600
            Visible         =   0   'False
            Width           =   1572
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
            Tip             =   "ChartCfg.frx":72B6
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   1
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":72D6
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTrendHistory 
            Height          =   195
            Left            =   1500
            TabIndex        =   175
            Top             =   1320
            Width           =   1920
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
            Caption         =   "ChartCfg.frx":72F2
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":7336
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":73F8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkShowInAllPanes 
            Height          =   255
            Left            =   240
            TabIndex        =   171
            Top             =   2100
            Width           =   2340
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
            Caption         =   "ChartCfg.frx":7414
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":7456
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7476
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectIcon gdSelectIcon 
            Height          =   345
            Left            =   2220
            TabIndex        =   121
            Top             =   1860
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   609
         End
         Begin HexUniControls.ctlUniFrameWL fraFakeDropdown 
            Height          =   375
            Left            =   120
            TabIndex        =   157
            Top             =   525
            Visible         =   0   'False
            Width           =   1665
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
            Caption         =   "ChartCfg.frx":7492
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "ChartCfg.frx":74B2
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":74D2
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniButtonImageXP cmdFakeDropdown 
               Height          =   270
               Left            =   1245
               TabIndex        =   158
               Top             =   30
               Visible         =   0   'False
               Width           =   300
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
               Caption         =   "ChartCfg.frx":74EE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "ChartCfg.frx":752C
               Style           =   1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "ChartCfg.frx":754C
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniTextBoxXP txtFakeDropdown 
               Height          =   315
               Left            =   0
               TabIndex        =   168
               Top             =   0
               Visible         =   0   'False
               Width           =   1575
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "ChartCfg.frx":7568
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
               Tip             =   "ChartCfg.frx":75AA
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "ChartCfg.frx":75CA
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid fgColors 
            Height          =   2055
            Left            =   75
            TabIndex        =   174
            Top             =   615
            Visible         =   0   'False
            Width           =   1575
            _cx             =   2778
            _cy             =   3625
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   0
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
            Rows            =   8
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   0
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
         Begin gdOCX.gdSelectColor clrFakeDropdown 
            Height          =   315
            Left            =   1530
            TabIndex        =   176
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP lblBaseLineY 
            Height          =   255
            Left            =   120
            Top             =   840
            Visible         =   0   'False
            Width           =   2055
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
            Caption         =   "ChartCfg.frx":75E6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":761C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":763C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblMarkerSize 
            Height          =   255
            Left            =   240
            Top             =   1890
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
            Caption         =   "ChartCfg.frx":7658
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":7682
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":76A2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblExtBars 
            Height          =   195
            Left            =   60
            Top             =   540
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
            Caption         =   "ChartCfg.frx":76BE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":76E6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7706
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBarsRight 
            Height          =   195
            Left            =   1980
            Top             =   120
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
            Caption         =   "ChartCfg.frx":7722
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":7758
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7778
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBarsLeft 
            Height          =   195
            Left            =   1140
            Top             =   120
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
            Caption         =   "ChartCfg.frx":7794
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":77C8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":77E8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblWidth 
            Height          =   255
            Left            =   240
            Top             =   1140
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
            Caption         =   "ChartCfg.frx":7804
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":7830
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7850
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblLabelMode 
            Height          =   255
            Left            =   240
            Top             =   1500
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
            Caption         =   "ChartCfg.frx":786C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":7898
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":78B8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblIndType 
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
            Caption         =   "ChartCfg.frx":78D4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":78FE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":791E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label36 
            Height          =   255
            Left            =   240
            Top             =   360
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
            Caption         =   "ChartCfg.frx":793A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":7966
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7986
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblExtAutoTrend 
            Height          =   195
            Left            =   0
            Top             =   180
            Visible         =   0   'False
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
            Caption         =   "ChartCfg.frx":79A2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":79E6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7A06
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraFunction 
         Height          =   2655
         Left            =   120
         TabIndex        =   55
         Top             =   3060
         Width           =   2895
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
         Caption         =   "ChartCfg.frx":7A22
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "ChartCfg.frx":7A52
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "ChartCfg.frx":7A72
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgInputs 
            Height          =   1395
            Left            =   120
            TabIndex        =   103
            Top             =   840
            Width           =   2655
            _cx             =   4683
            _cy             =   2461
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
         Begin HexUniControls.ctlUniFrameWL fraShift 
            Height          =   315
            Left            =   240
            TabIndex        =   104
            Top             =   2280
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
            Caption         =   "ChartCfg.frx":7A8E
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "ChartCfg.frx":7AAE
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7ACE
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtShiftBars 
               Height          =   288
               Left            =   720
               TabIndex        =   105
               Top             =   0
               Width           =   435
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "ChartCfg.frx":7AEA
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
               Tip             =   "ChartCfg.frx":7B0C
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "ChartCfg.frx":7B2C
            End
            Begin HexUniControls.ctlUniLabelXP lblShiftBars2 
               Height          =   255
               Left            =   1260
               Top             =   30
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
               Caption         =   "ChartCfg.frx":7B48
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "ChartCfg.frx":7B80
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "ChartCfg.frx":7BA0
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblShiftBars 
               Height          =   255
               Left            =   60
               Top             =   30
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
               Caption         =   "ChartCfg.frx":7BBC
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "ChartCfg.frx":7BEC
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "ChartCfg.frx":7C0C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniTextBoxXP txtFunction 
            Height          =   300
            Left            =   120
            TabIndex        =   58
            Top             =   270
            Width           =   1815
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "ChartCfg.frx":7C28
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
            Tip             =   "ChartCfg.frx":7C68
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7C88
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdEditFunction 
            Height          =   525
            Left            =   1980
            TabIndex        =   56
            Top             =   210
            Width           =   825
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
            Caption         =   "ChartCfg.frx":7CA4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "ChartCfg.frx":7CE4
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7D04
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblCustomFunction 
            Height          =   1335
            Left            =   120
            Top             =   900
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
            Caption         =   "ChartCfg.frx":7D20
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":7DB0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7DD0
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
         Begin HexUniControls.ctlUniLabelXP lblNameForDisplay 
            Height          =   255
            Left            =   240
            Top             =   570
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
            Caption         =   "ChartCfg.frx":7DEC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "ChartCfg.frx":7E30
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "ChartCfg.frx":7E50
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label2 
      Height          =   255
      Left            =   120
      Top             =   5760
      Visible         =   0   'False
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
      Caption         =   "ChartCfg.frx":7E6C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "ChartCfg.frx":7E9E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "ChartCfg.frx":7EBE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Begin VB.Menu mnuAddNewStudy 
         Caption         =   "Add &Study (prebuilt pane)"
      End
      Begin VB.Menu mnuAddIndicator 
         Caption         =   "Add &Indicator"
         Begin VB.Menu mnuAddToSelectedPane 
            Caption         =   "Add Indicator to Selected &Pane"
         End
         Begin VB.Menu mnuAddToSelectedIndicator 
            Caption         =   "&Attach new Indicator to Selected Indicator"
         End
         Begin VB.Menu mnuAddToNewPane 
            Caption         =   "Add Indicator to &New Pane"
         End
      End
      Begin VB.Menu mnuAddHorzLine 
         Caption         =   "Add &Horizontal Line to selected Pane"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompSymbol 
         Caption         =   "Add &Comparison Symbol"
      End
      Begin VB.Menu mnuAddHighlightBars 
         Caption         =   "Add Highlight&Bars to Selected Indicator"
      End
      Begin VB.Menu mnuAddSystem 
         Caption         =   "Add a &Trading Strategy"
      End
      Begin VB.Menu mnuSepRemove 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveStudy 
         Caption         =   "Save Pane as Custom Study"
      End
      Begin VB.Menu mnuNewPaneMove 
         Caption         =   "&Move to New Pane"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
   End
End
Attribute VB_Name = "frmChartCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const kFormWidth = 6795
Const kInvertedColors = True
Const kCollapseRows = True 'False

Const kMaxInputsFrame = 2

Const kUnknown = 0
Const kNewPane = 1
Const kGeneral = 2
Const kXaxis = 3
Const kSystem = 4
Const kLinkedInputs = 5
Const kPane = 10
Const kIndicator = 11
'Const kLinkedInd = 12

Const kBoxCol = 0
Const kShowCol = 1
Const kNameCol = 2

Const kInputsTab = 0
Const kStyleTab = 1
Const kScaleTab = 2
Const kPaneTab = 3
Const kDataTab = 4
Const kGenStyleTab = 5

Const kVseFullHeight = 5925

Public bNowAdding As Boolean

Private Type mPrivate
    Chart As cChart
    PaneWood As cPane
    EditedIndicator As cIndicator
    EditedPane As cPane
    iGenerateChart As Integer 'flag for timer to generate chart
    nRowWhenLeftForm As Long
    vFgInputsBeforeEditValue As Variant
    bClickOK As Boolean
    nEditCol As Long
    bSkipGenerate As Boolean
    bCondFuncNewInprog As Boolean           'flag to unload in case user cancelled from condition builder or function editor
    bSeasonal As Boolean
    aModifiedInputs As cGdArray ' to delay saving inputs until all have changed for the indicator (esp. for TAS indicators)
End Type

Private m As mPrivate

Private Function RowType(ByVal nRow&) As Integer
On Error GoTo ErrSection:

    Dim nNode&, nLevel&, nFixedRows&
    With fgSettings
        If nRow >= 0 And nRow < .Rows Then
            nNode = .RowData(nRow)
        End If
        nFixedRows = .FixedRows
        If nNode > 0 Then
            nLevel = m.Chart.Tree.NodeLevel(nNode)
            If nLevel = 0 Then
                RowType = kPane
            Else
                RowType = kIndicator
            End If
        ElseIf nRow < nFixedRows Then
            RowType = kUnknown 'header
        ElseIf nRow = nFixedRows Then
            RowType = kGeneral
        ElseIf nRow = nFixedRows + 1 Then
            RowType = kXaxis
        ElseIf nRow = nFixedRows + 2 Then
            RowType = kSystem
        ElseIf nRow = nFixedRows + 3 Then
            RowType = kLinkedInputs
        ElseIf nRow = .Rows - 1 Then
            RowType = kNewPane
        Else
            RowType = kUnknown
        End If
    End With
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmChartCfg.RowType", eGDRaiseError_Raise
        
End Function

Private Sub cboAnnotDefault_Click()

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    g.ChartGlobals.eDefaultAnnotStyle = cboAnnotDefault.ListIndex + 1
    m.iGenerateChart = 1

End Sub

Private Sub cboBarPeriod_Change()
On Error GoTo ErrSection:

    Dim nNewPer&
    
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    
    If m.Chart.ChangeBarPeriod(cboBarPeriod.Text, False) Then
        fgSettings.Redraw = flexRDNone '(so won't recurse)
        cboBarPeriod.Text = GetPeriodStr(m.Chart.Periodicity)
        fgSettings.Redraw = flexRDBuffered
        FixControls
        m.iGenerateChart = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboBarPeriod.Change", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboBarPeriod_Click()
On Error GoTo ErrSection:

    cboBarPeriod_Change

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboBarPeriod.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboBarsDefault_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    g.ChartGlobals.eDefaultBarsStyle = cboBarsDefault.ListIndex + 1
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboBarsDefault.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboBarsLabelMode_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
        
    m.EditedIndicator.IndLabelMode = cboBarsLabelMode.ListIndex
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboBarsLabelMode.Click", eGDRaiseError_Show
    Resume ErrExit


End Sub

Private Sub cboBoxPenStyle_Click()

    If m.EditedIndicator Is Nothing Then Exit Sub
    m.EditedIndicator.BoxPenStyle = cboBoxPenStyle.ListIndex
    m.iGenerateChart = True

End Sub

Private Sub cboHighlightBars_Click()
On Error GoTo ErrSection:
    
    Dim Ind As cIndicator
    Dim strMsg$
    
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    If m.EditedIndicator.DataType <> eINDIC_BooleanArray And Not m.EditedIndicator.IsHawkeyeAdds Then Exit Sub
    
    Set Ind = m.Chart.Tree(m.EditedIndicator.IndToColor)
    If Ind Is Nothing Then Exit Sub
    
    If cboHighlightBars.ListCount = 6 Then
        If cboHighlightBars.ListIndex = 0 Then
            m.EditedIndicator.DisplayType = eINDIC_HighlightBars
        ElseIf cboHighlightBars.ListIndex = 1 Then
            m.EditedIndicator.DisplayType = eINDIC_HighlightMarkers
        ElseIf cboHighlightBars.ListIndex = 2 Then
            m.EditedIndicator.DisplayType = eINDIC_HighlightBoxes
        ElseIf cboHighlightBars.ListIndex = 3 Then
            m.EditedIndicator.DisplayType = eINDIC_HighlightZones
            m.EditedIndicator.BoxFillStyle = 1
            If m.EditedIndicator.ShowInAllPanes = 0 Then m.EditedIndicator.ShowInAllPanes = 1
        ElseIf cboHighlightBars.ListIndex = 4 Then
            m.EditedIndicator.DisplayType = eINDIC_ValueMarkers
        Else
            m.EditedIndicator.DisplayType = eINDIC_NoStyle
        End If
    Else
    
    If cboHighlightBars.ListCount = 5 Then
        Select Case cboHighlightBars.ListIndex
            Case 0:
                m.EditedIndicator.DisplayType = eINDIC_HighlightBars
            Case 1:
                m.EditedIndicator.DisplayType = eINDIC_HighlightMarkers
            Case 2:
                m.EditedIndicator.DisplayType = eINDIC_HighlightBoxes
            Case 3:
                With m.EditedIndicator
                    .DisplayType = eINDIC_HighlightZones
                    .BoxFillStyle = 1
                    If .ShowInAllPanes = 0 Then .ShowInAllPanes = 1
                End With
            Case 4:
                m.EditedIndicator.DisplayType = eINDIC_NoStyle
        End Select
    Else
        Select Case cboHighlightBars.ListIndex
            Case 0:
                m.EditedIndicator.DisplayType = eINDIC_HighlightMarkers
            Case 1:
                m.EditedIndicator.DisplayType = eINDIC_HighlightBoxes
            Case 2:
                With m.EditedIndicator
                    .DisplayType = eINDIC_HighlightZones
                    .BoxFillStyle = 1
                    If .ShowInAllPanes = 0 Then .ShowInAllPanes = 1
                End With
            Case 3:
                m.EditedIndicator.DisplayType = eINDIC_NoStyle
        End Select
    End If
    
    End If
    
    FixControls
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboHighlightBars.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub cboHorzDefault_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    g.ChartGlobals.eDefaultHorzStyle = cboHorzDefault.ListIndex + 1
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboHorzDefault.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboIndLabelDefault_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    g.ChartGlobals.eDefaultLabelMode = cboIndLabelDefault.ListIndex + 1
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboIndLabelDefault_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboIndLabelMode_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    
    If Not m.EditedIndicator Is Nothing Then
        With m.EditedIndicator
             .IndLabelMode = cboIndLabelMode.ListIndex
        End With
        m.iGenerateChart = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboIndLabelMode.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboLineDefault_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    g.ChartGlobals.eDefaultIndStyle = cboLineDefault.ListIndex + 1
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboLineDefault.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboLineTypes_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
        
    m.EditedIndicator.Style = cboLineTypes.ListIndex
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboLineTypes.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboLineTypesBars_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
        
    m.EditedIndicator.Style = cboLineTypesBars.ListIndex
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboLineTypesBars.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboMarkerLoc_Click()
On Error GoTo ErrSection:

    If m.EditedIndicator Is Nothing Then Exit Sub
    m.EditedIndicator.MarkerLoc = cboMarkerLoc.ListIndex
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboMarkerLoc.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboMarkerSize_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    
    If Not m.EditedIndicator Is Nothing Then
        With m.EditedIndicator
            If .DisplayType = eINDIC_HighlightMarkers Or .DisplayType = eINDIC_ValueMarkers Then
                .MarkerSize = cboMarkerSize.ListIndex + 1
            End If
        End With
        m.iGenerateChart = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboMarkerSize_Click"
    Resume ErrExit

End Sub

Private Sub cboOHLCType_Click()
On Error GoTo ErrSection:

    Dim i&, nPeriods&, nPeriodType&
    Dim bGetPeriod As Boolean
    Dim eNewDisplayType As eIndicatorDisplayType, eOldDisplayType As eIndicatorDisplayType
    
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    eNewDisplayType = OHLCcombo
    If m.Chart.Tree.Key(m.EditedIndicator.geIndId) <> "PRICE" Then
        ' for secondary symbols, can't use special types
        If eNewDisplayType = eINDIC_Kagi Or eNewDisplayType = eINDIC_PNF Or eNewDisplayType = eINDIC_Renko Then
            InfBox "This display type can only be used|with the primary symbol of the chart.", "i", , cboOHLCType.Text
            Exit Sub
        End If
        m.EditedIndicator.DisplayType = eNewDisplayType
    Else
        ' for primary symbol, see if need to fix bar period
        eOldDisplayType = m.EditedIndicator.DisplayType
        nPeriodType = GetPeriodType(m.Chart.Periodicity)
        nPeriods = GetPeriodsPerBar(m.Chart.Periodicity)
        Select Case eNewDisplayType
        Case eINDIC_PNF
            bGetPeriod = True
            If nPeriodType >= ePRD_Days Then
                nPeriodType = ePRD_EodPF
                nPeriods = 100
            Else
                nPeriodType = ePRD_IntPF
                nPeriods = 10
            End If
        Case eINDIC_Kagi
            bGetPeriod = True
            If nPeriodType >= ePRD_Days Then
                nPeriodType = ePRD_EodKagi
                nPeriods = 100
            Else
                nPeriodType = ePRD_IntKagi
                nPeriods = 10
            End If
        Case eINDIC_Renko
            bGetPeriod = True
            If nPeriodType >= ePRD_Days Then
                nPeriodType = ePRD_EodRenko
                nPeriods = 100
            Else
                nPeriodType = ePRD_IntRenko
                nPeriods = 10
            End If
        Case Else
            If nPeriodType >= ePRD_EodRenko And nPeriodType <= ePRD_EodPF Then
                bGetPeriod = True
                nPeriodType = ePRD_Days
                nPeriods = 1
            ElseIf nPeriodType >= ePRD_IntRenko And nPeriodType <= ePRD_IntPF Then
                bGetPeriod = True
                nPeriodType = ePRD_Minutes
                nPeriods = 5
            End If
        End Select
        
        If Len(m.Chart.SpreadSymbols) > 0 Then
            ' call this in order to do special checks for Spread charts
            m.Chart.BarDisplayType = eNewDisplayType
        Else
            m.EditedIndicator.DisplayType = eNewDisplayType
            If Me.Visible And bGetPeriod Then
                i = frmBarPeriod.ShowMe(nPeriodType + nPeriods, m.Chart.Bars, m.Chart)          '6499
                If i = 0 Then
                    ' cancelled
                    m.EditedIndicator.DisplayType = eOldDisplayType
                ElseIf i <> m.Chart.Periodicity Then
                    m.Chart.ChangeBarPeriod i, False
                End If
            End If
        End If
    End If
    
    FixControls
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboOHLCType.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboPyramidStyle_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
        
    If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
        m.EditedIndicator.ProfileStyleTPO = cboPyramidStyle.ListIndex + 3
    Else
        m.EditedIndicator.Style = cboPyramidStyle.ListIndex
    End If
    
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboPyramidStyle_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboTemplate_Click()
On Error GoTo ErrSection:
    
    If Me.Visible Then
        Set m.Chart = Nothing
        LoadSettingsGrid
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboTemplate.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboType_Click()
On Error GoTo ErrSection:

    Dim i&, k&, nID&
    Dim Tree As cGdTree
    Dim Ind As cIndicator
    Dim aInd As cGdArray
    
    Dim ePrevDisplay As eIndicatorDisplayType
    Dim eNewDisplay As eIndicatorDisplayType
    
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    With m.EditedIndicator
        ePrevDisplay = .DisplayType
        
'JM 07-30-2015: think all this cluster code is obsolete (leave awhile then remove if all ok)
'        If .DisplayType = eINDIC_ClusterTime Or .DisplayType = eINDIC_ClusterPriceNone Then
'            'do nothing - display type can only be none for chart's editor purposes
'        ElseIf .DisplayType = eINDIC_ClusterPrice Then
'            If cboType.Text = "Line" Then .DisplayType = eINDIC_ClusterPriceLine
'        ElseIf .DisplayType = eINDIC_ClusterPriceLine Then
'            If cboType.Text <> "Line" Then .DisplayType = eINDIC_ClusterPrice
        If cboType.Text = "None" Then
            .DisplayType = eINDIC_NoStyle
        ElseIf cboType.Text = "Value Markers" Then
            .DisplayType = eINDIC_ValueMarkers
        Else
            .DisplayType = cboType.ListIndex
            eNewDisplay = .DisplayType
            
            If ePrevDisplay = eINDIC_Ribbon And eNewDisplay <> eINDIC_Ribbon Then
                 RibbonList Nothing, m.Chart, m.EditedIndicator, 2, 0       '6227
            ElseIf eNewDisplay = eINDIC_Ribbon Then
                Set Ind = Nothing
                
                Set aInd = New cGdArray
                RibbonList aInd, m.Chart, m.EditedIndicator, 1, 0
                If aInd.Size > 0 Then
                    Set Ind = frmRibbonInd.ShowMe(m.Chart, aInd)
                End If
                
                If Ind Is Nothing Then
                    m.EditedIndicator.DisplayType = ePrevDisplay
                ElseIf Not Ind.Display Or Not m.EditedIndicator.Display Then
                    'make sure both indicators are visible
                    If Not m.Chart Is Nothing Then Set Tree = m.Chart.Tree
                    
                    If Not Tree Is Nothing Then
                        'walk through grid and turn on the checkbox(es)
                        With fgSettings
                            For i = .TopRow To .BottomRow
                                If Not .RowHidden(i) Then
                                    If Ind Is Tree(.RowData(i)) Then
                                        If Not Ind.Display Then
                                            Ind.Display = True
                                            fgSettings.Cell(flexcpChecked, i, kBoxCol) = flexChecked
                                            fgSettings_AfterEdit i, kBoxCol
                                        End If
                                        k = k + 1
                                    ElseIf m.EditedIndicator Is Tree(.RowData(i)) Then
                                        If Not m.EditedIndicator.Display Then
                                            m.EditedIndicator.Display = True
                                            fgSettings.Cell(flexcpChecked, i, kBoxCol) = flexChecked
                                            fgSettings_AfterEdit i, kBoxCol
                                        End If
                                        k = k + 1
                                    End If
                                End If
                                If k = 2 Then Exit For
                            Next
                        End With
                    End If
                End If
                
            End If
        
        End If
    End With
    
    FixControls
    SyncPowerZones
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboType.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboVertGrid_Click()
On Error GoTo ErrSection:
    
    Dim s$, nIndex&, i&
    Dim frm As Form, frms As New cForms

    If fgSettings.Redraw = flexRDNone Then Exit Sub
        
    If txtPoints.Visible Then
        If Not m.EditedIndicator Is Nothing Then
            If m.Chart.Bars.IsIntraday Then
                If cboVertGrid.Text = "Monthly" Then
                    If ValOfText(txtBarWidth.Text) > 1 Then
                        txtBarWidth.Text = "1"
                    End If
                ElseIf cboVertGrid.Text = "Weekly" Then
                    If ValOfText(txtBarWidth.Text) > 4 Then
                        txtBarWidth.Text = "4"
                    End If
                ElseIf cboVertGrid.Text = "Daily" Then
                    If ValOfText(txtBarWidth.Text) > 30 Then
                        txtBarWidth.Text = "30"
                    End If
                End If
            End If
            i = GetPeriodicity(txtBarWidth.Text & " " & cboVertGrid.Text)
            If i < m.Chart.Periodicity Then GoTo ErrExit
            m.EditedIndicator.ProfilePeriodicityStr = txtBarWidth.Text & " " & cboVertGrid.Text
            m.iGenerateChart = True
        End If
    Else
        nIndex = cboVertGrid.ListIndex
        If m.Chart.VertGrid <> nIndex Then
            m.Chart.VertGrid = nIndex
            
            s = "Change the Vertical Gridlines for all the|charts in this chart page, or just for this chart?"
            If InfBox(s, "?", "All Charts|+-This Chart", "Horizontal Gridlines") = "A" Then
                frms.Init
                Do
                    Set frm = frms.NextForm
                    If frm Is Nothing Then Exit Do
                    ' then simply use the "frm" returned by .NextForm
                    If IsFrmChart(frm) Then
                        frm.Chart.VertGrid = nIndex
                    End If
                Loop
                Set frm = Nothing
                m.iGenerateChart = 1 ' to reset all charts
            Else
                m.iGenerateChart = True ' reset just this chart
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboVertGrid_Change", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkArtChartLabel_Click()
On Error GoTo ErrSection:

    If Not m.EditedIndicator Is Nothing Then
        If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
            If chkVolume_VA(1).Value = 0 Then
                m.EditedIndicator.ProfileShowHide(ePCStruct_Volume_POC) = chkArtChartLabel.Value
            ElseIf chkArtChartLabel.Value = 0 Then
                m.EditedIndicator.ProfileShowHide(ePCStruct_Volume_POC) = 3
            Else
                m.EditedIndicator.ProfileShowHide(ePCStruct_Volume_POC) = 2
            End If
            FixProfileChkboxes m.EditedIndicator
        Else
            If chkArtChartLabel.Value = vbChecked Then
                m.EditedIndicator.IndLabelMode = eINDIC_NoValue
            Else
                m.EditedIndicator.IndLabelMode = eINDIC_Nothing
            End If
            FixCtlArrayData m.EditedIndicator
        End If
        m.iGenerateChart = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkArtChartLabel_Click", eGDRaiseError_Show

End Sub

Private Sub chkBiColorBars_Click()
On Error GoTo ErrSection:

    If Not m.EditedIndicator Is Nothing Then
        If chkBiColorBars.Value = 1 Then
            m.EditedIndicator.HistogramColorBelow = gdFillColor.Color
        Else
            m.EditedIndicator.HistogramColorBelow = -1
        End If
        FixCtlArrayData m.EditedIndicator
        m.iGenerateChart = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkBiColorBars_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkColorPrice_Click()
On Error GoTo ErrSection:

    If Not m.EditedIndicator Is Nothing Then
        m.EditedIndicator.ColorPriceIndFlag = chkColorPrice.Value
        m.iGenerateChart = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkColorPrice_Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkCustomColors_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    m.Chart.UseCustomColors = chkCustomColors
    ShowChartColors
    m.Chart.FixColors
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkCustomColors.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkEmptyBars_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    m.Chart.ShowEmptyBars = chkEmptyBars * -1
    m.Chart.RedoMode = eRedo9_ReloadData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkEmptyBars.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkFlip_Click()
On Error GoTo ErrSection:

    Dim strMsg$, lRedrawSave&
    
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    
    lRedrawSave = fgSettings.Redraw
    fgSettings.Redraw = flexRDNone      'to prevent the overlay checkbox prompt
    
    If m.EditedIndicator.isPriceInd = 1 Then
        ' TLB 6/29/2016: re-using this control for new "Hide Wick" option
        m.EditedIndicator.HideWick = chkFlip.Value
    ElseIf chkFlip.Value = 1 And Not m.EditedIndicator.Overlayed Then
        strMsg = "This feature requires " & m.EditedIndicator.Name & " to be overlayed." & vbCrLf
        strMsg = strMsg & "Overlay and flip the data for this symbol?"
        If InfBox(strMsg, "?", "+Yes|-No", "Flip comparison symbol") = "Y" Then
            m.EditedIndicator.Overlayed = True
            m.EditedIndicator.Flip = True
            chkOverlayed(0).Value = 1
            chkFlip.Value = 1
            'cannot use module level flag because there can be multiple comparison symbols
            m.EditedIndicator.SyncOverlayFlip = True        '6519
        Else
            m.EditedIndicator.Flip = False
            chkFlip.Value = 0
            GoTo ErrExit
        End If
    ElseIf chkFlip.Value = 0 And m.EditedIndicator.SyncOverlayFlip Then
        m.EditedIndicator.Overlayed = False
        m.EditedIndicator.Flip = False
        chkOverlayed(0).Value = 0
    Else
        m.EditedIndicator.Flip = chkFlip.Value
    End If
    
    m.iGenerateChart = 1

ErrExit:
    fgSettings.Redraw = lRedrawSave
    Exit Sub
    
ErrSection:
    fgSettings.Redraw = lRedrawSave
    RaiseError "frmChartCfg.chkFlip_Click", eGDRaiseError_Show

End Sub

Private Sub chkGradient_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    
    If m.Chart.UseCustomColors Then
        m.Chart.UseGradient = chkGradient.Value
        
        ' do special gradients for Woodies templates
        If m.Chart.UseGradient Then
            If m.Chart.ChartBackColor = 8362872 Then ' older Woodies color
                m.Chart.ChartBackColor = RGB(171, 192, 176)
                m.Chart.ChartGradientColor = RGB(84, 112, 89)
                clrChartBack.Color = m.Chart.ChartBackColor
                clrChartGradient.Color = m.Chart.ChartGradientColor
            ElseIf m.Chart.ChartBackColor = 8421440 Then ' newer WCCI color
                m.Chart.ChartBackColor = RGB(80, 160, 160)
                m.Chart.ChartGradientColor = RGB(54, 108, 108)
                clrChartBack.Color = m.Chart.ChartBackColor
                clrChartGradient.Color = m.Chart.ChartGradientColor
            End If
        ElseIf m.Chart.ChartBackColor = RGB(171, 192, 176) Then
            m.Chart.ChartBackColor = 8362872
            clrChartBack.Color = m.Chart.ChartBackColor
        ElseIf m.Chart.ChartBackColor = RGB(80, 160, 160) Then
            m.Chart.ChartBackColor = 8421440
            clrChartBack.Color = m.Chart.ChartBackColor
        End If
    Else
        g.ChartGlobals.nUseGradient = chkGradient.Value
    End If
    
    m.iGenerateChart = 1

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkGradient_Click", eGDRaiseError_Show

End Sub

Private Sub chkHideSeparator_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedPane Is Nothing Then Exit Sub
    
    m.EditedPane.HideSeparator = chkHideSeparator
    'm.iGenerateChart = True
    m.Chart.Form.pbChart.AutoRedraw = False
    geDrawSeparator m.Chart.geChartObj, m.Chart.Form.pbChart.hDC, m.EditedPane.gePaneId, 1, chkHideSeparator
    m.Chart.Form.pbChart.AutoRedraw = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkHideSeparator.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkHorzGrid_Click()
On Error GoTo ErrSection:

    Dim s$
    Dim frm As Form, frms As New cForms

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    
    m.Chart.HorzGrid = chkHorzGrid * -1
    s = "Turn the Horizontal Gridlines OFF for all the|charts in this chart page, or just for this chart?"
    If m.Chart.HorzGrid Then
        s = Replace(s, " OFF ", " ON ")
    End If
    If InfBox(s, "?", "All Charts|+-This Chart", "Horizontal Gridlines") = "A" Then
        frms.Init
        Do
            Set frm = frms.NextForm
            If frm Is Nothing Then Exit Do
            ' then simply use the "frm" returned by .NextForm
            If IsFrmChart(frm) Then
                frm.Chart.HorzGrid = m.Chart.HorzGrid
            End If
        Loop
        Set frm = Nothing
        m.iGenerateChart = 1 ' to reset all charts
    Else
        m.iGenerateChart = True ' reset just this chart
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkHorzGrid.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkHorzLetters_Click()
On Error GoTo ErrSection:

    Dim bShow As Boolean

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
        If chkVolume_VA(0).Value = 0 Then
            m.EditedIndicator.ProfileShowHide(ePCStruct_Volume_VA) = chkHorzLetters.Value
        ElseIf chkHorzLetters.Value = 0 Then
            m.EditedIndicator.ProfileShowHide(ePCStruct_Volume_VA) = 3
        Else
            m.EditedIndicator.ProfileShowHide(ePCStruct_Volume_VA) = 2
        End If
        
        If m.EditedIndicator.ProfileShowHide(ePCStruct_Volume_VA) > 0 Then bShow = True
        txtPercentVolume_VA.Enabled = bShow
        lblArtLabels.Enabled = bShow
        FixProfileChkboxes m.EditedIndicator
    Else
        m.EditedIndicator.BoxFillStyle = chkHorzLetters.Value       'art reversal bars
    End If
    
    m.iGenerateChart = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkHorzLetters_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkOverlayed_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim strMsg$

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    m.EditedIndicator.Overlayed = chkOverlayed(Index)
    
    If chkOverlayed(Index).Value <> 0 And Me.Visible Then
        strMsg = "Since an overlayed indicator is not fixed |to the pane's scale, it can 'float' up |or down as the chart is scrolled." _
            & "||Therefore, the points where overlayed indicators cross other indicators or lines are not meaningful."
        InfBox strMsg, "i", , "Please Note ..."
        m.EditedIndicator.SyncOverlayFlip = False       '6519
    End If
    
    If Not m.EditedIndicator.Overlayed Then
        If m.EditedIndicator.Flip Then
            chkFlip.Value = 0
        End If
    End If
    
    m.iGenerateChart = 1
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkOverlayed.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub chkPriceTopMost_Click()

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    
    m.Chart.PriceTopMost = chkPriceTopMost.Value
    m.iGenerateChart = True

End Sub

Private Sub chkRemoveOvernightGap_Click()
    
    If fgSettings.Redraw = flexRDNone Then Exit Sub

    If Not m.Chart Is Nothing Then
        If chkRemoveOvernightGap.Value = vbChecked Then
            m.Chart.RemoveOvernightGap = True
        Else
            m.Chart.RemoveOvernightGap = False
        End If
        m.Chart.RedoMode = eRedo9_ReloadData
    End If

End Sub

Private Sub chkShowInAllPanes_Click()

    If Not m.EditedIndicator Is Nothing Then
        If chkShowInAllPanes.Value = 1 Then
            m.EditedIndicator.ShowInAllPanes = 2
        Else
            m.EditedIndicator.ShowInAllPanes = 1
        End If
        m.iGenerateChart = 1
    End If

End Sub

Private Sub chkSplitPane_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
        
    If Not m.Chart Is Nothing Then
        m.Chart.ShowSplitPane = chkSplitPane.Value
        cmdSplitPaneCfg.Enabled = chkSplitPane.Value
        m.iGenerateChart = True
    End If
        
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.chkSplitPane_Click"
    
End Sub

Private Sub chkSplitsRolls_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    g.ChartGlobals.bSplitsRolls = -chkSplitsRolls
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkSplitRolls.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkTips_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    g.ChartGlobals.bFloatingTips = -chkTips
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkTips.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkTrendHistory_Click()
    If Not m.EditedIndicator Is Nothing Then
        m.EditedIndicator.ShowTrendHistory = chkTrendHistory.Value
        m.iGenerateChart = True
    End If
End Sub

Private Sub chkTrueRange_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    If chkTrueRange.Value = 0 Then
        m.EditedIndicator.TrueRangeFlag = kNullData
    Else
        m.EditedIndicator.TrueRangeFlag = m.Chart.FirstTrueRangeClose
    End If
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkTrueRange.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkUnsplit_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    m.Chart.Unsplit = chkUnsplit * -1
    m.Chart.RedoMode = eRedo9_ReloadData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkUnsplit.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkUpDownColors_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    m.EditedIndicator.UpDownColorFlag = chkUpDownColors.Value
    FixCtlBarData m.EditedIndicator
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkUpDownColors.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkVolume_VA_Click(Index As Integer)
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub

    If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
        Select Case Index
            Case 0
                chkHorzLetters_Click
            Case 1
                chkArtChartLabel_Click
            Case 2, 3
                If chkVolume_VA(3).Value = vbUnchecked Then
                    'show previous values is OFF - set to whatever show currrent values checkbox is
                    m.EditedIndicator.ProfileShowHide(ePCStruct_Open) = chkVolume_VA(2).Value
                ElseIf chkVolume_VA(2).Value = vbUnchecked Then
                    'show current values is OFF
                    m.EditedIndicator.ProfileShowHide(ePCStruct_Open) = 3
                Else
                    'both check boxes are ON
                    m.EditedIndicator.ProfileShowHide(ePCStruct_Open) = 2
                End If
            Case 4
                m.EditedIndicator.ProfileShowHide(ePCStruct_Close) = chkVolume_VA(4).Value
        End Select
        m.iGenerateChart = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.chkVolume_VA_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkYscaleLabelAll_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedPane Is Nothing Then Exit Sub
    
    m.EditedPane.YscaleLabelAll = chkYscaleLabelAll
    m.Chart.Form.pbChart.AutoRedraw = False
    m.Chart.geDrawChart
    m.Chart.Form.pbChart.AutoRedraw = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkYscaleLabelAll.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub clrBorderBack_Changed()
On Error GoTo ErrSection:

    If m.Chart.UseCustomColors Then
        m.Chart.BorderBackColor = clrBorderBack.Color
    Else
        g.ChartGlobals.nBorderBackColor = clrBorderBack.Color
    End If
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrBorderBack.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub clrBorderFore_Changed()
On Error GoTo ErrSection:

    If m.Chart.UseCustomColors Then
        m.Chart.BorderForeColor = clrBorderFore.Color
    Else
        g.ChartGlobals.nBorderForeColor = clrBorderFore.Color
    End If
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrBorderFore.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub clrChartBack_Changed()
On Error GoTo ErrSection:

    Dim i&

    If m.Chart.UseCustomColors Then
        m.Chart.ChartBackColor = clrChartBack.Color
        m.Chart.FixColors
        m.iGenerateChart = True
    Else
        g.ChartGlobals.nChartBackColor = clrChartBack.Color
        For i = 0 To Forms.Count - 1
            If IsFrmChart(Forms(i)) Then
                Forms(i).Chart.FixColors
            End If
        Next
        m.iGenerateChart = 1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrChartBack.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub clrChartFore_Changed()
On Error GoTo ErrSection:

    If m.Chart.UseCustomColors Then
        m.Chart.ChartForeColor = clrChartFore.Color
    Else
        g.ChartGlobals.nChartForeColor = clrChartFore.Color
    End If
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrChartFore.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub clrChartGradient_Changed()
On Error GoTo ErrSection:

    If m.Chart.UseCustomColors Then
        m.Chart.ChartGradientColor = clrChartGradient.Color
    Else
        g.ChartGlobals.nChartGradientColor = clrChartGradient.Color
    End If
    m.iGenerateChart = 1

ErrExit:
    Exit Sub

ErrSection:
    RaiseError

End Sub

Private Sub clrFakeDropdown_Changed()
On Error GoTo ErrSection:

    Dim i&, iColor As Long
    Dim strColors As String

    Dim nColor As Long
    
    clrFakeDropdown.Visible = False
    iColor = clrFakeDropdown.Color
    If iColor = 0 Then iColor = -1  '0 is reserved color in flex grid control
    
    With fgColors
        If .Row >= 0 And .Row < .Rows Then
            If .Cell(flexcpBackColor, .Row, 1) <> iColor Then
                .Cell(flexcpBackColor, .Row, 1) = iColor
                If Not m.EditedIndicator Is Nothing Then
                    For i = 0 To .Rows - 1
                        strColors = strColors & Str(.Cell(flexcpBackColor, i, 1)) & ";"
                        m.EditedIndicator.HawkeyeLevelsColors = strColors
                    Next
                End If
            End If
            .Select .Row, 0
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrFakeDropdown_Changed", eGDRaiseError_Show

End Sub

Private Sub clrLong_Changed()
On Error GoTo ErrSection:

    g.ChartGlobals.nLongColor = clrLong.Color
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrLong.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub clrShort_Changed()
On Error GoTo ErrSection:

    g.ChartGlobals.nShortColor = clrShort.Color
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrShort.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub clrSplitsRolls_Changed()
On Error GoTo ErrSection:

    g.ChartGlobals.nSplitRollColor = clrSplitsRolls.Color
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrSplitsRolls.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    ShowAddPopup True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdAdd.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdBarPeriod_Click()
On Error GoTo ErrSection:

    Dim i&
    
    If m.Chart.ChangeBarPeriod("<", False) Then
        cboBarPeriod.Text = GetPeriodStr(m.Chart.Periodicity)
        FixControls
        m.iGenerateChart = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdBarPeriod.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    'must do this with a default button so the
    'Lost_Focus event will trigger for the active control
    MoveFocus cmdCancel
    DoEvents
    
    If Not m.Chart Is Nothing Then
        If m.Chart.TypeOfChart <> eTypeChart_Seasonal Then
            If m.Chart.Form.OrderBarMode = eOrdBarMode_PFP Then
                Unload Me           '5897
                Exit Sub
            ElseIf Not m.Chart.Form.IsInGameMode Then
                ' restore chart to previous settings if not in game mode
                If m.Chart.Zoomed = False Then
                    m.Chart.TemplateLoad
                    m.Chart.geResetPanes 'MJM - added for grapheng.dll
                End If
            End If
        End If
    End If
    
    cmdOK_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdEditFunction_Click()
On Error GoTo ErrSection:

    Dim nID&
    Dim frm As Form
    Dim eType As eIndicatorDisplayType
    
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    eType = m.EditedIndicator.DisplayType
    
'JM 07-30-2015: think all this cluster code is obsolete (leave awhile then remove if all ok)
'    If m.EditedIndicator.IsClusterPrice() Or eType = eINDIC_ClusterTime Then
'        Me.Hide
'        frmClusterCfg.ShowMe m.Chart, m.EditedIndicator
'        m.bClickOK = True
'        Exit Sub
'    End If
    
    With m.EditedIndicator
        nID = .FunctionID
        If nID = 0 Then
            .EditCustom m.Chart             '6499
            If Len(Trim(.CodedText & .Expression)) = 0 Then
                .Display = False
                .Expression = ""
            End If
            lblCustomFunction = .Expression
            FixControls
            .CodedText = "" '(so will rebuild in context)
        Else
            Select Case g.Functions.Item(Str(nID)).ImplementationTypeID
                Case 1
                    If Not ActivateEditor("frmFunctionMgr", nID) Then
                        Set frm = New frmFunctionMgr
                        frm.ShowMe nID, m.Chart                         '6499
                    End If
                Case 2
                    If Not ActivateEditor("frmFunctionMgrCT", nID) Then
                        Set frm = New frmFunctionMgrCT
                        frm.ShowMe nID, , , , , Me                      '6499
                    End If
                Case Else
                    Beep
            End Select
        End If
    End With
   
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdEditFunction.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdEditSystem_Click()
On Error GoTo ErrSection:

    Dim nID As Long
    Dim nLinkToChart As Long
    Dim frm As frmSystemManager
    
    nID = m.Chart.SystemID
    If nID = 0 Then
        Beep
    ElseIf Not ActivateEditor("frmSystemManager", nID) Then
        Set frm = New frmSystemManager
        nLinkToChart = GetIniFileProperty("LinkToChart", 1, "Systems", g.strIniFile)
        frm.ShowMe nID, , False, "", , True, m.Chart                '6499
        If nLinkToChart = 1 Then frm.UseChartSystem 1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdEditSystem.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdFakeDropdown_Click()
On Error GoTo ErrSection:

    If fgColors.Visible Then
        fgColors.Visible = False
        clrFakeDropdown.Visible = False
    Else
        SetColorsGrid
        fgColors.Move fraFakeDropdown.Left, fraFakeDropdown.Top + fraFakeDropdown.Height
        fgColors.Visible = True
        fgColors.ZOrder
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdFakeDropdown_Click", eGDRaiseError_Show

End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:
        
    Dim n&, nStyle&, strFont$
    
    Me.Font.Name = g.ChartGlobals.strFontName
    Me.Font.Size = g.ChartGlobals.nFontSize
    
    'JM 10-26-2015: Prior to dark, light themes, indicator label style was always hard coded to
    '   be bold. Glen wanted these labels to be un-bolded so need to save style in case existing
    '   users prefer bolded indicator labels. Italic remains unimplemented as per Tim.
    '   Bold, italic remain unimplemented for labels in y-scale (i.e. unchanged)
    nStyle = g.ChartGlobals.nFontStyle
    Select Case nStyle
        Case 1, 3:
            Me.Font.Bold = True
        Case Else:
            Me.Font.Italic = False
            Me.Font.Bold = False
    End Select
    
    'use control on this form so will show centered on this form - 6499
    If CommonDialogFont(CommonDialog1, Me.Font) Then
        'do not implement bold or italic per Tim
        n = Me.Font.Size
        If n < 4 Or n > 30 Then n = 8
        strFont = Me.Font.Name
        
        If Me.Font.Bold = True Then
            nStyle = 1
        Else
            nStyle = 0
        End If
        
        If n <> g.ChartGlobals.nFontSize Or nStyle <> g.ChartGlobals.nFontStyle Or strFont <> g.ChartGlobals.strFontName Then
            g.ChartGlobals.nFontSize = n
            g.ChartGlobals.nFontStyle = nStyle
            g.ChartGlobals.strFontName = strFont
            m.iGenerateChart = 1
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdFont_Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdMarketInfo_Click()
On Error GoTo ErrSection:

    ' TLB 3/17/2014: allow for changing mkt info for custom symbol
    If Left(m.Chart.Symbol, 1) = "*" Then
        If frmMarkets.ShowMe(m.Chart.Symbol, , m.Chart) Then
            ' to reload data (in case min move info changed)
            m.Chart.RedoMode = eRedo9_ReloadData
            m.iGenerateChart = True
        End If
    Else
        frmMarkets.ShowMe lblSymbol.Caption, , m.Chart       '5216, 6499
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdMarketInfo.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub cmdPyramidFont_Click()
On Error GoTo ErrSection:
        
    Dim nSize&, strFont$
    
    Dim bBold As Boolean
    Dim bItalic As Boolean
    
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    Me.Font.Name = m.EditedIndicator.FontName
    Me.Font.Size = m.EditedIndicator.FontSize
    Me.Font.Bold = m.EditedIndicator.FontBold
    Me.Font.Italic = m.EditedIndicator.FontItalic
    
    If CommonDialogFont(CommonDialog1, Me.Font) Then
        strFont = Me.Font.Name
        bBold = Me.Font.Bold
        bItalic = Me.Font.Italic
        
        nSize = Me.Font.Size
        If nSize < 4 Or nSize > 30 Then nSize = 8
        
        If nSize <> m.EditedIndicator.FontSize Then
            m.EditedIndicator.FontSize = nSize
            m.iGenerateChart = 1
        End If
        
        If bBold <> m.EditedIndicator.FontBold Then
            m.EditedIndicator.FontBold = bBold
            m.iGenerateChart = 1
        End If
        
        If bItalic <> m.EditedIndicator.FontItalic Then
            m.EditedIndicator.FontItalic = bItalic
            m.iGenerateChart = 1
        End If
        
        If strFont <> m.EditedIndicator.FontName Then
            m.EditedIndicator.FontName = strFont
            m.iGenerateChart = 1
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdPyramidFont_Click", eGDRaiseError_Show

End Sub

Private Sub cmdShowBidAsk_Click()
    frmChartBidAskCfg.ShowMe m.Chart
End Sub

Private Sub cmdSplitPaneCfg_Click()
On Error GoTo ErrSection:

    If Not m.Chart Is Nothing Then
        If Not m.Chart.Tree Is Nothing Then
            If m.PaneWood Is Nothing Then
                frmSplitPaneCfg.ShowMe m.Chart, m.Chart.Tree.Index("PRICE PANE"), 0
            ElseIf InfBox("Which pane would you like to configure?", "?", "+Price|-Woodies", "Select Pane") = "W" Then
                frmSplitPaneCfg.ShowMe m.Chart, m.PaneWood.gePaneId, 0
            Else
                frmSplitPaneCfg.ShowMe m.Chart, m.Chart.Tree.Index("PRICE PANE"), 0
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.cmdSplitPaneCfg_Click"

End Sub

Private Sub cmdStartStop_Click()
On Error GoTo ErrSection:
    
    If Not m.Chart Is Nothing Then
        If Not m.Chart.Bars Is Nothing Then
            If frmStartStopTimes.ShowMe(m.Chart.Bars, m.Chart) = True Then          '6499
                'm.bStartEndTimesChanged = True
                SetStartStopTimesText
                m.Chart.GenerateChart eRedo9_ReloadData         '5307
                Unload Me
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.cmdStartStop_Click"

End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    Dim idx&, idxRemove&, iLevel&, n&, nRow&, strKey$
    Dim bRemoved As Boolean, bPrice As Boolean
    
    'check for zoom
    If m.Chart.Zoomed Then m.Chart.UnzoomChart True
    
    nRow = fgSettings.Row
    
    ' see if removing system
    If RowType(nRow) = kSystem Then
        m.Chart.ShowTrades = False
        m.Chart.SystemID = 0
        LoadSystemInfo
        'fgSettings.Row = nRow - 1
        Exit Sub
    End If
    
    ' make sure Price is not being removed
    idx = fgSettings.RowData(nRow)
    If idx > 0 Then
        strKey = m.Chart.Tree.Key(idx)
        If UCase(Left(strKey, 5)) = "PRICE" Then
            bPrice = True
        ElseIf m.Chart.Tree.NodeLevel(strKey) > 0 Then
            bRemoved = RemoveIndicator(nRow, strKey)
        Else
            bRemoved = m.Chart.Tree.Remove(idx)
        End If
    End If

    If bRemoved Then
        LoadSettingsGrid
    Else
        Beep ' can't remove it
        If bPrice Then
            InfBox "'Price' cannot be removed.", "!"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdRemove.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrSection:

    vseSave_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdSave.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdMoveUp_Click()
On Error GoTo ErrSection:

    Dim idx&, idxNew&, idxRelative&
    
    idx = fgSettings.RowData(fgSettings.Row)
    idxNew = idx
    With m.Chart.Tree
        If idx > 0 Then
            idxRelative = .RelativeIndex(idx, eTREE_PrevSibling)
            If idxRelative > 0 Then
                idxNew = .Move(idx, idxRelative, eTREE_PrevSibling)
            ElseIf .NodeLevel(idx) = 1 And .Key(idx) <> "PRICE" Then
                ' only allow unlinked indicators to
                ' go to previous parent (Pane)
                idxRelative = .RelativeIndex(idx, eTREE_Parent)
                Do
                    idxRelative = .RelativeIndex(idxRelative, eTREE_PrevSibling)
                    If idxRelative = 0 Then Exit Do
                    ' but skip over non-displayed Panes
                Loop While Not .Item(idxRelative).Display
                If idxRelative > 0 Then
                    idxNew = .Move(idx, idxRelative, eTREE_LastChild)
                End If
            End If
        End If
    End With
    If idxNew = idx Then
        Beep 'invalid move
    Else
        LoadSettingsGrid
        SelectIdx idxNew
        ClearCodedText idxNew
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdMoveUp.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdMoveDown_Click()
On Error GoTo ErrSection:

    Dim idx&, idxNew&, idxRelative&
    
    idx = fgSettings.RowData(fgSettings.Row)
    idxNew = idx
    With m.Chart.Tree
        If idx > 0 Then
            idxRelative = .RelativeIndex(idx, eTREE_NextSibling)
            If idxRelative > 0 Then
                idxNew = .Move(idx, idxRelative, eTREE_NextSibling)
            ElseIf .NodeLevel(idx) = 1 And .Key(idx) <> "PRICE" Then
                ' only allow unlinked indicators to
                ' go to next parent (Pane)
                idxRelative = .RelativeIndex(idx, eTREE_Parent)
                Do
                    idxRelative = .RelativeIndex(idxRelative, eTREE_NextSibling)
                    If idxRelative = 0 Then Exit Do
                    ' but skip over non-displayed Panes
                Loop While Not .Item(idxRelative).Display
                If idxRelative > 0 Then
                    idxNew = .Move(idx, idxRelative, eTREE_FirstChild)
                End If
            End If
        End If
    End With
    If idxNew = idx Then
        Beep 'invalid move
    Else
        LoadSettingsGrid
        SelectIdx idxNew
        ClearCodedText idxNew
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdMoveDown.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim i&, s$
    
    tmrChartCfg.Enabled = False
    
    'must do this with a default button so the
    'Lost_Focus event will trigger for the active control
    MoveFocus cmdOK
    DoEvents
    
    SaveModifiedInputs
    
    If Not m.Chart Is Nothing Then
        If Not m.Chart.VerifyPanes(True) Then Exit Sub
        If Not m.Chart.Form.IsInGameMode Then
            With cboAccounts
                If .ListIndex >= 0 Then
                    m.Chart.TradeAccountID = .ItemData(.ListIndex)
                Else
                    m.Chart.TradeAccountID = 0
                End If
            End With
            Chart.Form.PfpReset ePfpReset_GridPfp
            Chart.Form.PfpReset ePfpReset_PpfAnnotInd

'JM 08-02-2011 - this is done in cChart.cls, not needed here - aardvark 6410 (leave awhile then remove if all ok)
'            Chart.Form.PfpReset ePfpReset_GridInd, True
        
        End If
    End If
    
    'check that min is < max if scale is in manual mode
    If Not m.EditedPane Is Nothing Then
        If m.EditedPane.Scaling = ePANE_ScaleModeManual Then
            If m.EditedPane.Min >= m.EditedPane.Max Then
                MsgBox "The Max value must be greater than the Min value. Reverting to auto scale."
                optAutoScale.Value = True
            End If
        End If
    End If
        
    m.nRowWhenLeftForm = fgSettings.Row

    Set m.EditedIndicator = Nothing
    Set m.EditedPane = Nothing
              
    If Not m.Chart Is Nothing Then
        Screen.MousePointer = vbHourglass
        If m.iGenerateChart = 1 Then
            ' do for all charts (if globals have been edited)
            UpdateVisibleCharts eRedo3_Settings
        Else
            ' just this chart
            m.Chart.geResetPanes        'aardvark 4139
            m.Chart.GenerateChart eRedo3_Settings
            
            If g.RealTime.Active Then
                m.Chart.geForceRecalc   '5364
                m.Chart.GenerateChart eRedo1_Scrolled
            End If
        End If
        m.Chart.SyncToolbar True, True
    End If
    m.iGenerateChart = False '(cleared)
    
    'Me.Visible = False
    'DockState(Me) = eHidden
    
    Dim bDontCare As Boolean
    
    If Not ActiveChart Is Nothing Then
        MoveFocus ActiveChart.pbChart   'ActiveChart.Peg    -MJM
        If ActiveChart.OrderBarMode = eOrdBarMode_Wizard Then
            ActiveChart.Chart.ChainFromFile bDontCare, True         '5228
            ActiveChart.Chart.GenerateChart eRedo1_Scrolled
        Else
            ActiveChart.FixOrderBarControls
        End If
    End If
    
    s = GetIniFileProperty("TradenavTheme", "", "General", g.strIniFile)
    If Len(s) > 0 Then SaveChartGlobals s
    
    'JM 12-16-2009: Do not remember why this call to resize the active chart is here.
    'It is, however, causing the system menu buttons to disappear from a maximized chart and
    'weird tabbed charts issue 5496. For now, safest fix seems to be to comment this out.
'    FormResize ActiveChart
    
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdOK.Click", eGDRaiseError_Show
    Unload Me
    Resume ErrExit
        
End Sub

Private Sub cmdSaveAs_Click()
On Error GoTo ErrSection:

    vseSaveAs_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdSaveAs.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdSaveStudy_Click()
On Error GoTo ErrSection:

    Dim idx&, strName$, strDesc$, strFile$, bSave As Boolean
    Dim Pane As cPane
    Dim Ind As cIndicator

    If cmdSaveStudy.Caption = "&Save Study" And Not HasGold Then Exit Sub

    If RowType(fgSettings.Row) = kPane Then
        idx = fgSettings.RowData(fgSettings.Row)
        If idx > 0 Then
            With m.Chart
                Set Pane = .Tree(idx)
                strName = Pane.Name
                If UCase(strName) = "PRICE" Then
                    strName = ""
                    strDesc = ""
                Else
                    strDesc = Pane.Desc
                End If
                If frmNameDesc.ShowMe(strName, strDesc, , m.Chart) Then
                    strFile = App.Path & "\Charts\Templates\" & strName & ".STU"
                    If Not FileExist(strFile) Then
                        bSave = True
                    ElseIf InfBox("Overwrite existing study:|" & strName, "?", _
                            "+Overwrite|-Cancel", "Confirm Overwrite") = "O" Then
                        bSave = True
                    End If
                    If bSave Then
                        Pane.Name = strName
                        Pane.Desc = strDesc
                        .TemplateSaveStudy .Tree.Key(idx), Pane.Name
                    End If
                End If
                Set Pane = Nothing
            End With
        End If
    ElseIf cmdSaveStudy.Caption = "&Group" Then
        If GroupIndicators Then cmdSaveStudy.Caption = "&Ungroup"
    ElseIf cmdSaveStudy.Caption = "&Ungroup" Then
        idx = fgSettings.RowData(fgSettings.Row)
        Set Ind = m.Chart.Tree(idx)
        If Not Ind Is Nothing Then
            CollapseRows Ind.GroupKey
            cmdSaveStudy.Caption = "&Group"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdSaveStudy.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdSelectSystem_Click()

On Error GoTo ErrSection:

    AddToChart eAdd_System, 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdSelectSystem.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub cmdSymbol_Click()
On Error GoTo ErrSection:

    Dim nRec&, strSymbol$
    Dim aSymbol As cGdArray
    
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
            
    If UCase(m.EditedIndicator.Name) = "PRICE" Then
        If m.Chart Is Nothing Then
            Set aSymbol = frmSymbolSelector.ShowMe(lblSymbol, False, , , True)
        Else
            If Len(m.Chart.SpreadSymbols) > 0 Then
                frmNewChart.ShowMe m.Chart.SpreadSymbols, True, m.Chart
            Else
                Set aSymbol = frmSymbolSelector.ShowMe(lblSymbol, False, , , True)
            End If
        End If
    Else
        Set aSymbol = frmSymbolSelector.ShowMe(lblSymbol, False, , "Comparison Symbol", True)
    End If
    
    If Not aSymbol Is Nothing Then
        If aSymbol.Size > 0 Then
            ' TLB 4/11/2012: allow if an external symbol (from hard drive)
            strSymbol = ""
            If InStr(aSymbol(0), "|") > 0 Then
                strSymbol = aSymbol(0)
                nRec = -1
            Else
                nRec = g.SymbolPool.PoolRecForSymbol(aSymbol(0))
                If nRec >= 0 Then
                    strSymbol = UCase(g.SymbolPool.Symbol(nRec))
                End If
            End If
            
            If Len(strSymbol) > 0 Then
                lblSymbol = Parse(strSymbol, "|", 1)
                lblDesc = g.SymbolPool.Desc(nRec)
                If UCase(m.EditedIndicator.Name) = "PRICE" Then
                    m.Chart.SetSymbol strSymbol
                    m.Chart.RedoMode = eRedo9_ReloadData
                Else
                    m.EditedIndicator.Name = strSymbol
                    fgSettings.TextMatrix(fgSettings.Row, kNameCol) = UCase(lblSymbol)
                End If
                'If SecurityType(lblSymbol.Caption) = "S" And InStr(strSymbol, "|") = 0 Then
                If InStr("SM", SecurityType(lblSymbol.Caption)) > 0 And InStr(strSymbol, "|") = 0 Then
                    chkUnsplit.Visible = True
                    cmdMarketInfo.Visible = False
                ElseIf Len(m.Chart.SpreadSymbols) > 0 Then
                    chkUnsplit.Visible = False
                    cmdMarketInfo.Visible = False
                Else
                    chkUnsplit.Visible = False
                    cmdMarketInfo.Visible = True
                End If
            End If
        End If
    End If
    
ErrExit:
    Set aSymbol = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdTemplate_Click()
On Error GoTo ErrSection:

    Dim i&

    i = fgSettings.Row
    frmTemplates.ShowMe eMode_Templates, m.Chart
    Set Chart = m.Chart
    If i >= fgSettings.FixedRows And i < fgSettings.Rows Then
        fgSettings.Row = i
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cmdTemplate.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdTradeSettings_Click()
On Error GoTo ErrSection:

    frmChartOrdBar.ShowMe m.Chart.Form

    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.cmdTradeSettings_Click"
End Sub

Private Sub dtFromDate_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If dtFromDate <> m.Chart.FromDate Then
        m.Chart.FromDate = dtFromDate
        m.Chart.RedoMode = eRedo9_ReloadData
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.dtFromDate.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub dtToDate_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If dtToDate <> m.Chart.ToDate Then
        m.Chart.ToDate = dtToDate
        m.Chart.RedoMode = eRedo9_ReloadData
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.dtToDate.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    With fgColors
        If .Row >= 0 And .Row < .Rows Then
            If .Col = 1 Then
                .Select .Row, .Col
                clrFakeDropdown.Width = .ColWidth(1) * 1.3
                clrFakeDropdown.Move .Left + .ClientWidth - clrFakeDropdown.Width, .Top + .RowHeight(.Row) * .Row + 1 - 30
                clrFakeDropdown.Color = .Cell(flexcpBackColor, .Row, .Col)
                clrFakeDropdown.Visible = True
                clrFakeDropdown.ZOrder
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgColors_MouseDown", eGDRaiseError_Show

End Sub

Private Sub fgInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim i&, strLinkedValue$, strName$, strValue$, strDefault$
    Dim aSymbols As New cGdArray
        
    m.nEditCol = 0
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    fgInputs.TextMatrix(Row, Col) = fgInputs.EditText
    If fgInputs.TextMatrix(Row, Col) = m.vFgInputsBeforeEditValue Then Exit Sub
       
    If Col = 1 Then
        i = fgInputs.RowData(Row)
        
        strValue = Trim(fgInputs.TextMatrix(Row, Col))
        strDefault = UCase(m.EditedIndicator.ParmDefault(i))
               
        If Left(strValue, 1) = "&" Or strValue = "<Linked Input>" Then
            strName = strValue
            strLinkedValue = HandleLinkedParm(i, strName)
            If Len(strLinkedValue) > 0 Then
                strValue = strLinkedValue
            Else
                strValue = m.EditedIndicator.Parm(i)
            End If
        ElseIf m.EditedIndicator.IsLinkedParm(i) Then
            If Left(m.vFgInputsBeforeEditValue, 1) = "&" Then
                m.EditedIndicator.LinkedParmsDelete m.vFgInputsBeforeEditValue, i
            End If
        End If
        
        If strDefault = "GOLD" Or m.EditedIndicator.ParmType(i) = 5 Then
            If UCase(strValue) = "< LOOKUP >" Then
                Set aSymbols = frmSymbolSelector.ShowMe("", False)
                strValue = aSymbols(0)
                If Len(strValue) = 0 Then
                    strValue = m.EditedIndicator.Parm(i)
                ' if price pane is hidden and symbol is first parm ...
                ElseIf m.Chart.Tree("PRICE PANE").Display = False Then
                    If m.EditedIndicator.ParmType(i) = 5 And i = 1 And _
                            Left(UCase(m.EditedIndicator.Name), 6) = "SPREAD" Then
                        ' ... then link to primary symbol
                        m.Chart.SetSymbol strValue
                        m.Chart.RedoMode = eRedo9_ReloadData
                    End If
                End If
            ElseIf UCase(strValue) = "DEFAULT" Then
                strValue = ""
            Else
                strValue = UCase(strValue)
            End If
            strValue = StripStr(strValue, Chr(34))
            If Right(strValue, 1) = "-" Then strValue = strValue & "067"
            If strValue = "" Then
                strValue = "default"
            ElseIf UCase(strValue) <> "DEFAULT" Then
                strValue = Chr(34) & strValue & Chr(34)
            End If
        End If
        
        If Len(strLinkedValue) > 0 Then
            fgInputs.TextMatrix(Row, Col) = strName
        Else
            fgInputs.TextMatrix(Row, Col) = StripStr(strValue, Chr(34))
        End If
        
        If UCase(m.EditedIndicator.Name) = "POWERZONES" Then
            m.EditedIndicator.Parm(i) = strValue
            SyncPowerZones
            If Not m.Chart Is Nothing Then m.Chart.RedoMode = eRedo3_Settings
            m.iGenerateChart = True
        Else
            ' TLB 9/15/2014: just queue up the modified inputs until the fgSettings grid loses focus
            ' (so won't actually change them until all have been reset -- helps esp. for the TAS indicators)
            'm.EditedIndicator.Parm(i) = strValue
            m.aModifiedInputs.Add Str(m.EditedIndicator.Data.ArrayHandle) & Chr(27) & Str(i) & Chr(27) & strValue
        End If
    End If
    
    ClearCodedText
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgInputs.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgInputs_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    fgInputs.EditCell

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgInputs.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo BeforeEditError:

    Dim strValue$, strDefault$, i&
    
    fgInputs.ComboList = ""
    If m.EditedIndicator Is Nothing Then
        Cancel = True
    Else
        With fgInputs
            If Col = 0 Or .Redraw = flexRDNone Or Me.ActiveControl <> fgInputs Then
                ShowInputTip -1
                Cancel = True
                .Row = 0                '5532
            ElseIf Row >= .FixedRows Then
                i = fgInputs.RowData(Row)
                ShowInputTip Row
                m.nEditCol = 1
                strValue = .TextMatrix(Row, Col)
                m.vFgInputsBeforeEditValue = strValue
                strDefault = UCase(m.EditedIndicator.ParmDefault(i))
                'If UCase(strValue) = "TRUE" Or UCase(strValue) = "FALSE" Then
                If m.EditedIndicator.ParmType(i) = 6 Then
                    If Left(strValue, 1) = "&" Then
                        .ComboList = "False|True|" & strValue & "|<Linked Input>"
                    Else
                        .ComboList = "False|True|<Linked Input>"
                    End If
                ElseIf m.EditedIndicator.ParmType(i) = 5 Or strDefault = "GOLD" Then
                    .ComboList = "|default|GC-067|TQ-067|US-067|TY-067|DX-067|SP-067|$DJIA|< Lookup >"
                ElseIf ParmDefaultIsBarsArray(strDefault) Then  ' = "CLOSE" Then ' And i = 1 Then
                    .ComboList = "Close|High|Low|Open|AvgHL|AvgHLC|AvgOHLC|AvgOC|WClose"
                Else
                    .ComboList = ""
                End If
            Else
                Cancel = True           '5531
            End If
        End With
    End If
    
    Exit Sub

BeforeEditError:
    Cancel = True
    Exit Sub

End Sub

Private Sub fgInputs_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If fgInputs.Redraw = flexRDNone Then Exit Sub
    If NewCol = 0 Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgInputs.BeforeRowColChange", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgInputs_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    GridScrollCheck fgInputs, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel

End Sub

Private Sub fgInputs_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If fgInputs.Redraw = flexRDNone Then Exit Sub
    If NewColSel = 0 Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgInputs.BeforeSelChange", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgInputs_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

    ' so "Lookup" will trigger immediately
    FinishEdit = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgInputs.ComboCloseUp", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgInputs_GotFocus()
On Error GoTo ErrSection:

    Set m.aModifiedInputs = New cGdArray
    fgInputs.EditCell

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgInputs.GotFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgInputs_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
On Error GoTo ErrSection:

    Dim strKey$, i&, strSymbol$
    Dim aStrings As New cGdArray
    
    If Not m.EditedIndicator Is Nothing Then
        With fgInputs
            strKey = Chr(KeyAscii)
'StatusMsg strKey & " " & .EditText & " " & Str(.EditSelStart) & " " & Str(.EditSelLength) & " " & .EditSelText
            If Col = 0 Or .Redraw = flexRDNone Or Me.ActiveControl <> fgInputs Then
                'ShowInputTip -1
                'Cancel = True
            ElseIf IsAlpha(strKey) Or InStr("$#", strKey) > 0 Then
                ' go right to symbol selector if entire text is being replaced
                If Len(.EditText) = .EditSelLength Then
                    i = fgInputs.RowData(Row)
                    If m.EditedIndicator.ParmType(i) = 5 Or m.EditedIndicator.ParmDefault(i) = "GOLD" Then
                        KeyAscii = 0
                        Set aStrings = frmSymbolSelector.ShowMe(strKey, False, , , , False)
                        strSymbol = aStrings(0)
                        If Len(strSymbol) > 0 Then
                            m.EditedIndicator.Parm(i) = Chr(34) & strSymbol & Chr(34)
                            .TextMatrix(Row, Col) = strSymbol
                            ' if price pane is hidden and symbol is first parm ...
                            If m.Chart.Tree("PRICE PANE").Display = False Then
                                If m.EditedIndicator.ParmType(i) = 5 And i = 1 And _
                                        Left(UCase(m.EditedIndicator.Name), 6) = "SPREAD" Then
                                    ' ... then link to primary symbol
                                    m.Chart.SetSymbol strSymbol
                                    m.Chart.RedoMode = eRedo9_ReloadData
                                End If
                            End If
                        End If
                        Set aStrings = Nothing
                    End If
                End If
            End If
        End With
    End If
    Exit Sub

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgInputs_KeyPressEdit", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub fgInputs_LostFocus()
    SaveModifiedInputs
    ShowInputTip -1
End Sub

Private Sub fgInputs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m.nEditCol = 0 Then ShowInputTip -1      'can show tip if desired: ShowInputTip fgInputs.MouseRow + 1
End Sub

Private Sub fgLinkedInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Dim i&, strName$, strValue$, strType$
    Dim bChanged As Boolean

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    
    With fgLinkedInputs
        If Row >= .FixedRows And Row < .Rows Then
            If Col = 1 Then
                If Not m.Chart Is Nothing Then
                    strName = .TextMatrix(Row, 0)
                    strType = .TextMatrix(Row, 2)
                    m.Chart.LinkedInputGet strName, strValue, strType
                    
                    If strType = "6" Then
                        If strValue = "1" Or strValue = "True" Then
                            If .TextMatrix(Row, 1) <> "True" Then
                                strValue = "0"
                                bChanged = True
                            End If
                        ElseIf .TextMatrix(Row, 1) <> "False" Then
                            strValue = "1"
                            bChanged = True
                        End If
                    ElseIf strValue <> .TextMatrix(Row, 1) Then
                        If strType = 5 Then
                            strValue = Chr(34) & UCase(.TextMatrix(Row, 1)) & Chr(34)
                            .TextMatrix(Row, 1) = StripStr(strValue, Chr(34))
                        Else
                            strValue = .TextMatrix(Row, 1)
                        End If
                        bChanged = True
                    End If
                    
                    If bChanged Then
                        m.Chart.LinkedInputSet strName, strValue, strType, True
                        m.Chart.RedoMode = eRedo3_Settings
                        m.iGenerateChart = True
                        ClearCodedText
                    End If
                End If
            End If
        End If
    End With

End Sub

Private Sub fgLinkedInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim i&
    Dim strName$, strVal$, strType$

    If fgSettings.Redraw = flexRDNone Then Exit Sub

    With fgLinkedInputs
        If Row >= .FixedRows And Row < .Rows Then
            If Col = 1 Then
                strName = .TextMatrix(Row, 0)
                If Not m.Chart Is Nothing Then m.Chart.LinkedInputGet strName, strVal, strType

                i = .TextMatrix(Row, 2)
                If i = 6 Then
                    .ComboList = "False|True"
                    If strVal = "TRUE" Or strVal = "T" Or strVal = "1" Then
                        .TextMatrix(Row, Col) = "True"
                    Else
                        .TextMatrix(Row, Col) = "False"
                    End If
                Else
                    .ComboList = ""
                End If
            Else
                .ComboList = ""
                Cancel = True
            End If
        End If
    End With

End Sub

Private Sub fgLinkedInputs_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    GridScrollCheck fgLinkedInputs, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel

End Sub

Private Sub fgLinkedInputs_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    FinishEdit = True
End Sub

Private Sub fgSettings_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim nNode&, bDisplay As Boolean
    Dim Pane As cPane
    Dim Ind As cIndicator

    If fgSettings.Redraw = flexRDNone Then Exit Sub
        
    If Col = kBoxCol Then
        With fgSettings
            bDisplay = CheckedCell(fgSettings, Row, kBoxCol)
            CheckedCell(fgSettings, Row, kShowCol) = bDisplay
            ColorSettingsRow Row
            nNode = .RowData(Row)
            If RowType(Row) = kPane Then
                Set Pane = m.Chart.Tree(nNode)
                Pane.Display = bDisplay
                CollapseRows
                m.iGenerateChart = True         '6453
            ElseIf RowType(Row) = kIndicator Then
                Set Ind = m.Chart.Tree(nNode)
                If Not Ind Is Nothing Then
                    Ind.Display = bDisplay
                    
                    If Ind.DataType <> eINDIC_BooleanArray Then
                    
                        If Not bDisplay And Ind.DisplayType = eINDIC_Ribbon Then
                            cboType.ListIndex = 0        '6227
                        End If
                    
                        Ind.ToggleBoolRefInd m.Chart.Tree, bDisplay 'toggle off/on any associated highlight bars
                        
                    ElseIf bDisplay = True And Ind.DataType = eINDIC_BooleanArray Then
                        Ind.ToggleIndToColor m.Chart.Tree, True     'turn on indicator this highlight bar applies to
                    End If
                    
                    'If Not Ind.Display Then CheckPaneDisplay nNode
                    CheckPaneDisplay nNode, Ind.Display
                    m.iGenerateChart = True     '6453
                End If
            End If
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgSettings.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgSettings_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim i&

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    
    If NewRow <> OldRow And fgSettings.Redraw <> 0 Then
        EnableTabs
        MoveFocus fgSettings
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgSettings.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgSettings_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Static strMsg As String

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    
    Dim Ind As cIndicator
    
    If Col = kNameCol Then
        Cancel = True
    ElseIf RowType(Row) = kIndicator Then
        If Len(strMsg) > 0 Then
            InfBox strMsg, , , "Volume Profile"
            strMsg = ""
        Else
            Set Ind = m.Chart.Tree(fgSettings.RowData(Row))
            If Not Ind Is Nothing Then
                If Ind.DataType = eINDIC_ProfileVolume Then
                    If SecurityType(m.Chart.Symbol) = "I" Then
                        'JM 05-14-2015: calling infbox here will cause message to display twice
                        strMsg = "Volume profile cannot be turned on for this chart."
                        Cancel = True
                    End If
                End If
            End If
        End If
    ElseIf RowType(Row) <> kPane Then
        Cancel = True
        'Beep ' invalid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgSettings.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgSettings_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:
    
    'If Button = 2 Then Exit Sub
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If fgSettings.MouseRow < 0 Then Exit Sub
    If Shift = 1 Then Exit Sub
    
    With fgSettings
        .Row = .MouseRow
        SelectRows
        .Refresh
        If Button = 2 Then
            Cancel = True
            ShowAddPopup False
        ElseIf .MouseCol = kNameCol Then
            Cancel = True
            .DragRow .Row
        End If
    End With
    
    If Not m.EditedIndicator Is Nothing Then
        '??? TLB: to fix problem where everything bogs down on page with lots of charts
        m.bSkipGenerate = True
        gdSelectIcon.Icon = gdSelectIcon.Icon
        m.bSkipGenerate = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgSettings.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgSettings_BeforeMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    Dim nFromRow&, nToRow&, idxFromNode&, idxToNode&
    Dim nParent&, nParentLevel&, idxMovedTo&
    Dim Pane As cPane
    
    Dim IndFrom As cIndicator, IndTo As cIndicator
    
    If Row = Position Then Exit Sub
    
    With m.Chart.Tree
        
        ' if moving to prior row, adjust so
        ' will move it to "after" the row before
        nFromRow = Row
        nToRow = Position
        If nToRow < nFromRow Then nToRow = nToRow - 1
         
        idxFromNode = fgSettings.RowData(nFromRow)
        idxToNode = fgSettings.RowData(nToRow)
        Select Case RowType(nFromRow)
            
            Case kPane:
                ' move the entire Pane
                If RowType(nToRow) = kNewPane Then
                    idxToNode = -1 ' move to end
                ElseIf idxToNode = 0 Then
                    ' above the tree, so move to beginning,
                    idxToNode = 0
                Else
                    ' move as next sibling of this root
                    idxToNode = .RelativeIndex(idxToNode, eTREE_Root)
                End If
                idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_NextSibling)
                
            Case kIndicator:
                Set IndFrom = m.Chart.Tree(idxFromNode)
                IndFrom.SaveGroupInfo
                
                If IndFrom.DisplayType = eINDIC_Ribbon Then
                    With m.Chart.Tree
                        If .RelativeIndex(idxFromNode, eTREE_Parent) <> .RelativeIndex(idxToNode, eTREE_Parent) Then
                            RibbonList Nothing, m.Chart, IndFrom, 2, 0          '6219
                        End If
                    End With
                End If
                
                ' check if indicator is highlight bars
                If IndFrom.DataType = eINDIC_BooleanArray Then
                    If .NodeLevel(idxToNode) > 0 Then   'must move to an indicator level
                        Set IndTo = m.Chart.Tree(idxToNode)
                        If IndTo.DataType <> eINDIC_Constant Then 'do not highlight horz indicator
                            If IndTo.DataType = eINDIC_BooleanArray Or _
                               .NodeLevel(idxToNode) = .NodeLevel(idxFromNode) Then     'fix for aardvark 689
                                idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_NextSibling)
                            Else
                                idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_FirstChild)
                                IndFrom.Display = IndTo.Display
                            End If
                        End If
                        If idxMovedTo > 0 Then IndFrom.geIndId = idxMovedTo
                    End If
                    Set IndTo = Nothing
                ' check if indicator is "unlinked"
                ElseIf .NodeLevel(idxFromNode) = 1 Then
                    If .Key(idxFromNode) = "PRICE" And _
                        .Key(.RelativeIndex(idxToNode, eTREE_Root)) <> "PRICE PANE" Then
                            Beep 'can't move price from price pane
                    ElseIf RowType(nToRow) = kNewPane Then
                        ' moving to a new pane
                        Set Pane = New cPane
                        Pane.Display = True
                        Pane.Scaling = ePANE_ScaleModeAuto
                        idxToNode = .Add(Pane)
                        idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_FirstChild)
                    ElseIf idxToNode = 0 Then
                        Beep ' above the tree is invalid
                    ElseIf RowType(nToRow) = kPane Then
                        ' move to be first child of this Pane
                        idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_FirstChild)
                    Else
                        ' get the level 1 ancestor of the ToNode
                        idxToNode = .AncestorIndex(idxToNode, 1)
                        idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_NextSibling)
                    End If
                Else
                    ' linked indicators MUST stay under their current
                    ' parent (and stay at the same level)
                    nParent = .RelativeIndex(idxFromNode, eTREE_Parent)
                    nParentLevel = .NodeLevel(nParent)
                    ' if ToNode is parent, move to be first child
                    If idxToNode = nParent Then
                        idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_FirstChild)
                    ' otherwise the ancestor of the ToNode at the parent
                    ' level must be the parent of the FromNode
                    ElseIf nParent = .AncestorIndex(idxToNode, nParentLevel) Then
                        ' then move to be next sibling of ToNode
                        ' ancestor at FromNode level
                        idxToNode = .AncestorIndex(idxToNode, nParentLevel + 1)
                        idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_NextSibling)
                    Else
                        Beep ' else an invalid move of a linked indicator
                    End If
                End If
            
            Case Else:
                ' can't move these rows
                Beep
        End Select
    
        Set Pane = Nothing
    
        If idxMovedTo > 0 And idxMovedTo <> idxFromNode Then
            If Not IndFrom Is Nothing Then IndFrom.CheckGroup
            LoadSettingsGrid
            SelectIdx idxMovedTo
            ClearCodedText idxMovedTo
        End If
    End With

    Position = Row ' do this so we don't actually change it in the grid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgSettings.BeforeMoveRow", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgSettings_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    GridScrollCheck fgSettings, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel

End Sub

Private Sub fgSettings_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    SelectRows

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.fgSettings.KeyUp", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_Activate()
On Error GoTo ErrSection:

    Dim i&
    Static bAlreadyDone As Boolean

    If Not bAlreadyDone Then
        bAlreadyDone = True
        'LoadSettingsGrid
    End If

    If m.Chart Is Nothing Then
        LoadSettingsGrid
    End If

    m.iGenerateChart = False '(cleared)
    tmrChartCfg.Enabled = True

    SetCaption

    DoEvents
    EnableTabs

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    m.iGenerateChart = False '(cleared)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.Form.Deactivate", eGDRaiseError_Show
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
    RaiseError "frmChartCfg.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim i&

    Me.Width = kFormWidth
    mnuAdd.Visible = False
    
    g.Styler.StyleForm Me
    
    'Me.Width = (lineGray.X2 - lineGray.X1) + lineGray.X1 * 2 + (Me.Width - Me.ScaleWidth)
    'Me.Height = vst.top + vst.Height + (Me.Height - Me.ScaleHeight) + 120
    
    If Screen.TwipsPerPixelY = 12 Then
        ' adjustments when screen set to "large fonts"
        picMoveUp.Left = picMoveUp.Left + 12
        picMoveUp.Top = picMoveUp.Top + 12
        picMoveDown.Left = picMoveDown.Left + 12
        picMoveDown.Top = picMoveDown.Top + 12
    End If
    
    SizeFormToControl Me, Corner
    
    CenterFormOnChart Me, Chart
    Me.Icon = Picture16("kToolsEdit")   'JM:03-30-2009 - this icon does not exist; form has no icon
        
    cmdMarketInfo.Top = chkUnsplit.Top - 60
    
    'picMoveUp.Picture = Picture16("kSortedUpArrow")
    'picMoveDown.Picture = Picture16("kSortedDownArrow")

    fgSettings.Redraw = flexRDNone

    With cboLineDefault
        .AddItem "Thin"
        .AddItem "Medium Thin"
        .AddItem "Medium"
        .AddItem "Medium Thick"
        .AddItem "Thick"
        .AddItem "Extra Thick"
        .AddItem "Dashed (Large)"
        .AddItem "Dashed (Small)"
        .AddItem "Dash Dot"
    End With
    
    With cboLineTypes
        .Clear
        .AddItem "(default)"
        .AddItem "Thin"
        .AddItem "Medium Thin"
        .AddItem "Medium"
        .AddItem "Medium Thick"
        .AddItem "Thick"
        .AddItem "Extra Thick"
    End With
    
    With cboProfitLines
        .AddItem "Thin"
        .AddItem "Medium"
        .AddItem "Thick"
        .AddItem "Dashed (Large)"
        .AddItem "Dashed (Small)"
        .AddItem "Dash Dot"
    End With
    
    With cboHorzDefault
        .AddItem "Thin"
        .AddItem "Medium Thin"
        .AddItem "Medium"
        .AddItem "Medium Thick"
        .AddItem "Thick"
        .AddItem "Extra Thick"
        .AddItem "Dashed (Large)"
        .AddItem "Dashed (Small)"
        .AddItem "Dash Dot"
    End With

    With cboBarsDefault
        .AddItem "Thin"
        .AddItem "Medium Thin"
        .AddItem "Medium"
        .AddItem "Medium Thick"
        .AddItem "Thick"
        .AddItem "Extra Thick"
        .AddItem "Auto (variable)"
    End With
           
    With cboLineTypesBars
        .AddItem "(default)"
        .AddItem "Thin"
        .AddItem "Medium Thin"
        .AddItem "Medium"
        .AddItem "Medium Thick"
        .AddItem "Thick"
        .AddItem "Extra Thick"
        .AddItem "Auto (variable)"
    End With
    
    With cboType
        .AddItem "Line"
        .AddItem "Histogram"
        .AddItem "Mountain"
        .AddItem "Points"
        .AddItem "Steps"
        .AddItem "Rectangles"
        .AddItem "Ribbon"
        .AddItem "None"
        .ListIndex = 0
    End With
    
    With cboOHLCType
        .AddItem "HL bars"
        .AddItem "HLC bars"
        .AddItem "OHLC bars"
        .AddItem "Candlesticks"
        .AddItem "Bollinger Bars"
        .AddItem "Close line"
        .AddItem "Histogram"
        .AddItem "Mountain"
        .AddItem "Points"
        .AddItem "Point & Figure"
        .AddItem "Kagi"
        .AddItem "Renko"
        .AddItem "None"
        .ListIndex = 2
        '.Move cboType.Left, cboType.Top
    End With
    
    With cboHighlightBars
        .AddItem "HighlightBars"
        .AddItem "HighlightMarkers"
        .AddItem "HighlightBoxes"
        .AddItem "None"
        .ListIndex = 0
        .Move cboType.Left, cboType.Top
    End With
    
    With cboMarkerLoc
        .Clear
        .AddItem "Above"
        .AddItem "Below"
        .ListIndex = 0
    End With
    
    With cboBoxPenStyle
        .AddItem "(default)"
        .AddItem "Thin"
        .AddItem "Medium"
        .AddItem "Thick"
        .AddItem "Dashed (Large)"
        .AddItem "Dashed (Small)"
        .AddItem "Dash Dot"
    End With
        
    With cboBarPeriod
        .AddItem "5 minute"
        .AddItem "10 minute"
        .AddItem "15 minute"
        .AddItem "30 minute"
        .AddItem "60 minute"
        .AddItem "Daily"
        .AddItem "Weekly"
        .AddItem "Monthly"
        .AddItem "Quarterly"
        .AddItem "Yearly"
        .AddItem "2 days"
        .AddItem "3 days"
        .AddItem "4 days"
    End With
    
    With cboIndLabelDefault        'this is in the chart's defaults setting dialog
        .Clear
        .AddItem "Value in axis"
        .AddItem "Value in label"
        .AddItem "Value in label/axis"
        .AddItem "No values"
        .AddItem "No labels"
        .AddItem "No labels or values"
        .AddItem "Only values"
    End With
        
    With cboIndLabelMode           'this is in the individual indicator setting dialog
        .Clear
        .AddItem "(default)"
        .AddItem "Value in axis"
        .AddItem "Value in label"
        .AddItem "Value in label/axis"
        .AddItem "No values"
        .AddItem "No labels"
        .AddItem "No labels or values"
        .AddItem "Only values"
    End With
    
    With cboBarsLabelMode           'this is in the individual indicator setting dialog
        .Clear
        .AddItem "(default)"
        .AddItem "Value in axis"
        .AddItem "Value in label"
        .AddItem "Value in label/axis"
        .AddItem "No values"
        .AddItem "No labels"
        .AddItem "No labels or values"
        .AddItem "Only values"
    End With
    
    With cboVertGrid
        .Clear
        .AddItem "Coarse"
        .AddItem "Fine"
        .AddItem "None"
    End With
    
    With cboAnnotDefault
        .Clear
        .AddItem "Thin"
        .AddItem "Medium"
        .AddItem "Thick"
        .AddItem "Dashed (Large)"
        .AddItem "Dashed (Small)"
        .AddItem "Dash Dot"
    End With
    
    With cboMarkerSize              'aardvark 3288
        .Clear
        .AddItem "Small"
        .AddItem "Medium"
        .AddItem "Large"
    End With
    
    'RH commented out fraShift.BorderStyle = 0

    With fraAppearance
        If g.nColorTheme = kDarkThemeColor Then .ForeColor = vbWhite
        fraLinearLog.ForeColor = .ForeColor
        fraAppearanceBars.ForeColor = .ForeColor
        fraLinkedInputs.ForeColor = .ForeColor
        fraTemplate = .ForeColor
        fraPaneScale.ForeColor = .ForeColor
        fraPaneDisplay.ForeColor = .ForeColor
        fraColors.ForeColor = .ForeColor
        fraMisc.ForeColor = .ForeColor
        fraDefaults.ForeColor = .ForeColor
        fraXaxis.ForeColor = .ForeColor
        fraTrades.ForeColor = .ForeColor
        fraSystem.ForeColor = .ForeColor
        fraProfitLines.ForeColor = .ForeColor
        fraDates.ForeColor = .ForeColor
        fraBars.ForeColor = .ForeColor
        fraFunction.ForeColor = .ForeColor
        fraArtPyramids.ForeColor = .ForeColor
        fraArtPyramids.Left = .Left
        fraArtPyramids.Top = .Top
        lblSymbol.ForeColor = .ForeColor
        lblDesc.ForeColor = .ForeColor
        lblCustomFunction.ForeColor = .ForeColor
        lblSystemName.ForeColor = .ForeColor
    End With
    
    With vseIndicator
        vseGeneral.Move .Left, .Top - 150   ', .Width, .Height
        vseXaxis.Move .Left, .Top, .Width, .Height
        vseSystem.Move .Left, .Top, .Width, .Height
        vseLinkedInputs.Move .Left, .Top, .Width, .Height - 130
        vsePane.Move .Left, .Top, .Width, .Height
        vseBars.Move .Left, .Top, .Width, .Height
        vseGeneral.Appearance = apFlat
        vseXaxis.Appearance = apFlat
        vseSystem.Appearance = apFlat
        vsePane.Appearance = apFlat
        vseBars.Appearance = apFlat
    End With
       
    ShowInputTip -1     'clear tip
    fraFunction.Left = fraAppearance.Left
    fgSettings.Redraw = flexRDBuffered
    gdSelectIcon.Move cboLineTypes.Left, cboLineTypes.Top - (gdSelectIcon.Height - cboLineTypes.Height) / 2
    lblUpColor.Move chkUpDownColors.Left, chkUpDownColors.Top
    lblUpColor.Visible = False
    chkOverlayed(1).Top = chkColorPrice.Top
    'position controls for highlight boxes
    cboBoxPenStyle.Move cboLineTypes.Left, cboLineTypes.Top
    lblBarsLeft.Move lblWidth.Left, chkColorPrice.Top - 20
    txtBarsLeft.Move cboLineTypes.Left, lblBarsLeft.Top - 50
    lblBarsRight.Move txtBarsLeft.Left + txtBarsLeft.Width + 80, lblBarsLeft.Top
    txtBarsRight.Move txtBarsRight.Left, lblBarsRight.Top - 50
    
    InitSettingsGrid
    InitInputsGrid
    InitColorsGrid
    InitLinkedInputsGrid
    InitTemplates
    InitCboAccounts
    
    gdSelectIcon.AllowCustom = True
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.Form.Load", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowInputTip -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        'Me.Hide
        'cmdOK_Click '(for now)'
        m.bClickOK = True
    End If
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    bNowAdding = False
    m.iGenerateChart = False '(cleared)
    tmrChartCfg.Enabled = False
    Set m.Chart = Nothing
    'frmMain.DockPro.RemoveForm Me.Name
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub CollapseRows(Optional ByVal strUngroup As String = "")
On Error GoTo ErrSection:

    Dim i&, nCurRow&, strKey$, strGroupKey$
    
    Dim bPaneHide As Boolean
    Dim bRowHide As Boolean
    
    Dim Pane As cPane
    Dim Ind As cIndicator
    Dim Tree As cGdTree
    
    If Not m.Chart Is Nothing Then Set Tree = m.Chart.Tree
    
    Set m.PaneWood = Nothing
    
    With fgSettings
        .Redraw = flexRDNone
        nCurRow = .Row
        bPaneHide = False
        For i = .FixedRows To .Rows - 1
            strKey = ""
            strGroupKey = ""
            Select Case RowType(i)
                Case kPane
                    bRowHide = False
                    bPaneHide = Not CheckedCell(fgSettings, i, kShowCol)
                    Set Pane = Tree(.RowData(i))
                    If Not Pane Is Nothing Then
                        If Tree.Key(.RowData(i)) = kClusterTimeKeyPane Then
                            bPaneHide = True
                            bRowHide = True
                        ElseIf Pane.SplitPaneType = ePANE_SplitPaneWood Then
                            Set m.PaneWood = Pane
                        End If
                    End If
                Case kIndicator
                    If Not Tree Is Nothing Then
                        Set Ind = Tree(.RowData(i))
                        If Not Ind Is Nothing Then
                            If Len(Ind.GroupKey) > 0 Then
                                strGroupKey = Ind.GroupKey
                                If Len(strUngroup) > 0 And strGroupKey = strUngroup Then
                                    Ind.GroupKey = ""
                                    strGroupKey = ""
                                ElseIf Len(strGroupKey) > 0 Then
                                    strKey = Tree.Key(.RowData(i))
                                    strGroupKey = Ind.GroupKey
                                End If
                            End If
                        End If
                    End If
                    If Tree.Key(.RowData(i)) = kClusterPriceKey Or m.bSeasonal Then
                        bRowHide = True
                    Else
                        bRowHide = bPaneHide
                    End If
                Case kSystem
                    bRowHide = Not HasGold(False, , False)
                    bRowHide = m.bSeasonal
                Case Else
                    If .TextMatrix(i, kNameCol) = "General Settings" Then
                        bRowHide = False
                    Else
                        bRowHide = m.bSeasonal
                    End If
            End Select
            If kCollapseRows Then
                .RowHidden(i) = bRowHide
            End If
            If bRowHide Then
                .Cell(flexcpChecked, i, kBoxCol) = flexNoCheckbox
                .Cell(flexcpForeColor, i, kNameCol) = RGB(128, 128, 128)
            ElseIf strKey = strGroupKey Then
                CheckedCell(fgSettings, i, kBoxCol) = CheckedCell(fgSettings, i, kShowCol)
                .Cell(flexcpForeColor, i, kNameCol) = .Cell(flexcpForeColor, i, kShowCol)
            Else
                .Cell(flexcpChecked, i, kBoxCol) = flexNoCheckbox
            End If
        Next
        .Cell(flexcpChecked, .FixedRows, kBoxCol) = flexNoCheckbox
        .Cell(flexcpChecked, .FixedRows + 1, kBoxCol) = flexNoCheckbox
        .Cell(flexcpChecked, .FixedRows + 2, kBoxCol) = flexNoCheckbox
        .Cell(flexcpChecked, .FixedRows + 3, kBoxCol) = flexNoCheckbox
        .Cell(flexcpChecked, .Rows - 1, kBoxCol) = flexNoCheckbox
        .Row = nCurRow
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.CollapseRows", eGDRaiseError_Raise
        
End Sub

Public Property Get Chart() As cChart
On Error GoTo ErrSection:

    Set Chart = m.Chart

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmChartCfg.Chart.Get", eGDRaiseError_Raise
        
End Property

Private Property Set Chart(vChart As cChart)
On Error GoTo ErrSection:

    Dim s$
    Dim Pane As cPane
    
    Set m.Chart = vChart
    If m.Chart Is Nothing Then GoTo ErrExit
    
    If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
        m.bSeasonal = False
    Else
        m.bSeasonal = True
    End If
    
    LoadSettingsGrid
    
    fgSettings.Redraw = flexRDNone
    With g.ChartGlobals
        cboBarsDefault.ListIndex = .eDefaultBarsStyle - 1
        cboLineDefault.ListIndex = .eDefaultIndStyle - 1
        cboHorzDefault.ListIndex = .eDefaultHorzStyle - 1
        cboIndLabelDefault.ListIndex = .eDefaultLabelMode - 1       'TODO: doublecheck this with Tim
        cboAnnotDefault.ListIndex = .eDefaultAnnotStyle - 1
        cboProfitLines.ListIndex = .eProfitLineStyle - 1
        chkTips = Abs(.bFloatingTips)
        chkChartTips = Abs(.bChartTips)
        chkSplitsRolls = Abs(.bSplitsRolls)
        clrSplitsRolls.Color = .nSplitRollColor
        clrLong.Color = .nLongColor
        clrShort.Color = .nShortColor
        clrLoss.Color = .nLossColor
        clrWin.Color = .nWinColor
    End With
    
    ShowChartColors
    
    With m.Chart
        fraDefaults.Enabled = Not m.bSeasonal
        
        lblBarsDefault.Enabled = Not m.bSeasonal
        lblLineDefault.Enabled = Not m.bSeasonal
        lblHorzDefault.Enabled = Not m.bSeasonal
        lblIndLabelDefault.Enabled = Not m.bSeasonal
        lblAnnotDefault.Enabled = Not m.bSeasonal
        
        cboBarsDefault.Enabled = Not m.bSeasonal
        cboLineDefault.Enabled = Not m.bSeasonal
        cboHorzDefault.Enabled = Not m.bSeasonal
        cboIndLabelDefault.Enabled = Not m.bSeasonal
        cboAnnotDefault.Enabled = Not m.bSeasonal
        
        picAdd.Visible = Not m.bSeasonal
        picAdd.Enabled = Not m.bSeasonal
        cmdAdd.Enabled = Not m.bSeasonal
        
        cmdTemplate.Enabled = Not m.bSeasonal
        cmdRemove.Enabled = Not m.bSeasonal
        cmdShowBidAsk.Enabled = Not m.bSeasonal
        
        chkSplitPane.Enabled = Not m.bSeasonal
        chkSplitsRolls.Enabled = Not m.bSeasonal
    
        LoadSystemInfo
        chkEmptyBars = Abs(.ShowEmptyBars)
        chkUnsplit = Abs(.Unsplit)
        chkHorzGrid = Abs(.HorzGrid)
        
        chkSplitPane = .ShowSplitPane
        If chkSplitPane.Value = False Then .ShowSplitPane = 0
        cmdSplitPaneCfg.Enabled = chkSplitPane.Value
        
        txtForecastBars = Str(.BlankBars)
        s = GetPeriodStr(.Periodicity)
        If Len(s) = 0 Then s = "Daily"
        cboBarPeriod.Text = s
        dtFromDate = .FromDate
        dtToDate = .ToDate
        If .ToEndOfData Then
            optEndOfData = True
        Else
            optToDate = True
        End If
        cboVertGrid.ListIndex = .VertGrid
        txtMaxDays.Text = Str(.MaxIntradayDays)
    End With
    
    fgSettings.Redraw = flexRDBuffered

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmChartCfg.Chart.Set", eGDRaiseError_Raise
        
End Property

Private Sub InitSettingsGrid()
On Error GoTo ErrSection:

    With fgSettings
        .Redraw = flexRDNone
        .FillStyle = flexFillRepeat
        .GridLines = flexGridFlatHorz
        .FixedCols = 0
        .FixedRows = 1
        .Rows = .FixedRows
        .AllowUserResizing = flexResizeColumns
        '.ExplorerBar = flexExSortShowAndMove
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .Editable = flexEDKbdMouse
        
        .Cols = 3
        .ColDataType(kShowCol) = flexDTBoolean
        .ColHidden(kShowCol) = True
        .TextMatrix(0, kBoxCol) = "Show"
        .TextMatrix(0, kNameCol) = "Settings to Edit ..."
        .AutoSize kBoxCol
        .Select 0, 0, 0, .Cols - 1
        .CellFontBold = True
        .CellForeColor = fraAppearance.ForeColor
        .ExtendLastCol = True
        
        .FillStyle = flexFillSingle
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.InitSettingsGrid", eGDRaiseError_Raise
        
End Sub

Private Sub InitInputsGrid()
On Error GoTo ErrSection:

    With fgInputs
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .AllowSelection = True ' False
        .HighLight = flexHighlightWithFocus
        '.AllowUserFreezing = flexFreezeColumns
        .SheetBorder = RGB(128, 128, 128)
        '.ExtendLastCol = True
        .Editable = flexEDKbdMouse
        
        .Rows = 5
        .Cols = 2
        .TextMatrix(0, 0) = "  Input Name"
        .TextMatrix(0, 1) = "  Value  "
        .FillStyle = flexFillRepeat
        .Select 0, 0, 0, .Cols - 1
        .CellFontBold = True
''        .CellForeColor = fraAppearance.ForeColor
        .AutoSize 1
        .ColWidth(0) = .Width - .ColWidth(1) - 4 * Screen.TwipsPerPixelX
        .ExtendLastCol = True
        .ColAlignment(1) = flexAlignCenterCenter
        
        .TextMatrix(1, 0) = "# Bars"
        .TextMatrix(1, 1) = "14"
        
        .FillStyle = flexFillSingle
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.InitInputsGrid", eGDRaiseError_Raise
        
End Sub

Private Sub EnableTabs()
On Error GoTo ErrSection:

    Static nPrevIndTab&

    Dim idx&, i&, iPos&, iSaveRedraw&, d#
    Dim strText$, strDontCare$
    Dim bCombo As Boolean
    
    Dim Indicator As cIndicator, IndParent As cIndicator
    Dim Pane As cPane
    
    SaveModifiedInputs
    
    Set m.EditedIndicator = Nothing
    Set m.EditedPane = Nothing
    If fgSettings.Row < 0 Then Exit Sub
    iSaveRedraw = fgSettings.Redraw
    fgSettings.Redraw = flexRDNone
    
    ' use tag to store which "frame" should be visible
    ' (since don't want it to flash by turning current off and back on again)
    vseGeneral.Tag = ""
    vseXaxis.Tag = ""
    vseSystem.Tag = ""
    vseLinkedInputs.Tag = ""
    vseIndicator.Tag = ""
    vseBars.Tag = ""
    vsePane.Tag = ""
    
    idx = fgSettings.RowData(fgSettings.Row)
        
    cmdSaveStudy.Caption = "&Save Study"        'default to this is so caption will be same as previous versions
    
    Select Case RowType(fgSettings.Row)
        Case kNewPane:
            cmdSaveStudy.Enabled = False
            cmdRemove.Enabled = False
    
        Case kGeneral:
            vseGeneral.Tag = "-1" ' True
            cmdSaveStudy.Enabled = False
            cmdRemove.Enabled = False
        
        Case kXaxis:
            vseXaxis.Tag = "-1"  ' True
            txtBarWidth.Visible = False     'volume profile
            txtBarWidth.Enabled = False     'volume profile
            txtForecastBars.Visible = True
            txtForecastBars.Enabled = True
            cmdSaveStudy.Enabled = False
            cmdRemove.Enabled = False
            cmdBarPeriod.Visible = True
            cboBarPeriod.Visible = True
            chkHorzGrid.Visible = True
            chkEmptyBars.Visible = True
            vseXaxis.Move vseIndicator.Left, vseIndicator.Top, vseXaxis.Width, kVseFullHeight
            
            If fraXaxis.Caption <> "Dates Axis" Then
                fraXaxis.Caption = "Dates Axis"
                fraXaxis.Height = 2175
                Label22.Top = 780
                cboVertGrid.Top = 720
                chkHorzGrid.Top = 1080
                chkEmptyBars.Top = 1380
                Label21.Top = 1740
                Label22.Caption = "Vertical grid lines:"
'                Label21.Caption = "Blank bars after end of data:"
                Label21.Visible = True
                With cboVertGrid
                    .Clear
                    .AddItem "Coarse"
                    .AddItem "Fine"
                    .AddItem "None"
                End With
                cboVertGrid.ListIndex = m.Chart.VertGrid
            End If
            
        Case kSystem:
            vseSystem.Tag = "-1"  ' True
            cmdSaveStudy.Enabled = False
            cmdRemove.Enabled = True
            cmdRemove.ToolTipText = "Remove Strategy"
        
        Case kLinkedInputs
            vseLinkedInputs.Tag = "-1"  ' True
            cmdSaveStudy.Enabled = False
            cmdRemove.Enabled = False
            
        Case kIndicator:
            Set Indicator = m.Chart.Tree(idx)
            
            vseIndicator.Height = kVseFullHeight
            
            cmdSaveStudy.Enabled = Not m.bSeasonal
            cmdRemove.Enabled = True
            cmdRemove.ToolTipText = "Remove Indicator"
            If Indicator.DataType = eINDIC_BarData Then
                vseBars.Tag = "-1"  ' True
            Else
                vseIndicator.Tag = "-1"  ' True
            End If
                
            With Indicator
                txtFunction = .Name
                If Len(Indicator.GroupKey) > 0 Then
                    cmdSaveStudy.Caption = "&Ungroup"
                Else
                    cmdSaveStudy.Caption = "&Group"
                End If
                If .DataType = eINDIC_BarData Then
                    If m.Chart.Tree.Key(idx) = "PRICE" Then
                        i = g.SymbolPool.PoolRecForSymbolID(m.Chart.SymbolID)
                    Else
                        i = g.SymbolPool.PoolRecForSymbol(.Name)
                    End If
                    If i >= 0 Then
                        lblSymbol = g.SymbolPool.Symbol(i)
                        lblDesc = g.SymbolPool.Desc(i)
                    Else
                        lblSymbol = ""
                        lblDesc = ""
                    End If
                    ' TLB 3/17/2014: allow for changing mkt info for custom symbol
                    If InStr("SM", SecurityType(lblSymbol.Caption)) > 0 And Left(m.Chart.Symbol, 1) <> "*" Then
                        chkUnsplit.Visible = True
                        cmdMarketInfo.Visible = False
                    ElseIf Len(m.Chart.SpreadSymbols) > 0 Then
                        chkUnsplit.Visible = False
                        cmdMarketInfo.Visible = False
                    Else
                        chkUnsplit.Visible = False
                        cmdMarketInfo.Visible = True
                    End If
                ElseIf .IsCustom Then
                    lblCustomFunction = .Expression
                ElseIf .DataType = eINDIC_ProfileVolume Then
                    vseXaxis.Tag = "-1"  ' True
                    'txtBarWidth is days per bar (reused name from frmPriceVolCfg)
                    txtBarWidth.Visible = True
                    txtBarWidth.Enabled = True
                    txtForecastBars.Visible = False
                    txtForecastBars.Enabled = False
                    cmdBarPeriod.Visible = False
                    cboBarPeriod.Visible = False
                    chkHorzGrid.Visible = False
                    chkEmptyBars.Visible = False
                    fraFunction.Visible = False
                    
                    If fraXaxis.Caption <> "Profile Periodicity" Then
                        fraXaxis.Caption = "Profile Periodicity"
                        fraXaxis.Height = 1035
                        
                        Label21.Visible = False
                        Label22.Caption = "Bar period:"
                        Label22.Top = 450
                        cboVertGrid.Top = 420
                        txtBarWidth.Top = 420
                        txtBarWidth.Left = cboVertGrid.Left - txtBarWidth.Width - 30
                    End If
                    
                    fraArtPyramids.Height = 4170
                    vseXaxis.Move vseIndicator.Left, vseIndicator.Top + fraArtPyramids.Height + 360, vseXaxis.Width, fraXaxis.Height + 90
                Else
                    fraFunction.Visible = True
                    fraArtPyramids.Height = 2715
                    ' show inputs
                    fgInputs.Redraw = flexRDNone
                    fgInputs.Rows = fgInputs.FixedRows
                    For i = 1 To .ParmCount
                        strText = ""
                        ' Real users should not be able to change the inputs for the TAS_Result indicators
                        If UCase(.CodedName) = "TAS_RESULT" And Not IsIDE Then
                            Exit For
                        End If
                        ' Per TLB: do not show Date parm for Hawkeye Adds & Hawkeye Levels
                        If .IsHawkeyeLevels And i = 2 Then
                            Exit For
                        ElseIf .IsHawkeyeAdds And i = 3 Then
                            Exit For
                        End If
                        If .DataType = eINDIC_Constant Then
                            If .IsLinkedParm(i) Then
                                If Not .LinkedParmGet(i, strText, strDontCare) Then strText = Replace(Trim(.Parm(i)), Chr(34), "")
                            Else
                                strText = Replace(Trim(.Parm(i)), Chr(34), "")
                            End If
                            fgInputs.AddItem "Value" & vbTab & strText
                            fgInputs.RowData(fgInputs.Rows - 1) = i
                        Else
                            Select Case .ParmType(i)
                            Case 1, 2 ' constant number or text
                                If .IsLinkedParm(i) Then
                                    If Not .LinkedParmGet(i, strText, strDontCare) Then strText = Replace(Trim(.Parm(i)), Chr(34), "")
                                Else
                                    strText = Replace(Trim(.Parm(i)), Chr(34), "")
                                End If
                                fgInputs.AddItem .ParmName(i) & vbTab & strText
                                fgInputs.RowData(fgInputs.Rows - 1) = i
                            Case 4 ' array of numbers
                                strText = UCase(.ParmDefault(i))
                                If ParmDefaultIsBarsArray(strText) Then
                                    bCombo = False
                                    If m.Chart.Tree.NodeLevel(idx) = 1 Then
                                        bCombo = True
                                    ElseIf m.Chart.Tree.NodeLevel(idx) = 2 Then
                                        Set IndParent = m.Chart.Tree.RelativeItem(idx, eTREE_Parent)
                                        If IndParent.DataType = eINDIC_BarData Then
                                            bCombo = True
                                        End If
                                    End If
                                    If bCombo Then
                                        strText = Trim(.Parm(i))
                                        ' convert if market parm changed to expression (e.g. RSI)
                                        If UCase(strText) = "MARKET1" Then
                                            strText = "Close"
                                        End If
                                        iPos = InStr(UCase(strText), " OF ")
                                        If iPos > 0 Then strText = Trim(Left(strText, iPos))
                                        fgInputs.AddItem .ParmName(i) & vbTab & strText
                                        fgInputs.RowData(fgInputs.Rows - 1) = i
                                    End If
                                ElseIf strText = "GOLD" Then
                                    If UCase(.Parm(i)) = "GOLD" Then .Parm(i) = "default"
                                    fgInputs.AddItem "Symbol" & vbTab & Trim(StripStr(.Parm(i), Chr(34)))
                                    fgInputs.RowData(fgInputs.Rows - 1) = i
                                ElseIf IsDigit(strText) Then
                                    ' if default is numeric, then let it be edited like a constant number
                                    strText = Replace(Trim(.Parm(i)), Chr(34), "")
                                    fgInputs.AddItem .ParmName(i) & vbTab & strText
                                    fgInputs.RowData(fgInputs.Rows - 1) = i
                                Else
                                    'fgInputs.AddItem "Symbol" & vbTab & Trim(.Parm(i))
                                    'fgInputs.RowData(fgInputs.Rows - 1) = i
                                End If
                                
                            Case 5 ' bars variable
                                ' if not Market1 and symbol not hard-coded (using double-quotes),
                                ' then add input for user to select the symbol for this bars variable
                                strText = UCase(.ParmName(i))
                                If strText <> "MARKET1" And strText <> "DAILY" And strText <> "WEEKLY" And strText <> "MONTHLY" And Left(strText, 1) <> Chr(34) Then
                                    'strText = StripStr(UCase(.Parm(i)), Chr(34))
                                    If .IsLinkedParm(i) Then
                                        If Not .LinkedParmGet(i, strText, strDontCare) Then strText = strText = StripStr(.Parm(i), Chr(34))
                                    Else
                                        strText = StripStr(.Parm(i), Chr(34))
                                    End If
                                    If strText = "DEFAULT" Then strText = "default"
                                    fgInputs.AddItem .ParmName(i) & vbTab & strText
                                    fgInputs.RowData(fgInputs.Rows - 1) = i
                                End If

                            Case 6 ' constant boolean
                                If .IsLinkedParm(i) Then
                                    If Not .LinkedParmGet(i, strText, strDontCare) Then strText = ""
                                End If
                                
                                If Len(strText) = 0 Then
                                    d = ValOfText(.Parm(i))
                                    If d = 0 Or d < -127 Then
                                        strText = "False"
                                    Else
                                        strText = "True"
                                    End If
                                End If
                                fgInputs.AddItem .ParmName(i) & vbTab & strText
                                fgInputs.RowData(fgInputs.Rows - 1) = i
                            End Select
                        End If
                    Next
                    With fgInputs
                        If .Rows > .FixedRows Then
                            .Col = 1
                            .FillStyle = flexFillRepeat
                            .Select .FixedRows, 1, .Rows - 1, 1
                            .CellFontBold = True
                            .FillStyle = flexFillSingle
                            .Row = 0
                        End If
                        .AutoSize 1
                        .Redraw = flexRDBuffered
                        .ColWidth(0) = .ClientWidth - .ColWidth(1) - 4 * Screen.TwipsPerPixelX
                        '.Redraw = flexRDBuffered
                    End With
                End If
                gdColor.Color = .Color
                txtShiftBars = CStr(Int(.ShiftBars))
                If .Overlayed > 1 Then
                    chkOverlayed(0) = 1
                    chkOverlayed(0).Enabled = False
                Else
                    chkOverlayed(0) = Abs(.Overlayed)
                    chkOverlayed(0).Enabled = True
                End If
                chkOverlayed(1) = chkOverlayed(0).Value
                chkOverlayed(1).Enabled = chkOverlayed(0).Enabled
                'controls for biColor histogram/area charts
                If .DisplayType = eINDIC_Area Or .DisplayType = eINDIC_Histogram Then
                    txtBaseLineY.Text = CStr(.BaseLineY)
                    txtColorSeperator.Text = CStr(.ColorSeperatorVal)
                    chkBiColorBars.Value = .IsBiColorHistogram
                    gdFillColor.Color = .HistogramColorBelow
                End If
            End With
            Set m.EditedIndicator = Indicator
            
        Case kPane:
            vsePane.Tag = "-1"  ' True
            cmdSaveStudy.Enabled = Not m.bSeasonal
            cmdRemove.Enabled = Not m.bSeasonal
            cmdRemove.ToolTipText = "Remove Pane"
            If m.Chart.Tree.Key(idx) = "PRICE PANE" Then
                cmdSaveStudy.ToolTipText = "Save as Study (displayed indicators)"
            Else
                cmdSaveStudy.ToolTipText = "Save as Study (entire pane)"
            End If
            Set Pane = m.Chart.Tree(idx)
            With Pane
                If .Min = 0 And .Max = 0 Then .Max = 100
                txtMin = CStr(.Min)
                txtMax = CStr(.Max)
                
                If .DisplayFormat > 2 Then
                    i = i
                Else
                    optDisplayFormat(.DisplayFormat) = True
                End If
                
                txtDecimals = CStr(.DisplayDecimals)
                chkHideSeparator = .HideSeparator
                If Pane.PricePaneFlag Then
                    chkYscaleLabelAll.Visible = True
                    chkYscaleLabelAll = .YscaleLabelAll
                    chkPriceTopMost.Visible = True
                    chkPriceTopMost.Enabled = True
                    chkPriceTopMost = m.Chart.PriceTopMost
                Else
                    chkYscaleLabelAll.Visible = False
                    chkYscaleLabelAll = 0
                    chkPriceTopMost.Visible = False
                    chkPriceTopMost.Enabled = False
                End If
            End With
            Set m.EditedPane = Pane
            
    End Select
    FixControls

    ' set only the frame being used to visible
    vseGeneral.Visible = Val(vseGeneral.Tag)
    vseXaxis.Visible = Val(vseXaxis.Tag)
    vseSystem.Visible = Val(vseSystem.Tag)
    vseLinkedInputs.Visible = Val(vseLinkedInputs.Tag)
    vseIndicator.Visible = Val(vseIndicator.Tag)
    vseBars.Visible = Val(vseBars.Tag)
    vsePane.Visible = Val(vsePane.Tag)

    Set Indicator = Nothing
    Set IndParent = Nothing
    Set Pane = Nothing
    fgSettings.Redraw = iSaveRedraw

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.EnableTabs", eGDRaiseError_Raise
                
End Sub

Private Sub LoadSettingsGrid()
On Error GoTo ErrSection:

    Dim n&, i&, iEndLoop&, strTemp$, idxNode&, nPane&, strReset$
    Dim nRow&, nOldRows&, nOldRow&
    Dim Pane As cPane
    Dim Indicator As cIndicator
    
    ' default (for now)
    If m.Chart Is Nothing Then
        Set m.Chart = New cChart
        m.Chart.AddStudy "PRICE", True
        'm.Chart.AddStudy "ADX", True
        'm.Chart.AddStudy "STOCHASTIC", True
        'm.Chart.AddStudy "PRO-GO", False
        'm.Chart.AddStudy "RED/GREEN LIGHT", False
        'm.Chart.AddStudy "WILL-VAL", False
        'm.Chart.AddStudy "COT", True
        'm.Chart.AddStudy "MACD", False
        'm.Chart.AddStudy "VOLUME", False
    End If

     With fgSettings
        .Redraw = flexRDNone
        .FillStyle = flexFillRepeat
        
        ' general settings
        nOldRows = .Rows
        nOldRow = .Row
        nRow = .FixedRows
        If nRow >= .Rows Then .Rows = nRow + 1
        CheckedCell(fgSettings, nRow, kShowCol) = True
        .TextMatrix(nRow, kNameCol) = "General Settings"
        .RowData(nRow) = 0
        ColorSettingsRow nRow
        
        ' Xaxis
        nRow = .FixedRows + 1
        If nRow >= .Rows Then .Rows = nRow + 1
        CheckedCell(fgSettings, nRow, kShowCol) = True
        .TextMatrix(nRow, kNameCol) = "Dates Axis"
        '.Cell(flexcpAlignment, nRow, kNameCol) = flexAlignCenterCenter
        .RowData(nRow) = 0
        ColorSettingsRow nRow
        
        ' System
        nRow = .FixedRows + 2
        If nRow >= .Rows Then .Rows = nRow + 1
        CheckedCell(fgSettings, nRow, kShowCol) = True
        .TextMatrix(nRow, kNameCol) = "Trading"
        '.TextMatrix(nRow, kNameCol) = "Trading Strategy"
        '.Cell(flexcpAlignment, nRow, kNameCol) = flexAlignCenterCenter
        .RowData(nRow) = 0
        ColorSettingsRow nRow
        
        ' linked inputs
        nRow = .FixedRows + 3
        If nRow >= .Rows Then .Rows = nRow + 1
        CheckedCell(fgSettings, nRow, kShowCol) = True
        .TextMatrix(nRow, kNameCol) = "Linked Inputs"
        .RowData(nRow) = 0
        ColorSettingsRow nRow
        
        iEndLoop = m.Chart.Tree.Count
        ' for each node in the tree
        For idxNode = 1 To iEndLoop
            If m.Chart.Tree.NodeLevel(idxNode) = 0 Then
                nRow = nRow + 1
                If nRow >= .Rows Then .Rows = nRow + 1
                .RowData(nRow) = idxNode
                ' a Pane
                Set Pane = m.Chart.Tree(idxNode)
                ' see if empty (no indicators)
                If idxNode = m.Chart.Tree.RelativeIndex(idxNode, eTREE_LastDescendant) Then
                    Pane.Display = False
                End If
                nPane = nPane + 1
                If Pane.PricePaneFlag = -1 Then
                    CheckedCell(fgSettings, nRow, kShowCol) = False
                Else
                    CheckedCell(fgSettings, nRow, kShowCol) = Pane.Display
                End If
                '.TextMatrix(nRow, kShowCol) = Pane.Display
                '.TextMatrix(nRow, kNameCol) = Pane.Name
                strTemp = "Pane " & CStr(nPane)
                If m.Chart.Tree.Key(idxNode) = "PRICE PANE" Then
                    strTemp = strTemp & ":  Price"
                ElseIf Len(Pane.Name) > 0 Then
                    strTemp = strTemp & ":  " & Pane.Name
                Else
                    ' get name of first indicator
                    i = m.Chart.Tree.RelativeIndex(idxNode, eTREE_FirstChild)
                    If i > 0 Then
                        Set Indicator = m.Chart.Tree(i)
                        strTemp = strTemp & ":  " & Indicator.Name
                    End If
                End If
                .TextMatrix(nRow, kNameCol) = strTemp
                .Select nRow, 1
            Else
                ' an indicator
                Set Indicator = m.Chart.Tree(idxNode)
                If Indicator.FixCodedName Then
                    ' we only need to show a message if there are non-market inputs
                    ' for this function that have been reset to their defaults
                    For i = Indicator.Inputs.Count To 1 Step -1
                        If Indicator.Inputs.Item(i).ParmTypeID <> 5 Then
                            strReset = strReset & vbCrLf & Indicator.Name
                            Exit For
                        End If
                    Next
                End If
                If Not Indicator.IsAlert Then
                    nRow = nRow + 1
                    If nRow >= .Rows Then .Rows = nRow + 1
                    .RowData(nRow) = idxNode
                    CheckedCell(fgSettings, nRow, kShowCol) = Indicator.Display
                    '.TextMatrix(nRow, kShowCol) = Indicator.Display
                    strTemp = "   "
                    For i = 2 To m.Chart.Tree.NodeLevel(idxNode)
                        strTemp = strTemp & "   "
                    Next
                    .TextMatrix(nRow, kNameCol) = strTemp & Indicator.Name
                End If
            End If
            ColorSettingsRow nRow
        Next
        Set Indicator = Nothing
        Set Pane = Nothing
        
        ' new pane row
        nRow = nRow + 1
        If nRow >= .Rows Then .Rows = nRow + 1
        .RowData(nRow) = 0
        CheckedCell(fgSettings, nRow, kShowCol) = 0
        '.TextMatrix(nRow, kShowCol) = 0
        .TextMatrix(nRow, kNameCol) = "( new chart pane )"
        ColorSettingsRow nRow
        .Cell(flexcpPictureAlignment, nRow, kShowCol) = flexAlignLeftCenter
        
        .Rows = nRow + 1
        If .Rows <> nOldRows Then
            If nOldRow < .FixedRows Then nOldRow = .FixedRows
            If nOldRow >= .Rows Then nOldRow = .Rows - 1
            SelectRows nOldRow
        End If
        
        CollapseRows
        EnableTabs
        
        .FillStyle = flexFillSingle
        .Redraw = flexRDBuffered
    End With

    If Len(strReset) > 0 Then
        InfBox "Some of the function inputs are being reset to their defaults -- please verify the settings for:" _
            & strReset, "!", , "Please Verify Settings"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.LoadSettingsGrid", eGDRaiseError_Raise
        
End Sub

Private Sub InitTemplates()
On Error GoTo ErrSection:

    cboTemplate.AddItem "My Favorite"
    cboTemplate.AddItem "Next Step"
    cboTemplate.AddItem "WST Workshop"
    cboTemplate.AddItem "SpreadNButter"
    cboTemplate.ListIndex = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.InitTemplates", eGDRaiseError_Raise
        
End Sub

Private Sub gdColor_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    With m.EditedIndicator
        .Color = gdColor.Color
        If .IsAutoSwingTrendlines Then
            .UpColor = gdColor.Color
        End If
    End With
    
    SyncPowerZones
    
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdColor.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub ColorSettingsRow(ByVal nRow&)
On Error GoTo ErrSection:

    Dim nType&, bDisplay As Boolean
    
    With fgSettings
        nType = RowType(nRow)
        bDisplay = CheckedCell(fgSettings, nRow, kShowCol)
        .FillStyle = flexFillRepeat
        .Select nRow, 0, nRow, .Cols - 1
        If nType = kGeneral Or nType = kXaxis Or nType = kSystem Or nType = kLinkedInputs Then
            .CellFontBold = True
            .Cell(flexcpPictureAlignment, nRow, 0) = flexAlignLeftCenter
        ElseIf nType = kPane Then
            .CellFontBold = False 'True
            .Cell(flexcpPictureAlignment, nRow, 0) = flexAlignLeftCenter
        Else
            .CellFontBold = False
            .Cell(flexcpPictureAlignment, nRow, 0) = flexAlignRightCenter
        End If
        If nType <> kPane And nType <> kNewPane Then
            ' normal rows (non-Panes)
            .CellBackColor = 0 ' standard
            If bDisplay Then
                .CellForeColor = 0 ' standard
            ElseIf 1 Then
                .CellForeColor = RGB(128, 128, 128) 'darker
            Else
                .CellForeColor = RGB(192, 192, 192) 'lighter
            End If
        ElseIf kInvertedColors Then
            '.CellBackColor = RGB(128, 128, 128)
            '.CellBackColor = &H808000         ' &HFF8080
            .CellBackColor = ALT_GRID_ROW_COLOR
            If bDisplay Then
                '.CellForeColor = vbWhite
                .CellForeColor = 0
            Else
                .CellForeColor = RGB(128, 128, 128)
                '.CellForeColor = vbWhite
            End If
        Else 'light gray background
            .CellBackColor = RGB(192, 192, 192)
            If bDisplay Then
                .CellForeColor = 0 ' standard
            Else
                .CellForeColor = RGB(128, 128, 128)
            End If
        End If
        
        .FillStyle = flexFillSingle
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.ColorSettingsRow", eGDRaiseError_Raise
        
End Sub

Private Sub gdColorBars_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    m.EditedIndicator.Color = gdColorBars.Color
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdColorBars.Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub gdColorUp_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    m.EditedIndicator.UpColor = gdColorUp.Color
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdColor1.Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub gdColorDown_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    m.EditedIndicator.DownColor = gdColorDown.Color
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdColor2.Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub gdColorVolume_POC_Changed()
On Error GoTo ErrSection:

    If Not m.EditedIndicator Is Nothing Then
        If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
            'POC color
            m.EditedIndicator.ProfileColor(ePCStruct_Volume_POC) = gdColorVolume_POC.Color
            m.iGenerateChart = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdColorVolume_POC_Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub gdColorVolume_VA_Changed()
On Error GoTo ErrSection:

    If Not m.EditedIndicator Is Nothing Then
        If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
            'VA color
            m.EditedIndicator.ProfileColor(ePCStruct_Volume_VA) = gdColorVolume_VA.Color
            m.iGenerateChart = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdColorVolume_VA_Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub gdFillColor_Changed()

    If Not m.EditedIndicator Is Nothing Then
        With m.EditedIndicator
            If .DisplayType = eINDIC_Area Or .DisplayType = eINDIC_Histogram Then
                .HistogramColorBelow = gdFillColor.Color
            ElseIf .IsAutoSwingTrendlines Then
                .DownColor = gdFillColor.Color
            Else
                .BoxFilLColor = gdFillColor.Color
            End If
        End With
        m.iGenerateChart = True
    End If

End Sub

Private Sub gdPyramidColorDown_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
        '05-19-2015 - to be implemented for gradient volume profile
        m.EditedIndicator.ProfileColor(ePCStruct_TPO_ColorTo) = gdPyramidColorDown.Color
    Else
        m.EditedIndicator.DownColor = gdPyramidColorDown.Color
    End If
    
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdPyramidColorDown_Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub gdPyramidColorLabel_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
        'POC text color
        m.EditedIndicator.ProfileColor(ePCStruct_TPO_POC) = gdPyramidColorLabel.Color
    Else
        m.EditedIndicator.Color = gdPyramidColorLabel.Color
    End If
    
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdPyramidColorLabel_Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub gdPyramidColorP_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
        'VA text color
        m.EditedIndicator.ProfileColor(ePCStruct_TPO_VA) = gdPyramidColorP.Color
    Else
        m.EditedIndicator.BoxFilLColor = gdPyramidColorP.Color
    End If
    
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdPyramidColorP_Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub gdPyramidColorUp_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    If m.EditedIndicator.DataType = eINDIC_ProfileVolume Then
        'volume profile color
        m.EditedIndicator.ProfileColor(ePCStruct_TPO) = gdPyramidColorUp.Color
    Else
        m.EditedIndicator.UpColor = gdPyramidColorUp.Color
    End If
    
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdPyramidColorUp_Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub gdSelectIcon_Changed()

    Dim nIcon&
    
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    nIcon = gdSelectIcon.Icon
    
    With m.EditedIndicator
        If nIcon = eCNI_Ascii Then
            .MarkerAscii = gdSelectIcon.Ascii
            .MarkerImage = eCNI_Ascii
        ElseIf nIcon < 8 Then
            .MarkerImage = eCNI_Arrow
            Select Case nIcon
                Case 0
                    .MarkerDir = eCNI_North
                Case 1
                    .MarkerDir = eCNI_South
                Case 2
                    .MarkerDir = eCNI_East
                Case 3
                    .MarkerDir = eCNI_West
                Case 4
                    .MarkerDir = eCNI_NorthEast
                Case 5
                    .MarkerDir = eCNI_SouthWest
                Case 6
                    .MarkerDir = eCNI_SouthEast
                Case 7
                    .MarkerDir = eCNI_NorthWest
            End Select
        ElseIf nIcon = 8 Then
            .MarkerImage = eCNI_Plus
        ElseIf nIcon = 9 Then
            .MarkerImage = eCNI_Cross
        ElseIf nIcon = 12 Or nIcon = 17 Then
            .MarkerImage = eCNI_Circle
            If nIcon = 12 Then
                .MarkerFill = 1
            Else
                .MarkerFill = 0
            End If
        ElseIf nIcon = 13 Or nIcon = 18 Then
            .MarkerImage = eCNI_Square
            If nIcon = 13 Then
                .MarkerFill = 1
            Else
                .MarkerFill = 0
            End If
        ElseIf nIcon = 14 Or nIcon = 19 Then
            .MarkerImage = eCNI_Diamond
            If nIcon = 14 Then
                .MarkerFill = 1
            Else
                .MarkerFill = 0
            End If
        Else
            .MarkerImage = eCNI_Triangle
            If nIcon = 10 Or nIcon = 11 Then
                .MarkerFill = 1
            Else
                .MarkerFill = 0
            End If
            If nIcon = 10 Or nIcon = 15 Then
                .MarkerDir = eCNI_North
            Else
                .MarkerDir = eCNI_South
            End If
        End If
    End With
    
    If Not m.bSkipGenerate Then
        m.iGenerateChart = True ' 1
    End If

End Sub

Private Sub gdTrueRange_Changed()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    m.EditedIndicator.trueRangeColor = gdTrueRange.Color
    m.iGenerateChart = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.gdTrueRange.Changed", eGDRaiseError_Show
    Resume ErrExit
            
End Sub

Private Sub mnuAddHorzLine_Click()
On Error GoTo ErrSection:

    Dim idx&, Ind As cIndicator
        
    'get idx of pane
    idx = fgSettings.RowData(fgSettings.Row)
    If idx > 0 Then
        idx = m.Chart.Tree.RelativeIndex(idx, eTREE_Root)
        If idx > 0 Then
            DoEvents
            'create new indicator
            Set Ind = New cIndicator
            With Ind
                .Name = "Horizontal Line"
                .Display = True
                .DisplayType = eINDIC_Line
                .DataType = eINDIC_Constant
                .Style = eINDIC_Default
                .Color = RGB(128, 128, 128)
                .Parm(1) = "0"
            End With
            'add to that pane (make it the last child)
            idx = m.Chart.Tree.Add(Ind, "", idx, eTREE_LastChild)
            LoadSettingsGrid
            SelectIdx idx
            Set Ind = Nothing
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuAddHorzLine.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub mnuAddNewStudy_Click()
On Error GoTo ErrSection:

    AddToChart eAdd_Study

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuAddNewStudy.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub mnuAddHighlightBars_Click()
On Error GoTo ErrSection:

    AddToChart eAdd_HighlightBars, eAddMode2_AttchToInd

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuAddHighlightBars.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub mnuAddSystem_Click()
On Error GoTo ErrSection:

    AddToChart eAdd_System, 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuAddSystem.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub mnuAddToNewPane_Click()
On Error GoTo ErrSection:

    AddToChart eAdd_Indicator, 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuAddToNewPane.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub mnuAddToSelectedIndicator_Click()
On Error GoTo ErrSection:

    AddToChart eAdd_AttachedInd, 2

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuAddToSelectedIndicator.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub mnuAddToSelectedPane_Click()
On Error GoTo ErrSection:

    AddToChart eAdd_Indicator, 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuAddtoSelectedPane.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub mnuCompSymbol_Click()
On Error GoTo ErrSection:

    Dim idx&
    
    If Not m.Chart Is Nothing Then
        idx = AddCompSymbol(m.Chart, False)
        If idx > 0 Then
            LoadSettingsGrid
            SelectIdx idx
            DoEvents
            'MoveFocus gdColor
            MoveFocus gdColorBars
            SendKeys " " '(to dropdown the color pallete)
            m.Chart.RedoMode = eRedo5_RecalcInd     '5442
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuCompSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub mnuNewPaneMove_Click()
On Error GoTo ErrSection:
    
    Dim idx&, idxTo&
    Dim Pane As cPane
    Dim Ind As cIndicator
    
    idx = fgSettings.RowData(fgSettings.Row)
    If idx > 0 Then
        With m.Chart.Tree
            ' verify indicator is "unlinked" and not the Price
            If .NodeLevel(idx) = 1 And .Key(idx) <> "PRICE" Then
                ' move to a new pane (insert as next pane)
                Set Pane = New cPane
                Pane.Display = True
                Pane.Scaling = ePANE_ScaleModeAuto
                idxTo = .RelativeIndex(idx, eTREE_Root) '(current pane)
                idxTo = .Add(Pane, , idxTo, eTREE_NextSibling) '(add new pane after current pane)
                idxTo = .Move(idx, idxTo, eTREE_FirstChild) '(move ind to new pane)
                If idxTo > 0 Then
                    ' make sure it's not overlayed
                    Set Ind = .Item(idxTo)
                    Ind.Overlayed = False
                    If Ind.DataType = eINDIC_BarData Then
                        Pane.DisplayFormat = ePANE_PriceFormat
                    End If
                    Set Ind = Nothing
                End If
                Set Pane = Nothing
            End If
        End With
    End If
    If idxTo > 0 Then
        LoadSettingsGrid
        SelectIdx idxTo
    Else
        Beep
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuNewPaneMove.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub mnuRemove_Click()
On Error GoTo ErrSection:

    cmdRemove_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuRemove.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub mnuSaveStudy_Click()
On Error GoTo ErrSection:

    cmdSaveStudy_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.mnuSaveStudy.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub optAutoScale_Click()
On Error GoTo ErrSection:

    If Not m.EditedPane Is Nothing And optAutoScale Then
        m.EditedPane.Scaling = ePANE_ScaleModeAuto
        FixControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.optAutoScale.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub optAutoScalePrice_Click()
On Error GoTo ErrSection:

    If Not m.EditedPane Is Nothing And optAutoScalePrice Then
        m.EditedPane.Scaling = ePANE_ScaleModeAutoPrice
        FixControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.optAutoScalePrice_Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub optBoxFill_Click(Index As Integer)

    If m.EditedIndicator Is Nothing Then Exit Sub
    
    If optBoxFill(0) = True Then
        m.EditedIndicator.BoxFillStyle = 0
    Else
        m.EditedIndicator.BoxFillStyle = 1
    End If
    m.iGenerateChart = True '1

End Sub

Private Sub optDisplayFormat_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.EditedPane Is Nothing Then Exit Sub
    m.EditedPane.DisplayFormat = Index
    FixControls
    If Index = 2 Then MoveFocus txtDecimals

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.optDisplayFormat.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub optEndOfData_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    m.Chart.ToEndOfData = optEndOfData
    m.Chart.RedoMode = eRedo9_ReloadData
    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.optEndOfData.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub optLineBox_Click(Index As Integer)

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    m.Chart.ProfitLineBox = Index
    m.iGenerateChart = 1

End Sub

Private Sub optManualScale_Click()
On Error GoTo ErrSection:

    If Not m.EditedPane Is Nothing And optManualScale Then
        m.EditedPane.Scaling = ePANE_ScaleModeManual
        FixControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.optManualScale.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub optScaleLinear_Click()
On Error GoTo ErrExit:
    
    If m.EditedPane Is Nothing Then Exit Sub
    If optScaleLinear.Value = False Then Exit Sub
    If m.EditedPane.PaneLogFlag = ePANE_LogFlagLinear Then Exit Sub
    
    m.EditedPane.PaneLogFlag = ePANE_LogFlagLinear
    FixControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.optScaleLinear_Click"

End Sub

Private Sub optScaleLog_Click()
On Error GoTo ErrExit:
    
    If m.EditedPane Is Nothing Then Exit Sub
    If optScaleLog.Value = False Then Exit Sub
    If m.EditedPane.PaneLogFlag = ePANE_LogFlagLog Then Exit Sub
    
    m.EditedPane.PaneLogFlag = ePANE_LogFlagLog
    FixControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.optScaleLog_Click"

End Sub

Private Sub optScalePercent_Click()
On Error GoTo ErrExit:
    
    If m.EditedPane Is Nothing Then Exit Sub
    If optScalePercent.Value = False Then Exit Sub
    If m.EditedPane.PaneLogFlag = ePANE_LogFlagPercent Then Exit Sub
    
    m.EditedPane.PaneLogFlag = ePANE_LogFlagPercent
    FixControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.optScalePercent_Click"

End Sub

Private Sub optTradesAccount_Click()
    FixTradesControls Nothing
End Sub

Private Sub optTradesNone_Click()
    FixTradesControls Nothing
End Sub

Private Sub optTradesStrategy_Click()
    FixTradesControls Nothing
End Sub

Private Sub optPtsPerBar_Click()
    
    m.EditedPane.PointsOrTicksFlag = 0
    FixControls

End Sub

Private Sub optSquareScale_Click()
On Error GoTo ErrSection:

    If m.EditedPane Is Nothing Then Exit Sub
    If optSquareScale.Value = False Then Exit Sub
    If m.EditedPane.Scaling = ePANE_ScaleModeSquare Then Exit Sub
    
    m.EditedPane.Scaling = ePANE_ScaleModeSquare
    m.EditedPane.LogFlagReset
    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.optSquareScale.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optTicksPerBar_Click()
    
    m.EditedPane.PointsOrTicksFlag = 1
    FixControls
    
End Sub

Private Sub optToDate_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    m.Chart.ToEndOfData = optEndOfData
    m.Chart.RedoMode = eRedo9_ReloadData
    FixControls
    MoveFocus dtToDate

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.optToDate.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub picAdd_Click()
On Error GoTo ErrSection:

    cmdAdd_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.picAdd.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub picMoveDown_Click()
On Error GoTo ErrSection:

    cmdMoveDown_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.picMoveDown.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub picMoveUp_Click()
On Error GoTo ErrSection:

    cmdMoveUp_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.picMoveUp.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub SelectIdx(ByVal idxToSelect&)
On Error GoTo ErrSection:

    ' select item if can find it
    Dim bFound As Boolean, bRoot As Boolean
    Dim nRow&, idx&, nSelectRow&, nRowsShowing&
    Dim nTopToShow&, nBottomToShow&, nTopShowing&
    Dim Tree As cGdTree
    Set Tree = m.Chart.Tree
    
    m.nRowWhenLeftForm = 0 '(clear this)
    
    With fgSettings
        For nRow = .FixedRows To .Rows - 1
            idx = .RowData(nRow)
            ' save top and bottom row currently showing
            If .RowIsVisible(nRow) Then
                If nTopShowing = 0 Then nTopShowing = nRow
                nRowsShowing = nRow - nTopShowing + 1
            ElseIf nBottomToShow > 0 And nRowsShowing > 0 Then
                Exit For ' done with everything
            End If
            ' save root before and after
            bRoot = False
            If idx <= 0 Then
                bRoot = True
            Else
                If Tree.NodeLevel(idx) = 0 Then
                    bRoot = True
                End If
            End If
            If bRoot Then
                If bFound Then
                    If nBottomToShow = 0 Then
                        nBottomToShow = nRow
                    End If
                Else
                    nTopToShow = nRow
                End If
            End If
            ' see if row is idxSelect
            If .RowData(nRow) = idxToSelect Then
                bFound = True
                nSelectRow = nRow
            End If
        Next
        If bFound Then
            SelectRows nSelectRow
            ' try to show whole Pane of selected item
            If nBottomToShow < nRowsShowing Then
                .TopRow = .FixedRows
            ElseIf nSelectRow - nTopToShow < nRowsShowing Then
                .TopRow = nTopToShow
            End If
            ' now just make sure selected row is showing
            .ShowCell nSelectRow, 0
        End If
    End With

ErrExit:
    Set Tree = Nothing
    Exit Sub
    
ErrSection:
    Set Tree = Nothing
    RaiseError "frmChartCfg.SelectIdx", eGDRaiseError_Raise
        
End Sub

Private Sub tmrChartCfg_Timer()
On Error GoTo ErrSection:

    Dim i&
    Static bInProgress As Boolean
    
    If bInProgress Then Exit Sub
    bInProgress = True
    
    'see if flagged to generate a chart
    '(done in timer so will not "hang" the control while generating)
    If m.bClickOK Then
        m.bClickOK = False
        cmdOK_Click
    ElseIf m.iGenerateChart <> 0 Then
        If Not m.Chart Is Nothing Then
            If m.iGenerateChart = 1 Then
                ' do for all charts (if globals have been edited)
                UpdateVisibleCharts eRedo3_Settings
            Else
                ' just this chart
                m.Chart.GenerateChart eRedo3_Settings
            End If
        End If
        m.iGenerateChart = False 'clear it
    ElseIf Not Me.Visible And m.bCondFuncNewInprog Then
        If Not FormIsLoaded("frmConditionBuilder") And Not FormIsLoaded("frmFunctionMgrCT") Then
            cmdOK_Click
        End If
    End If
    
ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "frmChartCfg.tmrChartCfg.Timer", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtBarsLeft_Change()

    If m.EditedIndicator Is Nothing Then Exit Sub
    m.EditedIndicator.HighlightBarsLeft = Int(Val(txtBarsLeft.Text))
    m.iGenerateChart = True

End Sub

Private Sub txtBarsRight_Change()

    If m.EditedIndicator Is Nothing Then Exit Sub
    m.EditedIndicator.HighlightBarsRight = Int(Val(txtBarsRight.Text))
    m.iGenerateChart = True

End Sub

Private Sub txtBarWidth_Change()
On Error GoTo ErrSection:

    Dim i&
    
    If Not m.EditedIndicator Is Nothing Then
        If m.Chart.Bars.IsIntraday Then
            i = ValOfText(txtBarWidth.Text)
            If cboVertGrid.Text = "Monthly" Then
                If i > 1 Then
                    InfBox "Profile periodicty on an intraday chart" & vbCrLf & "cannot exceed 1 Monthly"
                    txtBarWidth.Text = "1"
                    GoTo ErrExit
                End If
            ElseIf cboVertGrid.Text = "Weekly" Then
                If i > 4 Then
                    InfBox "Profile periodicty on an intraday chart" & vbCrLf & "cannot exceed 4 Weekly"
                    txtBarWidth.Text = "4"
                    GoTo ErrExit
                End If
            ElseIf cboVertGrid.Text = "Daily" Then
                If i > 30 Then
                    InfBox "Profile periodicty on an intraday chart" & vbCrLf & "cannot exceed 30 Daily"
                    txtBarWidth.Text = "30"
                    GoTo ErrExit
                End If
            End If
        End If
        i = GetPeriodicity(txtBarWidth.Text & " " & cboVertGrid.Text)
        If i < m.Chart.Periodicity Then GoTo ErrExit
        m.EditedIndicator.ProfilePeriodicityStr = txtBarWidth.Text & " " & cboVertGrid.Text
        m.iGenerateChart = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtBarWidth_Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtBaseLineY_Change()
On Error Resume Next:

    If Not m.EditedIndicator Is Nothing Then
        With m.EditedIndicator
            If Not .IsAutoSwingTrendlines Then
                If Len(txtBaseLineY.Text) = 0 Then
                    .BaseLineY = kNullData      'user cleared custom baseline, reset back to null
                Else
                    .BaseLineY = ValOfText(txtBaseLineY.Text)
                End If
            End If
        End With
        m.iGenerateChart = True
    End If

End Sub

Private Sub txtColorSeperator_Change()
On Error Resume Next:

    If Not m.EditedIndicator Is Nothing Then
        With m.EditedIndicator
            If .IsAutoSwingTrendlines Then
                .ExtendTrend = ValOfText(txtColorSeperator.Text)
            Else
                .ColorSeperatorVal = ValOfText(txtColorSeperator.Text)
            End If
        End With
        m.iGenerateChart = True
    End If

End Sub

Private Sub txtFunction_Change()
On Error GoTo ErrSection:

    Dim strName$

    If m.EditedIndicator Is Nothing Then Exit Sub
    txtFunction.Tag = "DIRTY"
    strName = txtFunction
    If Len(strName) = 0 Then
        If m.EditedIndicator.FunctionID > 0 Then
            On Error Resume Next
            strName = g.Functions.Item(CStr(m.EditedIndicator.FunctionID)).FunctionName
        End If
        If Len(strName) = 0 Then strName = "Function"
    End If
    m.EditedIndicator.Name = strName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtFunction.Change", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtFunction_GotFocus()
On Error GoTo ErrSection:

    txtFunction.Tag = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtFunction.GotFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtFunction_LostFocus()
On Error GoTo ErrSection:

    Dim nRow&

    If m.EditedIndicator Is Nothing Then Exit Sub
    'If txtFunction <> m.EditedIndicator.Name Then
    If txtFunction.Tag = "DIRTY" Then
        'm.EditedIndicator.Name = txtFunction
        nRow = fgSettings.Row
        LoadSettingsGrid
        fgSettings.Row = nRow
    End If
    txtFunction.Tag = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtFunction.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtDecimals_Change()
On Error GoTo ErrSection:

    If m.EditedPane Is Nothing Then Exit Sub
    m.EditedPane.DisplayDecimals = ValOfText(txtDecimals)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtDecimals.Change", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtDecimals_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtDecimals

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtDecimals.GotFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtForecastBars_Change()
On Error GoTo ErrSection:
    
    If fgSettings.Redraw = flexRDNone Then Exit Sub
        
    If ValOfText(txtForecastBars) <> m.Chart.BlankBars Then
        m.Chart.ResetLastScreenDate
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtForecastBars.Change", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtForecastBars_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtForecastBars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtForecastBars.GotFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtForecastBars_LostFocus()
On Error GoTo ErrSection:
   
    Dim dBars#, iBars&
   
    'aardvark 3295 fix:
    'need a double for users inputing value a long cannot accomodate (e.g. 999999999999)
    dBars = ValOfText(txtForecastBars)
    If dBars <= 0 Then
        iBars = 1
        txtForecastBars = iBars
    ElseIf dBars > 500 Then
        iBars = 500
        txtForecastBars = iBars
    Else
        iBars = Int(dBars)
    End If
    
    If iBars <> m.Chart.BlankBars Then
        m.Chart.BlankBars(Me) = iBars
        m.Chart.RedoMode = eRedo9_ReloadData
        'BenchMark
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtForecastBars.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtMax_Change()
On Error GoTo ErrSection:

    If Not m.EditedPane Is Nothing Then
        With m.EditedPane
            If .PricePaneFlag = 1 Then
                .Max = m.Chart.Bars.PriceFromString(txtMax)
            Else
                .Max = ValOfText(txtMax)
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtMax.Change", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtMaxDays_LostFocus()
On Error GoTo ErrSection:

    Dim dDays#
    Static nMaxAlreadyWarnedThisSession&
    
    dDays = ValOfText(txtMaxDays)
    If dDays > 9999 Then
        dDays = 9999
    ElseIf dDays <= 0 Then
        dDays = 10
    End If
    
    ' warn them if they're trying to increase it and it's > 10
    If dDays > 10 And dDays > m.Chart.MaxIntradayDays And dDays > nMaxAlreadyWarnedThisSession Then
        If InfBox("Please be aware that loading more intraday data can increase the time for loading the chart.", "!", "+OK|-Cancel", "Warning") = "C" Then
            dDays = m.Chart.MaxIntradayDays
        End If
        nMaxAlreadyWarnedThisSession = Round(dDays * 1.5) ' allow a higher number before warning again
    End If
    
    txtMaxDays = CStr(dDays)
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If dDays <> m.Chart.MaxIntradayDays Then
        m.Chart.MaxIntradayDays = dDays
        m.Chart.RedoMode = eRedo9_ReloadData
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtMaxDays.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtMin_Change()
On Error GoTo ErrSection:

    If Not m.EditedPane Is Nothing Then
        With m.EditedPane
            If .PricePaneFlag = 1 Then
                .Min = m.Chart.Bars.PriceFromString(txtMin)
            Else
                .Min = ValOfText(txtMin)
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtMin.Change", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtMin_LostFocus()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtMin.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtPerBar_Change()
On Error GoTo ErrSection:

    Dim i&, s$
    
    s = Trim(txtPerBar)
    For i = 1 To Len(s)
        If Not IsDigit(s, i) Then
            InfBox "Please use only positive integers for number of bars."
            txtPerBar = StripStr(s, Mid(s, i, 1))
            Exit Sub
        End If
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtPerBar.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtPerBar_LostFocus()
On Error GoTo ErrSection:

    SaveSquareTicksBar
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtPerBar.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtPercentVolume_VA_LostFocus()
On Error GoTo ErrSection:

    Dim str1$
    
    If Not m.EditedIndicator Is Nothing Then
        m.EditedIndicator.ProfileParm(ePCStruct_Volume_VA) = ValOfText(txtPercentVolume_VA.Text)
        m.iGenerateChart = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtPercentVolume_VA_LostFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtPtsOrTicks_LostFocus()
On Error GoTo ErrSection:

    SaveSquareTicksBar

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtPointsPerBar.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtShiftBars_Change()
On Error GoTo ErrSection:
    
    If fgSettings.Redraw = flexRDNone Then Exit Sub
    If m.EditedIndicator Is Nothing Then Exit Sub
    
    m.EditedIndicator.ShiftBars = ValOfText(txtShiftBars)
    
    ClearCodedText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.txtShiftBars.Change", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtPoints_LostFocus()

    If Not m.EditedIndicator Is Nothing Then
        m.EditedIndicator.TicksPerRow = ValOfText(txtPoints.Text)
        m.iGenerateChart = True
    End If

End Sub

Private Sub vseSave_Click()
On Error GoTo ErrSection:

    Dim strTemp$
    Dim sc1 As cPane
    Dim sc2 As cPane
    
    Set sc2 = m.Chart.Tree(1)
    Set sc1 = sc2.MakeCopy
    sc1.Name = "Testingg"
    strTemp = sc2.Name
    strTemp = sc1.Name

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.vseSave.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub vseSaveAs_Click()
On Error GoTo ErrSection:

    m.Chart.RedoMode = eRedo9_ReloadData
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.vseSaveAs.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub SelectRows(Optional ByVal nRow& = -1)
On Error GoTo ErrSection:

    Dim nEndRow&
    
    With fgSettings
        If nRow < 0 Then nRow = .Row
        If nRow >= 0 Then
            nEndRow = nRow + 1
            If RowType(nRow) = kPane Then
                ' select all rows with this Pane
                Do While nEndRow < .Rows
                    If RowType(nEndRow) <> kIndicator Then Exit Do
                    nEndRow = nEndRow + 1
                Loop
            ElseIf .IsSelected(nRow) Then
                If RowType(.Row) = kIndicator And RowType(nEndRow) = kIndicator Then
                    For nEndRow = .Row To .Rows - 1
                        If RowType(nEndRow) <> kIndicator Or Not .IsSelected(nEndRow) Then Exit For
                    Next
                End If
            End If
            .Select nRow, 0, nEndRow - 1, .Cols - 1
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.SelectRows", eGDRaiseError_Raise
        
End Sub

Private Sub ShowAddPopup(ByVal bAddButton As Boolean)
On Error GoTo ErrSection:

    Dim idx&, Ind As cIndicator
    
    If m.Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then Exit Sub

    'enable items
    mnuRemove.Enabled = True
    mnuAddToSelectedIndicator.Enabled = False
    mnuAddToSelectedPane.Enabled = True
    mnuAddHighlightBars.Enabled = False
    mnuAddHorzLine.Enabled = True
    mnuNewPaneMove.Enabled = False
    mnuNewPaneMove.Visible = False
    mnuSaveStudy.Visible = False
    Select Case RowType(fgSettings.Row)
        Case kIndicator
            mnuRemove.Caption = "&Remove Indicator"
            idx = fgSettings.RowData(fgSettings.Row)
            With m.Chart.Tree
                Set Ind = .Item(idx)
                If Ind.DataType <> eINDIC_Constant Then
                    mnuAddToSelectedIndicator.Enabled = True
                    If .NodeLevel(idx) = 1 And UCase(.Key(idx)) <> "PRICE" Then
                        mnuNewPaneMove.Enabled = True
                    End If
                    If Ind.DataType <> eINDIC_BooleanArray And Ind.Display = True Then
                        mnuAddHighlightBars.Enabled = True
                    End If
                End If
                Set Ind = Nothing
            End With
            mnuNewPaneMove.Visible = True
        Case kPane
            idx = fgSettings.RowData(fgSettings.Row)
            mnuRemove.Caption = "&Remove Pane"
            mnuSaveStudy.Visible = True
            mnuSaveStudy.Caption = cmdSaveStudy.ToolTipText
        Case Else
            mnuRemove.Enabled = False
            mnuAddToSelectedPane.Enabled = False
    End Select

    'show popup
    If bAddButton Then
        mnuSepRemove.Visible = False
        mnuRemove.Visible = False
        mnuNewPaneMove.Visible = False
        mnuSaveStudy.Visible = False
        mnuAddSystem.Visible = HasGold(False, , False)
        If 1 Then
            Me.PopupMenu mnuAdd, vbPopupMenuLeftAlign, _
                vseAdd.Left + cmdAdd.Left, _
                vseAdd.Top + cmdAdd.Top + cmdAdd.Height
        Else
            Me.PopupMenu mnuAdd, vbPopupMenuLeftAlign, _
                fgSettings.Left + fgSettings.Width / 2, _
                fgSettings.Top + fgSettings.Height / 2
        End If
    Else
        mnuSepRemove.Visible = True
        mnuRemove.Visible = True
        mnuAddSystem.Visible = False

        Me.PopupMenu mnuAdd
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.ShowAddPopup", eGDRaiseError_Raise
        
End Sub

Private Function DisplayTypeForInd(Ind As cIndicator) As eIndicatorDisplayType
On Error GoTo ErrSection:

    Dim eType As eIndicatorDisplayType

    eType = eINDIC_NoStyle

    If Not Ind Is Nothing Then
        If UCase(Ind.CodedName) = "PYRAMIDTP" Then
            eType = eINDIC_ArtPyramid
            Ind.IndLabelMode = eINDIC_NoValue
        ElseIf UCase(Ind.CodedName) = "REVERSEBARTP" Then
            eType = eINDIC_ArtReversal
            Ind.IndLabelMode = eINDIC_NoValue
        ElseIf Ind.DataType = eINDIC_Array Or Ind.DataType = eINDIC_DrawCommands Then
            eType = eINDIC_Line  'aardvark 933 fix
        ElseIf Ind.DisplayType = eINDIC_BollingerBar Or Ind.UpDownColorFlag <> 0 Then
            'Bollinger Bar indicators cannot have highlight bars
            If Ind.DisplayType = eINDIC_BollingerBar Or Ind.UpDownColorFlag = 1 Then
                eType = eINDIC_HighlightMarkers
            Else
                Ind.UpDownColorFlag = 0
                eType = eINDIC_HighlightBars
            End If
        Else
            eType = eINDIC_HighlightBars
        End If
    End If

ErrExit:
    DisplayTypeForInd = eType
    Exit Function
    
ErrSection:
    RaiseError "frmChartCfg.DisplayTypeForInd"

End Function

Private Function ProcessNewInd(ByRef Ind As cIndicator, _
    ByRef idx As Long, _
    ByRef eListType As eChartAddListType, _
    ByVal iMode As eChartAddListMode, _
    ByVal strItem As String, _
    ByVal strExpression As String, _
    ByVal strName As String) As Boolean
On Error GoTo ErrSection:

    Dim ParentInd As cIndicator
    Dim Pane As cPane
    
    Dim bCustom As Boolean
    
    Select Case eListType
        Case eAdd_Indicator, eAdd_AttachedInd, eAdd_HighlightBars
            Set Ind = New cIndicator
            
            If eListType = eAdd_HighlightBars Then
                Ind.DataType = eINDIC_BooleanArray
            Else
                Ind.DataType = eINDIC_Array
            End If
            
            If strItem = "..." Then
                bCustom = True
                Ind.InputsFrame = 2
                
                Set m.EditedIndicator = Ind
                
                If Len(strExpression) = 0 Then
                    'indicator datatype will change if user chose to add a custom indicator,
                    'but writes a function that returns array of boolean (or vice versa)
                    cmdEditFunction_Click
                Else
                    Ind.Expression = strExpression
                End If
            End If
        
        Case Else
            GoTo ErrExit
    
    End Select
    
    If Ind Is Nothing Then GoTo ErrExit
    
    Ind.Display = True
    
    If Not bCustom Then
        Ind.CodedName = strItem
    ElseIf Len(strName) > 0 Then
        Ind.Name = strName
    ElseIf Ind.DataType = eINDIC_BooleanArray Then
        Ind.Name = "Custom HighlightBars"
    Else
        Ind.Name = "Custom Indicator"
    End If
    
    If Ind.DataType = eINDIC_BooleanArray Then
        'set index of where to add this new in indicator
        Select Case iMode
            Case eAddMode0_NewPane:
                'new pane: attach to price indicator
                idx = m.Chart.Tree.Index("PRICE")
                Set ParentInd = m.Chart.Tree(idx)
            Case eAddMode1_SelectedPane:
                'get idx of selected pane
                idx = fgSettings.RowData(fgSettings.Row)
                idx = m.Chart.Tree.RelativeIndex(idx, eTREE_Root)
                'get idx of first child
                idx = m.Chart.Tree.RelativeIndex(idx, eTREE_FirstChild)
                Set ParentInd = m.Chart.Tree(idx)
                If Not ParentInd Is Nothing Then ParentInd.Display = True
            Case eAddMode2_AttchToInd:
                'get idx of selected indicator
                idx = fgSettings.RowData(fgSettings.Row)
                Set ParentInd = m.Chart.Tree(idx)
                If Not ParentInd Is Nothing Then ParentInd.Display = True
        End Select
        
        'the only time we care about the actual parent indicator is when it is price because
        'we have to check to see if the price bar display is up/down color or bolinger bars
        If UCase(ParentInd.Name) <> "PRICE" Then Set ParentInd = Ind
    
        Ind.DisplayType = DisplayTypeForInd(ParentInd)
        
        'processing specific to boolean indicator
        Ind.MarkerPrompt = True
        strItem = " " & UCase(Ind.Name) & " "
        If InStr(strItem, " SHORT ") > 0 Or InStr(strItem, " SELL ") > 0 Or InStr(strItem, " SELLING ") > 0 Then
            Ind.Color = g.ChartGlobals.nShortColor
        ElseIf InStr(strItem, " LONG ") > 0 Or InStr(strItem, " BUY ") > 0 Or InStr(strItem, " BUYING ") > 0 Then
            Ind.Color = g.ChartGlobals.nLongColor
        Else
            Ind.Color = vbBlue
        End If
    
        eListType = eAdd_HighlightBars
    Else
        
        If Ind.DataType = eINDIC_DrawCommands Then
            'JM 08-28-2014: for now TAS Market Map is the only indicator of this type, just add it to price pane
            idx = m.Chart.Tree.Index("PRICE PANE")
        Else
            'set index of where to add this new in indicator
            Select Case iMode
                Case eAddMode0_NewPane:
                    'create new pane
                    Set Pane = New cPane
                    With Pane
                        .Display = True
                        .Scaling = ePANE_ScaleModeAuto
                    End With
                    idx = m.Chart.Tree.Add(Pane, "", 1, eTREE_LastSibling)
                Case eAddMode1_SelectedPane:
                    'get idx of selected pane
                    idx = fgSettings.RowData(fgSettings.Row)
                    idx = m.Chart.Tree.RelativeIndex(idx, eTREE_Root)
                Case eAddMode2_AttchToInd:
                    'get idx of selected indicator
                    idx = fgSettings.RowData(fgSettings.Row)
            End Select
        End If
        
    
        Ind.DisplayType = DisplayTypeForInd(Ind)
        Ind.Color = vbBlue
        
        eListType = eAdd_Indicator
    End If
            
ErrExit:
    ProcessNewInd = bCustom
    Exit Function

ErrSection:
    ProcessNewInd = False
    RaiseError "frmChartCfg.ProcessNewInd"

End Function

Public Function AddToChart(Optional ByVal eListType As eChartAddListType = eAdd_Previous, _
        Optional ByVal iMode As eChartAddListMode = eAddMode0_NewPane, _
        Optional ByVal bHideWhenDone As Boolean = False, _
        Optional ByVal strExpression As String = "", _
        Optional ByVal strName As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bCustom As Boolean
    Dim bAdded As Boolean
    
    Dim i&, r&
    Dim strAdd$, strItem$
    Dim idx&, idxNew&, idxCustom&
    
    Dim Pane As cPane
    Dim Ind As cIndicator
    
    Dim eData As eIndicatorDataType
    Dim eDisplay As eIndicatorDisplayType
    
    If m.Chart Is Nothing Then Exit Function
    
    ' set this flag so "UnloadEditors" will not unload this
    ' form prematurely (can happen when in "hidden" mode)
    bNowAdding = True
    DoEvents
    
    ' Current grapheng.dll zoom code/design does not accomodate changes
    ' to the pane structures. This can be changed when/if desired.
    If m.Chart.Zoomed = True Then m.Chart.UnzoomChart True
    
    If eListType = eAdd_AttachedInd Then iMode = eAddMode2_AttchToInd
    
    If Not Me.Visible Then
        Set m.EditedIndicator = Nothing         '5635
        Set m.EditedPane = Nothing
    End If
    
    Select Case iMode
        Case eAddMode2_AttchToInd
            strItem = "Select indicator(s) to 'attach' to existing indicator:"
        Case eAddMode1_SelectedPane
            strItem = "Select indicator(s) to add to existing pane:"
        Case eAddMode4_InstReplay
            iMode = eAddMode0_NewPane      '5117 - iMode is 4 only from "A" key of chart & chart is in gamemode
            strItem = "ChartIsInGameMode"
        Case eAddMode3_CondBuilder, eAddMode5_NewFunction
            If Len(strExpression) = 0 Then
                strAdd = "Error"
            Else
                strAdd = "..."
            End If
        Case Else
            strItem = ""
    End Select
            
    If Len(strAdd) = 0 Then
        strAdd = frmAddToChart.ShowMe(eListType, strItem)
        If strAdd = "ShowFunctionEditor" Then
            frmFunctionMgrCT.ShowMe 0, , , , , Me
            tmrChartCfg.Enabled = False
            m.bCondFuncNewInprog = True
            GoTo ErrExit
        ElseIf strAdd = "ShowCondBuilder" Then
            frmConditionBuilder.ShowMe m.Chart, , , Me
            tmrChartCfg.Enabled = False
            m.bCondFuncNewInprog = True
            GoTo ErrExit
        ElseIf strAdd = "ShowCondBuilder..." Then
            frmConditionBuilder.ShowMe m.Chart, , eType_HighlightBars, Me
            tmrChartCfg.Enabled = False
            m.bCondFuncNewInprog = True
            GoTo ErrExit
        End If
        
        DoEvents
    ElseIf strAdd = "Error" Then
        strAdd = ""
    ElseIf Not m.EditedIndicator Is Nothing Then
        iMode = eAddMode2_AttchToInd
    ElseIf Not m.EditedPane Is Nothing Then
        iMode = eAddMode1_SelectedPane
    Else
        iMode = eAddMode0_NewPane
    End If
            
    For i = 1 To 99999
        bCustom = False        'fix for aardvark 1135
        strItem = Parse(strAdd, vbTab, i)
        If Len(strItem) = 0 Then
            If i > 1 Then
                With fgSettings
                    r = .Rows
                    LoadSettingsGrid
                    If .Rows > r Then
                        Select Case eListType
                            Case eAdd_Study:
                                SelectIdx .RowData(r)
                            Case Else
                                SelectIdx idxNew
                        End Select
                    End If
                End With
                m.iGenerateChart = True
            End If
            Exit For
        End If
        
        Select Case eListType
            Case eAdd_System
                m.Chart.SystemID = Val(strItem)
                m.Chart.ShowTrades = True
                LoadSystemInfo
                fgSettings.Row = fgSettings.FixedRows + 2
                Exit For
        
            Case eAdd_Study
                bAdded = m.Chart.TemplateAddStudy(strItem)
            
            Case eAdd_Indicator, eAdd_AttachedInd, eAdd_HighlightBars
                Set Ind = Nothing
                
                bCustom = ProcessNewInd(Ind, idx, eListType, iMode, strItem, strExpression, strName)
                
                If idx < 1 Or idx > m.Chart.Tree.Count Then
                    Beep
                ElseIf Not Ind Is Nothing Then
                    idxNew = m.Chart.Tree.Add(Ind, "", idx, eTREE_LastChild)
                    If bCustom Then idxCustom = idxNew
                    bAdded = True
                End If
        End Select
    Next
    
    If idxCustom > 0 Then
        SelectIdx idxCustom
        DoEvents
        If Not m.EditedIndicator Is Nothing Then
            If m.EditedIndicator.Expression = "" Then
                m.Chart.Tree.Remove idxCustom
                LoadSettingsGrid
                bAdded = False
            End If
        End If
    End If
        
    Set Ind = Nothing
    Set Pane = Nothing
    
    If bHideWhenDone Then
        cmdOK_Click
    ElseIf Me.Visible And bAdded Then
        Do While m.iGenerateChart
            Sleep 0
        Loop
        DoEvents
        MoveFocus gdColor
        If Not m.EditedIndicator Is Nothing Then
            If m.EditedIndicator.DisplayType <> eINDIC_ArtPyramid And m.EditedIndicator.DisplayType <> eINDIC_ArtReversal Then
                SendKeys " " '(to dropdown the color pallete)
            End If
        Else
            SendKeys " " '(to dropdown the color pallete)
        End If
    End If
        
    bNowAdding = False
    AddToChart = bAdded

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmChartCfg.AddToChart", eGDRaiseError_Raise
        
End Function

'Call this from a chart to load and show the form
Public Sub ShowMe(pChart As cChart, Optional ByVal nIdxToSelect& = 0, Optional ByVal bHidden As Boolean = False)
On Error GoTo ErrSection:

    Dim bGameModeLock As Boolean
    
    m.bCondFuncNewInprog = False    'reset
    
    If FormIsLoaded("frmEditAnnot") Then Unload frmEditAnnot
    
    If Not pChart Is Nothing Then Set Chart = pChart
    
    If nIdxToSelect > 0 Then
        SelectIdx nIdxToSelect
    ElseIf nIdxToSelect = -1 Then
        ' show trading system
        fgSettings.Row = fgSettings.FixedRows + 2
    ElseIf m.nRowWhenLeftForm > 0 And m.nRowWhenLeftForm < fgSettings.Rows Then
        fgSettings.Row = m.nRowWhenLeftForm
    ElseIf m.nRowWhenLeftForm = 0 Then
        fgSettings.Row = fgSettings.FixedRows
    End If
        
    If Not bHidden Then
        If Not m.Chart Is Nothing Then
            If Not m.Chart.Form.GameMode Is Nothing Then
                If m.Chart.Form.GameMode.CustomOrders > 0 Then
                    bGameModeLock = True
                End If
            End If
            If Not m.Chart.Form.IsInGameMode And Not m.bSeasonal Then
                ' first save existing template and any new price pane annots
                ' (so can restore old settings if cancelled)
                m.Chart.TemplateSave
            End If
            
            fraDates.Enabled = Not bGameModeLock
            cmdBarPeriod.Enabled = Not bGameModeLock
            cboBarPeriod.Enabled = Not bGameModeLock
            cmdSelectSystem.Enabled = Not bGameModeLock
            cmdSymbol.Enabled = Not bGameModeLock
            
            If m.bSeasonal Then fgSettings.Row = fgSettings.FixedRows
        End If
        
        SetCaption
        ShowForm Me             'JM 01-11-2016, ALT_GRID_ROW color is used only for panes
        MoveFocus fgSettings
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.ShowMe", eGDRaiseError_Show
        
End Sub

Private Sub FixControls()
On Error GoTo ErrSection:

    Dim Indicator As cIndicator
    Dim idx&, bEnable As Boolean
        
    If chkGradient.Value = 0 Then
        clrChartGradient.Visible = False
        chkGradient.Width = 2600
        chkGradient.Caption = "Use gradient background color"
        imgGradient.Visible = False
        imgGradientWhite.Visible = False
    Else
        clrChartGradient.Visible = True
        chkGradient.Caption = "Use gradient color:"
        chkGradient.Width = 1700
        If g.nColorTheme = kDarkThemeColor Then
            imgGradientWhite.Visible = True
        Else
            imgGradient.Visible = True
        End If
    End If
    
    If Not gdColor.Visible Then gdColor.Visible = True
    If fgColors.Visible Then fgColors.Visible = False
    ToggleFakeDropdown False
    
    If Not m.EditedIndicator Is Nothing Then
        'temporarily set m.EditedIndicator to Nothing
        'so controls being set will not trigger other things
        Set Indicator = m.EditedIndicator
        Set m.EditedIndicator = Nothing
        idx = fgSettings.RowData(fgSettings.Row)
        lblWidth.Caption = "Width:"
        FixCtlIndicator Indicator, idx
        Set m.EditedIndicator = Indicator
        Set Indicator = Nothing
    ElseIf Not m.EditedPane Is Nothing Then
        FixCtlPane
    ElseIf Not m.Chart Is Nothing Then
        If optToDate Then
            Enable dtToDate
        Else
            Disable dtToDate
            If Not optEndOfData Then optEndOfData = True
        End If
        If m.Chart.Periodicity < ePRD_Days Then
            bEnable = True
        Else
            bEnable = False
        End If
        
        SetStartStopTimesText
        
        Enable txtMaxDays, bEnable
        Enable lblMaxDays1, bEnable
        Enable lblMaxDays2, bEnable
        
        Enable lblStartStopInfo1, bEnable
        Enable lblStartStopInfo2, bEnable
        Enable lblStartStopTimes, bEnable
        
        If m.Chart.SymbolID <= 0 Then
            bEnable = False
        ElseIf Not m.Chart.Form.GameMode Is Nothing Then
            If m.Chart.Form.GameMode.CustomOrders > 0 Then
                bEnable = False        '4776
            End If
        End If
        Enable cmdStartStop, bEnable
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixControls", eGDRaiseError_Raise
        
End Sub

Private Sub ShowChartColors()
On Error GoTo ErrSection:

    Dim i&, bCustom As Boolean
    
    ' TLB: first set the CustomColor to our normal default
    ' (it will get overridden below if user has set it to their own custom color)
    clrChartGradient.CustomColor = GradientDefault
    
    If Not m.Chart Is Nothing Then
        If m.Chart.UseCustomColors Then bCustom = True
    End If
    If bCustom Then
        chkCustomColors = 1
        fraColors.Caption = "Custom Colors (this chart)"
        With m.Chart
            If .ChartBackColor = 0 And .ChartForeColor = 0 _
                And .BorderBackColor = 0 And .BorderForeColor = 0 Then
                ' first time: use defaults
                .ChartBackColor = g.ChartGlobals.nChartBackColor
                .ChartForeColor = g.ChartGlobals.nChartForeColor
                .BorderBackColor = g.ChartGlobals.nBorderBackColor
                .BorderForeColor = g.ChartGlobals.nBorderForeColor
            End If
            clrChartBack.Color = .ChartBackColor
            clrChartFore.Color = .ChartForeColor
            clrBorderBack.Color = .BorderBackColor
            clrBorderFore.Color = .BorderForeColor
            clrChartGradient.Color = .ChartGradientColor
            chkGradient.Value = .UseGradient
        End With
    Else
        chkCustomColors = 0
        fraColors.Caption = "Default Colors (all charts)"
        With g.ChartGlobals
            clrChartBack.Color = .nChartBackColor
            clrChartFore.Color = .nChartForeColor
            clrBorderBack.Color = .nBorderBackColor
            clrBorderFore.Color = .nBorderForeColor
            clrChartGradient.Color = .nChartGradientColor
            chkGradient.Value = .nUseGradient
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.ShowChartColors", eGDRaiseError_Raise
        
End Sub

Private Sub SetCaption()
On Error Resume Next

    Dim i&, strText$
    
    strText = Trim(Replace(m.Chart.Form.vseCaption.Caption, "&&", "&"))
    i = InStr(strText, ")")
    If i > 0 Then strText = Left(strText, i)
    strText = "Chart Settings:  " & StripStr(strText, ":")
    If Me.Caption <> strText Then Me.Caption = strText

End Sub

Private Sub LoadSystemInfo()
On Error GoTo ErrSection:

    Dim strName As String
    Dim lID As Long

    If m.Chart.SystemID > 0 And HasGold(False, , False) Then
        lID = m.Chart.SystemID
        SyncSystemInfo strName, lID
    End If
        
    If Len(strName) > 0 Then
        lblSystemName = strName
    Else
        lblSystemName = "" '"(click 'Select Strategy')"
    End If
    
    FixTradesControls m.Chart
   
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.LoadSystemInfo", eGDRaiseError_Raise
        
End Sub

Private Sub CheckPaneDisplay(ByVal nIndIdx&, ByVal bIndIsVisible As Boolean)
On Error GoTo ErrSection:

    Dim nParentIdx&, nVisibleInd&, i&
    Dim Pane As cPane
    Dim Ind As cIndicator
    Dim nParentRow&
    
    nParentIdx = m.Chart.Tree.AncestorIndex(nIndIdx, 0)
    If m.Chart.Tree.NodeLevel(nParentIdx) <> 0 Then
        Exit Sub        'tree/indices problems, can't do anything
    End If
    
    Set Pane = m.Chart.Tree(nParentIdx)
    If Pane Is Nothing Then Exit Sub    'precautionary, should never happen
    Pane.LogFlagReset
    If bIndIsVisible Then Exit Sub
    
    nParentRow = -1
    'find row that parent's pane is in
    With fgSettings
        For i = .Row To .FixedRows Step -1
            If .RowData(i) = nParentIdx Then
                nParentRow = i
                Exit For
            End If
        Next
        'count visible indicators
        If nParentRow > .FixedRows And nParentRow < .Rows Then
            For i = nParentRow + 1 To .Rows - 1
                If RowType(i) = kPane Then
                    Exit For
                ElseIf RowType(i) = kIndicator Then
                    If m.Chart.Tree.NodeLevel(.RowData(i)) <> 0 Then
                        Set Ind = m.Chart.Tree(.RowData(i))
                        If Not Ind Is Nothing Then
                            If Ind.Display Then
                                nVisibleInd = 1
                                Exit For
                            End If
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With

    If nVisibleInd = 0 And nParentRow > 0 Then
        If RowType(nParentRow) = kPane Then
            CheckedCell(fgSettings, nParentRow, kBoxCol) = False
            fgSettings_AfterEdit nParentRow, kBoxCol
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.LoadSystemCombo", eGDRaiseError_Raise
End Sub

' to clear coded text for a node and its descendendants
' (can pass -1 for entire tree or 0 for current node)
Private Sub ClearCodedText(Optional ByVal idxSubtree& = 0)
On Error GoTo ErrSection:

    Dim idx&, idxEnd&
    Dim Ind As cIndicator
    
    If m.Chart Is Nothing Then Exit Sub
    
    If idxSubtree = 0 Then
        With fgSettings
            If .Row >= .FixedRows And .Row < .Rows Then
                idxSubtree = .RowData(.Row)
            End If
        End With
    End If
    
    If idxSubtree > 0 And idxSubtree <= m.Chart.Tree.Count Then
        ' just for this subtree (node and all descendents)
        idxEnd = m.Chart.Tree.RelativeIndex(idxSubtree, eTREE_LastDescendant)
    Else
        ' for entire tree
        idxSubtree = 1
        idxEnd = m.Chart.Tree.Count
    End If
        
    ' clear coded text for all indicators in this part of the tree
    ' (so will get rebuilt next time chart gets generated)
    For idx = idxSubtree To idxEnd
        If TypeOf m.Chart.Tree(idx) Is cIndicator Then
            Set Ind = m.Chart.Tree(idx)
            If Not Ind.IsAlert Then
                Ind.UpdateAlertParms        '6580
                Ind.CodedText = ""
            End If
        End If
    Next
    Set Ind = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.ClearCodedText", eGDRaiseError_Raise
End Sub

Private Sub SaveSquareTicksBar()

    Dim nPerBar&, dTicks#
    
    If m.EditedPane Is Nothing Then Exit Sub
    
    nPerBar = Int(ValOfText(txtPerBar.Text))
    dTicks = ValOfText(txtPtsOrTicks.Text)
    
    If dTicks <= 0 Or nPerBar <= 0 Then
        FixControls
        Exit Sub
    End If
    
    If optTicksPerBar = True Then
        dTicks = m.Chart.Bars.PriceFromString(txtPtsOrTicks.Text)
    Else
        dTicks = ValOfText(txtPtsOrTicks.Text) / m.Chart.Bars.Prop(eBARS_TickMove)
    End If
    
    m.EditedPane.SquareTicks(m.Chart) = dTicks
    m.EditedPane.SquareBars(m.Chart) = nPerBar
    
    FixControls

End Sub

Private Sub clrLoss_Changed()
On Error GoTo ErrSection:

    g.ChartGlobals.nLossColor = clrLoss.Color
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrLong.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub clrWin_Changed()
On Error GoTo ErrSection:

    g.ChartGlobals.nWinColor = clrWin.Color
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.clrShort.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cboProfitLines_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    g.ChartGlobals.eProfitLineStyle = cboProfitLines.ListIndex + 1
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.cboLineDefault.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkChartTips_Click()
On Error GoTo ErrSection:

    If fgSettings.Redraw = flexRDNone Then Exit Sub
    g.ChartGlobals.bChartTips = -chkChartTips
    m.iGenerateChart = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.chkTips.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub FixCtlPane()
On Error GoTo ErrSection:

    Dim i&

    Disable optScaleLog
    Disable optScalePercent
    Disable optAutoScalePrice
    Disable optSquareScale
    Disable optPtsPerBar
    Disable optTicksPerBar
    Disable txtPtsOrTicks
    Disable txtPerBar
    Disable lblPer
    Disable lblBars
    
    Enable chkHideSeparator, Not m.bSeasonal
    Enable chkPriceTopMost, Not m.bSeasonal
    
    With m.EditedPane
        'Note: .Min/.Max does not reflect temporary scale values
        '  when shifting, scrunching or stretching in the y-scale
        If .PricePaneFlag = 1 Then
            Enable optAutoScalePrice, Not m.bSeasonal
            Enable optSquareScale, Not m.bSeasonal
            txtMin = m.Chart.Bars.PriceDisplay(.Min)
            txtMax = m.Chart.Bars.PriceDisplay(.Max)
        Else
            txtMin = CStr(.Min)
            txtMax = CStr(.Max)
        End If
        
        If .CanBeNonLinear(m.Chart) Then
            Enable optScaleLog
            Enable optScalePercent
        End If
        
        Select Case .Scaling
            Case ePANE_ScaleModeAuto, ePANE_ScaleModeAutoPrice:
                Enable optAutoScale
                Enable optManualScale
                If .Scaling = ePANE_ScaleModeAutoPrice Then
                    optAutoScalePrice = True
                Else
                    optAutoScale = True
                End If
            Case ePANE_ScaleModeManual:
                Enable optAutoScale
                Enable optManualScale
                optManualScale = True
            Case ePANE_ScaleModeSquare:
                Enable optPtsPerBar
                Enable optTicksPerBar
                Enable txtPtsOrTicks
                Enable txtPerBar
                Enable lblPer
                Enable lblBars
                optSquareScale = True
                optScaleLinear = True
                txtPerBar = CStr(.SquareBars)
                If .PointsOrTicksFlag = 0 Then
                    optPtsPerBar = True         'points
                    txtPtsOrTicks.Text = m.Chart.Bars.PriceDisplay(.SquareTicks * m.Chart.Bars.Prop(eBARS_TickMove), "0")
                Else
                    optTicksPerBar = True       'ticks
                    txtPtsOrTicks.Text = Format(.SquareTicks, "0")
                End If
            Case Else:
                optAutoScale = False
                optManualScale = False
                Disable optAutoScale
                Disable optManualScale
                Disable optScaleLog
                Disable optScalePercent
        End Select
        
        If optManualScale Then
            Enable txtMin
            Enable txtMax
            Enable lblMinValue
            Enable lblMaxValue
        Else
            Disable txtMin
            Disable txtMax
            Disable lblMinValue
            Disable lblMaxValue
        End If
        
        If optDisplayFormat(2).Value Then
            Enable txtDecimals
        Else
            Disable txtDecimals
        End If
        
        If .PaneLogFlag > 0 Then
            optScalePercent.Value = True
        ElseIf .PaneLogFlag < 0 Then
            optScaleLog.Value = True
        Else
            optScaleLinear.Value = True
        End If
        
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixCtlPane", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub FixCtlIndicator(Indicator As cIndicator, ByVal idx&)
On Error GoTo ErrSection:

    Dim i&
    
    With Indicator
        SetVisible chkShowInAllPanes, False
        If .DisplayType = eINDIC_HighlightBoxes Then
            SetVisible cboBoxPenStyle, True
        Else
            SetVisible cboBoxPenStyle, False
        End If
        
        Select Case .DataType
            Case eINDIC_BarData
                FixCtlBarData Indicator
            Case eINDIC_Array, eINDIC_DrawCommands, eINDIC_ProfileVolume
                FixCtlArrayData Indicator
            Case eINDIC_BooleanArray
                FixCtlBooleanData Indicator, idx
            Case eINDIC_Constant
                FixCtlConstData Indicator
        End Select
        
        ' Enable "edit" for a custom function or a coded text function
        ' but not for a DLL function (unless has the SDK)
        i = .FunctionID
        If .IsCustom Then
            Enable cmdEditFunction
        ElseIf Not g.Functions.Found(CStr(i)) Then
            Disable cmdEditFunction
        ElseIf g.Functions.Item(CStr(i)).ImplementationTypeID = 2 Then
            Enable cmdEditFunction
        ElseIf DirExist(App.Path & "\..\SDK") Then
            Enable cmdEditFunction
        Else
            Disable cmdEditFunction
        End If
                
        FixCtlLineType Indicator
        FixCtlLabelMode Indicator
    
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixCtlIndicator", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub FixCtlBarData(Indicator As cIndicator)
On Error GoTo ErrSection:

    Dim nUpDown As Long
    Dim bTrueRange As Boolean
    Dim bRemoveGapShow As Boolean
    
    bTrueRange = False
    nUpDown = 1
    With Indicator
        If m.Chart.Tree.Key(Indicator.geIndId) = "PRICE" Then
            Select Case GetPeriodType(m.Chart.Periodicity)
            Case ePRD_EodPF, ePRD_IntPF
                .DisplayType = eINDIC_PNF
            Case ePRD_EodKagi, ePRD_IntKagi
                .DisplayType = eINDIC_Kagi
            Case ePRD_EodRenko, ePRD_IntRenko
                .DisplayType = eINDIC_Renko
            Case Else
                If .DisplayType = eINDIC_Kagi Or .DisplayType = eINDIC_PNF Or .DisplayType = eINDIC_Renko Then
                    .DisplayType = eINDIC_OHLC
                End If
                If .DisplayType < 0 Then
                    bTrueRange = True
                Else
                    nUpDown = 0
                End If
            End Select
            If AllowRemoveOvernightGap Then
                bRemoveGapShow = True
                If m.Chart.RemoveOvernightGap Then
                    chkRemoveOvernightGap.Value = vbChecked
                Else
                    chkRemoveOvernightGap.Value = vbUnchecked
                End If
            End If
        ElseIf .DisplayType = eINDIC_Kagi Or .DisplayType = eINDIC_PNF Or .DisplayType = eINDIC_Renko Then
            .DisplayType = eINDIC_OHLC
            bTrueRange = True
        End If
        'color combo box
        gdColorBars.Color = .Color
        'display type combo box
        OHLCcombo = .DisplayType
        'true range controls
        If bTrueRange Then
            chkTrueRange.Value = .TrueRangeFlag
            gdTrueRange.Color = .trueRangeColor
            If Not chkTrueRange.Enabled Then chkTrueRange.Enabled = True
            gdTrueRange.Visible = True
        Else
            chkTrueRange.Value = 0
            If chkTrueRange.Enabled Then chkTrueRange.Enabled = False
            gdTrueRange.Visible = False
        End If
        
        'overlay & flip check boxes
        If .isPriceInd = 1 Then
            chkOverlayed(0).Visible = False
            chkOverlayed(0).Enabled = False
            If .DisplayType = eINDIC_Candlestick Then
                ' TLB 6/29/2016: re-use the Flip checkbox for new the HideWick option
                chkFlip.Caption = "Hide candle wick (high/low)"
                chkFlip.ToolTipText = "Hide the high/low wicks (so displays only the open/close body)"
                chkFlip.Left = chkTrueRange.Left
                chkFlip.Visible = True
                chkFlip.Enabled = True
                chkFlip.Value = Abs(.HideWick)
            Else
                chkFlip.Visible = False
            End If
        Else
            chkOverlayed(0).Visible = True
            chkOverlayed(0).Enabled = True
            chkFlip.Caption = "Flip price bars (as if upside down)"
            chkFlip.ToolTipText = "Display bar data as if prices were upside down"
            chkFlip.Left = Me.chkOverlayed(0).Left
            chkFlip.Visible = True
            chkFlip.Enabled = True
            chkFlip.Value = Abs(.Flip)
        End If

        'up/down colors controls
        FixCtlUpDownColors Indicator, nUpDown
    End With
    
    chkRemoveOvernightGap.Visible = bRemoveGapShow
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixCtlBarData", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub FixCtlArrayData(Indicator As cIndicator)
On Error GoTo ErrSection:

    Dim i&
    Dim bCluster As Boolean
    Dim bShow As Boolean

    If Indicator.IsHawkeyeAdds Then
        FixCtlBooleanData Indicator, Indicator.geIndpaneId
        Exit Sub
    ElseIf Indicator.IsHawkeyeLevels Then
        gdColor.Visible = False
        fraFakeDropdown.Move gdColor.Left, gdColor.Top
        ToggleFakeDropdown True
    Else
        gdColor.Visible = True
        ToggleFakeDropdown False
    End If

    If lblWidth.Caption <> "Width:" Then
        lblWidth.Caption = "Width:"
        lblWidth.Width = cboType.Width
    End If
    
    With Indicator

        cmdEditFunction.Caption = "  &EDIT Function"
        
        If .IsHawkeyeLevels Then
            chkOverlayed(1).Visible = False
        Else
            chkOverlayed(1).Visible = Not .IsAutoSwingTrendlines
        End If
        
        If .IsCustom Then
            SetVisible fgInputs, False
            SetVisible fraShift, False
            SetVisible lblCustomFunction, True
        ElseIf .DataType = eINDIC_ProfileVolume Then
            SetVisible fgInputs, False
            SetVisible fraShift, False
            SetVisible lblCustomFunction, False
'            SetVisible fraXaxis, True
        Else
            SetVisible fgInputs, Not bCluster
            SetVisible fraShift, Not bCluster
            SetVisible lblCustomFunction, False
        End If

        Enable cboType, Not .IsAutoSwingTrendlines
        Enable txtFunction, Not bCluster
        Enable lblIndType, Not bCluster
        Enable cboType, Not bCluster
        
        SetVisible lblNameForDisplay, Not bCluster
        SetVisible cboType, True
        SetVisible cboLineTypes, True
        
        SetVisible cboHighlightBars, False
        SetVisible chkColorPrice, False
        SetVisible cboMarkerLoc, False
        SetVisible gdSelectIcon, False
        'box highlight bars controls
        SetVisible lblBarsLeft, False
        SetVisible lblBarsRight, False
        SetVisible txtBarsLeft, False
        SetVisible txtBarsRight, False
        SetVisible optBoxFill(0), False
        SetVisible optBoxFill(1), False
        SetVisible gdFillColor, False
        'auto trend lines controls
        SetVisible chkTrendHistory, False
        SetVisible lblExtAutoTrend, False
        SetVisible lblExtBars, False
        
        If .DisplayType < 0 Or .DisplayType >= eINDIC_HighlightBars Then
'JM 07-30-2015: think all this cluster code is obsolete (leave awhile then remove if all ok)
'If .DisplayType <> eINDIC_NoStyle And .DisplayType <> eINDIC_ArtPyramid And .DisplayType <> eINDIC_ArtReversal And _
'   Not .IsClusterPrice() And .DisplayType <> eINDIC_ClusterTime And .DataType <> eINDIC_ProfileVolume Then
            
            If .DisplayType <> eINDIC_NoStyle And .DisplayType <> eINDIC_ArtPyramid And .DisplayType <> eINDIC_ArtReversal And _
               .DataType <> eINDIC_ProfileVolume Then
                
                .DisplayType = eINDIC_Line
            
            End If
        End If
                
        If .DisplayType = eINDIC_ArtPyramid Or .DisplayType = eINDIC_ArtReversal Or .DataType = eINDIC_ProfileVolume Then
            fraArtPyramids.Visible = True
            fraAppearance.Visible = False
            lblShiftBars.Visible = False
            lblShiftBars2.Visible = False
            txtShiftBars.Visible = False
            
            PopulateArtCbo Indicator
            If .DataType = eINDIC_ProfileVolume Then
                fraArtPyramids.Caption = "Profile Display"
                
                lblArtColorDown.Visible = False
                lblArtColorPending.Visible = False
                cmdPyramidFont.Visible = False
                gdPyramidColorDown.Visible = False
                
                lblArtStyle.Visible = True
                lblArtColorDown.Visible = True
                cboPyramidStyle.Visible = True
                chkHorzLetters.Visible = True
                gdColorVolume_POC.Visible = True
                gdColorVolume_POC.Enabled = True
                gdColorVolume_VA.Visible = True
                gdColorVolume_VA.Enabled = True
                gdPyramidColorLabel.Visible = True
                
                'txtPoints is ticks per row (reused name from frmEditAnnot)
                txtPoints.Visible = True
                txtPoints.Enabled = True
                txtPercentVolume_VA.Visible = True
                Label5.Visible = True
                
                lblArtStyle.Caption = "Style"
                lblArtStyle.Top = 345
                lblArtColorUp.Caption = "Volume Profile:"
                lblArtColorUp.Top = lblArtStyle.Top + lblArtStyle.Height + 195
                
                chkHorzLetters.Caption = "Volume VA"
                chkHorzLetters.Top = lblArtColorUp.Top + lblArtStyle.Height + 155
                chkHorzLetters.Left = chkArtChartLabel.Left
                chkHorzLetters.Width = 1275
                
                chkArtChartLabel.Caption = "Volume POC"
                chkArtChartLabel.Top = chkHorzLetters.Top + chkHorzLetters.Height + 195
                chkArtChartLabel.Width = 1275
                
                'profile style, style
                cboPyramidStyle.Top = 315
                gdPyramidColorUp.Top = cboPyramidStyle.Top + cboPyramidStyle.Height + 90
                gdPyramidColorUp.Width = 1350
                
                'VA color, VA text color
                gdColorVolume_VA.Top = gdPyramidColorUp.Top + cboPyramidStyle.Height + 90
                gdPyramidColorP.Top = gdColorVolume_VA.Top
                gdPyramidColorP.Width = gdColorVolume_POC.Width
                
                'POC color, POC text color
                gdColorVolume_POC.Top = gdColorVolume_VA.Top + cboPyramidStyle.Height + 90
                gdPyramidColorLabel.Top = chkArtChartLabel.Top - 60
                gdPyramidColorLabel.Width = gdColorVolume_POC.Width
                
                gdPyramidColorUp.Left = chkHorzLetters.Left + chkHorzLetters.Width      'profile color
                gdColorVolume_VA.Left = gdPyramidColorUp.Left                           'VA color
                gdPyramidColorP.Left = gdColorVolume_VA.Left + gdColorVolume_VA.Width   'VA text color
                gdColorVolume_POC.Left = gdPyramidColorUp.Left                          'POC color
                gdPyramidColorLabel.Left = gdColorVolume_POC.Left + gdColorVolume_POC.Width     'POC text color
                    
                'chkVolume_VA
                '[0] = show previous VA
                '[1] = show previous POC
                '[2] = show VA values
                '[3] = show previous VA values
                '[4] = right-align profiles
                lblArtColorDown.Caption = "Show previous"
                lblArtColorDown.Top = chkArtChartLabel.Top + chkArtChartLabel.Height + 195
                
                chkVolume_VA(0).Caption = "VA"
                chkVolume_VA(1).Caption = "POC"
                chkVolume_VA(0).Width = 675
                chkVolume_VA(1).Width = 675
                chkVolume_VA(0).Left = gdColorVolume_POC.Left
                chkVolume_VA(1).Left = gdPyramidColorLabel.Left
                
                chkVolume_VA(0).Visible = True
                chkVolume_VA(0).Enabled = True
                chkVolume_VA(1).Visible = True
                chkVolume_VA(1).Enabled = True
                chkVolume_VA(2).Visible = True
                chkVolume_VA(2).Enabled = True
                chkVolume_VA(3).Visible = True
                chkVolume_VA(3).Enabled = True
                chkVolume_VA(4).Visible = True
                chkVolume_VA(4).Enabled = True
                chkVolume_VA(0).Top = lblArtColorDown.Top
                chkVolume_VA(1).Top = chkVolume_VA(0).Top
                chkVolume_VA(2).Top = chkVolume_VA(1).Top + chkVolume_VA(1).Height + 120
                chkVolume_VA(3).Top = chkVolume_VA(2).Top + chkVolume_VA(2).Height + 120
                chkVolume_VA(4).Top = chkVolume_VA(3).Top + chkVolume_VA(3).Height + 120
                
                txtPercentVolume_VA.Top = chkVolume_VA(4).Top + chkVolume_VA(4).Height + 60
                lblArtLabels.Caption = "Volume VA Percent:"
                lblArtLabels.Top = txtPercentVolume_VA.Top + 60
                txtPoints.Top = txtPercentVolume_VA.Top + txtPercentVolume_VA.Height + 90
                Label5.Top = txtPoints.Top + 60
                
                If cboVertGrid.Text <> .ProfilePeriodicityStr Then
                    .PopulateCboProfilePeriodicity cboVertGrid, txtBarWidth
                End If
            Else
                fraArtPyramids.Caption = "ART Pyramids"
                lblArtColorUp.Top = 225
                lblArtColorDown.Top = 577
                lblArtColorPending.Top = 930
                lblArtStyle.Top = 1350
                lblArtLabels.Top = 1725
                gdPyramidColorUp.Top = 195
                gdPyramidColorDown.Top = 547
                gdPyramidColorP.Top = 900
                chkHorzLetters.Top = 2340
                chkHorzLetters.Width = 2520
                chkArtChartLabel.Top = 2395
                chkArtChartLabel.Width = 2520
                gdPyramidColorUp.Width = 960
                gdPyramidColorDown.Width = 960
                gdPyramidColorP.Width = 960
                gdPyramidColorP.Left = gdPyramidColorDown.Left
                gdPyramidColorUp.Left = gdPyramidColorDown.Left
                
                gdPyramidColorLabel.Top = 1950
                gdPyramidColorLabel.Left = cboPyramidStyle.Left
                gdPyramidColorLabel.Width = cboPyramidStyle.Width
                
                lblArtColorPending.Caption = "Potential Pyramid:"
                chkHorzLetters.Caption = "Draw letters side by side"
                chkArtChartLabel.Caption = "Show indicator name on chart"
                
                lblArtColorDown.Visible = True
                lblArtColorPending.Visible = True
                gdPyramidColorDown.Visible = True
                cmdPyramidFont.Visible = True
                lblArtLabels.Enabled = True
                
                chkVolume_VA(0).Visible = False
                chkVolume_VA(0).Enabled = False
                chkVolume_VA(1).Visible = False
                chkVolume_VA(1).Enabled = False
                chkVolume_VA(2).Visible = False
                chkVolume_VA(2).Enabled = False
                chkVolume_VA(3).Visible = False
                chkVolume_VA(3).Enabled = False
                chkVolume_VA(4).Visible = False
                chkVolume_VA(4).Enabled = False
                
                gdColorVolume_POC.Visible = False
                gdColorVolume_POC.Enabled = False
                gdColorVolume_VA.Visible = False
                gdColorVolume_VA.Enabled = False
                txtPercentVolume_VA.Visible = False
                txtPercentVolume_VA.Enabled = False
                txtPoints.Visible = False
                txtPoints.Enabled = False
                Label5.Visible = False
            End If
            
            If .DisplayType = eINDIC_ArtPyramid Then
                lblArtColorUp.Caption = "Pyramid Up:"
                lblArtColorDown.Caption = "Pyramid Down:"
                lblArtColorPending.Caption = "Potential Pyramid:"
                
                lblArtLabels.Caption = "Pyramid Label (P/MP):"
                lblArtLabels.Top = 1725
                cmdPyramidFont.Top = gdPyramidColorLabel.Top
                cmdPyramidFont.Left = lblArtLabels.Left
                
                lblArtStyle.Visible = True
                cboPyramidStyle.Visible = True
                gdPyramidColorLabel.Visible = True
                chkHorzLetters.Visible = False
                txtPoints.Visible = False
            ElseIf .DataType <> eINDIC_ProfileVolume Then
                lblArtColorUp.Caption = "Bullish Reversal:"
                lblArtColorDown.Caption = "Bearish Reversal:"
                lblArtColorPending.Caption = "Voided Reversal:"
                
                lblArtLabels.Caption = "Reversal Label:"
                lblArtLabels.Top = lblArtStyle.Top + 50
                cmdPyramidFont.Top = lblArtStyle.Top
                cmdPyramidFont.Left = gdPyramidColorP.Left
                
                chkHorzLetters.Move chkArtChartLabel.Left, chkArtChartLabel.Top - chkHorzLetters.Height - 50
                          
                lblArtStyle.Visible = False
                cboPyramidStyle.Visible = False
                gdPyramidColorLabel.Visible = False
                chkHorzLetters.Visible = True
                txtPoints.Visible = False
            End If
            
        Else
            fraArtPyramids.Visible = False
            fraAppearance.Visible = True
            lblShiftBars.Visible = True
            lblShiftBars2.Visible = True
            txtShiftBars.Visible = True
        End If
                
        If .DisplayType = eINDIC_Area Or .DisplayType = eINDIC_Histogram Or .IsAutoSwingTrendlines Then
            'SetVisible txtColorSeperator, True
            If .IsAutoSwingTrendlines Then
                SetVisible txtColorSeperator, False
                SetVisible txtBaseLineY, False
                gdColor.Color = .UpColor
                gdFillColor.Color = .DownColor
                
                chkColorPrice.Value = .ColorPriceIndFlag
                chkTrendHistory.Value = .ShowTrendHistory
                txtColorSeperator.Text = Str(.ExtendTrend)
                                
                chkColorPrice.Caption = "Highlight swing points"
                chkTrendHistory.Move chkBiColorBars.Left, chkColorPrice.Top + chkColorPrice.Height + 50
                lblExtAutoTrend.Move chkTrendHistory.Left, chkTrendHistory.Top + chkTrendHistory.Height + 50
                txtColorSeperator.Move lblExtAutoTrend.Left + lblExtAutoTrend.Width + 220, lblExtAutoTrend.Top - 80
                lblExtBars.Move txtColorSeperator.Left + txtColorSeperator.Width + 55, lblExtAutoTrend.Top
                txtColorSeperator.Enabled = True
                            
                SetVisible chkColorPrice, True
                SetVisible chkTrendHistory, True
                'SetVisible lblExtAutoTrend, True       'TODO: fix this
                'SetVisible lblExtBars, True
                
                SetVisible chkBiColorBars, False
                SetVisible lblBaseLineY, False
            Else
                SetVisible txtColorSeperator, True
                SetVisible txtBaseLineY, True
                
                txtBaseLineY.Text = CStr(.BaseLineY)
                txtColorSeperator.Text = CStr(.ColorSeperatorVal)
                
                chkBiColorBars.Move chkOverlayed(1).Left, chkOverlayed(1).Top + chkOverlayed(1).Height + 100
                txtColorSeperator.Move chkBiColorBars.Left + chkBiColorBars.Width + 10, chkBiColorBars.Top - 30
                txtColorSeperator.Enabled = chkBiColorBars.Value
                lblBaseLineY.Move chkBiColorBars.Left, chkBiColorBars.Top + chkBiColorBars.Height + 70
                txtBaseLineY.Move txtColorSeperator.Left, txtColorSeperator.Top + txtColorSeperator.Height + 5, txtColorSeperator.Width, txtColorSeperator.Height
            
                SetVisible chkBiColorBars, True
                SetVisible lblBaseLineY, True
            End If
            If .IsBiColorHistogram Or .IsAutoSwingTrendlines Then
                gdColor.Width = gdFillColor.Width
                gdFillColor.Move gdColor.Left + gdColor.Width, gdColor.Top
                gdFillColor.Visible = True
            Else
                gdColor.Width = cboType.Width
                gdFillColor.Visible = False
            End If
        ElseIf .DisplayType = eINDIC_ArtPyramid Or .DisplayType = eINDIC_ArtReversal Then
            gdPyramidColorP.Color = .BoxFilLColor
            gdPyramidColorUp.Color = .UpColor
            gdPyramidColorDown.Color = .DownColor
            gdPyramidColorLabel.Color = .Color
            If .IndLabelMode = eINDIC_NoValue Then
                chkArtChartLabel.Value = vbChecked
            Else
                chkArtChartLabel.Value = vbUnchecked
            End If
            If .DisplayType = eINDIC_ArtPyramid Then
                cboPyramidStyle.ListIndex = .Style
            Else
                chkHorzLetters.Value = .BoxFillStyle
            End If
        ElseIf .DataType = eINDIC_ProfileVolume Then
            gdPyramidColorUp.Color = .ProfileColor(ePCStruct_TPO)
            'VA color, VA text color
            gdColorVolume_VA.Color = .ProfileColor(ePCStruct_Volume_VA)
            gdPyramidColorP.Color = .ProfileColor(ePCStruct_TPO_VA)
            'POC color, POC text coor
            gdColorVolume_POC.Color = .ProfileColor(ePCStruct_Volume_POC)
            gdPyramidColorLabel.Color = .ProfileColor(ePCStruct_TPO_POC)
            
            txtPercentVolume_VA.Text = Str(.ProfileParm(ePCStruct_Volume_VA))
            txtPoints.Text = .TicksPerRow
            gdPyramidColorUp.Width = 1350

            chkHorzLetters.Value = .ChkboxValForProfileStruct(ePCStruct_Volume_VA, False)       'show/hide current VA
            chkVolume_VA(0).Value = .ChkboxValForProfileStruct(ePCStruct_Volume_VA, True)       'show/hide previous VA
            chkArtChartLabel.Value = .ChkboxValForProfileStruct(ePCStruct_Volume_POC, False)    'show/hide current POC
            chkVolume_VA(1).Value = .ChkboxValForProfileStruct(ePCStruct_Volume_POC, True)      'show/hide previous POC
            chkVolume_VA(4).Value = .ProfileShowHide(ePCStruct_Close)
            FixProfileChkboxes Indicator
            
            i = .ProfileShowHide(ePCStruct_Volume_VA)
            If i > 0 Then bShow = True
            txtPercentVolume_VA.Enabled = bShow
            lblArtLabels.Enabled = bShow
        Else
            SetVisible chkBiColorBars, False
            SetVisible txtColorSeperator, False
            SetVisible lblBaseLineY, False
            SetVisible txtBaseLineY, False
            SetVisible lblExtBars, False
            SetVisible chkTrendHistory, False
            gdFillColor.Visible = False
            gdColor.Width = cboType.Width
        End If
        
        If .DisplayType = eINDIC_ArtPyramid Or .DisplayType = eINDIC_ArtReversal Or .DataType = eINDIC_ProfileVolume Then
            'not allowing overlay ART Pyramids                  07/30/2008
            chkOverlayed(1) = 0
            .Overlayed = 0
            Disable chkOverlayed(1)
        Else
            Enable chkOverlayed(1)
        End If
        
'JM 07-30-2015: think all this cluster code is obsolete (leave awhile then remove if all ok)
'        If .IsClusterPrice Then
'            If .DisplayType = eINDIC_ClusterPriceLine Then
'                cboType.ListIndex = 0
'            Else
'                cboType.ListIndex = cboType.ListCount - 1
'            End If
                                                    'JM 07-30-2015: think all this cluster code is obsolete (leave awhile then remove if all ok)
        If .DisplayType = eINDIC_NoStyle Then       'Or .DisplayType = eINDIC_ClusterTime
            cboType.ListIndex = cboType.ListCount - 1
        ElseIf .DisplayType <> eINDIC_ArtPyramid And .DisplayType <> eINDIC_ArtReversal And .DataType <> eINDIC_ProfileVolume Then
            cboType.ListIndex = .DisplayType
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixCtlArrayData", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub FixCtlBooleanData(Indicator As cIndicator, ByVal idx&)
On Error GoTo ErrSection:

    Dim i&
    Dim Ind As cIndicator
    
    'true if user created as highlight bars, but got switched to markers due to Up/Down or Bollinger Bars
    Dim bMarkerOverride As Boolean          'aardvark 4572
    
    SetVisible lblBaseLineY, False          'aardvark 3910
    SetVisible txtBaseLineY, False
    
    If fraArtPyramids.Visible Then
        fraArtPyramids.Visible = False
        fraAppearance.Visible = True
        lblShiftBars.Visible = True
        lblShiftBars2.Visible = True
        txtShiftBars.Visible = True
    End If
    
    With Indicator
        If .IsCustom Then
            SetVisible fgInputs, False
            SetVisible fraShift, False
            SetVisible lblCustomFunction, True
        Else
            SetVisible fgInputs, True
            SetVisible fraShift, True
            SetVisible lblCustomFunction, False
        End If
        
        gdColor.Width = cboType.Width
        chkOverlayed(1).Visible = False
        SetVisible cboHighlightBars, True
        SetVisible cboType, False
        SetVisible cboLineTypes, False
        'bicolor histogram controls
        SetVisible chkBiColorBars, False
        SetVisible txtColorSeperator, False
        'auto trend lines controls
        SetVisible chkTrendHistory, False
        SetVisible lblExtAutoTrend, False
        SetVisible lblExtBars, False
        
        If .DisplayType = eINDIC_HighlightBoxes Then
            SetVisible cboMarkerLoc, False
            SetVisible gdSelectIcon, False
            lblWidth.Caption = "Width"
            'box highlight bars controls
            gdFillColor.Move optBoxFill(1).Left + optBoxFill(1).Width, optBoxFill(0).Top - 80
            SetVisible lblBarsLeft, True
            SetVisible lblBarsRight, True
            SetVisible txtBarsLeft, True
            SetVisible txtBarsRight, True
            SetVisible optBoxFill(0), True
            SetVisible optBoxFill(1), True
            SetVisible gdFillColor, True
        ElseIf .DisplayType = eINDIC_ValueMarkers Then
            SetVisible cboMarkerLoc, True
            SetVisible gdSelectIcon, False
            lblWidth.Caption = "Marker location:"
            lblWidth.Width = cboType.Width * 2
            'box highlight bars controls
            SetVisible lblBarsLeft, False
            SetVisible lblBarsRight, False
            SetVisible txtBarsLeft, False
            SetVisible txtBarsRight, False
            SetVisible optBoxFill(0), False
            SetVisible optBoxFill(1), False
            SetVisible gdFillColor, False
        Else
            SetVisible cboMarkerLoc, True
            SetVisible gdSelectIcon, True
            lblWidth.Caption = "Icon"
            'box highlight bars controls
            SetVisible lblBarsLeft, False
            SetVisible lblBarsRight, False
            SetVisible txtBarsLeft, False
            SetVisible txtBarsRight, False
            SetVisible optBoxFill(0), False
            SetVisible optBoxFill(1), False
            SetVisible gdFillColor, False
        End If
        cboHighlightBars.Clear
                
        If m.Chart.Tree.NodeLevel(.IndToColor) > 0 Then        'aardvark 2082 fix
            Set Ind = m.Chart.Tree(.IndToColor)
        Else
            Set Ind = m.Chart.Tree("PRICE")
        End If
        If Ind Is Nothing Then
            cboHighlightBars.AddItem "Highlight Bars"
            cboHighlightBars.AddItem "Highlight Markers"
            cboHighlightBars.AddItem "Highlight Boxes"
            cboHighlightBars.AddItem "Highlight Zones"
            cboHighlightBars.AddItem "None"
        Else
            'disallow highlight bars for BollingerBar bar AND price bars using Up/Down colors
            If Ind.DisplayType = eINDIC_BollingerBar Or Ind.UpDownColorFlag <> 0 Then
                cboHighlightBars.AddItem "Highlight Markers"
                cboHighlightBars.AddItem "Highlight Boxes"
                cboHighlightBars.AddItem "Highlight Zones"
                cboHighlightBars.AddItem "None"
                If Indicator.DisplayType = eINDIC_HighlightBars Then bMarkerOverride = True
            Else
                cboHighlightBars.AddItem "Highlight Bars"
                cboHighlightBars.AddItem "Highlight Markers"
                cboHighlightBars.AddItem "Highlight Boxes"
                cboHighlightBars.AddItem "Highlight Zones"
                If Indicator.IsHawkeyeAdds Then cboHighlightBars.AddItem "Value Markers"
                cboHighlightBars.AddItem "None"
            End If
        End If
        
        Enable cboHighlightBars, Not bMarkerOverride
        
        If .DisplayType = eINDIC_HighlightMarkers Or bMarkerOverride Then
            Enable lblWidth, True
            Enable gdSelectIcon, True
            Enable cboMarkerLoc, True
            If cboHighlightBars.ListCount = 6 Then
                cboHighlightBars.ListIndex = cboHighlightBars.ListCount - 5
            Else
                cboHighlightBars.ListIndex = cboHighlightBars.ListCount - 4
            End If
            cboMarkerLoc.ListIndex = .MarkerLoc
            gdSelectIcon.Ascii = .MarkerAscii
            gdSelectIcon.Icon = .MarkerImage(True)
        ElseIf .DisplayType = eINDIC_HighlightBoxes Then
            Enable lblWidth, True
            If cboHighlightBars.ListCount = 6 Then
                cboHighlightBars.ListIndex = cboHighlightBars.ListCount - 4
            Else
                cboHighlightBars.ListIndex = cboHighlightBars.ListCount - 3
            End If
            cboBoxPenStyle.ListIndex = .BoxPenStyle
            gdFillColor.Color = .BoxFilLColor
            txtBarsLeft.Text = Str(.HighlightBarsLeft)
            txtBarsRight.Text = Str(.HighlightBarsRight)
            Enable optBoxFill(0), True
            If .BoxFillStyle = 0 Then
                optBoxFill(0) = True
                optBoxFill(1) = False
            Else
                optBoxFill(0) = False
                optBoxFill(1) = True
            End If
        ElseIf .DisplayType = eINDIC_HighlightZones Then
            chkShowInAllPanes.Move chkShowInAllPanes.Left, lblBarsLeft.Top
            If cboHighlightBars.ListCount = 6 Then
                cboHighlightBars.ListIndex = cboHighlightBars.ListCount - 3
            Else
                cboHighlightBars.ListIndex = cboHighlightBars.ListCount - 2
            End If
            
            If .ShowInAllPanes > 1 Then
                chkShowInAllPanes.Value = 1
            Else
                chkShowInAllPanes.Value = 0
            End If
            optBoxFill(1) = True
            gdFillColor.Color = .BoxFilLColor
            Enable gdSelectIcon, False
            Enable cboMarkerLoc, False
            Enable optBoxFill(0), False
            chkShowInAllPanes.Visible = True
            gdFillColor.Visible = True
            optBoxFill(0).Visible = True
            optBoxFill(1).Visible = True
        ElseIf .DisplayType = eINDIC_ValueMarkers Then
            Enable cboMarkerLoc, True
            Enable gdSelectIcon, False
            cboMarkerLoc.ListIndex = .MarkerLoc
            cboHighlightBars.ListIndex = cboHighlightBars.ListCount - 2
        Else
            Enable lblWidth, False
            Enable gdSelectIcon, False
            Enable cboMarkerLoc, False
            If .DisplayType = eINDIC_HighlightBars Then
                cboHighlightBars.ListIndex = 0
            Else
                cboHighlightBars.ListIndex = cboHighlightBars.ListCount - 1
            End If
        End If
        
        i = m.Chart.Tree.RelativeIndex(idx, eTREE_Root)
        If m.Chart.Tree.Key(i) = "PRICE PANE" Or .DisplayType = eINDIC_HighlightBoxes _
            Or .DisplayType = eINDIC_HighlightZones Then
            SetVisible chkColorPrice, False
        Else
            chkColorPrice.Caption = "Highlight/mark price bars"
            chkColorPrice.Top = cboMarkerSize.Top + cboMarkerSize.Height + 50
            SetVisible chkColorPrice, True
            chkColorPrice.Value = .ColorPriceIndFlag
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixCtlBooleanData", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub FixCtlConstData(Indicator As cIndicator)
On Error GoTo ErrSection:

    If fraArtPyramids.Visible Then
        fraArtPyramids.Visible = False
        fraAppearance.Visible = True
    End If
    
    With Indicator
        gdColor.Width = cboType.Width
        chkOverlayed(1).Visible = False
        SetVisible fgInputs, True
        SetVisible fraShift, False
        SetVisible lblCustomFunction, False
        cboType.ListIndex = 0
        SetVisible cboType, True
        SetVisible cboLineTypes, True
        SetVisible cboHighlightBars, False
        SetVisible chkColorPrice, False
        SetVisible cboMarkerLoc, False
        SetVisible gdSelectIcon, False
        Disable cboType
        'box highlight bars controls
        SetVisible lblBarsLeft, False
        SetVisible lblBarsRight, False
        SetVisible txtBarsLeft, False
        SetVisible txtBarsRight, False
        SetVisible optBoxFill(0), False
        SetVisible optBoxFill(1), False
        SetVisible gdFillColor, False
        'bicolor histogram controls
        SetVisible txtColorSeperator, False
        SetVisible chkBiColorBars, False
        'auto swing trendlines controls
        SetVisible lblExtBars, False
        SetVisible chkTrendHistory, False

    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixCtlConstData", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub FixCtlLineType(Indicator As cIndicator)
On Error GoTo ErrSection:

    Dim cboControl As ctlUniComboBoxXP
    Dim eType As eIndicatorDisplayType
    
    If Indicator.DataType = eINDIC_BarData Then
        Set cboControl = cboLineTypesBars
        If Indicator.DisplayType = eINDIC_Area Or Indicator.DisplayType = eINDIC_BollingerBar Then
            Enable lblWidthBars, False
            Enable cboLineTypesBars, False
        Else
            Enable lblWidthBars, True
            Enable cboLineTypesBars, True
        End If
    ElseIf Indicator.DataType <> eINDIC_BooleanArray Then
        Set cboControl = cboLineTypes
        eType = Indicator.DisplayType
        
                                        'JM 07-30-2015: think all this cluster code is obsolete (leave awhile then remove if all ok)
        If eType = eINDIC_Area Then     'Or eType = eINDIC_ClusterTime Or Indicator.IsClusterPrice()
            Enable lblWidth, False
            Enable cboLineTypes, False
        Else
            Enable lblWidth, True
            Enable cboLineTypes, True
        End If
    End If
    
    If cboControl Is Nothing Then
        Exit Sub
    End If

    With cboControl
        'add dashed pen styles for display type line
        If Indicator.DisplayType = eINDIC_Line Then
            If .ListCount > 7 Then
                If Left(.List(7), 4) <> "Dash" Then
                    Do While .ListCount > 7
                        .RemoveItem 7
                    Loop
                End If
            End If
            If .ListCount <= 7 Then
                .AddItem "Dashed (Large)"
                .AddItem "Dashed (Small)"
                .AddItem "Dash Dot"
            End If
        Else
            'handle auto style
            If vseBars.Visible And Indicator.DisplayType < 0 And Indicator.DisplayType <> eINDIC_BollingerBar Then
                If .ListCount > 7 Then
                    If Left(.List(7), 4) <> "Auto" Then
                        Do While .ListCount > 7
                            .RemoveItem 7
                        Loop
                    End If
                End If
                If .ListCount <= 7 Then
                    .AddItem "Auto (variable)"
                End If
            Else
                'remove dashed pen and/or auto styles
                Do While .ListCount > 7
                    .RemoveItem 7
                Loop
            End If
        End If
        
        If .ListCount > Indicator.Style Then
            .ListIndex = Indicator.Style
        Else
            .ListIndex = 0
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixCtlLineType", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub FixCtlLabelMode(Indicator As cIndicator)
On Error GoTo ErrSection:

    Dim cboMode As ctlUniComboBoxXP
    Dim IndToColor As cIndicator
        
    If Indicator.DisplayType = eINDIC_HighlightMarkers Or Indicator.DisplayType = eINDIC_ValueMarkers Then
        lblMarkerSize.Visible = True
        cboMarkerSize.Visible = True
        cboMarkerSize.ListIndex = Indicator.MarkerSize - 1
    Else
        lblMarkerSize.Visible = False
        cboMarkerSize.Visible = False
    End If
    
    If Indicator.DataType = eINDIC_BarData Then
        Set cboMode = cboBarsLabelMode
    Else
        Set cboMode = cboIndLabelMode
    End If
    If Not cboMode Is Nothing Then
        If Indicator.DisplayType = eINDIC_Line And Indicator.DataType = eINDIC_Constant Then
            cboMode.ListIndex = eINDIC_Nothing
            Enable cboMode, False
        Else
            cboMode.ListIndex = Indicator.IndLabelMode
            Enable cboMode, True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixCtlLabelMode", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub FixCtlUpDownColors(Indicator As cIndicator, ByVal nUpDownFlag&)
On Error GoTo ErrSection:

    With Indicator
        'down color label
        lblDownColor.Visible = nUpDownFlag
        'up/down check box
        chkUpDownColors.Value = Abs(.UpDownColorFlag)
        If .DisplayType = eINDIC_BollingerBar Then
            SetVisible chkUpDownColors, False
            SetVisible lblUpColor, True
        Else
            SetVisible chkUpDownColors, True
            SetVisible lblUpColor, False
            Enable chkUpDownColors, nUpDownFlag
            If nUpDownFlag = 0 Or .UpDownColorFlag = 0 Then
                chkUpDownColors.Width = gdColorBars.Width
                chkUpDownColors.Caption = "Use Up/Down colors"
            Else
                chkUpDownColors.Width = gdColorUp.Width
                chkUpDownColors.Caption = "Up:"
            End If
        End If
        'up/down colors
        gdColorUp.Color = .UpColor
        gdColorDown.Color = .DownColor
        If .DisplayType = eINDIC_BollingerBar Then
            SetVisible gdColorUp, True
            SetVisible gdColorDown, True
        Else
            If nUpDownFlag = 0 Or .UpDownColorFlag = 0 Then
                SetVisible gdColorUp, False
                SetVisible gdColorDown, False
            Else
                SetVisible gdColorUp, True
                SetVisible gdColorDown, True
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartCfg.FixCtlUpDownColors", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Property Get OHLCcombo() As eIndicatorDisplayType
On Error GoTo ErrSection
    
    Select Case UCase(Left(Parse(cboOHLCType.Text, " ", 1), 6))
    Case "BOLLIN"
        OHLCcombo = eINDIC_BollingerBar
    Case "RENKO"
        OHLCcombo = eINDIC_Renko
    Case "KAGI"
        OHLCcombo = eINDIC_Kagi
    Case "POINT"
        OHLCcombo = eINDIC_PNF
    Case "OHLC"
        OHLCcombo = eINDIC_OHLC
    Case "HLC"
        OHLCcombo = eINDIC_HLC
    Case "HL"
        OHLCcombo = eINDIC_HL
    Case "CANDLE"
        OHLCcombo = eINDIC_Candlestick
    Case "CLOSE"
        OHLCcombo = eINDIC_Line
    Case "HISTOG"
        OHLCcombo = eINDIC_Histogram
    Case "MOUNTA"
        OHLCcombo = eINDIC_Area
    Case "POINTS"
        OHLCcombo = eINDIC_Points
    Case "NONE"
        OHLCcombo = eINDIC_NoStyle
    End Select
        
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmChartCfg.GetOHLCcombo", eGDRaiseError_Raise
    Resume ErrExit
End Property

Private Property Let OHLCcombo(ByVal eType As eIndicatorDisplayType)
On Error GoTo ErrSection
    Dim i&, s$

    Select Case eType
    Case eINDIC_BollingerBar '= -8
        s = "Bollinger"
    Case eINDIC_Renko '= -7
        s = "Renko"
    Case eINDIC_Kagi '= -6
        s = "Kagi"
    Case eINDIC_PNF '= -5
        s = "Point "
    Case eINDIC_OHLC '= -4
        s = "OHLC"
    Case eINDIC_HLC '= -3
        s = "HLC"
    Case eINDIC_HL '= -2
        s = "HL"
    Case eINDIC_Candlestick '= -1
        s = "Candle"
    Case eINDIC_Line '= 0
        s = "Close"
    Case eINDIC_Histogram '= 1
        s = "Histogram"
    Case eINDIC_Area '= 2
        s = "Mountain"
    Case eINDIC_Points '= 3
        s = "Points"
    Case eINDIC_NoStyle '=30
        s = "None"
    End Select
    
    With cboOHLCType
        For i = 0 To .ListCount - 1
            If Left(UCase(.List(i)), Len(s)) = UCase(s) Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmChartCfg.LetOHLCcombo", eGDRaiseError_Raise
    Resume ErrExit

End Property

Private Sub ShowInputTip(ByVal nRow As Long)
On Error GoTo ErrSection:

    Dim strDesc$, strTemp$
    Dim nTopOfGrid&, nTopOfRow&
    Dim nTop&, nLeft&, nRowHeight&
    Dim nCharPerLine&, nTwipsPerChar&, i&
    
    If nRow < 0 Or m.EditedIndicator Is Nothing Then
        vseTipInput.Top = -1000 - vseTipInput.Height    'move off screen
        Exit Sub
    End If
        
    i = fgInputs.RowData(nRow)
    If i > 0 Then
        strDesc = m.EditedIndicator.ParmDesc(i)
    End If
    If Len(strDesc) < 1 Then
        vseTipInput.Top = -1000 - vseTipInput.Height    'move off screen
        Exit Sub
    End If

    With vseTipInput
        strDesc = " " & Replace(strDesc, vbCrLf, vbCrLf & " ")
        .Width = Me.TextWidth(strDesc) + 120
        .Height = Me.TextHeight(strDesc) + 20
        
        'auto wrap text if too wide
        If .Width > Me.Width Then
            nTwipsPerChar = Me.TextWidth("W")   'this is an approximation
            nCharPerLine = Me.Width / nTwipsPerChar
            If Len(strDesc) < nCharPerLine Then nCharPerLine = Len(strDesc) / 2
            strTemp = strTemp & Mid(strDesc, 1, nCharPerLine)
            For i = nCharPerLine To Len(strDesc) Step nCharPerLine
                strTemp = strTemp & vbCrLf & "  " & Mid(strDesc, i, nCharPerLine)
            Next
            strDesc = strTemp
            .Width = Me.TextWidth(strDesc) + 120
            .Height = Me.TextHeight(strDesc) + 20
        End If
        
        lblTipInput.Caption = strDesc
        nRowHeight = fgInputs.RowHeight(0)
                
        nTopOfGrid = fraFunction.Top + fgInputs.Top
        nTopOfRow = nTopOfGrid + nRowHeight * nRow
        
        nLeft = Me.ScaleWidth - .Width
        nTop = nTopOfRow - .Height
        
        .Move nLeft, nTop
        
        lblTipInput.Move 0, 0, .Width, .Height
        vseTipInput.ZOrder
    End With


ErrExit:
    Exit Sub
        
ErrSection:
    RaiseError "frmChartCfg.ShowInputTip", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub FixTradesControls(Chart As cChart)
On Error GoTo ErrSection:

    Dim nShowTrades&    '0=none,1=strategy,2=trade account
    Dim nAccountID&, i&
    Dim bUpdateFlag As Boolean
    Dim bGameModeLock As Boolean
    Dim bGameMode As Boolean
    
    If Not m.Chart Is Nothing Then
        If m.Chart.Form.IsInGameMode Then
            bGameMode = True
            optTradesNone.Enabled = False
            cboAccounts.Clear
            cboAccounts.AddItem "GAME MODE"
            cboAccounts.BackColor = vbRed
            cboAccounts.ListIndex = 0
            cboAccounts.Enabled = False
            If Not m.Chart.Form.GameMode Is Nothing Then
                If m.Chart.Form.GameMode.CustomOrders > 0 Then
                    bGameModeLock = True
                End If
            End If
            cmdEditSystem.Enabled = False
        Else
            cmdEditSystem.Enabled = m.Chart.SystemID
            If m.Chart.IsInWhatIfMode Then
                optTradesAccount.Enabled = False
            Else
                If cboAccounts.ListCount = 0 Or Len(m.Chart.SpreadSymbols) <> 0 Or _
                   (m.Chart.Bars.SecurityType = "S" And g.nReplaySession > 0) Then
                    optTradesAccount.Enabled = False
                    If m.Chart.ShowTrades = 1 Then
                        m.Chart.ShowTrades = 0
                    End If
                End If
            End If
        End If
    End If
    
    If bGameModeLock Then
        fraSystem.Caption = "Game Mode"
        lblSystemName.Visible = False
        cmdSelectSystem.Visible = False
        cmdEditSystem.Visible = False
        cboAccounts.Visible = True
        fraSystem.Visible = True
        fraProfitLines.Visible = True
        fraTrades.Enabled = False
        optTradesStrategy.Enabled = False
        optTradesAccount.Enabled = False
        optTradesAccount.Value = True
    ElseIf Chart Is Nothing Then
        If optTradesStrategy.Value = True Then
            nShowTrades = 1
        ElseIf optTradesAccount.Value = True Then
            nShowTrades = 2
            If bGameMode Then
                If Not m.Chart Is Nothing Then
                    m.Chart.SystemID = 0    'this is like removing strategy to clear trade triangles
                End If
            End If
        End If
        bUpdateFlag = True
    Else
        nShowTrades = Chart.ShowTrades
        Select Case nShowTrades
            Case 1:
                optTradesStrategy.Value = True
            Case 2:
                optTradesAccount.Value = True
            Case Else:
                optTradesNone.Value = True
        End Select
        nAccountID = Chart.TradeAccountID
        For i = 0 To cboAccounts.ListCount - 1
            If cboAccounts.ItemData(i) = nAccountID Then
                cboAccounts.ListIndex = i
                Exit For
            End If
        Next
    End If
    
    If bGameModeLock Then Exit Sub
    
    If nShowTrades = 1 Or nShowTrades = 2 Then
        fraSystem.Visible = True
        fraProfitLines.Visible = True
        optLineBox(m.Chart.ProfitLineBox).Value = True
        If nShowTrades = 1 Then
            fraSystem.Caption = "Trading Strategy"
            cmdSelectSystem.Visible = True
            cmdEditSystem.Visible = True
            lblSystemName.Visible = True
            lblAccounts.Visible = False
            cboAccounts.Visible = False
            cmdTradeSettings.Enabled = False
        Else
            cmdSelectSystem.Visible = False
            cmdEditSystem.Visible = False
            lblSystemName.Visible = False
            lblAccounts.Visible = True
            cboAccounts.Visible = True
            If bGameMode Then
                fraSystem.Caption = "Game Mode"
                cmdTradeSettings.Enabled = False
            Else
                fraSystem.Caption = "Trading Account"
                cmdTradeSettings.Enabled = True
            End If
        End If
    Else
        fraSystem.Visible = False
        fraProfitLines.Visible = False
        cmdTradeSettings.Enabled = False
    End If
    
    If Not m.Chart Is Nothing Then
        If bUpdateFlag Then
            m.Chart.ShowTrades = nShowTrades
            m.Chart.GenerateChart eRedo1_Scrolled   'aardvark 4196
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.FixTradesControls"
    
End Sub

Private Sub InitCboAccounts()
On Error GoTo ErrSection:
    
    Dim nAccountID&
    
    If Not m.Chart Is Nothing Then
        nAccountID = m.Chart.TradeAccountID
    End If
    
    PopulateAccountsCbo cboAccounts, nAccountID
        
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.InitCboAccounts"

End Sub

Private Function GroupIndicators() As Boolean
On Error GoTo ErrSection:

    Dim i&, iFirst&, strGroupKey$
    
    Dim Ind As cIndicator
    Dim Tree As cGdTree
    
    If m.Chart Is Nothing Then Exit Function
    Set Tree = m.Chart.Tree
    If Tree Is Nothing Then Exit Function
    
    iFirst = -1

    With fgSettings
        If .SelectedRows > 1 Then
            For i = .FixedRows + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    If .IsSelected(i) And RowType(i) = kIndicator Then
                        Set Ind = Tree(.RowData(i))
                        If Not Ind Is Nothing Then
                            If iFirst = -1 Then
                                iFirst = i
                                strGroupKey = Ind.MyKey
                                .Cell(flexcpChecked, i, kBoxCol) = flexChecked
                            Else
                                .Cell(flexcpChecked, i, kBoxCol) = flexNoCheckbox
                            End If
                            Ind.Display = True
                            Ind.GroupKey = strGroupKey
                            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = 0
                        End If
                    ElseIf iFirst <> -1 Then
                        Exit For
                    End If
                End If
            Next
            If iFirst > .FixedRows And iFirst < .Rows Then .Select iFirst, 0, iFirst, .Cols - 1
            GroupIndicators = True
        Else
            InfBox "Use shift-click to select multiple indicators for grouping."
        End If
    End With

    Exit Function

ErrSection:
    RaiseError "frmChartCfg.GroupIndicators"

End Function

Private Sub SetStartStopTimesText()
On Error GoTo ErrSection:

    StartStopTimeLabel m.Chart, lblStartStopTimes, lblStartStopInfo2
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.SetStartStopTimesText"

End Sub

Private Function RemoveIndicator(ByVal nRow&, ByVal strKey$) As Boolean
On Error GoTo ErrSection:

    Dim idx As Long
    Dim strGroupKey$, strAnswer$
    
    Dim bRemoved As Boolean
    
    Dim Ind As cIndicator
    Dim IndLeader As cIndicator
        
    Set Ind = m.Chart.Tree(strKey)
    If Ind Is Nothing Then Exit Function
    
    If Ind.DisplayType = eINDIC_Ribbon Then RibbonList Nothing, m.Chart, Ind, 2, 0
    
    strGroupKey = Ind.GroupKey
    idx = fgSettings.RowData(nRow)
    
    If Len(strGroupKey) = 0 Then
        'not in a group
        If Ind.HasAssociatedAlerts(False, False) Then
            strAnswer = InfBox("Associated alert(s) will also be removed.", "?", "+Okay|-Cancel", "Confirmation")
            If strAnswer = "O" Then
                RemoveIndicator = m.Chart.Tree.Remove(idx)  'remove this first otherwise idx will be out-of-sync
                Ind.HasAssociatedAlerts True, True
            End If
        Else
            RemoveIndicator = m.Chart.Tree.Remove(idx)
        End If
        GoTo ErrExit
    End If
    
    If Ind.IAmGroupLeader Then
        bRemoved = Ind.RemoveMyGroup
    ElseIf InfBox("Remove all indicators in group?", "?", "-Yes|+No", "Confirmation") = "Y" Then
        bRemoved = Ind.RemoveMyGroup
    Else
        Set IndLeader = m.Chart.Tree(strGroupKey)
        If Not IndLeader Is Nothing Then
            bRemoved = m.Chart.Tree.Remove(idx)
            Ind.HasAssociatedAlerts True, True
            IndLeader.SaveGroupInfo
        End If
    End If
    
    RemoveIndicator = bRemoved

ErrExit:
    Set Ind = Nothing
    Set IndLeader = Nothing
    Exit Function

ErrSection:
    RaiseError "frmChartCfg.RemoveIndicator"
    Resume ErrExit

End Function

Private Sub InitColorsGrid()
On Error GoTo ErrSection:

    Dim i&

    With fgColors           'grid for hawkeye levels colors
        .Cols = 2
        .Rows = 9
        .FixedCols = 0
        .FixedRows = 0
        .ColWidth(0) = 300
        .ExtendLastCol = True
        .Font.Bold = True
        .HighLight = flexHighlightNever
        .Width = gdColor.Width
        .Height = .RowHeight(0) * 9
        
        For i = 0 To 8
            .TextMatrix(i, 0) = Str(i)
        Next
        .Visible = False
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.InitColorsGrid"

End Sub

Private Sub SetColorsGrid()
On Error GoTo ErrSection:

    Dim i&, j&, strColors$
    Dim aColors As cGdArray
    
    If Not m.EditedIndicator Is Nothing Then
        strColors = m.EditedIndicator.HawkeyeLevelsColors
        Set aColors = New cGdArray
        aColors.SplitFields strColors, ";"
        j = aColors.Size - 1
        If j > 8 Then j = 8
        For i = 0 To j
            fgColors.Cell(flexcpBackColor, i, 1) = Val(aColors(i))
        Next
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.SetColorsGrid"

End Sub

Private Sub ToggleFakeDropdown(ByVal bOn As Boolean)
On Error Resume Next

    If fraFakeDropdown.Visible <> bOn Then
        fraFakeDropdown.Visible = bOn
        fraFakeDropdown.Enabled = bOn
        cmdFakeDropdown.Visible = bOn
        cmdFakeDropdown.Enabled = bOn
        txtFakeDropdown.Visible = bOn
        txtFakeDropdown.Enabled = bOn
    End If
    
End Sub

Private Sub InitLinkedInputsGrid()
On Error GoTo ErrSection:

    Dim strLabel$, i&
    
    vseLinkedInputs.BorderWidth = 0
    fraLinkedInputs.Height = vseLinkedInputs.Height - 160

    With lblLinkedInputUsage
        .Move fgInputs.Left, .Top, fgInputs.Width
        .Caption = "To link inputs of multiple indicators:" & vbCrLf & _
            "- First enter the same variable name (starting with an '&') as the value for each input to be linked." _
            & vbCrLf & "- Then set the actual value for that variable in the table below."
    End With
    
    With fgLinkedInputs
        .Redraw = flexRDNone
        i = lblLinkedInputUsage.Top + lblLinkedInputUsage.Height
        .Move fgInputs.Left, i, fgInputs.Width, fraLinkedInputs.Height - i - 120
        
        .FixedCols = 0
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .AllowSelection = True ' False
        .HighLight = flexHighlightWithFocus
        .SheetBorder = RGB(128, 128, 128)
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
                
        .Cols = 3
        .ColWidth(0) = .Width / 2
        .ColAlignment(1) = flexAlignCenterCenter
        .ColHidden(2) = True
        
        .Rows = .FixedRows
        
        .TextMatrix(0, 0) = "Variable Name"
        .TextMatrix(0, 1) = "Value"
        .TextMatrix(0, 2) = "TypeID"
        
        .Select 0, 0, 0, .Cols - 1
        .CellFontBold = True
        
        .Redraw = flexRDBuffered
    End With
    
    PopulateLinkedInputsGrid

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.InitLinkedInputsGrid"

End Sub

Private Sub PopulateLinkedInputsGrid()
On Error GoTo ErrSection:

    Dim i&, strValue$, strType$
    Dim aInputs As New cGdArray

    If Not m.Chart Is Nothing Then
        aInputs.SplitFields m.Chart.LinkedInputString, "|"
        With fgLinkedInputs
            .Rows = .FixedRows
            For i = 0 To aInputs.Size - 1
                .Rows = .Rows + 1
                strValue = StripStr(Parse(aInputs(i), ";", 2), Chr(34))
                strType = Parse(aInputs(i), ";", 3)
                
                .TextMatrix(.Rows - 1, 0) = Parse(aInputs(i), ";", 1)
                
                If strType = 6 Then
                    If strValue = "1" Then
                        strValue = "True"
                    ElseIf strValue = "0" Then
                        strValue = "False"
                    End If
                End If
                
                .TextMatrix(.Rows - 1, 1) = strValue
                .TextMatrix(.Rows - 1, 2) = strType
            Next
            .Col = 1
            .FillStyle = flexFillRepeat
            .Select 0, 1, .Rows - 1, 1
            .CellFontBold = True
            .FillStyle = flexFillSingle
            .Row = 0
        End With
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.PopulateLinkedInputsGrid"

End Sub

Private Function HandleLinkedParm(ByVal nInputItem&, ByRef strName$) As String
On Error GoTo ErrSection:

    Dim strValue$, strType$, strParm$, s$
    Dim aSymbols As cGdArray
    Dim bNew As Boolean

    If m.EditedIndicator Is Nothing Or m.Chart Is Nothing Then
        InfBox "Internal error - unable to process linked input." & vbCrLf & "Please enter a value instead."
        GoTo ErrExit
    End If
    
    If Not m.EditedIndicator.LinkedParmAllowed(nInputItem) Then
        InfBox "Linked input cannot be used for input of this type."
        GoTo ErrExit
    End If
    
    If strName = "<Linked Input>" Then
        s = "i=? ; get=str ; msg=Enter linked input name starting with " & Chr(38) & " ; header=Linked Input Name"
        s = InfBox(s)
        If Len(s) > 0 Then
            If Left(s, 1) = Chr(38) Then
                strName = s
            Else
                strName = Chr(38) & s
            End If
        Else
            GoTo ErrExit
        End If
    End If
    
    m.Chart.LinkedInputGet strName, strValue, strType
    If Len(strValue) <= 0 Then
        strType = m.EditedIndicator.ParmType(nInputItem)
        strParm = m.EditedIndicator.Parm(nInputItem)
        If Len(strParm) > 0 Then strValue = strParm
    End If
    
    If Len(strValue) > 0 Then
        If m.EditedIndicator.LinkedParmSet(nInputItem, strName, strValue) Then
            'double check that parm was set correctly
            If m.EditedIndicator.Parm(nInputItem) = strValue Then
                HandleLinkedParm = strValue
                If Len(strParm) > 0 Then
                    'this is a new linked input - add it to the chart's template global list of linked inputs
                    strType = m.EditedIndicator.ParmType(nInputItem)
                    m.Chart.LinkedInputSet strName, strValue, strType
                    PopulateLinkedInputsGrid
                End If
                m.Chart.RedoMode = eRedo3_Settings
            Else
                InfBox "Linked input mismatched."
                GoTo ErrExit
            End If
        Else
            InfBox "Invalid linked input type."
            GoTo ErrExit
        End If
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmChartCfg.HandleLinkedParm"

End Function

Public Sub NewFunctionAdded(ByVal strExpression$, ByVal strName$, ByVal iFuncReturnType&)
On Error GoTo ErrSection:

    Dim strEditedIndName As String

    If Not m.EditedIndicator Is Nothing Then strEditedIndName = m.EditedIndicator.Name      '6509

    If strEditedIndName <> strName Or (Len(strEditedIndName) = 0 And Len(strName) = 0) Then     '6579
        'user added a new indicator to selected pane or attached new indicator to selected indicator
        If iFuncReturnType = 4 Then
            AddToChart eAdd_Indicator, eAddMode5_NewFunction, , strExpression, strName
        ElseIf iFuncReturnType = 3 Then
            AddToChart eAdd_HighlightBars, eAddMode5_NewFunction, , strExpression, strName
        End If
    End If
    
    Unload frmFunctionMgrCT
    
    If Not Me.Visible Then cmdOK_Click

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.NewFunctionAdded"

End Sub

Private Sub SyncPowerZones()
On Error GoTo ErrSection:

    Dim i&, iParm&, iIndLabelMode&
    Dim Ind As cIndicator
    
    ' only sync things if the edited indicator is the "group leader" of a PowerZones indicator group
    If m.EditedIndicator Is Nothing Then Exit Sub
    If UCase(m.EditedIndicator.Name) <> "POWERZONES" Then Exit Sub
    If Not m.EditedIndicator.IAmGroupLeader Then Exit Sub
    
    ' look for all the other PowerZone indicators in the same group
    iIndLabelMode = kNullData
    For i = 1 To m.Chart.Tree.Count
        If m.Chart.Tree.NodeLevel(i) > 0 Then
            Set Ind = m.Chart.Tree(i)
            If Not Ind Is Nothing Then
                If UCase(Ind.Name) = "POWERZONES" Then
                    If Ind.GroupKey = m.EditedIndicator.GroupKey And Not Ind.IAmGroupLeader Then
                        ' set to same color and display type
                        Ind.Color = m.EditedIndicator.Color
                        Ind.DisplayType = m.EditedIndicator.DisplayType
                        ' set IndLabelMode to same as first one after the group leader
                        If iIndLabelMode = kNullData Then
                            iIndLabelMode = Ind.IndLabelMode
                        Else
                            Ind.IndLabelMode = iIndLabelMode
                        End If
                        ' and set all parms except for ZoneID to the same values
                        If Ind.ParmCount = m.EditedIndicator.ParmCount Then
                            For iParm = 1 To Ind.ParmCount
                                If UCase(Ind.ParmName(iParm)) <> "ZONEID" Then
                                    Ind.Parm(iParm) = m.EditedIndicator.Parm(iParm)
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        End If
    Next
    Set Ind = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.SyncPowerZones"
End Sub

Private Sub SaveModifiedInputs()
On Error GoTo ErrSection:

    Dim j&, i&, s$, bModified As Boolean
    
    If Not m.aModifiedInputs Is Nothing And Not m.EditedIndicator Is Nothing Then
        For j = 0 To m.aModifiedInputs.Size - 1
            s = m.aModifiedInputs(j)
            ' make sure it's for the same indicator
            i = Val(Parse(s, Chr(27), 1))
            If i = m.EditedIndicator.Data.ArrayHandle Then
                ' save Value for Parm#
                i = Val(Parse(s, Chr(27), 2))
                If i > 0 Then
                    s = Parse(s, Chr(27), 3)
                    If m.EditedIndicator.Parm(i) <> s Then
                        m.EditedIndicator.Parm(i) = s
                        bModified = True
                    End If
                End If
            ElseIf IsIDE Then
                ' BUG!!!
                InfBox "This is a BUG -- show it to Tim!", "e", , "ERROR"
            End If
        Next
        m.aModifiedInputs.Size = 0
        
        If bModified Then
            If Not m.Chart Is Nothing Then m.Chart.RedoMode = eRedo3_Settings
            m.iGenerateChart = True
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.SaveModifiedInputs"
End Sub

Private Sub PopulateArtCbo(Indicator As cIndicator)
On Error GoTo ErrSection:

    Dim i&
    If Indicator Is Nothing Then Exit Sub

    If Indicator.DataType = eINDIC_ProfileVolume Then
        With cboPyramidStyle
            .Clear
            .AddItem "Volume (filled)"
            .AddItem "Volume (hollow)"
            .AddItem "Volume (outline)"
            .AddItem "Time (filled blocks)"
            .AddItem "Time (hollow blocks)"
            .AddItem "None"
            .ListIndex = 0
            .Width = 1800
            .Left = 1020
        End With
        i = Indicator.ProfileStyleTPO - 3
        If i >= 0 And i < cboPyramidStyle.ListCount Then
            cboPyramidStyle.ListIndex = i
        Else
            cboPyramidStyle.ListIndex = 0
        End If
    Else
        With cboPyramidStyle
            .Clear
            .AddItem "(default)"
            .AddItem "Thin"
            .AddItem "Medium Thin"
            .AddItem "Medium"
            .AddItem "Medium Thick"
            .AddItem "Thick"
            .AddItem "Extra Thick"
            .AddItem "Dashed (Large)"
            .AddItem "Dashed (Small)"
            .AddItem "Dash Dot"
            .Top = 1320
            .Left = 1320
            .Width = 1500
        End With
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.PopulateArtCbo"

End Sub

Private Sub FixProfileChkboxes(Ind As cIndicator)
On Error GoTo ErrSection

    Dim i&

    If Ind Is Nothing Then Exit Sub
    
    With Ind
        i = .ChkboxValForProfileStruct(ePCStruct_Open, False)   'show/hide current VA,POC values
        If i = vbGrayed Then
            chkVolume_VA(2).Value = vbUnchecked
            chkVolume_VA(2).Enabled = False
        Else
            chkVolume_VA(2).Value = i
            chkVolume_VA(2).Enabled = True
        End If
        i = .ChkboxValForProfileStruct(ePCStruct_Open, True)   'show/hide previous VA, POC values
        If i = vbGrayed Then
            chkVolume_VA(3).Value = vbUnchecked
            chkVolume_VA(3).Enabled = False
        Else
            chkVolume_VA(3).Value = i
            chkVolume_VA(3).Enabled = True
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartCfg.FixProfileChkboxes"

End Sub

