VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmChart2 
   ClientHeight    =   12720
   ClientLeft      =   165
   ClientTop       =   -615
   ClientWidth     =   14025
   Icon            =   "Chart2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12720
   ScaleWidth      =   14025
   Begin VB.Timer tmrProfileLoad 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   11880
   End
   Begin vsOcx6LibCtl.vsElastic vseSeasonal 
      Height          =   9495
      Left            =   11040
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   16748
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      BorderWidth     =   0
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
      Begin HexUniControls.ctlUniFrameWL fraTrendCycle 
         Height          =   4215
         Left            =   30
         TabIndex        =   6
         Top             =   2385
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
         Caption         =   "Chart2.frx":000C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":0054
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":0074
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkOverlayCycles 
            Height          =   255
            Left            =   1320
            TabIndex        =   13
            Top             =   3240
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
            Caption         =   "Chart2.frx":0090
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":00BE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":00DE
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkOverlayTrends 
            Height          =   255
            Left            =   1425
            TabIndex        =   24
            Top             =   270
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
            Caption         =   "Chart2.frx":00FA
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":0128
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0148
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboTrendStyle 
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   1800
            Width           =   1410
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
            Tip             =   "Chart2.frx":0164
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0184
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboTrendStyle 
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   1170
            Width           =   1410
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
            Tip             =   "Chart2.frx":01A0
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":01C0
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboTrendStyle 
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   2430
            Width           =   1410
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
            Tip             =   "Chart2.frx":01DC
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":01FC
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboTrendStyle 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   540
            Width           =   1410
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
            Tip             =   "Chart2.frx":0218
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0238
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTrendShow 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   270
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
            Caption         =   "Chart2.frx":0254
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":0286
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":02A6
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTrendShow 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   900
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
            Caption         =   "Chart2.frx":02C2
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":02FC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":031C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTrendShow 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   1530
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
            Caption         =   "Chart2.frx":0338
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":0372
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0392
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTrendShow 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   50
            Top             =   2160
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
            Caption         =   "Chart2.frx":03AE
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":03E8
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0408
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkShowCycles 
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   3240
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
            Caption         =   "Chart2.frx":0424
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":0454
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0474
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor gdTrendColor 
            Height          =   315
            Index           =   0
            Left            =   1605
            TabIndex        =   62
            Top             =   540
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP lblGradient 
            Height          =   255
            Left            =   120
            Top             =   3600
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
            Caption         =   "Chart2.frx":0490
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":04D0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":04F0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblGradientFrom 
            Height          =   255
            Left            =   120
            Top             =   3900
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
            Caption         =   "Chart2.frx":050C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":0534
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0554
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblGradientTo 
            Height          =   255
            Left            =   1320
            Top             =   3900
            Width           =   255
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":0570
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":0594
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":05B4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label15 
            Height          =   255
            Left            =   180
            Top             =   2950
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
            Caption         =   "Chart2.frx":05D0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":0606
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0626
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            X1              =   0
            X2              =   2330
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   0
            X2              =   120
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000010&
            X1              =   1235
            X2              =   2330
            Y1              =   3045
            Y2              =   3045
         End
      End
      Begin HexUniControls.ctlUniFrameWL Frame1 
         Height          =   2295
         Left            =   30
         TabIndex        =   63
         Top             =   0
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
         Caption         =   "Chart2.frx":0642
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":0682
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":06A2
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboImageXP cboBarType 
            Height          =   315
            Left            =   915
            TabIndex        =   64
            Top             =   1485
            Width           =   1335
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
            Tip             =   "Chart2.frx":06BE
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":06DE
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboCycle 
            Height          =   315
            Left            =   630
            TabIndex        =   65
            Top             =   495
            Width           =   1605
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
            Tip             =   "Chart2.frx":06FA
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":071A
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCycleNum 
            Height          =   315
            Left            =   120
            TabIndex        =   66
            Top             =   495
            Width           =   495
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "Chart2.frx":0736
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
            Tip             =   "Chart2.frx":0758
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0778
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSeasonalApply 
            Default         =   -1  'True
            Height          =   330
            Left            =   240
            TabIndex        =   67
            Top             =   1890
            Width           =   1890
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
            Caption         =   "Chart2.frx":0794
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":07CE
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":07EE
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectDate gdSeasonalDateFrom 
            Height          =   315
            Left            =   120
            TabIndex        =   70
            Top             =   1080
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            MaxDate         =   42611
            MaxDateIsToday  =   -1  'True
            Value           =   2
         End
         Begin HexUniControls.ctlUniLabelXP Label12 
            Height          =   255
            Left            =   90
            Top             =   1530
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
            Caption         =   "Chart2.frx":080A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":083C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":085C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label11 
            Height          =   255
            Left            =   120
            Top             =   270
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
            Caption         =   "Chart2.frx":0878
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":08B2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":08D2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   195
            Left            =   120
            Top             =   870
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
            Caption         =   "Chart2.frx":08EE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":0918
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0938
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin vsOcx6LibCtl.vsElastic vseSymbolLink 
      Height          =   240
      Left            =   12000
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Symbol Link"
      Top             =   630
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ForeColor       =   16777215
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   "S"
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
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
   Begin vsOcx6LibCtl.vsElastic vsePeriodLink 
      Height          =   240
      Left            =   12315
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Period Link"
      Top             =   645
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ForeColor       =   16777215
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   "P"
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
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
   Begin vsOcx6LibCtl.vsElastic vseDetach 
      Height          =   240
      Left            =   12615
      TabIndex        =   98
      TabStop         =   0   'False
      ToolTipText     =   "Detach chart"
      Top             =   660
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ForeColor       =   16777215
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Picture         =   "Chart2.frx":0954
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
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
   Begin VB.PictureBox pbTbBackDraw 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12240
      Index           =   0
      Left            =   13620
      ScaleHeight     =   12240
      ScaleWidth      =   405
      TabIndex        =   94
      Top             =   480
      Visible         =   0   'False
      Width           =   400
      Begin VB.Image imgTbBackDraw 
         Height          =   3675
         Index           =   0
         Left            =   60
         Stretch         =   -1  'True
         Top             =   6825
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin vsOcx6LibCtl.vsElastic vseMouse 
      Height          =   255
      Left            =   1815
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   675
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      Caption         =   "vseMouse"
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   0
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
   Begin vsOcx6LibCtl.vsElastic vseDay 
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   675
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   450
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
      Caption         =   "Wed"
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
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
   Begin vsOcx6LibCtl.vsElastic vseOrderBar 
      Height          =   12735
      Left            =   5940
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   22463
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
      BorderWidth     =   2
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
      Begin HexUniControls.ctlUniComboImageXP cboExchanges 
         Height          =   315
         Left            =   2640
         TabIndex        =   75
         Top             =   10920
         Visible         =   0   'False
         Width           =   1290
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
         Tip             =   "Chart2.frx":0AAE
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":0B2E
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraOrdWizard 
         Height          =   4725
         Left            =   2760
         TabIndex        =   84
         Top             =   5430
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
         Caption         =   "Chart2.frx":0B4A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":0B82
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":0BA2
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdRefreshGraph 
            Height          =   375
            Left            =   75
            TabIndex        =   101
            Top             =   3930
            Width           =   1320
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
            Caption         =   "Chart2.frx":0BBE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":0BF8
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0C18
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboRiskGraphType 
            Height          =   315
            Left            =   75
            TabIndex        =   100
            Top             =   3525
            Width           =   1320
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
            Tip             =   "Chart2.frx":0C34
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0C54
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdTicket 
            Height          =   375
            Left            =   75
            TabIndex        =   88
            Top             =   2670
            Width           =   1320
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
            Caption         =   "Chart2.frx":0C70
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":0CAA
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":0CCA
            RightToLeft     =   0   'False
         End
         Begin vsOcx6LibCtl.vsElastic vseBuyWizard 
            Height          =   540
            Left            =   45
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   0
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   953
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
            BackColor       =   16752800
            ForeColor       =   -2147483630
            FloodColor      =   192
            ForeColorDisabled=   -2147483631
            Picture         =   "Chart2.frx":0CE6
            Caption         =   "Click chart to BUY"
            Align           =   0
            Appearance      =   1
            AutoSizeChildren=   0
            BorderWidth     =   5
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   7
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   0   'False
            PicturePos      =   0
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            _GridInfo       =   ""
         End
         Begin vsOcx6LibCtl.vsElastic vseSellWizard 
            Height          =   540
            Left            =   45
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   540
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   953
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
            BackColor       =   10526975
            ForeColor       =   -2147483630
            FloodColor      =   192
            ForeColorDisabled=   -2147483631
            Picture         =   "Chart2.frx":1000
            Caption         =   "Click chart to SELL"
            Align           =   0
            Appearance      =   1
            AutoSizeChildren=   0
            BorderWidth     =   5
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   7
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   0   'False
            PicturePos      =   0
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            _GridInfo       =   ""
         End
         Begin HexUniControls.ctlUniLabelXP lblRiskGraph 
            Height          =   195
            Left            =   0
            Top             =   3255
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
            Caption         =   "Chart2.frx":131A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":134E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":136E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblMultiLeg 
            Height          =   195
            Left            =   0
            Top             =   1185
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
            Caption         =   "Chart2.frx":138A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":13C6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":13E6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraBrokerDisconnect 
         Height          =   4920
         Left            =   3150
         TabIndex        =   119
         Top             =   180
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
         Caption         =   "Chart2.frx":1402
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":1448
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":1468
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniLabelXP lblBrokerDisconnect 
            Height          =   1890
            Left            =   105
            Top             =   1425
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
            Caption         =   "Chart2.frx":1484
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":14F2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1512
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraPatternProfit 
         Height          =   8625
         Left            =   2340
         TabIndex        =   104
         Top             =   1455
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Chart2.frx":152E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":156E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":158E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdOptimizePFP 
            Height          =   315
            Left            =   465
            TabIndex        =   112
            Top             =   720
            Width           =   1575
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
            Caption         =   "Chart2.frx":15AA
            BackColor       =   12632064
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":15E2
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1602
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdMatchesPFP 
            Height          =   315
            Left            =   465
            TabIndex        =   111
            Top             =   405
            Width           =   1575
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
            Caption         =   "Chart2.frx":161E
            BackColor       =   16777088
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":1656
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1676
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdClosePFP 
            Height          =   315
            Left            =   465
            TabIndex        =   110
            Top             =   1065
            Width           =   1575
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
            Caption         =   "Chart2.frx":1692
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":16BC
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":16DC
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtForecastPFP 
            Height          =   330
            Left            =   1155
            TabIndex        =   109
            Top             =   1515
            Width           =   375
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "Chart2.frx":16F8
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
            Tip             =   "Chart2.frx":171A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":173A
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCorrPercentPFP 
            Height          =   330
            Left            =   1620
            TabIndex        =   108
            Top             =   1994
            Width           =   375
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "Chart2.frx":1756
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
            Tip             =   "Chart2.frx":1778
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1798
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdNewPatternPFP 
            Cancel          =   -1  'True
            Height          =   315
            Left            =   465
            TabIndex        =   107
            Top             =   90
            Width           =   1575
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
            Caption         =   "Chart2.frx":17B4
            BackColor       =   16777088
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":17F0
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1810
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkApplyFilter 
            Height          =   195
            Left            =   1155
            TabIndex        =   106
            Top             =   2805
            Visible         =   0   'False
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
            Caption         =   "Chart2.frx":182C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":1864
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1884
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdPfpSettings 
            Height          =   330
            Left            =   1620
            TabIndex        =   105
            Top             =   1515
            Width           =   805
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
            Caption         =   "Chart2.frx":18A0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":18D0
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":18F0
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label14 
            Height          =   285
            Left            =   2055
            Top             =   2055
            Width           =   165
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":190C
            BackColor       =   12632256
            ForeColor       =   -2147483640
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":192E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":194E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label10 
            Height          =   240
            Left            =   90
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
            Caption         =   "Chart2.frx":196A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":19A4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":19C4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblHitsFoundPFP 
            Height          =   270
            Left            =   45
            Top             =   4395
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
            Caption         =   "Chart2.frx":19E0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":1A0C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1A2C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPfpInd 
            Height          =   270
            Left            =   45
            Top             =   2805
            Width           =   2160
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":1A48
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":1A7C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1A9C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPatternLen 
            Height          =   255
            Left            =   90
            Top             =   2400
            Width           =   2325
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":1AB8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":1AF8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1B18
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label13 
            Height          =   255
            Left            =   90
            Top             =   2032
            Width           =   1560
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":1B34
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":1B7C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":1B9C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin VB.PictureBox pbRight 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   2685
         Picture         =   "Chart2.frx":1BB8
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   96
         ToolTipText     =   "Switch to Options mode"
         Top             =   615
         Width           =   250
      End
      Begin VB.PictureBox pbLeft 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1650
         Picture         =   "Chart2.frx":1F42
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   95
         ToolTipText     =   "Switch to Order mode"
         Top             =   615
         Width           =   250
      End
      Begin HexUniControls.ctlUniFrameWL fraOrderBarMode 
         Height          =   450
         Left            =   1950
         TabIndex        =   82
         Top             =   105
         Width           =   1050
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Chart2.frx":22CC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":230A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":232A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniLabelXP lblOrderBarMode 
            Height          =   195
            Left            =   0
            Top             =   163
            Width           =   1050
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":2346
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":237A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":239A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraOrderBtns 
         Height          =   12000
         Left            =   60
         TabIndex        =   39
         Top             =   615
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
         Caption         =   "Chart2.frx":23B6
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":23EE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":240E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraRithmic 
            Height          =   375
            Left            =   0
            TabIndex        =   76
            Top             =   9825
            Visible         =   0   'False
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
            Caption         =   "Chart2.frx":242A
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "Chart2.frx":245E
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":247E
            RightToLeft     =   0   'False
            Begin VB.Image imgOmne 
               Height          =   105
               Left            =   120
               Picture         =   "Chart2.frx":249A
               Stretch         =   -1  'True
               Top             =   210
               Width           =   1110
            End
            Begin VB.Image imgRithmic 
               Height          =   180
               Left            =   120
               Picture         =   "Chart2.frx":27B1
               Top             =   0
               Width           =   1005
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraTSO 
            Height          =   855
            Left            =   0
            TabIndex        =   81
            Top             =   11160
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":29C7
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "Chart2.frx":29F3
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":2A13
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniButtonImageXP cmdTSOEdit 
               Height          =   330
               Left            =   45
               TabIndex        =   83
               Top             =   60
               Width           =   1260
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
               Caption         =   "Chart2.frx":2A2F
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "Chart2.frx":2A61
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "Chart2.frx":2A81
               RightToLeft     =   0   'False
            End
            Begin vsOcx6LibCtl.vsElastic vseTSO4 
               Height          =   375
               Left            =   990
               TabIndex        =   87
               TabStop         =   0   'False
               Top             =   420
               Width           =   300
               _ExtentX        =   529
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
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "4"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
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
            Begin vsOcx6LibCtl.vsElastic vseTSO3 
               Height          =   375
               Left            =   675
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   420
               Width           =   300
               _ExtentX        =   529
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
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "3"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
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
            Begin vsOcx6LibCtl.vsElastic vseTSO2 
               Height          =   375
               Left            =   345
               TabIndex        =   97
               TabStop         =   0   'False
               Top             =   420
               Width           =   300
               _ExtentX        =   529
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
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "2"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
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
            Begin vsOcx6LibCtl.vsElastic vseTSO1 
               Height          =   375
               Left            =   30
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   420
               Width           =   300
               _ExtentX        =   529
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
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "1"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
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
         End
         Begin HexUniControls.ctlUniFrameWL fraExitFavorites 
            Height          =   495
            Left            =   0
            TabIndex        =   113
            Top             =   10440
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":2A9D
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "Chart2.frx":2ADD
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":2AFD
            RightToLeft     =   0   'False
            Begin vsOcx6LibCtl.vsElastic vseExitA 
               Height          =   375
               Left            =   30
               TabIndex        =   114
               TabStop         =   0   'False
               Top             =   60
               Width           =   300
               _ExtentX        =   529
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
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "A"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
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
            Begin vsOcx6LibCtl.vsElastic vseExitB 
               Height          =   375
               Left            =   350
               TabIndex        =   115
               TabStop         =   0   'False
               Top             =   60
               Width           =   300
               _ExtentX        =   529
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
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "B"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
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
            Begin vsOcx6LibCtl.vsElastic vseExitC 
               Height          =   375
               Left            =   670
               TabIndex        =   116
               TabStop         =   0   'False
               Top             =   60
               Width           =   300
               _ExtentX        =   529
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
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "C"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
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
            Begin vsOcx6LibCtl.vsElastic vseExitD 
               Height          =   375
               Left            =   990
               TabIndex        =   117
               TabStop         =   0   'False
               Top             =   60
               Width           =   300
               _ExtentX        =   529
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
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "D"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
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
         End
         Begin HexUniControls.ctlUniCheckXP chkConfirmOrder 
            Height          =   195
            Left            =   0
            TabIndex        =   118
            Top             =   6825
            Width           =   1410
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":2B19
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":2B55
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":2B75
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP lblTradePos 
            Height          =   225
            Left            =   0
            TabIndex        =   103
            Top             =   8155
            Width           =   1320
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   -1  'True
            Text            =   "Chart2.frx":2B91
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
            Tip             =   "Chart2.frx":2BC7
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":2BE7
         End
         Begin vsOcx6LibCtl.vsElastic vseBracketOrder 
            Height          =   540
            Left            =   0
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   9180
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   953
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
            BackColor       =   10526975
            ForeColor       =   -2147483630
            FloodColor      =   192
            ForeColorDisabled=   -2147483631
            Picture         =   "Chart2.frx":2C03
            Caption         =   ""
            Align           =   0
            Appearance      =   1
            AutoSizeChildren=   0
            BorderWidth     =   5
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   7
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
         Begin HexUniControls.ctlUniButtonImageXP cmdBuyMarket 
            Height          =   420
            Left            =   0
            TabIndex        =   45
            Top             =   60
            Width           =   1320
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
            Caption         =   "Chart2.frx":4D5D
            BackColor       =   16752800
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":4D91
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":4DB1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkAutoJournal 
            Height          =   180
            Left            =   0
            TabIndex        =   77
            Top             =   7065
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
            Caption         =   "Chart2.frx":4DCD
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":4E05
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":4E25
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraPrices 
            Height          =   555
            Left            =   -20
            TabIndex        =   61
            Top             =   5460
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
            Caption         =   "Chart2.frx":4E41
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "Chart2.frx":4E73
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":4E93
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniLabelXP Label9 
               Height          =   195
               Left            =   60
               Top             =   0
               Width           =   420
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Chart2.frx":4EAF
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "Chart2.frx":4ED7
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "Chart2.frx":4EF7
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label8 
               Height          =   195
               Left            =   60
               Top             =   180
               Width           =   420
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Chart2.frx":4F13
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "Chart2.frx":4F3B
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "Chart2.frx":4F5B
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label6 
               Height          =   195
               Left            =   60
               Top             =   360
               Width           =   360
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Chart2.frx":4F77
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "Chart2.frx":4F9F
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "Chart2.frx":4FBF
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblBid 
               Height          =   195
               Left            =   240
               Top             =   360
               Width           =   1020
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Chart2.frx":4FDB
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "Chart2.frx":5001
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "Chart2.frx":5021
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblMarket 
               Height          =   195
               Left            =   240
               Top             =   180
               Width           =   1020
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Chart2.frx":503D
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "Chart2.frx":5063
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "Chart2.frx":5083
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblAsk 
               Height          =   195
               Left            =   240
               Top             =   0
               Width           =   1020
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Chart2.frx":509F
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "Chart2.frx":50C5
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "Chart2.frx":50E5
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdClearQty 
            Height          =   300
            Left            =   0
            TabIndex        =   72
            Top             =   920
            Width           =   235
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
            Caption         =   "Chart2.frx":5101
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5123
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5143
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkAutoExit 
            Height          =   195
            Left            =   0
            TabIndex        =   71
            Top             =   7440
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
            Caption         =   "Chart2.frx":515F
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5193
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":51B3
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin vsOcx6LibCtl.vsElastic vseBuyChart 
            Height          =   540
            Left            =   0
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2085
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   953
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
            BackColor       =   16752800
            ForeColor       =   -2147483630
            FloodColor      =   192
            ForeColorDisabled=   -2147483631
            Picture         =   "Chart2.frx":51CF
            Caption         =   "Click chart to BUY"
            Align           =   0
            Appearance      =   1
            AutoSizeChildren=   0
            BorderWidth     =   5
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   7
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   0   'False
            PicturePos      =   0
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            _GridInfo       =   ""
         End
         Begin vsOcx6LibCtl.vsElastic cmdBailout 
            Height          =   375
            Left            =   0
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "Cancel all orders and exit current position"
            Top             =   4320
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   1
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   600
            BackColor       =   192
            ForeColor       =   65535
            FloodColor      =   192
            ForeColorDisabled=   -2147483631
            Caption         =   "FLATTEN"
            Align           =   0
            Appearance      =   1
            AutoSizeChildren=   0
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   4
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
         Begin HexUniControls.ctlUniButtonImageXP cmdCancelAll 
            Height          =   360
            Left            =   15
            TabIndex        =   41
            Top             =   3960
            Width           =   1320
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
            Caption         =   "Chart2.frx":54E9
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":551D
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":555F
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdQty1 
            Height          =   300
            Left            =   0
            TabIndex        =   53
            Top             =   1365
            Width           =   420
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
            Caption         =   "Chart2.frx":557B
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":559D
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":562B
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdQty2 
            Height          =   300
            Left            =   420
            TabIndex        =   52
            Top             =   1365
            Width           =   420
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
            Caption         =   "Chart2.frx":5647
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5669
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":56F7
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdQty3 
            Height          =   300
            Left            =   840
            TabIndex        =   51
            Top             =   1365
            Width           =   480
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
            Caption         =   "Chart2.frx":5713
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5737
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":57C5
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboOrderType 
            Height          =   315
            Left            =   480
            TabIndex        =   49
            Top             =   1725
            Width           =   840
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
            Tip             =   "Chart2.frx":57E1
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":584F
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin vsOcx6LibCtl.vsElastic vseSellChart 
            Height          =   540
            Left            =   0
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   2625
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   953
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
            BackColor       =   10526975
            ForeColor       =   -2147483630
            FloodColor      =   192
            ForeColorDisabled=   -2147483631
            Picture         =   "Chart2.frx":586B
            Caption         =   "Click chart to SELL"
            Align           =   0
            Appearance      =   1
            AutoSizeChildren=   0
            BorderWidth     =   5
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   7
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   0   'False
            PicturePos      =   0
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            _GridInfo       =   ""
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSellMarket 
            Height          =   420
            Left            =   0
            TabIndex        =   44
            Top             =   480
            Width           =   1320
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
            Caption         =   "Chart2.frx":5B85
            BackColor       =   10526975
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5BBB
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5BDB
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdReverse 
            Height          =   360
            Left            =   15
            TabIndex        =   43
            Top             =   3540
            Width           =   1320
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
            Caption         =   "Chart2.frx":5BF7
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5C25
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5C45
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtTradeQty 
            Height          =   300
            Left            =   260
            TabIndex        =   40
            Top             =   920
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   16777215
            ForeColor       =   0
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "Chart2.frx":5C61
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
            Tip             =   "Chart2.frx":5C81
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5CCF
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSellBid 
            Height          =   360
            Left            =   0
            TabIndex        =   58
            Top             =   6420
            Visible         =   0   'False
            Width           =   1320
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
            Caption         =   "Chart2.frx":5CEB
            BackColor       =   10526975
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5D21
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5D41
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdBuyBid 
            Height          =   360
            Left            =   0
            TabIndex        =   56
            Top             =   6060
            Visible         =   0   'False
            Width           =   1320
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
            Caption         =   "Chart2.frx":5D5D
            BackColor       =   16752800
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5D91
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5DB1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdBuyAsk 
            Height          =   360
            Left            =   0
            TabIndex        =   55
            Top             =   4740
            Visible         =   0   'False
            Width           =   1320
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
            Caption         =   "Chart2.frx":5DCD
            BackColor       =   16752800
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5E01
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5E21
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSellAsk 
            Height          =   360
            Left            =   0
            TabIndex        =   57
            Top             =   5100
            Visible         =   0   'False
            Width           =   1320
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
            Caption         =   "Chart2.frx":5E3D
            BackColor       =   10526975
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5E73
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5E93
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdExitPos 
            Height          =   360
            Left            =   1125
            TabIndex        =   42
            Top             =   3930
            Visible         =   0   'False
            Width           =   1320
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
            Caption         =   "Chart2.frx":5EAF
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":5EDD
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5EFD
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdScrollBar vscrQty 
            Height          =   300
            Left            =   1095
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   920
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   529
            Value           =   1
         End
         Begin HexUniControls.ctlUniLabelXP lblQty 
            Height          =   195
            Left            =   0
            Top             =   3225
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
            Caption         =   "Chart2.frx":5F19
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":5F4B
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5F6B
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblEquity 
            Height          =   255
            Left            =   0
            Top             =   8385
            Width           =   1320
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Chart2.frx":5F87
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   1
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":5FB7
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":5FD7
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAutoExit 
            Height          =   615
            Left            =   0
            Top             =   7680
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
            Caption         =   "Chart2.frx":5FF3
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":601B
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":603B
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblOrderType 
            Height          =   255
            Left            =   0
            Top             =   1785
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
            Caption         =   "Chart2.frx":6057
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":6081
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":60A1
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   0
         TabIndex        =   46
         Top             =   240
         Width           =   1320
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
         Tip             =   "Chart2.frx":60BD
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":60DD
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraFrontMonth 
         Height          =   6255
         Left            =   1605
         TabIndex        =   59
         Top             =   735
         Visible         =   0   'False
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
         Caption         =   "Chart2.frx":60F9
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":6125
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":6145
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdRollNow 
            Height          =   435
            Left            =   60
            TabIndex        =   74
            Top             =   2955
            Width           =   1260
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
            Caption         =   "Chart2.frx":6161
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":6191
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":61B1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdContracts 
            Height          =   435
            Left            =   60
            TabIndex        =   73
            Top             =   2370
            Width           =   1260
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
            Caption         =   "Chart2.frx":61CD
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":6209
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":6229
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblFrontMonth 
            Height          =   2340
            Left            =   195
            Top             =   3525
            Visible         =   0   'False
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
            Caption         =   "Chart2.frx":6245
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":6345
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":6365
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblOpenOrderPos 
            Height          =   1890
            Left            =   150
            Top             =   540
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
            Caption         =   "Chart2.frx":6381
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":640B
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":642B
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniLabelXP lblExchange 
         Height          =   255
         Left            =   3000
         Top             =   10560
         Visible         =   0   'False
         Width           =   1290
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Chart2.frx":6447
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "Chart2.frx":6479
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":6499
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAccounts 
         Height          =   195
         Left            =   0
         Top             =   0
         Width           =   1320
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Chart2.frx":64B5
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "Chart2.frx":64F5
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":6515
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraWizardPrompt 
      Height          =   465
      Left            =   1665
      TabIndex        =   89
      Top             =   8400
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
      Caption         =   "Chart2.frx":6531
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   8388608
      Tip             =   "Chart2.frx":656F
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Chart2.frx":658F
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPut 
         Height          =   420
         Left            =   1185
         TabIndex        =   92
         Top             =   30
         Width           =   465
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
         Caption         =   "Chart2.frx":65AB
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Chart2.frx":65D1
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":65F1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCall 
         Height          =   420
         Left            =   30
         TabIndex        =   91
         Top             =   30
         Width           =   465
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
         Caption         =   "Chart2.frx":660D
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Chart2.frx":6635
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":6655
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblWizardPrice 
         Height          =   435
         Left            =   495
         Top             =   15
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
         Caption         =   "Chart2.frx":6671
         BackColor       =   16777088
         ForeColor       =   -2147483640
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "Chart2.frx":66A5
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":66C5
         RightToLeft     =   0   'False
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox pbTbBack 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   14025
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   14025
      Begin HexUniControls.ctlUniComboBoxXP cboBarPeriod 
         Height          =   315
         Left            =   3810
         TabIndex        =   80
         Top             =   75
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
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
         Tip             =   "Chart2.frx":66E1
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
         MouseIcon       =   "Chart2.frx":6701
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin VB.PictureBox pbNotUsed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   180
         ScaleHeight     =   300
         ScaleWidth      =   510
         TabIndex        =   93
         Top             =   120
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Image imgTbBack 
         Height          =   585
         Index           =   0
         Left            =   7635
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5295
      End
   End
   Begin vsOcx6LibCtl.vsElastic vseInvalid 
      Height          =   1275
      Left            =   1140
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   2249
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      Caption         =   "Test a long caption to see how it wraps"
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
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
   Begin VSFlex7LCtl.VSFlexGrid fgChartFlex 
      Height          =   495
      Index           =   0
      Left            =   540
      TabIndex        =   68
      Top             =   5760
      Visible         =   0   'False
      Width           =   4440
      _cx             =   7832
      _cy             =   873
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
   Begin VB.Timer tmrGameMode 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1980
      Top             =   1020
   End
   Begin vsOcx6LibCtl.vsElastic vseTipChart 
      Height          =   240
      Left            =   4635
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   765
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
      Begin HexUniControls.ctlUniLabelXP lblTipChart 
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
         Caption         =   "Chart2.frx":671D
         BackColor       =   -2147483624
         ForeColor       =   -2147483625
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "Chart2.frx":6749
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":6769
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VB.PictureBox pbChart 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   1335
      TabIndex        =   11
      Top             =   660
      Visible         =   0   'False
      Width           =   1395
   End
   Begin gdOCX.gdScrollBar hsb 
      Height          =   240
      Left            =   0
      TabIndex        =   7
      Top             =   3420
      Visible         =   0   'False
      Width           =   4600
      _ExtentX        =   8123
      _ExtentY        =   423
      Horizontal      =   -1  'True
      ScrollTipAlign  =   1
   End
   Begin vsOcx6LibCtl.vsElastic vseScrollSep 
      Height          =   60
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3360
      Width           =   4600
      _ExtentX        =   8123
      _ExtentY        =   106
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
      BackColor       =   12632064
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
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2520
      Top             =   1080
   End
   Begin vsOcx6LibCtl.vsElastic vseTipY 
      Height          =   240
      Left            =   3840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
      _ExtentX        =   873
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
      Caption         =   "vseTipY"
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
      Begin HexUniControls.ctlUniLabelXP lblTipY 
         Height          =   132
         Left            =   60
         Top             =   60
         Width           =   372
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Chart2.frx":6785
         BackColor       =   -2147483624
         ForeColor       =   -2147483640
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "Chart2.frx":67B1
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":67D1
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdSettings 
      Height          =   255
      Left            =   4620
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
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
      Caption         =   "Chart2.frx":67ED
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "Chart2.frx":6815
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "Chart2.frx":6835
      RightToLeft     =   0   'False
   End
   Begin vsOcx6LibCtl.vsElastic vseTipX 
      Height          =   240
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
      _ExtentX        =   873
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
      Begin HexUniControls.ctlUniLabelXP lblTipX 
         Height          =   132
         Left            =   60
         Top             =   60
         Width           =   372
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Chart2.frx":6851
         BackColor       =   -2147483624
         ForeColor       =   -2147483625
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "Chart2.frx":687D
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":689D
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsElastic vseCaption 
      Height          =   240
      Left            =   3600
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483646
      ForeColor       =   -2147483639
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Picture         =   "Chart2.frx":68B9
      Caption         =   " NavSuite - frmChart2 (Form)"
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   0   'False
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   1
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin vsOcx6LibCtl.vsIndexTab vsTab 
      Height          =   1515
      Left            =   240
      TabIndex        =   8
      Top             =   3900
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   2672
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
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
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      CurrTab         =   0
      FirstTab        =   0
      Style           =   0
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   0   'False
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
   End
   Begin vsOcx6LibCtl.vsElastic vseReplay 
      Height          =   1080
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   1905
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ForeColor       =   -2147483639
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
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
      Begin HexUniControls.ctlUniFrameWL fraReplayControls 
         Height          =   675
         Left            =   0
         TabIndex        =   30
         Top             =   180
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
         Caption         =   "Chart2.frx":72CB
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":730D
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":732D
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdBack 
            Height          =   360
            Left            =   120
            TabIndex        =   34
            Top             =   210
            Width           =   435
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
            Caption         =   "Chart2.frx":7349
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":7377
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":73D1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdStop 
            Height          =   360
            Left            =   600
            TabIndex        =   33
            Top             =   210
            Width           =   435
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
            Caption         =   "Chart2.frx":73ED
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":741B
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7497
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdForward 
            Height          =   360
            Left            =   1560
            TabIndex        =   32
            Top             =   210
            Width           =   435
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
            Caption         =   "Chart2.frx":74B3
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":74E7
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7527
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdPlay 
            Height          =   360
            Left            =   1080
            TabIndex        =   31
            Top             =   210
            Width           =   435
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
            Caption         =   "Chart2.frx":7543
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":7571
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":75EF
            RightToLeft     =   0   'False
         End
         Begin MSComctlLib.Slider sldSpeed 
            Height          =   435
            Left            =   2460
            TabIndex        =   35
            ToolTipText     =   "Select replay speed (hotkey: Up/Down arrows)"
            Top             =   150
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   767
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   7
            SelStart        =   4
            Value           =   4
            TextPosition    =   1
         End
         Begin HexUniControls.ctlUniLabelXP Label2 
            Height          =   195
            Left            =   3870
            Top             =   270
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
            Caption         =   "Chart2.frx":760B
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":7633
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7653
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   195
            Left            =   2040
            Top             =   270
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
            Caption         =   "Chart2.frx":766F
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":7697
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":76B7
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraReplayProfit 
         Height          =   675
         Left            =   4400
         TabIndex        =   17
         Top             =   180
         Width           =   10500
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Chart2.frx":76D3
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "Chart2.frx":7711
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "Chart2.frx":7731
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdRevPos 
            Height          =   420
            Left            =   6318
            TabIndex        =   69
            Top             =   180
            Width           =   1275
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
            Caption         =   "Chart2.frx":774D
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":7783
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":77F5
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdReport 
            Height          =   420
            Left            =   8666
            TabIndex        =   23
            Top             =   180
            Width           =   795
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
            Caption         =   "Chart2.frx":7811
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":783D
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":789F
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdBuy 
            Height          =   420
            Left            =   3900
            TabIndex        =   22
            Top             =   180
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
            Caption         =   "Chart2.frx":78BB
            BackColor       =   16752800
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":78E1
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":792D
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSell 
            Height          =   420
            Left            =   4642
            TabIndex        =   21
            Top             =   180
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
            Caption         =   "Chart2.frx":7949
            BackColor       =   10526975
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":7971
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":79BD
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdExitNow 
            Height          =   420
            Left            =   5384
            TabIndex        =   20
            Top             =   180
            Width           =   915
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
            Caption         =   "Chart2.frx":79D9
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":7A09
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7A71
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdAutoExits 
            Height          =   420
            Left            =   7612
            TabIndex        =   19
            Top             =   180
            Width           =   1035
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
            Caption         =   "Chart2.frx":7A8D
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":7AC1
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7B33
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdQuitReplay 
            Height          =   420
            Left            =   9480
            TabIndex        =   18
            Top             =   180
            Width           =   735
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
            Caption         =   "Chart2.frx":7B4F
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "Chart2.frx":7B77
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7BBD
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblProfit 
            Height          =   195
            Left            =   60
            Top             =   390
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
            Caption         =   "Chart2.frx":7BD9
            BackColor       =   -2147483633
            ForeColor       =   255
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":7C13
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7C33
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label7 
            Height          =   195
            Left            =   240
            Top             =   150
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
            Caption         =   "Chart2.frx":7C4F
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":7C87
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7CA7
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPosition 
            Height          =   195
            Left            =   2460
            Top             =   390
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
            Caption         =   "Chart2.frx":7CC3
            BackColor       =   -2147483633
            ForeColor       =   12582912
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":7D03
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7D23
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   195
            Left            =   2460
            Top             =   150
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
            Caption         =   "Chart2.frx":7D3F
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":7D83
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7DA3
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label5 
            Height          =   195
            Left            =   1200
            Top             =   150
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
            Caption         =   "Chart2.frx":7DBF
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":7DF9
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7E19
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblOpenEquity 
            Height          =   195
            Left            =   1440
            Top             =   390
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
            Caption         =   "Chart2.frx":7E35
            BackColor       =   -2147483633
            ForeColor       =   32768
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "Chart2.frx":7E67
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "Chart2.frx":7E87
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Begin VB.Menu mnuSettings 
         Caption         =   "&EDIT Chart Settings"
      End
      Begin VB.Menu mnuChartCopyMove 
         Caption         =   "Copy/Move Chart"
      End
      Begin VB.Menu mnuDetAttChart 
         Caption         =   "Detach Chart"
      End
      Begin VB.Menu mnuShowToolbar 
         Caption         =   "Show toolbar"
      End
      Begin VB.Menu mnuHidePane 
         Caption         =   "&Hide Pane"
      End
      Begin VB.Menu mnuTemplates 
         Caption         =   "Apply &Template"
         Begin VB.Menu mnuTemplateManage 
            Caption         =   "< Manage chart templates >"
         End
         Begin VB.Menu mnuTemplateToOtherCharts 
            Caption         =   "Copy these settings to all other charts"
         End
         Begin VB.Menu mnuTemplateSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTemplate 
            Caption         =   "(no templates found)"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBarPeriods 
         Caption         =   "&Bar Time Period"
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "1 minute"
            Index           =   0
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "5 minute"
            Index           =   1
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "10 minute"
            Index           =   2
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "15 minute"
            Index           =   3
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "30 minute"
            Index           =   4
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "60 minute"
            Index           =   5
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "&Daily"
            Index           =   7
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "&Weekly"
            Index           =   8
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "&Monthly"
            Index           =   9
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "&Quarterly"
            Index           =   10
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "&Yearly"
            Index           =   11
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "-"
            Index           =   12
         End
         Begin VB.Menu mnuBarPeriod 
            Caption         =   "< &Custom >"
            Index           =   13
         End
      End
      Begin VB.Menu mnuBarDisplay 
         Caption         =   "Bar &Display Type"
         Begin VB.Menu mnuOHLC 
            Caption         =   "&OHLC bars"
         End
         Begin VB.Menu mnuCandlesticks 
            Caption         =   "&Candlesticks"
         End
         Begin VB.Menu mnuCloseLine 
            Caption         =   "Close &Line"
         End
         Begin VB.Menu mnuBollinger 
            Caption         =   "Bollinger Bars"
         End
      End
      Begin VB.Menu mnuPixelsPerBar 
         Caption         =   "Bar &Spacing (more/less bars)"
         Begin VB.Menu mnuPPB 
            Caption         =   "1 pixel"
            Index           =   0
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "3 pixels"
            Index           =   1
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "5 pixels"
            Index           =   2
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "7 pixels"
            Index           =   3
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "9 pixels"
            Index           =   4
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "11 pixels"
            Index           =   5
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "15 pixels"
            Index           =   6
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "21 pixels"
            Index           =   7
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "29 pixels"
            Index           =   8
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "39 pixels"
            Index           =   9
         End
         Begin VB.Menu mnuPPB 
            Caption         =   "51 pixels"
            Index           =   10
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCursor 
         Caption         =   "&Cursor"
         Begin VB.Menu mnuCursorArrow 
            Caption         =   "&Arrow"
         End
         Begin VB.Menu mnuCrosshairs 
            Caption         =   "&Crosshairs"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuCursorHoriz 
            Caption         =   "&Horizontal line"
         End
         Begin VB.Menu mnuCursorVert 
            Caption         =   "&Vertical line"
         End
      End
      Begin VB.Menu mnuAnnotS 
         Caption         =   "&Annotations"
         Begin VB.Menu mnuDeleteAnnot 
            Caption         =   "Delete &Last Annotation  (hotkey: Delete)"
            Index           =   0
         End
         Begin VB.Menu mnuSepAnnot 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHideAnnots 
            Caption         =   "&Hide Annotations  (toggles)"
            Index           =   0
         End
         Begin VB.Menu mnuSepAnnotFlags 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClearAnnotMultiFlag 
            Caption         =   "Clear Show on All Charts"
         End
      End
      Begin VB.Menu mnuUnzoom 
         Caption         =   "&Zoom Out"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoScale 
         Caption         =   "Auto-scale the price pane"
      End
      Begin VB.Menu mnuLogScale 
         Caption         =   "Semi-Log scale for Price pane"
      End
      Begin VB.Menu mnuLogModeDraw 
         Caption         =   "   Use semi-log for Drawing tools"
      End
      Begin VB.Menu mnuTips 
         Caption         =   "&Floating price/date tips"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHideScrollbars 
         Caption         =   "Hide scrollbars"
      End
      Begin VB.Menu mnuCoarseGrid 
         Caption         =   "Coarse vertical &Grid (dates)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUnsplit 
         Caption         =   "&Unsplit prices"
      End
      Begin VB.Menu mnuDisableRT 
         Caption         =   "Disable realtime"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrderBar 
         Caption         =   "Show Order Bar"
      End
      Begin VB.Menu mnuAccountBar 
         Caption         =   "Show Account Bar"
      End
      Begin VB.Menu mnuOrdBarSettings 
         Caption         =   "Settings for Order/Account Bars"
      End
      Begin VB.Menu mnuOrdBarDefaults 
         Caption         =   "Apply last saved Defaults"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrades 
         Caption         =   "Apply Trading Strategy to chart"
      End
      Begin VB.Menu mnuAutoTrade 
         Caption         =   "Auto-Trade Strategy on chart"
      End
      Begin VB.Menu mnuEditSystem 
         Caption         =   "Edit Trading Strategy"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Chart"
      End
      Begin VB.Menu mnuDataCopy 
         Caption         =   "C&opy chart data to clipboard"
      End
      Begin VB.Menu mnuSaveImage 
         Caption         =   "E&xport Chart (save image)"
      End
      Begin VB.Menu mnuScreenCapture 
         Caption         =   "Screen Capture (entire screen)"
      End
      Begin VB.Menu mnuChartCapture 
         Caption         =   "Screen Capture (this chart only)"
      End
      Begin VB.Menu mnuMoreDataSingleSym 
         Caption         =   "Get additional history"
      End
      Begin VB.Menu mnuHotKeys 
         Caption         =   "View charting Hot &Keys and Tips"
      End
   End
   Begin VB.Menu mnuAnnotEdit 
      Caption         =   "mnuAnnotEdit"
      Begin VB.Menu mnuAnnotMovePt 
         Caption         =   "Move point"
      End
      Begin VB.Menu mnuAnnotDeletePt 
         Caption         =   "Delete point"
      End
      Begin VB.Menu mnuAnnotAddPt 
         Caption         =   "Add point"
      End
      Begin VB.Menu mnuAnnotStyle 
         Caption         =   "Change style"
      End
      Begin VB.Menu mnuAnnotDuplicate 
         Caption         =   "Duplicate"
      End
   End
   Begin VB.Menu mnuOrderAction 
      Caption         =   "mnuOrderAction"
      Begin VB.Menu mnuOrderEdit 
         Caption         =   "Edit Order"
      End
      Begin VB.Menu mnuOrderSubmit 
         Caption         =   "Submit Order"
      End
      Begin VB.Menu mnuOrderCancel 
         Caption         =   "Cancel Order"
      End
      Begin VB.Menu mnuOrderPark 
         Caption         =   "Park Order"
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuManageXOS 
         Caption         =   "Manage Exit Order Strategies"
      End
      Begin VB.Menu mnuSelectXOS 
         Caption         =   "Select Exit Order Strategy"
      End
      Begin VB.Menu mnuRemoveXOS 
         Caption         =   "Remove Exit Order Strategy"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrderHistory 
         Caption         =   "Order History"
      End
      Begin VB.Menu mnuOrderJournal 
         Caption         =   "Journal for Order"
      End
      Begin VB.Menu mnuOrderCheckStatus 
         Caption         =   "Check Order Status"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrderAcctHistory 
         Caption         =   "Account History"
      End
      Begin VB.Menu mnuViewJournals 
         Caption         =   "View Journals"
      End
   End
   Begin VB.Menu mnuContractAction 
      Caption         =   "mnuContractAction"
      Begin VB.Menu mnuContracts 
         Caption         =   "contracts"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmChart2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE: the code for frmChart and frmChart2 are meant to
' be EXACTLY identical -- therefore it is usually easiest
' to make all desired changes to frmChart, then copy and
' paste ALL the code from frmChart over into frmChart2.
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/21/2010   MJM         removed the frame control "fraMiscCtls"
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private m As frmChartPrivateType

Public Property Get FlattenInProgress() As Boolean
    FlattenInProgress = g.FlattenQueue.IsGettingFlattened(g.Broker.AccountNumberForID(m.Chart.TradeAccountID), m.Chart.TradeSymbol, 0&, eGDFlattenQueueOperation_Flatten)
End Property

Public Property Get CancelAllInProgress() As Boolean
    CancelAllInProgress = g.FlattenQueue.IsGettingFlattened(g.Broker.AccountNumberForID(m.Chart.TradeAccountID), m.Chart.TradeSymbol, 0&, eGDFlattenQueueOperation_CancelAll)
End Property

Public Property Get ReverseInProgress() As Boolean
    ReverseInProgress = g.FlattenQueue.IsGettingFlattened(g.Broker.AccountNumberForID(m.Chart.TradeAccountID), m.Chart.TradeSymbol, 0&, eGDFlattenQueueOperation_Reverse)
End Property

Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    Static bInProgress As Boolean

    Dim i&, nAccountID&
    
    Dim ePrevAcctType As eGDTypeOfAccount
    Dim eSymbolPitType As eFutureSymbolType
    
    If bInProgress Then Exit Sub
    If Not GameMode Is Nothing Then Exit Sub
    
    bInProgress = True              'aardvark 6541
    
    If Not m.Chart Is Nothing Then
        nAccountID = m.Chart.TradeAccountID
        eSymbolPitType = m.Chart.SymbolPitType
        ePrevAcctType = TypeOfAccount(TradeAccountID)
        
        With cboAccounts
            If nAccountID <> .ItemData(.ListIndex) Then
                If True Then
                    m.Chart.TradeAccountID = .ItemData(.ListIndex)
                    m.Chart.SetTrackerTradesReload
                    m.Chart.GenerateChart eRedo3_Settings
                    SetFocusCtl
                    
                    SetAutoExit
                    
                    If vseOrderBar.Visible Then
                        If eSymbolPitType = ePitSymbol Then
                            If ePrevAcctType <> TypeOfAccount(TradeAccountID) Then
                                ToggleOrderBar True, True, eSymbolPitType
                            End If
                        ElseIf InStr(m.Chart.Symbol, "-0") Then
                            ToggleOrderBar True, True, eSymbolPitType       '6430
                        End If
                    End If
                    
                    FixOrderBarControls False, True
                    GetContractInformation
                Else
                    For i = 0 To .ListCount - 1
                        If .ItemData(i) = nAccountID Then
                            .ListIndex = i
                            Exit For
                        End If
                    Next
                End If
            End If
        End With
    End If
    
ErrExit:
    bInProgress = False
    Exit Sub

ErrSection:
    bInProgress = False
    RaiseError Me.Name & ".cboAccounts_Click"
    Resume ErrExit
    
End Sub

Private Sub cboBarPeriod_Click()
On Error GoTo ErrSection:
    
    BarPeriodClick Me, m.oBtnMouseLast, False
    m.Chart.ChangeBarPeriod cboBarPeriod.Text
    MoveFocus pbChart

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cboBarPeriod_Click"

End Sub

Private Sub cboBarPeriod_DropDown()
On Error GoTo ErrSection:
    
    BarPeriodClick Me, m.oBtnMouseLast, True

ErrExit:
    Exit Sub


ErrSection:
    RaiseError Me.Name & ".cboBarPeriod_DropDown"

End Sub

Private Sub cboBarPeriod_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyReturn Then cboBarPeriod_Click        '5166

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cboBarPeriod_KeyUp"

End Sub

Private Sub cboBarType_Click()
On Error Resume Next:
    
    If Not Me.Visible Then Exit Sub
    If m.eSeasonalCtrlsState <> eSeasonCtrlStatus_Updated And Not vseInvalid.Visible Then Exit Sub
    HandleSeasonalInput eSeasonalCtrl_BarType
End Sub

Private Sub cboBarType_DropDown()
    tmr.Enabled = False
End Sub

Private Sub cboCycle_Click()
On Error Resume Next:
    
    If Not Me.Visible Then Exit Sub
    If m.eSeasonalCtrlsState <> eSeasonCtrlStatus_Updated And Not vseInvalid.Visible Then Exit Sub
    HandleSeasonalInput eSeasonalCtrl_CycleType

End Sub

Private Sub cboCycle_DropDown()
    tmr.Enabled = False
End Sub

Private Sub cboOrderType_Click()
On Error GoTo ErrSection:

    If Not m.Chart Is Nothing Then m.Chart.PseudoOrderType = cboOrderType.ListIndex

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cboOrderType_Click"

End Sub

Private Sub cboRiskGraphType_Click()
On Error GoTo ErrSection:

    Dim Chart As cChart
    Dim Pane As cPane
    
    Dim dMin#, dMax#, dIncrement#, dTemp#, iPoints&
    Dim strLegs$, strType$, strMsg$
    Dim i&
    
    If fraBrokerDisconnect.Visible Then Exit Sub
    If m.eOrdBarMode <> eOrdBarMode_Wizard Then Exit Sub
    If cboRiskGraphType.ListIndex < 0 Then Exit Sub
    If cboRiskGraphType.ListIndex >= cboRiskGraphType.ListCount Then Exit Sub
    
    Set Chart = m.Chart
    If Chart Is Nothing Then Exit Sub
    
    Set Pane = Chart.Tree("Price Pane")
    If Pane Is Nothing Then Exit Sub

'risk graph type
    strType = cboRiskGraphType.Text
    strType = Replace(strType, "/", "")     'strip the '/' from profit/loss
    
    'the probability curve takes mutiple seconds to calculate in OE so don't request if unchanged
    If m.strPrevRiskGraph = "Probability" And strType = "Probability" Then Exit Sub
    
    If strType = "None" Then
        OptNavGraphClear
        Exit Sub
    End If
    
    m.strPrevRiskGraph = strType
    
'price range information
    dMin = Pane.geAdjustMin
    dMax = Pane.geAdjustMax
    dTemp = Chart.Bars.MinMove * 100
    
    'dMin = dMin - dTemp
    'dMax = dMax + dTemp
    
    iPoints = Int((dMax - dMin) / Chart.Bars.MinMove)       '# of minmove between min & max
    
    While iPoints > 1000
        iPoints = Int(iPoints / 10)        'divide by 10 until less than 1000 "data points"
    Wend
    If iPoints < 200 Then iPoints = 200
    
    dIncrement = (dMax - dMin) / iPoints    'get increment in price value for number of points wanted
    
    
    dTemp = Chart.Bars.MinMove              'round everything to minimimu move
    
    dMin = RoundToMinMove(dMin, dTemp)
    dMax = RoundToMinMove(dMax, dTemp)
    dIncrement = RoundToMinMove(dIncrement, dTemp)
    
'option legs
    strLegs = WizardGridLegInfo(Me, False)
    
'build message to send to OptNav
    If strType = "Probability" Then
        OptNavGraphClear 1          'clear graph, but keep blank split pane area so don't see left-right "jerk" effect when data comes back
        dIncrement = m.Chart.Bars(eBARS_Close, m.Chart.LastGoodDataBar(False))
        strMsg = Str(Chart.geChartObj) & vbTab & strType & vbTab & Str(dMin) & ";" & Str(dMax) & ";" & Str(dIncrement) & vbTab
        strMsg = strMsg & "45" & vbTab & m.Chart.TradeSymbol
        
        RequestRiskGraphStructure strMsg
        InfBox "Requesting data, please wait ...", "t", "", "Options Risk Graph", True, 3, , , , , , , , 5
    ElseIf Len(strLegs) > 0 Then
        strMsg = Str(Chart.geChartObj) & vbTab & strType & vbTab & Str(dMin) & ";" & Str(dMax) & ";" & Str(dIncrement) & vbTab
        strMsg = strMsg & strLegs
            
        RequestRiskGraphStructure strMsg
    Else
        OptNavGraphClear
    End If
    
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cboRiskGraphType_Click"

End Sub

Private Sub cboTrendStyle_Click(Index As Integer)
On Error Resume Next:
    
    Select Case Index
        Case 0
            HandleSeasonalInput eSeasonalCtrl_AvgTrendStyle
        Case 1
            HandleSeasonalInput eSeasonalCtrl_BullTrendStyle
        Case 2
            HandleSeasonalInput eSeasonalCtrl_BearTrendStyle
        Case 3
            HandleSeasonalInput eSeasonalCtrl_CurrCycleStyle
    End Select

End Sub

Private Sub cboTrendStyle_DropDown(Index As Integer)
    tmr.Enabled = False
End Sub

Private Sub chkApplyFilter_Click()
On Error GoTo ErrSection:

    Dim i&, j&, k&, s$
    Dim Annot As cAnnotation
    Dim fgIndicators As VSFlexGrid
    
    If g.bUnloading Or g.bStarting Or m.oPatternProfit Is Nothing Then GoTo ErrExit
    
    Set fgIndicators = fgChartFlex(eFlexGridIdx_PfpInd)
    If fgIndicators Is Nothing Then GoTo ErrExit
    If fgIndicators.Rows < 4 Then GoTo ErrExit
    
    Set Annot = m.oPatternProfit.PatternAnnotCheck(m.Chart, Me, False)
    If chkApplyFilter.Tag = "NoValidate" Then
        If Not Annot Is Nothing Then m.Chart.GenerateChart eRedo1_Scrolled
        GoTo ErrExit
    End If
    
    If chkApplyFilter.Value = vbChecked Then
        
        If m.oPatternProfit.StandardDev = 1 Then
            For i = 0 To 3
                s = UCase(fgIndicators.TextMatrix(i, 0))
                If s = "CLOSE" Then
                    j = j + 1
                ElseIf s = "NONE" Or s = "-999" Then
                    k = k + 1
                End If
            Next
            s = ""
            If j <> 1 Or k <> 3 Then
                s = "Heat map can only be used with Close." & vbCrLf & "Okay to change settings?"
            End If
        Else
            s = "Heat map can only be used with standard deviation = 1. Okay to change settings?"
        End If

        If Len(s) > 0 Then
            If chkApplyFilter.Tag = "NoPrompt" Then
                chkApplyFilter.Value = vbUnchecked
            ElseIf InfBox(s, "i", "+OK|-Cancel", "Pattern Forecasting") = "C" Then
                m.oPatternProfit.Heatmap = vbUnchecked
                chkApplyFilter.Value = vbUnchecked
                If Not Annot Is Nothing Then m.Chart.GenerateChart eRedo1_Scrolled
            Else
                With fgIndicators
                    .TextMatrix(0, 0) = "Close"
                    .Cell(flexcpData, 0, 0) = "#-888;Close"
    
                    .TextMatrix(1, 0) = "None"
                    .Cell(flexcpData, 1, 0) = "#-999;None"
    
                    .TextMatrix(2, 0) = "None"
                    .Cell(flexcpData, 2, 0) = "#-999;None"
    
                    .TextMatrix(3, 0) = "None"
                    .Cell(flexcpData, 3, 0) = "#-999;None"
                End With
                m.oPatternProfit.Heatmap = vbChecked
                cmdOptimizePFP_Click
            End If
        Else
            m.oPatternProfit.Heatmap = vbChecked
            If Not Annot Is Nothing Then m.Chart.GenerateChart eRedo1_Scrolled
        End If
    
    Else
        m.oPatternProfit.Heatmap = vbUnchecked
        If Not Annot Is Nothing Then m.Chart.GenerateChart eRedo1_Scrolled
    End If

'JM 8/27/2015: this checkbox originally used for testing, changed on 8/27/2015 to use for PFP heatmap display
'    If fgChartFlex(eFlexGridIdx_PfpHits).Rows <= fgChartFlex(eFlexGridIdx_PfpHits).FixedRows Then GoTo ErrExit
'    If Not m.oPatternProfit Is Nothing Then
'        strResult = m.oPatternProfit.FindMatches(m.Chart, fgChartFlex(eFlexGridIdx_PfpInd), fgChartFlex(eFlexGridIdx_PfpHits), ValOfText(txtCorrPercentPFP.Text))
'    End If
'    If Len(strResult) > 0 Then InfBox strResult, "E", , "Pattern for Profit"
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".chkApplyFilter_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAutoExit_Click
'' Description: When the user changes the auto exit, sync up the world
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAutoExit_Click()
On Error GoTo ErrSection:

    Dim strAutoExit As String           ' Auto Exit selected
    
    If Visible And (Not m.bSettingAutoExit) Then
        If chkAutoExit.Value = vbChecked Then
            strAutoExit = ActivateAutoExit(TradeAccountID, SymbolID, "Chart")
            If Len(strAutoExit) > 0 Then
                lblAutoExit.Caption = FileBase(strAutoExit)
            Else
                chkAutoExit.Value = vbUnchecked
                lblAutoExit.Caption = "None"
            End If
        Else
            g.Broker.BrokerDebug g.Broker.AccountTypeForID(TradeAccountID), "Auto Exit Deactivate from Chart (" & m.Chart.TradeSymbol & ", " & g.Broker.AccountNumberForID(TradeAccountID) & "): Done"
            g.OrderStrategies.DeactivateExit TradeAccountID, SymbolID, True, "Turned off on Chart"
            lblAutoExit.Caption = "None"
        End If
        
        If lblAutoExit.Caption = "None" And lblAutoExit.Enabled Then ExitCtrlAppearance Me, Nothing, ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".chkAutoExit_Click"
    
End Sub

Private Sub chkAutoJournal_Click()
On Error GoTo ErrSection:

    If Visible Then
        g.Broker.AutoJournalPopUp = (chkAutoJournal.Value = vbChecked)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".chkAutoJournal_Click"

End Sub

Private Sub chkConfirmOrder_Click()
On Error GoTo ErrSection:
    
    g.Broker.ConfirmManual = CheckBoxValue(chkConfirmOrder)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".chkConfirmOrder_Click"

End Sub

Private Sub chkOverlayCycles_Click()
On Error Resume Next
    
    HandleSeasonalInput eSeasonalCtrl_OverlayCycles

End Sub

Private Sub chkOverlayTrends_Click()
On Error Resume Next
    
    HandleSeasonalInput eSeasonalCtrl_OverlayTrends

End Sub

Private Sub chkShowCycles_Click()
On Error Resume Next
        
        HandleSeasonalInput eSeasonalCtrl_ShowCycles

End Sub

Private Sub chkTrendShow_Click(Index As Integer)
On Error Resume Next:

    If m.eSeasonalCtrlsState <> eSeasonCtrlStatus_Updated Then Exit Sub

    If chkTrendShow(0).Value = vbUnchecked And chkTrendShow(1).Value = vbUnchecked And _
        chkTrendShow(2).Value = vbUnchecked And chkTrendShow(3).Value = vbUnchecked And _
        Me.chkShowCycles.Value = vbUnchecked Then
        
        chkTrendShow(0).Value = vbChecked
        HandleSeasonalInput eSeasonalCtrl_AvgTrendCheckBox
    End If
    
    Select Case Index
        Case 0
            HandleSeasonalInput eSeasonalCtrl_AvgTrendCheckBox
        Case 1
            HandleSeasonalInput eSeasonalCtrl_BullTrendCheckBox
        Case 2
            HandleSeasonalInput eSeasonalCtrl_BearTrendCheckBox
        Case 3
            HandleSeasonalInput eSeasonalCtrl_CurrCycleCheckBox
    End Select
    
End Sub

Private Sub cmdAutoExits_Click()
On Error GoTo ErrSection:

    Dim eReplayModeSave As eGDReplayMode

    eReplayModeSave = m.oGameMode.GameReplayMode(tmrGameMode.Enabled)
    cmdStop_Click
    frmGameTargetLoss.ShowMe m.oGameMode, Me, Nothing
    MoveFocus sldSpeed
    
    If eReplayModeSave = eGDReplayMode_Play Then cmdPlay_Click
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdAutoExits_Click"
    Resume ErrExit

End Sub

Private Sub cmdBack_Click()
On Error GoTo ErrSection:

    m.oGameMode.UserForwardBack = -1
    cmdStop_Click
    
    m.oGameMode.GameBack
    UpdateGameModeLabels
    m.Chart.GenerateChart eRedo1_Scrolled
    MoveFocus sldSpeed
    m.oGameMode.UserForwardBack = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdBack_Click"
    Resume ErrExit

End Sub

Private Sub cmdBailOut_Click()
On Error GoTo ErrSection:

    cmdBailout.Enabled = False
    cmdBailout.BackColor = Me.BackColor
    
    g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.Chart.TradeAccountID), "Flattening Position for " & m.Chart.TradeSymbol & " in account " & g.Broker.AccountNameForID(m.Chart.TradeAccountID) & " from the Chart", True
    FlattenForSymbol m.Chart.TradeAccountID, m.Chart.TradeSymbolID(True), 0&
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdBailOut_Click"
    
End Sub

Private Sub cmdBrokerConnect_Click()
On Error GoTo ErrSection:
    
    g.Broker.Connect g.Broker.AccountTypeForID(m.Chart.TradeAccountID)
    If g.Broker.ConnectionStatusForAccount(m.Chart.TradeAccountID) = eGDConnectionStatus_Connected Then
        FixOrderBarControls False, True
        GetContractInformation
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdBrokerConnect_Click"
    
End Sub

Private Sub cmdBuyAsk_Click()
On Error GoTo ErrSection:

    HandleChartOrder True, eClickOrder_BuyAsk
    SetFocusCtl

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdBuyAsk_Click"

End Sub

Private Sub cmdBuyAsk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    cmdBuyAsk.ToolTipText = "Buy " & Trim(txtTradeQty) & " at the current Ask price or better"

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdBuyAsk_MouseMove"

End Sub

Private Sub cmdBuyBid_Click()
On Error GoTo ErrSection:

    HandleChartOrder True, eClickOrder_BuyBid
    SetFocusCtl

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdBuyBid_Click"

End Sub

Private Sub cmdBuyBid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    cmdBuyBid.ToolTipText = "Buy " & Trim(txtTradeQty) & " at the current Bid price or better"

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdBuyBid_MouseMove"

End Sub

Private Sub cmdBuyMarket_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    cmdBuyMarket.ToolTipText = "Buy " & Trim(txtTradeQty) & " at market"

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdBuyMarket_MouseMove"

End Sub

Private Sub cmdBuyMarket_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        HandleBuySellClick cmdBuyMarket, vbRightButton
    Else
        HandleChartOrder True, eClickOrder_BuyMkt
    End If
    
    SetFocusCtl

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdBuyMarket_MouseUp"

End Sub

Private Sub cmdCall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
        
    If m.AnnotOptions Is Nothing Then
        With cmdCall
            If Len(lblWizardPrice.Caption) = 0 Then
                .MousePointer = vbDefault
            Else
                .MousePointer = vbCustom
                .MouseIcon = pbChart.MouseIcon
            End If
        End With
    Else
        lblWizardPrice_MouseMove Button, Shift, X, Y
    End If

End Sub

Private Sub cmdCall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim eOrdType As enumOneClickOrder
    
    Dim Annot As cAnnotation
    Dim Pane As cPane
    Dim bBalloonTool As Boolean
    
    Dim coordInfo As coordinate_info
    Dim pt As POINTAPI
    
    Dim i&, dStrike#, dPrice#, strInfo$
    
    If Not m.AnnotOptions Is Nothing Then
        Set Annot = m.Chart.ClosestOptionAnnot(m.MouseLast.dDate)
        If Not Annot Is Nothing Then bBalloonTool = Annot.IsBalloonOptionsInfo
    End If
    
    If bBalloonTool Then
        If Button = vbRightButton Then
            ClearBuySellButtons
        ElseIf Not m.AnnotOptions Is Nothing Then
            dStrike = Val(Parse(lblWizardPrice.Caption, vbCrLf, 1))
            strInfo = Parse(lblWizardPrice.Caption, vbCrLf, 2)
            strInfo = Mid(strInfo, 2, Len(strInfo) - 1)
            dPrice = Val(strInfo)
            
            If m.AnnotOptions.BalloonPutStrike > 0 And m.AnnotOptions.BalloonPutCost > 0 Then
                'set values before clearing data & resetting flags
                m.AnnotOptions.BalloonCallStrike = dStrike
                m.AnnotOptions.BalloonCallCost = dPrice
                m.AnnotOptions.BalloonExpiration = Annot.dDate(1)
                
                ClearBuySellButtons
            Else
                'remove non-selected options expiration vertical lines if applicable
                m.Chart.RemoveAnnots False, eANNOT_VertLine, eANNOT_OptionInfo, False, Annot
                m.Chart.GenerateChart eRedo1_Scrolled
                'set values after redraw
                m.AnnotOptions.BalloonCallStrike = dStrike
                m.AnnotOptions.BalloonCallCost = dPrice
                m.AnnotOptions.BalloonExpiration = Annot.dDate(1)
                
                Set Pane = m.Chart.Tree("PRICE PANE")
                If Not Pane Is Nothing Then
                    coordInfo.paneId = Pane.gePaneId
                    coordInfo.x_value = m.AnnotOptions.geLeftX(0)
                    coordInfo.y_value = m.AnnotOptions.BalloonStockPrice - (dStrike - m.AnnotOptions.BalloonStockPrice)
                    If coordInfo.y_value < Pane.Min Then coordInfo.y_value = Pane.Min
                    i = geDataToCoord(m.Chart.geChartObj, coordInfo)
                    
                    m.MouseLast.dY = coordInfo.y_value
                    Y = coordInfo.y_pixels * Screen.TwipsPerPixelY
                    fraWizardPrompt.Top = Y + cmdPut.Height / 2
                    
                    pt.X = (fraWizardPrompt.Left + cmdPut.Width / 2) / Screen.TwipsPerPixelX
                    pt.Y = (fraWizardPrompt.Top - cmdPut.Height / 2) / Screen.TwipsPerPixelY
                    i = ClientToScreen(pbChart.hWnd, pt)
                    i = SetCursorPos(pt.X, pt.Y)
                    
                    HandleWizardPrompt X, coordInfo.y_pixels * Screen.TwipsPerPixelY
                End If
                
            End If
        End If

    ElseIf m.nActiveAnnotIdx >= 0 Then
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Not Annot Is Nothing Then
            dStrike = ValOfText(Parse(lblWizardPrice.Caption, vbCrLf, 2))
            strInfo = Annot.ClosestStrike(dStrike)          'price;SymCall,SymPut
            If dStrike = ValOfText(Parse(strInfo, ";", 1)) Then
                If vseBuyChart.Appearance = apInset Or vseBuyWizard.Appearance = apInset Then
                    eOrdType = eClickOrder_BuyCall
                Else
                    eOrdType = eClickOrder_SellCall
                End If
                WizardGridAdd eOrdType, Parse(lblWizardPrice, vbCrLf, 1) & " " & Parse(lblWizardPrice, vbCrLf, 2) & " C", _
                    Parse(strInfo, ";", 2), "L", Annot
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdCall_MouseUp"

End Sub

Private Sub cmdCancelAll_Click()
On Error GoTo ErrSection:

    cmdCancelAll.Enabled = False

    g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.Chart.TradeAccountID), "Cancelling All Orders for " & m.Chart.TradeSymbol & " in account " & g.Broker.AccountNameForID(m.Chart.TradeAccountID) & " from the Chart", True
    CancelAllForSymbol m.Chart.TradeAccountID, m.Chart.TradeSymbol, 0&

ErrExit:
    Exit Sub
        
ErrSection:
    RaiseError Me.Name & ".cmdCancelAll_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClearQty_Click
'' Description: Allow the user to start over with a new order quantity
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearQty_Click()
On Error GoTo ErrSection:

    'txtTradeQty.Text = ""       '10-06-2006 (reverse Harry's request for now)
    m.Quantity.Price = m.Quantity.Min

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdClearQty_Click"

End Sub

Private Sub cmdClosePFP_Click()
    OrderbarWrapper eOrdBarMode_PFP
End Sub

Private Sub cmdContracts_Click()
On Error GoTo ErrSection:

    Dim nSymbolID&

    If UCase(cmdContracts.Caption) = "GO TO CONTRACT" Then
        If Me.mnuContracts.Count = 1 Then
            mnuContracts_Click 0            '5898
        Else
            PopupMenu mnuContractAction
        End If
    Else
        nSymbolID = g.SymbolPool.SymbolIDforSymbol(cmdContracts.Caption)
        
        If nSymbolID > 0 Then
            m.Chart.SetSymbol nSymbolID, True
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdContracts_Click"

End Sub

Private Sub cmdExitNow_Click()
On Error GoTo ErrSection:

    Dim eReplayModeSave As eGDReplayMode
    
    eReplayModeSave = m.oGameMode.GameReplayMode(tmrGameMode.Enabled)
    
    cmdStop_Click
    m.oGameMode.ExitPosition lblPosition.Caption
    UpdateGameModeLabels
    EnableGameControls
    m.Chart.GenerateChart eRedo1_Scrolled
    MoveFocus sldSpeed
    
    If eReplayModeSave = eGDReplayMode_Play Then cmdPlay_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdExitNow_Click"
    Resume ErrExit

End Sub

Private Sub cmdExitPos_Click()
On Error GoTo ErrSection:
    
    Dim strPos$, nQty&
    
    strPos = UCase(Parse(lblTradePos.Text, " ", 1))
    nQty = ValOfText(Parse(lblTradePos.Text, " ", 2))
    If strPos = "LONG" Then
        HandleChartOrder True, eClickOrder_SellMkt, nQty
    ElseIf strPos = "SHORT" Then
        HandleChartOrder True, eClickOrder_BuyMkt, nQty
    'ElseIf g.ChartGlobals.eChartMode = eMode_ChartOrder Then
    '    ToolBarClick m.Chart.tbToolbar.Tools("ID_ZoomIn"), frmMain
    End If
    
ErrExit:
    SetFocusCtl
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdExitPos_Click"
    Resume ErrExit
End Sub

Private Sub cmdExitPos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim strPos$, nQty&
    On Error Resume Next
    strPos = UCase(Parse(lblTradePos.Text, " ", 1))
    nQty = ValOfText(Parse(lblTradePos.Text, " ", 2))
    If strPos = "LONG" Then
        cmdExitPos.ToolTipText = "Sell " & Str(nQty) & " to exit current position"
    ElseIf strPos = "SHORT" Then
        cmdExitPos.ToolTipText = "Buy " & Str(nQty) & " to exit current position"
    Else
        cmdExitPos.ToolTipText = ""
    End If
    

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdExitPos_MouseMove"

End Sub

Private Sub cmdForward_Click()
On Error GoTo ErrSection:

    m.oGameMode.UserForwardBack = 1
    cmdStop_Click
    
    m.oGameMode.GameForward
    UpdateGameModeLabels
    MoveFocus sldSpeed
    m.oGameMode.UserForwardBack = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdForward_Click"
    Resume ErrExit
    
End Sub

Private Sub cmdMatchesPFP_Click()
On Error GoTo ErrSection:

    Dim strResult$

    If m.Chart.Zoomed Then m.Chart.UnzoomChart True
    
    If Len(g.strActiveDraw) > 0 Then
        ClearAnnotFlags True, True
        g.strActiveDraw = ""
        SyncDrawTools
    End If
    
    PfpReset ePfpReset_GridPfp
    PfpReset ePfpReset_PpfAnnotInd
    DoEvents
    
    If Not m.oPatternProfit Is Nothing Then
        PfpReset ePfpReset_Forecastbars
        strResult = m.oPatternProfit.FindMatches(m.Chart, fgChartFlex(eFlexGridIdx_PfpInd), fgChartFlex(eFlexGridIdx_PfpHits), ValOfText(txtCorrPercentPFP.Text))
        If Len(strResult) > 0 Then
            InfBox strResult, "E", , "Pattern for Profit"
        Else
            m.Chart.GenerateChart eRedo1_Scrolled
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdMatchesPFP_Click"

End Sub

Private Sub cmdNewPatternPFP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
'mouse event sequence: MouseDown, MouseClick then MouseUp

    Set g.ChartGlobals.frmPfpSelPattern = Me
    
    If m.Chart Is Nothing Then Exit Sub
    
    If m.Chart.Zoomed Then m.Chart.UnzoomChart True
    
    If Len(g.strActiveDraw) > 0 And InStr(g.strActiveDraw, "PFP") = 0 Then
        ToolbarSetCursorGroup m.Chart.tbToolbar, False
        SyncDrawTools True
    End If

    PfpReset ePfpReset_PpfPattern
    PfpReset ePfpReset_GridPfp
    DoEvents
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdNewPatternPFP_Click"

End Sub

Private Sub cmdNewPatternPFP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If m.Chart Is Nothing Then Exit Sub
    
    If m.Chart.Zoomed Then m.Chart.UnzoomChart True
    
    If Not m.oPatternProfit Is Nothing Then txtCorrPercentPFP.Text = Str(m.oPatternProfit.PercentCorr)
    
    cmdMatchesPFP.Enabled = False
    
    PfpReset ePfpReset_PpfPattern
    PfpReset ePfpReset_GridPfp
    DoEvents
    
    ToolBarClick m.Chart.tbToolbar.Tools("ID_Rectangle"), frmMain, , eTbExtraInfo_PFPNewPattern

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdNewPatternPFP_MouseUp"

End Sub

Private Sub cmdOptimizePFP_Click()
On Error GoTo ErrSection:

    Dim strResult$
    Dim Annot As cAnnotation
    
    If m.Chart.Zoomed Then m.Chart.UnzoomChart True
    
    If Len(g.strActiveDraw) > 0 Then
        ClearAnnotFlags True, True
        g.strActiveDraw = ""
        SyncDrawTools
    End If
    
    If m.oPatternProfit Is Nothing Then
        strResult = "Internal error. Please close then reopen Pattern for Profit and try again."
        InfBox strResult, "E", , "Pattern for Profit"
    ElseIf frmPatternProfitOpt.ShowMe(Me) Then
    
        PfpReset ePfpReset_GridPfp
        PfpReset ePfpReset_PpfAnnotInd
        DoEvents
        
        Set Annot = m.oPatternProfit.PatternAnnotCheck(m.Chart, Me, True)
        PfpReset ePfpReset_Forecastbars
        strResult = m.oPatternProfit.FindMatches(m.Chart, fgChartFlex(eFlexGridIdx_PfpInd), fgChartFlex(eFlexGridIdx_PfpHits), ValOfText(txtCorrPercentPFP.Text), , True)
        
        If Len(strResult) > 0 Or Annot Is Nothing Then
            InfBox strResult, "E", , "Pattern for Profit"
        Else
            m.oPatternProfit.FindMatches m.Chart, fgChartFlex(eFlexGridIdx_PfpInd), fgChartFlex(eFlexGridIdx_PfpHits), ValOfText(txtCorrPercentPFP.Text), True
        End If
        
        If Not Annot Is Nothing Then cmdMatchesPFP.Enabled = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdOptimizePFP_Click"

End Sub

Private Sub cmdPfpSettings_Click()
On Error GoTo ErrSection:

    Dim i&
    
    frmPatProfitSettings.ShowMe Me
    
    If Not m.oPatternProfit Is Nothing Then
        i = m.oPatternProfit.Heatmap
        If i <> vbChecked Then
            If i = vbGrayed Then m.oPatternProfit.Heatmap = vbUnchecked
            If chkApplyFilter.Value = vbChecked Then
                chkApplyFilter.Tag = "NoValidate"
                chkApplyFilter.Value = vbUnchecked
                chkApplyFilter.Tag = ""
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdPfpSettings_Click"

End Sub

Private Sub cmdPlay_Click()
On Error GoTo ErrSection:

    Dim strText$

    If m.oGameMode.GameStart Then
        GameSpeed
        tmrGameMode.Enabled = True
        Enable cmdPlay, False
        Enable cmdStop, True
        Enable cmdReport, True
        If m.Chart.SystemID > 0 And m.Chart.SystemID <> m.oGameMode.GameStrategyID Then
            m.Chart.SystemID = m.oGameMode.GameStrategyID
            m.Chart.ShowTrades = True
        End If
        EnableGameControls
        MoveFocus sldSpeed
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdPlay_Click"
    Resume ErrExit
End Sub

Private Sub cmdPut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If m.AnnotOptions Is Nothing Then
        With cmdPut
            If Len(lblWizardPrice.Caption) = 0 Then
                .MousePointer = vbDefault
            Else
                .MousePointer = vbCustom
                .MouseIcon = pbChart.MouseIcon
            End If
        End With
    Else
        lblWizardPrice_MouseMove Button, Shift, X, Y
    End If

End Sub

Private Sub cmdPut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim eOrdType As enumOneClickOrder
    
    Dim Annot As cAnnotation
    Dim Pane As cPane
    Dim bBalloonTool As Boolean
    
    Dim coordInfo As coordinate_info
    Dim pt As POINTAPI
    
    Dim i&, dStrike#, dPrice#, strInfo$
    
    If Not m.AnnotOptions Is Nothing Then
        Set Annot = m.Chart.ClosestOptionAnnot(m.MouseLast.dDate)
        If Not Annot Is Nothing Then bBalloonTool = Annot.IsBalloonOptionsInfo
    End If
    
    If bBalloonTool Then
        If Button = vbRightButton Then
            ClearBuySellButtons
        ElseIf Not m.AnnotOptions Is Nothing Then
            dStrike = Val(Parse(lblWizardPrice.Caption, vbCrLf, 1))
            strInfo = Parse(lblWizardPrice.Caption, vbCrLf, 2)
            strInfo = Mid(strInfo, 2, Len(strInfo) - 1)
            dPrice = Val(strInfo)
            
            If m.AnnotOptions.BalloonCallStrike > 0 And m.AnnotOptions.BalloonCallCost > 0 Then
                'set values before clearing data & resetting flags
                m.AnnotOptions.BalloonPutStrike = dStrike
                m.AnnotOptions.BalloonPutCost = dPrice
                m.AnnotOptions.BalloonExpiration = Annot.dDate(1)
                
                ClearBuySellButtons
            Else
                'remove non-selected options expiration vertical lines if applicable
                m.Chart.RemoveAnnots False, eANNOT_VertLine, eANNOT_OptionInfo, False, Annot
                m.Chart.GenerateChart eRedo1_Scrolled
                'set values after redraw
                m.AnnotOptions.BalloonPutStrike = dStrike
                m.AnnotOptions.BalloonPutCost = dPrice
                m.AnnotOptions.BalloonExpiration = Annot.dDate(1)
            
                Set Pane = m.Chart.Tree("PRICE PANE")
                If Not Pane Is Nothing Then
                    coordInfo.paneId = Pane.gePaneId
                    coordInfo.x_value = m.AnnotOptions.geLeftX(0)
                    coordInfo.y_value = m.AnnotOptions.BalloonStockPrice + (m.AnnotOptions.BalloonStockPrice - dStrike)
                    If coordInfo.y_value > Pane.Max Then coordInfo.y_value = Pane.Max
                    i = geDataToCoord(m.Chart.geChartObj, coordInfo)
                    
                    m.MouseLast.dY = coordInfo.y_value
                    Y = coordInfo.y_pixels * Screen.TwipsPerPixelY
                    fraWizardPrompt.Top = Y + cmdPut.Height / 2
                    
                    pt.X = (fraWizardPrompt.Left + cmdPut.Width / 2) / Screen.TwipsPerPixelX
                    pt.Y = (fraWizardPrompt.Top - cmdPut.Height / 2) / Screen.TwipsPerPixelY
                    i = ClientToScreen(pbChart.hWnd, pt)
                    i = SetCursorPos(pt.X, pt.Y)
                    
                    HandleWizardPrompt X, coordInfo.y_pixels * Screen.TwipsPerPixelY
                End If
            End If
        End If
    ElseIf cmdPut.Caption = "Put" Then
        If m.nActiveAnnotIdx >= 0 Then
            Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
            If Not Annot Is Nothing Then
                dStrike = ValOfText(Parse(lblWizardPrice.Caption, vbCrLf, 2))
                strInfo = Annot.ClosestStrike(dStrike)          'strike;SymCall,SymPut
                If dStrike = ValOfText(Parse(strInfo, ";", 1)) Then
                    If vseBuyChart.Appearance = apInset Or vseBuyWizard.Appearance = apInset Then
                        eOrdType = eClickOrder_BuyPut
                    Else
                        eOrdType = eClickOrder_SellPut
                    End If
                    WizardGridAdd eOrdType, Parse(lblWizardPrice, vbCrLf, 1) & " " & Parse(lblWizardPrice, vbCrLf, 2) & " P", _
                        Parse(strInfo, ";", 3), "L", Annot
                End If
            End If
        End If
    Else
        lblWizardPrice_MouseUp Button, Shift, X, Y
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdPut_MouseUp"

End Sub

Private Sub cmdQty1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    cmdQty1.SetFocus

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdQty1_MouseDown"

End Sub

Private Sub cmdQty1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    UpdateTradeQuantity Button, m.lPreset1
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdQty1_MouseDown"

End Sub

Private Sub cmdQty2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    cmdQty2.SetFocus

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdQty2_MouseDown"

End Sub

Private Sub cmdQty2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    UpdateTradeQuantity Button, m.lPreset2
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdQty2_MouseUp"

End Sub

Private Sub cmdQty3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    cmdQty3.SetFocus

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdQty3_MouseDown"

End Sub

Private Sub cmdQty3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    UpdateTradeQuantity Button, m.lPreset3
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdQty3_MouseUp"

End Sub

Private Sub cmdQuitReplay_Click()
On Error Resume Next:

    Dim strFiles$
    
    If m.bGameMode Then
        strFiles = g.ChartGlobals.strCPCRoot & "\Charts\Replay.CHT"
        If FileExist(strFiles) Then KillFile strFiles
    End If

    Unload Me
    
End Sub

Private Sub cmdRefreshGraph_Click()
On Error GoTo ErrSection:
    
    Dim i&, strMsg$
    Dim Pane As cPane
    Dim Tree As cGdTree

    If cboRiskGraphType.Text = "None" Then Exit Sub
    
    If Not m.Chart Is Nothing Then
        Set Tree = m.Chart.Tree         '6060
        If Not Tree Is Nothing Then
            For i = 1 To Tree.Count
                If Tree.Key(i) = "PRICE PANE" Then
                ElseIf Tree.NodeLevel(i) = 0 Then
                    Set Pane = Tree(i)
                    If Not Pane Is Nothing Then
                        If Pane.Display Then
                            If Pane.SplitPaneType = ePANE_SplitPaneWood Then
                                strMsg = "Options risk graph cannot be shown when Woodies indicators are on chart."
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
            If Len(strMsg) > 0 Then
                InfBox strMsg, "I", , "Risk Graph"
                GoTo ErrExit
            End If
        End If
    End If
    
    If m.bFlexOrdBar Then
        With fgChartFlex(eFlexGridIdx_OrdWizard)
            For i = .FixedRows To .Rows - 1
                strMsg = .TextMatrix(i, 1)
                If Len(strMsg) > 0 Then Exit For
            Next
        End With
    End If

    If Len(strMsg) = 0 Then
        strMsg = "There are no strategy legs for " & cboRiskGraphType.Text & " graph."
        InfBox strMsg, "I", , "Risk Graph"
        OptNavGraphClear
    Else
        m.strPrevRiskGraph = ""         'to force request to OE for new data
        cboRiskGraphType_Click
    End If
    

ErrExit:
    Set Pane = Nothing
    Set Tree = Nothing
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdRefreshGraph_Click"

End Sub

Private Sub cmdReport_Click()
On Error GoTo ErrSection:
'this is for game mode report only

    GenerateGameReport
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdReport_Click"
    Resume ErrExit
End Sub

Private Sub cmdReverse_Click()
On Error GoTo ErrSection:
    
    cmdReverse.Enabled = False
    g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.Chart.TradeAccountID), "Reversing position for " & m.Chart.TradeSymbol & " in account " & g.Broker.AccountNameForID(m.Chart.TradeAccountID) & " from the Chart", True
    ReverseForSymbol m.Chart.TradeAccountID, m.Chart.TradeSymbolID(True), 0&
    
ErrExit:
    SetFocusCtl
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdReverse_Click"

End Sub

Private Sub cmdReverse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim strPos$, nQty&
    On Error Resume Next
    strPos = UCase(Parse(lblTradePos.Text, " ", 1))
    nQty = ValOfText(Parse(lblTradePos.Text, " ", 2)) * 2
    If strPos = "LONG" Then
        cmdReverse.ToolTipText = "Sell " & Str(nQty) & " to reverse current position"
    ElseIf strPos = "SHORT" Then
        cmdReverse.ToolTipText = "Buy " & Str(nQty) & " to reverse current position"
    Else
        cmdReverse.ToolTipText = ""
    End If
    

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdReverse_MouseMove"

End Sub

Private Sub cmdRevPos_Click()
On Error GoTo ErrSection:

    Dim eReplayModeSave As eGDReplayMode
    
    eReplayModeSave = m.oGameMode.GameReplayMode(tmrGameMode.Enabled)
    
    cmdStop_Click
    m.oGameMode.ExitPosition lblPosition.Caption, True
    UpdateGameModeLabels
    EnableGameControls
    m.Chart.GenerateChart eRedo1_Scrolled
    MoveFocus sldSpeed
    
    If eReplayModeSave = eGDReplayMode_Play Then cmdPlay_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdRevPos_Click"
    Resume ErrExit

End Sub

Private Sub cmdRollNow_Click()
On Error GoTo ErrSection:

    Dim nSymbolID As Long

    If m.bTradeContinuous Then
        If AllowRollContinuous(, True) Then ToggleOrderBar True, True
    Else
        ' for a continuous contract, look up current contract
        If Not m.Chart Is Nothing Then
            If Not m.Chart.Bars Is Nothing Then
                nSymbolID = g.SymbolPool.SymbolIDforSymbol(RollSymbolForDate(m.Chart.Symbol, m.Chart.Bars(eBARS_DateTime, m.Chart.Bars.Size - 1)))
                If nSymbolID > 0 Then
                    m.Chart.SetSymbol nSymbolID, True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdRollNow_Click"

End Sub

Private Sub cmdSeasonalApply_Click()
On Error GoTo ErrSection:

    Dim i&, strBarType$

    tmr.Enabled = False
    m.eSeasonalCtrlsState = eSeasonCtrlStatus_Unpopulated
    
    Select Case cboBarType.ListIndex
        Case 0
            If m.Chart.Bars.Prop(eBARS_PeriodicityStr) <> "Daily" Then strBarType = "Daily"
        Case 1
            If m.Chart.Bars.Prop(eBARS_PeriodicityStr) <> "Weekly" Then strBarType = "Weekly"
        Case 2
            If m.Chart.Bars.Prop(eBARS_PeriodicityStr) <> "Monthly" Then strBarType = "Monthly"
    End Select

    If m.Chart.SeasonalCycleTypeEnum <> cboCycle.ListIndex Then
        m.Chart.SeasonalCycle = ValOfText(txtCycleNum.Text) & " " & cboCycle.Text
    End If

    For i = 0 To Len(txtCycleNum.Text)
        If Not IsDigit(txtCycleNum, i) Then
            txtCycleNum.Text = Str(m.Chart.SeasonalCycleLen)
        End If
    Next
    
    i = Abs(Int(ValOfText(txtCycleNum.Text)))
    If m.Chart.SeasonalCycleLen <> i Then
        m.Chart.SeasonalCycle = i & " " & cboCycle.Text
    End If
    
    If Len(strBarType) = 0 Then
        m.Chart.GenerateChart eRedo9_ReloadData
    Else
        'chart object automatically reloads data when its bar period is changed
        m.Chart.ChangeBarPeriod strBarType
    End If
    
    cmdSeasonalApply.Enabled = False
    
    tmr.Enabled = True
    
    SetFocusCtl

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdSeasonalApply_Click"

End Sub

Private Sub cmdSell_Click()
On Error GoTo ErrSection:

    Dim eReplayModeSave As eGDReplayMode
    
    eReplayModeSave = m.oGameMode.GameReplayMode(tmrGameMode.Enabled)
    cmdStop_Click
    HandleGameOrder eGDTTEditOrderMode_GameNewOrder, , , 0
    If eReplayModeSave = eGDReplayMode_Play Then cmdPlay_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdSell_Click"
    Resume ErrExit

End Sub

Private Sub cmdSellAsk_Click()
On Error GoTo ErrSection:

    HandleChartOrder True, eClickOrder_SellAsk
    SetFocusCtl

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdSellAsk_Click"

End Sub

Private Sub cmdSellAsk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    cmdSellAsk.ToolTipText = "Sell " & Trim(txtTradeQty) & " at the current Ask price or better"

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdSellAsk_MouseMove"

End Sub

Private Sub cmdSellBid_Click()
On Error GoTo ErrSection:

    HandleChartOrder True, eClickOrder_SellBid
    SetFocusCtl

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdSellBid_Click"

End Sub

Private Sub cmdSellBid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    cmdSellBid.ToolTipText = "Sell " & Trim(txtTradeQty) & " at the current Bid price or better"

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdSellBid_MouseMove"

End Sub

Private Sub cmdSellMarket_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    cmdSellMarket.ToolTipText = "Sell " & Trim(txtTradeQty) & " at market"

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdSellMarket_MouseMove"

End Sub

Private Sub cmdSellMarket_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        HandleBuySellClick cmdSellMarket, Button
    Else
        HandleChartOrder True, eClickOrder_SellMkt
    End If

    SetFocusCtl

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdSellMarket_MouseUp"

End Sub

Private Sub cmdSettings_Click()
On Error GoTo ErrSection:

    TopMost = False
    
    SetFocusCtl
    tmr.Tag = "EditSettings"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdSettings_Click"
    Resume ErrExit
    
End Sub

Private Sub cmdSettings_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode <> vbEnter Then SetFocusCtl    'Peg    -MJM
    KeyPress KeyCode, Shift

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdSettings_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub cmdSettings_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdSettings_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub cmdStop_Click()
On Error GoTo ErrSection:

    tmrGameMode.Enabled = False
    
    If Not m.oGameMode Is Nothing Then
        If m.oGameMode.UserForwardBack = 0 Then
            If m.oGameMode.GameReplayMode(False) <> eGDReplayMode_Sync Then
                If InStr(Me.Caption, "PAUSED") = 0 Then
                    Me.Caption = Me.Caption & " - PAUSED"
                    vseCaption = vseCaption & " - PAUSED"
                End If
            End If
        End If
    End If
    
    Enable cmdPlay, True
    Enable cmdStop, False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdStop_Click"
    Resume ErrExit
    
End Sub

Private Sub cmdBuy_Click()
On Error GoTo ErrSection:

    Dim eReplayModeSave As eGDReplayMode
    
    eReplayModeSave = m.oGameMode.GameReplayMode(tmrGameMode.Enabled)
    cmdStop_Click
    HandleGameOrder eGDTTEditOrderMode_GameNewOrder, , , 1
    If eReplayModeSave = eGDReplayMode_Play Then cmdPlay_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdBuy_Click"
    Resume ErrExit
End Sub

Private Sub cmdTicket_Click()
On Error GoTo ErrSection:

    Dim strMsg$
    
    strMsg = WizardGridLegInfo(Me, True)
    
    If Len(strMsg) > 0 Then
        StatusMsg "Creating order ticket..."
        strMsg = Chr(27) & g.Broker.AccountNumberForID(m.Chart.TradeAccountID) & ";" & Str(m.Chart.geChartObj) & vbTab & strMsg       '4979, 5029 (added chart "ID")
    
        AllowSetForegroundWindow -1&            '4990
        
        If StartOptNav(m.Chart.Symbol, True) Then
            CreateTicketInOptNav strMsg
        End If
    Else
        strMsg = "There are no strategy legs to create ticket with."
        InfBox strMsg, "I", , "Create Ticket"
    End If
    
    StatusMsg ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdTicket_Click"

End Sub

Private Sub cmdTSOEdit_Click()
    TSOGrpFavoritesEdit Me, True
End Sub

Private Sub fgChartFlex_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection

    Dim strErr$
    Dim bLocked As Boolean

    If Index = eFlexGridIdx_PfpInd Then
        With fgChartFlex(eFlexGridIdx_PfpInd)
            If Row >= .FixedRows And Row < .Rows Then
                If .ComboIndex > 0 And .ComboIndex < .ComboCount Then
                    .Cell(flexcpData, Row, 0) = "#" & .TextMatrix(Row, 0) & ";" & .ComboItem(.ComboIndex)
                    .TextMatrix(Row, 0) = .Cell(flexcpTextDisplay, Row, 0)
                    SaveIndGridPFP Chart, fgChartFlex(eFlexGridIdx_PfpInd)
                End If
            End If
        End With
        chkApplyFilter.Tag = "NoPrompt"
        chkApplyFilter_Click
        chkApplyFilter.Tag = ""
    ElseIf Index = eFlexGridIdx_PfpHits Then
        If Col = eColsPFP_Use Then
            PfpReset ePfpReset_PpfAnnot
            strErr = "Internal error. Please close then reopen Pattern for Profit and try again."
            
            If m.oPatternProfit Is Nothing Then
                InfBox strErr, "E", , "Patterns for Profit"
            ElseIf m.oPatternProfit.RebuildComposite(fgChartFlex(eFlexGridIdx_PfpHits)) Then
                bLocked = LockWindowUpdate(Me.hWnd)
                m.Chart.GenerateChart eRedo1_Scrolled
                If bLocked Then LockWindowUpdate (0)
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".fgChartFlex_AfterEdit"

End Sub

Private Sub fgChartFlex_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim i&

    If Index = eFlexGridIdx_OrdWizard Then
        Select Case Col
            Case 0
                With fgChartFlex(eFlexGridIdx_OrdWizard)
                    If Row >= .FixedRows And Row < .Rows Then
                        If Len(.TextMatrix(Row, 1)) > 0 Then
                            .Redraw = flexRDNone
                            .Cell(flexcpPicture, Row, Col) = Nothing
                            .Cell(flexcpBackColor, Row, 1) = .Cell(flexcpBackColor, Row, 0)
                            .TextMatrix(Row, 1) = ""
                            .TextMatrix(Row, 2) = ""
                            .TextMatrix(Row, 3) = ""
                            .TextMatrix(Row, 4) = ""
                            .Redraw = flexRDBuffered
                        End If
                    End If
                End With
                cboRiskGraphType_Click
                Cancel = True
            Case 1
                Cancel = True
            Case 2
                If Len(fgChartFlex(eFlexGridIdx_OrdWizard).TextMatrix(Row, 1)) = 0 Then Cancel = True
        End Select
    ElseIf Index = eFlexGridIdx_PfpHits Then
        If Col <> eColsPFP_Use Then Cancel = True   '5887 - code moved to AfterEdit event
    End If

End Sub

Private Sub fgChartFlex_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim dDate#, i&, j&
    
    If Index = eFlexGridIdx_PfpInd Then
        With fgChartFlex(eFlexGridIdx_PfpInd)
            If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
                m.bDropdownPFP = True
                .Row = .MouseRow
                .Col = .FixedCols
                
                PfpReset ePfpReset_GridPfp
                PfpReset ePfpReset_PpfAnnotInd
                
                .EditCell
                SendKeys "{F4}"
            End If
        End With
    ElseIf Index = eFlexGridIdx_PfpHits Then
        With fgChartFlex(eFlexGridIdx_PfpHits)
            If .MouseCol <> eColsPFP_Use Then
                If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
                    dDate = .TextMatrix(.MouseRow, eColsPFP_DateDouble)
                    If Not m.oPatternProfit Is Nothing And Not m.Chart Is Nothing Then
                        m.oPatternProfit.MatchAnnotShow m.Chart, Me, dDate
                        i = m.Chart.Bars.FindDateTime(dDate)
                        j = m.oPatternProfit.PatternSelLength / 2
                        If i - j >= 0 Then dDate = m.Chart.Bars(eBARS_DateTime, i - j)
                    End If
                    CenterTheDate dDate
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".fgChartFlex_Click"

End Sub

Private Sub fgChartFlex_ComboCloseUp(Index As Integer, ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)

    If Index = eFlexGridIdx_PfpInd Then FinishEdit = True

End Sub

Private Sub fgChartFlex_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next

    If Index = eFlexGridIdx_PfpInd Then
        If m.bDropdownPFP Then
            m.bDropdownPFP = False
        ElseIf KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And KeyCode <> vbKeyReturn Then
            KeyCode = 0
        End If
    End If

End Sub

Private Sub fgChartFlex_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
On Error Resume Next

    If Index = eFlexGridIdx_PfpInd Then
        If m.bDropdownPFP Then
            m.bDropdownPFP = False
        ElseIf KeyAscii <> vbKeyUp And KeyAscii <> vbKeyDown And KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub fgChartFlex_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim dDate#, i&, j&

    If Index = eFlexGridIdx_Seasonal Then
        HandleSeasonalInput eSeasonalCtrl_SeasonalGrid
    ElseIf Index = eFlexGridIdx_PfpHits Then
        With fgChartFlex(eFlexGridIdx_PfpHits)
            If .Row >= .FixedRows And .Row < .Rows Then
                dDate = .TextMatrix(.Row, eColsPFP_DateDouble)
                If Not m.oPatternProfit Is Nothing And Not m.Chart Is Nothing Then
                    m.oPatternProfit.MatchAnnotShow m.Chart, Me, dDate
                    i = m.Chart.Bars.FindDateTime(dDate)
                    j = m.oPatternProfit.PatternSelLength / 2
                    If i - j >= 0 Then dDate = m.Chart.Bars(eBARS_DateTime, i - j)
                End If
                CenterTheDate dDate
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".fgChartFlex_KeyUp"

End Sub

Private Sub fgChartFlex_KeyUpEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next
    
    If Index = eFlexGridIdx_PfpInd Then
        If m.bDropdownPFP Then
            m.bDropdownPFP = False
        ElseIf KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And KeyCode <> vbKeyReturn Then
            KeyCode = 0
        End If
    End If

End Sub

Private Sub fgChartFlex_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Index
        Case eFlexGridIdx_AcctBar
            If Button = vbRightButton Then frmChartOrdBar.ShowMe Me
        Case eFlexGridIdx_Seasonal
            HandleSeasonalInput eSeasonalCtrl_SeasonalGrid
    End Select

End Sub

Private Sub Form_Activate()
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean
    
    If m.Chart Is Nothing Then Exit Sub
    
    ' #6922 10/17/2013: needed to move this section up (prior to the check for g.bSkipSetChartFocus)
    If Not m.bWindowLinkInitDone Then
        'If Not m.bGameMode Then        -5048
            ' subclass the window proc to intercept some of the messages
            m.WindowLink.Init Me
        'End If
        m.bWindowLinkInitDone = True
    End If
    
    ' TLB 8/6/2013: this check appears to be needed to help keep proper Zorder when loading a chart page:
    If g.bSkipSetChartFocus Then Exit Sub

    If Not Me Is ActiveChart Then
        If Me.WindowState <> vbMinimized And tmr.Tag <> "MINIMIZE_NOW" Then
            ActiveChartFormSet Me
        End If
    End If
    
    LogEvent "Activate"
    
'    If Not m.bAllowDetach Or m.eDetachStatus <> eDetached Then     07-01-2013: original; leave awhile then remove if all ok
    If m.eDetachStatus <> eDetached Then
        Me.pbTbBack(0).Visible = False
        Me.imgTbBack(0).Visible = False
    End If
    'make sure chart tips don't show
    vseTipY.Top = -1000 - vseTipY.Height
    vseTipX.Top = -1000 - vseTipX.Height
    vseTipChart.Top = -1000 - vseTipChart.Height        '4878

    m.Chart.SyncToolbar
        
    'LockWindowUpdate 0
    
    'Need to move focus to the chart control, otherwise for
    'some strange reason things don't work right (e.g. focus
    'moving back to the form when click on the chart).
    If Screen.ActiveForm Is Me Then
        'DoEvents '(no, if do this in here then scrollbar acts goofy)
        g.bDirtyChartPage = True
        
        ' but don't do this if maximized (will cause problems on XP and 2000)
        If Me.WindowState = 0 And Not Me.IsInGameMode Then
            If Not bAlreadyDone Then
                SetFocusCtl
            End If
        ElseIf Me.MDIChild And Me.WindowState = vbMaximized And m.Chart.ShowTrades <> 0 Then
            vseOrderBar.Visible = True      '4900
            Form_Resize
        End If
                       
        'sync chart on/off form
'        frmChartOnOff.ShowData
                        
        ' reset ChartCfg if visible and set to different chart
        'If DockState(frmChartCfg) <> eHidden Then
        '    If Not frmChartCfg.Chart Is m.Chart Then
        '        frmChartCfg.ShowMe m.Chart
        '    End If
        'End If
        UnloadEditors
    End If
    
    If m.MouseDown.dTickTime > gdTickCount - 250# Then
        frmChartData.ShowData m.MouseDown.nX
        frmPlanetData.ShowData m.MouseDown.dDate, m.Chart.Bars
    Else
        frmChartData.ShowData -1
        frmPlanetData.ShowData -1
    End If
    
'JM(11-02-2009) - this is now done in mMain; commented out to fix issue 5440
'    frmMain.SetWindowLink Me
    
    If frmImageServer.Active Then SetImgSrvSearcher
    
    m.Chart.EnablePerformanceButton
        
    tmr.Enabled = ChartTimers
        
    ' this needs to be the LAST thing in the Activate event
    If Me.WindowState = vbMaximized Then
        'CheckSystemMenu
        
        'this code ensures system menu is enabled and think it is more efficient swapping out charts etc.
        'also fixes problem of chart going temporarily blank when charts are getting swapped out
        GetSystemMenu Me.hWnd, 1
    End If

ErrExit:
    bAlreadyDone = True
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".Form_Activate"
    Resume ErrExit
    
End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    
    LogEvent "Deactivate"
    
    If Me.MDIChild Then TextIncDecUnregisterForm Me
           
    ' if an annotation has been started, then reset tools when changing chart windows
    If m.nActiveAnnotIdx = 0 Then
        ClearAnnotFlags True, False
    Else
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Annot Is Nothing Then
            ClearAnnotFlags True, True
        ElseIf Annot.AlertAddEditInprog Then
            GoTo ErrExit        '5632
        Else
            ClearAnnotFlags True, True
        End If
    End If
    
    'pause game mode
    If m.bGameMode Then cmdStop_Click
    
    'Disable tmr '(now need it for real-time)

    frmMain.tbToolbar.Tools("ID_Performance").Enabled = False
            

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".Form_Deactivate"
    Resume ErrExit
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    Else
        KeyPress KeyCode, Shift
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".Form_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If m.bGameMode And Chr(KeyAscii) = " " Then
        If Not m.oGameMode Is Nothing Then
            If m.oGameMode.GameReplayMode(tmrGameMode.Enabled) <> eGDReplayMode_Sync Then
                KeyAscii = 0
                If m.oGameMode.GameReplayMode(tmrGameMode.Enabled) = eGDReplayMode_Play Then
                    cmdStop_Click
                Else
                    cmdPlay_Click
                End If
            End If
        End If
    End If
    
    If KeyAscii <> 0 Then
        KeyPress KeyAscii
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".Form_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim i&, h&, w&, l&, t&

    g.Styler.StyleForm Me
    
    If m.Chart Is Nothing Then
        Me.Tag = "BOGUS"
        DebugLog "Bogus frmChart Load"
        If IsIDE Then
            StatusMsg "Bogus frmChart Load"
        End If
        Exit Sub     '5202
    End If
       
    'RH populate combos ' Hexagora does not allow pre-populating Comboboxes in design mode. They must be setup at runtime
    With cboOrderType
        .AddItem "Auto"
        .AddItem "Limit"
        .AddItem "Stop"
    End With
    
    With cboExchanges
        .AddItem "Auto"
        .AddItem "Limit"
        .AddItem "Stop"
    End With
       
       
    vseScrollSep.BackColor = Me.BackColor ' &H8000000F
    FixFormControls Me
    
    pbTbBack(0).Visible = False
    imgTbBack(0).Visible = False
    pbTbBackDraw(0).Visible = False
    imgTbBackDraw(0).Visible = False
    pbNotUsed.Visible = False
    cboBarPeriod.Visible = False
    cboBarPeriod.Text = ""
    
    'JM 05-19-2009: From Tim - requires RTG
    m.bTradeContinuous = HasModule("RTG")   'FileExist("TradeContinuous.flg")

    'JM 05-19-2009: From Tim - requires gold
    'JM 07-01-2013: From Glen - everyone gets detached chart
'    If HasLevel(eTN4_Gold, False) Or ExtremeCharts >= 1 Then
'        m.bAllowDetach = True
'    '10-06-2011: per Pete - also allow if TSU and Standard
'    ElseIf HasLevel(eTN3_Standard, False) And HasModule("TSU") Then
'        m.bAllowDetach = True
'    Else
'        m.bAllowDetach = False
'    End If
    
    'JM 05-19-2009: From Tim - requires gold
    m.bAllowOptWizard = HasModule("OPTNAV") 'FileExist("OptWizard.flg")
    m.bAllowRiskGraph = HasModule("OPTNAV")

    'JM 09-21-2010: for Schiff Andrews Fork
    m.bSchiffFork = HasModule("ADVFIB") 'FileExist("SchiffFork.flg")
    
    GameSpeed

    pbChart.BorderStyle = 0
    pbChart.AutoRedraw = True
    
    Set m.aTabs = New cGdArray
    
    'mnuLockValues.Visible = False 'obsolete: "Only display Last Bar's values"
    
    mnuPopUp.Visible = False
    If ExtremeCharts = 1 And Not HasModule("IT") Then
        mnuBarPeriod(0).Visible = False
        mnuBarPeriod(1).Visible = False
        mnuBarPeriod(2).Visible = False
        mnuBarPeriod(3).Visible = False
        mnuBarPeriod(4).Visible = False
        mnuBarPeriod(5).Visible = False
        mnuBarPeriod(6).Visible = False
    End If

    mnuAnnotEdit.Visible = False
    mnuOrderAction.Visible = False
    mnuContractAction.Visible = False
    Me.Icon = Picture16("kBlank")
    If UCase(g.strTitle) <> "EXTREME CHARTS" Then
        vseCaption.Picture = Picture16(ToolbarIcon("ID_About"))
    End If
    hsb.Visible = True
    vseMouse.Caption = ""
    vseDay.Caption = ""
    vseCaption.Top = 0          '5060
    
    'order bar mode (standard or OptNav wizard)
    fraOrderBarMode.Move 0, 0
    
    'broker disconnect
    fraBrokerDisconnect.Visible = False
'    fraBrokerDisconnect.Enabled = False
    'RH commented out fraBrokerDisconnect.BorderStyle = 0
    
    'order wizard controls
    fraWizardPrompt.Visible = False
    
    lblRiskGraph.Visible = m.bAllowRiskGraph
    cboRiskGraphType.Visible = m.bAllowRiskGraph
    cmdRefreshGraph.Visible = m.bAllowRiskGraph
    
    cboRiskGraphType.Enabled = False
    cmdRefreshGraph.Enabled = False
    cmdTicket.Enabled = False
    
    'order bar
    vseOrderBar.Visible = False
    fraExitFavorites.Visible = False
    'RH commented out fraExitFavorites.BorderStyle = 0
    'account bar
    fgChartFlex(eFlexGridIdx_AcctBar).HighLight = flexHighlightNever
    m.aABarColHeader.Size = 0
    GridBarHeader fgChartFlex(eFlexGridIdx_AcctBar), m.aABarColHeader, i, m.Chart.AcctBarCols
    ResetGridBar fgChartFlex(eFlexGridIdx_AcctBar), m.aABarColHeader, i
    fgChartFlex(eFlexGridIdx_AcctBar).ScrollBars = flexScrollBarNone
    'game mode labels
    lblPosition = "None"
    lblProfit = ""
    lblOpenEquity = ""
    cmdBailout.BackColor = cmdReverse.BackColor
    fraFrontMonth.Move fraOrderBtns.Left, fraOrderBtns.Top + 250
    
    If Not m.bTradeContinuous Then
        lblFrontMonth.Move lblOpenOrderPos.Left, lblOpenOrderPos.Top + 50
        cmdRollNow.Caption = "Active" & vbCrLf & "Contract"
        cmdContracts.Visible = False
        lblOpenOrderPos.Visible = False
        lblFrontMonth.Visible = True
    End If
            
    cmdBailout.BackColor = cmdReverse.BackColor
    cmdBailout.Enabled = False
    
    vseBuyChart.Caption = "Click chart " & vbCrLf & "to BUY  "
    vseSellChart.Caption = "Click chart " & vbCrLf & "to SELL  "
    
    'pattern for profit
    fraPatternProfit.Move 0, 0
    'RH commented out fraPatternProfit.BorderStyle = 0
    
    InitSeasonalControls Me              'do this here so won't see the grid move on initial show
    
    ' get chart loaded with defaults
    Chart.RedoMode = eRedo7_ReloadRT     'need this as fix for issue 6389

    vseTipY.Caption = ""
    vseTipX.Caption = ""
    vseTipChart.Caption = ""
    vseTipY.Top = -1000 - vseTipY.Height
    vseTipX.Top = -1000 - vseTipX.Height
    vseTipChart.Top = -1000 - vseTipChart.Height        '4878
    ChartTips eTiptype_None
    
    Screen.MousePointer = 11

    'set chart picture box properties
    pbChart.BackColor = g.ChartGlobals.nChartBackColor
    pbChart.Visible = True
    
    If AllowMIT() Then
        cboOrderType.AddItem "MIT"
    ElseIf m.Chart.PseudoOrderType > 2 Then
        m.Chart.PseudoOrderType = 0
    End If
    
    If m.Chart.PseudoOrderType >= 0 And m.Chart.PseudoOrderType <= 3 Then
        cboOrderType.ListIndex = m.Chart.PseudoOrderType
    Else
        cboOrderType.ListIndex = 0
    End If
    
    ' 10/05/2010 DAJ: For now, just hide the Rithmic images until we know what we are doing...
    fraRithmic.Visible = False
    
    Set m.Quantity = New cPriceEditor
    
    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    If IsIDE Then
        MsgBox "Check the Stack Window to see why the frmChart.Load error happened.", , "frmChart.Load ERROR"
        i = i
    Else
        RaiseError Me.Name & ".Form_Load"
    End If
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Dim i&, strFiles$, strText$
    Dim Annot As cAnnotation

    TopMost = False
    tmrGameMode.Enabled = False
                        
    If UnloadMode = vbFormControlMenu Then  'user chose the Close command from the system menu
        Screen.MousePointer = vbHourglass
        g.bDirtyChartPage = True            '6401
                
        ' save template (so all annots will get saved)
        If Not m.Chart Is Nothing Then
            m.Chart.TemplateSave
            m.Chart.RemoveAlerts True
            m.Chart.RemoveAnnots False      ', , , True     -bRemoveAll=true
        End If
        
        ' check if any drawing tools exist for this chart window only (for any symbol)
        'except don't remove annots for instant replay - aardvark 2950
        If m.bGameMode Then
            strFiles = g.ChartGlobals.strCPCRoot & "\Charts\Replay.CHT"
            KillFile strFiles
        Else
            strFiles = g.ChartGlobals.strCPCRoot & "\Charts\" & m.Chart.Template & "^*.ANO"
            If FileExist(strFiles) Then
                Screen.MousePointer = 0
                If UCase(Me.Tag) = "MOVED" Then
                    'chart was moved to another page, don't prompt for deletion of drawing tools
                    KillFile strFiles
                    strFiles = g.ChartGlobals.strCPCRoot & "\Charts\" & m.Chart.Template & ".CHT"
                    KillFile strFiles
                Else
                    strText = "Drawing objects drawn in this chart window| for any symbol will be permanently lost| (except those set as 'Show in all charts')."
                    If InfBox(strText, "i", "+OK|-Cancel", "Please Note ...") = "C" Then
                        Cancel = True
                        m.Chart.GenerateChart eRedo2_ReloadAnnots       '4643
                    Else
                        KillFile strFiles
                        strFiles = g.ChartGlobals.strCPCRoot & "\Charts\" & m.Chart.Template & ".CHT"
                        KillFile strFiles
                    End If
                End If
            Else
                'there are no annotations, delete the CHT file
                strFiles = g.ChartGlobals.strCPCRoot & "\Charts\" & m.Chart.Template & ".CHT"
                KillFile strFiles
            End If
        End If
    ElseIf Not m.Chart Is Nothing Then
        m.Chart.RemoveAlerts True         '4462
    End If
    
    If m.eDetachStatus <> eAttachInProg And m.eDetachStatus <> eDetachInProg Then
        If m.bGameMode Then
            If Not m.oGameMode Is Nothing Then
                If m.oGameMode.GameReplayMode(tmrGameMode.Enabled) <> eGDReplayMode_Sync Then
                    m.oGameMode.SaveGameResult
                End If
            End If
        End If
    End If
    
    If Cancel = 0 Then
        If Not g.bUnloading Then m.WindowLink.Unhook    '5551
        tmr.Enabled = False
        tmr.Tag = "UNLOADING"
               
        If Not g.bUnloading Then
            If m.Chart Is Nothing Then
                SetNextChartActive Me, ""
            Else
                SetNextChartActive Me, m.Chart.Symbol
            End If
        End If
        
        ClearChartPointers Me
        
    End If
    Screen.MousePointer = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".Form_QueryUnload"
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Static iPrevWindowState%
    
    Dim strText$, l&, t&, w&, h&, i&
    Dim iTabSpace&, iToolbarHeight&
    Dim bLockedWindow As Boolean
    Dim bTbDrawVisible As Boolean
    Dim iHsbHeight&

'vbNormal = 0
'vbMinimized = 1
'vbMaximized = 2
    
    If g.bUnloading Or g.bLoadingChartPage Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "UNLOADING" Then Exit Sub
    If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then Exit Sub
        
    bTbDrawVisible = frmMain.ToolbarVisible(kTbDraw)
    
    ' minimum/maximum size
    If Me.WindowState = vbNormal Then
        If LimitFormSize(Me, kMinChartWidth, kMinChartHeight) Then Exit Sub
        TopMost = False
    End If
    
    LogEvent "Resize"
        
    ' to help avoid some flicker
    If Not IsAtLeastVista Then
        If Not g.bStarting Then
            If m.eDetachStatus = eDetached Then
                bLockedWindow = LockWindowUpdate(GetDesktopWindow())
            ElseIf m.eDetachStatus = eNotDetached Then
                bLockedWindow = LockWindowUpdate(Me.hWnd)
            End If
        End If
    End If
            
    iToolbarHeight = ToolbarHeight(pbTbBack)
    ' need to do this in since the "Minimized" resize event isn't always
    ' triggered for other charts when active chart goes between Max and Normal
    If Me.WindowState <> iPrevWindowState Then
        If m.eDetachStatus = eNotDetached Then
            If Not ActiveChart Is Nothing Then
                If ActiveChart.DetachStatus = eDetached And iPrevWindowState = vbMaximized Then     '4883
                    'user clicked the "Restore" button in upper right of system menu of main app window while detached chart in focus
                    ActiveChart.SkipFocusFix = True
                    SendMessage Me.hWnd, WM_NCACTIVATE, 1, 0
                    SendMessage Me.hWnd, WM_MOUSEACTIVATE, 1, 0
                End If
            End If
            For i = 0 To Forms.Count - 1
                If IsFrmChartMDI(Forms(i)) And Not Forms(i) Is Me Then
                    If tmr.Tag <> "UNLOAD_NOW" And tmr.Tag <> "UNLOADING" Then
                        If Forms(i).WindowState = vbMinimized Then
                            FormResize Forms(i)
                        End If
                    End If
                End If
            Next
        ElseIf iPrevWindowState = vbMinimized Then
            FormResize Me           '6106 - see note about "miminimized" event above (same reason)
        End If
    End If
    
    If m.bGameMode And Not m.oGameMode Is Nothing Then
        If m.oGameMode.GameReplayMode(tmrGameMode.Enabled) <> eGDReplayMode_Sync Then
            'instant replay with replay data on all charts ON, but this is not the Replay chart so don't show Replay buttons
            vseReplay.Visible = True
        End If
    Else
        vseReplay.Visible = False
    End If
    
    If Not m.Chart Is Nothing Then
        If m.Chart.TypeOfChart = eTypeChart_Seasonal Then vseSeasonal.Visible = True
    End If
        
    ' TLB 5/12/2014: option to hide scrollbars
    If g.ChartGlobals.bHideScrollbars Then
        hsb.Visible = False
        iHsbHeight = 0
    Else
        hsb.Visible = True
        iHsbHeight = hsb.Height
    End If
        
    ' fake caption (when maximized)
    With vseCaption
        If Me.WindowState = vbMaximized And Me.MDIChild Then
            If m.eDetachStatus = eNotDetached Then
                vsTab.Move -30, -30, Me.ScaleWidth + 60, Me.ScaleHeight + 30
                ' when maximized, show fake caption (if not already)
                If vseCaption.Top <> 0 Then 'Or iPrevWindowState = vbMaximized Then
                    Me.Icon = Picture16("kBlank")
                    Me.Caption = ""
                    .Move 0, 0, Screen.Width * 3, .Height
                    .Visible = True
                    'SetChartTabs
                    'Me.Refresh
                    '.Refresh
                End If
                If iPrevWindowState <> vbMaximized Then
                    SetChartTabs
                End If
                If m.bGameMode Then
                    vseCaption = Space(5) & "Instant Replay"
                    If Not m.oGameMode Is Nothing Then
                        If m.oGameMode.GameReplayMode(tmrGameMode.Enabled) <> eGDReplayMode_Sync Then
                            With vseReplay
                                .Move 0, t - 10, Me.ScaleWidth, fraReplayProfit.Height + 200        '5048
                                fraReplayProfit.Move fraReplayProfit.Left, fraReplayProfit.Top, vseReplay.Width - fraReplayProfit.Left - 30
                                vsTab.Move 0, .Top + .Height, Me.ScaleWidth + 60, Me.ScaleHeight - .Height
                            End With
                        End If
                    End If
                End If
                If vsTab.Enabled Then
                    iTabSpace = 330
                End If
            ElseIf m.eDetachStatus = eDetached Then
                If m.Chart.ShowToolbar Then
                    ToolbarResize2 Me, pbTbBack, imgTbBack, m.aTbButtons, m.bToolbarWrap
                    ToolbarResize2 Me, pbTbBackDraw, imgTbBackDraw, m.aTbButtonsDraw, False
                End If
                Me.cboBarPeriod.SelLength = 0
            End If
        'ElseIf vseCaption.Top >= 0 Then
        ElseIf vseCaption.Top >= -1 Then ' TLB 6/20/2014: now need this condition in case still set to -1 from SetChartTabs
            ' when not maximized, hide the fake caption and chart tabs (this code executes when going from max to normal)
            vsTab.Enabled = False
            vsTab.Visible = False
            Me.Caption = Trim(.Caption)
            .Visible = False
            If UCase(g.strTitle) <> "EXTREME CHARTS" Then
                Me.Icon = Picture16(ToolbarIcon("ID_About"))
            End If
            .Top = -.Height
            If vseReplay.Visible Then
                vseReplay.Move 0, -275, Me.ScaleWidth, fraReplayProfit.Height + 200
                fraReplayProfit.Move fraReplayProfit.Left, fraReplayProfit.Top, vseReplay.Width - fraReplayProfit.Left - 30
            End If
        ElseIf m.bGameMode Then
            vsTab.Enabled = False   '(this code executes when chart is normal to start with)
            vsTab.Visible = False
            Me.Caption = Trim(.Caption)
            If vseReplay.Visible Then
                vseReplay.Move 0, -275, Me.ScaleWidth, fraReplayProfit.Height + 200
                fraReplayProfit.Move fraReplayProfit.Left, fraReplayProfit.Top, vseReplay.Width - fraReplayProfit.Left - 30
            End If
        ElseIf vseOrderBar.Visible Then
            vsTab.Enabled = False
            vsTab.Visible = False
            Me.Caption = Trim(.Caption)
        End If
        If m.eDetachStatus = eDetached Then
            If m.Chart.ShowToolbar Then
                ToolbarResize2 Me, pbTbBack, imgTbBack, m.aTbButtons, m.bToolbarWrap
                ToolbarResize2 Me, pbTbBackDraw, imgTbBackDraw, m.aTbButtonsDraw, False
            End If
            Me.cboBarPeriod.SelLength = 0
            If vseReplay.Visible Then
                t = -275 + iToolbarHeight
                'If pbTbBack(0).Visible Then t = -275 + pbTbBack(0).Height
                w = Me.ScaleWidth
                If pbTbBackDraw(0).Visible Then
                    If pbTbBackDraw(0).align = vbAlignTop Then
                        t = t + pbTbBackDraw(0).Height
                    Else
                        w = w - pbTbBackDraw(0).Width
                    End If
                End If
                vseReplay.Move 0, t, w, fraReplayProfit.Height + 200
                fraReplayProfit.Move fraReplayProfit.Left, fraReplayProfit.Top, vseReplay.Width - fraReplayProfit.Left - 30
                t = 0   'reset
                w = 0
            End If
        End If
    End With
    
    If m.Chart.ShowAccountBar And m.oGameMode Is Nothing Then
        With fgChartFlex(eFlexGridIdx_AcctBar)
            If m.eDetachStatus = eDetached Then
                t = iToolbarHeight
                If pbTbBackDraw(0).Visible And pbTbBackDraw(0).align = vbAlignTop Then
                    t = t + pbTbBackDraw(0).Height
                End If
            Else
                t = vseCaption.Top + vseCaption.Height
            End If
            w = Me.ScaleWidth / .Cols
            .Move -30, t, Me.ScaleWidth + 100
            .ColWidthMin = w
            .ScrollBars = flexScrollBarNone
            If Not .Visible Then .Visible = True
        End With
    Else
        fgChartFlex(eFlexGridIdx_AcctBar).Visible = False
    End If
    
    If Not g.bStarting Then
        iPrevWindowState = Me.WindowState
    End If
    
    ''If Me.Top < 0 Then Me.Top = 0
    If Me.WindowState = vbMinimized Then
        If bLockedWindow Then LockWindowUpdate 0
        Exit Sub
    End If
    
    If iTabSpace = 0 Then vsTab.Visible = False
    
    If vseOrderBar.Visible Then
        If m.eOrdBarMode = eOrdBarMode_PFP Then
            w = kPFPBarWidth
        Else
            w = kOrdBarWidth
        End If
        If Me.WindowState = vbMaximized Then
            If fgChartFlex(eFlexGridIdx_AcctBar).Visible Then
                vseOrderBar.Move Me.ScaleWidth - w, fgChartFlex(eFlexGridIdx_AcctBar).Top + fgChartFlex(eFlexGridIdx_AcctBar).Height, w, Me.ScaleHeight
            ElseIf m.eDetachStatus = eDetached And Me.pbTbBack(0).Visible Then
                'vseOrderBar.Move Me.ScaleWidth - w, vseCaption.Height + Me.pbTbBack(0).Height, w, Me.ScaleHeight - Me.pbTbBack(0).Height    '5015
                vseOrderBar.Move Me.ScaleWidth - w, vseCaption.Height + iToolbarHeight, w, Me.ScaleHeight - iToolbarHeight    '5015
            Else
                vseOrderBar.Move Me.ScaleWidth - w, vseCaption.Height, w, Me.ScaleHeight
            End If
            vsTab.Width = Me.ScaleWidth - vseOrderBar.Width
        ElseIf fgChartFlex(eFlexGridIdx_AcctBar).Visible Then
            vseOrderBar.Move Me.ScaleWidth - w, fgChartFlex(eFlexGridIdx_AcctBar).Top + fgChartFlex(eFlexGridIdx_AcctBar).Height, w, Me.ScaleHeight
        ElseIf m.eDetachStatus = eDetached And pbTbBack(0).Visible Then
            If pbTbBackDraw(0).Visible And pbTbBackDraw(0).align = vbAlignTop Then
                vseOrderBar.Move Me.ScaleWidth - w, pbTbBackDraw(0).Top + pbTbBackDraw(0).Height, w, Me.ScaleHeight
            Else
                vseOrderBar.Move Me.ScaleWidth - w, iToolbarHeight, w, Me.ScaleHeight     '4877 (order bar cut-off at top)
            End If
        Else
            vseOrderBar.Move Me.ScaleWidth - w, 0, w, Me.ScaleHeight
        End If
        w = Me.ScaleWidth - w
        If pbTbBackDraw(0).Visible And pbTbBackDraw(0).align = vbAlignRight Then
            vseOrderBar.Left = vseOrderBar.Left - pbTbBackDraw(0).Width
        End If
    ElseIf vseSeasonal.Visible Then
        w = vseSeasonal.Width
        If m.eDetachStatus = eDetached Then
            vseSeasonal.Move Me.ScaleWidth - w, 90, w, Me.ScaleHeight
        ElseIf Me.WindowState = vbMaximized Then
            vseSeasonal.Move Me.ScaleWidth - w, vseCaption.Top + vseCaption.Height + 60, w, Me.ScaleHeight
        Else
            vseSeasonal.Move Me.ScaleWidth - w, 75, w, Me.ScaleHeight
        End If
        w = Me.ScaleWidth - vseSeasonal.Width
        
        If pbTbBackDraw(0).Visible And pbTbBackDraw(0).align = vbAlignRight Then
            vseSeasonal.Left = vseSeasonal.Left - pbTbBackDraw(0).Width
        End If
        
        fgChartFlex(eFlexGridIdx_Seasonal).Top = fraTrendCycle.Top + fraTrendCycle.Height + 150
        fgChartFlex(eFlexGridIdx_Seasonal).Left = vseSeasonal.Left + 45
        fgChartFlex(eFlexGridIdx_Seasonal).Width = fraTrendCycle.Width - 30
        
        If Me.WindowState = vbMaximized Then
            fgChartFlex(eFlexGridIdx_Seasonal).Height = vseSeasonal.Height - (vseSeasonal.Top + fgChartFlex(eFlexGridIdx_Seasonal).Top) - iHsbHeight
            fgChartFlex(eFlexGridIdx_Seasonal).Top = fgChartFlex(eFlexGridIdx_Seasonal).Top + iHsbHeight
        Else
            fgChartFlex(eFlexGridIdx_Seasonal).Height = vseSeasonal.Height - (vseSeasonal.Top + fgChartFlex(eFlexGridIdx_Seasonal).Top) - iHsbHeight / 4
        End If
    Else
        w = Me.ScaleWidth '- l '- 30
    End If
    
    With vseMouse
        If vseReplay.Visible Then
            If vseCaption.Visible Then
                t = vsTab.Top
            Else
                t = vseReplay.Top + vseReplay.Height
            End If
        ElseIf fgChartFlex(eFlexGridIdx_AcctBar).Visible Then
            If m.eDetachStatus = eDetached Then
                'If pbTbBack(0).Visible Then t = pbTbBack(0).Top + pbTbBack(0).Height
                If pbTbBack(0).Visible Then t = iToolbarHeight
                If pbTbBackDraw(0).Visible And pbTbBackDraw(0).align = vbAlignTop Then
                    t = t + pbTbBackDraw(0).Height
                End If
            End If
            t = fgChartFlex(eFlexGridIdx_AcctBar).Top + fgChartFlex(eFlexGridIdx_AcctBar).Height
        ElseIf m.eDetachStatus = eDetached Then
            t = 0
            If pbTbBack(0).Visible Then t = iToolbarHeight
            If pbTbBackDraw(0).Visible And pbTbBackDraw(0).align = vbAlignTop Then
                t = t + pbTbBackDraw(0).Height
            End If
        Else
            t = vseCaption.Top + vseCaption.Height  '+ 30
        End If
        h = 200
        l = 480 ' 240 'cmdUnzoom.Left + cmdUnzoom.Width + 30
        If pbTbBackDraw(0).Visible And pbTbBackDraw(0).align = vbAlignRight Then
            .Move l, t, w - l - pbTbBackDraw(0).Width, h
            vseDay.Move 0, t, l, h
        ElseIf pbTbBackDraw(0).Visible And pbTbBackDraw(0).align = vbAlignLeft Then
            .Move l, t, w - l, h
            vseDay.Move 0 + pbTbBackDraw(0).Width, t, l, h
        Else
            .Move l, t, w - l, h
            vseDay.Move 0, t, l, h
        End If
        .ZOrder
        vseDay.ZOrder
    End With

    ' chart location
    't = vseCaption.Top + vseCaption.Height + 180 ' 150 ' 180
    t = vseMouse.Top + vseMouse.Height
    h = Me.ScaleHeight - t - iHsbHeight - iTabSpace + 30
    If pbTbBackDraw(0).Visible And (pbTbBackDraw(0).align = vbAlignRight Or pbTbBackDraw(0).align = vbAlignLeft) Then
        w = w - pbTbBackDraw(0).Width
    End If
    
    If pbTbBackDraw(0).Visible And pbTbBackDraw(0).align = vbAlignLeft Then
        pbChart.Move pbTbBackDraw(0).Width, t, w, h
    Else
        pbChart.Move 0, t, w, h
    End If
    
    vseInvalid.Move pbChart.Left, pbChart.Top, pbChart.Width, pbChart.Height
    
    With cmdSettings
        l = Me.ScaleWidth - .Width
        t = Me.ScaleHeight - .Height
        '.Move l, t
        '.ZOrder
    End With
    
    ' scrollbar
    With hsb
        l = 0
        t = Me.ScaleHeight - .Height - iTabSpace
        h = .Height
        '.Move l, t, w, h
        .Move pbChart.Left, t, w, h
        .ZOrder
    End With
    
    With vseScrollSep
        h = Screen.TwipsPerPixelY * 2
        ''w = cmdSettings.Left
        w = Me.ScaleWidth
        .Move 0, hsb.Top - h, w, h
    End With

    '???
    'fraCts.left = btnPrint.left - (fraCts.Width - btnPrint.Width)
    'fraCts.top = 0
    'fraCts.ZOrder

    ' floating tips
    With vseTipY
        If vseOrderBar.Visible Then
            .Left = Me.ScaleWidth - .Width - vseOrderBar.Width
        Else
            .Left = Me.ScaleWidth - .Width ' - Screen.TwipsPerPixelX '* 2
        End If
        .ZOrder
        lblTipY.Move 0, 0, .Width, .Height
    End With
    With vseTipX
        '.Top = vseScrollSep.Top - .Height - 30 ' * 2
        .ZOrder
        lblTipX.Move 0, 0, .Width, .Height
    End With

    With vseDetach
        .BackColor = &H8000000F     'button face
        If (WindowState = 2) And SubclassingEnabled And m.eDetachStatus = eNotDetached Then     'And m.bAllowDetach Then
            .Move Me.ScaleWidth - .Width - 30, (vseCaption.Height - .Height) \ 2
            .Visible = True
        Else
            .Visible = False
        End If
    End With
        
    With vsePeriodLink
        If (WindowState = 2) And SubclassingEnabled And Not m.bGameMode And _
            m.eDetachStatus = eNotDetached And m.Chart.TypeOfChart <> eTypeChart_Seasonal Then
            
'            If m.bAllowDetach Then
                .Move vseDetach.Left - .Width, vseDetach.Top
'            Else
'                .Move Me.ScaleWidth - .Width - 30, (vseCaption.Height - .Height) \ 2
'            End If
            If m.WindowLink.PeriodColor = 0 Then
                .BackColor = vseDetach.BackColor
                .ForeColor = 0
            Else
                .BackColor = m.WindowLink.PeriodColor
                .ForeColor = RGB(255, 255, 255)
            End If
            .Visible = True
        Else
            .Visible = False
        End If
    End With
    
    With vseSymbolLink
        If (WindowState = 2) And SubclassingEnabled And Not m.bGameMode And _
            m.eDetachStatus = eNotDetached And m.Chart.TypeOfChart <> eTypeChart_Seasonal Then
            
            .Move vsePeriodLink.Left - .Width, vsePeriodLink.Top
            If m.WindowLink.SymbolColor = 0 Then
                .BackColor = vseDetach.BackColor
                .ForeColor = 0
            Else
                .BackColor = m.WindowLink.SymbolColor
                .ForeColor = RGB(255, 255, 255)
            End If
            .Visible = True
        Else
            .Visible = False
        End If
    End With

    hsb.ZOrder
    vseTipX.ZOrder
    
    FixOrderBarControls True, True
    
    m.Chart.GenerateChart
    
    If bLockedWindow Then LockWindowUpdate 0
    
    If Me.MDIChild Then
        If Me Is ActiveChart Then
            If Me.WindowState = vbMaximized Then
                MoveFocus pbChart
            End If
        End If
    End If
    
    'If Screen.ActiveControl = hsb Then
        'move focus so "active" rectangle on scrollbar
        'won't look so goofy (since not resized)
        ' TLB 8/12/05: don't do this anymore -- it is causing
        ' problems when changing pages under certain conditions
        ''SetFocusCtl
    'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim strText$, frm As Form
    Dim i&, hNext&
    Dim bLocked As Boolean
    
    Disable tmr
    tmr.Tag = "UNLOADING"
    m.iAutoSize = False
                
    LogEvent "Unload"
        
    Set m.oToolTip = Nothing
    If Not g.bUnloading Then Set m.WindowLink = Nothing     '5551
                
    UnloadEditors
    TextIncDecUnregisterForm Me
        
    Set g.ChartGlobals.frmLastChartMouseMove = Nothing
    If m.bGameMode Then
        If Not m.oGameMode Is Nothing Then
            If m.oGameMode.GameReplayMode(tmrGameMode.Enabled) <> eGDReplayMode_Sync Then
                m.oGameMode.ClearReplayAll
                g.ChartGlobals.nGameInProg = 0
            End If
        End If
    End If
    If m.eDetachStatus = eNotDetached Then    '4929, 4976
        If Me.Visible And Me.WindowState = vbMaximized And Not g.bUnloading Then
            If vsTab.Enabled Then
                bLocked = LockWindowUpdate(frmMain.hWnd)
            End If
            'this code helps prevens system menu bar from getting grayed out under this one circumstance
            'a. close a maximized chart, b. open a modal dialog & close dialog --> menu gets grayed without this code
            'this code also fixes the order-bar not getting correctly placed when a chart is closed
            hNext = GetNextWindow(Me.hWnd, GW_HWNDNEXT)
            For i = 0 To Forms.Count - 1
                If Forms(i).hWnd = hNext Then
                    If IsFrmChartMDI(Forms(i)) Then
                        Forms(i).WindowState = vbMaximized
                    End If
                    Exit For
                End If
            Next
            If bLocked Then frmMain.tmrMain.Tag = "UnlockWindowUpdate"
        End If
    End If
    
    If fgChartFlex.Count > 1 Then
        If m.bFlexOrdBar Then Unload fgChartFlex(eFlexGridIdx_OrdWizard)
        If m.bFlexSeasonal Then Unload fgChartFlex(eFlexGridIdx_Seasonal)
        
        If m.bFlexPFP Then
            Unload fgChartFlex(eFlexGridIdx_PfpInd)
            Unload fgChartFlex(eFlexGridIdx_PfpHits)
        End If
        
        m.bFlexOrdBar = False
        m.bFlexSeasonal = False
        m.bFlexPFP = False
    End If
    
    If Not m.Chart Is Nothing Then m.Chart.DestroyChartRefs
    Set m.Chart = Nothing
    Set m.aTabs = Nothing
    Set m.oGameMode = Nothing
    Set m.AnnotOptions = Nothing
    Set m.aABarColHeader = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".Form_Unload"
    Resume ErrExit
    
End Sub

Private Sub fraOrderBtns_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        HandleBuySellClick fraOrderBtns, Button
    End If

End Sub

Private Sub gdSeasonalDateFrom_Changing()
On Error Resume Next:
    
    If Not Me.Visible Then Exit Sub
    If m.eSeasonalCtrlsState <> eSeasonCtrlStatus_Updated And Not vseInvalid.Visible Then Exit Sub
    HandleSeasonalInput eSeasonalCtrl_FromDate

End Sub

Private Sub gdTrendColor_Changed(Index As Integer)
On Error Resume Next:

    HandleSeasonalInput Index
    
End Sub

Private Sub gdTrendColor_DropDown(Index As Integer)
    tmr.Enabled = False
End Sub

Private Sub hsb_Change()
On Error GoTo ErrSection:

    Static bInProgress As Boolean

    If bInProgress Or m.Chart Is Nothing Then Exit Sub
    
    bInProgress = True
        
    'If label_mode = 0 Then
    '    lblMouse.Font.Bold = False
    '    ShowText lblMouse, drag_note
    '    ShowText lblBlank, ""
    'End If
    
    With hsb
        If .Value > .Max Then
            .Value = .Max
        ElseIf .Value < .Min Then
            .Value = .Min
        End If

        ' if holding down the scroll change, go to "faster" scroll
        If gdTickCount < m.dLastScrollTime + 60 Then
            .SmallChange = m.Chart.geChartPoints \ 20 + 1 'MJM - modified for use with grapheng.dll
        End If
    End With

    If m.Chart.Zoomed Then
        If m.Chart.AutoScale Then
            geZoomModeAuto m.Chart.geChartObj
        Else
            geZoomModeManual m.Chart.geChartObj
        End If
    End If
    
    If Not m.bChartMoveInProg And Not g.bStarting And m.Chart.RedoMode = eRedo1_Scrolled Then
        geAnnotMove m.Chart.geChartObj, 1
    End If
    m.Chart.GenerateChart eRedo1_Scrolled
        
    ' save time so won't do timer if scroll held down
    m.dLastScrollTime = gdTickCount

    
ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError Me.Name & ".hsb_Change"
    Resume ErrExit
    
End Sub

Private Sub hsb_GotFocus()
On Error GoTo ErrSection:

    'move focus so "active" rectangle on scrollbar won't look
    'so goofy (since doesn't seem to properly resize when it should)
    SetFocusCtl

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".hsb_GotFocus"
    Resume ErrExit
    
End Sub

Private Sub hsb_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    KeyPress KeyCode, Shift

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".hsb_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub hsb_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".hsb_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub hsb_Scroll(ByVal ScrollValue As Double, ScrollTipText As String)
On Error Resume Next

    Dim nMid&, strDate$, dDate#

    ' set scroll tip to date in middle of range
    nMid = hsb.Value - m.Chart.geChartPoints \ 2
    dDate = m.Chart.aXdate(nMid)
    If dDate > 0 Then
        dDate = DateTimeConvert(m.Chart.Bars, dDate)
        If IsIntraday(m.Chart.Periodicity) Then
            strDate = DateFormat(dDate, MM_DD_YY) & Format(dDate, " Hh:Nn")
        Else
            strDate = DateFormat(dDate, MM_DD_YYYY)
        End If
        ScrollTipText = strDate
    End If

End Sub

Private Sub imgTbBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:

    If InStr(g.strActiveDraw, "PFP") <> 0 Then
        ClearAnnotFlags True, True
        g.strActiveDraw = ""
        SyncDrawTools
    End If
    
    Set m.oBtnMouseLast = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_LBUTTONDOWN, Button, Index, X, Y, False)

End Sub

Private Sub imgTbBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:
    
    If InStr(g.strActiveDraw, "PFP") <> 0 Then
        ClearAnnotFlags True, True
        g.strActiveDraw = ""
        SyncDrawTools
    End If
    
    Set m.oBtnMouseLast = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_LBUTTONUP, Button, Index, X, Y, False)

End Sub

Private Sub imgTbBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:

    Dim oButton As cPicBoxButton

    If InStr(g.strActiveDraw, "PFP") <> 0 Then
        ClearAnnotFlags True, True
        g.strActiveDraw = ""
        SyncDrawTools
    End If
    
    Set oButton = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_MOUSEMOVE, Button, Index, X, Y, False)
    If Not oButton Is m.oBtnMouseLast Then
        ClearLastMouseButton Me, m.oBtnMouseLast
        Set m.oBtnMouseLast = oButton
        If Not m.oBtnMouseLast Is Nothing Then
            m.oBtnMouseLast.BtnToolTipShow Me, pbTbBack(Index)
        End If
    End If
    
End Sub

Private Sub imgTbBackDraw_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:

    If InStr(g.strActiveDraw, "PFP") <> 0 Then
        ClearAnnotFlags True, True
        g.strActiveDraw = ""
        SyncDrawTools
    End If
    
    Set m.oBtnMouseLast = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_LBUTTONDOWN, Button, Index, X, Y, True)

End Sub

Private Sub imgTbBackDraw_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:

    If InStr(g.strActiveDraw, "PFP") <> 0 Then
        ClearAnnotFlags True, True
        g.strActiveDraw = ""
        SyncDrawTools
    End If
    
    Set m.oBtnMouseLast = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_LBUTTONUP, Button, Index, X, Y, True)

End Sub

Private Sub imgTbBackDraw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:

    Dim oButton As cPicBoxButton

    If InStr(g.strActiveDraw, "PFP") <> 0 Then
        ClearAnnotFlags True, True
        g.strActiveDraw = ""
        SyncDrawTools
    End If
    
    Set oButton = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_MOUSEMOVE, Button, Index, X, Y, True)
    If Not oButton Is m.oBtnMouseLast Then
        ClearLastMouseButton Me, m.oBtnMouseLast
        Set m.oBtnMouseLast = oButton
        If Not m.oBtnMouseLast Is Nothing Then
            m.oBtnMouseLast.BtnToolTipShow Me, pbTbBack(Index)
        End If
    End If
    
End Sub

Private Sub lblAccounts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        HandleBuySellClick lblAccounts, Button
    End If
    
End Sub

Private Sub lblAutoExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        HandleBuySellClick lblAutoExit, Button
    End If

End Sub

Private Sub lblEquity_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        HandleBuySellClick lblEquity, Button
    End If


End Sub

Private Sub lblTradePos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        HandleBuySellClick lblTradePos, Button
    End If

End Sub

Private Sub lblTipChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Static nPrevTicks&
    Dim nTicks&
    
    nTicks = gdTickCount
    
    If nPrevTicks = 0 Or Abs(nTicks - nPrevTicks) > 300 Then
        ' get chart coordinates of mouse move
        m.MouseLast = GetChartCoordinates(X, Y, Shift)
        m.MouseLast.nButton = Button
        
        DoHitTest X, Y, False, Button
        nPrevTicks = nTicks
    End If
    
End Sub

Private Sub lblTipX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    'MJM - Q&A: need to duplicate this?
    'Peg_MouseMove Button, Shift, X + vseTipX.Left - Peg.Left, Y + vseTipX.Top - Peg.Top
    With vseTipX
        If .Top < vseScrollSep.Top Then
            .Top = vseScrollSep.Top
        Else
            .Top = vseScrollSep.Top - .Height - 30
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".lblTipX_MouseMove"
    Resume ErrExit
    
End Sub

Private Sub lblTipY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If FormIsLoaded("frmElliot") Then
        If frmElliot.Visible Then RefreshTips -1000, -1000      'so pbchart control will get mouse move event
    End If

End Sub

Private Sub lblWizardPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:

    Static dPrevX#, dPrevY#
    
    If X = dPrevX And Y = dPrevY Then Exit Sub
    
    Dim Annot As cAnnotation
    Dim pt As POINTAPI
    Dim nX As Single, nY As Single
    
    nX = 0
    nY = 0
    
    If cmdPut.Left = cmdCall.Left Then  '09-16-2014: as of this date, this is true only for Risk Reward Visualizer tool
        '7020: fix for jumpy mouse & excessive flashing
        
        nX = m.Chart.PixelsPerBar
        If nX > 5 Then
            nX = cmdPut.Width
        Else
            nX = nX * Screen.TwipsPerPixelX
        End If
        If Abs(X - dPrevX) < nX Then
            If Abs(Y - dPrevY) > 2 * Screen.TwipsPerPixelY Then
                If GetCursorPos(pt) Then
                    If ScreenToClient(pbChart.hWnd, pt) Then
                        nX = dPrevX
                        nY = pt.Y * Screen.TwipsPerPixelX
                        
                        Dim MouseCoord As ChartCoordinates
                        MouseCoord = GetChartCoordinates(nX, nY)
                        m.MouseLast.dY = MouseCoord.dY
                        
                        dPrevX = nX
                        
                        HandleWizardPrompt nX, nY
                    End If
                End If
            End If
            
            Exit Sub
        End If
    End If
    
    dPrevX = X
    dPrevY = Y
    
    With lblWizardPrice
        If Len(.Caption) = 0 Or .Width = kPriceWizWidth Then
            .MousePointer = vbDefault
        Else
            .MousePointer = vbCustom
            .MouseIcon = pbChart.MouseIcon
        End If
        
        If GetCursorPos(pt) Then
            If ScreenToClient(pbChart.hWnd, pt) Then
                nX = pt.X * Screen.TwipsPerPixelX
                nY = pt.Y * Screen.TwipsPerPixelX
            End If
        End If
        
        If nX = 0 Or nY = 0 Then
            StatusMsg "GetCursorPos or ScreenToClient failed"
            'JM 09-16-2014: original code is inexact; use as backup
            nX = .Left + .Width / 2 + fraWizardPrompt.Left - pbChart.Left
            nY = Y + .Top + fraWizardPrompt.Top - pbChart.Top
        End If
    End With
    pbChart_MouseMove Button, Shift, nX, nY

End Sub

Private Sub lblWizardPrice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim i&, strText$, strOrderType$
    Dim dMinMove#, dPriceLast#, dPriceNew#
    
    Dim Order As cPtOrder
    Dim eOrdType As enumOneClickOrder

    If Not m.AnnotOptions Is Nothing Then GoTo ErrExit

    With lblWizardPrice
        If .Width > kPriceWizWidth Then
            Set Order = New cPtOrder
            
            i = m.Chart.LastGoodDataBar(False)
            dMinMove = m.Chart.Bars.MinMove()
            dPriceLast = RoundToMinMove(m.Chart.Bars(eBARS_Close, i), dMinMove)
            dPriceNew = m.Chart.Bars.RoundToPrice(m.MouseDown.dY, m.MouseDown.dDate)
            
            SetOrderType Order, eClickOrder_None, dPriceNew, dPriceLast
          
            If Order.Buy Then
                eOrdType = eClickOrder_BuyMkt
            Else
                eOrdType = eClickOrder_SellMkt
            End If
            
            If Order.OrderType = eTT_OrderType_Stop Then
                strOrderType = "S"
            Else
                strOrderType = "L"
            End If
                        
            If InStr(.Caption, "-") Then
                strText = Parse(.Caption, "-", 1) & ";" & Parse(.Caption, vbCrLf, 2)
            Else
                strText = Parse(.Caption, vbCrLf, 1) & ";" & Parse(.Caption, vbCrLf, 2)
            End If
                
            WizardGridAdd eOrdType, strText, Parse(strText, ";", 1), strOrderType, Nothing, True
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".lblWizardPrice_MouseUp"

End Sub

Private Sub mnuAccountBar_Click()
On Error Resume Next

    m.Chart.ShowAccountBar = Not m.Chart.ShowAccountBar
    Form_Resize
    
End Sub

Private Sub mnuAnnotAddPt_Click()
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    Dim Alert As cAlert
    Dim nNewPoint&, i&
    
    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    
    If Not Annot Is Nothing Then
        If mnuAnnotAddPt.Caption = "Delete line" Then
            Annot.HideFibLine Annot.HitItemIndex
            m.Chart.SyncGlobalAnnots Nothing, Annot.MultiChartFlag       '4301
        ElseIf mnuAnnotAddPt.Caption = "Alert" Or Annot.Prop("FakePriceAlert") = 1 Then
            If Annot.MultiChartFlag Then
                MultiChartAlert m.Chart, Annot
            Else
                Set Alert = Annot.AlertObject()
                If Alert Is Nothing Then
                    Set Alert = Annot.AlertObject(True)
                    i = 1
                End If
                Annot.AlertAddEditInprog = True
                If Not frmAlerts.ShowMe(Alert, eGDAlertType_Annot) And i = 1 Then
                    'a new alert object was created but user cancelled or
                    'user does not have minimum required level/module -- remove it
                    Annot.UpdateAlert 0
                End If
                Annot.AlertAddEditInprog = False
            End If
        ElseIf Annot.eType = eANNOT_BalloonStrangle Then
            If Val(Annot.Prop("ShowCollapsed")) = 1 Then
                Annot.Prop("ShowCollapsed") = 0
            Else
                Annot.Prop("ShowCollapsed") = 1
            End If
            m.Chart.GenerateChart eRedo1_Scrolled
        ElseIf Annot.eType = eANNOT_Bracket Then
            i = ValOfText(Annot.Prop("BracketDirection"))
            If i = 0 Then
                Annot.Prop("BracketDirection") = 1
            Else
                Annot.Prop("BracketDirection") = 0
            End If
            m.Chart.GenerateChart eRedo1_Scrolled
        Else
            nNewPoint = Annot.AddPoint(m.MouseDown.dDate, m.MouseDown.dY + 0.01)
            If nNewPoint = -1 Then
                ClearAnnotFlags False
                m.nActiveIndIdx = 0
                m.nObjectMoving = 0
                m.Chart.SetCursor
            Else
                Annot.MenuAdd = True
                m.nActiveAnnotPt = nNewPoint
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuAnnotAddPt_Click"

End Sub

Private Sub mnuAnnotDeletePt_Click()
On Error GoTo ErrSection:

    Dim Annot As cAnnotation, i&
    Dim Alert As cAlert
    Dim bMultiChart As Boolean
    
    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    
    If Not Annot Is Nothing Then
        If mnuAnnotDeletePt.Caption = "Delete" Then
            If Annot.eUsage = eANNOT_IndicatorLabel Then
                HandleAnnotDblclk
            ElseIf Annot.eUsage = eANNOT_PriceAlert Then
                Set Alert = Annot.AlertObject
                If Not Alert Is Nothing Then
                    Alert.UpdateChartObject True
                    g.Alerts.Remove Alert.AlertKey
                    If FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid
                End If
            Else
                bMultiChart = Annot.MultiChartFlag
                Annot.geRemoveAnnotation m.Chart.geChartObj
                m.Chart.Annots.Remove (m.nActiveAnnotIdx)
                m.Chart.SyncGlobalAnnots Nothing, bMultiChart
            End If
        Else
            Annot.DeletePoint m.nActiveAnnotPt
            
            ClearAnnotFlags False
            m.nActiveIndIdx = 0
            m.nObjectMoving = 0
            
            Annot.geDrawAnn Chart
            Set Annot = Nothing
            m.Chart.SetCursor
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuAnnotDeletePt_Click"

End Sub

Private Sub mnuAnnotDuplicate_Click()
On Error GoTo ErrSection:

    HandleAnnotDblclk           '4560

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuAnnotDuplicate_Click"

End Sub

Private Sub mnuAnnotMovePt_Click()
On Error GoTo ErrSection:

    Dim sX As Single, sY As Single
    Dim Annot As cAnnotation
    Dim strSave$
    
    If mnuAnnotMovePt.Caption = "Edit" Then
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Not Annot Is Nothing Then
            If Annot.eType = eANNOT_HorzLine And Annot.Prop("FakePriceAlert") = 1 Then
                mnuAnnotAddPt_Click
            Else
                tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
            End If
        End If
    ElseIf mnuAnnotMovePt.Caption = "Delete" Then
        'JM 01-02-2014: this code should only execute for EWI label created by an end user who
        'shared chart page with another user who does not have EWL eneablement
        'do this to let click routine delete the label
        strSave = mnuAnnotDeletePt.Caption
        mnuAnnotDeletePt.Caption = "Delete"
        mnuAnnotDeletePt_Click
        mnuAnnotDeletePt.Caption = strSave
    ElseIf mnuAnnotMovePt.Caption = "Hide EWI" Then
        m.Chart.ShowHideAnnotsByType eANNOT_ElliotLabel, 1, True
        m.Chart.GenerateChart eRedo1_Scrolled
        m.Chart.SyncToolbar True
    Else
        sX = m.MouseDown.MouseX
        sY = m.MouseDown.MouseY + 1
        m.nObjectMoving = kMinMoveCount
        
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Not Annot Is Nothing Then
            Annot.MenuMove = True
            pbChart_MouseMove vbLeftButton, 0, sX, sY       'aardvark 3678
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuAnnotMovePt_Click"

End Sub

Private Sub mnuAnnotStyle_Click()
On Error GoTo ErrSection:

    Dim Annot As cAnnotation, i&
    
    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    
    If Not Annot Is Nothing Then
        Select Case Annot.eType
            Case eANNOT_Bracket
                i = ValOfText(Annot.Prop("BracketStyle"))
                If i = 0 Then
                    Annot.Prop("BracketStyle") = 1
                Else
                    Annot.Prop("BracketStyle") = 0
                End If
            Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                 eANNOT_Fibonacci4, eANNOT_FibExpansion, eANNOT_DanCodeFib, _
                 eANNOT_AdvRiskReward
                If mnuAnnotStyle.Caption = "Bold line" Then
                    Annot.FibLineBold = True
                Else
                    Annot.FibLineBold = False
                End If
        End Select
        
        m.Chart.GenerateChart eRedo1_Scrolled
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuAnnotStyle_Click"
    Resume ErrExit
End Sub

Private Sub mnuAutoScale_Click()
On Error GoTo ErrSection:

    m.Chart.AutoScale = Not m.Chart.AutoScale
    
    If m.Chart.Zoomed And m.Chart.AutoScale Then
        geZoomModeAuto m.Chart.geChartObj
    End If
    
    m.Chart.GenerateChart eRedo1_Scrolled

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuAutoScale_Click"
    Resume ErrExit
End Sub

Private Sub mnuAutoTrade_Click()
On Error GoTo ErrSection:

    m.Chart.AutoTrade = Not m.Chart.AutoTrade

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuAutoTrade_Click"
    Resume ErrExit
End Sub

Private Sub mnuBarPeriod_Click(Index As Integer)
On Error GoTo ErrSection:

    m.Chart.ChangeBarPeriod StripStr(mnuBarPeriod(Index).Caption, "&")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuBarPeriod_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuBollinger_Click()
On Error GoTo ErrSection:

    m.Chart.BarDisplayType = eINDIC_BollingerBar
    m.Chart.GenerateChart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuBollinger_Click"
    Resume ErrExit

End Sub

Private Sub mnuChartCapture_Click()
On Error GoTo ErrSection:

    Dim rc&, strFile$
    Dim nTop&, nLeft&, nWidth&, nHeight&
    Dim cy&

    DoEvents        'let the menu go away before capturing the screen

    strFile = App.Path & "\ScreenCapture.bmp"
    
    cy = GetSystemMetrics(55) + 2      'SM_CYMENUSIZE (winuser.h)
    
    If Me.WindowState = vbMaximized Then
        nLeft = 0
        nHeight = 0
        nWidth = pbChart.Width / Screen.TwipsPerPixelX + 5
        nHeight = vsTab.Height / Screen.TwipsPerPixelY + 5
    Else
        nLeft = Me.ScaleLeft / Screen.TwipsPerPixelX
        nTop = Me.ScaleTop / Screen.TwipsPerPixelY - cy
        nWidth = Me.ScaleWidth / Screen.TwipsPerPixelX
        nHeight = Me.ScaleHeight / Screen.TwipsPerPixelY + cy
    End If
    
    rc = geCaptureScreen(m.Chart.geChartObj, Me.hWnd, _
                strFile, _
                nLeft, _
                nTop, _
                nWidth, _
                nHeight, 1)
    
    If rc = 0 Then
        frmPrintPreview.ShowMe "", Me, "Screen", 0.5, 0.5, 0.5, 0.5, True
    Else
        MsgBox "Screen capture failed."
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuChartCapture_Click"
    Resume ErrExit

End Sub

Private Sub mnuChartCopyMove_Click()
On Error GoTo ErrSection:
    
    frmTemplates.ShowMe eMode_ChartCopyMove, m.Chart
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuChartCopyMove_Click"

End Sub

Private Sub mnuClearAnnotMultiFlag_Click()
On Error GoTo ErrSection:

    m.Chart.ClearAnnotMultiFlag

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuClearAnnotMultiFlag_Click"

End Sub

Private Sub mnuContracts_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim strSymbol$, nSymbolID&
    
    strSymbol = Parse(mnuContracts.Item(Index).Caption, "(", 1)
    If InStr(strSymbol, "Active Contract") <> 0 Then
        strSymbol = RollSymbolForDate(m.Chart.Symbol, m.Chart.Bars(eBARS_DateTime, m.Chart.Bars.Size - 1))
    End If
    
    nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
    
    If nSymbolID > 0 Then
        m.Chart.SetSymbol nSymbolID, True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuContracts_Click"
    
End Sub

Private Sub mnuDataCopy_Click()
On Error GoTo ErrSection:

    If HasModule("INDCOPY") Then            '6518
        If InfBox("Include indicators?", "?", "-Yes|+No", "Copy chart's data to clipboard") = "Y" Then
            DataToClipboard Me, True
        Else
            DataToClipboard Me, False
        End If
    Else
        DataToClipboard Me, False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuDataCopy_Click"
    Resume ErrExit

End Sub

Private Sub mnuDeleteAnnot_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim dDate As Double, bJustLastAnnot As Boolean

    If Index = 0 Then
        bJustLastAnnot = True
    ElseIf Index = 1 Then
        ' delete annotations prior to date (default to where the cursor is currently)
        dDate = Int(m.MouseLast.dDate)
        If dDate <= 0 Then dDate = Date
        dDate = frmAskDate.ShowMe("Delete annotations prior to date:", dDate)
        If dDate <= 0 Then Exit Sub
    End If

    If m.Chart.RemoveAnnots(bJustLastAnnot, , , , , dDate) > 0 Then
        m.Chart.SyncGlobalAnnots Nothing, True
    Else
        Beep
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuDeleteAnnot_Click"
    Resume ErrExit
End Sub

Private Sub mnuDetAttChart_Click()
On Error Resume Next

    If m.eDetachStatus = eNotDetached Then
        DetachChart Me
    Else
        AttachChart Me
    End If

End Sub

Private Sub mnuEditSystem_Click()
On Error GoTo ErrSection:

    Dim nID&
    Dim nLinkToChart As Long
    Dim frm As frmSystemManager
    
    If HasGold(True) Then
        nID = m.Chart.SystemID
        If nID > 0 Then
            If Not ActivateEditor("frmSystemManager", nID) Then
                Set frm = New frmSystemManager
                nLinkToChart = GetIniFileProperty("LinkToChart", 1, "Systems", g.strIniFile)
                frm.ShowMe nID, , False, "", , True
                If nLinkToChart = 1 Then frm.UseChartSystem 1
            End If
        Else
            Beep
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuEditSystem_Click"
    Resume ErrExit
End Sub

Private Sub mnuHideAnnots_Click(Index As Integer)
On Error GoTo ErrSection:

    If Index = 0 Then
        If mnuHideAnnots(0).Checked Then            '6984
            m.Chart.ShowHideAnnotsByType eANNOT_ElliotLabel, 0, True
        Else
            m.Chart.ShowHideAnnotsByType eANNOT_ElliotLabel, 1, True
        End If
        m.Chart.SyncToolbar True
        HideAnnotations (Not mnuHideAnnots(0).Checked)
    ElseIf Index = 1 Then
        If mnuHideAnnots(1).Caption = "Show EWI" Then
            m.Chart.ShowHideAnnotsByType eANNOT_ElliotLabel, 0, True
        Else
            m.Chart.ShowHideAnnotsByType eANNOT_ElliotLabel, 1, True
        End If
        m.Chart.GenerateChart eRedo1_Scrolled
        m.Chart.SyncToolbar True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuHideAnnots_Click"
    Resume ErrExit
End Sub

Private Sub mnuDisableRT_Click()
On Error GoTo ErrSection:

    m.Chart.DisableRT = Not mnuDisableRT.Checked
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuDisableRT_Click"

End Sub

Private Sub mnuHideScrollbars_Click()
On Error GoTo ErrSection:

    Dim i&

    g.ChartGlobals.bHideScrollbars = Not g.ChartGlobals.bHideScrollbars
    Form_Resize
    
    UpdateVisibleCharts -1 ' to resize all charts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuHideScrollbars_Click"
End Sub

Private Sub mnuLogModeDraw_Click()
On Error GoTo ErrSection:

    g.ChartGlobals.bLogModeDraw = Not g.ChartGlobals.bLogModeDraw
    UpdateVisibleCharts eRedo1_Scrolled

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuLogModeDraw_Click"

End Sub

Private Sub mnuLogScale_Click()
On Error GoTo ErrSection:

    If m.Chart.ChartLogFlag = ePANE_LogFlagLog Then
        m.Chart.ChartLogFlag = ePANE_LogFlagLinear
    Else
        m.Chart.ChartLogFlag = ePANE_LogFlagLog
    End If
    
    m.Chart.GenerateChart eRedo1_Scrolled

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuLogScale_Click"
    Resume ErrExit
End Sub

Private Sub mnuManageXOS_Click()
On Error GoTo ErrSection:

    frmOrderStrategies.ShowMeManage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuManageXOS_Click"
    
End Sub

Private Sub mnuMoreDataSingleSym_Click()
On Error GoTo ErrSection:

    frmSingleSymHistory.ShowMe m.Chart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuMoreDataSingleSym_Click"
    Resume ErrExit

End Sub

Private Sub mnuOrdBarDefaults_Click()
On Error GoTo ErrSection:
    
    m.Chart.OrdBarGetDefaults
    Form_Resize
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuOrdBarDefaults_Click"

End Sub

Private Sub mnuOrdBarSettings_Click()
On Error GoTo ErrSection:

    frmChartOrdBar.ShowMe Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuOrdBarSettings_Click"
    
End Sub

Private Sub mnuOrderAcctHistory_Click()
On Error GoTo ErrSection:

    frmTTPositions.ShowMe m.Chart.TradeAccountID, g.Broker.AccountTypeForID(m.Chart.TradeAccountID)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuOrderAcctHistory_Click"
    
End Sub

Private Sub mnuOrderBar_Click()
On Error Resume Next

    If m.eOrdBarMode = eOrdBarMode_PFP Or m.eOrdBarMode = eOrdBarMode_Undefined Then
        If m.Chart.ShowTrades = 2 Then m.Chart.ShowTrades = 0
        If m.eOrdBarMode = eOrdBarMode_Undefined Then m.eOrdBarMode = eOrdBarMode_Order
        ToggleOrderBar vseOrderBar.Visible, True
    Else
        If m.Chart.ShowTrades = 2 Then
            m.Chart.ShowTrades = 0
            If m.eOrdBarMode = eOrdBarMode_Wizard Then OrderBarModeToggle       '4935, 4998
        Else
            m.Chart.ShowTrades = 2
            m.eOrdBarMode = eOrdBarMode_Order
        End If
    End If
    
    m.Chart.GenerateChart eRedo1_Scrolled

End Sub

Private Sub mnuOrderCancel_Click()
On Error GoTo ErrSection:

    CancelOrderFromID m.nActiveOrderID, "Chart", True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuOrderCancel_Click"

End Sub

Private Sub mnuOrderCheckStatus_Click()
On Error GoTo ErrSection:

    g.Broker.CheckTradeServerOrders

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuOrderCheckStatus_Click"

End Sub

Private Sub mnuOrderEdit_Click()
On Error GoTo ErrSection:

    EditOrderFromID m.nActiveOrderID, "Chart"

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuOrderEdit_Click"

End Sub
 
Private Sub mnuOrderHistory_Click()
On Error GoTo ErrSection:

    frmOrderHistory.ShowMeForID m.nActiveOrderID

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuOrderHistory_Click"

End Sub

Private Sub mnuOrderJournal_Click()
On Error GoTo ErrSection:

    g.TnJournal.ShowOrderJournalForID m.nActiveOrderID

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuOrderJournal_Click"

End Sub

Private Sub mnuOrderPark_Click()
On Error GoTo ErrSection:

    ParkOrderFromID m.nActiveOrderID, "Chart"

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuOrderPark_Click"

End Sub

Private Sub mnuOrderSubmit_Click()
On Error GoTo ErrSection:

    SubmitOrderFromID m.nActiveOrderID, "Chart"

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".mnuOrderSubmit_Click"

End Sub

Private Sub mnuPPB_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim i&
    
    i = ValOfText(mnuPPB(Index).Caption)
    m.Chart.PixelsPerBar = i
    m.Chart.GenerateChart eRedo1_Scrolled

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuPPB_Click"
    Resume ErrExit
End Sub

Private Sub mnuCandlesticks_Click()
On Error GoTo ErrSection:

    m.Chart.BarDisplayType = eINDIC_Candlestick
    m.Chart.GenerateChart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuCandlesticks_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuCloseLine_Click()
On Error GoTo ErrSection:

    m.Chart.BarDisplayType = eINDIC_Line
    m.Chart.GenerateChart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuCloseLine_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuCoarseGrid_Click()
On Error GoTo ErrSection:

    With m.Chart
        If .VertGrid = 0 Then
            .VertGrid = 1
        Else
            .VertGrid = 0
        End If
        .GenerateChart
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuCoarseGrid_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuCrosshairs_Click()
On Error GoTo ErrSection:

    ToolbarSetCursorGroup frmMain.tbToolbar, False, "ID_CursorCrosshairs"
    m.Chart.SetCursor
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuCrosshairs_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuCursorArrow_Click()
On Error GoTo ErrSection:

    ToolbarSetCursorGroup frmMain.tbToolbar, False, "ID_CursorArrow"
    m.Chart.SetCursor

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuCursorArrow_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuCursorHoriz_Click()
On Error GoTo ErrSection:

    ToolbarSetCursorGroup frmMain.tbToolbar, False, "ID_CursorHorizLine"
    m.Chart.SetCursor

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuCursorHoriz_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuCursorVert_Click()
On Error GoTo ErrSection:

    ToolbarSetCursorGroup frmMain.tbToolbar, False, "ID_CursorVertLine"
    m.Chart.SetCursor

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuCursorVert_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuHidePane_Click()
On Error GoTo ErrSection:

    Dim strPane$, Pane As cPane
    
    strPane = Trim(mnuHidePane.Tag)
    Set Pane = m.Chart.Tree(strPane)
        
    If Pane Is Nothing Then Exit Sub
            
    If strPane = "PRICE PANE" Then
        m.Chart.HidePriceIndicators = Not m.Chart.HidePriceIndicators
        'm.Chart.GenerateChart
        g.bDirtyChartPage = True
    ElseIf m.Chart.geVisiblePaneCnt > 1 Then
        Pane.Display = False
        g.bDirtyChartPage = True
    End If
    
    If g.bDirtyChartPage Then
        m.Chart.geForceRecalc
        m.Chart.GenerateChart
    End If
    
    Set Pane = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuHidePane_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuHotKeys_Click()
On Error GoTo ErrSection:

    frmMessage.ShowMe "Charting Hot Keys and Tips", "@" & App.Path & "\Info\HotKeys"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuHotKeys_Click"
    Resume ErrExit
    
End Sub

#If 0 Then
Private Sub mnuLockValues_Click()
On Error GoTo ErrSection:

    m.bLockValuesDisplay = Not m.bLockValuesDisplay

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuLockValues_Click"
    Resume ErrExit
    
End Sub
#End If

Private Sub mnuOHLC_Click()
On Error GoTo ErrSection:

    m.Chart.BarDisplayType = eINDIC_OHLC
    m.Chart.GenerateChart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuOHLC_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuPrint_Click()
On Error GoTo ErrSection:

    m.Chart.PrintChart 2

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuPrint_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuRemoveXOS_Click()
On Error GoTo ErrSection:

    RemoveXOS m.Chart.TradeAccountID, m.Chart.SymbolID, "Chart"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuRemoveXOS_Click"
    
End Sub

Private Sub mnuSaveImage_Click()
On Error GoTo ErrSection:

    PrintReport 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuSaveImage_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuScreenCapture_Click()
On Error GoTo ErrSection:

    Dim rc&, strFile$

    DoEvents        'let the menu go away before capturing the screen

    strFile = App.Path & "\ScreenCapture.bmp"
    rc = geCaptureScreen(m.Chart.geChartObj, frmMain.hWnd, _
                strFile, _
                frmMain.Left / Screen.TwipsPerPixelX, _
                frmMain.Top / Screen.TwipsPerPixelY, _
                frmMain.Width / Screen.TwipsPerPixelX, _
                frmMain.Height / Screen.TwipsPerPixelY, 0)
    
    If rc = 0 Then
        frmPrintPreview.ShowMe "", Me, "Screen", 0.5, 0.5, 0.5, 0.5, True
    Else
        MsgBox "Screen capture failed."
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuScreenCapture_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuSelectXOS_Click()
On Error GoTo ErrSection:

    SelectXOS m.Chart.TradeAccountID, m.Chart.SymbolID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuSelectXOS_Click"
    
End Sub

Private Sub mnuSettings_Click()
On Error GoTo ErrSection:

    ' TLB 5/19/2008: since the Edit Settings form is not modal, we should be able
    ' to just call it from here immediately (rather than through a timer event)
    'tmr.Tag = "EditSettings"
    DoEvents '(but first do this to make the popup menu disappear quicker)
    EditSettings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuSettings_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuShowToolbar_Click()
On Error GoTo ErrSection:

    Dim i&

'This code toggles toolbar on/off and will only execute if chart is detached (i.e. in frmChart2).
'There is no need to call ToolbarReset here, but do need to call ToolbarInit2 to set up 5.0 toolbar.

    If m.Chart.ShowToolbar = 0 Then
        m.Chart.ShowToolbar = 1
        ToolbarInit2 Me, m.aTbButtons
        ToolbarInit2 Me, m.aTbButtonsDraw, , kTbDraw, , g.vbeTbAlignDraw
        If g.vbeTbAlignDraw = vbAlignTop Or g.vbeTbAlignDraw = vbAlignBottom Then
            pbTbBackDraw(0).align = vbAlignTop
        Else
            pbTbBackDraw(0).align = vbAlignRight
        End If
        
        ToolbarResize2 Me, pbTbBack, imgTbBack, m.aTbButtons, m.bToolbarWrap
        ToolbarResize2 Me, pbTbBackDraw, imgTbBackDraw, m.aTbButtonsDraw, m.bToolbarWrap
        
        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
            If Not g.ChartGlobals.frmActiveNonDetached.Chart Is Nothing Then
                g.ChartGlobals.frmActiveNonDetached.Chart.SyncToolbar True
            End If
        End If
    Else
        m.Chart.ShowToolbar = 0
        m.aTbButtons.Size = 0
        m.aTbButtonsDraw.Size = 0
        For i = pbTbBack.UBound To 0 Step -1
            pbTbBack(i).Visible = False
        Next
        pbTbBackDraw(0).Visible = False
        m.Chart.SyncToolbar True
    End If
    
    Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuShowToolbar_Click"

End Sub

Private Sub mnuTemplate_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim strTemplate$
        
    strTemplate = mnuTemplate(Index).Caption
    'If Left(Trim(strTemplate), 1) = "(" Then
    If InStr(strTemplate, "<") > 0 Then
        frmTemplates.ShowMe eMode_Templates, m.Chart
    Else
        If Not m.Chart.TemplateApply(strTemplate) Then
            Beep
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuTemplate_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuTemplateManage_Click()
On Error GoTo ErrSection:

    frmTemplates.ShowMe eMode_Templates, m.Chart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuTemplateManage_Click"
    Resume ErrExit
End Sub

Private Sub mnuTemplateToOtherCharts_Click()
On Error GoTo ErrSection:

    CopySettingsToOtherCharts Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuTemplateToOtherCharts_Click"
    Resume ErrExit
End Sub

Private Sub mnuTips_Click()
On Error GoTo ErrSection:

    g.ChartGlobals.bFloatingTips = Not g.ChartGlobals.bFloatingTips
    UpdateVisibleCharts eRedo1_Scrolled

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuTips_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuTrades_Click()
On Error GoTo ErrSection:

    If m.Chart.ShowTrades = 1 Then      'ShowTrades is a long: 0=none,1=strategy,2=manual trade account
        m.Chart.ShowTrades = 0
        m.Chart.GenerateChart eRedo1_Scrolled
    'ElseIf m.Chart.SystemID > 0 Then
    '    m.Chart.ShowTrades = True
    '    m.Chart.GenerateChart eRedo1_Scrolled
    Else
        tmr.Tag = "AddSystem"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuTrades_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuUnsplit_Click()
On Error GoTo ErrSection:

    m.Chart.Unsplit = Not mnuUnsplit.Checked
    m.Chart.GenerateChart eRedo9_ReloadData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuUnsplit_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuUnzoom_Click()
On Error GoTo ErrSection:

    'MJM [begin] - modified for grapheng.dll
    If m.Chart Is Nothing Then Exit Sub
    
    If m.Chart.Zoomed Then
        m.Chart.UnzoomChart True
    End If
    'MJM [end] - modified for grapheng.dll

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuUnzoom_Click"
    Resume ErrExit
    
End Sub

Private Sub RefreshTips(ByVal X#, ByVal Y#, _
        Optional ByVal bScrolling As Boolean = False)
On Error Resume Next

    Dim nWidth&, strTip$, nPix&, nMaxY&, nPane&, bShow As Boolean
    Static nMinYwidth&, nPrevPane&

    nPix = Screen.TwipsPerPixelX
    If nMinYwidth = 0 Then
        nMinYwidth = Me.TextWidth("99.9")
    End If

    ' X-axis (date) tip
    With vseTipX
        bShow = True
        If X = -1000 Then
            .Tag = ""
        Else
            strTip = Trim(.Tag)
        End If
        If Len(strTip) = 0 Then
            bShow = False
        ElseIf g.ChartGlobals.bFloatingTips = False And bScrolling = False Then
            bShow = False
        End If
        If bShow Then
            If strTip <> lblTipX.Caption Then
                nWidth = Me.TextWidth(strTip) + 120 '+ nPix * 2
                If .Width < nWidth Then
                    lblTipX.Width = nWidth '- nPix * 2
                Else
                    nWidth = .Width
                End If
                lblTipX.Caption = strTip
                X = X - nWidth \ 2 + pbChart.Left
                If X + nWidth > Me.ScaleWidth Then
                    X = Me.ScaleWidth - nWidth
                End If
                If X < 0 Then X = 0
                If m.MouseLast.nPaneID < 0 Then
                    .Move X, vseScrollSep.Top, nWidth
                Else
                    .Move X, hsb.Top - .Height, nWidth
                End If
            End If
            'If .Visible = False Then
            '    .Visible = True
                '.ZOrder
            'End If
        'ElseIf .Visible Then
        Else
            '.Visible = False
            .Top = -1000 - .Height
            lblTipX.Caption = ""
        End If
    End With

    ' Y-axis (price) tip
    With vseTipY
        bShow = True
        If Y = -1000 Then .Tag = ""
        If Len(.Tag) = 0 Then
            bShow = False
        ElseIf g.ChartGlobals.bFloatingTips = False Then
            bShow = False
        End If
        If bShow Then
            strTip = Parse(.Tag, vbTab, 1)
            nPane = Val(Parse(.Tag, vbTab, 2))
            nWidth = Me.TextWidth(strTip)
            If nWidth < nMinYwidth Then nWidth = nMinYwidth
            nWidth = nWidth + 120  '+ nPix * 2
            If nWidth < 600 Then nWidth = 600
            
            If .Width < nWidth Or nPane <> nPrevPane Then
'                X = .Left - (nWidth - .Width)
                lblTipY.Width = nWidth
            Else
                nWidth = .Width
 '               X = .Left
            End If
            
            X = pbChart.Width - nWidth
            
            lblTipY.Caption = " " & strTip
            Y = pbChart.Top + Y - .Height \ 2
            If Y < 0 Then
                Y = 0
            ElseIf Y >= vseScrollSep.Top - .Height * 2 - 30 Then
                Y = -1000 - .Height
            End If
            .Move X, Y, nWidth
            'If .Visible = False Then
                '.Visible = True
                '.ZOrder
            'End If
        'ElseIf .Visible Then
        Else
            '.Visible = False
            .Top = -1000 - .Height
        End If
        nPrevPane = nPane
    End With

End Sub

Private Sub mnuViewJournals_Click()
On Error GoTo ErrSection:

    g.TnJournal.ShowJournals

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".mnuViewJournals_Click"

End Sub

Private Sub pbChart_GotFocus()
On Error Resume Next

    If Not m.Chart Is Nothing Then m.Chart.SyncToolbar

End Sub

Private Sub pbLeft_Click()
On Error GoTo ErrSection:

    OrderBarModeToggle

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".pbLeft_Click"

End Sub

Private Sub pbRight_Click()
On Error GoTo ErrSection:

    If (Not m.Chart Is Nothing) And g.nOptNavStatus <> eGDOptNavStatus_Loaded Then
        StartOptNav m.Chart.Symbol
    End If
    OrderBarModeToggle

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".pbRight_Click"

End Sub

Private Sub sldSpeed_Change()
On Error GoTo ErrSection:

    GameSpeed
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".sldSpeed_Change"
End Sub

Private Sub sldSpeed_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If Chr(KeyAscii) = " " Then
        KeyAscii = 0
        If cmdPlay.Enabled Then
            cmdPlay_Click
        Else
            cmdStop_Click
        End If
    End If
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".sldSpeed_KeyPress"
End Sub

Private Sub sldSpeed_Scroll()
On Error GoTo ErrSection:

    GameSpeed
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".sldSpeed_Scroll"
End Sub

Private Sub tmr_Timer()
On Error Resume Next
    
    Dim dTimeDiff#, strTemp$, idx&, i&, X&, Y&
    Dim bLocked As Boolean
    Dim Annot As cAnnotation
    Dim Ind As cIndicator
    Dim Alert As cAlert
    Dim bSavePlacement As Boolean
    
    Static bAlreadyDone As Boolean
    Static bInProgress As Boolean
    Static bScrollInprog As Boolean
    Static nPrevLeft&, nPrevTop&, nPrevWidth&, nPrevHeight&, nSnapCount&
    
    TimerStart Me.Name & ".tmr." & m.Chart.Symbol & "." & Str(hWnd)

''   JM: This line was put in to supposedly fix 5303. It did not fix 5303, but instead caused issue 5429.
''   Leave awhile for reference so we don't wonder "what happened?" then remove if all okay.
'If MouseIsPressed Then Exit Sub         '5303

If vseInvalid.Visible Then
    If tmr.Tag = "DETACH_NOW" Then
        Me.Hide         'don't exit this is a form getting reused during page load
    Else
        Exit Sub         '5631
    End If
End If

If tmr.Tag = "RESIZE NOW" Then
    Form_Resize     'this tag is set when changing page collection so chart tabs will get shown correctly
    tmr.Tag = ""    'and to bring the manage chart page dialog back to the front
    frmTemplates.ZOrder
    Exit Sub
End If

If tmr.Tag = "UNLOAD_NOW" Then
    tmr.Tag = "UNLOADING"
    Unload Me
    Exit Sub
End If

    If g.bLoadingChartPage Then Exit Sub
    If tmr.Tag = "UNLOADING" Then Exit Sub
    If bInProgress Then Exit Sub

    If tmr.Tag = "ToggleOrderbarMode" Then
        OrderBarModeToggle
        tmr.Tag = ""
        Exit Sub
    End If

    bInProgress = True
    
Select Case tmr.Tag
    Case "DETACH_NOW"
        'This tag is set in RestoreCharts during startup or page loading so that detached charts will be
        'restored after all other startup processes have been completed. This works better because the
        'call to UpdateVisibleCharts -1 in RestoreCharts causes excessive flashing with detached charts
        'due to the fact that the detached chart is not a direct MDIChild of the main window (i.e. the
        'desktop gets repainted every time a top-level window is moved, sized etc.)
        
        g.bSkipSetChartFocus = True     'set this so dirty page flag will not get set
        
        If Not g.bStarting Then         'And Not bAlreadyDone Then
            DetachChart Me
            If tmr.Tag = "UNLOAD_NOW" Then
                g.ChartGlobals.nDetached = g.ChartGlobals.nDetached - 1     '5120
            End If
        End If
        
        g.bSkipSetChartFocus = False
        bInProgress = False
        Exit Sub

    Case "MINIMIZE_NOW"
        If Not g.bStarting Then
            Me.WindowState = vbMinimized
            tmr.Tag = ""
            bInProgress = False
            Exit Sub
        End If
        
    Case "RESTORE_NOW"
        If Not g.bStarting Then
            Me.WindowState = vbNormal
            tmr.Tag = ""
            bInProgress = False
            Exit Sub
        End If
End Select
       
    If Me Is ActiveChart And Not Me Is Screen.ActiveForm Then
        If vseOrderBar.Visible And m.eOrdBarMode = eOrdBarMode_Order Then
            If vseBracketOrder.Appearance = apInset Then
                If Not TypeOf Screen.ActiveForm Is frmAsk Then
                    'user left the chart before completing the bracket order (eg - user could cancel order from trade console)
                    ClearBuySellButtons True
                End If
            End If
        End If
    End If
    
    'check if mouse is still over last highlighted toolbar button
    If Not m.oBtnMouseLast Is Nothing Then
        If m.oBtnMouseLast.ToolBarName = kTbDraw Then
            If m.oBtnMouseLast.CursorCheckClear(Me, m.aTbButtonsDraw) Then Set m.oBtnMouseLast = Nothing
        Else
            If m.oBtnMouseLast.CursorCheckClear(Me, m.aTbButtons) Then Set m.oBtnMouseLast = Nothing
        End If
    End If
    
    If Not bAlreadyDone Then
        bAlreadyDone = True
        If 1 Then
            'for new charts, do a refresh since initial painting
            'seems unpredictable!
            Me.Refresh
            If Not m.Chart.PriceAlertsChecked Then m.Chart.PriceAlertAdd
            'cmdSettings.Refresh
            m.Chart.GenerateChart eRedo1_Scrolled
        End If
    ElseIf Not m.Chart.PriceAlertsChecked Then
        If m.Chart.PriceAlertAdd > 0 Then
            m.Chart.GenerateChart eRedo1_Scrolled
        End If
    End If
        
    ' Store current size and location (since moving the form doesn't trigger the resize event)
    ' -- but only under certain conditions: if MDI is not minimized and not starting/ending ...
    If frmMain.WindowState <> vbMinimized And Not g.bStarting And Not g.bUnloading _
                And Not g.bLoadingChartPage And frmMain.tmrAutoResize.Enabled Then
        ' and the position and/or size has changed since last stored
        If nPrevLeft <> Me.Left Or nPrevTop <> Me.Top Or nPrevWidth <> Me.Width Or nPrevHeight <> Me.Height Then
            ' and if active child window is in a "normal" state (not maximized or minimized)
            If Me.WindowState = vbNormal Then
                If m.eDetachStatus = eDetached Then
                    bSavePlacement = True
                ElseIf Me.DetachStatus = eNotDetached Then
                    If Not ActiveChart Is Nothing Then
                        If ActiveChart.DetachStatus = eNotDetached Then
                            If ActiveChart.WindowState = vbNormal Then
                                bSavePlacement = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If m.eDetachStatus = eDetached And m.Chart.ShowToolbar And Me.Width <> nPrevWidth Then
        ToolbarResize2 Me, Me.pbTbBack, Me.imgTbBack, m.aTbButtons, m.bToolbarWrap
        ToolbarResize2 Me, Me.pbTbBackDraw, Me.imgTbBackDraw, m.aTbButtonsDraw, m.bToolbarWrap
        nPrevWidth = Me.Width
    End If
    
    If bSavePlacement Then
        nPrevLeft = Me.Left
        nPrevTop = Me.Top
        nPrevWidth = Me.Width
        nPrevHeight = Me.Height
        If m.eDetachStatus = eDetached Then
            nSnapCount = 0
            If Len(m.strDetachedPlacment) > 0 And (m.strDetachedPlacment <> GetFormPlacement(Me)) Then
                g.bDirtyChartPage = True
            End If
            'need different string for NormalPlacement of detached vs. non-detached windows
            'because user can move non-detached form and overwrite the NormalPlacement string
            m.strDetachedPlacment = GetFormPlacement(Me)
        ElseIf m.eDetachStatus = eNotDetached Then
            If Len(m.strNormalPlacement) > 0 And (m.strNormalPlacement <> GetFormPlacement(Me)) Then
                g.bDirtyChartPage = True
            End If
            ' store position info as fixed # twips (better for startup)
            m.strNormalPlacement = GetFormPlacement(Me)
            ' and store position info as ratio of the MDI client area
            m.strRatioPlacement = Str(RoundNum(Me.Left / frmMain.ScaleWidth, 6)) & ";" _
                & Str(RoundNum(Me.Top / frmMain.ScaleHeight, 6)) & ";" _
                & Str(RoundNum(Me.Width / frmMain.ScaleWidth, 6)) & ";" _
                & Str(RoundNum(Me.Height / frmMain.ScaleHeight, 6)) & ";" _
                & Str(Me.WindowState) & ";" & Str(CInt(Me.Visible))
            
            If g.ChartGlobals.bSnapToDots Then
                nSnapCount = nSnapCount + 1             '6782
            End If
        End If
    ElseIf nSnapCount > 0 Then
        If g.ChartGlobals.bSnapToDots Then
            i = geSnapToDots(frmMain.Picture.Handle, frmMain.hWnd, Me.hWnd)     'grapheng will return 0 or 1
            If i = 0 Then
                nSnapCount = 0
            Else
                nSnapCount = nSnapCount + i
                If nSnapCount > 25 Then
                    nSnapCount = 0
                    g.ChartGlobals.bSnapToDots = False
                    strTemp = "Snap tiled charts to background grid points failed. "
                    strTemp = strTemp & "Please reconfigure the application background "
                    strTemp = strTemp & "before turning this feature back on."
                    InfBox strTemp, "I", "Ok", "Snap tiled charts to grid", True
                End If
            End If
        Else
            nSnapCount = 0
        End If
    End If

    ' exit if hsb scroll held down
    dTimeDiff = gdTickCount - m.dLastScrollTime
    If dTimeDiff < 300 Then
        bInProgress = False
        bScrollInprog = True
        Exit Sub
    ElseIf bScrollInprog Then
        'this is to redraw text for dollar lines after user scrolled via scroll or keyboard arrows
        geAnnotMove m.Chart.geChartObj, 0
        m.Chart.GenerateChart eRedo1_Scrolled
        bScrollInprog = False
    ElseIf m.Chart.IsPartiallyLoaded Then
        ' if data is still only partially loaded, call again to load more data
        m.Chart.GenerateChart eRedo5_RecalcInd
    End If
    
    ' otherwise, set hsb step back to 1
    hsb.SmallChange = 1
        
    strTemp = tmr.Tag
    If Len(strTemp) > 0 Then
        i = Val(Parse(strTemp, " ", 2))
        If strTemp = "DETACH_NOW" Then
            tmr.Tag = ""
            If g.ChartGlobals.nDetached <= 0 And Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                ToolbarResize2 frmMain, frmMain.pbTbBack, frmMain.imgTbBack, frmMain.TbButtonsArray(kTbGeneral), frmMain.ToolBarWrapGet(kTbGeneral)
            End If
            If m.Chart.ShowToolbar Then
                ToolbarResize2 Me, pbTbBack, imgTbBack, m.aTbButtons, m.bToolbarWrap
                ToolbarResize2 Me, pbTbBackDraw, imgTbBackDraw, m.aTbButtonsDraw, m.bToolbarWrap
            End If
        ElseIf Left(strTemp, 12) = "EditSettings" Then
            frmChartOnOff.ClearPrevChart
            EditSettings i
        ElseIf Left(strTemp, 9) = "EditAnnot" Then
            'JM (12-07-2011) - known bug as of version 5.0 or earlier (did not catch it in debugger until now)
            'Intermittently the tmr.tag is set with "EditAnnot" and no annotID (i.e. blank = zero).
            'This causes the Elliot Label editor to get displayed.
            If i > 0 Then
                Set Annot = m.Chart.Annots(i)
                
                If UCase(m.Chart.Annots.Key(i)) = "SYSTEM NAME" Then
                    idx = -1            '6701
                Else
                    idx = Val(m.Chart.Annots.Key(i))
                End If
                
                If idx = -1 Or idx > 0 Then
                    EditSettings idx
                ElseIf Annot Is Nothing Then
                    'invalid object, do nothing
                ElseIf Annot.eType = eANNOT_Icon Or Annot.eType = eANNOT_ElliotLabel Then
                    'edit existing icon
                    If Annot.Prop("ImageSize") = 999999 Then
                        If Annot.AllowEWI Or Annot.AllowGMP Or (Annot.IsEndUserEWI And HasModule("EWL")) Then
                            'JM 10-10-2013: only show editor if has flag file
                            If Not frmElliot.Visible Then
                                frmElliot.ShowMe Annot      'JM (12-07-2011): this executes of the tmr.tag has blank for Annot ID
                            End If
                        End If
                    ElseIf Annot.Prop("ImageType") = eCNI_Bell_Price Or Annot.Prop("ImageType") = eCNI_Bell_Price_Gray Then
                        Set Alert = g.Alerts(Annot.Prop("AnnotKey"))
                        If Alert Is Nothing Then
                            frmAlertsSetup.ShowMe
                        Else
                            frmAlerts.ShowMe Alert, eGDAlertType_Price
                        End If
                    Else
                        If Not frmIconAnnot.Visible Then
                            frmIconAnnot.ShowMe m.Chart, Annot
                        End If
                    End If
                ElseIf Annot.eUsage = eANNOT_FibClusters Then
                    frmClusterCfg.ShowMe m.Chart
                Else
                    Select Case Annot.eType
                        Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                            eANNOT_Fibonacci4, eANNOT_FibTimeRatio, eANNOT_FibArcs, _
                            eANNOT_FibExpansion, eANNOT_FibFan, eANNOT_ElliotTimeRatio, _
                            eANNOT_FibTimeZones, eANNOT_TimeCycle, eANNOT_SpResistFan, _
                            eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
                            eANNOT_DNExpansion4, eANNOT_DNRetracement, eANNOT_Pivot, _
                            eANNOT_DanCodeFib, eANNOT_FibABCD, eANNOT_Gartley, eANNOT_DanCodeZone, _
                            eANNOT_GannacciCycle, eANNOT_GannacciTime, eANNOT_GannacciSwing1, _
                            eANNOT_GannacciSwing2, eANNOT_BalloonStrangle, eANNOT_AdvRiskReward
                                                    
                            If FormIsLoaded("frmEditAnnotExt") Then
                                frmEditAnnotExt.Hide        '5260
                                Unload frmEditAnnotExt
                            End If
                                                    
                            frmEditAnnotExt.Edit m.Chart, i
                        Case eANNOT_AndrewFork
                            If m.bSchiffFork Then
                                frmAndrewFork.ShowMe Annot
                            Else
                                If FormIsLoaded("frmEditAnnotExt") Then
                                    frmEditAnnotExt.Hide        '5260
                                    Unload frmEditAnnotExt
                                End If
                                                        
                                frmEditAnnotExt.Edit m.Chart, i
                            End If
                            
                        Case Else
                                                
                            If FormIsLoaded("frmEditAnnot") Then
                                frmEditAnnot.Hide        '5260
                                Unload frmEditAnnot
                            End If
                            
                            If Annot.eType = eANNOT_HorzLine And Annot.Prop("FakePriceAlert") = 1 Then
                                m.nActiveAnnotIdx = i
                                mnuAnnotAddPt_Click
                                m.nActiveAnnotIdx = 0
                            Else
                                frmEditAnnot.Edit m.Chart, i
                            End If
                    End Select
                End If
                Set Annot = Nothing
            End If      'JM 12-07-2011 (end check for tmr.tag being set with Annot ID of blank, which converts to zero
        ElseIf strTemp = "AddItem" Then
            If Not m.Chart.IsProfileChart Then
                ' work through frmChartCfg
                frmChartOnOff.ClearPrevChart
                If frmChartCfg.Visible Then
                    frmChartCfg.AddToChart eAdd_Study
                Else
                    frmChartCfg.ShowMe m.Chart, 0, True         'This is the "A" hotkey
                    If m.bGameMode Then
                        frmChartCfg.AddToChart eAdd_Any, 4, True
                    Else
                        frmChartCfg.AddToChart eAdd_Any, 0, True
                    End If
                End If
            End If
        ElseIf strTemp = "AddSystem" Then
            If Not m.Chart.IsProfileChart Then
                ' work through frmChartCfg
                If frmChartCfg.Visible Then
                    frmChartCfg.AddToChart eAdd_System
                Else
                    frmChartCfg.ShowMe m.Chart, 0, True
                    frmChartCfg.AddToChart eAdd_System, 0, True
                End If
            End If
        ElseIf Left(strTemp, 8) = "EditData" Then
            EditData i
        ElseIf Left(strTemp, 13) = "GenerateChart" Then
            m.Chart.GenerateChart eRedo1_Scrolled
        ElseIf InStr(tmr.Tag, "Edit Price Cluster") <> 0 Then
            'do this towards the end & use instr to give double-click event a chance to clear the tag
            'otherwise the editor often gets brought up even on a double-click
            frmClusterCfg.ShowMe m.Chart, Chart.Tree(kClusterPriceKey)
        ElseIf InStr(tmr.Tag, "Edit Time Cluster") <> 0 Then
            frmClusterCfg.ShowMe m.Chart, Chart.Tree(kClusterTimeKeyInd)
        ElseIf strTemp = "WhatIfMoved" Then
            m.Chart.GenerateChart eRedo5_RecalcInd
        ElseIf strTemp = "OrderbarReset" And vseOrderBar.Visible Then
            '03-30-2011 - this tag is set when user makes changes to exit favorites
            If m.eOrdBarMode = eOrdBarMode_Order Then FixOrderBarControls False, True
        End If
        DoEvents
        If strTemp <> "MINIMIZE_NOW" Then tmr.Tag = ""
    End If
    
    ' if this is the chart looking at the image server queues
    If m.eImgSrv = eImgSrv_Searching Then
        m.Chart.ImageServerCheck
    ElseIf m.eCmdMode = eCmdMode_Off Then
        ' if not in Command-Line mode and if Mouse is not pressed
        'If Not MouseIsPressed And m.nActiveAnnotIdx = 0 And Not frmReplay.Visible Then
        If Not MouseIsPressed And m.nActiveAnnotIdx = 0 And Not m.bGameMode Then
            If m.Chart.UpdateNewTicks Then
                DoMouseLabel True
            End If
        End If
    End If

    ' do this just in case there is still a mouse label to process
    DoMouseLabel
    If geIsCursorInWnd(pbChart.hWnd, X, Y) <> 1 Then
        ChartTips eTiptype_None     '4284
    End If
    
    'do in case no new ticks but bid/ask prices changed
    If vseOrderBar.Visible Then
        If g.nReplaySession > 0 Or frmReplay.Visible Then
            'do this immediately because the account combo dropdown will not get
            'repopulated until after user is done downloading data for a specific day
            cboAccounts.Enabled = False
        End If
        If m.eOrdBarMode = eOrdBarMode_Wizard Then
            chkConfirmOrder.Visible = False
            'If g.nOptNavStatus = eGDOptNavStatus_Loaded Then
            If m.bOptNavLoaded Then
                'aardvark 6061 - once these controls are enabled, leave them enabled since OptionsNav stays loaded
                If m.bAllowRiskGraph Then
                    If Not cboRiskGraphType.Enabled Then cboRiskGraphType.Enabled = True
                    If Not cmdRefreshGraph.Enabled Then cmdRefreshGraph.Enabled = True
                Else
                    If cboRiskGraphType.Enabled Then cboRiskGraphType.Enabled = False
                    If cmdRefreshGraph.Enabled Then cmdRefreshGraph.Enabled = False
                End If
                If Not cmdTicket.Enabled Then
                    cmdTicket.Enabled = True
                    cmdTicket.ZOrder
                End If
            End If
        Else
            chkConfirmOrder.Visible = True
            m.Chart.UpdateTradePrices
            CheckBoxValue(chkConfirmOrder) = g.Broker.ConfirmManual
            'sync auto journal flag
            If g.Broker.AutoJournalPopUp Then
                chkAutoJournal.Value = vbChecked
            Else
                chkAutoJournal.Value = vbUnchecked
            End If
        End If
        CheckOrdBarColor        '5005
    End If
    
    'show bid/ask on chart if appropriate
    If Not m.bChartMoveInProg Then m.Chart.BidAskOnChart
    
    'update countdown timer value for Woodies template
    m.Chart.UpdateWoodCountDown
    
    If Me.MDIChild Then
        If Me Is ActiveChart Then
            If frmMain.cboBarPeriod.Visible And Len(frmMain.cboBarPeriod.Text) = 0 Then
                frmMain.cboBarPeriod.Text = GetPeriodStr(m.Chart.Periodicity)   '5213
            End If
        End If
    ElseIf cboBarPeriod.Visible Then
        If Len(cboBarPeriod.Text) = 0 Then
            cboBarPeriod.Text = GetPeriodStr(m.Chart.Periodicity)       '5213
        End If
    End If
    
    CheckOpenOrderPos Me
    InitQuantityEditor
    
    UpdateSeasonalControls
        
    bInProgress = False
    
    TimerEnd Me.Name & ".tmr." & m.Chart.Symbol & "." & Str(hWnd), tmr.Interval
    
End Sub

Private Sub tmrGameMode_Timer()
On Error GoTo ErrSection:

    TimerStart Me.Name & ".tmrGameMode." & m.Chart.Symbol & "." & Str(hWnd)
    If m.oGameMode Is Nothing Then Exit Sub
    
    If Not m.oGameMode.GameTimer Then
        cmdStop_Click
    Else
        UpdateGameModeLabels
    End If
    TimerEnd Me.Name & ".tmrGameMode." & m.Chart.Symbol & "." & Str(hWnd), tmrGameMode.Interval
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".tmrGameMode_Timer"
    Resume ErrExit

End Sub

Private Sub tmrProfileLoad_Timer()
On Error Resume Next

    TimerStart Me.Name & ".tmrProfileLoad." & m.Chart.Symbol & "." & Str(hWnd)
    If Not m.Chart Is Nothing Then
        m.Chart.LoadProfileHistory
        If Not m.Chart.IsPartiallyLoaded Then m.Chart.SetFormCaption
    End If
    TimerEnd Me.Name & ".tmrProfileLoad." & m.Chart.Symbol & "." & Str(hWnd), tmrProfileLoad.Interval

End Sub

Private Sub txtCorrPercentPFP_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    If fgChartFlex(eFlexGridIdx_PfpHits).Rows > fgChartFlex(eFlexGridIdx_PfpHits).FixedRows Then
        PfpReset ePfpReset_PpfAnnotInd
        PfpReset ePfpReset_GridPfp
    End If
    
    If Not m.oPatternProfit Is Nothing Then m.oPatternProfit.PercentCorr = ValOfText(txtCorrPercentPFP.Text)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".txtCorrPercentPFP_KeyUp"

End Sub

Private Sub txtCycleNum_Change()

    If Not Me.Visible Then Exit Sub
    If m.eSeasonalCtrlsState = eSeasonCtrlStatus_Updated Then
        HandleSeasonalInput eSeasonalCtrl_CycleNum
    End If
    
End Sub

Private Sub txtForecastPFP_Change()
On Error GoTo ErrSection:
    
    If fgChartFlex(eFlexGridIdx_PfpHits).Rows > fgChartFlex(eFlexGridIdx_PfpHits).FixedRows Then
        PfpReset ePfpReset_PpfAnnotInd
        PfpReset ePfpReset_GridPfp
    End If
    
    If Not m.oPatternProfit Is Nothing Then m.oPatternProfit.NumForecastBars = ValOfText(txtForecastPFP.Text)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".txtForecastPFP_Change"

End Sub

Private Sub txtTradeQty_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTradeQty

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".txtTradeQty_GotFocus"

End Sub

Private Sub txtTradeQty_LostFocus()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".txtTradeQty_LostFocus"

End Sub

Private Sub vseBracketOrder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    vseBracketOrder.ToolTipText = "Click on chart to place a bracket(OCO) buy/sell order."

End Sub

Private Sub vseBracketOrder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    HandleBuySellClick vseBracketOrder, Button

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseBracketOrder_MouseUp"

End Sub

Private Sub vseBuyChart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    HandleBuySellClick vseBuyChart, Button

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseBuyChart_MouseUp"

End Sub

Private Sub vseBuyWizard_Click()
On Error GoTo ErrSection:

    HandleBuySellClick vseBuyWizard

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseBuyWizard_Click"

End Sub

Private Sub vseCaption_DblClick()
On Error GoTo ErrSection:

    WindowStateX(Me) = wsNormal

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseCaption_DblClick"
    Resume ErrExit
    
End Sub

Private Sub vseCaption_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    SetFocusCtl   'Peg    -MJM
    KeyPress KeyCode, Shift

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseCaption_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub vseCaption_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseCaption_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub vseDay_Click()
On Error GoTo ErrSection:

    TopMost = False
    If KeyIsPressed(VK_CONTROL) Then
        m.Chart.SetRequired
    Else
        vseMouse_Click
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseDay_Click"
    Resume ErrExit
    
End Sub

Private Sub vseDay_DblClick()
On Error GoTo ErrSection:

    TopMost = False
    'm.Chart.ShowSpeed
    
    If FileExist(g.strAppPath & "\SCP.flg") Then
        m.Chart.Unpublishable = Not m.Chart.Unpublishable
        m.Chart.GenerateChart eRedo3_Settings
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseDay_DblClick"
    Resume ErrExit
    
End Sub

Private Sub vseDay_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    SetFocusCtl   'Peg    -MJM
    KeyPress KeyCode, Shift

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseDay_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub vseDay_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseDay_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub vseDay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ' clear tips
    RefreshTips -1000, -1000

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseDay_MouseMove"
    Resume ErrExit
    
End Sub

Private Sub vseDetach_Click()
On Error GoTo ErrSection:

    DetachChart Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseDetach_Click"
    Resume ErrExit

End Sub

Private Sub vseExitA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ExitFavoriteBtnClick Me, vseExitA, Button, "A", 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseExitA_MouseUp"

End Sub

Private Sub vseExitB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ExitFavoriteBtnClick Me, vseExitB, Button, "B", 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseExitB_MouseUp"

End Sub

Private Sub vseExitC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    ExitFavoriteBtnClick Me, vseExitC, Button, "C", 2

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseExitC_MouseUp"

End Sub

Private Sub vseExitD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ExitFavoriteBtnClick Me, vseExitD, Button, "D", 3

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseExitD_Click"

End Sub

Private Sub vseInvalid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    Else
        KeyPress KeyCode, Shift
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseInvalid_KeyDown"
    Resume ErrExit
End Sub

Private Sub vseInvalid_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseInvalid_KeyPress"
    Resume ErrExit
End Sub

Private Sub vseInvalid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        ShowPopup
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseInvalid_MouseDown"
    Resume ErrExit
End Sub

Private Sub vseMouse_Click()
On Error GoTo ErrSection:
    
    'Exit Sub
    'vseMouse.Caption = tmr.Enabled & ", " & tmr.Interval
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseMouse_Click"
    Resume ErrExit
    
End Sub

Private Sub vseMouse_DblClick()
On Error GoTo ErrSection:

    'vseMouse.Caption = str(Peg.Left) & str(Peg.Width) & str(Me.ScaleWidth)
    'If m.Chart.Zoomed Then
    '    Peg.PEactions = 19
    'Else
        m.Chart.ShowSpeed
    'End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseMouse_DblClick"
    Resume ErrExit
    
End Sub

Private Sub vseMouse_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    SetFocusCtl   'Peg    -MJM
    KeyPress KeyCode, Shift

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseMouse_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub vseMouse_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseMouse_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub vseMouse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ' clear tips
    RefreshTips -1000, -1000
        'MJM - Q&A: need to duplicate this?
    'Peg_MouseMove Button, Shift, X + vseMouse.Left - Peg.Left, Y + vseMouse.Top - Peg.Top

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseMouse_MouseMove"
    Resume ErrExit
    
End Sub

Private Sub EditSettings(Optional ByVal idxSetting& = 0)
On Error GoTo ErrSection:

    If m.bEditing Then Exit Sub
    m.bEditing = True
    
    TopMost = False
    'DoEvents
    
    If m.Chart.IsProfileChart Then
        frmNewChart.ShowMe m.Chart.Symbol, False, m.Chart, True
    Else
        frmChartCfg.ShowMe m.Chart, idxSetting
    End If
    RefreshTips -1000, -1000
    DoEvents
    tmr.Tag = ""
    m.bEditing = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".EditSettings", eGDRaiseError_Raise
    
End Sub

Public Sub DoMouseLabel(Optional ByVal bForceRedisplay As Boolean = False)
On Error GoTo ErrSection:
        
    Dim dDate#, nBar&, nPaneID&, nX&, dDiffTime#
    Dim strText$, strTag$, strMsg$
    Static dPrevTime#, strPrevTag$, nPrevPaneID&, nPrevColorTheme&
    
    ' skip (for now) if mouse is moving really fast
    dDiffTime = gdTickCount - dPrevTime
    If dDiffTime < 90 Then Exit Sub
        
    If g.nColorTheme <> nPrevColorTheme Then
        nPrevColorTheme = g.nColorTheme
        If g.nColorTheme = vbWhite Then
            vseMouse.ForeColor = 0
            vseMouse.Font.Bold = -1 * g.ChartGlobals.nFontStyle      'Glen does not want this bold by default
        ElseIf g.nColorTheme = kDarkThemeColor Then
            vseMouse.ForeColor = vbWhite
            vseMouse.Font.Bold = -1 * g.ChartGlobals.nFontStyle
        Else
            vseMouse.ForeColor = g.ChartGlobals.nBorderForeColor
            vseMouse.Font.Bold = True
        End If
        vseDay.Font.Bold = False '-1 * g.ChartGlobals.nFontStyle
        vseDay.ForeColor = vseMouse.ForeColor
    End If

strMsg = Trim(vseDay.Tag)
If Len(strMsg) > 0 Then
    If vseDay.Caption <> strMsg Then
        vseDay.Caption = strMsg
        vseDay.Width = Me.ScaleWidth
        vseDay.Font.Bold = True
    End If
ElseIf vseDay.Width <> vseMouse.Left Then
    vseDay.Width = vseMouse.Left
    vseDay.Caption = ""
    vseDay.Font.Bold = False
End If
        
    ' if something to do, it will be in the Tag
    strTag = Trim(vseMouse.Tag)
    If Not bForceRedisplay Then
        If Len(strTag) = 0 Then Exit Sub
        ' skip (for now) if same tag in last second
        If strTag = strPrevTag And dDiffTime < 1000 Then Exit Sub
    End If
    
    ' otherwise, do it now ...
    vseMouse.Tag = ""
    If Len(strTag) = 0 Then
        strTag = strPrevTag
    Else
        strPrevTag = strTag
    End If
    
    ' parse stuff from Tag
    nPaneID = Val(Parse(strTag, Chr(9), 1))
    If nPaneID = -99 Then
        nPaneID = nPrevPaneID
    Else
        nPrevPaneID = nPaneID
    End If
    nX = Val(Parse(strTag, Chr(9), 2))
    If nX < 0 Then
        If Len(strMsg) = 0 Then vseDay.Caption = ""
        vseMouse.Caption = ""
        Exit Sub
        'nX = m.Chart.LastGoodDataBar(True)
    End If
    
    ''nBar = m.Chart.aXbar(nX) 'could be < 0
       
    strText = m.Chart.GetMouseLabel(nX, nPaneID)
    If vseMouse.Caption <> strText Then
        vseMouse.Caption = Parse(strText, vbTab, 2)
        If Len(strMsg) = 0 Then vseDay.Caption = Parse(strText, vbTab, 1)
        'vseMouse.Refresh
    End If
    
    'If g.ChartGlobals.bAutoChartData And g.ChartGlobals.bChartDataSingleBar Then
    If g.ChartGlobals.bAutoChartData Then
        If Me Is ActiveChart Then
            dDate = m.Chart.aXdate(nX)
            frmChartData.ShowData nX
            frmPlanetData.ShowData dDate, m.Chart.Bars
        End If
    End If
        
    dPrevTime = gdTickCount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".DoMouseLabel", eGDRaiseError_Raise
   
End Sub

Public Property Get Chart() As cChart
On Error GoTo ErrSection:
    
    If m.Chart Is Nothing Then
        Set m.Chart = New cChart
        Set m.Chart.Form = Me
    End If
    Set Chart = m.Chart

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".Chart.Get", eGDRaiseError_Raise
    
End Property

Private Sub ShowPopup()
On Error GoTo ErrSection:

    Dim i&, iNondetachCount&, s$
    Dim aTemplates As cGdArray
    Dim Pane As cPane
    Dim Annot As cAnnotation
    
    TopMost = False
        
    'TLB 2/15/2008: for now, hide "Get additional data" for intraday charts
    'TLB 3/22/2010: show this menu item only if this is a daily chart and they have not installed full daily history
    i = False
    If Not m.Chart.Bars.IsIntraday Then
        Select Case SecurityType(m.Chart.Bars)
        Case "F" ' see if < 250 megs
            If FileLength(DataPath & "FUT_EOD.FPT") < 250000000 Then
                i = True
            End If
        Case "S" ' see if < 250 megs
            If FileLength(DataPath & "STK_EOD.FPT") < 250000000 Then
                i = True
            End If
        Case "I" ' see if < 25 megs
            If FileLength(DataPath & "IDX_EOD.FPT") < 25000000 Then
                i = True
            End If
        End Select
    End If
    mnuMoreDataSingleSym.Visible = i
    
'    If m.bAllowDetach Then
        If Me.MDIChild Then
            iNondetachCount = NonDetachCount
            mnuDetAttChart.Caption = "Detach chart"
            mnuShowToolbar.Visible = False
            If iNondetachCount <= 1 And g.ChartGlobals.frmActiveNonDetached Is Me And m.eDetachStatus <> eNotDetached Then
                mnuDetAttChart.Enabled = False          'work-around for weird issue at startup
                m.eDetachStatus = eNotDetached
'            ElseIf iNondetachCount > 1 Then            '5208
'                mnuDetAttChart.Enabled = True
            Else
                mnuDetAttChart.Enabled = True
            End If
        Else
            mnuDetAttChart.Caption = "Attach chart"
            mnuDetAttChart.Enabled = True
            mnuShowToolbar.Visible = True
            mnuShowToolbar.Checked = m.Chart.ShowToolbar
        End If
'    Else
'        mnuDetAttChart.Visible = False
'        mnuDetAttChart.Enabled = False
'        mnuShowToolbar.Visible = False
'    End If
        
    mnuUnzoom.Enabled = m.Chart.Zoomed
    
    mnuDisableRT.Visible = g.RealTime.Active
    mnuDisableRT.Checked = m.Chart.DisableRT
    
    'store current Pane
    If m.MouseLast.bOffChart Or m.Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then
        mnuHidePane.Enabled = False
        mnuHidePane.Tag = ""
    Else
        s = m.Chart.Tree.Key(m.MouseLast.nPaneID)
        mnuHidePane.Tag = s
        mnuHidePane.Enabled = True
    End If
    
    ' system
    If Not HasGold(False, , False) Or m.Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then
        mnuSep5.Visible = False
        mnuEditSystem.Visible = False
        mnuTrades.Visible = False
    Else 'If m.Chart.SystemID > 0 Then
        mnuSep5.Visible = True
        mnuEditSystem.Visible = True
        mnuTrades.Visible = True
        mnuTrades.Enabled = True
        If m.Chart.ShowTrades = 1 Then        'ShowTrades is a long: 0=none,1=strategy,2=manual trade account
            mnuTrades.Checked = vbChecked
            mnuEditSystem.Enabled = True
        Else
            mnuTrades.Checked = vbUnchecked
            mnuEditSystem.Enabled = False
        End If
    End If
    
    ' auto-trade
    i = False
    If HasPlatinum(False) And m.Chart.ShowTrades = 1 Then
        If FileExist(App.Path & "\Auto.trd") Then
            i = True
            mnuAutoTrade.Checked = m.Chart.AutoTrade
        End If
    End If
    mnuAutoTrade.Visible = i
        
    'Templates
    If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
        mnuTemplates.Visible = True
        Set aTemplates = GetAllowedList("T")
        If aTemplates.Size = 0 Then aTemplates.Add "(no templates found)"
        If Not aTemplates Is Nothing Then
            'add menu item for each template
            For i = 0 To aTemplates.Size - 1
                If i > mnuTemplate.UBound Then
                    Load mnuTemplate(i)
                    mnuTemplate(i).Visible = True
                End If
                mnuTemplate(i).Caption = Parse(aTemplates(i), vbTab, 1)
            Next
            'remove extras
            For i = mnuTemplate.UBound To aTemplates.Size Step -1
                If i > 0 Then Unload mnuTemplate(i)
            Next
            ' check if more than this chart exists
            mnuTemplateToOtherCharts.Visible = False
            For i = 0 To Forms.Count - 1
                If IsFrmChart(Forms(i)) Then
                    If Not Forms(i) Is Me Then
                        mnuTemplateToOtherCharts.Visible = True
                        Exit For
                    End If
                End If
            Next
        End If
    Else
        mnuTemplates.Visible = False
    End If
    'aTemplates.Add " < manage chart templates >", 0
    
    For i = 0 To mnuPPB.Count - 1
        If m.Chart.PixelsPerBar = ValOfText(mnuPPB(i).Caption) Then
            mnuPPB(i).Checked = True
        Else
            mnuPPB(i).Checked = False
        End If
    Next
    
    If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
        mnuBarPeriods.Visible = True
        s = UCase(m.Chart.Bars.Prop(eBARS_PeriodicityStr))
        For i = 0 To mnuBarPeriod.Count - 1
            If s = Left(StripStr(UCase(mnuBarPeriod(i).Caption), "&"), Len(s)) Then
                mnuBarPeriod(i).Checked = True
                s = "*FOUND*"
            Else
                mnuBarPeriod(i).Checked = False
            End If
        Next
        If s <> "*FOUND*" Then mnuBarPeriod(mnuBarPeriod.Count - 1).Checked = True
    Else
        mnuBarPeriods.Visible = False
    End If
    
    mnuCrosshairs.Checked = frmMain.tbToolbar.Tools("ID_Crosshair").State
    mnuTips.Checked = g.ChartGlobals.bFloatingTips
    'mnuLockValues.Checked = m.bLockValuesDisplay
    mnuHideScrollbars.Checked = g.ChartGlobals.bHideScrollbars
    mnuUnsplit.Checked = m.Chart.Unsplit
    If m.Chart.VertGrid = 0 Then
        mnuCoarseGrid.Checked = True
    Else
        mnuCoarseGrid.Checked = False
    End If
    
    mnuAutoScale.Checked = m.Chart.AutoScale
    If m.Chart.ChartLogFlag = ePANE_LogFlagLog Then
        mnuLogScale.Checked = vbChecked         '0=linear, -1=log, 1=percent change
        mnuLogModeDraw.Visible = True
        If g.ChartGlobals.bLogModeDraw Then
            mnuLogModeDraw.Checked = vbChecked
        Else
            mnuLogModeDraw.Checked = vbUnchecked
        End If
    Else
        mnuLogScale.Checked = vbUnchecked
        mnuLogModeDraw.Visible = False
    End If
    
    If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
        mnuLogScale.Visible = True
        mnuUnsplit.Visible = True
    
        mnuSep9.Visible = True
        mnuOrderBar.Visible = True
        mnuAccountBar.Visible = True
        mnuOrdBarSettings.Visible = True
        mnuOrdBarDefaults.Visible = True
    Else
        mnuLogScale.Visible = False
        mnuUnsplit.Visible = False
        
        mnuSep9.Visible = False
        mnuOrderBar.Visible = False
        mnuAccountBar.Visible = False
        mnuOrdBarSettings.Visible = False
        mnuOrdBarDefaults.Visible = False
    End If
    
    If Len(m.Chart.SpreadSymbols) > 0 Or m.Chart.IsInWhatIfMode Or _
        m.Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then

'JM 03-16-2015: we are streaming replay a limited number of stocks
'               allow order bar for all and if it is not in list then nothing happens
'        (m.Chart.Bars.SecurityType = "S" And g.nReplaySession > 0) Then
        
        mnuOrderBar.Enabled = False
        mnuAccountBar.Enabled = False
        mnuOrdBarSettings.Enabled = False
        mnuOrdBarDefaults.Enabled = False
    Else
        mnuOrderBar.Enabled = True
        mnuAccountBar.Enabled = True
        mnuOrdBarSettings.Enabled = True
        mnuOrdBarDefaults.Enabled = True
        If m.Chart.ShowTrades = 2 Then
            mnuOrderBar.Checked = True
        Else
            mnuOrderBar.Checked = False
        End If
        mnuAccountBar.Checked = m.Chart.ShowAccountBar
    End If
    
    mnuCursorArrow.Checked = False
    mnuCrosshairs.Checked = False
    mnuCursorVert.Checked = False
    mnuCursorHoriz.Checked = False
    With frmMain.tbToolbar
        If .Tools("ID_CursorCrosshairs").State = ssChecked Then
            mnuCrosshairs.Checked = True
        ElseIf .Tools("ID_CursorHorizLine").State = ssChecked Then
            mnuCursorHoriz.Checked = True
        ElseIf .Tools("ID_CursorVertLine").State = ssChecked Then
            mnuCursorVert.Checked = True
        Else
            mnuCursorArrow.Checked = True
        End If
    End With
    
    If mnuDeleteAnnot.UBound = 0 Then
        Load mnuDeleteAnnot(1)
        mnuDeleteAnnot(1).Caption = "Delete Annotations Prior to Date"
        Load mnuDeleteAnnot(2)
        mnuDeleteAnnot(2).Caption = "Delete All Annotations"
    End If
    
    Set Annot = m.Chart.Annots(1)
    If Not Annot Is Nothing Then
        'add Show/Hide EWI menu item if chart has EWI labels but no flag files
        If Not Annot.AllowEWI And Not Annot.AllowGMP Then
            i = m.Chart.HasHiddenAnnots(eANNOT_ElliotLabel)
            If -1 <> i Then
                If mnuHideAnnots.UBound = 0 Then
                    Load mnuHideAnnots(1)
                    mnuHideAnnots(1).Visible = True
                End If
                If i = 0 Then
                    mnuHideAnnots(1).Caption = "Hide EWI"
                Else
                    mnuHideAnnots(1).Caption = "Show EWI"
                End If
            End If
        End If
    End If
    
    If g.ChartGlobals.nHideAnnotations = 0 Then
        mnuHideAnnots(0).Checked = False
    Else
        mnuHideAnnots(0).Checked = True
    End If
    
    If mnuHideAnnots.UBound = 1 Then
        mnuHideAnnots(1).Checked = False
    End If
    
    If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
        mnuBarDisplay.Visible = True
        mnuOHLC.Checked = False
        mnuCandlesticks.Checked = False
        mnuCloseLine.Checked = False
        Select Case m.Chart.BarDisplayType
        Case eINDIC_Candlestick
            mnuCandlesticks.Checked = True
        Case eINDIC_Line
            mnuCloseLine.Checked = True
        Case eINDIC_BollingerBar
            mnuBollinger.Checked = True
        Case Else
            mnuOHLC.Checked = True
        End Select
    Else
        mnuBarDisplay.Visible = False
    End If
    
    If m.eDetachStatus = eDetached Then
        mnuScreenCapture.Enabled = False
    Else
        mnuScreenCapture.Enabled = True
    End If
    
    mnuDataCopy.Visible = True
        
    Me.PopupMenu mnuPopUp

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ShowPopup", eGDRaiseError_Raise
    
End Sub

Private Sub ShowAnnotPopup(ByVal nExtendable As Long, ByVal nButton&)
On Error GoTo ErrSection:

    Dim i&, s$
    Dim aTemplates As cGdArray
    Dim Pane As cPane
    Dim Annot As cAnnotation
    Dim bShowAddPtMenu As Boolean
        
    TopMost = False
        
    'set flag so if user decides to do nothing editor will not come up
    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    
    If Annot Is Nothing Then Exit Sub
    If Annot.eType = eANNOT_Rectangle And Annot.eUsage = eANNOT_PatternProfit Then Exit Sub
    
    mnuAnnotDuplicate.Visible = Annot.DuplicateAllow
    
    If Annot.eType = eANNOT_Bracket Then
        mnuAnnotStyle.Visible = True
    Else
        mnuAnnotStyle.Visible = False
    End If
        
    If Annot.IsFibType And nButton = vbRightButton Then
        mnuAnnotMovePt.Caption = "Edit"
        mnuAnnotDeletePt.Caption = "Delete"
        mnuAnnotAddPt.Caption = "Delete line"
        
        If Annot.HitItemIndex > 5 Then bShowAddPtMenu = True
        
        Select Case Annot.eType
            Case eANNOT_ElliotTimeRatio
                If Annot.HitItemIndex > 4 Then bShowAddPtMenu = True
            Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
                 eANNOT_DNExpansion4, eANNOT_FibABCD
                If Annot.HitItemIndex > 2 Then bShowAddPtMenu = True     '5602
            Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                 eANNOT_Fibonacci4, eANNOT_FibExpansion, eANNOT_DanCodeFib, _
                 eANNOT_AdvRiskReward
                If Annot.FibLineBold Then
                    mnuAnnotStyle.Caption = "Unbold line"
                Else
                    mnuAnnotStyle.Caption = "Bold line"
                End If
                mnuAnnotStyle.Visible = True
        End Select
        
        mnuAnnotAddPt.Visible = bShowAddPtMenu
        
''JM 12-01-2010: uncomment when ready to implement eraser tool for new andrews fork
''        ElseIf Annot.eType = eANNOT_AndrewFork And Annot.HitItemIndex > 2 Then
''            mnuAnnotAddPt.Visible = True
'
'        ElseIf Annot.HitItemIndex > 5 Then
'            mnuAnnotAddPt.Visible = True
'        Else
'            mnuAnnotAddPt.Visible = False
'        End If
        
        Me.PopupMenu mnuAnnotEdit
    ElseIf Annot.eUsage = eANNOT_FibClusters Then
        'do nothing
    ElseIf Annot.eType <> eANNOT_BellAlert And Annot.eUsage <> eANNOT_Trades And _
           Annot.eUsage <> eANNOT_WhatIf And Annot.eUsage <> eANNOT_OptionInfo Then
        If nButton = vbRightButton Then
            If Annot.eType = eANNOT_ElliotLabel And Not Annot.DuplicateAllow Then
                If Annot.IsEndUserEWI Then
                    'annot created by end user with EWL enablement, but being loaded by someone who does not have "EWL"
                    mnuAnnotMovePt.Caption = "Delete"
                    mnuAnnotAddPt.Visible = False
                    mnuAnnotDeletePt.Visible = False
                Else
                    'annot came from analyst
                    mnuAnnotMovePt.Caption = "Hide EWI"
                    mnuAnnotAddPt.Visible = False
                    mnuAnnotDeletePt.Visible = False
                End If
            ElseIf Annot.eType = eANNOT_BalloonStrangle Then
                mnuAnnotMovePt.Caption = "Edit"
                mnuAnnotDeletePt.Caption = "Delete"
                If Val(Annot.Prop("ShowCollapsed")) = 1 Then
                    mnuAnnotAddPt.Caption = "Expand"
                Else
                    mnuAnnotAddPt.Caption = "Collapse"
                End If
            Else
                mnuAnnotMovePt.Caption = "Edit"
                mnuAnnotDeletePt.Caption = "Delete"
                mnuAnnotDeletePt.Visible = True             '6945
                If Annot.eType = eANNOT_Bracket Then
                    mnuAnnotAddPt.Caption = "Change direction"
                Else
                    mnuAnnotAddPt.Caption = "Alert"
                End If
                
                If Annot.eType = eANNOT_HorzLine And Annot.Prop("FakePriceAlert") = 1 Then
                    mnuAnnotAddPt.Visible = False
                ElseIf Annot.CanHaveAlert Or Annot.eType = eANNOT_Bracket Or Annot.eType = eANNOT_TextEdit Or _
                   Annot.eType = eANNOT_TextEdit2 Or Annot.eType = eANNOT_TextEdit3 Or _
                   Annot.eType = eANNOT_TextEdit4 Then
                    mnuAnnotAddPt.Visible = True
                Else
                    mnuAnnotAddPt.Visible = False
                End If
            End If
        Else
            Annot.MenuMove = True
            'menu items for wave labels
            mnuAnnotMovePt.Caption = "Move point"
            mnuAnnotDeletePt.Caption = "Delete point"
            mnuAnnotAddPt.Caption = "Add point"
            
            mnuAnnotMovePt.Visible = True
            If Annot.TotalPoints > 2 Then
                mnuAnnotDeletePt.Visible = True
            Else
                mnuAnnotDeletePt.Visible = False
            End If
            If nExtendable > 1 Then
                mnuAnnotAddPt.Visible = True
            Else
                mnuAnnotAddPt.Visible = False
            End If
        End If
        Me.PopupMenu mnuAnnotEdit
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ShowAnnotPopup", eGDRaiseError_Raise
    
End Sub

Public Sub KeyPress(KeyAscii As Integer, Optional ByVal Shift As Integer = -1)
'On Error Resume Next
On Error GoTo ErrSection:

    Dim strCmd$, nRec&, strCaption$, i&, d#, bUnhandled As Boolean
    Dim aStrings As cGdArray
    Dim Ind As cIndicator
    Dim Annot As cAnnotation
    Dim Pane As cPane
    Dim bSymChangeOk As Boolean
        
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "UNLOADING" Then Exit Sub
    If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then Exit Sub
    
    If m.Chart Is Nothing Then Exit Sub
    If UCase(tmr.Tag) = "EDITSETTINGS" Then Exit Sub
    If m.bGameMode Then cmdStop_Click
        
    ' don't want to do any of this while zooming or drawing a trendline!
    If MouseIsPressed Then Exit Sub
    
    'aardvark 5627: if Ctrl-Key or Shift-Key is released at the exact same time as the mouse while
    '   moving/extending an annotation, such as $Line, Trendline, Arrow Text etc.
    '   this function gets invoked before the MouseUp event causing the trendline to get deleted
    If m.nActiveAnnotIdx > 0 Then
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Not Annot Is Nothing Then
            If Annot.ShiftKey <> 0 Then
                Set Annot = Nothing
                Exit Sub
            End If
        End If
    End If
    
    If KeyAscii = 27 And Len(g.strActiveDraw) > 0 Then m.bDrawToolJustCleared = True    'do not make this a local variable (timing issue)
    
    If KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight And _
        KeyAscii <> vbKeyHome And KeyAscii <> vbKeyEnd And _
        KeyAscii <> vbKeyPageUp And KeyAscii <> vbKeyPageDown Then
        
        'drawing tools in progress are getting cleared here (not in esc key code below)
        strCmd = UCase(Chr(KeyAscii))
        If strCmd = "T" Then
            ClearAnnotFlags True, False
        Else
            ClearAnnotFlags True
        End If
        
    End If
    
    If Screen.ActiveControl Is hsb Then
        'move focus so "active" rectangle on scrollbar
        'won't look so goofy (since not resized)
        SetFocusCtl
    End If
        
    TopMost = False
    
    ' See if currently in "Command Mode"
    If m.eCmdMode <> eCmdMode_Off Then
        ' make sure came from KeyPress event
        If Shift < 0 Then
            'parse current command from caption
            strCaption = vseCaption.Caption
            i = InStr(strCaption, ":  ")
            If i > 0 Then strCmd = Mid(strCaption, i + 3)
            
            Select Case KeyAscii
            Case 27 'Escape
                'cancelling CmdMode, so restore caption
                strCaption = "" 'strOriginalCaption
                m.eCmdMode = eCmdMode_Off
            Case 8 'Backspace
                If Len(strCmd) > 0 Then strCmd = Left(strCmd, Len(strCmd) - 1)
            Case 32 To 122
                'add the character typed
                strCmd = strCmd & UCase(Chr(KeyAscii))
            
            Case 13 'Enter
                'process command now
                strCmd = Trim(strCmd)
                If Len(strCmd) > 0 Then
                    Select Case m.eCmdMode
                
                    Case eCmdMode_Symbol
                        'load new symbol
                        ' go to symbol
                        If Right(strCmd, 1) = "-" Then strCmd = strCmd & "067"
                        nRec = g.SymbolPool.PoolRecForSymbol(strCmd, True)
                        If nRec >= 0 Then
                            m.Chart.SetSymbol g.SymbolPool.SymbolID(nRec), True
                        Else
                            Beep
                        End If
                
                    Case eCmdMode_BarPeriod
                        'different bar time period
                        m.Chart.ChangeBarPeriod strCmd
                    End Select
                End If
                'then turn CmdMode off
                strCaption = ""
                m.eCmdMode = eCmdMode_Off
                
            Case Else
                bUnhandled = True
            End Select
        End If
    
    ' See if came from KeyDown event (rather than from KeyPress)
    ElseIf Shift = 0 Then
        i = -999999
        Select Case KeyAscii
        Case vbKeyF11
            Set Pane = m.Chart.Tree("PRICE PANE")
            If Not Pane Is Nothing Then
                Pane.geIncDecMaxRatio 0.03
                Pane.geIncDecMinRatio -0.03
                If Pane.Scaling <> ePANE_ScaleModeManual Then
                    Pane.Scaling = ePANE_ScaleModeManual
                    m.Chart.SyncToolbar
                End If
                Set Pane = Nothing
                geAnnotMove m.Chart.geChartObj, 2
            End If
            m.Chart.GenerateChart eRedo1_Scrolled
            geAnnotMove m.Chart.geChartObj, 0
        Case vbKeyF12
            Set Pane = m.Chart.Tree("PRICE PANE")
            If Not Pane Is Nothing Then
                Pane.geIncDecMaxRatio -0.03
                Pane.geIncDecMinRatio 0.03
                If Pane.Scaling <> ePANE_ScaleModeManual Then
                    Pane.Scaling = ePANE_ScaleModeManual
                    m.Chart.SyncToolbar
                End If
                Set Pane = Nothing
                geAnnotMove m.Chart.geChartObj, 2
            End If
            m.Chart.GenerateChart eRedo1_Scrolled
            geAnnotMove m.Chart.geChartObj, 0
        
        Case vbKeyF2 To vbKeyF9 'Function keys
            mnuPPB_Click 7 - (KeyAscii - vbKeyF2)
        
        Case vbKeyDelete ' delete last user annot
            If m.Chart.RemoveAnnots(True) > 0 Then
                m.Chart.SyncGlobalAnnots Nothing, True
            Else
                Beep
            End If
            
        Case vbKeyUp, vbKeyDown
            'TLB: let's not do this anymore (with new QB style, too often
            'use arrows to move on QB and inadvertantly changes symbol on chart)
            If 0 Then
                With frmSymbolGrid.fgVirtual
                    nRec = -1
                    If KeyAscii = vbKeyUp Then
                        i = .Row - 1
                    Else
                        i = .Row + 1
                    End If
                    If i >= .FixedRows And i < .Rows Then
                        strCmd = Trim(.TextMatrix(i, kSymbolCol))
                        nRec = g.SymbolPool.PoolRecForSymbol(strCmd, True)
                        If nRec >= 0 Then
                            m.Chart.SetSymbol g.SymbolPool.SymbolID(nRec), True
                        End If
                    End If
                    If nRec < 0 Then Beep
                End With
            End If
            i = -999999
        
        'handle navigation keys as if scroll bar was active
        Case vbKeyLeft
            i = hsb.Value - hsb.SmallChange
        Case vbKeyRight
            i = hsb.Value + hsb.SmallChange
        Case vbKeyHome
            i = hsb.Min
        Case vbKeyEnd
            m.Chart.RestoreChartNormal vbKeyEnd
        Case vbKeyPageUp
            i = hsb.Value + hsb.LargeChange
        Case vbKeyPageDown
            i = hsb.Value - hsb.LargeChange
        
        Case Else
            bUnhandled = True
        End Select
    
        If i > -999999 Then
            If i > hsb.Max Then
                i = hsb.Max
            ElseIf i < hsb.Min Then
                i = hsb.Min
            End If
            hsb.Value = i
        End If
        
    ElseIf Shift = 2 Then ' if "Ctrl" is pressed
        Select Case KeyAscii
        Case vbKeyPageUp
            ' Can't load a new chart page from here since
            ' the new chart page may throw this form away
            ' which leaves the code nowhere to return to!
            frmMain.tmrMain.Tag = "LoadChartPage -"
            'LoadChartPage "-"
        Case vbKeyPageDown
            frmMain.tmrMain.Tag = "LoadChartPage +"
            'LoadChartPage "+"
        Case vbKeyLeft, vbKeyRight
            ' move left/right a chart tab
            If WindowState = 2 And vsTab.NumTabs > 1 Then
                If KeyAscii = vbKeyLeft Then
                    i = vsTab.CurrTab - 1
                Else
                    i = vsTab.CurrTab + 1
                End If
                If i < 0 Then
                    i = vsTab.NumTabs - 1
                ElseIf i >= vsTab.NumTabs Then
                    i = 0
                End If
                vsTab.CurrTab = i
            End If
        Case Asc("Z")
            If Not m.Chart Is Nothing Then m.Chart.LastEditedAnnotUndo
        Case Else
            bUnhandled = True
        End Select
    
    ElseIf Shift = 7 Then ' if Shift-Ctrl-Alt pressed
        Select Case KeyAscii
        Case vbKeyUp, vbKeyDown
            ' TLB: up/down arrows used for debugging (to act like realtime)
            If FileExist("c:\common\files.exe") Then
                i = m.Chart.LastGoodDataBar(False)
                d = m.Chart.Bars(eBARS_Close, i)
                If KeyAscii = vbKeyUp Then
                    d = d + m.Chart.Bars.MinMove
                Else
                    d = d - m.Chart.Bars.MinMove
                End If
                m.Chart.Bars(eBARS_Close, i) = d
                If m.Chart.Bars(eBARS_High, i) < d Then
                    m.Chart.Bars(eBARS_High, i) = d
                End If
                If m.Chart.Bars(eBARS_Low, i) > d Then
                    m.Chart.Bars(eBARS_Low, i) = d
                End If
                m.Chart.GenerateChart eRedo5_RecalcInd, True
            End If
        
        Case vbKeyLeft, vbKeyRight
            If m.Chart.ToEndOfData = False Then
                Do
                    If KeyAscii = vbKeyLeft Then
                        m.Chart.ToDate = m.Chart.ToDate - 1
                    Else
                        m.Chart.ToDate = m.Chart.ToDate + 1
                    End If
                Loop While IsWeekday(m.Chart.ToDate) = False
                m.Chart.GenerateChart eRedo9_ReloadData
            End If
        
        'Case vbKeyReturn
        '    m.Chart.VerifyResults
            
        Case 191 ' "?"
            g.bShowRecalcMsg = Not g.bShowRecalcMsg
            
        'Case Asc("N")
        '    m.Chart.UseNewRecalcMethod = True
        'Case Asc("O")
        '    m.Chart.UseNewRecalcMethod = False
        
        Case Asc("B")
            If FileExist("c:\Genesis\ChartData.GDB") Then
                m.Chart.Bars.ToFile "GDB", "c:\Genesis\ChartData.GDB"
            Else
                m.Chart.Bars.ToFile "GDB", App.Path & "\..\ChartData.GDB"
            End If
            
        Case Asc("Z")
            m.Chart.TestFZ
        
        Case Else
            bUnhandled = True
        End Select
    
    ElseIf Shift < 0 Then 'Came from KeyPress event
        strCmd = UCase(Chr(KeyAscii))
        Select Case strCmd
        Case Chr(27) 'Esc key - Clear (restore cursor, unzoom, etc.)
            If m.bDrawToolJustCleared Then
                m.bDrawToolJustCleared = False
            Else
            'JM: 06-05-2008: restored original code & added mod for
            '   - make "Esc" the hot-key for Auto-scale mode
                With m.Chart
                    .IsPartiallyLoaded = False
                    If .IsInWhatIfMode Then
                        .DeactivateWhatIf
                        .SyncToolbar
                    End If
                    If Not m.AnnotOptions Is Nothing Then
                        If m.AnnotOptions.eType = eANNOT_BalloonStrangle Then       '7021
                            ClearBuySellButtons
                            tmr.Tag = ""
                            .RemoveAnnots False, , eANNOT_OptionInfo, False
                        End If
                    End If
                    ToolbarSetCursorGroup frmMain.tbToolbar, False
                    For i = 0 To .Tree.Count
                        If .Tree.NodeLevel(i) = 0 Then
                            Set Pane = .Tree(i)
                            If Not Pane Is Nothing Then
                                Pane.geClearMaxRatio
                                Pane.geClearMinRatio
                                Pane.geIncDecMaxRatio 0.03 ' 0.05
                                Pane.geIncDecMinRatio -0.03 ' -0.05
                            End If
                        End If
                    Next
                    Set Pane = .Tree(.Tree.Key("PRICE PANE"))
                    If Not Pane Is Nothing Then
                        If Pane.Scaling = ePANE_ScaleModeManual Then
                            Pane.Scaling = Pane.ScaleAutoLastUsed
                            'm.Chart.SyncToolbar
                        End If
                    End If
                    Set Pane = Nothing
                    .ExtraPriceScale = 0
                    If .Zoomed Then
                        .UnzoomChart True
                    Else
                        .GenerateChart eRedo1_Scrolled
                    End If
                    
                    .SyncToolbar
                End With
            End If
        
        Case "."
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                If frmPlanetData.Visible Then
                    frmPlanetData.SetBase
                End If
            End If
            
        Case " " 'Space: toggle flags on symbol grid
'            If frmReplay.Visible Then
'                If frmReplay.ReplayMode = 2 Then
'                    frmReplay.StopPlay
'                ElseIf frmReplay.ReplayMode = 1 Then
'                    frmReplay.StartPlay
'                End If
            If m.bGameMode Then
                If m.oGameMode.GameReplayMode(tmrGameMode.Enabled) = eGDReplayMode_Play Then
                    cmdStop_Click
                ElseIf m.oGameMode.GameReplayMode(tmrGameMode.Enabled) = eGDReplayMode_Pause Then
                    cmdPlay_Click
                End If
            ElseIf DockState(frmSymbolGrid) <> eHidden Then
                With frmSymbolGrid
                    .ShowSymbol m.Chart.Symbol
                    .KeyPress KeyAscii
                    If .fgVirtual.Row > 0 Then .fgVirtual.ShowCell .fgVirtual.Row, 0
                End With
            Else
                Beep
            End If
                    
        Case "+", "=" 'More bars
            m.Chart.PixelsPerBar = -2
            m.Chart.GenerateChart eRedo1_Scrolled
            
        Case "-" 'Less bars
            m.Chart.PixelsPerBar = -1
            m.Chart.GenerateChart eRedo1_Scrolled
                
        Case "0" To "9" 'Templates
            ' change Templates
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                If Not m.Chart.TemplateApply(strCmd) Then
                    Beep
                End If
            End If
        
        Case "A" 'Add item
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                tmr.Tag = "AddItem"
            End If
        
        Case "B" 'Bar Display Type -- toggle: OHLC / Candlestick / Close Line
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                m.Chart.BarDisplayType = -999
                m.Chart.GenerateChart
            End If
        
        Case "C" 'Comparison Symbol
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                If AddCompSymbol(m.Chart, False, True) <> 0 Then        '6497
                    m.Chart.GenerateChart eRedo5_RecalcInd
                End If
            End If
        
        Case "D" 'Daily
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                m.Chart.ChangeBarPeriod "Daily"
            End If
            
        Case "E" 'Edit settings
            ' TLB 5/19/2008: since the Edit Settings form is not modal, we should be able
            ' to just call it from here immediately (rather than through a timer event)
            'tmr.Tag = "EditSettings"
            EditSettings
        
        Case "F" 'Floating price/date Tips
            g.ChartGlobals.bFloatingTips = Not g.ChartGlobals.bFloatingTips
            UpdateVisibleCharts eRedo1_Scrolled
        
        Case "G" 'Grid cycle: coarse/fine/none
            With m.Chart
                If .VertGrid = 2 Then
                    .VertGrid = 0
                Else
                    .VertGrid = .VertGrid + 1
                End If
                .GenerateChart
            End With
        
        Case "H" 'Hide pane
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                If m.MouseLast.bOffChart Or m.Chart.geVisiblePaneCnt < 1 Then
                    Beep
                    StatusMsg "At least one pane must be visible." 'TEST
                Else
                    mnuHidePane.Tag = m.Chart.Tree.Key(m.MouseLast.nPaneID)
                    mnuHidePane_Click
                End If
            End If
        
        Case "M" 'Monthly
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                m.Chart.ChangeBarPeriod "Monthly"
            End If
            
        Case "P" 'Bar time period
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                If frmMain.tbToolbar.Tools("ID_BarPeriod").TagVariant Then
                    ' if combo is on the toolbar, move focus to it
    '                If g.oAvailToolButtons Is Nothing Then
    '                    MoveFocus frmMain.tbToolbar.ToolBars(kTbChartSettings).Tools("ID_BarPeriod")
    '                Else
                        If m.eDetachStatus = eDetached And cboBarPeriod.Visible Then
                            MoveFocus cboBarPeriod
                            cboBarPeriod.SelLength = Len(cboBarPeriod.Text)
                        ElseIf frmMain.cboBarPeriod.Visible Then
                            MoveFocus frmMain.cboBarPeriod
                            frmMain.cboBarPeriod.SelLength = Len(frmMain.cboBarPeriod.Text)
                        Else
                            m.Chart.ChangeBarPeriod "Custom"
                        End If
    '                End If
                Else
                    m.Chart.ChangeBarPeriod "Custom"
                    ' otherwise go into "command mode"
                    'm.eCmdMode = eCmdMode_BarPeriod '(to start typing bar period)
                End If
            End If

        Case "Q" 'Qtrly
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                m.Chart.ChangeBarPeriod "Qtrly"
            End If
        
        Case "R" 'Rescale (if in manual scale)
            Set Pane = m.Chart.Tree("Price Pane")
            If Not Pane Is Nothing Then
                If Pane.Scaling = ePANE_ScaleModeManual Then
                    Pane.Scaling = ePANE_ScaleModeAuto
                    m.Chart.GenerateChart eRedo1_Scrolled
                    Pane.Scaling = ePANE_ScaleModeManual
                End If
            End If
            Set Pane = Nothing
        
        Case "S" 'Symbol
            If 0 Then
                m.eCmdMode = eCmdMode_Symbol '(to start typing symbol)
            Else
                bSymChangeOk = True
                If m.bGameMode And Not m.oGameMode Is Nothing Then
                    If m.oGameMode.CustomOrders > 0 Then bSymChangeOk = False
                End If
                If bSymChangeOk Then
                    If m.eOrdBarMode = eOrdBarMode_Wizard Then OrderBarModeToggle
                    If Len(m.Chart.SpreadSymbols) > 0 Then
                        frmNewChart.ShowMe m.Chart.SpreadSymbols, True, m.Chart
                    Else
                        Set aStrings = frmSymbolSelector.ShowMe(m.Chart.Symbol, False, True, "Symbol for the Chart", True, , , , True)
                        If aStrings.Size > 0 Then
                            If InStr(aStrings(0), "|") > 0 Then
                                m.Chart.SetSymbol aStrings(0), True
                            Else
                                nRec = g.SymbolPool.PoolRecForSymbol(aStrings(0), True)
                                If nRec >= 0 Then
                                    i = LockWindowUpdate(pbChart.hWnd)
                                    m.Chart.SetSymbol g.SymbolPool.SymbolID(nRec), True
                                    m.Chart.GenerateChart eRedo1_Scrolled           '4376
                                    If m.bGameMode And Not m.oGameMode Is Nothing Then
                                        m.oGameMode.InitGame Me                     '4094
                                    End If
                                    If i <> 0 Then LockWindowUpdate 0
                                Else
                                    Beep
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
        Case "T" 'Toggle drawing tool on/off
            If Len(g.strActiveDraw) = 0 Then
                ToolbarSetCursorGroup frmMain.tbToolbar, True '', "ID_Trendline"
            ElseIf InStr(g.strActiveDraw, "PFP") = 0 Then
                ToolbarSetCursorGroup frmMain.tbToolbar, False '', "ID_Trendline"
            ElseIf Not Me.fraPatternProfit.Visible Then
                ToolbarSetCursorGroup frmMain.tbToolbar, True       '5809
            End If
            
            'need to send mouse event to keep toolbar buttons in sync (5809)
            'there is no drawtool button for PFP so mouseevent is never sent
            Dim DrawButton As cPicBoxButton
            
            If Len(g.strActiveDraw) = 0 Then
                SyncDrawTools True
            ElseIf m.Chart.ShowToolbar Then
                Set DrawButton = ButtonByID(Me, g.strActiveDraw, kTbDraw)
                If Not DrawButton Is Nothing Then
                    ToolbarMouseEvent Me, m.oBtnMouseLast, WM_LBUTTONDOWN, vbLeftButton, DrawButton.PicboxIndex, 0, 0, True, DrawButton
                End If
            Else
                Set DrawButton = ButtonByID(frmMain, g.strActiveDraw, kTbDraw)
                If Not DrawButton Is Nothing Then
                    ToolbarMouseEvent frmMain, frmMain.LastMouseButton, WM_LBUTTONDOWN, vbLeftButton, DrawButton.PicboxIndex, 0, 0, True, DrawButton
                End If
            End If
                        
            MoveFocus pbChart
            m.Chart.SetCursor
            pbChart_MouseMove vbLeftButton, 0, m.MouseLast.MouseX + 1, m.MouseLast.MouseY
        
        Case "U" 'Unsplit: toggle between adjust/unadjust
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                m.Chart.Unsplit = Not m.Chart.Unsplit
                m.Chart.GenerateChart eRedo9_ReloadData
            End If
            
        Case "W" 'Weekly
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                m.Chart.ChangeBarPeriod "Weekly"
            End If
        
        Case "Y" 'Yearly
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                m.Chart.ChangeBarPeriod "Yearly"
            End If
                        
        Case Else
            If KeyAscii = vbKeyReturn Then
                m.Chart.RestoreChartNormal vbKeyReturn      '- change hot-key for our "chart reset" to "Enter"
            Else
                bUnhandled = True
            End If
        End Select
        strCmd = ""
    Else
        bUnhandled = True
    End If
    
    ' fix caption (only if from KeyPress event)
    If Shift < 0 Then
        Select Case m.eCmdMode
        Case eCmdMode_Symbol
            strCaption = "Symbol:  " & strCmd
        Case eCmdMode_BarPeriod
            strCaption = "Bar Time Period:  " & strCmd
        End Select
        m.Chart.SetFormCaption strCaption
    End If
    
    ' eat the key unless not handled here
    If Not bUnhandled Then
        KeyAscii = 0
        DoEvents
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".KeyPress"

End Sub

Public Property Get TopMost() As Boolean
On Error GoTo ErrSection:

    TopMost = m.bTopMost

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".TopMost.Get", eGDRaiseError_Raise
    
End Property

Public Property Let TopMost(ByVal bTopMost As Boolean)
On Error GoTo ErrSection:

'JM: 06-30-2009 - don't think this is used anymore; leave awhile then remove if all okay
'    If bTopMost Then
'        SetFormTopmost Me, True
'        If Not Me.MDIChild Then
'            If m.Chart.tbToolbar.Tools("ID_TopMost").State <> ssChecked Then
'                m.Chart.tbToolbar.Tools("ID_TopMost").State = ssChecked
'            End If
'        End If
'    ElseIf m.bTopMost Then
'        SetFormTopmost Me, False
'        If Not Me.MDIChild Then
'            If m.Chart.tbToolbar.Tools("ID_TopMost").State <> ssUnchecked Then
'                m.Chart.tbToolbar.Tools("ID_TopMost").State = ssUnchecked
'            End If
'        End If
'    End If
'
'    m.bTopMost = bTopMost

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".TopMost.Let", eGDRaiseError_Raise
    
End Property

Private Sub vseScrollSep_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    SetFocusCtl   'Peg    -MJM
    KeyPress KeyCode, Shift

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseScrollSep_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub vseScrollSep_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseScrollSep_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub vsePeriodLink_Click()
On Error GoTo ErrSection:

    frmWindowLink.ShowMe Me, eLink_Period, vsePeriodLink.Left, vsePeriodLink.Top + vsePeriodLink.Height

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vsePeriodLink_Click"
    Resume ErrExit
End Sub

Public Sub ClearBuySellButtons(Optional ByVal bClearNow As Boolean = False)
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    Dim bRedraw As Boolean
    Dim bBalloonTool As Boolean
    
    If Not m.AnnotOptions Is Nothing Then
        Set Annot = m.Chart.ClosestOptionAnnot(m.MouseLast.dDate)
        If Not Annot Is Nothing Then bBalloonTool = Annot.IsBalloonOptionsInfo
    End If
    
    'no need to do this if order bar is not on
    If Not fraOrderBtns.Visible And Not bBalloonTool Then
        Exit Sub
    End If
    
    If Not bClearNow Then
        If Not ActiveChart Is Nothing Then
            If ActiveChart.vseBracketOrder.Appearance = apInset Then
                Exit Sub
            End If
        End If
    End If
    
    If vseBuyChart.Appearance = apInset Or vseSellChart.Appearance = apInset Or _
       vseBuyWizard.Appearance = apInset Or vseSellWizard.Appearance = apInset Or _
       vseBracketOrder.Appearance = apInset Or bClearNow Then
       
        fraWizardPrompt.Visible = False
        fraWizardPrompt.Enabled = False
        
        vseBuyChart.Appearance = ap3D
        vseSellChart.Appearance = ap3D
        
        vseBuyWizard.Appearance = ap3D
        vseSellWizard.Appearance = ap3D
        
        vseBracketOrder.Appearance = ap3D
        
        ' MJM 02/04/2011: Per Dave issue 6159 is due to some changes in mTradeTracker.bas on 01/26/2011
        '   that unmasked this bug (he thinks bug has always been here). Dave's reccommendation is to
        '   reload order objects to get latest status. This fix works and bug must have always been around.
        
        If Not m.oBracketOrdOne Is Nothing Then
            m.oBracketOrdOne.Reload
            If m.oBracketOrdOne.Status = eTT_OrderStatus_Parked Then CancelOrder m.oBracketOrdOne, False
            bRedraw = True
        End If
        
        Set m.oBracketOrdOne = Nothing
        If Not m.oBracketOrdTwo Is Nothing Then
            m.oBracketOrdTwo.Reload
            'Issue 6159 is due to the 2nd order having a status of Parked and triggering the CancelOrder
            If m.oBracketOrdTwo.Status = eTT_OrderStatus_Parked Then CancelOrder m.oBracketOrdTwo, False
            bRedraw = True
        End If
        Set m.oBracketOrdTwo = Nothing
        
        If Not m.Chart Is Nothing Then
            If Not m.Chart.Annots Is Nothing Then
                Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
                If Not Annot Is Nothing Then
                    If Annot.eUsage = eANNOT_OptionInfo Then
                        If bBalloonTool Then m.Chart.RemoveAnnots False, , eANNOT_OptionInfo, False
                        m.nActiveAnnotIdx = -1         '5101
                    End If
                End If
            End If
            
            g.ChartGlobals.eChartMode = g.ChartGlobals.ePrevChartMode
            
            If Len(g.strActiveDraw) = 0 Then    '5306
                ToolbarSetCursorGroup frmMain.tbToolbar, False                  '4877
                m.Chart.SetCursor
                StatusMsg
            End If
            
            If bBalloonTool Then
                If lblOrderBarMode.Caption = "Underlying" Then
                    m.eOrdBarMode = eOrdBarMode_Order
                End If
                If Not m.AnnotOptions Is Nothing Then
                    tmr.Tag = "EditAnnot " & m.AnnotOptions.geAnnId
                    Set m.AnnotOptions = Nothing
                End If
            ElseIf bRedraw Then
                m.Chart.GenerateChart eRedo1_Scrolled
            End If
        End If

    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ClearBuySellButtons"
    Resume ErrExit
End Sub

Private Sub vseBuyChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    vseBuyChart.ToolTipText = "Click on chart to buy " & Trim(txtTradeQty) & " at selected price"

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".vseBuyChart_MouseMove"

End Sub

Private Sub vseSellChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    vseSellChart.ToolTipText = "Click on chart to sell " & Trim(txtTradeQty) & " at selected price"

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".vseSellChart_MouseMove"

End Sub

Private Sub vseSellChart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    HandleBuySellClick vseSellChart, Button

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseSellChart_MouseUp"

End Sub

Private Sub vseSellWizard_Click()
On Error GoTo ErrSection:

    HandleBuySellClick vseSellWizard

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseSellChart_Click"

End Sub

Private Sub vseSymbolLink_Click()
On Error GoTo ErrSection:

    frmWindowLink.ShowMe Me, eLink_Symbol, vseSymbolLink.Left, vseSymbolLink.Top + vseSymbolLink.Height

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseSymbolLink_Click"
    Resume ErrExit
End Sub

Private Sub vseTipX_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    SetFocusCtl   'Peg    -MJM
    KeyPress KeyCode, Shift

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseTipX_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub vseTipX_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseTipX_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub vseTipY_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    SetFocusCtl   'Peg    -MJM
    KeyPress KeyCode, Shift

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseTipY_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub vseTipY_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vseTipY_KeyPress"
    Resume ErrExit
    
End Sub

Public Property Get AutoSize() As Integer
On Error GoTo ErrSection:

    AutoSize = m.iAutoSize

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".AutoSize.Get", eGDRaiseError_Raise
   
End Property

Public Property Let AutoSize(ByVal iAutoSize As Integer)
On Error GoTo ErrSection:

    m.iAutoSize = iAutoSize

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".AutoSize.Let", eGDRaiseError_Raise
    
End Property

'Returns true if snapped
Private Function SnapToHiLoClose(MouseCoord As ChartCoordinates, _
        dY#, ByVal nHiLoClose&, Optional ByVal bPeakCheck As Boolean = False) As Boolean
On Error GoTo ErrSection:
'[IN] nHiLoClose: 0=snap to closest high or low
'                 1=snap to high
'                 2=snap to low
'                 3=snap to close
'                 4=snap to price (for free-floating fib extension)

    Dim bHasBars As Boolean, bSnapped As Boolean
    Dim dyMouse#, nPeakBar&
    
    dY = 0
    SnapToHiLoClose = False
    ' see if pane has bars data
    If MouseCoord.bOffChart = True Then Exit Function
    If m.Chart.Tree.Key(MouseCoord.nPaneID) = "PRICE PANE" Then
        bHasBars = True
    Else
        Exit Function
    End If
        
    dyMouse = MouseCoord.dY
    If nHiLoClose = 0 Then      'snap to closest hi or low
        m.nFocusHiLo = -1       'reset
        SetDnExHiLoFlag MouseCoord
        If m.nFocusHiLo = -1 Then Exit Function     'failed to identify whether point is close to high or low of bar
        nHiLoClose = m.nFocusHiLo
    End If
    
    bSnapped = True
    Select Case nHiLoClose
        Case 0:
            dY = dY         'just dummy code as place holder
        Case 1:
            If bPeakCheck Then
                ' see if 1 bar off from a high peak
                nPeakBar = MouseCoord.nBar
                If m.Chart.Bars(eBARS_High, MouseCoord.nBar - 1) = kNullData Then
                    nPeakBar = MouseCoord.nBar + 1
                ElseIf m.Chart.Bars(eBARS_High, MouseCoord.nBar + 1) = kNullData Then
                    nPeakBar = MouseCoord.nBar - 1
                ElseIf m.Chart.Bars(eBARS_High, MouseCoord.nBar - 1) > m.Chart.Bars(eBARS_High, MouseCoord.nBar + 1) Then
                    nPeakBar = MouseCoord.nBar - 1
                Else
                    nPeakBar = MouseCoord.nBar + 1
                End If
                If m.Chart.Bars(eBARS_High, nPeakBar) = kNullData Then
                    nPeakBar = MouseCoord.nBar
                ElseIf m.Chart.Bars(eBARS_High, nPeakBar) <= m.Chart.Bars(eBARS_High, MouseCoord.nBar) Then
                    nPeakBar = MouseCoord.nBar
                ElseIf nPeakBar > MouseCoord.nBar And m.Chart.Bars(eBARS_High, nPeakBar) <= m.Chart.Bars(eBARS_High, nPeakBar + 1) Then
                    nPeakBar = MouseCoord.nBar
                ElseIf nPeakBar < MouseCoord.nBar And m.Chart.Bars(eBARS_High, nPeakBar) <= m.Chart.Bars(eBARS_High, nPeakBar - 1) Then
                    nPeakBar = MouseCoord.nBar
                End If
                If nPeakBar <> MouseCoord.nBar Then
                    ' if y-value of mouse is closer to the peak, then snap to it
                    If dyMouse > (m.Chart.Bars(eBARS_High, nPeakBar) + m.Chart.Bars(eBARS_High, MouseCoord.nBar)) / 2# Then
                        MouseCoord.nBar = nPeakBar
                        MouseCoord.dDate = m.Chart.Bars(eBARS_DateTime, nPeakBar)
                    End If
                End If
            End If
            dY = m.Chart.Bars(eBARS_High, MouseCoord.nBar)
        Case 2:
            If bPeakCheck Then
                ' see if 1 bar off from a low peak
                nPeakBar = MouseCoord.nBar
                If m.Chart.Bars(eBARS_Low, MouseCoord.nBar - 1) = kNullData Then
                    nPeakBar = MouseCoord.nBar + 1
                ElseIf m.Chart.Bars(eBARS_Low, MouseCoord.nBar + 1) = kNullData Then
                    nPeakBar = MouseCoord.nBar - 1
                ElseIf m.Chart.Bars(eBARS_Low, MouseCoord.nBar - 1) < m.Chart.Bars(eBARS_Low, MouseCoord.nBar + 1) Then
                    nPeakBar = MouseCoord.nBar - 1
                Else
                    nPeakBar = MouseCoord.nBar + 1
                End If
                If m.Chart.Bars(eBARS_Low, nPeakBar) = kNullData Then
                    nPeakBar = MouseCoord.nBar
                ElseIf m.Chart.Bars(eBARS_Low, nPeakBar) >= m.Chart.Bars(eBARS_Low, MouseCoord.nBar) Then
                    nPeakBar = MouseCoord.nBar
                ElseIf nPeakBar > MouseCoord.nBar And m.Chart.Bars(eBARS_Low, nPeakBar) >= m.Chart.Bars(eBARS_Low, nPeakBar + 1) Then
                    nPeakBar = MouseCoord.nBar
                ElseIf nPeakBar < MouseCoord.nBar And m.Chart.Bars(eBARS_Low, nPeakBar) >= m.Chart.Bars(eBARS_Low, nPeakBar - 1) Then
                    nPeakBar = MouseCoord.nBar
                End If
                If nPeakBar <> MouseCoord.nBar Then
                    ' if y-value of mouse is closer to the peak, then snap to it
                    If dyMouse < (m.Chart.Bars(eBARS_Low, nPeakBar) + m.Chart.Bars(eBARS_Low, MouseCoord.nBar)) / 2# Then
                        MouseCoord.nBar = nPeakBar
                        MouseCoord.dDate = m.Chart.Bars(eBARS_DateTime, nPeakBar)
                    End If
                End If
            End If
            dY = m.Chart.Bars(eBARS_Low, MouseCoord.nBar)
        Case 3:
            dY = m.Chart.Bars(eBARS_Close, MouseCoord.nBar)
        Case 4:
            dY = SnapToPrice(m.MouseLast)
            If m.nPointCount = 0 Then SetDnExHiLoFlag MouseCoord
        Case Default:
            bSnapped = False
    End Select
    
    SnapToHiLoClose = bSnapped
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".SnapToHiLoClose", eGDRaiseError_Raise
    
End Function

' Returns price which is "snapped" to OHLC price (if close enough)
' else is "rounded" for certain types of annotations
Private Function SnapToPrice(MouseCoord As ChartCoordinates, _
        Optional ByVal eType As eAnnotType = eANNOT_UndefinedType) As Double
On Error GoTo ErrSection:

    Dim dCloseEnough#, dChk#, nBar&
    Dim dX#, dY#, nPaneID&, pixX&, pixY&, rc&, i&
    Dim bHasBars As Boolean, bSnapped As Boolean

    Dim cInfo As coordinate_info
    Dim Ind As cIndicator
    Dim eDisplay As eIndicatorDisplayType
    
    ' see if pane has bars data
    With MouseCoord
        If Not .bOffChart Then
            If m.Chart.Tree.Key(.nPaneID) = "PRICE PANE" Then
                bHasBars = True
                Set Ind = m.Chart.Tree("PRICE")
                If Not Ind Is Nothing Then
                    eDisplay = Ind.DisplayType
                Else
                    eDisplay = eINDIC_OHLC
                End If
            End If
        End If
    
        If bHasBars Then
            ' calculate distance of one pixel in Y direction
            ' (check up or down a pixel, but make sure to stay in same pane)
            pixY = (.MouseY \ Screen.TwipsPerPixelY)
            pixX = (.MouseX \ Screen.TwipsPerPixelX)
            If .dY > (.dMaxY + .dMinY) / 2 Then
                pixY = pixY + g.ChartGlobals.nMagnetValue '(add 1 to go "down" a pixel if in upper half)
            Else
                pixY = pixY - g.ChartGlobals.nMagnetValue '(subtract 1 to go "up" a pixel if in lower half)
            End If
            
            nPaneID = -1
            cInfo.paneId = -1
            cInfo.x_pixels = pixX
            cInfo.y_pixels = pixY
            rc = geCoordToData(m.Chart.geChartObj, cInfo)
            If rc = 0 Then
                nPaneID = cInfo.paneId
                dX = cInfo.x_value
                dY = cInfo.y_value
                ' convert back in order to see if mouse is right/left of bar
                rc = geDataToCoord(m.Chart.geChartObj, cInfo)
            End If
            
            If nPaneID = .nPaneID Then
                ' calculate "close enough" in # of pixels
                dCloseEnough = Abs(.dY - dY) * 1.05 '(in case of rounding issues)
            End If
            
            If .nBar < 0 Then .nBar = 0     '6941
            
            dY = .dY
            nBar = .nBar
            
            If g.ChartGlobals.nMagnetValue > 0 And dCloseEnough > 0 Then
                ' find nearest of OHLC to dY
                If eDisplay = eINDIC_HL Then
                    ' ignore open and close
                    dChk = m.Chart.Bars(eBARS_High, .nBar)
                ElseIf pixX < cInfo.x_pixels And Not eDisplay = eINDIC_HLC Then
                    ' look at open if mouse is left of bar
                    dChk = m.Chart.Bars(eBARS_Open, .nBar)
                Else
                    ' else look at close (mouse is right of bar)
                    dChk = m.Chart.Bars(eBARS_Close, .nBar)
                End If
                If Abs(m.Chart.Bars(eBARS_High, .nBar) - dY) < Abs(dChk - dY) Then
                    dChk = m.Chart.Bars(eBARS_High, .nBar)
                End If
                If Abs(m.Chart.Bars(eBARS_Low, .nBar) - dY) < Abs(dChk - dY) Then
                    dChk = m.Chart.Bars(eBARS_Low, .nBar)
                End If
                ' TLB 10/10/2013: if bars are squished, also check high/low of 1 bar on either side
                If m.Chart.PixelsPerBar <= 5 Then
                    If Abs(m.Chart.Bars(eBARS_High, .nBar + 1) - dY) < Abs(dChk - dY) Then
                        dChk = m.Chart.Bars(eBARS_High, .nBar + 1)
                        nBar = .nBar + 1
                    End If
                    If Abs(m.Chart.Bars(eBARS_Low, .nBar + 1) - dY) < Abs(dChk - dY) Then
                        dChk = m.Chart.Bars(eBARS_Low, .nBar + 1)
                        nBar = .nBar + 1
                    End If
                    If Abs(m.Chart.Bars(eBARS_High, .nBar - 1) - dY) < Abs(dChk - dY) Then
                        dChk = m.Chart.Bars(eBARS_High, .nBar - 1)
                        nBar = .nBar - 1
                    End If
                    If Abs(m.Chart.Bars(eBARS_Low, .nBar - 1) - dY) < Abs(dChk - dY) Then
                        dChk = m.Chart.Bars(eBARS_Low, .nBar - 1)
                        nBar = .nBar - 1
                    End If
                End If
                ' see if close enough
                If Abs(dChk - dY) < dCloseEnough Then
                    dY = dChk
                    bSnapped = True
                    ' if bar# changed, fix the MouseCoord properties
                    If nBar <> .nBar Then
                        .nX = .nX + (nBar - .nBar)
                        .nBar = nBar
                        .dDate = m.Chart.Bars(eBARS_DateTime, nBar)
                    End If
                End If
            End If
        End If
        
        If Not bSnapped Then
            ' "round" for certain types of annotations
            If eType = eANNOT_HorzLine Or eType = eANNOT_HorzLine2 Or eType = eANNOT_HorzLine3 Or eType = eANNOT_HorzLine4 Then
                dY = ValOfText(.strRoundedY)
            Else
                dY = .dY
            End If
        End If
    End With
    
    SnapToPrice = dY

ErrExit:
    Set Ind = Nothing
    Exit Function
    
ErrSection:
    Set Ind = Nothing
    RaiseError Me.Name & ".SnaptoPrice", eGDRaiseError_Raise
    
End Function

Private Function EditData(ByVal nX&) As Boolean
On Error GoTo ErrSection:

    Dim i&, nBar&, dDate#, dStartTime#, dEndTime#, d#
    Dim strSymbol$, nSymbolID&, bFromContinuous As Boolean
    Dim Bars As New cGdBars
    Dim dCrossOver As Double
    Dim dSessionEnd As Double
    
    nBar = m.Chart.aXBar(nX)
    If nBar >= 0 And nBar < m.Chart.Bars.Size Then
        If m.Chart.SymbolID <= 0 Then
            InfBox "Data for this symbol cannot be edited.", "i", , "Edit Data"
        ElseIf IsIntraday(m.Chart.Periodicity) Then
            ' Edit ticks:
            ' get underlying contract
            strSymbol = m.Chart.RollSymbol(nX)
            If Len(strSymbol) > 0 Then
                nSymbolID = SU_GetSymID(strSymbol)
                bFromContinuous = True
            Else
                nSymbolID = m.Chart.SymbolID
                bFromContinuous = False
            End If
            frmEditTicks.ShowMe nSymbolID, m.Chart.Bars, nBar
            EditData = True
        ElseIf m.Chart.Periodicity = ePRD_Days + 1 Then
            ' Edit daily data:
            ' get underlying contract
            strSymbol = m.Chart.RollSymbol(nX)
            If Len(strSymbol) > 0 Then
                nSymbolID = SU_GetSymID(strSymbol)
                bFromContinuous = True
            Else
                nSymbolID = m.Chart.SymbolID
                bFromContinuous = False
            End If
            dDate = Int(m.Chart.Bars(eBARS_DateTime, nBar))
            frmEditOHLC.ShowMe nSymbolID, dDate, bFromContinuous
            EditData = True
        Else
            InfBox "You must be on a Daily chart to edit the data.", "i", , "Edit Data"
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".EditData", eGDRaiseError_Raise
    
End Function

Public Property Get ImgSrvState() As eImgSrvState
On Error GoTo ErrSection:

    ImgSrvState = m.eImgSrv

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".ImgSrvState.Get", eGDRaiseError_Raise
    
End Property

Public Property Let ImgSrvState(ByVal eImgSrv As eImgSrvState)
On Error GoTo ErrSection:

    Dim nInterval&
    Static nInitialInterval&
    
    ' first time: save the initial setting
    If nInitialInterval = 0 Then nInitialInterval = tmr.Interval
    
    m.eImgSrv = eImgSrv
    
    If frmImageServer.Active Then
    'If m.eImgSrv = eImgSrv_Searching Then
        ' if this is the chart that's searching the image server
        ' queue, have it look every 1/10 second
        nInterval = 100
    Else
        ' otherwise, set back to initial setting
        nInterval = nInitialInterval
    End If
    If nInterval <> tmr.Interval Then tmr.Interval = nInterval
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".ImgSrvState.Let", eGDRaiseError_Raise
    
End Property

Public Sub SetChartTabs(Optional ByVal bChartNameChanged As Boolean)
On Error GoTo ErrSection:

    Dim i&, iPos&, strText$, nTab&, strTabs$, iMaxTabsWithPeriod&
    Dim bCustomOrder As Boolean
    Static bInProgress As Boolean

    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "UNLOADING" Then Exit Sub
    If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then Exit Sub
    If m.aTabs Is Nothing Then Exit Sub
    
    If Not Me.MDIChild Then Exit Sub
    If Me.WindowState <> vbMaximized Then Exit Sub
    If Not Me Is g.ChartGlobals.frmActiveNonDetached And Not bChartNameChanged Then
        If Me Is frmMain.ActiveForm Then
            SendMessage Me.hWnd, WM_NCACTIVATE, 1, 0
            SendMessage Me.hWnd, WM_MOUSEACTIVATE, 1, 0     '5052, 5505
        Else
            Exit Sub
        End If
    End If

    If bInProgress Then Exit Sub
    bInProgress = True
    
    bCustomOrder = GetCustomTabOrder(bChartNameChanged)
    ' show period with symbol for max # tabs based on MDI client width
    ' (approx: 800 = 6, 1024 = 8, 1280 = 10, 1400 = 11, 1600 = 13 tabs)
    iMaxTabsWithPeriod = Round(frmMain.ScaleWidth / 1800)
    
    If g.nColorTheme <> 0 Then vseCaption.Font.Bold = -1 * g.ChartGlobals.nFontStyle
    'to let resize event draw fake caption on load or create of chart page that has only one maximized chart
    vseCaption.Top = -1
    
    If m.aTabs.Size <= 1 Then
        vsTab.Visible = False
        vsTab.Enabled = False
'        Form_Resize
    Else
        vsTab.Enabled = True
        vsTab.Visible = True
'        Form_Resize
    
        If Not bCustomOrder Then
            m.aTabs.Sort eGdSort_IgnoreCase Or eGdSort_Stable
        End If
        For i = 0 To m.aTabs.Size - 1
            strText = Parse(m.aTabs(i), vbTab, 1)
            If m.aTabs.Size > iMaxTabsWithPeriod Then
                ' shorter name: only keep what's before the parenthesis
                iPos = InStr(strText, "(")
                If iPos > 0 Then strText = Trim(Left(strText, iPos - 1))
            Else
                strText = Replace(strText, "(  ", "(")
                strText = Replace(strText, "( ", "(")
            End If
            strText = Replace(strText, "&", "&&")
            strTabs = strTabs & "|" & strText
            strText = Parse(m.aTabs(i), vbTab, 3)
            If Val(strText) = Me.hWnd Then nTab = i
        Next
        
        vsTab.Caption = Mid(strTabs, 2)
        vsTab.FirstTab = 0
        vsTab.CurrTab = nTab
        vsTab.BoldCurrent = True
        'vsTab.Refresh
    End If

ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError Me.Name & ".SetChartTabs", eGDRaiseError_Raise
    
End Sub

Private Sub vseTSO1_Click()
    TSOGrpFavoritesBtnClick Me, vseTSO1, 0
End Sub

Private Sub vseTSO2_Click()
    TSOGrpFavoritesBtnClick Me, vseTSO2, 1
End Sub

Private Sub vseTSO3_Click()
    TSOGrpFavoritesBtnClick Me, vseTSO3, 2
End Sub

Private Sub vseTSO4_Click()
    TSOGrpFavoritesBtnClick Me, vseTSO4, 3
End Sub

Private Sub vsTab_GotFocus()
On Error GoTo ErrSection:

    'SetFocusCtl

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vsTab_GotFocus"
    Resume ErrExit
    
End Sub

Private Sub vsTab_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    KeyPress KeyCode, Shift

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vsTab_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub vsTab_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vsTab_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub vsTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim strTip$, nTab&
    Static strPrevTip$
    
    On Error Resume Next
    nTab = vsTab.MouseOver
    If nTab < 0 Then
        'strTip = "(can right-click chart tabs to rename them)"
        strTip = "<-- can right-click chart tabs to rename them"
    ElseIf nTab = vsTab.CurrTab Then
        strTip = Trim(vseCaption.Caption)
    Else
        strTip = Parse(m.aTabs(nTab), vbTab, 2)
    End If
    If strTip <> strPrevTip Then
        vsTab.ToolTipText = Replace(strTip, "&&", "&")
        strPrevTip = strTip
    End If

End Sub

Private Sub vsTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    'Dim i&, hWnd&, nTab&, strText$, strName$, strHdr$
    Dim nTab&, strText$, strName$, strTabCaption$
    Dim iNewTab As Integer, iCancel As Integer
    Dim frm As Form
    
    If Button = 2 Then
        nTab = vsTab.MouseOver
        If nTab >= 0 Then
            If nTab <> vsTab.CurrTab Then
                iNewTab = nTab
                vsTab_Switch vsTab.CurrTab, iNewTab, iCancel
            End If
            
            strTabCaption = Parse(vsTab.Caption, "|", nTab + 1)
            If ActiveChart Is g.ChartGlobals.frmActiveNonDetached And ActiveChart.WindowState = vbMaximized Then
                Set frm = ActiveChart
                strText = "Enter a custom name for the chart|(or click 'Clear' to revert to default name):"
                strName = InfBox(strText, "?", "+OK|-Clear", frm.Chart.ChartName(False), , , , , , "s", strTabCaption) 'strName)
                If Len(strName) = 0 Then strName = Trim(frm.Chart.ChartName(False, True))     '5506
                If Len(strName) > 0 Then
                    If strName <> strTabCaption Then
                        frm.Chart.SetChartName strName
                        frm.Chart.TemplateSave
                        frm.SetChartTabs
                        g.bDirtyChartPage = True
                    End If
                End If
            End If
            
            Set frm = Nothing
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".vsTab_MouseUp"
    Resume ErrExit
End Sub

Private Sub vsTab_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    Dim i&, hWnd&
    Static bInProgress As Boolean
    
    If Not Me.MDIChild Then Exit Sub
    If g.bStarting Or g.bLoadingChartPage Then Exit Sub
    If bInProgress Then Exit Sub ' TLB 12/9/2011: to fix Pete's issue (when someone clicks twice fast on the tab)
    bInProgress = True
    
    DoEvents ' this is needed to fix Aardvark #3053
    
    ' look for the chart to switch to
    hWnd = Val(Parse(m.aTabs(NewTab), vbTab, 3))
    If hWnd = Me.hWnd Then
        bInProgress = False
        Exit Sub
    End If
    For i = Forms.Count - 1 To 0 Step -1
        If Forms(i).hWnd = hWnd And hWnd <> 0 Then
            ' unbold the current tab in order to immediately show
            ' that we are doing something (since it will take a little
            ' while before the user sees the tab actually "switch")
            ActiveChartFormSet Forms(i)
            m.bSkipFocusFix = True
            vsTab.BoldCurrent = False
            vsTab.Refresh
            LockWindowUpdate (frmMain.hWnd)
            If Me.WindowState = vbMaximized Then
                ' first set caption blank so MDI window caption won't flicker
                Forms(i).Caption = ""
                Forms(i).WindowState = vbMaximized
            End If
            ' change focus to desired chart
            Cancel = True
            ' set tabs back the way they were for this chart
            vsTab.BoldCurrent = True
            MoveFocus Forms(i)
            TextIncDecRegisterForm Forms(i), True           '6463
            frmMain.tmrMain.Tag = "UnlockWindowUpdate"
            m.bSkipFocusFix = False
            bInProgress = False
            Exit Sub
        End If
    Next
    
    Cancel = True
    Beep

ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError Me.Name & ".vsTab_Switch"
    Resume ErrExit
    
End Sub

' for debugging purposes
Private Sub LogEvent(ByVal strEvent$)
    
    'DebugLog strEvent & ", " & CStr(Me.WindowState) & ", " & Parse(vseCaption.Caption, ":", 1)
    
End Sub

' If the system menu has magically grayed out (seems like bug with
' XP and 2000), we have to consider this form "corrupt" and we'll
' transfer the whole chart to a new form and get rid of this one!
' 9/6/02: I think we now have a fix in to keep the boxes from ever
'   getting grayed out in the first place, but we'll leave this
'   code in just in case it ever does happen.
Private Sub CheckSystemMenu()
On Error GoTo ErrSection:

    Dim hMenu&, uFlags&, strSymbol$
    Dim frm As Form
    
    If Not Me.MDIChild Then Exit Sub
        
    If Me.IsInGameMode Or g.bSkipSetChartFocus Then
        'don't want to do any of this for form that is game mode
        Exit Sub
    End If
    
    ' see if system menu has grayed out
    hMenu = GetSystemMenu(Me.hWnd, 0)
    If hMenu <> 0 Then
        uFlags = GetMenuState(hMenu, SC_MINIMIZE, 0)
        If (uFlags And MF_DISABLED) Or (uFlags And MF_GRAYED) Then
            ' if so, transfer chart to a new form
            Set frm = New frmChart          'must be this type for transferring a non-detached chart
            m.Chart.TemplateSave
            If Not frm.Chart.TemplateLoad(m.Chart.Template) Then
                Unload frm
            Else
                SetFormPlacement frm, GetFormPlacement(Me), "P"
                strSymbol = m.Chart.ExternalData
                If Len(strSymbol) = 0 Then
                    strSymbol = m.Chart.SpreadSymbols
                End If
                If Len(strSymbol) > 0 Then
                    frm.Chart.SetSymbol strSymbol
                Else
                    frm.Chart.SetSymbol m.Chart.SymbolID
                End If
                'If ws = 1 Then frm.WindowState = 1
                
                tmr.Enabled = False
                tmr.Tag = "UNLOADING"
                LockWindowUpdate frmMain.hWnd
                frm.Show
                Sleep 0 '(need this so won't go into infinite loop on last chart)
                Unload Me
                frmMain.tmrMain.Tag = "UnlockWindowUpdate"
                DebugLog "Chart transferred to new form"
            End If
            Set frm = Nothing
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".CheckSystemMenu", eGDRaiseError_Raise
    
End Sub

Public Property Get FocusHiLow() As Long
On Error GoTo ErrSection:

    FocusHiLow = m.nFocusHiLo

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".FocusHiLow.Get", eGDRaiseError_Raise
    
End Property

Private Sub SetGannActivePt(hitInfo As hittest_info)
On Error GoTo ErrSection:

    If hitInfo.location = 8 Then        'anchor point
        m.epbCursor = eCursor_Arrow4Way
        m.nActiveAnnotPt = 1
    ElseIf hitInfo.location = 11 Then   'non-anchor point
        m.epbCursor = eCursor_Arrow4Way
        m.nActiveAnnotPt = 2
    Else
        m.epbCursor = eCursor_Hand      'on line
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetGannActivePt", eGDRaiseError_Raise
    
End Sub

Private Sub SetPatternActivePt(Annot As cAnnotation, hitInfo As hittest_info)
On Error GoTo ErrSection:

    Annot.HitItemIndex = hitInfo.itemIndex
            
    If hitInfo.location = 8 Then
        If hitInfo.itemIndex = 1 And Val(Annot.Prop("ForecastBars")) = 0 Then
            m.epbCursor = eCursor_ArrowEW
            m.nActiveAnnotPt = 3
        Else
            m.epbCursor = eCursor_Hand
            m.nActiveAnnotPt = 0
        End If
    ElseIf hitInfo.itemIndex = 0 And hitInfo.location = 2 Then
        m.epbCursor = eCursor_Hand
        m.nActiveAnnotPt = 0
    ElseIf hitInfo.itemIndex = 1 And hitInfo.location = 2 Then
        m.epbCursor = eCursor_ArrowEW
        m.nActiveAnnotPt = 7
    ElseIf hitInfo.itemIndex = 1 And hitInfo.location = 6 Then
        m.epbCursor = eCursor_ArrowEW
        m.nActiveAnnotPt = 3
    Else
        m.Chart.SetCursor
        m.nActiveAnnotPt = -1
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetPatternActivePt", eGDRaiseError_Raise
    
End Sub

Private Sub SetRectActivePt(Annot As cAnnotation, hitInfo As hittest_info)
On Error GoTo ErrSection:

    If Annot.eType = eANNOT_TextEdit Or Annot.eType = eANNOT_TextEdit2 Or _
       Annot.eType = eANNOT_TextEdit3 Or Annot.eType = eANNOT_TextEdit4 Or _
       Annot.eType = eANNOT_ArrowLine Then
        If hitInfo.itemIndex = 0 Then
            m.epbCursor = eCursor_Hand  'text
            m.nActiveAnnotPt = 1
        ElseIf hitInfo.itemIndex = 1 Then
            If hitInfo.location = 11 And Annot.ArrowLocation = 2 Or _
               hitInfo.location = 10 And Annot.ArrowLocation = 1 Or _
               hitInfo.location = 10 And Annot.ArrowLocation = 2 And Annot.Y(1) < Annot.Y(2) Then
               'the last case is vertical arrow line pointing up
                m.epbCursor = eCursor_Arrow4Way
                m.nActiveAnnotPt = 2
            Else
                m.nActiveAnnotIdx = 0
            End If
        End If
        Exit Sub
    ElseIf Annot.eUsage = eANNOT_PatternProfit Then
        If Not Annot.OkayToMovePFP Then Exit Sub
    End If
        
    If Annot.eUsage = eANNOT_FibClusters Then
        If hitInfo.location = 2 Or hitInfo.location = 6 Then
            m.epbCursor = eCursor_ArrowEW
            m.nActiveAnnotPt = 7
        Else
            ClearAnnotFlags False, False
            m.Chart.SetCursor
        End If
    ElseIf hitInfo.location = 0 Or hitInfo.location = 4 Then
        If Annot.eUsage = eANNOT_PatternProfit Then
            ClearAnnotFlags False, False        'not allowing vertical resize (makes no sense)
            m.Chart.SetCursor
        Else
            m.epbCursor = eCursor_ArrowNS
            m.nActiveAnnotPt = 5
        End If
    ElseIf hitInfo.location = 2 Or hitInfo.location = 6 Then
        m.epbCursor = eCursor_ArrowEW
        m.nActiveAnnotPt = 7
    ElseIf hitInfo.location = 1 Or hitInfo.location = 3 Or _
           hitInfo.location = 5 Or hitInfo.location = 7 Then
        m.epbCursor = eCursor_Arrow4Way
        m.nActiveAnnotPt = 3
    ElseIf hitInfo.location = 8 Then
        m.epbCursor = eCursor_Hand     '8=center of annotation
        m.nActiveAnnotPt = 0
    Else
        If Annot.eType = eANNOT_Mirror Then ClearAnnotFlags False, False
        m.Chart.SetCursor
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetRectActivePt", eGDRaiseError_Raise
    
End Sub

Private Sub SetTDActivePt(Annot As cAnnotation, hitInfo As hittest_info)
On Error GoTo ErrSection:
   
    Dim bTextEdit As Boolean
   
    If Annot.eType = eANNOT_Ellipse Then
        If hitInfo.location = 0 Or hitInfo.location = 4 Then
            m.epbCursor = eCursor_Arrow4Way
            m.nActiveAnnotPt = 4
            Exit Sub
        End If
    End If
    
    If Annot.eType = eANNOT_RegressionLine Then
        If hitInfo.itemIndex <> 0 And hitInfo.itemIndex <> 3 And hitInfo.itemIndex <> 4 Then
            m.epbCursor = eCursor_Hand
            Exit Sub
        End If
    End If

    If Annot.eType = eANNOT_TextEdit Or Annot.eType = eANNOT_TextEdit2 Or _
       Annot.eType = eANNOT_TextEdit3 Or Annot.eType = eANNOT_TextEdit4 Then
       
       bTextEdit = True
       
    End If

    'TD stands for trend & dollar line (this was originally used only for TD)
    m.epbCursor = eCursor_Arrow4Way
    If Annot.X(2) > Annot.X(1) Then
        If hitInfo.location = 9 And bTextEdit And hitInfo.itemIndex = 0 Then
            m.nActiveAnnotPt = 1
        ElseIf hitInfo.location = 10 Then
            m.nActiveAnnotPt = 1
        ElseIf hitInfo.location = 11 Then
            If Annot.eType = eANNOT_FibFan Or Annot.eType = eANNOT_SpResistFan Then
                If hitInfo.itemIndex = 0 Then
                    m.nActiveAnnotPt = 2
                Else
                    m.epbCursor = eCursor_Hand
                End If
            Else
                m.nActiveAnnotPt = 2
            End If
        Else
            m.epbCursor = eCursor_Hand
        End If
    ElseIf Annot.X(2) < Annot.X(1) Then
        If hitInfo.location = 9 And bTextEdit And hitInfo.itemIndex = 0 Then
            m.nActiveAnnotPt = 1
        ElseIf hitInfo.location = 10 Then
            m.nActiveAnnotPt = 2
        ElseIf hitInfo.location = 11 Then
            If Annot.eType = eANNOT_FibFan Or Annot.eType = eANNOT_SpResistFan Then
                If hitInfo.itemIndex = 0 Then
                    m.nActiveAnnotPt = 1
                Else
                    m.epbCursor = eCursor_Hand
                End If
            Else
                m.nActiveAnnotPt = 1
            End If
        Else
            m.epbCursor = eCursor_Hand
        End If
    Else
        If Annot.Y(2) = Annot.Y(1) Then
            'true when text edit is drawn using a single click
            'grapheng.dll assigns a y-value to the arrow when this occurs
            Annot.FixTextEditYs     'this is fix for aardvark 722
        End If
        
        If Annot.Y(2) > Annot.Y(1) Then
            If hitInfo.location = 10 Then
                m.nActiveAnnotPt = 2
            ElseIf hitInfo.location = 11 Or _
               (hitInfo.location = 9 And bTextEdit And hitInfo.itemIndex = 0) Then
                m.nActiveAnnotPt = 1
            Else
                m.epbCursor = eCursor_Hand
            End If
        Else
            If hitInfo.location = 10 Or _
                (hitInfo.location = 9 And bTextEdit And hitInfo.itemIndex = 0) Then
                m.nActiveAnnotPt = 1
            ElseIf hitInfo.location = 11 Then
                m.nActiveAnnotPt = 2
            Else
                m.epbCursor = eCursor_Hand
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetTDActivePt", eGDRaiseError_Raise
    
End Sub

Private Sub SetAnForkActivePt(hitInfo As hittest_info)
On Error GoTo ErrSection:


    Dim rc&, X#
    Dim Annot As cAnnotation
    Dim cInfo As coordinate_info
    Dim aXs As New cGdArray
    Dim iLeft&, iRight&
    
    If hitInfo.location = 9 Then
        'aardvark 3336 fix
        Set Annot = m.Chart.Annots(hitInfo.ItemID)
        
        If Not Annot Is Nothing Then
            If Annot.eType = eANNOT_AndrewFork Or Annot.eType = eANNOT_ElliotTimeRatio Then
                m.epbCursor = eCursor_Hand
                m.nActiveAnnotPt = 0
            ElseIf Annot.eType = eANNOT_WaveLabels And hitInfo.annType = 0 Then
                m.epbCursor = eCursor_Hand          'aardvark 3777
                m.nActiveAnnotPt = 0
            Else
                cInfo.paneId = hitInfo.paneId
                cInfo.x_pixels = hitInfo.x_pixels
                cInfo.y_pixels = hitInfo.y_pixels
                
                rc = geCoordToData(m.Chart.geChartObj, cInfo)
                If rc = 0 Then
                    X = cInfo.x_value
                    
                    aXs.Add Annot.geGetX(0)
                    aXs.Add Annot.geGetX(2)
                    aXs.Add Annot.geGetX(1)
                    If Annot.eType = eANNOT_ChannelHighlight Then aXs.Add Annot.geGetX(3)
                    
                    aXs.Sort
                                        
                    If Annot.eType = eANNOT_ElliotTimeRatio Then
                        iLeft = aXs(0) - 1
                        iRight = aXs(aXs.Size - 1) + 1
                    Else
                        iLeft = aXs(0)
                        iRight = aXs(aXs.Size - 1)
                    End If
                    
                    'If X > aXs(0) And X < aXs(aXs.Size - 1) Then
                    If X > iLeft And X < iRight Then
                        m.epbCursor = eCursor_Hand
                        m.nActiveAnnotPt = 0
                    Else
                        ClearAnnotFlags False
                        m.Chart.SetCursor
                    End If
                End If
            End If
        End If
        
        Exit Sub
    End If
    
    Set Annot = m.Chart.Annots(hitInfo.ItemID)
    If Annot Is Nothing Then
        ClearAnnotFlags False
        m.Chart.SetCursor
    ElseIf Annot.eType = eANNOT_ElliotTimeRatio And hitInfo.itemIndex > 4 Then
        ClearAnnotFlags False
        m.Chart.SetCursor               'don't want 4-way arrow for ratio lines
    Else
        m.epbCursor = eCursor_Arrow4Way
        
        cInfo.paneId = hitInfo.paneId
        cInfo.x_pixels = hitInfo.x_pixels
        cInfo.y_pixels = hitInfo.y_pixels
        
        rc = geCoordToData(m.Chart.geChartObj, cInfo)
        If rc = 0 Then
            X = cInfo.x_value + m.Chart.ScreenStartX
            m.nActiveAnnotPt = Annot.ClosestPoint(m.Chart, X, cInfo.y_value)
        Else
            ClearAnnotFlags False
            m.Chart.SetCursor
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetAnForkActivePt", eGDRaiseError_Raise
    
End Sub

Private Sub SetCycleActivePt(hitInfo As hittest_info)
On Error GoTo ErrSection:

    Dim Annot As cAnnotation

    Set Annot = m.Chart.Annots(hitInfo.ItemID)
    If Annot Is Nothing Then
        ClearAnnotFlags False
        Exit Sub
    End If
    
    If hitInfo.location = 11 Then
        'dash for resizing left/right
        m.epbCursor = eCursor_ArrowEW
        m.nActiveAnnotPt = 4
    ElseIf hitInfo.location = 10 Then
        'end point at top for resizing arc height
        m.epbCursor = eCursor_Arrow4Way
        m.nActiveAnnotPt = 1
    Else
        m.epbCursor = eCursor_Hand
        m.nActiveAnnotPt = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetCycleActivePt", eGDRaiseError_Raise
    
End Sub

Private Sub SetDnExActivePt(Annot As cAnnotation, hitInfo As hittest_info)
On Error GoTo ErrSection:

    If Annot Is Nothing Then Exit Sub           'precautionary
    
    If Annot.geMoveFlag <> 1 Then Exit Sub
    
    If UseDiNapFib() And Annot.eType = eANNOT_DNExpansion Then
        m.epbCursor = eCursor_Hand
        Exit Sub
    End If
    
    If Len(g.strActiveDraw) = 0 Then m.nPointCount = -1
    
    Select Case Annot.eType
        Case eANNOT_GannacciSwing1
            If Annot.X(2) > Annot.X(1) Then
                If hitInfo.location = 10 Then
                    m.nActiveAnnotPt = 1
                    m.nPointCount = 0
                ElseIf hitInfo.location = 11 Then
                    m.nActiveAnnotPt = 2
                    m.nPointCount = 1
                Else
                    m.nActiveAnnotPt = 0
                    m.nPointCount = -1
                End If
            ElseIf Annot.X(2) < Annot.X(1) Then
                If hitInfo.location = 10 Then
                    m.nActiveAnnotPt = 2
                    m.nPointCount = 1
                ElseIf hitInfo.location = 11 Then
                    m.nActiveAnnotPt = 1
                    m.nPointCount = 0
                Else
                    m.nActiveAnnotPt = 0
                    m.nPointCount = -1
                End If
            End If
            
            If hitInfo.location = 10 Or hitInfo.location = 11 Then
                m.epbCursor = eCursor_Arrow4Way
            Else
                m.epbCursor = eCursor_Hand
            End If
        
        Case eANNOT_GannacciSwing2
            If hitInfo.itemIndex = 0 Then
                If hitInfo.location = 10 Then
                    If Annot.X(1) <= Annot.X(2) Then
                        m.nActiveAnnotPt = 1
                        m.nPointCount = 0
                    Else
                        m.nActiveAnnotPt = 2
                        m.nPointCount = 1
                    End If
                ElseIf hitInfo.location = 11 Then
                    If Annot.X(1) <= Annot.X(2) Then
                        m.nActiveAnnotPt = 2
                        m.nPointCount = 1
                    Else
                        m.nActiveAnnotPt = 1
                        m.nPointCount = 0
                    End If
                Else
                    m.nActiveAnnotPt = 0
                    m.nPointCount = -1
                End If
            ElseIf hitInfo.itemIndex = 1 Then
                If hitInfo.location = 10 Then
                    If Annot.X(1) <= Annot.X(2) Then
                        m.nActiveAnnotPt = 2
                        m.nPointCount = 1
                    Else
                        m.nActiveAnnotPt = 3
                        m.nPointCount = 2
                    End If
                ElseIf hitInfo.location = 11 Then
                    m.nActiveAnnotPt = 3
                    m.nPointCount = 2
                Else
                    m.nActiveAnnotPt = 0
                    m.nPointCount = -1
                End If
            End If

            If hitInfo.location = 10 Or hitInfo.location = 11 Then
                m.epbCursor = eCursor_Arrow4Way
            Else
                m.epbCursor = eCursor_Hand
            End If
        
        Case Else
            If hitInfo.itemIndex = 1 Then
                'point B can have text to left of circle if it is on same bar as C
                If Annot.dDate(2) = Annot.dDate(1) Or Annot.dDate(2) = Annot.DateFromArray(0) Then
                    'drawn from right to left
                    If hitInfo.location = 10 Then
                        'this is true when drawn with 1st & 2nd points on same bar
                        'then going back in time, i.e. left, before choosing 3rd pt
                        m.epbCursor = eCursor_Arrow4Way
                        m.nActiveAnnotPt = 999999
                    ElseIf hitInfo.location = 11 Then       'And Annot.Prop("FreeFloat") = 1 Then
                        m.epbCursor = eCursor_Arrow4Way
                        m.nActiveAnnotPt = 2
                        m.nPointCount = 1
                    Else
                        m.epbCursor = eCursor_Hand
                    End If
                ElseIf hitInfo.location = 11 Then
                    m.epbCursor = eCursor_Arrow4Way     'drawn left to right
                    m.nActiveAnnotPt = 12
                ElseIf hitInfo.location = 10 Then           'And Annot.Prop("FreeFloat") = 1 Then
                    m.epbCursor = eCursor_Arrow4Way
                    m.nActiveAnnotPt = 2
                    m.nPointCount = 1
                Else
                    m.epbCursor = eCursor_Hand
                End If
            ElseIf hitInfo.location = 11 Then
                m.epbCursor = eCursor_Arrow4Way
                m.nActiveAnnotPt = hitInfo.itemIndex + 11
            ElseIf hitInfo.location = 10 And (hitInfo.itemIndex = 0 Or hitInfo.itemIndex = 2) Then
                'itemIndex 0 = A, itemIndex 2 = C
                m.epbCursor = eCursor_Arrow4Way
                m.nActiveAnnotPt = hitInfo.itemIndex + 1
                m.nPointCount = hitInfo.itemIndex
            Else
                m.epbCursor = eCursor_Hand
            End If
    End Select
        
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".SetDnExActivePt"
    
End Sub

Private Sub SetFibActivePt(hitInfo As hittest_info)
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    Dim bShowDiag As Boolean

    Set Annot = m.Chart.Annots(hitInfo.ItemID)
    If Annot Is Nothing Then Exit Sub
    
    Select Case Annot.eType
        Case eANNOT_FibTimeZones, eANNOT_DanCodeZone
            If hitInfo.location = 10 Then
                m.epbCursor = eCursor_Arrow4Way
                m.nActiveAnnotPt = 1
            Else
                m.epbCursor = eCursor_Hand
            End If
            
            Exit Sub
        
        Case eANNOT_FibArcs
            'hit location: 8=line, 9=text, 10=left endpoint, 11=right endpoint
            If hitInfo.itemIndex = 5 Then
                Select Case hitInfo.location
                    Case 8, 9       '8=line, 9=text
                        m.epbCursor = eCursor_Hand
                        m.nActiveAnnotPt = 0
                    Case 10         'left endpoint
                        m.epbCursor = eCursor_Arrow4Way
                        If Annot.X(1) < Annot.X(2) Then
                            m.nActiveAnnotPt = 1
                        Else
                            m.nActiveAnnotPt = 2
                        End If
                     
                    Case 11
                        m.epbCursor = eCursor_Arrow4Way
                        If Annot.X(1) < Annot.X(2) Then
                            m.nActiveAnnotPt = 2
                        Else
                            m.nActiveAnnotPt = 1
                        End If
                        
                    Case Else
                        ClearAnnotFlags False, False
                End Select
            Else
                m.epbCursor = eCursor_Hand
                m.nActiveAnnotPt = 0
            End If
            
            Exit Sub
    End Select
    
    
    Dim rc&, idx&, location&
    Dim dyDiff#, dyDiff2#
    Dim cInfo As coordinate_info
    Dim cInfo2 As coordinate_info
    
    cInfo.paneId = Annot.gePaneId
    cInfo.x_pixels = hitInfo.x_pixels
    cInfo.y_pixels = hitInfo.y_pixels
    
    rc = geCoordToData(m.Chart.geChartObj, cInfo)
    If rc <> 0 Then Exit Sub
    
    dyDiff = Abs(cInfo.y_value - Annot.Y(1))
    dyDiff2 = Abs(cInfo.y_value - Annot.Y(2))
    
    'idx identifies which line was hit, location identifies which part of line was hit
    idx = hitInfo.itemIndex
    location = hitInfo.location
    
    If Annot.eType = eANNOT_FibTimeRatio Then
        If (idx = 2 Or idx = 3) And (location = 10 Or location = 11) Then
            'item index 2,3 = vert line 0/1
            'location 10, 11 = endpoints of vert line
            If idx = 2 Then
                If dyDiff < dyDiff2 Then
                    m.nActiveAnnotPt = 3
                Else
                    m.nActiveAnnotPt = 5
                End If
            Else
                If dyDiff < dyDiff2 Then
                    m.nActiveAnnotPt = 4
                Else
                    m.nActiveAnnotPt = 6
                End If
            End If
            m.epbCursor = eCursor_Arrow4Way         '4-way arrow for resizing
        Else
            m.nActiveAnnotPt = 0
            m.epbCursor = eCursor_Hand
        End If
        
        Exit Sub
    End If

'remaining code is for fib supp/resist, fib expansion & dancode fib
'originally idx=1 indicates hit on vertical line on price fibs (changed to diagonal line in 5/2012)
    If idx = 1 Then
        If Annot.X(1) = Annot.X(2) Or location = 9 Then     'fix for 6692
            m.epbCursor = eCursor_Hand
            cInfo.paneId = Annot.gePaneId
            cInfo.x_value = Annot.X(1)
            cInfo.y_value = Annot.Y(1)
            rc = geDataToCoord(m.Chart.geChartObj, cInfo)
            If rc = 0 Then
                If Abs(cInfo.y_pixels - hitInfo.y_pixels) <= 5 Then
                    m.epbCursor = eCursor_Arrow4Way
                    m.nActiveAnnotPt = 3
                Else
                    cInfo.x_value = Annot.X(2)
                    cInfo.y_value = Annot.Y(2)
                    rc = geDataToCoord(m.Chart.geChartObj, cInfo)
                    If rc = 0 Then
                        If Abs(cInfo.y_pixels - hitInfo.y_pixels) <= 5 Then
                            m.epbCursor = eCursor_Arrow4Way
                            m.nActiveAnnotPt = 6
                        End If
                    End If
                End If
            End If
        ElseIf location = 10 Then
            m.epbCursor = eCursor_Arrow4Way
            If Annot.dDate(1) < Annot.dDate(2) Then
                m.nActiveAnnotPt = 3
            Else
                m.nActiveAnnotPt = 6
            End If
        ElseIf location = 11 Then
            m.epbCursor = eCursor_Arrow4Way
            If Annot.dDate(1) < Annot.dDate(2) Then
                m.nActiveAnnotPt = 6
            Else
                m.nActiveAnnotPt = 3
            End If
        End If
    ElseIf idx = 2 Or idx = 3 Then
        'horizontal 0/1 lines
        Select Case location
            Case 2, 6
                m.epbCursor = eCursor_ArrowNS
                bShowDiag = True
                If dyDiff < dyDiff2 Then
                    m.nActiveAnnotPt = 1
                Else
                    m.nActiveAnnotPt = 2
                End If
            Case 10, 11

'JM (06-13-2012): original code, do not think is needed, leave awhile then remove if all ok
'                If Annot.Prop("Ext") = 4 Then
'                    'only extension right is shown, do not show hit for left end point
'                    m.epbCursor = eCursor_ArrowEW
'                    bShowDiag = True
'                    m.nActiveAnnotPt = 1
'                    ClearAnnotFlags False, False
'                    m.Chart.SetCursor
'                Else
                    Dim nBar&, dxDate#
                    
                    dxDate = gdGetNum(m.Chart.geDateArray, cInfo.x_value)
                    If m.Chart.aXdate.BinarySearch(gdFixDateTime(dxDate), nBar) Then
                        m.epbCursor = eCursor_Arrow4Way
                        bShowDiag = True
                        If idx = 2 Then
                            If nBar = Annot.X(1) Then
                                m.nActiveAnnotPt = 3
                            ElseIf nBar = Annot.X(2) Then
                                m.nActiveAnnotPt = 4
                            End If
                        ElseIf nBar = Annot.X(1) Then
                            m.nActiveAnnotPt = 5
                        ElseIf nBar = Annot.X(2) Then
                            m.nActiveAnnotPt = 6
                        End If
                    End If
'                End If
            Case 8
                m.epbCursor = eCursor_Hand      'text
                bShowDiag = True
            Case 9
                If Val(Annot.Prop("Ext")) = 4 Or Annot.X(1) = Annot.X(2) Then
                    'only extension right is shown or diag is perfectly vertical
                    ClearAnnotFlags False, False
                    m.Chart.SetCursor
                Else
                    m.epbCursor = eCursor_ArrowNS
                    bShowDiag = True
                    If dyDiff < dyDiff2 Then
                        m.nActiveAnnotPt = 1
                    Else
                        m.nActiveAnnotPt = 2
                    End If
                End If
            Case Else
                ClearAnnotFlags False, False
                m.Chart.SetCursor
        End Select
    Else
        'ratio lines or text (aardvark 6691 fix)
        Select Case location
            Case 2, 6, 8, 9
                m.epbCursor = eCursor_Hand
            Case 10
                m.epbCursor = eCursor_ArrowEW
                If Annot.X(1) < Annot.X(2) Then
                    m.nActiveAnnotPt = 7
                Else
                    m.nActiveAnnotPt = 8
                End If
            Case 11
                m.epbCursor = eCursor_ArrowEW
                If Annot.X(1) < Annot.X(2) Then
                    m.nActiveAnnotPt = 8
                Else
                    m.nActiveAnnotPt = 7
                End If
            Case Else
                ClearAnnotFlags False, False
        End Select
        bShowDiag = True
    End If

    If bShowDiag Then
        'Pete's email (03-22-2012): when hovering over any line, show where you drew 2 points with dashed diagonal line
        If Val(Annot.Prop("HideVerticalLine")) = 1 Then
            If m.AnnotOptions Is Nothing Then
                Annot.Prop("HideVerticalLine") = -1
                Chart.GenerateChart eRedo1_Scrolled
                Set m.AnnotOptions = Annot
            End If
        End If
    End If

ErrExit:
    Set Annot = Nothing
    Exit Sub
    
ErrSection:
    Set Annot = Nothing
    RaiseError Me.Name & ".SetFibActivePt", eGDRaiseError_Raise
    
End Sub

Private Sub SetTgtShooterActivePt(hitInfo As hittest_info)
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    Dim cInfo As coordinate_info
    Dim rc&, dDiff#, dDiff2#
    Dim dTop#, dxLeft#, dBottom#, dxRight#
    
    m.epbCursor = eCursor_Hand
    Set Annot = m.Chart.Annots(hitInfo.ItemID)
    If Annot Is Nothing Then Exit Sub
    
    cInfo.paneId = Annot.gePaneId
    cInfo.x_pixels = hitInfo.x_pixels
    cInfo.y_pixels = hitInfo.y_pixels
    rc = geCoordToData(m.Chart.geChartObj, cInfo)
    
    If rc <> 0 Then
        Set Annot = Nothing
        Exit Sub
    End If
        
    If hitInfo.itemIndex = 0 Then
        m.epbCursor = eCursor_Arrow4Way
        If hitInfo.location = 10 Or hitInfo.location = 11 Then
            dDiff = Abs(cInfo.y_value - Annot.Y(1))
            dDiff2 = Abs(cInfo.y_value - Annot.Y(2))
            If dDiff < dDiff2 Then
                m.nActiveAnnotPt = 1
            Else
                m.nActiveAnnotPt = 2
            End If
        Else
            m.epbCursor = eCursor_Hand
        End If
    ElseIf hitInfo.itemIndex > 5 And hitInfo.itemIndex < 12 Then
        m.epbCursor = eCursor_ArrowEW
        Annot.geGetDim dTop, dxLeft, dBottom, dxRight, hitInfo.itemIndex
        If hitInfo.location = 10 Or hitInfo.location = 11 Then
            If Abs(cInfo.x_value - dxRight) < 2 Then
                m.nActiveAnnotPt = 10
            Else
                m.epbCursor = eCursor_Hand
            End If
        Else
            m.epbCursor = eCursor_Hand
        End If
    End If
    
ErrExit:
    Set Annot = Nothing
    Exit Sub
    
ErrSection:
    Set Annot = Nothing
    RaiseError Me.Name & ".SetTgtShooterActivePt", eGDRaiseError_Raise
    
End Sub

Public Sub ClearAnnotFlags(ByVal bDelAnnotInprog As Boolean, _
        Optional ByVal bResetTools As Boolean = True)
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If m.Chart Is Nothing Then Exit Sub
    If m.Chart.Annots Is Nothing Then Exit Sub
    
    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    If bDelAnnotInprog = True Then
        If Not Annot Is Nothing Then
            If Annot.eType = eANNOT_DNExpansion Or Annot.eType = eANNOT_DNExpansion2 Or _
               Annot.eType = eANNOT_DNExpansion3 Or Annot.eType = eANNOT_DNExpansion4 Then
                If m.nPointCount < 3 Then
                    m.Chart.RemoveAnnots True
                Else
                    Annot.geMoveFlag = 1    'signals user done with drawing
                End If
            ElseIf Annot.eType = eANNOT_DNRetracement Then
                If m.nPointCount < 2 Then
                    m.Chart.RemoveAnnots True
                Else
                    Annot.geMoveFlag = 1
                End If
            Else
                m.Chart.RemoveAnnots True
            End If
            If Len(g.strActiveDraw) > 0 Then            '6283
                g.strActiveDraw = ""
                SyncDrawTools True
            End If
            pbChart.Refresh
        End If
    End If
    If Not Annot Is Nothing Then
        If Annot.eType = eANNOT_RegressionLine Then Annot.geMoveFlag = 1
        Annot.HitItemIndex = -1
        Annot.MenuAdd = False
        Annot.MenuMove = False
    End If
    
    m.nActiveAnnotIdx = 0
    m.nActiveAnnotPt = -1
    m.bAnnotCreated = False

    m.nFocusHiLo = -1
    m.nPointCount = 0

    Set Annot = Nothing
    
    If bResetTools = True Then
        If Len(g.strActiveDraw) > 0 And g.strActiveDraw <> "ID_Icon" And _
            g.strActiveDraw <> "ID_ElliotLabels" And g.strActiveDraw <> "ID_ElliotEndUser" And _
            frmMain.tbToolbar.Tools("ID_RepeatDraw").State = ssUnchecked Then
            ToolbarSetCursorGroup frmMain.tbToolbar, False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ClearAnnotFlags", eGDRaiseError_Raise
    
End Sub

Private Sub DoHitTest(ByVal X As Single, ByVal Y As Single, _
    ByVal bOnMousedown As Boolean, ByVal Button As Long)
On Error GoTo ErrSection:

    Dim rc&, idx&, hitInfo As hittest_info
    Dim Ind As cIndicator
    Dim Annot As cAnnotation
    Dim bExit As Boolean
    Dim OrderStruct As cOrderStruct
        
    If m.nActiveAnnotIdx < 0 Then
        bExit = True
    ElseIf m.nActiveBtmPane > 0 And m.epbCursor = eCursor_HorzSize Then
        bExit = True 'pane separator dragging inprogress
    ElseIf Len(g.strActiveDraw) > 0 And g.strActiveDraw <> "ID_Icon" And _
        g.strActiveDraw <> "ID_ElliotLabels" And g.strActiveDraw <> "ID_ElliotEndUser" Then
        bExit = True  'annotations draw in progress
    End If
    
    If bExit Then
        ChartTips eTiptype_None
        GoTo ErrExit
    End If
    
    m.nActiveAnnotIdx = 0
    m.nActiveIndIdx = 0
    m.nActiveTopPane = 0
    m.nActiveBtmPane = 0
    m.nActiveOrderID = 0
    m.nActiveOrderLoc = 0
    
    hitInfo.paneId = -1
    hitInfo.x_pixels = X \ Screen.TwipsPerPixelX
    hitInfo.y_pixels = Y \ Screen.TwipsPerPixelY
    hitInfo.location = -1
    
    If m.epbCursor = eCursor_OrderBuy Or m.epbCursor = eCursor_OrderSell Then
        GoTo ErrExit        'don't want to bring up annot editor or anything like that
    End If
    
    rc = geHitTest(m.Chart.geChartObj, hitInfo, pbChart.hDC)
    If rc = 0 And Button = vbLeftButton Then
        If bOnMousedown = True Then
            If frmMain.tbToolbar.Tools("ID_ChartMove").State = ssChecked Then
                HandleChartMove 0, 0, False, m.epbCursor
                m.epbCursor = eCursor_ChartMove
                m.bChartMoveInProg = True
            Else
                m.epbCursor = eCursor_NoDrop
            End If
        Else
            m.Chart.SetCursor
        End If
        ChartTips eTiptype_None
        GoTo ErrExit
    End If
    
    
'Aardvark 5963 Note:
'original fix on 01-31-2014 took away ability for EWI analyst to delete/move
'label while working on a new wave (see Rob's email in aardvark)
'code change on 02-27-2014 restores this functionality
    If FormIsLoaded("frmIconAnnot") Or FormIsLoaded("frmElliot") Then
        If Len(g.strActiveDraw) > 0 Then m.epbCursor = eCursor_Pencil
        If hitInfo.itemType = 4 Then
            'JM 02-27-2014: code change (5963)
            Set Annot = m.Chart.Annots(hitInfo.ItemID)
            If Annot.eType <> eANNOT_ElliotLabel Then
                Set Annot = Nothing
            End If
        End If
        If Annot Is Nothing Then
            GoTo ErrExit
        End If
    End If
    
    'pane separators
    If hitInfo.itemType = 2 Then
        If bOnMousedown = True Then
            If hitInfo.btmPaneId > -1 Then
                m.nActiveTopPane = hitInfo.topPaneId
                m.nActiveBtmPane = hitInfo.btmPaneId
            End If
        End If
        'check for eraser mode
        If g.ChartGlobals.eChartMode = eMode_Erase Then
            m.epbCursor = eCursor_Eraser
        Else
            m.epbCursor = eCursor_HorzSize
        End If
        ChartTips eTiptype_None
        GoTo ErrExit
    End If
    
    'indicators
    If hitInfo.itemType = 3 Then 'And Button = vbLeftButton Then
        If m.Chart.Tree.NodeLevel(hitInfo.ItemID) > 0 Then
            Set Ind = m.Chart.Tree(hitInfo.ItemID)
        Else
            Set Ind = Nothing
        End If
                
        rc = 0
        If Button = vbLeftButton Then rc = frmMain.tbToolbar.Tools("ID_ChartMove").State
                
        If Ind Is Nothing Then
            'do nothing
        ElseIf Ind.DataType = eINDIC_BarData Then
            m.nActiveIndIdx = hitInfo.ItemID
        ElseIf Ind.DataType = eINDIC_Constant Then
            rc = 0
            m.nActiveIndIdx = hitInfo.ItemID
            m.epbCursor = eCursor_Hand
            If bOnMousedown = True Then m.nObjectMoving = 1
        ElseIf Ind.DisplayType = eINDIC_ClusterTime Then
            rc = 0
            m.nActiveIndIdx = hitInfo.ItemID
            m.epbCursor = eCursor_Hand
            m.nObjectMoving = 0
        End If
        
        If bOnMousedown And rc = ssChecked Then
           m.epbCursor = eCursor_ChartMove
            m.bChartMoveInProg = False
        End If
        ChartTips eTiptype_None
        GoTo ErrExit
    End If
    
    'orders
    If hitInfo.itemType = 5 Then
        m.nActiveOrderID = hitInfo.ItemID
        m.nActiveOrderLoc = hitInfo.location
        If g.ChartGlobals.bChartTips Then
            If hitInfo.location = 1 Or hitInfo.location = 2 Then
                Set OrderStruct = m.Chart.OnlineOrders(Str(hitInfo.ItemID))
                If Not OrderStruct Is Nothing Then
                    OrderStruct.AlertTip m.oToolTip, pbChart, hitInfo.location
                End If
            End If
        End If
        If bOnMousedown Then m.nObjectMoving = 1
        m.epbCursor = eCursor_Hand
        ShowCursor
        GoTo ErrExit
    End If
        
    'do additional processing if hit item was an annotation
    If hitInfo.itemType <> 4 Then
        ChartTips eTiptype_None
        GoTo ErrExit
    End If
            
    'annotations
    m.nActiveAnnotPt = 0
    Set Annot = m.Chart.Annots(hitInfo.ItemID)
    If Annot Is Nothing Then
        GoTo ErrExit
    End If
    
    Annot.HitItemIndex = hitInfo.itemIndex
    If bOnMousedown = True Then
        m.Chart.LastEditCreate Annot, m.bAnnotCreated
        
        If Annot.eUsage = eANNOT_FibClusters And hitInfo.location <> 2 And hitInfo.location <> 6 Then
            'show zoom of move cursor if user is inside cluster zone rect
            ClearAnnotFlags True, False
            If g.ChartGlobals.eChartMode = eMode_Move Then
                m.epbCursor = eCursor_ChartMove
            ElseIf g.ChartGlobals.eChartMode = eMode_Zoom Then
                m.epbCursor = eCursor_NoDrop
            End If
            GoTo ErrExit
        Else
            m.nActiveAnnotIdx = hitInfo.ItemID
            m.nObjectMoving = 1
        End If
    End If
    
    'location 8=center, 9=on line, 10=left endpoint 11=rightendpoint
    Select Case Annot.eType
        Case eANNOT_Icon, eANNOT_ElliotLabel, eANNOT_GannacciCycle
            If Not MouseIsPressed Then
                If Annot.eUsage <> eANNOT_BidAskChart Then
                    m.epbCursor = eCursor_Hand
                    If Annot.eUsage = eANNOT_Trades Then
                        idx = Val(Parse(m.Chart.Annots.Key(Annot.geAnnId), " ", 2))
                        ChartTips eTipType_Trade, idx
                    ElseIf Annot.eUsage = eANNOT_Notation Or Annot.eUsage = eANNOT_IndicatorLabel Then
                        ChartTips eTipType_Annot, hitInfo.ItemID
                    End If
                End If
            End If
        Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4
            m.epbCursor = eCursor_Hand
            If hitInfo.itemIndex > 0 Then
                Annot.AlertTip m.oToolTip, pbChart, hitInfo.itemIndex
            Else
                ChartTips eTipType_Annot, hitInfo.ItemID
            End If
        Case eANNOT_VertLine
            If Annot.eUsage <> eANNOT_OptionInfo Then m.epbCursor = eCursor_Hand
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
             eANNOT_DNExpansion4, eANNOT_FibABCD, eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
            SetDnExActivePt Annot, hitInfo
            ChartTips eTipType_Annot, hitInfo.ItemID, hitInfo.itemIndex         '6336
        Case eANNOT_DNRetracement
            m.epbCursor = eCursor_Hand
        Case eANNOT_Rectangle, eANNOT_Mirror, eANNOT_Bracket
            SetRectActivePt Annot, hitInfo
        Case eANNOT_Pattern
            SetPatternActivePt Annot, hitInfo
        Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, eANNOT_TrendChannel, _
             eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4, _
             eANNOT_RegressionLine, eANNOT_Ellipse, eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, _
             eANNOT_TextEdit4, eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4, _
             eANNOT_FibFan, eANNOT_SpResistFan, eANNOT_Pivot, _
             eANNOT_ArrowLine, eANNOT_GannacciSwingSquare, eANNOT_BalloonStrangle
            
            If hitInfo.annType = 7 Then
                m.epbCursor = eCursor_Hand
                Annot.AlertTip m.oToolTip, pbChart, hitInfo.itemIndex
            Else
                SetTDActivePt Annot, hitInfo
            End If
        
        Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, eANNOT_Fibonacci4, _
            eANNOT_FibTimeRatio, eANNOT_FibTimeZones, eANNOT_FibArcs, _
            eANNOT_FibExpansion, eANNOT_DanCodeFib, eANNOT_DanCodeZone, _
            eANNOT_AdvRiskReward
            SetFibActivePt hitInfo
            ChartTips eTipType_Annot, hitInfo.ItemID, hitInfo.itemIndex
        Case eANNOT_TargetShooter
            SetTgtShooterActivePt hitInfo
            ChartTips eTipType_Annot, hitInfo.ItemID, hitInfo.itemIndex
        Case eANNOT_GannLines
            SetGannActivePt hitInfo
        Case eANNOT_TimeCycle
            SetCycleActivePt hitInfo
        Case eANNOT_AndrewFork, eANNOT_TriangleWedge, eANNOT_ChannelHighlight, _
            eANNOT_WaveLabels, eANNOT_ElliotTimeRatio
            SetAnForkActivePt hitInfo
            If Annot.eType = eANNOT_ElliotTimeRatio Or eANNOT_FibTimeRatio Then
                ChartTips eTipType_Annot, hitInfo.ItemID, hitInfo.itemIndex
            End If
        Case eANNOT_Gartley
            SetGartleyActivePt Annot, hitInfo
        Case eANNOT_RiskReward
            SetRiskRewardActivePt Annot, hitInfo
        Case eANNOT_BellAlert
            m.epbCursor = eCursor_Hand
            Annot.AlertTip m.oToolTip, pbChart, hitInfo.itemIndex
        Case eANNOT_SimpleLine
            If Annot.eUsage = eANNOT_WhatIf Then pbCursor = eCursor_ArrowNS
        Case eANNOT_GannacciTime
            m.epbCursor = eCursor_Hand
    End Select
    
    'check for eraser mode
    If g.ChartGlobals.eChartMode = eMode_Erase Then
        m.epbCursor = eCursor_Eraser
    End If

ErrExit:
    If Not m.AnnotOptions Is Nothing Then
        If Not Annot Is m.AnnotOptions Then
            Select Case m.AnnotOptions.eType
                Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                     eANNOT_Fibonacci4, eANNOT_FibExpansion, eANNOT_DanCodeFib, _
                     eANNOT_AdvRiskReward
                    If Val(m.AnnotOptions.Prop("HideVerticalLine")) = 1 Or Val(m.AnnotOptions.Prop("Ext")) = 4 Then
                        m.Chart.GenerateChart eRedo1_Scrolled           '6728 - turn off diagonal
                        Set m.AnnotOptions = Nothing
                    End If
            End Select
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".DoHitTest", eGDRaiseError_Raise
    
End Sub

Private Sub AddNewIndicator(ByVal nType&, ByVal bCtrlClk As Boolean)
On Error GoTo ErrSection:

    'using enum type from grapheng
    '0=points,1=line,2=hz line,3=area,4=histogram,5=candle,6=OHLC,7=HLC,8=HL
    'TOTHINK: consider moving this and addnewannot sub to chart class
    
    Dim Pane As cPane, Ind As cIndicator
    
    'only processing hz line for now
    If nType <> 2 Then Exit Sub
    
    Set Pane = m.Chart.Tree("PRICE PANE")
    If Pane Is Nothing Then Exit Sub
    
    'The following two checks are fixes for aardvark issue 511 (ctrl-click in y-scale area)
    If m.MouseLast.bOffChart Then Exit Sub
    If m.Chart.Tree.NodeLevel(m.MouseLast.nPaneID) <> 0 Then Exit Sub
    
    If Pane.gePaneId = m.MouseLast.nPaneID And bCtrlClk = True Then
        AddNewAnnot True, "ID_HorzLine"
        Exit Sub
    End If
    
    'create new indicator
    Set Ind = New cIndicator
    With Ind
        .Name = "Horizontal Line"
        .Display = True
        .DisplayType = eINDIC_Line      'aardvark fix issue 961
        .DataType = eINDIC_Constant
        .Style = eINDIC_Default
        .Color = RGB(128, 128, 128)
        .Parm(0) = m.MouseLast.strRoundedY
    End With
    'add to pane (make it the last child)
    m.Chart.Tree.Add Ind, "", m.MouseDown.nPaneID, eTREE_LastChild
    m.Chart.GenerateChart
    
ErrExit:
    Set Pane = Nothing
    Exit Sub
    
ErrSection:
    Set Pane = Nothing
    RaiseError Me.Name & ".AddNewIndicator", eGDRaiseError_Raise
    
End Sub

Private Sub AddNewAnnot(ByVal bOnMouseup As Boolean, Optional ByVal strAnnotType$ = "")
On Error GoTo ErrSection:

    Dim i&, nIndIdx&, rc&, strKey$, strSave$
    Dim dY#, dPoints#, dHigh#, dLow#
    
    Dim bExit As Boolean
    Dim bDynamicFib As Boolean
    
    Dim Alert As cAlert
    Dim Bars As cGdBars
    Dim Pane As cPane
    
    Dim eType As eAnnotType
    Dim Annot As cAnnotation
    
    eType = eANNOT_UndefinedType
    m.bAnnotCreated = False

    If g.strActiveDraw = "ID_PriceAlert" Then
        'this is not a drawing tool tied to a chart so skip all normal processing for drawing tools
        'just create the price alert and let the global alerts collection handle the rest
        If Not m.Chart Is Nothing Then
            If Len(m.Chart.SpreadSymbols) > 0 Then
                strSave = g.strActiveDraw
                g.strActiveDraw = "ID_HorzLine"
                strAnnotType = "ID_HorzLine"
            ElseIf Len(m.Chart.ExternalData) > 0 Then
                InfBox "Price alert not valid for external data.", "I"
                m.nPointCount = 2
                GoTo ErrExit
            Else
                Set Alert = New cAlert
                Alert.Symbol = m.Chart.Symbol
                Alert.Period = "Daily"
                Alert.GetsUpToPrice = m.MouseDown.dY
                Alert.GetsDownToPrice = m.MouseDown.dY
                Alert.ShowOnCharts = True
                
                Set Bars = m.Chart.Bars
                If Not Bars Is Nothing Then
                    i = m.Chart.LastGoodDataBar(False, False)
                    dY = Bars(eBARS_Close, i)
                    If m.MouseDown.dY > dY Then
                        Alert.UseGetsUpTo = True
                    ElseIf m.MouseDown.dY < dY Then
                        Alert.UseGetsDownTo = True
                    End If
                End If
                
                If frmAlerts.ShowMe(Alert, eGDAlertType_Price) Then
                    g.Alerts.Add Alert
                    If FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid
                End If
                m.nPointCount = 2
                GoTo ErrExit
            End If
        End If
    End If
    
    ' add new annotation
    If Len(strAnnotType) = 0 Then strAnnotType = g.strActiveDraw
    
    'create temporary annotation to get type
    Set Annot = New cAnnotation
    eType = Annot.AnnotTypeFromToolID(strAnnotType)
    
    If eType = eANNOT_UndefinedType Then GoTo ErrExit
    
    Select Case eType
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
             eANNOT_DNExpansion4, eANNOT_DNRetracement, eANNOT_RegressionLine
            i = m.Chart.LastGoodDataBar(False)
            'These annotations cannot start past last good data bar or on empty bars
            If m.Chart.Bars(eBARS_DateTime, i) < m.MouseDown.dDate Or m.MouseDown.nBar < 0 Then
                Beep
                ClearAnnotFlags True, True
                GoTo ErrExit
            End If
        
        Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4, _
             eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4, _
             eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4, _
             eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
             eANNOT_Fibonacci4, eANNOT_FibExpansion, eANNOT_DanCodeFib, eANNOT_AdvRiskReward
            'these annots can be created with single click, click-click or click-n-drag
            'click-n-drag is processed with mouse down/mouse up and minmove
            'single click and click-click are mutually exclusive
            'if keep-at-end or extend to right is true from LoadDefaults then single-click ends draw
            'otherwise must use second click in click-click sequence to end draw
            Annot.CreateNew m.Chart, eType, 1, 0, 0, 0, 0, , , "", , True
    
    End Select
    
    If bOnMouseup = True Then
        bExit = True
        
        Select Case eType
            Case eANNOT_BalloonStrangle
                If m.Chart.Bars.IsIntraday Or m.Chart.Periodicity >= ePRD_Months + 1 Then
                    InfBox "The Risk Reward Visualizer tool can only be added to a Daily chart.", "I", , "Information"
                    ClearAnnotFlags True, True
                    m.Chart.GenerateChart eRedo1_Scrolled
                    bExit = True
                    If Not Me.MDIChild Then SyncDrawTools
                ElseIf SecurityType(m.Chart.Symbol) = "F" Then
                    InfBox "The Risk Reward Visualizer tool is not available for futures.", "I", , "Information"
                    ClearAnnotFlags True, True
                    m.Chart.GenerateChart eRedo1_Scrolled
                    bExit = True
                    If Not Me.MDIChild Then SyncDrawTools
                Else
                    m.nPointCount = 2
                    bExit = False
                End If
                
            Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4, _
                 eANNOT_VertLine, eANNOT_Icon, eANNOT_ElliotLabel, eANNOT_FibTimeZones, _
                 eANNOT_DanCodeZone, eANNOT_GannacciCycle, eANNOT_GannacciTime
                
                m.nPointCount = 2
                bExit = False
        
            Case eANNOT_DNRetracement
                If HandleDnRetrace(bOnMouseup, False, dY) = True Then bExit = False
        
            Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, eANNOT_DNExpansion4, _
                 eANNOT_FibABCD, eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
                If HandleDnExpansion(bOnMouseup, dY) = True Then bExit = False
        
            Case eANNOT_Gartley
                If HandleGartley(bOnMouseup, dY) = True Then bExit = False
        
            Case eANNOT_AndrewFork, eANNOT_ElliotTimeRatio
                bExit = Not HandleAndrewFork(bOnMouseup)
        
            Case eANNOT_ChannelHighlight, eANNOT_TriangleWedge
                bExit = Not HandleTriChannel(bOnMouseup, False)
        
            Case eANNOT_WaveLabels
                bExit = Not HandleWaveLabels(False)

            Case Else
                If m.nActiveAnnotIdx < 0 And Len(g.strActiveDraw) > 0 Then
                    
                    If m.nPointCount = 0 Then
                        Select Case Annot.eType
                            Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4
                                If Annot.Prop("KeepAtEnd") = 1 Then
                                    m.nPointCount = 2
                                    bExit = False
                                End If

                            Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
                                If Annot.Prop("Ext") = 1 Then
                                    m.nPointCount = 2
                                    bExit = False
                                End If

                            Case eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4
                                If Annot.Prop("ArrowStyle") = 0 Then
                                    m.nPointCount = 2
                                    bExit = False
                                End If

                            Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                                 eANNOT_Fibonacci4, eANNOT_FibExpansion, eANNOT_DanCodeFib, _
                                 eANNOT_AdvRiskReward
                                If Annot.Prop("KeepAlive") = 1 Then
                                    bDynamicFib = True
                                    bExit = False
                                End If
                        
                        End Select
                        
                        If bExit Then
                            StatusMsg "Now click on the second point ...", -1
                            m.nPointCount = 1
                        End If
                        
                    Else
                        bExit = True
                    End If
                End If

        End Select
    Else
        Select Case eType
            Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                 eANNOT_Fibonacci4, eANNOT_FibExpansion, eANNOT_DanCodeFib, _
                 eANNOT_AdvRiskReward
                If Annot.Prop("KeepAlive") = 1 Then bDynamicFib = True
        End Select
    End If
    
    Set Annot = Nothing
    
    If bExit = True Then GoTo ErrExit
    
    m.epbCursor = eCursor_Hand
    ShowCursor
    
    Set Annot = New cAnnotation
    
    Select Case eType
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, eANNOT_DNExpansion4
        Case eANNOT_DNRetracement, eANNOT_FibABCD, eANNOT_Gartley
        Case eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
            'do nothing
        
        Case eANNOT_GannacciCycle
            SnapToHiLoClose m.MouseDown, dY, 0

        Case Else
             dY = SnapToPrice(m.MouseDown, eType)
    End Select
    
    If bDynamicFib Then
        m.nPointCount = 2
    Else
        i = m.MouseDown.nPaneID
        If Len(strSave) > 0 And i = 0 Then i = m.Chart.Tree("PRICE").geIndpaneId
        m.nActiveAnnotPt = Annot.CreateNew(m.Chart, eType, i, _
                m.MouseDown.dDate, dY, m.MouseDown.dDate, m.MouseDown.dY, , , , , , m.bSchiffFork)
        If m.nActiveAnnotPt < 0 Then
            ClearAnnotFlags True, True
            m.Chart.SetCursor
            GoTo ErrExit
        End If
    
        Annot.X(1) = m.MouseDown.nX
    End If

    Select Case eType
        
        Case eANNOT_RegressionLine
            If m.Chart.Tree.Key(m.MouseDown.nPaneID) = "PRICE PANE" Then
                nIndIdx = m.Chart.Tree("PRICE").geIndId
            Else
                rc = geClosestIndicatorIdx(m.Chart.geChartObj, _
                    m.MouseDown.MouseX / Screen.TwipsPerPixelX, _
                    m.MouseDown.MouseY / Screen.TwipsPerPixelY, 0, 1, nIndIdx)
            End If
            If rc = 0 Then
                Annot.geIndId = nIndIdx
                Annot.Prop("IndicatorKey") = m.Chart.Tree.Key(Annot.geIndId)
            End If
        
        Case eANNOT_TrendChannel
            Annot.Prop("ChannelLocation") = 3
            Annot.Prop("ChannelType") = 1
            Annot.Prop("ChannelCount") = 1
            
            Set Pane = m.Chart.Tree(m.MouseDown.nPaneID)
            If Not Pane Is Nothing Then dPoints = (Pane.Max - Pane.Min) * 0.1
            Annot.Prop("ChannelPoints") = dPoints
        
        Case eANNOT_GannacciCycle
            If m.nFocusHiLo = 1 Then
                Annot.Prop("ImageDir") = eCNI_North     'closest to high of bar
            Else
                Annot.Prop("ImageDir") = eCNI_South     'closest to low of bar
            End If
    
        Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
             eANNOT_Fibonacci4, eANNOT_FibExpansion, eANNOT_DanCodeFib, _
             eANNOT_AdvRiskReward
            If bDynamicFib Then
                dY = m.MouseDown.dY
                dHigh = m.Chart.Bars(eBARS_High, m.MouseDown.nBar)
                dLow = m.Chart.Bars(eBARS_Low, m.MouseDown.nBar)
                
                dPoints = (dHigh + dLow) / 2
                
                If dY > dPoints Then
                    dY = dLow
                Else
                    dY = dHigh
                End If
                
                Annot.dDate(1) = m.MouseDown.dDate
                Annot.Y(1) = m.MouseDown.dY
                Annot.dDate(2) = m.Chart.Bars(eBARS_DateTime, m.MouseDown.nBar + 1)
                Annot.Y(2) = dY
            
                m.nActiveAnnotPt = Annot.CreateNew(m.Chart, eType, m.MouseDown.nPaneID, _
                        m.MouseDown.dDate, m.MouseDown.dY, Annot.dDate(2), dY, , , , , , m.bSchiffFork)
            
                If m.nActiveAnnotPt < 0 Then
                    ClearAnnotFlags True, True
                    m.Chart.SetCursor
                    GoTo ErrExit
                End If
                
                Annot.X(1) = m.MouseDown.nX
            End If
    
        Case eANNOT_WaveLabels
            Annot.CopyCreateOptions m.AnnotOptions
    
        Case eANNOT_Icon
            If FormIsLoaded("frmIconAnnot") Then
                If frmIconAnnot.Visible Then
                    frmIconAnnot.GetSettings Annot
                    frmIconAnnot.NextIcon
                End If
            End If
    
        Case eANNOT_ElliotLabel
            If FormIsLoaded("frmElliot") Then
                If frmElliot.Visible Then
                    frmElliot.GetSettings Annot
                    
                    If g.strActiveDraw = "ID_ElliotEndUser" Then
                        Annot.Prop("EndUserEWI") = 1
                    End If
                End If
            End If
    
        Case eANNOT_Rectangle
            If g.strActiveDraw = "ID_FibClusters" Then
                Annot.eUsage = eANNOT_FibClusters
                Annot.Prop("Shape") = 0
                Annot.PreIndicator = 1
                Annot.Prop("FillPattern") = 1
                Annot.Color = Annot.ClusterPropDefault("ZoneColor")
                Annot.Prop("FillColor") = Annot.ClusterPropDefault("ZoneFillColor")
                strKey = kClusterZoneRect
            ElseIf InStr(g.strActiveDraw, "_PFP") <> 0 Then
                Annot.eUsage = eANNOT_PatternProfit
                Annot.Prop("Shape") = 0
                Annot.Prop("FillPattern") = 0
                Annot.PreIndicator = 1
                i = Annot.PatternLength(Nothing)        'this is just to get a flag set
                If Not m.oPatternProfit Is Nothing Then Annot.Color = m.oPatternProfit.ForecastColor
            End If
    
        Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4, eANNOT_RiskReward
            If IsForex(m.Chart.Symbol) And (Val(Annot.Prop("NumForexContracts")) <= 0) Then
                Annot.Prop("NumForexContracts") = g.Broker.DefaultOrderQuantity(m.Chart.TradeAccountID, m.Chart.SymbolID)
            ElseIf (Val(Annot.Prop("NumContracts")) <= 0) Then
                Annot.Prop("NumContracts") = g.Broker.DefaultOrderQuantity(m.Chart.TradeAccountID, m.Chart.SymbolID)
            End If

        Case eANNOT_BalloonStrangle
            If Not m.Chart Is Nothing Then
                If Not m.Chart.Bars Is Nothing Then
                    'get to the "next 3 Friday"     '7036
                    i = 0
                    dPoints = Annot.dDate(1)
                    Do While i < 3
                        dPoints = dPoints + 1
                        If Weekday(dPoints) = vbFriday Then i = i + 1
                    Loop
                    Annot.dDate(2) = dPoints
                End If
            End If
        
        Case eANNOT_HorzLine
            If Len(strSave) > 0 Then
                Annot.Prop("FakePriceAlert") = 1
                Set Alert = Annot.AlertObject(True)
            End If
    
    End Select
        
    ' add to annots and save key
    If Len(strKey) > 0 Then
        i = m.Chart.Annots.Add(Annot, strKey)
    Else
        i = m.Chart.Annots.Add(Annot)
        strKey = m.Chart.Annots.Key(i)
    End If
    Annot.Prop("AnnotKey") = strKey
    
    ' get index of annot (safer to find by key)
    m.nActiveAnnotIdx = m.Chart.Annots.Index(strKey)
    Annot.geAnnId = m.nActiveAnnotIdx
    
    'set default name for pattern annotations
    If Annot.eType = eANNOT_Pattern Then
        Annot.Text = "Pattern #" & Str(Annot.geAnnId)
    End If
    
    m.bAnnotCreated = True
    g.bDirtyChartPage = True
    
    If Len(strSave) > 0 Then
        ClearAnnotFlags False
        g.strActiveDraw = strSave
        m.Chart.GenerateChart eRedo1_Scrolled
    End If
        
ErrExit:
    Set Annot = Nothing
    Set Alert = Nothing
    Set Bars = Nothing
    Set m.AnnotOptions = Nothing
    
    Exit Sub
    
ErrSection:
    Set Annot = Nothing
    Set Alert = Nothing
    Set Bars = Nothing
    Set m.AnnotOptions = Nothing
    
    RaiseError Me.Name & ".AddNewAnnot", eGDRaiseError_Raise
    
End Sub

Private Sub pbChart_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    Else
        KeyPress KeyCode, Shift
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".pbChart_KeyDown"
    Resume ErrExit
    
End Sub

Private Sub pbChart_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".pbChart_KeyPress"
    Resume ErrExit
    
End Sub

Private Sub pbChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:

    'If m.bEditing Then Exit Sub    - Q & A: need this?
    
    Dim Annot As cAnnotation
    Dim NewAnnot As cAnnotation
    Dim Pane As cPane
    Dim Ind As cIndicator
    
    Dim eType As eAnnotType
    Dim eNewScaleFlag As enumScaleFlag
    
    Dim i&, rc&, dY#, strText$, strKey$, dDate#
    Dim nPixDiffX&, nPixDiffY&
    Dim bMoved As Boolean
    Dim bIconPalVisible As Boolean
    Dim hArrayDate As Long
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    
    Static dPrevX#, dPrevY#
    
    If FormIsLoaded("frmTemplatePage") Then
        Unload frmTemplatePage
    End If
    
    If FormIsLoaded("frmTbMoreButtons") Then
        Unload frmTbMoreButtons
    End If
    
    ' for some reason, the MouseMove is getting called continuously even
    ' when the mouse is not moving -- this will help eliminate doing this
    ' routine when we don't need to
    If X = dPrevX And Y = dPrevY Then Exit Sub
    dPrevX = X
    dPrevY = Y
    
    ' save which chart the mouse last moved over (for cross-hair synching while real-time)
    Set g.ChartGlobals.frmLastChartMouseMove = Me
        
    ' increment object move count if user is dragging on an indicator, annotation or y-scale
    If m.nObjectMoving > 0 Then
        If InStr(tmr.Tag, "EditAnnot") <> 0 Then tmr.Tag = ""
        m.nObjectMoving = m.nObjectMoving + 1
    End If
        
    ' get chart coordinates of mouse move
    m.MouseLast = GetChartCoordinates(X, Y, Shift)
    m.MouseLast.nButton = Button

    If FormIsLoaded("frmIconAnnot") Then
        If frmIconAnnot.Visible Then
            bIconPalVisible = True
            frmIconAnnot.chartObj = m.Chart
            g.strActiveDraw = "ID_Icon"
            
            If m.MouseLast.bOffChart Then
                m.epbCursor = eCursor_Default
            ElseIf Not m.Chart.Tree(m.MouseLast.nPaneID) Is Nothing Then
                '5514 - show on all chart is only for annotations in Price Pane
                frmIconAnnot.ToggleMultichart m.Chart.Tree(m.MouseLast.nPaneID).PricePaneFlag
            End If
        End If
    ElseIf FormIsLoaded("frmElliot") Then
        If frmElliot.Visible Then
            If frmMain.tbToolbar.Tools("ID_ChartMove").State = ssChecked Then
                g.strActiveDraw = ""
            End If
        End If
    End If
    
    ShowCursor
    
    If Not bIconPalVisible Then
        If m.epbCursor <> eCursor_ChartMove Then
            GlobalCursorSync
        End If
        
        If g.ChartGlobals.eChartMode = eMode_ChartOrder Or fraWizardPrompt.Visible Then
            HandleWizardPrompt X, Y
        ElseIf m.epbCursor = eCursor_NoDrop Then
            HandleZoom X, Y, False
            ShowCursor
        ElseIf m.epbCursor = eCursor_Magnify Then
            Me.pbChart.AutoRedraw = False
            HandleZoom X, Y, True
            Me.pbChart.AutoRedraw = True
        ElseIf m.epbCursor = eCursor_ChartMove Then
            HandleChartMove X, Y, False
        ElseIf m.bChartMoveInProg = True Then
            HandleChartMove X, Y, True  'user moved out of chart area before releasing mouse
        ElseIf m.eScaleFlag = eScaleY_Max Or m.eScaleFlag = eScaleY_Min Then
            'moving y-scale
            If 0 = m.nObjectMoving Mod 5 Then HandleScaleMoveY X, Y, Shift
        ElseIf m.eScaleFlag = eScaleX_MoreLessBars Then
            'moving x-scale
            If 0 = m.nObjectMoving Mod 5 Then HandleScaleMoveX X, Y, Shift
        End If
        
        If m.epbCursor >= eCursor_NoDrop And m.epbCursor < eCursor_HandDown _
            Or m.epbCursor = eCursor_ChartMove Then
            Exit Sub   'user is zooming, moving chart or manipulating scales
        End If

        If m.MouseLast.bOffChart Then
            ' mouse is off chart ...
            If Not m.bPrevOffChart Then
                ' turn tips off
                RefreshTips -1000, -1000
                m.bPrevOffChart = True
            End If
            'check if in scale areas
            If m.MouseLast.nScalePaneId = -6 Then
                Set Pane = m.Chart.Tree("PRICE PANE")
                If Pane.SplitPaneType = ePANE_SplitPaneOptGraph Then
                    m.Chart.SetCursor
                ElseIf m.nSplitPaneHittestID > 0 Then
                    m.epbCursor = eCursor_Hand
                Else
                    m.Chart.SetCursor
                End If
            ElseIf m.MouseLast.nScalePaneId <> 0 Then
                SetScaleFlag eNewScaleFlag, False
                If eNewScaleFlag <> eScale_Unhandled Then SetScaleCursor eNewScaleFlag, False
            End If
            If m.nActiveAnnotIdx > 0 And m.nObjectMoving > 0 Or m.bAnnotCreated = True Then
                Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
                If Not Annot Is Nothing Then
                    PbChart_MouseUp Button, Shift, pbChart.ScaleWidth, Y
                End If
            End If
            Exit Sub
        ElseIf m.bPrevOffChart And m.nActiveAnnotIdx = 0 Then
            m.bPrevOffChart = False
        End If
        
        'mouse is not off chart - reset cursor if necessary
        If m.epbCursor >= eCursor_HandDown Then
            m.Chart.SetCursor
        End If
    End If      'end if icon palette is not visible
    
    With m.MouseLast
        ' refresh date/price tips
        hArrayDate = m.Chart.geDateArray
        If .dDate > gdGetNum(hArrayDate, gdGetSize(hArrayDate) - 1) Then
            vseTipX.Tag = ""
        ElseIf .dDate < gdGetNum(hArrayDate, 0) Then
            vseTipX.Tag = ""
        ElseIf .dDate > 0 Then
            dDate = DateTimeConvert(m.Chart.Bars, .dDate)
            If IsIntraday(m.Chart.Periodicity) Then
                vseTipX.Tag = DateFormat(dDate, MM_DD_YY) & Format(dDate, " Hh:Nn")
            Else
                vseTipX.Tag = DateFormat(dDate, MM_DD_YYYY)
            End If
        End If
        
        ' price tips
        strText = .strRoundedY
                
        'Need this to prevent "Mismatch Type" error
        Set Pane = Nothing
        If m.Chart.Tree.NodeLevel(.nPaneID) = 0 Then
            Set Pane = m.Chart.Tree(.nPaneID)
        End If
         
        If Not Pane Is Nothing Then
            If Pane.DisplayFormat = ePANE_PriceFormat And Pane.PaneLogFlag <> ePANE_LogFlagPercent And m.Chart.TypeOfChart <> eTypeChart_Seasonal Then
                'strText = m.Chart.PriceDisplay(Pane.gePaneId, .dY) ', 6)
                strText = m.Chart.PriceDisplay(.nPaneID, .dY) ', 6)
            End If
            Set Pane = Nothing
        End If
        
        vseTipY.Tag = strText & vbTab & CStr(.nPaneID)
        RefreshTips X, Y
        ' refresh mouse label
        If m.bLockValuesDisplay Then
            i = -1 ' m.Chart.LastGoodDataBar(True)
        Else
            i = .nX
        End If
        vseMouse.Tag = Str(.nPaneID) & Chr(9) & Str(i)
        DoMouseLabel
    End With
    
    If m.nActiveAnnotIdx < 0 Then
        If g.strActiveDraw = "ID_DNRetracement" Or g.strActiveDraw = "ID_DNExpansion" _
            Or g.strActiveDraw = "ID_AndrewFork" Or g.strActiveDraw = "ID_FibABCD" Or _
            g.strActiveDraw = "ID_Gartley" Then
            Exit Sub
        ElseIf g.strActiveDraw = "ID_Triangle" Or g.strActiveDraw = "ID_ChannelHighlight" Then
            HandleTriChannel False, False
        End If
        AddNewAnnot False       'this handles a click-n-drag draw
    End If
       
    If m.bAnnotCreated = True Then
    
        Select Case g.strActiveDraw
            Case "ID_Icon", "ID_ElliotLabels", "ID_ElliotEndUser"
                Exit Sub
            Case "ID_Fibonacci", "ID_FibExpansion", "ID_DanCodeFib"
                Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
                If Not Annot Is Nothing Then
                    If Annot.Prop("KeepAlive") = 1 Then Exit Sub
                End If
        End Select
    
    End If
    
    If m.nActiveBtmPane > 0 Then
        Me.pbChart.AutoRedraw = False
        rc = geDragSeparator(m.Chart.geChartObj, 0, Me.pbChart.hDC, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
        Me.pbChart.AutoRedraw = True
        Exit Sub
    End If
        
    ' see if moving a horz line indicator
    If m.nActiveIndIdx > 0 And m.nObjectMoving > 0 Then
        If m.Chart.Tree.NodeLevel(m.nActiveIndIdx) > 0 Then
            Set Ind = m.Chart.Tree(m.nActiveIndIdx)
            If Ind Is Nothing Then
                m.nObjectMoving = 0
                m.nActiveIndIdx = 0
            ElseIf Ind.DataType = eINDIC_Constant And m.nObjectMoving > kMinMoveCount Then
                Ind.Parm(0) = m.MouseLast.strRoundedY
                pbChart.AutoRedraw = False
                m.Chart.geDrawChart
                pbChart.AutoRedraw = True
            End If
        End If
        Set Ind = Nothing
        Exit Sub
    End If
    
    'see if moving an order
    'using min move count > 3 is too high for detecting order movement -aardvark 4114
    If m.nActiveOrderID > 0 And m.nObjectMoving > kMinOrderMove Then
        Dim Order As cPtOrder
        
        Set Order = New cPtOrder
        If Order.Load(m.nActiveOrderID) Then
            If Not OrderIsPending(Order) Then
                If m.nActiveOrderLoc = 3 Or m.nActiveOrderLoc = 1 Then
                    Order.OrderPrice(True) = m.MouseLast.dY
                    m.Chart.UpdateOnlineOrder Order, False
                End If
            End If
        End If
        
        Exit Sub
    End If
    
    ' see if moving an annotation
    If m.nActiveAnnotIdx > 0 Then
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Annot.eType = eANNOT_DNExpansion Or Annot.eType = eANNOT_FibABCD Or _
            Annot.eType = eANNOT_DNExpansion2 Or Annot.eType = eANNOT_DNExpansion3 Or _
            Annot.eType = eANNOT_DNExpansion4 Or Annot.eType = eANNOT_GannacciSwing1 Or _
            Annot.eType = eANNOT_GannacciSwing2 Then
            HandleDnExpansion False, dY
            Exit Sub
        ElseIf Annot.eType = eANNOT_Gartley Then
            HandleGartley False, dY
            Exit Sub
        ElseIf Annot.eType = eANNOT_DNRetracement Then
            HandleDnRetrace False, False, dY
            Exit Sub
        ElseIf Annot.eType = eANNOT_AndrewFork Or Annot.eType = eANNOT_ElliotTimeRatio Then
            HandleAndrewFork False
            Exit Sub
        ElseIf Annot.eType = eANNOT_ChannelHighlight Or Annot.eType = eANNOT_TriangleWedge Then
            HandleTriChannel False, False
            Exit Sub
        ElseIf Annot.eType = eANNOT_WaveLabels Then
            HandleWaveLabels False
        ElseIf Annot.eType = eANNOT_SimpleLine Then
            HandleWhatIf Annot
        ElseIf Annot.eType = eANNOT_TrendChannel And m.nPointCount = 3 Then
            dY = SnapToPrice(m.MouseLast, Annot.eType)
            Annot.SetChannelLocationOnAdd dY
            bMoved = Annot.MovePoint(m.Chart, 0, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
            If bMoved = True Then Annot.geDrawAnn m.Chart
            Exit Sub
        End If
        
        ' Note: on a newly created annotation object move count will be zero
        ' as the hittest would not have identified it as a hit.
        ' move existing or newly created annotation
        If m.nObjectMoving > kMinMoveCount Or m.bAnnotCreated Then
            Annot.ShiftKey = Shift
            'aardvark 1824: let grapheng.dll knows not to auto position text
            geAnnotMove m.Chart.geChartObj, 1
            If Annot.eUsage = eANNOT_IndicatorLabel Then
                If m.Chart.Zoomed = True Then
                    m.nActiveAnnotIdx = 0
                Else
                    Annot.geMoveIndLabel m.Chart, X, Y
                End If
            ElseIf Annot.eUsage = eANNOT_Trades Then
                If m.bGameMode And Not tmrGameMode.Enabled Then
                    HandleGameMoveOrder False, Annot
                End If
            Else
                dY = SnapToPrice(m.MouseLast, Annot.eType)
                With Annot
                    Select Case .eType
                        Case eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4
                            'check for minimum annotation size
                            nPixDiffX = Abs(m.MouseLast.MouseX / Screen.TwipsPerPixelX - m.MouseDown.MouseX / Screen.TwipsPerPixelX)
                            nPixDiffY = Abs(m.MouseLast.MouseY / Screen.TwipsPerPixelY - m.MouseDown.MouseY / Screen.TwipsPerPixelY)
                            If nPixDiffX > kMinAnnotSize And nPixDiffY > kMinAnnotSize And Annot.geMoveFlag = 0 Then
                                'auto add arrow to text
                                .dDate(1) = m.MouseDown.dDate
                                .dDate(2) = m.MouseLast.dDate
                                .Y(1) = m.MouseDown.dY
                                .Y(2) = m.MouseLast.dY
                                .Prop("ArrowStyle") = 2    'triangle arrowhead
                            End If
                        Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4, _
                             eANNOT_RiskReward, eANNOT_GannacciSwingSquare
                            .geMoveFlag = 0     'set so grapheng will show text without auto-positioning
                        Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4
                            If Shift = 1 And Val(Annot.Prop("ChannelCount")) = 0 Then
                                Annot.AddChannelOnMove m.MouseDown.dY, dY
                            End If
                        Case eANNOT_Bracket
                            If m.bAnnotCreated Then
                                If m.MouseLast.MouseX > m.MouseDown.MouseX Then
                                    Annot.Prop("BracketDirection") = 0
                                Else
                                    Annot.Prop("BracketDirection") = 1
                                End If
                            End If
                    End Select
                End With
                
                If Annot.eType = eANNOT_Ellipse And m.nActiveAnnotPt = 4 Then
                    'user is resizing minor axis
                    bMoved = Annot.MovePoint(m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, X, Y)
                ElseIf Annot.eType = eANNOT_TimeCycle And m.nActiveAnnotPt = 4 Then
                    bMoved = Annot.MovePoint(m.Chart, m.nActiveAnnotPt, m.MouseDown.nPaneID, m.MouseDown.nX, Y)
                    If bMoved Then m.nActiveAnnotPt = 2
                ElseIf Shift = 2 And Annot.HitItemIndex = 0 And _
                        (Annot.eType = eANNOT_TrendChannel Or _
                         Annot.eType = eANNOT_Trendline Or Annot.eType = eANNOT_Trendline2 Or _
                         Annot.eType = eANNOT_Trendline3 Or Annot.eType = eANNOT_Trendline4) Then
                         
                    If Annot.X(1) = Annot.X(2) Then
                        bMoved = Annot.MovePoint(m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, Annot.X(1), dY)
                    Else
                        dY = (Annot.Y(2) - Annot.Y(1)) / (Annot.X(2) - Annot.X(1))
                        If m.nActiveAnnotPt = 1 Then
                            dY = Annot.Y(2) - dY * (Annot.X(2) - m.MouseLast.nX)
                        Else
                            dY = dY * (m.MouseLast.nX - Annot.X(1)) + Annot.Y(1)
                        End If
                        bMoved = Annot.MovePoint(m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
                    End If
                Else            'If Shift = 0 Then
                    bMoved = Annot.MovePoint(m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
                
                    If Val(Annot.Prop("HideVerticalLine")) = 1 Then
                        Set m.AnnotOptions = Annot
                        Annot.Prop("HideVerticalLine") = -1         '6728 - show diag only when moving
                    End If
                
                End If
                
                If bMoved = True Then
                    'If m.epbCursor = eCursor_Arrow4Way Then m.epbCursor = eCursor_Hand
                    If Len(g.strActiveDraw) = 0 Then m.epbCursor = eCursor_Blank         'aardvark 3642
                    If Annot.eType = eANNOT_ChannelHighlight Or Annot.eType = eANNOT_TriangleWedge Then
                        'want user to see which "point number" they are working on if a draw is in progress
                        If Len(g.strActiveDraw) = 0 Then
                            StatusMsg Annot.MoveCaption(m.Chart)
                        End If
                    Else
                        StatusMsg Annot.MoveCaption(m.Chart)
                    End If
                    If g.strActiveDraw = "ID_TrendChannel" And m.nPointCount = 0 Then m.nPointCount = 2
                    Annot.geDrawAnn m.Chart
                Else
                    m.Chart.SetFormCaption
                End If
            End If
        End If
        Set Annot = Nothing
        Exit Sub
    End If
                   
    DoHitTest X, Y, False, vbLeftButton
    
    ' TLB 12/9/2011: just so a toolbar button won't get stuck with the "hover" color
    ' (e.g. when something long is happening in the frmMain timer which disables it's checking)
    frmMain.CheckHighlightedToolButton
    
End Sub

Private Sub PbChart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim rc&, i&, idx&, nPaneID&, idxPane&, d#
    Dim strText$, strKey$
    
    Dim eType As eAnnotType
    Dim bSkipSplitCfg As Boolean
    Dim bDone As Boolean
    
    Dim Annot As cAnnotation
    Dim AnnotVertLine As cAnnotation
    Dim Ind As cIndicator
    Dim Pane As cPane
    
    Dim aResults As cGdArray
    Dim aStrings As New cGdArray
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "UNLOADING" Then Exit Sub
    If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then Exit Sub
    If m.bDrawToolSelected Then Exit Sub
    
    geAnnotMove m.Chart.geChartObj, 0           'clear flag (see note for aardvark 1824)
    
    'Check for zoom
    If m.epbCursor = eCursor_Magnify Then
        HandleZoom X, Y, False
        m.MouseLast = GetChartCoordinates(X, Y, Shift)      'aardvark 4120
        m.MouseLast.nButton = Button
        ShowCursor
        GlobalCursorSync
        GoTo ErrExit
    'check for chart move
    ElseIf m.epbCursor = eCursor_ChartMove Then
        HandleChartMove X, Y, True
        If vseOrderBar.Visible And m.eOrdBarMode = eOrdBarMode_Wizard Then cboRiskGraphType_Click
        bSkipSplitCfg = True
    'check for scale drag/move
    ElseIf m.eScaleFlag = eScale_Arrow Then
        If Button = vbRightButton Then
            frmChartPixPerBar.ShowMe Me
        Else
            If Not m.Chart Is Nothing Then m.Chart.RestoreChartNormal vbKeyClear
            If InStr(g.strActiveDraw, "PFP") <> 0 Then
                ClearAnnotFlags True, True
                g.strActiveDraw = ""
                SyncDrawTools
            End If
        End If
        m.eScaleFlag = eScale_Unhandled
    ElseIf m.eScaleFlag = eScale_Arrow_Area Then
        frmChartPixPerBar.ShowMe Me
        m.eScaleFlag = eScale_Unhandled
    ElseIf m.eScaleFlag <> eScale_Unhandled Then
        m.nObjectMoving = 0
        If m.eScaleFlag = eScaleY_Max Or m.eScaleFlag = eScaleY_Min Then
            If Not GetIniFileProperty("DontShowEscMsg", False, "Charting", g.strIniFile) Then
                If Right(InfBox("Hitting the 'Esc' key will clear the extra| space above/below all chart panes.", "i", , "Note ...", , , , , , , , , True), 1) = "-" Then
                    SetIniFileProperty "DontShowEscMsg", True, "Charting", g.strIniFile
                End If
            End If
            HandleScaleMoveY 0, 0, 0, True
            g.strActiveDraw = ""
            
            If vseOrderBar.Visible And m.eOrdBarMode = eOrdBarMode_Wizard Then cboRiskGraphType_Click
            
        End If
        m.eScaleFlag = eScale_Unhandled
    End If
            
    ' Get chart coordinates of mouse move
    m.MouseLast = GetChartCoordinates(X, Y, Shift)
    m.MouseLast.nButton = Button
    If m.MouseLast.nScalePaneId = -6 And m.MouseLast.bOffChart = False Then
        'JM 09-01-2015: fix for bug reported by Richard (orders in split pane not selectable)
        DoHitTest X, Y, False, Button
        If m.nActiveOrderID > 0 Then
            m.MouseLast.nPaneID = m.nSplitPaneHittestID
        End If
    End If
    
                
    'see if adding annotation with a single click
    'a)annotations that can be added with only a single click are:
    '   icon, fib time zones
    'b)annotations that can be added with a single click AND drag-n-click are:
    '   $-Line, SR-Line, Text, Gann Fan
    If m.nActiveAnnotIdx < 0 Then
        If g.strActiveDraw = "ID_Hawkeye" Then
            If Not m.Chart Is Nothing Then m.Chart.HandleHawkeyeButton m.MouseDown.dDate
            ClearAnnotFlags False
            SyncDrawTools
            m.Chart.GenerateChart eRedo5_RecalcInd
        ElseIf g.strActiveDraw = "ID_FibClusters" Then
            Set Annot = m.Chart.Annots(kClusterZoneRect)
            If Annot Is Nothing Then
                AddNewAnnot False
                'set first annot date to 30 bars back from where user clicked
                If m.nActiveAnnotIdx > 0 Then
                    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
                    If Not Annot Is Nothing Then
                        d = Chart.Bars(eBARS_DateTime, Chart.LastGoodDataBar(False))
                        If Annot.dDate(2) > d Then Annot.dDate(2) = d
                        i = Chart.Bars.FindDateTime(Annot.dDate(2))
                        If i - 30 >= 0 Then i = i - 30 Else i = 0
                        Annot.dDate(1) = Chart.Bars(eBARS_DateTime, i)
                        rc = Annot.PatternLength(Chart)
                    End If
                End If
            Else
                ClearAnnotFlags False
                SyncDrawTools
            End If
        Else
            AddNewAnnot True, g.strActiveDraw
            If g.strActiveDraw = "ID_PriceAlert" Then
                SyncDrawTools
            ElseIf g.strActiveDraw = "ID_DNRetracement" Or g.strActiveDraw = "ID_DNExpansion" _
                Or g.strActiveDraw = "ID_AndrewFork" Or g.strActiveDraw = "ID_Triangle" _
                Or g.strActiveDraw = "ID_ChannelHighlight" Or g.strActiveDraw = "ID_WaveLabels" _
                Or g.strActiveDraw = "ID_ElliotTimeRatio" Or g.strActiveDraw = "ID_FibABCD" Then
                GoTo ErrExit
            ElseIf m.nPointCount < 2 Then
                GoTo ErrExit
            End If
        End If
        'm.nActiveAnnotIdx = 0  - part of fix for Aardvark 1391, leave awhile then remove 05-19-2004
    End If
        
    'check for adding vertical or horizontal annotations
    ' (mouse doesn't need to have been moved for these types)
    If m.nActiveBtmPane <= 0 Then
       If Button = vbLeftButton Then
            If Shift = 1 And m.nActiveAnnotPt = -1 Then     'shift key
                AddNewAnnot True, "ID_VertLine"
            ElseIf Shift = 2 And m.nActiveAnnotPt = -1 Then    'control key
                'If g.ChartGlobals.eChartMode <> eMode_Erase Then
                    AddNewIndicator 2, True         '6183
                'End If
            End If
        End If
    End If
    
    'Currently the hand cursor is shown only for horz-line indicators
    'so can use as quick check to determine whether to do editing.
    If m.nActiveIndIdx > 0 And m.epbCursor = eCursor_Hand Then
        If m.nObjectMoving > 0 And m.nObjectMoving < kMinMoveCount Then
            tmr.Tag = "EditSettings " & CStr(m.nActiveIndIdx)
        ElseIf m.Chart.Tree.NodeLevel(m.nActiveIndIdx) > 0 Then
            Set Ind = m.Chart.Tree(m.nActiveIndIdx)
            If Not Ind Is Nothing Then
                If Ind.DisplayType = eINDIC_ClusterTime Then
                    frmClusterCfg.ShowMe m.Chart, Ind        'price cluster is in splitpane code below
                End If
            End If
            Set Ind = Nothing
        End If
        m.nObjectMoving = 0
    End If
    m.nActiveIndIdx = 0
    
    'see if editing an order
    If m.nActiveOrderID > 0 Then
        HandleChartOrder False
        GoTo ErrExit
    End If
       
    ' See if editing or moving an annotation
    If m.nActiveAnnotIdx > 0 Then
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Not Annot Is Nothing Then
            If Annot.eUsage = eANNOT_UserAdded Then
                g.bDirtyChartPage = True
            End If
            
            strKey = m.Chart.Annots.Key(m.nActiveAnnotIdx)
            idx = Val(strKey)
            If idx > 0 Then
                Set Ind = m.Chart.Tree(idx)
            Else
                Set Ind = Nothing
            End If
            
            If g.ChartGlobals.eChartMode = eMode_Erase And m.bAnnotCreated = False And m.bNewPatternMoving = False Then
                EraseAnnot Ind, Annot, Button, Shift
                m.Chart.GenerateChart eRedo1_Scrolled
                m.Chart.SetCursor
                ShowCursor
                GlobalCursorSync
            Else
                If EditMoveAnnot(Ind, Annot, X, Y) = True Then
                    If Len(g.strActiveDraw) = 0 Then        '6334
                        If Annot.MultiChartFlag Then m.Chart.SyncGlobalAnnots Annot     '6314,6315
                    End If
                    GoTo ErrExit
                End If
                If Not m.bNewPattern Then
                    m.nActiveAnnotIdx = 0              'aardvark 3671 - cursor not showing soon enough
                    m.Chart.SetCursor
                    ShowCursor
                End If
            End If

            If Left(UCase(tmr.Tag), 4) <> "EDIT" Then
                If Not Annot Is Nothing Then        '4244
                    Dim bSaveFlag As Boolean
                    bSaveFlag = m.bAnnotCreated
                    
                    m.Chart.SyncGlobalAnnots Annot
                    SyncDrawTools True  'this will clear out the annot created flag
                    ShowCursor
                    
                    If bSaveFlag Then
                        If Annot.eType = eANNOT_BalloonStrangle Then
                            If Year(Annot.dDate(2)) = Year(Date) Then
                                i = Month(Annot.dDate(2)) - Month(Date)
                                If i >= 0 And i < 2 Then
                                    i = 1
                                Else
                                    i = 0
                                End If
                            ElseIf Year(Annot.dDate(2)) > Year(Date) Then
                                If Month(Date) = 12 And Month(Annot.dDate(2)) = 1 Then
                                    i = 1
                                Else
                                    i = 0
                                End If
                            Else
                                i = 0
                            End If
                            If Annot.dDate(2) >= Date And i = 1 Then
                                If InfBox("Auto-fill using the options data for the next expiration?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                                    Set aResults = GetOptionChainBidAskData(m.Chart.Symbol, 0)
                                    aResults.Sort
                                    If aResults.Size > 0 Then
                                        i = 1
                                        strText = Parse(aResults(0), " ", 2)
                                        For idx = 0 To aResults.Size - 1
                                            'Aardvark: 7036
                                            'dDate2 contains furthest out option date requested
                                            'dDate2 - 14 allows 2 weeks back
                                            If DateOf(strText) < Annot.dDate(2) - 14 Then
                                                strText = Parse(aResults(idx), " ", 2)
                                            ElseIf Parse(aResults(idx), " ", 2) = strText Then
                                                aStrings.Add aResults(idx)
                                            Else
                                                Set AnnotVertLine = New cAnnotation
                                                AnnotVertLine.CreateNew m.Chart, eANNOT_VertLine, 1, DateOf(strText), 0, Month(DateOf(strText)), 0, vbBlue, , , eANNOT_OptionInfo
                                                AnnotVertLine.BalloonOptionsInfo = aStrings
                                                AnnotVertLine.Y(2) = Annot.BalloonStockPrice
                                                AnnotVertLine.Prop("FontSize") = 14
                                                
                                                bDone = IsWeekday(AnnotVertLine.dDate(1)) And Not m.Chart.Bars.IsHoliday(AnnotVertLine.dDate(1))
                                                While Not bDone
                                                    AnnotVertLine.dDate(1) = AnnotVertLine.dDate(1) - 1
                                                    bDone = IsWeekday(AnnotVertLine.dDate(1)) And Not m.Chart.Bars.IsHoliday(AnnotVertLine.dDate(1))
                                                Wend
                                                AnnotVertLine.dDate(2) = AnnotVertLine.dDate(1)
                                                
                                                m.Chart.Annots.Add AnnotVertLine
                                                Annot.BalloonExpiration = AnnotVertLine.dDate(1)
                                                
                                                i = i + 1
                                                If i > 3 Then
                                                    Exit For
                                                Else
                                                    aStrings.Clear
                                                    aStrings.Add aResults(idx)
                                                    strText = Parse(aResults(idx), " ", 2)
                                                End If
                                            End If
                                        Next
                                        
                                        i = AnnotVertLine.dDate(1) - gdGetNum(m.Chart.geDateArray, gdGetSize(m.Chart.geDateArray) - 1)
                                        If i > 0 Then
                                            m.Chart.ForecastBars(Me) = m.Chart.ForecastBars + i + 3
                                            m.Chart.GenerateChart eRedo7_ReloadRT
                                            hsb.Value = hsb.Max
                                        Else
                                            m.Chart.GenerateChart eRedo1_Scrolled
                                        End If
                                    End If
                                End If  'prompt for autofill
                            End If
                            
                            If AnnotVertLine Is Nothing Then
                                Annot.BalloonExpiration = Annot.dDate(2) - 7
                                tmr.Tag = "EditAnnot " & Annot.geAnnId
                            End If
                            
                        ElseIf Annot.eType = eANNOT_GannacciCycle Then
                            tmr.Tag = "EditAnnot " & Annot.geAnnId
                        End If
                    End If
                End If
            End If  'tag <> EDIG
        End If
    End If  'end checking for edit, move or erase annotation
                
    'see if sizing or erasing panes
    If m.nActiveBtmPane > 0 Then
        If m.Chart.Zoomed = False Then
            If g.ChartGlobals.eChartMode = eMode_Erase Then
                Set Pane = m.Chart.Tree(m.nActiveBtmPane)
                If Not Pane Is Nothing Then
                    Pane.Display = False
                    m.Chart.GenerateChart eRedo1_Scrolled
                End If
            Else
                ResizePanes X, Y, Shift
                m.Chart.GenerateChart eRedo1_Scrolled
            End If
        End If
        m.nActiveTopPane = 0
        m.nActiveBtmPane = 0
    End If
    
    'see if in split-pane area
    If Not bSkipSplitCfg Then
        If m.MouseLast.nScalePaneId = -6 And m.epbCursor = eCursor_Hand Then
            If m.nSplitPaneHittestID > 0 Then
                Set Pane = m.Chart.Tree("PRICE PANE")
                If Not Pane Is Nothing Then
                    If Pane.SplitPaneType = ePANE_SplitPaneCluster Then
                        frmClusterCfg.ShowMe m.Chart, m.Chart.Tree(kClusterPriceKey)
                    Else
                        frmSplitPaneCfg.ShowMe m.Chart, m.nSplitPaneHittestID, m.nSplitPaneLabelID
                    End If
                End If
            End If
        End If
    End If
    
    ' clear all flags
    ClearAnnotFlags False
    m.nActiveIndIdx = 0
    m.nObjectMoving = 0
    
    If Len(g.strActiveDraw) > 0 And g.strActiveDraw <> "ID_Icon" And _
        g.strActiveDraw <> "ID_ElliotLabels" And g.strActiveDraw <> "ID_ElliotEndUser" And _
        frmMain.tbToolbar.Tools("ID_RepeatDraw").State = ssUnchecked Then
        ToolbarSetCursorGroup frmMain.tbToolbar, False
    End If
    
    Dim sX&, sY&
    
    If m.bNewPattern Then
        If Annot Is Nothing Then
            m.Chart.SetCursor       'theoretically should never get here
            ShowCursor
            GlobalCursorSync
        ElseIf Annot.eType = eANNOT_Pattern Then
            geAnnotCursorPos Annot.geAnnotObject, pbChart.hWnd, sX, sY
            
            sX = sX * Screen.TwipsPerPixelX
            sY = sY * Screen.TwipsPerPixelY
            
            m.epbCursor = eCursor_Hand
            m.nObjectMoving = kMinMoveCount
            m.nActiveAnnotIdx = Annot.geAnnId
            m.nActiveAnnotPt = 0
            Annot.HitItemIndex = 0
            Annot.MovePoint m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, sX, sY
            Annot.geDrawAnn m.Chart
            m.bNewPatternMoving = True
        Else
            m.Chart.SetCursor       'theoretically should never get here
            ShowCursor
            GlobalCursorSync
        End If
        m.bNewPattern = False
    Else
        m.Chart.SetCursor
        ShowCursor
        GlobalCursorSync
    End If
    
    If Not AnnotVertLine Is Nothing Then
        Set m.AnnotOptions = Annot  'so don't have to keep looking for the balloon strangle annot
        m.epbCursor = eCursor_OrderBuy
        m.eOrdBarMode = eOrdBarMode_Wizard
        vseBuyWizard_Click
        
        X = m.MouseLast.nX
        Y = m.MouseLast.dY
        pbChart_MouseMove Button, Shift, X, Y
    
    ElseIf Not Ind Is Nothing Then
        m.Chart.GenerateChart eRedo1_Scrolled
    End If

ErrExit:
    Set Annot = Nothing
    Set AnnotVertLine = Nothing
    Set Ind = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".pbChart_MouseUp"
    Resume ErrExit
    
End Sub

Private Sub pbChart_DblClick()
On Error GoTo ErrSection:
    
    Dim dY#
         
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "UNLOADING" Then Exit Sub
    If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then Exit Sub
    
    tmr.Tag = ""
    DoHitTest m.MouseLast.MouseX, m.MouseLast.MouseY, True, vbLeftButton
    
    If Len(g.strActiveDraw) > 0 Then
        If g.strActiveDraw = "ID_DNRetracement" Then
            m.nActiveIndIdx = -1
            HandleDnRetrace False, True, dY
        ElseIf g.strActiveDraw = "ID_WaveLabels" Then
            HandleWaveLabels False, True
        End If
        Exit Sub        'a draw is in progress, don't duplicate
    End If
        
    If m.nActiveAnnotIdx < 1 And m.nActiveIndIdx < 1 Then Exit Sub
    
    If m.nActiveAnnotIdx > 0 Then
        HandleAnnotDblclk
    End If
    
    If m.nActiveIndIdx > 0 Then
        With m.MouseLast
            If m.Chart.Tree.Key(m.nActiveIndIdx) = "PRICE" Then
                m.bChartMoveInProg = False
                If .nBar >= 0 Then
                    'dFudge = (.dMaxY - .dMinY) / 100#
                    'If .dy >= m.Chart.Bars(eBARS_Low, .nBar) - dFudge And _
                    '    .dy <= m.Chart.Bars(eBARS_High, .nBar) + dFudge Then
                    '        tmr.Tag = "EditData " & CStr(.nX)
                    'End If
                    EditData .nX
                End If
            End If
        End With
    End If
    
    m.epbCursor = eCursor_Hand
    ShowCursor

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".pbChart_DblClick"
    Resume ErrExit
    
End Sub

Private Sub pbChart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim i&
    Dim eNewScaleFlag As enumScaleFlag
    Dim Pane As cPane
    Dim Ind As cIndicator
    Dim Annot As cAnnotation
    Dim bSkipHitTest As Boolean
        
    StatusMsg ""
        
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "UNLOADING" Then Exit Sub
    If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then Exit Sub
    
    TextIncDecRegisterForm Me, True
    
    If m.Chart.TypeOfChart = eTypeChart_Seasonal Then           '6401
        If Len(g.strActiveDraw) > 0 Then
            Select Case g.strActiveDraw
                Case "ID_Trendline", "ID_Trendline2", "ID_Trendline3", "ID_Trendline4", _
                     "ID_TrendChannel", "ID_SRLine", "ID_HorzLine", "ID_VertLine", "ID_ArrowLine", _
                     "ID_Text", "ID_Text2", "ID_Text3", "ID_Text4", "ID_Icon", "ID_ElliotLabels", _
                     "ID_Bracket", "ID_Ellipse", "ID_Rectangle", "ID_ElliotEndUser"
                     
                     'tools ok - continue
                     
                Case Else
                    InfBox kSeasonalUnavail, "I", "Ok", "Seasonal chart"
                    SyncDrawTools
            End Select
        End If
    End If
    
    ActiveChartFormSet Me
    If m.eDetachStatus = eNotDetached Then Set g.ChartGlobals.frmActiveNonDetached = Me
    m.bSkipFocusFix = False
    
    '04-16-2009: Request from LW per Chad
    If FormIsLoaded("frmPatternProfit") Then frmPatternProfit.WindowState = vbMinimized
    
    If m.bGameMode Then
        m.eReplayModeSave = m.oGameMode.GameReplayMode(tmrGameMode.Enabled)
        cmdStop_Click
    End If
    
    TopMost = False
    m.nObjectMoving = 0
    m.nScaleStartPixel = 0
    m.nScrollChange = 0
    m.bChartMoveInProg = False
    m.bDrawToolSelected = False
    
    If FormIsLoaded("frmTbMoreButtons") Then Unload frmTbMoreButtons
                        
    If Not ActiveChart Is Me Then
        m.MouseDown.bOffChart = True
        If Button = vbRightButton Then BringWindowToTop Me.hWnd
        bSkipHitTest = True
    ElseIf FormIsLoaded("frmTemplatePage") Then
        Unload frmTemplatePage
        DoEvents
        Exit Sub
    End If
        
    m.MouseDown = GetChartCoordinates(X, Y, Shift)
    
    m.MouseDown.nButton = Button
    
    If ActiveChart Is Me Then
        frmChartData.ShowData m.MouseDown.nX
        frmPlanetData.ShowData m.MouseDown.dDate, m.Chart.Bars
    End If
                      
    If Button = vbRightButton Then
        ClearBuySellButtons
        DoHitTest X, Y, True, Button
        If m.nActiveAnnotIdx > 0 Then
            If Len(g.strActiveDraw) > 0 Then
                If g.strActiveDraw = "ID_ElliotLabels" Then
                    Dim labelAnnot As cAnnotation
                    Set labelAnnot = m.Chart.Annots(m.nActiveAnnotIdx)
                    
                    If labelAnnot Is Nothing Then
                        ClearAnnotFlags True, False             '6283
                    ElseIf labelAnnot.eType = eANNOT_ElliotLabel Then
                        '02-27-2014 aardvark 5963 - delete annot if it is an ewi label
                        labelAnnot.geRemoveAnnotation m.Chart.geChartObj
                        m.Chart.Annots.Remove (m.nActiveAnnotIdx)
                        Set labelAnnot = Nothing
                        ClearAnnotFlags False, False
                        m.Chart.GenerateChart eRedo1_Scrolled
                    Else
                        ClearAnnotFlags True, False             '6283
                    End If
                
                Else
                    ClearAnnotFlags True, False             '6283
                End If
            Else
                'active draw string length is zero
                ShowAnnotPopup 0, vbRightButton
                ClearAnnotFlags False, False
                PbChart_MouseUp Button, Shift, X, Y
            End If
            
            Exit Sub
        ElseIf m.nActiveOrderID > 0 Then
            ShowOrderActionMenu
            Exit Sub
        ElseIf Shift = 0 And m.epbCursor <> eCursor_HorzSize And Not m.MouseDown.bOffChart Then
            BringWindowToTop Me.hWnd
            ShowPopup
            Exit Sub
        End If
    ElseIf Shift = 0 Then
        If Len(g.strActiveDraw) > 0 Then
            If MouseDownActiveDraw() Then
                Exit Sub    'a multi-click drawing tool is in progress
            End If
            If (g.strActiveDraw = "ID_Icon" Or g.strActiveDraw = "ID_ElliotLabels" Or g.strActiveDraw = "ID_ElliotEndUser") And _
               (m.epbCursor = eCursor_Hand Or m.epbCursor = eCursor_HandDown) Then
                'Note: eCursor_Hand - cursor is over a hot-spot on chart or in upper half of y-scale area
                '      eCursor_HandDown - cursor is over a hot-spot on chart or is in lower half of y-scale area
                'icon draw in progress - allow hand cursor for moving icon icon around
                '                        cursor changes to arrow when in y-scale area (dragging y-scale not allowed)
                'elliott label draw in progress - allow hand & hand_down cursor for moving elliott label & dragging in y-scale area
                m.nActiveAnnotIdx = 0
            ElseIf Len(g.strActiveDraw) > 0 And InStr(g.strActiveDraw, "PFP") = 0 Then
                m.nActiveAnnotIdx = -1
            ElseIf Me.fraPatternProfit.Visible Then
                m.nActiveAnnotIdx = -1          '5809
            End If
        End If
    End If

    'check for move/drag in scale areas
    If m.MouseDown.bOffChart = True And m.MouseDown.nScalePaneId <> 0 And m.Chart.Zoomed = False Then
        If m.MouseDown.nScalePaneId = -7 Then
            m.eScaleFlag = eScale_Arrow
        ElseIf m.MouseDown.nScalePaneId = -8 Then
            m.eScaleFlag = eScale_Arrow_Area
        Else
            SetScaleFlag eNewScaleFlag, True
            If m.eScaleFlag <> eScale_Unhandled Then
                m.nObjectMoving = 1
                SetScaleCursor m.eScaleFlag, True
                ShowCursor
            End If
        End If
        Exit Sub
    End If
        
    Dim iHwnd&
    
    If vseOrderBar.Visible And m.eOrdBarMode = eOrdBarMode_Wizard And _
        (vseBuyWizard.Appearance = apInset Or vseSellWizard.Appearance = apInset) Then      '4977
        If Not fraWizardPrompt.Visible And pbChart.MousePointer = vbCustom Then
            lblWizardPrice_MouseUp Button, Shift, X, Y
        Else
            'do nothing     -aardvark 4961
        End If
    ElseIf m.nActiveAnnotIdx <> 0 Then
        m.Chart.SetCursor 3 'pencil
        'user starting annot draw make sure global mode is not place order (aardvark 3585 fix)
        If g.ChartGlobals.eChartMode = eMode_ChartOrder Then g.ChartGlobals.eChartMode = g.ChartGlobals.ePrevChartMode
    Else
        'user extending horz lines left/right - skip hittest (aardvark 2579 fix)
        If m.epbCursor = eCursor_ArrowEW Then
            Set Annot = m.Chart.Tree(m.nActiveAnnotIdx)
            If Not Annot Is Nothing Then
                If m.nActiveAnnotPt > 0 Then
                    If Annot.eType = eANNOT_Fibonacci Or Annot.eType = eANNOT_Fibonacci2 Or Annot.eType = eANNOT_Fibonacci3 _
                       Or Annot.eType = eANNOT_Fibonacci4 Or Annot.eType = eANNOT_DanCodeFib Then
                        Exit Sub
                    End If
                End If
            End If
        ElseIf m.epbCursor = eCursor_OrderBuy Or m.epbCursor = eCursor_OrderSell Then
            iHwnd = ValOfText(frmMain.tmrCheckBuySellButtons.Tag)
            If iHwnd = Me.hWnd And vseOrderBar.Visible Then
                HandleChartOrder True
            Else
                ClearAllBuySellBtns
            End If
            Exit Sub
        End If
        
        If Not bSkipHitTest Then
            'if user chose toolbutton from some other chart then immeidately click
            'this chart then skip the hittest else chart may go into move mode
            DoHitTest X, Y, True, Button
            If m.nActiveAnnotIdx <= 0 And m.nActiveBtmPane <= 0 Then
                m.bPrevOffChart = True
                ShowCursor
            ElseIf m.epbCursor = eCursor_HorzSize And Button = vbRightButton Then
                Set Pane = m.Chart.Tree(m.nActiveBtmPane)
                If Not Pane Is Nothing Then
                    If Pane.HideSeparator = 0 Then
                        Pane.HideSeparator = 1
                    Else
                        Pane.HideSeparator = 0
                    End If
                    Me.pbChart.AutoRedraw = False
                    geDrawSeparator m.Chart.geChartObj, Me.pbChart.hDC, m.nActiveBtmPane, 1, Pane.HideSeparator
                    Me.pbChart.AutoRedraw = True
                End If
    
                Set Pane = Nothing
                m.nActiveBtmPane = 0
                m.nActiveTopPane = 0
            Else
                Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
                If Not Annot Is Nothing Then
                    Annot.SetMoveOffSet m.MouseDown.nX, m.MouseDown.dY, Shift
                    If Annot.eType = eANNOT_WaveLabels Then HandleWaveLabels False, False, True
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".pbChart_MouseDown"
    Resume ErrExit
    
End Sub

Public Sub ShowCursor()
On Error GoTo ErrSection:
                                                    
    Static eCursorPrev As enumCursor
    

    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If m.Chart Is Nothing Then Exit Sub

    If m.epbCursor = eCursor_OrderBuy Then
        If m.Chart.ChartBkIsLight Then
            If vseBracketOrder.Appearance = apInset Then
                pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderBuySell"))
            Else
                pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderBuy"))
            End If
        ElseIf vseBracketOrder.Appearance = apInset Then
            pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderBuySellWhite"))
        Else
            pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderBuyWhite"))
        End If
    ElseIf m.epbCursor = eCursor_OrderSell Then
        If m.Chart.ChartBkIsLight Then
            If vseBracketOrder.Appearance = apInset Then
                pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderBuySell"))
            Else
                pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderSell"))
            End If
        ElseIf vseBracketOrder.Appearance = apInset Then
            pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderBuySellWhite"))
        Else
            pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderSellWhite"))
        End If
    ElseIf Not ActiveChart Is Nothing Then
        'bracket order is in progress on active chart, don't change cursor
        If ActiveChart.vseBracketOrder.Appearance = apInset Then m.epbCursor = eCursorPrev
    End If
    
    If m.nActiveAnnotIdx < 0 Then
        m.epbCursor = eCursor_Pencil
    ElseIf Len(g.strActiveDraw) > 0 Then
        If InStr(g.strActiveDraw, "PFP") = 0 Then       '5809
            If g.strActiveDraw <> "ID_Icon" And g.strActiveDraw <> "ID_ElliotLabels" And g.strActiveDraw <> "ID_ElliotEndUser" Then
                m.epbCursor = eCursor_Pencil
            End If
        ElseIf fraPatternProfit.Visible And Me Is g.ChartGlobals.frmPfpSelPattern Then
            m.epbCursor = eCursor_Pencil
        ElseIf Me Is g.ChartGlobals.frmLastChartMouseMove Then
            SyncDrawTools           '5807
        End If
    ElseIf m.nActiveAnnotIdx > 0 And eCursorPrev = eCursor_Blank Then
        Exit Sub
    End If
        
    With frmMain.tbToolbar
        If .Tools("ID_CursorArrow").State = ssChecked Or m.epbCursor = eCursor_Pencil _
                Or m.epbCursor = eCursor_NoDrop Or m.epbCursor = eCursor_Magnify Then
            m.eCrossHairOn = eCursor_Default
        ElseIf .Tools("ID_CursorCrosshairs").State = ssChecked Then
            If m.eOrdBarMode = eOrdBarMode_Wizard And m.epbCursor = eCursor_Default Then
                m.eCrossHairOn = eCursor_Horizontal     'wizard prompt is active, turn off vertical crosshair
            Else
                m.eCrossHairOn = eCursor_CrossHair
            End If
        ElseIf .Tools("ID_CursorVertLine").State = ssChecked Then
            If m.eOrdBarMode = eOrdBarMode_Wizard And m.epbCursor = eCursor_Default Then
                m.eCrossHairOn = eCursor_Horizontal
            Else
                m.eCrossHairOn = eCursor_Vertical
            End If
        ElseIf .Tools("ID_CursorHorizLine").State = ssChecked Then
            m.eCrossHairOn = eCursor_Horizontal
        Else
            m.eCrossHairOn = eCursor_Default
        End If
    End With

    If m.epbCursor <> eCursorPrev Or m.epbCursor = eCursor_Pencil Then
        eCursorPrev = m.epbCursor
        pbChart.MousePointer = 99
        Select Case m.epbCursor
            Case eCursor_Default
                pbChart.MousePointer = 1
            'case 2 - not used, this is built-in, crosshair cursor id for picturebox
            Case eCursor_Pencil
                pbChart.MouseIcon = Picture16(ToolbarIcon("kPencil"))
                If g.strActiveDraw = "ID_WaveLabels" Then
                    WaveLabelsCreate
                End If
            Case eCursor_Hand
                pbChart.MouseIcon = Picture16(ToolbarIcon("kHand"))
            Case eCursor_CrossHair, eCursor_Vertical, eCursor_Horizontal
                pbChart.MousePointer = 2 '1
            Case eCursor_ArrowNS
                pbChart.MousePointer = vbSizeNS
            Case eCursor_ArrowEW
                pbChart.MousePointer = vbSizeWE
            Case eCursor_ArrowNE
                pbChart.MousePointer = vbSizeNESW
            Case eCursor_ArrowNW
                pbChart.MousePointer = vbSizeNWSE
            Case eCursor_HorzSize
                pbChart.MouseIcon = Picture16(ToolbarIcon("kHzSize"))    'separator drag-n-drop
            Case eCursor_NoDrop
                pbChart.MousePointer = vbNoDrop
            Case eCursor_Magnify
                pbChart.MouseIcon = Picture16(ToolbarIcon("kMagnify"))
            Case eCursor_Arrow4Way
                pbChart.MousePointer = vbSizePointer
            Case eCursor_HandMoveUp
                pbChart.MouseIcon = Picture16(ToolbarIcon("kHandMoveUp"))
            Case eCursor_HandMoveDown
                pbChart.MouseIcon = Picture16(ToolbarIcon("kHandMoveDown"))
            Case eCursor_HandMoveLeft
                pbChart.MouseIcon = Picture16(ToolbarIcon("kHandMoveLeft"))
            Case eCursor_HandMoveRight
                pbChart.MouseIcon = Picture16(ToolbarIcon("kHandMoveRight"))
            Case eCursor_HandDown
                pbChart.MouseIcon = Picture16(ToolbarIcon("kHandDown"))
            Case eCursor_HandLeft
                pbChart.MouseIcon = Picture16(ToolbarIcon("kHandLeft"))
            Case eCursor_HandRight
                pbChart.MouseIcon = Picture16(ToolbarIcon("kHandRight"))
            Case eCursor_ChartMove
                pbChart.MouseIcon = Picture16(ToolbarIcon("kDragMove"))
            Case eCursor_Eraser
                pbChart.MouseIcon = Picture16(ToolbarIcon("kCursorEraser"))
            Case eCursor_Blank
                pbChart.MouseIcon = Picture16("kBlank")
        End Select
    End If
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ShowCursor", eGDRaiseError_Raise
    
End Sub

Public Property Get cChartObj() As cChart
On Error GoTo ErrSection:

    Set cChartObj = m.Chart

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".cChartObj.Get", eGDRaiseError_Raise
    
End Property

Public Property Get pbCursor() As enumCursor
On Error GoTo ErrSection:

    pbCursor = m.epbCursor

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".pbCursor.Get", eGDRaiseError_Raise
   
End Property

Public Property Let pbCursor(ByVal eCursor As enumCursor)
On Error GoTo ErrSection:

    Dim iHwnd&
        
    If eCursor = eCursor_OrderBuy Or eCursor = eCursor_OrderSell Then
        iHwnd = ValOfText(frmMain.tmrCheckBuySellButtons.Tag)
        If Me.hWnd <> iHwnd Then
            Exit Property
        End If
    End If
        
    If eCursor <> m.epbCursor Then
        m.Chart.geDrawChart
        Me.pbChart.Refresh
    End If
    m.epbCursor = eCursor
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".pbCursor.Let", eGDRaiseError_Raise
    
End Property

Private Sub ResizePanes(ByVal X As Single, ByVal Y As Single, ByVal Shift As Integer)
On Error GoTo ErrSection:

    'active pane is the pane below the separator that was hit
    'Current pane separator dragDrop behavior:
    '1. separator dropped above chart's area --> no change
    '2. separator dropped below chart's area
    '   --> active pane is removed
    '3. separator dropped into pane immediately above active pane
    '   --> active pane is made larger, pane above is made smaller
    '4. separator dropped into active pane (i.e. same pane)
    '   --> active pane is made smaller, pane above is made larger
    '5. separator dropped into pane NOT immediately above active pane
    '   --> active pane moved below pane the separator was dropped in
    
    Dim btmPane As cPane, topPane As cPane, PricePane As cPane
    Dim nTop&, nLeft&, nBtm&, nRight&, nHeight&     'dimension of chart's area
    Dim pixY&                   'pixel value of current cursor position
    Dim dSizeOld#               'old height proportion of pane being resized
    Dim bMove As Boolean, i&
    
    bMove = False
    If m.nActiveBtmPane > 0 Then Set btmPane = m.Chart.Tree(m.nActiveBtmPane)
    If btmPane Is Nothing Then Exit Sub
    
    m.MouseLast = GetChartCoordinates(X, Y)
    
    pixY = Y \ Screen.TwipsPerPixelY
    m.Chart.geGetChartDim nTop, nLeft, nBtm, nRight
    nHeight = nBtm - nTop           'this is height of chart area
    
    If pixY < nTop Then Exit Sub    'case 1: no change
    
    If pixY > nBtm Then
        btmPane.Display = False     'case 2: hide pane
        Exit Sub
    End If
    
    Set topPane = Nothing
    Set PricePane = Nothing
    
    If m.nActiveTopPane > 0 Then
        If m.Chart.Tree.Exists(m.nActiveTopPane) Then
            'check for CTRL key
            If Shift = 2 And m.Chart.Tree.Key(m.nActiveTopPane) = "PRICE PANE" Then
                Set PricePane = m.Chart.Tree(m.nActiveTopPane)
            Else
                Set topPane = m.Chart.Tree(m.nActiveTopPane)
            End If
        End If
    End If
    
    'check for case 5
    If PricePane Is Nothing Then
        If m.MouseLast.nPaneID <> m.nActiveTopPane And _
           m.MouseLast.nPaneID <> m.nActiveBtmPane Then
            bMove = True
        End If
    End If
    
    If bMove Then
        'case 5: move pane -
        'active top pane can be -1 if separator at chart's top edge is dragged
        'make pane below separator the next sibling of pane separator was dropped in
        m.Chart.geResetPanes
        m.Chart.Tree.Move m.nActiveBtmPane, m.MouseLast.nPaneID, eTREE_NextSibling
    ElseIf PricePane Is Nothing Then
        'cases 3,4: resize pane
        dSizeOld = btmPane.Size
        btmPane.Size = (btmPane.geBtmSep - pixY) / nHeight  'this is height of pane as a percentage of chart's height
        
        'if there is a pane above the separator being dragged then
        'reduce or increase this pane's proportion by same amount
        If Not topPane Is Nothing Then
            topPane.Size = topPane.Size + (dSizeOld - btmPane.Size)
        End If

'JM 04-15-2011: not sure why this is here, but it causes issue 6190
'   commenting this out fixes 6190; leave awhile then remove if all ok
'        If Not topPane Is Nothing Then
'            topPane.Size = m.Chart.Tree("PRICE PANE").Size * 0.06
'        End If

        m.Chart.geForceRecalc
    Else
        'CTRL key is held down on price pane
        PricePane.Size = Abs(pixY - PricePane.geTopSep) / nHeight
        m.Chart.SetEqualSizePanes PricePane
        If Not topPane Is Nothing Then
            topPane.Size = PricePane.Size * 0.06
        End If
        m.Chart.geForceRecalc
    End If
        
ErrExit:
    Set btmPane = Nothing
    Set topPane = Nothing
    Set PricePane = Nothing
    Exit Sub
    
ErrSection:
    Set btmPane = Nothing
    Set topPane = Nothing
    Set PricePane = Nothing
    RaiseError Me.Name & ".ResizePanes", eGDRaiseError_Raise
    
End Sub

Private Sub SetScaleCursor(ByVal eFlagval As enumScaleFlag, ByVal bMouseDown As Boolean)
On Error GoTo ErrSection:

    Dim nTop&, nLeft&, nBtm&, nRight&
    Dim nPixX&, nPixY&
    
    If eFlagval = eScale_Unhandled Or m.Chart.Zoomed = True Then Exit Sub
            
    'check to see if cursor is in lower right corner of window
    'i.e. not on chart and not in scale areas
    m.Chart.geGetChartDim nTop, nLeft, nBtm, nRight, True
    nPixX = m.MouseLast.MouseX / Screen.TwipsPerPixelX
    nPixY = m.MouseLast.MouseY / Screen.TwipsPerPixelY
    
    If nPixX > nRight And nPixY > nBtm Then
        eFlagval = eScale_Unhandled
        m.Chart.SetCursor
        Exit Sub
    End If
    
    If bMouseDown = True Then
        Select Case m.eScaleFlag
            Case eScaleY_Max
                m.epbCursor = eCursor_HandMoveUp
            Case eScaleY_Min
                m.epbCursor = eCursor_HandMoveDown
            Case eScaleX_MoreLessBars
                m.epbCursor = eCursor_HandMoveLeft
            Case Else
                m.nObjectMoving = 0
        End Select
        Exit Sub
    End If
            
    Select Case eFlagval
        Case eScaleY_Max
            m.epbCursor = eCursor_Hand
        Case eScaleY_Min
            m.epbCursor = eCursor_HandDown
        Case eScaleX_MoreLessBars
            m.epbCursor = eCursor_HandLeft
    End Select
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetScaleCursor", eGDRaiseError_Raise
    
End Sub

Private Sub SetScaleFlag(eFlagval As enumScaleFlag, ByVal bMouseDown As Boolean)
On Error GoTo ErrSection:

    Dim Pane As cPane
    Dim MouseInfo As ChartCoordinates
    Dim geCoordInfo As coordinate_info
    Dim rc&, nTop&, nLeft&, nBtm&, nRight&
    Dim dYMid#, dPixMid#
    
    m.eScaleFlag = eScale_Unhandled
    eFlagval = eScale_Unhandled
    m.nScaleStartPixel = 0
    m.nBarsToRight = 0
    
    If m.Chart.Zoomed = True Then Exit Sub
    
    If bMouseDown = True Then
        MouseInfo = m.MouseDown
    Else
        MouseInfo = m.MouseLast
    End If
    
    If MouseInfo.nScalePaneId = ePANE_LogFlagLog Then
        eFlagval = eScaleX_MoreLessBars
        If bMouseDown = True Then   '3834 - show hand cursor even if in blank bar/timer area
            m.nScaleStartPixel = m.MouseDown.MouseX / Screen.TwipsPerPixelX
            m.nBarsToRight = m.Chart.geChartPoints - m.MouseDown.nScreenX
        End If
    ElseIf m.Chart.Tree.NodeLevel(MouseInfo.nScalePaneId) = 0 Then
        Set Pane = m.Chart.Tree(MouseInfo.nScalePaneId)
    End If
    
    If Pane Is Nothing Then
        If bMouseDown = True Then m.eScaleFlag = eFlagval
        Exit Sub
    End If
        
    If g.ChartGlobals.eDragModeY = eDragModeY_Each Then
        'flag whether to change y-scale's min or max
        If Pane.PaneLogFlag = ePANE_LogFlagLog Then
            m.Chart.geGetChartDim nTop, nLeft, nBtm, nRight
            dPixMid = (Pane.geTopSep + Pane.geBtmSep) / 2
            geCoordInfo.paneId = Pane.gePaneId
            geCoordInfo.x_pixels = nRight - 5   'don't care
            geCoordInfo.y_pixels = dPixMid      'this is what we are really after
            
            rc = geCoordToData(m.Chart.geChartObj, geCoordInfo)
            If rc = 0 Then
                dYMid = geCoordInfo.y_value
            Else
                Exit Sub
            End If
        Else
            dYMid = (Pane.gePaneMax + Pane.gePaneMin) / 2
        End If
        
        If MouseInfo.dY > dYMid Then
            eFlagval = eScaleY_Max
        Else
            eFlagval = eScaleY_Min
        End If
    Else
        eFlagval = eScaleY_Max
    End If
    
    If bMouseDown = True Then m.eScaleFlag = eFlagval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetScaleFlag", eGDRaiseError_Raise

End Sub

Private Sub HandleScaleMoveX(ByVal X As Single, ByVal Y As Single, Shift As Integer)
On Error GoTo ErrSection:

    Static nPrevX As Single
    
    If m.eScaleFlag <> eScaleX_MoreLessBars Then Exit Sub   'precautionary, should never happen
    
    If X < nPrevX Then
        m.Chart.PixelsPerBar = -1   'more bars (increase pix per bar)
        m.Chart.GenerateChart eRedo1_Scrolled
    ElseIf X > nPrevX Then
        m.Chart.PixelsPerBar = -2   'less bars (decrease pix per bar)
        m.Chart.GenerateChart eRedo1_Scrolled
    End If
    
    nPrevX = X
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".HandleScaleMoveX", eGDRaiseError_Raise

End Sub

Private Sub HandleScaleMoveY(ByVal X As Single, ByVal Y As Single, Shift As Integer, Optional ByVal bReset As Boolean = False)
On Error GoTo ErrSection:

    Static nPrevY As Single
    Dim Pane As cPane
    
    If bReset Then
        Set Pane = m.Chart.Tree(m.Chart.Tree.Key("PRICE PANE"))
        If Not Pane Is Nothing Then
            If Pane.Scaling <> ePANE_ScaleModeManual Then
                Pane.Scaling = ePANE_ScaleModeManual
                m.Chart.SyncToolbar
            End If
        End If
        geAnnotMove m.Chart.geChartObj, 2
        m.Chart.GenerateChart eRedo1_Scrolled       'fix auto-text on $Line not repositioning when squishing scale right after chart page change
        nPrevY = -1
        Exit Sub
    End If
    
    geAnnotMove m.Chart.geChartObj, 1
    If m.MouseDown.nScalePaneId = m.MouseLast.nScalePaneId Then
        If nPrevY = -1 Then
            nPrevY = Y
        ElseIf m.Chart.Tree.NodeLevel(m.MouseLast.nScalePaneId) = 0 Then
            Set Pane = m.Chart.Tree(m.MouseLast.nScalePaneId)
            If Not Pane Is Nothing Then
                If g.ChartGlobals.eDragModeY = eDragModeY_Each Then
                    Select Case m.eScaleFlag
                        Case eScaleY_Max
                            If nPrevY < Y Then
                                Pane.geIncDecMaxRatio 0.06, True
                            ElseIf nPrevY > Y Then
                                Pane.geIncDecMaxRatio -0.06, True
                            End If
                        Case eScaleY_Min
                            If nPrevY < Y Then
                                Pane.geIncDecMinRatio 0.06, True
                            ElseIf nPrevY > Y Then
                                Pane.geIncDecMinRatio -0.06, True
                            End If
                    End Select
                Else
                    If nPrevY < Y Then
                        Pane.geIncDecMaxRatio 0.1, True     'scrunching (F11)
                        Pane.geIncDecMinRatio -0.1, True
                    ElseIf nPrevY > Y Then
                        Pane.geIncDecMaxRatio -0.1, True
                        Pane.geIncDecMinRatio 0.1, True
                    End If
                End If
                nPrevY = Y
                m.Chart.GenerateChart eRedo1_Scrolled
            End If
        End If
    End If
    
    Set Pane = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".HandleScaleMoveY", eGDRaiseError_Raise

End Sub

Private Sub HandleZoom(ByVal X As Single, ByVal Y As Single, ByVal bDrawRect As Boolean)
On Error GoTo ErrSection:

    Dim nTop&, nLeft&, nBtm&, nRight&
    Dim nHeight&, nWidth&, nMinWidth&, rc&
        
        
    m.MouseLast = GetChartCoordinates(X, Y)
        
    If m.MouseLast.bOffChart And m.MouseLast.dY = 0 Then
        'cursor is above chart area where grapheng.dll does not get mouse messages
        'call draw zoom rect with -1 to clear zoom variables
        geDrawZoomRect m.Chart.geChartObj, pbChart.hDC, -1, -1, -1, -1
        m.Chart.SetCursor
        Exit Sub
    End If
    
    nTop = m.MouseDown.MouseY / Screen.TwipsPerPixelY
    nLeft = m.MouseDown.MouseX / Screen.TwipsPerPixelX
    nBtm = Y / Screen.TwipsPerPixelY
    nRight = X / Screen.TwipsPerPixelX
    
    If Chart.Zoomed = True Then
        nMinWidth = geZoomPixPerBar(m.Chart.geChartObj, pbChart.hDC, -1, 0)
    Else
        nMinWidth = Chart.PixelsPerBar * 3
    End If
    nWidth = Abs(nRight - nLeft)
    nHeight = Abs(nBtm - nTop)
        
    If nWidth < nMinWidth Or nHeight < 10 Then
        If m.epbCursor = eCursor_NoDrop Then Exit Sub
        If bDrawRect = True Then
            'user has decreased the size of a previously valid zoomed area
            'call draw zoom rect with -1 to clear zoom variables
            geDrawZoomRect m.Chart.geChartObj, pbChart.hDC, -1, -1, -1, -1
            m.epbCursor = eCursor_NoDrop
            Exit Sub
        End If
    End If
    
    If bDrawRect = True Then
        'clear cross hair cursor before drawing zoom rectangle
        geSyncCrossHairEx m.Chart.geChartObj, pbChart.hWnd, pbChart.hDC, kNullData, kNullData, kNullData, 1, 1, 1
        rc = geDrawZoomRect(m.Chart.geChartObj, Me.pbChart.hDC, _
            nTop, nLeft, nBtm, nRight)
        Exit Sub
    ElseIf m.epbCursor = eCursor_NoDrop Then
        m.epbCursor = eCursor_Magnify         'Zoom cursor
        Exit Sub
    End If
                                        
    rc = geDrawZoomChart(m.Chart.geChartObj, 0, Me.pbChart.hDC, nTop, nLeft, nBtm, nRight)
    
    If rc <> 0 Then Exit Sub
    
    Dim dLastGoodDate#
       
    If m.Chart.Zoomed = False Then
        m.Chart.Zoomed = True
        m.nHsbMaxSave = hsb.Max
    End If
    If m.MouseLast.bOffChart = False And m.MouseDown.bOffChart = False Then
        dLastGoodDate = m.Chart.Bars(eBARS_DateTime, m.Chart.LastGoodDataBar(False))
        If m.MouseLast.dDate < dLastGoodDate And m.MouseDown.dDate < dLastGoodDate Then
            m.Chart.aXdate.BinarySearch gdFixDateTime(m.MouseLast.dDate), nLeft
            m.Chart.aXdate.BinarySearch gdFixDateTime(m.MouseDown.dDate), nRight
            If nRight > nLeft Then
                rc = hsb.Value - nRight
            Else
                rc = hsb.Value - nLeft
            End If
            hsb.Max = m.nHsbMaxSave + rc - m.Chart.ForecastBars + 2
        End If
    End If
    
    m.Chart.SyncToolbar
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".HandleZoom", eGDRaiseError_Raise
    
End Sub

Private Sub HandleAnnotDblclk()
On Error GoTo ErrSection:

    Dim Annot As cAnnotation, NewAnnot As cAnnotation
    Dim OrigPatternAnnot As cAnnotation
    Dim Ind As cIndicator
    Dim OtherCluster As cIndicator
    Dim Pane As cPane
    Dim i&, idx&, strKey$, dDiff#, dMinMove#
    
    Dim bRemoveClusters As Boolean
    
    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    
    If Annot Is Nothing Then
        Exit Sub
    ElseIf Annot.eUsage = eANNOT_WhatIf Then
        Exit Sub       '4293 - don't allow duplicate whatif annotations
    ElseIf Annot.eUsage = eANNOT_FibClusters Then
        Exit Sub
    End If
    
    tmr.Tag = ""
        
    If Annot.eUsage = eANNOT_IndicatorLabel Then        '4575
        frmChartOnOff.ClearPrevChart
        strKey = m.Chart.Annots.Key(m.nActiveAnnotIdx)
        If UCase(strKey) = "SYSTEM NAME" Then
            m.Chart.ShowTrades = False
        Else
            idx = Val(strKey)
            Set Ind = m.Chart.Tree(idx)
            If Not Ind Is Nothing Then
                Ind.Display = False
                If Ind.DisplayType = eINDIC_Ribbon Then
                    RibbonList Nothing, m.Chart, Ind, 2, 0
                ElseIf m.Chart.Tree.Key(idx) = kClusterTimeKeyInd Or m.Chart.Tree.Key(idx) = kClusterPriceKey Then
                    If Ind.DisplayType = eINDIC_ClusterPrice Then
                        Set OtherCluster = m.Chart.Tree(kClusterTimeKeyInd)
                        m.Chart.ResetSplitPane
                        m.Chart.RestoreChartNormal vbKeyClear
                    Else
                        Set OtherCluster = m.Chart.Tree(kClusterPriceKey)
                    End If
                    
                    If OtherCluster Is Nothing Then
                        bRemoveClusters = True
                    ElseIf Not OtherCluster.Display Then
                        bRemoveClusters = True
                    End If
                End If
            End If
        End If
        Annot.geRemoveAnnotation (m.Chart.geChartObj)
        m.Chart.Annots.Remove m.nActiveAnnotIdx
        
        If bRemoveClusters Then m.Chart.RemoveFibClusters
        m.Chart.GenerateChart
        
        If g.RealTime.Active Then
            m.Chart.geForceRecalc       'don't want to call this in generatechart when realtime is on (5173)
            m.Chart.GenerateChart eRedo1_Scrolled
        End If
        
        Set Ind = Nothing
        Set Annot = Nothing
        m.nActiveAnnotIdx = 0
        Exit Sub
    ElseIf Annot.eType = eANNOT_WaveLabels Then
        HandleWaveLabels False, True
        Exit Sub
    ElseIf Not Annot.DuplicateAllow Then
        m.nActiveAnnotIdx = 0
        StatusMsg "This annotation cannot be duplicated."
        Set Annot = Nothing
        Exit Sub
    End If
        
    Set Pane = m.Chart.Tree(Annot.gePaneId)
    If Pane Is Nothing Then Exit Sub
    
    ' make a duplicate annotation and adjust location a little (so shows)
    Set NewAnnot = Annot.MakeCopy
    If Annot.eType = eANNOT_VertLine Or Annot.eType = eANNOT_RiskReward Then
        ' go back one bar
        m.Chart.aXdate.BinarySearch gdFixDateTime(Annot.dDate(1)), i
        dDiff = m.Chart.aXdate(i - 1)
        NewAnnot.dDate(1) = dDiff
        NewAnnot.dDate(2) = dDiff
        NewAnnot.geMoveFlag = 1
    Else
        ' move up 1/20th of the pane size
        dDiff = 0.05 * (Pane.gePaneMax - Pane.gePaneMin)
        If Annot.eType = eANNOT_DollarLine Or Annot.eType = eANNOT_DollarLine2 Or _
           Annot.eType = eANNOT_DollarLine3 Or Annot.eType = eANNOT_DollarLine4 Or _
           Annot.eType = eANNOT_GannacciSwingSquare Then
            ' for $Line, make sure diff is at least a min move of the price
            dMinMove = m.Chart.Bars.MinMove(Annot.dDate(1))
            If dDiff < dMinMove Then
                dDiff = dMinMove
            End If
        End If
        
        NewAnnot.Y(1) = Annot.Y(1) + dDiff
        If Pane.PaneLogFlag = -1 And Annot.Y(1) > 0 And Annot.Y(2) > 0 Then
            NewAnnot.Y(2) = Exp(Log(Annot.Y(2)) + Log(NewAnnot.Y(1)) - Log(Annot.Y(1)))
        Else
            NewAnnot.Y(2) = Annot.Y(2) + dDiff
        End If
        If Annot.eType = eANNOT_TextEdit Or Annot.eType = eANNOT_TextEdit2 Or _
           Annot.eType = eANNOT_TextEdit3 Or Annot.eType = eANNOT_TextEdit4 Then
            NewAnnot.geMoveFlag = 1   'aardvark 1225 fix
        ElseIf Annot.eType = eANNOT_TriangleWedge Or Annot.eType = eANNOT_AndrewFork Then
            NewAnnot.YFromArray(0) = NewAnnot.YFromArray(0) + dDiff
            NewAnnot.geMoveFlag = 1
        ElseIf Annot.eType = eANNOT_ChannelHighlight Then
            NewAnnot.YFromArray(0) = NewAnnot.YFromArray(0) + dDiff
            NewAnnot.YFromArray(1) = NewAnnot.YFromArray(1) + dDiff
            NewAnnot.geMoveFlag = 1
        End If
    End If
        
    ' add to annots and save key
    NewAnnot.geAddAnnotation m.Chart, Annot.gePaneId, m.Chart.Annots.Count + 1    'Q & A - valid assumption here?
    i = m.Chart.Annots.Add(NewAnnot)
    strKey = m.Chart.Annots.Key(i)
    NewAnnot.Prop("AnnotKey") = strKey
    NewAnnot.geAnnId = i
    
    m.Chart.LastEditCreate NewAnnot, True
                
    'clear flags
    m.nActiveAnnotPt = 0
    m.nActiveAnnotIdx = 0
    
    Set NewAnnot = Nothing
    Set Annot = Nothing
    Set Pane = Nothing
    
    m.Chart.GenerateChart eRedo1_Scrolled

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".HandleAnnotDblClk", eGDRaiseError_Raise
    
End Sub

Private Function HandleAndrewFork(ByVal bMouseUp As Boolean) As Boolean
On Error GoTo ErrSection:

    Static nPrevX&, nPrevY&
    Dim Annot As cAnnotation
    Dim bMoved As Boolean
    Dim dY#

    HandleAndrewFork = False
        
    'check if this is initial create
    If m.nActiveAnnotIdx < 0 Then
        StatusMsg "Now click on the second point ...", -1
        m.nPointCount = 1
        HandleAndrewFork = True
        nPrevX = m.MouseLast.MouseX / Screen.TwipsPerPixelX
        nPrevY = m.MouseLast.MouseY / Screen.TwipsPerPixelY
        Exit Function
    End If
    
    'check if editing
    If bMouseUp = True Then
       If Len(g.strActiveDraw) = 0 Then
            If m.nObjectMoving < kMinMoveCount Then
                tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
            Else
                Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
                If Not Annot Is Nothing Then m.Chart.SyncGlobalAnnots Annot
            End If
            ClearAnnotFlags False
            m.nActiveIndIdx = 0
            m.nObjectMoving = 0
            m.Chart.SetCursor
            Exit Function
        End If
    End If

    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    If Annot Is Nothing Then Exit Function
    
    If m.MouseLast.nPaneID <> Annot.gePaneId Then
        If m.bAnnotCreated And bMouseUp Then
            ClearAnnotFlags True
            Exit Function
        End If
    End If
    
    dY = SnapToPrice(m.MouseLast, Annot.eType)
    If Annot.geMoveFlag = 0 And bMouseUp = True Then
        'user still selecting points
        Select Case m.nPointCount
            Case 1:
                If Abs(m.MouseLast.MouseX / Screen.TwipsPerPixelX - nPrevX) > kMinAnnotSize Or _
                   Abs(m.MouseLast.MouseY / Screen.TwipsPerPixelY - nPrevY) > kMinAnnotSize Then
                    
                    bMoved = Annot.MovePoint(m.Chart, 2, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
                    If bMoved Then
                        If Annot.geMoveFlag = 0 Then StatusMsg "Now click on the third point ...", -1
                        m.nPointCount = 2
                        nPrevX = m.MouseLast.MouseX / Screen.TwipsPerPixelX
                        nPrevY = m.MouseLast.MouseY / Screen.TwipsPerPixelY
                    End If
                Else
                    StatusMsg "Now click on the second point ...", -1
                    bMoved = True       'so flags won't get cleared
                End If
                
            Case 2:
                If Abs(m.MouseLast.MouseX / Screen.TwipsPerPixelX - nPrevX) > kMinAnnotSize Or _
                   Abs(m.MouseLast.MouseY / Screen.TwipsPerPixelY - nPrevY) > kMinAnnotSize Then
                
                    bMoved = Annot.MovePoint(m.Chart, 3, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
                    If bMoved Then
                        If Annot.eType = eANNOT_ChannelHighlight Then
                            StatusMsg "Now click on the fourth point ...", -1
                            m.nPointCount = 3
                            nPrevX = m.MouseLast.MouseX / Screen.TwipsPerPixelX
                            nPrevY = m.MouseLast.MouseY / Screen.TwipsPerPixelY
                        Else
                            StatusMsg
                            Annot.AssignDateTime
                            ClearAnnotFlags False
                            Annot.geMoveFlag = 1    'set flag indicating points selection is complete
                            m.Chart.SyncGlobalAnnots Annot
                            SyncDrawTools
                        End If
                    End If
                Else
                    StatusMsg "Now click on the third point ...", -1
                    bMoved = True       'so flags won't get cleared
                End If
                
            Case 3:
                If Annot.eType = eANNOT_ChannelHighlight Then
                    If Abs(m.MouseLast.MouseX / Screen.TwipsPerPixelX - nPrevX) > kMinAnnotSize Or _
                       Abs(m.MouseLast.MouseY / Screen.TwipsPerPixelY - nPrevY) > kMinAnnotSize Then
                    
                        bMoved = Annot.MovePoint(m.Chart, 4, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
                        If bMoved Then
                            StatusMsg
                            Annot.AssignDateTime
                            ClearAnnotFlags False
                            Annot.geMoveFlag = 1    'set flag indicating points selection is complete
                            m.Chart.SyncGlobalAnnots Annot
                            SyncDrawTools
                        End If
                    Else
                        StatusMsg "Now click on the fourth point ...", -1
                        bMoved = True       'so flags won't get cleared
                    End If
                Else
                    bMoved = False
                    ClearAnnotFlags True, True
                End If
                
            Case Else
                bMoved = False
                ClearAnnotFlags True, True
        End Select
        
        If bMoved = False Then
            StatusMsg
            ClearAnnotFlags True, True
        End If
    ElseIf bMouseUp = False Then
        'moveflag = 1 means user is moving an existing, completed annotation
        If Annot.geMoveFlag = 1 Then
            bMoved = Annot.MovePoint(m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
        ElseIf Annot.eType = eANNOT_TriangleWedge Then
            If m.nPointCount = 2 Then Annot.geMoveFlag = 3
            bMoved = Annot.MovePoint(m.Chart, m.nPointCount + 1, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
        ElseIf Annot.eType = eANNOT_ChannelHighlight Then
            bMoved = Annot.MovePoint(m.Chart, m.nPointCount + 1, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
        End If
        If bMoved Then m.epbCursor = eCursor_Blank
        Annot.geDrawAnn m.Chart
        If Annot.geMoveFlag = 3 Then Annot.geMoveFlag = 0
    End If
            
    Set Annot = Nothing
    HandleAndrewFork = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".HandleAndrewFork", eGDRaiseError_Raise
    
End Function

Private Function HandleDnExpansion(ByVal bMouseUp As Boolean, dY#) As Boolean
On Error GoTo ErrSection:

    Dim i&, nFreeFloat&, nDrawInprog&
    Dim bSnapped As Boolean
    Dim bMoved As Boolean
    
    Dim Annot As cAnnotation
    Dim eType As eAnnotType
       
    HandleDnExpansion = False

    nDrawInprog = Len(g.strActiveDraw)
    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    If Not Annot Is Nothing Then eType = Annot.eType
    
    If UseDiNapFib Then
        If m.nPointCount < 1 And bMouseUp = False Then
            If eType <> eANNOT_FibABCD And eType <> eANNOT_GannacciSwing1 And eType <> eANNOT_GannacciSwing2 Then
                Exit Function
            End If
        End If
    End If
    
    'check if editing (currently this annotation is not moveable after creation)
    If nDrawInprog = 0 Then
        If bMouseUp = True Then
            If nDrawInprog = 0 Then
                If m.nObjectMoving = 1 Then
                    tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
                End If
                ClearAnnotFlags False
                m.nActiveIndIdx = 0
                m.nObjectMoving = 0
                m.Chart.SetCursor
                Exit Function
            End If
        End If
        
        If Annot Is Nothing Then Exit Function
        m.nFocusHiLo = 0
                
    End If
    
    If m.nPointCount > 0 Then
        If Annot Is Nothing Then Exit Function
        If m.MouseLast.nBar < 0 Then Exit Function      'empty bar
    End If
    
'04-26-2007: SnapToHiLoClose has optional PeakCheck boolean that causes 3rd point
'   to snap back one bar when last data bar is chosen. Don't think the PeakCheck
'   is needed for this tool. Commenting out all bMouseUp being passed to skip
'   the PeakCheck. Leave awhile then remove if all okay.

'TLB 8/2/2013: we need the PeakCheck functionality back (e.g. for Elliott Wave)

    If Annot Is Nothing Then
        Set Annot = New cAnnotation
        nFreeFloat = Annot.DefaultProp(Annot.AnnotTypeFromToolID(g.strActiveDraw), "FreeFloat")
    Else
        nFreeFloat = ValOfText(Annot.Prop("FreeFloat"))
    End If
    
    If m.nPointCount = 0 Then       'And (nDrawInprog > 0 Or nFreeFloat = 1) Then
        If nFreeFloat = 1 Then
            If nDrawInprog = 0 Then
                dY = SnapToPrice(m.MouseLast)
                bSnapped = True
            Else
                bSnapped = SnapToHiLoClose(m.MouseLast, dY, 4)
            End If
        ElseIf nDrawInprog = 0 Then
            If Annot.Y(1) > Annot.Y(2) Then
                bSnapped = SnapToHiLoClose(m.MouseLast, dY, 1, True)
            Else
                bSnapped = SnapToHiLoClose(m.MouseLast, dY, 2, True)
            End If
        Else
            bSnapped = SnapToHiLoClose(m.MouseLast, dY, 0, True)  ' bMouseUp) 'point A
        End If
        
        If bSnapped = False Then Exit Function
        
        If m.nFocusHiLo = -1 Or m.nFocusHiLo = 3 Then
            ClearAnnotFlags False
            Exit Function
        End If
        
        If Annot.eType = eANNOT_GannacciSwing1 Or Annot.eType = eANNOT_GannacciSwing2 Then
            bMoved = Annot.MovePoint(m.Chart, 1, m.MouseLast.nPaneID, m.MouseLast.dDate, dY)
            Annot.Prop("HiLo") = m.nFocusHiLo - 1
            If nDrawInprog > 0 Then bMouseUp = True             'in case user drawing with clik-n-drag
        ElseIf nDrawInprog > 0 Then
            bMoved = True
        Else
            bMoved = Annot.MovePoint(m.Chart, 1, m.MouseLast.nPaneID, m.MouseLast.dDate, dY)
            Annot.Prop("HiLo") = m.nFocusHiLo - 1
        End If
    ElseIf m.nPointCount = 1 Then
        If nFreeFloat = 1 Then
            dY = SnapToPrice(m.MouseLast)
            bSnapped = True
        ElseIf nDrawInprog = 0 Then
            If Annot.Y(1) > Annot.Y(2) Then
                bSnapped = SnapToHiLoClose(m.MouseLast, dY, 2, True)
            Else
                bSnapped = SnapToHiLoClose(m.MouseLast, dY, 1, True)
            End If
        Else
            If m.nFocusHiLo = 1 Or m.nFocusHiLo = 4 Then  'focus is bar's high
                bSnapped = SnapToHiLoClose(m.MouseLast, dY, 2, True)  ', bMouseUp)
            Else
                bSnapped = SnapToHiLoClose(m.MouseLast, dY, 1, True) ', bMouseUp)
            End If
        End If
        If bSnapped = True Then
            bMoved = Annot.MovePoint(m.Chart, 2, m.MouseLast.nPaneID, m.MouseLast.dDate, dY)
        End If
    ElseIf m.nPointCount = 2 Then
        If Annot.eType = eANNOT_GannacciSwing1 Then
            bMoved = True
        Else
            If nFreeFloat = 1 Then
                dY = SnapToPrice(m.MouseLast)
                bSnapped = True
            ElseIf nDrawInprog = 0 Then
                If Annot.Y(2) > Annot.YArray(0) Then
                    bSnapped = SnapToHiLoClose(m.MouseLast, dY, 2, True)
                Else
                    bSnapped = SnapToHiLoClose(m.MouseLast, dY, 1, True)
                End If
            Else
                If m.nFocusHiLo = 1 Or m.nFocusHiLo = 4 Then  'focus is bar's high
                    bSnapped = SnapToHiLoClose(m.MouseLast, dY, 1, True)  ', bMouseUp)
                Else
                    bSnapped = SnapToHiLoClose(m.MouseLast, dY, 2, True)  ', bMouseUp)
                End If
            End If
            If bSnapped = True Then
                bMoved = Annot.MovePoint(m.Chart, 3, m.MouseLast.nPaneID, m.MouseLast.dDate, dY)
            End If
        End If
    Else
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Not Annot Is Nothing Then
            dY = SnapToPrice(m.MouseLast)
            bMoved = Annot.MovePoint(m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, m.MouseLast.dDate, dY)
        End If
    End If
        
    Dim bDone As Boolean
        
    If bMoved = True Then
        If bMouseUp = True Then
            If m.nPointCount < 2 Then
                m.nPointCount = m.nPointCount + 1
                If m.nPointCount = 1 Then
                    StatusMsg "Now click on the second peak ...", -1
                ElseIf Annot.eType = eANNOT_GannacciSwing1 Then
                    bDone = True        'drawn with click & drag
                Else
                    StatusMsg "Now click on the third peak ...", -1
                End If
            Else
                bDone = True
            End If
            
            If bDone Then
                StatusMsg
                Annot.AssignDateTime
                ClearAnnotFlags False
                'set flag notifying grapheng.dll that points selection is complete
                Annot.geMoveFlag = 1
                m.Chart.SyncGlobalAnnots Annot
                SyncDrawTools
            End If
        Else
            Annot.geDrawAnn m.Chart
        End If
        
        ' TLB 8/2/2013: we need this to be done in order for the PeakCheck to work!
        m.MouseDown = m.MouseLast
    End If
        
    Set Annot = Nothing
    HandleDnExpansion = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".HandleDnExpansion", eGDRaiseError_Raise
    
End Function

Private Function HandleDnRetrace(ByVal bMouseUp As Boolean, _
        ByVal bDblClick As Boolean, dY#, _
        Optional ByVal bResetTool As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim dHighestHi#, dLowestLow#, dHiLowDate#
    Dim bMoved As Boolean, bSnapped As Boolean, bRC As Boolean
    Dim Annot As cAnnotation
    Dim nIsInBoundary As Long
    
    HandleDnRetrace = False
        
    If bDblClick = True Then
        StatusMsg
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Not Annot Is Nothing Then
            Annot.AssignDateTime
            Annot.geMoveFlag = 1    'sets flag that drawing is complete
        End If
        ClearAnnotFlags False
        'Annot.geDrawAnn m.Chart
        m.Chart.SyncGlobalAnnots Annot
        SyncDrawTools
        Set Annot = Nothing
        Exit Function
    End If
    
    If m.nPointCount < 1 And bMouseUp = False Then Exit Function
           
    'check if editing
    If bMouseUp = True Then
        If Len(g.strActiveDraw) = 0 Then
            tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
            ClearAnnotFlags False
            m.nActiveIndIdx = 0
            m.nObjectMoving = 0
            m.Chart.SetCursor
            Exit Function
        End If
    End If
    
    If m.nPointCount > 0 Then
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Annot Is Nothing Then Exit Function
        If m.MouseLast.nBar < 0 Then Exit Function      'empty bar
    End If
            
    If m.nPointCount = 0 Then
        bSnapped = SnapToHiLoClose(m.MouseLast, dY, 0, True) ' bMouseUp)  'focus point
        
        If bSnapped = False Then Exit Function
        
        If m.nFocusHiLo = -1 Or m.nFocusHiLo = 3 Then
            ClearAnnotFlags False
            Exit Function
        End If
        
        If m.nFocusHiLo = 1 Then
            bRC = m.Chart.HighestHiForwardDate(m.MouseLast.dDate, dHighestHi, dHiLowDate)
            If bRC = True Then
                If dY = dHighestHi Then m.nFocusHiLo = 4  'dynamic
            End If
        ElseIf m.nFocusHiLo = 2 Then
            bRC = m.Chart.LowestLowForwardDate(m.MouseLast.dDate, dLowestLow, dHiLowDate)
            If bRC = True Then
                If dY = dLowestLow Then m.nFocusHiLo = 5
            End If
        End If
        bMoved = True
    Else
        If m.nFocusHiLo = 1 Or m.nFocusHiLo = 4 Then  'focus is bar's high
            bSnapped = SnapToHiLoClose(m.MouseLast, dY, 2, True) 'bMouseUp)
        Else
            bSnapped = SnapToHiLoClose(m.MouseLast, dY, 1, True) 'bMouseUp)
        End If
        If bSnapped = True Then
            bMoved = Annot.MovePoint(m.Chart, m.nPointCount - 1, m.MouseLast.nPaneID, m.MouseLast.dDate, dY, nIsInBoundary)
            If nIsInBoundary = 0 And bMouseUp = True Then
                'bMoved = False
                Beep
                '(don't use InfBox here since triggers Deactivate)
                'MsgBox "This reaction point does not belong to the" & vbCrLf & "selected focus point (please refer to" & vbCrLf & "Chapter 10 of 'Trading With DiNapoli Levels').", _
                '    vbExclamation, "Invalid FibNode"
                StatusMsg "Invalid reaction for selected focus (see Chapter 10 of 'Trading with DiNapoli Levels')", vbRed
            ElseIf nIsInBoundary = 2 And bMouseUp = True And m.nPointCount > 1 Then
                StatusMsg
                Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
                If Not Annot Is Nothing Then
                    Annot.geMoveFlag = 1    'sets flag that drawing is complete
                End If
                ClearAnnotFlags False
                Annot.geDrawAnn m.Chart
                Set Annot = Nothing
                Exit Function
            End If
        End If
    End If

    If bMoved = True Then
        If bMouseUp = True Then
            m.nPointCount = m.nPointCount + 1
            If m.nPointCount <= 1 Then
                StatusMsg "Now click on a reaction point, or double-click to end the series ...", -1
            ElseIf nIsInBoundary = 1 Then
                StatusMsg "Click on the next reaction point, or double-click to end the series ...", -1
            End If
        Else
            Annot.geDrawAnn m.Chart
        End If
    
        ' TLB 8/2/2013: we need this to be done in order for the PeakCheck to work!
        m.MouseDown = m.MouseLast
    End If
    
    Set Annot = Nothing
    HandleDnRetrace = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".HandleDnRetrace", eGDRaiseError_Raise
    
End Function

Private Sub HandleIndClick(Ind As cIndicator, Annot As cAnnotation)
On Error GoTo ErrSection:

    Dim idxPane&, idxInd&, idxTemp&, idx&
    Dim TempInd As cIndicator
    Dim Pane As cPane
    
    If Ind Is Nothing Or m.Chart.Zoomed = True Then
        Exit Sub
    Else
        idxInd = Ind.geIndId
        Ind.SaveGroupInfo
    End If
       
    If m.MouseLast.bOffChart Then
        'scalepaneid holds information on where mouse went offchart
        idxPane = m.MouseLast.nScalePaneId
    Else
        idxPane = m.MouseLast.nPaneID
    End If
    
    frmChartOnOff.ClearPrevChart
    With m.Chart.Tree
        If .Key(idxInd) = "PRICE" And idxPane <> Ind.geIndpaneId Then
            Beep '(can't move price out of price pane)
            MsgBox "Cannot move price indicator out of price pane"
            Annot.geDrawAnn m.Chart
        ElseIf m.MouseLast.bOffChart And idxPane <> -1 Then '-1 indicates mouse went off in x-scale area
            ' hide the indicator that was moved
            Ind.Display = False
            m.Chart.GenerateChart
        ElseIf idxPane <> Ind.geIndpaneId Then
            If Ind.DisplayType = eINDIC_Ribbon Then
                RibbonList Nothing, m.Chart, Ind, 2, 0      '6219
            End If
            ' work with root indicator if linked
            If .NodeLevel(idxInd) > 1 Then
                If Ind.DataType = eINDIC_BooleanArray Then
                    'apply to first indicator in move-to pane
                    idxTemp = .RelativeIndex(idxPane, eTREE_FirstChild)
                    idxPane = idxTemp
                Else
                    idxTemp = .AncestorIndex(idxInd, 1)
                    If .Key(idxTemp) <> "PRICE" Then
                        idxInd = idxTemp
                    Else
                        ' if linked to Price, work with child of Price
                        idxInd = .AncestorIndex(idxInd, 2)
                    End If
                End If
            End If
            If idxPane >= 0 Then
                ' move to this pane
                idxInd = .Move(idxInd, idxPane, eTREE_LastChild)
            Else
                ' add to a newly created pane at bottom:
                ' - first find the last visible pane
                idxPane = -1
                For idxTemp = m.Chart.Tree.Count To 0 Step -1
                    Set Pane = Nothing
                    If m.Chart.Tree.NodeLevel(idxTemp) = 0 Then
                        Set Pane = m.Chart.Tree(idxTemp)
                        If Not Pane Is Nothing Then
                            If Pane.gePaneShow = 1 Then
                                idxPane = idxTemp
                                Exit For
                            End If
                        End If
                    End If
                Next
                ' - then add a new pane after it
                Set Pane = New cPane
                Pane.Display = True
                Pane.Scaling = ePANE_ScaleModeAuto
                idxPane = .Add(Pane, , idxPane, eTREE_NextSibling)
                ' - then move indicator to the new pane
                idxInd = .Move(idxInd, idxPane, eTREE_FirstChild)
                If idxInd > 0 Then
                    ' and make sure it's not overlayed
                    Ind.Overlayed = False
                    If Ind.DataType = eINDIC_BarData Then
                        Pane.DisplayFormat = ePANE_PriceFormat
                    End If
                End If
            End If
            ' clear coded text for moved part of tree (so will get rebuilt)
            idxTemp = .RelativeIndex(idxInd, eTREE_LastDescendant)
            For idx = idxInd To idxTemp
                Set TempInd = .Item(idx)
                TempInd.CodedText = ""
            Next
            Set TempInd = Nothing
            Ind.CheckGroup
            m.Chart.geResetPanes
            m.Chart.GenerateChart
        Else
            'set tag to bring up chart edit dialog
            If idxInd > 0 Then
                tmr.Tag = "EditSettings " & CStr(idxInd)
            Else
                tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
            End If
        End If
    End With

ErrExit:
    Set Pane = Nothing
    Exit Sub
    
ErrSection:
    Set Pane = Nothing
    RaiseError Me.Name & ".HandleIndClick", eGDRaiseError_Raise
    
End Sub

Private Function GetChartCoordinates(ByVal X As Single, ByVal Y As Single, _
        Optional ByVal iShift As Integer = 0) As ChartCoordinates
On Error GoTo ErrSection:

    Static nTickZeroCount As Long

    Dim rc&, dMin#, dMax#
    Dim pixX&, pixY&
    Dim coord As ChartCoordinates
    Dim geCoordInfo As coordinate_info
    Dim Pane As cPane
    
    Dim dPane#, dY#, dTickDiff#         'overflow error fix
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "UNLOADING" Then Exit Function
    If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then Exit Function
    
    ' initialize returned value
    GetChartCoordinates = m.MouseLast   'no changes
    
    'reset
    m.nSplitPaneHittestID = -1
    m.nSplitPaneLabelID = -1
    
    ' get chart coordinates of mouse move
    pixX = X \ Screen.TwipsPerPixelX
    pixY = Y \ Screen.TwipsPerPixelY
    
    geCoordInfo.paneId = -1
    geCoordInfo.x_pixels = pixX
    geCoordInfo.y_pixels = pixY
    
    'grapheng.dll returns
    '   paneId = -1 if a specific pane was specified and point is not in pane
    '   paneId = -2 if no pane was specified and point is above chart area
    '   paneId = -3 if no pane was specified and point is below chart area
    '   paneId = -4 if no pane was specified and point is left of chart area
    '   paneId = -5 if no pane was specified and point is right of chart in y-scale area
    '   paneID = -6 if no pane was specified and point is right of chart in split-pane area
    '   paneID = -7 if scale triangle was hit
    '   paneID = -8 if area where scale triangle would be drawn, but isn't, was hit
    '   paneID = -9 if no pane was specified and point is in y-tick mark area of chart
    rc = geCoordToData(m.Chart.geChartObj, geCoordInfo)
    If rc <> 0 Then
        Exit Function
    End If
    
    If rc = 0 Then
        If geCoordInfo.paneId > 0 Then
            If m.Chart.Tree.NodeLevel(geCoordInfo.paneId) = 0 Then
                coord.nPaneID = geCoordInfo.paneId
                Set Pane = m.Chart.Tree(geCoordInfo.paneId)
                If Not Pane Is Nothing Then
                    dMin = Pane.gePaneMin
                    dMax = Pane.gePaneMax
                End If
            Else
                'This code is hit when user changes templates, move panes etc. and we really
                'don't know where panes are. Original cause of "Type Mismatch" error.
                geCoordInfo.paneId = -1
            End If
        ElseIf geCoordInfo.paneId = -5 Then 'point is in y-scale area
            If m.Chart.Tree.NodeLevel(geCoordInfo.reserved) = 0 Then
                Set Pane = m.Chart.Tree(geCoordInfo.reserved)
                If Not Pane Is Nothing Then
                    coord.nScalePaneId = geCoordInfo.reserved    'this holds paneId when point is off-chart right
                    dMin = Pane.gePaneMin
                    dMax = Pane.gePaneMax
                End If
            Else
                geCoordInfo.paneId = -1 'don't want to process y-scale movement
            End If
            Set Pane = Nothing
        ElseIf geCoordInfo.paneId = -3 Then 'point is in x-scale area
            coord.nScalePaneId = -1
        ElseIf geCoordInfo.paneId = -6 Then 'point is in split-pane area
            coord.nScalePaneId = -6
            m.nSplitPaneHittestID = geCoordInfo.reserved
            m.nSplitPaneLabelID = geCoordInfo.x_value       'array index of label item that was hit else -1
            Set Pane = m.Chart.Tree(m.nSplitPaneHittestID)
            If Not Pane Is Nothing Then
                If Pane.SplitPaneType = ePANE_SplitPaneNone Then
                    m.nSplitPaneHittestID = -1
                    m.nSplitPaneLabelID = -1
                End If
            End If
        ElseIf geCoordInfo.paneId = -7 Then
            coord.nScalePaneId = -7
            pbCursor = eCursor_Hand
        ElseIf geCoordInfo.paneId = -8 Then
            coord.nScalePaneId = -8
            pbCursor = eCursor_Hand
        End If
    End If
        
    If Not Pane Is Nothing Then
        'if difference in tick time is zero two times in a row and difference in Y is 1% of height then switch to manual
        If Pane.Scaling = ePANE_ScaleModeAuto Or Pane.Scaling = ePANE_ScaleModeAutoPrice Then
            dTickDiff = Abs(m.MouseLast.dTickTime - gdTickCount)
            dY = Abs(m.MouseLast.MouseY / Screen.TwipsPerPixelY - pixY)
            dPane = (Pane.geBtmSep - Pane.geTopSep) * 0.02
            If m.epbCursor = eCursor_ChartMove Then
                If nTickZeroCount > 2 Then
                    If dY >= dPane Then
                        Pane.Scaling = ePANE_ScaleModeManual
                        'm.Chart.SyncToolbar
                        frmMain.tbToolbar.Tools("ID_ChartMove").Picture = Picture16(ToolbarIcon("ID_ChartMove"))
                    End If
                    nTickZeroCount = 0
                ElseIf dTickDiff = 0# Then
                    nTickZeroCount = nTickZeroCount + 1
                End If
            Else
                nTickZeroCount = 0          'reset
            End If
        Else
            nTickZeroCount = 0              'reset
        End If
    End If
    
    Set Pane = Nothing
    
    With coord
        .dTickTime = gdTickCount
        .iShift = iShift
        .MouseX = X
        .MouseY = Y
        '.nPaneID = geCoordInfo.paneId
        .dY = geCoordInfo.y_value
        .nScreenX = Int(geCoordInfo.x_value + 0.5)
        If m.Chart.aXBar.Size > m.Chart.geChartPoints Then      'Q & A - is this right fix for data buckets number > actual loaded data (e.g. load only 1 month of data)
            .nX = hsb.Value - m.Chart.geChartPoints + .nScreenX + 1
        Else
            .nX = .nScreenX
        End If
        .nBar = m.Chart.aXBar(.nX) 'could be < 0
        .dDate = m.Chart.aXdate(.nX)
        .dMinY = dMin
        .dMaxY = dMax
        .strRoundedY = RoundedValueStr(.dY, Abs(.dMaxY - .dMinY))
        
        ' check if out of range
        .bOffChart = False
        If geCoordInfo.paneId = -6 Then
            '09-01-2015 grapheng.dll changed to return 1/-1 for split pane labelID:
            '-1: mouse is over an order in split pane area
            ' 1: mouse is in split pane area, but not over an order
            'leave off chart flag = false when mouse is over an order to allow order edit/cancel
            If m.nSplitPaneLabelID > 0 Then .bOffChart = True
        ElseIf geCoordInfo.paneId < 0 Then
            .bOffChart = True
        End If
    End With
    
    GetChartCoordinates = coord

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".GetChartCoordinates", eGDRaiseError_Raise
    
End Function

Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim rc&
        
    If vArgs = "Screen" Then
        frmPrintPreview.CustomPrint = True
        m.bScreenCaptured = True
        ScrCapturePreview
    Else
        m.Chart.HdrPaneId 1
        With frmPrintPreview
            .CustomPrint = True
            .vp.Zoom = 50
            .vp.StartDoc
            rc = gePrintChart(m.Chart.geChartObj, Me.pbChart.hDC, .vp.hDC, 1)   'print preview
            .vp.EndDoc
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".GenerateReport", eGDRaiseError_Raise
    
End Sub

Public Sub PrintReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim rc&, dTop#, dLeft#, dBottom#, dRight#
    Dim prtDevice As Printer
    Dim bFound As Boolean
    Dim strMsg$, strSrc$, strDest$

    If m.Chart.Zoomed And Not m.bScreenCaptured Then
        InfBox "Please unzoom the chart before exporting.", "!", , "Export Chart"
        Exit Sub
    End If
    
    If vArgs = 1 Then   'print to file
        If m.bScreenCaptured Then
            strDest = CommonDialogFile(frmMain.CommonDialog1, True, "Bitmap Files (*.bmp)")
            If Len(strDest) > 0 Then
                strSrc = App.Path & "\ScreenCapture.bmp"
                If FileExist(strSrc) Then
                    If Right(strDest, 4) <> ".bmp" Then strDest = strDest & ".bmp"
                    'fs.CopyFile strSrc, strDest, True
                    FileCopy strSrc, strDest
                End If
            End If
        Else
            frmExportChart.ShowMe Me
            m.Chart.GenerateChart eRedo1_Scrolled
        End If
        Exit Sub
    ElseIf vArgs = 2 Then   'print to clipboard
        If m.bScreenCaptured Then
            strMsg = "You can now paste the image into |another application by selecting |'Edit-Paste'  (or hit 'Ctrl-V')."
            InfBox "i=i ; h=Process Graph Image ; " + strMsg
        Else
            m.Chart.PrintChart 1, False
        End If
        Exit Sub
    End If
    
    If m.bScreenCaptured Then
        With frmPrintPreview.vp
            .Preview = False
            .PrintDoc False, 1, 1
        End With
        Exit Sub
    End If
    
    For Each prtDevice In Printers
        If vArgs = prtDevice.DeviceName Then
            bFound = True
            Exit For
        End If
    Next
    
    If bFound = False Then
        InfBox "Error locating printer:|" & vArgs, "e", , "Printer Error"
        Exit Sub
    End If
    
    Set Printer = prtDevice
    
    Screen.MousePointer = 11
    With frmPrintPreview
        dTop = Val(.txtMargin(0))
        dLeft = Val(.txtMargin(1))
        dBottom = Val(.txtMargin(2))
        dRight = Val(.txtMargin(3))
        
        If .vp.Orientation = orLandscape Then
            Printer.Orientation = vbPRORLandscape
        Else
            Printer.Orientation = vbPRORPortrait
        End If
                
    End With
    
    m.Chart.geSetPrintMargin dTop, dLeft, dBottom, dRight
    Printer.ScaleMode = 3  'Pixels
    Printer.Print
    If Printer.Orientation = vbPRORLandscape Then
        m.Chart.PrintOrientation = 0
    Else
        m.Chart.PrintOrientation = 1
    End If
    
    strSrc = g.strAppPath & "\" & kSextantGif
    If FileExist(strSrc) Then geSextantFile m.Chart.geChartObj, strSrc
    
    rc = gePrintChart(m.Chart.geChartObj, Me.pbChart.hDC, Printer.hDC, 0)
    
    Printer.EndDoc
        
    Screen.MousePointer = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".PrintReport", eGDRaiseError_Raise
    
End Sub

Public Sub EndReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    m.Chart.HdrPaneId 0
    m.Chart.geForceRecalc
    m.Chart.geDrawChart
    m.bScreenCaptured = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".EndReport", eGDRaiseError_Raise
    
End Sub

Public Sub CenterTheDate(ByVal dDate#, Optional ByVal dY# = 0)
On Error Resume Next

    Dim nPos&, iPoint&, rc&
    Dim coordInfo As coordinate_info
    Dim bClear As Boolean
    Dim dDate2#

    dDate2 = RoundToSecond(dDate)
    
    'set scroll bar position for centering chart
    m.Chart.aXdate.BinarySearch gdFixDateTime(dDate2), nPos
    nPos = nPos + m.Chart.geChartPoints \ 2
    If nPos < hsb.Min Then
        nPos = hsb.Min
    ElseIf nPos > hsb.Max Then
        nPos = hsb.Max
    End If
    If hsb.Value = nPos Then bClear = True
    hsb.Value = nPos    'note: this triggers a generatechart re-draw
    
    'set point number & yvalue for passing to graphics engine to get x/y pixels
    coordInfo.paneId = m.Chart.Tree.Index("PRICE PANE")
    m.Chart.aXdate.BinarySearch dDate2, iPoint
    iPoint = iPoint - m.Chart.ScreenStartX
    coordInfo.x_value = iPoint
    If dY <> 0 Then
        coordInfo.y_value = dY
    Else
        gdBinarySearch m.Chart.Bars.ArrayHandle(eBARS_DateTime), dDate2, nPos&, eGdSort_Default, 0, 999999999
        coordInfo.y_value = m.Chart.Bars(eBARS_High, nPos)
    End If
    
    rc = geDataToCoord(m.Chart.geChartObj, coordInfo)
    If rc = 0 And coordInfo.paneId > 0 Then
        If bClear = True Then pbChart.Refresh
        pbChart.AutoRedraw = False
        'Developer Note: can pass 0=circle or 1=square (default is circle) for shape parameter
        rc = geDrawMarker(pbChart.hDC, coordInfo.x_pixels, coordInfo.y_pixels, 6, 2, vbGreen)
        pbChart.AutoRedraw = True
    End If

End Sub

Public Property Get MouseLastDate() As Double
On Error GoTo ErrSection:

    MouseLastDate = m.MouseLast.dDate

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".Property.MouseLastDate", eGDRaiseError_Raise
    
End Property

Private Sub CheckMinMoveOnCreate(Annot As cAnnotation)
On Error GoTo ErrSection:

    Dim nPixDiffX&, nPixDiffY&
    Dim iDir1%, iDir2%, iDir3%, iDir4%, i&
    
    If Annot.eType = eANNOT_Rectangle Then
        If Annot.eUsage <> eANNOT_PatternProfit And Annot.eUsage <> eANNOT_UserAdded Then Exit Sub
    ElseIf Annot.eUsage <> eANNOT_UserAdded Then
        Exit Sub
    End If
        
    If m.nPointCount >= 2 And m.nActiveAnnotIdx > 0 And m.bAnnotCreated Then
        nPixDiffX = kMinAnnotSize
        nPixDiffY = kMinAnnotSize
    Else
        nPixDiffX = Abs(m.MouseLast.MouseX / Screen.TwipsPerPixelX - m.MouseDown.MouseX / Screen.TwipsPerPixelX)
        nPixDiffY = Abs(m.MouseLast.MouseY / Screen.TwipsPerPixelY - m.MouseDown.MouseY / Screen.TwipsPerPixelY)
    End If
    
    'Note: these tools do not require a minimum mousemove
    'eANNOT_HorzLine, eANNOT_VertLine, eANNOT_Icon, eANNOT_WaveLabels
    'eANNOT_DNExpansion, eANNOT_DNRetracement, eANNOT_FibTimeZones
    'eANNOT_AndrewFork , eANNOT_TriangleWedge, eANNOT_ChannelHighlight
    'eANNOT_ElliotLabel, eANNOT_ElliotTimeRatio, eANNOT_DanCodeZone
    'eANNOT_BalloonStrangle
    
    Select Case Annot.eType
        Case eANNOT_Trendline, eANNOT_Trendline2, eANNOT_Trendline3, eANNOT_Trendline4, _
             eANNOT_TrendChannel, eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, _
             eANNOT_DollarLine4, eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
             eANNOT_Fibonacci4, eANNOT_FibTimeRatio, eANNOT_FibArcs, eANNOT_TargetShooter, _
             eANNOT_Rectangle, eANNOT_Ellipse, eANNOT_TimeCycle, eANNOT_FibExpansion, eANNOT_AdvRiskReward, _
             eANNOT_FibFan, eANNOT_SpResistFan, eANNOT_Mirror, eANNOT_Pattern, eANNOT_RiskReward, _
             eANNOT_Bracket, eANNOT_Pivot, eANNOT_DanCodeFib, eANNOT_GannacciSwingSquare       '4762
             
            
            If nPixDiffX < kMinAnnotSize And nPixDiffY < kMinAnnotSize Then
                If Annot.eType = eANNOT_DollarLine Or Annot.eType = eANNOT_DollarLine2 Or Annot.eType = eANNOT_DollarLine3 Or _
                   Annot.eType = eANNOT_DollarLine4 And Annot.Prop("KeepAtEnd") = 1 Then
                    Annot.Prop("KeepAtEnd") = 1
                    Annot.geMoveFlag = 1
                Else
                    Annot.geRemoveAnnotation m.Chart.geChartObj
                    m.Chart.Annots.Remove Annot.geAnnId
                    m.bAnnotCreated = False
                End If
            ElseIf Annot.eType = eANNOT_DollarLine Or Annot.eType = eANNOT_DollarLine2 Or _
                   Annot.eType = eANNOT_DollarLine3 Or Annot.eType = eANNOT_DollarLine4 Or _
                   Annot.eType = eANNOT_GannacciSwingSquare Or _
                   Annot.eType = eANNOT_RiskReward Then
                Annot.geMoveFlag = 1
            ElseIf Annot.eType = eANNOT_Pattern Then
                Annot.geMoveFlag = 1
                m.bNewPattern = True
            End If
        Case eANNOT_RegressionLine
            If nPixDiffX < kMinAnnotSize And nPixDiffY < kMinAnnotSize Then
                Annot.geRemoveAnnotation m.Chart.geChartObj
                m.Chart.Annots.Remove Annot.geAnnId
                m.bAnnotCreated = False
            Else
                'set move flag for regression line so grapheng.dll will not draw vertical lines
                Annot.geMoveFlag = 1
                If m.MouseLast.nX >= m.Chart.LastGoodDataBar(True, False) _
                    Or m.MouseDown.nX >= m.Chart.LastGoodDataBar(True, False) Then
                    Annot.Prop("KeepAtEnd") = 1
                ElseIf m.MouseLast.nBar < 0 Then        'empty bars
                    Annot.FixRegressionEndPoint m.Chart
                End If
            End If
        Case eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4
            Annot.AssignDateTime
            Annot.geMoveFlag = 1    'set this so mousemove will not auto add arrow
            If nPixDiffX < kMinAnnotSize And nPixDiffY < kMinAnnotSize Then
                Annot.dDate(1) = m.MouseDown.dDate
                Annot.dDate(2) = m.MouseDown.dDate
                Annot.Y(1) = m.MouseDown.dY
                Annot.Y(2) = m.MouseDown.dY
                'text edit was created with a single click, turn arrow off
                Annot.Prop("ArrowStyle") = 0
            Else
                m.bAnnotCreated = False
            End If
        Case eANNOT_GannLines
            Annot.geMoveFlag = 0    'set this so annotation class knows initial draw is complete
            If nPixDiffX < kMinAnnotSize And nPixDiffY < kMinAnnotSize Then
                Annot.Y(2) = Annot.Y(1)     'single click
            Else
                'clear default directions (DirX = 2 means the anchor line is in that quadrant)
                iDir1 = Val(Annot.Prop("DirNE"))
                iDir2 = Val(Annot.Prop("DirSE"))
                iDir3 = Val(Annot.Prop("DirNW"))
                iDir4 = Val(Annot.Prop("DirSW"))
                If iDir1 = 2 Or iDir2 = 2 Or iDir3 = 2 Or iDir4 = 2 Then
                    If Val(Annot.Prop("DirNE")) = 1 Then Annot.Prop("DirNE") = 0
                    If Val(Annot.Prop("DirSE")) = 1 Then Annot.Prop("DirSE") = 0
                    If Val(Annot.Prop("DirNW")) = 1 Then Annot.Prop("DirNW") = 0
                    If Val(Annot.Prop("DirSW")) = 1 Then Annot.Prop("DirSW") = 0
                End If
            End If
        Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
            If nPixDiffX < kMinAnnotSize And nPixDiffY < kMinAnnotSize Then
                Annot.dDate(2) = m.Chart.Bars(eBARS_DateTime, m.Chart.aXBar(m.Chart.ScreenEndX))     'single click, extend to right
                Annot.Y(2) = Annot.Y(1)
                Annot.Prop("Ext") = "1"
            End If
    End Select
    
    If m.bAnnotCreated Then
        Annot.AssignDateTime
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".CheckMinMoveOnCreate", eGDRaiseError_Raise
    
End Sub

Public Sub GlobalCursorSync(Optional ByVal bForceSync As Boolean = False)
On Error GoTo ErrSection:

    Static dPrevDate#, nPrevMouseY&
    
    Dim nHorz&, nVert&, nDrawFatBar&, nSplitPane&, i&, j&
    Dim frm As Form
    Dim hXDates As Long
    Dim dDate#, dTipDate#     ', dDateFound#
    Dim strTip$
    
    Dim Pane As cPane

    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "UNLOADING" Then Exit Sub
    If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then Exit Sub
    
    Set Pane = m.Chart.Tree("PRICE PANE")
    If Not Pane Is Nothing Then
        If Pane.SplitPaneType = ePANE_SplitPaneOptGraph Then
            If m.MouseLast.nScalePaneId = -6 Then
                nSplitPane = 1
            End If
        End If
    End If
    
    If m.MouseLast.bOffChart And nSplitPane = 0 Then
        Exit Sub
    End If

    'don't want to sync/re-draw the cross hair as it makes the new icons hard to see
    If FormIsLoaded("frmIconAnnot") Then
        If frmIconAnnot.Visible Then Exit Sub
    ElseIf FormIsLoaded("frmElliot") Then
        If frmElliot.Visible Then
            If g.ChartGlobals.eChartMode <> eMode_Move Then Exit Sub
        End If
    End If

    If m.nObjectMoving > 0 Then Exit Sub
    
    If m.eCrossHairOn = eCursor_CrossHair Or _
       m.eCrossHairOn = eCursor_Horizontal Then
        nHorz = 1
    End If
    If m.eCrossHairOn = eCursor_CrossHair Or _
       m.eCrossHairOn = eCursor_Vertical Then
        nVert = 1
    End If
                            
    'this check helps cuts down on flashing horz crosshair (aardvark 930)
    'it is not essential so can be removed if necessary
    'If nVert <> 1 And Abs(m.MouseLast.MouseY - nPrevMouseY) < kMinMoveCount Then Exit Sub
    
    dDate = m.MouseLast.dDate
    If bForceSync Then
        'currently only the chart object calls this with force sync set to true
        m.MouseLast = GetChartCoordinates(m.MouseLast.MouseX, m.MouseLast.MouseY)
        If dPrevDate = dDate And nPrevMouseY = m.MouseLast.MouseY And nSplitPane = 0 Then nDrawFatBar = 1  '4260 (ghostbar issue)
    ElseIf nSplitPane = 0 Then
        nDrawFatBar = 1     'always draw fat bar on mouse move unless mouse is in splitpane area
    End If
    
    nPrevMouseY = m.MouseLast.MouseY
    
    ' always need to do this chart
    If nVert = 1 Or nHorz = 1 Then
        pbChart.AutoRedraw = False
        j = geSyncCrossHairEx(Chart.geChartObj, pbChart.hWnd, pbChart.hDC, _
                dDate, m.MouseLast.MouseX / Screen.TwipsPerPixelX, _
                m.MouseLast.MouseY / Screen.TwipsPerPixelY, nVert, nHorz, nDrawFatBar, nSplitPane)
        pbChart.AutoRedraw = True
    End If
                           
    ' only need to resync other charts if date has changed and
    ' vert cursor is being used and this chart is not maximized
    If Not bForceSync And nSplitPane = 0 Then
        'JM: Original code, leave awhile then remove if all okay - 03-06-2009
        'If Me.WindowState = vbMaximized Or dDate = dPrevDate Then Exit Sub
        If dDate = dPrevDate Then Exit Sub
    End If
    dPrevDate = dDate
    
    ' resync other charts if not in splitpane area
    If nSplitPane = 1 Then
        strTip = ""
        ' refresh date/price tips
        vseTipX.Tag = ""
        vseTipY.Tag = m.Chart.Bars.PriceDisplay(m.MouseLast.dY)
        RefreshTips m.MouseLast.MouseX, m.MouseLast.MouseY
        'show values in chart tips
        If g.ChartGlobals.eChartMode = eMode_ChartOrder Then        '5645
            ChartTips eTiptype_None
        Else
            ChartTips eTipType_OptGraph
        End If
    Else
        For i = 0 To Forms.Count - 1
            If IsFrmChart(Forms(i)) Then
                Set frm = Forms(i)
                dDate = m.MouseLast.dDate
                If frm.hWnd <> Me.hWnd Then
                    If dDate > 0 Then
                        If frm.Chart.Bars.IsIntraday Then
                            If Not m.Chart.Bars.IsIntraday Then
                                dDate = dDate + Me.Chart.Bars.Prop(eBARS_EndTime) / 1440#
                            End If
                            dDate = ConvertTimeZone(dDate, Me.Chart.Bars.Prop(eBARS_ExchangeTimeZoneInf), frm.Chart.Bars.Prop(eBARS_ExchangeTimeZoneInf))
                        ElseIf m.Chart.Bars.IsIntraday Then
                            'adjust date for when chart cursor is in has intra-day data but other chart(s) are non-intraday
                            If Hour(dDate) * 60 + Minute(dDate) > Me.Chart.Bars.Prop(eBARS_CrossoverTime) Then
                                dDate = Int(dDate) + 1
                            Else
                                dDate = Int(dDate)
                            End If
                        End If
                    End If
                    
                    If nVert = 1 Then
                        'call graphics engine to draw crosshair
                        frm.pbChart.AutoRedraw = False
                        j = geSyncCrossHairEx(frm.Chart.geChartObj, frm.pbChart.hWnd, frm.pbChart.hDC, _
                            dDate, -1, m.MouseLast.dY, 1, 0, 1)
                        frm.pbChart.AutoRedraw = True
                    End If
                End If
                    
                ' refresh date/price tips
                strTip = ""
                If dDate > 0 Then
                    If frm.hWnd = Me.hWnd Then
                        dTipDate = DateTimeConvert(m.Chart.Bars, dDate)
                        If m.Chart.Bars.IsIntraday Then
                            strTip = DateFormat(dTipDate, MM_DD_YY) & Format(dTipDate, " Hh:Nn")
                        Else
                            strTip = DateFormat(dTipDate, MM_DD_YYYY)
                        End If
                    End If
                    
                    hXDates = frm.Chart.aXdate.ArrayHandle
                    gdBinarySearch hXDates, gdFixDateTime(dDate), j, eGdSort_Default, 0, 999999999
                    
                    frm.vseMouse.Tag = Str(-99) & Chr(9) & Str(j)
                    frm.DoMouseLabel True
                End If
                
                ' only show a tip for this chart
                If Len(strTip) > 0 Then
                    frm.vseTipX.Tag = strTip
                    RefreshTips m.MouseLast.MouseX, m.MouseLast.MouseY
                Else
                    ' hide the tips for other charts
                    With frm.vseTipX
                        .Caption = ""
                        .Top = -1000 - .Height
                    End With
                    With frm.vseTipY
                        .Caption = ""
                        .Top = -1000 - .Height
                    End With
                End If
            End If
        Next
    End If
    
ErrExit:
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".GlobalCursorSync", eGDRaiseError_Raise
    
End Sub

Private Sub HandleChartMove(ByVal X#, ByVal Y#, ByVal bMouseUp As Boolean, _
    Optional ByVal eCurrentCursor As enumCursor = eCursor_ChartMove)

On Error GoTo ErrSection:

    Static eCursorSave As enumCursor
    Static dLastMoveTime As Double
    
    Dim Pane As cPane
    Dim nGridModeSave&, nDiff&
    Dim nX As Single, nY As Single
    Dim bLocked As Boolean
                                                
    If eCurrentCursor <> eCursor_ChartMove Then
        eCursorSave = eCurrentCursor    'aardvark 1442 fix
        Exit Sub
    End If
    
    geAnnotMove m.Chart.geChartObj, 1
    'note: if user double clicks on an indicator we do not want to move the chart
    If m.bChartMoveInProg = False Then
        m.bChartMoveInProg = True
        Exit Sub
    End If
                   
    Set Pane = m.Chart.Tree(m.MouseLast.nPaneID)
    
    If Pane Is Nothing Then Exit Sub
                    
    If bMouseUp = True Then
        'check to see if extra forecast bars need to be added
        geAnnotMove m.Chart.geChartObj, 0
        If hsb.Value + m.nScrollChange > hsb.Max Then
            bLocked = LockWindowUpdate(Me.hWnd)
            
            If m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                m.Chart.ForecastBars(Me) = m.Chart.ForecastBars + hsb.Value + m.nScrollChange - hsb.Max
                m.Chart.GenerateChart eRedo7_ReloadRT
            Else
                hsb.Value = hsb.Max
                m.Chart.GenerateChart eRedo1_Scrolled
            End If
            

'05-01-2008:
'This check was a work-around. The real fix is in BuildBars (available 4/30/08 and later builds)
'Leave awhile for reference then remove if all okay.
'            If m.Chart.Bars.Prop(eBARS_PeriodType) = ePRD_Days And m.Chart.Bars.Prop(eBARS_PeriodsPerBar) > 1 Then
'                m.Chart.GenerateChart eRedo9_ReloadData     '4400
'            Else
'                m.Chart.GenerateChart eRedo7_ReloadRT
'            End If
            
            
            If bLocked Then LockWindowUpdate 0
            hsb.Value = hsb.Max
        Else
            hsb.Value = hsb.Value + m.nScrollChange
            hsb_Change
        End If
        m.bChartMoveInProg = False
        m.epbCursor = eCursorSave
        ShowCursor
        GlobalCursorSync
        m.Chart.SyncToolbar
        m.Chart.BidAskOnChart
        Exit Sub
    ElseIf m.nScaleStartPixel = 0 Then
        m.nScaleStartPixel = X / Screen.TwipsPerPixelX
        dLastMoveTime = 0
        Exit Sub
    ElseIf gdTickCount - dLastMoveTime < 30 Then
        Exit Sub
    End If
    
    'temporarily turn off grid & labels in y-scale area and bid/ask on chart
    nGridModeSave = m.Chart.geGridMode
    m.Chart.geGridMode = 6
    m.Chart.BidAskOnChart True
    
    'Note: nDiff is pos when moving left & neg when moving right
    nDiff = m.nScaleStartPixel - X / Screen.TwipsPerPixelX
    
    If nDiff < 0 And m.Chart.ScreenStartX <= 0 Then
        'we are at the beginning of available data for screen
        'adjust y-scale only if in price pane and scale mode is manual
        If Pane.PricePaneFlag = 1 And Pane.Scaling = ePANE_ScaleModeManual Then
            Pane.gePaneMax = Pane.gePaneMax + (m.MouseDown.dY - m.MouseLast.dY)
            Pane.gePaneMin = Pane.gePaneMin + (m.MouseDown.dY - m.MouseLast.dY)
            Pane.Max = Pane.gePaneMax
            Pane.Min = Pane.gePaneMin
        End If
        
        m.bChartMoveInProg = False
        
        GoTo ErrExit
    End If
    
    If Abs(nDiff) >= m.Chart.PixelsPerBar Then
        m.nScaleStartPixel = X / Screen.TwipsPerPixelX
        'fake shift in x-direction
        nDiff = nDiff / m.Chart.PixelsPerBar
        m.nScrollChange = m.nScrollChange + nDiff
        m.Chart.FakeForecastBarsChange nDiff * -1
    End If
    
    'adjust y-scale only if in price pane and scale mode is manual
    If Pane.PricePaneFlag = 1 And Pane.Scaling = ePANE_ScaleModeManual Then
        Pane.gePaneMax = Pane.gePaneMax + (m.MouseDown.dY - m.MouseLast.dY)
        Pane.gePaneMin = Pane.gePaneMin + (m.MouseDown.dY - m.MouseLast.dY)
        Pane.Max = Pane.gePaneMax
        Pane.Min = Pane.gePaneMin
    End If
    
ErrExit:
    
'JM 09-17-2007: Cannot remember why this IF check is here, but commenting it out fixes aardvark 4290
'    If nDiff < 0 And m.Chart.Bars.Prop(eBARS_PeriodType) = ePRD_Minutes And m.Chart.Bars.Prop(eBARS_PeriodsPerBar) = 1 Then
'        'do nothing
'    Else
        pbChart.AutoRedraw = False
        geDrawWindow m.Chart.geChartObj, pbChart.hWnd, pbChart.hDC
        pbChart.AutoRedraw = True
'    End If
    
    m.Chart.geGridMode = nGridModeSave
    
    dLastMoveTime = gdTickCount
               
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".HandleChartMove", eGDRaiseError_Raise
    
End Sub

Private Function EditMoveAnnot(Ind As cIndicator, Annot As cAnnotation, _
    ByVal X As Single, ByVal Y As Single) As Boolean
On Error GoTo ErrSection:

    Dim strKey$, strText$
    Dim dY#, dDate#, idx&, rc&
    Dim bExit As Boolean
    
    If Annot Is Nothing Then
        EditMoveAnnot = False
        Exit Function  'precautionary, should never happen
    End If
                
    strKey = m.Chart.Annots.Key(m.nActiveAnnotIdx)
    idx = Val(strKey)
    
    If strKey = "system name" Then
        tmr.Tag = "EditSettings -1"
    ElseIf Annot.eUsage = eANNOT_Notation Then
        m.Chart.ShowNotationMsg m.Chart.Annots.Key(m.nActiveAnnotIdx), True
    ElseIf Annot.eUsage = eANNOT_IndicatorLabel And InStr(Annot.Text, "Price Clusters") <> 0 Then
        tmr.Tag = "Edit Price Cluster"
    ElseIf Annot.eUsage = eANNOT_IndicatorLabel And InStr(Annot.Text, "Time Clusters") <> 0 Then
        tmr.Tag = "Edit Time Cluster"
    ElseIf Annot.eUsage = eANNOT_Trades Then
        idx = Val(Parse(m.Chart.Annots.Key(m.nActiveAnnotIdx), " ", 2))
        If m.bGameMode And tmrGameMode.Enabled = False Then
            strText = UCase(m.Chart.ShowTradeMsg(idx, True))
            If InStr(strText, "BUY") Or InStr(strText, "SELL") Then
                HandleGameOrder eGDTTEditOrderMode_GameEditOrder, idx, Annot
            End If
        Else
            m.Chart.ShowTradeMsg idx, True
        End If
    ElseIf Annot.eUsage = eANNOT_PriceAlert And m.nObjectMoving > kMinMoveCount Then
        If Not Annot.AlertObject Is Nothing Then
            Annot.AlertObject.UpdateChartObject False
        End If
    ElseIf Annot.eType = eANNOT_DNRetracement Then
        HandleDnRetrace True, False, dY
        bExit = True
    ElseIf Annot.eType = eANNOT_DNExpansion Or Annot.eType = eANNOT_FibABCD Or _
        Annot.eType = eANNOT_DNExpansion2 Or Annot.eType = eANNOT_DNExpansion3 Or Annot.eType = eANNOT_DNExpansion4 Or _
        Annot.eType = eANNOT_GannacciSwing1 Or Annot.eType = eANNOT_GannacciSwing2 Then
        HandleDnExpansion True, dY
        bExit = True
    ElseIf Annot.eType = eANNOT_Gartley Then
        HandleGartley True, dY
        bExit = True
    ElseIf Annot.eType = eANNOT_AndrewFork Or Annot.eType = eANNOT_ElliotTimeRatio Then
        HandleAndrewFork True
        bExit = True
    ElseIf Annot.eType = eANNOT_ChannelHighlight Or Annot.eType = eANNOT_TriangleWedge Then
        HandleTriChannel True, False
        bExit = True
    ElseIf Annot.eType = eANNOT_WaveLabels Then
        HandleWaveLabels True
        bExit = True
    ElseIf Annot.eType = eANNOT_BellAlert Then
        Annot.AlertAddEditInprog = True
        frmAlertsSetup.ShowMe m.Chart.Symbol        '5572
        Annot.AlertAddEditInprog = False
    ElseIf Annot.eType = eANNOT_SimpleLine And Annot.eUsage = eANNOT_WhatIf Then
        HandleWhatIf Annot, True
    ElseIf Annot.eType = eANNOT_TrendChannel And g.strActiveDraw = "ID_TrendChannel" Then   '4626
        If m.nPointCount = 2 Then
            Annot.Prop("ChannelType") = 0
            Annot.AddChannelOnMove m.MouseDown.dDate, m.MouseDown.dY        '6716
            StatusMsg "Now click the third point..."
            m.nPointCount = 3
            bExit = True
        ElseIf m.nPointCount = 3 Then
            Annot.geMoveFlag = 1        '6717
            Annot.AddChannelOnMove m.MouseDown.dDate, m.MouseDown.dY
            Annot.AssignDateTime
        End If
    ElseIf m.nObjectMoving > kMinMoveCount Then
        If Annot.eUsage = eANNOT_IndicatorLabel Then
            HandleIndClick Ind, Annot
'JM 05-21-2014: commented out for issue 6985
'        ElseIf m.MouseLast.bOffChart Then
'            'delete the annot if point being moved is off the chart
'            If m.nActiveAnnotPt = 0 Then
'                If Annot.eType <> eANNOT_Mirror Then        '4529 - issue 2
'                    m.Chart.Annots.Remove m.nActiveAnnotIdx
'                End If
'                rc = m.Chart.geDrawChart
'            End If
        ElseIf Annot.ShiftKey = 2 And _
            (Annot.eType = eANNOT_TrendChannel Or _
                Annot.eType = eANNOT_Trendline Or Annot.eType = eANNOT_Trendline2 Or _
                Annot.eType = eANNOT_Trendline3 Or Annot.eType = eANNOT_Trendline4) Then

            Annot.AssignDateTime
            Annot.UpdateAlert 2, True
            rc = Annot.geDrawAnn(m.Chart)
        
        Else
            ' move the annot
            dY = SnapToPrice(m.MouseLast, Annot.eType)
            If Annot.eType = eANNOT_Ellipse And m.nActiveAnnotPt = 4 Then
                Annot.MovePoint m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, X, Y
            Else
                Annot.MovePoint m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, m.MouseLast.nX, dY
            End If
            If Annot.eType = eANNOT_RegressionLine Then
                Annot.FixRegressionEndPoint m.Chart
            ElseIf Annot.eType = eANNOT_GannLines Then
                Annot.geMoveFlag = 0
            ElseIf Annot.eType = eANNOT_DollarLine Or Annot.eType = eANNOT_DollarLine2 Or _
                   Annot.eType = eANNOT_DollarLine3 Or Annot.eType = eANNOT_DollarLine4 Or _
                   Annot.eType = eANNOT_GannacciSwingSquare Or _
                   Annot.eType = eANNOT_Pattern Or Annot.eType = eANNOT_RiskReward Then
                'for $Line: grapheng.dll will auto-position text - aardvark 1366
                'for pattern: mousemove will make copy
                Annot.geMoveFlag = 1
            End If
            Annot.AssignDateTime
            Annot.UpdateAlert 2, True
            rc = Annot.geDrawAnn(m.Chart)
        End If
    Else
        If m.bAnnotCreated = True Then
            CheckMinMoveOnCreate Annot
            m.Chart.GenerateChart eRedo1_Scrolled
            If Annot.eType = eANNOT_TextEdit Or Annot.eType = eANNOT_TextEdit2 Or _
               Annot.eType = eANNOT_TextEdit3 Or Annot.eType = eANNOT_TextEdit4 Then
                m.nActiveAnnotIdx = Annot.geAnnId
                tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
            End If
            m.nPointCount = 0
            m.Chart.LastEditCreate Annot, True
        ElseIf m.bNewPatternMoving Then
            m.bNewPatternMoving = False
        Else
            ' edit the annot or indicator if the annot is an indicator label
            gePeekClickMsg pbChart.hWnd     'aardvark 1030 fix
            If idx > 0 Then
                tmr.Tag = "EditSettings " & CStr(idx)
            Else
                tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
            End If
        End If
    End If
    
    Select Case Annot.eType
        Case eANNOT_Rectangle
            If Annot.eUsage = eANNOT_PatternProfit Then
                PfpReset ePfpReset_GridPfp
                If Not m.oPatternProfit Is Nothing Then
                    If m.bAnnotCreated Then
                        m.oPatternProfit.PatternAnnotCheck m.Chart, Me, False, Annot
                    Else
                        m.oPatternProfit.ClearPFP
                    End If
                    PfpReset ePfpReset_PpfAnnotInd
                    
                    rc = Annot.PatternLength(m.Chart)
                    If rc = 0 Then
                        PfpReset ePfpReset_ClearAll
                    Else
                        Annot.geDrawAnn m.Chart
                        m.oPatternProfit.PatternDateFrom = Annot.dDate(1)
                        m.oPatternProfit.PatternDateTo = Annot.dDate(2)
                        m.oPatternProfit.PatternSelLength = rc
                        cmdMatchesPFP.Enabled = True
                    End If
                End If
            ElseIf Annot.eUsage = eANNOT_FibClusters Then
                If Not m.Chart Is Nothing Then
                    rc = Annot.PatternLength(Chart)
                    If rc > 0 Then
                        dDate = Annot.dDate(2)
                        If m.Chart.InitClusterInd(False, dDate, rc) > 0 Then m.Chart.RedoMode = eRedo5_RecalcInd
                        m.Chart.GenerateChart
                    Else
                        m.Chart.RemoveAnnots True, eANNOT_Rectangle, eANNOT_FibClusters
                    End If
                End If
            End If
    End Select
    
    EditMoveAnnot = bExit

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".EditMoveAnnot", eGDRaiseError_Raise
End Function

Private Sub EraseAnnot(Ind As cIndicator, Annot As cAnnotation, Button As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim bExit As Boolean
    Dim bRemoveAnnot As Boolean
    Dim bMultiChart As Boolean

    
    If Annot Is Nothing Or Annot.eType = eANNOT_UndefinedType Or _
        Annot.eType = eANNOT_IndicLabel Or Annot.eType = eANNOT_SimpleLine Then
        Exit Sub      'precautionary, should never happen
    End If
        
    bMultiChart = Annot.MultiChartFlag          '4116
    
    Select Case Annot.eUsage
        Case eANNOT_IndicatorLabel
            HandleAnnotDblclk
            bExit = True
        Case eANNOT_Trades
            m.Chart.ShowTrades = False
            bExit = True
        Case eANNOT_Notation
            g.ChartGlobals.bSplitsRolls = False
            bExit = True
    End Select
    
    'alternate way for hiding individual fib lines
    If Button = vbLeftButton And Shift = 2 Then
        If Annot.IsFibType Then Annot.HideFibLine Annot.HitItemIndex
        m.Chart.SyncGlobalAnnots Nothing, bMultiChart
        bExit = True    'always exit if ctrl-key is held down in eraser mode
    End If
    
    If bExit Then Exit Sub
    
    Select Case Annot.eType
        Case eANNOT_GannLines
            If Annot.HitItemIndex = 0 Then
                bRemoveAnnot = True
            Else
                Annot.HideGannLine Annot.HitItemIndex
            End If
        Case eANNOT_TargetShooter
            If Annot.HitItemIndex < 6 Then
                bRemoveAnnot = True
            Else
                Annot.HideTargetShooterLine Annot.HitItemIndex
            End If
        Case Else
            bRemoveAnnot = True
    End Select
    
    If bRemoveAnnot Then
        If Annot.eUsage = eANNOT_PriceAlert Then
            Dim Alert As cAlert
            Set Alert = Annot.AlertObject
            If Not Alert Is Nothing Then
                Alert.UpdateChartObject True
                g.Alerts.Remove Alert.AlertKey
                If FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid
            End If
        Else
            Annot.geRemoveAnnotation m.Chart.geChartObj
            m.Chart.Annots.Remove Annot.geAnnId
        End If
    End If
    
    m.Chart.SyncGlobalAnnots Nothing, bMultiChart
    Set Annot = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".EraseAnnot", eGDRaiseError_Raise
End Sub

Private Sub ChartTips(ByVal eType As enumTipType, _
    Optional ByVal nItem& = 0, Optional ByVal nSubItem& = 0)
On Error Resume Next

    Static strPrevTip$
    
    Dim strTip$, strFormat$, nWidth&, nHeight&, bToLeft As Boolean
    Dim dNow#, dThird#, dTwoThird#, dExpire#, dPrice#, dTemp#
    Dim dDateNow#, dDateThird#, dDateTwoThird#, dDateExpire#
    
    Dim Annot As cAnnotation
    Dim Pane As cPane
    
    If m.bChartMoveInProg Then Exit Sub

    ' get tip to display
    If eType <> eTiptype_None And (g.ChartGlobals.bChartTips Or eType = eTipType_OptGraph) Then
        Select Case eType
            Case eTipType_Trade
                strTip = m.Chart.ShowTradeMsg(nItem, False)
                With lblTipChart
                    .Alignment = 2       'center
                    .Move 0, 0, Me.TextWidth(strTip) + 120, Me.TextHeight(strTip) + 50
                End With
            Case eTipType_Annot
                If nSubItem >= 0 Then
                    Set Annot = m.Chart.Annots(nItem)
                    If Not Annot Is Nothing Then
                        If Annot.eUsage = eANNOT_Notation Then
                            strTip = m.Chart.Annots.Key(nItem)
                            strTip = m.Chart.ShowNotationMsg(strTip, False)
                            If Me.TextWidth(strTip) + m.MouseLast.MouseX > pbChart.ScaleWidth Then strTip = Chr(27) & strTip    '5353
                        ElseIf Annot.eUsage = eANNOT_IndicatorLabel Then
                            strTip = IndFunctionTip(Annot)
                        Else
                            strTip = Annot.ChartTip(m.Chart, nSubItem, m.MouseLast.dDate)
                        End If
                        If Annot.eUsage <> eANNOT_IndicatorLabel Then
                            With lblTipChart
                                .Alignment = vbRightJustify
                                .Move 0, 0, Me.TextWidth(strTip) + 120, Me.TextHeight(strTip) + 50
                            End With
                        End If
                    End If
                End If
            Case eTipType_OptGraph
                dNow = kNullData
                dThird = kNullData
                dTwoThird = kNullData
                dExpire = kNullData
                
                strTip = cboRiskGraphType.Text
                Select Case strTip
                    Case "Profit/Loss"
                        strFormat = "#,###,##0.00"
                    Case "Delta", "Theta"
                        strFormat = "##0.0000"
                    Case "Gamma", "Vega"
                        strFormat = "##0.000000"
                    Case Else
                        strFormat = "#####0.000000"
                End Select
                strTip = ""
                
                Set Pane = m.Chart.Tree("PRICE PANE")
                If Not Pane Is Nothing Then
                    If cboRiskGraphType.Text = "Probability" Then
                        strFormat = ""
                        If Pane.WoodPane Is Nothing Then
                            strTip = ""
                        Else
                            strTip = Pane.WoodPane.OptNavGraphToolTip
                        End If
                    ElseIf m.MouseLast.dY >= Pane.geAdjustMin And m.MouseLast.dY <= Pane.geAdjustMax Then
                        dPrice = RoundToMinMove(m.MouseLast.dY, m.Chart.Bars.MinMove)
                        Pane.OptNavGraphValues dPrice, dNow, dThird, dTwoThird, dExpire
                        Pane.OptNavGraphDate dDateNow, dDateThird, dDateTwoThird, dDateExpire
                        
                        If dNow <> kNullData Then
                            strTip = DateFormat(dDateNow) & ": " & Format(RoundNum(dNow, 6), strFormat)
                        End If
                        If dThird <> kNullData Then
                            If Len(strTip) = 0 Then
                                strTip = DateFormat(dDateThird) & ": " & Format(RoundNum(dThird, 6), strFormat)
                            Else
                                strTip = strTip & vbCrLf & DateFormat(dDateThird) & ": " & Format(RoundNum(dThird, 6), strFormat)
                            End If
                        End If
                        If dTwoThird <> kNullData Then
                            If Len(strTip) = 0 Then
                                strTip = DateFormat(dDateTwoThird) & ": " & Format(RoundNum(dTwoThird, 6), strFormat)
                            Else
                                strTip = strTip & vbCrLf & DateFormat(dDateTwoThird) & ": " & Format(RoundNum(dTwoThird, 6), strFormat)
                            End If
                        End If
                        If dExpire <> kNullData Then
                            If Len(strTip) = 0 Then
                                strTip = DateFormat(dDateExpire) & ": " & Format(RoundNum(dExpire, 6), strFormat)
                            Else
                                strTip = strTip & vbCrLf & DateFormat(dDateExpire) & ": " & Format(RoundNum(dExpire, 6), strFormat)
                            End If
                        End If
                    End If
                    
                    strTip = Chr(27) & strTip
                    With lblTipChart
                        .Alignment = vbLeftJustify
                        .Move 0, 0, Me.TextWidth(strTip) + 120, Me.TextHeight(strTip) + 50
                    End With
                    If Len(strTip) > 0 Then strPrevTip = ""
                End If
            Case Else
                strTip = ""
        End Select
        
        ' see if should be displayed to the left
        If Left(strTip, 1) = Chr(27) Then
            bToLeft = True
            strTip = Mid(strTip, 2)
        End If
    End If
                                                                        
    ' show tip
    If strTip <> strPrevTip Then
        strPrevTip = strTip
        lblTipChart.Caption = strTip
        With vseTipChart
            If Len(strTip) = 0 Then
                ' if no tip, then move off-screen
                .Top = -1000 - .Height
            Else
                .ZOrder
                nWidth = lblTipChart.Width
                nHeight = lblTipChart.Height
                
                If bToLeft And m.MouseLast.MouseX >= pbChart.Width / 3 Then
                    If pbChart.Top > m.MouseLast.MouseY - .Height Then
                        .Move m.MouseLast.MouseX - (nWidth + 120), m.MouseLast.MouseY + pbChart.Top, nWidth, nHeight
                    Else
                        .Move m.MouseLast.MouseX - (nWidth + 120), m.MouseLast.MouseY, nWidth, nHeight
                    End If
                Else
                    If pbChart.Top > m.MouseLast.MouseY + .Height Then
                        .Move m.MouseLast.MouseX + 200, m.MouseLast.MouseY, nWidth, nHeight
                    Else
                        .Move m.MouseLast.MouseX + 200, m.MouseLast.MouseY + pbChart.Top, nWidth, nHeight
                    End If
                End If
            End If
        End With
    End If

End Sub

Private Function IndFunctionTip(Annot As cAnnotation) As String
On Error GoTo ErrSection:

    Dim Ind As cIndicator
    Dim IndFunction As cFunction
    Dim strTip$, strUsage$, strDesc$
    Dim nTextWidth&
    Dim i&
    
    If m.Chart.Tree.NodeLevel(Annot.geIndId) > 0 Then
        Set Ind = m.Chart.Tree(Annot.geIndId)
    End If
    
    If Ind Is Nothing Then Exit Function
    
    Set IndFunction = New cFunction
    
    With IndFunction
        .FunctionID = Ind.FunctionID
        .Load
        strDesc = Trim(.Description)
        strUsage = Trim(.TradeSenseUsage)
    End With
    
    If Len(strDesc) = 0 And Len(strUsage) = 0 Then Exit Function
    
    strTip = ""
    Me.Font = lblTipChart.Font
    If Len(strUsage) > 0 Then
        strTip = "USAGE:  " & strUsage & " "
        i = 30 '(to set a minimum width if there is a longer description)
        If Len(strTip) < i And Len(strDesc) > i Then
            strTip = Left(strTip & Space(i), i)
        End If
        nTextWidth = Me.TextWidth(strTip)
    Else
        nTextWidth = 4020       'just a default that seems to look good
    End If
    
    strDesc = "INFO:  " & strDesc
    If Me.TextWidth(strDesc) > nTextWidth Then
        strDesc = strDesc & Space(100)
        If geWrapText(pbChart.hDC, strDesc, strTip) = 0 Then
            strTip = strTip & vbCrLf & Mid(strDesc, 2)
        End If
    Else
        strTip = strTip & vbCrLf & strDesc
    End If
   
    With lblTipChart
        'show tip to left of cursor if there is enough room
        If Abs(m.MouseLast.MouseX - pbChart.Left) > nTextWidth Then strTip = Chr(27) & strTip
        .Alignment = 0  'left
        .Move 0, 0, nTextWidth + 120, Me.TextHeight(strTip) + 50
    End With
    
    Set IndFunction = Nothing
    Set Ind = Nothing
    IndFunctionTip = strTip
    
    Exit Function

ErrSection:
    RaiseError Me.Name & ".IndFunctionTip"

End Function

Private Sub SetRiskRewardActivePt(Annot As cAnnotation, hitInfo As hittest_info)
On Error GoTo ErrSection:

    If hitInfo.location = 10 Or hitInfo.location = 11 Then
        If hitInfo.itemIndex = 0 Then
            m.nActiveAnnotPt = 2
        ElseIf hitInfo.itemIndex = 4 Or hitInfo.itemIndex = 3 Then
            m.nActiveAnnotPt = 1
        End If
        m.epbCursor = eCursor_Arrow4Way
    Else
        m.nActiveAnnotPt = 0
        m.epbCursor = eCursor_Hand
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetRiskRewardActivePt", eGDRaiseError_Raise
End Sub

Private Function HandleWaveLabels(ByVal bMouseUp As Boolean, _
    Optional ByVal bDblClick As Boolean = False, _
    Optional ByVal bMouseDown As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Static PrevX#, PrevY#
    Dim Annot As cAnnotation
    Dim bMoved As Boolean
    Dim bEnd As Boolean, bDelete As Boolean
    Dim aLabels As New cGdArray
    Dim nPoint&, dY#
    
    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    
    If Annot Is Nothing Then Exit Function
    
    If bDblClick = True Then
        Annot.VerifyWavePoints
        Annot.MenuMove = False
        Annot.MenuAdd = False
        bEnd = True
    ElseIf bMouseDown And Len(g.strActiveDraw) = 0 Then
        nPoint = Annot.Extendable(m.nActiveAnnotPt)
        If nPoint > 1 Then
            ShowAnnotPopup nPoint, vbLeftButton
        End If
    Else
        aLabels.SplitFields Annot.Text
        nPoint = m.nActiveAnnotPt
        dY = SnapToPrice(m.MouseLast, Annot.eType)
    End If
    
    If bMouseUp Then
        If m.MouseLast.nPaneID <> Annot.gePaneId Then
            bEnd = True
            If m.MouseLast.nPaneID > 0 Then bDelete = True
        ElseIf g.strActiveDraw = "ID_WaveLabels" Then
            If PrevX = m.MouseLast.nX And PrevY = dY Then
                HandleWaveLabels = False
                Exit Function
            Else
                PrevX = m.MouseLast.nX
                PrevY = dY
            End If
            'user selecting points for new annotation
            nPoint = nPoint + 1
            bMoved = Annot.MovePoint(m.Chart, nPoint, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
            m.nActiveAnnotPt = nPoint
        ElseIf Annot.MenuAdd Then
            nPoint = m.nActiveAnnotPt
            bMoved = Annot.MovePoint(m.Chart, nPoint, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
            nPoint = Annot.AddPoint(m.MouseLast.dDate, m.MouseLast.dY)
            m.nActiveAnnotPt = nPoint
            m.nObjectMoving = kMinMoveCount
        Else
            If m.nObjectMoving < kMinMoveCount Then
                'user single clicked to bring up editor or to end a move initiated from a menu
                If Not Annot.MenuMove Then
                    tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
                End If
            Else
                'user moving existing annotation
                bMoved = Annot.MovePoint(m.Chart, nPoint, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
            End If
            bEnd = True
        End If
    End If
    
    Annot.geDrawAnn m.Chart
    
    If bEnd Then
        ClearAnnotFlags bDelete
        m.nActiveIndIdx = 0
        m.nObjectMoving = 0
        m.epbCursor = eCursor_Default
        m.Chart.SetCursor
        PrevX = 0
        PrevY = 0
    End If
    
    Set Annot = Nothing
    
    HandleWaveLabels = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".HandleWaveLabels", eGDRaiseError_Raise

End Function

Private Sub ScrCapturePreview()
On Error GoTo ErrSection:
        
    Dim lAvailWidth As Long             ' Available width on the page
    Dim lAvailHeight As Long            ' Available height on the page
    Dim lPicHeight As Long              ' Picture height
    Dim lPicWidth As Long               ' Picture width
    Dim lResize As Long                 ' Resize amount for the picture
    
    With frmPrintPreview.vp
        .StartDoc
        ' Set the header and the footer
        DoPrintHeader
        ' Redraw the diagram on the given device context
        lAvailWidth = .PageWidth - .MarginLeft - .MarginRight
        lAvailHeight = .PageHeight - .CurrentY - .MarginBottom
        .CalcPicture = LoadPicture(AddSlash(App.Path) & "ScreenCapture.BMP")
        lPicHeight = .Y2 - .Y1
        lPicWidth = .X2 - .X1
        
        .X1 = 0
        .Y1 = 0
        
        If lPicHeight > lAvailHeight And lPicWidth > lAvailWidth Then
            .DrawPicture LoadPicture(AddSlash(App.Path) & "ScreenCapture.BMP"), _
                 .MarginLeft, .CurrentY, lAvailWidth, lAvailHeight, vppaZoom
        ElseIf lPicHeight > lAvailHeight Then
            lResize = (lPicWidth * (1 - (lAvailHeight / lPicHeight))) / 2
            .DrawPicture LoadPicture(AddSlash(App.Path) & "ScreenCapture.BMP"), _
                 .MarginLeft - lResize, .CurrentY, lPicWidth, lAvailHeight, vppaZoom
        ElseIf lPicWidth > lAvailWidth Then
            lResize = (lPicHeight * (1 - (lAvailWidth / lPicWidth))) / 2
            .DrawPicture LoadPicture(AddSlash(App.Path) & "ScreenCapture.BMP"), _
                 .MarginLeft, .CurrentY - lResize, lAvailWidth, lPicHeight, vppaZoom
        Else
            .DrawPicture LoadPicture(AddSlash(App.Path) & "ScreenCapture.BMP"), _
                 .MarginLeft, .CurrentY, lPicWidth, lPicHeight, vppaZoom
        End If
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ScrCapturePreview", eGDRaiseError_Raise

End Sub

Private Sub WaveLabelsCreate()
On Error GoTo ErrSection:

    If m.AnnotOptions Is Nothing Then
        Set m.AnnotOptions = New cAnnotation
        'call creatnew to get defaults
        m.AnnotOptions.CreateNew m.Chart, eANNOT_WaveLabels, 1, 0, 0, 0, 0, , , , , True
        
'Aardvark 3989: Don't show editor, just use defaults.
'        frmEditAnnot.ShowWaveLabels m.AnnotOptions
'        If Len(m.AnnotOptions.Text) > 0 Then
'            'reset this string as the form deactivate event clears it
'            g.strActiveDraw = "ID_WaveLabels"
'            If m.AnnotOptions.Prop("Lines") = 0 Then
'                m.AnnotOptions.Prop("LabelFirstPoint") = 1  'override default
'            End If
'        Else
'            'user cancelled
'            ClearAnnotFlags True
'            Set m.AnnotOptions = Nothing
'        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".WaveLabelsCreate", eGDRaiseError_Raise

End Sub

Private Function HandleTriChannel(ByVal bMouseUp As Boolean, _
    ByVal bMouseDown As Boolean) As Boolean
On Error GoTo ErrSection:

    Static nPrevX&, nPrevY&
    
    Dim Annot As cAnnotation
    Dim bMoved As Boolean
    Dim dY#, nValid&

    HandleTriChannel = False
        
    'check if this is initial create
    If m.nActiveAnnotIdx < 0 Then
        If m.MouseDown.bOffChart Then
            ClearAnnotFlags True
        Else
            StatusMsg "Now click on the second point ...", -1
            m.nPointCount = 1
            HandleTriChannel = True
            nPrevX = m.MouseDown.MouseX / Screen.TwipsPerPixelX
            nPrevY = m.MouseDown.MouseY / Screen.TwipsPerPixelY
        End If
        Exit Function
    End If
    
    'check if editing
    If bMouseUp = True Then
        If Len(g.strActiveDraw) = 0 Then
            If m.nObjectMoving < kMinMoveCount Then
                tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
            Else
                Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
                If Not Annot Is Nothing Then m.Chart.SyncGlobalAnnots Annot
            End If
            ClearAnnotFlags False
            m.nActiveIndIdx = 0
            m.nObjectMoving = 0
            m.Chart.SetCursor
            Exit Function
        End If
    End If

    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    If Annot Is Nothing Then Exit Function
     
    If m.MouseLast.bOffChart Then
        'user is selecting 2nd, 3rd or 4th point
        'mouse is in split-pane area, don't quit        -4240
    Else
        If m.MouseLast.nPaneID <> Annot.gePaneId Then
            If m.bAnnotCreated And bMouseUp Then
                ClearAnnotFlags True
                Exit Function
            End If
        End If
    End If
    
    dY = SnapToPrice(m.MouseLast, Annot.eType)
    If Annot.geMoveFlag = 0 And bMouseDown = True Then
        'user still selecting points
        Select Case m.nPointCount
            Case 1:
                If Abs(m.MouseDown.MouseX / Screen.TwipsPerPixelX - nPrevX) > kMinTriChannSize Or _
                   Abs(m.MouseDown.MouseY / Screen.TwipsPerPixelY - nPrevY) > kMinTriChannSize Then
                    
                    If m.MouseDown.bOffChart And m.MouseDown.nScalePaneId = -6 Then
                        bMoved = Annot.MovePoint(m.Chart, 2, m.nSplitPaneHittestID, 0, dY)
                    Else
                        bMoved = Annot.MovePoint(m.Chart, 2, m.MouseDown.nPaneID, m.MouseDown.nX, dY)
                        nPrevX = m.MouseDown.MouseX / Screen.TwipsPerPixelX
                    End If
                    If bMoved Then
                        If Annot.geMoveFlag = 0 Then StatusMsg "Now click on the third point ...", -1
                        m.nPointCount = 2
                        nPrevY = m.MouseDown.MouseY / Screen.TwipsPerPixelY
                    End If
                Else
                    StatusMsg "Now click on the second point ...", -1
                    bMoved = True       'so flags won't get cleared
                End If
                
            Case 2:
                If Abs(m.MouseDown.MouseX / Screen.TwipsPerPixelX - nPrevX) > kMinTriChannSize Or _
                   Abs(m.MouseDown.MouseY / Screen.TwipsPerPixelY - nPrevY) > kMinTriChannSize Then
                
                    If m.MouseDown.bOffChart And m.MouseDown.nScalePaneId = -6 Then
                        bMoved = Annot.MovePoint(m.Chart, 3, m.nSplitPaneHittestID, 0, dY)
                    Else
                        bMoved = Annot.MovePoint(m.Chart, 3, m.MouseDown.nPaneID, m.MouseDown.nX, dY)
                        nPrevX = m.MouseDown.MouseX / Screen.TwipsPerPixelX
                    End If
                    If bMoved Then
                        nPrevY = m.MouseDown.MouseY / Screen.TwipsPerPixelY
                        nValid = geValidTriangle(Annot.geAnnotObject)
                        If nValid = 1 Then
                            If Annot.eType = eANNOT_ChannelHighlight Then
                                StatusMsg "Now click on the fourth point ...", -1
                                m.nPointCount = 3
                            Else
                                StatusMsg
                                Annot.AssignDateTime
                                ClearAnnotFlags False
                                Annot.geMoveFlag = 1    'set flag indicating points selection is complete
                                m.Chart.SyncGlobalAnnots Annot
                                SyncDrawTools
                            End If
                        Else
                            StatusMsg "Now click on the third point ...", -1
                        End If
                    End If
                Else
                    StatusMsg "Now click on the third point ...", -1
                    bMoved = True       'so flags won't get cleared
                End If
                
            Case 3:
                If Annot.eType = eANNOT_ChannelHighlight Then
                    If Abs(m.MouseDown.MouseX / Screen.TwipsPerPixelX - nPrevX) > kMinTriChannSize Or _
                       Abs(m.MouseDown.MouseY / Screen.TwipsPerPixelY - nPrevY) > kMinTriChannSize Then
                       
                        If m.MouseDown.bOffChart And m.MouseDown.nScalePaneId = -1 Then
                            bMoved = Annot.MovePoint(m.Chart, 4, m.nSplitPaneHittestID, 0, dY)
                        Else
                            bMoved = Annot.MovePoint(m.Chart, 4, m.MouseDown.nPaneID, m.MouseDown.nX, dY)
                        End If
                        If bMoved Then
                            StatusMsg
                            Annot.AssignDateTime
                            ClearAnnotFlags False
                            Annot.geMoveFlag = 1    'set flag indicating points selection is complete
                            m.Chart.SyncGlobalAnnots Annot
                            SyncDrawTools
                        End If
                    Else
                        StatusMsg "Now click on the fourth point ...", -1
                        bMoved = True       'so flags won't get cleared
                    End If
                Else
                    bMoved = False
                    ClearAnnotFlags True, True
                End If
                
            Case Else
                bMoved = False
                ClearAnnotFlags True, True
        End Select
        
        If bMoved = False Then
            StatusMsg
            ClearAnnotFlags True, True
        End If
    ElseIf bMouseUp = False Then
        'moveflag = 1 means user is moving an existing, completed annotation
        If Annot.geMoveFlag = 1 Then
            bMoved = Annot.MovePoint(m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
        ElseIf Annot.eType = eANNOT_TriangleWedge Then
            If m.nPointCount = 2 Then Annot.geMoveFlag = 3
            bMoved = Annot.MovePoint(m.Chart, m.nPointCount + 1, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
        ElseIf Annot.eType = eANNOT_ChannelHighlight Then
            bMoved = Annot.MovePoint(m.Chart, m.nPointCount + 1, m.MouseLast.nPaneID, m.MouseLast.nX, dY)
        End If
        If bMoved Then m.epbCursor = eCursor_Blank
        Annot.geDrawAnn m.Chart
        If Annot.geMoveFlag = 3 Then Annot.geMoveFlag = 0
    End If
                
    Set Annot = Nothing
    
    HandleTriChannel = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".HandleTriChannel", eGDRaiseError_Raise
    
End Function

Private Function GetCustomTabOrder(ByVal bChartNameChanged As Boolean) As Boolean
On Error GoTo ErrSection:
    
    Dim i&, j&, iPos&
    Dim strText$, strFile$
    Dim frm As Form
    
    Dim tbForms As New cGdTable
    Dim aIdx As cGdArray
    Dim aCustomOrder As New cGdArray
    Dim bFound As Boolean

    m.aTabs.Size = 0
        
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\Page.ini"
    strText = GetIniFileProperty("ChartsOrder", "", "Tab", strFile)
    
    If Len(strText) > 0 Then
        aCustomOrder.SplitFields strText, ","
        tbForms.CreateField eGDARRAY_Strings, 0, "TemplateName"
        tbForms.CreateField eGDARRAY_Strings, 1, "TabInfo"
    End If
    
    For i = 0 To Forms.Count - 1
        If IsFrmChartMDI(Forms(i)) Then
            Set frm = Forms(i)
            If frm.Chart Is Nothing Then
                frm.tmr.Tag = "UNLOAD_NOW"
            ElseIf Len(frm.Chart.Symbol) = 0 Then
                frm.tmr.Tag = "UNLOAD_NOW"
            ElseIf frm.DetachStatus = eNotDetached Then
                If frm.tmr.Tag <> "UNLOADING" Then
                    If bChartNameChanged Then
                        strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & frm.Chart.Template & ".CHT"
                        strText = GetIniFileProperty("ChartName", "", "General", strFile)
                        frm.Chart.SetChartName strText
                        If Len(strText) = 0 Then
                            strText = frm.Chart.ChartName
                        End If
                    Else
                        strText = frm.Chart.ChartName
                    End If
                    If aCustomOrder.Size > 0 Then
                        tbForms.AddRecord ""
                        tbForms(0, tbForms.NumRecords - 1) = frm.Chart.Template
                        tbForms(1, tbForms.NumRecords - 1) = strText & vbTab & Trim(frm.vseCaption.Caption) _
                            & vbTab & CStr(frm.hWnd)
                    Else
                        ' TabTitle , Tooltip , hWnd
                        m.aTabs.Add strText & vbTab & Trim(frm.vseCaption.Caption) _
                            & vbTab & CStr(frm.hWnd)
                    End If
                End If
            End If
        End If
    Next
    
    If tbForms.NumRecords = 0 Then
        Set frm = Nothing
        GetCustomTabOrder = False
        Exit Function
    End If
    
    'look for new charts
    For i = 0 To tbForms.NumRecords - 1
        For j = 0 To aCustomOrder.Size - 1
            bFound = False
            If tbForms(0, i) = aCustomOrder(j) Then
                bFound = True
                Exit For
            End If
        Next
        If bFound = False Then
            aCustomOrder(aCustomOrder.Size) = tbForms(0, i)
        End If
    Next
    
    Set aIdx = tbForms.CreateSortedIndex(0, eGdSort_Default)
    
    strText = ""
    For i = 0 To aCustomOrder.Size - 1
        bFound = False
        If tbForms.SearchAsIndex(aIdx, 0, aCustomOrder(i), iPos) Then
            If tbForms(0, aIdx(iPos)) = aCustomOrder(i) Then
                bFound = True
                m.aTabs.Add tbForms(1, aIdx(iPos))
                strText = strText & aCustomOrder(i) & ","
            End If
        End If
    Next
    
    If Len(strText) > 0 Then
        strText = Left(strText, Len(strText) - 1) 'strip off last comma
    End If
        
    strFile = g.ChartGlobals.strCPCRoot & "\Charts\Page.ini"
    SetIniFileProperty "ChartsOrder", strText, "Tab", strFile
        
    Set frm = Nothing
    
    GetCustomTabOrder = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".GetCustomTabOrder", eGDRaiseError_Raise
    
End Function

Public Property Get WindowLink() As cWindowLink
On Error GoTo ErrSection:

    Set WindowLink = m.WindowLink

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".WindowLinkGet", eGDRaiseError_Raise

End Property

Public Property Get SymbolOrSymbolID() As Variant
On Error GoTo ErrSection:

    If Chart.SymbolID = 0& Then
        SymbolOrSymbolID = Chart.Symbol
    Else
        SymbolOrSymbolID = Chart.SymbolID
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".SymbolOrSymbolIDGet", eGDRaiseError_Raise

End Property

Public Property Get SymbolID() As Long
On Error GoTo ErrSection:

    SymbolID = Chart.SymbolID

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".SymbolIDGet", eGDRaiseError_Raise

End Property

Public Property Let SymbolID(ByVal nSymbolID As Long)
On Error GoTo ErrSection:

    If nSymbolID <> Chart.SymbolID Then
        Chart.SetSymbol nSymbolID, True
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError Me.Name & ".SymbolIDLet", eGDRaiseError_Raise

End Property

Public Property Get Periodicity() As Long
On Error GoTo ErrSection:

    Periodicity = Chart.Periodicity

ErrExit:
    Exit Property

ErrSection:
    RaiseError Me.Name & ".PeriodicityGet", eGDRaiseError_Raise

End Property

Public Property Let Periodicity(ByVal nPeriodicity As Long)
On Error GoTo ErrSection:
    
    If nPeriodicity <> Chart.Periodicity Then
        Chart.ChangeBarPeriod nPeriodicity, True
    End If
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".LetPeriodicity", eGDRaiseError_Raise
End Property

Public Property Get IsInGameMode() As Boolean
On Error GoTo ErrSection:

    IsInGameMode = m.bGameMode

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".IsInGameMode.Get", eGDRaiseError_Raise
    Resume ErrExit

End Property

Public Property Let IsInGameMode(ByVal bMode As Boolean)
On Error GoTo ErrSection:

    m.bGameMode = bMode
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".IsInGameMode.Let", eGDRaiseError_Raise

End Property

Public Property Get GameMode() As cGameMode
On Error GoTo ErrSection:

    Set GameMode = m.oGameMode

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".GameModeGet", eGDRaiseError_Raise

End Property

Public Property Let GameMode(oGameModeObj As cGameMode)
On Error GoTo ErrSection:

    Set m.oGameMode = oGameModeObj

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".GameModeLet", eGDRaiseError_Raise
    Resume ErrExit

End Property

Public Property Get GameReplayMode() As eGDReplayMode
On Error GoTo ErrSection:

    If m.oGameMode Is Nothing Then
        GameReplayMode = eGDReplayMode_Off
    Else
        GameReplayMode = m.oGameMode.GameReplayMode(tmrGameMode.Enabled)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".GameReplayMode", eGDRaiseError_Raise

End Property

Private Sub GameSpeed()
On Error GoTo ErrSection:
    
    Dim bEnabledSave As Boolean
    Dim strText$
    
    bEnabledSave = tmrGameMode.Enabled
    tmrGameMode.Enabled = False
    
    With sldSpeed
        tmrGameMode.Interval = Round((2 ^ (.Value - .Min)) * 62.5)
        If tmrGameMode.Interval = 1000 Then
            strText = "1 second"
        ElseIf tmr.Interval > 1000 Then
            strText = Str(tmrGameMode.Interval / 1000#) & " seconds"
        Else
            strText = "1/" & Str(Round(1000# / tmrGameMode.Interval)) & " second"
        End If
    End With
    
    tmrGameMode.Enabled = bEnabledSave
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".GameSpeed", eGDRaiseError_Raise

End Sub

Private Sub UpdateGameModeLabels()
On Error GoTo ErrSection:

    Dim dDollars#

    'set current position label
    If m.Chart.Position > 0 Then
        lblPosition = "Long " & Str(m.Chart.Position)
    ElseIf m.Chart.Position < 0 Then
        lblPosition = "Short " & Str(Abs(m.Chart.Position))
    Else
        lblPosition = "None"
    End If
    
    'set open equity label
    dDollars = m.Chart.OpenEquity
    If dDollars > 0 Then
        lblOpenEquity = Format(dDollars, "$#,##0;-$#,##0")
        lblOpenEquity.ForeColor = RGB(0, 128, 0)
    ElseIf dDollars < 0 Then
        lblOpenEquity = Format(dDollars, "$#,##0;-$#,##0")
        lblOpenEquity.ForeColor = vbRed
    Else
        lblOpenEquity = ""
    End If
    
    'set profit label
    dDollars = m.Chart.NetProfit
    If dDollars > 0 Then
        lblProfit = Format(dDollars, "$#,##0;-$#,##0")
        lblProfit.ForeColor = RGB(0, 128, 0)
    ElseIf dDollars < 0 Then
        lblProfit = Format(dDollars, "$#,##0;-$#,##0")
        lblProfit.ForeColor = vbRed
    Else
        lblProfit = ""
    End If
    
    EnableGameControls          '4767

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".updategamemodelabels", eGDRaiseError_Raise

End Sub

Private Sub HandleGameOrder(ByVal eOrderMode As eGDTTEditOrderMode, _
    Optional ByVal nIdx& = -1, Optional Annot As cAnnotation = Nothing, _
    Optional nBuy As Byte = 255)
On Error GoTo ErrSection:
    
    Dim strSymbol$, dDollars#
    Dim Order As cPtOrder               ' Order to create
    Dim nReturn As eGDEditOrderReturn   ' Return from the create order call
    Dim aFlds As cGdArray
    Dim dMarket#, strMarket$, dPrice#
    
    'aardvark 3026 fix
    If m.Chart.LastGoodDataBar(False) < 0 Then
        InfBox "Cannot place order until there is data on chart.", "E", , "Instant Replay"
        Exit Sub
    End If
    
    If m.bGameOrderMoving And Annot Is Nothing Then
        m.bGameOrderMoving = False
        Exit Sub
    ElseIf m.oGameMode.GameStrategyID > 0 And m.oGameMode.GameStrategyID <> kGameModeSysID Then
        Exit Sub
    End If
    
    cmdStop_Click
    
    If eOrderMode = eGDTTEditOrderMode_GameNewOrder Then
        Set Order = New cPtOrder
        strSymbol = RollSymbolForDate(GetSymbol(m.Chart.Symbol), Int(m.Chart.Bars(eBARS_DateTime, m.Chart.Bars.Size - 1)))
        If Len(strSymbol) = 0 Then strSymbol = m.Chart.Symbol
        
        Order.SymbolOrSymbolID = strSymbol
        Order.AccountID = 0
        Order.OrderID = 0&
    ElseIf eOrderMode = eGDTTEditOrderMode_GameEditOrder Then
        If nIdx <> 0 Then
            Set Order = m.oGameMode.OpenOrder(nIdx + 1000)
        End If
    End If
    
    If Order Is Nothing Then Exit Sub
    
    If m.bGameOrderMoving Then
        nReturn = eGDEditOrderReturn_Submit
        Set aFlds = New cGdArray
        aFlds.SplitFields Annot.Text
        dPrice = m.Chart.Bars.PriceFromString(aFlds(1))
        If Order.OrderType = eTT_OrderType_Limit Then
            Order.LimitPrice = dPrice
            Order.StopPrice = kNullData
        ElseIf Order.OrderType = eTT_OrderType_Stop Then
            Order.StopPrice = dPrice
            Order.LimitPrice = kNullData
        ElseIf Order.OrderType = eTT_OrderType_StopWithLimit Then
            If UCase(aFlds(2)) = "LIMIT" Then
                Order.LimitPrice = dPrice
            ElseIf UCase(aFlds(2)) = "STOP" Then
                Order.StopPrice = dPrice
            End If
        End If
        
        If m.oGameMode.IsTargetOrder(Order) Then
            m.oGameMode.CalcNewTargetValues dPrice
        ElseIf m.oGameMode.IsStopLossOrder(Order) Then
            m.oGameMode.CalcNewStopValues dPrice
        End If
        
        m.bGameOrderMoving = False
    ElseIf m.oGameMode.IsTargetOrder(Order) Or m.oGameMode.IsStopLossOrder(Order) Then
        nReturn = eGDEditOrderReturn_Cancel
        frmGameTargetLoss.ShowMe m.oGameMode, Me, Nothing
    Else
        'pass & overide limit/stop prices in order form
        If m.Chart.Bars.Prop(eBARS_PeriodType) = ePRD_Days And Not m.oGameMode.HasIntradayTicks Then
            If m.oGameMode.GameDataTime > m.oGameMode.MidSessionTime Then
                dMarket = m.Chart.Bars(eBARS_Close, m.Chart.LastGoodDataBar(False))
            Else
                dMarket = m.Chart.Bars(eBARS_Open, m.Chart.LastGoodDataBar(False))
            End If
        Else
            dMarket = m.Chart.Bars(eBARS_Close, m.Chart.LastGoodDataBar(False))
        End If
        If dMarket <> kNullData Then
            strMarket = m.Chart.Bars.PriceDisplay(dMarket)
        Else
            strMarket = ""
        End If
        nReturn = frmTTEditOrder.ShowMe(Order, nBuy, eOrderMode, m.Chart.Position, strMarket)
    End If

        
    If nReturn = eGDEditOrderReturn_Cancel Then Exit Sub
    
    If eOrderMode = eGDTTEditOrderMode_GameEditOrder Then
        m.Chart.DeleteGameOrders
        If nReturn = eGDEditOrderReturn_Submit Then
            m.oGameMode.GameUpdateOrder
        ElseIf nReturn = eGDEditOrderReturn_Park Then
            m.oGameMode.GameDeleteOrder Order

        End If
    ElseIf eOrderMode = eGDTTEditOrderMode_GameNewOrder And nReturn = eGDEditOrderReturn_Submit Then
        m.oGameMode.GameNewOrder Order
    End If
    
    If m.Chart.Position = 0 Then
        m.oGameMode.RemoveAutoExits True, True
    End If
    
    UpdateGameModeLabels
    EnableGameControls
    m.Chart.GenerateChart eRedo1_Scrolled
    If m.eReplayModeSave = eGDReplayMode_Play Then cmdPlay_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".HandleGameOrder", eGDRaiseError_Raise

End Sub

Private Sub HandleGameMoveOrder(ByVal bMouseUp As Boolean, Annot As cAnnotation)
On Error GoTo ErrSection:

    Dim dY#
    
    m.bGameOrderMoving = False
    
    If Annot Is Nothing Then
        Exit Sub
    ElseIf m.oGameMode.GameStrategyID > 0 And m.oGameMode.GameStrategyID <> kGameModeSysID Then
        Exit Sub
    End If

    If Not bMouseUp Then cmdStop_Click
    
    dY = SnapToPrice(m.MouseLast, Annot.eType)
    m.bGameOrderMoving = Annot.MovePoint(m.Chart, 0, Annot.gePaneId, Annot.X(1), dY#)
    
    If m.bGameOrderMoving Then
        Annot.geDrawAnn m.Chart
    ElseIf Annot.Y(1) = dY Then
        'cAnnotation.MovePoint returns false if existing Y values are same as passed in new y-value (ie dY)
        'for gamemode we don't care as long as the y-values are valid
        'setting this to true prevents profit loss dialog from showing - issue 6689
        m.bGameOrderMoving = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".HandleGameMoveOrder", eGDRaiseError_Raise

End Sub

Public Sub EnableGameControls()
On Error GoTo ErrSection:

    If m.Chart.SystemID = 0 Or m.Chart.SystemID = kGameModeSysID Then
        If m.Chart.Position = 0 Then
            Enable cmdBuy
            Enable cmdSell
            Disable cmdExitNow
            Disable cmdAutoExits
            Disable cmdRevPos
        Else
            Disable cmdBuy
            Disable cmdSell
            Enable cmdExitNow
            Enable cmdAutoExits
            Enable cmdRevPos
        End If
    Else
        Disable cmdBuy
        Disable cmdSell
    End If
    
    If m.Chart.NetProfit = 0 Then
        Disable cmdReport
    Else
        Enable cmdReport
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".EnableGameControls", eGDRaiseError_Raise

End Sub

Private Sub SetFocusCtl()
On Error Resume Next
    
    If g.bSkipSetChartFocus Then Exit Sub
    
    If Me.IsInGameMode Then
        MoveFocus sldSpeed
    Else
        MoveFocus pbChart
    End If

End Sub

Public Property Get GameRptEnable() As Boolean
On Error Resume Next

    GameRptEnable = cmdReport.Enabled
    
End Property

Public Sub GenerateGameReport()
On Error GoTo ErrSection:
    
    Dim strReport$

    cmdStop_Click
    
    SetFocusCtl
    
    strReport = m.oGameMode.GameReport
    
    If m.oGameMode.GameStrategyID > 0 And m.oGameMode.GameStrategyID <> kGameModeSysID Then
        m.Chart.ShowSystemReport False
    ElseIf Len(strReport) > 0 Then
        m.oGameMode.ShowGameReport strReport
    End If
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".GenerateGameReport"

End Sub

Public Sub ClearReplaySync()
On Error GoTo ErrSection:

    cmdStop_Click
    If Not m.oGameMode Is Nothing Then
        If m.oGameMode.GameReplayMode(False) = eGDReplayMode_Sync Then
            Set m.oGameMode = Nothing
            m.bGameMode = False
            m.Chart.ResetLastScreenDate
            m.Chart.GenerateChart eRedo9_ReloadData
            DoEvents
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".ClearReplaySync"
    
End Sub

Private Sub HandleChartOrder(ByVal bNewOrder As Boolean, _
    Optional ByVal eClickOrder As enumOneClickOrder = eClickOrder_None, _
    Optional ByVal nQty& = 0, _
    Optional ByVal strBuySell$ = "")
On Error GoTo ErrSection:
'This sub intended to be called when user is done adding an
'order on a chart or single-clicking to edit/cancel order (i.e.
'the mouse up event calls this). Do not call from mouse move or
'code that would make the call in rapid successions.
    
    Static bInProgress As Boolean
    
    Dim i&, nMoved&
    Dim strAnswer$, strOrdOne$, strOrdTwo$
    Dim dPriceLast#, dPriceNew#, dMinMove#
        
    Dim Order As cPtOrder
    Dim ExistingOrders As cGdTree
    
    Dim TrigOrds1 As cGdTree
    Dim TrigOrds2 As cGdTree
    
    Dim X&, Y&
    Dim bBracketOk As Boolean
        
    nMoved = m.nObjectMoving
    m.nObjectMoving = 0
                       
    If m.eOrdBarMode = eOrdBarMode_Wizard Then Exit Sub
    
    If bInProgress Then Exit Sub
                   
    If eClickOrder = eClickOrder_None Then
        If m.MouseLast.bOffChart Or m.Chart.Tree("PRICE PANE").gePaneId <> m.MouseLast.nPaneID Then
            ClearAnnotFlags False
            m.nActiveOrderID = 0
            m.nActiveOrderLoc = 0
            ClearBuySellButtons
            m.Chart.GenerateChart eRedo1_Scrolled
            Exit Sub
        End If
    End If
        
    Set Order = New cPtOrder
    
    bInProgress = True              'fix for double-click portion of 6159
    If bNewOrder Then
        i = m.Chart.LastGoodDataBar(False)
        dMinMove = m.Chart.Bars.MinMove()
        dPriceLast = RoundToMinMove(m.Chart.Bars(eBARS_Close, i), dMinMove)
        
        Order.AccountID = m.Chart.TradeAccountID
        Order.SymbolOrSymbolID = m.Chart.SymbolID
        Order.Exchange = cboExchanges.Text
        Order.OrderID = -1
        
        If eClickOrder = eClickOrder_None Then
            dPriceNew = m.Chart.Bars.RoundToPrice(m.MouseDown.dY, m.MouseDown.dDate)
            
            If SetOrderType(Order, eClickOrder, dPriceNew, dPriceLast) Then
                Order.OrderPrice(False) = dPriceNew
                Order.Quantity = m.Quantity.Price
                
                If vseBracketOrder.Appearance = apInset Then
                    If OkayToExecute(Order, dPriceLast, True, Me) Then
                        Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(m.Chart.TradeAccountID))
                        g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.Chart.TradeAccountID), "Creating Order from Chart: " & Order.OrderText, True
                        If m.oBracketOrdOne Is Nothing Then
                            Set m.oBracketOrdOne = Order
                            ParkOrder m.oBracketOrdOne
                        ElseIf Not m.oBracketOrdOne Is Nothing Then
                            bBracketOk = True
                            
                            If g.Broker.ConfirmOrder(TradeAccountID, SymbolOrSymbolID) Then
                                strOrdOne = m.oBracketOrdOne.OrderText
                                strOrdTwo = Order.OrderText
                                
                                GetPromptLocation X, Y
                                If InfBox(strOrdOne & vbCrLf & strOrdTwo, "?", "+Ok|-Cancel", "Bracket order confirm", , , , , , , , , , , X, Y) = "C" Then
                                    bBracketOk = False
                                    Set Order = Nothing
                                End If
                            End If
                            
                            If bBracketOk Then
                                Set m.oBracketOrdTwo = Order
                                ParkOrder m.oBracketOrdTwo
                                
                                If g.Broker.HoldOcoAtBroker(Order.AccountID) Then
                                    Order.BrokerCancelOrderID = -(m.oBracketOrdOne.OrderID)
                                    m.oBracketOrdOne.BrokerCancelOrderID = -(m.oBracketOrdTwo.OrderID)
                                Else
                                    Order.CancelOrderID = m.oBracketOrdOne.OrderID
                                    
                                    ' DAJ 09/02/2009: Must set the cancel order ID on order one to the
                                    ' ID of order two here because the chart is not getting refreshed
                                    ' on the save in the Park routine...
                                    m.oBracketOrdOne.CancelOrderID = m.oBracketOrdTwo.OrderID
                                End If
                            End If
                        End If
                    End If
                Else
                    If g.Broker.ConfirmOrder(TradeAccountID, SymbolOrSymbolID) Then
                        CreateOrder , , , Order, , "Chart"
                    ElseIf OkayToExecute(Order, dPriceLast) Then
                        Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(m.Chart.TradeAccountID))
                        g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.Chart.TradeAccountID), "Creating Order from Chart: " & Order.OrderText, True
                        SubmitOrder Order
                    End If
                End If
            Else
                'user cancelled
                Set Order = Nothing
            End If
        Else
            dPriceNew = SetClickOrderParam(Order, eClickOrder, nQty)
            If dPriceNew <> kNullData Then
                If g.Broker.ConfirmOrder(TradeAccountID, SymbolOrSymbolID) Then
                    CreateOrder , , , Order, , "Chart"
                ElseIf OkayToExecute(Order, dPriceLast) Then
                    Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(m.Chart.TradeAccountID))
                    g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.Chart.TradeAccountID), "Creating Order from Chart: " & Order.OrderText, True
                    SubmitOrder Order
                End If
            End If
        End If
    ElseIf Order.Load(m.nActiveOrderID) Then
        If m.nActiveOrderLoc = 2 Then           'always allow order-cancel
            g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.Chart.TradeAccountID), "Cancelling Order from Chart: " & Order.OrderText, True
            CancelOrder Order, g.Broker.ConfirmOrder(TradeAccountID, SymbolOrSymbolID), , True
            m.Chart.UpdateOnlineOrder Order, False  'not necessary, but gives a little faster feedback for 5841
        ElseIf OrderIsPending(Order) Then
            InfBox "This order cannot be modified because it is in a pending state.  Please wait for order confirmation.", "!", , "Chart Order Error"
        ElseIf nMoved > kMinOrderMove Then          '4131
            dPriceLast = Order.OrderPrice(True)
            dPriceNew = m.Chart.Bars.RoundToPrice(m.MouseLast.dY, m.MouseLast.dDate)
            
            g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.Chart.TradeAccountID), "Modifying Order from Chart: " & Order.OrderText, True
            If dPriceLast <> dPriceNew Then
                ModifyOrder Order, dPriceNew, , g.Broker.ConfirmOrder(TradeAccountID, SymbolOrSymbolID)
            End If
        Else
            g.Broker.BrokerDebug g.Broker.AccountTypeForID(m.Chart.TradeAccountID), "Modifying Order from Chart: " & Order.OrderText, True
            EditOrderFromOrder Order, "Chart"
        End If
    End If
        
    If vseBracketOrder.Appearance = apInset And Not Order Is Nothing Then
        If m.oBracketOrdTwo Is Nothing And Not m.oBracketOrdOne Is Nothing Then
            'user need to place 2nd order
            ToolbarSetCursorGroup frmMain.tbToolbar, False, "ID_ChartOrderBuy"
        Else
            If Not m.oBracketOrdOne Is Nothing And Not m.oBracketOrdTwo Is Nothing Then
                ' DAJ 09/02/2009: Allow the SubmitOrder routine to handle submitting the
                ' other side of the OCO, but tell it not to ask the user about it...
                SubmitOrder m.oBracketOrdOne, , , , False
                ''SubmitOrder m.oBracketOrdTwo
            End If
            ClearBuySellButtons True
            Set m.oBracketOrdOne = Nothing
            Set m.oBracketOrdTwo = Nothing
        End If
    Else
        ClearBuySellButtons True
        Set m.oBracketOrdOne = Nothing
        Set m.oBracketOrdTwo = Nothing
    End If
    
    m.Chart.GenerateChart eRedo1_Scrolled
        
ErrExit:
    Set Order = Nothing
    Set ExistingOrders = Nothing
    Set TrigOrds1 = Nothing
    Set TrigOrds2 = Nothing
    bInProgress = False
    
    m.nActiveOrderID = 0
    m.nActiveOrderLoc = 0
    
    Exit Sub

ErrSection:
    bInProgress = False
    RaiseError Me.Name & ".HandleChartOrder"
    Resume ErrExit
    
End Sub

Private Function SetClickOrderParam(Order As cPtOrder, ByVal eOneClickType As enumOneClickOrder, _
            Optional ByVal nQty& = 0) As Double
On Error GoTo ErrSection:

    Dim dPrice#, strErr$
    Dim bOkay As Boolean
    
    SetClickOrderParam = kNullData
    
    m.Chart.UpdateTradePrices
    
    If nQty = 0 Then nQty = m.Quantity.Price
    If nQty = 0 Then Exit Function
    
'    strErr = ValidSecForAccount(m.Chart.Bars, m.Chart.TradeAccountID, m.Chart.TradeAccountNumber)
'    If Len(strErr) <> 0 Then
'        InfBox strErr, "E", , "Error"
'        Exit Function
'    End If

    dPrice = kNullData
    Select Case eOneClickType
        Case eClickOrder_BuyAsk, eClickOrder_SellAsk
            dPrice = m.Chart.Bars.PriceFromString(lblAsk.Caption)
        Case eClickOrder_BuyBid, eClickOrder_SellBid
            dPrice = m.Chart.Bars.PriceFromString(lblBid.Caption)
        Case eClickOrder_BuyMkt, eClickOrder_SellMkt
            dPrice = m.Chart.Bars.PriceFromString(lblMarket.Caption)
    End Select
    
    If dPrice = kNullData Then
        SetClickOrderParam = kNullData
        Exit Function
    Else
        bOkay = True
    End If
    
    If eOneClickType = eClickOrder_BuyAsk Or _
       eOneClickType = eClickOrder_BuyBid Or _
       eOneClickType = eClickOrder_BuyMkt Then
        
        Order.Buy = True
    
    ElseIf eOneClickType = eClickOrder_SellAsk Or _
           eOneClickType = eClickOrder_SellBid Or _
           eOneClickType = eClickOrder_SellMkt Then
           
           Order.Buy = False
    Else
        bOkay = False
    End If
    
    If Not bOkay Then
        SetClickOrderParam = kNullData
        Exit Function
    End If
    
    If eOneClickType = eClickOrder_BuyMkt Or _
       eOneClickType = eClickOrder_SellMkt Then
       
        Order.OrderType = eTT_OrderType_Market
       
    Else
    
        Order.OrderType = eTT_OrderType_Limit
    End If
    
    Order.OrderPrice(False) = dPrice
    Order.Quantity = nQty
    
    SetClickOrderParam = dPrice
    
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".SetClickOrderParam", eGDRaiseError_Raise

End Function

Private Function MouseDownActiveDraw() As Boolean
On Error GoTo ErrSection:

    Dim Ind As cIndicator, i&

    If g.strActiveDraw = "ID_DNRetracement" Or g.strActiveDraw = "ID_DNExpansion" _
        Or g.strActiveDraw = "ID_DNExpansion2" Or g.strActiveDraw = "ID_DNExpansion3" _
        Or g.strActiveDraw = "ID_DNExpansion4" _
        Or g.strActiveDraw = "ID_AndrewFork" Or g.strActiveDraw = "ID_WaveLabels" _
        Or g.strActiveDraw = "ID_ElliotTimeRatio" Or g.strActiveDraw = "ID_FibABCD" _
        Or g.strActiveDraw = "ID_Gartley" Or g.strActiveDraw = "ID_GannacciSwing2" Then
        If m.nActiveAnnotIdx > 0 Then
            MouseDownActiveDraw = True
        End If
    ElseIf g.strActiveDraw = "ID_Triangle" Or g.strActiveDraw = "ID_ChannelHighlight" Then
        If m.nActiveAnnotIdx > 0 Then
            HandleTriChannel False, True
            MouseDownActiveDraw = True
        End If
    ElseIf g.strActiveDraw = "ID_FibClusters" Then
        If Not m.Chart.Annots(kClusterZoneRect) Is Nothing Then
            ClearAnnotFlags False
            SyncDrawTools
            MouseDownActiveDraw = True      'just so the mousedown event stop processing
        End If
    ElseIf g.strActiveDraw = "ID_Mirror" Or g.strActiveDraw = "ID_Pattern" Or g.strActiveDraw = "ID_PivotPoints" Then
        If m.Chart.Tree.Key(m.MouseDown.nPaneID) = "PRICE PANE" Then
            Set Ind = m.Chart.Tree("PRICE")
            If Ind Is Nothing Then
                g.strActiveDraw = ""
                ClearAnnotFlags True
                MouseDownActiveDraw = False
            ElseIf g.strActiveDraw = "ID_PivotPoints" Then
                If m.nActiveAnnotIdx > 0 Then
                    MouseDownActiveDraw = True
                    m.nPointCount = 2
                End If
            ElseIf Not Ind.OkayToMirror Then
                If g.strActiveDraw = "ID_Mirror" Then
                    MsgBox "Price bars can only be mirrored when displayed as" & vbCrLf & "Line, Points, HL, HLC, OHLC, Candle or Bollinger bars."
                Else
                    MsgBox "Price bars patterns can only be saved for copying" & vbCrLf & "when the price indicator is displayed as Line, Points," & vbCrLf & "HL, HLC, OHLC, Candle or Bollinger bars."
                End If
                ClearAnnotFlags True
                MouseDownActiveDraw = False
            ElseIf m.nActiveAnnotIdx > 0 Then
                MouseDownActiveDraw = True
                m.nPointCount = 2
            End If
        Else
            MsgBox "This tool can only be used in the price pane."
            ClearAnnotFlags True
            MouseDownActiveDraw = True
        End If
    ElseIf m.nActiveAnnotIdx > 0 Then
        MouseDownActiveDraw = True
        If m.nPointCount <= 2 Then m.nPointCount = 2
    End If
    
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".MouseDownActiveDraw"

End Function

Private Sub ShowOrderActionMenu()
On Error GoTo ErrSection:

    Dim OrderStruct As cOrderStruct
    
    TopMost = False
    
    Set OrderStruct = m.Chart.OnlineOrders(Str(m.nActiveOrderID))
    
    If Not OrderStruct Is Nothing Then
        If OrderStruct.ItemStatus = eTT_OrderStatus_TriggerPending Then
            mnuOrderSubmit.Enabled = True
            mnuOrderPark.Enabled = True
        ElseIf OrderStruct.ItemStatus = eTT_OrderStatus_Working Then
            mnuOrderSubmit.Enabled = False
            mnuOrderPark.Enabled = True
        Else
            mnuOrderSubmit.Enabled = False
            mnuOrderPark.Enabled = False
        End If
        
        mnuSep8.Visible = True 'FileExist(AddSlash(App.Path) & "AdvExit.FLG")
        mnuManageXOS.Visible = True 'FileExist(AddSlash(App.Path) & "AdvExit.FLG")
        mnuRemoveXOS.Visible = True 'FileExist(AddSlash(App.Path) & "AdvExit.FLG")
        mnuSelectXOS.Visible = True 'FileExist(AddSlash(App.Path) & "AdvExit.FLG")
        
        PopupMenu mnuOrderAction
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ShowOrderActionMenu"

End Sub

Public Sub FixOrderBarControls(Optional ByVal bUpdateAccts As Boolean = True, _
    Optional ByVal bPositionCtrls As Boolean = False, _
    Optional ByVal strCtrlsToShow As String = "", _
    Optional ByVal bOrderModeChanged As Boolean = False)
On Error GoTo ErrSection:
    
    Static bPrevReplayVisible As Boolean
    
    Dim Ctrl1 As Control
    Dim Ctrl2 As Control
    
    Dim bLocked As Boolean
    Dim bShowExchanges As Boolean       ' Show the exchange controls?
    Dim bNewChain As Boolean
        
    Dim Pane As cPane
    
    Dim iCtlIndex As Long
    Dim i&, strIndListPFP$
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "UNLOADING" Then Exit Sub
    If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then Exit Sub
    
    If m.eOrdBarMode = eOrdBarMode_PFP Then
        If fraFrontMonth.Visible Then fraFrontMonth.Visible = False
        If fraOrderBarMode.Visible Then fraOrderBarMode.Visible = False
        If fraOrderBtns.Visible Then fraOrderBtns.Visible = False
        
        If fraOrdWizard.Visible Then
            fraOrdWizard.Visible = False
            SetFlexVisible eFlexGridIdx_OrdWizard, False
        End If
            
        If fraPrices.Visible Then fraPrices.Visible = False
        If fraBrokerDisconnect.Visible Then fraBrokerDisconnect.Visible = False
        
        'show flex grids if necessare
        SetFlexVisible eFlexGridIdx_PfpInd, True
        SetFlexVisible eFlexGridIdx_PfpHits, True
        
        With fraPatternProfit
            .Width = vseOrderBar.Width
            .Height = vseOrderBar.Height        ' - 90
            
            With fgChartFlex(eFlexGridIdx_PfpInd)
                .Left = vseOrderBar.Left + 15
                .Width = fraPatternProfit.Width - 75
                .Height = fgChartFlex(eFlexGridIdx_PfpInd).RowHeight(0) * 4
                
                If Me.WindowState = vbMaximized Then
                    .Top = lblPfpInd.Top + lblPfpInd.Height + 225
                    lblHitsFoundPFP.Top = lblPfpInd.Top + lblPfpInd.Height + .Height + 150
                Else
                    .Top = lblPfpInd.Top + lblPfpInd.Height - 75
                    lblHitsFoundPFP.Top = lblPfpInd.Top + lblPfpInd.Height + .Height + 120
                End If
                
                fgChartFlex(eFlexGridIdx_PfpHits).Left = .Left
                fgChartFlex(eFlexGridIdx_PfpHits).Width = .Width
                fgChartFlex(eFlexGridIdx_PfpHits).Top = .Top + .Height + lblHitsFoundPFP.Height + 150
            End With
            
            '5809 - order bar showing instead of PFP bar
            '   original code did not check that value is non-negative before assigning it to grid's
            '   height, as result, code went to errsection & never got to code that made PFP bar visible
            i = .Height - lblHitsFoundPFP.Top - hsb.Height - lblHitsFoundPFP.Height
            If i >= fgChartFlex(eFlexGridIdx_PfpInd).Height Then
                fgChartFlex(eFlexGridIdx_PfpHits).Height = i
            Else
                fgChartFlex(eFlexGridIdx_PfpHits).Height = fgChartFlex(eFlexGridIdx_PfpInd).Height
            End If
            
'            fgChartFlex(eFlexGridIdx_PfpInd).Move 15, fgChartFlex(eFlexGridIdx_PfpInd).Top, .Width - 15
            lblPatternLen.Width = .Width - 65
            
            If Not .Visible Then
                If m.oPatternProfit Is Nothing Then Set m.oPatternProfit = New cPatternProfit
                txtForecastPFP.Text = Str(m.oPatternProfit.NumForecastBars)
                txtCorrPercentPFP.Text = Str(m.oPatternProfit.PercentCorr)
                strIndListPFP = LoadIndGridPFP(m.Chart, fgChartFlex(eFlexGridIdx_PfpInd), True, True)
                If m.oPatternProfit.ShowSettings Then
                    pbLeft.Visible = False
                    pbRight.Visible = False
                    lblAccounts.Visible = False
                    cboAccounts.Visible = False
                    
                    cmdPfpSettings_Click
                
                    pbLeft.Visible = True
                    pbRight.Visible = True
                    lblAccounts.Visible = True
                    cboAccounts.Visible = True
                End If
                i = m.oPatternProfit.Heatmap
                If i = vbGrayed Then
                    m.oPatternProfit.Heatmap = vbUnchecked
                    i = vbUnchecked
                End If
                chkApplyFilter.Tag = "NoPrompt"
                chkApplyFilter.Value = i
                chkApplyFilter_Click
                chkApplyFilter.Tag = ""
                .Visible = True
            End If
        End With
        
        GoTo ErrExit
    End If
    
    'make sure grids are not visible if not in pfp mode
    SetFlexVisible eFlexGridIdx_PfpInd, False
    SetFlexVisible eFlexGridIdx_PfpHits, False
    
    i = m.Chart.HighlightPos
    If i = -1 Then
        lblTradePos.BackColor = fraOrderBtns.BackColor
    Else
        lblTradePos.BackColor = i
    End If
    
    i = m.Chart.HighlightEquity
    If i = -1 Then
        lblEquity.BackColor = fraOrderBtns.BackColor
    Else
        lblEquity.BackColor = i
    End If
    
    'determine the type of order bar to be shown: order bar, option wizard or disconnect message frame
    If bPositionCtrls And Len(strCtrlsToShow) = 0 And Me Is ActiveChart Then        '4946
        If m.eOrdBarMode = eOrdBarMode_Order Then
            strCtrlsToShow = m.Chart.OrdBarCtrls
        ElseIf m.eOrdBarMode = eOrdBarMode_BrokerDisconnect Then
            strCtrlsToShow = kOrdBarDisconnect
        ElseIf m.Chart.Bars.IsIntraday Or m.Chart.Periodicity >= ePRD_Months + 1 Then
            m.eOrdBarMode = eOrdBarMode_Order
            strCtrlsToShow = m.Chart.OrdBarCtrls
            InfBox "Options order bar can only be used with Daily and Weekly charts.", "I", , "Information"         '4949
        Else
            If Not m.bOptNavLoaded Then
                If g.nOptNavStatus = eGDOptNavStatus_Loaded Then
                    m.bOptNavLoaded = True                  '6098
                ElseIf g.nOptNavStatus = eGDOptNavStatus_Loading Then
                    Do While g.nOptNavStatus <> eGDOptNavStatus_Loaded
                        DoEvents
                    Loop
                    m.bOptNavLoaded = True
                End If
            End If
            
            'check status in case user detached chart while this code is waiting OptNav startup
            If m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then GoTo ErrExit
            
            If m.bOptNavLoaded = True Then       'need this check because g.nOpNavStatus = eGDOptNavStatus_Unloaded when OptionsNav is hidden (6060, 6066)
                If m.Chart.ChainFromFile(bNewChain, bOrderModeChanged) Then
                    If bNewChain Or bOrderModeChanged Then WizardGridClear
                    strCtrlsToShow = kOrdWizardDefaults
                    m.Quantity.Price = m.Quantity.Min
                Else
                    WizardGridClear
                    strCtrlsToShow = m.Chart.OrdBarCtrls
                    m.eOrdBarMode = eOrdBarMode_Order
                End If
            End If
        End If
    End If
        
    If vseOrderBar.Visible Then
        If m.eOrdBarMode = eOrdBarMode_Order Or m.eOrdBarMode = eOrdBarMode_Wizard Then
            bShowExchanges = g.Broker.ShowExchangeControls(m.Chart.TradeAccountID, m.Chart.SymbolID)
        End If
        
        fraOrderBtns.Width = vseOrderBar.Width
        fraFrontMonth.Width = vseOrderBar.Width
        PopulateRiskGrapType
        'populate combo box when update flag is true or when streaming replay gets turned on and off
        If bUpdateAccts Or bPrevReplayVisible <> frmReplay.Visible Or cboAccounts.ListCount <= 0 Then       '4997
            bPrevReplayVisible = frmReplay.Visible      'this fixes infinite loop caused by cbo click event re-calling this
            PopulateAccountsCbo cboAccounts, m.Chart.TradeAccountID, True
        End If
        If cboAccounts.ListCount > 0 Then
            If g.nReplaySession > 0 Or frmReplay.Visible Then
                cboAccounts.Enabled = False
                If cboAccounts.ItemData(cboAccounts.ListIndex) <> g.nReplayAccountID Then
                    If g.nReplayAccountID <> 0 Then
                        'if user already has order bar open and starts downloading data for streaming replay,
                        'the replay account is not always set yet during downloading so don't want to do this
                        InfBox "Account " & Str(cboAccounts.ItemData(cboAccounts.ListIndex)) & " does not match Replay Account " & Str(g.nReplayAccountID), "E", "", "Account Error"
                        m.Chart.ShowTrades = 0
                        vseOrderBar.Visible = False
                        FormResize Me
                    End If
                    Exit Sub
'JM 03-16-2015: we are streaming replay a limited number of stocks
'               allow order bar for all and if it is not in list then nothing happens
'                ElseIf SecurityType(m.Chart.Symbol) = "S" Then
'                    m.Chart.ShowTrades = 0
'                    vseOrderBar.Visible = False
'                    FormResize Me
                Else
                    m.Chart.SetTrackerTradesReload
                End If
            Else
                cboAccounts.Enabled = True
                If cboAccounts.ItemData(cboAccounts.ListIndex) <> m.Chart.TradeAccountID Then
                    m.Chart.TradeAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
                End If
            End If
        Else
            InfBox "Unable to obtain trading account information.", "!", "", "Account Error"
            m.Chart.ShowTrades = 0
            vseOrderBar.Visible = False
            FormResize Me
            Exit Sub
        End If
        
        If cmdRollNow.Visible Then
            fraOrderBarMode.Visible = False
            fraOrdWizard.Visible = False
            SetFlexVisible eFlexGridIdx_OrdWizard, False
            lblAccounts.Move 60, 0
            cboAccounts.Move 60, lblAccounts.Top + lblAccounts.Height + 50
            
            If m.eOrdBarMode = eOrdBarMode_BrokerDisconnect Then        '6688
                With fraOrderBtns
                    fraBrokerDisconnect.Move .Left, .Top + lblOpenEquity.Height + lblTradePos.Height + 150, .Width, .Height
                End With
                fraBrokerDisconnect.Visible = True
            Else
                fraBrokerDisconnect.Visible = False
            End If
        ElseIf bPositionCtrls Then
            If m.bAllowOptWizard And m.eOrdBarMode <> eOrdBarMode_BrokerDisconnect Then
                fraOrdWizard.Width = vseOrderBar.Width
                fraOrdWizard.Left = vseOrderBar.Left - 10
                 
                If m.eOrdBarMode = eOrdBarMode_Wizard Then
                    If m.bFlexOrdBar Then
                        With fgChartFlex(eFlexGridIdx_OrdWizard)
                            .Width = fraOrdWizard.Width - 15
                            .Left = fraOrdWizard.Left
                            If Me.WindowState = vbMaximized Then
                                .Top = fraOrdWizard.Top + lblMultiLeg.Top + lblMultiLeg.Height * 2 + 60
                            Else
                                .Top = fraOrdWizard.Top + lblMultiLeg.Top + lblMultiLeg.Height
                            End If
                            
                            If fgChartFlex(eFlexGridIdx_AcctBar).Visible Then
                                .Top = .Top + fgChartFlex(eFlexGridIdx_AcctBar).Height + 30
                            End If
                        
                            .Visible = True
                            .ZOrder
                        End With
                    End If
                End If
                
                If bShowExchanges Then
                    lblExchange.Move 60, fraOrderBarMode.Top + fraOrderBarMode.Height + 15
                    cboExchanges.Move 60, lblExchange.Top + lblExchange.Height
                    lblAccounts.Move 60, cboExchanges.Top + cboExchanges.Height + 60
                    cboAccounts.Move 60, lblAccounts.Top + lblAccounts.Height
                Else
                    lblAccounts.Move 60, fraOrderBarMode.Top + fraOrderBarMode.Height + 100
                    cboAccounts.Move 60, lblAccounts.Top + lblAccounts.Height + 50
                End If
                
                fraOrderBarMode.Visible = True
                pbRight.Visible = True
                
                'center buttons in frame
                lblMultiLeg.Move 0, lblMultiLeg.Top, fraOrdWizard.Width
            Else
                fraOrderBarMode.Visible = False
                pbRight.Visible = False
                
                If bShowExchanges Then
                    lblExchange.Move 60, 0
                    cboExchanges.Move 60, lblExchange.Top + lblExchange.Height
                    lblAccounts.Move 60, cboExchanges.Top + cboExchanges.Height + 60
                    cboAccounts.Move 60, lblAccounts.Top + lblAccounts.Height
                Else
                    lblAccounts.Move 60, 0
                    cboAccounts.Move 60, lblAccounts.Top + lblAccounts.Height + 50
                End If
            End If
            
            lblExchange.Visible = bShowExchanges
            cboExchanges.Visible = bShowExchanges
            
            fraOrderBtns.Move 0, cboAccounts.Top + cboAccounts.Height + 100
            'when called externally, locking prevents order bar flashing
            bLocked = LockWindowUpdate(vseOrderBar.hWnd)
            
            Set Ctrl2 = NextOrdBarCtrl(iCtlIndex)
            While Not Ctrl2 Is Nothing
                If Ctrl2.Visible Then
                    PositionOrdBarCtrl Ctrl1, Ctrl2
                    Set Ctrl1 = Ctrl2
                End If
                
                If Not Ctrl1 Is Nothing Then
                    If Ctrl1.Name = "chkAutoExit" Then
                        Set Ctrl2 = lblAutoExit
                    Else
                        iCtlIndex = iCtlIndex + 1
                        Set Ctrl2 = NextOrdBarCtrl(iCtlIndex)
                    End If
                Else
                    iCtlIndex = iCtlIndex + 1
                    Set Ctrl2 = NextOrdBarCtrl(iCtlIndex)
                End If
            Wend
            
            'Development Note (10-05-2010):
            '   If it becomes necesarry to make this visible always then move the rithimic frame
            '   OUT OF the order bar buttons frame and resize in the CheckOrdBarColor routine.
            If Not Ctrl1 Is Nothing Then
                fraRithmic.Top = Ctrl1.Top + Ctrl1.Height + 90
                fraRithmic.Left = Ctrl1.Left
            End If
            
            'JM 01-17-2012 - primatry differences between Options Explorer enablement or not are:
            'Not enabled for OE: the little blue left/right buttons & fraOrderBarMode are never shown
            '   do not need to mess with turning these guys on/off & repositioning them
            'Enabled for OE: the little blue left/right buttons & fraOrderBarMode are shown only when connected to broker
            '   must reposition them relative to the account dropdown combo etc on every broker connect disconnect
            
            If m.bAllowOptWizard Then           '5047
                If m.eOrdBarMode = eOrdBarMode_Order Or m.eOrdBarMode = eOrdBarMode_BrokerDisconnect Then
                    lblOrderBarMode.Caption = OrdBarModeCaption
                    fraOrdWizard.Visible = False
                    SetFlexVisible eFlexGridIdx_OrdWizard, False
                    m.Chart.RemoveAnnots False, , eANNOT_OptionInfo
                    If m.Chart.ShowSplitPane = 1 Then
                        Set Pane = m.Chart.Tree("PRICE PANE")
                        If Pane Is Nothing Then
                            m.Chart.ShowSplitPane = 0
                        ElseIf Pane.SplitPaneType = ePANE_SplitPaneOptGraph Then
                            OptNavGraphClear
                        End If
                    End If
                    
                    'check disconnect
                    If m.eOrdBarMode = eOrdBarMode_BrokerDisconnect Then
                        With fraOrderBtns
                            fraBrokerDisconnect.Move .Left, .Top + lblOpenEquity.Height + lblTradePos.Height + 150, .Width, .Height
                        End With
                        fraBrokerDisconnect.Visible = True
                        fraOrderBarMode.Visible = False
                        pbRight.Visible = False
                        pbLeft.Visible = False
                    Else
                        fraBrokerDisconnect.Visible = False
                        fraOrderBarMode.Visible = True
                    End If
                    
                    m.Chart.GenerateChart eRedo1_Scrolled
                Else
                    lblOrderBarMode.Caption = OrdBarModeCaption
                    fraOrdWizard.Visible = True
                    SetFlexVisible eFlexGridIdx_OrdWizard, True
                    fraOrdWizard.Move 0, fraOrderBtns.Top + cmdClearQty.Top + cmdClearQty.Height + 150, fraOrderBtns.Width
                    m.Chart.GenerateChart eRedo1_Scrolled
                    If bNewChain Then OptionAnnotOnScreen
                    If m.Chart.ShowSplitPane = 1 Then
                        Set Pane = m.Chart.Tree("PRICE PANE")
                        If Pane Is Nothing Then
                            m.Chart.ShowSplitPane = 0
                        ElseIf Pane.SplitPaneType = ePANE_SplitPaneTimer Then
                            'just clear here, splitpane type will be set correctly when risk graph data received
                            Pane.SplitPaneType = ePANE_SplitPaneNone
                            m.Chart.ShowSplitPane = 0
                        End If
                    End If
                End If
            Else
                'when you do not have options explorer enablement, there is a lot less to do,
                'position of disconnect frame is slightly different & no need to call generatechart
                
                If lblQty.Visible Then lblQty.Visible = False
                
                If m.eOrdBarMode = eOrdBarMode_BrokerDisconnect Then        'aardvark 6568
                    With fraOrderBtns
                        fraBrokerDisconnect.Move .Left, .Top + lblOpenEquity.Height + lblTradePos.Height + 90, .Width, .Height
                    End With
                    fraBrokerDisconnect.Visible = True
                Else
                    fraBrokerDisconnect.Visible = False
                End If
            End If
            
            If m.eOrdBarMode = eOrdBarMode_Order And m.bResetOptWizardSpace Then
                m.Chart.RestoreChartNormal vbKeyReturn      '5121
                m.bResetOptWizardSpace = False
            End If
                                    
            If bLocked Then LockWindowUpdate (0)
        End If
    ElseIf m.bResetOptWizardSpace Then
        m.Chart.RestoreChartNormal vbKeyReturn
        m.bResetOptWizardSpace = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".FixOrderBarControls"

End Sub

Private Sub UpdateTradeQuantity(Button As Integer, ByVal lQuantity As Long)
On Error GoTo ErrSection:
    
    If Button = vbRightButton Then
        m.Quantity.Price = m.Quantity.Price - lQuantity
    Else
        m.Quantity.Price = m.Quantity.Price + lQuantity
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".UpdateTradeQuantity"

End Sub

Public Sub ToggleOrderBar(ByVal bVisible As Boolean, Optional ByVal bReset As Boolean = False, _
    Optional ByVal nSymbolPitType As Long = -1)
On Error GoTo ErrSection:
        
    Static nPrevSymbolID&
        
    Dim i&
    Dim bAllowTrade As Boolean
    Dim bAllowTradePit As Boolean
    
    Dim strSymbol As String
    
    Dim aOpenPos As cGdArray
    Dim aOpenOrders As cGdArray
    
    Dim eSymbolPitType As eFutureSymbolType
        
    If g.bUnloading Or g.bStarting Then
        Exit Sub
    ElseIf m.Chart Is Nothing Then
        Exit Sub
    ElseIf Len(m.Chart.SpreadSymbols) > 0 Then
        bVisible = False        '4310
    ElseIf m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then
        Exit Sub
    End If
    
    If m.eOrdBarMode <> eOrdBarMode_PFP And m.eOrdBarMode <> eOrdBarMode_Undefined Then
        If fraPatternProfit.Visible Then
            fraPatternProfit.Visible = False
            SetFlexVisible eFlexGridIdx_PfpInd, False
            SetFlexVisible eFlexGridIdx_PfpHits, False
            bReset = True
        End If
        If nSymbolPitType = -1 Then
            eSymbolPitType = m.Chart.SymbolPitType
        Else
            eSymbolPitType = nSymbolPitType
        End If
        
        If eSymbolPitType = ePitSymbol And TypeOfAccount(TradeAccountID) = eGDTypeOfAccount_Simulated Then
            bAllowTradePit = True
        End If
    End If
    
    If m.bGameMode Then     'Or InStr(m.Chart.Symbol, "-0") <> 0 Or Not g.RealTime.Active Then
        'don't show if in game mode or symbol is continuous contract
        If m.epbCursor = eCursor_OrderBuy Or m.epbCursor = eCursor_OrderSell Then
            m.epbCursor = eCursor_CrossHair
        End If
        If vseOrderBar.Visible Then
            vseOrderBar.Visible = False
            Form_Resize
            m.Chart.geForceRecalc
        End If
    ElseIf m.eOrdBarMode = eOrdBarMode_PFP Then
        If bReset Or Not vseOrderBar.Visible Then
            vseOrderBar.Visible = True
            If frmMain.WindowState <> vbMinimized Then
                Form_Resize
                'do this all the time so don't have to wait for real time
                m.Chart.geDrawChart     'fixes blank bars not showing when chart first shows
            End If
        End If
    ElseIf bReset Or (vseOrderBar.Visible <> bVisible) Or (m.Chart.SymbolID <> nPrevSymbolID) Then

        vseOrderBar.Visible = bVisible
        strSymbol = m.Chart.Symbol
        
        'aardvark 3032 fix (desktop flashes continually when realtime stream on)
        'explanation: this routine is called from generate chart everytime
        'only do the following if one of following is true:
        'a. user switched order bar from on to off or vice-versa
        'b. user changed symbol (S hotkey) from continuous to non-continuous or vice versa
        If bVisible Then
            If eSymbolPitType = eCombinedSymbol Or (eSymbolPitType = ePitSymbol And Not bAllowTradePit) Then
                cmdContracts.Caption = ConvertFutureSymbol(m.Chart.Symbol, eElectronicSymbol)
                cmdRollNow.Visible = False
                If Len(cmdContracts.Caption) > 0 Then
                    cmdContracts.Visible = True
                    lblOpenOrderPos.Caption = m.Chart.Symbol & " cannot be traded on this bar." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Go to"
                Else
                    cmdContracts.Visible = False        '5668
                    lblOpenOrderPos.Caption = m.Chart.Symbol & " cannot be traded on this bar."
                End If
            ElseIf InStr(strSymbol, "-0") Then
                If Not m.bTradeContinuous Then
                    bAllowTrade = False
                Else
                    'JM 01-15-2008: New code for trading continuous contract.
                    Set aOpenPos = New cGdArray
                    Set aOpenOrders = New cGdArray
                    bAllowTrade = AllowTradeContinuous(strSymbol, m.Chart.TradeAccountID, aOpenPos, aOpenOrders)
                    cmdContracts.Caption = "Go to Contract"
                    cmdRollNow.Visible = True
                    cmdContracts.Visible = True
                    lblOpenOrderPos.Caption = "You have  positions and/or orders in other contracts."
                End If
            ElseIf InStr(strSymbol, "$") <> 0 And Not IsForex(strSymbol) Then
                If m.eOrdBarMode = eOrdBarMode_Order Then
                    bAllowTrade = False
                    lblOpenOrderPos.Caption = "Indices cannot be traded on this bar."
                    cmdRollNow.Visible = False
                    cmdContracts.Visible = False
                Else
                    bAllowTrade = True
                End If
            Else
                bAllowTrade = True
            End If
            
            If bAllowTrade Then
                fraFrontMonth.Visible = False
                fraOrderBtns.Visible = True
            Else
                fraFrontMonth.Visible = True
                fraOrderBtns.Visible = False
                pbRight.Visible = False
                pbLeft.Visible = False
                
                If m.bTradeContinuous Then
                    If Not aOpenPos Is Nothing And Not aOpenOrders Is Nothing Then
                        'JM 01-15-2008: New code for trading continuous contract.
                        If aOpenPos.Size > 0 Or aOpenOrders.Size > 0 Then
                            PopulateMnuContracts aOpenPos, aOpenOrders
                        End If
                    End If
                End If
            End If
        End If
        
        'This is extra protection for the flashing issue 3032 in case the
        'ElseIf above somehow evaluates to true when application is minimized.
        'Theoretically that should not happen, but the protection does not hurt.
        If frmMain.WindowState <> vbMinimized Then
            Form_Resize
            'do this all the time so don't have to wait for real time
            m.Chart.geDrawChart     'fixes blank bars not showing when chart first shows
        End If
        
        nPrevSymbolID = m.Chart.SymbolID
        
    End If
        
    Set aOpenPos = Nothing
    Set aOpenOrders = Nothing
        
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ToggleOrderBar"
    
End Sub

Private Function SetOrderType(Order As cPtOrder, ByVal eClickOrder As enumOneClickOrder, _
    ByVal dPriceNew#, ByVal dPriceLast#) As Boolean
On Error GoTo ErrSection:

    Dim bWizardBuy As Boolean
    Dim bWizardSell As Boolean
    
    Dim strAnswer As String
    Dim X&, Y&
    
    If vseOrderBar.Visible And m.eOrdBarMode = eOrdBarMode_Wizard Then
        If pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderBuy")) Or pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderBuyWhite")) Then
            bWizardBuy = True
        ElseIf pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderSell")) Or pbChart.MouseIcon = Picture16(ToolbarIcon("kOrderSellWhite")) Then
            bWizardSell = True
        End If
    End If
            
    If eClickOrder = eClickOrder_None Then
        If vseBracketOrder.Appearance = apInset Then
            GetPromptLocation X, Y
            strAnswer = InfBox("Would you like to buy or sell?", "?", "Buy|Sell|Cancel", "Bracket Order", , , , , , , , , , , X, Y)
            If strAnswer = "B" Then
                m.epbCursor = eCursor_OrderBuy
            ElseIf strAnswer = "S" Then
                m.epbCursor = eCursor_OrderSell
            Else
                SetOrderType = False
                GoTo ErrExit
            End If
        End If
        If cboOrderType.Text = "Auto" Then
            If m.Chart.Bars.PriceDisplay(dPriceNew) = m.Chart.Bars.PriceDisplay(dPriceLast) Then
                Order.OrderType = eTT_OrderType_Limit
                If m.epbCursor = eCursor_OrderBuy Then
                    Order.Buy = True
                ElseIf m.epbCursor = eCursor_OrderSell Then
                    Order.Buy = False
                End If
            ElseIf dPriceNew > dPriceLast Then
                If m.epbCursor = eCursor_OrderBuy Or bWizardBuy Then
                    Order.OrderType = eTT_OrderType_Stop
                    Order.Buy = True
                ElseIf m.epbCursor = eCursor_OrderSell Or bWizardSell Then
                    Order.OrderType = eTT_OrderType_Limit
                    Order.Buy = False
                End If
            ElseIf dPriceNew < dPriceLast Then
                If m.epbCursor = eCursor_OrderBuy Or bWizardBuy Then
                    Order.OrderType = eTT_OrderType_Limit
                    Order.Buy = True
                ElseIf m.epbCursor = eCursor_OrderSell Or bWizardSell Then
                    Order.OrderType = eTT_OrderType_Stop
                    Order.Buy = False
                End If
            End If
        ElseIf cboOrderType.Text = "MIT" Then
            If m.epbCursor = eCursor_OrderBuy Then
                Order.Buy = True
                Order.OrderType = eTT_OrderType_MIT
            Else
                Order.Buy = False
                Order.OrderType = eTT_OrderType_MIT
            End If
        Else
            If m.epbCursor = eCursor_OrderBuy Then
                Order.Buy = True
            Else
                Order.Buy = False
            End If
            If cboOrderType.Text = "Stop" Then
                Order.OrderType = eTT_OrderType_Stop
            Else
                Order.OrderType = eTT_OrderType_Limit
            End If
        End If
    End If
    
    SetOrderType = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".SetOrderType"

End Function

Private Sub CheckOrdBarColor()
On Error GoTo ErrSection:

    Static nPrevSymID&
    
    Dim nBackColor&, nReverseColor&, strPos$
    Dim bReverseEnable As Boolean
    Dim bRithmic As Boolean
    
    nBackColor = -1
    strPos = UCase(Parse(g.Broker.PositionString(m.Chart.TradeAccountID, m.Chart.SymbolID, 0&), "|", 1))
    
    If TypeOfAccount(TradeAccountID) <> eGDTypeOfAccount_BrokerLive Or g.nReplaySession > 0 Or frmReplay.Visible Then
        nBackColor = Me.BackColor
    End If
    
    Select Case strPos
        Case "LONG"
            If nBackColor = -1 Then nBackColor = kFrameLong
            nReverseColor = kFrameShort
            bReverseEnable = True
        Case "SHORT"
            If nBackColor = -1 Then nBackColor = kFrameShort
            nReverseColor = kFrameLong
            bReverseEnable = True
        Case Else
            If nBackColor = -1 Then nBackColor = kFrameLive
            nReverseColor = Me.BackColor
            bReverseEnable = False
    End Select

    If vseOrderBar.BackColor <> nBackColor Then
        'set frames color
'        fraAccount.BackColor = nBackColor
        fraOrderBtns.BackColor = nBackColor
        fraPrices.BackColor = nBackColor
        fraTSO.BackColor = nBackColor
        fraFrontMonth.BackColor = nBackColor
        lblOpenOrderPos.BackColor = nBackColor              'this is text label about cannot trade continuous contract
        chkConfirmOrder.BackColor = nBackColor
        vseOrderBar.BackColor = nBackColor
        chkAutoExit.BackColor = nBackColor
        lblAutoExit.BackColor = nBackColor
        chkAutoJournal.BackColor = nBackColor
        
        fraOrdWizard.BackColor = nBackColor             '5005
        fraOrderBarMode.BackColor = nBackColor
        fraBrokerDisconnect.BackColor = nBackColor
        lblBrokerDisconnect.BackColor = nBackColor
        
        fraRithmic.BackColor = nBackColor
        fraExitFavorites.BackColor = nBackColor
    End If
    
    If pbRight.BackColor <> nBackColor Then
        'JM 12-07-2015: picture boxes are skipped when fixing colors for TN themes
        pbRight.BackColor = nBackColor
        pbLeft.BackColor = nBackColor
    End If
    
    If cmdReverse.BackColor <> nReverseColor Then
        cmdReverse.BackColor = nReverseColor      '5900
        If cmdReverse.Enabled <> bReverseEnable Then cmdReverse.Enabled = bReverseEnable
    End If
    
    bRithmic = ShowRithmic(m.Chart.TradeAccountID)      'for testing, just set bRithmic to True/False
    If bRithmic Then
'Development Note (10-06-2010):
'   move the rithmic frame OUT OF the order buttons frame then uncomment
'   this code to position rithmic frame at bottom of order bar always
'        i = vseOrderBar.Height - fraRithmic.Height
'        If fraRithmic.Top <> i Or fraRithmic.Left <> vseOrderBar.Left Then
'            fraRithmic.Move vseOrderBar.Left, vseOrderBar.Height - fraRithmic.Height
'        End If

        If Not fraRithmic.Visible Then fraRithmic.Visible = True
    ElseIf fraRithmic.Visible Then
        fraRithmic.Visible = False
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".CheckOrdBarColor"
    
End Sub

Private Function NextOrdBarCtrl(ByVal iIndex As Long) As Control
On Error GoTo ErrSection:

    Dim aControls As New cGdArray
    
    If m.eOrdBarMode = eOrdBarMode_Order Then
        aControls.SplitFields m.Chart.OrdBarCtrls, "|"
    ElseIf m.eOrdBarMode = eOrdBarMode_BrokerDisconnect Then
        aControls.SplitFields kOrdBarDisconnect, "|"
    Else
        aControls.SplitFields kOrdWizardDefaults, "|"
    End If
    Set NextOrdBarCtrl = OrdBarCtrlFromCode(aControls, Me, iIndex)
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError Me.Name & ".NextOrdBarCtrl"

End Function

Private Sub PositionOrdBarCtrl(PrevCtrl As Control, CurrCtrl As Control)
On Error GoTo ErrSection:

    Dim nTop&, nLeft&, i&
    
    If CurrCtrl.Name = "cmdReverse" Or CurrCtrl.Name = "cmdCancelAll" Then
        nLeft = 75
    Else
        nLeft = 60
    End If
    
    If PrevCtrl Is Nothing Then
        nTop = 0
    ElseIf PrevCtrl.Name = "lblOrderType" Then
        nTop = PrevCtrl.Top + PrevCtrl.Height + 50
    Else
        nTop = PrevCtrl.Top + PrevCtrl.Height + 20
    End If
    
    If Not CurrCtrl Is Nothing Then
        CurrCtrl.Top = nTop
        CurrCtrl.Left = nLeft
        If CurrCtrl.Name = "txtTradeQty" Then
            vscrQty.Top = nTop
            cmdClearQty.Top = nTop
        
            txtTradeQty.Left = cmdClearQty.Left + cmdClearQty.Width
            vscrQty.Left = txtTradeQty.Left + txtTradeQty.Width
        ElseIf CurrCtrl.Name = "cmdQty1" Then
            cmdQty2.Top = nTop
            cmdQty3.Top = nTop
        
            cmdQty2.Left = cmdQty1.Left + cmdQty1.Width
            cmdQty3.Left = cmdQty2.Left + cmdQty2.Width
        ElseIf CurrCtrl.Name = "lblOrderType" Then
            CurrCtrl.Top = CurrCtrl.Top + 30
            cboOrderType.Top = nTop
            cboOrderType.Left = lblOrderType.Left + 510     'lblOrderType.Width
        ElseIf CurrCtrl.Name = "cmdClearQty" Then
            txtTradeQty.Top = nTop
            vscrQty.Top = nTop
            
            txtTradeQty.Left = cmdClearQty.Left + cmdClearQty.Width
            vscrQty.Left = txtTradeQty.Left + txtTradeQty.Width
        ElseIf CurrCtrl.Name = "chkAutoExit" And fraExitFavorites.Visible Then
            fraExitFavorites.Top = nTop
            fraExitFavorites.Left = nLeft
            chkAutoExit.Left = nLeft
            chkAutoExit.Top = nTop + fraExitFavorites.Height
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".PositionOrdBarCtrl"
    
End Sub

Public Property Get AcctBarHeader() As cGdArray
On Error GoTo ErrSection:

    Set AcctBarHeader = m.aABarColHeader

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".AcctBarHeaderGet"

End Property

Public Sub SetAcctBarHeader(ByVal strText$)
On Error GoTo ErrSection:

    GridBarHeader fgChartFlex(eFlexGridIdx_AcctBar), m.aABarColHeader, 0, strText
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetAcctBarHeader", eGDRaiseError_Raise

End Sub

Public Property Get TradeAccountID() As Long
On Error GoTo ErrSection:

    TradeAccountID = Chart.TradeAccountID

ErrExit:
    Exit Property

ErrSection:
    RaiseError Me.Name & ".TradeAccountIDGet", eGDRaiseError_Raise

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAutoExit
'' Description: Attempt to select the auto exit in the combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAutoExit()
On Error GoTo ErrSection:

    Dim strAutoExitName As String       ' Auto exit name to select
    
    m.bSettingAutoExit = True
    
    strAutoExitName = g.OrderStrategies.ExitForAccountAndSymbol(TradeAccountID, SymbolID)
    If Len(strAutoExitName) > 0 Then
        lblAutoExit.Caption = strAutoExitName
        chkAutoExit.Value = vbChecked
        
        ExitCtrlAppearance Me, Nothing, strAutoExitName
    Else
        lblAutoExit.Caption = "None"
        chkAutoExit.Value = vbUnchecked
        
        'JM 03-30-2011: the autoexit checkbox & label are temporarily disabled when user
        '   clicks an exit favorite button, do not need to do this right now
        If chkAutoExit.Enabled Then ExitCtrlAppearance Me, Nothing, ""
    End If
    
    m.bSettingAutoExit = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetAutoExit"
    
End Sub

Public Property Get OrderMoveInProg() As Long
On Error GoTo ErrSection:
    
    If MouseIsPressed() Then OrderMoveInProg = m.nActiveOrderID     '5722

    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".OrderMoveInProg.Get"

End Property

Public Sub ToggleAccountBar()
On Error Resume Next
    
    mnuAccountBar_Click

End Sub

Public Function GetDetachedPlacement() As String
On Error GoTo ErrSection:
    
    GetDetachedPlacement = m.strDetachedPlacment

ErrExit:
    Exit Function

ErrSection:
    RaiseError Me.Name & ".GetDetachedPlacement"

End Function

Public Function GetNormalPlacement() As String
On Error GoTo ErrSection:

    GetNormalPlacement = m.strNormalPlacement

ErrExit:
    Exit Function

ErrSection:
    RaiseError Me.Name & ".GetNormalPlacement"

End Function

Public Function GetRatioPlacement() As String
On Error GoTo ErrSection:
    
    GetRatioPlacement = m.strRatioPlacement

ErrExit:
    Exit Function

ErrSection:
    RaiseError Me.Name & ".GetRatioPlacement"

End Function

Public Sub SetDetachedPlacement(ByVal strPlacement As String)
On Error GoTo ErrSection:

    m.strDetachedPlacment = strPlacement

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".SetDetachedPlacement"

End Sub

Public Sub SetNormalPlacement(ByVal strPlacement As String)
On Error GoTo ErrSection:

    m.strNormalPlacement = strPlacement

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".SetNormalPlacement"

End Sub

Public Sub SetRatioPlacement(ByVal strRatioPlacement As String, Optional ByVal bOnlySize As Boolean = False)
On Error Resume Next ' (just trying to be nice here -- don't want any errors being displayed)
    
    Dim d#, l&, t&, w&, h&
    m.strRatioPlacement = strRatioPlacement
    If frmMain.WindowState <> vbMinimized And InStr(m.strRatioPlacement, ";") > 0 Then
        If InStr(m.strRatioPlacement, ".") = 0 Then
            ' stored as fixed # of twips - this is for backwards compatibility
            ' the incoming string is called strRatioPlacement, but it could be
            ' either ratio values or twips, the check for the decimal point
            ' determines which type of values is in the string
            l = Val(Parse(m.strRatioPlacement, ";", 1))
            t = Val(Parse(m.strRatioPlacement, ";", 2))
            w = Val(Parse(m.strRatioPlacement, ";", 3))
            h = Val(Parse(m.strRatioPlacement, ";", 4))
        ElseIf g.ChartGlobals.bChartModeAutoSize Then
            'set chart size/location as ratio of MDI client area
            d = Val(Parse(m.strRatioPlacement, ";", 1))
            l = Int(d * frmMain.ScaleWidth + 0.5)
            d = Val(Parse(m.strRatioPlacement, ";", 2))
            t = Int(d * frmMain.ScaleHeight + 0.5)
            d = Val(Parse(m.strRatioPlacement, ";", 3))
            w = Int(d * frmMain.ScaleWidth + 0.5)
            d = Val(Parse(m.strRatioPlacement, ";", 4))
            h = Int(d * frmMain.ScaleHeight + 0.5)
        Else
            'set chart size/location as fixed # of twips (we ignore the incoming strRatioPlacement)
            l = Val(Parse(m.strNormalPlacement, ";", 1))
            t = Val(Parse(m.strNormalPlacement, ";", 2))
            w = Val(Parse(m.strNormalPlacement, ";", 3))
            h = Val(Parse(m.strNormalPlacement, ";", 4))
        End If
        
        If Not g.bStarting Then
            ' make sure chart is fully on-screen
            If w <= 0 Or w > frmMain.ScaleWidth Then
                If g.ChartGlobals.bChartModeAutoSize Then
                    w = frmMain.ScaleWidth
                End If
            End If
            If h <= 0 Or h > frmMain.ScaleHeight Then
                If g.ChartGlobals.bChartModeAutoSize Then
                    h = frmMain.ScaleHeight
                End If
            End If
            If l > frmMain.ScaleWidth - w Then
                l = frmMain.ScaleWidth - w
            End If
            If t > frmMain.ScaleHeight - h Then
                t = frmMain.ScaleHeight - h
            End If
            If l < 0 Then l = 0
            If t < 0 Then t = 0
        End If
        
        ' move chart (if different)
        If l <> Me.Left Or t <> Me.Top Or w <> Me.Width Or h <> Me.Height Then
            If bOnlySize Then
                Me.Move Me.Left, Me.Top, w, h
            Else
                Me.Move l, t, w, h
            End If
        End If
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetExchanges
'' Description: Load the exchange combo and set the default exchange
'' Inputs:      Exchange List, Default Exchange
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetExchanges(ByVal strExchangeList As String, ByVal strDefaultExchange As String)
On Error GoTo ErrSection:

    Dim astrExchanges As cGdArray       ' Array of exchanges
    Dim lIndex As Long                  ' Index into a for loop
    
    ' Split out the exchange information into an array...
    Set astrExchanges = New cGdArray
    astrExchanges.SplitFields strExchangeList, ","
    
    With cboExchanges
        .Clear
        
        ' Load up combo box with the comma delimited list of exchanges in strExchangeList...
        For lIndex = 0 To astrExchanges.Size - 1
            If Len(astrExchanges(lIndex)) > 0 Then
                cboExchanges.AddItem astrExchanges(lIndex)
            End If
        Next lIndex
        
        ' Set the combo box to the string passed in strDefaultExchange if it exists...
        If Len(strDefaultExchange) > 0 Then
            For lIndex = 0 To cboExchanges.ListCount - 1
                If cboExchanges.List(lIndex) = strDefaultExchange Then
                    cboExchanges.ListIndex = lIndex
                    Exit For
                End If
            Next lIndex
        End If
    End With
    
    InfBox ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetExchanges"
    
End Sub

Private Sub HandleWhatIf(Annot As cAnnotation, Optional ByVal bMouseUp As Boolean = False)
On Error Resume Next
    
    Dim Bars As cGdBars
    Dim Ind As cIndicator
    Dim dY#, iLastBar&
    
    Set Bars = m.Chart.Bars
    
    If Not Bars Is Nothing Then
        dY = RoundToMinMove(m.MouseLast.dY, Bars.MinMove)
        iLastBar = m.Chart.LastGoodDataBar(False)
                        
        If bMouseUp Then
            If tmr.Tag = "WhatIfMoved" Then tmr.Tag = ""
            m.Chart.GenerateChart eRedo5_RecalcInd
        Else
            Set Ind = m.Chart.Tree("PRICE")
            If Not Ind Is Nothing Then
                If Ind.WhatIfMove(dY) Then
                    Annot.Y(1) = dY
                    Annot.Y(2) = dY
                    tmr.Tag = "WhatIfMoved"
                ElseIf tmr.Tag = "WhatIfMoved" Then
                   tmr.Tag = ""
                End If
            End If
            Annot.geDrawAnn m.Chart
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContractInformation
'' Description: Get the contract information (if applicable) for given symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetContractInformation()
On Error GoTo ErrSection:

    If (m.Chart.TradeAccountID > 0) And (m.Chart.TradeSymbolID > 0) Then
        g.Broker.GetContractInfo g.Broker.AccountTypeForID(m.Chart.TradeAccountID), m.Chart.Symbol, True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".GetContractInformation"
    
End Sub

Private Function AllowTradeContinuous(ByVal strSymbol$, ByVal nTradeAccount&, _
    Optional aPositions As cGdArray = Nothing, _
    Optional aOrders As cGdArray = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim i&
    Dim iPositionsFound&, iOrdersFound&

    Dim nSymbolID&, nCurrContractID&, strCurrContract$
    Dim strContract$, strContractPos$, strContractOrder$
        
    Dim oPosCollection As cAccountPositions
    Dim oPosition As cAccountPosition
    
    Dim oOrderCollection As cPtOrders
    Dim oOrder As cPtOrder
                
    If InStr(strSymbol, "-") = 0 Then Exit Function       'invalid symbol of some sort
    
    nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
    If nSymbolID = 0 Then Exit Function
        
    strCurrContract = RollSymbolForDate(strSymbol)         '5167
    strCurrContract = ConvertToTradeSymbol(strCurrContract)
    nCurrContractID = g.SymbolPool.SymbolIDforSymbol(strCurrContract)
    
    If Not aOrders Is Nothing Then aOrders.Size = 0
    If Not aPositions Is Nothing Then aPositions.Size = 0
    
    Set oPosCollection = g.Broker.OpenPositionsForBaseSymbol(nTradeAccount, strSymbol, 0)
    
    If Not oPosCollection Is Nothing Then
        For i = 0 To oPosCollection.Count
            Set oPosition = oPosCollection(i)
            If Not oPosition Is Nothing Then
                strContract = oPosition.Symbol
                strContractPos = oPosition.CurrentPositionSnapshotString
                If strContract <> strCurrContract Then iPositionsFound = iPositionsFound + 1
                If Not aPositions Is Nothing Then
                    aPositions.Add Str(oPosition.AccountID) & ";" & strContract & ";" & strContractPos
                End If
            End If
        Next
    End If
    
'    If iPositionsFound = 0 Then
        Set oOrderCollection = g.Broker.WorkingOrdersForBaseSymbol(nTradeAccount, strSymbol, 0)
        If Not oOrderCollection Is Nothing Then
            For i = 0 To oOrderCollection.Count
                Set oOrder = oOrderCollection(i)
                If Not oOrder Is Nothing Then
                    strContract = oOrder.Symbol
                    strContractOrder = oOrder.OrderText
                    If strContract <> strCurrContract Then iOrdersFound = iOrdersFound + 1
                    If Not aOrders Is Nothing Then
                        aOrders.Add Str(oOrder.AccountID) & ";" & strContract & ";" & strContractOrder
                    End If
                End If
            Next
        End If
'    End If
        
    If iPositionsFound > 0 Or iOrdersFound > 0 Then
        AllowTradeContinuous = False
    Else
        AllowTradeContinuous = True
    End If
        
ErrExit:
    Exit Function

ErrSection:
    RaiseError Me.Name & ".AllowTradeContinuous"

End Function

Private Sub PopulateMnuContracts(aPositions As cGdArray, aOrders As cGdArray)
On Error GoTo ErrSection:

    Dim i&, k&
    Dim strCurrContract$, strContract$, strDetail$
    Dim aContracts As New cGdArray
    Dim bCurrContractFound As Boolean
    
    If aPositions Is Nothing And aOrders Is Nothing Then
        Exit Sub            'precautionary, should never happen
    End If
        
    k = 0
    strCurrContract = RollSymbolForDate(m.Chart.Symbol, m.Chart.Bars(eBARS_DateTime, m.Chart.LastGoodDataBar(False)))
    For i = 0 To aPositions.Size - 1
        strContract = Parse(aPositions(i), ";", 2)
        strDetail = Parse(aPositions(i), ";", 3)
        If k > mnuContracts.UBound Then
            Load mnuContracts(k)
            mnuContracts(k).Visible = True
        End If
        mnuContracts(k).Caption = strContract & " (" & strDetail & ")"
        k = k + 1
        aContracts.Add strContract
    
        If strContract = strCurrContract Then bCurrContractFound = True
    Next
    
    For i = 0 To aOrders.Size - 1
        strContract = Parse(aOrders(i), ";", 2)
        strDetail = Parse(aOrders(i), ";", 3)
        If k > mnuContracts.UBound Then
            Load mnuContracts(k)
            mnuContracts(k).Visible = True
        End If
        mnuContracts(k).Caption = strContract & " (" & strDetail & ")"
        k = k + 1
        aContracts.Add strContract
    
        If strContract = strCurrContract Then bCurrContractFound = True
    Next
    
    If Not bCurrContractFound Then          '6598
        If k > mnuContracts.UBound Then
            Load mnuContracts(k)
            mnuContracts(k).Visible = True
        End If
        mnuContracts(k).Caption = strCurrContract & " (Flat)"
        aContracts.Add strCurrContract
    End If
    
    'clear out any items in list that is no longer valid (6688)
    If k <= mnuContracts.UBound Then
        For i = mnuContracts.UBound To k Step -1
            Unload mnuContracts(i)
        Next
    End If
    
    cmdRollNow.Enabled = AllowRollContinuous(aContracts)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".PopulateMnuContracts"
    
End Sub

Private Function AllowRollContinuous(Optional ByVal aContracts As cGdArray, Optional ByVal bRollNow As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim i&, k&
    Dim strCaption$, strContract$, strCurrContract$
    Dim aUnique As cGdArray, a1 As cGdArray, a2 As cGdArray
    Dim strSymbol As String
    
'Design note from Tim:
'For continuous contract chart (only -055, -057, -065, and -067 types):
'- if any type of continuous other than the above list, we just show a "Go to Contract" button
    strSymbol = m.Chart.Symbol
    
    If InStr(strSymbol, "-055") Or InStr(strSymbol, "-057") Or InStr(strSymbol, "-065") Or InStr(strSymbol, "-067") Then
        strCurrContract = RollSymbolForDate(m.Chart.Symbol)         '5167
        
        If aContracts Is Nothing Then
            Set aUnique = New cGdArray
            For i = mnuContracts.LBound To mnuContracts.UBound
                strCaption = mnuContracts(i).Caption
                strContract = Parse(strCaption, "(", 1)
                aUnique.Add strContract
            Next
        Else
            Set aUnique = aContracts
        End If
    
        aUnique.Sort eGdSort_Default Or eGdSort_DeleteDuplicates
        
        k = -1
        For i = 0 To aUnique.Size - 1
            If strCurrContract = aUnique(i) Then
                k = i
                Exit For
            End If
        Next
        
        If aUnique.Size > 2 Then
            AllowRollContinuous = False
        ElseIf bRollNow Then
            If k = 0 Then
                strContract = aUnique(1)
            Else
                strContract = aUnique(0)
            End If
            AllowRollContinuous = RollPosition(m.Chart.TradeAccountID, strContract, strCurrContract, 0)
        Else
            AllowRollContinuous = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".AllowRollContinuous"
    
End Function

Private Sub SetDnExHiLoFlag(MouseCoord As ChartCoordinates)
On Error GoTo ErrSection:

    Dim cInfo As coordinate_info
    Dim dyHi#, dyLow#, dDiffHi#, dDiffLow#, dyMouse#, nPeakBar&
    Dim rc&
    
    cInfo.paneId = -1
    cInfo.x_pixels = MouseCoord.MouseX / Screen.TwipsPerPixelX
    cInfo.y_pixels = MouseCoord.MouseY / Screen.TwipsPerPixelY
    rc = geCoordToData(m.Chart.geChartObj, cInfo)
    
    If rc <> 0 Then Exit Sub
            
    If cInfo.paneId = MouseCoord.nPaneID Then
        dyHi = m.Chart.Bars(eBARS_High, MouseCoord.nBar)
        dyLow = m.Chart.Bars(eBARS_Low, MouseCoord.nBar)

        dDiffHi = dyHi - cInfo.y_value
        If dDiffHi < 0 Then dDiffHi = dDiffHi * -1
        dDiffLow = dyLow - cInfo.y_value
        If dDiffLow < 0 Then dDiffLow = dDiffLow * -1
        If dDiffLow < dDiffHi Then
            m.nFocusHiLo = 2
        Else
            m.nFocusHiLo = 1
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetDnExHiLoFlag"
    
End Sub

Public Sub ClearChartObject()
On Error GoTo ErrSection:

    Set m.Chart = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".ClearChartObject"

End Sub

Public Sub SetChartObject(NewChart As cChart)
On Error GoTo ErrSection:
    
    If Not NewChart Is Nothing Then Set m.Chart = NewChart

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".SetChartObject"

End Sub

Public Property Get DetachStatus() As enumDetachStatus
On Error GoTo ErrSection:
    
    DetachStatus = m.eDetachStatus

ErrExit:
    Exit Property

ErrSection:
    RaiseError Me.Name & ".DetachStatusGet"

End Property

Public Property Let DetachStatus(ByVal eStatus As enumDetachStatus)
On Error GoTo ErrSection:
    
    m.eDetachStatus = eStatus
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError Me.Name & ".DetachStatusLet"

End Property

Public Sub CopyPlacements(Optional frm As Form = Nothing, _
    Optional ByVal strNormal$ = "", Optional ByVal strRatio$ = "", Optional ByVal strDetach$ = "")
On Error GoTo ErrSection:

    If frm Is Nothing Then
        If Len(strNormal) > 0 Then m.strNormalPlacement = strNormal
        If Len(strRatio) > 0 Then m.strRatioPlacement = strRatio
        If Len(strDetach) > 0 Then m.strDetachedPlacment = strDetach
    Else
        m.strDetachedPlacment = frm.GetDetachedPlacement
        m.strNormalPlacement = frm.GetNormalPlacement
        m.strRatioPlacement = frm.GetRatioPlacement
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".CopyPlacements"

End Sub

Public Function ToolBarWrapGet(ByVal strToolbar$) As Boolean
'returns width, height in pixels
On Error GoTo ErrSection:

    ToolBarWrapGet = m.bToolbarWrap

ErrExit:
    Exit Function

ErrSection:
    RaiseError Me.Name & ".ToolBarWrapGet"

End Function

Public Sub ToolBarWrapSet(ByVal strToolbar$, ByVal bWrap As Boolean)
'sets width, height in pixels
On Error GoTo ErrSection:

    m.bToolbarWrap = bWrap
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".ToolBarWrapSet"

End Sub

Public Property Get SkipFocusFix() As Boolean
On Error Resume Next
    
    SkipFocusFix = m.bSkipFocusFix

End Property

Public Property Let SkipFocusFix(ByVal bSkip As Boolean)
On Error Resume Next
    
    m.bSkipFocusFix = bSkip

End Property

Public Property Get OrderBarMode() As eOrderBarMode
On Error GoTo ErrSection:
    
    OrderBarMode = m.eOrdBarMode

ErrExit:
    Exit Property

ErrSection:
    RaiseError Me.Name & ".OrderBarModeGet"

End Property

Public Property Let OrderBarMode(ByVal eMode As eOrderBarMode)
On Error GoTo ErrSection:
    
    m.eOrdBarMode = eMode

ErrExit:
    Exit Property

ErrSection:
    RaiseError Me.Name & ".OrderBarModeLet"

End Property

Private Property Get OrdBarModeCaption() As String
On Error GoTo ErrSection:
    
    Dim strText$
    
    If m.eOrdBarMode = eOrdBarMode_Order Then
        strText = "Underlying"
        
        fraOrderBarMode.Move lblAccounts.Left
        pbRight.Move (lblAccounts.Left + lblAccounts.Width) - pbRight.Width + 15, fraOrderBarMode.Top + 150
        
        pbRight.Visible = True
        pbRight.Enabled = True
        pbLeft.Visible = False
        pbLeft.Enabled = False
        
        If lblQty.Visible Then lblQty.Visible = False
    Else
        strText = "Options"
    
        fraOrderBarMode.Move (lblAccounts.Left + lblAccounts.Width) - fraOrderBarMode.Width + 15
        pbLeft.Move lblAccounts.Left, fraOrderBarMode.Top + 150
    
        pbRight.Visible = False
        pbRight.Enabled = False
        pbLeft.Visible = True
        pbLeft.Enabled = True
        
        cmdClearQty.Move 75, cmdClearQty.Top + 365      '5102
        txtTradeQty.Move cmdClearQty.Left + cmdClearQty.Width, cmdClearQty.Top
        vscrQty.Move txtTradeQty.Left + txtTradeQty.Width, cmdClearQty.Top
        
        lblQty.Move 0, txtTradeQty.Top - lblQty.Height - 45
        If Not lblQty.Visible Then lblQty.Visible = True
    End If
    
    OrdBarModeCaption = strText

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".OrdBarModeCaptionGet"

End Property

Private Sub HandleWizardPrompt(ByVal X As Single, ByVal Y As Single)
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    Dim coordInfo As coordinate_info
    
    Dim rc&, dStrike#, dPrice#
    Dim strMonth$, strInfo$
    
    Dim strCaption$, nWidth&
    Dim nLabelLeft&, nLabelTop&, nBkColor&
    Dim nFrameLeft&, nFrameTop&
    Dim bVisible As Boolean
    Dim bBalloonTool As Boolean
        
    With lblWizardPrice
        strCaption = .Caption
        nLabelLeft = .Left
        nLabelTop = .Top
        nBkColor = .BackColor
    End With
    
    With fraWizardPrompt
        nFrameLeft = .Left
        nFrameTop = .Top
    End With
    
    If Not m.AnnotOptions Is Nothing Then
        Set Annot = m.Chart.ClosestOptionAnnot(m.MouseLast.dDate)
        If Not Annot Is Nothing Then bBalloonTool = Annot.IsBalloonOptionsInfo
    End If
        
    If m.eOrdBarMode = eOrdBarMode_Wizard And Not m.MouseLast.bOffChart Then
        'skip code to hide wizard prompt when working with balloon tool
        If Not bBalloonTool And (m.MouseLast.nBar > 0) And (m.MouseLast.nBar < m.Chart.LastGoodDataBar(False, False)) And g.ChartGlobals.eChartMode = eMode_ChartOrder Then
            If fraWizardPrompt.Visible Then fraWizardPrompt.Visible = False

            If InStr(m.Chart.Symbol, "$") <> 0 And Not IsForex(m.Chart.Symbol) Then
                bVisible = False
                ClearBuySellButtons True
            Else
                nLabelLeft = cmdCall.Left
                nWidth = cmdCall.Width + kPriceWizWidth

                If InStr(m.Chart.Symbol, "-") = 0 Then
                    strCaption = m.Chart.Symbol & vbCrLf & m.Chart.PriceDisplay(m.MouseLast.nPaneID, m.MouseLast.dY)
                Else
                    strCaption = RollSymbolForDate(m.Chart.Symbol, m.Chart.Bars(eBARS_DateTime, m.Chart.Bars.Size - 1)) & vbCrLf & m.Chart.PriceDisplay(m.MouseLast.nPaneID, m.MouseLast.dY)
                End If

                bVisible = True
                nFrameLeft = X - fraWizardPrompt.Width + pbChart.Left
                nFrameTop = Y - (fraWizardPrompt.Height / 2) + pbChart.Top
            End If
        ElseIf m.MouseLast.nBar < m.Chart.LastBar Then
            
            nLabelLeft = kPriceWizLeft
            nBkColor = vbCyan
            
            If bBalloonTool Then
                nWidth = kPriceWizWidth + cmdPut.Width
                cmdPut.Left = cmdCall.Left
'                cmdPut.Top = cmdCall.Top
            Else
                cmdPut.Left = kPriceWizLeft + kPriceWizWidth
                nWidth = kPriceWizWidth
            End If
            
            If Annot Is Nothing Then Set Annot = m.Chart.ClosestOptionAnnot(m.MouseLast.dDate)
            
            strCaption = ""
            If Not Annot Is Nothing Then
                If bBalloonTool Then
                    If m.MouseLast.dY > Annot.Y(2) Then
                        strInfo = Annot.ClosestStrike(m.Chart.Bars.RoundToPrice(m.MouseLast.dY), True)
                    Else
                        strInfo = Annot.ClosestStrike(m.Chart.Bars.RoundToPrice(m.MouseLast.dY), False, True)
                    End If
                Else
                    strInfo = Annot.ClosestStrike(m.Chart.Bars.RoundToPrice(m.MouseLast.dY))
                End If
                
                dStrike = ValOfText(Parse(strInfo, ";", 1))
                If dStrike > 0 Then
                    m.nActiveAnnotIdx = Annot.geAnnId
                    
                    If bBalloonTool Then
                        nFrameLeft = X
                        If m.eDetachStatus = eDetached Then
                            nFrameTop = Y
                        Else
                            nFrameTop = Y + cmdPut.Height / 2
                        End If
                        If dStrike < Annot.Y(2) Then
                            If m.AnnotOptions Is Nothing Then
                                bVisible = False
                            ElseIf m.AnnotOptions.BalloonPutStrike > 0 Then
                                bVisible = False
                            Else
                                bVisible = True
                                dPrice = ValOfText(Parse(strInfo, ";", 5))
                                cmdPut.ZOrder
                            End If
                        Else
                            If m.AnnotOptions Is Nothing Then
                                bVisible = False
                            ElseIf m.AnnotOptions.BalloonCallStrike > 0 Then
                                bVisible = False
                            Else
                                bVisible = True
                                dPrice = ValOfText(Parse(strInfo, ";", 4))
                                cmdCall.ZOrder
                            End If
                        End If
                        strCaption = Str(dStrike) & vbCrLf & Format(dPrice, "$#,##0.00")
                    Else
                        strMonth = DateFormat(Annot.dDate(2), MMM_YY)
                        strCaption = Parse(strMonth, "-", 1) & vbCrLf & Str(dStrike)
                    End If
                    
                    coordInfo.paneId = Annot.gePaneId
                    coordInfo.x_value = Annot.geLeftX(0)
                    coordInfo.y_value = dStrike
                    
                    rc = geDataToCoord(m.Chart.geChartObj, coordInfo)
                    If rc = 0 And coordInfo.paneId > 0 Then
                        nFrameLeft = (coordInfo.x_pixels * Screen.TwipsPerPixelX) - fraWizardPrompt.Width / 2 + pbChart.Left
                        If Not bBalloonTool Then
                            nFrameTop = (coordInfo.y_pixels * Screen.TwipsPerPixelY) - (fraWizardPrompt.Height / 2) + pbChart.Top
                        End If
                    Else
                        nFrameLeft = X + 50 'so won't show
                    End If
                    
                    If bBalloonTool Then
                        fraWizardPrompt.Visible = bVisible
                        fraWizardPrompt.Enabled = bVisible
                    ElseIf Annot.FirstOptionAnnnot And X < nFrameLeft + 30 Then         'aardvark 5961
                        bVisible = False
                    Else
                        bVisible = True
                        With fraWizardPrompt
                            If Not .Visible Then .Visible = True
                            If Not .Enabled Then .Enabled = True
                        End With
                    End If
                End If      'end strike > 0
            End If
        End If
    End If
    
    If Len(strCaption) = 0 Then bVisible = False
    
    If bVisible Then
        With lblWizardPrice
            If .Caption <> strCaption Then .Caption = strCaption
            If .Left <> nLabelLeft Then .Left = nLabelLeft
            If .Top <> nLabelTop Then .Top = nLabelTop
            If .Width <> nWidth Then .Width = nWidth
            If .BackColor <> nBkColor Then .BackColor = nBkColor
        End With
        With fraWizardPrompt
            If .Left <> nFrameLeft Then .Left = nFrameLeft
            If .Top <> nFrameTop Then .Top = nFrameTop
        End With
    ElseIf fraWizardPrompt.Visible Then
        With fraWizardPrompt
            .Visible = False
            .Enabled = False
            .Move -1000, -1000
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".HandleWizardPrompt"

End Sub

Private Function WizardGridDataRow() As Long
On Error GoTo ErrSection:

    Dim i&, j&
    
    If m.bFlexOrdBar Then
        With fgChartFlex(eFlexGridIdx_OrdWizard)
            For i = .FixedRows To .Rows - 1
                If Len(.TextMatrix(i, 1)) > 0 Then j = j + 1
            Next
        End With
    End If

    WizardGridDataRow = j

ErrExit:
    Exit Function

ErrSection:
    RaiseError Me.Name & ".WizardGridDataRow"

End Function

Private Function WizardGridMaxed() As Boolean
On Error GoTo ErrSection:

    Dim bMaxed As Boolean
    Dim strMsg$

    If m.bFlexOrdBar Then
        With fgChartFlex(eFlexGridIdx_OrdWizard)
            If Len(.TextMatrix(1, 1)) > 0 And _
               Len(.TextMatrix(2, 1)) > 0 And _
               Len(.TextMatrix(3, 1)) > 0 And _
               Len(.TextMatrix(4, 1)) > 0 Then
               
                bMaxed = True
               
            End If
        End With
    
        If bMaxed Then
            strMsg = "There is a max of 4 legs." & vbCrLf & "Would you like to replace the last leg?"
            If InfBox(strMsg, "?", "+Yes|-No", "Confirmation") = "Y" Then
                With fgChartFlex(eFlexGridIdx_OrdWizard)
                    fgChartFlex_BeforeEdit eFlexGridIdx_OrdWizard, .Rows - 1, 0, bMaxed
                End With
                bMaxed = False
            Else
                ClearBuySellButtons True
            End If
        End If
    End If
    
    WizardGridMaxed = bMaxed

ErrExit:
    Exit Function

ErrSection:
    RaiseError Me.Name & ".WizardGridMaxed"

End Function

Private Sub WizardGridClear()
On Error GoTo ErrSection:
    
    Dim i&
    
    If m.bFlexOrdBar Then
        With fgChartFlex(eFlexGridIdx_OrdWizard)
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpPicture, i, 0) = Nothing
                .TextMatrix(i, 1) = ""
                .TextMatrix(i, 2) = ""
                .TextMatrix(i, 3) = ""
                .TextMatrix(i, 4) = ""
                .Cell(flexcpBackColor, i, 1) = .Cell(flexcpBackColor, i, 0)
            Next
        End With
    End If
    
    OptNavGraphClear

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".WizardGridClear"

End Sub

Private Sub WizardGridAdd(ByVal eOrdType As enumOneClickOrder, ByVal strOrdInfo$, ByVal strOptSym$, _
    ByVal strOrderType$, OptionAnnot As cAnnotation, _
    Optional ByVal bUnderlying As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, strSym$, strPrice$, strQty$
    
    Dim bAdded As Boolean
    Dim bSubmitTicket As Boolean
    Dim bGraphOk As Boolean


    If vseBuyChart.Appearance = apInset Or vseSellChart.Appearance = apInset Then
        bSubmitTicket = True
    End If

    ClearBuySellButtons
    
    If Len(strOrdInfo) = 0 Then Exit Sub
    
    strSym = Parse(strOrdInfo, ";", 1)
    strPrice = Parse(strOrdInfo, ";", 2)
    strQty = Str(m.Quantity.Price)
        
    With fgChartFlex(eFlexGridIdx_OrdWizard)
        
        If bUnderlying And Len(.TextMatrix(.FixedRows, 1)) > 0 Then
            'clear out first row for underlying
            For i = .Rows - 1 To .FixedRows + 1 Step -1
                If Len(.TextMatrix(i - 1, 1)) > 0 Then
                    .TextMatrix(i, 1) = .TextMatrix(i - 1, 1)
                    .TextMatrix(i, 2) = .TextMatrix(i - 1, 2)
                    .TextMatrix(i, 3) = .TextMatrix(i - 1, 3)
                    .TextMatrix(i, 4) = .TextMatrix(i - 1, 4)
                    .Cell(flexcpPicture, i, 0) = Picture16("kCancel")
                    .Cell(flexcpBackColor, i, 1) = .Cell(flexcpBackColor, i - 1, 1)
                End If
            Next
            
            .TextMatrix(.FixedRows, 1) = strSym & " " & strPrice
            .TextMatrix(.FixedRows, 2) = strQty
            .TextMatrix(.FixedRows, 3) = strOptSym
            .TextMatrix(.FixedRows, 4) = strOrderType
            If eOrdType = eClickOrder_BuyMkt Or eOrdType = eClickOrder_BuyCall Or eOrdType = eClickOrder_BuyPut Then
                .Cell(flexcpBackColor, i, 1) = vseBuyWizard.BackColor
            Else
                .Cell(flexcpBackColor, i, 1) = vseSellWizard.BackColor
            End If
            bAdded = True
            i = .FixedRows
        Else
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, 1) = "" Then
                    If eOrdType = eClickOrder_BuyMkt Or eOrdType = eClickOrder_SellMkt Or InStr(strSym, "-") = 0 Then
                        .TextMatrix(i, 1) = strSym & " " & strPrice
                    Else
                        .TextMatrix(i, 1) = strSym
                    End If
                    .TextMatrix(i, 2) = strQty
                    .TextMatrix(i, 3) = strOptSym
                    .TextMatrix(i, 4) = strOrderType
                    .Cell(flexcpPicture, i, 0) = Picture16("kCancel")
                    .Cell(flexcpPictureAlignment, i, 0) = flexAlignCenterCenter
                    If eOrdType = eClickOrder_BuyMkt Or eOrdType = eClickOrder_BuyCall Or eOrdType = eClickOrder_BuyPut Then
                        .Cell(flexcpBackColor, i, 1) = vseBuyWizard.BackColor
                    Else
                        .Cell(flexcpBackColor, i, 1) = vseSellWizard.BackColor
                    End If
                    bAdded = True
                    Exit For
                End If
            Next
        End If
        
        If bAdded Then
            If Not OptionAnnot Is Nothing Then
                If OptionAnnot.dDate(1) - m.Chart.Bars(eBARS_DateTime, m.Chart.LastGoodDataBar(False)) < 2 Then
                    InfBox "There are less than 2 days to expiration. Risk graphs not available", "I", , "Options Order Bar"
                    cboRiskGraphType.ListIndex = 0
                ElseIf cboRiskGraphType.ListIndex = 0 Then
                    cboRiskGraphType.ListIndex = 1
                Else
                    cboRiskGraphType_Click
                End If
            End If
            
            If bSubmitTicket Then
                cmdTicket_Click
            Else
                .Row = i
                .Col = 2
                .Select .Row, .Col
                .EditCell
            End If
        End If
        
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".WizardGridAdd"

End Sub

Private Sub WizardControlsSync(ByVal nColor&)
On Error Resume Next

    cmdCall.BackColor = nColor
    cmdPut.BackColor = nColor

    m.eCrossHairOn = eCursor_Horizontal         'temporarily change to horizontal crosshair

End Sub

Public Sub OptionDataAvailable(ByVal eMsgType As eGDOptNavMessageType, ByVal strMsg$)
On Error Resume Next:
        
    Dim iHandle&, strErr$
        
    If Not m.Chart Is Nothing Then
        Select Case eMsgType
            Case eGDOptNav_ChainBuilt           '5109 - message comes up too slowly
                StatusMsg ""
                strErr = Parse(strMsg, vbTab, 2)
                If Len(strErr) > 0 Then
                    If Len(Parse(strErr, ";", 2)) > 0 Then
                        strErr = Parse(strErr, ";", 2)
                    End If
                    InfBox strErr, "!", , "Option Order Bar"
                Else
                    OrderBarModeToggle
                End If
            Case eGDOptNav_RiskGraphBuilt
                iHandle = ValOfText(Parse(strMsg, vbTab, 1))
                If iHandle = m.Chart.geChartObj Then
                    If Parse(strMsg, vbTab, 2) = "-1" Then
                        InfBox "There was an error retrieving risk graph data.", "I", , "Options Risk Graph"
                        OptNavGraphClear
                    Else
                        m.Chart.OptNavGraphInfoSet strMsg
                        m.Chart.GenerateChart eRedo1_Scrolled
                    End If
                End If
            Case eGDOptNav_TicketSubmitted
                iHandle = ValOfText(Parse(strMsg, vbTab, 1))
                If iHandle = m.Chart.geChartObj Then WizardGridClear
        End Select
    End If
                
End Sub

Private Sub OrderBarModeToggle()
On Error GoTo ErrSection:

    If g.bStarting Or g.bUnloading Then Exit Sub

    If m.eOrdBarMode = eOrdBarMode_Order Then
        InitChartFlex Me, eFlexGridIdx_OrdWizard
        m.eOrdBarMode = eOrdBarMode_Wizard
    Else
        m.eOrdBarMode = eOrdBarMode_Order
    End If
    
    ClearBuySellButtons True
    FixOrderBarControls False, True, , True
    ToggleOrderBar True, True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".OrderBarModeToggle"

End Sub

Private Sub OptionAnnotOnScreen()
On Error GoTo ErrSection:
        
    Dim dDateLastBar#, dDateLastScreenBar#, i&
    Dim Annot As cAnnotation
    
    Set Annot = m.Chart.ClosestOptionAnnot(dDateLastBar)
    
    If Annot Is Nothing Then Exit Sub
    
    i = m.Chart.LastGoodDataBar(False, True)
    dDateLastBar = m.Chart.Bars(eBARS_DateTime, i)
    
    i = gdGetSize(m.Chart.geDateArray) - 1
    dDateLastScreenBar = gdGetNum(m.Chart.geDateArray, i)
        
    If dDateLastScreenBar < Annot.dDate(1) Then
        i = Annot.dDate(1) - dDateLastScreenBar + 1
        m.Chart.ForecastBars(Me) = m.Chart.ForecastBars + i
        m.Chart.GenerateChart eRedo7_ReloadRT
        hsb.Value = hsb.Max
        m.bResetOptWizardSpace = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".OptionAnnotOnScreen"

End Sub

Public Property Get AllowDetach() As Boolean
On Error GoTo ErrSection:

    AllowDetach = True     'm.bAllowDetach

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".AllowDetachGet"

End Property

Private Sub PopulateRiskGrapType()
On Error GoTo ErrSection:

    If cboRiskGraphType.ListCount = 0 Then
        cboRiskGraphType.AddItem "None"
        cboRiskGraphType.AddItem "Profit/Loss"
        cboRiskGraphType.AddItem "Delta"
        cboRiskGraphType.AddItem "Gamma"
        cboRiskGraphType.AddItem "Theta"
        cboRiskGraphType.AddItem "Vega"
        cboRiskGraphType.AddItem "Probability"
        cboRiskGraphType.ListIndex = 1
    End If

EErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".PopulateRiskGrapType"

End Sub

Private Sub OptNavGraphClear(Optional ByVal nShow As Long = 0)
On Error GoTo ErrSection:

    Dim Chart As cChart
    Dim Pane As cPane
    
    If Not m.Chart Is Nothing Then
        If m.Chart.ShowSplitPane = 1 Then
            Set Pane = m.Chart.Tree("PRICE PANE")
            If Not Pane Is Nothing Then
                If Pane.SplitPaneType = ePANE_SplitPaneOptGraph Then
                    If nShow = 0 Then
                        Pane.OptNavGraphInfoClear
                        m.Chart.ShowSplitPane = nShow
                        m.Chart.geForceRecalc
                        m.Chart.GenerateChart eRedo1_Scrolled       '5246
                        Form_Resize
                    Else
                        Pane.OptNavGraphInfoClear
                        m.Chart.GenerateChart eRedo1_Scrolled
                    End If
                End If
            End If
        End If
    End If
    
'    If cboRiskGraphType.ListIndex <> 0 Then cboRiskGraphType.ListIndex = 0
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".OptNavGraphClear"

End Sub

Public Property Get TbButtonsArray(ByVal strToolbar$) As cGdArray
On Error GoTo ErrSection:

    If strToolbar = kTbDraw Then
        Set TbButtonsArray = m.aTbButtonsDraw
    Else
        Set TbButtonsArray = m.aTbButtons
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError Me.Name & ".TbButtonsArray"

End Property

Public Property Get OkayToRefresh() As Boolean
On Error Resume Next

    Dim bOkay As Boolean
    
    If tmr.Tag = "UNLOADING" Or tmr.Tag = "UNLOAD_NOW" Or tmr.Tag = "ToggleOrderbarMode" Or tmr.Tag = "DETACH_NOW" Then
        bOkay = False
    ElseIf m.eDetachStatus = eAttachInProg Or m.eDetachStatus = eDetachInProg Then
        bOkay = False
    Else
        bOkay = True
    End If
    
    OkayToRefresh = bOkay
    
End Property

Public Sub SyncDrawTools(Optional ByVal bSyncNow As Boolean = False)
On Error Resume Next:
    
    Static bInProgress As Boolean
    
    Dim i&
    Dim oButton As cPicBoxButton            '5154
    Dim aToolArray As cGdArray
    Dim strToolID As String
    
    Dim bSync As Boolean
    
    If Len(g.strActiveDraw) And FormIsLoaded("frmTbMoreButtons") Then Exit Sub
            
    StatusMsg ""
    
    If bInProgress Then Exit Sub
    
    If FormIsLoaded("frmElliot") Then
        If frmElliot.Visible Then Exit Sub      '4937
    End If
    
    If FormIsLoaded("frmIconAnnot") Then
        If frmIconAnnot.Visible Then Exit Sub
    End If
    
    bInProgress = True
    bSync = bSyncNow
    
    If InStr(g.strActiveDraw, "PFP") = 0 Then
        If m.eDetachStatus = eDetached And m.Chart.ShowToolbar = 1 Then
            Set oButton = ButtonByID(Me, "ID_RepeatDraw", kTbDraw)
        Else
            Set oButton = ButtonByID(frmMain, "ID_RepeatDraw", kTbDraw)
        End If
        
        If oButton Is Nothing Then
            bSync = True        '5201
        ElseIf oButton.BtnState = eBtnState_Neutral Then
            bSync = True
        ElseIf oButton.BtnState = eBtnState_Selected Then
            ClearAnnotFlags False   '7016
            bSync = False
        End If
    Else
        g.strActiveDraw = ""        '5807
        Set g.ChartGlobals.frmPfpSelPattern = Nothing
    End If
    
    If bSync Then
        If g.ChartGlobals.eChartMode = eMode_Move Then
            strToolID = "ID_ChartMove"
        ElseIf g.ChartGlobals.eChartMode = eMode_Zoom Then
            strToolID = "ID_ZoomIn"
        ElseIf g.ChartGlobals.eChartMode = eMode_Erase Then
            strToolID = "ID_Eraser"
        End If
    
        If m.eDetachStatus = eDetached And m.Chart.ShowToolbar = 1 Then
            Set oButton = ButtonByID(Me, strToolID, kTbDraw)
            If Not oButton Is Nothing Then
                i = oButton.PicboxIndex
                If i >= 0 Then
                    If g.vbeTbAlignDraw = vbAlignBottom And Me.pbTbBack(0).Visible Then
                        If i <= Me.pbTbBack.UBound Then oButton.MouseDown Me, Me.pbTbBack(i), m.aTbButtonsDraw, m.aTbButtonsDraw(m.aTbButtonsDraw.Size - 1)
                    ElseIf i <= Me.pbTbBackDraw.UBound Then
                        oButton.MouseDown Me, Me.pbTbBackDraw(i), m.aTbButtonsDraw, m.aTbButtonsDraw(m.aTbButtonsDraw.Size - 1)
                    End If
                End If
            End If
        Else
            Set oButton = ButtonByID(frmMain, strToolID, kTbDraw)
            If Not oButton Is Nothing Then
                i = oButton.PicboxIndex
                Set aToolArray = frmMain.TbButtonsArray(kTbDraw)
                If i >= 0 And Not aToolArray Is Nothing Then
                    If g.vbeTbAlignDraw = vbAlignBottom And frmMain.pbTbBack(0).Visible Then
                        If i <= frmMain.pbTbBack.UBound Then oButton.MouseDown frmMain, frmMain.pbTbBack(i), aToolArray, aToolArray(aToolArray.Size - 1)
                    ElseIf i <= frmMain.pbTbBackDraw.UBound Then
                        oButton.MouseDown frmMain, frmMain.pbTbBackDraw(i), aToolArray, aToolArray(aToolArray.Size - 1)
                    End If
                End If
            End If
        End If
    End If
    
    bInProgress = False

End Sub

Public Sub DrawToolSelect(frmFrom As frmTbMoreButtons)
On Error GoTo ErrSection:
    
    m.bDrawToolSelected = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".DrawToolSelect"

End Sub

Public Sub ChangeSymWrapper()
On Error GoTo ErrSection:
    'temporary fix for issue 5223
    
    Dim aStrings As New cGdArray
    Dim i&, nRec&
    
    If m.bGameMode And Not m.oGameMode Is Nothing Then
        'do nothing
    Else
        If m.eOrdBarMode = eOrdBarMode_Wizard Then OrderBarModeToggle
        If Len(m.Chart.SpreadSymbols) > 0 Then
            frmNewChart.ShowMe m.Chart.SpreadSymbols, True, m.Chart
        Else
            Set aStrings = frmSymbolSelector.ShowMe(m.Chart.Symbol, False, True, "Symbol for the Chart", True, , , , True)
            If aStrings.Size > 0 Then
                If InStr(aStrings(0), "|") > 0 Then
                    m.Chart.SetSymbol aStrings(0), True
                Else
                    nRec = g.SymbolPool.PoolRecForSymbol(aStrings(0), True)
                    If nRec >= 0 Then
                        i = LockWindowUpdate(pbChart.hWnd)
                        m.Chart.SetSymbol g.SymbolPool.SymbolID(nRec), True
                        m.Chart.GenerateChart eRedo1_Scrolled           '4376
                        If m.bGameMode And Not m.oGameMode Is Nothing Then
                            m.oGameMode.InitGame Me                     '4094
                        End If
                        If i <> 0 Then LockWindowUpdate 0
                    Else
                        Beep
                    End If
                End If
            End If
        End If
    End If
    
    If FormIsLoaded("frmTbMoreButtons") Then Unload frmTbMoreButtons
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".ChangeSymWrapper"
    
End Sub

Public Property Get BracketOrderOne() As cPtOrder
On Error GoTo ErrSection:
    
    Set BracketOrderOne = m.oBracketOrdOne

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".BracketOrderOneGet"

End Property

Public Property Get BracketOrderTwo() As cPtOrder
On Error GoTo ErrSection:
    
    Set BracketOrderTwo = m.oBracketOrdTwo

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".BracketOrderTwoGet"

End Property

Private Sub HandleBuySellClick(ctlClicked As Control, Optional ByVal Button As Integer = vbLeftButton)
On Error GoTo ErrSection:

    Dim strToolID$
        
    If m.eOrdBarMode = eOrdBarMode_Wizard And WizardGridMaxed Then
        'do nothing
    ElseIf Button = vbRightButton Then
        ClearBuySellButtons True
        frmChartOrdBar.ShowMe Me
        Exit Sub
    ElseIf ctlClicked.Appearance = apInset Then
        ClearBuySellButtons True
    Else
    
        If vseBracketOrder.Appearance = apInset Then ClearBuySellButtons True
    
        ctlClicked.Appearance = apInset
        
        If m.eOrdBarMode = eOrdBarMode_Order Then
            If vseBuyChart <> ctlClicked And vseBuyChart.Appearance = apInset Then vseBuyChart.Appearance = ap3D
            If vseSellChart <> ctlClicked And vseSellChart.Appearance = apInset Then vseSellChart.Appearance = ap3D
            If vseBracketOrder <> ctlClicked And vseBracketOrder.Appearance = apInset Then vseBracketOrder.Appearance = ap3D
        Else
            If vseBuyWizard <> ctlClicked And vseBuyWizard.Appearance = apInset Then vseBuyWizard.Appearance = ap3D
            If vseSellWizard <> ctlClicked And vseSellWizard.Appearance = apInset Then vseSellWizard.Appearance = ap3D
        End If
        
        
        g.ChartGlobals.eChartMode = eMode_ChartOrder
        
        Select Case ctlClicked
            Case vseBuyChart
                StatusMsg "Click on chart at desired price to buy ..."
                strToolID = "ID_ChartOrderBuy"
            Case vseSellChart
                StatusMsg "Click on chart at desired price to sell ..."
                strToolID = "ID_ChartOrderSell"
            Case vseBracketOrder
                StatusMsg "Click on chart at desired price to buy or sell ..."
                strToolID = "ID_ChartOrderBuy"
            Case vseBuyWizard
                strToolID = "ID_ChartOrderBuy"
            Case vseSellWizard
                strToolID = "ID_ChartOrderSell"
        End Select
        
        If Len(strToolID) > 0 Then
            ToolbarSetCursorGroup frmMain.tbToolbar, False, strToolID      '4877
        End If
        
        ClearAllBuySellBtns ctlClicked, Me.hWnd
        If vseBracketOrder.Appearance = apInset Then frmMain.tmrCheckBuySellButtons.Enabled = False
        If m.eOrdBarMode = eOrdBarMode_Wizard Then WizardControlsSync ctlClicked.BackColor
    
    End If
    
    SetFocusCtl

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".HandleBuySellClick"

End Sub

Public Sub GetPromptLocation(X As Long, Y As Long)
On Error GoTo ErrSection:

    Dim pt As POINTAPI

    pt.X = m.MouseDown.MouseX / Screen.TwipsPerPixelX
    pt.Y = m.MouseDown.MouseY / Screen.TwipsPerPixelY
    ClientToScreen Me.hWnd, pt
    pt.X = pt.X * Screen.TwipsPerPixelX - frmAsk.Width / 2
    pt.Y = pt.Y * Screen.TwipsPerPixelY - frmAsk.Height / 2
    
    X = pt.X
    Y = pt.Y

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".GetPromptLocation"

End Sub

Public Sub OrderbarWrapper(Optional ByVal eMode As eOrderBarMode = eOrdBarMode_Undefined)
On Error GoTo ErrSection:

    Dim bMOdeChanged As Boolean
    
    bMOdeChanged = True
    If m.eOrdBarMode = eOrdBarMode_Order And eMode = eOrdBarMode_BrokerDisconnect Then
        bMOdeChanged = False
        m.eOrdBarMode = eOrdBarMode_BrokerDisconnect
        FixOrderBarControls False, True
    ElseIf m.eOrdBarMode = eOrdBarMode_BrokerDisconnect And eMode = eOrdBarMode_Order Then
        bMOdeChanged = False
        m.eOrdBarMode = eOrdBarMode_Order
        FixOrderBarControls False, True
    End If

    If m.eOrdBarMode = eOrdBarMode_PFP Then
        If FormIsLoaded("frmPatternProfitOpt") Then Unload frmPatternProfitOpt
        PfpReset ePfpReset_ClearAll
        ClearAnnotFlags True, True
        g.strActiveDraw = ""
        Chart.RestoreChartNormal vbKeyClear
    End If
    
    If eMode = eOrdBarMode_PFP Then
        If m.eOrdBarMode = eOrdBarMode_PFP Then
            m.eOrdBarMode = eOrdBarMode_Undefined
        ElseIf m.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
            InitChartFlex Me, eFlexGridIdx_PfpInd
            If m.eOrdBarMode = eOrdBarMode_Wizard Then          '6393
                m.eOrdBarMode = eOrdBarMode_Order
                ToggleOrderBar False, True
                Chart.RestoreChartNormal vbKeyClear
            End If
            m.eOrdBarMode = eOrdBarMode_PFP
        Else
            bMOdeChanged = False
        End If
    ElseIf m.eOrdBarMode = eOrdBarMode_PFP Or m.eOrdBarMode = eOrdBarMode_Undefined Then
        m.eOrdBarMode = eOrdBarMode_Order     'user clicked orderbar button while pattern for profit bar is on
    End If
    
    If bMOdeChanged Then mnuOrderBar_Click
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".OrderbarWrapper"

End Sub

Public Sub PfpReset(eMode As ePfpResetMode, Optional ByVal bResetGridInd As Boolean = False)
On Error GoTo ErrSection:

    Dim strText$, dDateTo#, i&
    Dim iPfpForecast&, iNewForecast&
    
    Dim bLocked As Boolean
    Dim bReset As Boolean
    
    Dim Annot As cAnnotation
    Dim AnnotPFP As cAnnotation
    Dim Tree As cGdTree
    
    If m.eOrdBarMode <> eOrdBarMode_PFP Then Exit Sub
    If Not m.bFlexPFP Then Exit Sub
    
    Select Case eMode
        Case ePfpReset_GridInd
            If Not m.oPatternProfit Is Nothing Then strText = LoadIndGridPFP(m.Chart, fgChartFlex(eFlexGridIdx_PfpInd), bResetGridInd)
        Case ePfpReset_GridPfp
            fgChartFlex(eFlexGridIdx_PfpHits).Rows = fgChartFlex(eFlexGridIdx_PfpHits).FixedRows
            lblHitsFoundPFP.Caption = "Found:"
            lblPatternLen.Caption = "Pattern length:"
        Case ePfpReset_Forecastbars
            If Not m.oPatternProfit Is Nothing Then
                iPfpForecast = ValOfText(txtForecastPFP)
                
                strText = m.oPatternProfit.FindMatches(m.Chart, fgChartFlex(eFlexGridIdx_PfpInd), fgChartFlex(eFlexGridIdx_PfpHits), ValOfText(txtCorrPercentPFP.Text), , True)
                
                If Len(strText) = 0 Then
                    dDateTo = m.oPatternProfit.PatternDateTo
                    i = m.Chart.Bars.FindDateTime(dDateTo)
                    
                    If i + iPfpForecast > m.Chart.Bars.Size Then
                        iNewForecast = m.Chart.ForecastBars + ((i + iPfpForecast) - m.Chart.Bars.Size)
                        If iNewForecast > m.Chart.geChartPoints - 3 Then
                            bReset = True
                            iPfpForecast = iPfpForecast - 1
                            While iNewForecast > m.Chart.geChartPoints - 3
                                iPfpForecast = iPfpForecast - 1
                                iNewForecast = m.Chart.ForecastBars + ((i + iPfpForecast) - m.Chart.Bars.Size)
                            Wend
                        End If
                        
                        If bReset Then
                            InfBox "Forecast bars too large adjusted to " & Str(iPfpForecast), "I", "Ok", "Pattern for Profit"
                            txtForecastPFP.Text = Str(iPfpForecast)
                        End If
                        
                        m.Chart.ForecastBars(Me) = m.Chart.ForecastBars + ((i + iPfpForecast) - m.Chart.Bars.Size)
                        
                        bLocked = LockWindowUpdate(Me.hWnd)
                        m.Chart.GenerateChart eRedo7_ReloadRT
                        If bLocked Then LockWindowUpdate 0
                        hsb.Value = hsb.Max
                    End If
                End If
            End If
        Case ePfpReset_PpfAnnot, ePfpReset_PpfAnnotInd, ePfpReset_PpfPattern
            'ePfpReset_PpfAnnot: clear matched PFP annots
            'ePfpReset_PpfAnnotInd: clear matched PFP annots & PFP indicators
            'ePfpReset_PpfPattern: clear all PFP annots & PFP indcators AND PFP pattern
            
            If Not m.Chart Is Nothing Then
                If eMode <> ePfpReset_PpfAnnot Then
                    If Not m.oPatternProfit Is Nothing Then m.oPatternProfit.ClearPFP   'this will clear out the PFP indicators
                End If
                
                If eMode <> ePfpReset_PpfPattern Then
                    'clears out the matched annotations, but not the pattern itself
                    Set Tree = m.Chart.Annots
                    For i = 1 To Tree.Count
                        Set Annot = Tree(i)
                        If Not Annot Is Nothing Then
                            If Annot.eUsage = eANNOT_PatternProfit Then
                                If InStr(Tree.Key(i), "PFP") = 0 Then
                                    Set AnnotPFP = Annot
                                    AnnotPFP.eUsage = eANNOT_UserAdded      'temporarily change usage so will not get deleted
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                ElseIf Len(g.strActiveDraw) > 0 Then
                    ClearAnnotFlags True, True
                    g.strActiveDraw = ""
                    SyncDrawTools
                End If
                
                If Not m.Chart Is Nothing Then m.Chart.RemoveAnnots False, , eANNOT_PatternProfit
                If Not AnnotPFP Is Nothing Then AnnotPFP.eUsage = eANNOT_PatternProfit
                
                m.Chart.GenerateChart eRedo1_Scrolled
            End If
        Case ePfpReset_ClearAll
            cmdMatchesPFP.Enabled = False
            
            If Not m.oPatternProfit Is Nothing Then m.oPatternProfit.ClearPFP
            If Not m.Chart Is Nothing Then m.Chart.RemoveAnnots False, , eANNOT_PatternProfit
            
            fgChartFlex(eFlexGridIdx_PfpHits).Rows = fgChartFlex(eFlexGridIdx_PfpHits).FixedRows
            lblHitsFoundPFP.Caption = "Found:"
            lblPatternLen.Caption = "Pattern length:"
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".PfpReset"

End Sub

Public Property Get PatternProfitObj() As cPatternProfit
    Set PatternProfitObj = m.oPatternProfit
End Property

Public Property Let PatternProfitObj(obj As cPatternProfit)
    Set m.oPatternProfit = obj
End Property

Private Function HandleGartley(ByVal bMouseUp As Boolean, dY#) As Boolean
On Error GoTo ErrSection:
    
    Dim bSnapped As Boolean
    Dim bMoved As Boolean
    Dim Annot As cAnnotation
    Dim nFreeFloat&
       
    HandleGartley = False
           
    'check if editing (currently this annotation is not moveable after creation)
    If Len(g.strActiveDraw) = 0 Then
        If bMouseUp = True Then
            If Len(g.strActiveDraw) = 0 Then
                If m.nObjectMoving = 1 Then
                    tmr.Tag = "EditAnnot " & CStr(m.nActiveAnnotIdx)
                End If
                ClearAnnotFlags False
                m.nActiveIndIdx = 0
                m.nObjectMoving = 0
                m.Chart.SetCursor
                Exit Function
            End If
        End If
    End If
                  
    Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
    If m.nPointCount > 0 Then
        If Annot Is Nothing Then Exit Function
        If m.MouseLast.nBar < 0 Then Exit Function      'empty bar
    End If
               
    If Annot Is Nothing Then
        Set Annot = New cAnnotation
        nFreeFloat = Annot.DefaultProp(Annot.AnnotTypeFromToolID(g.strActiveDraw), "FreeFloat")
    Else
        nFreeFloat = ValOfText(Annot.Prop("FreeFloat"))
    End If
    
    If m.nPointCount = 0 And (Len(g.strActiveDraw) > 0 Or nFreeFloat = 1) Then
        If nFreeFloat = 1 Then
            dY = SnapToPrice(m.MouseLast)
            bSnapped = True
        Else
            bSnapped = SnapToHiLoClose(m.MouseLast, dY, 0, True)
        End If
        If bSnapped = False Then Exit Function
        If m.nFocusHiLo = -1 Or m.nFocusHiLo = 3 Then
            ClearAnnotFlags False
            Exit Function
        End If
        If (Len(g.strActiveDraw) > 0) Then
            bMoved = True
        Else
            bMoved = Annot.MovePoint(m.Chart, 1, m.MouseLast.nPaneID, m.MouseLast.dDate, dY)
            Annot.Prop("HiLo") = m.nFocusHiLo - 1
        End If
    ElseIf m.nPointCount = 1 Then
        If nFreeFloat = 1 Then
            dY = SnapToPrice(m.MouseLast)
            bSnapped = True
        Else
            If m.nFocusHiLo = 1 Or m.nFocusHiLo = 4 Then  'focus is bar's high
                bSnapped = SnapToHiLoClose(m.MouseLast, dY, 2, True)  ', bMouseUp)
            Else
                bSnapped = SnapToHiLoClose(m.MouseLast, dY, 1, True)  ', bMouseUp)
            End If
        End If
        If bSnapped = True Then
            bMoved = Annot.MovePoint(m.Chart, 2, m.MouseLast.nPaneID, m.MouseLast.dDate, dY)
        End If
    ElseIf m.nPointCount = 2 Then
        'do not snap the last point(D) to high or low (user may want D to be beyond last data bar)
        'bSnapped = SnapToHiLoClose(m.MouseLast, dY, 4)
        dY = SnapToPrice(m.MouseLast)
        bMoved = Annot.MovePoint(m.Chart, 3, m.MouseLast.nPaneID, m.MouseLast.dDate, dY)
    Else
        Set Annot = m.Chart.Annots(m.nActiveAnnotIdx)
        If Not Annot Is Nothing Then
            dY = SnapToPrice(m.MouseLast)
            bMoved = Annot.MovePoint(m.Chart, m.nActiveAnnotPt, m.MouseLast.nPaneID, m.MouseLast.dDate, dY)
        End If
    End If
        
    If bMoved = True Then
        If bMouseUp = True Then
            If m.nPointCount < 2 Then
                m.nPointCount = m.nPointCount + 1
                If m.nPointCount = 1 Then
                    StatusMsg "Now click on the second peak ...", -1
                Else
                    StatusMsg "Now click on the third peak ...", -1
                End If
            Else
                StatusMsg
                ClearAnnotFlags False
                'set flag notifying grapheng.dll that points selection is complete
                Annot.geMoveFlag = 1
                m.Chart.SyncGlobalAnnots Annot
                SyncDrawTools
            End If
        Else
            Annot.geDrawAnn m.Chart
        End If
        
        ' TLB 8/2/2013: we need this to be done in order for the PeakCheck to work!
        m.MouseDown = m.MouseLast
    End If
        
    Set Annot = Nothing
    HandleGartley = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".HandleGartley"

End Function

Private Sub SetGartleyActivePt(Annot As cAnnotation, hitInfo As hittest_info)
On Error GoTo ErrSection:


    If hitInfo.location = 10 Or hitInfo.location = 11 Then
        m.epbCursor = eCursor_Arrow4Way
        m.nActiveAnnotPt = hitInfo.itemIndex
    ElseIf hitInfo.location = 9 Then
        m.epbCursor = eCursor_Hand
    Else
        m.nActiveAnnotPt = -1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetGartleyActivePt"

End Sub

Private Sub UpdateSeasonalControls()
On Error GoTo ErrSection:

    Dim strText$, d#, i&, j&
    Dim bShow As Boolean
    
    Dim Ind As cIndicator
    Dim eCycleType As eSeasonalCycleType
    

    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Or Not Me.Visible Then Exit Sub
    If m.Chart Is Nothing Then Exit Sub
    If m.eSeasonalCtrlsState = eSeasonCtrlStatus_Updated Then Exit Sub
    If vseInvalid.Visible Then Exit Sub
    
    If m.Chart.TypeOfChart <> eTypeChart_Seasonal Then
        m.eSeasonalCtrlsState = eSeasonCtrlStatus_Updated       'fix 'sluggish' issue reported by Vanessa
        Exit Sub
    End If
    
    StatusMsg ""
    
    'from date
    d = m.Chart.FromDate
    If gdSeasonalDateFrom.Value <> d Then
        gdSeasonalDateFrom.Value = d
    End If

    'cycle
    eCycleType = m.Chart.SeasonalCycleTypeEnum
    If eCycleType >= 0 And eCycleType < cboCycle.ListCount Then cboCycle.ListIndex = eCycleType
    txtCycleNum.Text = Str(m.Chart.SeasonalCycleLen)
    
    'bar type
    strText = m.Chart.Bars.Prop(eBARS_PeriodicityStr)
    If InStr(strText, " ") <> 0 Then strText = Parse(strText, " ", 2)
    
    If eCycleType = eCycleType_Year Or eCycleType = eCycleType_Quarter Then
        If cboBarType.Text <> strText Then
            Select Case strText
                Case "Daily"
                    cboBarType.ListIndex = 0
                Case "Weekly"
                    cboBarType.ListIndex = 1
                Case "Monthly"
                    cboBarType.ListIndex = 2
            End Select
        End If
        cboBarType.Enabled = True
    Else
        cboBarType.ListIndex = 0
        cboBarType.Enabled = False
    End If
    
    'current cycle is usually last in tree, do this first so don't have to walk through entire tree below
    Set Ind = m.Chart.Tree(kSeasonalCurrIndKey)
    If Not Ind Is Nothing Then
        j = Ind.Style
        If j <> cboTrendStyle(3).ListIndex And fgChartFlex(eFlexGridIdx_Seasonal).Row <> fgChartFlex(eFlexGridIdx_Seasonal).Rows - 1 Then
            If j >= 0 And j < cboTrendStyle(3).ListCount Then cboTrendStyle(3).ListIndex = j
        End If
        If Ind.Display Then
            chkTrendShow(3).Value = vbChecked
        Else
            chkTrendShow(3).Value = vbUnchecked
        End If
        gdTrendColor(3).Color = Ind.Color
    End If
    
    'check boxes, color, style for trends & other cycles
    With m.Chart
        For i = .Tree.Count To 1 Step -1
            If TypeOf .Tree(i) Is cIndicator Then
                Set Ind = .Tree(i)
                
                If Ind.MyKey = "PRICE" Or Ind.MyKey = kSeasonalCurrIndKey Or Ind.DataType = eINDIC_Constant Then
                    'do nothing
                ElseIf Ind.MyKey = kSeasonalAvgIndKey Then
                    'average trend AND gradient from/to colors
                    j = Ind.Style
                    If j <> cboTrendStyle(0).ListIndex Then 'fgchartflex(eFlexGridIdx_Seasonal).Row <> fgchartflex(eFlexGridIdx_Seasonal).FixedRows Then
                        If j >= 0 And j < cboTrendStyle(0).ListCount Then cboTrendStyle(0).ListIndex = j
                    End If
                    
                    If Ind.Display Then
                        chkTrendShow(0).Value = vbChecked
                    Else
                        chkTrendShow(0).Value = vbUnchecked
                    End If
                    
                    If Ind.Overlayed Then
                        chkOverlayTrends.Value = vbChecked
                    Else
                        chkOverlayTrends.Value = vbUnchecked
                    End If
                    
                    gdTrendColor(0).Color = Ind.Color
                    gdTrendColor(4).Color = Ind.UpColor
                    gdTrendColor(5).Color = Ind.DownColor
                
                ElseIf Ind.MyKey = kSeasonalBullIndKey Then
                    'bullish trend
                    j = Ind.Style
                    If j <> cboTrendStyle(1).ListIndex Then
                        If j >= 0 And j < cboTrendStyle(1).ListCount Then cboTrendStyle(1).ListIndex = j
                    End If
                    
                    If Ind.Display Then
                        chkTrendShow(1).Value = vbChecked
                    Else
                        chkTrendShow(1).Value = vbUnchecked
                    End If
                    
                    gdTrendColor(1).Color = Ind.Color
                
                ElseIf Ind.MyKey = kSeasonalBearIndKey Then
                    'bearish trend
                    j = Ind.Style
                    If j <> cboTrendStyle(2).ListIndex Then
                        If j >= 0 And j < cboTrendStyle(2).ListCount Then cboTrendStyle(2).ListIndex = j
                    End If
                    
                    If Ind.Display Then
                        chkTrendShow(2).Value = vbChecked
                    Else
                        chkTrendShow(2).Value = vbUnchecked
                    End If
                    
                    gdTrendColor(2).Color = Ind.Color
                    
                Else
                    'all other cycles
                    bShow = Ind.Display
                    If bShow Then
                        chkShowCycles.Value = vbChecked
                    Else
                        chkShowCycles.Value = vbUnchecked
                    End If
'                    chkOverlayCycles.Enabled = bShow
                
                    lblGradient.Enabled = bShow
                    lblGradientFrom.Enabled = bShow
                    lblGradientTo.Enabled = bShow
                    
                    gdTrendColor(4).Enabled = bShow
                    gdTrendColor(5).Enabled = bShow
                    
                    Exit For
                End If
            End If
        Next
    End With
    
    If Not m.Chart Is Nothing Then
        If m.Chart.SeasonalHasOverlay Then
            chkOverlayCycles.Value = vbChecked
        Else
            chkOverlayCycles.Value = vbUnchecked
        End If
    End If

ErrExit:
    m.eSeasonalCtrlsState = eSeasonCtrlStatus_Updated
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".UpdateSeasonalControls"

End Sub

Private Sub HandleSeasonalInput(ByVal eCtrlType As eSeasonalCtrlType)
On Error GoTo ErrExit:

    Static idxPrev&, iRowPrev&
    Static eStylePrev As eIndicatorStyle

    Dim idx&, iRedrawSave&, i&
    Dim iColorFrom&, iColorTo&, iCount&, iColorCounter&
    
    Dim Ind As cIndicator
    Dim IndPrev As cIndicator
    
    Dim bRedraw As Boolean
    Dim bOverlay As Boolean
    Dim bShow As Boolean
    Dim bShowAvgTrend As Boolean
    
    Dim RedoMode As eChartRedoMode
    
    If Not Me.Visible Then Exit Sub
    If m.Chart Is Nothing Then Exit Sub
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    
    If m.eSeasonalCtrlsState = eSeasonCtrlStatus_Unpopulated Then Exit Sub
    
    tmr.Enabled = False
    RedoMode = eRedo1_Scrolled
    
    Select Case eCtrlType
        Case eSeasonalCtrl_BarType, eSeasonalCtrl_CycleType, eSeasonalCtrl_CycleNum, eSeasonalCtrl_FromDate
            If Not vseInvalid.Visible Then
                cmdSeasonalApply.Enabled = True
                m.Chart.SeasonalIndClear
                If eCtrlType = eSeasonalCtrl_FromDate Then
                    MoveFocus gdSeasonalDateFrom
                ElseIf eCtrlType = eSeasonalCtrl_CycleNum Then
                    MoveFocus txtCycleNum
                End If
            End If
            
            If eCtrlType = eSeasonalCtrl_FromDate Then
                m.Chart.FromDate = CDbl(gdSeasonalDateFrom.Value)
            ElseIf eCtrlType = eSeasonalCtrl_CycleType Or eCtrlType = eSeasonalCtrl_CycleNum Then
                Select Case cboCycle.ListIndex
                    Case eCycleType_Year, eCycleType_Quarter
                        cboBarType.Enabled = True
                    Case eCycleType_Month, eCycleType_FullMoons, eCycleType_NewMoons
                        'for anything that is not year or quarter, only allow Daily bar type
                        If cboBarType.ListIndex <> 0 Then cboBarType.ListIndex = 0
                        cboBarType.Enabled = False
                    Case eCycleType_Week
                        If cboBarType.ListIndex <> 0 Then cboBarType.ListIndex = 0
                        cboBarType.Enabled = False
                        If Int(ValOfText(txtCycleNum.Text)) < 4 Then
                            StatusMsg "Weekly cycle length must be a minum of 4 weeks."
                            txtCycleNum.Text = 4
                        End If
                    Case eCycleType_Days
                        If cboBarType.ListIndex <> 0 Then cboBarType.ListIndex = 0
                        cboBarType.Enabled = False
                        If Int(ValOfText(txtCycleNum.Text)) < 20 Then
                            StatusMsg "Daily cycle length must be a minum of 20 days."
                            txtCycleNum.Text = 20
                        End If
                End Select
                
            End If
            
            GoTo ErrExit
        
        Case eSeasonalCtrl_SeasonalGrid
            If idxPrev > 0 And idxPrev <= m.Chart.Tree.Count Then
                If TypeOf m.Chart.Tree(idxPrev) Is cIndicator Then Set IndPrev = m.Chart.Tree(idxPrev)
            End If
        
            With fgChartFlex(eFlexGridIdx_Seasonal)
                If .Row >= .FixedRows And .Row < .Rows Then
                    idx = Val(.TextMatrix(.Row, 1))
                    If idx <> idxPrev Then
                        If TypeOf m.Chart.Tree(idx) Is cIndicator Then Set Ind = m.Chart.Tree(idx)
                            
                        If Not Ind Is Nothing Then
                            If Not IndPrev Is Nothing Then
                                IndPrev.Style = eStylePrev
                                
                                If IndPrev.MyKey = kSeasonalAvgIndKey Or _
                                   IndPrev.MyKey = kSeasonalBullIndKey Or _
                                   IndPrev.MyKey = kSeasonalBearIndKey Or _
                                   IndPrev.MyKey = kSeasonalCurrIndKey Then
                                   
                                    'do nothing
                                
                                ElseIf iRowPrev >= .FixedRows And iRowPrev < .Rows Then
                                    .Cell(flexcpFontBold, iRowPrev, 0) = False
                                    If chkShowCycles.Value = vbUnchecked And IndPrev.Display Then IndPrev.Display = False
                                End If
                            End If
                        
                            eStylePrev = Ind.Style
                            idxPrev = idx
                            Ind.Style = eINDIC_Thick
                            
                            .Cell(flexcpFontBold, .Row, 0) = True
                            
                            If Ind.MyKey = kSeasonalAvgIndKey Or _
                               Ind.MyKey = kSeasonalBullIndKey Or _
                               Ind.MyKey = kSeasonalBearIndKey Or _
                               Ind.MyKey = kSeasonalCurrIndKey Then
                               
                                .Cell(flexcpForeColor, .Row, 0) = Ind.Color
                            Else
                                .Cell(flexcpForeColor, .Row, 0) = Ind.UpColor
                                If Not Ind.Display Then Ind.Display = True
                            End If
                            bRedraw = True
                            iRowPrev = .Row
                        End If
                    End If
                End If
            End With
        
'avgtrend checkbox, style & color
        Case eSeasonalCtrl_OverlayTrends
            Set Ind = m.Chart.Tree(kSeasonalAvgIndKey)
            If Not Ind Is Nothing Then
                If chkOverlayTrends.Value = vbChecked Then
                    If Not Ind.Overlayed Then Ind.Overlayed = True
                    If chkShowCycles.Value = vbChecked Then
                        chkShowCycles.Value = vbUnchecked
                        With m.Chart
                            For i = 1 To .Tree.Count
                                If TypeOf .Tree(i) Is cIndicator Then
                                    Set Ind = .Tree(i)
                                    
                                    If Ind.MyKey = "PRICE" Or Ind.MyKey = kSeasonalCurrIndKey Or _
                                        Ind.MyKey = kSeasonalBullIndKey Or Ind.MyKey = kSeasonalBearIndKey Or _
                                        Ind.DataType = eINDIC_Constant Or Ind.MyKey = kSeasonalAvgIndKey Then
                                       
                                       'do nothing
                                       
                                    Else
                                        Ind.Display = False
                                    End If
                                End If
                            Next
                        End With
                        
                        Set Ind = m.Chart.Tree(kSeasonalAvgIndKey)
                    End If
                ElseIf Ind.Overlayed Then
                    Ind.Overlayed = False
                End If
                
                If Not m.Chart.Tree(kSeasonalBullIndKey) Is Nothing Then
                    m.Chart.Tree(kSeasonalBullIndKey).Overlayed = Ind.Overlayed
                End If
                
                If Not m.Chart.Tree(kSeasonalBearIndKey) Is Nothing Then
                    m.Chart.Tree(kSeasonalBearIndKey).Overlayed = Ind.Overlayed
                End If
                
                If Not m.Chart.Tree(kSeasonalCurrIndKey) Is Nothing Then
                    m.Chart.Tree(kSeasonalCurrIndKey).Overlayed = Ind.Overlayed
                End If
                
                RedoMode = eRedo5_RecalcInd
                bRedraw = True
            End If
        
        Case eSeasonalCtrl_AvgTrendCheckBox
            Set Ind = m.Chart.Tree(kSeasonalAvgIndKey)
            If Not Ind Is Nothing Then
                If chkTrendShow(0).Value = vbChecked Then
                    Ind.Display = True
                    If m.Chart.SeasonalHasOverlay Then RedoMode = eRedo9_ReloadData
                Else
                    Ind.Display = False
                End If
                bRedraw = True
            End If
            
        Case eSeasonalCtrl_AvgTrendStyle
            Set Ind = m.Chart.Tree(kSeasonalAvgIndKey)
            If Not Ind Is Nothing Then
                If Ind.Style <> cboTrendStyle(0).ListIndex Then
                    Ind.Style = cboTrendStyle(0).ListIndex
                    bRedraw = True
                End If
            End If

        Case eSeasonalCtrl_AvgTrendColor
            Set Ind = m.Chart.Tree(kSeasonalAvgIndKey)
            If Not Ind Is Nothing Then
                Ind.Color = gdTrendColor(0).Color
                fgChartFlex(eFlexGridIdx_Seasonal).Cell(flexcpForeColor, fgChartFlex(eFlexGridIdx_Seasonal).FixedRows, 0) = Ind.Color
                bRedraw = True
            End If

'bullish trend checkbox, style & color
        Case eSeasonalCtrl_BullTrendCheckBox
            Set Ind = m.Chart.Tree(kSeasonalBullIndKey)
            If Not Ind Is Nothing Then
                If chkTrendShow(1).Value = vbChecked Then
                    Ind.Display = True
                    If m.Chart.SeasonalHasOverlay Then RedoMode = eRedo9_ReloadData
                Else
                    Ind.Display = False
                End If
                bRedraw = True
            End If
        
        Case eSeasonalCtrl_BullTrendStyle
            Set Ind = m.Chart.Tree(kSeasonalBullIndKey)
            If Not Ind Is Nothing Then
                If Ind.Style <> cboTrendStyle(1).ListIndex Then
                    Ind.Style = cboTrendStyle(1).ListIndex
                    bRedraw = True
                End If
            End If
        
        Case eSeasonalCtrl_BullTrendColor
            Set Ind = m.Chart.Tree(kSeasonalBullIndKey)
            If Not Ind Is Nothing Then
                Ind.Color = gdTrendColor(1).Color
                fgChartFlex(eFlexGridIdx_Seasonal).Cell(flexcpForeColor, fgChartFlex(eFlexGridIdx_Seasonal).FixedRows + 1, 0) = Ind.Color
                bRedraw = True
            End If

'bearish trend checkbox, style & color
        Case eSeasonalCtrl_BearTrendCheckBox
            Set Ind = m.Chart.Tree(kSeasonalBearIndKey)
            If Not Ind Is Nothing Then
                If chkTrendShow(2).Value = vbChecked Then
                    Ind.Display = True
                    If m.Chart.SeasonalHasOverlay Then RedoMode = eRedo9_ReloadData
                Else
                    Ind.Display = False
                End If
                bRedraw = True
            End If
        
        Case eSeasonalCtrl_BearTrendStyle
            Set Ind = m.Chart.Tree(kSeasonalBearIndKey)
            If Not Ind Is Nothing Then
                If Ind.Style <> cboTrendStyle(2).ListIndex Then
                    Ind.Style = cboTrendStyle(2).ListIndex
                    bRedraw = True
                End If
            End If
        
        Case eSeasonalCtrl_BearTrendColor
            Set Ind = m.Chart.Tree(kSeasonalBearIndKey)
            If Not Ind Is Nothing Then
                Ind.Color = gdTrendColor(2).Color
                fgChartFlex(eFlexGridIdx_Seasonal).Cell(flexcpForeColor, fgChartFlex(eFlexGridIdx_Seasonal).FixedRows + 2, 0) = Ind.Color
                bRedraw = True
            End If

'current cycle style & color
        Case eSeasonalCtrl_CurrCycleCheckBox
            Set Ind = m.Chart.Tree(kSeasonalCurrIndKey)
            If Not Ind Is Nothing Then
                If chkTrendShow(3).Value = vbChecked Then
                    Ind.Display = True
                    If m.Chart.SeasonalHasOverlay Then RedoMode = eRedo9_ReloadData
                Else
                    Ind.Display = False
                End If
                bRedraw = True
            End If
        
        Case eSeasonalCtrl_CurrCycleStyle
            Set Ind = m.Chart.Tree(kSeasonalCurrIndKey)
            If Not Ind Is Nothing Then
                If Ind.Style <> cboTrendStyle(3).ListIndex Then
                    Ind.Style = cboTrendStyle(3).ListIndex
                    bRedraw = True
                End If
            End If

        Case eSeasonalCtrl_CurrCycleColor
            Set Ind = m.Chart.Tree(kSeasonalCurrIndKey)
            If Not Ind Is Nothing Then
                Ind.Color = gdTrendColor(3).Color
                fgChartFlex(eFlexGridIdx_Seasonal).Cell(flexcpForeColor, fgChartFlex(eFlexGridIdx_Seasonal).Rows - 1, 0, fgChartFlex(eFlexGridIdx_Seasonal).Rows - 1, 1) = Ind.Color
                bRedraw = True
            End If

'Overlay flag, From/To colors for other cycles
        Case eSeasonalCtrl_OtherColorTo, _
             eSeasonalCtrl_OtherColorFrom, _
             eSeasonalCtrl_OverlayCycles
        
            iColorFrom = gdTrendColor(4).Color
            iColorTo = gdTrendColor(5).Color
            
            If chkOverlayCycles.Value = vbChecked Then bOverlay = True
            If eCtrlType = eSeasonalCtrl_OverlayCycles Then
                If m.Chart.SeasonalHasOverlay <> bOverlay Then RedoMode = eRedo5_RecalcInd
            End If
            
            Set Ind = m.Chart.Tree(kSeasonalAvgIndKey)
            If Not Ind Is Nothing Then
                If Ind.UpColor <> iColorFrom Or Ind.DownColor <> iColorTo Or RedoMode = eRedo5_RecalcInd Then
                    Ind.UpColor = iColorFrom
                    Ind.DownColor = iColorTo
                    
                    iCount = m.Chart.Tree.Count - 4
                    iColorCounter = 1
                    For i = 1 To m.Chart.Tree.Count
                        If TypeOf m.Chart.Tree(i) Is cIndicator Then
                            Set Ind = m.Chart.Tree(i)
                            If Ind.MyKey = kSeasonalAvgIndKey Or Ind.MyKey = "PRICE" Or _
                               Ind.MyKey = kSeasonalBullIndKey Or Ind.MyKey = kSeasonalBearIndKey Or _
                               Ind.DataType = eINDIC_Constant Then
                                'do nothing
                            ElseIf Ind.MyKey = kSeasonalCurrIndKey Then
                                'only set overlay flag
                                Ind.Overlayed = bOverlay
                            Else
                                'set overlay & color
                                Ind.Color = GradientColor(100 * iColorCounter / iCount, iColorTo, iColorFrom)       '6360
                                Ind.Overlayed = bOverlay
                                iColorCounter = iColorCounter + 1
                            End If
                        End If
                    Next
                    
                    bRedraw = True
                    
                End If
            End If

        Case eSeasonalCtrl_ShowCycles
            If chkShowCycles.Value = vbChecked Then
                bShow = True
            ElseIf chkTrendShow(0).Value = vbUnchecked And _
                chkTrendShow(1).Value = vbUnchecked And _
                chkTrendShow(2).Value = vbUnchecked Then
                
                bShowAvgTrend = True
            End If
            
            With m.Chart
                For i = 1 To .Tree.Count
                    If TypeOf .Tree(i) Is cIndicator Then
                        Set Ind = .Tree(i)
                        
                        If Ind.MyKey = kSeasonalAvgIndKey Then
                            If bShowAvgTrend = True Then Ind.Display = True
                        ElseIf Ind.MyKey = "PRICE" Or Ind.MyKey = kSeasonalCurrIndKey Or _
                            Ind.MyKey = kSeasonalBullIndKey Or Ind.MyKey = kSeasonalBearIndKey Or _
                            Ind.DataType = eINDIC_Constant Then
                           
                           'do nothing
                           
                        Else
                            Ind.Display = bShow
                        End If
                    End If
                Next
            End With
            bRedraw = True
    
    End Select
    
    If chkShowCycles.Value = vbChecked Then
'07-13-2011: Glen wants overlay & show all to be independent (suggestion from Rob)
'        If Not chkOverlayCycles.Enabled Then chkOverlayCycles.Enabled = True
        
        lblGradient.Enabled = True
        lblGradientFrom.Enabled = True
        lblGradientTo.Enabled = True
        gdTrendColor(4).Enabled = True
        gdTrendColor(5).Enabled = True
    Else
'        If chkOverlayCycles.Enabled Then chkOverlayCycles.Enabled = False
                
        lblGradient.Enabled = False
        lblGradientFrom.Enabled = False
        lblGradientTo.Enabled = False
        gdTrendColor(4).Enabled = False
        gdTrendColor(5).Enabled = False
        
        If chkTrendShow(0).Value = vbUnchecked And chkTrendShow(1).Value = vbUnchecked And _
           chkTrendShow(2).Value = vbUnchecked And chkTrendShow(3).Value = vbUnchecked Then
           
            chkTrendShow(0).Value = vbChecked
            
        End If
    End If
    
ErrExit:
    If bRedraw Then m.Chart.GenerateChart RedoMode
    If Not vseInvalid.Visible Then tmr.Enabled = True
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".HandleSeasonalInput"

End Sub

Public Sub SeasonalControlsReset()
    m.eSeasonalCtrlsState = eSeasonCtrlStatus_Unpopulated
End Sub

Private Sub SetFlexVisible(ByVal eIndex As eChartFlexCtrlIndex, ByVal bVisible As Boolean)
On Error Resume Next

    Select Case eIndex
        
        Case eFlexGridIdx_OrdWizard
            If m.bFlexOrdBar Then
                If fgChartFlex(eFlexGridIdx_OrdWizard).Visible <> bVisible Then
                    fgChartFlex(eFlexGridIdx_OrdWizard).Visible = bVisible
                End If
            End If
        
        Case eFlexGridIdx_PfpInd, eFlexGridIdx_PfpHits
            If m.bFlexPFP Then
                If fgChartFlex(eIndex).Visible <> bVisible Then
                    fgChartFlex(eIndex).Visible = bVisible
                End If
            End If
    
    End Select

End Sub

Public Property Get GridLoaded(ByVal eIndex As eChartFlexCtrlIndex) As Boolean
On Error GoTo ErrSection:

    Select Case eIndex
        Case eFlexGridIdx_OrdWizard
            GridLoaded = m.bFlexOrdBar
        Case eFlexGridIdx_Seasonal
            GridLoaded = m.bFlexSeasonal
        Case eFlexGridIdx_PfpInd, eFlexGridIdx_PfpHits
            GridLoaded = m.bFlexPFP
    End Select

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".GridLoaded.Get"

End Property

Public Property Let GridLoaded(ByVal eIndex As eChartFlexCtrlIndex, ByVal bLoaded As Boolean)
On Error GoTo ErrSection:

    Select Case eIndex
        Case eFlexGridIdx_OrdWizard
            m.bFlexOrdBar = bLoaded
        Case eFlexGridIdx_Seasonal
            m.bFlexSeasonal = bLoaded
        Case eFlexGridIdx_PfpInd, eFlexGridIdx_PfpHits
            m.bFlexPFP = bLoaded
    End Select

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError Me.Name & ".GridLoaded.Let"

End Property

Public Sub GridTextIncrease()
On Error GoTo ErrSection:

    g.ChartGlobals.nFontSize = g.ChartGlobals.nFontSize + 1
    m.Chart.GenerateChart eRedo3_Settings

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".GridTextIncrease"

End Sub

Public Sub GridTextDecrease()
On Error GoTo ErrSection:

    g.ChartGlobals.nFontSize = g.ChartGlobals.nFontSize - 1
    m.Chart.GenerateChart eRedo3_Settings

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".GridTextDecrease"

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
Public Sub InitQuantityEditor()
On Error GoTo ErrSection:

    Static lAccountID As Long           ' Account ID
    Static vSymbolOrSymbolID As Variant ' Symbol or Symbol id
    Static strSecType As String         ' Security type for the information

    If (TradeAccountID <> lAccountID) Or (SymbolOrSymbolID <> vSymbolOrSymbolID) Then
        lAccountID = TradeAccountID
        vSymbolOrSymbolID = SymbolOrSymbolID
        
        g.Broker.InitQuantityEditor m.Quantity, vscrQty, txtTradeQty, TradeAccountID, SymbolOrSymbolID
        
        If strSecType <> g.Broker.TradeSecType(vSymbolOrSymbolID) Then
            strSecType = g.Broker.TradeSecType(vSymbolOrSymbolID)
            SetQuantityPresetButtons
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".InitQuantityEditor"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetQuantityPresetButtons
'' Description: Set the quantity preset buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetQuantityPresetButtons()
On Error GoTo ErrSection:

    Dim strSecType As String            ' Security type for the chart symbol

    strSecType = g.Broker.TradeSecType(SymbolOrSymbolID)
    g.Broker.GetQuantityPresets TradeAccountID, SymbolOrSymbolID, m.lPreset1, m.lPreset2, m.lPreset3
    cmdQty1.Caption = ShortDisplayNumber(m.lPreset1)
    cmdQty2.Caption = ShortDisplayNumber(m.lPreset2)
    cmdQty3.Caption = ShortDisplayNumber(m.lPreset3)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".SetQuantityPresetButtons"
    
End Sub

Public Sub ClearFormForReuse()
On Error GoTo ErrSection:

    Tag = ""
    tmr.Tag = ""
    PfpReset ePfpReset_ClearAll
    vseSeasonal.Visible = False
    SeasonalControlsReset
    If OrderBarMode = eOrdBarMode_Wizard Then
        OrderBarMode = eOrdBarMode_Order   '5231
    ElseIf OrderBarMode = eOrdBarMode_PFP Then
        OrderBarMode = eOrdBarMode_Undefined
    End If
    
    'clear grids in case form gets reused for a non-seasonal chart type (aardvark 6399)
    If GridLoaded(eFlexGridIdx_OrdWizard) Then
        Unload fgChartFlex(eFlexGridIdx_OrdWizard)
        GridLoaded(eFlexGridIdx_OrdWizard) = False
    End If
    
    If GridLoaded(eFlexGridIdx_Seasonal) Then
        Unload fgChartFlex(eFlexGridIdx_Seasonal)
        GridLoaded(eFlexGridIdx_Seasonal) = False
    End If
    
    If GridLoaded(eFlexGridIdx_PfpInd) Then
        Unload fgChartFlex(eFlexGridIdx_PfpInd)
        Unload fgChartFlex(eFlexGridIdx_PfpHits)
        GridLoaded(eFlexGridIdx_PfpInd) = False
    End If
    
    Set m.aTabs = New cGdArray

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".ClearFormForReuse"
    
End Sub

