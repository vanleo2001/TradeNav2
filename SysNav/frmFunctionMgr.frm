VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmFunctionMgr 
   Caption         =   "Function Manager"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9105
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   8580
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   5
      DisplayContextMenu=   0   'False
      Tools           =   "frmFunctionMgr.frx":0000
      ToolBars        =   "frmFunctionMgr.frx":01AE
   End
   Begin HexUniControls.ctlUniFrameWL fraFunctionInfo 
      Height          =   1815
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   8895
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmFunctionMgr.frx":02E4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFunctionMgr.frx":0304
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFunctionMgr.frx":0324
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cbCategory 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Top             =   60
         Width           =   2550
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
         Tip             =   "frmFunctionMgr.frx":0340
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgr.frx":0360
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDesc 
         DataField       =   "Description"
         Height          =   720
         Left            =   900
         TabIndex        =   4
         Top             =   420
         Width           =   7860
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmFunctionMgr.frx":037C
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
         Tip             =   "frmFunctionMgr.frx":039C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgr.frx":03BC
      End
      Begin HexUniControls.ctlUniRichTextBoxXP TradeSense 
         Height          =   570
         Left            =   900
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1215
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   1005
         BackColor       =   12632256
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmFunctionMgr.frx":03D8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   -1
         MultiLine       =   -1  'True
         Alignment       =   0
         ScrollBars      =   3
         PasswordChar    =   ""
         TrapTab         =   0   'False
         RaiseChangeEvent=   -1  'True
         RaiseUpdateEvent=   0   'False
         RaiseSelChangeEvent=   -1  'True
         Tip             =   "frmFunctionMgr.frx":03F8
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgr.frx":0418
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
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   210
         Index           =   0
         Left            =   0
         Top             =   1215
         Width           =   900
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmFunctionMgr.frx":0434
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFunctionMgr.frx":045E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgr.frx":047E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Index           =   5
         Left            =   0
         Top             =   90
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmFunctionMgr.frx":049A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFunctionMgr.frx":04CA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgr.frx":04EA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   210
         Index           =   6
         Left            =   0
         Top             =   420
         Width           =   900
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmFunctionMgr.frx":0506
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFunctionMgr.frx":053C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgr.frx":055C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
      Height          =   4185
      Left            =   60
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   7382
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
      Caption         =   "&Function|&Arguments (Inputs)|Ad&vanced"
      Align           =   0
      Appearance      =   1
      CurrTab         =   0
      FirstTab        =   0
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
      Begin HexUniControls.ctlUniFrameWL fraAdvanced 
         Height          =   3810
         Left            =   9825
         TabIndex        =   1
         Top             =   330
         Width           =   8790
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmFunctionMgr.frx":0578
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmFunctionMgr.frx":05A4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgr.frx":05C4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtRequiredMod 
            Height          =   315
            Left            =   1740
            TabIndex        =   3
            Top             =   330
            Width           =   5715
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmFunctionMgr.frx":05E0
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
            Tip             =   "frmFunctionMgr.frx":0600
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":0620
         End
         Begin HexUniControls.ctlUniLabelXP lblRequiredMod 
            Height          =   255
            Left            =   180
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
            Caption         =   "frmFunctionMgr.frx":063C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFunctionMgr.frx":067E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":069E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraInputs 
         Height          =   3810
         Left            =   9525
         TabIndex        =   7
         Top             =   330
         Width           =   8790
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmFunctionMgr.frx":06BA
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmFunctionMgr.frx":06E6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgr.frx":0706
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraInputButtons 
            Height          =   1515
            Left            =   7320
            TabIndex        =   9
            Top             =   120
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
            Caption         =   "frmFunctionMgr.frx":0722
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmFunctionMgr.frx":074E
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":076E
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
               CausesValidation=   0   'False
               Height          =   435
               Left            =   0
               TabIndex        =   30
               Top             =   960
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
               Caption         =   "frmFunctionMgr.frx":078A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":07C4
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":07E4
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
               CausesValidation=   0   'False
               Height          =   435
               Left            =   0
               TabIndex        =   29
               Top             =   480
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
               Caption         =   "frmFunctionMgr.frx":0800
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":0836
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0856
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdNew 
               CausesValidation=   0   'False
               Height          =   435
               Left            =   0
               TabIndex        =   28
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
               Caption         =   "frmFunctionMgr.frx":0872
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":08A6
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":08C6
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraMoveButtons 
            Height          =   375
            Left            =   2340
            TabIndex        =   16
            Top             =   3300
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
            Caption         =   "frmFunctionMgr.frx":08E2
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmFunctionMgr.frx":090E
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":092E
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniButtonImageXP cmdMoveUp 
               Height          =   375
               Left            =   0
               TabIndex        =   26
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
               Caption         =   "frmFunctionMgr.frx":094A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":097A
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":099A
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdMoveDown 
               Height          =   375
               Left            =   1440
               TabIndex        =   27
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
               Caption         =   "frmFunctionMgr.frx":09B6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":09EA
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0A0A
               RightToLeft     =   0   'False
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid vsInputs 
            Height          =   3030
            Left            =   120
            TabIndex        =   25
            Top             =   105
            Width           =   7080
            _cx             =   12488
            _cy             =   5345
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
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
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
      Begin HexUniControls.ctlUniFrameWL fraFunction 
         Height          =   3810
         Left            =   45
         TabIndex        =   31
         Top             =   330
         Width           =   8790
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmFunctionMgr.frx":0A26
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmFunctionMgr.frx":0A52
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgr.frx":0A72
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraInternal 
            Height          =   555
            Left            =   300
            TabIndex        =   20
            Top             =   2820
            Width           =   8235
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmFunctionMgr.frx":0A8E
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmFunctionMgr.frx":0ABA
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":0ADA
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkInternal 
               Height          =   330
               Left            =   5655
               TabIndex        =   24
               Top             =   240
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
               Caption         =   "frmFunctionMgr.frx":0AF6
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":0B38
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0B58
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkLateCalculating 
               Height          =   330
               Left            =   0
               TabIndex        =   21
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
               Caption         =   "frmFunctionMgr.frx":0B74
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":0BB4
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0BD4
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkUsesOpenNextBar 
               Height          =   330
               Left            =   1755
               TabIndex        =   22
               Top             =   240
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
               Caption         =   "frmFunctionMgr.frx":0BF0
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":0C34
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0C54
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkUsesHLCNextBar 
               Height          =   330
               Left            =   3735
               TabIndex        =   23
               Top             =   240
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
               Caption         =   "frmFunctionMgr.frx":0C70
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":0CB2
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0CD2
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblLateCalc 
               Height          =   255
               Left            =   15
               Top             =   0
               Width           =   6330
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmFunctionMgr.frx":0CEE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmFunctionMgr.frx":0D34
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0D54
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraReturnType 
            Height          =   600
            Left            =   240
            TabIndex        =   32
            Tag             =   "4"
            Top             =   1305
            Width           =   8370
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmFunctionMgr.frx":0D70
            Enabled         =   -1  'True
            ForeColor       =   -2147483641
            BackColor       =   -2147483633
            Tip             =   "frmFunctionMgr.frx":0DAA
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":0DCA
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP opReturnType 
               CausesValidation=   0   'False
               Height          =   330
               Index           =   5
               Left            =   5340
               TabIndex        =   15
               Tag             =   "8"
               Top             =   240
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
               Caption         =   "frmFunctionMgr.frx":0DE6
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":0E38
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0E58
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP opReturnType 
               CausesValidation=   0   'False
               Height          =   330
               Index           =   4
               Left            =   5340
               TabIndex        =   14
               Tag             =   "2"
               Top             =   -30
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
               Caption         =   "frmFunctionMgr.frx":0E74
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":0EB6
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0ED6
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP opReturnType 
               CausesValidation=   0   'False
               Height          =   330
               Index           =   1
               Left            =   2640
               TabIndex        =   12
               Tag             =   "6"
               Top             =   -30
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
               Caption         =   "frmFunctionMgr.frx":0EF2
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":0F3C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0F5C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP opReturnType 
               CausesValidation=   0   'False
               Height          =   330
               Index           =   3
               Left            =   2640
               TabIndex        =   13
               Tag             =   "3"
               Top             =   240
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
               Caption         =   "frmFunctionMgr.frx":0F78
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":0FD0
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":0FF0
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP opReturnType 
               CausesValidation=   0   'False
               Height          =   330
               Index           =   0
               Left            =   0
               TabIndex        =   10
               Tag             =   "1"
               Top             =   -45
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
               Caption         =   "frmFunctionMgr.frx":100C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgr.frx":104E
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":106E
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP opReturnType 
               CausesValidation=   0   'False
               Height          =   330
               Index           =   2
               Left            =   0
               TabIndex        =   11
               Tag             =   "4"
               Top             =   240
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
               Caption         =   "frmFunctionMgr.frx":108A
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmFunctionMgr.frx":10DC
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgr.frx":10FC
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniTextBoxXP txtDLLName 
            DataField       =   "CodedName"
            Height          =   345
            Left            =   285
            TabIndex        =   8
            Top             =   465
            Width           =   3840
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmFunctionMgr.frx":1118
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
            Tip             =   "frmFunctionMgr.frx":1138
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":1158
         End
         Begin HexUniControls.ctlUniCheckXP chkSystemTesting 
            Height          =   375
            Left            =   285
            TabIndex        =   17
            Top             =   2325
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
            Caption         =   "frmFunctionMgr.frx":1174
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmFunctionMgr.frx":11B4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":11D4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkCharting 
            Height          =   375
            Left            =   2085
            TabIndex        =   18
            Top             =   2325
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
            Caption         =   "frmFunctionMgr.frx":11F0
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmFunctionMgr.frx":1220
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":1240
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkCriteria 
            Height          =   375
            Left            =   3885
            TabIndex        =   19
            Top             =   2325
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
            Caption         =   "frmFunctionMgr.frx":125C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmFunctionMgr.frx":1298
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":12B8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDllName 
            Height          =   285
            Left            =   270
            Top             =   240
            Width           =   6330
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmFunctionMgr.frx":12D4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFunctionMgr.frx":1376
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":1396
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblReturnType 
            Height          =   255
            Left            =   270
            Top             =   1065
            Width           =   6330
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmFunctionMgr.frx":13B2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFunctionMgr.frx":1422
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":1442
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblUsage 
            Height          =   255
            Left            =   285
            Top             =   2085
            Width           =   7410
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmFunctionMgr.frx":145E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFunctionMgr.frx":1526
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgr.frx":1546
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuAddInput 
         Caption         =   "&Add Input"
      End
      Begin VB.Menu mnuRemoveInput 
         Caption         =   "&Remove Input"
      End
      Begin VB.Menu mnuEditInput 
         Caption         =   "&Edit Input"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmFunctionMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmFunctionMgr.frm
'' Description: Allows the user to edit a DLL function
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    lWidth As Long
    lHeight As Long
    
    Function As cFunction
    FunctionCategories As cFunctionCategories
    InputTypes As cParmTypes
    bReturnValue As Boolean
    strName As String
End Type
Private m As mPrivate

'Inputs columns
Private Enum eIGCols
    eIGCol_InputID = 0
    eIGCol_InputName = 1
    eIGCol_InputDesc = 2
    eIGCol_InputType = 3
    eIGCol_Req = 4
    eIGCol_Default = 5
    eIGCol_FromVal = 6
    eIGCol_ToVal = 7
    eIGCol_Order = 8
    eIGCol_InputTypeID = 9
    eIGCol_ListType = 10
    eIGCol_ListTypeID = 11
    eIGCol_FillPre = 12
    eIGCol_FillPost = 13
End Enum
Private Const kInputsGridCols = 14

'Usage masks
Private Enum eUsageMask
    eUsageMask_MM = 1
    eUsageMask_SystemTesting = 2
    eUsageMask_Charting = 4
    eUsageMask_Criteria = 8
End Enum

Public Property Get ID() As Long
    ID = m.Function.FunctionID
End Property

Private Function IGCol(ByVal lColumn As eIGCols) As Long
    IGCol = lColumn
End Function
Private Function UsageMask(ByVal lUsageMask As eUsageMask) As Long
    UsageMask = lUsageMask
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Gets the function manager ready to add a new function
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add()
On Error GoTo ErrSection:
    
    ' Clear the function manager and get a new object
    ClearFunction
    Set m.Function = New cFunction
    m.Function.FunctionID = 0
    
    ' Default the Usage flags to all on
    chkCriteria.Value = vbChecked
    chkCharting.Value = vbChecked
    chkSystemTesting.Value = vbChecked
    cbCategory.Text = "Indicator"
    
    vsIndexTab1.CurrTab = 0
    
    SetEditorCaption Me, "Function", ""
    
    InitInputsGrid
    TradeSense.Text = ""
    txtRequiredMod.Text = ""
    EnableToolbar False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.Add", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableToolbar
'' Description: Turn save button on/off as a dirty flag
'' Inputs:      Whether to turn the button on or off
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableToolbar(ByVal bEnabled As Boolean)
On Error GoTo ErrSection:
    
    tbToolbar.Tools("ID_Save").Enabled = bEnabled
    tbToolbar.Tools("ID_SaveAs").Enabled = (Trim(m.strName) <> "")
    tbToolbar.Tools("ID_Rename").Enabled = (Trim(m.strName) <> "")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.EnableToolbar", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadRec
'' Description: Load the function manager with an existing function
'' Inputs:      ID of the Function to Load
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadRec(pKey As Long) As Boolean
On Error GoTo ErrSection:
    
    'Load the function from collection to form...
    Set m.Function = New cFunction
    With m.Function
        .FunctionID = Val(pKey)
        .Load
        
        If Not g.Security.CanEdit(.SecurityLevel, .Password) Then
            GoTo ErrExit:
        End If
        
        ClearFunction
        
        m.strName = .FunctionName
        txtDesc.Text = .Description
        txtDLLName = .CodedName
        If .LateCalculating Then Me.chkLateCalculating.Value = 1 Else chkLateCalculating.Value = 0
        cbCategory.Text = m.FunctionCategories.Item(CStr(.FunctionCategoryID)).FunctionCategory
        SetReturnType .DataTypeID
        
        If .Usage And UsageMask(eUsageMask_Charting) Then chkCharting.Value = vbChecked
        If .Usage And UsageMask(eUsageMask_Criteria) Then chkCriteria.Value = vbChecked
        If .Usage And UsageMask(eUsageMask_SystemTesting) Then chkSystemTesting.Value = vbChecked
        
        If .UsesOpenNextBar Then chkUsesOpenNextBar = vbChecked Else chkUsesOpenNextBar = vbUnchecked
        If .UsesNextBarHLC Then chkUsesHLCNextBar = vbChecked Else chkUsesHLCNextBar = vbUnchecked
        
        If .ImplementationTypeID = 3 Then chkInternal.Value = vbChecked Else chkInternal.Value = vbUnchecked
        
        txtRequiredMod.Text = .RequiredMod
    End With
    
    SetEditorCaption Me, "Function", m.strName
    
    InitInputsGrid
    ShowParmLine TradeSense
    m.bReturnValue = LockWindowUpdate(0)
    EnableToolbar False
    LoadRec = True
    
ErrExit:
    m.bReturnValue = LockWindowUpdate(0)
    Exit Function

ErrSection:
    RaiseError "frmFunctionMgr.LoadRec", eGDRaiseError_Raise
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetReturnType
'' Description: Set the appropriate option button depending on the return type
'' Inputs:      Return Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnType(pType As Byte)
On Error GoTo ErrSection:

    Dim X   As Integer                  ' Index for a for loop
    
    For X = 0 To 5
        If opReturnType(X).Tag = pType Then
            opReturnType(X) = True
            Exit Sub
        End If
    Next X

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.SetReturnType", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the function to the database
'' Inputs:      Whether or not to show success message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save(ByVal bShowMsg As Boolean, ByVal strButton As String)
On Error GoTo ErrSection:
    
    Dim X       As Integer              ' Index into a for loop
    Dim Y       As Integer              ' Index into a for loop
    Dim Expr    As cExpression          ' Expression for coded text
    Dim lUsage  As Long                 ' Usage flags
    Dim strText As String               ' Text to show in a message box
    Dim strNewName As String            ' Return from the message box
    Dim bSaveAs As Boolean              ' Are we in SaveAs mode?
    Dim strError As String
    Dim lOldID As Long
    
    ' Verify DLL name
    If DLLNameExists(txtDLLName.Text) Then
        InfBox "A function with the DLL Name '" & txtDLLName.Text & "' already exists", "e", , "Error"
        MoveFocus txtDLLName
        Exit Sub
    End If
    
    ' Handle Rename/Save As
    m.strName = Trim(m.strName)
    If Len(m.strName) = 0 Then
        strText = "Save the current Function as..."
        strNewName = AskBox("h=Save ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    ElseIf strButton = "ID_SaveAs" Then
        strText = "Save a copy of the current Function as..."
        strNewName = AskBox("h=Save As ; i=? ; g=string ; d=" & m.strName & " #02" & " ; " & strText)
        If Trim(UCase(strNewName)) <> UCase(m.strName) Then
            bSaveAs = True
        End If
    ElseIf strButton = "ID_Rename" Then
        strText = "Rename the current Function as..."
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    Else
        strNewName = m.strName
    End If
    
    ' Verify that it is a good name
    Do While Len(Trim(strNewName)) > 0
        ' Strip out a colon if it exists in the name...
        If InStr(strNewName, ":") Then strNewName = Replace(strNewName, ":", "")
        strError = m.Function.ValidName(strNewName)
        If strError <> "" Then
            InfBox strError, "e", , "Error"
        ElseIf FunctionExists(strNewName) Then
            InfBox "'" & strNewName & "' already exists.", "e", , "Error"
        Else
            ' Name is OK so we can exit loop
            Exit Do
        End If
        strText = "Rename the current Function as..."
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & strNewName & " ; " & strText)
    Loop
    
    If Len(Trim(strNewName)) = 0 Then
        Exit Sub 'Err.Raise vbObjectError + 1000, , "You must enter in a name for the filter"
    End If
    m.strName = Trim(strNewName)
    SetEditorCaption Me, "Function", m.strName
        
    'If bSaveAs Then m.Function.FunctionID = 0
    If bSaveAs Then
        lOldID = m.Function.FunctionID
        Set m.Function = New cFunction
        m.Function.FunctionID = lOldID
        m.Function.Load
        m.Function.FunctionID = 0&
    End If
        
    m.bReturnValue = LockWindowUpdate(Me.hWnd)
    Screen.MousePointer = vbHourglass
        
    'Validate function fields
    With m.Function
    
        'User must be authorized to save (don't prompt for new functions or if
        'copying an existing function
        If .FunctionID = 0 Then
            .SecurityLevel = 0
            .CannotDelete = False
            .LibraryID = kSN_UserLibrary
            .Password = ""
        Else
            If Not g.Security.CanSave(.SecurityLevel, .Password) Then
                GoTo ErrExit:
            End If
        End If
        
        .FunctionName = m.strName
        .Description = txtDesc.Text
        .FunctionCategoryID = GetCatID(cbCategory.Text)
        .CodedName = txtDLLName.Text
        If chkInternal.Value = vbChecked Then
            .ImplementationTypeID = kSN_Internal
        Else
            .ImplementationTypeID = kSN_BuiltIn
        End If
        .DataTypeID = fraReturnType.Tag
        .ReturnTypeID = fraReturnType.Tag
        
        .LateCalculating = (chkLateCalculating.Value = vbChecked)
        .UsesOpenNextBar = (chkUsesOpenNextBar.Value = vbChecked)
        .UsesNextBarHLC = (chkUsesHLCNextBar.Value = vbChecked)
         
         lUsage = 0&
         If chkSystemTesting.Value = vbChecked Then lUsage = lUsage Or UsageMask(eUsageMask_SystemTesting)
         If chkCriteria.Value = vbChecked Then lUsage = lUsage Or UsageMask(eUsageMask_Criteria)
         If chkCharting.Value = vbChecked Then lUsage = lUsage Or UsageMask(eUsageMask_Charting)
         .Usage = lUsage
        
        .RequiredMod = txtRequiredMod.Text
    End With
    
#If 0 Then
    'Save inputs to table (already validated)
    For X = 1 To vsInputs.Rows - 1
        With vsInputs
            m.Function.SaveInput ValOfText(.TextMatrix(X, IGCol(eIGCol_InputID))), _
                ValOfText(.TextMatrix(X, IGCol(eIGCol_Order))), .TextMatrix(X, IGCol(eIGCol_InputName)), _
                .TextMatrix(X, IGCol(eIGCol_InputDesc)), ValOfText(.TextMatrix(X, IGCol(eIGCol_InputTypeID))), _
                .TextMatrix(X, IGCol(eIGCol_Default)), ValOfText(.TextMatrix(X, IGCol(eIGCol_FromVal))), _
                ValOfText(.TextMatrix(X, IGCol(eIGCol_ToVal))), ValOfText(.TextMatrix(X, IGCol(eIGCol_Req))), _
                ValOfText(.TextMatrix(X, IGCol(eIGCol_ListTypeID))), .TextMatrix(X, IGCol(eIGCol_FillPre)), _
                .TextMatrix(X, IGCol(eIGCol_FillPost))
        End With
    Next X
        
    'Remove old inputs
    For X = m.Function.Inputs.Count To 1 Step -1
        For Y = 1 To vsInputs.Rows - 1
            If m.Function.Inputs.Item(X).ParmID = vsInputs.TextMatrix(Y, IGCol(eIGCol_InputID)) Then
                m.Function.RemoveInput m.Function.Inputs.Item(X).ParmID
            End If
        Next Y
    Next X
#End If
        
    SaveOrder
        
    Set Expr = New cExpression
    With Expr
        .Functions = g.Functions
        .PortfolioNavigator = False
    End With
    
    With m.Function
        .CodedText = "N/A" 'Expr.BuiltinCodedText(.CodedName, .DataTypeID, .Inputs)
            
        ShowParmLine TradeSense
        .TradeSenseUsage = TradeSense.Tag
        .Reverify = False
        ' update LastModified (unless skip file exists -- e.g. when updating master mdb)
        If .LastModified <= 0 Or Not FileExist(App.Path & "\LastMod.SKP") Then
            .LastModified = Now
        Else
            StatusMsg "LastModified not changed"
        End If
        .Save
        g.bDirtyLibrariesMDB = True
    End With
    
    InitInputsGrid
    RefreshFunction m.Function
    RefreshReverify
    Screen.MousePointer = vbDefault
    
    EnableToolbar False
    'Unload Me
    
ErrExit:
    Set Expr = Nothing
    Screen.MousePointer = vbDefault
    m.bReturnValue = LockWindowUpdate(0)
    Exit Sub

ErrSection:
    Set Expr = Nothing
    Screen.MousePointer = vbDefault
    m.bReturnValue = LockWindowUpdate(0)
    RaiseError "frmFunctionMgr.Save", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetCatID
'' Description: Get a Category ID for a Category Name
'' Inputs:      Category Name
'' Returns:     Category ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetCatID(pName As String) As Long
On Error GoTo ErrSection:

    Dim X       As Integer              ' Index for a for loop
    
    GetCatID = 0
    For X = 1 To m.FunctionCategories.Count
        With m.FunctionCategories.Item(X)
            If .FunctionCategory = pName Then
                GetCatID = .FunctionCategoryID
                Exit For
            End If
        End With
    Next X

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFunctionMgr.GetCatID", eGDRaiseError_Raise
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowParmLine
'' Description: Show a Preview of the Parameter Line
'' Inputs:      Rich Text Box to put the Preview in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowParmLine(pControl As RichTextBox)
On Error GoTo ErrSection:

    Dim TradeSenseText  As String
    Dim LeftParen       As Integer
    Dim Lentext         As Integer
    Dim X               As Integer
    Dim Inputs          As Byte
    Dim fName           As String
    
    'Default the function name to beginning of TradeSense text
    TradeSenseText = m.strName & " ("
    LeftParen = 0
    Inputs = 0
    
    'Build function's parameters
#If 0 Then
    With m.Function
        If Not .Inputs Is Nothing Then
            If .Inputs.Count > 0 Then
                For X = 1 To .Inputs.Count
                    If .Inputs.Item(X).ParmTypeID <> kSN_RetTrades And _
                       .Inputs.Item(X).ParmTypeID <> kSN_RetBars Then
                        TradeSenseText = _
                            TradeSenseText & .Inputs.Item(X).ParmName & ", "
                        Inputs = Inputs + 1
                    End If
                Next X
            End If
        End If
    End With
#End If
    With vsInputs
        For X = .FixedRows To .Rows - 1
            If ValOfText(.TextMatrix(X, IGCol(eIGCol_InputTypeID))) <> kSN_RetTrades And _
                ValOfText(.TextMatrix(X, IGCol(eIGCol_InputTypeID))) <> kSN_RetBars Then
                    TradeSenseText = TradeSenseText & .Cell(flexcpText, X, IGCol(eIGCol_InputName)) & ", "
                    Inputs = Inputs + 1
            End If
        Next X
    End With
    
    'Add parm right paren to end of string
    If Inputs > 0 Then
        TradeSenseText = Left(TradeSenseText, Len(TradeSenseText) - 2) + ")"
        LeftParen = InStr(1, TradeSenseText, "(")
        Lentext = Len(TradeSenseText) - Len(Trim(m.strName))
    Else
        TradeSenseText = Left(TradeSenseText, Len(TradeSenseText) - 2)
    End If
    
    'Simulate text entered into RTF box...
    With pControl
        .Tag = TradeSenseText
        .Text = TradeSenseText
        If LeftParen > 0 Then
            .SelStart = LeftParen
            .SelLength = Lentext
            .SelItalic = True
            .SelLength = 0
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.ShowParmLine", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

Private Sub chkInternal_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.chkInternal.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkUsesHLCNextBar_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.chkUsesHLCNextBar.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkUsesOpenNextBar_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.chkUsesOpenNextBar.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: When the user clicks on the Edit button, allow them to Edit
''              the current Function Argument
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:
    
    EditInput

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNew_Click
'' Description: If the user clicks on the New button, allow them to add a New
''              Function Argument
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    AddInput
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.cmdNew.Click", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemove_Click
'' Description: If the user clicks on the Remove button, Remove the current
''              Function Argument
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    RemoveInput
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.cmdRemove.Click", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cbCategory_Click
'' Description: If the user clicks on the Category Combo Box, set the Dirty Flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cbCategory_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.cbCategory.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCharting_Click
'' Description: If the user clicks on the Charting Check Box, set the Dirty
''              Flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCharting_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.chkCharting.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCriteria_Click
'' Description: If the user clicks on the Criteria Check Box, set the Dirty
''              Flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCriteria_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.chkCriteria.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkLateCalculating_Click
'' Description: If the user clicks on the Late Calculating Check Box, set the
''              Dirty Flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkLateCalculating_Click()
On Error GoTo ErrSection:
    
    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.chkLateCalculating.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkSystemTesting_Click
'' Description: If the user clicks on the System Testing Check Box, set the
''              Dirty Flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkSystemTesting_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.chkSystemTesting.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdMoveDown_Click
'' Description: If the user clicks on the Move Down button, move the currently
''              selected input down one row in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdMoveDown_Click()
On Error GoTo ErrSection:

    With vsInputs
        If .RowSel = 1 And .TextMatrix(1, IGCol(eIGCol_InputTypeID)) = "5" Then Exit Sub
        If .RowSel > .FixedRows - 1 And .RowSel < .Rows - 1 Then
            .RowPosition(.RowSel) = .RowSel + 1
            .Row = .RowSel + 1
            .RowSel = .Row
        End If
    End With
    
    EnableToolbar True
    MoveFocus vsInputs

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.cmdMoveDown.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdMoveUp_Click
'' Description: If the user clicks on the Move Up button, move the currently
''              selected input up one row in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdMoveUp_Click()
On Error GoTo ErrSection:

    With vsInputs
        If .RowSel = 2 And .TextMatrix(1, IGCol(eIGCol_InputTypeID)) = "5" Then Exit Sub
        If .RowSel > .FixedRows Then ' + 1 Then
            .RowPosition(.RowSel) = .RowSel - 1
            .Row = .RowSel - 1
            .RowSel = .Row
        End If
    End With
    
    EnableToolbar True
    MoveFocus vsInputs

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.cmdMoveUp.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form gets activated, set the Dirty Flag off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    
    If g.Functions Is Nothing Then
        InitFunctions
    End If
    
    ' Quickly check the Reverify flag.  If on then force a reverify...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
        "WHERE [FunctionID]=" & m.Function.FunctionID & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblFunctions"
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        
        If rs!CheckSum = 0.5 Then
            EnableToolbar False
            Unload Me
            Err.Raise vbObjectError + 1000, , "This Function is no longer Valid"
        End If
        
        If rs!Reverify Then EnableToolbar True
    End If

ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.Form.Activate", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    Else
        frmMain.DockPro_ShortcutKeyDown KeyCode, Shift, Me.Name
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, do some initialization with it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim strText$
    Dim X As Long
    Dim strFont As String

    g.Styler.StyleForm Me
    
    Me.Icon = Picture16(ToolbarIcon("ID_Functions"), , True)
    With tbToolbar
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
    End With
        
    m.lWidth = Me.Width
    m.lHeight = Me.vsIndexTab1.Height + Me.fraFunctionInfo.Height + Me.fraFunctionInfo.Top * 3 _
                    + Me.Height - Me.ScaleHeight
    Me.Height = m.lHeight
    CenterTheForm Me

    vsIndexTab1.CurrTab = 0
    
    ' Only show the required module controls if running from IDE...
    vsIndexTab1.TabVisible(2) = IsIDE

    strText = GetIniFileProperty("frmFunctionMgr", "", "Placement", g.strIniFile)
    SetFormPlacement Me, strText, "LT"
    
    ' Set the grid font from the INI file...
    strFont = GetIniFileProperty("FunctionMgr", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString vsInputs.Font, strFont
    
    ' Only show the internal stuff if a Genesis user...
    fraInternal.Visible = FileExist("C:\Common\Files.EXE")

#If 0 Then
    If cExpression Is Nothing Then
        Set cExpression = New cExpression
        With cExpression
            .SourceID = kSN_Both
            .cFunctionsRef = g.Functions
        End With
    End If
#End If

    m.strName = ""
    
    Set m.InputTypes = New cParmTypes
    m.InputTypes.Load
    
    Set m.FunctionCategories = New cFunctionCategories
    m.FunctionCategories.Load
    For X = 1 To m.FunctionCategories.Count
        With m.FunctionCategories.Item(X)
            If .FunctionCategory <> "Reserved" And .FunctionCategory <> "Actions" Then
                cbCategory.AddItem m.FunctionCategories.Item(X).FunctionCategory
            ElseIf FileExist("C:\Common\Files.EXE") Then
                cbCategory.AddItem m.FunctionCategories.Item(X).FunctionCategory
            End If
        End With
    Next X
    
    mnuPopUp.Visible = False
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next

    If m.lHeight > 0 Then
        If Me.Height <> m.lHeight Or Me.Width <> m.lWidth Then
            Me.Move Me.Left, Me.Top, m.lWidth, m.lHeight
            Exit Sub
        End If
    End If

    'With vsIndexTab1
    '    .Move .Left, (fraFunctionInfo.Top * 2) + fraFunctionInfo.Height
    'End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, save the placement and category id
'' Inputs:      Whether or not to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    SetIniFileProperty "frmFunctionMgr", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "FunctionCategoryID", cbCategory.Text, "Misc", g.strIniFile
    SetIniFileProperty "FunctionMgr", FontToString(vsInputs.Font), "Fonts", g.strIniFile
    
    Set m.Function = Nothing
    Set m.FunctionCategories = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearFunction
'' Description: Clear the controls on the Function Manager form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearFunction()
On Error GoTo ErrSection:

    Dim CatText     As String
    
    m.strName = ""
    txtDesc.Text = ""
    
    'Get INI Defaults
    CatText = GetIniFileProperty("FunctionCategory", "", "Misc", g.strIniFile)
    If Len(CatText) > 0 Then
        cbCategory.Text = CatText
    End If
    
    'Clear Function tabs
    txtDLLName.Text = ""
    chkCriteria.Value = vbUnchecked
    chkCharting.Value = vbUnchecked
    chkSystemTesting.Value = vbUnchecked
    chkLateCalculating.Value = vbUnchecked
    
    ' Clear out the internal stuff...
    chkUsesOpenNextBar.Value = vbUnchecked
    chkUsesHLCNextBar.Value = vbUnchecked
    chkLateCalculating.Value = vbUnchecked
    chkInternal.Value = vbUnchecked
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.ClearFunction", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user wants to close the form and the function is dirty,
''              ask if they would like to save the function
'' Inputs:      Whether or not to Cancel the Unload, Unload Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    Dim strResponse As String
    
    If UnloadMode <> vbFormCode Then
        Cancel = AskToSave
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub lblDllName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    If Button = 2 Then
        fraInternal.Visible = True
    End If

End Sub

Private Sub mnuAddInput_Click()
On Error GoTo ErrSection:

    AddInput
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.mnuAddInput.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont vsInputs, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuEditInput_Click()
On Error GoTo ErrSection:

    EditInput
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.mnuEditInput.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuRemoveInput_Click()
On Error GoTo ErrSection:

    RemoveInput
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.mnuRemoveInput.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    opReturnType_Click
'' Description: If the user clicks on one of the Return Type Option Buttons,
''              set the dirty flag
'' Inputs:      Which option button was clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub opReturnType_Click(Index As Integer)
On Error GoTo ErrSection:
    
    EnableToolbar True
    fraReturnType.Tag = opReturnType(Index).Tag

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.opReturnType.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle the toolbar button click appropriately
'' Inputs:      Toolbar Clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim strID$

    ToggleFocus Me, Me.vsIndexTab1

    Select Case Tool.ID
        Case "ID_Save", "ID_SaveAs", "ID_Rename"
            Save True, Tool.ID
        
        Case "ID_Toolbox"
            If Not AskToSave Then
                strID = CStr(m.Function.FunctionID)
                Unload Me
                frmToolbox.ShowMe eTab_Functions, strID
            End If
        
        Case "ID_Close"
            If Not AskToSave Then
                Unload Me
            End If
    
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDesc_Change
'' Description: If the user changes the description, set the Dirty Flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDesc_Change()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.txtDesc.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDLLName_Change
'' Description: If the user changes the DLL Name, set the Dirty Flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDLLName_Change()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.txtDLLName.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HighlightParm
'' Description: Highlight the appropriate parameter
'' Inputs:      Which parameter to highlight
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HighlightParm(pInputName As String)
On Error GoTo ErrSection:
    
    Static hParm        As String
    Dim xSelStart       As Integer
    Dim iStart          As Integer
    
    'Unhighlight current parm
    If Len(hParm) > 0 Then
        With TradeSense
            xSelStart = .Find(hParm)
            If xSelStart > 0 Then
                .SelStart = xSelStart
                .SelLength = Len(hParm)
                .SelBold = False
                .SelItalic = False
                .SelLength = 0
            End If
        End With
    End If
    
    'Highlight current parm
    If Len(TradeSense) > 0 Then
        With TradeSense
            iStart = .Find(pInputName)
            If iStart <> -1 Then
                .SelStart = iStart
                .SelLength = Len(pInputName)
                .SelItalic = True
                .SelBold = True
                .SelLength = 0
            End If
        End With
        hParm = pInputName
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.HighlightParm", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FunctionExists
'' Description: Determine whether the function name already exists
'' Inputs:      Function Name to check
'' Returns:     True if it exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FunctionExists(ByVal strFuncName As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim QryDef As QueryDef              ' Query into the database
    
    FunctionExists = False
    
    Set QryDef = g.dbNav.QueryDefs("qryFunctionIDFromName")
    QryDef.Parameters(0).Value = strFuncName
    Set rs = QryDef.OpenRecordset
    
    If rs.RecordCount <> 0 Then
        If rs!FunctionID <> m.Function.FunctionID Then
            FunctionExists = True
        End If
    End If
    
ErrExit:
    Set rs = Nothing
    Set QryDef = Nothing
    Exit Function

ErrSection:
    RaiseError "frmFunctionMgr.FunctionExists", eGDRaiseError_Raise
    Resume ErrExit:

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitInputsGrid
'' Description: Initialize the inputs grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitInputsGrid()
On Error GoTo ErrSection:

    Dim xInput As cInput
    Dim X As Long
    
    With vsInputs
        .Redraw = False
        .AllowBigSelection = False
        .AllowSelection = True
        .HighLight = flexHighlightWithFocus
        .TabBehavior = flexTabCells
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .ExplorerBar = flexExMoveRows
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .Ellipsis = flexEllipsisEnd
        .Cols = kInputsGridCols
        .Rows = 1
        .FixedCols = 1
        .FixedRows = 1
        
        .ColHidden(IGCol(eIGCol_InputID)) = True
        '.ColHidden(IGCol(eIGCol_InputDesc)) = True
        .ColHidden(IGCol(eIGCol_Order)) = True
        .ColHidden(IGCol(eIGCol_InputTypeID)) = True
        .ColHidden(IGCol(eIGCol_ListTypeID)) = True
        .ColHidden(IGCol(eIGCol_FillPre)) = True
        .ColHidden(IGCol(eIGCol_FillPost)) = True
        
        .TextMatrix(0, IGCol(eIGCol_InputName)) = "Input"
        .TextMatrix(0, IGCol(eIGCol_Default)) = "Default"
        .TextMatrix(0, IGCol(eIGCol_InputDesc)) = "Description"
        .ColWidth(IGCol(eIGCol_Default)) = 1200
        .TextMatrix(0, IGCol(eIGCol_InputType)) = "Type"
        .TextMatrix(0, IGCol(eIGCol_Req)) = "Required"
        .TextMatrix(0, IGCol(eIGCol_FromVal)) = "Min"
        .TextMatrix(0, IGCol(eIGCol_ToVal)) = "Max"
        
        .ColAlignment(IGCol(eIGCol_InputName)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_Default)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_InputType)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_FromVal)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_ToVal)) = flexAlignLeftCenter
        .ColDataType(IGCol(eIGCol_Req)) = flexDTBoolean
        
        .AutoSize 0, .Cols - 1
    End With
    
    '6/2001 TradeSense compatilbility
    With vsInputs
        .ColHidden(IGCol(eIGCol_FromVal)) = True
        .ColHidden(IGCol(eIGCol_ToVal)) = True
        .ColHidden(IGCol(eIGCol_FillPost)) = True
    End With
    
    With vsInputs
        .Redraw = flexRDNone
        .Rows = m.Function.Inputs.Count + .FixedRows
        
        For X = 1 To m.Function.Inputs.Count
            Set xInput = m.Function.Inputs.Item(X)
            .TextMatrix(X, IGCol(eIGCol_InputID)) = xInput.ParmID
            .TextMatrix(X, IGCol(eIGCol_InputName)) = Left(xInput.ParmName, 40)
            .Cell(flexcpFontBold, X, IGCol(eIGCol_InputName)) = True
            .Cell(flexcpForeColor, X, IGCol(eIGCol_InputName)) = vbBlack
            .TextMatrix(X, IGCol(eIGCol_Req)) = xInput.Required
            .TextMatrix(X, IGCol(eIGCol_Order)) = xInput.ParmSeq
            .TextMatrix(X, IGCol(eIGCol_InputType)) = m.InputTypes.Item(CStr(xInput.ParmTypeID)).ParmType
            .TextMatrix(X, IGCol(eIGCol_InputTypeID)) = m.InputTypes.Item(CStr(xInput.ParmTypeID)).ParmTypeID
            .TextMatrix(X, IGCol(eIGCol_InputDesc)) = xInput.ParmDesc
            .TextMatrix(X, IGCol(eIGCol_FromVal)) = ""
            .TextMatrix(X, IGCol(eIGCol_ToVal)) = ""
            
            'Set the value (or default if one doesn't exist).  The bars and
            'trades type structure is always "Market1" and "Trades"
            Select Case xInput.ParmTypeID
            
                Case kSN_RetBars
                    .TextMatrix(X, IGCol(eIGCol_Default)) = xInput.ParmName
                    .RowPosition(X) = 1
            
                Case kSN_RetTrades
                    .TextMatrix(X, IGCol(eIGCol_Default)) = xInput.ParmName
                    
                Case kSN_RetNumericConstant
                    .TextMatrix(X, IGCol(eIGCol_Default)) = FormatNum(Val(xInput.DefaultValue))
                    ColorCell X, IGCol(eIGCol_Default)
                    
                Case kSN_RetNumeric
                    .TextMatrix(X, IGCol(eIGCol_Default)) = xInput.DefaultValue
                    
                Case kSN_RetTrueFalse, kSN_RetTrueFalseConstant
                    .TextMatrix(X, IGCol(eIGCol_Default)) = xInput.DefaultValue
                    
                Case Else
                    .TextMatrix(X, IGCol(eIGCol_Default)) = xInput.DefaultValue
                
            End Select
        Next X
        
        For X = .FixedRows To .Rows - 1
            If X = .FixedRows Then
                .TextMatrix(X, IGCol(eIGCol_Req)) = True
            ElseIf X = .FixedRows + 1 And .TextMatrix(.FixedRows, IGCol(eIGCol_Default)) = "Market1" Then
                .TextMatrix(X, IGCol(eIGCol_Req)) = True
            ElseIf .TextMatrix(X - 1, IGCol(eIGCol_Req)) = False Then
                .TextMatrix(X, IGCol(eIGCol_Req)) = False
            End If
        Next X
        
        .AutoSize 0, .Cols - 1
        cmdRemove.Enabled = .Rows > 1
        .Redraw = flexRDBuffered
    End With

    If vsInputs.Rows > 1 Then
        cmdEdit.Enabled = True
        cmdRemove.Enabled = True
        With vsInputs
            .Row = .FixedRows
            .RowSel = .Row
        End With
    Else
        cmdEdit.Enabled = False
        cmdRemove.Enabled = False
        cmdMoveUp.Enabled = False
        cmdMoveDown.Enabled = False
    End If
        
    ShowParmLine TradeSense
    
ErrExit:
    Set xInput = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.InitInputsGrid", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorCell
'' Description: If the number is negative, color the cell red, otherwise black
'' Inputs:      Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorCell(pRow As Long, pCol As Long)
On Error GoTo ErrSection:
    
    If ValOfText(vsInputs.TextMatrix(pRow, pCol)) < 0 Then
        vsInputs.Cell(flexcpForeColor, pRow, pCol) = vbRed
    Else
        vsInputs.Cell(flexcpForeColor, pRow, pCol) = vbBlack
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.ColorCell", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

Private Sub txtRequiredMod_Change()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.txtRequiredMod.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_AfterEdit
'' Description: After a user has edited a column, make sure that the required
''              flags are correct and save the information to the collection
'' Inputs:      Row and Column that was edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    Select Case Col
        Case IGCol(eIGCol_Req)
            With vsInputs
                If Row = 1 Then
                    .TextMatrix(Row, Col) = True
                    m.Function.Inputs.Item(.TextMatrix(Row, IGCol(eIGCol_InputID))).Required = True
                ElseIf Row = 2 And .TextMatrix(1, IGCol(eIGCol_InputTypeID)) = "5" Then
                    .TextMatrix(Row, Col) = True
                    m.Function.Inputs.Item(.TextMatrix(Row, IGCol(eIGCol_InputID))).Required = True
                Else
                    If .TextMatrix(Row, Col) = False Then
                        For lIndex = Row To .Rows - 1
                            .TextMatrix(lIndex, IGCol(eIGCol_Req)) = False
                            m.Function.Inputs.Item(.TextMatrix(Row, IGCol(eIGCol_InputID))).Required = False
                        Next lIndex
                    Else
                        For lIndex = Row To 1 Step -1
                            .TextMatrix(lIndex, IGCol(eIGCol_Req)) = True
                            m.Function.Inputs.Item(.TextMatrix(Row, IGCol(eIGCol_InputID))).Required = True
                        Next lIndex
                    End If
                End If
            End With
        
        Case IGCol(eIGCol_Default)
            With m.Function.Inputs.Item(vsInputs.TextMatrix(Row, IGCol(eIGCol_InputID)))
                .DefaultValue = vsInputs.TextMatrix(Row, Col)
            End With
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.vsInputs.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_AfterMoveRow
'' Description: After the user moves a row, make sure that the row is still the
''              one selected
'' Inputs:      Row moved, Position moved to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_AfterMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:
    
    With vsInputs
        EnableToolbar True
        .Row = Position
        .RowSel = Position
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.vsInputs.AfterMoveRow", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_AfterRowColChange
'' Description: As the user changes rows in the grid, Highlight the appropriate
''              parameter
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With vsInputs
        If .Row <> -1 Then
            ShowParmLine TradeSense
            HighlightParm .TextMatrix(.Row, IGCol(eIGCol_InputName))
            cmdMoveUp.Enabled = .Row > .FixedRows
            cmdMoveDown.Enabled = .Row > .FixedRows - 1 And .Row < .Rows - 1
            If .Col <> IGCol(eIGCol_Req) Then .EditCell
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.vsInputs.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_BeforeEdit
'' Description: Only allow the user to edit the first column
'' Inputs:      Row and Column user is trying to edit, Whether or not to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> IGCol(eIGCol_Req) And Col <> IGCol(eIGCol_Default) Then Cancel = True
    If Col = IGCol(eIGCol_Default) And Row = 1 And vsInputs.TextMatrix(Row, IGCol(eIGCol_InputName)) = "Market1" Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.vsInputs.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_BeforeMouseDown
'' Description: When the user presses the mouse button, get ready for a drag
'' Inputs:      Mouse button pressed, Shift/Ctrl/Alt status, X and Y location,
''              Whether or not to cancel the mouse down
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of new row
    Dim lRow As Long                    ' Row being moved
    
    With vsInputs
        lRow = .MouseRow
        If lRow >= .FixedRows And lRow < .Rows Then
            If Button = vbLeftButton Then
                .Row = lRow
                .RowSel = lRow
            
                .Refresh
                lPos = .DragRow(lRow)
                If lPos <> lRow Then
                    Cancel = True
                End If
            ElseIf Button = vbRightButton Then
                .Row = lRow
                .RowSel = lRow
                
                Cancel = True
                PopupMenu mnuPopUp
            Else
                Cancel = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.vsInputs.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_BeforeMoveRow
'' Description: If the first row is a Market type, do not let the user move it
'' Inputs:      Row that is being moved, Position being moved to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_BeforeMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    With vsInputs
        If Position = 1 And .TextMatrix(1, IGCol(eIGCol_InputTypeID)) = 5 Then Position = Row
        If Row = 1 And .TextMatrix(1, IGCol(eIGCol_InputTypeID)) = 5 Then Position = 1
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.vsInputs.BeforeMoveRow", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsInputs_ChangeEdit()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgr.vsInputs.ChangeEdit", eGDRaiseError_Show
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_DblClick
'' Description: If the user double clicks on a parameter in the grid, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_DblClick()
On Error GoTo ErrSection:

    If vsInputs.MouseRow <> -1 Then cmdEdit_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.vsInputs.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveOrder
'' Description: Save the order of the inputs
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveOrder()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With vsInputs
        For lIndex = .FixedRows To .Rows - 1
            .TextMatrix(lIndex, IGCol(eIGCol_Order)) = lIndex
            m.Function.Inputs.Item(.TextMatrix(lIndex, IGCol(eIGCol_InputID))).ParmSeq = lIndex
            
            If .TextMatrix(lIndex, IGCol(eIGCol_Req)) = False And .TextMatrix(lIndex, IGCol(eIGCol_Default)) = "" Then
                Select Case CLng(ValOfText(.TextMatrix(lIndex, IGCol(eIGCol_InputTypeID))))
                    Case 1: ' Single Number
                        .TextMatrix(lIndex, IGCol(eIGCol_InputTypeID)) = 0
                        m.Function.Inputs.Item(.TextMatrix(lIndex, IGCol(eIGCol_InputID))).DefaultValue = 0
                    Case 3: ' Series of True/False
                        .TextMatrix(lIndex, IGCol(eIGCol_InputTypeID)) = 0
                        m.Function.Inputs.Item(.TextMatrix(lIndex, IGCol(eIGCol_InputID))).DefaultValue = 0
                    Case 4: ' Series of Numbers
                        .TextMatrix(lIndex, IGCol(eIGCol_InputTypeID)) = "Close"
                        m.Function.Inputs.Item(.TextMatrix(lIndex, IGCol(eIGCol_InputID))).DefaultValue = "Close"
                    Case 5: ' Market
                        .TextMatrix(lIndex, IGCol(eIGCol_InputTypeID)) = "Market1"
                        m.Function.Inputs.Item(.TextMatrix(lIndex, IGCol(eIGCol_InputID))).DefaultValue = "Market1"
                    Case 6: ' Single True/False
                        .TextMatrix(lIndex, IGCol(eIGCol_InputTypeID)) = 0
                        m.Function.Inputs.Item(.TextMatrix(lIndex, IGCol(eIGCol_InputID))).DefaultValue = 0
                End Select
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.SaveOrder", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      ID of the Function to load
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal lFunctionID As Long, Optional Chart As cChart = Nothing)
On Error GoTo ErrSection:

    If lFunctionID = 0 Then
        Add
    Else
        If Not LoadRec(lFunctionID) Then GoTo ErrExit
    End If
    
    Screen.MousePointer = vbDefault
    EnableToolbar False
    
    If Not Chart Is Nothing Then CenterFormOnChart Me, Chart            '6499
    ShowForm Me, False, frmMain

ErrExit:
    ''Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmFunctionMgr.ShowMe", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DLLNameExists
'' Description: Determine whether another function already exists with the
''              given DLL Name (Coded Name)
'' Inputs:      DLL Name to check
'' Returns:     True if already exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DLLNameExists(ByVal strDLLName As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
            "WHERE [CodedName]='" & strDLLName & "';", dbOpenDynaset)
    If Not (rs.EOF And rs.BOF) Then
        If rs!FunctionID <> m.Function.FunctionID Then DLLNameExists = True
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmFunctionMgr.DLLNameExists", eGDRaiseError_Raise

End Function

' Returns True if Cancelled
Public Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim strResponse As String
    
    If tbToolbar.Tools("ID_Save").Enabled Then
        If WindowState = vbMinimized Then WindowState = vbNormal
    
        strResponse = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", Caption)
        Select Case strResponse
            Case "C"
                AskToSave = True
            Case "Y"
                Save True, "ID_Save"
        End Select
    End If
        
ErrExit:
    Exit Function

ErrSection:
    AskToSave = True
    RaiseError Me.Name & ".AskToSave"

End Function

Private Sub AddInput()
On Error GoTo ErrSection:

    Dim xInput As New cInput
    Dim lParmID As Long
    Dim astrArgNames As New cGdArray
    Dim lIndex As Long
    
    astrArgNames.Create eGDARRAY_Strings
    With vsInputs
        For lIndex = .FixedRows To .Rows - 1
            astrArgNames.Add .TextMatrix(lIndex, IGCol(eIGCol_InputName))
        Next lIndex
    End With
    
    If frmFunctionInput.ShowMe(xInput, astrArgNames) Then
        ' Find next unique Paramter ID
        Do
            lParmID = lParmID + 1
        Loop While m.Function.Inputs.Found(CStr(lParmID)) = True
        
        With xInput
            m.Function.Inputs.Add 0, vsInputs.Rows, .ParmName, .ParmDesc, lParmID, _
                    .Value, 0, 0, 0, 0, 0, 0, .ParmTypeID, .DefaultValue, _
                    .Required, 0, 0, .ListID, .FillPre, .FillPost
        End With
        ShowParmLine TradeSense
        InitInputsGrid
        EnableToolbar True
    End If
    
ErrExit:
    Set astrArgNames = Nothing
    Set xInput = Nothing
    Exit Sub
    
ErrSection:
    Set astrArgNames = Nothing
    Set xInput = Nothing
    RaiseError "frmFunctionMgr.AddInput", eGDRaiseError_Raise
    
End Sub

Private Sub EditInput()
On Error GoTo ErrSection:

    Dim xInput As New cInput
    Dim astrArgNames As New cGdArray
    Dim lIndex As Long
    
    If ValOfText(vsInputs.TextMatrix(vsInputs.Row, IGCol(eIGCol_InputID))) = 0 Then Exit Sub
    
    astrArgNames.Create eGDARRAY_Strings
    With vsInputs
        For lIndex = .FixedRows To .Rows - 1
            astrArgNames.Add .TextMatrix(lIndex, IGCol(eIGCol_InputName))
        Next lIndex
    End With
        
    Set xInput = m.Function.Inputs.Item(vsInputs.TextMatrix(vsInputs.Row, IGCol(eIGCol_InputID)))
    If frmFunctionInput.ShowMe(xInput, astrArgNames) Then
        With m.Function.Inputs.Item(vsInputs.TextMatrix(vsInputs.Row, IGCol(eIGCol_InputID)))
            .ParmName = xInput.ParmName
            .ParmDesc = xInput.ParmDesc
            .ParmSeq = xInput.ParmSeq
            .ParmTypeID = xInput.ParmTypeID
            .Required = xInput.Required
            .DefaultValue = xInput.DefaultValue
        End With
        ShowParmLine TradeSense
        InitInputsGrid
        EnableToolbar True
    End If

ErrExit:
    Set xInput = Nothing
    Set astrArgNames = Nothing
    Exit Sub

ErrSection:
    Set xInput = Nothing
    Set astrArgNames = Nothing
    RaiseError "frmFunctionMgr.EditInput", eGDRaiseError_Raise
    
End Sub

Private Sub RemoveInput()
On Error GoTo ErrSection:

    m.Function.Inputs.Remove vsInputs.TextMatrix(vsInputs.Row, IGCol(eIGCol_InputID))
    InitInputsGrid
    EnableToolbar True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.RemoveInput", eGDRaiseError_Raise
    
End Sub

Private Sub vsInputs_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    With vsInputs
        If KeyCode = vbKeyDelete Then
            RemoveInput
        ElseIf KeyCode = vbKeyInsert Then
            AddInput
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgr.vsInputs.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

