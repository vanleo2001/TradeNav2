VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmDateJournal 
   Caption         =   "Form1"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   14805
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniTextBoxXP txtJournal 
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7875
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmDateJournal.frx":0000
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
      Tip             =   "frmDateJournal.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDateJournal.frx":0040
   End
   Begin HexUniControls.ctlUniFrameWL fraCustomChecklist 
      Height          =   5895
      Left            =   6660
      TabIndex        =   36
      Top             =   3540
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
      Caption         =   "frmDateJournal.frx":005C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDateJournal.frx":0088
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDateJournal.frx":00A8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
         Height          =   255
         Index           =   1
         Left            =   6070
         TabIndex        =   39
         Top             =   30
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
         Caption         =   "frmDateJournal.frx":00C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDateJournal.frx":00F6
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":0116
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraChecklistWeekly 
         Height          =   2475
         Left            =   0
         TabIndex        =   40
         Top             =   540
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
         Caption         =   "frmDateJournal.frx":0132
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDateJournal.frx":016C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":018C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklySetup 
            Height          =   315
            Index           =   5
            Left            =   180
            TabIndex        =   51
            Top             =   2040
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":01A8
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":01C8
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklySetup 
            Height          =   315
            Index           =   4
            Left            =   180
            TabIndex        =   49
            Top             =   1680
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":01E4
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0204
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklySetup 
            Height          =   315
            Index           =   3
            Left            =   180
            TabIndex        =   47
            Top             =   1320
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0220
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0240
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklySetup 
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   45
            Top             =   960
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":025C
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":027C
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklySetup 
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   43
            Top             =   600
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0298
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":02B8
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklySetup 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   41
            Top             =   240
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":02D4
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":02F4
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklyValue 
            Height          =   315
            Index           =   0
            Left            =   4140
            TabIndex        =   42
            Top             =   240
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0310
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0330
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklyValue 
            Height          =   315
            Index           =   1
            Left            =   4140
            TabIndex        =   44
            Top             =   600
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":034C
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":036C
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklyValue 
            Height          =   315
            Index           =   3
            Left            =   4140
            TabIndex        =   48
            Top             =   1320
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0388
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":03A8
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklyValue 
            Height          =   315
            Index           =   4
            Left            =   4140
            TabIndex        =   50
            Top             =   1680
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":03C4
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":03E4
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklyValue 
            Height          =   315
            Index           =   5
            Left            =   4140
            TabIndex        =   52
            Top             =   2040
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0400
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0420
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboWeeklyValue 
            Height          =   315
            Index           =   2
            Left            =   4140
            TabIndex        =   46
            Top             =   960
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":043C
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":045C
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblWeeklyDots 
            Height          =   195
            Index           =   5
            Left            =   3780
            Top             =   2160
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
            Caption         =   "frmDateJournal.frx":0478
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":04A8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":04C8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblWeeklyDots 
            Height          =   195
            Index           =   4
            Left            =   3780
            Top             =   1800
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
            Caption         =   "frmDateJournal.frx":04E4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":0514
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":0534
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblWeeklyDots 
            Height          =   195
            Index           =   3
            Left            =   3780
            Top             =   1440
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
            Caption         =   "frmDateJournal.frx":0550
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":0580
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":05A0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblWeeklyDots 
            Height          =   195
            Index           =   2
            Left            =   3780
            Top             =   1080
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
            Caption         =   "frmDateJournal.frx":05BC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":05EC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":060C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblWeeklyDots 
            Height          =   195
            Index           =   1
            Left            =   3780
            Top             =   720
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
            Caption         =   "frmDateJournal.frx":0628
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":0658
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":0678
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblWeeklyDots 
            Height          =   195
            Index           =   0
            Left            =   3780
            Top             =   360
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
            Caption         =   "frmDateJournal.frx":0694
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":06C4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":06E4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraChecklistDaily 
         Height          =   2475
         Left            =   0
         TabIndex        =   53
         Top             =   3120
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
         Caption         =   "frmDateJournal.frx":0700
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDateJournal.frx":0738
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":0758
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboBoxXP cboDailySetup 
            Height          =   315
            Index           =   5
            Left            =   180
            TabIndex        =   2
            Top             =   2040
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0774
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0794
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailySetup 
            Height          =   315
            Index           =   4
            Left            =   180
            TabIndex        =   5
            Top             =   1680
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":07B0
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":07D0
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailySetup 
            Height          =   315
            Index           =   3
            Left            =   180
            TabIndex        =   9
            Top             =   1320
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":07EC
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":080C
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailySetup 
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   11
            Top             =   960
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0828
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0848
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailySetup 
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   13
            Top             =   600
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0864
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0884
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailySetup 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   54
            Top             =   240
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":08A0
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":08C0
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailyValue 
            Height          =   315
            Index           =   1
            Left            =   4140
            TabIndex        =   15
            Top             =   600
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":08DC
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":08FC
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailyValue 
            Height          =   315
            Index           =   2
            Left            =   4140
            TabIndex        =   17
            Top             =   960
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0918
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0938
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailyValue 
            Height          =   315
            Index           =   5
            Left            =   4140
            TabIndex        =   19
            Top             =   2040
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0954
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0974
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailyValue 
            Height          =   315
            Index           =   0
            Left            =   4140
            TabIndex        =   55
            Top             =   240
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0990
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":09B0
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailyValue 
            Height          =   315
            Index           =   3
            Left            =   4140
            TabIndex        =   22
            Top             =   1320
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":09CC
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":09EC
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboDailyValue 
            Height          =   315
            Index           =   4
            Left            =   4140
            TabIndex        =   24
            Top             =   1680
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0A08
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0A28
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDailyDots 
            Height          =   195
            Index           =   5
            Left            =   3780
            Top             =   2160
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
            Caption         =   "frmDateJournal.frx":0A44
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":0A74
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":0A94
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDailyDots 
            Height          =   195
            Index           =   4
            Left            =   3780
            Top             =   1800
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
            Caption         =   "frmDateJournal.frx":0AB0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":0AE0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":0B00
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDailyDots 
            Height          =   195
            Index           =   3
            Left            =   3780
            Top             =   1440
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
            Caption         =   "frmDateJournal.frx":0B1C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":0B4C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":0B6C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDailyDots 
            Height          =   195
            Index           =   2
            Left            =   3780
            Top             =   1080
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
            Caption         =   "frmDateJournal.frx":0B88
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":0BB8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":0BD8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDailyDots 
            Height          =   195
            Index           =   1
            Left            =   3780
            Top             =   720
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
            Caption         =   "frmDateJournal.frx":0BF4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":0C24
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":0C44
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDailyDots 
            Height          =   195
            Index           =   0
            Left            =   3780
            Top             =   360
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
            Caption         =   "frmDateJournal.frx":0C60
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":0C90
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":0CB0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   315
         Index           =   1
         Left            =   4140
         TabIndex        =   38
         Top             =   0
         Width           =   2200
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmDateJournal.frx":0CCC
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
         Tip             =   "frmDateJournal.frx":0D06
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":0D26
      End
      Begin HexUniControls.ctlUniLabelXP lblNotes 
         Height          =   255
         Index           =   1
         Left            =   0
         Top             =   5700
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
         Caption         =   "frmDateJournal.frx":0D42
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDateJournal.frx":0D6E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":0D8E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSymbol 
         Height          =   195
         Index           =   1
         Left            =   0
         Top             =   120
         Width           =   4155
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmDateJournal.frx":0DAA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDateJournal.frx":0E7E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":0E9E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraMoneyCode 
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   3540
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
      Caption         =   "frmDateJournal.frx":0EBA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDateJournal.frx":0EE6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDateJournal.frx":0F06
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraDaily 
         Height          =   2475
         Left            =   0
         TabIndex        =   21
         Top             =   3120
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
         Caption         =   "frmDateJournal.frx":0F22
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDateJournal.frx":0F5A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":0F7A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboBoxXP cboEntryTrigger 
            Height          =   315
            Left            =   5400
            TabIndex        =   32
            Top             =   1680
            Width           =   2355
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
            Tip             =   "frmDateJournal.frx":0F96
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0FB6
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboOverall 
            Height          =   315
            Left            =   4140
            TabIndex        =   29
            Top             =   1320
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":0FD2
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":0FF2
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboEntryDirection 
            Height          =   315
            Left            =   4140
            TabIndex        =   31
            Top             =   1680
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
            Tip             =   "frmDateJournal.frx":100E
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":102E
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboOpenInterest 
            Height          =   315
            Left            =   4140
            TabIndex        =   23
            Top             =   240
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":104A
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":106A
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboExitTrigger 
            Height          =   315
            Left            =   4140
            TabIndex        =   34
            Top             =   2040
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":1086
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":10A6
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboPercentR 
            Height          =   315
            Left            =   4140
            TabIndex        =   27
            Top             =   960
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":10C2
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":10E2
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboAccumulation 
            Height          =   315
            Left            =   4140
            TabIndex        =   25
            Top             =   600
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":10FE
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":111E
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblOverall 
            Height          =   195
            Left            =   120
            Top             =   1440
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
            Caption         =   "frmDateJournal.frx":113A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":11FC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":121C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblOpenInterest 
            Height          =   195
            Left            =   120
            Top             =   360
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
            Caption         =   "frmDateJournal.frx":1238
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":12E4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":1304
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblExit 
            Height          =   195
            Left            =   120
            Top             =   2160
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
            Caption         =   "frmDateJournal.frx":1320
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":13E6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":1406
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblEntry 
            Height          =   195
            Left            =   120
            Top             =   1800
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
            Caption         =   "frmDateJournal.frx":1422
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":14D8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":14F8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPercentR 
            Height          =   195
            Left            =   120
            Top             =   1080
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
            Caption         =   "frmDateJournal.frx":1514
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":15CA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":15EA
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAccumulation 
            Height          =   195
            Left            =   120
            Top             =   720
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
            Caption         =   "frmDateJournal.frx":1606
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":16B6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":16D6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraWeekly 
         Height          =   2475
         Left            =   0
         TabIndex        =   8
         Top             =   540
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
         Caption         =   "frmDateJournal.frx":16F2
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDateJournal.frx":172C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":174C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboBoxXP cboAdvisor 
            Height          =   315
            Left            =   4140
            TabIndex        =   14
            Top             =   960
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":1768
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":1788
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboSeasonalDirection 
            Height          =   315
            Left            =   4140
            TabIndex        =   20
            Top             =   2040
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":17A4
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":17C4
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboSeasonalSetup 
            Height          =   315
            Left            =   4140
            TabIndex        =   18
            Top             =   1680
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":17E0
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":1800
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboAgricultural 
            Height          =   315
            Left            =   4140
            TabIndex        =   16
            Top             =   1320
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":181C
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":183C
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboLargeTraders 
            Height          =   315
            Left            =   4140
            TabIndex        =   12
            Top             =   600
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":1858
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":1878
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP cboCot 
            Height          =   315
            Left            =   4140
            TabIndex        =   10
            Top             =   240
            Width           =   3615
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
            Tip             =   "frmDateJournal.frx":1894
            Sorted          =   0   'False
            HScroll         =   0   'False
            Style           =   4
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
            MouseIcon       =   "frmDateJournal.frx":18B4
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAdvisor 
            Height          =   195
            Left            =   120
            Top             =   1080
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
            Caption         =   "frmDateJournal.frx":18D0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":198C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":19AC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSeasonalDirection 
            Height          =   195
            Left            =   120
            Top             =   2160
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
            Caption         =   "frmDateJournal.frx":19C8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":1A7A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":1A9A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSeasonalSetup 
            Height          =   195
            Left            =   120
            Top             =   1800
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
            Caption         =   "frmDateJournal.frx":1AB6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":1B5E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":1B7E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAgricultural 
            Height          =   195
            Left            =   120
            Top             =   1440
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
            Caption         =   "frmDateJournal.frx":1B9A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":1C4A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":1C6A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblLargeTraders 
            Height          =   195
            Left            =   120
            Top             =   720
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
            Caption         =   "frmDateJournal.frx":1C86
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":1D3A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":1D5A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblCot 
            Height          =   195
            Left            =   120
            Top             =   360
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
            Caption         =   "frmDateJournal.frx":1D76
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDateJournal.frx":1E18
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDateJournal.frx":1E38
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
         Height          =   255
         Index           =   0
         Left            =   6070
         TabIndex        =   7
         Top             =   30
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
         Caption         =   "frmDateJournal.frx":1E54
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDateJournal.frx":1E86
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":1EA6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   315
         Index           =   0
         Left            =   4140
         TabIndex        =   6
         Top             =   0
         Width           =   2200
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmDateJournal.frx":1EC2
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
         Tip             =   "frmDateJournal.frx":1EFC
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":1F1C
      End
      Begin HexUniControls.ctlUniLabelXP lblSymbol 
         Height          =   195
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   4155
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmDateJournal.frx":1F38
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDateJournal.frx":200C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":202C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblNotes 
         Height          =   255
         Index           =   0
         Left            =   0
         Top             =   5700
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
         Caption         =   "frmDateJournal.frx":2048
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDateJournal.frx":2074
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":2094
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraImage 
      Height          =   315
      Left            =   180
      TabIndex        =   26
      Top             =   2400
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
      Caption         =   "frmDateJournal.frx":20B0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDateJournal.frx":20E6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDateJournal.frx":2106
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkAttachImage 
         Height          =   220
         Left            =   0
         TabIndex        =   28
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "frmDateJournal.frx":2122
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDateJournal.frx":2168
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":2188
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboCharts 
         Height          =   315
         Left            =   1800
         TabIndex        =   30
         Top             =   0
         Width           =   6075
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
         Tip             =   "frmDateJournal.frx":21A4
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":21C4
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
   End
   Begin gdOCX.gdSelectDate gdJournalTime 
      Height          =   315
      Left            =   5100
      TabIndex        =   3
      Top             =   180
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      ShowTime        =   2
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   960
      TabIndex        =   33
      Top             =   2880
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
      Caption         =   "frmDateJournal.frx":21E0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDateJournal.frx":220C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDateJournal.frx":222C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   1260
         TabIndex        =   35
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
         Caption         =   "frmDateJournal.frx":2248
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDateJournal.frx":2276
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":2296
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   37
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
         Caption         =   "frmDateJournal.frx":22B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDateJournal.frx":22D8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDateJournal.frx":22F8
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboCategories 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   180
      Width           =   2955
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
      Tip             =   "frmDateJournal.frx":2314
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmDateJournal.frx":2334
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblJournalTime 
      Height          =   195
      Left            =   4080
      Top             =   240
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
      Caption         =   "frmDateJournal.frx":2350
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmDateJournal.frx":238C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDateJournal.frx":23AC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblCategory 
      Height          =   255
      Left            =   120
      Top             =   210
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
      Caption         =   "frmDateJournal.frx":23C8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmDateJournal.frx":23FC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDateJournal.frx":241C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmDateJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDateJournal.frm
'' Description: Form that allows the user to enter in a journal entry
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/20/2011   DAJ         Changed icon
'' 09/22/2011   DAJ         Added the ability to attach a chart image
'' 09/22/2011   DAJ         Persist size/location, change format in charts combo, verify input
'' 09/23/2011   DAJ         Fix for exporting non-active maximized chart
'' 10/11/2011   DAJ         Add attach chart check box
'' 01/25/2012   DAJ         Money Code journal
'' 01/30/2012   DAJ         Option Nav Journal Image
'' 03/13/2012   DAJ         New version of the Money Code journal
'' 03/19/2012   DAJ         Added Symbol/Symbol ID to date journal object
'' 07/30/2013   DAJ         Moved out chart export code
'' 08/08/2013   DAJ         Custom checklist journal
'' 08/15/2013   DAJ         Allow for saving 20 weekly/daily setups for custom checklist
'' 08/30/2013   DAJ         Use same MRU list for all combo's in a custom checklist group
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/08/2014   DAJ         Use NavCore Image List; Use newer place/save form
'' 10/24/2014   DAJ         Core Application functions for DLL's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    dCurrentTime As Double              ' Current date/time
    nMode As eGDJournalCategoryTypes    ' Current mode of the form
    strJournalIni As String             ' Journal INI file
    strDefaultsIni As String            ' Defaults INI file
    bVisible As Boolean                 ' Form is visible
    
    strComboText As String              ' Text from the combo box
    
    Mrus As cGdTree                     ' Collection of MRU's
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Journal Entry, Journal Date
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(journalEntry As cDateJournal, ByVal dJournalDate As Double) As Boolean
On Error GoTo ErrSection:

    m.bVisible = False
    Caption = "Journal Entry"

    LoadCategories journalEntry.JournalCategoryID
    JournalEntryToControls journalEntry, dJournalDate
    LoadChartsCombo journalEntry.JournalImage(eGDJournalImageType_Chart)
    
    If g.bAppIsIde Then
        mGenesis.ShowForm Me, eForm_Modal
    Else
        g.TnCore.ShowForm Me, eForm_Modal
    End If
    
    m.bVisible = False
    If m.bOK Then
        JournalEntryFromControls journalEntry
        SaveMoneyCodeMru
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    g.TnCore.RaiseError "frmDateJournal.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboCategories_Click
'' Description: User has changed the category
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCategories_Click()
On Error GoTo ErrSection:

    Dim nNewCategory As eGDJournalCategoryTypes ' New category

    If (cboCategories.ListIndex > -1) Then
        nNewCategory = g.JournalCategories.TypeForId(cboCategories.ItemData(cboCategories.ListIndex))

        If nNewCategory = eGDJournalCategoryType_MoneyCode Then
            If m.nMode <> eGDJournalCategoryType_MoneyCode Then
                m.nMode = eGDJournalCategoryType_MoneyCode
                
                fraMoneyCode.Visible = True
                fraCustomChecklist.Visible = False
                
                Form_Resize
            End If
        ElseIf nNewCategory = eGDJournalCategoryType_CustomChecklist Then
            If m.nMode <> eGDJournalCategoryType_CustomChecklist Then
                m.nMode = eGDJournalCategoryType_CustomChecklist
            
                fraMoneyCode.Visible = False
                fraCustomChecklist.Visible = True
                
                Form_Resize
            End If
        Else
            If m.nMode <> eGDJournalCategoryType_Note Then
                m.nMode = eGDJournalCategoryType_Note
            
                fraMoneyCode.Visible = False
                fraCustomChecklist.Visible = False
                
                Form_Resize
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboCategories_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboCategories_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCategories_GotFocus()
On Error GoTo ErrSection:

    SelectAll cboCategories

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboCategories_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboDailySetup_Dropdown
'' Description: When the user drops down the combo, fill the list from the MRU
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDailySetup_DropDown(Index As Integer)
On Error GoTo ErrSection:

    ComboDropDown "DailySetup", cboDailySetup(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboDailySetup_DropDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboDailySetup_GotFocus
'' Description: When the combo gets focus, save off the current text
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDailySetup_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    m.strComboText = cboDailySetup(Index).Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboDailySetup_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboDailySetup_LostFocus
'' Description: When the combo loses focus, save the changes to the MRU list
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDailySetup_LostFocus(Index As Integer)
On Error GoTo ErrSection:

    ComboLostFocus "DailySetup", cboDailySetup(Index).Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboDailySetup_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboDailyValue_Dropdown
'' Description: When the user drops down the combo, fill the list from the MRU
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDailyValue_DropDown(Index As Integer)
On Error GoTo ErrSection:

    ComboDropDown "DailyValue", cboDailyValue(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboDailyValue_DropDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboDailyValue_GotFocus
'' Description: When the combo gets focus, save off the current text
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDailyValue_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    m.strComboText = cboDailyValue(Index).Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboDailyValue_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboDailyValue_LostFocus
'' Description: When the combo loses focus, save the changes to the MRU list
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDailyValue_LostFocus(Index As Integer)
On Error GoTo ErrSection:

    ComboLostFocus "DailyValue", cboDailyValue(Index).Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboDailyValue_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboWeeklySetup_Dropdown
'' Description: When the user drops down the combo, fill the list from the MRU
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboWeeklySetup_DropDown(Index As Integer)
On Error GoTo ErrSection:

    ComboDropDown "WeeklySetup", cboWeeklySetup(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboWeeklySetup_DropDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboWeeklySetup_GotFocus
'' Description: When the combo gets focus, save off the current text
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboWeeklySetup_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    m.strComboText = cboWeeklySetup(Index).Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboWeeklySetup_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboWeeklySetup_LostFocus
'' Description: When the combo loses focus, save the changes to the MRU list
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboWeeklySetup_LostFocus(Index As Integer)
On Error GoTo ErrSection:

    ComboLostFocus "WeeklySetup", cboWeeklySetup(Index).Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboWeeklySetup_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboWeeklyValue_Dropdown
'' Description: When the user drops down the combo, fill the list from the MRU
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboWeeklyValue_DropDown(Index As Integer)
On Error GoTo ErrSection:

    ComboDropDown "WeeklyValue", cboWeeklyValue(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboWeeklyValue_DropDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboWeeklyValue_GotFocus
'' Description: When the combo gets focus, save off the current text
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboWeeklyValue_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    m.strComboText = cboWeeklyValue(Index).Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboWeeklyValue_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboWeeklyValue_LostFocus
'' Description: When the combo loses focus, save the changes to the MRU list
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboWeeklyValue_LostFocus(Index As Integer)
On Error GoTo ErrSection:

    ComboLostFocus "WeeklyValue", cboWeeklyValue(Index).Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cboWeeklyValue_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: User has chosen to cancel the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: Allow the user to lookup a symbol
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookup_Click(Index As Integer)
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cmdLookup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: User has chosen to OK the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If cboCategories.ListIndex = -1& Then
        MoveFocus cboCategories
        InfBox "Please choose a category", "!", , "Error"
    ElseIf (m.nMode <> eGDJournalCategoryType_Note) And (InStr(txtJournal.Text, "=") <> 0) Then
        MoveFocus txtJournal
        InfBox "Notes cannot contain an equals sign (=)", "!", , "Error"
    ElseIf (m.nMode <> eGDJournalCategoryType_Note) And (InStr(txtJournal.Text, "|") <> 0) Then
        MoveFocus txtJournal
        InfBox "Notes cannot contain a pipe character (|)", "!", , "Error"
    Else
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Handle the form getting activated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean      ' Have we already done the first time stuff?
    
    If bAlreadyDone = False Then
        bAlreadyDone = True
        
        m.bVisible = True
        MoveFocus cboCategories
        
        ' For some reason, all of the combo boxes except entry direction want to
        ' have all of their text selected when the form comes up for a pre-existing
        ' Money Code journal.  Doing this here seems to stop that...
        cboCot.SelLength = 0
        cboLargeTraders.SelLength = 0
        cboAdvisor.SelLength = 0
        cboAgricultural.SelLength = 0
        cboSeasonalSetup.SelLength = 0
        cboSeasonalDirection.SelLength = 0
        cboOpenInterest.SelLength = 0
        cboAccumulation.SelLength = 0
        cboPercentR.SelLength = 0
        cboOverall.SelLength = 0
        cboEntryDirection.SelLength = 0
        cboEntryTrigger.SelLength = 0
        cboExitTrigger.SelLength = 0
    End If

ErrExit:
    Exit Sub

ErrSection:
    g.TnCore.RaiseError "frmDateJournal.Form_Activate"
    
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

    Dim strActiveChartSymbol As String  ' Active chart symbol

    m.nMode = kNullData
    Icon = g.CoreBridge.Picture16(g.TnCore.ToolbarIcon("kScroll"))
    
    g.Styler.StyleForm Me
    
    PlaceTheForm Me, g.strIniFile
    
    fraMoneyCode.Visible = False
    fraCustomChecklist.Visible = False
    
    m.strJournalIni = AddSlash(g.strAppPath) & "Journal.INI"
    m.strDefaultsIni = AddSlash(g.strAppPath) & "Provided\Journal.INI"

    Set m.Mrus = New cGdTree
    LoadMrus
    
    strActiveChartSymbol = g.AppBridge.ActiveChartSymbol
    If Len(strActiveChartSymbol) > 0 Then
        txtSymbol(0).Text = strActiveChartSymbol
        txtSymbol(1).Text = txtSymbol(0).Text
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the "X", hide the form
'' Inputs:      Cancel Unload?, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lTop As Long                    ' Top of the moveable controls
    Dim lMinWidth As Long               ' Minimum form width
    Dim lMinHeight As Long              ' Minimum form height
    Dim lLeft As Long                   ' Left of the notes box
    
    lMinWidth = 7875 + (120 * 2)
    If m.nMode <> eGDJournalCategoryType_Note Then
        lMinHeight = fraWeekly.Height * 3.5
    Else
        lMinHeight = fraButtons.Height * 5
    End If

    If LimitFormSize(Me, lMinWidth, lMinHeight) = False Then
        With fraButtons
            .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height - 120
        End With
        
        With fraImage
            .Move 120, fraButtons.Top - .Height - 120, ScaleWidth - 240
        End With
        
        With cboCharts
            .Move .Left, .Top, fraImage.Width - .Left - 120
        End With
        
        lTop = cboCategories.Top + cboCategories.Height + 120
        lLeft = 120
        If m.nMode = eGDJournalCategoryType_MoneyCode Then
            With fraMoneyCode
                .Move 120, lTop, ScaleWidth - 240
            End With
            
            With fraWeekly
                .Move .Left, .Top, fraMoneyCode.Width - .Left
            End With
            
            With cboCot
                .Move .Left, .Top, fraWeekly.Width - .Left - 60
            End With
            
            With cboLargeTraders
                .Move .Left, .Top, fraWeekly.Width - .Left - 60
            End With
            
            With cboAdvisor
                .Move .Left, .Top, fraWeekly.Width - .Left - 60
            End With
            
            With cboAgricultural
                .Move .Left, .Top, fraWeekly.Width - .Left - 60
            End With
            
            With cboSeasonalSetup
                .Move .Left, .Top, fraWeekly.Width - .Left - 60
            End With
            
            With cboSeasonalDirection
                .Move .Left, .Top, fraWeekly.Width - .Left - 60
            End With
            
            With fraDaily
                .Move .Left, .Top, fraMoneyCode.Width - .Left
            End With
            
            With cboOpenInterest
                .Move .Left, .Top, fraDaily.Width - .Left - 60
            End With
            
            With cboAccumulation
                .Move .Left, .Top, fraDaily.Width - .Left - 60
            End With
            
            With cboPercentR
                .Move .Left, .Top, fraDaily.Width - .Left - 60
            End With
            
            With cboOverall
                .Move .Left, .Top, fraDaily.Width - .Left - 60
            End With
            
            With cboEntryTrigger
                .Move .Left, .Top, fraDaily.Width - .Left - 60
            End With
            
            With cboExitTrigger
                .Move .Left, .Top, fraDaily.Width - .Left - 60
            End With
            
            lTop = fraMoneyCode.Top + fraMoneyCode.Height
            lLeft = 120 '480
        ElseIf m.nMode = eGDJournalCategoryType_CustomChecklist Then
            With fraCustomChecklist
                .Move 120, lTop, ScaleWidth - 240
            End With
            
            lTop = fraCustomChecklist.Top + fraCustomChecklist.Height
            lLeft = 120 '480
        End If
        
        With txtJournal
            .Move lLeft, lTop, ScaleWidth - lLeft - 120, ScaleHeight - lTop - fraImage.Height - fraButtons.Height - 360
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, save the placement and category id
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SaveTheFormPlacement Me, g.strIniFile

    Set m.Mrus = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtJournal_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtJournal_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtJournal

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.txtJournal_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_Click
'' Description: Allow the user to lookup a symbol
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_Click(Index As Integer, Button As Integer)
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.txtSymbol_Click", 0
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_GotFocus(Index As Integer)
On Error GoTo ErrSection:

    SelectAll txtSymbol(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.txtSymbol_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_KeyPress
'' Description: Allow the user to lookup a symbol
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_KeyPress(Index As Integer, KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LookupSymbol KeyAscii
    KeyAscii = 0

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.txtSymbol_KeyPress", 0
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCategories
'' Description: Load the categories into the combo box
'' Inputs:      Include Category ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCategories(Optional ByVal lIncludeCategoryID As Long = 0&)
On Error GoTo ErrSection:

    Dim JournalCategories As cJournalCategories
    Dim journalCategory As cJournalCategory
    Dim lIndex As Long                  ' Index into a for loop
    
    Set JournalCategories = g.JournalCategories
    
    cboCategories.Clear
    For lIndex = 1 To JournalCategories.Count
        Set journalCategory = JournalCategories(lIndex)
        
        If ((g.TnCore.HasModule(journalCategory.RequiredModule) = True) And (journalCategory.Active = True)) Or (journalCategory.ID = lIncludeCategoryID) Then
            cboCategories.AddItem journalCategory.Text
            cboCategories.ItemData(cboCategories.NewIndex) = journalCategory.ID
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.LoadCategories"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCategory
'' Description: Set the category combo from the given ID
'' Inputs:      Category ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCategory(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To cboCategories.ListCount - 1
        If cboCategories.ItemData(lIndex) = lCategoryID Then
            cboCategories.ListIndex = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.SetCategory"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    JournalEntryToControls
'' Description: Set the controls from the given journal entry
'' Inputs:      Journal Entry, Journal Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub JournalEntryToControls(ByVal journalEntry As cDateJournal, ByVal dJournalDate As Double)
On Error GoTo ErrSection:

    Dim dDateTime As Double             ' Date/Time of the journal entry

    If journalEntry.DateJournalID = 0 Then
        m.dCurrentTime = g.TnCore.CurrentTime
        
        txtJournal.Text = ""
        If Int(dJournalDate) = Int(m.dCurrentTime) Then
            gdJournalTime.Value = m.dCurrentTime
        Else
            gdJournalTime.Value = dJournalDate
        End If
        
        SetCategory GetIniFileProperty("Default", 1&, "Categories", m.strDefaultsIni)
    Else
        dDateTime = journalEntry.JournalDate + journalEntry.JournalTime
        
        gdJournalTime.Value = dDateTime
        SetCategory journalEntry.JournalCategoryID
        
        Select Case m.nMode
            Case eGDJournalCategoryType_MoneyCode
                MoneyCodeToControls journalEntry.Text
            Case eGDJournalCategoryType_CustomChecklist
                CustomChecklistToControls journalEntry.Text
            Case Else
                txtJournal.Text = journalEntry.Text
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.JournalEntryToControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    JournalEntryFromControls
'' Description: Set the journal entry from the controls
'' Inputs:      Journal Entry
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub JournalEntryFromControls(journalEntry As cDateJournal)
On Error GoTo ErrSection:

    Dim dDateTime As Double             ' Date and time of the journal entry
    Dim JournalImage As cJournalImage   ' Journal image information
    
    dDateTime = gdJournalTime.Value

    journalEntry.JournalCategoryID = cboCategories.ItemData(cboCategories.ListIndex)
    journalEntry.JournalDate = CDbl(Int(dDateTime))
    journalEntry.JournalTime = dDateTime - journalEntry.JournalDate
    
    If CheckBoxValue(chkAttachImage) = False Then
        Set JournalImage = journalEntry.JournalImage(eGDJournalImageType_Chart)
        If Not JournalImage Is Nothing Then
            If Len(JournalImage.FileName) > 0 Then
                KillFile JournalImage.FileName
            End If
            
            journalEntry.JournalImages.Remove Str(eGDJournalImageType_Chart)
        End If
    Else
        Set JournalImage = journalEntry.JournalImage(eGDJournalImageType_Chart)
        If JournalImage Is Nothing Then
            Set JournalImage = New cJournalImage
        ElseIf cboCharts.Text <> JournalImage.Caption Then
            If Len(JournalImage.FileName) > 0 Then
                KillFile JournalImage.FileName
            End If
        End If
                
        JournalImage.FileName = ExportChart
        If Len(JournalImage.FileName) > 0 Then
            JournalImage.Caption = cboCharts.Text
        End If
        JournalImage.DateJournalID = journalEntry.DateJournalID
        journalEntry.JournalImage(eGDJournalImageType_Chart) = JournalImage
    End If
    
    Select Case m.nMode
        Case eGDJournalCategoryType_MoneyCode
            journalEntry.Text = MoneyCodeFromControls
            journalEntry.SymbolOrSymbolID = Trim(txtSymbol(eGDJournalCategoryType_MoneyCode).Text)
        Case eGDJournalCategoryType_CustomChecklist
            journalEntry.Text = CustomChecklistFromControls
            journalEntry.SymbolOrSymbolID = Trim(txtSymbol(eGDJournalCategoryType_CustomChecklist).Text)
        Case Else
            journalEntry.Text = txtJournal.Text
            journalEntry.SymbolOrSymbolID = ""
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.JournalEntryFromControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadChartsCombo
'' Description: Load the charts combo box
'' Inputs:      Previously selected image caption
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadChartsCombo(ByVal ChartImage As cJournalImage)
On Error GoTo ErrSection:

    Dim strImageCaption As String       ' Image caption

    strImageCaption = ""
    If Not ChartImage Is Nothing Then
        strImageCaption = ChartImage.Caption
    End If
    
    g.AppBridge.LoadChartsCombo cboCharts, strImageCaption
    
    CheckBoxValue(chkAttachImage) = (Len(strImageCaption) > 0)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.LoadChartsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportChart
'' Description: Export the selected chart
'' Inputs:      None
'' Returns:     Filename
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ExportChart() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lHwnd As Long                   ' Hwnd for the window
    
    strReturn = ""
    
    If cboCharts.ListIndex >= 0 Then
        lHwnd = cboCharts.ItemData(cboCharts.ListIndex)
        strReturn = g.AppBridge.ChartImageForHwnd(lHwnd)
    End If
    
    ExportChart = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.ExportChart"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoneyCodeFromControls
'' Description: Build string to save for the Money Code controls
'' Inputs:      None
'' Returns:     Money Code String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MoneyCodeFromControls() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 16
    
    astrReturn(0) = "Version=1"
    astrReturn(1) = "Symbol=" & Trim(txtSymbol(eGDJournalCategoryType_MoneyCode).Text)
    astrReturn(2) = "Cot=" & Trim(cboCot.Text)
    astrReturn(3) = "LargeTraders=" & Trim(cboLargeTraders.Text)
    astrReturn(4) = "Advisor=" & Trim(cboAdvisor.Text)
    astrReturn(5) = "Agricultural=" & Trim(cboAgricultural.Text)
    astrReturn(6) = "SeasonalSetup=" & Trim(cboSeasonalSetup.Text)
    astrReturn(7) = "SeasonalDirection=" & Trim(cboSeasonalDirection.Text)
    astrReturn(8) = "OpenInterest=" & Trim(cboOpenInterest.Text)
    astrReturn(9) = "Accumulation=" & Trim(cboAccumulation.Text)
    astrReturn(10) = "PercentR=" & Trim(cboPercentR.Text)
    astrReturn(11) = "Overall=" & Trim(cboOverall.Text)
    astrReturn(12) = "EntryDirection=" & Trim(cboEntryDirection.Text)
    astrReturn(13) = "EntryTrigger=" & Trim(cboEntryTrigger.Text)
    astrReturn(14) = "ExitTrigger=" & Trim(cboExitTrigger.Text)
    astrReturn(15) = "Notes=" & Trim(txtJournal.Text)
    
    MoneyCodeFromControls = astrReturn.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.MoneyCodeFromControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoneyCodeToControls
'' Description: Set the Money Code controls from the string
'' Inputs:      Money Code string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MoneyCodeToControls(ByVal strMoneyCode As String)
On Error GoTo ErrSection:

    Dim moneyCodeFields As cGdTree      ' Dictionary of fields from the given string

    If UCase(Left(strMoneyCode, 8)) <> "VERSION=" Then
        MoneyCodeToControls0 strMoneyCode
    Else
        Set moneyCodeFields = New cGdTree
        moneyCodeFields.FromKeyValueString strMoneyCode, "|", "="
        
        If moneyCodeFields.Exists("Version") Then
            Select Case moneyCodeFields("Version")
                Case "1"
                    MoneyCodeToControls1 moneyCodeFields
            
            End Select
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.MoneyCodeToControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoneyCodeToControls0
'' Description: Set the Money Code controls from the version zero string
'' Inputs:      Version Zero Money Code string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MoneyCodeToControls0(ByVal strMoneyCode As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Array of fields
    Dim astrField As cGdArray           ' Array of field information
    
    If Len(strMoneyCode) > 0 Then
        Set astrField = New cGdArray
        Set astrFields = New cGdArray
        astrFields.SplitFields strMoneyCode, "|"
        
        astrField.SplitFields astrFields(1), ";"
        cboCot.Text = astrField(1)
        
        astrField.SplitFields astrFields(2), ";"
        cboLargeTraders.Text = astrField(1)
        
        astrField.SplitFields astrFields(3), ";"
        cboAgricultural.Text = astrField(1)
        
        astrField.SplitFields astrFields(4), ";"
        cboSeasonalSetup.Text = astrField(1)
        
        astrField.SplitFields astrFields(5), ";"
        cboOpenInterest.Text = astrField(1)
        
        astrField.SplitFields astrFields(6), ";"
        cboAccumulation.Text = astrField(1)
        
        astrField.SplitFields astrFields(7), ";"
        cboPercentR.Text = astrField(1)
        
        astrField.SplitFields astrFields(9), ";"
        txtJournal.Text = astrField(1)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.MoneyCodeToControls0"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoneyCodeToControls1
'' Description: Set the Money Code controls from the version one string
'' Inputs:      Version One Money Code string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MoneyCodeToControls1(ByVal moneyCodeFields As cGdTree)
On Error GoTo ErrSection:

    If Not moneyCodeFields Is Nothing Then
        If moneyCodeFields.Exists("Symbol") Then
            txtSymbol(eGDJournalCategoryType_MoneyCode).Text = moneyCodeFields("Symbol")
        End If
        If moneyCodeFields.Exists("Cot") Then
            SetMruCombo moneyCodeFields("Cot"), cboCot
        End If
        If moneyCodeFields.Exists("LargeTraders") Then
            SetMruCombo moneyCodeFields("LargeTraders"), cboLargeTraders
        End If
        If moneyCodeFields.Exists("Advisor") Then
            SetMruCombo moneyCodeFields("Advisor"), cboAdvisor
        End If
        If moneyCodeFields.Exists("Agricultural") Then
            SetMruCombo moneyCodeFields("Agricultural"), cboAgricultural
        End If
        If moneyCodeFields.Exists("SeasonalSetup") Then
            SetMruCombo moneyCodeFields("SeasonalSetup"), cboSeasonalSetup
        End If
        If moneyCodeFields.Exists("SeasonalDirection") Then
            SetMruCombo moneyCodeFields("SeasonalDirection"), cboSeasonalDirection
        End If
        If moneyCodeFields.Exists("OpenInterest") Then
            SetMruCombo moneyCodeFields("OpenInterest"), cboOpenInterest
        End If
        If moneyCodeFields.Exists("Accumulation") Then
            SetMruCombo moneyCodeFields("Accumulation"), cboAccumulation
        End If
        If moneyCodeFields.Exists("PercentR") Then
            SetMruCombo moneyCodeFields("PercentR"), cboPercentR
        End If
        If moneyCodeFields.Exists("Overall") Then
            SetMruCombo moneyCodeFields("Overall"), cboOverall
        End If
        If moneyCodeFields.Exists("EntryDirection") Then
            SetMruCombo moneyCodeFields("EntryDirection"), cboEntryDirection
        End If
        If moneyCodeFields.Exists("EntryTrigger") Then
            SetMruCombo moneyCodeFields("EntryTrigger"), cboEntryTrigger
        End If
        If moneyCodeFields.Exists("ExitTrigger") Then
            SetMruCombo moneyCodeFields("ExitTrigger"), cboExitTrigger
        End If
        If moneyCodeFields.Exists("Notes") Then
            txtJournal.Text = moneyCodeFields("Notes")
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.MoneyCodeToControls1"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CustomChecklistFromControls
'' Description: Build string to save for the Custom Checklist controls
'' Inputs:      None
'' Returns:     Custom Checklist String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CustomChecklistFromControls() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim strText As String               ' Text to store
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 15
    
    astrReturn(0) = "Version=1"
    astrReturn(1) = "Symbol=" & Trim(txtSymbol(eGDJournalCategoryType_CustomChecklist).Text)
    
    For lIndex = 0 To 5
        strText = cboWeeklySetup(lIndex).Text
        If Len(strText) > 0 Then
            strText = strText & ";" & cboWeeklyValue(lIndex).Text
        End If
        astrReturn(lIndex + 2) = "Weekly" & Str(lIndex) & "=" & strText

        strText = cboDailySetup(lIndex).Text
        If Len(strText) > 0 Then
            strText = strText & ";" & cboDailyValue(lIndex).Text
        End If
        astrReturn(lIndex + 8) = "Daily" & Str(lIndex) & "=" & strText
    Next lIndex
    
    astrReturn(14) = "Notes=" & Trim(txtJournal.Text)
    
    CustomChecklistFromControls = astrReturn.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.CustomChecklistFromControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CustomChecklistToControls
'' Description: Set the Custom Checklist controls from the string
'' Inputs:      Custom Checklist string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CustomChecklistToControls(ByVal strCustomChecklist As String)
On Error GoTo ErrSection:

    Dim customChecklistFields As cGdTree ' Dictionary of fields from the given string

    Set customChecklistFields = New cGdTree
    customChecklistFields.FromKeyValueString strCustomChecklist, "|", "="
    
    If customChecklistFields.Exists("Version") Then
        Select Case customChecklistFields("Version")
            Case "1"
                CustomChecklistToControls1 customChecklistFields
        
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.CustomChecklistToControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CustomChecklistToControls1
'' Description: Set the Custom Checklist controls from the version one string
'' Inputs:      Version One Custom Checklist string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CustomChecklistToControls1(ByVal customChecklistFields As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strKey As String                ' Key into the collection
    Dim strValue As String              ' Value for the key out of the collection

    If Not customChecklistFields Is Nothing Then
        If customChecklistFields.Exists("Symbol") Then
            txtSymbol(eGDJournalCategoryType_CustomChecklist).Text = customChecklistFields("Symbol")
        End If
        
        For lIndex = 0 To 5
            strKey = "Weekly" & Str(lIndex)
            If customChecklistFields.Exists(strKey) Then
                strValue = customChecklistFields(strKey)
                
                cboWeeklySetup(lIndex).Text = Parse(strValue, ";", 1)
                cboWeeklyValue(lIndex).Text = Parse(strValue, ";", 2)
            End If
        
            strKey = "Daily" & Str(lIndex)
            If customChecklistFields.Exists(strKey) Then
                strValue = customChecklistFields(strKey)
                
                cboDailySetup(lIndex).Text = Parse(strValue, ";", 1)
                cboDailyValue(lIndex).Text = Parse(strValue, ";", 2)
            End If
        Next lIndex
        
        If customChecklistFields.Exists("Notes") Then
            txtJournal.Text = customChecklistFields("Notes")
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.CustomChecklistToControls1"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMruCombo
'' Description: Load up the given MRU combo from the most recently used list
'' Inputs:      MRU Object, Combo Box
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadMruCombo(MRU As cMostRecentlyUsed, cboMru As ComboBox)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    cboMru.Clear
    For lIndex = 0 To MRU.Count - 1
        cboMru.AddItem MRU.RecentlyUsedList(lIndex)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.LoadMruCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetMruCombo
'' Description: Set the given MRU combo to the given value
'' Inputs:      Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetMruCombo(ByVal strValue As String, cboMru As ComboBox, Optional ByVal bForceSet As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lValueIndex As Long             ' Index for the given value
    
    lValueIndex = -1&
    If Len(strValue) > 0 Then
        For lIndex = 0 To cboMru.ListCount - 1
            If cboMru.List(lIndex) = strValue Then
                lValueIndex = lIndex
                Exit For
            End If
        Next lIndex
    End If

    If lValueIndex > -1& Then
        cboMru.ListIndex = lValueIndex
    ElseIf (cboMru.ListCount > 0) And (bForceSet = True) Then
        cboMru.ListIndex = 0
    Else
        cboMru.Text = ""
    End If
    
    cboMru.ToolTipText = cboMru.Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.SetMruCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveMoneyCodeMru
'' Description: Save the Money Code Mru values
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveMoneyCodeMru()
On Error GoTo ErrSection:

    Dim MRU As cMostRecentlyUsed        ' Most recently used object

    If Len(cboCot.Text) > 0 Then
        Set MRU = m.Mrus("COT")
        If MRU.LastUsed <> cboCot.Text Then
            MRU.LastUsed = cboCot.Text
            MRU.Save
        End If
    End If

    If Len(cboLargeTraders.Text) > 0 Then
        Set MRU = m.Mrus("LargeTraders")
        If MRU.LastUsed <> cboLargeTraders.Text Then
            MRU.LastUsed = cboLargeTraders.Text
            MRU.Save
        End If
    End If

    If Len(cboAdvisor.Text) > 0 Then
        Set MRU = m.Mrus("Advisor")
        If MRU.LastUsed <> cboAdvisor.Text Then
            MRU.LastUsed = cboAdvisor.Text
            MRU.Save
        End If
    End If

    If Len(cboAgricultural.Text) > 0 Then
        Set MRU = m.Mrus("Agricultural")
        If MRU.LastUsed <> cboAgricultural.Text Then
            MRU.LastUsed = cboAgricultural.Text
            MRU.Save
        End If
    End If

    If Len(cboSeasonalSetup.Text) > 0 Then
        Set MRU = m.Mrus("SeasonalSetup")
        If MRU.LastUsed <> cboSeasonalSetup.Text Then
            MRU.LastUsed = cboSeasonalSetup.Text
            MRU.Save
        End If
    End If

    If Len(cboSeasonalDirection.Text) > 0 Then
        Set MRU = m.Mrus("SeasonalDirection")
        If MRU.LastUsed <> cboSeasonalDirection.Text Then
            MRU.LastUsed = cboSeasonalDirection.Text
            MRU.Save
        End If
    End If

    If Len(cboOpenInterest.Text) > 0 Then
        Set MRU = m.Mrus("OpenInterest")
        If MRU.LastUsed <> cboOpenInterest.Text Then
            MRU.LastUsed = cboOpenInterest.Text
            MRU.Save
        End If
    End If

    If Len(cboAccumulation.Text) > 0 Then
        Set MRU = m.Mrus("Accumulation")
        If MRU.LastUsed <> cboAccumulation.Text Then
            MRU.LastUsed = cboAccumulation.Text
            MRU.Save
        End If
    End If

    If Len(cboPercentR.Text) > 0 Then
        Set MRU = m.Mrus("PercentR")
        If MRU.LastUsed <> cboPercentR.Text Then
            MRU.LastUsed = cboPercentR.Text
            MRU.Save
        End If
    End If

    If Len(cboOverall.Text) > 0 Then
        Set MRU = m.Mrus("Overall")
        If MRU.LastUsed <> cboOverall.Text Then
            MRU.LastUsed = cboOverall.Text
            MRU.Save
        End If
    End If

    If Len(cboEntryDirection.Text) > 0 Then
        Set MRU = m.Mrus("EntryDirection")
        If MRU.LastUsed <> cboEntryDirection.Text Then
            MRU.LastUsed = cboEntryDirection.Text
            MRU.Save
        End If
    End If

    If Len(cboEntryTrigger.Text) > 0 Then
        Set MRU = m.Mrus("EntryTrigger")
        If MRU.LastUsed <> cboEntryTrigger.Text Then
            MRU.LastUsed = cboEntryTrigger.Text
            MRU.Save
        End If
    End If

    If Len(cboExitTrigger.Text) > 0 Then
        Set MRU = m.Mrus("ExitTrigger")
        If MRU.LastUsed <> cboExitTrigger.Text Then
            MRU.LastUsed = cboExitTrigger.Text
            MRU.Save
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.SaveMoneyCodeMru"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupSymbol
'' Description: Allow the user to look up a symbol
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LookupSymbol(Optional ByVal KeyAscii As Long = 0&)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol the user selected
    
    strSymbol = g.AppBridge.LookupSymbol(KeyAscii, txtSymbol(m.nMode).Text)
    If Len(strSymbol) > 0 Then
        txtSymbol(m.nMode).Text = strSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.LookupSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMrus
'' Description: Load the MRU's
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadMrus()
On Error GoTo ErrSection:

    LoadMoneyCodeMrus
    LoadCustomChecklistMrus

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.LoadMrus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMoneyCodeMrus
'' Description: Load the MRU's for the Money Code checklist
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadMoneyCodeMrus()
On Error GoTo ErrSection:

    Dim MRU As cMostRecentlyUsed        ' Most recently used object
    
    Set MRU = New cMostRecentlyUsed
    MRU.Init "COT", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboCot
    m.Mrus.Add MRU, "COT"
    
    Set MRU = New cMostRecentlyUsed
    MRU.Init "LargeTraders", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboLargeTraders
    m.Mrus.Add MRU, "LargeTraders"
    
    Set MRU = New cMostRecentlyUsed
    MRU.Init "Advisor", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboAdvisor
    m.Mrus.Add MRU, "Advisor"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "Agricultural", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboAgricultural
    m.Mrus.Add MRU, "Agricultural"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "SeasonalSetup", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboSeasonalSetup
    m.Mrus.Add MRU, "SeasonalSetup"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "SeasonalDirection", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboSeasonalDirection
    m.Mrus.Add MRU, "SeasonalDirection"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "OpenInterest", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboOpenInterest
    m.Mrus.Add MRU, "OpenInterest"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "Accumulation", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboAccumulation
    m.Mrus.Add MRU, "Accumulation"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "PercentR", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboPercentR
    m.Mrus.Add MRU, "PercentR"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "Overall", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboOverall
    m.Mrus.Add MRU, "Overall"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "EntryDirection", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboEntryDirection
    m.Mrus.Add MRU, "EntryDirection"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "EntryTrigger", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboEntryTrigger
    m.Mrus.Add MRU, "EntryTrigger"
        
    Set MRU = New cMostRecentlyUsed
    MRU.Init "ExitTrigger", 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    LoadMruCombo MRU, cboExitTrigger
    m.Mrus.Add MRU, "ExitTrigger"

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.LoadMoneyCodeMrus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCustomChecklistMrus
'' Description: Load the MRU's for the custom checklist
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCustomChecklistMrus()
On Error GoTo ErrSection:

    Dim MRU As cMostRecentlyUsed        ' Most recently used object
    Dim strKey As String                ' Key into the collection

    strKey = "WeeklySetup"
    Set MRU = New cMostRecentlyUsed
    MRU.Init strKey, 20, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    m.Mrus.Add MRU, strKey

    strKey = "WeeklyValue"
    Set MRU = New cMostRecentlyUsed
    MRU.Init strKey, 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    m.Mrus.Add MRU, strKey

    strKey = "DailySetup"
    Set MRU = New cMostRecentlyUsed
    MRU.Init strKey, 20, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    m.Mrus.Add MRU, strKey

    strKey = "DailyValue"
    Set MRU = New cMostRecentlyUsed
    MRU.Init strKey, 10, m.strJournalIni, m.strDefaultsIni
    MRU.Load
    m.Mrus.Add MRU, strKey
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.LoadCustomChecklistMrus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ComboDropdown
'' Description: When the user drops down the combo, fill the list from the MRU
'' Inputs:      Key, Combo Box
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ComboDropDown(ByVal strKey As String, MruCombo As ComboBox)
On Error GoTo ErrSection:

    Dim strText As String               ' Text that was in the box

    ' The DropDown event happens before the LostFocus event of the other
    ' control.  Do a DoEvents here to allow the LostFocus event of the other
    ' control to happen first so that it can save off the MRU list for us
    ' to use here if the previous control was one of the other combo boxes in
    ' this family...
    DoEvents

    strText = MruCombo.Text
    LoadMruCombo m.Mrus(strKey), MruCombo
    
    If Len(strText) > 0 Then
        MruCombo.Text = strText
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.ComboDropDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ComboLostFocus
'' Description: When the combo loses focus, save the changes to the MRU list
'' Inputs:      Key, Text
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ComboLostFocus(ByVal strKey As String, ByVal strComboText As String)
On Error GoTo ErrSection:

    Dim MRU As cMostRecentlyUsed        ' Most recently used object
    
    If (m.bVisible = True) And (strComboText <> m.strComboText) Then
        Set MRU = m.Mrus(strKey)
        If MRU.LastUsed <> strComboText Then
            MRU.LastUsed = strComboText
            MRU.Save
            
            Set m.Mrus(strKey) = MRU
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournal.ComboLostFocus"
    
End Sub

