VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTTEditFill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraFill 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
      Caption         =   "frmTTEditFill.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditFill.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditFill.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraInformation 
         Height          =   1575
         Left            =   0
         TabIndex        =   1
         Top             =   0
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
         Caption         =   "frmTTEditFill.frx":0068
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTEditFill.frx":00AA
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditFill.frx":00CA
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraBuySell 
            Height          =   195
            Left            =   120
            TabIndex        =   2
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
            Caption         =   "frmTTEditFill.frx":00E6
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTEditFill.frx":0112
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0132
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optBuy 
               Height          =   220
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   735
               _ExtentX        =   1296
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
               Caption         =   "frmTTEditFill.frx":014E
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   16711680
               Pressed         =   0   'False
               Tip             =   "frmTTEditFill.frx":0176
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditFill.frx":0196
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optSell 
               Height          =   220
               Left            =   780
               TabIndex        =   4
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
               Caption         =   "frmTTEditFill.frx":01B2
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   255
               Pressed         =   0   'False
               Tip             =   "frmTTEditFill.frx":01DC
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditFill.frx":01FC
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniTextBoxXP txtQuantity 
            Height          =   315
            Left            =   840
            TabIndex        =   6
            Top             =   630
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditFill.frx":0218
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
            Tip             =   "frmTTEditFill.frx":0242
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0262
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
            Height          =   255
            Left            =   3870
            TabIndex        =   10
            Top             =   660
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
            Caption         =   "frmTTEditFill.frx":027E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTTEditFill.frx":02B0
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":02D0
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPrice 
            Height          =   285
            Left            =   660
            TabIndex        =   12
            Top             =   1125
            Width           =   1020
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditFill.frx":02EC
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
            Tip             =   "frmTTEditFill.frx":031C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":033C
         End
         Begin gdOCX.gdScrollBar sbQuantity 
            Height          =   360
            Left            =   1680
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   600
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
         Begin gdOCX.gdScrollBar sbPrice 
            Height          =   360
            Left            =   1680
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1080
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
         Begin gdOCX.gdSelectDate gdFillDate 
            Height          =   315
            Left            =   2520
            TabIndex        =   15
            Top             =   1110
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            ShowDayOfWeek   =   0   'False
            ShowTime        =   1
         End
         Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
            Height          =   315
            Left            =   2700
            TabIndex        =   9
            Top             =   630
            Width           =   1440
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditFill.frx":0358
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
            Tip             =   "frmTTEditFill.frx":038C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":03AC
         End
         Begin HexUniControls.ctlUniLabelXP lblSymbol 
            Height          =   195
            Left            =   2040
            Top             =   690
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
            Caption         =   "frmTTEditFill.frx":03C8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":03F8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0418
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblQuantity 
            Height          =   195
            Left            =   120
            Top             =   690
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
            Caption         =   "frmTTEditFill.frx":0434
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":0468
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0488
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDate 
            Height          =   255
            Left            =   2040
            Top             =   1140
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
            Caption         =   "frmTTEditFill.frx":04A4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":04D0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":04F0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPrice 
            Height          =   255
            Left            =   120
            Top             =   1140
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
            Caption         =   "frmTTEditFill.frx":050C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":053A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":055A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraMisc 
         Height          =   2475
         Left            =   0
         TabIndex        =   21
         Top             =   2520
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
         Caption         =   "frmTTEditFill.frx":0576
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTEditFill.frx":05BA
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditFill.frx":05DA
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboImageXP cboAccounts 
            Height          =   315
            Left            =   900
            TabIndex        =   5
            Top             =   1740
            Width           =   2895
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
            Tip             =   "frmTTEditFill.frx":05F6
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0616
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboLots 
            Height          =   315
            Left            =   900
            TabIndex        =   8
            Top             =   1320
            Width           =   2895
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
            Tip             =   "frmTTEditFill.frx":0632
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0652
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboCategory 
            Height          =   315
            Left            =   900
            TabIndex        =   11
            Top             =   600
            Width           =   2895
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
            Tip             =   "frmTTEditFill.frx":066E
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":068E
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtNotes 
            Height          =   1335
            Left            =   900
            TabIndex        =   14
            Top             =   1020
            Width           =   4275
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditFill.frx":06AA
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
            Tip             =   "frmTTEditFill.frx":06CA
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":06EA
         End
         Begin HexUniControls.ctlUniTextBoxXP txtFees 
            Height          =   315
            Left            =   1860
            TabIndex        =   23
            Top             =   240
            Width           =   855
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditFill.frx":0706
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
            Tip             =   "frmTTEditFill.frx":0726
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0746
         End
         Begin HexUniControls.ctlUniLabelXP lblAccount 
            Height          =   255
            Left            =   120
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
            Caption         =   "frmTTEditFill.frx":0762
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":0794
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":07B4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblLot 
            Height          =   255
            Left            =   120
            Top             =   1380
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
            Caption         =   "frmTTEditFill.frx":07D0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":07FA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":081A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblCategory 
            Height          =   255
            Left            =   120
            Top             =   660
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
            Caption         =   "frmTTEditFill.frx":0836
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":086A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":088A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblNotes 
            Height          =   255
            Left            =   120
            Top             =   1020
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
            Caption         =   "frmTTEditFill.frx":08A6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":08D4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":08F4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblFees 
            Height          =   255
            Left            =   120
            Top             =   270
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
            Caption         =   "frmTTEditFill.frx":0910
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":095C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":097C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraIdentification 
         Height          =   735
         Left            =   0
         TabIndex        =   16
         Top             =   1680
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
         Caption         =   "frmTTEditFill.frx":0998
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTEditFill.frx":09E4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditFill.frx":0A04
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtBrokerOrderID 
            Height          =   315
            Left            =   3300
            TabIndex        =   20
            Top             =   240
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditFill.frx":0A20
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
            Tip             =   "frmTTEditFill.frx":0A40
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0A60
         End
         Begin HexUniControls.ctlUniTextBoxXP txtBrokerFillID 
            Height          =   315
            Left            =   600
            TabIndex        =   18
            Top             =   240
            Width           =   1875
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditFill.frx":0A7C
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
            Tip             =   "frmTTEditFill.frx":0A9C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0ABC
         End
         Begin HexUniControls.ctlUniLabelXP lblBrokerOrderID 
            Height          =   255
            Left            =   2580
            Top             =   270
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
            Caption         =   "frmTTEditFill.frx":0AD8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":0B0C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0B2C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBrokerFillID 
            Height          =   255
            Left            =   120
            Top             =   270
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
            Caption         =   "frmTTEditFill.frx":0B48
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditFill.frx":0B7A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditFill.frx":0B9A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   855
      Left            =   5640
      TabIndex        =   17
      Top             =   120
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
      Caption         =   "frmTTEditFill.frx":0BB6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditFill.frx":0BE2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditFill.frx":0C02
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   480
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
         Caption         =   "frmTTEditFill.frx":0C1E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditFill.frx":0C4C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditFill.frx":0C6C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   0
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
         Caption         =   "frmTTEditFill.frx":0C88
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditFill.frx":0CAE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditFill.frx":0CCE
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmTTEditFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTTEditFill.frm
'' Description: Allows the user to edit a fill for an order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/11/2010   DAJ         Use global Trading Items collection, Trade Bars
'' 07/19/2011   DAJ         ShowMe now takes automated trading ID default
'' 02/11/2014   DAJ         Changed "Fees" label to "Commissions and Fees"
'' 06/06/2014   DAJ         Added ability to create/edit a fill from Cattle stuff
'' 07/10/2014   DAJ         Disable the date/time control if non-manual live fill
'' 01/16/2015   DAJ         In lookup symbol, make sure to add bars to the stream
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?

    Fill As cPtFill                     ' Fill being edited
    CattleFill As cBrokerMessage        ' Cattle fill being edited
    Bars As cGdBars                     ' Bars structure
    Price As cPriceEditor               ' Editor for price
    Qty As cPriceEditor                 ' Editor for quantity
    
    lAccountID As Long                  ' Account ID for this fill
    bForCattle As Boolean               ' Are we doing this for a cattle fill?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Fill, Symbol, Account ID, Use Previous Close Time?, Auto Trade ID
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Fill As cPtFill, ByVal strSymbol As String, ByVal lAccountID As Long, Optional ByVal bUsePrevCloseTime As Boolean = False, Optional ByVal lAutoTradeID As Long = 0&) As Boolean
On Error GoTo ErrSection:

    Dim dPrevCloseTime As Double        ' Time of the previous session close
    Dim dCurrentTime As Double          ' Current time
    Dim strCurrentTime As String        ' Current time as a string

    Caption = "Edit Fill"
    
    Set m.Fill = Fill
    Set m.CattleFill = Nothing
    m.bForCattle = False
    
    Set m.Bars = New cGdBars
    SetBarProperties m.Bars, strSymbol
    
    Set m.Price = New cPriceEditor
    m.Price.Init sbPrice, txtPrice, m.Bars
    
    Set m.Qty = New cPriceEditor
    m.Qty.Init sbQuantity, txtQuantity, Nothing
    
    m.lAccountID = lAccountID
    
    lblCategory.Visible = Not m.bForCattle
    cboCategory.Visible = Not m.bForCattle
    lblNotes.Visible = Not m.bForCattle
    txtNotes.Visible = Not m.bForCattle
    lblLot.Visible = m.bForCattle
    cboLots.Visible = m.bForCattle
    lblAccount.Visible = m.bForCattle
    cboAccounts.Visible = m.bForCattle
    
    LoadCategoryCombo
    ''cboCategory.Locked = True
    ''cboCategory.BackColor = vbButtonFace

    ' If this is a new fill, then set the controls to default values...
    If Fill.FillID = 0# Then
        optBuy.Value = True
        optSell.Value = False
        m.Qty.Price = 0
        txtSymbol.Text = strSymbol
        m.Price.Price = 0
        
        If bUsePrevCloseTime = False Then
            gdFillDate.Value = CurrentTime
        Else
            dCurrentTime = CurrentTime
            dPrevCloseTime = Int(dCurrentTime) + (m.Bars.Prop(eBARS_EndTime) / 1440#)
            dPrevCloseTime = ConvertTimeZone(dPrevCloseTime, m.Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            Do While dPrevCloseTime > dCurrentTime
                dPrevCloseTime = dPrevCloseTime - 1
            Loop
            gdFillDate.Value = dPrevCloseTime
        End If
        
        strCurrentTime = Format(CurrentTime, "YYYYMMDD_HHMMSS")
        txtBrokerFillID.Text = "MF_" & strCurrentTime
        txtBrokerOrderID.Text = "MO_" & strCurrentTime
        
        txtFees.Text = ""
        SetCategoryCombo lAutoTradeID
        txtNotes.Text = ""
        
        Enable optBuy, True
        Enable optSell, True
        Enable lblSymbol, True
        Enable txtSymbol, True
        Enable cmdLookup, True
        Enable lblBrokerFillID, True
        Enable txtBrokerFillID, True
        Enable lblBrokerOrderID, True
        Enable txtBrokerOrderID, True
        Enable lblDate, True
        Enable gdFillDate, True
        
    ' Otherwise fill in with the values from the fill passed in...
    Else
        optBuy.Value = Fill.Buy
        optSell.Value = Not Fill.Buy
        m.Qty.Price = Fill.Quantity
        txtSymbol.Text = Fill.Symbol
        m.Price.Price = Fill.Price
        gdFillDate.Value = Fill.FillDate
        
        txtBrokerFillID = Fill.BrokerID
        txtBrokerOrderID = Fill.BrokerOrderID
        
        txtFees.Text = Format(Fill.Fees, "$#,##0.00")
        SetCategoryCombo Fill.AutoTradingItemID
        txtNotes.Text = Fill.Notes
    
        Enable optBuy, Fill.IsManual
        Enable optSell, Fill.IsManual
        Enable lblSymbol, Fill.IsManual
        Enable txtSymbol, Fill.IsManual
        Enable cmdLookup, Fill.IsManual
        Enable lblBrokerFillID, Fill.IsManual
        Enable txtBrokerFillID, Fill.IsManual
        Enable lblBrokerOrderID, Fill.IsManual
        Enable txtBrokerOrderID, Fill.IsManual
        
        ' DAJ 07/10/2014: Only allow the date to be changed on an existing fill if it was
        ' a manually created fill OR the fill is in a Genesis simulated account...
        Enable lblDate, Fill.IsManual Or (Not g.Broker.IsLiveAccount(Fill.Broker))
        Enable gdFillDate, Fill.IsManual Or (Not g.Broker.IsLiveAccount(Fill.Broker))
    End If
    
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK = True Then
        Fill.AccountID = lAccountID
        Fill.AutoTradingItemID = cboCategory.ItemData(cboCategory.ListIndex)
        Fill.Buy = optBuy.Value
        Fill.Quantity = m.Qty.Price
        Fill.SymbolOrSymbolID = txtSymbol.Text
        Fill.Price = m.Price.Price
        Fill.FillDate = gdFillDate.Value
        If Fill.FillID = 0& Then Fill.IsManual = True
        Fill.SessionDate = m.Bars.SessionDateForTradeTime(ConvertBrokerDate(gdFillDate.Value, Fill.Broker, Fill.Symbol, False))
        
        Fill.BrokerID = txtBrokerFillID.Text
        Fill.BrokerOrderID = txtBrokerOrderID.Text
        
        Fill.Fees = ValOfText(txtFees.Text)
        Fill.Notes = txtNotes.Text
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTEditFill.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeCattle
'' Description: Setup and show the form for a cattle fill
'' Inputs:      Cattle Fill
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeCattle(CattleFill As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    Dim dPrevCloseTime As Double        ' Time of the previous session close
    Dim dCurrentTime As Double          ' Current time
    Dim strSymbol As String             ' Symbol to use
    Dim strCurrentTime As String        ' Current time as a string

    Caption = "Edit Fill"
    
    Set m.Fill = Nothing
    Set m.CattleFill = CattleFill
    m.bForCattle = True
    
    If Len(CattleFill("Symbol")) = 0 Then
        strSymbol = ConvertToTradeSymbol("LE-067", CurrentTime)
    Else
        strSymbol = CattleFill("Symbol")
    End If
    
    Set m.Bars = New cGdBars
    SetBarProperties m.Bars, strSymbol
    
    Set m.Price = New cPriceEditor
    m.Price.Init sbPrice, txtPrice, m.Bars
    
    Set m.Qty = New cPriceEditor
    m.Qty.Init sbQuantity, txtQuantity, Nothing
    
    m.lAccountID = 0&
    
    lblCategory.Visible = Not m.bForCattle
    cboCategory.Visible = Not m.bForCattle
    lblNotes.Visible = Not m.bForCattle
    txtNotes.Visible = Not m.bForCattle
    lblLot.Visible = m.bForCattle
    cboLots.Visible = m.bForCattle
    lblAccount.Visible = m.bForCattle
    cboAccounts.Visible = m.bForCattle
    
    g.CattleBridge.LoadLotsCombo cboLots, CattleFill("FeedYardLotID")
    g.CattleBridge.LoadAccountsCombo cboAccounts, CattleFill("BrokerAccountID")
    
    ' If this is a new fill, then set the controls to default values...
    If Len(CattleFill("BrokerFillID")) = 0 Then
        optBuy.Value = True
        optSell.Value = False
        m.Qty.Price = 0
        txtSymbol.Text = strSymbol
        m.Price.Price = 0
        
        gdFillDate.Value = CurrentTime
        strCurrentTime = Format(gdFillDate.Value, "YYYYMMDD_HHMMSS")
        
        txtBrokerFillID.Text = "MF_" & strCurrentTime
        txtBrokerOrderID.Text = "MO_" & strCurrentTime
        
        txtFees.Text = ""
        'SetCategoryCombo lAutoTradeID
        txtNotes.Text = ""
        
        Enable optBuy, True
        Enable optSell, True
        Enable lblSymbol, True
        Enable txtSymbol, True
        Enable cmdLookup, True
        Enable lblBrokerFillID, True
        Enable txtBrokerFillID, True
        Enable lblBrokerOrderID, True
        Enable txtBrokerOrderID, True
        
    ' Otherwise fill in with the values from the fill passed in...
    Else
        optBuy.Value = (CattleFill("IsBuy") = "1")
        optSell.Value = (CattleFill("IsBuy") = "0")
        m.Qty.Price = CLng(Val(CattleFill("Quantity")))
        txtSymbol.Text = CattleFill("Symbol")
        m.Price.Price = Val(CattleFill("Price"))
        gdFillDate.Value = Val(CattleFill("FillTime"))
        
        txtBrokerFillID.Text = CattleFill("BrokerFillID")
        txtBrokerOrderID.Text = CattleFill("BrokerOrderID")
        
        txtFees.Text = Format(Val(CattleFill("Commission")), "$#,##0.00")
        'SetCategoryCombo Fill.AutoTradingItemID
        txtNotes.Text = ""
    
        Enable optBuy, True
        Enable optSell, True
        Enable lblSymbol, True
        Enable txtSymbol, True
        Enable cmdLookup, True
        Enable lblBrokerFillID, True
        Enable txtBrokerFillID, True
        Enable lblBrokerOrderID, True
        Enable txtBrokerOrderID, True
    End If
    
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK = True Then
        'Fill.AccountID = lAccountID
        If optBuy.Value = True Then
            CattleFill.Add "IsBuy", "1"
        Else
            CattleFill.Add "IsBuy", "0"
        End If
        CattleFill.Add "Quantity", Str(m.Qty.Price)
        CattleFill.Add "Symbol", txtSymbol.Text
        CattleFill.Add "Price", Str(m.Price.Price)
        CattleFill.Add "FillTime", Str(gdFillDate.Value)
        CattleFill.Add "BrokerFillID", txtBrokerFillID.Text
        CattleFill.Add "BrokerOrderID", txtBrokerOrderID.Text
        CattleFill.Add "Commission", Str(ValOfText(txtFees.Text))
        CattleFill.Add "FeedYardLotID", Str(cboLots.ItemData(cboLots.ListIndex))
        CattleFill.Add "BrokerAccountID", Str(cboAccounts.ItemData(cboAccounts.ListIndex))
    End If
    
    ShowMeCattle = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTEditFill.ShowMeCattle"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allows the user to cancel out of the dialog
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
    RaiseError "frmTTEditFill.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: Allow the user to change the symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookup_Click()
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.cmdLookup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allows the user to exit dialog and keep changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If (optBuy.Value = False) And (optSell.Value = False) Then
        MoveFocus optBuy
        InfBox "Please specify whether this fill is a Buy or a Sell.", "!", , "Fill Error"
    
    ElseIf m.Qty.Price = 0 Then
        MoveFocus txtQuantity
        InfBox "Please specify a valid quantity for the fill.", "!", , "Fill Error"
    
    ElseIf Len(txtSymbol.Text) = 0 Then
        MoveFocus txtSymbol
        InfBox "Please specify a symbol for the fill.", "!", , "Fill Error"
    
    ElseIf m.Price.Price = 0 Then
        MoveFocus txtPrice
        InfBox "Please specify a valid price for the fill.", "!", , "Fill Error"
    
    ElseIf Len(Trim(txtBrokerFillID.Text)) = 0 Then
        MoveFocus txtBrokerFillID
        InfBox "Please specify a Broker Fill ID for the fill.", "!", , "Fill Error"

    ElseIf Len(Trim(txtBrokerOrderID.Text)) = 0 Then
        MoveFocus txtBrokerOrderID
        InfBox "Please specify a Broker Order ID for the fill.", "!", , "Fill Error"
    
    ElseIf BrokerFillIdExists = True Then
        MoveFocus txtBrokerFillID
        InfBox "A fill with the Broker Fill ID '" & txtBrokerFillID.Text & "' already exists in a " & BrokerName & " account.  Please specify a unique Broker Fill ID for the fill.", "!", , "Fill Error"
        
    Else
        m.bOK = True
        Me.Hide
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: If the user pressed F1, show the help menu
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
    RaiseError "frmTTEditFill.Form_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form and controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strTemp As String               ' Temporary string to get ini file property

    g.Styler.StyleForm Me
    
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)

    strTemp = GetIniFileProperty("TTEditFill", "", "Placement", g.strIniFile)
    If Len(strTemp) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strTemp
    End If
    
    lblLot.Top = lblCategory.Top
    cboLots.Top = cboCategory.Top
    lblAccount.Top = lblNotes.Top
    cboAccounts.Top = txtNotes.Top
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Cancel the unload and handle it ourselves
'' Inputs:      Whether or not to cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.Form_QueryLoad"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    ''If LimitFormSize(Me, 5715, 3825) Then Exit Sub
    If LimitFormSize(Me, 7080, 5340) Then Exit Sub

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save some settings when the form gets unloaded
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "TTEditFill", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.Form_Unload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtBrokerFillID_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtBrokerFillID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtBrokerFillID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.txtBrokerFillID_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtBrokerOrderID_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtBrokerOrderID_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtBrokerOrderID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.txtBrokerOrderID_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFees_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFees_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtFees

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.txtFees_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFees_LostFocus
'' Description: Reformat the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFees_LostFocus()
On Error GoTo ErrSection:

    txtFees.Text = Format(ValOfText(txtFees.Text), "$#,##0.00")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.txtFees_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPrice_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPrice_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPrice

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.txtPrice_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtQuantity_GotFocus
'' Description: Select all of the text when the control gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtQuantity_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQuantity

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.txtQuantity_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_Click
'' Description: Allow the user to change the symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.txtSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_GotFocus
'' Description: Select all of the text when the control gets the focus
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
    RaiseError "frmTTEditFill.txtSymbol_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_KeyPress
'' Description: Allow the user the change the symbol
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LookupSymbol KeyAscii
    KeyAscii = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.txtSymbol_KeyPress", 0
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupSymbol
'' Description: Lookup a symbol for the user to trade
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LookupSymbol(Optional ByVal KeyAscii As Long = 0&)
On Error GoTo ErrSection:

    Dim astrSymbol As New cGdArray      ' Array to get lookup symbol from
    Dim strSymbol As String
    
    If KeyAscii = 0& Then
        Set astrSymbol = frmSymbolSelector.ShowMe(txtSymbol.Text, False, True, "Symbol for this Fill", , , True)
    Else
        Set astrSymbol = frmSymbolSelector.ShowMe(Chr(KeyAscii), False, True, "Symbol for this Fill", , False, True)
    End If
    If astrSymbol.Size > 0 Then
        strSymbol = ConvertToTradeSymbol(astrSymbol(0), gdFillDate.Value)
        'If (InStr(Trim(astrSymbol(0)), " ") > 0) And (InStr(Trim(astrSymbol(0)), "-") > 0) And (Not FileExist(AddSlash(App.Path) & "TradeFO.FLG")) Then
        '    InfBox "Future Options are not currently allowed|to be traded", "!", , "Fill Error"
        'ElseIf strSymbol <> UCase(Trim(txtSymbol.Text)) Then
        If strSymbol <> UCase(Trim(txtSymbol.Text)) Then
            txtSymbol.Text = strSymbol
            
            ' DAJ 01/16/2015: We need to add to stream here because if we don't, the bars in the
            ' trade console collection get added, but never added to the stream and because of that,
            ' any subsequent orders or positions for this symbol won't get updated via the stream...
            Set m.Bars = GetTradeBars(strSymbol, True)
            'Set m.Bars = GetTradeBars(strSymbol, False)
            'DM_GetBars m.Bars, txtSymbol.Text
            'g.RealTime.SpliceBars m.Bars
            
            m.Price.Init sbPrice, txtPrice, m.Bars, m.Bars(eBARS_Close, m.Bars.Size - 1)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.LookupSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCategoryCombo
'' Description: Load up the category combo list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCategoryCombo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With cboCategory
        .AddItem "Manual"
        .ItemData(.NewIndex) = 0&
        
        For lIndex = 1 To g.TradingItems.Count
            If g.TradingItems(lIndex).AccountID = m.lAccountID Then
                .AddItem g.TradingItems(lIndex).Name
                .ItemData(.NewIndex) = g.TradingItems(lIndex).AutoTradeItemID
            End If
        Next lIndex
        
        .ListIndex = 0
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.LoadCategoryCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCategoryCombo
'' Description: Set the category combo to the given ID
'' Inputs:      Category ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCategoryCombo(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With cboCategory
        For lIndex = 0 To .ListCount - 1
            If .ItemData(lIndex) = lCategoryID Then
                .ListIndex = lIndex
                Exit For
            End If
        Next lIndex
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditFill.SetCategoryCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerFillIdExists
'' Description: Does the broker fill ID in the text box already exist?
'' Inputs:      None
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerFillIdExists() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    Dim strAccountID As String          ' Account ID

    bReturn = False

    If m.bForCattle = True Then
        If Not g.CattleBridge Is Nothing Then
            strAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
            bReturn = g.CattleBridge.CattleFillExists(txtBrokerFillID.Text, strAccountID)
        End If
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT tblFills.*, tblAccounts.AccountType " & _
                    "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID " & _
                    "WHERE tblFills.BrokerFillID='" & txtBrokerFillID.Text & "' AND tblAccounts.AccountType=" & Str(g.Broker.AccountTypeForID(m.lAccountID)) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!FillID <> m.Fill.FillID Then
                bReturn = True
            End If
            
            rs.MoveNext
        Loop
    End If
    
    BrokerFillIdExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditFill.BrokerFillIdExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerName
'' Description: Determine the appropriate broker name
'' Inputs:      None
'' Returns:     Broker Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerName() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strAccount As String            ' Account from the combo
    Dim iPos As Integer                 ' Position of something in a string
    Dim iLen As Integer                 ' Length of the string
    
    If m.bForCattle Then
        strAccount = cboAccounts.Text
        iPos = InStr(strAccount, "(")
        iLen = (Len(strAccount) - 1) - iPos
        strReturn = Mid(strAccount, iPos + 1, iLen)
    Else
        strReturn = g.Broker.BrokerName(g.Broker.AccountTypeForID(m.lAccountID))
    End If
    
    BrokerName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditFill.BrokerName"
    
End Function

