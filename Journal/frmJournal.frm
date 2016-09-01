VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmJournal 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraOptNavImage 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   5100
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmJournal.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmJournal.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmJournal.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtOptNavImageCaption 
         Height          =   285
         Left            =   1860
         TabIndex        =   2
         Top             =   0
         Width           =   3795
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmJournal.frx":0068
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
         Tip             =   "frmJournal.frx":0088
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":00A8
      End
      Begin HexUniControls.ctlUniCheckXP chkOptNavImage 
         Height          =   220
         Left            =   0
         TabIndex        =   4
         Top             =   15
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "frmJournal.frx":00C4
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmJournal.frx":010E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":012E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtOptNavImageFilename 
         Height          =   285
         Left            =   5820
         TabIndex        =   7
         Top             =   0
         Width           =   1395
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmJournal.frx":014A
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
         Tip             =   "frmJournal.frx":016A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":018A
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraImage 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   4740
      Width           =   7275
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmJournal.frx":01A6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmJournal.frx":01DC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmJournal.frx":01FC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkAttachImage 
         Height          =   220
         Left            =   0
         TabIndex        =   17
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
         Caption         =   "frmJournal.frx":0218
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmJournal.frx":025E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":027E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboCharts 
         Height          =   315
         Left            =   1860
         TabIndex        =   18
         Top             =   0
         Width           =   5415
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
         Tip             =   "frmJournal.frx":029A
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":02BA
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraNotes 
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   3420
      Width           =   7275
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmJournal.frx":02D6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmJournal.frx":0302
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmJournal.frx":0322
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtNotes 
         Height          =   915
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   7275
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmJournal.frx":033E
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
         Tip             =   "frmJournal.frx":035E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":037E
      End
      Begin HexUniControls.ctlUniLabelXP lblNotes 
         Height          =   195
         Left            =   0
         Top             =   0
         Width           =   4275
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournal.frx":039A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmJournal.frx":03C8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":03E8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraQuestions 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7275
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmJournal.frx":0404
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmJournal.frx":0430
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmJournal.frx":0450
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboBoxXP cboThoughts 
         Height          =   315
         Left            =   0
         TabIndex        =   10
         Top             =   1620
         Width           =   7215
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
         Tip             =   "frmJournal.frx":046C
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
         MouseIcon       =   "frmJournal.frx":048C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniComboBoxXP cboReasons 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   900
         Width           =   7215
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
         Tip             =   "frmJournal.frx":04A8
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
         MouseIcon       =   "frmJournal.frx":04C8
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniComboBoxXP cboFeelings 
         Height          =   315
         Left            =   660
         TabIndex        =   6
         Top             =   240
         Width           =   6555
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
         Tip             =   "frmJournal.frx":04E4
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
         MouseIcon       =   "frmJournal.frx":0504
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboEmotions 
         Height          =   315
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   615
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
         Tip             =   "frmJournal.frx":0520
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":0540
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboAction 
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   2280
         Width           =   1935
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
         Tip             =   "frmJournal.frx":055C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":057C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdCurrentTime 
         Height          =   315
         Left            =   3720
         TabIndex        =   9
         Top             =   2280
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         ShowTime        =   1
      End
      Begin HexUniControls.ctlUniLabelXP lblCurrentTime 
         Height          =   195
         Left            =   3720
         Top             =   2040
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
         Caption         =   "frmJournal.frx":0598
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmJournal.frx":05D4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":05F4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAction 
         Height          =   195
         Left            =   0
         Top             =   2040
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
         Caption         =   "frmJournal.frx":0610
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmJournal.frx":067E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":069E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblThoughts 
         Height          =   195
         Left            =   0
         Top             =   1380
         Width           =   4275
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournal.frx":06BA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmJournal.frx":072A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":074A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblReasons 
         Height          =   195
         Left            =   0
         Top             =   660
         Width           =   4275
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournal.frx":0766
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmJournal.frx":07C4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":07E4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFeeling 
         Height          =   195
         Left            =   0
         Top             =   0
         Width           =   4275
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournal.frx":0800
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmJournal.frx":089E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":08BE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraOrderInfo 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7275
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmJournal.frx":08DA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmJournal.frx":0906
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmJournal.frx":0926
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP lblBrokerID 
         Height          =   195
         Left            =   0
         Top             =   240
         Width           =   7215
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournal.frx":0942
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmJournal.frx":0976
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":0996
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOrder 
         Height          =   195
         Left            =   0
         Top             =   0
         Width           =   7215
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournal.frx":09B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmJournal.frx":09DE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":09FE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1800
      TabIndex        =   19
      Top             =   5580
      Width           =   3855
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmJournal.frx":0A1A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmJournal.frx":0A46
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmJournal.frx":0A66
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPrint 
         Height          =   495
         Left            =   1320
         TabIndex        =   11
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
         Caption         =   "frmJournal.frx":0A82
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmJournal.frx":0AAE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":0ACE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   2640
         TabIndex        =   14
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
         Caption         =   "frmJournal.frx":0AEA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmJournal.frx":0B18
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":0B38
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
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
         Caption         =   "frmJournal.frx":0B54
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmJournal.frx":0B7A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmJournal.frx":0B9A
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmJournal.frm
'' Description: Allows the user to create or edit a journal entry for an order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/20/2009   DAJ         Added support for the emotion number
'' 04/22/2009   DAJ         Fixed load if strings are Null, EmotionNumber = -1
'' 04/23/2009   DAJ         Changed from database to new journal objects
'' 03/11/2010   DAJ         Added the ShowMeForOrderID function
'' 09/22/2011   DAJ         Added the ability to attach a chart image
'' 09/22/2011   DAJ         Change format in charts combo
'' 09/22/2011   DAJ         Set the JournalDate in the journal object on a save
'' 09/23/2011   DAJ         Send journal to date journals form if loaded
'' 09/23/2011   DAJ         Fix for exporting non-active maximized chart
'' 10/11/2011   DAJ         Set defaults for combo boxes, add attach chart check box
'' 10/12/2011   DAJ         Implement MRU list for feelings, reasons, and thoughts
'' 10/13/2011   DAJ         Changed INI file for MRU persistence
'' 10/14/2011   DAJ         Auto drop-down combo boxes when they get focus for a new journal
'' 01/25/2012   DAJ         Utilize the Ini file for MRU defaults
'' 01/30/2012   DAJ         Option Nav Journal Image
'' 02/17/2012   DAJ         Fixed "With Block" error in Save when Chart attached
'' 03/13/2012   DAJ         Fixed off by one error in SetMruCombo
'' 03/19/2012   DAJ         Added Symbol/Symbol ID to the order journal object
'' 07/30/2013   DAJ         Moved out chart export code
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/08/2014   DAJ         Use NavCore Image List; Use newer place/save form
'' 10/24/2014   DAJ         Core Application functions for DLL's
'' 05/18/2015   DAJ         Pass frmPrintPreview.vp to DoPrintHeader
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user choose OK?
    lNewOrderID As Long                 ' New order ID in case changed while form was up
    lJournalID As Long                  ' Journal ID
    Journal As cJournal                 ' Journal object
    
    Order As cBrokerMessage             ' Order object for the form
    
    FeelingsMru As cMostRecentlyUsed    ' Most recently used feelings
    ReasonsMru As cMostRecentlyUsed     ' Most recently used reasons
    ThoughtsMru As cMostRecentlyUsed    ' Most recently used thoughts
End Type
Private m As mPrivate

Public Property Get OrderID() As Long
    If m.Order Is Nothing Then
        OrderID = 0&
    Else
        OrderID = CLng(Val(m.Order("OrderID")))
    End If
End Property

Public Property Get NewOrderID() As Long
    NewOrderID = m.lNewOrderID
End Property
Public Property Let NewOrderID(ByVal lNewOrderID As Long)
    m.lNewOrderID = lNewOrderID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and set controls and show the form
'' Inputs:      Order, Journal ID, Journal
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal Order As cBrokerMessage, Optional ByVal lJournalID As Long = -1&, Optional Journal As cJournal = Nothing) As Boolean
On Error GoTo ErrSection:

    m.lNewOrderID = 0&
    Set m.Order = Order
    If m.Order Is Nothing Then Set m.Order = New cBrokerMessage
    m.lJournalID = lJournalID
    Set m.Journal = Journal
    
    Caption = "Order Journal"
    
    If Len(Order("Symbol")) > 0 Then
        lblOrder.Caption = "Order: " & Order("OrderText") & " in Account " & Order("AccountName")
    Else
        lblOrder.Caption = "Order: "
    End If
    If Len(Order("BrokerOrderID")) > 0 Then
        lblBrokerID.Caption = "Broker ID: " & Order("BrokerOrderID")
    Else
        lblBrokerID.Caption = "Broker ID: <none>"
    End If
    
    Load
    
    If g.bAppIsIde Then
        mGenesis.ShowForm Me, eForm_Modal
    Else
        g.TnCore.ShowForm Me, eForm_Modal
    End If
    
    If m.bOK Then
        Save
        Set Journal = m.Journal
    ElseIf Len(Order("OptionNavImageFile")) > 0 Then
        If FileExist(Order("OptionNavImageFile")) Then
            KillFile Order("OptionNavImageFile")
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    g.TnCore.RaiseError "frmJournal.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeForOrderID
'' Description: Initialize and set controls and show the form
'' Inputs:      Order ID, Journal ID, Journal
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeForOrderID(ByVal lOrderID As Long, Optional ByVal lJournalID As Long = -1&, Optional Journal As cJournal = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim Order As cBrokerMessage         ' Order object for the given order ID
    
    Set Order = g.AppBridge.OrderForID(lOrderID)
    If Not Order Is Nothing Then
        ShowMeForOrderID = ShowMe(Order, lJournalID, Journal)
    End If

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    g.TnCore.RaiseError "frmJournal.ShowMeForOrderID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the journal
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection:

    PrintMe = frmPrintPreview.ShowMe("CNV Journal", frmJournal, , 0.75, 0.75, 0.75, 0.75)

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.PrintMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview control
'' Inputs:      Args to pass to print preview
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        g.TnCore.DoPrintHeader , frmPrintPreview.vp

        .FontName = "Times New Roman"
        .FontSize = 14
        .FontBold = True
        
        If Len(m.Order("BrokerOrderID")) > 0 Then
            .Text = vbLf & "Journal for Order #" & Str(m.Order("BrokerOrderID")) & ": " & m.Order("OrderText") & vbCrLf
        Else
            .Text = vbLf & "Journal for Order: " & m.Order("OrderText") & vbCrLf
        End If
        
        .FontSize = 12
        .FontBold = False
        
        .Paragraph = ""
        .Text = Replace(lblFeeling.Caption, "&", "") & vbLf
        .Text = cboFeelings.Text & vbCrLf
        
        .Text = Replace(lblReasons.Caption, "&", "") & vbLf
        .Text = cboReasons.Text & vbCrLf
        
        .Text = Replace(lblThoughts.Caption, "&", "") & vbLf
        .Text = cboThoughts.Text & vbCrLf
        
        .Text = Replace(lblAction.Caption, "&", "") & vbLf
        .Text = cboAction.Text & vbCrLf
        
        .Text = Replace(lblNotes.Caption, "&", "") & vbLf
        .Text = txtNotes.Text
        .Paragraph = ""

        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAction_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAction_GotFocus()
On Error GoTo ErrSection:

    SelectAll cboAction
    
    If m.lJournalID = -1& Then
        ShowDropDown cboAction
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.cboAction_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboCharts_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCharts_GotFocus()
On Error GoTo ErrSection:

    SelectAll cboCharts
    
    If m.lJournalID = -1& Then
        ShowDropDown cboCharts
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.cboCharts_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboEmotions_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboEmotions_GotFocus()
On Error GoTo ErrSection:

    SelectAll cboEmotions
    
    If m.lJournalID = -1& Then
        ShowDropDown cboEmotions
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.cboEmotions_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboFeelings_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboFeelings_GotFocus()
On Error GoTo ErrSection:

    SelectAll cboFeelings
    
    If m.lJournalID = -1& Then
        ShowDropDown cboFeelings
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.cboFeelings_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboReasons_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboReasons_GotFocus()
On Error GoTo ErrSection:

    SelectAll cboReasons
    
    If m.lJournalID = -1& Then
        ShowDropDown cboReasons
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.cboReasons_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboThoughts_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboThoughts_GotFocus()
On Error GoTo ErrSection:

    SelectAll cboThoughts
    
    If m.lJournalID = -1& Then
        ShowDropDown cboThoughts
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.cboThoughts_GotFocus"
    
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
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Save the information and unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: Allow the user to print the information
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
    g.TnCore.RaiseError "frmJournal.cmdPrint_Click"
    
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
        
        MoveFocus cboEmotions
    End If

ErrExit:
    Exit Sub

ErrSection:
    g.TnCore.RaiseError "frmJournal.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form when it gets loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strJournalIni As String         ' Ini file with journal values
    Dim strDefaultsIni As String        ' Ini file for journal defaults

    Icon = g.CoreBridge.Picture16(g.TnCore.ToolbarIcon("ID_TradeTracker"))
    
    g.Styler.StyleForm Me
    
    PlaceTheForm Me, g.strIniFile
    
    With cboEmotions
        .AddItem " "
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        .AddItem "10"
    End With
    
    With cboAction
        .AddItem "Entry"
        .AddItem "Exit"
        .AddItem "Reversal"
    End With
    
    strJournalIni = AddSlash(g.strAppPath) & "Journal.INI"
    strDefaultsIni = AddSlash(g.strAppPath) & "Provided\Journal.INI"
    
    Set m.FeelingsMru = New cMostRecentlyUsed
    m.FeelingsMru.Init "Feelings", 10, strJournalIni, strDefaultsIni
    m.FeelingsMru.Load
    LoadMruCombo m.FeelingsMru, cboFeelings
    
    Set m.ReasonsMru = New cMostRecentlyUsed
    m.ReasonsMru.Init "Reasons", 10, strJournalIni, strDefaultsIni
    m.ReasonsMru.Load
    LoadMruCombo m.ReasonsMru, cboReasons
    
    Set m.ThoughtsMru = New cMostRecentlyUsed
    m.ThoughtsMru.Init "Thoughts", 10, strJournalIni, strDefaultsIni
    m.ThoughtsMru.Load
    LoadMruCombo m.ThoughtsMru, cboThoughts
    
    txtOptNavImageFilename.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Allow ShowMe to close form if user clicks on the "X"
'' Inputs:      Whether to Cancel Unload, Mode of the Unload
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
    g.TnCore.RaiseError "frmJournal.Form_QueryUnload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form gets resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lNotesTop As Long               ' Top of the notes frame

    If Not LimitFormSize(Me, (fraButtons.Width * 2) + 240, (fraOrderInfo.Height * 4) + fraQuestions.Height + fraButtons.Height + 240) Then
        With fraOrderInfo
            .Move 120, 120, ScaleWidth - 240
        End With
        
        With lblOrder
            .Move 0, 0, fraOrderInfo.Width
        End With
        
        With fraQuestions
            .Move 120, fraOrderInfo.Height + 120, ScaleWidth - 240
        End With
        
        With cboEmotions
            .Move .Left, cboFeelings.Top
        End With
        
        With cboFeelings
            .Move .Left, .Top, fraQuestions.Width - .Left
        End With
        
        With cboReasons
            .Move 0, .Top, fraQuestions.Width
        End With
        
        With cboThoughts
            .Move 0, .Top, fraQuestions.Width
        End With
    
        With fraButtons
            .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height - 120
        End With
        
        If fraOptNavImage.Visible Then
            With fraOptNavImage
                .Move 120, fraButtons.Top - .Height - 120, ScaleWidth - 240
            End With
            
            With fraImage
                .Move 120, fraOptNavImage.Top - .Height - 60, ScaleWidth - 240
            End With
        Else
            With fraImage
                .Move 120, fraButtons.Top - .Height - 120, ScaleWidth - 240
            End With
        End If
        
        With cboCharts
            .Move .Left, .Top, fraImage.Width - .Left
        End With
        
        With txtOptNavImageCaption
            .Move .Left, .Top, fraOptNavImage.Width - .Left
        End With
        
        With fraNotes
            lNotesTop = fraOrderInfo.Height + fraQuestions.Height + 120
            .Move 120, lNotesTop, ScaleWidth - 240, fraImage.Top - lNotesTop
        End With
        
        With txtNotes
            .Move 0, .Top, fraNotes.Width, fraNotes.Height - .Top - 120
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

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the information from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Load()
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return from journal load
    Dim JournalImage As cJournalImage   ' Journal image
    Dim strFileName As String           ' Filename for the Option Nav image

    bReturn = True
    If m.Journal Is Nothing Then
        Set m.Journal = New cJournal
        bReturn = g.JournalDB.LoadOrderJournal(m.lJournalID, m.Journal)
    End If
    
    If bReturn = True Then
        If m.Journal.EmotionNumber = -1 Then
            cboEmotions.Text = " "
        Else
            cboEmotions.Text = Str(m.Journal.EmotionNumber)
        End If
        SetMruCombo m.Journal.Feelings, cboFeelings
        SetMruCombo m.Journal.WhyTrade, cboReasons
        SetMruCombo m.Journal.Thoughts, cboThoughts
        If Len(m.Journal.Action) > 0 Then
            cboAction.Text = m.Journal.Action
        Else
            cboAction.ListIndex = -1&
        End If
        gdCurrentTime.Value = m.Journal.NoteDate
        txtNotes.Text = m.Journal.Note
        
        Set JournalImage = m.Journal.JournalImage(eGDJournalImageType_OptionNavOrder)
        If JournalImage Is Nothing Then
            CheckBoxValue(chkOptNavImage) = False
            fraOptNavImage.Visible = False
        Else
            txtOptNavImageCaption.Text = JournalImage.Caption
            txtOptNavImageFilename.Text = JournalImage.FileName
            
            CheckBoxValue(chkOptNavImage) = True
            fraOptNavImage.Visible = True
        End If
    Else
        cboEmotions.ListIndex = 6
        SetMruCombo "", cboFeelings
        SetMruCombo "", cboReasons
        SetMruCombo "", cboThoughts
        cboAction.ListIndex = 0
        gdCurrentTime.Value = g.TnCore.CurrentTime
        txtNotes.Text = ""
        
        If (Len(m.Order("OptionNavImageFile")) = 0) Or (Not FileExist(m.Order("OptionNavImageFile"))) Then
            CheckBoxValue(chkOptNavImage) = False
            fraOptNavImage.Visible = False
        Else
            If DirExist(AddSlash(g.strAppPath) & "SavedImages") = False Then
                MkDir AddSlash(g.strAppPath) & "SavedImages"
            End If
            
            strFileName = AddSlash(g.strAppPath) & "SavedImages\" & Format(g.TnCore.CurrentTime, "YYYYMMDD HHMMSS") & ".JPG"
            FileCopy m.Order("OptionNavImageFile"), strFileName, True
            KillFile m.Order("OptionNavImageFile")
            
            txtOptNavImageCaption.Text = m.Order("GroupName") & " " & Format(g.TnCore.CurrentTime, "YYYY-MM-DD HH:MM:SS")
            txtOptNavImageFilename.Text = strFileName
            
            CheckBoxValue(chkOptNavImage) = True
            fraOptNavImage.Visible = True
        End If
    End If
    
    LoadChartsCombo m.Journal.JournalImage(eGDJournalImageType_Chart)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the information back to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save()
On Error GoTo ErrSection:

    Dim lOrderID As Long                ' Order ID
    Dim lEmotionNumber As Long          ' Emotion number out of the combo
    Dim JournalImage As cJournalImage   ' Journal image information
    
    If m.lNewOrderID = 0 Then
        lOrderID = CLng(Val(m.Order("OrderID")))
    Else
        lOrderID = m.lNewOrderID
    End If
    
    If Len(Trim(cboEmotions.Text)) = 0 Then
        lEmotionNumber = -1&
    Else
        lEmotionNumber = CLng(Val(cboEmotions.Text))
    End If
    
    If m.Journal.NoteDate = 0# Then
        m.Journal.NoteDate = gdCurrentTime.Value
    End If
    
    If m.lNewOrderID = 0& Then
        m.Journal.OrderID = CLng(Val(m.Order("OrderID")))
    Else
        m.Journal.OrderID = m.lNewOrderID
    End If
    m.Journal.AccountID = CLng(Val(m.Order("AccountID")))
    m.Journal.EmotionNumber = lEmotionNumber
    
    If Len(cboFeelings.Text) > 0 Then
        If cboFeelings.Text <> m.Journal.Feelings Then
            If m.FeelingsMru.LastUsed <> cboFeelings.Text Then
                m.FeelingsMru.LastUsed = cboFeelings.Text
                m.FeelingsMru.Save
            End If
        End If
    End If
    m.Journal.Feelings = cboFeelings.Text
    
    If Len(cboReasons.Text) > 0 Then
        If cboReasons.Text <> m.Journal.WhyTrade Then
            If m.ReasonsMru.LastUsed <> cboReasons.Text Then
                m.ReasonsMru.LastUsed = cboReasons.Text
                m.ReasonsMru.Save
            End If
        End If
    End If
    m.Journal.WhyTrade = cboReasons.Text
    
    If Len(cboThoughts.Text) > 0 Then
        If cboThoughts.Text <> m.Journal.Thoughts Then
            If m.ThoughtsMru.LastUsed <> cboThoughts.Text Then
                m.ThoughtsMru.LastUsed = cboThoughts.Text
                m.ThoughtsMru.Save
            End If
        End If
    End If
    m.Journal.Thoughts = cboThoughts.Text
    
    m.Journal.Action = cboAction.Text
    m.Journal.Note = txtNotes.Text
    m.Journal.JournalDate = CDbl(Int(m.Journal.NoteDate))
    m.Journal.SymbolOrSymbolID = m.Order("Symbol")
    
    If CheckBoxValue(chkAttachImage) = False Then
        Set JournalImage = m.Journal.JournalImage(eGDJournalImageType_Chart)
        If Not JournalImage Is Nothing Then
            If Len(JournalImage.FileName) > 0 Then
                KillFile JournalImage.FileName
            End If
            
            m.Journal.JournalImages.Remove Str(eGDJournalImageType_Chart)
        End If
    Else
        Set JournalImage = m.Journal.JournalImage(eGDJournalImageType_Chart)
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
        m.Journal.JournalImage(eGDJournalImageType_Chart) = JournalImage
    End If
    
    If CheckBoxValue(chkOptNavImage) = False Then
        Set JournalImage = m.Journal.JournalImage(eGDJournalImageType_OptionNavOrder)
        If Not JournalImage Is Nothing Then
            If Len(JournalImage.FileName) > 0 Then
                KillFile JournalImage.FileName
            End If
            
            m.Journal.JournalImages.Remove Str(eGDJournalImageType_OptionNavOrder)
        End If
    Else
        Set JournalImage = m.Journal.JournalImage(eGDJournalImageType_OptionNavOrder)
        If JournalImage Is Nothing Then
            Set JournalImage = New cJournalImage
        End If
        
        JournalImage.ImageType = eGDJournalImageType_OptionNavOrder
        JournalImage.FileName = txtOptNavImageFilename.Text
        JournalImage.Caption = txtOptNavImageCaption.Text
        
        m.Journal.JournalImage(eGDJournalImageType_OptionNavOrder) = JournalImage
    End If
    
    g.JournalDB.SaveOrderJournal m.Journal
    
    ' If the Trade Tracker form is up and for the same account as the order
    ' for this journal, send the new updated journal entry to that form...
    g.AppBridge.UpdateJournal m.Journal.JournalID
    
    ' If the date journals form is currently loaded, send this journal there
    ' as well...
    If FormIsLoaded("frmDateJournals") Then
        frmDateJournals.UpdateOrderJournal m.Journal
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.Save"
    
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
    g.TnCore.RaiseError "frmJournal.LoadChartsCombo"
    
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
    g.TnCore.RaiseError "frmJournal.ExportChart"
    
End Function

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
    g.TnCore.RaiseError "frmJournal.LoadMruCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetMruCombo
'' Description: Set the given MRU combo to the given value
'' Inputs:      Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetMruCombo(ByVal strValue As String, cboMru As ComboBox)
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
    ElseIf cboMru.ListCount > 0 Then
        cboMru.ListIndex = 0
    Else
        cboMru.Text = ""
    End If
    
    cboMru.ToolTipText = cboMru.Text

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournal.SetMruCombo"
    
End Sub

