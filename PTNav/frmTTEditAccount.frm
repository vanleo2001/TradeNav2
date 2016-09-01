VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTTEditAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraSecTypes 
      Height          =   795
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   3615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmTTEditAccount.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditAccount.frx":004C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditAccount.frx":006C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkStkOpts 
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   480
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
         Caption         =   "frmTTEditAccount.frx":0088
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditAccount.frx":00C0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":00E0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkFutOpts 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
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
         Caption         =   "frmTTEditAccount.frx":00FC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditAccount.frx":0134
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0154
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkForex 
         Height          =   255
         Left            =   2520
         TabIndex        =   18
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
         Caption         =   "frmTTEditAccount.frx":0170
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditAccount.frx":019C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":01BC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkStocks 
         Height          =   255
         Left            =   1320
         TabIndex        =   17
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
         Caption         =   "frmTTEditAccount.frx":01D8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditAccount.frx":0206
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0226
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkFutures 
         Height          =   255
         Left            =   120
         TabIndex        =   16
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
         Caption         =   "frmTTEditAccount.frx":0242
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditAccount.frx":0272
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0292
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraAccountInfo 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmTTEditAccount.frx":02AE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditAccount.frx":02DA
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditAccount.frx":02FA
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkFillRT 
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   2580
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
         Caption         =   "frmTTEditAccount.frx":0316
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditAccount.frx":0376
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0396
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtCommissions 
         Height          =   285
         Left            =   1380
         TabIndex        =   12
         Top             =   1805
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTTEditAccount.frx":03B2
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
         Tip             =   "frmTTEditAccount.frx":03D2
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":03F2
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBroker 
         Height          =   285
         Left            =   1380
         TabIndex        =   10
         Top             =   1450
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTTEditAccount.frx":040E
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
         Tip             =   "frmTTEditAccount.frx":042E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":044E
      End
      Begin HexUniControls.ctlUniComboImageXP cboAccountType 
         Height          =   315
         Left            =   1380
         TabIndex        =   14
         Top             =   2160
         Width           =   2235
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
         Tip             =   "frmTTEditAccount.frx":046A
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":048A
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtAccountNumber 
         Height          =   285
         Left            =   1380
         TabIndex        =   2
         Top             =   0
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTTEditAccount.frx":04A6
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
         Tip             =   "frmTTEditAccount.frx":04C6
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":04E6
      End
      Begin HexUniControls.ctlUniTextBoxXP txtName 
         Height          =   285
         Left            =   1380
         TabIndex        =   4
         Top             =   355
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTTEditAccount.frx":0502
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
         Tip             =   "frmTTEditAccount.frx":0522
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0542
      End
      Begin HexUniControls.ctlUniTextBoxXP txtStartBalance 
         Height          =   285
         Left            =   1380
         TabIndex        =   6
         Top             =   710
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTTEditAccount.frx":055E
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
         Tip             =   "frmTTEditAccount.frx":057E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":059E
      End
      Begin gdOCX.gdSelectDate gdStartDate 
         Height          =   315
         Left            =   1380
         TabIndex        =   8
         Top             =   1065
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
      End
      Begin HexUniControls.ctlUniLabelXP lblComms 
         Height          =   255
         Left            =   0
         Top             =   1820
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
         Caption         =   "frmTTEditAccount.frx":05BA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditAccount.frx":05F2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0612
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBroker 
         Height          =   255
         Left            =   0
         Top             =   1465
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
         Caption         =   "frmTTEditAccount.frx":062E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditAccount.frx":065E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":067E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAccountType 
         Height          =   255
         Left            =   0
         Top             =   2190
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
         Caption         =   "frmTTEditAccount.frx":069A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditAccount.frx":06D6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":06F6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblNumber 
         Height          =   255
         Left            =   0
         Top             =   15
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
         Caption         =   "frmTTEditAccount.frx":0712
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditAccount.frx":0752
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0772
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblName 
         Height          =   255
         Left            =   0
         Top             =   370
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
         Caption         =   "frmTTEditAccount.frx":078E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditAccount.frx":07C2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":07E2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStartBalance 
         Height          =   255
         Left            =   0
         Top             =   725
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
         Caption         =   "frmTTEditAccount.frx":07FE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditAccount.frx":0842
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0862
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStartDate 
         Height          =   255
         Left            =   0
         Top             =   1095
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
         Caption         =   "frmTTEditAccount.frx":087E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditAccount.frx":08BC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":08DC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1515
      Left            =   3900
      TabIndex        =   7
      Top             =   120
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
      Caption         =   "frmTTEditAccount.frx":08F8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditAccount.frx":0924
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditAccount.frx":0944
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdConfig 
         Height          =   435
         Left            =   0
         TabIndex        =   9
         Top             =   1080
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
         Caption         =   "frmTTEditAccount.frx":0960
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditAccount.frx":0996
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":09B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   435
         Left            =   0
         TabIndex        =   11
         Top             =   480
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
         Caption         =   "frmTTEditAccount.frx":09D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditAccount.frx":0A00
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0A20
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   435
         Left            =   0
         TabIndex        =   13
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
         Caption         =   "frmTTEditAccount.frx":0A3C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditAccount.frx":0A62
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditAccount.frx":0A82
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmTTEditAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    frmTTEditAccount.frm
'' Description: Allow the user to edit account information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 07/25/2002   DAJ         Created
'' 06/01/2009   DAJ         Added FO and SO to security type mask for trading
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 06/24/2011   DAJ         Utilize NextAccount functions from simulated objects
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean
    Account As cPtAccount
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Allow an outside caller to show the form
'' Inputs:      Account ID to load (0 for New Account)
'' Returns:     True if user clicked OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal lAccountID As Long, Optional ByVal nAccountType As eTT_AccountType = eTT_AccountType_SimStream) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim dOldStarting As Double          ' Old Starting Balance
    Dim lSecTypeMask As Long            ' Secutity type mask

    Set m.Account = New cPtAccount
    If m.Account.Load(lAccountID) Then
        With m.Account
            txtAccountNumber.Text = .AccountNumber
            txtName.Text = .Name
            txtStartBalance.Text = Format(.StartingBalance, "$#,##0.00")
            gdStartDate.Value = .StartingDate
            txtBroker.Text = .Broker
            txtCommissions.Text = Format(.Comms, "$#,##0.00")
            SetAccountTypeCombo .AccountType
            
            SetCheckBox chkFutures, GetBit(.SecTypeMask, 1)
            SetCheckBox chkStocks, GetBit(.SecTypeMask, 2)
            SetCheckBox chkForex, GetBit(.SecTypeMask, 3)
            SetCheckBox chkFutOpts, GetBit(.SecTypeMask, 4)
            SetCheckBox chkStkOpts, GetBit(.SecTypeMask, 5)
            
            If .FillRT = True Then chkFillRT.Value = vbChecked Else chkFillRT.Value = vbUnchecked
        End With
    
        lblAccountType.Enabled = False
        cboAccountType.Enabled = False
        lblNumber.Enabled = False
        txtAccountNumber.Enabled = False
    Else
        txtStartBalance.Text = "$25,000"
        gdStartDate.Value = Now
        SetAccountTypeCombo nAccountType
        
        Caption = g.Broker.BrokerName(nAccountType) & " Account Setup"
        If Not g.Broker.IsLiveAccount(nAccountType) Then
            chkFutures.Value = vbChecked
            chkStocks.Value = vbChecked
            chkForex.Value = vbChecked
            lblAccountType.Enabled = True
            cboAccountType.Enabled = True
        Else
            chkFutures.Value = vbChecked
            chkStocks.Value = vbUnchecked
            chkForex.Value = vbUnchecked
            lblAccountType.Enabled = False
            cboAccountType.Enabled = False
        End If
        
        lblNumber.Enabled = True
        txtAccountNumber.Enabled = True
    End If
    
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK Then
        With m.Account
            dOldStarting = .StartingBalance
            
            .AccountNumber = Trim(txtAccountNumber.Text)
            .Name = Trim(txtName.Text)
            .StartingBalance = ValOfText(txtStartBalance.Text)
            .StartingDate = gdStartDate.Value
            .Broker = Trim(txtBroker.Text)
            .Comms = ValOfText(txtCommissions.Text)
            .AccountType = cboAccountType.ItemData(cboAccountType.ListIndex)
            
            If lAccountID = 0& Then
                .CurrentBalance = .StartingBalance
            Else
                .CurrentBalance = .CurrentBalance + (.StartingBalance - dOldStarting)
            End If
            
            SetBit lSecTypeMask, 1, (chkFutures.Value = vbChecked)
            SetBit lSecTypeMask, 2, (chkStocks.Value = vbChecked)
            SetBit lSecTypeMask, 3, (chkForex.Value = vbChecked)
            SetBit lSecTypeMask, 4, (chkFutOpts.Value = vbChecked)
            SetBit lSecTypeMask, 5, (chkStkOpts.Value = vbChecked)
            .SecTypeMask = lSecTypeMask
            
            .FillRT = (chkFillRT.Value = vbChecked)
            
            .Save
            
            g.Broker.UpdateAccount m.Account
        End With
    End If
    
ErrExit:
    ShowMe = m.bOK
    Unload Me
    Exit Function

ErrSection:
    ShowMe = False
    Unload Me
    RaiseError "frmTTEditAccount.ShowMe"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccountType_Click
'' Description: Allow the user to change account types
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccountType_Click()
On Error GoTo ErrSection:

    cmdConfig.Visible = (UCase(cboAccountType.Text) <> "SIMTRADE")
        
    If cboAccountType.ItemData(cboAccountType.ListIndex) = eTT_AccountType_SimStream Then
        chkFillRT.Value = vbChecked
        If (m.Account.AccountID = 0&) Then
            txtAccountNumber.Text = g.SimTradeStream.NextAccount
        End If
    ElseIf cboAccountType.ItemData(cboAccountType.ListIndex) = eTT_AccountType_SimBroker Then
        chkFillRT.Value = vbChecked
        If (m.Account.AccountID = 0&) Then
            txtAccountNumber.Text = g.SimTradeTs.NextAccount
        End If
    Else
        chkFillRT.Value = vbUnchecked
        If (m.Account.AccountID = 0&) Then
            If (Len(Trim(txtAccountNumber.Text)) = 7) And ((Left(Trim(UCase(txtAccountNumber.Text)), 3) = "SIM") Or (Left(Trim(UCase(txtAccountNumber.Text)), 3) = "GEN")) Then
                txtAccountNumber.Text = ""
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.cboAccountType_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks Cancel, unload the form without saving
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
    RaiseError "frmTTEditAccount.cmdCancel_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdConfig_Click
'' Description: Allow the user to configure their online broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdConfig_Click()
On Error GoTo ErrSection:

    g.Broker.ShowBrokerConnectionInfo cboAccountType.ItemData(cboAccountType.ListIndex), False, m.Account.UserName, False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.cmdConfig_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks OK, save the information and unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    If ValOfText(txtStartBalance) > (2 ^ 31) Then
        MoveFocus txtStartBalance
        Err.Raise vbObjectError + 1000, , "Invalid Starting Balance"
    End If
    
    If Len(Trim(txtAccountNumber.Text)) = 0 Or Len(Trim(txtAccountNumber.Text)) > 20 Then
        MoveFocus txtAccountNumber
        Err.Raise vbObjectError + 1000, , "Account Number must be Between 1 and 20 Characters"
    End If

    If Len(Trim(txtName.Text)) = 0 Then txtName.Text = txtAccountNumber.Text
    
    If Len(Trim(txtName.Text)) > 50 Then
        MoveFocus txtName
        Err.Raise vbObjectError + 1000, , "Name must be Between 1 and 50 Characters"
    End If
    
    If InStr(txtName.Text, "'") > 0 Then
        MoveFocus txtName
        Err.Raise vbObjectError + 1000, , "Account Name cannot contain an apostrophe"
    End If
    
    If InStr(txtAccountNumber.Text, "'") > 0 Then
        MoveFocus txtAccountNumber
        Err.Raise vbObjectError + 1000, , "Account Number cannot contain an apostrophe"
    End If
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
            "WHERE [AccountNumber]='" & Trim(txtAccountNumber.Text) & "';", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        If rs!AccountID <> m.Account.AccountID Then
            MoveFocus txtAccountNumber
            Err.Raise vbObjectError + 1000, , "Account Number must be unique"
        End If
    End If

    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
            "WHERE [Name]='" & Trim(txtName.Text) & "';", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        If rs!AccountID <> m.Account.AccountID Then
            MoveFocus txtName
            Err.Raise vbObjectError + 1000, , "Account Name must be unique"
        End If
    End If
    
    m.bOK = True
    Me.Hide

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.cmdOK_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: If the user hits F1, show the help on this form
'' Inputs:      Key Pressed, Shift/Ctrl/Alt status
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
    RaiseError "frmTTEditAccount.Form_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, do some intialization and center it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim bShowAccountCombo As Boolean    ' Should we show the account combo?

    Me.Caption = "Edit Account Information"
    Me.Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    bShowAccountCombo = g.Broker.LoadBrokerCombo(cboAccountType)
    
    ' Only show the account type combo box if the PATS directory exists...
    lblAccountType.Visible = bShowAccountCombo
    cboAccountType.Visible = bShowAccountCombo
    
    ' Hide these until we get the automatic fee generation hooked up...
    lblComms.Visible = False
    txtCommissions.Visible = False
    
    ' Hide this control for now...
    chkFillRT.Visible = IsIDE

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the X, unload the form without saving
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAccountNumber_GotFocus
'' Description: When the text box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAccountNumber_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAccountNumber

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.txtAccountNumber_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAccountNumber_LostFocus
'' Description: When the control loses the focus, trim the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAccountNumber_LostFocus()
On Error GoTo ErrSection:

    txtAccountNumber.Text = Trim(txtAccountNumber.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.txtAccountNumber_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtBroker_GotFocus
'' Description: When the text box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtBroker_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtBroker

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.txtBroker_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtCommissions_GotFocus
'' Description: When the text box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtCommissions_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtCommissions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.txtCommissions_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtCommissions_LostFocus
'' Description: When the text box loses the focus, reformat the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtCommissions_LostFocus()
On Error GoTo ErrSection:

    txtCommissions.Text = Format(ValOfText(txtCommissions.Text), "$#,##0.00")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.txtCommissions_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_GotFocus
'' Description: When the text box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.txtName_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_LostFocus
'' Description: When the control loses the focus, trim the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_LostFocus()
On Error GoTo ErrSection:

    txtName.Text = Trim(txtName.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.txtName_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStartBalance_GotFocus
'' Description: When the text box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStartBalance_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtStartBalance

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.txtStartBalance_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStartBalance_LostFocus
'' Description: When the text box loses the focus, reformat the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStartBalance_LostFocus()
On Error GoTo ErrSection:

    txtStartBalance.Text = Format(ValOfText(txtStartBalance.Text), "$#,##0.00")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.txtStartBalance_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCheckBox
'' Description: Set the check box checked if value is true, uncheked otherwise
'' Inputs:      Check Box, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCheckBox(chk As ctlUniCheckXP, ByVal bValue As Boolean)  'RH was Checkbox
On Error GoTo ErrSection:

    If bValue = True Then
        chk.Value = vbChecked
    Else
        chk.Value = vbUnchecked
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.SetCheckBox"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAccountTypeCombo
'' Description: Set the account type combo box appropriately
'' Inputs:      Account Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAccountTypeCombo(ByVal nAccountType As eTT_AccountType)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
        
    For lIndex = 0 To cboAccountType.ListCount - 1
        If cboAccountType.ItemData(lIndex) = nAccountType Then
            cboAccountType.ListIndex = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditAccount.SetAccountTypeCombo"
    
End Sub

