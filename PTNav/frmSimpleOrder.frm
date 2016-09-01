VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSimpleOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdLookupAccount 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   120
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
      Caption         =   "frmSimpleOrder.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmSimpleOrder.frx":0040
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":0060
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboAccounts 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   120
      Width           =   3015
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
      Tip             =   "frmSimpleOrder.frx":007C
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":009C
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboLots 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   4320
      Width           =   3255
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
      Tip             =   "frmSimpleOrder.frx":00B8
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":00D8
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtQuantity 
      Height          =   315
      Left            =   1500
      TabIndex        =   6
      Top             =   1080
      Width           =   780
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmSimpleOrder.frx":00F4
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
      Tip             =   "frmSimpleOrder.frx":011E
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":013E
   End
   Begin HexUniControls.ctlUniTextBoxXP txtWithLimitPrice 
      Height          =   285
      Left            =   1500
      TabIndex        =   17
      Top             =   2925
      Width           =   1020
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmSimpleOrder.frx":015A
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
      Tip             =   "frmSimpleOrder.frx":018A
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":01AA
   End
   Begin HexUniControls.ctlUniTextBoxXP txtPrice 
      Height          =   285
      Left            =   1500
      TabIndex        =   14
      Top             =   2505
      Width           =   1020
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmSimpleOrder.frx":01C6
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
      Tip             =   "frmSimpleOrder.frx":01F6
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":0216
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   4800
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
      Caption         =   "frmSimpleOrder.frx":0232
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSimpleOrder.frx":025E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":027E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   1320
         TabIndex        =   5
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
         Caption         =   "frmSimpleOrder.frx":029A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSimpleOrder.frx":02C8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSimpleOrder.frx":02E8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSubmit 
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   8
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
         Caption         =   "frmSimpleOrder.frx":0304
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSimpleOrder.frx":0332
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSimpleOrder.frx":0352
         RightToLeft     =   0   'False
      End
   End
   Begin gdOCX.gdSelectDate gdExpirationDate 
      Height          =   315
      Left            =   1500
      TabIndex        =   11
      Top             =   3840
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
   End
   Begin HexUniControls.ctlUniComboImageXP cboTifs 
      Height          =   315
      Left            =   1500
      TabIndex        =   13
      Top             =   3360
      Width           =   795
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
      Tip             =   "frmSimpleOrder.frx":036E
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":038E
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboOrderTypes 
      Height          =   315
      Left            =   1500
      TabIndex        =   12
      Top             =   2040
      Width           =   1440
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
      Tip             =   "frmSimpleOrder.frx":03AA
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":03CA
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboSides 
      Height          =   315
      Left            =   1500
      TabIndex        =   4
      Top             =   600
      Width           =   1440
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
      Tip             =   "frmSimpleOrder.frx":03E6
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":0406
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
      Height          =   255
      Left            =   3390
      TabIndex        =   10
      Top             =   1590
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
      Caption         =   "frmSimpleOrder.frx":0422
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmSimpleOrder.frx":0454
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":0474
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   1560
      Width           =   2160
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmSimpleOrder.frx":0490
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
      Tip             =   "frmSimpleOrder.frx":04D0
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":04F0
   End
   Begin gdOCX.gdScrollBar sbPrice 
      Height          =   360
      Left            =   2520
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2460
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   635
   End
   Begin gdOCX.gdScrollBar sbWithLimitPrice 
      Height          =   360
      Left            =   2520
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2880
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   635
   End
   Begin gdOCX.gdScrollBar sbQuantity 
      Height          =   360
      Left            =   2280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1050
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   635
   End
   Begin HexUniControls.ctlUniLabelXP lblAccount 
      Height          =   195
      Left            =   180
      Top             =   180
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
      Caption         =   "frmSimpleOrder.frx":050C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":053E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":055E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblLot 
      Height          =   195
      Left            =   180
      Top             =   4380
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
      Caption         =   "frmSimpleOrder.frx":057A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":05A4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":05C4
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblExpirationDate 
      Height          =   195
      Left            =   180
      Top             =   3900
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
      Caption         =   "frmSimpleOrder.frx":05E0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":0622
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":0642
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblTimeInForce 
      Height          =   195
      Left            =   180
      Top             =   3420
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
      Caption         =   "frmSimpleOrder.frx":065E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":069C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":06BC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblWithLimitPrice 
      Height          =   195
      Left            =   180
      Top             =   2970
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
      Caption         =   "frmSimpleOrder.frx":06D8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":0710
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":0730
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblPrice 
      Height          =   195
      Left            =   180
      Top             =   2550
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
      Caption         =   "frmSimpleOrder.frx":074C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":0778
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":0798
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblOrderType 
      Height          =   195
      Left            =   180
      Top             =   2100
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
      Caption         =   "frmSimpleOrder.frx":07B4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":07EC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":080C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblSymbol 
      Height          =   195
      Left            =   180
      Top             =   1620
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
      Caption         =   "frmSimpleOrder.frx":0828
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":0858
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":0878
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblQuantity 
      Height          =   195
      Left            =   180
      Top             =   1140
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
      Caption         =   "frmSimpleOrder.frx":0894
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":08C8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":08E8
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblSide 
      Height          =   195
      Left            =   180
      Top             =   660
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
      Caption         =   "frmSimpleOrder.frx":0904
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSimpleOrder.frx":0930
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSimpleOrder.frx":0950
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmSimpleOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSimpleOrder.frm
'' Description: User interface for a simple order form
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 04/12/2012   DAJ         Added Genesis Symbols
'' 04/20/2012   DAJ         Mods for handling new order with information
'' 05/31/2012   DAJ         Turnkey implementation
'' 06/11/2012   DAJ         Add Account to simple order dialog
'' 09/11/2012   DAJ         Owner name in lot combo, Lot combo keyed by Lot ID
'' 09/27/2012   DAJ         Changed account to a combo box
'' 01/09/2013   DAJ         Made symbol text box wider to accomodate option symbols
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 03/07/2014   DAJ         Moved Cattle stuff into NavCattle.DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click OK?
    bNewOrder As Boolean                ' Is this a new order?
    bGenesisSymbol As Boolean           ' Is this dialog using a Genesis symbol?
    bLookupSymbolDone As Boolean        ' Was the symbol lookup done on activate?
    
    Bars As cGdBars                     ' Bars object
    Price As cPriceEditor               ' Price editor for the Price
    WithLimitPrice As cPriceEditor      ' Price editor for the With-Limit price
    Quantity As cPriceEditor            ' Price editor for the Quantity
End Type
Private m As mPrivate

Private Property Get AccountID() As Long
    If cboAccounts.ListIndex >= 0 Then
        AccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
    Else
        AccountID = -1&
    End If
End Property
Private Property Let AccountID(ByVal lAccountID As Long)
    SelectComboByItemData cboAccounts, lAccountID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Broker Message with the order, Account, New Price, New Quantity,
''              Feed Yard Lot ID
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(brokerMessage As cBrokerMessage, ByVal lAccountID As Long, Optional ByVal dNewPrice As Double = 0#, Optional ByVal lNewQuantity As Long = 0&, Optional strFeedYardLotID As String = "") As Boolean
On Error GoTo ErrSection:

    LoadCombo cboSides, "Buy,Sell"
    LoadCombo cboOrderTypes, "Market,Stop,Limit,StopWithLimit"
    LoadCombo cboTifs, "Day,GTC,IOC,FOK,GTD"
    CopyComboBox frmBrokerView.cboAccounts, cboAccounts
    
    If Len(brokerMessage("BrokerID")) = 0 Then
        m.bNewOrder = True
        Caption = "New Order"
    Else
        m.bNewOrder = False
        Caption = "Order #" & brokerMessage("BrokerID")
    End If
    If (FormIsLoaded("frmTurnkey") = True) And (m.bNewOrder = True) Then
        lblLot.Visible = True
        cboLots.Visible = True
        fraButtons.Top = 4785 ' 4320
        Height = 5910 ' 5445
        
        g.CattleBridge.LoadLotsCombo cboLots, , "None"
        If Len(strFeedYardLotID) > 0 Then
            SelectComboByItemData cboLots, CLng(Val(strFeedYardLotID))
        End If
    Else
        lblLot.Visible = False
        cboLots.Visible = False
        fraButtons.Top = lblLot.Top
        Height = 5445 ' 4980
    End If
    
    ControlsFromBrokerMessage brokerMessage, lAccountID, dNewPrice, lNewQuantity
    EnableControls
    
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK Then
        ControlsToBrokerMessage brokerMessage
        
        strFeedYardLotID = ""
        If cboLots.ListIndex >= 0 Then
            If cboLots.Text <> "None" Then
                strFeedYardLotID = cboLots.ItemData(cboLots.ListIndex)
            End If
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmSimpleOrder.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccounts_Click
'' Description: Handle the user changing the account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    If (Visible = True) And (FormIsLoaded("frmBrokerView") = True) Then
        frmBrokerView.ChangeAccount AccountID, "simple order"
        InitQuantityEditor
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.cboAccounts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboOrderTypes_Click
'' Description: Handle the user changing the order type
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboOrderTypes_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.cboOrderTypes_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboTifs_Click
'' Description: Handle the user changing the time-in-force
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboTifs_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.cboTifs_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Don't submit the order
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
    RaiseError "frmSimpleOrder.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: Allow the user to lookup a symbol
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
    RaiseError "frmSimpleOrder.cmdLookup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookupAccount_Click
'' Description: Allow the user to lookup an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookupAccount_Click()
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID selected
    
    lAccountID = frmBrokerView.LookupAccount
    If lAccountID > -1& Then
        AccountID = lAccountID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.cmdLookupAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSubmit_Click
'' Description: Submit the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSubmit_Click()
On Error GoTo ErrSection:

    If VerifyControls Then
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.cmdSubmit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Perform actions when form is activated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    If m.bLookupSymbolDone = False Then
        m.bLookupSymbolDone = True
        
        If (m.bNewOrder = True) And (Len(txtSymbol.Text) = 0) Then
            LookupSymbol
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement of the form
    
    strPlacement = GetIniFileProperty("frmSimpleOrder", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me
    
    Set m.Bars = New cGdBars
    Set m.Price = New cPriceEditor
    Set m.WithLimitPrice = New cPriceEditor
    Set m.Quantity = New cPriceEditor
    
    m.bLookupSymbolDone = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, let the ShowMe unload the form
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    m.bLookupSymbolDone = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_Click
'' Description: Allow the user to lookup a symbol
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
    RaiseError "frmSimpleOrder.txtSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_GotFocus
'' Description: Select all text when the text box gets the focus
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
    RaiseError "frmSimpleOrder.txtSymbol_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_KeyPress
'' Description: Allow the user to lookup a symbol
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
    RaiseError "frmSimpleOrder.txtSymbol_KeyPress", 0
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load the given combo with the valid values
'' Inputs:      Combo, Valid Values
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo(cboCombo As ctlUniComboImageXP, ByVal strValidValues As String)
On Error GoTo ErrSection:

    Dim astrValues As cGdArray          ' Valid values for the combo
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrValues = New cGdArray
    astrValues.SplitFields strValidValues, ","
    
    cboCombo.Clear
    For lIndex = 0 To astrValues.Size - 1
        cboCombo.AddItem astrValues(lIndex)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.LoadSidesCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Enable cboAccounts, m.bNewOrder
    Enable cmdLookupAccount, m.bNewOrder
    Enable lblSide, m.bNewOrder
    Enable cboSides, m.bNewOrder
    Enable lblSymbol, m.bNewOrder
    Enable txtSymbol, m.bNewOrder
    Enable cmdLookup, m.bNewOrder
    Enable lblPrice, UCase(cboOrderTypes.Text) <> "MARKET"
    Enable txtPrice, lblPrice.Enabled
    Enable sbPrice, lblPrice.Enabled
    Enable lblWithLimitPrice, UCase(cboOrderTypes.Text) = "STOPWITHLIMIT"
    Enable txtWithLimitPrice, lblWithLimitPrice.Enabled
    Enable sbWithLimitPrice, lblWithLimitPrice.Enabled
    Enable lblExpirationDate, UCase(cboTifs.Text) = "GTD"
    Enable gdExpirationDate, lblExpirationDate.Enabled

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ControlsFromBrokerMessage
'' Description: Fill the controls from the given broker message
'' Inputs:      Broker Message, Account, New Price, New Quantity
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ControlsFromBrokerMessage(ByVal brokerMessage As cBrokerMessage, ByVal lAccountID As Long, Optional ByVal dNewPrice As Double = 0#, Optional ByVal lNewQuantity As Long = 0&)
On Error GoTo ErrSection:

    Dim dOrderPrice As Double           ' Price of the order

    AccountID = lAccountID

    SelectComboByText cboSides, brokerMessage("Side")
    SelectComboByText cboOrderTypes, brokerMessage("Type")
    
    If (Len(brokerMessage("GenesisSymbol")) > 0) Or (m.bNewOrder = True) Then
        m.bGenesisSymbol = True
        cmdLookup.Visible = True
        If Len(brokerMessage("GenesisSymbol")) > 0 Then
            ChangeSymbol brokerMessage("GenesisSymbol")
        
            Select Case UCase(brokerMessage("Type"))
                Case "STOP"
                    If dNewPrice > 0# Then
                        dOrderPrice = dNewPrice
                    Else
                        dOrderPrice = Val(brokerMessage("GenesisStopPrice"))
                    End If
                    
                    m.Price.Price = dOrderPrice
                    m.WithLimitPrice.Price = 0
                
                Case "LIMIT"
                    If dNewPrice > 0# Then
                        dOrderPrice = dNewPrice
                    Else
                        dOrderPrice = Val(brokerMessage("GenesisLimitPrice"))
                    End If
                    
                    m.Price.Price = dOrderPrice
                    m.WithLimitPrice.Price = 0
                
                Case "STOPWITHLIMIT"
                    If dNewPrice > 0# Then
                        dOrderPrice = dNewPrice
                    Else
                        dOrderPrice = Val(brokerMessage("GenesisStopPrice"))
                    End If
                    
                    m.Price.Price = dOrderPrice
                    m.WithLimitPrice.Price = Val(brokerMessage("GenesisLimitPrice"))
            End Select
        Else
            txtSymbol.Text = ""
        End If
    Else
        m.bGenesisSymbol = False
        txtSymbol.Text = brokerMessage("Symbol")
        cmdLookup.Visible = False
        
        Select Case UCase(brokerMessage("Type"))
            Case "STOP"
                txtPrice.Text = brokerMessage("StopPrice")
                txtWithLimitPrice.Text = ""
            Case "LIMIT"
                txtPrice.Text = brokerMessage("LimitPrice")
                txtWithLimitPrice.Text = ""
            Case "STOPWITHLIMIT"
                txtPrice.Text = brokerMessage("StopPrice")
                txtWithLimitPrice.Text = brokerMessage("LimitPrice")
        End Select
    End If
    SelectComboByText cboTifs, brokerMessage("TIF")
    If Len(brokerMessage("Expiration")) > 0 Then
        gdExpirationDate.YYYYMMDD = CLng(Val(brokerMessage("Expiration")))
    Else
        gdExpirationDate.Value = Date
    End If
        
    InitQuantityEditor
    If lNewQuantity > 0& Then
        m.Quantity.Price = lNewQuantity
    ElseIf Len(brokerMessage("Quantity")) > 0 Then
        m.Quantity.Price = Val(brokerMessage("Quantity"))
    Else
        m.Quantity.Price = 1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.ControlsFromBrokerMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ControlsToBrokerMessage
'' Description: Fill the broker message from the controls
'' Inputs:      Broker Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ControlsToBrokerMessage(brokerMessage As cBrokerMessage)
On Error GoTo ErrSection:

    brokerMessage.Add "Side", cboSides.Text
    brokerMessage.Add "Quantity", Str(m.Quantity.Price)
    brokerMessage.Add "Type", cboOrderTypes.Text
    
    If m.bGenesisSymbol Then
        brokerMessage.Add "GenesisSymbol", Trim(txtSymbol.Text)
        Select Case UCase(brokerMessage("Type"))
            Case "STOP"
                brokerMessage.Add "GenesisStopPrice", Str(m.Price.Price)
                brokerMessage.Add "GenesisLimitPrice", ""
            Case "LIMIT"
                brokerMessage.Add "GenesisLimitPrice", Str(m.Price.Price)
                brokerMessage.Add "GenesisStopPrice", ""
            Case "STOPWITHLIMIT"
                brokerMessage.Add "GenesisStopPrice", Str(m.Price.Price)
                brokerMessage.Add "GenesisLimitPrice", Str(m.WithLimitPrice.Price)
        End Select
    Else
        brokerMessage.Add "Symbol", Trim(txtSymbol.Text)
        Select Case UCase(brokerMessage("Type"))
            Case "STOP"
                brokerMessage.Add "StopPrice", Trim(txtPrice.Text)
                brokerMessage.Add "LimitPrice", ""
            Case "LIMIT"
                brokerMessage.Add "LimitPrice", Trim(txtPrice.Text)
                brokerMessage.Add "StopPrice", ""
            Case "STOPWITHLIMIT"
                brokerMessage.Add "StopPrice", Trim(txtPrice.Text)
                brokerMessage.Add "LimitPrice", Trim(txtWithLimitPrice.Text)
        End Select
    End If
    
    brokerMessage.Add "TIF", cboTifs.Text
    If UCase(brokerMessage("TIF")) = "GTD" Then
        brokerMessage.Add "Expiration", Str(gdExpirationDate.YYYYMMDD)
    Else
        brokerMessage.Add "Expiration", ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.ControlsToBrokerMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyControls
'' Description: Verify the values in the controls
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyControls() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If cboAccounts.ListIndex < 0 Then
        bReturn = False
        MoveFocus cboAccounts
        InfBox "Please select an account for the order", "!", , "Error"
    ElseIf cboSides.ListIndex < 0 Then
        bReturn = False
        MoveFocus cboSides
        InfBox "Please select a side for the order", "!", , "Error"
    ElseIf Val(txtQuantity.Text) <= 0 Then
        bReturn = False
        MoveFocus txtQuantity
        InfBox "Please select a valid quantity for the order", "!", , "Error"
    ElseIf Len(Trim(txtSymbol.Text)) = 0 Then
        bReturn = False
        MoveFocus txtSymbol
        InfBox "Please select a symbol for the order", "!", , "Error"
    ElseIf cboOrderTypes.ListIndex < 0 Then
        bReturn = False
        MoveFocus cboOrderTypes
        InfBox "Please select an order type for the order", "!", , "Error"
    ElseIf (Val(Trim(txtPrice.Text)) <= 0) And (txtPrice.Enabled = True) Then
        bReturn = False
        MoveFocus txtPrice
        InfBox "Please select a valid price for the order", "!", , "Error"
    ElseIf (Val(Trim(txtWithLimitPrice.Text)) <= 0) And (txtWithLimitPrice.Enabled = True) Then
        bReturn = False
        MoveFocus txtWithLimitPrice
        InfBox "Please select a valid limit price for the order", "!", , "Error"
    ElseIf cboTifs.ListIndex < 0 Then
        bReturn = False
        MoveFocus cboTifs
        InfBox "Please select a time-in-force for the order", "!", , "Error"
    End If
    
    VerifyControls = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSimpleOrder.VerifyControls"
    
End Function

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
    
    If KeyAscii = 0& Then
        Set astrSymbol = frmSymbolSelector.ShowMe(txtSymbol.Text, False, True, "Symbol to Buy/Sell", , , True)
    Else
        Set astrSymbol = frmSymbolSelector.ShowMe(Chr(KeyAscii), False, True, "Symbol to Buy/Sell", False, False, True)
    End If
    If astrSymbol.Size > 0 Then
        ChangeSymbol ConvertToTradeSymbol(astrSymbol(0), Date)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.LookupSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeSymbol
'' Description: Change the symbol on the UI
'' Inputs:      New Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeSymbol(ByVal strNewSymbol As String)
On Error GoTo ErrSection:

    Dim dMin As Double                  ' Minimum value for the price editor
    Dim bShowIfZero As Boolean          ' Show the price in the editor if it is zero?

    If strNewSymbol <> UCase(Trim(txtSymbol.Text)) Then
        txtSymbol.Text = strNewSymbol
                
        DM_GetBars m.Bars, txtSymbol.Text, , LastDailyDownload
        
        If IsSpreadSymbol(m.Bars.Prop(eBARS_Symbol)) Then
            dMin = -999999#
            bShowIfZero = True
        Else
            dMin = m.Bars.MinMove
            bShowIfZero = False
        End If
                
        m.Price.Init sbPrice, txtPrice, m.Bars, m.Bars(eBARS_Close, m.Bars.Size - 1), dMin, , , bShowIfZero
        m.WithLimitPrice.Init sbWithLimitPrice, txtWithLimitPrice, m.Bars, m.Bars(eBARS_Close, m.Bars.Size - 1), dMin, , , bShowIfZero
        InitQuantityEditor
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.ChangeSymbol"
    
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
Private Sub InitQuantityEditor()
On Error GoTo ErrSection:

    g.Broker.InitQuantityEditor m.Quantity, sbQuantity, txtQuantity, AccountID, txtSymbol.Text

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSimpleOrder.InitQuantityEditor"
    
End Sub

