VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCotSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraCommercialsProxy 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1380
      Width           =   6315
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmCotSettings.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCotSettings.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtProxyLookback 
         Height          =   285
         Left            =   5340
         TabIndex        =   10
         Top             =   0
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmCotSettings.frx":0068
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
         Tip             =   "frmCotSettings.frx":0088
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":00A8
      End
      Begin HexUniControls.ctlUniTextBoxXP txtProxyAverage 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Top             =   0
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmCotSettings.frx":00C4
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
         Tip             =   "frmCotSettings.frx":00E4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":0104
      End
      Begin HexUniControls.ctlUniLabelXP lblProxyLookback 
         Height          =   195
         Left            =   3720
         Top             =   45
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
         Caption         =   "frmCotSettings.frx":0120
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCotSettings.frx":0168
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":0188
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblProxy 
         Height          =   255
         Left            =   0
         Top             =   15
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
         Caption         =   "frmCotSettings.frx":01A4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCotSettings.frx":020A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":022A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraWillVal 
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   2700
      Width           =   6435
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmCotSettings.frx":0246
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCotSettings.frx":0272
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0292
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtShortTermWV 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   0
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmCotSettings.frx":02AE
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
         Tip             =   "frmCotSettings.frx":02CE
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":02EE
      End
      Begin HexUniControls.ctlUniTextBoxXP txtLongTermWV 
         Height          =   285
         Left            =   3360
         TabIndex        =   2
         Top             =   0
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmCotSettings.frx":030A
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
         Tip             =   "frmCotSettings.frx":032A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":034A
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBarsWV 
         Height          =   285
         Left            =   5700
         TabIndex        =   4
         Top             =   0
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmCotSettings.frx":0366
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
         Tip             =   "frmCotSettings.frx":0386
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":03A6
      End
      Begin HexUniControls.ctlUniLabelXP lblWillVal 
         Height          =   255
         Left            =   0
         Top             =   30
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
         Caption         =   "frmCotSettings.frx":03C2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCotSettings.frx":040A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":042A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblLongTerm 
         Height          =   255
         Left            =   2460
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
         Caption         =   "frmCotSettings.frx":0446
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCotSettings.frx":047A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":049A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBarsWV 
         Height          =   255
         Left            =   4260
         Top             =   30
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
         Caption         =   "frmCotSettings.frx":04B6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCotSettings.frx":04F8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":0518
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtPercentD 
      Height          =   285
      Left            =   5820
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmCotSettings.frx":0534
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
      Tip             =   "frmCotSettings.frx":0554
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0574
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   375
      Left            =   1245
      TabIndex        =   7
      Top             =   3240
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
      Caption         =   "frmCotSettings.frx":0590
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCotSettings.frx":05BC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":05DC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDefaults 
         Height          =   375
         Left            =   1470
         TabIndex        =   9
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
         Caption         =   "frmCotSettings.frx":05F8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCotSettings.frx":0636
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":0656
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2940
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
         Caption         =   "frmCotSettings.frx":0672
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCotSettings.frx":06A0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":06C0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRun 
         Default         =   -1  'True
         Height          =   375
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
         Caption         =   "frmCotSettings.frx":06DC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCotSettings.frx":0712
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":0732
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtMktSentYears 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   570
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmCotSettings.frx":074E
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
      Tip             =   "frmCotSettings.frx":076E
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":078E
   End
   Begin HexUniControls.ctlUniTextBoxXP txtCotYears 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   990
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmCotSettings.frx":07AA
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
      Tip             =   "frmCotSettings.frx":07CA
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":07EA
   End
   Begin HexUniControls.ctlUniTextBoxXP txtPercentK 
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Top             =   2250
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmCotSettings.frx":0806
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
      Tip             =   "frmCotSettings.frx":0826
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0846
   End
   Begin HexUniControls.ctlUniTextBoxXP txtStochBars 
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Top             =   2250
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmCotSettings.frx":0862
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
      Tip             =   "frmCotSettings.frx":0882
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":08A2
   End
   Begin HexUniControls.ctlUniTextBoxXP txtADXBars 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   1830
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmCotSettings.frx":08BE
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
      Tip             =   "frmCotSettings.frx":08DE
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":08FE
   End
   Begin MSComctlLib.ImageCombo cboSymbolGroup 
      Height          =   330
      Left            =   5040
      TabIndex        =   15
      Top             =   180
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin gdOCX.gdSelectDate CalcDate 
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      AllowWeekends   =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraSentiment 
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmCotSettings.frx":091A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCotSettings.frx":0946
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0966
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optWilliams 
         Height          =   255
         Left            =   1980
         TabIndex        =   18
         Top             =   0
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
         Caption         =   "frmCotSettings.frx":0982
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCotSettings.frx":09DE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":09FE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optGenesis 
         Height          =   255
         Left            =   60
         TabIndex        =   21
         Top             =   0
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
         Caption         =   "frmCotSettings.frx":0A1A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmCotSettings.frx":0A66
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCotSettings.frx":0A86
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label2 
      Height          =   255
      Left            =   180
      Top             =   3660
      Visible         =   0   'False
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
      Caption         =   "frmCotSettings.frx":0AA2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0AEE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0B0E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   255
      Left            =   3720
      Top             =   2280
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
      Caption         =   "frmCotSettings.frx":0B2A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0B6E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0B8E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblPercentD 
      Height          =   255
      Left            =   5340
      Top             =   2280
      Visible         =   0   'False
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
      Caption         =   "frmCotSettings.frx":0BAA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0BD0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0BF0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblMktSentYears 
      Height          =   255
      Left            =   120
      Top             =   600
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
      Caption         =   "frmCotSettings.frx":0C0C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0C76
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0C96
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblCotYears 
      Height          =   255
      Left            =   120
      Top             =   1020
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
      Caption         =   "frmCotSettings.frx":0CB2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0D16
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0D36
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblPercentK 
      Height          =   255
      Left            =   2550
      Top             =   2280
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
      Caption         =   "frmCotSettings.frx":0D52
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0D78
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0D98
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblStoch 
      Height          =   255
      Left            =   120
      Top             =   2280
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
      Caption         =   "frmCotSettings.frx":0DB4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0DFC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0E1C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblADX 
      Height          =   255
      Left            =   120
      Top             =   1860
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
      Caption         =   "frmCotSettings.frx":0E38
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0E72
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0E92
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblSymbolGroup 
      Height          =   255
      Left            =   3840
      Top             =   225
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
      Caption         =   "frmCotSettings.frx":0EAE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0EE8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0F08
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblDate 
      Height          =   255
      Left            =   120
      Top             =   180
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
      Caption         =   "frmCotSettings.frx":0F24
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCotSettings.frx":0F5C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCotSettings.frx":0F7C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmCotSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCotSettings.frm
'' Description: Allows the user to set up and run the COT report
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Bvld Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/28/2002   D Jarmuth   Created
'' 01/26/2012   DAJ         Added Premium Spread columns
'' 01/27/2012   DAJ         Fixed show/hide of Premium Spread based on enablement
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean
    strDefaults As String
    strNew As String
    dDate As Double
    FieldTable As cGdTable
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcDate_LostFocus
'' Description: Force the calculation date to always be a Friday
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcDate_LostFocus()
On Error GoTo ErrSection:

    Do While Weekday(CalcDate) <> vbFriday
        CalcDate = CalcDate - 1
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.CalcDate.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: When the user clicks on the cancel button, unload the form
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
    RaiseError "frmCotSettings.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDefaults_Click
'' Description: If the user clicks on the Defaults button, set all of the
''              controls back to the default values
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDefaults_Click()
On Error GoTo ErrSection:

    Dim dDate As Double                 ' Previous Friday
    Dim strDefaults As String           ' Default string
    
    dDate = Date
    Do While Weekday(dDate) <> vbFriday
        dDate = dDate - 1
    Loop
    
'    If g.SymbolPool.FieldNumForID("GRP:COT067.GRP") <> -1 Then
'        strDefaults = "GRP:COT067.GRP;7;14;3;3;3;3;2;22;156;GC-067;TQ-067"
'    If g.SymbolPool.FieldNumForID("GRP:CONT067.GRP") <> -1 Then
'        strDefaults = "GRP:CONT067.GRP;7;14;3;3;3;3;2;22;156;GC-067;TQ-067"
'    Else
'        strDefaults = "GRP:ALL FUTURES.GRP;7;14;3;3;3;3;2;22;156;GC-067;TQ-067"
'    End If

    strDefaults = "GRP:CONT067.GRP;7;14;3;3;3;3;-1;2;22;156;GC-067;TQ-067;8;3"

    SetControls strDefaults
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.cmdDefaults.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRun_Click
'' Description: When the user chooses to run, save the defaults and unload this
''              form allowing the COT Report form to run
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRun_Click()
On Error GoTo ErrSection:

    Dim strDefaults As String           ' Default string
    
    cmdRun.SetFocus
    
    If ProcessIsBusy Then Exit Sub
    
    strDefaults = cboSymbolGroup.SelectedItem.Key
    strDefaults = strDefaults & ";" & txtADXBars.Text
    strDefaults = strDefaults & ";" & txtStochBars.Text
    strDefaults = strDefaults & ";" & txtPercentK.Text
    strDefaults = strDefaults & ";" & txtPercentD.Text
    strDefaults = strDefaults & ";" & txtCotYears.Text
    strDefaults = strDefaults & ";" & txtMktSentYears.Text
    strDefaults = strDefaults & ";" & Str(CInt(optGenesis.Value))
    strDefaults = strDefaults & ";" & txtShortTermWV.Text
    strDefaults = strDefaults & ";" & txtLongTermWV.Text
    strDefaults = strDefaults & ";" & txtBarsWV.Text
    strDefaults = strDefaults & ";"
    strDefaults = strDefaults & ";"
    strDefaults = strDefaults & ";" & txtProxyAverage.Text
    strDefaults = strDefaults & ";" & txtProxyLookback.Text
    
    m.strNew = strDefaults
    m.dDate = CalcDate.Value
    
    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.cmdRun.Click", eGDRaiseError_Show
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
    RaiseError "frmCotSettings.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center it, set the caption, and get
''              the values from the last run
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_COTReport"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Me.Caption = "COT Report Settings"
    
    CalcDate.MaxDate = Date + 1
    cboSymbolGroup.ImageList = frmMain.img16
    LoadCombo
    cboSymbolGroup.Locked = True
    
    ' Hide the symbol group option for now 04/23/2002 DAJ
    lblSymbolGroup.Visible = False
    cboSymbolGroup.Visible = False
    
'    If g.SymbolPool.FieldNumForID("GRP:COT067.GRP") <> -1 Then
'        m.strDefaults = "GRP:COT067.GRP;7;14;3;3;3;3;2;22;156;GC-067;TQ-067"
'    If g.SymbolPool.FieldNumForID("GRP:CONT067.GRP") <> -1 Then
'        m.strDefaults = "GRP:CONT067.GRP;7;14;3;3;3;3;2;22;156;GC-067;TQ-067"
'    Else
'        m.strDefaults = "GRP:ALL FUTURES.GRP;7;14;3;3;3;3;2;22;156;GC-067;TQ-067"
'    End If
    
    m.strDefaults = "GRP:CONT067.GRP;7;14;3;3;3;3;-1;2;22;156;GC-067;TQ-067;8;3"
    m.strDefaults = GetIniFileProperty("Defaults", m.strDefaults, "Defaults", AddSlash(App.Path) & "CotRpt.INI")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As Boolean
On Error GoTo ErrSection:

    Dim dDate As Double                 ' Temporary date
    Dim astrEnglish As New cGdArray     ' Array of english expressions
    Dim lIndex As Long
    Dim strIniFile As String
    Dim lMaxWeeks As Long
    Dim lTemp As Long
    Dim bClearLastRun As Boolean        ' Clear the last run?
        
    strIniFile = AddSlash(App.Path) & "CotRpt.INI"
        
    Set m.FieldTable = New cGdTable
    m.FieldTable.CreateField eGDARRAY_Strings, , "Active"
    m.FieldTable.CreateField eGDARRAY_Strings, , "Name"
    m.FieldTable.CreateField eGDARRAY_Strings, , "English"
    m.FieldTable.CreateField eGDARRAY_Strings, , "Show"
    m.FieldTable.CreateField eGDARRAY_Strings, , "FirstHeader"
    m.FieldTable.CreateField eGDARRAY_Strings, , "SecondHeader"
    m.FieldTable.CreateField eGDARRAY_Strings, , "Description"
    m.FieldTable.CreateField eGDARRAY_Strings, , "ColWidth"
    FillTable
        
    ' Make sure that the Continuous 67 Symbol Group exists
    If g.SymbolPool.FieldNumForID("GRP:CONT067.GRP") = -1 Then
        InfBox "Symbol Group CONT067 does not exist", "!", , "Error"
        Exit Function
    End If
    
    ' Force the date to be a Friday
    CalcDate = LastDailyDownload
    Do While Weekday(CalcDate) <> vbFriday
        CalcDate = CalcDate - 1
    Loop
    
    ' Set all of the controls to a default or the last run
    SetControls m.strDefaults
    fraWillVal.Visible = WillValExists
    fraCommercialsProxy.Visible = CommercialsProxyExists

    ' Show the form
    ShowForm Me, True
    
    ' If the user hit run, save the new values and run the report
    If m.bOK Then
        SetIniFileProperty "Defaults", m.strNew, "Defaults", strIniFile
        
        bClearLastRun = False
        For lIndex = 0 To m.FieldTable.NumRecords - 1
            Select Case m.FieldTable(1, lIndex)
                Case "Commercials"
                    m.FieldTable(2, lIndex) = "COT Commercials"
                Case "Commercials Change"
                    m.FieldTable(2, lIndex) = "COT Commercials - COT Commercials.1"
                Case "Commercials Index"
                    m.FieldTable(2, lIndex) = "COT Commercials Index(" & Parse(m.strNew, ";", 6) & ")"
                Case "Commercials Index Change"
                    m.FieldTable(2, lIndex) = "COT Commercials Index(" & Parse(m.strNew, ";", 6) & ") - COT Commercials Index(" & Parse(m.strNew, ";", 6) & ").1"
                Case "Commercials Proxy Index"
                    If CommercialsProxyExists = True Then
                        m.FieldTable(2, lIndex) = "Commercials Proxy Index(" & Parse(m.strNew, ";", 14) & ", " & Parse(m.strNew, ";", 15) & " * 52) Of Weekly"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "Commercials Proxy Index Change"
                    If CommercialsProxyExists = True Then
                        m.FieldTable(2, lIndex) = "Commercials Proxy Index(" & Parse(m.strNew, ";", 14) & ", " & Parse(m.strNew, ";", 15) & " * 52) Of Weekly - Commercials Proxy Index(" & Parse(m.strNew, ";", 14) & ", " & Parse(m.strNew, ";", 15) & " * 52).1 Of Weekly"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "Large Spec"
                    m.FieldTable(2, lIndex) = "COT Large Spec"
                Case "Large Spec Change"
                    m.FieldTable(2, lIndex) = "COT Large Spec - COT Large Spec.1"
                Case "Large Spec Index"
                    'm.FieldTable(2, lIndex) = "COT Large Spec Index(" & Parse(m.strNew, ";", 6) & ")"
                    m.FieldTable(2, lIndex) = "Stochastic Custom(COT Large Spec Of Weekly , " & Parse(m.strNew, ";", 6) & " * 52)"
                Case "Large Spec Index Change"
                    'm.FieldTable(2, lIndex) = "COT Large Spec Index(" & Parse(m.strNew, ";", 6) & ") - COT Large Spec Index(" & Parse(m.strNew, ";", 6) & ").1"
                    m.FieldTable(2, lIndex) = "Stochastic Custom(COT Large Spec Of Weekly, " & Parse(m.strNew, ";", 6) & " * 52) - Stochastic Custom(COT Large Spec Of Weekly, " & Parse(m.strNew, ";", 6) & " * 52).1"
                Case "Small Spec"
                    m.FieldTable(2, lIndex) = "COT Small Spec"
                Case "Small Spec Change"
                    m.FieldTable(2, lIndex) = "COT Small Spec - COT Small Spec.1"
                Case "Small Spec Index"
                    m.FieldTable(2, lIndex) = "Stochastic Custom(COT Small Spec Of Weekly, " & Parse(m.strNew, ";", 6) & " * 52)"
                Case "Small Spec Index Change"
                    m.FieldTable(2, lIndex) = "Stochastic Custom(COT Small Spec Of Weekly, " & Parse(m.strNew, ";", 6) & " * 52) - Stochastic Custom(COT Small Spec Of Weekly, " & Parse(m.strNew, ";", 6) & " * 52).1"
                Case "Genesis Sentiment"
                    m.FieldTable(2, lIndex) = "Genesis Sentiment"
                Case "Genesis Sentiment Change"
                    m.FieldTable(2, lIndex) = "Genesis Sentiment - Genesis Sentiment.1"
                Case "Genesis Sentiment Index"
                    m.FieldTable(2, lIndex) = "Genesis Sentiment Index(" & Parse(m.strNew, ";", 7) & ")"
                Case "Genesis Sentiment Index Change"
                    m.FieldTable(2, lIndex) = "Genesis Sentiment Index(" & Parse(m.strNew, ";", 7) & ") - Genesis Sentiment Index(" & Parse(m.strNew, ";", 7) & ").1"
                Case "Genesis Setup Strength"
                    m.FieldTable(2, lIndex) = "Genesis Sentiment Index(" & Parse(m.strNew, ";", 7) & ") - COT Commercials Index(" & Parse(m.strNew, ";", 6) & ")"
                Case "Genesis Setup Strength Change"
                    m.FieldTable(2, lIndex) = "(Genesis Sentiment Index(" & Parse(m.strNew, ";", 7) & ") - COT Commercials Index(" & Parse(m.strNew, ";", 6) & ")) - (Genesis Sentiment Index(" & Parse(m.strNew, ";", 7) & ") - COT Commercials Index(" & Parse(m.strNew, ";", 6) & ")).1"
                Case "Genesis Proxy Setup Strength"
                    If CommercialsProxyExists = True Then
                        m.FieldTable(2, lIndex) = "Genesis Sentiment Index(" & Parse(m.strNew, ";", 7) & ") - Commercials Proxy Index(" & Parse(m.strNew, ";", 14) & ", " & Parse(m.strNew, ";", 15) & " * 52) Of Weekly"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "Genesis Proxy Setup Strength Change"
                    If CommercialsProxyExists = True Then
                        m.FieldTable(2, lIndex) = "(Genesis Sentiment Index(" & Parse(m.strNew, ";", 7) & ") - Commercials Proxy Index(" & Parse(m.strNew, ";", 14) & ", " & Parse(m.strNew, ";", 15) & " * 52) Of Weekly) - (Genesis Sentiment Index(" & Parse(m.strNew, ";", 7) & ") - Commercials Proxy Index(" & Parse(m.strNew, ";", 14) & ", " & Parse(m.strNew, ";", 15) & " * 52) Of Weekly).1"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "TN Consensus"
                    m.FieldTable(2, lIndex) = "TN Consensus"
                Case "TN Consensus Change"
                    m.FieldTable(2, lIndex) = "TN Consensus - TN Consensus.1"
                Case "TN Consensus Index"
                    m.FieldTable(2, lIndex) = "Stochastic Custom(TN Consensus Of Weekly, " & LookBack(Parse(m.strNew, ";", 7)) & " * 52)"
                Case "TN Consensus Index Change"
                    m.FieldTable(2, lIndex) = "Stochastic Custom(TN Consensus Of Weekly, " & LookBack(Parse(m.strNew, ";", 7)) & " * 52) - Stochastic Custom(TN Consensus Of Weekly, " & LookBack(Parse(m.strNew, ";", 7)) & " * 52).1"
                Case "TN Setup Strength"
                    m.FieldTable(2, lIndex) = "Stochastic Custom(TN Consensus Of Weekly, " & LookBack(Parse(m.strNew, ";", 7)) & " * 52) - COT Commercials Index(" & Parse(m.strNew, ";", 6) & ")"
                Case "TN Setup Strength Change"
                    m.FieldTable(2, lIndex) = "(Stochastic Custom(TN Consensus Of Weekly, " & LookBack(Parse(m.strNew, ";", 7)) & " * 52) - COT Commercials Index(" & Parse(m.strNew, ";", 6) & ")) - (Stochastic Custom(TN Consensus Of Weekly, " & LookBack(Parse(m.strNew, ";", 7)) & " * 52) - COT Commercials Index(" & Parse(m.strNew, ";", 6) & ")).1"
                Case "LW Proxy Setup Strength"
                    If CommercialsProxyExists = True Then
                        m.FieldTable(2, lIndex) = "Stochastic Custom(TN Consensus Of Weekly, " & LookBack(Parse(m.strNew, ";", 7)) & " * 52) - Commercials Proxy Index(" & Parse(m.strNew, ";", 14) & ", " & Parse(m.strNew, ";", 15) & " * 52) Of Weekly"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "LW Proxy Setup Strength Change"
                    If CommercialsProxyExists = True Then
                        m.FieldTable(2, lIndex) = "(Stochastic Custom(TN Consensus Of Weekly, " & LookBack(Parse(m.strNew, ";", 7)) & " * 52) - Commercials Proxy Index(" & Parse(m.strNew, ";", 14) & ", " & Parse(m.strNew, ";", 15) & " * 52) Of Weekly) - (Stochastic Custom(TN Consensus Of Weekly, " & LookBack(Parse(m.strNew, ";", 7)) & " * 52) - Commercials Proxy Index(" & Parse(m.strNew, ";", 14) & ", " & Parse(m.strNew, ";", 15) & " * 52) Of Weekly).1"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "Will-Val"
                    If WillValExists Then
                        m.FieldTable(2, lIndex) = "WillVal(Close of GC, " & Parse(m.strNew, ";", 9) & ", " & Parse(m.strNew, ";", 10) & ", " & Parse(m.strNew, ";", 11) & ")"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "Will-Val Change"
                    If WillValExists Then
                        m.FieldTable(2, lIndex) = "WillVal(Close of GC, " & Parse(m.strNew, ";", 9) & ", " & Parse(m.strNew, ";", 10) & ", " & Parse(m.strNew, ";", 11) & ") - WillVal(Close of GC, " & Parse(m.strNew, ";", 9) & ", " & Parse(m.strNew, ";", 10) & ", " & Parse(m.strNew, ";", 11) & ").1"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "ADX"
                    m.FieldTable(2, lIndex) = "ADX(" & Parse(m.strNew, ";", 2) & ")"
                Case "ADX Change"
                    m.FieldTable(2, lIndex) = "ADX(" & Parse(m.strNew, ";", 2) & ") - ADX(" & Parse(m.strNew, ";", 2) & ").1"
                Case "Stoch"
                    m.FieldTable(2, lIndex) = "StochK(" & Parse(m.strNew, ";", 3) & ", " & Parse(m.strNew, ";", 4) & ")"
                Case "Stoch Change"
                    m.FieldTable(2, lIndex) = "StochK(" & Parse(m.strNew, ";", 3) & ", " & Parse(m.strNew, ";", 4) & ") - StochK(" & Parse(m.strNew, ";", 3) & ", " & Parse(m.strNew, ";", 4) & ").1"
                Case "Premium Spread One Contract Out"
                    If CheckPremiumSpreadRecord(lIndex, True) Then
                        bClearLastRun = True
                    End If
                    If CalendarSpreadExists Then
                        m.FieldTable(2, lIndex) = "Calendar Spread(1)"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "Premium Spread One Contract Out Change"
                    If CheckPremiumSpreadRecord(lIndex, False) Then
                        bClearLastRun = True
                    End If
                    If CalendarSpreadExists Then
                        m.FieldTable(2, lIndex) = "Calendar Spread(1) - Calendar Spread(1).1"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "Premium Spread Two Contracts Out"
                    If CheckPremiumSpreadRecord(lIndex, True) Then
                        bClearLastRun = True
                    End If
                    If CalendarSpreadExists Then
                        m.FieldTable(2, lIndex) = "Calendar Spread(2)"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
                Case "Premium Spread Two Contracts Out Change"
                    If CheckPremiumSpreadRecord(lIndex, False) Then
                        bClearLastRun = True
                    End If
                    If CalendarSpreadExists Then
                        m.FieldTable(2, lIndex) = "Calendar Spread(2) - Calendar Spread(2).1"
                    Else
                        m.FieldTable(2, lIndex) = ""
                    End If
            End Select
        Next lIndex
        
        ' Save the table to the Ini File
        SetIniFileProperty "NumFields", m.FieldTable.NumRecords, "Fields", strIniFile
        For lIndex = 0 To m.FieldTable.NumRecords - 1
            SetIniFileProperty "Field" & lIndex, m.FieldTable(0, lIndex) & ";" & _
                    m.FieldTable(1, lIndex) & ";" & m.FieldTable(2, lIndex) & ";" & _
                    m.FieldTable(3, lIndex) & ";" & m.FieldTable(4, lIndex) & _
                    ";" & m.FieldTable(5, lIndex) & ";" & m.FieldTable(6, lIndex), "Fields", strIniFile
        Next lIndex
        
        lMaxWeeks = 0
        lTemp = CLng(txtLongTermWV) * 3 + CLng(txtBarsWV)
        If lTemp > lMaxWeeks Then lMaxWeeks = lTemp
        lTemp = CLng(txtCotYears) * 52
        If lTemp > lMaxWeeks Then lMaxWeeks = lTemp
        lTemp = CLng(txtMktSentYears) * 52
        If lTemp > lMaxWeeks Then lMaxWeeks = lTemp
        lTemp = CLng(txtADXBars) * 10
        If lTemp > lMaxWeeks Then lMaxWeeks = lTemp
        lTemp = CLng(txtPercentK) + CLng(txtStochBars)
        If lTemp > lMaxWeeks Then lMaxWeeks = lTemp
        
        If bClearLastRun Then
            SetIniFileProperty "LastRun", "", "LastRun", strIniFile
        End If
        
'        frmCotReport.ShowMe Parse(m.strNew, ";", 1), astrEnglish, m.dDate
        frmCotReport.ShowMe Parse(m.strNew, ";", 1), m.FieldTable, m.dDate, lMaxWeeks
        Set astrEnglish = Nothing
    End If
    
    ShowMe = m.bOK
    Set m.FieldTable = Nothing
    Unload Me

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotSettings.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the 'X' in the corner, hide the form
''              and allow the ShowMe to finish
'' Inputs:      Whether or not to cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Loads the Symbol Group combo with all of the active Symbol
''              Groups from the symbol pool
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo()
On Error Resume Next

    Dim lIndex As Long                  ' Index for a for loop
    Dim strID As String                 ' Symbol pool ID for the field
    Dim strType As String               ' Type of thing (i.e. Filter, Criteria, etc)
    Dim strPicture As String            ' Picture to use in the combo box
    Dim strSelID As String              ' ID of the currently selected item
    Dim bSelExists As Boolean           ' Old selection still exists
    Dim iSortStart As Long              ' Where to start the sort
    Dim strItem As String               ' Item to add to the combo box
    Dim aItems As New cGdArray          ' Items to add to the combo box
    Dim obj As Object                   ' Symbol Pool Object
    Dim strKey As String                ' Key into the registry
    Dim bScans As Boolean               ' Are we doing scans?
   
    If cboSymbolGroup.ComboItems.Count > 0 Then
        strSelID = cboSymbolGroup.SelectedItem.Key
        cboSymbolGroup.ComboItems.Clear
    End If
    
    ' get list of items to put into combo list
    With g.SymbolPool
        For lIndex = 0 To .ArrayTable.NumFields - 1
            strID = .FieldID(lIndex)
            If Len(strID) = 0 Then
                strType = "" '???
            Else
                strType = Left(strID, 3)
                strPicture = ""
                Set obj = .PoolObject(strID)
                If UCase(strType) = "GRP" Then
                    If strID <> "GRP:_FLAGS_.GRP" And obj.IsActive = True Then
                        strPicture = ToolbarIcon("ID_SymbolGroups")
                        
                        If strID = strSelID Then
                            bSelExists = True
                        End If
                        
                        If iSortStart = 0 And lIndex >= g.SymbolPool.OtherFieldsStart Then
                            iSortStart = aItems.Size
                        End If
                        
                        aItems.Add .ArrayTable.FieldName(lIndex) & vbTab _
                                & strID & vbTab & strPicture
                    End If
                End If
            End If
        Next
    End With
    If iSortStart > 0 Then
        aItems.Sort eGdSort_IgnoreCase, iSortStart
    End If

    For lIndex = 0 To aItems.Size - 1
        strItem = aItems(lIndex)
        cboSymbolGroup.ComboItems.Add , Parse(strItem, vbTab, 2), _
            Parse(strItem, vbTab, 1), Parse(strItem, vbTab, 3)
    Next

    If bSelExists Then
        cboSymbolGroup.ComboItems(strSelID).Selected = True
    Else
        cboSymbolGroup.ComboItems(1).Selected = True
    End If

    cboSymbolGroup.Refresh

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetControls
'' Description: Set the controls according to the settings passed in
'' Inputs:      Settings to set the controls to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetControls(ByVal strSettings As String)
On Error GoTo ErrSection:

    Dim strTemp As String               ' Temporary string

    Set cboSymbolGroup.SelectedItem = cboSymbolGroup.ComboItems(Parse(strSettings, ";", 1))
    
    txtADXBars.Text = Parse(strSettings, ";", 2)
    txtStochBars.Text = Parse(strSettings, ";", 3)
    txtPercentK.Text = Parse(strSettings, ";", 4)
    txtPercentD.Text = Parse(strSettings, ";", 5)
    txtCotYears.Text = Parse(strSettings, ";", 6)
    txtMktSentYears.Text = Parse(strSettings, ";", 7)
    If Val(Parse(strSettings, ";", 8)) <> 0 Then optGenesis = True Else optWilliams = True
    txtShortTermWV.Text = Parse(strSettings, ";", 9)
    txtLongTermWV.Text = Parse(strSettings, ";", 10)
    txtBarsWV.Text = Parse(strSettings, ";", 11)
    strTemp = Parse(strSettings, ";", 14)
    If Len(strTemp) = 0 Then txtProxyAverage.Text = "8" Else txtProxyAverage.Text = strTemp
    strTemp = Parse(strSettings, ";", 15)
    If Len(strTemp) = 0 Then txtProxyLookback.Text = "3" Else txtProxyLookback.Text = strTemp

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotSettings.SetControls", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckControls
'' Description: Verify that all of the controls have "good" values
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckControls()
On Error GoTo ErrSection:

    If ValOfText(txtADXBars.Text) = 0 Then txtADXBars.Text = "7" Else txtADXBars.Text = Trim(CStr(ValOfText(txtADXBars.Text)))
    If ValOfText(txtStochBars.Text) = 0 Then txtStochBars.Text = "14" Else txtStochBars.Text = Trim(CStr(ValOfText(txtStochBars.Text)))
    If ValOfText(txtPercentD.Text) = 0 Then txtPercentD.Text = "3" Else txtPercentD.Text = Trim(CStr(ValOfText(txtPercentD.Text)))
    If ValOfText(txtPercentK.Text) = 0 Then txtPercentK.Text = "3" Else txtPercentK.Text = Trim(CStr(ValOfText(txtPercentK.Text)))
    If ValOfText(txtCotYears.Text) = 0 Then txtCotYears.Text = "3" Else txtCotYears.Text = Trim(CStr(ValOfText(txtCotYears.Text)))
    If ValOfText(txtMktSentYears.Text) = 0 Then txtMktSentYears.Text = "3" Else txtMktSentYears.Text = Trim(CStr(ValOfText(txtMktSentYears.Text)))
    If ValOfText(txtShortTermWV.Text) = 0 Then txtShortTermWV.Text = "2" Else txtShortTermWV.Text = Trim(CStr(ValOfText(txtShortTermWV.Text)))
    If ValOfText(txtLongTermWV.Text) = 0 Then txtLongTermWV.Text = "22" Else txtLongTermWV.Text = Trim(CStr(ValOfText(txtLongTermWV.Text)))
    If ValOfText(txtBarsWV.Text) = 0 Then txtBarsWV.Text = "156" Else txtBarsWV.Text = Trim(CStr(ValOfText(txtBarsWV.Text)))
    If ValOfText(txtProxyAverage.Text) = 0 Then txtProxyAverage.Text = "8" Else txtProxyAverage.Text = Trim(Str(ValOfText(txtProxyAverage.Text)))
    If ValOfText(txtProxyLookback.Text) = 0 Then txtProxyLookback.Text = "3" Else txtProxyLookback.Text = Trim(Str(ValOfText(txtProxyLookback.Text)))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.CheckControls", eGDRaiseError_Raise
    
End Sub

Private Sub txtADXBars_LostFocus()
On Error GoTo ErrSection:

    CheckControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.txtADXBars.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtStochBars_LostFocus()
On Error GoTo ErrSection:

    CheckControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.txtStochBars.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtPercentD_LostFocus()
On Error GoTo ErrSection:

    CheckControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.txtPercentD.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtPercentK_LostFocus()
On Error GoTo ErrSection:

    CheckControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.txtPercentK.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtCotYears_LostFocus()
On Error GoTo ErrSection:

    CheckControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.txtCotYears.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtMktSentYears_LostFocus()
On Error GoTo ErrSection:

    CheckControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.txtMktSentYears.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtShortTermWV_LostFocus()
On Error GoTo ErrSection:

    CheckControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.txtShortTermWV.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtLongTermWV_LostFocus()
On Error GoTo ErrSection:

    CheckControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.txtLongTermWV.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtBarsWV_LostFocus()
On Error GoTo ErrSection:

    CheckControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotSettings.txtBarsWV.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Function WillValExists() As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    If InStr(g.strAuthorizationString, ",INC,") <> 0 Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                    "WHERE [FunctionName]='WillVal';", dbOpenDynaset)
        If Not rs.EOF Then
            WillValExists = True
        End If
    End If

ErrExit:
    Set rs = Nothing
    Exit Function

ErrSection:
    RaiseError "frmCotSettings.WillValExists", eGDRaiseError_Raise

End Function

Private Function FillTable()
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim lField As Long
    Dim strTemp As String
    Dim lNumItems As Long
    Dim strIniFile As String
    Dim astrFields As cGdArray
    Dim bShowProxy As Boolean           ' Show the commercial proxy columns?
    Dim bShowPremium As Boolean         ' Show the premium spread columns?
    
    strIniFile = AddSlash(App.Path) & "CotRpt.INI"
    lNumItems = GetIniFileProperty("NumFields", 0, "Fields", strIniFile)
    Set astrFields = New cGdArray
    astrFields.Create eGDARRAY_Strings, 32
    
    If lNumItems > 0 Then
        For lIndex = 0 To lNumItems - 1
            strTemp = GetIniFileProperty("Field" & lIndex, "", "Fields", strIniFile)
            strTemp = Replace(strTemp, "LW Sentiment", "TN Consensus")
            strTemp = Replace(strTemp, "LW Setup", "TN Setup")
            strTemp = Replace(strTemp, "True", "1")
            strTemp = Replace(strTemp, "False", "0")
            
            astrFields(lIndex) = strTemp
            For lField = 0 To m.FieldTable.NumFields - 1
                m.FieldTable(lField, lIndex) = Parse(strTemp, ";", lField + 1)
            Next lField
            
            If m.FieldTable(7, lIndex) = m.FieldTable.FieldArray(7).NullValue Then
                m.FieldTable(7, lIndex) = ""
            End If
        Next lIndex
    Else
        astrFields(0) = "1" & ";" & "SYMBOL" & ";" & "" & ";" & "1" & ";" & "SYMBOL" & ";" & "SYMBOL" & ";" & "" & ";" & ""
        astrFields(1) = "1" & ";" & "Commercials" & ";" & "" & ";" & "1" & ";" & "Commercials" & ";" & "Now" & ";" & "Commercials Net Position (Longs - Shorts)" & ";" & ""
        astrFields(2) = "1" & ";" & "Commercials Change" & ";" & "" & ";" & "1" & ";" & "Commercials" & ";" & "Change" & ";" & "" & ";" & ""
        astrFields(3) = "1" & ";" & "Commercials Index" & ";" & "" & ";" & "1" & ";" & "Commercials" & ";" & "Idx" & ";" & "Commercials as percent between highest and lowest over lookback period" & ";" & ""
        astrFields(4) = "0" & ";" & "Commercials Index Change" & ";" & "" & ";" & "1" & ";" & "Commercials" & ";" & "Idx Chg" & ";" & "" & ";" & ""
        astrFields(5) = "0" & ";" & "Large Spec" & ";" & "" & ";" & "1" & ";" & "Large Specs" & ";" & "Now" & ";" & "Large Spec Net Position (Longs - Shorts)" & ";" & ""
        astrFields(6) = "0" & ";" & "Large Spec Change" & ";" & "" & ";" & "1" & ";" & "Large Specs" & ";" & "Change" & ";" & "" & ";" & ""
        astrFields(7) = "0" & ";" & "Large Spec Index" & ";" & "" & ";" & "1" & ";" & "Large Specs" & ";" & "Idx" & ";" & "Large Specs as percent between highest and lowest over lookback period" & ";" & ""
        astrFields(8) = "0" & ";" & "Large Spec Index Change" & ";" & "" & ";" & "1" & ";" & "Large Specs" & ";" & "Idx Chg" & ";" & "" & ";" & ""
        astrFields(9) = "0" & ";" & "Small Spec" & ";" & "" & ";" & "1" & ";" & "Small Specs" & ";" & "Now" & ";" & "Small Spec Net Position (Longs - Shorts)" & ";" & ""
        astrFields(10) = "0" & ";" & "Small Spec Change" & ";" & "" & ";" & "1" & ";" & "Small Specs" & ";" & "Change" & ";" & "" & ";" & ""
        astrFields(11) = "0" & ";" & "Small Spec Index" & ";" & "" & ";" & "1" & ";" & "Small Specs" & ";" & "Idx" & ";" & "Small Specs as percent between highest and lowest over lookback period" & ";" & ""
        astrFields(12) = "0" & ";" & "Small Spec Index Change" & ";" & "" & ";" & "1" & ";" & "Small Specs" & ";" & "Idx Chg" & ";" & "" & ";" & ""
        astrFields(13) = "1" & ";" & "Genesis Sentiment" & ";" & "" & ";" & "1" & ";" & "Genesis Sentiment" & ";" & "Now" & ";" & "Market Sentiment" & ";" & ""
        astrFields(14) = "0" & ";" & "Genesis Sentiment Change" & ";" & "" & ";" & "1" & ";" & "Genesis Sentiment" & ";" & "Change" & ";" & "" & ";" & ""
        astrFields(15) = "1" & ";" & "Genesis Sentiment Index" & ";" & "" & ";" & "1" & ";" & "Genesis Sentiment" & ";" & "Idx" & ";" & "Market Sentiment as percent between highest and lowest over lookback period" & ";" & ""
        astrFields(16) = "0" & ";" & "Genesis Sentiment Index Change" & ";" & "" & ";" & "1" & ";" & "Genesis Sentiment" & ";" & "Idx Chg" & ";" & "" & ";" & ""
        astrFields(17) = "1" & ";" & "Genesis Setup Strength" & ";" & "" & ";" & "1" & ";" & "Genesis Sentiment" & ";" & "SETUP" & ";" & "Setup Strength = Commercials (as %) - Sentiment (as %)" & ";" & ""
        astrFields(18) = "1" & ";" & "Genesis Setup Strength Change" & ";" & "" & ";" & "1" & ";" & "Genesis Sentiment" & ";" & "Setup Chg" & ";" & "" & ";" & ""
        astrFields(19) = "0" & ";" & "TN Consensus" & ";" & "" & ";" & "1" & ";" & "TN Consensus" & ";" & "Now" & ";" & "Market Sentiment" & ";" & ""
        astrFields(20) = "0" & ";" & "TN Consensus Change" & ";" & "" & ";" & "1" & ";" & "TN Consensus" & ";" & "Change" & ";" & "" & ";" & ""
        astrFields(21) = "0" & ";" & "TN Consensus Index" & ";" & "" & ";" & "1" & ";" & "TN Consensus" & ";" & "Idx" & ";" & "Market Sentiment as percent between highest and lowest over lookback period" & ";" & ""
        astrFields(22) = "0" & ";" & "TN Consensus Index Change" & ";" & "" & ";" & "1" & ";" & "TN Consensus" & ";" & "Idx Chg" & ";" & "" & ";" & ""
        astrFields(23) = "0" & ";" & "TN Setup Strength" & ";" & "" & ";" & "1" & ";" & "TN Consensus" & ";" & "SETUP" & ";" & "Setup Strength = Commercials (as %) - Sentiment (as %)" & ";" & ""
        astrFields(24) = "0" & ";" & "TN Setup Strength Change" & ";" & "" & ";" & "1" & ";" & "TN Consensus" & ";" & "Setup Chg" & ";" & "" & ";" & ""
        astrFields(25) = Str(CLng(WillValExists)) & ";" & "Will-Val" & ";" & "" & ";" & Str(CLng(WillValExists)) & ";" & "Will-Val" & ";" & "Value" & ";" & "WillVal: a spread indicator by Larry Williams" & ";" & ""
        astrFields(26) = "0" & ";" & "Will-Val Change" & ";" & "" & ";" & Str(CLng(WillValExists)) & ";" & "Will-Val" & ";" & "Change" & ";" & "" & ";" & ""
        astrFields(27) = Str(CLng(WillValExists)) & ";" & "Will-Val Symbol" & ";" & "" & ";" & Str(CLng(WillValExists)) & ";" & "Will-Val" & ";" & "Symbol" & ";" & "Click in this column to change the symbol for the WillVal spread" & ";" & ""
        astrFields(28) = "1" & ";" & "ADX" & ";" & "" & ";" & "1" & ";" & "ADX" & ";" & "Value" & ";" & "ADX: an indicator by Welles Wilder" & ";" & ""
        astrFields(29) = "0" & ";" & "ADX Change" & ";" & "" & ";" & "1" & ";" & "ADX" & ";" & "Change" & ";" & "" & ";" & ""
        astrFields(30) = "1" & ";" & "Stoch" & ";" & "" & ";" & "1" & ";" & "Stochastic" & ";" & "Value" & ";" & "StochK: the %K version of the stochastic indicator" & ";" & ""
        astrFields(31) = "0" & ";" & "Stoch Change" & ";" & "" & ";" & "1" & ";" & "Stochastic" & ";" & "Change" & ";" & "" & ";" & ""
        
        SetIniFileProperty "NumFields", astrFields.Size, "Fields", strIniFile
        For lIndex = 0 To astrFields.Size - 1
            SetIniFileProperty "Field" & lIndex, astrFields(lIndex), "Fields", strIniFile
            For lField = 0 To m.FieldTable.NumFields - 1
                m.FieldTable(lField, lIndex) = Parse(astrFields(lIndex), ";", lField + 1)
            Next lField
        Next lIndex
        
        lNumItems = 32
    End If
    
    If lNumItems = 32 Then
        lNumItems = 34
        bShowProxy = CommercialsProxyExists
        
        astrFields.Add Str(CLng(bShowProxy)) & ";" & "Commercials Proxy Index" & ";" & "" & ";" & Str(CLng(bShowProxy)) & ";" & "Commercials" & ";" & "Proxy" & ";" & "Commercials Proxy Index", 5
        astrFields.Add Str(CLng(bShowProxy)) & ";" & "Commercials Proxy Index Change" & ";" & "" & ";" & Str(CLng(bShowProxy)) & ";" & "Commercials" & ";" & "Proxy Chg" & ";" & "Commercials Proxy Index", 6
        
        SetIniFileProperty "NumFields", lNumItems, "Fields", strIniFile
        For lIndex = 0 To astrFields.Size - 1
            SetIniFileProperty "Field" & lIndex, astrFields(lIndex), "Fields", strIniFile
            For lField = 0 To m.FieldTable.NumFields - 1
                m.FieldTable(lField, lIndex) = Parse(astrFields(lIndex), ";", lField + 1)
            Next lField
            If m.FieldTable(7, lIndex) = m.FieldTable.FieldArray(7).NullValue Then
                m.FieldTable(7, lIndex) = ""
            End If
        Next lIndex
    End If
    
    If lNumItems = 34 Then
        lNumItems = 36
        bShowProxy = CommercialsProxyExists
        
        astrFields.Add Str(CLng(bShowProxy)) & ";" & "Genesis Proxy Setup Strength" & ";" & "" & ";" & Str(CLng(bShowProxy)) & ";" & "Genesis Sentiment" & ";" & "Proxy SETUP" & ";" & "Proxy Setup Strength = Commercials Proxy - Sentiment (as %)", 21
        astrFields.Add Str(CLng(bShowProxy)) & ";" & "Genesis Proxy Setup Strength Change" & ";" & "" & ";" & Str(CLng(bShowProxy)) & ";" & "Genesis Sentiment" & ";" & "Proxy Setup Chg", 22
        
        SetIniFileProperty "NumFields", lNumItems, "Fields", strIniFile
        For lIndex = 0 To astrFields.Size - 1
            SetIniFileProperty "Field" & lIndex, astrFields(lIndex), "Fields", strIniFile
            For lField = 0 To m.FieldTable.NumFields - 1
                m.FieldTable(lField, lIndex) = Parse(astrFields(lIndex), ";", lField + 1)
            Next lField
            If m.FieldTable(7, lIndex) = m.FieldTable.FieldArray(7).NullValue Then
                m.FieldTable(7, lIndex) = ""
            End If
        Next lIndex
    End If
    
    If lNumItems = 36 Then
        lNumItems = 38
        bShowProxy = CommercialsProxyExists
        
        astrFields.Add Str(CLng(bShowProxy)) & ";" & "LW Proxy Setup Strength" & ";" & "" & ";" & Str(CLng(bShowProxy)) & ";" & "TN Consensus" & ";" & "Proxy SETUP" & ";" & "Proxy Setup Strength = Commercials Proxy - Sentiment (as %)", 29
        astrFields.Add Str(CLng(bShowProxy)) & ";" & "LW Proxy Setup Strength Change" & ";" & "" & ";" & Str(CLng(bShowProxy)) & ";" & "TN Consensus" & ";" & "Proxy Setup Chg", 30
        
        SetIniFileProperty "NumFields", lNumItems, "Fields", strIniFile
        For lIndex = 0 To astrFields.Size - 1
            SetIniFileProperty "Field" & lIndex, astrFields(lIndex), "Fields", strIniFile
            For lField = 0 To m.FieldTable.NumFields - 1
                m.FieldTable(lField, lIndex) = Parse(astrFields(lIndex), ";", lField + 1)
            Next lField
            If m.FieldTable(7, lIndex) = m.FieldTable.FieldArray(7).NullValue Then
                m.FieldTable(7, lIndex) = ""
            End If
        Next lIndex
    End If
    
    If lNumItems = 38 Then
        lNumItems = 42
        bShowPremium = HasModule("LWMC") And CalendarSpreadExists
    
        astrFields.Add Str(CLng(bShowPremium)) & ";" & "Premium Spread One Contract Out" & ";" & "" & ";" & Str(CLng(bShowPremium)) & ";" & "Premium Spread" & ";" & "One Out" & ";" & "", 31
        astrFields.Add "0" & ";" & "Premium Spread One Contract Out Change" & ";" & "" & ";" & Str(CLng(bShowPremium)) & ";" & "Premium Spread" & ";" & "Change" & ";" & "", 32
        astrFields.Add Str(CLng(bShowPremium)) & ";" & "Premium Spread Two Contracts Out" & ";" & "" & ";" & Str(CLng(bShowPremium)) & ";" & "Premium Spread" & ";" & "Two Out" & ";" & "", 33
        astrFields.Add "0" & ";" & "Premium Spread Two Contracts Out Change" & ";" & "" & ";" & Str(CLng(bShowPremium)) & ";" & "Premium Spread" & ";" & "Change" & ";" & "", 34
        
        SetIniFileProperty "NumFields", lNumItems, "Fields", strIniFile
        For lIndex = 0 To astrFields.Size - 1
            SetIniFileProperty "Field" & lIndex, astrFields(lIndex), "Fields", strIniFile
            For lField = 0 To m.FieldTable.NumFields - 1
                m.FieldTable(lField, lIndex) = Parse(astrFields(lIndex), ";", lField + 1)
            Next lField
            If m.FieldTable(7, lIndex) = m.FieldTable.FieldArray(7).NullValue Then
                m.FieldTable(7, lIndex) = ""
            End If
        Next lIndex
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotSettings.FillTable", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookBack
'' Description: Determine the Larry Williams sentiment lookback
'' Inputs:      Value
'' Returns:     Lookback
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LookBack(ByVal strValue As String) As String
On Error GoTo ErrSection:

    If Val(strValue) > 1 Then
        LookBack = "1"
    Else
        LookBack = strValue
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotSettings.LookBack"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CommercialsProxyExists
'' Description: Does the CommercialsProxy function exist and is the user enabled?
'' Inputs:      None
'' Returns:     True if exists and enabled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CommercialsProxyExists() As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                "WHERE [FunctionName]='Commercials Proxy Index';", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        If HasModule(NullChk(rs!RequiredMod)) = True Then
            bReturn = True
        End If
    End If
    
    CommercialsProxyExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotSettings.CommercialsProxyExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalendarSpreadExists
'' Description: Does the Calendar Spread function exist?
'' Inputs:      None
'' Returns:     True if exists and enabled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CalendarSpreadExists() As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                "WHERE [FunctionName]='Calendar Spread';", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        If HasModule(NullChk(rs!RequiredMod)) = True Then
            bReturn = True
        End If
    End If
    
    CalendarSpreadExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotSettings.CalendarSpreadExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckPremiumSpreadRecord
'' Description: Check to see if enablement has changed since last time
'' Inputs:      Record, Force a show if now turned on?
'' Returns:     True if Changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckPremiumSpreadRecord(ByVal lRecord As Long, Optional ByVal bForceShow As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If HasModule("LWMC") And CalendarSpreadExists Then
        If m.FieldTable(3, lRecord) = "0" Then
            If bForceShow Then
                m.FieldTable(0, lRecord) = "1"
            End If
            bReturn = True
        End If
        m.FieldTable(3, lRecord) = "1"
    Else
        If m.FieldTable(3, lRecord) = "1" Then
            m.FieldTable(0, lRecord) = "0"
            bReturn = True
        End If
        m.FieldTable(3, lRecord) = "0"
    End If
    
    CheckPremiumSpreadRecord = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotSettings.CheckPremiumSpreadRecord"
    
End Function

