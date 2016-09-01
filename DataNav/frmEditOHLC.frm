VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmEditOHLC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Daily Data"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
      Height          =   435
      Left            =   1860
      TabIndex        =   1
      Top             =   2880
      Width           =   1395
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
      Caption         =   "frmEditOHLC.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmEditOHLC.frx":0036
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmEditOHLC.frx":0056
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL Frame4 
      Height          =   855
      Left            =   180
      TabIndex        =   4
      Top             =   0
      Width           =   4395
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmEditOHLC.frx":0072
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditOHLC.frx":0092
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditOHLC.frx":00B2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optTU 
         Height          =   220
         Left            =   1560
         TabIndex        =   2
         Top             =   555
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "frmEditOHLC.frx":00CE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditOHLC.frx":0108
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0128
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optDecimal 
         Height          =   220
         Left            =   2880
         TabIndex        =   3
         Top             =   555
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "frmEditOHLC.frx":0144
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditOHLC.frx":0172
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0192
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Left            =   480
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
         Caption         =   "frmEditOHLC.frx":01AE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":01EA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":020A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSymbol 
         Height          =   195
         Left            =   960
         Top             =   240
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
         Caption         =   "frmEditOHLC.frx":0226
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":0252
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0272
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDate 
         Height          =   195
         Left            =   3060
         Top             =   240
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
         Caption         =   "frmEditOHLC.frx":028E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":02C2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":02E2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   195
         Left            =   240
         Top             =   240
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
         Caption         =   "frmEditOHLC.frx":02FE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":032C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":034C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   195
         Left            =   2460
         Top             =   240
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
         Caption         =   "frmEditOHLC.frx":0368
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":0392
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":03B2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame3 
      Height          =   1695
      Left            =   2220
      TabIndex        =   5
      Top             =   1080
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
      Caption         =   "frmEditOHLC.frx":03CE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditOHLC.frx":03EE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditOHLC.frx":040E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtContOI 
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   1260
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditOHLC.frx":042A
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
         Tip             =   "frmEditOHLC.frx":044A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":046A
      End
      Begin HexUniControls.ctlUniTextBoxXP txtContVol 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   840
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditOHLC.frx":0486
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
         Tip             =   "frmEditOHLC.frx":04A6
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":04C6
      End
      Begin HexUniControls.ctlUniTextBoxXP txtOI 
         Height          =   315
         Left            =   1140
         TabIndex        =   10
         Top             =   420
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditOHLC.frx":04E2
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
         Tip             =   "frmEditOHLC.frx":0502
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0522
      End
      Begin HexUniControls.ctlUniTextBoxXP txtVol 
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditOHLC.frx":053E
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
         Tip             =   "frmEditOHLC.frx":055E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":057E
      End
      Begin HexUniControls.ctlUniLabelXP lblContOI 
         Height          =   255
         Left            =   20
         Top             =   1320
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
         Caption         =   "frmEditOHLC.frx":059A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":05D6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":05F6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblContVol 
         Height          =   255
         Left            =   20
         Top             =   900
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
         Caption         =   "frmEditOHLC.frx":0612
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":064E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":066E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOI 
         Height          =   195
         Left            =   20
         Top             =   480
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
         Caption         =   "frmEditOHLC.frx":068A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":06C6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":06E6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblVol 
         Height          =   195
         Left            =   0
         Top             =   60
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
         Caption         =   "frmEditOHLC.frx":0702
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":0730
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0750
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   3480
      TabIndex        =   0
      Top             =   2880
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
      Caption         =   "frmEditOHLC.frx":076C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmEditOHLC.frx":079A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmEditOHLC.frx":07BA
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdSave 
      Height          =   435
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   1395
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
      Caption         =   "frmEditOHLC.frx":07D6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmEditOHLC.frx":0810
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmEditOHLC.frx":0830
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblMessage 
      Height          =   465
      Left            =   300
      Top             =   3420
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
      Caption         =   "frmEditOHLC.frx":084C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmEditOHLC.frx":091E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditOHLC.frx":093E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   1080
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
      Caption         =   "frmEditOHLC.frx":095A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditOHLC.frx":097A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditOHLC.frx":099A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP lblOpen 
         Height          =   195
         Left            =   20
         Top             =   60
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
         Caption         =   "frmEditOHLC.frx":09B6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":09E0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0A00
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblHigh 
         Height          =   195
         Left            =   20
         Top             =   480
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
         Caption         =   "frmEditOHLC.frx":0A1C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":0A46
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0A66
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblLow 
         Height          =   195
         Left            =   20
         Top             =   900
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
         Caption         =   "frmEditOHLC.frx":0A82
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":0AAA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0ACA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblClose 
         Height          =   195
         Left            =   10
         Top             =   1320
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
         Caption         =   "frmEditOHLC.frx":0AE6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditOHLC.frx":0B12
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0B32
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtOpen 
         Height          =   315
         Left            =   660
         TabIndex        =   15
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditOHLC.frx":0B4E
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
         Tip             =   "frmEditOHLC.frx":0B6E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0B8E
      End
      Begin HexUniControls.ctlUniTextBoxXP txtHigh 
         Height          =   315
         Left            =   660
         TabIndex        =   14
         Top             =   420
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditOHLC.frx":0BAA
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
         Tip             =   "frmEditOHLC.frx":0BCA
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0BEA
      End
      Begin HexUniControls.ctlUniTextBoxXP txtLow 
         Height          =   315
         Left            =   660
         TabIndex        =   13
         Top             =   840
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditOHLC.frx":0C06
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
         Tip             =   "frmEditOHLC.frx":0C26
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0C46
      End
      Begin HexUniControls.ctlUniTextBoxXP txtClose 
         Height          =   315
         Left            =   660
         TabIndex        =   12
         Top             =   1260
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditOHLC.frx":0C62
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
         Tip             =   "frmEditOHLC.frx":0C82
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditOHLC.frx":0CA2
      End
   End
End
Attribute VB_Name = "frmEditOHLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    nSymbolID As Long
    lDate As Long
    strSecType As String
    strSymbol As String
    Bars As cGdBars
End Type
Private m As mPrivate

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    If InfBox("Are you sure you wish to delete this bar?", "?", "Delete|+-Cancel", "Confirm Delete") = "D" Then
        UpdateData True
        Unload Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.cmdDelete.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrSection:

    Dim bDirty As Boolean
    Dim i&, strPrice$, strErr$
    Dim dValue#, dMinMove#, dTickDiv#, dTicks#
    
    ' make sure LostFocus triggers for text boxes
    MoveFocus cmdSave
    DoEvents
    
    ' check for errors
    If ValOfText(txtLow.Tag) > ValOfText(txtHigh.Tag) Then
        strErr = "Low cannot be greater than High"
    ElseIf ValOfText(txtOpen.Tag) > ValOfText(txtHigh.Tag) Then
        strErr = "Open cannot be greater than High"
    ElseIf ValOfText(txtClose.Tag) > ValOfText(txtHigh.Tag) Then
        strErr = "Close cannot be greater than High"
    ElseIf ValOfText(txtOpen.Tag) < ValOfText(txtLow.Tag) Then
        strErr = "Open cannot be less than Low"
    ElseIf ValOfText(txtClose.Tag) < ValOfText(txtLow.Tag) Then
        strErr = "Close cannot be less than Low"
    End If
    If Len(strErr) > 0 Then
        InfBox strErr, "e", , "Error in Prices"
        Exit Sub
    End If
    
    If txtOpen.Enabled Then
        dValue = ValOfText(txtOpen.Tag)
        If dValue <> m.Bars(eBARS_Open, 0) Then
            m.Bars(eBARS_Open, 0) = dValue
            bDirty = True
        End If
    End If
    
    If txtHigh.Enabled Then
        dValue = ValOfText(txtHigh.Tag)
        If dValue <> m.Bars(eBARS_High, 0) Then
            m.Bars(eBARS_High, 0) = dValue
            bDirty = True
        End If
    End If
    
    If txtLow.Enabled Then
        dValue = ValOfText(txtLow.Tag)
        If dValue <> m.Bars(eBARS_Low, 0) Then
            m.Bars(eBARS_Low, 0) = dValue
            bDirty = True
        End If
    End If
    
    If txtClose.Enabled Then
        dValue = ValOfText(txtClose.Tag)
        If dValue <> m.Bars(eBARS_Close, 0) Then
            m.Bars(eBARS_Close, 0) = dValue
            bDirty = True
        End If
    End If
    
    If txtVol.Enabled Then
        dValue = ValOfText(txtVol.Tag)
        If dValue <> m.Bars(eBARS_Vol, 0) Then
            m.Bars(eBARS_Vol, 0) = dValue
            bDirty = True
        End If
    End If
    
    If txtOI.Enabled Then
        dValue = ValOfText(txtOI.Tag)
        If dValue <> m.Bars(eBARS_OI, 0) Then
            m.Bars(eBARS_OI, 0) = dValue
            bDirty = True
        End If
    End If
    
    If txtContVol.Enabled Then
        dValue = ValOfText(txtContVol.Tag)
        If dValue <> m.Bars(eBARS_ContVol, 0) Then
            m.Bars(eBARS_ContVol, 0) = dValue
            bDirty = True
        End If
    ElseIf m.Bars(eBARS_ContVol, 0) <> kNullData Then
        m.Bars(eBARS_ContVol, 0) = m.Bars(eBARS_Vol, 0)
    End If
    
    If txtContOI.Enabled Then
        dValue = ValOfText(txtContOI.Tag)
        If dValue <> m.Bars(eBARS_ContOI, 0) Then
            m.Bars(eBARS_ContOI, 0) = dValue
            bDirty = True
        End If
    End If
    
    If bDirty Then UpdateData

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.cmdSave.Click", eGDRaiseError_Show
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
    RaiseError "frmEditOHLC.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Me.Icon = Picture16("kBlank")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Sub ShowMe(ByVal nSymbolID&, ByVal dDate#, ByVal bFromContinuous As Boolean)
On Error GoTo ErrSection:

    ' set symbol and date info
    m.nSymbolID = nSymbolID
    m.strSymbol = SU_GetSymbol(nSymbolID)
    lblSymbol = m.strSymbol
    m.lDate = Int(dDate)
    lblDate = DateFormat(m.lDate)
    
    ' get data
    Set m.Bars = New cGdBars
    m.Bars.ArrayMask = eBARS_Eod
    DM_GetBars m.Bars, nSymbolID, , m.lDate, m.lDate, , False, True, True
    If m.Bars(eBARS_DateTime, 0) = m.lDate Then
        ' store values in tags
        txtOpen.Tag = CStr(m.Bars(eBARS_Open, 0))
        txtHigh.Tag = CStr(m.Bars(eBARS_High, 0))
        txtLow.Tag = CStr(m.Bars(eBARS_Low, 0))
        txtClose.Tag = CStr(m.Bars(eBARS_Close, 0))
        txtVol.Tag = CStr(m.Bars(eBARS_Vol, 0))
        txtOI.Tag = CStr(m.Bars(eBARS_OI, 0))
        txtContVol.Tag = CStr(m.Bars(eBARS_ContVol, 0))
        txtContOI.Tag = CStr(m.Bars(eBARS_ContOI, 0))
        
        ' disable unused controls
        m.strSecType = UCase(Chr(m.Bars.Prop(eBARS_SecurityType)))
        If m.strSecType <> "F" Then
            optDecimal = True
            optTU.Enabled = False
            lblOI.Enabled = False
            txtOI.Enabled = False
            lblContVol.Enabled = False
            txtContVol.Enabled = False
            lblContOI.Enabled = False
            txtContOI.Enabled = False
        End If
        Select Case m.strSecType
        Case "F"
            If Not bFromContinuous Then
                lblMessage = ""
            End If
        Case "S"
            lblMessage = "NOTE: data shown is unadjusted for splits."
            Me.Height = Me.Height - lblMessage.Height / 3
        Case Else
            lblMessage = ""
        End Select
        If Len(lblMessage) = 0 Then
            Me.Height = Me.Height - lblMessage.Height
        End If
        
        ' show form
        ShowData
        ShowForm Me, True
    Else
        Beep
    End If
    
    Set m.Bars = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.ShowMe", eGDRaiseError_Raise
    
End Sub

Private Sub ShowData()
On Error GoTo ErrSection:

    Dim dValue#

    dValue = ValOfText(txtOpen.Tag)
    If dValue <> kNullData And txtOpen.Enabled Then
        txtOpen = m.Bars.PriceDisplay(dValue, optTU)
    End If
        
    dValue = ValOfText(txtHigh.Tag)
    If dValue <> kNullData And txtHigh.Enabled Then
        txtHigh = m.Bars.PriceDisplay(dValue, optTU)
    End If
    
    dValue = ValOfText(txtLow.Tag)
    If dValue <> kNullData And txtLow.Enabled Then
        txtLow = m.Bars.PriceDisplay(dValue, optTU)
    End If
    
    dValue = ValOfText(txtClose.Tag)
    If dValue <> kNullData And txtClose.Enabled Then
        txtClose = m.Bars.PriceDisplay(dValue, optTU)
    End If
    
    dValue = ValOfText(txtVol.Tag)
    If dValue <> kNullData And txtVol.Enabled Then
        txtVol = CStr(dValue)
    End If

    dValue = ValOfText(txtOI.Tag)
    If dValue <> kNullData And txtOI.Enabled Then
        txtOI = CStr(dValue)
    End If

    dValue = ValOfText(txtContVol.Tag)
    If dValue <> kNullData And txtContVol.Enabled Then
        txtContVol = CStr(dValue)
    End If

    dValue = ValOfText(txtContOI.Tag)
    If dValue <> kNullData And txtContOI.Enabled Then
        txtContOI = CStr(dValue)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.ShowData", eGDRaiseError_Raise
    
End Sub

Private Sub optDecimal_Click()
On Error GoTo ErrSection:

    If Me.Visible Then ShowData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.optDecimal.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optTU_Click()
On Error GoTo ErrSection:

    If Me.Visible Then ShowData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.optTU.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub FixPrice(txt As ctlUniTextBoxXP) 'RH was TextBox
On Error GoTo ErrSection:

    Dim strPrice$
    Dim dPrice#, dMinMove#, dTickMult#, dTicks#
    
    ' get minimum move
    dMinMove = m.Bars.MinMove(m.lDate)
    
    If optTU And InStr(txt, "^") > 0 Then
        ' need to convert price from trading units,
        ' so find the tick multiplier (for post-caret)
        dTickMult = 0
        strPrice = m.Bars.PriceDisplay(dMinMove, True)
        dTicks = ValOfText(Parse(strPrice, "^", 2))
        If dTicks > 0 Then
            dTickMult = dMinMove / dTicks
        End If
    End If
    
    ' get price from text
    strPrice = Trim(txt)
    dPrice = ValOfText(Parse(strPrice, "^", 1))
    If dTickMult > 0 Then
        ' add/subtract post-caret
        dTicks = ValOfText(Parse(strPrice, "^", 2)) * dTickMult
        If Left(strPrice, 1) = "-" Then
            dPrice = dPrice - dTicks
        Else
            dPrice = dPrice + dTicks
        End If
    End If
    
    ' round to nearest min move
    'dPrice = Int(dPrice / dMinMove + 0.5) * dMinMove
    dPrice = RoundToMinMove(dPrice, dMinMove)
    
    ' store decimal in Tag
    txt.Tag = CStr(dPrice)
    
    ' and redisplay
    ShowData
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.FixPrice", eGDRaiseError_Raise
    
End Sub

Private Sub FixVol(txt As ctlUniTextBoxXP) 'RH was TextBox
On Error GoTo ErrSection:

    Dim dValue#
    
    ' store back in Tag
    On Error Resume Next
    dValue = Int(ValOfText(txt))
    txt.Tag = CStr(dValue)
    
    ' and redisplay
    ShowData
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.FixVol", eGDRaiseError_Raise
    
End Sub

Private Sub txtClose_LostFocus()
On Error GoTo ErrSection:

    FixPrice txtClose

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.txtClose.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtHigh_LostFocus()
On Error GoTo ErrSection:

    FixPrice txtHigh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.txtHigh.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtLow_LostFocus()
On Error GoTo ErrSection:

    FixPrice txtLow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.txtLow.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtOpen_LostFocus()
On Error GoTo ErrSection:

    FixPrice txtOpen

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.txtOpen.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtVol_LostFocus()
On Error GoTo ErrSection:

    FixVol txtVol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.txtVol.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtOI_LostFocus()
On Error GoTo ErrSection:

    FixVol txtOI

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.txtOI.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtContVol_LostFocus()
On Error GoTo ErrSection:

    FixVol txtContVol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.txtContVol.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtContOI_LostFocus()
On Error GoTo ErrSection:

    FixVol txtContOI

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.txtContOI.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub UpdateData(Optional bDeleteBar As Boolean = False)
On Error GoTo ErrSection:

    Dim lChkDate&, bSuccess As Boolean
    Dim strSymbol$, strSymbols$, i&
    Dim aSymbols As New cGdArray

    aSymbols.Create eGDARRAY_Longs

    ' save the changed data (that's now in m.Bars)
    If bDeleteBar Then
        ' delete data for m.lDate for m.nSymbolID
        bSuccess = DM_ClearEOD(g.DMS, m.nSymbolID, m.lDate, aSymbols.ArrayHandle)
    Else
        ' update data for m.nSymbolID for m.lDate using m.Bars
        bSuccess = DM_ChangeEOD(g.DMS, m.nSymbolID, m.Bars.BarsHandle, aSymbols.ArrayHandle)
    End If
    
    If Not bSuccess Then
        InfBox "Data changes could not be saved.", "e", , "ERROR"
    Else
        ' refresh things
        UpdateVisibleCharts
        
        ' if date was on or after the day before the LastDailyDownload,
        ' then refresh quote board (since prev close could have changed)
        lChkDate = LastDailyDownload - 1
        Do While Not IsWeekday(lChkDate)
            lChkDate = lChkDate - 1
        Loop
        If m.lDate >= lChkDate Then
            frmQuotes.TotalRefresh True
        End If
        
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        strSymbols = ""
        For i = 0 To aSymbols.Size - 1
            ' if symbol is in the pool, then recalc criteria for that symbol
            strSymbol = g.SymbolPool.SymbolForID(aSymbols(i))
            If Len(strSymbol) > 0 Then
                strSymbols = strSymbols & vbTab & strSymbol
            End If
        Next
        If Len(strSymbols) > 1 Then
            g.SymbolPool.RecalcDirtyCriteria False, Mid(strSymbols, 2)
        End If
        Screen.MousePointer = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditOHLC.UpdateData", eGDRaiseError_Raise
    
End Sub

