VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOptionCalc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options Calculator"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "frmOptionCalc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HexUniControls.ctlUniTextBoxXP txtStrikePrice 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   1200
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOptionCalc.frx":030A
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
      Tip             =   "frmOptionCalc.frx":0332
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0352
   End
   Begin HexUniControls.ctlUniTextBoxXP txtIntRate 
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Top             =   840
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOptionCalc.frx":036E
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
      Tip             =   "frmOptionCalc.frx":0396
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":03B6
   End
   Begin HexUniControls.ctlUniTextBoxXP txtAssetPrice 
      Height          =   285
      Left            =   3840
      TabIndex        =   7
      Top             =   480
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOptionCalc.frx":03D2
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
      Tip             =   "frmOptionCalc.frx":03FA
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":041A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1860
      TabIndex        =   8
      Top             =   5580
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
      Caption         =   "frmOptionCalc.frx":0436
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmOptionCalc.frx":0462
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0482
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP ckbAdvEdit 
      Height          =   255
      Left            =   2790
      TabIndex        =   9
      Top             =   1680
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
      Caption         =   "frmOptionCalc.frx":049E
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmOptionCalc.frx":04DE
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":04FE
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraGreeks 
      Height          =   1815
      Left            =   150
      TabIndex        =   5
      Top             =   3660
      Width           =   4575
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptionCalc.frx":051A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptionCalc.frx":0546
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0566
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP lblDelta 
         Height          =   285
         Left            =   930
         Top             =   240
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
         Caption         =   "frmOptionCalc.frx":0582
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":05AC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":05CC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTheta 
         Height          =   285
         Left            =   930
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
         Caption         =   "frmOptionCalc.frx":05E8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":0612
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0632
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblRho 
         Height          =   285
         Left            =   930
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
         Caption         =   "frmOptionCalc.frx":064E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":0674
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0694
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblGamma 
         Height          =   285
         Left            =   930
         Top             =   540
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
         Caption         =   "frmOptionCalc.frx":06B0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":06DA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":06FA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblVega 
         Height          =   285
         Left            =   930
         Top             =   840
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
         Caption         =   "frmOptionCalc.frx":0716
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":073E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":075E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDeltaValue 
         Height          =   285
         Left            =   1920
         Top             =   240
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
         Caption         =   "frmOptionCalc.frx":077A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":07A2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":07C2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblThetaValue 
         Height          =   285
         Left            =   1920
         Top             =   1140
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
         Caption         =   "frmOptionCalc.frx":07DE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":0806
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0826
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblRhoValue 
         Height          =   285
         Left            =   1920
         Top             =   1440
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
         Caption         =   "frmOptionCalc.frx":0842
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":086A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":088A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblVegaValue 
         Height          =   285
         Left            =   1920
         Top             =   840
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
         Caption         =   "frmOptionCalc.frx":08A6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":08CE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":08EE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblGammaValue 
         Height          =   285
         Left            =   1920
         Top             =   540
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
         Caption         =   "frmOptionCalc.frx":090A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":0932
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0952
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraTest 
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4575
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptionCalc.frx":096E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptionCalc.frx":09E0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0A00
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtVolatile 
         Height          =   315
         Left            =   2220
         TabIndex        =   10
         Top             =   900
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmOptionCalc.frx":0A1C
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
         Tip             =   "frmOptionCalc.frx":0A44
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0A64
      End
      Begin HexUniControls.ctlUniTextBoxXP txtOptPrice 
         Height          =   315
         Left            =   2220
         TabIndex        =   11
         Top             =   540
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmOptionCalc.frx":0A80
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
         Tip             =   "frmOptionCalc.frx":0AA8
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0AC8
      End
      Begin HexUniControls.ctlUniComboImageXP cboOptPrice 
         Height          =   315
         Left            =   480
         TabIndex        =   12
         Top             =   540
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
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
         Tip             =   "frmOptionCalc.frx":0AE4
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0B04
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCalcVolatile 
         Height          =   255
         Left            =   200
         TabIndex        =   3
         Top             =   960
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
         Caption         =   "frmOptionCalc.frx":0B20
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptionCalc.frx":0B64
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0B84
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCalcOptPrice 
         Height          =   435
         Left            =   200
         TabIndex        =   2
         Top             =   480
         Width           =   220
         _ExtentX        =   397
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptionCalc.frx":0BA0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptionCalc.frx":0BDE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0BFE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCalc 
         Default         =   -1  'True
         Height          =   720
         Left            =   3240
         TabIndex        =   1
         Top             =   540
         Width           =   1095
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
         Caption         =   "frmOptionCalc.frx":0C1A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptionCalc.frx":0C4E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0C6E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOptPrice 
         Height          =   255
         Left            =   480
         Top             =   315
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
         Caption         =   "frmOptionCalc.frx":0C8A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptionCalc.frx":0CD0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptionCalc.frx":0CF0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblIntRate 
      Height          =   255
      Left            =   2820
      Top             =   900
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
      Caption         =   "frmOptionCalc.frx":0D0C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":0D46
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0D66
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblStrikePrice 
      Height          =   255
      Left            =   2790
      Top             =   1260
      Width           =   885
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptionCalc.frx":0D82
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":0DBA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0DDA
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAssetSym 
      Height          =   255
      Left            =   2790
      Top             =   135
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
      Caption         =   "frmOptionCalc.frx":0DF6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":0E2E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0E4E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAssetSymValue 
      Height          =   255
      Left            =   3870
      Top             =   135
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
      Caption         =   "frmOptionCalc.frx":0E6A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":0E90
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0EB0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAssetPrice 
      Height          =   255
      Left            =   2820
      Top             =   525
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
      Caption         =   "frmOptionCalc.frx":0ECC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":0F02
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0F22
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblOptSym 
      Height          =   255
      Left            =   150
      Top             =   135
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
      Caption         =   "frmOptionCalc.frx":0F3E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":0F78
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":0F98
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblOptSymValue 
      Height          =   255
      Left            =   1350
      Top             =   135
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
      Caption         =   "frmOptionCalc.frx":0FB4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":0FE8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":1008
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblOptTypeValue 
      Height          =   255
      Left            =   1590
      Top             =   1680
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
      Caption         =   "frmOptionCalc.frx":1024
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":104C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":106C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblTodayValue 
      Height          =   255
      Left            =   1560
      Top             =   930
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
      Caption         =   "frmOptionCalc.frx":1088
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":10BC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":10DC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblConDateValue 
      Height          =   255
      Left            =   1560
      Top             =   530
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
      Caption         =   "frmOptionCalc.frx":10F8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":1126
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":1146
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblDaysLeftValue 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1590
      Top             =   1305
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
      Caption         =   "frmOptionCalc.frx":1162
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":118C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":11AC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblOptType 
      Height          =   255
      Left            =   150
      Top             =   1680
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
      Caption         =   "frmOptionCalc.frx":11C8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":11FE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":121E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblConDate 
      Height          =   255
      Left            =   150
      Top             =   530
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
      Caption         =   "frmOptionCalc.frx":123A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":1274
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":1294
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblToday 
      Height          =   255
      Left            =   150
      Top             =   930
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
      Caption         =   "frmOptionCalc.frx":12B0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":12DA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":12FA
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblDaysLeft 
      Height          =   255
      Left            =   150
      Top             =   1320
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
      Caption         =   "frmOptionCalc.frx":1316
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptionCalc.frx":135A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptionCalc.frx":137A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmOptionCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOptionsCalc.frm
'' Description: Allow the user to perform options calculations
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
     
Private Type mPrivate
    strOptType As String
    dExpDays As Double
    dAssetPrice As Double
    dIntRate As Double
    dStrikePrice As Double
    dOptionAsk As Double
    dOptionBid As Double
    dOptionLast As Double
    dOptionPrice As Double
    dVolatile As Double
    dDelta As Double
    dTheta As Double
    dRho As Double
    dGama As Double
    dVega As Double
    nAdvEdit As Integer
    nOptPriceCtlColor As Long
    nVolCtlColor As Long
    bCalcVolatile As Boolean
    bCalcOptPrice As Boolean
    bClearCalcItem As Boolean
    bChanged As Boolean
    bFirstShown As Boolean
    bIsFuture As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Symbol, Option Symbol, Today's Price, Month, Year, Underlying
''              Price, Strike Price, Ask, Bid, Last, Is it a Put?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(strAssetSym$, strOptSym$, ByVal dToday#, ByVal nMonth&, _
    ByVal nYear&, ByVal dAssetPrice#, ByVal dStrikePrice#, _
    ByVal dOptAsk#, ByVal dOptBid#, ByVal dOptLast#, ByVal bIsPut As Byte)
On Error GoTo ErrSection:

    Dim Bars As New cGdBars
    Dim strToday$, strConDate$
    Dim lExpDate As Long
            
    'get interest rate
    DM_GetBars Bars, "$IRX", 0, LastDailyDownload - 5
    m.dIntRate = Bars(eBARS_Close, Bars.Size - 1)
    If m.dIntRate < 0 Then
        m.dIntRate = 0
    End If
    
    'calculate number of days left till contract expiration
    'note1: Only stock options expire on the 3rd Friday of the month
    '       code will have to be added for futures etc.
    'note2: The day value passed in sometimes contain a fractional part
    '       representing time. Truncating the fractional part ensures
    '       that the calculation for days left does not 'decrease'
    '       the days when the time is past noon.
    If InStr(strOptSym, "-") > 0 Then
        m.bIsFuture = True
        If SU_GetFutureOptionExp(strOptSym, lExpDate) Then
            m.dExpDays = lExpDate - Int(dToday) + 1
        End If
    Else
        m.bIsFuture = False
        m.dExpDays = GetDateFromRule(nYear, nMonth, "3F") - Int(dToday) + 1
    End If
    strToday = DateFormat(dToday, 0)
    strConDate = nMonth & "/" & nYear
        
    m.dStrikePrice = dStrikePrice
    m.dOptionAsk = dOptAsk
    m.dOptionBid = dOptBid
    m.dOptionLast = dOptLast
    m.dOptionPrice = dOptAsk
    m.dAssetPrice = dAssetPrice
    m.dIntRate = RoundNum(m.dIntRate, 2)
    m.nAdvEdit = 0
    m.bCalcVolatile = True
    m.bCalcOptPrice = False
    
    If bIsPut Then
        m.strOptType = "PUT"
    Else
        m.strOptType = "CALL"
    End If
            
    'set default option price
    Me.cboOptPrice.ListIndex = 0
    SetTxtOptionPrice
    'set non-variable controls
    Me.lblOptSymValue.Caption = strOptSym
    Me.lblAssetSymValue.Caption = strAssetSym
    Me.lblConDateValue.Caption = strConDate
    Me.lblTodayValue.Caption = strToday
    Me.lblDaysLeftValue.Caption = m.dExpDays
    Me.lblOptTypeValue.Caption = m.strOptType
    'set variable controls
    SetAdvEditControls
    SetCalcControls
    'calculate volatility & greeks
    m.bChanged = True
    cmdCalc_Click
    m.bFirstShown = True
    
    ShowForm Me, eForm_Modal

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmOptionCalc.ShowMe", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearCalcItem
'' Description: Clear the appropriate item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearCalcItem()
On Error GoTo ErrSection:

    If m.bClearCalcItem = True Then
        If m.bCalcOptPrice = True Then
            txtOptPrice.Text = ""
        Else
            txtVolatile.Text = ""
        End If
        m.bChanged = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.ClearCalcItem", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidateInput
'' Description: Validate the user's inputs
'' Inputs:      None
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateInput() As Boolean
On Error GoTo ErrSection:

    If txtAssetPrice.Text = "" Then
        InfBox ("Asset price cannot be blank.")
        ValidateInput = False
        Exit Function
    End If

    If txtIntRate.Text = "" Then
        InfBox ("Interest rate cannot be blank.")
        ValidateInput = False
        Exit Function
    End If

    If txtStrikePrice.Text = "" Then
        InfBox ("Strike price cannot be blank.")
        ValidateInput = False
        Exit Function
    End If

    If m.bCalcOptPrice Then
        If txtVolatile.Text = "" Then
            InfBox ("Volatility cannot be blank.")
            ValidateInput = False
            Exit Function
        End If
        m.dVolatile = ValOfText(txtVolatile.Text) / 100#
    Else
        If txtOptPrice.Text = "" Then
            InfBox ("Option price cannot be blank.")
            ValidateInput = False
            Exit Function
        End If
        m.dOptionPrice = ValOfText(txtOptPrice.Text)
    End If

    m.dAssetPrice = ValOfText(txtAssetPrice.Text)
    m.dIntRate = ValOfText(txtIntRate.Text)
    m.dStrikePrice = ValOfText(txtStrikePrice.Text)
    
    ValidateInput = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptionCalc.ValidateInput", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCalcControls
'' Description: Set the calculator controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCalcControls()
On Error GoTo ErrSection:
    
    'determine controls color
    If m.bCalcOptPrice = True Then
        m.nOptPriceCtlColor = Me.BackColor
        m.bCalcVolatile = False
        m.nVolCtlColor = RGB(255, 255, 255)
    Else
        m.nVolCtlColor = Me.BackColor
        m.bCalcVolatile = True
        m.nOptPriceCtlColor = RGB(255, 255, 255)
    End If
        
    'set volatility & option price controls
    Me.txtOptPrice.Text = Format(m.dOptionPrice, "0.000")
    Me.txtVolatile.Text = Format(m.dVolatile * 100#, "0.00")
    
    Me.txtOptPrice.BackColor = m.nOptPriceCtlColor
    Me.txtVolatile.BackColor = m.nVolCtlColor
    
    Me.txtOptPrice.Locked = m.bCalcOptPrice
    Me.optCalcOptPrice.Value = m.bCalcOptPrice
    Me.txtVolatile.Locked = m.bCalcVolatile
    Me.optCalcVolatile.Value = m.bCalcVolatile
    
    'set greeks controls
    Me.lblDeltaValue.Caption = RoundNum(m.dDelta, 4)
    Me.lblThetaValue.Caption = RoundNum(m.dTheta / 365#, 4)
    Me.lblRhoValue.Caption = RoundNum(m.dRho / 100#, 4)
    Me.lblGammaValue.Caption = RoundNum(m.dGama, 4)
    Me.lblVegaValue.Caption = RoundNum(m.dVega / 100#, 4)

    'option price drop-down box
    '0=bid/ask avg, 1=bid, 2=ask, 3=last, 4=user input/calculated price
    If m.bCalcOptPrice = True Then
        Me.lblOptPrice.Visible = False
        Me.cboOptPrice.Enabled = False
        MoveFocus Me.txtVolatile
        SelectAll Me.txtVolatile
    Else
        Me.lblOptPrice.Visible = True
        Me.cboOptPrice.Enabled = True
        MoveFocus Me.txtOptPrice
        SelectAll Me.txtOptPrice
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.SetCalcControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAdvEditControls
'' Description: Set the advanced editor controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAdvEditControls()
On Error GoTo ErrSection:

    Dim editCtlColor&
    Dim lockCtl As Boolean
    
    If m.nAdvEdit = 0 Then
        editCtlColor = Me.BackColor
        lockCtl = True
    Else
        editCtlColor = RGB(255, 255, 255)       'white
        lockCtl = False
    End If
    
    'set advanced editing controls
    Me.txtAssetPrice.Text = Format(m.dAssetPrice, "0.00")
    Me.txtIntRate.Text = Format(m.dIntRate, "0.00")
    Me.txtStrikePrice.Text = Format(m.dStrikePrice, "0.00")
    
    Me.txtAssetPrice.BackColor = editCtlColor
    Me.txtIntRate.BackColor = editCtlColor
    Me.txtStrikePrice.BackColor = editCtlColor
    
    Me.txtAssetPrice.Locked = lockCtl
    Me.txtIntRate.Locked = lockCtl
    Me.txtStrikePrice.Locked = lockCtl

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.SetAdvEditControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ckbAdvEdit_Click
'' Description: Set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ckbAdvEdit_Click()
On Error GoTo ErrSection:

    m.nAdvEdit = ckbAdvEdit.Value
    SetAdvEditControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.ckbAdvEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show the help if the user presses F1
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
    RaiseError "frmOptionCalc.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, hide the form
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAssetPrice_Change
'' Description: Set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAssetPrice_Change()
On Error GoTo ErrSection:

    ClearCalcItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.txtAssetPrice.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optCalcOptPrice_Click
'' Description: Set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optCalcOptPrice_Click()
On Error GoTo ErrSection:

    m.bCalcOptPrice = Me.optCalcOptPrice.Value
    m.bCalcVolatile = Me.optCalcVolatile.Value / 100#
    SetCalcControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.optCalcOptPrice.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optCalcVolatile_Click
'' Description: Set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optCalcVolatile_Click()
On Error GoTo ErrSection:

    m.bCalcOptPrice = Me.optCalcOptPrice.Value
    m.bCalcVolatile = Me.optCalcVolatile.Value / 100#
    SetCalcControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.optCalcVolatile.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboOptPrice_Click
'' Description: Set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboOptPrice_Click()
On Error GoTo ErrSection:

    If Me.cboOptPrice.ListIndex = 4 Then
         Me.txtOptPrice.Text = ""
         MoveFocus Me.txtOptPrice
         Exit Sub
    End If
    SetTxtOptionPrice

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.cboOptPrice.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCalc_Click
'' Description: Perform the calculations
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCalc_Click()
On Error GoTo ErrSection:

    Dim bIsPut As Byte
    Dim dIntRatePct#
    Dim dCarryCost#
    
    'It is important to not "re-calculate" if user has not changed any values
    'because displayed values are truncated to 2 or 4 decimal places.
    If m.bChanged = False Then
        Exit Sub
    End If
   
    If ValidateInput = False Then
        Exit Sub
    End If
    
    If m.strOptType = "PUT" Then
        bIsPut = 1
    Else
        bIsPut = 0
    End If
    
    dIntRatePct = m.dIntRate / 100
    If m.bIsFuture = True Then
        dCarryCost = 0#
    Else
        dCarryCost = dIntRatePct
    End If
        
    If m.bCalcOptPrice = True Then
        Me.cboOptPrice.ListIndex = 4
        m.dOptionPrice = Opt_BlackSholes(m.dAssetPrice, m.dStrikePrice, _
            m.dExpDays, m.dVolatile, dIntRatePct, dCarryCost, bIsPut, m.dDelta)
    Else
        GetTxtOptionPrice
        m.dVolatile = Opt_GetVolatility(m.dOptionPrice, m.dAssetPrice, _
            m.dStrikePrice, m.dExpDays, dIntRatePct, dCarryCost, bIsPut, m.dDelta)
    End If
    
    m.dTheta = Opt_Theta(m.dAssetPrice, m.dStrikePrice, m.dExpDays, _
        m.dVolatile, dIntRatePct, dCarryCost, bIsPut)
    m.dRho = Opt_Rho(m.dAssetPrice, m.dStrikePrice, m.dExpDays, _
        m.dVolatile, dIntRatePct, dCarryCost, bIsPut)
    m.dGama = Opt_Gamma(m.dAssetPrice, m.dStrikePrice, m.dExpDays, _
        m.dVolatile, dIntRatePct, dCarryCost)
    m.dVega = Opt_Vega(m.dAssetPrice, m.dStrikePrice, m.dExpDays, _
        m.dVolatile, dIntRatePct, dCarryCost)
    
    m.bClearCalcItem = False
    SetCalcControls
    m.bClearCalcItem = True
    m.bChanged = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.cmdCalc.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Set the focus upon form activation
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    If m.bFirstShown Then
        MoveFocus Me.txtOptPrice
        SelectAll Me.txtOptPrice
        m.bFirstShown = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form upon loading
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    g.Styler.StyleForm Me
    
    'RH populate list
    With cboOptPrice
        .AddItem "Bid/Ask Avg"
        .AddItem "Bid"
        .AddItem "Ask"
        .AddItem "Last"
        .AddItem "Custom"
    End With
    
    lblDelta.ToolTipText = "Rate of change of option price with respect to change in asset price."
    lblGamma.ToolTipText = "Rate of change of Delta with respect to change in asset price."
    lblVega.ToolTipText = "Rate of change of option price with respect to change in volatility."
    lblTheta.ToolTipText = "Rate of change of option price with respect to change in days to expiration."
    lblRho.ToolTipText = "Rate of change of option price with respect to change in interest rate."

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtIntRate_Change
'' Description: Set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtIntRate_Change()
On Error GoTo ErrSection:

    ClearCalcItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.txtIntRate.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtOptPrice_Change
'' Description: Set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtOptPrice_Change()
On Error GoTo ErrSection:

    ClearCalcItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.txtOptPrice.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStrikePrice_Change
'' Description: Set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStrikePrice_Change()
On Error GoTo ErrSection:

    ClearCalcItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.txtStrikePrice.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTxtOptionPrice
'' Description: Get the option price
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetTxtOptionPrice()
On Error GoTo ErrSection:

    Dim dPrice#, dAvg#
    
    If m.bChanged = False Then
        Exit Sub    'no changes
    End If
    
    If Me.txtOptPrice.Text = "" Then
        InfBox ("Option price cannot be blank.")
        Exit Sub
    End If
    
    dPrice = ValOfText(txtOptPrice.Text)
    dAvg = (m.dOptionAsk + m.dOptionBid) / 2
    dAvg = RoundNum(dAvg, 4)
    
    m.dOptionPrice = dPrice
    
    If dPrice = dAvg Then
        Me.cboOptPrice.ListIndex = 0
        m.dOptionPrice = dAvg
    ElseIf dPrice = m.dOptionBid Then
        Me.cboOptPrice.ListIndex = 1
    ElseIf dPrice = m.dOptionAsk Then
        Me.cboOptPrice.ListIndex = 2
    ElseIf dPrice = m.dOptionLast Then
        Me.cboOptPrice.ListIndex = 3
    Else
        Me.cboOptPrice.ListIndex = 4
    End If
    
    m.dOptionPrice = dPrice
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.GetTxtOptionPrice", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTxtOptionPrice
'' Description: Set the option price
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTxtOptionPrice()
On Error GoTo ErrSection:

    Select Case Me.cboOptPrice.ListIndex
        Case 0
            m.dOptionPrice = (m.dOptionAsk + m.dOptionBid) / 2
            m.dOptionPrice = RoundNum(m.dOptionPrice, 4)
        Case 1
            m.dOptionPrice = m.dOptionBid
        Case 2
            m.dOptionPrice = m.dOptionAsk
        Case 3
            m.dOptionPrice = m.dOptionLast
    End Select
    Me.txtOptPrice.Text = Format(m.dOptionPrice, "0.000")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.SetTxtOptionPrice", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtVolatile_Change
'' Description: Set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtVolatile_Change()
On Error GoTo ErrSection:

    ClearCalcItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptionCalc.txtVolatile.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

