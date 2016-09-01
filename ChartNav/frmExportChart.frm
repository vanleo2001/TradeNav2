VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmExportChart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Chart"
   ClientHeight    =   4905
   ClientLeft      =   2055
   ClientTop       =   1590
   ClientWidth     =   5310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4905
   ScaleWidth      =   5310
   Begin HexUniControls.ctlUniCheckXP chkTimeStamp 
      Height          =   220
      Left            =   3503
      TabIndex        =   2
      Top             =   3765
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
      Caption         =   "frmExportChart.frx":0000
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmExportChart.frx":003E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmExportChart.frx":005E
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtFooter 
      Height          =   315
      Left            =   113
      TabIndex        =   14
      Top             =   4005
      Width           =   5085
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmExportChart.frx":007A
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
      Tip             =   "frmExportChart.frx":009A
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportChart.frx":00BA
   End
   Begin HexUniControls.ctlUniFrameWL fraImageFormat 
      Height          =   1620
      Left            =   113
      TabIndex        =   9
      Top             =   64
      Width           =   5085
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmExportChart.frx":00D6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportChart.frx":0136
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportChart.frx":0156
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkBlackWhite 
         Height          =   255
         Left            =   195
         TabIndex        =   4
         Top             =   810
         Width           =   2310
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmExportChart.frx":0172
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportChart.frx":01C2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":01E2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraBlackWhiteOpts 
         Height          =   600
         Left            =   120
         TabIndex        =   5
         Top             =   855
         Width           =   4755
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmExportChart.frx":01FE
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmExportChart.frx":0240
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":0260
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optBestQuality 
            Height          =   285
            Left            =   675
            TabIndex        =   7
            Top             =   240
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
            Caption         =   "frmExportChart.frx":027C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmExportChart.frx":02B4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExportChart.frx":02D4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optSmallSize 
            Height          =   285
            Left            =   2790
            TabIndex        =   15
            Top             =   240
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
            Caption         =   "frmExportChart.frx":02F0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmExportChart.frx":032A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmExportChart.frx":034A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniRadioXP optGif 
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   360
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
         Caption         =   "frmExportChart.frx":0366
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportChart.frx":038E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":03AE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optBmp 
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   360
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
         Caption         =   "frmExportChart.frx":03CA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportChart.frx":03F8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":0418
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optClipBoard 
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   360
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
         Caption         =   "frmExportChart.frx":0434
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportChart.frx":0468
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":0488
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optJpg 
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   360
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
         Caption         =   "frmExportChart.frx":04A4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportChart.frx":04CC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":04EC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPng 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
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
         Caption         =   "frmExportChart.frx":0508
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmExportChart.frx":0530
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":0550
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdExport 
      Height          =   345
      Left            =   1568
      TabIndex        =   0
      Top             =   4485
      Width           =   1020
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
      Caption         =   "frmExportChart.frx":056C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmExportChart.frx":059A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmExportChart.frx":05BA
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   2708
      TabIndex        =   1
      Top             =   4485
      Width           =   1020
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
      Caption         =   "frmExportChart.frx":05D6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmExportChart.frx":0604
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmExportChart.frx":0624
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraImageSize 
      Height          =   1755
      Left            =   113
      TabIndex        =   8
      Top             =   1830
      Width           =   5085
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmExportChart.frx":0640
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportChart.frx":069C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportChart.frx":06BC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optSizeOnScreen 
         Height          =   375
         Left            =   120
         TabIndex        =   17
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
         Caption         =   "frmExportChart.frx":06D8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportChart.frx":0724
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":0744
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCustomSize 
         Height          =   220
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
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
         Caption         =   "frmExportChart.frx":0760
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportChart.frx":07C6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":07E6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtImgHeight 
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   1350
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmExportChart.frx":0802
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
         Tip             =   "frmExportChart.frx":0822
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":0842
      End
      Begin HexUniControls.ctlUniTextBoxXP txtImgWidth 
         Height          =   300
         Left            =   1200
         TabIndex        =   3
         Top             =   990
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmExportChart.frx":085E
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
         Tip             =   "frmExportChart.frx":087E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":089E
      End
      Begin HexUniControls.ctlUniLabelXP lblHeight 
         Height          =   255
         Left            =   2040
         Top             =   1380
         Width           =   2835
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmExportChart.frx":08BA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportChart.frx":08EE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":090E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblWidth 
         Height          =   255
         Left            =   2040
         Top             =   1020
         Width           =   2835
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmExportChart.frx":092A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportChart.frx":0956
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":0976
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   120
         Top             =   1380
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
         Caption         =   "frmExportChart.frx":0992
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportChart.frx":09CE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":09EE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   120
         Top             =   1020
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
         Caption         =   "frmExportChart.frx":0A0A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportChart.frx":0A44
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportChart.frx":0A64
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label4 
      Height          =   255
      Left            =   120
      Top             =   3765
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
      Caption         =   "frmExportChart.frx":0A80
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmExportChart.frx":0AD0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportChart.frx":0AF0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmExportChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmPrintPreview.frm
'' Description: Generic Print Preview form that allows the user to change
''              different print settings before printing
''
'' Author:      Genesis Financial Data Services
''              425 Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' ??/??/??  M Thorne    Created
'' 02/28/01  D Jarmuth   Made generic
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type mPrivate
    frmObj As Form
    strFile As String
    strPrevFile As String
    dWidthInch As Double
    dHeightPercent As Double
    nWidthPix As Long
    nHeightPix As Long
    nLogPixX As Long
    nFooterAlign As Long
    nCustomSize As Long
    nMonochrome As Long
    nBitDepth As Long
End Type
Private m As mPrivate

Private Sub BrowseFile()
On Error GoTo ErrSection:

    Dim iPos&, strFilter$, strExt$, strPrompt$

    If optBmp.Value = True Then
        strFilter = "Bitmap Files (*.bmp)"
        strExt = ".BMP"
    ElseIf optJpg.Value = True Then
        strFilter = "JPG Files (*.jpg)"
        strExt = ".JPG"
    ElseIf optPng.Value = True Then
        strFilter = "PNG Files (*.png)"
        strExt = ".PNG"
    ElseIf optGif.Value = True Then
        strFilter = "GIF Files (*.gif)"
        strExt = ".GIF"
    Else
        strFilter = "Bitmap Files (*.bmp)|*.bmp|JPG Files (*.jpg)|*.jpg|PNG Files (*.png)|GIF Files (*.gif)"
    End If
    
    m.strFile = CommonDialogFile(frmMain.CommonDialog1, True, strFilter, m.strPrevFile)

    ' strip the extension
    iPos = At(m.strFile, ".", -1)
    If iPos > 0 And iPos >= Len(m.strFile) - 3 Then
        If iPos = 1 Then
            m.strFile = ""
        Else
            m.strFile = Left(m.strFile, iPos - 1)
        End If
    End If
    
    If FileExist(m.strFile & strExt) Then
        'aardvark 6919 & 6920
        strPrompt = Trim(GetIniFileProperty("ExportChartOverwrite", "", "DontAsk", g.strIniFile))
        If Len(strPrompt) = 0 Then
            ' TLB 8/2/2013: verify overwrite
            strPrompt = "Overwrite existing " & FileBase(m.strFile) & strExt & " file?"
            strPrompt = InfBox(strPrompt, "?", "+OK|-Cancel", "Overwrite", , , , , , , , , True)
            If UCase(Left(strPrompt, 1)) = "C" Then
                m.strFile = ""
            ElseIf Right(strPrompt, 1) = "-" Then
                ' don't ask anymore, store for future use
                Call SetIniFileProperty("ExportChartOverwrite", "Y", "DontAsk", g.strIniFile)
            End If
        End If
    End If
    
    If Len(m.strFile) > 0 Then
        m.strPrevFile = m.strFile
    End If
            
    Exit Sub
    
ErrSection:
    RaiseError "frmExportChart.BrowseFile"
    
End Sub

Private Sub chkBlackWhite_Click()

    If chkBlackWhite.Value = vbChecked Then
        m.nMonochrome = 1
    Else
        m.nMonochrome = 0
    End If
    
    SetMonochromeControls

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportChart.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdExport_Click()
On Error GoTo ErrSection:

    Dim rc&, nFormat&, nsaveWd, nsaveHt, nAlign
    Dim iFormatAdjust&
        
    'File Format -  0=clipboard, 1=bmp, 2=jpg, 3=png, 4=gif (color)
    '               5=clipboard, 6=bmp, 7=jpg, 8=png, 9=gif (BW best quality)
    '               10=clipboard, 11=bmp, 12=jpg, 13=png, 14=gif (BW smallest size)
    
    nFormat = -1
    iFormatAdjust = 0
    
    If m.nMonochrome Then
        iFormatAdjust = 5
        If m.nBitDepth = 0 Then iFormatAdjust = 10
    End If
    
    If optClipboard.Value = True Then
        nFormat = 0 + iFormatAdjust
        m.strFile = "Clipboard"     'this is set so grapheng.dll knows to export chart
    ElseIf optBmp.Value = True Then
        nFormat = 1 + iFormatAdjust
    ElseIf optJpg.Value = True Then
        nFormat = 2 + iFormatAdjust
    ElseIf optPng.Value = True Then
        nFormat = 3 + iFormatAdjust
    ElseIf optGif.Value = True Then
        nFormat = 4 + iFormatAdjust
    End If
        
    If nFormat <> 0 And nFormat <> 5 And nFormat <> 10 Then
        BrowseFile
        If Len(m.strFile) = 0 Then Exit Sub
    End If
        
    If Not m.frmObj Is Nothing Then
        Screen.MousePointer = vbHourglass
        Me.Hide
        DoEvents '(to allow this form to hide)
        
        'check for optional footer
        If Len(txtFooter.Text) > 0 Then
            If chkTimeStamp.Value = vbChecked Then
                m.frmObj.cChartObj.FooterPaneId 1, txtFooter.Text & Space(3) & DateFormat(Now, MM_DD_YYYY, HH_MM_SS), m.nFooterAlign
            Else
                m.frmObj.cChartObj.FooterPaneId 1, txtFooter.Text, m.nFooterAlign
            End If
        ElseIf chkTimeStamp.Value = vbChecked Then
            m.frmObj.cChartObj.FooterPaneId 1, Space(3) & DateFormat(Now, MM_DD_YYYY, HH_MM_SS), m.nFooterAlign
        Else
            m.frmObj.cChartObj.FooterPaneId 1, ""
        End If
        
        If m.nCustomSize = 0 Then
        'JM 07-25-2013: Glen wants header string to be smaller font size
            m.frmObj.cChartObj.HdrPaneId 1, 8, 6, 1     '1, 12, 8, 1
            rc = geSaveChart(m.frmObj.cChartObj.geChartObj, m.frmObj.pbChart.hWnd, _
                m.frmObj.pbChart.hDC, 0, 0, nFormat, m.strFile)
        Else
            ' make titles a little smaller than when printing
            rc = Int(ValOfText(txtImgWidth.Text))
            If rc > 6 Then
                m.frmObj.cChartObj.HdrPaneId 1, 12, 8, 1
            ElseIf rc > 4 Then
                m.frmObj.cChartObj.HdrPaneId 1, 10, 8, 1
            Else
                m.frmObj.cChartObj.HdrPaneId 1, 8, 8, 1
            End If
            
            rc = m.frmObj.cChartObj.LoadExportData(m.nWidthPix, m.nHeightPix)
            If rc = 0 Then
                rc = geSaveChart(m.frmObj.cChartObj.geChartObj, m.frmObj.pbChart.hWnd, _
                    m.frmObj.pbChart.hDC, m.nWidthPix, m.nHeightPix, nFormat, m.strFile)
            End If
        End If
        
        If rc <> 0 Then
            InfBox "Export chart failed - err code: " & CStr(rc), "e", , "Error"
        ElseIf nFormat = 0 Or nFormat = 5 Or nFormat = 10 Then
            InfBox "You can now paste the chart image into| another application  (choose 'Edit - Paste' |from that program's menu, or hit 'Ctrl-V')", "i", , "Chart copied to clipboard"
        End If
        
        m.frmObj.cChartObj.HdrPaneId 0
        m.frmObj.cChartObj.FooterPaneId 0, ""
        m.frmObj.cChartObj.RestoreScreenData
        
        Screen.MousePointer = vbDefault
        'save latest settings to file
        If optClipboard Then
            SetIniFileProperty "Format", "CLIP", "ExportChart", g.strIniFile
        ElseIf optJpg Then
            SetIniFileProperty "Format", "JPG", "ExportChart", g.strIniFile
        ElseIf optBmp Then
            SetIniFileProperty "Format", "BMP", "ExportChart", g.strIniFile
        ElseIf optGif Then
            SetIniFileProperty "Format", "GIF", "ExportChart", g.strIniFile
        Else
            SetIniFileProperty "Format", "PNG", "ExportChart", g.strIniFile
        End If
        SetIniFileProperty "Width", m.dWidthInch, "ExportChart", g.strIniFile
        SetIniFileProperty "HeightPercent", m.dHeightPercent, "ExportChart", g.strIniFile
        SetIniFileProperty "Footer", txtFooter.Text, "ExportChart", g.strIniFile
        SetIniFileProperty "FooterAlign", m.nFooterAlign, "ExportChart", g.strIniFile
        SetIniFileProperty "PrevFile", m.strPrevFile, "ExportChart", g.strIniFile
        SetIniFileProperty "CustomSize", m.nCustomSize, "ExportChart", g.strIniFile
        SetIniFileProperty "Monochrome", m.nMonochrome, "ExportChart", g.strIniFile
        SetIniFileProperty "BitDepth", m.nBitDepth, "ExportChart", g.strIniFile
        SetIniFileProperty "DateTimeStamp", chkTimeStamp.Value, "ExportChart", g.strIniFile
                
    End If
    
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportChart.cmdExport.Click", eGDRaiseError_Show
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
    RaiseError "frmExportChart.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_ExportData"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportChart.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Function ShowMe(ByVal frmObj As Form) As Boolean
On Error GoTo ErrSection:

    Dim strFormat$
    
    Set m.frmObj = frmObj
        
    ' set defaults
    If HasModule("CTP") Or HasModule("FIB") Then
        'DiNapoli defaults:
        m.dWidthInch = 7.5
        m.dHeightPercent = 70
        txtFooter = "'Coast Trading Package' -- www.fibtrader.com -- �2003 CIS, Inc. All rights reserved worldwide."
    Else
        m.dWidthInch = 6.5
        m.dHeightPercent = 70
        txtFooter = ""
    End If
    
    ' load settings from prior use
    strFormat = GetIniFileProperty("Format", "PNG", "ExportChart", g.strIniFile)
    Select Case UCase(strFormat)
    Case "BMP"
        optBmp = True
    Case "JPG"
        optJpg = True
    Case "CLIP"
        optClipboard = True
    Case "GIF"
        optGif = True
    Case Else
        optPng = True
    End Select
    m.dWidthInch = GetIniFileProperty("Width", m.dWidthInch, "ExportChart", g.strIniFile)
    m.dHeightPercent = GetIniFileProperty("HeightPercent", m.dHeightPercent, "ExportChart", g.strIniFile)
    txtFooter.Text = GetIniFileProperty("Footer", txtFooter.Text, "ExportChart", g.strIniFile)
    m.nFooterAlign = GetIniFileProperty("FooterAlign", 2, "ExportChart", g.strIniFile)
    m.strPrevFile = GetIniFileProperty("PrevFile", AddSlash(App.Path) & "..\Chart", "ExportChart", g.strIniFile)
    m.nCustomSize = GetIniFileProperty("CustomSize", 1, "ExportChart", g.strIniFile)
    m.nMonochrome = GetIniFileProperty("Monochrome", 0, "ExportChart", g.strIniFile)
    m.nBitDepth = GetIniFileProperty("BitDepth", 0, "ExportChart", g.strIniFile)
    
    chkTimeStamp.Value = GetIniFileProperty("DateTimeStamp", 0, "ExportChart", g.strIniFile)
    
    'set monochrome controls
    chkBlackWhite.Value = m.nMonochrome
    SetMonochromeControls
    
    'get pixels per logical inch
    m.nLogPixX = geDeviceCapsPixX(frmObj.pbChart.hDC)
    
    SetImageSizeControls True
    
    ShowForm Me, True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExportChart.ShowMe", eGDRaiseError_Raise
    
End Function

Private Sub SetMonochromeControls()

    If m.nMonochrome = 0 Then
        optBestQuality.Enabled = False
        optSmallSize.Enabled = False
    Else
        optBestQuality.Enabled = True
        optSmallSize.Enabled = Not optJpg.Value     'cannot do bpp 1 for JPG
        If m.nBitDepth = 0 Then
            optSmallSize.Value = True
        Else
            optBestQuality.Value = True
        End If
    End If

End Sub

Private Sub SetImageSizeControls(ByVal bSetTextCtrl As Boolean)
On Error GoTo ErrSection:

    'conver image size to pixels
    m.nWidthPix = m.dWidthInch * m.nLogPixX
    m.nHeightPix = Round(m.nWidthPix * m.dHeightPercent / 100)
        
    'width control / label
    If bSetTextCtrl = True Then txtImgWidth.Text = m.dWidthInch
    lblWidth.Caption = "inches (when printed) = " & CStr(m.nWidthPix) & " pixels"
    
    'height control / label
    If bSetTextCtrl = True Then txtImgHeight.Text = m.dHeightPercent
    lblHeight.Caption = "% of the width = " & CStr(m.nHeightPix) & " pixels"
    
    'custom size
    optCustomSize.Value = -1 * m.nCustomSize
    optSizeOnScreen.Value = Not optCustomSize.Value

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportChart.SetImageSizeControls", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optBestQuality_Click()
    m.nBitDepth = 1     'just anything > 0
End Sub

Private Sub optBmp_Click()
    SetMonochromeControls
End Sub

Private Sub optClipBoard_Click()
    SetMonochromeControls
End Sub

Private Sub optCustomSize_Click()
    m.nCustomSize = 1
End Sub

Private Sub optGif_Click()
    SetMonochromeControls
End Sub

Private Sub optJpg_Click()
    SetMonochromeControls
End Sub

Private Sub optPng_Click()
    SetMonochromeControls
End Sub

Private Sub optSizeOnScreen_Click()
    m.nCustomSize = 0
End Sub

Private Sub optSmallSize_Click()
    m.nBitDepth = 0
End Sub

Private Sub txtImgHeight_Change()
    m.dHeightPercent = ValOfText(txtImgHeight.Text)
    SetImageSizeControls False
End Sub

Private Sub txtImgWidth_Change()
    m.dWidthInch = ValOfText(txtImgWidth.Text)
    SetImageSizeControls False
End Sub


