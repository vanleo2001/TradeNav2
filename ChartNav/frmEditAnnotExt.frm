VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmEditAnnotExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Annotation"
   ClientHeight    =   12735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12975
   Icon            =   "frmEditAnnotExt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12735
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin HexUniControls.ctlUniFrameWL fraExt 
      Height          =   1335
      Left            =   5955
      TabIndex        =   8
      Top             =   5865
      Visible         =   0   'False
      Width           =   3375
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
      Caption         =   "frmEditAnnotExt.frx":030A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":0352
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":0372
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optExt 
         Height          =   220
         Index           =   4
         Left            =   1680
         TabIndex        =   179
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "frmEditAnnotExt.frx":038E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":03CC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":03EC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExt 
         Height          =   220
         Index           =   3
         Left            =   2520
         TabIndex        =   15
         Top             =   270
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
         Caption         =   "frmEditAnnotExt.frx":0408
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0430
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0450
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExt 
         Height          =   220
         Index           =   2
         Left            =   1830
         TabIndex        =   16
         Top             =   270
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
         Caption         =   "frmEditAnnotExt.frx":046C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0494
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":04B4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExt 
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   17
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
         Caption         =   "frmEditAnnotExt.frx":04D0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":04FA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":051A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExt 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   14
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
         Caption         =   "frmEditAnnotExt.frx":0536
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnotExt.frx":055E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":057E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor clrExt 
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   540
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniComboImageXP cboExtStyle 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   900
         Width           =   1635
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
         Tip             =   "frmEditAnnotExt.frx":059A
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":05BA
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblExtColor 
         Height          =   255
         Left            =   240
         Top             =   600
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
         Caption         =   "frmEditAnnotExt.frx":05D6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0616
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0636
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblExtStyle 
         Height          =   255
         Left            =   240
         Top             =   960
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
         Caption         =   "frmEditAnnotExt.frx":0652
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0692
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":06B2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFibExtDNE 
      Height          =   645
      Left            =   3840
      TabIndex        =   121
      Top             =   960
      Width           =   4800
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
      Caption         =   "frmEditAnnotExt.frx":06CE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":06EE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":070E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optExtDNE 
         Height          =   220
         Index           =   0
         Left            =   525
         TabIndex        =   125
         Top             =   375
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
         Caption         =   "frmEditAnnotExt.frx":072A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnotExt.frx":0752
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0772
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExtDNE 
         Height          =   220
         Index           =   1
         Left            =   1580
         TabIndex        =   124
         Top             =   375
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
         Caption         =   "frmEditAnnotExt.frx":078E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":07B8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":07D8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExtDNE 
         Height          =   220
         Index           =   2
         Left            =   2635
         TabIndex        =   123
         Top             =   375
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
         Caption         =   "frmEditAnnotExt.frx":07F4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":081C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":083C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExtDNE 
         Height          =   220
         Index           =   3
         Left            =   3690
         TabIndex        =   122
         Top             =   375
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
         Caption         =   "frmEditAnnotExt.frx":0858
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0880
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":08A0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   225
         Left            =   60
         Top             =   105
         Width           =   2640
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
         Caption         =   "frmEditAnnotExt.frx":08BC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0902
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0922
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFibText 
      Height          =   2250
      Left            =   6960
      TabIndex        =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   5145
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
      Caption         =   "frmEditAnnotExt.frx":093E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":0970
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":0990
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL Frame5 
         Height          =   285
         Left            =   2190
         TabIndex        =   1
         Top             =   1425
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
         Caption         =   "frmEditAnnotExt.frx":09AC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditAnnotExt.frx":09CC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":09EC
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optFibTextLoc 
            Height          =   220
            Index           =   7
            Left            =   0
            TabIndex        =   6
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmEditAnnotExt.frx":0A08
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":0A30
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":0A50
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optFibTextLoc 
            Height          =   220
            Index           =   8
            Left            =   870
            TabIndex        =   12
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmEditAnnotExt.frx":0A6C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":0A96
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":0AB6
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL Frame4 
         Height          =   285
         Left            =   2190
         TabIndex        =   13
         Top             =   1050
         Width           =   2055
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
         Caption         =   "frmEditAnnotExt.frx":0AD2
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditAnnotExt.frx":0AF2
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0B12
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optFibTextLoc 
            Height          =   220
            Index           =   5
            Left            =   0
            TabIndex        =   52
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmEditAnnotExt.frx":0B2E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":0B56
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":0B76
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optFibTextLoc 
            Height          =   220
            Index           =   6
            Left            =   870
            TabIndex        =   54
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmEditAnnotExt.frx":0B92
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":0BBC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":0BDC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL Frame3 
         Height          =   285
         Left            =   2190
         TabIndex        =   55
         Top             =   675
         Width           =   2805
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
         Caption         =   "frmEditAnnotExt.frx":0BF8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditAnnotExt.frx":0C18
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0C38
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optFibTextLoc 
            Height          =   220
            Index           =   3
            Left            =   870
            TabIndex        =   56
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmEditAnnotExt.frx":0C54
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":0C7E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":0C9E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optFibTextLoc 
            Height          =   220
            Index           =   2
            Left            =   0
            TabIndex        =   57
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmEditAnnotExt.frx":0CBA
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":0CE2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":0D02
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optFibTextLoc 
            Height          =   220
            Index           =   4
            Left            =   1800
            TabIndex        =   58
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmEditAnnotExt.frx":0D1E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":0D4C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":0D6C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniCheckXP chkFibTextMain 
         Height          =   285
         Left            =   480
         TabIndex        =   59
         Top             =   1800
         Width           =   4500
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
         Caption         =   "frmEditAnnotExt.frx":0D88
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0E20
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0E40
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkFibTextShow 
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   60
         Top             =   1425
         Width           =   1590
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
         Caption         =   "frmEditAnnotExt.frx":0E5C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0EA0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0EC0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkFibTextShow 
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   71
         Top             =   1050
         Width           =   1140
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
         Caption         =   "frmEditAnnotExt.frx":0EDC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0F12
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0F32
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkFibTextShow 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   75
         Top             =   300
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
         Caption         =   "frmEditAnnotExt.frx":0F4E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0F7A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":0F9A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkFibTextShow 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   91
         Top             =   675
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
         Caption         =   "frmEditAnnotExt.frx":0FB6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":0FE2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1002
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL Frame2 
         Height          =   285
         Left            =   2190
         TabIndex        =   93
         Top             =   300
         Width           =   2055
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
         Caption         =   "frmEditAnnotExt.frx":101E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditAnnotExt.frx":103E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":105E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optFibTextLoc 
            Height          =   220
            Index           =   0
            Left            =   0
            TabIndex        =   94
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmEditAnnotExt.frx":107A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":10A2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":10C2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optFibTextLoc 
            Height          =   220
            Index           =   1
            Left            =   870
            TabIndex        =   100
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "frmEditAnnotExt.frx":10DE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":1108
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":1128
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraBalloonStrangle 
      Height          =   3735
      Left            =   7680
      TabIndex        =   101
      Top             =   7200
      Width           =   5295
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
      Caption         =   "frmEditAnnotExt.frx":1144
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":117C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":119C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkShowNote 
         Height          =   255
         Left            =   675
         TabIndex        =   102
         Top             =   1830
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
         Caption         =   "frmEditAnnotExt.frx":11B8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":11EA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":120A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdOptionExpire 
         Height          =   285
         Left            =   1560
         TabIndex        =   126
         Top             =   1425
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBalloonCostCall 
         Height          =   285
         Left            =   3970
         TabIndex        =   132
         Top             =   330
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":1226
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
         Tip             =   "frmEditAnnotExt.frx":1246
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1266
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBalloonCostPut 
         Height          =   285
         Left            =   3970
         TabIndex        =   134
         Top             =   1020
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":1282
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
         Tip             =   "frmEditAnnotExt.frx":12A2
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":12C2
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBalloonPut 
         Height          =   285
         Left            =   1560
         TabIndex        =   137
         Top             =   1020
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":12DE
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
         Tip             =   "frmEditAnnotExt.frx":12FE
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":131E
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBalloonStock 
         Height          =   285
         Left            =   1560
         TabIndex        =   142
         Top             =   675
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":133A
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
         Tip             =   "frmEditAnnotExt.frx":135A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":137A
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBalloonCall 
         Height          =   285
         Left            =   1560
         TabIndex        =   151
         Top             =   330
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":1396
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
         Tip             =   "frmEditAnnotExt.frx":13B6
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":13D6
      End
      Begin HexUniControls.ctlUniRichTextBoxXP rtfText 
         Height          =   1485
         Left            =   675
         TabIndex        =   152
         Top             =   2100
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   2619
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":13F2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Tip             =   "frmEditAnnotExt.frx":1412
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1432
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
      Begin HexUniControls.ctlUniLabelXP Label25 
         Height          =   255
         Left            =   240
         Top             =   1440
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
         Caption         =   "frmEditAnnotExt.frx":144E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1490
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":14B0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label24 
         Height          =   255
         Left            =   2790
         Top             =   345
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
         Caption         =   "frmEditAnnotExt.frx":14CC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1504
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1524
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label23 
         Height          =   255
         Left            =   2790
         Top             =   1035
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
         Caption         =   "frmEditAnnotExt.frx":1540
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1576
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1596
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label22 
         Height          =   255
         Left            =   240
         Top             =   1035
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
         Caption         =   "frmEditAnnotExt.frx":15B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":15F2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1612
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label20 
         Height          =   255
         Left            =   240
         Top             =   690
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
         Caption         =   "frmEditAnnotExt.frx":162E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1664
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1684
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label6 
         Height          =   255
         Left            =   240
         Top             =   345
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
         Caption         =   "frmEditAnnotExt.frx":16A0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":16E4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1704
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBalloonCost 
         Height          =   255
         Left            =   2790
         Top             =   690
         Width           =   2205
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
         Caption         =   "frmEditAnnotExt.frx":1720
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":175C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":177C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraGannacciMultiply 
      Height          =   2085
      Left            =   9000
      TabIndex        =   157
      Top             =   7920
      Visible         =   0   'False
      Width           =   3375
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
      Caption         =   "frmEditAnnotExt.frx":1798
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":17B8
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":17D8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtDecimals 
         Height          =   285
         Left            =   1620
         TabIndex        =   180
         Top             =   735
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":17F4
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
         Tip             =   "frmEditAnnotExt.frx":1816
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1836
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdGannacciDebug 
         Height          =   255
         Left            =   2880
         TabIndex        =   165
         Top             =   120
         Width           =   495
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
         Caption         =   "frmEditAnnotExt.frx":1852
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1874
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1894
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtGannacciMultiply 
         Height          =   285
         Left            =   1440
         TabIndex        =   163
         Top             =   330
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":18B0
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmEditAnnotExt.frx":18D0
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":18F0
      End
      Begin HexUniControls.ctlUniCheckXP chkGannacciMultiply 
         Height          =   255
         Left            =   210
         TabIndex        =   158
         Top             =   0
         Width           =   2055
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
         Caption         =   "frmEditAnnotExt.frx":190C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1960
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1980
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDecimals 
         Height          =   285
         Left            =   480
         TabIndex        =   153
         Top             =   735
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
         Caption         =   "frmEditAnnotExt.frx":199C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":19CC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":19EC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPrice3 
         Height          =   255
         Left            =   210
         Top             =   1800
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
         Caption         =   "frmEditAnnotExt.frx":1A08
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1A38
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1A58
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblGannacciMultiply 
         Height          =   285
         Left            =   630
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
         Caption         =   "frmEditAnnotExt.frx":1A74
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1AAA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1ACA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPrice2 
         Height          =   255
         Left            =   210
         Top             =   1500
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
         Caption         =   "frmEditAnnotExt.frx":1AE6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1B16
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1B36
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPrice1 
         Height          =   255
         Left            =   210
         Top             =   1200
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
         Caption         =   "frmEditAnnotExt.frx":1B52
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1B82
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1BA2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDecimals 
         Height          =   285
         Left            =   2160
         Top             =   750
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
         Caption         =   "frmEditAnnotExt.frx":1BBE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1BEE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1C0E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraGannacciSwing 
      Height          =   4095
      Left            =   9000
      TabIndex        =   146
      Top             =   3840
      Width           =   3375
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
      Caption         =   "frmEditAnnotExt.frx":1C2A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":1C68
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":1C88
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkShowTextBorder 
         Height          =   255
         Left            =   150
         TabIndex        =   160
         Top             =   500
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
         Caption         =   "frmEditAnnotExt.frx":1CA4
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1CE4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1D04
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkIncludeFirstBar 
         Height          =   255
         Left            =   150
         TabIndex        =   159
         Top             =   1020
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
         Caption         =   "frmEditAnnotExt.frx":1D20
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1D7C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1D9C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraGannacciColors 
         Height          =   855
         Left            =   150
         TabIndex        =   150
         Top             =   3120
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
         Caption         =   "frmEditAnnotExt.frx":1DB8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditAnnotExt.frx":1E04
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1E24
         RightToLeft     =   0   'False
         Begin gdOCX.gdSelectColor gdGannacciLow 
            Height          =   255
            Left            =   180
            TabIndex        =   154
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdGannacciMed 
            Height          =   255
            Left            =   1140
            TabIndex        =   155
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdGannacciHigh 
            Height          =   255
            Left            =   2100
            TabIndex        =   156
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label19 
            Height          =   255
            Left            =   2100
            Top             =   240
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
            Caption         =   "frmEditAnnotExt.frx":1E40
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnotExt.frx":1E68
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":1E88
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label8 
            Height          =   255
            Left            =   1140
            Top             =   240
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
            Caption         =   "frmEditAnnotExt.frx":1EA4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnotExt.frx":1ECE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":1EEE
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label7 
            Height          =   255
            Left            =   180
            Top             =   240
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
            Caption         =   "frmEditAnnotExt.frx":1F0A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnotExt.frx":1F30
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":1F50
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid fgGannacciSwing 
         Height          =   1575
         Left            =   150
         TabIndex        =   149
         Top             =   1320
         Width           =   3015
         _cx             =   5318
         _cy             =   2778
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
      Begin HexUniControls.ctlUniCheckXP chkShowMarkers 
         Height          =   255
         Left            =   150
         TabIndex        =   148
         Top             =   760
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
         Caption         =   "frmEditAnnotExt.frx":1F6C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":1FA4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":1FC4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkShowText 
         Height          =   255
         Left            =   150
         TabIndex        =   147
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
         Caption         =   "frmEditAnnotExt.frx":1FE0
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2012
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2032
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRichTextBoxXP rtfGannacciDebug 
         Height          =   375
         Left            =   2280
         TabIndex        =   166
         Top             =   360
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":204E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   -1
         MultiLine       =   -1  'True
         Alignment       =   0
         ScrollBars      =   2
         PasswordChar    =   ""
         TrapTab         =   0   'False
         RaiseChangeEvent=   -1  'True
         RaiseUpdateEvent=   0   'False
         RaiseSelChangeEvent=   -1  'True
         Tip             =   "frmEditAnnotExt.frx":206E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":208E
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
   End
   Begin HexUniControls.ctlUniFrameWL fraGannacciTime 
      Height          =   2715
      Left            =   9480
      TabIndex        =   143
      Top             =   600
      Width           =   3300
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
      Caption         =   "frmEditAnnotExt.frx":20AA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":2112
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":2132
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL Frame1 
         Height          =   1785
         Left            =   195
         TabIndex        =   168
         Top             =   840
         Width           =   2900
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
         Caption         =   "frmEditAnnotExt.frx":214E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditAnnotExt.frx":21A2
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":21C2
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chk180 
            Height          =   255
            Left            =   420
            TabIndex        =   174
            Top             =   840
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
            Caption         =   "frmEditAnnotExt.frx":21DE
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":2204
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":2224
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chk90 
            Height          =   255
            Left            =   420
            TabIndex        =   173
            Top             =   540
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
            Caption         =   "frmEditAnnotExt.frx":2240
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":2264
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":2284
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chk45 
            Height          =   255
            Left            =   420
            TabIndex        =   172
            Top             =   240
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
            Caption         =   "frmEditAnnotExt.frx":22A0
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":22C4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":22E4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chk270 
            Height          =   255
            Left            =   420
            TabIndex        =   170
            Top             =   1140
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
            Caption         =   "frmEditAnnotExt.frx":2300
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":2326
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":2346
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chk360 
            Height          =   255
            Left            =   420
            TabIndex        =   169
            Top             =   1440
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
            Caption         =   "frmEditAnnotExt.frx":2362
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnotExt.frx":2388
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnotExt.frx":23A8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor gdGannacci45 
            Height          =   255
            Left            =   1380
            TabIndex        =   171
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdGannacci90 
            Height          =   255
            Left            =   1380
            TabIndex        =   175
            Top             =   540
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdGannacci180 
            Height          =   255
            Left            =   1380
            TabIndex        =   176
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdGannacci270 
            Height          =   255
            Left            =   1380
            TabIndex        =   177
            Top             =   1140
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdGannacci360 
            Height          =   255
            Left            =   1380
            TabIndex        =   178
            Top             =   1440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            CustomColor     =   255
         End
      End
      Begin HexUniControls.ctlUniCheckXP chkGannacciTimeBars 
         Height          =   255
         Left            =   150
         TabIndex        =   145
         Top             =   270
         Width           =   2775
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
         Caption         =   "frmEditAnnotExt.frx":23C4
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2400
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2420
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkGannacciTimeCalendar 
         Height          =   255
         Left            =   150
         TabIndex        =   144
         Top             =   510
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
         Caption         =   "frmEditAnnotExt.frx":243C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":24A8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":24C8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraGannacciCycle 
      Height          =   435
      Left            =   9960
      TabIndex        =   140
      Top             =   11880
      Width           =   1995
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
      Caption         =   "frmEditAnnotExt.frx":24E4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":2504
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":2524
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtGannacciYears 
         Height          =   285
         Left            =   1320
         TabIndex        =   141
         Top             =   105
         Width           =   495
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":2540
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
         Tip             =   "frmEditAnnotExt.frx":2564
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2584
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   255
         Left            =   0
         Top             =   120
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
         Caption         =   "frmEditAnnotExt.frx":25A0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":25E0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2600
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraZoneOptions 
      Height          =   900
      Left            =   4680
      TabIndex        =   136
      Top             =   10080
      Visible         =   0   'False
      Width           =   2955
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
      Caption         =   "frmEditAnnotExt.frx":261C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":263C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":265C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optFromPrevBar 
         Height          =   255
         Left            =   1440
         TabIndex        =   139
         Top             =   345
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
         Caption         =   "frmEditAnnotExt.frx":2678
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":26B2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":26D2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optFromFirstBar 
         Height          =   255
         Left            =   360
         TabIndex        =   138
         Top             =   345
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
         Caption         =   "frmEditAnnotExt.frx":26EE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2722
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2742
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   120
         Top             =   135
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
         Caption         =   "frmEditAnnotExt.frx":275E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":27B8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":27D8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraMainExtDNE 
      Height          =   645
      Left            =   5295
      TabIndex        =   127
      Top             =   1470
      Width           =   4800
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
      Caption         =   "frmEditAnnotExt.frx":27F4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":2814
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":2834
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optExtMainDNE 
         Height          =   220
         Index           =   3
         Left            =   3690
         TabIndex        =   131
         Top             =   375
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
         Caption         =   "frmEditAnnotExt.frx":2850
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2878
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2898
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExtMainDNE 
         Height          =   220
         Index           =   2
         Left            =   2635
         TabIndex        =   130
         Top             =   375
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
         Caption         =   "frmEditAnnotExt.frx":28B4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":28DC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":28FC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExtMainDNE 
         Height          =   220
         Index           =   1
         Left            =   1580
         TabIndex        =   129
         Top             =   375
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
         Caption         =   "frmEditAnnotExt.frx":2918
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2942
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2962
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExtMainDNE 
         Height          =   220
         Index           =   0
         Left            =   525
         TabIndex        =   128
         Top             =   375
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
         Caption         =   "frmEditAnnotExt.frx":297E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnotExt.frx":29A6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":29C6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   225
         Left            =   60
         Top             =   105
         Width           =   2640
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
         Caption         =   "frmEditAnnotExt.frx":29E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2A2E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2A4E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkDneCircles 
      Height          =   255
      Left            =   7140
      TabIndex        =   133
      Top             =   630
      Width           =   2115
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
      Caption         =   "frmEditAnnotExt.frx":2A6A
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnotExt.frx":2ABA
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":2ADA
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraPivotText 
      Height          =   900
      Left            =   4335
      TabIndex        =   111
      Top             =   11730
      Visible         =   0   'False
      Width           =   4905
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
      Caption         =   "frmEditAnnotExt.frx":2AF6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":2B3E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":2B5E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkPriceOnly 
         Height          =   285
         Left            =   3210
         TabIndex        =   117
         Top             =   510
         Width           =   1500
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
         Caption         =   "frmEditAnnotExt.frx":2B7A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2BAE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2BCE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkTextNextToMain 
         Height          =   285
         Left            =   270
         TabIndex        =   116
         Top             =   510
         Width           =   3060
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
         Caption         =   "frmEditAnnotExt.frx":2BEA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2C42
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2C62
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPivotText 
         Height          =   195
         Index           =   3
         Left            =   4080
         TabIndex        =   114
         Top             =   240
         Width           =   780
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
         Caption         =   "frmEditAnnotExt.frx":2C7E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2CA8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2CC8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPivotText 
         Height          =   195
         Index           =   2
         Left            =   3210
         TabIndex        =   113
         Top             =   240
         Width           =   780
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
         Caption         =   "frmEditAnnotExt.frx":2CE4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2D0C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2D2C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPivotText 
         Height          =   195
         Index           =   1
         Left            =   1185
         TabIndex        =   112
         Top             =   240
         Width           =   2055
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
         Caption         =   "frmEditAnnotExt.frx":2D48
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2D96
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2DB6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPivotText 
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   115
         Top             =   240
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
         Caption         =   "frmEditAnnotExt.frx":2DD2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnotExt.frx":2DFA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2E1A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraDNR 
      Height          =   4575
      Left            =   210
      TabIndex        =   18
      Top             =   8100
      Visible         =   0   'False
      Width           =   3975
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
      Caption         =   "frmEditAnnotExt.frx":2E36
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":2E8E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":2EAE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkDnrCircles 
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   2115
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
         Caption         =   "frmEditAnnotExt.frx":2ECA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":2F1A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2F3A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor clrDnrArc 
         Height          =   315
         Left            =   1035
         TabIndex        =   50
         Top             =   2820
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrDnr3 
         Height          =   315
         Left            =   1035
         TabIndex        =   46
         Top             =   2100
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrDnr2 
         Height          =   315
         Left            =   1035
         TabIndex        =   45
         Top             =   2460
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrDnr1 
         Height          =   315
         Left            =   1035
         TabIndex        =   44
         Top             =   1740
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDnrRatio1 
         Height          =   315
         Left            =   135
         TabIndex        =   40
         Top             =   1740
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":2F56
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
         Tip             =   "frmEditAnnotExt.frx":2F80
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":2FA0
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDnrRatio2 
         Height          =   315
         Left            =   135
         TabIndex        =   41
         Top             =   2460
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":2FBC
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
         Tip             =   "frmEditAnnotExt.frx":2FE6
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3006
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDnrRatio3 
         Height          =   315
         Left            =   135
         TabIndex        =   42
         Top             =   2100
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":3022
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
         Tip             =   "frmEditAnnotExt.frx":304C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":306C
      End
      Begin HexUniControls.ctlUniComboImageXP cboDnrStyle1 
         Height          =   315
         Left            =   2175
         TabIndex        =   47
         Top             =   1740
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":3088
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":30A8
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboDnrStyle2 
         Height          =   315
         Left            =   2175
         TabIndex        =   48
         Top             =   2460
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":30C4
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":30E4
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboDnrStyle3 
         Height          =   315
         Left            =   2175
         TabIndex        =   49
         Top             =   2100
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":3100
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3120
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboDnrStyleArc 
         Height          =   315
         Left            =   2175
         TabIndex        =   51
         Top             =   2820
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":313C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":315C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fgDnrLineage 
         Height          =   1215
         Left            =   135
         TabIndex        =   53
         Top             =   3300
         Width           =   3495
         _cx             =   6165
         _cy             =   2143
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
      Begin HexUniControls.ctlUniLabelXP Label12 
         Height          =   255
         Left            =   2235
         Top             =   1500
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
         Caption         =   "frmEditAnnotExt.frx":3178
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":31A2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":31C2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label11 
         Height          =   255
         Left            =   1095
         Top             =   1500
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
         Caption         =   "frmEditAnnotExt.frx":31DE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3208
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3228
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label10 
         Height          =   255
         Left            =   195
         Top             =   1500
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
         Caption         =   "frmEditAnnotExt.frx":3244
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":326E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":328E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label9 
         Height          =   255
         Left            =   615
         Top             =   2850
         Width           =   315
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
         Caption         =   "frmEditAnnotExt.frx":32AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":32D2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":32F2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   315
      Left            =   3975
      TabIndex        =   3
      Top             =   105
      Width           =   915
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
      Caption         =   "frmEditAnnotExt.frx":330E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmEditAnnotExt.frx":333C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":335C
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraTimeCycle 
      Height          =   1335
      Left            =   5955
      TabIndex        =   90
      Top             =   4500
      Width           =   3195
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
      Caption         =   "frmEditAnnotExt.frx":3378
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":33AC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":33CC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optArcs 
         Height          =   220
         Index           =   1
         Left            =   1200
         TabIndex        =   105
         Top             =   1020
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
         Caption         =   "frmEditAnnotExt.frx":33E8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3410
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3430
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optArcs 
         Height          =   220
         Index           =   0
         Left            =   180
         TabIndex        =   104
         Top             =   1020
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
         Caption         =   "frmEditAnnotExt.frx":344C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3476
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3496
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdGoToBaseLine 
         Height          =   285
         Left            =   2340
         TabIndex        =   96
         Top             =   600
         Width           =   735
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
         Caption         =   "frmEditAnnotExt.frx":34B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":34DC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":34FC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBaseLine 
         Height          =   285
         Left            =   1200
         TabIndex        =   95
         Top             =   600
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":3518
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
         Tip             =   "frmEditAnnotExt.frx":3552
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3572
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBarsSpacing 
         Height          =   285
         Left            =   1200
         TabIndex        =   92
         Top             =   255
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":358E
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
         Tip             =   "frmEditAnnotExt.frx":35CA
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":35EA
      End
      Begin HexUniControls.ctlUniLabelXP Label33 
         Height          =   195
         Left            =   180
         Top             =   660
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
         Caption         =   "frmEditAnnotExt.frx":3606
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":363A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":365A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label32 
         Height          =   195
         Left            =   2340
         Top             =   300
         Width           =   675
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
         Caption         =   "frmEditAnnotExt.frx":3676
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":36A2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":36C2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label30 
         Height          =   195
         Left            =   180
         Top             =   300
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
         Caption         =   "frmEditAnnotExt.frx":36DE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":370E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":372E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraSpResistFan 
      Height          =   1665
      Left            =   4320
      TabIndex        =   97
      Top             =   8115
      Visible         =   0   'False
      Width           =   3255
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
      Caption         =   "frmEditAnnotExt.frx":374A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":378A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":37AA
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboResistStyle 
         Height          =   315
         Left            =   1020
         TabIndex        =   103
         Top             =   1200
         Width           =   2055
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
         Tip             =   "frmEditAnnotExt.frx":37C6
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":37E6
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboResistCount 
         Height          =   315
         Left            =   1860
         TabIndex        =   99
         Top             =   330
         Width           =   675
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
         Tip             =   "frmEditAnnotExt.frx":3802
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3822
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor gdResistColor 
         Height          =   315
         Left            =   1020
         TabIndex        =   98
         Top             =   780
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label21 
         Height          =   195
         Left            =   360
         Top             =   1260
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
         Caption         =   "frmEditAnnotExt.frx":383E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3868
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3888
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   195
         Left            =   360
         Top             =   840
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
         Caption         =   "frmEditAnnotExt.frx":38A4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":38CE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":38EE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label18 
         Height          =   255
         Left            =   660
         Top             =   360
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
         Caption         =   "frmEditAnnotExt.frx":390A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3948
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3968
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkFreeFloat 
      Height          =   220
      Left            =   5085
      TabIndex        =   120
      Top             =   660
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
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
      Caption         =   "frmEditAnnotExt.frx":3984
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnotExt.frx":39BE
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":39DE
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboGreenBlattZones 
      Height          =   315
      Left            =   2835
      TabIndex        =   119
      Top             =   540
      Width           =   1740
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
      Tip             =   "frmEditAnnotExt.frx":39FA
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":3A1A
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkLucasNumbers 
      Height          =   220
      Left            =   5085
      TabIndex        =   118
      Top             =   405
      Width           =   1740
      _ExtentX        =   3069
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
      Caption         =   "frmEditAnnotExt.frx":3A36
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnotExt.frx":3A78
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":3A98
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraTimeLines 
      Height          =   1995
      Left            =   135
      TabIndex        =   83
      Top             =   2310
      Width           =   3195
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
      Caption         =   "frmEditAnnotExt.frx":3AB4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":3AF2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":3B12
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optFibZoneArcs 
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   107
         Top             =   1680
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
         Caption         =   "frmEditAnnotExt.frx":3B2E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3B56
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3B76
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optFibZoneArcs 
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   106
         Top             =   1680
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
         Caption         =   "frmEditAnnotExt.frx":3B92
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3BBC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3BDC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fgTimeLines 
         Height          =   1275
         Left            =   120
         TabIndex        =   84
         Top             =   300
         Width           =   2955
         _cx             =   5212
         _cy             =   2249
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   615
      Left            =   60
      TabIndex        =   7
      Top             =   1695
      Width           =   3375
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
      Caption         =   "frmEditAnnotExt.frx":3BF8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":3C18
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":3C38
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   330
         Left            =   2575
         TabIndex        =   39
         Top             =   180
         Width           =   750
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
         Caption         =   "frmEditAnnotExt.frx":3C54
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3C82
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3CA2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   330
         Left            =   60
         TabIndex        =   38
         Top             =   180
         Width           =   750
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
         Caption         =   "frmEditAnnotExt.frx":3CBE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3CE4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3D04
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSaveDefaults 
         Height          =   330
         Left            =   965
         TabIndex        =   37
         Top             =   180
         Width           =   1455
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
         Caption         =   "frmEditAnnotExt.frx":3D20
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3D62
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3D82
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkMultiChart 
      Height          =   220
      Left            =   180
      TabIndex        =   80
      Top             =   1200
      Visible         =   0   'False
      Width           =   4500
      _ExtentX        =   7938
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
      Caption         =   "frmEditAnnotExt.frx":3D9E
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnotExt.frx":3E0E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":3E2E
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkDynamic 
      Height          =   220
      Left            =   180
      TabIndex        =   81
      Top             =   1455
      Width           =   2955
      _ExtentX        =   5212
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
      Caption         =   "frmEditAnnotExt.frx":3E4A
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnotExt.frx":3EB4
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":3ED4
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraQuadrants 
      Height          =   885
      Left            =   5955
      TabIndex        =   85
      Top             =   7230
      Width           =   3870
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
      Caption         =   "frmEditAnnotExt.frx":3EF0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":3F22
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":3F42
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkQuadrant 
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   89
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
         Caption         =   "frmEditAnnotExt.frx":3F5E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":3F94
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":3FB4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkQuadrant 
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   88
         Top             =   570
         Width           =   1650
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
         Caption         =   "frmEditAnnotExt.frx":3FD0
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4006
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4026
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkQuadrant 
         Height          =   220
         Index           =   2
         Left            =   480
         TabIndex        =   87
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "frmEditAnnotExt.frx":4042
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4076
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4096
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkQuadrant 
         Height          =   220
         Index           =   3
         Left            =   480
         TabIndex        =   86
         Top             =   570
         Width           =   1530
         _ExtentX        =   2699
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
         Caption         =   "frmEditAnnotExt.frx":40B2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":40E6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4106
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkAllPanes 
      Height          =   255
      Left            =   5085
      TabIndex        =   82
      Top             =   90
      Width           =   1635
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
      Caption         =   "frmEditAnnotExt.frx":4122
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnotExt.frx":4164
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":4184
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdFont 
      Height          =   330
      Left            =   2640
      TabIndex        =   72
      Top             =   105
      Width           =   750
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
      Caption         =   "frmEditAnnotExt.frx":41A0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmEditAnnotExt.frx":41CA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":41EA
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkPreIndicator 
      Height          =   220
      Left            =   180
      TabIndex        =   74
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "frmEditAnnotExt.frx":4206
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnotExt.frx":4252
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":4272
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin VB.Timer tmrEditAnnot 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9225
      Top             =   105
   End
   Begin HexUniControls.ctlUniComboImageXP cboStyle 
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   540
      Width           =   1635
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
      Tip             =   "frmEditAnnotExt.frx":428E
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":42AE
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin gdOCX.gdSelectColor clrColor 
      Height          =   315
      Left            =   720
      TabIndex        =   73
      Top             =   135
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      CustomColor     =   255
   End
   Begin HexUniControls.ctlUniFrameWL fraDNE 
      Height          =   3675
      Left            =   105
      TabIndex        =   19
      Top             =   4350
      Visible         =   0   'False
      Width           =   5835
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
      Caption         =   "frmEditAnnotExt.frx":42CA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":430E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":432E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtDneValueXOP 
         Height          =   315
         Left            =   3585
         TabIndex        =   70
         Top             =   3300
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":434A
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
         Tip             =   "frmEditAnnotExt.frx":4386
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":43A6
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneValueOP 
         Height          =   315
         Left            =   3585
         TabIndex        =   69
         Top             =   2940
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":43C2
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
         Tip             =   "frmEditAnnotExt.frx":43FC
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":441C
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneValueCOP 
         Height          =   315
         Left            =   3585
         TabIndex        =   68
         Top             =   2580
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":4438
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
         Tip             =   "frmEditAnnotExt.frx":4474
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4494
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneValueC 
         Height          =   315
         Left            =   3585
         TabIndex        =   67
         Top             =   2100
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":44B0
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
         Tip             =   "frmEditAnnotExt.frx":44E8
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4508
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneValueB 
         Height          =   315
         Left            =   3585
         TabIndex        =   66
         Top             =   1740
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":4524
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
         Tip             =   "frmEditAnnotExt.frx":455C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":457C
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneValueA 
         Height          =   315
         Left            =   3585
         TabIndex        =   65
         Top             =   1380
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":4598
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
         Tip             =   "frmEditAnnotExt.frx":45D0
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":45F0
      End
      Begin HexUniControls.ctlUniComboImageXP cboDneStyleCOP 
         Height          =   315
         Left            =   2010
         TabIndex        =   64
         Top             =   2580
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":460C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":462C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneRatioCOP 
         Height          =   315
         Left            =   4680
         TabIndex        =   63
         Top             =   2580
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":4648
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
         Tip             =   "frmEditAnnotExt.frx":4672
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4692
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneLabelCOP 
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   2580
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":46AE
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
         Tip             =   "frmEditAnnotExt.frx":46D4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":46F4
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneRatioOP 
         Height          =   315
         Left            =   4680
         TabIndex        =   36
         Top             =   2940
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":4710
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
         Tip             =   "frmEditAnnotExt.frx":473A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":475A
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneRatioXOP 
         Height          =   315
         Left            =   4680
         TabIndex        =   35
         Top             =   3300
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":4776
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
         Tip             =   "frmEditAnnotExt.frx":47A0
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":47C0
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneLabelOP 
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   2940
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":47DC
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
         Tip             =   "frmEditAnnotExt.frx":4800
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4820
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneLabelXOP 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   3300
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":483C
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
         Tip             =   "frmEditAnnotExt.frx":4862
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4882
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneLabelC 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   2100
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":489E
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
         Tip             =   "frmEditAnnotExt.frx":48C0
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":48E0
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneLabelB 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   1740
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":48FC
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
         Tip             =   "frmEditAnnotExt.frx":491E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":493E
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDneLabelA 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   1380
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnotExt.frx":495A
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
         Tip             =   "frmEditAnnotExt.frx":497C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":499C
      End
      Begin HexUniControls.ctlUniComboImageXP cboDneStyleOP 
         Height          =   315
         Left            =   2010
         TabIndex        =   28
         Top             =   2940
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":49B8
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":49D8
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboDneStyleXOP 
         Height          =   315
         Left            =   2010
         TabIndex        =   27
         Top             =   3300
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":49F4
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4A14
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboDneStyleC 
         Height          =   315
         Left            =   2010
         TabIndex        =   26
         Top             =   2100
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":4A30
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4A50
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboDneStyleB 
         Height          =   315
         Left            =   2010
         TabIndex        =   25
         Top             =   1740
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":4A6C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4A8C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboDneStyleA 
         Height          =   315
         Left            =   2010
         TabIndex        =   24
         Top             =   1380
         Width           =   1515
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
         Tip             =   "frmEditAnnotExt.frx":4AA8
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4AC8
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor clrDneB 
         Height          =   315
         Left            =   855
         TabIndex        =   20
         Top             =   1740
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrDneC 
         Height          =   315
         Left            =   855
         TabIndex        =   21
         Top             =   2100
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrDneXOP 
         Height          =   315
         Left            =   855
         TabIndex        =   22
         Top             =   3300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrDneOP 
         Height          =   315
         Left            =   855
         TabIndex        =   23
         Top             =   2940
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrDneA 
         Height          =   315
         Left            =   855
         TabIndex        =   34
         Top             =   1380
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrDneCOP 
         Height          =   315
         Left            =   855
         TabIndex        =   62
         Top             =   2580
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label17 
         Height          =   255
         Left            =   3585
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
         Caption         =   "frmEditAnnotExt.frx":4AE4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4B0E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4B2E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label16 
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
         Caption         =   "frmEditAnnotExt.frx":4B4A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4B74
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4B94
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label15 
         Height          =   255
         Left            =   4680
         Top             =   2340
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
         Caption         =   "frmEditAnnotExt.frx":4BB0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4BDA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4BFA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label14 
         Height          =   255
         Left            =   840
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
         Caption         =   "frmEditAnnotExt.frx":4C16
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4C40
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4C60
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label13 
         Height          =   255
         Left            =   1965
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
         Caption         =   "frmEditAnnotExt.frx":4C7C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4CA6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4CC6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFib 
      Height          =   2505
      Left            =   3465
      TabIndex        =   9
      Top             =   1785
      Visible         =   0   'False
      Width           =   5295
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
      Caption         =   "frmEditAnnotExt.frx":4CE2
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnotExt.frx":4D20
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":4D40
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkCircular 
         Height          =   255
         Left            =   3600
         TabIndex        =   110
         Top             =   2205
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
         Caption         =   "frmEditAnnotExt.frx":4D5C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4D96
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4DB6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkExtendArcs 
         Height          =   255
         Left            =   2040
         TabIndex        =   108
         Top             =   2205
         Visible         =   0   'False
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
         Caption         =   "frmEditAnnotExt.frx":4DD2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4E08
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4E28
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkFibValues 
         Height          =   255
         Left            =   300
         TabIndex        =   79
         Top             =   2205
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
         Caption         =   "frmEditAnnotExt.frx":4E44
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4E82
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4EA2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   315
         Left            =   300
         TabIndex        =   78
         Top             =   840
         Width           =   1455
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
         Caption         =   "frmEditAnnotExt.frx":4EBE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4EF0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4F10
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
         Height          =   315
         Left            =   300
         TabIndex        =   77
         Top             =   1200
         Width           =   1455
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
         Caption         =   "frmEditAnnotExt.frx":4F2C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4F64
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":4F84
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRestoreOrig 
         Height          =   315
         Left            =   300
         TabIndex        =   76
         Top             =   1560
         Width           =   1455
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
         Caption         =   "frmEditAnnotExt.frx":4FA0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":4FE0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":5000
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboFibStyle 
         Height          =   315
         Left            =   540
         TabIndex        =   10
         Top             =   420
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
         Tip             =   "frmEditAnnotExt.frx":501C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":503C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor clrFib 
         Height          =   315
         Left            =   4440
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin VSFlex7LCtl.VSFlexGrid fgFib 
         Height          =   1905
         Left            =   2040
         TabIndex        =   109
         Top             =   210
         Width           =   3090
         _cx             =   5450
         _cy             =   3360
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin HexUniControls.ctlUniComboImageXP cboGartleyEndpoints 
         Height          =   315
         Left            =   30
         TabIndex        =   135
         Top             =   1950
         Width           =   2085
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
         Tip             =   "frmEditAnnotExt.frx":5058
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":5078
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkBalloonPrices 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   161
         Top             =   1200
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
         Caption         =   "frmEditAnnotExt.frx":5094
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":50CC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":50EC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor gdBalloonColorBE 
         Height          =   315
         Left            =   0
         TabIndex        =   162
         Top             =   720
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniCheckXP chkBalloonPrices 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   164
         Top             =   1500
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
         Caption         =   "frmEditAnnotExt.frx":5108
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":5146
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":5166
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkBalloonPrices 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   167
         Top             =   1800
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
         Caption         =   "frmEditAnnotExt.frx":5182
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnotExt.frx":51C0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":51E0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblGartleyEndpoints 
         Height          =   225
         Left            =   195
         Top             =   1830
         Width           =   1905
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
         Caption         =   "frmEditAnnotExt.frx":51FC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":524C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":526C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFibColor 
         Height          =   255
         Left            =   3240
         Top             =   2130
         Visible         =   0   'False
         Width           =   435
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
         Caption         =   "frmEditAnnotExt.frx":5288
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":52B4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":52D4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFibStyle 
         Height          =   255
         Left            =   120
         Top             =   450
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
         Caption         =   "frmEditAnnotExt.frx":52F0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnotExt.frx":531C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnotExt.frx":533C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblStyle 
      Height          =   255
      Left            =   180
      Top             =   600
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
      Caption         =   "frmEditAnnotExt.frx":5358
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmEditAnnotExt.frx":5384
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":53A4
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   255
      Left            =   180
      Top             =   180
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
      Caption         =   "frmEditAnnotExt.frx":53C0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmEditAnnotExt.frx":53EC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnotExt.frx":540C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmEditAnnotExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kFraFibWd = 5295
Private Const kFraFibDnExWd = 6095

Private Type mPrivate
    Chart As cChart
    Annot As cAnnotation
    nAnnotIdx As Long
    bTextChanged As Boolean
    bMultiChartOption As Boolean
    bIgnoreUnload As Boolean    '(so MDI activate won't unload this form when a modal form called from here)
    bCenterColorStyle As Boolean
    bWasMultiChart As Boolean
    bDescending As Boolean      'sort flag
    
    nFibRowDown As Long         'to keep track of color changes in grid
    nFibColDown As Long
End Type
Private m As mPrivate

Private Sub cboChannels_Click()
    Repaint
End Sub

Private Sub cboDneStyleA_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDneStyleA.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboDneStyleB_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDneStyleB.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboDneStyleC_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDneStyleC.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboDneStyleCOP_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDneStyleCOP.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboDneStyleOP_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDneStyleOP.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboDneStyleXOP_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDneStyleXOP.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboDnrStyle1_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDnrStyle1.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboDnrStyle2_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDnrStyle2.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboDnrStyle3_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDnrStyle3.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboDnrStyleArc_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboDneStyleArc.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboGartleyEndpoints_Click()
On Error GoTo ErrSection:
    
    Dim i&

    If Me.Visible Then
        If Not m.Annot Is Nothing Then
            i = Val(m.Annot.Prop("GartleyLabelStyle"))
            If i <> cboGartleyEndpoints.ListIndex Then Repaint
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboGartleyEndpoints_Click"

End Sub

Private Sub cboGreenBlattZones_Click()
On Error GoTo ErrSection:

    Dim i&, eZone As eGBZoneInUse
    
    If Not m.Annot Is Nothing Then
        i = cboGreenBlattZones.ListIndex
        If i >= eANNOT_GB_ZoneLucas And i <= eANNOT_GB_ZoneSqRoot Then
            eZone = i
            If m.Annot.ZoneInUse <> eZone Then
                m.Annot.ZoneInUse = eZone
                Repaint
                SetTimeZoneGrid
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboGreenBlattZones_Click"

End Sub

Private Sub cboR50Style_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboR50Style_Click", eGDRaiseError_Show
    
End Sub

Private Sub cboExtStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboExtStyle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboFibStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboFibStyle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboLineStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboLineStyle.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboResistCount_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboResistCount.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboResistStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboResistStyle.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboStdDevStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboStdDevStyle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboStyle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboTextJustify_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cboTextJustify.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chk1st_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chk1st.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chk2nd_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chk2nd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chk3rd_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chk3rd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chk45_Click()
    Repaint
End Sub

Private Sub chk90_Click()
    Repaint
End Sub

Private Sub chk180_Click()
    Repaint
End Sub

Private Sub chk270_Click()
    Repaint
End Sub

Private Sub chk360_Click()
    Repaint
End Sub

Private Sub chkAllPanes_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkAllPanes_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkAxes_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkAxes.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkBalloonPrices_Click(Index As Integer)
    Repaint
End Sub

Private Sub chkCircular_Click()
    Repaint
End Sub

Private Sub chkDisplaySRValue_Click()
    Repaint
End Sub

Private Sub chkDecimals_Click()
    Repaint
End Sub

Private Sub chkDneCircles_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkDneCircles.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkDnrCircles_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkDnrCircles.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkDynamic_Click()
On Error GoTo ErrSection:
    
    With m.Annot
        Repaint
        If .eType = eANNOT_FibTimeZones Then
            cmdFont.Enabled = chkDynamic.Value
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkDynamic.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkExtendArcs_Click()
    Repaint
End Sub

Private Sub chkExtendSRLine_Click()
    Repaint
End Sub

Private Sub chkFibTextMain_Click()
    Repaint
End Sub

Private Sub chkFibTextShow_Click(Index As Integer)
    Repaint
End Sub

Private Sub chkGannacciMultiply_Click()
    
    Dim bEnabled As Boolean
    
    If Me.Visible Then
        If chkGannacciMultiply.Value = vbChecked Then bEnabled = True
        
        txtGannacciMultiply.Enabled = bEnabled
        lblGannacciMultiply.Enabled = bEnabled
        
        chkDecimals.Enabled = bEnabled
        txtDecimals.Enabled = bEnabled
        lblDecimals.Enabled = bEnabled
    
        lblPrice1.Enabled = bEnabled
        lblPrice2.Enabled = bEnabled
        lblPrice3.Enabled = bEnabled
        
        Repaint
    End If
    
End Sub

Private Sub chkGannacciTimeBars_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        If chkGannacciTimeBars.Value = vbUnchecked Then
            If chkGannacciTimeCalendar.Value = vbUnchecked Then
                InfBox "One of the options: Number of bars or Number of calendar days must be selected.", "I", "Ok", "Gannacci Time Cycle"
                chkGannacciTimeBars.Value = vbChecked
            Else
                Repaint
            End If
        Else
            Repaint
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkGannacciTimeBars_Click", eGDRaiseError_Show

End Sub

Private Sub chkGannacciTimeCalendar_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        If chkGannacciTimeCalendar.Value = vbUnchecked Then
            If chkGannacciTimeBars.Value = vbUnchecked Then
                InfBox "One of the options: Number of bars or Number of calendar days must be selected.", "I", "Ok", "Gannacci Time Cycle"
                chkGannacciTimeCalendar.Value = vbChecked
            Else
                Repaint
            End If
        Else
            Repaint
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkGannacciTimeCalendar_Click", eGDRaiseError_Show

End Sub

Private Sub chkIncludeFirstBar_Click()
    Repaint
End Sub

Private Sub chkLucasNumbers_Click()
On Error GoTo ErrSection:

    Dim eZone As eGBZoneInUse

    If Not m.Annot Is Nothing Then
        If chkLucasNumbers.Value = vbChecked Then
            eZone = eANNOT_GB_ZoneLucas
        Else
            eZone = eANNOT_GB_ZoneFib
        End If
        If eZone <> m.Annot.ZoneInUse Then
            m.Annot.ZoneInUse = eZone
            Repaint
            SetTimeZoneGrid
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkLucasNumbers_Click", eGDRaiseError_Show

End Sub

Private Sub chkPatternName_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkPatternName.Click", eGDRaiseError_Show

End Sub

Private Sub chkPointArc_Click()
    Repaint
End Sub

Private Sub chkPriceOnly_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkPriceOnly_Click", eGDRaiseError_Show

End Sub

Private Sub chkProfitLossPercent_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkProfitLossPercent_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkQuadrant_Click(Index As Integer)
On Error GoTo ErrSection:
            
    Dim bRepaint As Boolean
                
    bRepaint = True
        
    If chkQuadrant(Index).Value = vbUnchecked Then
        If chkQuadrant(0).Value = vbUnchecked And _
           chkQuadrant(1).Value = vbUnchecked And _
           chkQuadrant(2).Value = vbUnchecked And _
           chkQuadrant(3).Value = vbUnchecked Then
                
            bRepaint = False
            chkQuadrant(Index).Value = vbChecked
        End If
    End If
                
    If bRepaint Then Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkQuadrant.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkPreIndicator_Click()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkPreIndicator.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkR50_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkR50.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkFib_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkFib.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkFibValues_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkFibValues.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkQtrLines_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkQtrLines.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkShowProfitLost_Click()
    Repaint
End Sub

Private Sub chkFreeFloat_Click()
On Error GoTo ErrSection:
    
    If Not Me.Visible Then Exit Sub
    
    With m.Annot
        If .eType = eANNOT_DNExpansion Or .eType = eANNOT_DNExpansion2 Or .eType = eANNOT_DNExpansion3 Or _
           .eType = eANNOT_DNExpansion4 Or .eType = eANNOT_FibABCD Then
        
            HandleFibRepaint
        
        End If
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnotExt.chkFreeFloat_Click"

End Sub

Private Sub chkShowValues_Click()
    Repaint
End Sub

Private Sub chkTextBorder_Click()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.chkTextBorder.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkShowMarkers_Click()
    Repaint
End Sub

Private Sub chkShowNote_Click()
    Repaint
End Sub

Private Sub chkShowText_Click()
    Repaint
End Sub

Private Sub chkShowTextBorder_Click()
    Repaint
End Sub

Private Sub chkTextNextToMain_Click()
    Repaint
End Sub

Private Sub clrDneA_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDneA.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrDneB_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDneB.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrDneC_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDneC.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrDneCOP_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDneCOP.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrDneOP_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDneOP.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrDneXOP_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDneXOP.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrDnr1_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDnr1.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrDnr2_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDnr2.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrDnr3_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDnr3.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrDnrArc_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrDnrArc.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrR50_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrR50.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrColor_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrColor.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrExt_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrExt.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrFib_Changed()
On Error GoTo ErrSection:

    Dim nColor As Long
    
    clrFib.Visible = False
    nColor = clrFib.Color
    If nColor = 0 Then nColor = -1  '0 is reserved color in flex grid control
    
    With fgFib
        If m.nFibColDown > 1 And m.nFibColDown < .Cols Then
            If .TextMatrix(0, m.nFibColDown) = "Color" Then
                If m.nFibRowDown >= .FixedRows And m.nFibRowDown < .Rows Then
                    .Cell(flexcpBackColor, m.nFibRowDown, m.nFibColDown) = nColor
                    .Select 0, 0
                End If
            End If
        End If
    End With
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.clrFib.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrFib_ColorClicked()

    clrFib.Visible = False
    fgFib.Select 0, 0
    
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    Dim dRatio#
    
    dRatio = ValOfText(InfBox("Enter new fibonacci ratio:", "", "", "", False, 0, -1, 0, "", "New Ratio"))
    
    If m.Annot.eType = eANNOT_FibArcs Then
        If dRatio = 0 Or dRatio = 1 Or dRatio < 0 Then
            InfBox "Ratio cannot be negative, 0 or 1.", "I", "Ok", "New Ratio"
            GoTo ErrExit
        End If
    ElseIf m.Annot.eType = eANNOT_FibTimeRatio Then
        If dRatio = 0 Or dRatio = -1 Then
            InfBox "Ratio cannot be 0 or -1.", "I", "Ok", "New Ratio"
            GoTo ErrExit
        End If
    ElseIf dRatio = 0 Or dRatio = 1 Then
        InfBox "Ratio cannot be 0 or 1.", "I", "Ok", "New Ratio"
        GoTo ErrExit
    End If
    
    If dRatio <> 0 Then
        fgFib.AddItem "1" & vbTab & CStr(dRatio)
        
        'aardvark 6279
        If m.Annot.eType <> eANNOT_DNExpansion And m.Annot.eType <> eANNOT_FibABCD And _
           m.Annot.eType <> eANNOT_DNExpansion2 And m.Annot.eType <> eANNOT_DNExpansion3 And _
           m.Annot.eType <> eANNOT_DNExpansion4 Then m.Annot.geMoveFlag = 1
        
        Repaint
        
        If m.Annot.eType <> eANNOT_DNExpansion And m.Annot.eType <> eANNOT_FibABCD And _
           m.Annot.eType <> eANNOT_DNExpansion2 And m.Annot.eType <> eANNOT_DNExpansion3 And _
           m.Annot.eType <> eANNOT_DNExpansion4 Then m.Annot.geMoveFlag = 0
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cmdAdd.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    Dim aCopies As New cGdArray
    Dim i&, j&, idx&
                    
    If Not m.Chart Is Nothing Then
        If m.nAnnotIdx <= m.Chart.Annots.Count And m.nAnnotIdx > 0 Then
            Set Annot = m.Chart.Annots(m.nAnnotIdx)
            If Not Annot Is Nothing Then
                'this is quicker than calling the chart's object remove annot routine
                Annot.geRemoveAnnotation (m.Chart.geChartObj)
                m.Chart.Annots.Remove m.nAnnotIdx
                m.Chart.SyncGlobalAnnots Annot, m.bWasMultiChart
            End If
        End If
    End If
    
    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cmdDelete.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:
    
    Dim nStyle&
    
    'set font currently in use
    Me.Font.Name = m.Annot.Prop("FontName")
    Me.Font.Size = Val(m.Annot.Prop("FontSize"))
    Me.Font.Underline = Val(m.Annot.Prop("FontUnderline"))
    nStyle = Val(m.Annot.Prop("FontStyle"))
    Select Case nStyle
        Case 0:
            Me.Font.Italic = False
            Me.Font.Bold = False
        Case 1:
            Me.Font.Italic = False
            Me.Font.Bold = True
        Case 2:
            Me.Font.Italic = True
            Me.Font.Bold = False
        Case 3:
            Me.Font.Italic = True
            Me.Font.Bold = True
    End Select
    
    m.bIgnoreUnload = True
    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        m.Annot.Prop("FontName") = Me.Font.Name
        m.Annot.Prop("FontSize") = Me.Font.Size
        m.Annot.Prop("FontUnderline") = Me.Font.Underline
        
        'style - 0=reg,1=bold,2=italic,3=bold italic
        nStyle = 0
        If Me.Font.Bold = True Then
            If Me.Font.Italic = True Then
                nStyle = 3
            Else
                nStyle = 1
            End If
        ElseIf Me.Font.Italic = True Then
            nStyle = 2
        End If
        
        m.Annot.Prop("FontStyle") = nStyle
        Repaint
    End If
    
ErrExit:
    DoEvents
    m.bIgnoreUnload = False
    Exit Sub
    
ErrSection:
    m.bIgnoreUnload = False
    RaiseError "frmEditAnnotExt.cmdFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdGannacciDebug_Click()

    Dim Str$

    With rtfGannacciDebug
        .Move 0, 0, fraGannacciSwing.Width, fraGannacciSwing.Height
        If .Visible Then
            .Visible = False
            .Enabled = False
            .ZOrder 1
        Else
            .Text = m.Annot.GannacciDebugText
            .Visible = True
            .Enabled = True
            .ZOrder
        End If
    End With

End Sub

Private Sub cmdGoToBaseLine_Click()
On Error GoTo ErrSection:
   
    m.Chart.Form.CenterTheDate m.Annot.dDate(1)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cmdGoToBaseLine.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Dim strMsg$
    
    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdOK
    
    If Not m.Annot Is Nothing Then
        Repaint '(still need this in order to save the changes)
        
        m.Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart
    End If

    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    Dim lRow&
    
    If m.nFibRowDown = -1 Or m.nFibColDown = -1 Then
        InfBox "Please select a ratio from the grid to remove.", "I", , Me.Caption
        GoTo ErrExit
    End If
    
    With fgFib
        If .Row >= .FixedRows And .Row <= .Rows Then
            lRow = .Row
        ElseIf m.nFibRowDown >= .FixedRows And m.nFibRowDown <= .Rows Then
            lRow = m.nFibRowDown
        End If
        
        If lRow >= .FixedRows And lRow <= .Rows Then
            If .Cell(flexcpBackColor, lRow) = ALT_GRID_ROW_COLOR Then
                'aardvark 6531 - should not allow removal of the anchor ratios
                InfBox "This ratio cannot be removed. The 'Use' checkbox in the grid can be unchecked to prevent it from showing on the chart..", "I", , Me.Caption
            ElseIf lRow = .Rows - 1 And .Cell(flexcpBackColor, lRow - 1) = ALT_GRID_ROW_COLOR Then
                'aardvark 6531 - should not allow removal of last ratio
                InfBox "This ratio cannot be removed. A minimum of " & Str(.Rows - 1) & " ratios is needed for the " & Me.Caption & " to work.", "I", , Me.Caption
            Else
                .RemoveItem lRow
                
                'aardvark 6279
                If m.Annot.eType <> eANNOT_DNExpansion And m.Annot.eType <> eANNOT_FibABCD And _
                   m.Annot.eType <> eANNOT_DNExpansion2 And m.Annot.eType <> eANNOT_DNExpansion3 And _
                   m.Annot.eType <> eANNOT_DNExpansion4 Then m.Annot.geMoveFlag = 1
                
                Repaint
                
                If m.Annot.eType <> eANNOT_DNExpansion And m.Annot.eType <> eANNOT_FibABCD And _
                   m.Annot.eType <> eANNOT_DNExpansion2 And m.Annot.eType <> eANNOT_DNExpansion3 And _
                   m.Annot.eType <> eANNOT_DNExpansion4 Then m.Annot.geMoveFlag = 0
            End If
            
            'aardvark 6531  - need to reset else will keep deleting saved m.nFibRowDown value
            m.nFibRowDown = -1
            m.nFibColDown = -1
        End If
        
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cmdRemove.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdRestoreOrig_Click()
On Error GoTo ErrSection:

    m.bDescending = False
    m.Annot.LoadRatios True, False
    m.Annot.RatiosToFibGrid fgFib, m.bDescending
    
    'aardvark 6279
    If m.Annot.eType <> eANNOT_DNExpansion And m.Annot.eType <> eANNOT_FibABCD And _
       m.Annot.eType <> eANNOT_DNExpansion2 And m.Annot.eType <> eANNOT_DNExpansion3 And _
       m.Annot.eType <> eANNOT_DNExpansion4 Then m.Annot.geMoveFlag = 1
    
    m.Chart.GenerateChart eRedo3_Settings
    
    If m.Annot.eType <> eANNOT_DNExpansion And m.Annot.eType <> eANNOT_FibABCD And _
       m.Annot.eType <> eANNOT_DNExpansion2 And m.Annot.eType <> eANNOT_DNExpansion3 And _
       m.Annot.eType <> eANNOT_DNExpansion4 Then m.Annot.geMoveFlag = 0
    
    m.nFibColDown = -1
    m.nFibColDown = -1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cmdRestoreOrig.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdSaveDefaults_Click()
On Error GoTo ErrSection:
    
    Repaint '(still need this in order to save the changes)
    
    If Not m.Annot Is Nothing Then
        m.Annot.SaveDefaults
        m.Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart      '6334
    End If
    
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.cmdSaveDefaults.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgFib_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrExit
    
    Dim i&
    
    If m.Annot.eType = eANNOT_FibTimeRatio And m.nFibRowDown = 1 And m.nFibColDown = 1 Then
        With fgFib
            If .TextMatrix(1, 1) = "-1" Then
                .TextMatrix(2, 1) = "0"
            Else
                .TextMatrix(2, 1) = "1"
            End If
        End With
        i = fgFib.TextMatrix(2, 1)
        If Val(m.Annot.Prop("RatioOneVal")) <> i Then Repaint
    Else
        Repaint
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.fgFib_AfterEdit", eGDRaiseError_Show

End Sub

Private Sub fgFib_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Row >= fgFib.FixedRows And Row < fgFib.Rows And Col >= fgFib.FixedCols And Col < fgFib.Cols Then
        m.nFibRowDown = Row
        m.nFibColDown = Col
        
        fgFib.Row = Row
        fgFib.Col = Col
    End If

    If Col = 1 Then
        clrFib.Visible = False
        If m.Annot.eType = eANNOT_FibTimeRatio And Row = 1 Then
            fgFib.ComboList = "-1|0"
        ElseIf m.Annot.eType = eANNOT_DNExpansion Or m.Annot.eType = eANNOT_DNExpansion2 Or _
               m.Annot.eType = eANNOT_DNExpansion3 Or m.Annot.eType = eANNOT_DNExpansion4 Or _
               m.Annot.eType = eANNOT_FibABCD Then
            If fgFib.TextMatrix(Row, 1) = "n/a" Then Cancel = True
        ElseIf m.Annot.eType = eANNOT_Pivot Then
            Cancel = True
        ElseIf fgFib.Cell(flexcpBackColor, Row) = ALT_GRID_ROW_COLOR Then
            Cancel = True
        End If
    ElseIf Col = 2 Then
        If m.Annot.eType = eANNOT_Pivot Then
            Cancel = True
        ElseIf fgFib.TextMatrix(0, Col) = "Color" Then
            If fgFib.Cell(flexcpBackColor, Row) = ALT_GRID_ROW_COLOR Then
                clrFib.Visible = False
                Cancel = True
            End If
        Else
            clrFib.Visible = False
            If m.Annot.eType = eANNOT_DNExpansion Or m.Annot.eType = eANNOT_DNExpansion2 Or _
               m.Annot.eType = eANNOT_DNExpansion3 Or m.Annot.eType = eANNOT_DNExpansion4 Then
                If fgFib.Cell(flexcpBackColor, Row) <> ALT_GRID_ROW_COLOR Then Cancel = True        '6144
            ElseIf m.Annot.eType = eANNOT_FibABCD Then
                If Row > 3 Then Cancel = True
            ElseIf fgFib.TextMatrix(Row, 1) <> "0" And fgFib.TextMatrix(Row, 1) <> "1" Then
                Cancel = True
            End If
        End If
    ElseIf Col = 4 Then
        If m.Annot.eType = eANNOT_Gartley And Row < 6 Then Cancel = True
    ElseIf Col = 5 Then
        clrFib.Visible = False
        If m.Annot.eType <> eANNOT_Gartley Then
            Cancel = True
        ElseIf Row > 5 Then
            Cancel = True
        End If
    ElseIf Col <> 4 Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.fgFib_BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgFib_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim nRow&, nCol&, i&

    With fgFib
        nRow = .MouseRow
        nCol = .MouseCol
        If nRow >= .FixedRows And nRow < .Rows Then
            If nCol = 1 Then
                .Row = nRow
                .Col = nCol
                If m.Annot.eType <> eANNOT_Pivot Then
                    If fgFib.TextMatrix(nRow, 1) <> "0" And fgFib.TextMatrix(nRow, 1) <> "1" Then
                        .EditCell
                    End If
                End If
            ElseIf nCol = 2 Then
                If .TextMatrix(0, nCol) <> "Color" Then
                    .Row = nRow
                    .Col = nCol
                    If m.Annot.eType <> eANNOT_Pivot Then
                        If fgFib.TextMatrix(nRow, 1) <> "0" Or fgFib.TextMatrix(nRow, 1) <> "1" Then
                            .EditCell
                        End If
                    End If
                End If
            ElseIf nCol = 0 Then
                If m.Annot.eType = eANNOT_Gartley And nRow < 6 Then
                    Cancel = True
                ElseIf m.Annot.eType = eANNOT_FibABCD Then
                    If .TextMatrix(nRow, 1) <> "n/a" Or .TextMatrix(nRow, 1) = "n/a" Then
                        CheckedCell(fgFib, nRow, nCol) = Not CheckedCell(fgFib, nRow, nCol)
                        Repaint
                    End If
                ElseIf m.Annot.eType = eANNOT_DNExpansion Or m.Annot.eType = eANNOT_DNExpansion2 Or _
                       m.Annot.eType = eANNOT_DNExpansion3 Or m.Annot.eType = eANNOT_DNExpansion4 Then
                    If .TextMatrix(nRow, 1) <> "n/a" Or (.TextMatrix(nRow, 1) = "n/a" And Not UseDiNapFib()) Then
                        CheckedCell(fgFib, nRow, nCol) = Not CheckedCell(fgFib, nRow, nCol)
                        Repaint
                    End If
                'ElseIf fgFib.TextMatrix(nRow, 1) <> "0" And fgFib.TextMatrix(nRow, 1) <> "1" Then
                Else
                    CheckedCell(fgFib, nRow, nCol) = Not CheckedCell(fgFib, nRow, nCol)
                    'ratio zero & one define same "line" on fib arc & fan
                    If Not m.Annot Is Nothing Then
                        If m.Annot.eType = eANNOT_FibArcs Or m.Annot.eType = eANNOT_FibFan Then
                            With fgFib
                                If .TextMatrix(nRow, 1) = "0" Then
                                    For i = nRow To .Rows - 1
                                        If .TextMatrix(i, 1) = "1" Then
                                            CheckedCell(fgFib, i, nCol) = CheckedCell(fgFib, nRow, nCol)
                                        End If
                                    Next
                                ElseIf .TextMatrix(nRow, 1) = "1" Then
                                    For i = nRow To .FixedRows Step -1
                                        If .TextMatrix(i, 1) = "0" Then
                                            CheckedCell(fgFib, i, nCol) = CheckedCell(fgFib, nRow, nCol)
                                        End If
                                    Next
                                End If
                            End With
                        End If
                    End If
                    Repaint
                End If
            ElseIf nCol = 4 Then
                If .TextMatrix(0, nCol) <> "Fill" Then .EditCell        'gartley 5th column is checkbox for triangle fill
            ElseIf nCol = 5 Then
                .EditCell           'gartley has 6 columns
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.fgFib_BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgFib_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    clrFib.Visible = False
End Sub

Private Sub fgFib_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:
    
    'let grid do sorting for column 0 (Use column)
    'let annotation object do sorting for column 1 (Ratios column)
    'disallow sorting for all other columns
    If Col <> 0 Then
        Order = flexSortNone
        If Col = 1 And Not m.Annot Is Nothing Then
            m.bDescending = Not m.bDescending
            m.Annot.RatiosToFibGrid fgFib, m.bDescending
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.fgFib_BeforeSort", eGDRaiseError_Show

End Sub

Private Sub fgFib_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    FinishEdit = True
End Sub

Private Sub fgFib_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim i&, j&, strText$
    Dim bReset As Boolean
    
    With fgFib
        If .Row > 0 And .Row < .Rows And .Col > 0 And .Col < .Cols Then
            .Select .Row, .Col
            m.nFibColDown = .Col
            m.nFibRowDown = .Row
            If .TextMatrix(0, .Col) = "Color" Then
                If m.Annot.eType = eANNOT_Gartley And .Row < 6 Then
                    bReset = True
                    GoTo ErrExit
                End If
                strText = .TextMatrix(.Row, 1)
                If m.Annot.eType = eANNOT_DNExpansion Or m.Annot.eType = eANNOT_DNExpansion2 Or _
                   m.Annot.eType = eANNOT_DNExpansion3 Or m.Annot.eType - eANNOT_DNExpansion4 Or _
                   m.Annot.eType = eANNOT_FibABCD Or _
                   (strText <> "0" And strText <> "1") Or _
                   (strText = "1" And (m.Annot.eType = eANNOT_ElliotTimeRatio Or m.Annot.eType = eANNOT_FibTimeRatio Or m.Annot.eType = eANNOT_BalloonStrangle)) Then
                    
                    If .TopRow > .FixedRows Then
                        i = .Row - .TopRow + .FixedRows
                    Else
                        i = .Row
                    End If
                    
                    If m.Annot.eType = eANNOT_Gartley Then
                        j = .ColWidth(.Cols - 1) * 2 + 195
                    ElseIf .Col = .Cols - 1 Then
                        j = 10
                    Else
                        j = 10 + .ColWidth(.Col)    'FibDnEx has label column
                    End If
                    clrFib.Move .Left + .ClientWidth - clrFib.Width - j, .Top + .RowHeight(.Row) * i
                    clrFib.Color = .Cell(flexcpBackColor, .Row, .Col)
                    clrFib.Visible = True
                End If
            End If
        Else
            bReset = True
        End If
    End With

ErrExit:
    
    If bReset Then
        m.nFibColDown = -1
        m.nFibRowDown = -1
        clrFib.Visible = False
    End If

End Sub

Private Sub fgFib_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If clrFib.Visible And m.nFibColDown > 0 And m.nFibRowDown > 0 And Not clrFib.DropDownVisible Then
        clrFib.UserControl_Click
    End If

End Sub

Private Sub fgGannacciSwing_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Row >= fgGannacciSwing.FixedRows And Row < fgGannacciSwing.Rows Then
        If Col = 0 Then Repaint
    End If
    
End Sub

Private Sub fgGannacciSwing_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Row < fgGannacciSwing.FixedRows Or Row > fgGannacciSwing.Rows Or Col <> 0 Then
        Cancel = True
    End If
    
End Sub

Private Sub fgTimeLines_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Repaint
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.fgTimeLines.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgTimeLines_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col = 0 Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.fgTimeLines.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgTimeLines_DblClick()
On Error GoTo ErrSection:

    Dim dDate#
    Dim SelAnnot As cAnnotation
       
    If m.Annot.eType = eANNOT_FibTimeZones Or m.Annot.eType = eANNOT_DanCodeZone Then
        With fgTimeLines
            dDate = .RowData(fgTimeLines.Row)
        End With
        m.Chart.Form.CenterTheDate dDate
    ElseIf m.Annot.eType = eANNOT_Pattern Then
        With fgTimeLines
            If .TextMatrix(.Row, 0) = "Original" Then
                dDate = m.Annot.DateFromArray(1)
            ElseIf .TextMatrix(.Row, 0) = "Copy" Then
                dDate = m.Annot.dDate(2)
            End If
            m.Chart.Form.CenterTheDate dDate
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.fgTimeLines.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Activate()

    DoEvents
    m.bIgnoreUnload = False

End Sub

Private Sub Form_Click()
On Error GoTo ErrSection:

    ' acts like "apply"
    If m.bTextChanged Then Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.Form.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    'Set m.Annot = Nothing
    'Set m.Chart = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.Form.Deactivate", eGDRaiseError_Show
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
    RaiseError "frmEditAnnotExt.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    cmdCancel.Top = -cmdCancel.Height * 2

    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If m.bIgnoreUnload Then
        Cancel = True
    ElseIf UnloadMode = 0 Then
        Cancel = True
        'cmdOK_Click
        tmrEditAnnot.Enabled = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
    If m.bCenterColorStyle Then CenterColorStyle
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

On Error GoTo ErrSection:

    If Not m.Annot Is Nothing Then
        If m.Annot.eType = eANNOT_BalloonStrangle Then
            If Not m.Chart Is Nothing Then
                If Not m.Chart.Form Is Nothing Then
                    m.Chart.Form.SyncDrawTools
                End If
            End If
        End If
    End If

ErrExit:
    Set m.Annot = Nothing
    Set m.Chart = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Sub Edit(Chart As cChart, ByVal nAnnotIdx&)
On Error GoTo ErrSection:
    
    Dim strText$, i&, eUsage As eAnnotUsage
    Dim FibTimeAnnot As cAnnotation
    
    m.bIgnoreUnload = False
    m.bCenterColorStyle = False
    If FormIsLoaded("frmChartCfg") Then
        If Not frmChartCfg.bNowAdding Then Unload frmChartCfg
    End If
   
    If nAnnotIdx <= 0 Or nAnnotIdx > Chart.Annots.Count Then Exit Sub
    
    eUsage = Chart.Annots(nAnnotIdx).eUsage
    If eUsage <> eANNOT_UserAdded Then Exit Sub
    
    DoEvents
    
    Set m.Chart = Chart
    Set m.Annot = Chart.Annots(nAnnotIdx)
    If m.Annot Is Nothing Then Exit Sub
     
    m.nAnnotIdx = m.Annot.geAnnId
    
    If m.Chart.SymbolID > 0 Then
        chkMultiChart.Caption = "Show for " & m.Chart.Symbol & " in all chart windows"
    ElseIf Len(m.Chart.SpreadSymbols) > 0 Then
        chkMultiChart.Caption = "Show in all chart windows for this spread"
    Else
        chkMultiChart.Caption = "Show in all chart windows for this symbol"
    End If
    
    HideOnInitialShow
    
    m.bDescending = False
    
    'Developer Note: there are 2 general types of subroutines: Init(xxx)Controls and
    '   Set(xxx)Controls. The "Init" type subroutines are intended to handle setting
    '   control values that do not need to be changed regardless of users input.
    '   The "Set" type subroutines are called to update controls values as users
    '   add/change/modify options.
    
    With m.Annot
        m.bWasMultiChart = .MultiChartFlag
        'set values for controls common to all annotations
        clrColor.Color = .Color
        chkPreIndicator.Value = .PreIndicator
        'show multichart option only if annotation is in price pane AND does not have alert
        If Chart.Tree.Key(m.Annot.gePaneId) = "PRICE PANE" And .AlertObject Is Nothing Then
            chkMultiChart.Value = Abs(.MultiChartFlag)
            m.bMultiChartOption = True
        Else
            m.bMultiChartOption = False
        End If
        chkMultiChart.Visible = m.bMultiChartOption
        
        LoadAnnotPenstyle cboStyle
        cboStyle.Width = clrColor.Width
        SetAnnotPenstyleCombo cboStyle, .Style
        
        ' other properties
        Select Case .eType
            Case eANNOT_DNExpansion
                If UseDiNapFib() Then
                    InitDneControls
                Else
                    InitFibControls
                End If

            Case eANNOT_DNExpansion2, eANNOT_DNExpansion3, eANNOT_DNExpansion4
                InitFibControls
            
            Case eANNOT_FibArcs, eANNOT_Gartley, eANNOT_FibABCD, _
                 eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                 eANNOT_Fibonacci4, eANNOT_FibExpansion, eANNOT_DanCodeFib, _
                 eANNOT_FibFan, eANNOT_FibTimeRatio, eANNOT_ElliotTimeRatio, _
                 eANNOT_BalloonStrangle, eANNOT_AdvRiskReward
                
                InitFibControls
            
            Case eANNOT_FibTimeZones, eANNOT_AndrewFork, eANNOT_DanCodeZone     '6571 (andrew fork got removed by mistake - thought not needed)
                
                InitFibForkControls
            
            Case eANNOT_DNRetracement
                InitDnrControls
            Case eANNOT_TimeCycle
                InitTimeCycleControls
            Case eANNOT_SpResistFan
                InitSpResistFan
            Case eANNOT_Pivot
                InitPivotControls
            
            Case eANNOT_GannacciCycle
                InitGannacciCycle
            Case eANNOT_GannacciTime
                InitGannacciTime
            Case eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
                InitGannacciSwing
        End Select
    End With
        
    If m.Annot.eType = eANNOT_DNExpansion And UseDiNapFib() Then
        Me.Width = fraDNE.Width
    Else
        Select Case m.Annot.eType
            Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                 eANNOT_Fibonacci4, eANNOT_FibTimeRatio, eANNOT_DanCodeFib, _
                 eANNOT_FibExpansion, eANNOT_AdvRiskReward, eANNOT_FibFan, eANNOT_FibArcs, _
                 eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, eANNOT_DNExpansion4, _
                 eANNOT_ElliotTimeRatio, eANNOT_FibABCD, eANNOT_BalloonStrangle
                 
                Me.Width = fraFib.Width + 280
                fraButtons.Left = fraButtons.Left + 150

            Case eANNOT_Pivot
                Me.Width = fraFib.Width + 220
                fraButtons.Left = fraButtons.Left + 950
        
            Case eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
                Me.Width = fraGannacciSwing.Width + 450
        
            Case Else
                Me.Width = fraButtons.Left * 2 + fraButtons.Width + Me.Width - Me.ScaleWidth
        End Select
    End If
        
    CenterFormOnChart Me, m.Chart        '6434
    ShowForm Me, , , , ALT_GRID_ROW_COLOR
    m.bTextChanged = False          '4234
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.Edit", eGDRaiseError_Raise
    
End Sub

Private Sub Repaint()
On Error GoTo ErrSection:
    
    Dim strText$
    Dim i&, j&
    Dim d#, dY1#, dY2#
    
    Dim aValues As New cGdArray
    Dim aDates As New cGdArray
    
    Dim strRatios$, strShow$, strColor$
    
    If Not Me.Visible Then Exit Sub
    If m.Chart Is Nothing Then Exit Sub
    If m.Annot Is Nothing Then Exit Sub
    
    With m.Annot
        ' main color, style & pre-indicator flag
        .Color = clrColor.Color
        .Style = cboStyle.ItemData(cboStyle.ListIndex)
        .PreIndicator = chkPreIndicator.Value
        If chkMultiChart.Value = 1 Then
            .MultiChartFlag = True
        Else
            .MultiChartFlag = False
        End If
        
        ' other properties
        Select Case m.Annot.eType
            Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, _
                eANNOT_Fibonacci4, eANNOT_FibTimeRatio, eANNOT_FibArcs, _
                eANNOT_FibExpansion, eANNOT_FibFan, eANNOT_ElliotTimeRatio, _
                eANNOT_SpResistFan, eANNOT_DanCodeFib, eANNOT_AdvRiskReward
                
                If (.eType = eANNOT_FibExpansion Or .eType = eANNOT_AdvRiskReward) And .Prop("ReverseDirection") <> chkCircular.Value Then
                    .Prop("ReverseDirection") = chkCircular.Value
                End If
                
                HandleFibRepaint
                
                 
            Case eANNOT_FibTimeZones, eANNOT_TimeCycle, eANNOT_DanCodeZone
                .Prop("ShowInAllPanes") = chkAllPanes.Value
                .Prop("FromPrevBar") = Abs(optFromPrevBar.Value)
                
                If .eType = eANNOT_TimeCycle Then
                    .TimeCycleSpace(m.Chart) = Int(ValOfText(txtBarsSpacing.Text))
                    .Prop("Arcs") = Abs(optArcs(1).Value)
                Else
                    .Prop("ShowValues") = chkDynamic.Value
                    .Prop("Arcs") = Abs(optFibZoneArcs(1).Value)
                End If
                FixZoneOptControls
                
            Case eANNOT_FibABCD
                HandleFibRepaint

            Case eANNOT_DNExpansion
                If UseDiNapFib() Then
                    .Prop("ShowHandle") = chkDneCircles
                    'pen color
                    .Prop("colorA") = clrDneA.Color
                    .Prop("colorB") = clrDneB.Color
                    .Prop("colorC") = clrDneC.Color
                    .Prop("colorCOP") = clrDneCOP.Color
                    .Prop("colorOP") = clrDneOP.Color
                    .Prop("colorXOP") = clrDneXOP.Color
                    'pen size
                    .Prop("penSizeA") = CboItem(cboDneStyleA)
                    .Prop("penSizeB") = CboItem(cboDneStyleB)
                    .Prop("penSizeC") = CboItem(cboDneStyleC)
                    .Prop("penSizeCOP") = CboItem(cboDneStyleCOP)
                    .Prop("penSizeOP") = CboItem(cboDneStyleOP)
                    .Prop("penSizeXOP") = CboItem(cboDneStyleXOP)
                    'text
                    .Prop("textA") = txtDneLabelA
                    .Prop("textB") = txtDneLabelB
                    .Prop("textC") = txtDneLabelC
                    .Prop("textCOP") = txtDneLabelCOP
                    .Prop("textOP") = txtDneLabelOP
                    .Prop("textXOP") = txtDneLabelXOP
                    'ratios
                    .Prop("ratioCOP") = Str(ValOfText(txtDneRatioCOP))
                    .Prop("ratioOP") = Str(ValOfText(txtDneRatioOP))
                    .Prop("ratioXOP") = Str(ValOfText(txtDneRatioXOP))
                    'reshow calculated values in case ratios changed
                    SetDneControls
                Else
                    HandleFibRepaint
                End If

            Case eANNOT_DNExpansion2, eANNOT_DNExpansion3, eANNOT_DNExpansion4
                HandleFibRepaint
            
            Case eANNOT_DNRetracement
                .Prop("ShowHandle") = chkDnrCircles
                .Prop("FocusDynamic") = chkDynamic
                'pen color
                .Prop("FibColorR2") = clrDnr1.Color
                .Prop("FibColorR0") = clrDnr2.Color
                .Prop("FibColorR1") = clrDnr3.Color
                .Prop("ArcColorR1") = clrDnrArc.Color
                .Prop("ArcColorR2") = clrDnrArc.Color
                'pen size
                .Prop("FibPenSizeR2") = CboItem(cboDnrStyle1)
                .Prop("FibPenSizeR0") = CboItem(cboDnrStyle2)
                .Prop("FibPenSizeR1") = CboItem(cboDnrStyle3)
                .Prop("ArcPensizeR1") = CboItem(cboDnrStyleArc)
                .Prop("ArcPensizeR2") = CboItem(cboDnrStyleArc)
                'ratios
                .Prop("FibR2") = Str(ValOfText(txtDnrRatio1))
                .Prop("FibR0") = Str(ValOfText(txtDnrRatio2))
                .Prop("FibR1") = Str(ValOfText(txtDnrRatio3))
                'lineage markings
                m.Annot.Text = ""
                For i = 1 To fgDnrLineage.Rows - 1
                    If fgDnrLineage.TextMatrix(i, 2) = "None" Then fgDnrLineage.TextMatrix(i, 2) = ""
                    m.Annot.Text = m.Annot.Text + fgDnrLineage.TextMatrix(i, 2) + "~"
                Next
            
            Case eANNOT_AndrewFork
                With fgTimeLines
                    For i = 1 To .Rows - .FixedRows
                        strText = .TextMatrix(i, 1)
                        If InStr(strText, "^") > 0 Then
                            aValues(i - 1) = m.Chart.Bars.PriceFromString(strText)
                        Else
                            aValues(i - 1) = ValOfText(.TextMatrix(i, 1))   'y-values
                        End If
                        aDates(i - 1) = .RowData(i)
                    Next
                End With
                .geSetAndrewForkChange aValues, aDates
                .Prop("ShowQtrLines") = chkAllPanes

            Case eANNOT_Pivot
                .Prop("PivotShow") = Str(fgFib.Cell(flexcpChecked, 1, 0) * 34)
                .Prop("R1Show") = Str(fgFib.Cell(flexcpChecked, 2, 0) * 34)
                .Prop("R2Show") = Str(fgFib.Cell(flexcpChecked, 3, 0) * 34)
                .Prop("R3Show") = Str(fgFib.Cell(flexcpChecked, 4, 0) * 34)
                .Prop("S1Show") = Str(fgFib.Cell(flexcpChecked, 5, 0) * 34)
                .Prop("S2Show") = Str(fgFib.Cell(flexcpChecked, 6, 0) * 34)
                .Prop("S3Show") = Str(fgFib.Cell(flexcpChecked, 7, 0) * 34)
                
                .Prop("PivotColor") = Str(fgFib.Cell(flexcpBackColor, 1, 3))
                .Prop("R1Color") = Str(fgFib.Cell(flexcpBackColor, 2, 3))
                .Prop("R2Color") = Str(fgFib.Cell(flexcpBackColor, 3, 3))
                .Prop("R3Color") = Str(fgFib.Cell(flexcpBackColor, 4, 3))
                .Prop("S1Color") = Str(fgFib.Cell(flexcpBackColor, 5, 3))
                .Prop("S2Color") = Str(fgFib.Cell(flexcpBackColor, 6, 3))
                .Prop("S3Color") = Str(fgFib.Cell(flexcpBackColor, 7, 3))
                
                If optExt(1) Then
                    .Prop("Ext") = 1
                ElseIf optExt(2) Then
                    .Prop("Ext") = 2
                ElseIf optExt(3) Then
                    .Prop("Ext") = 3
                Else
                    .Prop("Ext") = 0
                End If
                .Prop("ExtStyle") = CboItem(cboExtStyle)
                
                If optPivotText(1) Then
                    .Prop("ShowValues") = 1
                ElseIf optPivotText(2) Then
                    .Prop("ShowValues") = 2
                ElseIf optPivotText(3) Then
                    .Prop("ShowValues") = 3
                Else
                    .Prop("ShowValues") = 0
                End If
                
                .Prop("TextNextToMain") = chkTextNextToMain.Value
                .Prop("CalcMethod") = cboFibStyle.ListIndex + 1
                .Prop("PriceOnly") = chkPriceOnly.Value
            
            Case eANNOT_Gartley
                .Prop("ShowHandle") = chkDneCircles.Value
                .Prop("FreeFloat") = chkFreeFloat.Value
                .Prop("GartleyLabelStyle") = cboGartleyEndpoints.ListIndex
                
                .Color = clrColor.Color
                
                .FibGridToRatios fgFib
            
            Case eANNOT_GannacciCycle
                i = ValOfText(txtGannacciYears.Text)
                If i > 0 Then .Prop("GannacciYears") = i
                .Prop("Ext") = chkAllPanes.Value
            
            Case eANNOT_GannacciTime
                .Prop("ShowTB") = chkGannacciTimeBars.Value
                .Prop("ShowCD") = chkGannacciTimeCalendar.Value
                
                .Prop("Show45") = chk45.Value
                .Prop("Show90") = chk90.Value
                .Prop("Show180") = chk180.Value
                .Prop("Show270") = chk270.Value
                .Prop("Show360") = chk360.Value
                
                .Prop("ColorFor45") = gdGannacci45.Color
                .Prop("ColorFor90") = gdGannacci90.Color
                .Prop("ColorFor180") = gdGannacci180.Color
                .Prop("ColorFor270") = gdGannacci270.Color
                .Prop("ColorFor360") = gdGannacci360.Color
                
                .Prop("Ext") = chkAllPanes.Value
            
            Case eANNOT_GannacciSwing1, eANNOT_GannacciSwing2
                .Prop("ShowMarker") = chkShowMarkers.Value
                .Prop("ShowText") = chkShowText.Value
                .Prop("Border") = chkShowTextBorder.Value
                .Prop("IncludeBarOne") = chkIncludeFirstBar.Value
                .Prop("UseMutiplier") = chkGannacciMultiply.Value
                .Prop("RoundDecimals") = chkDecimals.Value
                
                .Prop("SignalColorLow") = gdGannacciLow.Color
                .Prop("SignalColorMedium") = gdGannacciMed.Color
                .Prop("SignalColorHigh") = gdGannacciHigh.Color
                
                If chkGannacciMultiply.Value = vbChecked Then
                    .Prop("MultiplierVal") = ValOfText(txtGannacciMultiply.Text)
                End If
                
                If chkDecimals.Value = vbChecked Then
                    .Prop("Decimals") = Int(ValOfText(txtDecimals.Text))    'presumably user knows this needs to be an integer
                End If
                
                lblPrice1.Caption = "Price 1 = " & .GannacciSwingPrice(1)
                lblPrice2.Caption = "Price 2 = " & .GannacciSwingPrice(2)
                
                If .eType = eANNOT_GannacciSwing1 Then
                    .Prop("R") = fgGannacciSwing.Cell(flexcpChecked, 1, 0)
                
                    .Prop("Diff_TBP2") = fgGannacciSwing.Cell(flexcpChecked, 2, 0)
                    .Prop("Diff_CDP2") = fgGannacciSwing.Cell(flexcpChecked, 3, 0)
                    .Prop("Diff_TBR") = fgGannacciSwing.Cell(flexcpChecked, 4, 0)
                    .Prop("Diff_CDR") = fgGannacciSwing.Cell(flexcpChecked, 5, 0)
                
                    .Prop("Equal_TBP1") = fgGannacciSwing.Cell(flexcpChecked, 6, 0)
                    .Prop("Equal_CDP1") = fgGannacciSwing.Cell(flexcpChecked, 7, 0)
                    .Prop("Equal_TBR") = fgGannacciSwing.Cell(flexcpChecked, 8, 0)
                    .Prop("Equal_CDR") = fgGannacciSwing.Cell(flexcpChecked, 9, 0)
                    .Prop("Equal_TBP2") = fgGannacciSwing.Cell(flexcpChecked, 10, 0)
                    .Prop("Equal_CDP2") = fgGannacciSwing.Cell(flexcpChecked, 11, 0)
                Else
                    lblPrice3.Caption = "Price 3 = " & .GannacciSwingPrice(3)
                
                    .Prop("TB1") = fgGannacciSwing.Cell(flexcpChecked, 1, 0)
                    .Prop("CD1") = fgGannacciSwing.Cell(flexcpChecked, 2, 0)
                    .Prop("TB2") = fgGannacciSwing.Cell(flexcpChecked, 3, 0)
                    .Prop("CD2") = fgGannacciSwing.Cell(flexcpChecked, 4, 0)
                    .Prop("R1") = fgGannacciSwing.Cell(flexcpChecked, 5, 0)
                    .Prop("R2") = fgGannacciSwing.Cell(flexcpChecked, 6, 0)
                    
                    .Prop("Diff_TB1_R2") = fgGannacciSwing.Cell(flexcpChecked, 7, 0)
                    .Prop("Diff_CD1_R2") = fgGannacciSwing.Cell(flexcpChecked, 8, 0)
                    .Prop("Diff_TB2_R1") = fgGannacciSwing.Cell(flexcpChecked, 9, 0)
                    .Prop("Diff_CD2_R1") = fgGannacciSwing.Cell(flexcpChecked, 10, 0)
                    
                    .Prop("Equal_TB1_R2") = fgGannacciSwing.Cell(flexcpChecked, 11, 0)
                    .Prop("Equal_CD1_R2") = fgGannacciSwing.Cell(flexcpChecked, 12, 0)
                    .Prop("Equal_TB2_R1") = fgGannacciSwing.Cell(flexcpChecked, 13, 0)
                    .Prop("Equal_CD2_R1") = fgGannacciSwing.Cell(flexcpChecked, 14, 0)
                End If
            
            Case eANNOT_BalloonStrangle
                .BalloonStockPrice = m.Chart.Bars.PriceFromString(txtBalloonStock.Text)
                .BalloonPutStrike = m.Chart.Bars.PriceFromString(txtBalloonPut.Text)
                .BalloonCallStrike = m.Chart.Bars.PriceFromString(txtBalloonCall.Text)
                .BalloonPutCost = m.Chart.Bars.PriceFromString(txtBalloonCostPut.Text)
                .BalloonCallCost = m.Chart.Bars.PriceFromString(txtBalloonCostCall.Text)
                
                .Prop("RiskBEStyle") = CboItem(cboFibStyle)
                .Prop("ShowPrices") = chkBalloonPrices(0).Value
                .Prop("ShowTradeCost") = chkBalloonPrices(1).Value
                .Prop("ShowPriceMove") = chkBalloonPrices(2).Value
                .Prop("ShowProfitLoss") = chkFibValues.Value
                .Prop("ShowText") = chkShowNote.Value
                .Prop("ShowBE") = chkExtendArcs.Value
                .Prop("RiskBEColor") = gdBalloonColorBE.Color
                
                If chkCircular.Value = vbChecked Then
                    .Prop("TextAlignment") = 6
                Else
                    .Prop("TextAlignment") = 7
                End If
                
                m.Annot.GridToBalloonRatios fgFib
                
                lblBalloonCost.Caption = "Cost of trade: " & Str(m.Annot.BalloonTradeCost)
                m.Annot.BalloonExpiration = gdOptionExpire.Value        '7018
                m.Annot.Text = rtfText.Text
                
        End Select
        m.bTextChanged = False
        
        '.AssignDateTime
        ' Do this since # of points could have changed
        ' (e.g. extensions being toggled)
        m.Chart.GenerateChart eRedo1_Scrolled
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.Repaint", eGDRaiseError_Raise
    
End Sub

Private Sub gdBalloonColorBE_Changed()
    Repaint
End Sub

Private Sub gdGannacci180_Changed()
    Repaint
End Sub

Private Sub gdGannacci270_Changed()
    Repaint
End Sub

Private Sub gdGannacci360_Changed()
    Repaint
End Sub

Private Sub gdGannacci45_Changed()
    Repaint
End Sub

Private Sub gdGannacci90_Changed()
    Repaint
End Sub

Private Sub gdGannacciHigh_Changed()
    Repaint
End Sub

Private Sub gdGannacciLow_Changed()
    Repaint
End Sub

Private Sub gdGannacciMed_Changed()
    Repaint
End Sub

Private Sub gdResistColor_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.gdResistColor.Changed", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optArcs_Click(Index As Integer)
    chkAllPanes.Enabled = optArcs(0).Value
    Repaint
End Sub

Private Sub optAuto_Click()
On Error GoTo ErrSection:

    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.optAuto.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optCenter_Click()
On Error GoTo ErrSection:

    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.optCenter.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optExt_Click(Index As Integer)
On Error GoTo ErrSection:

'    With m.Annot
'        If .eType = eANNOT_Fibonacci Or .eType = eANNOT_FibExpansion Or .eType = eANNOT_DanCodeFib Then
'            SetFibControls
'        End If
'    End With
    
    m.bTextChanged = True '(since only text or extensions)
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.optExt.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optExtDNE_Click(Index As Integer)
    Repaint
End Sub

Private Sub optExtMainDNE_Click(Index As Integer)
    Repaint
End Sub

Private Sub optFibTextLoc_Click(Index As Integer)
    If Me.Visible Then Repaint
End Sub

Private Sub optPivotText_Click(Index As Integer)
    Repaint
End Sub

Private Sub optFibZoneArcs_Click(Index As Integer)
    chkAllPanes.Enabled = optFibZoneArcs(0).Value
    Repaint
End Sub

Private Sub optLeft_Click()
On Error GoTo ErrSection:

    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.optLeft.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optRectangle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.optRectangle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optRight_Click()
On Error GoTo ErrSection:

    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.optRight.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optRounded_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.optRounded.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optSRLeft_Click()
    Repaint
End Sub

Private Sub optSRRight_Click()
    Repaint
End Sub

Private Sub optFromFirstBar_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If m.Annot Is Nothing Then Exit Sub
    
    If m.Annot.ZoneInUse = eANNOT_GB_ZoneFib Then
        Repaint
        SetTimeZoneGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.optFromFirstBar_Click"

End Sub

Private Sub optFromPrevBar_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If m.Annot Is Nothing Then Exit Sub
    
    If m.Annot.ZoneInUse = eANNOT_GB_ZoneFib Then
        Repaint
        SetTimeZoneGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.optFromPrevBar_Click"

End Sub

Private Sub tmrEditAnnot_Timer()
On Error GoTo ErrSection:
    
    tmrEditAnnot.Enabled = False
    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.tmrEditAnnot.Timer", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtAxisLen_LostFocus()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtAxisLen.LostFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtBalloonCall_Change()
    Repaint
End Sub

Private Sub txtBalloonCostCall_Change()
    Repaint
End Sub

Private Sub txtBalloonCostPut_Change()
    Repaint
End Sub

Private Sub txtBalloonPut_Change()
    Repaint
End Sub

Private Sub txtBalloonStock_Change()
    Repaint
End Sub

Private Sub txtBarsSpacing_Change()
On Error GoTo ErrSection:

    If Len(txtBarsSpacing.Text) > 0 Then Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtBarsSpacing.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDecimals_Change()
    Repaint
End Sub

Private Sub txtDneLabelA_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDneLabelA.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelB_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDneLabelB.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelC_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDneLabelC.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelCOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDneLabelCOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDneLabelOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelXOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDneLabelXOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneRatioCOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDneRatioCOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneRatioOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDneRatioOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneRatioXOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDneRatioXOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDnrRatio1_Change()
On Error GoTo ErrSection:

    Repaint
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDnrRatio1.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDnrRatio2_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDnrRatio2.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDnrRatio3_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.txtDnrRatio3.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub SetTimeZoneGrid()
On Error GoTo ErrSection:

    Dim dDate#, nSize&, i&, k&
    Dim aDates As cGdArray
    Dim aLines As cGdArray
    Dim bIntraday As Boolean
    
    If m.Annot Is Nothing Then Exit Sub
    
    Set aDates = New cGdArray
    Set aLines = New cGdArray

    m.Annot.geGetFibTimeLines m.Chart, aDates, aLines
    
    nSize = aDates.Size
        
    If nSize > 0 Then
        bIntraday = m.Chart.Bars.IsIntraday
        With fgTimeLines        'fib time zone
            .Rows = nSize + .FixedRows
            .Editable = flexEDNone
            .TextMatrix(0, 0) = "Line"
            .TextMatrix(0, 1) = "Date"
            dDate = m.Chart.Form.MouseLastDate
            k = -1
            For i = 0 To nSize - 1
                '.TextMatrix(i + 1, 0) = CStr(i + 1)
                .TextMatrix(i + 1, 0) = CStr(aLines(i))
                If bIntraday Then
                    .TextMatrix(i + 1, 1) = DateFormat(aDates(i), MM_DD_YYYY, HH_MM)
                Else
                    .TextMatrix(i + 1, 1) = DateFormat(aDates(i))
                End If
                .RowData(i + 1) = aDates(i)
                If dDate = aDates(i) Then k = i + 1
            Next
            If k > 0 Then
                .TopRow = k
                .Row = k
                .Col = 1
            End If
            fraTimeLines.Caption = "Total Lines(" + CStr(nSize) + ") - DblClk goes to line."
        End With
    End If
    
    Set aDates = Nothing
    Set aLines = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.SetTimeZoneGrid", eGDRaiseError_Raise

End Sub

Private Function ShowValue(ByVal vValue As Variant) As String
On Error GoTo ErrSection:

    Dim iPane&, nAxisLenData&, strValue$
    
    strValue = CStr(vValue)
    
    With m.Annot
        If .Pane = "PRICE PANE" Then
            iPane = m.Chart.Tree.Index(.Pane)
            If iPane > 0 Then
                strValue = m.Chart.PriceDisplay(iPane, vValue)
            End If
        Else
            strValue = Format(vValue, "0.00#")
        End If
    End With

    ShowValue = strValue

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditAnnotExt.ShowValue", eGDRaiseError_Raise
    
End Function

Private Sub SetBottom(ctlBottom As Control)
On Error GoTo ErrSection:

    If ctlBottom.Visible = False Then ctlBottom.Visible = True
    fraButtons.Top = ctlBottom.Top + ctlBottom.Height
    Me.Height = fraButtons.Top + fraButtons.Height + Me.Height - Me.ScaleHeight
    If Me.Visible Then Me.Refresh
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.SetBottom", eGDRaiseError_Raise
    
End Sub

Private Sub SetDneControls()
On Error GoTo ErrSection:

    Dim iPane&
    Dim dA#, DB#, dC#
    Dim dCOP#, dOp#, dXOP#, dRange#

    'Programmer Note:
    '   B is always X2 & Y2 in the annotation class
    '   C is always > in date than B & A is always < in date than B
    '   i.e. A,B & C are always drawn left to right in order
    iPane = m.Chart.Tree.Index(m.Annot.Pane)
    If iPane > 0 Then
        DB = m.Annot.Y(2)
        If m.Annot.X(2) > m.Annot.X(1) Then
            dA = m.Annot.Y(1)
            dC = m.Annot.Y(3)
        Else
            dC = m.Annot.Y(1)
            dA = m.Annot.Y(3)
        End If
        txtDneValueA.Text = m.Chart.PriceDisplay(iPane, dA)
        txtDneValueB.Text = m.Chart.PriceDisplay(iPane, DB)
        txtDneValueC.Text = m.Chart.PriceDisplay(iPane, dC)
        'caculate COP, OP & XOP
        dRange = Abs(dA - DB)
        If Val(m.Annot.Prop("HiLo") = 1) Then
            dCOP = dC + dRange * Val(m.Annot.Prop("ratioCOP"))
            dOp = dC + dRange * Val(m.Annot.Prop("ratioOP"))
            dXOP = dC + dRange * Val(m.Annot.Prop("ratioXOP"))
        Else
            dCOP = dC - dRange * Val(m.Annot.Prop("ratioCOP"))
            dOp = dC - dRange * Val(m.Annot.Prop("ratioOP"))
            dXOP = dC - dRange * Val(m.Annot.Prop("ratioXOP"))
        End If
        txtDneValueCOP.Text = m.Chart.PriceDisplay(iPane, dCOP)
        txtDneValueOP.Text = m.Chart.PriceDisplay(iPane, dOp)
        txtDneValueXOP.Text = m.Chart.PriceDisplay(iPane, dXOP)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.SetDneControls", eGDRaiseError_Raise

End Sub

Private Sub SetFibControls()
On Error GoTo ErrSection:
    
    Dim aFibRatios As New cGdArray, aFibValues As New cGdArray
    Dim aFibShow As New cGdArray
    Dim aFibColor As New cGdArray
    Dim aFibText As New cGdArray
    Dim i&, k&, dRatio#, strLabel$
    
    If m.Annot.eType = eANNOT_FibTimeZones Or m.Annot.eType = eANNOT_DanCodeZone Then Exit Sub
            
    'set extension options
    If m.Annot.eType = eANNOT_Gartley Then
        Enable cmdFont
        SetBottom fraFib
        
        m.Annot.RatiosToFibGrid fgFib, m.bDescending
        clrColor.Color = m.Annot.Color
    Else
        
        If m.Annot.eType = eANNOT_DNExpansion Or m.Annot.eType = eANNOT_DNExpansion2 Or _
           m.Annot.eType = eANNOT_DNExpansion3 Or m.Annot.eType = eANNOT_DNExpansion4 Or _
           m.Annot.eType = eANNOT_FibABCD Then
            Enable cmdFont
        ElseIf m.Annot.eType = eANNOT_Fibonacci Or m.Annot.eType = eANNOT_Fibonacci2 _
            Or m.Annot.eType = eANNOT_Fibonacci3 Or m.Annot.eType = eANNOT_Fibonacci4 _
            Or m.Annot.eType = eANNOT_FibExpansion Or m.Annot.eType = eANNOT_DanCodeFib _
            Or m.Annot.eType = eANNOT_AdvRiskReward Then
            If optExt(0) = True Then
                fraExt.Height = clrExt.Top
            Else
                fraExt.Height = 1335 - cboExtStyle.Height
            End If
            Enable cmdFont
            SetBottom fraExt
        Else
            If chkFibValues.Value = 1 Then
                Enable cmdFont
            Else
                Disable cmdFont
            End If
            SetBottom fraFib
        End If
        
        m.Annot.RatiosToFibGrid fgFib, m.bDescending
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.SetFibControls", eGDRaiseError_Raise

End Sub

Private Sub SetBalloonStrangleControls()
On Error GoTo ErrSection

    Dim Bars As cGdBars

    If Not m.Annot Is Nothing Then
        If Not m.Annot.AnnotChart Is Nothing Then
            Set Bars = m.Annot.AnnotChart.Bars
        End If
    End If
    
    Enable cmdFont
    
    clrColor.Color = m.Annot.Color
    
    lblBalloonCost.Caption = "Cost of trade: "
    gdOptionExpire.Value = m.Annot.BalloonExpiration
    
    If Bars Is Nothing Then
        If m.Annot.BalloonStockPrice > 0 Then
            txtBalloonStock.Text = RoundNum(m.Annot.BalloonStockPrice)
        End If
        If m.Annot.BalloonPutStrike > 0 Then
            txtBalloonPut.Text = RoundNum(m.Annot.BalloonPutStrike)
        End If
        If m.Annot.BalloonCallStrike > 0 Then
            txtBalloonCall.Text = RoundNum(m.Annot.BalloonCallStrike)
        End If
        If m.Annot.BalloonPutCost > 0 Then
            txtBalloonCostPut.Text = RoundNum(m.Annot.BalloonPutCost)
        End If
        If m.Annot.BalloonCallCost > 0 Then
            txtBalloonCostPut.Text = RoundNum(m.Annot.BalloonCallCost)
        End If
        If m.Annot.BalloonTradeCost > 0 Then
            lblBalloonCost.Caption = "Cost of trade: " & RoundNum(m.Annot.BalloonTradeCost)
        End If
    Else
        If m.Annot.BalloonStockPrice > 0 Then
            txtBalloonStock.Text = Bars.PriceDisplay(m.Annot.BalloonStockPrice)
        End If
        If m.Annot.BalloonPutStrike > 0 Then
            txtBalloonPut.Text = Bars.PriceDisplay(m.Annot.BalloonPutStrike)
        End If
        If m.Annot.BalloonCallStrike > 0 Then
            txtBalloonCall.Text = Bars.PriceDisplay(m.Annot.BalloonCallStrike)
        End If
        If m.Annot.BalloonPutCost > 0 Then
            txtBalloonCostPut.Text = Bars.PriceDisplay(m.Annot.BalloonPutCost)
        End If
        If m.Annot.BalloonCallCost > 0 Then
            txtBalloonCostCall.Text = Bars.PriceDisplay(m.Annot.BalloonCallCost)
        End If
        If m.Annot.BalloonTradeCost > 0 Then
            lblBalloonCost.Caption = "Cost of trade: " & Format(Bars.PriceDisplay(m.Annot.BalloonTradeCost), "#,##0.00")
        End If
    End If
    
    If Val(m.Annot.Prop("TextAlignment")) = 6 Then chkCircular.Value = vbChecked

    rtfText.Text = m.Annot.Text
    rtfText.Font.Name = m.Annot.Prop("FontName")
    rtfText.Font.Size = m.Annot.Prop("FontSize")
    rtfText.Font.Italic = m.Annot.Prop("FontUnderline")
    If m.Annot.Prop("FontStyle") = 1 Or m.Annot.Prop("FontStyle") = 3 Then
        rtfText.Font.Bold = True
    Else
        rtfText.Font.Bold = False
    End If
    
    
    chkBalloonPrices(0).Value = Val(m.Annot.Prop("ShowPrices"))
    chkBalloonPrices(1).Value = Val(m.Annot.Prop("ShowTradeCost"))
    chkBalloonPrices(2).Value = Val(m.Annot.Prop("ShowPriceMove"))
    chkFibValues.Value = Val(m.Annot.Prop("ShowProfitLoss"))
    chkExtendArcs.Value = Val(m.Annot.Prop("ShowBE"))
    chkShowNote.Value = Val(m.Annot.Prop("ShowText"))
    
    gdBalloonColorBE.Color = Val(m.Annot.Prop("RiskBEColor"))

    m.Annot.BalloonRatiosToGrid fgFib
    
    SetBottom fraFib

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnotExt.SetFibControls", eGDRaiseError_Raise

End Sub

Private Sub InitFibForkControls()
On Error GoTo ErrSection:
    
    Dim aDates As New cGdArray
    Dim aYs As New cGdArray
    Dim i&, strValue$
    Dim eZone As eGBZoneInUse
    
    'show neeeded frame(s)
    fraTimeLines.Visible = True
    If m.Annot.eType = eANNOT_FibTimeZones Or m.Annot.eType = eANNOT_DanCodeZone Then
        If m.Annot.eType = eANNOT_DanCodeZone Then
            Me.Caption = "Daniel Code Time Cycle"
            Me.Icon = Picture16(ToolbarIcon("ID_DanCodeZone"), , True)
        ElseIf m.Annot.IsGreenBlattTool Then
            Me.Caption = "Greenblatt Time Zone"
            Me.Icon = Picture16(ToolbarIcon("ID_FibTimeZones"), , True)
        Else
            Me.Caption = "Fibonacci Time Zone"
            Me.Icon = Picture16(ToolbarIcon("ID_FibTimeZones"), , True)
        End If
        
        cmdFont.Visible = True
        chkAllPanes.Visible = True
        chkDynamic.Visible = True
        chkDynamic.Caption = "Show numbers next to lines"
        
        eZone = m.Annot.ZoneInUse
        FixZoneOptControls
        
        If m.bMultiChartOption = True Then
            chkAllPanes.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 45
        Else
            chkAllPanes.Move chkPreIndicator.Left, chkPreIndicator.Top + chkMultiChart.Height + 45
        End If
        
        If m.Annot.IsGreenBlattTool Then
            chkLucasNumbers.Visible = False
            chkLucasNumbers.Enabled = False
            
            With fraZoneOptions
                .Visible = True
                .Move chkAllPanes.Left, chkAllPanes.Top + chkAllPanes.Height - 15, fraTimeLines.Width
                chkDynamic.Move .Left + 120, .Top + .Height - 255
                chkDynamic.ZOrder
            End With
            
            With cboGreenBlattZones
                .Visible = True
                .Enabled = True
                .Clear
                .AddItem "Lucas series"
                .AddItem "Fibonacci series"
                .AddItem "144 series"
                .AddItem "Ratio series"
                .AddItem "Square Root series"
                .ListIndex = eZone
                .Move fraZoneOptions.Left, chkDynamic.Top + chkDynamic.Height + 180
                fraTimeLines.Move fraZoneOptions.Left, .Top + .Height + 60
            End With
        ElseIf m.Annot.eType = eANNOT_DanCodeZone Then
            cboGreenBlattZones.Visible = False
            cboGreenBlattZones.Enabled = False
            chkLucasNumbers.Visible = False
            chkLucasNumbers.Enabled = False
            
            fraZoneOptions.Visible = False
            chkDynamic.Move chkAllPanes.Left, chkAllPanes.Top + chkAllPanes.Height + 60
            
            fraTimeLines.Move chkDynamic.Left, chkDynamic.Top + chkDynamic.Height + 100, fraTimeLines.Width, fraTimeLines.Height + fgTimeLines.RowHeight(0)
            fgTimeLines.Height = fgTimeLines.Height + fgTimeLines.RowHeight(0)
            optFibZoneArcs(0).Top = optFibZoneArcs(0).Top + fgTimeLines.RowHeight(0)
            optFibZoneArcs(1).Top = optFibZoneArcs(0).Top
        Else
            cboGreenBlattZones.Visible = False
            cboGreenBlattZones.Enabled = False
            
            With fraZoneOptions
                .Visible = True
                .Move chkAllPanes.Left, chkAllPanes.Top + chkAllPanes.Height - 15, fraTimeLines.Width
                chkDynamic.Move .Left + 120, .Top + .Height - 255
                chkDynamic.ZOrder
            End With
            
            With chkLucasNumbers
                .Visible = True
                .Enabled = True
                If eZone = eANNOT_GB_ZoneLucas Then
                    .Value = vbChecked
                Else
                    .Value = vbUnchecked
                End If
                .Move fraZoneOptions.Left + 30, fraZoneOptions.Top + fraZoneOptions.Height + 120
                fraTimeLines.Move fraZoneOptions.Left, .Top + .Height + 60
            End With
        End If
        
        chkAllPanes.Value = Int(Val(m.Annot.Prop("ShowInAllPanes")))
        i = Int(Val(m.Annot.Prop("ShowValues")))
        chkDynamic.Value = i
        cmdFont.Enabled = i
        optFibZoneArcs(Val(m.Annot.Prop("Arcs"))).Value = True
        chkAllPanes.Enabled = optFibZoneArcs(0).Value
        m.bCenterColorStyle = False
    Else
        Me.Icon = Picture16(ToolbarIcon("ID_AndrewFork"), , True)
        Me.Caption = "Andrews Pitchfork"
        'RH commented out fraTimeLines.BorderStyle = 0
        fraTimeLines.ZOrder 1
        fraTimeLines.Height = fgTimeLines.Height
        chkAllPanes.Visible = True
        chkAllPanes.Caption = "Show quarter lines"
        If m.bMultiChartOption = True Then
            chkAllPanes.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 70
        Else
            chkAllPanes.Move chkPreIndicator.Left, chkPreIndicator.Top + chkPreIndicator.Height + 70
        End If
        fraTimeLines.Move chkAllPanes.Left - 100, chkAllPanes.Top + chkAllPanes.Height - 150
        
        With m.Annot
            aDates(0) = .dDate(1)
            aDates(1) = .dDate(2)
            aDates(2) = .DateFromArray(0)
            aYs(0) = .Y(1)
            aYs(1) = .Y(2)
            aYs(2) = .YFromArray(0)
        End With
        chkAllPanes.Value = Int(Val(m.Annot.Prop("ShowQtrLines")))
        m.bCenterColorStyle = True
        fraTimeLines.Height = fgTimeLines.Height + 100
    End If
    
    SetBottom fraTimeLines
    
    With fgTimeLines
        SetupGrid Me.fgTimeLines, eGridMode_Grid
        .SelectionMode = flexSelectionFree
        .ColAlignment(0) = flexAlignCenterCenter
        .FixedCols = 0
        .FixedRows = 1
        .Cols = 2
    End With
    
    'show control values
    If m.Annot.eType = eANNOT_FibTimeZones Or m.Annot.eType = eANNOT_DanCodeZone Then
        SetTimeZoneGrid
    ElseIf aDates.Size > 0 Then
        With fgTimeLines
            .Rows = aDates.Size + .FixedRows
            .Editable = flexEDKbdMouse
            .Height = .Height - .RowHeight(1)
            .ColWidth(0) = 1500
            .ColAlignment(1) = flexAlignRightCenter
            .TextMatrix(0, 0) = "Date"
            .TextMatrix(0, 1) = "Value"
            For i = 0 To aDates.Size - 1
                strValue = ShowValue(aYs(i))
                .TextMatrix(i + 1, 0) = DateFormat(aDates(i))
                .TextMatrix(i + 1, 1) = strValue
                .RowData(i + 1) = aDates(i)
            Next
            .Sort = flexSortGenericAscending
            .Select .FixedRows, .FixedCols, .Rows - 1, .Cols - 1    'forces a sort
            .Select 1, 1        'reset selection
        End With
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.InitFibForkControls", eGDRaiseError_Raise

End Sub

Private Sub InitFibMultiClicks()
On Error GoTo ErrSection:

    Dim i&

    fraFib.Width = kFraFibDnExWd
    chkFreeFloat.Visible = True
    chkDneCircles.Visible = True
    cmdFont.Visible = True
    
    LoadAnnotPenstyle cboExtStyle
    SetAnnotPenstyleCombo cboExtStyle, Val(m.Annot.Prop("ExtStyle"))
    
    'reposition/resize controls as needed
    cboExtStyle.Move clrExt.Left, clrExt.Top
    lblExtStyle.Move lblExtColor.Left, lblExtColor.Top

    If m.bMultiChartOption = True Then
        fraFib.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 100
    Else
        fraFib.Move fraButtons.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
    End If
    fraButtons.Move fraButtons.Left + 850
    chkFibValues.Visible = False
    
    Select Case m.Annot.eType
        
        Case eANNOT_Gartley
            Me.Caption = "Gartley"
            Me.Icon = Picture16(ToolbarIcon("ID_Gartley"), , True)
            
            chkPreIndicator.Move chkPreIndicator.Left, cboStyle.Top + cboStyle.Height + 100
            chkMultiChart.Move chkMultiChart.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
            chkFreeFloat.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 100
            
            fraFib.Width = kFraFibWd
            
            chkDneCircles.Move chkFreeFloat.Left, chkFreeFloat.Top
            chkFreeFloat.Visible = False
            chkFreeFloat.Enabled = False
        
            chkFibValues.Caption = "Text on right"
            chkFibValues.Value = Int(Val(m.Annot.Prop("TextOnRight")))
            chkFreeFloat.Value = Val(m.Annot.Prop("FreeFloat"))
            chkDneCircles.Value = Val(m.Annot.Prop("ShowHandle"))
        
            With fgFib
                .Cols = 6
                .TextMatrix(0, 0) = "Use"
                .TextMatrix(0, 1) = "Ratio"
                .TextMatrix(0, 2) = "Value"
                .TextMatrix(0, 3) = "Color"
                .TextMatrix(0, 4) = "Fill"
                .TextMatrix(0, 5) = "Text"
                
                .ColWidth(0) = 600
                .ColWidth(1) = 750
                .ColWidth(2) = 930
                .ColWidth(3) = 720
                .ColWidth(4) = 450
                .ColWidth(5) = 600
                .Rows = 10
                .Width = 4500  '4290      'Me.fraGartleyOpt.Width
                .Height = .Rows * .RowHeight(0) + .RowHeight(0) / 3
                .ScrollBars = flexScrollBarNone
                .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
            End With
            
            fraExt.Visible = False
            fraFibExtDNE.Visible = False
            fraMainExtDNE.Visible = False
            
            lblFibStyle.Visible = False
            cboFibStyle.Visible = False
            
            lblGartleyEndpoints.Visible = True
            cboGartleyEndpoints.Visible = True
            
            cmdAdd.Visible = False
            cmdRemove.Visible = False
            cmdRestoreOrig.Visible = False
            
            cboGartleyEndpoints.Clear
            cboGartleyEndpoints.AddItem "None"
            cboGartleyEndpoints.AddItem "Label with text"
            cboGartleyEndpoints.AddItem "Label with text & value"
            
            i = Val(m.Annot.Prop("GartleyLabelStyle"))
            
            If i >= 0 And i < cboGartleyEndpoints.ListCount Then
                cboGartleyEndpoints.ListIndex = i
            Else
                cboGartleyEndpoints.ListIndex = 1
            End If
            
            fgFib.Left = 0
            lblGartleyEndpoints.Move 145, fgFib.Top + fgFib.Height + 165
            cboGartleyEndpoints.Move lblGartleyEndpoints.Width + 45, lblGartleyEndpoints.Top - 15
            
            fraFib.Caption = ""
            'RH commented out fraFib.BorderStyle = 0
            With chkFreeFloat
                fraFib.Move .Left + 180, .Top + .Height, fgFib.Width + 120, fgFib.Height + cboGartleyEndpoints.Height * 2.5
            End With
            
            SetFibControls
            
        Case eANNOT_DNExpansion, eANNOT_DNExpansion2, eANNOT_DNExpansion3, _
             eANNOT_DNExpansion4, eANNOT_FibABCD
            If m.Annot.eType = eANNOT_FibABCD Then
                Me.Caption = "Fibonacci AB=CD"
                Me.Icon = Picture16(ToolbarIcon("ID_FibABCD"), , True)
            ElseIf m.Annot.eType = eANNOT_DNExpansion Then
                Me.Caption = "Fibonacci Extension"
                Me.Icon = Picture16(ToolbarIcon("ID_DNExpansion"), , True)
            ElseIf m.Annot.eType = eANNOT_DNExpansion2 Then
                Me.Caption = "Fibonacci Extension 2"
                Me.Icon = Picture16(ToolbarIcon("ID_DNExpansion2"), , True)
            ElseIf m.Annot.eType = eANNOT_DNExpansion3 Then
                Me.Caption = "Fibonacci Extension 3"
                Me.Icon = Picture16(ToolbarIcon("ID_DNExpansion3"), , True)
            ElseIf m.Annot.eType = eANNOT_DNExpansion4 Then
                Me.Caption = "Fibonacci Extension 4"
                Me.Icon = Picture16(ToolbarIcon("ID_DNExpansion4"), , True)
            End If
        
            clrColor.Visible = False
            lblColor.Visible = False

            lblStyle.Move lblStyle.Left, lblColor.Top
            cboStyle.Move cboStyle.Left, clrColor.Top

            chkPreIndicator.Move chkPreIndicator.Left, cboStyle.Top + cboStyle.Height + 100
            chkMultiChart.Move chkMultiChart.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
            chkFreeFloat.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 100

            chkDneCircles.Move chkFreeFloat.Left + chkFreeFloat.Width, chkFreeFloat.Top
        
            chkFreeFloat.Value = Val(m.Annot.Prop("FreeFloat"))
            chkDneCircles.Value = Val(m.Annot.Prop("ShowHandle"))
                
            FibExtDNECtrls True
        
            With fgFib
                .Cols = 5
                .ColDataType(0) = flexDTBoolean
                .TextMatrix(0, 0) = "Use"
                .TextMatrix(0, 1) = "Ratio"
                .TextMatrix(0, 2) = "Value"
                .TextMatrix(0, 3) = "Color"
                .TextMatrix(0, 4) = "Label"
                .ColWidth(0) = 600
                .ColWidth(1) = 600
                .ColWidth(2) = 1100
                .ColWidth(3) = 300
                .ColWidth(3) = 600
                .ColWidth(4) = 600
                .Width = fraFib.Width - cboFibStyle.Width - lblStyle.Width + 10
                .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
            End With
            clrFib.Width = 650
        
            fraExt.Visible = False
            fraFibExtDNE.Visible = True
            fraMainExtDNE.Visible = True
            
            fraPivotText.Caption = "Show Text"
            fraPivotText.Visible = True
            
            optPivotText(1).Caption = "Left"
            optPivotText(1).Width = optPivotText(2).Width
            optPivotText(2).Caption = "Right"
            optPivotText(2).Left = optPivotText(1).Left + optPivotText(1).Width + 30
            optPivotText(3).Caption = "Values in y-scale (ratios on right)"
            optPivotText(3).Left = optPivotText(2).Left + optPivotText(2).Width + 60
            optPivotText(3).Width = optPivotText(2).Width * 3 + 330
            
            'hide controls inside fraPivotText that are not used
            chkTextNextToMain.Visible = False
            chkPriceOnly.Visible = False
            
            If Int(Val(m.Annot.Prop("ShowValues"))) = 0 Then
                optPivotText(0).Value = True
            Else
                i = Int(Val(m.Annot.Prop("TextOnRight")))
                If i >= 0 And i < 3 Then
                    optPivotText(i + 1).Value = True
                Else
                    optPivotText(1).Value = True    'default to left if invalid
                End If
            End If
            
            SetFibControls
            With fgFib
                If .Rows > .FixedRows Then
                    .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
                End If
            End With
            
            fraFib.Top = chkFreeFloat.Top + chkFreeFloat.Height + 100
            fraFib.Height = fgFib.RowHeight(0) * 9 + fraMainExtDNE.Height * 3
            
            fraMainExtDNE.Move cmdRestoreOrig.Left + 60, fraFib.Top + fraFib.Height - fraMainExtDNE.Height * 3, fgFib.Width + cmdRestoreOrig.Width + 300
            fraFibExtDNE.Move cmdRestoreOrig.Left + 60, fraMainExtDNE.Top + fraMainExtDNE.Height, fraMainExtDNE.Width
            fraPivotText.Move cmdRestoreOrig.Left + 60, fraFibExtDNE.Top + fraFibExtDNE.Height + 75, fraMainExtDNE.Width, fraMainExtDNE.Height
            fraPivotText.ZOrder
            
            'note: there is ciruclar reference to position the various frames relative to each other
            'so need to do this last for extending the fib's frame a few pixels below the text frame
            fraFib.Height = fraFib.Height + 195
            
            SetBottom fraFib
            
        Case Else
            GoTo ErrExit
    
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnotExt.InitFibMultiClicks"
    
End Sub

Private Sub InitFibFanArcTime()
On Error GoTo ErrSection:

    Dim i&

    fraFib.Width = kFraFibWd
    clrColor.Visible = True
    lblColor.Visible = True
    
    With fgFib
        If m.Annot.eType = eANNOT_FibFan Then
            .Cols = 4
            .TextMatrix(0, 2) = "Value"
            .TextMatrix(0, 3) = "Color"
            .ColWidth(0) = 600
            .ColWidth(1) = 600
            .ColWidth(2) = 1100
            .ColWidth(3) = 300
            clrFib.Width = 650
        Else
            .Cols = 3
            .TextMatrix(0, 2) = "Color"
            .ColWidth(0) = 900
            .ColWidth(1) = 1000
            .ColWidth(2) = 900
            clrFib.Width = 900
            .ScrollBars = flexScrollBarVertical
            cmdFont.Visible = True
        End If
    
        .TextMatrix(0, 0) = "Use"
        .TextMatrix(0, 1) = "Ratio"
        .ColDataType(0) = flexDTBoolean
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    
    Select Case m.Annot.eType
        
        Case eANNOT_FibFan
            Me.Caption = "Fibonacci Fan"
            Me.Icon = Picture16(ToolbarIcon("ID_FibFan"), , True)
        
            chkFibValues.Value = Val(m.Annot.Prop("ShowValues"))
            
            'reposition/resize controls as needed
            cboExtStyle.Move clrExt.Left, clrExt.Top
            lblExtStyle.Move lblExtColor.Left, lblExtColor.Top
            If m.bMultiChartOption = True Then
                fraFib.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 100
            Else
                fraFib.Move fraButtons.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
            End If

            chkFibValues.Caption = "Extensions"
            chkFibValues.Value = Val(m.Annot.Prop("Ext"))
            chkFibValues.Move lblFibStyle.Left, cmdAdd.Top + 100, cboFibStyle.Width
            cmdAdd.Visible = False
            cmdRemove.Visible = False
            cmdRestoreOrig.Visible = False
            fgFib.Height = fgFib.RowHeight(0) * 7 + 60
            fraFib.Height = 2100
            m.bCenterColorStyle = True
        
            LoadAnnotPenstyle cboExtStyle
            SetAnnotPenstyleCombo cboExtStyle, Val(m.Annot.Prop("ExtStyle"))
        
        Case eANNOT_FibArcs:
            Me.Caption = "Fibonacci Arcs/Circles"
            Me.Icon = Picture16(ToolbarIcon("ID_FibCircle"), , True)
        
            'chkExtendArcs.Visible = False
            chkCircular.Caption = "As ratios of diagonal line (circular)"
            chkCircular.Visible = True
            fraQuadrants.Visible = True
            fraQuadrants.Width = fraFib.Width
            If m.bMultiChartOption = True Then
                fraQuadrants.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 100
            Else
                fraQuadrants.Move fraButtons.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
            End If
        
            chkFibValues.Value = Val(m.Annot.Prop("ShowValues"))
            fraQuadrants.Caption = "Quadrants"
            chkQuadrant(0).Caption = "Upper Right"
            chkQuadrant(1).Caption = "Lower Right"
            chkQuadrant(2).Caption = "Upper Left"
            chkQuadrant(3).Caption = "Lower Left"
            With m.Annot
                chkQuadrant(0).Value = Val(.Prop("DirNE"))
                chkQuadrant(1).Value = Val(.Prop("DirSE"))
                chkQuadrant(2).Value = Val(.Prop("DirNW"))
                chkQuadrant(3).Value = Val(.Prop("DirSW"))
                'chkExtendArcs.Value = Val(.Prop("Ext"))
                chkCircular.Value = Val(.Prop("Circular"))
            End With
            
            chkCircular.Left = chkExtendArcs.Left
            chkCircular.Width = fgFib.Width
            fraFib.Move fraQuadrants.Left, fraQuadrants.Top + fraQuadrants.Height + 50
        
        Case eANNOT_FibTimeRatio, eANNOT_ElliotTimeRatio
            If m.Annot.eType = eANNOT_ElliotTimeRatio Then
                Me.Caption = "Fibonacci Time Extension"
                Me.Icon = Picture16(ToolbarIcon("ID_ElliotTimeRatio"), , True)
            Else
                Me.Caption = "Fibonacci Time Ratios"
                Me.Icon = Picture16(ToolbarIcon("ID_FibTimeRatio"), , True)
            End If

            chkAllPanes.Visible = True
            chkExtendArcs.Visible = True
            chkExtendArcs.Width = 1900
            chkExtendArcs.Caption = "Show # of bars"
            chkExtendArcs.Value = Val(m.Annot.Prop("ShowNumBars"))
            chkCircular.Visible = True
            chkCircular.Caption = "Show dates"
            chkCircular.Value = Val(m.Annot.Prop("ShowDates"))
            chkFibValues.Caption = "Show Values"
            If m.bMultiChartOption = True Then
                chkAllPanes.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 100
            Else
                chkAllPanes.Move chkPreIndicator.Left, chkPreIndicator.Top + chkMultiChart.Height + 100
            End If
            chkFibValues.Value = Val(m.Annot.Prop("ShowValues"))

            chkAllPanes.Value = Val(m.Annot.Prop("ShowInAllPanes"))
            fraFib.Move fraButtons.Left, chkAllPanes.Top + chkAllPanes.Height + 100
        
        Case Else
            GoTo ErrExit
    
    End Select
    
    SetFibControls
    With fgFib
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        End If
    End With
    
    'reposition/resize controls as needed
    fraButtons.Move fraButtons.Left + 850
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.InitFibFanArcTime", eGDRaiseError_Raise

End Sub

Private Sub InitFibControls()
On Error GoTo ErrSection:
        
    Dim i&, j&, s$
        
    m.bCenterColorStyle = True
    'hide controls within a frame not used by all
    lblExtColor.Visible = False
    clrExt.Visible = False
    chkExtendArcs.Visible = False   'arcs: extend arc, fib retracement/expansion: hide vertical line
    chkCircular.Visible = False     'arcs: keep circle, fib expansion: reverse
    optExt(4).Visible = False
    optExt(4).Enabled = False
    
    lblGartleyEndpoints.Visible = False
    cboGartleyEndpoints.Visible = False
    
    fraFib.Visible = True
    LoadAnnotPenstyle cboFibStyle
    SetAnnotPenstyleCombo cboFibStyle, Val(m.Annot.Prop("FibStyle"))
    clrFib.Color = Val(Parse(m.Annot.Prop("FibColor"), ",", 1))
    With fgFib
        SetupGrid Me.fgFib, eGridMode_Grid
        .SelectionMode = flexSelectionFree
        .Rows = 1
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
    End With
    
    Select Case m.Annot.eType
        Case eANNOT_Fibonacci, eANNOT_Fibonacci2, eANNOT_Fibonacci3, eANNOT_Fibonacci4
            s = ToolbarIcon("ID_Fibonacci")
            If m.Annot.eType = eANNOT_Fibonacci2 Then
                Me.Caption = "Fibonacci Support/Resistance 2"
                s = s & "2"
            ElseIf m.Annot.eType = eANNOT_Fibonacci3 Then
                Me.Caption = "Fibonacci Support/Resistance 3"
                s = s & "3"
            ElseIf m.Annot.eType = eANNOT_Fibonacci4 Then
                Me.Caption = "Fibonacci Support/Resistance 4"
                s = s & "4"
            Else
                Me.Caption = "Fibonacci Support/Resistance"
            End If
            Me.Icon = Picture16(s, , True)
        
        Case eANNOT_FibExpansion
            Me.Caption = "Fibonacci Expansion"
            Me.Icon = Picture16(ToolbarIcon("ID_FibExpansion"), , True)

        Case eANNOT_AdvRiskReward
            Me.Caption = "Advanced Risk Reward"
            Me.Icon = Picture16(ToolbarIcon("ID_AdvRiskReward"), , True)
        
        Case eANNOT_DanCodeFib
            Me.Caption = "Daniel Code Retracement"
            Me.Icon = Picture16(ToolbarIcon("ID_DanCodeFib"), , True)
            fraFib.Caption = "Retracement Lines"
        
        Case eANNOT_FibFan, eANNOT_FibArcs, eANNOT_FibTimeRatio, eANNOT_ElliotTimeRatio
            InitFibFanArcTime
            GoTo ErrExit
        
        Case eANNOT_Gartley, eANNOT_FibABCD, eANNOT_DNExpansion, _
             eANNOT_DNExpansion2, eANNOT_DNExpansion3, eANNOT_DNExpansion4
            InitFibMultiClicks
            GoTo ErrExit
        
        Case eANNOT_BalloonStrangle
            InitBalloonStrangleControls
            GoTo ErrExit
        
        Case Else
            GoTo ErrExit
    End Select
    
    If m.Annot.eType = eANNOT_AdvRiskReward Then
        fraFib.Width = kFraFibDnExWd
    Else
        fraFib.Width = kFraFibWd
    End If
    clrColor.Visible = True
    lblColor.Visible = True
    cmdFont.Visible = True
        
    'show neeed frame(s)
    If m.Annot.eType = eANNOT_FibExpansion Or m.Annot.eType = eANNOT_AdvRiskReward Then
        chkCircular.Caption = "Reverse direction"
        chkCircular.Move chkFibValues.Left, chkFibValues.Top - chkCircular.Height - 50
        chkCircular.Value = Int(Val(m.Annot.Prop("ReverseDirection")))
        chkCircular.Visible = True
    End If
    
    'show control values
    chkExtendArcs.Caption = "Hide diagonal line"
    If m.Annot.eType = eANNOT_Fibonacci Or m.Annot.eType = eANNOT_Fibonacci2 Or m.Annot.eType = eANNOT_Fibonacci3 _
       Or m.Annot.eType = eANNOT_Fibonacci4 Or m.Annot.eType = eANNOT_DanCodeFib Then
        chkExtendArcs.Move chkFibValues.Left, chkFibValues.Top - chkExtendArcs.Height, cmdRestoreOrig.Width + 100
    Else
        chkExtendArcs.Move chkFibValues.Left, chkExtendArcs.Top, cmdRestoreOrig.Width + 100
    End If
    chkExtendArcs.Value = Int(Val(m.Annot.Prop("HideVerticalLine")))
    chkExtendArcs.Visible = True
    chkFibValues.Visible = False
    chkFibValues.Enabled = False
    
    If m.Annot.eType = eANNOT_AdvRiskReward Then
        chkFibTextShow(2).Visible = True
        chkFibTextShow(2).Enabled = True
        chkFibTextShow(3).Visible = True
        chkFibTextShow(3).Enabled = True
        Frame4.Visible = True
        Frame5.Visible = True
        With fgFib
            .Cols = 5
            .ColDataType(0) = flexDTBoolean
            .TextMatrix(0, 0) = "Use"
            .TextMatrix(0, 1) = "Ratio"
            .TextMatrix(0, 2) = "Value"
            .TextMatrix(0, 3) = "Color"
            .TextMatrix(0, 4) = "Label"
            .ColWidth(0) = 600
            .ColWidth(1) = 600
            .ColWidth(2) = 1100
            .ColWidth(3) = 300
            .ColWidth(3) = 600
            .ColWidth(4) = 600
            .Width = fraFib.Width - cboFibStyle.Width - lblStyle.Width + 10
            .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        End With
        clrFib.Width = 650
    Else
        chkFibTextShow(2).Visible = False
        chkFibTextShow(2).Enabled = False
        chkFibTextShow(3).Visible = False
        chkFibTextShow(3).Enabled = False
        Frame4.Visible = False
        Frame5.Visible = False
        With fgFib
            .Cols = 4
            .TextMatrix(0, 2) = "Value"
            .TextMatrix(0, 3) = "Color"
            .ColWidth(0) = 600
            .ColWidth(1) = 600
            .ColWidth(2) = 1100
            .ColWidth(3) = 300
            clrFib.Width = 650
            .TextMatrix(0, 0) = "Use"
            .TextMatrix(0, 1) = "Ratio"
            .ColDataType(0) = flexDTBoolean
            .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        End With
    End If
    
    'reposition/resize controls as needed
    cboExtStyle.Move clrExt.Left - 60, clrExt.Top + 90, cboExtStyle.Width - 120
    lblExtStyle.Move lblExtColor.Left, lblExtColor.Top + 90
    
    If m.bMultiChartOption = True Then
        'multi-chart & keep alive are available only when fib tool drawn in price pane
        chkDynamic.Visible = True
        chkDynamic.Caption = "Dynamic (stay on last swing)"
        fraFib.Move fraButtons.Left, chkDynamic.Top + chkDynamic.Height + 100
        chkDynamic.Value = Val(m.Annot.Prop("KeepAlive"))
'        fraFib.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 100
    Else
        fraFib.Move fraButtons.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
    End If
    
    chkFibTextShow(0).Value = Int(Val(m.Annot.Prop("FibTextRatio")))
    chkFibTextShow(1).Value = Int(Val(m.Annot.Prop("FibTextPrice")))
    
    i = Int(Val(m.Annot.Prop("FibTextRatioLoc"))) * -1
    optFibTextLoc(0).Value = Not i
    optFibTextLoc(1).Value = i
    i = Int(Val(m.Annot.Prop("FibTextPriceLoc")))
    If i = 2 Then
        optFibTextLoc(4).Value = True
    Else
        i = -i
        optFibTextLoc(2).Value = Not i
        optFibTextLoc(3).Value = i
    End If
    
    If m.Annot.eType = eANNOT_AdvRiskReward Then
        Frame2.Left = Frame2.Left + 390
        Frame3.Left = Frame2.Left
        Frame4.Left = Frame2.Left
        Frame5.Left = Frame2.Left
        chkFibTextShow(2).Value = Int(Val(m.Annot.Prop("FibTextPL")))
        chkFibTextShow(3).Value = Int(Val(m.Annot.Prop("FibLabelPL")))
        i = Int(Val(m.Annot.Prop("FibTextPLLoc"))) * -1
        optFibTextLoc(5).Value = Not i
        optFibTextLoc(6).Value = i
        i = Int(Val(m.Annot.Prop("FibLabelPLLoc"))) * -1
        optFibTextLoc(7).Value = Not i
        optFibTextLoc(8).Value = i
        fraFibText.Move fraButtons.Left, fraFib.Top + fraFib.Height + 60, fraFib.Width
    Else
        Me.chkFibTextMain.Top = chkFibTextShow(2).Top
        fraFibText.Move fraButtons.Left, fraFib.Top + fraFib.Height + 60, fraFib.Width, fraFib.Height / 2 + 150
    End If
    
    fraExt.Caption = "Extensions"
    fraExt.Width = fraFib.Width
    fraExt.Move fraFibText.Left, fraFibText.Top + fraFibText.Height + 90
    
    optExt(1).Left = optExt(0).Left + optExt(0).Width + 135
    optExt(2).Left = optExt(1).Left + optExt(1).Width + 135
    optExt(3).Left = optExt(2).Left + optExt(2).Width + 135
    
    optExt(4).Left = optExt(3).Left + optExt(3).Width + 135      '+ 420
    optExt(4).Top = optExt(3).Top
    optExt(4).Visible = True
    optExt(4).Enabled = True
    
    j = Val(m.Annot.Prop("Ext"))
    If j < 0 Or j > 4 Then j = 0
    optExt(j) = True
    
    If j = 4 Then
        If chkFibTextMain.Enabled Then chkFibTextMain.Enabled = False
        chkFibTextMain.Value = 0
    Else
        If Not chkFibTextMain.Enabled Then chkFibTextMain.Enabled = True
        chkFibTextMain.Value = Val(m.Annot.Prop("TextNextToMain"))
    End If

                
    fraFibText.Visible = True
    fraExt.Visible = True
    
    LoadAnnotPenstyle cboExtStyle
    SetAnnotPenstyleCombo cboExtStyle, Val(m.Annot.Prop("ExtStyle"))

    SetFibControls
    
    With fgFib
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        End If
    End With
    
    'reposition/resize controls as needed
    fraButtons.Move fraButtons.Left + 850
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.InitFibControls", eGDRaiseError_Raise

End Sub

Private Sub InitTimeCycleControls()
On Error GoTo ErrSection:
   
    Me.Caption = "Cycle Finder"
    Me.Icon = Picture16(ToolbarIcon("ID_TimeCycle"), , True)
    
    'show neeed frame(s)/control(s)
    chkAllPanes.Visible = True
    fraTimeCycle.Visible = True
    If m.bMultiChartOption = True Then
        chkAllPanes.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 100
    Else
        chkAllPanes.Move chkPreIndicator.Left, chkPreIndicator.Top + chkMultiChart.Height + 100
    End If
    fraTimeCycle.Move chkAllPanes.Left, chkAllPanes.Top + chkAllPanes.Height + 50
    
    'show control values
    fraTimeCycle.Caption = "Options"
    chkAllPanes.Value = Val(m.Annot.Prop("ShowInAllPanes"))
    txtBarsSpacing.Text = Str(m.Annot.TimeCycleSpace)
    txtBaseLine.Text = DateFormat(m.Annot.dDate(1))
    
    optArcs(Val(m.Annot.Prop("Arcs"))).Value = True
    chkAllPanes.Enabled = optArcs(0).Value

    SetBottom fraTimeCycle
    m.bCenterColorStyle = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.InitTimeCycleControls", eGDRaiseError_Raise

End Sub

Private Sub InitDneControls()
On Error GoTo ErrSection:
    
    Me.Icon = Picture16(ToolbarIcon("ID_DNExpansion"), , True)
    If HasModule("FIB") Then
        Me.Caption = "Dinapoli Expansion"
    Else
        Me.Caption = "Fibonacci Extension"
    End If
    
    'hide controls not used by this annotation
    lblColor.Visible = False
    clrColor.Visible = False
    cboStyle.Visible = False
    
    'show needed frame(s) & controls not within a frame
    chkDneCircles.Visible = True
    With fraDNE
'RH commented out         .BorderStyle = 0
        .Move cmdOK.Left, clrColor.Top - 80
    End With
    SetBottom fraDNE
    
    'show controls not within any frame
    cmdFont.Visible = True
    
    'set controls values
    LoadAnnotPenstyle cboDneStyleA
    LoadAnnotPenstyle cboDneStyleB
    LoadAnnotPenstyle cboDneStyleC
    LoadAnnotPenstyle cboDneStyleCOP
    LoadAnnotPenstyle cboDneStyleOP
    LoadAnnotPenstyle cboDneStyleXOP
    
    SetCtl chkDneCircles, Val(m.Annot.Prop("ShowHandle"))
    
    clrDneA.Color = Val(m.Annot.Prop("ColorA"))
    clrDneB.Color = Val(m.Annot.Prop("ColorB"))
    clrDneC.Color = Val(m.Annot.Prop("ColorC"))
    clrDneCOP.Color = Val(m.Annot.Prop("colorCOP"))
    clrDneOP.Color = Val(m.Annot.Prop("colorOP"))
    clrDneXOP.Color = Val(m.Annot.Prop("colorXOP"))
    
    cboDneStyleA.ListIndex = Val(m.Annot.Prop("penSizeA")) '- 1
    cboDneStyleB.ListIndex = Val(m.Annot.Prop("penSizeB")) '- 1
    cboDneStyleC.ListIndex = Val(m.Annot.Prop("penSizeC")) '- 1
    cboDneStyleCOP.ListIndex = Val(m.Annot.Prop("penSizeCOP")) '- 1
    cboDneStyleOP.ListIndex = Val(m.Annot.Prop("penSizeOP")) '- 1
    cboDneStyleXOP.ListIndex = Val(m.Annot.Prop("penSizeXOP")) '- 1
    
    txtDneLabelA.Text = m.Annot.Prop("textA")
    txtDneLabelB.Text = m.Annot.Prop("textB")
    txtDneLabelC.Text = m.Annot.Prop("textC")
    txtDneLabelCOP.Text = m.Annot.Prop("textCOP")
    txtDneLabelOP.Text = m.Annot.Prop("textOP")
    txtDneLabelXOP.Text = m.Annot.Prop("textXOP")
    
    txtDneRatioCOP.Text = Val(m.Annot.Prop("ratioCOP"))
    txtDneRatioOP.Text = Val(m.Annot.Prop("ratioOP"))
    txtDneRatioXOP.Text = Val(m.Annot.Prop("ratioXOP"))
    SetDneControls

    'reposition/resize controls as needed
    chkDneCircles.Move txtDneLabelA.Left + 45, fraDNE.Top + 75
    chkPreIndicator.Move chkDneCircles.Left, chkDneCircles.Top + chkDneCircles.Height + 100
    chkMultiChart.Move chkPreIndicator.Left, chkPreIndicator.Top + chkPreIndicator.Height + 110
    cmdFont.Move txtDneRatioCOP.Left + 100, chkMultiChart.Top - 50
    fraButtons.Left = clrDneA.Left + 300
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.InitDneControls", eGDRaiseError_Raise
    
End Sub

Private Sub InitDnrControls()
On Error GoTo ErrSection:

    Dim i&, nReactPts&
        
    Me.Icon = Picture16(ToolbarIcon("ID_DNRetracement"), , True)
    Me.Caption = "Dinapoli Retracement"
    
    'hide controls within a frame not used by this annotation
    lblColor.Visible = False
    clrColor.Visible = False
    cboStyle.Visible = False
        
    'show needed frame(s)
    With fraDNR
        'RH commented out .BorderStyle = 0
        .Move cmdOK.Left, clrColor.Top - 120
        .ZOrder (1)
    End With
    SetBottom fraDNR
    
    'show controls not within any frame
    cmdFont.Visible = True
    chkDynamic.Visible = True
    
    'set controls values
    LoadAnnotPenstyle cboDnrStyle1
    LoadAnnotPenstyle cboDnrStyle2
    LoadAnnotPenstyle cboDnrStyle3
    LoadAnnotPenstyle cboDnrStyleArc
    
    chkDynamic.Caption = "Dynamic Focus Point"
    SetCtl chkDynamic, Val(m.Annot.Prop("FocusDynamic"))
    SetCtl chkDnrCircles, Val(m.Annot.Prop("ShowHandle"))
    
    clrDnr1.Color = Val(m.Annot.Prop("FibColorR2"))     'default ratio=0.382
    clrDnr2.Color = Val(m.Annot.Prop("FibColorR0"))     'default ratio=1.000
    clrDnr3.Color = Val(m.Annot.Prop("FibColorR1"))     'default ratio=0.618
    clrDnrArc.Color = Val(m.Annot.Prop("ArcColorR1"))   'arc color for ratios 1 & 2
    
    cboDnrStyle1.ListIndex = Val(m.Annot.Prop("FibPenSizeR2")) '- 1
    cboDnrStyle2.ListIndex = Val(m.Annot.Prop("FibPenSizeR0")) '- 1
    cboDnrStyle3.ListIndex = Val(m.Annot.Prop("FibPenSizeR1")) '- 1
    cboDnrStyleArc.ListIndex = Val(m.Annot.Prop("ArcPensizeR1")) '- 1
    
    txtDnrRatio1.Text = Val(m.Annot.Prop("FibR2"))
    txtDnrRatio2.Text = Val(m.Annot.Prop("FibR0"))
    txtDnrRatio3.Text = Val(m.Annot.Prop("FibR1"))
    
    nReactPts = m.Annot.DnrRptsCount - 1
    With Me.fgDnrLineage
        SetupGrid Me.fgDnrLineage, eGridMode_Grid
        .Rows = 1
        .Cols = 3
        .FixedCols = 2
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .TextMatrix(0, 0) = "Reaction"
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Lineage"
        For i = 0 To nReactPts
            .AddItem CStr(i + 1) & vbTab & m.Annot.DnrRptDate(i, True) & vbTab & m.Annot.DnrLineageText(i)
        Next
        .Row = -1
        .ColAlignment(0) = flexAlignCenterCenter
        .ColComboList(2) = "|G|M|T|F|f|d|m|*|None" '???
    End With
    
    'reposition/resize controls as needed
    chkDynamic.Move chkDnrCircles.Left + 55, chkDnrCircles.Top + chkDnrCircles.Height + 90
    chkPreIndicator.Move chkDnrCircles.Left + 55, chkDynamic.Top + chkDynamic.Height + 100
    chkMultiChart.Move chkPreIndicator.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
    cmdFont.Move cboDnrStyle1.Left + cboDnrStyle1.Width - cmdFont.Width, chkDnrCircles.Top - 10
    fraButtons.Left = fgDnrLineage.Left + 110

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.InitDnrControls", eGDRaiseError_Raise
    
End Sub

Private Sub InitBalloonStrangleControls()

    Me.Caption = "Risk Reward Visualizer"
    Me.Icon = Picture16(ToolbarIcon("ID_BalloonStrangle"), , True)
    
    LoadAnnotPenstyle cboFibStyle

    fraFib.Width = kFraFibWd
    fraBalloonStrangle.Width = kFraFibWd
    
    clrColor.Visible = True
    lblColor.Visible = True
    cmdFont.Visible = True

    chkDynamic.Visible = False
    chkDynamic.Enabled = False

    chkBalloonPrices(0).Visible = True
    chkBalloonPrices(0).Enabled = True
    chkBalloonPrices(1).Visible = True
    chkBalloonPrices(1).Enabled = True
    chkBalloonPrices(2).Visible = True
    chkBalloonPrices(2).Enabled = True
    
    chkExtendArcs.Visible = True
    chkExtendArcs.Enabled = True
    chkFibValues.Visible = True
    chkFibValues.Enabled = True
    chkCircular.Visible = True
    chkCircular.Enabled = True
    
    lblFibStyle.Visible = True
    cboFibStyle.Visible = True
    cboFibStyle.Enabled = True
    
    lblGartleyEndpoints.Visible = True      'just reusing this
    gdBalloonColorBE.Visible = True
    gdBalloonColorBE.Enabled = True
    
    cmdAdd.Visible = False
    cmdAdd.Enabled = False
    cmdRemove.Visible = False
    cmdRemove.Enabled = False
    cmdRestoreOrig.Visible = False
    cmdRestoreOrig.Enabled = False
    
    fraBalloonStrangle.Visible = True
    
    fraBalloonStrangle.Move fraButtons.Left, chkDynamic.Top + 60
    fraFib.Move fraButtons.Left, fraBalloonStrangle.Top + fraBalloonStrangle.Height + 60
    fraFib.Caption = "Labels options"
    fraFib.Height = 2995
    
    With fgFib
        .Top = .Top + 60
        .Height = fraFib.Height - 375
        .Cols = 3
        .TextMatrix(0, 0) = "Show"
        .TextMatrix(0, 1) = "Risk Level"
        .TextMatrix(0, 2) = "Color"
        .ColWidth(0) = 600
        .ColWidth(1) = 1110
        clrFib.Width = 1050
        .ColDataType(0) = flexDTBoolean
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    
    chkBalloonPrices(0).Top = fgFib.Top
    chkBalloonPrices(1).Top = chkBalloonPrices(0).Top + chkBalloonPrices(0).Height + 60
    chkBalloonPrices(2).Top = chkBalloonPrices(1).Top + chkBalloonPrices(0).Height + 60
    chkFibValues.Top = chkBalloonPrices(2).Top + chkBalloonPrices(0).Height + 60
    chkExtendArcs.Top = chkFibValues.Top + chkFibValues.Height + 60
    chkCircular.Top = chkExtendArcs.Top + chkExtendArcs.Height + 60
    
    gdBalloonColorBE.Top = chkCircular.Top + chkCircular.Height + 60
    gdBalloonColorBE.Left = cboFibStyle.Left
    gdBalloonColorBE.Width = cboFibStyle.Width
    
    lblGartleyEndpoints.Width = lblFibStyle.Width
    lblGartleyEndpoints.Left = lblFibStyle.Left
    lblGartleyEndpoints.Top = gdBalloonColorBE.Top + 60
    
    lblFibStyle.Top = gdBalloonColorBE.Top + gdBalloonColorBE.Height + 160
    cboFibStyle.Top = gdBalloonColorBE.Top + gdBalloonColorBE.Height + 100
    
    chkBalloonPrices(0).Left = 300
    chkBalloonPrices(1).Left = 300
    chkBalloonPrices(2).Left = 300
    chkExtendArcs.Left = 300
    chkFibValues.Left = 300
    chkCircular.Left = 300
    
    chkFibValues.Width = chkCircular.Width
    
    chkExtendArcs.Caption = "Show B/E"
    chkFibValues.Caption = "Show profit/loss"
    chkCircular.Caption = "Place text on right"
    lblGartleyEndpoints.Caption = "Color:"
    
    SetAnnotPenstyleCombo cboFibStyle, Val(m.Annot.Prop("RiskBEStyle"))
    SetBalloonStrangleControls
    
    With fgFib
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        End If
    End With
    
    fraButtons.Move fraButtons.Left + 850

End Sub

Private Function CboItem(cbo As ctlUniComboImageXP) As Long
On Error GoTo ErrSection:

    If cbo.ListIndex >= 0 And cbo.ListIndex < cbo.ListCount Then    '4346
        CboItem = cbo.ItemData(cbo.ListIndex)
    Else
        CboItem = 0
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditAnnotExt.CboItem", eGDRaiseError_Raise
    
End Function

Private Sub SetQuadrantChkBoxes()
On Error GoTo ErrSection:

    With m.Annot
        If Val(.Prop("DirNE")) = 0 Then
            chkQuadrant(0) = 0
        Else
            chkQuadrant(0) = 1
        End If
        If Val(.Prop("DirSE")) = 0 Then
            chkQuadrant(1) = 0
        Else
            chkQuadrant(1) = 1
        End If
        If Val(.Prop("DirNW")) = 0 Then
            chkQuadrant(2) = 0
        Else
            chkQuadrant(2) = 1
        End If
        If Val(.Prop("DirSW")) = 0 Then
            chkQuadrant(3) = 0
        Else
            chkQuadrant(3) = 1
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.SetQuadrantChkBoxes", eGDRaiseError_Raise
    
End Sub

Private Sub CenterColorStyle()
On Error Resume Next

    Dim nTotalWidth&, nLeft&
    
    lblColor.Width = Me.TextWidth(lblColor.Caption) + 100
        
    nTotalWidth = lblColor.Width + clrColor.Width
    
    nLeft = Me.ScaleLeft + Me.ScaleWidth / 2 - nTotalWidth / 2
    
    'center color controls
    lblColor.Move nLeft
    clrColor.Move lblColor.Left + lblColor.Width
    
    'center style controls
    lblStyle.Move lblColor.Left
    cboStyle.Move clrColor.Left
    
    If cmdFont.Visible Then
        cmdFont.Move Me.ScaleLeft + Me.ScaleWidth - cmdFont.Width - 150
    End If
    

End Sub

Private Sub InitSpResistFan()
On Error GoTo ErrSection:

    Me.Caption = "Speed Resistance Fan"
    Me.Icon = Picture16(ToolbarIcon("ID_SpResistFan"), , True)
    
    'show needed frame(s)
    fraSpResistFan.Visible = True
    If m.bMultiChartOption = True Then
        fraSpResistFan.Top = chkMultiChart.Top + chkMultiChart.Height + 100
    Else
        fraSpResistFan.Top = chkPreIndicator.Top + chkPreIndicator.Height + 100
    End If
    fraSpResistFan.Left = chkPreIndicator.Left - 50
    SetBottom fraSpResistFan
    
    'initialize combo boxes
    cboResistCount.Clear
    cboResistCount.AddItem "2"
    cboResistCount.AddItem "3"
    cboResistCount.AddItem "4"
    cboResistCount.AddItem "5"
    cboResistCount.AddItem "6"
    cboResistCount.AddItem "7"
    cboResistCount.AddItem "8"
    cboResistCount.AddItem "9"
    LoadAnnotPenstyle cboResistStyle
    
    'set controls values
    cboResistCount.ListIndex = Val(m.Annot.Prop("SpResistCount")) - 2
    SetAnnotPenstyleCombo cboResistStyle, Val(m.Annot.Prop("SpResistStyle"))
    gdResistColor.Color = Val(m.Annot.Prop("SpResistColor"))
    
    m.bCenterColorStyle = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.InitSpResistFan", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub HideOnInitialShow()
On Error GoTo ErrSection:

    'hide controls not within a frame
    cmdFont.Visible = False
    chkDynamic.Visible = False
    chkAllPanes.Visible = False
    chkLucasNumbers.Visible = False
    cboGreenBlattZones.Visible = False
    chkFreeFloat.Visible = False
    chkDneCircles.Visible = False
    
    chkBalloonPrices(0).Visible = False
    chkBalloonPrices(0).Enabled = False
    chkBalloonPrices(1).Visible = False
    chkBalloonPrices(1).Enabled = False
    chkBalloonPrices(2).Visible = False
    chkBalloonPrices(2).Enabled = False
    gdBalloonColorBE.Visible = False
    
    'hide all frames except buttons frame
    fraDNE.Visible = False
    fraDNR.Visible = False
    fraExt.Visible = False
    fraFib.Visible = False
    fraFibExtDNE.Visible = False
    fraMainExtDNE.Visible = False
    fraQuadrants.Visible = False
    fraTimeLines.Visible = False
    fraTimeCycle.Visible = False
    fraFibText.Visible = False
    fraPivotText.Visible = False
    fraGannacciCycle.Visible = False
    fraGannacciTime.Visible = False
    fraGannacciSwing.Visible = False
    fraGannacciMultiply.Visible = False
    fraBalloonStrangle.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.HideOnInitialShow", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub HandleFibRepaint()
On Error GoTo ErrSection:

    Dim i&, j&, d#, strRatios$, strShow$
    
    Dim eType As eAnnotType
    
    If m.Annot.eType = eANNOT_SpResistFan Then
        With m.Annot
            .Prop("SpResistColor") = gdResistColor.Color
            .Prop("SpResistStyle") = cboResistStyle.ListIndex
            j = cboResistCount
            For i = 1 To j - 1
                d = i / j
                If i = 1 Then
                    strRatios = Str(RoundNum(d, 2))
                    strShow = "1"
                Else
                    strRatios = strRatios & "," & Str(RoundNum(d, 2))
                    strShow = strShow & ",1"
                End If
            Next
            .Prop("SpResistCount") = j
            .SpeedResistRatiosChange strRatios, strShow
        End With
        
        GoTo ErrExit
    End If
    
    With m.Annot
        .Prop("FibStyle") = CboItem(cboFibStyle)
        
        If .eType = eANNOT_DNExpansion Or .eType = eANNOT_DNExpansion2 Or _
           .eType = eANNOT_DNExpansion3 Or .eType = eANNOT_DNExpansion4 Or _
           .eType = eANNOT_FibABCD Then
            If optPivotText(0).Value = True Then
                .Prop("ShowValues") = 0
            ElseIf optPivotText(1).Value = True Then
                .Prop("ShowValues") = 1
                .Prop("TextOnRight") = 0
            ElseIf optPivotText(2).Value = True Then
                .Prop("ShowValues") = 1
                .Prop("TextOnRight") = 1
            ElseIf optPivotText(3).Value = True Then
                .Prop("ShowValues") = 1
                .Prop("TextOnRight") = 2
            Else
                .Prop("ShowValues") = 1
                .Prop("TextOnRight") = 0    'default to left
            End If
            
            .Prop("ShowHandle") = chkDneCircles
            i = ValOfText(.Prop("FreeFloat"))
            If chkFreeFloat.Value = 0 And i = 1 Then .SnapFreeFloat Me
            .Prop("FreeFloat") = chkFreeFloat.Value
            FibExtDNECtrls False
        ElseIf chkFibValues.Visible Then
            .Prop("ShowValues") = chkFibValues.Value
        End If
        
        If .eType = eANNOT_Fibonacci Or .eType = eANNOT_Fibonacci2 Or .eType = eANNOT_Fibonacci3 _
            Or .eType = eANNOT_Fibonacci4 Or .eType = eANNOT_FibExpansion _
            Or .eType = eANNOT_DanCodeFib Or .eType = eANNOT_AdvRiskReward Then
            
            If optExt(1) Then
                .Prop("Ext") = 1
                .Prop("TextNextToMain") = chkFibTextMain.Value
                If Not chkFibTextMain.Enabled Then chkFibTextMain.Enabled = True
            ElseIf optExt(2) Then
                .Prop("Ext") = 2
                .Prop("TextNextToMain") = chkFibTextMain.Value
                If Not chkFibTextMain.Enabled Then chkFibTextMain.Enabled = True
            ElseIf optExt(3) Then
                .Prop("Ext") = 3
                .Prop("TextNextToMain") = chkFibTextMain.Value
                If Not chkFibTextMain.Enabled Then chkFibTextMain.Enabled = True
            ElseIf optExt(4) Then
                .Prop("Ext") = 4
                .Prop("TextNextToMain") = 0
                If chkFibTextMain.Value = vbChecked Then chkFibTextMain.Value = vbUnchecked
                If chkFibTextMain.Enabled Then chkFibTextMain.Enabled = False
            Else
                .Prop("Ext") = 0
                If Not chkFibTextMain.Enabled Then chkFibTextMain.Enabled = True
            End If
            .Prop("ExtStyle") = CboItem(cboExtStyle)
            
            .Prop("FibTextRatio") = chkFibTextShow(0).Value
            .Prop("FibTextPrice") = chkFibTextShow(1).Value
            .Prop("FibTextRatioLoc") = Abs(Not optFibTextLoc(0).Value)
            If optFibTextLoc(4).Value = True Then
                .Prop("FibTextPriceLoc") = 2
            Else
                .Prop("FibTextPriceLoc") = Abs(Not optFibTextLoc(2).Value)
            End If
            If .eType = eANNOT_AdvRiskReward Then
                .Prop("FibTextPL") = chkFibTextShow(2).Value
                .Prop("FibLabelPL") = chkFibTextShow(3).Value
                .Prop("FibTextPLLoc") = Abs(Not optFibTextLoc(5).Value)
                .Prop("FibLabelPLLoc") = Abs(Not optFibTextLoc(7).Value)
            End If
            
            .Prop("HideVerticalLine") = chkExtendArcs.Value
            
            If Val(.Prop("KeepAlive")) = 0 And chkDynamic.Value = vbChecked Then
                If .FibDynamicValidate Then
                    .Prop("KeepAlive") = 1
                Else
                    chkDynamic.Value = vbUnchecked
                End If
            Else
                .Prop("KeepAlive") = chkDynamic.Value
            End If
        
        ElseIf .eType = eANNOT_FibTimeRatio Or .eType = eANNOT_ElliotTimeRatio Then
            .Prop("ShowInAllPanes") = chkAllPanes.Value
            .Prop("ShowDates") = chkCircular.Value
            .Prop("ShowNumBars") = chkExtendArcs.Value
        ElseIf .eType = eANNOT_FibFan Then
            .Prop("Ext") = chkFibValues.Value
        ElseIf .eType = eANNOT_FibArcs Then
            .Prop("DirNE") = chkQuadrant(0).Value
            .Prop("DirSE") = chkQuadrant(1).Value
            .Prop("DirNW") = chkQuadrant(2).Value
            .Prop("DirSW") = chkQuadrant(3).Value
            '.Prop("Ext") = chkExtendArcs.Value
            .Prop("Circular") = chkCircular.Value
        End If
        
        .FibGridToRatios fgFib
        
        m.Chart.GenerateChart eRedo3_Settings
    End With
    
    SetFibControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.HandleFibRepaint"

End Sub

Private Sub InitPivotControls()
On Error GoTo ErrSection:

    Dim i&
    Dim Bars As cGdBars
    
    Me.Caption = "Pivot Points"
    Me.Icon = Picture16(ToolbarIcon("ID_PivotPoints"), , True)
    
    If Not m.Annot Is Nothing Then
        If Not m.Annot.AnnotChart Is Nothing Then
            Set Bars = m.Annot.AnnotChart.Bars
        End If
    End If

    clrColor.Visible = False
    lblColor.Visible = False
    
    lblStyle.Move lblStyle.Left, lblColor.Top
    cboStyle.Move cboStyle.Left, clrColor.Top
    
    chkPreIndicator.Move chkPreIndicator.Left, cboStyle.Top + cboStyle.Height + 150
    chkMultiChart.Move chkMultiChart.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
    
    'hide controls within a frame not used by this annotation
    lblExtColor.Visible = False
    clrExt.Visible = False
    chkExtendArcs.Visible = False
    chkCircular.Visible = False
    chkFibValues.Visible = False
    cmdAdd.Visible = False
    cmdRemove.Visible = False
    cmdRestoreOrig.Visible = False
    cboGartleyEndpoints.Visible = False
    
    fraFib.Caption = ""
    fraFib.Width = kFraFibWd
    fraPivotText.Width = fraFib.Width
    
    fraExt.Caption = "Extensions"
    fraExt.Width = fraFib.Width
    
    fraPivotText.Caption = "Show pivot labels/values"
    optPivotText(1).Caption = "Label left/value right"
    
    optExt(1).Left = optExt(0).Left + optExt(0).Width + 420
    optExt(2).Left = optExt(1).Left + optExt(1).Width + 420
    optExt(3).Left = optExt(2).Left + optExt(2).Width + 420
    optExt(4).Visible = False
    optExt(4).Enabled = False
    
    lblExtStyle.Move lblExtColor.Left, lblExtColor.Top
    cboExtStyle.Move clrExt.Left, clrExt.Top, cboExtStyle.Width
    
    lblFibStyle.Caption = "Calculation Method"
    lblFibStyle.Left = lblFibStyle.Left + 650
    lblFibStyle.Width = 1500
    lblFibStyle.Top = lblFibStyle.Top - 175
    cboFibStyle.Move lblFibStyle.Left + lblFibStyle.Width, cboFibStyle.Top - 175, cboFibStyle.Width + 750
    
    cboFibStyle.AddItem "Average HLC"
        
    i = m.Annot.AnnotChart.LastGoodDataBar(False)
    If m.Annot.dDate(1) < Bars(eBARS_DateTime, i) And m.Annot.dDate(2) < Bars(eBARS_DateTime, i) Then
        cboFibStyle.AddItem "Average HLC (nbOpen)"
        cboFibStyle.AddItem "Average HL (nbOpen)"
        cboFibStyle.AddItem "Average HL (nbOpen * 2)"
    End If
    
    cmdFont.Move fraFib.Width - cmdFont.Width, chkPreIndicator.Top + 150
    cmdFont.Visible = True
    cmdFont.ZOrder
    
    fgFib.Move lblFibStyle.Left - 650, cboFibStyle.Top + cboFibStyle.Height + 100, fraFib.Width - 250, fraFib.Height - 750
    With fgFib
        SetupGrid Me.fgFib, eGridMode_Grid
        .SelectionMode = flexSelectionFree
        .Rows = 8
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .Cols = 4
        .TextMatrix(0, 0) = "Display"
        .TextMatrix(0, 1) = "Label"
        .TextMatrix(0, 2) = "Value"
        .TextMatrix(0, 3) = "Color"
        .ColWidth(0) = 850
        .ColWidth(1) = 1100
        .ColWidth(2) = 1500
        .ColWidth(3) = 1100
        .ColDataType(0) = flexDTBoolean
        .Cell(flexcpAlignment, .Rows - 1, 0, 0, .Cols - 1) = flexAlignCenterCenter
                
        .TextMatrix(1, 1) = "Pivot"
        .TextMatrix(2, 1) = "R1"
        .TextMatrix(3, 1) = "R2"
        .TextMatrix(4, 1) = "R3"
        .TextMatrix(5, 1) = "S1"
        .TextMatrix(6, 1) = "S2"
        .TextMatrix(7, 1) = "S3"
        
        'annot's aYs[0/S2] [1/S3] [2/R1] [3/R2] [4/R3]
        If Bars Is Nothing Then
            .TextMatrix(1, 2) = m.Annot.Y(1)
            .TextMatrix(2, 2) = m.Annot.YFromArray(2)
            .TextMatrix(3, 2) = m.Annot.YFromArray(3)
            .TextMatrix(4, 2) = m.Annot.YFromArray(4)
            .TextMatrix(5, 2) = m.Annot.Y(2)
            .TextMatrix(6, 2) = m.Annot.YFromArray(0)
            .TextMatrix(7, 2) = m.Annot.YFromArray(1)
        Else
            .TextMatrix(1, 2) = Bars.PriceDisplay(m.Annot.Y(1))
            .TextMatrix(2, 2) = Bars.PriceDisplay(m.Annot.YFromArray(2))
            .TextMatrix(3, 2) = Bars.PriceDisplay(m.Annot.YFromArray(3))
            .TextMatrix(4, 2) = Bars.PriceDisplay(m.Annot.YFromArray(4))
            .TextMatrix(5, 2) = Bars.PriceDisplay(m.Annot.Y(2))
            .TextMatrix(6, 2) = Bars.PriceDisplay(m.Annot.YFromArray(0))
            .TextMatrix(7, 2) = Bars.PriceDisplay(m.Annot.YFromArray(1))
        End If
        
        .Cell(flexcpBackColor, 1, 3) = Val(m.Annot.Prop("PivotColor"))
        .Cell(flexcpBackColor, 2, 3) = Val(m.Annot.Prop("R1Color"))
        .Cell(flexcpBackColor, 3, 3) = Val(m.Annot.Prop("R2Color"))
        .Cell(flexcpBackColor, 4, 3) = Val(m.Annot.Prop("R3Color"))
        .Cell(flexcpBackColor, 5, 3) = Val(m.Annot.Prop("S1Color"))
        .Cell(flexcpBackColor, 6, 3) = Val(m.Annot.Prop("S2Color"))
        .Cell(flexcpBackColor, 7, 3) = Val(m.Annot.Prop("S3Color"))
        
        i = 1
        If Val(m.Annot.Prop("PivotShow")) <> 34 Then i = 2       '1=checked, 2=unchecked
        .Cell(flexcpChecked, 1, 0) = i
        
        i = 1
        If Val(m.Annot.Prop("R1Show")) <> 34 Then i = 2
        .Cell(flexcpChecked, 2, 0) = i
        
        i = 1
        If Val(m.Annot.Prop("R2Show")) <> 34 Then i = 2
        .Cell(flexcpChecked, 3, 0) = i
        
        i = 1
        If Val(m.Annot.Prop("R3Show")) <> 34 Then i = 2
        .Cell(flexcpChecked, 4, 0) = i
        
        i = 1
        If Val(m.Annot.Prop("S1Show")) <> 34 Then i = 2
        .Cell(flexcpChecked, 5, 0) = i
        
        i = 1
        If Val(m.Annot.Prop("S2Show")) <> 34 Then i = 2
        .Cell(flexcpChecked, 6, 0) = i
        
        i = 1
        If Val(m.Annot.Prop("S3Show")) <> 34 Then i = 2
        .Cell(flexcpChecked, 7, 0) = i
                        
    End With
    
    fraFib.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 50
    fraPivotText.Move fraButtons.Left, fraFib.Top + fraFib.Height + 100
    fraExt.Height = 1335 - cboExtStyle.Height
    fraExt.Move fraPivotText.Left, fraPivotText.Top + fraPivotText.Height + 100
    
    fraFib.Visible = True
    fraPivotText.Visible = True
    fraExt.Visible = True
    
    m.bCenterColorStyle = True
                        
'set control values
    'calculation method
    If cboFibStyle.ListCount > 1 Then
        cboFibStyle.ListIndex = Val(m.Annot.Prop("CalcMethod")) - 1
    Else
        cboFibStyle.ListIndex = 0
    End If
    
    'text options
    optPivotText(Int((Val(m.Annot.Prop("ShowValues"))))).Value = True
    chkTextNextToMain.Value = Int(Val(m.Annot.Prop("TextNextToMain")))
    chkPriceOnly.Value = Val(m.Annot.Prop("PriceOnly"))
    
    'extension options
    LoadAnnotPenstyle cboExtStyle
    SetAnnotPenstyleCombo cboExtStyle, Val(m.Annot.Prop("ExtStyle"))
    optExt(Val(m.Annot.Prop("Ext"))) = True
    
    SetBottom fraExt

    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnotExt.InitPivotControls"
    
End Sub

Private Sub FibExtDNECtrls(ByVal bSetCtrls As Boolean)
On Error GoTo ErrSection:
'1=extend right, 2=extend left, 3=extend both direction, 0=extend nothing
    Dim i&
        
    If bSetCtrls Then
        i = Int(Val(m.Annot.Prop("Ext")))
        Select Case i
            Case 1
                optExtDNE(1).Value = True
            Case 2
                optExtDNE(2).Value = True
            Case 3
                optExtDNE(3).Value = True
            Case Default
                optExtDNE(0).Value = True
        End Select
        
        i = Int(Val(m.Annot.Prop("ExtMain")))
        Select Case i
            Case 1
                optExtMainDNE(1).Value = True
            Case 2
                optExtMainDNE(2).Value = True
            Case 3
                optExtMainDNE(3).Value = True
            Case Default
                optExtMainDNE(0).Value = True
        End Select
                
        Exit Sub
    End If
    
    If optExtDNE(1).Value = True Then
        m.Annot.Prop("Ext") = 1
    ElseIf optExtDNE(2).Value = True Then
        m.Annot.Prop("Ext") = 2
    ElseIf optExtDNE(3).Value = True Then
        m.Annot.Prop("Ext") = 3
    Else
        m.Annot.Prop("Ext") = 0
    End If
    
    If optExtMainDNE(1).Value = True Then
        m.Annot.Prop("ExtMain") = 1
    ElseIf optExtMainDNE(2).Value = True Then
        m.Annot.Prop("ExtMain") = 2
    ElseIf optExtMainDNE(3).Value = True Then
        m.Annot.Prop("ExtMain") = 3
    Else
        m.Annot.Prop("ExtMain") = 0
    End If
    
ErrExit:
    Exit Sub
                
ErrSection:
    RaiseError "frmEditAnnotExt.FibExtDNECtrls"

End Sub

Private Sub FixZoneOptControls()
On Error GoTo ErrSection:

    If m.Annot Is Nothing Then Exit Sub
    If m.Annot.ZoneInUse = eANNOT_GB_ZoneDanCode Then Exit Sub
    If m.Annot.eType = eANNOT_TimeCycle Then Exit Sub

    If m.Annot.ZoneInUse = eANNOT_GB_ZoneFib Then
        optFromPrevBar.Enabled = True
        If Val(m.Annot.Prop("FromPrevBar")) = 0 Then
            optFromFirstBar.Value = True
            optFromPrevBar.Value = False
        Else
            optFromFirstBar.Value = False
            optFromPrevBar.Value = True
        End If
    Else
        optFromFirstBar.Value = True
        optFromPrevBar.Value = False
        optFromPrevBar.Enabled = False
    End If

ErrExit:
    Exit Sub
                
ErrSection:
    RaiseError "frmEditAnnotExt.FixZoneOptControls"

End Sub

Private Sub InitGannacciCycle()
On Error GoTo ErrSection:

    Me.Caption = "GANNacci Yearly Cycles"
    Me.Icon = Picture16(ToolbarIcon("ID_GannacciCycle"), , True)
    
    cmdFont.Visible = True
    chkAllPanes.Visible = True
    fraGannacciCycle.Visible = True
    'RH commented out fraGannacciCycle.BorderStyle = 0
    
    chkMultiChart.Top = chkPreIndicator.Top + chkPreIndicator.Height + 120
    chkAllPanes.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 45
    
    fraGannacciCycle.Left = chkMultiChart.Left
    fraGannacciCycle.Top = chkMultiChart.Top + chkMultiChart.Height + chkAllPanes.Height + 120
    
    chkAllPanes.Caption = "Extend vertical line"
    chkAllPanes.Value = Val(m.Annot.Prop("Ext"))
    
    txtGannacciYears.Text = Val(m.Annot.Prop("GannacciYears"))
    
    SetBottom fraGannacciCycle

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnotExt.InitGannacciCycle"

End Sub

Private Sub txtGannacciMultiply_Change()
    Repaint
End Sub

Private Sub txtGannacciYears_Change()
    Repaint
End Sub

Private Sub InitGannacciTime()
On Error GoTo ErrSection:

    Dim bEnabled As Boolean

    Me.Caption = "GANNacci Degree Cycles"
    Me.Icon = Picture16(ToolbarIcon("ID_GannacciTime"), , True)
    
    cmdFont.Visible = True
    chkAllPanes.Visible = True
    fraGannacciTime.Visible = True
    
    chkMultiChart.Top = chkPreIndicator.Top + chkPreIndicator.Height + 120
    chkAllPanes.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 45
    
    fraGannacciTime.Left = chkMultiChart.Left
    fraGannacciTime.Top = chkMultiChart.Top + chkMultiChart.Height + chkAllPanes.Height + 120
    
    'set values
    chk45.Value = Val(m.Annot.Prop("Show45"))
    chk90.Value = Val(m.Annot.Prop("Show90"))
    chk180.Value = Val(m.Annot.Prop("Show180"))
    chk270.Value = Val(m.Annot.Prop("Show270"))
    chk360.Value = Val(m.Annot.Prop("Show360"))
    
    gdGannacci45.Color = Val(m.Annot.Prop("ColorFor45"))
    gdGannacci90.Color = Val(m.Annot.Prop("ColorFor90"))
    gdGannacci180.Color = Val(m.Annot.Prop("ColorFor180"))
    gdGannacci270.Color = Val(m.Annot.Prop("ColorFor270"))
    gdGannacci360.Color = Val(m.Annot.Prop("ColorFor360"))
    
    chkAllPanes.Caption = "Extend vertical line"
    chkAllPanes.Value = Val(m.Annot.Prop("Ext"))
    
    chkGannacciTimeBars.Value = Val(m.Annot.Prop("ShowTB"))
    
    If Not m.Annot Is Nothing Then
        bEnabled = m.Annot.AllowCalendarDays
        chkGannacciTimeCalendar.Enabled = bEnabled
        chkGannacciTimeCalendar.Value = Val(m.Annot.Prop("ShowCD"))
    Else
        chkGannacciTimeCalendar.Enabled = False
        chkGannacciTimeCalendar.Enabled = False
    End If
    
    SetBottom fraGannacciTime

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnotExt.InitGannacciTime"

End Sub

Private Sub InitGannacciSwing()
On Error GoTo ErrSection:

    If m.Annot Is Nothing Then Exit Sub
    
    If m.Annot.eType = eANNOT_GannacciSwing1 Then
        Me.Caption = "GANNacci 1 Swing"
        Me.Icon = Picture16(ToolbarIcon("ID_GannacciSwing1"), , True)
    Else
        Me.Caption = "GANNacci 2 Swing"
        Me.Icon = Picture16(ToolbarIcon("ID_GannacciSwing2"), , True)
    End If
    
    cmdGannacciDebug.Visible = m.Annot.GannacciDebugFlag
    
    cmdFont.Visible = True
    fraGannacciSwing.Visible = True
    fraGannacciMultiply.Visible = True

    With fgGannacciSwing
        SetupGrid fgGannacciSwing, eGridMode_Grid
        .SelectionMode = flexSelectionFree
        .Editable = flexEDKbdMouse
        .ColAlignment(0) = flexAlignCenterCenter
        .FixedCols = 0
        .FixedRows = 1
        .Cols = 2
        
        If m.Annot.eType = eANNOT_GannacciSwing1 Then
            .Rows = 12
            
            .TextMatrix(0, 0) = "Use"
            .TextMatrix(0, 1) = "Item"
            
            .TextMatrix(1, 1) = "R (range)"
            
            .TextMatrix(2, 1) = "TB - P2"
            .TextMatrix(3, 1) = "CD - P2"
            .TextMatrix(4, 1) = "TB - R"
            .TextMatrix(5, 1) = "CD - R"
            
            .TextMatrix(6, 1) = "TB = P1"
            .TextMatrix(7, 1) = "CD = P1"
            .TextMatrix(8, 1) = "TB = R"
            .TextMatrix(9, 1) = "CD = R"
            .TextMatrix(10, 1) = "TB = P2"
            .TextMatrix(11, 1) = "CD = P2"
        Else
            .Rows = 15
            
            .TextMatrix(0, 0) = "Use"
            .TextMatrix(0, 1) = "Item"
            
            .TextMatrix(1, 1) = "TB1 (trading bars)"
            .TextMatrix(2, 1) = "CD1 (calendar days)"
            .TextMatrix(3, 1) = "TB2 (trading bars)"
            .TextMatrix(4, 1) = "CD2 (calendar days)"
            
            .TextMatrix(5, 1) = "R1 (price range)"
            .TextMatrix(6, 1) = "R2 (price range)"
            
            .TextMatrix(7, 1) = "TB1 - R2"
            .TextMatrix(8, 1) = "CD1 - R2"
            .TextMatrix(9, 1) = "TB2 - R1"
            .TextMatrix(10, 1) = "CD2 - R1"
            
            .TextMatrix(11, 1) = "TB1 = R2"
            .TextMatrix(12, 1) = "CD1 = R2"
            .TextMatrix(13, 1) = "TB2 = R1"
            .TextMatrix(14, 1) = "CD2 = R1"
        End If
        
        .ColWidth(0) = 600
        .Height = .RowHeight(0) * 6 - 180
    
        fraGannacciColors.Top = .Top + .Height + 90
    End With
    
    chkMultiChart.Top = chkPreIndicator.Top + chkPreIndicator.Height + 120
    fraGannacciSwing.Left = chkMultiChart.Left
    fraGannacciSwing.Top = chkMultiChart.Top + chkMultiChart.Height + 120
    fraGannacciSwing.Width = fraGannacciSwing.Width
    fraGannacciSwing.Height = fgGannacciSwing.Height + fraGannacciColors.Height + chkShowText.Height * 6 + 30
    
    With fraGannacciMultiply
        .Left = fraGannacciSwing.Left
        .Top = fraGannacciSwing.Top + fraGannacciSwing.Height + 30
        .Width = fraGannacciSwing.Width
        
        If m.Annot.eType = eANNOT_GannacciSwing1 Then
            .Height = .Height - lblPrice3.Height + 15
            lblPrice3.Visible = False
            lblPrice3.Enabled = False
        Else
            lblPrice3.Visible = True
            lblPrice3.Enabled = True
        End If
        
        lblPrice1.Left = chkGannacciMultiply.Left
        lblPrice2.Left = chkGannacciMultiply.Left
        lblPrice3.Left = chkGannacciMultiply.Left
    End With
    
    SetGannacciSwingCtrls
    
    SetBottom fraGannacciMultiply

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnotExt.InitGannacciSwing"

End Sub

Private Sub SetGannacciSwingCtrls()
On Error GoTo ErrSection:
    
    Dim i&
    Dim Bars As cGdBars
    Dim bEnabled As Boolean

    If m.Annot Is Nothing Then Exit Sub
    If m.Annot.AnnotChart Is Nothing Then Exit Sub
    
    Set Bars = m.Annot.AnnotChart.Bars
    If Bars Is Nothing Then Exit Sub
    
    With m.Annot
        clrColor.Color = .Color
        chkShowMarkers.Value = Val(.Prop("ShowMarker"))
        chkShowText.Value = Val(.Prop("ShowText"))
        chkIncludeFirstBar.Value = Val(.Prop("IncludeBarOne"))
        chkShowTextBorder.Value = Val(.Prop("Border"))
            
        i = Val(.Prop("UseMutiplier"))
        If i = 1 Then bEnabled = True
        
        chkGannacciMultiply.Value = i
        txtGannacciMultiply.Text = Val(.Prop("MultiplierVal"))
        
        chkDecimals.Value = Val(.Prop("RoundDecimals"))
        txtDecimals.Text = Val(.Prop("Decimals"))
        
        txtGannacciMultiply.Enabled = bEnabled
        lblGannacciMultiply.Enabled = bEnabled
        
        chkDecimals.Enabled = bEnabled
        txtDecimals.Enabled = bEnabled
        lblDecimals.Enabled = bEnabled
    
        lblPrice1.Enabled = bEnabled
        lblPrice2.Enabled = bEnabled
        lblPrice3.Enabled = bEnabled
        
        'price labels
        lblPrice1.Caption = "Price 1 = " & .GannacciSwingPrice(1)
        lblPrice2.Caption = "Price 2 = " & .GannacciSwingPrice(2)
        If .eType = eANNOT_GannacciSwing2 Then
            lblPrice3.Caption = "Price 3 = " & .GannacciSwingPrice(3)
        End If
        
        'signal strength colors
        gdGannacciLow.Color = Val(.Prop("SignalColorLow"))
        gdGannacciMed.Color = Val(.Prop("SignalColorMedium"))
        gdGannacciHigh.Color = Val(.Prop("SignalColorHigh"))
        
        If .eType = eANNOT_GannacciSwing1 Then
            SetGannacciGridSwing1
        Else
            SetGannacciGridSwing2
        End If

    End With
    
    For i = 1 To fgGannacciSwing.Rows - 1
        fgGannacciSwing.Cell(flexcpPictureAlignment, i, 0) = flexAlignCenterCenter
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnotExt.SetGannacciSwingCtrls"

End Sub

Private Sub SetGannacciGridSwing1()
On Error GoTo ErrSection:

    With m.Annot
        If Val(.Prop("R")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 1, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 1, 0) = flexUnchecked
        End If
        
        'diff properties
        If Val(.Prop("Diff_TBP2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 2, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 2, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Diff_CDP2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 3, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 3, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Diff_TBR")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 4, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 4, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Diff_CDR")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 5, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 5, 0) = flexUnchecked
        End If
        
        'equal properties
        If Val(.Prop("Equal_TBP1")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 6, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 6, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Equal_CDP1")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 7, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 7, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Equal_TBR")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 8, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 8, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Equal_CDR")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 9, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 9, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Equal_TBP2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 10, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 10, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Equal_CDP2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 11, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 11, 0) = flexUnchecked
        End If
    
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnotExt.SetGannacciGridSwing1"

End Sub

Private Sub SetGannacciGridSwing2()
On Error GoTo ErrSection:
            
    With m.Annot
        If Val(.Prop("TB1")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 1, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 1, 0) = flexUnchecked
        End If
        
        If Val(.Prop("CD1")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 2, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 2, 0) = flexUnchecked
        End If
        
        If Val(.Prop("TB2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 3, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 3, 0) = flexUnchecked
        End If
        
        If Val(.Prop("CD2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 4, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 4, 0) = flexUnchecked
        End If
        
        If Val(.Prop("R1")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 5, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 5, 0) = flexUnchecked
        End If
        
        If Val(.Prop("R2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 6, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 6, 0) = flexUnchecked
        End If
        
        'diff properties
        If Val(.Prop("Diff_TB1_R2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 7, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 7, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Diff_CD1_R2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 8, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 8, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Diff_TB2_R1")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 9, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 9, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Diff_CD2_R1")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 10, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 10, 0) = flexUnchecked
        End If
        
        'equal properties
        If Val(.Prop("Equal_TB1_R2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 11, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 11, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Equal_CD1_R2")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 12, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 12, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Equal_TB2_R1")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 13, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 13, 0) = flexUnchecked
        End If
        
        If Val(.Prop("Equal_CD2_R1")) = 1 Then
            fgGannacciSwing.Cell(flexcpChecked, 14, 0) = flexChecked
        Else
            fgGannacciSwing.Cell(flexcpChecked, 14, 0) = flexUnchecked
        End If
    
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnotExt.SetGannacciGridSwing2"

End Sub


