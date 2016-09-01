VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSymbolGroup 
   Caption         =   "Symbol Group Editor"
   ClientHeight    =   5265
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
      Height          =   375
      Left            =   180
      TabIndex        =   14
      Top             =   1980
      Width           =   2535
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
      Caption         =   "frmSymbolGroup.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmSymbolGroup.frx":0050
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGroup.frx":0070
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraTotals 
      Height          =   1215
      Left            =   4380
      TabIndex        =   13
      Top             =   3480
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
      Caption         =   "frmSymbolGroup.frx":008C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolGroup.frx":00B8
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGroup.frx":00D8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtCurrentVolume 
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   840
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmSymbolGroup.frx":00F4
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
         Tip             =   "frmSymbolGroup.frx":0114
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":0134
      End
      Begin HexUniControls.ctlUniTextBoxXP txtTotalVolume 
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   120
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmSymbolGroup.frx":0150
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
         Tip             =   "frmSymbolGroup.frx":0170
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":0190
      End
      Begin HexUniControls.ctlUniTextBoxXP txtVolDivisor 
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   480
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmSymbolGroup.frx":01AC
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
         Tip             =   "frmSymbolGroup.frx":01CC
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":01EC
      End
      Begin HexUniControls.ctlUniTextBoxXP txtCurrentValue 
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Top             =   840
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmSymbolGroup.frx":0208
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
         Tip             =   "frmSymbolGroup.frx":0228
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":0248
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDivisor 
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   480
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmSymbolGroup.frx":0264
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
         Tip             =   "frmSymbolGroup.frx":0284
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":02A4
      End
      Begin HexUniControls.ctlUniTextBoxXP txtTotalValue 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   120
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmSymbolGroup.frx":02C0
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
         Tip             =   "frmSymbolGroup.frx":02E0
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":0300
      End
      Begin HexUniControls.ctlUniLabelXP lblValue 
         Height          =   255
         Left            =   720
         Top             =   870
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
         Caption         =   "frmSymbolGroup.frx":031C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolGroup.frx":035A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":037A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDivisor 
         Height          =   255
         Left            =   720
         Top             =   510
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
         Caption         =   "frmSymbolGroup.frx":0396
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolGroup.frx":03C6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":03E6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTotal 
         Height          =   255
         Left            =   720
         Top             =   150
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
         Caption         =   "frmSymbolGroup.frx":0402
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolGroup.frx":043C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":045C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraGridView 
      Height          =   315
      Left            =   4380
      TabIndex        =   10
      Top             =   4800
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
      Caption         =   "frmSymbolGroup.frx":0478
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolGroup.frx":04A4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGroup.frx":04C4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optViewVolumes 
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   0
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
         Caption         =   "frmSymbolGroup.frx":04E0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolGroup.frx":051E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":053E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optViewPrices 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   0
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
         Caption         =   "frmSymbolGroup.frx":055A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmSymbolGroup.frx":0596
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":05B6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkCustomIndex 
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   2520
      Width           =   3735
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmSymbolGroup.frx":05D2
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmSymbolGroup.frx":064C
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGroup.frx":066C
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraAdd 
      Height          =   1755
      Left            =   100
      TabIndex        =   8
      Top             =   100
      Width           =   4095
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmSymbolGroup.frx":0688
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolGroup.frx":06C4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGroup.frx":06E4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optLookup 
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   360
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
         Caption         =   "frmSymbolGroup.frx":0700
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolGroup.frx":074C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":076C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   750
         Left            =   3240
         TabIndex        =   6
         Top             =   780
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
         Caption         =   "frmSymbolGroup.frx":0788
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolGroup.frx":07B0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":07D0
         RightToLeft     =   0   'False
      End
      Begin MSComctlLib.ImageCombo cboFilters 
         Height          =   330
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "ImageCombo1"
      End
      Begin HexUniControls.ctlUniRadioXP optSymbolGroup 
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   1230
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
         Caption         =   "frmSymbolGroup.frx":07EC
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolGroup.frx":0820
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":0840
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   285
         Left            =   1410
         TabIndex        =   1
         Top             =   780
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmSymbolGroup.frx":085C
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
         Tip             =   "frmSymbolGroup.frx":087C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":089C
      End
      Begin HexUniControls.ctlUniRadioXP optAddSelected 
         Height          =   255
         Left            =   1500
         TabIndex        =   3
         Top             =   180
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
         Caption         =   "frmSymbolGroup.frx":08B8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolGroup.frx":0924
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":0944
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAddSymbol 
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   780
         Width           =   2895
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmSymbolGroup.frx":0960
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmSymbolGroup.frx":0998
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":09B8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
         Height          =   435
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   2475
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
         Caption         =   "frmSymbolGroup.frx":09D4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolGroup.frx":0A20
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGroup.frx":0A40
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgSymbols 
      Height          =   2955
      Left            =   4380
      TabIndex        =   7
      Top             =   180
      Width           =   3255
      _cx             =   5741
      _cy             =   5212
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
      BackColorAlternate=   14742776
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
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   2880
      Top             =   3780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   7
      DisplayContextMenu=   0   'False
      Tools           =   "frmSymbolGroup.frx":0A5C
      ToolBars        =   "frmSymbolGroup.frx":0CD0
   End
   Begin HexUniControls.ctlUniLabelXP lblNumSymbols 
      Height          =   255
      Left            =   5100
      Top             =   3240
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
      Caption         =   "frmSymbolGroup.frx":0E83
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmSymbolGroup.frx":0EBB
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGroup.frx":0EDB
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuAddSymbol 
         Caption         =   "&Add Symbol"
      End
      Begin VB.Menu mnuRemoveSymbol 
         Caption         =   "&Remove Selected Symbols"
      End
      Begin VB.Menu mnuSetActiveChart 
         Caption         =   "&Set Active Chart"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmSymbolGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSymbolGroup.frm
'' Description: Allows the user to edit the symbols that are in a symbol
''              group
''
'' Author:      Genesis Financial Data Services
''              425 Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date      Author      Description
'' 01/31/01  DAJ/TLB     Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type mPrivate
    lMaxSymbols As Long
    dPrevious As Double
    dPrevDiv As Double
    
    SymbolGroup As cSymbolGroup
    strName As String
    strDescription As String
    eGroupType As eSymbolGroupType
    aSymbolIds As cGdArray
    
    bModal As Boolean
    bOK As Boolean
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_SymbolID = 0
    eGDCol_Symbol
    eGDCol_Description
    eGDCol_Price
    eGDCol_PriceWeight
    eGDCol_PriceValue
    eGDCol_Volume
    eGDCol_VolumeWeight
    eGDCol_VolumeValue
    eGDCol_Flags
End Enum
Private Const kGridCols = 10

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

Public Property Get ID() As String
    ID = m.SymbolGroup.ID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCustomIndex_Click
'' Description: When the user turns the Custom Index option on or off, show/hide
''              controls as necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCustomIndex_Click()
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Bars to hold data for a symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCount As Long
    Dim lFromDate As Long
    
    EnableToolbar True
    
    If chkCustomIndex.Value = vbChecked Then
        Screen.MousePointer = vbHourglass
        lFromDate = LastDailyDownload
        With fgSymbols
            For lIndex = .FixedRows To .Rows - 1
                If ValOfText(.Cell(flexcpText, lIndex, GDCol(eGDCol_Price))) = 0 Then
                    If DM_GetBars(Bars, .Cell(flexcpText, lIndex, GDCol(eGDCol_Symbol)), 0, lFromDate) Then
                        .Cell(flexcpText, lIndex, GDCol(eGDCol_Price)) = CStr(RoundToSigDigits(Bars(eBARS_Close, Bars.Size - 1)))
                        .Cell(flexcpText, lIndex, GDCol(eGDCol_PriceWeight)) = "1"
                        .Cell(flexcpText, lIndex, GDCol(eGDCol_Volume)) = Format(Bars(eBARS_Vol, Bars.Size - 1), "#,##0")
                        .Cell(flexcpText, lIndex, GDCol(eGDCol_VolumeWeight)) = "1"
                    End If
                End If
                lCount = lCount + 1
                If lCount Mod 10 = 0 Then
                    'StatusMsg CStr(lCount) & " symbols calculated"
                End If
            Next lIndex
        End With
        CalcTotals
        'StatusMsg
        Screen.MousePointer = vbDefault
    End If
    
    ShowColumns

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.chkCustomIndex.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:   cmdAdd_Click
'' Descrition: When the user clicks on the Add button, add the appropriate
''             symbol(s) depending on which option button is selected
'' Inputs:     None
'' Returns:    None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    AddSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.cmdAdd.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: If the user clicks on the Lookup option, bring up the symbol
''              selector form to allow them to choose symbol(s) to add to the
''              grid, then add the one(s) they selected
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
    RaiseError "frmSymbolGroup.cmdLookup.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemove_Click
'' Description: If the user clicks on the remove button, remove the selected
''              symbols in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    RemoveSelectedSymbols

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.cmdRemove.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: If the user clicks on the Save button, set the saved flag and
''              hide the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save(ByVal strButton As String)
On Error GoTo ErrSection:

    Dim strNewName As String            ' Return from the AskBox
    Dim strText As String               ' Text to send into AskBox
    Dim bSaveAs As Boolean              ' Are we in Save As mode?
    Dim bReload As Boolean
    Dim bJustAdded As Boolean
    Dim lIndex As Long
    Dim lSymbolID As Long
    Dim bRename As Boolean              ' Are we renaming the symbol group?
    Dim strOldName As String            ' Old name
    Dim strSymbol$
    
    If chkCustomIndex.Value = vbChecked Then
        strText = vbCrLf & vbCrLf & "(Since you have chosen to make a Custom Index out" _
                                & " of this symbol group, the name must be less than" _
                                & " eight characters in length with no spaces)" & vbCrLf
    End If
    
    m.strName = Trim(m.strName)
    If Len(m.strName) = 0 Then
        strText = "Save the current Symbol Group as..." & strText
        strNewName = AskBox("h=Save ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    ElseIf strButton = "ID_SaveAs" Then
        strText = "Save a copy of the current Symbol Group as..." & strText
        strNewName = AskBox("h=Save As ; i=? ; g=string ; d=" & "Copy of " & m.strName & " ; " & strText)
        If Trim(UCase(strNewName)) <> UCase(m.strName) Then
            bSaveAs = True
        End If
    ElseIf strButton = "ID_Rename" Then
        strText = "Rename the current Symbol Group as..." & strText
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    Else
        strNewName = m.strName
    End If
    
    Do While chkCustomIndex.Value = vbChecked
        If InStr(strNewName, " ") > 0 Then
            InfBox "Custom Index symbol cannot contain spaces", "!", , "Error"
        ElseIf InStr(strNewName, "'") > 0 Then
            InfBox "Custom Index symbol cannot contain apostrophes", "!", , "Error"
        ElseIf Len(strNewName) > 8 Then
            InfBox "Custom Index symbol cannot exceed eight characters in length", "!", , "Error"
        Else
            Exit Do
        End If
        strText = vbCrLf & vbCrLf & "(Since you have chosen to make a Custom Index out" _
                                & " of this symbol group, the name must be less than" _
                                & " eight characters in length with no spaces)" & vbCrLf
        strText = "Rename the current Symbol Group as..." & strText
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & strNewName & " ; " & strText)
    Loop
    
    If Left(strNewName, 1) = "#" Then
        strNewName = Right(strNewName, Len(strNewName) - 1)
    End If
    If Len(Trim(strNewName)) = 0 Then
        Exit Sub 'Err.Raise vbObjectError + 1000, , "You must enter in a name for the filter"
    End If
    If Trim(UCase(strNewName)) <> UCase(m.strName) And strButton = "ID_Rename" Then
        strOldName = UCase(m.strName)
        bRename = True
    Else
        strOldName = ""
        bRename = False
    End If
    m.strName = Trim(strNewName)
    SetEditorCaption Me, "Symbol Group", m.strName
    
    ' Get from form
    If bSaveAs Then Set m.SymbolGroup = m.SymbolGroup.MakeCopy
    With m.SymbolGroup
        If bSaveAs Then .ID = ""
        .Name = Trim(m.strName)
        If Left(.Name, 1) = "#" Then .Name = Right(.Name, Len(.Name) - 1)
        .Desc = Trim(m.strDescription)
    
        ' Get symbols from the symbol grid
        .SymbolIDs.Clear
        .Symbols.Clear
    
        ' Custom Index stuff added 11/8/2001 by DAJ
        .IsIndex = (chkCustomIndex.Value = vbChecked)
        .PriceDivisor = ValOfText(txtDivisor.Text)
        .VolDivisor = ValOfText(txtVolDivisor.Text)
        .PriceWeights.Clear
        .VolWeights.Clear
        .Flags.Clear
    
        For lIndex = fgSymbols.FixedRows To fgSymbols.Rows - 1
            If CLng(fgSymbols.TextMatrix(lIndex, GDCol(eGDCol_SymbolID))) <> 0 Then
                .SymbolIDs.Add CLng(fgSymbols.TextMatrix(lIndex, GDCol(eGDCol_SymbolID)))
                
                ' Custom Index stuff added 11/8/2001 by DAJ
                .PriceWeights.Add ValOfText(fgSymbols.TextMatrix(lIndex, GDCol(eGDCol_PriceWeight)))
                .VolWeights.Add ValOfText(fgSymbols.TextMatrix(lIndex, GDCol(eGDCol_VolumeWeight)))
                .Flags.Add CLng(ValOfText(fgSymbols.TextMatrix(lIndex, GDCol(eGDCol_Flags))))
            Else
                strSymbol = fgSymbols.TextMatrix(lIndex, GDCol(eGDCol_Symbol))
                If Left(strSymbol, 1) = "*" Then
                    strText = fgSymbols.TextMatrix(lIndex, GDCol(eGDCol_Description))
                    If InStr(strText, "|") > 0 Then
                        strSymbol = strSymbol & "|" & strText
                    End If
                End If
                .Symbols.Add strSymbol
            End If
        Next lIndex
    
        ' Save the Custom Index information
        bReload = False
        .SymbolID = g.SymbolPool.SymbolIDforSymbol(UCase("#" & .Name))
        bJustAdded = (.SymbolID = 0)
        If .IsIndex = True Then
            ' If the user is renaming the symbol group, delete the old custom index first...
            If bRename = True And Len(strOldName) > 0 Then
                lSymbolID = g.SymbolPool.SymbolIDforSymbol(UCase("#" & strOldName))
                If SU_DeleteComposite(lSymbolID, UCase("#" & strOldName)) Then
                    g.SymbolPool.RemoveCustomIndex lSymbolID
                End If
            End If
        
            lSymbolID = .SymbolID
            If SU_SetCompositeInf(lSymbolID, UCase("#" & .Name), .Desc, .PriceDivisor, _
                                    .SymbolIDs, .PriceWeights, .Flags, _
                                    .VolDivisor, .VolWeights) = False Then
                InfBox "Problems saving #" & .Name, "!", , "Error"
            Else
                .SymbolID = lSymbolID
                If bJustAdded Then
                    If g.SymbolPool.AddCustomIndex(.SymbolID, UCase("#" & .Name)) = False Then
                        InfBox "Problems saving #" & .Name, "!", , "Error"
                    End If
                Else
                    g.SymbolPool.RecalcDirtyCriteria False, UCase("#" & .Name)
                End If
                UpdateVisibleCharts eRedo9_ReloadData, .SymbolID
            End If
        Else
            If g.SymbolPool.PoolRecForSymbolID(.SymbolID) <> -1 Then
                bReload = True
                If SU_DeleteComposite(.SymbolID, UCase("#" & .Name)) = False Then
                    InfBox "Problems deleting #" & .Name, "!", , "Error"
                Else
                    If g.SymbolPool.RemoveCustomIndex(.SymbolID) = False Then
                        InfBox "Problems deleting #" & .Name, "!", , "Error"
                    Else
                        .SymbolID = 0&
                    End If
                End If
            End If
        End If
        
        ' Save to file
        .ToFile
        
        ' Add back into pool
        .AddToPool True
        
#If 0 Then
        If .GroupType = eGROUP_QuoteList Then
            With frmQuotes
                ' wait until TotalRefresh is not active
                Do While .IsBusy
                    Sleep 0.1
                Loop
                .fgQuotes.Redraw = flexRDNone
                If .LoadGrid Then .TotalRefresh False 'True
                .fgQuotes.Redraw = flexRDBuffered
            End With
        End If
#End If
    End With
        
    ' Refresh symbol grid dropdown and list
    frmSymbolGrid.RefreshGrid
    
    ' Refresh the filter tab of the quote board if this is the group on it...
    frmQuotes.UpdateFilter "GRP:" & m.SymbolGroup.ID
    
    m.bOK = True
    EnableToolbar False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.Save", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_AfterEdit
'' Description: After the user is done editing a cell, recalculate the totals
'' Inputs:      Row and Column of the cell that was changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    With fgSymbols
        .TextMatrix(Row, Col) = .EditText
        If ValOfText(.Cell(flexcpText, Row, Col)) <> m.dPrevious Then
            CalcTotals
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.fgSymbols.AfterEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_AfterRowColChange
'' Description: After a row or column change in the grid, try to edit the cell
'' Inputs:      Old Row, Old Column, New Row, New Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    fgSymbols.EditCell

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.fgSymbols.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_BeforeEdit
'' Description: Only allow editing in the Price Weight and Volume Weight columns
'' Inputs:      Row, Col, and Whether or not to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> 4 And Col <> 7 Then
        Cancel = True
    Else
        m.dPrevious = ValOfText(fgSymbols.Cell(flexcpText, Row, Col))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.fgSymbols.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgSymbols_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim lCol As Long
    
    With fgSymbols
        ' Capture the current mouse row and column...
        lRow = .MouseRow
        lCol = .MouseCol
        
        If lRow >= .FixedRows And lRow < .Rows - 1 Then
            If Button = vbRightButton Then
                .RowSel = lRow
                If .SelectedRows <= 1 Then .Row = lRow
                mnuSetActiveChart.Caption = "&Set Active Chart to " & .TextMatrix(lRow, GDCol(eGDCol_Symbol))
                PopupMenu mnuPopUp
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.fgSymbols.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgSymbols_ChangeEdit()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.fgSymbols.ChangeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgSymbols_DblClick()
On Error GoTo ErrSection:

    Dim lRow As Long

    With fgSymbols
        lRow = .MouseRow
        If lRow >= .FixedRows And lRow < .Rows Then
            .Row = lRow
            SetActiveChartSymbol .TextMatrix(lRow, GDCol(eGDCol_Symbol))
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.fgSymbols.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_KeyDown
'' Description: If the user hits the delete key in the grid, remove all of the
''              selected items in the grid
'' Inputs:      KeyCode of the key that was hit, Shift status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    With fgSymbols
        If KeyCode = vbKeyDelete Then
            RemoveSelectedSymbols
        ElseIf KeyCode = vbKeyInsert Then
            LookupSymbol
        ElseIf KeyCode <> vbKeyShift And KeyCode <> vbKeyControl Then
            If .Row > .FixedRows - 1 And (.Col = 4 Or .Col = 7) Then
                .EditCell
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.fgSymbols.KeyDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgSymbols_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strCol$
    
    Select Case fgSymbols.MouseCol
        Case eGDCol_Symbol
            strCol = "Symbol"
        Case eGDCol_Description
            strCol = "Description"
        Case eGDCol_Price
            strCol = "Price"
        Case eGDCol_PriceWeight
            strCol = "Price Weight"
        Case eGDCol_PriceValue
            strCol = "Price Total"
        Case eGDCol_Volume
            strCol = "Volume"
        Case eGDCol_VolumeWeight
            strCol = "Volume Weight"
        Case eGDCol_VolumeValue
            strCol = "Volume Total"
    End Select

    GridTooltip fgSymbols, , strCol

End Sub

Private Sub fgSymbols_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    With fgSymbols
        If .Row > .FixedRows - 1 And (.Col = 4 Or .Col = 7) Then .EditCell
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.fgSymbols.MouseUp", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, calculate and show the number of
''              symbols
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    ShowNumSymbols
    If GetActiveWindow = Me.hWnd Then MoveFocus txtSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.Form.Activate", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        g.Help.ShowF1Help Me
    Else
        frmMain.DockPro_ShortcutKeyDown KeyCode, Shift, Me.Name
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSymbolGroup.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Me.Icon = Picture16(ToolbarIcon("ID_SymbolGroups"), , True)
    
    With tbToolbar
        .Tools("ID_Description").Picture = Picture16(ToolbarIcon("ID_News"))
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
    End With
    
    ' Set the grid font from the INI file...
    strFont = GetIniFileProperty("SymbolGroup", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgSymbols.Font, strFont
    
    ' Always make sure to hide the PopUp menu...
    mnuPopUp.Visible = False
    
    cboFilters.ImageList = frmMain.img16
    cboFilters.Locked = True
    LoadCombo True
    
    txtTotalValue.Enabled = False
    txtCurrentValue.Enabled = False
    txtTotalVolume.Enabled = False
    txtCurrentVolume.Enabled = False
    
    'tbToolbar.Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Unless we are unloading from code, just hide the form
'' Inputs:      Whether or not to cancel the unload, Unload Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        If AskToSave Then
            Cancel = True
        ElseIf m.bModal Then
            Cancel = True
            Me.Hide
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: When the user resizes the form, resize the controls on the form
''              appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next
    
    Dim lHeight As Long
    Dim lWidth As Long
    
    lHeight = chkCustomIndex.Top + chkCustomIndex.Height + fraAdd.Top + fraAdd.Height
    lWidth = fraAdd.Width * 2 + fraAdd.Left ' + fraTotals.Width + (fraAdd.Left * 3)
    If LimitFormSize(Me, lWidth, lHeight) Then Exit Sub
    
    lWidth = ScaleWidth - fraAdd.Width - (fraAdd.Left * 3)
    With fraGridView
        .Move fraAdd.Width + (fraAdd.Left * 2), ScaleHeight - .Height - fraAdd.Top, lWidth
    End With
    With fraTotals
        .Move fraAdd.Width + (fraAdd.Left * 2), fraGridView.Top - .Height, lWidth
    End With
    With lblNumSymbols
        If chkCustomIndex.Value = vbChecked Then
            .Move ScaleWidth - .Width - fraAdd.Left, fraTotals.Top - .Height
        Else
            .Move ScaleWidth - .Width - fraAdd.Left, ScaleHeight - .Height - fraAdd.Top
        End If
    End With
    
    txtTotalValue.Move fraTotals.Width - txtTotalValue.Width
    txtCurrentValue.Move fraTotals.Width - txtCurrentValue.Width
    txtTotalVolume.Move fraTotals.Width - txtTotalVolume.Width
    txtCurrentVolume.Move fraTotals.Width - txtCurrentVolume.Width
    txtDivisor.Move fraTotals.Width - txtDivisor.Width
    txtVolDivisor.Move fraTotals.Width - txtVolDivisor.Width
    lblTotal.Move fraTotals.Width - txtVolDivisor.Width - lblTotal.Width
    lblDivisor.Move fraTotals.Width - txtVolDivisor.Width - lblDivisor.Width
    lblValue.Move fraTotals.Width - txtVolDivisor.Width - lblValue.Width
    
    optViewPrices.Move (fraTotals.Width / 2) - ((optViewPrices.Width + optViewVolumes.Width) / 2)
    optViewVolumes.Move optViewPrices.Left + optViewPrices.Width
    
    With fgSymbols
        lHeight = lblNumSymbols.Top - (fraAdd.Top * 2)
        .Move fraTotals.Left, fraAdd.Top, lWidth, lHeight
    End With
    
    'With cmdRemove
    '    .Move fgSymbols.Left + ((fgSymbols.Width - .Width) / 2), fraAdd.Top
    'End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "SymbolGroup", FontToString(fgSymbols.Font), "Fonts", g.strIniFile

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSymbolGroup.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuAddSymbol_Click()
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.mnuAddSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuRemoveSymbol_Click()
On Error GoTo ErrSection:

    RemoveSelectedSymbols

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.mnuRemoveSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuSetActiveChart_Click()
On Error GoTo ErrSection:

    SetActiveChartSymbol fgSymbols.TextMatrix(fgSymbols.RowSel, GDCol(eGDCol_Symbol))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.mnuSetActiveChart.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAddSelected_Click
'' Description: If the user clicks on the Add Selected option, disable both
''              the symbol group combo box and the symbol text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAddSelected_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.optAddSelected.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAddSymbol_Click
'' Description: If the user clicks on the Add Symbol option, disable the symbol
''              group combo box and enable the symbol text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAddSymbol_Click()
On Error GoTo ErrSection:

    EnableControls
    MoveFocus txtSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.optAddSymbol.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLookup_Click
'' Description: If the user clicks on the Lookup option, bring up the symbol
''              selector form to allow them to choose symbol(s) to add to the
''              grid, then add the one(s) they selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLookup_Click()
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.optLookup.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optSymbolGroup_Click()
On Error GoTo ErrSection:

    EnableControls
    MoveFocus cboFilters

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.optSymbolGroup.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optViewPrices_Click
'' Description: If the user chooses to view prices, show the price weighting
''              columns on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optViewPrices_Click()
On Error GoTo ErrSection:

    ShowColumns

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.optViewPrices.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optViewVolumes_Click
'' Description: If the user chooses to view volumes, show the volume weighting
''              columns on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optViewVolumes_Click()
On Error GoTo ErrSection:

    ShowColumns

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.optViewVolumes.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim strID$

    ToggleFocus Me, Me.chkCustomIndex

    Select Case Tool.ID
        Case "ID_Save", "ID_SaveAs", "ID_Rename"
            Save Tool.ID
        
        Case "ID_Toolbox"
            If Not AskToSave Then
                strID = m.SymbolGroup.ID
                Unload Me
                frmToolbox.ShowMe eTab_SymbolGroups, strID
            End If
        
        Case "ID_Print"
            PrintMe
            
        Case "ID_Description"
            m.strDescription = frmNotes.ShowMe(m.strDescription, "Description")
            EnableToolbar True
        
        Case "ID_Close"
            If Not AskToSave Then
                If m.bModal Then
                    Me.Hide
                Else
                    Unload Me
                End If
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDivisor_Change
'' Description: If the user changes the divisor, recalc the current value
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDivisor_Change()
On Error GoTo ErrSection:

    EnableToolbar True
    If ValOfText(txtDivisor.Text) <> 0 Then
        txtCurrentValue.Text = CStr(ValOfText(txtTotalValue.Text) / ValOfText(txtDivisor.Text))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.txtDivisor.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDivisor_GotFocus
'' Description: When the text box gets the focus, save the current value
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDivisor_GotFocus()
On Error GoTo ErrSection:

    m.dPrevDiv = ValOfText(txtDivisor.Text)
    SelectAll txtDivisor

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.txtDivisor.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDivisor_LostFocus
'' Description: Do not let the user enter in a divisor of zero
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDivisor_LostFocus()
On Error GoTo ErrSection:

    If ValOfText(txtDivisor.Text) = 0 Then
        InfBox "Divisor cannot be zero", "!", , "Error"
        txtDivisor.Text = CStr(m.dPrevDiv)
        txtCurrentValue.Text = CStr(ValOfText(txtTotalValue.Text) / ValOfText(txtDivisor.Text))
        MoveFocus txtDivisor
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.txtDivisor.LostFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtSymbol_Change()
On Error GoTo ErrSection:
    
    EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.txtSymbol.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtSymbol_GotFocus()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.txtSymbol.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:   txtSymbol_KeyPress,  0
'' Descriptin: If the user presses Enter while in the symbol text box, simulate
''             a Add button click with what is currently in the stock text box
'' Inputs:     Key that was pressed
'' Returns:    None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdAdd_Click
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.txtSymbol.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddToGrid
'' Description: Add a symbol to the grid
'' Inputs:      Record number of the symbol to add
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddToGrid(ByVal nSymbolID As Long, Optional ByVal strSymbol As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long, lRec&, lInsertAt&, iPos&
    Dim bFound As Boolean
    Dim SymInfo As vbSymbolInfo
    Dim Bars As New cGdBars
    
    If m.eGroupType = eGROUP_QuoteList Then
        If fgSymbols.Rows - 1 >= m.lMaxSymbols Then
            InfBox "h=Error ; i=! ; There cannot be more than " & CStr(m.lMaxSymbols) & " symbols in the Quote List"
            Exit Sub
        End If
    End If
    
    If nSymbolID = 0 Then
        If strSymbol <> "" Then
            With fgSymbols
                bFound = False
                For lIndex = 1 To .Rows - 1
                    If .Cell(flexcpText, lIndex, 1) = strSymbol Then
                        bFound = True
                        Exit For
                    End If
                Next lIndex
                
                If bFound = False Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Cell(flexcpText, .Rows - 1, 0) = "0"
                    ' check if a hard-drive symbol
                    If Left(strSymbol, 1) = "*" Then
                        iPos = InStr(strSymbol, "|")
                    Else
                        iPos = 0
                    End If
                    If iPos > 0 Then
                        .Cell(flexcpText, .Rows - 1, 1) = Left(strSymbol, iPos - 1)
                        .Cell(flexcpText, .Rows - 1, 2) = Mid(strSymbol, iPos + 1)
                    Else
                        .Cell(flexcpText, .Rows - 1, 1) = strSymbol
                        .Cell(flexcpText, .Rows - 1, 2) = ""
                    End If
                    EnableToolbar True
                End If
            End With
        End If
    Else
        With fgSymbols
            If nSymbolID < 0 Then
                If chkCustomIndex.Value = vbChecked Then
                    InfBox "Cannot add a Custom Index to a Symbol Group that is a Custom Index", "!", , "Error"
                    Exit Sub
                ElseIf m.eGroupType = eGROUP_QuoteList Then
                    InfBox "Cannot add a Custom Index to the Quote Board", "!", , "Error"
                    Exit Sub
                Else
                    chkCustomIndex.Enabled = False
                End If
            End If
            
            If Left(g.SymbolPool.SymbolForID(nSymbolID), 2) = "$-" Then
                If m.eGroupType = eGROUP_QuoteList Then
                    InfBox "Cannot add an Industry Sector to the|Quote Board", "!", , "Error"
                    Exit Sub
                End If
            End If
            
            If m.aSymbolIds.BinarySearch(nSymbolID, lInsertAt) = False Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Cell(flexcpText, .Rows - 1, 0) = CStr(nSymbolID)
                lRec = g.SymbolPool.PoolRecForSymbolID(nSymbolID)
                If lRec >= 0 Then
                    .Cell(flexcpText, .Rows - 1, 1) = g.SymbolPool.Symbol(lRec)
                    .Cell(flexcpText, .Rows - 1, 2) = g.SymbolPool.Desc(lRec)
                    If chkCustomIndex = vbChecked Then
                        .Cell(flexcpText, .Rows - 1, 4) = "1"
                        .Cell(flexcpText, .Rows - 1, 7) = "1"
                        If DM_GetBars(Bars, g.SymbolPool.Symbol(lRec), , LastDailyDownload) Then
                            .Cell(flexcpText, .Rows - 1, 3) = CStr(RoundToSigDigits(Bars(eBARS_Close, Bars.Size - 1)))
                            .Cell(flexcpText, .Rows - 1, 6) = Format(Bars(eBARS_Vol, Bars.Size - 1), "#,##0")
                        End If
                    End If
                Else
                    'if not in pool, see if in rest of DBF
                    If SU_GetSymbolInf(nSymbolID, SymInfo) Then
                        .Cell(flexcpText, .Rows - 1, 1) = SymInfo.Symbol
                        .Cell(flexcpText, .Rows - 1, 2) = SymInfo.Description
                    End If
                End If
                m.aSymbolIds.Add nSymbolID, lInsertAt
                EnableToolbar True
            End If
        End With
    End If
    
    ShowNumSymbols
    
    'TLB: can't do this here -- way too slow when adding large groups
    'CalcTotals

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.AddToGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSelected
'' Description: Add the symbols that are selected in the symbol grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSelected()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nSymbolID&
            
    With frmSymbolGrid.fgVirtual
        If m.eGroupType = eGROUP_QuoteList Then
            If .SelectedRows + fgSymbols.Rows - 2 > m.lMaxSymbols Then
                InfBox "h=Error ; i=! ; There cannot be more than " & CStr(m.lMaxSymbols) & " symbols in the Quote List"
                Exit Sub
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        
        For lIndex = 0 To .SelectedRows - 1
            nSymbolID = g.SymbolPool.SymbolIDforSymbol(.Cell(flexcpText, .SelectedRow(lIndex), kSymbolCol))
            If nSymbolID <> 0 Then AddToGrid nSymbolID
        Next lIndex
    End With
    optAddSelected.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.AddSelected", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowNumSymbols
'' Description: Update the number of symbols label with the number of symbols
''              currently in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowNumSymbols()
On Error GoTo ErrSection:

    lblNumSymbols = "# Symbols:  " & CStr(fgSymbols.Rows - fgSymbols.FixedRows)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.ShowNumSymbols", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load up the filters combo box with the symbol groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo(Optional ByVal bShowFilters As Boolean = False)
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
    Dim bScans As Boolean               ' Are we doing scans?
   
    bScans = ScansEnabled
        
    If cboFilters.ComboItems.Count > 0 Then
        strSelID = cboFilters.SelectedItem.Key
        cboFilters.ComboItems.Clear
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
                Select Case UCase(strType)
                    Case "GRP"
                        If obj.GroupType = eGROUP_Normal Or obj.GroupType = eGROUP_QuoteList Then
                            strPicture = ToolbarIcon("ID_SymbolGroups")
                        End If
                    Case "FIL"
                        If bScans And bShowFilters Then
                            strPicture = ToolbarIcon("ID_Filters")
                        End If
                End Select
                If Len(strPicture) > 0 And obj.IsActive = True Then
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
        Next
    End With
    If iSortStart > 0 Then
        aItems.Sort eGdSort_IgnoreCase, iSortStart
    End If

    For lIndex = 0 To aItems.Size - 1
        strItem = aItems(lIndex)
        cboFilters.ComboItems.Add , Parse(strItem, vbTab, 2), _
            Parse(strItem, vbTab, 1), Parse(strItem, vbTab, 3)
    Next


    If bSelExists Then
        cboFilters.ComboItems(strSelID).Selected = True
    Else
        cboFilters.ComboItems(1).Selected = True
    End If

    cboFilters.Refresh

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddGroup
'' Description: Add the symbols in another symbol group to this symbol group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddGroup()
On Error GoTo ErrSection:

    Dim lFieldNum As Long               ' Field number for the symbol group
    Dim aSymbols As cGdArray            ' Array of true/false in grid values
    Dim lIndex As Long                  ' Array into a for loop
    
    ' Get the field number for the symbol group
    lFieldNum = g.SymbolPool.FieldNumForID(cboFilters.SelectedItem.Key)
    Set aSymbols = g.SymbolPool.ArrayTable.FieldArray(lFieldNum)
    
    If m.eGroupType = eGROUP_QuoteList Then
        If aSymbols.CountOf(1) + fgSymbols.Rows - 1 > m.lMaxSymbols Then
            InfBox "h=Error ; i=! ; There cannot be more than " & CStr(m.lMaxSymbols) & " symbols in the Quote List"
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Add the symbols from that group into this one
    For lIndex = 0 To aSymbols.Size - 1
        If Abs(aSymbols(lIndex)) = 1 Then
            AddToGrid g.SymbolPool.SymbolID(lIndex)
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.AddGroup", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowColumns
'' Description: Show the appropriate columns in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowColumns()
On Error GoTo ErrSection:

    If chkCustomIndex.Value = vbChecked Then
        fraGridView.Visible = True
        fraTotals.Visible = True
               
        If optViewPrices.Value = True Then
            lblTotal.Caption = "Total Value:"
            lblValue.Caption = "Current Value:"
            
            txtTotalValue.Visible = True
            txtTotalVolume.Visible = False
            txtDivisor.Visible = True
            txtVolDivisor.Visible = False
            txtCurrentValue.Visible = True
            txtCurrentVolume.Visible = False
        Else
            lblTotal.Caption = "Total Volume:"
            lblValue.Caption = "Current Volume:"
            
            txtTotalValue.Visible = False
            txtTotalVolume.Visible = True
            txtDivisor.Visible = False
            txtVolDivisor.Visible = True
            txtCurrentValue.Visible = False
            txtCurrentVolume.Visible = True
        End If
    Else
        fraGridView.Visible = False
        fraTotals.Visible = False
    End If

    With fgSymbols
        ' Description
        .ColHidden(2) = (chkCustomIndex.Value = vbChecked)
        
        ' Price Columns
        .ColHidden(3) = (chkCustomIndex.Value = vbUnchecked) Or Not optViewPrices
        .ColHidden(4) = (chkCustomIndex.Value = vbUnchecked) Or Not optViewPrices
        .ColHidden(5) = (chkCustomIndex.Value = vbUnchecked) Or Not optViewPrices
        
        ' Volume Columns
        .ColHidden(6) = (chkCustomIndex.Value = vbUnchecked) Or Not optViewVolumes
        .ColHidden(7) = (chkCustomIndex.Value = vbUnchecked) Or Not optViewVolumes
        .ColHidden(8) = (chkCustomIndex.Value = vbUnchecked) Or Not optViewVolumes
    End With
    
    Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.ShowColumns", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcTotals
'' Description: Calculate the total price/volume and current price/volume of
''              the custom index
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalcTotals()
On Error GoTo ErrSection:

    Dim dTotalVolume As Double          ' Total Volume for the custom index
    Dim dTotalPrice As Double           ' Total Price for the custom index
    Dim lIndex As Long                  ' Index for a for loop
    
    dTotalPrice = 0#
    dTotalVolume = 0#
    
    With fgSymbols
        For lIndex = .FixedRows To .Rows - .FixedRows
            .Cell(flexcpText, lIndex, 5) = CStr(RoundToSigDigits(ValOfText(.Cell(flexcpText, lIndex, 3)) * ValOfText(.Cell(flexcpText, lIndex, 4))))
            dTotalPrice = dTotalPrice + ValOfText(.Cell(flexcpText, lIndex, 5))
            
            .Cell(flexcpText, lIndex, 8) = Format(ValOfText(.Cell(flexcpText, lIndex, 6)) * ValOfText(.Cell(flexcpText, lIndex, 7)), "#,##0")
            dTotalVolume = dTotalVolume + ValOfText(.Cell(flexcpText, lIndex, 8))
        Next lIndex
    End With
    
    txtTotalValue.Text = CStr(dTotalPrice)
    If ValOfText(txtDivisor.Text) <> 0 Then
        txtCurrentValue.Text = dTotalPrice / ValOfText(txtDivisor.Text)
    End If
    
    txtTotalVolume.Text = Format(dTotalVolume, "#,##0")
    If ValOfText(txtVolDivisor.Text) <> 0 Then
        txtCurrentVolume.Text = Format(dTotalVolume / ValOfText(txtVolDivisor.Text), "#,##0")
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.CalcTotals", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtVolDivisor_Change
'' Description: If the user changes the volume divisor, recalc the totals
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtVolDivisor_Change()
On Error GoTo ErrSection:

    EnableToolbar True
    If ValOfText(txtVolDivisor.Text) <> 0 Then
        txtCurrentVolume.Text = Format(ValOfText(txtTotalVolume.Text) / ValOfText(txtVolDivisor.Text), "#,##0")
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.txtVolDivisor.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtVolDivisor_GotFocus
'' Description: When the text box gets the focus, save the current value
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtVolDivisor_GotFocus()
On Error GoTo ErrSection:

    m.dPrevDiv = ValOfText(txtVolDivisor.Text)
    SelectAll txtVolDivisor

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.txtVolDivisor.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtVolDivisor_LostFocus
'' Description: Do not let the user enter in a divisor of zero
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtVolDivisor_LostFocus()
On Error GoTo ErrSection:

    If ValOfText(txtVolDivisor.Text) = 0 Then
        InfBox "Volume divisor cannot be zero", "!", , "Error"
        txtVolDivisor.Text = CStr(m.dPrevDiv)
        txtCurrentVolume.Text = Format(ValOfText(txtTotalVolume.Text) / ValOfText(txtVolDivisor.Text), "#,##0")
        MoveFocus txtVolDivisor
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.txtVolDivisor.LostFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub EnableControls()
On Error GoTo ErrSection:

    Enable cboFilters, optSymbolGroup
    Enable txtSymbol, optAddSymbol
    If optAddSymbol And Len(Trim(txtSymbol)) = 0 Then
        Disable cmdAdd
    Else
        Enable cmdAdd
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.EnableControls", eGDRaiseError_Raise

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With fgSymbols
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Cols = kGridCols
        .Rows = 1
        
        .FixedRows = 1
        .FixedCols = 0
        
        .Cell(flexcpText, 0, GDCol(eGDCol_SymbolID)) = "SymbolID"
        .Cell(flexcpText, 0, GDCol(eGDCol_Symbol)) = "Symbol"
        .Cell(flexcpText, 0, GDCol(eGDCol_Description)) = "Description"
        
        ' Custom Index stuff added 11/8/2001 by DAJ
        .Cell(flexcpText, 0, GDCol(eGDCol_Price)) = "Price"
        .Cell(flexcpText, 0, GDCol(eGDCol_PriceWeight)) = "Weight"
        .Cell(flexcpText, 0, GDCol(eGDCol_PriceValue)) = "Total"
        .Cell(flexcpText, 0, GDCol(eGDCol_Volume)) = "Volume"
        .Cell(flexcpText, 0, GDCol(eGDCol_VolumeWeight)) = "Weight"
        .Cell(flexcpText, 0, GDCol(eGDCol_VolumeValue)) = "Total"
        .Cell(flexcpText, 0, GDCol(eGDCol_Flags)) = "Flags"
        
        .ColHidden(GDCol(eGDCol_SymbolID)) = True
        
        ' Custom Index stuff added 11/8/2001 by DAJ
        If m.SymbolGroup.IsIndex = True Then
            .ColHidden(GDCol(eGDCol_Description)) = True
        Else
            .ColHidden(GDCol(eGDCol_Price)) = True
            .ColHidden(GDCol(eGDCol_PriceWeight)) = True
            .ColHidden(GDCol(eGDCol_PriceValue)) = True
        End If
        .ColHidden(GDCol(eGDCol_Volume)) = True
        .ColHidden(GDCol(eGDCol_VolumeWeight)) = True
        .ColHidden(GDCol(eGDCol_VolumeValue)) = True
        .ColHidden(GDCol(eGDCol_Flags)) = True
        
        .ExtendLastCol = True
        .ExplorerBar = flexExSortShow
        .ScrollTrack = True
        .SheetBorder = RGB(128, 128, 128)
        .SelectionMode = flexSelectionListBox
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeColumns
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.InitGrid", eGDRaiseError_Raise

End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRecNum As Long                 ' Record Number in the Symbol Pool
    Dim Bars As New cGdBars             ' Temporary Bars structure
    Dim bHasBars As Boolean
    Dim SymInfo As vbSymbolInfo
    Dim iPos&, strSymbol$

    With fgSymbols
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 0 To m.SymbolGroup.SymbolIDs.Size - 1
            If m.SymbolGroup.SymbolIDs(lIndex) < 0 And m.SymbolGroup.IsIndex Then
                InfBox "Cannot add a Custom Index to a Symbol Group that is a Custom Index", "!", , "Error"
            ElseIf m.SymbolGroup.SymbolIDs(lIndex) < 0 And m.eGroupType = eGROUP_QuoteList Then
                InfBox "Cannot add a Custom Index to the Quote Board", "!", , "Error"
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, GDCol(eGDCol_SymbolID)) = CStr(m.SymbolGroup.SymbolIDs(lIndex))
                lRecNum = g.SymbolPool.PoolRecForSymbolID(m.SymbolGroup.SymbolIDs(lIndex))
                If lRecNum >= 0 Then
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = g.SymbolPool.Symbol(lRecNum)
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Description)) = g.SymbolPool.Desc(lRecNum)
                    If m.eGroupType <> eGROUP_QuoteList And m.SymbolGroup.IsIndex Then
                        bHasBars = DM_GetBars(Bars, g.SymbolPool.Symbol(lRecNum))
                    End If
                Else
                    'if not in pool, see if in DBF
                    If SU_GetSymbolInf(m.SymbolGroup.SymbolIDs(lIndex), SymInfo) Then
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = SymInfo.Symbol
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Description)) = SymInfo.Description
                    End If
                End If
                
                ' Custom Index stuff added 11/8/2001 by DAJ
                If m.SymbolGroup.IsIndex = True Then
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Price)) = CStr(Bars(eBARS_Close, Bars.Size - 1))
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_PriceWeight)) = CStr(m.SymbolGroup.PriceWeights(lIndex))
                    
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Volume)) = Format(Bars(eBARS_Vol, Bars.Size - 1), "#,##0")
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_VolumeWeight)) = CStr(m.SymbolGroup.VolWeights(lIndex))
                    
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Flags)) = CStr(m.SymbolGroup.Flags(lIndex))
                End If
                If m.SymbolGroup.SymbolIDs(lIndex) < 0 Then chkCustomIndex.Enabled = False
            End If
        Next lIndex
        
        For lIndex = 0 To m.SymbolGroup.Symbols.Size - 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GDCol(eGDCol_SymbolID)) = "0"
            strSymbol = m.SymbolGroup.Symbols(lIndex)
            ' check if a hard-drive symbol
            If Left(strSymbol, 1) = "*" Then
                iPos = InStr(strSymbol, "|")
            Else
                iPos = 0
            End If
            If iPos > 0 Then
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = Left(strSymbol, iPos - 1)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Description)) = Mid(strSymbol, iPos + 1)
            Else
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = strSymbol
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Description)) = ""
            End If
        Next lIndex
        
        If .Rows > .FixedRows Then
            .Cell(flexcpSort, .FixedRows, 1, .Rows - .FixedRows, 1) = flexSortStringAscending
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.LoadGrid", eGDRaiseError_Raise
    
End Sub

Public Function ShowMe(ByVal strPath As String, ByVal strID As String, _
                    Optional ByVal bLoadSelected As Boolean = False, _
                    Optional alToAdd As cGdArray = Nothing, _
                    Optional ByVal bKeepExisting As Boolean = True, _
                    Optional ByVal bModal As Boolean = False, _
                    Optional ByVal bSaveNow As Boolean = False, _
                    Optional ByVal strScreenerName As String = "") As String
On Error GoTo ErrSection:

    Dim lIndex As Long
    
    Set m.SymbolGroup = New cSymbolGroup
    m.bModal = bModal

    ' Load the Symbol Group from the file...
    If Len(strID) > 0 Then
        If Not m.SymbolGroup.FromFile(strPath, strID, True) Then
            If Len(strScreenerName) = 0 Then
                Err.Raise vbObjectError + 1000, , strID & " could not be loaded"
            End If
        End If
    End If
    
    ' Get the maximum number of Quote Board symbols this user is allowed...
    m.lMaxSymbols = MaxSymbolsAllowed

    If Not alToAdd Is Nothing Then
        With m.SymbolGroup
            If bKeepExisting = True Then
                If .GroupType = eGROUP_QuoteList Then
                    If alToAdd.Size + .SymbolIDs.Size + .Symbols.Size > m.lMaxSymbols Then
                        Err.Raise vbObjectError + 1000, , "There cannot be more than " & m.lMaxSymbols & " symbols on the quote board"
                    End If
                End If
                
                For lIndex = 0 To alToAdd.Size - 1
                    .AddSymbolID alToAdd(lIndex)
                Next lIndex
            Else
                If .GroupType = eGROUP_QuoteList Then
                    If alToAdd.Size > m.lMaxSymbols Then
                        Err.Raise vbObjectError + 1000, , "There cannot be more than " & m.lMaxSymbols & " symbols on the quote board"
                    End If
                End If
                
                .SymbolIDs.Clear
                .PriceWeights.Clear
                .VolWeights.Clear
                .Flags.Clear
                
                For lIndex = 0 To alToAdd.Size - 1
                    .AddSymbolID alToAdd(lIndex)
                Next lIndex
            End If
        End With
    End If

    With m.SymbolGroup
        Screen.MousePointer = vbHourglass
        m.strName = .Name
        If Len(m.strName) = 0 Then
            m.strName = strScreenerName
        End If
        m.strDescription = .Desc
        m.eGroupType = .GroupType
        
        ' If there is nothing selected in the symbol grid, or it is not even visible,
        ' then disable the "Add Selected Symbols" option
        If frmSymbolGrid.Visible = False Or frmSymbolGrid.fgVirtual.SelectedRows = 0 Then
            optAddSelected.Enabled = False
        End If
        
        ' Default to the "Add Symbol" option
        optAddSymbol.Value = True
        cboFilters.Enabled = False
        
        fgSymbols.Redraw = flexRDNone
        InitGrid
        LoadGrid
        fgSymbols.Redraw = flexRDBuffered
        
        txtDivisor.Text = CStr(.PriceDivisor)
        txtVolDivisor.Text = CStr(.VolDivisor)
        
        If .IsIndex Then
            CalcTotals
            chkCustomIndex.Value = vbChecked
            fraGridView.Visible = True
            fraTotals.Visible = True
        Else
            chkCustomIndex.Value = vbUnchecked
            fraGridView.Visible = False
            fraTotals.Visible = False
        End If
        
        Set m.aSymbolIds = .SymbolIDs
        m.aSymbolIds.Sort
        
        ' Load the selected items from the symbol grid if we need to
        If bLoadSelected = True Then AddSelected
        
        If m.eGroupType = eGROUP_QuoteList Then
            chkCustomIndex.Visible = False
        End If
        
        Screen.MousePointer = vbDefault
    End With
    
    If bSaveNow Then
        tbToolbar_ToolClick tbToolbar.Tools("ID_Save")
    Else
        SetEditorCaption Me, "Symbol Group", Trim(m.strName)
        
        If (Trim(m.strName) = "" And m.aSymbolIds.Size > 0) Or (Not alToAdd Is Nothing) Or (bLoadSelected = True) Then
            EnableToolbar True '(symbols added from symbol grid to new group)
        Else
            EnableToolbar False
        End If
    End If
    
    m.bOK = False
    
    ShowForm Me, bModal, frmMain
    If bModal Then
        If m.bOK Then ShowMe = m.SymbolGroup.ID
        Unload Me
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmSymbolGroup.ShowMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableToolbar
'' Description: Enable/Disable the controls on the toolbar appropriately
'' Inputs:      Whether to Enable or Disable
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableToolbar(ByVal bEnableSave As Boolean)
On Error GoTo ErrSection:

    With tbToolbar
        .Tools("ID_Toolbox").Enabled = Not m.bModal
        .Tools("ID_Save").Enabled = bEnableSave
        .Tools("ID_SaveAs").Enabled = (Trim(m.strName) <> "")
        If m.eGroupType = eGROUP_QuoteList Then
            .Tools("ID_Rename").Enabled = False
        Else
            .Tools("ID_Rename").Enabled = (Trim(m.strName) <> "")
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".EnableToolbar", eGDRaiseError_Raise
    
End Sub

' Returns True if Cancelled
Public Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim strResponse As String
    
    If tbToolbar.Tools("ID_Save").Enabled Then
        If WindowState = vbMinimized Then WindowState = vbNormal
    
        strResponse = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", Caption)
        Select Case strResponse
            Case "C"
                AskToSave = True
            Case "Y"
                Save "ID_Save"
        End Select
    End If
        
ErrExit:
    Exit Function

ErrSection:
    AskToSave = True
    RaiseError Me.Name & ".AskToSave"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "CNV SymbolGroup", Me, 0
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.PrintMe", eGDRaiseError_Raise
            
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the Print Preview
'' Inputs:      Variant set of arguments from the Print Preview control
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim X As Integer                    ' Index into a for loop
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim strText As String               ' Text to add to the printer control
    Dim abColHidden() As Boolean        ' Array of column hidden values
    
    With frmPrintPreview.vp
        .StartDoc
        
        ' Header and Footer
        DoPrintHeader
        
        ' Report Heading and date/time...
        .Font.Name = "Times New Roman"
        .Font.Bold = True
        .Font.Size = 14
        .FontUnderline = True
        .Text = vbLf & "Symbol Group:"
        .FontUnderline = False
        .Text = "    " & Trim(m.strName) & vbLf
        .Font.Size = 12
        .Font.Bold = False
        .Text = "Description: " & Trim(m.strDescription) & vbCrLf & vbCrLf
        
        ' If printing or previewing show the grid, otherwise if printing to file
        ' walk through the grid outputing tab delimeted strings
        If frmPrintPreview.GoingToFile = False Then
            If chkCustomIndex = vbChecked Then
                ReDim abColHidden(fgSymbols.Cols) As Boolean
                
                For lCol = 0 To fgSymbols.Cols - 1
                    abColHidden(lCol) = fgSymbols.ColHidden(lCol)
                Next lCol
                
                fgSymbols.Redraw = flexRDNone
                fgSymbols.ExtendLastCol = False
                For lCol = 0 To fgSymbols.Cols - 1
                    fgSymbols.ColHidden(lCol) = False
                Next lCol
                fgSymbols.ColHidden(GDCol(eGDCol_Flags)) = True
                fgSymbols.ColHidden(GDCol(eGDCol_SymbolID)) = True
                                
                'fgSymbols.TextMatrix(0, GDCol(eGDCol_PriceWeight)) = "Price Weight"
                'fgSymbols.TextMatrix(0, GDCol(eGDCol_VolumeWeight)) = "Volume Weight"
                fgSymbols.AutoSize 0, fgSymbols.Cols - 1, False, 75
            End If
            .RenderControl = fgSymbols.hWnd
            If chkCustomIndex = vbChecked Then
                'fgSymbols.TextMatrix(0, GDCol(eGDCol_PriceWeight)) = "Weight"
                'fgSymbols.TextMatrix(0, GDCol(eGDCol_VolumeWeight)) = "Weight"
                For lCol = 0 To fgSymbols.Cols - 1
                    fgSymbols.ColHidden(lCol) = abColHidden(lCol)
                Next lCol
                fgSymbols.ExtendLastCol = True
                fgSymbols.AutoSize 0, fgSymbols.Cols - 1, False, 75
                fgSymbols.Redraw = flexRDBuffered
            End If
        Else
            With fgSymbols
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If lCol <> GDCol(eGDCol_Flags) And lCol <> GDCol(eGDCol_SymbolID) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        End If
        
        If chkCustomIndex = vbChecked Then
            .Text = vbCrLf
            .Text = "Total Price:" & vbTab & vbTab & vbTab & txtTotalValue.Text & vbLf
            .Text = "Overall Price Divisor:" & vbTab & vbTab & txtDivisor.Text & vbLf
            .Text = "Current Price of Index:" & vbTab & txtCurrentValue.Text & vbCrLf
            
            .Text = "Total Volume:" & vbTab & vbTab & vbTab & txtTotalVolume.Text & vbLf
            .Text = "Overall Volume Divisor:" & vbTab & txtVolDivisor.Text & vbLf
            .Text = "Current Volume of Index:" & vbTab & txtCurrentVolume.Text & vbLf
        End If
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.GenerateReport", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Change the font of the quotes grid if the user chooses to
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgSymbols, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub RemoveSelectedSymbols()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim alSelectedRows As New cGdArray  ' Array of selected rows in the grid
    Dim nSymbolID As Long, nItem As Long
    Dim bHasCustom As Boolean
    
    EnableToolbar True

    alSelectedRows.Create eGDARRAY_Longs
    
    With fgSymbols
        ' Save the selected rows from the grid
        For lIndex = 0 To .SelectedRows - 1
            alSelectedRows.Add .SelectedRow(lIndex)
        Next lIndex
        
        ' Remove all of the selected rows from the grid
        For lIndex = alSelectedRows.Size - 1 To 0 Step -1
            nSymbolID = ValOfText(.Cell(flexcpText, alSelectedRows(lIndex), 0))
            .RemoveItem alSelectedRows(lIndex)
            ' take it out of the local symbols array
            If m.aSymbolIds.BinarySearch(nSymbolID, nItem) Then
                m.aSymbolIds.Remove nItem
            End If
        Next lIndex
        
        ' Check to see if there are any Custom Indexes left
        For lIndex = .FixedRows To .Rows - .FixedRows
            If ValOfText(.Cell(flexcpText, lIndex, 0)) < 0 Then
                bHasCustom = True
                Exit For
            End If
        Next lIndex
        chkCustomIndex.Enabled = Not bHasCustom
    End With
    
    ' TLB 8/28/2009 for #5281
    If chkCustomIndex Then
        CalcTotals
    End If
    
    ShowNumSymbols

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.RemoveSelectedSymbols", eGDRaiseError_Raise
    
End Sub

Private Sub LookupSymbol()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbols that were selected to add
    
    ' Get symbols from symbol selector form
    Set astrSymbols = frmSymbolSelector.ShowMe("", True, True)
    'Set astrSymbols = frmSymbolSelector.ShowMe("", True, True, , True)
    
    txtSymbol.Text = astrSymbols.JoinFields(",")
    optAddSymbol.Value = True
    If Len(Trim(txtSymbol)) > 0 Then AddSymbol
    EnableControls

ErrExit:
    Set astrSymbols = Nothing
    Exit Sub
    
ErrSection:
    Set astrSymbols = Nothing
    RaiseError "frmSymbolGroup.LookupSymbol", eGDRaiseError_Raise
    
End Sub

Private Sub AddSymbol()
On Error GoTo ErrSection:

    Dim lRecNum As Long                 ' Record number for the symbol
    Dim strReturn As String             ' Return from an ask box
    Dim strSymbol As String, strTemp As String
    Dim aSymbols As New cGdArray, lSymbol&
    Dim aUnknowns As New cGdArray, i&
    Dim nSymbolID&, bUnknown As Boolean
    Dim Bars As New cGdBars
    
    ' Add the symbol that the user typed in
    If optAddSymbol.Value = True Then
        strSymbol = UCase(Trim(txtSymbol.Text))
        
        'see if a symbol or filename
        If Len(strSymbol) = 0 Then
            Err.Raise vbObjectError + 1000, , "You must enter a symbol first"
        ElseIf InStr(strSymbol, "\") > 0 And InStr(strSymbol, "|") = 0 Then
            'get symbols from file
            If FileExist(strSymbol) Then
                aSymbols.FromFile strSymbol
            Else
                Err.Raise vbObjectError + 1000, , "File does not exist:|" & strSymbol
            End If
            Screen.MousePointer = vbHourglass
        Else
            'comma or tab delimited string of symbols
            aSymbols.SplitFields strSymbol, "," & vbTab & Chr(10)
        End If
        
        For lSymbol = 0 To aSymbols.Size - 1
            strSymbol = ""
            strTemp = Trim(aSymbols(lSymbol))
            If InStr(strTemp, "=") = 0 Then
                strSymbol = Parse(strTemp, vbTab, 1)
                If IsDigit(strSymbol, 1) And Not IsAlpha(strSymbol) Then
                    strSymbol = Parse(strTemp, vbTab, 2)
                End If
            End If
            If Len(strSymbol) > 0 Then
                bUnknown = False
                If Left(strSymbol, 1) = "*" Then
                    ' check for a hard-drive symbol
                    If InStr(strSymbol, "|") > 0 Then
                        AddToGrid 0, strSymbol
                    End If
                ElseIf InStr(strSymbol, " ") = 0 Then
                    ' Not an option symbol:
                    ' set bars prop so will use gdSymbol class to convert to
                    ' Genesis symbology (stock classes, contract century, etc.)
                    If Right(strSymbol, 1) = "-" Then
                        strSymbol = strSymbol & "067"
                    End If
                    Bars.Prop(eBARS_Symbol) = strSymbol
                    strSymbol = Bars.Prop(eBARS_Symbol)
                    nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
                    If nSymbolID = 0 Then
                        'if not in pool, check rest of DBF
                        nSymbolID = SU_GetSymID(strSymbol)
                    End If
                    If nSymbolID <> 0 Then
                        AddToGrid nSymbolID
                    Else
                        bUnknown = True
                    End If
                Else
                    ' Option symbol?
                    If m.eGroupType <> eGROUP_QuoteList Then
                        bUnknown = True
                    ElseIf Len(Parse(strSymbol, " ", 1)) < 4 And Len(Parse(strSymbol, " ", 2)) = 2 Then
                        ' stock option
                        AddToGrid 0, strSymbol
                    '****ElseIf 0 Then
                    ElseIf InStr(strSymbol, "-") > 0 And DoFutOpts Then
                        ' future option
                        AddToGrid 0, strSymbol
                    Else
                        bUnknown = True
                    End If
                End If
                ' add unknown symbol to list of unknowns
                If bUnknown Then
                    If Not aUnknowns.BinarySearch(strSymbol) Then
                        aUnknowns.Add strSymbol
                        aUnknowns.Sort
                    End If
                End If
            End If
        Next
        txtSymbol = ""
        
    ' Add the selected symbols in the symbol grid
    ElseIf optAddSelected.Value = True Then
        AddSelected
    ElseIf optSymbolGroup.Value = True Then
        AddGroup
    End If
    
    If chkCustomIndex Then
        CalcTotals
    End If

    With fgSymbols
        .Redraw = flexRDNone
        If .Rows > .FixedRows Then
            strSymbol = ""
            If .Row >= .FixedRows Then
                strSymbol = .TextMatrix(.Row, GDCol(eGDCol_Symbol))
            End If
            .Cell(flexcpSort, .FixedRows, GDCol(eGDCol_Symbol), .Rows - 1, GDCol(eGDCol_Symbol)) = flexSortStringAscending
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, GDCol(eGDCol_Symbol)) = strSymbol Then
                    .Row = i
                    .ShowCell .Row, GDCol(eGDCol_Symbol)
                    Exit For
                End If
            Next
        End If
        .Redraw = flexRDBuffered
    End With

    Screen.MousePointer = vbDefault
    
    If aUnknowns.Size > 0 Then
        strSymbol = Trim(aUnknowns.JoinFields(" "))
        If m.eGroupType = eGROUP_QuoteList And aUnknowns.Size = 1 Then
            InfBox "h=Error ; i=[] ; Symbol does not exist:  " & strSymbol _
                & "||(Note: Stock option symbols need to| have a space, e.g. " _
                & Chr(34) & "IBM BP" & Chr(34) & ")"
        Else
            InfBox "h=Error ; i=[] ; Symbol does not exist: " & strSymbol
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGroup.AddSymbol", eGDRaiseError_Raise
    
End Sub


