VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOptChartEdit3d 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chart Settings"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraGridLines 
      Height          =   600
      Left            =   180
      TabIndex        =   3
      Top             =   3990
      Width           =   4440
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptChartEdit3d.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptChartEdit3d.frx":0034
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit3d.frx":0054
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optGridLines 
         Height          =   220
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   285
         Width           =   765
         _ExtentX        =   1349
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
         Caption         =   "frmOptChartEdit3d.frx":0070
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":0098
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":00B8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optGridLines 
         Height          =   220
         Index           =   1
         Left            =   1260
         TabIndex        =   6
         Top             =   285
         Width           =   765
         _ExtentX        =   1349
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
         Caption         =   "frmOptChartEdit3d.frx":00D4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":0100
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":0120
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optGridLines 
         Height          =   220
         Index           =   2
         Left            =   2385
         TabIndex        =   5
         Top             =   285
         Width           =   765
         _ExtentX        =   1349
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
         Caption         =   "frmOptChartEdit3d.frx":013C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":0168
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":0188
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optGridLines 
         Height          =   220
         Index           =   3
         Left            =   3510
         TabIndex        =   4
         Top             =   285
         Width           =   765
         _ExtentX        =   1349
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
         Caption         =   "frmOptChartEdit3d.frx":01A4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":01CC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":01EC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboRotateDetail 
      Height          =   315
      Left            =   3330
      TabIndex        =   11
      Top             =   930
      Width           =   1290
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
      Tip             =   "frmOptChartEdit3d.frx":0208
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   2
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
      MouseIcon       =   "frmOptChartEdit3d.frx":0228
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraLegendLocation 
      Height          =   600
      Left            =   180
      TabIndex        =   12
      Top             =   1365
      Width           =   4440
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptChartEdit3d.frx":0244
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptChartEdit3d.frx":0282
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit3d.frx":02A2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optLegend 
         Height          =   220
         Index           =   4
         Left            =   3555
         TabIndex        =   13
         Top             =   300
         Width           =   750
         _ExtentX        =   1323
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
         Caption         =   "frmOptChartEdit3d.frx":02BE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":02E6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":0306
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLegend 
         Height          =   220
         Index           =   3
         Left            =   2673
         TabIndex        =   25
         Top             =   300
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
         Caption         =   "frmOptChartEdit3d.frx":0322
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":034C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":036C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLegend 
         Height          =   220
         Index           =   2
         Left            =   1882
         TabIndex        =   26
         Top             =   300
         Width           =   645
         _ExtentX        =   1138
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
         Caption         =   "frmOptChartEdit3d.frx":0388
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":03B0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":03D0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLegend 
         Height          =   220
         Index           =   1
         Left            =   896
         TabIndex        =   27
         Top             =   300
         Width           =   840
         _ExtentX        =   1482
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
         Caption         =   "frmOptChartEdit3d.frx":03EC
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":0418
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":0438
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLegend 
         Height          =   220
         Index           =   0
         Left            =   90
         TabIndex        =   28
         Top             =   300
         Width           =   660
         _ExtentX        =   1164
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
         Caption         =   "frmOptChartEdit3d.frx":0454
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":047A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":049A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraColors 
      Height          =   1875
      Left            =   180
      TabIndex        =   14
      Top             =   2040
      Width           =   4440
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptChartEdit3d.frx":04B6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptChartEdit3d.frx":0522
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit3d.frx":0542
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkSingleColor 
         Height          =   285
         Left            =   105
         TabIndex        =   29
         Top             =   300
         Width           =   2490
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptChartEdit3d.frx":055E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":05A8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":05C8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   15
         Top             =   870
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   1
         Left            =   993
         TabIndex        =   16
         Top             =   870
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   2
         Left            =   1881
         TabIndex        =   17
         Top             =   870
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   3
         Left            =   2769
         TabIndex        =   18
         Top             =   870
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   4
         Left            =   3660
         TabIndex        =   19
         Top             =   870
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   5
         Left            =   105
         TabIndex        =   20
         Top             =   1470
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   6
         Left            =   993
         TabIndex        =   21
         Top             =   1470
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   7
         Left            =   1881
         TabIndex        =   22
         Top             =   1470
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   8
         Left            =   2769
         TabIndex        =   23
         Top             =   1470
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColors 
         Height          =   285
         Index           =   9
         Left            =   3660
         TabIndex        =   24
         Top             =   1470
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   9
         Left            =   3810
         Top             =   1260
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
         Caption         =   "frmOptChartEdit3d.frx":05E4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":060C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":062C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   8
         Left            =   2925
         Top             =   1260
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
         Caption         =   "frmOptChartEdit3d.frx":0648
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":066E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":068E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   7
         Left            =   2040
         Top             =   1260
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
         Caption         =   "frmOptChartEdit3d.frx":06AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":06D0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":06F0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   6
         Left            =   1155
         Top             =   1260
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
         Caption         =   "frmOptChartEdit3d.frx":070C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":0732
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":0752
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   5
         Left            =   270
         Top             =   1260
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
         Caption         =   "frmOptChartEdit3d.frx":076E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":0794
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":07B4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   4
         Left            =   3840
         Top             =   660
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
         Caption         =   "frmOptChartEdit3d.frx":07D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":07F6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":0816
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   3
         Left            =   2946
         Top             =   660
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
         Caption         =   "frmOptChartEdit3d.frx":0832
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":0858
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":0878
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   2
         Left            =   2054
         Top             =   660
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
         Caption         =   "frmOptChartEdit3d.frx":0894
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":08BA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":08DA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   1
         Left            =   1162
         Top             =   660
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
         Caption         =   "frmOptChartEdit3d.frx":08F6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":091C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":093C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblColors 
         Height          =   240
         Index           =   0
         Left            =   270
         Top             =   660
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
         Caption         =   "frmOptChartEdit3d.frx":0958
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":097E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":099E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboDataY 
      Height          =   315
      Left            =   1155
      TabIndex        =   10
      Top             =   90
      Width           =   3465
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
      Tip             =   "frmOptChartEdit3d.frx":09BA
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   2
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
      MouseIcon       =   "frmOptChartEdit3d.frx":09DA
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboChartType 
      Height          =   315
      Left            =   1155
      TabIndex        =   9
      Top             =   507
      Width           =   3465
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
      Tip             =   "frmOptChartEdit3d.frx":09F6
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   2
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
      MouseIcon       =   "frmOptChartEdit3d.frx":0A16
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboFontSize 
      Height          =   315
      Left            =   1155
      TabIndex        =   8
      Top             =   930
      Width           =   945
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
      Tip             =   "frmOptChartEdit3d.frx":0A32
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   2
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
      MouseIcon       =   "frmOptChartEdit3d.frx":0A52
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   600
      Left            =   892
      TabIndex        =   0
      Top             =   4620
      Width           =   3240
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptChartEdit3d.frx":0A6E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptChartEdit3d.frx":0AA2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit3d.frx":0AC2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   330
         Left            =   1650
         TabIndex        =   2
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
         Caption         =   "frmOptChartEdit3d.frx":0ADE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":0B0C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":0B2C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   330
         Left            =   765
         TabIndex        =   1
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
         Caption         =   "frmOptChartEdit3d.frx":0B48
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit3d.frx":0B6E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit3d.frx":0B8E
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label4 
      Height          =   225
      Left            =   2190
      Top             =   975
      Width           =   1350
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptChartEdit3d.frx":0BAA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit3d.frx":0BE8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit3d.frx":0C08
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAxisY 
      Height          =   225
      Left            =   192
      Top             =   135
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
      Caption         =   "frmOptChartEdit3d.frx":0C24
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit3d.frx":0C60
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit3d.frx":0C80
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label2 
      Height          =   225
      Left            =   214
      Top             =   552
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
      Caption         =   "frmOptChartEdit3d.frx":0C9C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit3d.frx":0CD2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit3d.frx":0CF2
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label3 
      Height          =   225
      Left            =   210
      Top             =   960
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
      Caption         =   "frmOptChartEdit3d.frx":0D0E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit3d.frx":0D40
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit3d.frx":0D60
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmOptChartEdit3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    fmOptChart As Form
    Pe3d As Pe3doa
    aColors As cGdArray
End Type
Private m As mPrivate

Public Sub ShowMe(fmForm As Form, aColors As cGdArray, Pe3d As Pe3doa)
On Error GoTo ErrSection:

    Set m.fmOptChart = fmForm
    Set m.Pe3d = Pe3d
    Set m.aColors = aColors
    
    If m.fmOptChart Is Nothing Or m.Pe3d Is Nothing Or _
        m.aColors Is Nothing Then
        Unload Me
        Exit Sub
    End If
    
    If m.aColors.Size < 1 Then
        Unload Me
        Exit Sub
    End If
    
    InitControls
        
    CenterTheForm Me
    Me.Show 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.ShowMe", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitControls()
On Error GoTo ErrSection:

    m.fmOptChart.InitComboY cboDataY
    InitChartType
    InitFontSize
    InitLegendLocation
    InitColors
    InitGridLines
    InitSingleColor
    InitRotateDetail

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.InitControls", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub UpdateChart()
On Error GoTo ErrSection:
    
    'data column to plot
    m.fmOptChart.IdxAxisY = cboDataY.ItemData(cboDataY.ListIndex)
    'flag for whether to use only one color for Bar chart type
    m.fmOptChart.SingleColor = chkSingleColor.Value
    
    SetChartType
    SetFontSize
    SetLegendLocation
    SetColors
    SetGridLines
    SetRotateDetail
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.UpdateChart", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboChartType_Click()
On Error GoTo ErrSection:

    If cboChartType.ListIndex > 2 Then
        chkSingleColor.Enabled = False      'area
        EnableLegend True
    Else
        chkSingleColor.Enabled = True       'bar
        EnableLegend False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.cboChartType.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:
    
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdExport_Click()
On Error GoTo ErrSection:
    
    m.Pe3d.PEactions = 6

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.cmdExport.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    UpdateChart
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitChartType()
On Error GoTo ErrSection:

    cboChartType.Clear
    
    cboChartType.AddItem "Bar (wire frame)"
    cboChartType.AddItem "Bar (surface)"
    cboChartType.AddItem "Bar (surface with shading)"
    
    If m.fmOptChart.AllowAreaChart = 1 Then
        cboChartType.AddItem "Area (wire frame)"
        cboChartType.AddItem "Area (surface)"
        cboChartType.AddItem "Area (surface with shading)"
        cboChartType.AddItem "Area (surface with contour)"
    End If
    
    If m.Pe3d.PolyMode = PEPM_3DBAR Then
        cboChartType.ListIndex = m.Pe3d.PlottingMethod
    ElseIf m.Pe3d.PolyMode = PEPM_SURFACEPOLYGONS Then
        Select Case m.Pe3d.PlottingMethod
            Case TDPM_0
                cboChartType.ListIndex = 3
            Case TDPM_1
                cboChartType.ListIndex = 4
            Case TDPM_2
                cboChartType.ListIndex = 5
            Case TDPM_4     'case 3 is surface with pixes (not using)
                cboChartType.ListIndex = 6
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.InitChartType", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitSingleColor()
On Error GoTo ErrSection:

    If m.Pe3d.PolyMode = PEPM_3DBAR Then
        chkSingleColor.Value = m.fmOptChart.SingleColor
        chkSingleColor.Enabled = True
    Else
        chkSingleColor.Enabled = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.InitSingleColor", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitColors()
On Error GoTo ErrSection:

    Dim i&

    For i = 0 To 9
        gdColors(i).Color = m.aColors(i)
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.InitColors", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitLegendLocation()
On Error GoTo ErrSection:

    Dim i&

    If m.Pe3d.PolyMode = PEPM_3DBAR Then
        EnableLegend False
    Else
        EnableLegend True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.InitLegendLocation", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitGridLines()
On Error GoTo ErrSection:

    Select Case m.Pe3d.GridLineControl
        Case PEGLC_BOTH
            optGridLines(0) = True
        Case PEGLC_XAXIS
            optGridLines(1) = True
        Case PEGLC_YAXIS
            optGridLines(2) = True
        Case PEGLC_NONE
            optGridLines(3) = True
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.InitGridLines", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitFontSize()
On Error GoTo ErrSection:

    cboFontSize.Clear
    
    cboFontSize.AddItem "Small"
    cboFontSize.AddItem "Medium"
    cboFontSize.AddItem "Large"

    Select Case m.Pe3d.Font.Size
        Case PEFS_SMALL
            cboFontSize.ListIndex = 0
        Case PEFS_MEDIUM
            cboFontSize.ListIndex = 1
        Case PEFS_LARGE
            cboFontSize.ListIndex = 2
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.InitFontSize", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitRotateDetail()
On Error GoTo ErrSection:

    cboRotateDetail.Clear
    
    cboRotateDetail.AddItem "Wire Frame"
    cboRotateDetail.AddItem "Plotting Style"
    cboRotateDetail.AddItem "Full Detail"
    
    cboRotateDetail.ListIndex = m.fmOptChart.RotateDetail

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.InitRotateDetail", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub EnableLegend(ByVal bEnable As Boolean)
On Error GoTo ErrSection:

    Dim i&

    fraLegendLocation.Enabled = bEnable
    For i = 0 To 4
        optLegend(i).Enabled = bEnable
    Next
    
    If bEnable = True Then
        i = m.fmOptChart.LegendLocation
        If i = -1 Then
            optLegend(4) = True
        Else
            optLegend(i) = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.EnableLegend", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub SetChartType()
On Error GoTo ErrSection:

    Dim nPolyMode&, nPlotMethod&

    If cboChartType.ListIndex < 3 Then
        nPolyMode = PEPM_3DBAR
    Else
        nPolyMode = PEPM_SURFACEPOLYGONS
    End If
    
    Select Case cboChartType.ListIndex
        Case 0, 3
            nPlotMethod = TDPM_0  'wire frame
        Case 1, 4
            nPlotMethod = TDPM_1  'surface
        Case 2, 5
            nPlotMethod = TDPM_2  'surface with shading
        Case 6
            nPlotMethod = TDPM_4  'surface with contour
        Case Else
            nPlotMethod = TDPM_0  'wire frame
    End Select
    
    m.fmOptChart.SetChartType3d nPolyMode, nPlotMethod

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.SetChartType", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub SetFontSize()
On Error GoTo ErrSection:

    Select Case cboFontSize.ListIndex
        Case 0
            m.fmOptChart.PeFontSize = PEFS_SMALL
        Case 1
            m.fmOptChart.PeFontSize = PEFS_MEDIUM
        Case 2
            m.fmOptChart.PeFontSize = PEFS_LARGE
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.SetFontSize", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub SetLegendLocation()
On Error GoTo ErrSection:

    If optLegend(0) = True Then
        m.fmOptChart.LegendLocation = PELL_TOP
    ElseIf optLegend(1) = True Then
        m.fmOptChart.LegendLocation = PELL_BOTTOM
    ElseIf optLegend(2) = True Then
        m.fmOptChart.LegendLocation = PELL_LEFT
    ElseIf optLegend(3) = True Then
        m.fmOptChart.LegendLocation = PELL_RIGHT
    ElseIf optLegend(4) = True Then
        m.fmOptChart.LegendLocation = -1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.SetLegendLocation", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub SetColors()
On Error GoTo ErrSection:

    Dim i&

    For i = 0 To 9
        m.aColors(i) = gdColors(i).Color
    Next
    
    m.fmOptChart.Color = m.aColors(0)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.SetColors", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub SetGridLines()
On Error GoTo ErrSection:

    If optGridLines(0) = True Then
        m.Pe3d.GridLineControl = PEGLC_BOTH
    ElseIf optGridLines(1) = True Then
        m.Pe3d.GridLineControl = PEGLC_XAXIS
    ElseIf optGridLines(2) = True Then
        m.Pe3d.GridLineControl = PEGLC_YAXIS
    ElseIf optGridLines(3) = True Then
        m.Pe3d.GridLineControl = PEGLC_NONE
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.SetGridLines", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub SetRotateDetail()
On Error GoTo ErrSection:
    
    m.fmOptChart.RotateDetail = cboRotateDetail.ListIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.SetRotateDetail", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Me.Icon = Picture16(ToolbarIcon("ID_News"))
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.fmOptChart = Nothing
    Set m.Pe3d = Nothing
    Set m.aColors = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit3d.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

