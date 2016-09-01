VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTemplates 
   Caption         =   "Manage Chart Templates"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   Icon            =   "frmTemplates.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Interval        =   2000
      Left            =   4800
      Top             =   5880
   End
   Begin HexUniControls.ctlUniFrameWL fraPageCollection 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   60
      Width           =   4935
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmTemplates.frx":0442
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTemplates.frx":048C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTemplates.frx":04AC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboCollection 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   1335
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
         Tip             =   "frmTemplates.frx":04C8
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":04E8
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCollectionDelete 
         Height          =   315
         Left            =   3840
         TabIndex        =   17
         Top             =   300
         Width           =   855
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
         Caption         =   "frmTemplates.frx":0504
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":0530
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0550
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCollectionImport 
         Height          =   315
         Left            =   2880
         TabIndex        =   16
         Top             =   300
         Width           =   855
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
         Caption         =   "frmTemplates.frx":056C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":0598
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":05B8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCollectionExport 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   300
         Width           =   855
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
         Caption         =   "frmTemplates.frx":05D4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":0600
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0620
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3915
      Left            =   3480
      TabIndex        =   7
      Top             =   960
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
      Caption         =   "frmTemplates.frx":063C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTemplates.frx":0670
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTemplates.frx":0690
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCopy 
         Height          =   375
         Left            =   60
         TabIndex        =   12
         Top             =   2040
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
         Caption         =   "frmTemplates.frx":06AC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":06DC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0734
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   1140
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
         Caption         =   "frmTemplates.frx":0750
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":0778
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":07F2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRename 
         Height          =   375
         Left            =   60
         TabIndex        =   5
         Top             =   2640
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
         Caption         =   "frmTemplates.frx":080E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":083C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0894
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   0
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
         Caption         =   "frmTemplates.frx":08B0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":08DC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":08FC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSwitch 
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   480
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
         Caption         =   "frmTemplates.frx":0918
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":094C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":09C8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   3120
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
         Caption         =   "frmTemplates.frx":09E4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":0A12
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0A6A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSave 
         Height          =   375
         Left            =   60
         TabIndex        =   4
         Top             =   1590
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
         Caption         =   "frmTemplates.frx":0A86
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":0AB6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0B36
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraMove 
      Height          =   1095
      Left            =   60
      TabIndex        =   8
      Top             =   5040
      Width           =   3795
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmTemplates.frx":0B52
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTemplates.frx":0B7E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTemplates.frx":0B9E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSort 
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   0
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
         Caption         =   "frmTemplates.frx":0BBA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":0BEC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0C0C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkPrompt 
         Height          =   225
         Left            =   0
         TabIndex        =   11
         Top             =   840
         Width           =   3795
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTemplates.frx":0C28
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmTemplates.frx":0C90
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0CB0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdMoveDown 
         Height          =   315
         Left            =   1380
         TabIndex        =   10
         Top             =   0
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
         Caption         =   "frmTemplates.frx":0CCC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":0CFE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0D1E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdMoveUp 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   0
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
         Caption         =   "frmTemplates.frx":0D3A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTemplates.frx":0D68
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0D88
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblNote 
         Height          =   435
         Left            =   60
         Top             =   360
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
         Caption         =   "frmTemplates.frx":0DA4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTemplates.frx":0E7C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTemplates.frx":0E9C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgList 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1455
      _cx             =   2566
      _cy             =   3836
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
Attribute VB_Name = "frmTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kMinWidth = 4875
Private Const kReservedName = "My Collection"
Private Const kCollectionInfbox = "Chart Page Collection"

Public Enum eTemplateFormMode
    eMode_Templates = 0
    eMode_Pages = 1
    eMode_ChartTabs = 2
    eMode_SecSubCom = 3         'sectors, subsectors, components
    eMode_ChartCopyMove = 4
End Enum

Private Enum eGDCols
    eGDCol_Favorites = 0
    eGDCol_Name = 1
    eGDCol_Group = 2
End Enum

Private Enum eFileAction
    eFileAction_Create
    eFileAction_DeleteFile
    eFileAction_Copy
End Enum

Private Type mPrivate
    eMode As eTemplateFormMode
    Chart As cChart
    strPath As String
    bDirty As Boolean
    bShowCPC As Boolean
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Sub cboCollection_Click()
On Error Resume Next

    If Me.Visible Then
        With cboCollection
            If .ListIndex >= 0 And .ListIndex < .ListCount Then
                PageCollectionChange cboCollection.Text
            End If
        End With
    End If
    
End Sub

Private Sub chkPrompt_Click()
On Error GoTo ErrSection:

    Dim strMsg$
    
    If Me.Visible Then
        If m.eMode = eMode_Templates Then
            If chkPrompt.Value = 0 Then
                strMsg = "Turning this setting off means you will |no longer be prompted to save changes |made to a chart when switching templates.||Are you sure?"
                If InfBox(strMsg, "i", "+OK|-Cancel", "WARNING ...") = "C" Then chkPrompt.Value = 1
            End If
            SetIniFileProperty "SaveTemplatePrompt", chkPrompt.Value, "Charting", g.strIniFile      '4695
        ElseIf m.eMode = eMode_Pages Then
            SetIniFileProperty "AutoSavePage", chkPrompt.Value, "Charting", g.strIniFile
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.chkPrompt.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdCollectionDelete_Click()
On Error GoTo ErrSection:

    Dim strTemp$, strCollection$, strErr$, cboIndex&
    ''Dim fsoSysObj As New FileSystemObject
    
    'get collection name from dropdown box
    cboIndex = cboCollection.ListIndex
    If cboIndex <= 0 Or cboIndex >= cboCollection.ListCount Then Exit Sub

    'prompt deletion confirm
    strCollection = cboCollection.Text
    strTemp = "Delete all chart pages from the collection:" & vbCrLf & vbTab & strCollection & "?"
    If MsgBox(strTemp, vbYesNo, kCollectionInfbox) <> vbYes Then Exit Sub

    'save CPC file info for the current collection
    strTemp = g.strAppPath & "\charts\collection\" & strCollection & ".cpc"
    strCollection = g.ChartGlobals.strCPCRoot
    
    'remove item from combo box
    cboCollection.RemoveItem cboIndex
    
    'delete CPC file
    strErr = "fsoSysObj.DeleteFile" & vbCrLf & strTemp
    ''If fsoSysObj.FileExists(strTemp) Then fsoSysObj.DeleteFile strTemp, True
    If FileExist(strTemp) Then DeleteFiles strTemp, True
    
    'delete shared collection folders
    strErr = "fsoSysObj.DeleteFolder" & vbCrLf & strCollection
    ''If fsoSysObj.FolderExists(strCollection) Then fsoSysObj.DeleteFolder strCollection, True
    
    If DirExist(strCollection) Then DeleteFolder AddSlash(strCollection), True
    
    'change to the default local collection
    'click event of cbo will trigger page collection load
    cboCollection.ListIndex = 0
    
    Me.Hide
    tmr.Tag = ""

ErrExit:
    tmr.Enabled = True          'aardvark 6914
    
    Exit Sub

ErrSection:
    
    RaiseError "frmTemplates.cmdCollectionDelete_Click"
    If Len(strErr) > 0 Then InfBox strErr, "E", , kCollectionInfbox

End Sub

Private Sub cmdCollectionExport_Click()
On Error GoTo ErrSection

    Dim i&, iCount&, dResult#, strErr$
    
    Dim strTemp$, strTempDir$, strFileCPC$
    Dim strSource$, strDest$
    
    ''Dim fsoSysObj As New FileSystemObject
    Dim bFavorites As Boolean
    
    strTemp = InfBox("Export favorites or all pages?", "?", "+Favorites|All|Cancel", kCollectionInfbox)
    If strTemp = "C" Then Exit Sub
    
    If strTemp = "F" Then
        bFavorites = True
    
        'make sure there are favorites
        With fgList
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, eGDCol_Favorites) = flexChecked Then
                    iCount = 1
                    Exit For
                End If
            Next
        End With
        If iCount = 0 Then
            InfBox "You do not have any pages marked as favorites.", "I", , kCollectionInfbox
            GoTo ErrExit
        End If
    End If
    
    'get name of current collection
    If cboCollection.Text <> kReservedName Then
        strTemp = Trim(cboCollection.Text)
    Else
        strTemp = ""
    End If
    
    'prompt for collection export name until valid or cancelled
    Do
        strTemp = Trim(InfBox("Enter name of collection owner", "?", _
            "+Ok|-Cancel", "Export chart page collection", , , , , , "s", strTemp))
        If Len(strTemp) = 0 Then GoTo ErrExit
        If strTemp = kReservedName Then
            InfBox strTemp & " is a reserved collection name. Please enter a different name.", "I", , kCollectionInfbox
        ElseIf IsValidFileBase(strTemp) Then
            Exit Do
        End If
    Loop
    
    'prompt for location to export collection TO
    strFileCPC = CommonDialogFile(frmMain.CommonDialog1, True, "Chart page collection (*.cpc)|*.cpc", strTemp)
    If Len(strFileCPC) = 0 Then GoTo ErrExit
    
    'prompt to overwrite existing
    If FileExist(strFileCPC) Then
        Do
            strTemp = InfBox("The collection " & strFileCPC & " exists. Overwrite?", "?", "+Yes|No|Cancel", kCollectionInfbox)
            
            If strTemp = "C" Then
                GoTo ErrExit
            ElseIf strTemp = "N" Then
                Do
                    'prompt for collection export name until valid or cancelled
                    strTemp = Trim(InfBox("Name for this chart page collection:", "?", _
                        "+Ok|-Cancel", "Save Chart Page Collection", , , , , , "s", strTemp))
                    If Len(strTemp) = 0 Then GoTo ErrExit
                    If IsValidFileBase(strTemp) Then Exit Do
                Loop
                strFileCPC = CommonDialogFile(frmMain.CommonDialog1, True, "Chart page collection (*.cpc)|*.cpc", strTemp)
                If Len(strFileCPC) = 0 Then GoTo ErrExit
            End If
            
            If Not FileExist(strFileCPC) Or strTemp = "Y" Then Exit Do
        Loop
    End If
    
    If bFavorites Then
        strTempDir = g.ChartGlobals.strCPCRoot & "\charts\pages\TempCollection"
        'delete temp directory recursively if exists
        ''If fsoSysObj.FolderExists(strTempDir) Then
        If DirExist(strTempDir) Then
            strErr = "fsoSysObj.DeleteFolder" & vbCrLf & strTempDir
            ''fsoSysObj.DeleteFolder strTempDir, True
            DeleteFolder AddSlash(strTempDir), True
        End If
        
        'create temp directory
        strErr = "fsoSysObj.CreateFolder" & vbCrLf & strTempDir
        ''fsoSysObj.CreateFolder strTempDir
        MakeDir strTempDir
        strErr = ""

        'copy favoites page files to temp directory
        With fgList
            For i = .FixedRows To .Rows - 1
                strTemp = ""
                strSource = ""
                strDest = ""
                If .Cell(flexcpChecked, i, eGDCol_Favorites) = flexChecked Then
                    strTemp = .TextMatrix(i, GDCol(eGDCol_Name))
                    strSource = g.ChartGlobals.strCPCRoot & "\charts\pages\" & strTemp & ".gzp"
                    strDest = strTempDir & "\" & strTemp & ".gzp"
                
                    strErr = "fsoSysObj.CopyFile" & vbCrLf & strSource & " : " & strDest
                    ''fsoSysObj.CopyFile strSource, strDest, True
                    CopyFiles strSource, strDest, True
                    strErr = ""
                End If
            Next
        End With
    End If
    
    If bFavorites Then
        'zip up favorite pages (*.gzp) files from temp directory
        dResult = ZipExecute("C", strFileCPC, strTempDir, "*.*")
        strErr = "fsoSysObj.DeleteFolder" & vbCrLf & strTempDir
        ''fsoSysObj.DeleteFolder strTempDir, True
        DeleteFolder AddSlash(strTempDir), True
        strErr = ""
    Else
        'zip up all pages (*.gzp) files
        strTempDir = g.ChartGlobals.strCPCRoot & "\charts\pages"
        dResult = ZipExecute("C", strFileCPC, strTempDir, "*.gzp")
    End If
    
    If Not FileExist(strFileCPC) Or dResult = 0 Then
        InfBox "Export to collection " & strFileCPC & "failed.", "E", , kCollectionInfbox
        If dResult = 0 And FileExist(strFileCPC) Then
            strErr = "fsoSysObj.DeleteFile" & vbCrLf & strFileCPC
            ''fsoSysObj.DeleteFile strFileCPC, True
            DeleteFiles strFileCPC, True
        End If
    Else
        'append files in current charts folder (*.cht, *.ano, charts.cfg, page.ini)
        strTempDir = g.ChartGlobals.strCPCRoot & "\Charts"
        dResult = ZipExecute("A", strFileCPC, strTempDir, "* /i=CHT,Charts.cfg,INI,ANO")
        If dResult = 0 Then
            InfBox "Export to collection " & strFileCPC & "failed.", "E", , kCollectionInfbox
            strErr = "fsoSysObj.DeleteFile" & vbCrLf & strFileCPC
            ''fsoSysObj.DeleteFile strFileCPC, True
            DeleteFiles strFileCPC, True
        End If
    End If
    
    tmr.Tag = "ZORDER"
    
ErrExit:
    tmr.Enabled = True          'aardvark 6914

    Exit Sub

ErrSection:
    RaiseError "frmTemplates.cmdCollectionExport_Click"
    
    'if there was file system error then the RaiseError will give system error msg
    'additional info is available in local string as set in code above
    'e.g. system may give invalid path error, strErr will give path & attempted action
    If Len(strErr) > 0 Then InfBox strErr, "E", , kCollectionInfbox
    
End Sub

Private Sub cmdCollectionImport_Click()
On Error GoTo ErrSection:
    
    Dim strErr$, i&
    Dim bCpcExists As Boolean
    
    Dim strFileCPC$, strDest$
    Dim strTempDir$, strSharedRoot$
    
    ''Dim fsoSysObj As New FileSystemObject

    'prompt for collection to import
    strFileCPC = CommonDialogFile(frmMain.CommonDialog1, False, "Chart page collection (*.cpc)|*.cpc")
    If Len(strFileCPC) = 0 Then GoTo ErrExit
    
    'make sure collection directory exists on local drive
    strTempDir = g.strAppPath & "\charts\Collection"
    ''If Not fsoSysObj.FolderExists(strTempDir) Then
    If Not DirExist(strTempDir) Then
        strErr = "fsoSysObj.CreateFolder" & vbCrLf & strTempDir
        ''fsoSysObj.CreateFolder strTempDir
        MakeDir strTempDir
    End If
    strErr = ""
    
    'prompt to overwrite CPC file if exists
    strDest = strTempDir & "\" & FileBase(strFileCPC) & ".cpc"
    bCpcExists = FileExist(strDest) ''fsoSysObj.FileExists(strDest)
    If bCpcExists Then
        If InfBox("The collection " & FileBase(strFileCPC) & " exists. Overwrite?", "?", "+Yes|No", kCollectionInfbox) = "N" Then Exit Sub
    End If
    
    'copy CPC file to local drive
    strErr = "fsoSysObj.CopyFile" & vbCrLf & strFileCPC & " : " & strDest
    ''fsoSysObj.CopyFile strFileCPC, strDest, True
    CopyFiles strFileCPC, strDest, True
    strErr = ""
    
    strFileCPC = FileBase(strFileCPC)
    ExtractCollection strFileCPC, True
    
    If Not bCpcExists Then
        cboCollection.AddItem strFileCPC
        cboCollection.ListIndex = cboCollection.ListCount - 1 'this will trigger chart page collection load
    ElseIf cboCollection.Text = strFileCPC Then
        'user imported a collection with same name as currently loaded collection
        'could be because user wants to "revert" back to original collection
        'or user just chose to overwrite the loaded collection for whatever reason
        PageCollectionChange strFileCPC, False
    Else
        For i = 1 To cboCollection.ListCount - 1
            If cboCollection.List(i) = strFileCPC Then
                cboCollection.ListIndex = i
                Exit For
            End If
        Next
    End If

    tmr.Tag = ""
    
ErrExit:
    tmr.Enabled = True          'aardvark 6914

    Exit Sub

ErrSection:
    RaiseError "frmTemplates.cmdCollectionImport_Click"
    
    'if there was file system error then the RaiseError will give system error msg
    'additional info is available in local string as set in code above
    'e.g. system may give invalid path error, strErr will give path & attempted action
    If Len(strErr) > 0 Then InfBox strErr, "E", , kCollectionInfbox

End Sub

Private Sub cmdCopy_Click()
On Error GoTo ErrSection:

    Dim i&
    Dim strPage$, strSource$, strDest$, strAns$
    
    If m.eMode = eMode_Pages Then
        strPage = Trim(fgList.TextMatrix(fgList.Row, GDCol(eGDCol_Name)))
        
        If Len(strPage) > 0 Then
            strSource = "Would you like to copy the selected page or all pages to the collection called 'My Collection'.?"
            
            strAns = InfBox(strSource, "?", "Selected|All|Cancel", "Confirmation")
            
            If strAns = "S" Then
                strDest = App.Path & "\Charts\Pages\" & strPage & ".GZP"
            
                If FileExist(strDest) Then
                    strSource = "This will overwrite the existing " & strPage & " page in the collection called 'My Collection'."
                    
                    strAns = InfBox(strSource, "I", "+Ok|-Cancel", "Confirmation")
                End If
                                
                If strAns = "S" Or strAns = "O" Then
                    strSource = g.ChartGlobals.strCPCRoot & "\charts\pages\" & strPage & ".GZP"
                    FileCopy strSource, strDest
                End If
            ElseIf strAns = "A" Then
                strSource = "This will overwrite pages in the 'My Collection' that have the same names."
                If InfBox(strSource, "I", "+Ok|-Cancel", "Confirmation") = "O" Then
                    strDest = App.Path & "\Charts\Pages\"
                    strSource = g.ChartGlobals.strCPCRoot & "\charts\pages\*.GZP"
                    FileCopy strSource, strDest
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdCopy_Click"
    Resume ErrExit

End Sub

Private Sub cmdSwitch_Click()
On Error GoTo ErrSection:

    If m.eMode = eMode_ChartCopyMove Then
        CopyChartNewPage        '6806
        Unload Me
        Exit Sub
    End If

    Dim strTemplate$, bCtrlKey As Boolean

    strTemplate = fgList.TextMatrix(fgList.Row, GDCol(eGDCol_Name))
    
    If InStr(strTemplate, "<") > 0 Then
        Beep
        Exit Sub
    End If
    
    bCtrlKey = KeyIsPressed(VK_CONTROL)
    
    Select Case m.eMode
    Case eMode_Pages
        If Not bCtrlKey Then Me.Hide
        LoadChartPage strTemplate
    Case Else
        If Not m.Chart.TemplateApply(strTemplate) Then
            Beep
            Exit Sub
        End If
        If Not bCtrlKey Then Me.Hide
    End Select
    
    tmr.Enabled = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdSwitch.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdClose.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Dim strTemplate$, strMsg$
    
    strTemplate = fgList.TextMatrix(fgList.Row, GDCol(eGDCol_Name))
    
    If InStr(strTemplate, "<") > 0 Then
        Beep
        Exit Sub
    End If
    
    If strTemplate = fgList.TextMatrix(fgList.Row, eGDCol_Name) Then
        Select Case m.eMode
            Case eMode_Pages
                strMsg = "Delete the following chart page:|" & strTemplate
            Case Else
                strMsg = "Delete the following template:|" & strTemplate
        End Select
        
        If AskBox("i=? ; h=Confirm Delete ; b=+Delete|-Cancel ; " & strMsg) = "C" Then
            GoTo ErrExit
        End If
        
        Select Case m.eMode
            Case eMode_Pages
                KillFile m.strPath & strTemplate & ".GZP"
            Case Else
                KillFile m.strPath & strTemplate & ".CHT"
        End Select
    
        fgList.RemoveItem fgList.Row
        m.bDirty = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdDelete.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdMoveDown_Click()
On Error GoTo ErrSection:

    Dim nRow&
    
    With fgList
        nRow = .Row
        If nRow >= FirstRow And nRow < .Rows - 1 Then
            .RowPosition(nRow) = nRow + 1
            .Row = nRow + 1
            CheckFavorites
            .ShowCell .Row, eGDCol_Name
            'SetHotkeys
        Else
            Beep
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdMoveDown.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdMoveUp_Click()
On Error GoTo ErrSection:

    Dim nRow&
    
    With fgList
        nRow = .Row
        If nRow > FirstRow And nRow < .Rows Then
            .RowPosition(nRow) = nRow - 1
            .Row = nRow - 1
            CheckFavorites
            .ShowCell .Row, eGDCol_Name
            'SetHotkeys
        Else
            Beep
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdMoveUp.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    If m.eMode = eMode_ChartCopyMove Then
        If CopyChart(False) Then Unload Me
    Else
        SaveTemplate ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdNew.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdRename_Click()
On Error GoTo ErrSection:

    Dim strOld$, strNew$, strExt$, strMsg$
    
    strOld = fgList.TextMatrix(fgList.Row, GDCol(eGDCol_Name))
    
    If InStr(strOld, "<") > 0 Then
        Beep
        Exit Sub
    End If
    
    Select Case m.eMode
        Case eMode_Pages
            strNew = GetTemplateName(strOld, "Enter new name for this chart page ...")
            strExt = ".GZP"
        Case Else
            strNew = GetTemplateName(strOld, "Enter new name for this template ...")
            strExt = ".CHT"
    End Select
    If Len(strNew) = 0 Or strNew = strOld Then GoTo ErrExit
    
    If UCase(strNew) = UCase(strOld) Then
        ' just changing the case
        If FileExist(m.strPath & strOld & strExt) Then
            KillFile m.strPath & strNew & ".TMP"
            Name m.strPath & strOld & strExt As m.strPath & strNew & ".TMP"
            KillFile m.strPath & strNew & strExt
            Name m.strPath & strNew & ".TMP" As m.strPath & strNew & strExt
        End If
    Else
        ' change name of file
        If FileExist(m.strPath & strNew & strExt) Then
            Select Case m.eMode
            Case eMode_Pages
                strMsg = "Overwrite existing chart page?|(" & strNew & ")"
            Case Else
                strMsg = "Overwrite existing template?|(" & strNew & ")"
            End Select
            If AskBox("i=? ; h=Overwrite? ; b=Overwrite|+-Cancel ; " & strMsg) = "C" Then
                GoTo ErrExit
            End If
        End If
        
        If FileExist(m.strPath & strOld & strExt) Then
            KillFile m.strPath & strNew & strExt
            Name m.strPath & strOld & strExt As m.strPath & strNew & strExt
            ' change name of current page if it was renamed
            If m.eMode = eMode_Pages And UCase(g.strChartPage) = UCase(strOld) Then
                g.strChartPage = strNew
            End If
        End If
    End If
    
    ' update grid & save to LST file
    If strNew <> strOld Then
        Dim aTemp As New cGdArray
        
        With fgList
            .TextMatrix(.Row, GDCol(eGDCol_Name)) = strNew
            aTemp.SplitFields .RowData(.Row), vbTab
            aTemp(0) = strNew
            .RowData(.Row) = aTemp.JoinFields(vbTab)
            m.bDirty = True
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdRename.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrSection:

    Dim strTemplate$
    
    If m.eMode = eMode_ChartCopyMove Then
        If CopyChart(True) Then
            m.bDirty = True
            Unload Me
        End If
    Else
        strTemplate = Trim(fgList.TextMatrix(fgList.Row, GDCol(eGDCol_Name)))
        'If UCase(strTemplate) = "DEFAULT" Then Exit Sub
        If InStr(strTemplate, "<") > 0 Then strTemplate = "" 'NEW
    
        SaveTemplate strTemplate
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdSave.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdSort_Click()
On Error GoTo ErrSection:

'JM (11-10-2008) - original code, leave awhile then remove if all okay
'    fgList.Select fgList.FixedRows, GDCol(eGDCol_Name)
'    fgList.Sort = flexSortGenericAscending

    Dim aFavorites As New cGdArray, aOthers As New cGdArray
    Dim strText$, i&
    
    With fgList
        
        .Redraw = flexRDNone
        
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, eGDCol_Favorites) = flexChecked Then
                aFavorites.Add .RowData(i)
            Else
                aOthers.Add .RowData(i)
            End If
        Next
        
        aFavorites.Sort eGdSort_Default Or eGdSort_IgnoreCase
        aOthers.Sort eGdSort_Default Or eGdSort_IgnoreCase
        
        .Rows = .FixedRows
        
        For i = 0 To aFavorites.Size - 1
            strText = aFavorites(i)
            .AddItem "" & vbTab & Parse(strText, vbTab, 1)
            .Cell(flexcpChecked, .Rows - 1, eGDCol_Favorites) = flexChecked
            .Cell(flexcpPictureAlignment, .Rows - 1, eGDCol_Favorites) = flexAlignCenterCenter
            .TextMatrix(.Rows - 1, eGDCol_Group) = "F"
            .RowData(.Rows - 1) = strText
        Next
        
        For i = 0 To aOthers.Size - 1
            strText = aOthers(i)
            .AddItem "" & vbTab & Parse(strText, vbTab, 1)
            .Cell(flexcpChecked, .Rows - 1, eGDCol_Favorites) = flexUnchecked
            .Cell(flexcpPictureAlignment, .Rows - 1, eGDCol_Favorites) = flexAlignCenterCenter
            .TextMatrix(.Rows - 1, eGDCol_Group) = "N"
            .RowData(.Rows - 1) = strText
        Next
        
        .Redraw = flexRDBuffered
    
    End With
    
    m.bDirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.cmdSort_Click"
    Resume ErrExit
End Sub

Private Sub fgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim aTemp As New cGdArray
    Dim tb As New cGdTable
    Dim strText$, i&
    
    Dim bMoved As Boolean
    
    If Col = eGDCol_Favorites Then
            
        With fgList
            If Row >= .FixedRows And Row < .Rows Then
                aTemp.SplitFields .RowData(Row), vbTab
                If .Cell(flexcpChecked, Row, Col) = flexChecked Then
                    .TextMatrix(Row, eGDCol_Group) = "F"
                    aTemp(4) = "F"
                Else
                    .TextMatrix(Row, eGDCol_Group) = "N"
                    aTemp(4) = "N"
                End If
                .RowData(Row) = aTemp.JoinFields(vbTab)
                
                If Not m.bShowCPC Then          '6872
                    .Row = Row
                    If aTemp(4) = "F" Then
                        For i = .Row To .FixedRows + 1 Step -1
                            If i - 1 >= .FixedRows Then         'precautionary check
                                If .Cell(flexcpChecked, i - 1, eGDCol_Favorites) = flexChecked Then
                                    .RowPosition(Row) = i
                                    bMoved = True
                                    Exit For
                                End If
                            End If
                        Next
                    Else
                        For i = .Row To .Rows - 2
                            If i + 1 < .Rows Then               'precautionary check
                                If .Cell(flexcpChecked, i + 1, eGDCol_Favorites) = flexUnchecked Then
                                    .RowPosition(Row) = i
                                    bMoved = True
                                    Exit For
                                End If
                            Else
                                i = i
                            End If
                        Next
                    End If
                    
                    If Not bMoved Then
                        If .Cell(flexcpChecked, Row, Col) = flexChecked Then
                            'all items unchecked and just checked an item in the grid
                            .RowPosition(Row) = .FixedRows
                        Else
                            'all items checked and just unchecked an item in the grid
                            .RowPosition(Row) = .Rows - 1
                        End If
                    End If
                End If
                
                m.bDirty = True
            End If
        End With
        
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.fgList_AfterEdit"

End Sub

Private Sub fgList_AfterMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:
    
    fgList.Row = Position
    fgList.RowSel = Position
    CheckFavorites
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.fgList.AfterMoveRow", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.fgList.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If Col <> eGDCol_Favorites Then
        If m.eMode <> eMode_ChartCopyMove Then Cancel = True
    ElseIf m.eMode = eMode_Templates Then
        With fgList
            If Row = .FixedRows And UCase(.TextMatrix(Row, eGDCol_Name)) = "DEFAULT" Then
                Cancel = True
            End If
        End With
    End If

End Sub

Private Sub fgList_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim nRow&, nPos&
       
    If m.eMode = eMode_ChartCopyMove Then Exit Sub
    
    With fgList
        nRow = .MouseRow
        If nRow >= FirstRow + 1 And .Rows > FirstRow + 2 Then
            .Row = nRow
            .Refresh
            nPos = .DragRow(nRow)
            If nPos <> nRow Then
                Cancel = True
                'SetHotkeys
                EnableButtons
            End If
        End If
    End With
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.fgList.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_BeforeMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    If Row = Position Or m.eMode = eMode_ChartCopyMove Then Exit Sub
    
    If Row <= FirstRow Or Position <= FirstRow Then
        Position = Row
        Beep
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.fgList.BeforeMoveRow", eGDRaiseError_Show
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
    RaiseError "frmTemplates.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement$, strMode$
    
    g.Styler.StyleForm Me
    
    Select Case m.eMode
    Case eMode_Pages
        strMode = "Pages"
    Case eMode_Templates
        strMode = "Templates"
    Case eMode_ChartCopyMove
        strMode = "ChartCopyMove"
    End Select
    If Len(strMode) > 0 Then
        strPlacement = GetIniFileProperty(strMode, "", "Placement", g.strIniFile)
    End If
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    fraPageCollection.Enabled = m.bShowCPC
    fraPageCollection.Visible = m.bShowCPC
    
    cmdCopy.Enabled = m.bShowCPC
    cmdCopy.Visible = m.bShowCPC
    
    Me.Icon = Picture16(ToolbarIcon("kSelect"))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Function ShowMe(eMode As eTemplateFormMode, Optional Chart As cChart = Nothing) As String
On Error GoTo ErrSection:

    m.eMode = eMode
    Set m.Chart = Nothing
    m.bDirty = False
    
    Select Case m.eMode
    Case eMode_Templates
        If Chart Is Nothing Then Exit Function
        Set m.Chart = Chart
        m.bShowCPC = False
        m.strPath = App.Path & "\Charts\Templates\"
        Me.Caption = "Manage Chart Templates"
        chkPrompt.Caption = "Always prompt to save template before switching"
        cmdSwitch.ToolTipText = "Switch current chart to the selected template"
        cmdSave.ToolTipText = "Save current chart settings to selected template"
        cmdNew.ToolTipText = "Save current chart settings to a new template"
        cmdRename.ToolTipText = "Rename the selected template"
        cmdDelete.ToolTipText = "Delete the selected template"
        lblNote = "(Note: the hot-keys can be used from a chart to quickly switch to the associated chart template)"
        chkPrompt.Value = GetIniFileProperty("SaveTemplatePrompt", 1, "Charting", g.strIniFile)
        
        fraMove.Visible = True
    
    Case eMode_Pages
        If Not HasGold(True, "Saving chart pages") Then
            Exit Function
        End If
        m.bShowCPC = FileExist(g.strAppPath & "\ShowCPC.flg")
        m.strPath = g.ChartGlobals.strCPCRoot & "\Charts\Pages\"
        Me.Caption = "Manage Chart Pages"
        chkPrompt.Caption = "Auto-save when switching between chart pages"
        cmdSwitch.ToolTipText = "Switch to the selected chart page"
        cmdSave.ToolTipText = "Save current chart windows to selected page"
        cmdNew.ToolTipText = "Save current chart windows to a new page"
        cmdRename.ToolTipText = "Rename the selected chart page"
        cmdDelete.ToolTipText = "Delete the selected chart page"
        lblNote = "(Note: 'Ctrl-PgUp' and 'Ctrl-PgDn' can be used from a chart to move to the previous or next chart page)"
        chkPrompt.Value = GetIniFileProperty("AutoSavePage", 0, "Charting", g.strIniFile)
        
        fraMove.Visible = True
        
        GetPageCollectionList
        
    Case eMode_ChartCopyMove
        If Chart Is Nothing Then Exit Function
        Set m.Chart = Chart
        
        m.bShowCPC = False
        m.strPath = g.ChartGlobals.strCPCRoot & "\Charts\Pages\"
        Me.Caption = "Copy/Move Chart To Page"
        
        cmdSwitch.Caption = "New Page"          '6806
        cmdSwitch.ToolTipText = "Copy active chart to a new page"
        cmdSwitch.Top = cmdSwitch.Top + 150
        
        cmdNew.Caption = "Copy"
        cmdNew.ToolTipText = "Copy active chart to selected page"
        
        cmdSave.Caption = "Move"
        cmdSave.ToolTipText = "Move active chart to selected page"
        
        fraMove.Visible = False
        
        GetPageCollectionList
        
    Case Else
        Set m.Chart = Nothing
        Exit Function
    End Select
    
    FillList
    tmr.Enabled = False
    
    If Not Chart Is Nothing Then CenterFormOnChart Me, Chart            '6499
    
    ShowForm Me, eForm_Modal, , , ALT_GRID_ROW_COLOR
    
'    If m.eMode = eMode_Pages Then SetMainCaption
'
'    Set m.Chart = Nothing
    'Unload Me

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTemplates.ShowMe", eGDRaiseError_Raise
    
End Function

Private Sub FillList()
On Error GoTo ErrSection:

'Design mod from Pete:
'Favorites check box in manage.
'If you check favorites on one, it has to be in the top list.  At the bottom of the favorites already checked.
'If you wanted it higher, you can move it with up down or drag it.
'Make pages and Templates  the same.
'Put in a Separator line for blank spaces.  If someone sorts, the sorting is between separator lines and favorites.

'Implementation mod:
'Old fields: Name, ReqMod, Date/Time, Desc
'New fields: Name, ReqMod, Date/Time, Group
'
'New Group Field: The names "F" and "N" are reserved to indicate Favorites or None
'If group field exists then use it, otherwise assume first 9 belong to Favorites and append to string to be saved.

    Dim i&, hHandle&, strItem$, strApplied$
    Dim aItems As cGdArray, aFields As New cGdArray
    Dim bAllowed As Boolean
    
    Select Case m.eMode
    Case eMode_Pages
        Set aItems = GetAllowedList("P")
        If Len(g.strChartPage) > 0 Then strApplied = g.strChartPage
    Case eMode_ChartCopyMove
        Set aItems = GetAllowedList("P")
    Case Else
        Set aItems = GetAllowedList("T")
        If Not m.Chart Is Nothing Then
            strApplied = m.Chart.TemplateApplied
        End If
    End Select
        
    With fgList
        .Redraw = flexRDNone
        InitGrid
        
        .Row = -1 ' .FixedRows
        For i = 0 To aItems.Size - 1
            strItem = Parse(aItems(i), vbTab, 1)
            
            If m.eMode = eMode_ChartCopyMove And strItem = g.strChartPage Then
                'this is the current page, do not show in list for chart copy/move mode
            Else
                .AddItem "" & vbTab & strItem
                
                If Parse(aItems(i), vbTab, 5) = "F" Then
                    .Cell(flexcpChecked, .Rows - 1, eGDCol_Favorites) = flexChecked
                Else
                    .Cell(flexcpChecked, .Rows - 1, eGDCol_Favorites) = flexUnchecked
                End If
                
                .TextMatrix(.Rows - 1, eGDCol_Group) = Parse(aItems(i), vbTab, 5)
                .Cell(flexcpPictureAlignment, .Rows - 1, eGDCol_Favorites) = flexAlignCenterCenter
                .RowData(.Rows - 1) = aItems(i)
                
                If UCase(strApplied) = UCase(strItem) Then
                    .Row = .Rows - 1
                End If
            End If
        Next
        
        .Redraw = flexRDBuffered
    
        If .Row >= .FixedRows Then
            .ShowCell .Row, 0
            .Select .Row, 0, .Row, .Cols - 1
        End If
        
    End With
    
    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.FillList", eGDRaiseError_Raise
    
End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgList
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
        
        .Rows = .FixedRows
        .Cols = 3
        .TextMatrix(0, GDCol(eGDCol_Favorites)) = "Favorites"
        Select Case m.eMode
        Case eMode_Pages
            .TextMatrix(0, GDCol(eGDCol_Name)) = "Page Name"
            .AllowSelection = False
        Case eMode_ChartCopyMove
            .TextMatrix(0, GDCol(eGDCol_Name)) = "Page Name"
            .BackColorAlternate = ALT_GRID_ROW_COLOR
            .ColHidden(eGDCol_Favorites) = True
        Case Else
            .TextMatrix(0, GDCol(eGDCol_Name)) = "Template Name"
            .AllowSelection = False
        End Select
        .TextMatrix(0, eGDCol_Group) = "Group"
        .ColHidden(eGDCol_Group) = True
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize eGDCol_Favorites
        '.Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.InitGrid", eGDRaiseError_Raise
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If m.bDirty Then
        If m.eMode <> eMode_ChartCopyMove Then
            SaveList
            Cancel = True
        End If
        Me.Hide
    End If

    If m.eMode = eMode_Pages Then SetMainCaption
    Set m.Chart = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    Dim w&, h&
    
    'check minimum size
    w = kMinWidth
    h = fraButtons.Top + fraButtons.Height + fraMove.Height - cmdMoveDown.Height - 60
    If LimitFormSize(Me, w, h) Then Exit Sub

    fraButtons.Left = Me.ScaleWidth - fraButtons.Width
    
    If m.bShowCPC Then
    
        With fgList
            .Move .Left, .Top, fraButtons.Left - .Left * 2, _
                            Me.ScaleHeight - .Top - fraMove.Height - 60
        End With
    
        With fraPageCollection
            .Move fgList.Left, .Top, Me.ScaleWidth - fgList.Left * 2
            
            cmdCollectionDelete.Left = .Width - fgList.Left - cmdCollectionDelete.Width
            cmdCollectionImport.Left = cmdCollectionDelete.Left - cmdCollectionImport.Width - fgList.Left
            cmdCollectionExport.Left = cmdCollectionImport.Left - cmdCollectionExport.Width - fgList.Left
            cboCollection.Width = .Width - cmdCollectionDelete.Width * 3 - fgList.Left * 5
        End With
    
    ElseIf m.eMode = eMode_ChartCopyMove Then
        
        With fgList
            .Move .Left, fraPageCollection.Top + 30, fraButtons.Left - .Left * 2, _
                            Me.ScaleHeight - fraMove.Height / 4 + 60
        End With
    
    Else
    
        With fgList
            .Move .Left, fraPageCollection.Top + 30, fraButtons.Left - .Left * 2, _
                            Me.ScaleHeight - .Top - fraMove.Height - 60
        End With
    
    End If
    
    fraButtons.Top = fgList.Top + 15
    
    With fraMove
        .Move .Left, Me.ScaleHeight - .Height
    End With
    
    Me.Refresh

End Sub

Private Sub EnableButtons()
On Error GoTo ErrSection:

    Dim strText$, Count&, i&
    Dim aCfgFile As cGdArray
    Dim bChartMoveOk As Boolean
    
    If m.eMode = eMode_ChartCopyMove Then
        
        cmdRename.Visible = False
        cmdRename.Enabled = False
        cmdDelete.Visible = False
        cmdDelete.Enabled = False
        
        If cmdSave.Enabled Then         'this is renamed to "Move" button when in copy/move mode
            
            'enable chart move only if this page has more than one chart
            For i = 0 To Forms.Count - 1
                If IsFrmChart(Forms(i)) Then
                    Count = Count + 1
                End If
                
                If Count > 1 Then
                    bChartMoveOk = True
                    Exit For
                End If
            Next
            cmdSave.Enabled = bChartMoveOk
        
        End If
    
    Else
    
        With fgList
            If .Row >= .FixedRows Then
                strText = Trim(.TextMatrix(.Row, GDCol(eGDCol_Name)))
            End If
            
            If .Row < .Rows - 1 And .Row >= FirstRow Then
                Enable cmdMoveDown
            Else
                Disable cmdMoveDown
            End If
            If .Row > FirstRow Then
                Enable cmdMoveUp
            Else
                Disable cmdMoveUp
            End If
        End With
        
        If Len(strText) = 0 Then
            cmdSwitch.Enabled = False
            cmdSave.Enabled = False
            cmdDelete.Enabled = False
            cmdRename.Enabled = False
            cmdCopy.Enabled = False
        ElseIf InStr(strText, "<") > 0 Then
            cmdSwitch.Enabled = False
            cmdSave.Enabled = True
            cmdDelete.Enabled = False
            cmdRename.Enabled = False
            cmdCopy.Enabled = False
        ElseIf UCase(strText) = "DEFAULT" And m.eMode = eMode_Templates Then
            cmdSwitch.Enabled = True
            cmdSave.Enabled = True 'False
            cmdDelete.Enabled = False
            cmdRename.Enabled = False
        Else
            cmdSwitch.Enabled = True
            cmdSave.Enabled = True
            cmdDelete.Enabled = True
            cmdRename.Enabled = True
            cmdCopy.Enabled = m.bShowCPC
        End If

    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.EnableButtons", eGDRaiseError_Raise
    
End Sub

Private Function GetTemplateName(ByVal strName$, ByVal strMsg$) As String
On Error GoTo ErrSection:

    Do
        Select Case m.eMode
        Case eMode_Pages
            strName = Trim(AskBox("i=? ; g=str ; h=Page Name ; d=" & strName & " ; " & strMsg))
        Case Else
            strName = Trim(AskBox("i=? ; g=str ; h=Template Name ; d=" & strName & " ; " & strMsg))
        End Select
        If Len(strName) = 0 Then Exit Do
        If IsValidFileBase(strName) Then Exit Do
    Loop

    GetTemplateName = strName

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTemplates.GetTemplateName", eGDRaiseError_Raise
    
End Function

Private Sub SaveTemplate(ByVal strTemplate$)
On Error GoTo ErrSection:

    Dim strMsg$, strOld$, strNew$
    Dim iRedrawSave&, i&
    Dim bNewFile As Boolean
    
    If Len(strTemplate) = 0 Then
        'New template
        If fgList.Rows >= fgList.FixedRows + 10 And m.eMode = eMode_Templates Then
            If Not HasGold(True, "Creating more Chart Templates") Then
                Exit Sub
            End If
        ElseIf m.eMode = eMode_Pages Then
            If Not HasGold(True, "Creating new Pages") Then
                Exit Sub
            End If
        End If
        
        Select Case m.eMode
        Case eMode_Pages
            strMsg = "Name for new page|to save current chart windows ..."
        Case Else
            strMsg = "Name for new template|to save current chart settings ..."
        End Select
        strTemplate = GetTemplateName("", strMsg)
        If strTemplate = "" Then Exit Sub
        
        If m.bDirty Then SaveList       '4724
        
        'insert name of new template or page at bottom of favorites
        With fgList
            iRedrawSave = .Redraw
            .Redraw = flexRDNone
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, eGDCol_Favorites) = flexUnchecked Then
                    .AddItem vbTab & strTemplate, i
                    fgList.RowData(i) = strTemplate & vbTab & Date & vbTab & vbTab & "N"
                    .Cell(flexcpChecked, i, eGDCol_Favorites) = flexUnchecked
                    .Cell(flexcpPictureAlignment, i, eGDCol_Favorites) = flexAlignCenterCenter
                    Exit For
                End If
            Next
            .Redraw = iRedrawSave
        End With
        bNewFile = True
    Else
        'Save as existing template (no change in position of item so no need to save list)
        Select Case m.eMode
        Case eMode_Pages
            strMsg = "Save current chart windows as page:|" & strTemplate
        Case Else
            strMsg = "Save current chart settings as template:|" & strTemplate
        End Select
        If AskBox("i=? ; h=Confirm Save ; b=+Save|-Cancel ; " & strMsg) = "C" Then
            Exit Sub
        End If
    End If
    
    Select Case m.eMode
        Case eMode_Pages
            SaveChartPage strTemplate
        Case Else
            'first save template
            m.Chart.TemplateApplied = strTemplate
            m.Chart.TemplateSave
            
            'then copy to new template
            strOld = g.ChartGlobals.strCPCRoot & "\Charts\" & m.Chart.Template & ".CHT"
            strNew = m.strPath & strTemplate & ".CHT"
            If Not FileExist(strOld) Then
                InfBox "i=[] ; File could not be found:|" & strOld
            Else
                FileCopy strOld, strNew
            End If
    End Select

    If bNewFile Then
        If m.eMode = eMode_Templates Then
            'The templates list has more files than what the user is allowed to see
            'GetAllowedList will pick up the newly created file and save it correctly
            GetAllowedList "T", False
            FillList
        ElseIf m.eMode = eMode_Pages Then
            'Pages can be save once on unload
            m.bDirty = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.SaveTemplate", eGDRaiseError_Raise
    
End Sub

Private Sub SaveList()
On Error GoTo ErrSection:

    Dim i&, strItem$
    Dim aItems As New cGdArray
        
    For i = fgList.FixedRows To fgList.Rows - 1
        ''strItem = fgList.TextMatrix(i, GDCol(eGDCol_Name))
        strItem = fgList.RowData(i)
        If InStr(strItem, "<") = 0 And Len(strItem) > 0 Then
            aItems.Add strItem
        End If
    Next
        
    SaveTempPageList aItems, m.eMode
    
ErrExit:
    m.bDirty = False
    Exit Sub
    
ErrSection:
    RaiseError "frmTemplates.SaveList", eGDRaiseError_Raise
    
End Sub

Private Function RequiredMod(ByVal strFile$) As String
On Error Resume Next

    Dim fh%, strText$
    
    If m.eMode <> eMode_Templates Then Exit Function
    
    strFile = App.Path & "\Charts\Templates\" & strFile
    If FileExist(strFile) Then
        fh = FreeFile
        Open strFile For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strText
            Select Case UCase(Parse(strText, "=", 1))
            Case "END"
                Exit Do
            Case "REQUIRED"
                RequiredMod = UCase(Parse(strText, "=", 2))
                Exit Do
            End Select
        Loop
        Close #fh
    End If

End Function

Private Function FirstRow() As Long
    
    Select Case m.eMode
    Case eMode_Templates
        FirstRow = 1 '2
    Case Else
        FirstRow = 0 '1
    End Select

End Function

Private Sub Form_Unload(Cancel As Integer)

    Dim strMode$, i&
    
    Select Case m.eMode
    Case eMode_Pages
        strMode = "Pages"
    Case eMode_Templates
        strMode = "Templates"
    End Select
    
    If Len(strMode) > 0 Then
        SetIniFileProperty strMode, GetFormPlacement(Me), "Placement", g.strIniFile
    End If
    
    If m.eMode = eMode_ChartCopyMove And m.bDirty Then
        If Not ActiveChart Is Nothing Then
            ActiveChart.Tag = "MOVED"
            SendMessage ActiveChart.hWnd, WM_CLOSE, 0, 0
        End If
    End If
    
End Sub

Private Sub CheckFavorites()

    Dim aTemp As New cGdArray
    Dim bFix As Boolean
    
    With fgList
        If .Row = .FixedRows Then
            If .Cell(flexcpChecked, .Row, eGDCol_Favorites) = flexUnchecked Then bFix = True
        ElseIf .Row = .Rows - 1 Then
            If .Cell(flexcpChecked, .Row, eGDCol_Favorites) = flexChecked Then bFix = True
        ElseIf .Cell(flexcpChecked, .Row - 1, eGDCol_Favorites) = flexChecked And .Cell(flexcpChecked, .Row + 1, eGDCol_Favorites) = flexChecked Then
            If .Cell(flexcpChecked, .Row, eGDCol_Favorites) = flexUnchecked Then bFix = True
        ElseIf .Cell(flexcpChecked, .Row - 1, eGDCol_Favorites) = flexUnchecked And .Cell(flexcpChecked, .Row + 1, eGDCol_Favorites) = flexUnchecked Then   '6576
            If .Cell(flexcpChecked, .Row, eGDCol_Favorites) = flexChecked Then bFix = True
        End If
        
        If bFix Then
            aTemp.SplitFields .RowData(.Row), vbTab
            If .Cell(flexcpChecked, .Row, eGDCol_Favorites) = flexChecked Then
                .Cell(flexcpChecked, .Row, eGDCol_Favorites) = flexUnchecked
                aTemp(4) = "N"
            Else
                .Cell(flexcpChecked, .Row, eGDCol_Favorites) = flexChecked
                aTemp(4) = "F"
            End If
            
            .Cell(flexcpPictureAlignment, .Row, eGDCol_Favorites) = flexAlignCenterCenter
            .RowData(.Row) = aTemp.JoinFields(vbTab)
        End If
        m.bDirty = True             '5695
    End With
    
End Sub

Private Sub GetPageCollectionList()
On Error GoTo ErrSection:

    Dim i&
    Dim iCboIndex&
    Dim strFileMask$
    Dim aFiles As New cGdArray
    
    If Not m.bShowCPC Then Exit Sub

    iCboIndex = -1
    
    cboCollection.Clear
    cboCollection.AddItem kReservedName

    strFileMask = App.Path & "\charts\collection\*.CPC"
    aFiles.GetMatchingFiles strFileMask, False, False, True
    aFiles.Sort eGdSort_IgnoreCase

    For i = 0 To aFiles.Size - 1
        strFileMask = Parse(aFiles(i), ".", 1)
        cboCollection.AddItem strFileMask
        If InStr(g.ChartGlobals.strCPCRoot, strFileMask) <> 0 Then iCboIndex = i + 1
    Next
    
    If iCboIndex >= 0 And iCboIndex < cboCollection.ListCount Then
        cboCollection.ListIndex = iCboIndex
    Else
        cboCollection.ListIndex = 0
    End If
    
    If cboCollection.Text = kReservedName Then
        cmdCollectionDelete.Enabled = False
        
        cmdCopy.Visible = False
        cmdCopy.Enabled = False
    Else
        cmdCollectionDelete.Enabled = True
        
        cmdCopy.Visible = m.bShowCPC
        cmdCopy.Enabled = m.bShowCPC
    End If
    
    MoveFocus cmdCollectionExport

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTemplates.GetPageCollectionList"

End Sub

Private Sub PageCollectionChange(ByVal strCollection$, Optional ByVal bSaveCurrent As Boolean = True)
On Error GoTo ErrSection

    Dim i&
    
    If Not ExtractCollection(strCollection, False) Then Exit Sub
    
    Me.Hide
    DoEvents
    
    If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        MoveFocus g.ChartGlobals.frmActiveNonDetached.pbChart   'do this so maximized chart gets saved correctly
    End If
    
    'save current charts (if collection just got deleted then directory will not exist)
    If bSaveCurrent And DirExist(g.ChartGlobals.strCPCRoot) Then
        SaveCharts
    End If
    
    'change to specified charts collection directory
    If strCollection = kReservedName Then
        cmdCollectionDelete.Enabled = False
        cmdCopy.Visible = False
        cmdCopy.Enabled = False
        
        g.ChartGlobals.strCPCRoot = g.strAppPath
    Else
        cmdCollectionDelete.Enabled = True
        cmdCopy.Visible = m.bShowCPC
        cmdCopy.Enabled = m.bShowCPC
        
        g.ChartGlobals.strCPCRoot = g.strAppPath & "\charts\shared pages\" & strCollection
    End If
    
    InfBox "Loading the """ & strCollection & """ collection ...", "t", , "Chart Page Collection", True
    
    PageCollectionLoad
    
'JM 06-13-2014: original code left over; leave awhile then remove if all ok
'   original intention was to leave form up after switching to new collection to
'   allow user to choose a page; too many issues with this so not doing it anymore
'
'    FillList
'    m.strPath = g.ChartGlobals.strCPCRoot & "\Charts\Pages\"
'    Me.Show
            
    SetMainCaption
    If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        'so chart tabs will show correctly after a page collection change
        FormResize g.ChartGlobals.frmActiveNonDetached
    End If
    
    tmr.Tag = ""
    tmr.Enabled = True
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTemplates.PageCollectionChange"

End Sub

Private Function ExtractCollection(ByVal strCollection$, ByVal bClearExisting As Boolean) As Boolean
On Error GoTo ErrSection

    Dim strTempDir$, strSharedRoot$, strErr$, strSource$
    ''Dim fsoSysObj As New FileSystemObject
    
    If strCollection = kReservedName Then
        ExtractCollection = True
        Exit Function
    End If
    
    strSource = g.strAppPath & "\charts\Collection\" & strCollection & ".cpc"
    ''If Not fsoSysObj.FileExists(strSource) Then
    If Not FileExist(strSource) Then
        InfBox "The collection " & strSource & " was not found.", "E", , kCollectionInfbox
        Exit Function
    End If
    
    'make sure local shared pages area exists
    strSharedRoot = g.strAppPath & "\charts\Shared Pages"
    ''If Not fsoSysObj.FolderExists(strSharedRoot) Then
    If Not DirExist(strSharedRoot) Then
        strErr = "fsoSysObj.CreateFolder" & vbCrLf & strSharedRoot
        ''fsoSysObj.CreateFolder strSharedRoot
        MakeDir strSharedRoot
    End If
    strErr = ""
    
    'clear out shared page area for this CPC if exists
    strTempDir = strSharedRoot & "\" & strCollection
    If bClearExisting And DirExist(strTempDir) Then ''fsoSysObj.FolderExists(strTempDir) Then
        strErr = "fsoSysObj.DeleteFolder" & vbCrLf & strTempDir
        ''fsoSysObj.DeleteFolder strTempDir, True
        DeleteFolder AddSlash(strTempDir), True
    End If
    strErr = ""
    
    'create shared page area for this CPC
    strErr = "fsoSysObj.CreateFolder" & vbCrLf & strTempDir
    ''If Not fsoSysObj.FolderExists(strTempDir) Then
    If Not DirExist(strTempDir) Then
        ''fsoSysObj.CreateFolder strTempDir
        MakeDir strTempDir
    End If
    
    strErr = "fsoSysObj.CreateFolder" & vbCrLf & strTempDir & "\charts"
    ''If Not fsoSysObj.FolderExists(strTempDir & "\charts") Then
    If Not DirExist(strTempDir & "\charts") Then
        ''fsoSysObj.CreateFolder strTempDir & "\charts"
        MakeDir strTempDir & "\charts"
        strErr = "ZipExecute files:" & vbCrLf & strTempDir & "\charts"
        ZipExecute "U", strSource, strTempDir & "\charts", "*.cht,*.ini,*.ano,charts.cfg"
    End If
    
    strErr = "fsoSysObj.CreateFolder" & vbCrLf & strTempDir & "\charts\pages"
    ''If Not fsoSysObj.FolderExists(strTempDir & "\charts\pages") Then
    If Not DirExist(strTempDir & "\charts\pages") Then
        ''fsoSysObj.CreateFolder strTempDir & "\charts\pages"
        MakeDir strTempDir & "\charts\pages"
        strErr = "ZipExecute files:" & vbCrLf & strTempDir & "\charts\pages"
        ZipExecute "U", strSource, strTempDir & "\charts\pages", "*.gzp"
    End If
    strErr = ""
    
    
ErrExit:
    ExtractCollection = True
    Exit Function

ErrSection:
    RaiseError "frmTemplates.ExtractCollection"
    
    If Len(strErr) > 0 Then InfBox strErr, "E", , kCollectionInfbox

End Function

Private Sub tmr_Timer()
On Error Resume Next

    tmr.Enabled = False
    
    If tmr.Tag = "ZORDER" Then
        Me.ZOrder
        tmr.Tag = ""
    Else
        Unload Me
    End If

End Sub

Private Function CopyChart(ByVal bMoveChart As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim i&, j&, k&
    Dim iSrcChtNum&, iDestChtNum&
    
    Dim strPagePath$, strPage$
    Dim strDestCfg$, strDestChtFile$, strDestAnnot$
    Dim strSrcTemplate$, strSrcAnnot$, strSrcLine$
    Dim strTemp$, strTempDir$, strName$
    Dim strAnnotBaseFileName$
    
    Dim aFile As New cGdArray
    Dim aPages As New cGdArray
    
    Dim frm As frmChart
    
    Dim bRemoveDirTemp As Boolean

    If m.Chart Is Nothing Then Exit Function
    If m.Chart.Form Is Nothing Then Exit Function
    
    If fgList.SelectedRows = 0 Then
        strPage = g.strChartPage        '6921
    Else
        '6981 - need copy-to page name up front
        For i = 0 To fgList.SelectedRows - 1
            strPage = fgList.TextMatrix(fgList.SelectedRow(i), eGDCol_Name)
            If Len(strPage) > 0 Then Exit For
        Next
    End If
    
    SaveCharts      '6981 - need to save any new charts to page so .cfg file gets updated
    
    strPagePath = g.ChartGlobals.strCPCRoot & "\Charts\Pages\"
    
    'get CHT file for this chart
    strSrcTemplate = g.ChartGlobals.strCPCRoot & "\charts\" & m.Chart.Template & ".cht"
    If Not FileExist(strSrcTemplate) Then Exit Function
    iSrcChtNum = Int(Replace(UCase(m.Chart.Template), "CUS", ""))
    
    'get ANO file for this chart
    strAnnotBaseFileName = m.Chart.Template & "^" & m.Chart.AnnotSymbol     '6917
    strSrcAnnot = g.ChartGlobals.strCPCRoot & "\Charts\" & strAnnotBaseFileName & ".ANO"
    
    If strPage = g.strChartPage And FileExist(strSrcTemplate) Then
        'get name for new chart template file & do file copy
        strName = GetUnusedChartName
        strDestChtFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strName & ".cht"
        FileCopy strSrcTemplate, strDestChtFile
        
        If FileExist(strDestChtFile) Then
            If FileExist(strSrcAnnot) Then
                strDestAnnot = g.ChartGlobals.strCPCRoot & "\charts\" & strName & "^" & m.Chart.AnnotSymbol & ".ANO"
                FileCopy strSrcAnnot, strDestAnnot
            End If
        Else
            InfBox "Attempt to copy " & strSrcTemplate & " to " & strDestChtFile & " failed."
            
            Me.Hide
            tmr.Tag = ""
            tmr.Enabled = True
            
            GoTo ErrExit
        End If
        
        Set frm = New frmChart
        If Not frm.Chart.TemplateLoad(strName) Then
            InfBox "Attempt to load template  " & strDestChtFile & " failed."
            
            Me.Hide
            tmr.Tag = ""
            tmr.Enabled = True
            
            GoTo ErrExit
        End If
    
    End If
    
    strTemp = g.ChartGlobals.strCPCRoot & "\charts\\charts.cfg"
    If Not FileExist(strTemp) Then GoTo ErrExit
    'extract line that contains source chart template from source Cfg file
    aFile.FromFile strTemp
    For i = 0 To aFile.Size - 1
        strSrcLine = aFile(i)
        If InStr(strSrcLine, m.Chart.Template) <> 0 Then Exit For
    Next
    If Len(strSrcLine) = 0 Then
        InfBox "Attempt to process Cfg file failed."
        
        Me.Hide
        tmr.Tag = ""
        tmr.Enabled = True
        
        GoTo ErrExit
    End If
    
    If strPage = g.strChartPage Then
        strDestCfg = strTemp
        'add a line for new chart to current Chart.Cfg file
        strTemp = Replace(strSrcLine, m.Chart.Template, strName)
        aFile.Add strTemp
        aFile.ToFile strDestCfg

        frm.Chart.SetSymbol m.Chart.Symbol
        aFile.SplitFields strTemp, vbTab
            
        'shift left, top a little so new chart will not be directly on top of old one
        aPages.SplitFields aFile(2), ";"
        i = Val(aPages(0)) + 180
        aPages(0) = Str(i)
        i = Val(aPages(1)) + 180
        aPages(1) = Str(i)
        aFile(2) = aPages.JoinFields(";")
        
        frm.WindowLink.SymbolColor = Val(aFile(5))
        frm.WindowLink.PeriodColor = Val(aFile(6))
        frm.Chart.ResetLastScreenDate
            
        If aFile.Size > 9 Then
            frm.CopyPlacements , aFile(2), aFile(7), aFile(9)
        Else
            frm.CopyPlacements , aFile(2), aFile(7)
        End If
        frm.SetRatioPlacement aFile(7)
    
        Me.Hide
        
        frm.Show
        
        If Not m.Chart.AutoScale And frm.Chart.AutoScale Then
            Dim srcPane As cPane
            Dim destPane As cPane
            
            frm.Chart.AutoScale = m.Chart.AutoScale
            
            'walk through each pane & copy individual pane's scale properties
            If m.Chart.Tree.Count = frm.Chart.Tree.Count Then
                For i = 1 To m.Chart.Tree.Count
                    If TypeOf m.Chart.Tree(i) Is cPane Then
                        Set srcPane = m.Chart.Tree(i)
                        If TypeOf frm.Chart.Tree(i) Is cPane Then
                            Set destPane = frm.Chart.Tree(i)
                            destPane.Scaling = srcPane.Scaling
                            destPane.Max = srcPane.Max
                            destPane.Min = srcPane.Min
                            destPane.geCopyScaleStructInfo srcPane
                        End If
                    End If
                Next
            End If
                            
            
            Set srcPane = Nothing
            Set destPane = Nothing
            
            frm.Chart.GenerateChart eRedo1_Scrolled
        End If

        Me.Hide

        tmr.Tag = ""
        tmr.Enabled = True
        
        GoTo ErrExit
        
    End If
        
    'create temp directory & set flag to remove it if does not already exist
    If Not DirExist(g.ChartGlobals.strCPCRoot & "\charts\temp") Then
        If Not MakeDir(g.ChartGlobals.strCPCRoot & "\charts\temp") Then GoTo ErrExit
        bRemoveDirTemp = True
    End If
    
    'walk through grid & copy this chart to selected pages
    With fgList
        
        For i = 0 To .SelectedRows - 1
            strPage = .TextMatrix(.SelectedRow(i), eGDCol_Name)
            
            If FileExist(strPagePath & strPage & ".GZP") Then
                strTempDir = g.ChartGlobals.strCPCRoot & "\charts\temp\" & strPage
                
                If DirExist(strTempDir) Then
                    'should not exist, give message
                ElseIf MakeDir(strTempDir) Then
                    'extract zipped page to temp directory
                    ZipExecute "U", strPagePath & strPage & ".GZP", strTempDir
                    
                    'read content of destination charts cfg file to make sure CHT source name does not already exist
                    strDestCfg = g.ChartGlobals.strCPCRoot & "\charts\temp\" & strPage & "\charts.cfg"
                    If FileExist(strDestCfg) Then
                        aFile.FromFile strDestCfg
                        
                        iDestChtNum = iSrcChtNum
                        For j = 1 To aFile.Size - 1
                            strTemp = UCase(Parse(aFile(j), vbTab, 2))       'grab the CUSxxxxx from line of the Cfg file
                            If InStr(strTemp, "CUS") <> 0 Then
                                k = Int(Replace(strTemp, "CUS", ""))            'extract the numeric portion of the CUSxxxxx
                                If k >= iDestChtNum Then iDestChtNum = k + 1    'increment the digit portion of CHT file name if necessary
                            End If
                        Next
                        
                        'modify destination Cfg file to include this chart
                        If iDestChtNum = iSrcChtNum Then
                            'the CUSxxxxx.CHT file for this chart does not exist in destination page
                            'just add line from source Cfg to destination Cfg file
                            aFile.Add strSrcLine
                            strDestChtFile = m.Chart.Template
                        Else
                            'the CUSxxxxx.CHT file for this chart exists in destination page
                            'tweak the CUSxxxxx.CHT name before adding it to destination Cfg file
                            strTemp = Parse(strSrcLine, vbTab, 2)
                            strDestChtFile = "Cus" & Format(iDestChtNum, "0000#")
                            aFile.Add Replace(strSrcLine, strTemp, strDestChtFile)
                        End If
                        
                        aFile.ToFile strDestCfg
                        
                        If FileExist(strTempDir & "\" & strDestChtFile & ".CHT") Then
                            KillFile strTempDir & "\" & strDestChtFile & ".CHT", True
                        End If
                        
                        FileCopy strSrcTemplate, strTempDir & "\" & strDestChtFile & ".CHT"
                        
                        If FileExist(strSrcAnnot) Then
                            strDestAnnot = strTempDir & "\" & strDestChtFile & "^" & m.Chart.AnnotSymbol & ".ANO"
                            If FileExist(strDestAnnot) Then
                                KillFile strDestAnnot, True
                            End If
                        
                            FileCopy strSrcAnnot, strDestAnnot
                        End If
                        ZipExecute "C", strPagePath & strPage & ".GZP", g.ChartGlobals.strCPCRoot & "\charts\temp\" & strPage, "*.*"
                    
                        KillFolder strTempDir, True
                        
                        'remove the page from cache so will load from file
                        CachePageRemove strPage
                        
                        aPages.Add strPage
                    End If
                
                End If
            End If
        Next
    
    End With
    
    If m.Chart.Form.WindowState = vbMaximized Then
        strTemp = Parse(m.Chart.Form.vseCaption.Caption, ")", 1)
    Else
        strTemp = Parse(m.Chart.Form.Caption, ")", 1)
    End If
    If InStr(strTemp, "(") <> 0 Then strTemp = strTemp & ")"
    
    If aPages.Size > 0 Then
        If bMoveChart Then
            strTemp = "The chart " & strTemp & " successfully moved to: " & vbCrLf
        Else
            strTemp = "The chart " & strTemp & " successfully copied to: " & vbCrLf
        End If
        
        For i = 0 To aPages.Size - 1
            strTemp = strTemp & vbCrLf & aPages(i)
        Next
        
        CopyChart = True
        
    ElseIf bMoveChart Then
        strTemp = "Attempt to move the chart " & strTemp & " failed"
    Else
        strTemp = "Attempt to copy the chart " & strTemp & " failed"
    End If
    
    InfBox strTemp

ErrExit:
    If bRemoveDirTemp Then KillFolder g.ChartGlobals.strCPCRoot & "\charts\temp", True
    Exit Function
    
ErrSection:
    RaiseError "frmTemplates.CopyChart"

End Function

Private Sub CopyChartNewPage()
On Error GoTo ErrSection:

    Dim bCfgReadSuccess As Boolean
    Dim dResult#, i&
    
    Dim strPage$, strPath$, strErr$
    Dim strSrcTemplate$, strSrcAnnot$
    
    Dim aFields As New cGdArray
    Dim aTempCfg As New cGdArray
    Dim aCfgSave As New cGdArray
    
    strPath = g.ChartGlobals.strCPCRoot & "\Charts\"
    
    If Not FileExist(strPath & "Charts.cfg") Then       '6880
        SaveChartPage ""
        GoTo ErrExit
    End If
    
    ' get name of page
    Do
        ' ask for name until valid or cancelled
        strPage = Trim(InfBox("Name for new chart page:", "?", _
            "+Save|-Cancel", "Save Chart Page", , , , , , "s", strPage))
        If Len(strPage) = 0 Then Exit Sub
        If IsValidFileBase(strPage) Then Exit Do
    Loop
    
    'read chart.cfg file for current chart page
    aCfgSave.FromFile strPath & "Charts.cfg"
    If aCfgSave.Size <= 0 Then
        strErr = "Attemp to read charts.cfg file failed."
        GoTo ErrExit
    End If
    bCfgReadSuccess = True
    
    'get CHT file for this chart
    strSrcTemplate = g.ChartGlobals.strCPCRoot & "\Charts\" & m.Chart.Template & ".cht"
    If Not FileExist(strSrcTemplate) Then
        strErr = "Attempt to locate " & strSrcTemplate & " failed"
        GoTo ErrExit
    End If
    
    'write out temporary chart.cfg for just this chart
    aTempCfg.Add aCfgSave(0)
    aTempCfg.Add aCfgSave(1)
    aFields.SplitFields aTempCfg(1), vbTab
    If aFields.Size > 2 Then
        aFields(0) = m.Chart.SymbolID
        aFields(1) = m.Chart.Template
        
        aTempCfg(1) = aFields.JoinFields(vbTab)
        
        aTempCfg.ToFile strPath & "\Charts.Cfg"
    End If
    
    strSrcAnnot = g.ChartGlobals.strCPCRoot & "\Charts\" & m.Chart.Template & "^" & m.Chart.SymbolID & ".ANO"
    ' zip up the files needed for just this chart page
    If FileExist(strSrcAnnot) Then
        ' (but not global annotation files, which start with a caret)
        dResult = ZipExecute("C", strPath & "Pages\" & strPage & ".GZP", strPath, "* /Chart.cfg," & strSrcTemplate & "," & strSrcAnnot)
    Else
        dResult = ZipExecute("C", strPath & "Pages\" & strPage & ".GZP", strPath, "* /Chart.cfg," & strSrcTemplate)
    End If
    
    If dResult > 0 Then
        Me.Hide
        If InfBox("Would you like to load the new chart page?", "?", "+Yes|-No", "Chart copy/move") = "Y" Then
            LoadChartPage strPage
        End If
    Else
        strErr = "Unknown error: zipexecute returned " & Str(dResult) & ". Copy to new page failed"
    End If

ErrExit:
    'restore original cfg file
    If bCfgReadSuccess Then aCfgSave.ToFile strPath & "Charts.Cfg"
    
    If Len(strErr) > 0 Then InfBox strErr, "E", , "Chart copy/move"

    Exit Sub

ErrSection:
    'restore original cfg file
    If bCfgReadSuccess Then aCfgSave.ToFile strPath & "Charts.Cfg"
    
    RaiseError Me.Name & ".CopyChartNewPage"
    
End Sub

