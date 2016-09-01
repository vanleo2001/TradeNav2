VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmMarkets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Market Information"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "frmMarkets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HexUniControls.ctlUniFrameWL fraUsePrev 
      Height          =   1395
      Left            =   180
      TabIndex        =   11
      Top             =   2820
      Width           =   4515
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmMarkets.frx":030A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMarkets.frx":033E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMarkets.frx":035E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdClear 
         Height          =   435
         Left            =   0
         TabIndex        =   7
         Top             =   960
         Width           =   1335
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
         Caption         =   "frmMarkets.frx":037A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMarkets.frx":03AE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":03CE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdMap 
         Height          =   435
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   1335
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
         Caption         =   "frmMarkets.frx":03EA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMarkets.frx":0428
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":0448
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdUsePrev 
         Height          =   435
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1335
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
         Caption         =   "frmMarkets.frx":0464
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMarkets.frx":049E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":04BE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblClear 
         Height          =   435
         Left            =   1500
         Top             =   960
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
         Caption         =   "frmMarkets.frx":04DA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMarkets.frx":0592
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":05B2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblMap 
         Height          =   435
         Left            =   1500
         Top             =   480
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
         Caption         =   "frmMarkets.frx":05CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMarkets.frx":067C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":069C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUsePrev 
         Height          =   435
         Left            =   1500
         Top             =   0
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
         Caption         =   "frmMarkets.frx":06B8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMarkets.frx":077A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":079A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgInfo 
      Height          =   1275
      Left            =   180
      TabIndex        =   10
      Top             =   1380
      Width           =   4515
      _cx             =   7964
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
   Begin HexUniControls.ctlUniFrameWL fraSymbolInfo 
      Height          =   1035
      Left            =   180
      TabIndex        =   3
      Top             =   180
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
      Caption         =   "frmMarkets.frx":07B6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMarkets.frx":07E2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMarkets.frx":0802
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   0
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   12632256
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   -1  'True
         Text            =   "frmMarkets.frx":081E
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
         Tip             =   "frmMarkets.frx":0850
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":0870
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDescription 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   12632256
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   -1  'True
         Text            =   "frmMarkets.frx":088C
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
         Tip             =   "frmMarkets.frx":08C8
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":08E8
      End
      Begin HexUniControls.ctlUniComboImageXP cboSecurityType 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   690
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
         Tip             =   "frmMarkets.frx":0904
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":0924
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   0
         Top             =   0
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
         Caption         =   "frmMarkets.frx":0940
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMarkets.frx":096E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":098E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   0
         Top             =   360
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
         Caption         =   "frmMarkets.frx":09AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMarkets.frx":09E4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":0A04
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label7 
         Height          =   255
         Left            =   0
         Top             =   720
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
         Caption         =   "frmMarkets.frx":0A20
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMarkets.frx":0A5E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":0A7E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   795
      Left            =   3660
      TabIndex        =   2
      Top             =   180
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
      Caption         =   "frmMarkets.frx":0A9A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMarkets.frx":0AC6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMarkets.frx":0AE6
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   0
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
         Caption         =   "frmMarkets.frx":0B02
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMarkets.frx":0B28
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":0B48
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   420
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
         Caption         =   "frmMarkets.frx":0B64
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMarkets.frx":0B92
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMarkets.frx":0BB2
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmMarkets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmMarkets.frm
'' Description: Allow the user to edit Market information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Type = 0
    eGDCol_Custom
    eGDCol_Value
    eGDCol_Default
    eGDCol_NumCols
End Enum

Private Enum eGDRows
    eGDRow_TickMove = 1
    eGDRow_TickValue
    eGDRow_MinMove
    eGDRow_Margin
    eGDRow_NumRows
End Enum

Private Type mPrivate
    strSymbol As String
    strDescription As String
    lSymbolID As Long
    Bars As cGdBars
    DefaultBars As cGdBars
    strPrevSettings As String
    bHasCustomSettings As Boolean
        
    bOK As Boolean
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Function GDRow(ByVal Row As eGDRows) As Long
    GDRow = Row
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Symbol, Security Type, and Description of symbol to edit
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal strSymbol As String, Optional ByVal strDescription As String = "", _
    Optional ByRef Chart As cChart = Nothing) As Boolean
On Error GoTo ErrSection:
    
    Dim astrSymbols As New cGdArray
    Dim tmpBars As New cGdBars
    Dim rs As Recordset
    Dim astrValues As New cGdArray
    
    m.strPrevSettings = GetIniFileProperty("Previous", "", "MarketInfo", g.strIniFile)
    m.strDescription = strDescription
    
    If Left(strSymbol, 1) = "*" Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarketInfo] " & _
                "WHERE [tblMarketInfo.Symbol] LIKE '[*]*';", dbOpenDynaset)
        If (rs.BOF And rs.EOF) Then
            m.bHasCustomSettings = False
        Else
            m.bHasCustomSettings = True
        End If
    Else
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarketInfo] " & _
                "WHERE [tblMarketInfo.Symbol] LIKE '[!*]*';", dbOpenDynaset)
        If (rs.BOF And rs.EOF) Then
            m.bHasCustomSettings = False
        Else
            m.bHasCustomSettings = True
        End If
    End If
    
    ' Get Market information for this symbol...
    Set m.Bars = New cGdBars
    m.strSymbol = strSymbol
    m.lSymbolID = GetMarketInfo(strSymbol, m.Bars)
    
    ' Get Default Market information for this symbol...
    Set m.DefaultBars = New cGdBars
    SetBarProperties m.DefaultBars, strSymbol, True
    
    ' If we don't recognize the symbol, try to get the user to give us a
    ' Genesis equivalent symbol that we can get the information from...
    'If m.lSymbolID = 0 And m.Bars.Prop(eBARS_TickMove) = 0 Then
    '    If InfBox("Trade Navigator does not recognize " & UCase(strSymbol) & ".  Would you like to " & _
    '            "search for an equivalent Genesis symbol?", "?", "+Yes|-No", strSymbol) = "Y" Then
    '        Set astrSymbols = frmSymbolSelector.ShowMe(, False, False, "Genesis Equivalent for " & UCase(strSymbol))
    '        If astrSymbols.Size > 0 Then
    '            Set tmpBars = New cGdBars
    '            If GetMarketInfo(astrSymbols(0), tmpBars) > 0 Then
    '                Set m.Bars = tmpBars.MakeCopy
    '            End If
    '        End If
    '    End If
    'End If
    
    ' Load the controls with the values...
    With m.Bars
        txtSymbol.Text = strSymbol
        
        txtDescription.Text = ""
        If Len(strDescription) > 0 Then txtDescription.Text = strDescription
        If Len(.Prop(eBARS_Desc)) > 0 Then txtDescription.Text = .Prop(eBARS_Desc)
        
        If .Prop(eBARS_TickMove) = 0 And .Prop(eBARS_TickValue) = 0 And Left(strSymbol, 1) = "*" Then
            SetCombo ""
        Else
            SetCombo .SecurityType
        End If
        
        fgInfo.Redraw = flexRDNone
        InitGrid
        LoadGrid
        fgInfo.Redraw = flexRDBuffered
    End With
        
    ' Allow the user to change the description if not one of our symbols...
    txtSymbol.TabStop = False
    If m.lSymbolID = 0 Then
        txtDescription.Locked = False
        txtDescription.BackColor = fgInfo.BackColor
        txtDescription.TabStop = True
        
        'RH commented out cboSecurityType.Locked = False
        cboSecurityType.Enabled = True
        cboSecurityType.BackColor = fgInfo.BackColor
        cboSecurityType.TabStop = True
        
        cmdUsePrev.Visible = True
        lblUsePrev.Visible = True
        cmdMap.Visible = True
        lblMap.Visible = True
        cmdClear.Visible = True
        cmdClear.Caption = "Clear &All"
        lblClear.Visible = True
        lblClear.Caption = "Clear ALL of the custom settings that you have entered for external symbols."
        fraUsePrev.Height = 1395
        Height = 4725
        
        MoveFocus txtDescription
    Else
        txtDescription.Locked = True
        txtDescription.BackColor = txtSymbol.BackColor
        txtDescription.TabStop = False
        
        'RH commented out cboSecurityType.Locked = True
        cboSecurityType.Enabled = False
        cboSecurityType.BackColor = txtSymbol.BackColor
        cboSecurityType.TabStop = False
        
        cmdUsePrev.Visible = False
        lblUsePrev.Visible = False
        cmdMap.Visible = False
        lblMap.Visible = False
        With cmdClear
            .Visible = True
            .Move .Left, cmdUsePrev.Top
            .Caption = "Restore &All"
        End With
        With lblClear
            .Visible = True
            .Move .Left, lblUsePrev.Top
            .Caption = "Restore ALL of the default settings for Genesis symbols."
        End With
        fraUsePrev.Height = 435
        Height = 3795
        
        fgInfo.Row = 1
        fgInfo.Col = 1
        MoveFocus fgInfo
    End If
        
    ' Show the form...
    If Chart Is Nothing Then
        CenterTheForm Me
    Else
        CenterFormOnChart Me, Chart                     '6499
    End If
    
    SetEditorCaption Me, "Market Information", strSymbol
    EnableControls
    ShowForm Me, True

    ' Save the Market information if the user hit OK...
    If m.bOK Then
        Save
        
        astrValues.Create eGDARRAY_Strings
        astrValues(0) = txtDescription.Text
        astrValues(1) = cboSecurityType.Text
        astrValues(2) = fgInfo.TextMatrix(1, GDCol(eGDCol_Value))
        astrValues(3) = fgInfo.TextMatrix(2, GDCol(eGDCol_Value))
        astrValues(4) = fgInfo.TextMatrix(3, GDCol(eGDCol_Value))
        astrValues(5) = fgInfo.TextMatrix(4, GDCol(eGDCol_Value))
        
        SetIniFileProperty "Previous", astrValues.JoinFields("|"), "MarketInfo", g.strIniFile
        
        If m.lSymbolID <> 0 Then
            UpdateVisibleCharts eRedo9_ReloadData, m.lSymbolID
        End If
    End If
    ShowMe = m.bOK
    
ErrExit:
    Set rs = Nothing
    Set astrSymbols = Nothing
    Set tmpBars = Nothing
    Set astrValues = Nothing
    Unload Me
    Exit Function
    
ErrSection:
    Set rs = Nothing
    Set astrSymbols = Nothing
    Set tmpBars = Nothing
    Set astrValues = Nothing
    Unload Me
    RaiseError "frmMarkets.ShowMe", eGDRaiseError_Raise

End Function

Private Sub cboSecurityType_GotFocus()
On Error Resume Next

    If cboSecurityType.TabStop = False Then
        fgInfo.Row = 1
        fgInfo.Col = 1
        MoveFocus fgInfo
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving information
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
    RaiseError "frmMarkets.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClear_Click
'' Description: Allow the user to clear out all of the custom overrides
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClear_Click()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop

    If Left(m.strSymbol, 1) = "*" Then
        If InfBox("This will clear out the settings for ALL of the external symbols that you have entered.  Do you want to continue?", "?", "+Yes|-No", "Confirmation") = "Y" Then
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarketInfo] " & _
                        "WHERE [tblMarketInfo.Symbol] LIKE '[*]*';", dbOpenDynaset)
            Do While Not rs.EOF
                rs.Delete
                rs.MoveNext
            Loop
            
            If Len(m.strDescription) > 0 Then
                txtDescription.Text = m.strDescription
            Else
                txtDescription.Text = m.strSymbol
            End If
            cboSecurityType.ListIndex = 0
            fgInfo.TextMatrix(GDRow(eGDRow_TickMove), GDCol(eGDCol_Value)) = "0.0"
            fgInfo.TextMatrix(GDRow(eGDRow_TickValue), GDCol(eGDCol_Value)) = "$0.00"
            fgInfo.TextMatrix(GDRow(eGDRow_MinMove), GDCol(eGDCol_Value)) = "1.0"
            fgInfo.TextMatrix(GDRow(eGDRow_Margin), GDCol(eGDCol_Value)) = "$0.00"
                        
            m.bHasCustomSettings = False
            EnableControls
        End If
    Else
        If InfBox("This will clear out the custom settings for ALL of the Genesis symbols that you have entered.  Do you want to continue?", "?", "+Yes|-No", "Confirmation") = "Y" Then
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarketInfo] " & _
                        "WHERE [tblMarketInfo.Symbol] LIKE '[!*]*';", dbOpenDynaset)
            Do While Not rs.EOF
                rs.Delete
                rs.MoveNext
            Loop
            
            For lIndex = fgInfo.FixedRows To fgInfo.Rows - 1
                fgInfo.TextMatrix(lIndex, GDCol(eGDCol_Value)) = fgInfo.TextMatrix(lIndex, GDCol(eGDCol_Default))
                CheckedCell(fgInfo, lIndex, GDCol(eGDCol_Custom)) = False
            Next lIndex
            
            m.bHasCustomSettings = False
            EnableControls
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.cmdClear.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdMap_Click
'' Description: Allow the user to map an external symbol to a Genesis symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdMap_Click()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Return from the symbol selector form
    Dim tmpBars As cGdBars              ' Temporary bars structure
    Dim frm As New frmSymbolSelector    ' New form for the symbol selector
    
    Set astrSymbols = frm.ShowMe("", False, False, "Genesis Equivalent for " & UCase(txtSymbol.Text))
    If astrSymbols.Size > 0 Then
        Set tmpBars = New cGdBars
        If GetMarketInfo(astrSymbols(0), tmpBars) > 0 Then
            Set m.Bars = tmpBars.MakeCopy
        
            ' Load the controls with the values...
            With m.Bars
                If Len(.Prop(eBARS_Desc)) > 0 Then txtDescription.Text = .Prop(eBARS_Desc)
                
                If .Prop(eBARS_TickMove) = 0 And .Prop(eBARS_TickValue) = 0 Then
                    SetCombo ""
                Else
                    SetCombo .SecurityType
                End If
                
                fgInfo.Redraw = flexRDNone
                InitGrid
                LoadGrid
                fgInfo.Redraw = flexRDBuffered
            End With
        End If
    End If

ErrExit:
    Set frm = Nothing
    Set astrSymbols = Nothing
    Set tmpBars = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.cmdMap.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOk_Click
'' Description: Unload the form and save the changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Dim dTemp As Double                 ' Temporary value
    
    MoveFocus cmdOK
    
    If Len(cboSecurityType.Text) = 0 Then
        Err.Raise vbObjectError + 1000, , "Must supply a security type"
    End If
    
    dTemp = ValOfText(fgInfo.TextMatrix(GDRow(eGDRow_TickMove), GDCol(eGDCol_Value)))
    If dTemp <= 0 Or dTemp > 1000000 Then
        Err.Raise vbObjectError + 1000, , _
            "Tick Move must be greater than 0 and less than 1,000,000"
    End If
        
    dTemp = ValOfText(fgInfo.TextMatrix(GDRow(eGDRow_TickValue), GDCol(eGDCol_Value)))
    If dTemp <= 0 Or dTemp > 1000000 Then
        Err.Raise vbObjectError + 1000, , _
            "Tick Value must be greater than 0 and less than 1,000,000"
    End If
    
    dTemp = ValOfText(fgInfo.TextMatrix(GDRow(eGDRow_MinMove), GDCol(eGDCol_Value)))
    If dTemp <= 0 Or dTemp > 1000000 Then
        Err.Raise vbObjectError + 1000, , _
            "Min Move must be greater than 0 and less than 1,000,000"
    End If
    
    dTemp = ValOfText(fgInfo.TextMatrix(GDRow(eGDRow_Margin), GDCol(eGDCol_Value)))
    If dTemp < 0 Or dTemp > 1000000 Then
        Err.Raise vbObjectError + 1000, , "Margin must be between 0 and 1,000,000"
    End If
    
    m.bOK = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdUsePrev_Click
'' Description: Use the previous values that the user used
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdUsePrev_Click()
On Error GoTo ErrSection:

    txtDescription.Text = Parse(m.strPrevSettings, "|", 1)
    cboSecurityType.Text = Parse(m.strPrevSettings, "|", 2)
    fgInfo.TextMatrix(1, GDCol(eGDCol_Value)) = Parse(m.strPrevSettings, "|", 3)
    fgInfo.TextMatrix(2, GDCol(eGDCol_Value)) = Parse(m.strPrevSettings, "|", 4)
    fgInfo.TextMatrix(3, GDCol(eGDCol_Value)) = Parse(m.strPrevSettings, "|", 5)
    fgInfo.TextMatrix(4, GDCol(eGDCol_Value)) = Parse(m.strPrevSettings, "|", 6)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.cmdUsePrev.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Select Case Col
        Case GDCol(eGDCol_Custom)
            If Not CheckedCell(fgInfo, Row, Col) Then
                fgInfo.TextMatrix(Row, GDCol(eGDCol_Value)) = fgInfo.TextMatrix(Row, GDCol(eGDCol_Default))
            End If
        
        Case GDCol(eGDCol_Value)
            Select Case Row
                Case GDRow(eGDRow_TickMove), GDRow(eGDRow_TickValue), GDRow(eGDRow_MinMove)
                    fgInfo.TextMatrix(Row, Col) = Format(ValOfText(fgInfo.TextMatrix(Row, Col)), "0.0#######")
                Case GDRow(eGDRow_Margin)
                    fgInfo.TextMatrix(Row, Col) = Format(ValOfText(fgInfo.TextMatrix(Row, Col)), "$#,##0.00")
            End Select
            
            CheckedCell(fgInfo, Row, GDCol(eGDCol_Custom)) = (ValOfText(fgInfo.TextMatrix(Row, GDCol(eGDCol_Value))) <> ValOfText(fgInfo.TextMatrix(Row, GDCol(eGDCol_Default))))
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.fgInfo.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    If NewCol = GDCol(eGDCol_Value) Then fgInfo.EditCell

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.fgInfo.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    Select Case Col
        Case GDCol(eGDCol_Type)
            Cancel = True
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.fgInfo.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgInfo_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    If NewCol = GDCol(eGDCol_Type) Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.fgInfo.BeforeRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgInfo
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        Select Case lMouseCol
            Case GDCol(eGDCol_Type), GDCol(eGDCol_Value)
                Select Case lMouseRow
                    Case GDRow(eGDRow_TickMove)
                        .ToolTipText = "Number of points this security moves in one tick"
                    Case GDRow(eGDRow_TickValue)
                        .ToolTipText = "Value of each Tick Move"
                    Case GDRow(eGDRow_MinMove)
                        .ToolTipText = "Minimum number of ticks this security can move"
                    Case GDRow(eGDRow_Margin)
                        .ToolTipText = "Margin required to trade this security"
                    Case Else
                        .ToolTipText = ""
                End Select
            
            Case GDCol(eGDCol_Custom)
                .ToolTipText = "Click here to override the default value"
            
            Case Else
                .ToolTipText = ""
        End Select
    End With
    
End Sub

Private Sub fgInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current grid row clicked on
    Dim lMouseCol As Long               ' Current grid column clicked on

    With fgInfo
        If Button = vbLeftButton Then
            lMouseRow = .MouseRow
            lMouseCol = .MouseCol
            
            If lMouseCol = GDCol(eGDCol_Custom) Then
                If CheckedCell(fgInfo, lMouseRow, lMouseCol) = True Then
                    .Row = lMouseRow
                    .Col = GDCol(eGDCol_Value)
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.fgInfo.MouseUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgInfo_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim dTemp As Double                 ' Temporary variable

    If Col = GDCol(eGDCol_Value) Then
        dTemp = ValOfText(fgInfo.EditText)
        Select Case Row
            Case GDRow(eGDRow_TickMove)
                If dTemp <= 0 Or dTemp > 1000000 Then
                    Err.Raise vbObjectError + 1000, , _
                        "Tick Move must be greater than 0 and less than 1,000,000"
                End If
            Case GDRow(eGDRow_TickValue)
                If dTemp <= 0 Or dTemp > 1000000 Then
                    Err.Raise vbObjectError + 1000, , _
                        "Tick Value must be greater than 0 and less than 1,000,000"
                End If
            Case GDRow(eGDRow_MinMove)
                If dTemp <= 0 Or dTemp > 1000000 Then
                    Err.Raise vbObjectError + 1000, , _
                        "Min Move must be greater than 0 and less than 1,000,000"
                End If
            Case GDRow(eGDRow_Margin)
                If dTemp < 0 Or dTemp > 1000000 Then
                    Err.Raise vbObjectError + 1000, , _
                        "Margin must be between 0 and 1,000,000"
                End If
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    Cancel = True
    RaiseError "frmMarkets.fgInfo.ValidateEdit", eGDRaiseError_Show
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
    RaiseError "frmMarkets.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    g.Styler.StyleForm Me

    With cboSecurityType
        .AddItem "", 0
        .AddItem "Future", 1
        .AddItem "Stock", 2
        .AddItem "Index / Forex", 3
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Unload the form without saving changes (if X clicked)
'' Inputs:      Whether or not to Cancel Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the information that the user entered
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    ' First delete all information about this symbol out of the database...
    g.dbNav.Execute "DELETE * FROM [tblMarketInfo] WHERE [Symbol]='" & m.strSymbol & "';"
    
    If ((Trim(txtDescription.Text) <> m.DefaultBars.Prop(eBARS_Desc)) And (Len(Trim(txtDescription.Text)) > 0) Or (m.lSymbolID = 0&)) Then
        UpdateMarketInfo m.strSymbol, m.lSymbolID, eMarketType_Desc, Trim(txtDescription.Text)
    End If
        
    If (Left(cboSecurityType.Text, 1) <> m.DefaultBars.SecurityType) Or (m.lSymbolID = 0&) Then
        UpdateMarketInfo m.strSymbol, m.lSymbolID, eMarketType_SecurityType, Left(cboSecurityType.Text, 1)
    End If
        
    For lIndex = fgInfo.FixedRows To GDRow(eGDRow_NumRows) - 1
        If CheckedCell(fgInfo, lIndex, GDCol(eGDCol_Custom)) Or (m.lSymbolID = 0&) Then
            UpdateMarketInfo m.strSymbol, m.lSymbolID, lIndex + 1, Str(ValOfText(fgInfo.TextMatrix(lIndex, GDCol(eGDCol_Value))))
        End If
    Next lIndex
    
    g.bDirtyLibrariesMDB = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarkets.Save", eGDRaiseError_Raise

End Sub

Private Sub txtDescription_GotFocus()
On Error Resume Next

    If txtDescription.TabStop = False Then
        fgInfo.Row = 1
        fgInfo.Col = 1
        MoveFocus fgInfo
    Else
        SelectAll txtDescription
    End If

End Sub

Private Sub txtSymbol_GotFocus()
On Error Resume Next

    fgInfo.Row = 1
    fgInfo.Col = 1
    MoveFocus fgInfo

End Sub

Private Sub SetCombo(ByVal strSecurityType As String)
On Error GoTo ErrSection:

    With cboSecurityType
        Select Case strSecurityType
            Case "F"
                .Text = "Future"
            Case "S"
                .Text = "Stock"
            Case "I"
                .Text = "Index / Forex"
            Case Else
                .ListIndex = 0
        End Select
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.SetCombo", eGDRaiseError_Raise
    
End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    
    With fgInfo
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowSelection = False
        .ExplorerBar = flexExNone
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        
        .Rows = GDRow(eGDRow_NumRows)
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .ColDataType(GDCol(eGDCol_Custom)) = flexDTBoolean
        .ColHidden(GDCol(eGDCol_Default)) = True
        .ColHidden(GDCol(eGDCol_Custom)) = (m.lSymbolID = 0)
        
        .TextMatrix(0, GDCol(eGDCol_Type)) = "Type"
        .TextMatrix(0, GDCol(eGDCol_Custom)) = "Custom"
        .TextMatrix(0, GDCol(eGDCol_Value)) = "Value"
        
        .TextMatrix(GDRow(eGDRow_TickMove), GDCol(eGDCol_Type)) = "Tick Move (points)"
        .TextMatrix(GDRow(eGDRow_TickValue), GDCol(eGDCol_Type)) = "Tick Value (dollars)"
        .TextMatrix(GDRow(eGDRow_MinMove), GDCol(eGDCol_Type)) = "Minimum Move in Ticks"
        .TextMatrix(GDRow(eGDRow_Margin), GDCol(eGDCol_Type)) = "Margin Required"
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.InitGrid", eGDRaiseError_Raise
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long

    With fgInfo
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .TextMatrix(GDRow(eGDRow_TickMove), GDCol(eGDCol_Value)) = Format(m.Bars.Prop(eBARS_TickMove), "0.0#######")
        .TextMatrix(GDRow(eGDRow_TickValue), GDCol(eGDCol_Value)) = Format(m.Bars.Prop(eBARS_TickValue), "$#,##0.00###")
        If m.Bars.Prop(eBARS_TickMove) = 0 And m.Bars.Prop(eBARS_MinMoveInTicks) = 0 Then
            .TextMatrix(GDRow(eGDRow_MinMove), GDCol(eGDCol_Value)) = "1.0"
        Else
            .TextMatrix(GDRow(eGDRow_MinMove), GDCol(eGDCol_Value)) = Format(m.Bars.Prop(eBARS_MinMoveInTicks), "0.0#######")
        End If
        .TextMatrix(GDRow(eGDRow_Margin), GDCol(eGDCol_Value)) = Format(m.Bars.Prop(eBARS_Margin), "$#,##0.00")
        
        .TextMatrix(GDRow(eGDRow_TickMove), GDCol(eGDCol_Default)) = Format(m.DefaultBars.Prop(eBARS_TickMove), "0.0#######")
        .TextMatrix(GDRow(eGDRow_TickValue), GDCol(eGDCol_Default)) = Format(m.DefaultBars.Prop(eBARS_TickValue), "$#,##0.00###")
        .TextMatrix(GDRow(eGDRow_MinMove), GDCol(eGDCol_Default)) = Format(m.DefaultBars.Prop(eBARS_MinMoveInTicks), "0.0#######")
        .TextMatrix(GDRow(eGDRow_Margin), GDCol(eGDCol_Default)) = Format(m.DefaultBars.Prop(eBARS_Margin), "$#,##0.00")
        
        CheckedCell(fgInfo, GDRow(eGDRow_TickMove), GDCol(eGDCol_Custom)) = Override(eMarketType_TickMove)
        CheckedCell(fgInfo, GDRow(eGDRow_TickValue), GDCol(eGDCol_Custom)) = Override(eMarketType_TickValue)
        CheckedCell(fgInfo, GDRow(eGDRow_MinMove), GDCol(eGDCol_Custom)) = Override(eMarketType_MinMoveInTicks)
        CheckedCell(fgInfo, GDRow(eGDRow_Margin), GDCol(eGDCol_Custom)) = Override(eMarketType_Margin)
        
        .AutoSize 0, .Cols - 1
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.LoadGrid", eGDRaiseError_Raise
    
End Sub

Private Function Override(ByVal MarketType As eMarketTypes) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    If m.lSymbolID = 0 Then
        Override = True
    Else
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarketInfo] " & _
                "WHERE [Symbol]='" & m.strSymbol & "' AND [SymbolID]=" & m.lSymbolID & _
                " AND [DataType]=" & MarketType & ";", dbOpenSnapshot)
        
        Override = Not (rs.BOF And rs.EOF)
    End If

ErrExit:
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmMarkets.Override", eGDRaiseError_Raise
    
End Function

Private Sub UpdateMarketInfo(ByVal strSymbol As String, ByVal lSymbolID As Long, _
                    ByVal nDataType As eMarketTypes, ByVal strValue As String)
On Error GoTo ErrSection:

    Dim rs As Recordset

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblMarketInfo];", dbOpenDynaset)
    
    rs.AddNew
    rs!Symbol = strSymbol
    rs!SymbolID = lSymbolID
    rs!DataType = nDataType
    rs!Value = strValue
    rs.Update
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmMarkets.UpdateMarketInfo", eGDRaiseError_Raise
                    
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

    Enable cmdUsePrev, (Len(m.strPrevSettings) > 0)
    Enable lblUsePrev, (Len(m.strPrevSettings) > 0)
    Enable cmdClear, m.bHasCustomSettings
    Enable lblClear, m.bHasCustomSettings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarkets.EnableControls", eGDRaiseError_Raise
    
End Sub

