VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmArrange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arrange Chart Windows"
   ClientHeight    =   5445
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4440
   Icon            =   "frmArrange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL Frame2 
      Height          =   735
      Left            =   2160
      TabIndex        =   9
      Top             =   4620
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
      Caption         =   "frmArrange.frx":014A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmArrange.frx":018E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmArrange.frx":01AE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPeriodicity 
         Height          =   375
         Left            =   1020
         TabIndex        =   5
         Top             =   240
         Width           =   975
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
         Caption         =   "frmArrange.frx":01CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmArrange.frx":01FE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmArrange.frx":021E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAlpha 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   795
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
         Caption         =   "frmArrange.frx":023A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmArrange.frx":0262
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmArrange.frx":0282
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   735
      Left            =   180
      TabIndex        =   8
      Top             =   4620
      Width           =   1875
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
      Caption         =   "frmArrange.frx":029E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmArrange.frx":02EC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmArrange.frx":030C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDown 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   240
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
         Caption         =   "frmArrange.frx":0328
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmArrange.frx":0350
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmArrange.frx":0370
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdUp 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
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
         Caption         =   "frmArrange.frx":038C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmArrange.frx":03B0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmArrange.frx":03D0
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgForms 
      Height          =   3315
      Left            =   180
      TabIndex        =   7
      Top             =   1200
      Width           =   4035
      _cx             =   7117
      _cy             =   5847
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
   Begin HexUniControls.ctlUniRadioXP optMaximize 
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   780
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
      Caption         =   "frmArrange.frx":03EC
      Enabled         =   -1  'True
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmArrange.frx":0444
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmArrange.frx":0464
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboRows 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   138
      Width           =   555
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
      Tip             =   "frmArrange.frx":0480
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmArrange.frx":04A0
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   3180
      TabIndex        =   3
      Top             =   660
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
      Caption         =   "frmArrange.frx":04BC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmArrange.frx":04EA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmArrange.frx":050A
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   435
      Left            =   3180
      TabIndex        =   2
      Top             =   120
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
      Caption         =   "frmArrange.frx":0526
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmArrange.frx":054C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmArrange.frx":056C
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniRadioXP optCascade 
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   480
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
      Caption         =   "frmArrange.frx":0588
      Enabled         =   -1  'True
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmArrange.frx":05C6
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmArrange.frx":05E6
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniRadioXP optTile 
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   183
      Width           =   1395
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
      Caption         =   "frmArrange.frx":0602
      Enabled         =   -1  'True
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmArrange.frx":0644
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmArrange.frx":0664
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   255
      Left            =   2250
      Top             =   195
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
      Caption         =   "frmArrange.frx":0680
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmArrange.frx":06A8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmArrange.frx":06C8
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eTbTileFields
    eTb_TileIndex = 0
    eTb_Top
    eTb_Left
    eTb_Caption
    eTb_TileTemplate
    eTb_TilePeriodicity
End Enum

Private Enum eTbTabFields
    eTb_TabIndex = 0
    eTb_DefaultTab
    eTb_CustomTab
    eTb_TabTemplate
    eTb_TabNameChanged
    eTb_TabPeriodicity
End Enum

Private Enum eGridCols
    eFg_Checkbox = 0
    eFg_TabName
    eFg_ChartName
    eFg_Index
End Enum

Private Type mPrivate
    tbTileForms As New cGdTable
    tbTabForms As New cGdTable
    aTabChartsOrder As New cGdArray
    
    nNumRows As Long
    nMinimizedHeight As Long
    nMouseDownRow As Long
    nForms As Long
    
    strIniFile As String
    strView As String
    
    bAlphabetical As Boolean
    bPeriodicity As Boolean
    bTabNameChanged As Boolean
End Type

Private m As mPrivate

Private Sub cmdAlpha_Click()
On Error GoTo ErrSection:

    m.bAlphabetical = True
    m.bPeriodicity = False
    ShowGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.cmdAlpha.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdDown_Click()
On Error GoTo ErrSection:

    Dim nTbIdx1&, nTbIdx2&
    Dim strRec1$, strRec2$
    Dim strCustomTab$, strDefaultTab$

    With fgForms
        If .Row + 1 <= .Rows - 1 Then
            nTbIdx1 = Val(.TextMatrix(.Row, eFg_Index))
            nTbIdx2 = Val(.TextMatrix(.Row + 1, eFg_Index))
        Else
            nTbIdx1 = -1
            nTbIdx2 = -1
        End If
    End With

    If nTbIdx1 >= 0 And nTbIdx2 >= 0 Then
        If m.strView = "Tab" Then
            strRec1 = m.tbTabForms.GetRecord(nTbIdx1, vbTab)
            strRec2 = m.tbTabForms.GetRecord(nTbIdx2, vbTab)
            m.tbTabForms.SetRecord strRec1, nTbIdx2, vbTab
            m.tbTabForms.SetRecord strRec2, nTbIdx1, vbTab
            With fgForms
                strDefaultTab = m.tbTabForms(eTb_DefaultTab, nTbIdx1)
                strCustomTab = m.tbTabForms(eTb_CustomTab, nTbIdx1)
                If strCustomTab <> strDefaultTab Then
                    .TextMatrix(.Row, eFg_TabName) = strCustomTab
                    .Cell(flexcpChecked, .Row, eFg_Checkbox) = flexChecked
                Else
                    .TextMatrix(.Row, eFg_TabName) = strDefaultTab
                    .Cell(flexcpChecked, .Row, eFg_Checkbox) = flexUnchecked
                End If
                .TextMatrix(.Row, eFg_Index) = Str(nTbIdx1)
            
                strDefaultTab = m.tbTabForms(eTb_DefaultTab, nTbIdx2)
                strCustomTab = m.tbTabForms(eTb_CustomTab, nTbIdx2)
                If strCustomTab <> strDefaultTab Then
                    .TextMatrix(.Row + 1, eFg_TabName) = strCustomTab
                    .Cell(flexcpChecked, .Row + 1, eFg_Checkbox) = flexChecked
                Else
                    .TextMatrix(.Row + 1, eFg_TabName) = strDefaultTab
                    .Cell(flexcpChecked, .Row + 1, eFg_Checkbox) = flexUnchecked
                End If
                .TextMatrix(.Row + 1, eFg_Index) = Str(nTbIdx2)
                .Row = .Row + 1
            End With
        Else
            strRec1 = m.tbTileForms.GetRecord(nTbIdx1, vbTab)
            strRec2 = m.tbTileForms.GetRecord(nTbIdx2, vbTab)
            m.tbTileForms.SetRecord strRec1, nTbIdx2, vbTab
            m.tbTileForms.SetRecord strRec2, nTbIdx1, vbTab
            With fgForms
                .TextMatrix(.Row, eFg_ChartName) = m.tbTileForms(eTb_Caption, nTbIdx1)
                .TextMatrix(.Row, eFg_Index) = nTbIdx1
                .TextMatrix(.Row + 1, eFg_ChartName) = m.tbTileForms(eTb_Caption, nTbIdx2)
                .TextMatrix(.Row + 1, eFg_Index) = nTbIdx2
                .Row = .Row + 1
            End With
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.cmdDown.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    MoveFocus cmdOK
    DoEvents

    If optCascade Then
        Me.Tag = "C"
    ElseIf optMaximize Then
        Me.Tag = "M"
        SaveTabOrder
    Else
        Me.Tag = "T"
    End If
    Me.Hide
    ArrangeCharts
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdPeriodicity_Click()
On Error GoTo ErrSection:

    m.bAlphabetical = False
    m.bPeriodicity = True
    ShowGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.cmdPeriodicity.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdUp_Click()
On Error GoTo ErrSection:

    Dim nTbIdx1&, nTbIdx2&
    Dim strRec1$, strRec2$
    Dim strDefaultTab$, strCustomTab$

    With fgForms
        If .Row - 1 >= .FixedRows Then
            nTbIdx1 = Val(.TextMatrix(.Row, eFg_Index))
            nTbIdx2 = Val(.TextMatrix(.Row - 1, eFg_Index))
        Else
            nTbIdx1 = -1
            nTbIdx2 = -1
        End If
    End With

    If nTbIdx1 >= 0 And nTbIdx2 >= 0 Then
        If m.strView = "Tab" Then
            strRec1 = m.tbTabForms.GetRecord(nTbIdx1, vbTab)
            strRec2 = m.tbTabForms.GetRecord(nTbIdx2, vbTab)
            m.tbTabForms.SetRecord strRec1, nTbIdx2, vbTab
            m.tbTabForms.SetRecord strRec2, nTbIdx1, vbTab
            With fgForms
                strDefaultTab = m.tbTabForms(eTb_DefaultTab, nTbIdx1)
                strCustomTab = m.tbTabForms(eTb_CustomTab, nTbIdx1)
                If strCustomTab <> strDefaultTab Then
                    .TextMatrix(.Row, eFg_TabName) = strCustomTab
                    .Cell(flexcpChecked, .Row, eFg_Checkbox) = flexChecked
                Else
                    .TextMatrix(.Row, eFg_TabName) = strDefaultTab
                    .Cell(flexcpChecked, .Row, eFg_Checkbox) = flexUnchecked
                End If
                .TextMatrix(.Row, eFg_Index) = Str(nTbIdx1)
            
                strDefaultTab = m.tbTabForms(eTb_DefaultTab, nTbIdx2)
                strCustomTab = m.tbTabForms(eTb_CustomTab, nTbIdx2)
                If strCustomTab <> strDefaultTab Then
                    .TextMatrix(.Row - 1, eFg_TabName) = strCustomTab
                    .Cell(flexcpChecked, .Row - 1, eFg_Checkbox) = flexChecked
                Else
                    .TextMatrix(.Row - 1, eFg_TabName) = strDefaultTab
                    .Cell(flexcpChecked, .Row - 1, eFg_Checkbox) = flexUnchecked
                End If
                .TextMatrix(.Row - 1, eFg_Index) = Str(nTbIdx2)
                .Row = .Row - 1
            End With
        Else
            strRec1 = m.tbTileForms.GetRecord(nTbIdx1, vbTab)
            strRec2 = m.tbTileForms.GetRecord(nTbIdx2, vbTab)
            m.tbTileForms.SetRecord strRec1, nTbIdx2, vbTab
            m.tbTileForms.SetRecord strRec2, nTbIdx1, vbTab
            With fgForms
                .TextMatrix(.Row, eFg_ChartName) = m.tbTileForms(eTb_Caption, nTbIdx1)
                .TextMatrix(.Row, eFg_Index) = Str(nTbIdx1)
                .TextMatrix(.Row - 1, eFg_ChartName) = m.tbTileForms(eTb_Caption, nTbIdx2)
                .TextMatrix(.Row - 1, eFg_Index) = Str(nTbIdx2)
                .Row = .Row - 1
            End With
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.cmdUp.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgForms_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim nIdx&

    With fgForms
        If .Col = eFg_TabName Then
            nIdx = Val(.TextMatrix(.Row, eFg_Index))
            If Len(.TextMatrix(.Row, eFg_TabName)) > 0 Then
                m.tbTabForms(eTb_CustomTab, nIdx) = .TextMatrix(.Row, eFg_TabName)
                m.tbTabForms(eTb_TabNameChanged, nIdx) = 1
                m.bTabNameChanged = True
                .Cell(flexcpChecked, Row, eFg_Checkbox) = 1
            Else
                .TextMatrix(.Row, eFg_TabName) = m.tbTabForms(eTb_DefaultTab, nIdx)
                m.tbTabForms(eTb_CustomTab, nIdx) = m.tbTabForms(eTb_DefaultTab, nIdx)
                m.bTabNameChanged = True
                .Cell(flexcpChecked, .Row, eFg_Checkbox) = 2
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.fgForms.AfterEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgForms_AfterMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:
    
    With fgForms
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .Row = .MouseRow
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.fgForms.AfterMoveRow", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgForms_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    With fgForms
        If Col = eFg_TabName Then
            If Not optMaximize Then
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.fgForms.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgForms_Click()
On Error GoTo ErrSection:

    With fgForms
        If .Col = eFg_Checkbox Then
            ToggleCheckbox
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.fgForms.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgForms_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ' save row when MouseDown occurred in order to start dragging in MouseMove
    m.nMouseDownRow = fgForms.MouseRow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.fgForms.MouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgForms_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim nRow As Long

    On Error Resume Next

    With fgForms
        If m.nMouseDownRow <> .MouseRow And m.nMouseDownRow >= .FixedRows And .MouseRow >= .FixedRows Then
            nRow = m.nMouseDownRow
            m.nMouseDownRow = 0
            .DragRow nRow
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.fgForms.MouseMove", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgForms_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    m.nMouseDownRow = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.fgForms.MouseUp", eGDRaiseError_Show
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
    RaiseError "frmArrange.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    'Can't do these when focus is in the grid (so can type in a custom tab name)
    If fgForms Is ActiveControl Then Exit Sub

    Select Case KeyAscii
    Case 27 '(Escape)
        KeyAscii = 0
        cmdCancel_Click
    Case Asc("C"), Asc("c")
        KeyAscii = 0
        Me.optCascade = True
        cmdOK_Click
    Case Asc("M"), Asc("m")
        KeyAscii = 0
        Me.optMaximize = True
        cmdOK_Click
    Case Asc("T"), Asc("t") ', 13
        KeyAscii = 0
        Me.optTile = True
        cmdOK_Click
    Case Asc("1") To Asc("9")
        cboRows.ListIndex = KeyAscii - Asc("1")
        KeyAscii = 0
        Me.optTile = True
        cmdOK_Click
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.Form.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("kSelect"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub ShowMe()
On Error GoTo ErrSection:

    Dim frm As frmChart, i&
    
    On Error Resume Next
    
    m.bTabNameChanged = False
    m.bAlphabetical = False
    m.bPeriodicity = False
    'get settings from INI
    m.strIniFile = g.ChartGlobals.strCPCRoot & "\Charts\Page.ini"
    m.nNumRows = GetIniFileProperty("ArrangeRows", 0, "TileCascade", m.strIniFile)
    m.strView = GetIniFileProperty("View", "", "TileCascade", m.strIniFile)
    
    Set frm = ActiveChart
    If Not frm Is Nothing Then
        If frm.WindowState = vbMaximized Then
            optMaximize.Value = True
            m.strView = "Tab"
        ElseIf m.strView = "Cascade" Then
            optCascade.Value = True
        Else
            optTile.Value = True    'default to tile if view info have never been saved
            m.strView = "Tile"
        End If
    End If
            
    LoadTable
    InitGrid
    ShowGrid
    
    'set number of rows drop-down
    cboRows.Clear
    For i = 1 To 9
        cboRows.AddItem CStr(i)
    Next
    cboRows.Text = CStr(m.nNumRows)
    
    ' arrange minimized chart icons
    frmMain.Arrange vbArrangeIcons
    
    ' show the form
    Me.Tag = ""
    ShowForm Me, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.ShowMe", eGDRaiseError_Raise

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        Me.Tag = ""
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optCascade_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        Me.Tag = "C"
        m.strView = "Cascade"
        ShowGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.optCascade.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optMaximize_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        Me.Tag = "M"
        m.strView = "Tab"
        ShowGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.optMaximize.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub optTile_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        Me.Tag = "T"
        m.strView = "Tile"
        ShowGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.optTile.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub ArrangeCharts(Optional ByVal bAutoTile As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, iForm&, nPerRow&, dRoundTo#, iNumRects&
    Dim nLeft&, nTop&, nHeight&, nWidth&
    Dim nOrigMdiWidth&, nOrigMdiHeight&, iLoops%
    Dim nIdx&, strTabs$, nScrollBar&
    Dim bClearInfBox As Boolean
    Dim frm As frmChart
    Dim Rects() As Rect
    
    Static nPrevWidth&, nPrevHeight&, bDisableAutoTile As Boolean
    
    If bAutoTile Then
        ' TLB 4/17/2008: AutoTile is now ONLY used when loading an older chart page that does not
        ' have the MDI client ratios assigned (i.e. for backwards-compatibility of pages that were
        ' intended to be tiled) -- so, we should now always start with bDisableAutoTile set to False.
        ' (and we are also now eliminating some of the conditions that we used to exit for)
        bDisableAutoTile = False
        
        ' only need to do this if client area of frmMain has changed
        If bDisableAutoTile Or g.bStarting Or g.bUnloading Then
            Exit Sub
        ElseIf frmMain.WindowState = vbMinimized Then
            Exit Sub
        'ElseIf nPrevWidth = frmMain.ScaleWidth And nPrevHeight = frmMain.ScaleHeight Then
        '    Exit Sub
        'ElseIf MouseIsPressed Then
        '    Exit Sub ' and wait until the mouse button is up (e.g. to wait until finished resizing)
        ElseIf ActiveChart Is Nothing Then
            Exit Sub
        ElseIf ActiveChart.WindowState = vbMaximized Then
            Exit Sub ' if a maximized chart, then quit
        End If
        
        ' get settings from INI file
        m.strIniFile = g.ChartGlobals.strCPCRoot & "\Charts\Page.ini"
        m.nNumRows = GetIniFileProperty("ArrangeRows", 0, "TileCascade", m.strIniFile)
        If m.nNumRows <= 0 Then m.nNumRows = 2
            
        ' make sure all charts are still same size and don't overlap
        ReDim Rects(Forms.Count) As Rect
        For iForm = 0 To Forms.Count - 1
            If TypeOf Forms(iForm) Is frmChart Then         'only want non-detached charts
                Set frm = Forms(iForm)
                If frm.DetachStatus = eNotDetached Then
                    If frm.WindowState = 0 Then '(ignore minimized and maximized)
                        iNumRects = iNumRects + 1
                        With Rects(iNumRects)
                            ' TLB: must convert to pixels since Twips not always rounding correctly
                            .Left = Round(frm.Left / Screen.TwipsPerPixelX)
                            .Top = Round(frm.Top / Screen.TwipsPerPixelY)
                            .Right = Round((frm.Left + frm.Width) / Screen.TwipsPerPixelX)
                            .Bottom = Round((frm.Top + frm.Height) / Screen.TwipsPerPixelY)
                        End With
                        ' see if this chart overlaps any other chart or is not the same size
                        For i = 1 To iNumRects - 1
                            If Abs((Rects(iNumRects).Right - Rects(iNumRects).Left) - (Rects(i).Right - Rects(i).Left)) >= 2 Then
                                bDisableAutoTile = True
                            ElseIf Abs((Rects(iNumRects).Bottom - Rects(iNumRects).Top) - (Rects(i).Bottom - Rects(i).Top)) >= 2 Then
                                bDisableAutoTile = True
                            ElseIf Rects(iNumRects).Left < Rects(i).Right - 1 And _
                                    Rects(iNumRects).Right > Rects(i).Left + 1 And _
                                    Rects(iNumRects).Top < Rects(i).Bottom - 1 And _
                                    Rects(iNumRects).Bottom > Rects(i).Top + 1 Then
                                bDisableAutoTile = True
                            End If
                            If bDisableAutoTile Then Exit For
                        Next
                    End If
                End If
                If bDisableAutoTile Then Exit For
            End If
        Next
        Set frm = Nothing
        ReDim Rects(0) As Rect
        If bDisableAutoTile Then
            Exit Sub
        End If
        
        Me.Tag = ""
        If m.tbTileForms.NumRecords = 0 Or m.nForms <> Forms.Count Then
            LoadTable True
        End If
        If m.tbTileForms.NumRecords <> fgForms.Rows - fgForms.FixedRows Then
            InitGrid
            LoadTileGrid
        End If
        If m.tbTileForms.NumRecords = fgForms.Rows - fgForms.FixedRows Then
            Me.Tag = "T"
            nScrollBar = GetSystemMetrics(2) * Screen.TwipsPerPixelX       'SM_CXVSCROLL (width of arrow bitmap on vertical scroll bar - winuser.h)
            If m.tbTileForms.NumRecords > 12 Then
                InfBox "Arranging Charts.  Please wait...", , , "Arranging Charts...", True
                bClearInfBox = True
            End If
        End If
    Else
        m.nNumRows = ValOfText(cboRows.Text)
        If m.nNumRows < 1 Then
            ' Shouldn't happen -- but if it does, don't try to continue
            Beep
            Exit Sub
        End If
    End If
        
    ChartTimers = False
        
    If Me.Tag = "M" Then
        ' maximize chart
        Set frm = ActiveChart
        If Not frm Is Nothing Then
            frm.WindowState = vbMaximized
            frm.SetChartTabs m.bTabNameChanged
        End If
    ElseIf Me.Tag = "C" Then
        m.strView = "Cascade"
        ' cascade charts
        LockWindowUpdate frmMain.hWnd
        'For i = 0 To m.tbForms.NumRecords - 1
        For i = fgForms.FixedRows To fgForms.Rows - 1
            nIdx = Val(fgForms.TextMatrix(i, eFg_Index))
            iForm = m.tbTileForms(eTb_TileIndex, nIdx)
            Set frm = Forms(iForm)
            frm.ZOrder
        Next
        frmMain.Arrange vbCascade
        DoEvents
        LockWindowUpdate 0
    ElseIf Me.Tag = "T" Then
        ' tile charts
        m.strView = "Tile"
        bDisableAutoTile = False
        'DoEvents
        LockWindowUpdate frmMain.hWnd
        g.bLoadingChartPage = True '(set flag so resizing will not generate chart until all are done)
        ' unmaximize
        Set frm = ActiveChart
        If Not frm Is Nothing Then
            If frm.WindowState = 2 Then frm.WindowState = 0
            Set frm = Nothing
        End If
        nOrigMdiWidth = frmMain.ScaleWidth
        nOrigMdiHeight = frmMain.ScaleHeight
        For iLoops = 1 To 2
            If g.bUnloading Then Exit For
            ' place charts
            If m.nNumRows <= 0 Then     '4145
                m.nNumRows = ValOfText(cboRows.Text)
                If m.nNumRows <= 0 Then m.nNumRows = 2      'theoretically should not happen, but double check anyways
            End If
            nPerRow = (m.tbTileForms.NumRecords - 1) \ m.nNumRows + 1 ' # columns
            If nPerRow <= 0 Or g.bStarting Or g.bUnloading Or frmMain.WindowState = vbMinimized Then
                Exit For        'aardvark 3675
            End If
            
            nHeight = (frmMain.ScaleHeight - m.nMinimizedHeight) \ m.nNumRows
            nWidth = frmMain.ScaleWidth \ nPerRow
            
            If nWidth < kMinChartWidth Then nWidth = kMinChartWidth
            If nHeight < kMinChartHeight Then nHeight = kMinChartHeight
            
            nLeft = 0
            nTop = -nHeight
            'For i = 0 To m.tbForms.NumRecords - 1
            For i = 0 To fgForms.Rows - 2
                nIdx = Val(fgForms.TextMatrix(fgForms.FixedRows + i, eFg_Index))
                iForm = m.tbTileForms(eTb_TileIndex, nIdx)
                If iForm < m.nForms Then        'precautionary
                    If TypeOf Forms(iForm) Is frmChart Then         'only want non-detached charts
                        Set frm = Forms(iForm)
                        If frm.DetachStatus = eNotDetached Then
                            If i Mod nPerRow = 0 Then
                                ' start a new row
                                nTop = nTop + nHeight
                                nLeft = -nWidth
                                ' if last row, shove to right
                                'If m.tbForms.NumRecords - i < nPerRow Then
                                    'nLeft = nLeft + nWidth * (nPerRow - (nNumCharts - i))
                                'End If
                            End If
                            nLeft = nLeft + nWidth
                            With frm
                                If .WindowState <> 0 Then .WindowState = 0
                                If .Left <> nLeft Or .Top <> nTop Or .Width <> nWidth Or .Height <> nHeight Then
                                    .Move nLeft, nTop, nWidth, nHeight
                                End If
                            End With
                        End If
                    End If
                End If
            Next
            Set frm = Nothing
            
            ' if size of MDI client area did NOT change
            ' (due to scrollbars disappearing), then done
            DoEvents
            If nOrigMdiWidth = frmMain.ScaleWidth And _
                nOrigMdiHeight = frmMain.ScaleHeight Then
                    Exit For
            End If
            ' otherwise need to do this one more time
        Next
        
        ' now call resize for all charts so they will
        ' all regenerate now (this way, it's just once)
        g.bLoadingChartPage = False '(clear flag)
        For i = 0 To m.tbTileForms.NumRecords - 1
            If g.bUnloading Then Exit For
            iForm = m.tbTileForms(eTb_TileIndex, i)
            If iForm < m.nForms Then        'precautionary
                FormResize Forms(iForm)
            End If
        Next
        DoEvents
        LockWindowUpdate 0
    End If
        
    ChartTimers = True

    'save settings
    If Not bAutoTile Then
        SetIniFileProperty "ArrangeRows", m.nNumRows, "TileCascade", m.strIniFile
        If m.strView <> "Tab" Then
            SetIniFileProperty "View", m.strView, "TileCascade", m.strIniFile
        End If
        
        If Not ActiveChart Is Nothing Then
            If ActiveChart.DetachStatus = eDetached Then
                Set frm = frmMain.ActiveForm
                If Not frm Is Nothing Then
                    If TypeOf frm Is frmChart Then              'only want non-detached charts
                        If frm.DetachStatus = eNotDetached Then
                            SendMessage frm.hWnd, WM_NCACTIVATE, 1, 0
                        End If
                    End If
                End If
            End If
        End If
        
        'set dirty flag (fix for aardvark 2500)
        g.bDirtyChartPage = True
    End If
    
    nPrevWidth = frmMain.ScaleWidth
    nPrevHeight = frmMain.ScaleHeight
    If bClearInfBox Then InfBox ""

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmArrange.ArrangeCharts", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgForms
        .Redraw = flexRDNone
        SetupGrid Me.fgForms, eGridMode_Grid
        .ExplorerBar = flexExNone
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .Rows = 1
        .Cols = 4
        'alignment
        .ColAlignment(0) = flexAlignCenterCenter    'check box
        'column headers
        .TextMatrix(0, eFg_Checkbox) = "Custom Tab"
        .TextMatrix(0, eFg_TabName) = "Tab Name"
        .TextMatrix(0, eFg_ChartName) = "Chart Name"
        .TextMatrix(0, eFg_Index) = "Idx"
        'hide last column
        .ColHidden(3) = True
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.InitGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub LoadTabGrid()
On Error GoTo ErrSection:
    
    Dim i&, nIdx&, nRow&
    Dim aIdx As cGdArray
    Dim strCustom$, strDefault$
    Dim strTemplate$
    Dim ActiveForm As frmChart
    
    If m.tbTabForms.NumRecords = 0 Then Exit Sub
    
    If m.bPeriodicity Then
        Set aIdx = m.tbTabForms.CreateSortedIndex(eTb_TabPeriodicity, eGdSort_Descending)   '6989
    ElseIf m.bAlphabetical Or m.aTabChartsOrder.Size = 0 Then
        Set aIdx = m.tbTabForms.CreateSortedIndex(eTb_CustomTab, eGdSort_Default)
    End If
    
    Set ActiveForm = ActiveChart
    If Not ActiveForm Is Nothing Then
        strTemplate = ActiveForm.Chart.Template
    End If
    Set ActiveForm = Nothing
        
    With fgForms
        .Rows = .FixedRows
        For i = 0 To m.tbTabForms.NumRecords - 1
            If m.bPeriodicity Or m.bAlphabetical Or m.aTabChartsOrder.Size = 0 Then
                nIdx = aIdx(i)
            Else
                nIdx = i
            End If
            .Rows = .Rows + 1
            strDefault = m.tbTabForms(eTb_DefaultTab, nIdx)
            strCustom = m.tbTabForms(eTb_CustomTab, nIdx)
            If strCustom <> strDefault Then
                .TextMatrix(.Rows - 1, eFg_TabName) = strCustom
                .Cell(flexcpChecked, .Rows - 1, eFg_Checkbox) = flexChecked
            Else
                .TextMatrix(.Rows - 1, eFg_TabName) = strDefault
                .Cell(flexcpChecked, .Rows - 1, eFg_Checkbox) = flexUnchecked
            End If
            .TextMatrix(.Rows - 1, eFg_Index) = Str(nIdx)
            'set row number to highlight
            If m.tbTabForms(eTb_TabTemplate, nIdx) = strTemplate Then nRow = .Rows - 1
        Next
        .Cell(flexcpPictureAlignment, .FixedRows, eFg_Checkbox, .Rows - 1, eFg_Checkbox) = flexPicAlignCenterCenter
        If nRow >= .FixedRows And nRow < .Rows Then
            .Row = nRow
        End If
    End With
        
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.LoadTabGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub LoadTileGrid()
On Error GoTo ErrSection:

    Dim i&, nIdx&, nRow&
    Dim aIdx As cGdArray
    Dim strTemplate$
    Dim ActiveForm As frmChart
    
    If m.bPeriodicity Then
        Set aIdx = m.tbTileForms.CreateSortedIndex(eTb_TilePeriodicity, eGdSort_Descending)     '6989
    ElseIf m.bAlphabetical Then
        Set aIdx = m.tbTileForms.CreateSortedIndex(eTb_Caption, eGdSort_Default)
    Else
        Set aIdx = m.tbTileForms.CreateSortedIndex(eTb_Top, eGdSort_Default, eTb_Left, eGdSort_Stable)
    End If
    
    Set ActiveForm = ActiveChart
    If Not ActiveForm Is Nothing Then
        strTemplate = ActiveForm.Chart.Template
    End If
    Set ActiveForm = Nothing
    
    With fgForms
        .Rows = .FixedRows
        For i = 0 To aIdx.Size - 1
            .Rows = .Rows + 1
            nIdx = aIdx(i)
            .TextMatrix(.Rows - 1, eFg_ChartName) = m.tbTileForms(eTb_Caption, nIdx)
            .TextMatrix(.Rows - 1, eFg_Index) = Str(nIdx)
            'set row to highlight
            If strTemplate = m.tbTileForms(eTb_TileTemplate, nIdx) Then
                nRow = .Rows - 1
            End If
        Next
        If nRow >= .FixedRows And nRow < .Rows Then
            .Row = nRow
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.LoadTileGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub ShowGrid()
On Error GoTo ErrSection:

    With fgForms
        .Redraw = flexRDNone
        If m.strView = "Tab" Then
            LoadTabGrid
            .ColHidden(eFg_Checkbox) = False
            .ColHidden(eFg_TabName) = False
            .ColHidden(eFg_ChartName) = True
        Else
            LoadTileGrid
            .ColHidden(eFg_Checkbox) = True
            .ColHidden(eFg_TabName) = True
            .ColHidden(eFg_ChartName) = False
        End If
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.ShowGrid", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub LoadTable(Optional ByVal bAutoTile As Boolean = False)
On Error GoTo ErrSection:

    Dim strDefaultTab$, strCustomTab$, strCaption$
    Dim frm As frmChart
    Dim tbTemp As New cGdTable
    Dim aIdx As cGdArray
    Dim iForm&, nNumCharts&, strChtOrder$
    Dim i&, h&, iPos&, nTop&, dRoundTo#
    Dim wp As WINDOWPLACEMENT
    
    'create fields for tile table
    If m.tbTileForms.NumFields = 0 Then
        m.tbTileForms.CreateField eGDARRAY_Longs, 0, "Index"
        m.tbTileForms.CreateField eGDARRAY_Longs, 1, "FormTop"
        m.tbTileForms.CreateField eGDARRAY_Longs, 2, "FormLeft"
        m.tbTileForms.CreateField eGDARRAY_Strings, 3, "FormCaption"
        m.tbTileForms.CreateField eGDARRAY_Strings, 4, "TemplateFileName"
        m.tbTileForms.CreateField eGDARRAY_Longs, 5, "Periodicity"
    End If
    m.tbTileForms.NumRecords = 0
    
    'create fields for tab table
    If m.tbTabForms.NumFields = 0 Then
        m.tbTabForms.CreateField eGDARRAY_Longs, 0, "Index"
        m.tbTabForms.CreateField eGDARRAY_Strings, 1, "DefaultTab"
        m.tbTabForms.CreateField eGDARRAY_Strings, 2, "CustomTab"
        m.tbTabForms.CreateField eGDARRAY_Strings, 3, "TemplateFileName"
        m.tbTabForms.CreateField eGDARRAY_Longs, 4, "TabNameChanged"
        m.tbTabForms.CreateField eGDARRAY_Longs, 5, "Periodicity"
    End If
    m.tbTabForms.NumRecords = 0

    'create fields for temporary table
    If Not bAutoTile Then
        tbTemp.CreateField eGDARRAY_Longs, 0, "Index"
        tbTemp.CreateField eGDARRAY_Strings, 1, "DefaultTab"
        tbTemp.CreateField eGDARRAY_Strings, 2, "CustomTab"
        tbTemp.CreateField eGDARRAY_Strings, 3, "TemplateFileName"
        tbTemp.CreateField eGDARRAY_Longs, 4, "TabNameChanged"
        tbTemp.CreateField eGDARRAY_Longs, 5, "Periodicity"
        tbTemp.NumRecords = 0
        
        strChtOrder = GetIniFileProperty("ChartsOrder", "", "Tab", m.strIniFile)
        m.aTabChartsOrder.SplitFields strChtOrder, ","
    End If
    
    ' set number of rows
    If m.nNumRows = 0 Then
        For iForm = 0 To Forms.Count - 1
            If TypeOf Forms(iForm) Is frmChart Then         'only want non-detached charts
                Set frm = Forms(iForm)
                If frm.Visible And frm.WindowState <> 1 And frm.DetachStatus = eNotDetached Then
                    nNumCharts = nNumCharts + 1
                End If
            End If
        Next
        If nNumCharts <= 1 Then
            m.nNumRows = 1
        'ElseIf m.nNumRows = 0 Or m.nNumRows > nNumCharts Then      'original code
        Else
            ' default to 2 charts per row
            m.nNumRows = (nNumCharts + 1) \ 2
            If m.nNumRows < 2 Then
                m.nNumRows = 2
            ElseIf m.nNumRows > 9 Then
                m.nNumRows = 9
            End If
        End If
    ElseIf m.nNumRows > 9 Then
        m.nNumRows = 9
    End If
    
    wp.Length = Len(wp)
    dRoundTo = frmMain.ScaleHeight \ m.nNumRows
    m.nForms = Forms.Count
    ' find all the non-minimized charts
    m.nMinimizedHeight = 0
    For iForm = 0 To Forms.Count - 1
        If TypeOf Forms(iForm) Is frmChart Then         'only want non-detached charts
            Set frm = Forms(iForm)
            If frm.Visible And frm.DetachStatus = eNotDetached Then
                If frm.WindowState = vbMinimized And m.strView <> "Tab" Then        'aardvark 3241
                    m.nMinimizedHeight = frm.Height - frm.ScaleHeight
                Else
                    ' save caption in order to sort
                    'aCharts.Add Format(frm.Top, "000000") & Format(frm.Left, "000000") & vbTab & Trim(Parse((frm.Caption), ":", 1)) & vbTab & CStr(iForm)
                    
                    strDefaultTab = Trim(frm.Chart.ChartName(False, True))
                    strCustomTab = Trim(frm.Chart.ChartName(False, False))
                    strCaption = Trim(frm.vseCaption.Caption)
                    strCaption = Trim(frm.Chart.ChartName(True))
                    
                    'populate tile table
                    h = frm.hWnd
                    GetWindowPlacement h, wp
                    m.tbTileForms.AddRecord ""
                    
                    nTop = wp.rcNormalPosition.Top * Screen.TwipsPerPixelY
                    nTop = Int(nTop / dRoundTo + 0.5) * dRoundTo
                    
                    m.tbTileForms(eTb_TileIndex, m.tbTileForms.NumRecords - 1) = iForm
                    m.tbTileForms(eTb_Top, m.tbTileForms.NumRecords - 1) = nTop
                    m.tbTileForms(eTb_Left, m.tbTileForms.NumRecords - 1) = wp.rcNormalPosition.Left * Screen.TwipsPerPixelX
                    m.tbTileForms(eTb_Caption, m.tbTileForms.NumRecords - 1) = strCaption
                    m.tbTileForms(eTb_TileTemplate, m.tbTileForms.NumRecords - 1) = frm.Chart.Template
                    m.tbTileForms(eTb_TilePeriodicity, m.tbTileForms.NumRecords - 1) = frm.Chart.Periodicity
                    
                    'populate temporary table
                    If Not bAutoTile Then
                        tbTemp.AddRecord ""
                        tbTemp(eTb_TabIndex, tbTemp.NumRecords - 1) = iForm
                        tbTemp(eTb_DefaultTab, tbTemp.NumRecords - 1) = strDefaultTab
                        tbTemp(eTb_CustomTab, tbTemp.NumRecords - 1) = strCustomTab
                        tbTemp(eTb_TabTemplate, tbTemp.NumRecords - 1) = frm.Chart.Template
                        tbTemp(eTb_TabNameChanged, tbTemp.NumRecords - 1) = 0
                        tbTemp(eTb_TabPeriodicity, tbTemp.NumRecords - 1) = frm.Chart.Periodicity
                    End If
                End If
            End If
        End If
    Next
    
    If Not bAutoTile Then
        If m.aTabChartsOrder.Size > 0 Then
            Set aIdx = tbTemp.CreateSortedIndex(eTb_TabTemplate, eGdSort_Default)
            For i = 0 To m.aTabChartsOrder.Size - 1
                If tbTemp.SearchAsIndex(aIdx, eTb_TabTemplate, m.aTabChartsOrder(i), iPos) Then
                    If tbTemp(eTb_TabTemplate, aIdx(iPos)) = m.aTabChartsOrder(i) Then
                        m.tbTabForms.AddRecord ""
                        m.tbTabForms.SetRecord tbTemp.GetRecord(aIdx(iPos)), m.tbTabForms.NumRecords - 1
                    End If
                End If
            Next
        Else
            Set aIdx = tbTemp.CreateSortedIndex(eTb_CustomTab, eGdSort_Default)
            For i = 0 To aIdx.Size - 1
                m.tbTabForms.AddRecord ""
                m.tbTabForms.SetRecord tbTemp.GetRecord(aIdx(i)), m.tbTabForms.NumRecords - 1
            Next
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.LoadTable", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub SaveTabOrder()
On Error GoTo ErrSection:

    Dim i&, nTbIdx&
    Dim strOrder$, strCustom$, strDefault$
    Dim strFile$, strTemplate$
    
    With fgForms
        For i = fgForms.FixedRows To fgForms.Rows - 1
            nTbIdx = Val(.TextMatrix(i, eFg_Index))
            strTemplate = m.tbTabForms(eTb_TabTemplate, nTbIdx)
            strDefault = m.tbTabForms(eTb_DefaultTab, nTbIdx)
            strCustom = m.tbTabForms(eTb_CustomTab, nTbIdx)
            If strDefault = strCustom Then strCustom = ""
            If m.tbTabForms(eTb_TabNameChanged, nTbIdx) = 1 Then
                strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strTemplate & ".CHT"
                SetIniFileProperty "ChartName", strCustom, "General", strFile
            End If
            strOrder = strOrder & m.tbTabForms(eTb_TabTemplate, nTbIdx) & ","
        Next
    End With

    If Len(strOrder) > 0 Then
        strOrder = Left(strOrder, Len(strOrder) - 1)  'remove last comma
        SetIniFileProperty "ChartsOrder", strOrder, "Tab", m.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.SaveTabOrder", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub ToggleCheckbox()
On Error GoTo ErrSection:

    Dim i&
    Dim strCustom$, strDefault$

    With fgForms
        i = Val(.TextMatrix(.Row, eFg_Index))
        strDefault = m.tbTabForms(eTb_DefaultTab, i)
        strCustom = m.tbTabForms(eTb_CustomTab, i)
        If strCustom = strDefault Then strCustom = ""
        If .Cell(flexcpChecked, .Row, eFg_Checkbox) = 1 Then
            .TextMatrix(.Row, eFg_TabName) = strDefault
            .Cell(flexcpChecked, .Row, eFg_Checkbox) = 2
            m.bTabNameChanged = True
            m.tbTabForms(eTb_TabNameChanged, i) = 1
            m.tbTabForms(eTb_CustomTab, i) = strDefault
        ElseIf .Cell(flexcpChecked, .Row, eFg_Checkbox) = 2 Then
            .TextMatrix(.Row, eFg_TabName) = strCustom
            .Cell(flexcpChecked, .Row, eFg_Checkbox) = 1
            .Col = eFg_TabName
            .EditCell
        End If
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmArrange.ToggleCheckBox", eGDRaiseError_Raise
    Resume ErrExit

End Sub

