VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmToolbar 
   Caption         =   "Customize Toolbars"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   Icon            =   "frmToolbar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin HexUniControls.ctlUniComboImageXP cboTemplate 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   3795
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
      Tip             =   "frmToolbar.frx":000C
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmToolbar.frx":002C
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraControls 
      Height          =   2820
      Left            =   233
      TabIndex        =   4
      Top             =   5580
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
      Caption         =   "frmToolbar.frx":0048
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmToolbar.frx":0074
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmToolbar.frx":0094
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   1830
         TabIndex        =   3
         Top             =   2355
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
         Caption         =   "frmToolbar.frx":00B0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbar.frx":00DE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbar.frx":0156
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   705
         TabIndex        =   2
         Top             =   2355
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
         Caption         =   "frmToolbar.frx":0172
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmToolbar.frx":0198
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmToolbar.frx":01B8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraOptions 
         Height          =   2250
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   3510
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
         Caption         =   "frmToolbar.frx":01D4
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmToolbar.frx":0212
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmToolbar.frx":0232
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkToolbarWrap 
            Height          =   255
            Left            =   240
            TabIndex        =   0
            Top             =   950
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
            Caption         =   "frmToolbar.frx":024E
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmToolbar.frx":029C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmToolbar.frx":02BC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkUseLargeIcons 
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   270
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
            Caption         =   "frmToolbar.frx":02D8
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmToolbar.frx":0316
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmToolbar.frx":0336
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkIncludeText 
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   610
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
            Caption         =   "frmToolbar.frx":0352
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmToolbar.frx":0392
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmToolbar.frx":03B2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboTbDrawAlign 
            Height          =   315
            Left            =   1500
            TabIndex        =   7
            Top             =   1395
            Width           =   1545
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
            Tip             =   "frmToolbar.frx":03CE
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmToolbar.frx":03EE
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboSkin 
            Height          =   315
            Left            =   1500
            TabIndex        =   6
            Top             =   1785
            Width           =   1545
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
            Tip             =   "frmToolbar.frx":040A
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmToolbar.frx":042A
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label2 
            Height          =   210
            Left            =   240
            Top             =   1815
            Width           =   930
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
            Caption         =   "frmToolbar.frx":0446
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmToolbar.frx":0480
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbar.frx":04A0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   210
            Left            =   240
            Top             =   1425
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
            Caption         =   "frmToolbar.frx":04BC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmToolbar.frx":04F8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmToolbar.frx":0518
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   4485
      Left            =   120
      TabIndex        =   1
      Top             =   930
      Width           =   3795
      _cx             =   6694
      _cy             =   7911
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
   Begin HexUniControls.ctlUniLabelXP lblTemplate 
      Height          =   255
      Left            =   120
      Top             =   0
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
      Caption         =   "frmToolbar.frx":0534
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmToolbar.frx":0590
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmToolbar.frx":05B0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   255
      Left            =   120
      Top             =   690
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
      Caption         =   "frmToolbar.frx":05CC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmToolbar.frx":063C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmToolbar.frx":065C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' All items on toolbars are added to this array once at startup (from ToolbarReset)
Public aItems As New cGdArray
Public iVersion As Integer  '(current toolbar version)

Private Const OptionsFrameHt = 900

Private Type mPrivate
    bChanged As Boolean
    bCancelled As Boolean
    bNewCustom As Boolean
    
    bTbGeneral As Boolean
    bTbWindows As Boolean
    bTbChart As Boolean
    bTbDraw As Boolean
    
    nTbGeneralGridRow As Long
    nTbWindowsGridRow As Long
    nTbChartingGridRow As Long
    nTbDrawGridRow As Long
    
    strTemplate As String
    nUseLargeIcons As Long
    
    tbContent As cGdTable
    aTbIndex As cGdArray
End Type
Private m As mPrivate

Private Sub cboSkin_Click()
On Error Resume Next:
    
    Dim i&
    
    If Me.Visible Then
        i = cboSkin.ListIndex
        If i <> g.eTbSkin Then
            If i >= 0 And i < cboSkin.ListCount Then
                SetIniFileProperty "ToolbarSkin", i, "Toolbars", g.strIniFile
                g.eTbSkin = i
                m.bChanged = True
            End If
            ToolbarMainUpdate True
        End If
    End If

End Sub

Private Sub cboTbDrawAlign_Click()
On Error Resume Next:
    
    Dim i&, iAlign&
    
    If Me.Visible Then
        i = cboTbDrawAlign.ListIndex
        Select Case i
            Case 1
                iAlign = vbAlignLeft
            Case 2
                iAlign = vbAlignTop
            Case 3
                'If optButtonsText.Value = vbChecked Then
                If chkIncludeText.Value = vbChecked Then
                    iAlign = vbAlignTop
                Else
                    iAlign = vbAlignBottom
                End If
            Case Else
                iAlign = vbAlignRight
        End Select
        
        If iAlign <> g.vbeTbAlignDraw Then
            m.bChanged = True
            SetIniFileProperty "ToolbarAlignDraw", iAlign, "Toolbars", g.strIniFile
            g.vbeTbAlignDraw = iAlign
            ToolbarMainUpdate True
        End If
    End If

End Sub

Private Sub cboTemplate_Click()
On Error GoTo ErrSection:

    Dim strFile$, i&
    Dim j&

    If Me.Visible Then
        If UCase(cboTemplate.Text) = "CUSTOM" Then
            If m.bNewCustom Then
                strFile = App.Path & "\toolbarnew.sav"
            Else
                strFile = App.Path & "\toolbarsho.sav"
            End If
        Else
            j = cboTemplate.ListIndex - 1
            If j >= 0 And j < m.aTbIndex.Size Then
                strFile = m.tbContent(5, m.aTbIndex(j))
            End If
        End If
        
        If FileExist(strFile) Then
            FileCopy strFile, App.Path & "\toolbar.sho"
            
            j = cboTemplate.ListIndex - 1
            If j >= 0 And j < m.aTbIndex.Size Then
                i = m.tbContent(3, m.aTbIndex(j))
                If i <> g.nTbLargeIcons Then
                    g.nTbLargeIcons = i
                    chkUseLargeIcons.Value = i
                    SetIniFileProperty "LargeIcons", i, "Toolbars", g.strIniFile
                End If
            
                i = m.tbContent(4, m.aTbIndex(j))
                If i <> g.nTbIncludeText Then
                    g.nTbLargeIcons = i
                    chkIncludeText.Value = i
                    SetIniFileProperty "LargeButtons", i, "Toolbars", g.strIniFile
                End If
            End If
            
            
            ResetToolbar True
            ToolbarMainUpdate True
            PopulateGrid
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.cboTemplate_Click", eGDRaiseError_Show

End Sub

Private Sub chkIncludeText_Click()
On Error GoTo ErrSection:

    Dim i&
    
    If Not Me.Visible Then GoTo ErrExit
    
    i = chkIncludeText.Value
    If i = g.nTbIncludeText Then GoTo ErrExit
    
    SetIniFileProperty "LargeButtons", i, "Toolbars", g.strIniFile
    AlignComboBoxSet
    ResetToolbar
        
    m.bChanged = True
    ToolbarMainUpdate True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.chkIncludeText_Click", eGDRaiseError_Show

End Sub

Private Sub chkToolbarWrap_Click()
On Error GoTo ErrSection

    Dim i&

    i = chkToolbarWrap.Value
    
    If Me.Visible Then
        SetIniFileProperty "ToolbarWrap", -(chkToolbarWrap), "Toolbars", g.strIniFile
        ResetToolbar
        m.bChanged = True
        ToolbarMainUpdate True
    End If

    Exit Sub
ErrSection:
    RaiseError "frmToolbar.chkToolbarWrap_Click", eGDRaiseError_Show

End Sub

Private Sub chkUseLargeIcons_Click()
On Error GoTo ErrSection:

    Dim i&
    
    If Not Me.Visible Then GoTo ErrExit
    
    i = chkUseLargeIcons.Value
    If i = g.nTbLargeIcons Then GoTo ErrExit
    
    SetIniFileProperty "LargeIcons", i, "Toolbars", g.strIniFile
        
    AlignComboBoxSet
    ResetToolbar
    
    m.bChanged = True
    ToolbarMainUpdate True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.chkUseLargeIcons_Click", eGDRaiseError_Show

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:
    
    If FileExist(App.Path & "\toolbarsho.sav") Then
        FileCopy App.Path & "\toolbarsho.sav", App.Path & "\toolbar.sho"
        ResetToolbar True
        m.bCancelled = True
        cmdOK_Click
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.cmdCancel_Click", eGDRaiseError_Show

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection

    MoveFocus cmdOK
    Unload Me

    Exit Sub

ErrSection:
    RaiseError "frmToolbar.cmdOK", eGDRaiseError_Show

End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection

    Dim nRow&, strTool$, strToolbar$, bVisible As Boolean
    
    Dim bTbWasVisible As Boolean
    Dim bTbIsVisible As Boolean
    
    Dim bUpdateDrawTool As Boolean
    
    If fg.Redraw = flexRDNone Then Exit Sub
    If Col <> 0 Or Row < 0 Then Exit Sub
    If fg.Cell(flexcpChecked, Row, 0) = flexNoCheckbox Then Exit Sub
    
    If m.bTbGeneral Or m.bTbWindows Or m.bTbChart Then bTbWasVisible = True
    
    With fg
        strToolbar = Parse(.TextMatrix(Row, 3), "|", 1)
        strTool = Parse(.TextMatrix(Row, 3), "|", 2)
        If CheckedCell(fg, Row, 0) Then
            bVisible = True
        End If
        
        If Len(.TextMatrix(Row, 0)) > 0 Then
            ' toolbar
            If strToolbar = kTbGeneral Then
                m.bTbGeneral = bVisible
                SetIniFileProperty kTbGeneralVisible, bVisible, "Toolbars", g.strIniFile
            ElseIf strToolbar = kTbWindows Then
                m.bTbWindows = bVisible
                SetIniFileProperty kTbWindowsVisible, bVisible, "Toolbars", g.strIniFile
            ElseIf strToolbar = kTbChartSettings Then
                m.bTbChart = bVisible
                SetIniFileProperty kTbChartSettingsVisible, bVisible, "Toolbars", g.strIniFile
            ElseIf strToolbar = kTbDraw Then
                m.bTbDraw = bVisible
                SetIniFileProperty kTbDrawVisible, bVisible, "Toolbars", g.strIniFile
            Else
                frmMain.tbToolbar.ToolBars(strToolbar).Visible = bVisible
            End If

            If bVisible And frmMain.tbToolbar.ToolBars(strToolbar).Tools.Count = 0 Then
                'user turned on a toolbar, but none of the buttons are turned on
                .Cell(flexcpChecked, Row + 1, 0) = flexChecked
                strTool = Parse(.TextMatrix(Row + 1, 3), "|", 2)
                frmMain.tbToolbar.Tools(strTool).TagVariant = CheckedCell(fg, Row, 0)
                frmMain.tbToolbar.Tools(strTool).TagVariant = bVisible
                ResetToolbar
            End If
            
            For nRow = Row + 1 To .Rows - 1
                If Len(.TextMatrix(nRow, 0)) > 0 Then Exit For
                .RowHidden(nRow) = Not bVisible
            Next
        Else
            .Refresh
            ' button
            frmMain.tbToolbar.Tools(strTool).TagVariant = CheckedCell(fg, Row, 0)
            frmMain.tbToolbar.Tools(strTool).TagVariant = bVisible
            ResetToolbar
        
            If Len(strToolbar) > 0 Then
                If frmMain.tbToolbar.ToolBars(strToolbar).Tools.Count = 0 Then      '5846
                    Select Case strToolbar
                        Case kTbGeneral
                            m.bTbGeneral = bVisible
                            SetIniFileProperty kTbGeneralVisible, bVisible, "Toolbars", g.strIniFile
                            .Cell(flexcpChecked, m.nTbGeneralGridRow, 0) = flexUnchecked
                            Row = m.nTbGeneralGridRow
                        Case kTbWindows
                            m.bTbWindows = bVisible
                            SetIniFileProperty kTbWindowsVisible, bVisible, "Toolbars", g.strIniFile
                            .Cell(flexcpChecked, m.nTbWindowsGridRow, 0) = flexUnchecked
                            Row = m.nTbWindowsGridRow
                        Case kTbChartSettings
                            m.bTbChart = bVisible
                            SetIniFileProperty kTbChartSettingsVisible, bVisible, "Toolbars", g.strIniFile
                            .Cell(flexcpChecked, m.nTbChartingGridRow, 0) = flexUnchecked
                            Row = m.nTbChartingGridRow
                        Case kTbDraw
                            m.bTbDraw = bVisible
                            SetIniFileProperty kTbDrawVisible, bVisible, "Toolbars", g.strIniFile
                            .Cell(flexcpChecked, m.nTbDrawGridRow, 0) = flexUnchecked
                            Row = m.nTbDrawGridRow
                    End Select
                    For nRow = Row + 1 To .Rows - 1
                        If Len(.TextMatrix(nRow, 0)) > 0 Then Exit For
                        .RowHidden(nRow) = Not bVisible
                    Next
                End If
            End If
        End If
    
    
    End With
    
    If m.bTbGeneral Or m.bTbWindows Or m.bTbChart Then bTbIsVisible = True

    If bTbWasVisible <> bTbIsVisible Then
        AlignComboBoxSet
        bUpdateDrawTool = True
    ElseIf strToolbar = kTbDraw Then
        bUpdateDrawTool = True
    End If
    ToolbarMainUpdate bUpdateDrawTool
    
    Dim aShow As New cGdArray
    Dim Tool As SSTool
    
    m.bChanged = True
    m.bNewCustom = True
    
    cboTemplate.ListIndex = 0
    ' save buttons to show for "new" custom choices
    aShow.Add Str(iVersion)
    For Each Tool In frmMain.tbToolbar.Tools
        If Tool.Group <> "HelpList" Then
            If Tool.TagVariant Then
                aShow.Add Tool.ID
            End If
        End If
    Next
    
    aShow.Sort eGdSort_IgnoreCase Or eGdSort_DeleteNullValues
    aShow.ToFile App.Path & "\toolbarnew.sav"
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.fg_AfterEdit", eGDRaiseError_Show

End Sub

Private Sub fg_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo ErrSection

    ' this is just to skip over the summary labels when moving up/down
    On Error Resume Next
    If NewRow >= 0 And fg.Redraw <> flexRDNone Then
        If fg.Cell(flexcpChecked, NewRow, 0) = flexNoCheckbox Then
            If OldRow <= NewRow Or NewRow <= 0 Then
                Cancel = True
                fg.Row = NewRow + 1
            Else
                Cancel = True
                fg.Row = NewRow - 1
            End If
        End If
    End If

    Exit Sub
ErrSection:
    RaiseError "frmToolbar.fg_BeforeRowColChange", eGDRaiseError_Show
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
    RaiseError "frmToolbar.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection
    
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Me.Icon = Picture16(ToolbarIcon("ID_CustomizeToolbar"), , True)
    
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.Form_Load", eGDRaiseError_Show

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If LimitFormSize(Me, fraControls.Left + fraControls.Width, fg.Top + 1200 + fraControls.Height) Then Exit Sub

    With fraControls
        .Move .Left, Me.ScaleHeight - .Height - 60
        fg.Move fg.Left, fg.Top, Me.ScaleWidth - fg.Left * 2, .Top - fg.Top - fg.Left
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection

    Dim aShow As New cGdArray
    Dim Tool As SSTool
                
    KillFile App.Path & "\toolbarsho.sav"
    KillFile App.Path & "\toolbarnew.sav"
    
    Set m.tbContent = Nothing
    Set m.aTbIndex = Nothing
    
    ' save buttons to show
    aShow.Add Str(iVersion)
    For Each Tool In frmMain.tbToolbar.Tools
        If Tool.Group <> "HelpList" Then
            If Tool.TagVariant Then
                aShow.Add Tool.ID
            End If
        End If
    Next
    
    aShow.Sort eGdSort_IgnoreCase Or eGdSort_DeleteNullValues
    aShow.ToFile App.Path & "\Toolbar.sho"
    
    If m.bCancelled Then
        If m.nUseLargeIcons <> g.nTbLargeIcons Then
            g.nTbLargeIcons = m.nUseLargeIcons
            SetIniFileProperty "LargeIcons", m.nUseLargeIcons, kTbIniSection, g.strIniFile
            ToolbarReset
        End If
        SetIniFileProperty kTbTemplate, m.strTemplate, kTbIniSection, g.strIniFile
        ToolbarMainUpdate True
    Else
        SetIniFileProperty kTbTemplate, cboTemplate.Text, kTbIniSection, g.strIniFile
        SetIniFileProperty kTbTemplateDate, CDbl(FileDate(App.Path & "\provided\" & cboTemplate.Text & ".sho")), kTbIniSection, g.strIniFile
    End If
    
    'update toolbar on detached charts without updating the main app toolbar again
    If m.bChanged Then ToolbarChartsUpdate
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.Form_Unload", eGDRaiseError_Show

End Sub

Public Sub ShowMe()
On Error GoTo ErrSection

    Dim nItem&, strItem$, strIcon$, strText$, strHidden$, strToolbar$
    Dim iLargeButtons&, iLargeIcons&
    
    Dim bShow As Boolean
    Dim bTbVisible As Boolean
    
    m.bCancelled = False
    m.bNewCustom = False

    chkUseLargeIcons.Value = Abs(g.nTbLargeIcons)       'double-protection for aardvark 6158
    chkIncludeText.Value = Abs(g.nTbIncludeText)
    m.nUseLargeIcons = g.nTbLargeIcons
    
    PopulateCboTemplate
    
    'skin combo box
    cboSkin.Clear
    cboSkin.AddItem "Silver"
    cboSkin.AddItem "Blue"
    cboSkin.AddItem "Aluminum Silver"
    cboSkin.AddItem "Aluminum Blue"
    cboSkin.AddItem "Dark Flat"
    cboSkin.AddItem "Light Flat"
    If g.eTbSkin = eTbSkin_Unknown Then
        cboSkin.ListIndex = 1
    Else
        cboSkin.ListIndex = g.eTbSkin
    End If
    
    chkToolbarWrap.Visible = True
    chkToolbarWrap = Abs(GetIniFileProperty("ToolbarWrap", False, "Toolbars", g.strIniFile))
    m.bTbGeneral = GetIniFileProperty(kTbGeneralVisible, True, "Toolbars", g.strIniFile)
    m.bTbWindows = GetIniFileProperty(kTbWindowsVisible, True, "Toolbars", g.strIniFile)
    m.bTbChart = GetIniFileProperty(kTbChartSettingsVisible, True, "Toolbars", g.strIniFile)
    m.bTbDraw = GetIniFileProperty(kTbDrawVisible, True, "Toolbars", g.strIniFile)
    
    AlignComboBoxSet
    
    With fg
        SetupGrid fg, eGridMode_Grid
        .Redraw = flexRDNone
        .HighLight = flexHighlightNever
        .SelectionMode = flexSelectionFree
        .GridLines = flexGridNone ' flexGridInsetVert
        .BackColor = frmMain.tbToolbar.BackColor
        .FixedCols = 0
        .Cols = 4
        .FixedRows = 0
        .Rows = aItems.Size + 1 '(for now)
        .ColWidth(0) = Screen.TwipsPerPixelY * 30
        .ColWidth(1) = Screen.TwipsPerPixelY * 35
        .ColAlignment(1) = flexAlignCenterCenter
        .ColHidden(3) = True
        .MergeCells = flexMergeSpill
        .Editable = flexEDKbdMouse
        .FillStyle = flexFillRepeat
        .Select 0, 0, .Rows - 1, .Cols - 1
        .Row = 1
        .Redraw = flexRDBuffered
    End With
    
    PopulateGrid

    ShowForm Me, False, frmMain

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.ShowMe", eGDRaiseError_Show
End Sub

Private Sub ToolbarChartsUpdate()
On Error Resume Next:

    Dim i&, iZoomIn&, iZoomOut&, iEraser&
    Dim frm As Form
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
        
'    ToolbarReset True, frmMain.tbToolbar
        
    For i = 0 To Forms.Count - 1
        DoEvents
        If IsFrmChart(Forms(i)) Then
            Set frm = Forms(i)
            If frm.DetachStatus = eDetached Then
                If frm.Chart.ShowToolbar Then
                    If g.vbeTbAlignDraw = vbAlignBottom Then
                        frm.pbTbBackDraw(0).align = vbAlignTop
                    Else
                        frm.pbTbBackDraw(0).align = g.vbeTbAlignDraw
                    End If
                    ToolbarInit2 frm, frm.TbButtonsArray(kTbGeneral)
                    ToolbarInit2 frm, frm.TbButtonsArray(kTbDraw), , kTbDraw, , g.vbeTbAlignDraw
                    
                    ToolbarResize2 frm, frm.pbTbBack, frm.imgTbBack, frm.TbButtonsArray(kTbGeneral), frm.ToolBarWrapGet(kTbGeneral)
                    ToolbarResize2 frm, frm.pbTbBackDraw, frm.imgTbBackDraw, frm.TbButtonsArray(kTbDraw), frm.ToolBarWrapGet(kTbDraw)
                    
                    FormResize frm
                    DoEvents
                End If
            End If
        End If
    Next

End Sub

Private Sub ToolbarMainUpdate(bUpdateDrawTool As Boolean)
On Error GoTo ErrSection

    Dim X&, Y&
    
    ToolbarInit2 frmMain, frmMain.TbButtonsArray(kTbGeneral)
    If bUpdateDrawTool Then ToolbarInit2 frmMain, frmMain.TbButtonsArray(kTbDraw), , kTbDraw, , g.vbeTbAlignDraw
    
    ToolbarResize2 frmMain, frmMain.pbTbBack, frmMain.imgTbBack, frmMain.TbButtonsArray(kTbGeneral), frmMain.ToolBarWrapGet(kTbGeneral)
    
    If bUpdateDrawTool Then
        If frmMain.pbTbBackDraw(0).align <> vbAlignTop Then
            'JM: for some reason, if this is not done here, the background picture box for drawing tools will not resize properly
            If g.vbeTbAlignDraw = vbAlignTop Or g.vbeTbAlignDraw = vbAlignBottom Then
                frmMain.ToolBarBtnSizeGet kTbDraw, X, Y
                frmMain.pbTbBackDraw(0).align = vbAlignTop
                frmMain.pbTbBackDraw(0).Height = (X + 2) * Screen.TwipsPerPixelX
            End If
        End If
        ToolbarResize2 frmMain, frmMain.pbTbBackDraw, frmMain.imgTbBackDraw, frmMain.TbButtonsArray(kTbDraw), frmMain.ToolBarWrapGet(kTbDraw)
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.ToolbarMainUpdate"
    
End Sub

Private Sub AlignComboBoxSet()
On Error GoTo ErrSection:

    'drawing tools alignment combo box
    cboTbDrawAlign.Clear
    cboTbDrawAlign.AddItem "Right"
    cboTbDrawAlign.AddItem "Left"
    cboTbDrawAlign.AddItem "Top"
    
    'If Me.optButtonsText.Value = vbUnchecked And (m.bTbGeneral Or m.bTbWindows Or m.bTbChart) Then
    If chkIncludeText.Value = vbUnchecked And (m.bTbGeneral Or m.bTbWindows Or m.bTbChart) Then
        cboTbDrawAlign.AddItem "Side by side"
    ElseIf g.vbeTbAlignDraw = vbAlignBottom Then
        g.vbeTbAlignDraw = vbAlignTop
    End If
    
    Select Case g.vbeTbAlignDraw
        Case vbAlignLeft
            cboTbDrawAlign.ListIndex = 1
        Case vbAlignTop
            cboTbDrawAlign.ListIndex = 2
        Case vbAlignBottom
            cboTbDrawAlign.ListIndex = 3
        Case Else
            cboTbDrawAlign.ListIndex = 0
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.AlignComboBoxSet"

End Sub

Private Sub ResetToolbar(Optional ByVal bReinit As Boolean = False)

'    Dim bLocked As Boolean
    
'JM 07-30-2010: locking window appear not needed, leave awhile then remove if all okay
'    bLocked = LockWindowUpdate(GetDesktopWindow())
        ToolbarReset bReinit
'    If bLocked Then LockWindowUpdate 0
End Sub

Private Sub PopulateCboTemplate()
On Error GoTo ErrSection:

    Dim i&, j&, iListIndex&
    Dim strFileMask$, strFile$, strName$, strEnable$
    Dim nOrder&, nLargeIcon&, nIncludeText&
    
    Dim aFiles As New cGdArray
    Dim aTemp As New cGdArray
    
    strFileMask = App.Path & "\provided\*.SHO"
    aFiles.GetMatchingFiles strFileMask, False, False, True
    
    m.strTemplate = GetIniFileProperty(kTbTemplate, "", kTbIniSection, g.strIniFile)
    
    'toolbar template combo box
    cboTemplate.Clear
    cboTemplate.AddItem "Custom"
    
    If aFiles.Size = 0 Then
        cboTemplate.Enabled = False
        lblTemplate.Enabled = False
    Else
        Set m.tbContent = New cGdTable
        Set m.aTbIndex = Nothing


        m.tbContent.CreateField eGDARRAY_Strings
        m.tbContent.CreateField eGDARRAY_Longs
        m.tbContent.CreateField eGDARRAY_Strings
        m.tbContent.CreateField eGDARRAY_Longs
        m.tbContent.CreateField eGDARRAY_Longs
        m.tbContent.CreateField eGDARRAY_Strings

        'read in first line of file to get display name, display order, enablement & iconsize
        For i = 0 To aFiles.Size - 1
            strFile = App.Path & "\provided\" & Parse(aFiles(i), ".sho", 1) & ".sho"
            aTemp.FromFile strFile
            
            strEnable = Parse(aTemp(0), vbTab, 4)
            
            If HasModule(strEnable) Then
                strName = Parse(aTemp(0), vbTab, 2)
                nOrder = Val(Parse(aTemp(0), vbTab, 3))
                nLargeIcon = Val(Parse(aTemp(0), vbTab, 5))
                nIncludeText = Val(Parse(aTemp(0), vbTab, 6))
                
                If nOrder <= 0 Then nOrder = 999999
            
                m.tbContent.AddRecord ""
                j = m.tbContent.NumRecords - 1
                
                m.tbContent(0, j) = strName
                m.tbContent(1, j) = nOrder
                m.tbContent(2, j) = strEnable
                m.tbContent(3, j) = nLargeIcon
                m.tbContent(4, j) = nIncludeText
                m.tbContent(5, j) = strFile
            End If
        Next
        
        Set m.aTbIndex = m.tbContent.CreateSortedIndex(1, eGdSort_Default Or eGdSort_Stable, 0, eGdSort_Default Or eGdSort_Stable)
        
        If m.aTbIndex Is Nothing Then
            cboTemplate.Enabled = False
            lblTemplate.Enabled = False
        ElseIf m.aTbIndex.Size <= 0 Then
            cboTemplate.Enabled = False
            lblTemplate.Enabled = False
        Else
            For i = 0 To m.aTbIndex.Size - 1
                strName = m.tbContent(0, m.aTbIndex(i))
                If Len(strName) > 0 Then
                    cboTemplate.AddItem strName
                    If UCase(m.strTemplate) = UCase(strName) Then iListIndex = cboTemplate.ListCount - 1
                End If
            Next
            cboTemplate.Enabled = True
            lblTemplate.Enabled = True
        End If
    End If
    
    FileCopy App.Path & "\toolbar.sho", App.Path & "\toolbarsho.sav", True
    
    cboTemplate.ListIndex = iListIndex
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.PopulateCboTemplate"

End Sub

Private Sub PopulateGrid()
On Error GoTo ErrSection:
    
    Dim nItem&
    Dim strItem$, strToolbar$, strText$, strIcon$
    
    Dim bTbVisible As Boolean
    Dim bShow As Boolean

    With fg
        .Redraw = flexRDNone
        .Row = 0
        For nItem = 0 To aItems.Size - 1
            strItem = aItems(nItem)
            If Left(strItem, 1) = "=" Then
                ' Toolbar
                strToolbar = Mid(strItem, 2)
                .RowHeight(.Row) = Screen.TwipsPerPixelY * 26
                .Col = 0
                
                If strToolbar = "General" Then
                    bTbVisible = m.bTbGeneral
                    m.nTbDrawGridRow = .Row
                ElseIf strToolbar = kTbWindows Then
                    bTbVisible = m.bTbWindows
                    m.nTbWindowsGridRow = .Row
                ElseIf strToolbar = kTbChartSettings Then
                    bTbVisible = m.bTbChart
                    m.nTbChartingGridRow = .Row
                ElseIf strToolbar = kTbDraw Then
                    bTbVisible = m.bTbDraw
                    m.nTbDrawGridRow = .Row
                Else
                    bTbVisible = frmMain.tbToolbar.ToolBars(strToolbar).Visible
                End If
                
                If 0 Then
                    .CellChecked = flexNoCheckbox
                    .CellAlignment = flexAlignLeftBottom
                ElseIf bTbVisible Then      'frmMain.tbToolbar.ToolBars(strToolbar).Visible Then
                    .CellChecked = flexChecked
                Else
                    .CellChecked = flexUnchecked
                End If
                .CellPictureAlignment = flexPicAlignLeftCenter
                .CellFontBold = True
                .TextMatrix(.Row, 0) = "" & strToolbar & " toolbar:         " '(need extra spaces at end to spill properly!)
                .TextMatrix(.Row, 3) = strToolbar
                .Row = .Row + 1
            ElseIf frmMain.tbToolbar.Tools(strItem).Visible Then
                ' Button on toolbar
                .RowHeight(.Row) = Screen.TwipsPerPixelY * 18
                .Col = 1
                .CellPictureAlignment = flexPicAlignCenterCenter
                strText = StripStr(frmMain.tbToolbar.Tools(strItem).Name, "&")
                strIcon = ToolbarIcon(strItem)
                Select Case strItem
                Case "ID_Daily"
                    strIcon = "Dly"
                    strText = "Daily"
                Case "ID_Weekly"
                    strIcon = "Wk"
                    strText = "Weekly"
                Case "ID_Monthly"
                    strIcon = "Mo"
                    strText = "Monthly"
                Case "ID_Quarterly"
                    strIcon = "Qtr"
                    strText = "Quarterly"
                Case "ID_Yearly"
                    strIcon = "Yr"
                    strText = "Yearly"
                Case "ID_CustomMinute"
                    strIcon = "??"
                    strText = "Custom # minutes"
                Case "ID_CustomPeriod"
                    strIcon = "Per"
                    strText = "Custom Bar Period"
                Case "ID_RealTime"
                    strText = "Data Streaming on/off"
                Case "ID_ChartData"
                    strText = "Chart Data window"
                Case "ID_EditChart"
                    strText = "Edit Chart Settings"
                Case "ID_LessAboveBelow"
                    strText = "Less Space above/below prices"
                Case "ID_MoreAboveBelow"
                    strText = "More Space above/below prices"
                Case "ID_Snapshot"
                    strText = "Snapshot (data & fundamentals)"
                Case "ID_Sectors"
                    strText = "Sectors"
                Case "ID_Subsectors"
                    strText = "Subsectors"
                Case "ID_Components"
                    strText = "Components"
                Case Else
                    If Right(strItem, 6) = "minute" Then
                        strIcon = Mid(Left(strItem, Len(strItem) - 6), 4)
                        strText = strIcon & " minute"
                    End If
                End Select
                If Left(strIcon, 1) = "k" Then
                    If g.nTbIconStyle = 1 Then
                        If g.nColorTheme = kDarkThemeColor Then
                            .CellPicture = g.CoreBridge.ImgListToolbarExt("Light", strIcon, "", 16).ExtractIcon
                        Else
                            .CellPicture = g.CoreBridge.ImgListToolbarExt("Dark", strIcon, "", 16).ExtractIcon
                        End If
                    Else
                        .CellPicture = g.CoreBridge.ImgListToolbarExt("Classic", strIcon, "", 16).ExtractIcon
                    End If
                Else
                    .TextMatrix(.Row, 1) = strIcon
                End If
                bShow = False
                On Error Resume Next
                bShow = frmMain.tbToolbar.ToolBars(strToolbar).Tools.Item(strItem).Visible
                .TextMatrix(.Row, 2) = strText
                .Col = 0
                .CellPictureAlignment = flexPicAlignRightCenter
                If bShow Then
                    .CellChecked = flexChecked
                Else
                    .CellChecked = flexUnchecked
                End If
                
                If strToolbar = "General" Then
                    bShow = m.bTbGeneral
                ElseIf strToolbar = kTbWindows Then
                    bShow = m.bTbWindows
                ElseIf strToolbar = kTbChartSettings Then
                    bShow = m.bTbChart
                ElseIf strToolbar = kTbDraw Then
                    bShow = m.bTbDraw
                Else
                    bShow = frmMain.tbToolbar.ToolBars(strToolbar).Visible
                End If

                .RowHidden(.Row) = Not bShow        'frmMain.tbToolbar.ToolBars(strToolbar).Visible
                .TextMatrix(.Row, 3) = strToolbar & "|" & strItem
                .Row = .Row + 1
            End If
        Next
        .Rows = .Row
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmToolbar.PopulateGrid"

End Sub

