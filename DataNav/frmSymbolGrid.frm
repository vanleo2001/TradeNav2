VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSymbolGrid 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Symbols"
   ClientHeight    =   5940
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   3555
   Icon            =   "frmSymbolGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdSettings 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   120
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
      Caption         =   "frmSymbolGrid.frx":038A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmSymbolGrid.frx":03BC
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGrid.frx":0444
      RightToLeft     =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid fgTree 
      Height          =   2895
      Left            =   60
      TabIndex        =   3
      Top             =   4560
      Width           =   2895
      _cx             =   5106
      _cy             =   5106
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
   Begin HexUniControls.ctlUniFrameWL fraFlags 
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   3720
      Width           =   3495
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
      Caption         =   "frmSymbolGrid.frx":0460
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolGrid.frx":048C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGrid.frx":04AC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSaveFlags 
         Height          =   300
         Left            =   1200
         TabIndex        =   8
         Top             =   300
         Width           =   675
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
         Caption         =   "frmSymbolGrid.frx":04C8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolGrid.frx":04F2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGrid.frx":057C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLoadFlags 
         Height          =   300
         Left            =   1920
         TabIndex        =   9
         Top             =   300
         Width           =   675
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
         Caption         =   "frmSymbolGrid.frx":0598
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolGrid.frx":05C2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGrid.frx":063E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClearAll 
         Height          =   300
         Left            =   480
         TabIndex        =   7
         Top             =   300
         Width           =   675
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
         Caption         =   "frmSymbolGrid.frx":065A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolGrid.frx":0686
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGrid.frx":06CA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFlags 
         Height          =   255
         Left            =   30
         Top             =   330
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
         Caption         =   "frmSymbolGrid.frx":06E6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolGrid.frx":0712
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGrid.frx":0732
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblNumFlagged 
         Height          =   255
         Left            =   30
         Top             =   45
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
         Caption         =   "frmSymbolGrid.frx":074E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolGrid.frx":07AE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGrid.frx":07CE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VB.Timer tmrSortCol 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1560
      Top             =   540
   End
   Begin MSComctlLib.ImageCombo cboList 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VSFlex7LCtl.VSFlexGrid fgVirtual 
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4335
      _cx             =   7646
      _cy             =   5318
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
   Begin VB.PictureBox Picture2 
      Height          =   135
      Left            =   300
      Picture         =   "frmSymbolGrid.frx":07EA
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   0
      Picture         =   "frmSymbolGrid.frx":08E0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.Menu mnuGrid 
      Caption         =   "Grid"
      Begin VB.Menu mnuLookup 
         Caption         =   "Lookup Symbol"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Select columns to display in grid"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarketInfo 
         Caption         =   "Market information for"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddToQuoteBoard 
         Caption         =   "Add selected symbols to Quote Board"
      End
      Begin VB.Menu mnuSymGroupAdd 
         Caption         =   "Add selected symbols to Symbol Group"
         Begin VB.Menu mnuSymGroup 
            Caption         =   "(add to new SymbolGroup)"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColumn 
         Caption         =   "COLUMN:"
      End
      Begin VB.Menu mnuAscending 
         Caption         =   "   Sort Ascending"
      End
      Begin VB.Menu mnuDescending 
         Caption         =   "   Sort Descending"
      End
      Begin VB.Menu mnuRemoveCol 
         Caption         =   "   Remove from grid"
      End
      Begin VB.Menu mnuEditObject 
         Caption         =   "   Edit"
      End
      Begin VB.Menu mnuSepColumn 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSymbol 
         Caption         =   "ROW:"
      End
      Begin VB.Menu mnuChart 
         Caption         =   "   New Chart"
      End
      Begin VB.Menu mnuSyncAll 
         Caption         =   "   Synchronize All Charts"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoChart 
         Caption         =   "Auto-Chart (chart/grid synchronization)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmSymbolGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSymbolGrid.frm
'' Description: Display symbols and information in a grid style format
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/27/2009   DAJ         Allow Save flags button to work for sector tree
'' 04/28/2009   DAJ         Fixed load flags and clear flags for sector tree
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDSymbolGridMode
    eGDSymbolGridMode_Virtual
    eGDSymbolGridMode_Tree
End Enum

Private Enum eGDTreeCols
    eGDTreeCol_TableIndex = 0
    eGDTreeCol_Flagged
    eGDTreeCol_Symbol
End Enum

Private Enum eTblCols
    eTblCol_Symbol = 0
    eTblCol_SymbolID
    eTblCol_PoolRec
    eTblCol_Description
    eTblCol_Level
    eTblCol_SortKey
    eTblCol_NumCols
End Enum

Private Type mPrivate
    WindowLink As New cWindowLink       ' Object to handle window linking

    SymbolGrid As cSymbolGrid           ' Object to handle the symbol grid information
    strLastSymbol As String             ' Last symbol selected
    nSymbolID As Long                   ' Symbol ID of the last symbol selected
    bAutoChart As Boolean               ' Change chart when user changes to a new symbol?
    bShowFlags As Boolean               ' Show flag column?

    tblSymbols As cGdTable              ' Table of symbols
    hSymbols As Long                    ' Handle to the table of symbols
    aSortedIndex As cGdArray            ' Sorted Index array on the table of symbols
    hSortedIndex As Long                ' Handle to the sorted index array
    alSectorPool As cGdArray            ' Array of sector pool information
    
    lNumSectors As Long                 ' Number of sectors
    Mode As eGDSymbolGridMode           ' Mode of the symbol grid (Regular vs Tree)
    lOpenRow As Long                    ' Currently open row in the tree
    lSortedCol As Long                  ' Sorted column in the grid
    bSortedDescending As Boolean        ' Sort Ascending or Descending?
End Type
Private m As mPrivate

Private Function GDTreeCol(ByVal Col As eGDTreeCols) As Long
    GDTreeCol = Col
End Function

Private Function TblCol(ByVal Col As eTblCols) As Long
    TblCol = Col
End Function

Private Property Get TableStr(ByVal nField As eTblCols, ByVal lRecord As Long) As String
    TableStr = gdGetTableString(m.hSymbols, nField, lRecord)
End Property
Private Property Let TableStr(ByVal nField As eTblCols, ByVal lRecord As Long, ByVal strValue As String)
    gdSetTableStr m.hSymbols, nField, lRecord, strValue
End Property
Private Property Get TableNum(ByVal nField As eTblCols, ByVal lRecord As Long) As Double
    TableNum = gdGetTableNum(m.hSymbols, nField, lRecord)
End Property
Private Property Let TableNum(ByVal nField As eTblCols, ByVal lRecord As Long, ByVal dValue As Double)
    gdSetTableNum m.hSymbols, nField, lRecord, dValue
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.nSymbolID
End Property
Public Property Let SymbolID(ByVal nSymbolID As Long)
On Error GoTo ErrSection:

    m.nSymbolID = nSymbolID
    ShowSymbol GetSymbol(nSymbolID)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmSymbolGrid.SymbolID.Let"

End Property

Public Property Get WindowLink() As cWindowLink
    Set WindowLink = m.WindowLink
End Property

Public Property Get SymbolGrid() As cSymbolGrid
    Set SymbolGrid = m.SymbolGrid
End Property

Public Property Get AutoChart() As Boolean
    AutoChart = m.bAutoChart
End Property

Public Property Let AutoChart(ByVal bAutoChart As Boolean)
On Error GoTo ErrSection:
    
    ' not used anymore
    ''m.bAutoChart = bAutoChart
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmSymbolGrid.AutoChart.Let"

End Property

Public Property Get SelectionKey() As String
    SelectionKey = cboList.SelectedItem.Key
End Property

Private Sub cboList_Change()
On Error GoTo ErrSection:

    ChangeList

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.cboList.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub cboList_Click()
On Error GoTo ErrSection:

    ChangeList

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.cboList.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboList_Dropdown()
On Error GoTo ErrSection:

    'reload list (to pick up any changes)
    'LoadCombo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.cboList.DropDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboList_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    ' certain characters (like "$") aren't getting picked up
    ' by the form's preview so they fall through to here
    KeyPress KeyAscii
    KeyAscii = 0
    MoveFocus fgVirtual

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.cboList.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClearAll_Click
'' Description: If the user clicks on the Clear All Button, clear all of the
''              flags in the symbol pool
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearAll_Click()
On Error GoTo ErrSection:

    m.SymbolGrid.ClearAllFlags
    
    If m.Mode = eGDSymbolGridMode_Tree Then
        RefreshTreeFlags
    End If
    
    UpdateFlagCount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.cmdClearAll_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLoadFlags_Click
'' Description: Allow the user to flag the symbols from a Symbol Group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLoadFlags_Click()
On Error GoTo ErrSection:

    Dim strID As String                 ' ID of the Selected Symbol Group
    Dim lFieldToLoad As Long            ' Field Number for the Selected Symbol Group
    Dim lFlagField As Long              ' Field Number for the Flag Group
    
    ' Have the user select the Symbol Group to Load From
    strID = frmSelect.ShowMe("GRP", eSelectMode_Select)
    
    If strID <> "" Then
        ' Get the pool fields for the flagged group and the group to load
        lFlagField = CLng(Val(fgVirtual.ColData(kFlagCol)))
        lFieldToLoad = g.SymbolPool.FieldNumForID("GRP:" & strID)
        
        If lFlagField <> -1 And lFieldToLoad <> -1 Then
            fgVirtual.Redraw = flexRDNone
            With g.SymbolPool.ArrayTable
                .AttachField .FieldArray(lFieldToLoad), lFlagField
            End With
            UpdateFlagCount
            fgVirtual.Redraw = flexRDBuffered
        End If
        
        If m.Mode = eGDSymbolGridMode_Tree Then
            RefreshTreeFlags
        End If
    End If
    
    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            MoveFocus fgVirtual
        Case eGDSymbolGridMode_Tree
            MoveFocus fgTree
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.cmdLoadFlags_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSaveFlags_Click
'' Description: Allow the user to save the flagged symbols to a Symbol Group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSaveFlags_Click()
On Error GoTo ErrSection:

    Dim alSymbolIds As New cGdArray     ' Array of symbol ids to send to a group
    Dim lIndex As Long                  ' Index into an array
    Dim lCounter As Long                ' Counter into a for loop
    Dim lField As Long                  ' Symbol pool field for the flagged array
    Dim alTemp As New cGdArray          ' Temporary array
    Dim lFilterFld As Long              ' field # of filter
           
    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            ' Get the field number for the flagged group
            lField = CLng(Val(fgVirtual.ColData(kFlagCol)))
            
            With g.SymbolPool
                ' create the filtered list of flagged symbols
                lFilterFld = .FieldNumForID(cboList.SelectedItem.Key)
                alTemp.Create eGDARRAY_TinyInts, .NumRecords
                alTemp.ArrayOperate .ArrayTable.FieldArray(lField), "AND", .ArrayTable.FieldArray(lFilterFld)
                
                ' Create the Symbol ID array
                alSymbolIds.Create eGDARRAY_Longs, alTemp.CountOf(1)
                lIndex = 0&
                
                ' If a symbol is flagged, add the Symbol ID to the array
                For lCounter = 0 To .NumRecords - 1
                    If alTemp(lCounter) = 1 Then
                        alSymbolIds(lIndex) = .SymbolID(lCounter)
                        lIndex = lIndex + 1
                    End If
                Next lCounter
            End With
                    
            ' Allow the user to select the group to add the flagged symbols to
            frmSelect.ShowMe "GRP", eSelectMode_SendTo, alSymbolIds
            
            MoveFocus fgVirtual
        
        Case eGDSymbolGridMode_Tree
            lField = CLng(Val(fgTree.ColData(GDTreeCol(eGDTreeCol_Flagged))))
            
            With g.SymbolPool
                alTemp.Create eGDARRAY_TinyInts, .NumRecords
                alTemp.ArrayOperate .ArrayTable.FieldArray(lField), "AND", m.alSectorPool
                
                ' Create the Symbol ID array
                alSymbolIds.Create eGDARRAY_Longs, alTemp.CountOf(1)
                lIndex = 0&
                
                ' If a symbol is flagged, add the Symbol ID to the array
                For lCounter = 0 To .NumRecords - 1
                    If alTemp(lCounter) = 1 Then
                        alSymbolIds(lIndex) = .SymbolID(lCounter)
                        lIndex = lIndex + 1
                    End If
                Next lCounter
            End With
                                
            ' Allow the user to select the group to add the flagged symbols to
            frmSelect.ShowMe "GRP", eSelectMode_SendTo, alSymbolIds
            
            MoveFocus fgTree
        
    End Select
    
ErrExit:
    Set alSymbolIds = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.cmdSaveFlags_Click"

End Sub

Private Sub cmdShow_Click()
On Error GoTo ErrSection:

    Dim astrAvailable As New cGdArray
    Dim astrUsed As New cGdArray
    Dim astrUsedSorted As New cGdArray
    Dim aFields As New cGdArray
    Dim lIndex As Long
    Dim nField&
    Dim strDisplayFields As String
    Dim strID As String
    Dim strName$, strIgnore$
    Dim bReturn As Boolean
    Dim bScansOn As Boolean
    Dim bSkip As Boolean
    Dim nVirtualRow&, nVirtualTopRow&
    Dim strTreeSymbol As String
    Dim obj As Object

    strIgnore = "||FLAGS|FLAGGED SYMBOLS|SYMINDEX|SYMBOL|DBRECNUM|ALL SYMBOLS|"
    
    bScansOn = ScansEnabled
    
    astrAvailable.Create eGDARRAY_Strings
    astrUsed.Create eGDARRAY_Strings
    astrUsedSorted.Create eGDARRAY_Strings

    ' Set up the available/used arrays
    strDisplayFields = GetGridFields
    aFields.SplitFields strDisplayFields, "|"
    For lIndex = 0 To aFields.Size - 1
        strID = Parse(aFields(lIndex), "\", 1)
        nField = g.SymbolPool.FieldNumForID(strID)
        If nField >= 0 Then
            strName = g.SymbolPool.ArrayTable.FieldName(nField)
            'skip certain fields
            If InStr(strIgnore, "|" & UCase(strName) & "|") = 0 Then
                astrUsed.Add strName
                astrUsedSorted.Add strName
            End If
        End If
    Next
    astrUsedSorted.Sort
    
    For lIndex = 0 To g.SymbolPool.ArrayTable.NumFields - 1
        strID = g.SymbolPool.FieldID(lIndex)
        If Len(strID) > 0 And Left(strID, 4) <> "DSP:" Then
            Set obj = g.SymbolPool.PoolObject(strID)
            
            If bScansOn = False And (Left(strID, 4) = "DSV:" Or Left(strID, 4) = "FIL:") Then
                bSkip = True
            Else
                bSkip = False
            End If
            
            If Not obj Is Nothing Then
                If bSkip = False Then bSkip = (obj.IsActive <> True)
            End If
            
            strName = g.SymbolPool.ArrayTable.FieldName(lIndex)
            'skip certain fields
            If InStr(strIgnore, "|" & UCase(strName) & "|") = 0 Then
                If astrUsedSorted.BinarySearch(strName) = False And bSkip = False Then
                    astrAvailable.Add strName
                End If
            End If
        End If
    Next lIndex
    astrAvailable.Sort eGdSort_IgnoreCase

    ' Call the add/remove form
    bReturn = frmAddRemove.ShowMe(astrAvailable, astrUsed, eOrderMode_Ordered, , "Arrange Columns of Grid")
    
    ' Make the new "Used" string and re-init grid if add/remove returned OK
    If bReturn = True Then
        'keep Symbol as first column
        strDisplayFields = "INF:SYMBOL"
        For lIndex = 0 To astrUsed.Size - 1
            nField = g.SymbolPool.ArrayTable.FieldNum(astrUsed(lIndex))
            strID = g.SymbolPool.FieldID(nField)
            If Len(strID) > 0 Then
                strDisplayFields = strDisplayFields & "|" & strID
            End If
        Next lIndex
        
        nVirtualRow = fgVirtual.Row
        nVirtualTopRow = fgVirtual.TopRow
        strTreeSymbol = fgTree.TextMatrix(fgTree.Row, GDTreeCol(eGDTreeCol_Symbol))
        
        m.SymbolGrid.InitGrid fgVirtual, strDisplayFields
        InitTreeGrid strDisplayFields
        LoadTreeGrid
        
        On Error Resume Next
        fgVirtual.TopRow = nVirtualTopRow
        fgVirtual.Row = nVirtualRow
        ShowTreeSymbol strTreeSymbol
        
        If m.Mode = eGDSymbolGridMode_Virtual Then
            MoveFocus fgVirtual
        Else
            MoveFocus fgTree
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.cmdShow.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdSettings_Click()
On Error GoTo ErrSection:

    Dim strGridFont As String           ' Font from the grid
    Dim strDisplayFields As String      ' Fields to display in the grid
    Dim strRankField As String          ' Field to sort on
    Dim bAscending As Boolean           ' Is the field sorted ascending?
    Dim bShowFlags As Boolean           ' Should we show the flag stuff?
    Dim bList As Boolean                ' Are we displaying the list style?
    Dim strFilterID As String           ' ID of the filter from the combo
    Dim lIndex As Long                  ' Index into a for loop
    Dim lTreeCol As Long                ' Column in the tree to sort
    Dim lVirtualCol As Long             ' Column in the virtual grid to sort
    
    ' Get the font information...
    strGridFont = FontToString(fgVirtual.Font)
    
    ' Get the field information...
    strDisplayFields = GetGridFields
    strRankField = Trim(fgTree.TextMatrix(0, m.lSortedCol))
    bAscending = Not m.bSortedDescending
    bShowFlags = m.bShowFlags
    
    ' Get the view information...
    strFilterID = cboList.SelectedItem.Key
    bList = (strFilterID <> "Sectors")
    
    If frmSymbolGridCfg.ShowMe(strGridFont, strDisplayFields, strRankField, bAscending, bShowFlags, bList, strFilterID) = True Then
        ' Change the font information...
        FontFromString fgVirtual.Font, strGridFont
        fgVirtual.Font = fgVirtual.Font
        FontFromString fgTree.Font, strGridFont
        fgTree.Font = fgTree.Font
        
        ' Change the field information...
        ChangeFields strDisplayFields
        For lIndex = 0 To fgTree.Cols - 1
            If Trim(fgTree.TextMatrix(0, lIndex)) = strRankField Then
                lTreeCol = lIndex
                Exit For
            End If
        Next lIndex
        For lIndex = 0 To fgVirtual.Cols - 1
            If Trim(fgVirtual.TextMatrix(0, lIndex)) = strRankField Then
                lVirtualCol = lIndex
                Exit For
            End If
        Next lIndex
        If bAscending = True Then
            SortTreeOnCol lTreeCol, 1
            m.SymbolGrid.SortOnCol lVirtualCol, 1
        Else
            SortTreeOnCol lTreeCol, -1
            m.SymbolGrid.SortOnCol lVirtualCol, -1
        End If
        m.bShowFlags = bShowFlags
        ShowFlags
        
        ' Change the view information...
        If bList Then
            cboList.ComboItems(strFilterID).Selected = True
        Else
            cboList.ComboItems("Sectors").Selected = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.cmdSettings.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgTree_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
On Error GoTo ErrSection:

    SetBackColors fgTree
    If Me.Visible And Row <> -1& Then
        ShowTreeData Row
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgTree.AfterCollapse", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgTree_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lTblIndex As Long               ' Index into the table
    Dim lPoolField As Long              ' Flag field in the pool
    Dim lPoolRec As Long                ' Record number in the symbol pool

    If Col = GDTreeCol(eGDTreeCol_Flagged) Then
        lPoolField = CLng(Val(Parse(fgTree.ColData(Col), vbTab, 1)))
        lTblIndex = m.aSortedIndex(CLng(fgTree.TextMatrix(Row, GDTreeCol(eGDTreeCol_TableIndex))))
        lPoolRec = m.tblSymbols(TblCol(eTblCol_PoolRec), lTblIndex)
        If CheckedCell(fgTree, Row, Col) = True Then
            g.SymbolPool.ArrayTable(lPoolField, lPoolRec) = 1
        Else
            g.SymbolPool.ArrayTable(lPoolField, lPoolRec) = 0
        End If
        
        ' Need to call this to enable/disable the buttons and update stuff...
        UpdateFlagCount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgTree.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgTree_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    UpdateNumberLabel
    
    KeyPress 27 'to restore caption
    If Me.Visible And fgTree.RowSel > 0 And NewRow <> OldRow And Not g.bStarting Then
        If fgTree.Redraw <> flexRDNone Then
            m.strLastSymbol = Trim(fgTree.TextMatrix(NewRow, GDTreeCol(eGDTreeCol_Symbol)))
            m.nSymbolID = GetSymbolID(m.strLastSymbol)
            If AutoChart Then
                SetActiveChartSymbol Trim(fgTree.TextMatrix(fgTree.RowSel, GDTreeCol(eGDTreeCol_Symbol)))
            End If
        End If
    ElseIf Me.Visible Then
        UpdateFlagCount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgTree.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgTree_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lParentRow As Long              ' Parent row of the current row
    Dim lIndex As Long
    Dim strSymbol As String
    Dim lNewRow As Long

    If Row >= fgTree.FixedRows And Row < fgTree.Rows Then
        fgTree.Row = Row
        fgTree.RowSel = Row
    End If

    If State = flexOutlineExpanded Then
        With fgTree
            lRedraw = .Redraw
            .Redraw = flexRDNone
            
            strSymbol = .TextMatrix(Row, GDTreeCol(eGDTreeCol_Symbol))

            ' Collapse all branches except for this one...
            'If m.lOpenRow <> -1& Then .IsCollapsed(m.lOpenRow) = flexOutlineCollapsed
            .Outline 1
            
            lParentRow = .GetNodeRow(Row, flexNTParent)
            If lParentRow <> -1 Then .IsCollapsed(lParentRow) = flexOutlineExpanded
                           
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, GDTreeCol(eGDTreeCol_Symbol)) = strSymbol Then
                    lNewRow = lIndex
                    Exit For
                End If
            Next lIndex
            
            If lNewRow <> Row Then
                .IsCollapsed(lNewRow) = flexOutlineExpanded
                .Row = lNewRow
                .RowSel = lNewRow
                Cancel = True
            Else
                ' Fill in the leaves if we need to...
                If .TextMatrix(lNewRow + 1, GDTreeCol(eGDTreeCol_Symbol)) = "(blank)" Then
                    ExpandRow lNewRow
                End If
                
                ' If the bottom node is off the screen, make this the top row...
                If .GetNodeRow(lNewRow, flexNTLastChild) > .BottomRow Then
                    .TopRow = .Row
                End If
            End If
            
            .Redraw = lRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgTree.BeforeCollapse", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgTree_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> GDTreeCol(eGDTreeCol_Flagged) Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgTree.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgTree_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim lMouseCol As Long               ' Current mouse column in the grid
        
    lMouseRow = fgTree.MouseRow
    lMouseCol = fgTree.MouseCol
    
    If Button = vbRightButton Then
        Cancel = True
        fgTree.Col = lMouseCol
        ShowGridPopup lMouseRow, lMouseCol
    ElseIf lMouseRow = 0 Then
        'trigger sort col on mouse up
        tmrSortCol.Enabled = True
    ElseIf lMouseCol = 0 Then
        ' 11/26/02: handle toggling the flag box here so can cancel the Row move
        ' (we no longer want to move the selected row when flag box is toggled)
        'Cancel = True
        'm.SymbolGrid.ToggleFlags lMouseRow
        'UpdateFlagCount
    ElseIf lMouseRow > 0 And lMouseRow < fgTree.Rows And fgTree.Row <> lMouseRow Then
        'NO: this messes up the multi-select!
        'fgTree.Row = lMouseRow
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgTree.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgTree_DblClick()
On Error GoTo ErrSection:
    
    If fgTree.MouseRow > 0 And Not AutoChart Then
        SetActiveChartSymbol Trim(fgTree.TextMatrix(fgTree.RowSel, GDTreeCol(eGDTreeCol_Symbol)))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgTree.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgTree_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If fgKeyDown(KeyCode, Shift) Then Exit Sub
       
    With fgTree
        Select Case KeyCode
            Case vbKeyRight
                KeyCode = 0
                If .RowOutlineLevel(.Row) < 3 Then
                    If .IsCollapsed(.Row) = flexOutlineExpanded Then
                        If .RowOutlineLevel(.Row + 1) >= .RowOutlineLevel(.Row) Then
                            'm.bChangeChart = True
                            .Row = .Row + 1
                            .ShowCell .Row, GDTreeCol(eGDTreeCol_Symbol)
                        End If
                    Else
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    End If
                ElseIf .Row + 1 < .Rows Then
                    If .RowOutlineLevel(.Row + 1) >= .RowOutlineLevel(.Row) Then
                        'm.bChangeChart = True
                        .Row = .Row + 1
                        .ShowCell .Row, GDTreeCol(eGDTreeCol_Symbol)
                    End If
                End If
            
            Case vbKeyLeft
                KeyCode = 0
                If .Row >= .FixedRows And .Row < .Rows Then
                    If .RowOutlineLevel(.Row) = 3 Then
                        If .GetNodeRow(.Row, flexNTParent) <> -1 Then
                            'm.bChangeChart = True
                            .Row = .GetNodeRow(.Row, flexNTParent)
                            .ShowCell .Row, GDTreeCol(eGDTreeCol_Symbol)
                        End If
                    ElseIf .IsCollapsed(.Row) = flexOutlineExpanded Then
                        .IsCollapsed(.Row) = flexOutlineCollapsed
                    ElseIf .GetNodeRow(.Row, flexNTParent) <> -1 Then
                        'm.bChangeChart = True
                        .Row = .GetNodeRow(.Row, flexNTParent)
                        .ShowCell .Row, GDTreeCol(eGDTreeCol_Symbol)
                    End If
                End If
            
        End Select
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgSymbols.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgTree_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgTree.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgTree_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgTree

End Sub

Private Sub fgVirtual_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:
    
    fgVirtual.FlexDataSource = m.SymbolGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgVirtual.AfterMoveColumn", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgVirtual_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim nSymbolID&
    
    UpdateNumberLabel
    
    KeyPress 27 'to restore caption
    If Me.Visible And fgVirtual.RowSel > 0 And NewRow <> OldRow And Not g.bStarting Then
        If fgVirtual.Redraw <> flexRDNone Then
            m.strLastSymbol = Trim(fgVirtual.TextMatrix(NewRow, kSymbolCol))
            m.nSymbolID = GetSymbolID(m.strLastSymbol)
            frmMain.SetWindowLink Me
            If AutoChart Then
                SetActiveChartSymbol Trim(fgVirtual.TextMatrix(fgVirtual.RowSel, kSymbolCol))
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgVirtual.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgVirtual_AfterUserFreeze()
On Error GoTo ErrSection:

    ' Make sure that the symbol and flag columns remain frozen
    If fgVirtual.FrozenCols < 2 Then fgVirtual.FrozenCols = 2

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgVirtual.AfterUserFreeze", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgVirtual_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

'11/26/02: this is now handled in the BeforeMouseDown event
#If 0 Then
    ' Only allow editing on the flag column
    With fgVirtual
        If Row < .FixedRows Or Row > .Rows - 1 Or Col <> kFlagCol Then
            Cancel = True
        ElseIf Row > 0 And Row < fgVirtual.Rows And fgVirtual.Row <> Row Then
            If fgVirtual.SelectedRows <= 1 Then
                fgVirtual.Row = Row
            End If
        End If
    End With
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgVirtual.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgVirtual_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim iMouseRow&, iMouseCol&
        
    iMouseRow = fgVirtual.MouseRow
    iMouseCol = fgVirtual.MouseCol
    
    If Button = 2 Then
        Cancel = True
        ShowGridPopup iMouseRow, fgVirtual.MouseCol
    ElseIf iMouseRow = 0 Then
        'trigger sort col on mouse up
        tmrSortCol.Enabled = True
    ElseIf fgVirtual.MouseCol = 0 Then
        ' 11/26/02: handle toggling the flag box here so can cancel the Row move
        ' (we no longer want to move the selected row when flag box is toggled)
        Cancel = True
        m.SymbolGrid.ToggleFlags iMouseRow
        UpdateFlagCount
    ElseIf iMouseRow > 0 And iMouseRow < fgVirtual.Rows And fgVirtual.Row <> iMouseRow Then
        'NO: this messes up the multi-select!
        ''fgVirtual.Row = iMouseRow
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgVirtual.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgVirtual_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:
    
    ' Keep the symbol and flag columns where they are
    If Col = kSymbolCol Then Position = kSymbolCol
    If Col = kFlagCol Then Position = kFlagCol
    If Col > kSymbolCol And Position <= kSymbolCol Then Position = Col
    
    tmrSortCol.Enabled = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgVirtual.BeforeMoveColumn", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgVirtual_BeforeScrollTip(ByVal Row As Long)
On Error GoTo ErrSection:
    
    fgVirtual.ScrollTipText = CStr(Row - fgVirtual.FixedRows) & " of " & CStr(fgVirtual.Rows - fgVirtual.FixedRows)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgVirtual.BeforeScrollTip", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgVirtual_DblClick()
On Error GoTo ErrSection:
    
    If fgVirtual.MouseRow > 0 And Not AutoChart Then
        SetActiveChartSymbol Trim(fgVirtual.TextMatrix(fgVirtual.RowSel, kSymbolCol))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgVirtual.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgVirtual_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If fgKeyDown(KeyCode, Shift) Then Exit Sub
       
    If KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Then
        If KeyCode = vbKeyHome Then
            fgVirtual.Row = 1
        Else
            fgVirtual.Row = fgVirtual.Rows - 1
        End If
        KeyCode = 0
        fgVirtual.ShowCell fgVirtual.Row, kSymbolCol
    End If

End Sub

Private Sub fgVirtual_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.fgVirtual.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgVirtual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgVirtual
    
End Sub

Private Sub Form_Activate()
On Error GoTo ErrSection:
        
    Static bAlreadyDone As Boolean
    Static bInProgress As Boolean

    '2/14/02: don't think this is needed now that StingRay is gone
    'm.nRowWhenActivated = fgVirtual.MouseRow
      
    m.WindowLink.Init Me
      
    If cboList.ComboItems.Count = 0 Then
        If bInProgress = False And Me.Visible = True Then
            bInProgress = True
            InfBox "There are no symbols to display", "i", , "Symbol Grid"
            frmMain.tbToolbar.Tools("ID_SymbolGrid").State = ssUnchecked
            bInProgress = False
        End If
    Else
        If Not bAlreadyDone Then
            bAlreadyDone = True
        
'            Me.Refresh
            If AutoChart Then
'                m.SymbolGrid.ShowChart fgVirtual.RowSel
            End If
        End If
    
        KeyPress 27 'to restore caption
    
        'cboList.ToolTipText = TipStr(CStr(fgVirtual.Rows - fgVirtual.FixedRows) & " symbols")
        SetComboTooltip
        
        ToolbarSync Me
    End If
    
    fgVirtual.BackColorAlternate = ALT_GRID_ROW_COLOR

    TextIncDecRegisterForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.Form.Activate", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:
    
    TextIncDecUnregisterForm Me
    
    tmrSortCol.Enabled = False
    SetPrevActiveForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.Form.Deactivate", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        g.Help.ShowF1Help Me
    ElseIf KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
        KeyPress KeyCode, Shift
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii
    MoveFocus fgVirtual

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.Form.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
  
    Dim strFont$, strKey$
    
    g.Styler.StyleForm Me
                
    mnuGrid.Visible = False
    cmdSettings.Height = cboList.Height
          
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    m.bAutoChart = False ' GetRegistryValue(rkLocalMachine, strKey, "AutoSync", True)
          
    Me.Width = fgVirtual.Width + fgVirtual.Left * 2
    'CenterTheForm Me
    Me.Top = 0
    Me.Left = 0
    If frmMain.ScaleHeight > 0 Then
        Me.Height = frmMain.ScaleHeight
    End If

    strFont = GetIniFileProperty("SymbolGrid", "", "Fonts", g.strIniFile)
    If strFont <> "" Then
        FontFromString fgVirtual.Font, strFont
        FontFromString fgTree.Font, strFont
    End If
    
    With fgTree
        .Move fgVirtual.Left, fgVirtual.Top, fgVirtual.Width, fgVirtual.Height
        .Visible = False
    End With
    
    'JM 12-28-2015: make sure plus/minus picture boxes do not show
    Picture1.Left = cmdSettings.Left + 30
    Picture1.Top = cmdSettings.Top + 30
    Picture2.Left = Picture1.Left
    Picture2.Top = Picture1.Top
    ''m.WindowLink.SymbolColor = GetIniFileProperty(Me.Name, 0&, "SymbolLink", g.strIniFile)
    
    InitForm

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    KeyPress 27 'to restore caption

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.Form.MouseUp", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    TextIncDecUnregisterForm Me

    If UnloadMode = 0 Then
        'Cancel = True
        'ToolbarSync Me, False
        ''Me.Hide
        'frmMain.DockPro.State(Me.Name) = DPHidden
        'AutoSizeChart
        frmMain.tbToolbar.Tools("ID_SymbolGrid").State = ssUnchecked
    End If
    
    If Cancel = 0 Then m.WindowLink.Unhook

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim h&, w&
    Static bInProgress As Boolean
    If bInProgress Then Exit Sub
    bInProgress = True
    
    With fgVirtual
        ''If FormX1.FrameType = otxDocked Then
        If DockState(Me) = eDocked Then
            w = Me.ScaleWidth - .Left
        Else
            w = Me.ScaleWidth - .Left * 2
        End If
        If m.bShowFlags Then
            h = Me.ScaleHeight - .Top - .Left - fraFlags.Height
        Else
            h = Me.ScaleHeight - .Top - .Left
        End If
        .Move .Left, .Top, w, h
    End With
    
    fgTree.Move fgVirtual.Left, fgVirtual.Top, fgVirtual.Width, fgVirtual.Height
    
    With fraFlags
        .Move cmdSettings.Left, Me.ScaleHeight - .Height - fgVirtual.Left
    End With
       
    'make dropdown end where grid does
    w = fgVirtual.Width - (cboList.Left - fgVirtual.Left)
    cboList.Width = w
    
    
    AutoSizeChart
   
    bInProgress = False

End Sub

Public Sub LoadCombo()
On Error Resume Next

    Dim i&, iMatch&, strID$, strType$, strPicture$, strKey$, nAt&, nNode&
    Dim strSelID$, bSelExists As Boolean
    Dim iSortStart&, strItem$, bScans As Boolean
    Dim aItems As New cGdArray
    Dim obj As Object
   
    bScans = ScansEnabled
        
    If cboList.ComboItems.Count > 0 Then
        ' save item currently set to
        strSelID = cboList.SelectedItem.Key
        cboList.ComboItems.Clear
    Else
        ' get item last set to (stored in INI file)
        strSelID = GetIniFileProperty("ComboID", "", "Grid", g.strIniFile)
    End If
    
    ' get list of items to put into combo list
    With g.SymbolPool
        For i = 0 To .ArrayTable.NumFields - 1
            strID = .FieldID(i)
            If Len(strID) = 0 Then
                strType = "" '???
            Else
                strType = Left(strID, 3)
                strPicture = ""
                Set obj = .PoolObject(strID)
                Select Case strType
                    Case "GRP":
                        strPicture = ToolbarIcon("ID_SymbolGroups")
                        If Len(strSelID) = 0 Then
                            If UCase(obj.Name) = "HUME" Then
                                strSelID = .FieldID(i)
                            End If
                        End If
                    Case "FIL":
                        If bScans Then
                            strPicture = ToolbarIcon("ID_Filters")
                        End If
                    Case "DSV":
                        If bScans Then
                            'only if boolean
                            strKey = Mid(strID, 5)
                            nNode = .Criterias.Index(strKey)
                            If nNode > 0 Then
                                If .Criterias(nNode).IsBoolean Then
                                    strPicture = ToolbarIcon("ID_Criteria")
                                End If
                            End If
                        End If
                End Select
                If Len(strPicture) > 0 Then
                    If obj.IsActive = True Then
                        If strID = strSelID Then
                            bSelExists = True
                        End If
                        
                        If iSortStart = 0 And i >= g.SymbolPool.OtherFieldsStart Then
                            iSortStart = aItems.Size
                        End If
                        
                        ' keep "flagged symbols" above where we sort
                        nAt = -1
                        If strID = "GRP:_FLAGS_.GRP" And iSortStart > 0 Then
                            nAt = iSortStart
                            iSortStart = iSortStart + 1
                        End If
                        
                        aItems.Add .ArrayTable.FieldName(i) & vbTab _
                                & strID & vbTab & strPicture, nAt
                    End If
                End If
            End If
        Next
    End With
    
    If aItems.Size > 0 Then
        aItems.ToFile App.Path & "\..\OptionNav\GroupNames.txt"
        aItems.Add "Stock SECTOR Tree" & vbTab & "Sectors" & vbTab & "kSectors", 4
        iSortStart = iSortStart + 1
        If strSelID = "Sectors" Then bSelExists = True
    End If
    
    If iSortStart > 0 Then
        aItems.Sort eGdSort_IgnoreCase, iSortStart
    End If

    For i = 0 To aItems.Size - 1
        strItem = aItems(i)
        cboList.ComboItems.Add , Parse(strItem, vbTab, 2), _
            Parse(strItem, vbTab, 1), Parse(strItem, vbTab, 3)
    Next

    If bSelExists Then
        cboList.ComboItems(strSelID).Selected = True
    Else
        cboList.ComboItems(1).Selected = True
    End If

    cboList.Refresh

End Sub

Private Sub OrderList()
On Error GoTo ErrSection:

    Dim n%, bDescending As Boolean
    
    n = cboList.SelectedItem.Index
    Select Case n
        Case 1:
            'm.SymbolGrid.SetList -99, 0, False
        Case 2:
            'm.SymbolGrid.SetList -99, 0, True
        Case 3:
            'm.SymbolGrid.SetList -99, -3, False
        Case 4:
            'm.SymbolGrid.SetList -99, -3, True
        Case 5:
            'm.SymbolGrid.SetList -99, -2, False
        Case 6:
            'm.SymbolGrid.SetList -99, -2, True
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.OrderList", eGDRaiseError_Raise

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    Dim strDisplayFields As String      ' Pipe delimited string of field nums from grid
    Dim lIndex As Long                  ' Index for a for loop
            
    ' store window link color and unhook the window proc
    ''SetIniFileProperty Me.Name, m.WindowLink.SymbolColor, "SymbolLink", g.strIniFile
    Set m.WindowLink = Nothing
            
    ToolbarSync Me, False
    
    ' Get the fields that are displayed in the grid
    strDisplayFields = GetGridFields
    
    ' Save the display field string in the INI file
    SetIniFileProperty "DisplayFields", strDisplayFields, "Grid", g.strIniFile
    If m.Mode = eGDSymbolGridMode_Tree Then
        SetIniFileProperty "InitialSymbol", fgTree.TextMatrix(fgTree.Row, GDTreeCol(eGDTreeCol_Symbol)), "Grid", g.strIniFile
    Else
        SetIniFileProperty "InitialSymbol", fgVirtual.TextMatrix(fgVirtual.Row, kSymbolCol), "Grid", g.strIniFile
    End If
    SetIniFileProperty "ComboID", cboList.SelectedItem.Key, "Grid", g.strIniFile
    SetIniFileProperty "SymbolGrid", FontToString(fgVirtual.Font), "Fonts", g.strIniFile
    SetIniFileProperty "ShowFlags", m.bShowFlags, "Grid", g.strIniFile
    
    Set m.SymbolGrid = Nothing
    Set m.tblSymbols = Nothing
    Set m.aSortedIndex = Nothing
    
    frmMain.DockPro.RemoveForm Me.Name

End Sub

Public Sub KeyPress(KeyAscii As Integer, Optional Shift As Integer = -1)
On Error Resume Next

    Dim strSymbol$, iRow&
    Dim astrSymbols As New cGdArray
    Dim frm As Form
    Dim bLookForChart As Boolean

    If KeyAscii = 0 Then Exit Sub

    If Shift >= 0 Then ' (came from KeyDown event)
        If KeyAscii >= vbKeyF2 And KeyAscii <= vbKeyF12 Then
            bLookForChart = True
        End If
    Else ' (came from KeyPress event)
        Select Case Asc(UCase(Chr(KeyAscii)))
            Case 13:        ' Enter Key
                If Not AutoChart Then
                    Select Case m.Mode
                        Case eGDSymbolGridMode_Virtual
                            SetActiveChartSymbol Trim(fgVirtual.TextMatrix(fgVirtual.RowSel, kSymbolCol))
                        
                        Case eGDSymbolGridMode_Tree
                            SetActiveChartSymbol Trim(fgTree.TextMatrix(fgTree.RowSel, GDTreeCol(eGDTreeCol_Symbol)))
                    
                    End Select
                End If
                KeyAscii = 0
            
            Case 32:        ' Space
                m.SymbolGrid.ToggleFlags
                UpdateFlagCount
                KeyAscii = 0
            
            Case 83:        ' S
                If AutoChart And Not ActiveChart Is Nothing Then
                    bLookForChart = True
                Else
                    Select Case m.Mode
                        Case eGDSymbolGridMode_Virtual
                            iRow = fgVirtual.RowSel
                            If iRow > fgVirtual.FixedRows And iRow < fgVirtual.Rows Then
                                strSymbol = fgVirtual.TextMatrix(iRow, kSymbolCol)
                            Else
                                strSymbol = ""
                            End If
                            Set astrSymbols = frmSymbolSelector.ShowMe(strSymbol, False)
                            If astrSymbols.Size > 0 Then
                                m.SymbolGrid.ShowRec g.SymbolPool.PoolRecForSymbol(astrSymbols(0))
                            End If
                            
                        Case eGDSymbolGridMode_Tree
                            iRow = fgTree.RowSel
                            If iRow > fgTree.FixedRows And iRow < fgTree.Rows Then
                                strSymbol = fgTree.TextMatrix(iRow, GDTreeCol(eGDTreeCol_Symbol))
                            Else
                                strSymbol = ""
                            End If
                            Set astrSymbols = frmSymbolSelector.ShowMe(strSymbol, False)
                            If astrSymbols.Size > 0 Then
                                ShowTreeSymbol astrSymbols(0)
                            End If
                            
                    End Select
                    
                    KeyAscii = 0
                End If
                
            Case 65 To 90, 48 To 57, 43, 45, 61:
                bLookForChart = True
        End Select
    End If
       
    If bLookForChart Then
        Set frm = ActiveChart
        If Not frm Is Nothing Then
            'MoveFocus frm
            'DoEvents
            frm.KeyPress KeyAscii, Shift
        End If
        KeyAscii = 0
    End If
       
    Set frm = Nothing
    Set astrSymbols = Nothing

End Sub

Private Function GetGridFields() As String
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim strReturn$, strField$
    
    For lIndex = 0 To fgVirtual.Cols - 1
        strField = fgVirtual.ColData(lIndex)
        strReturn = strReturn & Parse(strField, vbTab, 2) & "|"
    Next
    
    GetGridFields = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSymbolGrid.GetGridFields", eGDRaiseError_Raise

End Function

Private Sub lblFlags_Click()
On Error GoTo ErrSection:

    Static bTree As Boolean
    
    If bTree Then
        fgTree.Visible = False
        fgVirtual.Visible = True
        bTree = False
        m.Mode = eGDSymbolGridMode_Virtual
        MoveFocus fgVirtual
    Else
        RefreshTreeFlags
        fgVirtual.Visible = False
        fgTree.Visible = True
        bTree = True
        m.Mode = eGDSymbolGridMode_Tree
        MoveFocus fgTree
    End If
    
    UpdateFlagCount
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.lblFlags.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuAddToQuoteBoard_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRecNum As Long                 ' Symbol pool record number for symbol to add
    
    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            With fgVirtual
                For lIndex = 0 To .SelectedRows - 1
                    lRecNum = g.SymbolPool.PoolRecForSymbol(.TextMatrix(.SelectedRow(lIndex), kSymbolCol))
                    frmQuotes.AddSymbol lRecNum, "Daily"
                Next lIndex
            End With
            
        Case eGDSymbolGridMode_Tree
            With fgTree
                For lIndex = 0 To .SelectedRows - 1
                    lRecNum = g.SymbolPool.PoolRecForSymbol(.TextMatrix(.SelectedRow(lIndex), GDTreeCol(eGDTreeCol_Symbol)))
                    frmQuotes.AddSymbol lRecNum, "Daily"
                Next lIndex
            End With
    
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuAddToQuoteBoard.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuAscending_Click()
On Error GoTo ErrSection:

    Dim lTreeCol As Long
    Dim lVirtualCol As Long

    If m.Mode = eGDSymbolGridMode_Tree Then
        lTreeCol = CLng(Val(Parse(mnuGrid.Tag, vbTab, 2)))
        lVirtualCol = lTreeCol - (GDTreeCol(eGDTreeCol_Symbol) - kSymbolCol)
    Else
        lVirtualCol = CLng(Val(Parse(mnuGrid.Tag, vbTab, 2)))
        lTreeCol = lVirtualCol + (GDTreeCol(eGDTreeCol_Symbol) - kSymbolCol)
    End If

    SortTreeOnCol lTreeCol, 1
    m.SymbolGrid.SortOnCol lVirtualCol, 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuAscending.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuAutoChart_Click()
On Error GoTo ErrSection:

    AutoChart = Not AutoChart
    mnuAutoChart.Checked = AutoChart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuAutoChart.Click", eGDRaiseError_Show
    Resume ErrExit

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

    If ChangeGridFont(fgVirtual, True) Then
        fgTree.Font = fgVirtual.Font
        fgTree.Font = fgTree.Font
        fgTree.AutoSize GDTreeCol(eGDTreeCol_Symbol), fgTree.Cols - 1, False, 75
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuChart_Click()
On Error GoTo ErrSection:

    Dim nRow&, strSymbol$, nSymbolID&
    Dim frm As Form
    
    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            nRow = Val(Parse(mnuGrid.Tag, vbTab, 1))
            If nRow >= fgVirtual.FixedRows Then
                strSymbol = Trim(fgVirtual.TextMatrix(nRow, kSymbolCol))
                nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
            End If
            
        Case eGDSymbolGridMode_Tree
            nRow = Val(Parse(mnuGrid.Tag, vbTab, 1))
            If nRow >= fgTree.FixedRows Then
                strSymbol = Trim(fgTree.TextMatrix(nRow, GDTreeCol(eGDTreeCol_Symbol)))
                nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
            End If
    
    End Select
    
    If nSymbolID = 0 Then
        Beep
    ElseIf Not AutoChart Then
        ' just change symbol for the active chart
        SetActiveChartSymbol strSymbol
    Else
        ' use a new instance of frmChart
        Screen.MousePointer = vbHourglass
        Set frm = New frmChart          'new chart is always non-detached
        With frm
            .Chart.SetSymbol nSymbolID, True
            ShowForm frm, , , , ALT_GRID_ROW_COLOR
        End With
        Screen.MousePointer = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuChart.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuDescending_Click()
On Error GoTo ErrSection:

    Dim lTreeCol As Long
    Dim lVirtualCol As Long

    If m.Mode = eGDSymbolGridMode_Tree Then
        lTreeCol = CLng(Val(Parse(mnuGrid.Tag, vbTab, 2)))
        lVirtualCol = lTreeCol - (GDTreeCol(eGDTreeCol_Symbol) - kSymbolCol)
    Else
        lVirtualCol = CLng(Val(Parse(mnuGrid.Tag, vbTab, 2)))
        lTreeCol = lVirtualCol + (GDTreeCol(eGDTreeCol_Symbol) - kSymbolCol)
    End If

    SortTreeOnCol lTreeCol, -1
    m.SymbolGrid.SortOnCol lVirtualCol, -1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuDescending.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuEditObject_Click()
On Error GoTo ErrSection:
    
    Dim strID As String                 ' ID of the object to edit
    Dim strPath As String               ' Path of the object to edit
    Dim frm As Form                     ' Form of the appropriate editor
    
    strID = mnuEditObject.Tag
    If Len(strID) > 0 Then
        strPath = AddSlash(App.Path) & "Custom\"
        Select Case Left(strID, 3)
            Case "GRP":
                Set frm = New frmSymbolGroup
            Case "FIL":
                Set frm = New frmFilter
            Case "DSV", "DSP":
                Set frm = New frmCriteria
        End Select
        
        If frm Is Nothing Then
            Beep
        Else
            frm.ShowMe strPath, Mid(strID, 5)
            ''If Left(strID, 2) = "DS" Then CheckCriteria True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuEditObject.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuLookup_Click()
On Error GoTo ErrSection:

    Dim astrSymbol As New cGdArray
    
    Set astrSymbol = frmSymbolSelector.ShowMe("", False, True)
    
    If astrSymbol.Size > 0 Then
        Select Case m.Mode
            Case eGDSymbolGridMode_Virtual
                m.SymbolGrid.ShowRec g.SymbolPool.PoolRecForSymbol(astrSymbol(0), False)
            
            Case eGDSymbolGridMode_Tree
                ShowTreeSymbol astrSymbol(0)
                
        End Select
    End If
    
ErrExit:
    Set astrSymbol = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuLookup.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuMarketInfo_Click()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid
    
    lRow = ValOfText(Parse(mnuGrid.Tag, vbTab, 1))
    
    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            If lRow >= fgVirtual.FixedRows And lRow < fgVirtual.Rows Then
                frmMarkets.ShowMe fgVirtual.TextMatrix(lRow, kSymbolCol)
            End If
            
        Case eGDSymbolGridMode_Tree
            If lRow >= fgTree.FixedRows And lRow < fgTree.Rows Then
                frmMarkets.ShowMe fgTree.TextMatrix(lRow, GDTreeCol(eGDTreeCol_Symbol))
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuMarketInfo.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuRemoveCol_Click()
On Error GoTo ErrSection:

    Dim strID$, strDisplayFields$, i&
    
    strID = mnuEditObject.Tag
        
    strDisplayFields = GetGridFields
    i = InStr(UCase(strDisplayFields), "|" & UCase(strID) & "|")
    If i > 0 Then
        strDisplayFields = Left(strDisplayFields, i - 1) _
            & Mid(strDisplayFields, i + Len(strID) + 1)
        m.SymbolGrid.InitGrid fgVirtual, strDisplayFields
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuRemoveCol.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuShow_Click()
On Error GoTo ErrSection:

    cmdShow_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuShow.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuSymGroup_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim frm As New frmSymbolGroup
    
    frm.ShowMe AddSlash(App.Path) & "Custom\", mnuSymGroup(Index).Tag, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuSymGroup.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuSyncAll_Click()
On Error GoTo ErrSection:

    Dim i&, nRow&, strSymbol$, nSymbolID&
    Dim frm As Form, frmActive As Form
    
    nRow = Val(Parse(mnuGrid.Tag, vbTab, 1))
    
    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            If nRow >= fgVirtual.FixedRows Then
                strSymbol = Trim(fgVirtual.TextMatrix(nRow, kSymbolCol))
                nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
            End If
            
        Case eGDSymbolGridMode_Tree
            If nRow >= fgTree.FixedRows Then
                strSymbol = Trim(fgTree.TextMatrix(nRow, GDTreeCol(eGDTreeCol_Symbol)))
                nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
            End If
    
    End Select
    
    If nSymbolID = 0 Then
        Beep
    Else
        Set frmActive = ActiveChart
        For i = 0 To Forms.Count - 1
            If IsFrmChart(Forms(i)) Then
                Set frm = Forms(i)
                If Not frm Is frmActive Then
                    frm.Chart.SetSymbol nSymbolID, True
                End If
            End If
        Next
        ' do active chart last
        Set frm = frmActive
        If Not frm Is Nothing Then
            frm.Chart.SetSymbol nSymbolID, True
        End If
        Set frm = Nothing
        Set frmActive = Nothing
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.mnuSyncAll.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub tmrSortCol_Timer()
On Error GoTo ErrSection:

    Dim iCol&
    
    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            If Not MouseIsPressed Then
                tmrSortCol.Enabled = False
                iCol = fgVirtual.MouseCol
                If fgVirtual.MouseRow = 0 And iCol >= 0 And iCol < fgVirtual.Cols Then
                    m.SymbolGrid.SortOnCol iCol
                    SortTreeOnCol iCol + (GDTreeCol(eGDTreeCol_Symbol) - kSymbolCol)
                End If
            End If
            
        Case eGDSymbolGridMode_Tree
            If Not MouseIsPressed Then
                tmrSortCol.Enabled = False
                iCol = fgTree.MouseCol
                If fgTree.MouseRow = 0 And iCol >= 0 And iCol < fgTree.Cols Then
                    SortTreeOnCol iCol
                    m.SymbolGrid.SortOnCol iCol - (GDTreeCol(eGDTreeCol_Symbol) - kSymbolCol)
                End If
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.tmrSortCol.Timer", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub ShowGridPopup(ByVal nRow&, ByVal nCol&)
On Error GoTo ErrSection:

    Dim i&, strID$, nCount&
    Dim aItems As New cGdArray
    Dim obj As Object
    Dim strSymbol As String
    Dim strColHeader As String

    ' no longer needed
    mnuSyncAll.Visible = False
    mnuAutoChart.Visible = False
    
    mnuAutoChart.Checked = AutoChart

    ' store current row and column for reference
    mnuGrid.Tag = Str(nRow) & vbTab & Str(nCol)
    
    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            If nRow >= fgVirtual.FixedRows Then
                strSymbol = Trim(fgVirtual.TextMatrix(nRow, kSymbolCol))
            End If
            strColHeader = Trim(fgVirtual.TextMatrix(0, nCol))
            If nCol > kSymbolCol Then
                strID = Parse(fgVirtual.ColData(nCol), vbTab, 2)
            End If
            
        Case eGDSymbolGridMode_Tree
            If nRow >= fgTree.FixedRows Then
                strSymbol = Trim(fgTree.TextMatrix(nRow, GDTreeCol(eGDTreeCol_Symbol)))
            End If
            strColHeader = Trim(fgTree.TextMatrix(0, nCol))
            strID = ""
    
    End Select
    
    ' current symbol
    If Len(strSymbol) > 0 Then
        mnuSymbol.Caption = "ROW:   " & strSymbol & "  ..."
        mnuMarketInfo.Caption = "Market Information for " & strSymbol
        mnuMarketInfo.Enabled = True
    Else
        mnuSymbol.Caption = "ROW: "
        mnuMarketInfo.Caption = "MarketInformation"
        mnuMarketInfo.Enabled = False
    End If
    If AutoChart Then
        mnuChart.Caption = "   New Chart"
    Else
        mnuChart.Caption = "   Synchronize Chart Symbol"
    End If
    
    ' current column
    mnuColumn.Caption = "COLUMN:   " & strColHeader & "  ..."
    mnuRemoveCol.Enabled = False
    mnuEditObject.Visible = False
    If nCol > kSymbolCol And Len(strID) > 0 Then
        mnuRemoveCol.Enabled = True
        Set obj = g.SymbolPool.PoolObject(strID)
        If Not obj Is Nothing Then
            If obj.Custom Then
                Select Case Left(strID, 3)
                    Case "GRP":
                        mnuEditObject.Caption = "   Edit Symbol Group"
                    Case "FIL":
                        mnuEditObject.Caption = "   Edit Filter"
                    Case "DSV", "DSP":
                        mnuEditObject.Caption = "   Edit Criteria"
                End Select
                mnuEditObject.Visible = True
                mnuEditObject.Tag = strID
            End If
            Set obj = Nothing
        End If
    End If

    ' symbol groups
    aItems.Clear
    aItems.Add " (new Symbol Group)"
    For i = 1 To g.SymbolPool.SymbolGroups.Count
        With g.SymbolPool.SymbolGroups.Item(i)
            If .Custom And Len(.Name) > 0 Then
                aItems.Add .Name & vbTab & .ID
            End If
        End With
    Next
    aItems.Sort
    'add menu item for each
    For i = 0 To aItems.Size - 1
        If i > mnuSymGroup.UBound Then
            Load mnuSymGroup(i)
            mnuSymGroup(i).Visible = True
        End If
        mnuSymGroup(i).Caption = Parse(aItems(i), vbTab, 1)
        mnuSymGroup(i).Tag = Parse(aItems(i), vbTab, 2)
    Next
    'remove extras
    For i = mnuSymGroup.UBound To aItems.Size Step -1
        If i > 0 Then Unload mnuSymGroup(i)
    Next
    
    ' multiple charts?
    mnuSyncAll.Enabled = False
    nCount = 0
    For i = 0 To Forms.Count - 1
        If IsFrmChart(Forms(i)) Then
            nCount = nCount + 1
            If nCount > 1 Then
                mnuSyncAll.Enabled = True
                Exit For
            End If
        End If
    Next
    
    ' show popup
    Me.PopupMenu mnuGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ShowGridPopup", eGDRaiseError_Raise

End Sub

Public Sub RefreshGrid()
On Error GoTo ErrSection:

    Dim nSymbolIDtoShow&, nRow&
    Dim lPoolRec As Long
    Dim lIndex As Long
    
    LoadCombo
    
    ' Need to reset the pool records for the symbols in the table...
    For lIndex = 0 To m.tblSymbols.NumRecords - 1
        lPoolRec = g.SymbolPool.PoolRecForSymbolID(gdGetTableNum(m.hSymbols, eTblCol_SymbolID, lIndex))
        gdSetTableNum m.hSymbols, eTblCol_PoolRec, lIndex, lPoolRec
    Next lIndex
    
    ' Do this to force the new filter flag to get set on the grid
    SymbolGrid.FilterID = SymbolGrid.FilterID
    
    nSymbolIDtoShow = g.SymbolPool.SymbolIDforSymbol(m.strLastSymbol)
    SortTreeOnCol -1, 0
    m.SymbolGrid.SortOnCol -1, 0, nSymbolIDtoShow
    
    SetComboTooltip
    
    fgVirtual.Refresh
    fgTree.Refresh
    UpdateFlagCount

    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            nRow = fgVirtual.RowSel
            If nRow >= fgVirtual.FixedRows And nRow < fgVirtual.Rows Then
                If AutoChart Then
                    SetActiveChartSymbol Trim(fgVirtual.TextMatrix(fgVirtual.RowSel, kSymbolCol))
                End If
            End If
        
        Case eGDSymbolGridMode_Tree
            nRow = fgTree.RowSel
            If nRow >= fgTree.FixedRows And nRow < fgTree.Rows Then
                If AutoChart Then
                    SetActiveChartSymbol Trim(fgTree.TextMatrix(fgTree.RowSel, GDTreeCol(eGDTreeCol_Symbol)))
                End If
            End If
            
    End Select
       
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.RefreshGrid", eGDRaiseError_Raise

End Sub

Public Sub ShowSymbol(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Select Case m.Mode
        Case eGDSymbolGridMode_Tree
            If strSymbol <> fgTree.TextMatrix(fgTree.RowSel, GDTreeCol(eGDTreeCol_Symbol)) Then
                ShowTreeSymbol strSymbol
            End If
            
        Case eGDSymbolGridMode_Virtual
            ShowListSymbol strSymbol
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ShowSymbol", eGDRaiseError_Raise

End Sub

Public Sub ShowInitialSymbol()
On Error GoTo ErrSection:

    ShowSymbol ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ShowInitialSymbol", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFlagCount
'' Description: Update the flag frame according to how many symbols are flagged
'' Inputs:      Whether or not this is being called from Form_Load
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateFlagCount()
On Error Resume Next

    Dim lField As Long                  ' Pool field for the flag array
    Dim lFilterFld As Long              ' Pool field for the current filter
    Dim lNumFlagged As Long             ' Number of symbols that are flagged
    Dim lNumSymbols As Long             ' Number of symbols in the current filter
    Dim alTemp As New cGdArray          ' Temporary array
    Dim lRow As Long
    Dim lCount As Long
    
    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            lField = Val(fgVirtual.ColData(kFlagCol))
            
            With g.SymbolPool
                lFilterFld = .FieldNumForID(cboList.SelectedItem.Key)
                alTemp.Create eGDARRAY_TinyInts, .NumRecords
                alTemp.ArrayOperate .ArrayTable.FieldArray(lField), "AND", .ArrayTable.FieldArray(lFilterFld)
                lNumFlagged = alTemp.CountOf(1)
                lNumSymbols = .ArrayTable.FieldArray(lFilterFld).CountOf(1)
            End With
    
        Case eGDSymbolGridMode_Tree
            lField = Val(fgTree.ColData(GDTreeCol(eGDTreeCol_Flagged)))
            lNumSymbols = m.lNumSectors
            With g.SymbolPool
                alTemp.Create eGDARRAY_TinyInts, .NumRecords
                alTemp.ArrayOperate .ArrayTable.FieldArray(lField), "AND", m.alSectorPool
                lNumFlagged = alTemp.CountOf(1)
            End With
    
    End Select
    
    lblNumFlagged.Tag = CStr(lNumSymbols)
    UpdateNumberLabel
    
    Enable cmdClearAll, (lNumFlagged > 0)
    Enable cmdSaveFlags, (lNumFlagged > 0)
    If lNumFlagged = 1 Then
        cmdSaveFlags.ToolTipText = "Add 1 flagged symbol to a new or existing Symbol Group"
        lblFlags.ToolTipText = "1 symbol currently flagged"
    Else
        cmdSaveFlags.ToolTipText = "Add " & CStr(lNumFlagged) & " flagged symbols to a new or existing Symbol Group"
        lblFlags.ToolTipText = CStr(lNumFlagged) & " symbols currently flagged"
    End If
    
    Set alTemp = Nothing

End Sub

Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim bExtend As Boolean
    Dim alColWidths As New cGdArray
    Dim lIndex As Long
    Dim lRow As Long
    Dim lCol As Long
    Dim strText As String
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .TextAlign = taCenterMiddle
        .Text = cboList.Text
        .TextAlign = taLeftMiddle
        .Font.Bold = False
        
        .Paragraph = ""
        .Paragraph = ""
        
        Select Case m.Mode
            Case eGDSymbolGridMode_Virtual
                alColWidths.Create eGDARRAY_Longs, fgVirtual.Cols
                For lIndex = 0 To fgVirtual.Cols - 1
                    alColWidths(lIndex) = fgVirtual.ColWidth(lIndex)
                Next lIndex
                bExtend = fgVirtual.ExtendLastCol
                
                fgVirtual.ExtendLastCol = False
                fgVirtual.AutoSize 0, fgVirtual.Cols - 1
                
                If frmPrintPreview.GoingToFile Then
                    frmPrintPreview.GridToTable fgVirtual
                Else
                    .RenderControl = fgVirtual.hWnd
                End If
                
                fgVirtual.ExtendLastCol = bExtend
                
                For lIndex = 0 To fgVirtual.Cols - 1
                    fgVirtual.ColWidth(lIndex) = alColWidths(lIndex)
                Next lIndex
                
            Case eGDSymbolGridMode_Tree
                alColWidths.Create eGDARRAY_Longs, fgTree.Cols
                For lIndex = 0 To fgTree.Cols - 1
                    alColWidths(lIndex) = fgTree.ColWidth(lIndex)
                Next lIndex
                bExtend = fgTree.ExtendLastCol
                
                fgTree.ExtendLastCol = False
                fgTree.AutoSize GDTreeCol(eGDTreeCol_Symbol), fgTree.Cols - 1
                
                If frmPrintPreview.GoingToFile Then
                    frmPrintPreview.GridToTable fgTree
                Else
                    .RenderControl = fgTree.hWnd
                End If
                
                fgTree.ExtendLastCol = bExtend
                
                For lIndex = 0 To fgTree.Cols - 1
                    fgTree.ColWidth(lIndex) = alColWidths(lIndex)
                Next lIndex
        
        End Select
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.GenerateReport", eGDRaiseError_Raise

End Sub

Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    If fgVirtual.Rows > 550 Then
        If AskBox("h=Warning ; i=? ; b=+Yes|-No ; Printing this many symbols may take a while.||Do you want to continue?") = "N" Then
            Exit Function
        End If
    End If

    PrintMe = frmPrintPreview.ShowMe("CNV SymbolGrid", frmSymbolGrid)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSymbolGrid.PrintMe", eGDRaiseError_Raise
    
End Function

Private Sub SetComboTooltip()
On Error Resume Next

    Dim strID$, strDesc$
    Dim obj As Object
    
    strID = cboList.SelectedItem.Key
    If strID = "Sectors" Then
        strDesc = "A Tree View of Sectors, Subsectors, and Stocks"
    ElseIf Len(strID) > 0 Then
        Set obj = g.SymbolPool.PoolObject(strID)
        strDesc = obj.Desc
        If Len(strDesc) = 0 Then
            strDesc = obj.Name
        End If
        Set obj = Nothing
    End If
    cboList.ToolTipText = TipStr(strDesc)

End Sub

Public Sub InitForm()
On Error GoTo ErrSection:

    Dim i&
    Dim strDisplayFields$
       
    cboList.ImageList = frmMain.img16
    LoadCombo
    cboList.Locked = True
    
    m.bShowFlags = GetIniFileProperty("ShowFlags", True, "Grid", g.strIniFile)

    Set m.SymbolGrid = New cSymbolGrid
    With m.SymbolGrid
        If cboList.ComboItems.Count > 0 Then
            .FilterID = cboList.SelectedItem.Key
        End If
        
        strDisplayFields = GetIniFileProperty("DisplayFields", "", "Grid", g.strIniFile)
                
        .InitGrid fgVirtual, strDisplayFields
        .SortOnCol kSymbolCol, 1
    End With
    
    Set m.tblSymbols = New cGdTable
    With m.tblSymbols
        .CreateField eGDARRAY_Longs, TblCol(eTblCol_SymbolID), "SymbolID"
        .CreateField eGDARRAY_Longs, TblCol(eTblCol_PoolRec), "PoolRec"
        .CreateField eGDARRAY_Strings, TblCol(eTblCol_Symbol), "Symbol"
        .CreateField eGDARRAY_Strings, TblCol(eTblCol_Description), "Description"
        .CreateField eGDARRAY_TinyInts, TblCol(eTblCol_Level), "Level"
        .CreateField eGDARRAY_Strings, TblCol(eTblCol_SortKey), "SortKey"
    End With
    m.hSymbols = m.tblSymbols.TableHandle
    
    InitTreeGrid strDisplayFields
    LoadTreeSymbols
    LoadTreeGrid
    
    UpdateFlagCount
    ShowFlags

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.InitForm", eGDRaiseError_Raise

End Sub

Private Sub UpdateNumberLabel()
On Error GoTo ErrSection:

    Dim nRow&
    Dim lNum As Long
    Dim lCount As Long
    Dim lParentRow As Long
    Dim strParent As String

    Select Case m.Mode
        Case eGDSymbolGridMode_Virtual
            With fgVirtual
                nRow = .Row
                If nRow >= .FixedRows And nRow < .Rows Then
                    lblNumFlagged.Caption = CStr(nRow) & " of " & lblNumFlagged.Tag & " symbols"
                Else
                    lblNumFlagged.Caption = lblNumFlagged.Tag & " symbols"
                End If
            End With
            
        Case eGDSymbolGridMode_Tree
            With fgTree
                nRow = .Row
                If nRow >= .FixedRows And nRow < .Rows Then
                    GetSiblingNumber nRow, lNum, lCount
                    lParentRow = .GetNodeRow(nRow, flexNTParent)
                    If lParentRow = -1& Then
                        lblNumFlagged.Caption = CStr(lNum) & " of " & Str(lCount) & " sectors"
                    Else
                        lblNumFlagged.Caption = CStr(lNum) & " of " & Str(lCount) & " symbols in " & .TextMatrix(lParentRow, GDTreeCol(eGDTreeCol_Symbol))
                    End If
                Else
                    lblNumFlagged.Caption = lblNumFlagged.Tag & " sectors"
                End If
            End With
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.UpdateNumberLabel", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitTreeGrid
'' Description: Initialize the tree style grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitTreeGrid(ByVal strFields As String)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim astrFields As New cGdArray      ' Array of fields to display in the grid
    Dim strID As String                 ' ID of the item to show
    Dim lField As Long                  ' Field number of the item in the pool
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCount As Long                  ' Number of columns
    
    With fgTree
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        SetupGrid fgTree, eGridMode_Tree
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbApplicationWorkspace
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSort
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .OutlineBar = flexOutlineBarSimpleLeaf
        .Cols = GDTreeCol(eGDTreeCol_Symbol) - 1
        .FixedCols = 0
        .Rows = 1
        .FixedRows = 1

        If InStr(UCase(strFields), "INF:SYMBOL|") = 0 Then
            strFields = "GRP:_FLAGS_.GRP|INF:Symbol|INF:Description|INF:FirstDate|INF:LastDate"
        ElseIf Parse(UCase(strFields), "|", 1) <> "GRP:_FLAGS_.GRP" Then
            strFields = "GRP:_FLAGS_.GRP|" & strFields
        End If
        
        astrFields.SplitFields strFields, "|"
        lCount = 1&
        For lIndex = 0 To astrFields.Size - 1
            strID = Parse(astrFields(lIndex), "\", 1)
            lField = g.SymbolPool.FieldNumForID(strID)
        
            If lField >= 0 Then
                .Cols = lCount + 1
                .TextMatrix(0, lCount) = Space(4) & g.SymbolPool.ArrayTable.FieldName(lField)
                .ColData(lCount) = Str(lField) & vbTab & astrFields(lIndex)
                .ColAlignment(lCount) = flexAlignCenterCenter
                Select Case lField
                    Case 2: 'Symbol
                        .ColAlignment(lCount) = flexAlignLeftCenter
                    Case 4: 'Desc
                        .ColWidth(lCount) = 2.4 * .ColWidth(kSymbolCol)
                        .ColAlignment(lCount) = flexAlignLeftCenter
                    Case 7: 'SecType
                        .ColWidth(lCount) = 60 * 10
                    Case Else:
                        If strID = "GRP:_FLAGS_.GRP" Then
                            .TextMatrix(0, lCount) = "Flag"
                            .ColDataType(lCount) = flexDTBoolean
                            .ColWidth(lCount) = 40 * 10
                        End If
                End Select
                lCount = lCount + 1
            End If
        Next lIndex
        
        If .Cols > GDTreeCol(eGDTreeCol_Symbol) Then
            .OutlineCol = GDTreeCol(eGDTreeCol_Symbol)
            .FrozenCols = GDTreeCol(eGDTreeCol_Symbol) + 1
        End If
        
        .ColHidden(GDTreeCol(eGDTreeCol_TableIndex)) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftCenter
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.InitTreeGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTreeSymbols
'' Description: Load the sector symbols into the memory table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTreeSymbols()
On Error GoTo ErrSection:

    Dim alSectors As New cGdArray       ' Array of sectors from the symbol universe
    Dim alSubSectors As New cGdArray    ' Array of subsectors of a sector
    Dim alStocks As New cGdArray        ' Array of stocks in a subsector
    Dim hSectors As Long                ' Handle to the array of sectors
    Dim hSubSectors As Long             ' Handle to the array of subsectors
    Dim hStocks As Long                 ' Handle to the array of stocks
    Dim lSector As Long                 ' Index into a for loop
    Dim lSubSector As Long              ' Index into a for loop
    Dim lStock As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol for the SymbolID
    Dim strSectorSym As String          ' Symbol for the sector of the symbol
    Dim strSubSectorSym As String       ' Symbol for the subsector of the symbol
    Dim strDesc As String               ' Description for the SymbolID
    Dim lPoolRec As Long                ' Pool record for the SymbolID
    Dim lTblIndex As Long               ' Index into the table
    
    alSectors.Create eGDARRAY_Longs
    hSectors = alSectors.ArrayHandle
    alSubSectors.Create eGDARRAY_Longs
    hSubSectors = alSubSectors.ArrayHandle
    alStocks.Create eGDARRAY_Longs
    hStocks = alStocks.ArrayHandle
    
    Set m.alSectorPool = New cGdArray
    m.alSectorPool.Create eGDARRAY_TinyInts, g.SymbolPool.NumRecords, 0
    
    If SU_GetGroupChildren(0&, alSectors) Then
        m.tblSymbols.NumRecords = 20000
        m.lNumSectors = alSectors.Size
        For lSector = 0 To alSectors.Size - 1
            lPoolRec = g.SymbolPool.PoolRecForSymbolID(alSectors(lSector))
            m.alSectorPool(lPoolRec) = 1
            strSectorSym = g.SymbolPool.Symbol(lPoolRec)
            strDesc = g.SymbolPool.Desc(lPoolRec)
            
            gdSetTableNum m.hSymbols, eTblCol_SymbolID, lTblIndex, gdGetNum(hSectors, lSector)
            gdSetTableStr m.hSymbols, eTblCol_Symbol, lTblIndex, strSectorSym
            gdSetTableNum m.hSymbols, eTblCol_PoolRec, lTblIndex, lPoolRec
            gdSetTableStr m.hSymbols, eTblCol_Description, lTblIndex, strDesc
            gdSetTableNum m.hSymbols, eTblCol_Level, lTblIndex, 0
            gdSetTableStr m.hSymbols, eTblCol_SortKey, lTblIndex, Pad(strSectorSym, 14, "L") & Pad("", 14, "L") & Pad("", 14, "L")
            lTblIndex = lTblIndex + 1
            
            If SU_GetGroupChildren(alSectors(lSector), alSubSectors) Then
                For lSubSector = 0 To alSubSectors.Size - 1
                    lPoolRec = g.SymbolPool.PoolRecForSymbolID(alSubSectors(lSubSector))
                    m.alSectorPool(lPoolRec) = 1
                    strSubSectorSym = g.SymbolPool.Symbol(lPoolRec)
                    strDesc = g.SymbolPool.Desc(lPoolRec)

                    gdSetTableNum m.hSymbols, eTblCol_SymbolID, lTblIndex, gdGetNum(hSubSectors, lSubSector)
                    gdSetTableStr m.hSymbols, eTblCol_Symbol, lTblIndex, strSubSectorSym
                    gdSetTableNum m.hSymbols, eTblCol_PoolRec, lTblIndex, lPoolRec
                    gdSetTableStr m.hSymbols, eTblCol_Description, lTblIndex, strDesc
                    gdSetTableNum m.hSymbols, eTblCol_Level, lTblIndex, 1
                    gdSetTableStr m.hSymbols, eTblCol_SortKey, lTblIndex, Pad(strSectorSym, 14, "L") & Pad(strSubSectorSym, 14, "L") & Pad("", 14, "L")
                    lTblIndex = lTblIndex + 1
                    
                    If SU_GetGroupChildren(alSubSectors(lSubSector), alStocks) Then
                        For lStock = 0 To alStocks.Size - 1
                            lPoolRec = g.SymbolPool.PoolRecForSymbolID(alStocks(lStock))
                            m.alSectorPool(lPoolRec) = 1
                            strSymbol = g.SymbolPool.Symbol(lPoolRec)
                            strDesc = g.SymbolPool.Desc(lPoolRec)

                            If Len(strSymbol) > 0 Then
                                gdSetTableNum m.hSymbols, eTblCol_SymbolID, lTblIndex, gdGetNum(hStocks, lStock)
                                gdSetTableStr m.hSymbols, eTblCol_Symbol, lTblIndex, strSymbol
                                gdSetTableNum m.hSymbols, eTblCol_PoolRec, lTblIndex, lPoolRec
                                gdSetTableStr m.hSymbols, eTblCol_Description, lTblIndex, strDesc
                                gdSetTableNum m.hSymbols, eTblCol_Level, lTblIndex, 2
                                gdSetTableStr m.hSymbols, eTblCol_SortKey, lTblIndex, Pad(strSectorSym, 14, "L") & Pad(strSubSectorSym, 14, "L") & Pad(strSymbol, 14, "L")
                                lTblIndex = lTblIndex + 1
                            End If
                        Next lStock
                    End If
                Next lSubSector
            End If
        Next lSector
        
        m.tblSymbols.NumRecords = lTblIndex - 1&
        
        Set m.aSortedIndex = m.tblSymbols.CreateIndex
        m.tblSymbols.SortIndex m.aSortedIndex, TblCol(eTblCol_SortKey)
        m.hSortedIndex = m.aSortedIndex.ArrayHandle
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.LoadTreeSymbols", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTreeGrid
'' Description: Load the tree grid from the memory tables
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTreeGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lTblIndex As Long               ' Index into the sectors table
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgTree
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 0 To gdGetSize(m.hSortedIndex) - 1
            lTblIndex = gdGetNum(m.hSortedIndex, lIndex)
            
            If gdGetTableNum(m.hSymbols, eTblCol_Level, lTblIndex) = 0 Then
                .AddItem Str(lIndex)
                SetTreeGridRow .Rows - 1, lTblIndex
                
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 1
                
                .AddItem vbTab & vbTab & "(blank)"
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 2
            End If
        Next lIndex
        
        m.lOpenRow = -1&
        If .Cols > GDTreeCol(eGDTreeCol_Symbol) Then
            .Outline -1
            .AutoSize GDTreeCol(eGDTreeCol_Symbol), .Cols - 1, False, 75
            .Outline 2
            .Outline 1
        
            SortTreeOnCol GDTreeCol(eGDTreeCol_Symbol), 1
            SetBackColors fgTree
        End If

        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.LoadTreeGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExpandRow
'' Description: Expand the next level of a given row
'' Inputs:      Row in the grid to expand
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExpandRow(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim lTblIndex As Long               ' Index of the row into the table
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLevel As Long                  ' Level of the current row
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lInsertRow As Long              ' Row to insert at
    Dim lStart As Long                  ' Starting place in the table
    
    With fgTree
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        lStart = CLng(.TextMatrix(lRow, GDTreeCol(eGDTreeCol_TableIndex)))
        lLevel = .RowOutlineLevel(lRow) - 1
        
        .RemoveItem lRow + 1
        lInsertRow = lRow + 1
        
        For lIndex = lStart + 1 To m.tblSymbols.NumRecords - 1
            lTblIndex = gdGetNum(m.hSortedIndex, lIndex)
            If TableNum(eTblCol_Level, lTblIndex) = lLevel Then Exit For
            
            If TableNum(eTblCol_Level, lTblIndex) = lLevel + 1 Then
                .AddItem Str(lIndex), lInsertRow
                SetTreeGridRow lInsertRow, lTblIndex
                
                .IsSubtotal(lInsertRow) = True
                .RowOutlineLevel(lInsertRow) = TableNum(eTblCol_Level, lTblIndex) + 1
                lInsertRow = lInsertRow + 1
                
                If TableNum(eTblCol_Level, lTblIndex) < 2 Then
                    .AddItem vbTab & vbTab & "(blank)", lInsertRow
                    .IsSubtotal(lInsertRow) = True
                    .RowOutlineLevel(lInsertRow) = TableNum(eTblCol_Level, lTblIndex) + 2
                    lInsertRow = lInsertRow + 1
                    .IsCollapsed(lInsertRow - 2) = flexOutlineCollapsed
                End If
                
            End If
        Next lIndex
        
        .Row = lRow
        .RowSel = lRow
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ExpandRow", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowTreeSymbol
'' Description: Highlight the line with the given symbol
'' Inputs:      Symbol to highlight
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowTreeSymbol(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strSector As String             ' Symbol of the sector
    Dim strSubsector As String          ' Symbol of the subsector
    Dim lRow As Long                    ' Row where we found what we were looking for
    Dim lParentRow As Long              ' Row of the parent node
        
    With fgTree
        If Len(strSymbol) = 0 Then
            strSymbol = Trim(GetIniFileProperty("InitialSymbol", "", "Grid", g.strIniFile))
            If strSymbol = "" Or strSymbol = "$AD" Then strSymbol = "$DJIA"
        End If
        
        For lIndex = 0 To m.tblSymbols.NumRecords - 1
            If TableStr(eTblCol_Symbol, lIndex) = strSymbol Then
                strSector = Trim(Left(TableStr(eTblCol_SortKey, lIndex), 14))
                strSubsector = Trim(Mid(TableStr(eTblCol_SortKey, lIndex), 15, 14))
                Exit For
            End If
        Next lIndex
        
        If Len(strSector) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, GDTreeCol(eGDTreeCol_Symbol)) = strSector Then
                    lRow = lIndex
                    .IsCollapsed(lIndex) = flexOutlineExpanded
                    Exit For
                End If
            Next lIndex
        
            If Len(strSubsector) > 0 Then
                For lIndex = lRow To .Rows - 1
                    If .TextMatrix(lIndex, GDTreeCol(eGDTreeCol_Symbol)) = strSubsector Then
                        lRow = lIndex
                        .IsCollapsed(lIndex) = flexOutlineExpanded
                        Exit For
                    End If
                Next lIndex
            
                If Left(strSymbol, 1) <> "$" Then
                    For lIndex = lRow To .Rows - 1
                        If .TextMatrix(lIndex, GDTreeCol(eGDTreeCol_Symbol)) = strSymbol Then
                            lRow = lIndex
                            .Row = lIndex
                            .RowSel = lIndex
                            
                            lParentRow = .GetNodeRow(lIndex, flexNTParent)
                            If lParentRow <> -1 Then .TopRow = lParentRow
                            If lIndex < .TopRow Or lIndex > .BottomRow Then
                                .ShowCell lIndex, GDTreeCol(eGDTreeCol_Symbol)
                            End If
                            Exit For
                        End If
                    Next lIndex
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ShowTreeSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTreeGridRow
'' Description: Set a row in the grid to appropriate data from the pool
'' Inputs:      Row to set in the Grid, Row of data in the Table
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTreeGridRow(ByVal lGridRow As Long, ByVal lTblIndex As Long)
On Error GoTo ErrSection:

    Dim lPoolRec As Long                ' Record for this symbol in the pool
    Dim lField As Long                  ' Field number for item in the pool
    Dim lCol As Long                    ' Index into a for loop
    Dim vData As Variant                ' Data from the pool
    
    lPoolRec = gdGetTableNum(m.hSymbols, eTblCol_PoolRec, lTblIndex)
    
    With fgTree
        For lCol = GDTreeCol(eGDTreeCol_Flagged) To .Cols - 1
            lField = Val(.ColData(lCol))
            vData = g.SymbolPool.DataItem(lField, lPoolRec, "")
            Select Case VarType(vData)
                Case vbDate
                    .ColFormat(lCol) = DateFormat("Format", MM_DD_YY)
                    .Cell(flexcpText, lGridRow, lCol) = CDbl(vData)
                
                Case Else
                    .TextMatrix(lGridRow, lCol) = vData
            
            End Select
        
        Next lCol
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.SetTreeGridRow", eGDRaiseError_Raise
    
End Sub

Private Sub ClearRow(ByVal lGridRow As Long)
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgTree
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lCol = GDTreeCol(eGDTreeCol_Symbol) + 1 To .Cols - 1
            If UCase(Trim(.TextMatrix(0, lCol))) <> "DESCRIPTION" Then
                .TextMatrix(lGridRow, lCol) = ""
            End If
        Next lCol
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ClearRow", eGDRaiseError_Raise
    
End Sub

Private Sub ShowTreeData(ByVal lGridRow As Long)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim lTblIndex As Long               ' Index of the symbol in the table
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgTree
        .BackColor = fgVirtual.BackColor
        .BackColorAlternate = fgVirtual.BackColorAlternate
        If g.nColorTheme = kDarkThemeColor Then
            .NodeClosedPicture = Picture1
            .NodeOpenPicture = Picture2
        End If
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lRow = .FixedRows To .Rows - 1
            ClearRow lRow
        Next lRow

        lTblIndex = gdGetNum(m.hSortedIndex, CLng(Val((.TextMatrix(lGridRow, GDTreeCol(eGDTreeCol_TableIndex))))))
        SetTreeGridRow lGridRow, lTblIndex
        
        lRow = .GetNodeRow(lGridRow, flexNTParent)
        Do While lRow <> -1&
            lTblIndex = gdGetNum(m.hSortedIndex, CLng(Val((.TextMatrix(lRow, GDTreeCol(eGDTreeCol_TableIndex))))))
            SetTreeGridRow lRow, lTblIndex
            lRow = .GetNodeRow(lRow, flexNTParent)
        Loop
        
        Select Case .IsCollapsed(lGridRow)
            Case flexOutlineCollapsed
                m.lOpenRow = .GetNodeRow(lGridRow, flexNTParent)

                lRow = .GetNodeRow(lGridRow, flexNTFirstSibling)
                Do While lRow <> -1&
                    lTblIndex = gdGetNum(m.hSortedIndex, CLng(Val((.TextMatrix(lRow, GDTreeCol(eGDTreeCol_TableIndex))))))
                    SetTreeGridRow lRow, lTblIndex
                    lRow = .GetNodeRow(lRow, flexNTNextSibling)
                Loop
            
            Case flexOutlineExpanded
                m.lOpenRow = lGridRow

                lRow = .GetNodeRow(lGridRow, flexNTFirstChild)
                Do While lRow <> -1&
                    lTblIndex = gdGetNum(m.hSortedIndex, CLng(Val((.TextMatrix(lRow, GDTreeCol(eGDTreeCol_TableIndex))))))
                    SetTreeGridRow lRow, lTblIndex
                    lRow = .GetNodeRow(lRow, flexNTNextSibling)
                Loop
                If m.bSortedDescending Then
                    SortTreeOnCol m.lSortedCol, -1
                Else
                    SortTreeOnCol m.lSortedCol, 1
                End If
        
        End Select

        .Row = lGridRow
        .RowSel = lGridRow
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ShowTreeData", eGDRaiseError_Raise
    
End Sub

Private Sub SortTreeOnCol(ByVal lCol As Long, Optional ByVal lDirection As Long = 0&)
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lRedraw As Long
    Dim lFirstRow As Long
    Dim lLastRow As Long
    Dim lSibling As Long
    Dim lParent As Long
    Dim lGrandparent As Long
    Dim lParentSibling As Long
    Dim lLevel As Long
    Dim strSymbol As String
    Dim lParentRow As Long
    Static lPrevCol As Long
    Static bPrevDescending As Boolean
    
    If fgTree.Rows = fgTree.FixedRows Then Exit Sub

    If lCol = -1 Then
        lCol = m.lSortedCol
        If m.bSortedDescending Then
            lDirection = -1
        Else
            lDirection = 1
        End If
    Else
        lPrevCol = m.lSortedCol
        bPrevDescending = m.bSortedDescending
    End If

    If lCol < 0 Or lCol >= fgTree.Cols Then Exit Sub
        
    lRedraw = fgTree.Redraw
    fgTree.Redraw = flexRDNone
    
    If lCol = m.lSortedCol And lDirection = 0 Then
        m.bSortedDescending = Not m.bSortedDescending
    ElseIf lDirection = -1 Then
        m.bSortedDescending = True
    Else
        m.bSortedDescending = False
    End If
    m.lSortedCol = lCol

    With fgTree
        If .Row >= .FixedRows And .Row < .Rows Then
            strSymbol = .TextMatrix(.Row, GDTreeCol(eGDTreeCol_Symbol))
        End If
        
        If m.lOpenRow = -1& Then
            lFirstRow = .FixedRows
            lLastRow = .GetNodeRow(lFirstRow, flexNTLastSibling)
        ElseIf .GetNode(m.lOpenRow).Level < 3 Then
            lFirstRow = .GetNodeRow(m.lOpenRow, flexNTFirstChild)
            lLastRow = .GetNodeRow(m.lOpenRow, flexNTLastChild)
        Else
            lFirstRow = .GetNodeRow(m.lOpenRow, flexNTParent)
        End If
        
        lIndex = lFirstRow
        lLevel = .GetNode(lIndex).Level
        
        ' First, remove all of the child nodes of all but the last child of the parent...
        Do While lIndex <> -1&
            lSibling = .GetNodeRow(lIndex, flexNTNextSibling)
            
            For lIndex2 = lSibling - 1 To lIndex + 1 Step -1
                .RemoveItem lIndex2
            Next lIndex2
            lIndex = .GetNodeRow(lIndex, flexNTNextSibling)
        Loop
        lLastRow = .GetNodeRow(lFirstRow, flexNTLastSibling)
        
        ' If the current node has a sibling...
        If .GetNodeRow(lLastRow, flexNTParent) <> -1 Then
            lParent = .GetNodeRow(lLastRow, flexNTParent)
            lGrandparent = .GetNodeRow(lParent, flexNTParent)
            If lParent <> -1& Then
                lSibling = .GetNodeRow(lParent, flexNTNextSibling)
            Else
                lSibling = -1&
            End If
            If lGrandparent <> -1& Then
                lParentSibling = .GetNodeRow(.GetNodeRow(lParent, flexNTParent), flexNTNextSibling)
            Else
                lParentSibling = -1&
            End If
            
            ' If the parent has a next sibling, remove all child nodes between the last
            ' child of this parent and it's next sibling...
            If lSibling <> -1& Then
                For lIndex2 = lSibling - 1 To lLastRow + 1 Step -1
                    .RemoveItem lIndex2
                Next lIndex2
                
            ' Otherwise, if the parent's parent has a next sibling, remove all child nodes between the last
            ' child of this parent and it's parent's next sibling...
            ElseIf lParentSibling <> -1& Then
                For lIndex2 = lParentSibling - 1 To lLastRow + 1 Step -1
                    .RemoveItem lIndex2
                Next lIndex2
                
            ' Otherwise, remove all child nodes between the last child of this parent
            ' and the end of the grid...
            Else
                For lIndex2 = .Rows - 1 To lLastRow + 1 Step -1
                    .RemoveItem lIndex2
                Next lIndex2
            End If
        
        ' Otherwise, remove all child nodes between the last child of this parent
        ' and the end of the grid...
        Else
            For lIndex2 = .Rows - 1 To lLastRow + 1 Step -1
                .RemoveItem lIndex2
            Next lIndex2
        End If
        
        For lIndex = lFirstRow To lLastRow
            .IsSubtotal(lIndex) = False
        Next lIndex
        
        If m.bSortedDescending Then
            .Cell(flexcpSort, lFirstRow, lCol, lLastRow, lCol) = flexSortGenericDescending
        Else
            .Cell(flexcpSort, lFirstRow, lCol, lLastRow, lCol) = flexSortGenericAscending
        End If
        
        For lIndex = lFirstRow To lLastRow
            .IsSubtotal(lIndex) = True
        Next lIndex
        
        lIndex = lFirstRow
        If lLevel < 3 Then
            Do While lIndex <> -1&
                .AddItem vbTab & vbTab & "(blank)", lIndex + 1
                .IsSubtotal(lIndex + 1) = True
                .RowOutlineLevel(lIndex + 1) = lLevel + 1
                .IsCollapsed(lIndex) = flexOutlineCollapsed
                lIndex = .GetNodeRow(lIndex, flexNTNextSibling)
            Loop
        End If
        
        .FillStyle = flexFillSingle
        For lIndex = 0 To .Cols - 1
            .Select 0, lIndex
            If lIndex = m.lSortedCol Then
                If m.bSortedDescending Then
                    .CellPicture = frmMain.img16.ListImages("kSortedDownArrow").Picture
                Else
                    .CellPicture = frmMain.img16.ListImages("kSortedUpArrow").Picture
                End If
                .CellPictureAlignment = flexPicAlignLeftTop
                .PicturesOver = True
            Else
                .CellPicture = Nothing
            End If
        Next
    
        SetBackColors fgTree
        
        If Len(strSymbol) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, GDTreeCol(eGDTreeCol_Symbol)) = strSymbol Then
                    .Row = lIndex
                    .RowSel = lIndex
                    
                    lParentRow = .GetNodeRow(lIndex, flexNTParent)
                    If lParentRow <> -1 Then .TopRow = lParentRow
                    If lIndex < .TopRow Or lIndex > .BottomRow Then
                        .ShowCell lIndex, GDTreeCol(eGDTreeCol_Symbol)
                    End If
                    Exit For
                End If
            Next lIndex
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.SortTreeOnCol", eGDRaiseError_Raise
    
End Sub

Private Sub RefreshTreeFlags()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPoolRec As Long
    Dim lField As Long
    Dim lTblIndex As Long
    
    With fgTree
        For lIndex = .FixedRows To .Rows - 1
            If .RowIsVisible(lIndex) = True Then
                lTblIndex = gdGetNum(m.hSortedIndex, Val(.TextMatrix(lIndex, GDTreeCol(eGDTreeCol_TableIndex))))
                lPoolRec = gdGetTableNum(m.hSymbols, TblCol(eTblCol_PoolRec), lTblIndex)
                lField = CLng(Val(Parse(.ColData(GDTreeCol(eGDTreeCol_Flagged)), vbTab, 1)))
                If g.SymbolPool.ArrayTable(lField, lPoolRec) = 1 Then
                    CheckedCell(fgTree, lIndex, GDTreeCol(eGDTreeCol_Flagged)) = True
                Else
                    CheckedCell(fgTree, lIndex, GDTreeCol(eGDTreeCol_Flagged)) = False
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.RefreshTreeFlags", eGDRaiseError_Raise
    
End Sub

Private Sub GetSiblingNumber(ByVal lRow As Long, lNum As Long, lCount As Long)
On Error GoTo ErrSection:

    Dim lGridRow As Long
    
    lNum = 0&
    lCount = 0&
    lGridRow = fgTree.GetNodeRow(lRow, flexNTFirstSibling)
    Do While lGridRow <> -1
        If lGridRow <= lRow Then lNum = lNum + 1
        lCount = lCount + 1
        lGridRow = fgTree.GetNodeRow(lGridRow, flexNTNextSibling)
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.GetSiblingNumber", eGDRaiseError_Raise
    
End Sub

Private Sub ChangeFields(ByVal strDisplayFields As String)
On Error GoTo ErrSection:

    Dim lVirtualRow As Long             ' Current row in the virtual grid
    Dim lVirtualTopRow As Long          ' Current top row in the virtual grid
    Dim strTreeSymbol As String         ' Current symbol in the tree grid

    lVirtualRow = fgVirtual.Row
    lVirtualTopRow = fgVirtual.TopRow
    strTreeSymbol = fgTree.TextMatrix(fgTree.Row, GDTreeCol(eGDTreeCol_Symbol))
        
    m.SymbolGrid.InitGrid fgVirtual, strDisplayFields
    InitTreeGrid strDisplayFields
    LoadTreeGrid
        
    On Error Resume Next
    fgVirtual.TopRow = lVirtualTopRow
    fgVirtual.Row = lVirtualRow
    ShowTreeSymbol strTreeSymbol
        
    If m.Mode = eGDSymbolGridMode_Virtual Then
        MoveFocus fgVirtual
    Else
        MoveFocus fgTree
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ChangeFields", eGDRaiseError_Raise
    
End Sub

Private Sub ChangeList()
On Error GoTo ErrSection:

    If Not m.SymbolGrid Is Nothing Then ' And Me.Visible Then
        If cboList.SelectedItem.Key = "Sectors" Then
            If m.Mode = eGDSymbolGridMode_Virtual Then
                RefreshTreeFlags
                fgVirtual.Visible = False
                fgTree.Visible = True
                m.Mode = eGDSymbolGridMode_Tree
                MoveFocus fgTree
                
                If fgVirtual.Row >= fgVirtual.FixedRows And fgVirtual.Row < fgVirtual.Rows Then
                    ShowTreeSymbol fgVirtual.TextMatrix(fgVirtual.Row, kSymbolCol)
                End If
            
                If fgTree.RowSel >= fgTree.FixedRows And fgTree.RowSel < fgTree.Rows Then
                    m.strLastSymbol = Trim(fgTree.TextMatrix(fgTree.RowSel, GDTreeCol(eGDTreeCol_Symbol)))
                    If AutoChart Then
                        SetActiveChartSymbol Trim(fgTree.TextMatrix(fgTree.RowSel, GDTreeCol(eGDTreeCol_Symbol)))
                    End If
                    MoveFocus fgTree
                End If
            End If
        Else
            fgVirtual.Row = -1
            m.SymbolGrid.FilterID = cboList.SelectedItem.Key
            m.SymbolGrid.SortOnCol -1, 0
            
            If m.Mode = eGDSymbolGridMode_Tree Then
                fgTree.Visible = False
                fgVirtual.Visible = True
                m.Mode = eGDSymbolGridMode_Virtual
                If fgTree.Row >= fgTree.FixedRows And fgTree.Row < fgTree.Rows Then
                    ShowSymbol fgTree.TextMatrix(fgTree.Row, GDTreeCol(eGDTreeCol_Symbol))
                End If
            End If
            
            If fgVirtual.RowSel >= fgVirtual.FixedRows And fgVirtual.RowSel < fgVirtual.Rows Then
                m.strLastSymbol = Trim(fgVirtual.TextMatrix(fgVirtual.RowSel, kSymbolCol))
                If AutoChart Then
                    SetActiveChartSymbol Trim(fgVirtual.TextMatrix(fgVirtual.RowSel, kSymbolCol))
                End If
                MoveFocus fgVirtual
            End If
        End If
        
        SetComboTooltip
        UpdateFlagCount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ChangeList", eGDRaiseError_Raise
    
End Sub

Private Sub ShowFlags()
On Error GoTo ErrSection:

    fgTree.ColHidden(GDTreeCol(eGDTreeCol_Flagged)) = Not m.bShowFlags
    fgVirtual.ColHidden(kFlagCol) = Not m.bShowFlags
    fraFlags.Visible = m.bShowFlags
    Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ShowFlags", eGDRaiseError_Raise
    
End Sub

Private Sub ShowListSymbol(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim nRec&
    
    If g.SymbolPool.NumRecords > 0 And fgVirtual.Rows > fgVirtual.FixedRows Then
        If Len(strSymbol) > 0 Then
            If fgVirtual.Row >= 0 Then
                If strSymbol = fgVirtual.TextMatrix(fgVirtual.Row, kSymbolCol) Then
                    Exit Sub
                End If
            End If
            nRec = g.SymbolPool.PoolRecForSymbol(strSymbol, True)
        Else
            ' get initial symbol
            strSymbol = Trim(GetIniFileProperty("InitialSymbol", "", "Grid", g.strIniFile))
            If strSymbol = "" Or strSymbol = "$AD" Then strSymbol = "$DJIA"
            nRec = g.SymbolPool.PoolRecForSymbol(strSymbol, True)
            If nRec <= 0 Or nRec >= g.SymbolPool.NumRecords Then
                strSymbol = "$DJIA"
                nRec = g.SymbolPool.PoolRecForSymbol(strSymbol)
                If nRec < 0 Or nRec >= g.SymbolPool.NumRecords Then nRec = 0
            End If
        End If
        SymbolGrid.ShowRec nRec
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.ShowListSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameCriteria
'' Description: Handle a renamed criteria file (which also changes the ID)
'' Inputs:      Old Criteria ID, New Criteria ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RenameCriteria(ByVal strOldCriteriaID As String, ByVal strNewCriteriaID As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strDisplayFields As String      ' Display string for the grid
    Dim strDisplay As String            ' Display string for the grid
    Dim strSelection As String          ' Current selection in the combo

    ' Grab the current selection in the combo box...
    strSelection = SelectionKey
    
    ' Update the display fields if necessary...
    strDisplayFields = GetIniFileProperty("DisplayFields", "", "Grid", g.strIniFile)
    strDisplay = Replace(strDisplayFields, ":" & UCase(strOldCriteriaID) & "|", ":" & UCase(strNewCriteriaID) & "|")
    
    If strDisplay <> strDisplayFields Then
        SetIniFileProperty "DisplayFields", strDisplay, "Grid", g.strIniFile
        InitForm
    Else
        LoadCombo
    End If
    
    ' If the previous selection was the old criteria ID, select the new criteria ID...
    If UCase(Parse(strSelection, ":", 2)) = UCase(strOldCriteriaID) Then
        For lIndex = 1 To cboList.ComboItems.Count
            If UCase(Parse(cboList.ComboItems(lIndex).Key, ":", 2)) = UCase(strNewCriteriaID) Then
                cboList.ComboItems(lIndex).Selected = True
                Exit For
            End If
        Next lIndex
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGrid.RenameCriteria"
    
End Sub

