VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSelectFilter 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7LCtl.VSFlexGrid fgFilters 
      Height          =   2895
      Left            =   180
      TabIndex        =   0
      Top             =   120
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
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1815
      Left            =   4620
      TabIndex        =   1
      Top             =   60
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
      Caption         =   "frmSelectFilter.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSelectFilter.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSelectFilter.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdClearFilter 
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
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
         Caption         =   "frmSelectFilter.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSelectFilter.frx":00A2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSelectFilter.frx":00C2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   540
         Width           =   1215
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
         Caption         =   "frmSelectFilter.frx":00DE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSelectFilter.frx":010C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSelectFilter.frx":012C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1215
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
         Caption         =   "frmSelectFilter.frx":0148
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSelectFilter.frx":016E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSelectFilter.frx":018E
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmSelectFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSelectFilter.frm
'' Description: Allow the user to select a symbol group, filter, or boolean
''              criteria to add to the filter tab of the quote board
''
'' Author:      Genesis Financial Data Services
''              425 Wind Chime Pl
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_ID = 0
    eGDCol_Name
    eGDCol_Type
    eGDCol_NumCols
End Enum

Private Enum eGDButtons
    eGDButton_OK = 0
    eGDButton_Cancel
    eGDButton_ClearFilter
End Enum

Private Type mPrivate
    nButton As eGDButtons               ' Button the user pressed
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Current Filter ID to select
'' Returns:     Selected Filter ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal strFilterID As String) As String
On Error GoTo ErrSection:

    InitGrid
    LoadGrid strFilterID
    
    MoveFocus fgFilters
    
    cmdClearFilter.Visible = (Len(strFilterID) > 0)

    ShowForm Me, eForm_Modal, frmMain
    
    Select Case m.nButton
        Case eGDButton_OK
            ShowMe = fgFilters.TextMatrix(fgFilters.Row, GDCol(eGDCol_ID))
        Case eGDButton_ClearFilter
            ShowMe = "<clear>"
        Case eGDButton_Cancel
            ShowMe = ""
    End Select

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmSelectFilter.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow ShowMe to unload the form without saving information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.nButton = eGDButton_Cancel
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelectFilter.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClearFilter_Click
'' Description: Allow ShowMe to unload the form and notify to clear the filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearFilter_Click()
On Error GoTo ErrSection:

    m.nButton = eGDButton_ClearFilter
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelectFilter.cmdClearFilter_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allow ShowMe to unload the form and save the information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.nButton = eGDButton_OK
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelectFilter.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilters_DblClick
'' Description: If the user double clicks on an item, select it and OK it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilters_DblClick()
On Error GoTo ErrSection:

    With fgFilters
        .Row = .MouseRow
        .RowSel = .Row
    End With

    m.nButton = eGDButton_OK
    Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelectFilter.fgFilters_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilters_KeyPress
'' Description: If the user presses enter, OK the selection
'' Inputs:      Ascii version of Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilters_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        m.nButton = eGDButton_OK
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelectFilter.fgFilters_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Caption = "Filter Selection"
    Icon = Picture16("kBlank")
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelectFilter.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the "X", allow ShowMe to unload the form
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.nButton = eGDButton_Cancel
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelectFilter.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinScaleWidth As Long          ' Minimum scale width
    Dim lMinScaleHeight As Long         ' Minimum scale height
    
    lMinScaleWidth = fraButtons.Width * 5
    lMinScaleHeight = fraButtons.Height + 120
    
    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = True Then Exit Sub
    
    With fraButtons
        .Move ScaleWidth - .Width - 60, 60
    End With
    
    With fgFilters
        .Move 60, 60, ScaleWidth - fraButtons.Width - 180, ScaleHeight - 120
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgFilters
        .Redraw = flexRDNone
        
        SetupGrid fgFilters, eGridMode_List
        
        .FixedRows = 0
        .Rows = 0
        .FixedCols = 0
        .Cols = GDCol(eGDCol_NumCols)
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelectFilter.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load grid with symbol groups, filters, and boolean criteria
'' Inputs:      Filter ID to select
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(ByVal strFilterID As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRowSel As Long                 ' Row to select in the grid

    With fgFilters
        .Redraw = flexRDNone
        
        lRowSel = 0&
        
        ' Load up the symbol groups...
        For lIndex = 1 To g.SymbolPool.SymbolGroups.Count
            If (g.SymbolPool.SymbolGroups(lIndex).GroupType <> eGROUP_Builtin) And (g.SymbolPool.SymbolGroups(lIndex).GroupType <> eGROUP_Flag) Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = "GRP:" & g.SymbolPool.SymbolGroups(lIndex).ID
                .Cell(flexcpPicture, .Rows - 1, GDCol(eGDCol_Name)) = Picture16(ToolbarIcon("ID_SymbolGroups"))
                .Cell(flexcpPictureAlignment, .Rows - 1, GDCol(eGDCol_Name)) = flexAlignLeftCenter
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = g.SymbolPool.SymbolGroups(lIndex).Name
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = "Symbol Group"
                If .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = strFilterID Then lRowSel = .Rows - 1
            End If
        Next lIndex
        
        ' Load up the active filters...
        For lIndex = 1 To g.SymbolPool.Filters.Count
            If g.SymbolPool.Filters(lIndex).IsActive = True Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = "FIL:" & g.SymbolPool.Filters(lIndex).ID
                .Cell(flexcpPicture, .Rows - 1, GDCol(eGDCol_Name)) = Picture16(ToolbarIcon("ID_Filters"))
                .Cell(flexcpPictureAlignment, .Rows - 1, GDCol(eGDCol_Name)) = flexAlignLeftCenter
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = g.SymbolPool.Filters(lIndex).Name
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = "Filter"
                If .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = strFilterID Then lRowSel = .Rows - 1
            End If
        Next lIndex
        
        ' Load up the boolean criteria...
        For lIndex = 1 To g.SymbolPool.Criterias.Count
            If g.SymbolPool.Criterias(lIndex).IsActive = True And g.SymbolPool.Criterias(lIndex).IsBoolean = True Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = "DSV:" & g.SymbolPool.Criterias(lIndex).ID
                .Cell(flexcpPicture, .Rows - 1, GDCol(eGDCol_Name)) = Picture16(ToolbarIcon("ID_Criteria"))
                .Cell(flexcpPictureAlignment, .Rows - 1, GDCol(eGDCol_Name)) = flexAlignLeftCenter
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = g.SymbolPool.Criterias(lIndex).Name
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = "Criteria"
                If .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = strFilterID Then lRowSel = .Rows - 1
            End If
        Next lIndex
        
        .Col = GDCol(eGDCol_Name)
        .Sort = flexSortStringAscending
        
        ' Search for the row to select...
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_ID)) = strFilterID Then
                lRowSel = lIndex
                Exit For
            End If
        Next lIndex
        
        .Row = lRowSel
        .RowSel = lRowSel
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
        .ShowCell lRowSel, GDCol(eGDCol_Name)
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelectFilter.LoadGrid"
    
End Sub

