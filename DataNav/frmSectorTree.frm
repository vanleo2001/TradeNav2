VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmSectorTree 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7LCtl.VSFlexGrid fgSymbols 
      Height          =   2895
      Left            =   120
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
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuAddToQuoteBoard 
         Caption         =   "Add to &Quote Board"
      End
      Begin VB.Menu mnuAddToSymbolGroup 
         Caption         =   "Add to &Symbol Group"
         Begin VB.Menu mnuSymGroups 
            Caption         =   " (new Symbol Group)"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "E&xport Sector Tree"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmSectorTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSectorTree.frm
'' Description: Show the user sectors, sub-sectors, and stocks in a tree format
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80917
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Symbol = 0
    eGDCol_Description
    eGDCol_SymbolID
    eGDCol_TableIndex
    eGDCol_NumCols
End Enum

Private Enum eTblCols
    eTblCol_Symbol = 0
    eTblCol_SymbolID
    eTblCol_Description
    eTblCol_Level
    eTblCol_SortKey
    eTblCol_NumCols
End Enum

Private Type mPrivate
    tblSymbols As cGdTable
    hSymbols As Long
    aSortedIndex As cGdArray
    hSortedIndex As Long
    
    bChangeChart As Boolean
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(Optional ByVal strSymbol As String = "", Optional ByVal bJustExport As Boolean = False)
On Error GoTo ErrSection:

    Set m.tblSymbols = New cGdTable
    With m.tblSymbols
        .CreateField eGDARRAY_Longs, TblCol(eTblCol_SymbolID), "SymbolID"
        .CreateField eGDARRAY_Strings, TblCol(eTblCol_Symbol), "Symbol"
        .CreateField eGDARRAY_Strings, TblCol(eTblCol_Description), "Description"
        .CreateField eGDARRAY_TinyInts, TblCol(eTblCol_Level), "Level"
        .CreateField eGDARRAY_Strings, TblCol(eTblCol_SortKey), "SortKey"
    End With
    m.hSymbols = m.tblSymbols.TableHandle

    Screen.MousePointer = vbHourglass
    InitGrid
    LoadSymbols
    LoadGrid
    If Len(strSymbol) > 0 Then HighlightSymbol strSymbol
    Screen.MousePointer = vbDefault

    If bJustExport Then
        ExportSectorTree
    Else
        ShowForm Me, False, frmMain, , ALT_GRID_ROW_COLOR
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmSectorTree.ShowMe", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Bring up the print preview form so the user can print the grid
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection:

    If fgSymbols.Rows > 550 Then
        If AskBox("h=Warning ; i=? ; b=+Yes|-No ; Printing this many symbols may take a while.||Do you want to continue?") = "N" Then
            Exit Function
        End If
    End If

    PrintMe = frmPrintPreview.ShowMe("CNV SectorTree", Me)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSectorTree.PrintMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_AfterCollapse
'' Description: After expanding or collapsing, redo the background colors
'' Inputs:      Row expanded/collapsed, State
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
On Error GoTo ErrSection:

    SetBackColors fgSymbols

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.fgSymbols.AfterCollapse", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_AfterRowColChange
'' Description: After a row changes, change the chart if we need to
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If m.bChangeChart = True Then
        SetActiveChartSymbol CLng(fgSymbols.TextMatrix(NewRow, GDCol(eGDCol_SymbolID)))
        m.bChangeChart = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.fgSymbols.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_BeforeCollapse
'' Description: Before expanding, see if we need to fill in the children
'' Inputs:      Row expanded/collapsed, State, Whether to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lParentRow As Long              ' Parent row of the current row

    If State = flexOutlineExpanded Then
        With fgSymbols
            lRedraw = .Redraw
            .Redraw = flexRDNone

            ' Collapse all branches except for this one...
            CollapseAll
            lParentRow = .GetNodeRow(Row, flexNTParent)
            If lParentRow <> -1 Then .IsCollapsed(lParentRow) = flexOutlineExpanded
            
            ' Fill in the leaves if we need to...
            If .TextMatrix(Row + 1, GDCol(eGDCol_Symbol)) = "(blank)" Then
                ExpandRow Row
            End If
            
            ' If the bottom node is off the screen, make this the top row...
            If .GetNodeRow(Row, flexNTLastChild) > .BottomRow Then
                .TopRow = .Row
            End If
            
            .Redraw = lRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.fgSymbols.BeforeCollapse", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_BeforeMouseDown
'' Description: If the user clicks on the right button, bring up the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Mouse Location,
''              Whether to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row that the user clicked on
    Dim lMouseCol As Long               ' Column that the user clicked on
    Dim strSymbol As String             ' Symbol of the row the user clicked on
    Dim astrItems As New cGdArray       ' Array of symbol group items
    Dim lIndex As Long                  ' Index into a for loop

    ' Grab the row and column of the cell that the user clicked on...
    With fgSymbols
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            strSymbol = .TextMatrix(.RowSel, GDCol(eGDCol_Symbol))
        Else
            strSymbol = ""
        End If
    End With

    ' If the user pressed the right button, bring up the pop-up menu...
    If Button = vbRightButton Then
        If Len(strSymbol) > 0 Then
            mnuAddToQuoteBoard.Visible = True
            mnuAddToSymbolGroup.Visible = True
            mnuSep.Visible = True
            
            mnuAddToQuoteBoard.Caption = "Add " & strSymbol & " to &Quote Board"
            mnuAddToQuoteBoard.Enabled = (fgSymbols.RowOutlineLevel(lMouseRow) = 3)
            mnuAddToSymbolGroup.Caption = "Add " & strSymbol & " to &Symbol Group"
    
            ' Assemble Symbol Group list for menu...
            astrItems.Clear
            astrItems.Add " (new Symbol Group)"
            For lIndex = 1 To g.SymbolPool.SymbolGroups.Count
                With g.SymbolPool.SymbolGroups.Item(lIndex)
                    If .Custom And Len(.Name) > 0 Then
                        astrItems.Add .Name & vbTab & .ID
                    End If
                End With
            Next
            astrItems.Sort eGdSort_IgnoreCase
            
            ' Add menu item for each symbol group...
            For lIndex = 0 To astrItems.Size - 1
                If lIndex > mnuSymGroups.UBound Then
                    Load mnuSymGroups(lIndex)
                    mnuSymGroups(lIndex).Visible = True
                End If
                mnuSymGroups(lIndex).Caption = Parse(astrItems(lIndex), vbTab, 1)
                mnuSymGroups(lIndex).Tag = Parse(astrItems(lIndex), vbTab, 2)
            Next
            
            ' Remove extras...
            For lIndex = mnuSymGroups.UBound To astrItems.Size Step -1
                If lIndex > 0 Then Unload mnuSymGroups(lIndex)
            Next
        Else
            mnuAddToQuoteBoard.Visible = False
            mnuAddToSymbolGroup.Visible = False
            mnuSep.Visible = False
        End If

        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.fgSymbols.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_DblClick
'' Description: Set the active chart symbol on a double click
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_DblClick()
On Error GoTo ErrSection:
    
    Dim lMouseRow As Long               ' Row that the user clicked on
    
    With fgSymbols
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            SetActiveChartSymbol CLng(.TextMatrix(lMouseRow, GDCol(eGDCol_SymbolID)))
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.fgSymbols.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_KeyDown
'' Description: Send any key strokes to the fgKeyDown routine
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    With fgSymbols
        Select Case KeyCode
            Case vbKeyRight
                KeyCode = 0
                If .RowOutlineLevel(.Row) < 3 Then
                    If .IsCollapsed(.Row) = flexOutlineExpanded Then
                        If .RowOutlineLevel(.Row + 1) >= .RowOutlineLevel(.Row) Then
                            m.bChangeChart = True
                            .Row = .Row + 1
                            .ShowCell .Row, GDCol(eGDCol_Symbol)
                        End If
                    Else
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    End If
                ElseIf .Row + 1 < .Rows Then
                    If .RowOutlineLevel(.Row + 1) >= .RowOutlineLevel(.Row) Then
                        m.bChangeChart = True
                        .Row = .Row + 1
                        .ShowCell .Row, GDCol(eGDCol_Symbol)
                    End If
                End If
            
            Case vbKeyLeft
                KeyCode = 0
                If .RowOutlineLevel(.Row) = 3 Then
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then
                        m.bChangeChart = True
                        .Row = .GetNodeRow(.Row, flexNTParent)
                        .ShowCell .Row, GDCol(eGDCol_Symbol)
                    End If
                ElseIf .IsCollapsed(.Row) = flexOutlineExpanded Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then
                        m.bChangeChart = True
                        .Row = .GetNodeRow(.Row, flexNTParent)
                        .ShowCell .Row, GDCol(eGDCol_Symbol)
                    End If
                End If
            
            Case vbKeyUp
                m.bChangeChart = True
            
            Case vbKeyDown
                m.bChangeChart = True
        
        End Select
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.fgSymbols.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: If the user presses a function key, pass it on to KeyPress
'' Inputs:      Code of the Key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    RaiseError "frmSectorTree.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyPress
'' Description: If the user presses a key, pass it on to the global KeyPress
'' Inputs:      Ascii version of the Key pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.Form.KeyPress", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize form stuff upon loading
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String               ' Grid font information from ini file
    Dim strPlacement As String          ' Form placement information

    g.Styler.StyleForm Me
    
    Caption = "Sector Browser"
    Me.Icon = Picture16(ToolbarIcon("ID_SectorBrowser"), , True)
    mnuPopUp.Visible = False

    strPlacement = GetIniFileProperty("SectorTree", "", "Placement", g.strIniFile)
    If Len(strPlacement) > 0 Then SetFormPlacement Me, strPlacement, "LHT"

    strFont = GetIniFileProperty("SectorTree", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgSymbols.Font, strFont

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move/Resize the controls on the form as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    With fgSymbols
        .Move 0, 0, ScaleWidth, ScaleHeight
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up after ourselves before unloading
'' Inputs:      Whether to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "SectorTree", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "SectorTree", FontToString(fgSymbols.Font), "Fonts", g.strIniFile
    
    Set m.tblSymbols = Nothing
    m.hSymbols = 0
    
    frmMain.tbToolbar.Tools("ID_SectorBrowser").State = ssUnchecked

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
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

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgSymbols
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        SetupGrid fgSymbols, eGridMode_Tree
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbApplicationWorkspace
        .GridLines = flexGridNone
        .OutlineBar = flexOutlineBarSimpleLeaf
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        .Rows = 1
        .FixedRows = 1
        
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Description)) = "Description"
        
        .ColHidden(GDCol(eGDCol_SymbolID)) = True
        .ColHidden(GDCol(eGDCol_TableIndex)) = True
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.InitGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid from the memory tables
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lTblIndex As Long               ' Index into the sectors table
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgSymbols
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 0 To gdGetSize(m.hSortedIndex) - 1
            lTblIndex = gdGetNum(m.hSortedIndex, lIndex)
            
            If gdGetTableNum(m.hSymbols, eTblCol_Level, lTblIndex) = 0 Then
                .AddItem gdGetTableString(m.hSymbols, eTblCol_Symbol, lTblIndex) & vbTab & gdGetTableString(m.hSymbols, eTblCol_Description, lTblIndex) & vbTab & Str(gdGetTableNum(m.hSymbols, eTblCol_SymbolID, lTblIndex)) & vbTab & Str(lIndex)
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 1
                
                .AddItem "(blank)"
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 2
            End If
        Next lIndex
        
        .Outline -1
        .AutoSize 0, .Cols - 1, False, 75
        .Outline 2
        .Outline 1
        
        SetBackColors fgSymbols

        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.LoadGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function to set up the print preview form
'' Inputs:      Arguments
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgSymbols
        Else
            .RenderControl = fgSymbols.hWnd
        End If
        
        .EndDoc
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.GenerateReport", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAddToQuoteBoard_Click
'' Description: Allow the user to add a stock to the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddToQuoteBoard_Click()
On Error GoTo ErrSection:

    Dim lPoolRec As Long                ' Record of the symbol in the symbol pool

    With fgSymbols
        If .RowOutlineLevel(.RowSel) = 3 Then
            lPoolRec = g.SymbolPool.PoolRecForSymbolID(CLng(.TextMatrix(.RowSel, GDCol(eGDCol_SymbolID))))
            frmQuotes.AddSymbol lPoolRec, "Daily"
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.mnuAddToQuoteBoard_Click"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change the font on the grid
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
    RaiseError "frmSectorTree.mnuChangeFont_Click"
    Resume ErrExit
    
End Sub

Private Sub mnuExport_Click()
On Error GoTo ErrSection:

    Dim strFile$
    
    strFile = ExportSectorTree
    If Len(strFile) > 0 Then
        InfBox "Sector Tree info has been saved to|" & strFile, "i", , "Export Sector Tree"
    Else
        InfBox "Error exporting sector tree.", "e", , "Export Sector Tree"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.mnuExport_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSymGroups_Click
'' Description: Allow the user to add selected symbol to a symbol group
'' Inputs:      Index of the Chosen Symbol Group
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSymGroups_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim frm As New frmSymbolGroup       ' Symbol group form
    Dim alToAdd As New cGdArray         ' Array of symbols to add the symbol group
    
    ' Set up array of symbols to add to the symbol group...
    alToAdd.Create eGDARRAY_Longs
    alToAdd.Add CLng(fgSymbols.TextMatrix(fgSymbols.RowSel, GDCol(eGDCol_SymbolID)))
    
    ' Show the symbol group form and add the selected symbols...
    frm.ShowMe AddSlash(App.Path) & "Custom\", mnuSymGroups(Index).Tag, False, alToAdd

ErrExit:
    Set alToAdd = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.mnuSymGroups_Click"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSymbols
'' Description: Load the symbols into the memory table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSymbols()
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
    
    If SU_GetGroupChildren(0&, alSectors) Then
        m.tblSymbols.NumRecords = 20000
        For lSector = 0 To alSectors.Size - 1
            lPoolRec = g.SymbolPool.PoolRecForSymbolID(alSectors(lSector))
            strSectorSym = g.SymbolPool.Symbol(lPoolRec)
            If Len(strSectorSym) > 0 Then
                strDesc = g.SymbolPool.Desc(lPoolRec)
                gdSetTableNum m.hSymbols, eTblCol_SymbolID, lTblIndex, gdGetNum(hSectors, lSector)
                gdSetTableStr m.hSymbols, eTblCol_Symbol, lTblIndex, strSectorSym
                gdSetTableStr m.hSymbols, eTblCol_Description, lTblIndex, strDesc
                gdSetTableNum m.hSymbols, eTblCol_Level, lTblIndex, 0
                gdSetTableStr m.hSymbols, eTblCol_SortKey, lTblIndex, Pad(strSectorSym, 14, "L") & Pad("", 14, "L") & Pad("", 14, "L")
                lTblIndex = lTblIndex + 1
                
                If SU_GetGroupChildren(alSectors(lSector), alSubSectors) Then
                    For lSubSector = 0 To alSubSectors.Size - 1
                        lPoolRec = g.SymbolPool.PoolRecForSymbolID(alSubSectors(lSubSector))
                        strSubSectorSym = g.SymbolPool.Symbol(lPoolRec)
                        If Len(strSubSectorSym) > 0 Then
                            strDesc = g.SymbolPool.Desc(lPoolRec)
                            gdSetTableNum m.hSymbols, eTblCol_SymbolID, lTblIndex, gdGetNum(hSubSectors, lSubSector)
                            gdSetTableStr m.hSymbols, eTblCol_Symbol, lTblIndex, strSubSectorSym
                            gdSetTableStr m.hSymbols, eTblCol_Description, lTblIndex, strDesc
                            gdSetTableNum m.hSymbols, eTblCol_Level, lTblIndex, 1
                            gdSetTableStr m.hSymbols, eTblCol_SortKey, lTblIndex, Pad(strSectorSym, 14, "L") & Pad(strSubSectorSym, 14, "L") & Pad("", 14, "L")
                            lTblIndex = lTblIndex + 1
                            
                            If SU_GetGroupChildren(alSubSectors(lSubSector), alStocks) Then
                                For lStock = 0 To alStocks.Size - 1
                                    lPoolRec = g.SymbolPool.PoolRecForSymbolID(alStocks(lStock))
                                    strSymbol = g.SymbolPool.Symbol(lPoolRec)
                                    If Len(strSymbol) > 0 Then
                                        strDesc = g.SymbolPool.Desc(lPoolRec)
                                        gdSetTableNum m.hSymbols, eTblCol_SymbolID, lTblIndex, gdGetNum(hStocks, lStock)
                                        gdSetTableStr m.hSymbols, eTblCol_Symbol, lTblIndex, strSymbol
                                        gdSetTableStr m.hSymbols, eTblCol_Description, lTblIndex, strDesc
                                        gdSetTableNum m.hSymbols, eTblCol_Level, lTblIndex, 2
                                        gdSetTableStr m.hSymbols, eTblCol_SortKey, lTblIndex, Pad(strSectorSym, 14, "L") & Pad(strSubSectorSym, 14, "L") & Pad(strSymbol, 14, "L")
                                        lTblIndex = lTblIndex + 1
                                    End If
                                Next lStock
                            End If
                        End If
                    Next lSubSector
                End If
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
    RaiseError "frmSectorTree.LoadSymbols", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HighlightSymbol
'' Description: Highlight the line with the given symbol
'' Inputs:      Symbol to highlight
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HighlightSymbol(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strSector As String             ' Symbol of the sector
    Dim strSubsector As String          ' Symbol of the subsector
    Dim lRow As Long                    ' Row where we found what we were looking for
    Dim lParentRow As Long              ' Row of the parent node
    
    With fgSymbols
        For lIndex = 0 To m.tblSymbols.NumRecords - 1
            If TableStr(eTblCol_Symbol, lIndex) = strSymbol Then
                strSector = Trim(Left(TableStr(eTblCol_SortKey, lIndex), 14))
                strSubsector = Trim(Mid(TableStr(eTblCol_SortKey, lIndex), 15, 14))
                Exit For
            End If
        Next lIndex
        
        If Len(strSector) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, GDCol(eGDCol_Symbol)) = strSector Then
                    lRow = lIndex
                    .IsCollapsed(lIndex) = flexOutlineExpanded
                    Exit For
                End If
            Next lIndex
        
            If Len(strSubsector) > 0 Then
                For lIndex = lRow To .Rows - 1
                    If .TextMatrix(lIndex, GDCol(eGDCol_Symbol)) = strSubsector Then
                        lRow = lIndex
                        .IsCollapsed(lIndex) = flexOutlineExpanded
                        Exit For
                    End If
                Next lIndex
            
                If Left(strSymbol, 1) <> "$" Then
                    For lIndex = lRow To .Rows - 1
                        If .TextMatrix(lIndex, GDCol(eGDCol_Symbol)) = strSymbol Then
                            lRow = lIndex
                            .Row = lIndex
                            .RowSel = lIndex
                            
                            lParentRow = .GetNodeRow(lIndex, flexNTParent)
                            If lParentRow <> -1 Then .TopRow = lParentRow
                            If lIndex < .TopRow Or lIndex > .BottomRow Then
                                .ShowCell lIndex, GDCol(eGDCol_Symbol)
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
    RaiseError "frmSectorTree.HighlightSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KeyPress
'' Description: Perform an action based on the key the user pressed
'' Inputs:      Key Pressed, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub KeyPress(KeyAscii As Integer, Optional Shift As Integer = -1)
On Error Resume Next

    Dim frm As Form                     ' Charting Form
    Dim bLookForChart As Boolean        ' Should we look for the chart?
    Dim astrSymbols As New cGdArray     ' Symbol back from symbol selector

    If KeyAscii = 0 Then Exit Sub

    If Shift >= 0 Then ' (came from KeyDown event)
        If KeyAscii >= vbKeyF2 And KeyAscii <= vbKeyF12 Then
            bLookForChart = True
        End If
    Else ' (came from KeyPress event)
        Select Case Asc(UCase(Chr(KeyAscii)))
            
            Case 83:        ' S
                Set astrSymbols = frmSymbolSelector.ShowMe("", False)
                If astrSymbols.Size > 0 Then
                    HighlightSymbol astrSymbols(0)
                End If
                KeyAscii = 0
                
            Case 65 To 90, 48 To 57, 43, 45, 61:
                bLookForChart = True
        End Select
    End If
       
    If bLookForChart Then
        Set frm = ActiveChart
        If Not frm Is Nothing Then
            frm.KeyPress KeyAscii, Shift
        End If
        KeyAscii = 0
    End If
       
    Set frm = Nothing
    Set astrSymbols = Nothing

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
    
    With fgSymbols
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        lStart = CLng(.TextMatrix(lRow, GDCol(eGDCol_TableIndex)))
        lLevel = .RowOutlineLevel(lRow) - 1
        
        .RemoveItem lRow + 1
        lInsertRow = lRow + 1
        
        For lIndex = lStart + 1 To m.tblSymbols.NumRecords - 1
            lTblIndex = gdGetNum(m.hSortedIndex, lIndex)
            If TableNum(eTblCol_Level, lTblIndex) = lLevel Then Exit For
            
            If TableNum(eTblCol_Level, lTblIndex) = lLevel + 1 Then
                .AddItem gdGetTableString(m.hSymbols, eTblCol_Symbol, lTblIndex) & vbTab & gdGetTableString(m.hSymbols, eTblCol_Description, lTblIndex) & vbTab & Str(gdGetTableNum(m.hSymbols, eTblCol_SymbolID, lTblIndex)) & vbTab & Str(lIndex), lInsertRow
                .IsSubtotal(lInsertRow) = True
                .RowOutlineLevel(lInsertRow) = TableNum(eTblCol_Level, lTblIndex) + 1
                lInsertRow = lInsertRow + 1
                
                If TableNum(eTblCol_Level, lTblIndex) < 2 Then
                    .AddItem "(blank)", lInsertRow
                    .IsSubtotal(lInsertRow) = True
                    .RowOutlineLevel(lInsertRow) = TableNum(eTblCol_Level, lTblIndex) + 2
                    lInsertRow = lInsertRow + 1
                    .IsCollapsed(lInsertRow - 2) = flexOutlineCollapsed
                End If
            End If
        Next lIndex
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.ExpandRow", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CollapseAll
'' Description: Collapse all rows
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CollapseAll()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgSymbols
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If .RowOutlineLevel(lIndex) < 3 Then
                .IsCollapsed(lIndex) = flexOutlineCollapsed
            End If
        Next lIndex
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.CollapseAll", eGDRaiseError_Raise
    
End Sub

Private Function ExportSectorTree() As String
On Error GoTo ErrSection:

    Dim i&, nTblIndex&, nLevel&, nSymbolID&
    Dim s$, strSymbol$, strDesc$
    Dim aOut As New cGdArray

    ' export the tree
    For i = 0 To gdGetSize(m.hSortedIndex) - 1
        nTblIndex = gdGetNum(m.hSortedIndex, i)
        nLevel = gdGetTableNum(m.hSymbols, eTblCol_Level, nTblIndex)
        nSymbolID = gdGetTableNum(m.hSymbols, eTblCol_SymbolID, nTblIndex)
        strSymbol = gdGetTableString(m.hSymbols, eTblCol_Symbol, nTblIndex)
        strDesc = gdGetTableString(m.hSymbols, eTblCol_Description, nTblIndex)
        If Len(strSymbol) < 6 Then
            strSymbol = strSymbol & Space(6 - Len(strSymbol))
        End If
        s = Space(nLevel) & Str(nLevel) & " " & vbTab & strSymbol & vbTab & strDesc
        aOut.Add s
    Next
    s = FilePath(App.Path) ' e.g. "C:\Genesis\"
    s = AddSlash(s) & "SectorTree.txt"
    aOut.ToFile s
    ExportSectorTree = s
    
    If FileExist(App.Path & "\SectorExport.flg") Then
        ExportSectorInfo
    End If
    
ErrExit:
    Set aOut = Nothing
    Exit Function
    
ErrSection:
    RaiseError "frmSectorTree.ExportSectorTree"
    Resume ErrExit
End Function

Private Sub ExportSectorInfo()
On Error GoTo ErrSection:

    Dim rc&, i&, tbl&, iRec&, nSymbolID&, nSectorID&, nSubsectorID&, dValue#, nDate&
    Dim s$, strSymbol$, strSecType$, strDesc$, strSector$, strSubsector$
    Dim aOut As New cGdArray
    
    ' walk through the Symbols table
    tbl = g.Universe.tblSymbols
    If tbl = 0 Then Exit Sub
    rc = TagSelect(tbl, g.Universe.tagSymbolID)
    rc = d4top(tbl)
    Do While rc = r4success
        iRec = d4recNo(tbl)
        nSymbolID = f4long(g.Universe.fldSymbolID)
        'strSecType = UCase(Trim(f4str(g.Universe.fldSecType)))
        strSymbol = UCase(Trim(f4str(g.Universe.fldSymbol)))
        If Len(strSymbol) > 0 Then
            'strDesc = Trim(f4str(g.Universe.fldDesc))
            nSectorID = 0
            nSubsectorID = 0
            ' get symbol for sector
            If DM_GetSnap1(g.DMS, nSymbolID, 162, dValue, nDate) Then
                nSectorID = dValue
            End If
            If DM_GetSnap1(g.DMS, nSymbolID, 163, dValue, nDate) Then
                nSubsectorID = dValue
            End If
            If nSectorID > 0 Or nSubsectorID > 0 Then
                strSector = GetSymbol(nSectorID)
                strSubsector = GetSymbol(nSubsectorID)
                s = strSymbol & vbTab & Str(nSymbolID) & vbTab & strSector & vbTab & Str(nSectorID) & vbTab & strSubsector & vbTab & Str(nSubsectorID)
                If InStr(strSymbol, "@") > 0 Or Left(strSector, 3) <> "$--" Or Left(strSubsector, 2) <> "$-" Or SecurityType(strSymbol) <> "S" Then
                    s = s & vbTab & "###"
                    'AddList s
                End If
                aOut.Add s
            End If
        End If
        If iRec <> d4recNo(tbl) Then
            rc = d4go(tbl, iRec)
            If iRec <> d4recNo(tbl) Then
                rc = 0
            End If
        End If
        rc = d4skip(tbl, 1)
    Loop
   
    s = FilePath(App.Path) ' e.g. "C:\Genesis\"
    s = AddSlash(s) & "Sectors.txt"
    aOut.ToFile s
   
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSectorTree.ExportSectorInfo"
End Sub


