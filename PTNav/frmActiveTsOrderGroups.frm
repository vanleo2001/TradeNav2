VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmActiveTsOrderGroups 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMenu 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4140
      Top             =   2520
   End
   Begin VSFlex7LCtl.VSFlexGrid fgGroups 
      Height          =   2895
      Left            =   900
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
      Begin VB.Menu mnuManage 
         Caption         =   "Manage"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubmitGroup 
         Caption         =   "Submit"
      End
      Begin VB.Menu mnuCancelGroup 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuParkGroup 
         Caption         =   "Park"
      End
   End
End
Attribute VB_Name = "frmActiveTsOrderGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmActiveTsOrderGroups.frm
'' Description: Form to show user what groups are currently active or parked
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/15/2010   DAJ         Made several user interface changes
'' 06/24/2010   DAJ         Change to tree format with orders as subnodes
'' 08/09/2010   DAJ         Added check in GroupRow that RowData is correct type
'' 08/11/2010   DAJ         Fixed ToolTipText to display if in Name column
'' 08/12/2010   DAJ         Display order text in the action column
'' 08/18/2010   DAJ         Cancel confirmation for orders and groups
'' 06/28/2011   DAJ         Setup clickable cells like hyperlinks
'' 04/30/2013   DAJ         Employ Tim's fix for grid scrolling vs. streaming issue
'' 05/21/2013   DAJ         Changed the confirmations when cancelling a TSOG
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Name = 0
    eGDCol_Symbol
    eGDCol_Account
    eGdCol_Quantity
    eGDCol_Status
    eGDCol_Cancel
    eGDCol_Action
    eGDCol_NumCols
End Enum

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddGroup
'' Description: Add the given group to the grid
'' Inputs:      Group
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddGroup(ByVal tsoGroup As cActiveTsOrderGroup)
On Error GoTo ErrSection:

    GroupToGrid tsoGroup
    FilterGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.AddGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateGroup
'' Description: Update the given group in the grid
'' Inputs:      Group
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateGroup(ByVal tsoGroup As cActiveTsOrderGroup)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid to update
    Dim grdGroup As cActiveTsOrderGroup ' Group out of the grid
    
    With fgGroups
        lRow = -1&
        For lIndex = .FixedRows To .Rows - 1
            If IsGroupRow(lIndex) Then
                Set grdGroup = .RowData(lIndex)
                If grdGroup.Key = tsoGroup.Key Then
                    lRow = lIndex
                    Exit For
                End If
            End If
        Next lIndex
        
        GroupToGrid tsoGroup, lRow
        FilterGrid
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.UpdateGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateGroupOrder
'' Description: Update the given group and order in the grid
'' Inputs:      Group
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateGroupOrder(ByVal tsoGroup As cActiveTsOrderGroup, ByVal lOrderNumber As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid to update
    Dim grdGroup As cActiveTsOrderGroup ' Group out of the grid
    
    With fgGroups
        lRow = -1&
        For lIndex = .FixedRows To .Rows - 1
            If IsGroupRow(lIndex) Then
                Set grdGroup = .RowData(lIndex)
                If grdGroup.Key = tsoGroup.Key Then
                    lRow = lIndex
                    Exit For
                End If
            End If
        Next lIndex
        
        GroupToGrid tsoGroup, lRow
        If lRow <> -1& Then
            OrderToGrid tsoGroup.tsOrderGroup.Order(lOrderNumber), tsoGroup.DisplayStatus(lOrderNumber), tsoGroup.Status(lOrderNumber), tsoGroup.DisplayLevel(lOrderNumber), tsoGroup.Action(lOrderNumber), lRow + lOrderNumber
        End If
        
        FilterGrid
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.UpdateGroupOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveGroup
'' Description: Remove the given group from the grid
'' Inputs:      Group
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveGroup(ByVal tsoGroup As cActiveTsOrderGroup)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim grdGroup As cActiveTsOrderGroup ' Group out of the grid
    Dim lChild As Long                  ' Child row in the grid
    
    With fgGroups
        For lIndex = .FixedRows To .Rows - 1
            If IsGroupRow(lIndex) Then
                Set grdGroup = .RowData(lIndex)
                If grdGroup.Key = tsoGroup.Key Then
                    lChild = .GetNodeRow(lIndex, flexNTFirstChild)
                    Do While lChild <> -1&
                        .RemoveItem lChild
                        lChild = .GetNodeRow(lIndex, flexNTFirstChild)
                    Loop
                    
                    .RemoveItem lIndex
                    Exit For
                End If
            End If
        Next lIndex
            
        FilterGrid
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.RemoveGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrid
'' Description: Filter the grid based on broker entitlements
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FilterGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' State of the grids redraw

    With fgGroups
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        HandleManageRow
        .AutoSize 0, .Cols - 1, False, 75
        
        If Not g.ConsoleForms Is Nothing Then
            g.ConsoleForms.NumVisible(eGDConsoleForm_TradeSenseOrders) = NumVisible
        End If
        
        ChangeBackColors
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.FilterGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_AfterCollapse
'' Description: Do things after the user expands/collapses nodes in the tree
'' Inputs:      Row, State
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.fgGroups_AfterCollapse"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_BeforeMouseDown
'' Description: Bring up the popup menu on a right click by the user
'' Inputs:      Button, Shift/Ctrl/Alt Status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseCol As Long               ' Column that the mouse is in
    Dim grp As cActiveTsOrderGroup      ' Currently selected group

    fgGroups.Row = fgGroups.MouseRow

    If Button = vbRightButton Then
        EnableControls
        PopupMenu mnuPopUp
    Else
        lMouseCol = fgGroups.MouseCol
        
        If ValidRowSelected Then
            If IsGroupRow Then
                If lMouseCol = GDCol(eGDCol_Cancel) Then
                    CancelGroup
                ElseIf lMouseCol = GDCol(eGDCol_Action) Then
                    Set grp = SelectedGroup
                    If Not grp Is Nothing Then
                        If grp.Submitted = True Then
                            ParkGroup
                        Else
                            SubmitGroup
                        End If
                    End If
                End If
            ElseIf IsManageRow = True Then
                tmrMenu.Tag = "MANAGE"
                tmrMenu.Enabled = True
            Else
                If lMouseCol = GDCol(eGDCol_Cancel) Then
                    CancelOrder
                ElseIf lMouseCol = GDCol(eGDCol_Action) Then
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.fgGroups_BeforeMouseDown"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_BeforeScroll
'' Description: Make sure left col stays the same if no horizontal
'' Inputs:      Old Top Row, Old Left Col, New Top Row, New Left Col, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    GridScrollCheck fgGroups, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.fgGroups_BeforeScroll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_MouseMove
'' Description: Bring up a tool tip if the user hovers over an order
'' Inputs:      Button, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim lMouseCol As Long               ' Mouse column in the grid
    Dim tsOrder As cTradeSenseOrder     ' Trade sense order
    Dim strTooltip As String            ' Tooltip text
    
    lMouseRow = fgGroups.MouseRow
    lMouseCol = fgGroups.MouseCol
    
    strTooltip = ""
    If (IsOrderRow(lMouseRow) = True) And (lMouseCol >= GDCol(eGDCol_Name)) And (lMouseCol < GDCol(eGDCol_Status)) Then
        Set tsOrder = fgGroups.RowData(lMouseRow)
        If Not tsOrder Is Nothing Then
            strTooltip = tsOrder.ToolTip
        End If
    End If
    
    If strTooltip <> fgGroups.ToolTipText Then
        fgGroups.ToolTipText = strTooltip
    End If

    If Screen.MousePointer <> vbHourglass Then
        If (Me.MousePointer = vbDefault) And (ValidRow(lMouseRow) = True) And (ValidCol(lMouseCol) = True) Then
            If fgGroups.Cell(flexcpFontUnderline, lMouseRow, lMouseCol) = True Then
                Me.MousePointer = vbCustom
                Me.MouseIcon = Picture16(ToolbarIcon("kHand"))
            End If
        ElseIf Me.MousePointer = vbCustom Then
            If (ValidRow(lMouseRow) = False) Or (ValidCol(lMouseCol) = False) Then
                Me.MousePointer = vbDefault
            ElseIf fgGroups.Cell(flexcpFontUnderline, lMouseRow, lMouseCol) = False Then
                Me.MousePointer = vbDefault
            End If
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement of the form

    Caption = "Active Trade Sense Order Groups"
    
    g.Styler.StyleForm Me
    
    Icon = Picture16(ToolbarIcon("kTradeSenseOrders"))
    
    strPlacement = GetIniFileProperty("frmActiveTsOrderGroups", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    mnuPopUp.Visible = False

    InitGrid
    LoadGrid


ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_MouseMove
'' Description: If the mouse cursor has been set somewhere else, reset it
'' Inputs:      Button pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Me.MousePointer = vbCustom Then
        Me.MousePointer = vbDefault
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, re-attach the grid
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = vbFormControlMenu Then
        If Not g.ConsoleForms Is Nothing Then
            g.ConsoleForms.ShowForm(eGDConsoleForm_TradeSenseOrders) = False
        End If
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.Form_QueryUnload"
    
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

    fgGroups.Move 0, 0, ScaleWidth, ScaleHeight

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save and clean up when the form is unloaded
'' Inputs:      Cancel Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmActiveTsOrderGroups", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.Form_Unload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCancelGroup_Click
'' Description: Allow the user to cancel a group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCancelGroup_Click()
On Error GoTo ErrSection:

    CancelGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.mnuCancelGroup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuManage_Click
'' Description: Allow the user to manage groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuManage_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "MANAGE"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.mnuManage_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuParkGroup_Click
'' Description: Allow the user to Park a group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuParkGroup_Click()
On Error GoTo ErrSection:

    ParkGroup
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.mnuParkGroup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSubmitGroup_Click
'' Description: Allow the user to Submit a group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSubmitGroup_Click()
On Error GoTo ErrSection:

    SubmitGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.mnuSubmitGroup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMenu_Timer
'' Description: Perform menu action chosen by user
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMenu_Timer()
On Error GoTo ErrSection:

    tmrMenu.Enabled = False
    
    Select Case UCase(tmrMenu.Tag)
        Case "MANAGE"
            frmTradeSenseOrderGroups.ShowMe
            
    End Select
    
    tmrMenu.Tag = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.tmrMenu_Timer"
    
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

    With fgGroups
        .Redraw = flexRDNone
        
        SetupGrid fgGroups, eGridMode_Grid
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarSimpleLeaf
        .SelectionMode = flexSelectionFree
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, eGDCol_Name) = "Group"
        .TextMatrix(0, eGDCol_Symbol) = "Symbol"
        .TextMatrix(0, eGDCol_Account) = "Account"
        .TextMatrix(0, eGdCol_Quantity) = "Qty"
        .TextMatrix(0, eGDCol_Status) = "Status"
        .TextMatrix(0, GDCol(eGDCol_Cancel)) = "X"
        .TextMatrix(0, GDCol(eGDCol_Action)) = "Action"
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.InitGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GroupToGrid
'' Description: Set the row in the grid to the given group
'' Inputs:      Group, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GroupToGrid(ByVal tsoGroup As cActiveTsOrderGroup, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current grid redraw setting
    Dim lIndex As Long                  ' Index into a for loop
    Dim bNewRow As Boolean              ' Did we create a new row in the grid?
    Dim strActionText As String         ' Action text
    Dim strAccount As String            ' Account name

    With fgGroups
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        bNewRow = False
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
            bNewRow = True
        End If
        
        .RowData(lRow) = tsoGroup
        
        strAccount = g.Broker.AccountNameForID(tsoGroup.AccountID)
        strActionText = "group '" & tsoGroup.tsOrderGroup.Name & "' for " & tsoGroup.Symbol & " in account '" & strAccount & "'"
        
        .TextMatrix(lRow, GDCol(eGDCol_Name)) = tsoGroup.tsOrderGroup.Name
        .TextMatrix(lRow, GDCol(eGDCol_Symbol)) = tsoGroup.Symbol
        .TextMatrix(lRow, GDCol(eGDCol_Account)) = strAccount
        .TextMatrix(lRow, GDCol(eGdCol_Quantity)) = Str(tsoGroup.Quantity)
        If tsoGroup.Submitted Then
            .TextMatrix(lRow, GDCol(eGDCol_Status)) = "Working"
            .TextMatrix(lRow, GDCol(eGDCol_Action)) = "Park " & strActionText
        Else
            .TextMatrix(lRow, GDCol(eGDCol_Status)) = "Parked"
            .TextMatrix(lRow, GDCol(eGDCol_Action)) = "Submit " & strActionText
        End If
        .TextMatrix(lRow, GDCol(eGDCol_Cancel)) = "X"
        
        .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Cancel)) = vbRed
        If g.nColorTheme = kDarkThemeColor Then
            .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Action)) = vbCyan
        Else
            .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Action)) = vbBlue
        End If
        .Cell(flexcpFontUnderline, lRow, GDCol(eGDCol_Cancel), lRow, GDCol(eGDCol_Action)) = True
        
        .RowOutlineLevel(lRow) = 0
        .IsSubtotal(lRow) = True
        
        For lIndex = 1 To tsoGroup.tsOrderGroup.OrderCount
            If bNewRow Then
                OrderToGrid tsoGroup.tsOrderGroup.Order(lIndex), tsoGroup.DisplayStatus(lIndex), tsoGroup.Status(lIndex), tsoGroup.DisplayLevel(lIndex), tsoGroup.Action(lIndex)
            Else
                OrderToGrid tsoGroup.tsOrderGroup.Order(lIndex), tsoGroup.DisplayStatus(lIndex), tsoGroup.Status(lIndex), tsoGroup.DisplayLevel(lIndex), tsoGroup.Action(lIndex), lRow + lIndex
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.GroupToGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToGrid
'' Description: Set the row in the grid to the given order
'' Inputs:      Order, Display Status, Status, Level, Action, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderToGrid(ByVal tsOrder As cTradeSenseOrder, ByVal strDisplayStatus As String, ByVal nStatus As eGD_TsoStatus, ByVal lLevel As Long, ByVal strAction As String, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current grid redraw setting

    With fgGroups
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = tsOrder
        
        .Cell(flexcpText, lRow, GDCol(eGDCol_Name), lRow, GDCol(eGdCol_Quantity)) = tsOrder.Name
        .TextMatrix(lRow, GDCol(eGDCol_Status)) = strDisplayStatus
        
        .TextMatrix(lRow, GDCol(eGDCol_Cancel)) = "X"
        If nStatus = eGD_TsoStatus_Closed Then
            .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Cancel)) = RGB(128, 128, 128)
        Else
            .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Cancel)) = vbRed
        End If
        .Cell(flexcpFontUnderline, lRow, GDCol(eGDCol_Cancel)) = True
        
        .TextMatrix(lRow, GDCol(eGDCol_Action)) = strAction
        
        .MergeRow(lRow) = True
        .RowOutlineLevel(lRow) = lLevel
        .IsSubtotal(lRow) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.OrderToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With fgGroups
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To g.TsoGroups.Count
            GroupToGrid g.TsoGroups(lIndex)
        Next lIndex
        
        FilterGrid
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.LoadGrid"
    
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

    Dim grp As cActiveTsOrderGroup      ' Order group object
    
    If (ValidRowSelected = False) Or (IsManageRow = True) Then
        Disable mnuSubmitGroup
        Disable mnuCancelGroup
        Disable mnuParkGroup
    ElseIf IsGroupRow Then
        Set grp = SelectedGroup
        Enable mnuSubmitGroup, (grp.Submitted = False)
        Enable mnuCancelGroup, True
        Enable mnuParkGroup, (grp.Submitted = True)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedGroup
'' Description: Selected group in the grid
'' Inputs:      None
'' Returns:     Selected Group (Nothing if none)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedGroup() As cActiveTsOrderGroup
On Error GoTo ErrSection:

    Dim grp As cActiveTsOrderGroup      ' Order group object
    Dim lRow As Long                    ' Row in the grid
    
    With fgGroups
        Set grp = Nothing
        
        lRow = GetGroupRow(.Row)
        If (lRow <> -1&) Then
            Set grp = .RowData(lRow)
        End If
    End With
    
    Set SelectedGroup = grp

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.SelectedGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelGroup
'' Description: Cancel group in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelGroup()
On Error GoTo ErrSection:

    Dim grp As cActiveTsOrderGroup      ' Order group object
    
    Set grp = SelectedGroup
    If Not grp Is Nothing Then
        g.TsoGroups.CancelSubmittedGroup grp, "User Cancel"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.CancelGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkGroup
'' Description: Park group in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParkGroup()
On Error GoTo ErrSection:

    Dim grp As cActiveTsOrderGroup      ' Order group object
    
    Set grp = SelectedGroup
    If Not grp Is Nothing Then
        g.TsoGroups.ParkSubmittedGroup grp, "User parking group from Active TradeSense Order Groups"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.ParkGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitGroup
'' Description: Submit group in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitGroup()
On Error GoTo ErrSection:

    Dim grp As cActiveTsOrderGroup      ' Order group object
    
    Set grp = SelectedGroup
    If Not grp Is Nothing Then
        g.TsoGroups.SubmitParkedGroup grp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.SubmitGroup"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Cancel order in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelOrder()
On Error GoTo ErrSection:

    Dim grp As cActiveTsOrderGroup      ' Order group object
    Dim tsOrder As cTradeSenseOrder     ' TradeSense order
    Dim lOrderNumber As Long            ' Order number
    
    Set grp = SelectedGroup
    lOrderNumber = OrderNumber
    
    If (Not grp Is Nothing) And (lOrderNumber <> -1&) Then
        Set tsOrder = grp.tsOrderGroup.Order(lOrderNumber)
        If InfBox("Are you sure you want to cancel order '" & tsOrder.Name & "' in group '" & grp.tsOrderGroup.Name & "' for '" & grp.Symbol & "'?", "?", "+Yes|-No", "Cancel Confirmation") = "Y" Then
            grp.CancelOrder lOrderNumber, "User Cancel", False
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.CancelOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Park order in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParkOrder()
On Error GoTo ErrSection:

    Dim grp As cActiveTsOrderGroup      ' Order group object
    Dim lOrderNumber As Long            ' Order number
    
    Set grp = SelectedGroup
    lOrderNumber = OrderNumber
    
    If (Not grp Is Nothing) And (lOrderNumber <> -1&) Then
        grp.ParkOrder lOrderNumber
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.ParkOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrder
'' Description: Submit order in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitOrder()
On Error GoTo ErrSection:

    Dim grp As cActiveTsOrderGroup      ' Order group object
    Dim lOrderNumber As Long            ' Order number
    
    Set grp = SelectedGroup
    lOrderNumber = OrderNumber
    
    If (Not grp Is Nothing) And (lOrderNumber <> -1&) Then
        grp.ActivateOrder lOrderNumber
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.SubmitOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRowSelected
'' Description: Is the currently selected row valid?
'' Inputs:      None
'' Returns:     True if valid row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidRowSelected() As Boolean
On Error GoTo ErrSection:

    ValidRowSelected = ValidRow(fgGroups.Row)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.ValidRowSelected"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRow
'' Description: Is the given row valid?
'' Inputs:      Row
'' Returns:     True if valid row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidRow(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    ValidRow = (lRow >= fgGroups.FixedRows) And (lRow < fgGroups.Rows)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.ValidRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidCol
'' Description: Is the given column valid?
'' Inputs:      Row
'' Returns:     True if valid column, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidCol(ByVal lCol As Long) As Boolean
On Error GoTo ErrSection:

    ValidCol = (lCol >= fgGroups.FixedCols) And (lCol < fgGroups.Cols)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.ValidCol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleManageRow
'' Description: Make sure that the "Manage Row" is the last row in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleManageRow()
On Error GoTo ErrSection:

    Dim lManageRow As Long              ' Row in the grid containing the manage row
    Dim lIndex As Long                  ' Index into a for loop
    
    With fgGroups
        lManageRow = -1&
        For lIndex = .FixedRows To .Rows - 1
            If IsManageRow(lIndex) = True Then
                lManageRow = lIndex
                Exit For
            End If
        Next lIndex
        
        If lManageRow = -1& Then
            .Rows = .Rows + 1
            .RowOutlineLevel(.Rows - 1) = 0
            .IsSubtotal(.Rows - 1) = True
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "Click here to Manage TradeSense Orders"
            .Cell(flexcpFontUnderline, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            If g.nColorTheme = kDarkThemeColor Then
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbCyan
            Else
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
            End If
            .MergeRow(.Rows - 1) = True
        ElseIf lManageRow <> .Rows - 1 Then
            .RowPosition(lIndex) = .Rows - 1
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.HandleManageRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumVisible
'' Description: Number of active TradeSense order groups in the grid
'' Inputs:      None
'' Returns:     Number of Groups
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NumVisible() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = 0&
    With fgGroups
        For lIndex = .FixedRows To .Rows - 1
            If (.RowOutlineLevel(lIndex) = 0) And (.MergeRow(lIndex) = False) Then
                lReturn = lReturn + 1&
            End If
        Next lIndex
    End With
    
    NumVisible = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.NumVisible"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetGroupRow
'' Description: Get the group row for the given row
'' Inputs:      Row
'' Returns:     Group Row (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetGroupRow(ByVal lRow As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lParent As Long                 ' Parent row
    
    lReturn = -1&
    If IsGroupRow(lRow) Then
        lReturn = lRow
    ElseIf IsOrderRow(lRow) Then
        lReturn = -1&
        
        lParent = fgGroups.GetNodeRow(lRow, flexNTParent)
        Do While (lParent <> -1&)
            If IsGroupRow(lParent) Then
                lReturn = lParent
                Exit Do
            End If
            
            lParent = fgGroups.GetNodeRow(lParent, flexNTParent)
        Loop
    End If
    
    GetGroupRow = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.GetGroupRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsGroupRow
'' Description: Determine if the given row is a group row or not
'' Inputs:      Row
'' Returns:     True if Group Row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsGroupRow(Optional ByVal lRow As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    With fgGroups
        If lRow = -1& Then
            lRow = .Row
        End If
        
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            bReturn = ((.RowOutlineLevel(lRow) = 0) And (.MergeRow(lRow) = False) And (TypeOf .RowData(lRow) Is cActiveTsOrderGroup))
        End If
    End With
    
    IsGroupRow = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.IsGroupRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsManageRow
'' Description: Determine if the given row is a manage row or not
'' Inputs:      Row
'' Returns:     True if Manage Row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsManageRow(Optional ByVal lRow As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    With fgGroups
        If lRow = -1& Then
            lRow = .Row
        End If
        
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            bReturn = (.RowOutlineLevel(lRow) = 0) And (.MergeRow(lRow) = True)
        End If
    End With
    
    IsManageRow = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.IsManageRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOrderRow
'' Description: Determine if the given row is an order row or not
'' Inputs:      Row
'' Returns:     True if Order Row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsOrderRow(Optional ByVal lRow As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    With fgGroups
        If lRow = -1& Then
            lRow = .Row
        End If
        
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            bReturn = (.RowOutlineLevel(lRow) > 0) And (.MergeRow(lRow) = True)
        End If
    End With
    
    IsOrderRow = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.IsOrderRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderNumber
'' Description: Determine the order number for the given row
'' Inputs:      Row
'' Returns:     Order Number (-1 if invalid)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderNumber(Optional ByVal lRow As Long = -1&) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lParent As Long                 ' Parent row
    
    If lRow = -1& Then
        lRow = fgGroups.Row
    End If
    
    lParent = GetGroupRow(lRow)
    If lParent <> -1& Then
        lReturn = lRow - lParent
    End If
    
    OrderNumber = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.OrderNumber"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeBackColors
'' Description: Change the background colors on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeBackColors()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Redraw setting for the grid
    Dim nCurrent As OLE_COLOR           ' Current background color
    
    With fgGroups
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        nCurrent = .BackColorAlternate
        For lIndex = .FixedRows To .Rows - 1
            If .RowOutlineLevel(lIndex) = 0 Then
                If .MergeRow(lIndex) = True Then
                    .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = .BackColor
                Else
                    .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = vbCyan
                End If
            ElseIf .RowOutlineLevel(lIndex) = 1 Then
                If nCurrent = .BackColor Then
                    nCurrent = .BackColorAlternate
                Else
                    nCurrent = .BackColor
                End If
                .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = nCurrent
            Else
                .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = nCurrent
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsOrderGroups.ChangeBackColors"
    
End Sub

